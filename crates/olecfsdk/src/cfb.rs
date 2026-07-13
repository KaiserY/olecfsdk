use std::path::{Component, Path, PathBuf};

use crate::{
    Error, Result,
    common::{FileTime, Guid},
    limits::Limits,
};

mod allocation;
mod directory;
mod header;
mod sector;
mod stream;
mod writer;

pub use allocation::{Difat, Fat, FatEntry, FatMarkerMismatch, MiniFat, MiniFatEntry};
pub use directory::{
    Directory, DirectoryColor, DirectoryEntry, DirectoryObjectType, DirectoryPointer,
};
pub use header::Header;
pub use sector::{MiniSectorId, SectorId};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Version {
    V3,
    V4,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EntryKind {
    Root,
    Storage,
    Stream,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Entry {
    pub path: PathBuf,
    pub name: String,
    pub kind: EntryKind,
    pub clsid: Guid,
    pub state_bits: u32,
    pub created: FileTime,
    pub modified: FileTime,
    pub data: Vec<u8>,
}

impl Entry {
    pub fn is_stream(&self) -> bool {
        self.kind == EntryKind::Stream
    }
    pub fn is_storage(&self) -> bool {
        self.kind != EntryKind::Stream
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CompoundFile {
    version: Version,
    header: Header,
    difat: Difat,
    fat: Fat,
    mini_fat: MiniFat,
    directory: Directory,
    entries: Vec<Entry>,
    unallocated_sectors: Vec<Vec<u8>>,
    trailing_data: Vec<u8>,
}

impl CompoundFile {
    pub fn new(version: Version) -> Result<Self> {
        Self::from_bytes(&writer::write_empty_compound(version)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        if bytes.len() as u64 > limits.max_file_size {
            return Err(Error::Limit(format!(
                "file length {} exceeds {}",
                bytes.len(),
                limits.max_file_size
            )));
        }
        let header = Header::from_bytes(bytes)?;
        let sector_len = header.sector_len();
        let padded_len = bytes
            .len()
            .checked_add(sector_len - 1)
            .map(|len| len / sector_len * sector_len)
            .ok_or_else(|| Error::Limit("padded CFB length overflow".into()))?;
        let mut padded_bytes = bytes.to_vec();
        padded_bytes.resize(padded_len, 0);
        let sectors = sector::SectorSource::new(&padded_bytes, bytes.len(), &header)?;
        if sectors.sector_count() > u32::MAX as usize {
            return Err(Error::Limit("CFB sector count exceeds u32".into()));
        }
        sectors.full_sector(SectorId::new(header.first_directory_sector)?)?;
        let difat = Difat::read(&header, &sectors, limits)?;
        let fat = Fat::read(&difat, &sectors)?;
        let directory_sectors = fat.chain(header.first_directory_sector, sectors.sector_count())?;
        let mini_fat = MiniFat::read(&header, &fat, &sectors, limits)?;
        let directory = Directory::read(
            directory_sectors,
            header.number_of_directory_sectors,
            &sectors,
            limits,
        )?;
        let version = header.version();
        let entries = stream::read_entries(&header, &fat, &mini_fat, &directory, &sectors, limits)?;
        let mut unallocated_sectors = Vec::new();
        let mut unallocated_bytes = 0usize;
        for index in 0..sectors.sector_count() {
            let id = SectorId::new(index as u32)?;
            if sectors.is_partial(id) {
                continue;
            }
            let is_allocation_sector =
                difat.fat_sectors().contains(&id) || difat.difat_sectors().contains(&id);
            if !is_allocation_sector && fat.is_free_or_unaddressed(id) {
                unallocated_bytes = unallocated_bytes
                    .checked_add(sectors.sector_len())
                    .ok_or_else(|| Error::Limit("unallocated sector size overflow".into()))?;
                if unallocated_bytes > limits.max_allocation {
                    return Err(Error::Limit(format!(
                        "unallocated sector data {unallocated_bytes} exceeds {}",
                        limits.max_allocation
                    )));
                }
                unallocated_sectors.push(sectors.sector(id)?.to_vec());
            }
        }
        let trailing_data = sectors.unaccessed_partial_data().to_vec();
        Ok(Self {
            version,
            header,
            difat,
            fat,
            mini_fat,
            directory,
            entries,
            unallocated_sectors,
            trailing_data,
        })
    }

    pub fn version(&self) -> Version {
        self.version
    }
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }
    pub fn header(&self) -> &Header {
        &self.header
    }
    pub fn difat(&self) -> &Difat {
        &self.difat
    }
    pub fn fat(&self) -> &Fat {
        &self.fat
    }
    pub fn mini_fat(&self) -> &MiniFat {
        &self.mini_fat
    }
    pub fn directory_sectors(&self) -> &[SectorId] {
        self.directory.sectors()
    }
    pub fn directory(&self) -> &Directory {
        &self.directory
    }
    pub fn entry(&self, path: impl AsRef<Path>) -> Option<&Entry> {
        self.entries
            .iter()
            .find(|entry| entry.path == path.as_ref())
    }
    pub fn stream(&self, path: impl AsRef<Path>) -> Option<&[u8]> {
        self.entry(path)
            .filter(|entry| entry.is_stream())
            .map(|entry| entry.data.as_slice())
    }
    pub fn stream_mut(&mut self, path: impl AsRef<Path>) -> Option<&mut Vec<u8>> {
        let path = path.as_ref();
        self.entries
            .iter_mut()
            .find(|entry| entry.path == path && entry.is_stream())
            .map(|entry| &mut entry.data)
    }
    pub fn replace_stream(&mut self, path: impl AsRef<Path>, data: Vec<u8>) -> Result<Vec<u8>> {
        let path = path.as_ref();
        let entry = self
            .entries
            .iter_mut()
            .find(|entry| entry.path == path)
            .ok_or_else(|| {
                Error::invalid(0, format!("CFB entry {} does not exist", path.display()))
            })?;
        if !entry.is_stream() {
            return Err(Error::invalid(
                0,
                format!("CFB entry {} is not a stream", path.display()),
            ));
        }
        Ok(std::mem::replace(&mut entry.data, data))
    }
    pub fn create_storage(&mut self, path: impl AsRef<Path>) -> Result<()> {
        self.create_entry(path.as_ref(), EntryKind::Storage, Vec::new())
    }
    pub fn create_stream(&mut self, path: impl AsRef<Path>, data: Vec<u8>) -> Result<()> {
        self.create_entry(path.as_ref(), EntryKind::Stream, data)
    }

    pub fn rename_entry(&mut self, path: impl AsRef<Path>, new_name: &str) -> Result<()> {
        let path = path.as_ref();
        if path == Path::new("/") {
            return Err(Error::invalid(0, "CFB root entry cannot be renamed"));
        }
        writer::validate_entry_name(new_name)?;
        let index = self
            .entries
            .iter()
            .position(|entry| entry.path == path)
            .ok_or_else(|| {
                Error::invalid(0, format!("CFB entry {} does not exist", path.display()))
            })?;
        let parent = path
            .parent()
            .ok_or_else(|| Error::invalid(0, "CFB entry path has no parent"))?;
        if self.entries.iter().enumerate().any(|(sibling, entry)| {
            sibling != index
                && entry.path.parent() == Some(parent)
                && writer::names_equal(&entry.name, new_name)
        }) {
            return Err(Error::invalid(
                0,
                format!("CFB sibling name {new_name} already exists case-insensitively"),
            ));
        }
        let new_path = parent.join(new_name);
        let old_path = path.to_path_buf();
        for entry in &mut self.entries {
            let Ok(suffix) = entry.path.strip_prefix(&old_path) else {
                continue;
            };
            entry.path = if suffix.as_os_str().is_empty() {
                new_path.clone()
            } else {
                new_path.join(suffix)
            };
        }
        self.entries[index].name = new_name.to_owned();
        Ok(())
    }

    pub fn remove_entry(&mut self, path: impl AsRef<Path>) -> Result<Entry> {
        let path = path.as_ref();
        if path == Path::new("/") {
            return Err(Error::invalid(0, "CFB root entry cannot be removed"));
        }
        let index = self
            .entries
            .iter()
            .position(|entry| entry.path == path)
            .ok_or_else(|| {
                Error::invalid(0, format!("CFB entry {} does not exist", path.display()))
            })?;
        if self
            .entries
            .iter()
            .any(|entry| entry.path != path && entry.path.starts_with(path))
        {
            return Err(Error::invalid(
                0,
                format!("CFB storage {} is not empty", path.display()),
            ));
        }
        Ok(self.entries.remove(index))
    }

    fn create_entry(&mut self, path: &Path, kind: EntryKind, data: Vec<u8>) -> Result<()> {
        let mut components = path.components();
        if components.next() != Some(Component::RootDir)
            || !components
                .clone()
                .all(|part| matches!(part, Component::Normal(_)))
            || components.next().is_none()
        {
            return Err(Error::invalid(
                0,
                format!(
                    "CFB entry path {} is not canonical and absolute",
                    path.display()
                ),
            ));
        }
        let name = path
            .file_name()
            .and_then(|name| name.to_str())
            .ok_or_else(|| Error::invalid(0, "CFB entry name is not valid UTF-8"))?;
        writer::validate_entry_name(name)?;
        if self.entries.iter().any(|entry| entry.path == path) {
            return Err(Error::invalid(
                0,
                format!("CFB entry {} already exists", path.display()),
            ));
        }
        let parent = path
            .parent()
            .ok_or_else(|| Error::invalid(0, "CFB entry path has no parent"))?;
        let parent_entry = self
            .entries
            .iter()
            .find(|entry| entry.path == parent)
            .ok_or_else(|| {
                Error::invalid(0, format!("CFB parent {} does not exist", parent.display()))
            })?;
        if parent_entry.is_stream() {
            return Err(Error::invalid(
                0,
                format!("CFB parent {} is a stream", parent.display()),
            ));
        }
        if self.entries.iter().any(|entry| {
            entry.path.parent() == Some(parent) && writer::names_equal(&entry.name, name)
        }) {
            return Err(Error::invalid(
                0,
                format!("CFB sibling name {name} already exists case-insensitively"),
            ));
        }
        self.entries.push(Entry {
            path: path.to_path_buf(),
            name: name.to_owned(),
            kind,
            clsid: Guid::ZERO,
            state_bits: 0,
            created: FileTime::ZERO,
            modified: FileTime::ZERO,
            data,
        });
        Ok(())
    }
    pub fn trailing_data(&self) -> &[u8] {
        &self.trailing_data
    }
    pub fn unallocated_sectors(&self) -> &[Vec<u8>] {
        &self.unallocated_sectors
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        writer::write_compound(self)
    }

    pub fn logical_eq(&self, other: &Self) -> bool {
        self.version == other.version
            && self.entries == other.entries
            && self.unallocated_sectors == other.unallocated_sectors
            && self.trailing_data == other.trailing_data
    }
}

pub fn round_trip_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    let original = CompoundFile::from_bytes(bytes)?;
    let output = original.to_bytes()?;
    let reopened = CompoundFile::from_bytes(&output)?;
    if !original.logical_eq(&reopened) {
        return Err(Error::invalid(
            0,
            "CFB logical structure changed after round-trip",
        ));
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_compound_file_round_trips() {
        for (reference_version, version) in [
            (cfb::Version::V3, Version::V3),
            (cfb::Version::V4, Version::V4),
        ] {
            let source = cfb::CompoundFile::create_with_version(
                reference_version,
                std::io::Cursor::new(Vec::new()),
            )
            .unwrap()
            .into_inner()
            .into_inner();
            let output = round_trip_bytes(&source).unwrap();
            let parsed = CompoundFile::from_bytes(&output).unwrap();
            assert_eq!(parsed.version(), version);
        }
    }

    #[test]
    fn nested_streams_round_trip() {
        use std::io::Write;
        let mut source = cfb::CompoundFile::create(std::io::Cursor::new(Vec::new())).unwrap();
        source.create_storage("/Macros").unwrap();
        source.create_storage("/Macros/VBA").unwrap();
        source
            .create_stream("/WordDocument")
            .unwrap()
            .write_all(b"word")
            .unwrap();
        source
            .create_stream("/Macros/VBA/dir")
            .unwrap()
            .write_all(b"vba")
            .unwrap();
        source.flush().unwrap();
        let bytes = source.into_inner().into_inner();
        let output = round_trip_bytes(&bytes).unwrap();
        cfb::CompoundFile::open_strict(std::io::Cursor::new(output.clone())).unwrap();
        let parsed = CompoundFile::from_bytes(&output).unwrap();
        assert_eq!(parsed.entry("/WordDocument").unwrap().data, b"word");
        assert_eq!(parsed.entry("/Macros/VBA/dir").unwrap().data, b"vba");
    }

    #[test]
    fn native_stream_editing_crosses_mini_stream_cutoff_in_v3_and_v4() {
        use std::io::Write;

        for reference_version in [cfb::Version::V3, cfb::Version::V4] {
            let mut source = cfb::CompoundFile::create_with_version(
                reference_version,
                std::io::Cursor::new(Vec::new()),
            )
            .unwrap();
            source.create_storage("/Data").unwrap();
            source
                .create_stream("/Data/Small")
                .unwrap()
                .write_all(b"small")
                .unwrap();
            source
                .create_stream("/Data/Large")
                .unwrap()
                .write_all(&vec![0x11; 5_000])
                .unwrap();
            source.flush().unwrap();

            let mut compound = CompoundFile::from_bytes(&source.into_inner().into_inner()).unwrap();
            assert_eq!(compound.stream("/Data/Small"), Some(b"small".as_slice()));
            assert_eq!(compound.stream("/Data"), None);
            assert_eq!(
                compound
                    .replace_stream("/Data/Small", vec![0x22; 5_001])
                    .unwrap(),
                b"small"
            );
            *compound.stream_mut("/Data/Large").unwrap() = vec![0x33; 63];
            assert!(compound.replace_stream("/Data", Vec::new()).is_err());
            assert!(compound.replace_stream("/Missing", Vec::new()).is_err());

            let encoded = compound.to_bytes().unwrap();
            let reopened = CompoundFile::from_bytes(&encoded).unwrap();
            assert_eq!(
                reopened.stream("/Data/Small"),
                Some(vec![0x22; 5_001].as_slice())
            );
            assert_eq!(
                reopened.stream("/Data/Large"),
                Some(vec![0x33; 63].as_slice())
            );
            cfb::CompoundFile::open_strict(std::io::Cursor::new(encoded)).unwrap();
        }
    }

    #[test]
    fn native_directory_creation_builds_nested_v3_and_v4_trees() {
        for version in [Version::V3, Version::V4] {
            let mut compound = CompoundFile::new(version).unwrap();
            assert_eq!(compound.version(), version);
            assert_eq!(compound.entries().len(), 1);

            compound.create_storage("/Data").unwrap();
            compound.create_storage("/Data/Nested").unwrap();
            compound
                .create_stream("/Data/Nested/Mini", vec![0x44; 63])
                .unwrap();
            compound
                .create_stream("/Regular", vec![0x55; 4_096])
                .unwrap();
            compound.create_stream("/Empty", Vec::new()).unwrap();
            assert!(compound.create_storage("/").is_err());
            assert!(compound.create_storage("relative").is_err());
            assert!(compound.create_storage("/Missing/Child").is_err());
            assert!(compound.create_storage("/Regular/Child").is_err());
            assert!(compound.create_stream("/Data", Vec::new()).is_err());
            assert!(compound.create_stream("/data", Vec::new()).is_err());
            assert!(compound.create_stream("/Bad:Name", Vec::new()).is_err());
            assert!(compound.create_stream("/Bad\0Name", Vec::new()).is_err());
            assert!(compound.create_stream("/Data/../Oops", Vec::new()).is_err());
            assert!(
                compound
                    .create_stream(format!("/{}", "x".repeat(32)), Vec::new())
                    .is_err()
            );
            compound.rename_entry("/Data", "Archive").unwrap();
            assert!(compound.entry("/Data").is_none());
            assert!(compound.entry("/Archive/Nested/Mini").is_some());
            compound
                .rename_entry("/Archive/Nested/Mini", "mini")
                .unwrap();
            compound
                .create_stream("/Archive/Nested/Other", b"other".to_vec())
                .unwrap();
            assert!(
                compound
                    .rename_entry("/Archive/Nested/mini", "OTHER")
                    .is_err()
            );
            assert!(compound.rename_entry("/Regular", "archive").is_err());
            assert!(compound.rename_entry("/", "Root2").is_err());
            assert!(compound.rename_entry("/Missing", "Missing2").is_err());
            assert!(compound.rename_entry("/Regular", "Bad:Name").is_err());
            assert!(compound.remove_entry("/Archive").is_err());
            let removed = compound.remove_entry("/Empty").unwrap();
            assert_eq!(removed.kind, EntryKind::Stream);
            assert!(compound.remove_entry("/Empty").is_err());
            compound.create_storage("/Vacant").unwrap();
            assert_eq!(
                compound.remove_entry("/Vacant").unwrap().kind,
                EntryKind::Storage
            );
            assert!(compound.remove_entry("/").is_err());

            let encoded = compound.to_bytes().unwrap();
            cfb::CompoundFile::open_strict(std::io::Cursor::new(encoded.clone())).unwrap();
            let reopened = CompoundFile::from_bytes(&encoded).unwrap();
            assert_eq!(
                reopened.stream("/Archive/Nested/mini"),
                Some([0x44; 63].as_slice())
            );
            assert_eq!(
                reopened.stream("/Archive/Nested/Other"),
                Some(b"other".as_slice())
            );
            assert_eq!(
                reopened.stream("/Regular"),
                Some(vec![0x55; 4_096].as_slice())
            );
            assert!(reopened.stream("/Empty").is_none());
            assert!(reopened.entry("/Archive/Nested").unwrap().is_storage());
        }
    }

    #[test]
    fn trailing_data_is_preserved_outside_sector_space() {
        let source = cfb::CompoundFile::create(std::io::Cursor::new(Vec::new())).unwrap();
        let mut bytes = source.into_inner().into_inner();
        bytes.extend_from_slice(b"trailing");
        let parsed = CompoundFile::from_bytes(&bytes).unwrap();
        assert_eq!(parsed.trailing_data(), b"trailing");
        let output = parsed.to_bytes().unwrap();
        assert!(output.ends_with(b"trailing"));
        assert!(parsed.logical_eq(&CompoundFile::from_bytes(&output).unwrap()));
    }

    #[test]
    fn unallocated_physical_sectors_are_preserved() {
        let source = cfb::CompoundFile::create_with_version(
            cfb::Version::V3,
            std::io::Cursor::new(Vec::new()),
        )
        .unwrap();
        let mut bytes = source.into_inner().into_inner();
        bytes.extend_from_slice(&[0x5a; 512]);
        let parsed = CompoundFile::from_bytes(&bytes).unwrap();
        assert_eq!(parsed.unallocated_sectors(), [vec![0x5a; 512]]);
        let output = parsed.to_bytes().unwrap();
        let reopened = CompoundFile::from_bytes(&output).unwrap();
        assert!(parsed.logical_eq(&reopened));
    }
}
