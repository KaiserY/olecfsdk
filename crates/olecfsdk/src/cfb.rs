use std::{
    io::{Cursor, Read, Write},
    path::{Path, PathBuf},
};

use crate::{
    Error, Result,
    common::{FileTime, Guid},
    limits::Limits,
};

mod allocation;
mod directory;
mod header;
mod name;
mod sector;
mod stream;
mod writer;

pub use allocation::{Difat, Fat, FatEntry, FatMarkerMismatch, MiniFat, MiniFatEntry};
pub use directory::{
    Directory, DirectoryColor, DirectoryEntry, DirectoryObjectType, DirectoryPointer,
};
pub use header::Header;
pub use name::compare_names;
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
    header_padding_is_zero: bool,
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

    /// Parses a CFB image and additionally enforces canonical MS-CFB fields.
    ///
    /// Use [`Self::from_bytes`] for compatibility reading of legacy producer
    /// quirks that are normalized by the deterministic writer.
    pub fn from_bytes_strict(bytes: &[u8]) -> Result<Self> {
        let compound = Self::from_bytes(bytes)?;
        compound.validate_strict()?;
        Ok(compound)
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
        let header_padding_is_zero = if sector_len > header::HEADER_LEN {
            bytes
                .get(header::HEADER_LEN..sector_len)
                .is_some_and(|padding| padding.iter().all(|byte| *byte == 0))
        } else {
            true
        };
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
            header_padding_is_zero,
            unallocated_sectors,
            trailing_data,
        })
    }

    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_limits(reader, Limits::default())
    }

    pub fn from_reader_strict(reader: impl Read) -> Result<Self> {
        let compound = Self::from_reader(reader)?;
        compound.validate_strict()?;
        Ok(compound)
    }

    pub fn from_reader_with_limits(mut reader: impl Read, limits: Limits) -> Result<Self> {
        let maximum = limits.max_file_size.saturating_add(1);
        let mut bytes = Vec::new();
        reader.by_ref().take(maximum).read_to_end(&mut bytes)?;
        if bytes.len() as u64 > limits.max_file_size {
            return Err(Error::Limit(format!(
                "file length exceeds {}",
                limits.max_file_size
            )));
        }
        Self::from_bytes_with_limits(&bytes, limits)
    }

    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_reader(std::fs::File::open(path)?)
    }

    pub fn open_strict(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_reader_strict(std::fs::File::open(path)?)
    }

    pub fn version(&self) -> Version {
        self.version
    }
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }
    pub fn root_entry(&self) -> &Entry {
        &self.entries[0]
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
        self.entry_index(path.as_ref())
            .map(|index| &self.entries[index])
    }
    pub fn contains_entry(&self, path: impl AsRef<Path>) -> bool {
        self.entry(path).is_some()
    }
    pub fn is_stream(&self, path: impl AsRef<Path>) -> bool {
        self.entry(path).is_some_and(Entry::is_stream)
    }
    pub fn is_storage(&self, path: impl AsRef<Path>) -> bool {
        self.entry(path).is_some_and(Entry::is_storage)
    }
    pub fn children(&self, path: impl AsRef<Path>) -> Result<Vec<&Entry>> {
        let parent = self.required_entry_index(path.as_ref())?;
        if self.entries[parent].is_stream() {
            return Err(Error::invalid(
                0,
                format!(
                    "CFB entry {} is not a storage",
                    self.entries[parent].path.display()
                ),
            ));
        }
        let parent_path = self.entries[parent].path.as_path();
        let mut children: Vec<_> = self
            .entries
            .iter()
            .filter(|entry| entry.path.parent() == Some(parent_path))
            .collect();
        children.sort_by(|left, right| name::compare_names(&left.name, &right.name));
        Ok(children)
    }
    pub fn walk_storage(&self, path: impl AsRef<Path>) -> Result<Vec<&Entry>> {
        let root = self.required_entry_index(path.as_ref())?;
        if self.entries[root].is_stream() {
            return Err(Error::invalid(
                0,
                format!(
                    "CFB entry {} is not a storage",
                    self.entries[root].path.display()
                ),
            ));
        }
        let mut output = Vec::new();
        let mut stack = vec![root];
        while let Some(index) = stack.pop() {
            let entry = &self.entries[index];
            output.push(entry);
            if entry.is_stream() {
                continue;
            }
            let mut children: Vec<_> = self
                .entries
                .iter()
                .enumerate()
                .filter(|(_, candidate)| candidate.path.parent() == Some(entry.path.as_path()))
                .map(|(index, _)| index)
                .collect();
            children.sort_by(|left, right| {
                name::compare_names(&self.entries[*left].name, &self.entries[*right].name)
            });
            stack.extend(children.into_iter().rev());
        }
        Ok(output)
    }
    pub fn stream(&self, path: impl AsRef<Path>) -> Option<&[u8]> {
        self.entry(path)
            .filter(|entry| entry.is_stream())
            .map(|entry| entry.data.as_slice())
    }
    /// Opens an owned-model stream through the standard `Read + Seek` cursor API.
    pub fn open_stream(&self, path: impl AsRef<Path>) -> Result<Cursor<&[u8]>> {
        let index = self.required_entry_index(path.as_ref())?;
        if !self.entries[index].is_stream() {
            return Err(Error::invalid(0, "CFB entry is not a stream"));
        }
        Ok(Cursor::new(self.entries[index].data.as_slice()))
    }
    pub fn stream_mut(&mut self, path: impl AsRef<Path>) -> Option<&mut Vec<u8>> {
        let index = self.entry_index(path.as_ref())?;
        self.entries[index]
            .is_stream()
            .then_some(&mut self.entries[index].data)
    }
    /// Opens an owned-model stream through the standard `Read + Write + Seek` cursor API.
    pub fn open_stream_mut(&mut self, path: impl AsRef<Path>) -> Result<Cursor<&mut Vec<u8>>> {
        let index = self.required_entry_index(path.as_ref())?;
        if !self.entries[index].is_stream() {
            return Err(Error::invalid(0, "CFB entry is not a stream"));
        }
        Ok(Cursor::new(&mut self.entries[index].data))
    }
    pub fn replace_stream(&mut self, path: impl AsRef<Path>, data: Vec<u8>) -> Result<Vec<u8>> {
        let path = path.as_ref();
        let index = self.required_entry_index(path)?;
        let entry = &mut self.entries[index];
        if !entry.is_stream() {
            return Err(Error::invalid(
                0,
                format!("CFB entry {} is not a stream", path.display()),
            ));
        }
        Ok(std::mem::replace(&mut entry.data, data))
    }

    pub fn create_or_replace_stream(
        &mut self,
        path: impl AsRef<Path>,
        data: Vec<u8>,
    ) -> Result<Option<Vec<u8>>> {
        let path = path.as_ref();
        match self.entry_index(path) {
            Some(index) if self.entries[index].is_stream() => {
                Ok(Some(std::mem::replace(&mut self.entries[index].data, data)))
            }
            Some(index) => Err(Error::invalid(
                0,
                format!(
                    "CFB entry {} is not a stream",
                    self.entries[index].path.display()
                ),
            )),
            None => {
                self.create_stream(path, data)?;
                Ok(None)
            }
        }
    }

    pub fn replace_storage_class_id(
        &mut self,
        path: impl AsRef<Path>,
        class_id: Guid,
    ) -> Result<Guid> {
        let path = path.as_ref();
        let index = self.required_entry_index(path)?;
        let entry = &mut self.entries[index];
        if !entry.is_storage() {
            return Err(Error::invalid(
                0,
                format!("CFB entry {} is not a storage", path.display()),
            ));
        }
        Ok(std::mem::replace(&mut entry.clsid, class_id))
    }

    pub fn replace_state_bits(&mut self, path: impl AsRef<Path>, bits: u32) -> Result<u32> {
        let index = self.required_entry_index(path.as_ref())?;
        Ok(std::mem::replace(&mut self.entries[index].state_bits, bits))
    }

    pub fn replace_creation_time(
        &mut self,
        path: impl AsRef<Path>,
        time: FileTime,
    ) -> Result<FileTime> {
        let index = self.required_entry_index(path.as_ref())?;
        if self.entries[index].kind != EntryKind::Storage {
            return Err(Error::invalid(
                0,
                "CFB creation time is writable only for non-root storage entries",
            ));
        }
        Ok(std::mem::replace(&mut self.entries[index].created, time))
    }

    pub fn replace_modified_time(
        &mut self,
        path: impl AsRef<Path>,
        time: FileTime,
    ) -> Result<FileTime> {
        let index = self.required_entry_index(path.as_ref())?;
        if self.entries[index].is_stream() {
            return Err(Error::invalid(
                0,
                "CFB stream modified time must remain zero",
            ));
        }
        Ok(std::mem::replace(&mut self.entries[index].modified, time))
    }

    pub fn create_storage(&mut self, path: impl AsRef<Path>) -> Result<()> {
        self.create_entry(path.as_ref(), EntryKind::Storage, Vec::new())
    }
    pub fn create_storage_all(&mut self, path: impl AsRef<Path>) -> Result<()> {
        let names = name::path_components(path.as_ref()).ok_or_else(|| {
            Error::invalid(0, format!("invalid CFB path {}", path.as_ref().display()))
        })?;
        let mut parent = 0usize;
        for name in names {
            let parent_path = self.entries[parent].path.clone();
            let candidate = parent_path.join(&name);
            match self.entry_index(&candidate) {
                Some(index) if self.entries[index].is_storage() => parent = index,
                Some(index) => {
                    return Err(Error::invalid(
                        0,
                        format!(
                            "CFB entry {} is not a storage",
                            self.entries[index].path.display()
                        ),
                    ));
                }
                None => {
                    self.create_storage(&candidate)?;
                    parent = self.required_entry_index(&candidate)?;
                }
            }
        }
        Ok(())
    }
    pub fn create_stream(&mut self, path: impl AsRef<Path>, data: Vec<u8>) -> Result<()> {
        self.create_entry(path.as_ref(), EntryKind::Stream, data)
    }

    pub fn rename_entry(&mut self, path: impl AsRef<Path>, new_name: &str) -> Result<()> {
        let path = path.as_ref();
        let index = self.required_entry_index(path)?;
        if self.entries[index].kind == EntryKind::Root {
            return Err(Error::invalid(0, "CFB root entry cannot be renamed"));
        }
        name::validate_entry_name(new_name)?;
        let old_path = self.entries[index].path.clone();
        let parent = old_path
            .parent()
            .ok_or_else(|| Error::invalid(0, "CFB entry path has no parent"))?;
        if self.entries.iter().enumerate().any(|(sibling, entry)| {
            sibling != index
                && entry.path.parent() == Some(parent)
                && name::names_equal(&entry.name, new_name)
        }) {
            return Err(Error::invalid(
                0,
                format!("CFB sibling name {new_name} already exists case-insensitively"),
            ));
        }
        let new_path = parent.join(new_name);
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
        let index = self.required_entry_index(path)?;
        if self.entries[index].kind == EntryKind::Root {
            return Err(Error::invalid(0, "CFB root entry cannot be removed"));
        }
        let actual_path = self.entries[index].path.clone();
        if self
            .entries
            .iter()
            .any(|entry| entry.path != actual_path && entry.path.starts_with(&actual_path))
        {
            return Err(Error::invalid(
                0,
                format!("CFB storage {} is not empty", path.display()),
            ));
        }
        Ok(self.entries.remove(index))
    }

    pub fn remove_stream(&mut self, path: impl AsRef<Path>) -> Result<Entry> {
        let index = self.required_entry_index(path.as_ref())?;
        if !self.entries[index].is_stream() {
            return Err(Error::invalid(0, "CFB entry is not a stream"));
        }
        self.remove_entry(path)
    }

    pub fn remove_storage(&mut self, path: impl AsRef<Path>) -> Result<Entry> {
        let index = self.required_entry_index(path.as_ref())?;
        if self.entries[index].kind != EntryKind::Storage {
            return Err(Error::invalid(0, "CFB entry is not a non-root storage"));
        }
        self.remove_entry(path)
    }

    pub fn remove_storage_all(&mut self, path: impl AsRef<Path>) -> Result<Vec<Entry>> {
        let index = self.required_entry_index(path.as_ref())?;
        if self.entries[index].is_stream() {
            return Err(Error::invalid(0, "CFB entry is not a storage"));
        }
        let root = self.entries[index].path.clone();
        let remove_root = self.entries[index].kind != EntryKind::Root;
        let mut removed = Vec::new();
        let mut kept = Vec::with_capacity(self.entries.len());
        for entry in self.entries.drain(..) {
            let matches = entry.path.starts_with(&root) && (remove_root || entry.path != root);
            if matches {
                removed.push(entry);
            } else {
                kept.push(entry);
            }
        }
        self.entries = kept;
        removed.sort_by(|left, right| {
            right
                .path
                .components()
                .count()
                .cmp(&left.path.components().count())
        });
        Ok(removed)
    }

    fn create_entry(&mut self, path: &Path, kind: EntryKind, data: Vec<u8>) -> Result<()> {
        let mut names = name::path_components(path)
            .ok_or_else(|| Error::invalid(0, format!("invalid CFB path {}", path.display())))?;
        let name = names
            .pop()
            .ok_or_else(|| Error::invalid(0, "CFB root entry already exists"))?;
        name::validate_entry_name(&name)?;
        if self.entry_index(path).is_some() {
            return Err(Error::invalid(
                0,
                format!("CFB entry {} already exists", path.display()),
            ));
        }
        let parent_index = self.entry_index_from_components(&names).ok_or_else(|| {
            Error::invalid(
                0,
                format!("CFB parent for {} does not exist", path.display()),
            )
        })?;
        let parent_entry = &self.entries[parent_index];
        if parent_entry.is_stream() {
            return Err(Error::invalid(
                0,
                format!("CFB parent {} is a stream", parent_entry.path.display()),
            ));
        }
        if self.entries.iter().any(|entry| {
            entry.path.parent() == Some(parent_entry.path.as_path())
                && name::names_equal(&entry.name, &name)
        }) {
            return Err(Error::invalid(
                0,
                format!("CFB sibling name {name} already exists case-insensitively"),
            ));
        }
        let entry_path = parent_entry.path.join(&name);
        self.entries.push(Entry {
            path: entry_path,
            name,
            kind,
            clsid: Guid::ZERO,
            state_bits: 0,
            created: FileTime::ZERO,
            modified: FileTime::ZERO,
            data,
        });
        Ok(())
    }

    fn entry_index(&self, path: &Path) -> Option<usize> {
        let names = name::path_components(path)?;
        self.entry_index_from_components(&names)
    }

    fn entry_index_from_components(&self, names: &[String]) -> Option<usize> {
        let mut current = self
            .entries
            .iter()
            .position(|entry| entry.kind == EntryKind::Root)?;
        for requested in names {
            let parent = self.entries[current].path.as_path();
            current = self.entries.iter().position(|entry| {
                entry.path.parent() == Some(parent) && name::names_equal(&entry.name, requested)
            })?;
        }
        Some(current)
    }

    fn required_entry_index(&self, path: &Path) -> Result<usize> {
        self.entry_index(path).ok_or_else(|| {
            Error::invalid(0, format!("CFB entry {} does not exist", path.display()))
        })
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

    pub fn write_to(&self, mut writer: impl Write) -> Result<()> {
        writer.write_all(&self.to_bytes()?)?;
        Ok(())
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.write_to(std::fs::File::create(path)?)
    }

    /// Validates the retained physical source representation in strict mode.
    ///
    /// Logical edits are materialized and validated when serialized; the raw
    /// header, allocation tables, and directory accessors intentionally keep
    /// describing the source image until that output is reopened.
    pub fn validate_strict(&self) -> Result<()> {
        if self.header.clsid != [0; 16] {
            return Err(Error::invalid(8, "CFB header CLSID must be zero"));
        }
        if self.header.reserved != [0; 6] {
            return Err(Error::invalid(34, "CFB header reserved bytes must be zero"));
        }
        if !self.header_padding_is_zero {
            return Err(Error::invalid(
                header::HEADER_LEN as u64,
                "CFB v4 header padding must be zero",
            ));
        }
        if self.version == Version::V3 && self.header.number_of_directory_sectors != 0 {
            return Err(Error::invalid(
                40,
                "CFB v3 directory sector count must be zero",
            ));
        }
        if !self.fat.marker_mismatches().is_empty() {
            return Err(Error::invalid(
                0,
                "CFB allocation sectors have non-canonical FAT markers",
            ));
        }
        self.difat.validate_strict()?;
        self.fat.validate_strict()?;
        if !self.mini_fat.sector_count_matches_header() {
            return Err(Error::invalid(
                64,
                "CFB MiniFAT sector count does not match the header",
            ));
        }
        let root_mini_sector_count = self
            .directory
            .root()
            .effective_stream_size(self.version)
            .checked_div(64)
            .and_then(|count| usize::try_from(count).ok())
            .ok_or_else(|| Error::Limit("root mini stream size does not fit usize".into()))?;
        self.mini_fat.validate_strict(root_mini_sector_count)?;
        self.directory.validate_strict(self.version)
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
    let reopened = CompoundFile::from_bytes_strict(&output)?;
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
    use std::io::{Seek, SeekFrom};

    use super::*;

    #[test]
    fn empty_compound_file_round_trips() {
        for version in [Version::V3, Version::V4] {
            let source = CompoundFile::new(version).unwrap().to_bytes().unwrap();
            let output = round_trip_bytes(&source).unwrap();
            let parsed = CompoundFile::from_bytes(&output).unwrap();
            assert_eq!(parsed.version(), version);
        }
    }

    #[test]
    fn nested_streams_round_trip() {
        let mut source = CompoundFile::new(Version::V3).unwrap();
        source.create_storage("/Macros").unwrap();
        source.create_storage("/Macros/VBA").unwrap();
        source
            .create_stream("/WordDocument", b"word".to_vec())
            .unwrap();
        source
            .create_stream("/Macros/VBA/dir", b"vba".to_vec())
            .unwrap();
        let bytes = source.to_bytes().unwrap();
        let output = round_trip_bytes(&bytes).unwrap();
        let parsed = CompoundFile::from_bytes(&output).unwrap();
        assert_eq!(parsed.entry("/WordDocument").unwrap().data, b"word");
        assert_eq!(parsed.entry("/Macros/VBA/dir").unwrap().data, b"vba");
    }

    #[test]
    fn native_stream_editing_crosses_mini_stream_cutoff_in_v3_and_v4() {
        for version in [Version::V3, Version::V4] {
            let mut source = CompoundFile::new(version).unwrap();
            source.create_storage("/Data").unwrap();
            source
                .create_stream("/Data/Small", b"small".to_vec())
                .unwrap();
            source
                .create_stream("/Data/Large", vec![0x11; 5_000])
                .unwrap();

            let mut compound = CompoundFile::from_bytes(&source.to_bytes().unwrap()).unwrap();
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
            compound.create_storage("relative").unwrap();
            assert!(compound.entry("/RELATIVE").unwrap().is_storage());
            assert!(compound.create_storage("/Missing/Child").is_err());
            assert!(compound.create_storage("/Regular/Child").is_err());
            assert!(compound.create_stream("/Data", Vec::new()).is_err());
            assert!(compound.create_stream("/data", Vec::new()).is_err());
            assert!(compound.create_stream("/Bad:Name", Vec::new()).is_err());
            compound.create_stream("/Bad\0Name", Vec::new()).unwrap();
            compound.create_stream("/Data/../Oops", Vec::new()).unwrap();
            assert!(compound.entry("/Oops").unwrap().is_stream());
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
            compound.remove_storage("RELATIVE").unwrap();
            compound.create_storage("/Vacant").unwrap();
            assert_eq!(
                compound.remove_entry("/Vacant").unwrap().kind,
                EntryKind::Storage
            );
            assert!(compound.remove_entry("/").is_err());

            let encoded = compound.to_bytes().unwrap();
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
        let mut bytes = CompoundFile::new(Version::V3).unwrap().to_bytes().unwrap();
        bytes.extend_from_slice(b"trailing");
        let parsed = CompoundFile::from_bytes(&bytes).unwrap();
        assert_eq!(parsed.trailing_data(), b"trailing");
        let output = parsed.to_bytes().unwrap();
        assert!(output.ends_with(b"trailing"));
        CompoundFile::from_bytes_strict(&output).unwrap();
        assert!(parsed.logical_eq(&CompoundFile::from_bytes(&output).unwrap()));
    }

    #[test]
    fn unallocated_physical_sectors_are_preserved() {
        let mut bytes = CompoundFile::new(Version::V3).unwrap().to_bytes().unwrap();
        bytes.extend_from_slice(&[0x5a; 512]);
        let parsed = CompoundFile::from_bytes(&bytes).unwrap();
        assert_eq!(parsed.unallocated_sectors(), [vec![0x5a; 512]]);
        let output = parsed.to_bytes().unwrap();
        let reopened = CompoundFile::from_bytes(&output).unwrap();
        assert!(parsed.logical_eq(&reopened));

        let mut many = CompoundFile::new(Version::V3).unwrap();
        many.unallocated_sectors = vec![vec![0x6b; 512]; 130];
        let output = many.to_bytes().unwrap();
        let reopened = CompoundFile::from_bytes_strict(&output).unwrap();
        assert_eq!(reopened.unallocated_sectors().len(), 130);
        assert!(many.logical_eq(&reopened));
    }

    #[test]
    fn strict_open_rejects_compatibility_only_header_and_directory_shapes() {
        let bytes = CompoundFile::new(Version::V3).unwrap().to_bytes().unwrap();
        CompoundFile::from_bytes_strict(&bytes).unwrap();

        let mut reserved = bytes.clone();
        reserved[34] = 1;
        assert!(CompoundFile::from_bytes(&reserved).is_ok());
        assert!(CompoundFile::from_bytes_strict(&reserved).is_err());

        let mut noncanonical_difat_end = bytes.clone();
        noncanonical_difat_end[68..72].copy_from_slice(&header::FREE_SECTOR.to_le_bytes());
        assert!(CompoundFile::from_bytes(&noncanonical_difat_end).is_ok());
        assert!(CompoundFile::from_bytes_strict(&noncanonical_difat_end).is_err());

        let header = Header::from_bytes(&bytes).unwrap();
        let directory_offset =
            header.sector_len() + header.first_directory_sector as usize * header.sector_len();
        let mut unterminated = bytes.clone();
        unterminated[directory_offset + "Root Entry".encode_utf16().count() * 2] = 1;
        assert!(CompoundFile::from_bytes(&unterminated).is_ok());
        assert!(CompoundFile::from_bytes_strict(&unterminated).is_err());

        let mut noncanonical_free_entry = bytes.clone();
        noncanonical_free_entry[directory_offset + 2 * directory::DIRECTORY_ENTRY_LEN + 67] = 1;
        assert!(CompoundFile::from_bytes(&noncanonical_free_entry).is_ok());
        assert!(CompoundFile::from_bytes_strict(&noncanonical_free_entry).is_err());

        let fat_sector = header.difat[0] as usize;
        let fat_offset = header.sector_len() + fat_sector * header.sector_len();
        let file_sector_count = bytes.len() / header.sector_len() - 1;
        let mut noncanonical_fat_padding = bytes.clone();
        let padding = fat_offset + file_sector_count * 4;
        noncanonical_fat_padding[padding..padding + 4]
            .copy_from_slice(&allocation::END_OF_CHAIN.to_le_bytes());
        assert!(CompoundFile::from_bytes(&noncanonical_fat_padding).is_ok());
        assert!(CompoundFile::from_bytes_strict(&noncanonical_fat_padding).is_err());

        let mut nonzero_root_creation = reserved;
        nonzero_root_creation[directory_offset + 100] = 1;
        let compatible = CompoundFile::from_bytes(&nonzero_root_creation).unwrap();
        assert_eq!(compatible.root_entry().created, FileTime::ZERO);
        assert_eq!(compatible.directory().root().creation_time.ticks(), 1);
        assert!(CompoundFile::from_bytes_strict(&nonzero_root_creation).is_err());
        CompoundFile::from_bytes_strict(&compatible.to_bytes().unwrap()).unwrap();

        let mut tree = CompoundFile::new(Version::V3).unwrap();
        for name in ["A", "B", "C", "D", "E", "F"] {
            tree.create_stream(name, Vec::new()).unwrap();
        }
        let mut transitive_order_violation = tree.to_bytes().unwrap();
        let canonical_tree = CompoundFile::from_bytes(&transitive_order_violation).unwrap();
        let header = Header::from_bytes(&transitive_order_violation).unwrap();
        let directory_offset =
            header.sector_len() + header.first_directory_sector as usize * header.sector_len();
        assert_eq!(
            canonical_tree.directory().entries()[3].raw_name().unwrap(),
            "C"
        );
        transitive_order_violation[directory_offset + 3 * directory::DIRECTORY_ENTRY_LEN] = b'E';
        assert!(CompoundFile::from_bytes(&transitive_order_violation).is_ok());
        assert!(CompoundFile::from_bytes_strict(&transitive_order_violation).is_err());

        let mut v4_padding = CompoundFile::new(Version::V4).unwrap().to_bytes().unwrap();
        v4_padding[header::HEADER_LEN] = 1;
        assert!(CompoundFile::from_bytes(&v4_padding).is_ok());
        assert!(CompoundFile::from_bytes_strict(&v4_padding).is_err());
    }

    #[test]
    fn storage_walk_uses_cfb_preorder() {
        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound.create_storage("/Data").unwrap();
        compound.create_stream("/Data/B", Vec::new()).unwrap();
        compound.create_storage("/Data/C").unwrap();
        compound.create_stream("/Data/C/X", Vec::new()).unwrap();
        compound.create_stream("/Data/AA", Vec::new()).unwrap();

        let paths: Vec<_> = compound
            .walk_storage("/data")
            .unwrap()
            .into_iter()
            .map(|entry| entry.path.as_path())
            .collect();
        assert_eq!(
            paths,
            ["/Data", "/Data/B", "/Data/C", "/Data/C/X", "/Data/AA"].map(Path::new)
        );
    }

    #[test]
    fn logical_api_matches_cfb_path_and_metadata_semantics() {
        let mut compound = CompoundFile::new(Version::V4).unwrap();
        compound.create_storage_all("Data/Nested/Leaf").unwrap();
        assert!(compound.is_storage("/data/NESTED/leaf"));
        assert_eq!(compound.children("DATA/NESTED").unwrap().len(), 1);

        assert_eq!(
            compound
                .create_or_replace_stream("data/nested/leaf/Value", b"first".to_vec())
                .unwrap(),
            None
        );
        assert_eq!(
            compound.stream("/DATA/NESTED/LEAF/value"),
            Some(b"first".as_slice())
        );
        let mut reader = compound.open_stream("data/nested/leaf/value").unwrap();
        let mut streamed = Vec::new();
        reader.read_to_end(&mut streamed).unwrap();
        assert_eq!(streamed, b"first");
        {
            let mut stream = compound.open_stream_mut("data/nested/leaf/value").unwrap();
            stream.seek(SeekFrom::End(0)).unwrap();
            stream.write_all(b"!").unwrap();
        }
        assert_eq!(
            compound.stream("data/nested/leaf/value"),
            Some(b"first!".as_slice())
        );
        assert_eq!(
            compound
                .create_or_replace_stream("/Data/Nested/Leaf/VALUE", b"second".to_vec())
                .unwrap(),
            Some(b"first!".to_vec())
        );

        let clsid = Guid::from_fields(0x1234_5678, 0x9abc, 0xdef0, [1, 2, 3, 4, 5, 6, 7, 8]);
        assert_eq!(
            compound
                .replace_storage_class_id("/data/nested", clsid)
                .unwrap(),
            Guid::ZERO
        );
        assert_eq!(
            compound.replace_state_bits("DATA/NESTED", 0x55aa).unwrap(),
            0
        );
        assert_eq!(
            compound
                .replace_creation_time("data/nested", FileTime::from_ticks(11))
                .unwrap(),
            FileTime::ZERO
        );
        assert_eq!(
            compound
                .replace_modified_time("/", FileTime::from_ticks(22))
                .unwrap(),
            FileTime::ZERO
        );
        assert!(
            compound
                .replace_creation_time("/", FileTime::from_ticks(1))
                .is_err()
        );
        assert!(
            compound
                .replace_modified_time("data/nested/leaf/value", FileTime::from_ticks(1))
                .is_err()
        );

        let walked = compound.walk_storage("/DATA/NESTED").unwrap();
        assert_eq!(walked.len(), 3);

        let mut bytes = Vec::new();
        compound.write_to(&mut bytes).unwrap();
        let mut reopened = CompoundFile::from_reader(bytes.as_slice()).unwrap();
        let nested = reopened.entry("/data/NESTED").unwrap();
        assert_eq!(nested.clsid, clsid);
        assert_eq!(nested.state_bits, 0x55aa);
        assert_eq!(nested.created, FileTime::from_ticks(11));
        assert_eq!(reopened.root_entry().modified, FileTime::from_ticks(22));
        assert_eq!(
            reopened.stream("data/nested/leaf/value"),
            Some(b"second".as_slice())
        );

        assert_eq!(
            reopened
                .remove_stream("DATA/NESTED/LEAF/VALUE")
                .unwrap()
                .data,
            b"second"
        );
        let removed = reopened.remove_storage_all("data/nested").unwrap();
        assert_eq!(removed.len(), 2);
        assert!(reopened.entry("/Data/Nested").is_none());
        assert!(reopened.entry("/Data").is_some());
    }
}
