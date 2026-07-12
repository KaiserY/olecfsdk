use std::path::{Path, PathBuf};

use uuid::Uuid;
use web_time::SystemTime;

use crate::{Error, Result, limits::Limits};

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
    pub clsid: Uuid,
    pub state_bits: u32,
    pub created: SystemTime,
    pub modified: SystemTime,
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
