use std::{
    io::{Cursor, Read, Write},
    path::{Path, PathBuf},
};

use uuid::Uuid;
use web_time::SystemTime;

use crate::{Error, Result, limits::Limits};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Version {
    V3,
    V4,
}

impl From<cfb::Version> for Version {
    fn from(value: cfb::Version) -> Self {
        match value {
            cfb::Version::V3 => Self::V3,
            cfb::Version::V4 => Self::V4,
        }
    }
}

impl From<Version> for cfb::Version {
    fn from(value: Version) -> Self {
        match value {
            Version::V3 => Self::V3,
            Version::V4 => Self::V4,
        }
    }
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
    entries: Vec<Entry>,
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
        let mut compatible_bytes = bytes.to_vec();
        normalize_unaddressable_fat_padding(&mut compatible_bytes);
        let mut source = cfb::CompoundFile::open(Cursor::new(compatible_bytes))?;
        let version = source.version().into();
        let metadata: Vec<_> = source.walk().collect();
        if metadata.len() > limits.max_entries {
            return Err(Error::Limit(format!(
                "entry count {} exceeds {}",
                metadata.len(),
                limits.max_entries
            )));
        }

        let mut entries = Vec::with_capacity(metadata.len());
        for meta in metadata {
            if meta.len() > limits.max_stream_size {
                return Err(Error::Limit(format!(
                    "stream {} length {} exceeds {}",
                    meta.path().display(),
                    meta.len(),
                    limits.max_stream_size
                )));
            }
            let kind = if meta.is_root() {
                EntryKind::Root
            } else if meta.is_stream() {
                EntryKind::Stream
            } else {
                EntryKind::Storage
            };
            let mut data = Vec::new();
            if meta.is_stream() {
                let capacity = usize::try_from(meta.len())
                    .map_err(|_| Error::Limit("stream length does not fit usize".into()))?;
                if capacity > limits.max_allocation {
                    return Err(Error::Limit(format!(
                        "stream allocation {capacity} exceeds {}",
                        limits.max_allocation
                    )));
                }
                data.reserve_exact(capacity);
                source.open_stream(meta.path())?.read_to_end(&mut data)?;
            }
            entries.push(Entry {
                path: meta.path().to_path_buf(),
                name: meta.name().to_string(),
                kind,
                clsid: *meta.clsid(),
                state_bits: meta.state_bits(),
                created: meta.created(),
                modified: meta.modified(),
                data,
            });
        }
        entries.sort_by(|left, right| left.path.cmp(&right.path));
        Ok(Self { version, entries })
    }

    pub fn version(&self) -> Version {
        self.version
    }
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }
    pub fn entry(&self, path: impl AsRef<Path>) -> Option<&Entry> {
        self.entries
            .iter()
            .find(|entry| entry.path == path.as_ref())
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let cursor = Cursor::new(Vec::new());
        let mut target = cfb::CompoundFile::create_with_version(self.version.into(), cursor)?;

        let mut storages: Vec<_> = self
            .entries
            .iter()
            .filter(|entry| entry.kind == EntryKind::Storage)
            .collect();
        storages.sort_by_key(|entry| entry.path.components().count());
        for entry in storages {
            target.create_storage(&entry.path)?;
        }

        for entry in self.entries.iter().filter(|entry| entry.is_stream()) {
            let mut stream = target.create_new_stream(&entry.path)?;
            stream.write_all(&entry.data)?;
        }

        for entry in self.entries.iter().filter(|entry| entry.is_storage()) {
            target.set_storage_clsid(&entry.path, entry.clsid)?;
            target.set_state_bits(&entry.path, entry.state_bits)?;
            target.set_created_time(&entry.path, entry.created)?;
            target.set_modified_time(&entry.path, entry.modified)?;
        }
        for entry in self.entries.iter().filter(|entry| entry.is_stream()) {
            target.set_state_bits(&entry.path, entry.state_bits)?;
        }
        target.flush()?;
        Ok(target.into_inner().into_inner())
    }

    pub fn logical_eq(&self, other: &Self) -> bool {
        self.version == other.version && self.entries == other.entries
    }
}

const CFB_MAGIC: [u8; 8] = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREE_SECTOR: u32 = 0xffff_ffff;
const END_OF_CHAIN: u32 = 0xffff_fffe;
const MAX_REGULAR_SECTOR: u32 = 0xffff_fffa;

/// Real-world writers sometimes leave garbage in FAT slots whose sector IDs
/// are beyond EOF. Those slots cannot be reached by any valid chain. Office
/// and POI ignore them, so compatibility mode canonicalizes only that unused
/// padding before handing the image to the strict allocator implementation.
fn normalize_unaddressable_fat_padding(bytes: &mut [u8]) {
    if bytes.len() < 512 || bytes[..8] != CFB_MAGIC {
        return;
    }
    let Some(sector_shift) = read_u16_at(bytes, 30) else {
        return;
    };
    if !matches!(sector_shift, 9 | 12) {
        return;
    }
    let sector_len = 1usize << sector_shift;
    if bytes.len() < sector_len || !bytes.len().is_multiple_of(sector_len) {
        return;
    }
    let num_sectors = bytes.len() / sector_len - 1;
    let Some(num_fat_sectors) = read_u32_at(bytes, 44).map(|value| value as usize) else {
        return;
    };
    let Some(mut next_difat) = read_u32_at(bytes, 68) else {
        return;
    };
    let Some(num_difat_sectors) = read_u32_at(bytes, 72).map(|value| value as usize) else {
        return;
    };

    let mut fat_sectors = Vec::with_capacity(num_fat_sectors.min(num_sectors));
    for index in 0..109 {
        let Some(value) = read_u32_at(bytes, 76 + index * 4) else {
            return;
        };
        if value <= MAX_REGULAR_SECTOR {
            fat_sectors.push(value);
            if fat_sectors.len() == num_fat_sectors {
                break;
            }
        }
    }

    let mut seen = std::collections::BTreeSet::new();
    for _ in 0..num_difat_sectors.min(num_sectors) {
        if fat_sectors.len() == num_fat_sectors || next_difat > MAX_REGULAR_SECTOR {
            break;
        }
        if next_difat as usize >= num_sectors || !seen.insert(next_difat) {
            return;
        }
        let offset = (next_difat as usize + 1) * sector_len;
        for index in 0..(sector_len / 4 - 1) {
            let Some(value) = read_u32_at(bytes, offset + index * 4) else {
                return;
            };
            if value <= MAX_REGULAR_SECTOR {
                fat_sectors.push(value);
                if fat_sectors.len() == num_fat_sectors {
                    break;
                }
            }
        }
        let Some(value) = read_u32_at(bytes, offset + sector_len - 4) else {
            return;
        };
        next_difat = value;
        if matches!(next_difat, END_OF_CHAIN | FREE_SECTOR) {
            break;
        }
    }

    let entries_per_sector = sector_len / 4;
    for (table_index, sector_id) in fat_sectors.into_iter().enumerate() {
        if sector_id as usize >= num_sectors {
            return;
        }
        let offset = (sector_id as usize + 1) * sector_len;
        for slot in 0..entries_per_sector {
            if table_index * entries_per_sector + slot >= num_sectors {
                bytes[offset + slot * 4..offset + slot * 4 + 4]
                    .copy_from_slice(&FREE_SECTOR.to_le_bytes());
            }
        }
    }
}

fn read_u16_at(bytes: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        bytes.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn read_u32_at(bytes: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        bytes.get(offset..offset + 4)?.try_into().ok()?,
    ))
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
        for version in [cfb::Version::V3, cfb::Version::V4] {
            let source = cfb::CompoundFile::create_with_version(version, Cursor::new(Vec::new()))
                .unwrap()
                .into_inner()
                .into_inner();
            let output = round_trip_bytes(&source).unwrap();
            let parsed = CompoundFile::from_bytes(&output).unwrap();
            assert_eq!(parsed.version(), version.into());
        }
    }

    #[test]
    fn nested_streams_round_trip() {
        let mut source = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
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
        let parsed = CompoundFile::from_bytes(&output).unwrap();
        assert_eq!(parsed.entry("/WordDocument").unwrap().data, b"word");
        assert_eq!(parsed.entry("/Macros/VBA/dir").unwrap().data, b"vba");
    }
}
