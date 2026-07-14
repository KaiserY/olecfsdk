use std::{cmp::Ordering, collections::BTreeSet, io::Cursor, path::PathBuf};

use crate::{
    Error, Result, SdkEnum, SdkObject,
    common::{FileTime, Guid},
    io::{Reader, SdkRead, SdkSize, SdkWrite, Writer},
    limits::Limits,
};

use super::{SectorId, Version, name, sector::SectorSource};

pub const DIRECTORY_ENTRY_LEN: usize = 128;
pub const NO_STREAM: u32 = 0xffff_ffff;
pub const MAX_REGULAR_STREAM_ID: u32 = 0xffff_fffa;

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u8")]
pub enum DirectoryObjectType {
    Unallocated = 0,
    Storage = 1,
    Stream = 2,
    Root = 5,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u8")]
pub enum DirectoryColor {
    Red = 0,
    Black = 1,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DirectoryPointer {
    Entry(u32),
    None,
}

impl DirectoryPointer {
    pub fn from_raw(raw: u32) -> Result<Self> {
        match raw {
            0..=MAX_REGULAR_STREAM_ID => Ok(Self::Entry(raw)),
            NO_STREAM => Ok(Self::None),
            _ => Err(Error::invalid(
                0,
                format!("invalid directory stream ID 0x{raw:08x}"),
            )),
        }
    }

    pub fn raw(self) -> u32 {
        match self {
            Self::Entry(value) => value,
            Self::None => NO_STREAM,
        }
    }
}

impl SdkRead for DirectoryPointer {
    fn read_from<R: std::io::Read + std::io::Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let offset = reader.position()?;
        Self::from_raw(reader.read_u32()?)
            .map_err(|_| Error::invalid(offset, "invalid directory sibling or child stream ID"))
    }
}

impl SdkWrite for DirectoryPointer {
    fn write_to<W: std::io::Write + std::io::Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u32(self.raw())
    }
}

impl SdkSize for DirectoryPointer {
    fn sdk_size(&self) -> u64 {
        4
    }
}

/// Exact static representation of an MS-CFB 128-byte directory entry.
#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
#[sdk(validate = "validate_directory_entry")]
pub struct DirectoryEntry {
    pub name_buffer: [u16; 32],
    pub name_length: u16,
    pub object_type: DirectoryObjectType,
    pub color: DirectoryColor,
    pub left_sibling: DirectoryPointer,
    pub right_sibling: DirectoryPointer,
    pub child: DirectoryPointer,
    pub clsid: Guid,
    pub state_bits: u32,
    pub creation_time: FileTime,
    pub modified_time: FileTime,
    pub start_sector: u32,
    pub stream_size: u64,
}

impl DirectoryEntry {
    pub fn name(&self) -> Result<String> {
        if self.object_type == DirectoryObjectType::Root {
            return Ok("Root Entry".to_string());
        }
        let len = self.name_char_len()?;
        String::from_utf16(&self.name_buffer[..len])
            .map_err(|_| Error::invalid(0, "directory name is not valid UTF-16"))
    }

    pub fn raw_name(&self) -> Result<String> {
        let len = self.name_char_len()?;
        String::from_utf16(&self.name_buffer[..len])
            .map_err(|_| Error::invalid(0, "directory name is not valid UTF-16"))
    }

    pub fn effective_stream_size(&self, version: Version) -> u64 {
        match version {
            Version::V3 => self.stream_size & u32::MAX as u64,
            Version::V4 => self.stream_size,
        }
    }

    fn name_char_len(&self) -> Result<usize> {
        if self.name_length == 0 {
            return Ok(0);
        }
        if self.name_length > 64 || !self.name_length.is_multiple_of(2) {
            return Err(Error::invalid(64, "invalid CFB directory name length"));
        }
        Ok((self.name_length / 2 - 1) as usize)
    }
}

fn validate_directory_entry(entry: &DirectoryEntry) -> Result<()> {
    if entry.object_type != DirectoryObjectType::Unallocated {
        if entry.name_length < 2 {
            return Err(Error::invalid(64, "allocated directory entry has no name"));
        }
        entry.raw_name()?;
    } else if entry.name_length != 0 {
        entry.raw_name()?;
    }
    if entry.object_type == DirectoryObjectType::Stream && entry.child != DirectoryPointer::None {
        return Err(Error::invalid(76, "stream directory entry has a child"));
    }
    Ok(())
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Directory {
    sectors: Vec<SectorId>,
    entries: Vec<DirectoryEntry>,
    declared_sector_count: u32,
}

impl Directory {
    pub(crate) fn read(
        sectors: Vec<SectorId>,
        header_declared_sector_count: u32,
        source: &SectorSource<'_>,
        limits: Limits,
    ) -> Result<Self> {
        let entry_count = sectors
            .len()
            .checked_mul(source.sector_len() / DIRECTORY_ENTRY_LEN)
            .ok_or_else(|| Error::Limit("directory entry count overflow".into()))?;
        if entry_count > limits.max_entries {
            return Err(Error::Limit(format!(
                "directory entry count {entry_count} exceeds {}",
                limits.max_entries
            )));
        }
        let mut entries = Vec::with_capacity(entry_count);
        for &sector in &sectors {
            for raw in source
                .full_sector(sector)?
                .chunks_exact(DIRECTORY_ENTRY_LEN)
            {
                let mut reader = Reader::new(Cursor::new(raw))?;
                let entry = DirectoryEntry::read_from(&mut reader)?;
                debug_assert_eq!(entry.sdk_size(), DIRECTORY_ENTRY_LEN as u64);
                entries.push(entry);
            }
        }
        let directory = Self {
            sectors,
            entries,
            declared_sector_count: header_declared_sector_count,
        };
        directory.validate_tree()?;
        Ok(directory)
    }

    pub fn sectors(&self) -> &[SectorId] {
        &self.sectors
    }

    pub fn entries(&self) -> &[DirectoryEntry] {
        &self.entries
    }

    pub fn root(&self) -> &DirectoryEntry {
        &self.entries[0]
    }

    pub fn declared_sector_count(&self) -> u32 {
        self.declared_sector_count
    }

    pub(crate) fn paths(&self) -> Result<Vec<(u32, PathBuf)>> {
        enum Task {
            Siblings(DirectoryPointer, PathBuf),
            Entry(u32, PathBuf),
        }

        let mut output = vec![(0, PathBuf::from("/"))];
        let mut tasks = vec![Task::Siblings(self.root().child, PathBuf::from("/"))];
        while let Some(task) = tasks.pop() {
            match task {
                Task::Siblings(DirectoryPointer::None, _) => {}
                Task::Siblings(DirectoryPointer::Entry(start), parent) => {
                    let mut ordered = Vec::new();
                    let mut stack = Vec::new();
                    let mut current = Some(start);
                    while current.is_some() || !stack.is_empty() {
                        while let Some(id) = current {
                            stack.push(id);
                            current = match self.entries[id as usize].left_sibling {
                                DirectoryPointer::Entry(left) => Some(left),
                                DirectoryPointer::None => None,
                            };
                        }
                        let id = stack.pop().unwrap();
                        ordered.push(id);
                        current = match self.entries[id as usize].right_sibling {
                            DirectoryPointer::Entry(right) => Some(right),
                            DirectoryPointer::None => None,
                        };
                    }
                    for id in ordered.into_iter().rev() {
                        tasks.push(Task::Entry(id, parent.clone()));
                    }
                }
                Task::Entry(id, parent) => {
                    let entry = &self.entries[id as usize];
                    let path = parent.join(entry.name()?);
                    output.push((id, path.clone()));
                    if entry.object_type == DirectoryObjectType::Storage {
                        tasks.push(Task::Siblings(entry.child, path));
                    }
                }
            }
        }
        Ok(output)
    }

    pub(crate) fn validate_strict(&self, version: Version) -> Result<()> {
        if version == Version::V3 && self.declared_sector_count != 0 {
            return Err(Error::invalid(
                40,
                "CFB v3 directory sector count must be zero",
            ));
        }
        if version == Version::V4 && self.declared_sector_count as usize != self.sectors.len() {
            return Err(Error::invalid(
                40,
                "CFB v4 directory sector count does not match its chain",
            ));
        }

        for (id, entry) in self.entries.iter().enumerate() {
            if entry.object_type == DirectoryObjectType::Unallocated {
                if entry.name_buffer != [0; 32]
                    || entry.name_length != 0
                    || entry.color != DirectoryColor::Red
                    || entry.left_sibling != DirectoryPointer::None
                    || entry.right_sibling != DirectoryPointer::None
                    || entry.child != DirectoryPointer::None
                    || !entry.clsid.is_zero()
                    || entry.state_bits != 0
                    || entry.creation_time != FileTime::ZERO
                    || entry.modified_time != FileTime::ZERO
                    || entry.start_sector != 0
                    || entry.stream_size != 0
                {
                    return Err(Error::invalid(
                        id as u64 * DIRECTORY_ENTRY_LEN as u64,
                        "unallocated CFB directory entries must have canonical zero fields",
                    ));
                }
                continue;
            }
            let raw_name = entry.raw_name()?;
            let terminator = entry.name_char_len()?;
            if entry.name_buffer[terminator] != 0 {
                return Err(Error::invalid(
                    id as u64 * DIRECTORY_ENTRY_LEN as u64,
                    "CFB directory name is not NUL-terminated",
                ));
            }
            match entry.object_type {
                DirectoryObjectType::Root => {
                    if id != 0 || raw_name != "Root Entry" {
                        return Err(Error::invalid(
                            id as u64 * DIRECTORY_ENTRY_LEN as u64,
                            "CFB root entry has the wrong ID or name",
                        ));
                    }
                    if entry.left_sibling != DirectoryPointer::None
                        || entry.right_sibling != DirectoryPointer::None
                    {
                        return Err(Error::invalid(
                            id as u64 * DIRECTORY_ENTRY_LEN as u64 + 68,
                            "CFB root entry cannot have siblings",
                        ));
                    }
                    if entry.creation_time != FileTime::ZERO {
                        return Err(Error::invalid(
                            id as u64 * DIRECTORY_ENTRY_LEN as u64 + 100,
                            "CFB root creation time must be zero",
                        ));
                    }
                }
                DirectoryObjectType::Storage => {
                    name::validate_entry_name(&raw_name)?;
                    if entry.start_sector != 0 || entry.stream_size != 0 {
                        return Err(Error::invalid(
                            id as u64 * DIRECTORY_ENTRY_LEN as u64 + 116,
                            "CFB storage start sector and stream size must be zero",
                        ));
                    }
                }
                DirectoryObjectType::Stream => {
                    name::validate_entry_name(&raw_name)?;
                    if !entry.clsid.is_zero()
                        || entry.creation_time != FileTime::ZERO
                        || entry.modified_time != FileTime::ZERO
                    {
                        return Err(Error::invalid(
                            id as u64 * DIRECTORY_ENTRY_LEN as u64 + 80,
                            "CFB stream CLSID and timestamps must be zero",
                        ));
                    }
                }
                DirectoryObjectType::Unallocated => unreachable!(),
            }
            if version == Version::V3
                && matches!(
                    entry.object_type,
                    DirectoryObjectType::Root | DirectoryObjectType::Stream
                )
                && (entry.stream_size > 0x8000_0000 || entry.stream_size >> 32 != 0)
            {
                return Err(Error::invalid(
                    id as u64 * DIRECTORY_ENTRY_LEN as u64 + 120,
                    "CFB v3 stream size must fit the specified 2 GiB range",
                ));
            }
        }

        let mut visited = BTreeSet::new();
        self.validate_strict_node(0, false, None, None, &mut visited)?;
        if self.entries.iter().enumerate().any(|(id, entry)| {
            entry.object_type != DirectoryObjectType::Unallocated && !visited.contains(&(id as u32))
        }) {
            return Err(Error::invalid(
                0,
                "CFB directory contains an allocated entry outside the hierarchy",
            ));
        }
        Ok(())
    }

    fn validate_strict_node(
        &self,
        id: u32,
        parent_is_red: bool,
        lower_bound: Option<u32>,
        upper_bound: Option<u32>,
        visited: &mut BTreeSet<u32>,
    ) -> Result<()> {
        if !visited.insert(id) {
            return Err(Error::invalid(0, "CFB directory tree contains a cycle"));
        }
        let entry = self
            .entries
            .get(id as usize)
            .ok_or_else(|| Error::invalid(0, "CFB directory pointer is out of bounds"))?;
        let is_red = entry.color == DirectoryColor::Red;
        if parent_is_red && is_red {
            return Err(Error::invalid(
                id as u64 * DIRECTORY_ENTRY_LEN as u64 + 67,
                "CFB directory tree contains adjacent red nodes",
            ));
        }

        let entry_name = entry.name()?;
        for (bound, required) in [
            (lower_bound, Ordering::Greater),
            (upper_bound, Ordering::Less),
        ] {
            let Some(bound) = bound else { continue };
            let bound_entry = self
                .entries
                .get(bound as usize)
                .ok_or_else(|| Error::invalid(0, "CFB directory bound is out of bounds"))?;
            if name::compare_names(&entry_name, &bound_entry.name()?) != required {
                return Err(Error::invalid(
                    id as u64 * DIRECTORY_ENTRY_LEN as u64,
                    "CFB directory sibling names are not in MS-CFB order or are not unique",
                ));
            }
        }

        for (pointer, expected) in [
            (entry.left_sibling, Ordering::Less),
            (entry.right_sibling, Ordering::Greater),
        ] {
            let DirectoryPointer::Entry(sibling) = pointer else {
                continue;
            };
            let sibling_entry = self
                .entries
                .get(sibling as usize)
                .ok_or_else(|| Error::invalid(0, "CFB sibling pointer is out of bounds"))?;
            let ordering = name::compare_names(&sibling_entry.name()?, &entry.name()?);
            if ordering != expected {
                return Err(Error::invalid(
                    sibling as u64 * DIRECTORY_ENTRY_LEN as u64,
                    "CFB directory sibling names are not in MS-CFB order",
                ));
            }
            let (lower, upper) = if expected == Ordering::Less {
                (lower_bound, Some(id))
            } else {
                (Some(id), upper_bound)
            };
            self.validate_strict_node(sibling, is_red, lower, upper, visited)?;
        }
        if let DirectoryPointer::Entry(child) = entry.child {
            self.validate_strict_node(child, false, None, None, visited)?;
        }
        Ok(())
    }

    fn validate_tree(&self) -> Result<()> {
        let root = self
            .entries
            .first()
            .ok_or_else(|| Error::invalid(0, "CFB directory has no root entry"))?;
        if root.object_type != DirectoryObjectType::Root {
            return Err(Error::invalid(0, "directory entry zero is not the root"));
        }
        if !root.effective_stream_size(Version::V4).is_multiple_of(64) {
            return Err(Error::invalid(
                120,
                "root mini stream size is not 64-byte aligned",
            ));
        }

        let mut visited = BTreeSet::new();
        let mut stack = vec![0u32];
        while let Some(id) = stack.pop() {
            if !visited.insert(id) {
                return Err(Error::invalid(0, "directory tree contains a cycle"));
            }
            let entry = self
                .entries
                .get(id as usize)
                .ok_or_else(|| Error::invalid(0, "directory tree pointer is out of bounds"))?;
            if id != 0
                && !matches!(
                    entry.object_type,
                    DirectoryObjectType::Storage | DirectoryObjectType::Stream
                )
            {
                return Err(Error::invalid(
                    0,
                    "reachable non-root directory entry has an invalid object type",
                ));
            }
            for pointer in [entry.left_sibling, entry.right_sibling, entry.child] {
                if let DirectoryPointer::Entry(next) = pointer {
                    if next as usize >= self.entries.len() {
                        return Err(Error::invalid(0, "directory tree pointer is out of bounds"));
                    }
                    stack.push(next);
                }
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use crate::io::{SdkRead, SdkWrite};

    use super::*;

    fn root() -> DirectoryEntry {
        let mut name_buffer = [0; 32];
        for (target, source) in name_buffer.iter_mut().zip("Root Entry".encode_utf16()) {
            *target = source;
        }
        DirectoryEntry {
            name_buffer,
            name_length: 22,
            object_type: DirectoryObjectType::Root,
            color: DirectoryColor::Black,
            left_sibling: DirectoryPointer::None,
            right_sibling: DirectoryPointer::None,
            child: DirectoryPointer::None,
            clsid: Guid::ZERO,
            state_bits: 0,
            creation_time: FileTime::ZERO,
            modified_time: FileTime::ZERO,
            start_sector: super::super::allocation::END_OF_CHAIN,
            stream_size: 0,
        }
    }

    #[test]
    fn directory_entry_is_exactly_128_bytes() {
        let entry = root();
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        entry.write_to(&mut writer).unwrap();
        let bytes = writer.into_inner().into_inner();
        assert_eq!(bytes.len(), DIRECTORY_ENTRY_LEN);
        let mut reader = Reader::new(Cursor::new(bytes)).unwrap();
        assert_eq!(DirectoryEntry::read_from(&mut reader).unwrap(), entry);
        assert_eq!(entry.name().unwrap(), "Root Entry");
    }
}
