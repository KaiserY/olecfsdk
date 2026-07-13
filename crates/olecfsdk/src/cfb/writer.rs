use std::{cmp::Ordering, collections::BTreeMap, io::Cursor, path::Path};

use crate::{
    Error, Result,
    common::{FileTime, Guid},
    io::{SdkWrite, Writer},
};

use super::{
    CompoundFile, Entry, EntryKind, Version,
    allocation::{DIFAT_SECTOR, END_OF_CHAIN, FAT_SECTOR},
    directory::{
        DIRECTORY_ENTRY_LEN, DirectoryColor, DirectoryEntry, DirectoryObjectType, DirectoryPointer,
    },
    header::{BYTE_ORDER_LE, FREE_SECTOR, Header, MAGIC, MINI_SECTOR_SHIFT, MINI_STREAM_CUTOFF},
};

const MINI_SECTOR_LEN: usize = 64;
const HEADER_DIFAT_LEN: usize = 109;

pub(crate) fn write_compound(compound: &CompoundFile) -> Result<Vec<u8>> {
    write_logical_compound(
        compound.version,
        &compound.entries,
        &compound.unallocated_sectors,
        &compound.trailing_data,
    )
}

pub(crate) fn write_empty_compound(version: Version) -> Result<Vec<u8>> {
    let entries = [Entry {
        path: "/".into(),
        name: "Root Entry".into(),
        kind: EntryKind::Root,
        clsid: Guid::ZERO,
        state_bits: 0,
        created: FileTime::ZERO,
        modified: FileTime::ZERO,
        data: Vec::new(),
    }];
    write_logical_compound(version, &entries, &[], &[])
}

fn write_logical_compound(
    version: Version,
    entries: &[Entry],
    unallocated_sectors: &[Vec<u8>],
    trailing_data: &[u8],
) -> Result<Vec<u8>> {
    let sector_len = match version {
        Version::V3 => 512,
        Version::V4 => 4096,
    };
    let ordered = ordered_entries(entries)?;
    let mut starts = vec![END_OF_CHAIN; ordered.len()];
    let mut mini_starts = vec![END_OF_CHAIN; ordered.len()];
    let mut mini_stream = Vec::new();
    let mut mini_fat = Vec::new();

    for (index, entry) in ordered.iter().enumerate() {
        if entry.kind != EntryKind::Stream
            || entry.data.is_empty()
            || entry.data.len() >= MINI_STREAM_CUTOFF as usize
        {
            continue;
        }
        let start = u32_len(mini_fat.len(), "mini-sector count")?;
        mini_starts[index] = start;
        let count = div_ceil(entry.data.len(), MINI_SECTOR_LEN);
        for offset in 0..count {
            let current = start
                .checked_add(offset as u32)
                .ok_or_else(|| Error::Limit("mini-sector ID overflow".into()))?;
            mini_fat.push(if offset + 1 == count {
                END_OF_CHAIN
            } else {
                current + 1
            });
            let begin = offset * MINI_SECTOR_LEN;
            let end = (begin + MINI_SECTOR_LEN).min(entry.data.len());
            mini_stream.extend_from_slice(&entry.data[begin..end]);
            mini_stream.resize(mini_stream.len().next_multiple_of(MINI_SECTOR_LEN), 0);
        }
    }

    let mut sectors = Vec::<Vec<u8>>::new();
    let mut chains = Vec::<(u32, usize)>::new();
    let root_mini_start = append_payload(&mut sectors, &mut chains, &mini_stream, sector_len)?;

    let mini_fat_entries_per_sector = sector_len / 4;
    if !mini_fat.is_empty() {
        mini_fat.resize(
            mini_fat.len().next_multiple_of(mini_fat_entries_per_sector),
            FREE_SECTOR,
        );
    }
    let mut mini_fat_bytes = Vec::with_capacity(mini_fat.len() * 4);
    for value in &mini_fat {
        mini_fat_bytes.extend_from_slice(&value.to_le_bytes());
    }
    let mini_fat_start = append_payload(&mut sectors, &mut chains, &mini_fat_bytes, sector_len)?;
    let mini_fat_sector_count = if mini_fat_bytes.is_empty() {
        0
    } else {
        mini_fat_bytes.len() / sector_len
    };

    for (index, entry) in ordered.iter().enumerate() {
        if entry.kind == EntryKind::Stream && entry.data.len() >= MINI_STREAM_CUTOFF as usize {
            starts[index] = append_payload(&mut sectors, &mut chains, &entry.data, sector_len)?;
        } else if entry.kind == EntryKind::Stream {
            starts[index] = mini_starts[index];
        }
    }

    let directory_entries =
        build_directory_entries(&ordered, &starts, root_mini_start, mini_stream.len() as u64)?;
    let mut directory_bytes = Vec::new();
    for entry in &directory_entries {
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        entry.write_to(&mut writer)?;
        directory_bytes.extend_from_slice(&writer.into_inner().into_inner());
    }
    let entries_per_sector = sector_len / DIRECTORY_ENTRY_LEN;
    while !(directory_bytes.len() / DIRECTORY_ENTRY_LEN).is_multiple_of(entries_per_sector) {
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        unallocated_directory_entry().write_to(&mut writer)?;
        directory_bytes.extend_from_slice(&writer.into_inner().into_inner());
    }
    let directory_start = append_payload(&mut sectors, &mut chains, &directory_bytes, sector_len)?;
    let directory_sector_count = directory_bytes.len() / sector_len;

    let data_sector_count = sectors.len();
    let fat_entries_per_sector = sector_len / 4;
    let difat_entries_per_sector = fat_entries_per_sector - 1;
    let (fat_sector_count, difat_sector_count) = allocation_table_counts(
        data_sector_count,
        fat_entries_per_sector,
        difat_entries_per_sector,
    )?;
    let difat_start_index = data_sector_count;
    let fat_start_index = difat_start_index + difat_sector_count;
    let total_sector_count = fat_start_index + fat_sector_count;
    let mut fat = vec![FREE_SECTOR; fat_sector_count * fat_entries_per_sector];
    for &(start, count) in &chains {
        mark_chain(&mut fat, start, count)?;
    }
    for entry in fat.iter_mut().take(fat_start_index).skip(difat_start_index) {
        *entry = DIFAT_SECTOR;
    }
    for entry in fat
        .iter_mut()
        .take(total_sector_count)
        .skip(fat_start_index)
    {
        *entry = FAT_SECTOR;
    }

    let fat_sector_ids: Vec<u32> = (fat_start_index..total_sector_count)
        .map(|value| u32_len(value, "FAT sector ID"))
        .collect::<Result<_>>()?;
    let difat_sector_ids: Vec<u32> = (difat_start_index..fat_start_index)
        .map(|value| u32_len(value, "DIFAT sector ID"))
        .collect::<Result<_>>()?;
    let mut difat_sectors = Vec::with_capacity(difat_sector_count);
    let remaining_fat = &fat_sector_ids[HEADER_DIFAT_LEN.min(fat_sector_ids.len())..];
    for (index, &sector_id) in difat_sector_ids.iter().enumerate() {
        let mut values = vec![FREE_SECTOR; fat_entries_per_sector];
        let begin = index * difat_entries_per_sector;
        let end = (begin + difat_entries_per_sector).min(remaining_fat.len());
        values[..end - begin].copy_from_slice(&remaining_fat[begin..end]);
        values[difat_entries_per_sector] = difat_sector_ids
            .get(index + 1)
            .copied()
            .unwrap_or(END_OF_CHAIN);
        let bytes = u32_sector(&values, sector_len);
        debug_assert_eq!(sector_id as usize, data_sector_count + index);
        difat_sectors.push(bytes);
    }
    let fat_sectors: Vec<_> = fat
        .chunks_exact(fat_entries_per_sector)
        .map(|values| u32_sector(values, sector_len))
        .collect();

    let mut header_difat = [FREE_SECTOR; HEADER_DIFAT_LEN];
    let initial_count = fat_sector_ids.len().min(HEADER_DIFAT_LEN);
    header_difat[..initial_count].copy_from_slice(&fat_sector_ids[..initial_count]);
    let header = Header {
        signature: MAGIC,
        clsid: [0; 16],
        minor_version: 0x003e,
        major_version: match version {
            Version::V3 => 3,
            Version::V4 => 4,
        },
        byte_order: BYTE_ORDER_LE,
        sector_shift: match version {
            Version::V3 => 9,
            Version::V4 => 12,
        },
        mini_sector_shift: MINI_SECTOR_SHIFT,
        reserved: [0; 6],
        number_of_directory_sectors: if version == Version::V4 {
            u32_len(directory_sector_count, "directory sector count")?
        } else {
            0
        },
        number_of_fat_sectors: u32_len(fat_sector_count, "FAT sector count")?,
        first_directory_sector: directory_start,
        transaction_signature: 0,
        mini_stream_cutoff: MINI_STREAM_CUTOFF,
        first_mini_fat_sector: mini_fat_start,
        number_of_mini_fat_sectors: u32_len(mini_fat_sector_count, "MiniFAT sector count")?,
        first_difat_sector: difat_sector_ids.first().copied().unwrap_or(END_OF_CHAIN),
        number_of_difat_sectors: u32_len(difat_sector_count, "DIFAT sector count")?,
        difat: header_difat,
    };

    let mut writer = Writer::new(Cursor::new(Vec::new()));
    header.write_to(&mut writer)?;
    let mut output = writer.into_inner().into_inner();
    output.resize(sector_len, 0);
    for sector in sectors
        .iter()
        .chain(difat_sectors.iter())
        .chain(fat_sectors.iter())
    {
        output.extend_from_slice(sector);
    }
    for sector in unallocated_sectors {
        if sector.len() != sector_len {
            return Err(Error::invalid(0, "unallocated sector has the wrong size"));
        }
        output.extend_from_slice(sector);
    }
    output.extend_from_slice(trailing_data);
    Ok(output)
}

fn ordered_entries(entries: &[Entry]) -> Result<Vec<&Entry>> {
    let root = entries
        .iter()
        .find(|entry| entry.kind == EntryKind::Root && entry.path == Path::new("/"))
        .ok_or_else(|| Error::invalid(0, "logical CFB model has no root entry"))?;
    if entries
        .iter()
        .filter(|entry| entry.kind == EntryKind::Root)
        .count()
        != 1
    {
        return Err(Error::invalid(0, "logical CFB model must contain one root"));
    }
    let mut rest: Vec<_> = entries
        .iter()
        .filter(|entry| entry.kind != EntryKind::Root)
        .collect();
    rest.sort_by(|left, right| left.path.cmp(&right.path));
    let mut ordered = Vec::with_capacity(entries.len());
    ordered.push(root);
    ordered.extend(rest);
    Ok(ordered)
}

fn build_directory_entries(
    entries: &[&Entry],
    starts: &[u32],
    root_mini_start: u32,
    root_mini_len: u64,
) -> Result<Vec<DirectoryEntry>> {
    let mut ids = BTreeMap::new();
    for (index, entry) in entries.iter().enumerate() {
        if ids.insert(entry.path.clone(), index as u32).is_some() {
            return Err(Error::invalid(0, "duplicate logical CFB path"));
        }
    }
    let mut records = Vec::with_capacity(entries.len());
    for (index, entry) in entries.iter().enumerate() {
        let (object_type, start_sector, stream_size) = match entry.kind {
            EntryKind::Root => (DirectoryObjectType::Root, root_mini_start, root_mini_len),
            EntryKind::Storage => (DirectoryObjectType::Storage, 0, 0),
            EntryKind::Stream => (
                DirectoryObjectType::Stream,
                starts[index],
                entry.data.len() as u64,
            ),
        };
        records.push(DirectoryEntry {
            name_buffer: encode_name(if entry.kind == EntryKind::Root {
                "Root Entry"
            } else {
                &entry.name
            })?,
            name_length: name_length(if entry.kind == EntryKind::Root {
                "Root Entry"
            } else {
                &entry.name
            })?,
            object_type,
            color: DirectoryColor::Black,
            left_sibling: DirectoryPointer::None,
            right_sibling: DirectoryPointer::None,
            child: DirectoryPointer::None,
            clsid: if entry.kind == EntryKind::Stream {
                Guid::ZERO
            } else {
                entry.clsid
            },
            state_bits: entry.state_bits,
            creation_time: if entry.kind == EntryKind::Stream {
                FileTime::ZERO
            } else {
                entry.created
            },
            modified_time: if entry.kind == EntryKind::Stream {
                FileTime::ZERO
            } else {
                entry.modified
            },
            start_sector,
            stream_size,
        });
    }

    let mut children: BTreeMap<u32, Vec<u32>> = BTreeMap::new();
    for (index, entry) in entries.iter().enumerate().skip(1) {
        let parent = entry
            .path
            .parent()
            .and_then(|path| ids.get(path))
            .copied()
            .ok_or_else(|| {
                Error::invalid(0, format!("missing parent for {}", entry.path.display()))
            })?;
        if entries[parent as usize].kind == EntryKind::Stream {
            return Err(Error::invalid(
                0,
                "stream cannot contain directory children",
            ));
        }
        children.entry(parent).or_default().push(index as u32);
    }
    for (parent, child_ids) in &mut children {
        child_ids.sort_by(|left, right| {
            compare_names(
                &entries[*left as usize].name,
                &entries[*right as usize].name,
            )
        });
        for pair in child_ids.windows(2) {
            if compare_names(
                &entries[pair[0] as usize].name,
                &entries[pair[1] as usize].name,
            ) == Ordering::Equal
            {
                return Err(Error::invalid(0, "duplicate case-insensitive CFB name"));
            }
        }
        records[*parent as usize].child = build_sibling_tree(child_ids, &mut records);
    }
    Ok(records)
}

fn build_sibling_tree(ids: &[u32], records: &mut [DirectoryEntry]) -> DirectoryPointer {
    fn build(
        ids: &[u32],
        records: &mut [DirectoryEntry],
        depth: usize,
        depths: &mut Vec<(u32, usize)>,
    ) -> DirectoryPointer {
        if ids.is_empty() {
            return DirectoryPointer::None;
        }
        let middle = ids.len() / 2;
        let id = ids[middle];
        let left = build(&ids[..middle], records, depth + 1, depths);
        let right = build(&ids[middle + 1..], records, depth + 1, depths);
        records[id as usize].left_sibling = left;
        records[id as usize].right_sibling = right;
        depths.push((id, depth));
        DirectoryPointer::Entry(id)
    }

    let mut depths = Vec::new();
    let root = build(ids, records, 0, &mut depths);
    let max_depth = depths.iter().map(|(_, depth)| *depth).max().unwrap_or(0);
    if max_depth > 0 {
        for (id, depth) in depths {
            records[id as usize].color = if depth == max_depth {
                DirectoryColor::Red
            } else {
                DirectoryColor::Black
            };
        }
    }
    root
}

fn append_payload(
    sectors: &mut Vec<Vec<u8>>,
    chains: &mut Vec<(u32, usize)>,
    bytes: &[u8],
    sector_len: usize,
) -> Result<u32> {
    if bytes.is_empty() {
        return Ok(END_OF_CHAIN);
    }
    let start = u32_len(sectors.len(), "sector ID")?;
    let count = div_ceil(bytes.len(), sector_len);
    for index in 0..count {
        let begin = index * sector_len;
        let end = (begin + sector_len).min(bytes.len());
        let mut sector = vec![0; sector_len];
        sector[..end - begin].copy_from_slice(&bytes[begin..end]);
        sectors.push(sector);
    }
    chains.push((start, count));
    Ok(start)
}

fn allocation_table_counts(
    data_count: usize,
    fat_capacity: usize,
    difat_capacity: usize,
) -> Result<(usize, usize)> {
    let mut fat_count = 1usize;
    loop {
        let difat_count = if fat_count <= HEADER_DIFAT_LEN {
            0
        } else {
            div_ceil(fat_count - HEADER_DIFAT_LEN, difat_capacity)
        };
        let total = data_count
            .checked_add(fat_count)
            .and_then(|value| value.checked_add(difat_count))
            .ok_or_else(|| Error::Limit("CFB sector count overflow".into()))?;
        let needed = div_ceil(total, fat_capacity).max(1);
        if needed == fat_count {
            return Ok((fat_count, difat_count));
        }
        fat_count = needed;
    }
}

fn mark_chain(fat: &mut [u32], start: u32, count: usize) -> Result<()> {
    let start = start as usize;
    for offset in 0..count {
        let index = start
            .checked_add(offset)
            .ok_or_else(|| Error::Limit("FAT chain index overflow".into()))?;
        fat[index] = if offset + 1 == count {
            END_OF_CHAIN
        } else {
            u32_len(index + 1, "FAT next sector")?
        };
    }
    Ok(())
}

fn u32_sector(values: &[u32], sector_len: usize) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(sector_len);
    for value in values {
        bytes.extend_from_slice(&value.to_le_bytes());
    }
    debug_assert_eq!(bytes.len(), sector_len);
    bytes
}

fn unallocated_directory_entry() -> DirectoryEntry {
    DirectoryEntry {
        name_buffer: [0; 32],
        name_length: 0,
        object_type: DirectoryObjectType::Unallocated,
        color: DirectoryColor::Red,
        left_sibling: DirectoryPointer::None,
        right_sibling: DirectoryPointer::None,
        child: DirectoryPointer::None,
        clsid: Guid::ZERO,
        state_bits: 0,
        creation_time: FileTime::ZERO,
        modified_time: FileTime::ZERO,
        start_sector: 0,
        stream_size: 0,
    }
}

fn encode_name(name: &str) -> Result<[u16; 32]> {
    let chars: Vec<_> = name.encode_utf16().collect();
    if chars.len() > 31 {
        return Err(Error::invalid(0, "CFB name exceeds 31 UTF-16 code units"));
    }
    if name
        .chars()
        .any(|value| matches!(value, '/' | '\\' | ':' | '!'))
    {
        return Err(Error::invalid(0, "CFB name contains a forbidden character"));
    }
    let mut buffer = [0; 32];
    buffer[..chars.len()].copy_from_slice(&chars);
    Ok(buffer)
}

pub(crate) fn validate_entry_name(name: &str) -> Result<()> {
    if name.contains('\0') {
        return Err(Error::invalid(0, "new CFB name contains NUL"));
    }
    encode_name(name).map(|_| ())
}

fn name_length(name: &str) -> Result<u16> {
    let chars = name.encode_utf16().count();
    u16::try_from((chars + 1) * 2).map_err(|_| Error::invalid(0, "CFB name length overflow"))
}

fn compare_names(left: &str, right: &str) -> Ordering {
    let left_len = left.encode_utf16().count();
    let right_len = right.encode_utf16().count();
    left_len.cmp(&right_len).then_with(|| {
        left.chars()
            .map(cfb_simple_uppercase)
            .cmp(right.chars().map(cfb_simple_uppercase))
    })
}

pub(crate) fn names_equal(left: &str, right: &str) -> bool {
    compare_names(left, right) == Ordering::Equal
}

fn cfb_simple_uppercase(value: char) -> char {
    // MS-CFB uses one-to-one invariant uppercase mapping. In particular,
    // U+00DF must not expand to the full-uppercase string "SS".
    match value {
        'ß' => 'ß',
        value => value.to_uppercase().next().unwrap_or(value),
    }
}

fn div_ceil(value: usize, divisor: usize) -> usize {
    value.div_ceil(divisor)
}

fn u32_len(value: usize, what: &str) -> Result<u32> {
    u32::try_from(value).map_err(|_| Error::Limit(format!("{what} exceeds u32")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn allocation_counts_include_fat_and_difat_sectors() {
        assert_eq!(allocation_table_counts(2, 128, 127).unwrap(), (1, 0));
        let (fat, difat) = allocation_table_counts(20_000, 128, 127).unwrap();
        assert!(fat > HEADER_DIFAT_LEN);
        assert!(difat > 0);
        assert!(20_000 + fat + difat <= fat * 128);
    }

    #[test]
    fn directory_comparison_uses_simple_uppercase_for_sharp_s() {
        assert_eq!(compare_names("ßY", "UF"), Ordering::Greater);
    }
}
