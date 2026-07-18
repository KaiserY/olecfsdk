use crate::{
    Error, Result,
    common::{FileTime, Guid},
    limits::Limits,
};

use super::{
    Entry, EntryKind,
    allocation::{END_OF_CHAIN, Fat, MiniFat},
    directory::{Directory, DirectoryObjectType},
    header::{FREE_SECTOR, Header, MINI_STREAM_CUTOFF},
    sector::SectorRead,
};

const MINI_SECTOR_LEN: usize = 64;

pub(crate) fn read_entries<S: SectorRead + ?Sized>(
    header: &Header,
    fat: &Fat,
    mini_fat: &MiniFat,
    directory: &Directory,
    source: &mut S,
    limits: Limits,
) -> Result<Vec<Entry>> {
    let root = directory.root();
    let root_len = checked_stream_len(root.effective_stream_size(header.version()), limits)?;
    let root_mini_stream = read_regular_stream(fat, source, root.start_sector, root_len, limits)?;
    if !root_mini_stream.len().is_multiple_of(MINI_SECTOR_LEN) {
        return Err(Error::invalid(
            120,
            "root mini stream is not 64-byte aligned",
        ));
    }

    let mut entries = Vec::new();
    for (id, path) in directory.paths()? {
        let raw = directory
            .entries()
            .get(id as usize)
            .ok_or_else(|| Error::invalid(0, "directory path references a missing entry"))?;
        let kind = match raw.object_type {
            DirectoryObjectType::Root => EntryKind::Root,
            DirectoryObjectType::Storage => EntryKind::Storage,
            DirectoryObjectType::Stream => EntryKind::Stream,
            DirectoryObjectType::Unallocated => {
                return Err(Error::invalid(
                    0,
                    "unallocated entry is reachable from root",
                ));
            }
        };
        let stream_len = checked_stream_len(raw.effective_stream_size(header.version()), limits)?;
        let data = if kind != EntryKind::Stream || stream_len == 0 {
            Vec::new()
        } else if stream_len < MINI_STREAM_CUTOFF as usize {
            read_mini_stream(
                mini_fat,
                &root_mini_stream,
                raw.start_sector,
                stream_len,
                limits,
            )?
        } else {
            read_regular_stream(fat, source, raw.start_sector, stream_len, limits)?
        };
        let clsid = if kind == EntryKind::Stream {
            Guid::ZERO
        } else {
            raw.clsid
        };
        entries.push(Entry {
            path,
            name: raw.name()?,
            kind,
            clsid,
            state_bits: raw.state_bits,
            // Keep the raw directory entry available through `Directory`, but
            // expose a spec-normalized logical model for editing and rewrite.
            // MS-CFB requires both stream times and the root creation time to
            // be zero.
            created: if matches!(kind, EntryKind::Root | EntryKind::Stream) {
                FileTime::ZERO
            } else {
                raw.creation_time
            },
            modified: if kind == EntryKind::Stream {
                FileTime::ZERO
            } else {
                raw.modified_time
            },
            data: data.into(),
        });
    }
    entries.sort_by(|left, right| left.path.cmp(&right.path));
    Ok(entries)
}

fn read_regular_stream<S: SectorRead + ?Sized>(
    fat: &Fat,
    source: &mut S,
    start: u32,
    len: usize,
    limits: Limits,
) -> Result<Vec<u8>> {
    if len == 0 {
        return Ok(Vec::new());
    }
    if matches!(start, END_OF_CHAIN | FREE_SECTOR) {
        return Err(Error::invalid(
            0,
            "non-empty stream has no regular sector chain",
        ));
    }
    let chain = fat.chain(start, source.sector_count())?;
    let capacity = chain
        .len()
        .checked_mul(source.sector_len())
        .ok_or_else(|| Error::Limit("regular stream chain size overflow".into()))?;
    if capacity < len {
        return Err(Error::invalid(
            0,
            "regular stream chain is shorter than stream size",
        ));
    }
    if len > limits.max_allocation {
        return Err(Error::Limit(format!(
            "stream allocation {len} exceeds {}",
            limits.max_allocation
        )));
    }
    let mut data = Vec::with_capacity(len);
    for sector in chain {
        let remaining = len - data.len();
        if remaining == 0 {
            break;
        }
        let valid_len = source.valid_len(sector);
        let bytes = source.sector(sector)?;
        let bytes = bytes.as_ref();
        let needed = remaining.min(bytes.len());
        if needed > valid_len {
            return Err(Error::invalid(
                0,
                "stream data is truncated at physical EOF",
            ));
        }
        data.extend_from_slice(&bytes[..needed]);
    }
    Ok(data)
}

fn read_mini_stream(
    mini_fat: &MiniFat,
    root_stream: &[u8],
    start: u32,
    len: usize,
    limits: Limits,
) -> Result<Vec<u8>> {
    if len == 0 {
        return Ok(Vec::new());
    }
    let mini_sector_count = root_stream.len() / MINI_SECTOR_LEN;
    let chain = mini_fat.chain(start, mini_sector_count)?;
    let capacity = chain
        .len()
        .checked_mul(MINI_SECTOR_LEN)
        .ok_or_else(|| Error::Limit("mini stream chain size overflow".into()))?;
    if capacity < len {
        return Err(Error::invalid(
            0,
            "mini stream chain is shorter than stream size",
        ));
    }
    if len > limits.max_allocation {
        return Err(Error::Limit(format!(
            "stream allocation {len} exceeds {}",
            limits.max_allocation
        )));
    }
    let mut data = Vec::with_capacity(len);
    for mini_sector in chain {
        let start = mini_sector
            .get()
            .checked_mul(MINI_SECTOR_LEN as u32)
            .ok_or_else(|| Error::invalid(0, "mini-sector offset overflow"))?
            as usize;
        let end = start
            .checked_add(MINI_SECTOR_LEN)
            .ok_or_else(|| Error::invalid(0, "mini-sector end overflow"))?;
        let bytes = root_stream
            .get(start..end)
            .ok_or_else(|| Error::invalid(0, "mini-sector is outside the root stream"))?;
        let remaining = len - data.len();
        data.extend_from_slice(&bytes[..remaining.min(MINI_SECTOR_LEN)]);
        if data.len() == len {
            break;
        }
    }
    Ok(data)
}

fn checked_stream_len(len: u64, limits: Limits) -> Result<usize> {
    if len > limits.max_stream_size {
        return Err(Error::Limit(format!(
            "stream length {len} exceeds {}",
            limits.max_stream_size
        )));
    }
    usize::try_from(len).map_err(|_| Error::Limit("stream length does not fit usize".into()))
}
