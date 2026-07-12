use uuid::Uuid;
use web_time::{Duration, SystemTime, UNIX_EPOCH};

use crate::{Error, Result, limits::Limits};

use super::{
    Entry, EntryKind,
    allocation::{END_OF_CHAIN, Fat, MiniFat},
    directory::{Directory, DirectoryObjectType},
    header::{FREE_SECTOR, Header, MINI_STREAM_CUTOFF},
    sector::SectorSource,
};

const MINI_SECTOR_LEN: usize = 64;
const UNIX_EPOCH_FILETIME: u64 = 116_444_736_000_000_000;

pub(crate) fn read_entries(
    header: &Header,
    fat: &Fat,
    mini_fat: &MiniFat,
    directory: &Directory,
    source: &SectorSource<'_>,
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
            Uuid::nil()
        } else {
            clsid_from_bytes(raw.clsid)
        };
        let (created, modified) = if kind == EntryKind::Stream {
            (filetime_to_system_time(0), filetime_to_system_time(0))
        } else {
            (
                filetime_to_system_time(raw.creation_time),
                filetime_to_system_time(raw.modified_time),
            )
        };
        entries.push(Entry {
            path,
            name: raw.name()?,
            kind,
            clsid,
            state_bits: raw.state_bits,
            created,
            modified,
            data,
        });
    }
    entries.sort_by(|left, right| left.path.cmp(&right.path));
    Ok(entries)
}

fn read_regular_stream(
    fat: &Fat,
    source: &SectorSource<'_>,
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
        let bytes = source.sector(sector)?;
        let needed = remaining.min(bytes.len());
        if needed > source.valid_len(sector) {
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

fn clsid_from_bytes(bytes: [u8; 16]) -> Uuid {
    let d4 = [
        bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15],
    ];
    Uuid::from_fields(
        u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
        u16::from_le_bytes([bytes[4], bytes[5]]),
        u16::from_le_bytes([bytes[6], bytes[7]]),
        &d4,
    )
}

fn filetime_to_system_time(value: u64) -> SystemTime {
    let delta = |ticks: u64| Duration::new(ticks / 10_000_000, (ticks % 10_000_000) as u32 * 100);
    if value >= UNIX_EPOCH_FILETIME {
        UNIX_EPOCH
            .checked_add(delta(value - UNIX_EPOCH_FILETIME))
            .unwrap_or(UNIX_EPOCH)
    } else {
        UNIX_EPOCH
            .checked_sub(delta(UNIX_EPOCH_FILETIME - value))
            .unwrap_or(UNIX_EPOCH)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn filetime_unix_epoch_conversion_is_exact() {
        assert_eq!(filetime_to_system_time(UNIX_EPOCH_FILETIME), UNIX_EPOCH);
    }
}
