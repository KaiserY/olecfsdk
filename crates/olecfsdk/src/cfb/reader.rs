use std::{
    io::{self, BufRead, Read, Seek, SeekFrom},
    path::{Path, PathBuf},
    sync::Arc,
};

use crate::{
    Error, Result,
    common::{FileTime, Guid},
    limits::Limits,
};

use super::{
    CompoundFile, Difat, Directory, DirectoryObjectType, EntryKind, Fat, Header, MiniFat,
    MiniSectorId, SectorId, Version,
    allocation::END_OF_CHAIN,
    header::{FREE_SECTOR, HEADER_LEN, MINI_STREAM_CUTOFF},
    name,
    sector::{SectorRead, SeekSectorSource},
};

const DEFAULT_STREAM_BUFFER_SIZE: usize = 1024 * 1024;
const MINI_SECTOR_LEN: usize = 64;

/// Metadata for an MS-CFB directory entry without eagerly materializing its stream payload.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EntryInfo {
    pub stream_id: u32,
    pub path: PathBuf,
    pub name: String,
    pub kind: EntryKind,
    pub clsid: Guid,
    pub state_bits: u32,
    pub created: FileTime,
    pub modified: FileTime,
    pub stream_len: u64,
}

impl EntryInfo {
    pub fn is_stream(&self) -> bool {
        self.kind == EntryKind::Stream
    }

    pub fn is_storage(&self) -> bool {
        self.kind != EntryKind::Stream
    }
}

/// A seekable MS-CFB reader that keeps allocation metadata in memory and reads
/// stream payloads through bounded buffers.
pub struct CompoundFileReader<R> {
    version: Version,
    header: Header,
    difat: Difat,
    fat: Fat,
    mini_fat: MiniFat,
    directory: Directory,
    entries: Vec<EntryInfo>,
    root_mini_chain: Arc<[SectorId]>,
    source: SeekSectorSource<R>,
    header_padding_is_zero: bool,
    limits: Limits,
    stream_buffer_size: usize,
}

impl<R: Read + Seek> CompoundFileReader<R> {
    pub fn from_reader(reader: R) -> Result<Self> {
        Self::from_reader_with_limits(reader, Limits::default())
    }

    pub fn from_reader_strict(reader: R) -> Result<Self> {
        let compound = Self::from_reader(reader)?;
        compound.validate_strict()?;
        Ok(compound)
    }

    pub fn from_reader_with_limits(reader: R, limits: Limits) -> Result<Self> {
        Self::from_reader_with_buffer_size(reader, limits, DEFAULT_STREAM_BUFFER_SIZE)
    }

    pub fn from_reader_with_buffer_size(
        mut reader: R,
        limits: Limits,
        stream_buffer_size: usize,
    ) -> Result<Self> {
        let original_len = reader.seek(SeekFrom::End(0))?;
        if original_len > limits.max_file_size {
            return Err(Error::Limit(format!(
                "file length {original_len} exceeds {}",
                limits.max_file_size
            )));
        }
        reader.seek(SeekFrom::Start(0))?;
        let mut header_bytes = [0; HEADER_LEN];
        reader.read_exact(&mut header_bytes)?;
        let header = Header::from_bytes(&header_bytes)?;
        let sector_len = header.sector_len();
        if limits.max_allocation < sector_len {
            return Err(Error::Limit(format!(
                "stream buffer requires one {sector_len}-byte sector but max allocation is {}",
                limits.max_allocation
            )));
        }
        if original_len < sector_len as u64 {
            return Err(Error::invalid(
                0,
                "CFB file is shorter than its header sector",
            ));
        }
        let header_padding_is_zero = if sector_len > HEADER_LEN {
            let mut padding = vec![0; sector_len - HEADER_LEN];
            reader.read_exact(&mut padding)?;
            padding.iter().all(|byte| *byte == 0)
        } else {
            true
        };
        let mut source = SeekSectorSource::new(reader, original_len, &header)?;
        if source.sector_count() > u32::MAX as usize {
            return Err(Error::Limit("CFB sector count exceeds u32".into()));
        }
        source.full_sector(SectorId::new(header.first_directory_sector)?)?;
        let difat = Difat::read(&header, &mut source, limits)?;
        let fat = Fat::read(&difat, &mut source)?;
        let directory_sectors = fat.chain(header.first_directory_sector, source.sector_count())?;
        let mini_fat = MiniFat::read(&header, &fat, &mut source, limits)?;
        let directory = Directory::read(
            directory_sectors,
            header.number_of_directory_sectors,
            &mut source,
            limits,
        )?;
        let version = header.version();
        let entries = read_entry_info(&directory, version, limits)?;
        let root = directory.root();
        let root_len = root.effective_stream_size(version);
        let root_mini_chain = regular_chain(
            &fat,
            root.start_sector,
            root_len,
            source.sector_count(),
            source.sector_len(),
        )?;
        ensure_physical_capacity(&source, &root_mini_chain, root_len, "root mini")?;
        let root_mini_chain = Arc::from(root_mini_chain);
        let stream_buffer_size = stream_buffer_size
            .max(sector_len)
            .min(limits.max_allocation);
        Ok(Self {
            version,
            header,
            difat,
            fat,
            mini_fat,
            directory,
            entries,
            root_mini_chain,
            source,
            header_padding_is_zero,
            limits,
            stream_buffer_size,
        })
    }

    pub fn version(&self) -> Version {
        self.version
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

    pub fn directory(&self) -> &Directory {
        &self.directory
    }

    pub fn entries(&self) -> &[EntryInfo] {
        &self.entries
    }

    pub fn root_entry(&self) -> &EntryInfo {
        self.entries
            .iter()
            .find(|entry| entry.kind == EntryKind::Root)
            .expect("validated CFB directory has a root entry")
    }

    pub fn entry(&self, path: impl AsRef<Path>) -> Option<&EntryInfo> {
        self.entry_index(path.as_ref())
            .map(|index| &self.entries[index])
    }

    pub fn contains_entry(&self, path: impl AsRef<Path>) -> bool {
        self.entry(path).is_some()
    }

    pub fn is_stream(&self, path: impl AsRef<Path>) -> bool {
        self.entry(path).is_some_and(EntryInfo::is_stream)
    }

    pub fn is_storage(&self, path: impl AsRef<Path>) -> bool {
        self.entry(path).is_some_and(EntryInfo::is_storage)
    }

    pub fn children(&self, path: impl AsRef<Path>) -> Result<Vec<&EntryInfo>> {
        let parent = self.required_entry_index(path.as_ref())?;
        if self.entries[parent].is_stream() {
            return Err(Error::invalid(0, "CFB entry is not a storage"));
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

    pub fn walk_storage(&self, path: impl AsRef<Path>) -> Result<Vec<&EntryInfo>> {
        let root = self.required_entry_index(path.as_ref())?;
        if self.entries[root].is_stream() {
            return Err(Error::invalid(0, "CFB entry is not a storage"));
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

    pub fn open_stream(&mut self, path: impl AsRef<Path>) -> Result<CfbStream<'_, R>> {
        let index = self.required_entry_index(path.as_ref())?;
        let entry = &self.entries[index];
        if !entry.is_stream() {
            return Err(Error::invalid(0, "CFB entry is not a stream"));
        }
        let raw = self
            .directory
            .entries()
            .get(entry.stream_id as usize)
            .ok_or_else(|| Error::invalid(0, "CFB stream directory entry is missing"))?;
        let len = entry.stream_len;
        let chain = if len == 0 {
            StreamChain::Regular(Vec::new())
        } else if len < MINI_STREAM_CUTOFF as u64 {
            if matches!(raw.start_sector, END_OF_CHAIN | FREE_SECTOR) {
                return Err(Error::invalid(
                    0,
                    "non-empty stream has no mini-sector chain",
                ));
            }
            let root = self.directory.root();
            let root_len = root.effective_stream_size(self.version);
            let mini_sector_count = usize::try_from(root_len / MINI_SECTOR_LEN as u64)
                .map_err(|_| Error::Limit("root mini stream size does not fit usize".into()))?;
            let mini_chain = self.mini_fat.chain(raw.start_sector, mini_sector_count)?;
            ensure_chain_capacity(mini_chain.len(), MINI_SECTOR_LEN, len, "mini")?;
            StreamChain::Mini {
                mini_chain,
                root_chain: self.root_mini_chain.clone(),
                root_len,
            }
        } else {
            StreamChain::Regular(regular_chain(
                &self.fat,
                raw.start_sector,
                len,
                self.source.sector_count(),
                self.source.sector_len(),
            )?)
        };
        let buffer_capacity = usize::try_from(len)
            .unwrap_or(self.stream_buffer_size)
            .min(self.stream_buffer_size)
            .max(1);
        Ok(CfbStream {
            source: &mut self.source,
            chain,
            len,
            position: 0,
            buffer_start: 0,
            buffer_len: 0,
            buffer: vec![0; buffer_capacity],
        })
    }

    /// Consumes the streaming reader and materializes the existing full owned
    /// model. This is the explicit fallback for callers that need unrestricted
    /// logical editing before the file-backed editor is used.
    pub fn into_owned(self) -> Result<CompoundFile> {
        let limits = self.limits;
        let mut reader = self.source.into_inner();
        reader.seek(SeekFrom::Start(0))?;
        CompoundFile::from_reader_with_limits(reader, limits)
    }

    pub fn into_inner(self) -> R {
        self.source.into_inner()
    }

    pub fn validate_strict(&self) -> Result<()> {
        if self.header.clsid != [0; 16] {
            return Err(Error::invalid(8, "CFB header CLSID must be zero"));
        }
        if self.header.reserved != [0; 6] {
            return Err(Error::invalid(34, "CFB header reserved bytes must be zero"));
        }
        if !self.header_padding_is_zero {
            return Err(Error::invalid(
                HEADER_LEN as u64,
                "CFB v4 header padding must be zero",
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
            .checked_div(MINI_SECTOR_LEN as u64)
            .and_then(|count| usize::try_from(count).ok())
            .ok_or_else(|| Error::Limit("root mini stream size does not fit usize".into()))?;
        self.mini_fat.validate_strict(root_mini_sector_count)?;
        self.directory.validate_strict(self.version)?;
        self.validate_stream_chains(root_mini_sector_count)
    }

    fn entry_index(&self, path: &Path) -> Option<usize> {
        let names = name::path_components(path)?;
        let mut current = self
            .entries
            .iter()
            .position(|entry| entry.kind == EntryKind::Root)?;
        for requested in names {
            let parent = self.entries[current].path.as_path();
            current = self.entries.iter().position(|entry| {
                entry.path.parent() == Some(parent) && name::names_equal(&entry.name, &requested)
            })?;
        }
        Some(current)
    }

    fn required_entry_index(&self, path: &Path) -> Result<usize> {
        self.entry_index(path).ok_or_else(|| {
            Error::invalid(0, format!("CFB entry {} does not exist", path.display()))
        })
    }

    fn validate_stream_chains(&self, root_mini_sector_count: usize) -> Result<()> {
        for entry in self.entries.iter().filter(|entry| entry.is_stream()) {
            if entry.stream_len == 0 {
                continue;
            }
            let raw = self
                .directory
                .entries()
                .get(entry.stream_id as usize)
                .ok_or_else(|| Error::invalid(0, "CFB stream directory entry is missing"))?;
            if entry.stream_len < MINI_STREAM_CUTOFF as u64 {
                if matches!(raw.start_sector, END_OF_CHAIN | FREE_SECTOR) {
                    return Err(Error::invalid(
                        0,
                        "non-empty stream has no mini-sector chain",
                    ));
                }
                let chain = self
                    .mini_fat
                    .chain(raw.start_sector, root_mini_sector_count)?;
                ensure_chain_capacity(chain.len(), MINI_SECTOR_LEN, entry.stream_len, "mini")?;
            } else {
                let chain = regular_chain(
                    &self.fat,
                    raw.start_sector,
                    entry.stream_len,
                    self.source.sector_count(),
                    self.source.sector_len(),
                )?;
                ensure_physical_capacity(&self.source, &chain, entry.stream_len, "regular")?;
            }
        }
        Ok(())
    }
}

impl CompoundFileReader<std::fs::File> {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_reader(std::fs::File::open(path)?)
    }

    pub fn open_strict(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_reader_strict(std::fs::File::open(path)?)
    }
}

enum StreamChain {
    Regular(Vec<SectorId>),
    Mini {
        mini_chain: Vec<MiniSectorId>,
        root_chain: Arc<[SectorId]>,
        root_len: u64,
    },
}

/// A bounded-buffer reader for one stream object in a [`CompoundFileReader`].
pub struct CfbStream<'a, R> {
    source: &'a mut SeekSectorSource<R>,
    chain: StreamChain,
    len: u64,
    position: u64,
    buffer_start: u64,
    buffer_len: usize,
    buffer: Vec<u8>,
}

impl<R> CfbStream<'_, R> {
    pub fn len(&self) -> u64 {
        self.len
    }

    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    pub fn buffer_capacity(&self) -> usize {
        self.buffer.len()
    }
}

impl<R: Read + Seek> BufRead for CfbStream<'_, R> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.position == self.len {
            return Ok(&[]);
        }
        let buffer_end = self.buffer_start + self.buffer_len as u64;
        if self.buffer_len == 0 || self.position < self.buffer_start || self.position >= buffer_end
        {
            self.buffer_start = self.position;
            let remaining = self.len - self.position;
            let requested = usize::try_from(remaining)
                .unwrap_or(usize::MAX)
                .min(self.buffer.len());
            self.buffer_len = match &self.chain {
                StreamChain::Regular(chain) => read_regular_at(
                    self.source,
                    chain,
                    self.len,
                    self.position,
                    &mut self.buffer[..requested],
                ),
                StreamChain::Mini {
                    mini_chain,
                    root_chain,
                    root_len,
                } => read_mini_at(
                    self.source,
                    mini_chain,
                    root_chain,
                    *root_len,
                    self.len,
                    self.position,
                    &mut self.buffer[..requested],
                ),
            }
            .map_err(as_io_error)?;
        }
        let offset = usize::try_from(self.position - self.buffer_start)
            .map_err(|_| io::Error::other("CFB stream buffer offset does not fit usize"))?;
        Ok(&self.buffer[offset..self.buffer_len])
    }

    fn consume(&mut self, amount: usize) {
        debug_assert!(self.position + amount as u64 <= self.buffer_start + self.buffer_len as u64);
        self.position += amount as u64;
    }
}

impl<R: Read + Seek> Read for CfbStream<'_, R> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl<R: Read + Seek> Seek for CfbStream<'_, R> {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let candidate = match position {
            SeekFrom::Start(value) => i128::from(value),
            SeekFrom::End(delta) => i128::from(self.len) + i128::from(delta),
            SeekFrom::Current(delta) => i128::from(self.position) + i128::from(delta),
        };
        if candidate < 0 || candidate > i128::from(self.len) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "cannot seek to {candidate}; CFB stream length is {}",
                    self.len
                ),
            ));
        }
        self.position = candidate as u64;
        Ok(self.position)
    }
}

fn read_entry_info(
    directory: &Directory,
    version: Version,
    limits: Limits,
) -> Result<Vec<EntryInfo>> {
    let mut entries = Vec::new();
    for (stream_id, path) in directory.paths()? {
        let raw = directory
            .entries()
            .get(stream_id as usize)
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
        let stream_len = raw.effective_stream_size(version);
        if stream_len > limits.max_stream_size {
            return Err(Error::Limit(format!(
                "stream length {stream_len} exceeds {}",
                limits.max_stream_size
            )));
        }
        entries.push(EntryInfo {
            stream_id,
            path,
            name: raw.name()?,
            kind,
            clsid: if kind == EntryKind::Stream {
                Guid::ZERO
            } else {
                raw.clsid
            },
            state_bits: raw.state_bits,
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
            stream_len,
        });
    }
    entries.sort_by(|left, right| left.path.cmp(&right.path));
    Ok(entries)
}

fn regular_chain(
    fat: &Fat,
    start: u32,
    len: u64,
    sector_count: usize,
    sector_len: usize,
) -> Result<Vec<SectorId>> {
    if len == 0 {
        return Ok(Vec::new());
    }
    if matches!(start, END_OF_CHAIN | FREE_SECTOR) {
        return Err(Error::invalid(
            0,
            "non-empty stream has no regular sector chain",
        ));
    }
    let chain = fat.chain(start, sector_count)?;
    ensure_chain_capacity(chain.len(), sector_len, len, "regular")?;
    Ok(chain)
}

fn ensure_chain_capacity(count: usize, unit_len: usize, len: u64, kind: &str) -> Result<()> {
    let capacity = u64::try_from(count)
        .ok()
        .and_then(|count| count.checked_mul(unit_len as u64))
        .ok_or_else(|| Error::Limit(format!("{kind} stream chain size overflow")))?;
    if capacity < len {
        return Err(Error::invalid(
            0,
            format!("{kind} stream chain is shorter than stream size"),
        ));
    }
    Ok(())
}

fn ensure_physical_capacity<S: SectorRead + ?Sized>(
    source: &S,
    chain: &[SectorId],
    len: u64,
    kind: &str,
) -> Result<()> {
    if len == 0 {
        return Ok(());
    }
    let sector_len = source.sector_len() as u64;
    let final_index = usize::try_from((len - 1) / sector_len)
        .map_err(|_| Error::Limit(format!("{kind} stream chain index does not fit usize")))?;
    let final_sector = *chain
        .get(final_index)
        .ok_or_else(|| Error::invalid(0, format!("{kind} stream chain is too short")))?;
    let required = usize::try_from((len - 1) % sector_len + 1)
        .map_err(|_| Error::Limit(format!("{kind} stream tail length does not fit usize")))?;
    if source.valid_len(final_sector) < required {
        return Err(Error::invalid(
            0,
            format!("{kind} stream data is truncated at physical EOF"),
        ));
    }
    Ok(())
}

fn read_regular_at<S: SectorRead + ?Sized>(
    source: &mut S,
    chain: &[SectorId],
    stream_len: u64,
    offset: u64,
    output: &mut [u8],
) -> Result<usize> {
    if offset >= stream_len || output.is_empty() {
        return Ok(0);
    }
    let requested = usize::try_from((stream_len - offset).min(output.len() as u64))
        .map_err(|_| Error::Limit("stream read length does not fit usize".into()))?;
    let sector_len = source.sector_len() as u64;
    let mut logical = offset;
    let mut written = 0usize;
    while written < requested {
        let chain_index = usize::try_from(logical / sector_len)
            .map_err(|_| Error::Limit("stream chain index does not fit usize".into()))?;
        let within = usize::try_from(logical % sector_len)
            .map_err(|_| Error::Limit("sector offset does not fit usize".into()))?;
        let sector_id = *chain
            .get(chain_index)
            .ok_or_else(|| Error::invalid(0, "stream chain ended before stream size"))?;
        let valid_len = source.valid_len(sector_id);
        let count = (requested - written).min(source.sector_len() - within);
        if within + count > valid_len {
            return Err(Error::invalid(
                0,
                "stream data is truncated at physical EOF",
            ));
        }
        let bytes = source.sector(sector_id)?;
        output[written..written + count].copy_from_slice(&bytes.as_ref()[within..within + count]);
        logical += count as u64;
        written += count;
    }
    Ok(written)
}

fn read_mini_at<S: SectorRead + ?Sized>(
    source: &mut S,
    mini_chain: &[MiniSectorId],
    root_chain: &[SectorId],
    root_len: u64,
    stream_len: u64,
    offset: u64,
    output: &mut [u8],
) -> Result<usize> {
    if offset >= stream_len || output.is_empty() {
        return Ok(0);
    }
    let requested = usize::try_from((stream_len - offset).min(output.len() as u64))
        .map_err(|_| Error::Limit("mini stream read length does not fit usize".into()))?;
    let mut logical = offset;
    let mut written = 0usize;
    while written < requested {
        let chain_index = usize::try_from(logical / MINI_SECTOR_LEN as u64)
            .map_err(|_| Error::Limit("mini-chain index does not fit usize".into()))?;
        let within = usize::try_from(logical % MINI_SECTOR_LEN as u64)
            .map_err(|_| Error::Limit("mini-sector offset does not fit usize".into()))?;
        let mini_sector = *mini_chain
            .get(chain_index)
            .ok_or_else(|| Error::invalid(0, "mini stream chain ended before stream size"))?;
        let count = (requested - written).min(MINI_SECTOR_LEN - within);
        let root_offset = u64::from(mini_sector.get())
            .checked_mul(MINI_SECTOR_LEN as u64)
            .and_then(|value| value.checked_add(within as u64))
            .ok_or_else(|| Error::Limit("root mini stream offset overflow".into()))?;
        let read = read_regular_at(
            source,
            root_chain,
            root_len,
            root_offset,
            &mut output[written..written + count],
        )?;
        if read != count {
            return Err(Error::invalid(0, "mini-sector is outside the root stream"));
        }
        logical += count as u64;
        written += count;
    }
    Ok(written)
}

fn as_io_error(error: Error) -> io::Error {
    match error {
        Error::Io(error) => error,
        other => io::Error::new(io::ErrorKind::InvalidData, other),
    }
}

#[cfg(test)]
mod tests {
    use std::{
        cell::Cell,
        io::{Cursor, Read, Seek, SeekFrom},
        rc::Rc,
    };

    use super::*;

    struct CountingReader {
        inner: Cursor<Vec<u8>>,
        bytes_read: Rc<Cell<u64>>,
        largest_read: Rc<Cell<usize>>,
    }

    impl CountingReader {
        fn new(bytes: Vec<u8>) -> (Self, Rc<Cell<u64>>, Rc<Cell<usize>>) {
            let bytes_read = Rc::new(Cell::new(0));
            let largest_read = Rc::new(Cell::new(0));
            (
                Self {
                    inner: Cursor::new(bytes),
                    bytes_read: bytes_read.clone(),
                    largest_read: largest_read.clone(),
                },
                bytes_read,
                largest_read,
            )
        }
    }

    impl Read for CountingReader {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            let count = self.inner.read(output)?;
            self.bytes_read.set(self.bytes_read.get() + count as u64);
            self.largest_read.set(self.largest_read.get().max(count));
            Ok(count)
        }
    }

    impl Seek for CountingReader {
        fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
            self.inner.seek(position)
        }
    }

    fn compound_bytes() -> (Vec<u8>, Vec<u8>, Vec<u8>) {
        let large: Vec<_> = (0..4 * 1024 * 1024)
            .map(|index| (index % 251) as u8)
            .collect();
        let small: Vec<_> = (0..977).map(|index| (index % 239) as u8).collect();
        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound.create_storage_all("/Data/Nested").unwrap();
        compound
            .create_stream("/Data/Large", large.clone())
            .unwrap();
        compound
            .create_stream("/Data/Nested/Small", small.clone())
            .unwrap();
        (compound.to_bytes().unwrap(), large, small)
    }

    #[test]
    fn opening_seekable_cfb_does_not_materialize_stream_payloads() {
        let (bytes, large, _) = compound_bytes();
        let file_len = bytes.len() as u64;
        let (source, bytes_read, largest_read) = CountingReader::new(bytes);
        let mut compound =
            CompoundFileReader::from_reader_with_buffer_size(source, Limits::default(), 4096)
                .unwrap();

        assert!(bytes_read.get() < file_len / 8);
        assert!(largest_read.get() <= 4096);
        assert_eq!(
            compound.entry("/data/large").unwrap().stream_len,
            large.len() as u64
        );

        let mut stream = compound.open_stream("/DATA/LARGE").unwrap();
        assert_eq!(stream.len(), large.len() as u64);
        assert_eq!(stream.buffer_capacity(), 4096);
        let mut first = [0; 37];
        stream.read_exact(&mut first).unwrap();
        assert_eq!(first.as_slice(), &large[..first.len()]);
        stream.seek(SeekFrom::End(-53)).unwrap();
        let mut tail = Vec::new();
        stream.read_to_end(&mut tail).unwrap();
        assert_eq!(tail, large[large.len() - 53..]);
        assert!(bytes_read.get() < file_len / 4);
    }

    #[test]
    fn mini_stream_reads_and_seeks_across_mini_sector_boundaries() {
        let (bytes, _, small) = compound_bytes();
        let mut compound = CompoundFileReader::from_reader_strict(Cursor::new(bytes)).unwrap();
        let mut stream = compound.open_stream("/Data/Nested/Small").unwrap();
        assert_eq!(stream.len(), small.len() as u64);
        assert_eq!(stream.buffer_capacity(), small.len());
        stream.seek(SeekFrom::Start(61)).unwrap();
        let mut actual = vec![0; 197];
        stream.read_exact(&mut actual).unwrap();
        assert_eq!(actual, small[61..61 + 197]);
        assert!(stream.seek(SeekFrom::End(1)).is_err());
        assert!(stream.seek(SeekFrom::Current(-10_000)).is_err());
    }

    #[test]
    fn into_owned_is_an_explicit_full_feature_fallback() {
        let (bytes, large, small) = compound_bytes();
        let reader = CompoundFileReader::from_reader(Cursor::new(bytes)).unwrap();
        let owned = reader.into_owned().unwrap();
        assert_eq!(owned.stream("/Data/Large"), Some(large.as_slice()));
        assert_eq!(owned.stream("/Data/Nested/Small"), Some(small.as_slice()));
    }
}
