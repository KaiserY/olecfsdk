//! Typed file and substream roots for the Excel binary format.

use std::{ops::Range, path::Path};

use crate::{Error, Result, cfb::CompoundFile, limits::Limits};

use super::{BiffRecord, BiffRecordData, BiffStream, RevisionLogStream};

const WORKBOOK_STREAM: &str = "/Workbook";
const BOOK_STREAM: &str = "/Book";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum XlsStreamName {
    Workbook,
    Book,
}

impl XlsStreamName {
    pub const fn path(self) -> &'static str {
        match self {
            Self::Workbook => WORKBOOK_STREAM,
            Self::Book => BOOK_STREAM,
        }
    }
}

/// Specification-level kind of a BOF/EOF BIFF substream.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BiffSubstreamKind {
    WorkbookGlobals,
    WorksheetOrDialogSheet,
    ChartSheet,
    MacroSheet,
    /// BIFF producer/legacy kinds outside the current MS-XLS BOF table.
    Compatibility(u16),
}

/// A structural BOF/EOF node indexing records in [`BiffWorkbookTree::stream`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BiffSubstreamNode {
    pub kind: BiffSubstreamKind,
    /// Includes the opening BOF and closing EOF records.
    pub record_range: Range<usize>,
    pub children: Vec<BiffSubstreamNode>,
}

/// Full BIFF stream plus its lossless structural substream index.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BiffWorkbookTree {
    pub stream: BiffStream,
    pub substreams: Vec<BiffSubstreamNode>,
    /// Ranges not enclosed by BOF/EOF, retained for compatibility analysis.
    pub outside_substream_ranges: Vec<Range<usize>>,
}

/// Complete typed root for an Excel binary file.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsFile {
    pub compound_file: CompoundFile,
    /// Every root BIFF workbook stream. Some compatibility files contain both
    /// the modern `Workbook` name and the legacy `Book` name.
    pub workbooks: Vec<XlsWorkbookStream>,
    pub revision_log: Option<XlsRevisionLog>,
}

/// A named root BIFF stream and its full record/substream tree.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlsWorkbookStream {
    pub name: XlsStreamName,
    pub tree: BiffWorkbookTree,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlsRevisionLog {
    Parsed(RevisionLogStream),
    Compatibility { bytes: Vec<u8>, reason: String },
}

impl BiffSubstreamKind {
    pub const fn from_document_type(value: u16) -> Self {
        match value {
            0x0005 => Self::WorkbookGlobals,
            0x0010 => Self::WorksheetOrDialogSheet,
            0x0020 => Self::ChartSheet,
            0x0040 => Self::MacroSheet,
            value => Self::Compatibility(value),
        }
    }
}

impl BiffSubstreamNode {
    pub fn records<'a>(&self, tree: &'a BiffWorkbookTree) -> Option<&'a [BiffRecord]> {
        tree.stream.records.get(self.record_range.clone())
    }
}

impl BiffWorkbookTree {
    pub fn from_stream(stream: BiffStream) -> Result<Self> {
        let mut roots = Vec::new();
        let mut stack = Vec::<OpenSubstream>::new();
        let mut covered = vec![false; stream.records.len()];
        for (index, record) in stream.records.iter().enumerate() {
            match &record.data {
                BiffRecordData::Bof(bof) => stack.push(OpenSubstream {
                    kind: BiffSubstreamKind::from_document_type(bof.document_type),
                    start: index,
                    children: Vec::new(),
                }),
                BiffRecordData::LegacyBof { .. } => stack.push(OpenSubstream {
                    kind: BiffSubstreamKind::Compatibility(0xffff),
                    start: index,
                    children: Vec::new(),
                }),
                BiffRecordData::Eof => {
                    let open = stack.pop().ok_or_else(|| {
                        Error::invalid(record.offset.into(), "BIFF EOF has no matching BOF")
                    })?;
                    covered[open.start..=index].fill(true);
                    let node = BiffSubstreamNode {
                        kind: open.kind,
                        record_range: open.start..index + 1,
                        children: open.children,
                    };
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(node);
                    } else {
                        roots.push(node);
                    }
                }
                _ => {}
            }
        }
        if let Some(open) = stack.last() {
            return Err(Error::invalid(
                stream.records[open.start].offset.into(),
                "BIFF BOF has no matching EOF",
            ));
        }
        let outside_substream_ranges = matching_ranges(&covered, false);
        Ok(Self {
            stream,
            substreams: roots,
            outside_substream_ranges,
        })
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_stream(BiffStream::from_bytes(bytes)?)
    }

    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        Self::from_stream(BiffStream::from_bytes_with_limits(bytes, limits)?)
    }

    /// Rebuilds the BOF/EOF tree after inserting, removing, or moving records.
    pub fn reindex(&mut self) -> Result<()> {
        let indexed = Self::from_stream(self.stream.clone())?;
        self.substreams = indexed.substreams;
        self.outside_substream_ranges = indexed.outside_substream_ranges;
        Ok(())
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        // Re-index first so edits cannot leave a stale structural tree.
        let indexed = Self::from_stream(self.stream.clone())?;
        if indexed.substreams != self.substreams
            || indexed.outside_substream_ranges != self.outside_substream_ranges
        {
            return Err(Error::invalid(
                0,
                "BIFF substream index is stale after record-tree editing",
            ));
        }
        self.stream.to_bytes()
    }
}

impl XlsFile {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_compound_file(CompoundFile::open(path)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        Self::from_compound_file_with_limits(
            CompoundFile::from_bytes_with_limits(bytes, limits)?,
            limits,
        )
    }

    pub fn from_compound_file(compound_file: CompoundFile) -> Result<Self> {
        Self::from_compound_file_with_limits(compound_file, Limits::default())
    }

    pub fn from_compound_file_with_limits(
        compound_file: CompoundFile,
        limits: Limits,
    ) -> Result<Self> {
        let mut workbooks = Vec::new();
        for name in [XlsStreamName::Workbook, XlsStreamName::Book] {
            if let Some(bytes) = compound_file.stream(name.path()) {
                workbooks.push(XlsWorkbookStream {
                    name,
                    tree: BiffWorkbookTree::from_bytes_with_limits(bytes, limits)?,
                });
            }
        }
        if workbooks.is_empty() {
            return Err(Error::invalid(0, "Workbook/Book stream is missing"));
        }
        let revision_log = compound_file
            .stream(super::REVISION_LOG_STREAM_PATH)
            .map(
                |bytes| match RevisionLogStream::from_bytes_with_limits(bytes, limits) {
                    Ok(value) => XlsRevisionLog::Parsed(value),
                    Err(error) => XlsRevisionLog::Compatibility {
                        bytes: bytes.to_vec(),
                        reason: error.to_string(),
                    },
                },
            );
        Ok(Self {
            compound_file,
            workbooks,
            revision_log,
        })
    }

    pub fn to_compound_file(&self) -> Result<CompoundFile> {
        if self.workbooks.is_empty() {
            return Err(Error::invalid(0, "XLS file has no BIFF workbook stream"));
        }
        if [XlsStreamName::Workbook, XlsStreamName::Book]
            .into_iter()
            .any(|name| {
                self.workbooks
                    .iter()
                    .filter(|workbook| workbook.name == name)
                    .count()
                    > 1
            })
        {
            return Err(Error::invalid(
                0,
                "XLS file has duplicate BIFF stream names",
            ));
        }
        let mut compound = self.compound_file.clone();
        for name in [XlsStreamName::Workbook, XlsStreamName::Book] {
            if let Some(workbook) = self.workbooks.iter().find(|workbook| workbook.name == name) {
                compound.create_or_replace_stream(name.path(), workbook.tree.to_bytes()?)?;
            } else if compound.is_stream(name.path()) {
                compound.remove_stream(name.path())?;
            }
        }
        match &self.revision_log {
            Some(XlsRevisionLog::Parsed(log)) => log.save(&mut compound)?,
            Some(XlsRevisionLog::Compatibility { bytes, .. }) => {
                compound.replace_stream(super::REVISION_LOG_STREAM_PATH, bytes.clone())?;
            }
            None if compound.is_stream(super::REVISION_LOG_STREAM_PATH) => {
                compound.remove_stream(super::REVISION_LOG_STREAM_PATH)?;
            }
            None => {}
        }
        Ok(compound)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_compound_file()?.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.to_compound_file()?.save(path)
    }
}

#[derive(Debug)]
struct OpenSubstream {
    kind: BiffSubstreamKind,
    start: usize,
    children: Vec<BiffSubstreamNode>,
}

fn matching_ranges(values: &[bool], target: bool) -> Vec<Range<usize>> {
    let mut ranges = Vec::new();
    let mut start = None;
    for (index, value) in values.iter().copied().enumerate() {
        match (value == target, start) {
            (true, None) => start = Some(index),
            (false, Some(from)) => {
                ranges.push(from..index);
                start = None;
            }
            _ => {}
        }
    }
    if let Some(from) = start {
        ranges.push(from..values.len());
    }
    ranges
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xls::BofRecord;

    fn record(offset: u32, data: BiffRecordData) -> BiffRecord {
        BiffRecord { offset, data }
    }

    #[test]
    fn indexes_nested_bof_eof_without_flattening_records() {
        let stream = BiffStream {
            records: vec![
                record(
                    0,
                    BiffRecordData::Bof(BofRecord {
                        version: 0x0600,
                        document_type: 0x0010,
                        build_identifier: 0,
                        build_year: 0,
                        history_flags: 0,
                        lowest_version: 0,
                    }),
                ),
                record(
                    20,
                    BiffRecordData::Bof(BofRecord {
                        version: 0x0600,
                        document_type: 0x0020,
                        build_identifier: 0,
                        build_year: 0,
                        history_flags: 0,
                        lowest_version: 0,
                    }),
                ),
                record(40, BiffRecordData::Eof),
                record(44, BiffRecordData::Eof),
            ],
            trailing_padding: Vec::new(),
        };
        let mut tree = BiffWorkbookTree::from_stream(stream).unwrap();
        assert_eq!(tree.substreams.len(), 1);
        assert_eq!(tree.substreams[0].record_range, 0..4);
        assert_eq!(tree.substreams[0].children[0].record_range, 1..3);
        assert!(tree.outside_substream_ranges.is_empty());

        tree.stream
            .records
            .insert(2, record(30, BiffRecordData::CodePage { code_page: 1252 }));
        assert!(tree.to_bytes().is_err());
        tree.reindex().unwrap();
        assert_eq!(tree.substreams[0].record_range, 0..5);
        assert_eq!(tree.substreams[0].children[0].record_range, 1..4);
    }
}
