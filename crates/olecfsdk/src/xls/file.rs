//! Typed file and substream roots for the Excel binary format.

use std::{ops::Range, path::Path};

use crate::{
    Error, Result,
    cfb::CompoundFile,
    io::BinaryFormat,
    limits::Limits,
    office_art::{OfficeArtDrawingGraph, OfficeArtStream},
    parse::{
        ParseDiagnostic, ParseDiagnosticCode, ParseOptions, ParseOutcome, SpecificationReference,
        compound_from_bytes, compound_from_path, compound_outcome,
    },
    save::SaveOptions,
};

use super::{
    BiffRecord, BiffRecordData, BiffStream, DevModeW, FeatureHeaderData, HyperlinkObject,
    MsoDrawingData, PlsRecord, PrinterSettings, RevisionLogStream, SstCompletion,
};

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

    /// Rebuilds physical BIFF positions, specification file pointers, and the
    /// BOF/EOF substream index after record-tree edits.
    pub fn relayout(&mut self) -> Result<()> {
        let mut rebuilt = self.clone();
        rebuilt.stream.relayout()?;
        rebuilt.reindex()?;
        *self = rebuilt;
        Ok(())
    }

    /// Aggregates the workbook-global `OfficeArtDggContainer` and the
    /// `OfficeArtDgContainer` values owned by sheet substreams.
    ///
    /// The BIFF and OfficeArt record trees remain the editable source of
    /// truth. A workbook without OfficeArt returns `None`; partial or
    /// ambiguous OfficeArt framing is rejected instead of being presented as
    /// a complete drawing graph.
    pub fn drawing_graph(&self) -> Result<Option<OfficeArtDrawingGraph>> {
        let mut drawing_groups = Vec::<&OfficeArtStream>::new();
        let mut drawings = Vec::<&OfficeArtStream>::new();
        let mut incomplete_kinds = Vec::new();

        for record in &self.stream.records {
            match &record.data {
                BiffRecordData::MsoDrawingGroup(value) => match &value.data {
                    MsoDrawingData::Complete(stream) => drawing_groups.push(stream),
                    MsoDrawingData::Partial(_) => incomplete_kinds.push("partial MsoDrawingGroup"),
                    MsoDrawingData::Incomplete { .. } => {
                        incomplete_kinds.push("incomplete MsoDrawingGroup")
                    }
                },
                BiffRecordData::MsoDrawing(value) => match &value.data {
                    MsoDrawingData::Complete(stream) => drawings.push(stream),
                    MsoDrawingData::Partial(_) => incomplete_kinds.push("partial MsoDrawing"),
                    MsoDrawingData::Incomplete { .. } => {
                        incomplete_kinds.push("incomplete MsoDrawing")
                    }
                },
                _ => {}
            }
        }

        if drawing_groups.is_empty() && drawings.is_empty() && incomplete_kinds.is_empty() {
            return Ok(None);
        }
        if !incomplete_kinds.is_empty() {
            return Err(Error::invalid(
                0,
                format!(
                    "XLS drawing graph contains non-complete OfficeArt aggregates: {}",
                    incomplete_kinds.join(", ")
                ),
            ));
        }
        let [drawing_group] = drawing_groups.as_slice() else {
            return Err(Error::invalid(
                0,
                format!(
                    "XLS drawing graph contains {} complete MsoDrawingGroup records, expected 1",
                    drawing_groups.len()
                ),
            ));
        };
        OfficeArtDrawingGraph::from_streams(drawing_group, &drawings).map(Some)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut tree = self.clone();
        tree.relayout()?;
        tree.stream.to_bytes()
    }
}

impl XlsWorkbookStream {
    /// Returns the complete drawing graph for this Workbook Stream when one
    /// is present.
    pub fn drawing_graph(&self) -> Result<Option<OfficeArtDrawingGraph>> {
        self.tree.drawing_graph()
    }
}

impl XlsFile {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Ok(Self::open_with_options(path, ParseOptions::default())?.into_value())
    }

    pub fn open_compatible(path: impl AsRef<Path>) -> Result<ParseOutcome<Self>> {
        Self::open_with_options(path, ParseOptions::compatible(Limits::default()))
    }

    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: ParseOptions,
    ) -> Result<ParseOutcome<Self>> {
        let compound = compound_from_path(path.as_ref(), options, BinaryFormat::Xls)?;
        Self::from_compound_outcome(compound, options)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Ok(Self::from_bytes_with_options(bytes, ParseOptions::default())?.into_value())
    }

    pub fn from_bytes_compatible(bytes: &[u8]) -> Result<ParseOutcome<Self>> {
        Self::from_bytes_with_options(bytes, ParseOptions::compatible(Limits::default()))
    }

    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        Ok(Self::from_bytes_with_options(bytes, ParseOptions::strict(limits))?.into_value())
    }

    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: ParseOptions,
    ) -> Result<ParseOutcome<Self>> {
        let compound = compound_from_bytes(bytes, options, BinaryFormat::Xls)?;
        Self::from_compound_outcome(compound, options)
    }

    pub fn from_compound_file(compound_file: CompoundFile) -> Result<Self> {
        Ok(
            Self::from_compound_file_with_options(compound_file, ParseOptions::default())?
                .into_value(),
        )
    }

    pub fn from_compound_file_compatible(
        compound_file: CompoundFile,
    ) -> Result<ParseOutcome<Self>> {
        Self::from_compound_file_with_options(
            compound_file,
            ParseOptions::compatible(Limits::default()),
        )
    }

    pub fn from_compound_file_with_limits(
        compound_file: CompoundFile,
        limits: Limits,
    ) -> Result<Self> {
        Ok(
            Self::from_compound_file_with_options(compound_file, ParseOptions::strict(limits))?
                .into_value(),
        )
    }

    pub fn from_compound_file_with_options(
        compound_file: CompoundFile,
        options: ParseOptions,
    ) -> Result<ParseOutcome<Self>> {
        let compound = compound_outcome(compound_file, options, BinaryFormat::Xls)?;
        Self::from_compound_outcome(compound, options)
    }

    /// Rebuilds every managed BIFF stream's derived physical layout after
    /// callers edit the public Rust record trees.
    pub fn relayout(&mut self) -> Result<()> {
        let mut rebuilt = self.clone();
        for workbook in &mut rebuilt.workbooks {
            workbook.tree.relayout()?;
        }
        if let Some(XlsRevisionLog::Parsed(log)) = &mut rebuilt.revision_log {
            log.relayout()?;
        }
        *self = rebuilt;
        Ok(())
    }

    fn from_compound_outcome(
        compound: ParseOutcome<CompoundFile>,
        options: ParseOptions,
    ) -> Result<ParseOutcome<Self>> {
        let ParseOutcome {
            value: compound_file,
            mut diagnostics,
        } = compound;
        let mut workbooks = Vec::new();
        for name in [XlsStreamName::Workbook, XlsStreamName::Book] {
            if let Some(bytes) = compound_file.stream(name.path()) {
                let workbook = XlsWorkbookStream {
                    name,
                    tree: BiffWorkbookTree::from_bytes_with_limits(bytes, options.limits)?,
                };
                audit_workbook(&workbook, options.is_strict(), &mut diagnostics)?;
                workbooks.push(workbook);
            }
        }
        if workbooks.is_empty() {
            return Err(Error::invalid(0, "Workbook/Book stream is missing"));
        }
        let revision_log = match compound_file.stream(super::REVISION_LOG_STREAM_PATH) {
            Some(bytes) => match RevisionLogStream::from_bytes_with_limits(bytes, options.limits) {
                Ok(value) => Some(XlsRevisionLog::Parsed(value)),
                Err(error) if options.is_strict() => {
                    let offset = error.offset().unwrap_or(0);
                    return Err(Error::invalid(
                        offset,
                        format!("Revision Stream violates MS-XLS 2.1.7.14: {error}"),
                    ));
                }
                Err(error) => {
                    let offset = error.offset().unwrap_or(0);
                    diagnostics.push(ParseDiagnostic::warning(
                        ParseDiagnosticCode::InvalidStreamPreserved,
                        BinaryFormat::Xls,
                        Some(super::REVISION_LOG_STREAM_PATH),
                        Some(offset),
                        "Revision Stream",
                        SpecificationReference {
                            document: "MS-XLS",
                            section: "2.1.7.14",
                        },
                        format!("preserved an invalid Revision Stream: {error}"),
                    ));
                    Some(XlsRevisionLog::Compatibility {
                        bytes: bytes.to_vec(),
                        reason: error.to_string(),
                    })
                }
            },
            None => None,
        };
        Ok(ParseOutcome::new(
            Self {
                compound_file,
                workbooks,
                revision_log,
            },
            diagnostics,
        ))
    }

    pub fn to_compound_file(&self) -> Result<CompoundFile> {
        self.to_compound_file_with_options(SaveOptions::default())
    }

    pub fn to_compound_file_preserving_compatibility(&self) -> Result<CompoundFile> {
        self.to_compound_file_with_options(SaveOptions::preserving_compatibility())
    }

    pub fn to_compound_file_with_options(&self, options: SaveOptions) -> Result<CompoundFile> {
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
        if !options.preserves_compatibility() {
            for workbook in &self.workbooks {
                audit_workbook(workbook, true, &mut Vec::new())?;
            }
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
            Some(XlsRevisionLog::Compatibility { .. }) if !options.preserves_compatibility() => {
                return Err(Error::invalid(
                    0,
                    "strict save rejects an invalid Revision Stream",
                ));
            }
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

    pub fn to_bytes_preserving_compatibility(&self) -> Result<Vec<u8>> {
        self.to_compound_file_preserving_compatibility()?.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.to_compound_file()?.save(path)
    }

    pub fn save_preserving_compatibility(&self, path: impl AsRef<Path>) -> Result<()> {
        self.to_compound_file_preserving_compatibility()?.save(path)
    }
}

fn audit_workbook(
    workbook: &XlsWorkbookStream,
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<()> {
    audit_workbook_topology(workbook, strict, diagnostics)?;
    let biff8 = workbook.tree.stream.is_biff8();
    for record in &workbook.tree.stream.records {
        match &record.data {
            BiffRecordData::Bof(value) if biff8 => {
                audit_bof(workbook, record, value, strict, diagnostics)?;
            }
            BiffRecordData::Formula(value) | BiffRecordData::Formula4Compatibility(value) => {
                audit_missing_formula_extra(
                    workbook,
                    record,
                    value.tokens.rgce.missing_extra_count(),
                    strict,
                    diagnostics,
                    "Formula",
                    "2.4.127",
                )?;
            }
            BiffRecordData::SharedFormula(value) => audit_missing_formula_extra(
                workbook,
                record,
                value.tokens.rgce.missing_extra_count(),
                strict,
                diagnostics,
                "ShrFmla",
                "2.4.260",
            )?,
            BiffRecordData::Array(value) => audit_missing_formula_extra(
                workbook,
                record,
                value.tokens.rgce.missing_extra_count(),
                strict,
                diagnostics,
                "Array",
                "2.4.4",
            )?,
            BiffRecordData::DataValidation(value) => audit_missing_formula_extra(
                workbook,
                record,
                value.formula1.tokens.missing_extra_count()
                    + value.formula2.tokens.missing_extra_count(),
                strict,
                diagnostics,
                "Dv",
                "2.4.95",
            )?,
            BiffRecordData::ConditionalFormatting(value) => audit_missing_formula_extra(
                workbook,
                record,
                value.formula1.missing_extra_count() + value.formula2.missing_extra_count(),
                strict,
                diagnostics,
                "CF",
                "2.4.42",
            )?,
            BiffRecordData::ConditionalFormatting12(value) => audit_missing_formula_extra(
                workbook,
                record,
                value.formula1.missing_extra_count()
                    + value.formula2.missing_extra_count()
                    + value.active_formula.missing_extra_count(),
                strict,
                diagnostics,
                "CF12",
                "2.4.43",
            )?,
            BiffRecordData::Name(value) => audit_missing_formula_extra(
                workbook,
                record,
                value.formula.missing_extra_count(),
                strict,
                diagnostics,
                "Lbl",
                "2.4.150",
            )?,
            BiffRecordData::ChartLinkedData(value) => audit_missing_formula_extra(
                workbook,
                record,
                value.formula.missing_extra_count(),
                strict,
                diagnostics,
                "BRAI",
                "2.4.29",
            )?,
            BiffRecordData::ChartEndObject(value)
                if !matches!(value.object_kind, 0x0010..=0x0012) =>
            {
                report_record_issue(
                    workbook,
                    record,
                    strict,
                    diagnostics,
                    ParseDiagnosticCode::NonconformingRecord,
                    "EndObject",
                    "2.4.101",
                    format!(
                        "iObjectKind is {:#06x}, outside the specified 0x0010..=0x0012 range",
                        value.object_kind
                    ),
                )?;
            }
            BiffRecordData::ExtSst(value) => {
                let nonzero_reserved = value
                    .buckets
                    .iter()
                    .filter(|bucket| bucket.reserved != 0)
                    .count();
                let invalid_offsets = value
                    .buckets
                    .iter()
                    .filter(|bucket| u32::from(bucket.record_offset) >= bucket.stream_offset)
                    .count();
                let invalid_buckets = value
                    .buckets
                    .iter()
                    .filter(|bucket| {
                        bucket.reserved != 0
                            || u32::from(bucket.record_offset) >= bucket.stream_offset
                    })
                    .count();
                if invalid_buckets != 0 {
                    report_record_issue(
                        workbook,
                        record,
                        strict,
                        diagnostics,
                        ParseDiagnosticCode::NonconformingRecord,
                        "ISSTInf",
                        "2.5.167",
                        format!(
                            "ExtSST contains {invalid_buckets} nonconforming bucket(s): {nonzero_reserved} with a nonzero reserved field and {invalid_offsets} with cbOffset not less than ib"
                        ),
                    )?;
                }
            }
            BiffRecordData::Hyperlink(value) => match &value.object {
                HyperlinkObject::Parsed { .. } => {}
                HyperlinkObject::Truncated { payload, .. } => report_record_issue(
                    workbook,
                    record,
                    strict,
                    diagnostics,
                    ParseDiagnosticCode::TruncatedRecord,
                    "HLink",
                    "2.4.140",
                    format!(
                        "Hyperlink Object is truncated with {} retained bytes",
                        payload.len()
                    ),
                )?,
                HyperlinkObject::TruncatedUrlMoniker {
                    declared_byte_length,
                    address,
                    ..
                } => report_record_issue(
                    workbook,
                    record,
                    strict,
                    diagnostics,
                    ParseDiagnosticCode::TruncatedRecord,
                    "HLink",
                    "2.4.140",
                    format!(
                        "URL moniker declares {declared_byte_length} bytes but only {} UTF-16 units are available",
                        address.len()
                    ),
                )?,
                HyperlinkObject::Compatibility(bytes) => report_record_issue(
                    workbook,
                    record,
                    strict,
                    diagnostics,
                    ParseDiagnosticCode::NonconformingRecord,
                    "HLink",
                    "2.4.140",
                    format!("Hyperlink Object has {} nonconforming bytes", bytes.len()),
                )?,
            },
            BiffRecordData::FeatureHeader(value)
                if matches!(value.data, FeatureHeaderData::Malformed { .. }) =>
            {
                report_record_issue(
                    workbook,
                    record,
                    strict,
                    diagnostics,
                    ParseDiagnosticCode::NonconformingRecord,
                    "FeatHdr",
                    "2.4.112",
                    "FeatHdr contains a marker or payload outside its shared-feature schema".into(),
                )?;
            }
            BiffRecordData::Pls(value) => {
                if let Some(devmode) = pls_devmode(value) {
                    // MS-RPRN 2.2.2.1 explicitly requires consumers to accept
                    // _DEVMODE values with truncated public information. Only
                    // missing bytes from the declared driver-private tail make
                    // the containing Pls record incomplete.
                    if !devmode.driver_extra_complete {
                        report_record_issue(
                            workbook,
                            record,
                            strict,
                            diagnostics,
                            ParseDiagnosticCode::TruncatedRecord,
                            "Pls",
                            "2.4.199",
                            format!(
                                "DEVMODEW declares {} driver-private bytes but only {} are available",
                                devmode.declared_driver_extra_size,
                                devmode.driver_extra.len()
                            ),
                        )?;
                    }
                }
            }
            BiffRecordData::MsoDrawingGroup(value) => audit_drawing(
                workbook,
                record,
                value,
                strict,
                diagnostics,
                "MsoDrawingGroup",
                "2.4.171",
            )?,
            BiffRecordData::MsoDrawing(value) => audit_drawing(
                workbook,
                record,
                value,
                strict,
                diagnostics,
                "MsoDrawing",
                "2.4.170",
            )?,
            BiffRecordData::Sst(value) => {
                if let SstCompletion::Truncated {
                    first_unparsed_string,
                    reason,
                } = &value.completion
                {
                    report_record_issue(
                        workbook,
                        record,
                        strict,
                        diagnostics,
                        ParseDiagnosticCode::TruncatedRecord,
                        "SST",
                        "2.4.265",
                        format!("SST stopped at string {first_unparsed_string}: {reason}"),
                    )?;
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn audit_bof(
    workbook: &XlsWorkbookStream,
    record: &BiffRecord,
    value: &super::BofRecord,
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<()> {
    let mut violations = Vec::new();
    if value.version != 0x0600 {
        violations.push(format!("vers is {:#06x}, expected 0x0600", value.version));
    }
    if !matches!(value.document_type, 0x0005 | 0x0010 | 0x0020 | 0x0040) {
        violations.push(format!(
            "dt is {:#06x}, outside the four specified substream kinds",
            value.document_type
        ));
    }
    if !matches!(value.build_year, 0x07cc | 0x07cd) {
        violations.push(format!(
            "rupYear is {:#06x}, expected 0x07cc or 0x07cd",
            value.build_year
        ));
    }
    let history = value.history_flags;
    let must_be_zero = (history & 0x0000_0136) | (history & 0xfff8_0000);
    if history & 1 == 0 || must_be_zero != 0 {
        violations.push(format!(
            "history flags violate fWin/reserved MUST values (raw {history:#010x})"
        ));
    }
    let highest_version = (history >> 14) & 0x0f;
    if !matches!(highest_version, 0 | 1 | 2 | 3 | 4 | 6 | 7) {
        violations.push(format!("verXLHigh has reserved value {highest_version:#x}"));
    }
    let lowest = value.lowest_version;
    let lowest_biff = lowest & 0xff;
    let last_saved = (lowest >> 8) & 0x0f;
    if lowest_biff != 6
        || !matches!(last_saved, 0 | 1 | 2 | 3 | 4 | 6 | 7)
        || last_saved > highest_version
        || lowest & 0xffff_f000 != 0
    {
        violations.push(format!(
            "verLowestBiff/verLastXLSaved/reserved values are invalid (raw {lowest:#010x}, verXLHigh {highest_version:#x})"
        ));
    }
    if violations.is_empty() {
        return Ok(());
    }
    report_record_issue(
        workbook,
        record,
        strict,
        diagnostics,
        ParseDiagnosticCode::NonconformingRecord,
        "BOF",
        "2.4.21",
        violations.join("; "),
    )
}

fn audit_workbook_topology(
    workbook: &XlsWorkbookStream,
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<()> {
    let first_offset = workbook
        .tree
        .stream
        .records
        .first()
        .map_or(0, |record| u64::from(record.offset));
    if workbook.name == XlsStreamName::Book {
        report_workbook_issue(
            workbook,
            first_offset,
            strict,
            diagnostics,
            ParseDiagnosticCode::NonconformingRecord,
            "Workbook Stream",
            "2.1.7.20",
            "legacy stream name is Book; MS-XLS requires Workbook".into(),
        )?;
    }

    if !workbook.tree.stream.is_biff8() {
        report_workbook_issue(
            workbook,
            first_offset,
            strict,
            diagnostics,
            ParseDiagnosticCode::NonconformingRecord,
            "Workbook Stream",
            "2.1.7.20",
            "legacy BIFF stream is outside the current MS-XLS Workbook Stream grammar".into(),
        )?;
        return Ok(());
    }

    if !workbook.tree.outside_substream_ranges.is_empty() {
        report_workbook_issue(
            workbook,
            first_offset,
            strict,
            diagnostics,
            ParseDiagnosticCode::NonconformingRecord,
            "Workbook Stream",
            "2.1.7.20",
            format!(
                "records outside BOF/EOF substreams occur in ranges {:?}",
                workbook.tree.outside_substream_ranges
            ),
        )?;
    }

    let roots = &workbook.tree.substreams;
    let globals_count = roots
        .iter()
        .filter(|node| node.kind == BiffSubstreamKind::WorkbookGlobals)
        .count();
    if globals_count != 1
        || roots
            .first()
            .is_none_or(|node| node.kind != BiffSubstreamKind::WorkbookGlobals)
    {
        report_workbook_issue(
            workbook,
            first_offset,
            strict,
            diagnostics,
            ParseDiagnosticCode::NonconformingRecord,
            "Globals Substream",
            "2.1.7.20.3",
            format!(
                "Workbook Stream has {globals_count} top-level Globals Substreams and the first substream kind is {:?}",
                roots.first().map(|node| node.kind)
            ),
        )?;
    }

    let following = roots.get(1..).unwrap_or_default();
    let invalid_following = following
        .iter()
        .enumerate()
        .filter_map(|(index, node)| {
            (!matches!(
                node.kind,
                BiffSubstreamKind::WorksheetOrDialogSheet
                    | BiffSubstreamKind::ChartSheet
                    | BiffSubstreamKind::MacroSheet
            ))
            .then_some((index + 1, node.kind))
        })
        .collect::<Vec<_>>();
    if following.is_empty() || !invalid_following.is_empty() {
        report_workbook_issue(
            workbook,
            first_offset,
            strict,
            diagnostics,
            ParseDiagnosticCode::NonconformingRecord,
            "Workbook Stream",
            "2.1.7.20",
            format!(
                "Workbook Stream has {} following sheet substreams; invalid (index, kind) values are {invalid_following:?}",
                following.len()
            ),
        )?;
    }
    Ok(())
}

fn audit_missing_formula_extra(
    workbook: &XlsWorkbookStream,
    record: &BiffRecord,
    missing_extra_count: usize,
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
    structure: &'static str,
    section: &'static str,
) -> Result<()> {
    if missing_extra_count == 0 {
        return Ok(());
    }
    report_record_issue(
        workbook,
        record,
        strict,
        diagnostics,
        ParseDiagnosticCode::TruncatedRecord,
        structure,
        section,
        format!(
            "formula is missing {missing_extra_count} required MS-XLS 2.5.198.103 RgbExtra structures"
        ),
    )
}

fn pls_devmode(value: &PlsRecord) -> Option<&DevModeW> {
    match &value.settings {
        PrinterSettings::WindowsUnicode(value)
        | PrinterSettings::LengthPrefixedWindowsUnicode { devmode: value, .. } => Some(value),
        _ => None,
    }
}

fn audit_drawing(
    workbook: &XlsWorkbookStream,
    record: &BiffRecord,
    value: &super::MsoDrawingRecord,
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
    structure: &'static str,
    section: &'static str,
) -> Result<()> {
    let message = match &value.data {
        MsoDrawingData::Complete(_) => return Ok(()),
        MsoDrawingData::Partial(value) => format!(
            "OfficeArt sequence is partial: {} incomplete records and {} unparsed bytes",
            value.incomplete_record_count(),
            value.unparsed_byte_count()
        ),
        MsoDrawingData::Incomplete { bytes, reason } => {
            format!(
                "OfficeArt sequence retains {} incomplete bytes: {reason}",
                bytes.len()
            )
        }
    };
    report_record_issue(
        workbook,
        record,
        strict,
        diagnostics,
        ParseDiagnosticCode::TruncatedRecord,
        structure,
        section,
        message,
    )
}

#[allow(clippy::too_many_arguments)]
fn report_record_issue(
    workbook: &XlsWorkbookStream,
    record: &BiffRecord,
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
    code: ParseDiagnosticCode,
    structure: &'static str,
    section: &'static str,
    message: String,
) -> Result<()> {
    report_workbook_issue(
        workbook,
        u64::from(record.offset),
        strict,
        diagnostics,
        code,
        structure,
        section,
        message,
    )
}

#[allow(clippy::too_many_arguments)]
fn report_workbook_issue(
    workbook: &XlsWorkbookStream,
    offset: u64,
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
    code: ParseDiagnosticCode,
    structure: &'static str,
    section: &'static str,
    message: String,
) -> Result<()> {
    if strict {
        return Err(Error::invalid(
            offset,
            format!(
                "{} violates MS-XLS {section}: {message}",
                workbook.name.path()
            ),
        ));
    }
    diagnostics.push(ParseDiagnostic::warning(
        code,
        BinaryFormat::Xls,
        Some(workbook.name.path()),
        Some(offset),
        structure,
        SpecificationReference {
            document: "MS-XLS",
            section,
        },
        message,
    ));
    Ok(())
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
    use crate::{
        cfb::Version,
        xls::{
            BofRecord, CellHeader, ChartEndObjectRecord, DevModeFields, DevModeWPublic,
            ExtSstRecord, FormulaCachedResult, FormulaRecord, FormulaTokenStream, FormulaTokens,
            FrtFlags, FrtHeaderOld, IsstInf,
        },
    };

    fn record(offset: u32, data: BiffRecordData) -> BiffRecord {
        BiffRecord { offset, data }
    }

    fn bof(document_type: u16) -> BofRecord {
        BofRecord {
            version: 0x0600,
            document_type,
            build_identifier: 0,
            build_year: 0x07cc,
            history_flags: 1,
            lowest_version: 6,
        }
    }

    #[test]
    fn indexes_nested_bof_eof_without_flattening_records() {
        let stream = BiffStream {
            records: vec![
                record(0, BiffRecordData::Bof(bof(0x0010))),
                record(20, BiffRecordData::Bof(bof(0x0020))),
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
        tree.relayout().unwrap();
        assert_eq!(tree.substreams[0].record_range, 0..5);
        assert_eq!(tree.substreams[0].children[0].record_range, 1..4);
        assert_eq!(tree.stream.records[2].offset, 40);
        assert_eq!(tree.to_bytes().unwrap().len(), 54);
    }

    fn workbook_bytes() -> Vec<u8> {
        BiffStream {
            records: vec![
                record(0, BiffRecordData::Bof(bof(0x0005))),
                record(20, BiffRecordData::Eof),
                record(24, BiffRecordData::Bof(bof(0x0010))),
                record(44, BiffRecordData::Eof),
            ],
            trailing_padding: Vec::new(),
        }
        .to_bytes()
        .unwrap()
    }

    #[test]
    fn invalid_revision_stream_is_preserved_only_in_compatible_mode() {
        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound
            .create_or_replace_stream(WORKBOOK_STREAM, workbook_bytes())
            .unwrap();
        compound
            .create_or_replace_stream(super::super::REVISION_LOG_STREAM_PATH, vec![1, 2, 3])
            .unwrap();

        assert!(XlsFile::from_compound_file(compound.clone()).is_err());
        let outcome = XlsFile::from_compound_file_compatible(compound).unwrap();
        assert!(matches!(
            outcome.value.revision_log,
            Some(XlsRevisionLog::Compatibility { ref bytes, .. }) if bytes == &[1, 2, 3]
        ));
        assert_eq!(outcome.diagnostics.len(), 1);
        assert_eq!(
            outcome.diagnostics[0].code,
            ParseDiagnosticCode::InvalidStreamPreserved
        );
        assert_eq!(
            outcome.diagnostics[0].location.path.as_deref(),
            Some(super::super::REVISION_LOG_STREAM_PATH)
        );
        assert!(outcome.value.to_compound_file().is_err());
        let preserved = outcome
            .value
            .to_compound_file_preserving_compatibility()
            .unwrap();
        assert_eq!(
            preserved.stream(super::super::REVISION_LOG_STREAM_PATH),
            Some([1, 2, 3].as_slice())
        );
    }

    #[test]
    fn workbook_stream_lookup_uses_cfb_case_insensitive_names() {
        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound
            .create_or_replace_stream("/WORKBOOK", workbook_bytes())
            .unwrap();

        let file = XlsFile::from_compound_file(compound).unwrap();
        assert_eq!(file.workbooks.len(), 1);
        assert_eq!(file.workbooks[0].name, XlsStreamName::Workbook);
        let rebuilt = file.to_compound_file().unwrap();
        assert_eq!(rebuilt.entry("/Workbook").unwrap().name, "WORKBOOK");
    }

    #[test]
    fn workbook_name_and_substream_cardinality_use_compatible_diagnostics() {
        let mut legacy_name = CompoundFile::new(Version::V3).unwrap();
        legacy_name
            .create_or_replace_stream(BOOK_STREAM, workbook_bytes())
            .unwrap();
        assert!(XlsFile::from_compound_file(legacy_name.clone()).is_err());
        let outcome = XlsFile::from_compound_file_compatible(legacy_name).unwrap();
        assert_eq!(outcome.diagnostics.len(), 1);
        assert_eq!(outcome.diagnostics[0].structure, "Workbook Stream");
        assert_eq!(outcome.diagnostics[0].specification.section, "2.1.7.20");
        assert!(outcome.value.to_compound_file().is_err());
        assert!(
            outcome
                .value
                .to_compound_file_preserving_compatibility()
                .is_ok()
        );

        let sheet_only = BiffStream {
            records: vec![
                record(0, BiffRecordData::Bof(bof(0x0010))),
                record(20, BiffRecordData::Eof),
            ],
            trailing_padding: Vec::new(),
        }
        .to_bytes()
        .unwrap();
        let mut invalid_topology = CompoundFile::new(Version::V3).unwrap();
        invalid_topology
            .create_or_replace_stream(WORKBOOK_STREAM, sheet_only)
            .unwrap();
        assert!(XlsFile::from_compound_file(invalid_topology.clone()).is_err());
        let outcome = XlsFile::from_compound_file_compatible(invalid_topology).unwrap();
        assert_eq!(outcome.diagnostics.len(), 2);
        assert_eq!(outcome.diagnostics[0].structure, "Globals Substream");
        assert_eq!(outcome.diagnostics[0].specification.section, "2.1.7.20.3");
        assert_eq!(outcome.diagnostics[1].structure, "Workbook Stream");
    }

    #[test]
    fn bof_must_fields_use_the_root_strictness_gate() {
        let mut workbook = XlsWorkbookStream {
            name: XlsStreamName::Workbook,
            tree: BiffWorkbookTree::from_bytes(&workbook_bytes()).unwrap(),
        };
        let BiffRecordData::Bof(value) = &mut workbook.tree.stream.records[0].data else {
            panic!("first record is not BOF");
        };
        value.build_year = 0;
        value.history_flags = 0;
        value.lowest_version = 0;

        assert!(audit_workbook(&workbook, true, &mut Vec::new()).is_err());
        let mut diagnostics = Vec::new();
        audit_workbook(&workbook, false, &mut diagnostics).unwrap();
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0].structure, "BOF");
        assert_eq!(diagnostics[0].specification.section, "2.4.21");
    }

    #[test]
    fn end_object_kind_uses_the_root_strictness_gate() {
        let workbook = XlsWorkbookStream {
            name: XlsStreamName::Workbook,
            tree: BiffWorkbookTree::from_stream(BiffStream {
                records: vec![
                    record(0, BiffRecordData::Bof(bof(0x0005))),
                    record(20, BiffRecordData::Eof),
                    record(24, BiffRecordData::Bof(bof(0x0020))),
                    record(
                        44,
                        BiffRecordData::ChartEndObject(ChartEndObjectRecord {
                            header: FrtHeaderOld {
                                record_type: 0x0855,
                                flags: FrtFlags::empty(),
                            },
                            object_kind: 0x0013,
                            unused1: None,
                            unused2: None,
                            unused3: None,
                        }),
                    ),
                    record(54, BiffRecordData::Eof),
                ],
                trailing_padding: Vec::new(),
            })
            .unwrap(),
        };

        assert!(audit_workbook(&workbook, true, &mut Vec::new()).is_err());
        let mut diagnostics = Vec::new();
        audit_workbook(&workbook, false, &mut diagnostics).unwrap();
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0].structure, "EndObject");
        assert_eq!(diagnostics[0].specification.section, "2.4.101");
    }

    #[test]
    fn ext_sst_bucket_must_fields_use_the_root_strictness_gate() {
        let workbook = XlsWorkbookStream {
            name: XlsStreamName::Workbook,
            tree: BiffWorkbookTree::from_stream(BiffStream {
                records: vec![
                    record(0, BiffRecordData::Bof(bof(0x0005))),
                    record(
                        20,
                        BiffRecordData::ExtSst(ExtSstRecord {
                            strings_per_bucket: 8,
                            buckets: vec![IsstInf {
                                stream_offset: 4,
                                record_offset: 4,
                                reserved: 1,
                            }],
                        }),
                    ),
                    record(34, BiffRecordData::Eof),
                ],
                trailing_padding: Vec::new(),
            })
            .unwrap(),
        };

        assert!(audit_workbook(&workbook, true, &mut Vec::new()).is_err());
        let mut diagnostics = Vec::new();
        audit_workbook(&workbook, false, &mut diagnostics).unwrap();
        let issue = diagnostics
            .iter()
            .find(|diagnostic| diagnostic.structure == "ISSTInf")
            .expect("ISSTInf diagnostic");
        assert_eq!(issue.specification.section, "2.5.167");
    }

    #[test]
    fn spec_truncated_devmode_public_fields_are_not_a_damage_diagnostic() {
        let devmode = DevModeW {
            device_name: [0; 32],
            specification_version: 0x0401,
            driver_version: 0,
            declared_public_size: 76,
            declared_driver_extra_size: 0,
            fields: DevModeFields::empty(),
            public_fields: DevModeWPublic::Truncated(Vec::new()),
            driver_extra: Vec::new(),
            driver_extra_complete: true,
            trailing: Vec::new(),
        };
        let workbook = |devmode| XlsWorkbookStream {
            name: XlsStreamName::Workbook,
            tree: BiffWorkbookTree::from_stream(BiffStream {
                records: vec![
                    record(0, BiffRecordData::Bof(bof(0x0005))),
                    record(
                        20,
                        BiffRecordData::Pls(PlsRecord {
                            reserved: 0,
                            settings: PrinterSettings::WindowsUnicode(devmode),
                            physical_segment_lengths: vec![78],
                        }),
                    ),
                    record(102, BiffRecordData::Eof),
                    record(106, BiffRecordData::Bof(bof(0x0010))),
                    record(126, BiffRecordData::Eof),
                ],
                trailing_padding: Vec::new(),
            })
            .unwrap(),
        };

        let valid = workbook(devmode.clone());
        let mut diagnostics = Vec::new();
        audit_workbook(&valid, true, &mut diagnostics).unwrap();
        assert!(diagnostics.is_empty());

        let mut incomplete = devmode;
        incomplete.declared_driver_extra_size = 4;
        incomplete.driver_extra_complete = false;
        let invalid = workbook(incomplete);
        assert!(audit_workbook(&invalid, true, &mut Vec::new()).is_err());
        let mut diagnostics = Vec::new();
        audit_workbook(&invalid, false, &mut diagnostics).unwrap();
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0].code, ParseDiagnosticCode::TruncatedRecord);
    }

    #[test]
    fn missing_formula_rgb_extra_requires_explicit_compatibility() {
        let formula = FormulaRecord {
            cell: CellHeader {
                row: 0,
                column: 0,
                format_index: 0,
            },
            cached_result: FormulaCachedResult::NumberBits(0),
            flags: 0,
            calculation_chain_id: 0,
            tokens: FormulaTokens {
                rgce: FormulaTokenStream::from_bytes(&[0x60, 0, 0, 0, 0, 0, 0, 0]).unwrap(),
                rgcb_tail: Vec::new(),
            },
        };
        let bytes = BiffStream {
            records: vec![
                record(0, BiffRecordData::Bof(bof(0x0005))),
                record(20, BiffRecordData::Formula(formula)),
                record(54, BiffRecordData::Eof),
                record(58, BiffRecordData::Bof(bof(0x0010))),
                record(78, BiffRecordData::Eof),
            ],
            trailing_padding: Vec::new(),
        }
        .to_bytes()
        .unwrap();
        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound
            .create_or_replace_stream(WORKBOOK_STREAM, bytes)
            .unwrap();

        assert!(XlsFile::from_compound_file(compound.clone()).is_err());
        let outcome = XlsFile::from_compound_file_compatible(compound).unwrap();
        assert_eq!(outcome.diagnostics.len(), 1);
        assert_eq!(
            outcome.diagnostics[0].code,
            ParseDiagnosticCode::TruncatedRecord
        );
        assert_eq!(outcome.diagnostics[0].specification.section, "2.4.127");
        assert!(outcome.value.to_compound_file().is_err());
        assert!(
            outcome
                .value
                .to_compound_file_preserving_compatibility()
                .is_ok()
        );
    }
}
