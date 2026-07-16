//! Typed file and primary content-tree roots for the Word binary format.

use std::{
    collections::{BTreeMap, BTreeSet},
    ops::Range,
    path::{Path, PathBuf},
};

use crate::{
    Error, Result,
    cfb::CompoundFile,
    io::BinaryFormat,
    limits::Limits,
    parse::{
        ParseDiagnostic, ParseDiagnosticCode, ParseOptions, ParseOutcome, SpecificationReference,
        compound_from_bytes, compound_from_path, compound_outcome,
    },
    shared::MsoEnvelope,
};

use super::{
    AnnotationBookmarks, AnnotationExtendedData, AnnotationOwners, AnnotationReferenceTable,
    AssociatedStrings, AutoCaptionDefinitions, AutoSummaryRangeTable, Bookmarks,
    CaptionDefinitions, ChpxFkp, ChpxFkpRun, Clx, CommandCustomizations, CpOnlyTable,
    DocOfficeArtContent, DocumentProperties, EmbeddedFontTable, ExternalFileNameTable, Fib,
    FibBaseFlags, FibFcLcb, FieldDocumentPart, FieldTable, FkpPageNumber, FontTable,
    FormatConsistencyBookmarks, FrameAndListRecords, GrammarCheckerCookieTable, GrammarCookieStore,
    GrammarOptionSets, GrammarStateTable, GrpPrl, HeaderTextTable, KnownSprm,
    LanguageDetectionStateTable, LegacyGrammarCheckerCookieTable, LegacyGrammarOptionSets,
    ListDefinitions, ListNamesTable, ListOverrides, ListStyleTemplates, MailMergeState,
    NilPicfAndBinData, NilPicfFieldType, NoteReferenceTable, OfficeDataSource, OleControlInfos,
    OleObjectDescriptor, PapxFkp, PapxFkpRun, ParagraphGroupProperties, Picf, PicfAndOfficeArtData,
    PlcBte, PlcfSed, PrcData, PrinterDriverInfo, PrivateFieldType, Prm, RangeProtection,
    RepairBookmarks, RevisionAuthors, RevisionMessageThreading, RevisionSaveIdTable, SaveHistory,
    SelectionState, Sepx, ShapeAnchorTable, SmartTagBookmarks, SmartTagData,
    SmartTagRecognizerStateTable, SpellingStateTable, SprmGroup, SprmKind, SprmOperand,
    StructuredTagBookmarks, StyleFormatting, StyleSheet, SubdocumentTable,
    TableCharacterCacheTable, TextPiece, TextPieceCharacters, TextboxBreakTable,
    TextboxDocumentPart, TextboxStoryTable, UserInputMethods, UserVariables, XmlSchemaReferences,
    XmlTransformPath,
};

const WORD_DOCUMENT_STREAM: &str = "/WordDocument";
const TABLE0_STREAM: &str = "/0Table";
const TABLE1_STREAM: &str = "/1Table";
const DATA_STREAM: &str = "/Data";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DocTableStreamName {
    Table0,
    Table1,
}

impl DocTableStreamName {
    pub const fn path(self) -> &'static str {
        match self {
            Self::Table0 => TABLE0_STREAM,
            Self::Table1 => TABLE1_STREAM,
        }
    }
}

/// A typed value together with the FIB location that owns its physical bytes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocLocated<T> {
    pub location: FibFcLcb,
    pub value: T,
}

/// A typed 512-byte formatting page referenced by a PlcBte.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocFkpPage<T> {
    pub page: FkpPageNumber,
    pub value: T,
}

/// A text piece retains CP/FC boundaries and its physical character encoding.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocTextPiece {
    pub piece_index: usize,
    pub value: TextPiece,
}

/// A CHPX FKP text run mapped from its physical FC interval to the document
/// CP coordinate space. Pcd.Prm and style-derived properties remain separate
/// specification layers and are not folded into this value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocChpxRun {
    pub cp_start: u32,
    pub cp_end: u32,
    pub properties: Option<GrpPrl>,
}

/// A PAPX FKP paragraph, table-row, or table-cell run mapped from its physical
/// FC interval to the document CP coordinate space.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocPapxRun {
    pub cp_start: u32,
    pub cp_end: u32,
    pub paragraph_height_info: [u8; 12],
    pub properties: Option<super::PapxInFkp>,
}

/// The two normative direct-formatting layers at one MS-DOC character
/// position. This is intentionally not an "effective formatting" value:
/// styles, lists, table styles, and conditional table formatting remain
/// separate specification layers.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocDirectFormatting {
    pub part: FieldDocumentPart,
    pub local_cp: u32,
    pub global_cp: u32,
    pub piece_index: usize,
    pub paragraph: DocDirectParagraphFormatting,
    pub character: DocDirectCharacterFormatting,
}

/// Direct paragraph formatting in the order and source layers defined by
/// MS-DOC 2.4.6.1. Keeping PAPX and Pcd.Prm properties separate permits an
/// editor to write a change back to its original physical owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocDirectParagraphFormatting {
    pub style_index: u16,
    pub papx_properties: GrpPrl,
    pub piece_properties: GrpPrl,
    /// The normative direct-paragraph property array after recursively
    /// following sprmPHugePapx/sprmPTableProps and applying their stop rules.
    /// The physical source arrays above and the referenced `DocDataNode`s are
    /// retained unchanged for precise write-back.
    pub applied_properties: GrpPrl,
}

/// Direct character formatting in the order and source layers defined by
/// MS-DOC 2.4.6.1.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocDirectCharacterFormatting {
    pub chpx_properties: GrpPrl,
    pub piece_properties: GrpPrl,
}

/// The table-membership values produced by applying the direct paragraph
/// property array. `depth_is_explicit` distinguishes the default depth zero
/// from a value written by sprmPItap/sprmPDtap; this matters for diagnosing
/// legacy or nonconforming producers without inventing an implicit depth.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct DocDirectTableState {
    pub in_table: bool,
    pub depth: u32,
    pub depth_is_explicit: bool,
}

/// The inherited property arrays for one STSH style. `lineage` is ordered
/// base-first and the three property arrays follow the same order, so later
/// properties from the requested style retain normal SPRM precedence.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocStyleProperties {
    pub style_index: u16,
    pub style_kind: super::StyleKind,
    pub lineage: Vec<u16>,
    pub paragraph_properties: GrpPrl,
    pub character_properties: GrpPrl,
    pub table_properties: GrpPrl,
}

impl DocDirectParagraphFormatting {
    /// Applies sprmPFInTable, sprmPItap, and sprmPDtap in specification order.
    pub fn table_state(&self) -> Result<DocDirectTableState> {
        let mut in_table = false;
        let mut depth = 0i32;
        let mut depth_is_explicit = false;
        for property in &self.applied_properties.properties {
            match property.sprm.kind() {
                SprmKind::Known(KnownSprm::PFInTable) => {
                    let SprmOperand::Byte(value) = &property.operand else {
                        return Err(Error::invalid(0, "sprmPFInTable operand is not Bool8"));
                    };
                    if *value > 1 {
                        return Err(Error::invalid(0, "sprmPFInTable Bool8 operand exceeds one"));
                    }
                    in_table = *value != 0;
                }
                SprmKind::Known(KnownSprm::PItap) => {
                    let SprmOperand::Dword(value) = &property.operand else {
                        return Err(Error::invalid(0, "sprmPItap operand is not a signed dword"));
                    };
                    depth = i32::from_le_bytes(*value);
                    depth_is_explicit = true;
                    if depth < 0 {
                        return Err(Error::invalid(0, "sprmPItap table depth is negative"));
                    }
                }
                SprmKind::Known(KnownSprm::PDtap) => {
                    let SprmOperand::Dword(value) = &property.operand else {
                        return Err(Error::invalid(0, "sprmPDtap operand is not a signed dword"));
                    };
                    depth = depth
                        .checked_add(i32::from_le_bytes(*value))
                        .ok_or_else(|| Error::Limit("sprmPDtap table depth overflow".into()))?;
                    depth_is_explicit = true;
                    if depth < 0 {
                        return Err(Error::invalid(
                            0,
                            "sprmPDtap produces a negative table depth",
                        ));
                    }
                }
                _ => {}
            }
        }
        Ok(DocDirectTableState {
            in_table,
            depth: u32::try_from(depth)
                .map_err(|_| Error::invalid(0, "direct table depth is negative"))?,
            depth_is_explicit,
        })
    }
}

/// Section properties live in WordDocument while their SED index lives in the
/// selected Table stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocSectionProperties {
    pub section_index: usize,
    pub offset: i32,
    pub physical_len: usize,
    pub value: Option<Sepx>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocWordDocumentStream {
    pub fib: Fib,
    pub text_pieces: Vec<DocTextPiece>,
    /// Aggregated CHPX FKP runs. `None` is reserved for compatible-mode input
    /// whose nonconforming FKP ordering cannot form a CP tree.
    pub chpx_runs: Option<Vec<DocChpxRun>>,
    /// Aggregated PAPX FKP runs. `None` is reserved for compatible-mode input
    /// whose nonconforming FKP ordering cannot form a CP tree.
    pub papx_runs: Option<Vec<DocPapxRun>>,
    character_format_pages: Vec<DocFkpPage<ChpxFkp>>,
    paragraph_format_pages: Vec<DocFkpPage<PapxFkp>>,
    pub section_properties: Vec<DocSectionProperties>,
    physical_bytes: Vec<u8>,
    source_fib_len: usize,
    // Maps each current PlcPcd piece to its immutable index in the source CLX.
    // Current indices stay contiguous when an edit removes an entire piece.
    source_piece_indices: Vec<usize>,
    source_chpx_runs: Option<Vec<DocChpxRun>>,
    source_papx_runs: Option<Vec<DocPapxRun>>,
    // Each edit is expressed in the piece coordinate space produced by all
    // preceding edits. The ordered map is part of save semantics because FKP
    // FC boundaries still refer to the source WordDocument until layout.
    pending_text_edits: BTreeMap<usize, Vec<CpReplacement>>,
    rebuild_character_formatting: bool,
    rebuild_paragraph_formatting: bool,
}

impl DocWordDocumentStream {
    /// Physical ChpxFkp pages retained for byte-preserving diagnostics. Edit
    /// `chpx_runs` to change character formatting.
    pub fn chpx_fkp_pages(&self) -> &[DocFkpPage<ChpxFkp>] {
        &self.character_format_pages
    }

    /// Physical PapxFkp pages retained for byte-preserving diagnostics. Edit
    /// `papx_runs` to change paragraph, table-row, or table-cell formatting.
    pub fn papx_fkp_pages(&self) -> &[DocFkpPage<PapxFkp>] {
        &self.paragraph_format_pages
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ParagraphMarkEdit {
    PreserveAll,
    ExplicitPapx,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocTableStream {
    pub name: DocTableStreamName,
    pub clx: DocLocated<Clx>,
    pub character_bin_table: DocLocated<PlcBte>,
    pub paragraph_bin_table: DocLocated<PlcBte>,
    pub sections: DocLocated<PlcfSed>,
    pub styles: Option<DocLocated<StyleSheet>>,
    pub fonts: Option<DocLocated<FontTable>>,
    pub fields: BTreeMap<FieldDocumentPart, DocLocated<FieldTable>>,
    pub bookmarks: Option<DocLocatedBookmarks>,
    pub header_text: Option<DocLocated<HeaderTextTable>>,
    pub footnotes: Option<DocNoteTables>,
    pub endnotes: Option<DocNoteTables>,
    pub annotations: Option<DocAnnotationTables>,
    pub annotation_owners: Option<DocLocated<AnnotationOwners>>,
    pub annotation_bookmarks: Option<DocLocatedAnnotationBookmarks>,
    pub annotation_extended_data: Option<DocLocated<AnnotationExtendedData>>,
    pub textbox_stories: BTreeMap<TextboxDocumentPart, DocLocated<TextboxStoryTable>>,
    pub textbox_breaks: BTreeMap<TextboxDocumentPart, DocLocated<TextboxBreakTable>>,
    pub shape_anchors: BTreeMap<TextboxDocumentPart, DocLocated<ShapeAnchorTable>>,
    pub office_art: Option<DocLocated<DocOfficeArtContent>>,
    pub revision_authors: Option<DocLocated<RevisionAuthors>>,
    pub captions: Option<DocCaptionTables>,
    pub subdocuments: Option<DocLocated<SubdocumentTable>>,
    pub user_variables: Option<DocLocated<UserVariables>>,
    pub embedded_fonts: Option<DocLocated<EmbeddedFontTable>>,
    pub spelling_state: Option<DocLocated<SpellingStateTable>>,
    pub grammar_state: Option<DocLocated<GrammarStateTable>>,
    pub language_detection_state: Option<DocLocated<LanguageDetectionStateTable>>,
    pub list_definitions: Option<DocListDefinitions>,
    pub list_names: Option<DocLocated<ListNamesTable>>,
    pub list_overrides: Option<DocLocated<ListOverrides>>,
    pub document_properties: Option<DocLocated<DocumentProperties>>,
    pub associated_strings: Option<DocLocated<AssociatedStrings>>,
    pub external_file_names: Option<DocLocated<ExternalFileNameTable>>,
    pub mail_merge_state: Option<DocLocated<MailMergeState>>,
    pub new_mail_merge_state: Option<DocLocated<MailMergeState>>,
    pub office_data_source: Option<DocLocated<OfficeDataSource>>,
    pub printer_driver_info: Option<DocLocated<PrinterDriverInfo>>,
    pub ole_control_infos: Option<DocLocated<OleControlInfos>>,
    pub table_character_cache: Option<DocLocated<TableCharacterCacheTable>>,
    pub revision_message_threading: Option<DocLocated<RevisionMessageThreading>>,
    pub list_style_templates: Option<DocLocated<ListStyleTemplates>>,
    pub frame_and_list_records: Option<DocLocated<FrameAndListRecords>>,
    pub grammar_option_sets: Option<DocLocated<GrammarOptionSets>>,
    pub legacy_grammar_option_sets: Option<DocLocated<LegacyGrammarOptionSets>>,
    pub auto_summary_ranges: Option<DocLocated<AutoSummaryRangeTable>>,
    pub smart_tag_recognizer_state: Option<DocLocated<SmartTagRecognizerStateTable>>,
    pub xml_schema_references: Option<DocLocated<XmlSchemaReferences>>,
    pub xml_transform_path: Option<DocLocated<XmlTransformPath>>,
    pub paragraph_group_properties: Option<DocLocated<ParagraphGroupProperties>>,
    pub save_history: Option<DocLocated<SaveHistory>>,
    pub grammar_checker_cookies: Option<DocLocated<GrammarCheckerCookieTable>>,
    pub legacy_grammar_checker_cookies: Option<DocLocated<LegacyGrammarCheckerCookieTable>>,
    pub grammar_cookie_data: Option<DocLocated<GrammarCookieStore>>,
    pub smart_tag_data: Option<DocLocated<SmartTagData>>,
    pub revision_save_ids: Option<DocLocated<RevisionSaveIdTable>>,
    pub selection_state: Option<DocLocated<SelectionState>>,
    pub command_customizations: Option<DocLocated<CommandCustomizations>>,
    pub structured_tag_bookmarks: Option<DocBookmarkSet<StructuredTagBookmarks>>,
    pub range_protection: Option<DocRangeProtectionTables>,
    pub smart_tag_bookmarks: Option<DocBookmarkSet<SmartTagBookmarks>>,
    pub format_consistency_bookmarks: Option<DocBookmarkSet<FormatConsistencyBookmarks>>,
    pub repair_bookmarks: Option<DocBookmarkSet<RepairBookmarks>>,
    pub user_input_methods: Option<DocUserInputMethodTables>,
    pub mso_envelope: Option<DocLocated<MsoEnvelope>>,
    pub deprecated_numbering_field_cache: Option<DocDeprecatedNumberingFieldCache>,
    /// Nonconforming FIB-referenced payloads retained only in compatible mode.
    pub compatibility_tables: Vec<DocCompatibilityTable>,
    physical_bytes: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocLocatedBookmarks {
    pub names_location: FibFcLcb,
    pub starts_location: FibFcLcb,
    pub ends_location: FibFcLcb,
    pub value: Bookmarks,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocLocatedAnnotationBookmarks {
    pub infos_location: FibFcLcb,
    pub starts_location: FibFcLcb,
    pub ends_location: FibFcLcb,
    pub value: AnnotationBookmarks,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocCompatibilityTable {
    pub label: String,
    pub location: FibFcLcb,
    pub physical_bytes: Option<Vec<u8>>,
    pub reason: String,
}

/// A note story is indexed by a reference PLC and a CP-only text PLC.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocNoteTables {
    pub references: DocLocated<NoteReferenceTable>,
    pub text: DocLocated<CpOnlyTable>,
}

/// Comment references and their corresponding comment-story boundaries.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocAnnotationTables {
    pub references: DocLocated<AnnotationReferenceTable>,
    pub text: DocLocated<CpOnlyTable>,
}

/// User-defined caption labels and automatic-caption mappings.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocCaptionTables {
    pub definitions: DocLocated<CaptionDefinitions>,
    pub automatic: DocLocated<AutoCaptionDefinitions>,
}

/// `PlfLst` can place its variable-length LVL array immediately after the
/// FIB-declared range; both physical regions form one logical list tree.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocListDefinitions {
    pub location: FibFcLcb,
    pub value: ListDefinitions,
    trailing_levels_len: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocBookmarkSet<T> {
    pub metadata_location: FibFcLcb,
    pub starts_location: FibFcLcb,
    pub ends_location: FibFcLcb,
    pub value: T,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocRangeProtectionTables {
    pub permissions_location: FibFcLcb,
    pub starts_location: FibFcLcb,
    pub ends_location: FibFcLcb,
    pub users_location: FibFcLcb,
    pub value: RangeProtection,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocUserInputMethodTables {
    pub methods_location: FibFcLcb,
    pub service_guids_location: FibFcLcb,
    pub value: UserInputMethods,
}

/// MS-DOC marks this cache as deprecated and says it SHOULD be ignored. Its
/// bounded bytes remain explicit without inventing an undocumented layout.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocDeprecatedNumberingFieldCache {
    pub location: FibFcLcb,
    pub physical_bytes: Vec<u8>,
}

/// MS-DOC assigns structure-specific offsets into this stream. Until each
/// referenced payload is promoted, the stream remains one explicit physical
/// node rather than being mislabeled as arbitrary content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocDataStream {
    pub physical_bytes: Vec<u8>,
    pub nodes: Vec<DocDataNode>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocDataNode {
    pub offset: u32,
    pub physical_len: usize,
    pub value: DocDataNodeValue,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DocDataNodeValue {
    Picture(PicfAndOfficeArtData),
    Binary(Box<NilPicfAndBinData>),
    ParagraphProperties(PrcData),
}

struct RebuiltDataStream {
    bytes: Option<Vec<u8>>,
    relocations: BTreeMap<u32, u32>,
}

/// The MS-DOC ObjectPool storage aggregates each embedded object storage with
/// its required ObjInfo/ODT stream. Other streams remain CFB-managed payloads
/// for their owning format libraries and are identified by `entry_paths`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocObjectPoolStorage {
    pub path: PathBuf,
    pub objects: Vec<DocEmbeddedObjectStorage>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocEmbeddedObjectStorage {
    pub path: PathBuf,
    pub descriptor_stream_path: PathBuf,
    pub descriptor: OleObjectDescriptor,
    pub entry_paths: Vec<PathBuf>,
}

/// Complete file root with the primary MS-DOC content structures linked into
/// a Rust tree. Unmanaged CFB entries remain available in `compound_file`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocFile {
    pub compound_file: CompoundFile,
    pub word_document: DocWordDocumentStream,
    pub table: DocTableStream,
    pub data: Option<DocDataStream>,
    pub object_pool: Option<DocObjectPoolStorage>,
}

impl DocFile {
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
        let compound = compound_from_path(path.as_ref(), options, BinaryFormat::Doc)?;
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
        let compound = compound_from_bytes(bytes, options, BinaryFormat::Doc)?;
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
        let compound = compound_outcome(compound_file, options, BinaryFormat::Doc)?;
        Self::from_compound_outcome(compound, options)
    }

    fn from_compound_outcome(
        compound: ParseOutcome<CompoundFile>,
        options: ParseOptions,
    ) -> Result<ParseOutcome<Self>> {
        let ParseOutcome {
            value: compound_file,
            mut diagnostics,
        } = compound;
        let limits = options.limits;
        let word_bytes = required_stream(&compound_file, WORD_DOCUMENT_STREAM)?.to_vec();
        ensure_stream_limit("WordDocument", &word_bytes, limits)?;
        let fib = Fib::from_word_document(&word_bytes)?;
        if fib
            .base
            .flags
            .intersects(FibBaseFlags::ENCRYPTED | FibBaseFlags::OBFUSCATED)
        {
            return Err(Error::invalid(
                0,
                "encrypted or obfuscated DOC requires an MS-OFFCRYPTO layer",
            ));
        }
        let table_name = if fib.base.flags.contains(FibBaseFlags::USE_1_TABLE) {
            DocTableStreamName::Table1
        } else {
            DocTableStreamName::Table0
        };
        let table_bytes = required_stream(&compound_file, table_name.path())?.to_vec();
        ensure_stream_limit(table_name.path(), &table_bytes, limits)?;

        let clx = parse_required(&table_bytes, fib.clx_location(), "CLX", Clx::from_bytes)?;
        let character_bin_table = parse_required(
            &table_bytes,
            fib.chpx_bte_location(),
            "PlcBteChpx",
            PlcBte::from_bytes,
        )?;
        let paragraph_bin_table = parse_required(
            &table_bytes,
            fib.papx_bte_location(),
            "PlcBtePapx",
            PlcBte::from_bytes,
        )?;
        let sections = parse_required(
            &table_bytes,
            fib.section_table_location(),
            "PlcfSed",
            PlcfSed::from_bytes,
        )?;
        let styles = parse_optional(
            &table_bytes,
            fib.style_sheet_location(),
            "STSH",
            StyleSheet::from_bytes,
        )?;
        let fonts = parse_optional(
            &table_bytes,
            fib.font_table_location(),
            "SttbfFfn",
            FontTable::from_bytes,
        )?;
        let mut fields = BTreeMap::new();
        for (part, location) in fib.field_table_locations() {
            if let Some(field) = parse_optional(&table_bytes, Some(location), "Plcfld", |bytes| {
                FieldTable::from_bytes_with_compatibility(bytes, !options.is_strict())
            })? {
                for position in field.value.separator_flag_mismatches() {
                    report_doc_compatibility(
                        &mut diagnostics,
                        ParseDiagnosticCode::NonconformingRecord,
                        u64::from(location.fc),
                        "Plcfld",
                        format!(
                            "field ending at document-part CP {position} has a separator that disagrees with grffldEnd.fHasSep"
                        ),
                    );
                }
                fields.insert(part, field);
            }
        }
        let bookmarks = parse_bookmarks(&fib, &table_bytes)?;
        let mut compatibility_tables = Vec::new();
        let header_text = parse_optional(
            &table_bytes,
            fib.header_text_location(),
            "PlcfHdd",
            HeaderTextTable::from_bytes,
        )?;
        let footnotes = parse_note_tables(&table_bytes, fib.footnote_locations(), "footnote")?;
        let endnotes = parse_note_tables(&table_bytes, fib.endnote_locations(), "endnote")?;
        let annotations = parse_annotation_tables(&table_bytes, fib.annotation_locations())?;
        let annotation_owners = parse_optional(
            &table_bytes,
            fib.annotation_owner_location(),
            "GrpXstAtnOwners",
            AnnotationOwners::from_bytes,
        )?;
        let annotation_bookmarks =
            parse_annotation_bookmarks(&fib, &table_bytes, options, &mut diagnostics)?;
        let annotation_extended_data = parse_optional_compatible(
            &table_bytes,
            fib.annotation_extended_data_location(),
            "AtrdExtra",
            AnnotationExtendedData::from_bytes,
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let textbox_stories = parse_part_tables(
            &table_bytes,
            fib.textbox_story_locations(),
            "PlcfTxbxTxt",
            TextboxStoryTable::from_bytes,
        )?;
        let textbox_breaks = parse_part_tables(
            &table_bytes,
            fib.textbox_break_locations(),
            "PlcfTxbxBkd",
            TextboxBreakTable::from_bytes,
        )?;
        let shape_anchors = parse_part_tables(
            &table_bytes,
            fib.shape_anchor_locations(),
            "PlcSpa",
            ShapeAnchorTable::from_bytes,
        )?;
        let office_art = parse_optional(
            &table_bytes,
            fib.office_art_content_location(),
            "OfficeArtContent",
            DocOfficeArtContent::from_bytes,
        )?;
        let revision_authors = parse_optional(
            &table_bytes,
            fib.revision_authors_location(),
            "SttbfRMark",
            RevisionAuthors::from_bytes,
        )?;
        let captions = parse_caption_tables(
            &table_bytes,
            fib.caption_locations(),
            options,
            &mut diagnostics,
        )?;
        let subdocuments = parse_optional_compatible(
            &table_bytes,
            fib.subdocuments_location(),
            "PlcfWkb",
            SubdocumentTable::from_bytes,
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let user_variables = parse_optional(
            &table_bytes,
            fib.user_variables_location(),
            "StwUser",
            UserVariables::from_bytes,
        )?;
        let embedded_fonts = parse_optional(
            &table_bytes,
            fib.embedded_fonts_location(),
            "SttbTtmbd",
            EmbeddedFontTable::from_bytes,
        )?;
        let spelling_state = parse_optional(
            &table_bytes,
            fib.spelling_state_location(),
            "PlcfSpl",
            SpellingStateTable::from_bytes,
        )?;
        let grammar_state = parse_optional(
            &table_bytes,
            fib.grammar_state_location(),
            "PlcfGram",
            GrammarStateTable::from_bytes,
        )?;
        let language_detection_state = parse_optional(
            &table_bytes,
            fib.language_detection_state_location(),
            "PlcfLad",
            LanguageDetectionStateTable::from_bytes,
        )?;
        let list_definitions = parse_list_definitions(
            &table_bytes,
            fib.list_definition_location(),
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let list_names = parse_optional(
            &table_bytes,
            fib.list_names_location(),
            "SttbListNames",
            ListNamesTable::from_bytes,
        )?;
        let list_overrides = parse_optional(
            &table_bytes,
            fib.list_override_location(),
            "PlfLfo",
            ListOverrides::from_bytes,
        )?;
        let document_properties = parse_optional(
            &table_bytes,
            fib.document_properties_location(),
            "Dop",
            DocumentProperties::from_bytes,
        )?;
        let associated_strings = parse_optional(
            &table_bytes,
            fib.associated_strings_location(),
            "SttbfAssoc",
            AssociatedStrings::from_bytes,
        )?;
        let external_file_names = parse_optional(
            &table_bytes,
            fib.external_file_names_location(),
            "SttbFnm",
            ExternalFileNameTable::from_bytes,
        )?;
        macro_rules! parse_compatible_table {
            ($location:expr, $label:literal, $parser:path) => {
                parse_optional_compatible(
                    &table_bytes,
                    $location,
                    $label,
                    $parser,
                    options,
                    &mut diagnostics,
                    &mut compatibility_tables,
                )?
            };
        }
        let mail_merge_state = parse_compatible_table!(
            fib.mail_merge_state_location(),
            "Pms",
            MailMergeState::from_bytes
        );
        let new_mail_merge_state = parse_compatible_table!(
            fib.new_mail_merge_state_location(),
            "PmsNew",
            MailMergeState::from_bytes
        );
        let office_data_source = parse_compatible_table!(
            fib.office_data_source_location(),
            "Odso",
            OfficeDataSource::from_bytes
        );
        let printer_driver_info = parse_compatible_table!(
            fib.printer_driver_info_location(),
            "PrDrvr",
            PrinterDriverInfo::from_bytes
        );
        let ole_control_infos = parse_compatible_table!(
            fib.ole_control_info_location(),
            "RgxOcxInfo",
            OleControlInfos::from_bytes
        );
        let table_character_cache = parse_compatible_table!(
            fib.table_character_cache_location(),
            "PlcfTch",
            TableCharacterCacheTable::from_bytes
        );
        let revision_message_threading = parse_compatible_table!(
            fib.revision_message_threading_location(),
            "RmdThreading",
            RevisionMessageThreading::from_bytes
        );
        let list_style_templates = parse_compatible_table!(
            fib.list_style_templates_location(),
            "SttbRgtplc",
            ListStyleTemplates::from_bytes
        );
        let frame_and_list_records = parse_compatible_table!(
            fib.frame_and_list_records_location(),
            "RgDofr",
            FrameAndListRecords::from_bytes
        );
        let grammar_option_sets = parse_compatible_table!(
            fib.grammar_option_sets_location(),
            "PlfCosi",
            GrammarOptionSets::from_bytes
        );
        let legacy_grammar_option_sets = parse_compatible_table!(
            fib.legacy_grammar_option_sets_location(),
            "PlfGosl",
            LegacyGrammarOptionSets::from_bytes
        );
        let auto_summary_ranges = parse_compatible_table!(
            fib.auto_summary_ranges_location(),
            "PlcfAsumy",
            AutoSummaryRangeTable::from_bytes
        );
        let smart_tag_recognizer_state = parse_compatible_table!(
            fib.smart_tag_recognizer_state_location(),
            "PlcfFactoid",
            SmartTagRecognizerStateTable::from_bytes
        );
        let xml_schema_references = parse_compatible_table!(
            fib.xml_schema_references_location(),
            "Hplxsdr",
            XmlSchemaReferences::from_bytes
        );
        let xml_transform_path = parse_compatible_table!(
            fib.xml_transform_path_location(),
            "CustomXForm",
            XmlTransformPath::from_bytes
        );
        let paragraph_group_properties = parse_compatible_table!(
            fib.paragraph_group_properties_location(),
            "PlcfPgp",
            ParagraphGroupProperties::from_bytes
        );
        let save_history = parse_compatible_table!(
            fib.save_history_location(),
            "SttbSavedBy",
            SaveHistory::from_bytes
        );
        let grammar_checker_cookies = parse_compatible_table!(
            fib.grammar_checker_cookies_location(),
            "PlcfCookie",
            GrammarCheckerCookieTable::from_bytes
        );
        let legacy_grammar_checker_cookies = parse_compatible_table!(
            fib.legacy_grammar_checker_cookies_location(),
            "PlcfCookieOld",
            LegacyGrammarCheckerCookieTable::from_bytes
        );
        let grammar_cookie_data = parse_compatible_table!(
            fib.grammar_cookie_data_location(),
            "CookieData",
            GrammarCookieStore::from_bytes
        );
        let smart_tag_data = parse_compatible_table!(
            fib.smart_tag_data_location(),
            "FactoidData",
            SmartTagData::from_bytes
        );
        let revision_save_ids = parse_compatible_table!(
            fib.revision_save_ids_location(),
            "Plrsid",
            RevisionSaveIdTable::from_bytes
        );
        let selection_state = parse_compatible_table!(
            fib.selection_state_location(),
            "Wss",
            SelectionState::from_bytes
        );
        let command_customizations = parse_compatible_table!(
            fib.command_customizations_location(),
            "Cmds",
            CommandCustomizations::from_bytes
        );
        let structured_tag_bookmarks = parse_bookmark_set(
            &table_bytes,
            fib.structured_tag_bookmark_locations(),
            ["SttbfBkmkSdt", "PlcfBkfSdt", "PlcfBklSdt"],
            "structured-tag bookmarks",
            StructuredTagBookmarks::from_bytes,
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let range_protection = parse_range_protection(
            &table_bytes,
            fib.range_protection_locations(),
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let smart_tag_bookmarks = parse_bookmark_set(
            &table_bytes,
            fib.smart_tag_bookmark_locations(),
            ["SttbfBkmkFactoid", "PlcfBkfFactoid", "PlcfBklFactoid"],
            "smart-tag bookmarks",
            SmartTagBookmarks::from_bytes,
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let format_consistency_bookmarks = parse_bookmark_set(
            &table_bytes,
            fib.format_consistency_bookmark_locations(),
            ["SttbfBkmkFcc", "PlcfBkfFcc", "PlcfBklFcc"],
            "format-consistency bookmarks",
            FormatConsistencyBookmarks::from_bytes,
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let repair_bookmarks = parse_bookmark_set(
            &table_bytes,
            fib.repair_bookmark_locations(),
            ["SttbfBkmkBpRepairs", "PlcfBkfBpRepairs", "PlcfBklBpRepairs"],
            "repair bookmarks",
            RepairBookmarks::from_bytes,
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let user_input_methods = parse_user_input_methods(
            &table_bytes,
            fib.user_input_method_locations(),
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let mso_envelope = parse_optional_compatible(
            &table_bytes,
            fib.mso_envelope_location(),
            "MsoEnvelope",
            MsoEnvelope::from_bytes,
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?;
        let deprecated_numbering_field_cache = parse_optional_compatible(
            &table_bytes,
            fib.deprecated_numbering_field_cache_location(),
            "PlcfBteLvc",
            |bytes| Ok::<_, Error>(bytes.to_vec()),
            options,
            &mut diagnostics,
            &mut compatibility_tables,
        )?
        .map(|value| DocDeprecatedNumberingFieldCache {
            location: value.location,
            physical_bytes: value.value,
        });

        let text_pieces = parse_text_pieces(&clx.value, &word_bytes, limits)?;
        let character_format_pages =
            parse_fkp_pages(&character_bin_table.value, &word_bytes, ChpxFkp::from_bytes)?;
        let paragraph_format_pages =
            parse_fkp_pages(&paragraph_bin_table.value, &word_bytes, PapxFkp::from_bytes)?;
        validate_fkp_page_order(
            &character_format_pages,
            |page| &page.file_positions,
            "ChpxFkp rgfc",
            options,
            &mut diagnostics,
        )?;
        validate_fkp_page_order(
            &paragraph_format_pages,
            |page| &page.file_positions,
            "PapxFkp rgfc",
            options,
            &mut diagnostics,
        )?;
        let chpx_runs = match source_character_formatting_runs(&character_format_pages, &clx.value)
        {
            Ok(runs) => Some(runs),
            Err(error) if !options.is_strict() => {
                report_doc_compatibility(
                    &mut diagnostics,
                    ParseDiagnosticCode::NonconformingRecord,
                    u64::from(character_bin_table.location.fc),
                    "PlcBteChpx",
                    format!("CHPX FKP runs cannot form a CP tree: {error}"),
                );
                None
            }
            Err(error) => return Err(error),
        };
        let papx_runs = match source_paragraph_formatting_runs(
            &paragraph_format_pages,
            &clx.value,
            &word_bytes,
        ) {
            Ok(runs) => Some(runs),
            Err(error) if !options.is_strict() => {
                report_doc_compatibility(
                    &mut diagnostics,
                    ParseDiagnosticCode::NonconformingRecord,
                    u64::from(paragraph_bin_table.location.fc),
                    "PlcBtePapx",
                    format!("PAPX FKP runs cannot form a CP tree: {error}"),
                );
                None
            }
            Err(error) => return Err(error),
        };
        let section_properties = sections
            .value
            .sections
            .iter()
            .enumerate()
            .map(|(section_index, sed)| {
                let value = Sepx::from_word_document(&word_bytes, sed.sepx_offset)?;
                let physical_len = value
                    .as_ref()
                    .map(Sepx::to_bytes)
                    .transpose()?
                    .map_or(0, |bytes| bytes.len());
                Ok(DocSectionProperties {
                    section_index,
                    offset: sed.sepx_offset,
                    physical_len,
                    value,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        ensure_entry_limit("DOC text pieces", text_pieces.len(), limits)?;
        ensure_entry_limit("DOC fields", fields.len(), limits)?;

        let source_fib_len = fib.encoded_len();
        let source_piece_indices = (0..text_pieces.len()).collect();
        let source_chpx_runs = chpx_runs.clone();
        let source_papx_runs = papx_runs.clone();
        let word_document = DocWordDocumentStream {
            fib,
            text_pieces,
            chpx_runs,
            papx_runs,
            character_format_pages,
            paragraph_format_pages,
            section_properties,
            physical_bytes: word_bytes,
            source_fib_len,
            source_piece_indices,
            source_chpx_runs,
            source_papx_runs,
            pending_text_edits: BTreeMap::new(),
            rebuild_character_formatting: false,
            rebuild_paragraph_formatting: false,
        };
        let table = DocTableStream {
            name: table_name,
            clx,
            character_bin_table,
            paragraph_bin_table,
            sections,
            styles,
            fonts,
            fields,
            bookmarks,
            header_text,
            footnotes,
            endnotes,
            annotations,
            annotation_owners,
            annotation_bookmarks,
            annotation_extended_data,
            textbox_stories,
            textbox_breaks,
            shape_anchors,
            office_art,
            revision_authors,
            captions,
            subdocuments,
            user_variables,
            embedded_fonts,
            spelling_state,
            grammar_state,
            language_detection_state,
            list_definitions,
            list_names,
            list_overrides,
            document_properties,
            associated_strings,
            external_file_names,
            mail_merge_state,
            new_mail_merge_state,
            office_data_source,
            printer_driver_info,
            ole_control_infos,
            table_character_cache,
            revision_message_threading,
            list_style_templates,
            frame_and_list_records,
            grammar_option_sets,
            legacy_grammar_option_sets,
            auto_summary_ranges,
            smart_tag_recognizer_state,
            xml_schema_references,
            xml_transform_path,
            paragraph_group_properties,
            save_history,
            grammar_checker_cookies,
            legacy_grammar_checker_cookies,
            grammar_cookie_data,
            smart_tag_data,
            revision_save_ids,
            selection_state,
            command_customizations,
            structured_tag_bookmarks,
            range_protection,
            smart_tag_bookmarks,
            format_consistency_bookmarks,
            repair_bookmarks,
            user_input_methods,
            mso_envelope,
            deprecated_numbering_field_cache,
            compatibility_tables,
            physical_bytes: table_bytes,
        };
        let data = parse_data_stream(
            compound_file.stream(DATA_STREAM),
            &word_document,
            &table,
            options,
            &mut diagnostics,
        )?;
        let object_pool = parse_object_pool(&compound_file, options, &mut diagnostics)?;
        Ok(ParseOutcome::new(
            Self {
                compound_file,
                word_document,
                table,
                data,
                object_pool,
            },
            diagnostics,
        ))
    }

    pub fn to_compound_file(&self) -> Result<CompoundFile> {
        self.validate_links()?;
        let mut word = self.word_document.physical_bytes.clone();
        let mut table = TableLayout::new(self.table.physical_bytes.clone());
        let mut fib = self.word_document.fib.clone();
        let mut clx_table = self.table.clx.clone();
        let mut character_bin_table = self.table.character_bin_table.clone();
        let mut paragraph_bin_table = self.table.paragraph_bin_table.clone();
        let mut character_format_pages = self.word_document.character_format_pages.clone();
        let mut paragraph_format_pages = self.word_document.paragraph_format_pages.clone();
        let mut section_properties = self.word_document.section_properties.clone();
        let mut sections_table = self.table.sections.clone();
        let mut styles_table = self.table.styles.clone();
        let mut list_definitions = self.table.list_definitions.clone();
        let RebuiltDataStream {
            bytes: data_bytes,
            relocations: data_relocations,
        } = rebuild_data_stream(self.data.as_ref())?;
        relocate_root_data_references(
            &mut clx_table,
            &mut character_format_pages,
            &mut paragraph_format_pages,
            &mut section_properties,
            styles_table.as_mut(),
            list_definitions.as_mut(),
            &data_relocations,
        )?;
        if let Some(styles) = &styles_table {
            patch_located(&mut table, styles, StyleSheet::to_bytes, "STSH")?;
        }
        if let Some(fonts) = &self.table.fonts {
            patch_located(&mut table, fonts, FontTable::to_bytes, "SttbfFfn")?;
        }
        for field in self.table.fields.values() {
            patch_located(&mut table, field, FieldTable::to_bytes, "Plcfld")?;
        }
        if let Some(bookmarks) = &self.table.bookmarks {
            let (names, starts, ends) = bookmarks.value.to_bytes()?;
            patch_location(&mut table, bookmarks.names_location, names, "SttbfBkmk")?;
            patch_location(&mut table, bookmarks.starts_location, starts, "PlcfBkf")?;
            patch_location(&mut table, bookmarks.ends_location, ends, "PlcfBkl")?;
        }
        patch_optional_located(
            &mut table,
            self.table.header_text.as_ref(),
            HeaderTextTable::to_bytes,
            "PlcfHdd",
        )?;
        if let Some(notes) = &self.table.footnotes {
            patch_located(
                &mut table,
                &notes.references,
                NoteReferenceTable::to_bytes,
                "PlcfFndRef",
            )?;
            patch_located(&mut table, &notes.text, CpOnlyTable::to_bytes, "PlcfFndTxt")?;
        }
        if let Some(notes) = &self.table.endnotes {
            patch_located(
                &mut table,
                &notes.references,
                NoteReferenceTable::to_bytes,
                "PlcfEndRef",
            )?;
            patch_located(&mut table, &notes.text, CpOnlyTable::to_bytes, "PlcfEndTxt")?;
        }
        if let Some(annotations) = &self.table.annotations {
            patch_located(
                &mut table,
                &annotations.references,
                AnnotationReferenceTable::to_bytes,
                "PlcfAndRef",
            )?;
            patch_located(
                &mut table,
                &annotations.text,
                CpOnlyTable::to_bytes,
                "PlcfAndTxt",
            )?;
        }
        patch_optional_located(
            &mut table,
            self.table.annotation_owners.as_ref(),
            AnnotationOwners::to_bytes,
            "GrpXstAtnOwners",
        )?;
        if let Some(bookmarks) = &self.table.annotation_bookmarks {
            let (infos, starts, ends) = bookmarks.value.to_bytes()?;
            patch_location(&mut table, bookmarks.infos_location, infos, "SttbfAtnBkmk")?;
            patch_location(&mut table, bookmarks.starts_location, starts, "PlcfAtnBkf")?;
            patch_location(&mut table, bookmarks.ends_location, ends, "PlcfAtnBkl")?;
        }
        patch_optional_located(
            &mut table,
            self.table.annotation_extended_data.as_ref(),
            AnnotationExtendedData::to_bytes,
            "AtrdExtra",
        )?;
        patch_part_tables(
            &mut table,
            &self.table.textbox_stories,
            TextboxStoryTable::to_bytes,
            "PlcfTxbxTxt",
        )?;
        patch_part_tables(
            &mut table,
            &self.table.textbox_breaks,
            TextboxBreakTable::to_bytes,
            "PlcfTxbxBkd",
        )?;
        patch_part_tables(
            &mut table,
            &self.table.shape_anchors,
            ShapeAnchorTable::to_bytes,
            "PlcSpa",
        )?;
        if let Some(office_art) = &self.table.office_art {
            patch_located(
                &mut table,
                office_art,
                DocOfficeArtContent::to_bytes,
                "OfficeArtContent",
            )?;
        }
        patch_optional_located(
            &mut table,
            self.table.revision_authors.as_ref(),
            RevisionAuthors::to_bytes,
            "SttbfRMark",
        )?;
        if let Some(captions) = &self.table.captions {
            patch_located(
                &mut table,
                &captions.definitions,
                CaptionDefinitions::to_bytes,
                "SttbfCaption",
            )?;
            patch_located(
                &mut table,
                &captions.automatic,
                AutoCaptionDefinitions::to_bytes,
                "SttbfAutoCaption",
            )?;
        }
        patch_optional_located(
            &mut table,
            self.table.subdocuments.as_ref(),
            SubdocumentTable::to_bytes,
            "PlcfWkb",
        )?;
        patch_optional_located(
            &mut table,
            self.table.user_variables.as_ref(),
            UserVariables::to_bytes,
            "StwUser",
        )?;
        patch_optional_located(
            &mut table,
            self.table.embedded_fonts.as_ref(),
            EmbeddedFontTable::to_bytes,
            "SttbTtmbd",
        )?;
        patch_optional_located(
            &mut table,
            self.table.spelling_state.as_ref(),
            SpellingStateTable::to_bytes,
            "PlcfSpl",
        )?;
        patch_optional_located(
            &mut table,
            self.table.grammar_state.as_ref(),
            GrammarStateTable::to_bytes,
            "PlcfGram",
        )?;
        patch_optional_located(
            &mut table,
            self.table.language_detection_state.as_ref(),
            LanguageDetectionStateTable::to_bytes,
            "PlcfLad",
        )?;
        if let Some(definitions) = &list_definitions {
            let (base, levels) = definitions.value.to_bytes()?;
            patch_location(&mut table, definitions.location, base, "PlfLst")?;
            let levels_offset = usize::try_from(definitions.location.fc)
                .ok()
                .and_then(|offset| {
                    usize::try_from(definitions.location.lcb)
                        .ok()
                        .and_then(|length| offset.checked_add(length))
                })
                .ok_or_else(|| Error::Limit("PlfLst levels offset exceeds usize".into()))?;
            patch_at(
                &mut table,
                levels_offset,
                definitions.trailing_levels_len,
                levels,
                "PlfLst levels",
            )?;
        }
        patch_optional_located(
            &mut table,
            self.table.list_names.as_ref(),
            ListNamesTable::to_bytes,
            "SttbListNames",
        )?;
        patch_optional_located(
            &mut table,
            self.table.list_overrides.as_ref(),
            ListOverrides::to_bytes,
            "PlfLfo",
        )?;
        patch_optional_located(
            &mut table,
            self.table.document_properties.as_ref(),
            DocumentProperties::to_bytes,
            "Dop",
        )?;
        patch_optional_located(
            &mut table,
            self.table.associated_strings.as_ref(),
            AssociatedStrings::to_bytes,
            "SttbfAssoc",
        )?;
        patch_optional_located(
            &mut table,
            self.table.external_file_names.as_ref(),
            ExternalFileNameTable::to_bytes,
            "SttbFnm",
        )?;
        macro_rules! patch_table {
            ($field:ident, $type:ty, $label:literal) => {
                patch_optional_located(
                    &mut table,
                    self.table.$field.as_ref(),
                    <$type>::to_bytes,
                    $label,
                )?;
            };
        }
        patch_table!(mail_merge_state, MailMergeState, "Pms");
        patch_table!(new_mail_merge_state, MailMergeState, "PmsNew");
        patch_table!(office_data_source, OfficeDataSource, "Odso");
        patch_table!(printer_driver_info, PrinterDriverInfo, "PrDrvr");
        patch_table!(ole_control_infos, OleControlInfos, "RgxOcxInfo");
        patch_table!(table_character_cache, TableCharacterCacheTable, "PlcfTch");
        patch_table!(
            revision_message_threading,
            RevisionMessageThreading,
            "RmdThreading"
        );
        patch_table!(list_style_templates, ListStyleTemplates, "SttbRgtplc");
        patch_table!(frame_and_list_records, FrameAndListRecords, "RgDofr");
        patch_table!(grammar_option_sets, GrammarOptionSets, "PlfCosi");
        patch_table!(
            legacy_grammar_option_sets,
            LegacyGrammarOptionSets,
            "PlfGosl"
        );
        patch_table!(auto_summary_ranges, AutoSummaryRangeTable, "PlcfAsumy");
        patch_table!(
            smart_tag_recognizer_state,
            SmartTagRecognizerStateTable,
            "PlcfFactoid"
        );
        patch_table!(xml_schema_references, XmlSchemaReferences, "Hplxsdr");
        patch_table!(xml_transform_path, XmlTransformPath, "CustomXForm");
        patch_table!(
            paragraph_group_properties,
            ParagraphGroupProperties,
            "PlcfPgp"
        );
        patch_table!(save_history, SaveHistory, "SttbSavedBy");
        patch_table!(
            grammar_checker_cookies,
            GrammarCheckerCookieTable,
            "PlcfCookie"
        );
        patch_table!(
            legacy_grammar_checker_cookies,
            LegacyGrammarCheckerCookieTable,
            "PlcfCookieOld"
        );
        patch_table!(grammar_cookie_data, GrammarCookieStore, "CookieData");
        patch_table!(smart_tag_data, SmartTagData, "FactoidData");
        patch_table!(revision_save_ids, RevisionSaveIdTable, "Plrsid");
        patch_table!(selection_state, SelectionState, "Wss");
        patch_table!(command_customizations, CommandCustomizations, "Cmds");
        if let Some(value) = &self.table.structured_tag_bookmarks {
            let encoded = value.value.to_bytes()?;
            patch_location(
                &mut table,
                value.metadata_location,
                encoded.tags,
                "SttbfBkmkSdt",
            )?;
            patch_location(
                &mut table,
                value.starts_location,
                encoded.starts,
                "PlcfBkfSdt",
            )?;
            patch_location(&mut table, value.ends_location, encoded.ends, "PlcfBklSdt")?;
        }
        if let Some(value) = &self.table.range_protection {
            let encoded = value.value.to_bytes()?;
            patch_location(
                &mut table,
                value.permissions_location,
                encoded.permissions,
                "SttbfBkmkProt",
            )?;
            patch_location(
                &mut table,
                value.starts_location,
                encoded.starts,
                "PlcfBkfProt",
            )?;
            patch_location(&mut table, value.ends_location, encoded.ends, "PlcfBklProt")?;
            patch_location(
                &mut table,
                value.users_location,
                encoded.users,
                "SttbProtUser",
            )?;
        }
        if let Some(value) = &self.table.smart_tag_bookmarks {
            let (metadata, starts, ends) = value.value.to_bytes()?;
            patch_location(
                &mut table,
                value.metadata_location,
                metadata,
                "SttbfBkmkFactoid",
            )?;
            patch_location(&mut table, value.starts_location, starts, "PlcfBkfFactoid")?;
            patch_location(&mut table, value.ends_location, ends, "PlcfBklFactoid")?;
        }
        if let Some(value) = &self.table.format_consistency_bookmarks {
            let encoded = value.value.to_bytes()?;
            patch_location(
                &mut table,
                value.metadata_location,
                encoded.metadata,
                "SttbfBkmkFcc",
            )?;
            patch_location(
                &mut table,
                value.starts_location,
                encoded.starts,
                "PlcfBkfFcc",
            )?;
            patch_location(&mut table, value.ends_location, encoded.ends, "PlcfBklFcc")?;
        }
        if let Some(value) = &self.table.repair_bookmarks {
            let encoded = value.value.to_bytes()?;
            patch_location(
                &mut table,
                value.metadata_location,
                encoded.metadata,
                "SttbfBkmkBpRepairs",
            )?;
            patch_location(
                &mut table,
                value.starts_location,
                encoded.starts,
                "PlcfBkfBpRepairs",
            )?;
            patch_location(
                &mut table,
                value.ends_location,
                encoded.ends,
                "PlcfBklBpRepairs",
            )?;
        }
        if let Some(value) = &self.table.user_input_methods {
            let (methods, guids) = value.value.to_bytes()?;
            patch_location(&mut table, value.methods_location, methods, "PlcfUim")?;
            patch_location(
                &mut table,
                value.service_guids_location,
                guids,
                "PlfGuidUim",
            )?;
        }
        patch_optional_located(
            &mut table,
            self.table.mso_envelope.as_ref(),
            MsoEnvelope::to_bytes,
            "MsoEnvelope",
        )?;
        if let Some(value) = &self.table.deprecated_numbering_field_cache {
            patch_location(
                &mut table,
                value.location,
                value.physical_bytes.clone(),
                "PlcfBteLvc",
            )?;
        }

        let source_clx = Clx::from_bytes(bounded_slice(
            &self.table.physical_bytes,
            self.table.clx.location,
            "source CLX",
        )?)?;
        if self.word_document.source_piece_indices.len() != self.word_document.text_pieces.len() {
            return Err(Error::invalid(
                0,
                "DOC current/source text-piece identity cardinality changed",
            ));
        }
        if self.word_document.chpx_runs.is_none() && self.word_document.source_chpx_runs.is_some() {
            return Err(Error::invalid(0, "DOC CHPX CP tree was removed"));
        }
        if self.word_document.papx_runs.is_none() && self.word_document.source_papx_runs.is_some() {
            return Err(Error::invalid(0, "DOC PAPX CP tree was removed"));
        }
        let rebuild_character_formatting = self.word_document.rebuild_character_formatting
            || self.word_document.chpx_runs != self.word_document.source_chpx_runs;
        let rebuild_paragraph_formatting = self.word_document.rebuild_paragraph_formatting
            || self.word_document.papx_runs != self.word_document.source_papx_runs;
        let mut encoded_pieces = Vec::with_capacity(self.word_document.text_pieces.len());
        let mut text_layout_changed = false;
        for (current_piece_index, piece) in self.word_document.text_pieces.iter().enumerate() {
            let source_piece_index = self.word_document.source_piece_indices[current_piece_index];
            let bytes = piece.value.to_bytes();
            let expected_characters = piece
                .value
                .cp_end
                .checked_sub(piece.value.cp_start)
                .and_then(|count| usize::try_from(count).ok())
                .ok_or_else(|| {
                    Error::invalid(
                        u64::from(piece.value.file_offset),
                        "text piece CP range is invalid",
                    )
                })?;
            if piece.value.character_count() != expected_characters {
                return Err(Error::invalid(
                    u64::from(piece.value.file_offset),
                    "text piece character count changed",
                ));
            }
            let descriptor = source_clx
                .piece_table
                .pieces
                .get(source_piece_index)
                .ok_or_else(|| Error::invalid(0, "source text piece index is stale"))?;
            let source_cp_start = *source_clx
                .piece_table
                .character_positions
                .get(source_piece_index)
                .ok_or_else(|| Error::invalid(0, "source text piece CP start is missing"))?;
            let source_cp_end = *source_clx
                .piece_table
                .character_positions
                .get(source_piece_index + 1)
                .ok_or_else(|| Error::invalid(0, "source text piece CP limit is missing"))?;
            let source_characters = descriptor.text_piece(
                &self.word_document.physical_bytes,
                source_cp_start,
                source_cp_end,
            )?;
            let source_character_count = source_characters.character_count();
            let character_replacements = if let Some(edits) = self
                .word_document
                .pending_text_edits
                .get(&source_piece_index)
            {
                let relocated_count = relocate_character_position(
                    u32::try_from(source_character_count).map_err(|_| {
                        Error::Limit("source text piece character count exceeds u32".into())
                    })?,
                    edits,
                    "text piece character count",
                )?;
                if usize::try_from(relocated_count).ok() != Some(expected_characters) {
                    return Err(Error::invalid(
                        u64::from(piece.value.file_offset),
                        "text piece has an untracked variable-length edit after replace_text_range",
                    ));
                }
                edits.clone()
            } else {
                vec![text_piece_character_replacement(
                    &source_characters.characters,
                    &piece.value.characters,
                )?]
            };
            let source_width = if descriptor.file_position.compressed {
                1usize
            } else {
                2usize
            };
            let source_len = source_character_count
                .checked_mul(source_width)
                .ok_or_else(|| Error::Limit("text piece byte length overflow".into()))?;
            text_layout_changed |= bytes.len() != source_len;
            encoded_pieces.push(EncodedTextPiece {
                piece_index: current_piece_index,
                source_offset: descriptor.file_position.byte_offset(),
                source_len,
                source_width,
                source_character_count,
                destination_character_count: expected_characters,
                destination_start: None,
                character_replacements,
                compressed: matches!(&piece.value.characters, TextPieceCharacters::Compressed(_)),
                bytes,
            });
        }
        text_layout_changed |= rebuild_character_formatting || rebuild_paragraph_formatting;

        if text_layout_changed {
            let meaningful_end = usize::try_from(fib.rg_lw.cb_mac)
                .map_err(|_| Error::Limit("FIB cbMac exceeds usize".into()))?;
            if meaningful_end > word.len() {
                return Err(Error::invalid(
                    u64::from(fib.rg_lw.cb_mac),
                    "FIB cbMac exceeds WordDocument",
                ));
            }
            let mut source_order = encoded_pieces.iter().collect::<Vec<_>>();
            source_order.sort_by_key(|piece| piece.source_offset);
            for pair in source_order.windows(2) {
                let left_end = u64::from(pair[0].source_offset)
                    .checked_add(pair[0].source_len as u64)
                    .ok_or_else(|| Error::Limit("text piece source range overflow".into()))?;
                if left_end > u64::from(pair[1].source_offset) {
                    return Err(Error::invalid(
                        u64::from(pair[1].source_offset),
                        "overlapping text pieces cannot be relocated",
                    ));
                }
            }

            let appended_len = encoded_pieces.iter().try_fold(0usize, |length, piece| {
                length
                    .checked_add(piece.bytes.len())
                    .ok_or_else(|| Error::Limit("relocated text length overflow".into()))
            })?;
            let mut appended = Vec::with_capacity(appended_len);
            let mut relocations = Vec::with_capacity(encoded_pieces.len());
            for piece in &mut encoded_pieces {
                let new_offset = meaningful_end
                    .checked_add(appended.len())
                    .ok_or_else(|| Error::Limit("relocated text offset overflow".into()))?;
                let new_offset_u32 = u32::try_from(new_offset)
                    .map_err(|_| Error::Limit("relocated text offset exceeds u32".into()))?;
                let new_width = if piece.compressed { 1usize } else { 2usize };
                let descriptor = clx_table
                    .value
                    .piece_table
                    .pieces
                    .get_mut(piece.piece_index)
                    .ok_or_else(|| Error::invalid(0, "text piece index is stale"))?;
                descriptor.file_position.fc = if piece.compressed {
                    new_offset_u32.checked_mul(2).ok_or_else(|| {
                        Error::Limit("compressed text FC representation overflow".into())
                    })?
                } else {
                    new_offset_u32
                };
                if descriptor.file_position.fc > 0x3fff_ffff {
                    return Err(Error::Limit("relocated text FC exceeds 30 bits".into()));
                }
                descriptor.file_position.compressed = piece.compressed;
                piece.destination_start = Some(new_offset_u32);
                relocations.push(TextRelocation {
                    source_start: piece.source_offset,
                    source_len: piece.source_len,
                    source_width: piece.source_width,
                    destination_start: new_offset_u32,
                    destination_width: new_width,
                    source_character_count: piece.source_character_count,
                    destination_character_count: piece.destination_character_count,
                    character_replacements: piece.character_replacements.clone(),
                });
                appended.extend_from_slice(&piece.bytes);
            }
            if !rebuild_paragraph_formatting {
                relocate_text_file_positions(
                    &mut paragraph_bin_table.value.file_positions,
                    &relocations,
                )?;
                for page in &mut paragraph_format_pages {
                    relocate_text_file_positions(&mut page.value.file_positions, &relocations)?;
                }
            }
            word.splice(meaningful_end..meaningful_end, appended);
            fib.rg_lw.cb_mac = fib
                .rg_lw
                .cb_mac
                .checked_add(
                    u32::try_from(appended_len)
                        .map_err(|_| Error::Limit("relocated text length exceeds u32".into()))?,
                )
                .ok_or_else(|| Error::Limit("FIB cbMac overflow".into()))?;
            if rebuild_character_formatting {
                let rebuilt = rebuild_character_formatting_pages(
                    self.word_document.chpx_runs.as_deref().ok_or_else(|| {
                        Error::invalid(0, "DOC CHPX CP tree is unavailable for rebuild")
                    })?,
                    &self.word_document.text_pieces,
                    &encoded_pieces,
                    &mut word,
                    &mut fib,
                )?;
                character_bin_table.value = rebuilt.0;
                character_format_pages = rebuilt.1;
            } else {
                relocate_text_file_positions(
                    &mut character_bin_table.value.file_positions,
                    &relocations,
                )?;
                for page in &mut character_format_pages {
                    relocate_text_file_positions(&mut page.value.file_positions, &relocations)?;
                }
            }
            if rebuild_paragraph_formatting {
                let rebuilt = rebuild_paragraph_formatting_pages(
                    self.word_document.papx_runs.as_deref().ok_or_else(|| {
                        Error::invalid(0, "DOC PAPX CP tree is unavailable for rebuild")
                    })?,
                    &self.word_document.text_pieces,
                    &encoded_pieces,
                    &mut word,
                    &mut fib,
                )?;
                paragraph_bin_table.value = rebuilt.0;
                paragraph_format_pages = rebuilt.1;
            }
        } else {
            let relocations = encoded_pieces
                .iter()
                .map(|piece| TextRelocation {
                    source_start: piece.source_offset,
                    source_len: piece.source_len,
                    source_width: piece.source_width,
                    destination_start: piece.source_offset,
                    destination_width: piece.source_width,
                    source_character_count: piece.source_character_count,
                    destination_character_count: piece.destination_character_count,
                    character_replacements: piece.character_replacements.clone(),
                })
                .collect::<Vec<_>>();
            relocate_text_file_positions(
                &mut character_bin_table.value.file_positions,
                &relocations,
            )?;
            relocate_text_file_positions(
                &mut paragraph_bin_table.value.file_positions,
                &relocations,
            )?;
            for page in &mut character_format_pages {
                relocate_text_file_positions(&mut page.value.file_positions, &relocations)?;
            }
            for page in &mut paragraph_format_pages {
                relocate_text_file_positions(&mut page.value.file_positions, &relocations)?;
            }
            for piece in encoded_pieces {
                patch_at(
                    &mut word,
                    usize::try_from(piece.source_offset)
                        .map_err(|_| Error::Limit("text piece offset exceeds usize".into()))?,
                    piece.source_len,
                    piece.bytes,
                    "text piece",
                )?;
            }
        }
        for page in &character_format_pages {
            patch_at(
                &mut word,
                page.page.byte_offset()?,
                512,
                page.value.to_bytes()?,
                "ChpxFkp",
            )?;
        }
        for page in &paragraph_format_pages {
            patch_at(
                &mut word,
                page.page.byte_offset()?,
                512,
                page.value.to_bytes()?,
                "PapxFkp",
            )?;
        }
        for section in &section_properties {
            if let Some(value) = &section.value {
                let encoded = value.to_bytes()?;
                if encoded.len() == section.physical_len {
                    patch_at(
                        &mut word,
                        usize::try_from(section.offset)
                            .map_err(|_| Error::invalid(0, "negative Sepx offset"))?,
                        section.physical_len,
                        encoded,
                        "Sepx",
                    )?;
                } else {
                    let meaningful_end = usize::try_from(fib.rg_lw.cb_mac)
                        .map_err(|_| Error::Limit("FIB cbMac exceeds usize".into()))?;
                    if meaningful_end > word.len() {
                        return Err(Error::invalid(
                            u64::from(fib.rg_lw.cb_mac),
                            "FIB cbMac exceeds WordDocument",
                        ));
                    }
                    let new_offset = i32::try_from(meaningful_end)
                        .map_err(|_| Error::Limit("Sepx offset exceeds i32".into()))?;
                    let encoded_len = u32::try_from(encoded.len())
                        .map_err(|_| Error::Limit("Sepx length exceeds u32".into()))?;
                    word.splice(meaningful_end..meaningful_end, encoded);
                    fib.rg_lw.cb_mac = fib
                        .rg_lw
                        .cb_mac
                        .checked_add(encoded_len)
                        .ok_or_else(|| Error::Limit("FIB cbMac overflow".into()))?;
                    sections_table
                        .value
                        .sections
                        .get_mut(section.section_index)
                        .ok_or_else(|| Error::invalid(0, "section property index is stale"))?
                        .sepx_offset = new_offset;
                }
            }
        }

        patch_located(&mut table, &clx_table, Clx::to_bytes, "CLX")?;
        patch_located(
            &mut table,
            &character_bin_table,
            PlcBte::to_bytes,
            "PlcBteChpx",
        )?;
        patch_located(
            &mut table,
            &paragraph_bin_table,
            PlcBte::to_bytes,
            "PlcBtePapx",
        )?;
        patch_located(&mut table, &sections_table, PlcfSed::to_bytes, "PlcfSed")?;
        let (table, relocation) = table.finish()?;
        fib.relocate_table_locations(|location| relocation.relocate(location))?;
        patch_prefix(
            &mut word,
            self.word_document.source_fib_len,
            fib.to_bytes()?,
            "FIB",
        )?;

        let mut compound = self.compound_file.clone();
        compound.replace_stream(WORD_DOCUMENT_STREAM, word)?;
        compound.replace_stream(self.table.name.path(), table)?;
        match data_bytes {
            Some(bytes) => {
                compound.create_or_replace_stream(DATA_STREAM, bytes)?;
            }
            None if compound.is_stream(DATA_STREAM) => {
                compound.remove_stream(DATA_STREAM)?;
            }
            None => {}
        }
        if let Some(object_pool) = &self.object_pool {
            for object in &object_pool.objects {
                compound
                    .replace_stream(&object.descriptor_stream_path, object.descriptor.to_bytes())?;
            }
        }
        Ok(compound)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_compound_file()?.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.to_compound_file()?.save(path)
    }

    /// Replaces a character range in the Main Document and relocates every
    /// currently materialized CP reference whose coordinate space is affected.
    ///
    /// The range is relative to the Main Document, as specified by MS-DOC
    /// section 2.3.1. The replacement must use the physical encoding of the
    /// first affected text piece. CHPX and PAPX boundaries are rebuilt from
    /// logical runs; the paragraph/cell/section terminator sequence must stay
    /// unchanged so paragraph formatting inheritance is unambiguous. A
    /// cross-piece edit removes any emptied PlcPcd descriptors and rebuilds
    /// the logical formatting tables before serialization.
    pub fn replace_main_text_range(
        &mut self,
        range: Range<u32>,
        replacement: TextPieceCharacters,
    ) -> Result<()> {
        self.replace_text_range(FieldDocumentPart::Main, range, replacement)
    }

    /// Replaces a character range whose CPs are relative to an MS-DOC
    /// document part and relocates both global and part-local CP references.
    pub fn replace_text_range(
        &mut self,
        part: FieldDocumentPart,
        range: Range<u32>,
        replacement: TextPieceCharacters,
    ) -> Result<()> {
        let mut edited = self.clone();
        edited.replace_text_range_composed(
            part,
            range,
            replacement,
            ParagraphMarkEdit::PreserveAll,
        )?;
        edited.realign_papx_runs()?;
        *self = edited;
        Ok(())
    }

    /// Replaces Main Document text while allowing paragraph-mark (0x000D)
    /// insertion or deletion. `papx_runs` is the complete PAPX FKP CP tree in
    /// the post-edit global coordinate space; no formatting inheritance is
    /// inferred. Cell marks (0x0007) and section marks (0x000C) must remain
    /// unchanged because their other owning structures are not supplied here.
    pub fn replace_main_text_range_with_papx_runs(
        &mut self,
        range: Range<u32>,
        replacement: TextPieceCharacters,
        papx_runs: Vec<DocPapxRun>,
    ) -> Result<()> {
        self.replace_text_range_with_papx_runs(
            FieldDocumentPart::Main,
            range,
            replacement,
            papx_runs,
        )
    }

    /// Replaces text in an MS-DOC document part while allowing paragraph-mark
    /// (0x000D) insertion or deletion. `papx_runs` is the complete PAPX FKP CP
    /// tree in the post-edit global coordinate space; no formatting
    /// inheritance is inferred. The target part's official story/note/comment/
    /// textbox boundaries and guard characters are validated transactionally.
    /// Cell marks (0x0007) and section marks (0x000C) must remain unchanged.
    pub fn replace_text_range_with_papx_runs(
        &mut self,
        part: FieldDocumentPart,
        range: Range<u32>,
        replacement: TextPieceCharacters,
        papx_runs: Vec<DocPapxRun>,
    ) -> Result<()> {
        let mut edited = self.clone();
        edited.replace_text_range_composed(
            part,
            range,
            replacement,
            ParagraphMarkEdit::ExplicitPapx,
        )?;
        edited.word_document.papx_runs = Some(papx_runs);
        edited.validate_current_papx_runs()?;
        edited.validate_document_part_structure(part)?;
        *self = edited;
        Ok(())
    }

    /// Resolves the direct paragraph and character formatting at a CP that is
    /// local to an MS-DOC document part.
    ///
    /// The result follows MS-DOC 2.4.6.1: PAPX/CHPX properties come first and
    /// paragraph/character properties selected by the containing Pcd.Prm come
    /// second. It deliberately stops before style, list, and table-style
    /// evaluation so every returned Rust node still has one physical owner.
    pub fn direct_formatting_at_cp(
        &self,
        part: FieldDocumentPart,
        local_cp: u32,
    ) -> Result<DocDirectFormatting> {
        let (part_start, part_len) = document_part_range(&self.word_document.fib, part)?;
        if local_cp >= part_len {
            return Err(Error::invalid(
                u64::from(local_cp),
                "direct-formatting CP exceeds its MS-DOC document part",
            ));
        }
        let global_cp = part_start
            .checked_add(local_cp)
            .ok_or_else(|| Error::Limit("DOC direct-formatting CP overflow".into()))?;
        let (piece_index, _) = self
            .word_document
            .text_pieces
            .iter()
            .enumerate()
            .find(|(_, piece)| {
                u32::try_from(piece.value.cp_start).is_ok_and(|start| start <= global_cp)
                    && u32::try_from(piece.value.cp_end).is_ok_and(|end| global_cp < end)
            })
            .ok_or_else(|| {
                Error::invalid(
                    u64::from(global_cp),
                    "direct-formatting CP has no containing PlcPcd text piece",
                )
            })?;
        let descriptor = self
            .table
            .clx
            .value
            .piece_table
            .pieces
            .get(piece_index)
            .ok_or_else(|| {
                Error::invalid(
                    u64::try_from(piece_index).unwrap_or(u64::MAX),
                    "direct-formatting text piece has no Pcd descriptor",
                )
            })?;
        let piece_properties = descriptor
            .property_modifier
            .property_modifications(&self.table.clx.value)?;

        let papx = self
            .word_document
            .papx_runs
            .as_ref()
            .ok_or_else(|| Error::invalid(0, "DOC PAPX CP tree is unavailable"))?
            .iter()
            .find(|run| run.cp_start <= global_cp && global_cp < run.cp_end)
            .ok_or_else(|| {
                Error::invalid(
                    u64::from(global_cp),
                    "direct-formatting CP has no containing PAPX run",
                )
            })?;
        let (style_index, papx_properties) = papx
            .properties
            .as_ref()
            .map(|properties| (properties.style_index, properties.properties.clone()))
            .unwrap_or_else(|| {
                (
                    0,
                    GrpPrl {
                        properties: Vec::new(),
                    },
                )
            });

        let chpx = self
            .word_document
            .chpx_runs
            .as_ref()
            .ok_or_else(|| Error::invalid(0, "DOC CHPX CP tree is unavailable"))?
            .iter()
            .find(|run| run.cp_start <= global_cp && global_cp < run.cp_end)
            .ok_or_else(|| {
                Error::invalid(
                    u64::from(global_cp),
                    "direct-formatting CP has no containing CHPX run",
                )
            })?;

        let piece_paragraph_properties = grpprl_for_group(&piece_properties, SprmGroup::Paragraph);
        let mut applied_paragraph_properties = expand_direct_paragraph_properties(
            &papx_properties,
            self.data.as_ref(),
            Some(style_index),
        )?;
        applied_paragraph_properties.properties.extend(
            expand_direct_paragraph_properties(
                &piece_paragraph_properties,
                self.data.as_ref(),
                None,
            )?
            .properties,
        );

        Ok(DocDirectFormatting {
            part,
            local_cp,
            global_cp,
            piece_index,
            paragraph: DocDirectParagraphFormatting {
                style_index,
                papx_properties,
                piece_properties: piece_paragraph_properties,
                applied_properties: applied_paragraph_properties,
            },
            character: DocDirectCharacterFormatting {
                chpx_properties: chpx.properties.clone().unwrap_or_else(|| GrpPrl {
                    properties: Vec::new(),
                }),
                piece_properties: grpprl_for_group(&piece_properties, SprmGroup::Character),
            },
        })
    }

    /// Resolves an STSH style hierarchy into base-first property arrays as
    /// specified by MS-DOC 2.4.6.5.
    pub fn style_properties(&self, style_index: u16) -> Result<DocStyleProperties> {
        let styles = self
            .table
            .styles
            .as_ref()
            .ok_or_else(|| Error::invalid(0, "DOC has no STSH style sheet"))?;
        let mut active = BTreeSet::new();
        let mut lineage = Vec::new();
        let mut paragraph_properties = Vec::new();
        let mut character_properties = Vec::new();
        let mut table_properties = Vec::new();
        collect_style_properties(
            &styles.value,
            style_index,
            &mut active,
            &mut lineage,
            &mut paragraph_properties,
            &mut character_properties,
            &mut table_properties,
        )?;
        let definition = styles
            .value
            .styles
            .get(usize::from(style_index))
            .and_then(|style| style.definition.as_ref())
            .ok_or_else(|| {
                Error::invalid(
                    u64::from(style_index),
                    "requested STSH style is unavailable",
                )
            })?;
        Ok(DocStyleProperties {
            style_index,
            style_kind: definition.base.style_kind,
            lineage,
            paragraph_properties: GrpPrl {
                properties: paragraph_properties,
            },
            character_properties: GrpPrl {
                properties: character_properties,
            },
            table_properties: GrpPrl {
                properties: table_properties,
            },
        })
    }

    /// Determines the effective sprmCFSpec Boolean at a document-part CP.
    /// The paragraph style hierarchy is applied first, followed by CHPX and
    /// Pcd.Prm direct character formatting. Toggle operands 0x80/0x81 retain
    /// their normative relationship to the current style value.
    pub fn effective_cf_spec_at_cp(&self, part: FieldDocumentPart, local_cp: u32) -> Result<bool> {
        let direct = self.direct_formatting_at_cp(part, local_cp)?;
        let mut style_index = direct.paragraph.style_index;
        let papx_properties = expand_direct_paragraph_properties(
            &direct.paragraph.papx_properties,
            self.data.as_ref(),
            Some(style_index),
        )?;
        for property in &papx_properties.properties {
            if property.sprm.kind() != SprmKind::Known(KnownSprm::PIstdPermute) {
                continue;
            }
            let SprmOperand::StylePermutation(permutation) = &property.operand else {
                return Err(Error::invalid(
                    0,
                    "sprmPIstdPermute operand is not SPPOperand",
                ));
            };
            if let Some(remapped) = permutation.remap(style_index) {
                style_index = remapped;
            }
        }
        let style = self.style_properties(style_index)?;
        if style.style_kind != super::StyleKind::Paragraph {
            return Err(Error::invalid(
                u64::from(style_index),
                "paragraph formatting references a non-paragraph style",
            ));
        }
        let style_value = apply_cf_spec_toggles(
            &style.character_properties,
            false,
            false,
            "style character properties",
        )?;
        let value = apply_cf_spec_toggles(
            &direct.character.chpx_properties,
            style_value,
            style_value,
            "CHPX character properties",
        )?;
        apply_cf_spec_toggles(
            &direct.character.piece_properties,
            value,
            style_value,
            "Pcd.Prm character properties",
        )
    }

    /// Validates the text and boundary invariants owned by one MS-DOC
    /// document part. CPs stored in the part's PLCs are interpreted relative
    /// to that part, while text lookup is performed in the aggregate global CP
    /// space. Direct paragraph formatting is assembled from PAPX, Pcd.Prm, and
    /// referenced PrcData so comment-ending table depth can be validated. STSH
    /// inheritance and direct character formatting are also evaluated for the
    /// comment marker's effective sprmCFSpec. This method does not claim to
    /// produce the complete list/table-style/conditional formatting state for
    /// arbitrary content.
    pub fn validate_document_part_structure(&self, part: FieldDocumentPart) -> Result<()> {
        self.validate_main_document_final_paragraph_mark()?;
        if part == FieldDocumentPart::Main {
            return Ok(());
        }
        if part == FieldDocumentPart::Macro {
            return Err(Error::invalid(
                0,
                "the macro field table is not an MS-DOC document part",
            ));
        }
        self.validate_additional_document_paragraph_mark()?;
        let (part_start, part_len) = document_part_range(&self.word_document.fib, part)?;
        match part {
            FieldDocumentPart::Header => self.validate_header_document(part_start, part_len),
            FieldDocumentPart::Footnote | FieldDocumentPart::Endnote => {
                self.validate_note_document(part, part_start, part_len)
            }
            FieldDocumentPart::Comment => self.validate_comment_document(part_start, part_len),
            FieldDocumentPart::Textbox | FieldDocumentPart::HeaderTextbox => {
                self.validate_textbox_document(part, part_start, part_len)
            }
            FieldDocumentPart::Main | FieldDocumentPart::Macro => unreachable!(),
        }
    }

    fn realign_papx_runs(&mut self) -> Result<()> {
        let Some(runs) = &mut self.word_document.papx_runs else {
            if self.word_document.rebuild_paragraph_formatting {
                return Err(Error::invalid(
                    0,
                    "DOC PAPX CP tree is unavailable for a structural text edit",
                ));
            }
            return Ok(());
        };
        let ranges = current_paragraph_ranges(&self.word_document.text_pieces)?;
        if runs.len() != ranges.len() {
            return Err(Error::invalid(
                0,
                "paragraph terminator cardinality changed during PAPX realignment",
            ));
        }
        for (run, (cp_start, cp_end)) in runs.iter_mut().zip(ranges) {
            run.cp_start = cp_start;
            run.cp_end = cp_end;
        }
        Ok(())
    }

    fn validate_current_papx_runs(&self) -> Result<()> {
        let runs = self
            .word_document
            .papx_runs
            .as_ref()
            .ok_or_else(|| Error::invalid(0, "DOC PAPX CP tree is unavailable"))?;
        let ranges = current_paragraph_ranges(&self.word_document.text_pieces)?;
        if runs.len() != ranges.len()
            || runs
                .iter()
                .zip(ranges)
                .any(|(run, range)| (run.cp_start, run.cp_end) != range)
        {
            return Err(Error::invalid(
                0,
                "supplied PAPX CP tree does not match post-edit paragraph ranges",
            ));
        }
        Ok(())
    }

    fn validate_main_document_final_paragraph_mark(&self) -> Result<()> {
        let main_len = u32::try_from(self.word_document.fib.rg_lw.ccp_text)
            .map_err(|_| Error::invalid(0, "Main Document character count is negative"))?;
        if main_len == 0
            || text_value_at_cp(&self.word_document.text_pieces, main_len - 1)? != 0x000d
        {
            return Err(Error::invalid(
                u64::from(main_len),
                "Main Document must end with a paragraph mark",
            ));
        }
        Ok(())
    }

    fn validate_additional_document_paragraph_mark(&self) -> Result<()> {
        let parts = document_part_lengths(&self.word_document.fib)?;
        if !parts.iter().skip(1).any(|(_, length)| *length != 0) {
            return Ok(());
        }
        let end = parts.iter().try_fold(0u32, |total, (_, length)| {
            total
                .checked_add(*length)
                .ok_or_else(|| Error::Limit("DOC document-part CP limit overflow".into()))
        })?;
        if text_value_at_cp(&self.word_document.text_pieces, end)? != 0x000d {
            return Err(Error::invalid(
                u64::from(end),
                "non-empty secondary document parts require an additional paragraph mark",
            ));
        }
        Ok(())
    }

    fn validate_header_document(&self, part_start: u32, part_len: u32) -> Result<()> {
        let Some(table) = &self.table.header_text else {
            return if part_len == 0 {
                Ok(())
            } else {
                Err(Error::invalid(
                    0,
                    "non-empty Header Document has no PlcfHdd",
                ))
            };
        };
        if part_len == 0 {
            return Err(Error::invalid(0, "empty Header Document has a PlcfHdd"));
        }
        if table.value.boundaries.len() < 2 {
            return Err(Error::invalid(0, "PlcfHdd has fewer than two terminal CPs"));
        }
        let story_count = table.value.boundaries.len() - 2;
        let expected_story_count = self
            .table
            .sections
            .value
            .sections
            .len()
            .checked_mul(6)
            .and_then(|count| count.checked_add(6))
            .ok_or_else(|| Error::Limit("Header Document story count overflow".into()))?;
        if story_count != expected_story_count {
            return Err(Error::invalid(
                0,
                format!(
                    "PlcfHdd contains {story_count} stories; {expected_story_count} are required for the section table"
                ),
            ));
        }
        let positions = table.value.boundaries[..table.value.boundaries.len() - 1]
            .iter()
            .map(|boundary| match boundary {
                super::HeaderStoryBoundary::Position(value) => Ok(*value),
                super::HeaderStoryBoundary::Missing => Err(Error::invalid(
                    0,
                    "PlcfHdd has an undefined CP before its final ignored CP",
                )),
            })
            .collect::<Result<Vec<_>>>()?;
        validate_nondecreasing_part_positions(&positions, part_len, "PlcfHdd CP")?;
        if positions.last().copied() != Some(part_len - 1) {
            return Err(Error::invalid(
                0,
                "PlcfHdd second-to-last CP is not ccpHdd - 1",
            ));
        }
        for (index, range) in positions.windows(2).enumerate() {
            let [start, end] = [range[0], range[1]];
            if start == end {
                continue;
            }
            require_part_character(
                &self.word_document.text_pieces,
                part_start,
                end - 1,
                0x000d,
                "Header Document story guard",
            )?;
            if index >= 6 {
                if end - start < 2 {
                    return Err(Error::invalid(
                        u64::from(part_start + start),
                        "non-empty header/footer story lacks a content paragraph mark and guard",
                    ));
                }
                require_part_character(
                    &self.word_document.text_pieces,
                    part_start,
                    end - 2,
                    0x000d,
                    "header/footer content paragraph mark",
                )?;
            }
        }
        Ok(())
    }

    fn validate_note_document(
        &self,
        part: FieldDocumentPart,
        part_start: u32,
        part_len: u32,
    ) -> Result<()> {
        let tables = match part {
            FieldDocumentPart::Footnote => self.table.footnotes.as_ref(),
            FieldDocumentPart::Endnote => self.table.endnotes.as_ref(),
            _ => unreachable!(),
        };
        let Some(tables) = tables else {
            return if part_len == 0 {
                Ok(())
            } else {
                Err(Error::invalid(
                    0,
                    format!("non-empty {part:?} Document has no text/reference PLCs"),
                ))
            };
        };
        if part_len == 0 {
            return Err(Error::invalid(
                0,
                format!("empty {part:?} Document has text/reference PLCs"),
            ));
        }
        let positions = &tables.text.value.positions;
        if positions.len() < 2 || positions.len() - 2 != tables.references.value.indices.len() {
            return Err(Error::invalid(
                0,
                format!("{part:?} text/reference cardinality differs"),
            ));
        }
        validate_strict_part_positions(
            &positions[..positions.len() - 1],
            part_len,
            &format!("{part:?} text CP"),
        )?;
        if positions[positions.len() - 2] != part_len - 1 {
            return Err(Error::invalid(
                0,
                format!("{part:?} second-to-last text CP is not ccp - 1"),
            ));
        }
        for range in positions[..positions.len() - 1].windows(2) {
            require_part_character(
                &self.word_document.text_pieces,
                part_start,
                range[1] - 1,
                0x000d,
                &format!("{part:?} range paragraph mark"),
            )?;
        }
        Ok(())
    }

    fn validate_comment_document(&self, part_start: u32, part_len: u32) -> Result<()> {
        let Some(tables) = &self.table.annotations else {
            return if part_len == 0 {
                Ok(())
            } else {
                Err(Error::invalid(
                    0,
                    "non-empty Comment Document has no PlcfandTxt/PlcfandRef",
                ))
            };
        };
        if part_len == 0 {
            return Err(Error::invalid(
                0,
                "empty Comment Document has PlcfandTxt/PlcfandRef",
            ));
        }
        let positions = &tables.text.value.positions;
        if positions.len() < 2 || positions.len() - 2 != tables.references.value.annotations.len() {
            return Err(Error::invalid(
                0,
                "comment text/reference cardinality differs",
            ));
        }
        validate_strict_part_positions(
            &positions[..positions.len() - 1],
            part_len,
            "PlcfandTxt CP",
        )?;
        if positions[positions.len() - 2] != part_len - 1 {
            return Err(Error::invalid(
                0,
                "PlcfandTxt second-to-last CP is not ccpAtn - 1",
            ));
        }
        for range in positions[..positions.len() - 1].windows(2) {
            require_part_character(
                &self.word_document.text_pieces,
                part_start,
                range[0],
                0x0005,
                "comment range marker",
            )?;
            if !self.effective_cf_spec_at_cp(FieldDocumentPart::Comment, range[0])? {
                return Err(Error::invalid(
                    u64::from(part_start + range[0]),
                    "comment range marker does not have effective sprmCFSpec=1",
                ));
            }
            require_part_character(
                &self.word_document.text_pieces,
                part_start,
                range[1] - 1,
                0x000d,
                "comment range paragraph mark",
            )?;
            let table_state = self
                .direct_formatting_at_cp(FieldDocumentPart::Comment, range[1] - 1)?
                .paragraph
                .table_state()?;
            if table_state.in_table || table_state.depth != 0 {
                return Err(Error::invalid(
                    u64::from(part_start + range[1] - 1),
                    "comment range does not end at table depth zero",
                ));
            }
        }
        Ok(())
    }

    fn validate_textbox_document(
        &self,
        part: FieldDocumentPart,
        part_start: u32,
        part_len: u32,
    ) -> Result<()> {
        let key = match part {
            FieldDocumentPart::Textbox => TextboxDocumentPart::Main,
            FieldDocumentPart::HeaderTextbox => TextboxDocumentPart::Header,
            _ => unreachable!(),
        };
        let Some(table) = self.table.textbox_stories.get(&key) else {
            return if part_len == 0 {
                Ok(())
            } else {
                Err(Error::invalid(
                    0,
                    format!("non-empty {part:?} Document has no textbox story PLC"),
                ))
            };
        };
        if part_len == 0 {
            return Err(Error::invalid(
                0,
                format!("empty {part:?} Document has a textbox story PLC"),
            ));
        }
        let positions = &table.value.positions;
        if positions.len() != table.value.stories.len().saturating_add(1)
            || table.value.stories.is_empty()
        {
            return Err(Error::invalid(
                0,
                format!("{part:?} textbox CP/FTXBXS cardinality differs"),
            ));
        }
        validate_strict_textbox_positions(positions, part_len, &format!("{part:?} textbox CP"))?;
        for range in positions.windows(2).take(table.value.stories.len() - 1) {
            require_part_character(
                &self.word_document.text_pieces,
                part_start,
                range[1] - 1,
                0x000d,
                &format!("{part:?} textbox range paragraph mark"),
            )?;
        }
        Ok(())
    }

    fn replace_text_range_composed(
        &mut self,
        part: FieldDocumentPart,
        range: Range<u32>,
        replacement: TextPieceCharacters,
        paragraph_marks: ParagraphMarkEdit,
    ) -> Result<()> {
        if range.start > range.end {
            return Err(Error::invalid(0, "DOC text replacement range is reversed"));
        }
        let (part_start, part_len) = document_part_range(&self.word_document.fib, part)?;
        if range.end > part_len {
            return Err(Error::invalid(
                u64::from(range.end),
                "DOC text replacement exceeds its document part",
            ));
        }
        let global_start = part_start
            .checked_add(range.start)
            .ok_or_else(|| Error::Limit("DOC global text edit start overflow".into()))?;
        let global_end = part_start
            .checked_add(range.end)
            .ok_or_else(|| Error::Limit("DOC global text edit limit overflow".into()))?;
        let removed_terminators = paragraph_terminators_in_piece_range(
            &self.word_document.text_pieces,
            global_start,
            global_end,
        )?;
        let replacement_terminators = paragraph_terminators(&replacement);
        let terminators_match = match paragraph_marks {
            ParagraphMarkEdit::PreserveAll => removed_terminators == replacement_terminators,
            ParagraphMarkEdit::ExplicitPapx => {
                non_paragraph_terminators(&removed_terminators)
                    == non_paragraph_terminators(&replacement_terminators)
            }
        };
        if !terminators_match {
            return Err(Error::invalid(
                u64::from(global_start),
                match paragraph_marks {
                    ParagraphMarkEdit::PreserveAll => {
                        "DOC text replacement changes the paragraph/cell/section terminator sequence"
                    }
                    ParagraphMarkEdit::ExplicitPapx => {
                        "DOC explicit PAPX replacement changes a cell or section mark"
                    }
                },
            ));
        }
        let segments = self
            .word_document
            .text_pieces
            .iter()
            .filter_map(|piece| {
                let start = u32::try_from(piece.value.cp_start).ok()?;
                let end = u32::try_from(piece.value.cp_end).ok()?;
                let overlap_start = global_start.max(start);
                let overlap_end = global_end.min(end);
                (overlap_start < overlap_end).then_some((
                    overlap_start,
                    overlap_end,
                    matches!(piece.value.characters, TextPieceCharacters::Compressed(_)),
                ))
            })
            .collect::<Vec<_>>();
        if segments.len() <= 1 {
            return self.replace_text_range_inner(part, range, replacement);
        }
        for (index, (start, end, compressed)) in segments.iter().copied().enumerate().rev() {
            let piece_replacement = if index == 0 {
                replacement.clone()
            } else if compressed {
                TextPieceCharacters::Compressed(Vec::new())
            } else {
                TextPieceCharacters::Utf16(Vec::new())
            };
            self.replace_text_range_inner(
                part,
                (start - part_start)..(end - part_start),
                piece_replacement,
            )?;
        }
        Ok(())
    }

    fn replace_text_range_inner(
        &mut self,
        part: FieldDocumentPart,
        range: Range<u32>,
        replacement: TextPieceCharacters,
    ) -> Result<()> {
        if range.start > range.end {
            return Err(Error::invalid(0, "DOC text replacement range is reversed"));
        }
        if !self.table.compatibility_tables.is_empty() {
            return Err(Error::invalid(
                0,
                "DOC text relocation cannot preserve opaque compatibility tables",
            ));
        }
        if matches!(
            self.word_document.fib.version(),
            super::FibVersion::Compatibility(_)
        ) {
            return Err(Error::invalid(
                0,
                "DOC text relocation requires a documented FIB version",
            ));
        }
        let (part_start, part_len) = document_part_range(&self.word_document.fib, part)?;
        if range.end > part_len {
            return Err(Error::invalid(
                u64::from(range.end),
                "DOC text replacement exceeds its document part",
            ));
        }
        let global_start = part_start
            .checked_add(range.start)
            .ok_or_else(|| Error::Limit("DOC global text edit start overflow".into()))?;
        let global_end = part_start
            .checked_add(range.end)
            .ok_or_else(|| Error::Limit("DOC global text edit limit overflow".into()))?;
        let source_clx = Clx::from_bytes(bounded_slice(
            &self.table.physical_bytes,
            self.table.clx.location,
            "source CLX",
        )?)?;
        let piece_index = self
            .word_document
            .text_pieces
            .iter()
            .position(|piece| {
                let start = u32::try_from(piece.value.cp_start).ok();
                let end = u32::try_from(piece.value.cp_end).ok();
                start.is_some_and(|start| start <= global_start)
                    && end.is_some_and(|end| global_end <= end)
            })
            .ok_or_else(|| {
                Error::invalid(
                    u64::from(global_start),
                    "DOC text replacement crosses a text-piece boundary",
                )
            })?;
        let source_piece_index = *self
            .word_document
            .source_piece_indices
            .get(piece_index)
            .ok_or_else(|| Error::invalid(0, "source text piece identity is missing"))?;
        let piece = &self.word_document.text_pieces[piece_index].value;
        let source_piece_start = *source_clx
            .piece_table
            .character_positions
            .get(source_piece_index)
            .ok_or_else(|| Error::invalid(0, "source text piece CP start is missing"))?;
        let source_piece_end = *source_clx
            .piece_table
            .character_positions
            .get(source_piece_index + 1)
            .ok_or_else(|| Error::invalid(0, "source text piece CP limit is missing"))?;
        let source_piece_count = source_piece_end
            .checked_sub(source_piece_start)
            .and_then(|value| usize::try_from(value).ok())
            .ok_or_else(|| Error::invalid(0, "source text piece CP range is invalid"))?;
        let prior_edits = self
            .word_document
            .pending_text_edits
            .get(&source_piece_index)
            .cloned()
            .unwrap_or_default();
        let current_piece_count = relocate_character_position(
            u32::try_from(source_piece_count).map_err(|_| {
                Error::Limit("source text piece character count exceeds u32".into())
            })?,
            &prior_edits,
            "text piece character count",
        )?;
        if usize::try_from(current_piece_count).ok() != Some(piece.character_count()) {
            return Err(Error::invalid(
                u64::from(global_start),
                "text piece has an untracked variable-length edit after replace_text_range",
            ));
        }
        let piece_start = u32::try_from(piece.cp_start)
            .map_err(|_| Error::invalid(0, "text piece begins at a negative CP"))?;
        let local_start = usize::try_from(global_start - piece_start)
            .map_err(|_| Error::Limit("text replacement start exceeds usize".into()))?;
        let local_end = usize::try_from(global_end - piece_start)
            .map_err(|_| Error::Limit("text replacement end exceeds usize".into()))?;
        let replacement_len = replacement.character_count();
        let local_edit = CpReplacement::new(
            u32::try_from(local_start)
                .map_err(|_| Error::Limit("text replacement start exceeds u32".into()))?,
            u32::try_from(local_end)
                .map_err(|_| Error::Limit("text replacement end exceeds u32".into()))?,
            u32::try_from(replacement_len)
                .map_err(|_| Error::Limit("replacement character count exceeds u32".into()))?,
        )?;
        let source_descriptor = source_clx
            .piece_table
            .pieces
            .get(source_piece_index)
            .ok_or_else(|| Error::invalid(0, "source text piece descriptor is missing"))?;
        let rebuild_formatting = validate_formatting_edit(
            self,
            source_descriptor.file_position.byte_offset(),
            if source_descriptor.file_position.compressed {
                1
            } else {
                2
            },
            source_piece_count,
            &prior_edits,
            &local_edit,
        )?;
        let piece = &mut self.word_document.text_pieces[piece_index].value;
        match (&mut piece.characters, replacement) {
            (
                TextPieceCharacters::Compressed(characters),
                TextPieceCharacters::Compressed(value),
            ) => {
                characters.splice(local_start..local_end, value);
            }
            (TextPieceCharacters::Utf16(characters), TextPieceCharacters::Utf16(value)) => {
                characters.splice(local_start..local_end, value);
            }
            _ => {
                return Err(Error::invalid(
                    u64::from(range.start),
                    "DOC text replacement encoding differs from its text piece",
                ));
            }
        }
        let remove_piece = piece.character_count() == 0;
        if remove_piece && self.word_document.text_pieces.len() == 1 {
            return Err(Error::invalid(
                u64::from(global_start),
                "DOC text replacement would remove the final PlcPcd text piece",
            ));
        }
        let replacement_len = u32::try_from(replacement_len)
            .map_err(|_| Error::Limit("DOC replacement character count exceeds u32".into()))?;
        let global_edit = CpReplacement::new(global_start, global_end, replacement_len)?;
        let part_edit = CpReplacement::new(range.start, range.end, replacement_len)?;

        for position in self
            .table
            .clx
            .value
            .piece_table
            .character_positions
            .iter_mut()
            .skip(piece_index + 1)
        {
            *position = global_edit.relocate_i32(*position, "PlcPcd CP")?;
        }
        for text_piece in self.word_document.text_pieces.iter_mut().skip(piece_index) {
            if text_piece.piece_index == piece_index {
                text_piece.value.cp_end =
                    global_edit.relocate_i32(text_piece.value.cp_end, "text piece CP limit")?;
            } else {
                text_piece.value.cp_start =
                    global_edit.relocate_i32(text_piece.value.cp_start, "text piece CP start")?;
                text_piece.value.cp_end =
                    global_edit.relocate_i32(text_piece.value.cp_end, "text piece CP limit")?;
            }
        }
        if remove_piece {
            self.table.clx.value.piece_table.pieces.remove(piece_index);
            self.table
                .clx
                .value
                .piece_table
                .character_positions
                .remove(piece_index + 1);
            self.word_document.text_pieces.remove(piece_index);
            self.word_document.source_piece_indices.remove(piece_index);
            for (index, piece) in self
                .word_document
                .text_pieces
                .iter_mut()
                .enumerate()
                .skip(piece_index)
            {
                piece.piece_index = index;
            }
        }
        set_document_part_length(
            &mut self.word_document.fib,
            part,
            part_edit.relocate_u32(part_len, "FIB document-part character count")?,
        )?;
        if part == FieldDocumentPart::Main {
            relocate_main_document_cps(&mut self.table, &part_edit)?;
        } else {
            relocate_global_document_cps(&mut self.table, &global_edit)?;
            relocate_non_main_document_part_cps(&mut self.table, part, &part_edit)?;
        }
        if remove_piece {
            self.word_document
                .pending_text_edits
                .remove(&source_piece_index);
        } else {
            self.word_document
                .pending_text_edits
                .entry(source_piece_index)
                .or_default()
                .push(local_edit);
        }
        if let Some(runs) = &mut self.word_document.chpx_runs {
            apply_character_run_edit(runs, &global_edit)?;
        } else if rebuild_formatting.character || remove_piece {
            return Err(Error::invalid(
                u64::from(global_start),
                "DOC CHPX CP tree is unavailable for a structural text edit",
            ));
        }
        self.word_document.rebuild_character_formatting |=
            rebuild_formatting.character || remove_piece;
        self.word_document.rebuild_paragraph_formatting |=
            rebuild_formatting.paragraph || remove_piece;
        Ok(())
    }

    fn validate_links(&self) -> Result<()> {
        let fib = &self.word_document.fib;
        validate_compatibility_tables(
            &self.table.physical_bytes,
            &self.table.compatibility_tables,
        )?;
        validate_object_pool_links(&self.compound_file, self.object_pool.as_ref())?;
        validate_data_links(&self.word_document, &self.table, self.data.as_ref())?;
        let expected_name = if fib.base.flags.contains(FibBaseFlags::USE_1_TABLE) {
            DocTableStreamName::Table1
        } else {
            DocTableStreamName::Table0
        };
        if self.table.name != expected_name
            || fib.clx_location() != Some(self.table.clx.location)
            || fib.chpx_bte_location() != Some(self.table.character_bin_table.location)
            || fib.papx_bte_location() != Some(self.table.paragraph_bin_table.location)
            || fib.section_table_location() != Some(self.table.sections.location)
            || fib.style_sheet_location().filter(|v| v.lcb != 0)
                != self.table.styles.as_ref().map(|v| v.location)
            || fib.font_table_location().filter(|v| v.lcb != 0)
                != self.table.fonts.as_ref().map(|v| v.location)
            || fib.office_art_content_location().filter(|v| v.lcb != 0)
                != self.table.office_art.as_ref().map(|v| v.location)
        {
            return Err(Error::invalid(0, "DOC FIB/content-tree links changed"));
        }
        let expected_fields = fib
            .field_table_locations()
            .into_iter()
            .filter(|(_, location)| location.lcb != 0)
            .collect::<BTreeMap<_, _>>();
        let actual_fields = self
            .table
            .fields
            .iter()
            .map(|(part, field)| (*part, field.location))
            .collect::<BTreeMap<_, _>>();
        if actual_fields != expected_fields {
            return Err(Error::invalid(0, "DOC FIB/field-table links changed"));
        }
        let expected_bookmarks = match fib.bookmark_locations() {
            None => None,
            Some(locations) => {
                let locations = [locations.0, locations.1, locations.2];
                if locations.iter().all(|location| location.lcb == 0) {
                    None
                } else if locations.iter().all(|location| location.lcb != 0) {
                    Some(locations)
                } else {
                    return Err(Error::invalid(0, "DOC bookmark locations are incomplete"));
                }
            }
        };
        let actual_bookmarks = self.table.bookmarks.as_ref().map(|bookmarks| {
            [
                bookmarks.names_location,
                bookmarks.starts_location,
                bookmarks.ends_location,
            ]
        });
        if expected_bookmarks != actual_bookmarks {
            return Err(Error::invalid(0, "DOC FIB/bookmark links changed"));
        }
        validate_optional_location(
            fib.header_text_location(),
            self.table.header_text.as_ref(),
            "header text",
        )?;
        validate_note_locations(
            fib.footnote_locations(),
            self.table.footnotes.as_ref(),
            "footnote",
        )?;
        validate_note_locations(
            fib.endnote_locations(),
            self.table.endnotes.as_ref(),
            "endnote",
        )?;
        validate_annotation_locations(fib.annotation_locations(), self.table.annotations.as_ref())?;
        validate_optional_location(
            fib.annotation_owner_location(),
            self.table.annotation_owners.as_ref(),
            "annotation owners",
        )?;
        let expected_annotation_bookmarks =
            fib.annotation_bookmark_locations()
                .and_then(|(infos, starts, ends)| {
                    [infos, starts, ends]
                        .iter()
                        .all(|location| location.lcb != 0)
                        .then_some([infos, starts, ends])
                });
        let actual_annotation_bookmarks = self.table.annotation_bookmarks.as_ref().map(|value| {
            [
                value.infos_location,
                value.starts_location,
                value.ends_location,
            ]
        });
        if expected_annotation_bookmarks != actual_annotation_bookmarks {
            return Err(Error::invalid(
                0,
                "DOC FIB/annotation-bookmark links changed",
            ));
        }
        validate_optional_location(
            managed_expected_location(
                fib.annotation_extended_data_location(),
                &self.table.compatibility_tables,
                "AtrdExtra",
            ),
            self.table.annotation_extended_data.as_ref(),
            "annotation extended data",
        )?;
        validate_part_locations(
            fib.textbox_story_locations(),
            &self.table.textbox_stories,
            "textbox stories",
        )?;
        validate_part_locations(
            fib.textbox_break_locations(),
            &self.table.textbox_breaks,
            "textbox breaks",
        )?;
        validate_part_locations(
            fib.shape_anchor_locations(),
            &self.table.shape_anchors,
            "shape anchors",
        )?;
        validate_optional_location(
            fib.revision_authors_location(),
            self.table.revision_authors.as_ref(),
            "revision authors",
        )?;
        validate_caption_locations(fib.caption_locations(), self.table.captions.as_ref())?;
        validate_optional_location(
            managed_expected_location(
                fib.subdocuments_location(),
                &self.table.compatibility_tables,
                "PlcfWkb",
            ),
            self.table.subdocuments.as_ref(),
            "subdocuments",
        )?;
        validate_optional_location(
            fib.user_variables_location(),
            self.table.user_variables.as_ref(),
            "user variables",
        )?;
        validate_optional_location(
            fib.embedded_fonts_location(),
            self.table.embedded_fonts.as_ref(),
            "embedded fonts",
        )?;
        validate_optional_location(
            fib.spelling_state_location(),
            self.table.spelling_state.as_ref(),
            "spelling state",
        )?;
        validate_optional_location(
            fib.grammar_state_location(),
            self.table.grammar_state.as_ref(),
            "grammar state",
        )?;
        validate_optional_location(
            fib.language_detection_state_location(),
            self.table.language_detection_state.as_ref(),
            "language detection state",
        )?;
        if managed_expected_location(
            fib.list_definition_location(),
            &self.table.compatibility_tables,
            "PlfLst",
        ) != self
            .table
            .list_definitions
            .as_ref()
            .map(|value| value.location)
        {
            return Err(Error::invalid(0, "DOC FIB/list-definition link changed"));
        }
        validate_optional_location(
            fib.list_names_location(),
            self.table.list_names.as_ref(),
            "list names",
        )?;
        validate_optional_location(
            fib.list_override_location(),
            self.table.list_overrides.as_ref(),
            "list overrides",
        )?;
        validate_optional_location(
            fib.document_properties_location(),
            self.table.document_properties.as_ref(),
            "document properties",
        )?;
        validate_optional_location(
            fib.associated_strings_location(),
            self.table.associated_strings.as_ref(),
            "associated strings",
        )?;
        validate_optional_location(
            fib.external_file_names_location(),
            self.table.external_file_names.as_ref(),
            "external file names",
        )?;
        macro_rules! validate_table {
            ($location:expr, $field:ident, $physical:literal, $link:literal) => {
                validate_compatible_location(
                    $location,
                    self.table.$field.as_ref(),
                    &self.table.compatibility_tables,
                    $physical,
                    $link,
                )?;
            };
        }
        validate_table!(
            fib.mail_merge_state_location(),
            mail_merge_state,
            "Pms",
            "mail merge state"
        );
        validate_table!(
            fib.new_mail_merge_state_location(),
            new_mail_merge_state,
            "PmsNew",
            "new mail merge state"
        );
        validate_table!(
            fib.office_data_source_location(),
            office_data_source,
            "Odso",
            "office data source"
        );
        validate_table!(
            fib.printer_driver_info_location(),
            printer_driver_info,
            "PrDrvr",
            "printer driver info"
        );
        validate_table!(
            fib.ole_control_info_location(),
            ole_control_infos,
            "RgxOcxInfo",
            "OLE control infos"
        );
        validate_table!(
            fib.table_character_cache_location(),
            table_character_cache,
            "PlcfTch",
            "table character cache"
        );
        validate_table!(
            fib.revision_message_threading_location(),
            revision_message_threading,
            "RmdThreading",
            "revision message threading"
        );
        validate_table!(
            fib.list_style_templates_location(),
            list_style_templates,
            "SttbRgtplc",
            "list style templates"
        );
        validate_table!(
            fib.frame_and_list_records_location(),
            frame_and_list_records,
            "RgDofr",
            "frame and list records"
        );
        validate_table!(
            fib.grammar_option_sets_location(),
            grammar_option_sets,
            "PlfCosi",
            "grammar option sets"
        );
        validate_table!(
            fib.legacy_grammar_option_sets_location(),
            legacy_grammar_option_sets,
            "PlfGosl",
            "legacy grammar option sets"
        );
        validate_table!(
            fib.auto_summary_ranges_location(),
            auto_summary_ranges,
            "PlcfAsumy",
            "auto summary ranges"
        );
        validate_table!(
            fib.smart_tag_recognizer_state_location(),
            smart_tag_recognizer_state,
            "PlcfFactoid",
            "smart-tag recognizer state"
        );
        validate_table!(
            fib.xml_schema_references_location(),
            xml_schema_references,
            "Hplxsdr",
            "XML schema references"
        );
        validate_table!(
            fib.xml_transform_path_location(),
            xml_transform_path,
            "CustomXForm",
            "XML transform path"
        );
        validate_table!(
            fib.paragraph_group_properties_location(),
            paragraph_group_properties,
            "PlcfPgp",
            "paragraph group properties"
        );
        validate_table!(
            fib.save_history_location(),
            save_history,
            "SttbSavedBy",
            "save history"
        );
        validate_table!(
            fib.grammar_checker_cookies_location(),
            grammar_checker_cookies,
            "PlcfCookie",
            "grammar checker cookies"
        );
        validate_table!(
            fib.legacy_grammar_checker_cookies_location(),
            legacy_grammar_checker_cookies,
            "PlcfCookieOld",
            "legacy grammar checker cookies"
        );
        validate_table!(
            fib.grammar_cookie_data_location(),
            grammar_cookie_data,
            "CookieData",
            "grammar cookie data"
        );
        validate_table!(
            fib.smart_tag_data_location(),
            smart_tag_data,
            "FactoidData",
            "smart-tag data"
        );
        validate_table!(
            fib.revision_save_ids_location(),
            revision_save_ids,
            "Plrsid",
            "revision save IDs"
        );
        validate_table!(
            fib.selection_state_location(),
            selection_state,
            "Wss",
            "selection state"
        );
        validate_table!(
            fib.command_customizations_location(),
            command_customizations,
            "Cmds",
            "command customizations"
        );
        let expected = managed_expected_locations(
            fib.structured_tag_bookmark_locations(),
            &self.table.compatibility_tables,
            ["SttbfBkmkSdt", "PlcfBkfSdt", "PlcfBklSdt"],
        );
        let actual = self.table.structured_tag_bookmarks.as_ref().map(|value| {
            [
                value.metadata_location,
                value.starts_location,
                value.ends_location,
            ]
        });
        if expected != actual {
            return Err(Error::invalid(
                0,
                "DOC FIB/structured-tag bookmark links changed",
            ));
        }
        let expected = managed_expected_locations(
            fib.range_protection_locations(),
            &self.table.compatibility_tables,
            [
                "SttbfBkmkProt",
                "PlcfBkfProt",
                "PlcfBklProt",
                "SttbProtUser",
            ],
        );
        let actual = self.table.range_protection.as_ref().map(|value| {
            [
                value.permissions_location,
                value.starts_location,
                value.ends_location,
                value.users_location,
            ]
        });
        if expected != actual {
            return Err(Error::invalid(0, "DOC FIB/range-protection links changed"));
        }
        let expected = managed_expected_locations(
            fib.smart_tag_bookmark_locations(),
            &self.table.compatibility_tables,
            ["SttbfBkmkFactoid", "PlcfBkfFactoid", "PlcfBklFactoid"],
        );
        let actual = self.table.smart_tag_bookmarks.as_ref().map(|value| {
            [
                value.metadata_location,
                value.starts_location,
                value.ends_location,
            ]
        });
        if expected != actual {
            return Err(Error::invalid(
                0,
                "DOC FIB/smart-tag bookmark links changed",
            ));
        }
        let expected = managed_expected_locations(
            fib.format_consistency_bookmark_locations(),
            &self.table.compatibility_tables,
            ["SttbfBkmkFcc", "PlcfBkfFcc", "PlcfBklFcc"],
        );
        let actual = self
            .table
            .format_consistency_bookmarks
            .as_ref()
            .map(|value| {
                [
                    value.metadata_location,
                    value.starts_location,
                    value.ends_location,
                ]
            });
        if expected != actual {
            return Err(Error::invalid(
                0,
                "DOC FIB/format-consistency bookmark links changed",
            ));
        }
        let expected = managed_expected_locations(
            fib.repair_bookmark_locations(),
            &self.table.compatibility_tables,
            ["SttbfBkmkBpRepairs", "PlcfBkfBpRepairs", "PlcfBklBpRepairs"],
        );
        let actual = self.table.repair_bookmarks.as_ref().map(|value| {
            [
                value.metadata_location,
                value.starts_location,
                value.ends_location,
            ]
        });
        if expected != actual {
            return Err(Error::invalid(0, "DOC FIB/repair bookmark links changed"));
        }
        let expected = managed_expected_locations(
            fib.user_input_method_locations(),
            &self.table.compatibility_tables,
            ["PlcfUim", "PlfGuidUim"],
        );
        let actual = self
            .table
            .user_input_methods
            .as_ref()
            .map(|value| [value.methods_location, value.service_guids_location]);
        if expected != actual {
            return Err(Error::invalid(0, "DOC FIB/user-input-method links changed"));
        }
        validate_compatible_location(
            fib.mso_envelope_location(),
            self.table.mso_envelope.as_ref(),
            &self.table.compatibility_tables,
            "MsoEnvelope",
            "MsoEnvelope",
        )?;
        if managed_expected_location(
            fib.deprecated_numbering_field_cache_location(),
            &self.table.compatibility_tables,
            "PlcfBteLvc",
        ) != self
            .table
            .deprecated_numbering_field_cache
            .as_ref()
            .map(|value| value.location)
        {
            return Err(Error::invalid(
                0,
                "DOC FIB/deprecated numbering field cache link changed",
            ));
        }

        let pieces = &self.table.clx.value.piece_table;
        if pieces.pieces.len() != self.word_document.text_pieces.len()
            || pieces.character_positions.len() != pieces.pieces.len() + 1
            || self.word_document.source_piece_indices.len() != self.word_document.text_pieces.len()
            || self
                .word_document
                .source_piece_indices
                .windows(2)
                .any(|indices| indices[0] >= indices[1])
        {
            return Err(Error::invalid(0, "CLX/text-piece tree cardinality changed"));
        }
        for (index, (descriptor, piece)) in pieces
            .pieces
            .iter()
            .zip(&self.word_document.text_pieces)
            .enumerate()
        {
            if piece.piece_index != index
                || piece.value.cp_start != pieces.character_positions[index]
                || piece.value.cp_end != pieces.character_positions[index + 1]
                || piece.value.file_offset != descriptor.file_position.byte_offset()
            {
                return Err(Error::invalid(0, "CLX/text-piece link changed"));
            }
        }
        if let Some(runs) = &self.word_document.chpx_runs {
            let normalized = normalize_logical_character_runs(runs.clone())?;
            if normalized != *runs {
                return Err(Error::invalid(0, "CHPX CP tree is not canonical"));
            }
        }
        if let Some(runs) = &self.word_document.papx_runs {
            let ranges = current_paragraph_ranges(&self.word_document.text_pieces)?;
            if runs.len() != ranges.len()
                || runs
                    .iter()
                    .zip(ranges)
                    .any(|(run, range)| (run.cp_start, run.cp_end) != range)
            {
                return Err(Error::invalid(
                    0,
                    "PAPX CP tree does not match the document paragraphs",
                ));
            }
        }
        if self
            .word_document
            .character_format_pages
            .iter()
            .map(|page| page.page)
            .ne(self.table.character_bin_table.value.pages.iter().copied())
            || self
                .word_document
                .paragraph_format_pages
                .iter()
                .map(|page| page.page)
                .ne(self.table.paragraph_bin_table.value.pages.iter().copied())
        {
            return Err(Error::invalid(0, "PlcBte/FKP page links changed"));
        }
        if self.word_document.section_properties.len() != self.table.sections.value.sections.len() {
            return Err(Error::invalid(0, "SED/Sepx tree cardinality changed"));
        }
        for section in &self.word_document.section_properties {
            let sed = self
                .table
                .sections
                .value
                .sections
                .get(section.section_index)
                .ok_or_else(|| Error::invalid(0, "section property index is stale"))?;
            if sed.sepx_offset != section.offset
                || (section.offset == -1) != section.value.is_none()
            {
                return Err(Error::invalid(0, "SED/Sepx link changed"));
            }
        }
        Ok(())
    }
}

fn document_part_lengths(fib: &Fib) -> Result<[(FieldDocumentPart, u32); 7]> {
    let values = [
        (FieldDocumentPart::Main, fib.rg_lw.ccp_text),
        (FieldDocumentPart::Footnote, fib.rg_lw.ccp_footnote),
        (FieldDocumentPart::Header, fib.rg_lw.ccp_header),
        (FieldDocumentPart::Comment, fib.rg_lw.ccp_comment),
        (FieldDocumentPart::Endnote, fib.rg_lw.ccp_endnote),
        (FieldDocumentPart::Textbox, fib.rg_lw.ccp_textbox),
        (
            FieldDocumentPart::HeaderTextbox,
            fib.rg_lw.ccp_header_textbox,
        ),
    ];
    let mut lengths = [(FieldDocumentPart::Main, 0); 7];
    for (index, (part, value)) in values.into_iter().enumerate() {
        lengths[index] = (
            part,
            u32::try_from(value).map_err(|_| {
                Error::invalid(0, format!("FIB character count for {part:?} is negative"))
            })?,
        );
    }
    Ok(lengths)
}

fn document_part_range(fib: &Fib, target: FieldDocumentPart) -> Result<(u32, u32)> {
    if target == FieldDocumentPart::Macro {
        return Err(Error::invalid(
            0,
            "the macro field table has no MS-DOC document-part text range",
        ));
    }
    let mut start = 0u32;
    for (part, length) in document_part_lengths(fib)? {
        if part == target {
            return Ok((start, length));
        }
        start = start
            .checked_add(length)
            .ok_or_else(|| Error::Limit("DOC document-part CP range overflow".into()))?;
    }
    Err(Error::invalid(0, "unknown MS-DOC document part"))
}

fn grpprl_for_group(properties: &GrpPrl, group: SprmGroup) -> GrpPrl {
    GrpPrl {
        properties: properties
            .properties
            .iter()
            .filter(|property| property.sprm.group == group)
            .cloned()
            .collect(),
    }
}

#[allow(clippy::too_many_arguments)]
fn collect_style_properties(
    styles: &StyleSheet,
    style_index: u16,
    active: &mut BTreeSet<u16>,
    lineage: &mut Vec<u16>,
    paragraph: &mut Vec<super::Prl>,
    character: &mut Vec<super::Prl>,
    table: &mut Vec<super::Prl>,
) -> Result<()> {
    if !active.insert(style_index) {
        return Err(Error::invalid(
            u64::from(style_index),
            "STSH base-style references contain a cycle",
        ));
    }
    let definition = styles
        .styles
        .get(usize::from(style_index))
        .and_then(|style| style.definition.as_ref())
        .ok_or_else(|| {
            Error::invalid(
                u64::from(style_index),
                "STSH style hierarchy references an unavailable style",
            )
        })?;
    if definition.base.base_style_index != 0x0fff {
        collect_style_properties(
            styles,
            definition.base.base_style_index,
            active,
            lineage,
            paragraph,
            character,
            table,
        )?;
    }
    lineage.push(style_index);
    match &definition.formatting {
        StyleFormatting::Paragraph {
            paragraph: value,
            character: value_character,
        } => {
            paragraph.extend(value.properties.properties.iter().cloned());
            character.extend(value_character.properties.properties.iter().cloned());
        }
        StyleFormatting::Character { character: value } => {
            character.extend(value.properties.properties.iter().cloned());
        }
        StyleFormatting::RevisionParagraph {
            paragraph: value,
            character: value_character,
            ..
        } => {
            paragraph.extend(value.properties.properties.iter().cloned());
            character.extend(value_character.properties.properties.iter().cloned());
        }
        StyleFormatting::RevisionCharacter {
            character: value, ..
        } => {
            character.extend(value.properties.properties.iter().cloned());
        }
        StyleFormatting::Table {
            table: value_table,
            paragraph: value_paragraph,
            character: value_character,
        } => {
            table.extend(value_table.properties.properties.iter().cloned());
            paragraph.extend(value_paragraph.properties.properties.iter().cloned());
            character.extend(value_character.properties.properties.iter().cloned());
        }
        StyleFormatting::Numbering { paragraph: value } => {
            paragraph.extend(value.properties.properties.iter().cloned());
        }
    }
    active.remove(&style_index);
    Ok(())
}

fn apply_cf_spec_toggles(
    properties: &GrpPrl,
    mut value: bool,
    style_value: bool,
    source: &str,
) -> Result<bool> {
    for property in &properties.properties {
        if property.sprm.kind() != SprmKind::Known(KnownSprm::CFSpec) {
            continue;
        }
        let SprmOperand::Toggle(operand) = &property.operand else {
            return Err(Error::invalid(
                0,
                format!("sprmCFSpec in {source} does not have a ToggleOperand"),
            ));
        };
        value = match *operand {
            0x00 => false,
            0x01 => true,
            0x80 => style_value,
            0x81 => !style_value,
            _ => {
                return Err(Error::invalid(
                    0,
                    format!("sprmCFSpec in {source} has an invalid ToggleOperand"),
                ));
            }
        };
    }
    Ok(value)
}

fn expand_direct_paragraph_properties(
    properties: &GrpPrl,
    data: Option<&DocDataStream>,
    papx_style_index: Option<u16>,
) -> Result<GrpPrl> {
    let mut active_offsets = BTreeSet::new();
    Ok(GrpPrl {
        properties: expand_direct_paragraph_property_array(
            properties,
            data,
            papx_style_index,
            &mut active_offsets,
        )?,
    })
}

fn expand_direct_paragraph_property_array(
    properties: &GrpPrl,
    data: Option<&DocDataStream>,
    papx_style_index: Option<u16>,
    active_offsets: &mut BTreeSet<u32>,
) -> Result<Vec<super::Prl>> {
    let mut applied = Vec::new();
    for (index, property) in properties.properties.iter().enumerate() {
        let reference_kind = match property.sprm.kind() {
            SprmKind::Known(KnownSprm::PHugePapx) if index != 0 => continue,
            SprmKind::Known(KnownSprm::PHugePapx) => {
                if papx_style_index.is_some_and(|style_index| style_index != 0) {
                    return Err(Error::invalid(
                        0,
                        "sprmPHugePapx in PapxInFkp requires style index zero",
                    ));
                }
                "sprmPHugePapx"
            }
            SprmKind::Known(KnownSprm::PTableProps) => "sprmPTableProps",
            _ => {
                applied.push(property.clone());
                continue;
            }
        };
        let SprmOperand::Dword(raw_offset) = &property.operand else {
            return Err(Error::invalid(
                0,
                format!("{reference_kind} operand is not a Data-stream offset"),
            ));
        };
        let offset = u32::from_le_bytes(*raw_offset);
        if !active_offsets.insert(offset) {
            return Err(Error::invalid(
                u64::from(offset),
                "paragraph-property Data references contain a cycle",
            ));
        }
        let node = data
            .and_then(|data| data.nodes.iter().find(|node| node.offset == offset))
            .ok_or_else(|| {
                Error::invalid(
                    u64::from(offset),
                    format!("{reference_kind} target PrcData is unavailable"),
                )
            })?;
        let DocDataNodeValue::ParagraphProperties(referenced) = &node.value else {
            return Err(Error::invalid(
                u64::from(offset),
                format!("{reference_kind} target is not a PrcData node"),
            ));
        };
        if node.physical_len < 12 {
            return Err(Error::invalid(
                u64::from(offset),
                format!("{reference_kind} target PrcData cbGrpprl is less than 10"),
            ));
        }
        applied.extend(expand_direct_paragraph_property_array(
            &referenced.properties,
            data,
            None,
            active_offsets,
        )?);
        active_offsets.remove(&offset);
        // Both reference SPRMs terminate processing of their containing array
        // when the referenced PrcData is processed.
        break;
    }
    Ok(applied)
}

fn set_document_part_length(fib: &mut Fib, part: FieldDocumentPart, length: u32) -> Result<()> {
    let length = i32::try_from(length)
        .map_err(|_| Error::Limit("document-part character count exceeds i32".into()))?;
    match part {
        FieldDocumentPart::Main => fib.rg_lw.ccp_text = length,
        FieldDocumentPart::Footnote => fib.rg_lw.ccp_footnote = length,
        FieldDocumentPart::Header => fib.rg_lw.ccp_header = length,
        FieldDocumentPart::Comment => fib.rg_lw.ccp_comment = length,
        FieldDocumentPart::Endnote => fib.rg_lw.ccp_endnote = length,
        FieldDocumentPart::Textbox => fib.rg_lw.ccp_textbox = length,
        FieldDocumentPart::HeaderTextbox => fib.rg_lw.ccp_header_textbox = length,
        FieldDocumentPart::Macro => {
            return Err(Error::invalid(
                0,
                "the macro field table has no document-part character count",
            ));
        }
    }
    Ok(())
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct CpReplacement {
    old_start: u32,
    old_end: u32,
    new_end: u32,
}

#[derive(Clone, Copy, Debug, Default)]
struct FormattingRebuild {
    character: bool,
    paragraph: bool,
}

impl CpReplacement {
    fn new(old_start: u32, old_end: u32, replacement_len: u32) -> Result<Self> {
        let new_end = old_start
            .checked_add(replacement_len)
            .ok_or_else(|| Error::Limit("DOC replacement CP limit overflow".into()))?;
        Ok(Self {
            old_start,
            old_end,
            new_end,
        })
    }

    fn relocate_u32(self, value: u32, label: &str) -> Result<u32> {
        if value < self.old_start || (value == self.old_start && self.old_start != self.old_end) {
            return Ok(value);
        }
        if value < self.old_end {
            return Err(Error::invalid(
                u64::from(value),
                format!("{label} falls inside the replaced text range"),
            ));
        }
        if self.new_end >= self.old_end {
            value
                .checked_add(self.new_end - self.old_end)
                .ok_or_else(|| Error::Limit(format!("{label} CP overflow")))
        } else {
            value
                .checked_sub(self.old_end - self.new_end)
                .ok_or_else(|| Error::Limit(format!("{label} CP underflow")))
        }
    }

    fn relocate_i32(self, value: i32, label: &str) -> Result<i32> {
        let value =
            u32::try_from(value).map_err(|_| Error::invalid(0, format!("{label} is negative")))?;
        i32::try_from(self.relocate_u32(value, label)?)
            .map_err(|_| Error::Limit(format!("{label} exceeds i32")))
    }
}

fn relocate_cp_positions(positions: &mut [u32], edit: &CpReplacement, label: &str) -> Result<()> {
    for position in positions {
        *position = edit.relocate_u32(*position, label)?;
    }
    Ok(())
}

fn relocate_bookmark_positions(
    starts: &mut super::BookmarkStartTable,
    ends: &mut super::BookmarkEndTable,
    edit: &CpReplacement,
    label: &str,
) -> Result<()> {
    relocate_cp_positions(&mut starts.positions, edit, label)?;
    relocate_cp_positions(&mut ends.positions, edit, label)
}

fn relocate_field(field: &mut super::Field, edit: &CpReplacement) -> Result<()> {
    field.begin.position = edit.relocate_u32(field.begin.position, "Plcfld begin CP")?;
    for nested in &mut field.instruction_fields {
        relocate_field(nested, edit)?;
    }
    if let Some(separator) = &mut field.separator {
        separator.position = edit.relocate_u32(separator.position, "Plcfld separator CP")?;
    }
    for nested in &mut field.result_fields {
        relocate_field(nested, edit)?;
    }
    field.end.position = edit.relocate_u32(field.end.position, "Plcfld end CP")?;
    Ok(())
}

fn validate_formatting_edit(
    file: &DocFile,
    source_start: u32,
    source_width: usize,
    source_character_count: usize,
    prior_edits: &[CpReplacement],
    edit: &CpReplacement,
) -> Result<FormattingRebuild> {
    let source_len = source_character_count
        .checked_mul(source_width)
        .ok_or_else(|| Error::Limit("text piece source length overflow".into()))?;
    let mut rebuild = FormattingRebuild::default();
    let mut validate_positions = |positions: &[u32], label: &str, kind: &str| -> Result<()> {
        let source_end = u64::from(source_start)
            .checked_add(source_len as u64)
            .ok_or_else(|| Error::Limit("text piece source limit overflow".into()))?;
        for position in positions {
            let position = u64::from(*position);
            if position < u64::from(source_start) || position > source_end {
                continue;
            }
            let byte_offset = usize::try_from(position - u64::from(source_start))
                .map_err(|_| Error::Limit("formatting boundary offset exceeds usize".into()))?;
            if !byte_offset.is_multiple_of(source_width) {
                return Err(Error::invalid(
                    position,
                    format!("{label} is not aligned to the text-piece encoding"),
                ));
            }
            let current_position = relocate_character_position(
                u32::try_from(byte_offset / source_width).map_err(|_| {
                    Error::Limit("formatting boundary character offset exceeds u32".into())
                })?,
                prior_edits,
                label,
            )?;
            if let Err(error) = edit.relocate_u32(current_position, label) {
                match kind {
                    "character" => rebuild.character = true,
                    "paragraph" => rebuild.paragraph = true,
                    _ => return Err(error),
                }
            }
        }
        Ok(())
    };
    validate_positions(
        &file.table.character_bin_table.value.file_positions,
        "PlcBteChpx FC boundary",
        "character",
    )?;
    validate_positions(
        &file.table.paragraph_bin_table.value.file_positions,
        "PlcBtePapx FC boundary",
        "paragraph",
    )?;
    for page in &file.word_document.character_format_pages {
        validate_positions(
            &page.value.file_positions,
            "ChpxFkp FC boundary",
            "character",
        )?;
    }
    for page in &file.word_document.paragraph_format_pages {
        validate_positions(
            &page.value.file_positions,
            "PapxFkp FC boundary",
            "paragraph",
        )?;
    }
    Ok(rebuild)
}

fn relocate_document_properties_cp(
    properties: &mut DocumentProperties,
    part: FieldDocumentPart,
    edit: &CpReplacement,
) -> Result<()> {
    use super::DocumentPropertiesExtension;

    properties.word97.base.document_flags.exact_statistics = false;
    properties.word97.display_flags.list_cache_invalid = true;
    if part == FieldDocumentPart::Main && properties.word97.maximum_list_cache_position >= 0 {
        properties.word97.maximum_list_cache_position = edit.relocate_i32(
            properties.word97.maximum_list_cache_position,
            "DOP cpMaxListCacheMainDoc",
        )?;
    }
    let word2002 = match &mut properties.extension {
        DocumentPropertiesExtension::None | DocumentPropertiesExtension::Word2000(_) => None,
        DocumentPropertiesExtension::Word2002(value)
        | DocumentPropertiesExtension::Compatibility600 {
            word2002: value, ..
        }
        | DocumentPropertiesExtension::Compatibility610 {
            word2002: value, ..
        } => Some(value),
        DocumentPropertiesExtension::Word2003(value)
        | DocumentPropertiesExtension::Word2003WithTrailingByte {
            word2003: value, ..
        } => Some(&mut value.word2002),
        DocumentPropertiesExtension::Word2007(value) => Some(&mut value.word2003.word2002),
        DocumentPropertiesExtension::Word2010(value) => Some(&mut value.word2007.word2003.word2002),
        DocumentPropertiesExtension::Word2013(value) => {
            Some(&mut value.word2010.word2007.word2003.word2002)
        }
    };
    if let Some(value) = word2002 {
        let (position, label) = match part {
            FieldDocumentPart::Main => (
                &mut value.minimum_revision_positions.main,
                "DOP cpMinRMText",
            ),
            FieldDocumentPart::Footnote => (
                &mut value.minimum_revision_positions.footnote,
                "DOP cpMinRMFtn",
            ),
            FieldDocumentPart::Header => (
                &mut value.minimum_revision_positions.header,
                "DOP cpMinRMHdd",
            ),
            FieldDocumentPart::Comment => (
                &mut value.minimum_revision_positions.comment,
                "DOP cpMinRMAtn",
            ),
            FieldDocumentPart::Endnote => (
                &mut value.minimum_revision_positions.endnote,
                "DOP cpMinRMEdn",
            ),
            FieldDocumentPart::Textbox => (
                &mut value.minimum_revision_positions.textbox,
                "DOP cpMinRMTxbx",
            ),
            FieldDocumentPart::HeaderTextbox => (
                &mut value.minimum_revision_positions.header_textbox,
                "DOP cpMinRMHdrTxbx",
            ),
            FieldDocumentPart::Macro => return Ok(()),
        };
        *position = edit.relocate_u32(*position, label)?;
    }
    Ok(())
}

fn relocate_main_document_cps(table: &mut DocTableStream, edit: &CpReplacement) -> Result<()> {
    for position in &mut table.sections.value.character_positions {
        *position = edit.relocate_i32(*position, "PlcfSed CP")?;
    }
    if let Some(fields) = table.fields.get_mut(&FieldDocumentPart::Main) {
        for field in &mut fields.value.fields {
            relocate_field(field, edit)?;
        }
        fields.value.terminal_position =
            edit.relocate_u32(fields.value.terminal_position, "Plcfld terminal CP")?;
    }
    if let Some(bookmarks) = &mut table.bookmarks {
        relocate_bookmark_positions(
            &mut bookmarks.value.starts,
            &mut bookmarks.value.ends,
            edit,
            "main bookmark CP",
        )?;
    }
    if let Some(notes) = &mut table.footnotes {
        relocate_cp_positions(
            &mut notes.references.value.positions,
            edit,
            "footnote reference CP",
        )?;
    }
    if let Some(notes) = &mut table.endnotes {
        relocate_cp_positions(
            &mut notes.references.value.positions,
            edit,
            "endnote reference CP",
        )?;
    }
    if let Some(annotations) = &mut table.annotations {
        relocate_cp_positions(
            &mut annotations.references.value.positions,
            edit,
            "comment reference CP",
        )?;
    }
    if let Some(bookmarks) = &mut table.annotation_bookmarks {
        relocate_bookmark_positions(
            &mut bookmarks.value.starts,
            &mut bookmarks.value.ends,
            edit,
            "annotation bookmark CP",
        )?;
    }
    if let Some(anchors) = table.shape_anchors.get_mut(&TextboxDocumentPart::Main) {
        relocate_cp_positions(&mut anchors.value.positions, edit, "main shape anchor CP")?;
    }
    if let Some(subdocuments) = &mut table.subdocuments {
        relocate_cp_positions(&mut subdocuments.value.positions, edit, "subdocument CP")?;
    }
    if let Some(overrides) = &mut table.list_overrides {
        for value in &mut overrides.value.overrides {
            value.data.first_paragraph_position = edit.relocate_u32(
                value.data.first_paragraph_position,
                "LFO first-paragraph CP",
            )?;
        }
    }
    if let Some(properties) = &mut table.document_properties {
        relocate_document_properties_cp(&mut properties.value, FieldDocumentPart::Main, edit)?;
    }
    if let Some(cache) = &mut table.table_character_cache {
        relocate_cp_positions(&mut cache.value.positions, edit, "table-character cache CP")?;
    }
    if let Some(selection) = &mut table.selection_state {
        selection.value.first_character =
            edit.relocate_i32(selection.value.first_character, "selection cpFirst")?;
        selection.value.character_limit =
            edit.relocate_i32(selection.value.character_limit, "selection cpLim")?;
        selection.value.anchor_character =
            edit.relocate_i32(selection.value.anchor_character, "selection cpAnchor")?;
        if selection.value.flags.block && !selection.value.flags.table {
            selection.value.shrink_anchor_character = edit.relocate_i32(
                selection.value.shrink_anchor_character,
                "selection cpAnchorShrink",
            )?;
        }
    }
    relocate_global_document_cps(table, edit)
}

fn relocate_global_document_cps(table: &mut DocTableStream, edit: &CpReplacement) -> Result<()> {
    if let Some(state) = &mut table.spelling_state {
        relocate_cp_positions(&mut state.value.positions, edit, "spelling state CP")?;
    }
    if let Some(state) = &mut table.grammar_state {
        relocate_cp_positions(&mut state.value.positions, edit, "grammar state CP")?;
    }
    if let Some(state) = &mut table.language_detection_state {
        relocate_cp_positions(
            &mut state.value.positions,
            edit,
            "language-detection state CP",
        )?;
    }
    if let Some(summary) = &mut table.auto_summary_ranges {
        relocate_cp_positions(&mut summary.value.positions, edit, "AutoSummary CP")?;
    }
    if let Some(state) = &mut table.smart_tag_recognizer_state {
        relocate_cp_positions(&mut state.value.positions, edit, "smart-tag state CP")?;
    }
    if let Some(cookies) = &mut table.grammar_checker_cookies {
        relocate_cp_positions(&mut cookies.value.positions, edit, "grammar-cookie CP")?;
    }
    if let Some(cookies) = &mut table.legacy_grammar_checker_cookies {
        relocate_cp_positions(
            &mut cookies.value.positions,
            edit,
            "legacy grammar-cookie CP",
        )?;
    }
    if let Some(bookmarks) = &mut table.structured_tag_bookmarks {
        relocate_bookmark_positions(
            &mut bookmarks.value.starts,
            &mut bookmarks.value.ends,
            edit,
            "structured-tag bookmark CP",
        )?;
    }
    if let Some(bookmarks) = &mut table.range_protection {
        relocate_bookmark_positions(
            &mut bookmarks.value.starts,
            &mut bookmarks.value.ends,
            edit,
            "range-protection bookmark CP",
        )?;
    }
    if let Some(bookmarks) = &mut table.smart_tag_bookmarks {
        relocate_cp_positions(
            &mut bookmarks.value.starts.positions,
            edit,
            "smart-tag bookmark start CP",
        )?;
        relocate_cp_positions(
            &mut bookmarks.value.ends.positions,
            edit,
            "smart-tag bookmark end CP",
        )?;
    }
    if let Some(bookmarks) = &mut table.format_consistency_bookmarks {
        relocate_bookmark_positions(
            &mut bookmarks.value.starts,
            &mut bookmarks.value.ends,
            edit,
            "format-consistency bookmark CP",
        )?;
    }
    if let Some(bookmarks) = &mut table.repair_bookmarks {
        relocate_bookmark_positions(
            &mut bookmarks.value.starts,
            &mut bookmarks.value.ends,
            edit,
            "repair bookmark CP",
        )?;
    }
    if let Some(methods) = &mut table.user_input_methods {
        relocate_cp_positions(&mut methods.value.positions, edit, "user-input-method CP")?;
    }
    Ok(())
}

fn relocate_non_main_document_part_cps(
    table: &mut DocTableStream,
    part: FieldDocumentPart,
    edit: &CpReplacement,
) -> Result<()> {
    if let Some(fields) = table.fields.get_mut(&part) {
        for field in &mut fields.value.fields {
            relocate_field(field, edit)?;
        }
        fields.value.terminal_position =
            edit.relocate_u32(fields.value.terminal_position, "Plcfld terminal CP")?;
    }
    match part {
        FieldDocumentPart::Footnote => {
            if let Some(notes) = &mut table.footnotes {
                relocate_cp_positions(&mut notes.text.value.positions, edit, "footnote text CP")?;
            }
        }
        FieldDocumentPart::Header => {
            if let Some(headers) = &mut table.header_text {
                for boundary in &mut headers.value.boundaries {
                    if let super::HeaderStoryBoundary::Position(position) = boundary {
                        *position = edit.relocate_u32(*position, "header story CP")?;
                    }
                }
            }
            if let Some(anchors) = table.shape_anchors.get_mut(&TextboxDocumentPart::Header) {
                relocate_cp_positions(
                    &mut anchors.value.positions,
                    edit,
                    "header shape anchor CP",
                )?;
            }
        }
        FieldDocumentPart::Comment => {
            if let Some(annotations) = &mut table.annotations {
                relocate_cp_positions(
                    &mut annotations.text.value.positions,
                    edit,
                    "comment text CP",
                )?;
            }
        }
        FieldDocumentPart::Endnote => {
            if let Some(notes) = &mut table.endnotes {
                relocate_cp_positions(&mut notes.text.value.positions, edit, "endnote text CP")?;
            }
        }
        FieldDocumentPart::Textbox => {
            if let Some(stories) = table.textbox_stories.get_mut(&TextboxDocumentPart::Main) {
                relocate_cp_positions(&mut stories.value.positions, edit, "textbox story CP")?;
            }
            if let Some(breaks) = table.textbox_breaks.get_mut(&TextboxDocumentPart::Main) {
                relocate_cp_positions(&mut breaks.value.positions, edit, "textbox break CP")?;
            }
        }
        FieldDocumentPart::HeaderTextbox => {
            if let Some(stories) = table.textbox_stories.get_mut(&TextboxDocumentPart::Header) {
                relocate_cp_positions(
                    &mut stories.value.positions,
                    edit,
                    "header-textbox story CP",
                )?;
            }
            if let Some(breaks) = table.textbox_breaks.get_mut(&TextboxDocumentPart::Header) {
                relocate_cp_positions(
                    &mut breaks.value.positions,
                    edit,
                    "header-textbox break CP",
                )?;
            }
        }
        FieldDocumentPart::Main | FieldDocumentPart::Macro => {
            return Err(Error::invalid(
                0,
                "invalid non-main document-part CP relocation",
            ));
        }
    }
    if let Some(properties) = &mut table.document_properties {
        relocate_document_properties_cp(&mut properties.value, part, edit)?;
    }
    Ok(())
}

fn parse_text_pieces(clx: &Clx, word: &[u8], limits: Limits) -> Result<Vec<DocTextPiece>> {
    if clx.piece_table.character_positions.len() != clx.piece_table.pieces.len() + 1 {
        return Err(Error::invalid(0, "PlcPcd CP/Pcd cardinality mismatch"));
    }
    ensure_entry_limit("DOC text pieces", clx.piece_table.pieces.len(), limits)?;
    clx.piece_table
        .pieces
        .iter()
        .enumerate()
        .map(|(piece_index, descriptor)| {
            Ok(DocTextPiece {
                piece_index,
                value: descriptor.text_piece(
                    word,
                    clx.piece_table.character_positions[piece_index],
                    clx.piece_table.character_positions[piece_index + 1],
                )?,
            })
        })
        .collect()
}

fn parse_object_pool(
    compound: &CompoundFile,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<Option<DocObjectPoolStorage>> {
    let Some(pool) = compound.entry("/ObjectPool") else {
        return Ok(None);
    };
    if !pool.is_storage() {
        let error = Error::invalid(0, "ObjectPool is not a CFB storage");
        if options.is_strict() {
            return Err(error);
        }
        report_object_pool_compatibility(diagnostics, &pool.path, error.to_string());
        return Ok(None);
    }

    let mut objects = Vec::new();
    for entry in compound.children(&pool.path)? {
        if !entry.is_storage() {
            let error = Error::invalid(0, "ObjectPool contains a direct stream");
            if options.is_strict() {
                return Err(error);
            }
            report_object_pool_compatibility(diagnostics, &entry.path, error.to_string());
            continue;
        }
        let descriptor_entry = compound
            .children(&entry.path)?
            .into_iter()
            .find(|child| child.is_stream() && child.name.eq_ignore_ascii_case("\u{3}ObjInfo"));
        let Some(descriptor_entry) = descriptor_entry else {
            let error = Error::invalid(0, "embedded object storage has no ObjInfo stream");
            if options.is_strict() {
                return Err(error);
            }
            report_object_pool_compatibility(diagnostics, &entry.path, error.to_string());
            continue;
        };
        let descriptor = match OleObjectDescriptor::from_bytes(&descriptor_entry.data) {
            Ok(value) => value,
            Err(error) if options.is_strict() => return Err(error),
            Err(error) => {
                report_object_pool_compatibility(
                    diagnostics,
                    &descriptor_entry.path,
                    error.to_string(),
                );
                continue;
            }
        };
        let mut entry_paths = compound
            .entries()
            .iter()
            .filter(|candidate| candidate.path.starts_with(&entry.path))
            .map(|candidate| candidate.path.clone())
            .collect::<Vec<_>>();
        entry_paths.sort();
        objects.push(DocEmbeddedObjectStorage {
            path: entry.path.clone(),
            descriptor_stream_path: descriptor_entry.path.clone(),
            descriptor,
            entry_paths,
        });
    }
    ensure_entry_limit("DOC ObjectPool objects", objects.len(), options.limits)?;
    objects.sort_by(|left, right| left.path.cmp(&right.path));
    Ok(Some(DocObjectPoolStorage {
        path: pool.path.clone(),
        objects,
    }))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DataReferenceKind {
    PictureOrBinary,
    ParagraphProperties,
    Ambiguous,
}

fn rebuild_data_stream(data: Option<&DocDataStream>) -> Result<RebuiltDataStream> {
    let Some(data) = data else {
        return Ok(RebuiltDataStream {
            bytes: None,
            relocations: BTreeMap::new(),
        });
    };
    let mut measured = TableLayout::new(data.physical_bytes.clone());
    for node in &data.nodes {
        measured.replace(
            usize::try_from(node.offset)
                .map_err(|_| Error::Limit("Data node offset exceeds usize".into()))?,
            node.physical_len,
            encode_data_node_value(&node.value)?,
            "Data node",
        )?;
    }
    let (_, measured_relocation) = measured.finish()?;
    let mut relocations = BTreeMap::new();
    for node in &data.nodes {
        let location = FibFcLcb {
            fc: node.offset,
            lcb: u32::try_from(node.physical_len)
                .map_err(|_| Error::Limit("Data node length exceeds u32".into()))?,
        };
        let relocated = measured_relocation
            .relocate(location)?
            .ok_or_else(|| Error::invalid(u64::from(node.offset), "Data node exceeds stream"))?;
        relocations.insert(node.offset, relocated.fc);
    }

    let mut emitted = TableLayout::new(data.physical_bytes.clone());
    for node in &data.nodes {
        let mut value = node.value.clone();
        if let DocDataNodeValue::ParagraphProperties(properties) = &mut value {
            relocate_grpprl_data_references(&mut properties.properties, &relocations)?;
        }
        emitted.replace(
            usize::try_from(node.offset)
                .map_err(|_| Error::Limit("Data node offset exceeds usize".into()))?,
            node.physical_len,
            encode_data_node_value(&value)?,
            "Data node",
        )?;
    }
    let (bytes, emitted_relocation) = emitted.finish()?;
    for node in &data.nodes {
        let location = FibFcLcb {
            fc: node.offset,
            lcb: u32::try_from(node.physical_len)
                .map_err(|_| Error::Limit("Data node length exceeds u32".into()))?,
        };
        let emitted = emitted_relocation
            .relocate(location)?
            .ok_or_else(|| Error::invalid(u64::from(node.offset), "Data node exceeds stream"))?;
        if emitted.fc != relocations[&node.offset] {
            return Err(Error::invalid(
                0,
                "Data layout changed between measure and emit",
            ));
        }
    }
    Ok(RebuiltDataStream {
        bytes: Some(bytes),
        relocations,
    })
}

fn encode_data_node_value(value: &DocDataNodeValue) -> Result<Vec<u8>> {
    match value {
        DocDataNodeValue::Picture(value) => {
            let mut value = value.clone();
            let picture_len = value.picture.to_bytes()?.len();
            let name_len = value
                .shape_file_name
                .as_ref()
                .map_or(0usize, |name| name.len() + 1);
            let total_len = Picf::ENCODED_LEN
                .checked_add(name_len)
                .and_then(|length| length.checked_add(picture_len))
                .ok_or_else(|| Error::Limit("PICFAndOfficeArtData length overflow".into()))?;
            value.picf.total_length = i32::try_from(total_len)
                .map_err(|_| Error::Limit("PICFAndOfficeArtData length exceeds i32".into()))?;
            value.to_bytes()
        }
        DocDataNodeValue::Binary(value) => {
            let mut value = value.clone();
            let total_len = NilPicfAndBinData::HEADER_LEN
                .checked_add(value.binary_len()?)
                .ok_or_else(|| Error::Limit("NilPICFAndBinData length overflow".into()))?;
            value.total_length = i32::try_from(total_len)
                .map_err(|_| Error::Limit("NilPICFAndBinData length exceeds i32".into()))?;
            value.to_bytes()
        }
        DocDataNodeValue::ParagraphProperties(value) => value.to_bytes(),
    }
}

fn relocate_root_data_references(
    clx: &mut DocLocated<Clx>,
    character_pages: &mut [DocFkpPage<ChpxFkp>],
    paragraph_pages: &mut [DocFkpPage<PapxFkp>],
    sections: &mut [DocSectionProperties],
    styles: Option<&mut DocLocated<StyleSheet>>,
    lists: Option<&mut DocListDefinitions>,
    relocations: &BTreeMap<u32, u32>,
) -> Result<()> {
    if relocations.is_empty() {
        return Ok(());
    }
    for property_run in &mut clx.value.property_runs {
        relocate_grpprl_data_references(&mut property_run.properties, relocations)?;
    }
    for page in character_pages {
        for run in &mut page.value.runs {
            if let Some(properties) = &mut run.properties {
                relocate_grpprl_data_references(properties, relocations)?;
            }
        }
    }
    for page in paragraph_pages {
        for run in &mut page.value.runs {
            if let Some(properties) = &mut run.properties {
                relocate_grpprl_data_references(&mut properties.properties, relocations)?;
            }
        }
    }
    for section in sections {
        if let Some(value) = &mut section.value {
            relocate_grpprl_data_references(&mut value.properties, relocations)?;
        }
    }
    if let Some(styles) = styles {
        if let Some(properties) = &mut styles.value.info.standard_character_properties {
            relocate_grpprl_data_references(properties, relocations)?;
        }
        if let Some(properties) = &mut styles.value.info.standard_paragraph_properties {
            relocate_grpprl_data_references(properties, relocations)?;
        }
        for style in &mut styles.value.styles {
            if let Some(definition) = &mut style.definition {
                relocate_style_data_references(&mut definition.formatting, relocations)?;
            }
        }
    }
    if let Some(lists) = lists {
        for definition in &mut lists.value.definitions {
            for level in &mut definition.levels {
                relocate_grpprl_data_references(&mut level.paragraph_properties, relocations)?;
                relocate_grpprl_data_references(&mut level.number_properties, relocations)?;
            }
        }
    }
    Ok(())
}

fn relocate_style_data_references(
    formatting: &mut StyleFormatting,
    relocations: &BTreeMap<u32, u32>,
) -> Result<()> {
    let visit = |properties: &mut GrpPrl| relocate_grpprl_data_references(properties, relocations);
    match formatting {
        StyleFormatting::Paragraph {
            paragraph,
            character,
        } => {
            visit(&mut paragraph.properties)?;
            visit(&mut character.properties)?;
        }
        StyleFormatting::Character { character } => visit(&mut character.properties)?,
        StyleFormatting::RevisionParagraph {
            paragraph,
            character,
            original_paragraph,
            original_character,
            ..
        } => {
            visit(&mut paragraph.properties)?;
            visit(&mut character.properties)?;
            visit(&mut original_paragraph.properties)?;
            visit(&mut original_character.properties)?;
        }
        StyleFormatting::RevisionCharacter {
            character,
            original_character,
            ..
        } => {
            visit(&mut character.properties)?;
            visit(&mut original_character.properties)?;
        }
        StyleFormatting::Table {
            table,
            paragraph,
            character,
        } => {
            visit(&mut table.properties)?;
            visit(&mut paragraph.properties)?;
            visit(&mut character.properties)?;
        }
        StyleFormatting::Numbering { paragraph } => visit(&mut paragraph.properties)?,
    }
    Ok(())
}

fn relocate_grpprl_data_references(
    properties: &mut GrpPrl,
    relocations: &BTreeMap<u32, u32>,
) -> Result<()> {
    let special_character = properties.properties.iter().rev().find_map(|property| {
        (property.sprm.opcode().ok() == Some(0x0855)).then_some(match property.operand {
            SprmOperand::Toggle(value) => value & 1 != 0,
            _ => false,
        })
    });
    let ole_object = properties.properties.iter().rev().find_map(|property| {
        (property.sprm.opcode().ok() == Some(0x080a)).then_some(match property.operand {
            SprmOperand::Toggle(value) => value & 1 != 0,
            _ => false,
        })
    });
    for property in &mut properties.properties {
        let opcode = property.sprm.opcode()?;
        let is_reference = matches!(opcode, 0x646b | 0x6646)
            || (opcode == 0x6a03 && special_character == Some(true) && ole_object != Some(true));
        if is_reference {
            let SprmOperand::Dword(raw_offset) = &mut property.operand else {
                return Err(Error::invalid(
                    0,
                    "Data-reference SPRM operand is not a dword",
                ));
            };
            let offset = u32::from_le_bytes(*raw_offset);
            if let Some(relocated) = relocations.get(&offset) {
                *raw_offset = relocated.to_le_bytes();
            }
        }
        if let SprmOperand::CharacterMajority(nested) = &mut property.operand {
            relocate_grpprl_data_references(nested, relocations)?;
        }
    }
    Ok(())
}

#[derive(Clone, Copy, Debug, Default)]
struct CharacterDataReferenceState {
    picture_offset: Option<u32>,
    ole_object: Option<bool>,
}

impl CharacterDataReferenceState {
    fn apply(&mut self, properties: &GrpPrl) {
        for property in &properties.properties {
            match (property.sprm.opcode().ok(), &property.operand) {
                (Some(0x6a03), SprmOperand::Dword(bytes)) => {
                    self.picture_offset = Some(u32::from_le_bytes(*bytes));
                }
                (Some(0x080a), SprmOperand::Toggle(value)) => {
                    self.ole_object = Some(value & 1 != 0);
                }
                (_, SprmOperand::CharacterMajority(nested)) => self.apply(nested),
                _ => {}
            }
        }
    }
}

fn collect_nil_picf_field_types(
    word: &DocWordDocumentStream,
    table: &DocTableStream,
) -> BTreeMap<u32, Option<NilPicfFieldType>> {
    let mut references = BTreeMap::new();
    let mut picture_characters = Vec::new();
    for (descriptor, piece) in table
        .clx
        .value
        .piece_table
        .pieces
        .iter()
        .zip(&word.text_pieces)
    {
        let mut base = CharacterDataReferenceState::default();
        if let Prm::Complex { property_run_index } = descriptor.property_modifier
            && let Some(properties) = table
                .clx
                .value
                .property_runs
                .get(usize::from(property_run_index))
        {
            base.apply(&properties.properties);
        }
        picture_characters.extend(
            picture_characters_in_piece(&piece.value)
                .into_iter()
                .map(|(cp, fc)| (cp, fc, base)),
        );
    }
    for page in &word.character_format_pages {
        for (index, run) in page.value.runs.iter().enumerate() {
            let Some((&fc_start, &fc_end)) = page
                .value
                .file_positions
                .get(index)
                .zip(page.value.file_positions.get(index + 1))
            else {
                continue;
            };
            for &(cp, fc, base) in &picture_characters {
                if !(fc_start..fc_end).contains(&fc) {
                    continue;
                }
                let mut state = base;
                if let Some(properties) = &run.properties {
                    state.apply(properties);
                }
                let Some(offset) = state
                    .picture_offset
                    .filter(|_| state.ole_object != Some(true))
                else {
                    continue;
                };
                let Some(field_type) = nil_picf_field_type_at_cp(&word.fib, table, cp)
                    .or_else(|| private_field_type_at_cp(word, cp))
                else {
                    continue;
                };
                references
                    .entry(offset)
                    .and_modify(|existing| {
                        if *existing != Some(field_type) {
                            *existing = None;
                        }
                    })
                    .or_insert(Some(field_type));
            }
        }
    }
    references
}

fn picture_characters_in_piece(piece: &TextPiece) -> Vec<(u32, u32)> {
    let (width, characters): (u64, Box<dyn Iterator<Item = bool> + '_>) = match &piece.characters {
        TextPieceCharacters::Compressed(values) => {
            (1, Box::new(values.iter().map(|value| *value == 1)))
        }
        TextPieceCharacters::Utf16(values) => (2, Box::new(values.iter().map(|value| *value == 1))),
    };
    characters
        .enumerate()
        .filter_map(|(index, is_picture)| {
            if !is_picture {
                return None;
            }
            let fc = u64::from(piece.file_offset).checked_add(index as u64 * width)?;
            let cp = i64::from(piece.cp_start).checked_add(index as i64)?;
            Some((u32::try_from(cp).ok()?, u32::try_from(fc).ok()?))
        })
        .collect()
}

fn nil_picf_field_type_at_cp(
    fib: &Fib,
    table: &DocTableStream,
    global_cp: u32,
) -> Option<NilPicfFieldType> {
    let counts = [
        (FieldDocumentPart::Main, fib.rg_lw.ccp_text),
        (FieldDocumentPart::Footnote, fib.rg_lw.ccp_footnote),
        (FieldDocumentPart::Header, fib.rg_lw.ccp_header),
        (FieldDocumentPart::Comment, fib.rg_lw.ccp_comment),
        (FieldDocumentPart::Endnote, fib.rg_lw.ccp_endnote),
        (FieldDocumentPart::Textbox, fib.rg_lw.ccp_textbox),
        (
            FieldDocumentPart::HeaderTextbox,
            fib.rg_lw.ccp_header_textbox,
        ),
    ];
    let mut start = 0u32;
    for (part, count) in counts {
        let count = u32::try_from(count).ok()?;
        let end = start.checked_add(count)?;
        if (start..end).contains(&global_cp) {
            let relative_cp = global_cp - start;
            let field = table.fields.get(&part)?.value.innermost_at(relative_cp)?;
            return NilPicfFieldType::from_field_type(field.begin.field_type);
        }
        start = end;
    }
    None
}

fn private_field_type_at_cp(
    word: &DocWordDocumentStream,
    picture_cp: u32,
) -> Option<NilPicfFieldType> {
    let mut nesting = 0usize;
    let mut begin = None;
    for cp in (0..picture_cp).rev() {
        match text_character_at_cp(word, cp)? {
            0x0015 => nesting = nesting.saturating_add(1),
            0x0013 if nesting == 0 => {
                begin = Some(cp);
                break;
            }
            0x0013 => nesting = nesting.saturating_sub(1),
            _ => {}
        }
    }
    let begin = begin?;
    let mut instruction = String::new();
    let mut nested = 0usize;
    for cp in begin + 1..picture_cp {
        match text_character_at_cp(word, cp)? {
            0x0013 => nested = nested.saturating_add(1),
            0x0015 if nested != 0 => nested = nested.saturating_sub(1),
            0x0014 | 0x0015 if nested == 0 => break,
            value if nested == 0 && value <= 0x007f => {
                instruction.push(char::from(value as u8));
            }
            _ => {}
        }
    }
    instruction
        .trim_start()
        .to_ascii_uppercase()
        .starts_with("PRIVATE")
        .then_some(NilPicfFieldType::Private(PrivateFieldType::Private))
}

fn text_character_at_cp(word: &DocWordDocumentStream, cp: u32) -> Option<u16> {
    let cp = i32::try_from(cp).ok()?;
    let piece = word
        .text_pieces
        .iter()
        .find(|piece| piece.value.cp_start <= cp && cp < piece.value.cp_end)?;
    let index = usize::try_from(cp - piece.value.cp_start).ok()?;
    match &piece.value.characters {
        TextPieceCharacters::Compressed(values) => values.get(index).copied().map(u16::from),
        TextPieceCharacters::Utf16(values) => values.get(index).copied(),
    }
}

fn parse_data_stream(
    bytes: Option<&[u8]>,
    word: &DocWordDocumentStream,
    table: &DocTableStream,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<Option<DocDataStream>> {
    let mut references = collect_data_references(word, table)?;
    let binary_field_types = collect_nil_picf_field_types(word, table);
    let Some(bytes) = bytes else {
        if references.is_empty() {
            return Ok(None);
        }
        let error = Error::invalid(0, "SPRM references exist but the Data stream is missing");
        if options.is_strict() {
            return Err(error);
        }
        report_data_compatibility(diagnostics, 0, error.to_string());
        return Ok(None);
    };
    ensure_stream_limit("Data", bytes, options.limits)?;

    let mut pending = references
        .iter()
        .map(|(offset, kind)| (*offset, *kind))
        .collect::<Vec<_>>();
    let mut index = 0usize;
    let mut nodes = Vec::new();
    while index < pending.len() {
        let (offset, kind) = pending[index];
        index += 1;
        let parsed = parse_data_node(bytes, offset, kind);
        let mut node = match parsed {
            Ok(value) => value,
            Err(error) if options.is_strict() => return Err(error),
            Err(error) => {
                report_data_compatibility(diagnostics, u64::from(offset), error.to_string());
                continue;
            }
        };
        if let DocDataNodeValue::Binary(value) = &mut node.value {
            match binary_field_types.get(&offset) {
                Some(Some(field_type)) => value.interpret(*field_type),
                Some(None) | None if options.is_strict() => {
                    return Err(Error::invalid(
                        u64::from(offset),
                        "NilPICF picture character is not inside one permitted field type",
                    ));
                }
                Some(None) | None => {
                    value.mark_invalid_context();
                    report_data_compatibility(
                        diagnostics,
                        u64::from(offset),
                        "NilPICF picture character is not inside one permitted field type".into(),
                    );
                }
            }
        }
        if let DocDataNodeValue::ParagraphProperties(value) = &node.value {
            let mut nested = BTreeMap::new();
            collect_grpprl_data_references(&value.properties, &mut nested)?;
            for (nested_offset, nested_kind) in nested {
                match references.get(&nested_offset) {
                    Some(existing) if *existing != nested_kind => {
                        references.insert(nested_offset, DataReferenceKind::Ambiguous);
                        if let Some(pending_entry) = pending
                            .iter_mut()
                            .find(|(pending_offset, _)| *pending_offset == nested_offset)
                        {
                            pending_entry.1 = DataReferenceKind::Ambiguous;
                        }
                    }
                    Some(_) => {}
                    None => {
                        references.insert(nested_offset, nested_kind);
                        pending.push((nested_offset, nested_kind));
                    }
                }
            }
        }
        nodes.push(node);
    }
    nodes.sort_by_key(|node| node.offset);
    let mut retained = Vec::<DocDataNode>::with_capacity(nodes.len());
    for node in nodes {
        let overlaps = retained.last().is_some_and(|previous| {
            u64::from(previous.offset) + previous.physical_len as u64 > u64::from(node.offset)
        });
        if overlaps {
            let error = Error::invalid(
                u64::from(node.offset),
                "referenced Data stream structures overlap",
            );
            if options.is_strict() {
                return Err(error);
            }
            report_data_compatibility(diagnostics, u64::from(node.offset), error.to_string());
        } else {
            retained.push(node);
        }
    }
    ensure_entry_limit("DOC Data nodes", retained.len(), options.limits)?;
    Ok(Some(DocDataStream {
        physical_bytes: bytes.to_vec(),
        nodes: retained,
    }))
}

fn parse_data_node(bytes: &[u8], offset: u32, kind: DataReferenceKind) -> Result<DocDataNode> {
    if kind == DataReferenceKind::Ambiguous {
        let picture = parse_data_node(bytes, offset, DataReferenceKind::PictureOrBinary);
        let paragraph = parse_data_node(bytes, offset, DataReferenceKind::ParagraphProperties);
        return match (picture, paragraph) {
            (Ok(value), Err(_)) | (Err(_), Ok(value)) => Ok(value),
            (Ok(_), Ok(_)) => Err(Error::invalid(
                u64::from(offset),
                "Data offset is structurally ambiguous between picture and PrcData",
            )),
            (Err(picture_error), Err(paragraph_error)) => Err(Error::invalid(
                u64::from(offset),
                format!(
                    "Data offset matches neither picture nor PrcData ({picture_error}; {paragraph_error})"
                ),
            )),
        };
    }
    let start = usize::try_from(offset)
        .map_err(|_| Error::Limit("Data node offset exceeds usize".into()))?;
    let remaining = bytes
        .get(start..)
        .ok_or_else(|| Error::invalid(u64::from(offset), "Data node offset exceeds stream"))?;
    let (physical_len, value) = match kind {
        DataReferenceKind::PictureOrBinary => {
            let length_bytes = remaining.get(..4).ok_or_else(|| {
                Error::invalid(u64::from(offset), "Data picture length is missing")
            })?;
            let length = i32::from_le_bytes(length_bytes.try_into().expect("four bytes checked"));
            let length = usize::try_from(length).map_err(|_| {
                Error::invalid(u64::from(offset), "Data picture length is negative")
            })?;
            let encoded = remaining.get(..length).ok_or_else(|| {
                Error::invalid(u64::from(offset), "Data picture exceeds the stream")
            })?;
            let mapping_mode = encoded
                .get(6..8)
                .map(|value| i16::from_le_bytes(value.try_into().expect("two bytes checked")));
            let value = if matches!(mapping_mode, Some(0x0064 | 0x0066)) {
                DocDataNodeValue::Picture(PicfAndOfficeArtData::from_bytes(encoded)?)
            } else {
                DocDataNodeValue::Binary(Box::new(NilPicfAndBinData::from_bytes(encoded)?))
            };
            (length, value)
        }
        DataReferenceKind::ParagraphProperties => {
            let length_bytes = remaining
                .get(..2)
                .ok_or_else(|| Error::invalid(u64::from(offset), "PrcData length is missing"))?;
            let length = i16::from_le_bytes(length_bytes.try_into().expect("two bytes checked"));
            if !(0..=0x3fa2).contains(&length) {
                return Err(Error::invalid(
                    u64::from(offset),
                    "PrcData cbGrpprl is outside 0..=0x3FA2",
                ));
            }
            let physical_len = length as usize + 2;
            let encoded = remaining.get(..physical_len).ok_or_else(|| {
                Error::invalid(u64::from(offset), "PrcData exceeds the Data stream")
            })?;
            (
                physical_len,
                DocDataNodeValue::ParagraphProperties(PrcData::from_bytes(encoded)?),
            )
        }
        DataReferenceKind::Ambiguous => unreachable!("handled above"),
    };
    Ok(DocDataNode {
        offset,
        physical_len,
        value,
    })
}

fn collect_data_references(
    word: &DocWordDocumentStream,
    table: &DocTableStream,
) -> Result<BTreeMap<u32, DataReferenceKind>> {
    let mut references = BTreeMap::new();
    for property_run in &table.clx.value.property_runs {
        collect_grpprl_data_references(&property_run.properties, &mut references)?;
    }
    for page in &word.character_format_pages {
        for run in &page.value.runs {
            if let Some(properties) = &run.properties {
                collect_grpprl_data_references(properties, &mut references)?;
            }
        }
    }
    for page in &word.paragraph_format_pages {
        for run in &page.value.runs {
            if let Some(properties) = &run.properties {
                collect_grpprl_data_references(&properties.properties, &mut references)?;
            }
        }
    }
    for section in &word.section_properties {
        if let Some(value) = &section.value {
            collect_grpprl_data_references(&value.properties, &mut references)?;
        }
    }
    if let Some(styles) = &table.styles {
        if let Some(properties) = &styles.value.info.standard_character_properties {
            collect_grpprl_data_references(properties, &mut references)?;
        }
        if let Some(properties) = &styles.value.info.standard_paragraph_properties {
            collect_grpprl_data_references(properties, &mut references)?;
        }
        for style in &styles.value.styles {
            if let Some(definition) = &style.definition {
                collect_style_data_references(&definition.formatting, &mut references)?;
            }
        }
    }
    if let Some(lists) = &table.list_definitions {
        for definition in &lists.value.definitions {
            for level in &definition.levels {
                collect_grpprl_data_references(&level.paragraph_properties, &mut references)?;
                collect_grpprl_data_references(&level.number_properties, &mut references)?;
            }
        }
    }
    Ok(references)
}

fn collect_style_data_references(
    formatting: &StyleFormatting,
    references: &mut BTreeMap<u32, DataReferenceKind>,
) -> Result<()> {
    let mut visit = |properties: &GrpPrl| collect_grpprl_data_references(properties, references);
    match formatting {
        StyleFormatting::Paragraph {
            paragraph,
            character,
        } => {
            visit(&paragraph.properties)?;
            visit(&character.properties)?;
        }
        StyleFormatting::Character { character } => visit(&character.properties)?,
        StyleFormatting::RevisionParagraph {
            paragraph,
            character,
            original_paragraph,
            original_character,
            ..
        } => {
            visit(&paragraph.properties)?;
            visit(&character.properties)?;
            visit(&original_paragraph.properties)?;
            visit(&original_character.properties)?;
        }
        StyleFormatting::RevisionCharacter {
            character,
            original_character,
            ..
        } => {
            visit(&character.properties)?;
            visit(&original_character.properties)?;
        }
        StyleFormatting::Table {
            table,
            paragraph,
            character,
        } => {
            visit(&table.properties)?;
            visit(&paragraph.properties)?;
            visit(&character.properties)?;
        }
        StyleFormatting::Numbering { paragraph } => visit(&paragraph.properties)?,
    }
    Ok(())
}

fn collect_grpprl_data_references(
    properties: &GrpPrl,
    references: &mut BTreeMap<u32, DataReferenceKind>,
) -> Result<()> {
    let special_character = properties.properties.iter().rev().find_map(|property| {
        (property.sprm.opcode().ok() == Some(0x0855)).then_some(match property.operand {
            SprmOperand::Toggle(value) => value & 1 != 0,
            _ => false,
        })
    });
    let ole_object = properties.properties.iter().rev().find_map(|property| {
        (property.sprm.opcode().ok() == Some(0x080a)).then_some(match property.operand {
            SprmOperand::Toggle(value) => value & 1 != 0,
            _ => false,
        })
    });
    for property in &properties.properties {
        let opcode = property.sprm.opcode()?;
        let kind = match opcode {
            0x6a03 if special_character == Some(true) && ole_object != Some(true) => {
                Some(DataReferenceKind::PictureOrBinary)
            }
            0x646b | 0x6646 => Some(DataReferenceKind::ParagraphProperties),
            _ => None,
        };
        if let Some(kind) = kind {
            let SprmOperand::Dword(raw_offset) = property.operand else {
                return Err(Error::invalid(
                    0,
                    "Data-reference SPRM operand is not a dword",
                ));
            };
            let offset = u32::from_le_bytes(raw_offset);
            references
                .entry(offset)
                .and_modify(|existing| {
                    if *existing != kind {
                        *existing = DataReferenceKind::Ambiguous;
                    }
                })
                .or_insert(kind);
        }
        if let SprmOperand::CharacterMajority(nested) = &property.operand {
            collect_grpprl_data_references(nested, references)?;
        }
    }
    Ok(())
}

fn report_data_compatibility(diagnostics: &mut Vec<ParseDiagnostic>, offset: u64, message: String) {
    diagnostics.push(ParseDiagnostic::warning(
        ParseDiagnosticCode::InvalidReference,
        BinaryFormat::Doc,
        Some(DATA_STREAM),
        Some(offset),
        "Data Stream",
        SpecificationReference {
            document: "MS-DOC",
            section: "2.1.3",
        },
        message,
    ));
}

fn report_object_pool_compatibility(
    diagnostics: &mut Vec<ParseDiagnostic>,
    path: &Path,
    message: String,
) {
    let path = path.display().to_string();
    diagnostics.push(ParseDiagnostic::warning(
        ParseDiagnosticCode::InvalidStreamPreserved,
        BinaryFormat::Doc,
        Some(&path),
        None,
        "ObjectPool Storage",
        SpecificationReference {
            document: "MS-DOC",
            section: "2.1.4",
        },
        message,
    ));
}

fn validate_object_pool_links(
    compound: &CompoundFile,
    object_pool: Option<&DocObjectPoolStorage>,
) -> Result<()> {
    let physical_pool = compound
        .entry("/ObjectPool")
        .filter(|entry| entry.is_storage());
    if physical_pool.map(|entry| &entry.path) != object_pool.map(|pool| &pool.path) {
        return Err(Error::invalid(0, "DOC ObjectPool root link changed"));
    }
    let Some(object_pool) = object_pool else {
        return Ok(());
    };

    let mut physical_objects = Vec::new();
    for entry in compound.children(&object_pool.path)? {
        if !entry.is_storage() {
            continue;
        }
        let Some(descriptor_entry) = compound.children(&entry.path)?.into_iter().find(|child| {
            child.is_stream()
                && child.name.eq_ignore_ascii_case("\u{3}ObjInfo")
                && OleObjectDescriptor::from_bytes(&child.data).is_ok()
        }) else {
            continue;
        };
        physical_objects.push((entry.path.clone(), descriptor_entry.path.clone()));
    }
    physical_objects.sort();
    let managed_objects = object_pool
        .objects
        .iter()
        .map(|object| (object.path.clone(), object.descriptor_stream_path.clone()))
        .collect::<Vec<_>>();
    if physical_objects != managed_objects {
        return Err(Error::invalid(0, "DOC ObjectPool object links changed"));
    }
    for object in &object_pool.objects {
        let mut physical_paths = compound
            .entries()
            .iter()
            .filter(|entry| entry.path.starts_with(&object.path))
            .map(|entry| entry.path.clone())
            .collect::<Vec<_>>();
        physical_paths.sort();
        if physical_paths != object.entry_paths {
            return Err(Error::invalid(
                0,
                "DOC ObjectPool embedded entry links changed",
            ));
        }
    }
    Ok(())
}

fn validate_data_links(
    word: &DocWordDocumentStream,
    table: &DocTableStream,
    data: Option<&DocDataStream>,
) -> Result<()> {
    let Some(data) = data else {
        return Ok(());
    };
    let mut diagnostics = Vec::new();
    let expected = parse_data_stream(
        Some(&data.physical_bytes),
        word,
        table,
        ParseOptions::compatible(Limits::default()),
        &mut diagnostics,
    )?
    .expect("a supplied Data stream produces a root");
    if expected.nodes.len() != data.nodes.len()
        || expected
            .nodes
            .iter()
            .zip(&data.nodes)
            .any(|(expected, actual)| {
                expected.offset != actual.offset
                    || expected.physical_len != actual.physical_len
                    || !same_data_node_kind(&expected.value, &actual.value)
            })
    {
        return Err(Error::invalid(0, "DOC SPRM/Data node links changed"));
    }
    Ok(())
}

fn same_data_node_kind(left: &DocDataNodeValue, right: &DocDataNodeValue) -> bool {
    matches!(
        (left, right),
        (DocDataNodeValue::Picture(_), DocDataNodeValue::Picture(_))
            | (DocDataNodeValue::Binary(_), DocDataNodeValue::Binary(_))
            | (
                DocDataNodeValue::ParagraphProperties(_),
                DocDataNodeValue::ParagraphProperties(_)
            )
    )
}

fn parse_fkp_pages<T>(
    table: &PlcBte,
    word: &[u8],
    parse: impl Fn(&[u8]) -> Result<T>,
) -> Result<Vec<DocFkpPage<T>>> {
    table
        .pages
        .iter()
        .copied()
        .map(|page| {
            let offset = page.byte_offset()?;
            let bytes = word
                .get(offset..offset + 512)
                .ok_or_else(|| Error::invalid(offset as u64, "FKP page exceeds WordDocument"))?;
            Ok(DocFkpPage {
                page,
                value: parse(bytes)?,
            })
        })
        .collect()
}

fn validate_fkp_page_order<T>(
    pages: &[DocFkpPage<T>],
    positions: impl Fn(&T) -> &[u32],
    structure: &'static str,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<()> {
    for page in pages {
        let Some((index, pair)) = positions(&page.value)
            .windows(2)
            .enumerate()
            .find(|(_, pair)| pair[0] >= pair[1])
        else {
            continue;
        };
        let offset = page
            .page
            .byte_offset()?
            .checked_add((index + 1) * 4)
            .ok_or_else(|| Error::Limit("FKP boundary offset overflow".into()))?;
        let error = Error::invalid(
            offset as u64,
            format!(
                "{structure} values are not strictly increasing: {} then {}",
                pair[0], pair[1]
            ),
        );
        if options.is_strict() {
            return Err(error);
        }
        diagnostics.push(ParseDiagnostic::warning(
            ParseDiagnosticCode::NonconformingRecord,
            BinaryFormat::Doc,
            Some("WordDocument Stream"),
            Some(offset as u64),
            structure,
            SpecificationReference {
                document: "MS-DOC",
                section: "2.9",
            },
            error.to_string(),
        ));
    }
    Ok(())
}

fn parse_bookmarks(fib: &Fib, table: &[u8]) -> Result<Option<DocLocatedBookmarks>> {
    let Some((names, starts, ends)) = fib.bookmark_locations() else {
        return Ok(None);
    };
    if names.lcb == 0 && starts.lcb == 0 && ends.lcb == 0 {
        return Ok(None);
    }
    if names.lcb == 0 || starts.lcb == 0 || ends.lcb == 0 {
        return Err(Error::invalid(0, "bookmark table locations are incomplete"));
    }
    Ok(Some(DocLocatedBookmarks {
        names_location: names,
        starts_location: starts,
        ends_location: ends,
        value: Bookmarks::from_bytes(
            bounded_slice(table, names, "SttbfBkmk")?,
            bounded_slice(table, starts, "PlcfBkf")?,
            bounded_slice(table, ends, "PlcfBkl")?,
        )?,
    }))
}

fn parse_note_tables(
    table: &[u8],
    locations: Option<(FibFcLcb, FibFcLcb)>,
    label: &'static str,
) -> Result<Option<DocNoteTables>> {
    let Some((references, text)) = complete_locations(locations, label)? else {
        return Ok(None);
    };
    Ok(Some(DocNoteTables {
        references: DocLocated {
            location: references,
            value: NoteReferenceTable::from_bytes(bounded_slice(
                table,
                references,
                &format!("{label} reference table"),
            )?)?,
        },
        text: DocLocated {
            location: text,
            value: CpOnlyTable::from_bytes(bounded_slice(
                table,
                text,
                &format!("{label} text table"),
            )?)?,
        },
    }))
}

fn parse_annotation_tables(
    table: &[u8],
    locations: Option<(FibFcLcb, FibFcLcb)>,
) -> Result<Option<DocAnnotationTables>> {
    let Some((references, text)) = complete_locations(locations, "annotation")? else {
        return Ok(None);
    };
    Ok(Some(DocAnnotationTables {
        references: DocLocated {
            location: references,
            value: AnnotationReferenceTable::from_bytes(bounded_slice(
                table,
                references,
                "annotation reference table",
            )?)?,
        },
        text: DocLocated {
            location: text,
            value: CpOnlyTable::from_bytes(bounded_slice(table, text, "annotation text table")?)?,
        },
    }))
}

fn parse_annotation_bookmarks(
    fib: &Fib,
    table: &[u8],
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<Option<DocLocatedAnnotationBookmarks>> {
    let Some((infos, starts, ends)) = fib.annotation_bookmark_locations() else {
        return Ok(None);
    };
    let [infos, starts, ends] =
        match complete_location_array([infos, starts, ends], "annotation bookmark tables") {
            Ok(Some(locations)) => locations,
            Ok(None) => return Ok(None),
            Err(error) if options.is_strict() => return Err(error),
            Err(error) => {
                report_doc_compatibility(
                    diagnostics,
                    ParseDiagnosticCode::InvalidReference,
                    0,
                    "annotation bookmark tables",
                    format!("preserved incomplete annotation bookmark tables: {error}"),
                );
                return Ok(None);
            }
        };
    Ok(Some(DocLocatedAnnotationBookmarks {
        infos_location: infos,
        starts_location: starts,
        ends_location: ends,
        value: AnnotationBookmarks::from_bytes(
            bounded_slice(table, infos, "SttbfAtnBkmk")?,
            bounded_slice(table, starts, "PlcfAtnBkf")?,
            bounded_slice(table, ends, "PlcfAtnBkl")?,
        )?,
    }))
}

fn parse_optional_compatible<T>(
    bytes: &[u8],
    location: Option<FibFcLcb>,
    label: &'static str,
    parse: impl Fn(&[u8]) -> Result<T>,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
    compatibility_tables: &mut Vec<DocCompatibilityTable>,
) -> Result<Option<DocLocated<T>>> {
    match parse_optional(bytes, location, label, parse) {
        Ok(value) => Ok(value),
        Err(error) if options.is_strict() => Err(error),
        Err(error) => {
            let reason = error.to_string();
            report_doc_compatibility(
                diagnostics,
                ParseDiagnosticCode::InvalidReference,
                location.map_or(0, |value| u64::from(value.fc)),
                label,
                format!("preserved an invalid {label} reference: {reason}"),
            );
            if let Some(location) = location.filter(|value| value.lcb != 0) {
                compatibility_tables.push(DocCompatibilityTable {
                    label: label.to_owned(),
                    location,
                    physical_bytes: bounded_slice(bytes, location, label)
                        .ok()
                        .map(<[u8]>::to_vec),
                    reason,
                });
            }
            Ok(None)
        }
    }
}

fn parse_list_definitions(
    table: &[u8],
    location: Option<FibFcLcb>,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
    compatibility_tables: &mut Vec<DocCompatibilityTable>,
) -> Result<Option<DocListDefinitions>> {
    let Some(location) = location.filter(|value| value.lcb != 0) else {
        return Ok(None);
    };
    match ListDefinitions::from_table_stream(table, location) {
        Ok(value) => {
            let (base, levels) = value.to_bytes()?;
            if base.len()
                != usize::try_from(location.lcb)
                    .map_err(|_| Error::Limit("PlfLst length exceeds usize".into()))?
            {
                return Err(Error::invalid(
                    u64::from(location.fc),
                    "PlfLst static base does not match its FIB length",
                ));
            }
            Ok(Some(DocListDefinitions {
                location,
                value,
                trailing_levels_len: levels.len(),
            }))
        }
        Err(error) if options.is_strict() => Err(error),
        Err(error) => {
            let reason = error.to_string();
            report_doc_compatibility(
                diagnostics,
                ParseDiagnosticCode::NonconformingRecord,
                u64::from(location.fc),
                "PlfLst",
                format!("preserved an invalid PlfLst: {reason}"),
            );
            compatibility_tables.push(DocCompatibilityTable {
                label: "PlfLst".to_owned(),
                location,
                physical_bytes: bounded_slice(table, location, "PlfLst")
                    .ok()
                    .map(<[u8]>::to_vec),
                reason,
            });
            Ok(None)
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn parse_bookmark_set<T>(
    table: &[u8],
    locations: Option<[FibFcLcb; 3]>,
    labels: [&'static str; 3],
    structure: &'static str,
    parse: impl Fn(&[u8], &[u8], &[u8]) -> Result<T>,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
    compatibility_tables: &mut Vec<DocCompatibilityTable>,
) -> Result<Option<DocBookmarkSet<T>>> {
    let Some(locations) = locations else {
        return Ok(None);
    };
    let locations = match complete_location_array(locations, structure) {
        Ok(Some(value)) => value,
        Ok(None) => return Ok(None),
        Err(error) => {
            return preserve_compatibility_group(
                table,
                locations,
                labels,
                structure,
                error,
                options,
                diagnostics,
                compatibility_tables,
            );
        }
    };
    let parsed = (|| {
        parse(
            bounded_slice(table, locations[0], labels[0])?,
            bounded_slice(table, locations[1], labels[1])?,
            bounded_slice(table, locations[2], labels[2])?,
        )
    })();
    match parsed {
        Ok(value) => Ok(Some(DocBookmarkSet {
            metadata_location: locations[0],
            starts_location: locations[1],
            ends_location: locations[2],
            value,
        })),
        Err(error) => preserve_compatibility_group(
            table,
            locations,
            labels,
            structure,
            error,
            options,
            diagnostics,
            compatibility_tables,
        ),
    }
}

#[allow(clippy::too_many_arguments)]
fn preserve_compatibility_group<T, const N: usize>(
    table: &[u8],
    locations: [FibFcLcb; N],
    labels: [&'static str; N],
    structure: &'static str,
    error: Error,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
    compatibility_tables: &mut Vec<DocCompatibilityTable>,
) -> Result<Option<T>> {
    if options.is_strict() {
        return Err(error);
    }
    let reason = error.to_string();
    report_doc_compatibility(
        diagnostics,
        ParseDiagnosticCode::NonconformingRecord,
        locations
            .iter()
            .find(|location| location.lcb != 0)
            .map_or(0, |location| u64::from(location.fc)),
        structure,
        format!("preserved invalid {structure}: {reason}"),
    );
    for (location, label) in locations.into_iter().zip(labels) {
        if location.lcb != 0 {
            compatibility_tables.push(DocCompatibilityTable {
                label: label.to_owned(),
                location,
                physical_bytes: bounded_slice(table, location, label)
                    .ok()
                    .map(<[u8]>::to_vec),
                reason: reason.clone(),
            });
        }
    }
    Ok(None)
}

fn parse_range_protection(
    table: &[u8],
    locations: Option<[FibFcLcb; 4]>,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
    compatibility_tables: &mut Vec<DocCompatibilityTable>,
) -> Result<Option<DocRangeProtectionTables>> {
    let Some(locations) = locations else {
        return Ok(None);
    };
    let labels = [
        "SttbfBkmkProt",
        "PlcfBkfProt",
        "PlcfBklProt",
        "SttbProtUser",
    ];
    let locations = match complete_location_array(locations, "range-protection tables") {
        Ok(Some(value)) => value,
        Ok(None) => return Ok(None),
        Err(error) => {
            return preserve_compatibility_group(
                table,
                locations,
                labels,
                "range-protection tables",
                error,
                options,
                diagnostics,
                compatibility_tables,
            );
        }
    };
    let parsed = (|| {
        RangeProtection::from_bytes(
            bounded_slice(table, locations[0], labels[0])?,
            bounded_slice(table, locations[1], labels[1])?,
            bounded_slice(table, locations[2], labels[2])?,
            bounded_slice(table, locations[3], labels[3])?,
        )
    })();
    match parsed {
        Ok(value) => Ok(Some(DocRangeProtectionTables {
            permissions_location: locations[0],
            starts_location: locations[1],
            ends_location: locations[2],
            users_location: locations[3],
            value,
        })),
        Err(error) => preserve_compatibility_group(
            table,
            locations,
            labels,
            "range-protection tables",
            error,
            options,
            diagnostics,
            compatibility_tables,
        ),
    }
}

fn parse_user_input_methods(
    table: &[u8],
    locations: Option<[FibFcLcb; 2]>,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
    compatibility_tables: &mut Vec<DocCompatibilityTable>,
) -> Result<Option<DocUserInputMethodTables>> {
    let Some(locations) = locations else {
        return Ok(None);
    };
    let labels = ["PlcfUim", "PlfGuidUim"];
    let locations = match complete_location_array(locations, "user-input-method tables") {
        Ok(Some(value)) => value,
        Ok(None) => return Ok(None),
        Err(error) => {
            return preserve_compatibility_group(
                table,
                locations,
                labels,
                "user-input-method tables",
                error,
                options,
                diagnostics,
                compatibility_tables,
            );
        }
    };
    let parsed = (|| {
        UserInputMethods::from_bytes(
            bounded_slice(table, locations[0], labels[0])?,
            bounded_slice(table, locations[1], labels[1])?,
        )
    })();
    match parsed {
        Ok(value) => Ok(Some(DocUserInputMethodTables {
            methods_location: locations[0],
            service_guids_location: locations[1],
            value,
        })),
        Err(error) => preserve_compatibility_group(
            table,
            locations,
            labels,
            "user-input-method tables",
            error,
            options,
            diagnostics,
            compatibility_tables,
        ),
    }
}

fn report_doc_compatibility(
    diagnostics: &mut Vec<ParseDiagnostic>,
    code: ParseDiagnosticCode,
    offset: u64,
    structure: &'static str,
    message: String,
) {
    diagnostics.push(ParseDiagnostic::warning(
        code,
        BinaryFormat::Doc,
        Some("Table Stream"),
        Some(offset),
        structure,
        SpecificationReference {
            document: "MS-DOC",
            section: "2.8",
        },
        message,
    ));
}

fn parse_part_tables<T>(
    table: &[u8],
    locations: Vec<(TextboxDocumentPart, FibFcLcb)>,
    label: &str,
    parse: impl Fn(&[u8]) -> Result<T>,
) -> Result<BTreeMap<TextboxDocumentPart, DocLocated<T>>> {
    locations
        .into_iter()
        .filter(|(_, location)| location.lcb != 0)
        .map(|(part, location)| {
            Ok((
                part,
                DocLocated {
                    location,
                    value: parse(bounded_slice(table, location, label)?)?,
                },
            ))
        })
        .collect()
}

fn parse_caption_tables(
    table: &[u8],
    locations: Option<(FibFcLcb, FibFcLcb)>,
    options: ParseOptions,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<Option<DocCaptionTables>> {
    let (definitions, automatic) = match complete_locations(locations, "caption") {
        Ok(Some(locations)) => locations,
        Ok(None) => return Ok(None),
        Err(error) if options.is_strict() => return Err(error),
        Err(error) => {
            report_doc_compatibility(
                diagnostics,
                ParseDiagnosticCode::InvalidReference,
                0,
                "caption tables",
                format!("preserved incomplete caption tables: {error}"),
            );
            return Ok(None);
        }
    };
    Ok(Some(DocCaptionTables {
        definitions: DocLocated {
            location: definitions,
            value: CaptionDefinitions::from_bytes(bounded_slice(
                table,
                definitions,
                "SttbfCaption",
            )?)?,
        },
        automatic: DocLocated {
            location: automatic,
            value: AutoCaptionDefinitions::from_bytes(bounded_slice(
                table,
                automatic,
                "SttbfAutoCaption",
            )?)?,
        },
    }))
}

fn complete_locations(
    locations: Option<(FibFcLcb, FibFcLcb)>,
    label: &str,
) -> Result<Option<(FibFcLcb, FibFcLcb)>> {
    let Some(locations) = locations else {
        return Ok(None);
    };
    match (locations.0.lcb == 0, locations.1.lcb == 0) {
        (true, true) => Ok(None),
        (false, false) => Ok(Some(locations)),
        _ => Err(Error::invalid(
            0,
            format!("{label} table locations are incomplete"),
        )),
    }
}

fn complete_location_array<const N: usize>(
    locations: [FibFcLcb; N],
    label: &str,
) -> Result<Option<[FibFcLcb; N]>> {
    if locations.iter().all(|location| location.lcb == 0) {
        Ok(None)
    } else if locations.iter().all(|location| location.lcb != 0) {
        Ok(Some(locations))
    } else {
        Err(Error::invalid(0, format!("{label} are incomplete")))
    }
}

fn validate_optional_location<T>(
    expected: Option<FibFcLcb>,
    actual: Option<&DocLocated<T>>,
    label: &str,
) -> Result<()> {
    if expected.filter(|location| location.lcb != 0) != actual.map(|value| value.location) {
        return Err(Error::invalid(0, format!("DOC FIB/{label} link changed")));
    }
    Ok(())
}

fn validate_compatible_location<T>(
    expected: Option<FibFcLcb>,
    actual: Option<&DocLocated<T>>,
    compatibility: &[DocCompatibilityTable],
    physical_label: &str,
    link_label: &str,
) -> Result<()> {
    validate_optional_location(
        managed_expected_location(expected, compatibility, physical_label),
        actual,
        link_label,
    )
}

fn managed_expected_location(
    expected: Option<FibFcLcb>,
    compatibility: &[DocCompatibilityTable],
    label: &str,
) -> Option<FibFcLcb> {
    expected.filter(|location| {
        location.lcb != 0
            && !compatibility
                .iter()
                .any(|value| value.label == label && value.location == *location)
    })
}

fn managed_expected_locations<const N: usize>(
    expected: Option<[FibFcLcb; N]>,
    compatibility: &[DocCompatibilityTable],
    labels: [&str; N],
) -> Option<[FibFcLcb; N]> {
    expected.filter(|locations| {
        locations.iter().all(|location| location.lcb != 0)
            && !locations.iter().zip(labels).any(|(location, label)| {
                compatibility
                    .iter()
                    .any(|value| value.label == label && value.location == *location)
            })
    })
}

fn validate_compatibility_tables(
    table: &[u8],
    compatibility: &[DocCompatibilityTable],
) -> Result<()> {
    for value in compatibility {
        let physical = bounded_slice(table, value.location, &value.label).ok();
        if physical.map(<[u8]>::to_vec) != value.physical_bytes {
            return Err(Error::invalid(
                u64::from(value.location.fc),
                format!("DOC compatibility table {} link changed", value.label),
            ));
        }
    }
    Ok(())
}

fn validate_note_locations(
    expected: Option<(FibFcLcb, FibFcLcb)>,
    actual: Option<&DocNoteTables>,
    label: &str,
) -> Result<()> {
    let expected = expected.and_then(|(references, text)| {
        (references.lcb != 0 && text.lcb != 0).then_some((references, text))
    });
    let actual = actual.map(|value| (value.references.location, value.text.location));
    if expected != actual {
        return Err(Error::invalid(0, format!("DOC FIB/{label} links changed")));
    }
    Ok(())
}

fn validate_annotation_locations(
    expected: Option<(FibFcLcb, FibFcLcb)>,
    actual: Option<&DocAnnotationTables>,
) -> Result<()> {
    let expected = expected.and_then(|(references, text)| {
        (references.lcb != 0 && text.lcb != 0).then_some((references, text))
    });
    let actual = actual.map(|value| (value.references.location, value.text.location));
    if expected != actual {
        return Err(Error::invalid(0, "DOC FIB/annotation links changed"));
    }
    Ok(())
}

fn validate_caption_locations(
    expected: Option<(FibFcLcb, FibFcLcb)>,
    actual: Option<&DocCaptionTables>,
) -> Result<()> {
    let expected = expected.and_then(|(definitions, automatic)| {
        (definitions.lcb != 0 && automatic.lcb != 0).then_some((definitions, automatic))
    });
    let actual = actual.map(|value| (value.definitions.location, value.automatic.location));
    if expected != actual {
        return Err(Error::invalid(0, "DOC FIB/caption links changed"));
    }
    Ok(())
}

fn validate_part_locations<T>(
    expected: Vec<(TextboxDocumentPart, FibFcLcb)>,
    actual: &BTreeMap<TextboxDocumentPart, DocLocated<T>>,
    label: &str,
) -> Result<()> {
    let expected = expected
        .into_iter()
        .filter(|(_, location)| location.lcb != 0)
        .collect::<BTreeMap<_, _>>();
    let actual = actual
        .iter()
        .map(|(part, value)| (*part, value.location))
        .collect::<BTreeMap<_, _>>();
    if expected != actual {
        return Err(Error::invalid(0, format!("DOC FIB/{label} links changed")));
    }
    Ok(())
}

fn parse_required<T>(
    bytes: &[u8],
    location: Option<FibFcLcb>,
    label: &str,
    parse: impl Fn(&[u8]) -> Result<T>,
) -> Result<DocLocated<T>> {
    let location = location
        .filter(|value| value.lcb != 0)
        .ok_or_else(|| Error::invalid(0, format!("{label} location is missing")))?;
    Ok(DocLocated {
        location,
        value: parse(bounded_slice(bytes, location, label)?)?,
    })
}

fn parse_optional<T>(
    bytes: &[u8],
    location: Option<FibFcLcb>,
    label: &str,
    parse: impl Fn(&[u8]) -> Result<T>,
) -> Result<Option<DocLocated<T>>> {
    location
        .filter(|value| value.lcb != 0)
        .map(|location| {
            Ok(DocLocated {
                location,
                value: parse(bounded_slice(bytes, location, label)?)?,
            })
        })
        .transpose()
}

fn required_stream<'a>(compound: &'a CompoundFile, path: &str) -> Result<&'a [u8]> {
    compound
        .stream(path)
        .ok_or_else(|| Error::invalid(0, format!("required CFB stream {path} is missing")))
}

fn bounded_slice<'a>(bytes: &'a [u8], location: FibFcLcb, label: &str) -> Result<&'a [u8]> {
    let start = usize::try_from(location.fc)
        .map_err(|_| Error::Limit(format!("{label} offset exceeds usize")))?;
    let len = usize::try_from(location.lcb)
        .map_err(|_| Error::Limit(format!("{label} length exceeds usize")))?;
    bytes
        .get(start..start.saturating_add(len))
        .ok_or_else(|| Error::invalid(u64::from(location.fc), format!("{label} exceeds stream")))
}

#[derive(Debug)]
struct EncodedTextPiece {
    piece_index: usize,
    source_offset: u32,
    source_len: usize,
    source_width: usize,
    source_character_count: usize,
    destination_character_count: usize,
    destination_start: Option<u32>,
    character_replacements: Vec<CpReplacement>,
    compressed: bool,
    bytes: Vec<u8>,
}

#[derive(Clone, Debug)]
struct TextRelocation {
    source_start: u32,
    source_len: usize,
    source_width: usize,
    destination_start: u32,
    destination_width: usize,
    source_character_count: usize,
    destination_character_count: usize,
    character_replacements: Vec<CpReplacement>,
}

#[derive(Clone, Debug)]
struct PhysicalCharacterRun {
    start: u32,
    end: u32,
    properties: Option<GrpPrl>,
}

#[derive(Clone, Debug)]
struct PhysicalParagraphRun {
    start: u32,
    end: u32,
    formatting: PapxFkpRun,
}

fn rebuild_paragraph_formatting_pages(
    logical_runs: &[DocPapxRun],
    current_pieces: &[DocTextPiece],
    encoded_pieces: &[EncodedTextPiece],
    word: &mut Vec<u8>,
    fib: &mut Fib,
) -> Result<(PlcBte, Vec<DocFkpPage<PapxFkp>>)> {
    let current_ranges = current_paragraph_ranges(current_pieces)?;
    if logical_runs.len() != current_ranges.len()
        || logical_runs
            .iter()
            .zip(&current_ranges)
            .any(|(run, range)| (run.cp_start, run.cp_end) != *range)
    {
        return Err(Error::invalid(
            0,
            "PAPX CP tree does not match the current paragraph ranges",
        ));
    }
    let physical_runs =
        destination_paragraph_formatting_runs(logical_runs, current_pieces, encoded_pieces)?;
    let pages = paginate_paragraph_formatting_runs(&physical_runs)?;
    append_paragraph_formatting_pages(pages, word, fib)
}

fn source_paragraph_formatting_runs(
    pages: &[DocFkpPage<PapxFkp>],
    clx: &Clx,
    source_word: &[u8],
) -> Result<Vec<DocPapxRun>> {
    let mut fragments = Vec::<DocPapxRun>::new();
    for page in pages {
        for (range, run) in page.value.file_positions.windows(2).zip(&page.value.runs) {
            if range[0] >= range[1] {
                return Err(Error::invalid(
                    u64::from(range[1]),
                    "cannot rebuild a PapxFkp with non-increasing rgfc",
                ));
            }
            for (piece_index, descriptor) in clx.piece_table.pieces.iter().enumerate() {
                let cp_start = u32::try_from(clx.piece_table.character_positions[piece_index])
                    .map_err(|_| Error::invalid(0, "source text piece CP is negative"))?;
                let cp_end = u32::try_from(clx.piece_table.character_positions[piece_index + 1])
                    .map_err(|_| Error::invalid(0, "source text piece CP is negative"))?;
                let character_count = cp_end.checked_sub(cp_start).ok_or_else(|| {
                    Error::invalid(0, "source text piece CP range is not increasing")
                })?;
                let width = if descriptor.file_position.compressed {
                    1u64
                } else {
                    2u64
                };
                let piece_start = u64::from(descriptor.file_position.byte_offset());
                let piece_end = piece_start
                    .checked_add(u64::from(character_count) * width)
                    .ok_or_else(|| Error::Limit("source text piece FC limit overflow".into()))?;
                let start = u64::from(range[0]).max(piece_start);
                let end = u64::from(range[1]).min(piece_end);
                if start >= end {
                    continue;
                }
                let start_delta = start - piece_start;
                let end_delta = end - piece_start;
                if start_delta % width != 0 || end_delta % width != 0 {
                    return Err(Error::invalid(
                        start,
                        "PapxFkp boundary is not aligned to its text-piece encoding",
                    ));
                }
                fragments.push(DocPapxRun {
                    cp_start: cp_start
                        .checked_add(u32::try_from(start_delta / width).map_err(|_| {
                            Error::Limit("logical paragraph run start exceeds u32".into())
                        })?)
                        .ok_or_else(|| {
                            Error::Limit("logical paragraph run start overflow".into())
                        })?,
                    cp_end: cp_start
                        .checked_add(u32::try_from(end_delta / width).map_err(|_| {
                            Error::Limit("logical paragraph run limit exceeds u32".into())
                        })?)
                        .ok_or_else(|| {
                            Error::Limit("logical paragraph run limit overflow".into())
                        })?,
                    paragraph_height_info: run.paragraph_height_info,
                    properties: run.properties.clone(),
                });
            }
        }
    }
    fragments.sort_by_key(|run| (run.cp_start, run.cp_end));
    let source_ranges = source_paragraph_ranges(clx, source_word)?;
    source_ranges
        .into_iter()
        .map(|(start, end)| {
            let marker = end.saturating_sub(1);
            let formatting = fragments
                .iter()
                .find(|run| run.cp_start <= marker && marker < run.cp_end)
                .or_else(|| fragments.iter().find(|run| run.cp_end == end))
                .ok_or_else(|| {
                    Error::invalid(u64::from(start), "paragraph has no PapxFkp formatting run")
                })?;
            Ok(DocPapxRun {
                cp_start: start,
                cp_end: end,
                paragraph_height_info: formatting.paragraph_height_info,
                properties: formatting.properties.clone(),
            })
        })
        .collect()
}

fn source_paragraph_ranges(clx: &Clx, word: &[u8]) -> Result<Vec<(u32, u32)>> {
    let mut ranges = Vec::new();
    let mut paragraph_start = u32::try_from(
        *clx.piece_table
            .character_positions
            .first()
            .ok_or_else(|| Error::invalid(0, "PlcPcd has no CP positions"))?,
    )
    .map_err(|_| Error::invalid(0, "first PlcPcd CP is negative"))?;
    let mut document_end = paragraph_start;
    for (index, descriptor) in clx.piece_table.pieces.iter().enumerate() {
        let cp_start = clx.piece_table.character_positions[index];
        let cp_end = clx.piece_table.character_positions[index + 1];
        let piece = descriptor.text_piece(word, cp_start, cp_end)?;
        let cp_start = u32::try_from(cp_start)
            .map_err(|_| Error::invalid(0, "source text piece CP is negative"))?;
        for (offset, value) in text_piece_u16_values(&piece.characters).enumerate() {
            let cp = cp_start
                .checked_add(
                    u32::try_from(offset)
                        .map_err(|_| Error::Limit("paragraph CP exceeds u32".into()))?,
                )
                .ok_or_else(|| Error::Limit("paragraph CP overflow".into()))?;
            if is_paragraph_terminator(value) {
                let end = cp
                    .checked_add(1)
                    .ok_or_else(|| Error::Limit("paragraph limit overflow".into()))?;
                ranges.push((paragraph_start, end));
                paragraph_start = end;
            }
        }
        document_end = u32::try_from(cp_end)
            .map_err(|_| Error::invalid(0, "source text piece CP is negative"))?;
    }
    if paragraph_start < document_end {
        ranges.push((paragraph_start, document_end));
    }
    Ok(ranges)
}

fn current_paragraph_ranges(pieces: &[DocTextPiece]) -> Result<Vec<(u32, u32)>> {
    let mut ranges = Vec::new();
    let mut paragraph_start = pieces
        .first()
        .map(|piece| u32::try_from(piece.value.cp_start))
        .transpose()
        .map_err(|_| Error::invalid(0, "first text piece CP is negative"))?
        .unwrap_or(0);
    let mut document_end = paragraph_start;
    for piece in pieces {
        let cp_start = u32::try_from(piece.value.cp_start)
            .map_err(|_| Error::invalid(0, "text piece CP is negative"))?;
        for (offset, value) in text_piece_u16_values(&piece.value.characters).enumerate() {
            let cp = cp_start
                .checked_add(
                    u32::try_from(offset)
                        .map_err(|_| Error::Limit("paragraph CP exceeds u32".into()))?,
                )
                .ok_or_else(|| Error::Limit("paragraph CP overflow".into()))?;
            if is_paragraph_terminator(value) {
                let end = cp
                    .checked_add(1)
                    .ok_or_else(|| Error::Limit("paragraph limit overflow".into()))?;
                ranges.push((paragraph_start, end));
                paragraph_start = end;
            }
        }
        document_end = u32::try_from(piece.value.cp_end)
            .map_err(|_| Error::invalid(0, "text piece CP is negative"))?;
    }
    if paragraph_start < document_end {
        ranges.push((paragraph_start, document_end));
    }
    Ok(ranges)
}

fn text_piece_u16_values(characters: &TextPieceCharacters) -> Box<dyn Iterator<Item = u16> + '_> {
    match characters {
        TextPieceCharacters::Compressed(values) => Box::new(values.iter().copied().map(u16::from)),
        TextPieceCharacters::Utf16(values) => Box::new(values.iter().copied()),
    }
}

fn destination_paragraph_formatting_runs(
    logical_runs: &[DocPapxRun],
    pieces: &[DocTextPiece],
    encoded_pieces: &[EncodedTextPiece],
) -> Result<Vec<PhysicalParagraphRun>> {
    logical_runs
        .iter()
        .map(|run| {
            Ok(PhysicalParagraphRun {
                start: destination_fc_for_cp(run.cp_start, pieces, encoded_pieces)?,
                end: destination_fc_for_cp(run.cp_end, pieces, encoded_pieces)?,
                formatting: PapxFkpRun {
                    property_offset: None,
                    paragraph_height_info: run.paragraph_height_info,
                    properties: run.properties.clone(),
                },
            })
        })
        .collect()
}

fn destination_fc_for_cp(
    cp: u32,
    pieces: &[DocTextPiece],
    encoded_pieces: &[EncodedTextPiece],
) -> Result<u32> {
    for piece in pieces {
        let start = u32::try_from(piece.value.cp_start)
            .map_err(|_| Error::invalid(0, "destination text piece CP is negative"))?;
        let end = u32::try_from(piece.value.cp_end)
            .map_err(|_| Error::invalid(0, "destination text piece CP is negative"))?;
        if cp < start || cp > end {
            continue;
        }
        let encoded = encoded_pieces
            .iter()
            .find(|encoded| encoded.piece_index == piece.piece_index)
            .ok_or_else(|| Error::invalid(0, "destination text piece layout is missing"))?;
        let destination_start = encoded
            .destination_start
            .ok_or_else(|| Error::invalid(0, "destination text piece has no FC"))?;
        let width = if encoded.compressed { 1u32 } else { 2u32 };
        return destination_start
            .checked_add(
                (cp - start)
                    .checked_mul(width)
                    .ok_or_else(|| Error::Limit("destination paragraph FC overflow".into()))?,
            )
            .ok_or_else(|| Error::Limit("destination paragraph FC overflow".into()));
    }
    Err(Error::invalid(
        u64::from(cp),
        "paragraph CP is outside destination text pieces",
    ))
}

fn paginate_paragraph_formatting_runs(runs: &[PhysicalParagraphRun]) -> Result<Vec<PapxFkp>> {
    let mut pages = Vec::new();
    let mut current = Vec::<PhysicalParagraphRun>::new();
    for run in runs {
        let mut candidate = current.clone();
        candidate.push(run.clone());
        let fits = candidate.len() <= 0x1d && build_paragraph_formatting_page(&candidate).is_ok();
        if !fits && !current.is_empty() {
            pages.push(build_paragraph_formatting_page(&current)?);
            current.clear();
        }
        current.push(run.clone());
        build_paragraph_formatting_page(&current)?;
    }
    if !current.is_empty() {
        pages.push(build_paragraph_formatting_page(&current)?);
    }
    Ok(pages)
}

fn build_paragraph_formatting_page(runs: &[PhysicalParagraphRun]) -> Result<PapxFkp> {
    let mut positions = Vec::with_capacity(runs.len() + 1);
    let mut page_runs = Vec::with_capacity(runs.len());
    for (index, run) in runs.iter().enumerate() {
        if index == 0 {
            positions.push(run.start);
        } else if positions.last().copied() != Some(run.start) {
            return Err(Error::invalid(
                u64::from(run.start),
                "paragraph formatting page runs are not adjacent",
            ));
        }
        positions.push(run.end);
        let mut formatting = run.formatting.clone();
        formatting.property_offset = None;
        page_runs.push(formatting);
    }
    PapxFkp::with_canonical_layout(positions, page_runs)
}

fn append_paragraph_formatting_pages(
    pages: Vec<PapxFkp>,
    word: &mut Vec<u8>,
    fib: &mut Fib,
) -> Result<(PlcBte, Vec<DocFkpPage<PapxFkp>>)> {
    if pages.is_empty() {
        return Err(Error::invalid(0, "paragraph formatting has no pages"));
    }
    let meaningful_end = usize::try_from(fib.rg_lw.cb_mac)
        .map_err(|_| Error::Limit("FIB cbMac exceeds usize".into()))?;
    if meaningful_end > word.len() {
        return Err(Error::invalid(
            u64::from(fib.rg_lw.cb_mac),
            "FIB cbMac exceeds WordDocument",
        ));
    }
    let page_start = meaningful_end
        .checked_add(511)
        .map(|value| value & !511)
        .ok_or_else(|| Error::Limit("PapxFkp alignment overflow".into()))?;
    let page_bytes = pages
        .len()
        .checked_mul(512)
        .ok_or_else(|| Error::Limit("PapxFkp page bytes overflow".into()))?;
    let page_end = page_start
        .checked_add(page_bytes)
        .ok_or_else(|| Error::Limit("PapxFkp page end overflow".into()))?;
    word.splice(
        meaningful_end..meaningful_end,
        vec![0; page_end - meaningful_end],
    );
    let first_page_number = page_start / 512;
    let mut located_pages = Vec::with_capacity(pages.len());
    let mut bin_positions = Vec::with_capacity(pages.len() + 1);
    let mut page_numbers = Vec::with_capacity(pages.len());
    for (index, page) in pages.into_iter().enumerate() {
        let offset = page_start
            .checked_add(
                index
                    .checked_mul(512)
                    .ok_or_else(|| Error::Limit("PapxFkp page offset overflow".into()))?,
            )
            .ok_or_else(|| Error::Limit("PapxFkp page offset overflow".into()))?;
        word[offset..offset + 512].copy_from_slice(&page.to_bytes()?);
        bin_positions.push(page.file_positions[0]);
        let page_ref = FkpPageNumber {
            page_number: u32::try_from(
                first_page_number
                    .checked_add(index)
                    .ok_or_else(|| Error::Limit("PapxFkp page number overflow".into()))?,
            )
            .map_err(|_| Error::Limit("PapxFkp page number exceeds u32".into()))?,
            unused: 0,
        };
        page_numbers.push(page_ref);
        located_pages.push(DocFkpPage {
            page: page_ref,
            value: page,
        });
    }
    bin_positions.push(
        located_pages
            .last()
            .and_then(|page| page.value.file_positions.last())
            .copied()
            .ok_or_else(|| Error::invalid(0, "paragraph formatting has no physical runs"))?,
    );
    fib.rg_lw.cb_mac =
        u32::try_from(page_end).map_err(|_| Error::Limit("FIB cbMac exceeds u32".into()))?;
    Ok((
        PlcBte {
            file_positions: bin_positions,
            pages: page_numbers,
        },
        located_pages,
    ))
}

fn rebuild_character_formatting_pages(
    logical_runs: &[DocChpxRun],
    current_pieces: &[DocTextPiece],
    encoded_pieces: &[EncodedTextPiece],
    word: &mut Vec<u8>,
    fib: &mut Fib,
) -> Result<(PlcBte, Vec<DocFkpPage<ChpxFkp>>)> {
    let logical_runs = normalize_logical_character_runs(logical_runs.to_vec())?;
    let physical_runs =
        destination_character_formatting_runs(&logical_runs, current_pieces, encoded_pieces)?;
    let pages = paginate_character_formatting_runs(&physical_runs)?;
    let meaningful_end = usize::try_from(fib.rg_lw.cb_mac)
        .map_err(|_| Error::Limit("FIB cbMac exceeds usize".into()))?;
    if meaningful_end > word.len() {
        return Err(Error::invalid(
            u64::from(fib.rg_lw.cb_mac),
            "FIB cbMac exceeds WordDocument",
        ));
    }
    let page_start = meaningful_end
        .checked_add(511)
        .map(|value| value & !511)
        .ok_or_else(|| Error::Limit("ChpxFkp alignment overflow".into()))?;
    let page_bytes = pages
        .len()
        .checked_mul(512)
        .ok_or_else(|| Error::Limit("ChpxFkp page bytes overflow".into()))?;
    let page_end = page_start
        .checked_add(page_bytes)
        .ok_or_else(|| Error::Limit("ChpxFkp page end overflow".into()))?;
    word.splice(
        meaningful_end..meaningful_end,
        vec![0; page_end - meaningful_end],
    );

    let first_page_number = page_start / 512;
    let mut located_pages = Vec::with_capacity(pages.len());
    let mut bin_positions = Vec::with_capacity(pages.len() + 1);
    let mut page_numbers = Vec::with_capacity(pages.len());
    for (index, page) in pages.into_iter().enumerate() {
        let offset = page_start
            .checked_add(
                index
                    .checked_mul(512)
                    .ok_or_else(|| Error::Limit("ChpxFkp page offset overflow".into()))?,
            )
            .ok_or_else(|| Error::Limit("ChpxFkp page offset overflow".into()))?;
        let bytes = page.to_bytes()?;
        word[offset..offset + 512].copy_from_slice(&bytes);
        bin_positions.push(page.file_positions[0]);
        let page_number = u32::try_from(
            first_page_number
                .checked_add(index)
                .ok_or_else(|| Error::Limit("ChpxFkp page number overflow".into()))?,
        )
        .map_err(|_| Error::Limit("ChpxFkp page number exceeds u32".into()))?;
        let page_ref = FkpPageNumber {
            page_number,
            unused: 0,
        };
        page_numbers.push(page_ref);
        located_pages.push(DocFkpPage {
            page: page_ref,
            value: page,
        });
    }
    let last_position = located_pages
        .last()
        .and_then(|page| page.value.file_positions.last())
        .copied()
        .ok_or_else(|| Error::invalid(0, "character formatting has no physical runs"))?;
    bin_positions.push(last_position);
    fib.rg_lw.cb_mac =
        u32::try_from(page_end).map_err(|_| Error::Limit("FIB cbMac exceeds u32".into()))?;
    Ok((
        PlcBte {
            file_positions: bin_positions,
            pages: page_numbers,
        },
        located_pages,
    ))
}

fn source_character_formatting_runs(
    pages: &[DocFkpPage<ChpxFkp>],
    clx: &Clx,
) -> Result<Vec<DocChpxRun>> {
    let mut runs = Vec::new();
    for page in pages {
        for (range, run) in page.value.file_positions.windows(2).zip(&page.value.runs) {
            if range[0] >= range[1] {
                return Err(Error::invalid(
                    u64::from(range[1]),
                    "cannot rebuild a ChpxFkp with non-increasing rgfc",
                ));
            }
            for (piece_index, descriptor) in clx.piece_table.pieces.iter().enumerate() {
                let cp_start = u32::try_from(clx.piece_table.character_positions[piece_index])
                    .map_err(|_| Error::invalid(0, "source text piece CP is negative"))?;
                let cp_end = u32::try_from(clx.piece_table.character_positions[piece_index + 1])
                    .map_err(|_| Error::invalid(0, "source text piece CP is negative"))?;
                let character_count = cp_end.checked_sub(cp_start).ok_or_else(|| {
                    Error::invalid(0, "source text piece CP range is not increasing")
                })?;
                let width = if descriptor.file_position.compressed {
                    1u64
                } else {
                    2u64
                };
                let piece_start = u64::from(descriptor.file_position.byte_offset());
                let piece_end = piece_start
                    .checked_add(u64::from(character_count) * width)
                    .ok_or_else(|| Error::Limit("source text piece FC limit overflow".into()))?;
                let start = u64::from(range[0]).max(piece_start);
                let end = u64::from(range[1]).min(piece_end);
                if start >= end {
                    continue;
                }
                let start_delta = start - piece_start;
                let end_delta = end - piece_start;
                if start_delta % width != 0 || end_delta % width != 0 {
                    return Err(Error::invalid(
                        start,
                        "ChpxFkp boundary is not aligned to its text-piece encoding",
                    ));
                }
                runs.push(DocChpxRun {
                    cp_start: cp_start
                        .checked_add(u32::try_from(start_delta / width).map_err(|_| {
                            Error::Limit("logical character run start exceeds u32".into())
                        })?)
                        .ok_or_else(|| {
                            Error::Limit("logical character run start overflow".into())
                        })?,
                    cp_end: cp_start
                        .checked_add(u32::try_from(end_delta / width).map_err(|_| {
                            Error::Limit("logical character run limit exceeds u32".into())
                        })?)
                        .ok_or_else(|| {
                            Error::Limit("logical character run limit overflow".into())
                        })?,
                    properties: run.properties.clone(),
                });
            }
        }
    }
    normalize_logical_character_runs(runs)
}

fn normalize_logical_character_runs(mut runs: Vec<DocChpxRun>) -> Result<Vec<DocChpxRun>> {
    runs.sort_by_key(|run| (run.cp_start, run.cp_end));
    let mut normalized: Vec<DocChpxRun> = Vec::with_capacity(runs.len());
    for run in runs {
        if run.cp_start >= run.cp_end {
            return Err(Error::invalid(
                u64::from(run.cp_start),
                "logical character formatting run is empty",
            ));
        }
        if let Some(previous) = normalized.last_mut() {
            if run.cp_start < previous.cp_end {
                return Err(Error::invalid(
                    u64::from(run.cp_start),
                    "logical character formatting runs overlap",
                ));
            }
            if run.cp_start == previous.cp_end && run.properties == previous.properties {
                previous.cp_end = run.cp_end;
                continue;
            }
        }
        normalized.push(run);
    }
    Ok(normalized)
}

fn apply_character_run_edit(runs: &mut Vec<DocChpxRun>, edit: &CpReplacement) -> Result<()> {
    let replacement_properties = runs
        .iter()
        .find(|run| run.cp_start <= edit.old_start && edit.old_start < run.cp_end)
        .or_else(|| runs.last().filter(|run| run.cp_end == edit.old_start))
        .and_then(|run| run.properties.clone());
    let mut edited = Vec::with_capacity(runs.len() + 1);
    for run in runs.iter() {
        if run.cp_end <= edit.old_start {
            edited.push(run.clone());
            continue;
        }
        if run.cp_start >= edit.old_end {
            edited.push(DocChpxRun {
                cp_start: edit.relocate_u32(run.cp_start, "character formatting run start")?,
                cp_end: edit.relocate_u32(run.cp_end, "character formatting run limit")?,
                properties: run.properties.clone(),
            });
            continue;
        }
        if run.cp_start < edit.old_start {
            edited.push(DocChpxRun {
                cp_start: run.cp_start,
                cp_end: edit.old_start,
                properties: run.properties.clone(),
            });
        }
        if run.cp_end > edit.old_end {
            edited.push(DocChpxRun {
                cp_start: edit.new_end,
                cp_end: edit.relocate_u32(run.cp_end, "character formatting run limit")?,
                properties: run.properties.clone(),
            });
        }
    }
    if edit.new_end > edit.old_start {
        edited.push(DocChpxRun {
            cp_start: edit.old_start,
            cp_end: edit.new_end,
            properties: replacement_properties,
        });
    }
    *runs = normalize_logical_character_runs(edited)?;
    Ok(())
}

fn destination_character_formatting_runs(
    logical_runs: &[DocChpxRun],
    current_pieces: &[DocTextPiece],
    encoded_pieces: &[EncodedTextPiece],
) -> Result<Vec<PhysicalCharacterRun>> {
    let mut runs = Vec::new();
    for logical in logical_runs {
        for piece in current_pieces {
            let piece_start = u32::try_from(piece.value.cp_start)
                .map_err(|_| Error::invalid(0, "destination text piece CP is negative"))?;
            let piece_end = u32::try_from(piece.value.cp_end)
                .map_err(|_| Error::invalid(0, "destination text piece CP is negative"))?;
            let start = logical.cp_start.max(piece_start);
            let end = logical.cp_end.min(piece_end);
            if start >= end {
                continue;
            }
            let encoded = encoded_pieces
                .iter()
                .find(|encoded| encoded.piece_index == piece.piece_index)
                .ok_or_else(|| Error::invalid(0, "destination text piece layout is missing"))?;
            let destination_start = encoded.destination_start.ok_or_else(|| {
                Error::invalid(0, "destination text piece was not assigned an FC")
            })?;
            let width = if encoded.compressed { 1u32 } else { 2u32 };
            let physical_start =
                destination_start
                    .checked_add((start - piece_start).checked_mul(width).ok_or_else(|| {
                        Error::Limit("destination character run FC overflow".into())
                    })?)
                    .ok_or_else(|| Error::Limit("destination character run FC overflow".into()))?;
            let physical_end =
                destination_start
                    .checked_add((end - piece_start).checked_mul(width).ok_or_else(|| {
                        Error::Limit("destination character run FC overflow".into())
                    })?)
                    .ok_or_else(|| Error::Limit("destination character run FC overflow".into()))?;
            runs.push(PhysicalCharacterRun {
                start: physical_start,
                end: physical_end,
                properties: logical.properties.clone(),
            });
        }
    }
    runs.sort_by_key(|run| (run.start, run.end));
    let mut normalized: Vec<PhysicalCharacterRun> = Vec::with_capacity(runs.len());
    for run in runs {
        if let Some(previous) = normalized.last() {
            let previous_end = previous.end;
            let same_properties = run.properties == previous.properties;
            if run.start < previous_end {
                return Err(Error::invalid(
                    u64::from(run.start),
                    "destination character formatting runs overlap",
                ));
            }
            if run.start > previous_end {
                normalized.push(PhysicalCharacterRun {
                    start: previous_end,
                    end: run.start,
                    properties: None,
                });
            } else if same_properties {
                normalized.last_mut().expect("previous run exists").end = run.end;
                continue;
            }
        }
        normalized.push(run);
    }
    if normalized.is_empty() {
        return Err(Error::invalid(
            0,
            "character formatting has no destination runs",
        ));
    }
    Ok(normalized)
}

fn paginate_character_formatting_runs(runs: &[PhysicalCharacterRun]) -> Result<Vec<ChpxFkp>> {
    let mut pages = Vec::new();
    let mut current = Vec::<PhysicalCharacterRun>::new();
    for run in runs {
        let mut candidate = current.clone();
        candidate.push(run.clone());
        let fits = candidate.len() <= 0x65 && build_character_formatting_page(&candidate).is_ok();
        if !fits && !current.is_empty() {
            pages.push(build_character_formatting_page(&current)?);
            current.clear();
        }
        current.push(run.clone());
        build_character_formatting_page(&current)?;
    }
    if !current.is_empty() {
        pages.push(build_character_formatting_page(&current)?);
    }
    Ok(pages)
}

fn build_character_formatting_page(runs: &[PhysicalCharacterRun]) -> Result<ChpxFkp> {
    let mut positions = Vec::with_capacity(runs.len() + 1);
    let mut page_runs = Vec::with_capacity(runs.len());
    for (index, run) in runs.iter().enumerate() {
        if index == 0 {
            positions.push(run.start);
        } else if positions.last().copied() != Some(run.start) {
            return Err(Error::invalid(
                u64::from(run.start),
                "character formatting page runs are not adjacent",
            ));
        }
        positions.push(run.end);
        page_runs.push(ChpxFkpRun {
            property_offset: None,
            properties: run.properties.clone(),
        });
    }
    ChpxFkp::with_canonical_layout(positions, page_runs)
}

fn relocate_character_position(
    mut position: u32,
    edits: &[CpReplacement],
    label: &str,
) -> Result<u32> {
    for edit in edits {
        position = edit.relocate_u32(position, label)?;
    }
    Ok(position)
}

fn is_paragraph_terminator(value: u16) -> bool {
    matches!(value, 0x0007 | 0x000c | 0x000d)
}

fn paragraph_terminators(characters: &TextPieceCharacters) -> Vec<u16> {
    match characters {
        TextPieceCharacters::Compressed(values) => values
            .iter()
            .copied()
            .map(u16::from)
            .filter(|value| is_paragraph_terminator(*value))
            .collect(),
        TextPieceCharacters::Utf16(values) => values
            .iter()
            .copied()
            .filter(|value| is_paragraph_terminator(*value))
            .collect(),
    }
}

fn non_paragraph_terminators(terminators: &[u16]) -> Vec<u16> {
    terminators
        .iter()
        .copied()
        .filter(|value| *value != 0x000d)
        .collect()
}

fn text_value_at_cp(pieces: &[DocTextPiece], cp: u32) -> Result<u16> {
    for piece in pieces {
        let start = u32::try_from(piece.value.cp_start)
            .map_err(|_| Error::invalid(0, "text piece begins at a negative CP"))?;
        let end = u32::try_from(piece.value.cp_end)
            .map_err(|_| Error::invalid(0, "text piece ends at a negative CP"))?;
        if cp < start || cp >= end {
            continue;
        }
        let index = usize::try_from(cp - start)
            .map_err(|_| Error::Limit("text CP exceeds usize".into()))?;
        return match &piece.value.characters {
            TextPieceCharacters::Compressed(values) => {
                values.get(index).copied().map(u16::from).ok_or_else(|| {
                    Error::invalid(u64::from(cp), "text CP exceeds compressed piece")
                })
            }
            TextPieceCharacters::Utf16(values) => values
                .get(index)
                .copied()
                .ok_or_else(|| Error::invalid(u64::from(cp), "text CP exceeds UTF-16 piece")),
        };
    }
    Err(Error::invalid(
        u64::from(cp),
        "text CP is outside PlcPcd pieces",
    ))
}

fn validate_nondecreasing_part_positions(
    positions: &[u32],
    part_len: u32,
    label: &str,
) -> Result<()> {
    if positions.is_empty()
        || positions.iter().any(|position| *position >= part_len)
        || positions.windows(2).any(|pair| pair[0] > pair[1])
    {
        return Err(Error::invalid(
            0,
            format!("{label} values are outside the document part or not sorted"),
        ));
    }
    Ok(())
}

fn validate_strict_part_positions(positions: &[u32], part_len: u32, label: &str) -> Result<()> {
    if positions.is_empty()
        || positions.iter().any(|position| *position >= part_len)
        || positions.windows(2).any(|pair| pair[0] >= pair[1])
    {
        return Err(Error::invalid(
            0,
            format!("{label} values are outside the document part or not strictly increasing"),
        ));
    }
    Ok(())
}

fn validate_strict_textbox_positions(positions: &[u32], part_len: u32, label: &str) -> Result<()> {
    if positions.is_empty()
        || positions.iter().enumerate().any(|(index, position)| {
            if index + 1 == positions.len() {
                *position > part_len
            } else {
                *position >= part_len
            }
        })
        || positions.windows(2).any(|pair| pair[0] >= pair[1])
    {
        return Err(Error::invalid(
            0,
            format!("{label} values are outside the document part or not strictly increasing"),
        ));
    }
    Ok(())
}

fn require_part_character(
    pieces: &[DocTextPiece],
    part_start: u32,
    local_cp: u32,
    expected: u16,
    label: &str,
) -> Result<()> {
    let global_cp = part_start
        .checked_add(local_cp)
        .ok_or_else(|| Error::Limit(format!("{label} CP overflow")))?;
    if text_value_at_cp(pieces, global_cp)? != expected {
        return Err(Error::invalid(
            u64::from(global_cp),
            format!("{label} is not U+{expected:04X}"),
        ));
    }
    Ok(())
}

fn paragraph_terminators_in_range(
    characters: &TextPieceCharacters,
    start: usize,
    end: usize,
) -> Result<Vec<u16>> {
    match characters {
        TextPieceCharacters::Compressed(values) => values
            .get(start..end)
            .ok_or_else(|| Error::invalid(start as u64, "text replacement range exceeds piece"))
            .map(|values| {
                values
                    .iter()
                    .copied()
                    .map(u16::from)
                    .filter(|value| is_paragraph_terminator(*value))
                    .collect()
            }),
        TextPieceCharacters::Utf16(values) => values
            .get(start..end)
            .ok_or_else(|| Error::invalid(start as u64, "text replacement range exceeds piece"))
            .map(|values| {
                values
                    .iter()
                    .copied()
                    .filter(|value| is_paragraph_terminator(*value))
                    .collect()
            }),
    }
}

fn paragraph_terminators_in_piece_range(
    pieces: &[DocTextPiece],
    start: u32,
    end: u32,
) -> Result<Vec<u16>> {
    let mut terminators = Vec::new();
    for piece in pieces {
        let piece_start = u32::try_from(piece.value.cp_start)
            .map_err(|_| Error::invalid(0, "text piece begins at a negative CP"))?;
        let piece_end = u32::try_from(piece.value.cp_end)
            .map_err(|_| Error::invalid(0, "text piece ends at a negative CP"))?;
        let overlap_start = start.max(piece_start);
        let overlap_end = end.min(piece_end);
        if overlap_start >= overlap_end {
            continue;
        }
        terminators.extend(paragraph_terminators_in_range(
            &piece.value.characters,
            usize::try_from(overlap_start - piece_start)
                .map_err(|_| Error::Limit("text replacement start exceeds usize".into()))?,
            usize::try_from(overlap_end - piece_start)
                .map_err(|_| Error::Limit("text replacement limit exceeds usize".into()))?,
        )?);
    }
    Ok(terminators)
}

fn text_piece_character_replacement(
    source: &TextPieceCharacters,
    destination: &TextPieceCharacters,
) -> Result<CpReplacement> {
    let source_len = source.character_count();
    let destination_len = destination.character_count();
    if source_len == destination_len {
        let end = u32::try_from(source_len)
            .map_err(|_| Error::Limit("text piece character count exceeds u32".into()))?;
        return CpReplacement::new(end, end, 0);
    }
    let (prefix, suffix) = match (source, destination) {
        (TextPieceCharacters::Compressed(source), TextPieceCharacters::Compressed(destination)) => {
            common_prefix_and_suffix(source, destination)
        }
        (TextPieceCharacters::Utf16(source), TextPieceCharacters::Utf16(destination)) => {
            common_prefix_and_suffix(source, destination)
        }
        _ => {
            return Err(Error::invalid(
                0,
                "text piece encoding and character count changed together",
            ));
        }
    };
    let old_start =
        u32::try_from(prefix).map_err(|_| Error::Limit("text edit start exceeds u32".into()))?;
    let old_end = u32::try_from(source_len - suffix)
        .map_err(|_| Error::Limit("text edit limit exceeds u32".into()))?;
    let replacement_len = u32::try_from(destination_len - prefix - suffix)
        .map_err(|_| Error::Limit("replacement character count exceeds u32".into()))?;
    CpReplacement::new(old_start, old_end, replacement_len)
}

fn common_prefix_and_suffix<T: PartialEq>(source: &[T], destination: &[T]) -> (usize, usize) {
    let prefix = source
        .iter()
        .zip(destination)
        .take_while(|(left, right)| left == right)
        .count();
    let suffix_limit = source
        .len()
        .saturating_sub(prefix)
        .min(destination.len().saturating_sub(prefix));
    let suffix = source
        .iter()
        .rev()
        .zip(destination.iter().rev())
        .take(suffix_limit)
        .take_while(|(left, right)| left == right)
        .count();
    (prefix, suffix)
}

fn relocate_text_file_positions(
    positions: &mut [u32],
    relocations: &[TextRelocation],
) -> Result<()> {
    for position in positions {
        for relocation in relocations {
            let source_end = u64::from(relocation.source_start)
                .checked_add(relocation.source_len as u64)
                .ok_or_else(|| Error::Limit("text relocation source range overflow".into()))?;
            let raw_position = u64::from(*position);
            if raw_position < u64::from(relocation.source_start) || raw_position > source_end {
                continue;
            }
            let delta = usize::try_from(raw_position - u64::from(relocation.source_start))
                .map_err(|_| Error::Limit("text relocation delta exceeds usize".into()))?;
            if !delta.is_multiple_of(relocation.source_width) {
                return Err(Error::invalid(
                    raw_position,
                    "formatting FC is not aligned to its text-piece encoding",
                ));
            }
            let character_offset = delta / relocation.source_width;
            if character_offset > relocation.source_character_count {
                return Err(Error::invalid(
                    raw_position,
                    "formatting FC exceeds its source text piece",
                ));
            }
            let relocated_character_offset = usize::try_from(relocate_character_position(
                u32::try_from(character_offset)
                    .map_err(|_| Error::Limit("formatting character offset exceeds u32".into()))?,
                &relocation.character_replacements,
                "formatting boundary",
            )?)
            .map_err(|_| Error::Limit("relocated character offset exceeds usize".into()))?;
            if relocated_character_offset > relocation.destination_character_count {
                return Err(Error::invalid(
                    raw_position,
                    "relocated formatting FC exceeds its destination text piece",
                ));
            }
            let relocated_delta = relocated_character_offset
                .checked_mul(relocation.destination_width)
                .ok_or_else(|| Error::Limit("relocated formatting FC overflow".into()))?;
            *position = relocation
                .destination_start
                .checked_add(
                    u32::try_from(relocated_delta)
                        .map_err(|_| Error::Limit("relocated formatting FC exceeds u32".into()))?,
                )
                .ok_or_else(|| Error::Limit("relocated formatting FC overflow".into()))?;
            break;
        }
    }
    Ok(())
}

trait PatchSink {
    fn replace(
        &mut self,
        offset: usize,
        expected: usize,
        encoded: Vec<u8>,
        label: &str,
    ) -> Result<()>;
}

impl PatchSink for Vec<u8> {
    fn replace(
        &mut self,
        offset: usize,
        expected: usize,
        encoded: Vec<u8>,
        label: &str,
    ) -> Result<()> {
        if encoded.len() != expected {
            return Err(Error::invalid(
                offset as u64,
                format!("{label} size changed; WordDocument relocation is not implemented"),
            ));
        }
        self.get_mut(offset..offset.saturating_add(expected))
            .ok_or_else(|| Error::invalid(offset as u64, format!("{label} exceeds stream")))?
            .copy_from_slice(&encoded);
        Ok(())
    }
}

#[derive(Debug)]
struct PendingTableReplacement {
    offset: usize,
    expected: usize,
    encoded: Vec<u8>,
    label: String,
}

#[derive(Debug)]
struct TableLayout {
    original: Vec<u8>,
    replacements: Vec<PendingTableReplacement>,
}

#[derive(Clone, Debug)]
struct AppliedTableReplacement {
    old_offset: usize,
    old_len: usize,
    new_offset: usize,
    new_len: usize,
}

#[derive(Clone, Debug)]
struct TableRelocation {
    original_len: usize,
    changed_layout: bool,
    replacements: Vec<AppliedTableReplacement>,
}

impl TableLayout {
    fn new(original: Vec<u8>) -> Self {
        Self {
            original,
            replacements: Vec::new(),
        }
    }

    fn finish(mut self) -> Result<(Vec<u8>, TableRelocation)> {
        self.replacements.sort_by_key(|value| value.offset);
        for pair in self.replacements.windows(2) {
            let previous_end = pair[0]
                .offset
                .checked_add(pair[0].expected)
                .ok_or_else(|| Error::Limit("Table Stream replacement end overflow".into()))?;
            if previous_end > pair[1].offset {
                return Err(Error::invalid(
                    pair[1].offset as u64,
                    format!(
                        "Table Stream replacements {} and {} overlap",
                        pair[0].label, pair[1].label
                    ),
                ));
            }
        }

        let original_len = self.original.len();
        let mut output = Vec::new();
        let mut cursor = 0usize;
        let mut applied = Vec::with_capacity(self.replacements.len());
        for replacement in self.replacements {
            let old_end = replacement
                .offset
                .checked_add(replacement.expected)
                .ok_or_else(|| Error::Limit("Table Stream replacement end overflow".into()))?;
            let unchanged = self
                .original
                .get(cursor..replacement.offset)
                .ok_or_else(|| {
                    Error::invalid(
                        replacement.offset as u64,
                        format!("{} exceeds Table Stream", replacement.label),
                    )
                })?;
            if old_end > original_len {
                return Err(Error::invalid(
                    replacement.offset as u64,
                    format!("{} exceeds Table Stream", replacement.label),
                ));
            }
            output.extend_from_slice(unchanged);
            let new_offset = output.len();
            let new_len = replacement.encoded.len();
            output.extend_from_slice(&replacement.encoded);
            applied.push(AppliedTableReplacement {
                old_offset: replacement.offset,
                old_len: replacement.expected,
                new_offset,
                new_len,
            });
            cursor = old_end;
        }
        output.extend_from_slice(&self.original[cursor..]);
        Ok((
            output,
            TableRelocation {
                original_len,
                changed_layout: applied.iter().any(|value| value.old_len != value.new_len),
                replacements: applied,
            },
        ))
    }
}

impl PatchSink for TableLayout {
    fn replace(
        &mut self,
        offset: usize,
        expected: usize,
        encoded: Vec<u8>,
        label: &str,
    ) -> Result<()> {
        if expected == 0 && encoded.is_empty() {
            return Ok(());
        }
        if offset
            .checked_add(expected)
            .is_none_or(|end| end > self.original.len())
        {
            return Err(Error::invalid(
                offset as u64,
                format!("{label} exceeds Table Stream"),
            ));
        }
        self.replacements.push(PendingTableReplacement {
            offset,
            expected,
            encoded,
            label: label.to_owned(),
        });
        Ok(())
    }
}

impl TableRelocation {
    fn relocate(&self, location: FibFcLcb) -> Result<Option<FibFcLcb>> {
        let old_offset = usize::try_from(location.fc)
            .map_err(|_| Error::Limit("Table Stream location exceeds usize".into()))?;
        let old_len = usize::try_from(location.lcb)
            .map_err(|_| Error::Limit("Table Stream length exceeds usize".into()))?;
        let old_end = old_offset
            .checked_add(old_len)
            .ok_or_else(|| Error::Limit("Table Stream location end overflow".into()))?;
        if old_end > self.original_len {
            return Ok(None);
        }
        if !self.changed_layout {
            return Ok(Some(location));
        }
        if let Some(replacement) = self
            .replacements
            .iter()
            .find(|value| value.old_offset == old_offset && value.old_len == old_len)
        {
            return Ok(Some(FibFcLcb {
                fc: u32::try_from(replacement.new_offset)
                    .map_err(|_| Error::Limit("relocated Table offset exceeds u32".into()))?,
                lcb: u32::try_from(replacement.new_len)
                    .map_err(|_| Error::Limit("relocated Table length exceeds u32".into()))?,
            }));
        }
        for replacement in &self.replacements {
            let replacement_end = replacement.old_offset + replacement.old_len;
            if old_offset < replacement_end && replacement.old_offset < old_end {
                return Err(Error::invalid(
                    old_offset as u64,
                    "an FIB Table range partially overlaps a relocated structure",
                ));
            }
        }
        let delta = self
            .replacements
            .iter()
            .take_while(|value| value.old_offset + value.old_len <= old_offset)
            .fold(0i64, |delta, value| {
                delta + value.new_len as i64 - value.old_len as i64
            });
        let new_offset = i64::try_from(old_offset)
            .ok()
            .and_then(|offset| offset.checked_add(delta))
            .and_then(|offset| u32::try_from(offset).ok())
            .ok_or_else(|| Error::Limit("relocated Table offset exceeds u32".into()))?;
        Ok(Some(FibFcLcb {
            fc: new_offset,
            lcb: location.lcb,
        }))
    }
}

fn patch_located<T, S: PatchSink + ?Sized>(
    target: &mut S,
    located: &DocLocated<T>,
    encode: impl Fn(&T) -> Result<Vec<u8>>,
    label: &str,
) -> Result<()> {
    patch_location(target, located.location, encode(&located.value)?, label)
}

fn patch_optional_located<T, S: PatchSink + ?Sized>(
    target: &mut S,
    located: Option<&DocLocated<T>>,
    encode: impl Fn(&T) -> Result<Vec<u8>>,
    label: &str,
) -> Result<()> {
    if let Some(located) = located {
        patch_located(target, located, encode, label)?;
    }
    Ok(())
}

fn patch_part_tables<T, S: PatchSink + ?Sized>(
    target: &mut S,
    tables: &BTreeMap<TextboxDocumentPart, DocLocated<T>>,
    encode: impl Fn(&T) -> Result<Vec<u8>>,
    label: &str,
) -> Result<()> {
    for table in tables.values() {
        patch_located(target, table, &encode, label)?;
    }
    Ok(())
}

fn patch_location<S: PatchSink + ?Sized>(
    target: &mut S,
    location: FibFcLcb,
    encoded: Vec<u8>,
    label: &str,
) -> Result<()> {
    patch_at(
        target,
        usize::try_from(location.fc)
            .map_err(|_| Error::Limit(format!("{label} offset exceeds usize")))?,
        usize::try_from(location.lcb)
            .map_err(|_| Error::Limit(format!("{label} length exceeds usize")))?,
        encoded,
        label,
    )
}

fn patch_prefix<S: PatchSink + ?Sized>(
    target: &mut S,
    expected: usize,
    encoded: Vec<u8>,
    label: &str,
) -> Result<()> {
    patch_at(target, 0, expected, encoded, label)
}

fn patch_at<S: PatchSink + ?Sized>(
    target: &mut S,
    offset: usize,
    expected: usize,
    encoded: Vec<u8>,
    label: &str,
) -> Result<()> {
    target.replace(offset, expected, encoded, label)
}

fn ensure_stream_limit(label: &str, bytes: &[u8], limits: Limits) -> Result<()> {
    if bytes.len() as u64 > limits.max_stream_size {
        return Err(Error::Limit(format!(
            "{label} stream length {} exceeds {}",
            bytes.len(),
            limits.max_stream_size
        )));
    }
    Ok(())
}

fn ensure_entry_limit(label: &str, count: usize, limits: Limits) -> Result<()> {
    if count > limits.max_entries {
        return Err(Error::Limit(format!(
            "{label} count {count} exceeds {}",
            limits.max_entries
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_layout_rebuilds_ranges_and_relocates_fib_coordinates() {
        let mut layout = TableLayout::new(b"abcdefghij".to_vec());
        layout.replace(2, 2, b"XYZ".to_vec(), "grow").unwrap();
        layout.replace(7, 1, Vec::new(), "remove").unwrap();
        let (bytes, relocation) = layout.finish().unwrap();
        assert_eq!(bytes, b"abXYZefgij");
        assert_eq!(
            relocation.relocate(FibFcLcb { fc: 2, lcb: 2 }).unwrap(),
            Some(FibFcLcb { fc: 2, lcb: 3 })
        );
        assert_eq!(
            relocation.relocate(FibFcLcb { fc: 4, lcb: 2 }).unwrap(),
            Some(FibFcLcb { fc: 5, lcb: 2 })
        );
        assert_eq!(
            relocation.relocate(FibFcLcb { fc: 7, lcb: 1 }).unwrap(),
            Some(FibFcLcb { fc: 8, lcb: 0 })
        );
        assert!(relocation.relocate(FibFcLcb { fc: 1, lcb: 2 }).is_err());
    }

    #[test]
    fn same_size_table_layout_preserves_compatibility_overlaps() {
        let mut layout = TableLayout::new(b"abcdef".to_vec());
        layout.replace(2, 2, b"XY".to_vec(), "same size").unwrap();
        let (bytes, relocation) = layout.finish().unwrap();
        assert_eq!(bytes, b"abXYef");
        let overlapping = FibFcLcb { fc: 1, lcb: 4 };
        assert_eq!(relocation.relocate(overlapping).unwrap(), Some(overlapping));
    }

    #[test]
    fn cp_replacement_relocates_boundaries_and_rejects_interior_references() {
        let replacement = CpReplacement::new(3, 6, 1).unwrap();
        assert_eq!(replacement.relocate_u32(2, "test").unwrap(), 2);
        assert_eq!(replacement.relocate_u32(3, "test").unwrap(), 3);
        assert!(replacement.relocate_u32(4, "test").is_err());
        assert_eq!(replacement.relocate_u32(6, "test").unwrap(), 4);
        assert_eq!(replacement.relocate_u32(10, "test").unwrap(), 8);

        let insertion = CpReplacement::new(3, 3, 2).unwrap();
        assert_eq!(insertion.relocate_u32(2, "test").unwrap(), 2);
        assert_eq!(insertion.relocate_u32(3, "test").unwrap(), 5);
    }

    #[test]
    fn malformed_fkp_order_is_strictly_rejected_and_compatibly_diagnosed() {
        let pages = vec![DocFkpPage {
            page: FkpPageNumber {
                page_number: 2,
                unused: 0,
            },
            value: vec![100, 100],
        }];
        let mut diagnostics = Vec::new();
        assert!(
            validate_fkp_page_order(
                &pages,
                Vec::as_slice,
                "test FKP rgfc",
                ParseOptions::default(),
                &mut diagnostics,
            )
            .is_err()
        );
        assert!(diagnostics.is_empty());

        validate_fkp_page_order(
            &pages,
            Vec::as_slice,
            "test FKP rgfc",
            ParseOptions::compatible(Limits::default()),
            &mut diagnostics,
        )
        .unwrap();
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(
            diagnostics[0].code,
            ParseDiagnosticCode::NonconformingRecord
        );
    }

    #[test]
    fn text_relocation_maps_source_fc_through_character_edit() {
        let relocation = TextRelocation {
            source_start: 100,
            source_len: 10,
            source_width: 1,
            destination_start: 200,
            destination_width: 2,
            source_character_count: 10,
            destination_character_count: 11,
            character_replacements: vec![CpReplacement::new(3, 4, 2).unwrap()],
        };
        let mut positions = [100, 103, 104, 110];
        relocate_text_file_positions(&mut positions, &[relocation]).unwrap();
        assert_eq!(positions, [200, 206, 210, 222]);
    }

    #[test]
    fn text_relocation_composes_multiple_character_edits() {
        let relocation = TextRelocation {
            source_start: 100,
            source_len: 10,
            source_width: 1,
            destination_start: 200,
            destination_width: 1,
            source_character_count: 10,
            destination_character_count: 11,
            character_replacements: vec![
                CpReplacement::new(2, 3, 3).unwrap(),
                CpReplacement::new(8, 10, 1).unwrap(),
            ],
        };
        let mut positions = [100, 102, 103, 106, 108, 110];
        relocate_text_file_positions(&mut positions, &[relocation]).unwrap();
        assert_eq!(positions, [200, 202, 205, 208, 209, 211]);
    }

    #[test]
    fn character_run_edits_inherit_the_start_run_and_remove_inner_boundaries() {
        let property = |value| {
            Some(GrpPrl {
                properties: vec![super::super::Prl {
                    sprm: super::super::Sprm::from_opcode(0x0835),
                    operand: SprmOperand::Toggle(value),
                }],
            })
        };
        let mut runs = vec![
            DocChpxRun {
                cp_start: 0,
                cp_end: 3,
                properties: property(1),
            },
            DocChpxRun {
                cp_start: 3,
                cp_end: 6,
                properties: property(0),
            },
            DocChpxRun {
                cp_start: 6,
                cp_end: 10,
                properties: None,
            },
        ];
        apply_character_run_edit(&mut runs, &CpReplacement::new(2, 8, 2).unwrap()).unwrap();
        assert_eq!(
            runs.iter()
                .map(|run| (run.cp_start, run.cp_end, run.properties.clone()))
                .collect::<Vec<_>>(),
            vec![(0, 4, property(1)), (4, 6, None)]
        );

        let mut insertion_runs = vec![
            DocChpxRun {
                cp_start: 0,
                cp_end: 3,
                properties: property(1),
            },
            DocChpxRun {
                cp_start: 3,
                cp_end: 6,
                properties: property(0),
            },
        ];
        apply_character_run_edit(&mut insertion_runs, &CpReplacement::new(3, 3, 2).unwrap())
            .unwrap();
        assert_eq!(
            insertion_runs
                .iter()
                .map(|run| (run.cp_start, run.cp_end, run.properties.clone()))
                .collect::<Vec<_>>(),
            vec![(0, 3, property(1)), (3, 8, property(0))]
        );
    }

    #[test]
    fn paragraph_terminator_inventory_distinguishes_text_from_structure() {
        let compressed = TextPieceCharacters::Compressed(vec![b'A', 0x0d, 0x07, b'B']);
        assert_eq!(paragraph_terminators(&compressed), vec![0x000d, 0x0007]);
        assert_eq!(
            non_paragraph_terminators(&paragraph_terminators(&compressed)),
            vec![0x0007]
        );
        assert_eq!(
            paragraph_terminators_in_range(&compressed, 1, 3).unwrap(),
            vec![0x000d, 0x0007]
        );
        let utf16 = TextPieceCharacters::Utf16(vec![0x4e2d, 0x000c, 0x6587]);
        assert_eq!(paragraph_terminators(&utf16), vec![0x000c]);
        assert!(paragraph_terminators_in_range(&utf16, 0, 4).is_err());
    }

    #[test]
    fn direct_paragraph_properties_follow_prcdata_and_stop_the_source_array() {
        let byte = |sprm: KnownSprm, value: u8| super::super::Prl {
            sprm: super::super::Sprm::from_opcode(sprm.opcode()),
            operand: SprmOperand::Byte(value),
        };
        let dword = |sprm: KnownSprm, value: u32| super::super::Prl {
            sprm: super::super::Sprm::from_opcode(sprm.opcode()),
            operand: SprmOperand::Dword(value.to_le_bytes()),
        };
        let inner = GrpPrl {
            properties: vec![
                dword(KnownSprm::PItap, 1),
                byte(KnownSprm::PFKeep, 1),
                byte(KnownSprm::PFKeepFollow, 1),
            ],
        };
        let outer = GrpPrl {
            properties: vec![
                byte(KnownSprm::PFInTable, 1),
                dword(KnownSprm::PTableProps, 20),
                byte(KnownSprm::PFPageBreakBefore, 1),
            ],
        };
        let data = DocDataStream {
            physical_bytes: Vec::new(),
            nodes: vec![
                DocDataNode {
                    offset: 4,
                    physical_len: outer.to_bytes().unwrap().len() + 2,
                    value: DocDataNodeValue::ParagraphProperties(PrcData { properties: outer }),
                },
                DocDataNode {
                    offset: 20,
                    physical_len: inner.to_bytes().unwrap().len() + 2,
                    value: DocDataNodeValue::ParagraphProperties(PrcData { properties: inner }),
                },
            ],
        };
        let root = GrpPrl {
            properties: vec![dword(KnownSprm::PHugePapx, 4)],
        };

        let applied = expand_direct_paragraph_properties(&root, Some(&data), Some(0)).unwrap();
        assert_eq!(
            applied
                .properties
                .iter()
                .map(|property| property.sprm.kind())
                .collect::<Vec<_>>(),
            vec![
                SprmKind::Known(KnownSprm::PFInTable),
                SprmKind::Known(KnownSprm::PItap),
                SprmKind::Known(KnownSprm::PFKeep),
                SprmKind::Known(KnownSprm::PFKeepFollow),
            ]
        );
        assert_eq!(
            DocDirectParagraphFormatting {
                style_index: 0,
                papx_properties: root.clone(),
                piece_properties: GrpPrl {
                    properties: Vec::new(),
                },
                applied_properties: applied,
            }
            .table_state()
            .unwrap(),
            DocDirectTableState {
                in_table: true,
                depth: 1,
                depth_is_explicit: true,
            }
        );
        assert!(expand_direct_paragraph_properties(&root, Some(&data), Some(1)).is_err());
        assert!(expand_direct_paragraph_properties(&root, None, Some(0)).is_err());
    }
}
