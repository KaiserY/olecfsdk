//! BIFF8 workbook stream framing and initial static record types.

pub mod formula;

pub use formula::{
    BiffConstant, FormulaArray, FormulaMemExtra, FormulaOperator, FormulaRange, FormulaToken,
    FormulaTokenData, FormulaTokenStream,
};

use std::{
    collections::BTreeSet,
    io::{Cursor, Read, Seek, Write},
};

use emfsdk::{DeviceIndependentBitmap, DibColorUsage};

use crate::{
    Error, Result, SdkEnum, SdkObject,
    io::{Reader, SdkEnumValue, SdkRead, SdkSize, SdkWrite, Writer},
    limits::Limits,
    office_art::{OfficeArtPartialStream, OfficeArtRecordHeader, OfficeArtStream},
};

pub const MAX_BIFF_RECORD_DATA: usize = 8224;

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const FORMULA: u16 = 0x0006;
const EXTERN_SHEET: u16 = 0x0017;
const EXTERN_NAME: u16 = 0x0023;
const HEADER: u16 = 0x0014;
const FOOTER: u16 = 0x0015;
const VERTICAL_PAGE_BREAKS: u16 = 0x001a;
const HORIZONTAL_PAGE_BREAKS: u16 = 0x001b;
const NAME: u16 = 0x0018;
const NOTE: u16 = 0x001c;
const SELECTION: u16 = 0x001d;
const FILE_PASS: u16 = 0x002f;
const FONT: u16 = 0x0031;
const FONT30_COMPATIBILITY: u16 = 0x0030;
const WINDOW1: u16 = 0x003d;
const CONTINUE: u16 = 0x003c;
const PANE: u16 = 0x0041;
const CODE_PAGE: u16 = 0x0042;
const PLS: u16 = 0x004d;
const DCON: u16 = 0x0050;
const DCON_REF: u16 = 0x0051;
const DCONN: u16 = 0x0876;
const TXT_QRY: u16 = 0x0805;
const QSI_SX_TAG: u16 = 0x0802;
const SX_VIEW_EX9: u16 = 0x0810;
const DB_QUERY_EXT: u16 = 0x0803;
const HLINK_TOOLTIP: u16 = 0x0800;
const CONTINUE_FRT12: u16 = 0x087f;
const SX_ADDL: u16 = 0x0864;
const ENT_EX_U2: u16 = 0x01c2;
const BK_HIM: u16 = 0x00e9;
const IM_DATA: u16 = 0x007f;
const REAL_TIME_DATA: u16 = 0x0813;
const SORT: u16 = 0x0090;
const LH_RECORD: u16 = 0x0094;
const SORT_DATA: u16 = 0x0895;
const AUTO_FILTER: u16 = 0x009e;
const SX_FORMAT: u16 = 0x00fb;
const W_OPT: u16 = 0x080b;
const TABLE: u16 = 0x0236;
const EXTERN_COUNT: u16 = 0x0016;
const FORMULA4: u16 = 0x0406;
const QSI: u16 = 0x01ad;
const PARAM_QRY: u16 = 0x00dc;
const SX_SELECT: u16 = 0x00f7;
const CRN_COUNT: u16 = 0x0059;
const CRN: u16 = 0x005a;
const FILE_SHARING: u16 = 0x005b;
const WRITE_ACCESS: u16 = 0x005c;
const COL_INFO: u16 = 0x007d;
const GUTS: u16 = 0x0080;
const BOUND_SHEET8: u16 = 0x0085;
const BOUND_SHEET8_2085_COMPATIBILITY: u16 = 0x2085;
const COUNTRY: u16 = 0x008c;
const PALETTE: u16 = 0x0092;
const SCL: u16 = 0x00a0;
const PRINT_SETUP: u16 = 0x00a1;
const SCEN_MAN: u16 = 0x00ae;
const SX_VIEW: u16 = 0x00b0;
const MUL_RK: u16 = 0x00bd;
const MUL_BLANK: u16 = 0x00be;
const MMS: u16 = 0x00c1;
const SX_STREAM_ID: u16 = 0x00d5;
const DB_CELL: u16 = 0x00d7;
const SX_VS: u16 = 0x00e3;
const XF: u16 = 0x00e0;
const XF_E4_COMPATIBILITY: u16 = 0x00e4;
const XF_EE_COMPATIBILITY: u16 = 0x00ee;
const XF_E8E0_COMPATIBILITY: u16 = 0xe8e0;
const OLE_OBJECT_SIZE: u16 = 0x00de;
const MSO_DRAWING_SELECTION: u16 = 0x00ed;
const MERGE_CELLS: u16 = 0x00e5;
const MSO_DRAWING_GROUP: u16 = 0x00eb;
const MSO_DRAWING: u16 = 0x00ec;
const MSO_DRAWING_AC_COMPATIBILITY: u16 = 0x00ac;
const OBJ: u16 = 0x005d;
const OBJ_DC5D_COMPATIBILITY: u16 = 0xdc5d;
const TXO: u16 = 0x01b6;
const SUP_BOOK: u16 = 0x01ae;
const CF: u16 = 0x01b1;
const COND_FMT: u16 = 0x01b0;
const HLINK: u16 = 0x01b8;
const DV: u16 = 0x01be;
const PHONETIC_INFO: u16 = 0x00ef;
const SST: u16 = 0x00fc;
const EXT_SST: u16 = 0x00ff;
const XF_EXT: u16 = 0x087d;
const XF_CRC: u16 = 0x087c;
const HF_PICTURE: u16 = 0x0866;
const FEAT_HDR: u16 = 0x0867;
const FEAT: u16 = 0x0868;
const BOOK_EXT: u16 = 0x0863;
const TABLE_STYLES: u16 = 0x088e;
const STYLE_EXT: u16 = 0x0892;
const DXF: u16 = 0x088d;
const CF_EX: u16 = 0x087b;
const CF12: u16 = 0x087a;
const COND_FMT12: u16 = 0x0879;
const THEME: u16 = 0x0896;
const HEADER_FOOTER_EXT: u16 = 0x089c;
const SHAPE_PROPS_STREAM: u16 = 0x08a4;
const TEXT_PROPS_STREAM: u16 = 0x08a5;
const COMPAT12: u16 = 0x088c;
const PLV: u16 = 0x088b;
const PLV_MAC: u16 = 0x08c8;
const LNEXT: u16 = 0x08c9;
const MKR_EXT: u16 = 0x08ca;
const CRT_CO_OPT: u16 = 0x08cb;
const FRT_ARCH_ID: u16 = 0x08d6;
const CRT_LAYOUT12: u16 = 0x089d;
const CRT_LAYOUT12_A: u16 = 0x08a7;
const MTR_SETTINGS: u16 = 0x089a;
const FORCE_FULL_CALCULATION: u16 = 0x08a3;
const COMPRESS_PICTURES: u16 = 0x089b;
const CRT_ML_FRT: u16 = 0x089e;
const GEL_FRAME: u16 = 0x1066;
const LABEL_SST: u16 = 0x00fd;
const LABEL: u16 = 0x0204;
const SXVI: u16 = 0x00b2;
const SX_IVD: u16 = 0x00b4;
const SX_LI: u16 = 0x00b5;
const SX_PI: u16 = 0x00b6;
const SX_DI: u16 = 0x00c5;
const SX_STRING: u16 = 0x00cd;
const SX_RULE: u16 = 0x00f0;
const SX_EX: u16 = 0x00f1;
const SX_FILT: u16 = 0x00f2;
const SX_DXF: u16 = 0x00f4;
const SX_ITM: u16 = 0x00f5;
const RECALC_ID: u16 = 0x01c1;
const SXVD_EX: u16 = 0x0100;
const SXVD: u16 = 0x00b1;
const CODE_NAME: u16 = 0x01ba;
const ARRAY: u16 = 0x0221;
const USER_SVIEW_BEGIN: u16 = 0x01aa;
const USER_SVIEW_END: u16 = 0x01ab;
const USER_BVIEW: u16 = 0x01a9;
const RR_TAB_ID: u16 = 0x013d;
const SHEET_EXT: u16 = 0x0862;
const CHART_DATA_LABEL_EXT_CONTENTS: u16 = 0x086b;
const CELL_WATCH: u16 = 0x086c;
const FEAT_HDR11: u16 = 0x0871;
const FEATURE11: u16 = 0x0872;
const LIST12: u16 = 0x0877;
const DROP_DOWN_OBJ_IDS: u16 = 0x0874;
const DATA_VALIDATION_HEADER: u16 = 0x01b2;
const RICH_TEXT_STREAM: u16 = 0x08a6;
const GUID_TYPE_LIB: u16 = 0x0897;
const NAME_COMMENT: u16 = 0x0894;
const DIMENSIONS: u16 = 0x0200;
const INDEX: u16 = 0x020b;
const BLANK: u16 = 0x0201;
const NUMBER: u16 = 0x0203;
const BOOL_ERR: u16 = 0x0205;
const STRING_VALUE: u16 = 0x0207;
const ROW: u16 = 0x0208;
const DEFAULT_ROW_HEIGHT: u16 = 0x0225;
const WINDOW2: u16 = 0x023e;
const STYLE: u16 = 0x0293;
const RK: u16 = 0x027e;
const FORMAT: u16 = 0x041e;
const SHARED_FORMULA: u16 = 0x04bc;
const CHART: u16 = 0x1002;
const CHART_DATA_FORMAT: u16 = 0x1006;
const CHART_LINE_FORMAT: u16 = 0x1007;
const CHART_MARKER_FORMAT: u16 = 0x1009;
const CHART_AREA_FORMAT: u16 = 0x100a;
const CHART_PIE_FORMAT: u16 = 0x100b;
const CHART_ATTACHED_LABEL: u16 = 0x100c;
const CHART_SERIES_TEXT: u16 = 0x100d;
const CHART_SERIES: u16 = 0x1003;
const CHART_SERIES_1103_COMPATIBILITY: u16 = 0x1103;
const CHART_FORMAT: u16 = 0x1014;
const CHART_SERIES_LIST: u16 = 0x1016;
const CHART_BAR: u16 = 0x1017;
const CHART_LINE: u16 = 0x1018;
const CHART_PIE: u16 = 0x1019;
const CHART_AREA: u16 = 0x101a;
const CHART_SCATTER: u16 = 0x101b;
const CHART_CRT_LINE: u16 = 0x101c;
const CHART_CRT_LINK: u16 = 0x1022;
const CHART_LEGEND: u16 = 0x1015;
const CHART_AXIS: u16 = 0x101d;
const CHART_TICK: u16 = 0x101e;
const CHART_VALUE_RANGE: u16 = 0x101f;
const CHART_LABEL_RANGE: u16 = 0x1020;
const CHART_AXIS_LINE: u16 = 0x1021;
const CHART_DEFAULT_TEXT: u16 = 0x1024;
const CHART_TEXT: u16 = 0x1025;
const CHART_FONT: u16 = 0x1026;
const CHART_OBJECT_LINK: u16 = 0x1027;
const CHART_FRAME: u16 = 0x1032;
const CHART_3D: u16 = 0x103a;
const CHART_DROP_BAR: u16 = 0x103d;
const CHART_SURF: u16 = 0x103f;
const CHART_LEGEND_EXCEPTION: u16 = 0x1043;
const CHART_AXIS_PARENT: u16 = 0x1041;
const CHART_SHEET_PROPERTIES: u16 = 0x1044;
const CHART_SERIES_GROUP_INDEX: u16 = 0x1045;
const CHART_AXIS_USED: u16 = 0x1046;
const CHART_NUMBER_FORMAT_INDEX: u16 = 0x104e;
const CHART_SERIES_PARENT: u16 = 0x104a;
const CHART_SERIES_AUX_TREND: u16 = 0x104b;
const CHART_POSITION: u16 = 0x104f;
const CHART_FONT_BASIS: u16 = 0x1060;
const CHART_3D_BAR_SHAPE: u16 = 0x105f;
const CHART_SERIES_FORMAT: u16 = 0x105d;
const CHART_SERIES_AUX_ERROR_BAR: u16 = 0x105b;
const CHART_CLRT_CLIENT: u16 = 0x105c;
const CHART_AXIS_OPTIONS: u16 = 0x1062;
const CHART_DAT: u16 = 0x1063;
const CHART_PLOT_GROWTH: u16 = 0x1064;
const CHART_LINKED_DATA: u16 = 0x1051;
const CHART_AL_RUNS: u16 = 0x1050;
const CHART_SERIES_INDEX: u16 = 0x1065;
const START_BLOCK: u16 = 0x0852;
const END_BLOCK: u16 = 0x0853;
const CHART_FRT_INFO: u16 = 0x0850;
const CHART_CAT_LAB: u16 = 0x0856;
const CHART_START_OBJECT: u16 = 0x0854;
const CHART_END_OBJECT: u16 = 0x0855;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BiffStream {
    pub records: Vec<BiffRecord>,
    pub trailing_padding: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BiffRecord {
    pub offset: u32,
    pub data: BiffRecordData,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum BiffRecordData {
    Bof(BofRecord),
    LegacyBof {
        payload: Vec<u8>,
    },
    Eof,
    Formula(FormulaRecord),
    Formula4Compatibility(FormulaRecord),
    SharedFormula(SharedFormulaRecord),
    SupBook(SupBookRecord),
    ConditionalFormatting(ConditionalFormattingRecord),
    ConditionalFormattingGroup(ConditionalFormattingGroupRecord),
    ExternSheet(ExternSheetRecord),
    ExternName(ExternNameRecord),
    Hyperlink(HyperlinkRecord),
    DataValidation(DataValidationRecord),
    Name(NameRecord),
    Pls(PlsRecord),
    MsoDrawingGroup(MsoDrawingRecord),
    MsoDrawing(MsoDrawingRecord),
    Obj(ObjRecord),
    ObjCompatibility {
        record_type: u16,
        value: ObjRecord,
    },
    Txo(TxoRecord),
    Header(HeaderFooterRecord),
    Footer(HeaderFooterRecord),
    VerticalPageBreaks(VerticalPageBreaksRecord),
    HorizontalPageBreaks(HorizontalPageBreaksRecord),
    DCon(DConRecord),
    DConRef(DConRefRecord),
    DConn(DConnRecord),
    TextQuery(TextQuery),
    QsiSxTag(QsiSxTagRecord),
    SxViewEx9(SxViewEx9Record),
    DbQueryExt(DbQueryExtRecord),
    HyperlinkTooltip(HyperlinkTooltipRecord),
    ContinueFrt12(ContinueFrt12Record),
    SxAddl(SxAddlRecord),
    EntExU2(EntExU2Record),
    BkHim(BkHimRecord),
    ImData(ImDataRecord),
    RealTimeData(RealTimeDataRecord),
    Sort(SortRecord),
    LhRecord(LhRecord),
    SortData(SortDataRecord),
    AutoFilter(AutoFilterRecord),
    SxFormat(SxFormatRecord),
    WOpt(WOptRecord),
    Table(TableRecord),
    ExternCount(ExternCountRecord),
    Qsi(QsiRecord),
    ParamQry(ParamQryRecord),
    SxSelect(SxSelectRecord),
    FileSharing(FileSharingRecord),
    OleObjectSize(OleObjectSizeRecord),
    MsoDrawingSelection(MsoDrawingSelectionRecord),
    ScenMan(ScenManRecord),
    SxView(SxViewRecord),
    CodePage {
        code_page: u16,
    },
    BoundSheet8(BoundSheet8Record),
    BoundSheet8Compatibility {
        record_type: u16,
        value: BoundSheet8Record,
    },
    Dimensions(DimensionsRecord),
    Blank(BlankRecord),
    Number(NumberRecord),
    BoolErr(BoolErrRecord),
    Label(LabelRecord),
    Sxvi(SxviRecord),
    SxIvd(SxIvdRecord),
    SxLi(SxLiRecord),
    SxPi(SxPiRecord),
    SxDi(SxDiRecord),
    SxString(SxStringRecord),
    RrTabId(RrTabIdRecord),
    SxRule(SxRuleRecord),
    SxEx(SxExRecord),
    SxFilt(SxFiltRecord),
    SxDxf(SxDxfRecord),
    SxItm(SxItmRecord),
    SxStreamId(SxStreamIdRecord),
    SxVs(SxVsRecord),
    RecalcId(RecalcIdRecord),
    SxvdEx(SxvdExRecord),
    Sxvd(SxvdRecord),
    CodeName(CodeNameRecord),
    Array(ArrayRecord),
    UserSViewBegin(UserSViewBeginRecord),
    UserSViewEnd(UserSViewEndRecord),
    UserBView(UserBViewRecord),
    SheetExt(SheetExtRecord),
    ChartDataLabelExtContents(ChartDataLabelExtContentsRecord),
    CellWatch(CellWatchRecord),
    FeatureHeader11(FeatureHeader11Record),
    Feature11(Feature11Record),
    List12(List12Record),
    DropDownObjIds(DropDownObjIdsRecord),
    DataValidationHeader(DataValidationHeaderRecord),
    RichTextStream(RichTextStreamRecord),
    GuidTypeLib(GuidTypeLibRecord),
    NameComment(NameCommentRecord),
    LabelSst(LabelSstRecord),
    Rk(RkRecord),
    Row(RowRecord),
    Window1(Window1Record),
    Pane(PaneRecord),
    ColInfo(ColInfoRecord),
    Guts(GutsRecord),
    Country(CountryRecord),
    Palette(PaletteRecord),
    Scl(SclRecord),
    PrintSetup(PrintSetupRecord),
    MulRk(MulRkRecord),
    MulBlank(MulBlankRecord),
    Xf(XfRecord),
    XfCompatibility {
        record_type: u16,
        value: XfRecord,
    },
    Crn(CrnRecord),
    XfExt(XfExtRecord),
    XfCrc(XfCrcRecord),
    TableStyles(TableStylesRecord),
    StyleExt(StyleExtRecord),
    Dxf(DxfRecord),
    ConditionalFormattingExtension(ConditionalFormattingExtensionRecord),
    ConditionalFormatting12(ConditionalFormatting12Record),
    ConditionalFormattingGroup12(ConditionalFormattingGroup12Record),
    Theme(ThemeRecord),
    ExtendedHeaderFooter(ExtendedHeaderFooterRecord),
    ShapePropsStream(ShapePropsStreamRecord),
    TextPropsStream(TextPropsStreamRecord),
    Compat12(Compat12Record),
    Plv(PlvRecord),
    PlvMac(PlvMacRecord),
    Lnext(LnextRecord),
    MkrExt(MkrExtRecord),
    CrtCoOpt(CrtCoOptRecord),
    FrtArchId(FrtArchIdRecord),
    CrtLayout12(CrtLayout12Record),
    CrtLayout12A(CrtLayout12ARecord),
    MtrSettings(MtrSettingsRecord),
    ForceFullCalculation(ForceFullCalculationRecord),
    CompressPictures(CompressPicturesRecord),
    CrtMlFrt(CrtMlFrtRecord),
    ChartFrtInfo(ChartFrtInfoRecord),
    ChartCatLab(ChartCatLabRecord),
    ChartStartObject(ChartStartObjectRecord),
    ChartEndObject(ChartEndObjectRecord),
    GelFrame(OfficeArtStream),
    HfPicture(HfPictureRecord),
    FeatureHeader(FeatureHeaderRecord),
    Feature(FeatureRecord),
    BookExt(BookExtRecord),
    Chart(ChartRecord),
    ChartAreaFormat(ChartAreaFormatRecord),
    ChartAttachedLabel(ChartAttachedLabelRecord),
    ChartDataFormat(ChartDataFormatRecord),
    ChartFormat(ChartFormatRecord),
    ChartSeriesList(ChartSeriesListRecord),
    ChartBar(ChartBarRecord),
    ChartLine(ChartLineRecord),
    ChartPie(ChartPieRecord),
    ChartArea(ChartAreaRecord),
    ChartScatter(ChartScatterRecord),
    ChartCrtLine(ChartCrtLineRecord),
    ChartCrtLink(ChartCrtLinkRecord),
    ChartLegend(ChartLegendRecord),
    ChartAxis(ChartAxisRecord),
    ChartTick(ChartTickRecord),
    ChartValueRange(ChartValueRangeRecord),
    ChartLabelRange(ChartLabelRangeRecord),
    ChartAxisLine(ChartAxisLineRecord),
    ChartDefaultText(ChartDefaultTextRecord),
    ChartFont(ChartFontRecord),
    ChartLineFormat(ChartLineFormatRecord),
    ChartMarkerFormat(ChartMarkerFormatRecord),
    ChartObjectLink(ChartObjectLinkRecord),
    ChartFrame(ChartFrameRecord),
    Chart3D(Chart3DRecord),
    ChartDropBar(ChartDropBarRecord),
    ChartSurf(ChartSurfRecord),
    ChartLegendException(ChartLegendExceptionRecord),
    ChartAxisParent(ChartAxisParentRecord),
    ChartSheetProperties(ChartSheetPropertiesRecord),
    ChartSeriesGroupIndex(ChartSeriesGroupIndexRecord),
    ChartAxisUsed(ChartAxisUsedRecord),
    ChartNumberFormatIndex(ChartNumberFormatIndexRecord),
    ChartSeriesParent(ChartSeriesParentRecord),
    ChartSeriesAuxTrend(ChartSeriesAuxTrendRecord),
    ChartPosition(ChartPositionRecord),
    ChartFontBasis(ChartFontBasisRecord),
    Chart3DBarShape(Chart3DBarShapeRecord),
    ChartSeriesFormat(ChartSeriesFormatRecord),
    ChartSeriesAuxErrorBar(ChartSeriesAuxErrorBarRecord),
    ChartClrtClient(ChartClrtClientRecord),
    ChartAxisOptions(ChartAxisOptionsRecord),
    ChartDat(ChartDatRecord),
    ChartPieFormat(ChartPieFormatRecord),
    ChartPlotGrowth(ChartPlotGrowthRecord),
    ChartLinkedData(ChartLinkedDataRecord),
    ChartAlRuns(ChartAlRunsRecord),
    ChartSeriesIndex(ChartSeriesIndexRecord),
    ChartSeries(ChartSeriesRecord),
    ChartSeriesCompatibility {
        record_type: u16,
        value: ChartSeriesRecord,
    },
    ChartSeriesText(ChartSeriesTextRecord),
    ChartText(ChartTextRecord),
    StartBlock(StartBlockRecord),
    EndBlock(EndBlockRecord),
    DbCell(DbCellRecord),
    Font(FontRecord),
    FontCompatibility {
        record_type: u16,
        value: FontRecord,
    },
    Format(FormatRecord),
    Style(StyleRecord),
    CrnCount(CrnCountRecord),
    DefaultRowHeight(DefaultRowHeightRecord),
    WriteAccess(WriteAccessRecord),
    Window2(Window2Record),
    Selection(SelectionRecord),
    MergeCells(MergeCellsRecord),
    Mms(MmsRecord),
    PhoneticInfo(PhoneticInfoRecord),
    Sst(SstRecord),
    ExtSst(ExtSstRecord),
    Index(IndexRecord),
    StringValue(StringValueRecord),
    FixedU16 {
        kind: FixedU16RecordKind,
        value: u16,
    },
    FixedF64Bits {
        kind: FixedF64RecordKind,
        bits: u64,
    },
    Empty {
        kind: EmptyRecordKind,
        reserved: Option<u16>,
    },
    FilePass {
        payload: Vec<u8>,
    },
    Encrypted {
        record_type: u16,
        payload: Vec<u8>,
    },
    Continue {
        payload: Vec<u8>,
    },
    Unknown {
        record_type: u16,
        payload: Vec<u8>,
    },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct BofRecord {
    pub version: u16,
    pub document_type: u16,
    pub build_identifier: u16,
    pub build_year: u16,
    /// Exact flag word, including specification-defined ignored bits.
    pub history_flags: u32,
    /// Exact lowest-version word, including reserved bits.
    pub lowest_version: u32,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BoundSheet8Record {
    pub sheet_bof_offset: u32,
    /// Low two bits are hsState; upper six bits are retained ignored bits.
    pub state: u8,
    pub sheet_type: u8,
    pub name: ShortXlUnicodeString,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ShortXlUnicodeString {
    /// Exact option flags. Bit 0 selects UTF-16; other bits are retained.
    pub flags: u8,
    pub characters: XlStringCharacters,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlStringCharacters {
    Compressed(Vec<u8>),
    Unicode(Vec<u16>),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum HeaderFooterRecord {
    /// The common encoding for an empty header/footer: no payload at all.
    EmptyPayload,
    /// Compatibility encoding containing only a zero character count.
    EmptyCountOnly,
    Text {
        /// Exact option flags. Bit 0 selects UTF-16; other bits are retained.
        flags: u8,
        characters: XlStringCharacters,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormulaRecord {
    pub cell: CellHeader,
    pub cached_result: FormulaCachedResult,
    pub flags: u16,
    pub calculation_chain_id: u32,
    pub tokens: FormulaTokens,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SharedFormulaRecord {
    pub first_row: u16,
    pub last_row: u16,
    pub first_column: u8,
    pub last_column: u8,
    pub reserved: u8,
    pub use_count: u8,
    pub tokens: FormulaTokens,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SupBookRecord {
    pub sheet_count: u16,
    pub link: SupBookLink,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SupBookLink {
    SelfReference,
    AddInFunctions,
    VirtualPath {
        declared_character_count: u16,
        path: BiffUnicodeString,
        sheet_names: Vec<SupBookSheetName>,
        trailing: Vec<u8>,
    },
    Compatibility {
        tag: u16,
        payload: Vec<u8>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SupBookSheetName {
    pub declared_character_count: u16,
    pub name: BiffUnicodeString,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct ExternSheetRecord {
    pub reference_count: u16,
    #[sdk(count = "reference_count")]
    pub references: Vec<ExternSheetReference>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ExternSheetReference {
    pub sup_book_index: u16,
    pub first_sheet_index: u16,
    pub last_sheet_index: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ExternNameFlags: u16 {
        const BUILT_IN = 0x0001;
        const WANT_ADVISE = 0x0002;
        const WANT_PICTURE = 0x0004;
        const DDE_NO_OPERATION = 0x0008;
        const OLE_LINK = 0x0010;
        const CLIPBOARD_FORMAT = 0x7fe0;
        const ICON = 0x8000;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExternNameRecord {
    pub flags: ExternNameFlags,
    pub sheet_index: u16,
    pub reserved: u16,
    pub declared_name_character_count: u8,
    pub name: BiffUnicodeString,
    pub body: ExternNameBody,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ExternNameBody {
    Empty,
    ParsedFormula {
        declared_length: u16,
        value: Option<ExternNameFormulaValue>,
    },
    CachedLinkValues {
        last_column: u8,
        last_row: u16,
        values: Vec<BiffConstant>,
        trailing: Vec<u8>,
    },
    /// Context-selected AddinUdf/OLE compatibility body retained inside a typed envelope.
    Compatibility(Vec<u8>),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExternNameFormulaValue {
    Reference3d {
        sheet_pair: u32,
        row: u16,
        column: u16,
    },
    Area3d {
        sheet_pair: u32,
        first_row: u16,
        last_row: u16,
        first_column: u16,
        last_column: u16,
    },
    DeletedReference3d {
        sheet_pair: u32,
        reserved: u32,
    },
    DeletedArea3d {
        sheet_pair: u32,
        reserved1: u32,
        reserved2: u32,
    },
    Error {
        code: u8,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HyperlinkRecord {
    pub first_row: u16,
    pub last_row: u16,
    pub first_column: u16,
    pub last_column: u16,
    pub class_id: [u8; 16],
    pub object: HyperlinkObject,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct HyperlinkFlags: u32 {
        const HAS_MONIKER = 0x0000_0001;
        const ABSOLUTE = 0x0000_0002;
        const SITE_GAVE_DISPLAY_NAME = 0x0000_0004;
        const HAS_LOCATION = 0x0000_0008;
        const HAS_DISPLAY_NAME = 0x0000_0010;
        const HAS_GUID = 0x0000_0020;
        const HAS_CREATION_TIME = 0x0000_0040;
        const HAS_TARGET_FRAME = 0x0000_0080;
        const MONIKER_SAVED_AS_STRING = 0x0000_0100;
        const ABSOLUTE_FROM_RELATIVE = 0x0000_0200;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum HyperlinkObject {
    Parsed {
        stream_version: u32,
        flags: HyperlinkFlags,
        display_name: Option<HyperlinkString>,
        target_frame_name: Option<HyperlinkString>,
        moniker: Option<Box<HyperlinkMoniker>>,
        location: Option<HyperlinkString>,
        guid: Option<[u8; 16]>,
        creation_time: Option<u64>,
        trailing: Vec<u8>,
    },
    Truncated {
        stream_version: u32,
        flags: HyperlinkFlags,
        payload: Vec<u8>,
    },
    TruncatedUrlMoniker {
        stream_version: u32,
        flags: HyperlinkFlags,
        class_id: [u8; 16],
        declared_byte_length: u32,
        address: Vec<u16>,
    },
    Compatibility(Vec<u8>),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HyperlinkString {
    pub declared_character_count: u32,
    pub characters: Vec<u16>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum HyperlinkMoniker {
    String(HyperlinkString),
    Url {
        class_id: [u8; 16],
        declared_byte_length: u32,
        address: Vec<u16>,
        tail: Vec<u8>,
    },
    File {
        class_id: [u8; 16],
        options: u16,
        declared_short_name_length: u32,
        short_name: Vec<u8>,
        tail: [u8; 24],
        long_path: Option<HyperlinkLongPath>,
    },
    Standard {
        class_id: [u8; 16],
        options: u16,
        declared_data_length: u32,
        data: Vec<u8>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HyperlinkLongPath {
    pub declared_total_length: u32,
    pub declared_character_bytes: u32,
    pub key: u16,
    pub characters: Vec<u16>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct DataValidationOptions {
    pub validation_type: u8,
    pub error_style: u8,
    pub string_lookup: bool,
    pub allow_blank: bool,
    pub suppress_combo: bool,
    pub ime_mode: u8,
    pub show_input_message: bool,
    pub show_error_message: bool,
    pub operator: u8,
    pub reserved: u8,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlUnicodeString {
    pub text: BiffUnicodeString,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataValidationFormula {
    pub unused: u16,
    pub tokens: FormulaTokenStream,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataValidationRecord {
    pub options: DataValidationOptions,
    pub prompt_title: XlUnicodeString,
    pub error_title: XlUnicodeString,
    pub prompt: XlUnicodeString,
    pub error: XlUnicodeString,
    pub formula1: DataValidationFormula,
    pub formula2: DataValidationFormula,
    pub ranges: Vec<CellRange>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DxfFlags: u64 {
        const NUMBER_FORMAT = 1 << 25;
        const FONT = 1 << 26;
        const ALIGNMENT = 1 << 27;
        const BORDER = 1 << 28;
        const PATTERN = 1 << 29;
        const PROTECTION = 1 << 30;
        const USER_NUMBER_FORMAT = 1 << 32;
        const NEW_BORDER = 1 << 34;
        const ZERO_INITIALIZED = 1 << 47;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DxfNumberFormat {
    BuiltIn {
        unused: u8,
        format_index: u8,
    },
    UserDefined {
        declared_size: u16,
        format: XlUnicodeString,
    },
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct DxfFont {
    pub font_name: [u8; 64],
    pub font_height: i32,
    pub font_options: u32,
    pub font_weight: u16,
    pub escapement: u16,
    pub underline_and_unused: u32,
    pub color_index: i32,
    pub reserved: u32,
    pub option_flags: u32,
    pub escapement_modified: u32,
    pub underline_modified: u32,
    pub weight_modified: u32,
    pub unused1: u32,
    pub unused2: u32,
    pub unused3: u32,
    pub formatting_end: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct DxfAlignment {
    pub bits: u64,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct DxfBorder {
    pub styles_and_side_colors: u32,
    pub remaining_colors_and_diagonal: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct DxfPattern {
    pub pattern_style: u16,
    pub color_indexes: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct DxfProtection {
    pub bits: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DxfN {
    pub flags: DxfFlags,
    pub number_format: Option<DxfNumberFormat>,
    pub font: Option<DxfFont>,
    pub alignment: Option<DxfAlignment>,
    pub border: Option<DxfBorder>,
    pub pattern: Option<DxfPattern>,
    pub protection: Option<DxfProtection>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ConditionalFormattingRecord {
    pub condition_type: u8,
    pub comparison_operator: u8,
    pub format: DxfN,
    pub formula1: FormulaTokenStream,
    pub formula2: FormulaTokenStream,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ConditionalFormattingGroupRecord {
    pub rule_count: u16,
    /// Bit 0 is fToughRecalc; bits 1..=15 contain nID.
    pub flags_and_id: u16,
    pub bounds: CellRange,
    pub ranges: Vec<CellRange>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XfExtNoFrt {
    pub reserved1: u16,
    pub reserved2: u16,
    pub reserved3: u16,
    pub properties: Vec<ExtProperty>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DxfN12 {
    Empty {
        reserved: u16,
    },
    Formatting {
        format: Box<DxfN>,
        extension: Option<XfExtNoFrt>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum CfExTemplateParams {
    Filter {
        flags: u8,
        parameter: u16,
        reserved: [u8; 13],
    },
    Text {
        comparison_type: u16,
        reserved: [u8; 14],
    },
    Date {
        operation: u16,
        reserved: [u8; 14],
    },
    Averages {
        parameter: u16,
        reserved: [u8; 14],
    },
    Default {
        unused: [u8; 16],
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CfExNonCf12 {
    pub rule_index: u16,
    pub comparison_operator: u8,
    pub template_id: u8,
    pub priority: u16,
    pub flags: u8,
    pub format: Option<DxfN12>,
    pub declared_template_parameter_size: u8,
    pub template_parameters: CfExTemplateParams,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ConditionalFormattingExtensionRecord {
    pub header: FrtHeader,
    pub is_cf12: u32,
    pub group_id: u16,
    pub content: Option<CfExNonCf12>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CfVo {
    pub value_type: u8,
    pub formula: FormulaTokenStream,
    pub value_bits: Option<u64>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CfGradientInterpolationItem {
    pub value: CfVo,
    pub domain_bits: u64,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CfGradientItem {
    pub range_bits: u64,
    pub color: CfColor,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct CfGradientFlags: u8 {
        const CLAMP = 0x01;
        const BACKGROUND = 0x02;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct CfDataBarFlags: u8 {
        const RIGHT_TO_LEFT = 0x01;
        const SHOW_VALUE = 0x02;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct CfFilterFlags: u8 {
        const TOP = 0x01;
        const PERCENT = 0x02;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct CfMultistateFlags: u8 {
        const ICON_ONLY = 0x01;
        const REVERSE = 0x04;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CfGradient {
    pub unused: u16,
    pub reserved: u8,
    pub interpolation_count: u8,
    pub gradient_count: u8,
    pub flags: CfGradientFlags,
    pub interpolation: Vec<CfGradientInterpolationItem>,
    pub gradient: Vec<CfGradientItem>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CfDataBar {
    pub unused: u16,
    pub reserved: u8,
    pub flags: CfDataBarFlags,
    pub minimum_percent: u8,
    pub maximum_percent: u8,
    pub color: CfColor,
    pub minimum: CfVo,
    pub maximum: CfVo,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CfFilter {
    pub declared_size: u16,
    pub reserved: u8,
    pub flags: CfFilterFlags,
    pub parameter: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CfMultistateItem {
    pub value: CfVo,
    pub equal: u8,
    pub unused: u32,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CfMultistate {
    pub unused: u16,
    pub reserved: u8,
    pub state_count: u8,
    pub icon_set: u8,
    pub flags: CfMultistateFlags,
    pub states: Vec<CfMultistateItem>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Cf12ConditionData {
    None,
    Gradient(CfGradient),
    DataBar(CfDataBar),
    Filter(CfFilter),
    Multistate(CfMultistate),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ConditionalFormatting12Record {
    pub header: FrtRefHeaderU,
    pub condition_type: u8,
    pub comparison_operator: u8,
    pub formula1: FormulaTokenStream,
    pub formula2: FormulaTokenStream,
    pub format: DxfN12,
    pub active_formula: FormulaTokenStream,
    pub option_flags: u8,
    pub priority: u16,
    pub template_id: u16,
    pub template_parameter_size: u8,
    pub template_parameters: Option<CfExTemplateParams>,
    pub condition_data: Cf12ConditionData,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ConditionalFormattingGroup12Record {
    pub header: FrtRefHeaderU,
    pub group: ConditionalFormattingGroupRecord,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FormulaCachedResult {
    /// Exact IEEE-754 bits for a numeric cached result.
    NumberBits(u64),
    Special(FormulaSpecialCachedResult),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FormulaSpecialCachedResult {
    pub kind: u8,
    pub reserved1: u8,
    pub value: u8,
    pub reserved2: [u8; 3],
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormulaTokens {
    pub rgce: FormulaTokenStream,
    /// Bounded producer-specific bytes after all statically parsed Array values.
    pub rgcb_tail: Vec<u8>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct NameFlags: u16 {
        const HIDDEN = 0x0001;
        const FUNCTION = 0x0002;
        const VBA_MACRO = 0x0004;
        const PROCEDURE = 0x0008;
        const CALCULATION_EXPRESSION = 0x0010;
        const BUILT_IN = 0x0020;
        const PUBLISHED = 0x2000;
        const WORKBOOK_PARAMETER = 0x4000;
    }
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct NameStringFlags: u8 {
        const HIGH_BYTE = 0x01;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NameRecord {
    pub flags: NameFlags,
    pub keyboard_shortcut: u8,
    pub declared_name_character_count: u8,
    pub declared_formula_byte_count: u16,
    pub reserved3: u16,
    pub sheet_index: u16,
    pub custom_menu: NameTrailingText,
    pub description: NameTrailingText,
    pub help_topic: NameTrailingText,
    pub status_bar: NameTrailingText,
    pub name_flags: NameStringFlags,
    pub name: NameValue,
    pub formula: FormulaTokenStream,
    pub formula_extra_tail: Vec<u8>,
    pub physical_segment_lengths: Vec<u16>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum NameValue {
    BuiltIn(u8),
    User(XlStringCharacters),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NameTrailingText {
    pub declared_character_count: u8,
    pub characters: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PlsRecord {
    pub reserved: u16,
    pub settings: PrinterSettings,
    pub physical_segment_lengths: Vec<u16>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MsoDrawingRecord {
    pub data: MsoDrawingData,
    pub physical_segments: Vec<MsoDrawingSegment>,
    /// BIFF Obj/TxO/Note records interleaved at OfficeArt client-data boundaries.
    pub host_records: Vec<MsoDrawingHostRecord>,
    /// First BIFF record outside this aggregate, when one was present.
    pub following_record_type: Option<u16>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct MsoDrawingSegment {
    pub record_type: u16,
    pub payload_length: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MsoDrawingHostRecord {
    /// Number of OfficeArt physical segments emitted before this host record.
    pub after_segment: usize,
    pub data: MsoDrawingHostData,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum MsoDrawingHostData {
    Obj(ObjRecord),
    ObjCompatibility { record_type: u16, value: ObjRecord },
    Txo(TxoRecord),
    Note(NoteRecord),
    Raw { record_type: u16, payload: Vec<u8> },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ObjRecord {
    pub subrecords: Vec<ObjSubrecord>,
    pub trailing: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ObjSubrecord {
    pub subrecord_type: u16,
    pub declared_length: u16,
    pub data: ObjSubrecordData,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ObjSubrecordData {
    End,
    Common(ObjCommonData),
    Macro(ObjFormula),
    GroupMarker {
        unused: u16,
    },
    ClipboardFormat {
        format: u16,
    },
    PictureFlags(ObjPictureFlags),
    /// One-byte FtPioGrbit found in a damaged legacy fixture.
    TruncatedPictureFlags {
        low_byte: u8,
    },
    /// Zero-length nonstandard marker found in a damaged standalone Obj fixture.
    EmptyCompatibilityMarker,
    PictureFormula(ObjPictureFormula),
    CheckBox(ObjCheckBoxStructure),
    RadioButton {
        unused1: u32,
        unused2: u16,
    },
    ScrollBar(ObjScrollBarData),
    Note(ObjNoteData),
    ScrollBarFormula(ObjFormula),
    GroupBox(ObjGroupBoxData),
    EditBox(ObjEditBoxData),
    RadioButtonData(ObjRadioButtonData),
    CheckBoxData(ObjCheckBoxData),
    ListBox(ObjListBoxData),
    CheckBoxFormula(ObjFormula),
    /// Producer-specific but structurally bounded subrecord retained as typed words.
    Compatibility(ObjCompatibilityData),
    Raw(Vec<u8>),
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ObjCommonFlags: u16 {
        const LOCKED = 0x0001;
        const DEFAULT_SIZE = 0x0004;
        const PUBLISHED = 0x0008;
        const PRINTABLE = 0x0010;
        const DISABLED = 0x0080;
        const UI_OBJECT = 0x0100;
        const RECALCULATE_ON_LOAD = 0x0200;
        const RECALCULATE_ALWAYS = 0x1000;
        const AUTO_LINE = 0x2000;
        const AUTO_FILL = 0x4000;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ObjCommonData {
    pub object_type: u16,
    pub object_id: u16,
    pub flags: ObjCommonFlags,
    pub reserved1: u32,
    pub reserved2: u32,
    pub reserved3: u32,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ObjPictureFlags: u16 {
        const AUTO_PICTURE = 0x0001;
        const DDE = 0x0002;
        const PRINT_CALCULATE = 0x0004;
        const ICON = 0x0008;
        const CONTROL = 0x0010;
        const CONTROL_STREAM = 0x0020;
        const CAMERA = 0x0080;
        const DEFAULT_SIZE = 0x0100;
        const AUTO_LOAD = 0x0200;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ObjPictureFormula {
    pub formula: ObjFormula,
    /// Context-selected PictFmlaEmbedInfo/PictFmlaKey fields not yet expanded.
    pub trailing: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ObjCheckBoxStructure {
    Legacy {
        unused1: u32,
        unused2: u32,
    },
    Full {
        unused1: u32,
        unused2: u32,
        unused3: u32,
    },
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ObjScrollBarFlags: u16 {
        const DRAW = 0x0001;
        const DRAW_SLIDER_ONLY = 0x0002;
        const TRACK_ELEVATOR = 0x0004;
        const NO_3D = 0x0008;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ObjScrollBarData {
    pub unused: u32,
    pub value: i16,
    pub minimum: i16,
    pub maximum: i16,
    pub increment: i16,
    pub page_increment: i16,
    pub horizontal: u16,
    pub width: i16,
    pub flags: ObjScrollBarFlags,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ObjNoteData {
    pub guid: [u8; 16],
    pub shared: u16,
    pub unused: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ObjGroupBoxData {
    pub accelerator: u16,
    pub reserved: u16,
    pub flags: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ObjEditBoxData {
    pub validation_type: u16,
    pub multiline: u16,
    pub vertical_scroll: u16,
    pub list_object_id: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ObjRadioButtonData {
    pub next_object_id: u16,
    pub first_button: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ObjCheckBoxData {
    pub checked: u16,
    pub accelerator: u16,
    pub reserved: u16,
    pub flags: u16,
    pub trailing: Vec<u8>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ObjListBoxFlags: u16 {
        const USE_CALLBACK = 0x0001;
        const VALID_STRING_ARRAY = 0x0002;
        const VALID_EDIT_ID = 0x0004;
        const NO_3D = 0x0008;
        const SELECTION_TYPE_0 = 0x0010;
        const SELECTION_TYPE_1 = 0x0020;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ObjListBoxData {
    /// Exact cbFContinued word, including the common AutoFilter value 0x1FEE.
    pub continued_size: u16,
    pub formula: ObjFormula,
    pub line_count: u16,
    pub selected_index: u16,
    pub flags: ObjListBoxFlags,
    pub edit_object_id: u16,
    pub drop_data: Option<ObjListDropData>,
    pub lines: Vec<ObjListString>,
    pub selections: Vec<u8>,
    pub trailing: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ObjListDropData {
    pub style: u16,
    pub line_count: u16,
    pub minimum_width: u16,
    pub declared_character_count: u16,
    pub text: BiffUnicodeString,
    pub unused: Option<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ObjListString {
    pub declared_character_count: u16,
    pub text: BiffUnicodeString,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ObjCompatibilityData {
    pub words: Vec<u16>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TxoRecord {
    pub options: TxoOptions,
    pub rotation: TxoRotation,
    pub context: TxoContext,
    pub declared_text_length: u16,
    pub declared_run_data_length: u16,
    pub empty_font_index: u16,
    pub formula: ObjFormula,
    pub trailing: Vec<u8>,
    pub text_chunks: Vec<BiffUnicodeString>,
    pub runs: Vec<TxoRun>,
    pub last_run: Option<TxoRun>,
    pub formatting_segment_lengths: Vec<u16>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct TxoOptions: u16 {
        const HORIZONTAL_ALIGNMENT = 0x000e;
        const VERTICAL_ALIGNMENT = 0x0070;
        const LOCK_TEXT = 0x0200;
        const JUSTIFY_LAST_LINE = 0x4000;
        const SECRET_EDIT = 0x8000;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TxoRotation(pub u16);

impl TxoRotation {
    pub const NONE: Self = Self(0);
    pub const STACKED: Self = Self(1);
    pub const COUNTER_CLOCKWISE_90: Self = Self(2);
    pub const CLOCKWISE_90: Self = Self(3);
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TxoContext {
    Reserved {
        reserved4: u16,
        reserved5: u32,
    },
    Control(ObjControlInfo),
    /// No usable preceding FtCmo was available; exact bounded words remain explicit.
    Undetermined {
        word1: u16,
        word2: u16,
        word3: u16,
    },
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ObjControlInfoFlags: u16 {
        const DEFAULT = 0x0001;
        const HELP = 0x0002;
        const CANCEL = 0x0004;
        const DISMISS = 0x0008;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ObjControlInfo {
    pub flags: ObjControlInfoFlags,
    pub accelerator: i16,
    pub reserved: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ObjFormula {
    pub declared_length: u16,
    pub data: ObjFormulaData,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ObjFormulaData {
    Empty,
    Parsed {
        cce_and_reserved: u16,
        unused: u32,
        tokens: FormulaTokenStream,
        padding: Vec<u8>,
    },
    Opaque(Vec<u8>),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TxoRun {
    pub format: FormatRun,
    pub unused1: u16,
    pub unused2: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NoteRecord {
    pub row: u16,
    pub column: u16,
    pub flags: NoteFlags,
    pub object_id: u16,
    pub declared_author_length: u16,
    pub author: BiffUnicodeString,
    pub unused: u8,
    pub trailing: Vec<u8>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct NoteFlags: u16 {
        const SHOW = 0x0002;
        const ROW_HIDDEN = 0x0080;
        const COLUMN_HIDDEN = 0x0100;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum MsoDrawingData {
    Complete(OfficeArtStream),
    Partial(OfficeArtPartialStream),
    Incomplete { bytes: Vec<u8>, reason: String },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PrinterSettings {
    WindowsUnicode(DevModeW),
    LengthPrefixedWindowsUnicode {
        declared_length: u16,
        devmode: DevModeW,
    },
    MacXmlPlist(Vec<u8>),
    MacPrintRecord([u8; 120]),
    LegacyPageLayout([u32; 7]),
    PlatformSpecific(Vec<u8>),
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DevModeFields: u32 {
        const ORIENTATION = 0x0000_0001;
        const PAPER_SIZE = 0x0000_0002;
        const PAPER_LENGTH = 0x0000_0004;
        const PAPER_WIDTH = 0x0000_0008;
        const SCALE = 0x0000_0010;
        const POSITION = 0x0000_0020;
        const NUP = 0x0000_0040;
        const DISPLAY_ORIENTATION = 0x0000_0080;
        const COPIES = 0x0000_0100;
        const DEFAULT_SOURCE = 0x0000_0200;
        const PRINT_QUALITY = 0x0000_0400;
        const COLOR = 0x0000_0800;
        const DUPLEX = 0x0000_1000;
        const Y_RESOLUTION = 0x0000_2000;
        const TT_OPTION = 0x0000_4000;
        const COLLATE = 0x0000_8000;
        const FORM_NAME = 0x0001_0000;
        const LOG_PIXELS = 0x0002_0000;
        const BITS_PER_PEL = 0x0004_0000;
        const PELS_WIDTH = 0x0008_0000;
        const PELS_HEIGHT = 0x0010_0000;
        const DISPLAY_FLAGS = 0x0020_0000;
        const DISPLAY_FREQUENCY = 0x0040_0000;
        const ICM_METHOD = 0x0080_0000;
        const ICM_INTENT = 0x0100_0000;
        const MEDIA_TYPE = 0x0200_0000;
        const DITHER_TYPE = 0x0400_0000;
        const PANNING_WIDTH = 0x0800_0000;
        const PANNING_HEIGHT = 0x1000_0000;
        const DISPLAY_FIXED_OUTPUT = 0x2000_0000;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DevModeW {
    pub device_name: [u16; 32],
    pub specification_version: u16,
    pub driver_version: u16,
    pub declared_public_size: u16,
    pub declared_driver_extra_size: u16,
    pub fields: DevModeFields,
    pub public_fields: DevModeWPublic,
    /// Printer-driver-defined data, bounded by dmDriverExtra.
    pub driver_extra: Vec<u8>,
    pub driver_extra_complete: bool,
    pub trailing: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DevModeWPublic {
    Full(Box<DevModeWFull>),
    Legacy212(Box<DevModeWLegacy212>),
    Core100(DevModeWCore100),
    Truncated(Vec<u8>),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DevModeWCore100 {
    pub orientation: u16,
    pub paper_size: u16,
    pub paper_length: u16,
    pub paper_width: u16,
    pub scale: u16,
    pub copies: u16,
    pub default_source: u16,
    pub print_quality: u16,
    pub color: u16,
    pub duplex: u16,
    pub y_resolution: u16,
    pub tt_option: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DevModeWLegacy212 {
    pub core: DevModeWCore100,
    pub collate: u16,
    pub form_name: [u16; 32],
    pub log_pixels: u16,
    pub bits_per_pel: u32,
    pub pels_width: u32,
    pub pels_height: u32,
    pub display_flags_or_nup: u32,
    pub display_frequency: u32,
    pub icm_method: u32,
    pub icm_intent: u32,
    pub media_type: u32,
    pub dither_type: u32,
    pub reserved1: u32,
    pub reserved2: u32,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DevModeWFull {
    pub orientation: u16,
    pub paper_size: u16,
    pub paper_length: u16,
    pub paper_width: u16,
    pub scale: u16,
    pub copies: u16,
    pub default_source: u16,
    pub print_quality: u16,
    pub color: u16,
    pub duplex: u16,
    pub y_resolution: u16,
    pub tt_option: u16,
    pub collate: u16,
    pub form_name: [u16; 32],
    pub log_pixels: u16,
    pub bits_per_pel: u32,
    pub pels_width: u32,
    pub pels_height: u32,
    pub display_flags_or_nup: u32,
    pub display_frequency: u32,
    pub icm_method: u32,
    pub icm_intent: u32,
    pub media_type: u32,
    pub dither_type: u32,
    pub reserved1: u32,
    pub reserved2: u32,
    pub panning_width: u32,
    pub panning_height: u32,
    pub public_extension: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FixedU16RecordKind {
    CalcCount,
    CalcMode,
    CalcPrecision,
    RefMode,
    Iteration,
    Protect,
    Password,
    WindowProtect,
    Date1904,
    PrintHeaders,
    PrintGridlines,
    Backup,
    DefaultColWidth,
    Uncalced,
    SaveRecalc,
    ObjectProtect,
    Gridset,
    HCenter,
    VCenter,
    HideObj,
    FnGroupCount,
    AutoFilterInfo,
    BookBool,
    ScenarioProtect,
    InterfaceHdr,
    UseSelFs,
    Dsf,
    ProtectionRev4,
    RefreshAll,
    PasswordRev4,
    ChartUnits,
    WsBool,
    PrintSize,
    StandardWidth,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FixedF64RecordKind {
    CalcDelta,
    LeftMargin,
    RightMargin,
    TopMargin,
    BottomMargin,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EmptyRecordKind {
    /// Zero-length record type 0 emitted between margin records by one legacy producer.
    NullCompatibility,
    WriteProtect,
    InterfaceEnd,
    Template,
    FilterMode,
    ObjectProject,
    VbaProjectEmpty,
    Excel9File,
    ChartBegin,
    ChartEnd,
    ChartPlotArea,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CellHeader {
    pub row: u16,
    pub column: u16,
    pub format_index: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct DimensionsRecord {
    pub first_row: u32,
    pub last_row_exclusive: u32,
    pub first_column: u16,
    pub last_column_exclusive: u16,
    pub reserved: u16,
    /// POI-61045 compatibility field found after the specified payload.
    pub compatibility_extra: Option<u16>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct BlankRecord {
    pub cell: CellHeader,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct NumberRecord {
    pub cell: CellHeader,
    /// Exact IEEE-754 bits.
    pub value_bits: u64,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct BoolErrRecord {
    pub cell: CellHeader,
    pub value: BoolErrValue,
    pub is_error: u8,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BoolErrValue {
    Byte(u8),
    /// Compatibility form used by some producers and accepted by Apache POI.
    Word(u16),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct LabelSstRecord {
    pub cell: CellHeader,
    pub shared_string_index: u32,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LabelRecord {
    pub cell: CellHeader,
    pub text: BiffUnicodeString,
    /// Compatibility UTF-16 NUL emitted by some producers after the declared text.
    pub trailing_null: Option<u16>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct RkRecord {
    pub cell: CellHeader,
    /// Encoded RK number including integer/divide-by-100 flag bits.
    pub value: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct RowRecord {
    pub row: u16,
    pub first_column: u16,
    pub last_column_exclusive: u16,
    pub height: u16,
    pub reserved1: u16,
    pub reserved2: u16,
    /// Exact row flags, outline level and format index bit field.
    pub flags: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct Window1Record {
    pub horizontal_position: u16,
    pub vertical_position: u16,
    pub width: u16,
    pub height: u16,
    pub flags: u16,
    pub active_sheet: u16,
    pub first_visible_tab: u16,
    pub selected_tab_count: u16,
    pub tab_width_ratio: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct PaneRecord {
    pub horizontal_split: u16,
    pub vertical_split: u16,
    pub top_row: u16,
    pub left_column: u16,
    pub active_pane: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ColInfoRecord {
    pub first_column: u16,
    pub last_column: u16,
    pub width: u16,
    pub format_index: u16,
    pub flags: u16,
    pub reserved: ColInfoReserved,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ColInfoReserved {
    Missing,
    Byte(u8),
    Word(u16),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct GutsRecord {
    pub row_gutter: u16,
    pub column_gutter: u16,
    pub maximum_row_outline_level: u16,
    pub maximum_column_outline_level: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CountryRecord {
    pub default_country: u16,
    pub current_country: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PaletteRecord {
    pub colors: Vec<u32>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct SclRecord {
    pub numerator: u16,
    pub denominator: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct PrintSetupRecord {
    pub paper_size: u16,
    pub scale: u16,
    pub page_start: u16,
    pub fit_width: u16,
    pub fit_height: u16,
    pub flags: u16,
    pub horizontal_resolution: u16,
    pub vertical_resolution: u16,
    /// Exact IEEE-754 bits.
    pub header_margin_bits: u64,
    /// Exact IEEE-754 bits.
    pub footer_margin_bits: u64,
    pub copies: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MulRkRecord {
    pub row: u16,
    pub first_column: u16,
    pub cells: Vec<MulRkCell>,
    pub last_column: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct MulRkCell {
    pub format_index: u16,
    /// Encoded RK number including integer/divide-by-100 flag bits.
    pub value: u32,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MulBlankRecord {
    pub row: u16,
    pub first_column: u16,
    pub format_indices: Vec<u16>,
    pub last_column: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct XfRecord {
    pub font_index: u16,
    pub number_format_index: u16,
    pub cell_flags: u16,
    pub alignment_flags: u16,
    pub indentation_flags: u16,
    pub border_style_flags: u16,
    pub border_color_flags: u16,
    pub additional_border_color_flags: u32,
    pub fill_flags: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CrnRecord {
    pub last_column: u8,
    pub first_column: u8,
    pub row: u16,
    pub values: Vec<BiffConstant>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxviFlags: u16 {
        const HIDDEN = 0x0001;
        const HIDE_DETAIL = 0x0002;
        const FORMULA = 0x0008;
        const MISSING = 0x0010;
    }
}

/// Pivot item metadata from an SXVI record.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxviRecord {
    pub item_type: i16,
    pub flags: SxviFlags,
    pub cache_index: i16,
    /// 0xFFFF represents a NULL name; every other value is the character count.
    pub declared_name_length: u16,
    pub name: Option<BiffUnicodeString>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxIvdRecord {
    pub field_indices: Vec<i16>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u16")]
pub enum SxLineItemType {
    Data = 0x0000,
    DefaultSubtotal = 0x0001,
    Sum = 0x0002,
    CountA = 0x0003,
    Count = 0x0004,
    Average = 0x0005,
    Maximum = 0x0006,
    Minimum = 0x0007,
    Product = 0x0008,
    StandardDeviation = 0x0009,
    StandardDeviationPopulation = 0x000a,
    Variance = 0x000b,
    VariancePopulation = 0x000c,
    GrandTotal = 0x000d,
    Blank = 0x000e,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxLineFlags: u16 {
        const MULTI_DATA_NAME = 0x0001;
        const SUBTOTAL = 0x0200;
        const BLOCK_TOTAL = 0x0400;
        const GRAND_TOTAL = 0x0800;
        const MULTI_DATA_ON_AXIS = 0x1000;
        const UNUSED1 = 0x2000;
        const UNUSED2 = 0x4000;
        const RESERVED2 = 0x8000;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxLiItem {
    pub shared_prefix_count: i16,
    pub item_type: SxLineItemType,
    pub reserved1: bool,
    pub displayed_item_count: i16,
    pub flags: SxLineFlags,
    pub data_item_index: u8,
    pub item_indices: Vec<i16>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxLiRecord {
    pub axis_dimension_count: u16,
    pub items: Vec<SxLiItem>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct SxPiItem {
    pub field_index: i16,
    pub item_index: i16,
    pub object_id: i16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxPiRecord {
    pub items: Vec<SxPiItem>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u16")]
pub enum SxDataAggregation {
    Sum = 0,
    Count = 1,
    Average = 2,
    Maximum = 3,
    Minimum = 4,
    Product = 5,
    CountNumbers = 6,
    StandardDeviation = 7,
    StandardDeviationPopulation = 8,
    Variance = 9,
    VariancePopulation = 10,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u16")]
pub enum SxDataDisplayCalculation {
    Normal = 0,
    Difference = 1,
    Percentage = 2,
    PercentageDifference = 3,
    RunningTotal = 4,
    PercentageOfRow = 5,
    PercentageOfColumn = 6,
    PercentageOfTotal = 7,
    Index = 8,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxDiRecord {
    pub data_field_index: i16,
    pub aggregation: SxDataAggregation,
    pub display_calculation: SxDataDisplayCalculation,
    pub calculation_field_index: i16,
    pub calculation_item_index: i16,
    pub number_format_index: u16,
    pub declared_name_length: u16,
    pub name: Option<BiffUnicodeString>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxStringRecord {
    pub declared_character_count: u16,
    pub segment: Option<BiffUnicodeString>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RrTabIdRecord {
    pub sheet_ids: Vec<u16>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxRuleAxis: u8 {
        const ROW = 0x01;
        const COLUMN = 0x02;
        const PAGE = 0x04;
        const DATA = 0x08;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxRuleFlags: u8 {
        const PART = 0x01;
        const DATA_ONLY = 0x02;
        const LABEL_ONLY = 0x04;
        const GRAND_ROW = 0x08;
        const GRAND_COLUMN = 0x10;
        const SAVED_GRAND_ROW = 0x20;
        const CACHE_BASED = 0x40;
        const SAVED_GRAND_COLUMN = 0x80;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct SxRulePartialRange {
    pub first_row: u8,
    pub last_row: u8,
    pub first_column: u8,
    pub last_column: u8,
}

/// A statically decoded PivotTable rule selector.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SxRuleRecord {
    pub dimension: u8,
    pub field_index: u8,
    pub axis: SxRuleAxis,
    /// PivotTable area kind (`sxrType`), in the range 0 through 6.
    pub area_type: u8,
    pub flags: SxRuleFlags,
    pub reserved: u16,
    pub filter_count: u16,
    pub partial_range: Option<SxRulePartialRange>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxExPageLayoutFlags: u16 {
        const ACROSS_PAGE = 0x0001;
        const UNUSED = 0x0200;
        const RESERVED1 = 0x0400;
        const RESERVED2 = 0xf800;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxExFlags: u8 {
        const ENABLE_WIZARD = 0x01;
        const ENABLE_DRILLDOWN = 0x02;
        const ENABLE_FIELD_DIALOG = 0x04;
        const PRESERVE_FORMATTING = 0x08;
        const MERGE_LABELS = 0x10;
        const DISPLAY_ERROR_STRING = 0x20;
        const DISPLAY_NULL_STRING = 0x40;
        const SUBTOTAL_HIDDEN_PAGE_ITEMS = 0x80;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxExRecord {
    pub format_count: u16,
    pub error_string_length: u16,
    pub null_string_length: u16,
    pub tag_length: u16,
    pub selection_count: u16,
    pub page_row_count: u16,
    pub page_column_count: u16,
    pub page_layout_flags: SxExPageLayoutFlags,
    pub page_wrap_count: u8,
    pub flags: SxExFlags,
    pub reserved3: u8,
    pub page_field_style_length: u16,
    pub table_style_length: u16,
    pub vacate_style_length: u16,
    pub error_string: Option<BiffUnicodeString>,
    pub null_string: Option<BiffUnicodeString>,
    pub tag: Option<BiffUnicodeString>,
    pub page_field_style: Option<BiffUnicodeString>,
    pub table_style: Option<BiffUnicodeString>,
    pub vacate_style: Option<BiffUnicodeString>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxFiltSubtotalFlags: u16 {
        const DATA = 0x0001;
        const DEFAULT = 0x0002;
        const SUM = 0x0004;
        const COUNTA = 0x0008;
        const AVERAGE = 0x0010;
        const MAX = 0x0020;
        const MIN = 0x0040;
        const PRODUCT = 0x0080;
        const COUNT = 0x0100;
        const STDEV = 0x0200;
        const STDEVP = 0x0400;
        const VARIANCE = 0x0800;
        const VARIANCEP = 0x1000;
        const BLANK = 0x4000;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SxFiltRecord {
    pub axis: SxRuleAxis,
    pub reserved1: u8,
    pub dimension: i16,
    pub field_index: i16,
    pub selected: bool,
    pub reserved2: bool,
    pub reserved3: u8,
    pub subtotal_flags: SxFiltSubtotalFlags,
    pub item_count: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxDxfRecord {
    pub format: DxfN12List,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxItmRecord {
    pub item_indices: Vec<u16>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct SxStreamIdRecord {
    pub stream_id: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct HorizontalPageBreak {
    pub row: u16,
    pub first_column: u16,
    pub last_column: u16,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct HorizontalPageBreaksRecord {
    pub break_count: u16,
    #[sdk(count = "break_count")]
    pub breaks: Vec<HorizontalPageBreak>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct VerticalPageBreak {
    pub column: u16,
    pub first_row: u16,
    pub last_row: u16,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct VerticalPageBreaksRecord {
    pub break_count: u16,
    #[sdk(count = "break_count")]
    pub breaks: Vec<VerticalPageBreak>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u16")]
pub enum DConFunction {
    Average = 0,
    CountNumbers = 1,
    Count = 2,
    Maximum = 3,
    Minimum = 4,
    Product = 5,
    StandardDeviation = 6,
    StandardDeviationPopulation = 7,
    Sum = 8,
    Variance = 9,
    VariancePopulation = 10,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u16")]
pub enum DConBoolean {
    False = 0,
    True = 1,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct DConRecord {
    pub function: DConFunction,
    pub use_left_column_labels: DConBoolean,
    pub use_top_row_labels: DConBoolean,
    pub create_source_references: DConBoolean,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct RefU {
    pub first_row: u16,
    pub last_row: u16,
    pub first_column: u8,
    pub last_column: u8,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DConSelfReferenceUnused {
    Compressed(u8),
    Unicode(u16),
    /// Producer compatibility: the required unused unit was omitted.
    Missing,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DConRefRecord {
    pub reference: RefU,
    pub declared_file_character_count: u16,
    pub file: BiffUnicodeString,
    pub self_reference_unused: Option<DConSelfReferenceUnused>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u16")]
pub enum DataSourceType {
    Odbc = 0x0001,
    Dao = 0x0002,
    Web = 0x0004,
    OleDb = 0x0005,
    Text = 0x0006,
    Ado = 0x0007,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DConnFlags: u16 {
        const SAVE_PASSWORD = 1 << 0;
        const TABLES_ONLY_HTML = 1 << 1;
        const TABLE_NAMES = 1 << 2;
        const DELETED = 1 << 3;
        const STAND_ALONE = 1 << 4;
        const ALWAYS_USE_CONNECTION_FILE = 1 << 5;
        const BACKGROUND_QUERY = 1 << 6;
        const REFRESH_ON_LOAD = 1 << 7;
        const SAVE_DATA = 1 << 8;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DConnSecondaryFlags: u8 {
        const MAINTAIN_CONNECTION = 1 << 0;
        const NEW_QUERY = 1 << 1;
        const IMPORT_XML_SOURCE = 1 << 2;
        const SHAREPOINT_LIST_SOURCE = 1 << 3;
        const SHAREPOINT_REINITIALIZE_CACHE = 1 << 4;
        const SOURCE_IS_XML = 1 << 7;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DConnWebFlags: u16 {
        const PARSE_PRE_FORMATTED = 1 << 0;
        const CONSECUTIVE_DELIMITERS = 1 << 1;
        const SAME_SETTINGS = 1 << 2;
        const XL97_FORMAT = 1 << 3;
        const NO_DATE_RECOGNITION = 1 << 4;
        const REFRESHED_IN_XL9 = 1 << 5;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DConnOleDbFlags: u16 {
        const LOCAL_CONNECTION = 1 << 3;
        const NO_REFRESH_CUBE = 1 << 4;
        const USE_OFFICE_LCID = 1 << 5;
        const SERVER_FORMAT_NUMBER = 1 << 6;
        const SERVER_FORMAT_BACKGROUND = 1 << 7;
        const SERVER_FORMAT_FOREGROUND = 1 << 8;
        const SERVER_FORMAT_FLAGS = 1 << 9;
        const SUPPORTS_LANGUAGE_CELL_PROPERTY = 1 << 10;
        const SERVER_SUPPORTS_CLIENT_CUBE = 1 << 11;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DConnAdoFlags: u8 {
        const REFRESHABLE = 1 << 0;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DConnQueryFlags {
    Unused(u16),
    Web(DConnWebFlags),
    OleDb(DConnOleDbFlags),
    Ado { reserved1: u8, flags: DConnAdoFlags },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DConnUnicodeStringSegmented {
    pub total_character_count: u32,
    pub segments: Vec<XlUnicodeString>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DConnStringSequence {
    pub strings: Vec<DConnUnicodeStringSegmented>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct TextQueryField {
    pub field_type: u32,
    pub field_start: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TextQueryOptions {
    pub file: bool,
    pub delimited: bool,
    pub code_page_kind: u8,
    pub prompt_for_file: bool,
    pub new_code_page: u16,
    pub use_new_code_page: bool,
    pub unused: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TextQueryDelimiterOptions {
    pub tab: bool,
    pub space: bool,
    pub comma: bool,
    pub semicolon: bool,
    pub custom: bool,
    pub consecutive: bool,
    pub text_delimiter: u8,
    pub custom_character: u16,
    pub unused: u8,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextQuery {
    pub record_type: u16,
    pub reserved: u16,
    pub options: TextQueryOptions,
    pub starting_row: i32,
    pub delimiter_options: TextQueryDelimiterOptions,
    pub fields: Vec<TextQueryField>,
    pub decimal_separator: u8,
    pub thousands_separator: u8,
    pub source_file: XlUnicodeString,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DConnOleDbConnection {
    pub drillthrough_row_limit: u32,
    pub valid_connection_types: Vec<u16>,
    pub invalid_connection_types: Vec<u16>,
    pub unused: u16,
    pub connection_strings: Vec<DConnUnicodeStringSegmented>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DConnWebConnection {
    pub url: DConnStringSequence,
    pub post_method: DConnStringSequence,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DConnConnection {
    Odbc(DConnUnicodeStringSegmented),
    Web(DConnWebConnection),
    OleDb(DConnOleDbConnection),
    Text(TextQuery),
    None,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DConnParameterBindingValue {
    Numeric(u64),
    String {
        reserved: u64,
        value: DConnUnicodeStringSegmented,
    },
    Boolean {
        value: u8,
        reserved1: [u8; 3],
        reserved2: u32,
    },
    Integer {
        value: u32,
        reserved: u32,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DConnParameterBinding {
    Prompt(DConnUnicodeStringSegmented),
    Value {
        value_type: u16,
        value: DConnParameterBindingValue,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DConnParameter {
    pub name: DConnUnicodeStringSegmented,
    pub parameter_type: u8,
    pub reserved: u16,
    pub sql_type: i16,
    pub default_name: bool,
    pub unused: u16,
    pub binding: DConnParameterBinding,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DConnIdentifier {
    None,
    QueryTable(DConnUnicodeStringSegmented),
    PivotCache(u16),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DConnRecord {
    pub header: FrtHeaderOld,
    pub data_source_type: DataSourceType,
    pub flags: DConnFlags,
    pub parameters_count: u16,
    pub reserved1: u16,
    pub secondary_flags: DConnSecondaryFlags,
    pub reserved2: u8,
    pub query_flags: DConnQueryFlags,
    pub edited_version: u8,
    pub refreshed_version: u8,
    pub minimum_refreshable_version: u8,
    pub refresh_interval_minutes: u16,
    pub html_format: u16,
    pub reconnection_method: u32,
    pub credential_method: u8,
    pub reserved3: u8,
    pub source_data_file: DConnUnicodeStringSegmented,
    pub source_connection_file: DConnUnicodeStringSegmented,
    pub connection_name: DConnUnicodeStringSegmented,
    pub connection_description: DConnUnicodeStringSegmented,
    pub sso_application_id: DConnUnicodeStringSegmented,
    pub table_names: Option<DConnUnicodeStringSegmented>,
    pub parameters: Vec<DConnParameter>,
    pub connection: DConnConnection,
    pub sql: DConnStringSequence,
    pub saved_sql: DConnStringSequence,
    pub edit_web_page: DConnStringSequence,
    pub identifier: DConnIdentifier,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct QsiSxTagFlags: u16 {
        const ENABLE_REFRESH = 1 << 0;
        const INVALID = 1 << 1;
        const OLAP_PIVOT_TABLE = 1 << 2;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct QueryTableFutureFlags: u32 {
        const PRESERVE_FORMATTING = 1 << 0;
        const AUTO_FIT = 1 << 1;
        const EXTERNAL_DATA_LIST = 1 << 4;
        const CREATE_QUERY_TABLE_LIST = 1 << 6;
        const DUMMY_LIST = 1 << 7;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct PivotTableFutureFlags: u32 {
        const NO_STENCIL = 1 << 0;
        const HIDE_TOTAL_ANNOTATION = 1 << 1;
        const INCLUDE_EMPTY_ROWS = 1 << 3;
        const INCLUDE_EMPTY_COLUMNS = 1 << 4;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum QsiSxTagFutureFlags {
    QueryTable(QueryTableFutureFlags),
    PivotTable(PivotTableFutureFlags),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QsiSxTagRecord {
    pub header: FrtHeaderOld,
    pub pivot_table: DConBoolean,
    pub flags: QsiSxTagFlags,
    pub future_flags: QsiSxTagFutureFlags,
    pub last_updated_version: u8,
    pub minimum_updatable_version: u8,
    pub name_character_offset: u8,
    pub reserved: u8,
    pub name: XlUnicodeString,
    pub unused: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxViewEx9HeaderFlags: u16 {
        const FUTURE_RECORD_ALERT = 1 << 1;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxViewEx9Flags: u32 {
        const PRINT_TITLES = 1 << 1;
        const OUTLINE_MODE = 1 << 2;
        const REPEAT_ITEMS_ON_PRINTED_PAGES = 1 << 5;
    }
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct SxViewEx9Record {
    pub record_type: u16,
    #[sdk(bitflags = "u16")]
    pub header_flags: SxViewEx9HeaderFlags,
    pub reserved: u32,
    #[sdk(bitflags = "u32")]
    pub flags: SxViewEx9Flags,
    pub autoformat_index: u16,
    pub grand_total_caption: XlUnicodeString,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DbQueryExtConnectionFlags: u16 {
        const MAINTAIN_CONNECTION = 1 << 0;
        const NEW_QUERY = 1 << 1;
        const IMPORT_XML_SOURCE = 1 << 2;
        const SHAREPOINT_LIST_SOURCE = 1 << 3;
        const SHAREPOINT_REINITIALIZE_CACHE = 1 << 4;
        const SOURCE_IS_XML = 1 << 7;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DbQueryExtFlags: u16 {
        const TEXT_WIZARD = 1 << 0;
        const TABLE_NAMES = 1 << 1;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct DbQueryParameterFlags {
    pub parameter_type: u8,
    pub auto_refresh: bool,
    pub need_refresh: bool,
    pub reserved: u16,
}

impl DbQueryParameterFlags {
    fn from_bits(bits: u16) -> Self {
        Self {
            parameter_type: (bits & 0x0007) as u8,
            auto_refresh: bits & (1 << 3) != 0,
            need_refresh: bits & (1 << 4) != 0,
            reserved: bits >> 5,
        }
    }

    fn bits(self) -> u16 {
        u16::from(self.parameter_type & 0x07)
            | (u16::from(self.auto_refresh) << 3)
            | (u16::from(self.need_refresh) << 4)
            | (self.reserved << 5)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DbQueryExtRecord {
    pub header: FrtHeaderOld,
    pub data_source_type: DataSourceType,
    pub connection_flags: DbQueryExtConnectionFlags,
    pub query_flags: DConnQueryFlags,
    pub flags: DbQueryExtFlags,
    pub edited_version: u8,
    pub refreshed_version: u8,
    pub minimum_refreshable_version: u8,
    pub reserved4: u8,
    pub reserved5: u16,
    pub ole_db_connection_count: u16,
    pub future_byte_count: u16,
    pub refresh_interval_minutes: u16,
    pub html_format: u16,
    pub parameter_count: u16,
    pub parameters: Vec<DbQueryParameterFlags>,
    /// Explicit future-version area sized by `future_byte_count`.
    pub future_bytes: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct FrtRefHeaderNoGrbit {
    pub record_type: u16,
    pub range: CellRange,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HyperlinkTooltipRecord {
    pub header: FrtRefHeaderNoGrbit,
    /// UTF-16 tooltip characters including the terminating NUL.
    pub tooltip: Vec<u16>,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct ContinueFrt12Record {
    pub header: FrtRefHeaderU,
    /// Continuation bytes whose schema belongs to the preceding future record.
    #[sdk(remaining)]
    pub continuation: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct SxAddlHeader {
    pub future_header: FrtHeaderOld,
    pub class: u8,
    pub data_type: u8,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxAddlString {
    pub total_character_count: u32,
    pub reserved: u16,
    pub segment: XlUnicodeString,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxAddlViewVer10Flags: u32 {
        const DISPLAY_IMMEDIATE_ITEMS = 1 << 0;
        const ENABLE_DATA_EDITING = 1 << 1;
        const DISABLE_FIELD_LIST = 1 << 2;
        const REENTER_ON_LOAD_ONCE = 1 << 3;
        const HIDE_CALCULATED_MEMBERS = 1 << 4;
        const NON_VISUAL_TOTALS = 1 << 5;
        const PAGE_MULTIPLE_ITEM_LABEL = 1 << 6;
        const USE_TENSOR_FILL_COLOR = 1 << 7;
        const HIDE_DROPDOWN_DATA = 1 << 8;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxAddlViewVer12Flags: u32 {
        const DEFAULT_COMPACT = 1 << 0;
        const DEFAULT_OUTLINE = 1 << 1;
        const OUTLINE_DATA = 1 << 2;
        const COMPACT_DATA = 1 << 3;
        const NEW_DROP_ZONES = 1 << 4;
        const PUBLISHED = 1 << 5;
        const TURN_OFF_IMMERSIVE = 1 << 6;
        const SINGLE_FILTER_PER_FIELD = 1 << 7;
        const NON_DEFAULT_SORT_IN_FIELD_LIST = 1 << 8;
        const DONT_USE_CUSTOM_LISTS = 1 << 10;
        const HIDE_DRILL_INDICATORS = 1 << 20;
        const PRINT_DRILL_INDICATORS = 1 << 21;
        const MEMBER_PROPERTIES_IN_TOOLTIPS = 1 << 22;
        const NO_PIVOT_TOOLTIPS = 1 << 23;
        const NO_HEADERS = 1 << 31;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxAddlTableStyleFlags: u16 {
        const UNUSED = 1 << 0;
        const LAST_COLUMN = 1 << 1;
        const ROW_STRIPES = 1 << 2;
        const COLUMN_STRIPES = 1 << 3;
        const ROW_HEADERS = 1 << 4;
        const COLUMN_HEADERS = 1 << 5;
        const DEFAULT_STYLE = 1 << 6;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxAddlCacheInfo12Flags: u32 {
        const SHEET_DATA = 1 << 0;
        const SERVER_SUPPORTS_ATTRIBUTE_DRILLDOWN = 1 << 1;
        const SERVER_SUPPORTS_SUBQUERY = 1 << 2;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxAddlCacheRefreshFlags: u32 {
        const ENABLE_REFRESH = 1 << 0;
        const INVALID = 1 << 1;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxAddlField12Flags: u32 {
        const HIDDEN_LEVEL = 1 << 1;
        const USE_MEMBER_PROPERTY_CAPTION = 1 << 2;
        const COMPACT = 1 << 3;
        const SORT_ITEMS_ON_NEXT_SORT = 1 << 4;
        const FILTER_INCLUSIVE = 1 << 5;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SxAddlData {
    End {
        reserved: [u8; 6],
    },
    ViewId(SxAddlString),
    VersionUpdateInvalidates {
        version: u8,
        reserved1: u8,
        reserved2: u16,
        reserved3: u16,
    },
    ViewVersion10 {
        created_version: u8,
        flags: SxAddlViewVer10Flags,
        reserved: u16,
    },
    ViewVersion12 {
        flags: SxAddlViewVer12Flags,
        reserved: u16,
    },
    ViewTableStyle {
        reserved: [u8; 6],
        flags: SxAddlTableStyleFlags,
        style_name: LpWideString,
    },
    CacheId {
        cache_stream_id: u32,
        reserved: u16,
    },
    CacheVersion10 {
        reserved1: [u8; 6],
        ghost_item_limit: i32,
        last_refresh_version: u8,
        minimum_refreshable_version: u8,
        refresh_date_bits: u64,
        reserved2: u16,
    },
    CacheVersionMacro {
        version: u8,
        reserved1: u8,
        reserved2: u16,
        reserved3: u16,
    },
    CacheInvalidRefresh {
        flags: SxAddlCacheRefreshFlags,
        reserved: u16,
    },
    CacheInfo12 {
        flags: SxAddlCacheInfo12Flags,
        reserved: u16,
    },
    Field12Id(SxAddlString),
    Field12Version12 {
        flags: SxAddlField12Flags,
        reserved: u16,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxAddlRecord {
    pub header: SxAddlHeader,
    pub data: SxAddlData,
}

/// Application-specific, rebuildable cache bytes from an EntExU2 record.
#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct EntExU2Record {
    #[sdk(remaining)]
    pub cache: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum BkHimImage {
    Bitmap(DeviceIndependentBitmap),
    /// Native image bytes whose application-specific format is opaque by specification.
    Native(Vec<u8>),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BkHimRecord {
    pub reserved: u16,
    pub image: BkHimImage,
    /// Payload lengths of the BkHim record followed by each Continue record.
    pub physical_segment_lengths: Vec<u16>,
}

/// Legacy BIFF ImgData uses the same image envelope as BkHim under record ID 0x007F.
pub type ImDataRecord = BkHimRecord;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RtdTopicString {
    /// Count of encoded units, including one length unit per substring.
    pub declared_unit_count: u32,
    /// Exact option flags. Bit 0 selects UTF-16; remaining bits are reserved.
    pub flags: u8,
    pub substrings: Vec<XlStringCharacters>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum RtdOperation {
    Number {
        bits: u64,
    },
    ShortString(BiffUnicodeString),
    Boolean(u32),
    Error(i32),
    ErrorWithCorruptDiscriminator {
        discriminator: u32,
        value: i32,
    },
    Integer(i32),
    LongString(BiffUnicodeString),
    /// A structurally bounded operation with a discriminator outside MS-XLS.
    Malformed {
        discriminator: u32,
        payload: Vec<u8>,
    },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct RtdCellReference {
    pub row: u16,
    pub column: u16,
    pub sheet_index: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RealTimeDataRecord {
    pub header: FrtHeader,
    pub shared_prefix_character_count: u32,
    pub topic: RtdTopicString,
    pub operation: RtdOperation,
    pub cells: Vec<RtdCellReference>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SortOptions {
    pub sort_columns: bool,
    pub descending: [bool; 3],
    pub case_sensitive: bool,
    /// Signed 5-bit index into the environment-specific custom sort lists.
    pub custom_list_index: i8,
    pub alternate_method: bool,
    /// The five ignored high bits retained for byte-exact producer compatibility.
    pub reserved: u8,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SortRecord {
    pub options: SortOptions,
    pub keys: [Option<BiffUnicodeString>; 3],
    pub reserved: u8,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SortFieldParent {
    Sheet,
    Table,
    AutoFilter,
    QueryTable,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SortDataOptions {
    pub sort_columns: bool,
    pub case_sensitive: bool,
    pub alternate_method: bool,
    pub parent: SortFieldParent,
    /// Undefined high ten bits retained for byte-exact producer compatibility.
    pub unused: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct Rfx {
    pub first_row: u32,
    pub last_row: u32,
    pub first_column: u32,
    pub last_column: u32,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SortDataRecord {
    pub header: FrtHeader,
    pub options: SortDataOptions,
    pub range: Rfx,
    pub condition_count: u32,
    pub parent_id: u32,
    pub conditions: Vec<SortConditionContinuation>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SortConditionContinuation {
    pub header: FrtRefHeaderU,
    pub condition: SortCondition,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SortConditionData {
    Value { value: u32, reserved: u32 },
    CellColor { dxf_index: u32, reserved: u32 },
    FontColor { dxf_index: u32, reserved: u32 },
    Icon { icon_set: u32, icon_index: i32 },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SortCondition {
    pub descending: bool,
    pub reserved: u16,
    pub range: Rfx,
    pub data: SortConditionData,
    pub custom_list: Option<BiffUnicodeString>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct AutoFilterOptions {
    pub join_or: bool,
    pub simple: [bool; 2],
    pub top_n: bool,
    pub top: bool,
    pub percent: bool,
    pub top_count: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AutoFilterOperandValue {
    Unused {
        bytes: [u8; 8],
    },
    Rk {
        value: u32,
        unused: u32,
    },
    Number {
        bits: u64,
    },
    String {
        unused1: u32,
        declared_character_count: u8,
        compare_without_wildcards: u8,
        reserved: u8,
        unused2: u8,
    },
    BooleanOrError {
        value: u16,
        unused1: u16,
        unused2: u32,
    },
    Blanks {
        reserved: u64,
    },
    NonBlanks {
        reserved: u64,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AutoFilterOperand {
    pub comparison: u8,
    pub value: AutoFilterOperandValue,
    pub string: Option<BiffUnicodeString>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AutoFilterRecord {
    pub entry_index: u16,
    pub options: AutoFilterOptions,
    pub operands: [AutoFilterOperand; 2],
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SxFormatRecord {
    pub formatting_applied: bool,
    /// Ignored high twelve bits retained for producer compatibility.
    pub reserved: u16,
    pub differential_format_byte_count: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct WOptFlags: u16 {
        const RELY_ON_CSS = 0x0001;
        const ORGANIZE_IN_FOLDER = 0x0002;
        const USE_LONG_FILE_NAMES = 0x0004;
        const DOWNLOAD_COMPONENTS = 0x0008;
        const RELY_ON_VML = 0x0010;
        const ALLOW_PNG = 0x0020;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u8")]
pub enum WebScreenSize {
    Pixels544x376 = 0,
    Pixels640x480 = 1,
    Pixels720x512 = 2,
    Pixels800x600 = 3,
    Pixels1024x768 = 4,
    Pixels1152x882 = 5,
    Pixels1152x900 = 6,
    Pixels1280x1024 = 7,
    Pixels1600x1200 = 8,
    Pixels1800x1440 = 9,
    Pixels1920x1200 = 10,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct WOptRecord {
    pub header: FrtHeaderOld,
    pub flags: WOptFlags,
    pub screen_size: WebScreenSize,
    pub reserved: u8,
    pub pixels_per_inch: u32,
    pub code_page: u32,
    pub component_location: LpWideString,
    pub future: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct TableRange {
    pub first_row: u16,
    pub last_row: u16,
    pub first_column: u8,
    pub last_column: u8,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TableOptions {
    pub always_calculate: bool,
    pub reserved1: bool,
    pub row_input: bool,
    pub two_variable: bool,
    pub first_input_deleted: bool,
    pub second_input_deleted: bool,
    pub reserved2: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct TableInputReference {
    pub row: u16,
    pub column: i16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TableRecord {
    pub range: TableRange,
    pub options: TableOptions,
    pub row_input: TableInputReference,
    pub column_input: TableInputReference,
    /// Two-byte padding emitted by one legacy producer in the corpus.
    pub compatibility_padding: Option<u16>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ExternCountRecord {
    pub external_sheet_count: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct QsiFlags: u16 {
        const TITLES = 0x0001;
        const ROW_NUMBERS = 0x0002;
        const DISABLE_REFRESH = 0x0004;
        const ASYNC = 0x0008;
        const NEW_ASYNC = 0x0010;
        const AUTO_REFRESH = 0x0020;
        const SHRINK = 0x0040;
        const FILL = 0x0080;
        const AUTO_FORMAT = 0x0100;
        const SAVE_DATA = 0x0200;
        const DISABLE_EDIT = 0x0400;
        const OVERWRITE = 0x2000;
    }
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct QsiFormattingFlags: u16 {
        const NUMBER = 0x0001;
        const FONT = 0x0002;
        const ALIGNMENT = 0x0004;
        const BORDER = 0x0008;
        const PATTERN = 0x0010;
        const PROTECTION = 0x0020;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QsiRecord {
    pub flags: QsiFlags,
    pub auto_format_index: u16,
    pub formatting_flags: QsiFormattingFlags,
    pub reserved: u32,
    pub name: XlUnicodeString,
    pub unused: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ParamQryFixed {
    pub sql_type: u16,
    pub parameter_type: u8,
    pub unused1: bool,
    pub non_default_name: bool,
    pub unused2: u16,
    pub value_type: u16,
    pub boolean_value: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ParamQryData {
    Prompt { text: SxStringRecord, unused: u8 },
    Number { bits: u64 },
    String { text: SxStringRecord, unused: u8 },
    Boolean,
    Integer(i32),
    Reference(FormulaTokenStream),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParamQryRecord {
    pub fixed: ParamQryFixed,
    pub data: ParamQryData,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u8")]
pub enum PaneType {
    BottomRight = 0,
    TopRight = 1,
    BottomLeft = 2,
    TopLeft = 3,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SxSelectOptions {
    pub click_count: u8,
    pub label_only: bool,
    pub data_only: bool,
    pub toggle_data_header: bool,
    pub selection_click: bool,
    pub extendable: bool,
    pub unused: u8,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SxSelectRecord {
    pub reserved1: u16,
    pub pane: PaneType,
    pub reserved2: u8,
    pub axis: SxAxis,
    pub active_dimension: u16,
    pub line_start: u16,
    pub active_line: u16,
    pub minimum_line: u16,
    pub maximum_line: u16,
    pub clicked_row: u16,
    pub clicked_column: u16,
    pub previous_clicked_row: u16,
    pub previous_clicked_column: u16,
    pub options: SxSelectOptions,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum FileSharingData {
    NoPassword { reserved: u16 },
    Password { user_name: XlUnicodeString },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FileSharingRecord {
    pub read_only_recommended: DConBoolean,
    pub password_verifier: u16,
    pub data: FileSharingData,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct OleObjectSizeRecord {
    pub unused: u16,
    pub visible_range: RefU,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u32")]
pub enum MsoDrawingSelectionMode {
    Normal = 0,
    Rotate = 1,
    Reshape = 2,
    Crop = 7,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MsoDrawingSelectionRecord {
    pub header: OfficeArtRecordHeader,
    pub shape_count_unused: u32,
    pub mode: MsoDrawingSelectionMode,
    pub focus_shape_id: u32,
    pub selected_shape_ids: Vec<u32>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ScenManRecord {
    pub scenario_count: i16,
    pub current_scenario_index: i16,
    pub shown_scenario_index: i16,
    pub result_reference_count: i16,
    pub result_references: Vec<CellRange>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxViewFlags: u16 {
        const ROW_GRAND_TOTALS = 0x0001;
        const COLUMN_GRAND_TOTALS = 0x0002;
        const UNUSED1 = 0x0004;
        const AUTO_FORMAT = 0x0008;
        const AUTO_FORMAT_NUMBER = 0x0010;
        const AUTO_FORMAT_FONT = 0x0020;
        const AUTO_FORMAT_ALIGNMENT = 0x0040;
        const AUTO_FORMAT_BORDER = 0x0080;
        const AUTO_FORMAT_PATTERN = 0x0100;
        const AUTO_FORMAT_DIMENSIONS = 0x0200;
        const UNUSED2 = 0xfc00;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxViewRecord {
    pub report_body: CellRange,
    pub first_header_row: u16,
    pub first_data_row: u16,
    pub first_data_column: u16,
    pub cache_index: i16,
    pub reserved: u16,
    pub data_axis: SxAxis,
    pub data_position: i16,
    pub field_count: i16,
    pub row_field_count: u16,
    pub column_field_count: u16,
    pub page_field_count: u16,
    pub data_field_count: i16,
    pub row_line_count: u16,
    pub column_line_count: u16,
    pub flags: SxViewFlags,
    pub auto_format_index: u16,
    pub declared_table_name_length: u16,
    pub declared_data_name_length: u16,
    pub table_name: BiffUnicodeString,
    pub data_name: BiffUnicodeString,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u16")]
pub enum SxSourceType {
    Worksheet = 0x0001,
    External = 0x0002,
    Consolidation = 0x0004,
    Scenario = 0x0010,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct SxVsRecord {
    pub source_type: SxSourceType,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct RecalcIdRecord {
    pub record_type: u16,
    pub reserved: u16,
    pub engine_build: u32,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct CodeNameRecord {
    pub name: XlUnicodeString,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ArrayFlags: u16 {
        const ALWAYS_CALCULATE = 0x0001;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ArrayRecord {
    pub first_row: u16,
    pub last_row: u16,
    pub first_column: u8,
    pub last_column: u8,
    pub flags: ArrayFlags,
    pub unused: u32,
    pub tokens: FormulaTokens,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct UserSViewFlags: u32 {
        const SHOW_PAGE_BREAKS = 1 << 0;
        const SHOW_FORMULAS = 1 << 1;
        const SHOW_GRIDLINES = 1 << 2;
        const SHOW_ROW_COLUMN_HEADINGS = 1 << 3;
        const SHOW_OUTLINE_SYMBOLS = 1 << 4;
        const SHOW_ZERO_VALUES = 1 << 5;
        const CENTER_HORIZONTALLY = 1 << 6;
        const CENTER_VERTICALLY = 1 << 7;
        const PRINT_ROW_COLUMN_HEADINGS = 1 << 8;
        const PRINT_GRIDLINES = 1 << 9;
        const FIT_TO_PAGE = 1 << 10;
        const HAS_PRINT_AREA = 1 << 11;
        const ONE_PRINT_AREA = 1 << 12;
        const FILTER_MODE = 1 << 13;
        const SHOW_AUTO_FILTER = 1 << 14;
        const FROZEN = 1 << 15;
        const FROZEN_NO_SPLIT = 1 << 16;
        const SPLIT_VERTICAL = 1 << 17;
        const SPLIT_HORIZONTAL = 1 << 18;
        const HIDDEN_ROWS_MASK = 0x0018_0000;
        const HIDDEN_COLUMNS = 1 << 21;
        const FILTER_UNIQUE = 1 << 25;
        const PAGE_BREAK_PREVIEW = 1 << 26;
        const PAGE_LAYOUT_VIEW = 1 << 27;
        const RULER = 1 << 29;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct UserSViewBeginRecord {
    pub guid: [u8; 16],
    pub sheet_id: u16,
    pub reserved1: u16,
    pub zoom_scale: u32,
    pub gridline_color_index: u16,
    pub reserved2: u16,
    pub selected_pane: u8,
    pub reserved3: u16,
    pub reserved4: u8,
    #[sdk(bitflags = "u32")]
    pub flags: UserSViewFlags,
    pub top_left_range: CellRange,
    pub split_x_bits: u64,
    pub split_y_bits: u64,
    pub right_pane_column: u16,
    pub bottom_pane_row: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct UserSViewEndRecord {
    pub reserved: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct UserBViewFlags: u16 {
        const SHOW_FORMULA_BAR = 1 << 0;
        const SHOW_STATUS_BAR = 1 << 1;
        const NOTE_DISPLAY_MASK = 0x000c;
        const SHOW_HORIZONTAL_SCROLLBAR = 1 << 4;
        const SHOW_VERTICAL_SCROLLBAR = 1 << 5;
        const SHOW_SHEET_TABS = 1 << 6;
        const MAXIMIZED = 1 << 7;
        const HIDE_OBJECTS_MASK = 0x0300;
        const INCLUDE_PRINT_SETTINGS = 1 << 10;
        const INCLUDE_ROW_COLUMN_SETTINGS = 1 << 11;
        const INVALID_SHEET_ID = 1 << 12;
        const TIMED_UPDATE = 1 << 13;
        const ALL_MEMBER_CHANGES = 1 << 14;
        const ONLY_SYNCHRONIZE = 1 << 15;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct UserBViewWindowFlags: u16 {
        const PERSONAL_VIEW = 1 << 0;
        const MINIMIZED = 1 << 1;
    }
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct UserBViewRecord {
    pub unused1: u32,
    pub active_sheet_id: u16,
    pub reserved1: u16,
    pub guid: [u8; 16],
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
    pub tab_ratio: u16,
    #[sdk(bitflags = "u16")]
    pub flags: UserBViewFlags,
    pub unused2: u16,
    #[sdk(bitflags = "u16")]
    pub window_flags: UserBViewWindowFlags,
    pub merge_interval_minutes: u16,
    pub name: XlUnicodeString,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SheetExtColorIndex {
    pub color_index: u8,
    pub reserved: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SheetExtOptionalFlags {
    pub color_index: u8,
    pub calculate_conditional_formats: bool,
    pub not_published: bool,
    pub reserved: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CfColor {
    pub color_type: u32,
    pub color_value: u32,
    pub tint_bits: u64,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SheetExtOptional {
    pub flags: SheetExtOptionalFlags,
    pub color: CfColor,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SheetExtRecord {
    pub header: FrtHeader,
    pub declared_size: u32,
    pub tab_color: SheetExtColorIndex,
    pub optional: Option<SheetExtOptional>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct FrtRefHeaderU {
    pub record_type: u16,
    #[sdk(bitflags = "u16")]
    pub flags: FrtFlags,
    pub range: CellRange,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CellWatchRecord {
    pub header: FrtRefHeaderU,
    pub reserved: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct FeatureHeader11Record {
    pub header: FrtHeader,
    pub shared_feature_type: u16,
    pub reserved1: u8,
    pub reserved2: u32,
    pub reserved3: u32,
    pub next_list_id: u32,
    pub reserved4: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct TableFeatureFlags: u32 {
        const UNUSED2 = 1 << 0;
        const AUTO_FILTER = 1 << 1;
        const PERSIST_AUTO_FILTER = 1 << 2;
        const SHOW_INSERT_ROW = 1 << 3;
        const INSERT_ROW_INSERTS_CELLS = 1 << 4;
        const LOAD_DELETED_IDS = 1 << 5;
        const SHOWN_TOTAL_ROW = 1 << 6;
        const NEEDS_COMMIT = 1 << 8;
        const SINGLE_CELL = 1 << 9;
        const APPLY_AUTO_FILTER = 1 << 11;
        const FORCE_INSERT_VISIBLE = 1 << 12;
        const COMPRESSED_XML = 1 << 13;
        const LOAD_CSP_NAME = 1 << 14;
        const LOAD_CHANGED_IDS = 1 << 15;
        const VERSION_MASK = 0x000f_0000;
        const LOAD_ENTRY_ID = 1 << 20;
        const LOAD_INVALID_CELLS = 1 << 21;
        const GOOD_BUILD = 1 << 22;
        const UNUSED3 = 1 << 23;
        const PUBLISHED = 1 << 24;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct Feature11FieldFlags: u32 {
        const AUTO_FILTER = 1 << 0;
        const AUTO_FILTER_HIDDEN = 1 << 1;
        const LOAD_XML_MAP = 1 << 2;
        const LOAD_FORMULA = 1 << 3;
        const LOAD_TOTAL_FORMULA = 1 << 7;
        const LOAD_TOTAL_ARRAY = 1 << 8;
        const SAVE_STYLE_NAME = 1 << 9;
        const LOAD_TOTAL_STRING = 1 << 10;
        const AUTO_CREATE_CALCULATED_COLUMN = 1 << 11;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DxfN12List {
    pub format: Box<DxfN>,
    pub extension: Option<XfExtNoFrt>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Feature11AutoFilter {
    pub declared_size: u32,
    pub unused: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Feature11FieldDataItem {
    pub field_id: u32,
    pub web_data_type: u32,
    pub xml_data_type: u32,
    pub total_aggregation: u32,
    pub aggregate_format_size: u32,
    pub aggregate_style_index: u32,
    pub flags: Feature11FieldFlags,
    pub insert_row_format_size: u32,
    pub insert_row_style_index: u32,
    pub field_name: XlUnicodeString,
    pub caption: Option<XlUnicodeString>,
    pub aggregate_format: Option<DxfN12List>,
    pub insert_row_format: Option<DxfN12List>,
    pub auto_filter: Option<Feature11AutoFilter>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableFeatureType {
    pub source_type: u32,
    pub list_id: u32,
    pub header_row_count: u32,
    pub total_row_count: u32,
    pub next_field_id: u32,
    pub fixed_data_size: u32,
    pub writer_build: u16,
    pub unused1: u16,
    pub flags: TableFeatureFlags,
    pub cache_stream_offset: u32,
    pub cache_stream_size: u32,
    pub cache_character_count: u32,
    pub edit_mode: u32,
    pub hash_parameters: [u8; 16],
    pub name: XlUnicodeString,
    pub field_count: u16,
    pub csp_name: Option<XlUnicodeString>,
    pub entry_id: Option<XlUnicodeString>,
    pub fields: Vec<Feature11FieldDataItem>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Feature11Record {
    pub header: FrtRefHeaderU,
    pub shared_feature_type: u16,
    pub reserved1: u8,
    pub reserved2: u32,
    pub reference_count: u16,
    pub declared_feature_size: u32,
    pub reserved3: u16,
    pub references: Vec<CellRange>,
    pub feature: TableFeatureType,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct List12BlockLevel {
    pub header_format_size: i32,
    pub header_style_index: i32,
    pub data_format_size: i32,
    pub data_style_index: i32,
    pub aggregate_format_size: i32,
    pub aggregate_style_index: i32,
    pub border_format_size: i32,
    pub header_border_format_size: i32,
    pub aggregate_border_format_size: i32,
    pub header_format: Option<DxfN12List>,
    pub data_format: Option<DxfN12List>,
    pub aggregate_format: Option<DxfN12List>,
    pub border_format: Option<DxfN12List>,
    pub header_border_format: Option<DxfN12List>,
    pub aggregate_border_format: Option<DxfN12List>,
    pub header_style_name: Option<XlUnicodeString>,
    pub data_style_name: Option<XlUnicodeString>,
    pub aggregate_style_name: Option<XlUnicodeString>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct List12TableStyleFlags: u16 {
        const FIRST_COLUMN = 0x0001;
        const LAST_COLUMN = 0x0002;
        const ROW_STRIPES = 0x0004;
        const COLUMN_STRIPES = 0x0008;
        const DEFAULT_STYLE = 0x0040;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum List12Data {
    BlockLevel(Box<List12BlockLevel>),
    TableStyle {
        flags: List12TableStyleFlags,
        style_name: XlUnicodeString,
    },
    DisplayName {
        list_name: XlUnicodeString,
        comment: XlUnicodeString,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct List12Record {
    pub header: FrtHeader,
    pub data_type: u16,
    pub list_id: u32,
    pub data: List12Data,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct DropDownObjIdsRecord {
    pub header: FrtHeader,
    pub object_id_count: u16,
    #[sdk(count = "object_id_count")]
    pub object_ids: Vec<u16>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DataValidationHeaderFlags: u16 {
        const INPUT_WINDOW_CLOSED = 0x0001;
        const UNUSED = 0x0004;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct DataValidationHeaderRecord {
    #[sdk(bitflags = "u16")]
    pub flags: DataValidationHeaderFlags,
    pub input_window_x: u32,
    pub input_window_y: u32,
    pub dropdown_object_id: i32,
    pub validation_count: u32,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlUnicodeStringMin2 {
    pub declared_character_count: u16,
    pub text: Option<BiffUnicodeString>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartDataLabelExtContentsFlags: u16 {
        const SERIES_NAME = 0x0001;
        const CATEGORY_NAME = 0x0002;
        const VALUE = 0x0004;
        const PERCENT = 0x0008;
        const BUBBLE_SIZE = 0x0010;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ChartDataLabelExtContentsRecord {
    pub header: FrtHeader,
    pub flags: ChartDataLabelExtContentsFlags,
    pub separator: XlUnicodeStringMin2,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct Chart3DBarShapeRecord {
    pub riser: u8,
    pub taper: u8,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxvdExFlags: u32 {
        const SHOW_ALL_ITEMS = 1 << 0;
        const DRAG_TO_ROW = 1 << 1;
        const DRAG_TO_COLUMN = 1 << 2;
        const DRAG_TO_PAGE = 1 << 3;
        const DRAG_TO_HIDE = 1 << 4;
        const NOT_DRAG_TO_DATA = 1 << 5;
        const SERVER_BASED = 1 << 7;
        const AUTO_SORT = 1 << 9;
        const ASCENDING_SORT = 1 << 10;
        const AUTO_SHOW = 1 << 11;
        const TOP_AUTO_SHOW = 1 << 12;
        const CALCULATED_FIELD = 1 << 13;
        const PAGE_BREAKS_BETWEEN_ITEMS = 1 << 14;
        const HIDE_NEW_ITEMS = 1 << 15;
        const OUTLINE = 1 << 21;
        const INSERT_BLANK_ROW = 1 << 22;
        const SUBTOTAL_AT_TOP = 1 << 23;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxvdExOptional {
    pub declared_name_length: u16,
    pub reserved1: u32,
    pub reserved2: u32,
    pub subtotal_name: Option<BiffUnicodeString>,
}

/// Extended PivotTable field properties from an SXVDEx record.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxvdExRecord {
    pub flags: SxvdExFlags,
    pub auto_show_count: u8,
    pub auto_sort_data_item: i16,
    pub auto_show_data_item: i16,
    pub number_format_index: u16,
    pub optional: Option<SxvdExOptional>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxAxis: u16 {
        const ROW = 0x0001;
        const COLUMN = 0x0002;
        const PAGE = 0x0004;
        const DATA = 0x0008;
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SxvdSubtotalFlags: u16 {
        const DEFAULT = 1 << 0;
        const SUM = 1 << 1;
        const COUNTA = 1 << 2;
        const AVERAGE = 1 << 3;
        const MAX = 1 << 4;
        const MIN = 1 << 5;
        const PRODUCT = 1 << 6;
        const COUNT = 1 << 7;
        const STDEV = 1 << 8;
        const STDEVP = 1 << 9;
        const VARIANCE = 1 << 10;
        const VARIANCEP = 1 << 11;
    }
}

/// Pivot field properties from an Sxvd record.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SxvdRecord {
    pub axis: SxAxis,
    pub subtotal_count: u16,
    pub subtotal_flags: SxvdSubtotalFlags,
    pub item_count: i16,
    pub declared_name_length: u16,
    pub name: Option<BiffUnicodeString>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct FrtHeader {
    pub record_type: u16,
    #[sdk(bitflags = "u16")]
    pub flags: FrtFlags,
    pub reserved: u64,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct FrtHeaderOld {
    pub record_type: u16,
    #[sdk(bitflags = "u16")]
    pub flags: FrtFlags,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CFrtId {
    pub first_record_type: u16,
    pub last_record_type: u16,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartFrtInfoRecord {
    pub header: FrtHeaderOld,
    pub originator_version: u8,
    pub writer_version: u8,
    pub record_type_range_count: u16,
    #[sdk(count = "record_type_range_count")]
    pub record_type_ranges: Vec<CFrtId>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartCatLabFlags: u16 {
        const AUTO_CATEGORY_LABEL_COUNT = 0x0001;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ChartCatLabRecord {
    pub header: FrtHeaderOld,
    pub offset_percent: u16,
    pub alignment: u16,
    pub flags: ChartCatLabFlags,
    /// Older producers omit the final reserved word.
    pub reserved: Option<u16>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartStartObjectRecord {
    pub header: FrtHeaderOld,
    pub object_kind: u16,
    pub object_context: u16,
    pub object_instance1: u16,
    pub object_instance2: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ChartEndObjectRecord {
    pub header: FrtHeaderOld,
    pub object_kind: u16,
    /// Older producers sometimes omit all three unused words.
    pub unused: Option<[u16; 3]>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct Compat12Record {
    pub header: FrtHeader,
    pub no_compatibility_check: u32,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct PlvFlags: u16 {
        const PAGE_LAYOUT_VIEW = 0x0001;
        const RULER_VISIBLE = 0x0002;
        const WHITESPACE_HIDDEN = 0x0004;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct PlvRecord {
    pub header: FrtHeader,
    pub zoom_scale: u16,
    #[sdk(bitflags = "u16")]
    pub flags: PlvFlags,
}

bitflags::bitflags! {
    /// Page Layout View flags written by Mac Excel 11.
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct PlvMacFlags: u8 {
        const MOVE = 0x01;
        const ONE_PAGE = 0x02;
        const RULER = 0x04;
        const PRINT_SCALE_NOT_SHEET_SCALE = 0x08;
    }
}

/// Page Layout View settings written by Mac Excel 11.
#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct PlvMacRecord {
    pub header: FrtHeader,
    #[sdk(bitflags = "u8")]
    pub flags: PlvMacFlags,
    pub zoom_scale: u32,
}

/// Extended line or border properties written by Mac Office 11.
#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct LnextRecord {
    pub header: FrtHeader,
    pub color_rgb: u32,
    pub opacity: u32,
    pub line_width: u32,
}

/// Extended chart marker properties written by Mac Office 11.
#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct MkrExtRecord {
    pub header: FrtHeader,
    pub foreground_rgb: u32,
    pub background_rgb: u32,
    pub opacity: u32,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct CrtCoOptFlags: u16 {
        const SHADED = 0x0001;
        const GRAYSCALE = 0x0002;
    }
}

/// Chart series color options written by Mac Office 11.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CrtCoOptRecord {
    pub header: FrtHeader,
    pub color_scheme: u32,
    pub flags: CrtCoOptFlags,
    /// Two-byte zero padding emitted by one Mac producer in the corpus.
    pub compatibility_padding: Option<u16>,
}

/// Producer architecture identifier carried in the internal FRTArchId$ record.
#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct FrtArchIdRecord {
    pub header: FrtHeader,
    pub architecture_id: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u8")]
pub enum LhGraphType {
    Xy = 0,
    Bar = 1,
    Pie = 2,
    Line = 4,
    StackedBar = 5,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u8")]
pub enum LhGraphGrid {
    None = 0,
    Horizontal = 1,
    Vertical = 2,
    Both = 3,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u8")]
pub enum LhGraphLineFormat {
    None = 0,
    Line = 1,
    Symbol = 2,
    LineAndSymbol = 3,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u8")]
pub enum LhGraphLabelAlignment {
    Center = 0,
    Right = 1,
    Below = 2,
    Left = 3,
    Above = 4,
}

/// The WKS GRAPH tail embedded in an LHRECORD graph-view subrecord.
#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct LhGraphViewCore {
    pub reference_defined: [u16; 13],
    pub graph_type: LhGraphType,
    pub grid: LhGraphGrid,
    /// 0 selects monochrome and 0xFF selects color in the WKS format.
    pub color: u8,
    pub line_format_a: LhGraphLineFormat,
    pub line_format_b: LhGraphLineFormat,
    pub line_format_c: LhGraphLineFormat,
    pub line_format_d: LhGraphLineFormat,
    pub line_format_e: LhGraphLineFormat,
    pub line_format_f: LhGraphLineFormat,
    pub label_alignment_a: LhGraphLabelAlignment,
    pub label_alignment_b: LhGraphLabelAlignment,
    pub label_alignment_c: LhGraphLabelAlignment,
    pub label_alignment_d: LhGraphLabelAlignment,
    pub label_alignment_e: LhGraphLabelAlignment,
    pub label_alignment_f: LhGraphLabelAlignment,
    /// 0 selects automatic scaling and 0xFF selects manual scaling.
    pub x_scale: u8,
    pub x_lower_limit_bits: u64,
    pub x_upper_limit_bits: u64,
    /// 0 selects automatic scaling and 0xFF selects manual scaling.
    pub y_scale: u8,
    pub y_lower_limit_bits: u64,
    pub y_upper_limit_bits: u64,
    pub first_title: [u8; 40],
    pub second_title: [u8; 40],
    pub x_title: [u8; 40],
    pub y_title: [u8; 40],
    pub legend_a: [u8; 20],
    pub legend_b: [u8; 20],
    pub legend_c: [u8; 20],
    pub legend_d: [u8; 20],
    pub legend_e: [u8; 20],
    pub legend_f: [u8; 20],
    pub x_format: u8,
    pub y_format: u8,
    pub skip_factor: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct LhGraphViewRecord {
    pub core: LhGraphViewCore,
    /// Three producer-specific words appended after the specified WKS GRAPH tail.
    pub compatibility_extension: Option<[u16; 3]>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LhMarginKind {
    Left,
    Right,
    Top,
    Bottom,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
#[sdk(repr = "u16")]
pub enum LhTableType {
    None = 0,
    Table1 = 1,
    Table2 = 2,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LhReservedKind {
    Type1,
    Type10,
    Type12,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum LhSubrecordData {
    HeaderString(Vec<u8>),
    FooterString(Vec<u8>),
    Margin {
        kind: LhMarginKind,
        value_bits: u64,
    },
    GraphView(Box<LhGraphViewRecord>),
    GlobalColumnWidth(u16),
    TableType(LhTableType),
    Reserved {
        kind: LhReservedKind,
        words: Vec<u16>,
    },
    /// Observed Excel extension beyond the published 0x01..=0x0C table.
    UndocumentedType13(u16),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LhRecord {
    pub subrecords: Vec<LhSubrecordData>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CrtLayout12Record {
    pub header: FrtHeader,
    pub checksum: u32,
    /// Bit 0 is unused, bits 1..=4 contain the automatic layout type.
    pub layout_flags: u16,
    pub x_mode: u16,
    pub y_mode: u16,
    pub width_mode: u16,
    pub height_mode: u16,
    pub x_bits: u64,
    pub y_bits: u64,
    pub width_bits: u64,
    pub height_bits: u64,
    pub reserved: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct CrtLayout12AFlags: u16 {
        const LAYOUT_TARGET_INNER = 0x0001;
    }
}

/// Plot-area layout information from a CrtLayout12A record.
#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CrtLayout12ARecord {
    pub header: FrtHeader,
    pub checksum: u32,
    #[sdk(bitflags = "u16")]
    pub flags: CrtLayout12AFlags,
    pub top_left_x: i16,
    pub top_left_y: i16,
    pub bottom_right_x: i16,
    pub bottom_right_y: i16,
    pub x_mode: u16,
    pub y_mode: u16,
    pub width_mode: u16,
    pub height_mode: u16,
    pub x_bits: u64,
    pub y_bits: u64,
    pub width_bits: u64,
    pub height_bits: u64,
    pub reserved: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct MtrSettingsRecord {
    pub header: FrtHeader,
    pub enabled: u32,
    pub user_set_thread_count: u32,
    pub thread_count: i32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ForceFullCalculationRecord {
    pub header: FrtHeader,
    pub ignore_dependencies: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct XfCrcRecord {
    pub header: FrtHeader,
    pub reserved: u16,
    pub xf_count: u16,
    pub checksum: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CompressPicturesRecord {
    pub header: FrtHeader,
    pub auto_compress_pictures: u32,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XmlTkData {
    Start,
    End,
    Boolean { value: u8, unused: u8 },
    Double { unused: u32, value_bits: u64 },
    DWord(i32),
    String(Vec<u16>),
    Token(u16),
    Blob(Vec<u8>),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlTkRecord {
    pub unused: u8,
    pub tag: u16,
    pub data: XmlTkData,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlTkChain {
    pub record_version: u8,
    pub unused: u8,
    pub parent: u16,
    pub records: Vec<XmlTkRecord>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CrtMlFrtRecord {
    pub header: FrtHeader,
    pub chain: XmlTkChain,
    pub unused: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct StartBlockRecord {
    pub header: FrtHeaderOld,
    pub object_kind: u16,
    pub object_context: u16,
    pub object_instance1: u16,
    pub object_instance2: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EndBlockRecord {
    pub header: FrtHeaderOld,
    pub object_kind: u16,
    pub optional_unused: Option<[u16; 3]>,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct ThemeRecord {
    pub header: FrtHeader,
    pub version: u32,
    /// ECMA-376 theme part bytes for custom themes; empty for the built-in theme.
    #[sdk(remaining)]
    pub contents: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct TableStylesRecord {
    pub header: FrtHeader,
    pub total_style_count: u32,
    pub default_table_style_character_count: u16,
    pub default_pivot_style_character_count: u16,
    #[sdk(count = "default_table_style_character_count")]
    pub default_table_style: Vec<u16>,
    #[sdk(count = "default_pivot_style_character_count")]
    pub default_pivot_style: Vec<u16>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ExtendedHeaderFooterFlags: u16 {
        const DIFFERENT_ODD_EVEN = 0x0001;
        const DIFFERENT_FIRST = 0x0002;
        const SCALE_WITH_DOCUMENT = 0x0004;
        const ALIGN_MARGINS = 0x0008;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExtendedHeaderFooterRecord {
    pub header: FrtHeader,
    pub sheet_view_guid: [u8; 16],
    pub flags: ExtendedHeaderFooterFlags,
    pub even_header_character_count: u16,
    pub even_footer_character_count: u16,
    pub first_header_character_count: u16,
    pub first_footer_character_count: u16,
    pub even_header: Option<BiffUnicodeString>,
    pub even_footer: Option<BiffUnicodeString>,
    pub first_header: Option<BiffUnicodeString>,
    pub first_footer: Option<BiffUnicodeString>,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct ShapePropsStreamRecord {
    pub header: FrtHeader,
    pub object_context: u16,
    pub unused: u16,
    pub checksum: u32,
    pub xml_length: u32,
    /// ECMA-376 chart shape-property XML stream bytes.
    #[sdk(count = "xml_length")]
    pub xml: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct TextPropsStreamRecord {
    pub header: FrtHeader,
    pub checksum: u32,
    pub xml_length: u32,
    /// ECMA-376 chart text-property XML stream bytes.
    #[sdk(count = "xml_length")]
    pub xml: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct RichTextStreamRecord {
    pub header: FrtHeader,
    pub checksum: u32,
    pub xml_length: u32,
    #[sdk(count = "xml_length")]
    pub xml: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct GuidTypeLibRecord {
    pub header: FrtHeader,
    pub guid: [u8; 16],
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NameCommentRecord {
    pub header: FrtHeader,
    pub declared_name_length: u16,
    pub declared_comment_length: u16,
    pub name: BiffUnicodeString,
    pub comment: BiffUnicodeString,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct HfPictureFlags: u8 {
        const DRAWING = 0x01;
        const DRAWING_GROUP = 0x02;
        const CONTINUATION = 0x04;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HfPictureRecord {
    pub header: FrtHeader,
    pub flags: HfPictureFlags,
    pub reserved: u8,
    pub drawing: MsoDrawingData,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct EnhancedProtectionFlags: u32 {
        const OBJECTS = 1 << 0;
        const SCENARIOS = 1 << 1;
        const FORMAT_CELLS = 1 << 2;
        const FORMAT_COLUMNS = 1 << 3;
        const FORMAT_ROWS = 1 << 4;
        const INSERT_COLUMNS = 1 << 5;
        const INSERT_ROWS = 1 << 6;
        const INSERT_HYPERLINKS = 1 << 7;
        const DELETE_COLUMNS = 1 << 8;
        const DELETE_ROWS = 1 << 9;
        const SELECT_LOCKED_CELLS = 1 << 10;
        const SORT = 1 << 11;
        const AUTO_FILTER = 1 << 12;
        const PIVOT_TABLES = 1 << 13;
        const SELECT_UNLOCKED_CELLS = 1 << 14;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PbStringCharacters {
    Ansi(Vec<u8>),
    Unicode(Vec<u16>),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PbString {
    /// Includes the terminating NUL as stored by MS-OSHARED.
    pub characters: PbStringCharacters,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FactoidType {
    pub id: u32,
    pub uri: PbString,
    pub tag: PbString,
    pub download_url: PbString,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PropertyBagStore {
    pub factoid_types: Vec<FactoidType>,
    pub header_size: u16,
    pub version: u16,
    pub factoid_count: u32,
    pub strings: Vec<PbString>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum FeatureHeaderData {
    None,
    EnhancedProtection(EnhancedProtectionFlags),
    PropertyBagStore(PropertyBagStore),
    /// Explicitly malformed header marker and bounded remainder from a damaged fixture.
    Malformed {
        marker: u32,
        payload: Vec<u8>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FeatureHeaderRecord {
    pub header: FrtHeader,
    pub shared_feature_type: u16,
    pub reserved: u8,
    pub data: FeatureHeaderData,
}

bitflags::bitflags! {
    /// Formula conditions ignored by the background error checker (FFErrorCheck).
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct FormulaErrorCheckFlags: u32 {
        const CALCULATION_ERRORS = 1 << 0;
        const EMPTY_CELL_REFERENCES = 1 << 1;
        const NUMBERS_STORED_AS_TEXT = 1 << 2;
        const INCONSISTENT_RANGES = 1 << 3;
        const INCONSISTENT_FORMULAS = 1 << 4;
        const TEXT_DATE_FORMATS = 1 << 5;
        const UNPROTECTED_FORMULAS = 1 << 6;
        const DATA_VALIDATION = 1 << 7;
    }
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct FeatureProtectionFlags: u32 {
        const HAS_SECURITY_DESCRIPTOR = 1 << 0;
    }
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SecurityDescriptorControl: u16 {
        const OWNER_DEFAULTED = 0x0001;
        const GROUP_DEFAULTED = 0x0002;
        const DACL_PRESENT = 0x0004;
        const DACL_DEFAULTED = 0x0008;
        const SACL_PRESENT = 0x0010;
        const SACL_DEFAULTED = 0x0020;
        const DACL_AUTO_INHERIT_REQUIRED = 0x0100;
        const SACL_AUTO_INHERIT_REQUIRED = 0x0200;
        const DACL_AUTO_INHERITED = 0x0400;
        const SACL_AUTO_INHERITED = 0x0800;
        const DACL_PROTECTED = 0x1000;
        const SACL_PROTECTED = 0x2000;
        const RM_CONTROL_VALID = 0x4000;
        const SELF_RELATIVE = 0x8000;
    }
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct AceFlags: u8 {
        const OBJECT_INHERIT = 0x01;
        const CONTAINER_INHERIT = 0x02;
        const NO_PROPAGATE_INHERIT = 0x04;
        const INHERIT_ONLY = 0x08;
        const INHERITED = 0x10;
        const SUCCESSFUL_ACCESS = 0x40;
        const FAILED_ACCESS = 0x80;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SecurityIdentifier {
    pub revision: u8,
    pub identifier_authority: [u8; 6],
    pub sub_authorities: Vec<u32>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BasicAceType {
    AccessAllowed,
    AccessDenied,
    SystemAudit,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BasicAce {
    pub ace_type: BasicAceType,
    pub flags: AceFlags,
    pub access_mask: u32,
    pub trustee: SecurityIdentifier,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AccessControlList {
    pub revision: u8,
    pub reserved1: u8,
    pub reserved2: u16,
    pub entries: Vec<BasicAce>,
    /// Bytes in the declared ACL size after its typed ACE array.
    pub padding: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OffsetSecurityIdentifier {
    pub offset: u32,
    pub value: SecurityIdentifier,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OffsetAccessControlList {
    pub offset: u32,
    pub value: AccessControlList,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SecurityDescriptorPadding {
    pub offset: u32,
    pub bytes: Vec<u8>,
}

/// A self-relative SECURITY_DESCRIPTOR from MS-DTYP section 2.4.6.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SecurityDescriptor {
    pub revision: u8,
    pub reserved: u8,
    pub control: SecurityDescriptorControl,
    pub owner: Option<OffsetSecurityIdentifier>,
    pub group: Option<OffsetSecurityIdentifier>,
    pub sacl: Option<OffsetAccessControlList>,
    pub dacl: Option<OffsetAccessControlList>,
    /// Explicit inter-component alignment/reserved regions in offset order.
    pub padding: Vec<SecurityDescriptorPadding>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SecurityDescriptorContainer {
    pub declared_size: u32,
    pub descriptor: SecurityDescriptor,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FeatureProtection {
    pub flags: FeatureProtectionFlags,
    pub password_verifier: u32,
    pub title: XlUnicodeString,
    pub security_descriptor: Option<SecurityDescriptorContainer>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct SmartTagProperty {
    pub key_index: u32,
    pub value_index: u32,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct SmartTagPropertyBag {
    pub factoid_type_id: u16,
    pub property_count: u16,
    pub reserved: u16,
    #[sdk(count = "property_count")]
    pub properties: Vec<SmartTagProperty>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct FactoidDataFlags: u8 {
        const DELETED = 1 << 0;
        const XML_BASED = 1 << 1;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FactoidData {
    pub flags: FactoidDataFlags,
    pub property_bag: SmartTagPropertyBag,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FeatureSmartTag {
    pub hash_value: u32,
    pub factoids: Vec<FactoidData>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum FeatureData {
    Protection(Box<FeatureProtection>),
    FormulaErrors(FormulaErrorCheckFlags),
    SmartTags(FeatureSmartTag),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FeatureRecord {
    pub header: FrtHeader,
    pub shared_feature_type: u16,
    pub reserved1: u8,
    pub reserved2: u32,
    pub feature_data_size: u32,
    pub reserved3: u16,
    pub references: Vec<CellRange>,
    pub data: FeatureData,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct BookExtFlags: u32 {
        const DONT_AUTO_RECOVER = 1 << 0;
        const HIDE_PIVOT_LIST = 1 << 1;
        const FILTER_PRIVACY = 1 << 2;
        const EMBED_FACTOIDS = 1 << 3;
        const FACTOID_DISPLAY_MASK = 0x0000_0030;
        const SAVED_DURING_RECOVERY = 1 << 6;
        const CREATED_VIA_MINIMAL_SAVE = 1 << 7;
        const OPENED_VIA_DATA_RECOVERY = 1 << 8;
        const OPENED_VIA_SAFE_LOAD = 1 << 9;
    }
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct BookExtConditional11Flags: u8 {
        const WARN_ABOUT_SOLUTION = 0x01;
        const SHOW_INK_ANNOTATION = 0x02;
    }
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct BookExtConditional12Flags: u8 {
        const PUBLISHED_BOOK_ITEMS = 0x02;
        const SHOW_PIVOT_CHART_FILTER = 0x04;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BookExtRecord {
    pub header: FrtHeader,
    pub declared_size: u32,
    pub flags: BookExtFlags,
    pub conditional11: Option<BookExtConditional11Flags>,
    pub conditional12: Option<BookExtConditional12Flags>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExtSstRecord {
    pub strings_per_bucket: u16,
    pub buckets: Vec<ExtSstBucket>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ExtSstBucket {
    pub stream_offset: u32,
    pub record_offset: u16,
    pub reserved: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartAreaFlags: u16 {
        const AUTO = 0x0001;
        const INVERT_NEGATIVE = 0x0002;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartAreaFormatRecord {
    pub foreground_rgb: u32,
    pub background_rgb: u32,
    pub fill_pattern: u16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartAreaFlags,
    pub foreground_color_index: u16,
    pub background_color_index: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartRecord {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartAttachedLabelFlags: u16 {
        const SHOW_VALUE = 0x0001;
        const SHOW_PERCENT = 0x0002;
        const SHOW_LABEL_AND_PERCENT = 0x0004;
        const SHOW_LABEL = 0x0010;
        const SHOW_BUBBLE_SIZE = 0x0020;
        const SHOW_SERIES_NAME = 0x0040;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartAttachedLabelRecord {
    #[sdk(bitflags = "u16")]
    pub flags: ChartAttachedLabelFlags,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartDataFormatFlags: u16 {
        const EXCEL4 = 0x0001;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartDataFormatRecord {
    pub point_index: u16,
    pub series_index: u16,
    pub series_number: u16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartDataFormatFlags,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartFormatFlags: u16 {
        const VARIED = 0x0001;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartFormatRecord {
    pub reserved1: u32,
    pub reserved2: u32,
    pub reserved3: u32,
    pub reserved4: u32,
    #[sdk(bitflags = "u16")]
    pub flags: ChartFormatFlags,
    pub drawing_order: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartLegendFlags: u16 {
        const AUTO_POSITION = 0x0001;
        const AUTO_SERIES = 0x0002;
        const AUTO_X = 0x0004;
        const AUTO_Y = 0x0008;
        const VERTICAL = 0x0010;
        const DATA_TABLE = 0x0020;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartLegendRecord {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
    pub legend_type: u8,
    pub spacing: u8,
    #[sdk(bitflags = "u16")]
    pub flags: ChartLegendFlags,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartAxisRecord {
    pub axis_type: u16,
    pub reserved1: u32,
    pub reserved2: u32,
    pub reserved3: u32,
    pub reserved4: u32,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartTickFlags: u16 {
        const AUTO_TEXT_COLOR = 0x0001;
        const AUTO_TEXT_BACKGROUND = 0x0002;
        const ROTATION = 0x001c;
        const AUTO_ROTATION = 0x0020;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartTickRecord {
    pub major_tick_type: u8,
    pub minor_tick_type: u8,
    pub label_position: u8,
    pub background_mode: u8,
    pub label_color_rgb: u32,
    pub reserved1: u32,
    pub reserved2: u32,
    pub reserved3: u32,
    pub reserved4: u32,
    #[sdk(bitflags = "u16")]
    pub flags: ChartTickFlags,
    pub color_index: u16,
    pub reserved5: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartValueRangeFlags: u16 {
        const AUTO_MINIMUM = 0x0001;
        const AUTO_MAXIMUM = 0x0002;
        const AUTO_MAJOR = 0x0004;
        const AUTO_MINOR = 0x0008;
        const AUTO_CROSS = 0x0010;
        const LOGARITHMIC = 0x0020;
        const REVERSED = 0x0040;
        const MAXIMUM_CROSS = 0x0080;
        const BIT9 = 0x0100;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartValueRangeRecord {
    pub minimum_bits: u64,
    pub maximum_bits: u64,
    pub major_unit_bits: u64,
    pub minor_unit_bits: u64,
    pub cross_value_bits: u64,
    #[sdk(bitflags = "u16")]
    pub flags: ChartValueRangeFlags,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartLabelRangeFlags: u16 {
        const BETWEEN_CATEGORIES = 0x0001;
        const MAXIMUM_CROSS = 0x0002;
        const REVERSED = 0x0004;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartLabelRangeRecord {
    pub crossing_point: u16,
    pub label_frequency: u16,
    pub tick_frequency: u16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartLabelRangeFlags,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartAxisLineRecord {
    pub line_kind: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartAxisParentRecord {
    pub axis_group: u16,
    pub unused1: u32,
    pub unused2: u32,
    pub unused3: u32,
    pub unused4: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartPositionRecord {
    pub top_left_mode: u16,
    pub bottom_right_mode: u16,
    pub x1: i32,
    pub y1: i32,
    pub x2: i32,
    pub y2: i32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartFontBasisRecord {
    pub x_basis: i16,
    pub y_basis: i16,
    pub height_basis: i16,
    pub scale: i16,
    pub font_index: i16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartFontRecord {
    pub font_index: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartLineFlags: u16 {
        const AUTO = 0x0001;
        const AXIS_ON = 0x0004;
        const AUTO_COLOR = 0x0008;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartLineFormatRecord {
    pub color_rgb: u32,
    pub line_style: u16,
    pub weight: i16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartLineFlags,
    pub color_index: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartMarkerFlags: u16 {
        const AUTO = 0x0001;
        const HIDE_INTERIOR = 0x0010;
        const HIDE_BORDER = 0x0020;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartMarkerFormatRecord {
    pub foreground_rgb: u32,
    pub background_rgb: u32,
    pub marker_type: u16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartMarkerFlags,
    pub foreground_color_index: u16,
    pub background_color_index: u16,
    pub marker_size_twips: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartObjectLinkRecord {
    pub object_type: u16,
    pub series_index: u16,
    pub category_index: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartPieFormatRecord {
    pub explode_percent: i16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartPlotGrowthRecord {
    pub horizontal_growth: u32,
    pub vertical_growth: u32,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartLinkedDataFlags: u16 {
        const CUSTOM_NUMBER_FORMAT = 0x0001;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ChartLinkedDataRecord {
    pub link_type: u8,
    pub reference_type: u8,
    pub flags: ChartLinkedDataFlags,
    pub number_format_index: u16,
    pub formula: FormulaTokenStream,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ChartSeriesListRecord {
    pub series_numbers: Vec<u16>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartBarFlags: u16 {
        const HORIZONTAL = 0x0001;
        const STACKED = 0x0002;
        const DISPLAY_AS_PERCENTAGE = 0x0004;
        const SHADOW = 0x0008;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartBarRecord {
    pub bar_space: i16,
    pub category_space: i16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartBarFlags,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartLineGroupFlags: u16 {
        const STACKED = 0x0001;
        const DISPLAY_AS_PERCENTAGE = 0x0002;
        const SHADOW = 0x0004;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartLineRecord {
    #[sdk(bitflags = "u16")]
    pub flags: ChartLineGroupFlags,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartPieFlags: u16 {
        const SHADOW = 0x0001;
        const SHOW_LEADER_LINES = 0x0002;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartPieRecord {
    pub starting_angle: u16,
    pub doughnut_hole_percent: u16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartPieFlags,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartScatterFlags: u16 {
        const BUBBLES = 0x0001;
        const SHOW_NEGATIVE_BUBBLES = 0x0002;
        const SHADOW = 0x0004;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartScatterRecord {
    pub bubble_size_ratio: u16,
    pub bubble_size_representation: u16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartScatterFlags,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartCrtLineRecord {
    pub line_type: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartCrtLinkRecord {
    pub unused: [u8; 10],
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartAreaRecordFlags: u16 {
        const STACKED = 0x0001;
        const DISPLAY_AS_PERCENTAGE = 0x0002;
        const SHADOW = 0x0004;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartAreaRecord {
    #[sdk(bitflags = "u16")]
    pub flags: ChartAreaRecordFlags,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartDropBarRecord {
    pub gap_width_percent: i16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct Chart3DFlags: u16 {
        const PERSPECTIVE = 0x0001;
        const CLUSTERED = 0x0002;
        const AUTO_SCALING = 0x0004;
        const NOT_PIE_CHART = 0x0010;
        const TWO_DIMENSIONAL_WALLS = 0x0020;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct Chart3DRecord {
    pub rotation: i16,
    pub elevation: i16,
    pub field_of_view: i16,
    pub height_percent: u16,
    pub depth_percent: i16,
    pub gap_percent: u16,
    #[sdk(bitflags = "u16")]
    pub flags: Chart3DFlags,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartAlRunsRecord {
    pub run_count: u16,
    #[sdk(count = "run_count")]
    pub runs: Vec<FormatRun>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartSurfFlags: u16 {
        const FILL_SURFACE = 0x0001;
        const PHONG_SHADING = 0x0002;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartSurfRecord {
    #[sdk(bitflags = "u16")]
    pub flags: ChartSurfFlags,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartLegendExceptionFlags: u16 {
        const DELETED = 0x0001;
        const FORMATTED_LABEL = 0x0002;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartLegendExceptionRecord {
    pub legend_entry_index: u16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartLegendExceptionFlags,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartSeriesParentRecord {
    pub series_index: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartSeriesAuxTrendRecord {
    pub regression_type: u8,
    pub order_or_period: u8,
    pub intercept_bits: u64,
    pub show_equation: u8,
    pub show_r_squared: u8,
    pub forecast_bits: u64,
    pub backcast_bits: u64,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartSeriesFormatFlags: u16 {
        const SMOOTHED_LINE = 0x0001;
        const THREE_DIMENSIONAL_BUBBLES = 0x0002;
        const SHADOW = 0x0004;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartSeriesFormatRecord {
    #[sdk(bitflags = "u16")]
    pub flags: ChartSeriesFormatFlags,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartSeriesAuxErrorBarRecord {
    pub direction: u8,
    pub source_type: u8,
    pub tee_top: u8,
    pub reserved: u8,
    pub value_bits: u64,
    pub value_count: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartClrtClientRecord {
    pub color_count: i16,
    pub series_or_data_point_color: u32,
    pub chart_area_color: u32,
    pub plot_area_color: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartDefaultTextRecord {
    pub category_data_type: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartFrameFlags: u16 {
        const AUTO_SIZE = 0x0001;
        const AUTO_POSITION = 0x0002;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartFrameRecord {
    pub border_type: u16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartFrameFlags,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartSheetPropertyFlags: u16 {
        const MANUALLY_FORMATTED = 0x0001;
        const PLOT_VISIBLE_ONLY = 0x0002;
        const DO_NOT_SIZE_WITH_WINDOW = 0x0004;
        const DEFAULT_PLOT_DIMENSIONS = 0x0008;
        const AUTO_PLOT_AREA = 0x0010;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartSheetPropertiesRecord {
    #[sdk(bitflags = "u16")]
    pub flags: ChartSheetPropertyFlags,
    pub empty_cell_display_mode: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartSeriesGroupIndexRecord {
    pub chart_group_index: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartAxisUsedRecord {
    pub axis_count: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartNumberFormatIndexRecord {
    pub format_index: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartAxisOptionFlags: u16 {
        const DEFAULT_MINIMUM = 0x0001;
        const DEFAULT_MAXIMUM = 0x0002;
        const DEFAULT_MAJOR = 0x0004;
        const DEFAULT_MINOR_UNIT = 0x0008;
        const IS_DATE = 0x0010;
        const DEFAULT_BASE = 0x0020;
        const DEFAULT_CROSS = 0x0040;
        const DEFAULT_DATE_SETTINGS = 0x0080;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartAxisOptionsRecord {
    pub minimum_category: i16,
    pub maximum_category: i16,
    pub major_unit_value: i16,
    pub major_unit: i16,
    pub minor_unit_value: i16,
    pub minor_unit: i16,
    pub base_unit: i16,
    pub crossing_point: i16,
    #[sdk(bitflags = "u16")]
    pub flags: ChartAxisOptionFlags,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartDatFlags: u16 {
        const HORIZONTAL_BORDER = 0x0001;
        const VERTICAL_BORDER = 0x0002;
        const BORDER = 0x0004;
        const SHOW_SERIES_KEY = 0x0008;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartDatRecord {
    #[sdk(bitflags = "u16")]
    pub flags: ChartDatFlags,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartSeriesIndexRecord {
    pub index: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartSeriesRecord {
    pub category_data_type: u16,
    pub value_data_type: u16,
    pub category_count: u16,
    pub value_count: u16,
    pub bubble_data_type: u16,
    pub bubble_count: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ChartSeriesTextRecord {
    pub reserved: u16,
    pub text: BiffUnicodeString,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct ChartTextFlags: u16 {
        const AUTO_COLOR = 0x0001;
        const SHOW_KEY = 0x0002;
        const SHOW_VALUE = 0x0004;
        const AUTO_TEXT = 0x0010;
        const GENERATED = 0x0020;
        const DELETED = 0x0040;
        const AUTO_MODE = 0x0080;
        const SHOW_LABEL_AND_PERCENT = 0x0800;
        const SHOW_PERCENT = 0x1000;
        const SHOW_BUBBLE_SIZE = 0x2000;
        const SHOW_LABEL = 0x4000;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct ChartTextRecord {
    pub horizontal_alignment: u8,
    pub vertical_alignment: u8,
    pub background_mode: u16,
    pub text_rgb: u32,
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
    #[sdk(bitflags = "u16")]
    pub flags: ChartTextFlags,
    pub color_index: u16,
    pub reading_order_flags: u16,
    pub rotation: u16,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct FrtFlags: u16 {
        const HAS_CELL_RANGE = 0x0001;
        const ALERT_ON_UNRECOGNIZED_SAVE = 0x0002;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XfExtRecord {
    pub header: FrtHeader,
    pub reserved1: u16,
    pub xf_index: u16,
    pub reserved2: u16,
    pub properties: Vec<ExtProperty>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExtProperty {
    pub data: ExtPropertyData,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ExtPropertyData {
    FullColor {
        property_type: u16,
        color: FullColorExt,
    },
    Gradient {
        payload: Vec<u8>,
    },
    FontScheme(ExtFontScheme),
    Indentation(u16),
    Unknown {
        property_type: u16,
        payload: Vec<u8>,
    },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExtFontScheme {
    Byte(u8),
    Word(u16),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct FullColorExt {
    pub color_type: u16,
    pub tint: i16,
    pub color_value: u32,
    pub unused: u64,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct StyleExtFlags: u8 {
        const BUILT_IN = 0x01;
        const HIDDEN = 0x02;
        const CUSTOM = 0x04;
    }
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct LpWideString {
    pub character_count: u16,
    #[sdk(count = "character_count")]
    pub characters: Vec<u16>,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct StyleExtRecord {
    pub header: FrtHeader,
    #[sdk(bitflags = "u8")]
    pub flags: StyleExtFlags,
    pub category: u8,
    pub built_in_data: u16,
    pub name: LpWideString,
    pub xf_properties_reserved: u16,
    pub property_count: u16,
    #[sdk(count = "property_count")]
    pub properties: Vec<XfProperty>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct DxfRecordFlags: u16 {
        const UNUSED1 = 0x0001;
        const NEW_BORDER = 0x0002;
        const UNUSED2 = 0x0004;
    }
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct DxfRecord {
    pub header: FrtHeader,
    #[sdk(bitflags = "u16")]
    pub flags: DxfRecordFlags,
    pub xf_properties_reserved: u16,
    pub property_count: u16,
    #[sdk(count = "property_count")]
    pub properties: Vec<XfProperty>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XfProperty {
    pub property_type: u16,
    pub data: XfPropertyData,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XfPropertyData {
    Color(XfPropColor),
    Border(XfPropBorder),
    Byte(u8),
    Word(u16),
    DWord(u32),
    WideString(LpWideString),
    FontScheme(u8),
    Unparsed(Vec<u8>),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct XfPropColor {
    /// Bit 0 is fValidRGBA; bits 1..7 are xclrType.
    pub flags_and_color_type: u8,
    pub color_index: u8,
    pub tint: i16,
    pub rgba: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct XfPropBorder {
    pub color: XfPropColor,
    pub border_style: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DbCellRecord {
    pub row_offset: u32,
    pub cell_offsets: Vec<u16>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BiffUnicodeString {
    /// Exact option flags. Bit 0 selects UTF-16; other bits are retained.
    pub flags: u8,
    pub characters: XlStringCharacters,
    /// Compatibility byte found after an odd-length UTF-16 character sequence.
    pub trailing_byte: Option<u8>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct FontAttributes: u16 {
        const ITALIC = 0x0002;
        const STRIKEOUT = 0x0008;
        const MAC_OUTLINE = 0x0010;
        const MAC_SHADOW = 0x0020;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FontRecord {
    pub height_twips: u16,
    pub attributes: FontAttributes,
    pub color_index: u16,
    pub bold_weight: u16,
    pub escapement: u16,
    pub underline: u8,
    pub family: u8,
    pub charset: u8,
    pub reserved: u8,
    pub name: BiffUnicodeString,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormatRecord {
    pub format_index: u16,
    pub declared_character_count: u16,
    pub format_string: BiffUnicodeString,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StyleRecord {
    /// Low 12 bits are ixfe; bit 15 is fBuiltIn; all bits are retained.
    pub xf_and_flags: u16,
    pub data: StyleData,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CrnCountRecord {
    /// Exact signed/flagged count word used by historical producers.
    pub count_word: u16,
    pub sheet_table_index: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct DefaultRowHeightRecord {
    pub flags: u16,
    pub height: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct WriteAccessRecord {
    pub declared_character_count: u16,
    pub flags: u8,
    pub name: XlStringCharacters,
    pub trailing_name_byte: Option<u8>,
    pub unused: Vec<u8>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct Window2Flags: u16 {
        const DISPLAY_FORMULAS = 0x0001;
        const DISPLAY_GRIDLINES = 0x0002;
        const DISPLAY_ROW_COLUMN_HEADINGS = 0x0004;
        const FREEZE_PANES = 0x0008;
        const DISPLAY_ZEROS = 0x0010;
        const DEFAULT_HEADER = 0x0020;
        const RIGHT_TO_LEFT = 0x0040;
        const DISPLAY_OUTLINE = 0x0080;
        const FREEZE_NO_SPLIT = 0x0100;
        const SELECTED = 0x0200;
        const ACTIVE = 0x0400;
        const PAGE_BREAK_PREVIEW = 0x0800;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Window2Record {
    pub flags: Window2Flags,
    pub top_row: u16,
    pub left_column: u16,
    pub header_color: u32,
    pub extension: Window2Extension,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Window2Extension {
    None,
    Zoom {
        page_break_zoom: u16,
        normal_zoom: u16,
        reserved: Option<u32>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct SelectionRecord {
    pub pane: u8,
    pub active_row: u16,
    pub active_column: u16,
    pub active_reference_index: u16,
    pub reference_count: u16,
    #[sdk(count = "reference_count")]
    pub references: Vec<SelectionReference>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct SelectionReference {
    pub first_row: u16,
    pub last_row: u16,
    pub first_column: u8,
    pub last_column: u8,
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct MergeCellsRecord {
    pub range_count: u16,
    #[sdk(count = "range_count")]
    pub ranges: Vec<CellRange>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct CellRange {
    pub first_row: u16,
    pub last_row: u16,
    pub first_column: u16,
    pub last_column: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct MmsRecord {
    pub reserved1: u8,
    pub reserved2: u8,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct PhoneticFlags: u16 {
        const TYPE_LOW = 0x0001;
        const TYPE_HIGH = 0x0002;
        const ALIGNMENT_LOW = 0x0004;
        const ALIGNMENT_HIGH = 0x0008;
    }
}

#[derive(Clone, Debug, PartialEq, Eq, SdkObject)]
pub struct PhoneticInfoRecord {
    pub font_index: u16,
    #[sdk(bitflags = "u16")]
    pub flags: PhoneticFlags,
    pub range_count: u16,
    #[sdk(count = "range_count")]
    pub ranges: Vec<CellRange>,
}

bitflags::bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct SstStringFlags: u8 {
        const HIGH_BYTE = 0x01;
        const EXTENDED = 0x04;
        const RICH_TEXT = 0x08;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SstRecord {
    pub total_string_count: u32,
    pub unique_string_count: u32,
    pub strings: Vec<SstString>,
    pub completion: SstCompletion,
    /// Explicit producer bytes after the declared string table.
    pub trailing: Vec<u8>,
    /// Exact SST/Continue physical layout, excluding continuation encoding bytes.
    pub physical_segments: Vec<SstSegmentLayout>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SstCompletion {
    Complete,
    Truncated {
        first_unparsed_string: u32,
        reason: String,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SstString {
    pub declared_character_count: u16,
    pub flags: SstStringFlags,
    pub declared_format_run_count: Option<u16>,
    pub declared_extension_length: Option<u32>,
    pub character_chunks: Vec<SstCharacterChunk>,
    pub format_runs: Vec<FormatRun>,
    pub extension: SstExtensionData,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SstExtensionData {
    None,
    ExtRst(ExtRst),
    Unparsed(Vec<u8>),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExtRst {
    pub reserved: u16,
    pub body: ExtRstBody,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ExtRstBody {
    Phonetic {
        declared_data_size: u16,
        font_index: u16,
        formatting_flags: PhoneticFlags,
        declared_run_count: u16,
        declared_character_count: u16,
        lpwide_character_count: u16,
        phonetic_text: Vec<u16>,
        runs: Vec<PhoneticRun>,
        extra_data_word: Option<u16>,
        inner_trailing: Vec<u8>,
        outer_trailing: Vec<u8>,
    },
    OldStyle {
        payload: Vec<u8>,
    },
    TruncatedPhoneticHeader {
        declared_data_size: u16,
        font_index: u16,
        formatting_flags: PhoneticFlags,
        declared_run_count: u16,
        declared_character_count: u16,
    },
    InvalidMarker {
        payload: Vec<u8>,
    },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct PhoneticRun {
    pub phonetic_text_first_character: u16,
    pub source_text_first_character: u16,
    pub source_text_character_count: u16,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SstCharacterChunk {
    pub flags: u8,
    pub characters: XlStringCharacters,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, SdkObject)]
pub struct FormatRun {
    pub character_index: u16,
    pub font_index: u16,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SstSegmentLayout {
    pub logical_byte_count: u16,
    pub continuation_encoding: Option<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct IndexRecord {
    pub reserved1: u32,
    pub first_row: u32,
    pub last_row_exclusive: u32,
    pub reserved2: u32,
    pub dbcell_offsets: Vec<u32>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StringValueRecord {
    pub declared_character_count: u16,
    pub chunks: Vec<ContinuedStringChunk>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ContinuedStringChunk {
    /// Exact option flags for this physical BIFF record segment.
    pub flags: u8,
    pub characters: XlStringCharacters,
    /// Explicit producer bytes after the declared string ends in this segment.
    pub trailing: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum StyleData {
    BuiltIn {
        style_id: u8,
        outline_level: u8,
    },
    UserDefined {
        declared_character_count: u16,
        /// None preserves the Crystal Reports zero-length form without an option byte.
        name: Option<BiffUnicodeString>,
    },
}

impl SdkRead for DimensionsRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let first_row = reader.read_u32()?;
        let last_row_exclusive = reader.read_u32()?;
        let first_column = reader.read_u16()?;
        let last_column_exclusive = reader.read_u16()?;
        let reserved = reader.read_u16()?;
        let compatibility_extra = match reader.remaining()? {
            0 => None,
            2 => Some(reader.read_u16()?),
            remaining => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("Dimensions compatibility tail has {remaining} bytes"),
                ));
            }
        };
        Ok(Self {
            first_row,
            last_row_exclusive,
            first_column,
            last_column_exclusive,
            reserved,
            compatibility_extra,
        })
    }
}

impl SdkRead for FormulaRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let cell = CellHeader::read_from(reader)?;
        let cached_result = FormulaCachedResult::from_bits(reader.read_u64()?)?;
        let flags = reader.read_u16()?;
        let calculation_chain_id = reader.read_u32()?;
        let token_length = usize::from(reader.read_u16()?);
        let mut rgce = FormulaTokenStream::from_bytes(&reader.read_vec(token_length)?)?;
        let rgcb_length = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("Formula rgcb length exceeds usize".into()))?;
        let rgcb_tail = rgce.parse_extra_data(&reader.read_vec(rgcb_length)?)?;
        Ok(Self {
            cell,
            cached_result,
            flags,
            calculation_chain_id,
            tokens: FormulaTokens { rgce, rgcb_tail },
        })
    }
}

impl SdkWrite for FormulaRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.cell.write_to(writer)?;
        writer.write_u64(self.cached_result.bits())?;
        writer.write_u16(self.flags)?;
        writer.write_u32(self.calculation_chain_id)?;
        let rgce = self.tokens.rgce.to_bytes()?;
        writer.write_u16(
            u16::try_from(rgce.len())
                .map_err(|_| Error::Limit("Formula rgce length exceeds u16".into()))?,
        )?;
        writer.write_all(&rgce)?;
        writer.write_all(&self.tokens.rgce.extra_data_to_bytes()?)?;
        writer.write_all(&self.tokens.rgcb_tail)?;
        Ok(())
    }
}

impl SdkRead for SharedFormulaRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let first_row = reader.read_u16()?;
        let last_row = reader.read_u16()?;
        let first_column = reader.read_u8()?;
        let last_column = reader.read_u8()?;
        let reserved = reader.read_u8()?;
        let use_count = reader.read_u8()?;
        let token_length = usize::from(reader.read_u16()?);
        let mut rgce = FormulaTokenStream::from_bytes(&reader.read_vec(token_length)?)?;
        let rgcb_length = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("ShrFmla rgcb length exceeds usize".into()))?;
        let rgcb_tail = rgce.parse_extra_data(&reader.read_vec(rgcb_length)?)?;
        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
            reserved,
            use_count,
            tokens: FormulaTokens { rgce, rgcb_tail },
        })
    }
}

impl SdkWrite for SharedFormulaRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.first_row)?;
        writer.write_u16(self.last_row)?;
        writer.write_u8(self.first_column)?;
        writer.write_u8(self.last_column)?;
        writer.write_u8(self.reserved)?;
        writer.write_u8(self.use_count)?;
        let rgce = self.tokens.rgce.to_bytes()?;
        writer.write_u16(
            u16::try_from(rgce.len())
                .map_err(|_| Error::Limit("ShrFmla rgce length exceeds u16".into()))?,
        )?;
        writer.write_all(&rgce)?;
        writer.write_all(&self.tokens.rgce.extra_data_to_bytes()?)?;
        writer.write_all(&self.tokens.rgcb_tail)?;
        Ok(())
    }
}

impl SdkRead for ArrayRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let first_row = reader.read_u16()?;
        let last_row = reader.read_u16()?;
        let first_column = reader.read_u8()?;
        let last_column = reader.read_u8()?;
        let flags = ArrayFlags::from_bits_retain(reader.read_u16()?);
        let unused = reader.read_u32()?;
        let token_length = usize::from(reader.read_u16()?);
        let mut rgce = FormulaTokenStream::from_bytes(&reader.read_vec(token_length)?)?;
        let rgcb_length = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("Array rgcb length exceeds usize".into()))?;
        let rgcb_tail = rgce.parse_extra_data(&reader.read_vec(rgcb_length)?)?;
        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
            flags,
            unused,
            tokens: FormulaTokens { rgce, rgcb_tail },
        })
    }
}

impl SdkWrite for ArrayRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let rgce = self.tokens.rgce.to_bytes()?;
        writer.write_u16(self.first_row)?;
        writer.write_u16(self.last_row)?;
        writer.write_u8(self.first_column)?;
        writer.write_u8(self.last_column)?;
        writer.write_u16(self.flags.bits())?;
        writer.write_u32(self.unused)?;
        writer.write_u16(
            u16::try_from(rgce.len())
                .map_err(|_| Error::Limit("Array rgce length exceeds u16".into()))?,
        )?;
        writer.write_all(&rgce)?;
        writer.write_all(&self.tokens.rgce.extra_data_to_bytes()?)?;
        writer.write_all(&self.tokens.rgcb_tail)?;
        Ok(())
    }
}

impl SdkRead for SupBookRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let sheet_count = reader.read_u16()?;
        let tag = reader.read_u16()?;
        let link = match tag {
            0x0401 if reader.remaining()? == 0 => SupBookLink::SelfReference,
            0x3a01 if reader.remaining()? == 0 => SupBookLink::AddInFunctions,
            1..=0x00ff => {
                let path = BiffUnicodeString::read(reader, usize::from(tag))?;
                let remaining_length = usize::try_from(reader.remaining()?)
                    .map_err(|_| Error::Limit("SupBook remaining length exceeds usize".into()))?;
                let remaining = reader.read_vec(remaining_length)?;
                let mut names_reader = Reader::new(Cursor::new(remaining.as_slice()))?;
                names_reader.ensure_allocation(usize::from(sheet_count), 3)?;
                let parsed_names = (|| -> Result<Vec<SupBookSheetName>> {
                    let mut names = Vec::with_capacity(usize::from(sheet_count));
                    for _ in 0..sheet_count {
                        let declared_character_count = names_reader.read_u16()?;
                        let name = BiffUnicodeString::read(
                            &mut names_reader,
                            usize::from(declared_character_count),
                        )?;
                        names.push(SupBookSheetName {
                            declared_character_count,
                            name,
                        });
                    }
                    if names_reader.remaining()? != 0 {
                        return Err(Error::invalid(0, "SupBook sheet names have trailing bytes"));
                    }
                    Ok(names)
                })();
                let (sheet_names, trailing) = match parsed_names {
                    Ok(names) => (names, Vec::new()),
                    Err(_) => (Vec::new(), remaining),
                };
                SupBookLink::VirtualPath {
                    declared_character_count: tag,
                    path,
                    sheet_names,
                    trailing,
                }
            }
            _ => {
                let length = usize::try_from(reader.remaining()?).map_err(|_| {
                    Error::Limit("SupBook compatibility payload exceeds usize".into())
                })?;
                SupBookLink::Compatibility {
                    tag,
                    payload: reader.read_vec(length)?,
                }
            }
        };
        Ok(Self { sheet_count, link })
    }
}

impl SdkWrite for SupBookRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.sheet_count)?;
        match &self.link {
            SupBookLink::SelfReference => writer.write_u16(0x0401)?,
            SupBookLink::AddInFunctions => writer.write_u16(0x3a01)?,
            SupBookLink::VirtualPath {
                declared_character_count,
                path,
                sheet_names,
                trailing,
            } => {
                if path.character_count() != usize::from(*declared_character_count) {
                    return Err(Error::invalid(0, "SupBook virtual path length mismatch"));
                }
                if !sheet_names.is_empty() && sheet_names.len() != usize::from(self.sheet_count) {
                    return Err(Error::invalid(0, "SupBook sheet name count mismatch"));
                }
                writer.write_u16(*declared_character_count)?;
                path.write(writer)?;
                for sheet in sheet_names {
                    if sheet.name.character_count() != usize::from(sheet.declared_character_count) {
                        return Err(Error::invalid(0, "SupBook sheet name length mismatch"));
                    }
                    writer.write_u16(sheet.declared_character_count)?;
                    sheet.name.write(writer)?;
                }
                writer.write_all(trailing)?;
            }
            SupBookLink::Compatibility { tag, payload } => {
                writer.write_u16(*tag)?;
                writer.write_all(payload)?;
            }
        }
        Ok(())
    }
}

impl SdkRead for ExternNameRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let flags = ExternNameFlags::from_bits_retain(reader.read_u16()?);
        let sheet_index = reader.read_u16()?;
        let reserved = reader.read_u16()?;
        let declared_name_character_count = reader.read_u8()?;
        let name = BiffUnicodeString::read(reader, usize::from(declared_name_character_count))?;
        let body_length = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("ExternName body length exceeds usize".into()))?;
        let bytes = reader.read_vec(body_length)?;
        let body = if bytes.is_empty() {
            ExternNameBody::Empty
        } else if flags.intersects(ExternNameFlags::DDE_NO_OPERATION | ExternNameFlags::OLE_LINK) {
            ExternNameBody::Compatibility(bytes)
        } else if flags.contains(ExternNameFlags::WANT_ADVISE) {
            Self::parse_cached_link_values(&bytes).unwrap_or(ExternNameBody::Compatibility(bytes))
        } else {
            Self::parse_formula_body(&bytes).unwrap_or(ExternNameBody::Compatibility(bytes))
        };
        Ok(Self {
            flags,
            sheet_index,
            reserved,
            declared_name_character_count,
            name,
            body,
        })
    }
}

impl ExternNameRecord {
    fn parse_formula_body(bytes: &[u8]) -> Result<ExternNameBody> {
        if bytes.len() < 2 {
            return Err(Error::invalid(0, "ExternName formula body is truncated"));
        }
        let declared_length = u16::from_le_bytes([bytes[0], bytes[1]]);
        let end = 2usize
            .checked_add(usize::from(declared_length))
            .ok_or_else(|| Error::Limit("ExternName formula length overflow".into()))?;
        let encoded = bytes
            .get(2..end)
            .ok_or_else(|| Error::invalid(0, "ExternName formula is truncated"))?;
        if end != bytes.len() {
            return Err(Error::invalid(0, "ExternName formula has trailing bytes"));
        }
        let value = if encoded.is_empty() {
            None
        } else {
            Some(ExternNameFormulaValue::parse(encoded)?)
        };
        Ok(ExternNameBody::ParsedFormula {
            declared_length,
            value,
        })
    }

    fn parse_cached_link_values(bytes: &[u8]) -> Result<ExternNameBody> {
        if bytes.len() < 3 {
            return Err(Error::invalid(0, "ExternName MOper is truncated"));
        }
        let last_column = bytes[0];
        let last_row = u16::from_le_bytes([bytes[1], bytes[2]]);
        let count = (usize::from(last_column) + 1)
            .checked_mul(usize::from(last_row) + 1)
            .ok_or_else(|| Error::Limit("ExternName MOper value count overflow".into()))?;
        if count > MAX_BIFF_RECORD_DATA {
            return Err(Error::Limit(
                "ExternName MOper value count exceeds limit".into(),
            ));
        }
        let mut cursor = 3usize;
        let mut values = Vec::with_capacity(count);
        for _ in 0..count {
            values.push(BiffConstant::read(bytes, &mut cursor)?);
        }
        Ok(ExternNameBody::CachedLinkValues {
            last_column,
            last_row,
            values,
            trailing: bytes[cursor..].to_vec(),
        })
    }
}

impl ExternNameFormulaValue {
    fn parse(bytes: &[u8]) -> Result<Self> {
        let u16_at = |offset| u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
        let u32_at =
            |offset| u32::from_le_bytes(bytes[offset..offset + 4].try_into().expect("four bytes"));
        Ok(match (bytes[0], bytes.len()) {
            (0x3a, 9) => Self::Reference3d {
                sheet_pair: u32_at(1),
                row: u16_at(5),
                column: u16_at(7),
            },
            (0x3b, 13) => Self::Area3d {
                sheet_pair: u32_at(1),
                first_row: u16_at(5),
                last_row: u16_at(7),
                first_column: u16_at(9),
                last_column: u16_at(11),
            },
            (0x3c, 9) => Self::DeletedReference3d {
                sheet_pair: u32_at(1),
                reserved: u32_at(5),
            },
            (0x3d, 13) => Self::DeletedArea3d {
                sheet_pair: u32_at(1),
                reserved1: u32_at(5),
                reserved2: u32_at(9),
            },
            (0x1c, 2) => Self::Error { code: bytes[1] },
            (opcode, length) => {
                return Err(Error::invalid(
                    0,
                    format!("unsupported ExtNameParsedFormula 0x{opcode:02x}/{length} bytes"),
                ));
            }
        })
    }

    fn to_bytes(self) -> Vec<u8> {
        let mut bytes = Vec::new();
        match self {
            Self::Reference3d {
                sheet_pair,
                row,
                column,
            } => {
                bytes.push(0x3a);
                bytes.extend_from_slice(&sheet_pair.to_le_bytes());
                bytes.extend_from_slice(&row.to_le_bytes());
                bytes.extend_from_slice(&column.to_le_bytes());
            }
            Self::Area3d {
                sheet_pair,
                first_row,
                last_row,
                first_column,
                last_column,
            } => {
                bytes.push(0x3b);
                bytes.extend_from_slice(&sheet_pair.to_le_bytes());
                bytes.extend_from_slice(&first_row.to_le_bytes());
                bytes.extend_from_slice(&last_row.to_le_bytes());
                bytes.extend_from_slice(&first_column.to_le_bytes());
                bytes.extend_from_slice(&last_column.to_le_bytes());
            }
            Self::DeletedReference3d {
                sheet_pair,
                reserved,
            } => {
                bytes.push(0x3c);
                bytes.extend_from_slice(&sheet_pair.to_le_bytes());
                bytes.extend_from_slice(&reserved.to_le_bytes());
            }
            Self::DeletedArea3d {
                sheet_pair,
                reserved1,
                reserved2,
            } => {
                bytes.push(0x3d);
                bytes.extend_from_slice(&sheet_pair.to_le_bytes());
                bytes.extend_from_slice(&reserved1.to_le_bytes());
                bytes.extend_from_slice(&reserved2.to_le_bytes());
            }
            Self::Error { code } => bytes.extend_from_slice(&[0x1c, code]),
        }
        bytes
    }
}

impl SdkWrite for ExternNameRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.name.character_count() != usize::from(self.declared_name_character_count) {
            return Err(Error::invalid(0, "ExternName name length mismatch"));
        }
        writer.write_u16(self.flags.bits())?;
        writer.write_u16(self.sheet_index)?;
        writer.write_u16(self.reserved)?;
        writer.write_u8(self.declared_name_character_count)?;
        self.name.write(writer)?;
        match &self.body {
            ExternNameBody::Empty => {}
            ExternNameBody::ParsedFormula {
                declared_length,
                value,
            } => {
                let encoded = value.map_or_else(Vec::new, ExternNameFormulaValue::to_bytes);
                if encoded.len() != usize::from(*declared_length) {
                    return Err(Error::invalid(0, "ExternName formula length mismatch"));
                }
                writer.write_u16(*declared_length)?;
                writer.write_all(&encoded)?;
            }
            ExternNameBody::CachedLinkValues {
                last_column,
                last_row,
                values,
                trailing,
            } => {
                let expected = (usize::from(*last_column) + 1)
                    .checked_mul(usize::from(*last_row) + 1)
                    .ok_or_else(|| Error::Limit("ExternName MOper value count overflow".into()))?;
                if values.len() != expected {
                    return Err(Error::invalid(0, "ExternName MOper value count mismatch"));
                }
                writer.write_u8(*last_column)?;
                writer.write_u16(*last_row)?;
                let mut bytes = Vec::new();
                for value in values {
                    value.write(&mut bytes)?;
                }
                writer.write_all(&bytes)?;
                writer.write_all(trailing)?;
            }
            ExternNameBody::Compatibility(bytes) => writer.write_all(bytes)?,
        }
        Ok(())
    }
}

const URL_MONIKER_CLASS_ID: [u8; 16] = [
    0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];
const FILE_MONIKER_CLASS_ID: [u8; 16] = [
    0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];
const STANDARD_MONIKER_CLASS_ID: [u8; 16] = [
    0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];
const URL_MONIKER_TAIL_GUID: [u8; 16] = [
    0x79, 0x58, 0x81, 0xf4, 0x3b, 0x1d, 0x7f, 0x48, 0xaf, 0x2c, 0x82, 0x5d, 0xc4, 0x85, 0x27, 0x63,
];

impl SdkRead for HyperlinkRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let first_row = reader.read_u16()?;
        let last_row = reader.read_u16()?;
        let first_column = reader.read_u16()?;
        let last_column = reader.read_u16()?;
        let class_id = reader.read_array()?;
        let length = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("HLink object length exceeds usize".into()))?;
        let bytes = reader.read_vec(length)?;
        let object = HyperlinkObject::parse(&bytes).unwrap_or_else(|_| {
            if let Some(value) = HyperlinkObject::parse_truncated_url_moniker(&bytes) {
                return value;
            }
            if bytes.len() >= 8 {
                HyperlinkObject::Truncated {
                    stream_version: u32::from_le_bytes(bytes[0..4].try_into().expect("four bytes")),
                    flags: HyperlinkFlags::from_bits_retain(u32::from_le_bytes(
                        bytes[4..8].try_into().expect("four bytes"),
                    )),
                    payload: bytes[8..].to_vec(),
                }
            } else {
                HyperlinkObject::Compatibility(bytes)
            }
        });
        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
            class_id,
            object,
        })
    }
}

impl SdkWrite for HyperlinkRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.first_row)?;
        writer.write_u16(self.last_row)?;
        writer.write_u16(self.first_column)?;
        writer.write_u16(self.last_column)?;
        writer.write_all(&self.class_id)?;
        writer.write_all(&self.object.to_bytes()?)?;
        Ok(())
    }
}

impl HyperlinkObject {
    pub(crate) fn parse(bytes: &[u8]) -> Result<Self> {
        let mut reader = Reader::new(Cursor::new(bytes))?;
        let stream_version = reader.read_u32()?;
        let flags = HyperlinkFlags::from_bits_retain(reader.read_u32()?);
        let display_name = if flags.contains(HyperlinkFlags::HAS_DISPLAY_NAME) {
            Some(HyperlinkString::read(&mut reader)?)
        } else {
            None
        };
        let target_frame_name = if flags.contains(HyperlinkFlags::HAS_TARGET_FRAME) {
            Some(HyperlinkString::read(&mut reader)?)
        } else {
            None
        };
        let moniker = if flags.contains(HyperlinkFlags::HAS_MONIKER) {
            Some(Box::new(
                if flags.contains(HyperlinkFlags::MONIKER_SAVED_AS_STRING) {
                    HyperlinkMoniker::String(HyperlinkString::read(&mut reader)?)
                } else {
                    HyperlinkMoniker::read(&mut reader)?
                },
            ))
        } else {
            None
        };
        let location = if flags.contains(HyperlinkFlags::HAS_LOCATION) {
            Some(HyperlinkString::read(&mut reader)?)
        } else {
            None
        };
        let guid = if flags.contains(HyperlinkFlags::HAS_GUID) {
            Some(reader.read_array()?)
        } else {
            None
        };
        let creation_time = if flags.contains(HyperlinkFlags::HAS_CREATION_TIME) {
            Some(reader.read_u64()?)
        } else {
            None
        };
        let trailing_length = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("Hyperlink trailing length exceeds usize".into()))?;
        let trailing = reader.read_vec(trailing_length)?;
        Ok(Self::Parsed {
            stream_version,
            flags,
            display_name,
            target_frame_name,
            moniker,
            location,
            guid,
            creation_time,
            trailing,
        })
    }

    pub(crate) fn to_bytes(&self) -> Result<Vec<u8>> {
        match self {
            Self::Compatibility(bytes) => Ok(bytes.clone()),
            Self::Truncated {
                stream_version,
                flags,
                payload,
            } => {
                let mut bytes = stream_version.to_le_bytes().to_vec();
                bytes.extend_from_slice(&flags.bits().to_le_bytes());
                bytes.extend_from_slice(payload);
                Ok(bytes)
            }
            Self::TruncatedUrlMoniker {
                stream_version,
                flags,
                class_id,
                declared_byte_length,
                address,
            } => {
                if *class_id != URL_MONIKER_CLASS_ID
                    || !flags.contains(HyperlinkFlags::HAS_MONIKER)
                    || flags.contains(HyperlinkFlags::MONIKER_SAVED_AS_STRING)
                    || address.len().checked_mul(2).is_none_or(|available| {
                        available >= usize::try_from(*declared_byte_length).unwrap_or(usize::MAX)
                    })
                {
                    return Err(Error::invalid(
                        0,
                        "truncated URL moniker invariants changed",
                    ));
                }
                let mut bytes = stream_version.to_le_bytes().to_vec();
                bytes.extend_from_slice(&flags.bits().to_le_bytes());
                bytes.extend_from_slice(class_id);
                bytes.extend_from_slice(&declared_byte_length.to_le_bytes());
                for unit in address {
                    bytes.extend_from_slice(&unit.to_le_bytes());
                }
                Ok(bytes)
            }
            Self::Parsed {
                stream_version,
                flags,
                display_name,
                target_frame_name,
                moniker,
                location,
                guid,
                creation_time,
                trailing,
            } => {
                if flags.contains(HyperlinkFlags::HAS_DISPLAY_NAME) != display_name.is_some()
                    || flags.contains(HyperlinkFlags::HAS_TARGET_FRAME)
                        != target_frame_name.is_some()
                    || flags.contains(HyperlinkFlags::HAS_MONIKER) != moniker.is_some()
                    || flags.contains(HyperlinkFlags::HAS_LOCATION) != location.is_some()
                    || flags.contains(HyperlinkFlags::HAS_GUID) != guid.is_some()
                    || flags.contains(HyperlinkFlags::HAS_CREATION_TIME) != creation_time.is_some()
                {
                    return Err(Error::invalid(
                        0,
                        "Hyperlink flags and optional fields disagree",
                    ));
                }
                let mut bytes = stream_version.to_le_bytes().to_vec();
                bytes.extend_from_slice(&flags.bits().to_le_bytes());
                if let Some(value) = display_name {
                    bytes.extend_from_slice(&value.to_bytes()?);
                }
                if let Some(value) = target_frame_name {
                    bytes.extend_from_slice(&value.to_bytes()?);
                }
                if let Some(value) = moniker {
                    bytes.extend_from_slice(&value.to_bytes()?);
                }
                if let Some(value) = location {
                    bytes.extend_from_slice(&value.to_bytes()?);
                }
                if let Some(value) = guid {
                    bytes.extend_from_slice(value);
                }
                if let Some(value) = creation_time {
                    bytes.extend_from_slice(&value.to_le_bytes());
                }
                bytes.extend_from_slice(trailing);
                Ok(bytes)
            }
        }
    }

    fn parse_truncated_url_moniker(bytes: &[u8]) -> Option<Self> {
        let fixed = bytes.get(..28)?;
        let stream_version = u32::from_le_bytes(fixed[0..4].try_into().ok()?);
        let flags =
            HyperlinkFlags::from_bits_retain(u32::from_le_bytes(fixed[4..8].try_into().ok()?));
        if !flags.contains(HyperlinkFlags::HAS_MONIKER)
            || flags.intersects(
                HyperlinkFlags::HAS_DISPLAY_NAME
                    | HyperlinkFlags::HAS_TARGET_FRAME
                    | HyperlinkFlags::MONIKER_SAVED_AS_STRING,
            )
        {
            return None;
        }
        let class_id: [u8; 16] = fixed[8..24].try_into().ok()?;
        if class_id != URL_MONIKER_CLASS_ID {
            return None;
        }
        let declared_byte_length = u32::from_le_bytes(fixed[24..28].try_into().ok()?);
        let available = &bytes[28..];
        if !available.len().is_multiple_of(2)
            || available.len() >= usize::try_from(declared_byte_length).ok()?
        {
            return None;
        }
        Some(Self::TruncatedUrlMoniker {
            stream_version,
            flags,
            class_id,
            declared_byte_length,
            address: available
                .chunks_exact(2)
                .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
                .collect(),
        })
    }
}

impl HyperlinkString {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let declared_character_count = reader.read_u32()?;
        let count = usize::try_from(declared_character_count)
            .map_err(|_| Error::Limit("HyperlinkString count exceeds usize".into()))?;
        reader.ensure_allocation(count, 2)?;
        let mut characters = Vec::with_capacity(count);
        for _ in 0..count {
            characters.push(reader.read_u16()?);
        }
        Ok(Self {
            declared_character_count,
            characters,
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.characters.len()
            != usize::try_from(self.declared_character_count)
                .map_err(|_| Error::Limit("HyperlinkString count exceeds usize".into()))?
        {
            return Err(Error::invalid(0, "HyperlinkString length mismatch"));
        }
        let mut bytes = self.declared_character_count.to_le_bytes().to_vec();
        for character in &self.characters {
            bytes.extend_from_slice(&character.to_le_bytes());
        }
        Ok(bytes)
    }
}

impl HyperlinkMoniker {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let class_id = reader.read_array()?;
        if class_id == URL_MONIKER_CLASS_ID {
            let declared_byte_length = reader.read_u32()?;
            let length = usize::try_from(declared_byte_length)
                .map_err(|_| Error::Limit("URLMoniker length exceeds usize".into()))?;
            let data = reader.read_vec(length)?;
            let tail_length = if data.len() >= 24
                && data[data.len() - 24..data.len() - 8] == URL_MONIKER_TAIL_GUID
            {
                24
            } else {
                0
            };
            let address_bytes = &data[..data.len() - tail_length];
            if !address_bytes.len().is_multiple_of(2) {
                return Err(Error::invalid(
                    0,
                    "URLMoniker address is not UTF-16 aligned",
                ));
            }
            Ok(Self::Url {
                class_id,
                declared_byte_length,
                address: address_bytes
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect(),
                tail: data[data.len() - tail_length..].to_vec(),
            })
        } else if class_id == FILE_MONIKER_CLASS_ID {
            let options = reader.read_u16()?;
            let declared_short_name_length = reader.read_u32()?;
            let short_name =
                reader
                    .read_vec(usize::try_from(declared_short_name_length).map_err(|_| {
                        Error::Limit("FileMoniker short name exceeds usize".into())
                    })?)?;
            let tail = reader.read_array()?;
            let declared_total_length = reader.read_u32()?;
            let long_path = if declared_total_length == 0 {
                None
            } else {
                let declared_character_bytes = reader.read_u32()?;
                let key = reader.read_u16()?;
                let byte_count = usize::try_from(declared_character_bytes).map_err(|_| {
                    Error::Limit("FileMoniker long path length exceeds usize".into())
                })?;
                if !byte_count.is_multiple_of(2) {
                    return Err(Error::invalid(0, "FileMoniker long path is misaligned"));
                }
                let data = reader.read_vec(byte_count)?;
                Some(HyperlinkLongPath {
                    declared_total_length,
                    declared_character_bytes,
                    key,
                    characters: data
                        .chunks_exact(2)
                        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                        .collect(),
                })
            };
            Ok(Self::File {
                class_id,
                options,
                declared_short_name_length,
                short_name,
                tail,
                long_path,
            })
        } else if class_id == STANDARD_MONIKER_CLASS_ID {
            let options = reader.read_u16()?;
            let declared_data_length = reader.read_u32()?;
            let data = reader.read_vec(
                usize::try_from(declared_data_length)
                    .map_err(|_| Error::Limit("standard moniker length exceeds usize".into()))?,
            )?;
            Ok(Self::Standard {
                class_id,
                options,
                declared_data_length,
                data,
            })
        } else {
            Err(Error::invalid(0, "unsupported HyperlinkMoniker CLSID"))
        }
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        match self {
            Self::String(value) => return value.to_bytes(),
            Self::Url {
                class_id,
                declared_byte_length,
                address,
                tail,
            } => {
                let length = address
                    .len()
                    .checked_mul(2)
                    .and_then(|length| length.checked_add(tail.len()))
                    .ok_or_else(|| Error::Limit("URLMoniker length overflow".into()))?;
                if length != usize::try_from(*declared_byte_length).unwrap_or(usize::MAX) {
                    return Err(Error::invalid(0, "URLMoniker length mismatch"));
                }
                bytes.extend_from_slice(class_id);
                bytes.extend_from_slice(&declared_byte_length.to_le_bytes());
                for character in address {
                    bytes.extend_from_slice(&character.to_le_bytes());
                }
                bytes.extend_from_slice(tail);
            }
            Self::File {
                class_id,
                options,
                declared_short_name_length,
                short_name,
                tail,
                long_path,
            } => {
                if short_name.len()
                    != usize::try_from(*declared_short_name_length).unwrap_or(usize::MAX)
                {
                    return Err(Error::invalid(0, "FileMoniker short-name length mismatch"));
                }
                bytes.extend_from_slice(class_id);
                bytes.extend_from_slice(&options.to_le_bytes());
                bytes.extend_from_slice(&declared_short_name_length.to_le_bytes());
                bytes.extend_from_slice(short_name);
                bytes.extend_from_slice(tail);
                if let Some(path) = long_path {
                    if path.characters.len().checked_mul(2)
                        != Some(
                            usize::try_from(path.declared_character_bytes).unwrap_or(usize::MAX),
                        )
                    {
                        return Err(Error::invalid(0, "FileMoniker long-path length mismatch"));
                    }
                    bytes.extend_from_slice(&path.declared_total_length.to_le_bytes());
                    bytes.extend_from_slice(&path.declared_character_bytes.to_le_bytes());
                    bytes.extend_from_slice(&path.key.to_le_bytes());
                    for character in &path.characters {
                        bytes.extend_from_slice(&character.to_le_bytes());
                    }
                } else {
                    bytes.extend_from_slice(&0u32.to_le_bytes());
                }
            }
            Self::Standard {
                class_id,
                options,
                declared_data_length,
                data,
            } => {
                if data.len() != usize::try_from(*declared_data_length).unwrap_or(usize::MAX) {
                    return Err(Error::invalid(0, "standard moniker length mismatch"));
                }
                bytes.extend_from_slice(class_id);
                bytes.extend_from_slice(&options.to_le_bytes());
                bytes.extend_from_slice(&declared_data_length.to_le_bytes());
                bytes.extend_from_slice(data);
            }
        }
        Ok(bytes)
    }
}

impl DataValidationOptions {
    fn from_bits(bits: u32) -> Self {
        Self {
            validation_type: (bits & 0x0f) as u8,
            error_style: ((bits >> 4) & 0x07) as u8,
            string_lookup: bits & (1 << 7) != 0,
            allow_blank: bits & (1 << 8) != 0,
            suppress_combo: bits & (1 << 9) != 0,
            ime_mode: ((bits >> 10) & 0xff) as u8,
            show_input_message: bits & (1 << 18) != 0,
            show_error_message: bits & (1 << 19) != 0,
            operator: ((bits >> 20) & 0x0f) as u8,
            reserved: (bits >> 24) as u8,
        }
    }

    fn bits(self) -> u32 {
        u32::from(self.validation_type & 0x0f)
            | (u32::from(self.error_style & 0x07) << 4)
            | (u32::from(self.string_lookup) << 7)
            | (u32::from(self.allow_blank) << 8)
            | (u32::from(self.suppress_combo) << 9)
            | (u32::from(self.ime_mode) << 10)
            | (u32::from(self.show_input_message) << 18)
            | (u32::from(self.show_error_message) << 19)
            | (u32::from(self.operator & 0x0f) << 20)
            | (u32::from(self.reserved) << 24)
    }
}

impl SdkRead for XlUnicodeString {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let character_count = usize::from(reader.read_u16()?);
        Ok(Self {
            text: BiffUnicodeString::read(reader, character_count)?,
        })
    }
}

impl SdkWrite for XlUnicodeString {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(
            u16::try_from(self.text.character_count())
                .map_err(|_| Error::Limit("XLUnicodeString exceeds u16 characters".into()))?,
        )?;
        self.text.write(writer)
    }
}

impl SdkSize for XlUnicodeString {
    fn sdk_size(&self) -> u64 {
        let character_bytes = match &self.text.characters {
            XlStringCharacters::Compressed(values) => values.len() as u64,
            XlStringCharacters::Unicode(values) => (values.len() as u64).saturating_mul(2),
        };
        3u64.saturating_add(character_bytes)
            .saturating_add(u64::from(self.text.trailing_byte.is_some()))
    }
}

impl SdkRead for DataValidationFormula {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let byte_count = usize::from(reader.read_u16()?);
        let unused = reader.read_u16()?;
        let tokens = FormulaTokenStream::from_bytes(&reader.read_vec(byte_count)?)?;
        Ok(Self { unused, tokens })
    }
}

impl SdkWrite for DataValidationFormula {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let bytes = self.tokens.to_bytes()?;
        writer.write_u16(
            u16::try_from(bytes.len())
                .map_err(|_| Error::Limit("DV formula exceeds u16 bytes".into()))?,
        )?;
        writer.write_u16(self.unused)?;
        writer.write_all(&bytes)?;
        Ok(())
    }
}

impl SdkRead for DataValidationRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let options = DataValidationOptions::from_bits(reader.read_u32()?);
        let prompt_title = XlUnicodeString::read_from(reader)?;
        let error_title = XlUnicodeString::read_from(reader)?;
        let prompt = XlUnicodeString::read_from(reader)?;
        let error = XlUnicodeString::read_from(reader)?;
        let formula1 = DataValidationFormula::read_from(reader)?;
        let formula2 = DataValidationFormula::read_from(reader)?;
        let range_count = usize::from(reader.read_u16()?);
        let mut ranges = Vec::with_capacity(range_count);
        for _ in 0..range_count {
            ranges.push(CellRange::read_from(reader)?);
        }
        Ok(Self {
            options,
            prompt_title,
            error_title,
            prompt,
            error,
            formula1,
            formula2,
            ranges,
        })
    }
}

impl SdkWrite for DataValidationRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u32(self.options.bits())?;
        self.prompt_title.write_to(writer)?;
        self.error_title.write_to(writer)?;
        self.prompt.write_to(writer)?;
        self.error.write_to(writer)?;
        self.formula1.write_to(writer)?;
        self.formula2.write_to(writer)?;
        writer.write_u16(
            u16::try_from(self.ranges.len())
                .map_err(|_| Error::Limit("DV range count exceeds u16".into()))?,
        )?;
        for range in &self.ranges {
            range.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for DxfN {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let raw_flags = u64::from(reader.read_u32()?) | (u64::from(reader.read_u16()?) << 32);
        let flags = DxfFlags::from_bits_retain(raw_flags);
        let number_format = if flags.contains(DxfFlags::NUMBER_FORMAT) {
            Some(if flags.contains(DxfFlags::USER_NUMBER_FORMAT) {
                let declared_size = reader.read_u16()?;
                let payload_size = declared_size.checked_sub(2).ok_or_else(|| {
                    Error::invalid(reader.position().unwrap_or(0), "DXFNumUsr size is below 2")
                })?;
                let mut child = reader.sub_reader(u64::from(payload_size))?;
                let format = XlUnicodeString::read_from(&mut child)?;
                if child.remaining()? != 0 {
                    return Err(Error::invalid(
                        child.position()?,
                        "DXFNumUsr has trailing bytes",
                    ));
                }
                DxfNumberFormat::UserDefined {
                    declared_size,
                    format,
                }
            } else {
                DxfNumberFormat::BuiltIn {
                    unused: reader.read_u8()?,
                    format_index: reader.read_u8()?,
                }
            })
        } else {
            None
        };
        let font = flags
            .contains(DxfFlags::FONT)
            .then(|| DxfFont::read_from(reader))
            .transpose()?;
        let alignment = flags
            .contains(DxfFlags::ALIGNMENT)
            .then(|| DxfAlignment::read_from(reader))
            .transpose()?;
        let border = flags
            .contains(DxfFlags::BORDER)
            .then(|| DxfBorder::read_from(reader))
            .transpose()?;
        let pattern = flags
            .contains(DxfFlags::PATTERN)
            .then(|| DxfPattern::read_from(reader))
            .transpose()?;
        let protection = flags
            .contains(DxfFlags::PROTECTION)
            .then(|| DxfProtection::read_from(reader))
            .transpose()?;
        Ok(Self {
            flags,
            number_format,
            font,
            alignment,
            border,
            pattern,
            protection,
        })
    }
}

impl SdkWrite for DxfN {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        for (flag, present, name) in [
            (
                DxfFlags::NUMBER_FORMAT,
                self.number_format.is_some(),
                "number format",
            ),
            (DxfFlags::FONT, self.font.is_some(), "font"),
            (DxfFlags::ALIGNMENT, self.alignment.is_some(), "alignment"),
            (DxfFlags::BORDER, self.border.is_some(), "border"),
            (DxfFlags::PATTERN, self.pattern.is_some(), "pattern"),
            (
                DxfFlags::PROTECTION,
                self.protection.is_some(),
                "protection",
            ),
        ] {
            if self.flags.contains(flag) != present {
                return Err(Error::invalid(
                    0,
                    format!("DXFN {name} flag disagrees with field"),
                ));
            }
        }
        let bits = self.flags.bits();
        writer.write_u32(bits as u32)?;
        writer.write_u16((bits >> 32) as u16)?;
        if let Some(number_format) = &self.number_format {
            match number_format {
                DxfNumberFormat::BuiltIn {
                    unused,
                    format_index,
                } => {
                    if self.flags.contains(DxfFlags::USER_NUMBER_FORMAT) {
                        return Err(Error::invalid(
                            0,
                            "DXFN built-in number format has user flag",
                        ));
                    }
                    writer.write_u8(*unused)?;
                    writer.write_u8(*format_index)?;
                }
                DxfNumberFormat::UserDefined {
                    declared_size,
                    format,
                } => {
                    if !self.flags.contains(DxfFlags::USER_NUMBER_FORMAT) {
                        return Err(Error::invalid(0, "DXFN user number format lacks user flag"));
                    }
                    writer.write_u16(*declared_size)?;
                    format.write_to(writer)?;
                }
            }
        }
        if let Some(value) = &self.font {
            value.write_to(writer)?;
        }
        if let Some(value) = &self.alignment {
            value.write_to(writer)?;
        }
        if let Some(value) = &self.border {
            value.write_to(writer)?;
        }
        if let Some(value) = &self.pattern {
            value.write_to(writer)?;
        }
        if let Some(value) = &self.protection {
            value.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for ConditionalFormattingRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let condition_type = reader.read_u8()?;
        let comparison_operator = reader.read_u8()?;
        let formula1_length = usize::from(reader.read_u16()?);
        let formula2_length = usize::from(reader.read_u16()?);
        let format = DxfN::read_from(reader)?;
        let formula1 = FormulaTokenStream::from_bytes(&reader.read_vec(formula1_length)?)?;
        let formula2 = FormulaTokenStream::from_bytes(&reader.read_vec(formula2_length)?)?;
        Ok(Self {
            condition_type,
            comparison_operator,
            format,
            formula1,
            formula2,
        })
    }
}

impl SdkWrite for ConditionalFormattingRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let formula1 = self.formula1.to_bytes()?;
        let formula2 = self.formula2.to_bytes()?;
        writer.write_u8(self.condition_type)?;
        writer.write_u8(self.comparison_operator)?;
        writer.write_u16(
            u16::try_from(formula1.len())
                .map_err(|_| Error::Limit("CF formula1 exceeds u16 bytes".into()))?,
        )?;
        writer.write_u16(
            u16::try_from(formula2.len())
                .map_err(|_| Error::Limit("CF formula2 exceeds u16 bytes".into()))?,
        )?;
        self.format.write_to(writer)?;
        writer.write_all(&formula1)?;
        writer.write_all(&formula2)?;
        Ok(())
    }
}

impl SdkRead for ConditionalFormattingGroupRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let rule_count = reader.read_u16()?;
        let flags_and_id = reader.read_u16()?;
        let bounds = CellRange::read_from(reader)?;
        let range_count = usize::from(reader.read_u16()?);
        let maximum_count = usize::try_from(reader.remaining()? / 8)
            .map_err(|_| Error::Limit("CondFmt range length exceeds usize".into()))?;
        if range_count > maximum_count {
            return Err(Error::invalid(
                reader.position()?,
                "CondFmt range count exceeds remaining bytes",
            ));
        }
        let mut ranges = Vec::with_capacity(range_count);
        for _ in 0..range_count {
            ranges.push(CellRange::read_from(reader)?);
        }
        Ok(Self {
            rule_count,
            flags_and_id,
            bounds,
            ranges,
        })
    }
}

impl SdkWrite for ConditionalFormattingGroupRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.rule_count)?;
        writer.write_u16(self.flags_and_id)?;
        self.bounds.write_to(writer)?;
        writer.write_u16(
            u16::try_from(self.ranges.len())
                .map_err(|_| Error::Limit("CondFmt range count exceeds u16".into()))?,
        )?;
        for range in &self.ranges {
            range.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for XfExtNoFrt {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let reserved1 = reader.read_u16()?;
        let reserved2 = reader.read_u16()?;
        let reserved3 = reader.read_u16()?;
        let property_count = usize::from(reader.read_u16()?);
        let mut properties = Vec::with_capacity(property_count);
        for _ in 0..property_count {
            properties.push(ExtProperty::read(reader)?);
        }
        Ok(Self {
            reserved1,
            reserved2,
            reserved3,
            properties,
        })
    }
}

impl SdkWrite for XfExtNoFrt {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.reserved1)?;
        writer.write_u16(self.reserved2)?;
        writer.write_u16(self.reserved3)?;
        writer.write_u16(
            u16::try_from(self.properties.len())
                .map_err(|_| Error::Limit("XFExtNoFRT property count exceeds u16".into()))?,
        )?;
        for property in &self.properties {
            property.write(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for DxfN12 {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let byte_count = reader.read_u32()?;
        if byte_count == 0 {
            return Ok(Self::Empty {
                reserved: reader.read_u16()?,
            });
        }
        let mut child = reader.sub_reader(u64::from(byte_count))?;
        let format = DxfN::read_from(&mut child)?;
        let extension = if child.remaining()? == 0 {
            None
        } else {
            Some(XfExtNoFrt::read_from(&mut child)?)
        };
        if child.remaining()? != 0 {
            return Err(Error::invalid(
                child.position()?,
                "DXFN12 has trailing bytes",
            ));
        }
        Ok(Self::Formatting {
            format: Box::new(format),
            extension,
        })
    }
}

impl SdkWrite for DxfN12 {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::Empty { reserved } => {
                writer.write_u32(0)?;
                writer.write_u16(*reserved)?;
            }
            Self::Formatting { format, extension } => {
                let mut child = Writer::new(Cursor::new(Vec::new()));
                format.write_to(&mut child)?;
                if let Some(value) = extension {
                    value.write_to(&mut child)?;
                }
                let bytes = child.into_inner().into_inner();
                writer.write_u32(
                    u32::try_from(bytes.len())
                        .map_err(|_| Error::Limit("DXFN12 exceeds u32 bytes".into()))?,
                )?;
                writer.write_all(&bytes)?;
            }
        }
        Ok(())
    }
}

impl CfExTemplateParams {
    fn read<R: Read + Seek>(reader: &mut Reader<R>, template_id: u8) -> Result<Self> {
        Ok(match template_id {
            0x05 => Self::Filter {
                flags: reader.read_u8()?,
                parameter: reader.read_u16()?,
                reserved: reader.read_array()?,
            },
            0x08 => Self::Text {
                comparison_type: reader.read_u16()?,
                reserved: reader.read_array()?,
            },
            0x0f..=0x18 => Self::Date {
                operation: reader.read_u16()?,
                reserved: reader.read_array()?,
            },
            0x19 | 0x1a | 0x1d | 0x1e => Self::Averages {
                parameter: reader.read_u16()?,
                reserved: reader.read_array()?,
            },
            _ => Self::Default {
                unused: reader.read_array()?,
            },
        })
    }

    fn write<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::Filter {
                flags,
                parameter,
                reserved,
            } => {
                writer.write_u8(*flags)?;
                writer.write_u16(*parameter)?;
                writer.write_all(reserved)?;
            }
            Self::Text {
                comparison_type,
                reserved,
            } => {
                writer.write_u16(*comparison_type)?;
                writer.write_all(reserved)?;
            }
            Self::Date {
                operation,
                reserved,
            } => {
                writer.write_u16(*operation)?;
                writer.write_all(reserved)?;
            }
            Self::Averages {
                parameter,
                reserved,
            } => {
                writer.write_u16(*parameter)?;
                writer.write_all(reserved)?;
            }
            Self::Default { unused } => writer.write_all(unused)?,
        }
        Ok(())
    }
}

impl SdkRead for CfExNonCf12 {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let rule_index = reader.read_u16()?;
        let comparison_operator = reader.read_u8()?;
        let template_id = reader.read_u8()?;
        let priority = reader.read_u16()?;
        let flags = reader.read_u8()?;
        let has_format = reader.read_u8()?;
        let format = if has_format == 0 {
            None
        } else {
            Some(DxfN12::read_from(reader)?)
        };
        let declared_template_parameter_size = reader.read_u8()?;
        let template_parameters = CfExTemplateParams::read(reader, template_id)?;
        Ok(Self {
            rule_index,
            comparison_operator,
            template_id,
            priority,
            flags,
            format,
            declared_template_parameter_size,
            template_parameters,
        })
    }
}

impl SdkWrite for CfExNonCf12 {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.rule_index)?;
        writer.write_u8(self.comparison_operator)?;
        writer.write_u8(self.template_id)?;
        writer.write_u16(self.priority)?;
        writer.write_u8(self.flags)?;
        writer.write_u8(u8::from(self.format.is_some()))?;
        if let Some(value) = &self.format {
            value.write_to(writer)?;
        }
        writer.write_u8(self.declared_template_parameter_size)?;
        self.template_parameters.write(writer)
    }
}

impl SdkRead for ConditionalFormattingExtensionRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let is_cf12 = reader.read_u32()?;
        let group_id = reader.read_u16()?;
        let content = if is_cf12 == 0 {
            Some(CfExNonCf12::read_from(reader)?)
        } else {
            None
        };
        Ok(Self {
            header,
            is_cf12,
            group_id,
            content,
        })
    }
}

impl SdkWrite for ConditionalFormattingExtensionRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if (self.is_cf12 == 0) != self.content.is_some() {
            return Err(Error::invalid(0, "CFEx fIsCF12 and content disagree"));
        }
        self.header.write_to(writer)?;
        writer.write_u32(self.is_cf12)?;
        writer.write_u16(self.group_id)?;
        if let Some(value) = &self.content {
            value.write_to(writer)?;
        }
        Ok(())
    }
}

impl CfVo {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let value_type = reader.read_u8()?;
        let formula_length = usize::from(reader.read_u16()?);
        let formula = FormulaTokenStream::from_bytes(&reader.read_vec(formula_length)?)?;
        let value_bits = (formula_length == 0 && !matches!(value_type, 0x02 | 0x03))
            .then(|| reader.read_u64())
            .transpose()?;
        Ok(Self {
            value_type,
            formula,
            value_bits,
        })
    }

    fn write<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let formula = self.formula.to_bytes()?;
        let expects_value = formula.is_empty() && !matches!(self.value_type, 0x02 | 0x03);
        if expects_value != self.value_bits.is_some() {
            return Err(Error::invalid(
                writer.position()?,
                "CFVO value presence disagrees with its type and formula",
            ));
        }
        writer.write_u8(self.value_type)?;
        writer.write_u16(
            u16::try_from(formula.len())
                .map_err(|_| Error::Limit("CFVO formula exceeds u16 bytes".into()))?,
        )?;
        writer.write_all(&formula)?;
        if let Some(value) = self.value_bits {
            writer.write_u64(value)?;
        }
        Ok(())
    }
}

impl SdkRead for CfGradient {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let unused = reader.read_u16()?;
        let reserved = reader.read_u8()?;
        let interpolation_count = reader.read_u8()?;
        let gradient_count = reader.read_u8()?;
        let flags = CfGradientFlags::from_bits_retain(reader.read_u8()?);
        reader.ensure_allocation(usize::from(interpolation_count), 12)?;
        let mut interpolation = Vec::with_capacity(usize::from(interpolation_count));
        for _ in 0..interpolation_count {
            interpolation.push(CfGradientInterpolationItem {
                value: CfVo::read(reader)?,
                domain_bits: reader.read_u64()?,
            });
        }
        reader.ensure_allocation(usize::from(gradient_count), 24)?;
        let mut gradient = Vec::with_capacity(usize::from(gradient_count));
        for _ in 0..gradient_count {
            gradient.push(CfGradientItem::read_from(reader)?);
        }
        Ok(Self {
            unused,
            reserved,
            interpolation_count,
            gradient_count,
            flags,
            interpolation,
            gradient,
        })
    }
}

impl SdkWrite for CfGradient {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if usize::from(self.interpolation_count) != self.interpolation.len()
            || usize::from(self.gradient_count) != self.gradient.len()
        {
            return Err(Error::invalid(
                writer.position()?,
                "CFGradient counts do not match their arrays",
            ));
        }
        writer.write_u16(self.unused)?;
        writer.write_u8(self.reserved)?;
        writer.write_u8(self.interpolation_count)?;
        writer.write_u8(self.gradient_count)?;
        writer.write_u8(self.flags.bits())?;
        for item in &self.interpolation {
            item.value.write(writer)?;
            writer.write_u64(item.domain_bits)?;
        }
        for item in &self.gradient {
            item.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for CfDataBar {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        Ok(Self {
            unused: reader.read_u16()?,
            reserved: reader.read_u8()?,
            flags: CfDataBarFlags::from_bits_retain(reader.read_u8()?),
            minimum_percent: reader.read_u8()?,
            maximum_percent: reader.read_u8()?,
            color: CfColor::read_from(reader)?,
            minimum: CfVo::read(reader)?,
            maximum: CfVo::read(reader)?,
        })
    }
}

impl SdkWrite for CfDataBar {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.unused)?;
        writer.write_u8(self.reserved)?;
        writer.write_u8(self.flags.bits())?;
        writer.write_u8(self.minimum_percent)?;
        writer.write_u8(self.maximum_percent)?;
        self.color.write_to(writer)?;
        self.minimum.write(writer)?;
        self.maximum.write(writer)
    }
}

impl SdkRead for CfFilter {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let declared_size = reader.read_u16()?;
        let mut child = reader.sub_reader(u64::from(declared_size))?;
        let value = Self {
            declared_size,
            reserved: child.read_u8()?,
            flags: CfFilterFlags::from_bits_retain(child.read_u8()?),
            parameter: child.read_u16()?,
        };
        if child.remaining()? != 0 {
            return Err(Error::invalid(
                child.position()?,
                "CFFilter has trailing bytes",
            ));
        }
        Ok(value)
    }
}

impl SdkWrite for CfFilter {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.declared_size != 4 {
            return Err(Error::invalid(
                writer.position()?,
                "CFFilter cbFilter is not 4",
            ));
        }
        writer.write_u16(self.declared_size)?;
        writer.write_u8(self.reserved)?;
        writer.write_u8(self.flags.bits())?;
        writer.write_u16(self.parameter)?;
        Ok(())
    }
}

impl SdkRead for CfMultistate {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let unused = reader.read_u16()?;
        let reserved = reader.read_u8()?;
        let state_count = reader.read_u8()?;
        let icon_set = reader.read_u8()?;
        let flags = CfMultistateFlags::from_bits_retain(reader.read_u8()?);
        reader.ensure_allocation(usize::from(state_count), 8)?;
        let mut states = Vec::with_capacity(usize::from(state_count));
        for _ in 0..state_count {
            states.push(CfMultistateItem {
                value: CfVo::read(reader)?,
                equal: reader.read_u8()?,
                unused: reader.read_u32()?,
            });
        }
        Ok(Self {
            unused,
            reserved,
            state_count,
            icon_set,
            flags,
            states,
        })
    }
}

impl SdkWrite for CfMultistate {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if usize::from(self.state_count) != self.states.len() {
            return Err(Error::invalid(
                writer.position()?,
                "CFMultistate cStates does not match rgStates",
            ));
        }
        writer.write_u16(self.unused)?;
        writer.write_u8(self.reserved)?;
        writer.write_u8(self.state_count)?;
        writer.write_u8(self.icon_set)?;
        writer.write_u8(self.flags.bits())?;
        for state in &self.states {
            state.value.write(writer)?;
            writer.write_u8(state.equal)?;
            writer.write_u32(state.unused)?;
        }
        Ok(())
    }
}

impl SdkRead for ConditionalFormatting12Record {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtRefHeaderU::read_from(reader)?;
        let condition_type = reader.read_u8()?;
        let comparison_operator = reader.read_u8()?;
        let formula1_length = usize::from(reader.read_u16()?);
        let formula2_length = usize::from(reader.read_u16()?);
        let format = DxfN12::read_from(reader)?;
        let formula1 = FormulaTokenStream::from_bytes(&reader.read_vec(formula1_length)?)?;
        let formula2 = FormulaTokenStream::from_bytes(&reader.read_vec(formula2_length)?)?;
        let active_formula_length = usize::from(reader.read_u16()?);
        let active_formula =
            FormulaTokenStream::from_bytes(&reader.read_vec(active_formula_length)?)?;
        let option_flags = reader.read_u8()?;
        let priority = reader.read_u16()?;
        let template_id = reader.read_u16()?;
        let template_parameter_size = reader.read_u8()?;
        let template_parameters = if template_parameter_size == 0 {
            None
        } else {
            let mut child = reader.sub_reader(u64::from(template_parameter_size))?;
            let value = CfExTemplateParams::read(&mut child, template_id as u8)?;
            if child.remaining()? != 0 {
                return Err(Error::invalid(
                    child.position()?,
                    "CF12 template parameters have trailing bytes",
                ));
            }
            Some(value)
        };
        let condition_data = match condition_type {
            0x01 | 0x02 => Cf12ConditionData::None,
            0x03 => Cf12ConditionData::Gradient(CfGradient::read_from(reader)?),
            0x04 => Cf12ConditionData::DataBar(CfDataBar::read_from(reader)?),
            0x05 => Cf12ConditionData::Filter(CfFilter::read_from(reader)?),
            0x06 => Cf12ConditionData::Multistate(CfMultistate::read_from(reader)?),
            value => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("unsupported CF12 condition type 0x{value:02x}"),
                ));
            }
        };
        Ok(Self {
            header,
            condition_type,
            comparison_operator,
            formula1,
            formula2,
            format,
            active_formula,
            option_flags,
            priority,
            template_id,
            template_parameter_size,
            template_parameters,
            condition_data,
        })
    }
}

impl SdkWrite for ConditionalFormatting12Record {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let formula1 = self.formula1.to_bytes()?;
        let formula2 = self.formula2.to_bytes()?;
        let active_formula = self.active_formula.to_bytes()?;
        let expected_condition_type = match &self.condition_data {
            Cf12ConditionData::None => self.condition_type,
            Cf12ConditionData::Gradient(_) => 0x03,
            Cf12ConditionData::DataBar(_) => 0x04,
            Cf12ConditionData::Filter(_) => 0x05,
            Cf12ConditionData::Multistate(_) => 0x06,
        };
        if expected_condition_type != self.condition_type
            || (matches!(&self.condition_data, Cf12ConditionData::None)
                && !matches!(self.condition_type, 0x01 | 0x02))
        {
            return Err(Error::invalid(
                writer.position()?,
                "CF12 condition type disagrees with rgbCT",
            ));
        }
        let mut template_bytes = Vec::new();
        if let Some(parameters) = &self.template_parameters {
            let mut child = Writer::new(Cursor::new(Vec::new()));
            parameters.write(&mut child)?;
            template_bytes = child.into_inner().into_inner();
        }
        if template_bytes.len() != usize::from(self.template_parameter_size) {
            return Err(Error::invalid(
                writer.position()?,
                "CF12 cbTemplateParm does not match rgbTemplateParms",
            ));
        }
        self.header.write_to(writer)?;
        writer.write_u8(self.condition_type)?;
        writer.write_u8(self.comparison_operator)?;
        writer.write_u16(
            u16::try_from(formula1.len())
                .map_err(|_| Error::Limit("CF12 formula1 exceeds u16 bytes".into()))?,
        )?;
        writer.write_u16(
            u16::try_from(formula2.len())
                .map_err(|_| Error::Limit("CF12 formula2 exceeds u16 bytes".into()))?,
        )?;
        self.format.write_to(writer)?;
        writer.write_all(&formula1)?;
        writer.write_all(&formula2)?;
        writer.write_u16(
            u16::try_from(active_formula.len())
                .map_err(|_| Error::Limit("CF12 active formula exceeds u16 bytes".into()))?,
        )?;
        writer.write_all(&active_formula)?;
        writer.write_u8(self.option_flags)?;
        writer.write_u16(self.priority)?;
        writer.write_u16(self.template_id)?;
        writer.write_u8(self.template_parameter_size)?;
        writer.write_all(&template_bytes)?;
        match &self.condition_data {
            Cf12ConditionData::None => {}
            Cf12ConditionData::Gradient(value) => value.write_to(writer)?,
            Cf12ConditionData::DataBar(value) => value.write_to(writer)?,
            Cf12ConditionData::Filter(value) => value.write_to(writer)?,
            Cf12ConditionData::Multistate(value) => value.write_to(writer)?,
        }
        Ok(())
    }
}

impl SdkRead for ConditionalFormattingGroup12Record {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        Ok(Self {
            header: FrtRefHeaderU::read_from(reader)?,
            group: ConditionalFormattingGroupRecord::read_from(reader)?,
        })
    }
}

impl SdkWrite for ConditionalFormattingGroup12Record {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        self.group.write_to(writer)
    }
}

impl SdkRead for ChartLinkedDataRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let link_type = reader.read_u8()?;
        let reference_type = reader.read_u8()?;
        let flags = ChartLinkedDataFlags::from_bits_retain(reader.read_u16()?);
        let number_format_index = reader.read_u16()?;
        let formula_length = usize::from(reader.read_u16()?);
        let formula = FormulaTokenStream::from_bytes(&reader.read_vec(formula_length)?)?;
        Ok(Self {
            link_type,
            reference_type,
            flags,
            number_format_index,
            formula,
        })
    }
}

impl SdkWrite for ChartLinkedDataRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let formula = self.formula.to_bytes()?;
        writer.write_u8(self.link_type)?;
        writer.write_u8(self.reference_type)?;
        writer.write_u16(self.flags.bits())?;
        writer.write_u16(self.number_format_index)?;
        writer.write_u16(
            u16::try_from(formula.len())
                .map_err(|_| Error::Limit("LinkedData formula exceeds u16 bytes".into()))?,
        )?;
        writer.write_all(&formula)?;
        Ok(())
    }
}

impl SdkRead for ChartSeriesListRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let count = usize::from(reader.read_u16()?);
        let mut series_numbers = Vec::with_capacity(count);
        for _ in 0..count {
            series_numbers.push(reader.read_u16()?);
        }
        Ok(Self { series_numbers })
    }
}

impl SdkWrite for ChartSeriesListRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(
            u16::try_from(self.series_numbers.len())
                .map_err(|_| Error::Limit("SeriesList count exceeds u16".into()))?,
        )?;
        for number in &self.series_numbers {
            writer.write_u16(*number)?;
        }
        Ok(())
    }
}

impl FormulaCachedResult {
    fn from_bits(bits: u64) -> Result<Self> {
        if bits & 0xffff_0000_0000_0000 != 0xffff_0000_0000_0000 {
            return Ok(Self::NumberBits(bits));
        }
        let bytes = bits.to_le_bytes();
        if bytes[0] > 3 {
            return Err(Error::invalid(
                0,
                format!(
                    "Formula cached result has invalid special kind {}",
                    bytes[0]
                ),
            ));
        }
        Ok(Self::Special(FormulaSpecialCachedResult {
            kind: bytes[0],
            reserved1: bytes[1],
            value: bytes[2],
            reserved2: [bytes[3], bytes[4], bytes[5]],
        }))
    }

    fn bits(self) -> u64 {
        match self {
            Self::NumberBits(bits) => bits,
            Self::Special(value) => u64::from_le_bytes([
                value.kind,
                value.reserved1,
                value.value,
                value.reserved2[0],
                value.reserved2[1],
                value.reserved2[2],
                0xff,
                0xff,
            ]),
        }
    }
}

impl SdkRead for HeaderFooterRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        if reader.remaining()? == 0 {
            return Ok(Self::EmptyPayload);
        }
        let count = usize::from(reader.read_u16()?);
        if count == 0 && reader.remaining()? == 0 {
            return Ok(Self::EmptyCountOnly);
        }
        let flags = reader.read_u8()?;
        let characters = if flags & 1 == 0 {
            XlStringCharacters::Compressed(reader.read_vec(count)?)
        } else {
            let byte_count = count
                .checked_mul(2)
                .ok_or_else(|| Error::Limit("header/footer byte count overflow".into()))?;
            let bytes = reader.read_vec(byte_count)?;
            XlStringCharacters::Unicode(
                bytes
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect(),
            )
        };
        Ok(Self::Text { flags, characters })
    }
}

impl SdkWrite for HeaderFooterRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::EmptyPayload => Ok(()),
            Self::EmptyCountOnly => writer.write_u16(0),
            Self::Text { flags, characters } => {
                let count = match characters {
                    XlStringCharacters::Compressed(values) => {
                        if flags & 1 != 0 {
                            return Err(Error::invalid(
                                writer.position()?,
                                "compressed header/footer has UTF-16 flag",
                            ));
                        }
                        values.len()
                    }
                    XlStringCharacters::Unicode(values) => {
                        if flags & 1 == 0 {
                            return Err(Error::invalid(
                                writer.position()?,
                                "Unicode header/footer lacks UTF-16 flag",
                            ));
                        }
                        values.len()
                    }
                };
                writer.write_u16(u16::try_from(count).map_err(|_| {
                    Error::Limit("header/footer character count exceeds u16".into())
                })?)?;
                writer.write_u8(*flags)?;
                match characters {
                    XlStringCharacters::Compressed(values) => writer.write_all(values)?,
                    XlStringCharacters::Unicode(values) => {
                        for value in values {
                            writer.write_u16(*value)?;
                        }
                    }
                }
                Ok(())
            }
        }
    }
}

impl SdkWrite for DimensionsRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u32(self.first_row)?;
        writer.write_u32(self.last_row_exclusive)?;
        writer.write_u16(self.first_column)?;
        writer.write_u16(self.last_column_exclusive)?;
        writer.write_u16(self.reserved)?;
        if let Some(value) = self.compatibility_extra {
            writer.write_u16(value)?;
        }
        Ok(())
    }
}

impl SdkRead for BoolErrRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let cell = CellHeader::read_from(reader)?;
        let value = match reader.remaining()? {
            2 => BoolErrValue::Byte(reader.read_u8()?),
            3 => BoolErrValue::Word(reader.read_u16()?),
            remaining => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("BoolErr value and flag have invalid length {remaining}"),
                ));
            }
        };
        let is_error = reader.read_u8()?;
        Ok(Self {
            cell,
            value,
            is_error,
        })
    }
}

impl SdkWrite for BoolErrRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.cell.write_to(writer)?;
        match self.value {
            BoolErrValue::Byte(value) => writer.write_u8(value)?,
            BoolErrValue::Word(value) => writer.write_u16(value)?,
        }
        writer.write_u8(self.is_error)
    }
}

impl SdkRead for ColInfoRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let first_column = reader.read_u16()?;
        let last_column = reader.read_u16()?;
        let width = reader.read_u16()?;
        let format_index = reader.read_u16()?;
        let flags = reader.read_u16()?;
        let reserved = match reader.remaining()? {
            0 => ColInfoReserved::Missing,
            1 => ColInfoReserved::Byte(reader.read_u8()?),
            2 => ColInfoReserved::Word(reader.read_u16()?),
            remaining => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("ColInfo reserved field has invalid length {remaining}"),
                ));
            }
        };
        Ok(Self {
            first_column,
            last_column,
            width,
            format_index,
            flags,
            reserved,
        })
    }
}

impl SdkRead for MulRkRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let row = reader.read_u16()?;
        let first_column = reader.read_u16()?;
        let remaining = reader.remaining()?;
        if remaining < 8 || (remaining - 2) % 6 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                format!("MulRk cell region has invalid length {remaining}"),
            ));
        }
        let count = usize::try_from((remaining - 2) / 6)
            .map_err(|_| Error::Limit("MulRk cell count exceeds usize".into()))?;
        reader.ensure_allocation(count, 6)?;
        let mut cells = Vec::with_capacity(count);
        for _ in 0..count {
            cells.push(MulRkCell::read_from(reader)?);
        }
        let last_column = reader.read_u16()?;
        Ok(Self {
            row,
            first_column,
            cells,
            last_column,
        })
    }
}

impl SdkWrite for MulRkRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.row)?;
        writer.write_u16(self.first_column)?;
        for cell in &self.cells {
            cell.write_to(writer)?;
        }
        writer.write_u16(self.last_column)
    }
}

impl SdkRead for MulBlankRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let row = reader.read_u16()?;
        let first_column = reader.read_u16()?;
        let remaining = reader.remaining()?;
        if remaining < 4 || (remaining - 2) % 2 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                format!("MulBlank format-index region has invalid length {remaining}"),
            ));
        }
        let count = usize::try_from((remaining - 2) / 2)
            .map_err(|_| Error::Limit("MulBlank format-index count exceeds usize".into()))?;
        reader.ensure_allocation(count, 2)?;
        let mut format_indices = Vec::with_capacity(count);
        for _ in 0..count {
            format_indices.push(reader.read_u16()?);
        }
        let last_column = reader.read_u16()?;
        Ok(Self {
            row,
            first_column,
            format_indices,
            last_column,
        })
    }
}

impl SdkWrite for MulBlankRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.row)?;
        writer.write_u16(self.first_column)?;
        for format_index in &self.format_indices {
            writer.write_u16(*format_index)?;
        }
        writer.write_u16(self.last_column)
    }
}

impl SdkWrite for ColInfoRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.first_column)?;
        writer.write_u16(self.last_column)?;
        writer.write_u16(self.width)?;
        writer.write_u16(self.format_index)?;
        writer.write_u16(self.flags)?;
        match self.reserved {
            ColInfoReserved::Missing => {}
            ColInfoReserved::Byte(value) => writer.write_u8(value)?,
            ColInfoReserved::Word(value) => writer.write_u16(value)?,
        }
        Ok(())
    }
}

impl SdkRead for CrnRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let last_column = reader.read_u8()?;
        let first_column = reader.read_u8()?;
        let row = reader.read_u16()?;
        let count = last_column
            .checked_sub(first_column)
            .map_or(0usize, |difference| usize::from(difference) + 1);
        if count == 0 {
            return Err(Error::invalid(
                reader.position()?,
                "CRN last column precedes first column",
            ));
        }
        let remaining = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("CRN constant data length exceeds usize".into()))?;
        let bytes = reader.read_vec(remaining)?;
        let mut cursor = 0usize;
        let mut values = Vec::with_capacity(count);
        for _ in 0..count {
            values.push(BiffConstant::read(&bytes, &mut cursor)?);
        }
        if cursor != bytes.len() {
            return Err(Error::invalid(
                cursor as u64,
                "CRN contains trailing constant data",
            ));
        }
        Ok(Self {
            last_column,
            first_column,
            row,
            values,
        })
    }
}

impl SdkWrite for CrnRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected = self
            .last_column
            .checked_sub(self.first_column)
            .map_or(0usize, |difference| usize::from(difference) + 1);
        if expected == 0 || self.values.len() != expected {
            return Err(Error::invalid(
                writer.position()?,
                "CRN column range and constant count disagree",
            ));
        }
        writer.write_u8(self.last_column)?;
        writer.write_u8(self.first_column)?;
        writer.write_u16(self.row)?;
        let mut bytes = Vec::new();
        for value in &self.values {
            value.write(&mut bytes)?;
        }
        writer.write_all(&bytes)?;
        Ok(())
    }
}

impl SdkRead for XfExtRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let reserved1 = reader.read_u16()?;
        let xf_index = reader.read_u16()?;
        let reserved2 = reader.read_u16()?;
        let count = usize::from(reader.read_u16()?);
        reader.ensure_allocation(count, 4)?;
        let mut properties = Vec::with_capacity(count);
        for _ in 0..count {
            properties.push(ExtProperty::read(reader)?);
        }
        Ok(Self {
            header,
            reserved1,
            xf_index,
            reserved2,
            properties,
        })
    }
}

impl SdkRead for ExtendedHeaderFooterRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let sheet_view_guid = reader.read_array()?;
        let flags = ExtendedHeaderFooterFlags::from_bits_retain(reader.read_u16()?);
        let even_header_character_count = reader.read_u16()?;
        let even_footer_character_count = reader.read_u16()?;
        let first_header_character_count = reader.read_u16()?;
        let first_footer_character_count = reader.read_u16()?;
        let even_header = if even_header_character_count == 0 {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::from(even_header_character_count),
            )?)
        };
        let even_footer = if even_footer_character_count == 0 {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::from(even_footer_character_count),
            )?)
        };
        let first_header = if first_header_character_count == 0 {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::from(first_header_character_count),
            )?)
        };
        let first_footer = if first_footer_character_count == 0 {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::from(first_footer_character_count),
            )?)
        };
        Ok(Self {
            header,
            sheet_view_guid,
            flags,
            even_header_character_count,
            even_footer_character_count,
            first_header_character_count,
            first_footer_character_count,
            even_header,
            even_footer,
            first_header,
            first_footer,
        })
    }
}

impl SdkRead for HfPictureRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let flags = HfPictureFlags::from_bits_retain(reader.read_u8()?);
        let reserved = reader.read_u8()?;
        let length = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("HFPicture drawing length exceeds usize".into()))?;
        let bytes = reader.read_vec(length)?;
        let drawing = match OfficeArtStream::from_bytes(&bytes) {
            Ok(stream) => MsoDrawingData::Complete(stream),
            Err(error) => match OfficeArtPartialStream::from_bytes_with_limits(
                &bytes,
                Limits::default(),
                error.to_string(),
            ) {
                Ok(partial) => MsoDrawingData::Partial(partial),
                Err(_) => MsoDrawingData::Incomplete {
                    bytes,
                    reason: error.to_string(),
                },
            },
        };
        Ok(Self {
            header,
            flags,
            reserved,
            drawing,
        })
    }
}

impl SdkWrite for HfPictureRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        writer.write_u8(self.flags.bits())?;
        writer.write_u8(self.reserved)?;
        match &self.drawing {
            MsoDrawingData::Complete(stream) => writer.write_all(&stream.to_bytes()?)?,
            MsoDrawingData::Partial(stream) => writer.write_all(&stream.to_bytes()?)?,
            MsoDrawingData::Incomplete { bytes, .. } => writer.write_all(bytes)?,
        }
        Ok(())
    }
}

impl SdkRead for PbString {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = reader.read_u16()?;
        let count = usize::from(header & 0x7fff);
        let characters = if header & 0x8000 != 0 {
            PbStringCharacters::Ansi(reader.read_vec(count)?)
        } else {
            let byte_count = count
                .checked_mul(2)
                .ok_or_else(|| Error::Limit("PBString byte count overflow".into()))?;
            let bytes = reader.read_vec(byte_count)?;
            PbStringCharacters::Unicode(
                bytes
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect(),
            )
        };
        Ok(Self { characters })
    }
}

impl SdkWrite for PbString {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        match &self.characters {
            PbStringCharacters::Ansi(values) => {
                let count = u16::try_from(values.len())
                    .map_err(|_| Error::Limit("ANSI PBString exceeds u16".into()))?;
                if count > 0x7fff {
                    return Err(Error::Limit("ANSI PBString exceeds 15-bit count".into()));
                }
                writer.write_u16(count | 0x8000)?;
                writer.write_all(values)?;
            }
            PbStringCharacters::Unicode(values) => {
                let count = u16::try_from(values.len())
                    .map_err(|_| Error::Limit("Unicode PBString exceeds u16".into()))?;
                if count > 0x7fff {
                    return Err(Error::Limit("Unicode PBString exceeds 15-bit count".into()));
                }
                writer.write_u16(count)?;
                for value in values {
                    writer.write_u16(*value)?;
                }
            }
        }
        Ok(())
    }
}

impl SdkRead for FactoidType {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let payload_size = reader.read_u32()?;
        let mut child = reader.sub_reader(u64::from(payload_size))?;
        let id = child.read_u32()?;
        let uri = PbString::read_from(&mut child)?;
        let tag = PbString::read_from(&mut child)?;
        let download_url = PbString::read_from(&mut child)?;
        if child.remaining()? != 0 {
            return Err(Error::invalid(
                child.position()?,
                "FactoidType has trailing bytes",
            ));
        }
        Ok(Self {
            id,
            uri,
            tag,
            download_url,
        })
    }
}

impl SdkWrite for FactoidType {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut payload_writer = Writer::new(Cursor::new(Vec::new()));
        payload_writer.write_u32(self.id)?;
        self.uri.write_to(&mut payload_writer)?;
        self.tag.write_to(&mut payload_writer)?;
        self.download_url.write_to(&mut payload_writer)?;
        let payload = payload_writer.into_inner().into_inner();
        writer.write_u32(
            u32::try_from(payload.len())
                .map_err(|_| Error::Limit("FactoidType exceeds u32 bytes".into()))?,
        )?;
        writer.write_all(&payload)?;
        Ok(())
    }
}

impl SdkRead for PropertyBagStore {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let factoid_type_count = usize::try_from(reader.read_u32()?)
            .map_err(|_| Error::Limit("FactoidType count exceeds usize".into()))?;
        let maximum_count = usize::try_from(reader.remaining()? / 10)
            .map_err(|_| Error::Limit("PropertyBagStore length exceeds usize".into()))?;
        if factoid_type_count > maximum_count {
            return Err(Error::invalid(
                reader.position()?,
                "FactoidType count exceeds remaining bytes",
            ));
        }
        let mut factoid_types = Vec::with_capacity(factoid_type_count);
        for _ in 0..factoid_type_count {
            factoid_types.push(FactoidType::read_from(reader)?);
        }
        let header_size = reader.read_u16()?;
        let version = reader.read_u16()?;
        let factoid_count = reader.read_u32()?;
        let string_count = usize::try_from(reader.read_u32()?)
            .map_err(|_| Error::Limit("PropertyBagStore string count exceeds usize".into()))?;
        let maximum_strings = usize::try_from(reader.remaining()? / 2)
            .map_err(|_| Error::Limit("PropertyBagStore length exceeds usize".into()))?;
        if string_count > maximum_strings {
            return Err(Error::invalid(
                reader.position()?,
                "PropertyBagStore string count exceeds remaining bytes",
            ));
        }
        let mut strings = Vec::with_capacity(string_count);
        for _ in 0..string_count {
            strings.push(PbString::read_from(reader)?);
        }
        Ok(Self {
            factoid_types,
            header_size,
            version,
            factoid_count,
            strings,
        })
    }
}

impl SdkWrite for PropertyBagStore {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u32(
            u32::try_from(self.factoid_types.len())
                .map_err(|_| Error::Limit("FactoidType count exceeds u32".into()))?,
        )?;
        for value in &self.factoid_types {
            value.write_to(writer)?;
        }
        writer.write_u16(self.header_size)?;
        writer.write_u16(self.version)?;
        writer.write_u32(self.factoid_count)?;
        writer.write_u32(
            u32::try_from(self.strings.len())
                .map_err(|_| Error::Limit("PropertyBagStore string count exceeds u32".into()))?,
        )?;
        for value in &self.strings {
            value.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for FeatureHeaderRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let shared_feature_type = reader.read_u16()?;
        let reserved = reader.read_u8()?;
        let header_data_marker = reader.read_u32()?;
        let data = match header_data_marker {
            0 => FeatureHeaderData::None,
            u32::MAX => match shared_feature_type {
                0x0002 => FeatureHeaderData::EnhancedProtection(
                    EnhancedProtectionFlags::from_bits_retain(reader.read_u32()?),
                ),
                0x0004 => FeatureHeaderData::PropertyBagStore(PropertyBagStore::read_from(reader)?),
                value => {
                    return Err(Error::invalid(
                        reader.position()?,
                        format!("FeatHdr type 0x{value:04x} has unsupported header data"),
                    ));
                }
            },
            value => {
                let remaining = usize::try_from(reader.remaining()?)
                    .map_err(|_| Error::Limit("FeatHdr remainder exceeds usize".into()))?;
                FeatureHeaderData::Malformed {
                    marker: value,
                    payload: reader.read_vec(remaining)?,
                }
            }
        };
        Ok(Self {
            header,
            shared_feature_type,
            reserved,
            data,
        })
    }
}

impl SdkWrite for FeatureHeaderRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        writer.write_u16(self.shared_feature_type)?;
        writer.write_u8(self.reserved)?;
        match &self.data {
            FeatureHeaderData::None => writer.write_u32(0)?,
            FeatureHeaderData::EnhancedProtection(flags) => {
                if self.shared_feature_type != 0x0002 {
                    return Err(Error::invalid(
                        0,
                        "EnhancedProtection FeatHdr type mismatch",
                    ));
                }
                writer.write_u32(u32::MAX)?;
                writer.write_u32(flags.bits())?;
            }
            FeatureHeaderData::PropertyBagStore(value) => {
                if self.shared_feature_type != 0x0004 {
                    return Err(Error::invalid(0, "PropertyBagStore FeatHdr type mismatch"));
                }
                writer.write_u32(u32::MAX)?;
                value.write_to(writer)?;
            }
            FeatureHeaderData::Malformed { marker, payload } => {
                writer.write_u32(*marker)?;
                writer.write_all(payload)?;
            }
        }
        Ok(())
    }
}

impl SecurityIdentifier {
    fn from_bytes_at(bytes: &[u8], offset: usize) -> Result<(Self, usize)> {
        let mut cursor = offset;
        let revision = take_u8(bytes, &mut cursor, "truncated SID revision")?;
        let count = usize::from(take_u8(
            bytes,
            &mut cursor,
            "truncated SID sub-authority count",
        )?);
        let authority = take_bytes(bytes, &mut cursor, 6, "truncated SID authority")?;
        let mut identifier_authority = [0u8; 6];
        identifier_authority.copy_from_slice(authority);
        let maximum_count = bytes.len().saturating_sub(cursor) / 4;
        if count > maximum_count {
            return Err(Error::invalid(
                cursor as u64,
                "SID sub-authorities exceed bounds",
            ));
        }
        let mut sub_authorities = Vec::with_capacity(count);
        for _ in 0..count {
            sub_authorities.push(take_u32(bytes, &mut cursor, "truncated SID sub-authority")?);
        }
        Ok((
            Self {
                revision,
                identifier_authority,
                sub_authorities,
            },
            cursor - offset,
        ))
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::with_capacity(8 + self.sub_authorities.len() * 4);
        bytes.push(self.revision);
        bytes.push(
            u8::try_from(self.sub_authorities.len())
                .map_err(|_| Error::Limit("SID has more than 255 sub-authorities".into()))?,
        );
        bytes.extend_from_slice(&self.identifier_authority);
        for value in &self.sub_authorities {
            bytes.extend_from_slice(&value.to_le_bytes());
        }
        Ok(bytes)
    }
}

impl BasicAceType {
    fn from_byte(value: u8, offset: usize) -> Result<Self> {
        match value {
            0x00 => Ok(Self::AccessAllowed),
            0x01 => Ok(Self::AccessDenied),
            0x02 => Ok(Self::SystemAudit),
            _ => Err(Error::invalid(
                offset as u64,
                format!("unsupported basic ACE type 0x{value:02x}"),
            )),
        }
    }

    const fn to_byte(self) -> u8 {
        match self {
            Self::AccessAllowed => 0x00,
            Self::AccessDenied => 0x01,
            Self::SystemAudit => 0x02,
        }
    }
}

impl BasicAce {
    fn from_bytes_at(bytes: &[u8], offset: usize) -> Result<(Self, usize)> {
        let mut cursor = offset;
        let ace_type =
            BasicAceType::from_byte(take_u8(bytes, &mut cursor, "truncated ACE type")?, offset)?;
        let flags = AceFlags::from_bits_retain(take_u8(bytes, &mut cursor, "truncated ACE flags")?);
        let declared_size = usize::from(take_u16(bytes, &mut cursor, "truncated ACE size")?);
        if declared_size < 16 || offset.saturating_add(declared_size) > bytes.len() {
            return Err(Error::invalid(offset as u64, "ACE size exceeds ACL bounds"));
        }
        let access_mask = take_u32(bytes, &mut cursor, "truncated ACE access mask")?;
        let (trustee, sid_size) = SecurityIdentifier::from_bytes_at(bytes, cursor)?;
        cursor = cursor
            .checked_add(sid_size)
            .ok_or_else(|| Error::Limit("ACE cursor overflow".into()))?;
        if cursor != offset + declared_size {
            return Err(Error::invalid(
                cursor as u64,
                "basic ACE size does not match its SID",
            ));
        }
        Ok((
            Self {
                ace_type,
                flags,
                access_mask,
                trustee,
            },
            declared_size,
        ))
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        let sid = self.trustee.to_bytes()?;
        let size = 8usize
            .checked_add(sid.len())
            .ok_or_else(|| Error::Limit("ACE size overflow".into()))?;
        let mut bytes = Vec::with_capacity(size);
        bytes.push(self.ace_type.to_byte());
        bytes.push(self.flags.bits());
        bytes.extend_from_slice(
            &u16::try_from(size)
                .map_err(|_| Error::Limit("ACE exceeds u16 bytes".into()))?
                .to_le_bytes(),
        );
        bytes.extend_from_slice(&self.access_mask.to_le_bytes());
        bytes.extend_from_slice(&sid);
        Ok(bytes)
    }
}

impl AccessControlList {
    fn from_bytes_at(bytes: &[u8], offset: usize) -> Result<(Self, usize)> {
        let mut cursor = offset;
        let revision = take_u8(bytes, &mut cursor, "truncated ACL revision")?;
        let reserved1 = take_u8(bytes, &mut cursor, "truncated ACL reserved byte")?;
        let declared_size = usize::from(take_u16(bytes, &mut cursor, "truncated ACL size")?);
        let entry_count = usize::from(take_u16(bytes, &mut cursor, "truncated ACL entry count")?);
        let reserved2 = take_u16(bytes, &mut cursor, "truncated ACL reserved word")?;
        if declared_size < 8 || offset.saturating_add(declared_size) > bytes.len() {
            return Err(Error::invalid(
                offset as u64,
                "ACL size exceeds descriptor bounds",
            ));
        }
        let acl_end = offset + declared_size;
        let mut entries = Vec::with_capacity(entry_count);
        for _ in 0..entry_count {
            let (entry, size) = BasicAce::from_bytes_at(&bytes[..acl_end], cursor)?;
            cursor = cursor
                .checked_add(size)
                .ok_or_else(|| Error::Limit("ACL cursor overflow".into()))?;
            entries.push(entry);
        }
        let padding = bytes[cursor..acl_end].to_vec();
        Ok((
            Self {
                revision,
                reserved1,
                reserved2,
                entries,
                padding,
            },
            declared_size,
        ))
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut encoded_entries = Vec::with_capacity(self.entries.len());
        let mut size = 8usize;
        for entry in &self.entries {
            let bytes = entry.to_bytes()?;
            size = size
                .checked_add(bytes.len())
                .ok_or_else(|| Error::Limit("ACL size overflow".into()))?;
            encoded_entries.push(bytes);
        }
        size = size
            .checked_add(self.padding.len())
            .ok_or_else(|| Error::Limit("ACL padding size overflow".into()))?;
        let mut bytes = Vec::with_capacity(size);
        bytes.push(self.revision);
        bytes.push(self.reserved1);
        bytes.extend_from_slice(
            &u16::try_from(size)
                .map_err(|_| Error::Limit("ACL exceeds u16 bytes".into()))?
                .to_le_bytes(),
        );
        bytes.extend_from_slice(
            &u16::try_from(self.entries.len())
                .map_err(|_| Error::Limit("ACL has more than u16 ACEs".into()))?
                .to_le_bytes(),
        );
        bytes.extend_from_slice(&self.reserved2.to_le_bytes());
        for entry in encoded_entries {
            bytes.extend_from_slice(&entry);
        }
        bytes.extend_from_slice(&self.padding);
        Ok(bytes)
    }
}

impl SecurityDescriptor {
    fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() < 20 {
            return Err(Error::invalid(
                0,
                "security descriptor is shorter than 20 bytes",
            ));
        }
        let mut cursor = 0usize;
        let revision = take_u8(bytes, &mut cursor, "truncated security descriptor revision")?;
        let reserved = take_u8(
            bytes,
            &mut cursor,
            "truncated security descriptor reserved byte",
        )?;
        let control = SecurityDescriptorControl::from_bits_retain(take_u16(
            bytes,
            &mut cursor,
            "truncated security descriptor control",
        )?);
        let owner_offset = take_u32(bytes, &mut cursor, "truncated owner offset")?;
        let group_offset = take_u32(bytes, &mut cursor, "truncated group offset")?;
        let sacl_offset = take_u32(bytes, &mut cursor, "truncated SACL offset")?;
        let dacl_offset = take_u32(bytes, &mut cursor, "truncated DACL offset")?;
        let owner = parse_offset_sid(bytes, owner_offset, "owner SID")?;
        let group = parse_offset_sid(bytes, group_offset, "group SID")?;
        let sacl = parse_offset_acl(bytes, sacl_offset, "SACL")?;
        let dacl = parse_offset_acl(bytes, dacl_offset, "DACL")?;

        let mut occupied = vec![false; bytes.len()];
        occupied[..20].fill(true);
        if let Some(value) = &owner {
            mark_security_descriptor_range(
                &mut occupied,
                usize::try_from(value.offset)
                    .map_err(|_| Error::Limit("owner SID offset exceeds usize".into()))?,
                value.value.to_bytes()?.len(),
            )?;
        }
        if let Some(value) = &group {
            mark_security_descriptor_range(
                &mut occupied,
                usize::try_from(value.offset)
                    .map_err(|_| Error::Limit("group SID offset exceeds usize".into()))?,
                value.value.to_bytes()?.len(),
            )?;
        }
        if let Some(value) = &sacl {
            mark_security_descriptor_range(
                &mut occupied,
                usize::try_from(value.offset)
                    .map_err(|_| Error::Limit("SACL offset exceeds usize".into()))?,
                value.value.to_bytes()?.len(),
            )?;
        }
        if let Some(value) = &dacl {
            mark_security_descriptor_range(
                &mut occupied,
                usize::try_from(value.offset)
                    .map_err(|_| Error::Limit("DACL offset exceeds usize".into()))?,
                value.value.to_bytes()?.len(),
            )?;
        }
        let mut padding = Vec::new();
        let mut index = 20usize;
        while index < bytes.len() {
            if occupied[index] {
                index += 1;
                continue;
            }
            let start = index;
            while index < bytes.len() && !occupied[index] {
                index += 1;
            }
            padding.push(SecurityDescriptorPadding {
                offset: u32::try_from(start)
                    .map_err(|_| Error::Limit("security descriptor offset exceeds u32".into()))?,
                bytes: bytes[start..index].to_vec(),
            });
        }
        Ok(Self {
            revision,
            reserved,
            control,
            owner,
            group,
            sacl,
            dacl,
            padding,
        })
    }

    fn to_bytes(&self, declared_size: usize) -> Result<Vec<u8>> {
        if declared_size < 20 {
            return Err(Error::invalid(
                0,
                "security descriptor is shorter than 20 bytes",
            ));
        }
        let mut bytes = vec![0u8; declared_size];
        let mut occupied = vec![false; declared_size];
        let owner_offset = self.owner.as_ref().map_or(0, |value| value.offset);
        let group_offset = self.group.as_ref().map_or(0, |value| value.offset);
        let sacl_offset = self.sacl.as_ref().map_or(0, |value| value.offset);
        let dacl_offset = self.dacl.as_ref().map_or(0, |value| value.offset);
        let mut header = Vec::with_capacity(20);
        header.push(self.revision);
        header.push(self.reserved);
        header.extend_from_slice(&self.control.bits().to_le_bytes());
        header.extend_from_slice(&owner_offset.to_le_bytes());
        header.extend_from_slice(&group_offset.to_le_bytes());
        header.extend_from_slice(&sacl_offset.to_le_bytes());
        header.extend_from_slice(&dacl_offset.to_le_bytes());
        place_security_descriptor_bytes(&mut bytes, &mut occupied, 0, &header, "header")?;
        for value in &self.padding {
            place_security_descriptor_bytes(
                &mut bytes,
                &mut occupied,
                usize::try_from(value.offset).map_err(|_| {
                    Error::Limit("security descriptor padding offset exceeds usize".into())
                })?,
                &value.bytes,
                "padding",
            )?;
        }
        if let Some(value) = &self.owner {
            place_security_descriptor_bytes(
                &mut bytes,
                &mut occupied,
                usize::try_from(value.offset)
                    .map_err(|_| Error::Limit("owner SID offset exceeds usize".into()))?,
                &value.value.to_bytes()?,
                "owner SID",
            )?;
        }
        if let Some(value) = &self.group {
            place_security_descriptor_bytes(
                &mut bytes,
                &mut occupied,
                usize::try_from(value.offset)
                    .map_err(|_| Error::Limit("group SID offset exceeds usize".into()))?,
                &value.value.to_bytes()?,
                "group SID",
            )?;
        }
        if let Some(value) = &self.sacl {
            place_security_descriptor_bytes(
                &mut bytes,
                &mut occupied,
                usize::try_from(value.offset)
                    .map_err(|_| Error::Limit("SACL offset exceeds usize".into()))?,
                &value.value.to_bytes()?,
                "SACL",
            )?;
        }
        if let Some(value) = &self.dacl {
            place_security_descriptor_bytes(
                &mut bytes,
                &mut occupied,
                usize::try_from(value.offset)
                    .map_err(|_| Error::Limit("DACL offset exceeds usize".into()))?,
                &value.value.to_bytes()?,
                "DACL",
            )?;
        }
        Ok(bytes)
    }
}

fn parse_offset_sid(
    bytes: &[u8],
    offset: u32,
    label: &str,
) -> Result<Option<OffsetSecurityIdentifier>> {
    if offset == 0 {
        return Ok(None);
    }
    let offset_usize = usize::try_from(offset)
        .map_err(|_| Error::Limit(format!("{label} offset exceeds usize")))?;
    let (value, _) = SecurityIdentifier::from_bytes_at(bytes, offset_usize)?;
    Ok(Some(OffsetSecurityIdentifier { offset, value }))
}

fn parse_offset_acl(
    bytes: &[u8],
    offset: u32,
    label: &str,
) -> Result<Option<OffsetAccessControlList>> {
    if offset == 0 {
        return Ok(None);
    }
    let offset_usize = usize::try_from(offset)
        .map_err(|_| Error::Limit(format!("{label} offset exceeds usize")))?;
    let (value, _) = AccessControlList::from_bytes_at(bytes, offset_usize)?;
    Ok(Some(OffsetAccessControlList { offset, value }))
}

fn mark_security_descriptor_range(occupied: &mut [bool], offset: usize, size: usize) -> Result<()> {
    let end = offset
        .checked_add(size)
        .ok_or_else(|| Error::Limit("security descriptor range overflow".into()))?;
    let Some(range) = occupied.get_mut(offset..end) else {
        return Err(Error::invalid(
            offset as u64,
            "security descriptor component exceeds bounds",
        ));
    };
    range.fill(true);
    Ok(())
}

fn place_security_descriptor_bytes(
    output: &mut [u8],
    occupied: &mut [bool],
    offset: usize,
    value: &[u8],
    label: &str,
) -> Result<()> {
    let end = offset
        .checked_add(value.len())
        .ok_or_else(|| Error::Limit(format!("security descriptor {label} range overflow")))?;
    let Some(destination) = output.get_mut(offset..end) else {
        return Err(Error::invalid(
            offset as u64,
            format!("security descriptor {label} exceeds declared size"),
        ));
    };
    let Some(markers) = occupied.get_mut(offset..end) else {
        return Err(Error::invalid(
            offset as u64,
            "security descriptor marker range exceeds bounds",
        ));
    };
    for ((destination, marker), source) in destination
        .iter_mut()
        .zip(markers.iter_mut())
        .zip(value.iter().copied())
    {
        if *marker && *destination != source {
            return Err(Error::invalid(
                offset as u64,
                format!("overlapping security descriptor {label} bytes disagree"),
            ));
        }
        *destination = source;
        *marker = true;
    }
    Ok(())
}

impl SdkRead for SecurityDescriptorContainer {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let declared_size = reader.read_u32()?;
        let size = usize::try_from(declared_size)
            .map_err(|_| Error::Limit("security descriptor size exceeds usize".into()))?;
        let bytes = reader.read_vec(size)?;
        Ok(Self {
            declared_size,
            descriptor: SecurityDescriptor::from_bytes(&bytes)?,
        })
    }
}

impl SdkWrite for SecurityDescriptorContainer {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let size = usize::try_from(self.declared_size)
            .map_err(|_| Error::Limit("security descriptor size exceeds usize".into()))?;
        let bytes = self.descriptor.to_bytes(size)?;
        writer.write_u32(self.declared_size)?;
        writer.write_all(&bytes)?;
        Ok(())
    }
}

impl SdkRead for FeatureProtection {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let flags = FeatureProtectionFlags::from_bits_retain(reader.read_u32()?);
        let password_verifier = reader.read_u32()?;
        let title = XlUnicodeString::read_from(reader)?;
        let security_descriptor = if flags.contains(FeatureProtectionFlags::HAS_SECURITY_DESCRIPTOR)
        {
            Some(SecurityDescriptorContainer::read_from(reader)?)
        } else {
            None
        };
        Ok(Self {
            flags,
            password_verifier,
            title,
            security_descriptor,
        })
    }
}

impl SdkWrite for FeatureProtection {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let has_descriptor = self
            .flags
            .contains(FeatureProtectionFlags::HAS_SECURITY_DESCRIPTOR);
        if has_descriptor != self.security_descriptor.is_some() {
            return Err(Error::invalid(
                0,
                "FeatProtection fSD does not match the security descriptor",
            ));
        }
        writer.write_u32(self.flags.bits())?;
        writer.write_u32(self.password_verifier)?;
        self.title.write_to(writer)?;
        if let Some(value) = &self.security_descriptor {
            value.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for FactoidData {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        Ok(Self {
            flags: FactoidDataFlags::from_bits_retain(reader.read_u8()?),
            property_bag: SmartTagPropertyBag::read_from(reader)?,
        })
    }
}

impl SdkWrite for FactoidData {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u8(self.flags.bits())?;
        self.property_bag.write_to(writer)
    }
}

impl SdkRead for FeatureSmartTag {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let hash_value = reader.read_u32()?;
        let count = usize::from(reader.read_u8()?);
        let mut factoids = Vec::with_capacity(count);
        for _ in 0..count {
            factoids.push(FactoidData::read_from(reader)?);
        }
        Ok(Self {
            hash_value,
            factoids,
        })
    }
}

impl SdkWrite for FeatureSmartTag {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u32(self.hash_value)?;
        writer.write_u8(
            u8::try_from(self.factoids.len())
                .map_err(|_| Error::Limit("FeatSmartTag has more than 255 factoids".into()))?,
        )?;
        for value in &self.factoids {
            value.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for FeatureRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let shared_feature_type = reader.read_u16()?;
        let reserved1 = reader.read_u8()?;
        let reserved2 = reader.read_u32()?;
        let reference_count = usize::from(reader.read_u16()?);
        let feature_data_size = reader.read_u32()?;
        let reserved3 = reader.read_u16()?;
        let maximum_references = usize::try_from(reader.remaining()? / 8)
            .map_err(|_| Error::Limit("Feat reference bounds exceed usize".into()))?;
        if reference_count > maximum_references {
            return Err(Error::invalid(
                reader.position()?,
                "Feat reference count exceeds remaining bytes",
            ));
        }
        let mut references = Vec::with_capacity(reference_count);
        for _ in 0..reference_count {
            references.push(CellRange::read_from(reader)?);
        }
        let data = match shared_feature_type {
            0x0002 => FeatureData::Protection(Box::new(FeatureProtection::read_from(reader)?)),
            0x0003 => FeatureData::FormulaErrors(FormulaErrorCheckFlags::from_bits_retain(
                reader.read_u32()?,
            )),
            0x0004 => FeatureData::SmartTags(FeatureSmartTag::read_from(reader)?),
            value => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("unsupported Feat shared feature type 0x{value:04x}"),
                ));
            }
        };
        Ok(Self {
            header,
            shared_feature_type,
            reserved1,
            reserved2,
            feature_data_size,
            reserved3,
            references,
            data,
        })
    }
}

impl SdkWrite for FeatureRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_type = match self.data {
            FeatureData::Protection(_) => 0x0002,
            FeatureData::FormulaErrors(_) => 0x0003,
            FeatureData::SmartTags(_) => 0x0004,
        };
        if self.shared_feature_type != expected_type {
            return Err(Error::invalid(
                0,
                "Feat shared feature type does not match its data",
            ));
        }
        self.header.write_to(writer)?;
        writer.write_u16(self.shared_feature_type)?;
        writer.write_u8(self.reserved1)?;
        writer.write_u32(self.reserved2)?;
        writer.write_u16(
            u16::try_from(self.references.len())
                .map_err(|_| Error::Limit("Feat has more than u16 references".into()))?,
        )?;
        writer.write_u32(self.feature_data_size)?;
        writer.write_u16(self.reserved3)?;
        for value in &self.references {
            value.write_to(writer)?;
        }
        match &self.data {
            FeatureData::Protection(value) => value.write_to(writer)?,
            FeatureData::FormulaErrors(value) => writer.write_u32(value.bits())?,
            FeatureData::SmartTags(value) => value.write_to(writer)?,
        }
        Ok(())
    }
}

impl SdkRead for DConnUnicodeStringSegmented {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let total_character_count = reader.read_u32()?;
        let total = usize::try_from(total_character_count)
            .map_err(|_| Error::Limit("DConn segmented string length exceeds usize".into()))?;
        let mut character_count = 0usize;
        let mut segments = Vec::new();
        while character_count < total {
            let segment = XlUnicodeString::read_from(reader)?;
            let segment_count = segment.text.character_count();
            if segment_count == 0 {
                return Err(Error::invalid(
                    reader.position()?,
                    "DConn segmented string contains an empty segment",
                ));
            }
            character_count = character_count
                .checked_add(segment_count)
                .ok_or_else(|| Error::Limit("DConn segmented string length overflow".into()))?;
            if character_count > total {
                return Err(Error::invalid(
                    reader.position()?,
                    "DConn segmented string exceeds its declared character count",
                ));
            }
            segments.push(segment);
        }
        Ok(Self {
            total_character_count,
            segments,
        })
    }
}

impl SdkWrite for DConnUnicodeStringSegmented {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let actual_count = self.segments.iter().try_fold(0usize, |count, segment| {
            if segment.text.character_count() == 0 {
                return Err(Error::invalid(
                    0,
                    "DConn segmented string contains an empty segment",
                ));
            }
            count
                .checked_add(segment.text.character_count())
                .ok_or_else(|| Error::Limit("DConn segmented string length overflow".into()))
        })?;
        if u32::try_from(actual_count)
            .map_err(|_| Error::Limit("DConn segmented string length exceeds u32".into()))?
            != self.total_character_count
        {
            return Err(Error::invalid(
                0,
                "DConn segmented string character count mismatch",
            ));
        }
        writer.write_u32(self.total_character_count)?;
        for segment in &self.segments {
            segment.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for DConnStringSequence {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let count = usize::from(reader.read_u16()?);
        let mut strings = Vec::with_capacity(count);
        for _ in 0..count {
            strings.push(DConnUnicodeStringSegmented::read_from(reader)?);
        }
        Ok(Self { strings })
    }
}

impl SdkWrite for DConnStringSequence {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(
            u16::try_from(self.strings.len())
                .map_err(|_| Error::Limit("DConn string sequence exceeds u16 entries".into()))?,
        )?;
        for value in &self.strings {
            value.write_to(writer)?;
        }
        Ok(())
    }
}

impl DConnQueryFlags {
    fn from_raw(data_source_type: DataSourceType, raw: u16) -> Self {
        match data_source_type {
            DataSourceType::Web => Self::Web(DConnWebFlags::from_bits_retain(raw)),
            DataSourceType::OleDb => Self::OleDb(DConnOleDbFlags::from_bits_retain(raw)),
            DataSourceType::Ado => Self::Ado {
                reserved1: raw as u8,
                flags: DConnAdoFlags::from_bits_retain((raw >> 8) as u8),
            },
            _ => Self::Unused(raw),
        }
    }

    fn bits_for(self, data_source_type: DataSourceType) -> Result<u16> {
        match (data_source_type, self) {
            (DataSourceType::Web, Self::Web(value)) => Ok(value.bits()),
            (DataSourceType::OleDb, Self::OleDb(value)) => Ok(value.bits()),
            (DataSourceType::Ado, Self::Ado { reserved1, flags }) => {
                Ok(u16::from(reserved1) | (u16::from(flags.bits()) << 8))
            }
            (
                DataSourceType::Odbc | DataSourceType::Dao | DataSourceType::Text,
                Self::Unused(value),
            ) => Ok(value),
            _ => Err(Error::invalid(
                0,
                "connection query flags do not match the data source type",
            )),
        }
    }
}

impl TextQueryOptions {
    fn from_bits(bits: u32) -> Self {
        Self {
            file: bits & (1 << 0) != 0,
            delimited: bits & (1 << 1) != 0,
            code_page_kind: ((bits >> 2) & 0x03) as u8,
            prompt_for_file: bits & (1 << 4) != 0,
            new_code_page: ((bits >> 5) & 0x03ff) as u16,
            use_new_code_page: bits & (1 << 15) != 0,
            unused: (bits >> 16) as u16,
        }
    }

    fn bits(self) -> u32 {
        u32::from(self.file)
            | (u32::from(self.delimited) << 1)
            | (u32::from(self.code_page_kind & 0x03) << 2)
            | (u32::from(self.prompt_for_file) << 4)
            | (u32::from(self.new_code_page & 0x03ff) << 5)
            | (u32::from(self.use_new_code_page) << 15)
            | (u32::from(self.unused) << 16)
    }
}

impl TextQueryDelimiterOptions {
    fn from_bits(bits: u32) -> Self {
        Self {
            tab: bits & (1 << 0) != 0,
            space: bits & (1 << 1) != 0,
            comma: bits & (1 << 2) != 0,
            semicolon: bits & (1 << 3) != 0,
            custom: bits & (1 << 4) != 0,
            consecutive: bits & (1 << 5) != 0,
            text_delimiter: ((bits >> 6) & 0x03) as u8,
            custom_character: ((bits >> 8) & 0xffff) as u16,
            unused: (bits >> 24) as u8,
        }
    }

    fn bits(self) -> u32 {
        u32::from(self.tab)
            | (u32::from(self.space) << 1)
            | (u32::from(self.comma) << 2)
            | (u32::from(self.semicolon) << 3)
            | (u32::from(self.custom) << 4)
            | (u32::from(self.consecutive) << 5)
            | (u32::from(self.text_delimiter & 0x03) << 6)
            | (u32::from(self.custom_character) << 8)
            | (u32::from(self.unused) << 24)
    }
}

impl SdkRead for TextQuery {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let record_type = reader.read_u16()?;
        let reserved = reader.read_u16()?;
        let options = TextQueryOptions::from_bits(reader.read_u32()?);
        let starting_row = reader.read_i32()?;
        let delimiter_options = TextQueryDelimiterOptions::from_bits(reader.read_u32()?);
        let field_count = usize::try_from(reader.read_i32()?).map_err(|_| {
            Error::invalid(
                reader.position().unwrap_or(reader.start()),
                "negative TxtQry field count",
            )
        })?;
        let decimal_separator = reader.read_u8()?;
        let thousands_separator = reader.read_u8()?;
        let maximum_fields = usize::try_from(reader.remaining()? / 8)
            .map_err(|_| Error::Limit("TxtQry field bounds exceed usize".into()))?;
        if field_count > maximum_fields {
            return Err(Error::invalid(
                reader.position()?,
                "TxtQry field count exceeds remaining bytes",
            ));
        }
        let mut fields = Vec::with_capacity(field_count);
        for _ in 0..field_count {
            fields.push(TextQueryField::read_from(reader)?);
        }
        let source_file = XlUnicodeString::read_from(reader)?;
        Ok(Self {
            record_type,
            reserved,
            options,
            starting_row,
            delimiter_options,
            fields,
            decimal_separator,
            thousands_separator,
            source_file,
        })
    }
}

impl SdkWrite for TextQuery {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.record_type)?;
        writer.write_u16(self.reserved)?;
        writer.write_u32(self.options.bits())?;
        writer.write_i32(self.starting_row)?;
        writer.write_u32(self.delimiter_options.bits())?;
        writer.write_i32(
            i32::try_from(self.fields.len())
                .map_err(|_| Error::Limit("TxtQry has more than i32 fields".into()))?,
        )?;
        writer.write_u8(self.decimal_separator)?;
        writer.write_u8(self.thousands_separator)?;
        for value in &self.fields {
            value.write_to(writer)?;
        }
        self.source_file.write_to(writer)
    }
}

impl SdkRead for DConnOleDbConnection {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let drillthrough_row_limit = reader.read_u32()?;
        let count = usize::from(reader.read_u16()?);
        if count > 4 {
            return Err(Error::invalid(
                reader.position()?,
                "DConn OLE DB connection count exceeds four",
            ));
        }
        let mut valid_connection_types = Vec::with_capacity(count);
        for _ in 0..count {
            valid_connection_types.push(reader.read_u16()?);
        }
        let mut invalid_connection_types = Vec::with_capacity(4 - count);
        for _ in count..4 {
            invalid_connection_types.push(reader.read_u16()?);
        }
        let unused = reader.read_u16()?;
        let mut connection_strings = Vec::with_capacity(count);
        for _ in 0..count {
            connection_strings.push(DConnUnicodeStringSegmented::read_from(reader)?);
        }
        Ok(Self {
            drillthrough_row_limit,
            valid_connection_types,
            invalid_connection_types,
            unused,
            connection_strings,
        })
    }
}

impl SdkWrite for DConnOleDbConnection {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let count = self.valid_connection_types.len();
        if count > 4
            || self.invalid_connection_types.len() != 4 - count
            || self.connection_strings.len() != count
        {
            return Err(Error::invalid(
                0,
                "DConn OLE DB connection array sizes disagree",
            ));
        }
        writer.write_u32(self.drillthrough_row_limit)?;
        writer.write_u16(
            u16::try_from(count)
                .map_err(|_| Error::Limit("DConn OLE DB connection count exceeds u16".into()))?,
        )?;
        for value in &self.valid_connection_types {
            writer.write_u16(*value)?;
        }
        for value in &self.invalid_connection_types {
            writer.write_u16(*value)?;
        }
        writer.write_u16(self.unused)?;
        for value in &self.connection_strings {
            value.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for DConnWebConnection {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        Ok(Self {
            url: DConnStringSequence::read_from(reader)?,
            post_method: DConnStringSequence::read_from(reader)?,
        })
    }
}

impl SdkWrite for DConnWebConnection {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.url.write_to(writer)?;
        self.post_method.write_to(writer)
    }
}

impl SdkRead for DConnParameter {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let name = DConnUnicodeStringSegmented::read_from(reader)?;
        let type_word = reader.read_u16()?;
        let parameter_type = (type_word & 0x0007) as u8;
        let reserved = type_word >> 3;
        let sql_type = reader.read_i16()?;
        let name_word = reader.read_u16()?;
        let default_name = name_word & 1 != 0;
        let unused = name_word >> 1;
        let binding = match parameter_type {
            0 => DConnParameterBinding::Prompt(DConnUnicodeStringSegmented::read_from(reader)?),
            1 => {
                let value_type = reader.read_u16()?;
                let value = match value_type {
                    0x0001 => DConnParameterBindingValue::Numeric(reader.read_u64()?),
                    0x0002 => DConnParameterBindingValue::String {
                        reserved: reader.read_u64()?,
                        value: DConnUnicodeStringSegmented::read_from(reader)?,
                    },
                    0x0004 => {
                        let value = reader.read_u8()?;
                        let bytes = reader.read_vec(3)?;
                        let reserved1 = [bytes[0], bytes[1], bytes[2]];
                        DConnParameterBindingValue::Boolean {
                            value,
                            reserved1,
                            reserved2: reader.read_u32()?,
                        }
                    }
                    0x0800 => DConnParameterBindingValue::Integer {
                        value: reader.read_u32()?,
                        reserved: reader.read_u32()?,
                    },
                    value => {
                        return Err(Error::invalid(
                            reader.position()?,
                            format!("unsupported DConn parameter value type 0x{value:04x}"),
                        ));
                    }
                };
                DConnParameterBinding::Value { value_type, value }
            }
            value => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("unsupported DConn parameter binding type {value}"),
                ));
            }
        };
        Ok(Self {
            name,
            parameter_type,
            reserved,
            sql_type,
            default_name,
            unused,
            binding,
        })
    }
}

impl SdkWrite for DConnParameter {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_type = match self.binding {
            DConnParameterBinding::Prompt(_) => 0,
            DConnParameterBinding::Value { .. } => 1,
        };
        if self.parameter_type != expected_type || self.parameter_type > 7 {
            return Err(Error::invalid(
                0,
                "DConn parameter type does not match its binding",
            ));
        }
        self.name.write_to(writer)?;
        writer.write_u16(u16::from(self.parameter_type) | (self.reserved << 3))?;
        writer.write_i16(self.sql_type)?;
        writer.write_u16(u16::from(self.default_name) | (self.unused << 1))?;
        match &self.binding {
            DConnParameterBinding::Prompt(value) => value.write_to(writer)?,
            DConnParameterBinding::Value { value_type, value } => {
                let expected_value_type = match value {
                    DConnParameterBindingValue::Numeric(_) => 0x0001,
                    DConnParameterBindingValue::String { .. } => 0x0002,
                    DConnParameterBindingValue::Boolean { .. } => 0x0004,
                    DConnParameterBindingValue::Integer { .. } => 0x0800,
                };
                if *value_type != expected_value_type {
                    return Err(Error::invalid(
                        0,
                        "DConn parameter value type does not match its value",
                    ));
                }
                writer.write_u16(*value_type)?;
                match value {
                    DConnParameterBindingValue::Numeric(bits) => writer.write_u64(*bits)?,
                    DConnParameterBindingValue::String { reserved, value } => {
                        writer.write_u64(*reserved)?;
                        value.write_to(writer)?;
                    }
                    DConnParameterBindingValue::Boolean {
                        value,
                        reserved1,
                        reserved2,
                    } => {
                        writer.write_u8(*value)?;
                        writer.write_all(reserved1)?;
                        writer.write_u32(*reserved2)?;
                    }
                    DConnParameterBindingValue::Integer { value, reserved } => {
                        writer.write_u32(*value)?;
                        writer.write_u32(*reserved)?;
                    }
                }
            }
        }
        Ok(())
    }
}

impl SdkRead for DConnIdentifier {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        match reader.read_u8()? {
            0 => Ok(Self::None),
            1 => Ok(Self::QueryTable(DConnUnicodeStringSegmented::read_from(
                reader,
            )?)),
            2 => Ok(Self::PivotCache(reader.read_u16()?)),
            value => Err(Error::invalid(
                reader.position()?,
                format!("unsupported DConn identifier type {value}"),
            )),
        }
    }
}

impl SdkWrite for DConnIdentifier {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::None => writer.write_u8(0)?,
            Self::QueryTable(value) => {
                writer.write_u8(1)?;
                value.write_to(writer)?;
            }
            Self::PivotCache(value) => {
                writer.write_u8(2)?;
                writer.write_u16(*value)?;
            }
        }
        Ok(())
    }
}

impl SdkRead for DConnRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeaderOld::read_from(reader)?;
        let data_source_type = DataSourceType::read_from(reader)?;
        let flags = DConnFlags::from_bits_retain(reader.read_u16()?);
        let parameters_count = reader.read_u16()?;
        let reserved1 = reader.read_u16()?;
        let secondary_flags = DConnSecondaryFlags::from_bits_retain(reader.read_u8()?);
        let reserved2 = reader.read_u8()?;
        let raw_query_flags = reader.read_u16()?;
        let query_flags = DConnQueryFlags::from_raw(data_source_type, raw_query_flags);
        let edited_version = reader.read_u8()?;
        let refreshed_version = reader.read_u8()?;
        let minimum_refreshable_version = reader.read_u8()?;
        let refresh_interval_minutes = reader.read_u16()?;
        let html_format = reader.read_u16()?;
        let reconnection_method = reader.read_u32()?;
        let credential_method = reader.read_u8()?;
        let reserved3 = reader.read_u8()?;
        let source_data_file = DConnUnicodeStringSegmented::read_from(reader)?;
        let source_connection_file = DConnUnicodeStringSegmented::read_from(reader)?;
        let connection_name = DConnUnicodeStringSegmented::read_from(reader)?;
        let connection_description = DConnUnicodeStringSegmented::read_from(reader)?;
        let sso_application_id = DConnUnicodeStringSegmented::read_from(reader)?;
        let table_names = if flags.contains(DConnFlags::TABLE_NAMES) {
            Some(DConnUnicodeStringSegmented::read_from(reader)?)
        } else {
            None
        };
        let parameter_count = usize::from(parameters_count);
        let mut parameters = Vec::with_capacity(parameter_count);
        for _ in 0..parameter_count {
            parameters.push(DConnParameter::read_from(reader)?);
        }
        let connection = match data_source_type {
            DataSourceType::Odbc => {
                DConnConnection::Odbc(DConnUnicodeStringSegmented::read_from(reader)?)
            }
            DataSourceType::Web => DConnConnection::Web(DConnWebConnection::read_from(reader)?),
            DataSourceType::OleDb => {
                DConnConnection::OleDb(DConnOleDbConnection::read_from(reader)?)
            }
            DataSourceType::Text => DConnConnection::Text(TextQuery::read_from(reader)?),
            DataSourceType::Dao | DataSourceType::Ado => DConnConnection::None,
        };
        let sql = DConnStringSequence::read_from(reader)?;
        let saved_sql = DConnStringSequence::read_from(reader)?;
        let edit_web_page = DConnStringSequence::read_from(reader)?;
        let identifier = DConnIdentifier::read_from(reader)?;
        Ok(Self {
            header,
            data_source_type,
            flags,
            parameters_count,
            reserved1,
            secondary_flags,
            reserved2,
            query_flags,
            edited_version,
            refreshed_version,
            minimum_refreshable_version,
            refresh_interval_minutes,
            html_format,
            reconnection_method,
            credential_method,
            reserved3,
            source_data_file,
            source_connection_file,
            connection_name,
            connection_description,
            sso_application_id,
            table_names,
            parameters,
            connection,
            sql,
            saved_sql,
            edit_web_page,
            identifier,
        })
    }
}

impl SdkWrite for DConnRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if usize::from(self.parameters_count) != self.parameters.len() {
            return Err(Error::invalid(0, "DConn parameter count mismatch"));
        }
        if self.flags.contains(DConnFlags::TABLE_NAMES) != self.table_names.is_some() {
            return Err(Error::invalid(0, "DConn table-name flag mismatch"));
        }
        let query_flags = self.query_flags.bits_for(self.data_source_type)?;
        let connection_matches = matches!(
            (self.data_source_type, &self.connection),
            (DataSourceType::Odbc, DConnConnection::Odbc(_))
                | (DataSourceType::Web, DConnConnection::Web(_))
                | (DataSourceType::OleDb, DConnConnection::OleDb(_))
                | (DataSourceType::Text, DConnConnection::Text(_))
                | (
                    DataSourceType::Dao | DataSourceType::Ado,
                    DConnConnection::None
                )
        );
        if !connection_matches {
            return Err(Error::invalid(
                0,
                "DConn connection does not match the data source type",
            ));
        }
        self.header.write_to(writer)?;
        self.data_source_type.write_to(writer)?;
        writer.write_u16(self.flags.bits())?;
        writer.write_u16(self.parameters_count)?;
        writer.write_u16(self.reserved1)?;
        writer.write_u8(self.secondary_flags.bits())?;
        writer.write_u8(self.reserved2)?;
        writer.write_u16(query_flags)?;
        writer.write_u8(self.edited_version)?;
        writer.write_u8(self.refreshed_version)?;
        writer.write_u8(self.minimum_refreshable_version)?;
        writer.write_u16(self.refresh_interval_minutes)?;
        writer.write_u16(self.html_format)?;
        writer.write_u32(self.reconnection_method)?;
        writer.write_u8(self.credential_method)?;
        writer.write_u8(self.reserved3)?;
        self.source_data_file.write_to(writer)?;
        self.source_connection_file.write_to(writer)?;
        self.connection_name.write_to(writer)?;
        self.connection_description.write_to(writer)?;
        self.sso_application_id.write_to(writer)?;
        if let Some(value) = &self.table_names {
            value.write_to(writer)?;
        }
        for value in &self.parameters {
            value.write_to(writer)?;
        }
        match &self.connection {
            DConnConnection::Odbc(value) => value.write_to(writer)?,
            DConnConnection::Web(value) => value.write_to(writer)?,
            DConnConnection::OleDb(value) => value.write_to(writer)?,
            DConnConnection::Text(value) => value.write_to(writer)?,
            DConnConnection::None => {}
        }
        self.sql.write_to(writer)?;
        self.saved_sql.write_to(writer)?;
        self.edit_web_page.write_to(writer)?;
        self.identifier.write_to(writer)
    }
}

impl SdkRead for QsiSxTagRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeaderOld::read_from(reader)?;
        let pivot_table = DConBoolean::read_from(reader)?;
        let flags = QsiSxTagFlags::from_bits_retain(reader.read_u16()?);
        let raw_future_flags = reader.read_u32()?;
        let future_flags = match pivot_table {
            DConBoolean::False => QsiSxTagFutureFlags::QueryTable(
                QueryTableFutureFlags::from_bits_retain(raw_future_flags),
            ),
            DConBoolean::True => QsiSxTagFutureFlags::PivotTable(
                PivotTableFutureFlags::from_bits_retain(raw_future_flags),
            ),
        };
        Ok(Self {
            header,
            pivot_table,
            flags,
            future_flags,
            last_updated_version: reader.read_u8()?,
            minimum_updatable_version: reader.read_u8()?,
            name_character_offset: reader.read_u8()?,
            reserved: reader.read_u8()?,
            name: XlUnicodeString::read_from(reader)?,
            unused: reader.read_u16()?,
        })
    }
}

impl SdkWrite for QsiSxTagRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let future_flags = match (self.pivot_table, self.future_flags) {
            (DConBoolean::False, QsiSxTagFutureFlags::QueryTable(value)) => value.bits(),
            (DConBoolean::True, QsiSxTagFutureFlags::PivotTable(value)) => value.bits(),
            _ => {
                return Err(Error::invalid(0, "QsiSXTag future flags do not match fSx"));
            }
        };
        self.header.write_to(writer)?;
        self.pivot_table.write_to(writer)?;
        writer.write_u16(self.flags.bits())?;
        writer.write_u32(future_flags)?;
        writer.write_u8(self.last_updated_version)?;
        writer.write_u8(self.minimum_updatable_version)?;
        writer.write_u8(self.name_character_offset)?;
        writer.write_u8(self.reserved)?;
        self.name.write_to(writer)?;
        writer.write_u16(self.unused)?;
        Ok(())
    }
}

impl SdkRead for DbQueryExtRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeaderOld::read_from(reader)?;
        let data_source_type = DataSourceType::read_from(reader)?;
        let connection_flags = DbQueryExtConnectionFlags::from_bits_retain(reader.read_u16()?);
        let query_flags = DConnQueryFlags::from_raw(data_source_type, reader.read_u16()?);
        let flags = DbQueryExtFlags::from_bits_retain(reader.read_u16()?);
        let edited_version = reader.read_u8()?;
        let refreshed_version = reader.read_u8()?;
        let minimum_refreshable_version = reader.read_u8()?;
        let reserved4 = reader.read_u8()?;
        let reserved5 = reader.read_u16()?;
        let ole_db_connection_count = reader.read_u16()?;
        let future_byte_count = reader.read_u16()?;
        let refresh_interval_minutes = reader.read_u16()?;
        let html_format = reader.read_u16()?;
        let parameter_count = reader.read_u16()?;
        let count = usize::from(parameter_count);
        let maximum_parameters = usize::try_from(reader.remaining()? / 2)
            .map_err(|_| Error::Limit("DBQueryExt parameter bounds exceed usize".into()))?;
        if count > maximum_parameters {
            return Err(Error::invalid(
                reader.position()?,
                "DBQueryExt parameter count exceeds remaining bytes",
            ));
        }
        let mut parameters = Vec::with_capacity(count);
        for _ in 0..count {
            parameters.push(DbQueryParameterFlags::from_bits(reader.read_u16()?));
        }
        let future_bytes = reader.read_vec(usize::from(future_byte_count))?;
        Ok(Self {
            header,
            data_source_type,
            connection_flags,
            query_flags,
            flags,
            edited_version,
            refreshed_version,
            minimum_refreshable_version,
            reserved4,
            reserved5,
            ole_db_connection_count,
            future_byte_count,
            refresh_interval_minutes,
            html_format,
            parameter_count,
            parameters,
            future_bytes,
        })
    }
}

impl SdkWrite for DbQueryExtRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if usize::from(self.parameter_count) != self.parameters.len() {
            return Err(Error::invalid(0, "DBQueryExt parameter count mismatch"));
        }
        if usize::from(self.future_byte_count) != self.future_bytes.len() {
            return Err(Error::invalid(0, "DBQueryExt future-byte count mismatch"));
        }
        self.header.write_to(writer)?;
        self.data_source_type.write_to(writer)?;
        writer.write_u16(self.connection_flags.bits())?;
        writer.write_u16(self.query_flags.bits_for(self.data_source_type)?)?;
        writer.write_u16(self.flags.bits())?;
        writer.write_u8(self.edited_version)?;
        writer.write_u8(self.refreshed_version)?;
        writer.write_u8(self.minimum_refreshable_version)?;
        writer.write_u8(self.reserved4)?;
        writer.write_u16(self.reserved5)?;
        writer.write_u16(self.ole_db_connection_count)?;
        writer.write_u16(self.future_byte_count)?;
        writer.write_u16(self.refresh_interval_minutes)?;
        writer.write_u16(self.html_format)?;
        writer.write_u16(self.parameter_count)?;
        for value in &self.parameters {
            writer.write_u16(value.bits())?;
        }
        writer.write_all(&self.future_bytes)?;
        Ok(())
    }
}

impl SdkRead for HyperlinkTooltipRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtRefHeaderNoGrbit::read_from(reader)?;
        let remaining = reader.remaining()?;
        if remaining % 2 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "HLinkTooltip UTF-16 payload has an odd byte count",
            ));
        }
        let count = usize::try_from(remaining / 2)
            .map_err(|_| Error::Limit("HLinkTooltip length exceeds usize".into()))?;
        let mut tooltip = Vec::with_capacity(count);
        for _ in 0..count {
            tooltip.push(reader.read_u16()?);
        }
        Ok(Self { header, tooltip })
    }
}

impl SdkWrite for HyperlinkTooltipRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        for value in &self.tooltip {
            writer.write_u16(*value)?;
        }
        Ok(())
    }
}

impl SdkRead for SxAddlString {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        Ok(Self {
            total_character_count: reader.read_u32()?,
            reserved: reader.read_u16()?,
            segment: XlUnicodeString::read_from(reader)?,
        })
    }
}

impl SdkWrite for SxAddlString {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u32(self.total_character_count)?;
        writer.write_u16(self.reserved)?;
        self.segment.write_to(writer)
    }
}

fn read_sx_addl_reserved6<R: Read + Seek>(reader: &mut Reader<R>) -> Result<[u8; 6]> {
    let bytes = reader.read_vec(6)?;
    Ok([bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5]])
}

impl SdkRead for SxAddlRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = SxAddlHeader::read_from(reader)?;
        let data = match (header.class, header.data_type) {
            (0x00, 0x00) => SxAddlData::ViewId(SxAddlString::read_from(reader)?),
            (0x17, 0x00) => SxAddlData::Field12Id(SxAddlString::read_from(reader)?),
            (0x00 | 0x03 | 0x17, 0x01) => SxAddlData::VersionUpdateInvalidates {
                version: reader.read_u8()?,
                reserved1: reader.read_u8()?,
                reserved2: reader.read_u16()?,
                reserved3: reader.read_u16()?,
            },
            (0x00, 0x02) => {
                let created_version = reader.read_u8()?;
                let bytes = reader.read_vec(3)?;
                let raw_flags = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], 0]);
                SxAddlData::ViewVersion10 {
                    created_version,
                    flags: SxAddlViewVer10Flags::from_bits_retain(raw_flags),
                    reserved: reader.read_u16()?,
                }
            }
            (0x00, 0x19) => SxAddlData::ViewVersion12 {
                flags: SxAddlViewVer12Flags::from_bits_retain(reader.read_u32()?),
                reserved: reader.read_u16()?,
            },
            (0x00, 0x1e) => SxAddlData::ViewTableStyle {
                reserved: read_sx_addl_reserved6(reader)?,
                flags: SxAddlTableStyleFlags::from_bits_retain(reader.read_u16()?),
                style_name: LpWideString::read_from(reader)?,
            },
            (0x03, 0x00) => SxAddlData::CacheId {
                cache_stream_id: reader.read_u32()?,
                reserved: reader.read_u16()?,
            },
            (0x03, 0x02) => SxAddlData::CacheVersion10 {
                reserved1: read_sx_addl_reserved6(reader)?,
                ghost_item_limit: reader.read_i32()?,
                last_refresh_version: reader.read_u8()?,
                minimum_refreshable_version: reader.read_u8()?,
                refresh_date_bits: reader.read_u64()?,
                reserved2: reader.read_u16()?,
            },
            (0x03, 0x18) => SxAddlData::CacheVersionMacro {
                version: reader.read_u8()?,
                reserved1: reader.read_u8()?,
                reserved2: reader.read_u16()?,
                reserved3: reader.read_u16()?,
            },
            (0x03, 0x34) => SxAddlData::CacheInvalidRefresh {
                flags: SxAddlCacheRefreshFlags::from_bits_retain(reader.read_u32()?),
                reserved: reader.read_u16()?,
            },
            (0x03, 0x41) => SxAddlData::CacheInfo12 {
                flags: SxAddlCacheInfo12Flags::from_bits_retain(reader.read_u32()?),
                reserved: reader.read_u16()?,
            },
            (0x17, 0x19) => SxAddlData::Field12Version12 {
                flags: SxAddlField12Flags::from_bits_retain(reader.read_u32()?),
                reserved: reader.read_u16()?,
            },
            (_, 0xff) => SxAddlData::End {
                reserved: read_sx_addl_reserved6(reader)?,
            },
            (class, data_type) => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("unsupported SXAddl class/type 0x{class:02x}/0x{data_type:02x}"),
                ));
            }
        };
        Ok(Self { header, data })
    }
}

impl SdkWrite for SxAddlRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let matches_header = match &self.data {
            SxAddlData::End { .. } => self.header.data_type == 0xff,
            SxAddlData::ViewId(_) => (self.header.class, self.header.data_type) == (0x00, 0x00),
            SxAddlData::VersionUpdateInvalidates { .. } => {
                matches!(self.header.class, 0x00 | 0x03 | 0x17) && self.header.data_type == 0x01
            }
            SxAddlData::ViewVersion10 { .. } => {
                (self.header.class, self.header.data_type) == (0x00, 0x02)
            }
            SxAddlData::ViewVersion12 { .. } => {
                (self.header.class, self.header.data_type) == (0x00, 0x19)
            }
            SxAddlData::ViewTableStyle { .. } => {
                (self.header.class, self.header.data_type) == (0x00, 0x1e)
            }
            SxAddlData::CacheId { .. } => {
                (self.header.class, self.header.data_type) == (0x03, 0x00)
            }
            SxAddlData::CacheVersion10 { .. } => {
                (self.header.class, self.header.data_type) == (0x03, 0x02)
            }
            SxAddlData::CacheVersionMacro { .. } => {
                (self.header.class, self.header.data_type) == (0x03, 0x18)
            }
            SxAddlData::CacheInvalidRefresh { .. } => {
                (self.header.class, self.header.data_type) == (0x03, 0x34)
            }
            SxAddlData::CacheInfo12 { .. } => {
                (self.header.class, self.header.data_type) == (0x03, 0x41)
            }
            SxAddlData::Field12Id(_) => (self.header.class, self.header.data_type) == (0x17, 0x00),
            SxAddlData::Field12Version12 { .. } => {
                (self.header.class, self.header.data_type) == (0x17, 0x19)
            }
        };
        if !matches_header {
            return Err(Error::invalid(
                0,
                "SXAddl header does not match its typed data",
            ));
        }
        self.header.write_to(writer)?;
        match &self.data {
            SxAddlData::End { reserved } => writer.write_all(reserved)?,
            SxAddlData::ViewId(value) | SxAddlData::Field12Id(value) => {
                value.write_to(writer)?;
            }
            SxAddlData::VersionUpdateInvalidates {
                version,
                reserved1,
                reserved2,
                reserved3,
            } => {
                writer.write_u8(*version)?;
                writer.write_u8(*reserved1)?;
                writer.write_u16(*reserved2)?;
                writer.write_u16(*reserved3)?;
            }
            SxAddlData::ViewVersion10 {
                created_version,
                flags,
                reserved,
            } => {
                writer.write_u8(*created_version)?;
                let bytes = flags.bits().to_le_bytes();
                writer.write_all(&bytes[..3])?;
                writer.write_u16(*reserved)?;
            }
            SxAddlData::ViewVersion12 { flags, reserved } => {
                writer.write_u32(flags.bits())?;
                writer.write_u16(*reserved)?;
            }
            SxAddlData::ViewTableStyle {
                reserved,
                flags,
                style_name,
            } => {
                writer.write_all(reserved)?;
                writer.write_u16(flags.bits())?;
                style_name.write_to(writer)?;
            }
            SxAddlData::CacheId {
                cache_stream_id,
                reserved,
            } => {
                writer.write_u32(*cache_stream_id)?;
                writer.write_u16(*reserved)?;
            }
            SxAddlData::CacheVersion10 {
                reserved1,
                ghost_item_limit,
                last_refresh_version,
                minimum_refreshable_version,
                refresh_date_bits,
                reserved2,
            } => {
                writer.write_all(reserved1)?;
                writer.write_i32(*ghost_item_limit)?;
                writer.write_u8(*last_refresh_version)?;
                writer.write_u8(*minimum_refreshable_version)?;
                writer.write_u64(*refresh_date_bits)?;
                writer.write_u16(*reserved2)?;
            }
            SxAddlData::CacheVersionMacro {
                version,
                reserved1,
                reserved2,
                reserved3,
            } => {
                writer.write_u8(*version)?;
                writer.write_u8(*reserved1)?;
                writer.write_u16(*reserved2)?;
                writer.write_u16(*reserved3)?;
            }
            SxAddlData::CacheInvalidRefresh { flags, reserved } => {
                writer.write_u32(flags.bits())?;
                writer.write_u16(*reserved)?;
            }
            SxAddlData::CacheInfo12 { flags, reserved } => {
                writer.write_u32(flags.bits())?;
                writer.write_u16(*reserved)?;
            }
            SxAddlData::Field12Version12 { flags, reserved } => {
                writer.write_u32(flags.bits())?;
                writer.write_u16(*reserved)?;
            }
        }
        Ok(())
    }
}

impl SdkRead for BookExtRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let declared_size = reader.read_u32()?;
        let flags = BookExtFlags::from_bits_retain(reader.read_u32()?);
        let conditional11 = if reader.remaining()? >= 1 {
            Some(BookExtConditional11Flags::from_bits_retain(
                reader.read_u8()?,
            ))
        } else {
            None
        };
        let conditional12 = if reader.remaining()? >= 1 {
            Some(BookExtConditional12Flags::from_bits_retain(
                reader.read_u8()?,
            ))
        } else {
            None
        };
        Ok(Self {
            header,
            declared_size,
            flags,
            conditional11,
            conditional12,
        })
    }
}

impl SdkWrite for BookExtRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.conditional12.is_some() && self.conditional11.is_none() {
            return Err(Error::invalid(
                0,
                "BookExt conditional12 requires conditional11",
            ));
        }
        self.header.write_to(writer)?;
        writer.write_u32(self.declared_size)?;
        writer.write_u32(self.flags.bits())?;
        if let Some(value) = self.conditional11 {
            writer.write_u8(value.bits())?;
        }
        if let Some(value) = self.conditional12 {
            writer.write_u8(value.bits())?;
        }
        Ok(())
    }
}

impl SdkRead for XmlTkRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let data_type = reader.read_u8()?;
        let unused = reader.read_u8()?;
        let tag = reader.read_u16()?;
        let data = match data_type {
            0x00 => XmlTkData::Start,
            0x01 => XmlTkData::End,
            0x02 => XmlTkData::Boolean {
                value: reader.read_u8()?,
                unused: reader.read_u8()?,
            },
            0x03 => XmlTkData::Double {
                unused: reader.read_u32()?,
                value_bits: reader.read_u64()?,
            },
            0x04 => XmlTkData::DWord(reader.read_i32()?),
            0x05 => {
                let count = usize::try_from(reader.read_u32()?)
                    .map_err(|_| Error::Limit("XmlTkString count exceeds usize".into()))?;
                let maximum = usize::try_from(reader.remaining()? / 2)
                    .map_err(|_| Error::Limit("XmlTkString length exceeds usize".into()))?;
                if count > maximum {
                    return Err(Error::invalid(
                        reader.position()?,
                        "XmlTkString count exceeds remaining bytes",
                    ));
                }
                let mut values = Vec::with_capacity(count);
                for _ in 0..count {
                    values.push(reader.read_u16()?);
                }
                XmlTkData::String(values)
            }
            0x06 => XmlTkData::Token(reader.read_u16()?),
            0x07 => {
                let count = usize::try_from(reader.read_u32()?)
                    .map_err(|_| Error::Limit("XmlTkBlob count exceeds usize".into()))?;
                XmlTkData::Blob(reader.read_vec(count)?)
            }
            value => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("unknown XmlTk data type 0x{value:02x}"),
                ));
            }
        };
        Ok(Self { unused, tag, data })
    }
}

impl SdkWrite for XmlTkRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let data_type = match &self.data {
            XmlTkData::Start => 0x00,
            XmlTkData::End => 0x01,
            XmlTkData::Boolean { .. } => 0x02,
            XmlTkData::Double { .. } => 0x03,
            XmlTkData::DWord(_) => 0x04,
            XmlTkData::String(_) => 0x05,
            XmlTkData::Token(_) => 0x06,
            XmlTkData::Blob(_) => 0x07,
        };
        writer.write_u8(data_type)?;
        writer.write_u8(self.unused)?;
        writer.write_u16(self.tag)?;
        match &self.data {
            XmlTkData::Start | XmlTkData::End => {}
            XmlTkData::Boolean { value, unused } => {
                writer.write_u8(*value)?;
                writer.write_u8(*unused)?;
            }
            XmlTkData::Double { unused, value_bits } => {
                writer.write_u32(*unused)?;
                writer.write_u64(*value_bits)?;
            }
            XmlTkData::DWord(value) => writer.write_i32(*value)?,
            XmlTkData::String(values) => {
                writer.write_u32(
                    u32::try_from(values.len())
                        .map_err(|_| Error::Limit("XmlTkString count exceeds u32".into()))?,
                )?;
                for value in values {
                    writer.write_u16(*value)?;
                }
            }
            XmlTkData::Token(value) => writer.write_u16(*value)?,
            XmlTkData::Blob(values) => {
                writer.write_u32(
                    u32::try_from(values.len())
                        .map_err(|_| Error::Limit("XmlTkBlob count exceeds u32".into()))?,
                )?;
                writer.write_all(values)?;
            }
        }
        Ok(())
    }
}

impl SdkRead for XmlTkChain {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let record_version = reader.read_u8()?;
        let unused = reader.read_u8()?;
        let parent = reader.read_u16()?;
        let mut records = Vec::new();
        while reader.remaining()? != 0 {
            records.push(XmlTkRecord::read_from(reader)?);
        }
        Ok(Self {
            record_version,
            unused,
            parent,
            records,
        })
    }
}

impl SdkWrite for XmlTkChain {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u8(self.record_version)?;
        writer.write_u8(self.unused)?;
        writer.write_u16(self.parent)?;
        for record in &self.records {
            record.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for CrtMlFrtRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let chain_length = reader.read_u32()?;
        let mut child = reader.sub_reader(u64::from(chain_length))?;
        let chain = XmlTkChain::read_from(&mut child)?;
        let unused = reader.read_u32()?;
        Ok(Self {
            header,
            chain,
            unused,
        })
    }
}

impl SdkWrite for CrtMlFrtRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut child = Writer::new(Cursor::new(Vec::new()));
        self.chain.write_to(&mut child)?;
        let bytes = child.into_inner().into_inner();
        self.header.write_to(writer)?;
        writer.write_u32(
            u32::try_from(bytes.len())
                .map_err(|_| Error::Limit("XmlTkChain exceeds u32 bytes".into()))?,
        )?;
        writer.write_all(&bytes)?;
        writer.write_u32(self.unused)?;
        Ok(())
    }
}

impl SdkRead for ExtSstRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let strings_per_bucket = reader.read_u16()?;
        let remaining = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("ExtSST bucket bytes exceed usize".into()))?;
        if !remaining.is_multiple_of(8) {
            return Err(Error::invalid(
                reader.position()?,
                "ExtSST bucket array is not 8-byte aligned",
            ));
        }
        let count = remaining / 8;
        reader.ensure_allocation(count, 8)?;
        let mut buckets = Vec::with_capacity(count);
        for _ in 0..count {
            buckets.push(ExtSstBucket::read_from(reader)?);
        }
        Ok(Self {
            strings_per_bucket,
            buckets,
        })
    }
}

impl SdkWrite for ExtSstRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.strings_per_bucket)?;
        for bucket in &self.buckets {
            bucket.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkWrite for ExtendedHeaderFooterRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        writer.write_all(&self.sheet_view_guid)?;
        writer.write_u16(self.flags.bits())?;
        writer.write_u16(self.even_header_character_count)?;
        writer.write_u16(self.even_footer_character_count)?;
        writer.write_u16(self.first_header_character_count)?;
        writer.write_u16(self.first_footer_character_count)?;
        let fields = [
            (self.even_header_character_count, &self.even_header),
            (self.even_footer_character_count, &self.even_footer),
            (self.first_header_character_count, &self.first_header),
            (self.first_footer_character_count, &self.first_footer),
        ];
        for (count, value) in fields {
            if (count != 0) != value.is_some() {
                return Err(Error::invalid(
                    0,
                    "HeaderFooter character count and string presence disagree",
                ));
            }
            if let Some(value) = value {
                if value.character_count() != usize::from(count) {
                    return Err(Error::invalid(0, "HeaderFooter string length mismatch"));
                }
                value.write(writer)?;
            }
        }
        Ok(())
    }
}

impl SdkWrite for XfExtRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        writer.write_u16(self.reserved1)?;
        writer.write_u16(self.xf_index)?;
        writer.write_u16(self.reserved2)?;
        writer.write_u16(
            u16::try_from(self.properties.len())
                .map_err(|_| Error::Limit("XFExt property count exceeds u16".into()))?,
        )?;
        for property in &self.properties {
            property.write(writer)?;
        }
        Ok(())
    }
}

impl ExtProperty {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let property_type = reader.read_u16()?;
        let size = usize::from(reader.read_u16()?);
        let position = reader.position()?;
        let payload_len = size
            .checked_sub(4)
            .ok_or_else(|| Error::invalid(position, "ExtProp size is smaller than header"))?;
        let payload = reader.read_vec(payload_len)?;
        let data = match property_type {
            0x0004 | 0x0005 | 0x0007..=0x000b | 0x000d | 0x0048 => {
                Self::decode_full_color(property_type, &payload)?
            }
            0x0006 => ExtPropertyData::Gradient { payload },
            0x000e if payload.len() == 1 => {
                ExtPropertyData::FontScheme(ExtFontScheme::Byte(payload[0]))
            }
            0x000e if payload.len() == 2 => {
                ExtPropertyData::FontScheme(ExtFontScheme::Word(u16::from_le_bytes([
                    payload[0], payload[1],
                ])))
            }
            0x000f if payload.len() == 2 => {
                ExtPropertyData::Indentation(u16::from_le_bytes([payload[0], payload[1]]))
            }
            _ => ExtPropertyData::Unknown {
                property_type,
                payload,
            },
        };
        Ok(Self { data })
    }

    fn decode_full_color(property_type: u16, payload: &[u8]) -> Result<ExtPropertyData> {
        let color = parse_sdk(payload, 0, property_type)?;
        Ok(ExtPropertyData::FullColor {
            property_type,
            color,
        })
    }

    fn write<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let (property_type, payload) = match &self.data {
            ExtPropertyData::FullColor {
                property_type,
                color,
            } => (*property_type, encode_sdk(color)?),
            ExtPropertyData::Gradient { payload } => (0x0006, payload.clone()),
            ExtPropertyData::FontScheme(ExtFontScheme::Byte(value)) => (0x000e, vec![*value]),
            ExtPropertyData::FontScheme(ExtFontScheme::Word(value)) => {
                (0x000e, value.to_le_bytes().to_vec())
            }
            ExtPropertyData::Indentation(value) => (0x000f, value.to_le_bytes().to_vec()),
            ExtPropertyData::Unknown {
                property_type,
                payload,
            } => (*property_type, payload.clone()),
        };
        let size = payload
            .len()
            .checked_add(4)
            .ok_or_else(|| Error::Limit("ExtProp size overflow".into()))?;
        writer.write_u16(property_type)?;
        writer.write_u16(
            u16::try_from(size).map_err(|_| Error::Limit("ExtProp size exceeds u16".into()))?,
        )?;
        writer.write_all(&payload)?;
        Ok(())
    }
}

impl SdkRead for XfProperty {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let property_type = reader.read_u16()?;
        let size = usize::from(reader.read_u16()?);
        let position = reader.position()?;
        let payload_len = size
            .checked_sub(4)
            .ok_or_else(|| Error::invalid(position, "XFProp size is smaller than header"))?;
        let payload = reader.read_vec(payload_len)?;
        let data = match property_type {
            0x0001 | 0x0002 | 0x0005 if payload.len() == 8 => {
                XfPropertyData::Color(parse_sdk(&payload, 0, property_type)?)
            }
            0x0006..=0x000c if payload.len() == 10 => {
                XfPropertyData::Border(parse_sdk(&payload, 0, property_type)?)
            }
            0x0000 | 0x000d..=0x0011 | 0x0013..=0x0017 | 0x001c..=0x0023 if payload.len() == 1 => {
                XfPropertyData::Byte(payload[0])
            }
            0x0012 | 0x0019 | 0x001a | 0x001b | 0x0029 | 0x002a if payload.len() == 2 => {
                XfPropertyData::Word(u16::from_le_bytes([payload[0], payload[1]]))
            }
            0x0024 if payload.len() == 4 => XfPropertyData::DWord(u32::from_le_bytes(
                payload.as_slice().try_into().expect("four-byte XFProp"),
            )),
            0x0018 | 0x0026 => XfPropertyData::WideString(parse_sdk(&payload, 0, property_type)?),
            0x0025 if payload.len() == 1 => XfPropertyData::FontScheme(payload[0]),
            _ => XfPropertyData::Unparsed(payload),
        };
        Ok(Self {
            property_type,
            data,
        })
    }
}

impl SdkWrite for XfProperty {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let payload = match &self.data {
            XfPropertyData::Color(value) => encode_sdk(value)?,
            XfPropertyData::Border(value) => encode_sdk(value)?,
            XfPropertyData::Byte(value) => vec![*value],
            XfPropertyData::Word(value) => value.to_le_bytes().to_vec(),
            XfPropertyData::DWord(value) => value.to_le_bytes().to_vec(),
            XfPropertyData::WideString(value) => encode_sdk(value)?,
            XfPropertyData::FontScheme(value) => vec![*value],
            XfPropertyData::Unparsed(payload) => payload.clone(),
        };
        let size = payload
            .len()
            .checked_add(4)
            .ok_or_else(|| Error::Limit("XFProp size overflow".into()))?;
        writer.write_u16(self.property_type)?;
        writer.write_u16(
            u16::try_from(size).map_err(|_| Error::Limit("XFProp size exceeds u16".into()))?,
        )?;
        writer.write_all(&payload)?;
        Ok(())
    }
}

impl SdkSize for XfProperty {
    fn sdk_size(&self) -> u64 {
        4 + match &self.data {
            XfPropertyData::Color(_) => 8,
            XfPropertyData::Border(_) => 10,
            XfPropertyData::Byte(_) => 1,
            XfPropertyData::Word(_) => 2,
            XfPropertyData::DWord(_) => 4,
            XfPropertyData::WideString(value) => value.sdk_size(),
            XfPropertyData::FontScheme(_) => 1,
            XfPropertyData::Unparsed(payload) => payload.len() as u64,
        }
    }
}

impl SdkRead for DbCellRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let row_offset = reader.read_u32()?;
        let remaining = reader.remaining()?;
        if remaining % 2 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "DBCell offset array has an odd byte length",
            ));
        }
        let count = usize::try_from(remaining / 2)
            .map_err(|_| Error::Limit("DBCell offset count exceeds usize".into()))?;
        reader.ensure_allocation(count, 2)?;
        let mut cell_offsets = Vec::with_capacity(count);
        for _ in 0..count {
            cell_offsets.push(reader.read_u16()?);
        }
        Ok(Self {
            row_offset,
            cell_offsets,
        })
    }
}

impl SdkWrite for DbCellRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u32(self.row_offset)?;
        for offset in &self.cell_offsets {
            writer.write_u16(*offset)?;
        }
        Ok(())
    }
}

impl BiffUnicodeString {
    fn read<R: Read + Seek>(reader: &mut Reader<R>, count: usize) -> Result<Self> {
        let flags = reader.read_u8()?;
        let byte_count = if flags & 1 == 0 {
            count
        } else {
            count
                .checked_mul(2)
                .ok_or_else(|| Error::Limit("BIFF string byte count overflow".into()))?
        };
        let remaining = reader.remaining()?;
        if remaining < byte_count as u64 {
            return Err(Error::invalid(
                reader.position()?,
                format!(
                    "BIFF string flags 0x{flags:02x} declare {count} characters requiring {byte_count} bytes but only {remaining} remain"
                ),
            ));
        }
        let characters = if flags & 1 == 0 {
            XlStringCharacters::Compressed(reader.read_vec(count)?)
        } else {
            let bytes = reader.read_vec(byte_count)?;
            XlStringCharacters::Unicode(
                bytes
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect(),
            )
        };
        Ok(Self {
            flags,
            characters,
            trailing_byte: None,
        })
    }

    fn read_remaining<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let flags = reader.read_u8()?;
        let remaining = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("BIFF string remaining length exceeds usize".into()))?;
        let bytes = reader.read_vec(remaining)?;
        let (characters, trailing_byte) = if flags & 1 == 0 {
            (XlStringCharacters::Compressed(bytes), None)
        } else {
            let pairs = bytes.chunks_exact(2);
            let remainder = pairs.remainder().first().copied();
            (
                XlStringCharacters::Unicode(
                    pairs
                        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                        .collect(),
                ),
                remainder,
            )
        };
        Ok(Self {
            flags,
            characters,
            trailing_byte,
        })
    }

    pub fn character_count(&self) -> usize {
        match &self.characters {
            XlStringCharacters::Compressed(values) => values.len(),
            XlStringCharacters::Unicode(values) => values.len(),
        }
    }

    fn write<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u8(self.flags)?;
        self.write_characters(writer)
    }

    fn write_characters<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        match &self.characters {
            XlStringCharacters::Compressed(values) => {
                if self.flags & 1 != 0 {
                    return Err(Error::invalid(
                        writer.position()?,
                        "compressed BIFF string has UTF-16 flag",
                    ));
                }
                writer.write_all(values)?;
            }
            XlStringCharacters::Unicode(values) => {
                if self.flags & 1 == 0 {
                    return Err(Error::invalid(
                        writer.position()?,
                        "Unicode BIFF string lacks UTF-16 flag",
                    ));
                }
                for value in values {
                    writer.write_u16(*value)?;
                }
            }
        }
        if let Some(value) = self.trailing_byte {
            writer.write_u8(value)?;
        }
        Ok(())
    }
}

impl SdkRead for SxviRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let item_type = reader.read_u16()? as i16;
        let flags = SxviFlags::from_bits_retain(reader.read_u16()?);
        let cache_index = reader.read_u16()? as i16;
        let declared_name_length = reader.read_u16()?;
        let name = if declared_name_length == u16::MAX {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::from(declared_name_length),
            )?)
        };
        Ok(Self {
            item_type,
            flags,
            cache_index,
            declared_name_length,
            name,
        })
    }
}

impl SdkWrite for SxviRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_name_length = match &self.name {
            None => u16::MAX,
            Some(name) => u16::try_from(name.character_count())
                .map_err(|_| Error::Limit("SXVI name exceeds u16 characters".into()))?,
        };
        if self.declared_name_length != expected_name_length {
            return Err(Error::invalid(
                writer.position()?,
                "SXVI cchName does not match its static name",
            ));
        }
        writer.write_u16(self.item_type as u16)?;
        writer.write_u16(self.flags.bits())?;
        writer.write_u16(self.cache_index as u16)?;
        writer.write_u16(self.declared_name_length)?;
        if let Some(name) = &self.name {
            name.write(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for SxIvdRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let remaining = reader.remaining()?;
        if remaining % 2 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "SxIvd field-index array has an odd byte length",
            ));
        }
        let count = usize::try_from(remaining / 2)
            .map_err(|_| Error::Limit("SxIvd field count exceeds usize".into()))?;
        reader.ensure_allocation(count, 2)?;
        let mut field_indices = Vec::with_capacity(count);
        for _ in 0..count {
            field_indices.push(reader.read_i16()?);
        }
        Ok(Self { field_indices })
    }
}

impl SdkWrite for SxIvdRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        for field_index in &self.field_indices {
            writer.write_u16(*field_index as u16)?;
        }
        Ok(())
    }
}

impl SxLiRecord {
    fn read_with_axis_dimension<R: Read + Seek>(
        reader: &mut Reader<R>,
        axis_dimension_count: u16,
    ) -> Result<Self> {
        let mut items = Vec::new();
        while reader.remaining()? != 0 {
            if reader.remaining()? < 8 {
                return Err(Error::invalid(
                    reader.position()?,
                    "SXLI has a truncated item header",
                ));
            }
            let shared_prefix_count = reader.read_i16()?;
            let item_type_and_reserved = reader.read_u16()?;
            let item_type_raw = item_type_and_reserved & 0x7fff;
            let item_type = SxLineItemType::from_raw(item_type_raw).ok_or_else(|| {
                Error::invalid(
                    reader
                        .position()
                        .unwrap_or(reader.start())
                        .saturating_sub(2),
                    format!("unknown SXLI item type 0x{item_type_raw:04x}"),
                )
            })?;
            let reserved1 = item_type_and_reserved & 0x8000 != 0;
            let displayed_item_count = reader.read_i16()?;
            let packed_flags = reader.read_u16()?;
            let flags = SxLineFlags::from_bits_retain(packed_flags & !0x01fe);
            let data_item_index = ((packed_flags >> 1) & 0x00ff) as u8;
            let count = usize::from(axis_dimension_count);
            reader.ensure_allocation(count, 2)?;
            let mut item_indices = Vec::with_capacity(count);
            for _ in 0..count {
                item_indices.push(reader.read_i16()?);
            }
            items.push(SxLiItem {
                shared_prefix_count,
                item_type,
                reserved1,
                displayed_item_count,
                flags,
                data_item_index,
                item_indices,
            });
        }
        Ok(Self {
            axis_dimension_count,
            items,
        })
    }
}

impl SdkWrite for SxLiRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        for item in &self.items {
            if item.item_indices.len() != usize::from(self.axis_dimension_count) {
                return Err(Error::invalid(
                    writer.position()?,
                    "SXLI item index array does not match its axis dimension",
                ));
            }
            if item.flags.bits() & 0x01fe != 0 {
                return Err(Error::invalid(
                    writer.position()?,
                    "SXLI flags overlap the static data-item index field",
                ));
            }
            writer.write_u16(item.shared_prefix_count as u16)?;
            writer.write_u16(item.item_type.raw() | if item.reserved1 { 0x8000 } else { 0 })?;
            writer.write_u16(item.displayed_item_count as u16)?;
            writer.write_u16(item.flags.bits() | (u16::from(item.data_item_index) << 1))?;
            for item_index in &item.item_indices {
                writer.write_u16(*item_index as u16)?;
            }
        }
        Ok(())
    }
}

impl SdkRead for SxPiRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let remaining = reader.remaining()?;
        if remaining % 6 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "SXPI item array is not a multiple of six bytes",
            ));
        }
        let count = usize::try_from(remaining / 6)
            .map_err(|_| Error::Limit("SXPI item count exceeds usize".into()))?;
        reader.ensure_allocation(count, 6)?;
        let mut items = Vec::with_capacity(count);
        for _ in 0..count {
            items.push(SxPiItem::read_from(reader)?);
        }
        Ok(Self { items })
    }
}

impl SdkWrite for SxPiRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        for item in &self.items {
            item.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for SxDiRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let data_field_index = reader.read_i16()?;
        let aggregation = SxDataAggregation::read_from(reader)?;
        let display_calculation = SxDataDisplayCalculation::read_from(reader)?;
        let calculation_field_index = reader.read_i16()?;
        let calculation_item_index = reader.read_i16()?;
        let number_format_index = reader.read_u16()?;
        let declared_name_length = reader.read_u16()?;
        let name = if declared_name_length == u16::MAX {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::from(declared_name_length),
            )?)
        };
        Ok(Self {
            data_field_index,
            aggregation,
            display_calculation,
            calculation_field_index,
            calculation_item_index,
            number_format_index,
            declared_name_length,
            name,
        })
    }
}

impl SdkWrite for SxDiRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_name_length = match &self.name {
            None => u16::MAX,
            Some(name) => u16::try_from(name.character_count())
                .map_err(|_| Error::Limit("SXDI name exceeds u16 characters".into()))?,
        };
        if self.declared_name_length != expected_name_length {
            return Err(Error::invalid(
                writer.position()?,
                "SXDI cchName does not match its static name",
            ));
        }
        writer.write_u16(self.data_field_index as u16)?;
        self.aggregation.write_to(writer)?;
        self.display_calculation.write_to(writer)?;
        writer.write_u16(self.calculation_field_index as u16)?;
        writer.write_u16(self.calculation_item_index as u16)?;
        writer.write_u16(self.number_format_index)?;
        writer.write_u16(self.declared_name_length)?;
        if let Some(name) = &self.name {
            name.write(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for SxStringRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let declared_character_count = reader.read_u16()?;
        let segment = if declared_character_count == u16::MAX {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::from(declared_character_count),
            )?)
        };
        Ok(Self {
            declared_character_count,
            segment,
        })
    }
}

impl SdkWrite for SxStringRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_count = match &self.segment {
            None => u16::MAX,
            Some(segment) => u16::try_from(segment.character_count())
                .map_err(|_| Error::Limit("SXString segment exceeds u16 characters".into()))?,
        };
        if self.declared_character_count != expected_count {
            return Err(Error::invalid(
                writer.position()?,
                "SXString cch does not match its static segment",
            ));
        }
        writer.write_u16(self.declared_character_count)?;
        if let Some(segment) = &self.segment {
            segment.write(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for RrTabIdRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let remaining = reader.remaining()?;
        if remaining % 2 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "RRTabId sheet identifier array has an odd byte length",
            ));
        }
        let count = usize::try_from(remaining / 2)
            .map_err(|_| Error::Limit("RRTabId sheet count exceeds usize".into()))?;
        reader.ensure_allocation(count, 2)?;
        let mut sheet_ids = Vec::with_capacity(count);
        for _ in 0..count {
            sheet_ids.push(reader.read_u16()?);
        }
        Ok(Self { sheet_ids })
    }
}

impl SdkWrite for RrTabIdRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        for sheet_id in &self.sheet_ids {
            writer.write_u16(*sheet_id)?;
        }
        Ok(())
    }
}

impl SdkRead for SxRuleRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let dimension = reader.read_u8()?;
        let field_index = reader.read_u8()?;
        let axis_and_type = reader.read_u8()?;
        let area_type = axis_and_type >> 4;
        if area_type > 6 {
            return Err(Error::invalid(
                reader.position()?.saturating_sub(1),
                "SxRule sxrType is outside the defined range",
            ));
        }
        let axis = SxRuleAxis::from_bits_retain(axis_and_type & 0x0f);
        let flags = SxRuleFlags::from_bits_retain(reader.read_u8()?);
        let reserved = reader.read_u16()?;
        let filter_count = reader.read_u16()?;
        let partial_range = if flags.contains(SxRuleFlags::PART) {
            Some(SxRulePartialRange::read_from(reader)?)
        } else {
            None
        };
        Ok(Self {
            dimension,
            field_index,
            axis,
            area_type,
            flags,
            reserved,
            filter_count,
            partial_range,
        })
    }
}

impl SdkWrite for SxRuleRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.area_type > 6 {
            return Err(Error::invalid(
                writer.position()?,
                "SxRule sxrType is outside the defined range",
            ));
        }
        if self.flags.contains(SxRuleFlags::PART) != self.partial_range.is_some() {
            return Err(Error::invalid(
                writer.position()?,
                "SxRule fPart does not match its static partial range",
            ));
        }
        writer.write_u8(self.dimension)?;
        writer.write_u8(self.field_index)?;
        writer.write_u8(self.axis.bits() | (self.area_type << 4))?;
        writer.write_u8(self.flags.bits())?;
        writer.write_u16(self.reserved)?;
        writer.write_u16(self.filter_count)?;
        if let Some(partial_range) = &self.partial_range {
            partial_range.write_to(writer)?;
        }
        Ok(())
    }
}

fn read_sx_ex_string<R: Read + Seek>(
    reader: &mut Reader<R>,
    declared_length: u16,
) -> Result<Option<BiffUnicodeString>> {
    if declared_length == u16::MAX {
        Ok(None)
    } else {
        Ok(Some(BiffUnicodeString::read(
            reader,
            usize::from(declared_length),
        )?))
    }
}

fn write_sx_ex_string<W: Write + Seek>(
    writer: &mut Writer<W>,
    declared_length: u16,
    value: &Option<BiffUnicodeString>,
) -> Result<()> {
    let expected_length = match value {
        None => u16::MAX,
        Some(value) => u16::try_from(value.character_count())
            .map_err(|_| Error::Limit("SXEx string exceeds u16 characters".into()))?,
    };
    if declared_length != expected_length {
        return Err(Error::invalid(
            writer.position()?,
            "SXEx string length does not match its static value",
        ));
    }
    if let Some(value) = value {
        value.write(writer)?;
    }
    Ok(())
}

impl SdkRead for SxExRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let format_count = reader.read_u16()?;
        let error_string_length = reader.read_u16()?;
        let null_string_length = reader.read_u16()?;
        let tag_length = reader.read_u16()?;
        let selection_count = reader.read_u16()?;
        let page_row_count = reader.read_u16()?;
        let page_column_count = reader.read_u16()?;
        let packed_page_layout = reader.read_u16()?;
        let page_layout_flags = SxExPageLayoutFlags::from_bits_retain(packed_page_layout & !0x01fe);
        let page_wrap_count = ((packed_page_layout >> 1) & 0xff) as u8;
        let flags = SxExFlags::from_bits_retain(reader.read_u8()?);
        let reserved3 = reader.read_u8()?;
        let page_field_style_length = reader.read_u16()?;
        let table_style_length = reader.read_u16()?;
        let vacate_style_length = reader.read_u16()?;
        let error_string = read_sx_ex_string(reader, error_string_length)?;
        let null_string = read_sx_ex_string(reader, null_string_length)?;
        let tag = read_sx_ex_string(reader, tag_length)?;
        let page_field_style = read_sx_ex_string(reader, page_field_style_length)?;
        let table_style = read_sx_ex_string(reader, table_style_length)?;
        let vacate_style = read_sx_ex_string(reader, vacate_style_length)?;
        Ok(Self {
            format_count,
            error_string_length,
            null_string_length,
            tag_length,
            selection_count,
            page_row_count,
            page_column_count,
            page_layout_flags,
            page_wrap_count,
            flags,
            reserved3,
            page_field_style_length,
            table_style_length,
            vacate_style_length,
            error_string,
            null_string,
            tag,
            page_field_style,
            table_style,
            vacate_style,
        })
    }
}

impl SdkWrite for SxExRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.format_count)?;
        writer.write_u16(self.error_string_length)?;
        writer.write_u16(self.null_string_length)?;
        writer.write_u16(self.tag_length)?;
        writer.write_u16(self.selection_count)?;
        writer.write_u16(self.page_row_count)?;
        writer.write_u16(self.page_column_count)?;
        writer.write_u16(self.page_layout_flags.bits() | (u16::from(self.page_wrap_count) << 1))?;
        writer.write_u8(self.flags.bits())?;
        writer.write_u8(self.reserved3)?;
        writer.write_u16(self.page_field_style_length)?;
        writer.write_u16(self.table_style_length)?;
        writer.write_u16(self.vacate_style_length)?;
        write_sx_ex_string(writer, self.error_string_length, &self.error_string)?;
        write_sx_ex_string(writer, self.null_string_length, &self.null_string)?;
        write_sx_ex_string(writer, self.tag_length, &self.tag)?;
        write_sx_ex_string(writer, self.page_field_style_length, &self.page_field_style)?;
        write_sx_ex_string(writer, self.table_style_length, &self.table_style)?;
        write_sx_ex_string(writer, self.vacate_style_length, &self.vacate_style)?;
        Ok(())
    }
}

fn sign_extend_10(value: u32) -> i16 {
    let value = (value & 0x03ff) as i16;
    if value & 0x0200 != 0 {
        value | !0x03ff
    } else {
        value
    }
}

fn encode_signed_10(value: i16, position: u64, field: &str) -> Result<u32> {
    if !(-512..=511).contains(&value) {
        return Err(Error::invalid(
            position,
            format!("SxFilt {field} does not fit its signed 10-bit field"),
        ));
    }
    Ok(u32::from(value as u16) & 0x03ff)
}

impl SdkRead for SxFiltRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let packed = reader.read_u32()?;
        Ok(Self {
            axis: SxRuleAxis::from_bits_retain((packed & 0x0f) as u8),
            reserved1: ((packed >> 4) & 0x03) as u8,
            dimension: sign_extend_10(packed >> 6),
            field_index: sign_extend_10(packed >> 16),
            selected: packed & (1 << 26) != 0,
            reserved2: packed & (1 << 27) != 0,
            reserved3: ((packed >> 28) & 0x0f) as u8,
            subtotal_flags: SxFiltSubtotalFlags::from_bits_retain(reader.read_u16()?),
            item_count: reader.read_u16()?,
        })
    }
}

impl SdkWrite for SxFiltRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.reserved1 > 0x03 || self.reserved3 > 0x0f {
            return Err(Error::invalid(
                writer.position()?,
                "SxFilt reserved bit field exceeds its static width",
            ));
        }
        let position = writer.position()?;
        let packed = u32::from(self.axis.bits())
            | (u32::from(self.reserved1) << 4)
            | (encode_signed_10(self.dimension, position, "iDim")? << 6)
            | (encode_signed_10(self.field_index, position, "isxvd")? << 16)
            | (u32::from(self.selected) << 26)
            | (u32::from(self.reserved2) << 27)
            | (u32::from(self.reserved3) << 28);
        writer.write_u32(packed)?;
        writer.write_u16(self.subtotal_flags.bits())?;
        writer.write_u16(self.item_count)?;
        Ok(())
    }
}

impl SdkRead for SxDxfRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let byte_count = u32::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("SxDXF payload exceeds u32 bytes".into()))?;
        Ok(Self {
            format: DxfN12List::read_sized(reader, byte_count)?,
        })
    }
}

impl SdkWrite for SxDxfRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_all(&self.format.to_bytes()?)?;
        Ok(())
    }
}

impl SdkRead for SxItmRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let remaining = reader.remaining()?;
        if remaining % 2 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "SxItm index array has an odd byte length",
            ));
        }
        let count = usize::try_from(remaining / 2)
            .map_err(|_| Error::Limit("SxItm item count exceeds usize".into()))?;
        reader.ensure_allocation(count, 2)?;
        let mut item_indices = Vec::with_capacity(count);
        for _ in 0..count {
            item_indices.push(reader.read_u16()?);
        }
        Ok(Self { item_indices })
    }
}

impl SdkWrite for SxItmRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        for item_index in &self.item_indices {
            writer.write_u16(*item_index)?;
        }
        Ok(())
    }
}

fn dcon_file_is_self_reference(file: &BiffUnicodeString) -> bool {
    match &file.characters {
        XlStringCharacters::Compressed(characters) => characters.first().copied() == Some(0x02),
        XlStringCharacters::Unicode(characters) => characters.first().copied() == Some(0x0002),
    }
}

impl SdkRead for DConRefRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let reference = RefU::read_from(reader)?;
        let declared_file_character_count = reader.read_u16()?;
        let file = BiffUnicodeString::read(reader, usize::from(declared_file_character_count))?;
        let self_reference_unused = if dcon_file_is_self_reference(&file) {
            if reader.remaining()? == 0 {
                Some(DConSelfReferenceUnused::Missing)
            } else if file.flags & 1 == 0 {
                Some(DConSelfReferenceUnused::Compressed(reader.read_u8()?))
            } else {
                Some(DConSelfReferenceUnused::Unicode(reader.read_u16()?))
            }
        } else {
            None
        };
        Ok(Self {
            reference,
            declared_file_character_count,
            file,
            self_reference_unused,
        })
    }
}

impl SdkWrite for DConRefRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_count = u16::try_from(self.file.character_count())
            .map_err(|_| Error::Limit("DConRef file exceeds u16 characters".into()))?;
        if self.declared_file_character_count != expected_count {
            return Err(Error::invalid(
                writer.position()?,
                "DConRef cchFile does not match its static file string",
            ));
        }
        let expected_self_reference = dcon_file_is_self_reference(&self.file);
        let padding_matches_encoding = matches!(
            (self.file.flags & 1, self.self_reference_unused),
            (0, Some(DConSelfReferenceUnused::Compressed(_)))
                | (1, Some(DConSelfReferenceUnused::Unicode(_)))
                | (_, Some(DConSelfReferenceUnused::Missing))
        );
        if expected_self_reference != self.self_reference_unused.is_some()
            || (expected_self_reference && !padding_matches_encoding)
        {
            return Err(Error::invalid(
                writer.position()?,
                "DConRef self-reference padding does not match its static file string",
            ));
        }
        self.reference.write_to(writer)?;
        writer.write_u16(self.declared_file_character_count)?;
        self.file.write(writer)?;
        match self.self_reference_unused {
            Some(DConSelfReferenceUnused::Compressed(value)) => writer.write_u8(value)?,
            Some(DConSelfReferenceUnused::Unicode(value)) => writer.write_u16(value)?,
            Some(DConSelfReferenceUnused::Missing) => {}
            None => {}
        }
        Ok(())
    }
}

impl SdkRead for FileSharingRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let read_only_recommended = DConBoolean::read_from(reader)?;
        let password_verifier = reader.read_u16()?;
        let data = if password_verifier == 0 {
            FileSharingData::NoPassword {
                reserved: reader.read_u16()?,
            }
        } else {
            FileSharingData::Password {
                user_name: XlUnicodeString::read_from(reader)?,
            }
        };
        Ok(Self {
            read_only_recommended,
            password_verifier,
            data,
        })
    }
}

impl SdkWrite for FileSharingRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let has_password = self.password_verifier != 0;
        if has_password != matches!(self.data, FileSharingData::Password { .. }) {
            return Err(Error::invalid(
                writer.position()?,
                "FileSharing password verifier does not match its static payload",
            ));
        }
        self.read_only_recommended.write_to(writer)?;
        writer.write_u16(self.password_verifier)?;
        match &self.data {
            FileSharingData::NoPassword { reserved } => writer.write_u16(*reserved)?,
            FileSharingData::Password { user_name } => user_name.write_to(writer)?,
        }
        Ok(())
    }
}

impl SdkRead for MsoDrawingSelectionRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let version_and_instance = reader.read_u16()?;
        let header = OfficeArtRecordHeader {
            version: (version_and_instance & 0x000f) as u8,
            instance: version_and_instance >> 4,
            record_type: reader.read_u16()?,
            declared_length: reader.read_u32()?,
        };
        let shape_count_unused = reader.read_u32()?;
        let mode = MsoDrawingSelectionMode::read_from(reader)?;
        let focus_shape_id = reader.read_u32()?;
        let remaining = reader.remaining()?;
        if remaining % 4 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "MsoDrawingSelection shape list has a partial shape identifier",
            ));
        }
        let count = usize::try_from(remaining / 4)
            .map_err(|_| Error::Limit("drawing selection shape count exceeds usize".into()))?;
        reader.ensure_allocation(count, 4)?;
        let mut selected_shape_ids = Vec::with_capacity(count);
        for _ in 0..count {
            selected_shape_ids.push(reader.read_u32()?);
        }
        Ok(Self {
            header,
            shape_count_unused,
            mode,
            focus_shape_id,
            selected_shape_ids,
        })
    }
}

impl SdkWrite for MsoDrawingSelectionRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.header.version > 0x0f || self.header.instance > 0x0fff {
            return Err(Error::invalid(
                writer.position()?,
                "MsoDrawingSelection OfficeArt header bit field exceeds its width",
            ));
        }
        writer.write_u16(u16::from(self.header.version) | (self.header.instance << 4))?;
        writer.write_u16(self.header.record_type)?;
        writer.write_u32(self.header.declared_length)?;
        writer.write_u32(self.shape_count_unused)?;
        self.mode.write_to(writer)?;
        writer.write_u32(self.focus_shape_id)?;
        for shape_id in &self.selected_shape_ids {
            writer.write_u32(*shape_id)?;
        }
        Ok(())
    }
}

impl SdkRead for ScenManRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let scenario_count = reader.read_i16()?;
        let current_scenario_index = reader.read_i16()?;
        let shown_scenario_index = reader.read_i16()?;
        let result_reference_count = reader.read_i16()?;
        let count_position = reader.position()?.saturating_sub(2);
        let count = usize::try_from(result_reference_count).map_err(|_| {
            Error::invalid(count_position, "ScenMan result reference count is negative")
        })?;
        reader.ensure_allocation(count, 8)?;
        let mut result_references = Vec::with_capacity(count);
        for _ in 0..count {
            result_references.push(CellRange::read_from(reader)?);
        }
        Ok(Self {
            scenario_count,
            current_scenario_index,
            shown_scenario_index,
            result_reference_count,
            result_references,
        })
    }
}

impl SdkWrite for ScenManRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_count = i16::try_from(self.result_references.len())
            .map_err(|_| Error::Limit("ScenMan result reference count exceeds i16".into()))?;
        if self.result_reference_count != expected_count {
            return Err(Error::invalid(
                writer.position()?,
                "ScenMan result reference count does not match its static array",
            ));
        }
        writer.write_u16(self.scenario_count as u16)?;
        writer.write_u16(self.current_scenario_index as u16)?;
        writer.write_u16(self.shown_scenario_index as u16)?;
        writer.write_u16(self.result_reference_count as u16)?;
        for reference in &self.result_references {
            reference.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for SxViewRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let report_body = CellRange::read_from(reader)?;
        let first_header_row = reader.read_u16()?;
        let first_data_row = reader.read_u16()?;
        let first_data_column = reader.read_u16()?;
        let cache_index = reader.read_i16()?;
        let reserved = reader.read_u16()?;
        let data_axis = SxAxis::from_bits_retain(reader.read_u16()?);
        let data_position = reader.read_i16()?;
        let field_count = reader.read_i16()?;
        let row_field_count = reader.read_u16()?;
        let column_field_count = reader.read_u16()?;
        let page_field_count = reader.read_u16()?;
        let data_field_count = reader.read_i16()?;
        let row_line_count = reader.read_u16()?;
        let column_line_count = reader.read_u16()?;
        let flags = SxViewFlags::from_bits_retain(reader.read_u16()?);
        let auto_format_index = reader.read_u16()?;
        let declared_table_name_length = reader.read_u16()?;
        let declared_data_name_length = reader.read_u16()?;
        let table_name = BiffUnicodeString::read(reader, usize::from(declared_table_name_length))?;
        let data_name = BiffUnicodeString::read(reader, usize::from(declared_data_name_length))?;
        Ok(Self {
            report_body,
            first_header_row,
            first_data_row,
            first_data_column,
            cache_index,
            reserved,
            data_axis,
            data_position,
            field_count,
            row_field_count,
            column_field_count,
            page_field_count,
            data_field_count,
            row_line_count,
            column_line_count,
            flags,
            auto_format_index,
            declared_table_name_length,
            declared_data_name_length,
            table_name,
            data_name,
        })
    }
}

impl SdkWrite for SxViewRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let table_name_length = u16::try_from(self.table_name.character_count())
            .map_err(|_| Error::Limit("SxView table name exceeds u16 characters".into()))?;
        let data_name_length = u16::try_from(self.data_name.character_count())
            .map_err(|_| Error::Limit("SxView data name exceeds u16 characters".into()))?;
        if self.declared_table_name_length != table_name_length
            || self.declared_data_name_length != data_name_length
        {
            return Err(Error::invalid(
                writer.position()?,
                "SxView string count does not match its static string",
            ));
        }
        self.report_body.write_to(writer)?;
        writer.write_u16(self.first_header_row)?;
        writer.write_u16(self.first_data_row)?;
        writer.write_u16(self.first_data_column)?;
        writer.write_u16(self.cache_index as u16)?;
        writer.write_u16(self.reserved)?;
        writer.write_u16(self.data_axis.bits())?;
        writer.write_u16(self.data_position as u16)?;
        writer.write_u16(self.field_count as u16)?;
        writer.write_u16(self.row_field_count)?;
        writer.write_u16(self.column_field_count)?;
        writer.write_u16(self.page_field_count)?;
        writer.write_u16(self.data_field_count as u16)?;
        writer.write_u16(self.row_line_count)?;
        writer.write_u16(self.column_line_count)?;
        writer.write_u16(self.flags.bits())?;
        writer.write_u16(self.auto_format_index)?;
        writer.write_u16(self.declared_table_name_length)?;
        writer.write_u16(self.declared_data_name_length)?;
        self.table_name.write(writer)?;
        self.data_name.write(writer)?;
        Ok(())
    }
}

impl SdkRead for SxvdExRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let flags_and_count = reader.read_u32()?;
        let flags = SxvdExFlags::from_bits_retain(flags_and_count & 0x00ff_ffff);
        let auto_show_count = (flags_and_count >> 24) as u8;
        let auto_sort_data_item = reader.read_u16()? as i16;
        let auto_show_data_item = reader.read_u16()? as i16;
        let number_format_index = reader.read_u16()?;
        let optional = if reader.remaining()? == 0 {
            None
        } else {
            let declared_name_length = reader.read_u16()?;
            let reserved1 = reader.read_u32()?;
            let reserved2 = reader.read_u32()?;
            let subtotal_name = if declared_name_length == u16::MAX {
                None
            } else {
                Some(BiffUnicodeString::read(
                    reader,
                    usize::from(declared_name_length),
                )?)
            };
            Some(SxvdExOptional {
                declared_name_length,
                reserved1,
                reserved2,
                subtotal_name,
            })
        };
        Ok(Self {
            flags,
            auto_show_count,
            auto_sort_data_item,
            auto_show_data_item,
            number_format_index,
            optional,
        })
    }
}

impl SdkWrite for SxvdExRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u32(self.flags.bits() | (u32::from(self.auto_show_count) << 24))?;
        writer.write_u16(self.auto_sort_data_item as u16)?;
        writer.write_u16(self.auto_show_data_item as u16)?;
        writer.write_u16(self.number_format_index)?;
        if let Some(optional) = &self.optional {
            let expected_name_length = match &optional.subtotal_name {
                None => u16::MAX,
                Some(name) => u16::try_from(name.character_count())
                    .map_err(|_| Error::Limit("SXVDEx subtotal name is too long".into()))?,
            };
            if optional.declared_name_length != expected_name_length {
                return Err(Error::invalid(
                    writer.position()?,
                    "SXVDEx cchSubName does not match its static name",
                ));
            }
            writer.write_u16(optional.declared_name_length)?;
            writer.write_u32(optional.reserved1)?;
            writer.write_u32(optional.reserved2)?;
            if let Some(name) = &optional.subtotal_name {
                name.write(writer)?;
            }
        }
        Ok(())
    }
}

impl SdkRead for SxvdRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let axis = SxAxis::from_bits_retain(reader.read_u16()?);
        let subtotal_count = reader.read_u16()?;
        let subtotal_flags = SxvdSubtotalFlags::from_bits_retain(reader.read_u16()?);
        let item_count = reader.read_u16()? as i16;
        let declared_name_length = reader.read_u16()?;
        let name = if declared_name_length == u16::MAX {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::from(declared_name_length),
            )?)
        };
        Ok(Self {
            axis,
            subtotal_count,
            subtotal_flags,
            item_count,
            declared_name_length,
            name,
        })
    }
}

impl SdkWrite for SxvdRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_name_length = match &self.name {
            None => u16::MAX,
            Some(name) => u16::try_from(name.character_count())
                .map_err(|_| Error::Limit("Sxvd name exceeds u16 characters".into()))?,
        };
        if self.declared_name_length != expected_name_length {
            return Err(Error::invalid(
                writer.position()?,
                "Sxvd cchName does not match its static name",
            ));
        }
        writer.write_u16(self.axis.bits())?;
        writer.write_u16(self.subtotal_count)?;
        writer.write_u16(self.subtotal_flags.bits())?;
        writer.write_u16(self.item_count as u16)?;
        writer.write_u16(self.declared_name_length)?;
        if let Some(name) = &self.name {
            name.write(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for ChartCatLabRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeaderOld::read_from(reader)?;
        let offset_percent = reader.read_u16()?;
        let alignment = reader.read_u16()?;
        let flags = ChartCatLabFlags::from_bits_retain(reader.read_u16()?);
        let reserved = match reader.remaining()? {
            0 => None,
            2 => Some(reader.read_u16()?),
            remaining => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("CatLab has {remaining} bytes after its fixed fields"),
                ));
            }
        };
        Ok(Self {
            header,
            offset_percent,
            alignment,
            flags,
            reserved,
        })
    }
}

impl SdkWrite for ChartCatLabRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        writer.write_u16(self.offset_percent)?;
        writer.write_u16(self.alignment)?;
        writer.write_u16(self.flags.bits())?;
        if let Some(reserved) = self.reserved {
            writer.write_u16(reserved)?;
        }
        Ok(())
    }
}

impl SdkRead for ChartEndObjectRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeaderOld::read_from(reader)?;
        let object_kind = reader.read_u16()?;
        let unused = match reader.remaining()? {
            0 => None,
            6 => Some([reader.read_u16()?, reader.read_u16()?, reader.read_u16()?]),
            remaining => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("EndObject has {remaining} bytes after iObjectKind"),
                ));
            }
        };
        Ok(Self {
            header,
            object_kind,
            unused,
        })
    }
}

impl SdkWrite for ChartEndObjectRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        writer.write_u16(self.object_kind)?;
        if let Some(unused) = self.unused {
            for value in unused {
                writer.write_u16(value)?;
            }
        }
        Ok(())
    }
}

impl SheetExtColorIndex {
    fn from_bits(bits: u32) -> Self {
        Self {
            color_index: (bits & 0x7f) as u8,
            reserved: bits >> 7,
        }
    }

    fn bits(self) -> u32 {
        u32::from(self.color_index & 0x7f) | (self.reserved << 7)
    }
}

impl SheetExtOptionalFlags {
    fn from_bits(bits: u32) -> Self {
        Self {
            color_index: (bits & 0x7f) as u8,
            calculate_conditional_formats: bits & 0x80 != 0,
            not_published: bits & 0x100 != 0,
            reserved: bits >> 9,
        }
    }

    fn bits(self) -> u32 {
        u32::from(self.color_index & 0x7f)
            | (u32::from(self.calculate_conditional_formats) << 7)
            | (u32::from(self.not_published) << 8)
            | (self.reserved << 9)
    }
}

impl SdkRead for SheetExtRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let declared_size = reader.read_u32()?;
        let tab_color = SheetExtColorIndex::from_bits(reader.read_u32()?);
        let optional = match reader.remaining()? {
            0 => None,
            20 => Some(SheetExtOptional {
                flags: SheetExtOptionalFlags::from_bits(reader.read_u32()?),
                color: CfColor::read_from(reader)?,
            }),
            remaining => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("SheetExt has {remaining} bytes after its fixed fields"),
                ));
            }
        };
        Ok(Self {
            header,
            declared_size,
            tab_color,
            optional,
        })
    }
}

impl SdkWrite for SheetExtRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_size = if self.optional.is_some() { 40 } else { 20 };
        if self.declared_size != expected_size {
            return Err(Error::invalid(
                writer.position()?,
                "SheetExt cb does not match its static shape",
            ));
        }
        self.header.write_to(writer)?;
        writer.write_u32(self.declared_size)?;
        writer.write_u32(self.tab_color.bits())?;
        if let Some(optional) = self.optional {
            writer.write_u32(optional.flags.bits())?;
            optional.color.write_to(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for XlUnicodeStringMin2 {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let declared_character_count = reader.read_u16()?;
        let text = if declared_character_count == 0 {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::from(declared_character_count),
            )?)
        };
        Ok(Self {
            declared_character_count,
            text,
        })
    }
}

impl SdkWrite for XlUnicodeStringMin2 {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_count = match &self.text {
            None => 0,
            Some(text) => u16::try_from(text.character_count())
                .map_err(|_| Error::Limit("XLUnicodeStringMin2 is too long".into()))?,
        };
        if self.declared_character_count != expected_count {
            return Err(Error::invalid(
                writer.position()?,
                "XLUnicodeStringMin2 cch does not match its static text",
            ));
        }
        writer.write_u16(self.declared_character_count)?;
        if let Some(text) = &self.text {
            text.write(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for NameCommentRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let declared_name_length = reader.read_u16()?;
        let declared_comment_length = reader.read_u16()?;
        let name = BiffUnicodeString::read(reader, usize::from(declared_name_length))?;
        let comment = BiffUnicodeString::read(reader, usize::from(declared_comment_length))?;
        Ok(Self {
            header,
            declared_name_length,
            declared_comment_length,
            name,
            comment,
        })
    }
}

impl SdkWrite for NameCommentRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if usize::from(self.declared_name_length) != self.name.character_count()
            || usize::from(self.declared_comment_length) != self.comment.character_count()
        {
            return Err(Error::invalid(
                writer.position()?,
                "NameCmt character counts do not match their strings",
            ));
        }
        self.header.write_to(writer)?;
        writer.write_u16(self.declared_name_length)?;
        writer.write_u16(self.declared_comment_length)?;
        self.name.write(writer)?;
        self.comment.write(writer)?;
        Ok(())
    }
}

impl SdkRead for ChartDataLabelExtContentsRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        Ok(Self {
            header: FrtHeader::read_from(reader)?,
            flags: ChartDataLabelExtContentsFlags::from_bits_retain(reader.read_u16()?),
            separator: XlUnicodeStringMin2::read_from(reader)?,
        })
    }
}

impl SdkWrite for ChartDataLabelExtContentsRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        writer.write_u16(self.flags.bits())?;
        self.separator.write_to(writer)
    }
}

impl DxfN12List {
    fn read_sized<R: Read + Seek>(reader: &mut Reader<R>, byte_count: u32) -> Result<Self> {
        let mut child = reader.sub_reader(u64::from(byte_count))?;
        let format = Box::new(DxfN::read_from(&mut child)?);
        let extension = if child.remaining()? == 0 {
            None
        } else {
            Some(XfExtNoFrt::read_from(&mut child)?)
        };
        if child.remaining()? != 0 {
            return Err(Error::invalid(
                child.position()?,
                "DXFN12List has trailing bytes",
            ));
        }
        Ok(Self { format, extension })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        self.format.write_to(&mut writer)?;
        if let Some(extension) = &self.extension {
            extension.write_to(&mut writer)?;
        }
        Ok(writer.into_inner().into_inner())
    }
}

impl Feature11AutoFilter {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let declared_size = reader.read_u32()?;
        let unused = reader.read_u16()?;
        if declared_size != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "Feature11 embedded AutoFilter criteria are not implemented yet",
            ));
        }
        Ok(Self {
            declared_size,
            unused,
        })
    }

    fn write<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.declared_size != 0 {
            return Err(Error::invalid(
                writer.position()?,
                "Feature11 embedded AutoFilter criteria are not implemented yet",
            ));
        }
        writer.write_u32(self.declared_size)?;
        writer.write_u16(self.unused)
    }
}

impl Feature11FieldDataItem {
    fn read<R: Read + Seek>(
        reader: &mut Reader<R>,
        source_type: u32,
        table_flags: TableFeatureFlags,
        header_row_count: u32,
    ) -> Result<Self> {
        let field_id = reader.read_u32()?;
        let web_data_type = reader.read_u32()?;
        let xml_data_type = reader.read_u32()?;
        let total_aggregation = reader.read_u32()?;
        let aggregate_format_size = reader.read_u32()?;
        let aggregate_style_index = reader.read_u32()?;
        let flags = Feature11FieldFlags::from_bits_retain(reader.read_u32()?);
        let insert_row_format_size = reader.read_u32()?;
        let insert_row_style_index = reader.read_u32()?;
        let field_name = XlUnicodeString::read_from(reader)?;
        let caption = (!table_flags.contains(TableFeatureFlags::SINGLE_CELL))
            .then(|| XlUnicodeString::read_from(reader))
            .transpose()?;
        let aggregate_format = (aggregate_format_size != 0)
            .then(|| DxfN12List::read_sized(reader, aggregate_format_size))
            .transpose()?;
        let insert_row_format = (insert_row_format_size != 0)
            .then(|| DxfN12List::read_sized(reader, insert_row_format_size))
            .transpose()?;
        let auto_filter = flags
            .contains(Feature11FieldFlags::AUTO_FILTER)
            .then(|| Feature11AutoFilter::read(reader))
            .transpose()?;
        let unsupported_flags = flags
            & (Feature11FieldFlags::LOAD_XML_MAP
                | Feature11FieldFlags::LOAD_FORMULA
                | Feature11FieldFlags::LOAD_TOTAL_FORMULA
                | Feature11FieldFlags::LOAD_TOTAL_ARRAY
                | Feature11FieldFlags::SAVE_STYLE_NAME
                | Feature11FieldFlags::LOAD_TOTAL_STRING);
        if !unsupported_flags.is_empty() || source_type != 0 || header_row_count == 0 {
            return Err(Error::invalid(
                reader.position()?,
                format!(
                    "Feature11 field requires unimplemented source/header branches: source {source_type}, header rows {header_row_count}, flags 0x{:08x}",
                    unsupported_flags.bits()
                ),
            ));
        }
        Ok(Self {
            field_id,
            web_data_type,
            xml_data_type,
            total_aggregation,
            aggregate_format_size,
            aggregate_style_index,
            flags,
            insert_row_format_size,
            insert_row_style_index,
            field_name,
            caption,
            aggregate_format,
            insert_row_format,
            auto_filter,
        })
    }

    fn write<W: Write + Seek>(
        &self,
        writer: &mut Writer<W>,
        source_type: u32,
        table_flags: TableFeatureFlags,
        header_row_count: u32,
    ) -> Result<()> {
        if source_type != 0 || header_row_count == 0 {
            return Err(Error::invalid(
                writer.position()?,
                "Feature11 source/header branch is not implemented yet",
            ));
        }
        writer.write_u32(self.field_id)?;
        writer.write_u32(self.web_data_type)?;
        writer.write_u32(self.xml_data_type)?;
        writer.write_u32(self.total_aggregation)?;
        writer.write_u32(self.aggregate_format_size)?;
        writer.write_u32(self.aggregate_style_index)?;
        writer.write_u32(self.flags.bits())?;
        writer.write_u32(self.insert_row_format_size)?;
        writer.write_u32(self.insert_row_style_index)?;
        self.field_name.write_to(writer)?;
        if table_flags.contains(TableFeatureFlags::SINGLE_CELL) != self.caption.is_none() {
            return Err(Error::invalid(
                writer.position()?,
                "Feature11 caption disagrees with fSingleCell",
            ));
        }
        if let Some(caption) = &self.caption {
            caption.write_to(writer)?;
        }
        for (declared_size, value, label) in [
            (
                self.aggregate_format_size,
                self.aggregate_format.as_ref(),
                "aggregate",
            ),
            (
                self.insert_row_format_size,
                self.insert_row_format.as_ref(),
                "insert-row",
            ),
        ] {
            if let Some(value) = value {
                let bytes = value.to_bytes()?;
                if bytes.len() != usize::try_from(declared_size).unwrap_or(usize::MAX) {
                    return Err(Error::invalid(
                        writer.position()?,
                        format!("Feature11 {label} DXFN12List size mismatch"),
                    ));
                }
                writer.write_all(&bytes)?;
            } else if declared_size != 0 {
                return Err(Error::invalid(
                    writer.position()?,
                    format!("Feature11 {label} DXFN12List is missing"),
                ));
            }
        }
        if self.flags.contains(Feature11FieldFlags::AUTO_FILTER) != self.auto_filter.is_some() {
            return Err(Error::invalid(
                writer.position()?,
                "Feature11 AutoFilter flag disagrees with its field",
            ));
        }
        if let Some(auto_filter) = self.auto_filter {
            auto_filter.write(writer)?;
        }
        Ok(())
    }
}

impl SdkRead for TableFeatureType {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let source_type = reader.read_u32()?;
        let list_id = reader.read_u32()?;
        let header_row_count = reader.read_u32()?;
        let total_row_count = reader.read_u32()?;
        let next_field_id = reader.read_u32()?;
        let fixed_data_size = reader.read_u32()?;
        let writer_build = reader.read_u16()?;
        let unused1 = reader.read_u16()?;
        let flags = TableFeatureFlags::from_bits_retain(reader.read_u32()?);
        let cache_stream_offset = reader.read_u32()?;
        let cache_stream_size = reader.read_u32()?;
        let cache_character_count = reader.read_u32()?;
        let edit_mode = reader.read_u32()?;
        let hash_parameters = reader.read_array()?;
        let name = XlUnicodeString::read_from(reader)?;
        let field_count = reader.read_u16()?;
        let csp_name = flags
            .contains(TableFeatureFlags::LOAD_CSP_NAME)
            .then(|| XlUnicodeString::read_from(reader))
            .transpose()?;
        let entry_id = flags
            .contains(TableFeatureFlags::LOAD_ENTRY_ID)
            .then(|| XlUnicodeString::read_from(reader))
            .transpose()?;
        if fixed_data_size != 64 {
            return Err(Error::invalid(
                reader.position()?,
                format!("TableFeatureType cbFSData is {fixed_data_size}, expected 64"),
            ));
        }
        let unsupported_table_flags = flags
            & (TableFeatureFlags::LOAD_DELETED_IDS
                | TableFeatureFlags::LOAD_CHANGED_IDS
                | TableFeatureFlags::LOAD_INVALID_CELLS);
        if !unsupported_table_flags.is_empty() {
            return Err(Error::invalid(
                reader.position()?,
                format!(
                    "TableFeatureType requires unimplemented trailing arrays 0x{:08x}",
                    unsupported_table_flags.bits()
                ),
            ));
        }
        let mut fields = Vec::with_capacity(usize::from(field_count));
        reader.ensure_allocation(usize::from(field_count), 36)?;
        for _ in 0..field_count {
            fields.push(Feature11FieldDataItem::read(
                reader,
                source_type,
                flags,
                header_row_count,
            )?);
        }
        Ok(Self {
            source_type,
            list_id,
            header_row_count,
            total_row_count,
            next_field_id,
            fixed_data_size,
            writer_build,
            unused1,
            flags,
            cache_stream_offset,
            cache_stream_size,
            cache_character_count,
            edit_mode,
            hash_parameters,
            name,
            field_count,
            csp_name,
            entry_id,
            fields,
        })
    }
}

impl SdkWrite for TableFeatureType {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.fixed_data_size != 64 || usize::from(self.field_count) != self.fields.len() {
            return Err(Error::invalid(
                writer.position()?,
                "TableFeatureType fixed size or field count mismatch",
            ));
        }
        for (flag, present, name) in [
            (
                TableFeatureFlags::LOAD_CSP_NAME,
                self.csp_name.is_some(),
                "CSP name",
            ),
            (
                TableFeatureFlags::LOAD_ENTRY_ID,
                self.entry_id.is_some(),
                "entry ID",
            ),
        ] {
            if self.flags.contains(flag) != present {
                return Err(Error::invalid(
                    writer.position()?,
                    format!("TableFeatureType {name} flag mismatch"),
                ));
            }
        }
        writer.write_u32(self.source_type)?;
        writer.write_u32(self.list_id)?;
        writer.write_u32(self.header_row_count)?;
        writer.write_u32(self.total_row_count)?;
        writer.write_u32(self.next_field_id)?;
        writer.write_u32(self.fixed_data_size)?;
        writer.write_u16(self.writer_build)?;
        writer.write_u16(self.unused1)?;
        writer.write_u32(self.flags.bits())?;
        writer.write_u32(self.cache_stream_offset)?;
        writer.write_u32(self.cache_stream_size)?;
        writer.write_u32(self.cache_character_count)?;
        writer.write_u32(self.edit_mode)?;
        writer.write_all(&self.hash_parameters)?;
        self.name.write_to(writer)?;
        writer.write_u16(self.field_count)?;
        if let Some(csp_name) = &self.csp_name {
            csp_name.write_to(writer)?;
        }
        if let Some(entry_id) = &self.entry_id {
            entry_id.write_to(writer)?;
        }
        for field in &self.fields {
            field.write(writer, self.source_type, self.flags, self.header_row_count)?;
        }
        Ok(())
    }
}

impl SdkRead for Feature11Record {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtRefHeaderU::read_from(reader)?;
        let shared_feature_type = reader.read_u16()?;
        let reserved1 = reader.read_u8()?;
        let reserved2 = reader.read_u32()?;
        let reference_count = reader.read_u16()?;
        let declared_feature_size = reader.read_u32()?;
        let reserved3 = reader.read_u16()?;
        reader.ensure_allocation(usize::from(reference_count), 8)?;
        let mut references = Vec::with_capacity(usize::from(reference_count));
        for _ in 0..reference_count {
            references.push(CellRange::read_from(reader)?);
        }
        let feature = if declared_feature_size == 0 {
            TableFeatureType::read_from(reader)?
        } else {
            let mut child = reader.sub_reader(u64::from(declared_feature_size))?;
            let feature = TableFeatureType::read_from(&mut child)?;
            if child.remaining()? != 0 {
                return Err(Error::invalid(
                    child.position()?,
                    "Feature11 rgbFeat has trailing bytes",
                ));
            }
            feature
        };
        Ok(Self {
            header,
            shared_feature_type,
            reserved1,
            reserved2,
            reference_count,
            declared_feature_size,
            reserved3,
            references,
            feature,
        })
    }
}

impl SdkWrite for Feature11Record {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if usize::from(self.reference_count) != self.references.len() {
            return Err(Error::invalid(
                writer.position()?,
                "Feature11 cref2 does not match refs2",
            ));
        }
        let feature_bytes = encode_sdk(&self.feature)?;
        if self.declared_feature_size != 0
            && usize::try_from(self.declared_feature_size).unwrap_or(usize::MAX)
                != feature_bytes.len()
        {
            return Err(Error::invalid(
                writer.position()?,
                "Feature11 cbFeatData does not match rgbFeat",
            ));
        }
        self.header.write_to(writer)?;
        writer.write_u16(self.shared_feature_type)?;
        writer.write_u8(self.reserved1)?;
        writer.write_u32(self.reserved2)?;
        writer.write_u16(self.reference_count)?;
        writer.write_u32(self.declared_feature_size)?;
        writer.write_u16(self.reserved3)?;
        for reference in &self.references {
            reference.write_to(writer)?;
        }
        writer.write_all(&feature_bytes)?;
        Ok(())
    }
}

fn list12_format_size(value: i32, label: &str) -> Result<u32> {
    u32::try_from(value).map_err(|_| Error::invalid(0, format!("negative List12 {label} size")))
}

impl List12BlockLevel {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header_format_size = reader.read_u32()? as i32;
        let header_style_index = reader.read_u32()? as i32;
        let data_format_size = reader.read_u32()? as i32;
        let data_style_index = reader.read_u32()? as i32;
        let aggregate_format_size = reader.read_u32()? as i32;
        let aggregate_style_index = reader.read_u32()? as i32;
        let border_format_size = reader.read_u32()? as i32;
        let header_border_format_size = reader.read_u32()? as i32;
        let aggregate_border_format_size = reader.read_u32()? as i32;
        let header_size = list12_format_size(header_format_size, "header")?;
        let data_size = list12_format_size(data_format_size, "data")?;
        let aggregate_size = list12_format_size(aggregate_format_size, "aggregate")?;
        let border_size = list12_format_size(border_format_size, "border")?;
        let header_border_size = list12_format_size(header_border_format_size, "header-border")?;
        let aggregate_border_size =
            list12_format_size(aggregate_border_format_size, "aggregate-border")?;
        let header_format = (header_size != 0)
            .then(|| DxfN12List::read_sized(reader, header_size))
            .transpose()?;
        let data_format = (data_size != 0)
            .then(|| DxfN12List::read_sized(reader, data_size))
            .transpose()?;
        let aggregate_format = (aggregate_size != 0)
            .then(|| DxfN12List::read_sized(reader, aggregate_size))
            .transpose()?;
        let border_format = (border_size != 0)
            .then(|| DxfN12List::read_sized(reader, border_size))
            .transpose()?;
        let header_border_format = (header_border_size != 0)
            .then(|| DxfN12List::read_sized(reader, header_border_size))
            .transpose()?;
        let aggregate_border_format = (aggregate_border_size != 0)
            .then(|| DxfN12List::read_sized(reader, aggregate_border_size))
            .transpose()?;
        let header_style_name = (header_style_index != -1)
            .then(|| XlUnicodeString::read_from(reader))
            .transpose()?;
        let data_style_name = (data_style_index != -1)
            .then(|| XlUnicodeString::read_from(reader))
            .transpose()?;
        let aggregate_style_name = (aggregate_style_index != -1)
            .then(|| XlUnicodeString::read_from(reader))
            .transpose()?;
        Ok(Self {
            header_format_size,
            header_style_index,
            data_format_size,
            data_style_index,
            aggregate_format_size,
            aggregate_style_index,
            border_format_size,
            header_border_format_size,
            aggregate_border_format_size,
            header_format,
            data_format,
            aggregate_format,
            border_format,
            header_border_format,
            aggregate_border_format,
            header_style_name,
            data_style_name,
            aggregate_style_name,
        })
    }

    fn write<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        for value in [
            self.header_format_size,
            self.header_style_index,
            self.data_format_size,
            self.data_style_index,
            self.aggregate_format_size,
            self.aggregate_style_index,
            self.border_format_size,
            self.header_border_format_size,
            self.aggregate_border_format_size,
        ] {
            writer.write_u32(value as u32)?;
        }
        for (size, value, label) in [
            (
                self.header_format_size,
                self.header_format.as_ref(),
                "header",
            ),
            (self.data_format_size, self.data_format.as_ref(), "data"),
            (
                self.aggregate_format_size,
                self.aggregate_format.as_ref(),
                "aggregate",
            ),
        ] {
            write_list12_dxf_list(writer, size, value, label)?;
        }
        write_list12_dxf_list(
            writer,
            self.border_format_size,
            self.border_format.as_ref(),
            "border",
        )?;
        for (size, value, label) in [
            (
                self.header_border_format_size,
                self.header_border_format.as_ref(),
                "header-border",
            ),
            (
                self.aggregate_border_format_size,
                self.aggregate_border_format.as_ref(),
                "aggregate-border",
            ),
        ] {
            write_list12_dxf_list(writer, size, value, label)?;
        }
        for (index, value, label) in [
            (
                self.header_style_index,
                self.header_style_name.as_ref(),
                "header",
            ),
            (self.data_style_index, self.data_style_name.as_ref(), "data"),
            (
                self.aggregate_style_index,
                self.aggregate_style_name.as_ref(),
                "aggregate",
            ),
        ] {
            if (index != -1) != value.is_some() {
                return Err(Error::invalid(
                    writer.position()?,
                    format!("List12 {label} style name presence mismatch"),
                ));
            }
            if let Some(value) = value {
                value.write_to(writer)?;
            }
        }
        Ok(())
    }
}

fn write_list12_dxf_list<W: Write + Seek>(
    writer: &mut Writer<W>,
    size: i32,
    value: Option<&DxfN12List>,
    label: &str,
) -> Result<()> {
    let size = list12_format_size(size, label)?;
    if let Some(value) = value {
        let bytes = value.to_bytes()?;
        if bytes.len() != usize::try_from(size).unwrap_or(usize::MAX) {
            return Err(Error::invalid(
                writer.position()?,
                format!("List12 {label} DXFN12List size mismatch"),
            ));
        }
        writer.write_all(&bytes)?;
    } else if size != 0 {
        return Err(Error::invalid(
            writer.position()?,
            format!("List12 {label} DXFN12List is missing"),
        ));
    }
    Ok(())
}

impl SdkRead for List12Record {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let data_type = reader.read_u16()?;
        let list_id = reader.read_u32()?;
        let data = match data_type {
            0 => List12Data::BlockLevel(Box::new(List12BlockLevel::read(reader)?)),
            1 => List12Data::TableStyle {
                flags: List12TableStyleFlags::from_bits_retain(reader.read_u16()?),
                style_name: XlUnicodeString::read_from(reader)?,
            },
            2 => List12Data::DisplayName {
                list_name: XlUnicodeString::read_from(reader)?,
                comment: XlUnicodeString::read_from(reader)?,
            },
            value => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("unsupported List12 lsd 0x{value:04x}"),
                ));
            }
        };
        Ok(Self {
            header,
            data_type,
            list_id,
            data,
        })
    }
}

impl SdkWrite for List12Record {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let expected_type = match &self.data {
            List12Data::BlockLevel(_) => 0,
            List12Data::TableStyle { .. } => 1,
            List12Data::DisplayName { .. } => 2,
        };
        if self.data_type != expected_type {
            return Err(Error::invalid(
                writer.position()?,
                "List12 lsd disagrees with rgb",
            ));
        }
        self.header.write_to(writer)?;
        writer.write_u16(self.data_type)?;
        writer.write_u32(self.list_id)?;
        match &self.data {
            List12Data::BlockLevel(value) => value.write(writer)?,
            List12Data::TableStyle { flags, style_name } => {
                writer.write_u16(flags.bits())?;
                style_name.write_to(writer)?;
            }
            List12Data::DisplayName { list_name, comment } => {
                list_name.write_to(writer)?;
                comment.write_to(writer)?;
            }
        }
        Ok(())
    }
}

impl SdkRead for LabelRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let cell = CellHeader::read_from(reader)?;
        let character_count = usize::from(reader.read_u16()?);
        let text = BiffUnicodeString::read(reader, character_count)?;
        let trailing_null = match reader.remaining()? {
            0 => None,
            2 => Some(reader.read_u16()?),
            remaining => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("Label has {remaining} bytes after its declared text"),
                ));
            }
        };
        Ok(Self {
            cell,
            text,
            trailing_null,
        })
    }
}

impl SdkRead for PaletteRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let color_count = usize::from(reader.read_u16()?);
        let mut colors = Vec::with_capacity(color_count);
        for _ in 0..color_count {
            colors.push(reader.read_u32()?);
        }
        Ok(Self { colors })
    }
}

impl SdkWrite for PaletteRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(
            u16::try_from(self.colors.len())
                .map_err(|_| Error::Limit("Palette color count exceeds u16".into()))?,
        )?;
        for color in &self.colors {
            writer.write_u32(*color)?;
        }
        Ok(())
    }
}

impl SdkWrite for LabelRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.cell.write_to(writer)?;
        writer.write_u16(
            u16::try_from(self.text.character_count())
                .map_err(|_| Error::Limit("Label text exceeds u16 characters".into()))?,
        )?;
        self.text.write(writer)?;
        if let Some(value) = self.trailing_null {
            writer.write_u16(value)?;
        }
        Ok(())
    }
}

impl SdkRead for ChartSeriesTextRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let reserved = reader.read_u16()?;
        let character_count = usize::from(reader.read_u8()?);
        let text = BiffUnicodeString::read(reader, character_count)?;
        Ok(Self { reserved, text })
    }
}

impl SdkWrite for ChartSeriesTextRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.reserved)?;
        writer.write_u8(
            u8::try_from(self.text.character_count())
                .map_err(|_| Error::Limit("SeriesText exceeds u8 characters".into()))?,
        )?;
        self.text.write(writer)
    }
}

impl SdkRead for EndBlockRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeaderOld::read_from(reader)?;
        let object_kind = reader.read_u16()?;
        let optional_unused = match reader.remaining()? {
            0 => None,
            6 => Some([reader.read_u16()?, reader.read_u16()?, reader.read_u16()?]),
            remaining => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("EndBlock has {remaining} bytes after object kind; expected 0 or 6"),
                ));
            }
        };
        Ok(Self {
            header,
            object_kind,
            optional_unused,
        })
    }
}

impl SdkWrite for EndBlockRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        writer.write_u16(self.object_kind)?;
        if let Some(values) = self.optional_unused {
            for value in values {
                writer.write_u16(value)?;
            }
        }
        Ok(())
    }
}

impl SdkRead for FontRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let height_twips = reader.read_u16()?;
        let attributes = FontAttributes::from_bits_retain(reader.read_u16()?);
        let color_index = reader.read_u16()?;
        let bold_weight = reader.read_u16()?;
        let escapement = reader.read_u16()?;
        let underline = reader.read_u8()?;
        let family = reader.read_u8()?;
        let charset = reader.read_u8()?;
        let reserved = reader.read_u8()?;
        let count = usize::from(reader.read_u8()?);
        let name = BiffUnicodeString::read(reader, count)?;
        Ok(Self {
            height_twips,
            attributes,
            color_index,
            bold_weight,
            escapement,
            underline,
            family,
            charset,
            reserved,
            name,
        })
    }
}

impl SdkWrite for FontRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.height_twips)?;
        writer.write_u16(self.attributes.bits())?;
        writer.write_u16(self.color_index)?;
        writer.write_u16(self.bold_weight)?;
        writer.write_u16(self.escapement)?;
        writer.write_u8(self.underline)?;
        writer.write_u8(self.family)?;
        writer.write_u8(self.charset)?;
        writer.write_u8(self.reserved)?;
        writer.write_u8(
            u8::try_from(self.name.character_count())
                .map_err(|_| Error::Limit("Font name exceeds u8".into()))?,
        )?;
        self.name.write(writer)
    }
}

impl SdkRead for FormatRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let format_index = reader.read_u16()?;
        let declared_character_count = reader.read_u16()?;
        let format_string = BiffUnicodeString::read_remaining(reader)?;
        Ok(Self {
            format_index,
            declared_character_count,
            format_string,
        })
    }
}

impl SdkWrite for FormatRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.format_index)?;
        writer.write_u16(self.declared_character_count)?;
        self.format_string.write(writer)
    }
}

impl SdkRead for StyleRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let xf_and_flags = reader.read_u16()?;
        let data = if xf_and_flags & 0x8000 != 0 {
            StyleData::BuiltIn {
                style_id: reader.read_u8()?,
                outline_level: reader.read_u8()?,
            }
        } else {
            let declared_character_count = reader.read_u16()?;
            let name = if reader.remaining()? == 0 {
                if declared_character_count != 0 {
                    return Err(Error::invalid(
                        reader.position()?,
                        "user-defined Style name is missing",
                    ));
                }
                None
            } else {
                Some(BiffUnicodeString::read(
                    reader,
                    usize::from(declared_character_count),
                )?)
            };
            StyleData::UserDefined {
                declared_character_count,
                name,
            }
        };
        Ok(Self { xf_and_flags, data })
    }
}

impl SdkWrite for StyleRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.xf_and_flags)?;
        match &self.data {
            StyleData::BuiltIn {
                style_id,
                outline_level,
            } => {
                if self.xf_and_flags & 0x8000 == 0 {
                    return Err(Error::invalid(
                        writer.position()?,
                        "built-in Style lacks fBuiltIn flag",
                    ));
                }
                writer.write_u8(*style_id)?;
                writer.write_u8(*outline_level)?;
            }
            StyleData::UserDefined {
                declared_character_count,
                name,
            } => {
                if self.xf_and_flags & 0x8000 != 0 {
                    return Err(Error::invalid(
                        writer.position()?,
                        "user-defined Style has fBuiltIn flag",
                    ));
                }
                writer.write_u16(*declared_character_count)?;
                if let Some(name) = name {
                    name.write(writer)?;
                }
            }
        }
        Ok(())
    }
}

impl SdkRead for WriteAccessRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let declared_character_count = reader.read_u16()?;
        let flags = reader.read_u8()?;
        let remaining = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("WriteAccess length exceeds usize".into()))?;
        let requested = usize::from(declared_character_count)
            .checked_mul(if flags & 1 == 0 { 1 } else { 2 })
            .ok_or_else(|| Error::Limit("WriteAccess name length overflow".into()))?;
        let mut name_len = requested.min(remaining);
        if flags & 1 != 0 {
            name_len -= name_len % 2;
        }
        let raw_name = reader.read_vec(name_len)?;
        let name = if flags & 1 == 0 {
            XlStringCharacters::Compressed(raw_name)
        } else {
            XlStringCharacters::Unicode(
                raw_name
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect(),
            )
        };
        let unused_len = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("WriteAccess padding length exceeds usize".into()))?;
        Ok(Self {
            declared_character_count,
            flags,
            name,
            trailing_name_byte: None,
            unused: reader.read_vec(unused_len)?,
        })
    }
}

impl SdkWrite for WriteAccessRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let name_bytes = match &self.name {
            XlStringCharacters::Compressed(values) => values.len(),
            XlStringCharacters::Unicode(values) => values
                .len()
                .checked_mul(2)
                .ok_or_else(|| Error::Limit("WriteAccess Unicode name length overflow".into()))?,
        } + usize::from(self.trailing_name_byte.is_some());
        let total = 3usize
            .checked_add(name_bytes)
            .and_then(|value| value.checked_add(self.unused.len()))
            .ok_or_else(|| Error::Limit("WriteAccess size overflow".into()))?;
        if total != 112 {
            return Err(Error::invalid(
                writer.position()?,
                "WriteAccess must contain exactly 112 bytes",
            ));
        }
        writer.write_u16(self.declared_character_count)?;
        let string = BiffUnicodeString {
            flags: self.flags,
            characters: self.name.clone(),
            trailing_byte: self.trailing_name_byte,
        };
        string.write(writer)?;
        writer.write_all(&self.unused)?;
        Ok(())
    }
}

impl SdkRead for Window2Record {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let flags = Window2Flags::from_bits_retain(reader.read_u16()?);
        let top_row = reader.read_u16()?;
        let left_column = reader.read_u16()?;
        let header_color = reader.read_u32()?;
        let extension = match reader.remaining()? {
            0 => Window2Extension::None,
            4 => Window2Extension::Zoom {
                page_break_zoom: reader.read_u16()?,
                normal_zoom: reader.read_u16()?,
                reserved: None,
            },
            8 => Window2Extension::Zoom {
                page_break_zoom: reader.read_u16()?,
                normal_zoom: reader.read_u16()?,
                reserved: Some(reader.read_u32()?),
            },
            remaining => {
                return Err(Error::invalid(
                    reader.position()?,
                    format!("Window2 has invalid extension length {remaining}"),
                ));
            }
        };
        Ok(Self {
            flags,
            top_row,
            left_column,
            header_color,
            extension,
        })
    }
}

impl SdkWrite for Window2Record {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.flags.bits())?;
        writer.write_u16(self.top_row)?;
        writer.write_u16(self.left_column)?;
        writer.write_u32(self.header_color)?;
        match self.extension {
            Window2Extension::None => {}
            Window2Extension::Zoom {
                page_break_zoom,
                normal_zoom,
                reserved,
            } => {
                writer.write_u16(page_break_zoom)?;
                writer.write_u16(normal_zoom)?;
                if let Some(value) = reserved {
                    writer.write_u32(value)?;
                }
            }
        }
        Ok(())
    }
}

impl SdkRead for IndexRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let reserved1 = reader.read_u32()?;
        let first_row = reader.read_u32()?;
        let last_row_exclusive = reader.read_u32()?;
        let reserved2 = reader.read_u32()?;
        let remaining = reader.remaining()?;
        if remaining % 4 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "Index DBCell pointer array has invalid byte length",
            ));
        }
        let count = usize::try_from(remaining / 4)
            .map_err(|_| Error::Limit("Index DBCell pointer count exceeds usize".into()))?;
        reader.ensure_allocation(count, 4)?;
        let mut dbcell_offsets = Vec::with_capacity(count);
        for _ in 0..count {
            dbcell_offsets.push(reader.read_u32()?);
        }
        Ok(Self {
            reserved1,
            first_row,
            last_row_exclusive,
            reserved2,
            dbcell_offsets,
        })
    }
}

impl SdkWrite for IndexRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u32(self.reserved1)?;
        writer.write_u32(self.first_row)?;
        writer.write_u32(self.last_row_exclusive)?;
        writer.write_u32(self.reserved2)?;
        for offset in &self.dbcell_offsets {
            writer.write_u32(*offset)?;
        }
        Ok(())
    }
}

impl StringValueRecord {
    fn from_sequence(first: &[u8], following: &[BiffRecord]) -> Result<(Self, usize)> {
        if first.len() < 3 {
            return Err(Error::invalid(0, "String record header is truncated"));
        }
        let declared_character_count = u16::from_le_bytes([first[0], first[1]]);
        let mut remaining = usize::from(declared_character_count);
        let mut chunks = Vec::new();
        let mut consumed_continues = 0usize;
        let mut payload = first;
        let mut data_offset = 3usize;
        let mut flags = first[2];
        loop {
            let (chunk, character_count) =
                ContinuedStringChunk::parse(payload, data_offset, flags, remaining)?;
            remaining -= character_count;
            chunks.push(chunk);
            if remaining == 0 {
                break;
            }
            let next = following
                .get(consumed_continues)
                .ok_or_else(|| Error::invalid(0, "String record is missing a Continue record"))?;
            let BiffRecordData::Continue {
                payload: next_payload,
            } = &next.data
            else {
                return Err(Error::invalid(
                    u64::from(next.offset),
                    "String record is not followed by Continue",
                ));
            };
            if next_payload.is_empty() {
                return Err(Error::invalid(
                    u64::from(next.offset),
                    "String Continue record lacks option flags",
                ));
            }
            consumed_continues += 1;
            payload = next_payload;
            flags = next_payload[0];
            data_offset = 1;
        }
        Ok((
            Self {
                declared_character_count,
                chunks,
            },
            consumed_continues,
        ))
    }

    fn encode_physical(&self) -> Result<Vec<EncodedBiffRecord>> {
        if self.chunks.is_empty() {
            return Err(Error::invalid(0, "String record has no physical chunks"));
        }
        let actual_count: usize = self
            .chunks
            .iter()
            .map(ContinuedStringChunk::character_count)
            .sum();
        if actual_count != usize::from(self.declared_character_count) {
            return Err(Error::invalid(
                0,
                "String declared and actual character counts disagree",
            ));
        }
        let mut encoded = Vec::with_capacity(self.chunks.len());
        for (index, chunk) in self.chunks.iter().enumerate() {
            let mut payload = Vec::new();
            if index == 0 {
                payload.extend_from_slice(&self.declared_character_count.to_le_bytes());
            }
            payload.push(chunk.flags);
            chunk.write_characters(&mut payload)?;
            payload.extend_from_slice(&chunk.trailing);
            encoded.push(EncodedBiffRecord {
                record_type: if index == 0 { STRING_VALUE } else { CONTINUE },
                payload,
            });
        }
        Ok(encoded)
    }
}

impl ContinuedStringChunk {
    fn parse(
        payload: &[u8],
        data_offset: usize,
        flags: u8,
        remaining_characters: usize,
    ) -> Result<(Self, usize)> {
        let data = payload
            .get(data_offset..)
            .ok_or_else(|| Error::invalid(0, "String segment header exceeds payload"))?;
        let available = if flags & 1 == 0 {
            data.len()
        } else {
            data.len() / 2
        };
        let count = remaining_characters.min(available);
        let consumed_bytes = count * if flags & 1 == 0 { 1 } else { 2 };
        if count < remaining_characters && consumed_bytes != data.len() {
            return Err(Error::invalid(
                consumed_bytes as u64,
                "UTF-16 String segment ends with an incomplete character",
            ));
        }
        let raw = &data[..consumed_bytes];
        let characters = if flags & 1 == 0 {
            XlStringCharacters::Compressed(raw.to_vec())
        } else {
            XlStringCharacters::Unicode(
                raw.chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect(),
            )
        };
        Ok((
            Self {
                flags,
                characters,
                trailing: data[consumed_bytes..].to_vec(),
            },
            count,
        ))
    }

    fn character_count(&self) -> usize {
        match &self.characters {
            XlStringCharacters::Compressed(values) => values.len(),
            XlStringCharacters::Unicode(values) => values.len(),
        }
    }

    fn write_characters(&self, bytes: &mut Vec<u8>) -> Result<()> {
        match &self.characters {
            XlStringCharacters::Compressed(values) => {
                if self.flags & 1 != 0 {
                    return Err(Error::invalid(0, "compressed String chunk has UTF-16 flag"));
                }
                bytes.extend_from_slice(values);
            }
            XlStringCharacters::Unicode(values) => {
                if self.flags & 1 == 0 {
                    return Err(Error::invalid(0, "Unicode String chunk lacks UTF-16 flag"));
                }
                for value in values {
                    bytes.extend_from_slice(&value.to_le_bytes());
                }
            }
        }
        Ok(())
    }
}

impl NameRecord {
    fn from_sequence(first: &[u8], continues: &[&[u8]], limits: Limits) -> Result<Self> {
        let total_len = std::iter::once(first)
            .chain(continues.iter().copied())
            .try_fold(0usize, |total, payload| {
                total
                    .checked_add(payload.len())
                    .ok_or_else(|| Error::Limit("Name logical record length overflow".into()))
            })?;
        if total_len > limits.max_allocation {
            return Err(Error::Limit(format!(
                "Name logical record length exceeds {}",
                limits.max_allocation
            )));
        }
        let mut bytes = Vec::with_capacity(total_len);
        bytes.extend_from_slice(first);
        for payload in continues {
            bytes.extend_from_slice(payload);
        }
        let physical_segment_lengths = std::iter::once(first)
            .chain(continues.iter().copied())
            .map(|payload| {
                u16::try_from(payload.len())
                    .map_err(|_| Error::Limit("Name physical segment exceeds u16".into()))
            })
            .collect::<Result<Vec<_>>>()?;
        let mut cursor = 0usize;
        let flags =
            NameFlags::from_bits_retain(take_u16(&bytes, &mut cursor, "truncated Name flags")?);
        let keyboard_shortcut = take_u8(&bytes, &mut cursor, "truncated Name shortcut")?;
        let declared_name_character_count =
            take_u8(&bytes, &mut cursor, "truncated Name character count")?;
        let declared_formula_byte_count =
            take_u16(&bytes, &mut cursor, "truncated Name formula length")?;
        let reserved3 = take_u16(&bytes, &mut cursor, "truncated Name reserved3")?;
        let sheet_index = take_u16(&bytes, &mut cursor, "truncated Name sheet index")?;
        let custom_menu_count = take_u8(&bytes, &mut cursor, "truncated Name reserved4")?;
        let description_count = take_u8(&bytes, &mut cursor, "truncated Name reserved5")?;
        let help_topic_count = take_u8(&bytes, &mut cursor, "truncated Name reserved6")?;
        let status_bar_count = take_u8(&bytes, &mut cursor, "truncated Name reserved7")?;
        let name_flags = NameStringFlags::from_bits_retain(take_u8(
            &bytes,
            &mut cursor,
            "truncated Name string flags",
        )?);
        let name = if flags.contains(NameFlags::BUILT_IN) {
            NameValue::BuiltIn(take_u8(
                &bytes,
                &mut cursor,
                "truncated built-in Name index",
            )?)
        } else if name_flags.contains(NameStringFlags::HIGH_BYTE) {
            let byte_count = usize::from(declared_name_character_count)
                .checked_mul(2)
                .ok_or_else(|| Error::Limit("Name Unicode byte count overflow".into()))?;
            let raw = take_bytes(&bytes, &mut cursor, byte_count, "truncated Unicode Name")?;
            NameValue::User(XlStringCharacters::Unicode(
                raw.chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect(),
            ))
        } else {
            NameValue::User(XlStringCharacters::Compressed(
                take_bytes(
                    &bytes,
                    &mut cursor,
                    usize::from(declared_name_character_count),
                    "truncated compressed Name",
                )?
                .to_vec(),
            ))
        };
        let rgce = take_bytes(
            &bytes,
            &mut cursor,
            usize::from(declared_formula_byte_count),
            "truncated Name formula tokens",
        )?;
        let trailing_text_len = [
            custom_menu_count,
            description_count,
            help_topic_count,
            status_bar_count,
        ]
        .into_iter()
        .map(usize::from)
        .sum::<usize>();
        let formula_region_end = bytes
            .len()
            .checked_sub(trailing_text_len)
            .filter(|end| *end >= cursor)
            .ok_or_else(|| Error::invalid(cursor as u64, "Name trailing text lengths overflow"))?;
        let mut formula = FormulaTokenStream::from_bytes(rgce)?;
        let formula_extra_tail = formula.parse_extra_data(&bytes[cursor..formula_region_end])?;
        cursor = formula_region_end;
        let mut read_text = |count: u8, context: &str| -> Result<NameTrailingText> {
            Ok(NameTrailingText {
                declared_character_count: count,
                characters: take_bytes(&bytes, &mut cursor, usize::from(count), context)?.to_vec(),
            })
        };
        let custom_menu = read_text(custom_menu_count, "truncated Name custom-menu text")?;
        let description = read_text(description_count, "truncated Name description text")?;
        let help_topic = read_text(help_topic_count, "truncated Name help-topic text")?;
        let status_bar = read_text(status_bar_count, "truncated Name status-bar text")?;
        debug_assert_eq!(cursor, bytes.len());
        Ok(Self {
            flags,
            keyboard_shortcut,
            declared_name_character_count,
            declared_formula_byte_count,
            reserved3,
            sheet_index,
            custom_menu,
            description,
            help_topic,
            status_bar,
            name_flags,
            name,
            formula,
            formula_extra_tail,
            physical_segment_lengths,
        })
    }

    fn logical_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&self.flags.bits().to_le_bytes());
        bytes.push(self.keyboard_shortcut);
        bytes.push(self.declared_name_character_count);
        bytes.extend_from_slice(&self.declared_formula_byte_count.to_le_bytes());
        bytes.extend_from_slice(&self.reserved3.to_le_bytes());
        bytes.extend_from_slice(&self.sheet_index.to_le_bytes());
        for text in [
            &self.custom_menu,
            &self.description,
            &self.help_topic,
            &self.status_bar,
        ] {
            if usize::from(text.declared_character_count) != text.characters.len() {
                return Err(Error::invalid(0, "Name trailing text count mismatch"));
            }
            bytes.push(text.declared_character_count);
        }
        bytes.push(self.name_flags.bits());
        match &self.name {
            NameValue::BuiltIn(index) => {
                if !self.flags.contains(NameFlags::BUILT_IN) {
                    return Err(Error::invalid(0, "built-in Name lacks BUILT_IN flag"));
                }
                bytes.push(*index);
            }
            NameValue::User(XlStringCharacters::Compressed(characters)) => {
                if self.flags.contains(NameFlags::BUILT_IN)
                    || self.name_flags.contains(NameStringFlags::HIGH_BYTE)
                    || usize::from(self.declared_name_character_count) != characters.len()
                {
                    return Err(Error::invalid(0, "compressed Name fields mismatch"));
                }
                bytes.extend_from_slice(characters);
            }
            NameValue::User(XlStringCharacters::Unicode(characters)) => {
                if self.flags.contains(NameFlags::BUILT_IN)
                    || !self.name_flags.contains(NameStringFlags::HIGH_BYTE)
                    || usize::from(self.declared_name_character_count) != characters.len()
                {
                    return Err(Error::invalid(0, "Unicode Name fields mismatch"));
                }
                for character in characters {
                    bytes.extend_from_slice(&character.to_le_bytes());
                }
            }
        }
        let rgce = self.formula.to_bytes()?;
        if usize::from(self.declared_formula_byte_count) != rgce.len() {
            return Err(Error::invalid(0, "Name formula byte count mismatch"));
        }
        bytes.extend_from_slice(&rgce);
        bytes.extend_from_slice(&self.formula.extra_data_to_bytes()?);
        bytes.extend_from_slice(&self.formula_extra_tail);
        for text in [
            &self.custom_menu,
            &self.description,
            &self.help_topic,
            &self.status_bar,
        ] {
            bytes.extend_from_slice(&text.characters);
        }
        Ok(bytes)
    }

    fn encode_physical(&self) -> Result<Vec<EncodedBiffRecord>> {
        let logical = self.logical_bytes()?;
        let declared_len = self
            .physical_segment_lengths
            .iter()
            .map(|length| usize::from(*length))
            .sum::<usize>();
        if self.physical_segment_lengths.is_empty() || declared_len != logical.len() {
            return Err(Error::invalid(0, "Name physical layout mismatch"));
        }
        let mut offset = 0usize;
        let mut records = Vec::with_capacity(self.physical_segment_lengths.len());
        for (index, length) in self.physical_segment_lengths.iter().enumerate() {
            let end = offset + usize::from(*length);
            records.push(EncodedBiffRecord {
                record_type: if index == 0 { NAME } else { CONTINUE },
                payload: logical[offset..end].to_vec(),
            });
            offset = end;
        }
        Ok(records)
    }
}

impl PlsRecord {
    fn from_sequence(first: &[u8], continues: &[&[u8]], limits: Limits) -> Result<Self> {
        let total_len = std::iter::once(first)
            .chain(continues.iter().copied())
            .try_fold(0usize, |total, payload| {
                total
                    .checked_add(payload.len())
                    .ok_or_else(|| Error::Limit("Pls logical record length overflow".into()))
            })?;
        if total_len > limits.max_allocation {
            return Err(Error::Limit(format!(
                "Pls logical record length exceeds {}",
                limits.max_allocation
            )));
        }
        let mut bytes = Vec::with_capacity(total_len);
        bytes.extend_from_slice(first);
        for payload in continues {
            bytes.extend_from_slice(payload);
        }
        let physical_segment_lengths = std::iter::once(first)
            .chain(continues.iter().copied())
            .map(|payload| {
                u16::try_from(payload.len())
                    .map_err(|_| Error::Limit("Pls physical segment exceeds u16".into()))
            })
            .collect::<Result<Vec<_>>>()?;
        let mut cursor = 0usize;
        let reserved = take_u16(&bytes, &mut cursor, "truncated Pls reserved field")?;
        let rgb = &bytes[cursor..];
        let settings = if reserved == 0 {
            DevModeW::from_bytes(rgb)
                .map(PrinterSettings::WindowsUnicode)
                .unwrap_or_else(|| PrinterSettings::PlatformSpecific(rgb.to_vec()))
        } else if usize::from(reserved) == rgb.len() {
            DevModeW::from_bytes(rgb)
                .map(|devmode| PrinterSettings::LengthPrefixedWindowsUnicode {
                    declared_length: reserved,
                    devmode,
                })
                .unwrap_or_else(|| PrinterSettings::PlatformSpecific(rgb.to_vec()))
        } else if reserved == 0x1003 && rgb.starts_with(b"<?xml") {
            PrinterSettings::MacXmlPlist(rgb.to_vec())
        } else if reserved == 1 && rgb.len() == 120 {
            PrinterSettings::MacPrintRecord(rgb.try_into().expect("120 bytes"))
        } else if reserved == 0x2003 && rgb.len() == 28 {
            let mut values = [0u32; 7];
            for (index, value) in values.iter_mut().enumerate() {
                let offset = index * 4;
                *value =
                    u32::from_le_bytes(rgb[offset..offset + 4].try_into().expect("four bytes"));
            }
            PrinterSettings::LegacyPageLayout(values)
        } else {
            PrinterSettings::PlatformSpecific(rgb.to_vec())
        };
        Ok(Self {
            reserved,
            settings,
            physical_segment_lengths,
        })
    }

    fn logical_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = self.reserved.to_le_bytes().to_vec();
        match &self.settings {
            PrinterSettings::WindowsUnicode(value) => bytes.extend_from_slice(&value.to_bytes()?),
            PrinterSettings::LengthPrefixedWindowsUnicode {
                declared_length,
                devmode,
            } => {
                if *declared_length != self.reserved {
                    return Err(Error::invalid(0, "Pls DEVMODE length prefix mismatch"));
                }
                bytes.extend_from_slice(&devmode.to_bytes()?);
            }
            PrinterSettings::MacXmlPlist(payload) => bytes.extend_from_slice(payload),
            PrinterSettings::MacPrintRecord(payload) => bytes.extend_from_slice(payload),
            PrinterSettings::LegacyPageLayout(values) => {
                for value in values {
                    bytes.extend_from_slice(&value.to_le_bytes());
                }
            }
            PrinterSettings::PlatformSpecific(payload) => bytes.extend_from_slice(payload),
        }
        Ok(bytes)
    }

    fn encode_physical(&self) -> Result<Vec<EncodedBiffRecord>> {
        let logical = self.logical_bytes()?;
        let declared_len = self
            .physical_segment_lengths
            .iter()
            .map(|length| usize::from(*length))
            .sum::<usize>();
        if self.physical_segment_lengths.is_empty() || declared_len != logical.len() {
            return Err(Error::invalid(0, "Pls physical layout mismatch"));
        }
        let mut offset = 0usize;
        let mut records = Vec::with_capacity(self.physical_segment_lengths.len());
        for (index, length) in self.physical_segment_lengths.iter().enumerate() {
            let end = offset + usize::from(*length);
            records.push(EncodedBiffRecord {
                record_type: if index == 0 { PLS } else { CONTINUE },
                payload: logical[offset..end].to_vec(),
            });
            offset = end;
        }
        Ok(records)
    }
}

impl BkHimRecord {
    fn from_sequence(
        first: &[u8],
        records: &[BiffRecord],
        limits: Limits,
    ) -> Result<(Self, usize)> {
        if first.len() < 8 {
            return Err(Error::invalid(0, "truncated BkHim header"));
        }
        let image_format = i16::from_le_bytes(first[0..2].try_into().expect("two bytes"));
        let reserved = u16::from_le_bytes(first[2..4].try_into().expect("two bytes"));
        if reserved != 1 {
            return Err(Error::invalid(2, "BkHim reserved field must be 1"));
        }
        let declared_size = i32::from_le_bytes(first[4..8].try_into().expect("four bytes"));
        let image_size = usize::try_from(declared_size)
            .map_err(|_| Error::invalid(4, "BkHim image size must be positive"))?;
        if image_size == 0 {
            return Err(Error::invalid(4, "BkHim image size must be positive"));
        }
        if image_size > limits.max_allocation {
            return Err(Error::Limit(format!(
                "BkHim image size exceeds {}",
                limits.max_allocation
            )));
        }
        let logical_size = image_size
            .checked_add(8)
            .ok_or_else(|| Error::Limit("BkHim logical size overflow".into()))?;
        if first.len() > logical_size {
            return Err(Error::invalid(
                0,
                "BkHim payload exceeds its declared image size",
            ));
        }

        let mut logical = Vec::with_capacity(logical_size);
        logical.extend_from_slice(first);
        let mut physical_segment_lengths = vec![
            u16::try_from(first.len())
                .map_err(|_| Error::Limit("BkHim physical segment exceeds u16".into()))?,
        ];
        let mut consumed = 0usize;
        while logical.len() < logical_size {
            let payload = continue_payload(records, consumed, "BkHim")?;
            let next_size = logical
                .len()
                .checked_add(payload.len())
                .ok_or_else(|| Error::Limit("BkHim logical size overflow".into()))?;
            if next_size > logical_size {
                return Err(Error::invalid(
                    0,
                    "BkHim Continue exceeds its declared image size",
                ));
            }
            logical.extend_from_slice(payload);
            physical_segment_lengths.push(
                u16::try_from(payload.len())
                    .map_err(|_| Error::Limit("BkHim physical segment exceeds u16".into()))?,
            );
            consumed += 1;
        }

        let image_blob = &logical[8..];
        let image = match image_format {
            0x0009 => BkHimImage::Bitmap(
                DeviceIndependentBitmap::from_packed_slice(image_blob, DibColorUsage::RgbColors)
                    .map_err(|error| Error::invalid(8, format!("invalid BkHim bitmap: {error}")))?,
            ),
            0x000e => BkHimImage::Native(image_blob.to_vec()),
            _ => return Err(Error::invalid(0, "unsupported BkHim image format")),
        };
        Ok((
            Self {
                reserved,
                image,
                physical_segment_lengths,
            },
            consumed,
        ))
    }

    fn logical_bytes(&self) -> Result<Vec<u8>> {
        if self.reserved != 1 {
            return Err(Error::invalid(2, "BkHim reserved field must be 1"));
        }
        let (image_format, image_blob) = match &self.image {
            BkHimImage::Bitmap(bitmap) => (
                0x0009i16,
                bitmap
                    .to_packed_bytes()
                    .map_err(|error| Error::invalid(8, format!("invalid BkHim bitmap: {error}")))?,
            ),
            BkHimImage::Native(bytes) => (0x000ei16, bytes.clone()),
        };
        if image_blob.is_empty() {
            return Err(Error::invalid(4, "BkHim image must not be empty"));
        }
        let declared_size = i32::try_from(image_blob.len())
            .map_err(|_| Error::Limit("BkHim image size exceeds i32".into()))?;
        let mut bytes = Vec::with_capacity(image_blob.len() + 8);
        bytes.extend_from_slice(&image_format.to_le_bytes());
        bytes.extend_from_slice(&self.reserved.to_le_bytes());
        bytes.extend_from_slice(&declared_size.to_le_bytes());
        bytes.extend_from_slice(&image_blob);
        Ok(bytes)
    }

    fn encode_physical(&self, first_record_type: u16) -> Result<Vec<EncodedBiffRecord>> {
        let logical = self.logical_bytes()?;
        let physical_size =
            self.physical_segment_lengths
                .iter()
                .try_fold(0usize, |total, length| {
                    total
                        .checked_add(usize::from(*length))
                        .ok_or_else(|| Error::Limit("BkHim physical layout size overflow".into()))
                })?;
        if self.physical_segment_lengths.is_empty() || physical_size != logical.len() {
            return Err(Error::invalid(0, "BkHim physical layout mismatch"));
        }
        let mut offset = 0usize;
        self.physical_segment_lengths
            .iter()
            .enumerate()
            .map(|(index, length)| {
                let end = offset + usize::from(*length);
                let encoded = EncodedBiffRecord {
                    record_type: if index == 0 {
                        first_record_type
                    } else {
                        CONTINUE
                    },
                    payload: logical[offset..end].to_vec(),
                };
                offset = end;
                Ok(encoded)
            })
            .collect()
    }
}

impl RtdTopicString {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let declared_unit_count = reader.read_u32()?;
        let unit_count = usize::try_from(declared_unit_count)
            .map_err(|_| Error::Limit("RTD topic unit count exceeds usize".into()))?;
        let flags = reader.read_u8()?;
        if flags & !1 != 0 {
            return Err(Error::invalid(
                reader.position()?.saturating_sub(1),
                "RTD topic has nonzero reserved flags",
            ));
        }
        let mut consumed = 0usize;
        let mut substrings = Vec::new();
        while consumed < unit_count {
            let count = if flags & 1 == 0 {
                usize::from(reader.read_u8()?)
            } else {
                usize::from(reader.read_u16()?)
            };
            consumed = consumed
                .checked_add(1 + count)
                .ok_or_else(|| Error::Limit("RTD topic unit count overflow".into()))?;
            if consumed > unit_count {
                return Err(Error::invalid(
                    reader.position()?,
                    "RTD topic substring exceeds declared unit count",
                ));
            }
            let characters = if flags & 1 == 0 {
                XlStringCharacters::Compressed(reader.read_vec(count)?)
            } else {
                let byte_count = count
                    .checked_mul(2)
                    .ok_or_else(|| Error::Limit("RTD topic byte count overflow".into()))?;
                let bytes = reader.read_vec(byte_count)?;
                XlStringCharacters::Unicode(
                    bytes
                        .chunks_exact(2)
                        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                        .collect(),
                )
            };
            substrings.push(characters);
        }
        Ok(Self {
            declared_unit_count,
            flags,
            substrings,
        })
    }

    fn write<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.flags & !1 != 0 {
            return Err(Error::invalid(0, "RTD topic has nonzero reserved flags"));
        }
        let mut actual_units = 0usize;
        for substring in &self.substrings {
            let count = match (self.flags & 1, substring) {
                (0, XlStringCharacters::Compressed(values)) => values.len(),
                (1, XlStringCharacters::Unicode(values)) => values.len(),
                _ => return Err(Error::invalid(0, "RTD topic encoding mismatch")),
            };
            actual_units = actual_units
                .checked_add(1 + count)
                .ok_or_else(|| Error::Limit("RTD topic unit count overflow".into()))?;
        }
        if usize::try_from(self.declared_unit_count).ok() != Some(actual_units) {
            return Err(Error::invalid(0, "RTD topic unit count mismatch"));
        }
        writer.write_u32(self.declared_unit_count)?;
        writer.write_u8(self.flags)?;
        for substring in &self.substrings {
            match substring {
                XlStringCharacters::Compressed(values) => {
                    writer.write_u8(u8::try_from(values.len()).map_err(|_| {
                        Error::Limit("compressed RTD substring exceeds u8".into())
                    })?)?;
                    writer.write_all(values)?;
                }
                XlStringCharacters::Unicode(values) => {
                    writer.write_u16(
                        u16::try_from(values.len()).map_err(|_| {
                            Error::Limit("Unicode RTD substring exceeds u16".into())
                        })?,
                    )?;
                    for value in values {
                        writer.write_u16(*value)?;
                    }
                }
            }
        }
        Ok(())
    }
}

impl SdkRead for RealTimeDataRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        if header.record_type != REAL_TIME_DATA {
            return Err(Error::invalid(0, "RealTimeData FRT record type mismatch"));
        }
        let shared_prefix_character_count = reader.read_u32()?;
        let topic = RtdTopicString::read(reader)?;
        let discriminator = reader.read_u32()?;
        let operation = match discriminator {
            0x0000_0001 => RtdOperation::Number {
                bits: reader.read_u64()?,
            },
            0x0000_0002 | 0x0000_1000 => {
                let count = usize::try_from(reader.read_u32()?)
                    .map_err(|_| Error::Limit("RTD operation string exceeds usize".into()))?;
                let value = BiffUnicodeString::read(reader, count)?;
                if discriminator == 2 {
                    RtdOperation::ShortString(value)
                } else {
                    RtdOperation::LongString(value)
                }
            }
            0x0000_0004 => RtdOperation::Boolean(reader.read_u32()?),
            0x0000_0010 => RtdOperation::Error(reader.read_i32()?),
            value
                if value & 0x0000_00ff == 0x10
                    && reader.remaining()? >= 4
                    && (reader.remaining()? - 4).is_multiple_of(6) =>
            {
                RtdOperation::ErrorWithCorruptDiscriminator {
                    discriminator: value,
                    value: reader.read_i32()?,
                }
            }
            0x0000_0800 => RtdOperation::Integer(reader.read_i32()?),
            _ => {
                let length = usize::try_from(reader.remaining()?)
                    .map_err(|_| Error::Limit("malformed RTD operation exceeds usize".into()))?;
                return Ok(Self {
                    header,
                    shared_prefix_character_count,
                    topic,
                    operation: RtdOperation::Malformed {
                        discriminator,
                        payload: reader.read_vec(length)?,
                    },
                    cells: Vec::new(),
                });
            }
        };
        let remaining = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("RTD cell array exceeds usize".into()))?;
        if remaining % 6 != 0 {
            return Err(Error::invalid(
                reader.position()?,
                "RTD cell array length is not divisible by 6",
            ));
        }
        let mut cells = Vec::with_capacity(remaining / 6);
        for _ in 0..remaining / 6 {
            cells.push(RtdCellReference::read_from(reader)?);
        }
        Ok(Self {
            header,
            shared_prefix_character_count,
            topic,
            operation,
            cells,
        })
    }
}

impl SdkWrite for RealTimeDataRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.header.record_type != REAL_TIME_DATA {
            return Err(Error::invalid(0, "RealTimeData FRT record type mismatch"));
        }
        self.header.write_to(writer)?;
        writer.write_u32(self.shared_prefix_character_count)?;
        self.topic.write(writer)?;
        match &self.operation {
            RtdOperation::Number { bits } => {
                writer.write_u32(1)?;
                writer.write_u64(*bits)?;
            }
            RtdOperation::ShortString(value) | RtdOperation::LongString(value) => {
                writer.write_u32(if matches!(self.operation, RtdOperation::ShortString(_)) {
                    2
                } else {
                    0x1000
                })?;
                writer.write_u32(u32::try_from(value.character_count()).map_err(|_| {
                    Error::Limit("RTD operation string character count exceeds u32".into())
                })?)?;
                value.write(writer)?;
            }
            RtdOperation::Boolean(value) => {
                writer.write_u32(4)?;
                writer.write_u32(*value)?;
            }
            RtdOperation::Error(value) => {
                writer.write_u32(0x10)?;
                writer.write_i32(*value)?;
            }
            RtdOperation::ErrorWithCorruptDiscriminator {
                discriminator,
                value,
            } => {
                if discriminator & 0x0000_00ff != 0x10 {
                    return Err(Error::invalid(
                        0,
                        "RTD corrupt error discriminator lost its Error low byte",
                    ));
                }
                writer.write_u32(*discriminator)?;
                writer.write_i32(*value)?;
            }
            RtdOperation::Integer(value) => {
                writer.write_u32(0x800)?;
                writer.write_i32(*value)?;
            }
            RtdOperation::Malformed {
                discriminator,
                payload,
            } => {
                if !self.cells.is_empty() {
                    return Err(Error::invalid(
                        0,
                        "malformed RTD operation has cell entries",
                    ));
                }
                writer.write_u32(*discriminator)?;
                writer.write_all(payload)?;
                return Ok(());
            }
        }
        for cell in &self.cells {
            cell.write_to(writer)?;
        }
        Ok(())
    }
}

impl SortOptions {
    fn from_bits(bits: u16) -> Self {
        let raw_order = ((bits >> 5) & 0x1f) as i8;
        Self {
            sort_columns: bits & 0x0001 != 0,
            descending: [bits & 0x0002 != 0, bits & 0x0004 != 0, bits & 0x0008 != 0],
            case_sensitive: bits & 0x0010 != 0,
            custom_list_index: if raw_order & 0x10 != 0 {
                raw_order - 0x20
            } else {
                raw_order
            },
            alternate_method: bits & 0x0400 != 0,
            reserved: (bits >> 11) as u8,
        }
    }

    fn bits(self) -> Result<u16> {
        if !(-16..=15).contains(&self.custom_list_index) {
            return Err(Error::invalid(
                0,
                "Sort custom-list index exceeds signed 5-bit range",
            ));
        }
        if self.reserved > 0x1f {
            return Err(Error::invalid(0, "Sort reserved flags exceed 5 bits"));
        }
        let mut bits = u16::from(self.sort_columns)
            | (u16::from(self.descending[0]) << 1)
            | (u16::from(self.descending[1]) << 2)
            | (u16::from(self.descending[2]) << 3)
            | (u16::from(self.case_sensitive) << 4)
            | (u16::from(self.alternate_method) << 10)
            | (u16::from(self.reserved) << 11);
        bits |= (i16::from(self.custom_list_index) as u16 & 0x1f) << 5;
        Ok(bits)
    }
}

impl SdkRead for SortRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let options = SortOptions::from_bits(reader.read_u16()?);
        let counts = [reader.read_u8()?, reader.read_u8()?, reader.read_u8()?];
        let mut keys = [None, None, None];
        for (key, count) in keys.iter_mut().zip(counts) {
            if count != 0 {
                *key = Some(BiffUnicodeString::read(reader, usize::from(count))?);
            }
        }
        Ok(Self {
            options,
            keys,
            reserved: reader.read_u8()?,
        })
    }
}

impl SdkWrite for SortRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.options.bits()?)?;
        for key in &self.keys {
            let count = key.as_ref().map_or(0, BiffUnicodeString::character_count);
            writer.write_u8(
                u8::try_from(count)
                    .map_err(|_| Error::Limit("Sort key character count exceeds u8".into()))?,
            )?;
        }
        for value in self.keys.iter().flatten() {
            value.write(writer)?;
        }
        writer.write_u8(self.reserved)?;
        Ok(())
    }
}

impl SortDataOptions {
    fn from_bits(bits: u16) -> Result<Self> {
        let parent = match (bits >> 3) & 0x7 {
            0 => SortFieldParent::Sheet,
            1 => SortFieldParent::Table,
            2 => SortFieldParent::AutoFilter,
            3 => SortFieldParent::QueryTable,
            _ => return Err(Error::invalid(0, "SortData has an invalid parent kind")),
        };
        Ok(Self {
            sort_columns: bits & 1 != 0,
            case_sensitive: bits & 2 != 0,
            alternate_method: bits & 4 != 0,
            parent,
            unused: bits >> 6,
        })
    }

    fn bits(self) -> Result<u16> {
        if self.unused > 0x03ff {
            return Err(Error::invalid(0, "SortData unused flags exceed 10 bits"));
        }
        let parent = match self.parent {
            SortFieldParent::Sheet => 0,
            SortFieldParent::Table => 1,
            SortFieldParent::AutoFilter => 2,
            SortFieldParent::QueryTable => 3,
        };
        Ok(u16::from(self.sort_columns)
            | (u16::from(self.case_sensitive) << 1)
            | (u16::from(self.alternate_method) << 2)
            | (parent << 3)
            | (self.unused << 6))
    }
}

impl SdkRead for SortDataRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        if header.record_type != SORT_DATA {
            return Err(Error::invalid(0, "SortData FRT record type mismatch"));
        }
        Ok(Self {
            header,
            options: SortDataOptions::from_bits(reader.read_u16()?)?,
            range: Rfx::read_from(reader)?,
            condition_count: reader.read_u32()?,
            parent_id: reader.read_u32()?,
            conditions: Vec::new(),
        })
    }
}

impl SdkWrite for SortDataRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.header.record_type != SORT_DATA {
            return Err(Error::invalid(0, "SortData FRT record type mismatch"));
        }
        self.header.write_to(writer)?;
        writer.write_u16(self.options.bits()?)?;
        self.range.write_to(writer)?;
        writer.write_u32(self.condition_count)?;
        writer.write_u32(self.parent_id)
    }
}

impl SdkRead for SortCondition {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let flags = reader.read_u16()?;
        let sort_on = ((flags >> 1) & 0x0f) as u8;
        let range = Rfx::read_from(reader)?;
        let data = match sort_on {
            0 => SortConditionData::Value {
                value: reader.read_u32()?,
                reserved: reader.read_u32()?,
            },
            1 => SortConditionData::CellColor {
                dxf_index: reader.read_u32()?,
                reserved: reader.read_u32()?,
            },
            2 => SortConditionData::FontColor {
                dxf_index: reader.read_u32()?,
                reserved: reader.read_u32()?,
            },
            3 => SortConditionData::Icon {
                icon_set: reader.read_u32()?,
                icon_index: reader.read_i32()?,
            },
            _ => return Err(Error::invalid(0, "SortCond12 has an invalid sort-on value")),
        };
        let count = reader.read_i32()?;
        if count < 0 {
            return Err(Error::invalid(
                0,
                "SortCond12 custom-list length is negative",
            ));
        }
        let custom_list = if count == 0 {
            None
        } else {
            Some(BiffUnicodeString::read(
                reader,
                usize::try_from(count)
                    .map_err(|_| Error::Limit("SortCond12 custom list exceeds usize".into()))?,
            )?)
        };
        Ok(Self {
            descending: flags & 1 != 0,
            reserved: flags >> 5,
            range,
            data,
            custom_list,
        })
    }
}

impl SdkWrite for SortCondition {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.reserved > 0x07ff {
            return Err(Error::invalid(
                0,
                "SortCond12 reserved flags exceed 11 bits",
            ));
        }
        let sort_on = match self.data {
            SortConditionData::Value { .. } => 0u16,
            SortConditionData::CellColor { .. } => 1,
            SortConditionData::FontColor { .. } => 2,
            SortConditionData::Icon { .. } => 3,
        };
        if sort_on != 0 && self.custom_list.is_some() {
            return Err(Error::invalid(
                0,
                "non-value SortCond12 has a custom sort list",
            ));
        }
        writer.write_u16(u16::from(self.descending) | (sort_on << 1) | (self.reserved << 5))?;
        self.range.write_to(writer)?;
        match self.data {
            SortConditionData::Value { value, reserved } => {
                writer.write_u32(value)?;
                writer.write_u32(reserved)?;
            }
            SortConditionData::CellColor {
                dxf_index,
                reserved,
            }
            | SortConditionData::FontColor {
                dxf_index,
                reserved,
            } => {
                writer.write_u32(dxf_index)?;
                writer.write_u32(reserved)?;
            }
            SortConditionData::Icon {
                icon_set,
                icon_index,
            } => {
                writer.write_u32(icon_set)?;
                writer.write_i32(icon_index)?;
            }
        }
        let count = self
            .custom_list
            .as_ref()
            .map_or(0, BiffUnicodeString::character_count);
        writer.write_i32(
            i32::try_from(count)
                .map_err(|_| Error::Limit("SortCond12 custom list exceeds i32".into()))?,
        )?;
        if let Some(value) = &self.custom_list {
            value.write(writer)?;
        }
        Ok(())
    }
}

impl SortDataRecord {
    fn encode_physical(&self) -> Result<Vec<EncodedBiffRecord>> {
        if usize::try_from(self.condition_count).ok() != Some(self.conditions.len()) {
            return Err(Error::invalid(0, "SortData condition count mismatch"));
        }
        let mut records = vec![EncodedBiffRecord {
            record_type: SORT_DATA,
            payload: encode_sdk(self)?,
        }];
        for continuation in &self.conditions {
            if continuation.header.record_type != CONTINUE_FRT12 {
                return Err(Error::invalid(
                    0,
                    "SortCond12 continuation record type mismatch",
                ));
            }
            let mut writer = Writer::new(Cursor::new(Vec::new()));
            continuation.header.write_to(&mut writer)?;
            continuation.condition.write_to(&mut writer)?;
            records.push(EncodedBiffRecord {
                record_type: CONTINUE_FRT12,
                payload: writer.into_inner().into_inner(),
            });
        }
        Ok(records)
    }
}

impl AutoFilterOptions {
    fn from_bits(bits: u16) -> Result<Self> {
        if bits & 0x0003 > 1 {
            return Err(Error::invalid(0, "AutoFilter has an invalid join value"));
        }
        Ok(Self {
            join_or: bits & 1 != 0,
            simple: [bits & 0x0004 != 0, bits & 0x0008 != 0],
            top_n: bits & 0x0010 != 0,
            top: bits & 0x0020 != 0,
            percent: bits & 0x0040 != 0,
            top_count: bits >> 7,
        })
    }

    fn bits(self) -> Result<u16> {
        if self.top_count > 0x01ff {
            return Err(Error::invalid(0, "AutoFilter Top-N count exceeds 9 bits"));
        }
        Ok(u16::from(self.join_or)
            | (u16::from(self.simple[0]) << 2)
            | (u16::from(self.simple[1]) << 3)
            | (u16::from(self.top_n) << 4)
            | (u16::from(self.top) << 5)
            | (u16::from(self.percent) << 6)
            | (self.top_count << 7))
    }
}

impl AutoFilterOperand {
    fn read_fixed<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let value_type = reader.read_u8()?;
        let comparison = reader.read_u8()?;
        let value = match value_type {
            0x00 => AutoFilterOperandValue::Unused {
                bytes: reader.read_vec(8)?.try_into().expect("exactly eight bytes"),
            },
            0x02 => AutoFilterOperandValue::Rk {
                value: reader.read_u32()?,
                unused: reader.read_u32()?,
            },
            0x04 => AutoFilterOperandValue::Number {
                bits: reader.read_u64()?,
            },
            0x06 => AutoFilterOperandValue::String {
                unused1: reader.read_u32()?,
                declared_character_count: reader.read_u8()?,
                compare_without_wildcards: reader.read_u8()?,
                reserved: reader.read_u8()?,
                unused2: reader.read_u8()?,
            },
            0x08 => AutoFilterOperandValue::BooleanOrError {
                value: reader.read_u16()?,
                unused1: reader.read_u16()?,
                unused2: reader.read_u32()?,
            },
            0x0c => AutoFilterOperandValue::Blanks {
                reserved: reader.read_u64()?,
            },
            0x0e => AutoFilterOperandValue::NonBlanks {
                reserved: reader.read_u64()?,
            },
            _ => return Err(Error::invalid(0, "AutoFilter has an invalid operand type")),
        };
        Ok(Self {
            comparison,
            value,
            string: None,
        })
    }

    fn write_fixed<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let value_type = match self.value {
            AutoFilterOperandValue::Unused { .. } => 0x00,
            AutoFilterOperandValue::Rk { .. } => 0x02,
            AutoFilterOperandValue::Number { .. } => 0x04,
            AutoFilterOperandValue::String { .. } => 0x06,
            AutoFilterOperandValue::BooleanOrError { .. } => 0x08,
            AutoFilterOperandValue::Blanks { .. } => 0x0c,
            AutoFilterOperandValue::NonBlanks { .. } => 0x0e,
        };
        writer.write_u8(value_type)?;
        writer.write_u8(self.comparison)?;
        match self.value {
            AutoFilterOperandValue::Unused { bytes } => writer.write_all(&bytes)?,
            AutoFilterOperandValue::Rk { value, unused } => {
                writer.write_u32(value)?;
                writer.write_u32(unused)?;
            }
            AutoFilterOperandValue::Number { bits } => writer.write_u64(bits)?,
            AutoFilterOperandValue::String {
                unused1,
                declared_character_count,
                compare_without_wildcards,
                reserved,
                unused2,
            } => {
                let string = self
                    .string
                    .as_ref()
                    .ok_or_else(|| Error::invalid(0, "AutoFilter string operand lacks text"))?;
                if string.character_count() != usize::from(declared_character_count) {
                    return Err(Error::invalid(0, "AutoFilter string length mismatch"));
                }
                writer.write_u32(unused1)?;
                writer.write_u8(declared_character_count)?;
                writer.write_u8(compare_without_wildcards)?;
                writer.write_u8(reserved)?;
                writer.write_u8(unused2)?;
            }
            AutoFilterOperandValue::BooleanOrError {
                value,
                unused1,
                unused2,
            } => {
                writer.write_u16(value)?;
                writer.write_u16(unused1)?;
                writer.write_u32(unused2)?;
            }
            AutoFilterOperandValue::Blanks { reserved }
            | AutoFilterOperandValue::NonBlanks { reserved } => writer.write_u64(reserved)?,
        }
        if !matches!(self.value, AutoFilterOperandValue::String { .. }) && self.string.is_some() {
            return Err(Error::invalid(0, "non-string AutoFilter operand has text"));
        }
        Ok(())
    }
}

impl SdkRead for AutoFilterRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let entry_index = reader.read_u16()?;
        let options = AutoFilterOptions::from_bits(reader.read_u16()?)?;
        let mut operands = [
            AutoFilterOperand::read_fixed(reader)?,
            AutoFilterOperand::read_fixed(reader)?,
        ];
        for operand in &mut operands {
            if let AutoFilterOperandValue::String {
                declared_character_count,
                ..
            } = operand.value
            {
                if declared_character_count == 0 {
                    return Err(Error::invalid(0, "AutoFilter string length is zero"));
                }
                operand.string = Some(BiffUnicodeString::read(
                    reader,
                    usize::from(declared_character_count),
                )?);
            }
        }
        Ok(Self {
            entry_index,
            options,
            operands,
        })
    }
}

impl SdkWrite for AutoFilterRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.entry_index)?;
        writer.write_u16(self.options.bits()?)?;
        for operand in &self.operands {
            operand.write_fixed(writer)?;
        }
        for operand in &self.operands {
            if let Some(value) = &operand.string {
                value.write(writer)?;
            }
        }
        Ok(())
    }
}

impl SdkRead for SxFormatRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let flags = reader.read_u16()?;
        let formatting_applied = match flags & 0x000f {
            0 => false,
            1 => true,
            _ => return Err(Error::invalid(0, "SxFormat has an invalid rule type")),
        };
        Ok(Self {
            formatting_applied,
            reserved: flags >> 4,
            differential_format_byte_count: reader.read_u16()?,
        })
    }
}

impl SdkWrite for SxFormatRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.reserved > 0x0fff {
            return Err(Error::invalid(0, "SxFormat reserved flags exceed 12 bits"));
        }
        if !self.formatting_applied && self.differential_format_byte_count != 0 {
            return Err(Error::invalid(
                0,
                "cleared SxFormat has a nonzero differential-format size",
            ));
        }
        writer.write_u16(u16::from(self.formatting_applied) | (self.reserved << 4))?;
        writer.write_u16(self.differential_format_byte_count)
    }
}

impl SdkRead for WOptRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeaderOld::read_from(reader)?;
        if header.record_type != W_OPT {
            return Err(Error::invalid(0, "WOpt FRT record type mismatch"));
        }
        let flags = WOptFlags::from_bits_retain(reader.read_u16()?);
        let screen_size = WebScreenSize::read_from(reader)?;
        let reserved = reader.read_u8()?;
        let pixels_per_inch = reader.read_u32()?;
        let code_page = reader.read_u32()?;
        let component_location = LpWideString::read_from(reader)?;
        let future_length = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("WOpt future data exceeds usize".into()))?;
        Ok(Self {
            header,
            flags,
            screen_size,
            reserved,
            pixels_per_inch,
            code_page,
            component_location,
            future: reader.read_vec(future_length)?,
        })
    }
}

impl SdkWrite for WOptRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.header.record_type != W_OPT {
            return Err(Error::invalid(0, "WOpt FRT record type mismatch"));
        }
        self.header.write_to(writer)?;
        writer.write_u16(self.flags.bits())?;
        self.screen_size.write_to(writer)?;
        writer.write_u8(self.reserved)?;
        writer.write_u32(self.pixels_per_inch)?;
        writer.write_u32(self.code_page)?;
        self.component_location.write_to(writer)?;
        writer.write_all(&self.future)?;
        Ok(())
    }
}

impl TableOptions {
    fn from_bits(bits: u16) -> Self {
        Self {
            always_calculate: bits & 0x0001 != 0,
            reserved1: bits & 0x0002 != 0,
            row_input: bits & 0x0004 != 0,
            two_variable: bits & 0x0008 != 0,
            first_input_deleted: bits & 0x0010 != 0,
            second_input_deleted: bits & 0x0020 != 0,
            reserved2: bits >> 6,
        }
    }

    fn bits(self) -> Result<u16> {
        if self.reserved2 > 0x03ff {
            return Err(Error::invalid(0, "Table reserved flags exceed 10 bits"));
        }
        Ok(u16::from(self.always_calculate)
            | (u16::from(self.reserved1) << 1)
            | (u16::from(self.row_input) << 2)
            | (u16::from(self.two_variable) << 3)
            | (u16::from(self.first_input_deleted) << 4)
            | (u16::from(self.second_input_deleted) << 5)
            | (self.reserved2 << 6))
    }
}

impl SdkRead for TableRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let range = TableRange::read_from(reader)?;
        let options = TableOptions::from_bits(reader.read_u16()?);
        let row_input = TableInputReference::read_from(reader)?;
        let column_input = TableInputReference::read_from(reader)?;
        let compatibility_padding = match reader.remaining()? {
            0 => None,
            2 => Some(reader.read_u16()?),
            _ => return Err(Error::invalid(0, "Table has invalid trailing data")),
        };
        Ok(Self {
            range,
            options,
            row_input,
            column_input,
            compatibility_padding,
        })
    }
}

impl SdkWrite for TableRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.range.write_to(writer)?;
        writer.write_u16(self.options.bits()?)?;
        self.row_input.write_to(writer)?;
        self.column_input.write_to(writer)?;
        if let Some(value) = self.compatibility_padding {
            writer.write_u16(value)?;
        }
        Ok(())
    }
}

impl SdkRead for CrtCoOptRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let header = FrtHeader::read_from(reader)?;
        let color_scheme = reader.read_u32()?;
        let flags = CrtCoOptFlags::from_bits_retain(reader.read_u16()?);
        let compatibility_padding = match reader.remaining()? {
            0 => None,
            2 => Some(reader.read_u16()?),
            _ => return Err(Error::invalid(0, "CrtCoOpt has invalid trailing data")),
        };
        Ok(Self {
            header,
            color_scheme,
            flags,
            compatibility_padding,
        })
    }
}

impl SdkWrite for CrtCoOptRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.header.write_to(writer)?;
        writer.write_u32(self.color_scheme)?;
        writer.write_u16(self.flags.bits())?;
        if let Some(value) = self.compatibility_padding {
            writer.write_u16(value)?;
        }
        Ok(())
    }
}

impl SdkRead for LhGraphViewRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let core = LhGraphViewCore::read_from(reader)?;
        let compatibility_extension = match reader.remaining()? {
            0 => None,
            6 => Some([reader.read_u16()?, reader.read_u16()?, reader.read_u16()?]),
            _ => {
                return Err(Error::invalid(
                    reader.position()?,
                    "LH graph view has an invalid compatibility extension",
                ));
            }
        };
        Ok(Self {
            core,
            compatibility_extension,
        })
    }
}

impl SdkWrite for LhGraphViewRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.core.write_to(writer)?;
        if let Some(words) = self.compatibility_extension {
            for word in words {
                writer.write_u16(word)?;
            }
        }
        Ok(())
    }
}

fn read_lh_reserved<R: Read + Seek>(
    reader: &mut Reader<R>,
    kind: LhReservedKind,
) -> Result<LhSubrecordData> {
    if reader.remaining()? % 2 != 0 {
        return Err(Error::invalid(
            reader.position()?,
            "LH reserved subrecord has an odd byte count",
        ));
    }
    let mut words = Vec::new();
    while reader.remaining()? != 0 {
        words.push(reader.read_u16()?);
    }
    Ok(LhSubrecordData::Reserved { kind, words })
}

impl SdkRead for LhRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let mut subrecords = Vec::new();
        while reader.remaining()? != 0 {
            if reader.remaining()? < 4 {
                return Err(Error::invalid(
                    reader.position()?,
                    "truncated LH subrecord header",
                ));
            }
            let subrecord_type = reader.read_u16()?;
            let byte_count = u64::from(reader.read_u16()?);
            let data = {
                let mut subreader = reader.sub_reader(byte_count)?;
                let data = match subrecord_type {
                    0x0001 => read_lh_reserved(&mut subreader, LhReservedKind::Type1)?,
                    0x0002 => {
                        let len = usize::try_from(subreader.remaining()?)
                            .map_err(|_| Error::Limit("LH header string is too large".into()))?;
                        LhSubrecordData::HeaderString(subreader.read_vec(len)?)
                    }
                    0x0003 => {
                        let len = usize::try_from(subreader.remaining()?)
                            .map_err(|_| Error::Limit("LH footer string is too large".into()))?;
                        LhSubrecordData::FooterString(subreader.read_vec(len)?)
                    }
                    0x0004..=0x0007 => LhSubrecordData::Margin {
                        kind: match subrecord_type {
                            0x0004 => LhMarginKind::Left,
                            0x0005 => LhMarginKind::Right,
                            0x0006 => LhMarginKind::Top,
                            _ => LhMarginKind::Bottom,
                        },
                        value_bits: subreader.read_u64()?,
                    },
                    0x0008 => LhSubrecordData::GraphView(Box::new(LhGraphViewRecord::read_from(
                        &mut subreader,
                    )?)),
                    0x0009 => LhSubrecordData::GlobalColumnWidth(subreader.read_u16()?),
                    0x000a => read_lh_reserved(&mut subreader, LhReservedKind::Type10)?,
                    0x000b => LhSubrecordData::TableType(LhTableType::read_from(&mut subreader)?),
                    0x000c => read_lh_reserved(&mut subreader, LhReservedKind::Type12)?,
                    0x000d => LhSubrecordData::UndocumentedType13(subreader.read_u16()?),
                    _ => {
                        return Err(Error::invalid(
                            subreader.position()?,
                            format!("unsupported LH subrecord type 0x{subrecord_type:04x}"),
                        ));
                    }
                };
                if subreader.remaining()? != 0 {
                    return Err(Error::invalid(
                        subreader.position()?,
                        format!("LH subrecord 0x{subrecord_type:04x} has trailing data"),
                    ));
                }
                data
            };
            subrecords.push(data);
        }
        Ok(Self { subrecords })
    }
}

impl LhSubrecordData {
    fn encode(&self) -> Result<(u16, Vec<u8>)> {
        Ok(match self {
            Self::HeaderString(bytes) => (0x0002, bytes.clone()),
            Self::FooterString(bytes) => (0x0003, bytes.clone()),
            Self::Margin { kind, value_bits } => (
                match kind {
                    LhMarginKind::Left => 0x0004,
                    LhMarginKind::Right => 0x0005,
                    LhMarginKind::Top => 0x0006,
                    LhMarginKind::Bottom => 0x0007,
                },
                value_bits.to_le_bytes().to_vec(),
            ),
            Self::GraphView(value) => (0x0008, encode_sdk(value.as_ref())?),
            Self::GlobalColumnWidth(value) => (0x0009, value.to_le_bytes().to_vec()),
            Self::TableType(value) => (0x000b, encode_sdk(value)?),
            Self::Reserved { kind, words } => {
                let mut bytes = Vec::with_capacity(words.len() * 2);
                for word in words {
                    bytes.extend_from_slice(&word.to_le_bytes());
                }
                (
                    match kind {
                        LhReservedKind::Type1 => 0x0001,
                        LhReservedKind::Type10 => 0x000a,
                        LhReservedKind::Type12 => 0x000c,
                    },
                    bytes,
                )
            }
            Self::UndocumentedType13(value) => (0x000d, value.to_le_bytes().to_vec()),
        })
    }
}

impl SdkWrite for LhRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        for subrecord in &self.subrecords {
            let (subrecord_type, bytes) = subrecord.encode()?;
            writer.write_u16(subrecord_type)?;
            writer.write_u16(
                u16::try_from(bytes.len())
                    .map_err(|_| Error::Limit("LH subrecord exceeds u16 length".into()))?,
            )?;
            writer.write_all(&bytes)?;
        }
        Ok(())
    }
}

impl SdkRead for QsiRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        Ok(Self {
            flags: QsiFlags::from_bits_retain(reader.read_u16()?),
            auto_format_index: reader.read_u16()?,
            formatting_flags: QsiFormattingFlags::from_bits_retain(reader.read_u16()?),
            reserved: reader.read_u32()?,
            name: XlUnicodeString::read_from(reader)?,
            unused: reader.read_u16()?,
        })
    }
}

impl SdkWrite for QsiRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.flags.bits())?;
        writer.write_u16(self.auto_format_index)?;
        writer.write_u16(self.formatting_flags.bits())?;
        writer.write_u32(self.reserved)?;
        self.name.write_to(writer)?;
        writer.write_u16(self.unused)
    }
}

impl ParamQryFixed {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let sql_type = reader.read_u16()?;
        let flags = reader.read_u16()?;
        let parameter_type = (flags & 0x0003) as u8;
        if parameter_type > 2 {
            return Err(Error::invalid(0, "ParamQry has an invalid parameter type"));
        }
        Ok(Self {
            sql_type,
            parameter_type,
            unused1: flags & 0x0004 != 0,
            non_default_name: flags & 0x0008 != 0,
            unused2: flags >> 4,
            value_type: reader.read_u16()?,
            boolean_value: reader.read_u16()?,
        })
    }

    fn write<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.parameter_type > 2 || self.unused2 > 0x0fff {
            return Err(Error::invalid(
                0,
                "ParamQry fixed flags exceed their bit widths",
            ));
        }
        writer.write_u16(self.sql_type)?;
        writer.write_u16(
            u16::from(self.parameter_type)
                | (u16::from(self.unused1) << 2)
                | (u16::from(self.non_default_name) << 3)
                | (self.unused2 << 4),
        )?;
        writer.write_u16(self.value_type)?;
        writer.write_u16(self.boolean_value)
    }
}

impl SdkRead for ParamQryRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let fixed = ParamQryFixed::read(reader)?;
        let data = match (fixed.parameter_type, fixed.value_type) {
            (0, _) => ParamQryData::Prompt {
                text: SxStringRecord::read_from(reader)?,
                unused: reader.read_u8()?,
            },
            (1, 0x0001) => ParamQryData::Number {
                bits: reader.read_u64()?,
            },
            (1, 0x0002) => ParamQryData::String {
                text: SxStringRecord::read_from(reader)?,
                unused: reader.read_u8()?,
            },
            (1, 0x0004) => ParamQryData::Boolean,
            (1, 0x0800) => ParamQryData::Integer(reader.read_i32()?),
            (2, _) => {
                let length = usize::from(reader.read_u16()?);
                ParamQryData::Reference(FormulaTokenStream::from_bytes(&reader.read_vec(length)?)?)
            }
            _ => return Err(Error::invalid(0, "ParamQry has an invalid value type")),
        };
        Ok(Self { fixed, data })
    }
}

impl SdkWrite for ParamQryRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        let valid = matches!(
            (&self.data, self.fixed.parameter_type, self.fixed.value_type),
            (ParamQryData::Prompt { .. }, 0, _)
                | (ParamQryData::Number { .. }, 1, 0x0001)
                | (ParamQryData::String { .. }, 1, 0x0002)
                | (ParamQryData::Boolean, 1, 0x0004)
                | (ParamQryData::Integer(_), 1, 0x0800)
                | (ParamQryData::Reference(_), 2, _)
        );
        if !valid {
            return Err(Error::invalid(
                0,
                "ParamQry data disagrees with fixed fields",
            ));
        }
        self.fixed.write(writer)?;
        match &self.data {
            ParamQryData::Prompt { text, unused } | ParamQryData::String { text, unused } => {
                text.write_to(writer)?;
                writer.write_u8(*unused)?;
            }
            ParamQryData::Number { bits } => writer.write_u64(*bits)?,
            ParamQryData::Boolean => {}
            ParamQryData::Integer(value) => writer.write_i32(*value)?,
            ParamQryData::Reference(formula) => {
                let bytes = formula.to_bytes()?;
                writer.write_u16(u16::try_from(bytes.len()).map_err(|_| {
                    Error::Limit("ParamQry reference formula exceeds u16".into())
                })?)?;
                writer.write_all(&bytes)?;
            }
        }
        Ok(())
    }
}

impl SxSelectOptions {
    fn from_bits(bits: u16) -> Self {
        Self {
            click_count: (bits & 0x001f) as u8,
            label_only: bits & 0x0020 != 0,
            data_only: bits & 0x0040 != 0,
            toggle_data_header: bits & 0x0080 != 0,
            selection_click: bits & 0x0100 != 0,
            extendable: bits & 0x0200 != 0,
            unused: (bits >> 10) as u8,
        }
    }

    fn bits(self) -> Result<u16> {
        if self.click_count > 0x1f || self.unused > 0x3f {
            return Err(Error::invalid(
                0,
                "SxSelect options exceed their bit widths",
            ));
        }
        Ok(u16::from(self.click_count)
            | (u16::from(self.label_only) << 5)
            | (u16::from(self.data_only) << 6)
            | (u16::from(self.toggle_data_header) << 7)
            | (u16::from(self.selection_click) << 8)
            | (u16::from(self.extendable) << 9)
            | (u16::from(self.unused) << 10))
    }
}

impl SdkRead for SxSelectRecord {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        Ok(Self {
            reserved1: reader.read_u16()?,
            pane: PaneType::read_from(reader)?,
            reserved2: reader.read_u8()?,
            axis: SxAxis::from_bits_retain(reader.read_u16()?),
            active_dimension: reader.read_u16()?,
            line_start: reader.read_u16()?,
            active_line: reader.read_u16()?,
            minimum_line: reader.read_u16()?,
            maximum_line: reader.read_u16()?,
            clicked_row: reader.read_u16()?,
            clicked_column: reader.read_u16()?,
            previous_clicked_row: reader.read_u16()?,
            previous_clicked_column: reader.read_u16()?,
            options: SxSelectOptions::from_bits(reader.read_u16()?),
        })
    }
}

impl SdkWrite for SxSelectRecord {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_u16(self.reserved1)?;
        self.pane.write_to(writer)?;
        writer.write_u8(self.reserved2)?;
        writer.write_u16(self.axis.bits())?;
        writer.write_u16(self.active_dimension)?;
        writer.write_u16(self.line_start)?;
        writer.write_u16(self.active_line)?;
        writer.write_u16(self.minimum_line)?;
        writer.write_u16(self.maximum_line)?;
        writer.write_u16(self.clicked_row)?;
        writer.write_u16(self.clicked_column)?;
        writer.write_u16(self.previous_clicked_row)?;
        writer.write_u16(self.previous_clicked_column)?;
        writer.write_u16(self.options.bits()?)
    }
}

impl MsoDrawingRecord {
    fn from_segments(segments: &[(u16, &[u8])], limits: Limits) -> Result<Self> {
        let total_len =
            segments
                .iter()
                .map(|(_, payload)| *payload)
                .try_fold(0usize, |total, payload| {
                    total.checked_add(payload.len()).ok_or_else(|| {
                        Error::Limit("MsoDrawing logical record length overflow".into())
                    })
                })?;
        if total_len > limits.max_allocation {
            return Err(Error::Limit(format!(
                "MsoDrawing logical record length exceeds {}",
                limits.max_allocation
            )));
        }
        let mut bytes = Vec::with_capacity(total_len);
        for (_, payload) in segments {
            bytes.extend_from_slice(payload);
        }
        let physical_segments = segments
            .iter()
            .map(|(record_type, payload)| {
                Ok(MsoDrawingSegment {
                    record_type: *record_type,
                    payload_length: u16::try_from(payload.len()).map_err(|_| {
                        Error::Limit("MsoDrawing physical segment exceeds u16".into())
                    })?,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(Self {
            data: match OfficeArtStream::from_bytes_with_limits(&bytes, limits) {
                Ok(value) => MsoDrawingData::Complete(value),
                Err(error) => match OfficeArtPartialStream::from_bytes_with_limits(
                    &bytes,
                    limits,
                    error.to_string(),
                ) {
                    Ok(partial) => MsoDrawingData::Partial(partial),
                    Err(_) => MsoDrawingData::Incomplete {
                        bytes,
                        reason: error.to_string(),
                    },
                },
            },
            physical_segments,
            host_records: Vec::new(),
            following_record_type: None,
        })
    }

    fn from_interleaved(
        segments: &[(u16, &[u8])],
        host_records: Vec<MsoDrawingHostRecord>,
        following_record_type: Option<u16>,
        limits: Limits,
    ) -> Result<Self> {
        let mut value = Self::from_segments(segments, limits)?;
        if host_records.iter().any(|host| {
            host.after_segment == 0 || host.after_segment > value.physical_segments.len()
        }) {
            return Err(Error::invalid(0, "invalid MsoDrawing host-record layout"));
        }
        value.host_records = host_records;
        value.following_record_type = following_record_type;
        Ok(value)
    }

    fn encode_physical(&self, expected_record_type: u16) -> Result<Vec<EncodedBiffRecord>> {
        let logical = match &self.data {
            MsoDrawingData::Complete(value) => value.to_bytes()?,
            MsoDrawingData::Partial(value) => value.to_bytes()?,
            MsoDrawingData::Incomplete { bytes, .. } => bytes.clone(),
        };
        let declared_len = self
            .physical_segments
            .iter()
            .map(|segment| usize::from(segment.payload_length))
            .sum::<usize>();
        if self.physical_segments.is_empty()
            || !(self.physical_segments[0].record_type == expected_record_type
                || (expected_record_type == MSO_DRAWING
                    && self.physical_segments[0].record_type == MSO_DRAWING_AC_COMPATIBILITY))
            || declared_len != logical.len()
        {
            return Err(Error::invalid(0, "MsoDrawing physical layout mismatch"));
        }
        let mut offset = 0usize;
        if expected_record_type == MSO_DRAWING_GROUP && !self.host_records.is_empty() {
            return Err(Error::invalid(
                0,
                "MsoDrawingGroup cannot contain BIFF host records",
            ));
        }
        let mut records =
            Vec::with_capacity(self.physical_segments.len() + self.host_records.len());
        let mut host_index = 0usize;
        for (segment_index, segment) in self.physical_segments.iter().enumerate() {
            if !matches!(
                segment.record_type,
                MSO_DRAWING_GROUP | MSO_DRAWING | MSO_DRAWING_AC_COMPATIBILITY | CONTINUE
            ) {
                return Err(Error::invalid(
                    0,
                    "invalid MsoDrawing physical segment type",
                ));
            }
            let end = offset + usize::from(segment.payload_length);
            records.push(EncodedBiffRecord {
                record_type: segment.record_type,
                payload: logical[offset..end].to_vec(),
            });
            offset = end;
            while self
                .host_records
                .get(host_index)
                .is_some_and(|host| host.after_segment == segment_index + 1)
            {
                let host = &self.host_records[host_index];
                records.extend(host.data.encode_physical()?);
                host_index += 1;
            }
        }
        if host_index != self.host_records.len() {
            return Err(Error::invalid(0, "MsoDrawing host-record order mismatch"));
        }
        Ok(records)
    }
}

impl DevModeW {
    fn from_bytes(bytes: &[u8]) -> Option<Self> {
        if bytes.len() < 76 {
            return None;
        }
        let read_u16_at = |offset: usize| u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
        let declared_public_size = read_u16_at(68);
        let declared_driver_extra_size = read_u16_at(70);
        let public_size = usize::from(declared_public_size);
        let driver_extra_size = usize::from(declared_driver_extra_size);
        if public_size < 76 || !public_size.is_multiple_of(4) || public_size > bytes.len() {
            return None;
        }
        let mut device_name = [0u16; 32];
        for (index, value) in device_name.iter_mut().enumerate() {
            *value = read_u16_at(index * 2);
        }
        let fields = DevModeFields::from_bits_retain(u32::from_le_bytes(
            bytes[72..76].try_into().expect("four bytes"),
        ));
        let public_fields = if matches!(public_size, 100 | 212) || public_size >= 220 {
            let mut cursor = 76usize;
            let mut take_word = || {
                let value = u16::from_le_bytes([bytes[cursor], bytes[cursor + 1]]);
                cursor += 2;
                value
            };
            let core = DevModeWCore100 {
                orientation: take_word(),
                paper_size: take_word(),
                paper_length: take_word(),
                paper_width: take_word(),
                scale: take_word(),
                copies: take_word(),
                default_source: take_word(),
                print_quality: take_word(),
                color: take_word(),
                duplex: take_word(),
                y_resolution: take_word(),
                tt_option: take_word(),
            };
            if public_size == 100 {
                DevModeWPublic::Core100(core)
            } else {
                let collate = take_word();
                let mut form_name = [0u16; 32];
                for value in &mut form_name {
                    *value = take_word();
                }
                let log_pixels = take_word();
                let mut take_dword = || {
                    let value = u32::from_le_bytes(
                        bytes[cursor..cursor + 4].try_into().expect("four bytes"),
                    );
                    cursor += 4;
                    value
                };
                let legacy = DevModeWLegacy212 {
                    core: core.clone(),
                    collate,
                    form_name,
                    log_pixels,
                    bits_per_pel: take_dword(),
                    pels_width: take_dword(),
                    pels_height: take_dword(),
                    display_flags_or_nup: take_dword(),
                    display_frequency: take_dword(),
                    icm_method: take_dword(),
                    icm_intent: take_dword(),
                    media_type: take_dword(),
                    dither_type: take_dword(),
                    reserved1: take_dword(),
                    reserved2: take_dword(),
                };
                if public_size == 212 {
                    DevModeWPublic::Legacy212(Box::new(legacy))
                } else {
                    let full = DevModeWFull {
                        orientation: core.orientation,
                        paper_size: core.paper_size,
                        paper_length: core.paper_length,
                        paper_width: core.paper_width,
                        scale: core.scale,
                        copies: core.copies,
                        default_source: core.default_source,
                        print_quality: core.print_quality,
                        color: core.color,
                        duplex: core.duplex,
                        y_resolution: core.y_resolution,
                        tt_option: core.tt_option,
                        collate,
                        form_name,
                        log_pixels,
                        bits_per_pel: legacy.bits_per_pel,
                        pels_width: legacy.pels_width,
                        pels_height: legacy.pels_height,
                        display_flags_or_nup: legacy.display_flags_or_nup,
                        display_frequency: legacy.display_frequency,
                        icm_method: legacy.icm_method,
                        icm_intent: legacy.icm_intent,
                        media_type: legacy.media_type,
                        dither_type: legacy.dither_type,
                        reserved1: legacy.reserved1,
                        reserved2: legacy.reserved2,
                        panning_width: take_dword(),
                        panning_height: take_dword(),
                        public_extension: bytes[220..public_size].to_vec(),
                    };
                    DevModeWPublic::Full(Box::new(full))
                }
            }
        } else {
            DevModeWPublic::Truncated(bytes[76..public_size].to_vec())
        };
        let declared_driver_end = public_size.checked_add(driver_extra_size)?;
        let driver_extra_complete = declared_driver_end <= bytes.len();
        let driver_end = declared_driver_end.min(bytes.len());
        Some(Self {
            device_name,
            specification_version: read_u16_at(64),
            driver_version: read_u16_at(66),
            declared_public_size,
            declared_driver_extra_size,
            fields,
            public_fields,
            driver_extra: bytes[public_size..driver_end].to_vec(),
            driver_extra_complete,
            trailing: if driver_extra_complete {
                bytes[driver_end..].to_vec()
            } else {
                Vec::new()
            },
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.driver_extra_complete
            && usize::from(self.declared_driver_extra_size) != self.driver_extra.len()
        {
            return Err(Error::invalid(0, "DEVMODEW dmDriverExtra mismatch"));
        }
        if !self.driver_extra_complete
            && self.driver_extra.len() >= usize::from(self.declared_driver_extra_size)
        {
            return Err(Error::invalid(
                0,
                "truncated DEVMODEW driver extra is not shorter than declared",
            ));
        }
        let mut bytes = Vec::new();
        for value in self.device_name {
            bytes.extend_from_slice(&value.to_le_bytes());
        }
        bytes.extend_from_slice(&self.specification_version.to_le_bytes());
        bytes.extend_from_slice(&self.driver_version.to_le_bytes());
        bytes.extend_from_slice(&self.declared_public_size.to_le_bytes());
        bytes.extend_from_slice(&self.declared_driver_extra_size.to_le_bytes());
        bytes.extend_from_slice(&self.fields.bits().to_le_bytes());
        match &self.public_fields {
            DevModeWPublic::Truncated(public) => bytes.extend_from_slice(public),
            DevModeWPublic::Core100(core) => core.write_to(&mut bytes),
            DevModeWPublic::Legacy212(legacy) => legacy.write_to(&mut bytes),
            DevModeWPublic::Full(value) => {
                for word in [
                    value.orientation,
                    value.paper_size,
                    value.paper_length,
                    value.paper_width,
                    value.scale,
                    value.copies,
                    value.default_source,
                    value.print_quality,
                    value.color,
                    value.duplex,
                    value.y_resolution,
                    value.tt_option,
                    value.collate,
                ] {
                    bytes.extend_from_slice(&word.to_le_bytes());
                }
                for word in value.form_name {
                    bytes.extend_from_slice(&word.to_le_bytes());
                }
                bytes.extend_from_slice(&value.log_pixels.to_le_bytes());
                for dword in [
                    value.bits_per_pel,
                    value.pels_width,
                    value.pels_height,
                    value.display_flags_or_nup,
                    value.display_frequency,
                    value.icm_method,
                    value.icm_intent,
                    value.media_type,
                    value.dither_type,
                    value.reserved1,
                    value.reserved2,
                    value.panning_width,
                    value.panning_height,
                ] {
                    bytes.extend_from_slice(&dword.to_le_bytes());
                }
                bytes.extend_from_slice(&value.public_extension);
            }
        }
        if bytes.len() != usize::from(self.declared_public_size) {
            return Err(Error::invalid(0, "DEVMODEW dmSize mismatch"));
        }
        bytes.extend_from_slice(&self.driver_extra);
        bytes.extend_from_slice(&self.trailing);
        Ok(bytes)
    }
}

impl DevModeWCore100 {
    fn write_to(&self, bytes: &mut Vec<u8>) {
        for word in [
            self.orientation,
            self.paper_size,
            self.paper_length,
            self.paper_width,
            self.scale,
            self.copies,
            self.default_source,
            self.print_quality,
            self.color,
            self.duplex,
            self.y_resolution,
            self.tt_option,
        ] {
            bytes.extend_from_slice(&word.to_le_bytes());
        }
    }
}

impl DevModeWLegacy212 {
    fn write_to(&self, bytes: &mut Vec<u8>) {
        self.core.write_to(bytes);
        bytes.extend_from_slice(&self.collate.to_le_bytes());
        for word in self.form_name {
            bytes.extend_from_slice(&word.to_le_bytes());
        }
        bytes.extend_from_slice(&self.log_pixels.to_le_bytes());
        for dword in [
            self.bits_per_pel,
            self.pels_width,
            self.pels_height,
            self.display_flags_or_nup,
            self.display_frequency,
            self.icm_method,
            self.icm_intent,
            self.media_type,
            self.dither_type,
            self.reserved1,
            self.reserved2,
        ] {
            bytes.extend_from_slice(&dword.to_le_bytes());
        }
    }
}

impl SstExtensionData {
    fn from_bytes(bytes: Vec<u8>) -> Self {
        let Some(raw_reserved) = bytes.get(..2) else {
            return Self::Unparsed(bytes);
        };
        let reserved = u16::from_le_bytes([raw_reserved[0], raw_reserved[1]]);
        if reserved == u16::MAX {
            return Self::ExtRst(ExtRst {
                reserved,
                body: ExtRstBody::OldStyle {
                    payload: bytes[2..].to_vec(),
                },
            });
        }
        if reserved != 1 {
            if bytes.len() == 12 {
                let declared_data_size = u16::from_le_bytes([bytes[2], bytes[3]]);
                if declared_data_size == 8 {
                    return Self::ExtRst(ExtRst {
                        reserved,
                        body: ExtRstBody::TruncatedPhoneticHeader {
                            declared_data_size,
                            font_index: u16::from_le_bytes([bytes[4], bytes[5]]),
                            formatting_flags: PhoneticFlags::from_bits_retain(u16::from_le_bytes(
                                [bytes[6], bytes[7]],
                            )),
                            declared_run_count: u16::from_le_bytes([bytes[8], bytes[9]]),
                            declared_character_count: u16::from_le_bytes([bytes[10], bytes[11]]),
                        },
                    });
                }
            }
            return Self::ExtRst(ExtRst {
                reserved,
                body: ExtRstBody::InvalidMarker {
                    payload: bytes[2..].to_vec(),
                },
            });
        }
        if bytes.len() < 14 {
            return Self::Unparsed(bytes);
        }
        let declared_data_size = u16::from_le_bytes([bytes[2], bytes[3]]);
        let declared_end = 4usize.saturating_add(usize::from(declared_data_size));
        if declared_end > bytes.len() || declared_end < 14 {
            return Self::Unparsed(bytes);
        }
        let font_index = u16::from_le_bytes([bytes[4], bytes[5]]);
        let formatting_flags =
            PhoneticFlags::from_bits_retain(u16::from_le_bytes([bytes[6], bytes[7]]));
        let declared_run_count = u16::from_le_bytes([bytes[8], bytes[9]]);
        let declared_character_count = u16::from_le_bytes([bytes[10], bytes[11]]);
        let lpwide_character_count = u16::from_le_bytes([bytes[12], bytes[13]]);
        let text_byte_count = usize::from(declared_character_count).saturating_mul(2);
        let text_end = 14usize.saturating_add(text_byte_count);
        if text_end > declared_end {
            return Self::Unparsed(bytes);
        }
        let phonetic_text = bytes[14..text_end]
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        let run_region = &bytes[text_end..declared_end];
        let run_byte_count = usize::from(declared_run_count).saturating_mul(6);
        if run_byte_count > run_region.len() {
            return Self::Unparsed(bytes);
        }
        let runs = run_region[..run_byte_count]
            .chunks_exact(6)
            .map(|run| PhoneticRun {
                phonetic_text_first_character: u16::from_le_bytes([run[0], run[1]]),
                source_text_first_character: u16::from_le_bytes([run[2], run[3]]),
                source_text_character_count: u16::from_le_bytes([run[4], run[5]]),
            })
            .collect::<Vec<_>>();
        let remaining = &run_region[run_byte_count..];
        let (extra_data_word, inner_trailing) = match remaining {
            [low, high] => (Some(u16::from_le_bytes([*low, *high])), Vec::new()),
            _ => (None, remaining.to_vec()),
        };
        Self::ExtRst(ExtRst {
            reserved,
            body: ExtRstBody::Phonetic {
                declared_data_size,
                font_index,
                formatting_flags,
                declared_run_count,
                declared_character_count,
                lpwide_character_count,
                phonetic_text,
                runs,
                extra_data_word,
                inner_trailing,
                outer_trailing: bytes[declared_end..].to_vec(),
            },
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        match self {
            Self::None => Ok(Vec::new()),
            Self::Unparsed(bytes) => Ok(bytes.clone()),
            Self::ExtRst(value) => value.to_bytes(),
        }
    }

    pub fn unparsed_byte_count(&self) -> usize {
        match self {
            Self::None => 0,
            Self::Unparsed(bytes) => bytes.len(),
            Self::ExtRst(ExtRst {
                body:
                    ExtRstBody::Phonetic {
                        inner_trailing,
                        outer_trailing,
                        ..
                    },
                ..
            }) => inner_trailing.len() + outer_trailing.len(),
            Self::ExtRst(ExtRst {
                body: ExtRstBody::TruncatedPhoneticHeader { .. },
                ..
            }) => 0,
            Self::ExtRst(ExtRst {
                body: ExtRstBody::OldStyle { payload } | ExtRstBody::InvalidMarker { payload },
                ..
            }) => payload.len(),
        }
    }
}

impl ExtRst {
    fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&self.reserved.to_le_bytes());
        match &self.body {
            ExtRstBody::OldStyle { payload } | ExtRstBody::InvalidMarker { payload } => {
                bytes.extend_from_slice(payload);
            }
            ExtRstBody::TruncatedPhoneticHeader {
                declared_data_size,
                font_index,
                formatting_flags,
                declared_run_count,
                declared_character_count,
            } => {
                bytes.extend_from_slice(&declared_data_size.to_le_bytes());
                bytes.extend_from_slice(&font_index.to_le_bytes());
                bytes.extend_from_slice(&formatting_flags.bits().to_le_bytes());
                bytes.extend_from_slice(&declared_run_count.to_le_bytes());
                bytes.extend_from_slice(&declared_character_count.to_le_bytes());
            }
            ExtRstBody::Phonetic {
                declared_data_size,
                font_index,
                formatting_flags,
                declared_run_count,
                declared_character_count,
                lpwide_character_count,
                phonetic_text,
                runs,
                extra_data_word,
                inner_trailing,
                outer_trailing,
            } => {
                if usize::from(*declared_character_count) != phonetic_text.len() {
                    return Err(Error::invalid(
                        0,
                        "ExtRst phonetic character count does not match text",
                    ));
                }
                bytes.extend_from_slice(&declared_data_size.to_le_bytes());
                bytes.extend_from_slice(&font_index.to_le_bytes());
                bytes.extend_from_slice(&formatting_flags.bits().to_le_bytes());
                bytes.extend_from_slice(&declared_run_count.to_le_bytes());
                bytes.extend_from_slice(&declared_character_count.to_le_bytes());
                bytes.extend_from_slice(&lpwide_character_count.to_le_bytes());
                for character in phonetic_text {
                    bytes.extend_from_slice(&character.to_le_bytes());
                }
                for run in runs {
                    bytes.extend_from_slice(&run.phonetic_text_first_character.to_le_bytes());
                    bytes.extend_from_slice(&run.source_text_first_character.to_le_bytes());
                    bytes.extend_from_slice(&run.source_text_character_count.to_le_bytes());
                }
                if let Some(value) = extra_data_word {
                    bytes.extend_from_slice(&value.to_le_bytes());
                }
                bytes.extend_from_slice(inner_trailing);
                if bytes.len() != 4 + usize::from(*declared_data_size) {
                    return Err(Error::invalid(
                        0,
                        "ExtRst declared data size does not match fields",
                    ));
                }
                bytes.extend_from_slice(outer_trailing);
            }
        }
        Ok(bytes)
    }
}

#[derive(Clone)]
struct SstSequenceCursor<'a> {
    segments: Vec<&'a [u8]>,
    segment_index: usize,
    offset: usize,
    continuation_encodings: Vec<Option<u8>>,
}

impl<'a> SstSequenceCursor<'a> {
    fn new(segments: Vec<&'a [u8]>) -> Result<Self> {
        if segments.is_empty() {
            return Err(Error::invalid(0, "SST has no physical record"));
        }
        let segment_count = segments.len();
        Ok(Self {
            segments,
            segment_index: 0,
            offset: 0,
            continuation_encodings: vec![None; segment_count],
        })
    }

    fn remaining_in_segment(&self) -> usize {
        self.segments[self.segment_index].len() - self.offset
    }

    fn advance_plain(&mut self) -> Result<()> {
        self.segment_index += 1;
        self.offset = 0;
        if self.segment_index >= self.segments.len() {
            return Err(Error::invalid(0, "SST data exceeds its Continue records"));
        }
        Ok(())
    }

    fn advance_characters(&mut self) -> Result<u8> {
        self.advance_plain()?;
        let encoding = *self.segments[self.segment_index]
            .first()
            .ok_or_else(|| Error::invalid(0, "SST character Continue record is empty"))?;
        if encoding > 1 {
            return Err(Error::invalid(
                0,
                format!("SST character Continue encoding {encoding} is invalid"),
            ));
        }
        self.offset = 1;
        self.continuation_encodings[self.segment_index] = Some(encoding);
        Ok(encoding)
    }

    fn ensure_contiguous(&mut self, byte_count: usize, context: &str) -> Result<()> {
        if self.remaining_in_segment() == 0 {
            self.advance_plain()?;
        }
        if self.remaining_in_segment() < byte_count {
            return Err(Error::invalid(
                0,
                format!("{context} crosses an SST physical record boundary"),
            ));
        }
        Ok(())
    }

    fn read_contiguous(&mut self, byte_count: usize, context: &str) -> Result<&'a [u8]> {
        self.ensure_contiguous(byte_count, context)?;
        let start = self.offset;
        self.offset += byte_count;
        Ok(&self.segments[self.segment_index][start..self.offset])
    }

    fn read_u16(&mut self, context: &str) -> Result<u16> {
        let bytes = self.read_contiguous(2, context)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn read_u32(&mut self, context: &str) -> Result<u32> {
        let bytes = self.read_contiguous(4, context)?;
        Ok(u32::from_le_bytes(bytes.try_into().expect("four bytes")))
    }

    fn read_bytes(&mut self, mut byte_count: usize, context: &str) -> Result<Vec<u8>> {
        let mut bytes = Vec::with_capacity(byte_count);
        while byte_count != 0 {
            if self.remaining_in_segment() == 0 {
                self.advance_plain().map_err(|_| {
                    Error::invalid(0, format!("{context} exceeds SST Continue records"))
                })?;
            }
            let count = byte_count.min(self.remaining_in_segment());
            let start = self.offset;
            self.offset += count;
            bytes.extend_from_slice(&self.segments[self.segment_index][start..self.offset]);
            byte_count -= count;
        }
        Ok(bytes)
    }

    fn read_characters(
        &mut self,
        character_count: usize,
        initial_encoding: u8,
    ) -> Result<Vec<SstCharacterChunk>> {
        let mut remaining = character_count;
        let mut encoding = initial_encoding;
        let mut chunks = Vec::new();
        while remaining != 0 {
            if self.remaining_in_segment() == 0 {
                encoding = self.advance_characters()?;
            }
            let width = if encoding == 0 { 1 } else { 2 };
            let available = self.remaining_in_segment() / width;
            if available == 0 {
                return Err(Error::invalid(
                    0,
                    "SST UTF-16 characters split at a non-character boundary",
                ));
            }
            let count = remaining.min(available);
            let byte_count = count * width;
            let start = self.offset;
            self.offset += byte_count;
            let data = &self.segments[self.segment_index][start..self.offset];
            let characters = if encoding == 0 {
                XlStringCharacters::Compressed(data.to_vec())
            } else {
                XlStringCharacters::Unicode(
                    data.chunks_exact(2)
                        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                        .collect(),
                )
            };
            chunks.push(SstCharacterChunk {
                flags: encoding,
                characters,
            });
            remaining -= count;
            if remaining != 0 && self.remaining_in_segment() != 0 {
                return Err(Error::invalid(
                    0,
                    "SST character data ended before its physical segment",
                ));
            }
        }
        Ok(chunks)
    }

    fn read_string(&mut self, limits: Limits) -> Result<SstString> {
        let fixed = self.read_contiguous(3, "SST string header")?;
        let declared_character_count = u16::from_le_bytes([fixed[0], fixed[1]]);
        let flags = SstStringFlags::from_bits_retain(fixed[2]);
        let declared_format_run_count = if flags.contains(SstStringFlags::RICH_TEXT) {
            Some(self.read_u16("SST rich-text run count")?)
        } else {
            None
        };
        let declared_extension_length = if flags.contains(SstStringFlags::EXTENDED) {
            Some(self.read_u32("SST extension length")?)
        } else {
            None
        };
        let character_chunks = self.read_characters(
            usize::from(declared_character_count),
            flags.bits() & SstStringFlags::HIGH_BYTE.bits(),
        )?;
        let run_count = usize::from(declared_format_run_count.unwrap_or(0));
        if run_count > limits.max_entries {
            return Err(Error::Limit(format!(
                "SST format run count exceeds {}",
                limits.max_entries
            )));
        }
        let mut format_runs = Vec::with_capacity(run_count);
        for _ in 0..run_count {
            format_runs.push(FormatRun {
                character_index: self.read_u16("SST format-run character index")?,
                font_index: self.read_u16("SST format-run font index")?,
            });
        }
        let extension_length = usize::try_from(declared_extension_length.unwrap_or(0))
            .map_err(|_| Error::Limit("SST extension length exceeds usize".into()))?;
        if extension_length > limits.max_allocation {
            return Err(Error::Limit(format!(
                "SST extension length exceeds {}",
                limits.max_allocation
            )));
        }
        let extension_bytes = self.read_bytes(extension_length, "SST extension data")?;
        let extension = if declared_extension_length.is_some() {
            SstExtensionData::from_bytes(extension_bytes)
        } else {
            SstExtensionData::None
        };
        Ok(SstString {
            declared_character_count,
            flags,
            declared_format_run_count,
            declared_extension_length,
            character_chunks,
            format_runs,
            extension,
        })
    }

    fn read_trailing(mut self) -> Result<(Vec<u8>, Vec<SstSegmentLayout>)> {
        let mut trailing = Vec::new();
        loop {
            trailing.extend_from_slice(&self.segments[self.segment_index][self.offset..]);
            if self.segment_index + 1 == self.segments.len() {
                break;
            }
            self.advance_plain()?;
        }
        let layouts = self
            .segments
            .iter()
            .zip(self.continuation_encodings)
            .map(|(payload, continuation_encoding)| {
                let logical_len = payload
                    .len()
                    .checked_sub(usize::from(continuation_encoding.is_some()))
                    .expect("continuation encoding is part of payload");
                Ok(SstSegmentLayout {
                    logical_byte_count: u16::try_from(logical_len)
                        .map_err(|_| Error::Limit("SST segment length exceeds u16".into()))?,
                    continuation_encoding,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        Ok((trailing, layouts))
    }
}

impl SstRecord {
    fn from_sequence(first: &[u8], continues: &[&[u8]], limits: Limits) -> Result<Self> {
        let mut segments = Vec::with_capacity(continues.len() + 1);
        segments.push(first);
        segments.extend_from_slice(continues);
        let mut cursor = SstSequenceCursor::new(segments)?;
        let header = cursor.read_contiguous(8, "SST header")?;
        let total_string_count = u32::from_le_bytes(header[0..4].try_into().expect("four bytes"));
        let unique_string_count = u32::from_le_bytes(header[4..8].try_into().expect("four bytes"));
        let unique_count = usize::try_from(unique_string_count).unwrap_or(usize::MAX);
        let mut strings = Vec::with_capacity(unique_count.min(limits.max_entries));
        let mut completion = SstCompletion::Complete;
        for index in 0..unique_count {
            if index >= limits.max_entries {
                return Err(Error::Limit(format!(
                    "parsed SST string count exceeds {}",
                    limits.max_entries
                )));
            }
            let checkpoint = cursor.clone();
            match cursor.read_string(limits) {
                Ok(string) => strings.push(string),
                Err(error) => {
                    cursor = checkpoint;
                    completion = SstCompletion::Truncated {
                        first_unparsed_string: u32::try_from(index).unwrap_or(u32::MAX),
                        reason: error.to_string(),
                    };
                    break;
                }
            }
        }
        let (trailing, physical_segments) = cursor.read_trailing()?;
        Ok(Self {
            total_string_count,
            unique_string_count,
            strings,
            completion,
            trailing,
            physical_segments,
        })
    }

    fn logical_bytes(&self) -> Result<Vec<u8>> {
        match self.completion {
            SstCompletion::Complete
                if usize::try_from(self.unique_string_count).ok() != Some(self.strings.len()) =>
            {
                return Err(Error::invalid(
                    0,
                    "complete SST unique string count does not match strings",
                ));
            }
            SstCompletion::Truncated {
                first_unparsed_string,
                ..
            } if usize::try_from(first_unparsed_string).ok() != Some(self.strings.len()) => {
                return Err(Error::invalid(
                    0,
                    "truncated SST index does not match parsed strings",
                ));
            }
            _ => {}
        }
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&self.total_string_count.to_le_bytes());
        bytes.extend_from_slice(&self.unique_string_count.to_le_bytes());
        for string in &self.strings {
            bytes.extend_from_slice(&string.declared_character_count.to_le_bytes());
            bytes.push(string.flags.bits());
            let extension_bytes = string.extension.to_bytes()?;
            match (
                string.flags.contains(SstStringFlags::RICH_TEXT),
                string.declared_format_run_count,
            ) {
                (true, Some(count)) if usize::from(count) == string.format_runs.len() => {
                    bytes.extend_from_slice(&count.to_le_bytes());
                }
                (false, None) => {}
                _ => return Err(Error::invalid(0, "SST rich-text count/flags mismatch")),
            }
            match (
                string.flags.contains(SstStringFlags::EXTENDED),
                string.declared_extension_length,
            ) {
                (true, Some(length))
                    if usize::try_from(length).ok() == Some(extension_bytes.len()) =>
                {
                    bytes.extend_from_slice(&length.to_le_bytes());
                }
                (false, None) if matches!(&string.extension, SstExtensionData::None) => {}
                _ => return Err(Error::invalid(0, "SST extension length/flags mismatch")),
            }
            let mut actual_character_count = 0usize;
            for chunk in &string.character_chunks {
                match &chunk.characters {
                    XlStringCharacters::Compressed(values) => {
                        if chunk.flags != 0 {
                            return Err(Error::invalid(0, "compressed SST chunk has Unicode flag"));
                        }
                        actual_character_count += values.len();
                        bytes.extend_from_slice(values);
                    }
                    XlStringCharacters::Unicode(values) => {
                        if chunk.flags != 1 {
                            return Err(Error::invalid(0, "Unicode SST chunk lacks Unicode flag"));
                        }
                        actual_character_count += values.len();
                        for value in values {
                            bytes.extend_from_slice(&value.to_le_bytes());
                        }
                    }
                }
            }
            if actual_character_count != usize::from(string.declared_character_count) {
                return Err(Error::invalid(
                    0,
                    "SST character count does not match chunks",
                ));
            }
            for run in &string.format_runs {
                bytes.extend_from_slice(&run.character_index.to_le_bytes());
                bytes.extend_from_slice(&run.font_index.to_le_bytes());
            }
            bytes.extend_from_slice(&extension_bytes);
        }
        bytes.extend_from_slice(&self.trailing);
        Ok(bytes)
    }

    fn encode_physical(&self) -> Result<Vec<EncodedBiffRecord>> {
        let logical = self.logical_bytes()?;
        let declared_logical_len =
            self.physical_segments
                .iter()
                .try_fold(0usize, |total, layout| {
                    total
                        .checked_add(usize::from(layout.logical_byte_count))
                        .ok_or_else(|| Error::Limit("SST physical layout length overflow".into()))
                })?;
        if declared_logical_len != logical.len() || self.physical_segments.is_empty() {
            return Err(Error::invalid(
                0,
                "SST physical layout does not match logical bytes",
            ));
        }
        let mut offset = 0usize;
        let mut encoded = Vec::with_capacity(self.physical_segments.len());
        for (index, layout) in self.physical_segments.iter().enumerate() {
            if index == 0 && layout.continuation_encoding.is_some() {
                return Err(Error::invalid(
                    0,
                    "first SST segment has a continuation encoding",
                ));
            }
            let end = offset + usize::from(layout.logical_byte_count);
            let mut payload = Vec::with_capacity(
                usize::from(layout.logical_byte_count)
                    + usize::from(layout.continuation_encoding.is_some()),
            );
            if let Some(encoding) = layout.continuation_encoding {
                if encoding > 1 {
                    return Err(Error::invalid(0, "invalid SST continuation encoding"));
                }
                payload.push(encoding);
            }
            payload.extend_from_slice(&logical[offset..end]);
            encoded.push(EncodedBiffRecord {
                record_type: if index == 0 { SST } else { CONTINUE },
                payload,
            });
            offset = end;
        }
        Ok(encoded)
    }
}

impl BiffStream {
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        if bytes.len() as u64 > limits.max_stream_size {
            return Err(Error::Limit(format!(
                "BIFF stream length {} exceeds {}",
                bytes.len(),
                limits.max_stream_size
            )));
        }
        let mut records = Vec::new();
        let mut cursor = 0usize;
        let mut encrypted = false;
        let mut biff8 = false;
        let mut sx_li_dimensions: Option<(u16, u16, u8)> = None;
        while cursor < bytes.len() {
            if bytes[cursor..].iter().all(|value| *value == 0) {
                if biff8 {
                    stitch_continued_records(&mut records, limits)?;
                }
                return Ok(Self {
                    records,
                    trailing_padding: bytes[cursor..].to_vec(),
                });
            }
            let offset = cursor;
            let record_type = take_u16(bytes, &mut cursor, "truncated BIFF record type")?;
            let size = usize::from(take_u16(bytes, &mut cursor, "truncated BIFF record size")?);
            if size > MAX_BIFF_RECORD_DATA {
                return Err(Error::invalid(
                    offset as u64 + 2,
                    "BIFF record data exceeds 8224 bytes",
                ));
            }
            let payload = take_bytes(bytes, &mut cursor, size, "truncated BIFF record data")?;
            if record_type == BOF {
                sx_li_dimensions = None;
            }
            let data = if record_type == SX_LI && biff8 && !encrypted {
                let (row_dimensions, column_dimensions, next_axis) =
                    sx_li_dimensions.as_mut().ok_or_else(|| {
                        Error::invalid(
                            offset as u64,
                            "SXLI has no preceding SxView dimension context",
                        )
                    })?;
                let axis_dimension_count = match *next_axis {
                    0 => *row_dimensions,
                    1 => *column_dimensions,
                    _ => {
                        return Err(Error::invalid(
                            offset as u64,
                            "more than two SXLI records follow one SxView",
                        ));
                    }
                };
                *next_axis += 1;
                BiffRecordData::SxLi(parse_sx_li(payload, offset, axis_dimension_count)?)
            } else {
                decode_record(record_type, payload, encrypted, biff8, offset)?
            };
            if records.is_empty() {
                biff8 = matches!(
                    &data,
                    BiffRecordData::Bof(BofRecord {
                        version: 0x0600,
                        ..
                    })
                );
            }
            if record_type == FILE_PASS {
                encrypted = true;
            }
            if let BiffRecordData::SxView(view) = &data {
                sx_li_dimensions = Some((view.row_field_count, view.column_field_count, 0));
            }
            records.push(BiffRecord {
                offset: u32::try_from(offset)
                    .map_err(|_| Error::Limit("BIFF record offset exceeds u32".into()))?,
                data,
            });
            if records.len() > limits.max_entries {
                return Err(Error::Limit(format!(
                    "BIFF record count exceeds {}",
                    limits.max_entries
                )));
            }
        }
        if biff8 {
            stitch_continued_records(&mut records, limits)?;
        }
        Ok(Self {
            records,
            trailing_padding: Vec::new(),
        })
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        for record in &self.records {
            if usize::try_from(record.offset).ok() != Some(bytes.len()) {
                return Err(Error::invalid(
                    bytes.len() as u64,
                    "BIFF record offset changed",
                ));
            }
            for encoded in record.data.encode_physical()? {
                if encoded.payload.len() > MAX_BIFF_RECORD_DATA {
                    return Err(Error::invalid(
                        bytes.len() as u64,
                        "BIFF record data exceeds 8224 bytes",
                    ));
                }
                bytes.extend_from_slice(&encoded.record_type.to_le_bytes());
                bytes.extend_from_slice(&(encoded.payload.len() as u16).to_le_bytes());
                bytes.extend_from_slice(&encoded.payload);
            }
        }
        bytes.extend_from_slice(&self.trailing_padding);
        Ok(bytes)
    }

    pub fn is_biff8(&self) -> bool {
        matches!(
            self.records.first().map(|record| &record.data),
            Some(BiffRecordData::Bof(BofRecord {
                version: 0x0600,
                ..
            }))
        )
    }

    pub fn unknown_record_types(&self) -> BTreeSet<u16> {
        self.records
            .iter()
            .filter_map(|record| match record.data {
                BiffRecordData::Unknown { record_type, .. } => Some(record_type),
                _ => None,
            })
            .collect()
    }
}

impl BiffRecordData {
    fn encode_physical(&self) -> Result<Vec<EncodedBiffRecord>> {
        if let Self::StringValue(value) = self {
            return value.encode_physical();
        }
        if let Self::Sst(value) = self {
            return value.encode_physical();
        }
        if let Self::Name(value) = self {
            return value.encode_physical();
        }
        if let Self::Pls(value) = self {
            return value.encode_physical();
        }
        if let Self::BkHim(value) = self {
            return value.encode_physical(BK_HIM);
        }
        if let Self::ImData(value) = self {
            return value.encode_physical(IM_DATA);
        }
        if let Self::SortData(value) = self {
            return value.encode_physical();
        }
        if let Self::Txo(value) = self {
            return value.encode_physical();
        }
        if let Self::MsoDrawingGroup(value) = self {
            return value.encode_physical(MSO_DRAWING_GROUP);
        }
        if let Self::MsoDrawing(value) = self {
            return value.encode_physical(MSO_DRAWING);
        }
        let (record_type, payload) = self.encode()?;
        Ok(vec![EncodedBiffRecord {
            record_type,
            payload,
        }])
    }

    fn encode(&self) -> Result<(u16, Vec<u8>)> {
        Ok(match self {
            Self::Bof(value) => {
                let mut bytes = Vec::with_capacity(16);
                bytes.extend_from_slice(&value.version.to_le_bytes());
                bytes.extend_from_slice(&value.document_type.to_le_bytes());
                bytes.extend_from_slice(&value.build_identifier.to_le_bytes());
                bytes.extend_from_slice(&value.build_year.to_le_bytes());
                bytes.extend_from_slice(&value.history_flags.to_le_bytes());
                bytes.extend_from_slice(&value.lowest_version.to_le_bytes());
                (BOF, bytes)
            }
            Self::LegacyBof { payload } => (BOF, payload.clone()),
            Self::Eof => (EOF, Vec::new()),
            Self::Formula(value) => (FORMULA, encode_sdk(value)?),
            Self::Formula4Compatibility(value) => (FORMULA4, encode_sdk(value)?),
            Self::SharedFormula(value) => (SHARED_FORMULA, encode_sdk(value)?),
            Self::SupBook(value) => (SUP_BOOK, encode_sdk(value)?),
            Self::ConditionalFormatting(value) => (CF, encode_sdk(value)?),
            Self::ConditionalFormattingGroup(value) => (COND_FMT, encode_sdk(value)?),
            Self::ExternSheet(value) => (EXTERN_SHEET, encode_sdk(value)?),
            Self::ExternName(value) => (EXTERN_NAME, encode_sdk(value)?),
            Self::Hyperlink(value) => (HLINK, encode_sdk(value)?),
            Self::DataValidation(value) => (DV, encode_sdk(value)?),
            Self::Name(value) => {
                let first = value
                    .encode_physical()?
                    .into_iter()
                    .next()
                    .ok_or_else(|| Error::invalid(0, "Name has no physical segments"))?;
                (NAME, first.payload)
            }
            Self::Pls(value) => {
                let first = value
                    .encode_physical()?
                    .into_iter()
                    .next()
                    .ok_or_else(|| Error::invalid(0, "Pls has no physical segments"))?;
                (PLS, first.payload)
            }
            Self::MsoDrawingGroup(value) => {
                let first = value
                    .encode_physical(MSO_DRAWING_GROUP)?
                    .into_iter()
                    .next()
                    .ok_or_else(|| Error::invalid(0, "MsoDrawingGroup has no physical segments"))?;
                (MSO_DRAWING_GROUP, first.payload)
            }
            Self::MsoDrawing(value) => {
                let first = value
                    .encode_physical(MSO_DRAWING)?
                    .into_iter()
                    .next()
                    .ok_or_else(|| Error::invalid(0, "MsoDrawing has no physical segments"))?;
                (MSO_DRAWING, first.payload)
            }
            Self::Obj(value) => (OBJ, value.to_bytes()?),
            Self::ObjCompatibility { record_type, value } => (*record_type, value.to_bytes()?),
            Self::Txo(value) => {
                let first = value
                    .encode_physical()?
                    .into_iter()
                    .next()
                    .ok_or_else(|| Error::invalid(0, "TxO has no physical segments"))?;
                (TXO, first.payload)
            }
            Self::Header(value) => (HEADER, encode_sdk(value)?),
            Self::Footer(value) => (FOOTER, encode_sdk(value)?),
            Self::VerticalPageBreaks(value) => (VERTICAL_PAGE_BREAKS, encode_sdk(value)?),
            Self::HorizontalPageBreaks(value) => (HORIZONTAL_PAGE_BREAKS, encode_sdk(value)?),
            Self::DCon(value) => (DCON, encode_sdk(value)?),
            Self::DConRef(value) => (DCON_REF, encode_sdk(value)?),
            Self::DConn(value) => (DCONN, encode_sdk(value)?),
            Self::TextQuery(value) => (TXT_QRY, encode_sdk(value)?),
            Self::QsiSxTag(value) => (QSI_SX_TAG, encode_sdk(value)?),
            Self::SxViewEx9(value) => (SX_VIEW_EX9, encode_sdk(value)?),
            Self::DbQueryExt(value) => (DB_QUERY_EXT, encode_sdk(value)?),
            Self::HyperlinkTooltip(value) => (HLINK_TOOLTIP, encode_sdk(value)?),
            Self::ContinueFrt12(value) => (CONTINUE_FRT12, encode_sdk(value)?),
            Self::SxAddl(value) => (SX_ADDL, encode_sdk(value)?),
            Self::EntExU2(value) => (ENT_EX_U2, encode_sdk(value)?),
            Self::BkHim(value) => {
                let first = value
                    .encode_physical(BK_HIM)?
                    .into_iter()
                    .next()
                    .ok_or_else(|| Error::invalid(0, "BkHim has no physical segments"))?;
                (BK_HIM, first.payload)
            }
            Self::ImData(value) => {
                let first = value
                    .encode_physical(IM_DATA)?
                    .into_iter()
                    .next()
                    .ok_or_else(|| Error::invalid(0, "ImData has no physical segments"))?;
                (IM_DATA, first.payload)
            }
            Self::RealTimeData(value) => (REAL_TIME_DATA, encode_sdk(value)?),
            Self::Sort(value) => (SORT, encode_sdk(value)?),
            Self::LhRecord(value) => (LH_RECORD, encode_sdk(value)?),
            Self::SortData(value) => (SORT_DATA, encode_sdk(value)?),
            Self::AutoFilter(value) => (AUTO_FILTER, encode_sdk(value)?),
            Self::SxFormat(value) => (SX_FORMAT, encode_sdk(value)?),
            Self::WOpt(value) => (W_OPT, encode_sdk(value)?),
            Self::Table(value) => (TABLE, encode_sdk(value)?),
            Self::ExternCount(value) => (EXTERN_COUNT, encode_sdk(value)?),
            Self::Qsi(value) => (QSI, encode_sdk(value)?),
            Self::ParamQry(value) => (PARAM_QRY, encode_sdk(value)?),
            Self::SxSelect(value) => (SX_SELECT, encode_sdk(value)?),
            Self::FileSharing(value) => (FILE_SHARING, encode_sdk(value)?),
            Self::OleObjectSize(value) => (OLE_OBJECT_SIZE, encode_sdk(value)?),
            Self::MsoDrawingSelection(value) => (MSO_DRAWING_SELECTION, encode_sdk(value)?),
            Self::ScenMan(value) => (SCEN_MAN, encode_sdk(value)?),
            Self::SxView(value) => (SX_VIEW, encode_sdk(value)?),
            Self::CodePage { code_page } => (CODE_PAGE, code_page.to_le_bytes().to_vec()),
            Self::BoundSheet8(value) => (BOUND_SHEET8, value.to_bytes()?),
            Self::BoundSheet8Compatibility { record_type, value } => {
                (*record_type, value.to_bytes()?)
            }
            Self::Dimensions(value) => (DIMENSIONS, encode_sdk(value)?),
            Self::Blank(value) => (BLANK, encode_sdk(value)?),
            Self::Number(value) => (NUMBER, encode_sdk(value)?),
            Self::BoolErr(value) => (BOOL_ERR, encode_sdk(value)?),
            Self::Label(value) => (LABEL, encode_sdk(value)?),
            Self::Sxvi(value) => (SXVI, encode_sdk(value)?),
            Self::SxIvd(value) => (SX_IVD, encode_sdk(value)?),
            Self::SxLi(value) => (SX_LI, encode_sdk(value)?),
            Self::SxPi(value) => (SX_PI, encode_sdk(value)?),
            Self::SxDi(value) => (SX_DI, encode_sdk(value)?),
            Self::SxString(value) => (SX_STRING, encode_sdk(value)?),
            Self::RrTabId(value) => (RR_TAB_ID, encode_sdk(value)?),
            Self::SxRule(value) => (SX_RULE, encode_sdk(value)?),
            Self::SxEx(value) => (SX_EX, encode_sdk(value)?),
            Self::SxFilt(value) => (SX_FILT, encode_sdk(value)?),
            Self::SxDxf(value) => (SX_DXF, encode_sdk(value)?),
            Self::SxItm(value) => (SX_ITM, encode_sdk(value)?),
            Self::SxStreamId(value) => (SX_STREAM_ID, encode_sdk(value)?),
            Self::SxVs(value) => (SX_VS, encode_sdk(value)?),
            Self::RecalcId(value) => (RECALC_ID, encode_sdk(value)?),
            Self::SxvdEx(value) => (SXVD_EX, encode_sdk(value)?),
            Self::Sxvd(value) => (SXVD, encode_sdk(value)?),
            Self::CodeName(value) => (CODE_NAME, encode_sdk(value)?),
            Self::Array(value) => (ARRAY, encode_sdk(value)?),
            Self::UserSViewBegin(value) => (USER_SVIEW_BEGIN, encode_sdk(value)?),
            Self::UserSViewEnd(value) => (USER_SVIEW_END, encode_sdk(value)?),
            Self::UserBView(value) => (USER_BVIEW, encode_sdk(value)?),
            Self::SheetExt(value) => (SHEET_EXT, encode_sdk(value)?),
            Self::ChartDataLabelExtContents(value) => {
                (CHART_DATA_LABEL_EXT_CONTENTS, encode_sdk(value)?)
            }
            Self::CellWatch(value) => (CELL_WATCH, encode_sdk(value)?),
            Self::FeatureHeader11(value) => (FEAT_HDR11, encode_sdk(value)?),
            Self::Feature11(value) => (FEATURE11, encode_sdk(value)?),
            Self::List12(value) => (LIST12, encode_sdk(value)?),
            Self::DropDownObjIds(value) => (DROP_DOWN_OBJ_IDS, encode_sdk(value)?),
            Self::DataValidationHeader(value) => (DATA_VALIDATION_HEADER, encode_sdk(value)?),
            Self::RichTextStream(value) => (RICH_TEXT_STREAM, encode_sdk(value)?),
            Self::GuidTypeLib(value) => (GUID_TYPE_LIB, encode_sdk(value)?),
            Self::NameComment(value) => (NAME_COMMENT, encode_sdk(value)?),
            Self::LabelSst(value) => (LABEL_SST, encode_sdk(value)?),
            Self::Rk(value) => (RK, encode_sdk(value)?),
            Self::Row(value) => (ROW, encode_sdk(value)?),
            Self::Window1(value) => (WINDOW1, encode_sdk(value)?),
            Self::Pane(value) => (PANE, encode_sdk(value)?),
            Self::ColInfo(value) => (COL_INFO, encode_sdk(value)?),
            Self::Guts(value) => (GUTS, encode_sdk(value)?),
            Self::Country(value) => (COUNTRY, encode_sdk(value)?),
            Self::Palette(value) => (PALETTE, encode_sdk(value)?),
            Self::Scl(value) => (SCL, encode_sdk(value)?),
            Self::PrintSetup(value) => (PRINT_SETUP, encode_sdk(value)?),
            Self::MulRk(value) => (MUL_RK, encode_sdk(value)?),
            Self::MulBlank(value) => (MUL_BLANK, encode_sdk(value)?),
            Self::Xf(value) => (XF, encode_sdk(value)?),
            Self::XfCompatibility { record_type, value } => (*record_type, encode_sdk(value)?),
            Self::Crn(value) => (CRN, encode_sdk(value)?),
            Self::XfExt(value) => (XF_EXT, encode_sdk(value)?),
            Self::XfCrc(value) => (XF_CRC, encode_sdk(value)?),
            Self::TableStyles(value) => (TABLE_STYLES, encode_sdk(value)?),
            Self::StyleExt(value) => (STYLE_EXT, encode_sdk(value)?),
            Self::Dxf(value) => (DXF, encode_sdk(value)?),
            Self::ConditionalFormattingExtension(value) => (CF_EX, encode_sdk(value)?),
            Self::ConditionalFormatting12(value) => (CF12, encode_sdk(value)?),
            Self::ConditionalFormattingGroup12(value) => (COND_FMT12, encode_sdk(value)?),
            Self::Theme(value) => (THEME, encode_sdk(value)?),
            Self::ExtendedHeaderFooter(value) => (HEADER_FOOTER_EXT, encode_sdk(value)?),
            Self::ShapePropsStream(value) => (SHAPE_PROPS_STREAM, encode_sdk(value)?),
            Self::TextPropsStream(value) => (TEXT_PROPS_STREAM, encode_sdk(value)?),
            Self::Compat12(value) => (COMPAT12, encode_sdk(value)?),
            Self::Plv(value) => (PLV, encode_sdk(value)?),
            Self::PlvMac(value) => (PLV_MAC, encode_sdk(value)?),
            Self::Lnext(value) => (LNEXT, encode_sdk(value)?),
            Self::MkrExt(value) => (MKR_EXT, encode_sdk(value)?),
            Self::CrtCoOpt(value) => (CRT_CO_OPT, encode_sdk(value)?),
            Self::FrtArchId(value) => (FRT_ARCH_ID, encode_sdk(value)?),
            Self::CrtLayout12(value) => (CRT_LAYOUT12, encode_sdk(value)?),
            Self::CrtLayout12A(value) => (CRT_LAYOUT12_A, encode_sdk(value)?),
            Self::MtrSettings(value) => (MTR_SETTINGS, encode_sdk(value)?),
            Self::ForceFullCalculation(value) => (FORCE_FULL_CALCULATION, encode_sdk(value)?),
            Self::CompressPictures(value) => (COMPRESS_PICTURES, encode_sdk(value)?),
            Self::CrtMlFrt(value) => (CRT_ML_FRT, encode_sdk(value)?),
            Self::ChartFrtInfo(value) => (CHART_FRT_INFO, encode_sdk(value)?),
            Self::ChartCatLab(value) => (CHART_CAT_LAB, encode_sdk(value)?),
            Self::ChartStartObject(value) => (CHART_START_OBJECT, encode_sdk(value)?),
            Self::ChartEndObject(value) => (CHART_END_OBJECT, encode_sdk(value)?),
            Self::GelFrame(value) => (GEL_FRAME, value.to_bytes()?),
            Self::HfPicture(value) => (HF_PICTURE, encode_sdk(value)?),
            Self::FeatureHeader(value) => (FEAT_HDR, encode_sdk(value)?),
            Self::Feature(value) => (FEAT, encode_sdk(value)?),
            Self::BookExt(value) => (BOOK_EXT, encode_sdk(value)?),
            Self::Chart(value) => (CHART, encode_sdk(value)?),
            Self::ChartAreaFormat(value) => (CHART_AREA_FORMAT, encode_sdk(value)?),
            Self::ChartAttachedLabel(value) => (CHART_ATTACHED_LABEL, encode_sdk(value)?),
            Self::ChartDataFormat(value) => (CHART_DATA_FORMAT, encode_sdk(value)?),
            Self::ChartFormat(value) => (CHART_FORMAT, encode_sdk(value)?),
            Self::ChartSeriesList(value) => (CHART_SERIES_LIST, encode_sdk(value)?),
            Self::ChartBar(value) => (CHART_BAR, encode_sdk(value)?),
            Self::ChartLine(value) => (CHART_LINE, encode_sdk(value)?),
            Self::ChartPie(value) => (CHART_PIE, encode_sdk(value)?),
            Self::ChartArea(value) => (CHART_AREA, encode_sdk(value)?),
            Self::ChartScatter(value) => (CHART_SCATTER, encode_sdk(value)?),
            Self::ChartCrtLine(value) => (CHART_CRT_LINE, encode_sdk(value)?),
            Self::ChartCrtLink(value) => (CHART_CRT_LINK, encode_sdk(value)?),
            Self::ChartLegend(value) => (CHART_LEGEND, encode_sdk(value)?),
            Self::ChartAxis(value) => (CHART_AXIS, encode_sdk(value)?),
            Self::ChartTick(value) => (CHART_TICK, encode_sdk(value)?),
            Self::ChartValueRange(value) => (CHART_VALUE_RANGE, encode_sdk(value)?),
            Self::ChartLabelRange(value) => (CHART_LABEL_RANGE, encode_sdk(value)?),
            Self::ChartAxisLine(value) => (CHART_AXIS_LINE, encode_sdk(value)?),
            Self::ChartDefaultText(value) => (CHART_DEFAULT_TEXT, encode_sdk(value)?),
            Self::ChartFont(value) => (CHART_FONT, encode_sdk(value)?),
            Self::ChartLineFormat(value) => (CHART_LINE_FORMAT, encode_sdk(value)?),
            Self::ChartMarkerFormat(value) => (CHART_MARKER_FORMAT, encode_sdk(value)?),
            Self::ChartObjectLink(value) => (CHART_OBJECT_LINK, encode_sdk(value)?),
            Self::ChartFrame(value) => (CHART_FRAME, encode_sdk(value)?),
            Self::Chart3D(value) => (CHART_3D, encode_sdk(value)?),
            Self::ChartDropBar(value) => (CHART_DROP_BAR, encode_sdk(value)?),
            Self::ChartSurf(value) => (CHART_SURF, encode_sdk(value)?),
            Self::ChartLegendException(value) => (CHART_LEGEND_EXCEPTION, encode_sdk(value)?),
            Self::ChartAxisParent(value) => (CHART_AXIS_PARENT, encode_sdk(value)?),
            Self::ChartSheetProperties(value) => (CHART_SHEET_PROPERTIES, encode_sdk(value)?),
            Self::ChartSeriesGroupIndex(value) => (CHART_SERIES_GROUP_INDEX, encode_sdk(value)?),
            Self::ChartAxisUsed(value) => (CHART_AXIS_USED, encode_sdk(value)?),
            Self::ChartNumberFormatIndex(value) => (CHART_NUMBER_FORMAT_INDEX, encode_sdk(value)?),
            Self::ChartSeriesParent(value) => (CHART_SERIES_PARENT, encode_sdk(value)?),
            Self::ChartSeriesAuxTrend(value) => (CHART_SERIES_AUX_TREND, encode_sdk(value)?),
            Self::ChartPosition(value) => (CHART_POSITION, encode_sdk(value)?),
            Self::ChartFontBasis(value) => (CHART_FONT_BASIS, encode_sdk(value)?),
            Self::Chart3DBarShape(value) => (CHART_3D_BAR_SHAPE, encode_sdk(value)?),
            Self::ChartSeriesFormat(value) => (CHART_SERIES_FORMAT, encode_sdk(value)?),
            Self::ChartSeriesAuxErrorBar(value) => (CHART_SERIES_AUX_ERROR_BAR, encode_sdk(value)?),
            Self::ChartClrtClient(value) => (CHART_CLRT_CLIENT, encode_sdk(value)?),
            Self::ChartAxisOptions(value) => (CHART_AXIS_OPTIONS, encode_sdk(value)?),
            Self::ChartDat(value) => (CHART_DAT, encode_sdk(value)?),
            Self::ChartPieFormat(value) => (CHART_PIE_FORMAT, encode_sdk(value)?),
            Self::ChartPlotGrowth(value) => (CHART_PLOT_GROWTH, encode_sdk(value)?),
            Self::ChartLinkedData(value) => (CHART_LINKED_DATA, encode_sdk(value)?),
            Self::ChartAlRuns(value) => (CHART_AL_RUNS, encode_sdk(value)?),
            Self::ChartSeriesIndex(value) => (CHART_SERIES_INDEX, encode_sdk(value)?),
            Self::ChartSeries(value) => (CHART_SERIES, encode_sdk(value)?),
            Self::ChartSeriesCompatibility { record_type, value } => {
                (*record_type, encode_sdk(value)?)
            }
            Self::ChartSeriesText(value) => (CHART_SERIES_TEXT, encode_sdk(value)?),
            Self::ChartText(value) => (CHART_TEXT, encode_sdk(value)?),
            Self::StartBlock(value) => (START_BLOCK, encode_sdk(value)?),
            Self::EndBlock(value) => (END_BLOCK, encode_sdk(value)?),
            Self::DbCell(value) => (DB_CELL, encode_sdk(value)?),
            Self::Font(value) => (FONT, encode_sdk(value)?),
            Self::FontCompatibility { record_type, value } => (*record_type, encode_sdk(value)?),
            Self::Format(value) => (FORMAT, encode_sdk(value)?),
            Self::Style(value) => (STYLE, encode_sdk(value)?),
            Self::CrnCount(value) => (CRN_COUNT, encode_sdk(value)?),
            Self::DefaultRowHeight(value) => (DEFAULT_ROW_HEIGHT, encode_sdk(value)?),
            Self::WriteAccess(value) => (WRITE_ACCESS, encode_sdk(value)?),
            Self::Window2(value) => (WINDOW2, encode_sdk(value)?),
            Self::Selection(value) => (SELECTION, encode_sdk(value)?),
            Self::MergeCells(value) => (MERGE_CELLS, encode_sdk(value)?),
            Self::Mms(value) => (MMS, encode_sdk(value)?),
            Self::PhoneticInfo(value) => (PHONETIC_INFO, encode_sdk(value)?),
            Self::Sst(value) => {
                let first = value
                    .encode_physical()?
                    .into_iter()
                    .next()
                    .ok_or_else(|| Error::invalid(0, "SST has no physical segments"))?;
                (SST, first.payload)
            }
            Self::ExtSst(value) => (EXT_SST, encode_sdk(value)?),
            Self::Index(value) => (INDEX, encode_sdk(value)?),
            Self::StringValue(value) => {
                let first = value
                    .encode_physical()?
                    .into_iter()
                    .next()
                    .ok_or_else(|| Error::invalid(0, "String record has no physical chunks"))?;
                (STRING_VALUE, first.payload)
            }
            Self::FixedU16 { kind, value } => (kind.id(), value.to_le_bytes().to_vec()),
            Self::FixedF64Bits { kind, bits } => (kind.id(), bits.to_le_bytes().to_vec()),
            Self::Empty { kind, reserved } => (
                kind.id(),
                reserved.map_or_else(Vec::new, |value| value.to_le_bytes().to_vec()),
            ),
            Self::FilePass { payload } => (FILE_PASS, payload.clone()),
            Self::Encrypted {
                record_type,
                payload,
            }
            | Self::Unknown {
                record_type,
                payload,
            } => (*record_type, payload.clone()),
            Self::Continue { payload } => (CONTINUE, payload.clone()),
        })
    }
}

struct EncodedBiffRecord {
    record_type: u16,
    payload: Vec<u8>,
}

impl TxoContext {
    fn parse(bytes: [u8; 6], object_type: Option<u16>) -> Self {
        let word1 = u16::from_le_bytes([bytes[0], bytes[1]]);
        let word2 = u16::from_le_bytes([bytes[2], bytes[3]]);
        let word3 = u16::from_le_bytes([bytes[4], bytes[5]]);
        if matches!(
            object_type,
            Some(0x0000 | 0x0005 | 0x0007 | 0x000b | 0x000c | 0x000e)
        ) {
            Self::Control(ObjControlInfo {
                flags: ObjControlInfoFlags::from_bits_retain(word1),
                accelerator: i16::from_le_bytes([bytes[2], bytes[3]]),
                reserved: word3,
            })
        } else if object_type.is_some() {
            Self::Reserved {
                reserved4: word1,
                reserved5: u32::from_le_bytes(bytes[2..6].try_into().expect("four bytes")),
            }
        } else {
            Self::Undetermined {
                word1,
                word2,
                word3,
            }
        }
    }

    fn to_bytes(self) -> [u8; 6] {
        let mut bytes = [0; 6];
        match self {
            Self::Reserved {
                reserved4,
                reserved5,
            } => {
                bytes[0..2].copy_from_slice(&reserved4.to_le_bytes());
                bytes[2..6].copy_from_slice(&reserved5.to_le_bytes());
            }
            Self::Control(value) => {
                bytes[0..2].copy_from_slice(&value.flags.bits().to_le_bytes());
                bytes[2..4].copy_from_slice(&value.accelerator.to_le_bytes());
                bytes[4..6].copy_from_slice(&value.reserved.to_le_bytes());
            }
            Self::Undetermined {
                word1,
                word2,
                word3,
            } => {
                bytes[0..2].copy_from_slice(&word1.to_le_bytes());
                bytes[2..4].copy_from_slice(&word2.to_le_bytes());
                bytes[4..6].copy_from_slice(&word3.to_le_bytes());
            }
        }
        bytes
    }
}

impl TxoRecord {
    fn from_sequence(
        payload: &[u8],
        following: &[BiffRecord],
        object_type: Option<u16>,
        limits: Limits,
    ) -> Result<(Self, usize)> {
        if payload.len() < 18 {
            return Err(Error::invalid(0, "TxO record is shorter than 18 bytes"));
        }
        let declared_text_length = u16::from_le_bytes([payload[10], payload[11]]);
        let declared_run_data_length = u16::from_le_bytes([payload[12], payload[13]]);
        let formula_length = usize::from(u16::from_le_bytes([payload[16], payload[17]]));
        let formula_end = 18usize
            .checked_add(formula_length)
            .ok_or_else(|| Error::Limit("TxO formula length overflow".into()))?;
        if formula_end > payload.len() {
            return Err(Error::invalid(16, "TxO ObjFmla is truncated"));
        }
        if formula_length > limits.max_allocation {
            return Err(Error::Limit("TxO formula exceeds allocation limit".into()));
        }

        let mut consumed = 0usize;
        let mut remaining_characters = usize::from(declared_text_length);
        let mut text_chunks = Vec::new();
        while remaining_characters != 0 {
            let continuation = continue_payload(following, consumed, "TxO text")?;
            let Some((&flags, character_bytes)) = continuation.split_first() else {
                return Err(Error::invalid(0, "TxO text Continue is empty"));
            };
            let width = if flags & 1 == 0 { 1 } else { 2 };
            if character_bytes.len() % width != 0 {
                return Err(Error::invalid(0, "TxO text character bytes are misaligned"));
            }
            let character_count = character_bytes.len() / width;
            if character_count == 0 || character_count > remaining_characters {
                return Err(Error::invalid(0, "TxO text Continue exceeds cchText"));
            }
            let mut reader = Reader::new(Cursor::new(continuation))?;
            text_chunks.push(BiffUnicodeString::read(&mut reader, character_count)?);
            if reader.remaining()? != 0 {
                return Err(Error::invalid(0, "TxO text Continue has trailing bytes"));
            }
            remaining_characters -= character_count;
            consumed += 1;
        }

        let mut remaining_runs = usize::from(declared_run_data_length);
        let mut formatting_segment_lengths = Vec::new();
        let mut run_bytes = Vec::with_capacity(remaining_runs);
        if remaining_runs > limits.max_allocation {
            return Err(Error::Limit("TxO run data exceeds allocation limit".into()));
        }
        while remaining_runs != 0 {
            let continuation = continue_payload(following, consumed, "TxO formatting")?;
            if continuation.is_empty() || continuation.len() > remaining_runs {
                return Err(Error::invalid(0, "TxO formatting exceeds cbRuns"));
            }
            formatting_segment_lengths.push(
                u16::try_from(continuation.len()).map_err(|_| {
                    Error::Limit("TxO formatting Continue length exceeds u16".into())
                })?,
            );
            run_bytes.extend_from_slice(continuation);
            remaining_runs -= continuation.len();
            consumed += 1;
        }
        if !run_bytes.len().is_multiple_of(8) {
            return Err(Error::invalid(0, "TxO run data is not 8-byte aligned"));
        }
        let mut parsed_runs = run_bytes
            .chunks_exact(8)
            .map(TxoRun::parse)
            .collect::<Vec<_>>();
        let last_run = parsed_runs.pop();

        Ok((
            Self {
                options: TxoOptions::from_bits_retain(u16::from_le_bytes([payload[0], payload[1]])),
                rotation: TxoRotation(u16::from_le_bytes([payload[2], payload[3]])),
                context: TxoContext::parse(
                    payload[4..10].try_into().expect("six bytes"),
                    object_type,
                ),
                declared_text_length,
                declared_run_data_length,
                empty_font_index: u16::from_le_bytes([payload[14], payload[15]]),
                formula: ObjFormula::parse(
                    u16::from_le_bytes([payload[16], payload[17]]),
                    &payload[18..formula_end],
                )?,
                trailing: payload[formula_end..].to_vec(),
                text_chunks,
                runs: parsed_runs,
                last_run,
                formatting_segment_lengths,
            },
            consumed,
        ))
    }

    fn encode_physical(&self) -> Result<Vec<EncodedBiffRecord>> {
        let formula_data = self.formula.to_bytes()?;
        if usize::from(self.formula.declared_length) != formula_data.len() {
            return Err(Error::invalid(0, "TxO ObjFmla length mismatch"));
        }
        let text_count = self
            .text_chunks
            .iter()
            .map(BiffUnicodeString::character_count)
            .sum::<usize>();
        if text_count != usize::from(self.declared_text_length) {
            return Err(Error::invalid(0, "TxO cchText mismatch"));
        }
        let mut payload = Vec::new();
        payload.extend_from_slice(&self.options.bits().to_le_bytes());
        payload.extend_from_slice(&self.rotation.0.to_le_bytes());
        payload.extend_from_slice(&self.context.to_bytes());
        payload.extend_from_slice(&self.declared_text_length.to_le_bytes());
        payload.extend_from_slice(&self.declared_run_data_length.to_le_bytes());
        payload.extend_from_slice(&self.empty_font_index.to_le_bytes());
        payload.extend_from_slice(&self.formula.declared_length.to_le_bytes());
        payload.extend_from_slice(&formula_data);
        payload.extend_from_slice(&self.trailing);
        let mut records = vec![EncodedBiffRecord {
            record_type: TXO,
            payload,
        }];
        for chunk in &self.text_chunks {
            records.push(EncodedBiffRecord {
                record_type: CONTINUE,
                payload: chunk.to_bytes()?,
            });
        }
        let mut run_bytes = Vec::new();
        for run in self.runs.iter().chain(self.last_run.iter()) {
            run.write(&mut run_bytes);
        }
        if run_bytes.len() != usize::from(self.declared_run_data_length)
            || self
                .formatting_segment_lengths
                .iter()
                .map(|length| usize::from(*length))
                .sum::<usize>()
                != run_bytes.len()
        {
            return Err(Error::invalid(0, "TxO cbRuns or physical layout mismatch"));
        }
        let mut offset = 0usize;
        for length in &self.formatting_segment_lengths {
            let end = offset + usize::from(*length);
            records.push(EncodedBiffRecord {
                record_type: CONTINUE,
                payload: run_bytes[offset..end].to_vec(),
            });
            offset = end;
        }
        Ok(records)
    }
}

impl BiffUnicodeString {
    fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        self.write(&mut writer)?;
        Ok(writer.into_inner().into_inner())
    }
}

impl ObjFormula {
    fn parse(declared_length: u16, bytes: &[u8]) -> Result<Self> {
        let data = if bytes.is_empty() {
            ObjFormulaData::Empty
        } else if bytes.len() >= 6 {
            let cce_and_reserved = u16::from_le_bytes([bytes[0], bytes[1]]);
            let cce = usize::from(cce_and_reserved & 0x7fff);
            let rgce_end = 6usize
                .checked_add(cce)
                .ok_or_else(|| Error::Limit("ObjFmla rgce length overflow".into()))?;
            if rgce_end <= bytes.len() {
                ObjFormulaData::Parsed {
                    cce_and_reserved,
                    unused: u32::from_le_bytes(bytes[2..6].try_into().expect("four bytes")),
                    tokens: FormulaTokenStream::from_bytes(&bytes[6..rgce_end])?,
                    padding: bytes[rgce_end..].to_vec(),
                }
            } else {
                ObjFormulaData::Opaque(bytes.to_vec())
            }
        } else {
            ObjFormulaData::Opaque(bytes.to_vec())
        };
        Ok(Self {
            declared_length,
            data,
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        let bytes = match &self.data {
            ObjFormulaData::Empty => Vec::new(),
            ObjFormulaData::Parsed {
                cce_and_reserved,
                unused,
                tokens,
                padding,
            } => {
                let rgce = tokens.to_bytes()?;
                if usize::from(cce_and_reserved & 0x7fff) != rgce.len() {
                    return Err(Error::invalid(0, "ObjFmla cce mismatch"));
                }
                let mut bytes = Vec::with_capacity(6 + rgce.len() + padding.len());
                bytes.extend_from_slice(&cce_and_reserved.to_le_bytes());
                bytes.extend_from_slice(&unused.to_le_bytes());
                bytes.extend_from_slice(&rgce);
                bytes.extend_from_slice(padding);
                bytes
            }
            ObjFormulaData::Opaque(bytes) => bytes.clone(),
        };
        Ok(bytes)
    }
}

impl TxoRun {
    fn parse(bytes: &[u8]) -> Self {
        Self {
            format: FormatRun {
                character_index: u16::from_le_bytes([bytes[0], bytes[1]]),
                font_index: u16::from_le_bytes([bytes[2], bytes[3]]),
            },
            unused1: u16::from_le_bytes([bytes[4], bytes[5]]),
            unused2: u16::from_le_bytes([bytes[6], bytes[7]]),
        }
    }

    fn write(&self, bytes: &mut Vec<u8>) {
        bytes.extend_from_slice(&self.format.character_index.to_le_bytes());
        bytes.extend_from_slice(&self.format.font_index.to_le_bytes());
        bytes.extend_from_slice(&self.unused1.to_le_bytes());
        bytes.extend_from_slice(&self.unused2.to_le_bytes());
    }
}

impl NoteRecord {
    fn from_bytes(bytes: &[u8], limits: Limits) -> Result<Self> {
        if bytes.len() < 12 {
            return Err(Error::invalid(0, "NoteSh record is truncated"));
        }
        let declared_author_length = u16::from_le_bytes([bytes[8], bytes[9]]);
        if usize::from(declared_author_length) > limits.max_allocation {
            return Err(Error::Limit(
                "NoteSh author exceeds allocation limit".into(),
            ));
        }
        let mut reader = Reader::new(Cursor::new(&bytes[10..]))?;
        let author = BiffUnicodeString::read(&mut reader, usize::from(declared_author_length))?;
        let consumed = usize::try_from(reader.position()?)
            .map_err(|_| Error::Limit("NoteSh author length exceeds usize".into()))?;
        let unused_offset = 10usize
            .checked_add(consumed)
            .ok_or_else(|| Error::Limit("NoteSh author offset overflow".into()))?;
        let unused = *bytes
            .get(unused_offset)
            .ok_or_else(|| Error::invalid(unused_offset as u64, "NoteSh unused byte is missing"))?;
        Ok(Self {
            row: u16::from_le_bytes([bytes[0], bytes[1]]),
            column: u16::from_le_bytes([bytes[2], bytes[3]]),
            flags: NoteFlags::from_bits_retain(u16::from_le_bytes([bytes[4], bytes[5]])),
            object_id: u16::from_le_bytes([bytes[6], bytes[7]]),
            declared_author_length,
            author,
            unused,
            trailing: bytes[unused_offset + 1..].to_vec(),
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.author.character_count() != usize::from(self.declared_author_length) {
            return Err(Error::invalid(0, "NoteSh author length mismatch"));
        }
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&self.row.to_le_bytes());
        bytes.extend_from_slice(&self.column.to_le_bytes());
        bytes.extend_from_slice(&self.flags.bits().to_le_bytes());
        bytes.extend_from_slice(&self.object_id.to_le_bytes());
        bytes.extend_from_slice(&self.declared_author_length.to_le_bytes());
        bytes.extend_from_slice(&self.author.to_bytes()?);
        bytes.push(self.unused);
        bytes.extend_from_slice(&self.trailing);
        Ok(bytes)
    }
}

impl ObjRecord {
    fn object_type(&self) -> Option<u16> {
        self.subrecords.iter().find_map(|subrecord| {
            if let ObjSubrecordData::Common(common) = &subrecord.data {
                Some(common.object_type)
            } else {
                None
            }
        })
    }

    fn from_bytes(bytes: &[u8], limits: Limits) -> Result<Self> {
        let mut cursor = 0usize;
        let mut subrecords = Vec::new();
        let mut object_type = None;
        while cursor < bytes.len() {
            if bytes.len() - cursor < 4 {
                break;
            }
            let subrecord_type = u16::from_le_bytes([bytes[cursor], bytes[cursor + 1]]);
            let declared_length = u16::from_le_bytes([bytes[cursor + 2], bytes[cursor + 3]]);
            cursor += 4;
            // FtLbsData's cbFContinued is explicitly not an ordinary payload length and the
            // structure terminates the Obj stream. Its bytes therefore occupy the remainder.
            let length = if subrecord_type == 0x0013 {
                bytes.len() - cursor
            } else {
                usize::from(declared_length)
            };
            if length > limits.max_allocation {
                return Err(Error::Limit(
                    "Obj subrecord exceeds allocation limit".into(),
                ));
            }
            let end = cursor
                .checked_add(length)
                .ok_or_else(|| Error::Limit("Obj subrecord offset overflow".into()))?;
            let payload = bytes
                .get(cursor..end)
                .ok_or_else(|| Error::invalid(cursor as u64, "Obj subrecord is truncated"))?;
            cursor = end;
            let data = ObjSubrecordData::parse(
                subrecord_type,
                declared_length,
                payload,
                object_type,
                limits,
            )?;
            if let ObjSubrecordData::Common(common) = &data {
                object_type = Some(common.object_type);
            }
            subrecords.push(ObjSubrecord {
                subrecord_type,
                declared_length,
                data,
            });
            if subrecord_type == 0x0000 || subrecord_type == 0x0013 {
                break;
            }
            if subrecords.len() > limits.max_entries {
                return Err(Error::Limit("Obj subrecord count exceeds limit".into()));
            }
        }
        if subrecords.is_empty() {
            return Err(Error::invalid(0, "Obj has no complete subrecord"));
        }
        Ok(Self {
            subrecords,
            trailing: bytes[cursor..].to_vec(),
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        for subrecord in &self.subrecords {
            let payload = subrecord.data.to_bytes()?;
            if subrecord.subrecord_type != 0x0013
                && usize::from(subrecord.declared_length) != payload.len()
            {
                return Err(Error::invalid(0, "Obj subrecord length mismatch"));
            }
            bytes.extend_from_slice(&subrecord.subrecord_type.to_le_bytes());
            bytes.extend_from_slice(&subrecord.declared_length.to_le_bytes());
            bytes.extend_from_slice(&payload);
        }
        bytes.extend_from_slice(&self.trailing);
        Ok(bytes)
    }
}

impl ObjSubrecordData {
    fn parse(
        subrecord_type: u16,
        declared_length: u16,
        bytes: &[u8],
        object_type: Option<u16>,
        limits: Limits,
    ) -> Result<Self> {
        let exact = |expected: usize, name: &str| {
            if bytes.len() == expected {
                Ok(())
            } else {
                Err(Error::invalid(
                    0,
                    format!("{name} has {} bytes, expected {expected}", bytes.len()),
                ))
            }
        };
        Ok(match subrecord_type {
            0x0000 if bytes.is_empty() => Self::End,
            0x0004 => Self::Macro(ObjFormula::parse(declared_length, bytes)?),
            0x0006 => {
                exact(2, "FtGmo")?;
                Self::GroupMarker {
                    unused: u16::from_le_bytes([bytes[0], bytes[1]]),
                }
            }
            0x0007 => {
                exact(2, "FtCf")?;
                Self::ClipboardFormat {
                    format: u16::from_le_bytes([bytes[0], bytes[1]]),
                }
            }
            0x0008 => match bytes {
                [low, high] => {
                    Self::PictureFlags(ObjPictureFlags::from_bits_retain(u16::from_le_bytes([
                        *low, *high,
                    ])))
                }
                [low] => Self::TruncatedPictureFlags { low_byte: *low },
                _ => {
                    return Err(Error::invalid(
                        0,
                        format!("FtPioGrbit has {} bytes, expected 1 or 2", bytes.len()),
                    ));
                }
            },
            0x0001 | 0x0106 if bytes.is_empty() => Self::EmptyCompatibilityMarker,
            0x0009 if bytes.len() >= 2 => {
                let formula_length = u16::from_le_bytes([bytes[0], bytes[1]]);
                let end = 2usize
                    .checked_add(usize::from(formula_length))
                    .ok_or_else(|| Error::Limit("FtPictFmla formula length overflow".into()))?;
                let formula_bytes = bytes
                    .get(2..end)
                    .ok_or_else(|| Error::invalid(0, "FtPictFmla formula is truncated"))?;
                Self::PictureFormula(ObjPictureFormula {
                    formula: ObjFormula::parse(formula_length, formula_bytes)?,
                    trailing: bytes[end..].to_vec(),
                })
            }
            0x000a if bytes.len() == 8 => Self::CheckBox(ObjCheckBoxStructure::Legacy {
                unused1: u32::from_le_bytes(bytes[0..4].try_into().expect("four bytes")),
                unused2: u32::from_le_bytes(bytes[4..8].try_into().expect("four bytes")),
            }),
            0x000a if bytes.len() == 12 => Self::CheckBox(ObjCheckBoxStructure::Full {
                unused1: u32::from_le_bytes(bytes[0..4].try_into().expect("four bytes")),
                unused2: u32::from_le_bytes(bytes[4..8].try_into().expect("four bytes")),
                unused3: u32::from_le_bytes(bytes[8..12].try_into().expect("four bytes")),
            }),
            0x000b => {
                exact(6, "FtRbo")?;
                Self::RadioButton {
                    unused1: u32::from_le_bytes(bytes[0..4].try_into().expect("four bytes")),
                    unused2: u16::from_le_bytes([bytes[4], bytes[5]]),
                }
            }
            0x000c => Self::ScrollBar(ObjScrollBarData::parse(bytes)?),
            0x000d => Self::Note(ObjNoteData::parse(bytes)?),
            0x000e => Self::ScrollBarFormula(ObjFormula::parse(declared_length, bytes)?),
            0x000f => Self::GroupBox(ObjGroupBoxData::parse(bytes)?),
            0x0010 => Self::EditBox(ObjEditBoxData::parse(bytes)?),
            0x0011 => Self::RadioButtonData(ObjRadioButtonData::parse(bytes)?),
            0x0012 if bytes.len() >= 8 => Self::CheckBoxData(ObjCheckBoxData::parse(bytes)),
            0x0013 => Self::ListBox(ObjListBoxData::parse(
                declared_length,
                bytes,
                object_type,
                limits,
            )?),
            0x0014 => Self::CheckBoxFormula(ObjFormula::parse(declared_length, bytes)?),
            0x0015 if bytes.len() == 18 => Self::Common(ObjCommonData::parse(bytes)),
            0x003f if bytes.len().is_multiple_of(2) => Self::Compatibility(ObjCompatibilityData {
                words: bytes
                    .chunks_exact(2)
                    .map(|word| u16::from_le_bytes([word[0], word[1]]))
                    .collect(),
            }),
            _ => Self::Raw(bytes.to_vec()),
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(match self {
            Self::End => Vec::new(),
            Self::Common(value) => value.to_bytes().to_vec(),
            Self::Macro(value) | Self::ScrollBarFormula(value) | Self::CheckBoxFormula(value) => {
                value.to_bytes()?
            }
            Self::GroupMarker { unused } => unused.to_le_bytes().to_vec(),
            Self::ClipboardFormat { format } => format.to_le_bytes().to_vec(),
            Self::PictureFlags(flags) => flags.bits().to_le_bytes().to_vec(),
            Self::TruncatedPictureFlags { low_byte } => vec![*low_byte],
            Self::EmptyCompatibilityMarker => Vec::new(),
            Self::PictureFormula(value) => value.to_bytes()?,
            Self::CheckBox(value) => value.to_bytes(),
            Self::RadioButton { unused1, unused2 } => {
                let mut bytes = unused1.to_le_bytes().to_vec();
                bytes.extend_from_slice(&unused2.to_le_bytes());
                bytes
            }
            Self::ScrollBar(value) => value.to_bytes().to_vec(),
            Self::Note(value) => value.to_bytes().to_vec(),
            Self::GroupBox(value) => value.to_bytes().to_vec(),
            Self::EditBox(value) => value.to_bytes().to_vec(),
            Self::RadioButtonData(value) => value.to_bytes().to_vec(),
            Self::CheckBoxData(value) => value.to_bytes(),
            Self::ListBox(value) => value.to_bytes()?,
            Self::Compatibility(value) => value
                .words
                .iter()
                .flat_map(|word| word.to_le_bytes())
                .collect(),
            Self::Raw(bytes) => bytes.clone(),
        })
    }
}

impl ObjCommonData {
    fn parse(bytes: &[u8]) -> Self {
        Self {
            object_type: u16::from_le_bytes([bytes[0], bytes[1]]),
            object_id: u16::from_le_bytes([bytes[2], bytes[3]]),
            flags: ObjCommonFlags::from_bits_retain(u16::from_le_bytes([bytes[4], bytes[5]])),
            reserved1: u32::from_le_bytes(bytes[6..10].try_into().expect("four bytes")),
            reserved2: u32::from_le_bytes(bytes[10..14].try_into().expect("four bytes")),
            reserved3: u32::from_le_bytes(bytes[14..18].try_into().expect("four bytes")),
        }
    }

    fn to_bytes(self) -> [u8; 18] {
        let mut bytes = [0; 18];
        bytes[0..2].copy_from_slice(&self.object_type.to_le_bytes());
        bytes[2..4].copy_from_slice(&self.object_id.to_le_bytes());
        bytes[4..6].copy_from_slice(&self.flags.bits().to_le_bytes());
        bytes[6..10].copy_from_slice(&self.reserved1.to_le_bytes());
        bytes[10..14].copy_from_slice(&self.reserved2.to_le_bytes());
        bytes[14..18].copy_from_slice(&self.reserved3.to_le_bytes());
        bytes
    }
}

impl ObjPictureFormula {
    fn to_bytes(&self) -> Result<Vec<u8>> {
        let formula = self.formula.to_bytes()?;
        if formula.len() != usize::from(self.formula.declared_length) {
            return Err(Error::invalid(0, "FtPictFmla ObjFmla length mismatch"));
        }
        let mut bytes = self.formula.declared_length.to_le_bytes().to_vec();
        bytes.extend_from_slice(&formula);
        bytes.extend_from_slice(&self.trailing);
        Ok(bytes)
    }
}

impl ObjCheckBoxStructure {
    fn to_bytes(self) -> Vec<u8> {
        let (unused1, unused2, unused3) = match self {
            Self::Legacy { unused1, unused2 } => (unused1, unused2, None),
            Self::Full {
                unused1,
                unused2,
                unused3,
            } => (unused1, unused2, Some(unused3)),
        };
        let mut bytes = unused1.to_le_bytes().to_vec();
        bytes.extend_from_slice(&unused2.to_le_bytes());
        if let Some(value) = unused3 {
            bytes.extend_from_slice(&value.to_le_bytes());
        }
        bytes
    }
}

impl ObjScrollBarData {
    fn parse(bytes: &[u8]) -> Result<Self> {
        if bytes.len() != 20 {
            return Err(Error::invalid(0, "FtSbs must contain 20 bytes"));
        }
        let word = |offset| i16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
        Ok(Self {
            unused: u32::from_le_bytes(bytes[0..4].try_into().expect("four bytes")),
            value: word(4),
            minimum: word(6),
            maximum: word(8),
            increment: word(10),
            page_increment: word(12),
            horizontal: u16::from_le_bytes([bytes[14], bytes[15]]),
            width: word(16),
            flags: ObjScrollBarFlags::from_bits_retain(u16::from_le_bytes([bytes[18], bytes[19]])),
        })
    }

    fn to_bytes(self) -> [u8; 20] {
        let mut bytes = [0; 20];
        bytes[0..4].copy_from_slice(&self.unused.to_le_bytes());
        bytes[4..6].copy_from_slice(&self.value.to_le_bytes());
        bytes[6..8].copy_from_slice(&self.minimum.to_le_bytes());
        bytes[8..10].copy_from_slice(&self.maximum.to_le_bytes());
        bytes[10..12].copy_from_slice(&self.increment.to_le_bytes());
        bytes[12..14].copy_from_slice(&self.page_increment.to_le_bytes());
        bytes[14..16].copy_from_slice(&self.horizontal.to_le_bytes());
        bytes[16..18].copy_from_slice(&self.width.to_le_bytes());
        bytes[18..20].copy_from_slice(&self.flags.bits().to_le_bytes());
        bytes
    }
}

impl ObjNoteData {
    fn parse(bytes: &[u8]) -> Result<Self> {
        if bytes.len() != 22 {
            return Err(Error::invalid(0, "FtNts must contain 22 bytes"));
        }
        Ok(Self {
            guid: bytes[0..16].try_into().expect("sixteen bytes"),
            shared: u16::from_le_bytes([bytes[16], bytes[17]]),
            unused: u32::from_le_bytes(bytes[18..22].try_into().expect("four bytes")),
        })
    }

    fn to_bytes(self) -> [u8; 22] {
        let mut bytes = [0; 22];
        bytes[0..16].copy_from_slice(&self.guid);
        bytes[16..18].copy_from_slice(&self.shared.to_le_bytes());
        bytes[18..22].copy_from_slice(&self.unused.to_le_bytes());
        bytes
    }
}

impl ObjGroupBoxData {
    fn parse(bytes: &[u8]) -> Result<Self> {
        if bytes.len() != 6 {
            return Err(Error::invalid(0, "FtGboData must contain 6 bytes"));
        }
        Ok(Self {
            accelerator: u16::from_le_bytes([bytes[0], bytes[1]]),
            reserved: u16::from_le_bytes([bytes[2], bytes[3]]),
            flags: u16::from_le_bytes([bytes[4], bytes[5]]),
        })
    }

    fn to_bytes(self) -> [u8; 6] {
        let mut bytes = [0; 6];
        bytes[0..2].copy_from_slice(&self.accelerator.to_le_bytes());
        bytes[2..4].copy_from_slice(&self.reserved.to_le_bytes());
        bytes[4..6].copy_from_slice(&self.flags.to_le_bytes());
        bytes
    }
}

impl ObjEditBoxData {
    fn parse(bytes: &[u8]) -> Result<Self> {
        if bytes.len() != 8 {
            return Err(Error::invalid(0, "FtEdoData must contain 8 bytes"));
        }
        Ok(Self {
            validation_type: u16::from_le_bytes([bytes[0], bytes[1]]),
            multiline: u16::from_le_bytes([bytes[2], bytes[3]]),
            vertical_scroll: u16::from_le_bytes([bytes[4], bytes[5]]),
            list_object_id: u16::from_le_bytes([bytes[6], bytes[7]]),
        })
    }

    fn to_bytes(self) -> [u8; 8] {
        let mut bytes = [0; 8];
        bytes[0..2].copy_from_slice(&self.validation_type.to_le_bytes());
        bytes[2..4].copy_from_slice(&self.multiline.to_le_bytes());
        bytes[4..6].copy_from_slice(&self.vertical_scroll.to_le_bytes());
        bytes[6..8].copy_from_slice(&self.list_object_id.to_le_bytes());
        bytes
    }
}

impl ObjRadioButtonData {
    fn parse(bytes: &[u8]) -> Result<Self> {
        if bytes.len() != 4 {
            return Err(Error::invalid(0, "FtRboData must contain 4 bytes"));
        }
        Ok(Self {
            next_object_id: u16::from_le_bytes([bytes[0], bytes[1]]),
            first_button: u16::from_le_bytes([bytes[2], bytes[3]]),
        })
    }

    fn to_bytes(self) -> [u8; 4] {
        let mut bytes = [0; 4];
        bytes[0..2].copy_from_slice(&self.next_object_id.to_le_bytes());
        bytes[2..4].copy_from_slice(&self.first_button.to_le_bytes());
        bytes
    }
}

impl ObjCheckBoxData {
    fn parse(bytes: &[u8]) -> Self {
        Self {
            checked: u16::from_le_bytes([bytes[0], bytes[1]]),
            accelerator: u16::from_le_bytes([bytes[2], bytes[3]]),
            reserved: u16::from_le_bytes([bytes[4], bytes[5]]),
            flags: u16::from_le_bytes([bytes[6], bytes[7]]),
            trailing: bytes[8..].to_vec(),
        }
    }

    fn to_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + self.trailing.len());
        bytes.extend_from_slice(&self.checked.to_le_bytes());
        bytes.extend_from_slice(&self.accelerator.to_le_bytes());
        bytes.extend_from_slice(&self.reserved.to_le_bytes());
        bytes.extend_from_slice(&self.flags.to_le_bytes());
        bytes.extend_from_slice(&self.trailing);
        bytes
    }
}

impl ObjListBoxData {
    fn parse(
        continued_size: u16,
        bytes: &[u8],
        object_type: Option<u16>,
        limits: Limits,
    ) -> Result<Self> {
        if continued_size == 0 {
            if !bytes.is_empty() {
                return Err(Error::invalid(0, "empty FtLbsData has payload bytes"));
            }
            return Ok(Self {
                continued_size,
                formula: ObjFormula::parse(0, &[])?,
                line_count: 0,
                selected_index: 0,
                flags: ObjListBoxFlags::empty(),
                edit_object_id: 0,
                drop_data: None,
                lines: Vec::new(),
                selections: Vec::new(),
                trailing: Vec::new(),
            });
        }
        let mut reader = Reader::new(Cursor::new(bytes))?;
        let formula_length = reader.read_u16()?;
        let formula_bytes = reader.read_vec(usize::from(formula_length))?;
        let formula = ObjFormula::parse(formula_length, &formula_bytes)?;
        let line_count = reader.read_u16()?;
        if usize::from(line_count) > limits.max_entries {
            return Err(Error::Limit("FtLbsData line count exceeds limit".into()));
        }
        let selected_index = reader.read_u16()?;
        let flags = ObjListBoxFlags::from_bits_retain(reader.read_u16()?);
        let edit_object_id = reader.read_u16()?;
        let drop_data = if object_type == Some(0x0014) {
            Some(ObjListDropData::read(&mut reader)?)
        } else {
            None
        };
        let mut lines = Vec::new();
        if flags.contains(ObjListBoxFlags::VALID_STRING_ARRAY) {
            lines.reserve(usize::from(line_count));
            for _ in 0..line_count {
                lines.push(ObjListString::read(&mut reader)?);
            }
        }
        let selections = if flags
            .intersects(ObjListBoxFlags::SELECTION_TYPE_0 | ObjListBoxFlags::SELECTION_TYPE_1)
        {
            reader.read_vec(usize::from(line_count))?
        } else {
            Vec::new()
        };
        let trailing_length = usize::try_from(reader.remaining()?)
            .map_err(|_| Error::Limit("FtLbsData trailing length exceeds usize".into()))?;
        let trailing = reader.read_vec(trailing_length)?;
        Ok(Self {
            continued_size,
            formula,
            line_count,
            selected_index,
            flags,
            edit_object_id,
            drop_data,
            lines,
            selections,
            trailing,
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.continued_size == 0 {
            if self.formula.declared_length != 0
                || !matches!(self.formula.data, ObjFormulaData::Empty)
                || self.line_count != 0
                || self.selected_index != 0
                || !self.flags.is_empty()
                || self.edit_object_id != 0
                || self.drop_data.is_some()
                || !self.lines.is_empty()
                || !self.selections.is_empty()
                || !self.trailing.is_empty()
            {
                return Err(Error::invalid(0, "empty FtLbsData has populated fields"));
            }
            return Ok(Vec::new());
        }
        let formula = self.formula.to_bytes()?;
        if formula.len() != usize::from(self.formula.declared_length) {
            return Err(Error::invalid(0, "FtLbsData ObjFmla length mismatch"));
        }
        let expected_lines = if self.flags.contains(ObjListBoxFlags::VALID_STRING_ARRAY) {
            usize::from(self.line_count)
        } else {
            0
        };
        if self.lines.len() != expected_lines {
            return Err(Error::invalid(0, "FtLbsData string array count mismatch"));
        }
        let has_selections = self
            .flags
            .intersects(ObjListBoxFlags::SELECTION_TYPE_0 | ObjListBoxFlags::SELECTION_TYPE_1);
        let expected_selections = if has_selections {
            usize::from(self.line_count)
        } else {
            0
        };
        if self.selections.len() != expected_selections {
            return Err(Error::invalid(
                0,
                "FtLbsData selection array count mismatch",
            ));
        }
        let mut bytes = self.formula.declared_length.to_le_bytes().to_vec();
        bytes.extend_from_slice(&formula);
        bytes.extend_from_slice(&self.line_count.to_le_bytes());
        bytes.extend_from_slice(&self.selected_index.to_le_bytes());
        bytes.extend_from_slice(&self.flags.bits().to_le_bytes());
        bytes.extend_from_slice(&self.edit_object_id.to_le_bytes());
        if let Some(drop_data) = &self.drop_data {
            bytes.extend_from_slice(&drop_data.to_bytes()?);
        }
        for line in &self.lines {
            bytes.extend_from_slice(&line.to_bytes()?);
        }
        bytes.extend_from_slice(&self.selections);
        bytes.extend_from_slice(&self.trailing);
        Ok(bytes)
    }
}

impl ObjListDropData {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let style = reader.read_u16()?;
        let line_count = reader.read_u16()?;
        let minimum_width = reader.read_u16()?;
        let declared_character_count = reader.read_u16()?;
        let text = BiffUnicodeString::read(reader, usize::from(declared_character_count))?;
        let encoded_size = 3usize
            .checked_add(if text.flags & 1 == 0 {
                usize::from(declared_character_count)
            } else {
                usize::from(declared_character_count) * 2
            })
            .ok_or_else(|| Error::Limit("LbsDropData string size overflow".into()))?;
        let unused = if encoded_size % 2 != 0 {
            Some(reader.read_u8()?)
        } else {
            None
        };
        Ok(Self {
            style,
            line_count,
            minimum_width,
            declared_character_count,
            text,
            unused,
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.text.character_count() != usize::from(self.declared_character_count) {
            return Err(Error::invalid(0, "LbsDropData string length mismatch"));
        }
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&self.style.to_le_bytes());
        bytes.extend_from_slice(&self.line_count.to_le_bytes());
        bytes.extend_from_slice(&self.minimum_width.to_le_bytes());
        bytes.extend_from_slice(&self.declared_character_count.to_le_bytes());
        bytes.extend_from_slice(&self.text.to_bytes()?);
        if let Some(unused) = self.unused {
            bytes.push(unused);
        }
        Ok(bytes)
    }
}

impl ObjListString {
    fn read<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self> {
        let declared_character_count = reader.read_u16()?;
        let text = BiffUnicodeString::read(reader, usize::from(declared_character_count))?;
        Ok(Self {
            declared_character_count,
            text,
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.text.character_count() != usize::from(self.declared_character_count) {
            return Err(Error::invalid(0, "FtLbsData list string length mismatch"));
        }
        let mut bytes = self.declared_character_count.to_le_bytes().to_vec();
        bytes.extend_from_slice(&self.text.to_bytes()?);
        Ok(bytes)
    }
}

impl MsoDrawingHostData {
    fn encode_physical(&self) -> Result<Vec<EncodedBiffRecord>> {
        match self {
            Self::Obj(value) => Ok(vec![EncodedBiffRecord {
                record_type: OBJ,
                payload: value.to_bytes()?,
            }]),
            Self::ObjCompatibility { record_type, value } => Ok(vec![EncodedBiffRecord {
                record_type: *record_type,
                payload: value.to_bytes()?,
            }]),
            Self::Txo(value) => value.encode_physical(),
            Self::Note(value) => Ok(vec![EncodedBiffRecord {
                record_type: NOTE,
                payload: value.to_bytes()?,
            }]),
            Self::Raw {
                record_type,
                payload,
            } => Ok(vec![EncodedBiffRecord {
                record_type: *record_type,
                payload: payload.clone(),
            }]),
        }
    }
}

fn continue_payload<'a>(
    records: &'a [BiffRecord],
    index: usize,
    context: &str,
) -> Result<&'a [u8]> {
    match records.get(index).map(|record| &record.data) {
        Some(BiffRecordData::Continue { payload }) => Ok(payload),
        _ => Err(Error::invalid(0, format!("missing {context} Continue"))),
    }
}

fn stitch_continued_records(records: &mut Vec<BiffRecord>, limits: Limits) -> Result<()> {
    let mut index = 0usize;
    while index < records.len() {
        let sort_condition_count = match &records[index].data {
            BiffRecordData::SortData(value) if value.conditions.is_empty() => {
                usize::try_from(value.condition_count)
                    .map_err(|_| Error::Limit("SortData condition count exceeds usize".into()))?
            }
            _ => 0,
        };
        if sort_condition_count != 0 {
            let mut conditions = Vec::with_capacity(sort_condition_count);
            for condition_index in 0..sort_condition_count {
                let continuation = match records.get(index + 1 + condition_index) {
                    Some(BiffRecord {
                        data: BiffRecordData::ContinueFrt12(value),
                        ..
                    }) => value,
                    _ => {
                        return Err(Error::invalid(
                            0,
                            "SortData lacks a ContinueFrt12 condition",
                        ));
                    }
                };
                conditions.push(SortConditionContinuation {
                    header: continuation.header,
                    condition: parse_sdk(&continuation.continuation, 0, CONTINUE_FRT12)?,
                });
            }
            let BiffRecordData::SortData(value) = &mut records[index].data else {
                unreachable!("SortData condition count came from this record")
            };
            value.conditions = conditions;
            records.drain(index + 1..index + 1 + sort_condition_count);
        }
        let unknown = match &records[index].data {
            BiffRecordData::Unknown {
                record_type,
                payload,
            } => Some((*record_type, payload.clone())),
            _ => None,
        };
        if let Some((record_type, first_payload)) = unknown {
            if matches!(record_type, BK_HIM | IM_DATA) {
                let (value, consumed) =
                    BkHimRecord::from_sequence(&first_payload, &records[index + 1..], limits)?;
                records[index].data = if record_type == BK_HIM {
                    BiffRecordData::BkHim(value)
                } else {
                    BiffRecordData::ImData(value)
                };
                if consumed != 0 {
                    records.drain(index + 1..index + 1 + consumed);
                }
            } else if record_type == STRING_VALUE {
                let (value, consumed) =
                    StringValueRecord::from_sequence(&first_payload, &records[index + 1..])?;
                records[index].data = BiffRecordData::StringValue(value);
                if consumed != 0 {
                    records.drain(index + 1..index + 1 + consumed);
                }
            } else if record_type == MSO_DRAWING_GROUP {
                let following_count = records[index + 1..]
                    .iter()
                    .take_while(|record| {
                        matches!(record.data, BiffRecordData::Continue { .. })
                            || matches!(
                                record.data,
                                BiffRecordData::Unknown {
                                    record_type: MSO_DRAWING_GROUP,
                                    ..
                                }
                            )
                    })
                    .count();
                let mut segments = Vec::with_capacity(following_count + 1);
                segments.push((MSO_DRAWING_GROUP, first_payload.as_slice()));
                for record in &records[index + 1..index + 1 + following_count] {
                    match &record.data {
                        BiffRecordData::Continue { payload } => {
                            segments.push((CONTINUE, payload.as_slice()));
                        }
                        BiffRecordData::Unknown {
                            record_type: MSO_DRAWING_GROUP,
                            payload,
                        } => segments.push((MSO_DRAWING_GROUP, payload.as_slice())),
                        _ => unreachable!("following_count only includes drawing segments"),
                    }
                }
                records[index].data = BiffRecordData::MsoDrawingGroup(
                    MsoDrawingRecord::from_segments(&segments, limits)?,
                );
                if following_count != 0 {
                    records.drain(index + 1..index + 1 + following_count);
                }
            } else if matches!(record_type, MSO_DRAWING | MSO_DRAWING_AC_COMPATIBILITY) {
                let mut cursor = index;
                let mut owned_segments = Vec::<(u16, Vec<u8>)>::new();
                let mut host_records = Vec::new();
                let mut preceding_object_type = None;
                while cursor < records.len() {
                    match &records[cursor].data {
                        BiffRecordData::Unknown {
                            record_type: segment_type @ (MSO_DRAWING | MSO_DRAWING_AC_COMPATIBILITY),
                            payload,
                        } => {
                            owned_segments.push((*segment_type, payload.clone()));
                            cursor += 1;
                            while let Some(BiffRecord {
                                data: BiffRecordData::Continue { payload },
                                ..
                            }) = records.get(cursor)
                            {
                                owned_segments.push((CONTINUE, payload.clone()));
                                cursor += 1;
                            }
                        }
                        BiffRecordData::Unknown {
                            record_type: obj_record_type @ (OBJ | OBJ_DC5D_COMPATIBILITY),
                            payload,
                        } if !owned_segments.is_empty() => {
                            let after_segment = owned_segments.len();
                            let parsed = ObjRecord::from_bytes(payload, limits);
                            preceding_object_type =
                                parsed.as_ref().ok().and_then(ObjRecord::object_type);
                            host_records.push(MsoDrawingHostRecord {
                                after_segment,
                                data: parsed.map_or_else(
                                    |_| MsoDrawingHostData::Raw {
                                        record_type: *obj_record_type,
                                        payload: payload.clone(),
                                    },
                                    |value| {
                                        if *obj_record_type == OBJ {
                                            MsoDrawingHostData::Obj(value)
                                        } else {
                                            MsoDrawingHostData::ObjCompatibility {
                                                record_type: *obj_record_type,
                                                value,
                                            }
                                        }
                                    },
                                ),
                            });
                            cursor += 1;
                        }
                        BiffRecordData::Unknown {
                            record_type: TXO,
                            payload,
                        } if !owned_segments.is_empty() => {
                            let after_segment = owned_segments.len();
                            match TxoRecord::from_sequence(
                                payload,
                                &records[cursor + 1..],
                                preceding_object_type,
                                limits,
                            ) {
                                Ok((txo, consumed)) => {
                                    host_records.push(MsoDrawingHostRecord {
                                        after_segment,
                                        data: MsoDrawingHostData::Txo(txo),
                                    });
                                    cursor += 1 + consumed;
                                }
                                Err(_) => {
                                    host_records.push(MsoDrawingHostRecord {
                                        after_segment,
                                        data: MsoDrawingHostData::Raw {
                                            record_type: TXO,
                                            payload: payload.clone(),
                                        },
                                    });
                                    cursor += 1;
                                }
                            }
                        }
                        BiffRecordData::Continue { payload } if !owned_segments.is_empty() => {
                            owned_segments.push((CONTINUE, payload.clone()));
                            cursor += 1;
                        }
                        BiffRecordData::Unknown {
                            record_type: NOTE,
                            payload,
                        } if !owned_segments.is_empty() => {
                            host_records.push(MsoDrawingHostRecord {
                                after_segment: owned_segments.len(),
                                data: NoteRecord::from_bytes(payload, limits).map_or_else(
                                    |_| MsoDrawingHostData::Raw {
                                        record_type: NOTE,
                                        payload: payload.clone(),
                                    },
                                    MsoDrawingHostData::Note,
                                ),
                            });
                            cursor += 1;
                        }
                        _ => break,
                    }
                }
                let segments = owned_segments
                    .iter()
                    .map(|(record_type, payload)| (*record_type, payload.as_slice()))
                    .collect::<Vec<_>>();
                let following_record_type = records
                    .get(cursor)
                    .and_then(|record| record.data.encode().ok())
                    .map(|(record_type, _)| record_type);
                records[index].data =
                    BiffRecordData::MsoDrawing(MsoDrawingRecord::from_interleaved(
                        &segments,
                        host_records,
                        following_record_type,
                        limits,
                    )?);
                if cursor > index + 1 {
                    records.drain(index + 1..cursor);
                }
            } else if matches!(record_type, SST | NAME | PLS) {
                let continue_count = records[index + 1..]
                    .iter()
                    .take_while(|record| matches!(record.data, BiffRecordData::Continue { .. }))
                    .count();
                let continues = records[index + 1..index + 1 + continue_count]
                    .iter()
                    .map(|record| match &record.data {
                        BiffRecordData::Continue { payload } => payload.as_slice(),
                        _ => unreachable!("continue_count only includes Continue records"),
                    })
                    .collect::<Vec<_>>();
                records[index].data = match record_type {
                    SST => BiffRecordData::Sst(SstRecord::from_sequence(
                        &first_payload,
                        &continues,
                        limits,
                    )?),
                    NAME => BiffRecordData::Name(NameRecord::from_sequence(
                        &first_payload,
                        &continues,
                        limits,
                    )?),
                    PLS => BiffRecordData::Pls(PlsRecord::from_sequence(
                        &first_payload,
                        &continues,
                        limits,
                    )?),
                    _ => unreachable!("matched stitchable logical record"),
                };
                if continue_count != 0 {
                    records.drain(index + 1..index + 1 + continue_count);
                }
            } else if matches!(record_type, OBJ | OBJ_DC5D_COMPATIBILITY) {
                let value = ObjRecord::from_bytes(&first_payload, limits)?;
                records[index].data = if record_type == OBJ {
                    BiffRecordData::Obj(value)
                } else {
                    BiffRecordData::ObjCompatibility { record_type, value }
                };
            } else if record_type == TXO {
                let (value, consumed) =
                    TxoRecord::from_sequence(&first_payload, &records[index + 1..], None, limits)?;
                records[index].data = BiffRecordData::Txo(value);
                if consumed != 0 {
                    records.drain(index + 1..index + 1 + consumed);
                }
            }
        }
        index += 1;
    }
    Ok(())
}

impl BoundSheet8Record {
    fn from_bytes(bytes: &[u8], offset: usize) -> Result<Self> {
        let mut cursor = 0usize;
        let sheet_bof_offset = take_u32(bytes, &mut cursor, "truncated BoundSheet8 offset")?;
        let state = take_u8(bytes, &mut cursor, "truncated BoundSheet8 state")?;
        let sheet_type = take_u8(bytes, &mut cursor, "truncated BoundSheet8 type")?;
        let name = ShortXlUnicodeString::read(bytes, &mut cursor)?;
        if cursor != bytes.len() {
            return Err(Error::invalid(
                offset as u64 + cursor as u64,
                "unexpected trailing bytes in BoundSheet8",
            ));
        }
        Ok(Self {
            sheet_bof_offset,
            state,
            sheet_type,
            name,
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&self.sheet_bof_offset.to_le_bytes());
        bytes.push(self.state);
        bytes.push(self.sheet_type);
        self.name.write(&mut bytes)?;
        Ok(bytes)
    }
}

impl ShortXlUnicodeString {
    fn read(bytes: &[u8], cursor: &mut usize) -> Result<Self> {
        let count = usize::from(take_u8(bytes, cursor, "truncated sheet-name length")?);
        let flags = take_u8(bytes, cursor, "truncated sheet-name flags")?;
        let characters = if flags & 1 == 0 {
            XlStringCharacters::Compressed(
                take_bytes(bytes, cursor, count, "truncated compressed sheet name")?.to_vec(),
            )
        } else {
            let byte_count = count
                .checked_mul(2)
                .ok_or_else(|| Error::Limit("sheet-name byte count overflow".into()))?;
            let raw = take_bytes(bytes, cursor, byte_count, "truncated Unicode sheet name")?;
            XlStringCharacters::Unicode(
                raw.chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect(),
            )
        };
        Ok(Self { flags, characters })
    }

    fn write(&self, bytes: &mut Vec<u8>) -> Result<()> {
        let count = match &self.characters {
            XlStringCharacters::Compressed(values) => {
                if self.flags & 1 != 0 {
                    return Err(Error::invalid(0, "compressed sheet name has UTF-16 flag"));
                }
                values.len()
            }
            XlStringCharacters::Unicode(values) => {
                if self.flags & 1 == 0 {
                    return Err(Error::invalid(0, "Unicode sheet name lacks UTF-16 flag"));
                }
                values.len()
            }
        };
        bytes.push(u8::try_from(count).map_err(|_| Error::Limit("sheet name exceeds u8".into()))?);
        bytes.push(self.flags);
        match &self.characters {
            XlStringCharacters::Compressed(values) => bytes.extend_from_slice(values),
            XlStringCharacters::Unicode(values) => {
                for value in values {
                    bytes.extend_from_slice(&value.to_le_bytes());
                }
            }
        }
        Ok(())
    }
}

fn decode_record(
    record_type: u16,
    payload: &[u8],
    encrypted: bool,
    biff8: bool,
    offset: usize,
) -> Result<BiffRecordData> {
    if encrypted && !matches!(record_type, BOF | EOF) {
        return Ok(BiffRecordData::Encrypted {
            record_type,
            payload: payload.to_vec(),
        });
    }
    if biff8 {
        let typed =
            match record_type {
                DIMENSIONS => Some(BiffRecordData::Dimensions(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FORMULA => Some(BiffRecordData::Formula(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FORMULA4 => Some(BiffRecordData::Formula4Compatibility(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SHARED_FORMULA => Some(BiffRecordData::SharedFormula(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SUP_BOOK => Some(BiffRecordData::SupBook(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CF => Some(BiffRecordData::ConditionalFormatting(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                COND_FMT => Some(BiffRecordData::ConditionalFormattingGroup(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                EXTERN_SHEET => Some(BiffRecordData::ExternSheet(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                EXTERN_NAME => Some(BiffRecordData::ExternName(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                HLINK => Some(BiffRecordData::Hyperlink(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DV => Some(BiffRecordData::DataValidation(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                HEADER => Some(BiffRecordData::Header(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FOOTER => Some(BiffRecordData::Footer(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                VERTICAL_PAGE_BREAKS => Some(BiffRecordData::VerticalPageBreaks(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                HORIZONTAL_PAGE_BREAKS => Some(BiffRecordData::HorizontalPageBreaks(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DCON => Some(BiffRecordData::DCon(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DCON_REF => Some(BiffRecordData::DConRef(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DCONN => Some(BiffRecordData::DConn(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                TXT_QRY => Some(BiffRecordData::TextQuery(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                QSI_SX_TAG => Some(BiffRecordData::QsiSxTag(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_VIEW_EX9 => Some(BiffRecordData::SxViewEx9(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DB_QUERY_EXT => Some(BiffRecordData::DbQueryExt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                HLINK_TOOLTIP => Some(BiffRecordData::HyperlinkTooltip(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CONTINUE_FRT12 => Some(BiffRecordData::ContinueFrt12(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_ADDL => Some(BiffRecordData::SxAddl(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                ENT_EX_U2 => Some(BiffRecordData::EntExU2(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                REAL_TIME_DATA => Some(BiffRecordData::RealTimeData(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SORT => Some(BiffRecordData::Sort(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                LH_RECORD => Some(BiffRecordData::LhRecord(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SORT_DATA => Some(BiffRecordData::SortData(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                AUTO_FILTER => Some(BiffRecordData::AutoFilter(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_FORMAT => Some(BiffRecordData::SxFormat(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                W_OPT => Some(BiffRecordData::WOpt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                TABLE => Some(BiffRecordData::Table(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                EXTERN_COUNT => Some(BiffRecordData::ExternCount(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                QSI => Some(BiffRecordData::Qsi(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                PARAM_QRY => Some(BiffRecordData::ParamQry(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_SELECT => Some(BiffRecordData::SxSelect(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FILE_SHARING => Some(BiffRecordData::FileSharing(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                OLE_OBJECT_SIZE => Some(BiffRecordData::OleObjectSize(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                MSO_DRAWING_SELECTION => Some(BiffRecordData::MsoDrawingSelection(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SCEN_MAN => Some(BiffRecordData::ScenMan(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_VIEW => Some(BiffRecordData::SxView(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                BLANK => Some(BiffRecordData::Blank(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                NUMBER => Some(BiffRecordData::Number(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                BOOL_ERR => Some(BiffRecordData::BoolErr(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                LABEL => Some(BiffRecordData::Label(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SXVI => Some(BiffRecordData::Sxvi(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_IVD => Some(BiffRecordData::SxIvd(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_PI => Some(BiffRecordData::SxPi(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_DI => Some(BiffRecordData::SxDi(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_STRING => Some(BiffRecordData::SxString(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                RR_TAB_ID => Some(BiffRecordData::RrTabId(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_RULE => Some(BiffRecordData::SxRule(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_EX => Some(BiffRecordData::SxEx(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_FILT => Some(BiffRecordData::SxFilt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_DXF => Some(BiffRecordData::SxDxf(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_ITM => Some(BiffRecordData::SxItm(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_STREAM_ID => Some(BiffRecordData::SxStreamId(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SX_VS => Some(BiffRecordData::SxVs(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                RECALC_ID => Some(BiffRecordData::RecalcId(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SXVD_EX => Some(BiffRecordData::SxvdEx(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SXVD => Some(BiffRecordData::Sxvd(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CODE_NAME => Some(BiffRecordData::CodeName(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                ARRAY => Some(BiffRecordData::Array(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                USER_SVIEW_BEGIN => Some(BiffRecordData::UserSViewBegin(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                USER_SVIEW_END => Some(BiffRecordData::UserSViewEnd(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                USER_BVIEW => Some(BiffRecordData::UserBView(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SHEET_EXT => Some(BiffRecordData::SheetExt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_DATA_LABEL_EXT_CONTENTS => Some(BiffRecordData::ChartDataLabelExtContents(
                    parse_sdk(payload, offset, record_type)?,
                )),
                CELL_WATCH => Some(BiffRecordData::CellWatch(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FEAT_HDR11 => Some(BiffRecordData::FeatureHeader11(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FEATURE11 => Some(BiffRecordData::Feature11(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                LIST12 => Some(BiffRecordData::List12(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DROP_DOWN_OBJ_IDS => Some(BiffRecordData::DropDownObjIds(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DATA_VALIDATION_HEADER => Some(BiffRecordData::DataValidationHeader(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                RICH_TEXT_STREAM => Some(BiffRecordData::RichTextStream(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                GUID_TYPE_LIB => Some(BiffRecordData::GuidTypeLib(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                NAME_COMMENT => Some(BiffRecordData::NameComment(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                LABEL_SST => Some(BiffRecordData::LabelSst(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                EXT_SST => Some(BiffRecordData::ExtSst(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                RK => Some(BiffRecordData::Rk(parse_sdk(payload, offset, record_type)?)),
                ROW => Some(BiffRecordData::Row(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                WINDOW1 => Some(BiffRecordData::Window1(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                PANE => Some(BiffRecordData::Pane(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                COL_INFO => Some(BiffRecordData::ColInfo(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                GUTS => Some(BiffRecordData::Guts(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                COUNTRY => Some(BiffRecordData::Country(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                PALETTE => Some(BiffRecordData::Palette(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SCL => Some(BiffRecordData::Scl(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                PRINT_SETUP => Some(BiffRecordData::PrintSetup(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                MUL_RK => Some(BiffRecordData::MulRk(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                MUL_BLANK => Some(BiffRecordData::MulBlank(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                XF => Some(BiffRecordData::Xf(parse_sdk(payload, offset, record_type)?)),
                record_type @ (XF_E4_COMPATIBILITY
                | XF_EE_COMPATIBILITY
                | XF_E8E0_COMPATIBILITY) => Some(BiffRecordData::XfCompatibility {
                    record_type,
                    value: parse_sdk(payload, offset, record_type)?,
                }),
                CRN => Some(BiffRecordData::Crn(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                XF_EXT => Some(BiffRecordData::XfExt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                XF_CRC => Some(BiffRecordData::XfCrc(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                TABLE_STYLES => Some(BiffRecordData::TableStyles(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                STYLE_EXT => Some(BiffRecordData::StyleExt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DXF => Some(BiffRecordData::Dxf(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CF_EX => Some(BiffRecordData::ConditionalFormattingExtension(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CF12 => Some(BiffRecordData::ConditionalFormatting12(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                COND_FMT12 => Some(BiffRecordData::ConditionalFormattingGroup12(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                THEME => Some(BiffRecordData::Theme(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                HEADER_FOOTER_EXT => Some(BiffRecordData::ExtendedHeaderFooter(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SHAPE_PROPS_STREAM => Some(BiffRecordData::ShapePropsStream(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                TEXT_PROPS_STREAM => Some(BiffRecordData::TextPropsStream(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                COMPAT12 => Some(BiffRecordData::Compat12(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                PLV => Some(BiffRecordData::Plv(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                PLV_MAC => Some(BiffRecordData::PlvMac(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                LNEXT => Some(BiffRecordData::Lnext(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                MKR_EXT => Some(BiffRecordData::MkrExt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CRT_CO_OPT => Some(BiffRecordData::CrtCoOpt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FRT_ARCH_ID => Some(BiffRecordData::FrtArchId(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CRT_LAYOUT12 => Some(BiffRecordData::CrtLayout12(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CRT_LAYOUT12_A => Some(BiffRecordData::CrtLayout12A(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                MTR_SETTINGS => Some(BiffRecordData::MtrSettings(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FORCE_FULL_CALCULATION => Some(BiffRecordData::ForceFullCalculation(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                COMPRESS_PICTURES => Some(BiffRecordData::CompressPictures(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CRT_ML_FRT => Some(BiffRecordData::CrtMlFrt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_FRT_INFO => Some(BiffRecordData::ChartFrtInfo(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_CAT_LAB => Some(BiffRecordData::ChartCatLab(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_START_OBJECT => Some(BiffRecordData::ChartStartObject(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_END_OBJECT => Some(BiffRecordData::ChartEndObject(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                GEL_FRAME => Some(BiffRecordData::GelFrame(OfficeArtStream::from_bytes(
                    payload,
                )?)),
                HF_PICTURE => Some(BiffRecordData::HfPicture(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FEAT_HDR => Some(BiffRecordData::FeatureHeader(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FEAT => Some(BiffRecordData::Feature(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                BOOK_EXT => Some(BiffRecordData::BookExt(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART => Some(BiffRecordData::Chart(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_AREA_FORMAT => Some(BiffRecordData::ChartAreaFormat(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_ATTACHED_LABEL => Some(BiffRecordData::ChartAttachedLabel(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_DATA_FORMAT => Some(BiffRecordData::ChartDataFormat(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_FORMAT => Some(BiffRecordData::ChartFormat(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SERIES_LIST => Some(BiffRecordData::ChartSeriesList(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_BAR => Some(BiffRecordData::ChartBar(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_LINE => Some(BiffRecordData::ChartLine(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_PIE => Some(BiffRecordData::ChartPie(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_AREA => Some(BiffRecordData::ChartArea(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SCATTER => Some(BiffRecordData::ChartScatter(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_CRT_LINE => Some(BiffRecordData::ChartCrtLine(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_CRT_LINK => Some(BiffRecordData::ChartCrtLink(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_LEGEND => Some(BiffRecordData::ChartLegend(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_AXIS => Some(BiffRecordData::ChartAxis(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_TICK => Some(BiffRecordData::ChartTick(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_VALUE_RANGE => Some(BiffRecordData::ChartValueRange(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_LABEL_RANGE => Some(BiffRecordData::ChartLabelRange(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_AXIS_LINE => Some(BiffRecordData::ChartAxisLine(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_DEFAULT_TEXT => Some(BiffRecordData::ChartDefaultText(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_FONT => Some(BiffRecordData::ChartFont(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_LINE_FORMAT => Some(BiffRecordData::ChartLineFormat(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_MARKER_FORMAT => Some(BiffRecordData::ChartMarkerFormat(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_OBJECT_LINK => Some(BiffRecordData::ChartObjectLink(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_FRAME => Some(BiffRecordData::ChartFrame(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_3D => Some(BiffRecordData::Chart3D(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_DROP_BAR => Some(BiffRecordData::ChartDropBar(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SURF => Some(BiffRecordData::ChartSurf(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_LEGEND_EXCEPTION => Some(BiffRecordData::ChartLegendException(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_AXIS_PARENT => Some(BiffRecordData::ChartAxisParent(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SHEET_PROPERTIES => Some(BiffRecordData::ChartSheetProperties(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SERIES_GROUP_INDEX => Some(BiffRecordData::ChartSeriesGroupIndex(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_AXIS_USED => Some(BiffRecordData::ChartAxisUsed(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_NUMBER_FORMAT_INDEX => Some(BiffRecordData::ChartNumberFormatIndex(
                    parse_sdk(payload, offset, record_type)?,
                )),
                CHART_SERIES_PARENT => Some(BiffRecordData::ChartSeriesParent(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SERIES_AUX_TREND => Some(BiffRecordData::ChartSeriesAuxTrend(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_POSITION => Some(BiffRecordData::ChartPosition(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_FONT_BASIS => Some(BiffRecordData::ChartFontBasis(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_3D_BAR_SHAPE => Some(BiffRecordData::Chart3DBarShape(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SERIES_FORMAT => Some(BiffRecordData::ChartSeriesFormat(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SERIES_AUX_ERROR_BAR => Some(BiffRecordData::ChartSeriesAuxErrorBar(
                    parse_sdk(payload, offset, record_type)?,
                )),
                CHART_CLRT_CLIENT => Some(BiffRecordData::ChartClrtClient(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_AXIS_OPTIONS => Some(BiffRecordData::ChartAxisOptions(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_DAT => Some(BiffRecordData::ChartDat(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_PIE_FORMAT => Some(BiffRecordData::ChartPieFormat(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_PLOT_GROWTH => Some(BiffRecordData::ChartPlotGrowth(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_LINKED_DATA => Some(BiffRecordData::ChartLinkedData(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_AL_RUNS => Some(BiffRecordData::ChartAlRuns(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SERIES_INDEX => Some(BiffRecordData::ChartSeriesIndex(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SERIES => Some(BiffRecordData::ChartSeries(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_SERIES_1103_COMPATIBILITY => Some(BiffRecordData::ChartSeriesCompatibility {
                    record_type,
                    value: parse_sdk(payload, offset, record_type)?,
                }),
                CHART_SERIES_TEXT => Some(BiffRecordData::ChartSeriesText(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CHART_TEXT => Some(BiffRecordData::ChartText(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                START_BLOCK => Some(BiffRecordData::StartBlock(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                END_BLOCK => Some(BiffRecordData::EndBlock(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DB_CELL => Some(BiffRecordData::DbCell(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FONT => Some(BiffRecordData::Font(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                FONT30_COMPATIBILITY => Some(BiffRecordData::FontCompatibility {
                    record_type,
                    value: parse_sdk(payload, offset, record_type)?,
                }),
                FORMAT => Some(BiffRecordData::Format(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                STYLE => Some(BiffRecordData::Style(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                CRN_COUNT => Some(BiffRecordData::CrnCount(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                DEFAULT_ROW_HEIGHT => Some(BiffRecordData::DefaultRowHeight(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                WRITE_ACCESS => Some(BiffRecordData::WriteAccess(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                WINDOW2 => Some(BiffRecordData::Window2(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                SELECTION => Some(BiffRecordData::Selection(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                MERGE_CELLS => Some(BiffRecordData::MergeCells(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                MMS => Some(BiffRecordData::Mms(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                PHONETIC_INFO => Some(BiffRecordData::PhoneticInfo(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                INDEX => Some(BiffRecordData::Index(parse_sdk(
                    payload,
                    offset,
                    record_type,
                )?)),
                _ => None,
            };
        if let Some(typed) = typed {
            return Ok(typed);
        }
        if let Some(kind) = FixedU16RecordKind::from_id(record_type) {
            if payload.len() != 2 {
                return Err(Error::invalid(
                    offset as u64,
                    format!("fixed u16 BIFF record 0x{record_type:04x} must contain 2 bytes"),
                ));
            }
            return Ok(BiffRecordData::FixedU16 {
                kind,
                value: u16::from_le_bytes([payload[0], payload[1]]),
            });
        }
        if let Some(kind) = FixedF64RecordKind::from_id(record_type) {
            if payload.len() != 8 {
                return Err(Error::invalid(
                    offset as u64,
                    format!("fixed f64 BIFF record 0x{record_type:04x} must contain 8 bytes"),
                ));
            }
            return Ok(BiffRecordData::FixedF64Bits {
                kind,
                bits: u64::from_le_bytes(payload.try_into().expect("eight bytes")),
            });
        }
        if let Some(kind) = EmptyRecordKind::from_id(record_type) {
            let reserved = match (kind, payload) {
                (_, []) => None,
                (EmptyRecordKind::WriteProtect | EmptyRecordKind::InterfaceEnd, [low, high]) => {
                    Some(u16::from_le_bytes([*low, *high]))
                }
                _ => {
                    return Err(Error::invalid(
                        offset as u64,
                        format!(
                            "empty BIFF marker record 0x{record_type:04x} has an invalid payload"
                        ),
                    ));
                }
            };
            return Ok(BiffRecordData::Empty { kind, reserved });
        }
    }
    Ok(match record_type {
        BOF if payload.len() == 16 => BiffRecordData::Bof(BofRecord {
            version: u16::from_le_bytes([payload[0], payload[1]]),
            document_type: u16::from_le_bytes([payload[2], payload[3]]),
            build_identifier: u16::from_le_bytes([payload[4], payload[5]]),
            build_year: u16::from_le_bytes([payload[6], payload[7]]),
            history_flags: u32::from_le_bytes(payload[8..12].try_into().expect("four bytes")),
            lowest_version: u32::from_le_bytes(payload[12..16].try_into().expect("four bytes")),
        }),
        BOF => BiffRecordData::LegacyBof {
            payload: payload.to_vec(),
        },
        EOF if payload.is_empty() => BiffRecordData::Eof,
        EOF => {
            return Err(Error::invalid(
                offset as u64,
                "EOF record must have no data",
            ));
        }
        CODE_PAGE if payload.len() == 2 => BiffRecordData::CodePage {
            code_page: u16::from_le_bytes([payload[0], payload[1]]),
        },
        CODE_PAGE => {
            return Err(Error::invalid(
                offset as u64,
                "CodePage record must contain 2 bytes",
            ));
        }
        BOUND_SHEET8 if biff8 => {
            BiffRecordData::BoundSheet8(BoundSheet8Record::from_bytes(payload, offset + 4)?)
        }
        BOUND_SHEET8_2085_COMPATIBILITY if biff8 => BiffRecordData::BoundSheet8Compatibility {
            record_type,
            value: BoundSheet8Record::from_bytes(payload, offset + 4)?,
        },
        FILE_PASS => BiffRecordData::FilePass {
            payload: payload.to_vec(),
        },
        CONTINUE => BiffRecordData::Continue {
            payload: payload.to_vec(),
        },
        _ => BiffRecordData::Unknown {
            record_type,
            payload: payload.to_vec(),
        },
    })
}

fn parse_sdk<T: SdkRead>(payload: &[u8], offset: usize, record_type: u16) -> Result<T> {
    let mut reader = Reader::new(Cursor::new(payload))?;
    let value = T::read_from(&mut reader).map_err(|error| match error {
        Error::InvalidData {
            offset: inner_offset,
            message,
        } => Error::invalid(
            offset as u64 + 4 + inner_offset,
            format!("typed BIFF record 0x{record_type:04x}: {message}"),
        ),
        Error::Io(error) => Error::invalid(
            offset as u64 + 4,
            format!("typed BIFF record 0x{record_type:04x}: {error}"),
        ),
        Error::Limit(message) => {
            Error::Limit(format!("typed BIFF record 0x{record_type:04x}: {message}"))
        }
    })?;
    if reader.remaining()? != 0 {
        return Err(Error::invalid(
            offset as u64,
            format!("unexpected trailing bytes in typed BIFF record 0x{record_type:04x}"),
        ));
    }
    Ok(value)
}

fn parse_sx_li(payload: &[u8], offset: usize, axis_dimension_count: u16) -> Result<SxLiRecord> {
    let mut reader = Reader::new(Cursor::new(payload))?;
    let value = SxLiRecord::read_with_axis_dimension(&mut reader, axis_dimension_count).map_err(
        |error| match error {
            Error::InvalidData {
                offset: inner_offset,
                message,
            } => Error::invalid(
                offset as u64 + 4 + inner_offset,
                format!("typed BIFF record 0x{SX_LI:04x}: {message}"),
            ),
            Error::Io(error) => Error::invalid(
                offset as u64 + 4,
                format!("typed BIFF record 0x{SX_LI:04x}: {error}"),
            ),
            Error::Limit(message) => {
                Error::Limit(format!("typed BIFF record 0x{SX_LI:04x}: {message}"))
            }
        },
    )?;
    if reader.remaining()? != 0 {
        return Err(Error::invalid(
            offset as u64,
            "unexpected trailing bytes in typed SXLI record",
        ));
    }
    Ok(value)
}

fn encode_sdk<T: SdkWrite>(value: &T) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    value.write_to(&mut writer)?;
    Ok(writer.into_inner().into_inner())
}

impl FixedU16RecordKind {
    fn from_id(id: u16) -> Option<Self> {
        Some(match id {
            0x000c => Self::CalcCount,
            0x000d => Self::CalcMode,
            0x000e => Self::CalcPrecision,
            0x000f => Self::RefMode,
            0x0011 => Self::Iteration,
            0x0012 => Self::Protect,
            0x0013 => Self::Password,
            0x0019 => Self::WindowProtect,
            0x0022 => Self::Date1904,
            0x002a => Self::PrintHeaders,
            0x002b => Self::PrintGridlines,
            0x0040 => Self::Backup,
            0x0055 => Self::DefaultColWidth,
            0x005e => Self::Uncalced,
            0x005f => Self::SaveRecalc,
            0x0063 => Self::ObjectProtect,
            0x0082 => Self::Gridset,
            0x0083 => Self::HCenter,
            0x0084 => Self::VCenter,
            0x008d => Self::HideObj,
            0x009c => Self::FnGroupCount,
            0x009d => Self::AutoFilterInfo,
            0x00da => Self::BookBool,
            0x00dd => Self::ScenarioProtect,
            0x00e1 => Self::InterfaceHdr,
            0x0160 => Self::UseSelFs,
            0x0161 => Self::Dsf,
            0x01af => Self::ProtectionRev4,
            0x01b7 => Self::RefreshAll,
            0x01bc => Self::PasswordRev4,
            0x1001 => Self::ChartUnits,
            0x0081 => Self::WsBool,
            0x0033 => Self::PrintSize,
            0x0099 => Self::StandardWidth,
            _ => return None,
        })
    }

    fn id(self) -> u16 {
        match self {
            Self::CalcCount => 0x000c,
            Self::CalcMode => 0x000d,
            Self::CalcPrecision => 0x000e,
            Self::RefMode => 0x000f,
            Self::Iteration => 0x0011,
            Self::Protect => 0x0012,
            Self::Password => 0x0013,
            Self::WindowProtect => 0x0019,
            Self::Date1904 => 0x0022,
            Self::PrintHeaders => 0x002a,
            Self::PrintGridlines => 0x002b,
            Self::Backup => 0x0040,
            Self::DefaultColWidth => 0x0055,
            Self::Uncalced => 0x005e,
            Self::SaveRecalc => 0x005f,
            Self::ObjectProtect => 0x0063,
            Self::Gridset => 0x0082,
            Self::HCenter => 0x0083,
            Self::VCenter => 0x0084,
            Self::HideObj => 0x008d,
            Self::FnGroupCount => 0x009c,
            Self::AutoFilterInfo => 0x009d,
            Self::BookBool => 0x00da,
            Self::ScenarioProtect => 0x00dd,
            Self::InterfaceHdr => 0x00e1,
            Self::UseSelFs => 0x0160,
            Self::Dsf => 0x0161,
            Self::ProtectionRev4 => 0x01af,
            Self::RefreshAll => 0x01b7,
            Self::PasswordRev4 => 0x01bc,
            Self::ChartUnits => 0x1001,
            Self::WsBool => 0x0081,
            Self::PrintSize => 0x0033,
            Self::StandardWidth => 0x0099,
        }
    }
}

impl FixedF64RecordKind {
    fn from_id(id: u16) -> Option<Self> {
        Some(match id {
            0x0010 => Self::CalcDelta,
            0x0026 => Self::LeftMargin,
            0x0027 => Self::RightMargin,
            0x0028 => Self::TopMargin,
            0x0029 => Self::BottomMargin,
            _ => return None,
        })
    }

    fn id(self) -> u16 {
        match self {
            Self::CalcDelta => 0x0010,
            Self::LeftMargin => 0x0026,
            Self::RightMargin => 0x0027,
            Self::TopMargin => 0x0028,
            Self::BottomMargin => 0x0029,
        }
    }
}

impl EmptyRecordKind {
    fn from_id(id: u16) -> Option<Self> {
        Some(match id {
            0x0000 => Self::NullCompatibility,
            0x0086 => Self::WriteProtect,
            0x00e2 => Self::InterfaceEnd,
            0x0060 => Self::Template,
            0x009b => Self::FilterMode,
            0x00d3 => Self::ObjectProject,
            0x01bd => Self::VbaProjectEmpty,
            0x01c0 => Self::Excel9File,
            0x1033 => Self::ChartBegin,
            0x1034 => Self::ChartEnd,
            0x1035 => Self::ChartPlotArea,
            _ => return None,
        })
    }

    fn id(self) -> u16 {
        match self {
            Self::NullCompatibility => 0x0000,
            Self::WriteProtect => 0x0086,
            Self::InterfaceEnd => 0x00e2,
            Self::Template => 0x0060,
            Self::FilterMode => 0x009b,
            Self::ObjectProject => 0x00d3,
            Self::VbaProjectEmpty => 0x01bd,
            Self::Excel9File => 0x01c0,
            Self::ChartBegin => 0x1033,
            Self::ChartEnd => 0x1034,
            Self::ChartPlotArea => 0x1035,
        }
    }
}

fn take_u8(bytes: &[u8], cursor: &mut usize, message: &str) -> Result<u8> {
    Ok(take_bytes(bytes, cursor, 1, message)?[0])
}

fn take_u16(bytes: &[u8], cursor: &mut usize, message: &str) -> Result<u16> {
    let value = take_bytes(bytes, cursor, 2, message)?;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}

fn take_u32(bytes: &[u8], cursor: &mut usize, message: &str) -> Result<u32> {
    let value = take_bytes(bytes, cursor, 4, message)?;
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

fn take_bytes<'a>(
    bytes: &'a [u8],
    cursor: &mut usize,
    len: usize,
    message: &str,
) -> Result<&'a [u8]> {
    let end = cursor
        .checked_add(len)
        .ok_or_else(|| Error::Limit("BIFF cursor overflow".into()))?;
    let value = bytes
        .get(*cursor..end)
        .ok_or_else(|| Error::invalid(*cursor as u64, message))?;
    *cursor = end;
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn initial_biff8_records_round_trip() {
        let stream = BiffStream {
            records: vec![
                BiffRecord {
                    offset: 0,
                    data: BiffRecordData::Bof(BofRecord {
                        version: 0x0600,
                        document_type: 5,
                        build_identifier: 0x1234,
                        build_year: 0x07cc,
                        history_flags: 0x41,
                        lowest_version: 6,
                    }),
                },
                BiffRecord {
                    offset: 20,
                    data: BiffRecordData::CodePage { code_page: 1200 },
                },
                BiffRecord {
                    offset: 26,
                    data: BiffRecordData::BoundSheet8(BoundSheet8Record {
                        sheet_bof_offset: 48,
                        state: 0,
                        sheet_type: 0,
                        name: ShortXlUnicodeString {
                            flags: 0,
                            characters: XlStringCharacters::Compressed(b"Sheet1".to_vec()),
                        },
                    }),
                },
                BiffRecord {
                    offset: 44,
                    data: BiffRecordData::Eof,
                },
            ],
            trailing_padding: Vec::new(),
        };
        let bytes = stream.to_bytes().unwrap();
        let parsed = BiffStream::from_bytes(&bytes).unwrap();
        assert!(parsed.is_biff8());
        assert_eq!(parsed, stream);
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn truncated_and_oversized_records_are_rejected() {
        assert!(BiffStream::from_bytes(&[9, 8, 1]).is_err());
        assert!(BiffStream::from_bytes(&[1, 0, 0xff, 0xff]).is_err());
    }

    #[test]
    fn txo_continuations_stop_before_the_next_drawing_fragment() {
        let mut txo = vec![0; 18];
        txo[10..12].copy_from_slice(&3u16.to_le_bytes());
        txo[12..14].copy_from_slice(&16u16.to_le_bytes());
        let following = vec![
            BiffRecord {
                offset: 0,
                data: BiffRecordData::Continue {
                    payload: vec![0, b'a', b'b', b'c'],
                },
            },
            BiffRecord {
                offset: 0,
                data: BiffRecordData::Continue {
                    payload: vec![0; 16],
                },
            },
            BiffRecord {
                offset: 0,
                data: BiffRecordData::Continue {
                    payload: vec![0x0f, 0, 4, 0xf0, 0, 0, 0, 0],
                },
            },
        ];
        let (parsed, consumed) =
            TxoRecord::from_sequence(&txo, &following, None, Limits::default()).unwrap();
        assert_eq!(consumed, 2);
        assert_eq!(parsed.text_chunks.len(), 1);
        assert_eq!(parsed.runs.len(), 1);
        assert!(parsed.last_run.is_some());
    }

    #[test]
    fn trailing_sector_padding_is_preserved() {
        let bytes = [0x09, 0x08, 0, 0, 0, 0, 0];
        let parsed = BiffStream::from_bytes(&bytes).unwrap();
        assert_eq!(parsed.trailing_padding, [0, 0, 0]);
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn producer_compatibility_shapes_are_typed_and_exact() {
        let dimensions = [1, 0, 0, 0, 2, 0, 0, 0, 3, 0, 4, 0, 0, 0, 0x34, 0x12];
        let record = decode_record(DIMENSIONS, &dimensions, false, true, 0).unwrap();
        assert_eq!(record.encode().unwrap(), (DIMENSIONS, dimensions.to_vec()));

        let bool_err = [1, 0, 2, 0, 3, 0, 0x34, 0x12, 1];
        let record = decode_record(BOOL_ERR, &bool_err, false, true, 0).unwrap();
        assert_eq!(record.encode().unwrap(), (BOOL_ERR, bool_err.to_vec()));

        for col_info in [
            vec![1, 0, 2, 0, 3, 0, 4, 0, 5, 0],
            vec![1, 0, 2, 0, 3, 0, 4, 0, 5, 0, 6],
            vec![1, 0, 2, 0, 3, 0, 4, 0, 5, 0, 6, 0],
        ] {
            let record = decode_record(COL_INFO, &col_info, false, true, 0).unwrap();
            assert_eq!(record.encode().unwrap(), (COL_INFO, col_info));
        }
    }

    #[test]
    fn continued_string_preserves_physical_segments_and_encoding_changes() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&BOF.to_le_bytes());
        bytes.extend_from_slice(&16u16.to_le_bytes());
        bytes.extend_from_slice(&0x0600u16.to_le_bytes());
        bytes.extend_from_slice(&0x0005u16.to_le_bytes());
        bytes.extend_from_slice(&0x1234u16.to_le_bytes());
        bytes.extend_from_slice(&0x07ccu16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&6u32.to_le_bytes());
        bytes.extend_from_slice(&STRING_VALUE.to_le_bytes());
        bytes.extend_from_slice(&5u16.to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&[0, b'a', b'b']);
        bytes.extend_from_slice(&CONTINUE.to_le_bytes());
        bytes.extend_from_slice(&5u16.to_le_bytes());
        bytes.extend_from_slice(&[1, b'c', 0, b'd', 0]);
        bytes.extend_from_slice(&EOF.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());

        let parsed = BiffStream::from_bytes(&bytes).unwrap();
        assert_eq!(parsed.records.len(), 3);
        let BiffRecordData::StringValue(value) = &parsed.records[1].data else {
            panic!("expected typed String record");
        };
        assert_eq!(value.chunks.len(), 2);
        assert_eq!(value.declared_character_count, 4);
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn sst_preserves_rich_string_continue_layout_and_encoding_changes() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&BOF.to_le_bytes());
        bytes.extend_from_slice(&16u16.to_le_bytes());
        bytes.extend_from_slice(&0x0600u16.to_le_bytes());
        bytes.extend_from_slice(&0x0005u16.to_le_bytes());
        bytes.extend_from_slice(&0x1234u16.to_le_bytes());
        bytes.extend_from_slice(&0x07ccu16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&6u32.to_le_bytes());

        let mut first = Vec::new();
        first.extend_from_slice(&2u32.to_le_bytes());
        first.extend_from_slice(&2u32.to_le_bytes());
        first.extend_from_slice(&4u16.to_le_bytes());
        first.push(SstStringFlags::RICH_TEXT.bits());
        first.extend_from_slice(&1u16.to_le_bytes());
        first.extend_from_slice(b"ab");
        bytes.extend_from_slice(&SST.to_le_bytes());
        bytes.extend_from_slice(&(first.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&first);

        let second = [1, b'c', 0, b'd', 0, 0, 0, 2, 0];
        bytes.extend_from_slice(&CONTINUE.to_le_bytes());
        bytes.extend_from_slice(&(second.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&second);

        let third = [1, 0, 0, b'z'];
        bytes.extend_from_slice(&CONTINUE.to_le_bytes());
        bytes.extend_from_slice(&(third.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&third);
        bytes.extend_from_slice(&EOF.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());

        let parsed = BiffStream::from_bytes(&bytes).unwrap();
        assert_eq!(parsed.records.len(), 3);
        let BiffRecordData::Sst(sst) = &parsed.records[1].data else {
            panic!("expected typed SST record");
        };
        assert_eq!(sst.strings.len(), 2);
        assert_eq!(sst.strings[0].character_chunks.len(), 2);
        assert_eq!(sst.strings[0].format_runs.len(), 1);
        assert_eq!(sst.physical_segments[1].continuation_encoding, Some(1));
        assert_eq!(sst.physical_segments[2].continuation_encoding, None);
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn ext_rst_round_trips_static_phonetic_fields_and_extra_word() {
        let extension = SstExtensionData::ExtRst(ExtRst {
            reserved: 1,
            body: ExtRstBody::Phonetic {
                declared_data_size: 20,
                font_index: 3,
                formatting_flags: PhoneticFlags::from_bits_retain(0x8005),
                declared_run_count: 1,
                declared_character_count: 1,
                lpwide_character_count: 1,
                phonetic_text: vec![0x3042],
                runs: vec![PhoneticRun {
                    phonetic_text_first_character: 0,
                    source_text_first_character: 0,
                    source_text_character_count: 1,
                }],
                extra_data_word: Some(0x1234),
                inner_trailing: Vec::new(),
                outer_trailing: Vec::new(),
            },
        });
        let bytes = extension.to_bytes().unwrap();
        assert_eq!(SstExtensionData::from_bytes(bytes), extension);
        assert_eq!(extension.unparsed_byte_count(), 0);
    }

    #[test]
    fn truncated_ext_rst_phonetic_header_is_static_and_exact() {
        let bytes = vec![
            0x0c, 0x00, 0x08, 0x00, 0x37, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        let extension = SstExtensionData::from_bytes(bytes.clone());
        assert_eq!(
            extension,
            SstExtensionData::ExtRst(ExtRst {
                reserved: 0x000c,
                body: ExtRstBody::TruncatedPhoneticHeader {
                    declared_data_size: 8,
                    font_index: 0x0037,
                    formatting_flags: PhoneticFlags::empty(),
                    declared_run_count: 0,
                    declared_character_count: 0,
                },
            })
        );
        assert_eq!(extension.unparsed_byte_count(), 0);
        assert_eq!(extension.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn name_record_preserves_formula_and_plain_continue_segments() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&BOF.to_le_bytes());
        bytes.extend_from_slice(&16u16.to_le_bytes());
        bytes.extend_from_slice(&0x0600u16.to_le_bytes());
        bytes.extend_from_slice(&0x0005u16.to_le_bytes());
        bytes.extend_from_slice(&0x1234u16.to_le_bytes());
        bytes.extend_from_slice(&0x07ccu16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&6u32.to_le_bytes());

        let mut logical = Vec::new();
        logical.extend_from_slice(&NameFlags::BUILT_IN.bits().to_le_bytes());
        logical.extend_from_slice(&[0, 1]);
        logical.extend_from_slice(&3u16.to_le_bytes());
        logical.extend_from_slice(&0u16.to_le_bytes());
        logical.extend_from_slice(&1u16.to_le_bytes());
        logical.extend_from_slice(&[1, 0, 0, 0]);
        logical.push(0);
        logical.push(6);
        logical.extend_from_slice(&[0x1e, 7, 0]);
        logical.push(b'm');
        let split = 17;
        bytes.extend_from_slice(&NAME.to_le_bytes());
        bytes.extend_from_slice(&(split as u16).to_le_bytes());
        bytes.extend_from_slice(&logical[..split]);
        bytes.extend_from_slice(&CONTINUE.to_le_bytes());
        bytes.extend_from_slice(&((logical.len() - split) as u16).to_le_bytes());
        bytes.extend_from_slice(&logical[split..]);
        bytes.extend_from_slice(&EOF.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());

        let parsed = BiffStream::from_bytes(&bytes).unwrap();
        let BiffRecordData::Name(name) = &parsed.records[1].data else {
            panic!("expected typed Name record");
        };
        assert_eq!(name.physical_segment_lengths.len(), 2);
        assert!(name.formula.unparsed_tail.is_empty());
        assert_eq!(name.custom_menu.characters, b"m");
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn devmodew_core_preserves_fields_and_driver_private_data() {
        let mut device_name = [0u16; 32];
        device_name[0] = b'P' as u16;
        let value = DevModeW {
            device_name,
            specification_version: 0x0401,
            driver_version: 7,
            declared_public_size: 100,
            declared_driver_extra_size: 3,
            fields: DevModeFields::from_bits_retain(0x8000_0001),
            public_fields: DevModeWPublic::Core100(DevModeWCore100 {
                orientation: 1,
                paper_size: 9,
                paper_length: 0,
                paper_width: 0,
                scale: 100,
                copies: 1,
                default_source: 7,
                print_quality: u16::MAX,
                color: 2,
                duplex: 1,
                y_resolution: 600,
                tt_option: 2,
            }),
            driver_extra: vec![1, 2, 3],
            driver_extra_complete: true,
            trailing: Vec::new(),
        };
        let bytes = value.to_bytes().unwrap();
        assert_eq!(DevModeW::from_bytes(&bytes), Some(value));
    }

    #[test]
    fn autofilter_obj_uses_lbs_continued_size_as_metadata() {
        let bytes = [
            0x15, 0x00, 0x12, 0x00, 0x14, 0x00, 0x01, 0x00, 0x01, 0x21, 0x00, 0x00, 0x00, 0x00,
            0xa0, 0x9b, 0xb4, 0x05, 0x00, 0x00, 0x00, 0x00, 0x0c, 0x00, 0x14, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x64, 0x00, 0x01, 0x00, 0x0a, 0x00, 0x00, 0x00,
            0x10, 0x00, 0x01, 0x00, 0x13, 0x00, 0xee, 0x1f, 0x00, 0x00, 0x00, 0x00, 0x04, 0x00,
            0x01, 0x03, 0x00, 0x00, 0x02, 0x00, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        let value = ObjRecord::from_bytes(&bytes, Limits::default()).unwrap();
        let ObjSubrecordData::ListBox(list) = &value.subrecords[2].data else {
            panic!("expected typed FtLbsData");
        };
        assert_eq!(list.continued_size, 0x1fee);
        assert_eq!(list.line_count, 0);
        assert_eq!(list.selected_index, 4);
        assert_eq!(list.drop_data.as_ref().unwrap().line_count, 8);
        assert!(
            value
                .subrecords
                .iter()
                .all(|subrecord| !matches!(subrecord.data, ObjSubrecordData::Raw(_)))
        );
        assert_eq!(value.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn self_relative_security_descriptor_is_fully_typed_and_exact() {
        let bytes = [
            0x01, 0x00, 0x04, 0x80, 0x40, 0x00, 0x00, 0x00, 0x4c, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x14, 0x00, 0x00, 0x00, 0x02, 0x00, 0x2c, 0x00, 0x01, 0x00, 0x00, 0x00,
            0x01, 0x00, 0x24, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x05, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x05, 0x15, 0x00, 0x00, 0x00, 0x99, 0xb8, 0xcb, 0x8a, 0x98, 0x00, 0xc0, 0x1a,
            0xfe, 0x9a, 0x22, 0x80, 0xeb, 0x03, 0x00, 0x00, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x00,
        ];
        let descriptor = SecurityDescriptor::from_bytes(&bytes).unwrap();
        assert!(
            descriptor
                .control
                .contains(SecurityDescriptorControl::SELF_RELATIVE)
        );
        let dacl = descriptor.dacl.as_ref().unwrap();
        assert_eq!(dacl.value.entries.len(), 1);
        assert_eq!(dacl.value.entries[0].ace_type, BasicAceType::AccessDenied);
        assert_eq!(dacl.value.entries[0].trustee.sub_authorities.len(), 5);
        assert_eq!(descriptor.owner.as_ref().unwrap().offset, 0x40);
        assert_eq!(descriptor.group.as_ref().unwrap().offset, 0x4c);
        assert!(descriptor.padding.is_empty());
        assert_eq!(descriptor.to_bytes(bytes.len()).unwrap(), bytes);
    }

    #[test]
    fn real_time_data_valid_variant_is_fully_typed_and_exact() {
        let value = RealTimeDataRecord {
            header: FrtHeader {
                record_type: REAL_TIME_DATA,
                flags: FrtFlags::empty(),
                reserved: 0,
            },
            shared_prefix_character_count: 0,
            topic: RtdTopicString {
                declared_unit_count: 12,
                flags: 0,
                substrings: vec![
                    XlStringCharacters::Compressed(b"Prog".to_vec()),
                    XlStringCharacters::Compressed(Vec::new()),
                    XlStringCharacters::Compressed(b"topic".to_vec()),
                ],
            },
            operation: RtdOperation::Number {
                bits: 42.5f64.to_bits(),
            },
            cells: vec![RtdCellReference {
                row: 7,
                column: 5,
                sheet_index: 2,
            }],
        };
        let bytes = encode_sdk(&value).unwrap();
        let decoded: RealTimeDataRecord = parse_sdk(&bytes, 0, REAL_TIME_DATA).unwrap();
        assert_eq!(decoded, value);
        assert_eq!(encode_sdk(&decoded).unwrap(), bytes);
    }

    #[test]
    fn real_time_data_recovers_error_with_corrupt_discriminator() {
        let value = RealTimeDataRecord {
            header: FrtHeader {
                record_type: REAL_TIME_DATA,
                flags: FrtFlags::empty(),
                reserved: 0,
            },
            shared_prefix_character_count: 0,
            topic: RtdTopicString {
                declared_unit_count: 1,
                flags: 0,
                substrings: vec![XlStringCharacters::Compressed(Vec::new())],
            },
            operation: RtdOperation::ErrorWithCorruptDiscriminator {
                discriminator: 0x0000_dd10,
                value: 42,
            },
            cells: vec![RtdCellReference {
                row: 7,
                column: 5,
                sheet_index: 0,
            }],
        };
        let bytes = encode_sdk(&value).unwrap();
        let decoded: RealTimeDataRecord = parse_sdk(&bytes, 0, REAL_TIME_DATA).unwrap();
        assert_eq!(decoded, value);
        assert_eq!(encode_sdk(&decoded).unwrap(), bytes);
    }

    #[test]
    fn hyperlink_recovers_truncated_url_moniker() {
        let mut bytes = 2u32.to_le_bytes().to_vec();
        bytes.extend_from_slice(
            &(HyperlinkFlags::HAS_MONIKER | HyperlinkFlags::ABSOLUTE)
                .bits()
                .to_le_bytes(),
        );
        bytes.extend_from_slice(&URL_MONIKER_CLASS_ID);
        bytes.extend_from_slice(&10u32.to_le_bytes());
        bytes.extend_from_slice(&[b'A', 0, b'B', 0]);

        let value = HyperlinkObject::parse_truncated_url_moniker(&bytes).unwrap();
        assert_eq!(
            value,
            HyperlinkObject::TruncatedUrlMoniker {
                stream_version: 2,
                flags: HyperlinkFlags::HAS_MONIKER | HyperlinkFlags::ABSOLUTE,
                class_id: URL_MONIKER_CLASS_ID,
                declared_byte_length: 10,
                address: vec![u16::from(b'A'), u16::from(b'B')],
            }
        );
        assert_eq!(value.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn sort_preserves_signed_custom_order_and_unicode_keys() {
        let value = SortRecord {
            options: SortOptions {
                sort_columns: true,
                descending: [true, false, true],
                case_sensitive: true,
                custom_list_index: -3,
                alternate_method: true,
                reserved: 0x1f,
            },
            keys: [
                Some(BiffUnicodeString {
                    flags: 1,
                    characters: XlStringCharacters::Unicode(vec![0x5217, 0x4e00]),
                    trailing_byte: None,
                }),
                None,
                Some(BiffUnicodeString {
                    flags: 0,
                    characters: XlStringCharacters::Compressed(b"C".to_vec()),
                    trailing_byte: None,
                }),
            ],
            reserved: 0xa5,
        };
        let bytes = encode_sdk(&value).unwrap();
        let decoded: SortRecord = parse_sdk(&bytes, 0, SORT).unwrap();
        assert_eq!(decoded, value);
        assert_eq!(encode_sdk(&decoded).unwrap(), bytes);
    }

    #[test]
    fn sort_conditions_cover_custom_lists_and_icon_sets() {
        let range = Rfx {
            first_row: 1,
            last_row: 9,
            first_column: 2,
            last_column: 2,
        };
        let values = [
            SortCondition {
                descending: true,
                reserved: 0x155,
                range,
                data: SortConditionData::Value {
                    value: 0,
                    reserved: 0x1234_5678,
                },
                custom_list: Some(BiffUnicodeString {
                    flags: 1,
                    characters: XlStringCharacters::Unicode(vec![0x7532, 0x4e59]),
                    trailing_byte: None,
                }),
            },
            SortCondition {
                descending: false,
                reserved: 0,
                range,
                data: SortConditionData::Icon {
                    icon_set: 7,
                    icon_index: -1,
                },
                custom_list: None,
            },
        ];
        for value in values {
            let bytes = encode_sdk(&value).unwrap();
            let decoded: SortCondition = parse_sdk(&bytes, 0, CONTINUE_FRT12).unwrap();
            assert_eq!(decoded, value);
            assert_eq!(encode_sdk(&decoded).unwrap(), bytes);
        }
    }

    #[test]
    fn auto_filter_non_string_operands_are_static_and_exact() {
        let value = AutoFilterRecord {
            entry_index: 4,
            options: AutoFilterOptions {
                join_or: true,
                simple: [false, true],
                top_n: false,
                top: true,
                percent: false,
                top_count: 17,
            },
            operands: [
                AutoFilterOperand {
                    comparison: 3,
                    value: AutoFilterOperandValue::Rk {
                        value: 0x1234_567b,
                        unused: 0xaabb_ccdd,
                    },
                    string: None,
                },
                AutoFilterOperand {
                    comparison: 5,
                    value: AutoFilterOperandValue::BooleanOrError {
                        value: 1,
                        unused1: 0x2345,
                        unused2: 0x6789_abcd,
                    },
                    string: None,
                },
            ],
        };
        let bytes = encode_sdk(&value).unwrap();
        let decoded: AutoFilterRecord = parse_sdk(&bytes, 0, AUTO_FILTER).unwrap();
        assert_eq!(decoded, value);
        assert_eq!(encode_sdk(&decoded).unwrap(), bytes);
    }

    #[test]
    fn cleared_sx_format_is_static_and_exact() {
        let value = SxFormatRecord {
            formatting_applied: false,
            reserved: 0x0abc,
            differential_format_byte_count: 0,
        };
        let bytes = encode_sdk(&value).unwrap();
        let decoded: SxFormatRecord = parse_sdk(&bytes, 0, SX_FORMAT).unwrap();
        assert_eq!(decoded, value);
        assert_eq!(encode_sdk(&decoded).unwrap(), bytes);
    }

    #[test]
    fn wopt_preserves_future_bytes_and_screen_enum() {
        let value = WOptRecord {
            header: FrtHeaderOld {
                record_type: W_OPT,
                flags: FrtFlags::empty(),
            },
            flags: WOptFlags::ALLOW_PNG | WOptFlags::from_bits_retain(0x8000),
            screen_size: WebScreenSize::Pixels1920x1200,
            reserved: 0xa5,
            pixels_per_inch: 240,
            code_page: 65001,
            component_location: LpWideString {
                character_count: 1,
                characters: vec!['/' as u16],
            },
            future: vec![1, 2, 3, 4],
        };
        let bytes = encode_sdk(&value).unwrap();
        let decoded: WOptRecord = parse_sdk(&bytes, 0, W_OPT).unwrap();
        assert_eq!(decoded, value);
        assert_eq!(encode_sdk(&decoded).unwrap(), bytes);
    }

    #[test]
    fn param_qry_value_and_reference_variants_are_static() {
        let values = [
            ParamQryRecord {
                fixed: ParamQryFixed {
                    sql_type: 8,
                    parameter_type: 1,
                    unused1: true,
                    non_default_name: false,
                    unused2: 0x123,
                    value_type: 1,
                    boolean_value: 0xabcd,
                },
                data: ParamQryData::Number {
                    bits: 12.5f64.to_bits(),
                },
            },
            ParamQryRecord {
                fixed: ParamQryFixed {
                    sql_type: 4,
                    parameter_type: 2,
                    unused1: false,
                    non_default_name: true,
                    unused2: 0,
                    value_type: 0,
                    boolean_value: 0,
                },
                data: ParamQryData::Reference(
                    FormulaTokenStream::from_bytes(&[0x1e, 7, 0]).unwrap(),
                ),
            },
        ];
        for value in values {
            let bytes = encode_sdk(&value).unwrap();
            let decoded: ParamQryRecord = parse_sdk(&bytes, 0, PARAM_QRY).unwrap();
            assert_eq!(decoded, value);
            assert_eq!(encode_sdk(&decoded).unwrap(), bytes);
        }
    }

    #[test]
    fn mac_office_11_future_records_are_static_and_exact() {
        let header = FrtHeader {
            record_type: PLV_MAC,
            flags: FrtFlags::empty(),
            reserved: 0,
        };
        let plv = PlvMacRecord {
            header,
            flags: PlvMacFlags::PRINT_SCALE_NOT_SHEET_SCALE,
            zoom_scale: 100,
        };
        let plv_bytes = encode_sdk(&plv).unwrap();
        assert_eq!(plv_bytes.len(), 17);
        assert_eq!(
            parse_sdk::<PlvMacRecord>(&plv_bytes, 0, PLV_MAC).unwrap(),
            plv
        );

        let lnext = LnextRecord {
            header: FrtHeader {
                record_type: LNEXT,
                ..header
            },
            color_rgb: 0x00a1_b2c3,
            opacity: 0x0000_8000,
            line_width: 0x0001_0000,
        };
        let lnext_bytes = encode_sdk(&lnext).unwrap();
        assert_eq!(lnext_bytes.len(), 24);
        assert_eq!(
            parse_sdk::<LnextRecord>(&lnext_bytes, 0, LNEXT).unwrap(),
            lnext
        );

        let marker = MkrExtRecord {
            header: FrtHeader {
                record_type: MKR_EXT,
                ..header
            },
            foreground_rgb: 0x0011_2233,
            background_rgb: 0x0044_5566,
            opacity: 0x0000_ffff,
        };
        let marker_bytes = encode_sdk(&marker).unwrap();
        assert_eq!(marker_bytes.len(), 24);
        assert_eq!(
            parse_sdk::<MkrExtRecord>(&marker_bytes, 0, MKR_EXT).unwrap(),
            marker
        );

        let color_options = CrtCoOptRecord {
            header: FrtHeader {
                record_type: CRT_CO_OPT,
                ..header
            },
            color_scheme: 7,
            flags: CrtCoOptFlags::SHADED | CrtCoOptFlags::GRAYSCALE,
            compatibility_padding: Some(0),
        };
        let color_options_bytes = encode_sdk(&color_options).unwrap();
        assert_eq!(color_options_bytes.len(), 20);
        assert_eq!(
            parse_sdk::<CrtCoOptRecord>(&color_options_bytes, 0, CRT_CO_OPT).unwrap(),
            color_options
        );

        let architecture = FrtArchIdRecord {
            header: FrtHeader {
                record_type: FRT_ARCH_ID,
                ..header
            },
            architecture_id: 2,
        };
        let architecture_bytes = encode_sdk(&architecture).unwrap();
        assert_eq!(architecture_bytes.len(), 16);
        assert_eq!(
            parse_sdk::<FrtArchIdRecord>(&architecture_bytes, 0, FRT_ARCH_ID).unwrap(),
            architecture
        );
    }
}
