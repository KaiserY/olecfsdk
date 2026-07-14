//! Typed file root for the PowerPoint binary format.

use std::path::Path;

use crate::{
    Error, Result,
    cfb::CompoundFile,
    io::BinaryFormat,
    limits::Limits,
    parse::{
        ParseDiagnostic, ParseDiagnosticCode, ParseOptions, ParseOutcome, SpecificationReference,
        compound_from_bytes, compound_from_path, compound_outcome,
    },
    save::SaveOptions,
};

use super::{
    BinaryTagData, CurrentUserData, CurrentUserStream, ExternalStorageAtom, PicturesStream,
    PowerPointDocument, PptRecord, PptRecordData, PptRecordSequence,
};

const DOCUMENT_STREAM: &str = "/PowerPoint Document";
const CURRENT_USER_STREAM: &str = "/Current User";
const PICTURES_STREAM: &str = "/Pictures";

/// Complete typed root for a PowerPoint binary file.
///
/// The document stream remains a recursive [`super::PptRecordSequence`]; no
/// content is flattened into text, slide summaries, or image shortcuts.
#[derive(Clone, Debug, PartialEq)]
pub struct PptFile {
    pub compound_file: CompoundFile,
    pub document: PowerPointDocument,
    pub current_user: CurrentUserStream,
    pub pictures: Option<PicturesStream>,
}

impl PptFile {
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
        let compound = compound_from_path(path.as_ref(), options, BinaryFormat::Ppt)?;
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
        let compound = compound_from_bytes(bytes, options, BinaryFormat::Ppt)?;
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
        let compound = compound_outcome(compound_file, options, BinaryFormat::Ppt)?;
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
        let document = compound_file
            .stream(DOCUMENT_STREAM)
            .ok_or_else(|| Error::invalid(0, "PowerPoint Document stream is missing"))
            .and_then(|bytes| PowerPointDocument::from_bytes_with_limits(bytes, options.limits))?;
        audit_record_sequence(&document.records, 0, options.is_strict(), &mut diagnostics)?;
        let current_user = compound_file
            .stream(CURRENT_USER_STREAM)
            .ok_or_else(|| {
                Error::invalid(0, "required Current User Stream is missing (MS-PPT 2.1.1)")
            })
            .and_then(CurrentUserStream::from_bytes)?;
        match &current_user.data {
            CurrentUserData::Parsed(atom) => {
                audit_current_user(&current_user, options.is_strict(), &mut diagnostics)?;
                if let Err(error) = document.incremental_save_chain(atom) {
                    let offset = error
                        .offset()
                        .unwrap_or(u64::from(atom.offset_to_current_edit));
                    if options.is_strict() {
                        return Err(Error::invalid(
                            offset,
                            format!(
                                "PowerPoint Document Stream violates the user-edit chain in MS-PPT 2.1.2: {error}"
                            ),
                        ));
                    }
                    diagnostics.push(ParseDiagnostic::warning(
                        ParseDiagnosticCode::InvalidReference,
                        BinaryFormat::Ppt,
                        Some(DOCUMENT_STREAM),
                        Some(offset),
                        "UserEditAtom chain",
                        SpecificationReference {
                            document: "MS-PPT",
                            section: "2.1.2",
                        },
                        format!("preserved a broken user-edit chain: {error}"),
                    ));
                }
            }
            CurrentUserData::Compatibility(_) if options.is_strict() => {
                return Err(Error::invalid(
                    0,
                    "Current User Stream does not contain a conforming CurrentUserAtom",
                ));
            }
            CurrentUserData::Compatibility(_) => diagnostics.push(ParseDiagnostic::warning(
                ParseDiagnosticCode::NonconformingRecord,
                BinaryFormat::Ppt,
                Some(CURRENT_USER_STREAM),
                Some(0),
                "CurrentUserAtom",
                SpecificationReference {
                    document: "MS-PPT",
                    section: "2.3.2",
                },
                "preserved a nonconforming CurrentUserAtom body",
            )),
            CurrentUserData::Truncated(_) if options.is_strict() => {
                return Err(Error::invalid(
                    0,
                    "CurrentUserAtom body is shorter than RecordHeader.recLen",
                ));
            }
            CurrentUserData::Truncated(_) => diagnostics.push(ParseDiagnostic::warning(
                ParseDiagnosticCode::TruncatedRecord,
                BinaryFormat::Ppt,
                Some(CURRENT_USER_STREAM),
                Some(0),
                "CurrentUserAtom",
                SpecificationReference {
                    document: "MS-PPT",
                    section: "2.3.2",
                },
                "preserved the available prefix of a truncated CurrentUserAtom",
            )),
        }
        let pictures = compound_file
            .stream(PICTURES_STREAM)
            .map(|bytes| PicturesStream::from_bytes_with_limits(bytes, options.limits))
            .transpose()?;
        if let Some(PicturesStream::Partial(partial)) = &pictures {
            if options.is_strict() {
                return Err(Error::invalid(
                    0,
                    format!(
                        "Pictures Stream violates MS-PPT 2.1.3 and MS-ODRAW 2.2.21: {}",
                        partial.reason
                    ),
                ));
            }
            diagnostics.push(ParseDiagnostic::warning(
                ParseDiagnosticCode::InvalidStreamPreserved,
                BinaryFormat::Ppt,
                Some(PICTURES_STREAM),
                Some(0),
                "OfficeArtBStoreDelay",
                SpecificationReference {
                    document: "MS-PPT",
                    section: "2.1.3",
                },
                format!("preserved a partial Pictures Stream: {}", partial.reason),
            ));
        }
        Ok(ParseOutcome::new(
            Self {
                compound_file,
                document,
                current_user,
                pictures,
            },
            diagnostics,
        ))
    }

    /// Rebuilds all managed streams from their typed trees and returns CFB.
    pub fn to_compound_file(&self) -> Result<CompoundFile> {
        self.to_compound_file_with_options(SaveOptions::default())
    }

    pub fn to_compound_file_preserving_compatibility(&self) -> Result<CompoundFile> {
        self.to_compound_file_with_options(SaveOptions::preserving_compatibility())
    }

    pub fn to_compound_file_with_options(&self, options: SaveOptions) -> Result<CompoundFile> {
        if !options.preserves_compatibility() {
            if !matches!(&self.current_user.data, CurrentUserData::Parsed(_)) {
                return Err(Error::invalid(
                    0,
                    "strict save rejects a nonconforming CurrentUserAtom",
                ));
            }
            audit_current_user(&self.current_user, true, &mut Vec::new())?;
            audit_record_sequence(&self.document.records, 0, true, &mut Vec::new())?;
            if let CurrentUserData::Parsed(atom) = &self.current_user.data {
                self.document.incremental_save_chain(atom)?;
            }
            if matches!(&self.pictures, Some(PicturesStream::Partial(_))) {
                return Err(Error::invalid(
                    0,
                    "strict save rejects a partial Pictures Stream",
                ));
            }
        }
        let mut compound = self.compound_file.clone();
        compound.replace_stream(DOCUMENT_STREAM, self.document.to_bytes()?)?;
        compound.replace_stream(CURRENT_USER_STREAM, self.current_user.to_bytes()?)?;
        sync_optional_stream(
            &mut compound,
            PICTURES_STREAM,
            self.pictures.as_ref().map(PicturesStream::to_bytes),
        )?;
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

fn audit_current_user(
    stream: &CurrentUserStream,
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<()> {
    let CurrentUserData::Parsed(atom) = &stream.data else {
        return Ok(());
    };
    let mut violations = Vec::new();
    if stream.header.version != 0 || stream.header.instance != 0 {
        violations.push(format!(
            "RecordHeader has recVer {:#x} and recInstance {:#x}, expected 0 and 0",
            stream.header.version, stream.header.instance
        ));
    }
    if atom.fixed_size != 20 {
        violations.push(format!("size is {}, expected 20", atom.fixed_size));
    }
    if !matches!(atom.header_token, 0xe391_c05f | 0xf3d1_c4df) {
        violations.push(format!(
            "headerToken is {:#010x}, outside the two specified values",
            atom.header_token
        ));
    }
    if atom.declared_user_name_byte_length > 255 {
        violations.push(format!(
            "lenUserName is {}, greater than 255",
            atom.declared_user_name_byte_length
        ));
    }
    if atom.document_file_version != 0x03f4 || atom.major_version != 3 || atom.minor_version != 0 {
        violations.push(format!(
            "file/storage version is {:#06x}/{}.{}, expected 0x03f4/3.0",
            atom.document_file_version, atom.major_version, atom.minor_version
        ));
    }
    if !matches!(atom.release_version, 8 | 9) {
        violations.push(format!(
            "relVersion is {}, expected 8 or 9",
            atom.release_version
        ));
    }
    if !atom.trailing.is_empty() {
        violations.push(format!(
            "CurrentUserAtom contains {} trailing bytes after its specified fields",
            atom.trailing.len()
        ));
    }
    if !violations.is_empty() {
        report_current_user_issue(
            strict,
            diagnostics,
            ParseDiagnosticCode::NonconformingRecord,
            violations.join("; "),
        )?;
    }
    if let Some(unicode) = &atom.unicode_user_name
        && !unicode.is_complete
    {
        report_current_user_issue(
            strict,
            diagnostics,
            ParseDiagnosticCode::TruncatedRecord,
            format!(
                "unicodeUserName contains {} of the required {} UTF-16 code units",
                unicode.code_units.len(),
                atom.declared_user_name_byte_length
            ),
        )?;
    }
    Ok(())
}

fn report_current_user_issue(
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
    code: ParseDiagnosticCode,
    message: String,
) -> Result<()> {
    if strict {
        return Err(Error::invalid(
            0,
            format!("Current User Stream violates MS-PPT 2.3.2: {message}"),
        ));
    }
    diagnostics.push(ParseDiagnostic::warning(
        code,
        BinaryFormat::Ppt,
        Some(CURRENT_USER_STREAM),
        Some(0),
        "CurrentUserAtom",
        SpecificationReference {
            document: "MS-PPT",
            section: "2.3.2",
        },
        message,
    ));
    Ok(())
}

fn audit_record_sequence(
    sequence: &PptRecordSequence,
    base_offset: u64,
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Result<()> {
    for record in &sequence.records {
        match &record.data {
            PptRecordData::MalformedSpecRecord(value) => report_record_issue(
                strict,
                diagnostics,
                ParseDiagnosticCode::NonconformingRecord,
                record.offset,
                "Record",
                "2.3",
                format!(
                    "record type 0x{:04X} has a body that violates its MS-PPT structure",
                    value.record_type
                ),
            )?,
            PptRecordData::Truncated(bytes) => report_record_issue(
                strict,
                diagnostics,
                ParseDiagnosticCode::TruncatedRecord,
                record.offset,
                "RecordHeader",
                "2.3.1",
                format!(
                    "record type 0x{:04X} declares {} body bytes but only {} remain",
                    record.header.record_type,
                    record.header.declared_length,
                    bytes.len()
                ),
            )?,
            PptRecordData::MalformedTextSpecialInfo(bytes) => report_record_issue(
                strict,
                diagnostics,
                ParseDiagnosticCode::NonconformingRecord,
                record.offset,
                "TextSpecialInfoAtom",
                "2.9.54",
                format!(
                    "TextSpecialInfoAtom contains {} bytes that do not form its rgSIRun array",
                    bytes.len()
                ),
            )?,
            PptRecordData::MalformedStyleTextProp(_)
            | PptRecordData::UnresolvedStyleTextProp(_) => report_record_issue(
                strict,
                diagnostics,
                ParseDiagnosticCode::NonconformingRecord,
                record.offset,
                "StyleTextPropAtom",
                "2.9.44",
                "StyleTextPropAtom does not form the runs required for its corresponding text"
                    .into(),
            )?,
            PptRecordData::MalformedTextMasterStyle(_) => report_record_issue(
                strict,
                diagnostics,
                ParseDiagnosticCode::NonconformingRecord,
                record.offset,
                "TextMasterStyleAtom",
                "2.9.35",
                "TextMasterStyleAtom body does not satisfy its level structures".into(),
            )?,
            PptRecordData::MalformedTextRuler(_) => report_record_issue(
                strict,
                diagnostics,
                ParseDiagnosticCode::NonconformingRecord,
                record.offset,
                "TextRulerAtom",
                "2.9.29",
                "TextRulerAtom body does not satisfy its masked field layout".into(),
            )?,
            PptRecordData::MalformedStyleTextProp9(_) => report_record_issue(
                strict,
                diagnostics,
                ParseDiagnosticCode::NonconformingRecord,
                record.offset,
                "StyleTextProp9Atom",
                "2.9.67",
                "StyleTextProp9Atom body does not satisfy its run layout".into(),
            )?,
            PptRecordData::MalformedTimeVariant(_) => report_record_issue(
                strict,
                diagnostics,
                ParseDiagnosticCode::NonconformingRecord,
                record.offset,
                "TimeVariant",
                "2.8.78",
                "TimeVariant body does not match its discriminant".into(),
            )?,
            PptRecordData::MalformedBlipEntity9 { reason, .. } => report_record_issue(
                strict,
                diagnostics,
                ParseDiagnosticCode::NonconformingRecord,
                record.offset,
                "BlipEntityAtom",
                "2.9.73",
                format!("BlipEntityAtom body is invalid: {reason}"),
            )?,
            PptRecordData::ExternalStorage(storage) => {
                let issue = match storage {
                    ExternalStorageAtom::Parsed(_) => None,
                    ExternalStorageAtom::MalformedCompressed { reason, .. }
                    | ExternalStorageAtom::InvalidCompressed { reason, .. } => {
                        Some(("ExOleObjStgCompressedAtom", "2.10.36", reason.as_str()))
                    }
                    ExternalStorageAtom::InvalidUncompressed { reason, .. } => {
                        Some(("ExOleObjStgUncompressedAtom", "2.10.35", reason.as_str()))
                    }
                    ExternalStorageAtom::UnsupportedInstance { .. } => Some((
                        "ExOleObjStg",
                        "2.10.34",
                        "recInstance is neither compressed nor uncompressed",
                    )),
                };
                if let Some((structure, section, reason)) = issue {
                    report_record_issue(
                        strict,
                        diagnostics,
                        ParseDiagnosticCode::NonconformingRecord,
                        record.offset,
                        structure,
                        section,
                        format!("preserved an invalid external storage record: {reason}"),
                    )?;
                }
            }
            PptRecordData::Container(children) | PptRecordData::ProgTags(children) => {
                audit_record_sequence(
                    children,
                    record.offset.saturating_add(8),
                    strict,
                    diagnostics,
                )?;
            }
            PptRecordData::ProgBinaryTag(value) => audit_record_sequence(
                &value.records,
                record.offset.saturating_add(8),
                strict,
                diagnostics,
            )?,
            PptRecordData::BinaryTagData(BinaryTagData::Records(children)) => {
                audit_record_sequence(
                    children,
                    record.offset.saturating_add(8),
                    strict,
                    diagnostics,
                )?;
            }
            _ => {}
        }
    }
    if !sequence.trailing_header_bytes.is_empty() {
        let offset = sequence
            .records
            .last()
            .map(record_physical_end)
            .unwrap_or(base_offset);
        report_record_issue(
            strict,
            diagnostics,
            ParseDiagnosticCode::TruncatedRecord,
            offset,
            "RecordHeader",
            "2.3.1",
            format!(
                "{} trailing bytes cannot form the required 8-byte RecordHeader",
                sequence.trailing_header_bytes.len()
            ),
        )?;
    }
    Ok(())
}

fn record_physical_end(record: &PptRecord) -> u64 {
    let body_len = match &record.data {
        PptRecordData::Truncated(bytes) => bytes.len() as u64,
        _ => u64::from(record.header.declared_length),
    };
    record.offset.saturating_add(8).saturating_add(body_len)
}

fn report_record_issue(
    strict: bool,
    diagnostics: &mut Vec<ParseDiagnostic>,
    code: ParseDiagnosticCode,
    offset: u64,
    structure: &'static str,
    section: &'static str,
    message: String,
) -> Result<()> {
    if strict {
        return Err(Error::invalid(
            offset,
            format!("PowerPoint Document Stream violates MS-PPT {section}: {message}"),
        ));
    }
    diagnostics.push(ParseDiagnostic::warning(
        code,
        BinaryFormat::Ppt,
        Some(DOCUMENT_STREAM),
        Some(offset),
        structure,
        SpecificationReference {
            document: "MS-PPT",
            section,
        },
        message,
    ));
    Ok(())
}

fn sync_optional_stream(
    compound: &mut CompoundFile,
    path: &str,
    bytes: Option<Result<Vec<u8>>>,
) -> Result<()> {
    match bytes {
        Some(bytes) => {
            compound.create_or_replace_stream(path, bytes?)?;
        }
        None if compound.is_stream(path) => {
            compound.remove_stream(path)?;
        }
        None => {}
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        cfb::Version,
        ppt::{
            CURRENT_USER_ATOM, CurrentUserAtom, DOCUMENT_ATOM, EXTERNAL_OLE_OBJECT_STORAGE,
            PERSIST_DIRECTORY_ATOM, PersistDirectoryAtom, PptRecordHeader, USER_EDIT_ATOM,
            UserEditAtom,
        },
    };

    fn current_user_stream() -> CurrentUserStream {
        let atom = CurrentUserAtom {
            fixed_size: 20,
            header_token: 0xe391_c05f,
            offset_to_current_edit: 16,
            declared_user_name_byte_length: 3,
            document_file_version: 0x03f4,
            major_version: 3,
            minor_version: 0,
            unused: 0,
            ansi_user_name: b"Ada".to_vec(),
            release_version: 8,
            unicode_user_name: None,
            trailing: Vec::new(),
        };
        let (body, following) = atom.to_parts().unwrap();
        assert!(following.is_empty());
        CurrentUserStream {
            header: PptRecordHeader {
                version: 0,
                instance: 0,
                record_type: CURRENT_USER_ATOM,
                declared_length: body.len() as u32,
            },
            data: CurrentUserData::Parsed(atom),
            padding: Vec::new(),
        }
    }

    fn document_with_minimal_chain(mut suffix: Vec<u8>) -> Vec<u8> {
        let persist_body = PersistDirectoryAtom {
            entries: Vec::new(),
        }
        .to_bytes()
        .unwrap();
        assert!(persist_body.is_empty());
        let user_body = UserEditAtom {
            last_slide_id_ref: 0,
            version: 0,
            minor_version: 0,
            major_version: 3,
            offset_last_edit: 0,
            offset_persist_directory: 8,
            doc_persist_id_ref: 0,
            persist_id_seed: 1,
            last_view: 0,
            unused: 0,
            encrypt_session_persist_id_ref: None,
        }
        .to_bytes();
        let mut document = Vec::new();
        PptRecordHeader {
            version: 0,
            instance: 0,
            record_type: 0x779f,
            declared_length: 0,
        }
        .write(&mut document)
        .unwrap();
        PptRecordHeader {
            version: 0,
            instance: 0,
            record_type: PERSIST_DIRECTORY_ATOM,
            declared_length: 0,
        }
        .write(&mut document)
        .unwrap();
        PptRecordHeader {
            version: 0,
            instance: 0,
            record_type: USER_EDIT_ATOM,
            declared_length: user_body.len() as u32,
        }
        .write(&mut document)
        .unwrap();
        document.extend_from_slice(&user_body);
        document.append(&mut suffix);
        document
    }

    fn compound_with_document(document: Vec<u8>) -> CompoundFile {
        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound
            .create_or_replace_stream(DOCUMENT_STREAM, document)
            .unwrap();
        compound
            .create_or_replace_stream(
                CURRENT_USER_STREAM,
                current_user_stream().to_bytes().unwrap(),
            )
            .unwrap();
        compound
    }

    #[test]
    fn file_root_round_trips_the_typed_document_stream() {
        let compound = compound_with_document(document_with_minimal_chain(Vec::new()));
        let file = PptFile::from_compound_file(compound).unwrap();
        assert_eq!(file.document.records.records.len(), 3);
        let reopened = PptFile::from_bytes(&file.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.document, file.document);
    }

    #[test]
    fn unknown_record_is_spec_allowed_in_strict_mode() {
        let mut document = Vec::new();
        PptRecordHeader {
            version: 0,
            instance: 0,
            record_type: 0x779f,
            declared_length: 3,
        }
        .write(&mut document)
        .unwrap();
        document.extend_from_slice(&[1, 2, 3]);

        let document = document_with_minimal_chain(document);
        let file = PptFile::from_compound_file(compound_with_document(document.clone())).unwrap();
        assert!(matches!(
            file.document.records.records.last().unwrap().data,
            PptRecordData::Unknown(_)
        ));
        assert_eq!(
            file.to_compound_file().unwrap().stream(DOCUMENT_STREAM),
            Some(document.as_slice())
        );
    }

    #[test]
    fn malformed_and_truncated_document_records_require_compatible_mode() {
        let mut document = Vec::new();
        PptRecordHeader {
            version: 0,
            instance: 0,
            record_type: DOCUMENT_ATOM,
            declared_length: 1,
        }
        .write(&mut document)
        .unwrap();
        document.push(0xff);
        PptRecordHeader {
            version: 0,
            instance: 0,
            record_type: 0x1234,
            declared_length: 5,
        }
        .write(&mut document)
        .unwrap();
        document.extend_from_slice(&[1, 2, 3]);

        let document = document_with_minimal_chain(document);
        let compound = compound_with_document(document.clone());
        assert!(PptFile::from_compound_file(compound.clone()).is_err());
        let outcome = PptFile::from_compound_file_compatible(compound).unwrap();
        assert_eq!(outcome.diagnostics.len(), 2);
        assert_eq!(
            outcome.diagnostics[0].code,
            ParseDiagnosticCode::NonconformingRecord
        );
        assert_eq!(
            outcome.diagnostics[1].code,
            ParseDiagnosticCode::TruncatedRecord
        );
        assert_eq!(
            outcome.diagnostics[1].location.offset,
            Some((document.len() - 11) as u64)
        );
        assert!(outcome.value.to_compound_file().is_err());
        assert_eq!(
            outcome
                .value
                .to_compound_file_preserving_compatibility()
                .unwrap()
                .stream(DOCUMENT_STREAM),
            Some(document.as_slice())
        );
    }

    #[test]
    fn trailing_record_header_prefix_is_diagnostic() {
        let document = document_with_minimal_chain(vec![1, 2, 3, 4]);
        let compound = compound_with_document(document.clone());
        assert!(PptFile::from_compound_file(compound.clone()).is_err());
        let outcome = PptFile::from_compound_file_compatible(compound).unwrap();
        assert_eq!(outcome.diagnostics.len(), 1);
        assert_eq!(
            outcome.diagnostics[0].code,
            ParseDiagnosticCode::TruncatedRecord
        );
        assert_eq!(
            outcome.diagnostics[0].location.offset,
            Some((document.len() - 4) as u64)
        );
        assert_eq!(
            outcome
                .value
                .to_compound_file_preserving_compatibility()
                .unwrap()
                .stream(DOCUMENT_STREAM),
            Some(document.as_slice())
        );
    }

    #[test]
    fn partial_pictures_stream_requires_compatible_mode() {
        let mut pictures = Vec::new();
        PptRecordHeader {
            version: 0,
            instance: 0,
            record_type: 0xf01e,
            declared_length: 5,
        }
        .write(&mut pictures)
        .unwrap();
        pictures.extend_from_slice(&[1, 2, 3]);
        let mut compound = compound_with_document(document_with_minimal_chain(Vec::new()));
        compound
            .create_or_replace_stream(PICTURES_STREAM, pictures.clone())
            .unwrap();

        assert!(PptFile::from_compound_file(compound.clone()).is_err());
        let outcome = PptFile::from_compound_file_compatible(compound).unwrap();
        assert!(matches!(
            outcome.value.pictures,
            Some(PicturesStream::Partial(_))
        ));
        assert_eq!(outcome.diagnostics.len(), 1);
        assert_eq!(
            outcome.diagnostics[0].code,
            ParseDiagnosticCode::InvalidStreamPreserved
        );
        assert_eq!(
            outcome.diagnostics[0].location.path.as_deref(),
            Some(PICTURES_STREAM)
        );
        assert!(outcome.value.to_compound_file().is_err());
        assert_eq!(
            outcome
                .value
                .to_compound_file_preserving_compatibility()
                .unwrap()
                .stream(PICTURES_STREAM),
            Some(pictures.as_slice())
        );
    }

    #[test]
    fn broken_current_edit_reference_requires_compatible_mode() {
        let compound = compound_with_document(Vec::new());
        assert!(PptFile::from_compound_file(compound.clone()).is_err());
        let outcome = PptFile::from_compound_file_compatible(compound).unwrap();
        assert_eq!(outcome.diagnostics.len(), 1);
        assert_eq!(
            outcome.diagnostics[0].code,
            ParseDiagnosticCode::InvalidReference
        );
        assert_eq!(outcome.diagnostics[0].location.offset, Some(16));
        assert!(outcome.value.to_compound_file().is_err());
        assert!(
            outcome
                .value
                .to_compound_file_preserving_compatibility()
                .is_ok()
        );
    }

    #[test]
    fn invalid_external_storage_requires_compatible_mode() {
        let mut suffix = Vec::new();
        PptRecordHeader {
            version: 0,
            instance: 1,
            record_type: EXTERNAL_OLE_OBJECT_STORAGE,
            declared_length: 5,
        }
        .write(&mut suffix)
        .unwrap();
        suffix.extend_from_slice(&[1, 0, 0, 0, 0xff]);
        let document = document_with_minimal_chain(suffix);
        let compound = compound_with_document(document.clone());

        assert!(PptFile::from_compound_file(compound.clone()).is_err());
        let outcome = PptFile::from_compound_file_compatible(compound).unwrap();
        assert_eq!(outcome.diagnostics.len(), 1);
        assert_eq!(
            outcome.diagnostics[0].code,
            ParseDiagnosticCode::NonconformingRecord
        );
        assert_eq!(
            outcome.diagnostics[0].structure,
            "ExOleObjStgCompressedAtom"
        );
        assert!(outcome.value.to_compound_file().is_err());
        assert_eq!(
            outcome
                .value
                .to_compound_file_preserving_compatibility()
                .unwrap()
                .stream(DOCUMENT_STREAM),
            Some(document.as_slice())
        );
    }

    #[test]
    fn named_malformed_variants_share_the_root_strictness_gate() {
        let sequence = PptRecordSequence {
            records: vec![PptRecord {
                offset: 123,
                header: PptRecordHeader {
                    version: 0,
                    instance: 0,
                    record_type: 0xf142,
                    declared_length: 1,
                },
                data: PptRecordData::MalformedTimeVariant(vec![0xff]),
            }],
            trailing_header_bytes: Vec::new(),
        };
        assert!(audit_record_sequence(&sequence, 0, true, &mut Vec::new()).is_err());
        let mut diagnostics = Vec::new();
        audit_record_sequence(&sequence, 0, false, &mut diagnostics).unwrap();
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(
            diagnostics[0].code,
            ParseDiagnosticCode::NonconformingRecord
        );
        assert_eq!(diagnostics[0].structure, "TimeVariant");
        assert_eq!(diagnostics[0].specification.section, "2.8.78");
    }

    #[test]
    fn required_current_user_stream_is_not_invented_by_compatibility_mode() {
        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound
            .create_or_replace_stream(DOCUMENT_STREAM, Vec::new())
            .unwrap();
        assert!(PptFile::from_compound_file(compound.clone()).is_err());
        assert!(PptFile::from_compound_file_compatible(compound).is_err());
    }

    #[test]
    fn truncated_current_user_is_compatible_only_and_diagnostic() {
        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound
            .create_or_replace_stream(DOCUMENT_STREAM, Vec::new())
            .unwrap();
        let mut bytes = Vec::new();
        PptRecordHeader {
            version: 0,
            instance: 0,
            record_type: CURRENT_USER_ATOM,
            declared_length: 5,
        }
        .write(&mut bytes)
        .unwrap();
        bytes.extend_from_slice(&[1, 2, 3]);
        compound
            .create_or_replace_stream(CURRENT_USER_STREAM, bytes)
            .unwrap();

        assert!(PptFile::from_compound_file(compound.clone()).is_err());
        let outcome = PptFile::from_compound_file_compatible(compound).unwrap();
        assert!(matches!(
            outcome.value.current_user.data,
            CurrentUserData::Truncated(ref bytes) if bytes == &[1, 2, 3]
        ));
        assert_eq!(outcome.diagnostics.len(), 1);
        assert_eq!(
            outcome.diagnostics[0].code,
            ParseDiagnosticCode::TruncatedRecord
        );
        assert_eq!(
            outcome.diagnostics[0].location.path.as_deref(),
            Some(CURRENT_USER_STREAM)
        );
        assert!(outcome.value.to_compound_file().is_err());
        let preserved = outcome
            .value
            .to_compound_file_preserving_compatibility()
            .unwrap();
        assert_eq!(
            preserved.stream(CURRENT_USER_STREAM),
            outcome.value.compound_file.stream(CURRENT_USER_STREAM)
        );
    }

    #[test]
    fn current_user_must_fields_require_compatible_mode() {
        let mut current = current_user_stream();
        let CurrentUserData::Parsed(atom) = &mut current.data else {
            panic!("test Current User stream is not parsed");
        };
        atom.document_file_version = 0;
        atom.release_version = 7;
        let mut compound = compound_with_document(document_with_minimal_chain(Vec::new()));
        compound
            .replace_stream(CURRENT_USER_STREAM, current.to_bytes().unwrap())
            .unwrap();

        assert!(PptFile::from_compound_file(compound.clone()).is_err());
        let outcome = PptFile::from_compound_file_compatible(compound).unwrap();
        assert_eq!(outcome.diagnostics.len(), 1);
        assert_eq!(
            outcome.diagnostics[0].code,
            ParseDiagnosticCode::NonconformingRecord
        );
        assert_eq!(outcome.diagnostics[0].structure, "CurrentUserAtom");
        assert!(outcome.value.to_compound_file().is_err());
        assert!(
            outcome
                .value
                .to_compound_file_preserving_compatibility()
                .is_ok()
        );
    }
}
