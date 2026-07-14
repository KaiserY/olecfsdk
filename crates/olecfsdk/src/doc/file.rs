//! Typed file and primary content-tree roots for the Word binary format.

use std::{collections::BTreeMap, path::Path};

use crate::{
    Error, Result,
    cfb::CompoundFile,
    io::BinaryFormat,
    limits::Limits,
    parse::{
        ParseOptions, ParseOutcome, compound_from_bytes, compound_from_path, compound_outcome,
    },
};

use super::{
    Bookmarks, ChpxFkp, Clx, DocOfficeArtContent, Fib, FibBaseFlags, FibFcLcb, FieldDocumentPart,
    FieldTable, FkpPageNumber, FontTable, PapxFkp, PlcBte, PlcfSed, Sepx, StyleSheet, TextPiece,
    TextPieceCharacters,
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
    pub character_format_pages: Vec<DocFkpPage<ChpxFkp>>,
    pub paragraph_format_pages: Vec<DocFkpPage<PapxFkp>>,
    pub section_properties: Vec<DocSectionProperties>,
    physical_bytes: Vec<u8>,
    source_fib_len: usize,
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
    pub office_art: Option<DocLocated<DocOfficeArtContent>>,
    physical_bytes: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocLocatedBookmarks {
    pub names_location: FibFcLcb,
    pub starts_location: FibFcLcb,
    pub ends_location: FibFcLcb,
    pub value: Bookmarks,
}

/// MS-DOC assigns structure-specific offsets into this stream. Until each
/// referenced payload is promoted, the stream remains one explicit physical
/// node rather than being mislabeled as arbitrary content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocDataStream {
    pub physical_bytes: Vec<u8>,
}

/// Complete file root with the primary MS-DOC content structures linked into
/// a Rust tree. Unmanaged CFB entries remain available in `compound_file`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocFile {
    pub compound_file: CompoundFile,
    pub word_document: DocWordDocumentStream,
    pub table: DocTableStream,
    pub data: Option<DocDataStream>,
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
            diagnostics,
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
            if let Some(field) = parse_optional(
                &table_bytes,
                Some(location),
                "Plcfld",
                FieldTable::from_bytes,
            )? {
                fields.insert(part, field);
            }
        }
        let bookmarks = parse_bookmarks(&fib, &table_bytes)?;
        let office_art = parse_optional(
            &table_bytes,
            fib.office_art_content_location(),
            "OfficeArtContent",
            DocOfficeArtContent::from_bytes,
        )?;

        let text_pieces = parse_text_pieces(&clx.value, &word_bytes, limits)?;
        let character_format_pages =
            parse_fkp_pages(&character_bin_table.value, &word_bytes, ChpxFkp::from_bytes)?;
        let paragraph_format_pages =
            parse_fkp_pages(&paragraph_bin_table.value, &word_bytes, PapxFkp::from_bytes)?;
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
        let word_document = DocWordDocumentStream {
            fib,
            text_pieces,
            character_format_pages,
            paragraph_format_pages,
            section_properties,
            physical_bytes: word_bytes,
            source_fib_len,
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
            office_art,
            physical_bytes: table_bytes,
        };
        let data = compound_file
            .stream(DATA_STREAM)
            .map(|bytes| DocDataStream {
                physical_bytes: bytes.to_vec(),
            });
        Ok(ParseOutcome::new(
            Self {
                compound_file,
                word_document,
                table,
                data,
            },
            diagnostics,
        ))
    }

    pub fn to_compound_file(&self) -> Result<CompoundFile> {
        self.validate_links()?;
        let mut word = self.word_document.physical_bytes.clone();
        let mut table = self.table.physical_bytes.clone();

        patch_prefix(
            &mut word,
            self.word_document.source_fib_len,
            self.word_document.fib.to_bytes()?,
            "FIB",
        )?;
        patch_located(&mut table, &self.table.clx, Clx::to_bytes, "CLX")?;
        patch_located(
            &mut table,
            &self.table.character_bin_table,
            PlcBte::to_bytes,
            "PlcBteChpx",
        )?;
        patch_located(
            &mut table,
            &self.table.paragraph_bin_table,
            PlcBte::to_bytes,
            "PlcBtePapx",
        )?;
        patch_located(
            &mut table,
            &self.table.sections,
            PlcfSed::to_bytes,
            "PlcfSed",
        )?;
        if let Some(styles) = &self.table.styles {
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
        if let Some(office_art) = &self.table.office_art {
            patch_located(
                &mut table,
                office_art,
                DocOfficeArtContent::to_bytes,
                "OfficeArtContent",
            )?;
        }

        for piece in &self.word_document.text_pieces {
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
            patch_at(
                &mut word,
                usize::try_from(piece.value.file_offset)
                    .map_err(|_| Error::Limit("text piece offset exceeds usize".into()))?,
                bytes.len(),
                bytes,
                "text piece",
            )?;
        }
        for page in &self.word_document.character_format_pages {
            patch_at(
                &mut word,
                page.page.byte_offset()?,
                512,
                page.value.to_bytes()?,
                "ChpxFkp",
            )?;
        }
        for page in &self.word_document.paragraph_format_pages {
            patch_at(
                &mut word,
                page.page.byte_offset()?,
                512,
                page.value.to_bytes()?,
                "PapxFkp",
            )?;
        }
        for section in &self.word_document.section_properties {
            if let Some(value) = &section.value {
                patch_at(
                    &mut word,
                    usize::try_from(section.offset)
                        .map_err(|_| Error::invalid(0, "negative Sepx offset"))?,
                    section.physical_len,
                    value.to_bytes()?,
                    "Sepx",
                )?;
            }
        }

        let mut compound = self.compound_file.clone();
        compound.replace_stream(WORD_DOCUMENT_STREAM, word)?;
        compound.replace_stream(self.table.name.path(), table)?;
        match &self.data {
            Some(data) => {
                compound.create_or_replace_stream(DATA_STREAM, data.physical_bytes.clone())?;
            }
            None if compound.is_stream(DATA_STREAM) => {
                compound.remove_stream(DATA_STREAM)?;
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

    fn validate_links(&self) -> Result<()> {
        let fib = &self.word_document.fib;
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

        let pieces = &self.table.clx.value.piece_table;
        if pieces.pieces.len() != self.word_document.text_pieces.len()
            || pieces.character_positions.len() != pieces.pieces.len() + 1
        {
            return Err(Error::invalid(0, "CLX/text-piece tree cardinality changed"));
        }
        for (index, (descriptor, piece)) in pieces
            .pieces
            .iter()
            .zip(&self.word_document.text_pieces)
            .enumerate()
        {
            let compressed = matches!(&piece.value.characters, TextPieceCharacters::Compressed(_));
            if piece.piece_index != index
                || piece.value.cp_start != pieces.character_positions[index]
                || piece.value.cp_end != pieces.character_positions[index + 1]
                || piece.value.file_offset != descriptor.file_position.byte_offset()
                || compressed != descriptor.file_position.compressed
            {
                return Err(Error::invalid(0, "CLX/text-piece link changed"));
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

fn patch_located<T>(
    target: &mut [u8],
    located: &DocLocated<T>,
    encode: impl Fn(&T) -> Result<Vec<u8>>,
    label: &str,
) -> Result<()> {
    patch_location(target, located.location, encode(&located.value)?, label)
}

fn patch_location(
    target: &mut [u8],
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

fn patch_prefix(target: &mut [u8], expected: usize, encoded: Vec<u8>, label: &str) -> Result<()> {
    patch_at(target, 0, expected, encoded, label)
}

fn patch_at(
    target: &mut [u8],
    offset: usize,
    expected: usize,
    encoded: Vec<u8>,
    label: &str,
) -> Result<()> {
    if encoded.len() != expected {
        return Err(Error::invalid(
            offset as u64,
            format!("{label} size changed; relocation is not implicit"),
        ));
    }
    target
        .get_mut(offset..offset.saturating_add(expected))
        .ok_or_else(|| Error::invalid(offset as u64, format!("{label} exceeds stream")))?
        .copy_from_slice(&encoded);
    Ok(())
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
