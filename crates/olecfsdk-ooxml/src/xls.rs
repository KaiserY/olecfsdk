use olecfsdk::{
    office_art::{OfficeArtClientAnchor, OfficeArtImageFormat, OfficeArtShapeFlags},
    xls::{
        BiffSubstreamKind, CellErrorCode, ExtFontScheme, ExtPropertyData, FontAttributes,
        FontRecord, FormulaOperator, FormulaTokenData, FormulaTokenStream, FullColorExt,
        XfExtRecord, XfRecord, XlStringCharacters, XlsCellValue, XlsCellValueRef, XlsFile,
        XlsFormulaCachedValue, XlsFormulaDefinitionRef, XlsFormulaRef, XlsHyperlinkTarget,
        XlsPictureImageLink, XlsPictureRef, XlsWorkbookView,
    },
};
use ooxmlsdk::{
    common::XmlNamespace,
    namespaces::XmlKnownNamespace,
    parts::spreadsheet_document::SpreadsheetDocument,
    parts::{
        drawings_part::DrawingsPart, image_part::ImagePart,
        shared_string_table_part::SharedStringTablePart, workbook_part::WorkbookPart,
        workbook_styles_part::WorkbookStylesPart, worksheet_comments_part::WorksheetCommentsPart,
        worksheet_part::WorksheetPart,
    },
    schemas::{
        opc_relationships::TargetMode,
        schemas_openxmlformats_org_drawingml_2006_main as a,
        schemas_openxmlformats_org_drawingml_2006_spreadsheet_drawing as xdr,
        schemas_openxmlformats_org_spreadsheetml_2006_main::{
            self as x, Author, Authors, Cell, CellFormula, CellValue, CellValues, Comment,
            CommentList, CommentText, Comments, Hyperlink, Hyperlinks, MergeCell, MergeCells, Row,
            SharedStringItem, SharedStringTable, Sheet, SheetData, SheetStateValues, Sheets, Text,
            Workbook, Worksheet, XstringType,
        },
        xml::SpaceProcessingModeValues,
    },
    sdk::SpreadsheetDocumentType,
    simple_type::{BooleanValue, CoordinateValue},
};

use crate::{
    ConversionCode, ConversionOptions, ConversionOutput, ConversionReport, Disposition, Error,
    LossPolicy, Result, SourceLocation, metadata::convert_core_properties,
};

#[derive(Default)]
struct XlsMediaState {
    parts: Vec<(u32, ImagePart)>,
}

/// Converts a typed BIFF8 workbook root into a SpreadsheetML package.
///
/// The default policy rejects the first known semantic loss. Use
/// [`convert_xls_with_options`] to request an explicit diagnostic report.
pub fn convert_xls(source: &XlsFile) -> Result<ConversionOutput<SpreadsheetDocument>> {
    convert_xls_with_options(source, ConversionOptions::default())
}

/// Converts a typed BIFF8 workbook root with an explicit loss policy.
pub fn convert_xls_with_options(
    source: &XlsFile,
    options: ConversionOptions,
) -> Result<ConversionOutput<SpreadsheetDocument>> {
    let mut report = ConversionReport::default();
    let workbook = source
        .workbooks
        .first()
        .ok_or_else(|| olecfsdk::Error::invalid(0, "XLS has no Workbook stream"))?;
    for workbook_index in 1..source.workbooks.len() {
        unsupported(
            &mut report,
            options,
            ConversionCode::AdditionalWorkbookStreamNotMapped,
            SourceLocation::XlsWorkbook { workbook_index },
        )?;
    }
    let view = workbook.relationships()?;
    if has_unmapped_workbook_features(source, &view) {
        unsupported(
            &mut report,
            options,
            ConversionCode::WorkbookFeatureNotMapped,
            SourceLocation::XlsWorkbook { workbook_index: 0 },
        )?;
    }

    let mut document = SpreadsheetDocument::create(SpreadsheetDocumentType::Workbook);
    let workbook_part = document.add_new_part_auto_id::<WorkbookPart>()?;
    let styles = convert_stylesheet(&view)?;
    let source_pictures = view.pictures()?;
    let mut media = XlsMediaState::default();
    let mut target_sheets = Vec::with_capacity(view.sheets().len());
    for (sheet_index, source_sheet) in view.sheets().iter().copied().enumerate() {
        let source_location = SourceLocation::XlsSheet {
            workbook_index: 0,
            sheet_index,
        };
        if source_sheet.kind() != BiffSubstreamKind::WorksheetOrDialogSheet
            || source_sheet.metadata().sheet_type != 0
        {
            unsupported(
                &mut report,
                options,
                ConversionCode::SheetKindNotMapped,
                source_location,
            )?;
            continue;
        }

        // Row/column dimensions, views, non-comment drawings, and broader
        // sheet-level formatting remain outside this vertical slice.
        unsupported(
            &mut report,
            options,
            ConversionCode::WorksheetFeatureNotMapped,
            source_location,
        )?;
        let index = source_sheet.sparse_cell_index()?;
        let merge_cell = source_sheet
            .merged_cells()
            .map(|range| MergeCell {
                reference: cell_range_reference(
                    range.first_row,
                    range.first_column,
                    range.last_row,
                    range.last_column,
                ),
            })
            .collect::<Vec<_>>();
        let merge_cells = if merge_cell.is_empty() {
            None
        } else {
            Some(MergeCells {
                count: Some(u32::try_from(merge_cell.len()).map_err(|_| {
                    olecfsdk::Error::Limit("XLS merged range count exceeds u32".into())
                })?),
                merge_cell,
            })
        };
        let source_hyperlinks = source_sheet.hyperlinks()?;
        let source_comments = source_sheet.comments()?;
        let mut rows = Vec::new();
        for source_row in index.rows() {
            let mut cells = Vec::new();
            for source_cell in source_row.cells() {
                cells.push(convert_cell(
                    &view,
                    &index,
                    source_cell,
                    sheet_index,
                    options,
                    &mut report,
                    &styles.xf_unmapped,
                )?);
            }
            rows.push(Row {
                row_index: Some(u32::from(source_row.row()) + 1),
                cell: cells,
                ..Default::default()
            });
            report.record(Disposition::Mapped);
        }

        let worksheet_part =
            workbook_part.add_new_part_auto_id::<_, WorksheetPart>(&mut document)?;
        let mut hyperlinks = Vec::with_capacity(source_hyperlinks.len());
        for source_hyperlink in source_hyperlinks {
            let range = source_hyperlink.value();
            let reference = cell_range_reference(
                range.first_row,
                range.first_column,
                range.last_row,
                range.last_column,
            );
            if source_hyperlink.target_frame_name.is_some() {
                unsupported(
                    &mut report,
                    options,
                    ConversionCode::HyperlinkFrameNotMapped,
                    source_location,
                )?;
            }
            let target = match source_hyperlink.target {
                Some(XlsHyperlinkTarget::String(value) | XlsHyperlinkTarget::Url(value)) => {
                    Some(value)
                }
                Some(XlsHyperlinkTarget::File {
                    long_path: Some(value),
                    ..
                }) => Some(value),
                Some(
                    XlsHyperlinkTarget::File {
                        long_path: None, ..
                    }
                    | XlsHyperlinkTarget::Standard { .. },
                ) => {
                    unsupported(
                        &mut report,
                        options,
                        ConversionCode::HyperlinkTargetNotMapped,
                        source_location,
                    )?;
                    None
                }
                None => None,
            };
            let id = target
                .map(|target| {
                    worksheet_part
                        .add_hyperlink_relationship_auto_id(
                            &mut document,
                            target,
                            TargetMode::External,
                        )
                        .map(|relationship| relationship.id().to_owned())
                })
                .transpose()?;
            hyperlinks.push(Hyperlink {
                reference,
                id,
                location: source_hyperlink.location,
                display: source_hyperlink.display_name,
                ..Default::default()
            });
            report.record(Disposition::Mapped);
        }
        let drawing = convert_sheet_pictures(
            source_pictures
                .iter()
                .copied()
                .filter(|picture| picture.sheet().id() == source_sheet.id()),
            sheet_index,
            &worksheet_part,
            &mut document,
            &mut media,
            options,
            &mut report,
        )?;
        worksheet_part.set_root_element(
            &mut document,
            Worksheet {
                xmlns: vec![XmlNamespace::known(XmlKnownNamespace::R)],
                sheet_data: SheetData { row: rows },
                merge_cells,
                hyperlinks: (!hyperlinks.is_empty()).then_some(Hyperlinks {
                    hyperlink: hyperlinks,
                }),
                drawing,
                ..Default::default()
            },
        )?;
        if let Some(comments) =
            convert_comments(source_comments, sheet_index, options, &mut report)?
        {
            let comments_part =
                worksheet_part.add_new_part_auto_id::<_, WorksheetCommentsPart>(&mut document)?;
            comments_part.set_root_element(&mut document, comments)?;
        }
        let relationship_id = workbook_part
            .get_id_of_part(&document, &worksheet_part)
            .expect("a newly added worksheet has a relationship id")
            .to_owned();
        let state = convert_sheet_state(
            source_sheet.metadata().state,
            source_location,
            options,
            &mut report,
        )?;
        target_sheets.push(Sheet {
            name: source_sheet.metadata().name.value.clone(),
            sheet_id: u32::try_from(sheet_index + 1)
                .map_err(|_| olecfsdk::Error::Limit("XLS sheet index exceeds u32".into()))?,
            state,
            id: relationship_id,
            ..Default::default()
        });
        report.record(Disposition::Mapped);
    }

    workbook_part.set_root_element(
        &mut document,
        Workbook {
            xmlns: vec![XmlNamespace::known(XmlKnownNamespace::R)],
            sheets: Sheets {
                sheet: target_sheets,
            },
            ..Default::default()
        },
    )?;
    if let Some(shared_strings) = convert_shared_strings(&view, options, &mut report)? {
        let part = workbook_part.add_new_part_auto_id::<_, SharedStringTablePart>(&mut document)?;
        part.set_root_element(&mut document, shared_strings)?;
    }
    let styles_part = workbook_part.add_new_part_auto_id::<_, WorkbookStylesPart>(&mut document)?;
    styles_part.set_root_element(&mut document, styles.root)?;
    report.record(Disposition::Mapped);
    if let Some(properties) = convert_core_properties(&source.shared, options, &mut report)? {
        let properties_part = document.add_core_file_properties_part()?;
        properties_part.set_root_element(&mut document, properties)?;
    }
    Ok(ConversionOutput { document, report })
}

fn convert_sheet_pictures<'a>(
    source: impl Iterator<Item = XlsPictureRef<'a>>,
    sheet_index: usize,
    worksheet_part: &WorksheetPart,
    document: &mut SpreadsheetDocument,
    media: &mut XlsMediaState,
    options: ConversionOptions,
    report: &mut ConversionReport,
) -> Result<Option<x::Drawing>> {
    let mut source = source.peekable();
    if source.peek().is_none() {
        return Ok(None);
    }
    let mut drawing_part = None;
    let mut anchors = Vec::new();
    for picture in source {
        let location = SourceLocation::XlsDrawing {
            workbook_index: 0,
            sheet_index,
            shape_id: picture.shape().shape_id,
        };
        let image = match picture.image() {
            XlsPictureImageLink::Resolved(image) => image,
            XlsPictureImageLink::Delayed { .. }
            | XlsPictureImageLink::Unsupported
            | XlsPictureImageLink::Missing => {
                unsupported(
                    report,
                    options,
                    ConversionCode::SpreadsheetPictureNotMapped,
                    location,
                )?;
                continue;
            }
        };
        let Some(content_type) = xls_image_content_type(image.format) else {
            unsupported(
                report,
                options,
                ConversionCode::SpreadsheetPictureNotMapped,
                location,
            )?;
            continue;
        };
        let OfficeArtClientAnchor::Words18 { flags, coordinates } = picture.anchor() else {
            unsupported(
                report,
                options,
                ConversionCode::SpreadsheetPictureAnchorNotMapped,
                location,
            )?;
            continue;
        };
        let edit_as = match flags {
            0 => xdr::EditAsValues::TwoCell,
            2 => xdr::EditAsValues::OneCell,
            3 => xdr::EditAsValues::Absolute,
            _ => {
                unsupported(
                    report,
                    options,
                    ConversionCode::SpreadsheetPictureAnchorNotMapped,
                    location,
                )?;
                xdr::EditAsValues::TwoCell
            }
        };
        if coordinates[1] > 1_023
            || coordinates[3] > 255
            || coordinates[5] > 1_023
            || coordinates[7] > 255
        {
            unsupported(
                report,
                options,
                ConversionCode::SpreadsheetPictureAnchorNotMapped,
                location,
            )?;
        }
        if drawing_part.is_none() {
            let part = worksheet_part.add_new_part_auto_id::<_, DrawingsPart>(document)?;
            let relationship_id = worksheet_part
                .get_id_of_part(document, &part)
                .expect("a newly added drawing part has a relationship ID")
                .to_owned();
            drawing_part = Some((part, relationship_id));
        }
        let Some((drawings_part, _)) = drawing_part.as_ref() else {
            unreachable!("the XLS drawing part was initialized immediately above")
        };
        let relationship_id = add_xls_image_relationship(
            picture.blip_identifier(),
            image,
            content_type,
            drawings_part,
            document,
            media,
        )?;
        let mut formatting_loss = picture.shape_type() != 75;
        let source_rectangle = xls_picture_crop(picture, &mut formatting_loss);
        if formatting_loss {
            unsupported(
                report,
                options,
                ConversionCode::SpreadsheetPictureFormattingNotMapped,
                location,
            )?;
        }
        let shape_flags = picture.shape().flags;
        let transform2_d = (shape_flags
            .intersects(OfficeArtShapeFlags::FLIP_HORIZONTAL | OfficeArtShapeFlags::FLIP_VERTICAL))
        .then(|| {
            Box::new(a::Transform2D {
                horizontal_flip: shape_flags
                    .contains(OfficeArtShapeFlags::FLIP_HORIZONTAL)
                    .then_some(true.into()),
                vertical_flip: shape_flags
                    .contains(OfficeArtShapeFlags::FLIP_VERTICAL)
                    .then_some(true.into()),
                ..Default::default()
            })
        });
        anchors.push(xdr::WorksheetDrawingChoice::TwoCellAnchor(Box::new(
            xdr::TwoCellAnchor {
                edit_as: Some(edit_as),
                from_marker: Box::new(xls_from_marker(coordinates)),
                to_marker: Box::new(xls_to_marker(coordinates)),
                two_cell_anchor_choice: Some(xdr::TwoCellAnchorChoice::Picture(Box::new(
                    xdr::Picture {
                        non_visual_picture_properties: Box::new(xdr::NonVisualPictureProperties {
                            non_visual_drawing_properties: Box::new(
                                xdr::NonVisualDrawingProperties {
                                    id: picture.shape().shape_id,
                                    name: format!("Legacy Picture {}", picture.shape().shape_id),
                                    ..Default::default()
                                },
                            ),
                            non_visual_picture_drawing_properties: Box::default(),
                        }),
                        blip_fill: Some(Box::new(xdr::BlipFill {
                            blip: Some(Box::new(a::Blip {
                                embed: Some(relationship_id),
                                ..Default::default()
                            })),
                            source_rectangle,
                            blip_fill_choice: Some(xdr::BlipFillChoice::Stretch(Box::new(
                                a::Stretch {
                                    fill_rectangle: Some(Default::default()),
                                    ..Default::default()
                                },
                            ))),
                            ..Default::default()
                        })),
                        shape_properties: Box::new(xdr::ShapeProperties {
                            transform2_d,
                            shape_properties_choice1: Some(
                                xdr::ShapePropertiesChoice::PresetGeometry(Box::new(
                                    a::PresetGeometry {
                                        preset: a::ShapeTypeValues::Rectangle,
                                        adjust_value_list: Some(Default::default()),
                                        ..Default::default()
                                    },
                                )),
                            ),
                            shape_properties_choice2: Some(xdr::ShapePropertiesChoice2::NoFill(
                                a::NoFill::default(),
                            )),
                            outline: Some(Box::new(a::Outline {
                                outline_choice1: Some(a::OutlineChoice::NoFill(
                                    a::NoFill::default(),
                                )),
                                ..Default::default()
                            })),
                            ..Default::default()
                        }),
                        ..Default::default()
                    },
                ))),
                client_data: xdr::ClientData::default(),
            },
        )));
        report.record(Disposition::Mapped);
    }
    let Some((drawings_part, relationship_id)) = drawing_part else {
        return Ok(None);
    };
    drawings_part.set_root_element(
        document,
        xdr::WorksheetDrawing {
            xmlns: vec![
                XmlNamespace::known(XmlKnownNamespace::A),
                XmlNamespace::known(XmlKnownNamespace::R),
            ],
            worksheet_drawing_choice: anchors,
        },
    )?;
    Ok(Some(x::Drawing {
        xmlns: vec![XmlNamespace::known(XmlKnownNamespace::R)],
        id: relationship_id,
    }))
}

fn add_xls_image_relationship(
    blip_identifier: u32,
    source: olecfsdk::office_art::OfficeArtImageRef<'_>,
    content_type: &'static str,
    host: &DrawingsPart,
    document: &mut SpreadsheetDocument,
    media: &mut XlsMediaState,
) -> Result<String> {
    let image_part = if let Some((_, part)) = media
        .parts
        .iter()
        .find(|(identifier, _)| *identifier == blip_identifier)
    {
        host.add_part(document, part.clone())?
    } else {
        let part = host.add_image_part(document, content_type)?;
        part.set_data(document, source.data.to_vec())?;
        media.parts.push((blip_identifier, part.clone()));
        part
    };
    Ok(host
        .get_id_of_part(document, &image_part)
        .expect("a newly related XLS image has a relationship ID")
        .to_owned())
}

fn xls_from_marker(coordinates: [u16; 8]) -> xdr::FromMarker {
    xdr::FromMarker {
        column_id: i32::from(coordinates[0]),
        column_offset: xls_column_offset(coordinates[1]),
        row_id: i32::from(coordinates[2]),
        row_offset: xls_row_offset(coordinates[3]),
    }
}

fn xls_to_marker(coordinates: [u16; 8]) -> xdr::ToMarker {
    xdr::ToMarker {
        column_id: i32::from(coordinates[4]),
        column_offset: xls_column_offset(coordinates[5]),
        row_id: i32::from(coordinates[6]),
        row_offset: xls_row_offset(coordinates[7]),
    }
}

fn xls_column_offset(value: u16) -> CoordinateValue {
    // BIFF client-anchor dx is 1/1024 of the host column. The current XLSX
    // vertical slice uses Excel's 64-pixel default column until dimensions
    // are mapped, so retain that same physical fraction in EMUs.
    CoordinateValue::Emu(i64::from(value) * 64 * 9_525 / 1_024)
}

fn xls_row_offset(value: u16) -> CoordinateValue {
    // BIFF client-anchor dy is 1/256 of the host row; 15 points is the Excel
    // default row height used by the current target worksheet.
    CoordinateValue::Emu(i64::from(value) * 15 * 12_700 / 256)
}

fn xls_picture_crop(
    picture: XlsPictureRef<'_>,
    formatting_loss: &mut bool,
) -> Option<a::SourceRectangle> {
    let crop = picture.crop();
    let values = [crop.left(), crop.top(), crop.right(), crop.bottom()];
    if values.iter().all(|value| *value == 0) {
        return None;
    }
    let converted = values.map(fixed_16_16_to_percentage);
    let [Some(left), Some(top), Some(right), Some(bottom)] = converted else {
        *formatting_loss = true;
        return None;
    };
    Some(a::SourceRectangle {
        left: Some(left),
        top: Some(top),
        right: Some(right),
        bottom: Some(bottom),
        ..Default::default()
    })
}

fn fixed_16_16_to_percentage(
    value: i32,
) -> Option<ooxmlsdk::simple_type::DrawingmlPercentageValue> {
    let scaled = i64::from(value) * 100_000;
    let rounded = if scaled < 0 {
        (scaled - 32_768) / 65_536
    } else {
        (scaled + 32_768) / 65_536
    };
    i32::try_from(rounded)
        .ok()
        .map(ooxmlsdk::simple_type::DrawingmlPercentageValue::Decimal)
}

const fn xls_image_content_type(format: OfficeArtImageFormat) -> Option<&'static str> {
    match format {
        OfficeArtImageFormat::Emf => Some("image/x-emf"),
        OfficeArtImageFormat::Wmf => Some("image/x-wmf"),
        OfficeArtImageFormat::Jpeg => Some("image/jpeg"),
        OfficeArtImageFormat::Png => Some("image/png"),
        OfficeArtImageFormat::Tiff => Some("image/tiff"),
        OfficeArtImageFormat::Pict | OfficeArtImageFormat::Dib => None,
    }
}

fn convert_comments(
    source: Vec<olecfsdk::xls::XlsCommentRef<'_>>,
    sheet_index: usize,
    options: ConversionOptions,
    report: &mut ConversionReport,
) -> Result<Option<Comments>> {
    if source.is_empty() {
        return Ok(None);
    }
    let mut author_values = Vec::<String>::new();
    let mut comments = Vec::with_capacity(source.len());
    for comment in source {
        let row = comment.note().row;
        let column = comment.note().column;
        let location = SourceLocation::XlsCell {
            workbook_index: 0,
            sheet_index,
            row,
            column,
        };
        unsupported(
            report,
            options,
            ConversionCode::CommentFormattingNotMapped,
            location,
        )?;
        let author_id = if let Some(index) = author_values
            .iter()
            .position(|author| author == &comment.author)
        {
            u32::try_from(index).map_err(|_| {
                olecfsdk::Error::Limit("XLS comment author index exceeds u32".into())
            })?
        } else {
            let index = u32::try_from(author_values.len()).map_err(|_| {
                olecfsdk::Error::Limit("XLS comment author count exceeds u32".into())
            })?;
            author_values.push(comment.author);
            index
        };
        comments.push(Comment {
            reference: cell_reference(row, column),
            author_id,
            comment_text: Box::new(CommentText {
                text: Some(Text(xstring(comment.content))),
                ..Default::default()
            }),
            ..Default::default()
        });
        report.record(Disposition::Mapped);
    }
    Ok(Some(Comments {
        xmlns: vec![XmlNamespace::known(XmlKnownNamespace::X)],
        authors: Authors {
            author: author_values
                .into_iter()
                .map(|value| Author(xstring(value)))
                .collect(),
        },
        comment_list: CommentList { comment: comments },
        ..Default::default()
    }))
}

fn has_unmapped_workbook_features(source: &XlsFile, view: &XlsWorkbookView<'_>) -> bool {
    !view.unresolved_sheets().is_empty()
        || !view.unlinked_substreams().is_empty()
        || !view.supporting_links().is_empty()
        || !view.external_sheets().is_empty()
        || !view.defined_names().is_empty()
        || !view.pivot_cache_definitions().is_empty()
        || !view.custom_views().is_empty()
        || !source.pivot_caches.is_empty()
        || source.revision_log.is_some()
        || source.user_names.is_some()
}

fn convert_shared_strings(
    view: &XlsWorkbookView<'_>,
    options: ConversionOptions,
    report: &mut ConversionReport,
) -> Result<Option<SharedStringTable>> {
    let Some(source) = view.shared_string_table()? else {
        return Ok(None);
    };
    let mut items = Vec::with_capacity(source.strings.len());
    for (string_index, source_string) in source.strings.iter().enumerate() {
        let string_index = u32::try_from(string_index)
            .map_err(|_| olecfsdk::Error::Limit("XLS SST index exceeds u32".into()))?;
        let source_location = SourceLocation::XlsSharedString {
            workbook_index: 0,
            string_index,
        };
        if !source_string.format_runs.is_empty()
            || !matches!(
                source_string.extension,
                olecfsdk::xls::SstExtensionData::None
            )
        {
            unsupported(
                report,
                options,
                ConversionCode::SharedStringRichTextNotMapped,
                source_location,
            )?;
        }
        let value = view
            .shared_string_value(string_index)?
            .expect("the SST index comes from the same shared string table");
        items.push(SharedStringItem {
            text: Some(Text(xstring(value))),
            ..Default::default()
        });
        report.record(Disposition::Mapped);
    }
    Ok(Some(SharedStringTable {
        count: Some(source.total_string_count),
        unique_count: Some(source.unique_string_count),
        shared_string_item: items,
        ..Default::default()
    }))
}

struct ConvertedStylesheet {
    root: x::Stylesheet,
    xf_unmapped: Vec<bool>,
}

fn convert_stylesheet(view: &XlsWorkbookView<'_>) -> Result<ConvertedStylesheet> {
    let source_fonts = view.fonts().collect::<Vec<_>>();
    let mut fonts = Vec::with_capacity(source_fonts.len());
    let mut font_unmapped = Vec::with_capacity(source_fonts.len());
    for source in source_fonts {
        let (font, unmapped) = convert_font(source, None)?;
        fonts.push(x::FontsChoice::Font(font));
        font_unmapped.push(unmapped);
    }
    if fonts.is_empty() {
        fonts.push(x::FontsChoice::Font(x::Font::default()));
        font_unmapped.push(false);
    }

    let xfs = view.xfs().collect::<Vec<_>>();
    let mut extensions = vec![None; xfs.len()];
    let mut duplicate_extension = vec![false; xfs.len()];
    for extension in view.xf_extensions() {
        let index = usize::from(extension.xf_index);
        if let Some(slot) = extensions.get_mut(index)
            && slot.replace(extension).is_some()
        {
            duplicate_extension[index] = true;
        }
    }
    let mut style_positions = vec![None; xfs.len()];
    let mut next_style = 0u32;
    for (index, xf) in xfs.iter().enumerate() {
        if xf.cell_flags & 0x0004 != 0 {
            style_positions[index] = Some(next_style);
            next_style = next_style
                .checked_add(1)
                .ok_or_else(|| olecfsdk::Error::Limit("XLS style XF count exceeds u32".into()))?;
        }
    }
    let mut cell_formats = Vec::with_capacity(xfs.len().max(1));
    let mut style_formats = Vec::with_capacity(xfs.len());
    let mut fills = Vec::with_capacity(xfs.len() + 2);
    fills.push(x::FillsChoice::Fill(Box::new(pattern_fill(
        x::PatternValues::None,
        0,
        0,
        None,
    ))));
    fills.push(x::FillsChoice::Fill(Box::new(pattern_fill(
        x::PatternValues::Gray125,
        0,
        0,
        None,
    ))));
    let mut borders = Vec::with_capacity(xfs.len().max(1));
    let mut xf_unmapped = Vec::with_capacity(xfs.len());
    for (index, xf) in xfs.iter().copied().enumerate() {
        let extension = extensions[index];
        let source_font_index = font_position(xf.font_index);
        let (target_font_id, mut unmapped) = match source_font_index {
            Some(source_font_index) => {
                let source_font_unmapped = font_unmapped
                    .get(source_font_index)
                    .copied()
                    .unwrap_or(true);
                if extension.is_some_and(has_font_extension) {
                    let source_font = view.font(xf.font_index);
                    if let Some(source_font) = source_font {
                        let (font, derived_unmapped) = convert_font(source_font, extension)?;
                        let target_font_id = u32::try_from(fonts.len()).map_err(|_| {
                            olecfsdk::Error::Limit("XLS target font count exceeds u32".into())
                        })?;
                        fonts.push(x::FontsChoice::Font(font));
                        (target_font_id, source_font_unmapped | derived_unmapped)
                    } else {
                        (0, true)
                    }
                } else {
                    (
                        u32::try_from(source_font_index).map_err(|_| {
                            olecfsdk::Error::Limit("XLS font index exceeds u32".into())
                        })?,
                        source_font_unmapped,
                    )
                }
            }
            None => (0, true),
        };
        unmapped |= xf_has_unmapped_base(xf, &style_positions)
            || duplicate_extension[index]
            || extension.is_some() != (xf.additional_border_color_flags & 0x0200_0000 != 0)
            || extension.is_some_and(extension_has_unmapped_property);
        let format = convert_xf(xf, index, &style_positions, target_font_id, extension);
        let pattern =
            u8::try_from((xf.additional_border_color_flags >> 26) & 0x3f).expect("six bits fit u8");
        let pattern_type = fill_pattern(pattern);
        if pattern_type.is_none() {
            unmapped = true;
        }
        fills.push(x::FillsChoice::Fill(Box::new(pattern_fill(
            pattern_type.unwrap_or_default(),
            u32::from(xf.fill_flags & 0x007f),
            u32::from((xf.fill_flags >> 7) & 0x007f),
            extension,
        ))));
        let (border, border_unmapped) = convert_border(xf, extension);
        borders.push(border);
        unmapped |= border_unmapped;
        if xf.cell_flags & 0x0004 != 0 {
            style_formats.push(convert_xf(
                xf,
                index,
                &style_positions,
                target_font_id,
                extension,
            ));
        }
        cell_formats.push(x::CellFormatsChoice::CellFormat(Box::new(format)));
        xf_unmapped.push(unmapped);
    }
    if cell_formats.is_empty() {
        cell_formats.push(x::CellFormatsChoice::CellFormat(Box::default()));
        xf_unmapped.push(false);
        borders.push(Default::default());
    }
    if style_formats.is_empty() {
        style_formats.push(Default::default());
    }

    let numbering_format = view
        .formats()
        .map(|format| {
            Ok(x::NumberingFormat {
                number_format_id: u32::from(format.format_index),
                format_code: String::try_from(&format.format_string)?,
                ..Default::default()
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let count = |value: usize, label: &str| {
        u32::try_from(value)
            .map_err(|_| olecfsdk::Error::Limit(format!("XLS {label} count exceeds u32")))
    };
    Ok(ConvertedStylesheet {
        root: x::Stylesheet {
            numbering_formats: (!numbering_format.is_empty()).then_some(x::NumberingFormats {
                count: Some(count(numbering_format.len(), "number format")?),
                numbering_format,
            }),
            fonts: Some(x::Fonts {
                count: Some(count(fonts.len(), "font")?),
                known_fonts: None,
                xml_children: fonts,
            }),
            fills: Some(x::Fills {
                count: Some(count(fills.len(), "fill")?),
                xml_children: fills,
            }),
            borders: Some(x::Borders {
                count: Some(count(borders.len(), "border")?),
                border: borders,
            }),
            cell_style_formats: Some(x::CellStyleFormats {
                count: Some(count(style_formats.len(), "style XF")?),
                cell_format: style_formats,
            }),
            cell_formats: Some(x::CellFormats {
                count: Some(count(cell_formats.len(), "cell XF")?),
                xml_children: cell_formats,
            }),
            cell_styles: Some(x::CellStyles {
                count: Some(1),
                cell_style: vec![x::CellStyle {
                    name: Some("Normal".into()),
                    format_id: 0,
                    builtin_id: Some(0),
                    ..Default::default()
                }],
            }),
            ..Default::default()
        },
        xf_unmapped,
    })
}

const fn font_position(index: u16) -> Option<usize> {
    match index {
        0..=3 => Some(index as usize),
        4 => None,
        _ => Some((index - 1) as usize),
    }
}

fn xf_has_unmapped_base(source: &XfRecord, style_positions: &[Option<u32>]) -> bool {
    let style = source.cell_flags & 0x0004 != 0;
    let parent = usize::from(source.cell_flags >> 4);
    (!style && parent != 0x0fff && style_positions.get(parent).copied().flatten().is_none())
        || source.indentation_flags & 0x0320 != 0
        || source.fill_flags & 0x8000 != 0
        || horizontal_alignment(source.alignment_flags & 0x0007).is_none()
            && source.alignment_flags & 0x0007 != 0
        || vertical_alignment((source.alignment_flags >> 4) & 0x0007).is_none()
}

fn has_font_extension(extension: &XfExtRecord) -> bool {
    extension.properties.iter().any(|property| {
        matches!(
            property.data,
            ExtPropertyData::FullColor {
                property_type: 0x000d,
                ..
            } | ExtPropertyData::FontScheme(_)
        )
    })
}

fn extension_font_scheme(extension: &XfExtRecord) -> Option<u16> {
    extension
        .properties
        .iter()
        .rev()
        .find_map(|property| match property.data {
            ExtPropertyData::FontScheme(ExtFontScheme::Byte(value)) => Some(u16::from(value)),
            ExtPropertyData::FontScheme(ExtFontScheme::Word(value)) => Some(value),
            _ => None,
        })
}

fn extension_indentation(extension: &XfExtRecord) -> Option<u16> {
    extension
        .properties
        .iter()
        .rev()
        .find_map(|property| match property.data {
            ExtPropertyData::Indentation(value) => Some(value),
            _ => None,
        })
}

fn extension_property_type(property: &ExtPropertyData) -> u16 {
    match property {
        ExtPropertyData::FullColor { property_type, .. }
        | ExtPropertyData::Unknown { property_type, .. } => *property_type,
        ExtPropertyData::Gradient { .. } => 0x0006,
        ExtPropertyData::FontScheme(_) => 0x000e,
        ExtPropertyData::Indentation(_) => 0x000f,
    }
}

fn extension_has_unmapped_property(extension: &XfExtRecord) -> bool {
    extension
        .properties
        .iter()
        .enumerate()
        .any(|(index, property)| {
            let property_type = extension_property_type(&property.data);
            extension.properties[..index]
                .iter()
                .any(|previous| extension_property_type(&previous.data) == property_type)
                || match &property.data {
                    ExtPropertyData::FullColor {
                        property_type:
                            0x0004 | 0x0005 | 0x0007 | 0x0008 | 0x0009 | 0x000a | 0x000b | 0x000d,
                        color,
                    } => convert_full_color(color).1,
                    ExtPropertyData::FontScheme(value) => !matches!(
                        value,
                        ExtFontScheme::Byte(0..=2) | ExtFontScheme::Word(0..=2)
                    ),
                    ExtPropertyData::Indentation(value) => *value > 250,
                    ExtPropertyData::FullColor { .. }
                    | ExtPropertyData::Gradient { .. }
                    | ExtPropertyData::Unknown { .. } => true,
                }
        })
}

fn extension_color(extension: &XfExtRecord, target_type: u16) -> Option<x::Color> {
    extension
        .properties
        .iter()
        .rev()
        .find_map(|property| match &property.data {
            ExtPropertyData::FullColor {
                property_type,
                color,
            } if *property_type == target_type => convert_full_color(color).0,
            _ => None,
        })
}

fn convert_full_color(source: &FullColorExt) -> (Option<x::Color>, bool) {
    let tint = if source.tint < 0 {
        f64::from(source.tint) / 32_768.0
    } else {
        f64::from(source.tint) / 32_767.0
    };
    let tint = (source.tint != 0).then_some(tint);
    let mut color = x::Color {
        tint,
        ..Default::default()
    };
    match source.color_type {
        0 => color.auto = Some(BooleanValue::True),
        1 => color.indexed = Some(source.color_value),
        2 => {
            let [red, green, blue, alpha] = source.color_value.to_le_bytes();
            color.rgb = Some(format!("{alpha:02X}{red:02X}{green:02X}{blue:02X}"));
        }
        3 => color.theme = Some(source.color_value),
        4 => return (None, source.color_value != 0 || source.tint != 0),
        _ => return (None, true),
    }
    (Some(color), false)
}

fn convert_font(source: &FontRecord, extension: Option<&XfExtRecord>) -> Result<(x::Font, bool)> {
    let mut choices = Vec::new();
    let mut unmapped = source.attributes.bits()
        & !(FontAttributes::ITALIC
            | FontAttributes::STRIKEOUT
            | FontAttributes::MAC_OUTLINE
            | FontAttributes::MAC_SHADOW
            | FontAttributes::CONDENSE
            | FontAttributes::EXTEND)
            .bits()
        & !0xff05
        != 0;
    if source.bold_weight >= 700 {
        choices.push(x::FontChoice::Bold(Default::default()));
    }
    unmapped |= !matches!(source.bold_weight, 400 | 700);
    if source.attributes.contains(FontAttributes::ITALIC) {
        choices.push(x::FontChoice::Italic(Default::default()));
    }
    if source.attributes.contains(FontAttributes::STRIKEOUT) {
        choices.push(x::FontChoice::Strike(Default::default()));
    }
    if source.attributes.contains(FontAttributes::MAC_OUTLINE) {
        choices.push(x::FontChoice::Outline(Default::default()));
    }
    if source.attributes.contains(FontAttributes::MAC_SHADOW) {
        choices.push(x::FontChoice::Shadow(Default::default()));
    }
    if source.attributes.contains(FontAttributes::CONDENSE) {
        choices.push(x::FontChoice::Condense(x::Condense {
            val: Some(BooleanValue::True),
        }));
    }
    if source.attributes.contains(FontAttributes::EXTEND) {
        choices.push(x::FontChoice::Extend(x::Extend {
            val: Some(BooleanValue::True),
        }));
    }
    if let Some(value) = font_underline(source.underline) {
        choices.push(x::FontChoice::Underline(x::Underline { val: Some(value) }));
    } else if source.underline != 0 {
        unmapped = true;
    }
    if let Some(value) = font_vertical_alignment(source.escapement) {
        choices.push(x::FontChoice::VerticalTextAlignment(
            x::VerticalTextAlignment { val: value },
        ));
    } else if source.escapement != 0 {
        unmapped = true;
    }
    choices.push(x::FontChoice::FontSize(x::FontSize {
        val: f64::from(source.height_twips) / 20.0,
    }));
    let extended_color = extension.and_then(|extension| {
        extension.properties.iter().rev().find_map(|property| {
            let ExtPropertyData::FullColor {
                property_type: 0x000d,
                color,
            } = &property.data
            else {
                return None;
            };
            Some(color)
        })
    });
    let color = if let Some(color) = extended_color {
        let (color, color_unmapped) = convert_full_color(color);
        unmapped |= color_unmapped;
        color.unwrap_or_default()
    } else {
        x::Color {
            auto: (source.color_index == 0x7fff).then_some(BooleanValue::True),
            indexed: (source.color_index != 0x7fff).then_some(u32::from(source.color_index)),
            ..Default::default()
        }
    };
    choices.push(x::FontChoice::Color(color));
    let name = String::try_from(&source.name)?;
    if name.is_empty() {
        unmapped = true;
    } else {
        choices.push(x::FontChoice::FontName(x::FontName { val: name }));
    }
    if source.family <= 5 {
        choices.push(x::FontChoice::FontFamilyNumbering(x::FontFamilyNumbering {
            val: i32::from(source.family),
        }));
    } else {
        unmapped = true;
    }
    choices.push(x::FontChoice::FontCharSet(x::FontCharSet {
        val: i32::from(source.charset),
    }));
    if let Some(scheme) = extension.and_then(extension_font_scheme) {
        let scheme = match scheme {
            0 => x::FontSchemeValues::None,
            1 => x::FontSchemeValues::Major,
            2 => x::FontSchemeValues::Minor,
            _ => {
                unmapped = true;
                x::FontSchemeValues::None
            }
        };
        choices.push(x::FontChoice::FontScheme(x::FontScheme { val: scheme }));
    }
    Ok((
        x::Font {
            font_choice: choices,
        },
        unmapped,
    ))
}

fn convert_xf(
    source: &XfRecord,
    xf_index: usize,
    style_positions: &[Option<u32>],
    target_font_id: u32,
    extension: Option<&XfExtRecord>,
) -> x::CellFormat {
    let style = source.cell_flags & 0x0004 != 0;
    let parent = usize::from(source.cell_flags >> 4);
    let format_id = (!style)
        .then(|| style_positions.get(parent).copied().flatten())
        .flatten();
    let horizontal = horizontal_alignment(source.alignment_flags & 0x0007);
    let vertical = vertical_alignment((source.alignment_flags >> 4) & 0x0007);
    let alignment = x::Alignment {
        horizontal,
        vertical,
        text_rotation: Some(u32::from(source.alignment_flags >> 8)),
        wrap_text: Some(BooleanValue::from_bool(
            source.alignment_flags & 0x0008 != 0,
        )),
        indent: Some(u32::from(
            extension
                .and_then(extension_indentation)
                .unwrap_or(source.indentation_flags & 0x000f),
        )),
        justify_last_line: Some(BooleanValue::from_bool(
            source.alignment_flags & 0x0080 != 0,
        )),
        shrink_to_fit: Some(BooleanValue::from_bool(
            source.indentation_flags & 0x0010 != 0,
        )),
        reading_order: Some(u32::from((source.indentation_flags >> 6) & 0x0003)),
        ..Default::default()
    };
    let fill_id = u32::try_from(xf_index + 2).unwrap_or(u32::MAX);
    let border_id = u32::try_from(xf_index).unwrap_or(u32::MAX);
    x::CellFormat {
        number_format_id: Some(u32::from(source.number_format_index)),
        font_id: Some(target_font_id),
        fill_id: Some(fill_id),
        border_id: Some(border_id),
        format_id: format_id.or(Some(0)),
        quote_prefix: Some(BooleanValue::from_bool(source.cell_flags & 0x0008 != 0)),
        pivot_button: Some(BooleanValue::from_bool(source.fill_flags & 0x4000 != 0)),
        apply_number_format: Some(BooleanValue::from_bool(
            source.indentation_flags & 0x0400 != 0,
        )),
        apply_font: Some(BooleanValue::from_bool(
            source.indentation_flags & 0x0800 != 0,
        )),
        apply_alignment: Some(BooleanValue::from_bool(
            source.indentation_flags & 0x1000 != 0,
        )),
        apply_border: Some(BooleanValue::from_bool(
            source.indentation_flags & 0x2000 != 0,
        )),
        apply_fill: Some(BooleanValue::from_bool(
            source.indentation_flags & 0x4000 != 0,
        )),
        apply_protection: Some(BooleanValue::from_bool(
            source.indentation_flags & 0x8000 != 0,
        )),
        alignment: Some(alignment),
        protection: Some(x::Protection {
            locked: Some(BooleanValue::from_bool(source.cell_flags & 0x0001 != 0)),
            hidden: Some(BooleanValue::from_bool(source.cell_flags & 0x0002 != 0)),
        }),
        ..Default::default()
    }
}

fn pattern_fill(
    pattern: x::PatternValues,
    foreground: u32,
    background: u32,
    extension: Option<&XfExtRecord>,
) -> x::Fill {
    let extended_foreground = extension.and_then(|extension| extension_color(extension, 0x0004));
    let extended_background = extension.and_then(|extension| extension_color(extension, 0x0005));
    let x::Color {
        auto: foreground_auto,
        indexed: foreground_indexed,
        rgb: foreground_rgb,
        theme: foreground_theme,
        tint: foreground_tint,
    } = extended_foreground.unwrap_or(x::Color {
        indexed: Some(foreground),
        ..Default::default()
    });
    let x::Color {
        auto: background_auto,
        indexed: background_indexed,
        rgb: background_rgb,
        theme: background_theme,
        tint: background_tint,
    } = extended_background.unwrap_or(x::Color {
        indexed: Some(background),
        ..Default::default()
    });
    x::Fill {
        fill_choice: Some(x::FillChoice::PatternFill(Box::new(x::PatternFill {
            pattern_type: Some(pattern),
            foreground_color: Some(x::ForegroundColor {
                auto: foreground_auto,
                indexed: foreground_indexed,
                rgb: foreground_rgb,
                theme: foreground_theme,
                tint: foreground_tint,
            }),
            background_color: Some(x::BackgroundColor {
                auto: background_auto,
                indexed: background_indexed,
                rgb: background_rgb,
                theme: background_theme,
                tint: background_tint,
            }),
        }))),
    }
}

fn convert_border(source: &XfRecord, extension: Option<&XfExtRecord>) -> (x::Border, bool) {
    let style = |value| border_style(value);
    let left_style = style(source.border_style_flags & 0x000f);
    let right_style = style((source.border_style_flags >> 4) & 0x000f);
    let top_style = style((source.border_style_flags >> 8) & 0x000f);
    let bottom_style = style((source.border_style_flags >> 12) & 0x000f);
    let diagonal_style = style(
        u16::try_from((source.additional_border_color_flags >> 21) & 0x000f)
            .expect("four bits fit u16"),
    );
    let diagonal = (source.border_color_flags >> 14) & 0x0003;
    let border = x::Border {
        diagonal_down: Some(BooleanValue::from_bool(diagonal & 1 != 0)),
        diagonal_up: Some(BooleanValue::from_bool(diagonal & 2 != 0)),
        left_border: Some(Box::new(x::LeftBorder {
            style: left_style,
            color: extension
                .and_then(|extension| extension_color(extension, 0x0009))
                .or_else(|| border_color(u32::from(source.border_color_flags & 0x007f))),
        })),
        right_border: Some(Box::new(x::RightBorder {
            style: right_style,
            color: extension
                .and_then(|extension| extension_color(extension, 0x000a))
                .or_else(|| border_color(u32::from((source.border_color_flags >> 7) & 0x007f))),
        })),
        top_border: Some(Box::new(x::TopBorder {
            style: top_style,
            color: extension
                .and_then(|extension| extension_color(extension, 0x0007))
                .or_else(|| border_color(source.additional_border_color_flags & 0x007f)),
        })),
        bottom_border: Some(Box::new(x::BottomBorder {
            style: bottom_style,
            color: extension
                .and_then(|extension| extension_color(extension, 0x0008))
                .or_else(|| border_color((source.additional_border_color_flags >> 7) & 0x007f)),
        })),
        diagonal_border: Some(Box::new(x::DiagonalBorder {
            style: diagonal_style,
            color: extension
                .and_then(|extension| extension_color(extension, 0x000b))
                .or_else(|| border_color((source.additional_border_color_flags >> 14) & 0x007f)),
        })),
        ..Default::default()
    };
    (
        border,
        [
            left_style,
            right_style,
            top_style,
            bottom_style,
            diagonal_style,
        ]
        .iter()
        .enumerate()
        .any(|(index, value)| {
            value.is_none()
                && [
                    source.border_style_flags & 0x000f,
                    (source.border_style_flags >> 4) & 0x000f,
                    (source.border_style_flags >> 8) & 0x000f,
                    (source.border_style_flags >> 12) & 0x000f,
                    ((source.additional_border_color_flags >> 21) & 0x000f) as u16,
                ][index]
                    != 0
        }),
    )
}

fn border_color(index: u32) -> Option<x::Color> {
    (index != 0).then_some(x::Color {
        indexed: Some(index),
        ..Default::default()
    })
}

const fn border_style(value: u16) -> Option<x::BorderStyleValues> {
    Some(match value {
        0 => return None,
        1 => x::BorderStyleValues::Thin,
        2 => x::BorderStyleValues::Medium,
        3 => x::BorderStyleValues::Dashed,
        4 => x::BorderStyleValues::Dotted,
        5 => x::BorderStyleValues::Thick,
        6 => x::BorderStyleValues::Double,
        7 => x::BorderStyleValues::Hair,
        8 => x::BorderStyleValues::MediumDashed,
        9 => x::BorderStyleValues::DashDot,
        10 => x::BorderStyleValues::MediumDashDot,
        11 => x::BorderStyleValues::DashDotDot,
        12 => x::BorderStyleValues::MediumDashDotDot,
        13 => x::BorderStyleValues::SlantDashDot,
        _ => return None,
    })
}

const fn fill_pattern(value: u8) -> Option<x::PatternValues> {
    Some(match value {
        0 => x::PatternValues::None,
        1 => x::PatternValues::Solid,
        2 => x::PatternValues::MediumGray,
        3 => x::PatternValues::DarkGray,
        4 => x::PatternValues::LightGray,
        5 => x::PatternValues::DarkHorizontal,
        6 => x::PatternValues::DarkVertical,
        7 => x::PatternValues::DarkDown,
        8 => x::PatternValues::DarkUp,
        9 => x::PatternValues::DarkGrid,
        10 => x::PatternValues::DarkTrellis,
        11 => x::PatternValues::LightHorizontal,
        12 => x::PatternValues::LightVertical,
        13 => x::PatternValues::LightDown,
        14 => x::PatternValues::LightUp,
        15 => x::PatternValues::LightGrid,
        16 => x::PatternValues::LightTrellis,
        17 => x::PatternValues::Gray125,
        18 => x::PatternValues::Gray0625,
        _ => return None,
    })
}

const fn horizontal_alignment(value: u16) -> Option<x::HorizontalAlignmentValues> {
    Some(match value {
        0 => return None,
        1 => x::HorizontalAlignmentValues::Left,
        2 => x::HorizontalAlignmentValues::Center,
        3 => x::HorizontalAlignmentValues::Right,
        4 => x::HorizontalAlignmentValues::Fill,
        5 => x::HorizontalAlignmentValues::Justify,
        6 => x::HorizontalAlignmentValues::CenterContinuous,
        7 => x::HorizontalAlignmentValues::Distributed,
        _ => return None,
    })
}

const fn vertical_alignment(value: u16) -> Option<x::VerticalAlignmentValues> {
    Some(match value {
        0 => x::VerticalAlignmentValues::Top,
        1 => x::VerticalAlignmentValues::Center,
        2 => x::VerticalAlignmentValues::Bottom,
        3 => x::VerticalAlignmentValues::Justify,
        4 => x::VerticalAlignmentValues::Distributed,
        _ => return None,
    })
}

const fn font_underline(value: u8) -> Option<x::UnderlineValues> {
    Some(match value {
        0 => return None,
        1 => x::UnderlineValues::Single,
        2 => x::UnderlineValues::Double,
        0x21 => x::UnderlineValues::SingleAccounting,
        0x22 => x::UnderlineValues::DoubleAccounting,
        _ => return None,
    })
}

const fn font_vertical_alignment(value: u16) -> Option<x::VerticalAlignmentRunValues> {
    Some(match value {
        0 => return None,
        1 => x::VerticalAlignmentRunValues::Superscript,
        2 => x::VerticalAlignmentRunValues::Subscript,
        _ => return None,
    })
}

fn render_cell_formula(formula: XlsFormulaRef<'_>, row: u16, column: u16) -> Option<CellFormula> {
    if formula.formula().flags & 0x0020 != 0 {
        return None;
    }
    let (tokens, shared) = match formula.definition() {
        XlsFormulaDefinitionRef::Inline(tokens) => (tokens, false),
        XlsFormulaDefinitionRef::Shared(formula) => (&formula.tokens, true),
        XlsFormulaDefinitionRef::Array(_)
        | XlsFormulaDefinitionRef::Table(_)
        | XlsFormulaDefinitionRef::UnresolvedExp { .. }
        | XlsFormulaDefinitionRef::UnresolvedTable { .. } => return None,
    };
    if !tokens.rgcb_tail.is_empty() {
        return None;
    }
    Some(CellFormula {
        calculate_cell: (formula.formula().flags & 0x0001 != 0).then_some(BooleanValue::True),
        xml_content: render_formula_tokens(&tokens.rgce, row, column, shared),
        ..Default::default()
    })
    .filter(|formula| formula.xml_content.is_some())
}

fn render_formula_tokens(
    source: &FormulaTokenStream,
    row: u16,
    column: u16,
    shared: bool,
) -> Option<String> {
    if !source.unparsed_tail.is_empty()
        || source.missing_extra_count() != 0
        || source.nonconforming_token_count() != 0
    {
        return None;
    }
    let mut stack = Vec::with_capacity(source.tokens.len());
    for token in &source.tokens {
        match &token.data {
            FormulaTokenData::Operator(operator) => {
                apply_formula_operator(&mut stack, *operator)?;
            }
            FormulaTokenData::String { flags, characters } if flags & !1 == 0 => {
                let value = formula_string(characters)?;
                stack.push(format!("\"{}\"", value.replace('"', "\"\"")));
            }
            FormulaTokenData::Error(value) => stack.push(formula_error(*value)?.into()),
            FormulaTokenData::Boolean(0) => stack.push("FALSE".into()),
            FormulaTokenData::Boolean(1) => stack.push("TRUE".into()),
            FormulaTokenData::Integer(value) => stack.push(value.to_string()),
            FormulaTokenData::NumberBits(bits) => {
                let value = f64::from_bits(*bits);
                if !value.is_finite() {
                    return None;
                }
                stack.push(value.to_string());
            }
            FormulaTokenData::Reference {
                row: target_row,
                column: target_column,
            } => stack.push(formula_reference(
                *target_row,
                *target_column,
                row,
                column,
                false,
            )?),
            FormulaTokenData::RelativeReference {
                row: target_row,
                column: target_column,
            } => stack.push(formula_reference(
                *target_row,
                *target_column,
                row,
                column,
                shared,
            )?),
            FormulaTokenData::Area {
                first_row,
                last_row,
                first_column,
                last_column,
            } => stack.push(formula_area(
                (*first_row, *first_column),
                (*last_row, *last_column),
                row,
                column,
                false,
            )?),
            FormulaTokenData::RelativeArea {
                first_row,
                last_row,
                first_column,
                last_column,
            } => stack.push(formula_area(
                (*first_row, *first_column),
                (*last_row, *last_column),
                row,
                column,
                shared,
            )?),
            FormulaTokenData::ReferenceError { .. } | FormulaTokenData::AreaError { .. } => {
                stack.push("#REF!".into());
            }
            FormulaTokenData::Function { function_index } => {
                let (name, argument_count) = fixed_function(*function_index)?;
                apply_formula_function(&mut stack, name, argument_count)?;
            }
            FormulaTokenData::FunctionVar {
                argument_count,
                function_index,
            } if function_index & 0x8000 == 0 => {
                apply_formula_function(
                    &mut stack,
                    formula_function_name(*function_index)?,
                    usize::from(*argument_count),
                )?;
            }
            FormulaTokenData::Attribute { options, .. } if options & 0x1e == 0x10 => {
                apply_formula_function(&mut stack, "SUM", 1)?;
            }
            FormulaTokenData::Attribute { options, .. } if options & 0x1e == 0 => {}
            FormulaTokenData::MemArea { .. }
            | FormulaTokenData::MemNoMem { .. }
            | FormulaTokenData::MemFunction { .. } => {}
            FormulaTokenData::UnknownZero
            | FormulaTokenData::Exp { .. }
            | FormulaTokenData::Table { .. }
            | FormulaTokenData::String { .. }
            | FormulaTokenData::PivotName { .. }
            | FormulaTokenData::NaturalLanguage { .. }
            | FormulaTokenData::Attribute { .. }
            | FormulaTokenData::Boolean(_)
            | FormulaTokenData::Array { .. }
            | FormulaTokenData::Name { .. }
            | FormulaTokenData::MemError { .. }
            | FormulaTokenData::ExternalName { .. }
            | FormulaTokenData::Reference3d { .. }
            | FormulaTokenData::Area3d { .. }
            | FormulaTokenData::DeletedReference3d { .. }
            | FormulaTokenData::DeletedArea3d { .. }
            | FormulaTokenData::FunctionVar { .. } => return None,
        }
    }
    let mut expression = stack.pop()?;
    if !stack.is_empty() {
        return None;
    }
    if expression.starts_with('(') && expression.ends_with(')') {
        expression.remove(0);
        expression.pop();
    }
    Some(expression)
}

fn apply_formula_operator(stack: &mut Vec<String>, operator: FormulaOperator) -> Option<()> {
    if operator == FormulaOperator::MissingArgument {
        stack.push(String::new());
        return Some(());
    }
    let binary = match operator {
        FormulaOperator::Add => Some("+"),
        FormulaOperator::Subtract => Some("-"),
        FormulaOperator::Multiply => Some("*"),
        FormulaOperator::Divide => Some("/"),
        FormulaOperator::Power => Some("^"),
        FormulaOperator::Concat => Some("&"),
        FormulaOperator::LessThan => Some("<"),
        FormulaOperator::LessEqual => Some("<="),
        FormulaOperator::Equal => Some("="),
        FormulaOperator::GreaterEqual => Some(">="),
        FormulaOperator::GreaterThan => Some(">"),
        FormulaOperator::NotEqual => Some("<>"),
        FormulaOperator::Intersection => Some(" "),
        FormulaOperator::Union => Some(","),
        FormulaOperator::Range => Some(":"),
        _ => None,
    };
    if let Some(operator) = binary {
        let right = stack.pop()?;
        let left = stack.pop()?;
        stack.push(format!("({left}{operator}{right})"));
        return Some(());
    }
    let value = stack.pop()?;
    stack.push(match operator {
        FormulaOperator::UnaryPlus => format!("(+{value})"),
        FormulaOperator::UnaryMinus => format!("(-{value})"),
        FormulaOperator::Percent => format!("({value}%)"),
        FormulaOperator::Parenthesis => format!("({value})"),
        FormulaOperator::MissingArgument => unreachable!("handled before popping an operand"),
        _ => return None,
    });
    Some(())
}

fn apply_formula_function(
    stack: &mut Vec<String>,
    name: &str,
    argument_count: usize,
) -> Option<()> {
    let first = stack.len().checked_sub(argument_count)?;
    let arguments = stack.drain(first..).collect::<Vec<_>>().join(",");
    stack.push(format!("{name}({arguments})"));
    Some(())
}

fn formula_string(value: &XlStringCharacters) -> Option<String> {
    match value {
        XlStringCharacters::Compressed(value) => {
            Some(value.iter().copied().map(char::from).collect())
        }
        XlStringCharacters::Unicode(value) => String::from_utf16(value).ok(),
    }
}

fn formula_reference(
    source_row: u16,
    source_column: u16,
    formula_row: u16,
    formula_column: u16,
    offsets: bool,
) -> Option<String> {
    let column_relative = source_column & 0x4000 != 0;
    let row_relative = source_column & 0x8000 != 0;
    let raw_column = source_column & 0x3fff;
    let row = if offsets && row_relative {
        formula_row.wrapping_add(source_row)
    } else {
        source_row
    };
    let column = if offsets && column_relative {
        let offset = i32::from(((raw_column << 2) as i16) >> 2);
        u16::try_from((i32::from(formula_column) + offset).rem_euclid(256)).ok()?
    } else {
        raw_column
    };
    if column > 255 {
        return None;
    }
    let mut reference = String::new();
    if !column_relative {
        reference.push('$');
    }
    append_column_name(&mut reference, column);
    if !row_relative {
        reference.push('$');
    }
    reference.push_str(&(u32::from(row) + 1).to_string());
    Some(reference)
}

fn formula_area(
    first: (u16, u16),
    last: (u16, u16),
    formula_row: u16,
    formula_column: u16,
    offsets: bool,
) -> Option<String> {
    Some(format!(
        "{}:{}",
        formula_reference(first.0, first.1, formula_row, formula_column, offsets)?,
        formula_reference(last.0, last.1, formula_row, formula_column, offsets)?
    ))
}

fn append_column_name(target: &mut String, column: u16) {
    let mut value = u32::from(column) + 1;
    let mut letters = [0u8; 3];
    let mut start = letters.len();
    while value != 0 {
        value -= 1;
        start -= 1;
        letters[start] = b'A' + u8::try_from(value % 26).expect("modulo 26 fits u8");
        value /= 26;
    }
    target.push_str(std::str::from_utf8(&letters[start..]).expect("ASCII column name"));
}

const fn formula_error(value: u8) -> Option<&'static str> {
    Some(match value {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0f => "#VALUE!",
        0x17 => "#REF!",
        0x1d => "#NAME?",
        0x24 => "#NUM!",
        0x2a => "#N/A",
        0x2b => "#GETTING_DATA",
        _ => return None,
    })
}

const fn fixed_function(index: u16) -> Option<(&'static str, usize)> {
    Some(match index {
        2 => ("ISNA", 1),
        3 => ("ISERROR", 1),
        8 => ("ROW", 0),
        9 => ("COLUMN", 0),
        10 => ("NA", 0),
        15 => ("SIN", 1),
        16 => ("COS", 1),
        17 => ("TAN", 1),
        18 => ("ATAN", 1),
        19 => ("PI", 0),
        20 => ("SQRT", 1),
        21 => ("EXP", 1),
        22 => ("LN", 1),
        23 => ("LOG10", 1),
        24 => ("ABS", 1),
        25 => ("INT", 1),
        26 => ("SIGN", 1),
        32 => ("LEN", 1),
        33 => ("VALUE", 1),
        34 => ("TRUE", 0),
        35 => ("FALSE", 0),
        38 => ("NOT", 1),
        39 => ("MOD", 2),
        _ => return None,
    })
}

const fn formula_function_name(index: u16) -> Option<&'static str> {
    Some(match index {
        0 => "COUNT",
        1 => "IF",
        4 => "SUM",
        5 => "AVERAGE",
        6 => "MIN",
        7 => "MAX",
        11 => "NPV",
        12 => "STDEV",
        13 => "DOLLAR",
        14 => "FIXED",
        27 => "ROUND",
        28 => "LOOKUP",
        29 => "INDEX",
        30 => "REPT",
        31 => "MID",
        36 => "AND",
        37 => "OR",
        46 => "VAR",
        48 => "TEXT",
        56 => "PV",
        57 => "FV",
        58 => "NPER",
        59 => "PMT",
        60 => "RATE",
        61 => "MIRR",
        62 => "IRR",
        63 => "RAND",
        64 => "MATCH",
        65 => "DATE",
        66 => "TIME",
        67 => "DAY",
        68 => "MONTH",
        69 => "YEAR",
        70 => "WEEKDAY",
        71 => "HOUR",
        72 => "MINUTE",
        73 => "SECOND",
        74 => "NOW",
        97 => "ATAN2",
        98 => "ASIN",
        99 => "ACOS",
        100 => "CHOOSE",
        101 => "HLOOKUP",
        102 => "VLOOKUP",
        109 => "LOG",
        111 => "CHAR",
        112 => "LOWER",
        113 => "UPPER",
        114 => "PROPER",
        115 => "LEFT",
        116 => "RIGHT",
        117 => "EXACT",
        118 => "TRIM",
        119 => "REPLACE",
        120 => "SUBSTITUTE",
        121 => "CODE",
        124 => "FIND",
        125 => "CELL",
        126 => "ISERR",
        127 => "ISTEXT",
        128 => "ISNUMBER",
        129 => "ISBLANK",
        130 => "T",
        131 => "N",
        148 => "INDIRECT",
        162 => "CLEAN",
        163 => "MDETERM",
        164 => "MINVERSE",
        165 => "MMULT",
        169 => "COUNTA",
        183 => "PRODUCT",
        184 => "FACT",
        190 => "ISNONTEXT",
        193 => "STDEVP",
        194 => "VARP",
        197 => "TRUNC",
        212 => "ROUNDUP",
        213 => "ROUNDDOWN",
        216 => "RANK",
        220 => "DAYS360",
        221 => "TODAY",
        228 => "SUMPRODUCT",
        269 => "AVEDEV",
        276 => "COMBIN",
        279 => "EVEN",
        298 => "ODD",
        300 => "POISSON",
        303 => "SUMXMY2",
        318 => "DEVSQ",
        319 => "GEOMEAN",
        320 => "HARMEAN",
        321 => "SUMSQ",
        325 => "LARGE",
        326 => "SMALL",
        327 => "QUARTILE",
        328 => "PERCENTILE",
        329 => "PERCENTRANK",
        330 => "MODE",
        331 => "TRIMMEAN",
        336 => "CONCATENATE",
        337 => "POWER",
        342 => "RADIANS",
        343 => "DEGREES",
        344 => "SUBTOTAL",
        345 => "SUMIF",
        346 => "COUNTIF",
        347 => "COUNTBLANK",
        _ => return None,
    })
}

fn convert_cell(
    view: &XlsWorkbookView<'_>,
    index: &olecfsdk::xls::XlsSparseCellIndex<'_>,
    source: olecfsdk::xls::XlsCellRef<'_>,
    sheet_index: usize,
    options: ConversionOptions,
    report: &mut ConversionReport,
    xf_unmapped: &[bool],
) -> Result<Cell> {
    let header = source.cell();
    let source_location = SourceLocation::XlsCell {
        workbook_index: 0,
        sheet_index,
        row: header.row,
        column: header.column,
    };
    if xf_unmapped
        .get(usize::from(header.format_index))
        .copied()
        .unwrap_or(true)
    {
        unsupported(
            report,
            options,
            ConversionCode::CellFormattingNotMapped,
            source_location,
        )?;
    }
    let cell_formula = if matches!(source.value(), XlsCellValueRef::Formula(_)) {
        index
            .resolve_cell_formula(source)?
            .and_then(|formula| render_cell_formula(formula, header.row, header.column))
    } else {
        None
    };
    if matches!(
        source.value(),
        XlsCellValueRef::Formula(_) | XlsCellValueRef::Formula4Compatibility(_)
    ) && cell_formula.is_none()
    {
        unsupported(
            report,
            options,
            ConversionCode::FormulaNotMapped,
            source_location,
        )?;
    }

    let (data_type, value) = if let Some(label) = source.label_sst() {
        (
            Some(CellValues::SharedString),
            Some(label.shared_string_index.to_string()),
        )
    } else {
        let value = view.resolve_cell_value(index, source)?;
        convert_cell_value(value, source_location, options, report)?
    };
    report.record(Disposition::Mapped);
    Ok(Cell {
        cell_reference: Some(cell_reference(header.row, header.column)),
        style_index: Some(u32::from(header.format_index)),
        data_type,
        cell_formula,
        cell_value: value.map(|value| CellValue(xstring(value))),
        ..Default::default()
    })
}

fn convert_cell_value(
    value: XlsCellValue,
    source: SourceLocation,
    options: ConversionOptions,
    report: &mut ConversionReport,
) -> Result<(Option<CellValues>, Option<String>)> {
    match value {
        XlsCellValue::Blank => Ok((None, None)),
        XlsCellValue::Number(value) => Ok((Some(CellValues::Number), Some(value.to_string()))),
        XlsCellValue::Boolean(value) => Ok((
            Some(CellValues::Boolean),
            Some(if value { "1" } else { "0" }.to_owned()),
        )),
        XlsCellValue::Error(value) => Ok((
            Some(CellValues::Error),
            Some(cell_error_value(value).to_owned()),
        )),
        XlsCellValue::String(value) => Ok((Some(CellValues::String), Some(value))),
        XlsCellValue::Formula(value) => convert_formula_cache(value),
        XlsCellValue::CompatibilityBoolErr { .. } => {
            unsupported(
                report,
                options,
                ConversionCode::CompatibilityCellValue,
                source,
            )?;
            Ok((None, None))
        }
    }
}

fn convert_formula_cache(
    value: XlsFormulaCachedValue,
) -> Result<(Option<CellValues>, Option<String>)> {
    Ok(match value {
        XlsFormulaCachedValue::Number(value) => (Some(CellValues::Number), Some(value.to_string())),
        XlsFormulaCachedValue::String(value) => (Some(CellValues::String), Some(value)),
        XlsFormulaCachedValue::Boolean(value) => (
            Some(CellValues::Boolean),
            Some(if value { "1" } else { "0" }.to_owned()),
        ),
        XlsFormulaCachedValue::Error(value) => (
            Some(CellValues::Error),
            Some(cell_error_value(value).to_owned()),
        ),
        XlsFormulaCachedValue::Empty => (None, None),
    })
}

fn convert_sheet_state(
    state: u8,
    source: SourceLocation,
    options: ConversionOptions,
    report: &mut ConversionReport,
) -> Result<Option<SheetStateValues>> {
    if state & !0x03 != 0 || state & 0x03 == 3 {
        unsupported(report, options, ConversionCode::SheetStateNotMapped, source)?;
        return Ok(None);
    }
    Ok(match state & 0x03 {
        0 => None,
        1 => Some(SheetStateValues::Hidden),
        2 => Some(SheetStateValues::VeryHidden),
        _ => unreachable!(),
    })
}

fn cell_reference(row: u16, column: u16) -> String {
    let mut column = u32::from(column) + 1;
    let mut reversed = [0_u8; 4];
    let mut len = 0;
    while column != 0 {
        column -= 1;
        reversed[len] = b'A' + (column % 26) as u8;
        len += 1;
        column /= 26;
    }
    let mut value = String::with_capacity(len + 5);
    value.extend(reversed[..len].iter().rev().map(|value| char::from(*value)));
    value.push_str(&(u32::from(row) + 1).to_string());
    value
}

fn cell_range_reference(
    first_row: u16,
    first_column: u16,
    last_row: u16,
    last_column: u16,
) -> String {
    let first = cell_reference(first_row, first_column);
    let last = cell_reference(last_row, last_column);
    let mut value = String::with_capacity(first.len() + last.len() + 1);
    value.push_str(&first);
    value.push(':');
    value.push_str(&last);
    value
}

fn xstring(value: String) -> XstringType {
    let preserve = value.starts_with(char::is_whitespace)
        || value.ends_with(char::is_whitespace)
        || value.contains("  ");
    XstringType {
        space: preserve.then_some(SpaceProcessingModeValues::Preserve),
        xml_content: Some(value),
    }
}

const fn cell_error_value(value: CellErrorCode) -> &'static str {
    match value {
        CellErrorCode::Null => "#NULL!",
        CellErrorCode::DivisionByZero => "#DIV/0!",
        CellErrorCode::Value => "#VALUE!",
        CellErrorCode::Reference => "#REF!",
        CellErrorCode::Name => "#NAME?",
        CellErrorCode::Number => "#NUM!",
        CellErrorCode::NotAvailable => "#N/A",
        CellErrorCode::GettingData => "#GETTING_DATA",
    }
}

fn unsupported(
    report: &mut ConversionReport,
    options: ConversionOptions,
    code: ConversionCode,
    source: SourceLocation,
) -> Result<()> {
    match options.unsupported {
        LossPolicy::Reject => Err(Error::Unsupported {
            code,
            location: source,
        }),
        LossPolicy::Report => {
            report.issue(Disposition::Unsupported, code, source);
            Ok(())
        }
    }
}
