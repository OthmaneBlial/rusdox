use std::io::{Cursor, Read, Write};
use std::path::Path;

use rusdox::config::RusdoxConfig;
use rusdox::studio::Studio;
use rusdox::{
    Border, BorderStyle, Document, DocumentMode, HeaderFooter, PageNumberFormat, PageNumbering,
    PageSetup, Paragraph, ParagraphAlignment, ParagraphList, ParagraphStyle,
    ParagraphStyleProperties, Run, RunStyle, RunStyleProperties, Stylesheet, Table, TableBorders,
    TableCell, TableRow, TableStyle, TableStyleProperties, UnderlineStyle, Visual, VisualKind,
};
use tempfile::tempdir;
use zip::write::SimpleFileOptions;
use zip::ZipArchive;
use zip::{CompressionMethod, ZipWriter};

#[test]
fn round_trip_complex_document_in_memory() -> Result<(), rusdox::DocxError> {
    let grid = TableBorders::new()
        .top(Border::new(BorderStyle::Single).size(8).color("111827"))
        .bottom(Border::new(BorderStyle::Single).size(8).color("111827"))
        .left(Border::new(BorderStyle::Single).size(8).color("111827"))
        .right(Border::new(BorderStyle::Single).size(8).color("111827"));

    let mut document = Document::new();
    document.push_paragraph(
        Paragraph::new()
            .with_alignment(ParagraphAlignment::Center)
            .add_run(Run::from_text(" Rus").bold())
            .add_run(Run::from_text("Dox ").italic())
            .add_run(Run::from_text("launch").underline(UnderlineStyle::Single)),
    );

    document.push_table(
        Table::new()
            .width(9_360)
            .borders(grid)
            .add_row(
                TableRow::new()
                    .add_cell(
                        TableCell::new()
                            .width(4_680)
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("A1"))),
                    )
                    .add_cell(
                        TableCell::new()
                            .width(4_680)
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("B1"))),
                    ),
            )
            .add_row(
                TableRow::new()
                    .add_cell(
                        TableCell::new()
                            .width(4_680)
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("A2"))),
                    )
                    .add_cell(
                        TableCell::new()
                            .width(4_680)
                            .add_paragraph(Paragraph::new().add_run(Run::from_text("B2"))),
                    ),
            ),
    );

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadWrite)?;

    assert_eq!(reopened.paragraphs().count(), 1);
    assert_eq!(reopened.tables().count(), 1);
    assert_eq!(
        reopened.paragraphs().next().map(Paragraph::text),
        Some(" RusDox launch".to_string())
    );
    assert!(reopened.text().contains("A1"));
    assert!(reopened.text().contains("B2"));

    Ok(())
}

#[test]
fn save_preserves_non_document_parts() -> Result<(), rusdox::DocxError> {
    let mut document = Document::new();
    document.push_paragraph(Paragraph::new().add_run(Run::from_text("hello")));

    let mut source = Cursor::new(Vec::new());
    document.save_to_writer(&mut source)?;
    source.set_position(0);

    let mut reopened = Document::open_from_reader(&mut source, DocumentMode::ReadWrite)?;
    reopened.push_paragraph(Paragraph::new().add_run(Run::from_text("world")));

    let mut output = Cursor::new(Vec::new());
    reopened.save_to_writer(&mut output)?;
    output.set_position(0);

    let mut archive = ZipArchive::new(output)?;
    let mut parts = Vec::new();
    for index in 0..archive.len() {
        parts.push(archive.by_index(index)?.name().to_string());
    }

    assert!(parts.iter().any(|name| name == "[Content_Types].xml"));
    assert!(parts.iter().any(|name| name == "_rels/.rels"));
    assert!(parts.iter().any(|name| name == "word/document.xml"));
    assert!(parts.iter().any(|name| name == "word/styles.xml"));

    Ok(())
}

#[test]
fn layout_controls_round_trip_and_emit_header_footer_parts() -> Result<(), rusdox::DocxError> {
    let page_setup = PageSetup::new(13_000, 17_000)
        .margins(900, 950, 1_000, 1_050)
        .header_footer_distances(500, 550)
        .gutter(120);
    let header = HeaderFooter::new("Quarterly Review").with_alignment(ParagraphAlignment::Center);
    let footer =
        HeaderFooter::new("Page {page} of {pages}").with_alignment(ParagraphAlignment::Right);
    let numbering = PageNumbering::new(PageNumberFormat::UpperRoman).start_at(7);

    let mut document = Document::new()
        .with_page_setup(page_setup.clone())
        .with_header(header.clone())
        .with_footer(footer.clone())
        .with_page_numbering(numbering.clone());
    document.push_paragraph(Paragraph::new().add_run(Run::from_text("Body")));

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadWrite)?;
    assert_eq!(reopened.page_setup(), &page_setup);
    assert_eq!(reopened.header(), Some(&header));
    assert_eq!(reopened.footer(), Some(&footer));
    assert_eq!(reopened.page_numbering(), Some(&numbering));

    buffer.set_position(0);
    let mut archive = ZipArchive::new(buffer)?;

    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")?
        .read_to_string(&mut document_xml)?;
    assert!(document_xml
        .contains(r#"w:headerReference w:type="default" r:id="rIdRusDoxHeaderDefault""#));
    assert!(document_xml
        .contains(r#"w:footerReference w:type="default" r:id="rIdRusDoxFooterDefault""#));
    assert!(document_xml.contains(r#"<w:pgSz w:w="13000" w:h="17000"/>"#));
    assert!(document_xml.contains(r#"<w:pgMar w:top="900" w:right="950" w:bottom="1000" w:left="1050" w:header="500" w:footer="550" w:gutter="120"/>"#));
    assert!(document_xml.contains(r#"<w:pgNumType w:fmt="upperRoman" w:start="7"/>"#));

    let mut header_xml = String::new();
    archive
        .by_name("word/header1.xml")?
        .read_to_string(&mut header_xml)?;
    assert!(header_xml.contains("Quarterly Review"));
    assert!(header_xml.contains(r#"<w:jc w:val="center"/>"#));

    let mut footer_xml = String::new();
    archive
        .by_name("word/footer1.xml")?
        .read_to_string(&mut footer_xml)?;
    assert!(footer_xml.contains(" PAGE "));
    assert!(footer_xml.contains(" NUMPAGES "));
    assert!(footer_xml.contains(r#"<w:jc w:val="right"/>"#));

    let mut rels_xml = String::new();
    archive
        .by_name("word/_rels/document.xml.rels")?
        .read_to_string(&mut rels_xml)?;
    assert!(rels_xml.contains("relationships/header"));
    assert!(rels_xml.contains("relationships/footer"));

    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")?
        .read_to_string(&mut content_types)?;
    assert!(content_types.contains("/word/header1.xml"));
    assert!(content_types.contains("/word/footer1.xml"));

    Ok(())
}

#[test]
fn document_xml_contains_expected_whitespace_handling() -> Result<(), rusdox::DocxError> {
    let mut document = Document::new();
    document.push_paragraph(Paragraph::new().add_run(Run::from_text(" padded ")));

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let mut archive = ZipArchive::new(buffer)?;
    let mut entry = archive.by_name("word/document.xml")?;
    let mut xml = String::new();
    entry.read_to_string(&mut xml)?;

    assert!(xml.contains(r#"xml:space="preserve""#));
    Ok(())
}

#[test]
fn open_read_only_mode_rejects_save() -> Result<(), rusdox::DocxError> {
    let mut document = Document::new();
    document.push_paragraph(Paragraph::new().add_run(Run::from_text("immutable")));

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadOnly)?;
    let error = reopened.save_to_writer(Cursor::new(Vec::new())).err();

    assert!(error.is_some());
    Ok(())
}

#[test]
fn rich_formatting_round_trips() -> Result<(), rusdox::DocxError> {
    let mut document = Document::new();
    document.push_paragraph(
        Paragraph::new()
            .with_alignment(ParagraphAlignment::Center)
            .spacing_before(120)
            .spacing_after(240)
            .page_break_before()
            .add_run(
                Run::from_text("Headline")
                    .font("Arial")
                    .size_points(18)
                    .color("0F172A"),
            ),
    );
    document.push_table(
        Table::new().add_row(
            TableRow::new().add_cell(
                TableCell::new()
                    .background("E2E8F0")
                    .add_paragraph(Paragraph::new().add_run(Run::from_text("Metric"))),
            ),
        ),
    );

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadWrite)?;
    let paragraph = reopened
        .paragraphs()
        .next()
        .expect("paragraph should exist");
    let run = paragraph.runs().next().expect("run should exist");
    let table = reopened.tables().next().expect("table should exist");
    let cell = table
        .rows()
        .next()
        .expect("row should exist")
        .cells()
        .next()
        .expect("cell should exist");

    assert_eq!(paragraph.spacing_before_value(), Some(120));
    assert_eq!(paragraph.spacing_after_value(), Some(240));
    assert!(paragraph.has_page_break_before());
    assert_eq!(run.properties().font_family.as_deref(), Some("Arial"));
    assert_eq!(run.properties().font_size, Some(36));
    assert_eq!(
        cell.properties().background_color.as_deref(),
        Some("E2E8F0")
    );

    Ok(())
}

#[test]
fn named_styles_round_trip_and_preserve_style_references() -> Result<(), rusdox::DocxError> {
    let border = Border::new(BorderStyle::Single).size(8).color("CBD5E1");
    let styles = Stylesheet::new()
        .add_paragraph_style(
            ParagraphStyle::new("Lead")
                .name("Lead")
                .based_on("Normal")
                .next("Body")
                .paragraph(ParagraphStyleProperties {
                    list: Some(ParagraphList::bullet_with_id(12)),
                    alignment: Some(ParagraphAlignment::Center),
                    spacing_before: Some(120),
                    spacing_after: Some(240),
                    keep_next: Some(true),
                    page_break_before: Some(false),
                })
                .run(RunStyleProperties::new().bold().color("112233")),
        )
        .add_paragraph_style(ParagraphStyle::new("Body").based_on("Lead").paragraph(
            ParagraphStyleProperties {
                keep_next: Some(false),
                ..ParagraphStyleProperties::default()
            },
        ))
        .add_run_style(
            RunStyle::new("Accent")
                .based_on("DefaultParagraphFont")
                .properties(RunStyleProperties::new().italic().color("AA5500")),
        )
        .add_table_style(
            TableStyle::new("DataGrid")
                .based_on("TableNormal")
                .properties(
                    TableStyleProperties::new().width(9_360).borders(
                        TableBorders::new()
                            .top(border.clone())
                            .bottom(border.clone()),
                    ),
                ),
        );

    let mut document = Document::new().with_styles(styles);
    document.push_paragraph(
        Paragraph::new()
            .with_style("Lead")
            .add_run(Run::from_text("Quarterly").with_style("Accent")),
    );
    document.push_table(
        Table::new()
            .style("DataGrid")
            .add_row(TableRow::new().add_cell(
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("ARR"))),
            )),
    );

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadWrite)?;
    let paragraph = reopened
        .paragraphs()
        .next()
        .expect("paragraph should exist");
    let run = paragraph.runs().next().expect("run should exist");
    let table = reopened.tables().next().expect("table should exist");

    assert_eq!(paragraph.style_id(), Some("Lead"));
    assert_eq!(paragraph.alignment(), None);
    assert_eq!(paragraph.spacing_before_value(), None);
    assert_eq!(paragraph.spacing_after_value(), None);
    assert!(!paragraph.has_keep_next());
    assert!(!paragraph.has_page_break_before());
    assert_eq!(run.style_id(), Some("Accent"));
    assert_eq!(run.properties().color, None);
    assert_eq!(run.properties().font_family, None);
    assert_eq!(table.style_id(), Some("DataGrid"));
    assert_eq!(table.properties().width, None);
    assert_eq!(table.properties().borders, None);

    let lead = reopened
        .styles()
        .paragraph_style("Lead")
        .expect("lead style should exist");
    assert_eq!(lead.based_on.as_deref(), Some("Normal"));
    assert_eq!(lead.next.as_deref(), Some("Body"));
    assert_eq!(lead.paragraph.list, Some(ParagraphList::bullet_with_id(12)));
    assert_eq!(lead.run.bold, Some(true));
    assert_eq!(lead.run.color.as_deref(), Some("112233"));

    let accent = reopened
        .styles()
        .run_style("Accent")
        .expect("accent style should exist");
    assert_eq!(accent.based_on.as_deref(), Some("DefaultParagraphFont"));
    assert_eq!(accent.properties.italic, Some(true));
    assert_eq!(accent.properties.color.as_deref(), Some("AA5500"));

    let grid = reopened
        .styles()
        .table_style("DataGrid")
        .expect("table style should exist");
    assert_eq!(grid.based_on.as_deref(), Some("TableNormal"));
    assert_eq!(grid.properties.width, Some(9_360));
    assert_eq!(
        grid.properties
            .borders
            .as_ref()
            .and_then(|borders| borders.top.as_ref()),
        Some(&border)
    );

    buffer.set_position(0);
    let mut archive = ZipArchive::new(buffer)?;

    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")?
        .read_to_string(&mut document_xml)?;
    assert!(document_xml.contains(r#"<w:pStyle w:val="Lead"/>"#));
    assert!(document_xml.contains(r#"<w:rStyle w:val="Accent"/>"#));
    assert!(document_xml.contains(r#"<w:tblStyle w:val="DataGrid"/>"#));

    let mut styles_xml = String::new();
    archive
        .by_name("word/styles.xml")?
        .read_to_string(&mut styles_xml)?;
    assert!(styles_xml.contains(r#"<w:style w:type="paragraph" w:styleId="Lead""#));
    assert!(styles_xml.contains(r#"<w:basedOn w:val="Normal"/>"#));
    assert!(styles_xml.contains(r#"<w:style w:type="character" w:styleId="Accent""#));
    assert!(styles_xml.contains(r#"<w:style w:type="table" w:styleId="DataGrid""#));
    assert!(styles_xml.contains(r#"<w:keepNext w:val="0"/>"#));
    assert!(styles_xml.contains(r#"<w:pageBreakBefore w:val="0"/>"#));

    let mut numbering_xml = String::new();
    archive
        .by_name("word/numbering.xml")?
        .read_to_string(&mut numbering_xml)?;
    assert!(numbering_xml.contains(r#"<w:num w:numId="12">"#));

    Ok(())
}

#[test]
fn table_row_pagination_properties_round_trip_and_emit_ooxml() -> Result<(), rusdox::DocxError> {
    let mut document = Document::new();
    document.push_table(
        Table::new()
            .add_row(TableRow::new().repeat_as_header().add_cell(
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("Header"))),
            ))
            .add_row(TableRow::new().allow_split_across_pages(false).add_cell(
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("Body"))),
            )),
    );

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadWrite)?;
    let table = reopened.tables().next().expect("table should exist");
    let rows = table.rows().collect::<Vec<_>>();

    assert_eq!(rows.len(), 2);
    assert!(rows[0].properties().repeat_as_header);
    assert!(rows[0].properties().allow_split_across_pages);
    assert!(!rows[1].properties().repeat_as_header);
    assert!(!rows[1].properties().allow_split_across_pages);

    buffer.set_position(0);
    let mut archive = ZipArchive::new(buffer)?;
    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")?
        .read_to_string(&mut document_xml)?;

    assert!(document_xml.contains("<w:tblHeader/>"));
    assert!(document_xml.contains("<w:cantSplit/>"));

    Ok(())
}

#[test]
fn visual_blocks_round_trip_and_emit_media_parts() -> Result<(), rusdox::DocxError> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let mut document = Document::new();
    document.push_visual(
        Visual::logo(root.join("assets/rusdox-mark.svg"))
            .alt_text_text("RusDox logo")
            .max_width_twips(2_200),
    );
    document.push_visual(
        Visual::from_path(root.join("assets/template-gallery.png"))
            .alt_text_text("RusDox gallery")
            .max_width_twips(7_200),
    );
    document.push_visual(
        Visual::chart(root.join("assets/benchmark-stress-1000-pages.svg"))
            .alt_text_text("RusDox benchmark chart")
            .max_width_twips(7_200),
    );
    document.push_visual(
        Visual::signature(root.join("assets/signature-demo.svg"))
            .alt_text_text("Approval signature")
            .max_width_twips(2_800),
    );

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadWrite)?;
    let kinds = reopened
        .visuals()
        .map(|visual| visual.kind())
        .collect::<Vec<_>>();
    assert_eq!(
        kinds,
        vec![
            VisualKind::Logo,
            VisualKind::Image,
            VisualKind::Chart,
            VisualKind::Signature,
        ]
    );
    assert_eq!(
        reopened
            .visuals()
            .next()
            .and_then(|visual| visual.alt_text()),
        Some("RusDox logo")
    );

    buffer.set_position(0);
    let mut archive = ZipArchive::new(buffer)?;
    let mut media_parts = Vec::new();
    for index in 0..archive.len() {
        let name = archive.by_index(index)?.name().to_string();
        if name.starts_with("word/media/") {
            media_parts.push(name);
        }
    }
    assert_eq!(media_parts.len(), 4);
    assert!(media_parts.iter().all(|name| name.ends_with(".png")));

    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")?
        .read_to_string(&mut document_xml)?;
    assert!(document_xml.contains("<w:drawing>"));
    assert!(document_xml.contains("RusDox Logo 1"));
    assert!(document_xml.contains("rIdRusDoxImage1"));

    let mut rels_xml = String::new();
    archive
        .by_name("word/_rels/document.xml.rels")?
        .read_to_string(&mut rels_xml)?;
    assert!(rels_xml.contains("relationships/image"));

    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")?
        .read_to_string(&mut content_types)?;
    assert!(content_types.contains(r#"Extension="png""#));

    Ok(())
}

#[test]
fn pdf_preview_renders_visual_blocks_as_image_xobjects() -> Result<(), rusdox::DocxError> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let temp = tempdir()?;
    let mut config = RusdoxConfig::default();
    config.output.docx_dir = temp.path().join("generated").to_string_lossy().to_string();
    config.output.pdf_dir = temp.path().join("rendered").to_string_lossy().to_string();
    config.output.emit_pdf_preview = true;
    let studio = Studio::new(config);

    let mut document = Document::new();
    document.push_visual(
        Visual::logo(root.join("assets/rusdox-mark.svg"))
            .alt_text_text("RusDox logo")
            .max_width_twips(2_200),
    );
    document.push_visual(
        Visual::chart(root.join("assets/benchmark-stress-1000-pages.svg"))
            .alt_text_text("RusDox benchmark chart")
            .max_width_twips(7_200),
    );

    let docx_path = temp.path().join("generated/visual-preview.docx");
    studio.save_with_pdf_stats(&document, &docx_path)?;

    let pdf_path = temp.path().join("rendered/visual-preview.pdf");
    let pdf = std::fs::read(pdf_path)?;
    let pdf_text = String::from_utf8_lossy(&pdf);

    assert!(pdf_text.matches("/Subtype /Image").count() >= 2);
    assert!(pdf_text.contains("/XObject"));

    Ok(())
}

#[test]
fn tabs_and_line_breaks_round_trip_and_emit_structured_ooxml() -> Result<(), rusdox::DocxError> {
    let mut document = Document::new();
    document
        .push_paragraph(Paragraph::new().add_run(Run::from_text("Alpha\tBeta\r\nGamma\nDelta")));

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadWrite)?;
    let paragraph = reopened
        .paragraphs()
        .next()
        .expect("paragraph should exist");
    let run = paragraph.runs().next().expect("run should exist");

    assert_eq!(run.text(), "Alpha\tBeta\nGamma\nDelta");

    buffer.set_position(0);
    let mut archive = ZipArchive::new(buffer)?;
    let mut entry = archive.by_name("word/document.xml")?;
    let mut xml = String::new();
    entry.read_to_string(&mut xml)?;

    assert!(xml.contains("<w:tab/>"));
    assert!(xml.matches("<w:br/>").count() >= 2);
    assert!(!xml.contains("Alpha\tBeta"));

    Ok(())
}

#[test]
fn semantic_lists_round_trip_and_emit_numbering_parts() -> Result<(), rusdox::DocxError> {
    let mut document = Document::new();
    document.push_paragraph(
        Paragraph::new()
            .with_list(ParagraphList::bullet_with_id(3))
            .add_run(Run::from_text("First bullet")),
    );
    document.push_paragraph(
        Paragraph::new()
            .with_list(ParagraphList::bullet_with_id(3))
            .add_run(Run::from_text("Second bullet")),
    );
    document.push_paragraph(
        Paragraph::new()
            .with_list(ParagraphList::numbered_with_id(4).with_level(1))
            .add_run(Run::from_text("Nested number")),
    );

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadWrite)?;
    let paragraphs: Vec<_> = reopened.paragraphs().collect();
    assert_eq!(paragraphs.len(), 3);
    assert_eq!(
        paragraphs[0].list(),
        Some(&ParagraphList::bullet_with_id(3))
    );
    assert_eq!(
        paragraphs[1].list(),
        Some(&ParagraphList::bullet_with_id(3))
    );
    assert_eq!(
        paragraphs[2].list(),
        Some(&ParagraphList::numbered_with_id(4).with_level(1))
    );
    assert_eq!(paragraphs[0].text(), "First bullet");

    buffer.set_position(0);
    let mut archive = ZipArchive::new(buffer)?;

    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")?
        .read_to_string(&mut document_xml)?;
    assert!(document_xml.contains("<w:numPr>"));
    assert!(!document_xml.contains("• First bullet"));

    let mut numbering_xml = String::new();
    archive
        .by_name("word/numbering.xml")?
        .read_to_string(&mut numbering_xml)?;
    assert!(numbering_xml.contains(r#"w:numId="3""#));
    assert!(numbering_xml.contains(r#"w:numId="4""#));

    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")?
        .read_to_string(&mut content_types)?;
    assert!(content_types.contains("/word/numbering.xml"));

    let mut rels = String::new();
    archive
        .by_name("word/_rels/document.xml.rels")?
        .read_to_string(&mut rels)?;
    assert!(rels.contains("relationships/numbering"));

    Ok(())
}

#[test]
fn custom_package_part_is_preserved_across_open_modify_save() -> Result<(), rusdox::DocxError> {
    let mut original = Document::new();
    original.push_paragraph(Paragraph::new().add_run(Run::from_text("base")));

    let mut base_buffer = Cursor::new(Vec::new());
    original.save_to_writer(&mut base_buffer)?;
    base_buffer.set_position(0);

    let mut base_archive = ZipArchive::new(base_buffer)?;
    let mut with_custom = Cursor::new(Vec::new());
    {
        let mut writer = ZipWriter::new(&mut with_custom);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for index in 0..base_archive.len() {
            let mut file = base_archive.by_index(index)?;
            let name = file.name().to_string();
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)?;
            writer.start_file(name, options)?;
            writer.write_all(&bytes)?;
        }

        writer.start_file("custom/metadata.bin", options)?;
        writer.write_all(&[1_u8, 2, 3, 4, 5])?;
        writer.finish()?;
    }
    with_custom.set_position(0);

    let mut reopened = Document::open_from_reader(&mut with_custom, DocumentMode::ReadWrite)?;
    reopened.push_paragraph(Paragraph::new().add_run(Run::from_text("updated")));

    let mut out = Cursor::new(Vec::new());
    reopened.save_to_writer(&mut out)?;
    out.set_position(0);

    let mut archive = ZipArchive::new(out)?;
    let mut custom = archive.by_name("custom/metadata.bin")?;
    let mut bytes = Vec::new();
    custom.read_to_end(&mut bytes)?;
    assert_eq!(bytes, vec![1, 2, 3, 4, 5]);
    Ok(())
}

#[test]
fn invalid_zip_input_returns_zip_error() {
    let mut invalid = Cursor::new(vec![0x01_u8, 0x02, 0x03, 0x04]);
    let error = Document::open_from_reader(&mut invalid, DocumentMode::ReadWrite)
        .expect_err("invalid zip should fail");
    assert!(matches!(error, rusdox::DocxError::Zip(_)));
}

#[test]
fn save_to_path_rejects_read_only_document() -> Result<(), rusdox::DocxError> {
    let mut document = Document::new();
    document.push_paragraph(Paragraph::new().add_run(Run::from_text("immutable")));

    let mut buffer = Cursor::new(Vec::new());
    document.save_to_writer(&mut buffer)?;
    buffer.set_position(0);

    let reopened = Document::open_from_reader(&mut buffer, DocumentMode::ReadOnly)?;
    let temp = tempdir().expect("temp dir");
    let path = temp.path().join("readonly.docx");
    let error = reopened.save(&path).expect_err("save must fail");

    assert!(matches!(error, rusdox::DocxError::Parse(message) if message.contains("read-only")));
    Ok(())
}

#[test]
fn open_from_disk_sets_source_path_and_open_read_only_mode() -> Result<(), rusdox::DocxError> {
    let temp = tempdir().expect("temp dir");
    let path = temp.path().join("sample.docx");

    let mut document = Document::new();
    document.push_paragraph(Paragraph::new().add_run(Run::from_text("disk")));
    document.save(&path)?;

    let read_write = Document::open(&path)?;
    let read_only = Document::open_read_only(&path)?;

    assert_eq!(read_write.mode(), DocumentMode::ReadWrite);
    assert_eq!(read_only.mode(), DocumentMode::ReadOnly);
    assert_eq!(read_write.source_path(), Some(path.as_path()));
    assert_eq!(read_only.source_path(), Some(path.as_path()));
    Ok(())
}
