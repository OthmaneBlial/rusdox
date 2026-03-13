use std::io::{Cursor, Read, Write};

use rusdox::{
    Border, BorderStyle, Document, DocumentMode, HeaderFooter, PageNumberFormat, PageNumbering,
    PageSetup, Paragraph, ParagraphAlignment, ParagraphList, Run, Table, TableBorders, TableCell,
    TableRow, UnderlineStyle,
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
