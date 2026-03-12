use crate::paragraph::Paragraph;

/// A supported border style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BorderStyle {
    /// No visible border.
    None,
    /// A single line border.
    Single,
    /// A double line border.
    Double,
    /// A dotted border.
    Dotted,
    /// A dashed border.
    Dashed,
    /// Preserve or emit a custom OOXML border value.
    Custom(String),
}

impl BorderStyle {
    pub(crate) fn from_xml(value: &str) -> Self {
        match value {
            "nil" | "none" => Self::None,
            "single" => Self::Single,
            "double" => Self::Double,
            "dotted" => Self::Dotted,
            "dashed" => Self::Dashed,
            other => Self::Custom(other.to_string()),
        }
    }

    pub(crate) fn as_xml_value(&self) -> &str {
        match self {
            Self::None => "nil",
            Self::Single => "single",
            Self::Double => "double",
            Self::Dotted => "dotted",
            Self::Dashed => "dashed",
            Self::Custom(value) => value.as_str(),
        }
    }
}

/// A single table or cell border.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Border {
    /// The border style.
    pub style: BorderStyle,
    /// Optional line size in eighths of a point.
    pub size: Option<u16>,
    /// Optional hexadecimal RGB color value.
    pub color: Option<String>,
}

impl Border {
    /// Creates a border with the provided style.
    pub fn new(style: BorderStyle) -> Self {
        Self {
            style,
            size: None,
            color: None,
        }
    }

    /// Sets the border size.
    pub fn size(mut self, size: u16) -> Self {
        self.size = Some(size);
        self
    }

    /// Sets the border color.
    pub fn color(mut self, color: impl Into<String>) -> Self {
        self.color = Some(color.into());
        self
    }
}

/// Border collection for tables and table cells.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableBorders {
    /// Top border.
    pub top: Option<Border>,
    /// Bottom border.
    pub bottom: Option<Border>,
    /// Left border.
    pub left: Option<Border>,
    /// Right border.
    pub right: Option<Border>,
    /// Horizontal internal border.
    pub inside_horizontal: Option<Border>,
    /// Vertical internal border.
    pub inside_vertical: Option<Border>,
}

impl TableBorders {
    /// Creates an empty border set.
    pub fn new() -> Self {
        Self::default()
    }

    /// Sets the top border.
    pub fn top(mut self, border: Border) -> Self {
        self.top = Some(border);
        self
    }

    /// Sets the bottom border.
    pub fn bottom(mut self, border: Border) -> Self {
        self.bottom = Some(border);
        self
    }

    /// Sets the left border.
    pub fn left(mut self, border: Border) -> Self {
        self.left = Some(border);
        self
    }

    /// Sets the right border.
    pub fn right(mut self, border: Border) -> Self {
        self.right = Some(border);
        self
    }

    /// Sets the internal horizontal border.
    pub fn inside_horizontal(mut self, border: Border) -> Self {
        self.inside_horizontal = Some(border);
        self
    }

    /// Sets the internal vertical border.
    pub fn inside_vertical(mut self, border: Border) -> Self {
        self.inside_vertical = Some(border);
        self
    }

    pub(crate) fn has_serialized_content(&self) -> bool {
        self.top.is_some()
            || self.bottom.is_some()
            || self.left.is_some()
            || self.right.is_some()
            || self.inside_horizontal.is_some()
            || self.inside_vertical.is_some()
    }
}

/// Properties attached to a table.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableProperties {
    /// Optional table width in DXA units.
    pub width: Option<u32>,
    /// Optional table borders.
    pub borders: Option<TableBorders>,
}

/// Properties attached to a table cell.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableCellProperties {
    /// Optional cell width in DXA units.
    pub width: Option<u32>,
    /// Optional grid span.
    pub grid_span: Option<u32>,
    /// Optional cell borders.
    pub borders: Option<TableBorders>,
    /// Optional cell background color in hexadecimal RGB form.
    pub background_color: Option<String>,
}

/// A table cell containing one or more paragraphs.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableCell {
    paragraphs: Vec<Paragraph>,
    properties: TableCellProperties,
}

impl TableCell {
    /// Creates an empty cell.
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds a paragraph using a builder-style API.
    pub fn add_paragraph(mut self, paragraph: Paragraph) -> Self {
        self.paragraphs.push(paragraph);
        self
    }

    /// Appends a paragraph in place.
    pub fn push_paragraph(&mut self, paragraph: Paragraph) -> &mut Self {
        self.paragraphs.push(paragraph);
        self
    }

    /// Returns immutable access to cell paragraphs.
    pub fn paragraphs(&self) -> std::slice::Iter<'_, Paragraph> {
        self.paragraphs.iter()
    }

    /// Returns mutable access to cell paragraphs.
    pub fn paragraphs_mut(&mut self) -> std::slice::IterMut<'_, Paragraph> {
        self.paragraphs.iter_mut()
    }

    /// Sets the cell width in DXA units.
    pub fn width(mut self, width: u32) -> Self {
        self.properties.width = Some(width);
        self
    }

    /// Sets the cell grid span.
    pub fn grid_span(mut self, grid_span: u32) -> Self {
        self.properties.grid_span = Some(grid_span);
        self
    }

    /// Applies cell borders.
    pub fn borders(mut self, borders: TableBorders) -> Self {
        self.properties.borders = Some(borders);
        self
    }

    /// Applies a background color to the cell.
    pub fn background(mut self, color: impl Into<String>) -> Self {
        self.properties.background_color = Some(color.into());
        self
    }

    /// Returns the cell properties.
    pub fn properties(&self) -> &TableCellProperties {
        &self.properties
    }

    /// Returns mutable access to cell properties.
    pub fn properties_mut(&mut self) -> &mut TableCellProperties {
        &mut self.properties
    }

    /// Extracts plain text from all cell paragraphs.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(Paragraph::text)
            .collect::<Vec<_>>()
            .join("\n")
    }

    pub(crate) fn from_parts(paragraphs: Vec<Paragraph>, properties: TableCellProperties) -> Self {
        Self {
            paragraphs,
            properties,
        }
    }
}

/// A row in a table.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableRow {
    cells: Vec<TableCell>,
}

impl TableRow {
    /// Creates an empty row.
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds a cell using a builder-style API.
    pub fn add_cell(mut self, cell: TableCell) -> Self {
        self.cells.push(cell);
        self
    }

    /// Appends a cell in place.
    pub fn push_cell(&mut self, cell: TableCell) -> &mut Self {
        self.cells.push(cell);
        self
    }

    /// Returns immutable access to row cells.
    pub fn cells(&self) -> std::slice::Iter<'_, TableCell> {
        self.cells.iter()
    }

    /// Returns mutable access to row cells.
    pub fn cells_mut(&mut self) -> std::slice::IterMut<'_, TableCell> {
        self.cells.iter_mut()
    }

    pub(crate) fn from_parts(cells: Vec<TableCell>) -> Self {
        Self { cells }
    }
}

/// A document table.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Table {
    rows: Vec<TableRow>,
    properties: TableProperties,
}

impl Table {
    /// Creates an empty table.
    ///
    /// ```rust
    /// use rusdox::Table;
    ///
    /// let table = Table::new();
    /// assert!(table.rows().next().is_none());
    /// ```
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds a row using a builder-style API.
    pub fn add_row(mut self, row: TableRow) -> Self {
        self.rows.push(row);
        self
    }

    /// Appends a row in place.
    pub fn push_row(&mut self, row: TableRow) -> &mut Self {
        self.rows.push(row);
        self
    }

    /// Sets the table width in DXA units.
    pub fn width(mut self, width: u32) -> Self {
        self.properties.width = Some(width);
        self
    }

    /// Applies table borders.
    pub fn borders(mut self, borders: TableBorders) -> Self {
        self.properties.borders = Some(borders);
        self
    }

    /// Returns immutable access to rows.
    pub fn rows(&self) -> std::slice::Iter<'_, TableRow> {
        self.rows.iter()
    }

    /// Returns mutable access to rows.
    pub fn rows_mut(&mut self) -> std::slice::IterMut<'_, TableRow> {
        self.rows.iter_mut()
    }

    /// Returns the table properties.
    pub fn properties(&self) -> &TableProperties {
        &self.properties
    }

    /// Returns mutable access to table properties.
    pub fn properties_mut(&mut self) -> &mut TableProperties {
        &mut self.properties
    }

    /// Extracts plain text from the table.
    pub fn text(&self) -> String {
        self.rows
            .iter()
            .map(|row| {
                row.cells()
                    .map(TableCell::text)
                    .collect::<Vec<_>>()
                    .join("\t")
            })
            .collect::<Vec<_>>()
            .join("\n")
    }

    pub(crate) fn from_parts(rows: Vec<TableRow>, properties: TableProperties) -> Self {
        Self { rows, properties }
    }
}

#[cfg(test)]
mod tests {
    use super::{
        Border, BorderStyle, Table, TableBorders, TableCell, TableCellProperties, TableProperties,
        TableRow,
    };
    use crate::{Paragraph, Run};

    #[test]
    fn border_style_round_trips_known_values() {
        let cases = [
            ("nil", BorderStyle::None, "nil"),
            ("none", BorderStyle::None, "nil"),
            ("single", BorderStyle::Single, "single"),
            ("double", BorderStyle::Double, "double"),
            ("dotted", BorderStyle::Dotted, "dotted"),
            ("dashed", BorderStyle::Dashed, "dashed"),
        ];

        for (xml, expected, roundtrip_xml) in cases {
            let parsed = BorderStyle::from_xml(xml);
            assert_eq!(parsed, expected);
            assert_eq!(parsed.as_xml_value(), roundtrip_xml);
        }
    }

    #[test]
    fn border_style_custom_value_is_preserved() {
        let parsed = BorderStyle::from_xml("thickThinLargeGap");
        assert_eq!(parsed, BorderStyle::Custom("thickThinLargeGap".to_string()));
        assert_eq!(parsed.as_xml_value(), "thickThinLargeGap");
    }

    #[test]
    fn border_builder_sets_size_and_color() {
        let border = Border::new(BorderStyle::Single).size(16).color("AABBCC");
        assert_eq!(border.style, BorderStyle::Single);
        assert_eq!(border.size, Some(16));
        assert_eq!(border.color.as_deref(), Some("AABBCC"));
    }

    #[test]
    fn table_borders_builder_and_serialization_flag() {
        let empty = TableBorders::new();
        assert!(!empty.has_serialized_content());

        let border = Border::new(BorderStyle::Single).size(8).color("111111");
        let filled = TableBorders::new()
            .top(border.clone())
            .bottom(border.clone())
            .left(border.clone())
            .right(border.clone())
            .inside_horizontal(border.clone())
            .inside_vertical(border);
        assert!(filled.has_serialized_content());
        assert!(filled.top.is_some());
        assert!(filled.bottom.is_some());
        assert!(filled.left.is_some());
        assert!(filled.right.is_some());
        assert!(filled.inside_horizontal.is_some());
        assert!(filled.inside_vertical.is_some());
    }

    #[test]
    fn table_cell_builder_sets_all_properties_and_text() {
        let borders = TableBorders::new().top(Border::new(BorderStyle::Single));
        let mut cell = TableCell::new()
            .width(1234)
            .grid_span(2)
            .borders(borders.clone())
            .background("DDEEFF")
            .add_paragraph(Paragraph::new().add_run(Run::from_text("A")))
            .add_paragraph(Paragraph::new().add_run(Run::from_text("B")));

        cell.push_paragraph(Paragraph::new().add_run(Run::from_text("C")));
        cell.properties_mut().width = Some(5678);

        assert_eq!(cell.properties().width, Some(5678));
        assert_eq!(cell.properties().grid_span, Some(2));
        assert_eq!(cell.properties().borders.as_ref(), Some(&borders));
        assert_eq!(
            cell.properties().background_color.as_deref(),
            Some("DDEEFF")
        );
        assert_eq!(cell.text(), "A\nB\nC");
    }

    #[test]
    fn table_row_builder_and_cells_mut_allow_changes() {
        let mut row = TableRow::new()
            .add_cell(TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("L"))))
            .add_cell(
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("R"))),
            );

        row.push_cell(
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("X"))),
        );

        for cell in row.cells_mut() {
            if cell.text() == "R" {
                cell.push_paragraph(Paragraph::new().add_run(Run::from_text("2")));
            }
        }

        let texts: Vec<_> = row.cells().map(TableCell::text).collect();
        assert_eq!(
            texts,
            vec!["L".to_string(), "R\n2".to_string(), "X".to_string()]
        );
    }

    #[test]
    fn table_builder_sets_properties_and_formats_text_grid() {
        let borders = TableBorders::new().top(Border::new(BorderStyle::Single));
        let mut table = Table::new().width(9360).borders(borders.clone()).add_row(
            TableRow::new()
                .add_cell(
                    TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("H1"))),
                )
                .add_cell(
                    TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("H2"))),
                ),
        );

        table.push_row(
            TableRow::new()
                .add_cell(
                    TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("V1"))),
                )
                .add_cell(
                    TableCell::new().add_paragraph(Paragraph::new().add_run(Run::from_text("V2"))),
                ),
        );
        table.properties_mut().width = Some(9000);

        assert_eq!(table.properties().width, Some(9000));
        assert_eq!(table.properties().borders.as_ref(), Some(&borders));
        assert_eq!(table.text(), "H1\tH2\nV1\tV2");
    }

    #[test]
    fn from_parts_builders_preserve_exact_state() {
        let cell_properties = TableCellProperties {
            width: Some(1000),
            grid_span: Some(3),
            borders: Some(TableBorders::new().left(Border::new(BorderStyle::Double))),
            background_color: Some("ABCDEF".to_string()),
        };
        let cell = TableCell::from_parts(
            vec![Paragraph::new().add_run(Run::from_text("value"))],
            cell_properties.clone(),
        );
        assert_eq!(cell.properties(), &cell_properties);

        let row = TableRow::from_parts(vec![cell.clone()]);
        assert_eq!(row.cells().count(), 1);
        assert_eq!(row.cells().next(), Some(&cell));

        let table_properties = TableProperties {
            width: Some(7777),
            borders: Some(TableBorders::new().right(Border::new(BorderStyle::Dashed))),
        };
        let table = Table::from_parts(vec![row], table_properties.clone());
        assert_eq!(table.properties(), &table_properties);
        assert_eq!(table.rows().count(), 1);
    }
}
