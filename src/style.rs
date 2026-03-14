use serde::{Deserialize, Serialize};

use crate::{
    DocxError, Paragraph, ParagraphAlignment, ParagraphList, Result, Run, RunProperties, Table,
    TableBorders, TableProperties, UnderlineStyle, VerticalAlign,
};

const DEFAULT_PARAGRAPH_STYLE_ID: &str = "Normal";
const DEFAULT_RUN_STYLE_ID: &str = "DefaultParagraphFont";
const DEFAULT_TABLE_STYLE_ID: &str = "TableNormal";

/// A collection of reusable paragraph, run, and table styles.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct Stylesheet {
    pub paragraph: Vec<ParagraphStyle>,
    pub run: Vec<RunStyle>,
    pub table: Vec<TableStyle>,
}

impl Stylesheet {
    /// Creates an empty stylesheet.
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns whether the stylesheet has no custom style definitions.
    pub fn is_empty(&self) -> bool {
        self.paragraph.is_empty() && self.run.is_empty() && self.table.is_empty()
    }

    /// Returns paragraph styles in document order.
    pub fn paragraph_styles(&self) -> std::slice::Iter<'_, ParagraphStyle> {
        self.paragraph.iter()
    }

    /// Returns run styles in document order.
    pub fn run_styles(&self) -> std::slice::Iter<'_, RunStyle> {
        self.run.iter()
    }

    /// Returns table styles in document order.
    pub fn table_styles(&self) -> std::slice::Iter<'_, TableStyle> {
        self.table.iter()
    }

    /// Inserts or replaces a paragraph style by id.
    pub fn define_paragraph_style(&mut self, style: ParagraphStyle) -> &mut Self {
        upsert_style(&mut self.paragraph, style);
        self
    }

    /// Inserts or replaces a paragraph style by id in builder style.
    pub fn add_paragraph_style(mut self, style: ParagraphStyle) -> Self {
        self.define_paragraph_style(style);
        self
    }

    /// Inserts or replaces a run style by id.
    pub fn define_run_style(&mut self, style: RunStyle) -> &mut Self {
        upsert_style(&mut self.run, style);
        self
    }

    /// Inserts or replaces a run style by id in builder style.
    pub fn add_run_style(mut self, style: RunStyle) -> Self {
        self.define_run_style(style);
        self
    }

    /// Inserts or replaces a table style by id.
    pub fn define_table_style(&mut self, style: TableStyle) -> &mut Self {
        upsert_style(&mut self.table, style);
        self
    }

    /// Inserts or replaces a table style by id in builder style.
    pub fn add_table_style(mut self, style: TableStyle) -> Self {
        self.define_table_style(style);
        self
    }

    /// Returns a paragraph style by id.
    pub fn paragraph_style(&self, style_id: &str) -> Option<&ParagraphStyle> {
        self.paragraph.iter().find(|style| style.id == style_id)
    }

    /// Returns a run style by id.
    pub fn run_style(&self, style_id: &str) -> Option<&RunStyle> {
        self.run.iter().find(|style| style.id == style_id)
    }

    /// Returns a table style by id.
    pub fn table_style(&self, style_id: &str) -> Option<&TableStyle> {
        self.table.iter().find(|style| style.id == style_id)
    }

    pub(crate) fn resolve_paragraph(
        &self,
        paragraph: &Paragraph,
    ) -> Result<ResolvedParagraphStyle> {
        let mut resolved = self.resolve_paragraph_style(paragraph.style_id())?;
        if let Some(list) = paragraph.list().copied() {
            resolved.list = Some(list);
        }
        if let Some(alignment) = paragraph.alignment().cloned() {
            resolved.alignment = Some(alignment);
        }
        if let Some(spacing_before) = paragraph.spacing_before_value() {
            resolved.spacing_before = Some(spacing_before);
        }
        if let Some(spacing_after) = paragraph.spacing_after_value() {
            resolved.spacing_after = Some(spacing_after);
        }
        if paragraph.has_keep_next() {
            resolved.keep_next = true;
        }
        if paragraph.has_page_break_before() {
            resolved.page_break_before = true;
        }
        Ok(resolved)
    }

    pub(crate) fn resolve_run(&self, paragraph: &Paragraph, run: &Run) -> Result<RunProperties> {
        let mut resolved = self.resolve_paragraph_style(paragraph.style_id())?.run;
        let run_style = self.resolve_run_style(run.style_id())?;
        apply_run_style_properties(&mut resolved, &run_style);
        apply_direct_run_properties(&mut resolved, run.properties());
        Ok(resolved)
    }

    pub(crate) fn resolve_table(&self, table: &Table) -> Result<TableProperties> {
        let mut resolved = self.resolve_table_style(table.style_id())?;
        if let Some(width) = table.properties().width {
            resolved.width = Some(width);
        }
        if let Some(borders) = table.properties().borders.clone() {
            resolved.borders = Some(borders);
        }
        resolved.style_id = None;
        Ok(resolved)
    }

    pub(crate) fn resolve_paragraph_style(
        &self,
        style_id: Option<&str>,
    ) -> Result<ResolvedParagraphStyle> {
        match style_id {
            None => Ok(ResolvedParagraphStyle::default()),
            Some(style_id) if self.is_implicit_default_paragraph_style(style_id) => {
                Ok(ResolvedParagraphStyle::default())
            }
            Some(style_id) => self.resolve_paragraph_style_recursive(style_id, &mut Vec::new()),
        }
    }

    pub(crate) fn resolve_run_style(&self, style_id: Option<&str>) -> Result<RunStyleProperties> {
        match style_id {
            None => Ok(RunStyleProperties::default()),
            Some(style_id) if self.is_implicit_default_run_style(style_id) => {
                Ok(RunStyleProperties::default())
            }
            Some(style_id) => self.resolve_run_style_recursive(style_id, &mut Vec::new()),
        }
    }

    pub(crate) fn resolve_table_style(&self, style_id: Option<&str>) -> Result<TableProperties> {
        match style_id {
            None => Ok(TableProperties::default()),
            Some(style_id) if self.is_implicit_default_table_style(style_id) => {
                Ok(TableProperties::default())
            }
            Some(style_id) => self.resolve_table_style_recursive(style_id, &mut Vec::new()),
        }
    }

    fn resolve_paragraph_style_recursive(
        &self,
        style_id: &str,
        stack: &mut Vec<String>,
    ) -> Result<ResolvedParagraphStyle> {
        if self.is_implicit_default_paragraph_style(style_id) {
            return Ok(ResolvedParagraphStyle::default());
        }
        if stack.iter().any(|entry| entry == style_id) {
            return Err(DocxError::parse(format!(
                "paragraph style inheritance cycle detected at '{style_id}'"
            )));
        }
        let style = self
            .paragraph_style(style_id)
            .ok_or_else(|| DocxError::parse(format!("unknown paragraph style '{style_id}'")))?;
        stack.push(style_id.to_string());
        let mut resolved = if let Some(parent_id) = style.based_on.as_deref() {
            self.resolve_paragraph_style_recursive(parent_id, stack)?
        } else {
            ResolvedParagraphStyle::default()
        };
        let _ = stack.pop();
        apply_paragraph_style_properties(&mut resolved, &style.paragraph);
        apply_run_style_properties(&mut resolved.run, &style.run);
        Ok(resolved)
    }

    fn resolve_run_style_recursive(
        &self,
        style_id: &str,
        stack: &mut Vec<String>,
    ) -> Result<RunStyleProperties> {
        if self.is_implicit_default_run_style(style_id) {
            return Ok(RunStyleProperties::default());
        }
        if stack.iter().any(|entry| entry == style_id) {
            return Err(DocxError::parse(format!(
                "run style inheritance cycle detected at '{style_id}'"
            )));
        }
        let style = self
            .run_style(style_id)
            .ok_or_else(|| DocxError::parse(format!("unknown run style '{style_id}'")))?;
        stack.push(style_id.to_string());
        let mut resolved = if let Some(parent_id) = style.based_on.as_deref() {
            self.resolve_run_style_recursive(parent_id, stack)?
        } else {
            RunStyleProperties::default()
        };
        let _ = stack.pop();
        merge_run_style_properties(&mut resolved, &style.properties);
        Ok(resolved)
    }

    fn resolve_table_style_recursive(
        &self,
        style_id: &str,
        stack: &mut Vec<String>,
    ) -> Result<TableProperties> {
        if self.is_implicit_default_table_style(style_id) {
            return Ok(TableProperties::default());
        }
        if stack.iter().any(|entry| entry == style_id) {
            return Err(DocxError::parse(format!(
                "table style inheritance cycle detected at '{style_id}'"
            )));
        }
        let style = self
            .table_style(style_id)
            .ok_or_else(|| DocxError::parse(format!("unknown table style '{style_id}'")))?;
        stack.push(style_id.to_string());
        let mut resolved = if let Some(parent_id) = style.based_on.as_deref() {
            self.resolve_table_style_recursive(parent_id, stack)?
        } else {
            TableProperties::default()
        };
        let _ = stack.pop();
        apply_table_style_properties(&mut resolved, &style.properties);
        Ok(resolved)
    }

    fn is_implicit_default_paragraph_style(&self, style_id: &str) -> bool {
        style_id == DEFAULT_PARAGRAPH_STYLE_ID && self.paragraph_style(style_id).is_none()
    }

    fn is_implicit_default_run_style(&self, style_id: &str) -> bool {
        style_id == DEFAULT_RUN_STYLE_ID && self.run_style(style_id).is_none()
    }

    fn is_implicit_default_table_style(&self, style_id: &str) -> bool {
        style_id == DEFAULT_TABLE_STYLE_ID && self.table_style(style_id).is_none()
    }
}

/// A reusable paragraph style definition.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct ParagraphStyle {
    pub id: String,
    pub name: Option<String>,
    pub based_on: Option<String>,
    pub next: Option<String>,
    pub paragraph: ParagraphStyleProperties,
    pub run: RunStyleProperties,
}

impl ParagraphStyle {
    pub fn new(id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            ..Self::default()
        }
    }

    pub fn name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    pub fn based_on(mut self, style_id: impl Into<String>) -> Self {
        self.based_on = Some(style_id.into());
        self
    }

    pub fn next(mut self, style_id: impl Into<String>) -> Self {
        self.next = Some(style_id.into());
        self
    }

    pub fn paragraph(mut self, properties: ParagraphStyleProperties) -> Self {
        self.paragraph = properties;
        self
    }

    pub fn run(mut self, properties: RunStyleProperties) -> Self {
        self.run = properties;
        self
    }
}

/// A reusable run style definition.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct RunStyle {
    pub id: String,
    pub name: Option<String>,
    pub based_on: Option<String>,
    pub properties: RunStyleProperties,
}

impl RunStyle {
    pub fn new(id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            ..Self::default()
        }
    }

    pub fn name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    pub fn based_on(mut self, style_id: impl Into<String>) -> Self {
        self.based_on = Some(style_id.into());
        self
    }

    pub fn properties(mut self, properties: RunStyleProperties) -> Self {
        self.properties = properties;
        self
    }
}

/// A reusable table style definition.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct TableStyle {
    pub id: String,
    pub name: Option<String>,
    pub based_on: Option<String>,
    pub properties: TableStyleProperties,
}

impl TableStyle {
    pub fn new(id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            ..Self::default()
        }
    }

    pub fn name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    pub fn based_on(mut self, style_id: impl Into<String>) -> Self {
        self.based_on = Some(style_id.into());
        self
    }

    pub fn properties(mut self, properties: TableStyleProperties) -> Self {
        self.properties = properties;
        self
    }
}

/// Paragraph-level formatting that can participate in style inheritance.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct ParagraphStyleProperties {
    pub list: Option<ParagraphList>,
    pub alignment: Option<ParagraphAlignment>,
    pub spacing_before: Option<u32>,
    pub spacing_after: Option<u32>,
    pub keep_next: Option<bool>,
    pub page_break_before: Option<bool>,
}

impl ParagraphStyleProperties {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn list(mut self, list: ParagraphList) -> Self {
        self.list = Some(list);
        self
    }

    pub fn alignment(mut self, alignment: ParagraphAlignment) -> Self {
        self.alignment = Some(alignment);
        self
    }

    pub fn spacing_before(mut self, twips: u32) -> Self {
        self.spacing_before = Some(twips);
        self
    }

    pub fn spacing_after(mut self, twips: u32) -> Self {
        self.spacing_after = Some(twips);
        self
    }

    pub fn keep_next(mut self) -> Self {
        self.keep_next = Some(true);
        self
    }

    pub fn page_break_before(mut self) -> Self {
        self.page_break_before = Some(true);
        self
    }
}

/// Run-level formatting that can participate in style inheritance.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct RunStyleProperties {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<UnderlineStyle>,
    pub strikethrough: Option<bool>,
    pub small_caps: Option<bool>,
    pub shadow: Option<bool>,
    pub color: Option<String>,
    pub font_size: Option<u16>,
    pub font_family: Option<String>,
    pub vertical_align: Option<VerticalAlign>,
}

impl RunStyleProperties {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn bold(mut self) -> Self {
        self.bold = Some(true);
        self
    }

    pub fn italic(mut self) -> Self {
        self.italic = Some(true);
        self
    }

    pub fn underline(mut self, underline: UnderlineStyle) -> Self {
        self.underline = Some(underline);
        self
    }

    pub fn strikethrough(mut self) -> Self {
        self.strikethrough = Some(true);
        self
    }

    pub fn small_caps(mut self) -> Self {
        self.small_caps = Some(true);
        self
    }

    pub fn shadow(mut self) -> Self {
        self.shadow = Some(true);
        self
    }

    pub fn color(mut self, color: impl Into<String>) -> Self {
        self.color = Some(color.into());
        self
    }

    pub fn font(mut self, family: impl Into<String>) -> Self {
        self.font_family = Some(family.into());
        self
    }

    pub fn size_points(mut self, points: u16) -> Self {
        self.font_size = Some(points.saturating_mul(2));
        self
    }

    pub fn size_half_points(mut self, half_points: u16) -> Self {
        self.font_size = Some(half_points);
        self
    }

    pub fn vertical_align(mut self, vertical_align: VerticalAlign) -> Self {
        self.vertical_align = Some(vertical_align);
        self
    }

    pub fn superscript(mut self) -> Self {
        self.vertical_align = Some(VerticalAlign::Superscript);
        self
    }

    pub fn subscript(mut self) -> Self {
        self.vertical_align = Some(VerticalAlign::Subscript);
        self
    }
}

/// Table-level formatting that can participate in style inheritance.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct TableStyleProperties {
    pub width: Option<u32>,
    pub borders: Option<TableBorders>,
}

impl TableStyleProperties {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn width(mut self, width: u32) -> Self {
        self.width = Some(width);
        self
    }

    pub fn borders(mut self, borders: TableBorders) -> Self {
        self.borders = Some(borders);
        self
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct ResolvedParagraphStyle {
    pub(crate) list: Option<ParagraphList>,
    pub(crate) alignment: Option<ParagraphAlignment>,
    pub(crate) spacing_before: Option<u32>,
    pub(crate) spacing_after: Option<u32>,
    pub(crate) keep_next: bool,
    pub(crate) page_break_before: bool,
    pub(crate) run: RunProperties,
}

fn upsert_style<T>(styles: &mut Vec<T>, style: T)
where
    T: StyleDefinition,
{
    let id = style.id().to_string();
    if let Some(index) = styles.iter().position(|existing| existing.id() == id) {
        styles[index] = style;
    } else {
        styles.push(style);
    }
}

trait StyleDefinition {
    fn id(&self) -> &str;
}

impl StyleDefinition for ParagraphStyle {
    fn id(&self) -> &str {
        self.id.as_str()
    }
}

impl StyleDefinition for RunStyle {
    fn id(&self) -> &str {
        self.id.as_str()
    }
}

impl StyleDefinition for TableStyle {
    fn id(&self) -> &str {
        self.id.as_str()
    }
}

fn apply_paragraph_style_properties(
    resolved: &mut ResolvedParagraphStyle,
    properties: &ParagraphStyleProperties,
) {
    if let Some(list) = properties.list {
        resolved.list = Some(list);
    }
    if let Some(alignment) = properties.alignment.clone() {
        resolved.alignment = Some(alignment);
    }
    if let Some(spacing_before) = properties.spacing_before {
        resolved.spacing_before = Some(spacing_before);
    }
    if let Some(spacing_after) = properties.spacing_after {
        resolved.spacing_after = Some(spacing_after);
    }
    if let Some(keep_next) = properties.keep_next {
        resolved.keep_next = keep_next;
    }
    if let Some(page_break_before) = properties.page_break_before {
        resolved.page_break_before = page_break_before;
    }
}

fn apply_run_style_properties(resolved: &mut RunProperties, properties: &RunStyleProperties) {
    if let Some(bold) = properties.bold {
        resolved.bold = bold;
    }
    if let Some(italic) = properties.italic {
        resolved.italic = italic;
    }
    if let Some(underline) = properties.underline.clone() {
        resolved.underline = Some(underline);
    }
    if let Some(strikethrough) = properties.strikethrough {
        resolved.strikethrough = strikethrough;
    }
    if let Some(small_caps) = properties.small_caps {
        resolved.small_caps = small_caps;
    }
    if let Some(shadow) = properties.shadow {
        resolved.shadow = shadow;
    }
    if let Some(color) = properties.color.clone() {
        resolved.color = Some(color);
    }
    if let Some(font_size) = properties.font_size {
        resolved.font_size = Some(font_size);
    }
    if let Some(font_family) = properties.font_family.clone() {
        resolved.font_family = Some(font_family);
    }
    if let Some(vertical_align) = properties.vertical_align {
        resolved.vertical_align = Some(vertical_align);
    }
}

fn merge_run_style_properties(resolved: &mut RunStyleProperties, properties: &RunStyleProperties) {
    if properties.bold.is_some() {
        resolved.bold = properties.bold;
    }
    if properties.italic.is_some() {
        resolved.italic = properties.italic;
    }
    if properties.underline.is_some() {
        resolved.underline = properties.underline.clone();
    }
    if properties.strikethrough.is_some() {
        resolved.strikethrough = properties.strikethrough;
    }
    if properties.small_caps.is_some() {
        resolved.small_caps = properties.small_caps;
    }
    if properties.shadow.is_some() {
        resolved.shadow = properties.shadow;
    }
    if properties.color.is_some() {
        resolved.color = properties.color.clone();
    }
    if properties.font_size.is_some() {
        resolved.font_size = properties.font_size;
    }
    if properties.font_family.is_some() {
        resolved.font_family = properties.font_family.clone();
    }
    if properties.vertical_align.is_some() {
        resolved.vertical_align = properties.vertical_align;
    }
}

fn apply_direct_run_properties(resolved: &mut RunProperties, properties: &RunProperties) {
    if properties.bold {
        resolved.bold = true;
    }
    if properties.italic {
        resolved.italic = true;
    }
    if let Some(underline) = properties.underline.clone() {
        resolved.underline = Some(underline);
    }
    if properties.strikethrough {
        resolved.strikethrough = true;
    }
    if properties.small_caps {
        resolved.small_caps = true;
    }
    if properties.shadow {
        resolved.shadow = true;
    }
    if let Some(color) = properties.color.clone() {
        resolved.color = Some(color);
    }
    if let Some(font_size) = properties.font_size {
        resolved.font_size = Some(font_size);
    }
    if let Some(font_family) = properties.font_family.clone() {
        resolved.font_family = Some(font_family);
    }
    if let Some(vertical_align) = properties.vertical_align {
        resolved.vertical_align = Some(vertical_align);
    }
}

fn apply_table_style_properties(resolved: &mut TableProperties, properties: &TableStyleProperties) {
    if let Some(width) = properties.width {
        resolved.width = Some(width);
    }
    if let Some(borders) = properties.borders.clone() {
        resolved.borders = Some(borders);
    }
}

#[cfg(test)]
mod tests {
    use crate::{
        Border, BorderStyle, Paragraph, ParagraphAlignment, ParagraphList, Run, Table, TableBorders,
    };

    use super::{
        ParagraphStyle, ParagraphStyleProperties, RunStyle, RunStyleProperties, Stylesheet,
        TableStyle, TableStyleProperties,
    };

    #[test]
    fn paragraph_style_inheritance_merges_paragraph_and_run_properties() {
        let stylesheet = Stylesheet::new()
            .add_paragraph_style(
                ParagraphStyle::new("base")
                    .paragraph(
                        ParagraphStyleProperties::new()
                            .alignment(ParagraphAlignment::Center)
                            .spacing_after(180),
                    )
                    .run(RunStyleProperties::new().bold().color("334155")),
            )
            .add_paragraph_style(
                ParagraphStyle::new("child")
                    .based_on("base")
                    .paragraph(ParagraphStyleProperties::new().keep_next()),
            );

        let paragraph = Paragraph::new()
            .with_style("child")
            .add_run(Run::from_text("Styled"));
        let resolved = stylesheet
            .resolve_paragraph(&paragraph)
            .expect("resolve paragraph");
        let run = stylesheet
            .resolve_run(&paragraph, paragraph.runs().next().expect("run"))
            .expect("resolve run");

        assert_eq!(resolved.alignment, Some(ParagraphAlignment::Center));
        assert_eq!(resolved.spacing_after, Some(180));
        assert!(resolved.keep_next);
        assert!(run.bold);
        assert_eq!(run.color.as_deref(), Some("334155"));
    }

    #[test]
    fn run_style_inheritance_and_direct_formatting_compose() {
        let stylesheet = Stylesheet::new()
            .add_run_style(RunStyle::new("accent").properties(RunStyleProperties::new().italic()))
            .add_run_style(
                RunStyle::new("strong")
                    .based_on("accent")
                    .properties(RunStyleProperties::new().bold().color("0F172A")),
            );

        let paragraph = Paragraph::new();
        let run = Run::from_text("Hello")
            .with_style("strong")
            .underline(crate::UnderlineStyle::Single);
        let resolved = stylesheet
            .resolve_run(&paragraph, &run)
            .expect("resolve run");

        assert!(resolved.bold);
        assert!(resolved.italic);
        assert_eq!(resolved.color.as_deref(), Some("0F172A"));
        assert_eq!(resolved.underline, Some(crate::UnderlineStyle::Single));
    }

    #[test]
    fn table_style_inheritance_merges_width_and_borders() {
        let border = Border::new(BorderStyle::Single).color("CBD5E1");
        let stylesheet = Stylesheet::new()
            .add_table_style(
                TableStyle::new("base").properties(
                    TableStyleProperties::new().width(8_800).borders(
                        TableBorders::new()
                            .top(border.clone())
                            .bottom(border.clone()),
                    ),
                ),
            )
            .add_table_style(TableStyle::new("child").based_on("base"));

        let table = Table::new().style("child");
        let resolved = stylesheet.resolve_table(&table).expect("resolve table");

        assert_eq!(resolved.width, Some(8_800));
        assert_eq!(
            resolved
                .borders
                .as_ref()
                .and_then(|borders| borders.top.as_ref()),
            Some(&border)
        );
    }

    #[test]
    fn inheritance_cycles_return_errors() {
        let stylesheet = Stylesheet::new().add_paragraph_style(
            ParagraphStyle::new("loop")
                .based_on("loop")
                .paragraph(ParagraphStyleProperties::new().spacing_before(120)),
        );

        let error = stylesheet
            .resolve_paragraph_style(Some("loop"))
            .expect_err("cycle must fail");
        assert!(matches!(error, crate::DocxError::Parse(message) if message.contains("cycle")));
    }

    #[test]
    fn direct_paragraph_formatting_overrides_style_defaults() {
        let stylesheet = Stylesheet::new().add_paragraph_style(
            ParagraphStyle::new("body").paragraph(
                ParagraphStyleProperties::new()
                    .alignment(ParagraphAlignment::Justified)
                    .spacing_after(120),
            ),
        );
        let paragraph = Paragraph::new()
            .with_style("body")
            .with_alignment(ParagraphAlignment::Right)
            .spacing_after(360);

        let resolved = stylesheet.resolve_paragraph(&paragraph).expect("resolve");
        assert_eq!(resolved.alignment, Some(ParagraphAlignment::Right));
        assert_eq!(resolved.spacing_after, Some(360));
    }

    #[test]
    fn paragraph_style_can_inherit_lists() {
        let stylesheet = Stylesheet::new().add_paragraph_style(
            ParagraphStyle::new("bullets")
                .paragraph(ParagraphStyleProperties::new().list(ParagraphList::bullet_with_id(9))),
        );
        let paragraph = Paragraph::new().with_style("bullets");
        let resolved = stylesheet.resolve_paragraph(&paragraph).expect("resolve");
        assert_eq!(resolved.list, Some(ParagraphList::bullet_with_id(9)));
    }

    #[test]
    fn styles_can_inherit_from_implicit_builtin_defaults() {
        let stylesheet = Stylesheet::new()
            .add_paragraph_style(
                ParagraphStyle::new("lead").based_on("Normal").paragraph(
                    ParagraphStyleProperties::new()
                        .alignment(ParagraphAlignment::Center)
                        .spacing_after(180),
                ),
            )
            .add_run_style(
                RunStyle::new("accent")
                    .based_on("DefaultParagraphFont")
                    .properties(RunStyleProperties::new().italic().color("0F172A")),
            )
            .add_table_style(
                TableStyle::new("grid")
                    .based_on("TableNormal")
                    .properties(TableStyleProperties::new().width(9_360)),
            );

        let paragraph = Paragraph::new()
            .with_style("lead")
            .add_run(Run::from_text("Styled").with_style("accent"));
        let run = paragraph.runs().next().expect("run");
        let table = Table::new().style("grid");

        let resolved_paragraph = stylesheet.resolve_paragraph(&paragraph).expect("paragraph");
        let resolved_run = stylesheet.resolve_run(&paragraph, run).expect("run");
        let resolved_table = stylesheet.resolve_table(&table).expect("table");

        assert_eq!(
            resolved_paragraph.alignment,
            Some(ParagraphAlignment::Center)
        );
        assert_eq!(resolved_paragraph.spacing_after, Some(180));
        assert!(resolved_run.italic);
        assert_eq!(resolved_run.color.as_deref(), Some("0F172A"));
        assert_eq!(resolved_table.width, Some(9_360));
    }
}
