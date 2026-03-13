use serde::{Deserialize, Serialize};

use crate::run::Run;

const MAX_LIST_LEVEL: u8 = 8;
const DEFAULT_BULLET_LIST_ID: u32 = 1;
const DEFAULT_NUMBERED_LIST_ID: u32 = 2;

/// Paragraph alignment options.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ParagraphAlignment {
    /// Left aligned text.
    Left,
    /// Center aligned text.
    Center,
    /// Right aligned text.
    Right,
    /// Justified text.
    Justified,
    /// Preserve or emit a custom OOXML value.
    Custom(String),
}

impl ParagraphAlignment {
    pub(crate) fn from_xml(value: &str) -> Self {
        match value {
            "left" => Self::Left,
            "center" => Self::Center,
            "right" => Self::Right,
            "both" | "justify" => Self::Justified,
            other => Self::Custom(other.to_string()),
        }
    }

    pub(crate) fn as_xml_value(&self) -> &str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Justified => "both",
            Self::Custom(value) => value.as_str(),
        }
    }
}

/// Semantic paragraph list kinds supported by the DOCX writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphListKind {
    /// Bulleted list formatting.
    Bullet,
    /// Decimal numbered list formatting.
    Decimal,
}

impl ParagraphListKind {
    pub(crate) fn from_number_format(value: &str) -> Option<Self> {
        match value {
            "bullet" => Some(Self::Bullet),
            "decimal" => Some(Self::Decimal),
            _ => None,
        }
    }

    pub(crate) fn as_number_format(self) -> &'static str {
        match self {
            Self::Bullet => "bullet",
            Self::Decimal => "decimal",
        }
    }
}

/// Semantic DOCX numbering metadata for a paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParagraphList {
    kind: ParagraphListKind,
    level: u8,
    id: u32,
}

impl ParagraphList {
    /// Creates a top-level bullet list item using the default shared list id.
    pub fn bullet() -> Self {
        Self::bullet_with_id(DEFAULT_BULLET_LIST_ID)
    }

    /// Creates a top-level bullet list item with an explicit list id.
    pub fn bullet_with_id(id: u32) -> Self {
        Self {
            kind: ParagraphListKind::Bullet,
            level: 0,
            id: sanitize_list_id(id),
        }
    }

    /// Creates a top-level decimal numbered list item using the default shared list id.
    pub fn numbered() -> Self {
        Self::numbered_with_id(DEFAULT_NUMBERED_LIST_ID)
    }

    /// Creates a top-level decimal numbered list item with an explicit list id.
    pub fn numbered_with_id(id: u32) -> Self {
        Self {
            kind: ParagraphListKind::Decimal,
            level: 0,
            id: sanitize_list_id(id),
        }
    }

    /// Returns the list kind.
    pub fn kind(&self) -> ParagraphListKind {
        self.kind
    }

    /// Returns the nesting level.
    pub fn level(&self) -> u8 {
        self.level
    }

    /// Returns the logical list id.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Sets the nesting level, clamped to the OOXML-supported range `0..=8`.
    pub fn with_level(mut self, level: u8) -> Self {
        self.level = level.min(MAX_LIST_LEVEL);
        self
    }

    pub(crate) fn from_parts(kind: ParagraphListKind, id: u32, level: u8) -> Self {
        Self {
            kind,
            level: level.min(MAX_LIST_LEVEL),
            id: sanitize_list_id(id),
        }
    }
}

fn sanitize_list_id(id: u32) -> u32 {
    id.max(1)
}

/// A block-level text container made up of one or more runs.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Paragraph {
    runs: Vec<Run>,
    list: Option<ParagraphList>,
    alignment: Option<ParagraphAlignment>,
    spacing_before: Option<u32>,
    spacing_after: Option<u32>,
    page_break_before: bool,
}

impl Paragraph {
    /// Creates an empty paragraph.
    ///
    /// ```rust
    /// use rusdox::Paragraph;
    ///
    /// let paragraph = Paragraph::new();
    /// assert!(paragraph.runs().next().is_none());
    /// ```
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds a run using a builder-style API.
    pub fn add_run(mut self, run: Run) -> Self {
        self.runs.push(run);
        self
    }

    /// Appends a run in place.
    pub fn push_run(&mut self, run: Run) -> &mut Self {
        self.runs.push(run);
        self
    }

    /// Returns immutable access to the paragraph runs.
    pub fn runs(&self) -> std::slice::Iter<'_, Run> {
        self.runs.iter()
    }

    /// Returns mutable access to the paragraph runs.
    pub fn runs_mut(&mut self) -> std::slice::IterMut<'_, Run> {
        self.runs.iter_mut()
    }

    /// Returns the semantic list metadata, if present.
    pub fn list(&self) -> Option<&ParagraphList> {
        self.list.as_ref()
    }

    /// Sets semantic list metadata in a builder-friendly way.
    pub fn with_list(mut self, list: ParagraphList) -> Self {
        self.list = Some(list);
        self
    }

    /// Sets semantic list metadata in place.
    pub fn set_list(&mut self, list: ParagraphList) -> &mut Self {
        self.list = Some(list);
        self
    }

    /// Removes semantic list metadata in place.
    pub fn clear_list(&mut self) -> &mut Self {
        self.list = None;
        self
    }

    /// Returns the paragraph alignment, if present.
    pub fn alignment(&self) -> Option<&ParagraphAlignment> {
        self.alignment.as_ref()
    }

    /// Sets paragraph alignment in a builder-friendly way.
    pub fn with_alignment(mut self, alignment: ParagraphAlignment) -> Self {
        self.alignment = Some(alignment);
        self
    }

    /// Sets paragraph alignment in place.
    pub fn set_alignment(&mut self, alignment: ParagraphAlignment) -> &mut Self {
        self.alignment = Some(alignment);
        self
    }

    /// Returns the concatenated plain text of all runs.
    pub fn text(&self) -> String {
        self.runs.iter().map(Run::text).collect()
    }

    /// Sets spacing before the paragraph in twips.
    pub fn spacing_before(mut self, twips: u32) -> Self {
        self.spacing_before = Some(twips);
        self
    }

    /// Sets spacing after the paragraph in twips.
    pub fn spacing_after(mut self, twips: u32) -> Self {
        self.spacing_after = Some(twips);
        self
    }

    /// Forces the paragraph to start on a new page.
    pub fn page_break_before(mut self) -> Self {
        self.page_break_before = true;
        self
    }

    /// Returns spacing before the paragraph, if present.
    pub fn spacing_before_value(&self) -> Option<u32> {
        self.spacing_before
    }

    /// Returns spacing after the paragraph, if present.
    pub fn spacing_after_value(&self) -> Option<u32> {
        self.spacing_after
    }

    /// Returns whether the paragraph starts on a new page.
    pub fn has_page_break_before(&self) -> bool {
        self.page_break_before
    }

    pub(crate) fn from_parts(
        runs: Vec<Run>,
        list: Option<ParagraphList>,
        alignment: Option<ParagraphAlignment>,
        spacing_before: Option<u32>,
        spacing_after: Option<u32>,
        page_break_before: bool,
    ) -> Self {
        Self {
            runs,
            list,
            alignment,
            spacing_before,
            spacing_after,
            page_break_before,
        }
    }

    pub(crate) fn has_properties(&self) -> bool {
        self.list.is_some()
            || self.alignment.is_some()
            || self.spacing_before.is_some()
            || self.spacing_after.is_some()
            || self.page_break_before
    }
}

#[cfg(test)]
mod tests {
    use super::{Paragraph, ParagraphAlignment, ParagraphList, ParagraphListKind};
    use crate::Run;

    #[test]
    fn alignment_round_trips_known_values() {
        let cases = [
            ("left", ParagraphAlignment::Left, "left"),
            ("center", ParagraphAlignment::Center, "center"),
            ("right", ParagraphAlignment::Right, "right"),
            ("both", ParagraphAlignment::Justified, "both"),
            ("justify", ParagraphAlignment::Justified, "both"),
        ];

        for (xml, expected, roundtrip_xml) in cases {
            let parsed = ParagraphAlignment::from_xml(xml);
            assert_eq!(parsed, expected);
            assert_eq!(parsed.as_xml_value(), roundtrip_xml);
        }
    }

    #[test]
    fn alignment_custom_value_is_preserved() {
        let parsed = ParagraphAlignment::from_xml("distribute");
        assert_eq!(parsed, ParagraphAlignment::Custom("distribute".to_string()));
        assert_eq!(parsed.as_xml_value(), "distribute");
    }

    #[test]
    fn add_and_push_run_preserve_order() {
        let mut paragraph = Paragraph::new().add_run(Run::from_text("A"));
        paragraph.push_run(Run::from_text("B"));
        paragraph.push_run(Run::from_text("C"));

        let texts: Vec<_> = paragraph.runs().map(Run::text).collect();
        assert_eq!(texts, vec!["A", "B", "C"]);
        assert_eq!(paragraph.text(), "ABC");
    }

    #[test]
    fn runs_mut_allows_in_place_run_updates() {
        let mut paragraph = Paragraph::new()
            .add_run(Run::from_text("Hello"))
            .add_run(Run::from_text("World"));

        for run in paragraph.runs_mut() {
            if run.text() == "World" {
                run.set_text("RusDox");
            }
        }

        assert_eq!(paragraph.text(), "HelloRusDox");
    }

    #[test]
    fn builder_sets_spacing_alignment_and_page_break() {
        let paragraph = Paragraph::new()
            .with_list(ParagraphList::bullet().with_level(2))
            .with_alignment(ParagraphAlignment::Center)
            .spacing_before(120)
            .spacing_after(240)
            .page_break_before()
            .add_run(Run::from_text("x"));

        assert_eq!(
            paragraph.list(),
            Some(&ParagraphList::bullet().with_level(2))
        );
        assert_eq!(paragraph.alignment(), Some(&ParagraphAlignment::Center));
        assert_eq!(paragraph.spacing_before_value(), Some(120));
        assert_eq!(paragraph.spacing_after_value(), Some(240));
        assert!(paragraph.has_page_break_before());
    }

    #[test]
    fn set_alignment_overwrites_existing_alignment() {
        let mut paragraph = Paragraph::new().with_alignment(ParagraphAlignment::Left);
        paragraph.set_alignment(ParagraphAlignment::Right);
        assert_eq!(paragraph.alignment(), Some(&ParagraphAlignment::Right));
    }

    #[test]
    fn has_properties_detects_any_non_default_property() {
        assert!(!Paragraph::new().has_properties());
        assert!(Paragraph::new()
            .with_list(ParagraphList::bullet())
            .has_properties());
        assert!(Paragraph::new()
            .with_alignment(ParagraphAlignment::Center)
            .has_properties());
        assert!(Paragraph::new().spacing_before(100).has_properties());
        assert!(Paragraph::new().spacing_after(100).has_properties());
        assert!(Paragraph::new().page_break_before().has_properties());
    }

    #[test]
    fn from_parts_constructs_full_paragraph_state() {
        let runs = vec![Run::from_text("one"), Run::from_text("two")];
        let paragraph = Paragraph::from_parts(
            runs,
            Some(ParagraphList::numbered_with_id(9).with_level(1)),
            Some(ParagraphAlignment::Justified),
            Some(160),
            Some(180),
            true,
        );

        assert_eq!(paragraph.text(), "onetwo");
        assert_eq!(
            paragraph.list(),
            Some(&ParagraphList::numbered_with_id(9).with_level(1))
        );
        assert_eq!(paragraph.alignment(), Some(&ParagraphAlignment::Justified));
        assert_eq!(paragraph.spacing_before_value(), Some(160));
        assert_eq!(paragraph.spacing_after_value(), Some(180));
        assert!(paragraph.has_page_break_before());
    }

    #[test]
    fn list_builders_clamp_invalid_ids_and_levels() {
        let bullet = ParagraphList::bullet_with_id(0).with_level(99);
        let numbered = ParagraphList::numbered_with_id(7).with_level(3);

        assert_eq!(bullet.kind(), ParagraphListKind::Bullet);
        assert_eq!(bullet.id(), 1);
        assert_eq!(bullet.level(), 8);
        assert_eq!(numbered.kind(), ParagraphListKind::Decimal);
        assert_eq!(numbered.id(), 7);
        assert_eq!(numbered.level(), 3);
    }

    #[test]
    fn clear_list_removes_semantic_numbering() {
        let mut paragraph = Paragraph::new().with_list(ParagraphList::bullet());
        assert!(paragraph.list().is_some());
        paragraph.clear_list();
        assert!(paragraph.list().is_none());
    }
}
