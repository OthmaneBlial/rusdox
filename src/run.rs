/// Supported underline styles for a text run.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum UnderlineStyle {
    /// A single underline.
    Single,
    /// A double underline.
    Double,
    /// A dotted underline.
    Dotted,
    /// A dashed underline.
    Dash,
    /// A wavy underline.
    Wavy,
    /// Underline words but not spaces.
    Words,
    /// Explicitly disable underline.
    None,
    /// Preserve or emit a custom OOXML underline value.
    Custom(String),
}

impl UnderlineStyle {
    pub(crate) fn from_xml(value: &str) -> Self {
        match value {
            "single" => Self::Single,
            "double" => Self::Double,
            "dotted" => Self::Dotted,
            "dash" => Self::Dash,
            "wave" => Self::Wavy,
            "words" => Self::Words,
            "none" => Self::None,
            other => Self::Custom(other.to_string()),
        }
    }

    pub(crate) fn as_xml_value(&self) -> &str {
        match self {
            Self::Single => "single",
            Self::Double => "double",
            Self::Dotted => "dotted",
            Self::Dash => "dash",
            Self::Wavy => "wave",
            Self::Words => "words",
            Self::None => "none",
            Self::Custom(value) => value.as_str(),
        }
    }
}

/// Vertical alignment for a run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlign {
    /// Superscript text.
    Superscript,
    /// Subscript text.
    Subscript,
    /// Baseline text.
    Baseline,
}

impl VerticalAlign {
    pub(crate) fn from_xml(value: &str) -> Self {
        match value {
            "superscript" => Self::Superscript,
            "subscript" => Self::Subscript,
            _ => Self::Baseline,
        }
    }

    pub(crate) fn as_xml_value(&self) -> &str {
        match self {
            Self::Superscript => "superscript",
            Self::Subscript => "subscript",
            Self::Baseline => "baseline",
        }
    }
}

/// Typed formatting properties for a text run.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RunProperties {
    /// Whether the run is bold.
    pub bold: bool,
    /// Whether the run is italic.
    pub italic: bool,
    /// Underline configuration.
    pub underline: Option<UnderlineStyle>,
    /// Whether the run is struck through.
    pub strikethrough: bool,
    /// Whether the run uses small caps.
    pub small_caps: bool,
    /// Whether the run has a shadow effect.
    pub shadow: bool,
    /// Text color in hexadecimal RGB form.
    pub color: Option<String>,
    /// Text size in OOXML half-point units.
    pub font_size: Option<u16>,
    /// Typeface to use for the run.
    pub font_family: Option<String>,
    /// Optional vertical alignment.
    pub vertical_align: Option<VerticalAlign>,
}

impl RunProperties {
    pub(crate) fn has_serialized_content(&self) -> bool {
        self.bold
            || self.italic
            || self.underline.is_some()
            || self.strikethrough
            || self.small_caps
            || self.shadow
            || self.color.is_some()
            || self.font_size.is_some()
            || self.font_family.is_some()
            || self.vertical_align.is_some()
    }
}

/// An inline text fragment in a paragraph.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Run {
    text: String,
    properties: RunProperties,
}

impl Run {
    /// Creates an empty run.
    ///
    /// ```rust
    /// use rusdox::Run;
    ///
    /// let run = Run::new();
    /// assert_eq!(run.text(), "");
    /// ```
    pub fn new() -> Self {
        Self::default()
    }

    /// Creates a run with text content.
    ///
    /// ```rust
    /// use rusdox::Run;
    ///
    /// let run = Run::from_text("Hello");
    /// assert_eq!(run.text(), "Hello");
    /// ```
    pub fn from_text(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            properties: RunProperties::default(),
        }
    }

    /// Sets the run text in a builder-friendly way.
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = text.into();
        self
    }

    /// Returns the run text.
    pub fn text(&self) -> &str {
        self.text.as_str()
    }

    /// Replaces the run text in place.
    pub fn set_text(&mut self, text: impl Into<String>) -> &mut Self {
        self.text = text.into();
        self
    }

    /// Returns the run properties.
    pub fn properties(&self) -> &RunProperties {
        &self.properties
    }

    /// Returns mutable access to the run properties.
    pub fn properties_mut(&mut self) -> &mut RunProperties {
        &mut self.properties
    }

    /// Enables bold formatting.
    pub fn bold(mut self) -> Self {
        self.properties.bold = true;
        self
    }

    /// Enables italic formatting.
    pub fn italic(mut self) -> Self {
        self.properties.italic = true;
        self
    }

    /// Applies an underline style.
    pub fn underline(mut self, style: UnderlineStyle) -> Self {
        self.properties.underline = Some(style);
        self
    }

    /// Enables strikethrough formatting.
    pub fn strikethrough(mut self) -> Self {
        self.properties.strikethrough = true;
        self
    }

    /// Enables small caps formatting.
    pub fn small_caps(mut self) -> Self {
        self.properties.small_caps = true;
        self
    }

    /// Enables a shadow effect.
    pub fn shadow(mut self) -> Self {
        self.properties.shadow = true;
        self
    }

    /// Sets an RGB text color.
    pub fn color(mut self, value: impl Into<String>) -> Self {
        self.properties.color = Some(value.into());
        self
    }

    /// Sets the run font family.
    pub fn font(mut self, value: impl Into<String>) -> Self {
        self.properties.font_family = Some(value.into());
        self
    }

    /// Sets the run font size in points.
    pub fn size_points(mut self, points: u16) -> Self {
        self.properties.font_size = Some(points.saturating_mul(2));
        self
    }

    /// Sets the run font size using raw OOXML half-point units.
    pub fn size_half_points(mut self, half_points: u16) -> Self {
        self.properties.font_size = Some(half_points);
        self
    }

    /// Sets the run to superscript.
    pub fn superscript(mut self) -> Self {
        self.properties.vertical_align = Some(VerticalAlign::Superscript);
        self
    }

    /// Sets the run to subscript.
    pub fn subscript(mut self) -> Self {
        self.properties.vertical_align = Some(VerticalAlign::Subscript);
        self
    }

    pub(crate) fn from_parts(text: String, properties: RunProperties) -> Self {
        Self { text, properties }
    }

    pub(crate) fn needs_space_preserve(&self) -> bool {
        let starts_with_ws = self.text.chars().next().is_some_and(char::is_whitespace);
        let ends_with_ws = self.text.chars().last().is_some_and(char::is_whitespace);
        starts_with_ws || ends_with_ws
    }
}

#[cfg(test)]
mod tests {
    use super::{Run, RunProperties, UnderlineStyle, VerticalAlign};

    #[test]
    fn underline_style_round_trips_known_values() {
        let cases = [
            ("single", UnderlineStyle::Single),
            ("double", UnderlineStyle::Double),
            ("dotted", UnderlineStyle::Dotted),
            ("dash", UnderlineStyle::Dash),
            ("wave", UnderlineStyle::Wavy),
            ("words", UnderlineStyle::Words),
            ("none", UnderlineStyle::None),
        ];

        for (xml, expected) in cases {
            let parsed = UnderlineStyle::from_xml(xml);
            assert_eq!(parsed, expected);
            assert_eq!(parsed.as_xml_value(), xml);
        }
    }

    #[test]
    fn underline_style_preserves_custom_xml_value() {
        let parsed = UnderlineStyle::from_xml("thick");
        assert_eq!(parsed, UnderlineStyle::Custom("thick".to_string()));
        assert_eq!(parsed.as_xml_value(), "thick");
    }

    #[test]
    fn vertical_align_round_trips_known_values() {
        let superscript = VerticalAlign::from_xml("superscript");
        let subscript = VerticalAlign::from_xml("subscript");
        let baseline = VerticalAlign::from_xml("anything-else");

        assert_eq!(superscript, VerticalAlign::Superscript);
        assert_eq!(subscript, VerticalAlign::Subscript);
        assert_eq!(baseline, VerticalAlign::Baseline);

        assert_eq!(superscript.as_xml_value(), "superscript");
        assert_eq!(subscript.as_xml_value(), "subscript");
        assert_eq!(baseline.as_xml_value(), "baseline");
    }

    #[test]
    fn run_properties_serialization_flag_tracks_all_fields() {
        let mut properties = RunProperties::default();
        assert!(!properties.has_serialized_content());

        properties.bold = true;
        assert!(properties.has_serialized_content());
        properties.bold = false;

        properties.italic = true;
        assert!(properties.has_serialized_content());
        properties.italic = false;

        properties.underline = Some(UnderlineStyle::Single);
        assert!(properties.has_serialized_content());
        properties.underline = None;

        properties.strikethrough = true;
        assert!(properties.has_serialized_content());
        properties.strikethrough = false;

        properties.small_caps = true;
        assert!(properties.has_serialized_content());
        properties.small_caps = false;

        properties.shadow = true;
        assert!(properties.has_serialized_content());
        properties.shadow = false;

        properties.color = Some("FF0000".to_string());
        assert!(properties.has_serialized_content());
        properties.color = None;

        properties.font_size = Some(24);
        assert!(properties.has_serialized_content());
        properties.font_size = None;

        properties.font_family = Some("Arial".to_string());
        assert!(properties.has_serialized_content());
        properties.font_family = None;

        properties.vertical_align = Some(VerticalAlign::Superscript);
        assert!(properties.has_serialized_content());
    }

    #[test]
    fn run_builder_methods_apply_expected_properties() {
        let run = Run::from_text("x")
            .bold()
            .italic()
            .underline(UnderlineStyle::Double)
            .strikethrough()
            .small_caps()
            .shadow()
            .color("AABBCC")
            .font("Inter")
            .size_points(12)
            .superscript();

        assert_eq!(run.text(), "x");
        assert!(run.properties().bold);
        assert!(run.properties().italic);
        assert_eq!(run.properties().underline, Some(UnderlineStyle::Double));
        assert!(run.properties().strikethrough);
        assert!(run.properties().small_caps);
        assert!(run.properties().shadow);
        assert_eq!(run.properties().color.as_deref(), Some("AABBCC"));
        assert_eq!(run.properties().font_family.as_deref(), Some("Inter"));
        assert_eq!(run.properties().font_size, Some(24));
        assert_eq!(
            run.properties().vertical_align,
            Some(VerticalAlign::Superscript)
        );
    }

    #[test]
    fn size_points_saturates_and_half_points_is_exact() {
        let saturated = Run::new().size_points(u16::MAX);
        let exact = Run::new().size_half_points(65530);
        assert_eq!(saturated.properties().font_size, Some(u16::MAX));
        assert_eq!(exact.properties().font_size, Some(65530));
    }

    #[test]
    fn set_text_and_properties_mut_work_in_place() {
        let mut run = Run::new().with_text("first");
        run.set_text("second");
        run.properties_mut().bold = true;
        run.properties_mut().font_family = Some("Arial".to_string());

        assert_eq!(run.text(), "second");
        assert!(run.properties().bold);
        assert_eq!(run.properties().font_family.as_deref(), Some("Arial"));
    }

    #[test]
    fn needs_space_preserve_detects_boundary_whitespace() {
        assert!(Run::from_text(" leading").needs_space_preserve());
        assert!(Run::from_text("trailing ").needs_space_preserve());
        assert!(Run::from_text("\nline").needs_space_preserve());
        assert!(!Run::from_text("in ter nal").needs_space_preserve());
        assert!(!Run::from_text("plain").needs_space_preserve());
    }

    #[test]
    fn subscript_overwrites_previous_vertical_alignment() {
        let run = Run::from_text("chem").superscript().subscript();
        assert_eq!(
            run.properties().vertical_align,
            Some(VerticalAlign::Subscript)
        );
    }

    #[test]
    fn from_parts_preserves_text_and_properties() {
        let properties = RunProperties {
            bold: true,
            italic: true,
            underline: Some(UnderlineStyle::Single),
            strikethrough: false,
            small_caps: true,
            shadow: false,
            color: Some("112233".to_string()),
            font_size: Some(28),
            font_family: Some("Calibri".to_string()),
            vertical_align: Some(VerticalAlign::Baseline),
        };

        let run = Run::from_parts("abc".to_string(), properties.clone());
        assert_eq!(run.text(), "abc");
        assert_eq!(run.properties(), &properties);
    }
}
