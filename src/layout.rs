use serde::{Deserialize, Serialize};

use crate::paragraph::ParagraphAlignment;

/// Page setup values stored in OOXML twips.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct PageSetup {
    pub width_twips: u32,
    pub height_twips: u32,
    pub margin_top_twips: u32,
    pub margin_right_twips: u32,
    pub margin_bottom_twips: u32,
    pub margin_left_twips: u32,
    pub header_twips: u32,
    pub footer_twips: u32,
    pub gutter_twips: u32,
}

impl Default for PageSetup {
    fn default() -> Self {
        Self {
            width_twips: 12_240,
            height_twips: 15_840,
            margin_top_twips: 1_440,
            margin_right_twips: 1_440,
            margin_bottom_twips: 1_440,
            margin_left_twips: 1_440,
            header_twips: 720,
            footer_twips: 720,
            gutter_twips: 0,
        }
    }
}

impl PageSetup {
    /// Creates a page setup with explicit page size and default margins.
    pub fn new(width_twips: u32, height_twips: u32) -> Self {
        Self {
            width_twips,
            height_twips,
            ..Self::default()
        }
    }

    /// Sets page margins in twips.
    pub fn margins(
        mut self,
        top_twips: u32,
        right_twips: u32,
        bottom_twips: u32,
        left_twips: u32,
    ) -> Self {
        self.margin_top_twips = top_twips;
        self.margin_right_twips = right_twips;
        self.margin_bottom_twips = bottom_twips;
        self.margin_left_twips = left_twips;
        self
    }

    /// Sets the header/footer distance from the page edge in twips.
    pub fn header_footer_distances(mut self, header_twips: u32, footer_twips: u32) -> Self {
        self.header_twips = header_twips;
        self.footer_twips = footer_twips;
        self
    }

    /// Sets the gutter margin in twips.
    pub fn gutter(mut self, gutter_twips: u32) -> Self {
        self.gutter_twips = gutter_twips;
        self
    }
}

/// A simple header or footer paragraph template.
///
/// The `text` field supports `{page}` and `{pages}` placeholders, which are
/// emitted as Word field codes.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct HeaderFooter {
    pub text: String,
    pub alignment: ParagraphAlignment,
}

impl Default for HeaderFooter {
    fn default() -> Self {
        Self {
            text: String::new(),
            alignment: ParagraphAlignment::Left,
        }
    }
}

impl HeaderFooter {
    /// Creates a header or footer with left alignment by default.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Self::default()
        }
    }

    /// Sets paragraph alignment for the rendered header or footer.
    pub fn with_alignment(mut self, alignment: ParagraphAlignment) -> Self {
        self.alignment = alignment;
        self
    }
}

/// Page numbering configuration emitted into section properties.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct PageNumbering {
    pub start_at: Option<u32>,
    pub format: PageNumberFormat,
}

impl Default for PageNumbering {
    fn default() -> Self {
        Self {
            start_at: None,
            format: PageNumberFormat::Decimal,
        }
    }
}

impl PageNumbering {
    /// Creates numbering settings using the supplied format.
    pub fn new(format: PageNumberFormat) -> Self {
        Self {
            format,
            ..Self::default()
        }
    }

    /// Sets the starting page number for the section.
    pub fn start_at(mut self, start_at: u32) -> Self {
        self.start_at = Some(start_at);
        self
    }
}

/// Word page number formats exposed through the public API and spec.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum PageNumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
}

impl PageNumberFormat {
    pub(crate) fn from_xml(value: &str) -> Self {
        match value {
            "upperRoman" => Self::UpperRoman,
            "lowerRoman" => Self::LowerRoman,
            "upperLetter" => Self::UpperLetter,
            "lowerLetter" => Self::LowerLetter,
            _ => Self::Decimal,
        }
    }

    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Decimal => "decimal",
            Self::UpperRoman => "upperRoman",
            Self::LowerRoman => "lowerRoman",
            Self::UpperLetter => "upperLetter",
            Self::LowerLetter => "lowerLetter",
        }
    }
}
