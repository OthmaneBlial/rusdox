use std::fs;
use std::path::{Path, PathBuf};

use image::codecs::png::PngEncoder;
use image::imageops::FilterType;
use image::{ColorType, GenericImageView, ImageEncoder};
use resvg::tiny_skia::{Pixmap, Transform};
use resvg::usvg;

use crate::error::{DocxError, Result};
use crate::paragraph::ParagraphAlignment;

const SVG_FALLBACK_DPI: u32 = 144;
const TWIPS_PER_PIXEL_AT_96_DPI: u32 = 15;

/// The semantic role of a visual asset in the document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VisualKind {
    /// A general-purpose image or illustration.
    Image,
    /// A brand mark or company logo.
    Logo,
    /// A handwritten or typed signature image.
    Signature,
    /// A chart or infographic asset.
    Chart,
}

impl VisualKind {
    pub(crate) fn as_docx_name(self) -> &'static str {
        match self {
            Self::Image => "RusDox Image",
            Self::Logo => "RusDox Logo",
            Self::Signature => "RusDox Signature",
            Self::Chart => "RusDox Chart",
        }
    }

    pub(crate) fn from_docx_name(value: &str) -> Self {
        let lower = value.trim().to_ascii_lowercase();
        if lower.starts_with("rusdox logo") {
            Self::Logo
        } else if lower.starts_with("rusdox signature") {
            Self::Signature
        } else if lower.starts_with("rusdox chart") {
            Self::Chart
        } else {
            Self::Image
        }
    }

    fn default_alignment(self) -> ParagraphAlignment {
        match self {
            Self::Image | Self::Chart => ParagraphAlignment::Center,
            Self::Logo => ParagraphAlignment::Left,
            Self::Signature => ParagraphAlignment::Right,
        }
    }

    fn default_max_width_twips(self, content_width_twips: u32) -> u32 {
        match self {
            Self::Image | Self::Chart => content_width_twips,
            Self::Logo => content_width_twips.min(2_880),
            Self::Signature => content_width_twips.min(3_600),
        }
    }

    fn default_max_height_twips(self, content_height_twips: u32) -> u32 {
        match self {
            Self::Image => content_height_twips.min(6_480),
            Self::Chart => content_height_twips.min(5_760),
            Self::Logo => content_height_twips.min(1_440),
            Self::Signature => content_height_twips.min(1_080),
        }
    }
}

/// Supported on-disk or embedded visual formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VisualFormat {
    /// A PNG image.
    Png,
    /// A JPEG image.
    Jpeg,
    /// An SVG document.
    Svg,
}

impl VisualFormat {
    pub(crate) fn content_type(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Svg => "image/svg+xml",
        }
    }

    pub(crate) fn extension(self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpg",
            Self::Svg => "svg",
        }
    }

    pub(crate) fn from_path(path: &Path) -> Option<Self> {
        match path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase()
            .as_str()
        {
            "png" => Some(Self::Png),
            "jpg" | "jpeg" => Some(Self::Jpeg),
            "svg" => Some(Self::Svg),
            _ => None,
        }
    }

    pub(crate) fn guess(bytes: &[u8]) -> Option<Self> {
        if bytes.starts_with(b"\x89PNG\r\n\x1A\n") {
            Some(Self::Png)
        } else if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
            Some(Self::Jpeg)
        } else {
            let sample = String::from_utf8_lossy(&bytes[..bytes.len().min(256)]);
            sample.contains("<svg").then_some(Self::Svg)
        }
    }
}

/// A visual source loaded from a file path or embedded directly in memory.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum VisualSource {
    /// Resolve the visual at save/render time from the provided path.
    Path(PathBuf),
    /// Use bytes already stored in memory.
    Embedded {
        format: VisualFormat,
        bytes: Vec<u8>,
    },
}

/// Size constraints attached to a visual block.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct VisualSizing {
    width_twips: Option<u32>,
    height_twips: Option<u32>,
    max_width_twips: Option<u32>,
    max_height_twips: Option<u32>,
}

impl VisualSizing {
    /// Returns the explicit width in twips, if present.
    pub fn width_twips(&self) -> Option<u32> {
        self.width_twips
    }

    /// Returns the explicit height in twips, if present.
    pub fn height_twips(&self) -> Option<u32> {
        self.height_twips
    }

    /// Returns the maximum width in twips, if present.
    pub fn max_width_twips(&self) -> Option<u32> {
        self.max_width_twips
    }

    /// Returns the maximum height in twips, if present.
    pub fn max_height_twips(&self) -> Option<u32> {
        self.max_height_twips
    }
}

/// A top-level visual block rendered as an image in DOCX and PDF output.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Visual {
    kind: VisualKind,
    source: VisualSource,
    alt_text: Option<String>,
    alignment: ParagraphAlignment,
    sizing: VisualSizing,
}

impl Visual {
    /// Creates a generic image block from a path.
    pub fn from_path(path: impl Into<PathBuf>) -> Self {
        Self {
            kind: VisualKind::Image,
            source: VisualSource::Path(path.into()),
            alt_text: None,
            alignment: VisualKind::Image.default_alignment(),
            sizing: VisualSizing::default(),
        }
    }

    /// Creates a visual block from already-embedded bytes.
    pub fn from_bytes(bytes: Vec<u8>, format: VisualFormat) -> Self {
        Self {
            kind: VisualKind::Image,
            source: VisualSource::Embedded { format, bytes },
            alt_text: None,
            alignment: VisualKind::Image.default_alignment(),
            sizing: VisualSizing::default(),
        }
    }

    /// Creates a semantic logo block from a path.
    pub fn logo(path: impl Into<PathBuf>) -> Self {
        Self::from_path(path).with_kind(VisualKind::Logo)
    }

    /// Creates a semantic signature block from a path.
    pub fn signature(path: impl Into<PathBuf>) -> Self {
        Self::from_path(path).with_kind(VisualKind::Signature)
    }

    /// Creates a semantic chart block from a path.
    pub fn chart(path: impl Into<PathBuf>) -> Self {
        Self::from_path(path).with_kind(VisualKind::Chart)
    }

    /// Returns the semantic visual kind.
    pub fn kind(&self) -> VisualKind {
        self.kind
    }

    /// Sets the semantic visual kind.
    pub fn with_kind(mut self, kind: VisualKind) -> Self {
        self.kind = kind;
        if self.alignment == VisualKind::Image.default_alignment()
            || self.alignment == VisualKind::Logo.default_alignment()
            || self.alignment == VisualKind::Signature.default_alignment()
            || self.alignment == VisualKind::Chart.default_alignment()
        {
            self.alignment = kind.default_alignment();
        }
        self
    }

    /// Returns the underlying source reference.
    pub fn source(&self) -> &VisualSource {
        &self.source
    }

    /// Returns the visual alt text when present.
    pub fn alt_text(&self) -> Option<&str> {
        self.alt_text.as_deref()
    }

    /// Sets the visual alt text.
    pub fn alt_text_text(mut self, alt_text: impl Into<String>) -> Self {
        self.alt_text = Some(alt_text.into());
        self
    }

    /// Returns the paragraph alignment used to place this visual.
    pub fn alignment(&self) -> &ParagraphAlignment {
        &self.alignment
    }

    /// Sets the paragraph alignment used to place this visual.
    pub fn with_alignment(mut self, alignment: ParagraphAlignment) -> Self {
        self.alignment = alignment;
        self
    }

    /// Sets the explicit width in twips.
    pub fn width_twips(mut self, width_twips: u32) -> Self {
        self.sizing.width_twips = Some(width_twips);
        self
    }

    /// Sets the explicit height in twips.
    pub fn height_twips(mut self, height_twips: u32) -> Self {
        self.sizing.height_twips = Some(height_twips);
        self
    }

    /// Sets the maximum width in twips.
    pub fn max_width_twips(mut self, max_width_twips: u32) -> Self {
        self.sizing.max_width_twips = Some(max_width_twips);
        self
    }

    /// Sets the maximum height in twips.
    pub fn max_height_twips(mut self, max_height_twips: u32) -> Self {
        self.sizing.max_height_twips = Some(max_height_twips);
        self
    }

    /// Returns the configured size constraints.
    pub fn sizing(&self) -> &VisualSizing {
        &self.sizing
    }

    /// Returns the current source path when the visual is path-backed.
    pub fn source_path(&self) -> Option<&Path> {
        match &self.source {
            VisualSource::Path(path) => Some(path.as_path()),
            VisualSource::Embedded { .. } => None,
        }
    }

    pub(crate) fn docx_name(&self) -> &'static str {
        self.kind.as_docx_name()
    }

    pub(crate) fn resolved_dimensions_twips(
        &self,
        content_width_twips: u32,
        content_height_twips: u32,
    ) -> Result<(u32, u32)> {
        let (intrinsic_width_px, intrinsic_height_px) = self.intrinsic_dimensions()?;
        Ok(resolve_dimensions_from_intrinsic(
            self,
            pixels_to_twips(intrinsic_width_px),
            pixels_to_twips(intrinsic_height_px),
            content_width_twips,
            content_height_twips,
        ))
    }

    pub(crate) fn intrinsic_dimensions(&self) -> Result<(u32, u32)> {
        let loaded = load_visual_source(&self.source)?;
        intrinsic_dimensions_for_source(&loaded)
    }

    pub(crate) fn docx_media(
        &self,
        display_width_twips: u32,
        display_height_twips: u32,
    ) -> Result<(VisualFormat, Vec<u8>)> {
        let loaded = load_visual_source(&self.source)?;
        match loaded.format {
            VisualFormat::Png | VisualFormat::Jpeg => Ok((loaded.format, loaded.bytes)),
            VisualFormat::Svg => {
                let raster = rasterize_svg(
                    &loaded.bytes,
                    loaded.resources_dir.as_deref(),
                    twips_to_pixels_at_dpi(display_width_twips, SVG_FALLBACK_DPI),
                    twips_to_pixels_at_dpi(display_height_twips, SVG_FALLBACK_DPI),
                )?;
                Ok((VisualFormat::Png, encode_png(&raster)?))
            }
        }
    }

    pub(crate) fn pdf_raster(
        &self,
        display_width_twips: u32,
        display_height_twips: u32,
    ) -> Result<RasterizedVisual> {
        let loaded = load_visual_source(&self.source)?;
        let target_width = twips_to_pixels_at_dpi(display_width_twips, SVG_FALLBACK_DPI);
        let target_height = twips_to_pixels_at_dpi(display_height_twips, SVG_FALLBACK_DPI);

        match loaded.format {
            VisualFormat::Svg => rasterize_svg(
                &loaded.bytes,
                loaded.resources_dir.as_deref(),
                target_width,
                target_height,
            ),
            VisualFormat::Png | VisualFormat::Jpeg => {
                rasterize_raster_image(&loaded.bytes, target_width, target_height)
            }
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RasterizedVisual {
    pub(crate) width_px: u32,
    pub(crate) height_px: u32,
    pub(crate) rgba: Vec<u8>,
}

#[derive(Debug)]
struct LoadedVisualSource {
    format: VisualFormat,
    bytes: Vec<u8>,
    resources_dir: Option<PathBuf>,
}

fn load_visual_source(source: &VisualSource) -> Result<LoadedVisualSource> {
    match source {
        VisualSource::Path(path) => {
            let bytes = fs::read(path)?;
            let format = VisualFormat::from_path(path)
                .or_else(|| VisualFormat::guess(&bytes))
                .ok_or_else(|| {
                    DocxError::parse(format!(
                        "unsupported visual format for {} (expected PNG, JPEG, or SVG)",
                        path.display()
                    ))
                })?;
            Ok(LoadedVisualSource {
                format,
                bytes,
                resources_dir: path.parent().map(Path::to_path_buf),
            })
        }
        VisualSource::Embedded { format, bytes } => Ok(LoadedVisualSource {
            format: *format,
            bytes: bytes.clone(),
            resources_dir: None,
        }),
    }
}

fn intrinsic_dimensions_for_source(source: &LoadedVisualSource) -> Result<(u32, u32)> {
    match source.format {
        VisualFormat::Png | VisualFormat::Jpeg => {
            let image = image::load_from_memory(&source.bytes)
                .map_err(|error| DocxError::parse(format!("failed to decode visual: {error}")))?;
            Ok(image.dimensions())
        }
        VisualFormat::Svg => {
            let tree = parse_svg_tree(&source.bytes, source.resources_dir.as_deref())?;
            let size = tree.size().to_int_size();
            Ok((size.width(), size.height()))
        }
    }
}

fn parse_svg_tree(bytes: &[u8], resources_dir: Option<&Path>) -> Result<usvg::Tree> {
    let mut options = usvg::Options {
        resources_dir: resources_dir.map(Path::to_path_buf),
        ..usvg::Options::default()
    };
    options.fontdb_mut().load_system_fonts();
    usvg::Tree::from_data(bytes, &options)
        .map_err(|error| DocxError::parse(format!("failed to parse SVG visual: {error}")))
}

fn rasterize_svg(
    bytes: &[u8],
    resources_dir: Option<&Path>,
    width_px: u32,
    height_px: u32,
) -> Result<RasterizedVisual> {
    let tree = parse_svg_tree(bytes, resources_dir)?;
    let width_px = width_px.max(1);
    let height_px = height_px.max(1);
    let mut pixmap = Pixmap::new(width_px, height_px)
        .ok_or_else(|| DocxError::parse("failed to allocate SVG render surface"))?;
    let source_size = tree.size();
    let transform = Transform::from_scale(
        width_px as f32 / source_size.width(),
        height_px as f32 / source_size.height(),
    );
    resvg::render(&tree, transform, &mut pixmap.as_mut());
    Ok(RasterizedVisual {
        width_px,
        height_px,
        rgba: pixmap.data().to_vec(),
    })
}

fn rasterize_raster_image(bytes: &[u8], width_px: u32, height_px: u32) -> Result<RasterizedVisual> {
    let image = image::load_from_memory(bytes)
        .map_err(|error| DocxError::parse(format!("failed to decode visual: {error}")))?;
    let width_px = width_px.max(1);
    let height_px = height_px.max(1);
    let resized = if image.width() == width_px && image.height() == height_px {
        image
    } else {
        image.resize_exact(width_px, height_px, FilterType::Lanczos3)
    };
    Ok(RasterizedVisual {
        width_px,
        height_px,
        rgba: resized.to_rgba8().into_raw(),
    })
}

fn encode_png(raster: &RasterizedVisual) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    PngEncoder::new(&mut bytes)
        .write_image(
            &raster.rgba,
            raster.width_px,
            raster.height_px,
            ColorType::Rgba8.into(),
        )
        .map_err(|error| DocxError::parse(format!("failed to encode PNG visual: {error}")))?;
    Ok(bytes)
}

pub(crate) fn resolve_dimensions_from_intrinsic(
    visual: &Visual,
    intrinsic_width_twips: u32,
    intrinsic_height_twips: u32,
    content_width_twips: u32,
    content_height_twips: u32,
) -> (u32, u32) {
    let intrinsic_width_twips = intrinsic_width_twips.max(1);
    let intrinsic_height_twips = intrinsic_height_twips.max(1);

    let (mut width, mut height) = match (visual.sizing.width_twips, visual.sizing.height_twips) {
        (Some(width), Some(height)) => (width.max(1), height.max(1)),
        (Some(width), None) => (
            width.max(1),
            scale_dimension(width.max(1), intrinsic_height_twips, intrinsic_width_twips),
        ),
        (None, Some(height)) => (
            scale_dimension(height.max(1), intrinsic_width_twips, intrinsic_height_twips),
            height.max(1),
        ),
        (None, None) => (intrinsic_width_twips, intrinsic_height_twips),
    };

    let max_width = visual
        .sizing
        .max_width_twips
        .unwrap_or_else(|| visual.kind.default_max_width_twips(content_width_twips))
        .min(content_width_twips.max(1));
    let max_height = visual
        .sizing
        .max_height_twips
        .unwrap_or_else(|| visual.kind.default_max_height_twips(content_height_twips))
        .min(content_height_twips.max(1));

    if width > max_width || height > max_height {
        let width_ratio = max_width as f64 / width as f64;
        let height_ratio = max_height as f64 / height as f64;
        let scale = width_ratio.min(height_ratio);
        width = ((width as f64 * scale).round() as u32).max(1);
        height = ((height as f64 * scale).round() as u32).max(1);
    }

    (width.max(1), height.max(1))
}

fn scale_dimension(base: u32, numerator: u32, denominator: u32) -> u32 {
    if denominator == 0 {
        base.max(1)
    } else {
        ((u64::from(base) * u64::from(numerator) + u64::from(denominator / 2))
            / u64::from(denominator)) as u32
    }
    .max(1)
}

pub(crate) fn pixels_to_twips(pixels: u32) -> u32 {
    pixels.saturating_mul(TWIPS_PER_PIXEL_AT_96_DPI).max(1)
}

pub(crate) fn twips_to_pixels_at_dpi(twips: u32, dpi: u32) -> u32 {
    (u64::from(twips) * u64::from(dpi)).div_ceil(1_440) as u32
}

#[cfg(test)]
mod tests {
    use std::fs;

    use tempfile::tempdir;

    use super::{
        pixels_to_twips, resolve_dimensions_from_intrinsic, twips_to_pixels_at_dpi, Visual,
        VisualFormat, VisualKind,
    };
    use crate::ParagraphAlignment;

    const SIMPLE_SVG: &str = r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 40">
  <rect width="120" height="40" fill="#F8FAFC"/>
  <path d="M12 28 L28 12 L44 28" stroke="#0F766E" stroke-width="6" fill="none" stroke-linecap="round"/>
  <text x="52" y="26" font-size="16" fill="#0F172A">RusDox</text>
</svg>"##;

    #[test]
    fn visual_kind_defaults_are_stable() {
        assert_eq!(
            Visual::from_path("hero.png").alignment(),
            &ParagraphAlignment::Center
        );
        assert_eq!(
            Visual::logo("mark.svg").alignment(),
            &ParagraphAlignment::Left
        );
        assert_eq!(
            Visual::signature("sig.svg").alignment(),
            &ParagraphAlignment::Right
        );
        assert_eq!(
            Visual::chart("bench.svg").alignment(),
            &ParagraphAlignment::Center
        );
    }

    #[test]
    fn visual_dimensions_fit_within_kind_defaults() {
        let image = Visual::from_path("placeholder.png");
        let logo = Visual::logo("placeholder.svg");
        let signature = Visual::signature("placeholder.svg");
        let chart = Visual::chart("placeholder.svg");

        let image_size = resolve_dimensions_from_intrinsic(&image, 9_000, 4_500, 6_000, 10_000);
        let logo_size = resolve_dimensions_from_intrinsic(&logo, 9_000, 4_500, 6_000, 10_000);
        let signature_size =
            resolve_dimensions_from_intrinsic(&signature, 9_000, 4_500, 6_000, 10_000);
        let chart_size = resolve_dimensions_from_intrinsic(&chart, 9_000, 4_500, 6_000, 10_000);

        assert_eq!(image_size, (6_000, 3_000));
        assert!(logo_size.0 <= 2_880);
        assert!(logo_size.1 <= 1_440);
        assert!(signature_size.0 <= 3_600);
        assert!(signature_size.1 <= 1_080);
        assert_eq!(chart_size, (6_000, 3_000));
    }

    #[test]
    fn explicit_visual_dimension_preserves_aspect_ratio() {
        let visual = Visual::from_path("photo.png").width_twips(2_400);
        let size = resolve_dimensions_from_intrinsic(&visual, 4_800, 1_600, 8_000, 10_000);
        assert_eq!(size, (2_400, 800));
    }

    #[test]
    fn pixel_twip_conversions_match_expected_document_units() {
        assert_eq!(pixels_to_twips(96), 1_440);
        assert_eq!(twips_to_pixels_at_dpi(1_440, 144), 144);
        assert_eq!(twips_to_pixels_at_dpi(2_880, 144), 288);
    }

    #[test]
    fn svg_visual_supports_intrinsic_dimensions_and_rasterization() {
        let visual = Visual::from_bytes(SIMPLE_SVG.as_bytes().to_vec(), VisualFormat::Svg)
            .with_kind(VisualKind::Logo);
        let intrinsic = visual.intrinsic_dimensions().expect("svg dimensions");
        assert_eq!(intrinsic, (120, 40));

        let raster = visual.pdf_raster(2_400, 800).expect("svg raster");
        assert_eq!(raster.width_px, 240);
        assert_eq!(raster.height_px, 80);
        assert_eq!(raster.rgba.len(), 240 * 80 * 4);
    }

    #[test]
    fn docx_media_rasterizes_svg_to_png() {
        let visual = Visual::from_bytes(SIMPLE_SVG.as_bytes().to_vec(), VisualFormat::Svg)
            .with_kind(VisualKind::Chart);
        let (format, bytes) = visual.docx_media(4_800, 1_600).expect("docx media");
        assert_eq!(format, VisualFormat::Png);
        assert!(bytes.starts_with(b"\x89PNG\r\n\x1A\n"));
    }

    #[test]
    fn path_backed_visual_uses_relative_svg_resources_directory() {
        let temp = tempdir().expect("temp dir");
        let svg_path = temp.path().join("logo.svg");
        fs::write(&svg_path, SIMPLE_SVG).expect("write svg");

        let visual = Visual::logo(&svg_path);
        let intrinsic = visual.intrinsic_dimensions().expect("svg dimensions");
        assert_eq!(intrinsic, (120, 40));
    }
}
