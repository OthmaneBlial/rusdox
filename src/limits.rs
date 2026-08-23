/// Resource ceilings applied to untrusted document, spec, and visual inputs.
///
/// The defaults are deliberately generous for normal business documents while
/// bounding memory amplification from ZIP/XML/YAML/image inputs. Callers that
/// intentionally process larger trusted files can pass a customized value to
/// the `*_with_limits` APIs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InputLimits {
    /// Maximum compressed DOCX archive size.
    pub max_docx_archive_bytes: u64,
    /// Maximum number of entries in a DOCX ZIP package.
    pub max_docx_entries: usize,
    /// Maximum uncompressed size of one DOCX ZIP entry.
    pub max_docx_entry_bytes: u64,
    /// Maximum total uncompressed size across a DOCX package.
    pub max_docx_total_bytes: u64,
    /// Maximum uncompressed-to-compressed ratio for one ZIP entry.
    pub max_zip_compression_ratio: u64,
    /// Maximum bytes accepted for one XML or relationships part.
    pub max_xml_bytes: u64,
    /// Maximum bytes accepted for a root spec or included YAML fragment.
    pub max_spec_bytes: u64,
    /// Maximum recursive YAML include depth.
    pub max_include_depth: usize,
    /// Maximum number of YAML includes expanded by one root spec.
    pub max_include_files: usize,
    /// Maximum source bytes for one PNG or JPEG visual.
    pub max_image_bytes: u64,
    /// Maximum source bytes for one SVG visual.
    pub max_svg_bytes: u64,
    /// Maximum decoded or target raster pixel count for one visual.
    pub max_image_pixels: u64,
}

impl Default for InputLimits {
    fn default() -> Self {
        Self {
            max_docx_archive_bytes: 64 * 1024 * 1024,
            max_docx_entries: 4_096,
            max_docx_entry_bytes: 64 * 1024 * 1024,
            max_docx_total_bytes: 256 * 1024 * 1024,
            max_zip_compression_ratio: 200,
            max_xml_bytes: 16 * 1024 * 1024,
            max_spec_bytes: 8 * 1024 * 1024,
            max_include_depth: 32,
            max_include_files: 128,
            max_image_bytes: 32 * 1024 * 1024,
            max_svg_bytes: 8 * 1024 * 1024,
            max_image_pixels: 64_000_000,
        }
    }
}
