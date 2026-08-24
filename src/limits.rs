use std::fs;
use std::path::Path;

use serde::{Deserialize, Serialize};

use crate::{DocxError, Result};

/// Resource ceilings applied to untrusted document, spec, and visual inputs.
///
/// The defaults are deliberately generous for normal business documents while
/// bounding memory amplification from ZIP/XML/YAML/image inputs. Callers that
/// intentionally process larger trusted files can pass a customized value to
/// the `*_with_limits` APIs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
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
    /// Maximum placeholder, loop, condition, and partial expansions per template render.
    pub max_template_expansions: usize,
    /// Maximum rendered bytes for one textual OOXML template part.
    pub max_template_output_xml_bytes: u64,
    /// Maximum nested partial expansion depth.
    pub max_template_partial_depth: usize,
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
            max_template_expansions: 100_000,
            max_template_output_xml_bytes: 64 * 1024 * 1024,
            max_template_partial_depth: 32,
        }
    }
}

impl InputLimits {
    /// Returns conservative ceilings for a multi-tenant or local service.
    pub const fn hosted() -> Self {
        Self {
            max_docx_archive_bytes: 16 * 1024 * 1024,
            max_docx_entries: 1_024,
            max_docx_entry_bytes: 16 * 1024 * 1024,
            max_docx_total_bytes: 64 * 1024 * 1024,
            max_zip_compression_ratio: 100,
            max_xml_bytes: 4 * 1024 * 1024,
            max_spec_bytes: 2 * 1024 * 1024,
            max_include_depth: 12,
            max_include_files: 32,
            max_image_bytes: 8 * 1024 * 1024,
            max_svg_bytes: 2 * 1024 * 1024,
            max_image_pixels: 16_000_000,
            max_template_expansions: 10_000,
            max_template_output_xml_bytes: 8 * 1024 * 1024,
            max_template_partial_depth: 12,
        }
    }

    /// Loads a complete TOML or JSON limit profile and rejects unknown fields.
    pub fn load_from_path(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let content = fs::read_to_string(path)?;
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or("toml")
            .to_ascii_lowercase();
        let limits: Self = match extension.as_str() {
            "toml" | "" => toml::from_str(&content)
                .map_err(|error| DocxError::parse(format!("invalid TOML limits: {error}")))?,
            "json" => serde_json::from_str(&content)
                .map_err(|error| DocxError::parse(format!("invalid JSON limits: {error}")))?,
            other => {
                return Err(DocxError::parse(format!(
                    "unsupported limits extension '{other}', expected .toml or .json"
                )))
            }
        };
        limits.validate()?;
        Ok(limits)
    }

    /// Rejects zero ceilings and internally inconsistent ZIP limits.
    pub fn validate(&self) -> Result<()> {
        let positive = [
            ("max_docx_archive_bytes", self.max_docx_archive_bytes),
            ("max_docx_entry_bytes", self.max_docx_entry_bytes),
            ("max_docx_total_bytes", self.max_docx_total_bytes),
            ("max_zip_compression_ratio", self.max_zip_compression_ratio),
            ("max_xml_bytes", self.max_xml_bytes),
            ("max_spec_bytes", self.max_spec_bytes),
            ("max_image_bytes", self.max_image_bytes),
            ("max_svg_bytes", self.max_svg_bytes),
            ("max_image_pixels", self.max_image_pixels),
            (
                "max_template_output_xml_bytes",
                self.max_template_output_xml_bytes,
            ),
        ];
        if let Some((name, _)) = positive.into_iter().find(|(_, value)| *value == 0) {
            return Err(DocxError::resource_limit(format!(
                "{name} must be greater than zero"
            )));
        }
        let counts = [
            ("max_docx_entries", self.max_docx_entries),
            ("max_include_depth", self.max_include_depth),
            ("max_include_files", self.max_include_files),
            ("max_template_expansions", self.max_template_expansions),
            (
                "max_template_partial_depth",
                self.max_template_partial_depth,
            ),
        ];
        if let Some((name, _)) = counts.into_iter().find(|(_, value)| *value == 0) {
            return Err(DocxError::resource_limit(format!(
                "{name} must be greater than zero"
            )));
        }
        if self.max_docx_entry_bytes > self.max_docx_total_bytes {
            return Err(DocxError::resource_limit(
                "max_docx_entry_bytes cannot exceed max_docx_total_bytes".to_string(),
            ));
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::InputLimits;

    #[test]
    fn hosted_profile_is_stricter_and_valid() {
        let hosted = InputLimits::hosted();
        let default = InputLimits::default();
        hosted.validate().expect("hosted profile");
        assert!(hosted.max_docx_total_bytes < default.max_docx_total_bytes);
        assert!(hosted.max_image_pixels < default.max_image_pixels);
    }

    #[test]
    fn invalid_profiles_fail_closed() {
        let invalid = InputLimits {
            max_spec_bytes: 0,
            ..InputLimits::default()
        };
        assert!(invalid.validate().is_err());
    }

    #[test]
    fn serialized_profiles_must_be_complete_and_known() {
        assert!(toml::from_str::<InputLimits>("max_spec_bytes = 64").is_err());
        let mut profile = toml::to_string(&InputLimits::hosted()).expect("serialize profile");
        profile.push_str("unknown_budget = 1\n");
        assert!(toml::from_str::<InputLimits>(&profile).is_err());
    }

    #[test]
    fn checked_in_hosted_profile_matches_the_builtin() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("examples/config/hosted-limits.toml");
        assert_eq!(
            InputLimits::load_from_path(path).expect("hosted limits fixture"),
            InputLimits::hosted()
        );
    }
}
