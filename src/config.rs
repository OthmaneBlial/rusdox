//! Configuration model for RusDox document and preview generation.
//!
//! The recommended format is TOML because it supports comments and is easy to
//! hand-edit. JSON is also supported for machine-generated workflows.

use std::fs;
use std::path::{Path, PathBuf};

use serde::{Deserialize, Serialize};

use crate::{DocxError, Result};

/// Top-level RusDox configuration.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct RusdoxConfig {
    /// Friendly profile name shown by tooling.
    pub profile_name: String,
    /// Output folder behavior for generated artifacts.
    pub output: OutputConfig,
    /// Font and size settings for common document blocks.
    pub typography: TypographyConfig,
    /// Paragraph and cell spacing settings in twips (DOCX units).
    pub spacing: SpacingConfig,
    /// Shared color palette used by helper APIs.
    pub colors: ColorConfig,
    /// Table styling and sizing defaults.
    pub table: TableConfig,
    /// PDF preview rendering settings.
    pub pdf: PdfConfig,
}

impl Default for RusdoxConfig {
    fn default() -> Self {
        Self {
            profile_name: "RusDox Default".to_string(),
            output: OutputConfig::default(),
            typography: TypographyConfig::default(),
            spacing: SpacingConfig::default(),
            colors: ColorConfig::default(),
            table: TableConfig::default(),
            pdf: PdfConfig::default(),
        }
    }
}

impl RusdoxConfig {
    /// Loads a configuration from a file path.
    ///
    /// `.toml` and `.json` are supported.
    pub fn load_from_path(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let content = fs::read_to_string(path)?;
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();

        match extension.as_str() {
            "json" => Self::from_json_str(&content),
            "toml" | "" => Self::from_toml_str(&content),
            other => Err(DocxError::parse(format!(
                "unsupported config extension '{other}', expected .toml or .json"
            ))),
        }
    }

    /// Loads configuration if the file exists, otherwise returns defaults.
    pub fn load_from_path_or_default(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        if path.exists() {
            Self::load_from_path(path)
        } else {
            Ok(Self::default())
        }
    }

    /// Loads a local config when present, otherwise falls back to the user-level config path.
    ///
    /// If neither exists, defaults are returned.
    pub fn load_local_or_user_default(local_path: impl AsRef<Path>) -> Result<Self> {
        Self::load_local_or_user_default_with_user_path(
            local_path.as_ref(),
            default_user_config_path().as_deref(),
        )
    }

    /// Parses a TOML configuration string.
    pub fn from_toml_str(content: &str) -> Result<Self> {
        toml::from_str(content)
            .map_err(|error| DocxError::parse(format!("invalid TOML config: {error}")))
    }

    /// Parses a JSON configuration string.
    pub fn from_json_str(content: &str) -> Result<Self> {
        serde_json::from_str(content)
            .map_err(|error| DocxError::parse(format!("invalid JSON config: {error}")))
    }

    /// Serializes the current configuration to JSON.
    pub fn to_json_pretty(&self) -> Result<String> {
        serde_json::to_string_pretty(self)
            .map_err(|error| DocxError::parse(format!("failed to serialize JSON config: {error}")))
    }

    /// Serializes the current configuration to TOML.
    pub fn to_toml_pretty(&self) -> Result<String> {
        toml::to_string_pretty(self)
            .map_err(|error| DocxError::parse(format!("failed to serialize TOML config: {error}")))
    }

    /// Saves the current configuration to disk.
    ///
    /// `.toml` and `.json` are supported. If no extension is provided,
    /// TOML is used by default.
    pub fn save_to_path(&self, path: impl AsRef<Path>) -> Result<()> {
        let path = path.as_ref();
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)?;
        }

        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or("toml")
            .to_ascii_lowercase();

        let content = match extension.as_str() {
            "json" => self.to_json_pretty()?,
            "toml" | "" => self.to_toml_pretty()?,
            other => {
                return Err(DocxError::parse(format!(
                    "unsupported config extension '{other}', expected .toml or .json"
                )))
            }
        };

        fs::write(path, content)?;
        Ok(())
    }

    /// Writes a commented default TOML template to disk.
    pub fn write_toml_template(path: impl AsRef<Path>) -> Result<()> {
        let path = path.as_ref();
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)?;
        }
        fs::write(path, Self::default_toml_template())?;
        Ok(())
    }

    /// Returns the commented default TOML template.
    pub fn default_toml_template() -> &'static str {
        DEFAULT_TOML_TEMPLATE
    }

    fn load_local_or_user_default_with_user_path(
        local_path: &Path,
        user_path: Option<&Path>,
    ) -> Result<Self> {
        if local_path.exists() {
            return Self::load_from_path(local_path);
        }

        if let Some(user_path) = user_path {
            if user_path.exists() {
                return Self::load_from_path(user_path);
            }
        }

        Ok(Self::default())
    }
}

/// Returns the default user-level config path.
///
/// Linux/macOS: `~/rusdox/config.toml`
/// Windows: `%USERPROFILE%\rusdox\config.toml`
pub fn default_user_config_path() -> Option<PathBuf> {
    #[cfg(windows)]
    {
        std::env::var_os("USERPROFILE")
            .map(PathBuf::from)
            .map(|home| home.join("rusdox").join("config.toml"))
    }

    #[cfg(not(windows))]
    {
        std::env::var_os("HOME")
            .map(PathBuf::from)
            .map(|home| home.join("rusdox").join("config.toml"))
    }
}

/// Output path and preview behavior.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct OutputConfig {
    /// Target directory for generated `.docx` files.
    pub docx_dir: String,
    /// Target directory for generated `.pdf` previews.
    pub pdf_dir: String,
    /// Whether to emit PDF previews alongside DOCX output.
    pub emit_pdf_preview: bool,
}

impl Default for OutputConfig {
    fn default() -> Self {
        Self {
            docx_dir: "generated".to_string(),
            pdf_dir: "rendered".to_string(),
            emit_pdf_preview: true,
        }
    }
}

/// Typography defaults used by the style helpers.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct TypographyConfig {
    /// Default run font family used by helper methods.
    pub font_family: String,
    pub cover_title_size_pt: f32,
    pub title_size_pt: f32,
    pub subtitle_size_pt: f32,
    pub hero_size_pt: f32,
    pub page_heading_size_pt: f32,
    pub section_size_pt: f32,
    pub body_size_pt: f32,
    pub tagline_size_pt: f32,
    pub note_size_pt: f32,
    pub table_size_pt: f32,
    pub metric_label_size_pt: f32,
    pub metric_value_size_pt: f32,
}

impl Default for TypographyConfig {
    fn default() -> Self {
        Self {
            font_family: "Arial".to_string(),
            cover_title_size_pt: 30.0,
            title_size_pt: 26.0,
            subtitle_size_pt: 11.0,
            hero_size_pt: 14.0,
            page_heading_size_pt: 20.0,
            section_size_pt: 15.0,
            body_size_pt: 11.0,
            tagline_size_pt: 11.0,
            note_size_pt: 10.0,
            table_size_pt: 10.0,
            metric_label_size_pt: 10.0,
            metric_value_size_pt: 18.0,
        }
    }
}

/// Spacing defaults expressed in twips (1/20th point).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct SpacingConfig {
    pub cover_title_before_twips: u32,
    pub cover_title_after_twips: u32,
    pub title_before_twips: u32,
    pub title_after_twips: u32,
    pub subtitle_after_twips: u32,
    pub hero_after_twips: u32,
    pub page_heading_after_twips: u32,
    pub section_before_twips: u32,
    pub section_after_twips: u32,
    pub body_after_twips: u32,
    pub bullet_after_twips: u32,
    pub label_value_after_twips: u32,
    pub tagline_after_twips: u32,
    pub spacer_after_twips: u32,
    pub note_after_twips: u32,
    pub metric_label_before_twips: u32,
    pub metric_label_after_twips: u32,
    pub metric_value_after_twips: u32,
    pub table_header_before_twips: u32,
    pub table_header_after_twips: u32,
    pub table_data_before_twips: u32,
    pub table_data_after_twips: u32,
    pub table_status_before_twips: u32,
    pub table_status_after_twips: u32,
}

impl Default for SpacingConfig {
    fn default() -> Self {
        Self {
            cover_title_before_twips: 1_200,
            cover_title_after_twips: 180,
            title_before_twips: 240,
            title_after_twips: 120,
            subtitle_after_twips: 280,
            hero_after_twips: 240,
            page_heading_after_twips: 140,
            section_before_twips: 220,
            section_after_twips: 90,
            body_after_twips: 100,
            bullet_after_twips: 80,
            label_value_after_twips: 70,
            tagline_after_twips: 80,
            spacer_after_twips: 120,
            note_after_twips: 120,
            metric_label_before_twips: 60,
            metric_label_after_twips: 20,
            metric_value_after_twips: 80,
            table_header_before_twips: 50,
            table_header_after_twips: 50,
            table_data_before_twips: 40,
            table_data_after_twips: 40,
            table_status_before_twips: 40,
            table_status_after_twips: 40,
        }
    }
}

/// Palette used by helper functions and presets.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct ColorConfig {
    pub ink: String,
    pub slate: String,
    pub muted: String,
    pub accent: String,
    pub gold: String,
    pub red: String,
    pub green: String,
    pub soft: String,
    pub pale: String,
    pub mint: String,
    pub amber: String,
    pub rose: String,
    pub table_border: String,
}

impl Default for ColorConfig {
    fn default() -> Self {
        Self {
            ink: "0F172A".to_string(),
            slate: "475569".to_string(),
            muted: "64748B".to_string(),
            accent: "0F766E".to_string(),
            gold: "B45309".to_string(),
            red: "B91C1C".to_string(),
            green: "166534".to_string(),
            soft: "E2E8F0".to_string(),
            pale: "F8FAFC".to_string(),
            mint: "DCFCE7".to_string(),
            amber: "FEF3C7".to_string(),
            rose: "FEE2E2".to_string(),
            table_border: "CBD5E1".to_string(),
        }
    }
}

/// Table defaults used by helper APIs and PDF preview rendering.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct TableConfig {
    pub default_width_twips: u32,
    pub metric_cell_width_twips: u32,
    pub grid_border_size_eighth_pt: u32,
    pub card_border_size_eighth_pt: u32,
    pub pdf_cell_padding_x_pt: f32,
    pub pdf_cell_padding_y_pt: f32,
    pub pdf_after_spacing_pt: f32,
    pub pdf_grid_stroke_width_pt: f32,
}

impl Default for TableConfig {
    fn default() -> Self {
        Self {
            default_width_twips: 9_360,
            metric_cell_width_twips: 3_120,
            grid_border_size_eighth_pt: 8,
            card_border_size_eighth_pt: 10,
            pdf_cell_padding_x_pt: 7.0,
            pdf_cell_padding_y_pt: 6.0,
            pdf_after_spacing_pt: 12.0,
            pdf_grid_stroke_width_pt: 0.75,
        }
    }
}

/// PDF preview renderer settings.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct PdfConfig {
    pub page_width_pt: f32,
    pub page_height_pt: f32,
    pub margin_x_pt: f32,
    pub margin_top_pt: f32,
    pub margin_bottom_pt: f32,
    pub default_text_size_pt: f32,
    pub default_line_height_pt: f32,
    pub line_height_multiplier: f32,
    pub baseline_factor: f32,
    pub text_width_bias_regular: f32,
    pub text_width_bias_bold: f32,
}

impl Default for PdfConfig {
    fn default() -> Self {
        Self {
            page_width_pt: 612.0,
            page_height_pt: 792.0,
            margin_x_pt: 54.0,
            margin_top_pt: 54.0,
            margin_bottom_pt: 54.0,
            default_text_size_pt: 11.0,
            default_line_height_pt: 14.0,
            line_height_multiplier: 1.35,
            baseline_factor: 0.82,
            text_width_bias_regular: 1.0,
            text_width_bias_bold: 1.03,
        }
    }
}

const DEFAULT_TOML_TEMPLATE: &str = r#"# RusDox configuration template
# Most users should use the CLI wizard instead of editing this by hand first:
#   rusdox config wizard --level basic
#   rusdox config wizard --level advanced
# Create or edit a project-local override in the current folder:
#   rusdox config wizard --path ./rusdox.toml --level basic
# Print the user-level config path:
#   rusdox config path
# Config load order while rendering documents:
#   1. ./rusdox.toml
#   2. ~/rusdox/config.toml
#   3. built-in defaults
# Every field below is optional: missing values fall back to defaults.

profile_name = "RusDox Default"

[output]
# Directory for generated DOCX files
docx_dir = "generated"
# Directory for generated PDF preview files
pdf_dir = "rendered"
# Set false if you only want DOCX output
emit_pdf_preview = true

[typography]
# Font family used by helper builders
font_family = "Arial"
cover_title_size_pt = 30.0
title_size_pt = 26.0
subtitle_size_pt = 11.0
hero_size_pt = 14.0
page_heading_size_pt = 20.0
section_size_pt = 15.0
body_size_pt = 11.0
tagline_size_pt = 11.0
note_size_pt = 10.0
table_size_pt = 10.0
metric_label_size_pt = 10.0
metric_value_size_pt = 18.0

[spacing]
# DOCX spacing values in twips (1/20 point)
cover_title_before_twips = 1200
cover_title_after_twips = 180
title_before_twips = 240
title_after_twips = 120
subtitle_after_twips = 280
hero_after_twips = 240
page_heading_after_twips = 140
section_before_twips = 220
section_after_twips = 90
body_after_twips = 100
bullet_after_twips = 80
label_value_after_twips = 70
tagline_after_twips = 80
spacer_after_twips = 120
note_after_twips = 120
metric_label_before_twips = 60
metric_label_after_twips = 20
metric_value_after_twips = 80
table_header_before_twips = 50
table_header_after_twips = 50
table_data_before_twips = 40
table_data_after_twips = 40
table_status_before_twips = 40
table_status_after_twips = 40

[colors]
# Six-digit hex values without '#'
ink = "0F172A"
slate = "475569"
muted = "64748B"
accent = "0F766E"
gold = "B45309"
red = "B91C1C"
green = "166534"
soft = "E2E8F0"
pale = "F8FAFC"
mint = "DCFCE7"
amber = "FEF3C7"
rose = "FEE2E2"
table_border = "CBD5E1"

[table]
default_width_twips = 9360
metric_cell_width_twips = 3120
grid_border_size_eighth_pt = 8
card_border_size_eighth_pt = 10
pdf_cell_padding_x_pt = 7.0
pdf_cell_padding_y_pt = 6.0
pdf_after_spacing_pt = 12.0
pdf_grid_stroke_width_pt = 0.75

[pdf]
# Letter page size by default
page_width_pt = 612.0
page_height_pt = 792.0
margin_x_pt = 54.0
margin_top_pt = 54.0
margin_bottom_pt = 54.0
default_text_size_pt = 11.0
default_line_height_pt = 14.0
line_height_multiplier = 1.35
baseline_factor = 0.82
text_width_bias_regular = 1.0
text_width_bias_bold = 1.03
"#;

#[cfg(test)]
mod tests {
    use std::fs;

    use tempfile::tempdir;

    use super::{
        default_user_config_path, ColorConfig, OutputConfig, RusdoxConfig, SpacingConfig,
        TypographyConfig,
    };
    use crate::DocxError;

    #[test]
    fn default_values_are_stable_and_complete() {
        let config = RusdoxConfig::default();
        assert_eq!(config.profile_name, "RusDox Default");
        assert_eq!(config.output.docx_dir, "generated");
        assert_eq!(config.output.pdf_dir, "rendered");
        assert!(config.output.emit_pdf_preview);
        assert_eq!(config.typography.font_family, "Arial");
        assert_eq!(config.typography.cover_title_size_pt, 30.0);
        assert_eq!(config.colors.accent, "0F766E");
        assert_eq!(config.spacing.page_heading_after_twips, 140);
        assert_eq!(config.table.default_width_twips, 9_360);
        assert_eq!(config.pdf.page_width_pt, 612.0);
    }

    #[test]
    fn from_toml_supports_partial_config_with_defaults() {
        let config = RusdoxConfig::from_toml_str(
            r#"
profile_name = "Custom"

[output]
docx_dir = "out/docx"
"#,
        )
        .expect("toml should parse");

        assert_eq!(config.profile_name, "Custom");
        assert_eq!(config.output.docx_dir, "out/docx");
        assert_eq!(config.output.pdf_dir, "rendered");
        assert!(config.output.emit_pdf_preview);
    }

    #[test]
    fn json_round_trip_preserves_key_fields() {
        let config = RusdoxConfig {
            profile_name: "JSON Profile".to_string(),
            output: OutputConfig {
                emit_pdf_preview: false,
                ..OutputConfig::default()
            },
            colors: ColorConfig {
                ink: "ABCDEF".to_string(),
                ..ColorConfig::default()
            },
            ..RusdoxConfig::default()
        };

        let json = config.to_json_pretty().expect("json serialize");
        let parsed = RusdoxConfig::from_json_str(&json).expect("json parse");
        assert_eq!(parsed, config);
    }

    #[test]
    fn toml_round_trip_preserves_key_fields() {
        let config = RusdoxConfig {
            profile_name: "TOML Profile".to_string(),
            typography: TypographyConfig {
                body_size_pt: 12.5,
                ..TypographyConfig::default()
            },
            spacing: SpacingConfig {
                section_after_twips: 333,
                ..SpacingConfig::default()
            },
            ..RusdoxConfig::default()
        };

        let toml = config.to_toml_pretty().expect("toml serialize");
        let parsed = RusdoxConfig::from_toml_str(&toml).expect("toml parse");
        assert_eq!(parsed, config);
    }

    #[test]
    fn load_from_path_uses_extension_based_parser() {
        let temp = tempdir().expect("temp dir");
        let toml_path = temp.path().join("config.toml");
        let json_path = temp.path().join("config.json");

        fs::write(
            &toml_path,
            r#"
profile_name = "Toml Path"
[output]
emit_pdf_preview = false
"#,
        )
        .expect("write toml");
        fs::write(
            &json_path,
            r#"{"profile_name":"Json Path","output":{"emit_pdf_preview":false}}"#,
        )
        .expect("write json");

        let toml_config = RusdoxConfig::load_from_path(&toml_path).expect("load toml");
        let json_config = RusdoxConfig::load_from_path(&json_path).expect("load json");

        assert_eq!(toml_config.profile_name, "Toml Path");
        assert!(!toml_config.output.emit_pdf_preview);
        assert_eq!(json_config.profile_name, "Json Path");
        assert!(!json_config.output.emit_pdf_preview);
    }

    #[test]
    fn load_from_path_without_extension_defaults_to_toml() {
        let temp = tempdir().expect("temp dir");
        let path = temp.path().join("rusdox_config");
        fs::write(
            &path,
            r#"
profile_name = "NoExt"
[output]
docx_dir = "x"
"#,
        )
        .expect("write file");

        let config = RusdoxConfig::load_from_path(&path).expect("load no-extension file");
        assert_eq!(config.profile_name, "NoExt");
        assert_eq!(config.output.docx_dir, "x");
    }

    #[test]
    fn unsupported_extension_errors_on_load_and_save() {
        let temp = tempdir().expect("temp dir");
        let load_path = temp.path().join("config.yaml");
        fs::write(&load_path, "profile_name: x").expect("write yaml");

        let load_error = RusdoxConfig::load_from_path(&load_path).expect_err("must fail");
        assert!(
            matches!(load_error, DocxError::Parse(message) if message.contains("unsupported config extension"))
        );

        let save_path = temp.path().join("config.ini");
        let save_error = RusdoxConfig::default()
            .save_to_path(&save_path)
            .expect_err("must fail");
        assert!(
            matches!(save_error, DocxError::Parse(message) if message.contains("unsupported config extension"))
        );
    }

    #[test]
    fn load_from_path_or_default_returns_default_when_missing() {
        let temp = tempdir().expect("temp dir");
        let missing = temp.path().join("missing.toml");
        let config = RusdoxConfig::load_from_path_or_default(missing).expect("default config");
        assert_eq!(config, RusdoxConfig::default());
    }

    #[test]
    fn load_local_or_user_default_prefers_local_then_user_then_default() {
        let temp = tempdir().expect("temp dir");
        let local_path = temp.path().join("local.toml");
        let user_path = temp.path().join("user.toml");

        fs::write(
            &user_path,
            r#"
profile_name = "User"
"#,
        )
        .expect("write user config");

        let from_user =
            RusdoxConfig::load_local_or_user_default_with_user_path(&local_path, Some(&user_path))
                .expect("load user fallback");
        assert_eq!(from_user.profile_name, "User");

        fs::write(
            &local_path,
            r#"
profile_name = "Local"
"#,
        )
        .expect("write local config");

        let from_local =
            RusdoxConfig::load_local_or_user_default_with_user_path(&local_path, Some(&user_path))
                .expect("load local config");
        assert_eq!(from_local.profile_name, "Local");

        fs::remove_file(&local_path).expect("remove local config");
        fs::remove_file(&user_path).expect("remove user config");

        let defaulted =
            RusdoxConfig::load_local_or_user_default_with_user_path(&local_path, Some(&user_path))
                .expect("load default config");
        assert_eq!(defaulted, RusdoxConfig::default());
    }

    #[test]
    fn save_to_path_creates_parent_directories_and_writes_toml_json() {
        let temp = tempdir().expect("temp dir");
        let nested_toml = temp.path().join("a/b/c/config.toml");
        let nested_json = temp.path().join("a/b/c/config.json");

        let config = RusdoxConfig {
            profile_name: "Persisted".to_string(),
            ..RusdoxConfig::default()
        };
        config.save_to_path(&nested_toml).expect("save toml");
        config.save_to_path(&nested_json).expect("save json");

        let parsed_toml = RusdoxConfig::load_from_path(&nested_toml).expect("load toml");
        let parsed_json = RusdoxConfig::load_from_path(&nested_json).expect("load json");

        assert_eq!(parsed_toml.profile_name, "Persisted");
        assert_eq!(parsed_json.profile_name, "Persisted");
    }

    #[test]
    fn write_toml_template_writes_expected_content() {
        let temp = tempdir().expect("temp dir");
        let path = temp.path().join("template/rusdox.toml");
        RusdoxConfig::write_toml_template(&path).expect("write template");

        let content = fs::read_to_string(&path).expect("read template");
        assert!(content.contains("# RusDox configuration template"));
        assert!(content.contains("rusdox config wizard --level basic"));
        assert!(content.contains("rusdox config wizard --path ./rusdox.toml --level basic"));
        assert!(content.contains("Config load order while rendering documents"));
        assert!(content.contains("[output]"));
        assert!(content.contains("[pdf]"));
        assert_eq!(content, RusdoxConfig::default_toml_template());
    }

    #[test]
    fn default_user_config_path_shape_is_consistent() {
        if let Some(path) = default_user_config_path() {
            assert!(path.ends_with(std::path::Path::new("rusdox").join("config.toml")));
        }
    }

    #[test]
    fn invalid_toml_and_json_return_parse_errors() {
        let toml_error = RusdoxConfig::from_toml_str("not = [valid").expect_err("invalid toml");
        let json_error = RusdoxConfig::from_json_str("{invalid json").expect_err("invalid json");

        assert!(
            matches!(toml_error, DocxError::Parse(message) if message.contains("invalid TOML"))
        );
        assert!(
            matches!(json_error, DocxError::Parse(message) if message.contains("invalid JSON"))
        );
    }
}
