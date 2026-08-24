//! Stable, transport-neutral rendering boundary for Rust embedding and local protocols.

use std::io::Cursor;
use std::path::PathBuf;
use std::time::{Duration, Instant};

use serde::{Deserialize, Serialize};

use crate::config::RusdoxConfig;
use crate::spec::DocumentSpec;
use crate::studio::Studio;
use crate::validate::{validate_config, validate_spec_with_config, ValidationIssue};
use crate::{attach_source_spans, attach_source_spans_from_str, DocxError, InputLimits, Result};

/// Version of the transport-neutral Rust rendering request.
pub const RENDERER_API_VERSION: u32 = 1;

/// Supported document-spec serialization formats for inline requests.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SpecFormat {
    /// YAML document spec.
    Yaml,
    /// JSON document spec.
    Json,
    /// TOML document spec.
    Toml,
}

impl SpecFormat {
    fn as_str(self) -> &'static str {
        match self {
            Self::Yaml => "yaml",
            Self::Json => "json",
            Self::Toml => "toml",
        }
    }
}

/// File-backed or in-memory source accepted by every renderer implementation.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case", deny_unknown_fields)]
pub enum RenderSource {
    /// Load a local spec path, preserving include and asset resolution.
    Path {
        /// YAML, JSON, or TOML spec path.
        path: PathBuf,
    },
    /// Parse a self-contained spec without filesystem access.
    Inline {
        /// Inline serialization format.
        format: SpecFormat,
        /// Complete document-spec source.
        content: String,
    },
}

/// Versioned request for a transport-neutral renderer.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct RenderRequest {
    /// Must equal [`RENDERER_API_VERSION`].
    pub renderer_api_version: u32,
    /// Spec source, either local path or inline text.
    pub source: RenderSource,
    /// Whether the native PDF artifact is required alongside DOCX.
    #[serde(default = "default_true")]
    pub emit_pdf: bool,
}

/// Validation result shared by Rust callers and protocol transports.
#[derive(Debug, Clone, Serialize)]
pub struct RendererValidation {
    /// Whether no error-severity diagnostic was produced.
    pub valid: bool,
    /// Semantic diagnostics, including source spans when available.
    pub diagnostics: Vec<ValidationIssue>,
    /// Spec parsing time.
    pub parse_duration: Duration,
    /// Semantic/config validation time.
    pub validation_duration: Duration,
}

/// In-memory artifacts returned by a renderer implementation.
#[derive(Debug, Clone)]
pub struct RenderedDocument {
    /// Editable OOXML package bytes.
    pub docx: Vec<u8>,
    /// Native PDF bytes when requested.
    pub pdf: Option<Vec<u8>>,
    /// Diagnostics from the validation gate.
    pub diagnostics: Vec<ValidationIssue>,
    /// Spec parsing time.
    pub parse_duration: Duration,
    /// Semantic/config validation time.
    pub validation_duration: Duration,
    /// Typed document composition time.
    pub compose_duration: Duration,
    /// DOCX package writing time.
    pub docx_duration: Duration,
    /// Native PDF rendering time.
    pub pdf_duration: Duration,
}

/// Object-safe boundary implemented by native, future WASM, or alternate backends.
pub trait Renderer: Send + Sync {
    /// Parse and validate without producing artifacts.
    fn validate(&self, request: &RenderRequest) -> Result<RendererValidation>;
    /// Validate, compose, and return requested artifacts in memory.
    fn render(&self, request: &RenderRequest) -> Result<RenderedDocument>;
}

/// Pure-Rust renderer used by the CLI and local JSON transports.
#[derive(Debug, Clone)]
pub struct NativeRenderer {
    config: RusdoxConfig,
    limits: InputLimits,
}

impl NativeRenderer {
    /// Create a renderer with explicit configuration and default resource limits.
    pub fn new(config: RusdoxConfig) -> Self {
        Self {
            config,
            limits: InputLimits::default(),
        }
    }

    /// Override resource ceilings for untrusted or hosted inputs.
    pub fn with_limits(mut self, limits: InputLimits) -> Self {
        self.limits = limits;
        self
    }

    fn inspect(&self, request: &RenderRequest) -> Result<(DocumentSpec, RendererValidation)> {
        ensure_renderer_version(request.renderer_api_version)?;
        let parse_started = Instant::now();
        let spec = match &request.source {
            RenderSource::Path { path } => {
                DocumentSpec::load_from_path_with_limits(path, self.limits)?
            }
            RenderSource::Inline { format, content } => match format {
                SpecFormat::Yaml => DocumentSpec::from_yaml_str_with_limits(content, self.limits)?,
                SpecFormat::Json => DocumentSpec::from_json_str_with_limits(content, self.limits)?,
                SpecFormat::Toml => DocumentSpec::from_toml_str_with_limits(content, self.limits)?,
            },
        };
        let parse_duration = parse_started.elapsed();
        let validation_started = Instant::now();
        let mut report = validate_spec_with_config(&spec, &self.config);
        match &request.source {
            RenderSource::Path { path } => attach_source_spans(path, &mut report)?,
            RenderSource::Inline { format, content } => {
                attach_source_spans_from_str(content, format.as_str(), &mut report);
            }
        }
        let mut config_report = validate_config(&self.config);
        config_report.issues.extend(report.issues);
        let report = config_report;
        let validation_duration = validation_started.elapsed();
        Ok((
            spec,
            RendererValidation {
                valid: !report.has_errors(),
                diagnostics: report.issues,
                parse_duration,
                validation_duration,
            },
        ))
    }
}

impl Renderer for NativeRenderer {
    fn validate(&self, request: &RenderRequest) -> Result<RendererValidation> {
        self.inspect(request).map(|(_, validation)| validation)
    }

    fn render(&self, request: &RenderRequest) -> Result<RenderedDocument> {
        let (spec, validation) = self.inspect(request)?;
        if !validation.valid {
            return Err(DocxError::Parse(format!(
                "rendering rejected by {} validation error(s)",
                validation
                    .diagnostics
                    .iter()
                    .filter(|issue| issue.severity == crate::ValidationSeverity::Error)
                    .count()
            )));
        }
        let studio = Studio::new(self.config.clone());
        let compose_started = Instant::now();
        let document = studio.compose(&spec);
        let compose_duration = compose_started.elapsed();

        let docx_started = Instant::now();
        let mut cursor = Cursor::new(Vec::new());
        document.save_to_writer(&mut cursor)?;
        let docx = cursor.into_inner();
        let docx_duration = docx_started.elapsed();

        let (pdf, pdf_duration) = if request.emit_pdf {
            let pdf_started = Instant::now();
            let workspace = tempfile::tempdir()?;
            let pdf_path = workspace.path().join("rendered.pdf");
            studio.render_pdf_with_evidence(&document, &pdf_path, None)?;
            (Some(std::fs::read(pdf_path)?), pdf_started.elapsed())
        } else {
            (None, Duration::ZERO)
        };

        Ok(RenderedDocument {
            docx,
            pdf,
            diagnostics: validation.diagnostics,
            parse_duration: validation.parse_duration,
            validation_duration: validation.validation_duration,
            compose_duration,
            docx_duration,
            pdf_duration,
        })
    }
}

fn ensure_renderer_version(version: u32) -> Result<()> {
    if version == RENDERER_API_VERSION {
        Ok(())
    } else {
        Err(DocxError::Parse(format!(
            "unsupported renderer_api_version {version}; expected {RENDERER_API_VERSION}"
        )))
    }
}

const fn default_true() -> bool {
    true
}

#[cfg(test)]
mod tests {
    use super::{
        NativeRenderer, RenderRequest, RenderSource, Renderer, SpecFormat, RENDERER_API_VERSION,
    };
    use crate::config::RusdoxConfig;

    fn inline_request(content: &str) -> RenderRequest {
        RenderRequest {
            renderer_api_version: RENDERER_API_VERSION,
            source: RenderSource::Inline {
                format: SpecFormat::Yaml,
                content: content.to_string(),
            },
            emit_pdf: true,
        }
    }

    #[test]
    fn native_renderer_returns_docx_and_pdf_in_memory() {
        let renderer = NativeRenderer::new(RusdoxConfig::default());
        let artifacts = renderer
            .render(&inline_request(
                "version: 1\noutput_name: protocol\nblocks:\n  - type: body\n    text: Hello protocol\n",
            ))
            .expect("render");
        assert!(artifacts.docx.starts_with(b"PK"));
        assert!(artifacts
            .pdf
            .as_deref()
            .is_some_and(|pdf| pdf.starts_with(b"%PDF")));
    }

    #[test]
    fn native_renderer_exposes_inline_source_spans_before_render() {
        let renderer = NativeRenderer::new(RusdoxConfig::default());
        let validation = renderer
            .validate(&inline_request(
                "version: 1\noutput_name: protocol\nblocks:\n  - type: body\n    text: ''\n",
            ))
            .expect("validate");
        assert!(validation.valid);
        assert!(validation
            .diagnostics
            .iter()
            .any(|issue| issue.source.is_some_and(|source| source.line == 5)));
    }

    #[test]
    fn renderer_rejects_unknown_api_versions() {
        let renderer = NativeRenderer::new(RusdoxConfig::default());
        let mut request = inline_request("version: 1\nblocks: []\n");
        request.renderer_api_version = 99;
        assert!(renderer.validate(&request).is_err());
    }
}
