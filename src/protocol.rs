//! Versioned local JSON protocol shared by stdin/stdout and loopback HTTP transports.

use std::path::{Component, Path, PathBuf};

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::config::RusdoxConfig;
use crate::renderer::{
    NativeRenderer, RenderRequest, RenderSource, Renderer, RENDERER_API_VERSION,
};
use crate::{atomic_write_file, ValidationIssue};

/// Stable protocol version accepted by local integrations.
pub const PROTOCOL_VERSION: u32 = 1;

/// Maximum UTF-8 JSON request size accepted by bundled transports.
pub const MAX_PROTOCOL_REQUEST_BYTES: usize = 2 * 1024 * 1024;

/// Supported protocol operations.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ProtocolOperation {
    /// Parse and validate without producing files.
    Validate,
    /// Produce atomic DOCX and optional PDF artifacts.
    Render,
}

/// Output selection for a render operation.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct ProtocolOutput {
    /// Relative directory beneath the service output root.
    #[serde(default = "default_output_directory")]
    pub directory: PathBuf,
    /// Safe artifact stem. Derived from the source when omitted.
    pub name: Option<String>,
    /// Generate native PDF alongside DOCX.
    #[serde(default = "default_true")]
    pub pdf: bool,
}

/// One request in the stable local protocol.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct ProtocolRequest {
    /// Must equal [`PROTOCOL_VERSION`].
    pub protocol_version: u32,
    /// Caller-selected correlation identifier, limited to 128 printable characters.
    pub request_id: String,
    /// Validation or render operation.
    pub operation: ProtocolOperation,
    /// File-backed or inline spec source.
    pub source: RenderSource,
    /// Optional complete config object. Defaults are used when absent.
    #[serde(default)]
    pub config: Option<RusdoxConfig>,
    /// Required only for render operations.
    #[serde(default)]
    pub output: Option<ProtocolOutput>,
}

/// One written artifact and its integrity metadata.
#[derive(Debug, Clone, Serialize)]
pub struct ProtocolArtifact {
    /// `docx` or `pdf`.
    pub kind: String,
    /// Absolute local output path.
    pub path: String,
    /// SHA-256 of the written bytes.
    pub sha256: String,
    /// Written byte count.
    pub bytes: usize,
}

/// Per-stage timings in milliseconds.
#[derive(Debug, Clone, Serialize)]
pub struct ProtocolTimings {
    pub parse_ms: f64,
    pub validate_ms: f64,
    pub compose_ms: f64,
    pub docx_ms: f64,
    pub pdf_ms: f64,
}

/// Machine-readable protocol failure.
#[derive(Debug, Clone, Serialize)]
pub struct ProtocolError {
    /// Stable error category.
    pub code: String,
    /// Human-readable local diagnostic.
    pub message: String,
}

/// One response emitted for every request, including malformed or rejected work.
#[derive(Debug, Clone, Serialize)]
pub struct ProtocolResponse {
    pub protocol_version: u32,
    pub request_id: String,
    pub ok: bool,
    pub diagnostics: Vec<ValidationIssue>,
    pub artifacts: Vec<ProtocolArtifact>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub timings: Option<ProtocolTimings>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<ProtocolError>,
}

/// Execute one validated request beneath a fixed local output root.
pub fn execute_protocol_request(
    request: ProtocolRequest,
    output_root: impl AsRef<Path>,
) -> ProtocolResponse {
    let request_id = request.request_id.clone();
    if let Err(message) = validate_request_envelope(&request) {
        return protocol_failure(request_id, "invalid_request", message);
    }
    let renderer = NativeRenderer::new(request.config.clone().unwrap_or_default());
    let render_request = RenderRequest {
        renderer_api_version: RENDERER_API_VERSION,
        source: request.source.clone(),
        emit_pdf: request.output.as_ref().is_none_or(|output| output.pdf),
    };
    let validation = match renderer.validate(&render_request) {
        Ok(validation) => validation,
        Err(error) => return protocol_failure(request_id, "parse_error", error.to_string()),
    };
    let validation_timings = ProtocolTimings {
        parse_ms: duration_ms(validation.parse_duration),
        validate_ms: duration_ms(validation.validation_duration),
        compose_ms: 0.0,
        docx_ms: 0.0,
        pdf_ms: 0.0,
    };
    if !validation.valid {
        return ProtocolResponse {
            protocol_version: PROTOCOL_VERSION,
            request_id,
            ok: false,
            diagnostics: validation.diagnostics,
            artifacts: Vec::new(),
            timings: Some(validation_timings),
            error: Some(ProtocolError {
                code: "validation_failed".to_string(),
                message: "document spec contains validation errors".to_string(),
            }),
        };
    }
    if request.operation == ProtocolOperation::Validate {
        return ProtocolResponse {
            protocol_version: PROTOCOL_VERSION,
            request_id,
            ok: true,
            diagnostics: validation.diagnostics,
            artifacts: Vec::new(),
            timings: Some(validation_timings),
            error: None,
        };
    }

    let output = request.output.expect("render output checked by envelope");
    let destination = match safe_output_directory(output_root.as_ref(), &output.directory) {
        Ok(path) => path,
        Err(message) => return protocol_failure(request_id, "invalid_output", message),
    };
    let name = output
        .name
        .unwrap_or_else(|| default_artifact_name(&request.source));
    if !is_safe_artifact_name(&name) {
        return protocol_failure(
            request_id,
            "invalid_output",
            "output name must use only ASCII letters, digits, '-' or '_'".to_string(),
        );
    }
    let rendered = match renderer.render(&render_request) {
        Ok(rendered) => rendered,
        Err(error) => return protocol_failure(request_id, "render_failed", error.to_string()),
    };
    if let Err(error) = std::fs::create_dir_all(&destination) {
        return protocol_failure(request_id, "write_failed", error.to_string());
    }
    let docx_path = destination.join(format!("{name}.docx"));
    if let Err(error) = atomic_write_file(&docx_path, &rendered.docx) {
        return protocol_failure(request_id, "write_failed", error.to_string());
    }
    let mut artifacts = vec![protocol_artifact("docx", &docx_path, &rendered.docx)];
    if let Some(pdf) = &rendered.pdf {
        let pdf_path = destination.join(format!("{name}.pdf"));
        if let Err(error) = atomic_write_file(&pdf_path, pdf) {
            return protocol_failure(request_id, "write_failed", error.to_string());
        }
        artifacts.push(protocol_artifact("pdf", &pdf_path, pdf));
    }
    ProtocolResponse {
        protocol_version: PROTOCOL_VERSION,
        request_id,
        ok: true,
        diagnostics: rendered.diagnostics,
        artifacts,
        timings: Some(ProtocolTimings {
            parse_ms: duration_ms(rendered.parse_duration),
            validate_ms: duration_ms(rendered.validation_duration),
            compose_ms: duration_ms(rendered.compose_duration),
            docx_ms: duration_ms(rendered.docx_duration),
            pdf_ms: duration_ms(rendered.pdf_duration),
        }),
        error: None,
    }
}

/// Build a response for malformed transport input that could not be deserialized.
pub fn protocol_failure(
    request_id: impl Into<String>,
    code: impl Into<String>,
    message: impl Into<String>,
) -> ProtocolResponse {
    ProtocolResponse {
        protocol_version: PROTOCOL_VERSION,
        request_id: request_id.into(),
        ok: false,
        diagnostics: Vec::new(),
        artifacts: Vec::new(),
        timings: None,
        error: Some(ProtocolError {
            code: code.into(),
            message: message.into(),
        }),
    }
}

fn validate_request_envelope(request: &ProtocolRequest) -> std::result::Result<(), String> {
    if request.protocol_version != PROTOCOL_VERSION {
        return Err(format!(
            "unsupported protocol_version {}; expected {PROTOCOL_VERSION}",
            request.protocol_version
        ));
    }
    if request.request_id.is_empty()
        || request.request_id.chars().count() > 128
        || request
            .request_id
            .chars()
            .any(|character| character.is_control())
    {
        return Err("request_id must contain 1-128 printable characters".to_string());
    }
    if request.operation == ProtocolOperation::Render && request.output.is_none() {
        return Err("render requests require an output object".to_string());
    }
    Ok(())
}

fn safe_output_directory(root: &Path, relative: &Path) -> std::result::Result<PathBuf, String> {
    if relative.is_absolute()
        || relative.components().any(|component| {
            matches!(
                component,
                Component::ParentDir | Component::RootDir | Component::Prefix(_)
            )
        })
    {
        return Err("output directory must stay relative to the service output root".to_string());
    }
    let root = if root.is_absolute() {
        root.to_path_buf()
    } else {
        std::env::current_dir()
            .map_err(|error| error.to_string())?
            .join(root)
    };
    std::fs::create_dir_all(&root).map_err(|error| error.to_string())?;
    let root = root.canonicalize().map_err(|error| error.to_string())?;
    let mut destination = root.clone();
    for component in relative.components() {
        if let Component::Normal(segment) = component {
            destination.push(segment);
            if std::fs::symlink_metadata(&destination)
                .is_ok_and(|metadata| metadata.file_type().is_symlink())
            {
                return Err("output directory must not traverse a symbolic link".to_string());
            }
            if destination.exists() {
                let canonical = destination
                    .canonicalize()
                    .map_err(|error| error.to_string())?;
                if !canonical.starts_with(&root) {
                    return Err("output directory escapes the service output root".to_string());
                }
            }
        }
    }
    Ok(destination)
}

fn is_safe_artifact_name(value: &str) -> bool {
    !value.is_empty()
        && value.len() <= 128
        && value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_'))
}

fn default_artifact_name(source: &RenderSource) -> String {
    match source {
        RenderSource::Path { path } => path
            .file_stem()
            .and_then(|stem| stem.to_str())
            .filter(|stem| is_safe_artifact_name(stem))
            .unwrap_or("document")
            .to_string(),
        RenderSource::Inline { .. } => "document".to_string(),
    }
}

fn protocol_artifact(kind: &str, path: &Path, bytes: &[u8]) -> ProtocolArtifact {
    ProtocolArtifact {
        kind: kind.to_string(),
        path: path.display().to_string(),
        sha256: format!("{:x}", Sha256::digest(bytes)),
        bytes: bytes.len(),
    }
}

fn duration_ms(duration: std::time::Duration) -> f64 {
    duration.as_secs_f64() * 1000.0
}

fn default_output_directory() -> PathBuf {
    PathBuf::from(".")
}

const fn default_true() -> bool {
    true
}

#[cfg(test)]
mod tests {
    use tempfile::tempdir;

    use super::{execute_protocol_request, ProtocolOperation, ProtocolOutput, ProtocolRequest};
    use crate::renderer::{RenderSource, SpecFormat};

    fn request(operation: ProtocolOperation) -> ProtocolRequest {
        ProtocolRequest {
            protocol_version: 1,
            request_id: "test-1".to_string(),
            operation,
            source: RenderSource::Inline {
                format: SpecFormat::Yaml,
                content:
                    "version: 1\noutput_name: protocol\nblocks:\n  - type: body\n    text: Hello\n"
                        .to_string(),
            },
            config: None,
            output: None,
        }
    }

    #[test]
    fn validate_is_side_effect_free() {
        let root = tempdir().expect("temp dir");
        let response = execute_protocol_request(request(ProtocolOperation::Validate), root.path());
        assert!(response.ok);
        assert!(response.artifacts.is_empty());
        assert_eq!(
            std::fs::read_dir(root.path()).expect("read root").count(),
            0
        );
    }

    #[test]
    fn render_writes_hash_described_artifacts_beneath_root() {
        let root = tempdir().expect("temp dir");
        let mut request = request(ProtocolOperation::Render);
        request.output = Some(ProtocolOutput {
            directory: "nested".into(),
            name: Some("hello".to_string()),
            pdf: true,
        });
        let response = execute_protocol_request(request, root.path());
        assert!(response.ok, "{:?}", response.error);
        assert_eq!(response.artifacts.len(), 2);
        assert!(root.path().join("nested/hello.docx").is_file());
        assert!(root.path().join("nested/hello.pdf").is_file());
    }

    #[test]
    fn render_rejects_output_escape() {
        let root = tempdir().expect("temp dir");
        let mut request = request(ProtocolOperation::Render);
        request.output = Some(ProtocolOutput {
            directory: "../outside".into(),
            name: Some("hello".to_string()),
            pdf: false,
        });
        let response = execute_protocol_request(request, root.path());
        assert!(!response.ok);
        assert_eq!(response.error.expect("error").code, "invalid_output");
    }

    #[cfg(unix)]
    #[test]
    fn render_rejects_symlink_escape_beneath_output_root() {
        use std::os::unix::fs::symlink;

        let root = tempdir().expect("root");
        let outside = tempdir().expect("outside");
        symlink(outside.path(), root.path().join("linked")).expect("symlink");
        let mut request = request(ProtocolOperation::Render);
        request.output = Some(ProtocolOutput {
            directory: "linked/nested".into(),
            name: Some("hello".to_string()),
            pdf: false,
        });
        let response = execute_protocol_request(request, root.path());
        assert!(!response.ok);
        assert_eq!(response.error.expect("error").code, "invalid_output");
        assert!(!outside.path().join("nested/hello.docx").exists());
    }
}
