use std::fs;
use std::io::{BufRead, BufReader, Cursor, Read, Write};
use std::net::{TcpListener, TcpStream};
use std::path::Path;
use std::process::{Command, Stdio};
use std::thread;
use std::time::Duration;

use ed25519_dalek::{Signer, SigningKey};
use sha2::{Digest, Sha256};
use tempfile::tempdir;
use zip::ZipArchive;

fn rusdox_bin() -> &'static str {
    env!("CARGO_BIN_EXE_rusdox")
}

fn run_cli(args: &[&str], cwd: &Path) -> std::process::Output {
    Command::new(rusdox_bin())
        .args(args)
        .current_dir(cwd)
        .output()
        .expect("failed to run rusdox binary")
}

fn spawn_cli(args: &[&str], cwd: &Path) -> std::process::Child {
    Command::new(rusdox_bin())
        .args(args)
        .current_dir(cwd)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("failed to spawn rusdox binary")
}

fn script_source() -> &'static str {
    r#"use rusdox::{Document, Paragraph, Run};
use rusdox::studio::Studio;

pub fn build_document(_studio: &Studio) -> rusdox::Result<Document> {
    let mut doc = Document::new();
    doc.push_paragraph(Paragraph::new().add_run(Run::from_text("hello")));
    Ok(doc)
}
"#
}

fn spec_source() -> &'static str {
    r#"output_name: mydoc
blocks:
  - type: title
    text: Hello from YAML
  - type: body
    text: This file should render without Rust source.
"#
}

fn create_signed_local_registry(directory: &Path) -> (std::path::PathBuf, String) {
    let repo = Path::new(env!("CARGO_MANIFEST_DIR"));
    let mut registry: serde_json::Value = serde_json::from_slice(
        &fs::read(repo.join("registry/v1/index.json")).expect("read registry"),
    )
    .expect("registry JSON");
    let base_url = registry["base_url"].as_str().expect("base URL").to_string();
    for entry in registry["templates"]
        .as_array_mut()
        .expect("template entries")
    {
        for pointer in [
            "/preview/url",
            "/files/template/url",
            "/files/sample_data/url",
            "/verified_outputs/docx/url",
            "/verified_outputs/pdf/url",
            "/verified_outputs/parity_json/url",
            "/verified_outputs/parity_html/url",
        ] {
            let url = entry
                .pointer_mut(pointer)
                .and_then(|value| value.as_str())
                .expect("asset URL")
                .to_string();
            let relative = url.strip_prefix(&base_url).expect("same-origin asset");
            let local_path = repo.join(relative);
            let local_sha256 = format!(
                "{:x}",
                Sha256::digest(fs::read(&local_path).expect("local registry asset"))
            );
            *entry.pointer_mut(pointer).expect("asset pointer") =
                serde_json::Value::String(local_path.display().to_string());
            let hash_pointer = pointer.replace("/url", "/sha256");
            *entry
                .pointer_mut(&hash_pointer)
                .expect("asset hash pointer") = serde_json::Value::String(local_sha256);
        }
    }
    let index = serde_json::to_vec_pretty(&registry).expect("serialize registry");
    let seed = [37_u8; 32];
    let signing_key = SigningKey::from_bytes(&seed);
    let signature = signing_key.sign(&index);
    let public_key = signing_key.verifying_key().to_bytes();
    let public_key_hex = public_key
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    let signature_hex = signature
        .to_bytes()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    let digest = format!("{:x}", Sha256::digest(&index));
    let index_path = directory.join("index.json");
    fs::write(&index_path, &index).expect("local registry");
    fs::write(
        directory.join("index.sig.json"),
        serde_json::to_vec_pretty(&serde_json::json!({
            "schema_version": 1,
            "algorithm": "ed25519",
            "key_id": "integration-test",
            "manifest_sha256": digest,
            "signature": signature_hex,
        }))
        .expect("signature JSON"),
    )
    .expect("signature file");
    (index_path, public_key_hex)
}

#[test]
fn init_doc_creates_template_file() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    let output = run_cli(
        &["init-doc", spec_path.to_string_lossy().as_ref()],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let content = fs::read_to_string(&spec_path).expect("spec should exist");
    assert!(content.contains("output_name: my-document"));
    assert!(content.contains("type: title"));
}

#[test]
fn demo_creates_verified_artifacts_and_refuses_an_existing_destination() {
    let temp = tempdir().expect("temp dir");
    let demo_root = temp.path().join("first-result");
    let demo_root_arg = demo_root.to_string_lossy();
    let output = run_cli(&["demo", demo_root_arg.as_ref()], temp.path());
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    for relative in [
        "product-launch-brief.yaml",
        "rusdox.toml",
        "generated/product-launch-brief.docx",
        "rendered/product-launch-brief.pdf",
        "reports/product-launch-brief-parity.html",
        "reports/product-launch-brief-parity.json",
        "reports/product-launch-brief-pages/page-001.png",
    ] {
        let path = demo_root.join(relative);
        assert!(path.is_file(), "expected {}", path.display());
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("parity verification passed"));
    assert!(stdout.contains("Demo ready"));
    assert!(stdout.contains("rusdox verify"));

    let spec_before = fs::read(demo_root.join("product-launch-brief.yaml")).expect("demo spec");
    let repeated = run_cli(&["demo", demo_root_arg.as_ref()], temp.path());
    assert!(!repeated.status.success());
    assert!(String::from_utf8_lossy(&repeated.stderr).contains("already exists"));
    assert_eq!(
        fs::read(demo_root.join("product-launch-brief.yaml")).expect("preserved demo spec"),
        spec_before
    );
}

#[test]
fn init_script_creates_template_file() {
    let temp = tempdir().expect("temp dir");
    let script_path = temp.path().join("mydoc.rs");
    let output = run_cli(
        &["init-script", script_path.to_string_lossy().as_ref()],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let content = fs::read_to_string(&script_path).expect("script should exist");
    assert!(content.contains("pub fn build_document("));
    assert!(content.contains("rusdox mydoc.rs"));
}

#[test]
fn init_script_refuses_overwrite_without_force() {
    let temp = tempdir().expect("temp dir");
    let script_path = temp.path().join("mydoc.rs");
    fs::write(&script_path, "existing").expect("write existing");

    let output = run_cli(
        &["init-script", script_path.to_string_lossy().as_ref()],
        temp.path(),
    );
    assert!(!output.status.success());
    assert!(String::from_utf8_lossy(&output.stderr).contains("already exists"));
}

#[test]
fn schema_command_emits_versioned_authoring_contract() {
    let temp = tempdir().expect("temp dir");
    let output = run_cli(&["schema"], temp.path());
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let schema: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("generated JSON Schema");
    assert_eq!(schema["x-rusdox-spec-version"], 1);
    assert_eq!(schema["properties"]["version"]["const"], 1);
    assert!(schema["required"]
        .as_array()
        .is_some_and(|required| required.contains(&serde_json::json!("version"))));
    assert!(String::from_utf8_lossy(&output.stdout).contains("\"when\""));
}

#[test]
fn migrate_upgrades_yaml_atomically_and_check_detects_current_version() {
    let temp = tempdir().expect("temp dir");
    let spec = temp.path().join("legacy.rusdox.yaml");
    fs::write(
        &spec,
        "# keep this author comment\noutput_name: legacy\nblocks: []\n",
    )
    .expect("write legacy spec");

    let check = run_cli(
        &["migrate", spec.to_string_lossy().as_ref(), "--check"],
        temp.path(),
    );
    assert!(!check.status.success());
    assert!(String::from_utf8_lossy(&check.stderr).contains("needs migration"));

    let migrate = run_cli(
        &["migrate", spec.to_string_lossy().as_ref(), "--in-place"],
        temp.path(),
    );
    assert!(
        migrate.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&migrate.stderr)
    );
    let migrated = fs::read_to_string(&spec).expect("migrated spec");
    assert!(migrated.starts_with("# keep this author comment\nversion: 1\n"));

    let current = run_cli(
        &["migrate", spec.to_string_lossy().as_ref(), "--check"],
        temp.path(),
    );
    assert!(current.status.success());
}

#[test]
fn validate_json_reports_source_line_and_column() {
    let temp = tempdir().expect("temp dir");
    let spec = temp.path().join("future.rusdox.yaml");
    fs::write(&spec, "version: 9\nblocks: []\n").expect("write future spec");
    let output = run_cli(
        &[
            "validate",
            spec.to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(!output.status.success());
    let report: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("validation JSON");
    let issue = &report["files"][0]["issues"][0];
    assert_eq!(issue["path"], "version");
    assert_eq!(issue["source"]["line"], 1);
    assert_eq!(issue["source"]["column"], 1);
}

#[test]
fn config_init_and_show_work_with_explicit_path() {
    let temp = tempdir().expect("temp dir");
    let config_path = temp.path().join("cfg.toml");

    let init = run_cli(
        &[
            "config",
            "init",
            "--path",
            config_path.to_string_lossy().as_ref(),
            "--force",
        ],
        temp.path(),
    );
    assert!(
        init.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&init.stderr)
    );
    assert!(config_path.exists());

    let show = run_cli(
        &[
            "config",
            "show",
            "--path",
            config_path.to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        show.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&show.stderr)
    );
    let stdout = String::from_utf8_lossy(&show.stdout);
    assert!(stdout.contains("\"profile_name\""));
}

#[test]
fn run_script_docx_only_writes_docx_without_pdf() {
    let temp = tempdir().expect("temp dir");
    let script_path = temp.path().join("mydoc.rs");
    fs::write(&script_path, script_source()).expect("write script");

    let output = run_cli(
        &[script_path.to_string_lossy().as_ref(), "--docx-only"],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let docx_path = temp.path().join("mydoc.docx");
    let pdf_path = temp.path().join("mydoc.pdf");
    assert!(docx_path.exists(), "expected {}", docx_path.display());
    assert!(!pdf_path.exists(), "did not expect {}", pdf_path.display());
}

#[test]
fn run_spec_docx_only_writes_docx_without_pdf() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    fs::write(&spec_path, spec_source()).expect("write spec");

    let output = run_cli(
        &[spec_path.to_string_lossy().as_ref(), "--docx-only"],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let docx_path = temp.path().join("generated").join("mydoc.docx");
    let pdf_path = temp.path().join("rendered").join("mydoc.pdf");
    assert!(docx_path.exists(), "expected {}", docx_path.display());
    assert!(!pdf_path.exists(), "did not expect {}", pdf_path.display());
}

#[test]
fn run_script_with_pdf_overrides_config_emit_flag() {
    let temp = tempdir().expect("temp dir");
    let script_path = temp.path().join("mydoc.rs");
    fs::write(&script_path, script_source()).expect("write script");

    let config_path = temp.path().join("rusdox.toml");
    fs::write(
        &config_path,
        r#"
[output]
emit_pdf_preview = false
"#,
    )
    .expect("write config");

    let output = run_cli(
        &[
            script_path.to_string_lossy().as_ref(),
            "--config",
            config_path.to_string_lossy().as_ref(),
            "--with-pdf",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let docx_path = temp.path().join("mydoc.docx");
    let pdf_path = temp.path().join("mydoc.pdf");
    assert!(docx_path.exists(), "expected {}", docx_path.display());
    assert!(pdf_path.exists(), "expected {}", pdf_path.display());
}

#[test]
fn run_script_rejects_missing_entrypoint() {
    let temp = tempdir().expect("temp dir");
    let script_path = temp.path().join("broken.rs");
    fs::write(&script_path, "fn nope() {}").expect("write script");

    let output = run_cli(&[script_path.to_string_lossy().as_ref()], temp.path());
    assert!(!output.status.success());
    assert!(String::from_utf8_lossy(&output.stderr).contains("build_document"));
}

#[test]
fn run_script_rejects_non_rs_file() {
    let temp = tempdir().expect("temp dir");
    let script_path = temp.path().join("doc.txt");
    fs::write(&script_path, "not rust").expect("write file");

    let output = run_cli(&[script_path.to_string_lossy().as_ref()], temp.path());
    assert!(!output.status.success());
    assert!(String::from_utf8_lossy(&output.stderr).contains("unsupported input type"));
}

#[test]
fn run_script_rejects_missing_file() {
    let temp = tempdir().expect("temp dir");
    let missing = temp.path().join("missing.rs");

    let output = run_cli(&[missing.to_string_lossy().as_ref()], temp.path());
    assert!(!output.status.success());
    assert!(String::from_utf8_lossy(&output.stderr).contains("input not found"));
}

#[test]
fn run_example_spec_with_visual_assets_resolves_paths_relative_to_spec() {
    let temp = tempdir().expect("temp dir");
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let spec_path = manifest_dir.join("examples/visual_assets_showcase.yaml");
    let output_docx = temp.path().join("visual-assets.docx");

    let output = run_cli(
        &[
            spec_path.to_string_lossy().as_ref(),
            "--output",
            output_docx.to_string_lossy().as_ref(),
            "--with-pdf",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let pdf_path = temp.path().join("rendered").join("visual-assets.pdf");
    assert!(output_docx.exists(), "expected {}", output_docx.display());
    assert!(pdf_path.exists(), "expected {}", pdf_path.display());

    let docx = fs::read(&output_docx).expect("read docx");
    let pdf = fs::read(&pdf_path).expect("read pdf");
    assert!(docx
        .windows("word/media/".len())
        .any(|window| window == b"word/media/"));
    assert!(pdf.starts_with(b"%PDF-"));
    assert!(
        pdf.len() > 2_000,
        "expected rendered pdf to contain real content"
    );
}

#[test]
fn run_example_spec_with_named_styles_emits_style_parts_and_pdf() {
    let temp = tempdir().expect("temp dir");
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let spec_path = manifest_dir.join("examples/named_styles_showcase.yaml");
    let output_docx = temp.path().join("named-styles.docx");

    let output = run_cli(
        &[
            spec_path.to_string_lossy().as_ref(),
            "--output",
            output_docx.to_string_lossy().as_ref(),
            "--with-pdf",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let pdf_path = temp.path().join("rendered").join("named-styles.pdf");
    assert!(output_docx.exists(), "expected {}", output_docx.display());
    assert!(pdf_path.exists(), "expected {}", pdf_path.display());

    let docx = fs::read(&output_docx).expect("read docx");
    let mut archive = ZipArchive::new(Cursor::new(docx)).expect("open docx zip");

    let mut styles_xml = String::new();
    archive
        .by_name("word/styles.xml")
        .expect("styles part should exist")
        .read_to_string(&mut styles_xml)
        .expect("read styles xml");
    assert!(styles_xml.contains(r#"w:styleId="cover_title""#));
    assert!(styles_xml.contains(r#"w:styleId="accent""#));
    assert!(styles_xml.contains(r#"w:styleId="dashboard_grid""#));

    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document part should exist")
        .read_to_string(&mut document_xml)
        .expect("read document xml");
    assert!(document_xml.contains(r#"<w:pStyle w:val="cover_title"/>"#));
    assert!(document_xml.contains(r#"<w:rStyle w:val="accent"/>"#));
    assert!(document_xml.contains(r#"<w:tblStyle w:val="dashboard_grid"/>"#));

    let pdf = fs::read(&pdf_path).expect("read pdf");
    assert!(pdf.starts_with(b"%PDF-"));
    assert!(
        pdf.len() > 2_000,
        "expected rendered pdf to contain real content"
    );
}

#[test]
fn run_example_spec_with_yaml_composition_emits_metadata_and_pdf() {
    let temp = tempdir().expect("temp dir");
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let spec_path = manifest_dir.join("examples/yaml_composition_showcase.yaml");
    let output_docx = temp.path().join("yaml-composition.docx");

    let output = run_cli(
        &[
            spec_path.to_string_lossy().as_ref(),
            "--output",
            output_docx.to_string_lossy().as_ref(),
            "--with-pdf",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let pdf_path = temp.path().join("rendered").join("yaml-composition.pdf");
    assert!(output_docx.exists(), "expected {}", output_docx.display());
    assert!(pdf_path.exists(), "expected {}", pdf_path.display());

    let docx = fs::read(&output_docx).expect("read docx");
    let mut archive = ZipArchive::new(Cursor::new(docx)).expect("open docx zip");

    let mut core_xml = String::new();
    archive
        .by_name("docProps/core.xml")
        .expect("core properties should exist")
        .read_to_string(&mut core_xml)
        .expect("read core xml");
    assert!(core_xml.contains("<dc:title>Northwind Health Rollout Plan</dc:title>"));
    assert!(core_xml.contains("<dc:subject>Q4 2026 regional rollout</dc:subject>"));
    assert!(core_xml.contains("<cp:keywords>yaml, composition, Q4 2026</cp:keywords>"));

    let mut custom_xml = String::new();
    archive
        .by_name("docProps/custom.xml")
        .expect("custom properties should exist")
        .read_to_string(&mut custom_xml)
        .expect("read custom xml");
    assert!(custom_xml.contains(r#"name="Client""#));
    assert!(custom_xml.contains("Northwind Health"));
    assert!(custom_xml.contains(r#"name="Sponsor""#));
    assert!(custom_xml.contains("Maya Chen"));

    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document part should exist")
        .read_to_string(&mut document_xml)
        .expect("read document xml");
    assert!(document_xml.contains("North America"));
    assert!(document_xml.contains("EMEA"));
    assert!(document_xml.contains("APAC"));

    let pdf = fs::read(&pdf_path).expect("read pdf");
    assert!(pdf.starts_with(b"%PDF-"));
}

#[test]
fn validate_reports_semantic_errors_for_invalid_spec() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("invalid.yaml");
    fs::write(
        &spec_path,
        r##"output_name: invalid
styles:
  run:
    - id: accent
      properties:
        color: "#AA5500"
blocks:
  - type: paragraph
    spec:
      runs:
        - text: Broken
          color: XYZ123
  - type: table
    spec:
      columns:
        - label: Only
          width: 1200
      rows:
        - cells:
            - kind: text
              text: A
            - kind: text
              text: B
"##,
    )
    .expect("write invalid spec");

    let output = run_cli(
        &["validate", spec_path.to_string_lossy().as_ref()],
        temp.path(),
    );
    assert!(
        !output.status.success(),
        "stdout: {}",
        String::from_utf8_lossy(&output.stdout)
    );

    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("invalid color '#AA5500'"));
    assert!(stderr.contains("invalid color 'XYZ123'"));
    assert!(stderr.contains("row has 2 cells but the table only defines 1 columns"));
}

#[test]
fn validate_json_reports_success_for_valid_spec() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    fs::write(&spec_path, spec_source()).expect("write spec");

    let output = run_cli(
        &[
            "validate",
            spec_path.to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let json: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("valid validation json");
    assert_eq!(json["errors"], 0);
    assert_eq!(json["warnings"], 0);
    assert_eq!(json["specs"], 1);
}

#[test]
fn validate_reports_config_errors_in_json_mode() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    let config_path = temp.path().join("rusdox.toml");
    fs::write(&spec_path, spec_source()).expect("write spec");
    fs::write(
        &config_path,
        r##"
[colors]
accent = "#12GG45"
"##,
    )
    .expect("write invalid config");

    let output = run_cli(
        &[
            "validate",
            spec_path.to_string_lossy().as_ref(),
            "--config",
            config_path.to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        !output.status.success(),
        "stdout: {}",
        String::from_utf8_lossy(&output.stdout)
    );

    let json: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("valid validation json");
    assert_eq!(json["specs"], 1);
    assert!(json["errors"].as_u64().unwrap_or_default() >= 1);
    assert!(json["config_issues"]
        .as_array()
        .expect("config_issues array")
        .iter()
        .any(|issue| issue["message"]
            == "invalid color '#12GG45', expected six hex digits without '#'"));
}

#[test]
fn render_rejects_semantic_validation_errors_before_writing_output() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("bad.yaml");
    fs::write(
        &spec_path,
        r#"output_name: bad
blocks:
  - type: image
    path: missing.png
"#,
    )
    .expect("write invalid spec");

    let output = run_cli(&[spec_path.to_string_lossy().as_ref()], temp.path());
    assert!(!output.status.success());

    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("rendering aborted because the spec has validation errors"));
    assert!(stderr.contains("visual asset does not exist"));
    assert!(!temp.path().join("generated").join("bad.docx").exists());
}

#[test]
fn bench_outputs_json_summary_without_leaving_artifacts_by_default() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    fs::write(&spec_path, spec_source()).expect("write spec");

    let output = run_cli(
        &[
            "bench",
            spec_path.to_string_lossy().as_ref(),
            "--docx-only",
            "--iterations",
            "2",
            "--warmup",
            "1",
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let json: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("valid benchmark json");
    assert_eq!(json["schema_version"], 2);
    assert_eq!(json["pipeline"], "docx");
    assert_eq!(
        json["input_sha256"].as_str().map(str::len),
        Some(64),
        "input hash should be a full SHA-256"
    );
    assert_eq!(json["specs"], 1);
    assert_eq!(json["iterations"], 2);
    assert_eq!(json["warmup"], 1);
    assert_eq!(json["emit_pdf"], false);
    assert!(json["parse_ms"]["avg"].as_f64().unwrap_or_default() >= 0.0);
    assert!(json["validate_ms"]["avg"].as_f64().unwrap_or_default() >= 0.0);
    assert!(json["compose_ms"]["avg"].as_f64().unwrap_or_default() >= 0.0);
    assert!(json["docx_ms"]["avg"].as_f64().unwrap_or_default() >= 0.0);
    assert!(json["docx_ms"]["median"].as_f64().unwrap_or_default() >= 0.0);
    assert_eq!(json["pdf_ms"]["avg"].as_f64().unwrap_or_default(), 0.0);
    assert!(!temp.path().join("generated").exists());
    assert!(!temp.path().join("rendered").exists());
}

#[test]
fn bench_isolates_pdf_dual_validation_and_existing_docx_pipelines() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    fs::write(&spec_path, spec_source()).expect("write spec");

    for pipeline in ["validation", "pdf", "dual"] {
        let output = run_cli(
            &[
                "bench",
                spec_path.to_string_lossy().as_ref(),
                "--pipeline",
                pipeline,
                "--iterations",
                "1",
                "--format",
                "json",
            ],
            temp.path(),
        );
        assert!(
            output.status.success(),
            "{pipeline} stderr: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let json: serde_json::Value =
            serde_json::from_slice(&output.stdout).expect("valid benchmark json");
        assert_eq!(json["pipeline"], pipeline);
        match pipeline {
            "validation" => {
                assert_eq!(json["docx_bytes"]["avg"], 0.0);
                assert_eq!(json["pdf_bytes"]["avg"], 0.0);
            }
            "pdf" => {
                assert_eq!(json["docx_bytes"]["avg"], 0.0);
                assert!(json["pdf_bytes"]["avg"].as_f64().unwrap_or_default() > 0.0);
            }
            "dual" => {
                assert!(json["docx_bytes"]["avg"].as_f64().unwrap_or_default() > 0.0);
                assert!(json["pdf_bytes"]["avg"].as_f64().unwrap_or_default() > 0.0);
            }
            _ => unreachable!(),
        }
    }

    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/external-macos-textutil.docx");
    let output = run_cli(
        &[
            "bench",
            fixture.to_string_lossy().as_ref(),
            "--pipeline",
            "existing-docx",
            "--iterations",
            "1",
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "existing-docx stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let json: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("valid benchmark json");
    assert_eq!(json["pipeline"], "existing-docx");
    assert!(
        json["existing_docx_open_ms"]["avg"]
            .as_f64()
            .unwrap_or_default()
            > 0.0
    );
    assert!(
        json["existing_docx_save_ms"]["avg"]
            .as_f64()
            .unwrap_or_default()
            > 0.0
    );
}

#[test]
fn bench_keep_output_writes_artifacts_when_requested() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    fs::write(&spec_path, spec_source()).expect("write spec");

    let output = run_cli(
        &[
            "bench",
            spec_path.to_string_lossy().as_ref(),
            "--docx-only",
            "--iterations",
            "1",
            "--keep-output",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("benchmark target:"));
    assert!(temp.path().join("generated").join("mydoc.docx").exists());
    assert!(!temp.path().join("rendered").join("mydoc.pdf").exists());
}

#[test]
fn template_cli_inspects_and_verifies_word_native_docx() {
    let temp = tempdir().expect("temp dir");
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let template = root.join("templates/proposal/template.docx");
    let data = root.join("templates/proposal/data.json");

    let inspect = run_cli(
        &[
            "template",
            "inspect",
            template.to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        inspect.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&inspect.stderr)
    );
    let inspection: serde_json::Value =
        serde_json::from_slice(&inspect.stdout).expect("inspection JSON");
    assert_eq!(inspection["syntax_version"], "1");
    assert!(inspection["placeholders"].as_array().is_some_and(|items| {
        items
            .iter()
            .any(|item| item["expression"] == "proposal.title")
    }));
    assert_eq!(inspection["diagnostics"].as_array().map(Vec::len), Some(0));

    let verify = run_cli(
        &[
            "template",
            "verify",
            template.to_string_lossy().as_ref(),
            data.to_string_lossy().as_ref(),
            "--name",
            "proposal",
            "--strict",
            "--output-root",
            temp.path().to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        verify.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&verify.stderr)
    );
    let result: serde_json::Value =
        serde_json::from_slice(&verify.stdout).expect("template result JSON");
    assert_eq!(result["command"], "verify");
    assert_eq!(result["passed"], true);
    assert_eq!(result["strict"], true);
    assert_eq!(result["checks"], 21);
    assert_eq!(result["failed_checks"], 0);
    assert!(result["replacements"].as_u64().unwrap_or_default() > 0);
    assert!(result["expanded_blocks"].as_u64().unwrap_or_default() > 0);
    assert!(temp.path().join("generated/proposal.docx").is_file());
    assert!(temp.path().join("rendered/proposal.pdf").is_file());
    assert!(temp.path().join("reports/proposal-parity.html").is_file());
    assert!(temp.path().join("reports/proposal-parity.json").is_file());
    assert!(temp
        .path()
        .join("reports/proposal-pages/page-001.png")
        .is_file());
}

#[test]
fn template_cli_strict_failure_preserves_previous_docx() {
    let temp = tempdir().expect("temp dir");
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let template = root.join("templates/proposal/template.docx");
    let data = temp.path().join("missing.json");
    let output = temp.path().join("generated/proposal.docx");
    fs::create_dir_all(output.parent().expect("output parent")).expect("create output parent");
    fs::write(&data, "{}").expect("write missing data");
    fs::write(&output, b"known-good-docx").expect("seed previous output");

    let render = run_cli(
        &[
            "template",
            "render",
            template.to_string_lossy().as_ref(),
            data.to_string_lossy().as_ref(),
            "--strict",
            "--output-root",
            temp.path().to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(!render.status.success());
    let report: serde_json::Value =
        serde_json::from_slice(&render.stdout).expect("strict render report JSON");
    assert_eq!(report["written"], false);
    assert!(report["diagnostics"].as_array().is_some_and(|diagnostics| {
        diagnostics.iter().any(|diagnostic| {
            diagnostic["part"] == "word/document.xml"
                && diagnostic["location"]
                    .as_str()
                    .is_some_and(|location| location.starts_with("paragraph "))
                && diagnostic["suggestion"]
                    .as_str()
                    .is_some_and(|suggestion| !suggestion.is_empty())
        })
    }));
    assert_eq!(
        fs::read(output).expect("read previous output"),
        b"known-good-docx"
    );
}

#[test]
fn template_registry_lists_and_installs_hash_verified_assets() {
    let temp = tempdir().expect("temp dir");
    let (registry, public_key) = create_signed_local_registry(temp.path());
    let registry_arg = registry.to_string_lossy();

    let list = run_cli(
        &[
            "template",
            "list",
            "--registry",
            registry_arg.as_ref(),
            "--public-key",
            &public_key,
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        list.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&list.stderr)
    );
    let listed: serde_json::Value = serde_json::from_slice(&list.stdout).expect("list JSON");
    assert_eq!(listed["signed"], true);
    assert_eq!(listed["templates"].as_array().map(Vec::len), Some(3));

    let install_root = temp.path().join("installed");
    let add = run_cli(
        &[
            "template",
            "add",
            "invoice",
            "--registry",
            registry_arg.as_ref(),
            "--public-key",
            &public_key,
            "--install-root",
            install_root.to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        add.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&add.stderr)
    );
    let result: serde_json::Value = serde_json::from_slice(&add.stdout).expect("install JSON");
    assert_eq!(result[0]["status"], "installed");
    assert!(install_root.join("invoice/template.docx").is_file());
    assert!(install_root.join("invoice/data.json").is_file());
    assert!(install_root.join("invoice/manifest.json").is_file());

    let second = run_cli(
        &[
            "template",
            "update",
            "invoice",
            "--registry",
            registry_arg.as_ref(),
            "--public-key",
            &public_key,
            "--install-root",
            install_root.to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(second.status.success());
    let result: serde_json::Value = serde_json::from_slice(&second.stdout).expect("update JSON");
    assert_eq!(result[0]["status"], "up-to-date");
}

#[test]
fn template_registry_rejects_tampering_before_install() {
    let temp = tempdir().expect("temp dir");
    let (registry, public_key) = create_signed_local_registry(temp.path());
    let mut bytes = fs::read(&registry).expect("registry");
    let position = bytes.iter().position(|byte| *byte == b'{').expect("JSON");
    bytes[position] = b' ';
    fs::write(&registry, bytes).expect("tamper registry");
    let install_root = temp.path().join("installed");
    let output = run_cli(
        &[
            "template",
            "add",
            "invoice",
            "--registry",
            registry.to_string_lossy().as_ref(),
            "--public-key",
            &public_key,
            "--install-root",
            install_root.to_string_lossy().as_ref(),
        ],
        temp.path(),
    );
    assert!(!output.status.success());
    assert!(String::from_utf8_lossy(&output.stderr).contains("SHA-256 does not match"));
    assert!(!install_root.exists());
}

#[test]
fn stdio_protocol_renders_one_bounded_json_request() {
    let temp = tempdir().expect("temp dir");
    let output_root = temp.path().join("service-output");
    let mut child = spawn_cli(
        &[
            "serve",
            "stdio",
            "--output-root",
            output_root.to_string_lossy().as_ref(),
            "--max-requests",
            "1",
        ],
        temp.path(),
    );
    let request = serde_json::json!({
        "protocol_version": 1,
        "request_id": "node-1",
        "operation": "render",
        "source": {
            "kind": "inline",
            "format": "yaml",
            "content": "version: 1\noutput_name: protocol\nblocks:\n  - type: body\n    text: Hello protocol\n"
        },
        "output": { "directory": "node", "name": "hello", "pdf": true }
    });
    writeln!(
        child.stdin.take().expect("stdin"),
        "{}",
        serde_json::to_string(&request).expect("request JSON")
    )
    .expect("request");
    let output = child.wait_with_output().expect("protocol output");
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let response: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("response JSON");
    assert_eq!(response["protocol_version"], 1);
    assert_eq!(response["request_id"], "node-1");
    assert_eq!(response["ok"], true);
    assert_eq!(response["artifacts"].as_array().map(Vec::len), Some(2));
    assert!(output_root.join("node/hello.docx").is_file());
    assert!(output_root.join("node/hello.pdf").is_file());
}

#[test]
fn stdio_protocol_applies_the_operator_limits_file() {
    let temp = tempdir().expect("temp dir");
    let output_root = temp.path().join("service-output");
    let profile_path = temp.path().join("limits.toml");
    let profile = fs::read_to_string(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("examples/config/hosted-limits.toml"),
    )
    .expect("hosted profile")
    .replace("max_spec_bytes = 2097152", "max_spec_bytes = 64");
    fs::write(&profile_path, profile).expect("custom limits");
    let mut child = spawn_cli(
        &[
            "serve",
            "stdio",
            "--limits-file",
            profile_path.to_string_lossy().as_ref(),
            "--output-root",
            output_root.to_string_lossy().as_ref(),
            "--max-requests",
            "1",
        ],
        temp.path(),
    );
    let request = serde_json::json!({
        "protocol_version": 1,
        "request_id": "limited-1",
        "operation": "validate",
        "source": {
            "kind": "inline",
            "format": "yaml",
            "content": format!("version: 1\nblocks: []\n# {}", "x".repeat(80))
        }
    });
    writeln!(
        child.stdin.take().expect("stdin"),
        "{}",
        serde_json::to_string(&request).expect("request JSON")
    )
    .expect("request");
    let output = child.wait_with_output().expect("protocol output");
    assert!(output.status.success());
    let response: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("response JSON");
    assert_eq!(response["ok"], false);
    assert_eq!(response["error"]["code"], "parse_error");
    assert!(
        response["error"]["message"]
            .as_str()
            .is_some_and(|message| message.contains("limit is 64 bytes")),
        "response: {response:#}"
    );
    assert!(!output_root.exists());
}

#[test]
fn loopback_http_protocol_reuses_the_v1_request_contract() {
    let temp = tempdir().expect("temp dir");
    let output_root = temp.path().join("service-output");
    let probe = TcpListener::bind("127.0.0.1:0").expect("reserve port");
    let port = probe.local_addr().expect("address").port();
    let port_string = port.to_string();
    drop(probe);
    let mut child = spawn_cli(
        &[
            "serve",
            "http",
            "--port",
            &port_string,
            "--output-root",
            output_root.to_string_lossy().as_ref(),
            "--max-requests",
            "1",
        ],
        temp.path(),
    );
    let mut readiness = String::new();
    BufReader::new(child.stderr.take().expect("stderr"))
        .read_line(&mut readiness)
        .expect("ready line");
    assert!(readiness.contains("protocol v1 listening"));

    let body = serde_json::to_vec(&serde_json::json!({
        "protocol_version": 1,
        "request_id": "http-1",
        "operation": "validate",
        "source": {
            "kind": "inline",
            "format": "yaml",
            "content": "version: 1\noutput_name: protocol\nblocks:\n  - type: body\n    text: Hello HTTP\n"
        }
    }))
    .expect("request JSON");
    let mut stream = TcpStream::connect(("127.0.0.1", port)).expect("connect");
    write!(
        stream,
        "POST /v1/request HTTP/1.1\r\nHost: 127.0.0.1\r\nContent-Type: application/json\r\nContent-Length: {}\r\nConnection: close\r\n\r\n",
        body.len()
    )
    .expect("headers");
    stream.write_all(&body).expect("body");
    let mut response = String::new();
    stream.read_to_string(&mut response).expect("response");
    assert!(response.starts_with("HTTP/1.1 200 OK"));
    let json = response.split("\r\n\r\n").nth(1).expect("response body");
    let parsed: serde_json::Value = serde_json::from_str(json).expect("response JSON");
    assert_eq!(parsed["request_id"], "http-1");
    assert_eq!(parsed["ok"], true);
    assert!(parsed["artifacts"].as_array().is_some_and(Vec::is_empty));
    assert!(child.wait().expect("server exit").success());
}

#[test]
fn loopback_http_protocol_rejects_non_json_content_type() {
    let temp = tempdir().expect("temp dir");
    let output_root = temp.path().join("service-output");
    let probe = TcpListener::bind("127.0.0.1:0").expect("reserve port");
    let port = probe.local_addr().expect("address").port();
    let port_string = port.to_string();
    drop(probe);
    let mut child = spawn_cli(
        &[
            "serve",
            "http",
            "--port",
            &port_string,
            "--output-root",
            output_root.to_string_lossy().as_ref(),
            "--max-requests",
            "1",
        ],
        temp.path(),
    );
    let mut readiness = String::new();
    BufReader::new(child.stderr.take().expect("stderr"))
        .read_line(&mut readiness)
        .expect("ready line");
    assert!(readiness.contains("protocol v1 listening"));

    let fixture = fs::read_to_string(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("contributor-fixtures/http/wrong-content-type.http"),
    )
    .expect("read HTTP fixture");
    let request = fixture.replace('\n', "\r\n");
    let mut stream = TcpStream::connect(("127.0.0.1", port)).expect("connect");
    stream
        .set_read_timeout(Some(Duration::from_secs(5)))
        .expect("read timeout");
    stream.write_all(request.as_bytes()).expect("request");
    let mut response = String::new();
    stream.read_to_string(&mut response).expect("response");

    assert!(response.starts_with("HTTP/1.1 415 Unsupported Media Type"));
    let json = response.split("\r\n\r\n").nth(1).expect("response body");
    let parsed: serde_json::Value = serde_json::from_str(json).expect("response JSON");
    assert_eq!(parsed["ok"], false);
    assert_eq!(parsed["error"]["code"], "unsupported_media_type");
    assert!(child.wait().expect("server exit").success());
}

#[test]
fn verify_writes_docx_pdf_html_json_and_page_evidence() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    let output_root = temp.path().join("evidence");
    fs::write(&spec_path, spec_source()).expect("write spec");

    let output = run_cli(
        &[
            "verify",
            spec_path.to_string_lossy().as_ref(),
            "--output-root",
            output_root.to_string_lossy().as_ref(),
            "--format",
            "json",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let summary: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("valid verify summary");
    assert_eq!(summary["passed"], true);
    assert_eq!(summary["specs"], 1);

    let docx = output_root.join("generated/mydoc.docx");
    let pdf = output_root.join("rendered/mydoc.pdf");
    let html = output_root.join("reports/mydoc-parity.html");
    let json = output_root.join("reports/mydoc-parity.json");
    let page = output_root.join("reports/mydoc-pages/page-001.png");
    for path in [&docx, &pdf, &html, &json, &page] {
        assert!(path.exists(), "expected {}", path.display());
    }

    let report: serde_json::Value =
        serde_json::from_slice(&fs::read(json).expect("read report")).expect("valid report json");
    assert_eq!(report["report_version"], "1");
    assert_eq!(report["passed"], true);
    assert_eq!(report["pdf"]["page_count"], 1);
    assert!(report["checks"]
        .as_array()
        .is_some_and(|checks| !checks.is_empty()));
    assert!(fs::read_to_string(html)
        .expect("read html")
        .contains("RusDox parity contract v1"));
}

#[test]
fn verify_uses_exit_code_two_for_a_completed_parity_failure() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    let output_root = temp.path().join("evidence");
    let empty_baseline = temp.path().join("empty-baseline");
    fs::write(&spec_path, spec_source()).expect("write spec");
    fs::create_dir_all(&empty_baseline).expect("create baseline dir");

    let output = run_cli(
        &[
            "verify",
            spec_path.to_string_lossy().as_ref(),
            "--output-root",
            output_root.to_string_lossy().as_ref(),
            "--visual-baseline",
            empty_baseline.to_string_lossy().as_ref(),
        ],
        temp.path(),
    );
    assert_eq!(output.status.code(), Some(2));
    assert!(String::from_utf8_lossy(&output.stderr).contains("parity verification failed"));
    assert!(output_root.join("reports/mydoc-parity.html").exists());
    assert!(output_root.join("reports/mydoc-parity.json").exists());
}

#[test]
fn watch_rebuilds_after_spec_changes() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    fs::write(&spec_path, spec_source()).expect("write spec");

    let child = spawn_cli(
        &[
            "watch",
            spec_path.to_string_lossy().as_ref(),
            "--docx-only",
            "--poll-interval-ms",
            "100",
            "--max-builds",
            "2",
        ],
        temp.path(),
    );

    let docx_path = temp.path().join("generated").join("mydoc.docx");
    for _ in 0..50 {
        if docx_path.exists() {
            break;
        }
        thread::sleep(Duration::from_millis(100));
    }
    assert!(docx_path.exists(), "expected {}", docx_path.display());

    fs::write(
        &spec_path,
        r#"output_name: mydoc
blocks:
  - type: title
    text: Hello from YAML
  - type: body
    text: Updated watch content.
"#,
    )
    .expect("update spec");

    let output = child.wait_with_output().expect("watch should exit");
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("watch build 1"));
    assert!(stdout.contains("watch build 2"));
    assert!(stdout.contains("change detected"));
    assert!(stdout.contains("succeeded"));
    assert!(docx_path.exists(), "expected {}", docx_path.display());
}

#[test]
fn watch_rebuilds_after_config_changes() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    let config_path = temp.path().join("rusdox.toml");
    fs::write(&spec_path, spec_source()).expect("write spec");
    fs::write(
        &config_path,
        r#"
[typography]
body_size_pt = 11
"#,
    )
    .expect("write config");

    let child = spawn_cli(
        &[
            "watch",
            spec_path.to_string_lossy().as_ref(),
            "--config",
            config_path.to_string_lossy().as_ref(),
            "--docx-only",
            "--poll-interval-ms",
            "100",
            "--max-builds",
            "2",
        ],
        temp.path(),
    );

    let docx_path = temp.path().join("generated").join("mydoc.docx");
    for _ in 0..50 {
        if docx_path.exists() {
            break;
        }
        thread::sleep(Duration::from_millis(100));
    }
    assert!(docx_path.exists(), "expected {}", docx_path.display());

    fs::write(
        &config_path,
        r#"
[typography]
body_size_pt = 13
"#,
    )
    .expect("update config");

    let output = child.wait_with_output().expect("watch should exit");
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("watch build 1"));
    assert!(stdout.contains("watch build 2"));
    assert!(stdout.contains(config_path.to_string_lossy().as_ref()));
    assert!(stdout.contains("succeeded"));
}

#[test]
fn dev_json_preserves_last_success_when_rebuild_fails() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    fs::write(&spec_path, spec_source()).expect("write spec");

    let child = spawn_cli(
        &[
            "dev",
            spec_path.to_string_lossy().as_ref(),
            "--docx-only",
            "--json",
            "--port",
            "0",
            "--poll-interval-ms",
            "50",
            "--debounce-ms",
            "120",
            "--max-builds",
            "2",
        ],
        temp.path(),
    );

    let docx_path = temp.path().join("generated").join("mydoc.docx");
    for _ in 0..80 {
        if docx_path.exists() {
            break;
        }
        thread::sleep(Duration::from_millis(50));
    }
    let successful_docx = fs::read(&docx_path).expect("initial successful DOCX");

    fs::write(
        &spec_path,
        r#"version: 2
output_name: mydoc
blocks:
  - type: title
    text: Invalid future version
"#,
    )
    .expect("write invalid spec");

    let output = child.wait_with_output().expect("dev should exit");
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(
        fs::read(&docx_path).expect("preserved successful DOCX"),
        successful_docx
    );

    let events = String::from_utf8(output.stdout)
        .expect("UTF-8 JSON Lines")
        .lines()
        .map(|line| serde_json::from_str::<serde_json::Value>(line).expect("valid JSON event"))
        .collect::<Vec<_>>();
    assert_eq!(events[0]["event"], "listening");
    assert_eq!(events[1]["status"], "success");
    assert_eq!(events[2]["status"], "failed");
    assert!(events[2]["reason"]
        .as_str()
        .is_some_and(|reason| reason.contains("input changed")));
    assert!(events[2]["error"]
        .as_str()
        .is_some_and(|error| error.contains("unsupported document spec version")));
    assert_eq!(events[2]["artifacts"], events[1]["artifacts"]);
}

#[test]
fn dev_quiet_mode_is_ci_friendly() {
    let temp = tempdir().expect("temp dir");
    let spec_path = temp.path().join("mydoc.yaml");
    fs::write(&spec_path, spec_source()).expect("write spec");
    let output = run_cli(
        &[
            "dev",
            spec_path.to_string_lossy().as_ref(),
            "--docx-only",
            "--quiet",
            "--port",
            "0",
            "--max-builds",
            "1",
        ],
        temp.path(),
    );
    assert!(
        output.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(output.stdout.is_empty());
    assert!(output.stderr.is_empty());
    assert!(temp.path().join("generated/mydoc.docx").is_file());
}
