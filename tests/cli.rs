use std::fs;
use std::io::{Cursor, Read};
use std::path::Path;
use std::process::{Command, Stdio};
use std::thread;
use std::time::Duration;

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
    assert_eq!(json["specs"], 1);
    assert_eq!(json["iterations"], 2);
    assert_eq!(json["warmup"], 1);
    assert_eq!(json["emit_pdf"], false);
    assert!(json["parse_ms"]["avg"].as_f64().unwrap_or_default() >= 0.0);
    assert!(json["validate_ms"]["avg"].as_f64().unwrap_or_default() >= 0.0);
    assert!(json["compose_ms"]["avg"].as_f64().unwrap_or_default() >= 0.0);
    assert!(json["docx_ms"]["avg"].as_f64().unwrap_or_default() >= 0.0);
    assert_eq!(json["pdf_ms"]["avg"].as_f64().unwrap_or_default(), 0.0);
    assert!(!temp.path().join("generated").exists());
    assert!(!temp.path().join("rendered").exists());
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
