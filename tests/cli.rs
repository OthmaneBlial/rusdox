use std::fs;
use std::path::Path;
use std::process::Command;

use tempfile::tempdir;

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
