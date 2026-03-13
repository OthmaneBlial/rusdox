use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

use clap::{Args, Parser, Subcommand, ValueEnum};
use dialoguer::{Confirm, Input, Select};
use rusdox::config::{default_user_config_path, RusdoxConfig};
use rusdox::spec::DocumentSpec;
use rusdox::studio::{Studio, DEFAULT_CONFIG_FILE};
use rusdox::{DocxError, Result};
use tempfile::tempdir;

#[derive(Debug, Parser)]
#[command(
    name = "rusdox",
    version,
    about = "RusDox CLI for document specs, configuration, and legacy script execution"
)]
struct Cli {
    /// Document spec file (.yaml/.yml/.json/.toml), spec directory, or legacy Rust script (.rs).
    #[arg(value_name = "INPUT")]
    input: Option<PathBuf>,
    /// Optional explicit output DOCX path for a single input file.
    #[arg(long)]
    output: Option<PathBuf>,
    /// Force DOCX-only generation (disable PDF output).
    #[arg(long)]
    docx_only: bool,
    /// Force PDF generation (overrides config if disabled).
    #[arg(long, conflicts_with = "docx_only")]
    with_pdf: bool,
    /// Optional config path for script execution.
    #[arg(long)]
    config: Option<PathBuf>,
    /// Build script in release mode.
    #[arg(long)]
    release: bool,
    #[command(subcommand)]
    command: Option<Commands>,
}

#[derive(Debug, Subcommand)]
enum Commands {
    /// Manage RusDox configuration files.
    Config {
        #[command(subcommand)]
        command: ConfigCommand,
    },
    /// Create a starter document spec compatible with `rusdox mydoc.yaml`.
    InitDoc(InitDocArgs),
    /// Create a starter script compatible with `rusdox mydoc.rs`.
    InitScript(InitScriptArgs),
}

#[derive(Debug, Subcommand)]
enum ConfigCommand {
    /// Initialize a config file with defaults.
    Init(InitArgs),
    /// Launch interactive wizard to edit config.
    Wizard(WizardArgs),
    /// Print effective config (loaded from file or defaults).
    Show(ShowArgs),
    /// Print the default user config path.
    Path,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
enum ConfigFormat {
    Toml,
    Json,
}

impl ConfigFormat {
    fn extension(self) -> &'static str {
        match self {
            Self::Toml => "toml",
            Self::Json => "json",
        }
    }
}

#[derive(Debug, Clone, Copy, ValueEnum)]
enum DocumentFormat {
    Yaml,
    Json,
    Toml,
}

impl DocumentFormat {
    fn extension(self) -> &'static str {
        match self {
            Self::Yaml => "yaml",
            Self::Json => "json",
            Self::Toml => "toml",
        }
    }
}

#[derive(Debug, Clone, Copy, ValueEnum)]
enum WizardLevel {
    Basic,
    Advanced,
}

#[derive(Debug, Args)]
struct InitArgs {
    /// Optional config path. Default: ~/rusdox/config.toml. Use `--path ./rusdox.toml` for a project override.
    #[arg(long)]
    path: Option<PathBuf>,
    /// Output format.
    #[arg(long, value_enum, default_value_t = ConfigFormat::Toml)]
    format: ConfigFormat,
    /// Write commented full template for TOML format.
    #[arg(long)]
    template: bool,
    /// Overwrite file if it already exists.
    #[arg(long)]
    force: bool,
}

#[derive(Debug, Args)]
struct WizardArgs {
    /// Optional config path. Default: ~/rusdox/config.toml. Use `--path ./rusdox.toml` for a project override.
    #[arg(long)]
    path: Option<PathBuf>,
    /// Wizard depth.
    #[arg(long, value_enum, default_value_t = WizardLevel::Basic)]
    level: WizardLevel,
    /// Save as TOML or JSON.
    #[arg(long, value_enum, default_value_t = ConfigFormat::Toml)]
    format: ConfigFormat,
}

#[derive(Debug, Args)]
struct ShowArgs {
    /// Optional config path. Defaults to the effective runtime config: `./rusdox.toml`, then ~/rusdox/config.toml, then defaults.
    #[arg(long)]
    path: Option<PathBuf>,
    /// Print as TOML or JSON.
    #[arg(long, value_enum, default_value_t = ConfigFormat::Toml)]
    format: ConfigFormat,
}

#[derive(Debug, Args)]
struct InitDocArgs {
    /// Spec path to create, for example `mydoc.yaml`.
    path: PathBuf,
    /// Starter document format.
    #[arg(long, value_enum, default_value_t = DocumentFormat::Yaml)]
    format: DocumentFormat,
    /// Overwrite if the spec already exists.
    #[arg(long)]
    force: bool,
}

#[derive(Debug, Args)]
struct InitScriptArgs {
    /// Script path to create, for example `mydoc.rs`.
    path: PathBuf,
    /// Overwrite if script already exists.
    #[arg(long)]
    force: bool,
}

fn main() {
    if let Err(error) = run() {
        eprintln!("error: {error}");
        std::process::exit(1);
    }
}

fn run() -> Result<()> {
    let cli = Cli::parse();
    if let Some(command) = cli.command {
        return match command {
            Commands::Config { command } => run_config_command(command),
            Commands::InitDoc(args) => init_doc(args),
            Commands::InitScript(args) => init_script(args),
        };
    }

    let input = cli.input.ok_or_else(|| {
        DocxError::Parse(
            "missing input path (usage: rusdox mydoc.yaml or rusdox mydoc.rs)".to_string(),
        )
    })?;
    run_input(
        input,
        cli.output,
        cli.config,
        cli.docx_only,
        cli.with_pdf,
        cli.release,
    )
}

fn init_doc(args: InitDocArgs) -> Result<()> {
    let path = resolve_doc_path(args.path, Some(args.format));
    if path.exists() && !args.force {
        return Err(DocxError::Parse(format!(
            "document spec already exists at {} (use --force to overwrite)",
            path.display()
        )));
    }

    match args.format {
        DocumentFormat::Yaml => DocumentSpec::write_yaml_template(&path)?,
        DocumentFormat::Json | DocumentFormat::Toml => {
            starter_document_spec().save_to_path(&path)?
        }
    }

    println!("{}", path.display());
    Ok(())
}

fn init_script(args: InitScriptArgs) -> Result<()> {
    let path = args.path;
    if path.exists() && !args.force {
        return Err(DocxError::Parse(format!(
            "script already exists at {} (use --force to overwrite)",
            path.display()
        )));
    }

    if let Some(parent) = path.parent() {
        if !parent.as_os_str().is_empty() {
            fs::create_dir_all(parent)?;
        }
    }
    fs::write(&path, default_script_template())?;
    println!("{}", path.display());
    Ok(())
}

fn run_config_command(command: ConfigCommand) -> Result<()> {
    match command {
        ConfigCommand::Init(args) => init_config(args),
        ConfigCommand::Wizard(args) => run_wizard(args),
        ConfigCommand::Show(args) => show_config(args),
        ConfigCommand::Path => {
            let path = default_path_with_fallback();
            println!("{}", path.display());
            Ok(())
        }
    }
}

fn init_config(args: InitArgs) -> Result<()> {
    let path = resolve_path(args.path, Some(args.format));
    if path.exists() && !args.force {
        return Err(DocxError::Parse(format!(
            "config already exists at {} (use --force to overwrite)",
            path.display()
        )));
    }

    if args.template && matches!(args.format, ConfigFormat::Toml) {
        RusdoxConfig::write_toml_template(&path)?;
    } else {
        RusdoxConfig::default().save_to_path(&path)?;
    }

    println!("{}", path.display());
    Ok(())
}

fn run_wizard(args: WizardArgs) -> Result<()> {
    let path = resolve_path(args.path, Some(args.format));
    let mut config = RusdoxConfig::load_from_path_or_default(&path)?;

    match args.level {
        WizardLevel::Basic => run_basic_wizard(&mut config)?,
        WizardLevel::Advanced => run_advanced_wizard(&mut config)?,
    }

    config.save_to_path(&path)?;
    println!("{}", path.display());
    Ok(())
}

fn show_config(args: ShowArgs) -> Result<()> {
    let config = if let Some(path) = args.path.as_ref() {
        let path = resolve_path(Some(path.clone()), None);
        RusdoxConfig::load_from_path_or_default(&path)?
    } else {
        load_runtime_config(None)?
    };
    match args.format {
        ConfigFormat::Toml => println!("{}", config.to_toml_pretty()?),
        ConfigFormat::Json => println!("{}", config.to_json_pretty()?),
    }
    Ok(())
}

fn run_input(
    input: PathBuf,
    output: Option<PathBuf>,
    config_path: Option<PathBuf>,
    docx_only: bool,
    with_pdf: bool,
    release: bool,
) -> Result<()> {
    if !input.exists() {
        return Err(DocxError::Parse(format!(
            "input not found: {}",
            input.display()
        )));
    }

    if input.is_dir() {
        if output.is_some() {
            return Err(DocxError::Parse(
                "--output is only supported for a single file input".to_string(),
            ));
        }
        return run_spec_dir(input, config_path, docx_only, with_pdf);
    }

    match input
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase()
        .as_str()
    {
        "rs" => run_script(input, output, config_path, docx_only, with_pdf, release),
        "yaml" | "yml" | "json" | "toml" => {
            run_spec_file(input, output, config_path, docx_only, with_pdf)
        }
        _ => Err(DocxError::Parse(format!(
            "unsupported input type: {} (expected .yaml, .yml, .json, .toml, .rs, or a directory)",
            input.display()
        ))),
    }
}

fn run_spec_file(
    spec_path: PathBuf,
    output: Option<PathBuf>,
    config_path: Option<PathBuf>,
    docx_only: bool,
    with_pdf: bool,
) -> Result<()> {
    let spec = DocumentSpec::load_from_path(&spec_path)?;
    let config = runtime_config(config_path.as_deref(), docx_only, with_pdf)?;
    let studio = Studio::new(config);

    if let Some(output_path) = output {
        let output_path = to_absolute_path(&output_path)?;
        if let Some(parent) = output_path.parent() {
            fs::create_dir_all(parent)?;
        }
        let document = studio.compose(&spec);
        studio.save_with_pdf(&document, output_path)
    } else {
        let output_name = spec
            .output_name
            .clone()
            .unwrap_or_else(|| default_output_name_for_spec(&spec_path));
        studio.save_spec_named(&spec, &output_name).map(|_| ())
    }
}

fn run_spec_dir(
    dir: PathBuf,
    config_path: Option<PathBuf>,
    docx_only: bool,
    with_pdf: bool,
) -> Result<()> {
    let mut entries = fs::read_dir(&dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| path.is_file() && is_spec_path(path))
        .collect::<Vec<_>>();
    entries.sort();

    if entries.is_empty() {
        return Err(DocxError::Parse(format!(
            "no document spec files found in {}",
            dir.display()
        )));
    }

    let config = runtime_config(config_path.as_deref(), docx_only, with_pdf)?;
    let studio = Studio::new(config);

    for spec_path in entries {
        let spec = DocumentSpec::load_from_path(&spec_path)?;
        let output_name = spec
            .output_name
            .clone()
            .unwrap_or_else(|| default_output_name_for_spec(&spec_path));
        studio.save_spec_named(&spec, &output_name)?;
    }

    Ok(())
}

fn run_script(
    script: PathBuf,
    output: Option<PathBuf>,
    config_path: Option<PathBuf>,
    docx_only: bool,
    with_pdf: bool,
    release: bool,
) -> Result<()> {
    if !script.exists() {
        return Err(DocxError::Parse(format!(
            "script not found: {}",
            script.display()
        )));
    }
    if script.extension().and_then(|ext| ext.to_str()) != Some("rs") {
        return Err(DocxError::Parse(format!(
            "script must be a .rs file: {}",
            script.display()
        )));
    }

    let script_source = fs::read_to_string(&script)?;
    if !script_source.contains("build_document") {
        return Err(DocxError::Parse(
            "script must define `build_document(&Studio) -> rusdox::Result<Document>`".to_string(),
        ));
    }

    let script_path = fs::canonicalize(&script)?;
    let output_path = output.unwrap_or_else(|| default_output_for_script(&script_path));
    let output_path = to_absolute_path(&output_path)?;
    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)?;
    }

    let mut config = runtime_config(config_path.as_deref(), docx_only, with_pdf)?;

    let output_dir = output_path
        .parent()
        .unwrap_or_else(|| Path::new("."))
        .to_string_lossy()
        .to_string();
    config.output.docx_dir = output_dir.clone();
    config.output.pdf_dir = output_dir;

    let temp = tempdir()?;
    let runner_dir = temp.path();
    let manifest_path = runner_dir.join("Cargo.toml");
    let src_dir = runner_dir.join("src");
    fs::create_dir_all(&src_dir)?;

    let runner_config_path = runner_dir.join("rusdox-runtime-config.toml");
    config.save_to_path(&runner_config_path)?;

    fs::write(&manifest_path, build_runner_manifest())?;
    fs::write(
        src_dir.join("main.rs"),
        build_runner_source(&script_path, &output_path, &runner_config_path),
    )?;

    let mut command = Command::new("cargo");
    command.arg("run");
    if release {
        command.arg("--release");
    }
    command.arg("--quiet");
    command.arg("--manifest-path");
    command.arg(&manifest_path);

    let status = command.status()?;
    if !status.success() {
        return Err(DocxError::Parse(format!(
            "script execution failed with status {status}"
        )));
    }
    Ok(())
}

fn default_output_for_script(script_path: &Path) -> PathBuf {
    let mut path = script_path.to_path_buf();
    path.set_extension("docx");
    path
}

fn default_output_name_for_spec(spec_path: &Path) -> String {
    spec_path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .unwrap_or("document")
        .replace('_', "-")
}

fn is_spec_path(path: &Path) -> bool {
    matches!(
        path.extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase()
            .as_str(),
        "yaml" | "yml" | "json" | "toml"
    )
}

fn runtime_config(path: Option<&Path>, docx_only: bool, with_pdf: bool) -> Result<RusdoxConfig> {
    let mut config = load_runtime_config(path)?;
    if docx_only {
        config.output.emit_pdf_preview = false;
    }
    if with_pdf {
        config.output.emit_pdf_preview = true;
    }
    Ok(config)
}

fn load_runtime_config(path: Option<&Path>) -> Result<RusdoxConfig> {
    if let Some(path) = path {
        return RusdoxConfig::load_from_path_or_default(path);
    }

    RusdoxConfig::load_local_or_user_default(DEFAULT_CONFIG_FILE)
}

fn to_absolute_path(path: &Path) -> Result<PathBuf> {
    if path.is_absolute() {
        Ok(path.to_path_buf())
    } else {
        Ok(std::env::current_dir()?.join(path))
    }
}

fn build_runner_manifest() -> String {
    format!(
        r#"[package]
name = "rusdox-script-runner"
version = "0.1.0"
edition = "2021"

[dependencies]
rusdox = {}
"#,
        runner_dependency_spec()
    )
}

fn runner_dependency_spec() -> String {
    let local_path = Path::new(env!("CARGO_MANIFEST_DIR"));
    if local_path.join("Cargo.toml").exists() {
        format!(
            "{{ path = \"{}\" }}",
            escape_toml(local_path.to_string_lossy().as_ref())
        )
    } else {
        format!("\"{}\"", env!("CARGO_PKG_VERSION"))
    }
}

fn escape_toml(value: &str) -> String {
    value.replace('\\', "\\\\").replace('\"', "\\\"")
}

fn build_runner_source(script_path: &Path, output_path: &Path, config_path: &Path) -> String {
    let script_literal = escape_rust_string(script_path.to_string_lossy().as_ref());
    let output_literal = escape_rust_string(output_path.to_string_lossy().as_ref());
    let config_literal = escape_rust_string(config_path.to_string_lossy().as_ref());
    format!(
        r#"use std::path::PathBuf;

use rusdox::config::RusdoxConfig;
use rusdox::studio::Studio;

mod user_script {{
    include!("{script_literal}");
}}

fn main() -> rusdox::Result<()> {{
    let output = PathBuf::from("{output_literal}");
    let config_path = PathBuf::from("{config_literal}");
    let config = RusdoxConfig::load_from_path(config_path)?;
    let studio = Studio::new(config);
    let document = user_script::build_document(&studio)?;
    studio.save_with_pdf(&document, output)
}}
"#
    )
}

fn escape_rust_string(value: &str) -> String {
    value.replace('\\', "\\\\").replace('\"', "\\\"")
}

fn default_script_template() -> &'static str {
    r#"use rusdox::{Document, Paragraph, Run};
use rusdox::studio::Studio;

/// Build and return a document.
///
/// Run with:
///   rusdox mydoc.rs
///   rusdox mydoc.rs --docx-only
pub fn build_document(studio: &Studio) -> rusdox::Result<Document> {
    let mut doc = Document::new();
    doc.push_paragraph(studio.title("My RusDox Document"));
    doc.push_paragraph(studio.subtitle("Generated from a single .rs file"));
    doc.push_paragraph(studio.section("Summary"));
    doc.push_paragraph(studio.body("Edit this file and rerun `rusdox mydoc.rs`."));
    doc.push_paragraph(
        Paragraph::new()
            .add_run(Run::from_text("You can use full Rust + RusDox APIs. ").bold())
            .add_run(Run::from_text("Tables, styles, and rich layouts are supported.")),
    );
    Ok(doc)
}
"#
}

fn resolve_path(path: Option<PathBuf>, format: Option<ConfigFormat>) -> PathBuf {
    let mut resolved = path.unwrap_or_else(default_path_with_fallback);
    if let Some(format) = format {
        resolved.set_extension(format.extension());
    }
    resolved
}

fn resolve_doc_path(path: PathBuf, format: Option<DocumentFormat>) -> PathBuf {
    let mut resolved = path;
    if let Some(format) = format {
        resolved.set_extension(format.extension());
    }
    resolved
}

fn default_path_with_fallback() -> PathBuf {
    default_user_config_path().unwrap_or_else(|| PathBuf::from("rusdox.toml"))
}

fn starter_document_spec() -> DocumentSpec {
    DocumentSpec {
        output_name: Some("my-document".to_string()),
        blocks: vec![
            rusdox::spec::title("My Document"),
            rusdox::spec::subtitle("Written as data, rendered by Rust"),
            rusdox::spec::section("Summary"),
            rusdox::spec::body("Replace this with your real content."),
            rusdox::spec::bullets([
                "Keep content in order.",
                "Let config handle styling.",
                "Render to DOCX and PDF with one command.",
            ]),
        ],
    }
}

fn run_basic_wizard(config: &mut RusdoxConfig) -> Result<()> {
    config.profile_name = prompt_string("Profile name", &config.profile_name)?;
    config.output.docx_dir = prompt_string("DOCX output directory", &config.output.docx_dir)?;
    config.output.emit_pdf_preview =
        prompt_bool("Generate PDF previews too", config.output.emit_pdf_preview)?;
    if config.output.emit_pdf_preview {
        config.output.pdf_dir = prompt_string("PDF output directory", &config.output.pdf_dir)?;
    }

    config.typography.font_family =
        prompt_string("Default font family", &config.typography.font_family)?;
    config.typography.title_size_pt =
        prompt_f32("Title size (pt)", config.typography.title_size_pt)?;
    config.typography.body_size_pt = prompt_f32("Body size (pt)", config.typography.body_size_pt)?;

    config.colors.ink = prompt_color("Primary text color (hex)", &config.colors.ink)?;
    config.colors.accent = prompt_color("Accent color (hex)", &config.colors.accent)?;
    config.spacing.body_after_twips = prompt_u32(
        "Body paragraph spacing after (twips)",
        config.spacing.body_after_twips,
    )?;

    Ok(())
}

fn run_advanced_wizard(config: &mut RusdoxConfig) -> Result<()> {
    if prompt_bool("Run quick basic setup first", true)? {
        run_basic_wizard(config)?;
    }

    loop {
        let menu = [
            "Output",
            "Typography",
            "Spacing",
            "Colors",
            "Tables",
            "PDF renderer",
            "Finish",
        ];
        let choice = Select::new()
            .with_prompt("Advanced settings section")
            .items(menu)
            .default(0)
            .interact()
            .map_err(dialog_err)?;

        match choice {
            0 => edit_output(config)?,
            1 => edit_typography(config)?,
            2 => edit_spacing(config)?,
            3 => edit_colors(config)?,
            4 => edit_table(config)?,
            5 => edit_pdf(config)?,
            _ => break,
        }
    }

    Ok(())
}

fn edit_output(config: &mut RusdoxConfig) -> Result<()> {
    config.output.docx_dir = prompt_string("DOCX output directory", &config.output.docx_dir)?;
    config.output.emit_pdf_preview =
        prompt_bool("Generate PDF previews too", config.output.emit_pdf_preview)?;
    config.output.pdf_dir = prompt_string("PDF output directory", &config.output.pdf_dir)?;
    Ok(())
}

fn edit_typography(config: &mut RusdoxConfig) -> Result<()> {
    config.typography.font_family =
        prompt_string("Default font family", &config.typography.font_family)?;
    config.typography.cover_title_size_pt = prompt_f32(
        "Cover title size (pt)",
        config.typography.cover_title_size_pt,
    )?;
    config.typography.title_size_pt =
        prompt_f32("Title size (pt)", config.typography.title_size_pt)?;
    config.typography.subtitle_size_pt =
        prompt_f32("Subtitle size (pt)", config.typography.subtitle_size_pt)?;
    config.typography.hero_size_pt = prompt_f32("Hero size (pt)", config.typography.hero_size_pt)?;
    config.typography.page_heading_size_pt = prompt_f32(
        "Page heading size (pt)",
        config.typography.page_heading_size_pt,
    )?;
    config.typography.section_size_pt = prompt_f32(
        "Section heading size (pt)",
        config.typography.section_size_pt,
    )?;
    config.typography.body_size_pt = prompt_f32("Body size (pt)", config.typography.body_size_pt)?;
    config.typography.tagline_size_pt =
        prompt_f32("Tagline size (pt)", config.typography.tagline_size_pt)?;
    config.typography.note_size_pt = prompt_f32("Note size (pt)", config.typography.note_size_pt)?;
    config.typography.table_size_pt =
        prompt_f32("Table text size (pt)", config.typography.table_size_pt)?;
    config.typography.metric_label_size_pt = prompt_f32(
        "Metric label size (pt)",
        config.typography.metric_label_size_pt,
    )?;
    config.typography.metric_value_size_pt = prompt_f32(
        "Metric value size (pt)",
        config.typography.metric_value_size_pt,
    )?;
    Ok(())
}

fn edit_spacing(config: &mut RusdoxConfig) -> Result<()> {
    config.spacing.cover_title_before_twips = prompt_u32(
        "cover_title_before_twips",
        config.spacing.cover_title_before_twips,
    )?;
    config.spacing.cover_title_after_twips = prompt_u32(
        "cover_title_after_twips",
        config.spacing.cover_title_after_twips,
    )?;
    config.spacing.title_before_twips =
        prompt_u32("title_before_twips", config.spacing.title_before_twips)?;
    config.spacing.title_after_twips =
        prompt_u32("title_after_twips", config.spacing.title_after_twips)?;
    config.spacing.subtitle_after_twips =
        prompt_u32("subtitle_after_twips", config.spacing.subtitle_after_twips)?;
    config.spacing.hero_after_twips =
        prompt_u32("hero_after_twips", config.spacing.hero_after_twips)?;
    config.spacing.page_heading_after_twips = prompt_u32(
        "page_heading_after_twips",
        config.spacing.page_heading_after_twips,
    )?;
    config.spacing.section_before_twips =
        prompt_u32("section_before_twips", config.spacing.section_before_twips)?;
    config.spacing.section_after_twips =
        prompt_u32("section_after_twips", config.spacing.section_after_twips)?;
    config.spacing.body_after_twips =
        prompt_u32("body_after_twips", config.spacing.body_after_twips)?;
    config.spacing.bullet_after_twips =
        prompt_u32("bullet_after_twips", config.spacing.bullet_after_twips)?;
    config.spacing.label_value_after_twips = prompt_u32(
        "label_value_after_twips",
        config.spacing.label_value_after_twips,
    )?;
    config.spacing.tagline_after_twips =
        prompt_u32("tagline_after_twips", config.spacing.tagline_after_twips)?;
    config.spacing.spacer_after_twips =
        prompt_u32("spacer_after_twips", config.spacing.spacer_after_twips)?;
    config.spacing.note_after_twips =
        prompt_u32("note_after_twips", config.spacing.note_after_twips)?;
    config.spacing.metric_label_before_twips = prompt_u32(
        "metric_label_before_twips",
        config.spacing.metric_label_before_twips,
    )?;
    config.spacing.metric_label_after_twips = prompt_u32(
        "metric_label_after_twips",
        config.spacing.metric_label_after_twips,
    )?;
    config.spacing.metric_value_after_twips = prompt_u32(
        "metric_value_after_twips",
        config.spacing.metric_value_after_twips,
    )?;
    config.spacing.table_header_before_twips = prompt_u32(
        "table_header_before_twips",
        config.spacing.table_header_before_twips,
    )?;
    config.spacing.table_header_after_twips = prompt_u32(
        "table_header_after_twips",
        config.spacing.table_header_after_twips,
    )?;
    config.spacing.table_data_before_twips = prompt_u32(
        "table_data_before_twips",
        config.spacing.table_data_before_twips,
    )?;
    config.spacing.table_data_after_twips = prompt_u32(
        "table_data_after_twips",
        config.spacing.table_data_after_twips,
    )?;
    config.spacing.table_status_before_twips = prompt_u32(
        "table_status_before_twips",
        config.spacing.table_status_before_twips,
    )?;
    config.spacing.table_status_after_twips = prompt_u32(
        "table_status_after_twips",
        config.spacing.table_status_after_twips,
    )?;
    Ok(())
}

fn edit_colors(config: &mut RusdoxConfig) -> Result<()> {
    config.colors.ink = prompt_color("ink", &config.colors.ink)?;
    config.colors.slate = prompt_color("slate", &config.colors.slate)?;
    config.colors.muted = prompt_color("muted", &config.colors.muted)?;
    config.colors.accent = prompt_color("accent", &config.colors.accent)?;
    config.colors.gold = prompt_color("gold", &config.colors.gold)?;
    config.colors.red = prompt_color("red", &config.colors.red)?;
    config.colors.green = prompt_color("green", &config.colors.green)?;
    config.colors.soft = prompt_color("soft", &config.colors.soft)?;
    config.colors.pale = prompt_color("pale", &config.colors.pale)?;
    config.colors.mint = prompt_color("mint", &config.colors.mint)?;
    config.colors.amber = prompt_color("amber", &config.colors.amber)?;
    config.colors.rose = prompt_color("rose", &config.colors.rose)?;
    config.colors.table_border = prompt_color("table_border", &config.colors.table_border)?;
    Ok(())
}

fn edit_table(config: &mut RusdoxConfig) -> Result<()> {
    config.table.default_width_twips =
        prompt_u32("default_width_twips", config.table.default_width_twips)?;
    config.table.metric_cell_width_twips = prompt_u32(
        "metric_cell_width_twips",
        config.table.metric_cell_width_twips,
    )?;
    config.table.grid_border_size_eighth_pt = prompt_u32(
        "grid_border_size_eighth_pt",
        config.table.grid_border_size_eighth_pt,
    )?;
    config.table.card_border_size_eighth_pt = prompt_u32(
        "card_border_size_eighth_pt",
        config.table.card_border_size_eighth_pt,
    )?;
    config.table.pdf_cell_padding_x_pt =
        prompt_f32("pdf_cell_padding_x_pt", config.table.pdf_cell_padding_x_pt)?;
    config.table.pdf_cell_padding_y_pt =
        prompt_f32("pdf_cell_padding_y_pt", config.table.pdf_cell_padding_y_pt)?;
    config.table.pdf_after_spacing_pt =
        prompt_f32("pdf_after_spacing_pt", config.table.pdf_after_spacing_pt)?;
    config.table.pdf_grid_stroke_width_pt = prompt_f32(
        "pdf_grid_stroke_width_pt",
        config.table.pdf_grid_stroke_width_pt,
    )?;
    Ok(())
}

fn edit_pdf(config: &mut RusdoxConfig) -> Result<()> {
    config.pdf.page_width_pt = prompt_f32("page_width_pt", config.pdf.page_width_pt)?;
    config.pdf.page_height_pt = prompt_f32("page_height_pt", config.pdf.page_height_pt)?;
    config.pdf.margin_x_pt = prompt_f32("margin_x_pt", config.pdf.margin_x_pt)?;
    config.pdf.margin_top_pt = prompt_f32("margin_top_pt", config.pdf.margin_top_pt)?;
    config.pdf.margin_bottom_pt = prompt_f32("margin_bottom_pt", config.pdf.margin_bottom_pt)?;
    config.pdf.default_text_size_pt =
        prompt_f32("default_text_size_pt", config.pdf.default_text_size_pt)?;
    config.pdf.default_line_height_pt =
        prompt_f32("default_line_height_pt", config.pdf.default_line_height_pt)?;
    config.pdf.line_height_multiplier =
        prompt_f32("line_height_multiplier", config.pdf.line_height_multiplier)?;
    config.pdf.baseline_factor = prompt_f32("baseline_factor", config.pdf.baseline_factor)?;
    config.pdf.text_width_bias_regular = prompt_f32(
        "text_width_bias_regular",
        config.pdf.text_width_bias_regular,
    )?;
    config.pdf.text_width_bias_bold =
        prompt_f32("text_width_bias_bold", config.pdf.text_width_bias_bold)?;
    Ok(())
}

fn prompt_string(prompt: &str, default: &str) -> Result<String> {
    Input::new()
        .with_prompt(prompt)
        .default(default.to_string())
        .interact_text()
        .map_err(dialog_err)
}

fn prompt_u32(prompt: &str, default: u32) -> Result<u32> {
    Input::new()
        .with_prompt(prompt)
        .default(default)
        .interact_text()
        .map_err(dialog_err)
}

fn prompt_f32(prompt: &str, default: f32) -> Result<f32> {
    Input::new()
        .with_prompt(prompt)
        .default(default)
        .interact_text()
        .map_err(dialog_err)
}

fn prompt_bool(prompt: &str, default: bool) -> Result<bool> {
    Confirm::new()
        .with_prompt(prompt)
        .default(default)
        .interact()
        .map_err(dialog_err)
}

fn prompt_color(prompt: &str, default: &str) -> Result<String> {
    let candidate = prompt_string(prompt, default)?;
    normalize_color_hex(&candidate)
}

fn normalize_color_hex(raw: &str) -> Result<String> {
    let normalized = raw.trim().trim_start_matches('#').to_ascii_uppercase();
    if normalized.len() != 6 || !normalized.chars().all(|ch| ch.is_ascii_hexdigit()) {
        return Err(DocxError::Parse(format!(
            "invalid color '{raw}', expected six hex digits"
        )));
    }
    Ok(normalized)
}

fn dialog_err(error: dialoguer::Error) -> DocxError {
    DocxError::Parse(format!("interactive prompt failed: {error}"))
}

#[cfg(test)]
mod tests {
    use std::path::Path;

    use super::{
        build_runner_manifest, build_runner_source, default_script_template, escape_rust_string,
        normalize_color_hex, resolve_path, ConfigFormat,
    };

    #[test]
    fn normalize_color_hex_accepts_hash_and_lowercase() {
        assert_eq!(
            normalize_color_hex("#a1b2c3").expect("valid color"),
            "A1B2C3"
        );
        assert_eq!(
            normalize_color_hex("ff00ff").expect("valid color"),
            "FF00FF"
        );
    }

    #[test]
    fn normalize_color_hex_rejects_invalid_values() {
        assert!(normalize_color_hex("12345").is_err());
        assert!(normalize_color_hex("1234567").is_err());
        assert!(normalize_color_hex("GG0000").is_err());
    }

    #[test]
    fn resolve_path_applies_format_extension() {
        let base_dir = std::env::temp_dir().join("rusdox-cli-tests");

        let path = resolve_path(
            Some(base_dir.join("config.anything")),
            Some(ConfigFormat::Toml),
        );
        assert_eq!(
            path.file_name().and_then(|name| name.to_str()),
            Some("config.toml")
        );
        assert_eq!(path.parent(), Some(base_dir.as_path()));

        let path = resolve_path(Some(base_dir.join("config")), Some(ConfigFormat::Json));
        assert_eq!(
            path.file_name().and_then(|name| name.to_str()),
            Some("config.json")
        );
        assert_eq!(path.parent(), Some(base_dir.as_path()));
    }

    #[test]
    fn default_script_template_exposes_expected_entry_point() {
        let template = default_script_template();
        assert!(template.contains("pub fn build_document("));
        assert!(template.contains("rusdox mydoc.rs"));
        assert!(template.contains("use rusdox::studio::Studio;"));
    }

    #[test]
    fn runner_manifest_contains_dependency_section() {
        let manifest = build_runner_manifest();
        assert!(manifest.contains("[package]"));
        assert!(manifest.contains("[dependencies]"));
        assert!(manifest.contains("rusdox = "));
    }

    #[test]
    fn runner_source_embeds_paths_and_calls_build_document() {
        let script = Path::new(r#"path\with\"quote\script.rs"#);
        let output = Path::new(r#"path\with\out.docx"#);
        let config = Path::new(r#"path\with\rusdox.toml"#);
        let source = build_runner_source(script, output, config);
        let expected_include = format!(
            "include!(\"{}\")",
            escape_rust_string(script.to_string_lossy().as_ref())
        );
        assert!(source.contains(&expected_include));
        assert!(source.contains("let document = user_script::build_document(&studio)?;"));
        assert!(source.contains("studio.save_with_pdf(&document, output)"));
    }
}
