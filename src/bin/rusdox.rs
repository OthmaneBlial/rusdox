use std::collections::BTreeSet;
use std::fs;
use std::hash::{DefaultHasher, Hash, Hasher};
use std::io::{BufRead, Read, Write};
use std::net::{IpAddr, Ipv4Addr, SocketAddr, TcpListener, TcpStream};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::{Arc, RwLock};
use std::thread;
use std::time::{Duration, Instant};

use clap::{Args, Parser, Subcommand, ValueEnum};
use dialoguer::{Confirm, Input, Select};
use ed25519_dalek::{Signature, Verifier, VerifyingKey};
use rusdox::config::{default_user_config_path, RusdoxConfig};
use rusdox::parity::{compare_visual_pages, ArtifactEvidence, DocumentProjection, ParityReport};
use rusdox::protocol::{
    execute_protocol_request, protocol_failure, ProtocolRequest, ProtocolResponse,
    MAX_PROTOCOL_REQUEST_BYTES, PROTOCOL_VERSION,
};
use rusdox::spec::DocumentSpec;
use rusdox::spec::SPEC_VERSION;
use rusdox::studio::{OutputStats, Studio, DEFAULT_CONFIG_FILE};
use rusdox::{
    atomic_write_file, attach_source_spans, document_spec_schema_pretty, validate_config,
    validate_spec, Document, DocxError, DocxTemplate, Result, ValidationIssue, ValidationReport,
    ValidationSeverity,
};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use tempfile::tempdir;

const DEFAULT_TEMPLATE_REGISTRY: &str = "https://othmaneblial.github.io/rusdox/registry/index.json";
const DEFAULT_TEMPLATE_REGISTRY_PUBLIC_KEY: &str =
    "6a765396357492a6aa4239c8fc042e0471259528f5b25ab742b964b735662353";
const MAX_REGISTRY_BYTES: u64 = 2 * 1024 * 1024;
const MAX_TEMPLATE_DOWNLOAD_BYTES: u64 = 64 * 1024 * 1024;

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
    /// Print or write the generated document-spec JSON Schema.
    Schema(SchemaArgs),
    /// Upgrade a legacy document spec to the current version.
    Migrate(MigrateArgs),
    /// Validate a document spec or spec directory without rendering output.
    Validate(ValidateArgs),
    /// Rebuild a document spec automatically when the spec or config changes.
    Watch(WatchArgs),
    /// Rebuild with a local PDF/status dashboard and structured feedback.
    Dev(DevArgs),
    /// Run the stable local JSON protocol over stdin/stdout or loopback HTTP.
    Serve(ServeArgs),
    /// Reproducibly measure one validation or rendering pipeline.
    Bench(BenchArgs),
    /// Inspect or render a Word-native DOCX template from JSON data.
    Template {
        #[command(subcommand)]
        command: TemplateCommand,
    },
    /// Generate DOCX, native PDF, deterministic page snapshots, and parity evidence.
    Verify(VerifyArgs),
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

#[derive(Debug, Subcommand)]
enum TemplateCommand {
    /// List placeholders and structural diagnostics without rendering.
    Inspect(TemplateInspectArgs),
    /// Generate DOCX, native PDF, page snapshots, and parity evidence.
    Render(TemplateRenderArgs),
    /// Render and require every enabled DOCX/PDF parity check to pass.
    Verify(TemplateRenderArgs),
    /// List verified entries from the signed curated registry.
    List(TemplateListArgs),
    /// Search verified registry metadata by text or category.
    Search(TemplateSearchArgs),
    /// Download and hash-verify one curated Word template and its sample data.
    Add(TemplateAddArgs),
    /// Refresh one or all installed templates from the signed registry.
    Update(TemplateUpdateArgs),
}

#[derive(Debug, Args)]
struct TemplateListArgs {
    /// Signed registry index URL or local path.
    #[arg(long, default_value = DEFAULT_TEMPLATE_REGISTRY)]
    registry: String,
    /// Trusted Ed25519 public key as 64 lowercase hexadecimal characters.
    #[arg(long)]
    public_key: Option<String>,
    /// Command output format.
    #[arg(long, value_enum, default_value_t = ReportFormat::Text)]
    format: ReportFormat,
}

#[derive(Debug, Args)]
struct TemplateSearchArgs {
    /// Case-insensitive title, description, category, tag, author, or ID query.
    query: String,
    /// Signed registry index URL or local path.
    #[arg(long, default_value = DEFAULT_TEMPLATE_REGISTRY)]
    registry: String,
    /// Trusted Ed25519 public key as 64 lowercase hexadecimal characters.
    #[arg(long)]
    public_key: Option<String>,
    /// Command output format.
    #[arg(long, value_enum, default_value_t = ReportFormat::Text)]
    format: ReportFormat,
}

#[derive(Debug, Args)]
struct TemplateAddArgs {
    /// Registry template identifier.
    id: String,
    /// Signed registry index URL or local path.
    #[arg(long, default_value = DEFAULT_TEMPLATE_REGISTRY)]
    registry: String,
    /// Trusted Ed25519 public key as 64 lowercase hexadecimal characters.
    #[arg(long)]
    public_key: Option<String>,
    /// Installation root. Defaults to the platform user data directory.
    #[arg(long)]
    install_root: Option<PathBuf>,
    /// Replace an installed template even when the version is unchanged.
    #[arg(long)]
    force: bool,
    /// Command output format.
    #[arg(long, value_enum, default_value_t = ReportFormat::Text)]
    format: ReportFormat,
}

#[derive(Debug, Args)]
struct TemplateUpdateArgs {
    /// Installed template identifier. Omit only with --all.
    id: Option<String>,
    /// Update every installed template that appears in the registry.
    #[arg(long, conflicts_with = "id", required_unless_present = "id")]
    all: bool,
    /// Signed registry index URL or local path.
    #[arg(long, default_value = DEFAULT_TEMPLATE_REGISTRY)]
    registry: String,
    /// Trusted Ed25519 public key as 64 lowercase hexadecimal characters.
    #[arg(long)]
    public_key: Option<String>,
    /// Installation root. Defaults to the platform user data directory.
    #[arg(long)]
    install_root: Option<PathBuf>,
    /// Command output format.
    #[arg(long, value_enum, default_value_t = ReportFormat::Text)]
    format: ReportFormat,
}

#[derive(Debug, Args)]
struct TemplateInspectArgs {
    /// Word-native `.docx` template.
    template: PathBuf,
    /// Inspection report format.
    #[arg(long, value_enum, default_value_t = ReportFormat::Text)]
    format: ReportFormat,
}

#[derive(Debug, Args)]
struct TemplateRenderArgs {
    /// Word-native `.docx` template.
    template: PathBuf,
    /// JSON object containing values and optional `$partials` strings.
    data: PathBuf,
    /// Root for generated/, rendered/, and reports/ artifacts.
    #[arg(long, default_value = ".")]
    output_root: PathBuf,
    /// Artifact stem. Defaults to the DOCX template file stem.
    #[arg(long)]
    name: Option<String>,
    /// Fail instead of replacing missing/null values with empty strings.
    #[arg(long)]
    strict: bool,
    /// Optional config path used by the native PDF renderer.
    #[arg(long)]
    config: Option<PathBuf>,
    /// Command summary format.
    #[arg(long, value_enum, default_value_t = ReportFormat::Text)]
    format: ReportFormat,
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
enum ReportFormat {
    Text,
    Json,
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

#[derive(Debug, Args)]
struct SchemaArgs {
    /// Write the schema atomically instead of printing it to stdout.
    #[arg(long)]
    output: Option<PathBuf>,
}

#[derive(Debug, Args)]
struct MigrateArgs {
    /// YAML, JSON, or TOML document spec to migrate.
    input: PathBuf,
    /// Write the migrated spec to a separate path.
    #[arg(long, conflicts_with_all = ["in_place", "check"])]
    output: Option<PathBuf>,
    /// Atomically replace the input after a successful migration and parse check.
    #[arg(long, conflicts_with_all = ["output", "check"])]
    in_place: bool,
    /// Exit non-zero when the input needs migration; never write.
    #[arg(long, conflicts_with_all = ["output", "in_place"])]
    check: bool,
    /// Allow --output to replace an existing file.
    #[arg(long, requires = "output")]
    force: bool,
}

#[derive(Debug, Args)]
struct ValidateArgs {
    /// Spec file or directory to validate.
    input: PathBuf,
    /// Optional config path used for config-aware validation.
    #[arg(long)]
    config: Option<PathBuf>,
    /// Output report format.
    #[arg(long, value_enum, default_value_t = ReportFormat::Text)]
    format: ReportFormat,
}

#[derive(Debug, Args)]
struct WatchArgs {
    /// Spec file or directory to watch.
    input: PathBuf,
    /// Optional explicit output DOCX path for a single watched spec file.
    #[arg(long)]
    output: Option<PathBuf>,
    /// Optional config path used while rebuilding.
    #[arg(long)]
    config: Option<PathBuf>,
    /// Force DOCX-only generation (disable PDF output).
    #[arg(long)]
    docx_only: bool,
    /// Force PDF generation (overrides config if disabled).
    #[arg(long, conflicts_with = "docx_only")]
    with_pdf: bool,
    /// Poll interval in milliseconds.
    #[arg(long, default_value_t = 750)]
    poll_interval_ms: u64,
    /// Wait for a burst of file writes to settle before rebuilding.
    #[arg(long, default_value_t = 150)]
    debounce_ms: u64,
    /// Stop after this many build attempts, including the initial build.
    #[arg(long)]
    max_builds: Option<u32>,
}

#[derive(Debug, Args)]
struct DevArgs {
    /// Spec file or directory to rebuild.
    input: PathBuf,
    /// Optional explicit output DOCX path for a single watched spec file.
    #[arg(long)]
    output: Option<PathBuf>,
    /// Optional config path used while rebuilding.
    #[arg(long)]
    config: Option<PathBuf>,
    /// Force DOCX-only generation (the dashboard still reports status and DOCX paths).
    #[arg(long)]
    docx_only: bool,
    /// Force PDF generation even when disabled by config.
    #[arg(long, conflicts_with = "docx_only")]
    with_pdf: bool,
    /// Poll interval in milliseconds.
    #[arg(long, default_value_t = 250)]
    poll_interval_ms: u64,
    /// Wait for a burst of file writes to settle before rebuilding.
    #[arg(long, default_value_t = 150)]
    debounce_ms: u64,
    /// Local dashboard port. Use 0 to let the operating system choose a free port.
    #[arg(long, default_value_t = 4174)]
    port: u16,
    /// Open the local dashboard in the default browser.
    #[arg(long)]
    open: bool,
    /// Emit one machine-readable JSON object per lifecycle/build event.
    #[arg(long, conflicts_with = "quiet")]
    json: bool,
    /// Suppress normal terminal output while retaining the local dashboard.
    #[arg(long, short = 'q', conflicts_with = "json")]
    quiet: bool,
    /// Stop after this many build attempts, including the initial build.
    #[arg(long)]
    max_builds: Option<u32>,
}

#[derive(Debug, Args)]
struct ServeArgs {
    /// Explicit local transport.
    #[arg(value_enum)]
    transport: ServeTransport,
    /// Fixed root beneath which render requests may write relative outputs.
    #[arg(long, default_value = ".rusdox-service")]
    output_root: PathBuf,
    /// Loopback HTTP port. Use 0 to let the operating system choose.
    #[arg(long, default_value_t = 0)]
    port: u16,
    /// Stop after a bounded number of non-empty requests (useful for jobs and tests).
    #[arg(long)]
    max_requests: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, ValueEnum)]
enum ServeTransport {
    Stdio,
    Http,
}

#[derive(Debug, Args)]
struct BenchArgs {
    /// Spec file/directory, or a DOCX file with --pipeline existing-docx.
    input: PathBuf,
    /// Optional config path used while benchmarking.
    #[arg(long)]
    config: Option<PathBuf>,
    /// Force DOCX-only generation (disable PDF output).
    #[arg(long)]
    docx_only: bool,
    /// Force PDF generation (overrides config if disabled).
    #[arg(long, conflicts_with = "docx_only")]
    with_pdf: bool,
    /// Pipeline to isolate. Without this flag, legacy DOCX/PDF config behavior is preserved.
    #[arg(long, value_enum, conflicts_with_all = ["docx_only", "with_pdf"])]
    pipeline: Option<BenchPipeline>,
    /// Number of measured iterations.
    #[arg(long, default_value_t = 3)]
    iterations: u32,
    /// Number of warmup iterations to discard before measuring.
    #[arg(long, default_value_t = 0)]
    warmup: u32,
    /// Output report format.
    #[arg(long, value_enum, default_value_t = ReportFormat::Text)]
    format: ReportFormat,
    /// Keep benchmark artifacts in the configured output folders instead of using a temporary workspace.
    #[arg(long)]
    keep_output: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, ValueEnum)]
enum BenchPipeline {
    Validation,
    Docx,
    Pdf,
    Dual,
    ExistingDocx,
}

impl BenchPipeline {
    fn as_str(self) -> &'static str {
        match self {
            Self::Validation => "validation",
            Self::Docx => "docx",
            Self::Pdf => "pdf",
            Self::Dual => "dual",
            Self::ExistingDocx => "existing-docx",
        }
    }
}

#[derive(Debug, Args)]
struct VerifyArgs {
    /// Spec file or directory to render and verify.
    input: PathBuf,
    /// Optional config path used while rendering.
    #[arg(long)]
    config: Option<PathBuf>,
    /// Root for generated/, rendered/, and reports/ artifacts.
    #[arg(long, default_value = ".")]
    output_root: PathBuf,
    /// Optional page-snapshot baseline directory. For directory inputs, use one subdirectory per spec stem.
    #[arg(long)]
    visual_baseline: Option<PathBuf>,
    /// Maximum fraction of pixels that may differ on each deterministic page snapshot.
    #[arg(long, default_value_t = 0.0)]
    visual_threshold: f64,
    /// Command summary format. HTML and JSON report files are always written.
    #[arg(long, value_enum, default_value_t = ReportFormat::Text)]
    format: ReportFormat,
}

#[derive(Debug)]
struct SpecInspection {
    spec: DocumentSpec,
    parse_duration: Duration,
    validation_duration: Duration,
    report: ValidationReport,
}

#[derive(Debug, Clone, Copy)]
struct BuildSummary {
    documents: usize,
    parse_duration: Duration,
    validation_duration: Duration,
    compose_duration: Duration,
    output_stats: OutputStats,
    total_duration: Duration,
    warning_count: usize,
}

#[derive(Debug, Clone, Copy)]
struct BenchSample {
    parse_duration: Duration,
    validation_duration: Duration,
    compose_duration: Duration,
    docx_duration: Duration,
    pdf_duration: Duration,
    total_duration: Duration,
    docx_bytes: u64,
    pdf_bytes: u64,
    existing_docx_open_duration: Duration,
    existing_docx_save_duration: Duration,
}

#[derive(Debug, Clone, Serialize)]
struct ValidationFileResult {
    path: String,
    parse_ms: f64,
    validate_ms: f64,
    issues: Vec<ValidationIssue>,
}

#[derive(Debug, Clone, Serialize)]
struct ValidationCommandResult {
    target: String,
    specs: usize,
    errors: usize,
    warnings: usize,
    config_issues: Vec<ValidationIssue>,
    files: Vec<ValidationFileResult>,
}

#[derive(Debug, Clone, Serialize)]
struct NumericSummary {
    avg: f64,
    min: f64,
    median: f64,
    max: f64,
}

#[derive(Debug, Clone, Serialize)]
struct BenchCommandResult {
    schema_version: u32,
    target: String,
    pipeline: String,
    input_sha256: String,
    input_bytes: u64,
    specs: usize,
    iterations: u32,
    warmup: u32,
    emit_pdf: bool,
    keep_output: bool,
    parse_ms: NumericSummary,
    validate_ms: NumericSummary,
    compose_ms: NumericSummary,
    docx_ms: NumericSummary,
    pdf_ms: NumericSummary,
    existing_docx_open_ms: NumericSummary,
    existing_docx_save_ms: NumericSummary,
    total_ms: NumericSummary,
    docx_bytes: NumericSummary,
    pdf_bytes: NumericSummary,
}

#[derive(Debug, Clone, Serialize)]
struct TemplateCommandResult {
    command: String,
    template: String,
    data: String,
    passed: bool,
    strict: bool,
    replacements: usize,
    expanded_blocks: usize,
    diagnostics: Vec<rusdox::TemplateDiagnostic>,
    docx: String,
    pdf: String,
    html_report: String,
    json_report: String,
    page_snapshots: String,
    checks: usize,
    failed_checks: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistry {
    schema_version: u32,
    registry_id: String,
    generated_at: String,
    base_url: String,
    categories: Vec<TemplateRegistryCategory>,
    template_of_the_month: String,
    templates: Vec<TemplateRegistryEntry>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryCategory {
    id: String,
    label: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryEntry {
    id: String,
    title: String,
    version: String,
    description: String,
    categories: Vec<String>,
    tags: Vec<String>,
    author: TemplateRegistryAuthor,
    license: TemplateRegistryLicense,
    supported_rusdox: TemplateRegistryVersionRange,
    preview: TemplateRegistryPreview,
    inputs: Vec<TemplateRegistryInput>,
    files: TemplateRegistryFiles,
    verified_outputs: TemplateRegistryOutputs,
    accessibility: TemplateRegistryAccessibility,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryAuthor {
    name: String,
    url: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryLicense {
    spdx: String,
    url: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryVersionRange {
    minimum: String,
    maximum_exclusive: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryAsset {
    url: String,
    sha256: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryPreview {
    url: String,
    sha256: String,
    alt: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryInput {
    path: String,
    #[serde(rename = "type")]
    input_type: String,
    required: bool,
    description: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryFiles {
    template: TemplateRegistryAsset,
    sample_data: TemplateRegistryAsset,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryOutputs {
    docx: TemplateRegistryAsset,
    pdf: TemplateRegistryAsset,
    parity_json: TemplateRegistryAsset,
    parity_html: TemplateRegistryAsset,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistryAccessibility {
    language: String,
    reading_order_reviewed: bool,
    color_only_meaning: bool,
    notes: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(deny_unknown_fields)]
struct TemplateRegistrySignature {
    schema_version: u32,
    algorithm: String,
    key_id: String,
    manifest_sha256: String,
    signature: String,
}

#[derive(Debug, Clone, Serialize)]
struct TemplateInstallResult {
    id: String,
    version: String,
    status: String,
    install_dir: String,
    template: String,
    sample_data: String,
}

#[derive(Debug, Clone, Serialize)]
struct VerifyFileResult {
    source: String,
    passed: bool,
    docx: String,
    pdf: String,
    html_report: String,
    json_report: String,
    page_snapshots: String,
    checks: usize,
    failed_checks: usize,
}

#[derive(Debug, Clone, Serialize)]
struct VerifyCommandResult {
    target: String,
    passed: bool,
    specs: usize,
    files: Vec<VerifyFileResult>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct WatchSnapshot {
    states: Vec<(PathBuf, u64)>,
}

#[derive(Debug, Clone, Serialize)]
struct DevTimings {
    parse_ms: f64,
    validate_ms: f64,
    compose_ms: f64,
    docx_ms: f64,
    pdf_ms: f64,
    total_ms: f64,
}

#[derive(Debug, Clone, Serialize)]
struct DevArtifact {
    source: String,
    docx: String,
    pdf: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
struct DevEvent {
    event: String,
    build: u32,
    status: String,
    reason: String,
    changed: Vec<String>,
    documents: usize,
    warnings: usize,
    timings: Option<DevTimings>,
    artifacts: Vec<DevArtifact>,
    error: Option<String>,
    dashboard: String,
}

#[derive(Debug, Clone)]
struct DevServerState {
    event: DevEvent,
    latest_pdf: Option<PathBuf>,
    latest_docx: Option<PathBuf>,
}

struct DevServer {
    url: String,
    state: Arc<RwLock<DevServerState>>,
}

fn main() {
    if let Err(error) = run() {
        eprintln!("error: {error}");
        std::process::exit(if matches!(error, DocxError::Parity(_)) {
            2
        } else {
            1
        });
    }
}

fn run() -> Result<()> {
    let cli = Cli::parse();
    if let Some(command) = cli.command {
        return match command {
            Commands::Config { command } => run_config_command(command),
            Commands::InitDoc(args) => init_doc(args),
            Commands::InitScript(args) => init_script(args),
            Commands::Schema(args) => run_schema(args),
            Commands::Migrate(args) => run_migrate(args),
            Commands::Validate(args) => run_validate(args),
            Commands::Watch(args) => run_watch(args),
            Commands::Dev(args) => run_dev(args),
            Commands::Serve(args) => run_serve(args),
            Commands::Bench(args) => run_bench(args),
            Commands::Template { command } => run_template_command(command),
            Commands::Verify(args) => run_verify(args),
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
    rusdox::atomic_write_file(&path, default_script_template().as_bytes())?;
    println!("{}", path.display());
    Ok(())
}

fn run_schema(args: SchemaArgs) -> Result<()> {
    let schema = document_spec_schema_pretty()?;
    if let Some(output) = args.output {
        atomic_write_file(&output, schema.as_bytes())?;
        println!("{}", output.display());
    } else {
        print!("{schema}");
    }
    Ok(())
}

fn run_migrate(args: MigrateArgs) -> Result<()> {
    if !args.input.is_file() {
        return Err(DocxError::Parse(format!(
            "migration input is not a file: {}",
            args.input.display()
        )));
    }
    let content = fs::read_to_string(&args.input)?;
    let extension = args
        .input
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or("yaml")
        .to_ascii_lowercase();
    let (version, migrated) = migrate_spec_content(&content, &extension)?;
    let needs_migration = version != Some(SPEC_VERSION);

    if args.check {
        if needs_migration {
            return Err(DocxError::Parse(format!(
                "{} needs migration to spec version {SPEC_VERSION}",
                args.input.display()
            )));
        }
        println!(
            "{} already uses spec version {SPEC_VERSION}",
            args.input.display()
        );
        return Ok(());
    }

    if let Some(output) = args.output {
        if output.exists() && !args.force {
            return Err(DocxError::Parse(format!(
                "migration output already exists at {} (use --force to replace it)",
                output.display()
            )));
        }
        atomic_write_file(&output, migrated.as_bytes())?;
        println!("{}", output.display());
    } else if args.in_place {
        atomic_write_file(&args.input, migrated.as_bytes())?;
        println!("{}", args.input.display());
    } else {
        print!("{migrated}");
    }
    Ok(())
}

fn migrate_spec_content(content: &str, extension: &str) -> Result<(Option<u32>, String)> {
    match extension {
        "yaml" | "yml" | "" => migrate_yaml_spec(content),
        "json" => migrate_json_spec(content),
        "toml" => migrate_toml_spec(content),
        other => Err(DocxError::Parse(format!(
            "unsupported migration extension '{other}', expected .yaml, .yml, .json, or .toml"
        ))),
    }
}

fn migrate_yaml_spec(content: &str) -> Result<(Option<u32>, String)> {
    let value: serde_yaml::Value = serde_yaml::from_str(content)
        .map_err(|error| DocxError::Parse(format!("invalid YAML document spec: {error}")))?;
    let version = mapping_spec_version(value.as_mapping(), "YAML")?;
    ensure_migratable_version(version)?;
    if version == Some(SPEC_VERSION) {
        return Ok((version, content.to_string()));
    }

    let mut lines = content.lines().map(str::to_string).collect::<Vec<_>>();
    if let Some(index) = lines.iter().position(|line| {
        let trimmed = line.trim_start();
        line.len() == trimmed.len() && trimmed.starts_with("version:")
    }) {
        lines[index] = format!("version: {SPEC_VERSION}");
    } else {
        let insertion = lines
            .iter()
            .position(|line| {
                let trimmed = line.trim();
                !trimmed.is_empty() && !trimmed.starts_with('#') && trimmed != "---"
            })
            .unwrap_or(lines.len());
        lines.insert(insertion, format!("version: {SPEC_VERSION}"));
    }
    let mut migrated = lines.join("\n");
    if content.ends_with('\n') || !migrated.is_empty() {
        migrated.push('\n');
    }
    serde_yaml::from_str::<serde_yaml::Value>(&migrated)
        .map_err(|error| DocxError::Parse(format!("migrated YAML failed to parse: {error}")))?;
    Ok((version, migrated))
}

fn migrate_json_spec(content: &str) -> Result<(Option<u32>, String)> {
    let value: serde_json::Value = serde_json::from_str(content)
        .map_err(|error| DocxError::Parse(format!("invalid JSON document spec: {error}")))?;
    let object = value
        .as_object()
        .ok_or_else(|| DocxError::Parse("JSON document spec root must be an object".to_string()))?;
    let version = json_spec_version(object.get("version"), "JSON")?;
    ensure_migratable_version(version)?;
    if version == Some(SPEC_VERSION) {
        return Ok((version, content.to_string()));
    }
    let mut migrated = serde_json::Map::new();
    migrated.insert("version".to_string(), serde_json::Value::from(SPEC_VERSION));
    for (key, value) in object {
        if key != "version" {
            migrated.insert(key.clone(), value.clone());
        }
    }
    let mut rendered = serde_json::to_string_pretty(&migrated)
        .map_err(|error| DocxError::Parse(format!("failed to serialize migrated JSON: {error}")))?;
    rendered.push('\n');
    Ok((version, rendered))
}

fn migrate_toml_spec(content: &str) -> Result<(Option<u32>, String)> {
    let mut value: toml::Value = toml::from_str(content)
        .map_err(|error| DocxError::Parse(format!("invalid TOML document spec: {error}")))?;
    let table = value
        .as_table_mut()
        .ok_or_else(|| DocxError::Parse("TOML document spec root must be a table".to_string()))?;
    let version = match table.get("version") {
        Some(value) => Some(
            value
                .as_integer()
                .and_then(|version| u32::try_from(version).ok())
                .ok_or_else(|| {
                    DocxError::Parse("TOML spec version must be a non-negative integer".to_string())
                })?,
        ),
        None => None,
    };
    ensure_migratable_version(version)?;
    if version == Some(SPEC_VERSION) {
        return Ok((version, content.to_string()));
    }
    table.insert(
        "version".to_string(),
        toml::Value::Integer(i64::from(SPEC_VERSION)),
    );
    let mut rendered = toml::to_string_pretty(&value)
        .map_err(|error| DocxError::Parse(format!("failed to serialize migrated TOML: {error}")))?;
    if !rendered.ends_with('\n') {
        rendered.push('\n');
    }
    Ok((version, rendered))
}

fn mapping_spec_version(
    mapping: Option<&serde_yaml::Mapping>,
    format: &str,
) -> Result<Option<u32>> {
    let mapping = mapping.ok_or_else(|| {
        DocxError::Parse(format!("{format} document spec root must be a mapping"))
    })?;
    let value = mapping.get(serde_yaml::Value::String("version".to_string()));
    match value {
        Some(serde_yaml::Value::Number(value)) => value
            .as_u64()
            .and_then(|version| u32::try_from(version).ok())
            .map(Some)
            .ok_or_else(|| {
                DocxError::Parse(format!(
                    "{format} spec version must be a non-negative integer"
                ))
            }),
        Some(_) => Err(DocxError::Parse(format!(
            "{format} spec version must be a non-negative integer"
        ))),
        None => Ok(None),
    }
}

fn json_spec_version(value: Option<&serde_json::Value>, format: &str) -> Result<Option<u32>> {
    match value {
        Some(value) => value
            .as_u64()
            .and_then(|version| u32::try_from(version).ok())
            .map(Some)
            .ok_or_else(|| {
                DocxError::Parse(format!(
                    "{format} spec version must be a non-negative integer"
                ))
            }),
        None => Ok(None),
    }
}

fn ensure_migratable_version(version: Option<u32>) -> Result<()> {
    if let Some(version) = version.filter(|version| *version > SPEC_VERSION) {
        return Err(DocxError::Parse(format!(
            "cannot migrate future spec version {version}; this build supports version {SPEC_VERSION}"
        )));
    }
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

fn run_validate(args: ValidateArgs) -> Result<()> {
    let config = load_runtime_config(args.config.as_deref())?;
    let config_report = validate_config(&config);
    let spec_paths = collect_spec_inputs(&args.input)?;
    let mut files = Vec::with_capacity(spec_paths.len());

    for spec_path in &spec_paths {
        let inspection = inspect_spec(spec_path)?;
        files.push(ValidationFileResult {
            path: spec_path.display().to_string(),
            parse_ms: duration_ms(inspection.parse_duration),
            validate_ms: duration_ms(inspection.validation_duration),
            issues: inspection.report.issues,
        });
    }

    let errors = config_report.error_count()
        + files
            .iter()
            .map(|file| {
                file.issues
                    .iter()
                    .filter(|issue| issue.severity == ValidationSeverity::Error)
                    .count()
            })
            .sum::<usize>();
    let warnings = config_report.warning_count()
        + files
            .iter()
            .map(|file| {
                file.issues
                    .iter()
                    .filter(|issue| issue.severity == ValidationSeverity::Warning)
                    .count()
            })
            .sum::<usize>();

    let result = ValidationCommandResult {
        target: args.input.display().to_string(),
        specs: spec_paths.len(),
        errors,
        warnings,
        config_issues: config_report.issues,
        files,
    };

    match args.format {
        ReportFormat::Text => {
            if result.errors > 0 {
                return Err(DocxError::Parse(format_validation_result_text(&result)));
            }
            println!("{}", format_validation_result_text(&result));
        }
        ReportFormat::Json => {
            println!(
                "{}",
                serde_json::to_string_pretty(&result).map_err(|error| {
                    DocxError::Parse(format!("failed to serialize validation report: {error}"))
                })?
            );
            if result.errors > 0 {
                return Err(DocxError::Parse("validation failed".to_string()));
            }
        }
    }

    Ok(())
}

fn run_watch(args: WatchArgs) -> Result<()> {
    if args.input.is_dir() && args.output.is_some() {
        return Err(DocxError::Parse(
            "--output is only supported for a single watched spec file".to_string(),
        ));
    }

    let poll_interval = Duration::from_millis(args.poll_interval_ms.max(50));
    let mut snapshot = capture_watch_snapshot(&args.input, args.config.as_deref())?;
    let mut build_attempts = 0_u32;
    let mut pending_reason = "initial build".to_string();

    loop {
        build_attempts += 1;
        println!("watch build {build_attempts}: {pending_reason}");

        let config = runtime_config(args.config.as_deref(), args.docx_only, args.with_pdf)?;
        match build_spec_input(&args.input, args.output.as_deref(), &config, true, true) {
            Ok(summary) => {
                println!(
                    "watch build {build_attempts} succeeded in {} across {} spec(s) (warnings: {})",
                    format_duration(summary.total_duration),
                    summary.documents,
                    summary.warning_count
                );
            }
            Err(error) => {
                eprintln!("watch build {build_attempts} failed: {error}");
            }
        }

        if args
            .max_builds
            .is_some_and(|limit| build_attempts >= limit.max(1))
        {
            break;
        }

        let (next_snapshot, changed) = wait_for_watch_change(
            &args.input,
            args.config.as_deref(),
            &snapshot,
            poll_interval,
            Duration::from_millis(args.debounce_ms),
        )?;
        pending_reason = if changed.is_empty() {
            "change detected".to_string()
        } else {
            format!(
                "change detected in {}",
                changed
                    .iter()
                    .map(|path| path.display().to_string())
                    .collect::<Vec<_>>()
                    .join(", ")
            )
        };
        snapshot = next_snapshot;
    }

    Ok(())
}

fn run_dev(args: DevArgs) -> Result<()> {
    if args.input.is_dir() && args.output.is_some() {
        return Err(DocxError::Parse(
            "--output is only supported for a single watched spec file".to_string(),
        ));
    }

    let poll_interval = Duration::from_millis(args.poll_interval_ms.max(50));
    let debounce = Duration::from_millis(args.debounce_ms);
    let mut snapshot = capture_watch_snapshot(&args.input, args.config.as_deref())?;
    let server = start_dev_server(args.port)?;
    emit_dev_lifecycle(&args, "listening", &server.url)?;
    if args.open {
        open_dev_dashboard(&server.url)?;
    }

    let mut build_attempts = 0_u32;
    let mut changed = Vec::new();
    let mut reason = "initial build".to_string();

    loop {
        build_attempts += 1;
        let config = runtime_config(args.config.as_deref(), args.docx_only, args.with_pdf);
        let result = config.and_then(|config| {
            build_spec_input(&args.input, args.output.as_deref(), &config, false, false)
                .map(|summary| (config, summary))
        });

        let changed_strings = changed
            .iter()
            .map(|path: &PathBuf| path.display().to_string())
            .collect::<Vec<_>>();
        match result {
            Ok((config, summary)) => {
                let resolved = resolve_dev_artifacts(&args.input, args.output.as_deref(), &config)?;
                let artifacts = resolved
                    .iter()
                    .map(|(artifact, _, _)| artifact.clone())
                    .collect::<Vec<_>>();
                let latest_docx = resolved.first().map(|(_, path, _)| path.clone());
                let latest_pdf = resolved
                    .first()
                    .and_then(|(_, _, path)| path.as_ref().cloned());
                let event = DevEvent {
                    event: "build".to_string(),
                    build: build_attempts,
                    status: "success".to_string(),
                    reason: reason.clone(),
                    changed: changed_strings,
                    documents: summary.documents,
                    warnings: summary.warning_count,
                    timings: Some(dev_timings(summary)),
                    artifacts,
                    error: None,
                    dashboard: server.url.clone(),
                };
                replace_dev_server_state(&server, event.clone(), latest_pdf, latest_docx)?;
                emit_dev_event(&args, &event)?;
            }
            Err(error) => {
                let previous = server
                    .state
                    .read()
                    .map_err(|_| DocxError::Parse("dev server state is poisoned".to_string()))?
                    .clone();
                let event = DevEvent {
                    event: "build".to_string(),
                    build: build_attempts,
                    status: "failed".to_string(),
                    reason: reason.clone(),
                    changed: changed_strings,
                    documents: previous.event.documents,
                    warnings: previous.event.warnings,
                    timings: previous.event.timings.clone(),
                    artifacts: previous.event.artifacts.clone(),
                    error: Some(error.to_string()),
                    dashboard: server.url.clone(),
                };
                replace_dev_server_state(
                    &server,
                    event.clone(),
                    previous.latest_pdf,
                    previous.latest_docx,
                )?;
                emit_dev_event(&args, &event)?;
            }
        }

        if args
            .max_builds
            .is_some_and(|limit| build_attempts >= limit.max(1))
        {
            break;
        }

        let (next_snapshot, next_changed) = wait_for_watch_change(
            &args.input,
            args.config.as_deref(),
            &snapshot,
            poll_interval,
            debounce,
        )?;
        reason = dev_change_reason(&next_changed, &args.input, args.config.as_deref());
        snapshot = next_snapshot;
        changed = next_changed;
    }

    Ok(())
}

fn dev_timings(summary: BuildSummary) -> DevTimings {
    DevTimings {
        parse_ms: duration_ms(summary.parse_duration),
        validate_ms: duration_ms(summary.validation_duration),
        compose_ms: duration_ms(summary.compose_duration),
        docx_ms: duration_ms(summary.output_stats.docx_write),
        pdf_ms: duration_ms(summary.output_stats.pdf_render),
        total_ms: duration_ms(summary.total_duration),
    }
}

fn emit_dev_lifecycle(args: &DevArgs, event: &str, dashboard: &str) -> Result<()> {
    if args.json {
        println!(
            "{}",
            serde_json::json!({ "event": event, "dashboard": dashboard })
        );
    } else if !args.quiet {
        println!("RusDox dev dashboard: {dashboard}");
    }
    Ok(())
}

fn emit_dev_event(args: &DevArgs, event: &DevEvent) -> Result<()> {
    if args.json {
        println!(
            "{}",
            serde_json::to_string(event).map_err(|error| {
                DocxError::Parse(format!("failed to serialize dev event: {error}"))
            })?
        );
    } else if !args.quiet {
        if event.status == "success" {
            println!(
                "dev build {} succeeded in {:.2} ms · {} · {} document(s), {} warning(s)",
                event.build,
                event
                    .timings
                    .as_ref()
                    .map_or(0.0, |timings| timings.total_ms),
                event.reason,
                event.documents,
                event.warnings,
            );
            for artifact in &event.artifacts {
                println!("  DOCX: {}", artifact.docx);
                if let Some(pdf) = &artifact.pdf {
                    println!("  PDF:  {pdf}");
                }
            }
        } else {
            eprintln!(
                "dev build {} failed · {}\n{}\nlast successful output is still available at {}",
                event.build,
                event.reason,
                event.error.as_deref().unwrap_or("unknown build failure"),
                event.dashboard,
            );
        }
    }
    Ok(())
}

fn replace_dev_server_state(
    server: &DevServer,
    event: DevEvent,
    latest_pdf: Option<PathBuf>,
    latest_docx: Option<PathBuf>,
) -> Result<()> {
    *server
        .state
        .write()
        .map_err(|_| DocxError::Parse("dev server state is poisoned".to_string()))? =
        DevServerState {
            event,
            latest_pdf,
            latest_docx,
        };
    Ok(())
}

fn resolve_dev_artifacts(
    input: &Path,
    output: Option<&Path>,
    config: &RusdoxConfig,
) -> Result<Vec<(DevArtifact, PathBuf, Option<PathBuf>)>> {
    let inputs = collect_spec_inputs(input)?;
    let mut artifacts = Vec::with_capacity(inputs.len());
    for spec_path in inputs {
        let spec = DocumentSpec::load_from_path(&spec_path)?;
        let output_name = spec
            .output_name
            .unwrap_or_else(|| default_output_name_for_spec(&spec_path));
        let docx_path = if let Some(path) = output {
            to_absolute_path(path)?
        } else {
            to_absolute_path(
                &Path::new(&config.output.docx_dir).join(format!("{output_name}.docx")),
            )?
        };
        let pdf_name = if output.is_some() {
            docx_path
                .file_stem()
                .and_then(|stem| stem.to_str())
                .unwrap_or(&output_name)
        } else {
            &output_name
        };
        let pdf_path = config.output.emit_pdf_preview.then(|| {
            to_absolute_path(&Path::new(&config.output.pdf_dir).join(format!("{pdf_name}.pdf")))
        });
        let pdf_path = match pdf_path {
            Some(path) => Some(path?),
            None => None,
        };
        artifacts.push((
            DevArtifact {
                source: spec_path.display().to_string(),
                docx: docx_path.display().to_string(),
                pdf: pdf_path.as_ref().map(|path| path.display().to_string()),
            },
            docx_path,
            pdf_path,
        ));
    }
    Ok(artifacts)
}

fn run_serve(args: ServeArgs) -> Result<()> {
    match args.transport {
        ServeTransport::Stdio => run_serve_stdio(&args),
        ServeTransport::Http => run_serve_http(&args),
    }
}

fn run_serve_stdio(args: &ServeArgs) -> Result<()> {
    let stdin = std::io::stdin();
    let mut reader = stdin.lock();
    let stdout = std::io::stdout();
    let mut writer = stdout.lock();
    let mut handled = 0_u32;
    while let Some((line, oversized)) = read_bounded_protocol_line(&mut reader)? {
        if line.iter().all(u8::is_ascii_whitespace) && !oversized {
            continue;
        }
        let response = if oversized {
            protocol_failure(
                "",
                "request_too_large",
                format!(
                    "JSON request exceeds the {} byte limit",
                    MAX_PROTOCOL_REQUEST_BYTES
                ),
            )
        } else {
            match serde_json::from_slice::<ProtocolRequest>(&line) {
                Ok(request) => execute_protocol_request(request, &args.output_root),
                Err(error) => protocol_failure("", "invalid_json", error.to_string()),
            }
        };
        serde_json::to_writer(&mut writer, &response).map_err(|error| {
            DocxError::Parse(format!("failed to serialize protocol response: {error}"))
        })?;
        writer.write_all(b"\n")?;
        writer.flush()?;
        handled += 1;
        if args.max_requests.is_some_and(|maximum| handled >= maximum) {
            break;
        }
    }
    Ok(())
}

fn read_bounded_protocol_line<R: BufRead>(
    reader: &mut R,
) -> std::io::Result<Option<(Vec<u8>, bool)>> {
    let mut line = Vec::new();
    let mut oversized = false;
    let mut saw_bytes = false;
    loop {
        let available = reader.fill_buf()?;
        if available.is_empty() {
            return if saw_bytes {
                Ok(Some((line, oversized)))
            } else {
                Ok(None)
            };
        }
        saw_bytes = true;
        let consumed = available
            .iter()
            .position(|byte| *byte == b'\n')
            .map_or(available.len(), |position| position + 1);
        if !oversized {
            let remaining = MAX_PROTOCOL_REQUEST_BYTES
                .saturating_add(1)
                .saturating_sub(line.len());
            line.extend_from_slice(&available[..consumed.min(remaining)]);
            oversized = line.len() > MAX_PROTOCOL_REQUEST_BYTES || consumed > remaining;
        }
        let ended = available[..consumed].ends_with(b"\n");
        reader.consume(consumed);
        if ended {
            while matches!(line.last(), Some(b'\n' | b'\r')) {
                line.pop();
            }
            return Ok(Some((line, oversized)));
        }
    }
}

fn run_serve_http(args: &ServeArgs) -> Result<()> {
    let listener = TcpListener::bind(SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), args.port))?;
    let address = listener.local_addr()?;
    eprintln!(
        "rusdox protocol v{PROTOCOL_VERSION} listening on http://127.0.0.1:{}/v1/request",
        address.port()
    );
    for (index, incoming) in listener.incoming().enumerate() {
        let mut stream = incoming?;
        stream.set_read_timeout(Some(Duration::from_secs(5)))?;
        stream.set_write_timeout(Some(Duration::from_secs(5)))?;
        serve_protocol_http_request(&mut stream, &args.output_root)?;
        if args
            .max_requests
            .is_some_and(|maximum| index + 1 >= maximum as usize)
        {
            break;
        }
    }
    Ok(())
}

struct ProtocolHttpRequest {
    method: String,
    path: String,
    content_type: Option<String>,
    body: Vec<u8>,
}

fn serve_protocol_http_request(stream: &mut TcpStream, output_root: &Path) -> Result<()> {
    let request = match read_protocol_http_request(stream) {
        Ok(request) => request,
        Err(error) => {
            return write_protocol_http_response(
                stream,
                "400 Bad Request",
                &protocol_failure("", "invalid_http", error.to_string()),
            );
        }
    };
    if request.method == "GET" && request.path == "/health" {
        let body = serde_json::to_vec(&serde_json::json!({
            "protocol_version": PROTOCOL_VERSION,
            "status": "ok",
            "transport": "loopback_http",
        }))
        .map_err(|error| {
            DocxError::Parse(format!("failed to serialize health response: {error}"))
        })?;
        return write_http_response(stream, "200 OK", "application/json; charset=utf-8", &body);
    }
    if request.method != "POST" || request.path != "/v1/request" {
        return write_protocol_http_response(
            stream,
            "404 Not Found",
            &protocol_failure("", "not_found", "use POST /v1/request or GET /health"),
        );
    }
    if !request
        .content_type
        .as_deref()
        .is_some_and(|value| value.to_ascii_lowercase().starts_with("application/json"))
    {
        return write_protocol_http_response(
            stream,
            "415 Unsupported Media Type",
            &protocol_failure(
                "",
                "unsupported_media_type",
                "Content-Type must be application/json",
            ),
        );
    }
    let response = match serde_json::from_slice::<ProtocolRequest>(&request.body) {
        Ok(request) => execute_protocol_request(request, output_root),
        Err(error) => protocol_failure("", "invalid_json", error.to_string()),
    };
    let status = if response.ok {
        "200 OK"
    } else if response.error.as_ref().is_some_and(|error| {
        matches!(
            error.code.as_str(),
            "validation_failed" | "render_failed" | "write_failed"
        )
    }) {
        "422 Unprocessable Content"
    } else {
        "400 Bad Request"
    };
    write_protocol_http_response(stream, status, &response)
}

fn read_protocol_http_request(stream: &mut TcpStream) -> Result<ProtocolHttpRequest> {
    const MAX_HEADERS: usize = 16 * 1024;
    let mut bytes = Vec::with_capacity(8192);
    let mut chunk = [0_u8; 8192];
    let (header_end, content_length) = loop {
        let read = stream.read(&mut chunk)?;
        if read == 0 {
            return Err(DocxError::Parse("incomplete HTTP request".to_string()));
        }
        bytes.extend_from_slice(&chunk[..read]);
        if bytes.len() > MAX_HEADERS + MAX_PROTOCOL_REQUEST_BYTES {
            return Err(DocxError::ResourceLimit(format!(
                "HTTP request exceeds {} bytes",
                MAX_HEADERS + MAX_PROTOCOL_REQUEST_BYTES
            )));
        }
        if let Some(position) = bytes.windows(4).position(|window| window == b"\r\n\r\n") {
            let header_end = position + 4;
            if header_end > MAX_HEADERS {
                return Err(DocxError::ResourceLimit(
                    "HTTP headers exceed 16384 bytes".to_string(),
                ));
            }
            let headers = std::str::from_utf8(&bytes[..position])
                .map_err(|error| DocxError::Parse(format!("invalid HTTP headers: {error}")))?;
            let content_length = headers
                .lines()
                .skip(1)
                .find_map(|line| {
                    line.split_once(':').and_then(|(name, value)| {
                        name.trim()
                            .eq_ignore_ascii_case("content-length")
                            .then(|| value.trim())
                    })
                })
                .map_or(Ok(0_usize), |value| {
                    value.parse::<usize>().map_err(|error| {
                        DocxError::Parse(format!("invalid Content-Length: {error}"))
                    })
                })?;
            if content_length > MAX_PROTOCOL_REQUEST_BYTES {
                return Err(DocxError::ResourceLimit(format!(
                    "JSON request exceeds the {} byte limit",
                    MAX_PROTOCOL_REQUEST_BYTES
                )));
            }
            break (header_end, content_length);
        }
    };
    while bytes.len() < header_end + content_length {
        let read = stream.read(&mut chunk)?;
        if read == 0 {
            return Err(DocxError::Parse("incomplete HTTP body".to_string()));
        }
        bytes.extend_from_slice(&chunk[..read]);
        if bytes.len() > header_end + MAX_PROTOCOL_REQUEST_BYTES {
            return Err(DocxError::ResourceLimit(format!(
                "JSON request exceeds the {} byte limit",
                MAX_PROTOCOL_REQUEST_BYTES
            )));
        }
    }
    let header_text = std::str::from_utf8(&bytes[..header_end - 4])
        .map_err(|error| DocxError::Parse(format!("invalid HTTP headers: {error}")))?;
    let mut lines = header_text.lines();
    let request_line = lines
        .next()
        .ok_or_else(|| DocxError::Parse("missing HTTP request line".to_string()))?;
    let mut request_parts = request_line.split_whitespace();
    let method = request_parts.next().unwrap_or_default().to_string();
    let path = request_parts
        .next()
        .unwrap_or_default()
        .split('?')
        .next()
        .unwrap_or_default()
        .to_string();
    if request_parts.next().is_none() || method.is_empty() || path.is_empty() {
        return Err(DocxError::Parse("invalid HTTP request line".to_string()));
    }
    let content_type = lines.find_map(|line| {
        line.split_once(':').and_then(|(name, value)| {
            name.trim()
                .eq_ignore_ascii_case("content-type")
                .then(|| value.trim().to_string())
        })
    });
    Ok(ProtocolHttpRequest {
        method,
        path,
        content_type,
        body: bytes[header_end..header_end + content_length].to_vec(),
    })
}

fn write_protocol_http_response(
    stream: &mut TcpStream,
    status: &str,
    response: &ProtocolResponse,
) -> Result<()> {
    let body = serde_json::to_vec(response).map_err(|error| {
        DocxError::Parse(format!("failed to serialize protocol response: {error}"))
    })?;
    write_http_response(stream, status, "application/json; charset=utf-8", &body)
}

fn start_dev_server(port: u16) -> Result<DevServer> {
    let listener = TcpListener::bind(SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port))?;
    listener.set_nonblocking(true)?;
    let address = listener.local_addr()?;
    let url = format!("http://127.0.0.1:{}/", address.port());
    let state = Arc::new(RwLock::new(DevServerState {
        event: DevEvent {
            event: "build".to_string(),
            build: 0,
            status: "starting".to_string(),
            reason: "waiting for initial build".to_string(),
            changed: Vec::new(),
            documents: 0,
            warnings: 0,
            timings: None,
            artifacts: Vec::new(),
            error: None,
            dashboard: url.clone(),
        },
        latest_pdf: None,
        latest_docx: None,
    }));
    let server_state = Arc::clone(&state);
    thread::spawn(move || loop {
        match listener.accept() {
            Ok((stream, _)) => {
                if let Err(error) = serve_dev_request(stream, &server_state) {
                    eprintln!("rusdox dev server error: {error}");
                }
            }
            Err(error) if error.kind() == std::io::ErrorKind::WouldBlock => {
                thread::sleep(Duration::from_millis(25));
            }
            Err(error) => {
                eprintln!("rusdox dev server stopped: {error}");
                break;
            }
        }
    });
    Ok(DevServer { url, state })
}

fn serve_dev_request(mut stream: TcpStream, state: &Arc<RwLock<DevServerState>>) -> Result<()> {
    stream.set_nonblocking(false)?;
    stream.set_read_timeout(Some(Duration::from_secs(2)))?;
    let mut request = [0_u8; 8192];
    let bytes = match stream.read(&mut request) {
        Ok(0) => return Ok(()),
        Ok(bytes) => bytes,
        Err(error)
            if matches!(
                error.kind(),
                std::io::ErrorKind::WouldBlock | std::io::ErrorKind::TimedOut
            ) =>
        {
            return Ok(());
        }
        Err(error) => return Err(error.into()),
    };
    let request = String::from_utf8_lossy(&request[..bytes]);
    let path = request
        .lines()
        .next()
        .and_then(|line| line.split_whitespace().nth(1))
        .unwrap_or("/")
        .split('?')
        .next()
        .unwrap_or("/");
    let state = state
        .read()
        .map_err(|_| DocxError::Parse("dev server state is poisoned".to_string()))?
        .clone();

    match path {
        "/" | "/index.html" => write_http_response(
            &mut stream,
            "200 OK",
            "text/html; charset=utf-8",
            render_dev_dashboard(&state).as_bytes(),
        )?,
        "/status.json" => {
            let body = serde_json::to_vec_pretty(&state.event).map_err(|error| {
                DocxError::Parse(format!("failed to serialize dev status: {error}"))
            })?;
            write_http_response(&mut stream, "200 OK", "application/json", &body)?;
        }
        "/latest.pdf" => {
            write_dev_artifact(&mut stream, state.latest_pdf.as_deref(), "application/pdf")?
        }
        "/latest.docx" => write_dev_artifact(
            &mut stream,
            state.latest_docx.as_deref(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )?,
        _ => write_http_response(
            &mut stream,
            "404 Not Found",
            "text/plain; charset=utf-8",
            b"Not found",
        )?,
    }
    Ok(())
}

fn write_dev_artifact(
    stream: &mut TcpStream,
    path: Option<&Path>,
    content_type: &str,
) -> Result<()> {
    let Some(path) = path.filter(|path| path.is_file()) else {
        return write_http_response(
            stream,
            "404 Not Found",
            "text/plain; charset=utf-8",
            b"No successful artifact is available yet.",
        );
    };
    let body = fs::read(path)?;
    write_http_response(stream, "200 OK", content_type, &body)
}

fn write_http_response(
    stream: &mut TcpStream,
    status: &str,
    content_type: &str,
    body: &[u8],
) -> Result<()> {
    write!(
        stream,
        "HTTP/1.1 {status}\r\nContent-Type: {content_type}\r\nContent-Length: {}\r\nCache-Control: no-store\r\nX-Content-Type-Options: nosniff\r\nContent-Security-Policy: default-src 'none'; style-src 'unsafe-inline'; frame-src 'self'; base-uri 'none'; form-action 'none'\r\nConnection: close\r\n\r\n",
        body.len()
    )?;
    stream.write_all(body)?;
    Ok(())
}

fn render_dev_dashboard(state: &DevServerState) -> String {
    let event = &state.event;
    let status_class = if event.status == "success" {
        "success"
    } else if event.status == "failed" {
        "failed"
    } else {
        "starting"
    };
    let timings = event.timings.as_ref().map_or_else(
        || "<p class=muted>No successful timing sample yet.</p>".to_string(),
        |timings| {
            format!(
                "<dl class=metrics><div><dt>Total</dt><dd>{:.2} ms</dd></div><div><dt>Parse</dt><dd>{:.2} ms</dd></div><div><dt>Validate</dt><dd>{:.2} ms</dd></div><div><dt>Compose</dt><dd>{:.2} ms</dd></div><div><dt>DOCX</dt><dd>{:.2} ms</dd></div><div><dt>PDF</dt><dd>{:.2} ms</dd></div></dl>",
                timings.total_ms,
                timings.parse_ms,
                timings.validate_ms,
                timings.compose_ms,
                timings.docx_ms,
                timings.pdf_ms,
            )
        },
    );
    let artifacts = if event.artifacts.is_empty() {
        "<p class=muted>No successful output yet.</p>".to_string()
    } else {
        event
            .artifacts
            .iter()
            .map(|artifact| {
                format!(
                    "<li><strong>{}</strong><code>{}</code>{}</li>",
                    dev_escape_html(&artifact.source),
                    dev_escape_html(&artifact.docx),
                    artifact
                        .pdf
                        .as_ref()
                        .map_or_else(String::new, |pdf| format!(
                            "<code>{}</code>",
                            dev_escape_html(pdf)
                        )),
                )
            })
            .collect::<String>()
    };
    let issue = event.error.as_ref().map_or_else(
        || "<p class=success-copy>Latest build completed successfully.</p>".to_string(),
        |error| {
            format!(
                "<pre class=error-copy>{}</pre><p class=muted>The previous successful DOCX/PDF remains available.</p>",
                dev_escape_html(error)
            )
        },
    );
    let preview = if state.latest_pdf.as_ref().is_some_and(|path| path.is_file()) {
        format!(
            "<iframe title=\"Latest RusDox PDF\" src=\"/latest.pdf?build={}\"></iframe>",
            event.build
        )
    } else {
        "<div class=empty>PDF preview is disabled or no successful PDF exists yet.</div>"
            .to_string()
    };

    format!(
        r#"<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="refresh" content="2"><title>RusDox dev</title><style>
:root{{--ink:#182019;--muted:#687068;--paper:#fffdf8;--bg:#eee9df;--line:#d5cdbf;--accent:#b85c30;--green:#17633a;--red:#9d2f2f}}*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.55 system-ui,sans-serif}}main{{width:min(1500px,calc(100% - 28px));margin:24px auto;display:grid;gap:16px}}header,.panel{{background:var(--paper);border:1px solid var(--line);padding:20px}}header{{display:flex;align-items:end;justify-content:space-between;gap:20px;border-top:5px solid var(--accent)}}h1,h2{{margin:.2rem 0;font-family:Georgia,serif}}h1{{font-size:clamp(2rem,5vw,4rem)}}h2{{font-size:1.3rem}}.eyebrow{{margin:0;color:var(--accent);font:bold .72rem ui-monospace,monospace;text-transform:uppercase;letter-spacing:.12em}}.badge{{padding:.65rem .8rem;border:2px solid currentColor;font:bold .75rem ui-monospace,monospace;text-transform:uppercase}}.badge.success{{color:var(--green)}}.badge.failed{{color:var(--red)}}.badge.starting{{color:var(--accent)}}.grid{{display:grid;grid-template-columns:minmax(320px,.68fr) minmax(0,1.32fr);gap:16px}}.stack{{display:grid;gap:16px;align-content:start}}.muted{{color:var(--muted)}}.metrics{{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:12px 0 0}}.metrics div{{padding:10px;background:#f5f0e7}}dt{{color:var(--muted);font-size:.72rem;text-transform:uppercase}}dd{{margin:3px 0 0;font-weight:800}}ul{{list-style:none;padding:0;margin:0;display:grid;gap:8px}}li{{display:grid;gap:4px;padding:10px;background:#f5f0e7}}code,pre{{font-family:ui-monospace,monospace;overflow-wrap:anywhere}}.error-copy{{max-height:260px;overflow:auto;padding:14px;background:#2a1818;color:#ffdada;white-space:pre-wrap}}.success-copy{{color:var(--green);font-weight:700}}iframe,.empty{{width:100%;height:calc(100vh - 170px);min-height:640px;border:1px solid var(--line);background:white}}.empty{{display:grid;place-items:center;padding:30px;color:var(--muted)}}.actions{{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}}a{{padding:9px 12px;background:var(--accent);color:white;text-decoration:none;font-weight:700}}@media(max-width:900px){{.grid{{grid-template-columns:1fr}}iframe,.empty{{height:70vh;min-height:480px}}}}@media(max-width:520px){{header{{align-items:start;flex-direction:column}}.metrics{{grid-template-columns:repeat(2,1fr)}}}}
</style></head><body><main><header><div><p class="eyebrow">RusDox local feedback loop</p><h1>Build {build}</h1><p class="muted">{reason}</p></div><strong class="badge {status_class}">{status}</strong></header><div class="grid"><div class="stack"><section class="panel"><p class="eyebrow">Validation</p><h2>Current status</h2>{issue}</section><section class="panel"><p class="eyebrow">Timings</p><h2>Latest successful build</h2>{timings}</section><section class="panel"><p class="eyebrow">Artifacts</p><h2>Output paths</h2><ul>{artifacts}</ul><div class="actions"><a href="/latest.docx">DOCX</a><a href="/status.json">status.json</a></div></section></div><section class="panel"><p class="eyebrow">Preview</p><h2>Latest successful PDF</h2>{preview}</section></div></main></body></html>"#,
        build = event.build,
        reason = dev_escape_html(&event.reason),
        status_class = status_class,
        status = dev_escape_html(&event.status),
        issue = issue,
        timings = timings,
        artifacts = artifacts,
        preview = preview,
    )
}

fn dev_escape_html(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&#39;")
}

fn open_dev_dashboard(url: &str) -> Result<()> {
    let mut command = if cfg!(target_os = "macos") {
        Command::new("open")
    } else if cfg!(target_os = "windows") {
        let mut command = Command::new("cmd");
        command.args(["/C", "start", ""]);
        command
    } else {
        Command::new("xdg-open")
    };
    command.arg(url).spawn().map_err(|error| {
        DocxError::Parse(format!(
            "could not open the dev dashboard at {url}: {error}"
        ))
    })?;
    Ok(())
}

fn dev_change_reason(changed: &[PathBuf], input: &Path, config: Option<&Path>) -> String {
    if changed.is_empty() {
        return "change detected".to_string();
    }
    changed
        .iter()
        .map(|path| {
            let kind = if is_config_watch_path(path, config) {
                "config"
            } else if path == input || (input.is_dir() && path.parent() == Some(input)) {
                "input"
            } else {
                "asset/include"
            };
            format!("{kind} changed: {}", path.display())
        })
        .collect::<Vec<_>>()
        .join("; ")
}

fn is_config_watch_path(path: &Path, config: Option<&Path>) -> bool {
    config.is_some_and(|config| path == config)
        || path == Path::new(DEFAULT_CONFIG_FILE)
        || default_user_config_path().as_deref() == Some(path)
}

fn run_bench(args: BenchArgs) -> Result<()> {
    if args.iterations == 0 {
        return Err(DocxError::Parse(
            "--iterations must be greater than zero".to_string(),
        ));
    }

    let mut config = runtime_config(args.config.as_deref(), args.docx_only, args.with_pdf)?;
    let pipeline = args.pipeline.unwrap_or({
        if config.output.emit_pdf_preview {
            BenchPipeline::Dual
        } else {
            BenchPipeline::Docx
        }
    });
    let inputs = if pipeline == BenchPipeline::ExistingDocx {
        if !args.input.is_file()
            || args
                .input
                .extension()
                .and_then(|extension| extension.to_str())
                .is_none_or(|extension| !extension.eq_ignore_ascii_case("docx"))
        {
            return Err(DocxError::Parse(
                "--pipeline existing-docx requires one .docx file".to_string(),
            ));
        }
        vec![args.input.clone()]
    } else {
        collect_spec_inputs(&args.input)?
    };

    config.output.emit_pdf_preview = pipeline == BenchPipeline::Dual;
    let temp = if args.keep_output {
        None
    } else {
        let temp = tempdir()?;
        config.output.docx_dir = temp.path().join("generated").to_string_lossy().to_string();
        config.output.pdf_dir = temp.path().join("rendered").to_string_lossy().to_string();
        Some(temp)
    };

    if pipeline != BenchPipeline::ExistingDocx {
        let config_report = validate_config(&config);
        handle_validation_issues(
            "config",
            &config_report,
            true,
            "benchmark target has validation errors",
        )?;
    }

    for _ in 0..args.warmup {
        let _ = bench_once(&inputs, &config, pipeline, false)?;
    }

    let mut samples = Vec::with_capacity(args.iterations as usize);
    for iteration in 0..args.iterations {
        let sample = bench_once(&inputs, &config, pipeline, iteration == 0)?;
        samples.push(sample);
    }

    drop(temp);
    let (input_sha256, input_bytes) = input_fingerprint(&args.input, &inputs)?;

    let result = BenchCommandResult {
        schema_version: 2,
        target: args.input.display().to_string(),
        pipeline: pipeline.as_str().to_string(),
        input_sha256,
        input_bytes,
        specs: inputs.len(),
        iterations: args.iterations,
        warmup: args.warmup,
        emit_pdf: matches!(pipeline, BenchPipeline::Pdf | BenchPipeline::Dual),
        keep_output: args.keep_output,
        parse_ms: summarize_f64(
            samples
                .iter()
                .map(|sample| duration_ms(sample.parse_duration)),
        ),
        validate_ms: summarize_f64(
            samples
                .iter()
                .map(|sample| duration_ms(sample.validation_duration)),
        ),
        compose_ms: summarize_f64(
            samples
                .iter()
                .map(|sample| duration_ms(sample.compose_duration)),
        ),
        docx_ms: summarize_f64(
            samples
                .iter()
                .map(|sample| duration_ms(sample.docx_duration)),
        ),
        pdf_ms: summarize_f64(
            samples
                .iter()
                .map(|sample| duration_ms(sample.pdf_duration)),
        ),
        existing_docx_open_ms: summarize_f64(
            samples
                .iter()
                .map(|sample| duration_ms(sample.existing_docx_open_duration)),
        ),
        existing_docx_save_ms: summarize_f64(
            samples
                .iter()
                .map(|sample| duration_ms(sample.existing_docx_save_duration)),
        ),
        total_ms: summarize_f64(
            samples
                .iter()
                .map(|sample| duration_ms(sample.total_duration)),
        ),
        docx_bytes: summarize_f64(samples.iter().map(|sample| sample.docx_bytes as f64)),
        pdf_bytes: summarize_f64(samples.iter().map(|sample| sample.pdf_bytes as f64)),
    };

    match args.format {
        ReportFormat::Text => println!("{}", format_bench_result_text(&result)),
        ReportFormat::Json => println!(
            "{}",
            serde_json::to_string_pretty(&result).map_err(|error| {
                DocxError::Parse(format!("failed to serialize benchmark report: {error}"))
            })?
        ),
    }

    Ok(())
}

fn run_template_command(command: TemplateCommand) -> Result<()> {
    match command {
        TemplateCommand::Inspect(args) => run_template_inspect(args),
        TemplateCommand::Render(args) => run_template_render("render", args),
        TemplateCommand::Verify(args) => run_template_render("verify", args),
        TemplateCommand::List(args) => run_template_list(args),
        TemplateCommand::Search(args) => run_template_search(args),
        TemplateCommand::Add(args) => run_template_add(args),
        TemplateCommand::Update(args) => run_template_update(args),
    }
}

fn run_template_list(args: TemplateListArgs) -> Result<()> {
    let registry = load_template_registry(&args.registry, args.public_key.as_deref())?;
    print_template_registry_entries(&registry, &registry.templates, args.format)
}

fn run_template_search(args: TemplateSearchArgs) -> Result<()> {
    let registry = load_template_registry(&args.registry, args.public_key.as_deref())?;
    let query = args.query.trim().to_ascii_lowercase();
    let matches = registry
        .templates
        .iter()
        .filter(|entry| {
            [
                entry.id.as_str(),
                entry.title.as_str(),
                entry.description.as_str(),
                entry.author.name.as_str(),
            ]
            .iter()
            .any(|value| value.to_ascii_lowercase().contains(&query))
                || entry
                    .categories
                    .iter()
                    .chain(&entry.tags)
                    .any(|value| value.to_ascii_lowercase().contains(&query))
        })
        .collect::<Vec<_>>();
    print_template_registry_entries(&registry, &matches, args.format)
}

fn run_template_add(args: TemplateAddArgs) -> Result<()> {
    let registry = load_template_registry(&args.registry, args.public_key.as_deref())?;
    let entry = registry
        .templates
        .iter()
        .find(|entry| entry.id == args.id)
        .ok_or_else(|| DocxError::Parse(format!("template '{}' was not found", args.id)))?;
    require_supported_template_version(entry)?;
    let install_root = resolve_template_install_root(args.install_root)?;
    let result = install_registry_template(entry, &install_root, args.force)?;
    print_template_install_results(&[result], args.format)
}

fn run_template_update(args: TemplateUpdateArgs) -> Result<()> {
    let registry = load_template_registry(&args.registry, args.public_key.as_deref())?;
    let install_root = resolve_template_install_root(args.install_root)?;
    let ids = if args.all {
        installed_template_ids(&install_root)?
    } else {
        vec![args
            .id
            .ok_or_else(|| DocxError::Parse("provide a template ID or use --all".to_string()))?]
    };
    let mut results = Vec::new();
    for id in ids {
        let entry = registry
            .templates
            .iter()
            .find(|entry| entry.id == id)
            .ok_or_else(|| {
                DocxError::Parse(format!(
                    "installed template '{id}' is not present in the trusted registry"
                ))
            })?;
        require_supported_template_version(entry)?;
        results.push(install_registry_template(entry, &install_root, false)?);
    }
    print_template_install_results(&results, args.format)
}

fn print_template_registry_entries<T>(
    registry: &TemplateRegistry,
    entries: &[T],
    format: ReportFormat,
) -> Result<()>
where
    T: std::borrow::Borrow<TemplateRegistryEntry> + Serialize,
{
    match format {
        ReportFormat::Json => println!(
            "{}",
            serde_json::to_string_pretty(&serde_json::json!({
                "registry": registry.registry_id,
                "signed": true,
                "template_of_the_month": registry.template_of_the_month,
                "templates": entries,
            }))
            .map_err(|error| DocxError::Parse(format!("failed to serialize registry: {error}")))?
        ),
        ReportFormat::Text => {
            println!(
                "{} verified template(s) · signed registry {}",
                entries.len(),
                registry.registry_id
            );
            for value in entries {
                let entry = value.borrow();
                let featured = if entry.id == registry.template_of_the_month {
                    " · template of the month"
                } else {
                    ""
                };
                println!(
                    "{} v{} · {} · by {}{}\n  {}\n  categories: {} · license: {}",
                    entry.id,
                    entry.version,
                    entry.title,
                    entry.author.name,
                    featured,
                    entry.description,
                    entry.categories.join(", "),
                    entry.license.spdx,
                );
            }
        }
    }
    Ok(())
}

fn print_template_install_results(
    results: &[TemplateInstallResult],
    format: ReportFormat,
) -> Result<()> {
    match format {
        ReportFormat::Json => println!(
            "{}",
            serde_json::to_string_pretty(results).map_err(|error| {
                DocxError::Parse(format!("failed to serialize install results: {error}"))
            })?
        ),
        ReportFormat::Text => {
            if results.is_empty() {
                println!("no installed templates to update");
            }
            for result in results {
                println!(
                    "{} v{} · {}\n  template: {}\n  sample data: {}",
                    result.id, result.version, result.status, result.template, result.sample_data,
                );
            }
        }
    }
    Ok(())
}

fn load_template_registry(source: &str, public_key: Option<&str>) -> Result<TemplateRegistry> {
    let bytes = read_registry_resource(source, MAX_REGISTRY_BYTES)?;
    let signature_source = registry_signature_source(source)?;
    let signature_bytes = read_registry_resource(&signature_source, 64 * 1024)?;
    let signature: TemplateRegistrySignature = serde_json::from_slice(&signature_bytes)
        .map_err(|error| DocxError::Parse(format!("invalid registry signature file: {error}")))?;
    verify_template_registry_signature(
        &bytes,
        &signature,
        public_key.unwrap_or(DEFAULT_TEMPLATE_REGISTRY_PUBLIC_KEY),
    )?;
    let registry: TemplateRegistry = serde_json::from_slice(&bytes)
        .map_err(|error| DocxError::Parse(format!("invalid template registry: {error}")))?;
    validate_template_registry(&registry)?;
    Ok(registry)
}

fn verify_template_registry_signature(
    manifest: &[u8],
    signature: &TemplateRegistrySignature,
    public_key_hex: &str,
) -> Result<()> {
    if signature.schema_version != 1
        || signature.algorithm != "ed25519"
        || signature.key_id.trim().is_empty()
    {
        return Err(DocxError::Parse(
            "unsupported template registry signature contract".to_string(),
        ));
    }
    let digest = format!("{:x}", Sha256::digest(manifest));
    if digest != signature.manifest_sha256 {
        return Err(DocxError::Parse(
            "template registry SHA-256 does not match its signature record".to_string(),
        ));
    }
    let public_key = decode_fixed_hex::<32>(public_key_hex, "registry public key")?;
    let signature_bytes = decode_fixed_hex::<64>(&signature.signature, "registry signature")?;
    let verifying_key = VerifyingKey::from_bytes(&public_key)
        .map_err(|error| DocxError::Parse(format!("invalid registry public key: {error}")))?;
    verifying_key
        .verify(manifest, &Signature::from_bytes(&signature_bytes))
        .map_err(|_| DocxError::Parse("template registry signature verification failed".into()))
}

fn decode_fixed_hex<const N: usize>(value: &str, label: &str) -> Result<[u8; N]> {
    let value = value.trim();
    if value.len() != N * 2 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(DocxError::Parse(format!(
            "{label} must contain exactly {} hexadecimal characters",
            N * 2
        )));
    }
    let mut output = [0_u8; N];
    for (index, byte) in output.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16)
            .map_err(|error| DocxError::Parse(format!("invalid {label}: {error}")))?;
    }
    Ok(output)
}

fn registry_signature_source(source: &str) -> Result<String> {
    if source.starts_with("https://") || source.starts_with("http://") {
        let (prefix, _) = source.rsplit_once('/').ok_or_else(|| {
            DocxError::Parse("registry URL must include an index filename".to_string())
        })?;
        Ok(format!("{prefix}/index.sig.json"))
    } else {
        let path = source.strip_prefix("file://").unwrap_or(source);
        let mut signature = PathBuf::from(path);
        signature.set_file_name("index.sig.json");
        Ok(signature.display().to_string())
    }
}

fn read_registry_resource(source: &str, limit: u64) -> Result<Vec<u8>> {
    if source.starts_with("http://") && !is_loopback_url(source) {
        return Err(DocxError::Parse(
            "template registry downloads require HTTPS; plain HTTP is allowed only on loopback"
                .to_string(),
        ));
    }
    if source.starts_with("https://") || source.starts_with("http://") {
        let mut response = ureq::get(source)
            .call()
            .map_err(|error| DocxError::Parse(format!("failed to download {source}: {error}")))?;
        response
            .body_mut()
            .with_config()
            .limit(limit)
            .read_to_vec()
            .map_err(|error| {
                DocxError::Parse(format!(
                    "failed to read bounded response from {source}: {error}"
                ))
            })
    } else {
        let path = Path::new(source.strip_prefix("file://").unwrap_or(source));
        let metadata = fs::metadata(path)?;
        if metadata.len() > limit {
            return Err(DocxError::ResourceLimit(format!(
                "{} is {} bytes; registry limit is {limit} bytes",
                path.display(),
                metadata.len()
            )));
        }
        fs::read(path).map_err(Into::into)
    }
}

fn is_loopback_url(source: &str) -> bool {
    source
        .strip_prefix("http://")
        .and_then(|rest| rest.split('/').next())
        .is_some_and(|authority| {
            matches!(authority, "127.0.0.1" | "localhost" | "[::1]")
                || authority.starts_with("127.0.0.1:")
                || authority.starts_with("localhost:")
                || authority.starts_with("[::1]:")
        })
}

fn validate_template_registry(registry: &TemplateRegistry) -> Result<()> {
    if registry.schema_version != 1 || registry.registry_id.trim().is_empty() {
        return Err(DocxError::Parse(
            "unsupported or unnamed template registry".to_string(),
        ));
    }
    let required_categories = [
        "invoices",
        "proposals",
        "reports",
        "compliance",
        "hr",
        "education",
        "operations",
    ];
    let category_ids = registry
        .categories
        .iter()
        .map(|category| category.id.as_str())
        .collect::<BTreeSet<_>>();
    for category in required_categories {
        if !category_ids.contains(category) {
            return Err(DocxError::Parse(format!(
                "template registry is missing required category '{category}'"
            )));
        }
    }
    let mut ids = BTreeSet::new();
    for entry in &registry.templates {
        if !is_safe_template_id(&entry.id) || !ids.insert(entry.id.as_str()) {
            return Err(DocxError::Parse(format!(
                "template ID '{}' is unsafe or duplicated",
                entry.id
            )));
        }
        if entry.title.trim().is_empty()
            || entry.description.trim().is_empty()
            || entry.author.name.trim().is_empty()
            || entry.author.url.trim().is_empty()
            || entry.license.spdx.trim().is_empty()
            || entry.license.url.trim().is_empty()
            || entry.preview.alt.trim().is_empty()
            || entry.inputs.is_empty()
            || entry.accessibility.language.trim().is_empty()
            || entry.accessibility.notes.trim().is_empty()
            || !entry.accessibility.reading_order_reviewed
            || entry.accessibility.color_only_meaning
        {
            return Err(DocxError::Parse(format!(
                "template '{}' does not satisfy the curated metadata contract",
                entry.id
            )));
        }
        if entry.categories.is_empty()
            || entry
                .categories
                .iter()
                .any(|category| !category_ids.contains(category.as_str()))
        {
            return Err(DocxError::Parse(format!(
                "template '{}' has an unknown or empty category set",
                entry.id
            )));
        }
        for asset in [
            TemplateRegistryAsset {
                url: entry.preview.url.clone(),
                sha256: entry.preview.sha256.clone(),
            },
            entry.files.template.clone(),
            entry.files.sample_data.clone(),
            entry.verified_outputs.docx.clone(),
            entry.verified_outputs.pdf.clone(),
            entry.verified_outputs.parity_json.clone(),
            entry.verified_outputs.parity_html.clone(),
        ] {
            validate_registry_asset(&entry.id, &asset)?;
        }
    }
    if !ids.contains(registry.template_of_the_month.as_str()) {
        return Err(DocxError::Parse(
            "template_of_the_month must reference a registry entry".to_string(),
        ));
    }
    Ok(())
}

fn validate_registry_asset(template_id: &str, asset: &TemplateRegistryAsset) -> Result<()> {
    if asset.url.trim().is_empty()
        || asset.sha256.len() != 64
        || !asset.sha256.bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        return Err(DocxError::Parse(format!(
            "template '{template_id}' contains an invalid URL or SHA-256"
        )));
    }
    Ok(())
}

fn is_safe_template_id(value: &str) -> bool {
    !value.is_empty()
        && !value.starts_with('-')
        && value
            .bytes()
            .all(|byte| byte.is_ascii_lowercase() || byte.is_ascii_digit() || byte == b'-')
}

fn require_supported_template_version(entry: &TemplateRegistryEntry) -> Result<()> {
    let current = parse_numeric_version(env!("CARGO_PKG_VERSION"))?;
    let minimum = parse_numeric_version(&entry.supported_rusdox.minimum)?;
    let maximum = parse_numeric_version(&entry.supported_rusdox.maximum_exclusive)?;
    if current < minimum || current >= maximum {
        return Err(DocxError::Parse(format!(
            "template '{}' v{} supports RusDox >= {} and < {}; current version is {}",
            entry.id,
            entry.version,
            entry.supported_rusdox.minimum,
            entry.supported_rusdox.maximum_exclusive,
            env!("CARGO_PKG_VERSION"),
        )));
    }
    Ok(())
}

fn parse_numeric_version(value: &str) -> Result<(u64, u64, u64)> {
    let core = value.split(['-', '+']).next().unwrap_or(value);
    let mut parts = core.split('.');
    let parse = |part: Option<&str>| {
        part.unwrap_or("0").parse::<u64>().map_err(|error| {
            DocxError::Parse(format!("invalid numeric version '{value}': {error}"))
        })
    };
    let parsed = (
        parse(parts.next())?,
        parse(parts.next())?,
        parse(parts.next())?,
    );
    if parts.next().is_some() {
        return Err(DocxError::Parse(format!(
            "invalid numeric version '{value}'"
        )));
    }
    Ok(parsed)
}

fn resolve_template_install_root(explicit: Option<PathBuf>) -> Result<PathBuf> {
    if let Some(path) = explicit {
        return to_absolute_path(&path);
    }
    default_template_install_root().ok_or_else(|| {
        DocxError::Parse(
            "could not resolve a user data directory; pass --install-root explicitly".to_string(),
        )
    })
}

fn default_template_install_root() -> Option<PathBuf> {
    if cfg!(target_os = "windows") {
        std::env::var_os("LOCALAPPDATA")
            .map(PathBuf::from)
            .map(|path| path.join("RusDox/templates"))
    } else if cfg!(target_os = "macos") {
        std::env::var_os("HOME")
            .map(PathBuf::from)
            .map(|path| path.join("Library/Application Support/RusDox/templates"))
    } else {
        std::env::var_os("XDG_DATA_HOME")
            .map(PathBuf::from)
            .or_else(|| {
                std::env::var_os("HOME")
                    .map(PathBuf::from)
                    .map(|path| path.join(".local/share"))
            })
            .map(|path| path.join("rusdox/templates"))
    }
}

fn install_registry_template(
    entry: &TemplateRegistryEntry,
    install_root: &Path,
    force: bool,
) -> Result<TemplateInstallResult> {
    let install_dir = install_root.join(&entry.id);
    let manifest_path = install_dir.join("manifest.json");
    if !force && manifest_path.is_file() {
        let installed: TemplateRegistryEntry = serde_json::from_slice(&fs::read(&manifest_path)?)
            .map_err(|error| {
            DocxError::Parse(format!(
                "invalid installed manifest at {}: {error}",
                manifest_path.display()
            ))
        })?;
        if installed.version == entry.version
            && installed.files.template.sha256 == entry.files.template.sha256
            && installed.files.sample_data.sha256 == entry.files.sample_data.sha256
        {
            return Ok(template_install_result(entry, &install_dir, "up-to-date"));
        }
    }

    let template = download_and_verify_registry_asset(&entry.files.template)?;
    let sample_data = download_and_verify_registry_asset(&entry.files.sample_data)?;
    let _: serde_json::Value = serde_json::from_slice(&sample_data).map_err(|error| {
        DocxError::Parse(format!(
            "template '{}' sample data is not valid JSON: {error}",
            entry.id
        ))
    })?;
    let manifest = serde_json::to_vec_pretty(entry).map_err(|error| {
        DocxError::Parse(format!("failed to serialize installed manifest: {error}"))
    })?;

    fs::create_dir_all(&install_dir)?;
    atomic_write_file(install_dir.join("template.docx"), &template)?;
    atomic_write_file(install_dir.join("data.json"), &sample_data)?;
    atomic_write_file(&manifest_path, &manifest)?;
    Ok(template_install_result(entry, &install_dir, "installed"))
}

fn template_install_result(
    entry: &TemplateRegistryEntry,
    install_dir: &Path,
    status: &str,
) -> TemplateInstallResult {
    TemplateInstallResult {
        id: entry.id.clone(),
        version: entry.version.clone(),
        status: status.to_string(),
        install_dir: install_dir.display().to_string(),
        template: install_dir.join("template.docx").display().to_string(),
        sample_data: install_dir.join("data.json").display().to_string(),
    }
}

fn download_and_verify_registry_asset(asset: &TemplateRegistryAsset) -> Result<Vec<u8>> {
    let bytes = read_registry_resource(&asset.url, MAX_TEMPLATE_DOWNLOAD_BYTES)?;
    let actual = format!("{:x}", Sha256::digest(&bytes));
    if actual != asset.sha256 {
        return Err(DocxError::Parse(format!(
            "download hash mismatch for {}: expected {}, got {actual}",
            asset.url, asset.sha256
        )));
    }
    Ok(bytes)
}

fn installed_template_ids(install_root: &Path) -> Result<Vec<String>> {
    if !install_root.exists() {
        return Ok(Vec::new());
    }
    let mut ids = fs::read_dir(install_root)?
        .filter_map(|entry| entry.ok())
        .filter(|entry| entry.path().join("manifest.json").is_file())
        .filter_map(|entry| entry.file_name().to_str().map(ToOwned::to_owned))
        .collect::<Vec<_>>();
    ids.sort();
    Ok(ids)
}

fn run_template_inspect(args: TemplateInspectArgs) -> Result<()> {
    let template = DocxTemplate::open(&args.template)?;
    let inspection = template.inspect();
    match args.format {
        ReportFormat::Json => println!(
            "{}",
            serde_json::to_string_pretty(&inspection).map_err(|error| {
                DocxError::Parse(format!("failed to serialize template inspection: {error}"))
            })?
        ),
        ReportFormat::Text => {
            println!("template syntax: v{}", inspection.syntax_version);
            println!("placeholders: {}", inspection.placeholders.len());
            for placeholder in &inspection.placeholders {
                println!(
                    "  {} · {} · {} · {}",
                    placeholder.part,
                    placeholder.location,
                    placeholder.kind,
                    placeholder.expression
                );
            }
            for diagnostic in &inspection.diagnostics {
                eprintln!(
                    "  [{:?}] {} · {} · {}: {} ({})",
                    diagnostic.severity,
                    diagnostic.part,
                    diagnostic.location,
                    diagnostic.placeholder,
                    diagnostic.message,
                    diagnostic.suggestion
                );
            }
        }
    }
    if inspection.has_errors() {
        Err(DocxError::Parse(
            "template inspection found structural errors".to_string(),
        ))
    } else {
        Ok(())
    }
}

fn run_template_render(command: &str, args: TemplateRenderArgs) -> Result<()> {
    let data_bytes = fs::read(&args.data)?;
    let max_data_bytes = rusdox::InputLimits::default().max_spec_bytes;
    if data_bytes.len() as u64 > max_data_bytes {
        return Err(DocxError::ResourceLimit(format!(
            "template data is {} bytes; limit is {max_data_bytes} bytes",
            data_bytes.len()
        )));
    }
    let data: serde_json::Value = serde_json::from_slice(&data_bytes)
        .map_err(|error| DocxError::Parse(format!("invalid template JSON data: {error}")))?;
    if !data.is_object() {
        return Err(DocxError::Parse(
            "template JSON data must be an object".to_string(),
        ));
    }

    let template = DocxTemplate::open(&args.template)?;
    let inspection = template.inspect();
    if inspection.has_errors() {
        return Err(DocxError::Parse(format!(
            "template inspection failed: {}",
            inspection
                .diagnostics
                .iter()
                .map(|diagnostic| format!(
                    "{} · {} · {}",
                    diagnostic.part, diagnostic.location, diagnostic.message
                ))
                .collect::<Vec<_>>()
                .join("; ")
        )));
    }

    let output_root = to_absolute_path(&args.output_root)?;
    let generated_dir = output_root.join("generated");
    let rendered_dir = output_root.join("rendered");
    let reports_dir = output_root.join("reports");
    fs::create_dir_all(&generated_dir)?;
    fs::create_dir_all(&rendered_dir)?;
    fs::create_dir_all(&reports_dir)?;
    let output_name = args
        .name
        .as_deref()
        .map(safe_output_name)
        .filter(|name| !name.is_empty())
        .unwrap_or_else(|| {
            args.template
                .file_stem()
                .and_then(|value| value.to_str())
                .map(safe_output_name)
                .filter(|name| !name.is_empty())
                .unwrap_or_else(|| "rendered-template".to_string())
        });
    let docx_path = generated_dir.join(format!("{output_name}.docx"));
    let pdf_path = rendered_dir.join(format!("{output_name}.pdf"));
    let page_snapshots_dir = reports_dir.join(format!("{output_name}-pages"));
    let html_report_path = reports_dir.join(format!("{output_name}-parity.html"));
    let json_report_path = reports_dir.join(format!("{output_name}-parity.json"));

    let render = template.render_to_path(&data, &docx_path, args.strict)?;
    if !render.written {
        match args.format {
            ReportFormat::Json => println!(
                "{}",
                serde_json::to_string_pretty(&render).map_err(|error| {
                    DocxError::Parse(format!("failed to serialize template render: {error}"))
                })?
            ),
            ReportFormat::Text => {
                for diagnostic in &render.diagnostics {
                    eprintln!(
                        "{} · {} · {}: {} ({})",
                        diagnostic.part,
                        diagnostic.location,
                        diagnostic.placeholder,
                        diagnostic.message,
                        diagnostic.suggestion
                    );
                }
            }
        }
        return Err(DocxError::Parse(
            "template render failed; no output was written".to_string(),
        ));
    }

    let package = rusdox::validate_docx_package(&docx_path)?;
    let document = Document::open_read_only(&docx_path)?;
    let expected = DocumentProjection::from_document(&document);
    let docx = expected.clone();
    let mut config = load_runtime_config(args.config.as_deref())?;
    config.output.emit_pdf_preview = true;
    handle_validation_issues(
        "config",
        &validate_config(&config),
        true,
        "template PDF rendering aborted because the active config has validation errors",
    )?;
    let studio = Studio::new(config);
    let pdf = studio.render_pdf_with_evidence(&document, &pdf_path, Some(&page_snapshots_dir))?;
    let mut visual_diff = compare_visual_pages(&page_snapshots_dir, None, 0.0)?;
    for page in &mut visual_diff.pages {
        page.current = format!("reports/{output_name}-pages/page-{:03}.png", page.page);
    }
    let mut docx_artifact = ArtifactEvidence::from_path(&docx_path)?;
    docx_artifact.path = format!("generated/{output_name}.docx");
    let mut pdf_artifact = ArtifactEvidence::from_path(&pdf_path)?;
    pdf_artifact.path = format!("rendered/{output_name}.pdf");
    let report = ParityReport::compare(
        args.template.display().to_string(),
        expected,
        docx,
        pdf,
        visual_diff,
        vec![docx_artifact, pdf_artifact],
        package.valid,
        verify_pdf_file(&pdf_path)?,
    );
    rusdox::atomic_write_file(&json_report_path, report.to_json_pretty()?.as_bytes())?;
    let canonical =
        format!("https://othmaneblial.github.io/rusdox/templates/{output_name}-parity.html");
    rusdox::atomic_write_file(&html_report_path, report.to_html(&canonical).as_bytes())?;

    let failed_checks = report
        .checks
        .iter()
        .filter(|check| check.status == rusdox::parity::CheckStatus::Failed)
        .count();
    let result = TemplateCommandResult {
        command: command.to_string(),
        template: args.template.display().to_string(),
        data: args.data.display().to_string(),
        passed: report.passed,
        strict: args.strict,
        replacements: render.replacements,
        expanded_blocks: render.expanded_blocks,
        diagnostics: render.diagnostics,
        docx: docx_path.display().to_string(),
        pdf: pdf_path.display().to_string(),
        html_report: html_report_path.display().to_string(),
        json_report: json_report_path.display().to_string(),
        page_snapshots: page_snapshots_dir.display().to_string(),
        checks: report.checks.len(),
        failed_checks,
    };
    match args.format {
        ReportFormat::Json => println!(
            "{}",
            serde_json::to_string_pretty(&result).map_err(|error| {
                DocxError::Parse(format!("failed to serialize template result: {error}"))
            })?
        ),
        ReportFormat::Text => println!("{}", format_template_result_text(&result)),
    }
    if result.passed {
        Ok(())
    } else {
        Err(DocxError::Parity(format!(
            "template parity failed for {}",
            args.template.display()
        )))
    }
}

fn format_template_result_text(result: &TemplateCommandResult) -> String {
    [
        format!("template {}: {}", result.command, result.template),
        format!("data: {}", result.data),
        format!("strict: {}", result.strict),
        format!("replacements: {}", result.replacements),
        format!("expanded blocks: {}", result.expanded_blocks),
        format!("diagnostics: {}", result.diagnostics.len()),
        format!("docx: {}", result.docx),
        format!("pdf: {}", result.pdf),
        format!("parity html: {}", result.html_report),
        format!("parity json: {}", result.json_report),
        format!("page snapshots: {}", result.page_snapshots),
        format!(
            "checks: {} (failed: {})",
            result.checks, result.failed_checks
        ),
        format!("status: {}", if result.passed { "PASS" } else { "FAIL" }),
    ]
    .join("\n")
}

fn run_verify(args: VerifyArgs) -> Result<()> {
    if !args.visual_threshold.is_finite() || !(0.0..=1.0).contains(&args.visual_threshold) {
        return Err(DocxError::Parse(
            "--visual-threshold must be a finite value between 0 and 1".to_string(),
        ));
    }

    let spec_paths = collect_spec_inputs(&args.input)?;
    let input_is_directory = args.input.is_dir();
    let output_root = to_absolute_path(&args.output_root)?;
    let generated_dir = output_root.join("generated");
    let rendered_dir = output_root.join("rendered");
    let reports_dir = output_root.join("reports");
    fs::create_dir_all(&generated_dir)?;
    fs::create_dir_all(&rendered_dir)?;
    fs::create_dir_all(&reports_dir)?;

    let mut config = load_runtime_config(args.config.as_deref())?;
    config.output.emit_pdf_preview = true;
    handle_validation_issues(
        "config",
        &validate_config(&config),
        true,
        "verification aborted because the active config has validation errors",
    )?;

    let mut files = Vec::with_capacity(spec_paths.len());
    for spec_path in &spec_paths {
        let inspection = inspect_spec(spec_path)?;
        handle_validation_issues(
            &spec_path.display().to_string(),
            &inspection.report,
            true,
            "verification aborted because the spec has validation errors",
        )?;

        let output_name = inspection
            .spec
            .output_name
            .as_deref()
            .map(safe_output_name)
            .filter(|name| !name.is_empty())
            .unwrap_or_else(|| default_output_name_for_spec(spec_path));
        let docx_path = generated_dir.join(format!("{output_name}.docx"));
        let pdf_path = rendered_dir.join(format!("{output_name}.pdf"));
        let page_snapshots_dir = reports_dir.join(format!("{output_name}-pages"));
        let html_report_path = reports_dir.join(format!("{output_name}-parity.html"));
        let json_report_path = reports_dir.join(format!("{output_name}-parity.json"));

        let studio = Studio::new(config.clone());
        let document = studio.compose(&inspection.spec);
        let expected = DocumentProjection::from_document(&document);
        document.save(&docx_path)?;
        let pdf =
            studio.render_pdf_with_evidence(&document, &pdf_path, Some(&page_snapshots_dir))?;

        let reopened = rusdox::Document::open_read_only(&docx_path)?;
        let docx = DocumentProjection::from_document(&reopened);
        let visual_baseline = args.visual_baseline.as_deref().map(|root| {
            if input_is_directory {
                root.join(&output_name)
            } else {
                root.to_path_buf()
            }
        });
        let mut visual_diff = compare_visual_pages(
            &page_snapshots_dir,
            visual_baseline.as_deref(),
            args.visual_threshold,
        )?;
        for page in &mut visual_diff.pages {
            page.current = format!("reports/{output_name}-pages/page-{:03}.png", page.page);
        }
        let mut docx_artifact = ArtifactEvidence::from_path(&docx_path)?;
        docx_artifact.path = format!("generated/{output_name}.docx");
        let mut pdf_artifact = ArtifactEvidence::from_path(&pdf_path)?;
        pdf_artifact.path = format!("rendered/{output_name}.pdf");
        let artifacts = vec![docx_artifact, pdf_artifact];
        let report = ParityReport::compare(
            spec_path.display().to_string(),
            expected,
            docx,
            pdf,
            visual_diff,
            artifacts,
            verify_docx_package(&docx_path)?,
            verify_pdf_file(&pdf_path)?,
        );

        rusdox::atomic_write_file(&json_report_path, report.to_json_pretty()?.as_bytes())?;
        let canonical =
            format!("https://othmaneblial.github.io/rusdox/parity/{output_name}-parity.html");
        rusdox::atomic_write_file(&html_report_path, report.to_html(&canonical).as_bytes())?;

        let failed_checks = report
            .checks
            .iter()
            .filter(|check| check.status == rusdox::parity::CheckStatus::Failed)
            .count();
        files.push(VerifyFileResult {
            source: spec_path.display().to_string(),
            passed: report.passed,
            docx: docx_path.display().to_string(),
            pdf: pdf_path.display().to_string(),
            html_report: html_report_path.display().to_string(),
            json_report: json_report_path.display().to_string(),
            page_snapshots: page_snapshots_dir.display().to_string(),
            checks: report.checks.len(),
            failed_checks,
        });
    }

    let result = VerifyCommandResult {
        target: args.input.display().to_string(),
        passed: files.iter().all(|file| file.passed),
        specs: files.len(),
        files,
    };
    match args.format {
        ReportFormat::Text => println!("{}", format_verify_result_text(&result)),
        ReportFormat::Json => println!(
            "{}",
            serde_json::to_string_pretty(&result).map_err(|error| {
                DocxError::Parse(format!("failed to serialize verification summary: {error}"))
            })?
        ),
    }

    if result.passed {
        Ok(())
    } else {
        Err(DocxError::Parity(format!(
            "{} of {} generated document(s) failed; inspect the HTML/JSON reports",
            result.files.iter().filter(|file| !file.passed).count(),
            result.specs
        )))
    }
}

fn safe_output_name(value: &str) -> String {
    value
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() || matches!(character, '-' | '_') {
                character
            } else {
                '-'
            }
        })
        .collect::<String>()
        .trim_matches('-')
        .to_string()
}

fn verify_docx_package(path: &Path) -> Result<bool> {
    Ok(rusdox::validate_docx_package(path)?.valid)
}

fn verify_pdf_file(path: &Path) -> Result<bool> {
    let bytes = fs::read(path)?;
    let has_page = bytes
        .windows(b"/Type /Page".len())
        .any(|window| window == b"/Type /Page");
    Ok(bytes.starts_with(b"%PDF-")
        && bytes
            .windows(b"%%EOF".len())
            .any(|window| window == b"%%EOF")
        && has_page)
}

fn inspect_spec(spec_path: &Path) -> Result<SpecInspection> {
    let parse_start = Instant::now();
    let spec = DocumentSpec::load_from_path(spec_path)?;
    let parse_duration = parse_start.elapsed();

    let validate_start = Instant::now();
    let mut report = validate_spec(&spec);
    attach_source_spans(spec_path, &mut report)?;
    let validation_duration = validate_start.elapsed();

    Ok(SpecInspection {
        spec,
        parse_duration,
        validation_duration,
        report,
    })
}

fn build_spec_input(
    input: &Path,
    output: Option<&Path>,
    config: &RusdoxConfig,
    announce_outputs: bool,
    print_warnings: bool,
) -> Result<BuildSummary> {
    handle_validation_issues(
        "config",
        &validate_config(config),
        print_warnings,
        "rendering aborted because the active config has validation errors",
    )?;

    if input.is_dir() {
        if output.is_some() {
            return Err(DocxError::Parse(
                "--output is only supported for a single file input".to_string(),
            ));
        }

        let started = Instant::now();
        let spec_paths = collect_spec_inputs(input)?;
        let mut parse_duration = Duration::ZERO;
        let mut validation_duration = Duration::ZERO;
        let mut compose_duration = Duration::ZERO;
        let mut output_stats = OutputStats {
            docx_write: Duration::ZERO,
            pdf_render: Duration::ZERO,
            docx_bytes: 0,
            pdf_bytes: 0,
        };
        let mut warning_count = 0usize;

        for spec_path in &spec_paths {
            let summary =
                build_spec_file(spec_path, None, config, announce_outputs, print_warnings)?;
            parse_duration += summary.parse_duration;
            validation_duration += summary.validation_duration;
            compose_duration += summary.compose_duration;
            output_stats.docx_write += summary.output_stats.docx_write;
            output_stats.pdf_render += summary.output_stats.pdf_render;
            output_stats.docx_bytes += summary.output_stats.docx_bytes;
            output_stats.pdf_bytes += summary.output_stats.pdf_bytes;
            warning_count += summary.warning_count;
        }

        Ok(BuildSummary {
            documents: spec_paths.len(),
            parse_duration,
            validation_duration,
            compose_duration,
            output_stats,
            total_duration: started.elapsed(),
            warning_count,
        })
    } else {
        build_spec_file(input, output, config, announce_outputs, print_warnings)
    }
}

fn build_spec_file(
    spec_path: &Path,
    output: Option<&Path>,
    config: &RusdoxConfig,
    announce_outputs: bool,
    print_warnings: bool,
) -> Result<BuildSummary> {
    if !is_spec_path(spec_path) {
        return Err(DocxError::Parse(format!(
            "unsupported input type: {} (expected .yaml, .yml, .json, or .toml)",
            spec_path.display()
        )));
    }

    let started = Instant::now();
    let inspection = inspect_spec(spec_path)?;
    handle_validation_issues(
        &spec_path.display().to_string(),
        &inspection.report,
        print_warnings,
        "rendering aborted because the spec has validation errors",
    )?;

    let studio = Studio::new(config.clone());
    let compose_start = Instant::now();
    let document = studio.compose(&inspection.spec);
    let compose_duration = compose_start.elapsed();

    let output_stats = if let Some(output_path) = output {
        let output_path = to_absolute_path(output_path)?;
        if let Some(parent) = output_path.parent() {
            fs::create_dir_all(parent)?;
        }
        if announce_outputs {
            studio.save_with_pdf_stats(&document, &output_path)?
        } else {
            studio.save_with_pdf_stats_quiet(&document, &output_path)?
        }
    } else {
        let output_name = inspection
            .spec
            .output_name
            .clone()
            .unwrap_or_else(|| default_output_name_for_spec(spec_path));
        if announce_outputs {
            studio.save_named(&document, &output_name)?
        } else {
            studio.save_named_quiet(&document, &output_name)?
        }
    };

    Ok(BuildSummary {
        documents: 1,
        parse_duration: inspection.parse_duration,
        validation_duration: inspection.validation_duration,
        compose_duration,
        output_stats,
        total_duration: started.elapsed(),
        warning_count: inspection.report.warning_count(),
    })
}

fn bench_once(
    inputs: &[PathBuf],
    config: &RusdoxConfig,
    pipeline: BenchPipeline,
    print_warnings: bool,
) -> Result<BenchSample> {
    let started = Instant::now();
    let mut parse_duration = Duration::ZERO;
    let mut validation_duration = Duration::ZERO;
    let mut compose_duration = Duration::ZERO;
    let mut docx_duration = Duration::ZERO;
    let mut pdf_duration = Duration::ZERO;
    let mut docx_bytes = 0_u64;
    let mut pdf_bytes = 0_u64;
    let mut existing_docx_open_duration = Duration::ZERO;
    let mut existing_docx_save_duration = Duration::ZERO;

    for input in inputs {
        if pipeline == BenchPipeline::ExistingDocx {
            let open_start = Instant::now();
            let document = Document::open(input)?;
            existing_docx_open_duration += open_start.elapsed();

            let output_path = Path::new(&config.output.docx_dir).join(format!(
                "{}-roundtrip.docx",
                input
                    .file_stem()
                    .and_then(|stem| stem.to_str())
                    .unwrap_or("existing")
            ));
            let save_start = Instant::now();
            document.save(&output_path)?;
            existing_docx_save_duration += save_start.elapsed();
            docx_bytes += fs::metadata(output_path)?.len();
            continue;
        }

        let inspection = inspect_spec(input)?;
        parse_duration += inspection.parse_duration;
        validation_duration += inspection.validation_duration;
        handle_validation_issues(
            &input.display().to_string(),
            &inspection.report,
            print_warnings,
            "benchmark target has validation errors",
        )?;

        if pipeline == BenchPipeline::Validation {
            continue;
        }

        let studio = Studio::new(config.clone());
        let compose_start = Instant::now();
        let document = studio.compose(&inspection.spec);
        compose_duration += compose_start.elapsed();
        let output_name = inspection
            .spec
            .output_name
            .clone()
            .unwrap_or_else(|| default_output_name_for_spec(input));

        match pipeline {
            BenchPipeline::Docx => {
                let docx_path =
                    Path::new(&config.output.docx_dir).join(format!("{output_name}.docx"));
                let write_start = Instant::now();
                document.save(&docx_path)?;
                docx_duration += write_start.elapsed();
                docx_bytes += fs::metadata(docx_path)?.len();
            }
            BenchPipeline::Pdf => {
                let pdf_path = Path::new(&config.output.pdf_dir).join(format!("{output_name}.pdf"));
                let render_start = Instant::now();
                let _evidence = studio.render_pdf_with_evidence(&document, &pdf_path, None)?;
                pdf_duration += render_start.elapsed();
                pdf_bytes += fs::metadata(pdf_path)?.len();
            }
            BenchPipeline::Dual => {
                let docx_path =
                    Path::new(&config.output.docx_dir).join(format!("{output_name}.docx"));
                let stats = studio.save_with_pdf_stats_quiet(&document, docx_path)?;
                docx_duration += stats.docx_write;
                pdf_duration += stats.pdf_render;
                docx_bytes += stats.docx_bytes;
                pdf_bytes += stats.pdf_bytes;
            }
            BenchPipeline::Validation | BenchPipeline::ExistingDocx => unreachable!(),
        }
    }

    Ok(BenchSample {
        parse_duration,
        validation_duration,
        compose_duration,
        docx_duration,
        pdf_duration,
        total_duration: started.elapsed(),
        docx_bytes,
        pdf_bytes,
        existing_docx_open_duration,
        existing_docx_save_duration,
    })
}

fn input_fingerprint(input: &Path, inputs: &[PathBuf]) -> Result<(String, u64)> {
    if !input.is_dir() && inputs.len() == 1 {
        let bytes = fs::read(&inputs[0])?;
        return Ok((format!("{:x}", Sha256::digest(&bytes)), bytes.len() as u64));
    }

    let mut digest = Sha256::new();
    let mut total_bytes = 0_u64;

    for path in inputs {
        let relative = if input.is_dir() {
            path.strip_prefix(input).unwrap_or(path.as_path())
        } else {
            path.file_name().map(Path::new).unwrap_or(path.as_path())
        };
        let bytes = fs::read(path)?;
        let relative = relative.to_string_lossy();
        digest.update((relative.len() as u64).to_le_bytes());
        digest.update(relative.as_bytes());
        digest.update((bytes.len() as u64).to_le_bytes());
        digest.update(&bytes);
        total_bytes += bytes.len() as u64;
    }

    Ok((format!("{:x}", digest.finalize()), total_bytes))
}

fn collect_spec_inputs(input: &Path) -> Result<Vec<PathBuf>> {
    if !input.exists() {
        return Err(DocxError::Parse(format!(
            "input not found: {}",
            input.display()
        )));
    }

    if input.is_dir() {
        let mut entries = fs::read_dir(input)?
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .filter(|path| path.is_file() && is_spec_path(path))
            .collect::<Vec<_>>();
        entries.sort();
        if entries.is_empty() {
            return Err(DocxError::Parse(format!(
                "no document spec files found in {}",
                input.display()
            )));
        }
        Ok(entries)
    } else if is_spec_path(input) {
        Ok(vec![input.to_path_buf()])
    } else {
        Err(DocxError::Parse(format!(
            "unsupported input type: {} (expected .yaml, .yml, .json, .toml, or a directory)",
            input.display()
        )))
    }
}

fn handle_validation_issues(
    label: &str,
    report: &ValidationReport,
    print_warnings: bool,
    error_prefix: &str,
) -> Result<()> {
    if report.has_warnings() && print_warnings {
        eprintln!("{}", format_issue_list(label, report, true));
    }
    if report.has_errors() {
        return Err(DocxError::Parse(format!(
            "{error_prefix}\n{}",
            format_issue_list(label, report, false)
        )));
    }
    Ok(())
}

fn capture_watch_snapshot(input: &Path, config_path: Option<&Path>) -> Result<WatchSnapshot> {
    let spec_inputs = collect_spec_inputs(input)?;
    let mut watched_paths = spec_inputs.clone();
    let mut visited = BTreeSet::new();
    for spec_path in spec_inputs {
        collect_path_dependencies(&spec_path, &mut watched_paths, &mut visited, 0)?;
    }
    if let Some(config_path) = config_path {
        watched_paths.push(config_path.to_path_buf());
    } else {
        watched_paths.push(PathBuf::from(DEFAULT_CONFIG_FILE));
        if let Some(user_path) = default_user_config_path() {
            watched_paths.push(user_path);
        }
    }
    watched_paths.sort();
    watched_paths.dedup();

    let mut states = Vec::with_capacity(watched_paths.len());
    for path in watched_paths {
        states.push((path.clone(), hash_path_state(&path)?));
    }
    Ok(WatchSnapshot { states })
}

fn collect_path_dependencies(
    source_path: &Path,
    watched_paths: &mut Vec<PathBuf>,
    visited: &mut BTreeSet<PathBuf>,
    depth: usize,
) -> Result<()> {
    if depth >= 16 || !source_path.is_file() || !visited.insert(source_path.to_path_buf()) {
        return Ok(());
    }
    let source = fs::read_to_string(source_path)?;
    let parent = source_path.parent().unwrap_or_else(|| Path::new("."));
    for value in source.lines().filter_map(extract_path_value) {
        if value.contains("://") || value.starts_with("data:") {
            continue;
        }
        let path = PathBuf::from(value);
        let path = if path.is_absolute() {
            path
        } else {
            parent.join(path)
        };
        if !watched_paths.contains(&path) {
            watched_paths.push(path.clone());
        }
        if is_spec_path(&path) {
            collect_path_dependencies(&path, watched_paths, visited, depth + 1)?;
        }
    }
    Ok(())
}

fn extract_path_value(line: &str) -> Option<String> {
    let trimmed = line.trim();
    let value = if let Some(value) = trimmed.strip_prefix("path:") {
        value
    } else if let Some(value) = trimmed.strip_prefix("path =") {
        value
    } else {
        trimmed.strip_prefix("\"path\":")?
    };
    let value = value
        .trim()
        .trim_end_matches(',')
        .trim()
        .trim_matches('"')
        .trim_matches('\'')
        .trim();
    (!value.is_empty()).then(|| value.to_string())
}

fn hash_path_state(path: &Path) -> Result<u64> {
    let mut hasher = DefaultHasher::new();
    path.hash(&mut hasher);
    if path.exists() {
        let bytes = fs::read(path)?;
        1_u8.hash(&mut hasher);
        bytes.hash(&mut hasher);
    } else {
        0_u8.hash(&mut hasher);
    }
    Ok(hasher.finish())
}

fn changed_paths(previous: &WatchSnapshot, next: &WatchSnapshot) -> Vec<PathBuf> {
    let previous = previous
        .states
        .iter()
        .cloned()
        .collect::<std::collections::BTreeMap<_, _>>();
    let next = next
        .states
        .iter()
        .cloned()
        .collect::<std::collections::BTreeMap<_, _>>();
    let mut changed = Vec::new();

    for path in previous.keys().chain(next.keys()) {
        if previous.get(path) != next.get(path) && !changed.iter().any(|item| item == path) {
            changed.push(path.clone());
        }
    }

    changed
}

fn wait_for_watch_change(
    input: &Path,
    config_path: Option<&Path>,
    previous: &WatchSnapshot,
    poll_interval: Duration,
    debounce: Duration,
) -> Result<(WatchSnapshot, Vec<PathBuf>)> {
    let mut candidate = loop {
        thread::sleep(poll_interval);
        let next = capture_watch_snapshot(input, config_path)?;
        if &next != previous {
            break next;
        }
    };
    let mut changed = changed_paths(previous, &candidate);
    let mut stable_since = Instant::now();

    while stable_since.elapsed() < debounce {
        let remaining = debounce.saturating_sub(stable_since.elapsed());
        thread::sleep(poll_interval.min(remaining).max(Duration::from_millis(1)));
        let next = capture_watch_snapshot(input, config_path)?;
        if next != candidate {
            for path in changed_paths(&candidate, &next) {
                if !changed.contains(&path) {
                    changed.push(path);
                }
            }
            candidate = next;
            stable_since = Instant::now();
        }
    }
    changed.sort();
    Ok((candidate, changed))
}

fn format_validation_result_text(result: &ValidationCommandResult) -> String {
    let mut lines = vec![format!(
        "validated {} spec(s) under {}: {} error(s), {} warning(s)",
        result.specs, result.target, result.errors, result.warnings
    )];

    if !result.config_issues.is_empty() {
        lines.push("config:".to_string());
        for issue in &result.config_issues {
            lines.push(format_issue_line(issue));
        }
    }

    for file in &result.files {
        lines.push(format!("{}:", file.path));
        if file.issues.is_empty() {
            lines.push("  [ok] no semantic issues".to_string());
        } else {
            for issue in &file.issues {
                lines.push(format_issue_line(issue));
            }
        }
    }

    lines.join("\n")
}

fn format_issue_list(label: &str, report: &ValidationReport, warnings_only: bool) -> String {
    let mut lines = vec![format!("{label}:")];
    let issues = report
        .issues
        .iter()
        .filter(|issue| !warnings_only || issue.severity == ValidationSeverity::Warning)
        .collect::<Vec<_>>();
    for issue in issues {
        lines.push(format_issue_line(issue));
    }
    lines.join("\n")
}

fn format_issue_line(issue: &ValidationIssue) -> String {
    let severity = match issue.severity {
        ValidationSeverity::Error => "error",
        ValidationSeverity::Warning => "warning",
    };
    let location = issue.source.map_or_else(String::new, |source| {
        format!(" (line {}, column {})", source.line, source.column)
    });
    format!("  [{severity}] {}{location}: {}", issue.path, issue.message)
}

fn summarize_f64(values: impl Iterator<Item = f64>) -> NumericSummary {
    let mut collected = values.collect::<Vec<_>>();
    collected.sort_by(f64::total_cmp);
    let count = collected.len().max(1) as f64;
    let sum = collected.iter().sum::<f64>();
    let min = collected.iter().copied().fold(f64::INFINITY, f64::min);
    let max = collected.iter().copied().fold(f64::NEG_INFINITY, f64::max);
    NumericSummary {
        avg: sum / count,
        min: if collected.is_empty() { 0.0 } else { min },
        median: if collected.is_empty() {
            0.0
        } else if collected.len() % 2 == 0 {
            let upper = collected.len() / 2;
            (collected[upper - 1] + collected[upper]) / 2.0
        } else {
            collected[collected.len() / 2]
        },
        max: if collected.is_empty() { 0.0 } else { max },
    }
}

fn format_bench_result_text(result: &BenchCommandResult) -> String {
    [
        format!("benchmark target: {}", result.target),
        format!("pipeline: {}", result.pipeline),
        format!("input sha256: {}", result.input_sha256),
        format!("input bytes: {}", result.input_bytes),
        format!("specs: {}", result.specs),
        format!("iterations: {}", result.iterations),
        format!("warmup: {}", result.warmup),
        format!("pdf enabled: {}", result.emit_pdf),
        format!("keep output: {}", result.keep_output),
        format!(
            "parse: {} avg, {} min, {} median, {} max",
            format_ms(result.parse_ms.avg),
            format_ms(result.parse_ms.min),
            format_ms(result.parse_ms.median),
            format_ms(result.parse_ms.max)
        ),
        format!(
            "validate: {} avg, {} min, {} median, {} max",
            format_ms(result.validate_ms.avg),
            format_ms(result.validate_ms.min),
            format_ms(result.validate_ms.median),
            format_ms(result.validate_ms.max)
        ),
        format!(
            "compose: {} avg, {} min, {} median, {} max",
            format_ms(result.compose_ms.avg),
            format_ms(result.compose_ms.min),
            format_ms(result.compose_ms.median),
            format_ms(result.compose_ms.max)
        ),
        format!(
            "docx write: {} avg, {} min, {} median, {} max",
            format_ms(result.docx_ms.avg),
            format_ms(result.docx_ms.min),
            format_ms(result.docx_ms.median),
            format_ms(result.docx_ms.max)
        ),
        format!(
            "pdf render: {} avg, {} min, {} median, {} max",
            format_ms(result.pdf_ms.avg),
            format_ms(result.pdf_ms.min),
            format_ms(result.pdf_ms.median),
            format_ms(result.pdf_ms.max)
        ),
        format!(
            "existing DOCX open: {} avg, {} min, {} median, {} max",
            format_ms(result.existing_docx_open_ms.avg),
            format_ms(result.existing_docx_open_ms.min),
            format_ms(result.existing_docx_open_ms.median),
            format_ms(result.existing_docx_open_ms.max)
        ),
        format!(
            "existing DOCX save: {} avg, {} min, {} median, {} max",
            format_ms(result.existing_docx_save_ms.avg),
            format_ms(result.existing_docx_save_ms.min),
            format_ms(result.existing_docx_save_ms.median),
            format_ms(result.existing_docx_save_ms.max)
        ),
        format!(
            "total: {} avg, {} min, {} median, {} max",
            format_ms(result.total_ms.avg),
            format_ms(result.total_ms.min),
            format_ms(result.total_ms.median),
            format_ms(result.total_ms.max)
        ),
        format!(
            "docx bytes: {} avg, {} min, {} max",
            format_bytes(result.docx_bytes.avg.round() as u64),
            format_bytes(result.docx_bytes.min.round() as u64),
            format_bytes(result.docx_bytes.max.round() as u64)
        ),
        format!(
            "pdf bytes: {} avg, {} min, {} max",
            format_bytes(result.pdf_bytes.avg.round() as u64),
            format_bytes(result.pdf_bytes.min.round() as u64),
            format_bytes(result.pdf_bytes.max.round() as u64)
        ),
    ]
    .join("\n")
}

fn format_verify_result_text(result: &VerifyCommandResult) -> String {
    let mut lines = vec![format!(
        "parity verification {}: {} spec(s) under {}",
        if result.passed { "passed" } else { "failed" },
        result.specs,
        result.target
    )];
    for file in &result.files {
        lines.push(format!(
            "  [{}] {} ({} checks, {} failed)",
            if file.passed { "pass" } else { "fail" },
            file.source,
            file.checks,
            file.failed_checks
        ));
        lines.push(format!("    DOCX: {}", file.docx));
        lines.push(format!("    PDF: {}", file.pdf));
        lines.push(format!("    HTML: {}", file.html_report));
        lines.push(format!("    JSON: {}", file.json_report));
        lines.push(format!("    pages: {}", file.page_snapshots));
    }
    lines.join("\n")
}

fn duration_ms(duration: Duration) -> f64 {
    duration.as_secs_f64() * 1000.0
}

fn format_duration(duration: Duration) -> String {
    format_ms(duration_ms(duration))
}

fn format_ms(value: f64) -> String {
    format!("{value:.2} ms")
}

fn format_bytes(bytes: u64) -> String {
    let kib = 1024.0;
    let mib = kib * 1024.0;
    let gib = mib * 1024.0;
    let value = bytes as f64;

    if value >= gib {
        format!("{:.2} GiB", value / gib)
    } else if value >= mib {
        format!("{:.2} MiB", value / mib)
    } else if value >= kib {
        format!("{:.2} KiB", value / kib)
    } else {
        format!("{bytes} B")
    }
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
    let config = runtime_config(config_path.as_deref(), docx_only, with_pdf)?;
    let _summary = build_spec_input(&spec_path, output.as_deref(), &config, true, true)?;
    Ok(())
}

fn run_spec_dir(
    dir: PathBuf,
    config_path: Option<PathBuf>,
    docx_only: bool,
    with_pdf: bool,
) -> Result<()> {
    let config = runtime_config(config_path.as_deref(), docx_only, with_pdf)?;
    let _summary = build_spec_input(&dir, None, &config, true, true)?;
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

    let runner_dir = cached_script_runner_dir(&script_path);
    let runner_package_name = runner_dir
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("rusdox-script-runner");
    let manifest_path = runner_dir.join("Cargo.toml");
    let src_dir = runner_dir.join("src");
    fs::create_dir_all(&src_dir)?;

    let runner_config_path = runner_dir.join("rusdox-runtime-config.toml");
    config.save_to_path(&runner_config_path)?;

    fs::write(&manifest_path, build_runner_manifest(runner_package_name))?;
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
    command.current_dir(&runner_dir);
    if std::env::var_os("CARGO_TARGET_DIR").is_none() {
        command.env(
            "CARGO_TARGET_DIR",
            std::env::temp_dir().join("rusdox-script-runner-target"),
        );
    }

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

fn cached_script_runner_dir(script_path: &Path) -> PathBuf {
    let mut hasher = DefaultHasher::new();
    script_path.hash(&mut hasher);
    std::env::temp_dir().join(format!("rusdox-script-runner-{:016x}", hasher.finish()))
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

fn build_runner_manifest(package_name: &str) -> String {
    format!(
        r#"[package]
name = "{package_name}"
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
    let mut spec = DocumentSpec::new();
    spec.output_name = Some("my-document".to_string());
    spec.blocks = vec![
        rusdox::spec::title("My Document"),
        rusdox::spec::subtitle("Written as data, rendered by Rust"),
        rusdox::spec::section("Summary"),
        rusdox::spec::body("Replace this with your real content."),
        rusdox::spec::bullets([
            "Keep content in order.",
            "Let config handle styling.",
            "Render to DOCX and PDF with one command.",
        ]),
    ];
    spec
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
    use std::fs;
    use std::io::{Read, Write};
    use std::net::TcpStream;
    use std::path::Path;

    use tempfile::tempdir;

    use super::{
        build_runner_manifest, build_runner_source, capture_watch_snapshot,
        default_script_template, dev_change_reason, escape_rust_string, is_loopback_url,
        normalize_color_hex, resolve_path, start_dev_server, summarize_f64, ConfigFormat,
    };

    #[test]
    fn numeric_summary_reports_real_extrema_and_median() {
        let summary = summarize_f64([9.0, 1.0, 5.0, 3.0].into_iter());
        assert_eq!(summary.avg, 4.5);
        assert_eq!(summary.min, 1.0);
        assert_eq!(summary.median, 4.0);
        assert_eq!(summary.max, 9.0);
    }

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
    fn registry_http_allows_only_exact_loopback_hosts() {
        assert!(is_loopback_url("http://127.0.0.1:8080/index.json"));
        assert!(is_loopback_url("http://localhost/index.json"));
        assert!(is_loopback_url("http://[::1]:8080/index.json"));
        assert!(!is_loopback_url("http://127.0.0.1.example/index.json"));
        assert!(!is_loopback_url("http://localhost.example/index.json"));
        assert!(!is_loopback_url("https://localhost/index.json"));
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
        let manifest = build_runner_manifest("rusdox-script-runner-test");
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

    #[test]
    fn watch_snapshot_tracks_local_asset_dependencies() {
        let temp = tempdir().expect("temp dir");
        let spec = temp.path().join("document.yaml");
        let asset = temp.path().join("chart.svg");
        fs::write(&asset, "<svg/>").expect("asset");
        fs::write(
            &spec,
            "version: 1\nblocks:\n  - type: chart\n    path: chart.svg\n",
        )
        .expect("spec");
        let before = capture_watch_snapshot(&spec, None).expect("before");
        fs::write(&asset, "<svg><path/></svg>").expect("change asset");
        let after = capture_watch_snapshot(&spec, None).expect("after");
        let changed = super::changed_paths(&before, &after);
        assert_eq!(changed, vec![asset.clone()]);
        assert!(dev_change_reason(&changed, &spec, None).contains("asset/include changed"));
    }

    #[test]
    fn dev_server_exposes_script_free_status_dashboard() {
        let server = start_dev_server(0).expect("server");
        let address = server
            .url
            .trim_start_matches("http://")
            .trim_end_matches('/');
        let mut stream = TcpStream::connect(address).expect("connect");
        stream
            .write_all(b"GET /status.json HTTP/1.1\r\nHost: localhost\r\n\r\n")
            .expect("request");
        let mut response = String::new();
        stream.read_to_string(&mut response).expect("response");
        assert!(response.contains("200 OK"));
        assert!(response.contains("Content-Security-Policy: default-src 'none'"));
        assert!(response.contains("\"status\": \"starting\""));
    }
}
