use std::fs;
use std::path::Path;
use std::time::{Duration, Instant};

use rusdox::spec::DocumentSpec;
use rusdox::studio::Studio;

const TOTAL_PAGES: usize = 1000;
const UNIQUE_TEMPLATES: usize = 10;
const SPEC_PATH: &str = "examples/stress/stress_1000_pages.yaml";

fn main() -> Result<(), rusdox::DocxError> {
    let spec_path = Path::new(SPEC_PATH);
    if !spec_path.exists() {
        return Err(rusdox::DocxError::Parse(format!(
            "stress spec not found at {} (run scripts/generate_stress_yaml.sh first)",
            spec_path.display()
        )));
    }

    let studio = Studio::from_default_file_or_default()?;
    let overall_start = Instant::now();

    let parse_start = Instant::now();
    let spec = DocumentSpec::load_from_path(spec_path)?;
    let parse_duration = parse_start.elapsed();

    let compose_start = Instant::now();
    let document = studio.compose(&spec);
    let compose_duration = compose_start.elapsed();

    let output_name = spec.output_name.as_deref().unwrap_or("stress-1000-pages");
    let stats = studio.save_named(&document, output_name)?;
    let total_duration = overall_start.elapsed();
    let yaml_bytes = fs::metadata(spec_path)?.len();

    println!("stress spec: {}", spec_path.display());
    println!("logical pages: {TOTAL_PAGES}");
    println!("unique templates: {UNIQUE_TEMPLATES}");
    println!("yaml parse: {}", format_duration(parse_duration));
    println!("doc compose: {}", format_duration(compose_duration));
    println!("docx write: {}", format_duration(stats.docx_write));
    println!("pdf render: {}", format_duration(stats.pdf_render));
    println!("output total: {}", format_duration(total_duration));
    println!("yaml size: {}", format_bytes(yaml_bytes));
    println!("docx size: {}", format_bytes(stats.docx_bytes));
    println!("pdf size: {}", format_bytes(stats.pdf_bytes));

    Ok(())
}

fn format_duration(duration: Duration) -> String {
    let milliseconds = duration.as_secs_f64() * 1000.0;
    format!("{milliseconds:.2} ms")
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
