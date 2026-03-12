use std::fs;

use rusdox::spec::{
    body, bullets, col, hero, label_values, page_heading, paragraph, row, section, status,
    subtitle, table, title, BlockSpec, DocumentSpec, ParagraphAlignmentSpec, ParagraphSpec,
    RowSpec, RunSpec, Tone,
};

const TOTAL_PAGES: usize = 1000;
const UNIQUE_TEMPLATES: usize = 10;
const OUTPUT_PATH: &str = "examples/stress/stress_1000_pages.yaml";

#[derive(Clone, Copy)]
struct PageTemplate {
    title: &'static str,
    subtitle: &'static str,
    highlight: &'static str,
    summary: &'static str,
    bullets: [&'static str; 3],
    rows: [MetricRow; 3],
}

#[derive(Clone, Copy)]
struct MetricRow {
    label: &'static str,
    value: &'static str,
    status: &'static str,
    tone: Tone,
}

fn main() -> Result<(), rusdox::DocxError> {
    let spec = build_spec();
    spec.save_to_path(OUTPUT_PATH)?;
    let bytes = fs::metadata(OUTPUT_PATH)?.len();

    println!("{OUTPUT_PATH}");
    println!("logical pages: {TOTAL_PAGES}");
    println!("unique templates: {UNIQUE_TEMPLATES}");
    println!("yaml size: {}", format_bytes(bytes));
    Ok(())
}

fn build_spec() -> DocumentSpec {
    let templates = build_templates();
    let mut blocks = Vec::with_capacity(TOTAL_PAGES * 9);

    for page_index in 0..TOTAL_PAGES {
        let template = &templates[page_index % templates.len()];
        append_page(&mut blocks, template, page_index);
    }

    DocumentSpec {
        output_name: Some("stress-1000-pages".to_string()),
        blocks,
    }
}

fn build_templates() -> Vec<PageTemplate> {
    vec![
        template(
            "Executive Pulse",
            "Revenue, delivery, and retention health",
            "Expansion revenue is offsetting slower top-of-funnel conversion.",
            "The operating cadence remains stable under rising enterprise volume, with support load and launch readiness staying within target range.",
            [
                "Enterprise expansion grew faster than plan.",
                "Support backlog remained below the weekly ceiling.",
                "Launch blockers are concentrated in security review.",
            ],
            [
                metric_row("Revenue", "$4.8M", "On Track", Tone::Positive),
                metric_row("Gross Margin", "68%", "Watch", Tone::Warning),
                metric_row("Net Retention", "117%", "On Track", Tone::Positive),
            ],
        ),
        template(
            "Board Snapshot",
            "Capital allocation and operating leverage",
            "Capital is being pushed into areas with immediate commercial leverage.",
            "Leadership is holding the current burn posture while prioritizing product releases with visible expansion pull-through.",
            [
                "Security work reduces late-stage deal risk.",
                "Customer operations still needs leadership backfill.",
                "Margin profile improved after infrastructure changes.",
            ],
            [
                metric_row("ARR", "$18.7M", "On Track", Tone::Positive),
                metric_row("Cash Runway", "24 mo", "On Track", Tone::Positive),
                metric_row("Hiring Delay", "1 role", "Risk", Tone::Risk),
            ],
        ),
        template(
            "Launch Readiness",
            "Cross-functional prelaunch brief",
            "Commercial launch can move once documentation and FAQ sign-off are complete.",
            "The GTM plan is fully sequenced, and the main remaining dependency is final wording approval in the security pack.",
            [
                "Pilot feedback is already feeding into enablement.",
                "Support macros are drafted and being refined.",
                "Pricing and packaging have been signed off.",
            ],
            [
                metric_row("Pricing", "Approved", "On Track", Tone::Positive),
                metric_row("Enablement", "In progress", "Watch", Tone::Warning),
                metric_row("Security FAQ", "Pending", "Risk", Tone::Risk),
            ],
        ),
        template(
            "Client Proposal",
            "Delivery plan and commercial structure",
            "The phased rollout keeps implementation risk low while preserving speed.",
            "The proposal emphasizes structured delivery, fixed implementation scope, and an optional managed support layer for teams that want ongoing help.",
            [
                "Week one is discovery and document mapping.",
                "Weeks two to four are generator implementation.",
                "Final week is UAT and rollout support.",
            ],
            [
                metric_row("Implementation", "$24,000", "Approved", Tone::Positive),
                metric_row("Managed Support", "$3,000/mo", "Optional", Tone::Warning),
                metric_row("Decision Window", "This week", "Watch", Tone::Warning),
            ],
        ),
        template(
            "Talent Profile",
            "Leadership experience and operating depth",
            "The candidate combines executive reporting depth with launch operations rigor.",
            "This profile format is meant to show that long-form career narratives and structured capability tables can be generated at scale as well.",
            [
                "Built weekly operating packs for leadership.",
                "Ran launch reviews across product and support.",
                "Automated recurring reporting workflows.",
            ],
            [
                metric_row("Executive Reporting", "Expert", "Strong", Tone::Positive),
                metric_row("Launch Ops", "Expert", "Strong", Tone::Positive),
                metric_row("Systems Design", "Advanced", "Strong", Tone::Positive),
            ],
        ),
        template(
            "Program Review",
            "Milestones, delivery status, and owners",
            "Program momentum is good, but vendor response time still creates schedule drag.",
            "The delivery team is tracking toward the quarter target with most schedule pressure isolated to external approvals rather than internal execution.",
            [
                "Two milestones closed ahead of schedule.",
                "Vendor response time remains the main drag.",
                "No critical regressions were found this week.",
            ],
            [
                metric_row("Migration", "82%", "On Track", Tone::Positive),
                metric_row("Vendor Approval", "Delayed", "Watch", Tone::Warning),
                metric_row("Regression Risk", "Low", "On Track", Tone::Positive),
            ],
        ),
        template(
            "Operations Review",
            "SLA, backlog, and team throughput",
            "Backlog remains manageable despite a short-term spike in incoming volume.",
            "Staffing adjustments and workflow automation kept response time inside target even after a large customer migration weekend.",
            [
                "Average handling time improved again.",
                "Escalations stayed below threshold.",
                "Internal tooling removed repetitive triage steps.",
            ],
            [
                metric_row("First Response SLA", "96%", "On Track", Tone::Positive),
                metric_row("Backlog", "142", "Watch", Tone::Warning),
                metric_row("Escalation Rate", "2.1%", "On Track", Tone::Positive),
            ],
        ),
        template(
            "Portfolio Summary",
            "Cross-initiative scorecard",
            "The overall portfolio is healthy, with risk concentrated in one external dependency chain.",
            "The team is intentionally narrowing scope in lower-signal workstreams to preserve capacity for the initiatives with the clearest commercial return.",
            [
                "Three workstreams are cleanly on plan.",
                "One workstream needs executive escalation.",
                "Resourcing is being shifted rather than increased.",
            ],
            [
                metric_row("Core Product", "Green", "On Track", Tone::Positive),
                metric_row("Compliance Stream", "Amber", "Watch", Tone::Warning),
                metric_row("Partner Integration", "Red", "Risk", Tone::Risk),
            ],
        ),
        template(
            "Customer Health",
            "Retention signals and expansion opportunity",
            "The install base is stable and the top cohort is expanding earlier than forecast.",
            "The operating question for the next period is not churn defense but whether the onboarding path can be simplified enough to increase early activation.",
            [
                "Retention held above plan.",
                "Expansion signals are concentrated in enterprise.",
                "Activation remains the biggest growth lever.",
            ],
            [
                metric_row("Gross Retention", "94%", "On Track", Tone::Positive),
                metric_row("Expansion Leads", "41", "On Track", Tone::Positive),
                metric_row("Activation Rate", "62%", "Watch", Tone::Warning),
            ],
        ),
        template(
            "Automation Brief",
            "Document generation system stress profile",
            "The system is producing large outputs with stable structure and repeatable timing.",
            "This template exists specifically to pressure the document pipeline with repeated styled pages, structured tables, and predictable output patterns.",
            [
                "Templates are intentionally reused for scale.",
                "Each logical page begins on a new page.",
                "The output includes both DOCX and PDF artifacts.",
            ],
            [
                metric_row("Logical Pages", "1000", "On Track", Tone::Positive),
                metric_row("Templates", "10", "On Track", Tone::Positive),
                metric_row("Pipeline", "Pure Rust", "On Track", Tone::Positive),
            ],
        ),
    ]
}

fn template(
    title_text: &'static str,
    subtitle_text: &'static str,
    highlight: &'static str,
    summary: &'static str,
    bullets_text: [&'static str; 3],
    rows: [MetricRow; 3],
) -> PageTemplate {
    PageTemplate {
        title: title_text,
        subtitle: subtitle_text,
        highlight,
        summary,
        bullets: bullets_text,
        rows,
    }
}

fn metric_row(
    label: &'static str,
    value: &'static str,
    status_text: &'static str,
    tone: Tone,
) -> MetricRow {
    MetricRow {
        label,
        value,
        status: status_text,
        tone,
    }
}

fn append_page(blocks: &mut Vec<BlockSpec>, template: &PageTemplate, page_index: usize) {
    let page_number = page_index + 1;
    let page_subtitle = format!(
        "{}  |  Page {} of {}",
        template.subtitle, page_number, TOTAL_PAGES
    );
    let repeated_template_note = format!(
        "Stress template {} repeated at page {}",
        (page_index % UNIQUE_TEMPLATES) + 1,
        page_number
    );

    if page_index == 0 {
        blocks.push(title(template.title));
    } else {
        blocks.push(page_heading(template.title));
    }
    blocks.push(subtitle(page_subtitle));
    blocks.push(hero(template.highlight));
    blocks.push(body(template.summary));
    blocks.push(section("Key Takeaways"));
    blocks.push(bullets(template.bullets));
    blocks.push(section("Structured Snapshot"));
    blocks.push(summary_table_block(template.rows));
    blocks.push(section("Operating Notes"));
    blocks.push(label_values([
        (
            "Cadence",
            "Weekly executive review with daily owner-level checkpoints",
        ),
        (
            "Confidence",
            "High confidence in the repeatability of the document generation path",
        ),
    ]));
    blocks.push(paragraph(ParagraphSpec {
        runs: vec![RunSpec {
            text: repeated_template_note,
            italic: true,
            color: Some("B45309".to_string()),
            size_pt: Some(9.0),
            ..RunSpec::default()
        }],
        alignment: Some(ParagraphAlignmentSpec::Center),
        spacing_before_twips: Some(60),
        spacing_after_twips: Some(0),
        ..ParagraphSpec::default()
    }));
}

fn summary_table_block(rows: [MetricRow; 3]) -> BlockSpec {
    table(
        [
            col("Dimension", 3_200),
            col("Current", 2_000),
            col("Status", 1_760),
            col("Comment", 2_400),
        ],
        [
            summary_row(rows[0]),
            summary_row(rows[1]),
            summary_row(rows[2]),
        ],
    )
}

fn summary_row(row_data: MetricRow) -> RowSpec {
    row((
        row_data.label,
        row_data.value,
        status(row_data.status, row_data.tone),
        commentary_for(row_data.label, row_data.status),
    ))
}

fn commentary_for(label: &str, status_text: &str) -> &'static str {
    match (label, status_text) {
        ("Hiring Delay", _) => "Leadership backfill remains the only people risk on the page.",
        (_, "Risk") => "This area needs escalation before the next operating review.",
        (_, "Watch") => "This area is trending within tolerance but still needs attention.",
        _ => "This area is within the expected operating range.",
    }
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
