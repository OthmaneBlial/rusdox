const DOCS = [
  {
    title: "Getting started",
    summary: "Install the CLI and generate an editable DOCX plus a native PDF.",
    href: "docs/getting-started.html",
  },
  {
    title: "YAML guide",
    summary: "Learn blocks, tables, visuals, styles, and reusable authoring patterns.",
    href: "docs/yaml-guide.html",
  },
  {
    title: "Compatibility matrix",
    summary: "See what works in DOCX, PDF, both outputs, or not yet.",
    href: "docs/compatibility.html",
  },
  {
    title: "Troubleshooting",
    summary: "Diagnose install, font, asset, viewer, and large-file problems.",
    href: "docs/troubleshooting.html",
  },
];

const EXAMPLES = [
  {
    title: "Board Report",
    summary: "Two-page leadership packet with narrative, metrics, and scorecard tables.",
    preview: "assets/gallery/board-report.png",
    yaml: "examples/board_report.yaml",
    docx: "generated/board-report.docx",
    pdf: "rendered/board-report.pdf",
    tags: ["cover page", "metrics", "board packet"],
  },
  {
    title: "Executive Dashboard",
    summary: "Multi-section KPI summary with narrative, delivery status, and risk tables.",
    preview: "assets/gallery/executive-dashboard.png",
    yaml: "examples/executive_dashboard.yaml",
    docx: "generated/executive-dashboard.docx",
    pdf: "rendered/executive-dashboard.pdf",
    tags: ["dashboard", "status tables", "KPI"],
  },
  {
    title: "Product Launch Brief",
    summary: "Launch packet with milestones, readiness checks, and customer-facing narrative.",
    preview: "assets/gallery/product-launch-brief.png",
    yaml: "examples/product_launch_brief.yaml",
    docx: "generated/product-launch-brief.docx",
    pdf: "rendered/product-launch-brief.pdf",
    tags: ["launch", "milestones", "readiness"],
  },
  {
    title: "Talent Profile",
    summary: "Resume-style example combining narrative, tables, and a signature line.",
    preview: "assets/gallery/talent-profile.png",
    yaml: "examples/talent_profile.yaml",
    docx: "generated/talent-profile.docx",
    pdf: "rendered/talent-profile.pdf",
    tags: ["profile", "resume", "table layout"],
  },
  {
    title: "Invoice",
    summary: "Compact billing layout with label-value metadata, line items, and totals.",
    preview: "assets/gallery/invoice.png",
    yaml: "examples/invoice.yaml",
    docx: "generated/invoice.docx",
    pdf: "rendered/invoice.pdf",
    tags: ["invoice", "billing", "table"],
  },
  {
    title: "Meeting Notes",
    summary: "Short operational note with metadata, agenda bullets, and a decision log.",
    preview: "assets/gallery/meeting-notes.png",
    yaml: "examples/meeting_notes.yaml",
    docx: "generated/meeting-notes.docx",
    pdf: "rendered/meeting-notes.pdf",
    tags: ["notes", "agenda", "decision log"],
  },
];

renderDocs();
renderExamples();
wireCopyButtons();

function renderDocs() {
  const root = document.querySelector("[data-doc-preview]");
  if (!root) return;

  root.innerHTML = DOCS.map(
    (doc) => `
      <article class="doc-preview-item">
        <a href="${doc.href}">${escapeHtml(doc.title)}</a>
        <span class="muted">${escapeHtml(doc.summary)}</span>
      </article>`,
  ).join("");
}

function renderExamples() {
  const root = document.querySelector("[data-example-grid]");
  if (!root) return;

  root.innerHTML = EXAMPLES.map(
    (example) => `
      <article class="example-card">
        <div class="example-preview">
          <img src="${example.preview}" alt="Preview of the ${escapeHtml(example.title)} PDF" loading="lazy" />
        </div>
        <div class="example-body">
          <div>
            <div class="tag-row">
              ${example.tags.map((tag) => `<span class="tag">${escapeHtml(tag)}</span>`).join("")}
            </div>
            <h3>${escapeHtml(example.title)}</h3>
            <p class="muted">${escapeHtml(example.summary)}</p>
          </div>
          <div class="action-row">
            <a class="button button-secondary" href="${example.yaml}">View YAML</a>
            <a class="button button-secondary" href="${example.pdf}">Open PDF</a>
            <a class="button button-primary" href="${example.docx}">Download DOCX</a>
          </div>
        </div>
      </article>`,
  ).join("");
}

function wireCopyButtons() {
  document.querySelectorAll("[data-copy]").forEach((button) => {
    button.addEventListener("click", async () => {
      const value = button.getAttribute("data-copy");
      if (!value) return;

      try {
        await navigator.clipboard.writeText(value);
        const previous = button.textContent;
        button.textContent = "Copied";
        window.setTimeout(() => {
          button.textContent = previous;
        }, 1400);
      } catch {
        button.textContent = "Copy failed";
      }
    });
  });
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
