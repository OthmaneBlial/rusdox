const DOC_GROUPS = [
  {
    id: "overview",
    label: "Overview",
    docs: [
      {
        slug: "readme",
        title: "Project Overview",
        category: "Overview",
        summary: "The root README that positions RusDox, the benchmark, examples, and docs map.",
        path: "README.md",
      },
      {
        slug: "docs-index",
        title: "Docs Index",
        category: "Overview",
        summary: "A quick route through the official documentation set and the core RusDox model.",
        path: "docs/README.md",
      },
    ],
  },
  {
    id: "guides",
    label: "Guides",
    docs: [
      {
        slug: "getting-started",
        title: "Getting Started",
        category: "Guide",
        summary: "Install the CLI, generate a starter YAML file, render outputs, and adjust config.",
        path: "docs/getting-started.md",
      },
      {
        slug: "yaml-guide",
        title: "YAML Guide",
        category: "Guide",
        summary: "Block types, tables, custom paragraphs, and authoring patterns for readable document specs.",
        path: "docs/yaml-guide.md",
      },
      {
        slug: "configuration",
        title: "Configuration",
        category: "Guide",
        summary: "How `rusdox.toml` and user config files control typography, spacing, output, and PDF behavior.",
        path: "docs/configuration.md",
      },
      {
        slug: "cli-guide",
        title: "CLI Guide",
        category: "Guide",
        summary: "Common commands, output modes, project overrides, and the advanced script entrypoint.",
        path: "docs/cli.md",
      },
      {
        slug: "template-gallery",
        title: "Template Gallery",
        category: "Guide",
        summary: "Image-backed previews of real RusDox outputs generated from the bundled example files.",
        path: "docs/gallery.md",
      },
      {
        slug: "rust-api",
        title: "Rust API",
        category: "Guide",
        summary: "When to use the CLI, `DocumentSpec`, `Studio`, or the low-level Rust document model.",
        path: "docs/rust-api.md",
      },
      {
        slug: "github-setup",
        title: "GitHub Setup",
        category: "Guide",
        summary: "The repository presentation checklist for GitHub topics, previews, discussions, and profile pinning.",
        path: "docs/github-setup.md",
      },
    ],
  },
  {
    id: "examples",
    label: "Examples",
    docs: [
      {
        slug: "examples-readme",
        title: "Examples Guide",
        category: "Examples",
        summary: "What lives in `examples/`, how to render it, and which sample specs are most representative.",
        path: "examples/README.md",
      },
    ],
  },
  {
    id: "community",
    label: "Community",
    docs: [
      {
        slug: "contributing",
        title: "Contributing",
        category: "Community",
        summary: "Contribution expectations, local setup, pull request checklist, and areas that matter most.",
        path: "CONTRIBUTING.md",
      },
      {
        slug: "support",
        title: "Support",
        category: "Community",
        summary: "Where to ask for help, what to include, and when a security issue should stay private.",
        path: "SUPPORT.md",
      },
      {
        slug: "security",
        title: "Security",
        category: "Community",
        summary: "Supported versions, vulnerability reporting process, and what falls inside project scope.",
        path: "SECURITY.md",
      },
      {
        slug: "code-of-conduct",
        title: "Code of Conduct",
        category: "Community",
        summary: "Contributor Covenant expectations and the enforcement ladder for unacceptable behavior.",
        path: "CODE_OF_CONDUCT.md",
      },
    ],
  },
];

const EXAMPLES = [
  {
    title: "Board Report",
    summary: "Two-page leadership packet with cover title, board narrative, metrics, and scorecard tables.",
    preview: "assets/gallery/board-report.png",
    yaml: "examples/board_report.yaml",
    docx: "generated/board-report.docx",
    pdf: "rendered/board-report.pdf",
    tags: ["cover page", "metrics", "board packet"],
  },
  {
    title: "Executive Dashboard",
    summary: "Multi-section KPI summary with narrative, metric cards, delivery status, and risk tables.",
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
    summary: "Resume-style example that mixes narrative sections, tables, and a strong final signature line.",
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

const ALL_DOCS = DOC_GROUPS.flatMap((group) => group.docs);
const DOC_BY_SLUG = new Map(ALL_DOCS.map((doc) => [doc.slug, doc]));
const DOC_BY_PATH = new Map(ALL_DOCS.map((doc) => [normalizePath(doc.path), doc]));
const activePage = document.body.dataset.page || "home";

const docGroupsRoot = document.querySelector("[data-doc-groups]");
const docSearch = document.querySelector("[data-doc-search]");
const docPath = document.querySelector("[data-doc-path]");
const docKicker = document.querySelector("[data-doc-kicker]");
const docTitle = document.querySelector("[data-doc-title]");
const docSummary = document.querySelector("[data-doc-summary]");
const docSource = document.querySelector("[data-doc-source]");
const docContent = document.querySelector("[data-doc-content]");
const docPreviewRoot = document.querySelector("[data-doc-preview]");
const exampleGrid = document.querySelector("[data-example-grid]");

let activeSlug = "readme";

bootstrap();

function bootstrap() {
  renderExamples();
  renderDocPreview();
  wireCopyButtons();
  applyCounts();

  if (docGroupsRoot && docSearch && docContent) {
    initializeDocsPage();
  }
}

function applyCounts() {
  document.querySelectorAll("[data-doc-count]").forEach((node) => {
    node.textContent = String(ALL_DOCS.length);
  });
  document.querySelectorAll("[data-example-count]").forEach((node) => {
    node.textContent = String(EXAMPLES.length);
  });
}

function initializeDocsPage() {
  renderDocNav();

  docSearch.addEventListener("input", () => {
    renderDocNav(docSearch.value.trim().toLowerCase());
  });

  window.addEventListener("hashchange", handleHash);
  handleHash();
}

function handleHash() {
  const hash = window.location.hash.replace(/^#/, "");
  const match = /^doc:(.+)$/.exec(hash);
  const slug = match?.[1] || activeSlug || "readme";

  if (DOC_BY_SLUG.has(slug)) {
    loadDoc(slug);
    return;
  }

  loadDoc("readme");
}

function renderExamples() {
  if (!exampleGrid) {
    return;
  }

  exampleGrid.innerHTML = EXAMPLES.map(
    (example) => `
      <article class="example-card">
        <div class="example-preview">
          <img src="${example.preview}" alt="${escapeHtml(example.title)} preview" loading="lazy" />
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
            <a class="button button-secondary" href="${example.yaml}" target="_blank" rel="noreferrer">View YAML</a>
            <a class="button button-secondary" href="${example.pdf}" target="_blank" rel="noreferrer">Open PDF</a>
            <a class="button button-primary" href="${example.docx}" target="_blank" rel="noreferrer">Download DOCX</a>
          </div>
        </div>
      </article>
    `,
  ).join("");
}

function renderDocPreview() {
  if (!docPreviewRoot) {
    return;
  }

  const featuredDocs = [
    DOC_BY_SLUG.get("getting-started"),
    DOC_BY_SLUG.get("yaml-guide"),
    DOC_BY_SLUG.get("configuration"),
    DOC_BY_SLUG.get("rust-api"),
  ].filter(Boolean);

  docPreviewRoot.innerHTML = featuredDocs
    .map(
      (doc) => `
        <article class="doc-preview-item">
          <a href="docs.html#doc:${doc.slug}">${escapeHtml(doc.title)}</a>
          <span class="muted">${escapeHtml(doc.summary)}</span>
        </article>
      `,
    )
    .join("");
}

function renderDocNav(filterText = "") {
  const renderedGroups = DOC_GROUPS.map((group) => {
    const docs = group.docs.filter((doc) => {
      const haystack = `${doc.title} ${doc.summary} ${doc.category} ${doc.path}`.toLowerCase();
      return !filterText || haystack.includes(filterText);
    });

    if (!docs.length) {
      return "";
    }

    return `
      <section class="doc-group">
        <h3 class="doc-group-title">${escapeHtml(group.label)}</h3>
        <div class="doc-nav">
          ${docs.map(renderDocButton).join("")}
        </div>
      </section>
    `;
  }).join("");

  docGroupsRoot.innerHTML =
    renderedGroups ||
    `<div class="empty-state">No documentation pages match this filter.</div>`;

  docGroupsRoot.querySelectorAll("[data-doc-slug]").forEach((button) => {
    button.addEventListener("click", () => {
      const slug = button.getAttribute("data-doc-slug");
      if (slug) {
        window.location.hash = `doc:${slug}`;
      }
    });
  });
}

function renderDocButton(doc) {
  const activeClass = doc.slug === activeSlug ? " is-active" : "";
  return `
    <button class="doc-button${activeClass}" type="button" data-doc-slug="${doc.slug}">
      <span class="doc-button-title">${escapeHtml(doc.title)}</span>
      <span class="doc-button-meta">${escapeHtml(doc.summary)}</span>
    </button>
  `;
}

async function loadDoc(slug) {
  const doc = DOC_BY_SLUG.get(slug);
  if (!doc) {
    return;
  }

  activeSlug = slug;
  renderDocNav(docSearch?.value.trim().toLowerCase() || "");

  if (docKicker) {
    docKicker.textContent = doc.category;
  }
  if (docTitle) {
    docTitle.textContent = doc.title;
  }
  if (docSummary) {
    docSummary.textContent = doc.summary;
  }
  if (docPath) {
    docPath.textContent = doc.path;
  }
  if (docSource) {
    docSource.href = doc.path;
  }
  if (!docContent) {
    return;
  }

  docContent.innerHTML = `<p>Loading ${escapeHtml(doc.title)}…</p>`;

  try {
    const response = await fetch(doc.path);
    if (!response.ok) {
      throw new Error(`Failed to load ${doc.path}`);
    }
    const markdown = await response.text();
    docContent.innerHTML = renderMarkdown(markdown, doc.path);
    enhanceDocLinks();
  } catch (error) {
    docContent.innerHTML = `
      <div class="empty-state">
        <strong>Could not load ${escapeHtml(doc.path)}</strong>
        <p>${escapeHtml(error.message)}</p>
      </div>
    `;
  }
}

function enhanceDocLinks() {
  if (!docContent) {
    return;
  }

  docContent.querySelectorAll("a[data-doc-target]").forEach((link) => {
    link.addEventListener("click", (event) => {
      if (activePage !== "docs") {
        return;
      }
      const slug = link.getAttribute("data-doc-target");
      if (!slug) {
        return;
      }
      event.preventDefault();
      window.location.hash = `doc:${slug}`;
    });
  });
}

function renderMarkdown(markdown, basePath) {
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const blocks = [];

  for (let index = 0; index < lines.length; ) {
    const line = lines[index];

    if (!line.trim()) {
      index += 1;
      continue;
    }

    const codeFence = line.match(/^```([\w-]*)\s*$/);
    if (codeFence) {
      const language = codeFence[1] || "";
      const codeLines = [];
      index += 1;
      while (index < lines.length && !/^```/.test(lines[index])) {
        codeLines.push(lines[index]);
        index += 1;
      }
      index += 1;
      blocks.push(
        `<pre><code class="language-${escapeHtml(language)}">${escapeHtml(codeLines.join("\n"))}</code></pre>`,
      );
      continue;
    }

    const heading = line.match(/^(#{1,6})\s+(.*)$/);
    if (heading) {
      const level = heading[1].length;
      blocks.push(`<h${level}>${renderInline(heading[2], basePath)}</h${level}>`);
      index += 1;
      continue;
    }

    const orderedMatch = line.match(/^\d+\.\s+/);
    const unorderedMatch = line.match(/^[-*+]\s+/);
    if (orderedMatch || unorderedMatch) {
      const tag = orderedMatch ? "ol" : "ul";
      const items = [];
      while (index < lines.length) {
        const candidate = lines[index];
        const itemMatch = orderedMatch
          ? candidate.match(/^\d+\.\s+(.*)$/)
          : candidate.match(/^[-*+]\s+(.*)$/);
        if (!itemMatch) {
          break;
        }
        items.push(`<li>${renderInline(itemMatch[1], basePath)}</li>`);
        index += 1;
      }
      blocks.push(`<${tag}>${items.join("")}</${tag}>`);
      continue;
    }

    const imageOnly = line.match(/^!\[(.*)\]\((.*)\)\s*$/);
    if (imageOnly) {
      const images = [];
      while (index < lines.length) {
        const candidate = lines[index].match(/^!\[(.*)\]\((.*)\)\s*$/);
        if (!candidate) {
          break;
        }
        images.push({
          alt: candidate[1],
          src: resolveAssetPath(candidate[2], basePath),
        });
        index += 1;
      }

      if (images.length > 1) {
        blocks.push(
          `<div class="badge-row">${images
            .map(
              (image) =>
                `<img src="${escapeAttribute(image.src)}" alt="${escapeAttribute(image.alt)}" loading="lazy" />`,
            )
            .join("")}</div>`,
        );
      } else {
        const image = images[0];
        blocks.push(
          `<figure><img src="${escapeAttribute(image.src)}" alt="${escapeAttribute(image.alt)}" loading="lazy" /><figcaption>${escapeHtml(image.alt)}</figcaption></figure>`,
        );
      }
      continue;
    }

    const paragraphLines = [];
    while (index < lines.length && shouldContinueParagraph(lines[index])) {
      paragraphLines.push(lines[index].trim());
      index += 1;
    }
    blocks.push(`<p>${renderInline(paragraphLines.join(" "), basePath)}</p>`);
  }

  return blocks.join("\n");
}

function shouldContinueParagraph(line) {
  if (!line.trim()) {
    return false;
  }
  if (/^(#{1,6})\s+/.test(line)) {
    return false;
  }
  if (/^```/.test(line)) {
    return false;
  }
  if (/^\d+\.\s+/.test(line)) {
    return false;
  }
  if (/^[-*+]\s+/.test(line)) {
    return false;
  }
  return true;
}

function renderInline(text, basePath) {
  const codeTokens = [];
  let html = text
    .replace(/`([^`]+)`/g, (_, code) => {
      const token = `@@CODE${codeTokens.length}@@`;
      codeTokens.push(`<code>${escapeHtml(code)}</code>`);
      return token;
    })
    .replace(/!\[([^\]]*)\]\(([^)]+)\)/g, (_, alt, rawHref) => {
      const src = resolveAssetPath(rawHref, basePath);
      return `<img src="${escapeAttribute(src)}" alt="${escapeAttribute(alt)}" loading="lazy" />`;
    })
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_, label, rawHref) => renderLink(label, rawHref, basePath))
    .replace(/\*\*([^*]+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*([^*]+?)\*/g, "<em>$1</em>");

  codeTokens.forEach((token, index) => {
    html = html.replace(`@@CODE${index}@@`, token);
  });

  return html;
}

function renderLink(label, rawHref, basePath) {
  const resolved = resolveAssetPath(rawHref, basePath);
  const pathDoc = DOC_BY_PATH.get(normalizePath(resolved));
  if (pathDoc) {
    return `<a href="docs.html#doc:${pathDoc.slug}" data-doc-target="${pathDoc.slug}">${escapeHtml(label)}</a>`;
  }

  const external = /^(https?:)?\/\//.test(rawHref);
  const target = external ? ' target="_blank" rel="noreferrer"' : ' target="_blank" rel="noreferrer"';
  return `<a href="${escapeAttribute(resolved)}"${target}>${escapeHtml(label)}</a>`;
}

function resolveAssetPath(rawHref, basePath) {
  if (/^(https?:)?\/\//.test(rawHref) || rawHref.startsWith("#")) {
    return rawHref;
  }

  const cleanedBase = normalizePath(basePath);
  const baseParts = cleanedBase.split("/");
  baseParts.pop();
  const hrefParts = rawHref.split("/");
  const parts = [...baseParts];

  hrefParts.forEach((part) => {
    if (!part || part === ".") {
      return;
    }
    if (part === "..") {
      parts.pop();
      return;
    }
    parts.push(part);
  });

  return parts.join("/");
}

function normalizePath(path) {
  return path.replace(/\\/g, "/").replace(/^\.\//, "");
}

function wireCopyButtons() {
  document.querySelectorAll("[data-copy]").forEach((button) => {
    button.addEventListener("click", async () => {
      const value = button.getAttribute("data-copy");
      if (!value) {
        return;
      }

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

function escapeAttribute(value) {
  return escapeHtml(value);
}
