#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const siteRoot = path.join(root, "site");
const baseUrl = "https://othmaneblial.github.io/rusdox/";
const checkOnly = process.argv.includes("--check");

const pages = [
  ["docs/README.md", "docs.html", "Documentation", "Start here", "Choose the shortest path from installation to a trustworthy DOCX and PDF workflow.", "Overview"],
  ["README.md", "docs/project-overview.html", "Project overview", "Overview", "The product promise, benchmark, examples, installation paths, and current support boundary.", "Overview"],
  ["docs/getting-started.md", "docs/getting-started.html", "Getting started", "Guide", "Install RusDox and generate the first editable DOCX and native PDF in minutes.", "Authoring"],
  ["docs/yaml-guide.md", "docs/yaml-guide.html", "YAML guide", "Guide", "Learn document blocks, styles, composition, tables, visuals, and reusable authoring patterns.", "Authoring"],
  ["docs/spec-versioning.md", "docs/spec-versioning.html", "Spec versioning", "Guide", "Generate schemas, migrate legacy specs, use deterministic expressions, and configure editor feedback.", "Authoring"],
  ["docs/stability.md", "docs/stability.html", "Stability and support", "Contract", "Understand v1 SemVer guarantees, deprecation windows, MSRV, output stability, and supported releases.", "Trust & operations"],
  ["docs/release-checklist.md", "docs/release-checklist.html", "Release checklist", "Operations", "Require contract, parity, compatibility, performance, supply-chain, and publication evidence for every release.", "Trust & operations"],
  ["docs/wasm-feasibility.md", "docs/wasm-feasibility.html", "Browser rendering feasibility", "Decision", "Understand what the browser playground proves, why edited files still use the CLI, and the gates for a full WASM renderer.", "Trust & operations"],
  ["docs/configuration.md", "docs/configuration.html", "Configuration", "Guide", "Control typography, spacing, color, tables, output paths, and PDF rendering.", "Authoring"],
  ["docs/word-templates.md", "docs/word-templates.html", "Word-native templates", "Guide", "Turn designer-authored DOCX files and JSON data into editable Word, native PDF, and parity evidence.", "Authoring"],
  ["docs/template-registry.md", "docs/template-registry.html", "Curated template registry", "Guide", "Discover, verify, install, update, and contribute signed Word templates without expanding the core crate.", "Examples"],
  ["docs/github-action.md", "docs/github-action.html", "GitHub Action", "Integration", "Validate specs, annotate pull-request lines, render outputs, and retain private parity evidence inside GitHub Actions.", "Trust & operations"],
  ["docs/integrations.md", "docs/integrations.html", "Integration protocol", "Integration", "Embed the renderer or call one stable local JSON protocol from Node, Python, Go, CI, and loopback HTTP.", "Reference"],
  ["docs/cli.md", "docs/cli.html", "CLI reference", "Reference", "Render, validate, watch, benchmark, initialize, and configure documents.", "Reference"],
  ["docs/rust-api.md", "docs/rust-api.html", "Rust API", "Reference", "Choose between DocumentSpec, Studio, and the low-level typed document model.", "Reference"],
  ["docs/compatibility.md", "docs/compatibility.html", "Compatibility matrix", "Trust", "See exactly what works in DOCX, PDF, both outputs, or not yet.", "Trust & operations"],
  ["docs/compatibility-scorecard.md", "docs/compatibility-scorecard.html", "Viewer compatibility scorecard", "Trust", "Review dated, hash-pinned evidence from real document viewers without universal claims.", "Trust & operations"],
  ["docs/international-accessibility.md", "docs/international-accessibility.html", "International and accessible output", "Trust", "Review script graduation gates, font embedding, alternative text, language metadata, and accessibility parity.", "Trust & operations"],
  ["docs/pdf-conformance-research.md", "docs/pdf-conformance-research.html", "PDF conformance research", "Decision", "Understand the tagged PDF, PDF/UA, PDF/A, validator, and human-review gates before any conformance claim.", "Trust & operations"],
  ["docs/input-safety.md", "docs/input-safety.html", "Input safety and limits", "Security", "Understand resource ceilings, fuzz targets, and atomic output recovery for untrusted inputs.", "Trust & operations"],
  ["docs/security-review-v1.md", "docs/security-review-v1.html", "v1 security review", "Security", "Review the ZIP, XML, visual, template, protocol, batch, and release threat model with residual risks.", "Trust & operations"],
  ["docs/production.md", "docs/production.html", "Production and batch rendering", "Operations", "Run bounded concurrent jobs with service-owned limits, ordered results, and cooperative cancellation.", "Trust & operations"],
  ["docs/parity.md", "docs/parity.html", "Parity verification", "Trust", "Generate machine-readable semantic checks and deterministic rendered-page diffs for DOCX and PDF.", "Trust & operations"],
  ["docs/performance.md", "docs/performance.html", "Reproducible performance", "Trust", "Reproduce isolated benchmark tiers, inspect raw evidence, and understand material regression thresholds.", "Trust & operations"],
  ["docs/troubleshooting.md", "docs/troubleshooting.html", "Troubleshooting", "Operations", "Diagnose installers, paths, fonts, viewer differences, large files, and CI failures.", "Trust & operations"],
  ["docs/gallery.md", "docs/gallery.html", "Template gallery", "Examples", "Browse real YAML inputs and generated DOCX/PDF output previews.", "Examples"],
  ["examples/README.md", "docs/examples.html", "Examples guide", "Examples", "Understand every bundled document fixture and how to render it.", "Examples"],
  ["ROADMAP.md", "docs/roadmap.html", "Roadmap", "Project", "Follow the path to verified parity, Word templates, integrations, and v1.", "Project"],
  ["CHANGELOG.md", "docs/changelog.html", "Changelog", "Project", "Review user-facing additions, changes, fixes, and migrations.", "Project"],
  ["CONTRIBUTING.md", "docs/contributing.html", "Contributing", "Community", "Set up the repository and submit a focused contribution.", "Community"],
  ["docs/architecture.md", "docs/architecture.html", "Architecture", "Community", "Follow one typed document model from authoring and validation through DOCX, PDF, parity, and every adapter.", "Community"],
  ["GOVERNANCE.md", "docs/governance.html", "Governance", "Community", "Review the decision model, path to committer, release authority, and conflict policy.", "Community"],
  ["CONTRIBUTORS.md", "docs/contributors.html", "Contributors", "Community", "Meet the people credited by the actual Git history and learn how new contributions are recognized.", "Community"],
  ["SUPPORT.md", "docs/support.html", "Support", "Community", "Find the right place for questions, reproducible bugs, and security reports.", "Community"],
  ["SECURITY.md", "docs/security.html", "Security policy", "Community", "Understand supported releases, reporting scope, and private disclosure.", "Community"],
  ["CODE_OF_CONDUCT.md", "docs/code-of-conduct.html", "Code of conduct", "Community", "Behavior and enforcement standards for a healthy community.", "Community"],
].map(([source, output, title, category, summary, group], index) => ({
  source,
  output,
  title,
  category,
  summary,
  group,
  index,
}));

const pageBySource = new Map(pages.map((page) => [normalize(page.source), page]));
const generated = new Map();

for (const page of pages) {
  const markdown = fs.readFileSync(path.join(root, page.source), "utf8");
  const rendered = renderMarkdown(markdown.replace(/^#\s+.*\n+/, ""), page);
  generated.set(page.output, renderPage(page, rendered));
}

generated.set("sitemap.xml", renderSitemap());
generated.set("robots.txt", `User-agent: *\nAllow: /\nSitemap: ${baseUrl}sitemap.xml\n`);
generated.set("llms.txt", renderLlmsIndex());
generated.set("llms-full.txt", renderLlmsFull());
generated.set("assets/quick-demo.svg", fs.readFileSync(path.join(root, "assets/quick-demo.svg"), "utf8"));
if (fs.existsSync(path.join(root, "assets/benchmark-history.svg"))) {
  generated.set("assets/benchmark-history.svg", fs.readFileSync(path.join(root, "assets/benchmark-history.svg"), "utf8"));
}

const paritySource = path.join(root, "reports", "gallery");
if (fs.existsSync(paritySource)) {
  for (const file of walkFiles(paritySource)) {
    const relative = normalize(path.relative(paritySource, file));
    generated.set(`parity/${relative}`, fs.readFileSync(file));
  }
}

const compatibilitySource = path.join(root, "compatibility");
if (fs.existsSync(compatibilitySource)) {
  for (const file of walkFiles(compatibilitySource)) {
    const relative = normalize(path.relative(compatibilitySource, file));
    generated.set(`compatibility/${relative}`, fs.readFileSync(file));
  }
}

for (const [sourceName, outputName] of [
  ["templates", "templates"],
  ["template-evidence", "template-evidence"],
  ["schema", "schema"],
  ["editors/vscode", "editors/vscode"],
  ["playground", "playground"],
  ["registry", "registry"],
]) {
  const sourceRoot = path.join(root, sourceName);
  if (fs.existsSync(sourceRoot)) {
    for (const file of walkFiles(sourceRoot)) {
      if (file.endsWith("-summary.json")) continue;
      const relative = normalize(path.relative(sourceRoot, file));
      generated.set(`${outputName}/${relative}`, fs.readFileSync(file));
    }
  }
}

const sourceCopies = [
  "README.md",
  "ROADMAP.md",
  "CHANGELOG.md",
  "CONTRIBUTING.md",
  "CONTRIBUTORS.md",
  "GOVERNANCE.md",
  "SUPPORT.md",
  "SECURITY.md",
  "CODE_OF_CONDUCT.md",
  ...fs.readdirSync(path.join(root, "docs")).filter((name) => name.endsWith(".md")).map((name) => `docs/${name}`),
  ...walkFiles(path.join(root, "examples"))
    .filter((file) => file.endsWith(".md") || file.endsWith(".yaml") || file.endsWith(".toml"))
    .map((file) => normalize(path.relative(root, file))),
  "fuzz/README.md",
  ...walkFiles(path.join(root, "benchmarks"))
    .filter((file) => file.endsWith(".json") || file.endsWith(".md"))
    .map((file) => normalize(path.relative(root, file))),
  "scripts/check_benchmark_regression.mjs",
  "scripts/build_benchmark_dashboard.mjs",
  "scripts/test_benchmark_contract.mjs",
  "scripts/render_benchmark_history.mjs",
  "scripts/run_benchmark_protocol.mjs",
  "scripts/check_compatibility_contract.mjs",
  "scripts/check_accessibility_contract.mjs",
  "scripts/check_security_review.mjs",
  "scripts/test_reproducible_release.mjs",
  "scripts/package_release.py",
  "scripts/verify_reproducible_build.py",
  "scripts/generate_word_templates.sh",
  "scripts/verify_word_templates.sh",
  "scripts/generate_schema.sh",
  "action.yml",
  "scripts/github_action.mjs",
  "scripts/github_action_comment.mjs",
  "scripts/test_github_action.mjs",
  "scripts/test_integrations.mjs",
  "scripts/starter_issues.mjs",
  "scripts/contributor_lab.mjs",
  "scripts/check_contributors.mjs",
  "scripts/check_stable_contracts.mjs",
  "scripts/test_contributor_lab.mjs",
  ".github/starter-issues.json",
  ...walkFiles(path.join(root, "contributor-fixtures"))
    .map((file) => normalize(path.relative(root, file))),
  ...walkFiles(path.join(root, "examples", "github-actions"))
    .map((file) => normalize(path.relative(root, file))),
  ...walkFiles(path.join(root, "examples", "integrations"))
    .map((file) => normalize(path.relative(root, file))),
];

for (const source of sourceCopies) {
  generated.set(source, fs.readFileSync(path.join(root, source), "utf8"));
}

const stale = [];
for (const [relativePath, content] of generated) {
  const destination = path.join(siteRoot, relativePath);
  if (checkOnly) {
    const current = fs.existsSync(destination) ? fs.readFileSync(destination) : null;
    const expected = Buffer.isBuffer(content) ? content : Buffer.from(content, "utf8");
    if (!current || !current.equals(expected)) {
      stale.push(relativePath);
    }
  } else {
    fs.mkdirSync(path.dirname(destination), { recursive: true });
    fs.writeFileSync(destination, content);
  }
}

function walkFiles(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(directory, entry.name);
    return entry.isDirectory() ? walkFiles(fullPath) : [fullPath];
  });
}

if (checkOnly && stale.length) {
  console.error("Static site outputs are stale:");
  stale.forEach((file) => console.error(`  ${file}`));
  console.error("Run: node scripts/build_site.mjs");
  process.exit(1);
}

console.log(
  checkOnly
    ? `Static site is current (${generated.size} checked files).`
    : `Generated ${generated.size} static site files.`,
);

function renderPage(page, rendered) {
  const depth = page.output.split("/").length - 1;
  const prefix = depth ? "../".repeat(depth) : "";
  const canonical = new URL(page.output, baseUrl).href;
  const previous = pages[page.index - 1];
  const next = pages[page.index + 1];
  const sourceUrl = `https://github.com/OthmaneBlial/rusdox/blob/main/${page.source}`;
  const structuredData = JSON.stringify({
    "@context": "https://schema.org",
    "@type": "TechArticle",
    headline: `RusDox ${page.title}`,
    description: page.summary,
    url: canonical,
    author: { "@type": "Person", name: "Othmane Blial", url: "https://github.com/OthmaneBlial" },
    isPartOf: { "@type": "WebSite", name: "RusDox", url: baseUrl },
  }).replaceAll("<", "\\u003c");

  const toc =
    rendered.headings.length > 1
      ? `<nav class="docs-toc" aria-label="On this page"><p>On this page</p><ol>${rendered.headings
          .map((heading) => `<li class="toc-level-${heading.level}"><a href="#${heading.id}">${escapeHtml(heading.text)}</a></li>`)
          .join("")}</ol></nav>`
      : "";
  const indexCards = page.index === 0 ? renderIndexCards(page) : "";
  const pager = `<nav class="docs-pager" aria-label="Documentation pagination">
    ${previous ? `<a href="${pageHref(page, previous)}"><span>Previous</span><strong>${escapeHtml(previous.title)}</strong></a>` : "<span></span>"}
    ${next ? `<a class="docs-pager-next" href="${pageHref(page, next)}"><span>Next</span><strong>${escapeHtml(next.title)}</strong></a>` : "<span></span>"}
  </nav>`;

  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${escapeHtml(page.title)} | RusDox Documentation</title>
    <meta name="description" content="${escapeAttribute(page.summary)}" />
    <link rel="canonical" href="${canonical}" />
    <meta property="og:title" content="${escapeAttribute(page.title)} | RusDox" />
    <meta property="og:description" content="${escapeAttribute(page.summary)}" />
    <meta property="og:image" content="${baseUrl}assets/social-preview-rusdox.png" />
    <meta property="og:url" content="${canonical}" />
    <meta property="og:type" content="article" />
    <meta name="twitter:card" content="summary_large_image" />
    <meta name="theme-color" content="#b85c30" />
    <link rel="icon" href="${prefix}assets/rusdox-mark.svg" type="image/svg+xml" />
    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
    <link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=IBM+Plex+Mono:wght@400;500;600&family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
    <link rel="stylesheet" href="${prefix}styles.css" />
    <script type="application/ld+json">${structuredData}</script>
  </head>
  <body data-page="docs-static">
    <a class="skip-link" href="#documentation">Skip to documentation</a>
    <div class="site-shell">
      <header class="site-header">
        <a class="brand" href="${prefix}index.html" aria-label="RusDox home">
          <span class="brand-mark">R</span>
          <span class="brand-text"><strong>RusDox</strong><span>One spec. Two trustworthy documents.</span></span>
        </a>
        <nav class="site-nav" aria-label="Primary">
          <a href="${prefix}index.html">Overview</a>
          <a class="is-active" href="${prefix}docs.html">Docs</a>
          <a href="${prefix}index.html#examples">Examples</a>
          <a href="${prefix}playground/">Playground</a>
          <a href="https://github.com/OthmaneBlial/rusdox">GitHub</a>
        </nav>
      </header>
      <main class="docs-static-main" id="documentation">
        <section class="docs-page-head docs-static-head">
          <div>
            <p class="eyebrow">${escapeHtml(page.category)}</p>
            <h1>${escapeHtml(page.title)}</h1>
            <p class="lede">${escapeHtml(page.summary)}</p>
          </div>
          <a class="button button-secondary" href="${sourceUrl}">View Markdown source</a>
        </section>
        <div class="docs-static-layout">
          ${renderNavigation(page, prefix)}
          <article class="docs-article docs-static-article">
            ${toc}
            <div class="docs-content prose">${rendered.html}${indexCards}</div>
            ${pager}
          </article>
        </div>
      </main>
      <footer class="site-footer">
        <p>RusDox documentation. Static, linkable, and usable without JavaScript.</p>
        <div class="footer-links">
          <a href="${prefix}docs/compatibility.html">Compatibility</a>
          <a href="${prefix}docs/troubleshooting.html">Troubleshooting</a>
          <a href="https://github.com/OthmaneBlial/rusdox/discussions">Discussions</a>
        </div>
      </footer>
    </div>
  </body>
</html>
`;
}

function renderNavigation(activePage, prefix) {
  const groups = [...new Set(pages.map((page) => page.group))];
  return `<aside class="docs-static-nav" aria-label="Documentation navigation">
    <p class="docs-nav-label">Field guide</p>
    ${groups
      .map((group) => `<section><h2>${escapeHtml(group)}</h2><ul>${pages
        .filter((page) => page.group === group)
        .map((page) => `<li><a${page === activePage ? ' class="is-active" aria-current="page"' : ""} href="${pageHref(activePage, page)}">${escapeHtml(page.title)}</a></li>`)
        .join("")}</ul></section>`)
      .join("")}
    <a class="docs-nav-home" href="${prefix}index.html">← Back to overview</a>
  </aside>`;
}

function renderIndexCards(activePage) {
  const groups = [...new Set(pages.filter((page) => page.index !== 0).map((page) => page.group))];
  return `<section class="docs-index-map" aria-labelledby="documentation-map">
    <p class="eyebrow">Documentation map</p>
    <h2 id="documentation-map">Find the exact layer you need.</h2>
    ${groups
      .map((group) => `<section class="docs-index-group"><h3>${escapeHtml(group)}</h3><div class="docs-index-grid">${pages
        .filter((page) => page.group === group && page.index !== 0)
        .map((page) => `<a class="docs-index-card" href="${pageHref(activePage, page)}"><span>${escapeHtml(page.category)}</span><strong>${escapeHtml(page.title)}</strong><p>${escapeHtml(page.summary)}</p></a>`)
        .join("")}</div></section>`)
      .join("")}
  </section>`;
}

function renderMarkdown(markdown, page) {
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const blocks = [];
  const headings = [];
  const ids = new Map();

  for (let index = 0; index < lines.length; ) {
    const line = lines[index];
    if (!line.trim()) {
      index += 1;
      continue;
    }

    const fence = line.match(/^```([\w-]*)\s*$/);
    if (fence) {
      const code = [];
      index += 1;
      while (index < lines.length && !/^```/.test(lines[index])) {
        code.push(lines[index]);
        index += 1;
      }
      index += 1;
      blocks.push(`<pre><code class="language-${escapeAttribute(fence[1] || "text")}">${escapeHtml(code.join("\n"))}</code></pre>`);
      continue;
    }

    const heading = line.match(/^(#{1,6})\s+(.*)$/);
    if (heading) {
      const level = heading[1].length;
      const text = stripInline(heading[2]);
      const id = uniqueSlug(text, ids);
      if (level <= 3) headings.push({ level, text, id });
      blocks.push(`<h${level} id="${id}">${renderInline(heading[2], page)}</h${level}>`);
      index += 1;
      continue;
    }

    if (/^\s*(---+|\*\*\*+)\s*$/.test(line)) {
      blocks.push("<hr />");
      index += 1;
      continue;
    }

    if (isTable(lines, index)) {
      const headers = splitRow(lines[index]);
      const rows = [];
      index += 2;
      while (index < lines.length && lines[index].includes("|") && lines[index].trim()) {
        rows.push(splitRow(lines[index]));
        index += 1;
      }
      blocks.push(`<div class="table-scroll"><table><thead><tr>${headers.map((cell) => `<th>${renderInline(cell, page)}</th>`).join("")}</tr></thead><tbody>${rows
        .map((row) => `<tr>${headers.map((_, cellIndex) => `<td>${renderInline(row[cellIndex] || "", page)}</td>`).join("")}</tr>`)
        .join("")}</tbody></table></div>`);
      continue;
    }

    if (/^>\s?/.test(line)) {
      const quote = [];
      while (index < lines.length && /^>\s?/.test(lines[index])) {
        quote.push(lines[index].replace(/^>\s?/, ""));
        index += 1;
      }
      blocks.push(`<blockquote>${renderInline(quote.join(" "), page)}</blockquote>`);
      continue;
    }

    const ordered = /^\s*\d+\.\s+/.test(line);
    const unordered = /^\s*[-*+]\s+/.test(line);
    if (ordered || unordered) {
      const tag = ordered ? "ol" : "ul";
      const pattern = ordered ? /^\s*\d+\.\s+(.*)$/ : /^\s*[-*+]\s+(.*)$/;
      const items = [];
      while (index < lines.length) {
        const match = lines[index].match(pattern);
        if (!match) break;
        let item = match[1];
        index += 1;
        while (index < lines.length && /^\s{2,}\S/.test(lines[index]) && !/^\s*[-*+]\s+/.test(lines[index]) && !/^\s*\d+\.\s+/.test(lines[index])) {
          item += " " + lines[index].trim();
          index += 1;
        }
        const task = item.match(/^\[([ xX])\]\s+(.*)$/);
        items.push(task
          ? `<li class="task-item"><input type="checkbox" disabled${task[1].toLowerCase() === "x" ? " checked" : ""} /> ${renderInline(task[2], page)}</li>`
          : `<li>${renderInline(item, page)}</li>`);
      }
      blocks.push(`<${tag}>${items.join("")}</${tag}>`);
      continue;
    }

    const image = line.match(/^!\[([^\]]*)\]\(([^)]+)\)\s*$/);
    if (image) {
      blocks.push(`<figure><img src="${escapeAttribute(resolveHref(image[2], page))}" alt="${escapeAttribute(image[1])}" loading="lazy" /><figcaption>${escapeHtml(image[1])}</figcaption></figure>`);
      index += 1;
      continue;
    }

    const paragraph = [];
    while (index < lines.length && isParagraph(lines, index)) {
      paragraph.push(lines[index].trim());
      index += 1;
    }
    blocks.push(`<p>${renderInline(paragraph.join(" "), page)}</p>`);
  }
  return { html: blocks.join("\n"), headings };
}

function renderInline(value, page) {
  const tokens = [];
  const reserve = (html) => {
    const token = `\u0000TOKEN${tokens.length}\u0000`;
    tokens.push(html);
    return token;
  };
  const restoreTokens = (value) => {
    let restored = value;
    tokens.forEach((token, index) => {
      restored = restored.replaceAll(`\u0000TOKEN${index}\u0000`, token);
    });
    return restored;
  };
  let text = value
    .replace(/`([^`]+)`/g, (_, code) => reserve(`<code>${escapeHtml(code)}</code>`))
    .replace(/!\[([^\]]*)\]\(([^)]+)\)/g, (_, alt, href) => reserve(`<img src="${escapeAttribute(resolveHref(href, page))}" alt="${escapeAttribute(alt)}" loading="lazy" />`))
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_, label, href) => {
      const external = /^(https?:|mailto:)/.test(href);
      const labelHtml = restoreTokens(escapeHtml(label));
      return reserve(`<a href="${escapeAttribute(resolveHref(href, page))}"${external ? ' target="_blank" rel="noreferrer"' : ""}>${labelHtml}</a>`);
    });
  text = escapeHtml(text)
    .replace(/\*\*([^*]+?)\*\*/g, "<strong>$1</strong>")
    .replace(/~~([^~]+?)~~/g, "<del>$1</del>")
    .replace(/(^|[^*])\*([^*]+?)\*/g, "$1<em>$2</em>");
  return restoreTokens(text);
}

function resolveHref(rawHref, page) {
  const href = rawHref.trim().replace(/^<|>$/g, "");
  if (/^(https?:|mailto:)/.test(href) || href.startsWith("#")) return href;
  const [pathname, fragment = ""] = href.split("#", 2);
  let resolved = normalize(path.posix.join(path.posix.dirname(page.source), pathname));
  if (resolved.startsWith("site/")) {
    resolved = resolved.slice("site/".length);
  }
  const targetPage = pageBySource.get(resolved);
  if (targetPage) {
    const target = pageHref(page, targetPage);
    return fragment ? `${target}#${fragment}` : target;
  }
  const target = path.posix.relative(path.posix.dirname(page.output), resolved) || ".";
  return fragment ? `${target}#${fragment}` : target;
}

function pageHref(from, to) {
  return path.posix.relative(path.posix.dirname(from.output), to.output) || path.posix.basename(to.output);
}

function isTable(lines, index) {
  return index + 1 < lines.length && lines[index].includes("|") && /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(lines[index + 1]);
}

function splitRow(line) {
  return line.trim().replace(/^\||\|$/g, "").split("|").map((cell) => cell.trim());
}

function isParagraph(lines, index) {
  const line = lines[index];
  if (!line.trim()) return false;
  if (/^(#{1,6})\s+/.test(line) || /^```/.test(line)) return false;
  if (/^\s*(---+|\*\*\*+)\s*$/.test(line) || /^>\s?/.test(line)) return false;
  if (/^\s*[-*+]\s+/.test(line) || /^\s*\d+\.\s+/.test(line)) return false;
  if (isTable(lines, index)) return false;
  return true;
}

function renderSitemap() {
  const urls = ["", ...pages.map((page) => page.output)];
  return `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${urls
    .map((url) => `  <url><loc>${new URL(url, baseUrl).href}</loc></url>`)
    .join("\n")}\n</urlset>\n`;
}

function renderLlmsIndex() {
  return `# RusDox\n\n> Pure-Rust document engine: one readable spec to editable DOCX and native PDF without Word or LibreOffice.\n\n## Documentation\n\n${pages
    .map((page) => `- [${page.title}](${new URL(page.output, baseUrl).href}): ${page.summary}`)
    .join("\n")}\n\n## Source\n\n- Repository: https://github.com/OthmaneBlial/rusdox\n- License: MIT\n`;
}

function renderLlmsFull() {
  return pages
    .map((page) => `# ${page.title}\n\nSource: ${new URL(page.output, baseUrl).href}\n\n${fs
      .readFileSync(path.join(root, page.source), "utf8")
      .replace(/^#\s+.*\n+/, "")
      .trim()}`)
    .join("\n\n---\n\n");
}

function uniqueSlug(value, ids) {
  const base = value.toLowerCase().normalize("NFKD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "section";
  const count = ids.get(base) || 0;
  ids.set(base, count + 1);
  return count ? `${base}-${count + 1}` : base;
}

function stripInline(value) {
  return value.replace(/`([^`]+)`/g, "$1").replace(/\[([^\]]+)\]\([^)]+\)/g, "$1").replace(/[*_~]/g, "").trim();
}

function normalize(value) {
  return value.replaceAll("\\", "/").replace(/^\.?\//, "");
}

function escapeHtml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value);
}
