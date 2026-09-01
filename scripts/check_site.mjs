#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const siteRoot = path.join(root, "site");
const files = walk(siteRoot);
const htmlFiles = files.filter((file) => file.endsWith(".html"));
const problems = [];
const sitemap = fs.readFileSync(path.join(siteRoot, "sitemap.xml"), "utf8");
const sitemapUrls = [...sitemap.matchAll(/<loc>([^<]+)<\/loc>/g)].map((match) => match[1]);
const sitemapHtml = new Set(sitemapUrls.map(publicUrlToFile).filter((file) => file?.endsWith(".html")));

for (const file of htmlFiles) {
  const relative = path.relative(siteRoot, file);
  const html = fs.readFileSync(file, "utf8");

  if (!/<title>[^<]+<\/title>/.test(html)) {
    problems.push(`${relative}: missing a non-empty title`);
  }
  if (!/<link rel="canonical" href="https:\/\/othmaneblial\.github\.io\/rusdox\//.test(html)) {
    problems.push(`${relative}: missing the canonical RusDox URL`);
  }
  if (!/<html\s+lang="[^"]+"/i.test(html)) {
    problems.push(`${relative}: missing document language`);
  }
  if (sitemapHtml.has(file) && !/<meta\s+[^>]*name="description"[^>]*content="[^"]+"/i.test(html)) {
    problems.push(`${relative}: sitemap page is missing a meta description`);
  }
  const canonical = html.match(/<link rel="canonical" href="(https:\/\/othmaneblial\.github\.io\/rusdox\/[^"]*)"/i)?.[1];
  if (canonical) {
    const canonicalFile = publicUrlToFile(canonical);
    if (!canonicalFile || !fs.existsSync(canonicalFile)) {
      problems.push(`${relative}: canonical target is not published: ${canonical}`);
    }
  }
  if (html.includes("\u0000") || /TOKEN\d+/.test(html)) {
    problems.push(`${relative}: unresolved generated-content placeholder`);
  }
  const interactiveEntry = relative === "index.html" || relative === path.join("playground", "index.html");
  if (!interactiveEntry && /<script(?![^>]+type="application\/ld\+json")/i.test(html)) {
    problems.push(`${relative}: documentation must not require executable JavaScript`);
  }

  const references = html.matchAll(/\b(?:href|src)="([^"]+)"/g);
  for (const [, reference] of references) {
    if (
      reference.startsWith("#") ||
      reference.startsWith("data:") ||
      reference.startsWith("mailto:") ||
      /^https?:\/\//.test(reference)
    ) {
      continue;
    }

    const clean = reference.split(/[?#]/, 1)[0];
    if (!clean) continue;
    const target = path.resolve(path.dirname(file), clean);
    if (!target.startsWith(`${siteRoot}${path.sep}`) && target !== siteRoot) {
      problems.push(`${relative}: reference escapes the site bundle: ${reference}`);
    } else if (!fs.existsSync(target)) {
      problems.push(`${relative}: missing local target: ${reference}`);
    }
  }
}

for (const url of sitemapUrls) {
  const file = publicUrlToFile(url);
  if (!file || !fs.existsSync(file)) problems.push(`sitemap.xml: missing published target: ${url}`);
}

const llmsFull = fs.readFileSync(path.join(siteRoot, "llms-full.txt"), "utf8");
for (const match of llmsFull.matchAll(/!?\[[^\]]*\]\(([^)]+)\)/g)) {
  const href = match[1].trim();
  if (!/^(https?:|mailto:)/.test(href)) problems.push(`llms-full.txt: relative Markdown target: ${href}`);
}

if (/data-(?:doc-preview|example-grid)/.test(fs.readFileSync(path.join(siteRoot, "index.html"), "utf8"))) {
  problems.push("index.html: critical homepage content must be present without JavaScript");
}

if (problems.length) {
  console.error("Static site verification failed:");
  problems.forEach((problem) => console.error(`  ${problem}`));
  process.exit(1);
}

console.log(`Static site verification passed (${htmlFiles.length} HTML pages, ${files.length} bundled files).`);

function walk(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(directory, entry.name);
    return entry.isDirectory() ? walk(fullPath) : [fullPath];
  });
}

function publicUrlToFile(rawUrl) {
  let url;
  try {
    url = new URL(rawUrl);
  } catch {
    return null;
  }
  if (url.origin !== "https://othmaneblial.github.io" || !url.pathname.startsWith("/rusdox/")) return null;
  let relative = decodeURIComponent(url.pathname.slice("/rusdox/".length));
  if (!relative || relative.endsWith("/")) relative += "index.html";
  const file = path.resolve(siteRoot, relative);
  return file.startsWith(`${siteRoot}${path.sep}`) ? file : null;
}
