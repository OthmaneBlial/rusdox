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

for (const file of htmlFiles) {
  const relative = path.relative(siteRoot, file);
  const html = fs.readFileSync(file, "utf8");

  if (!/<title>[^<]+<\/title>/.test(html)) {
    problems.push(`${relative}: missing a non-empty title`);
  }
  if (!/<link rel="canonical" href="https:\/\/othmaneblial\.github\.io\/rusdox\//.test(html)) {
    problems.push(`${relative}: missing the canonical RusDox URL`);
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
