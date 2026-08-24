#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const directoryArgumentIndex = process.argv.indexOf("--registry-dir");
const relativeDirectory = directoryArgumentIndex >= 0 ? process.argv[directoryArgumentIndex + 1] : "registry";
assert(relativeDirectory, "--registry-dir requires a value");
const registryDirectory = path.resolve(root, relativeDirectory);
assert(
  registryDirectory === path.join(root, "registry") || registryDirectory.startsWith(`${path.join(root, "registry")}${path.sep}`),
  "registry directory must be registry/ or one of its descendants",
);
const indexPath = path.join(registryDirectory, "index.json");
const signaturePath = path.join(registryDirectory, "index.sig.json");
const publicKeyPath = path.join(registryDirectory, "public-key.pem");
const outputPath = path.join(registryDirectory, "preview.html");
const registryRelative = path.relative(root, registryDirectory).replaceAll(path.sep, "/");
const checkOnly = process.argv.includes("--check");
const requiredCategories = ["invoices", "proposals", "reports", "compliance", "hr", "education", "operations"];

const indexBytes = fs.readFileSync(indexPath);
const registry = JSON.parse(indexBytes);
const signature = JSON.parse(fs.readFileSync(signaturePath, "utf8"));
const publicKey = fs.readFileSync(publicKeyPath, "utf8");
const digest = crypto.createHash("sha256").update(indexBytes).digest("hex");

assert(signature.schema_version === 1, "signature schema_version must be 1");
assert(signature.algorithm === "ed25519", "signature algorithm must be ed25519");
assert(signature.manifest_sha256 === digest, "manifest SHA-256 does not match signature record");
assert(
  crypto.verify(null, indexBytes, publicKey, Buffer.from(signature.signature, "hex")),
  "Ed25519 registry signature is invalid",
);
assert(registry.schema_version === 1, "registry schema_version must be 1");

const categoryIds = new Set(registry.categories.map((category) => category.id));
requiredCategories.forEach((category) => assert(categoryIds.has(category), `missing category ${category}`));
const ids = new Set();

for (const entry of registry.templates) {
  assert(/^[a-z0-9][a-z0-9-]*$/.test(entry.id), `unsafe template ID ${entry.id}`);
  assert(!ids.has(entry.id), `duplicate template ID ${entry.id}`);
  ids.add(entry.id);
  assert(entry.title && entry.description && entry.author?.name && entry.author?.url, `${entry.id}: missing identity or author`);
  assert(entry.license?.spdx && entry.license?.url, `${entry.id}: missing license`);
  assert(entry.preview?.alt, `${entry.id}: missing preview alt text`);
  assert(entry.inputs?.length, `${entry.id}: missing documented inputs`);
  assert(entry.accessibility?.language && entry.accessibility?.notes, `${entry.id}: missing accessibility notes`);
  assert(entry.accessibility.reading_order_reviewed === true, `${entry.id}: reading order not reviewed`);
  assert(entry.accessibility.color_only_meaning === false, `${entry.id}: color-only meaning is not allowed`);
  entry.categories.forEach((category) => assert(categoryIds.has(category), `${entry.id}: unknown category ${category}`));

  for (const asset of [
    entry.preview,
    entry.files.template,
    entry.files.sample_data,
    entry.verified_outputs.docx,
    entry.verified_outputs.pdf,
    entry.verified_outputs.parity_json,
    entry.verified_outputs.parity_html,
  ]) {
    assert(/^[a-f0-9]{64}$/.test(asset.sha256), `${entry.id}: invalid SHA-256 for ${asset.url}`);
    const local = localPathForUrl(asset.url, registry.base_url);
    assert(fs.existsSync(local), `${entry.id}: missing local asset ${path.relative(root, local)}`);
    const actual = crypto.createHash("sha256").update(fs.readFileSync(local)).digest("hex");
    assert(actual === asset.sha256, `${entry.id}: stale hash for ${path.relative(root, local)}`);
  }
}

assert(ids.has(registry.template_of_the_month), "template_of_the_month is not a registry entry");
const rootPrefix = path.relative(registryDirectory, root).replaceAll(path.sep, "/");
const html = renderRegistry(registry, digest, signature.key_id, rootPrefix, `${registry.base_url}${registryRelative}/preview.html`);
if (checkOnly) {
  assert(fs.existsSync(outputPath), "registry/preview.html is missing");
  assert(fs.readFileSync(outputPath, "utf8") === html, "registry/preview.html is stale");
  console.log(`${relativeDirectory} is current (${registry.templates.length} signed templates).`);
} else {
  fs.writeFileSync(outputPath, html);
  console.log(`Generated ${relativeDirectory}/preview.html with ${registry.templates.length} entries.`);
}

function localPathForUrl(url, baseUrl) {
  assert(url.startsWith(baseUrl), `asset URL is outside registry base: ${url}`);
  const relative = decodeURIComponent(url.slice(baseUrl.length));
  const resolved = path.resolve(root, relative);
  assert(resolved.startsWith(root + path.sep), `asset URL escapes repository: ${url}`);
  return resolved;
}

function renderRegistry(value, sha256, keyId, rootPrefix, canonicalUrl) {
  const cards = value.templates.map((entry) => {
    const featured = entry.id === value.template_of_the_month ? '<span class="featured">Template of the month</span>' : "";
    const inputs = entry.inputs.map((input) => `<li><code>${escapeHtml(input.path)}</code><span>${escapeHtml(input.description)}</span></li>`).join("");
    return `<article class="template-card">
      <img src="${rootPrefix}/${relativeUrl(entry.preview.url, value.base_url)}" alt="${escapeAttribute(entry.preview.alt)}" loading="lazy" />
      <div class="template-copy">${featured}<p class="category">${entry.categories.map(escapeHtml).join(" · ")}</p><h2>${escapeHtml(entry.title)}</h2><p>${escapeHtml(entry.description)}</p>
      <p class="credit">By <a href="${escapeAttribute(entry.author.url)}">${escapeHtml(entry.author.name)}</a> · ${escapeHtml(entry.license.spdx)} · v${escapeHtml(entry.version)}</p>
      <details><summary>Documented inputs</summary><ul>${inputs}</ul><p>${escapeHtml(entry.accessibility.notes)}</p></details>
      <div class="actions"><a href="${rootPrefix}/${relativeUrl(entry.files.template.url, value.base_url)}">Word template</a><a href="${rootPrefix}/${relativeUrl(entry.files.sample_data.url, value.base_url)}">Sample JSON</a><a href="${rootPrefix}/${relativeUrl(entry.verified_outputs.parity_html.url, value.base_url)}">Parity evidence</a></div></div>
    </article>`;
  }).join("");
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Verified template registry | RusDox</title><meta name="description" content="Signed, hash-verified Word templates for RusDox."><link rel="canonical" href="${escapeAttribute(canonicalUrl)}"><style>
:root{--ink:#17211b;--paper:#fffdf8;--bg:#f2ede4;--line:#d9cfc0;--accent:#b85c30;--green:#17633a}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.6 system-ui,sans-serif}main{width:min(1180px,calc(100% - 28px));margin:40px auto 80px}header{display:grid;grid-template-columns:1fr auto;gap:24px;align-items:end;padding-bottom:24px;border-bottom:2px solid var(--ink)}h1,h2{font-family:Georgia,serif;letter-spacing:-.025em}h1{max-width:12ch;margin:.3rem 0;font-size:clamp(2.8rem,7vw,6rem);line-height:.92}.eyebrow,.category{color:var(--accent);font:bold .72rem ui-monospace,monospace;text-transform:uppercase;letter-spacing:.12em}.trust{padding:16px;border:1px solid #a9c8ae;background:#eaf3e9}.trust strong{color:var(--green)}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:20px;margin-top:28px}.template-card{overflow:hidden;border:1px solid var(--line);background:var(--paper);box-shadow:0 12px 30px #392b1c12}.template-card img{width:100%;aspect-ratio:1.55/1;object-fit:cover;border-bottom:1px solid var(--line)}.template-copy{padding:22px}.template-copy h2{margin:.4rem 0;font-size:1.8rem}.featured{display:inline-block;padding:6px 8px;background:var(--accent);color:white;font:bold .68rem ui-monospace,monospace;text-transform:uppercase}.credit{color:#657067}.credit a{color:var(--accent)}details{margin:18px 0;padding:12px;border:1px solid var(--line)}summary{cursor:pointer;font-weight:700}details ul{padding:0;list-style:none}details li{display:grid;gap:2px;margin:9px 0}details span{color:#657067}.actions{display:flex;flex-wrap:wrap;gap:8px}.actions a{padding:9px 11px;background:#eee5d8;color:var(--ink);text-decoration:none;font-weight:700}.actions a:first-child{background:var(--accent);color:white}footer{margin-top:32px;color:#657067}@media(max-width:760px){header,.grid{grid-template-columns:1fr}h1{max-width:none}}
</style></head><body><main><header><div><p class="eyebrow">RusDox curated registry</p><h1>Word templates with receipts.</h1><p>Every entry has a license, author, documented inputs, preview, output hashes, parity evidence, and accessibility notes.</p></div><div class="trust"><strong>Ed25519 verified</strong><br><code>${escapeHtml(keyId)}</code><br><small>SHA-256 ${escapeHtml(sha256.slice(0, 16))}…</small></div></header><section class="grid">${cards}</section><footer><p>Install with <code>rusdox template add &lt;id&gt;</code>. Templates remain outside the core crate.</p><p><a href="${rootPrefix}/docs/template-registry.html">Registry trust and contribution contract</a> · <a href="https://github.com/OthmaneBlial/rusdox">Source</a></p></footer></main></body></html>`;
}

function relativeUrl(url, baseUrl) {
  return url.slice(baseUrl.length);
}

function escapeHtml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value);
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}
