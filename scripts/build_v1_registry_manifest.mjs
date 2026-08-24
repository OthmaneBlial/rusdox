#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const legacy = JSON.parse(fs.readFileSync(path.join(root, "registry/index.json"), "utf8"));
const outputPath = path.join(root, "registry/v1/index.json");
const checkOnly = process.argv.includes("--check");

const registry = structuredClone(legacy);
registry.registry_id = "rusdox-curated-v1";
registry.generated_at = "2026-08-24T00:00:00Z";

for (const entry of registry.templates) {
  entry.supported_rusdox = {
    minimum: "1.0.0",
    maximum_exclusive: "2.0.0",
  };
  for (const [kind, asset] of Object.entries(entry.verified_outputs)) {
    const suffix = {
      docx: `generated/${entry.id}.docx`,
      pdf: `rendered/${entry.id}.pdf`,
      parity_json: `reports/${entry.id}-parity.json`,
      parity_html: `reports/${entry.id}-parity.html`,
    }[kind];
    asset.url = `${registry.base_url}registry/v1/evidence/${entry.id}/${suffix}`;
    asset.sha256 = sha256(localPathForUrl(asset.url, registry.base_url));
  }
}

const output = `${JSON.stringify(registry, null, 2)}\n`;
if (checkOnly) {
  assert(fs.existsSync(outputPath), "registry/v1/index.json is missing");
  assert(fs.readFileSync(outputPath, "utf8") === output, "registry/v1/index.json is stale");
  const publicKeyPath = path.join(root, "registry/v1/public-key.pem");
  assert(fs.existsSync(publicKeyPath), "registry/v1/public-key.pem is missing");
  const publicKeyDer = crypto
    .createPublicKey(fs.readFileSync(publicKeyPath))
    .export({ type: "spki", format: "der" });
  const publicKeyHex = publicKeyDer.subarray(publicKeyDer.length - 32).toString("hex");
  const cli = fs.readFileSync(path.join(root, "src/bin/rusdox.rs"), "utf8");
  assert(
    cli.includes('https://othmaneblial.github.io/rusdox/registry/v1/index.json'),
    "CLI does not pin the immutable v1 registry URL",
  );
  assert(cli.includes(publicKeyHex), "CLI does not pin the immutable v1 registry public key");
  console.log(`v1 registry manifest is current (${registry.templates.length} templates).`);
} else {
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, output);
  console.log(`Generated ${path.relative(root, outputPath)}.`);
}

function localPathForUrl(url, baseUrl) {
  assert(url.startsWith(baseUrl), `asset URL is outside registry base: ${url}`);
  const resolved = path.resolve(root, decodeURIComponent(url.slice(baseUrl.length)));
  assert(resolved.startsWith(`${root}${path.sep}`), `asset URL escapes repository: ${url}`);
  assert(fs.existsSync(resolved), `registry asset is missing: ${path.relative(root, resolved)}`);
  return resolved;
}

function sha256(file) {
  return crypto.createHash("sha256").update(fs.readFileSync(file)).digest("hex");
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}
