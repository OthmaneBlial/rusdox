#!/usr/bin/env node

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputRoot = path.join(root, "target/contributor-lab-test");
const issues = JSON.parse(run(["scripts/contributor_lab.mjs", "list", "--json"]));
assert.equal(issues.length, 10);
assert.equal(new Set(issues.map((issue) => issue.id)).size, 10);

const prepared = run([
  "scripts/contributor_lab.mjs",
  "prepare",
  "protocol-inline-toml",
  "--output-root",
  outputRoot,
]).trim();
assert.ok(prepared.startsWith(outputRoot));
assert.ok(fs.statSync(path.join(prepared, "inline-toml.json")).isFile());
assert.match(fs.readFileSync(path.join(prepared, "TASK.md"), "utf8"), /Acceptance criteria/);

const dryRun = JSON.parse(run([
  "scripts/contributor_lab.mjs",
  "visual-diff",
  "--spec",
  "examples/hello_world.yaml",
  "--baseline",
  "tests/golden/pages/linux-x86_64/hello-world",
  "--dry-run",
]));
assert.ok(dryRun.command.includes("--visual-baseline"));
assert.ok(dryRun.command.includes("--visual-threshold"));
console.log("Contributor fixture and visual-diff lab passed.");

function run(args) {
  const result = spawnSync("node", args, { cwd: root, encoding: "utf8" });
  if (result.status !== 0) {
    process.stderr.write(result.stdout || "");
    process.stderr.write(result.stderr || "");
    process.exit(result.status ?? 1);
  }
  return result.stdout;
}
