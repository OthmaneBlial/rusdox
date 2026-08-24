#!/usr/bin/env node

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const binary = path.resolve(process.env.RUSDOX_BIN || path.join(root, "target/debug/rusdox"));
const outputRoot = path.join(root, "target/integration-examples");
const environment = { ...process.env, RUSDOX_BIN: binary };

run("node", ["examples/integrations/node.mjs", outputRoot]);
run("python3", ["examples/integrations/python.py", outputRoot]);
if (commandExists("go")) run("go", ["run", "./examples/integrations/go/main.go", outputRoot]);
run("sh", ["examples/integrations/ci.sh", outputRoot]);

for (const relative of [
  "node/node-report.docx",
  "node/node-report.pdf",
  "python/python-report.docx",
  "python/python-report.pdf",
  "ci/ci-report.docx",
  "ci/ci-report.pdf",
]) {
  assert.ok(fs.statSync(path.join(outputRoot, relative)).size > 0, `missing ${relative}`);
}
if (commandExists("go")) {
  for (const relative of ["go/go-report.docx", "go/go-report.pdf"]) {
    assert.ok(fs.statSync(path.join(outputRoot, relative)).size > 0, `missing ${relative}`);
  }
}
console.log(`Integration examples passed${commandExists("go") ? " (Node, Python, Go, CI)" : " (Node, Python, CI; Go unavailable)"}.`);

function run(command, args) {
  const result = spawnSync(command, args, {
    cwd: root,
    env: environment,
    encoding: "utf8",
    maxBuffer: 16 * 1024 * 1024,
  });
  if (result.status !== 0) {
    process.stderr.write(result.stdout || "");
    process.stderr.write(result.stderr || "");
    process.exit(result.status ?? 1);
  }
  const response = JSON.parse(result.stdout.trim());
  assert.equal(response.ok, true, `${command} integration response`);
}

function commandExists(command) {
  return spawnSync(command, ["version"], { stdio: "ignore" }).status === 0;
}
