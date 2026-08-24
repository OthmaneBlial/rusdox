#!/usr/bin/env node

import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const binary = fs.existsSync(path.join(root, "target/debug/rusdox.exe")) ? "target/debug/rusdox.exe" : "target/debug/rusdox";
assert.ok(fs.existsSync(path.join(root, binary)), "build target/debug/rusdox before this test");
const outputRoot = path.join(root, "target/reproducible-release-test");
fs.mkdirSync(outputRoot, { recursive: true });

for (const format of ["tar.gz", "zip"]) {
  const first = `target/reproducible-release-test/first.${format}`;
  const second = `target/reproducible-release-test/second.${format}`;
  for (const output of [first, second]) {
    const result = spawnSync("python3", [
      "scripts/package_release.py", "--binary", binary, "--format", format,
      "--output", output, "--source-date-epoch", "1787529600",
    ], { cwd: root, encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
  }
  assert.equal(sha(first), sha(second), `${format} archive bytes must be reproducible`);
}
console.log("Deterministic tar.gz and zip release packaging passed.");

function sha(relative) {
  return crypto.createHash("sha256").update(fs.readFileSync(path.join(root, relative))).digest("hex");
}
