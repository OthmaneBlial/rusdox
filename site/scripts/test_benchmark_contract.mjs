#!/usr/bin/env node

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const baselinePath = "benchmarks/baselines/ubuntu-latest.json";
const pass = run(["scripts/check_benchmark_regression.mjs", "--current", baselinePath, "--baseline", baselinePath]);
assert.equal(pass.status, 0, pass.stderr);
const report = JSON.parse(pass.stdout);
assert.equal(report.passed, true);
assert.equal(report.comparisons.length, 13);
assert.ok(report.comparisons.every((item) => item.runtime_budget_ms > 0 && item.memory_budget_bytes > 0));

const oversized = JSON.parse(fs.readFileSync(path.join(root, baselinePath), "utf8"));
oversized.scenarios.find((item) => item.id === "small-dual").result.total_ms.median = 31;
const fixture = path.join(root, "target/benchmark-contract-over-budget.json");
fs.mkdirSync(path.dirname(fixture), { recursive: true });
fs.writeFileSync(fixture, JSON.stringify(oversized));
const fail = run(["scripts/check_benchmark_regression.mjs", "--current", fixture, "--baseline", baselinePath]);
assert.equal(fail.status, 1);
assert.ok(JSON.parse(fail.stdout).failures.includes("small-dual"));

const dashboard = run(["scripts/build_benchmark_dashboard.mjs", "--check"]);
assert.equal(dashboard.status, 0, dashboard.stderr);
console.log("Benchmark regression budgets and dashboard contract passed.");

function run(args) {
  return spawnSync("node", args, { cwd: root, encoding: "utf8" });
}
