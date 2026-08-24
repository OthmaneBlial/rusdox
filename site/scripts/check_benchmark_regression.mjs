#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = parseArgs(process.argv.slice(2));
if (!args.current || !args.baseline) {
  fail("usage: check_benchmark_regression.mjs --current REPORT --baseline REPORT");
}
const current = readJson(args.current);
const baseline = readJson(args.baseline);
const protocol = readJson("benchmarks/protocol.json");
const budgets = readJson("benchmarks/budgets.json");
const policy = protocol.regression_policy;

assertComparable(current, baseline);
const baselineById = new Map(baseline.scenarios.map((scenario) => [scenario.id, scenario]));
const budgetById = new Map(budgets.scenarios.map((scenario) => [scenario.id, scenario]));
const failures = [];
const comparisons = [];

for (const scenario of current.scenarios) {
  const previous = baselineById.get(scenario.id);
  if (!previous) fail(`baseline is missing scenario ${scenario.id}`);
  const budget = budgetById.get(scenario.id);
  if (!budget) fail(`budget is missing scenario ${scenario.id}`);
  const currentMs = scenario.result.total_ms.median;
  const baselineMs = previous.result.total_ms.median;
  const runtimeDelta = currentMs - baselineMs;
  const runtimeLimit = Math.max(
    policy.runtime_absolute_ms,
    baselineMs * policy.runtime_relative,
  );
  const memoryDelta = scenario.peak_memory_bytes - previous.peak_memory_bytes;
  const memoryLimit = Math.max(
    policy.memory_absolute_bytes,
    previous.peak_memory_bytes * policy.memory_relative,
  );
  const runtimeRegression = runtimeDelta > runtimeLimit;
  const memoryRegression = memoryDelta > memoryLimit;
  const runtimeBudgetExceeded = currentMs > budget.max_median_ms;
  const memoryBudgetExceeded = scenario.peak_memory_bytes > budget.max_peak_memory_bytes;
  comparisons.push({
    id: scenario.id,
    baseline_ms: baselineMs,
    current_ms: currentMs,
    runtime_delta_ms: runtimeDelta,
    baseline_memory_bytes: previous.peak_memory_bytes,
    current_memory_bytes: scenario.peak_memory_bytes,
    memory_delta_bytes: memoryDelta,
    runtime_budget_ms: budget.max_median_ms,
    memory_budget_bytes: budget.max_peak_memory_bytes,
    runtime_regression: runtimeRegression,
    memory_regression: memoryRegression,
    runtime_budget_exceeded: runtimeBudgetExceeded,
    memory_budget_exceeded: memoryBudgetExceeded,
  });
  if (runtimeRegression || memoryRegression || runtimeBudgetExceeded || memoryBudgetExceeded) {
    failures.push(scenario.id);
  }
}

console.log(JSON.stringify({
  passed: failures.length === 0,
  baseline: baseline.generated_at_utc,
  current: current.generated_at_utc,
  policy,
  budgets: "benchmarks/budgets.json",
  failures,
  comparisons,
}, null, 2));

if (failures.length > 0) process.exit(1);

function assertComparable(left, right) {
  if (left.protocol_version !== right.protocol_version) {
    fail(`protocol mismatch: ${left.protocol_version} != ${right.protocol_version}`);
  }
  for (const key of ["platform", "architecture"]) {
    if (left.host[key] !== right.host[key]) {
      fail(`host ${key} mismatch: ${left.host[key]} != ${right.host[key]}`);
    }
  }
  if (!Array.isArray(left.scenarios) || !Array.isArray(right.scenarios)) {
    fail("benchmark reports must contain scenario arrays");
  }
  for (const key of ["platform", "architecture"]) {
    if (left.host[key] !== budgets.reference_host[key]) {
      fail(`current host ${key} does not match the explicit budget host`);
    }
  }
  const reportIds = new Set(left.scenarios.map((scenario) => scenario.id));
  const budgetIds = new Set(budgets.scenarios.map((scenario) => scenario.id));
  if (reportIds.size !== budgetIds.size || [...reportIds].some((id) => !budgetIds.has(id))) {
    fail("benchmark budgets must exactly cover the current protocol scenarios");
  }
}

function readJson(file) {
  const absolute = path.resolve(root, file);
  try {
    return JSON.parse(fs.readFileSync(absolute, "utf8"));
  } catch (error) {
    fail(`cannot read ${file}: ${error.message}`);
  }
}

function parseArgs(argv) {
  const parsed = {};
  for (let index = 0; index < argv.length; index += 2) {
    const key = argv[index];
    const value = argv[index + 1];
    if (!key?.startsWith("--") || value === undefined) fail(`invalid arguments near ${key}`);
    parsed[key.slice(2)] = value;
  }
  return parsed;
}

function fail(message) {
  console.error(message);
  process.exit(1);
}
