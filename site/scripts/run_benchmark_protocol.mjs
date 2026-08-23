#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const options = parseArgs(process.argv.slice(2));
const protocolPath = path.join(root, "benchmarks/protocol.json");
const protocol = JSON.parse(fs.readFileSync(protocolPath, "utf8"));
const binary = path.resolve(root, options.binary ?? "target/release/rusdox");
const output = path.resolve(root, options.output ?? "target/benchmarks/latest.json");
const iterations = integerOption(options.iterations, protocol.iterations, "iterations");
const warmup = integerOption(options.warmup, protocol.warmup, "warmup");

if (!fs.existsSync(binary)) {
  fail(`benchmark binary not found: ${binary}\nBuild it with: cargo build --release --locked --bin rusdox`);
}
if (!fs.existsSync("/usr/bin/time")) {
  fail("/usr/bin/time is required to capture peak resident memory");
}
if (!new Set(["linux", "darwin"]).has(process.platform)) {
  fail(`the benchmark protocol currently supports Linux and macOS, not ${process.platform}`);
}

const scenarios = [];
for (const tier of protocol.tiers) {
  for (const pipeline of protocol.pipelines) {
    scenarios.push(runScenario({
      id: `${tier.id}-${pipeline}`,
      label: `${tier.label} / ${pipeline}`,
      tier: tier.id,
      logicalPages: tier.logical_pages,
      pipeline,
      input: tier.input,
    }));
  }
}
scenarios.push(runScenario({
  id: protocol.existing_docx.id,
  label: protocol.existing_docx.label,
  tier: "existing-docx",
  logicalPages: null,
  pipeline: "existing-docx",
  input: protocol.existing_docx.input,
}));

const report = {
  schema_version: 1,
  protocol_version: protocol.protocol_version,
  generated_at_utc: new Date().toISOString(),
  git_commit: command("git", ["rev-parse", "HEAD"]),
  git_dirty: command("git", ["status", "--porcelain"]).length > 0,
  host: hostMetadata(),
  command: {
    runner: "node scripts/run_benchmark_protocol.mjs",
    binary: path.relative(root, binary),
    profile: "release",
    iterations,
    warmup,
  },
  scenarios,
};

fs.mkdirSync(path.dirname(output), { recursive: true });
const temporary = `${output}.tmp-${process.pid}`;
fs.writeFileSync(temporary, `${JSON.stringify(report, null, 2)}\n`);
fs.renameSync(temporary, output);
console.log(`Wrote ${path.relative(root, output)} (${scenarios.length} isolated scenarios).`);

function runScenario(scenario) {
  const input = path.resolve(root, scenario.input);
  if (!fs.existsSync(input)) fail(`benchmark input not found: ${input}`);

  const commandFlags = [
    "bench",
    scenario.input,
    "--pipeline",
    scenario.pipeline,
    "--iterations",
    String(iterations),
    "--warmup",
    String(warmup),
    "--format",
    "json",
  ];
  const memoryFile = path.join(os.tmpdir(), `rusdox-bench-memory-${process.pid}-${scenario.id}.txt`);
  const timeArgs = process.platform === "linux"
    ? ["-f", "%M", "-o", memoryFile, binary, ...commandFlags]
    : ["-l", binary, ...commandFlags];
  process.stderr.write(`benchmark ${scenario.id}\n`);
  const run = spawnSync("/usr/bin/time", timeArgs, {
    cwd: root,
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
  });
  if (run.status !== 0) {
    fail(`scenario ${scenario.id} failed (${run.status}):\n${run.stderr || run.stdout}`);
  }

  let result;
  try {
    result = JSON.parse(run.stdout);
  } catch (error) {
    fail(`scenario ${scenario.id} returned invalid JSON: ${error.message}\n${run.stdout}`);
  }
  const peakMemoryBytes = peakMemory(run.stderr, memoryFile);
  fs.rmSync(memoryFile, { force: true });

  return {
    id: scenario.id,
    label: scenario.label,
    tier: scenario.tier,
    logical_pages: scenario.logicalPages,
    pipeline: scenario.pipeline,
    input: scenario.input,
    input_sha256: sha256(input),
    input_bytes: fs.statSync(input).size,
    command_flags: commandFlags,
    peak_memory_bytes: peakMemoryBytes,
    result,
  };
}

function peakMemory(stderr, memoryFile) {
  if (process.platform === "linux") {
    const kib = Number.parseInt(fs.readFileSync(memoryFile, "utf8").trim(), 10);
    if (!Number.isFinite(kib) || kib <= 0) fail("GNU time did not report peak resident memory");
    return kib * 1024;
  }
  const match = stderr.match(/(\d+)\s+maximum resident set size/);
  if (!match) fail(`BSD time did not report peak resident memory:\n${stderr}`);
  return Number.parseInt(match[1], 10);
}

function hostMetadata() {
  const cpu = os.cpus()[0]?.model?.trim() || "unknown";
  return {
    platform: process.platform,
    os_release: os.release(),
    os_version: typeof os.version === "function" ? os.version() : os.release(),
    architecture: os.arch(),
    cpu,
    logical_cpu_count: os.cpus().length,
    total_memory_bytes: os.totalmem(),
    rustc: command("rustc", ["--version", "--verbose"]),
    cargo: command("cargo", ["--version"]),
  };
}

function command(program, args) {
  const result = spawnSync(program, args, { cwd: root, encoding: "utf8" });
  if (result.status !== 0) fail(`${program} ${args.join(" ")} failed: ${result.stderr}`);
  return result.stdout.trim();
}

function sha256(file) {
  return crypto.createHash("sha256").update(fs.readFileSync(file)).digest("hex");
}

function integerOption(raw, fallback, name) {
  const value = raw === undefined ? fallback : Number.parseInt(raw, 10);
  if (!Number.isInteger(value) || value < 0 || (name === "iterations" && value === 0)) {
    fail(`--${name} must be ${name === "iterations" ? "positive" : "non-negative"}`);
  }
  return value;
}

function parseArgs(args) {
  const parsed = {};
  for (let index = 0; index < args.length; index += 1) {
    const key = args[index];
    if (!key.startsWith("--")) fail(`unexpected argument: ${key}`);
    const value = args[index + 1];
    if (value === undefined || value.startsWith("--")) fail(`${key} requires a value`);
    parsed[key.slice(2).replaceAll("-", "_")] = value;
    index += 1;
  }
  return parsed;
}

function fail(message) {
  console.error(message);
  process.exit(1);
}
