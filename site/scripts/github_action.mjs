#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

export function annotationsForReport(report, workspace = process.cwd()) {
  const annotations = [];
  for (const issue of report.config_issues ?? []) {
    annotations.push(annotationForIssue(issue, "rusdox.toml", workspace));
  }
  for (const file of report.files ?? []) {
    for (const issue of file.issues ?? []) {
      annotations.push(annotationForIssue(issue, file.path, workspace));
    }
  }
  return annotations;
}

export function renderSummary(validation, verification, runUrl = "") {
  const validationPassed = Number(validation.errors ?? 0) === 0;
  const passed = validationPassed && Boolean(verification?.passed);
  const lines = [
    "<!-- rusdox-parity -->",
    "## RusDox document verification",
    "",
    `${passed ? "✅" : "❌"} **${passed ? "Passed" : "Needs attention"}** — ${validation.errors ?? 0} validation error(s), ${validation.warnings ?? 0} warning(s).`,
    "",
  ];

  if (verification?.files?.length) {
    lines.push("| Spec | Parity checks | Failed |", "|---|---:|---:|");
    for (const file of verification.files) {
      lines.push(
        `| \`${escapeMarkdownCell(file.source)}\` | ${Number(file.checks ?? 0)} | ${Number(file.failed_checks ?? 0)} |`,
      );
    }
    lines.push("");
  } else if (!validationPassed) {
    const issues = annotationsForReport(validation).filter((entry) => entry.level === "error");
    lines.push("Rendering did not start because validation failed.", "");
    for (const issue of issues.slice(0, 10)) {
      lines.push(`- \`${escapeMarkdownCell(issue.file)}:${issue.line ?? 1}\` ${escapeMarkdownCell(issue.message)}`);
    }
    if (issues.length > 10) lines.push(`- …and ${issues.length - 10} more error(s).`);
    lines.push("");
  }

  lines.push(
    "The PR comment contains only check metadata. Generated DOCX/PDF files are not uploaded by this action; parity reports stay inside the repository's GitHub Actions run.",
  );
  if (runUrl) lines.push("", `[Open the workflow run](${runUrl})`);
  lines.push("", "{{RUSDOX_ARTIFACT_LINK}}", "");
  return lines.join("\n");
}

function annotationForIssue(issue, file, workspace) {
  const source = issue.source ?? {};
  return {
    level: issue.severity === "warning" ? "warning" : "error",
    file: relativeWorkspacePath(file, workspace),
    line: source.line,
    endLine: source.end_line,
    column: source.column,
    endColumn: source.end_column,
    title: `RusDox ${issue.severity ?? "error"}`,
    message: `${issue.path}: ${issue.message}`,
  };
}

function relativeWorkspacePath(value, workspace) {
  const resolved = path.resolve(value);
  const relative = path.relative(path.resolve(workspace), resolved);
  if (relative && !relative.startsWith("..") && !path.isAbsolute(relative)) return relative;
  return value;
}

function emitAnnotation(annotation) {
  const properties = [
    ["file", annotation.file],
    ["line", annotation.line],
    ["endLine", annotation.endLine],
    ["col", annotation.column],
    ["endColumn", annotation.endColumn],
    ["title", annotation.title],
  ]
    .filter(([, value]) => value !== undefined && value !== null && value !== "")
    .map(([key, value]) => `${key}=${escapeWorkflowProperty(String(value))}`)
    .join(",");
  process.stdout.write(`::${annotation.level} ${properties}::${escapeWorkflowMessage(annotation.message)}\n`);
}

function runRusDox(binary, args) {
  const result = spawnSync(binary, args, {
    cwd: process.cwd(),
    encoding: "utf8",
    maxBuffer: 32 * 1024 * 1024,
  });
  if (result.error) throw result.error;
  return result;
}

function parseJsonOutput(result, label) {
  try {
    return JSON.parse(result.stdout);
  } catch (error) {
    if (result.stdout) process.stdout.write(result.stdout);
    if (result.stderr) process.stderr.write(result.stderr);
    throw new Error(`${label} did not return valid JSON: ${error.message}`);
  }
}

function writeActionOutput(name, value) {
  if (process.env.GITHUB_OUTPUT) {
    fs.appendFileSync(process.env.GITHUB_OUTPUT, `${name}=${String(value).replaceAll("\n", "%0A")}\n`);
  }
}

function writeSummary(markdown, outputRoot) {
  fs.mkdirSync(outputRoot, { recursive: true });
  fs.writeFileSync(path.join(outputRoot, "rusdox-comment.md"), markdown);
  if (process.env.GITHUB_STEP_SUMMARY) {
    fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, `${markdown.replace("{{RUSDOX_ARTIFACT_LINK}}", "")}\n`);
  }
}

function escapeWorkflowMessage(value) {
  return String(value).replaceAll("%", "%25").replaceAll("\r", "%0D").replaceAll("\n", "%0A");
}

function escapeWorkflowProperty(value) {
  return escapeWorkflowMessage(value).replaceAll(":", "%3A").replaceAll(",", "%2C");
}

function escapeMarkdownCell(value) {
  return String(value).replaceAll("|", "\\|").replaceAll("\n", " ").replaceAll("\r", " ");
}

function parseArgs(argv) {
  const values = {};
  for (let index = 0; index < argv.length; index += 2) {
    const key = argv[index];
    const value = argv[index + 1];
    if (!key?.startsWith("--") || value === undefined) throw new Error(`invalid argument near ${key ?? "end"}`);
    values[key.slice(2)] = value;
  }
  return values;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  for (const required of ["binary", "input", "output-root", "visual-threshold"]) {
    if (!args[required]) throw new Error(`--${required} is required`);
  }
  const outputRoot = path.resolve(args["output-root"]);
  const reportsPath = path.join(outputRoot, "reports");
  const validationArgs = ["validate", args.input, "--format", "json"];
  if (args.config) validationArgs.push("--config", args.config);

  process.stdout.write(`RusDox: validating ${args.input}\n`);
  const validationResult = runRusDox(args.binary, validationArgs);
  const validation = parseJsonOutput(validationResult, "rusdox validate");
  for (const annotation of annotationsForReport(validation)) emitAnnotation(annotation);

  const runUrl = process.env.GITHUB_SERVER_URL && process.env.GITHUB_REPOSITORY && process.env.GITHUB_RUN_ID
    ? `${process.env.GITHUB_SERVER_URL}/${process.env.GITHUB_REPOSITORY}/actions/runs/${process.env.GITHUB_RUN_ID}`
    : "";

  if (validationResult.status !== 0 || Number(validation.errors ?? 0) > 0) {
    writeSummary(renderSummary(validation, null, runUrl), outputRoot);
    writeActionOutput("passed", "false");
    writeActionOutput("reports-path", reportsPath);
    process.exitCode = validationResult.status || 1;
    return;
  }

  const verifyArgs = ["verify", args.input, "--output-root", outputRoot, "--format", "json"];
  if (args.config) verifyArgs.push("--config", args.config);
  if (args["visual-baseline"]) verifyArgs.push("--visual-baseline", args["visual-baseline"]);
  verifyArgs.push("--visual-threshold", args["visual-threshold"]);

  process.stdout.write(`RusDox: rendering DOCX/PDF and parity evidence\n`);
  const verificationResult = runRusDox(args.binary, verifyArgs);
  const verification = parseJsonOutput(verificationResult, "rusdox verify");
  const passed = verificationResult.status === 0 && verification.passed === true;
  writeSummary(renderSummary(validation, verification, runUrl), outputRoot);
  writeActionOutput("passed", passed ? "true" : "false");
  writeActionOutput("reports-path", reportsPath);
  if (verificationResult.stderr) process.stderr.write(verificationResult.stderr);
  if (!passed) process.exitCode = verificationResult.status || 2;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  main().catch((error) => {
    process.stderr.write(`RusDox action error: ${error.message}\n`);
    writeActionOutput("passed", "false");
    process.exitCode = 1;
  });
}
