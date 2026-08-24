#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const manifest = JSON.parse(fs.readFileSync(path.join(root, ".github/starter-issues.json"), "utf8"));
const [command, ...args] = process.argv.slice(2);

switch (command) {
  case "list":
    list(args.includes("--json"));
    break;
  case "prepare":
    prepare(args);
    break;
  case "visual-diff":
    visualDiff(args);
    break;
  default:
    usage();
    process.exitCode = 2;
}

function list(json) {
  if (json) {
    console.log(JSON.stringify(manifest.issues.map(({ id, title, fixture }) => ({ id, title, fixture })), null, 2));
    return;
  }
  for (const issue of manifest.issues) console.log(`${issue.id}\t${issue.fixture}\t${issue.title}`);
}

function prepare(values) {
  const id = values.find((value) => !value.startsWith("--"));
  const issue = manifest.issues.find((candidate) => candidate.id === id);
  if (!issue) throw new Error(`unknown starter issue '${id ?? ""}'; run contributor_lab.mjs list`);
  const outputRoot = path.resolve(valueAfter(values, "--output-root") || path.join(root, "target/contributor-lab"));
  const destination = path.join(outputRoot, issue.id);
  fs.mkdirSync(destination, { recursive: true });
  fs.copyFileSync(path.join(root, issue.fixture), path.join(destination, path.basename(issue.fixture)));
  fs.writeFileSync(
    path.join(destination, "TASK.md"),
    [
      `# ${issue.title}`,
      "",
      issue.why,
      "",
      `Source fixture: \`${issue.fixture}\``,
      "",
      "## Acceptance criteria",
      "",
      ...issue.acceptance.map((criterion) => `- [ ] ${criterion}`),
      "",
    ].join("\n"),
  );
  console.log(destination);
}

function visualDiff(values) {
  const spec = requiredValue(values, "--spec");
  const baseline = requiredValue(values, "--baseline");
  const outputRoot = valueAfter(values, "--output-root") || "target/contributor-visual-diff";
  const binary = valueAfter(values, "--binary") || path.join(root, "target/debug/rusdox");
  const threshold = valueAfter(values, "--threshold") || "0";
  const command = [
    binary,
    "verify",
    spec,
    "--output-root",
    outputRoot,
    "--visual-baseline",
    baseline,
    "--visual-threshold",
    threshold,
  ];
  if (values.includes("--dry-run")) {
    console.log(JSON.stringify({ command }, null, 2));
    return;
  }
  const result = spawnSync(command[0], command.slice(1), { cwd: root, stdio: "inherit" });
  if (result.status !== 0) process.exit(result.status ?? 1);
}

function requiredValue(values, flag) {
  const value = valueAfter(values, flag);
  if (!value) throw new Error(`${flag} is required`);
  return value;
}

function valueAfter(values, flag) {
  const index = values.indexOf(flag);
  return index >= 0 ? values[index + 1] : undefined;
}

function usage() {
  console.error("Usage:");
  console.error("  node scripts/contributor_lab.mjs list [--json]");
  console.error("  node scripts/contributor_lab.mjs prepare ID [--output-root DIR]");
  console.error("  node scripts/contributor_lab.mjs visual-diff --spec FILE --baseline DIR [--threshold N] [--output-root DIR] [--binary FILE] [--dry-run]");
}
