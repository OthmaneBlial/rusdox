#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const manifestPath = path.join(root, ".github/starter-issues.json");
const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
const publish = process.argv.includes("--publish");
const explicitRepo = valueAfter("--repo");

validateManifest(manifest);
if (!publish) {
  console.log(`Starter issue contract passed (${manifest.issues.length} issues).`);
  process.exit(0);
}

const repository = explicitRepo || runGh(["repo", "view", "--json", "nameWithOwner", "--jq", ".nameWithOwner"]).trim();
const existing = JSON.parse(runGh([
  "issue", "list", "--repo", repository, "--state", "all", "--limit", "200", "--json", "title,url",
]));
const byTitle = new Map(existing.map((issue) => [issue.title, issue.url]));

for (const issue of manifest.issues) {
  if (byTitle.has(issue.title)) {
    console.log(`exists ${byTitle.get(issue.title)}`);
    continue;
  }
  const body = renderIssue(issue, repository);
  const args = ["issue", "create", "--repo", repository, "--title", issue.title, "--body", body];
  for (const label of issue.labels) args.push("--label", label);
  const url = runGh(args).trim();
  console.log(`created ${url}`);
}

function validateManifest(value) {
  assert(value.schema_version === 1, "schema_version must be 1");
  assert(Array.isArray(value.issues) && value.issues.length === 10, "exactly 10 starter issues are required");
  const ids = new Set();
  const titles = new Set();
  for (const issue of value.issues) {
    assert(/^[a-z0-9][a-z0-9-]+$/.test(issue.id), `unsafe issue id ${issue.id}`);
    assert(!ids.has(issue.id), `duplicate issue id ${issue.id}`);
    assert(!titles.has(issue.title), `duplicate issue title ${issue.title}`);
    ids.add(issue.id);
    titles.add(issue.title);
    assert(issue.labels.includes("good first issue"), `${issue.id}: missing good first issue label`);
    assert(issue.labels.includes("help wanted"), `${issue.id}: missing help wanted label`);
    assert(issue.why?.length >= 40, `${issue.id}: why is too short`);
    assert(Array.isArray(issue.acceptance) && issue.acceptance.length === 3, `${issue.id}: require three acceptance criteria`);
    const fixture = path.resolve(root, issue.fixture);
    assert(fixture.startsWith(`${root}${path.sep}`), `${issue.id}: fixture escapes repository`);
    assert(fs.statSync(fixture).isFile(), `${issue.id}: missing fixture ${issue.fixture}`);
  }
}

function renderIssue(issue, repository) {
  const fixtureUrl = `https://github.com/${repository}/blob/main/${issue.fixture}`;
  return [
    "## Why",
    "",
    issue.why,
    "",
    "## Starter fixture",
    "",
    `[\`${issue.fixture}\`](${fixtureUrl})`,
    "",
    "Prepare an isolated copy with:",
    "",
    "```bash",
    `node scripts/contributor_lab.mjs prepare ${issue.id}`,
    "```",
    "",
    "## Acceptance criteria",
    "",
    ...issue.acceptance.map((criterion) => `- [ ] ${criterion}`),
    "",
    "## Scope",
    "",
    "This is intended to stay a small first contribution. Comment before expanding the contract or touching unrelated rendering behavior.",
  ].join("\n");
}

function runGh(args) {
  const result = spawnSync("gh", args, { cwd: root, encoding: "utf8", maxBuffer: 8 * 1024 * 1024 });
  if (result.status !== 0) {
    process.stderr.write(result.stdout || "");
    process.stderr.write(result.stderr || "");
    process.exit(result.status ?? 1);
  }
  return result.stdout;
}

function valueAfter(flag) {
  const index = process.argv.indexOf(flag);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}
