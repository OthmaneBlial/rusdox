#!/usr/bin/env node

import fs from "node:fs";
import process from "node:process";

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

function workflowWarning(message) {
  const escaped = String(message).replaceAll("%", "%25").replaceAll("\r", "%0D").replaceAll("\n", "%0A");
  process.stdout.write(`::warning title=RusDox PR comment::${escaped}\n`);
}

async function githubRequest(path, options = {}) {
  const response = await fetch(`https://api.github.com${path}`, {
    ...options,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${process.env.GITHUB_TOKEN}`,
      "X-GitHub-Api-Version": "2022-11-28",
      "User-Agent": "rusdox-github-action",
      ...options.headers,
    },
  });
  if (!response.ok) throw new Error(`GitHub API ${response.status}: ${(await response.text()).slice(0, 500)}`);
  return response.status === 204 ? null : response.json();
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args["body-file"] || !fs.existsSync(args["body-file"])) return;
  if (!process.env.GITHUB_TOKEN || !process.env.GITHUB_REPOSITORY || !process.env.GITHUB_EVENT_PATH) {
    workflowWarning("Missing token, repository, or event context; parity comment was skipped.");
    return;
  }
  const event = JSON.parse(fs.readFileSync(process.env.GITHUB_EVENT_PATH, "utf8"));
  const pullNumber = event.pull_request?.number;
  if (!pullNumber) return;
  const [owner, repo] = process.env.GITHUB_REPOSITORY.split("/");
  const artifactLink = args["artifact-url"]
    ? `[Download the authenticated parity-report artifact](${args["artifact-url"]})`
    : "Parity-report upload was disabled or no report was produced.";
  const body = fs.readFileSync(args["body-file"], "utf8").replace("{{RUSDOX_ARTIFACT_LINK}}", artifactLink);

  try {
    const comments = await githubRequest(`/repos/${owner}/${repo}/issues/${pullNumber}/comments?per_page=100`);
    const existing = comments.find((comment) => comment.body?.includes("<!-- rusdox-parity -->"));
    const request = { method: existing ? "PATCH" : "POST", body: JSON.stringify({ body }) };
    const target = existing
      ? `/repos/${owner}/${repo}/issues/comments/${existing.id}`
      : `/repos/${owner}/${repo}/issues/${pullNumber}/comments`;
    await githubRequest(target, request);
    process.stdout.write(existing ? "Updated RusDox parity comment.\n" : "Created RusDox parity comment.\n");
  } catch (error) {
    workflowWarning(`${error.message}. Grant pull-requests: write to enable comments; verification results are unaffected.`);
  }
}

main().catch((error) => {
  workflowWarning(error.message);
});
