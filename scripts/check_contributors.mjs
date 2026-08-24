#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const result = spawnSync("git", ["shortlog", "-sne", "HEAD"], { cwd: root, encoding: "utf8" });
if (result.status !== 0) throw new Error(result.stderr || "git shortlog failed");
const contributorText = fs.readFileSync(path.join(root, "CONTRIBUTORS.md"), "utf8").toLocaleLowerCase("en");
const humans = result.stdout
  .split("\n")
  .map((line) => line.trim().replace(/^\d+\s+/, ""))
  .filter(Boolean)
  .filter((identity) => !/\[bot\]|dependabot/i.test(identity));

for (const identity of humans) {
  const name = identity.replace(/\s*<[^>]+>\s*$/, "").trim().toLocaleLowerCase("en");
  if (!contributorText.includes(name)) throw new Error(`CONTRIBUTORS.md is missing Git author ${identity}`);
}
console.log(`Contributor credit is current (${humans.length} human Git identity/identities).`);
