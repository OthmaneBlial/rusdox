#!/usr/bin/env node

import { spawn } from "node:child_process";
import path from "node:path";
import process from "node:process";

const binary = process.env.RUSDOX_BIN || "rusdox";
const outputRoot = path.resolve(process.argv[2] || "target/integration-node");
const child = spawn(binary, ["serve", "stdio", "--output-root", outputRoot, "--max-requests", "1"], {
  stdio: ["pipe", "pipe", "inherit"],
});
const request = {
  protocol_version: 1,
  request_id: "node-example",
  operation: "render",
  source: {
    kind: "inline",
    format: "yaml",
    content: "version: 1\noutput_name: node-report\nblocks:\n  - type: title\n    text: Node integration\n  - type: body\n    text: One local JSON line produced both editable DOCX and native PDF.\n",
  },
  output: { directory: "node", name: "node-report", pdf: true },
};
child.stdin.end(`${JSON.stringify(request)}\n`);
let stdout = "";
child.stdout.setEncoding("utf8");
child.stdout.on("data", (chunk) => { stdout += chunk; });
child.on("error", (error) => {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
});
child.on("close", (status) => {
  if (status !== 0) process.exit(status ?? 1);
  const response = JSON.parse(stdout);
  if (!response.ok) throw new Error(response.error?.message || "RusDox request failed");
  process.stdout.write(`${JSON.stringify(response, null, 2)}\n`);
});
