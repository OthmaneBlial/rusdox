#!/usr/bin/env node

import assert from "node:assert/strict";
import { annotationsForReport, renderSummary } from "./github_action.mjs";

const validation = {
  errors: 1,
  warnings: 1,
  config_issues: [],
  files: [
    {
      path: "/workspace/docs/invoice.yaml",
      issues: [
        {
          severity: "error",
          path: "blocks[2].rows[0]",
          message: "row width does not match",
          source: { line: 18, column: 7, end_line: 18, end_column: 21 },
        },
        {
          severity: "warning",
          path: "blocks[1].text",
          message: "text is blank",
          source: { line: 9, column: 5, end_line: 9, end_column: 9 },
        },
      ],
    },
  ],
};

const annotations = annotationsForReport(validation, "/workspace");
assert.deepEqual(
  annotations.map(({ level, file, line, column }) => ({ level, file, line, column })),
  [
    { level: "error", file: "docs/invoice.yaml", line: 18, column: 7 },
    { level: "warning", file: "docs/invoice.yaml", line: 9, column: 5 },
  ],
);

const failed = renderSummary(validation, null, "https://github.com/example/project/actions/runs/1");
assert.match(failed, /Rendering did not start/);
assert.match(failed, /docs\/invoice.yaml:18/);
assert.match(failed, /contains only check metadata/);

const passed = renderSummary(
  { errors: 0, warnings: 0, config_issues: [], files: [] },
  {
    passed: true,
    files: [{ source: "reports/monthly.yaml", checks: 19, failed_checks: 0 }],
  },
);
assert.match(passed, /✅ \*\*Passed\*\*/);
assert.match(passed, /reports\/monthly.yaml/);
assert.match(passed, /\| 19 \| 0 \|/);

console.log("GitHub Action annotations and summaries passed.");
