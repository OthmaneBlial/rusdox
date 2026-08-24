"use strict";

const assert = require("node:assert/strict");
const service = require("../language-service");

assert(service.BLOCK_TYPES.includes("when"));
assert(service.completionItems("    type: ").some((item) => item.label === "table"));
assert.equal(service.hoverFor("version").includes("version: 1"), true);

const clean = service.analyze("version: 1\nblocks:\n  - type: body\n    text: '{{ customer.name | upper }}'\n");
assert.deepEqual(clean, []);

const invalid = service.analyze("blocks:\n  - type: mystery\n    text: '{{ value | shell }}'\n");
assert.equal(invalid.filter((item) => item.severity === "error").length, 2);
assert.equal(invalid.filter((item) => item.severity === "warning").length, 1);

console.log("RusDox VS Code language service tests passed.");
