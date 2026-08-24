#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const check = process.argv.includes("--check");
const history = read("benchmarks/history.json");
const baseline = read("benchmarks/baselines/ubuntu-latest.json");
const budgets = read("benchmarks/budgets.json");
const budgetById = new Map(budgets.scenarios.map((item) => [item.id, item]));
const rows = baseline.scenarios.map((scenario) => {
  const budget = budgetById.get(scenario.id);
  if (!budget) throw new Error(`missing benchmark budget for ${scenario.id}`);
  const runtime = scenario.result.total_ms.median;
  const memory = scenario.peak_memory_bytes;
  return {
    id: scenario.id,
    runtime_ms: runtime,
    runtime_budget_ms: budget.max_median_ms,
    runtime_headroom_ratio: 1 - runtime / budget.max_median_ms,
    peak_memory_bytes: memory,
    memory_budget_bytes: budget.max_peak_memory_bytes,
    memory_headroom_ratio: 1 - memory / budget.max_peak_memory_bytes,
    passed: runtime <= budget.max_median_ms && memory <= budget.max_peak_memory_bytes,
  };
});
if (rows.length !== budgets.scenarios.length) throw new Error("baseline and budgets must have identical scenario coverage");

const summary = {
  schema_version: 1,
  generated_from: ["benchmarks/history.json", "benchmarks/baselines/ubuntu-latest.json", "benchmarks/budgets.json"],
  baseline_at: baseline.generated_at_utc,
  reviewed_at: budgets.reviewed_at,
  reference_host: budgets.reference_host,
  passed: rows.every((row) => row.passed),
  scenarios: rows,
};
const json = `${JSON.stringify(summary, null, 2)}\n`;
const html = render(summary, history);
writeOrCheck("benchmarks/dashboard.json", json);
writeOrCheck("site/benchmarks/index.html", html);
console.log(check ? `Benchmark dashboard is current (${rows.length} budgets).` : `Generated benchmark dashboard (${rows.length} budgets).`);

function render(data, historyData) {
  const passed = data.scenarios.filter((row) => row.passed).length;
  const worstRuntime = data.scenarios.reduce((a, b) => a.runtime_headroom_ratio < b.runtime_headroom_ratio ? a : b);
  const worstMemory = data.scenarios.reduce((a, b) => a.memory_headroom_ratio < b.memory_headroom_ratio ? a : b);
  const tableRows = data.scenarios.map((row) => `<tr>
    <th scope="row"><code>${escape(row.id)}</code></th>
    <td>${formatMs(row.runtime_ms)}</td><td>${formatMs(row.runtime_budget_ms)}</td><td>${bar(row.runtime_ms / row.runtime_budget_ms)}</td>
    <td>${formatBytes(row.peak_memory_bytes)}</td><td>${formatBytes(row.memory_budget_bytes)}</td><td><span class="status ${row.passed ? "pass" : "fail"}">${row.passed ? "Within budget" : "Exceeded"}</span></td>
  </tr>`).join("");
  const runs = historyData.runs.map((run) => `<article><p>${escape(run.generated_at_utc.slice(0, 10))}</p><strong>${escape(run.host.platform)}/${escape(run.host.architecture)}</strong><span>${escape(run.host.cpu)}</span><small>${run.scenarios.length} isolated scenarios · commit ${escape(run.git_commit.slice(0, 7))}</small></article>`).join("");
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>RusDox performance budgets</title><meta name="description" content="Long-term RusDox runtime and memory budgets backed by reproducible benchmark evidence."><link rel="canonical" href="https://othmaneblial.github.io/rusdox/benchmarks/"><style>
  :root{color-scheme:dark;--bg:#07111e;--panel:#0d1b2b;--line:#23354b;--text:#f4f0e8;--muted:#a7b4c5;--mint:#5eead4;--amber:#fbbf24;--red:#fb7185}*{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at 82% 0,#16334a 0,transparent 36rem),var(--bg);color:var(--text);font:15px/1.55 Inter,system-ui,sans-serif}main{width:min(1180px,calc(100% - 32px));margin:auto;padding:48px 0 80px}a{color:var(--mint)}header{display:grid;gap:16px;padding:44px;border:1px solid var(--line);border-radius:28px;background:linear-gradient(145deg,#11283a,#0a1725)}.eyebrow{margin:0;color:var(--mint);font:700 12px ui-monospace,monospace;text-transform:uppercase;letter-spacing:.13em}h1{font-size:clamp(38px,7vw,76px);line-height:.98;max-width:850px;margin:0}header>p:not(.eyebrow){max-width:760px;color:var(--muted);font-size:18px}.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.stat{padding:20px;border:1px solid var(--line);border-radius:18px;background:#081522}.stat strong{display:block;font-size:28px}.stat span{color:var(--muted)}section{margin-top:48px}h2{font-size:30px}p{color:var(--muted)}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:20px;background:var(--panel)}table{border-collapse:collapse;width:100%;min-width:920px}th,td{text-align:left;padding:14px 16px;border-bottom:1px solid var(--line)}thead th{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.08em}.meter{display:flex;align-items:center;gap:8px}.track{width:90px;height:7px;border-radius:9px;background:#23354b;overflow:hidden}.fill{height:100%;background:var(--mint)}.status{font-weight:700}.pass{color:var(--mint)}.fail{color:var(--red)}.runs{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px}.runs article{display:grid;gap:4px;padding:20px;border:1px solid var(--line);border-radius:18px;background:var(--panel)}.runs p{margin:0;color:var(--amber);font:12px ui-monospace,monospace}.runs span,.runs small{color:var(--muted)}footer{margin-top:48px;padding-top:24px;border-top:1px solid var(--line);display:flex;gap:20px;flex-wrap:wrap}@media(max-width:720px){main{width:min(100% - 20px,1180px);padding-top:10px}header{padding:28px 22px}.stats{grid-template-columns:1fr}h1{font-size:44px}}
  </style></head><body><main><header><p class="eyebrow">Performance contract · reviewed ${escape(data.reviewed_at)}</p><h1>Fast is a budget, not a screenshot.</h1><p>Every scheduled Ubuntu run is compared with a comparable baseline, a relative regression floor, and these absolute runtime and peak-memory ceilings. A threshold change requires review and history.</p><div class="stats"><div class="stat"><strong>${passed}/${data.scenarios.length}</strong><span>scenarios within budget</span></div><div class="stat"><strong>${percent(worstRuntime.runtime_headroom_ratio)}</strong><span>tightest runtime headroom · ${escape(worstRuntime.id)}</span></div><div class="stat"><strong>${percent(worstMemory.memory_headroom_ratio)}</strong><span>tightest memory headroom · ${escape(worstMemory.id)}</span></div></div></header><section><p class="eyebrow">Reference host</p><h2>Explicit ceilings for ${escape(data.reference_host.runner)}</h2><p>Baseline ${escape(data.baseline_at)} · ${escape(data.reference_host.platform)}/${escape(data.reference_host.architecture)}. Cross-host observations remain visible below but are never compared as regressions.</p><div class="table-wrap" tabindex="0" aria-label="Scrollable performance budget table"><table><thead><tr><th>Scenario</th><th>Median</th><th>Runtime budget</th><th>Usage</th><th>Peak memory</th><th>Memory budget</th><th>Result</th></tr></thead><tbody>${tableRows}</tbody></table></div></section><section><p class="eyebrow">Long-term evidence</p><h2>Checked-in runs, separated by host.</h2><div class="runs">${runs}</div></section><footer><a href="../docs/performance.html">Methodology and limitations</a><a href="dashboard.json">Machine-readable dashboard</a><a href="../benchmarks/history.json">Full history</a><a href="https://github.com/OthmaneBlial/rusdox/actions/workflows/benchmark.yml">Scheduled workflow</a></footer></main></body></html>\n`;
}

function bar(ratio) { const value=Math.max(0,Math.min(1,ratio)); return `<div class="meter"><div class="track"><div class="fill" style="width:${(value*100).toFixed(1)}%"></div></div><span>${(ratio*100).toFixed(0)}%</span></div>`; }
function percent(value) { return `${Math.round(value * 100)}%`; }
function formatMs(value) { return value >= 1000 ? `${(value/1000).toFixed(2)} s` : `${value.toFixed(value < 10 ? 2 : 0)} ms`; }
function formatBytes(value) { return `${(value/1048576).toFixed(value < 104857600 ? 1 : 0)} MiB`; }
function escape(value) { return String(value).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;"); }
function read(relative) { return JSON.parse(fs.readFileSync(path.join(root, relative), "utf8")); }
function writeOrCheck(relative, content) { const file=path.join(root,relative); if(check){if(!fs.existsSync(file)||fs.readFileSync(file,"utf8")!==content)throw new Error(`${relative} is stale`);}else{fs.mkdirSync(path.dirname(file),{recursive:true});fs.writeFileSync(file,content);} }
