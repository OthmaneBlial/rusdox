#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const resultsDir = path.join(root, "benchmarks/results");
const historyPath = path.join(root, "benchmarks/history.json");
const chartPath = path.join(root, "assets/benchmark-history.svg");
const reports = fs.existsSync(resultsDir)
  ? fs.readdirSync(resultsDir)
      .filter((name) => name.endsWith(".json"))
      .sort()
      .map((name) => JSON.parse(fs.readFileSync(path.join(resultsDir, name), "utf8")))
  : [];

if (reports.length === 0) {
  throw new Error("benchmarks/results must contain at least one protocol report");
}

const history = {
  schema_version: 1,
  generated_from: "benchmarks/results/*.json",
  runs: reports.map((report) => ({
    generated_at_utc: report.generated_at_utc,
    git_commit: report.git_commit,
    protocol_version: report.protocol_version,
    host: report.host,
    scenarios: report.scenarios.map((scenario) => ({
      id: scenario.id,
      tier: scenario.tier,
      pipeline: scenario.pipeline,
      input_sha256: scenario.input_sha256,
      median_total_ms: scenario.result.total_ms.median,
      peak_memory_bytes: scenario.peak_memory_bytes,
      docx_bytes: scenario.result.docx_bytes.median,
      pdf_bytes: scenario.result.pdf_bytes.median,
    })),
  })),
};

writeAtomic(historyPath, `${JSON.stringify(history, null, 2)}\n`);
writeAtomic(chartPath, renderChart(history.runs));
console.log(`Rendered ${path.relative(root, historyPath)} and ${path.relative(root, chartPath)}.`);

function renderChart(runs) {
  const width = 1200;
  const height = 680;
  const margin = { left: 96, right: 40, top: 112, bottom: 96 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const ids = ["small-dual", "medium-dual", "pages-1000-dual", "external-docx-open-save"];
  const labels = ["Small dual", "Medium dual", "1,000-page dual", "Existing DOCX"];
  const colors = ["#2dd4bf", "#60a5fa", "#f59e0b", "#f472b6"];
  const values = runs.flatMap((run) => ids.map((id) => scenarioValue(run, id))).filter(Number.isFinite);
  const maxLog = Math.max(1, ...values.map((value) => Math.log10(value + 1)));
  const xInset = Math.min(96, plotWidth / 4);
  const x = (index) => margin.left + (runs.length === 1
    ? plotWidth / 2
    : xInset + ((plotWidth - xInset * 2) * index) / (runs.length - 1));
  const y = (value) => margin.top + plotHeight - (Math.log10(value + 1) / maxLog) * plotHeight;
  const grid = [0, 0.25, 0.5, 0.75, 1].map((ratio) => {
    const gy = margin.top + plotHeight * (1 - ratio);
    const value = 10 ** (maxLog * ratio) - 1;
    return `<line x1="${margin.left}" y1="${gy}" x2="${width - margin.right}" y2="${gy}" class="grid"/><text x="${margin.left - 16}" y="${gy + 5}" text-anchor="end" class="axis">${formatMs(value)}</text>`;
  }).join("");
  const series = ids.map((id, seriesIndex) => {
    const points = runs.map((run, index) => ({ x: x(index), y: y(scenarioValue(run, id)), value: scenarioValue(run, id), host: hostKey(run) }));
    const pathData = points.map((point, index) => `${index === 0 || point.host !== points[index - 1].host ? "M" : "L"}${point.x.toFixed(1)},${point.y.toFixed(1)}`).join(" ");
    const circles = points.map((point) => `<circle cx="${point.x}" cy="${point.y}" r="6" fill="${colors[seriesIndex]}"><title>${labels[seriesIndex]}: ${point.value.toFixed(2)} ms</title></circle>`).join("");
    return `<path d="${pathData}" fill="none" stroke="${colors[seriesIndex]}" stroke-width="3"/>${circles}`;
  }).join("");
  const dates = runs.map((run, index) => `<text x="${x(index)}" y="${height - margin.bottom + 28}" text-anchor="middle" class="axis">${escapeXml(run.generated_at_utc.slice(0, 16).replace("T", " "))}Z</text><text x="${x(index)}" y="${height - margin.bottom + 49}" text-anchor="middle" class="note">${escapeXml(run.host.platform)}/${escapeXml(run.host.architecture)}</text>`).join("");
  const legend = labels.map((label, index) => `<g transform="translate(${margin.left + index * 250},72)"><circle r="6" fill="${colors[index]}"/><text x="14" y="5" class="legend">${label}</text></g>`).join("");
  const hostCount = new Set(runs.map(hostKey)).size;

  return `<svg xmlns="http://www.w3.org/2000/svg" role="img" aria-labelledby="title desc" viewBox="0 0 ${width} ${height}">
  <title id="title">RusDox reproducible benchmark history</title>
  <desc id="desc">Median end-to-end runtime by dated protocol run on a logarithmic millisecond scale. Data is generated from checked-in JSON reports.</desc>
  <style>.bg{fill:#08111f}.grid{stroke:#233248;stroke-width:1}.axis{fill:#94a3b8;font:14px ui-monospace,SFMono-Regular,Menlo,monospace}.title{fill:#f8fafc;font:700 30px system-ui,sans-serif}.subtitle{fill:#94a3b8;font:15px system-ui,sans-serif}.legend{fill:#dbeafe;font:15px system-ui,sans-serif}.note{fill:#64748b;font:13px system-ui,sans-serif}</style>
  <rect class="bg" width="${width}" height="${height}" rx="24"/>
  <text x="${margin.left}" y="42" class="title">Reproducible benchmark history</text>
  <text x="${margin.left}" y="102" class="subtitle">Median total runtime · log scale · lower is better</text>
  ${legend}${grid}${series}${dates}
  <text x="${margin.left}" y="${height - 18}" class="note">${runs.length} checked-in run${runs.length === 1 ? "" : "s"} · ${hostCount} host ${hostCount === 1 ? "identity" : "identities"} · lines only connect comparable hosts</text>
</svg>
`;
}

function scenarioValue(run, id) {
  return run.scenarios.find((scenario) => scenario.id === id)?.median_total_ms ?? 0;
}

function hostKey(run) {
  return [run.host.platform, run.host.architecture, run.host.cpu].join("|");
}

function formatMs(value) {
  if (value >= 1000) return `${(value / 1000).toFixed(1)}s`;
  if (value >= 10) return `${Math.round(value)}ms`;
  return `${value.toFixed(1)}ms`;
}

function escapeXml(value) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function writeAtomic(destination, content) {
  fs.mkdirSync(path.dirname(destination), { recursive: true });
  const temporary = `${destination}.tmp-${process.pid}`;
  fs.writeFileSync(temporary, content);
  fs.renameSync(temporary, destination);
}
