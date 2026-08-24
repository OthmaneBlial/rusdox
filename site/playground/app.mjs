import { cliCommand, downloadableName, parseSpecPreview, previewBlockModel } from "./core.mjs";

const EXAMPLES = [
  example("board-report", "Board Report", "board_report.yaml"),
  example("executive-dashboard", "Executive Dashboard", "executive_dashboard.yaml"),
  example("product-launch-brief", "Product Launch Brief", "product_launch_brief.yaml"),
  example("talent-profile", "Talent Profile", "talent_profile.yaml"),
  example("invoice", "Invoice", "invoice.yaml"),
  example("meeting-notes", "Meeting Notes", "meeting_notes.yaml"),
];

const elements = {
  picker: document.querySelector("[data-example-picker]"),
  editor: document.querySelector("[data-editor]"),
  preview: document.querySelector("[data-preview]"),
  diagnostics: document.querySelector("[data-diagnostics]"),
  state: document.querySelector("[data-state]"),
  cli: document.querySelector("[data-cli]"),
  pdf: document.querySelector("[data-pdf]"),
  docx: document.querySelector("[data-docx]"),
  parity: document.querySelector("[data-parity]"),
};

let selected = null;
let verifiedSource = "";

for (const item of EXAMPLES) {
  const option = document.createElement("option");
  option.value = item.id;
  option.textContent = item.title;
  elements.picker.append(option);
}

elements.picker.addEventListener("change", () => loadExample(elements.picker.value));
elements.editor.addEventListener("input", render);
document.querySelector("[data-download-yaml]").addEventListener("click", downloadYaml);
document.querySelector("[data-copy-cli]").addEventListener("click", copyCli);
document.querySelectorAll("[data-insert]").forEach((button) => {
  button.addEventListener("click", () => insertBlock(button.dataset.insert));
});

const requested = new URLSearchParams(location.search).get("example");
const initial = EXAMPLES.some((item) => item.id === requested) ? requested : EXAMPLES[0].id;
elements.picker.value = initial;
loadExample(initial);

async function loadExample(id) {
  selected = EXAMPLES.find((item) => item.id === id) || EXAMPLES[0];
  setLoading(true);
  try {
    const response = await fetch("../examples/" + selected.file, { cache: "no-store", credentials: "same-origin" });
    if (!response.ok) throw new Error("HTTP " + response.status);
    verifiedSource = await response.text();
    elements.editor.value = verifiedSource;
    history.replaceState(null, "", "?example=" + selected.id);
    render();
  } catch (error) {
    elements.diagnostics.textContent = "Could not load this checked-in example: " + error.message;
  } finally {
    setLoading(false);
  }
}

function render() {
  const source = elements.editor.value;
  const model = parseSpecPreview(source);
  const pristine = source === verifiedSource;
  elements.preview.replaceChildren(...renderBlocks(model));
  elements.diagnostics.textContent = model.diagnostics.length ? model.diagnostics.join(" ") : "Structural preview updated locally.";
  elements.state.textContent = pristine ? "Verified checked-in source" : "Edited locally — verified downloads paused";
  elements.state.classList.toggle("is-edited", !pristine);
  elements.cli.textContent = cliCommand("examples/" + selected.file, model.outputName);
  updateArtifact(elements.pdf, "../rendered/" + selected.id + ".pdf", pristine);
  updateArtifact(elements.docx, "../generated/" + selected.id + ".docx", pristine);
  updateArtifact(elements.parity, "../parity/" + selected.id + "-parity.html", pristine);
}

function renderBlocks(model) {
  if (!model.blocks.length) return [messageNode("Add a title, body, bullet list, metric set, or table to begin.")];
  return model.blocks.map((block) => {
    const view = previewBlockModel(block);
    if (view.kind === "metrics" || view.kind === "pairs") {
      const group = node("div", view.kind === "metrics" ? "preview-metrics" : "preview-pairs");
      view.items.forEach((item) => {
        const card = node("div", "preview-pair");
        card.append(node("span", "", item.label || "Item"), node("strong", "", item.value || ""));
        group.append(card);
      });
      return group;
    }
    if (view.kind === "bullets") {
      const list = node("ul", "preview-list");
      view.items.forEach((item) => list.append(node("li", "", item)));
      return list;
    }
    if (view.kind === "table") {
      const box = node("div", "preview-table");
      view.cells.forEach((cell) => box.append(node("span", "", cell)));
      if (!view.cells.length) box.append(node("span", "", "Table structure"));
      return box;
    }
    const tag = ["title", "cover_title"].includes(view.kind) ? "h1" : ["section", "page_heading"].includes(view.kind) ? "h2" : ["subtitle", "hero", "tagline"].includes(view.kind) ? "h3" : "p";
    return node(tag, "preview-" + view.kind.replaceAll("_", "-"), view.text || view.kind.replaceAll("_", " "));
  });
}

function insertBlock(type) {
  const snippets = {
    section: "\n  - type: section\n    text: New Section\n",
    body: "\n  - type: body\n    text: Add a concise paragraph here.\n",
    bullets: "\n  - type: bullets\n    items:\n      - First point\n      - Second point\n",
    metrics: "\n  - type: metrics\n    items:\n      - label: Metric\n        value: 42\n        tone: positive\n",
  };
  const snippet = snippets[type];
  if (!snippet) return;
  const start = elements.editor.selectionStart;
  elements.editor.setRangeText(snippet, start, elements.editor.selectionEnd, "end");
  elements.editor.focus();
  render();
}

function updateArtifact(link, href, enabled) {
  link.href = enabled ? href : "";
  link.setAttribute("aria-disabled", String(!enabled));
  link.tabIndex = enabled ? 0 : -1;
}

function downloadYaml() {
  const model = parseSpecPreview(elements.editor.value);
  const url = URL.createObjectURL(new Blob([elements.editor.value], { type: "application/yaml" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = downloadableName(model.outputName);
  link.click();
  URL.revokeObjectURL(url);
}

async function copyCli(event) {
  try {
    await navigator.clipboard.writeText(elements.cli.textContent);
    event.currentTarget.textContent = "Copied";
    window.setTimeout(() => { event.currentTarget.textContent = "Copy command"; }, 1200);
  } catch {
    event.currentTarget.textContent = "Copy failed";
  }
}

function setLoading(value) {
  elements.editor.disabled = value;
  elements.state.textContent = value ? "Loading checked-in source…" : elements.state.textContent;
}

function example(id, title, file) {
  return { id, title, file };
}

function node(tag, className, text) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (text) element.textContent = text;
  return element;
}

function messageNode(text) {
  return node("p", "preview-empty", text);
}
