const SCALAR_PATTERN = /^([a-zA-Z_][a-zA-Z0-9_-]*):\s*(.*)$/;

export function parseSpecPreview(source) {
  const lines = source.replaceAll("\r\n", "\n").split("\n");
  const document = { version: null, outputName: "document", blocks: [], diagnostics: [] };
  let current = null;
  let collection = null;

  for (let index = 0; index < lines.length; index += 1) {
    const raw = lines[index];
    const trimmed = raw.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;

    const topLevel = raw.match(/^([a-zA-Z_][a-zA-Z0-9_-]*):\s*(.*)$/);
    if (topLevel && !raw.startsWith(" ")) {
      const value = parseScalar(topLevel[2]);
      if (topLevel[1] === "version") document.version = Number(value);
      if (topLevel[1] === "output_name") document.outputName = String(value || "document");
      continue;
    }

    const blockStart = raw.match(/^  - type:\s*([a-zA-Z0-9_-]+)\s*$/);
    if (blockStart) {
      current = { type: blockStart[1], text: "", items: [], cells: [] };
      document.blocks.push(current);
      collection = null;
      continue;
    }
    if (!current) continue;

    const property = raw.match(/^    ([a-zA-Z_][a-zA-Z0-9_-]*):\s*(.*)$/);
    if (property) {
      const key = property[1];
      const value = parseScalar(property[2]);
      if (key === "text") current.text = String(value);
      collection = key === "items" || key === "spec" ? key : null;
      continue;
    }

    if (collection === "items") {
      const item = raw.match(/^      -\s*(.*)$/);
      if (item) {
        const pair = item[1].match(SCALAR_PATTERN);
        current.items.push(pair ? { [pair[1]]: parseScalar(pair[2]) } : parseScalar(item[1]));
        continue;
      }
      const nestedProperty = raw.match(/^        ([a-zA-Z_][a-zA-Z0-9_-]*):\s*(.*)$/);
      if (nestedProperty && current.items.length) {
        const last = current.items[current.items.length - 1];
        if (typeof last === "object") last[nestedProperty[1]] = parseScalar(nestedProperty[2]);
      }
      continue;
    }

    if (collection === "spec") {
      const cellText = raw.match(/^\s{10,}text:\s*(.*)$/);
      if (cellText) current.cells.push(String(parseScalar(cellText[1])));
    }
  }

  if (document.version !== 1) document.diagnostics.push("Expected version: 1.");
  if (!document.blocks.length) document.diagnostics.push("No blocks found under blocks:.");
  if (!document.outputName.trim()) document.diagnostics.push("output_name must not be empty.");
  return document;
}

export function previewBlockModel(block) {
  if (block.type === "metrics") {
    return { kind: "metrics", items: block.items.filter((item) => typeof item === "object") };
  }
  if (block.type === "label_values") {
    return { kind: "pairs", items: block.items.filter((item) => typeof item === "object") };
  }
  if (block.type === "bullets") {
    return { kind: "bullets", items: block.items.map((item) => String(item)) };
  }
  if (block.type === "table") return { kind: "table", cells: block.cells };
  return { kind: block.type, text: block.text };
}

export function cliCommand(examplePath, outputName) {
  const safePath = String(examplePath).replace(/^\.\.\//, "");
  return "rusdox " + shellQuote(safePath) + "\n# writes generated/" + outputName + ".docx and rendered/" + outputName + ".pdf";
}

export function downloadableName(outputName) {
  const value = String(outputName || "document").replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/^-+|-+$/g, "");
  return (value || "document") + ".yaml";
}

function shellQuote(value) {
  if (/^[a-zA-Z0-9_./-]+$/.test(value)) return value;
  return "'" + value.replaceAll("'", "'\\''") + "'";
}

function parseScalar(value) {
  const trimmed = value.trim();
  if (!trimmed) return "";
  if ((trimmed.startsWith('"') && trimmed.endsWith('"')) || (trimmed.startsWith("'") && trimmed.endsWith("'"))) {
    return trimmed.slice(1, -1);
  }
  if (trimmed === "true") return true;
  if (trimmed === "false") return false;
  return trimmed;
}
