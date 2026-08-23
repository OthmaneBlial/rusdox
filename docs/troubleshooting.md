# Troubleshooting

This guide covers the most common installation, authoring, rendering, and viewer problems.

## Installation fails

### Release download returns 404

Installers download a platform archive and `SHA256SUMS` from GitHub Releases. Confirm that the requested version exists:

```bash
RUSDOX_VERSION=v0.1.0 sh scripts/install.sh
```

For a source build while diagnosing release availability:

```bash
cargo install --path . --locked
```

### The binary is installed but not on `PATH`

The Unix installer uses `/usr/local/bin` when writable and otherwise uses `$HOME/.local/bin`. Add the printed directory to your shell profile, then start a new terminal.

On Windows, the installer updates the user `PATH`. Existing PowerShell windows may need to be reopened.

### Checksum verification fails

Do not bypass the check. Delete the downloaded archive, retry on a trusted network, and confirm that the archive and `SHA256SUMS` come from the same release. If it still fails, report the release tag, platform, and architecture without attaching the corrupted binary.

## RusDox cannot find an asset

Relative visual paths are resolved from the document spec directory, not necessarily the current shell directory.

```yaml
blocks:
  - type: image
    path: assets/chart.png
```

For `reports/weekly.yaml`, that resolves to `reports/assets/chart.png`. Run validation for an exact path diagnostic:

```bash
rusdox validate reports/weekly.yaml
```

Supported visual formats are PNG, JPEG, and SVG.

## Fonts differ across machines

PDF rendering uses fonts available on the machine. A missing configured family can trigger fallback and change line wrapping.

For repeatable output:

- install the same licensed fonts on every rendering host;
- use the same `rusdox.toml`;
- record the OS and font package versions in CI;
- keep a rendered regression fixture for important documents.

RusDox does not redistribute arbitrary commercial fonts.

## DOCX and PDF do not look identical

DOCX is laid out by the viewer, while RusDox lays out PDF itself. Current parity limitations are listed in the [compatibility matrix](compatibility.md). In particular, document page setup, headers, footers, and page-number fields do not yet have full PDF parity.

When reporting a mismatch, include:

- the smallest input that reproduces it;
- the active config;
- the Word/LibreOffice/PDF viewer and version;
- screenshots of the exact region;
- whether text content, pagination, or visual styling differs.

## A table overflows or splits badly

- Check that column widths fit within the configured content width.
- Shorten unbreakable strings such as URLs and identifiers.
- Validate the spec for row/column shape warnings.
- Use repeating header rows through the Rust API for long tables.
- Reduce cell padding or font size only after checking the source data.

Complex merged/nested tables are not yet a stable high-level feature.

## A large document is slow or uses too much memory

Measure instead of guessing:

```bash
rusdox bench report.yaml --iterations 5 --warmup 1 --format json
```

To isolate the DOCX path:

```bash
rusdox report.yaml --docx-only
```

Large raster images often dominate memory and output size. Resize source images before rendering and avoid embedding the same oversized asset repeatedly.

## Watch mode does not rebuild

RusDox watches the input spec and active config with a polling loop. Confirm the path and reduce the interval while diagnosing:

```bash
rusdox watch report.yaml --poll-interval-ms 250
```

If a generated file is being written into the watched input directory by another tool, move output folders elsewhere to avoid noisy rebuilds.

## Validation output in CI

Use JSON and preserve the non-zero exit code:

```bash
rusdox validate docs --format json
```

Do not pipe through a command that masks the RusDox exit status.

## Existing DOCX content changes after save

The Rust API preserves non-document package parts, but it rewrites body markup using the features RusDox understands. Arbitrary Word documents are not yet guaranteed to round-trip losslessly. Work from a copy and reduce the document to a safe fixture before reporting a parser issue.

## Asking for help safely

Use [GitHub Discussions](https://github.com/OthmaneBlial/rusdox/discussions/categories/q-a) for authoring questions and [Issues](https://github.com/OthmaneBlial/rusdox/issues) for reproducible bugs.

Before sharing a document:

- remove names, emails, customer data, secrets, and hidden metadata;
- replace logos and signatures;
- inspect headers, footers, comments, and custom properties;
- prefer a newly created minimal fixture.

Report security-sensitive behavior through [private vulnerability reporting](https://github.com/OthmaneBlial/rusdox/security/advisories/new).
