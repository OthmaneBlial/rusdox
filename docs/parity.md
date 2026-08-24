# Parity Verification

`rusdox verify` turns the dual-output promise into a CI contract:

```bash
rusdox verify board-report.yaml
```

It writes:

```text
generated/board-report.docx
rendered/board-report.pdf
reports/board-report-parity.html
reports/board-report-parity.json
reports/board-report-pages/page-001.png
```

The HTML report is standalone and requires no JavaScript. The JSON report uses the versioned parity report contract and is suitable for CI, release evidence, or downstream tooling.

## What is compared

RusDox compares three views:

1. the typed document composed from the source spec;
2. the generated DOCX reopened through the RusDox OOXML reader;
3. the semantic projection actually consumed by the native PDF layout path.

The current contract contains 21 named checks:

- normalized text and heading sequence;
- top-level block order;
- every table row and cell;
- image count and semantic kind, plus a dedicated non-blank alt-text equality check;
- explicit page-break positions;
- explicit section-break positions;
- hyperlinks, bookmarks, dynamic fields, and footnote text;
- repeating/splittable row controls, grid spans, rich-cell paragraph counts, and nested-table presence;
- document metadata plus a dedicated language equality check across the source,
  reopened DOCX, PDF projection, and rendered PDF catalog;
- page setup, headers, footers, and numbering settings;
- required DOCX parts and package relationships;
- PDF header, trailer, and page evidence;
- optional deterministic rendered-page visual thresholds.

Each generated DOCX and PDF also receives a SHA-256 digest in the report.
The PDF evidence additionally lists resolved embedded fonts, their OpenType
embedding permission, mapped-glyph count, missing characters, and explicit
`false` values for the currently unclaimed tagged-PDF and PDF/A states. See the
[international and accessibility contract](international-accessibility.md).

## Visual regression thresholds

Every verification emits deterministic PNG snapshots from the PDF renderer's real layout operations. These geometry rasters show line wrapping, text positions, tables, images, spacing, and page boundaries without depending on a particular desktop PDF viewer.

The repository checks in 26 Linux baselines covering every one of the 17 top-level examples under `tests/golden/pages/linux-x86_64/`. Ubuntu CI compares them at a zero-pixel threshold. Baselines are platform-scoped because installed font files can differ; updating them requires reviewer inspection and `scripts/update_visual_baselines.sh`.

Compare a single spec to an approved baseline:

```bash
rusdox verify report.yaml \
  --visual-baseline tests/visual-baselines/report \
  --visual-threshold 0.002
```

For a directory input, the baseline root must contain one subdirectory per output name:

```text
tests/visual-baselines/
  board-report/page-001.png
  invoice/page-001.png
```

The threshold is the maximum fraction of pixels that may differ on each page. `0` is exact. A small non-zero threshold should be justified in the fixture documentation; do not increase it just to hide a regression.

Geometry snapshots are not Microsoft Word, LibreOffice, Preview, or Acrobat screenshots. Viewer evidence is tracked separately in the compatibility scorecard.

## Deterministic and intentionally variable data

Saving the same in-memory document twice produces byte-identical DOCX and PDF files in the regression suite. Stable metadata defaults, ordered maps, deterministic object identifiers, and fixed ZIP entry ordering are part of that contract.

The following are intentionally outside cross-machine byte determinism:

- benchmark durations, peak-memory samples, and host metadata;
- font fallback and page geometry when different font files are installed;
- metadata already present in an externally authored DOCX when the caller chooses to preserve or change it;
- compatibility screenshots and application-generated resaves from different viewer versions.

Semantic projections and OOXML relationship/content-type checks remain the portable contract across hosts.

The executable boundary fixture is [`examples/dual_output_contract.yaml`](../examples/dual_output_contract.yaml). Script coverage and known shaping limits are captured by [`examples/international_scripts.yaml`](../examples/international_scripts.yaml).

## Machine-readable output and exit codes

Print a compact JSON command summary while still writing the full HTML and JSON reports:

```bash
rusdox verify examples --output-root target/parity --format json
```

Exit codes are stable:

| Code | Meaning |
|---:|---|
| `0` | All enabled parity checks passed. |
| `1` | Verification could not complete because input, config, validation, I/O, or rendering failed. |
| `2` | Outputs were generated and at least one parity check failed. |

This distinction lets CI separate a broken build from a real document regression.

## Privacy boundary

Verification is local. RusDox does not upload the source, output documents, report, or snapshots. Uploading reports as CI artifacts is an explicit repository workflow decision; treat reports as sensitive whenever their normalized text or metadata is sensitive.
