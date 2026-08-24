# Compatibility Matrix

This page is the public RusDox support contract. It distinguishes implemented behavior from parity-tested behavior and planned work.

The same v1 surface is checked as machine-readable data in
[`compatibility/v1-feature-contract.json`](../compatibility/v1-feature-contract.json).
Its executable gate requires evidence for every capability and maps every
`BlockSpec` variant exactly once, so a new public block cannot bypass this page.

The document feature matrix exposes four independent states instead of hiding nuance behind one label:

- **DOCX**: the feature has an implemented DOCX path.
- **PDF**: the feature has an implemented native PDF path.
- **Parity-tested**: the source, reopened DOCX, and PDF projection are covered by the published parity contract.
- **Intentionally unsupported**: RusDox deliberately excludes the feature from its current product contract. A blank value means planned or incomplete, not silently supported.

Within a supported column, **Partial** or **Experimental** identifies a narrower boundary that must be read with the notes.

## Input formats and composition

| Capability | Status | Notes |
|---|---|---|
| YAML specs | Supported | Recommended authoring format. |
| JSON specs | Supported | Same serializable document model without YAML-only expansion. |
| TOML specs | Supported | Same serializable document model without YAML-only expansion. |
| Variables | Supported in YAML | Scalar interpolation through the YAML expansion layer. |
| Includes | Supported | Relative YAML, JSON, and TOML includes with cycle detection. |
| Repeaters | Supported | Cross-format sequence expansion with index and item variables. |
| Conditions and filters | Supported | Specs and Word-native templates support bounded conditions, nested paths, and upper, lower, title, trim, and default filters. |
| Versioned spec field | Supported | Current authoring contract is version 1 with a documented compatibility and migration policy. |
| JSON Schema / editor completion | Supported | Generated schema plus bundled zero-dependency VS Code completion, hover, enum, and diagnostic service. |

## Document features

| Feature | DOCX | PDF | Parity-tested | Intentionally unsupported | Notes |
|---|---:|---:|---:|---:|---|
| Paragraphs and multiple runs | Supported | Supported | Yes |  | Semantic order plus deterministic page snapshots. |
| Bold, italic, underline, strike, size, color | Supported | Supported | Yes |  | PDF maps the common supported run properties. |
| Paragraph alignment and spacing | Supported | Supported | Yes |  | Viewer scorecard remains separate from renderer parity. |
| Bulleted and numbered lists | Supported | Supported | Yes |  | Semantic OOXML numbering is emitted for DOCX. |
| Named paragraph/run/table styles | Supported | Supported | Yes |  | Style inheritance is tested; viewer matrix remains separate. |
| Tables and repeating header rows | Supported | Supported | Yes |  | High-level rows expose repeat and split controls; oversized unsplittable rows fail with a measured diagnostic. |
| Rich paragraphs inside table cells | Supported | Supported | Yes |  | YAML `kind: rich` cells preserve multiple formatted paragraphs. |
| Merged cells | Supported | Supported | Yes |  | Horizontal `grid_span` is supported and parity-tested. Vertical merges are not yet supported. |
| Nested tables | Supported | Bounded | Yes |  | DOCX uses native nested tables. PDF renders nested rows inside the parent cell; exact nested-grid geometry is not claimed. |
| PNG and JPEG images | Supported | Supported | Yes |  | Count, semantic kind, and alt text are compared. |
| SVG visuals | Supported | Supported | Yes |  | Rasterized for output; complex SVG compatibility varies. |
| Image alt text | Supported | Semantic projection | Yes |  | Non-blank alt text is required and compared, but the current PDF has no structure tree. |
| Document metadata | Supported | Supported | Yes |  | Core PDF title/author/subject/keywords plus catalog language are emitted. |
| Page size, orientation, and margins | Supported | Supported | Yes |  | Both outputs consume the same `PageSetup`; width/height remain explicit and orientation is validated. |
| Headers and footers | Supported | Supported | Yes |  | PDF renders the same text/alignment template on every page. |
| Page number and total-page fields | Supported | Supported | Yes |  | `{page}` and `{pages}` respect restart and decimal/Roman/letter formats. |
| Explicit page breaks | Supported | Supported | Yes |  | Dedicated `page_break` block plus paragraph-level control. |
| Next-page section breaks | Supported | Supported | Yes |  | Dedicated `section_break` block reuses the active section controls; independent per-section geometry is not yet supported. |
| Hyperlinks and bookmarks | Supported | Supported | Yes |  | DOCX uses field links/bookmarks; PDF emits URI/GoTo annotations and bookmark outlines. |
| Automatic table of contents | Supported | Supported fallback | Yes |  | DOCX emits an updateable TOC field. PDF freezes the spec heading list at render time without computed page numbers. |
| Footnotes | Supported | Supported fallback | Yes |  | DOCX emits a real footnotes part. PDF uses inline markers plus a deterministic endnotes page. |
| Comments and tracked changes | No | No | No | Yes | Deliberately deferred beyond common generated-report features. |
| Word-native placeholders | Supported | Supported subset | Yes |  | Scalars, nested values, complete paragraph/row loops, conditions, and partials preserve untouched DOCX parts; PDF covers the parsed body subset. |

## Existing DOCX packages

The Rust API can open a DOCX in read-only or read/write mode. Non-document package parts are retained when saving. Automated fixtures cover both synthetic custom parts and an external DOCX produced by the macOS text system, then opened, modified, saved, structurally validated, and reopened by RusDox.

This is not yet a promise of lossless editing for arbitrary Word files. RusDox parses the body features it understands and rewrites `word/document.xml`; unsupported body markup may therefore be normalized or lost. Keep an original copy and test representative templates before production use.

## Text and fonts

| Area | Status | Notes |
|---|---|---|
| Latin text | Supported | Primary tested path. |
| Unicode font embedding in PDF | Supported | Embeddable outline TrueType fonts and ToUnicode maps are tested; restricted or outline-forbidden fonts are rejected. |
| Font fallback | Partial | Requested, generic, and script-specific fallbacks are tried. Exact resolution depends on installed fonts and missing characters are recorded in parity evidence. |
| Mixed scripts | Partial | The `international-scripts` fixture covers Latin, Arabic, Hebrew, CJK, emoji, and mixed lines through DOCX reopen and native PDF extraction. |
| Arabic and RTL layout | Experimental | Unicode text and right alignment are preserved, but contextual Arabic shaping and the Unicode bidi algorithm are not yet implemented by the PDF renderer. |
| CJK | Experimental | Glyph fallback is exercised; language-specific line-breaking and kinsoku rules are not implemented. |
| Emoji | Experimental | Monochrome glyph fallback may work. Color/ZWJ emoji sequences are not guaranteed and vary with installed fonts. |

The machine-readable
[`international-readiness.json`](../compatibility/international-readiness.json)
prevents RTL or CJK from graduating while any shaping, line-breaking, font, or
viewer gate is unmet. Read the complete [font, language, and alt-text contract](international-accessibility.md)
and the [tagged PDF/PDF/A research](pdf-conformance-research.md). Current PDFs
are not claimed as tagged PDF, PDF/UA, or PDF/A.

## Viewers and operating systems

CI currently builds and tests on current GitHub-hosted Ubuntu, macOS, and Windows runners. That proves compilation and automated behavior on those platforms; it does not prove visual fidelity in every document viewer.

The [dated viewer scorecard](compatibility-scorecard.md) records exact app versions, fixture hashes, successful smoke tests, blocked environments, and viewers that were not run. Untested rows remain explicitly unclaimed. Resource ceilings and fuzz coverage are documented in [Input Safety and Resource Limits](input-safety.md).

## Reporting a compatibility problem

Open a [bug report](https://github.com/OthmaneBlial/rusdox/issues/new?template=bug_report.yml) with:

- RusDox version and installation method;
- operating system and architecture;
- viewer name and version;
- minimal YAML/Rust input;
- expected and actual output;
- sanitized screenshots or generated files when safe to share.

See [Troubleshooting](troubleshooting.md) before attaching private business documents.
