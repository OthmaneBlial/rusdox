# Compatibility Matrix

This page is the public RusDox support contract. It distinguishes implemented behavior from parity-tested behavior and planned work.

Status meanings:

- **Supported**: implemented and covered by automated tests.
- **Partial**: useful behavior exists, but the documented edge cases or DOCX/PDF parity are incomplete.
- **Not supported**: RusDox rejects, ignores, or does not expose the feature today.
- **Planned**: accepted roadmap work, not a current capability.

## Input formats and composition

| Capability | Status | Notes |
|---|---|---|
| YAML specs | Supported | Recommended authoring format. |
| JSON specs | Supported | Same serializable document model without YAML-only expansion. |
| TOML specs | Supported | Same serializable document model without YAML-only expansion. |
| Variables | Supported in YAML | Scalar interpolation through the YAML expansion layer. |
| Includes | Supported in YAML | Relative includes with cycle detection. |
| Repeaters | Supported in YAML | Sequence expansion with index and item variables. |
| Conditions and filters | Planned | Targeted for the schema/template milestone. |
| Stable spec version field | Planned | Pre-1.0 specs do not yet declare a schema version. |
| JSON Schema / editor completion | Planned | Targeted for v0.3. |

## Document features

| Feature | DOCX | PDF | Parity-tested | Notes |
|---|---:|---:|---:|---|
| Paragraphs and multiple runs | Supported | Supported | Partial | Text order is tested; visual parity reporting is planned. |
| Bold, italic, underline, strike, size, color | Supported | Supported | Partial | PDF maps the common supported run properties. |
| Paragraph alignment and spacing | Supported | Supported | Partial | Viewer-level snapshots are not yet a release gate. |
| Bulleted and numbered lists | Supported | Supported | Partial | Semantic OOXML numbering is emitted for DOCX. |
| Named paragraph/run/table styles | Supported | Supported | Partial | Style inheritance is tested; viewer matrix is pending. |
| Tables and repeating header rows | Supported | Supported | Partial | Long-row and complex overflow behavior is still being hardened. |
| Rich paragraphs inside table cells | Supported in Rust API | Partial | No | High-level YAML cells are currently text/status-oriented. |
| Merged cells | Partial in Rust API | Partial | No | Grid-span parsing/layout exists internally; no stable high-level contract yet. |
| Nested tables | Not supported | Not supported | No | Planned only after common table parity is stable. |
| PNG and JPEG images | Supported | Supported | Partial | Path and embedded-byte Rust APIs are covered. |
| SVG visuals | Supported | Supported | Partial | Rasterized for output; complex SVG compatibility varies. |
| Image alt text | Supported | Partial | No | Preserved in DOCX; tagged-PDF semantics are not implemented. |
| Document metadata | Supported | Not supported | No | DOCX core/custom properties are implemented. |
| Page size and margins | Supported | Partial | No | PDF currently reads PDF config rather than the document page setup. |
| Headers and footers | Supported | Not supported | No | DOCX supports one default header/footer template. |
| Page number fields | Supported | Not supported | No | DOCX supports page and total-page fields. |
| Explicit page breaks | Supported in Rust API | Supported | Partial | A dedicated high-level block is planned. |
| Multiple sections / section breaks | Not supported | Not supported | No | A document currently has one section property set. |
| Hyperlinks and bookmarks | Not supported | Not supported | No | Planned for the parity milestone. |
| Automatic table of contents | Not supported | Not supported | No | Planned with an explicit field/update contract. |
| Footnotes | Not supported | Not supported | No | Planned before comments/tracked changes. |
| Comments and tracked changes | Not supported | Not supported | No | Explicitly deferred beyond the common report feature set. |
| Word-native placeholders | Not supported | Not supported | No | Existing DOCX packages can be opened in Rust; template rendering is planned. |

## Existing DOCX packages

The Rust API can open a DOCX in read-only or read/write mode. Non-document package parts are retained when saving, and automated fixtures cover custom-part preservation.

This is not yet a promise of lossless editing for arbitrary Word files. RusDox parses the body features it understands and rewrites `word/document.xml`; unsupported body markup may therefore be normalized or lost. Keep an original copy and test representative templates before production use.

## Text and fonts

| Area | Status | Notes |
|---|---|---|
| Latin text | Supported | Primary tested path. |
| Unicode font embedding in PDF | Supported | TrueType fonts and ToUnicode maps are tested. |
| Font fallback | Partial | Depends on fonts installed on the rendering machine. |
| Mixed scripts | Partial | Shaping has automated coverage, but no published viewer matrix yet. |
| Arabic and RTL layout | Experimental | Do not assume correct bidirectional paragraph behavior yet. |
| CJK | Experimental | Glyph coverage depends on installed fonts; line-breaking rules are incomplete. |
| Emoji | Experimental | Color emoji and fallback differ by platform and viewer. |

## Viewers and operating systems

CI currently builds and tests on current GitHub-hosted Ubuntu, macOS, and Windows runners. That proves compilation and automated behavior on those platforms; it does not prove visual fidelity in every document viewer.

A dated manual compatibility scorecard for Microsoft Word, LibreOffice, Apple Preview, and Adobe Acrobat is planned for v0.2. Until then, report viewer-specific problems with the smallest non-sensitive fixture possible.

## Reporting a compatibility problem

Open a [bug report](https://github.com/OthmaneBlial/rusdox/issues/new?template=bug_report.yml) with:

- RusDox version and installation method;
- operating system and architecture;
- viewer name and version;
- minimal YAML/Rust input;
- expected and actual output;
- sanitized screenshots or generated files when safe to share.

See [Troubleshooting](troubleshooting.md) before attaching private business documents.
