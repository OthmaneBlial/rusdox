# Compatibility Matrix

This page is the public RusDox support contract. It distinguishes implemented behavior from parity-tested behavior and planned work.

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
| Includes | Supported in YAML | Relative includes with cycle detection. |
| Repeaters | Supported in YAML | Sequence expansion with index and item variables. |
| Conditions and filters | Planned | Targeted for the schema/template milestone. |
| Stable spec version field | Planned | Pre-1.0 specs do not yet declare a schema version. |
| JSON Schema / editor completion | Planned | Targeted for v0.3. |

## Document features

| Feature | DOCX | PDF | Parity-tested | Intentionally unsupported | Notes |
|---|---:|---:|---:|---:|---|
| Paragraphs and multiple runs | Supported | Supported | Yes |  | Semantic order plus deterministic page snapshots. |
| Bold, italic, underline, strike, size, color | Supported | Supported | Yes |  | PDF maps the common supported run properties. |
| Paragraph alignment and spacing | Supported | Supported | Yes |  | Viewer scorecard remains separate from renderer parity. |
| Bulleted and numbered lists | Supported | Supported | Yes |  | Semantic OOXML numbering is emitted for DOCX. |
| Named paragraph/run/table styles | Supported | Supported | Yes |  | Style inheritance is tested; viewer matrix remains separate. |
| Tables and repeating header rows | Supported | Supported | Yes |  | Complex overflow hardening continues under the same contract. |
| Rich paragraphs inside table cells | Supported in Rust API | Partial | Partial |  | High-level YAML cells remain text/status-oriented. |
| Merged cells | Partial in Rust API | Partial | Partial |  | Grid-span parsing/layout exists; the high-level contract is not stable. |
| Nested tables | No | No | No |  | Planned only after common table parity is stable. |
| PNG and JPEG images | Supported | Supported | Yes |  | Count, semantic kind, and alt text are compared. |
| SVG visuals | Supported | Supported | Yes |  | Rasterized for output; complex SVG compatibility varies. |
| Image alt text | Supported | Semantic projection | Yes |  | Tagged-PDF accessibility semantics are not yet advertised. |
| Document metadata | Supported | Supported | Yes |  | Core PDF title/author/subject/keywords are emitted. |
| Page size and margins | Supported | Partial | Yes |  | Semantic settings match; the PDF physical-layout gap remains tracked. |
| Headers and footers | Supported | No | Yes |  | Semantic projection is checked; visible PDF rendering is pending. |
| Page number fields | Supported | No | Yes |  | Semantic projection is checked; visible PDF rendering is pending. |
| Explicit page breaks | Supported in Rust API | Supported | Yes |  | A dedicated high-level block remains planned. |
| Multiple sections / section breaks | No | No | No |  | A document currently has one section property set. |
| Hyperlinks and bookmarks | No | No | No |  | Planned for the parity milestone. |
| Automatic table of contents | No | No | No |  | Planned with an explicit field/update contract. |
| Footnotes | No | No | No |  | Planned before comments or tracked changes. |
| Comments and tracked changes | No | No | No | Yes | Deliberately deferred beyond common generated-report features. |
| Word-native placeholders | No | No | No |  | Existing DOCX packages can be opened; template rendering is planned. |

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
