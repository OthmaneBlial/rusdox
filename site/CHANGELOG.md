# Changelog

All notable user-facing changes to RusDox are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and releases follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Shared DOCX/PDF page setup, visible headers and footers, configurable page fields, explicit page and section breaks, hyperlinks, bookmarks, TOC fields, and footnotes.
- Rich and merged table cells, nested tables, repeating headers, row-splitting controls, and measured overflow diagnostics.
- Executable dual-output and international-script contract fixtures with 19-check HTML/JSON parity reports.
- Structural OOXML validation for XML well-formedness, content types, and internal relationships; an external-DOCX round-trip fixture; dated viewer evidence; exact Linux page baselines; and buildable fuzz targets.
- Default, customizable resource ceilings for DOCX ZIP/XML, specs/includes, PNG/JPEG, SVG, and raster allocations.
- Versioned small/medium/1,000-page benchmarks with isolated validation, DOCX, PDF, dual-output, and existing-DOCX pipelines; machine-readable host/hash/memory evidence; scheduled regression detection; and generated history charts.
- Word-native DOCX template inspection and strict/lenient rendering from JSON, including split-run placeholders, nested values, complete paragraph/table-row loops, conditions, deterministic filters, reusable partials, native PDF output, and parity reports.
- Three externally authored starter templates for invoices, proposals, and board/compliance reporting.
- Versioned document specs with a generated JSON Schema shared by YAML, JSON, TOML, and a bundled zero-dependency VS Code extension.
- Schema generation, atomic migration, source-located validation, bounded conditional branches, nested paths, deterministic filters, and literal-brace escaping.
- A zero-install, local-first spec playground with six verified examples, structural preview, editable YAML downloads, exact CLI reproduction commands, and artifact integrity guards.
- A dated WASM feasibility spike and explicit browser-rendering capability boundary.
- A `rusdox dev` feedback loop with a loopback-only PDF/status dashboard, debounced input/config/include/asset watching, last-success preservation, per-stage timings, JSON Lines events, browser opening, and quiet bounded runs.
- A signed curated template registry with hash-verified `template list`, `search`, `add`, and `update` commands; seven discovery categories; contributor credit; accessibility metadata; a featured monthly template; and pull-request preview artifacts.
- A reusable source-built GitHub Action that validates and renders specs, annotates exact pull-request lines, uploads parity-only evidence, and upserts a metadata-only visual-diff summary, plus release-note, invoice, compliance, and scheduled-report recipes.
- An object-safe, in-memory Rust renderer boundary and stable local JSON protocol v1 over newline-delimited stdin/stdout or opt-in loopback HTTP, with atomic hash-described outputs and executable Node, Python, Go, and CI examples.
- A contributor field guide with an end-to-end architecture map, ten fixture-backed starter tasks, isolated fixture preparation, visual-diff commands, contributor credit checks, and a published governance policy.
- Frozen v1 spec, template, renderer, and local-protocol contracts; cross-surface SemVer, deprecation, support, and release policies; an exact Rust 1.88.0 MSRV gate; complete public-item rustdoc; and release-time API compatibility scanning.
- A machine-enforced 30-capability v1 feature contract covering every document block variant and its DOCX/PDF/parity status.
- Bounded concurrent batch rendering with aggregate source-memory preflight, ordered outcomes, and cooperative cancellation between native render stages.
- Default and hosted TOML/JSON resource profiles, including expansion/depth/output ceilings for Word templates and operator-owned limits for local protocol workers.
- A dated, executable v1 security review with DTD rejection, SVG external-resource denial, weekly RustSec auditing, and explicit residual-risk tracking.
- Reproducibility receipts for twice-built release binaries, deterministic tar/ZIP packaging, release provenance verification before installer smoke tests, and per-target receipts attached to releases.
- An accessible long-term benchmark dashboard backed by absolute runtime and peak-memory budgets for all 13 protocol scenarios.

### Contributors

- Othmane Blial — maintainer and author of the current unreleased document engine, parity, distribution, integration, and contributor-experience work.
- Dependabot — automated dependency update pull requests.

No external human contribution has landed in this unreleased cycle yet. New
contributors will be added here when a meaningful pull request is merged.

### Changed

- `PageSetup`, `Run`, `RowSpec`, and related low-level constructors now carry the expanded dual-output contract; this pre-1.0 Rust API change is incorporated into the v1 contract.
- PDF page geometry now comes from the document's `PageSetup`; PDF config geometry remains a compatibility fallback when no document-level value is supplied.
- PDF nested tables use a bounded flattened-row representation and do not claim exact nested-grid geometry.
- Benchmark summaries now report true extrema plus medians; the previous multi-iteration maximum accidentally reported the sum of samples.
- Updated `anyhow` and `memmap2` lockfile entries to versions that clear the 2026-08-24 RustSec unsoundness advisories; upstream `rustybuzz`/`ttf-parser` maintenance warnings remain tracked in the security review.

### Fixed

- Script-mode CLI tests now share a Cargo target cache and use unique package names, preventing multi-gigabyte temporary build accumulation and binary collisions.
- DOCX, PDF, and spec writes now preserve the previous valid destination when a render/write is interrupted, including atomic replacement on Windows instead of deleting the old file first.
- Signed registry text assets retain LF bytes on every checkout, and local registry fixtures hash the checked-out bytes so Windows CRLF policy cannot create false integrity failures.

## [0.1.1] - 2026-08-23

### Changed

- User installs now expose only the `rusdox` executable; stress generators remain available to maintainers through the `maintainer-tools` feature.
- crates.io releases now use short-lived GitHub OIDC credentials through Trusted Publishing instead of a long-lived repository secret.
- Installer fixtures derive the expected version from `Cargo.toml` so release checks cannot silently drift.

### Fixed

- `cargo binstall rusdox` can consume the native GitHub Release archive without falling back to a source build for missing maintainer-only binaries.

## [0.1.0] - 2026-08-23

### Added

- A public product roadmap focused on trustworthy DOCX/PDF parity, Word-native templates, and production adoption.
- Lean crates.io packaging that excludes the duplicated showcase bundle and generated binary documents.
- Structured community routes for support questions and private vulnerability reports.
- Pure-Rust DOCX generation and native PDF preview rendering.
- YAML, JSON, and TOML document specifications.
- Variables, includes, repeaters, named styles, metadata, visuals, and tables.
- `validate`, `watch`, and `bench` CLI workflows.
- Cross-platform release installers for Linux, macOS, and Windows.

### Changed

- Repository, documentation, CI, and security references now consistently target the `main` branch.
- CI verifies the exact locked dependency graph and the publishable crate package.

[Unreleased]: https://github.com/OthmaneBlial/rusdox/compare/v0.1.1...HEAD
[0.1.1]: https://github.com/OthmaneBlial/rusdox/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/OthmaneBlial/rusdox/releases/tag/v0.1.0
