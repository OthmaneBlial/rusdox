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

### Changed

- `PageSetup`, `Run`, `RowSpec`, and related low-level constructors now carry the expanded dual-output contract; this is a breaking Rust API change planned for `v0.2.0`.
- PDF page geometry now comes from the document's `PageSetup`; PDF config geometry remains a compatibility fallback when no document-level value is supplied.
- PDF nested tables use a bounded flattened-row representation and do not claim exact nested-grid geometry.
- Benchmark summaries now report true extrema plus medians; the previous multi-iteration maximum accidentally reported the sum of samples.

### Fixed

- Script-mode CLI tests now share a Cargo target cache and use unique package names, preventing multi-gigabyte temporary build accumulation and binary collisions.
- DOCX, PDF, and spec writes now preserve the previous valid destination when a render/write is interrupted, including atomic replacement on Windows instead of deleting the old file first.

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
