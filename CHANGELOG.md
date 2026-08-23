# Changelog

All notable user-facing changes to RusDox are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and releases follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- A public product roadmap focused on trustworthy DOCX/PDF parity, Word-native templates, and production adoption.
- Lean crates.io packaging that excludes the duplicated showcase bundle and generated binary documents.
- Structured community routes for support questions and private vulnerability reports.

### Changed

- Repository, documentation, CI, and security references now consistently target the `main` branch.
- CI verifies the exact locked dependency graph and the publishable crate package.

## [0.1.0] - 2026-08-23

### Added

- Pure-Rust DOCX generation and native PDF preview rendering.
- YAML, JSON, and TOML document specifications.
- Variables, includes, repeaters, named styles, metadata, visuals, and tables.
- `validate`, `watch`, and `bench` CLI workflows.
- Cross-platform release installers for Linux, macOS, and Windows.

[Unreleased]: https://github.com/OthmaneBlial/rusdox/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/OthmaneBlial/rusdox/releases/tag/v0.1.0
