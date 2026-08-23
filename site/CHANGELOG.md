# Changelog

All notable user-facing changes to RusDox are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and releases follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
