# Contributing

Thanks for contributing to RusDox.

The project goal is simple:

- author documents as YAML
- render DOCX and PDF in pure Rust
- keep the experience easy for normal users

## Before You Start

Read these first:

- [README.md](README.md)
- [docs/README.md](docs/README.md)
- [docs/yaml-guide.md](docs/yaml-guide.md)
- [docs/rust-api.md](docs/rust-api.md)

## Best First Contributions

Good contributions include:

- new YAML document examples
- docs improvements
- bug fixes with regression tests
- better config wizard wording
- higher-quality output rendering
- template gallery improvements

## Before Opening A Large PR

For substantial changes, open an issue first and explain:

- the problem
- the proposed solution
- user-facing impact
- whether it changes YAML, config, or rendering behavior

This keeps the project direction consistent.

## Local Setup

```bash
cargo fmt --all
cargo clippy --all-targets --all-features -- -D warnings
cargo test
```

If you touch visual assets or examples, also run:

```bash
cargo run -- examples
./scripts/generate_gallery_assets.sh
```

## Contribution Rules

- Keep the user-facing story YAML-first.
- Do not add complexity unless it clearly improves authoring or output quality.
- Prefer readable examples over clever abstractions.
- Keep configuration centralized instead of scattering styling through examples.
- Add or update tests when behavior changes.
- Update docs when changing YAML, config, CLI, or output behavior.

## Pull Request Checklist

- explain what changed and why
- mention any user-visible behavior change
- include tests or explain why tests were not needed
- update docs/examples if relevant
- keep the worktree clean and focused

## Areas That Matter Most

- YAML authoring experience
- output quality
- speed
- cross-platform reliability
- docs clarity
- template usefulness

## Questions

If you are not sure where to start, read [SUPPORT.md](SUPPORT.md).
