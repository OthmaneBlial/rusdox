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
- [docs/architecture.md](docs/architecture.md)
- [docs/stability.md](docs/stability.md)
- [GOVERNANCE.md](GOVERNANCE.md)

## Best First Contributions

Good contributions include:

- new YAML document examples
- docs improvements
- bug fixes with regression tests
- better config wizard wording
- higher-quality output rendering
- template gallery improvements

The maintained [`good first issue`](https://github.com/OthmaneBlial/rusdox/issues?q=is%3Aissue+is%3Aopen+label%3A%22good+first+issue%22)
queue contains ten bounded tasks, each with a checked-in fixture and three
acceptance criteria. List or prepare one without modifying the source fixture:

```bash
node scripts/contributor_lab.mjs list
node scripts/contributor_lab.mjs prepare protocol-inline-toml
```

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
node scripts/check_compatibility_contract.mjs
node scripts/check_accessibility_contract.mjs
node scripts/build_v1_registry_manifest.mjs --check
node scripts/build_template_registry.mjs --registry-dir registry/v1 --check
node scripts/test_benchmark_contract.mjs
node scripts/test_reproducible_release.mjs
node scripts/check_security_review.mjs
```

If you touch visual assets or examples, also run:

```bash
cargo run -- examples
./scripts/generate_gallery_assets.sh
```

For a compatibility or rendering change, generate the normal parity bundle and
compare it with a dated page baseline:

```bash
target/debug/rusdox verify examples/hello_world.yaml --output-root target/contributor-parity
node scripts/contributor_lab.mjs visual-diff \
  --spec examples/hello_world.yaml \
  --baseline tests/golden/pages/linux-x86_64/hello-world
```

Use `--threshold` only when the pull request explains why a non-zero visual
budget is justified. Template changes follow the fixture and evidence commands
in [the registry guide](docs/template-registry.md).

Security-sensitive parser, template, protocol, or dependency changes must also
keep the [v1 security review](docs/security-review-v1.md) accurate and pass
`cargo audit --deny unsound`. Performance-budget changes require a measured
comparable-host report; do not increase a ceiling solely to make CI green.

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
- include parity or viewer evidence when output behavior changes
- confirm that contributor credit remains current with `node scripts/check_contributors.mjs`

## Areas That Matter Most

- YAML authoring experience
- output quality
- speed
- cross-platform reliability
- docs clarity
- template usefulness

## Questions

If you are not sure where to start, read [SUPPORT.md](SUPPORT.md).
Accepted contributors are credited in [CONTRIBUTORS.md](CONTRIBUTORS.md) and the
release notes for their merged change.
