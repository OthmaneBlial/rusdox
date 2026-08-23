# GitHub Maintainer Setup

This checklist contains repository-owner tasks. It is intentionally separate from end-user documentation.

## Repository presentation

- Description: `Generate editable DOCX and native PDF from one readable spec, at Rust speed.`
- Homepage: `https://othmaneblial.github.io/rusdox/`
- Topics: `rust`, `docx`, `pdf`, `yaml`, `ooxml`, `document-generation`, `cli`, `templates`, `reporting`, `office-automation`.
- Social preview: `assets/social-preview-rusdox.png`.
- Pin the repository only after the current install command and release assets are verified.

## Community

- Enable Discussions with Announcements, Q&A, Ideas, and Show and tell.
- Keep usage questions in Discussions and actionable bugs in Issues.
- Maintain `good first issue`, `help wanted`, `compatibility`, `template`, and `docs` labels.
- Give every starter issue a fixture, expected behavior, and testable acceptance criteria.

## Security

- Enable Dependabot alerts and security updates.
- Enable secret scanning and push protection.
- Enable CodeQL/default code scanning when Rust analysis is supported for the repository.
- Enable private vulnerability reporting.
- Keep `SECURITY.md` aligned with supported release lines.

## Release prerequisites

- Add `CARGO_REGISTRY_TOKEN` to the protected `crates-io` environment.
- Protect the environment with an approval rule if another maintainer is available.
- Enable immutable releases after the first draft-to-published workflow succeeds.
- Confirm Actions has permission to create attestations and release assets.

## Release checklist

1. Update `Cargo.toml`, `Cargo.lock`, `CHANGELOG.md`, and the version badge.
2. Run `cargo fmt --check`, Clippy, the complete test suite, and `cargo publish --dry-run`.
3. Regenerate gallery, static docs, compatibility fixtures, and benchmark data.
4. Commit and push the release state.
5. Create and push the signed `vX.Y.Z` tag.
6. Wait for build, crate publication, attestations, release publication, and clean-install smoke jobs.
7. Verify `SHA256SUMS`, `gh attestation verify`, crates.io, docs.rs, and the live Pages site independently.
8. Publish the release announcement only after all links work.
