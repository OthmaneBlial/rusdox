# Release checklist

Use this checklist for every supported release. Record links or hashes in the
release pull request; a checked box without evidence is not a release receipt.

## Contract identity

- [ ] Cargo version matches the intended tag and changelog heading.
- [ ] Document spec, template syntax, renderer API, and local protocol versions
      are recorded; any version change has a migration.
- [ ] `node scripts/check_stable_contracts.mjs` passes.
- [ ] `cargo-semver-checks` passes against the latest published crate, or the
      release is a correctly documented major version.
- [ ] Every public library item has rustdoc and no broken intra-doc links.
- [ ] Rust 1.88.0 compiles all features.

## Product evidence

- [ ] Full locked tests, clippy, format, site, playground, registry, Action, and
      integration contracts pass.
- [ ] DOCX/PDF semantic parity and exact Linux page baselines pass.
- [ ] Compatibility claims link to dated viewer, version, OS, fixture hash, and
      outcome evidence.
- [ ] Accessibility checks pass for every supported semantic.
- [ ] Benchmark history is regenerated; no explicit runtime or memory budget is
      exceeded without a reviewed rationale.
- [ ] Load, concurrency, cancellation, and atomic recovery tests pass.

## Supply chain and publication

- [ ] The publishable crate contains only the intended files.
- [ ] Release archives are produced on all supported targets.
- [ ] SHA-256 checksums verify after download.
- [ ] SPDX SBOM and GitHub build provenance attestations cover every archive.
- [ ] Installers consume the release archive and verify its checksum.
- [ ] The crates.io dry run passes before the OIDC publish job.
- [ ] Release notes credit every meaningful contributor from merged history.
- [ ] The GitHub release, crates.io version, documentation, Pages site, and
      installer smoke jobs are independently checked after publication.
