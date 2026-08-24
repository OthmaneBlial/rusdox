# Stability, versioning, and support

The RusDox v1 contract has five public surfaces and takes effect with the
`v1.0.0` release. Each surface is versioned deliberately, so a crate version
bump never silently changes the meaning of a document specification or local
protocol request.

## Contract map

| Surface | Version | Compatible change | Breaking change |
|---|---:|---|---|
| Rust crate API | Cargo SemVer | New documented APIs; additions to explicitly non-exhaustive types | Removing/renaming a public item, narrowing accepted inputs, or changing observable semantics |
| CLI | crate SemVer | New commands or optional flags; clearer human text | Removing/renaming commands or flags, changing defaults, paths, exit codes, or versioned JSON fields |
| document spec/schema | `version: 1` | Optional fields, new blocks with explicit parity behavior, stricter rejection only for previously invalid input | Reinterpreting a valid field, default, enum, expression, escaping rule, or output semantic |
| Word template syntax | syntax `1` | New opt-in filters or diagnostics | Changing placeholder, loop, condition, strictness, escaping, or preservation behavior |
| renderer/local protocol | API/protocol `1` | New optional request fields and response fields that tolerant readers can ignore | Reusing a field with new meaning, removing fields, or changing validation/output confinement |

DOCX and PDF bytes are deterministic for the same RusDox version, inputs,
configuration, and font environment. Cross-version compatibility promises the
document semantics listed in the compatibility matrix, not an identical ZIP or
PDF byte stream. Any intentional layout change requires parity evidence and a
changelog entry.

## Frozen spec version 1

`SPEC_VERSION`, the generated schema identifier, every checked-in example, and
the migration command agree on version 1. Unknown future versions fail before
an output is replaced. Version 1 will not reinterpret existing valid input.

An additive feature may stay in spec v1 only when all of these are true:

1. old valid documents keep the same meaning;
2. the new field is optional or has a behavior-preserving default;
3. YAML, JSON, TOML, the schema, editor tooling, validation, DOCX, PDF, and
   parity evidence move together;
4. unsupported readers fail clearly rather than guessing;
5. migration is unnecessary for existing input.

Anything else requires spec version 2 and an atomic migration path. The current
pre-1.0 migration remains available as `rusdox migrate --check`, stdout,
`--output`, or `--in-place`; future migrations follow the same recoverable
transaction model.

## Deprecation window

A v1 API, CLI, or syntax deprecation is announced in rustdoc or `--help`, the
changelog, and migration documentation. Deprecated behavior remains functional
for at least two minor releases and six months, whichever is longer. Removal
waits for the next major release. Security fixes may reject previously accepted
malicious input immediately, with a security advisory and narrow migration note.

Human-readable CLI prose may improve without notice. Scripts must consume the
documented JSON forms and their version fields, not parse terminal sentences.

## Supported releases

- The latest v1 minor receives bug, compatibility, and security fixes.
- The previous v1 minor receives critical security fixes for six months after
  its successor, when a previous minor exists.
- Older minors and pre-1.0 releases are unsupported; users receive a migration
  path but no promise of backports.
- Release archives, checksums, SBOM, and GitHub build provenance are retained.
- Viewer claims are tied to dated, hash-pinned evidence rather than support by
  brand name alone.

## Minimum supported Rust version

The MSRV is Rust 1.88.0 and is declared through Cargo's `rust-version` field.
CI compiles all features on that exact toolchain. Raising the MSRV is a minor
release change, announced at least 90 days in advance when a security dependency
does not force an earlier move. Patch releases do not raise it.

## Rust API documentation and SemVer gate

The library denies missing public-item docs and broken intra-doc links. CI builds
the full library rustdoc with those lints as errors, so the supported public API
has 100% item coverage rather than an aspirational badge. Release tags also run
`cargo-semver-checks` against the latest published crate before crates.io upload.

The automated SemVer scan supplements review; it cannot judge every behavioral
contract. A release PR must still classify changes across all five surfaces,
include migrations where required, and link parity evidence for output changes.

## Release decision

The release checklist must record the crate/spec/template/protocol versions,
MSRV result, rustdoc result, semantic scan, migrations, compatibility evidence,
security review, benchmark budgets, and signed artifact verification. A green
build without those receipts is not sufficient to publish v1.
