# Spec Versioning and Authoring Tools

RusDox document specs have an explicit contract version:

    version: 1

The same model is supported in YAML, JSON, and TOML. New specs and published
examples must declare the version. During the pre-1.0 transition, a missing
field is read as legacy version 1 so old files keep rendering, but the schema
and editor tooling report the omission and recommend migration.

## Compatibility policy

- This release supports document spec version 1.
- Unknown future versions fail semantic validation before any output is written.
- Additive optional fields and new block variants may appear in compatible
  RusDox minor releases.
- Existing field meaning, enum meaning, and escaping behavior do not change
  inside a spec version.
- A removal, incompatible default, or expression-semantics change requires a
  new spec version and a migration path.
- The Rust API and CLI still follow crate SemVer independently of the spec
  version.

This is the frozen v1 contract. Read [Stability, versioning, and support](stability.md)
for the cross-surface SemVer rules, deprecation window, MSRV, and release policy.

Generate the canonical schema directly from the installed binary:

    rusdox schema
    rusdox schema --output rusdox-spec-v1.schema.json

The checked-in schema is generated from the Rust/serde types. Its public ID is
https://othmaneblial.github.io/rusdox/schema/rusdox-spec-v1.schema.json and it
applies to YAML, JSON, and TOML object shapes.

## Migration

Inspect a file without changing it:

    rusdox migrate legacy.yaml --check

Print the migrated form to stdout:

    rusdox migrate legacy.yaml

Write atomically in place or to a separate file:

    rusdox migrate legacy.yaml --in-place
    rusdox migrate legacy.json --output current.json

YAML migration preserves leading author comments. JSON and TOML are parsed and
pretty-printed. Future versions are never silently downgraded.

## Deterministic expressions

Nested paths work in values:

    text: "{{ customer.address.city }}"

Supported filters are upper, lower, title, trim, and default("text"):

    text: "{{ customer.name | trim | title }}"
    text: "{{ customer.owner | default(\"unassigned\") | upper }}"

Use doubled delimiters for literal braces:

    text: "Show {{{{ customer.name }}}} literally"

This becomes Show {{ customer.name }} literally. Values are always inserted as
document text and escaped by the DOCX/PDF writers. There is no raw OOXML, shell,
function-call, network, or general-purpose evaluation mode.

The bounded conditional block selects one branch by truthiness or scalar
equality:

    - type: when
      path: customer.active
      equals: true
      blocks:
        - type: body
          text: Active account
      otherwise:
        - type: body
          text: Inactive account

Conditions and filters share the same behavior in YAML, JSON, and TOML.

## VS Code

The dependency-free extension in editors/vscode bundles the schema and provides
block/field completion, hover help, enum suggestions, and inline diagnostics.
Its local language-service tests run with:

    npm test --prefix editors/vscode

The extension never uploads document content.
