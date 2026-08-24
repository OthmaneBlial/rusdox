# Architecture

RusDox has one typed document model and several deliberately thin entry points.
The CLI, Rust API, Word-template workflow, GitHub Action, stdin protocol, and
loopback service converge before packaging; they do not maintain separate DOCX
or PDF renderers.

## End-to-end flow

```text
YAML / JSON / TOML / Word template / Rust caller
                    │
                    ▼
        parse + bounded composition expansion
                    │
                    ▼
       semantic validation + source locations
                    │
                    ▼
          typed Document + Stylesheet model
                  ┌─┴───────────────┐
                  ▼                 ▼
       OOXML package writer    native PDF layout
       XML + rels + ZIP        fonts + pages + links
                  └─┬───────────────┘
                    ▼
       structural + semantic + visual parity
                    │
                    ▼
       CLI / Rust bytes / local protocol response
```

## Authoring and parsing

`src/spec.rs` owns the versioned high-level `DocumentSpec`. It parses YAML,
JSON, and TOML under `InputLimits`. `src/spec_expand.rs` performs bounded
variables, includes, repeaters, conditions, and deterministic filters before
deserialization. It rejects cycles, excess includes, future spec versions, and
general-purpose expression syntax.

`src/source.rs` maps semantic paths back to one-based line/column spans. The
generated schema in `src/schema.rs` is the contract consumed by the CLI, checked
JSON Schema, and VS Code tooling. Migration belongs in the CLI because it is a
filesystem transaction; version interpretation belongs in the library.

## Validation gate

`src/validate.rs` checks the spec and active config before any durable render.
It catches invalid references, colors, table shapes, blank or missing values,
assets, and geometry. Render commands call the same gate used by `validate`;
adapters must not invent a weaker validation path.

The `Renderer` boundary in `src/renderer.rs` accepts file-backed or inline
sources, exposes validation diagnostics, composes through `Studio`, and returns
DOCX/PDF bytes. `src/protocol.rs` adds correlation IDs, output confinement,
hashes, and JSON response semantics without duplicating rendering behavior.

## Composition model

`src/studio.rs` maps `DocumentSpec` blocks and configured styles into the typed
`Document` model. `Document` retains ordered paragraphs, tables, metadata,
sections, headers/footers, page numbering, visuals, relationships, and styles.
Lower-level types live in focused modules such as `paragraph.rs`, `run.rs`,
`table.rs`, `visual.rs`, `layout.rs`, and `metadata.rs`.

Word-native templates enter through `src/template.rs`. Template expansion
changes only supported placeholders and structural blocks, then reuses the same
`Document`/PDF/parity path. Untouched ZIP parts and relationships are preserved.

## DOCX packaging

`src/document.rs` coordinates save/open behavior and atomic replacement.
`src/xml_utils.rs` serializes and parses WordprocessingML. Media, numbering,
styles, metadata, headers, footers, relationships, content types, and custom
parts are assembled into the OOXML ZIP package. `src/package_validate.rs`
checks XML well-formedness, declared content types, and internal relationship
targets under resource ceilings.

The DOCX writer is not a generic ZIP passthrough: open/modify/save preserves
parts the model does not own, while generated parts remain deterministic.

## Native PDF layout

The PDF path in `src/studio.rs` lays out the same `Document`, not a DOCX
conversion. It resolves page geometry, shapes text with configured/system font
fallback, paginates paragraphs and tables, draws images and vector-derived
rasters, emits bookmarks and links, and can write deterministic page snapshots.
Unsupported equivalence is documented in the compatibility matrix rather than
hidden by an HTML or office-runtime conversion.

## Parity and regression evidence

`src/parity.rs` compares the expected typed projection with the reopened DOCX
and PDF evidence. Reports include semantic checks, package/PDF checks, artifact
hashes, and optional rendered-page comparisons. Golden pages live under
`tests/golden/pages/<platform>/`; compatibility fixtures and dated viewer
results live under `compatibility/`.

The CLI `verify` command is the public orchestration layer. The GitHub Action,
template evidence scripts, gallery generator, and contributor visual-diff lab
all call it instead of reimplementing comparisons.

## Boundaries and invariants

- Author data stays local unless a caller explicitly uploads its own artifacts.
- Every parser and transport has a declared size or count limit.
- Durable file replacement is atomic; a failed build preserves the previous
  successful destination where the command promises recovery.
- DOCX and PDF claims require executable parity evidence.
- Browser previews and external viewer observations are labeled separately from
  native render proof.
- Protocol and schema versions reject unknown futures instead of guessing.
- Core rendering remains independent of GitHub, Pages, Node, Python, or Go.

## Where to make a change

| Change | Start here | Required evidence |
|---|---|---|
| New spec field or block | `src/spec.rs`, `src/schema.rs` | parse round-trip, migration decision, validation, DOCX/PDF parity |
| OOXML behavior | `src/xml_utils.rs`, `src/document.rs` | package validation, open/save preservation, viewer fixture when relevant |
| PDF layout | `src/studio.rs` | semantic parity plus platform page baseline |
| Validation rule | `src/validate.rs`, `src/source.rs` | valid/invalid fixtures and exact source location |
| Word template syntax | `src/template.rs` | strict/lenient diagnostics, package preservation, template evidence |
| CLI or protocol | `src/bin/rusdox.rs`, `src/protocol.rs` | text/JSON contract test, bounded failure, cross-platform path behavior |
| Documentation site | `scripts/build_site.mjs` | generated-site check, local links, mobile/browser QA |

Read [Contributing](../CONTRIBUTING.md) for commands and
[Governance](../GOVERNANCE.md) for decision and maintainer policy.
