# RusDox Roadmap

> From a capable document engine to the most trustworthy way to generate an editable DOCX and a faithful PDF from one source.

Last updated: 2026-08-24

## The honest starting point

RusDox already has a stronger technical foundation than its public adoption suggests:

- a pure-Rust DOCX writer and PDF renderer;
- YAML, JSON, and TOML document specs;
- variables, includes, repeaters, named styles, metadata, visuals, and tables;
- the ability to open and preserve existing DOCX packages through the Rust API;
- `validate`, `watch`, and `bench` CLI workflows;
- 169 test functions and a three-platform CI matrix;
- real generated examples, a gallery, install scripts, and a live project site;
- a measured 1,000-page dual-output benchmark.

The problem is not a lack of code. The problem is that the public product journey currently breaks before most visitors can experience that code.

Public snapshot on 2026-08-23:

- 4 GitHub stars, 0 forks, and no open issues or pull requests;
- no GitHub Release and therefore no downloadable installer assets;
- no published `rusdox` crate on crates.io or docs.rs;
- the primary install command points to a missing release asset;
- `cargo add rusdox` is documented but cannot currently work;
- the default branch is `main` while the CI badge, security policy, crate metadata, and issue links still reference `master`;
- the live documentation is loaded client-side, which weakens indexing, direct linking, and no-JavaScript access;
- a local `cargo package --no-verify` currently collects 164 files—24.5 MiB unpacked and 7.9 MiB compressed—including duplicated site files and generated DOCX/PDF artifacts instead of a deliberately small publish set;
- the speed claim is visible, but there is no public compatibility/parity report proving how the DOCX and PDF behave across real viewers.

That means a large launch today would amplify broken onboarding. Distribution and trust come first.

## The category RusDox should own

### Positioning

**One spec → editable DOCX + faithful PDF + automated parity proof, at Rust speed, without an office runtime.**

“Fast DOCX generation” is useful but crowded. “YAML to PDF” is also crowded. The unusual and defensible combination is:

1. the DOCX remains editable by the recipient;
2. the PDF is generated natively, without Word or LibreOffice;
3. both outputs come from the same typed model;
4. RusDox can prove that important content and layout survived in both;
5. the whole pipeline is deterministic, local, scriptable, and CI-friendly.

The primary audience is backend, platform, operations, finance, and developer-tooling teams that generate recurring business documents: invoices, proposals, compliance packs, reports, dashboards, and customer deliverables.

### The memorable demo

The flagship workflow should eventually be:

```bash
rusdox verify board-report.yaml
```

It should produce:

```text
generated/board-report.docx
rendered/board-report.pdf
reports/board-report-parity.html
```

The parity report should verify normalized text, headings, tables, images, page breaks, page count, and selected geometry. This is the feature people can understand, trust, share, and build CI around.

### Product principles

- **Trust before breadth.** A smaller supported surface with visible compatibility proof is better than nominal support for the entire OOXML specification.
- **Two-minute first success.** A new visitor should install RusDox and generate an impressive document in under two minutes.
- **Visual quality is a feature.** Rendering fidelity matters more to document users than a benchmark alone.
- **Editable by default.** DOCX is not an intermediate artifact; it is one of the two first-class outputs.
- **Local and deterministic.** No account, cloud service, office runtime, or hidden network dependency should be required.
- **One model, explicit parity.** A feature is complete only when its DOCX behavior, PDF behavior, validation, docs, and tests are defined.
- **Errors must teach.** Diagnostics should identify the file, YAML path, line/column when possible, cause, and likely fix.
- **Stars are an outcome.** Activation, successful installs, real examples, repeat contributors, and downstream usage are the operating metrics.

## Priority map

| Priority | Outcome | User impact | Effort | Why now |
|---|---|---:|---:|---|
| P0 | Make every advertised install path real | Very high | Low | The current primary path ends in a missing release |
| P0 | Fix `main`/`master` and public trust inconsistencies | High | Low | Broken badges and links damage confidence immediately |
| P0 | Publish a lean crate, API docs, and first release | Very high | Medium | Unlocks Rust discovery and dependable distribution |
| P0 | Define and test the DOCX/PDF parity contract | Very high | High | Turns the dual renderer into a credible differentiator |
| P1 | Add Word-native template rendering | Very high | High | Solves a proven business-document workflow |
| P1 | Add schema-driven authoring and precise diagnostics | High | Medium | Makes YAML pleasant in real editors and CI |
| P1 | Ship a zero-install interactive demo | Very high | High | Converts curiosity into a visible result immediately |
| P2 | Build a curated template and integration ecosystem | High | Medium | Creates reasons to return, contribute, and share |
| P2 | Stabilize a v1 specification and compatibility policy | High | High | Enables production adoption |

---

## Milestone 0 — Repair the front door

Target: 1–2 weeks  
Release: `v0.1.0`  
Goal: every public promise works on a clean machine.

Status: **complete (2026-08-23).** `v0.1.0` delivered the public launch; `v0.1.1` completed the native `cargo binstall` path and moved crates.io publishing to short-lived OIDC credentials.

### Distribution

- [x] Create a real `v0.1.0` tag and GitHub Release using the existing Linux, macOS Intel, macOS Apple Silicon, and Windows build matrix.
- [x] Add SHA-256 checksums for every archive.
- [x] Add build provenance attestations and an SBOM to the release workflow.
- [x] Smoke-test `scripts/install.sh` and `scripts/install.ps1` against the published assets.
- [x] Publish the library and binary to crates.io.
- [x] Verify `cargo install rusdox --locked`, `cargo add rusdox`, and the generated docs.rs site from clean environments.
- [x] Test `cargo binstall rusdox` against actual release asset names; add package metadata only if the defaults do not resolve correctly.
- [x] Add an explicit Cargo `include` list so crates.io receives source, essential docs, and required examples—not the duplicated `site/` tree or generated binary documents.
- [x] Add `CHANGELOG.md` with release dates, user-facing changes, breaking changes, and migration notes.

### Repository trust

- [x] Replace every stale `master` reference with `main`.
- [x] Fix the CI badge and confirm it reports the real workflow state.
- [x] Update `Cargo.toml` homepage/documentation URLs to the live site and docs.rs.
- [x] Enable GitHub Discussions with Announcements, Q&A, Ideas, and Show and tell.
- [x] Enable private vulnerability reporting, Dependabot alerts/security updates, and CodeQL default setup.
- [x] Route usage questions to Discussions and security reports to the private reporting channel.
- [x] Add `good first issue`, `help wanted`, `compatibility`, `template`, and `docs` labels.

### README conversion

- [x] Put a verified “Install in 10 seconds” block before the long benchmark section.
- [x] Add a 20–30 second terminal/demo recording: install → initialize → edit one value → receive DOCX and PDF.
- [x] Show one YAML input beside large, readable previews of both real outputs.
- [x] Replace the generic “Foundation” badge with a version/stability statement tied to the release.
- [x] Add a concise “Why RusDox instead of X?” table covering DOCX editability, native PDF, office-runtime dependency, templates, and parity verification.
- [x] Link the live playground/site, crates.io, docs.rs, releases, roadmap, discussions, and gallery above the fold.
- [x] State unsupported features honestly and link to the compatibility matrix.

### Documentation

- [x] Generate crawlable static HTML with a stable URL for every guide and example.
- [x] Ensure core documentation remains usable without JavaScript.
- [x] Add canonical metadata, social cards, sitemap, and searchable page titles.
- [x] Split “owner GitHub setup” instructions out of end-user documentation.
- [x] Add a troubleshooting page for fonts, asset paths, viewer differences, large documents, and install/PATH problems.

#### Exit gate

Milestone 0 is complete only when:

- all README install commands succeed in CI or on disposable clean systems;
- the four release archives, checksums, and attestations are downloadable;
- crates.io and docs.rs show `v0.1.0`;
- all public links and badges resolve;
- the packaged crate is deliberately scoped and below the registry size limit;
- a visitor can reach a real output in under two minutes.

Evidence: release workflow [32667470624](https://github.com/OthmaneBlial/rusdox/actions/runs/32667470624), three-platform main CI [32667142635](https://github.com/OthmaneBlial/rusdox/actions/runs/32667142635), CodeQL [32667142592](https://github.com/OthmaneBlial/rusdox/actions/runs/32667142592), and the [v0.1.1 release](https://github.com/OthmaneBlial/rusdox/releases/tag/v0.1.1). Published archives were independently checksum- and attestation-verified; clean `cargo install`, `cargo add`, and archive-only `cargo binstall` checks passed.

---

## Milestone 1 — Make trust measurable

Target: weeks 2–6  
Release: `v0.2.0`  
Theme: parity, compatibility, and quality.

### Define the parity contract

- [x] Publish a feature matrix with four states: supported in DOCX, supported in PDF, parity-tested, and intentionally unsupported.
- [x] Introduce `rusdox verify <spec>`.
- [x] Generate an HTML parity report for every gallery example.
- [x] Compare normalized text, block order, table content, image count/alt text, page breaks, and document metadata.
- [x] Add optional rendered-page visual diffs with documented thresholds.
- [x] Make parity failures machine-readable through JSON and meaningful exit codes.
- [x] Upload parity reports and rendered diffs as CI artifacts.

Section evidence: `rusdox verify examples` passes all 18 supported top-level inputs (17 authored document examples plus the hosted-limit TOML fixture). Each report now evaluates 21 semantic and structural checks, and every verification emits HTML/JSON evidence plus deterministic page snapshots. Nine representative reports are bundled into the static site; CI uploads the complete generated evidence set. Exit codes `0`, `1`, and `2` distinguish success, execution failure, and completed parity failure.

### Close current dual-output gaps

This section unified the two renderers around one documented model. Remaining limits are explicit contract boundaries—such as per-section geometry and advanced bidirectional shaping—not silent differences between advertised outputs.

- [x] Make page size, orientation, and margins come from the same document model in both outputs.
- [x] Render headers, footers, page numbers, and total-page fields in PDF.
- [x] Add explicit page and section break blocks.
- [x] Add hyperlinks and bookmarks in both outputs.
- [x] Add automatic table of contents/field support with a documented PDF behavior.
- [x] Improve long-table pagination, repeating headers, unsplittable rows, and overflow diagnostics.
- [x] Add merged cells, rich cell paragraphs, per-row properties, and nested tables where parity can be guaranteed.
- [x] Add footnotes before comments or tracked changes; they are more common in generated reports and compliance documents.
- [x] Test Latin, Arabic/RTL, CJK, emoji fallback, and mixed-script shaping; publish what is and is not supported.

Section evidence: `rusdox verify examples` passes all 18 supported top-level inputs with 21 checks and zero failures per report; the optional baseline diff is explicitly skipped when no baseline is supplied. `dual_output_contract.yaml` proves shared landscape geometry, visible page fields, breaks, PDF URI/GoTo annotations and outlines, TOC fields, DOCX footnotes, rich/merged/nested cells, and row pagination controls. `international_scripts.yaml` preserves representative Unicode text and records the exact shaping/fallback boundaries. The regression suite passes 189 tests and strict Clippy; unsplittable-row overflow has a measured diagnostic test. PDF nested-table layout remains deliberately bounded and complex bidirectional shaping remains intentionally unsupported rather than overstated.

### Compatibility and regression lab

- [x] Maintain representative golden fixtures, including documents opened and modified from real DOCX packages.
- [x] Validate OOXML packages and assert relationships/content types, not only successful ZIP creation.
- [x] Exercise release fixtures in available current document viewers and record pass, blocked, and not-run states for Microsoft Word, LibreOffice, Apple Preview, Adobe Acrobat, and platform-native alternatives.
- [x] Publish a dated compatibility scorecard instead of claiming universal compatibility.
- [x] Add rendered-page snapshots for every gallery example.
- [x] Add deterministic-output tests and document which metadata is intentionally variable.
- [x] Add fuzz targets and size/decompression limits for untrusted DOCX, YAML, XML, ZIP, image, and SVG inputs.
- [x] Test interrupted writes and confirm output files remain atomic and recoverable.

Section evidence: the suite opens, modifies, saves, structurally validates, and reopens an external macOS-produced DOCX while preserving untouched theme/meta parts. The public OOXML validator checks XML well-formedness, content types, unique relationship IDs, safe targets, and target existence. The 2026-08-24 scorecard pins fixture hashes and distinguishes a Pages open smoke, a Preview/Quick Look open-and-raster pass, a blocked Word for Mac run, and unavailable Word/Windows, LibreOffice, and Acrobat environments without inventing passes. Ubuntu CI owns 26 exact page baselines for all 17 examples. DOCX and PDF are byte-deterministic in same-host tests; intentional cross-host variability is documented. Three libFuzzer targets compile and each completed a 100-run smoke campaign. Default/custom resource ceilings cover DOCX/ZIP/XML/spec/includes/PNG/JPEG/SVG/raster allocations, and simulated interrupted writes preserve the prior DOCX/PDF destination. The complete suite passes 189 tests plus strict Clippy.

### Reproducible performance

- [x] Replace the single-machine benchmark claim with a reproducible benchmark protocol.
- [x] Record CPU, OS, Rust version, input hash, output sizes, peak memory, and command flags.
- [x] Add small, medium, and 1,000-page benchmark tiers.
- [x] Track DOCX, PDF, dual-output, validation, and existing-DOCX open/save separately.
- [x] Run scheduled benchmark CI and flag material regressions without making noisy PR checks.
- [x] Publish benchmark history as data and charts, not a manually updated screenshot only.

Section evidence: protocol `2026-08-24` runs 13 isolated release processes: four pipelines for 1-page, 4-page, and 1,000-page tiers plus external-DOCX open/save. Each raw report pins host/toolchain metadata, input SHA-256 and bytes, exact flags, per-stage distributions, output sizes, and normalized peak RSS. Clean checked-in runs cover Apple M2/darwin-arm64 and GitHub Actions AMD EPYC/Linux-x64; manual Ubuntu run `32673209069` passed and supplied the comparable scheduled baseline. The Monday/manual-only workflow applies dual relative/absolute floors (20% + 5 ms runtime; 25% + 16 MiB RSS) so pull requests stay quiet while material regressions fail. `benchmarks/history.json` and the accessible SVG chart are regenerated from immutable raw JSON rather than hand-edited.

#### Exit gate

- every gallery example has a green parity report;
- core page controls behave consistently in DOCX and PDF;
- compatibility claims link to dated viewer evidence;
- benchmark results are independently reproducible;
- untrusted input has documented resource limits and fuzz coverage.

Exit evidence: all 18 supported top-level inputs pass 21 parity checks, while the 17 authored document examples own 26 exact Ubuntu page snapshots; shared page controls are covered by the dual-output contract; viewer claims link to the 2026-08-24 hash-pinned scorecard; two clean benchmark hosts implement the documented protocol; and the input-safety contract is backed by limits, atomic-write recovery tests, and three buildable fuzz targets.

---

## Milestone 2 — Build the “wow” workflow

Target: weeks 6–12  
Release: `v0.3.0`  
Theme: Word-native templates, live feedback, and excellent authoring.

### Word-native template rendering

RusDox already opens existing DOCX packages and preserves untouched parts at the Rust layer. Turn that technical capability into a product workflow:

```bash
rusdox template inspect proposal.docx
rusdox template render proposal.docx data.json
rusdox template verify proposal.docx data.json
```

- [x] Define a minimal, versioned placeholder syntax.
- [x] Support scalar replacements, nested values, loops, conditions, and reusable subtemplates.
- [x] Preserve designer-authored styles, sections, headers/footers, media, relationships, and untouched package parts.
- [x] Report template errors with part name, paragraph/table location, placeholder, and suggested fix.
- [x] Add safe handling for missing values and an explicit strict mode.
- [x] Generate the edited DOCX, native PDF, and parity report from one command.
- [x] Ship at least three designer-authored templates: invoice, proposal, and compliance/board report.

Section evidence: template syntax v1 supports split-run scalar placeholders, nested paths, one-based loop indices, complete paragraph/table-row loops, truthy conditions with else, five deterministic filters, and cycle-safe inline partials without a raw-XML escape hatch. `template inspect/render/verify` expose human and JSON reports; strict failures name the OOXML part, paragraph/row, expression, fix, and preserve the previous destination. Integration tests prove byte preservation for untouched styles, headers, footers, media, relationships, and section properties. Three styled RTF sources exported by the macOS text system (invoice, proposal, landscape board report) render from sample JSON into visually inspected DOCX/PDF output, deterministic snapshots, and green parity reports.

This is a higher-value differentiator than adding dozens of low-level OOXML builders with no end-user workflow.

### Schema-first authoring

- [x] Add a version field to document specs and publish the compatibility policy.
- [x] Publish generated JSON Schema for YAML, JSON, and TOML authoring.
- [x] Add `rusdox schema` and `rusdox migrate` commands.
- [x] Provide autocomplete, hover documentation, enum suggestions, and inline diagnostics through a lightweight VS Code extension.
- [x] Add source spans so validation errors point to line and column.
- [x] Add declarative conditions, nested path access, filters, and escaping rules.
- [x] Keep expressions intentionally small and deterministic; do not invent a general-purpose programming language.

Section evidence: all published specs now declare version 1; missing versions remain a documented legacy bridge while future versions fail before output. The canonical 2020-12 JSON Schema is generated from the Rust/serde model, includes bounded include/repeat/when authoring blocks, ships in the crate and VS Code extension, and is identical for YAML, JSON, and TOML object shapes. The schema and migrate commands support stdout, atomic writes, in-place migration, and CI checks. Semantic validation reports one-based source spans. Cross-format tests cover nested paths, when branches, upper/lower/title/trim/default filters, literal braces, rejection of unknown filters, and a dependency-free editor service with completion, hover, enums, and inline diagnostics.

### Zero-install playground

- [x] Run a feasibility spike for a WASM renderer with bundled, redistributable fonts.
- [x] If feasible, add a fully local browser playground: edit YAML, preview PDF, download DOCX/PDF, and load gallery templates.
- [x] If full browser rendering is not yet feasible, ship an interactive spec builder with pre-generated verified outputs and a transparent capability boundary.
- [x] Make every playground example reproducible with the CLI.
- [x] Add “Open this example” links from the README, gallery, docs, and parity reports.
- [x] Never upload user document content by default.

Section evidence: `cargo check --locked --target wasm32-unknown-unknown --lib` succeeds on Rust 1.95.0, but the dated feasibility record rejects a premature full-render claim because PDF fonts still come from the host filesystem, output APIs remain path-based, and no licensed multiscript font bundle has passed browser parity. The shipped fallback is a same-origin, dependency-free playground for six gallery fixtures: edits and downloads use browser memory only, the HTML preview is visibly structural rather than PDF layout, and verified DOCX/PDF/parity links are disabled after any source edit. Every fixture shows its exact CLI command; pure parser/safety tests cover all six inputs; README, gallery, generated docs, site cards, and generated parity HTML link directly into the selected example.

### Developer feedback loop

- [x] Add `rusdox dev` as the polished successor to raw watch mode.
- [x] Open a local preview showing the latest PDF, validation issues, timings, and output paths.
- [x] Preserve the current successful output when a rebuild fails.
- [x] Debounce file changes and explain which input/config/asset triggered each rebuild.
- [x] Add `--open`, `--json`, and quiet CI-friendly modes.

Section evidence: `rusdox dev` serves a dependency-free dashboard from loopback only with no-store headers, a restrictive CSP, a live status JSON route, fixed latest-DOCX/PDF routes, explicit failure output, and parse/validate/compose/DOCX/PDF/total timings. The polling watcher now follows local `path:` includes and assets recursively, coalesces save bursts, and labels triggers as input, config, or asset/include. JSON Lines tests prove a future-version failure preserves the byte-identical previous DOCX and previous artifact manifest; server tests cover the CSP/status contract; a quiet one-build test emits no terminal noise; `watch` remains backward compatible on the shared debouncer.

#### Exit gate

- a non-Rust user can customize a Word-designed template from JSON;
- a YAML author gets autocomplete and source-located errors;
- the flagship demo works without cloning the repository;
- one command produces the editable DOCX, faithful PDF, and parity evidence.

---

## Milestone 3 — Create an ecosystem, not a pile of examples

Target: months 3–5  
Release: `v0.4.0`  
Theme: templates, integrations, and contribution loops.

### Curated template registry

- [x] Define a small signed manifest format with license, author, preview, supported RusDox version, inputs, and output hashes.
- [x] Add `rusdox template search/list/add/update`.
- [x] Start curated: every template must have screenshots, sample data, tests, parity evidence, accessibility notes, and a clear license.
- [x] Add categories for invoices, proposals, reports, compliance, HR, education, and operations.
- [x] Generate a preview page automatically for template pull requests.
- [x] Feature a “template of the month” and credit contributors prominently.
- [x] Keep third-party templates outside the core crate so the engine remains lean.

Section evidence: `registry/index.json` is an Ed25519-signed v1 contract for three hash-pinned Word templates and seven curated discovery categories. The CLI pins the default public key, accepts explicit keys for private registries, rejects non-loopback HTTP, bounds registry and asset reads, enforces RusDox version ranges, verifies source hashes before atomic installation, and supports text or JSON list/search/add/update output. Integration tests prove a signed local registry installs verified bytes and a one-byte manifest change fails before installation. The Node registry builder independently verifies the signature, 21 source/evidence hashes, preview alt text, licenses, documented inputs, reading order, language, and non-color-only semantics; it generates an accessible public gallery and a pull-request review artifact. Registry content and verified outputs remain outside the published crate.

### GitHub-native automation

- [x] Ship a reusable GitHub Action that validates specs, renders outputs, and uploads parity reports.
- [x] Add examples for release notes, invoices, compliance evidence, and scheduled reporting workflows.
- [x] Support annotations that attach RusDox validation errors to pull-request lines.
- [x] Offer a pull-request visual-diff comment without uploading private documents to a third-party service.
- [x] Publish minimal Docker and OCI images only if real users need them; the native binary remains the default.

Section evidence: the root `action.yml` is a source-built composite action that builds the native binary selected by the caller's pinned `uses` ref, or accepts a prebuilt binary for a faster trusted job. It validates before render, converts JSON source spans to GitHub error/warning annotations, produces the full DOCX/PDF/parity bundle, uploads only `reports/` as an immutable Actions artifact, writes a Job Summary, and upserts a metadata-only PR comment whose authenticated link never exposes a public document URL. Upload and comment are independently optional for confidential or read-only fork runs. Node contract tests cover relative file annotations and pass/fail summaries; the main CI workflow dogfoods the local action. Four copyable workflows cover release notes, invoices, strict compliance baselines, and scheduled reports. A dated public issue/Discussion check found no container demand, so no speculative Docker/OCI image or second patching surface was published; the documented evidence gate keeps the native binary as the default.

### Integration surfaces

- [x] Stabilize a renderer boundary that can support Rust embedding, WASM, and a local stdin/stdout JSON protocol.
- [x] Add a tiny opt-in local HTTP service only after the core protocol is stable.
- [x] Provide official examples for calling the binary from Node, Python, Go, and CI without maintaining four premature native SDKs.
- [x] Evaluate a Node/WASM package based on demonstrated demand and browser feasibility.
- [x] Add Markdown/rich-text ingestion only after parity behavior is defined for each supported construct.

Section evidence: public `Renderer`/`NativeRenderer` types accept path-backed or self-contained YAML/JSON/TOML sources and return DOCX plus optional PDF bytes in memory, leaving durable writes to adapters. Protocol v1 maps that boundary to one request/response JSON contract with IDs, diagnostics, timings, atomic output paths, byte counts, and SHA-256. `rusdox serve stdio` reserves stdout for newline-delimited responses; the later HTTP adapter reuses the exact request at `POST /v1/request`, binds only `127.0.0.1`, disables CORS, sets defensive response headers, caps headers/JSON, and confines relative outputs beneath a fixed root. Library and CLI tests cover inline rendering, validation source spans, version rejection, path-escape rejection, real stdio DOCX/PDF output, and real loopback HTTP validation. Dependency-free Node, Python, Go, and shell/CI examples execute as contract tests. A dated demand check found no Node/npm/WASM request, and the browser feasibility blockers remain, so no misleading package ships; the inline source and object-safe trait preserve a future path. Markdown also remains deliberately unimplemented after publishing an explicit construct-by-construct DOCX/PDF parity gate—raw HTML, remote assets, scripts, and CSS are out of scope rather than ambiguously interpreted.

### Contributor experience

- [x] Publish 10 genuinely small starter issues with fixtures and acceptance criteria.
- [x] Document the architecture from spec parsing through composition, DOCX packaging, PDF layout, validation, and CLI.
- [x] Add fixture-generation and visual-diff commands to the contributor workflow.
- [x] Recognize contributors in release notes and the site.
- [x] Use Discussions to collect real output examples and prioritize compatibility work.
- [x] Publish a short governance and maintainer policy before adding committers.

Section evidence: `.github/starter-issues.json` defines exactly ten idempotently
published `good first issue` tasks, each bound to one immutable fixture, a user
reason, and three acceptance criteria. `contributor_lab.mjs` lists and prepares
isolated tasks and exposes the same `verify` visual-baseline path used by CI;
contract tests prevent the queue or commands from drifting. The architecture
map traces every input through validation, the typed model, OOXML packaging,
native PDF layout, parity, and adapters, while the governance policy documents
review, release, security, committer, inactivity, and conflict decisions before
access expands. Git-derived contributor credit is checked in CI and rendered in
the release notes and static site. Dedicated Show-and-tell and compatibility
Discussions request only reproducible, non-confidential evidence. The final M3
adoption gate remains honestly open until an external human contribution is
actually merged and released.

#### Exit gate

- external contributors can add a verified template without understanding OOXML internals;
- another repository can use RusDox through an official GitHub Action;
- the first integration protocol is stable and documented;
- at least one release includes meaningful external contributions.

---

## Milestone 4 — Production contract

Target: months 5–8  
Release: `v1.0.0`  
Theme: stability, international documents, accessibility, and long-term trust.

### Stable contracts

- [x] Freeze and document spec version 1.
- [x] Publish SemVer guarantees for the Rust API, CLI, spec schema, template syntax, and output behavior.
- [x] Ship migration tooling for pre-1.0 specs.
- [x] Define a minimum supported Rust version and test it in CI.
- [x] Reach strong rustdoc coverage for the supported public API.
- [x] Publish a deprecation window and supported release policy.

Section evidence: one policy maps the Rust API, CLI, document spec/schema, Word
template syntax, renderer/protocol, and output semantics to explicit compatibility
and breaking-change rules. An executable gate keeps all four independent format
constants at v1, pins the canonical required-version schema, and verifies the
editor copy byte-for-byte. The existing migration command covers check, stdout,
separate output, and atomic in-place upgrades while rejecting future versions.
Cargo declares Rust 1.88 and a dedicated CI job compiles every feature on that
exact toolchain. The library denies missing public-item documentation and broken
intra-doc links—rustdoc now reports complete item coverage—and tagged releases
run `cargo-semver-checks` before crates.io publication. The six-month/two-minor
deprecation window, supported-minor policy, MSRV notice rule, security exception,
and evidence-based release checklist are all public.

### Production quality

- [x] Complete the compatibility matrix for the supported v1 feature set.
- [x] Add load, memory, concurrency, and cancellation tests for batch generation.
- [x] Support configurable resource limits for hosted or multi-tenant use.
- [x] Complete the security review of ZIP/XML/image parsing and template expansion.
- [x] Provide reproducible builds and signed release artifacts.
- [x] Add a long-term benchmark dashboard with explicit regression budgets.

Section evidence: `compatibility/v1-feature-contract.json` records 31
capabilities and maps every one of the 23 spec block variants exactly once.
`BatchRenderer` preflights job/per-job/aggregate-source budgets, preserves order,
caps workers, renders a 16-document real load, and proves queued cancellation;
native cancellation is observed between parse, compose, DOCX, and PDF stages.
`InputLimits::hosted()` plus complete TOML/JSON profiles keep ceilings under
operator control, including template expansion/depth/output limits. The dated
v1 security review is backed by a source contract, DTD rejection, SVG external
resource denial, three fuzz targets, CodeQL, and the official weekly RustSec
audit; two unsound dependency advisories were removed and two upstream
maintenance notices remain explicit residual risks. Release targets build twice
to identical bytes, use deterministic archives, publish checksums/SPDX SBOM and
GitHub attestations, and verify provenance before install. The accessible
dashboard applies absolute runtime and peak-memory budgets to all 13 reproducible
Ubuntu scenarios in addition to the existing relative regression policy.

### International and accessible output

- [x] Graduate RTL and CJK support from experimental only when parity fixtures pass.
- [x] Define font embedding, fallback, licensing, and substitution behavior.
- [x] Preserve meaningful image alt text and document language metadata.
- [x] Research tagged PDF and PDF/A requirements before advertising accessible or archival output.
- [x] Add accessibility checks to parity reports where the formats expose equivalent semantics.

Section evidence: `compatibility/international-readiness.json` makes RTL and
CJK graduation executable; both remain experimental because bidi/shaping or
kinsoku/font-profile and representative-viewer gates are still false even
though semantic and deterministic baseline fixtures pass. PDF rendering admits
only embeddable outline TrueType fonts, records each resolved family,
permission, glyph count, and missing character, and documents that operators
own font licensing and pinned-font deployment. Non-blank visual alt text is
required and preserved in both DOCX drawing descriptions; BCP 47-style document
language reaches DOCX `dc:language` and PDF `/Lang`. The 21-check parity contract
compares alt text and language wherever applicable. Sourced PDF/UA and PDF/A
research records the required structure tree, reading order, artifacts, XMP,
ICC, font, exact veraPDF flavour, and human assistive-technology gates; current
PDF evidence keeps tagged-PDF and PDF/A claims false.

#### Exit gate

`v1.0.0` is justified only when teams can upgrade with confidence, supported viewers behave predictably, the spec has a migration story, and security/performance limits are explicit.

---

## Growth loop after the product gates are green

Do not run a large “please star this” campaign. Build a repeatable proof-and-sharing loop:

1. **Show a surprising result.** A short clip turns readable YAML or a Word template into an editable DOCX, faithful PDF, and green parity report.
2. **Make it instantly reproducible.** One verified install command and one small input file.
3. **Publish the evidence.** Link the benchmark data, compatibility scorecard, and generated artifacts.
4. **Give visitors a useful starting point.** A polished template or GitHub Action they can use in their own repository.
5. **Invite a specific contribution.** Ask for a viewer fixture, template, language case, or integration—not generic “help wanted.”
6. **Ship on a cadence.** Each release should have one clear story, a real demo, and a small set of excellent changes.

Launch channels can include the Rust community, developer-tool communities, document-automation communities, Show HN, and relevant GitHub showcases, but only after the installation and parity gates pass. The message should be the technical result and reproducible proof, not the star count.

### Metrics that predict durable adoption

Track:

- successful clean-machine installs by supported method;
- median time from README arrival to first generated document;
- playground example opens and completed downloads, without collecting document content;
- crates.io downloads and downstream dependents;
- GitHub Action usage in external repositories;
- parity reports generated in CI;
- template installs and template contributors;
- issue response time, first-time contributors, and repeat contributors;
- release adoption and upgrade lag;
- stars and forks as lagging discovery signals.

Do not add invasive analytics to a local document tool. Prefer public ecosystem signals, explicit opt-in feedback, and privacy-preserving aggregate site metrics if metrics are needed at all.

## Explicitly not now

These ideas would dilute the wedge before it is proven:

- a cloud account system or hosted document storage;
- a full collaborative WYSIWYG word processor;
- AI-generated content as a core feature;
- support for every Pandoc input/output format;
- a plugin system before the core render/parity contracts are stable;
- dozens of SDKs before the stdin/stdout or WASM boundary is proven;
- comments and tracked changes before common report features and fidelity are solid;
- a giant uncurated template marketplace;
- benchmark marketing without reproducible data and visual-quality evidence.

## Next 10 issues to open

These are the best first concrete issues, in order:

1. **Release v0.1.0 and smoke-test every advertised installer.**
2. **Publish a lean crates.io package and verify docs.rs.**
3. **Replace every `master` link with `main` and repair the CI badge.**
4. **Generate static, linkable documentation pages.**
5. **Define the DOCX/PDF parity matrix and report format.**
6. **Make PDF honor document page setup, headers, footers, and numbering.**
7. **Add hyperlinks, bookmarks, explicit breaks, and viewer fixtures.**
8. **Generate JSON Schema and line/column validation diagnostics.**
9. **Prototype `rusdox template inspect/render` on an existing DOCX.**
10. **Prototype the local/WASM playground and document feasibility findings.**

## Definition of success

RusDox becomes genuinely interesting when a developer can say:

> “I can give it one readable spec or designer-authored Word template, receive an editable Word file and a faithful PDF without Office installed, and prove in CI that both outputs contain what I expect.”

If RusDox delivers that experience reliably, quickly, and visibly, stars have a reason to follow.

## Public reference snapshot

Current project state:

- [RusDox repository](https://github.com/OthmaneBlial/rusdox)
- [RusDox releases](https://github.com/OthmaneBlial/rusdox/releases)
- [RusDox live site](https://othmaneblial.github.io/rusdox/)
- [crates.io package URL](https://crates.io/crates/rusdox)
- [docs.rs package URL](https://docs.rs/rusdox)

Relevant product comparisons:

- [docx-rs](https://github.com/bokuweb/docx-rs)
- [Typst](https://github.com/typst/typst)
- [docx for TypeScript](https://github.com/dolanmiu/docx)
- [python-docx](https://github.com/python-openxml/python-docx)
- [Docxtemplater](https://github.com/open-xml-templating/docxtemplater)
- [python-docx-template](https://github.com/elapouya/python-docx-template)
- [Pandoc](https://github.com/jgm/pandoc)

Distribution and community guidance:

- [Publishing on crates.io](https://doc.rust-lang.org/cargo/reference/publishing.html)
- [cargo-binstall support conventions](https://github.com/cargo-bins/cargo-binstall/blob/main/SUPPORT.md)
- [GitHub artifact attestations](https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations)
- [GitHub Discussions quickstart](https://docs.github.com/en/discussions/quickstart)
