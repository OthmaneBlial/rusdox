# v1 Security Review

Review date: **2026-09-01**
Scope: RusDox v1 document/spec ingestion, Word templates, visual decoding,
local integration transports, batch execution, and release supply chain.

RusDox is a local document renderer, not an isolation boundary. Its parsers
treat every input as untrusted and apply finite work and allocation ceilings,
but callers processing hostile documents should still run the process with
least privilege, a private working directory, and operating-system resource
limits.

## Threat model and trust boundaries

The attacker may control a DOCX/ZIP package, OOXML parts and relationships, a
YAML/JSON/TOML spec, local include contents, image bytes, SVG markup, Word
template expressions, and a protocol request. They may attempt path traversal,
entity expansion, decompression or image bombs, recursive expansion, excessive
concurrency, overwrite outside an output root, or access to local files through
an SVG reference. Network services, remote asset fetching, macros, embedded
executables, and raw OOXML template injection are outside the supported input
model.

## ZIP and OOXML

- A DOCX is rejected before parsing when the compressed archive, entry count,
  one uncompressed entry, aggregate uncompressed bytes, or per-entry expansion
  ratio exceeds `InputLimits`.
- Every member path is enclosed and normalized; absolute paths, parent
  traversal, backslash aliases, duplicate names, unsafe relationship targets,
  missing targets, and undeclared content types fail closed.
- XML and relationships parts have a byte ceiling. `DOCTYPE` declarations are
  forbidden, so DTD-defined entities and external entities cannot enter the
  OOXML parser. Only the XML parser's built-in character references remain.
- Opening an existing package preserves untouched parts, but validates the
  package before exposing it. Saving uses a same-directory temporary file,
  flush/sync, and atomic replacement so failure does not destroy the previous
  valid destination.

## Images and SVG

- PNG/JPEG/SVG source bytes are bounded before decode. Declared dimensions and
  requested raster targets are checked against the pixel budget before large
  allocations.
- The SVG resolver accepts bounded `data:` resources but deliberately refuses
  string/file references. An SVG therefore cannot read an arbitrary local path
  through an image `href`; RusDox does not fetch remote SVG resources.
- System font parsing remains native dependency code. Hosted deployments should
  expose only a curated font directory through their container or sandbox.

## Template expansion

- Template syntax is data substitution, not code execution: it has no network,
  filesystem, shell, evaluator, or raw-OOXML escape hatch.
- Package, XML, and data inputs retain their normal ceilings. Separate limits
  bound total placeholder/block expansions, nested partial depth, and output
  bytes for every rendered textual OOXML part.
- Partial cycles, unclosed blocks, unsupported placements, missing strict-mode
  values, and output-limit violations return structured diagnostics without
  replacing the last good output.

## Protocol, hosted use, and batch work

- `serve` defaults to the conservative `hosted` resource profile. Limits come
  from operator-owned CLI flags or a complete TOML/JSON file; a request cannot
  raise its own ceilings.
- HTTP binds to `127.0.0.1`, grants no CORS access, bounds headers and request
  bodies, sets defensive response headers, and confines relative writes below a
  fixed output root with symlink-aware containment. It is intentionally not an
  authenticated public or LAN service.
- `BatchRenderer` preflights job count, per-job bytes, and aggregate source
  bytes before workers start. A fixed worker budget bounds concurrency. A
  cooperative token cancels queued jobs and is checked between native render
  stages; it is not hard CPU preemption inside a decoder already executing.

## Verification and supply chain

- The crate's own `src/` contains no `unsafe` blocks. This is enforced by
  `scripts/check_security_review.mjs`; third-party native/parser code remains
  inside the threat model.
- Unit/integration tests cover malformed packages, DTD rejection, SVG local-file
  denial, template expansion ceilings, service-owned limits, batch load,
  concurrency, memory preflight, cancellation, and interrupted writes. Three
  libFuzzer targets compile in CI, and CodeQL scans the repository.
- `cargo audit --deny unsound` passes on the review date: there are no known
  vulnerabilities or unsound dependencies in `Cargo.lock`. RustSec separately
  marks `rustybuzz 0.20.1` (`RUSTSEC-2026-0206`) and `ttf-parser 0.25.1`
  (`RUSTSEC-2026-0192`) as unmaintained. They remain in the rendering stack
  while supported upstream replacements are assessed. The maintainer exception
  must be resolved or renewed with evidence before v1.2.0 or 2026-12-01,
  whichever comes first; any vulnerability or unsoundness advisory blocks a
  release immediately.
- Tagged binaries are built twice from the lockfile and compared byte-for-byte.
  Deterministic archives, SHA-256 checksums, SPDX SBOM, GitHub build-provenance
  attestations, and SBOM attestations are attached to every release. Installer
  smoke jobs verify the provenance before execution.

## Residual risks and non-claims

- Cancellation is stage-bound and does not forcibly interrupt a parser or image
  decoder in the middle of one call.
- Viewer rendering and font substitution can differ by operating system and
  installed fonts; reproducible bytes do not mean identical visual behavior on
  every host.
- Existing-DOCX preservation is bounded, not a guarantee that every optional or
  proprietary Word extension is understood.
- PDFs are not currently claimed to be tagged PDF, PDF/UA, or PDF/A. The
  accessibility/archival requirements and present gap are documented before any
  such marketing claim.

Report a suspected vulnerability privately through the process in
[`SECURITY.md`](../SECURITY.md). Do not attach confidential source documents to
a public issue.
