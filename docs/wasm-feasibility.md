# Browser rendering feasibility

Date: 2026-08-24

## Decision

RusDox ships an interactive, local-first spec playground now. It does not claim to render edited DOCX or PDF files in the browser yet.

The Rust library compiles for the browser target, which makes a future WASM renderer plausible. The current end-to-end render path is not browser-ready, however: PDF rendering discovers fonts from the host operating system, public PDF APIs write through filesystem paths, and the project does not yet ship a pinned redistributable font bundle with license evidence and multiscript coverage.

The playground therefore follows the explicit fallback in the roadmap. It previews the document structure, keeps edits in the current browser tab, downloads edited YAML, and exposes pre-generated DOCX/PDF artifacts only while the source exactly matches a checked-in example. This avoids presenting an HTML approximation as PDF parity.

## Reproduce the spike

The spike ran successfully on Rust 1.95.0 with target wasm32-unknown-unknown:

    rustup target add wasm32-unknown-unknown
    cargo check --locked --target wasm32-unknown-unknown --lib

Success means the library dependency graph compiles for the target. It does not prove that browser APIs, font loading, memory behavior, downloads, or layout parity work.

## Gates for a full renderer

A future full browser renderer must satisfy all of these gates:

1. Return DOCX and PDF bytes from public in-memory APIs without temporary paths.
2. Bundle pinned, redistributable TrueType fonts and their license texts.
3. Cover the scripts promised by the compatibility contract, including transparent shaping limitations.
4. Expose a small wasm-bindgen API with deterministic errors and bounded allocations.
5. Pass the same semantic parity fixtures as the CLI in real Chromium, Firefox, and WebKit runs.
6. Prove through a network test that author content is never uploaded.

Until those gates pass, the CLI remains the source of truth for edited DOCX/PDF output.

## Integration update

The transport redesign gate is now smaller: public `Renderer` and
`NativeRenderer` APIs accept an inline YAML/JSON/TOML `RenderSource` and return
DOCX plus optional PDF bytes, and JSON protocol v1 carries the same source over
stdin/stdout. The native PDF implementation still uses a temporary filesystem
path internally, so this does not satisfy the browser runtime gate or change the
playground's claim.

A 2026-08-24 search of public RusDox issues and Discussions found no Node, npm,
WASM, or WebAssembly request. With no demonstrated package demand and font,
image, memory, and cross-browser fixtures still open, an npm/WASM package would
currently expose a partial product. The package decision remains “not yet”; the
new boundary avoids locking a future implementation to CLI-specific JSON.
