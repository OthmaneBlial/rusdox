# Integration protocol

RusDox exposes one versioned, local-first rendering boundary instead of four
language-specific SDKs. Rust callers use the `Renderer` trait and
`NativeRenderer`; processes use the same JSON v1 request over newline-delimited
stdin/stdout or an explicitly started loopback HTTP service.

## One request, one response

Start a bounded worker:

```bash
rusdox serve stdio --limits-profile hosted \
  --output-root target/rusdox-output --max-requests 1
```

Send one compact JSON object followed by a newline:

```json
{
  "protocol_version": 1,
  "request_id": "invoice-42",
  "operation": "render",
  "source": {
    "kind": "inline",
    "format": "yaml",
    "content": "version: 1\noutput_name: invoice\nblocks:\n  - type: body\n    text: Paid\n"
  },
  "output": {
    "directory": "invoices",
    "name": "invoice-42",
    "pdf": true
  }
}
```

The response is one JSON line with the same `request_id`, `ok`, diagnostics,
per-stage timings, and each artifact's absolute path, byte count, and SHA-256.
`validate` uses the same source object but never writes files and does not
require `output`.

`source.kind` can be `inline` with `yaml`, `json`, or `toml`, or `path` with a
local spec path. Inline content is the transport-neutral/WASM-ready boundary;
path sources preserve includes and relative asset resolution for local tools.
Unknown fields, unknown versions, oversized requests, unsafe output names, and
output paths containing absolute roots or `..` are rejected explicitly.

## Loopback HTTP

Start the opt-in service only when a local tool cannot manage a child process:

```bash
rusdox serve http --limits-profile hosted \
  --port 4175 --output-root target/rusdox-output
curl -fsS http://127.0.0.1:4175/health
curl -fsS -H 'Content-Type: application/json' \
  --data-binary @request.json \
  http://127.0.0.1:4175/v1/request
```

The server always binds `127.0.0.1`, has no CORS permission, reads at most 16
KiB of headers and 2 MiB of JSON, returns `no-store` and `nosniff` headers, and
processes requests sequentially. It has no authentication and is deliberately
not a hosted, LAN, daemon, or multi-tenant API. Use stdin/stdout when possible.

`serve` uses the conservative hosted resource profile by default. A service
owner can select `--limits-profile default` for trusted local inputs or pass a
complete TOML/JSON profile through `--limits-file`; request JSON cannot increase
these ceilings. The `/health` response exposes the effective profile values.

## Rust embedding

The public object-safe boundary is independent of either transport:

```rust
use rusdox::config::RusdoxConfig;
use rusdox::{NativeRenderer, RenderRequest, RenderSource, Renderer, SpecFormat,
    RENDERER_API_VERSION};

let renderer = NativeRenderer::new(RusdoxConfig::default());
let output = renderer.render(&RenderRequest {
    renderer_api_version: RENDERER_API_VERSION,
    source: RenderSource::Inline {
        format: SpecFormat::Yaml,
        content: "version: 1\nblocks: []\n".into(),
    },
    emit_pdf: true,
})?;
assert!(output.docx.starts_with(b"PK"));
# Ok::<(), rusdox::DocxError>(())
```

The result keeps DOCX and optional PDF bytes in memory. Durable caller-visible
artifact writes belong to the protocol adapter, which lets a future WASM
implementation reuse request semantics without pretending the current native
font/PDF stack or its internal temporary path already runs in browsers.

Rust job runners can wrap the same object-safe renderer with `BatchRenderer`.
It validates batch/job source-byte budgets before starting, applies a fixed
worker ceiling, preserves result order, and accepts a shared
`CancellationToken`. See [Production and batch rendering](production.md) for a
complete example and the deployment boundary.

## Node, Python, Go, and CI

The executable examples use only each language's process standard library:

- [`node.mjs`](../examples/integrations/node.mjs)
- [`python.py`](../examples/integrations/python.py)
- [`go/main.go`](../examples/integrations/go/main.go)
- [`ci.sh`](../examples/integrations/ci.sh)

Run all available examples against the current debug binary:

```bash
node scripts/test_integrations.mjs
```

They are contract tests, not SDK scaffolds. Keeping the JSON boundary small
avoids version skew across four wrappers while demand is still forming.

## Node/WASM decision

The dated browser feasibility spike already shows that the current renderer
depends on filesystem, native font discovery, rasterization, and PDF paths. A
2026-08-24 check of public issues and Discussions found no Node package, npm,
WASM, or WebAssembly request. RusDox therefore ships no package that would expose
partial or misleading browser behavior. The new inline `RenderSource` and
object-safe `Renderer` trait remove protocol redesign as a future blocker; a
package still requires demand plus passing browser font/image/PDF fixtures.

## Markdown and rich text

RusDox does not silently reinterpret Markdown today. Before ingestion can ship,
each admitted construct must have defined DOCX/PDF semantics and parity checks:
headings map to named heading styles; paragraphs preserve hard/soft break rules;
emphasis maps to bold/italic runs; links remain interactive in both outputs;
ordered/unordered lists use semantic numbering; code keeps whitespace and font
policy; tables define alignment, wrapping, and page splitting; images require
alt text, sizing, and bounded local resources. Raw HTML, scripts, remote images,
and arbitrary CSS remain out of scope. This contract is the implementation gate,
not a claim that Markdown ingestion already exists.
