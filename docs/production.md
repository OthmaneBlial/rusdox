# Production and Batch Rendering

RusDox exposes bounded building blocks for a job runner without pretending that
the local HTTP adapter is a hosted service. A production owner supplies the
queue, authentication, tenant isolation, process/container limits, retention,
and observability. RusDox supplies deterministic preflight, bounded concurrency,
cooperative cancellation, and in-memory artifacts.

## Bounded Rust batches

```rust
use std::sync::Arc;
use rusdox::config::RusdoxConfig;
use rusdox::{
    BatchLimits, BatchRenderer, BatchRequest, CancellationToken, NativeRenderer,
    RenderRequest, RenderSource, SpecFormat, RENDERER_API_VERSION,
};

let batch = BatchRenderer::new(
    Arc::new(NativeRenderer::new(RusdoxConfig::default())),
    BatchLimits {
        max_jobs: 100,
        max_concurrency: 4,
        max_source_bytes_per_job: 2 * 1024 * 1024,
        max_total_source_bytes: 32 * 1024 * 1024,
    },
)?;
let cancellation = CancellationToken::new();
let results = batch.render(vec![BatchRequest {
    id: "invoice-42".into(),
    request: RenderRequest {
        renderer_api_version: RENDERER_API_VERSION,
        source: RenderSource::Inline {
            format: SpecFormat::Yaml,
            content: "version: 1\nblocks: []\n".into(),
        },
        emit_pdf: true,
    },
}], &cancellation)?;
assert_eq!(results[0].id, "invoice-42");
# Ok::<(), rusdox::DocxError>(())
```

Results preserve request order. Invalid capacity or aggregate-source limits fail
before worker creation. Per-item parse/layout failures do not erase successful
siblings. Cancelling a token prevents queued jobs and is observed between parse,
composition, DOCX, and PDF stages.

The test suite renders a 16-document real batch, proves a three-worker ceiling
under contention, cancels queued work, and rejects an aggregate source-memory
budget before any renderer is called. It does not claim hard memory isolation:
use a process/container memory limit when tenants are mutually untrusted.

## Service-owned profiles

The local protocol defaults to the stricter hosted profile:

```bash
rusdox serve stdio --limits-profile hosted --output-root target/jobs
rusdox serve http --limits-file examples/config/hosted-limits.toml \
  --port 4175 --output-root target/jobs
```

`--limits-file` accepts one complete TOML or JSON `InputLimits` object, rejects
unknown fields and zero/inconsistent ceilings, and overrides
`--limits-profile`. Profiles belong to the service owner; protocol clients have
no field that can increase them. The HTTP health response publishes effective
limits so operators can confirm what is active.

## Operational boundary

Prefer one isolated worker process per trust domain. Keep the HTTP adapter on
loopback behind an authenticated owner-controlled gateway, provide a private
output root, bound `--max-requests` to recycle long-lived workers, and use
read-only/curated font and template directories. Treat artifacts as sensitive
tenant data and delete them according to the owner's retention policy.

See [Input safety](input-safety.md), the [v1 security review](security-review-v1.md),
and the [integration protocol](integrations.md) for the complete boundary.
