# Input Safety and Resource Limits

RusDox treats DOCX packages, document specs, includes, XML, ZIP entries, images, and SVG as untrusted input. Default limits are applied before large reads or raster allocations.

## Default ceilings

| Boundary | Default |
|---|---:|
| Compressed DOCX archive | 64 MiB |
| ZIP entries per DOCX | 4,096 |
| One uncompressed ZIP entry | 64 MiB |
| Total uncompressed DOCX | 256 MiB |
| ZIP expansion ratio per entry | 200:1 |
| One XML/relationships part | 16 MiB |
| Root spec or one YAML include | 8 MiB |
| YAML include depth | 32 |
| YAML include count | 128 |
| One PNG/JPEG source | 32 MiB |
| One SVG source | 8 MiB |
| Decoded or target visual raster | 64 million pixels |
| Template substitutions/block expansions | 100,000 |
| One rendered template XML part | 64 MiB |
| Nested template partial depth | 32 |

Unsafe ZIP paths, duplicate entries, missing required package parts, malformed XML, unresolved internal relationships, and missing content-type declarations are rejected by the package validator.

## Custom limits

Trusted workflows that legitimately need larger inputs can opt in explicitly:

```rust
use rusdox::{Document, InputLimits};

let limits = InputLimits {
    max_docx_archive_bytes: 128 * 1024 * 1024,
    max_docx_total_bytes: 512 * 1024 * 1024,
    ..InputLimits::default()
};

let document = Document::open_with_limits("large-but-trusted.docx", limits)?;
# Ok::<(), rusdox::DocxError>(())
```

`DocumentSpec::load_from_path_with_limits` and the `from_*_str_with_limits` methods cover YAML/JSON/TOML. `Visual::from_path_with_limits` and `Visual::from_bytes_with_limits` apply explicit visual ceilings.

## Hosted profile

`InputLimits::hosted()` is the conservative service profile and the default for
`rusdox serve`. Compared with the library defaults it caps a compressed DOCX at
16 MiB, total expanded DOCX data at 64 MiB, one spec at 2 MiB, one image at 8
MiB, one SVG at 2 MiB, a decoded visual at 16 million pixels, template work at
10,000 expansions, rendered template XML at 8 MiB, and include/partial depth at
12. The complete checked-in profile is [`examples/config/hosted-limits.toml`](../examples/config/hosted-limits.toml).

```bash
rusdox serve stdio --limits-profile hosted --output-root target/jobs
rusdox serve http --limits-file examples/config/hosted-limits.toml \
  --port 4175 --output-root target/jobs
```

A limit file is a complete TOML or JSON `InputLimits` value. Unknown fields,
zero ceilings, and an entry ceiling larger than the aggregate DOCX ceiling are
rejected. It is operator-owned and cannot be overridden by a protocol request.

For concurrent Rust workloads, [`BatchRenderer`](production.md) also preflights
the number of jobs, one source's bytes, aggregate source bytes, and worker
concurrency before rendering starts.

## Fuzzing

The `fuzz/` workspace contains buildable libFuzzer targets for:

- DOCX ZIP/XML open paths;
- YAML, JSON, and TOML parsing;
- PNG, JPEG, and SVG inspection/rasterization.

CI compiles all targets under nightly Rust. Maintainers can run bounded campaigns with the commands in [`fuzz/README.md`](../fuzz/README.md).

## Atomic outputs

DOCX, PDF, and serialized spec writes use a same-directory temporary file, flush and sync it, then atomically replace the destination. If rendering or writing fails, the previous known-good output remains in place and the temporary file is cleaned. Tests simulate an interrupted write and a failed PDF layout to enforce this contract.
