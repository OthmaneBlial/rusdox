# RusDox Examples

`examples/` is now a folder of document spec files, not Rust programs.

For the full user docs, start with [../docs/README.md](../docs/README.md).

You can also browse visual previews in [../docs/gallery.md](../docs/gallery.md).

Each example renders to:

- `.docx` files in `generated/`
- `.pdf` previews in `rendered/`

All examples inherit styling from config. The easiest way to tweak it is:

```bash
cargo run -- config wizard --level basic
```

That updates the user config at `~/rusdox/config.toml`. If you want this repo only to use a local override, run:

```bash
cargo run -- config wizard --path ./rusdox.toml --level basic
```

Local `./rusdox.toml` overrides the user config.

Render one example:

```bash
cargo run -- examples/board_report.yaml
```

Render every example in the folder:

```bash
cargo run -- examples
```

Create your own starter document file:

```bash
cargo run -- init-doc mydoc.yaml
```

Example files:

- `board_report.yaml`
- `client_proposal.yaml`
- `configured_studio.yaml`
- `executive_dashboard.yaml`
- `formatting_showcase.yaml`
- `hello_world.yaml`
- `invoice.yaml`
- `meeting_notes.yaml`
- `product_launch_brief.yaml`
- `project_brief.yaml`
- `table_report.yaml`
- `talent_profile.yaml`

Showcase examples:

- `executive_dashboard.yaml`: multi-section KPI summary with shaded metric cards and status tables
- `board_report.yaml`: two-page leadership report using a cover page and page break
- `client_proposal.yaml`: styled proposal with phases, pricing, and delivery plan
- `product_launch_brief.yaml`: launch narrative, milestones, readiness checks, and metrics
- `talent_profile.yaml`: polished resume/profile style document with experience and skills sections

Benchmark note:

- The generated YAML stress spec lives at `examples/stress/stress_1000_pages.yaml`
- Regenerate it with `./scripts/generate_stress_yaml.sh`
- Run the full stress flow with `./scripts/run_stress_yaml.sh`
