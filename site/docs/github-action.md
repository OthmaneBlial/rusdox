# GitHub Action

The reusable RusDox action validates specs, attaches source-located errors and
warnings to pull-request lines, renders editable DOCX plus native PDF, runs the
parity contract, and keeps review evidence inside the calling repository's
GitHub Actions run.

## Minimal workflow

```yaml
name: Verify documents

on:
  pull_request:
    paths: ["documents/**"]

permissions:
  contents: read

jobs:
  documents:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v5
      - uses: OthmaneBlial/rusdox@c74a0f44bf03065fe5ca4d4d215bd78cac59f8b5 # v1.0.0
        with:
          input: documents
          upload-reports: "false"
          comment: "false"
```

The example is read-only and pins the reviewed v1.0.0 commit rather than mutable
`main`. Opt in to report retention with `upload-reports: "true"`. PR comments
also require `pull-requests: write`, `comment: "true"`, and a `github-token`.

## What it does

The action builds the native binary from the exact RusDox ref selected in
`uses`, unless `binary` points to an already-built executable. It runs
`rusdox validate --format json` first, converts source locations into native
GitHub error and warning annotations, and only renders valid specs. It then runs
`rusdox verify`, writes the same HTML/JSON/page evidence used locally, adds a Job
Summary, and optionally upserts one concise pull-request comment.

By default, only `reports/` is uploaded as an immutable GitHub Actions artifact
for 14 days. Editable DOCX and rendered PDF files stay on the ephemeral runner.
The PR comment contains filenames and check counts—not document contents—and no
external diff service receives the files. Page snapshots and HTML reports can
still reveal document content, so set `upload-reports: false` for confidential
workloads that must not be retained even in the repository's Actions storage.

## Inputs and outputs

Common inputs:

- `input`: required path to one spec or a directory;
- `config`: optional `rusdox.toml` path;
- `output-root`: defaults to `.rusdox-artifacts`;
- `visual-baseline` and `visual-threshold`: optional deterministic page diff;
- `upload-reports`: keeps parity evidence in GitHub Actions, default `true`;
- `comment`: updates a PR summary, default `true` when a token is provided;
- `github-token`: token with `pull-requests: write` for comments;
- `retention-days`: report retention, default 14.

Outputs are `passed`, `reports-path`, and `artifact-url`. A validation or parity
failure fails the workflow even though the evidence upload and PR summary still
run.

## Visual baselines

Page snapshots are renderer- and platform-specific. Check in a baseline from
the same runner image and pass its directory explicitly:

```yaml
      - uses: OthmaneBlial/rusdox@c74a0f44bf03065fe5ca4d4d215bd78cac59f8b5 # v1.0.0
        with:
          input: documents/monthly.yaml
          visual-baseline: tests/rusdox-pages/linux-x86_64
          visual-threshold: "0.002"
          github-token: ${{ secrets.GITHUB_TOKEN }}
```

The comment links to the authenticated artifact rather than embedding a public
image URL. A fork pull request with a read-only token still gets line
annotations, parity checks, and an artifact; only the optional comment is
skipped with a warning.

## Workflow recipes

Copy the complete, non-executing examples from
[`examples/github-actions/`](../examples/github-actions/):

- release notes generated when a GitHub Release is published;
- an on-demand invoice verification job;
- compliance evidence with strict visual baselines on pull requests;
- a scheduled operating report with short artifact retention.

## Why there is no container image

The native binary is the default distribution. As of 2026-08-24, public issues
and Discussions contain no Docker, container, or OCI request. RusDox therefore
does not publish a speculative image with a second patching and provenance
surface. Open a feature request with the runtime, architecture, and deployment
constraint if an image would solve a real adoption blocker; that evidence is
the publication gate.
