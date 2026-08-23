# Benchmark Evidence

`protocol.json` is the versioned scenario contract. `results/` contains immutable raw reports produced by `scripts/run_benchmark_protocol.mjs`; `history.json` and `assets/benchmark-history.svg` are derived by `scripts/render_benchmark_history.mjs`.

Reports from different platforms, architectures, CPUs, or materially changed runner images are not direct regression baselines. The scheduled Ubuntu workflow compares only with `baselines/ubuntu-latest.json` and uses the materiality floors in the protocol.

To contribute a result:

1. start from a clean checkout of a named commit;
2. build the locked release binary;
3. run the full default protocol on an otherwise idle supported host;
4. inspect the JSON and ensure `git_dirty` is `false`;
5. choose a dated, host-specific filename and regenerate history;
6. explain any baseline replacement in the commit or pull request.

Do not hand-edit measured values, derived history, or the chart.
