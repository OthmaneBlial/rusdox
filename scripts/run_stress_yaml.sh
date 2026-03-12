#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
repo_root="$(cd "$script_dir/.." && pwd)"
cd "$repo_root"

release_flag=()
if [[ "${1:-}" == "--release" ]]; then
  release_flag=(--release)
fi

echo "Generating examples/stress/stress_1000_pages.yaml"
"$script_dir/generate_stress_yaml.sh" "${1:-}"

echo "Running 1000-page YAML stress benchmark"
cargo run "${release_flag[@]}" --bin stress_1000_pages
