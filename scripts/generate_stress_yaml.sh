#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
repo_root="$(cd "$script_dir/.." && pwd)"
cd "$repo_root"

release_flag=()
if [[ "${1:-}" == "--release" ]]; then
  release_flag=(--release)
fi

cargo run "${release_flag[@]}" --quiet --bin generate_stress_1000_pages_yaml
