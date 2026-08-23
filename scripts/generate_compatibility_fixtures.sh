#!/bin/sh
set -eu

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
repo_root=$(CDPATH= cd -- "$script_dir/.." && pwd)
scratch_dir=$(mktemp -d "${TMPDIR:-/tmp}/rusdox-compatibility.XXXXXX")
trap 'rm -rf "$scratch_dir"' EXIT HUP INT TERM

cd "$repo_root"
cargo run --quiet --locked -- verify examples/dual_output_contract.yaml \
  --output-root "$scratch_dir" \
  --format json >/dev/null

destination="$repo_root/compatibility/fixtures"
mkdir -p "$destination"
cp "$scratch_dir/generated/dual-output-contract.docx" "$destination/dual-output-contract.docx"
cp "$scratch_dir/rendered/dual-output-contract.pdf" "$destination/dual-output-contract.pdf"

printf 'Updated compatibility fixtures in %s\n' "$destination"
