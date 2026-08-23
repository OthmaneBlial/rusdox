#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd "$(dirname "$BASH_SOURCE")" && pwd)"
repo_root="$(cd "$script_dir/.." && pwd)"
cd "$repo_root"
output_root="template-evidence"

mkdir -p "$output_root"

for name in invoice proposal board-report; do
  cargo run --locked -- template verify \
    "templates/$name/template.docx" \
    "templates/$name/data.json" \
    --name "$name" \
    --strict \
    --output-root "$output_root/$name" \
    --format json > "$output_root/$name-summary.json"
done

echo "$output_root"
