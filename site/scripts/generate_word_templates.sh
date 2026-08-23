#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd "$(dirname "$BASH_SOURCE")" && pwd)"
repo_root="$(cd "$script_dir/.." && pwd)"

if ! command -v textutil >/dev/null 2>&1; then
  echo "textutil is required to regenerate externally authored DOCX templates." >&2
  exit 1
fi

for name in invoice proposal board-report; do
  source_path="$repo_root/templates/$name/source.rtf"
  output_path="$repo_root/templates/$name/template.docx"
  textutil -convert docx -output "$output_path" "$source_path"
  echo "$output_path"
done
