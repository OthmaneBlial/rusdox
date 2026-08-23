#!/bin/sh
set -eu

if [ "$#" -ne 1 ]; then
  printf 'Usage: %s <verify-reports-directory>\n' "$0" >&2
  exit 2
fi

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
repo_root=$(CDPATH= cd -- "$script_dir/.." && pwd)
source_dir=$1
destination="$repo_root/tests/golden/pages/linux-x86_64"

if [ ! -d "$source_dir" ]; then
  printf 'Reports directory does not exist: %s\n' "$source_dir" >&2
  exit 2
fi

mkdir -p "$destination"
find "$destination" -type f -name 'page-*.png' -delete
find "$destination" -mindepth 1 -type d -empty -delete

copied=0
for pages_dir in "$source_dir"/*-pages; do
  [ -d "$pages_dir" ] || continue
  output_name=$(basename "$pages_dir" -pages)
  target="$destination/$output_name"
  mkdir -p "$target"
  for page in "$pages_dir"/page-*.png; do
    [ -f "$page" ] || continue
    cp "$page" "$target/$(basename "$page")"
    copied=$((copied + 1))
  done
done

if [ "$copied" -eq 0 ]; then
  printf 'No page snapshots found under %s\n' "$source_dir" >&2
  exit 2
fi

printf 'Updated %s Linux visual baselines in %s\n' "$copied" "$destination"
