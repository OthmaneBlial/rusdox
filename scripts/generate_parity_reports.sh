#!/bin/sh
set -eu

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
repo_root=$(CDPATH= cd -- "$script_dir/.." && pwd)
scratch_dir=$(mktemp -d "${TMPDIR:-/tmp}/rusdox-parity-site.XXXXXX")
trap 'rm -rf "$scratch_dir"' EXIT HUP INT TERM

cd "$repo_root"
cargo run --quiet --locked -- verify examples --output-root "$scratch_dir" >/dev/null

destination="$repo_root/reports/gallery"
mkdir -p "$destination"

for name in \
  board-report \
  executive-dashboard \
  product-launch-brief \
  talent-profile \
  invoice \
  meeting-notes \
  dual-output-contract \
  international-scripts
do
  rm -f "$destination/$name-parity.html" "$destination/$name-parity.json"
  rm -rf "$destination/$name-pages"
  cp "$scratch_dir/reports/$name-parity.html" "$destination/$name-parity.html"
  cp "$scratch_dir/reports/$name-parity.json" "$destination/$name-parity.json"
  cp -R "$scratch_dir/reports/$name-pages" "$destination/$name-pages"
done

printf 'Generated public parity reports in %s\n' "$destination"
