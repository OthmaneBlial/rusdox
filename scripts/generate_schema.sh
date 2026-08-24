#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd "$(dirname "$BASH_SOURCE")" && pwd)"
repo_root="$(cd "$script_dir/.." && pwd)"
cd "$repo_root"

mkdir -p schema editors/vscode/schema
cargo run --locked -- schema --output schema/rusdox-spec-v1.schema.json
cp schema/rusdox-spec-v1.schema.json editors/vscode/schema/rusdox-spec-v1.schema.json

echo "schema/rusdox-spec-v1.schema.json"
echo "editors/vscode/schema/rusdox-spec-v1.schema.json"
