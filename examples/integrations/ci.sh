#!/usr/bin/env sh
set -eu

rusdox_bin=${RUSDOX_BIN:-rusdox}
output_root=${1:-target/integration-ci}
request='{"protocol_version":1,"request_id":"ci-example","operation":"render","source":{"kind":"path","path":"examples/hello_world.yaml"},"output":{"directory":"ci","name":"ci-report","pdf":true}}'

response=$(printf '%s\n' "$request" | "$rusdox_bin" serve stdio \
  --output-root "$output_root" --max-requests 1)
printf '%s\n' "$response"
printf '%s\n' "$response" | grep -q '"ok":true'
test -f "$output_root/ci/ci-report.docx"
test -f "$output_root/ci/ci-report.pdf"
