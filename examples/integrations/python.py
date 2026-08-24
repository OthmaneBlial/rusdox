#!/usr/bin/env python3

import json
import os
import pathlib
import subprocess
import sys

binary = os.environ.get("RUSDOX_BIN", "rusdox")
output_root = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "target/integration-python").resolve()
request = {
    "protocol_version": 1,
    "request_id": "python-example",
    "operation": "render",
    "source": {
        "kind": "inline",
        "format": "yaml",
        "content": "version: 1\noutput_name: python-report\nblocks:\n  - type: title\n    text: Python integration\n  - type: body\n    text: No native SDK or office runtime is required.\n",
    },
    "output": {"directory": "python", "name": "python-report", "pdf": True},
}
completed = subprocess.run(
    [binary, "serve", "stdio", "--output-root", str(output_root), "--max-requests", "1"],
    input=json.dumps(request) + "\n",
    text=True,
    capture_output=True,
    check=True,
)
response = json.loads(completed.stdout)
if not response["ok"]:
    raise RuntimeError(response.get("error", {}).get("message", "RusDox request failed"))
print(json.dumps(response, indent=2))
