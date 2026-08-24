#!/usr/bin/env python3

import argparse
from datetime import datetime, timezone
import gzip
from pathlib import Path
import subprocess
import tarfile
import zipfile


def main() -> int:
    parser = argparse.ArgumentParser(description="Create one deterministic RusDox release archive")
    parser.add_argument("--binary", required=True)
    parser.add_argument("--format", choices=("tar.gz", "zip"), required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--source-date-epoch", type=int)
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    binary = (root / args.binary).resolve()
    output = (root / args.output).resolve()
    if not binary.is_file():
        raise SystemExit(f"binary does not exist: {binary}")
    epoch = args.source_date_epoch
    if epoch is None:
        epoch = int(subprocess.check_output(
            ["git", "log", "-1", "--format=%ct"], cwd=root, text=True
        ).strip())
    output.parent.mkdir(parents=True, exist_ok=True)
    data = binary.read_bytes()
    if args.format == "tar.gz":
        with output.open("wb") as raw:
            with gzip.GzipFile(filename="", mode="wb", fileobj=raw, mtime=0) as compressed:
                with tarfile.open(fileobj=compressed, mode="w", format=tarfile.PAX_FORMAT) as archive:
                    info = tarfile.TarInfo(binary.name)
                    info.size = len(data)
                    info.mode = 0o755
                    info.uid = info.gid = 0
                    info.uname = info.gname = ""
                    info.mtime = epoch
                    from io import BytesIO
                    archive.addfile(info, BytesIO(data))
    else:
        moment = datetime.fromtimestamp(max(epoch, 315532800), tz=timezone.utc)
        info = zipfile.ZipInfo(binary.name, moment.timetuple()[:6])
        info.compress_type = zipfile.ZIP_DEFLATED
        info.create_system = 3
        info.external_attr = 0o100755 << 16
        with zipfile.ZipFile(output, "w") as archive:
            archive.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    print(output.relative_to(root))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
