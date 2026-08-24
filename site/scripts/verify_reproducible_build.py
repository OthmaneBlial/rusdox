#!/usr/bin/env python3

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys


def main() -> int:
    parser = argparse.ArgumentParser(description="Build RusDox twice and compare release bytes")
    parser.add_argument("--target", required=True)
    parser.add_argument("--binary", required=True)
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    epoch = subprocess.check_output(
        ["git", "log", "-1", "--format=%ct"], cwd=root, text=True
    ).strip()
    command = [
        "cargo", "build", "--release", "--locked", "--bin", "rusdox",
        "--target", args.target,
    ]
    target_dir = root / "target" / "reproducible-build" / args.target
    diagnostics = root / "target" / "reproducibility-diagnostics" / args.target
    clean(diagnostics)
    clean(target_dir)
    first_binary = build(command, root, epoch, target_dir, args.target, args.binary)
    first = digest(first_binary)
    first_snapshot = diagnostics / "first" / args.binary
    first_snapshot.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(first_binary, first_snapshot)
    clean(target_dir)
    second_binary = build(command, root, epoch, target_dir, args.target, args.binary)
    second = digest(second_binary)
    second_snapshot = diagnostics / "second" / args.binary
    second_snapshot.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(second_binary, second_snapshot)
    canonical = root / "target" / args.target / "release" / args.binary
    canonical.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(second_binary, canonical)
    if first != second:
        print(f"release binary is not reproducible: {first} != {second}", file=sys.stderr)
        print(f"diagnostic binaries preserved in {diagnostics}", file=sys.stderr)
        return 1
    clean(diagnostics)

    report = {
        "schema_version": 1,
        "target": args.target,
        "binary": args.binary,
        "source_date_epoch": int(epoch),
        "cargo_locked": True,
        "incremental": False,
        "build_path_strategy": "fixed-clean-target-directory",
        "clean_target_directory_between_builds": True,
        "msvc_linker_brepro": args.target.endswith("-pc-windows-msvc"),
        "builds_compared": 2,
        "sha256": first,
        "reproducible": True,
    }
    destination = root / args.report
    destination.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(f"reproducible {args.target} {first}")
    return 0


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def clean(target_dir: Path) -> None:
    if target_dir.exists():
        shutil.rmtree(target_dir)


def build(command, root: Path, epoch: str, target_dir: Path, target: str, binary: str) -> Path:
    environment = os.environ.copy()
    environment["SOURCE_DATE_EPOCH"] = epoch
    environment["CARGO_INCREMENTAL"] = "0"
    environment["CARGO_TARGET_DIR"] = str(target_dir)
    if target.endswith("-pc-windows-msvc"):
        flags = environment.get("RUSTFLAGS", "").strip()
        environment["RUSTFLAGS"] = f"{flags} -C link-arg=/Brepro".strip()
    subprocess.run(command, cwd=root, env=environment, check=True)
    artifact = target_dir / target / "release" / binary
    if not artifact.is_file():
        raise FileNotFoundError(f"release binary was not created: {artifact}")
    return artifact


if __name__ == "__main__":
    raise SystemExit(main())
