#!/usr/bin/env python3
"""PCCM Phase 1 build entry point.

Usage:
    python pccm/builder/build_skeleton.py [--spec PATH] [--out PATH] [--quiet]

Reads the workbook manifest, generates the structural skeleton and runs
structural verification against the manifest that produced it. Exits non-zero
on a specification error or a verification failure.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pccm_builder import (  # noqa: E402
    BUILDER_VERSION,
    SpecError,
    build_workbook,
    load_spec,
    verify_workbook,
)

PCCM_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SPEC = PCCM_ROOT / "spec" / "workbook.yaml"
DEFAULT_BUILD_DIR = PCCM_ROOT / "build"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build the PCCM workbook skeleton.")
    parser.add_argument("--spec", type=Path, default=DEFAULT_SPEC,
                        help=f"path to the workbook manifest (default: {DEFAULT_SPEC})")
    parser.add_argument("--out", type=Path, default=None,
                        help="output path (default: <pccm>/build/<manifest filename>)")
    parser.add_argument("--quiet", action="store_true", help="suppress progress output")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    say = (lambda *a: None) if args.quiet else (lambda *a: print(*a))

    say(f"PCCM skeleton builder {BUILDER_VERSION}")
    say(f"  manifest : {args.spec}")

    try:
        spec = load_spec(args.spec)
    except SpecError as error:
        print(f"SPECIFICATION ERROR: {error}", file=sys.stderr)
        return 2

    out_path = args.out or (DEFAULT_BUILD_DIR / spec.skeleton_filename)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    say(f"  model    : {spec.model['short_name']} {spec.model['model_version']}")
    say(f"  phase    : {spec.model['build_phase']}")
    say(f"  sheets   : {len(spec.sheets)}")

    workbook, metadata = build_workbook(spec)
    workbook.save(out_path)
    workbook.close()

    say(f"  built    : {out_path}")
    say(f"  stamped  : builder {metadata.builder_version}, {metadata.build_timestamp}")
    say("")
    say("Structural verification:")

    result = verify_workbook(out_path, spec)
    say(result.report())

    if not result.ok:
        print("BUILD FAILED: structural verification did not pass", file=sys.stderr)
        return 1

    say("")
    say("Phase 1 skeleton build complete.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
