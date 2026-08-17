#!/usr/bin/env python3
"""PCCM Stage A build entry point.

Usage:
    python pccm/builder/build_stage_a.py [--spec PATH] [--contract PATH]
                                         [--out PATH] [--quiet]

Reads the structural manifest and the input contract, generates the Stage A
workbook and runs structural verification against both. Exits non-zero on a
specification error (2) or a verification failure (1).

Stage A produces .xlsx only. The .xlsm, VBA and CodeName assignment belong to
Stage B on Windows, which this script never touches.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pccm_builder import (  # noqa: E402
    BUILDER_VERSION,
    ContractError,
    SpecError,
    build_workbook,
    load_contract,
    load_spec,
    verify_workbook,
)

PCCM_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SPEC = PCCM_ROOT / "spec" / "workbook.yaml"
DEFAULT_CONTRACT = PCCM_ROOT / "spec" / "input_contract.yaml"
DEFAULT_BUILD_DIR = PCCM_ROOT / "build"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build the PCCM Stage A workbook.")
    parser.add_argument("--spec", type=Path, default=DEFAULT_SPEC,
                        help=f"path to the workbook manifest (default: {DEFAULT_SPEC})")
    parser.add_argument("--contract", type=Path, default=DEFAULT_CONTRACT,
                        help=f"path to the input contract (default: {DEFAULT_CONTRACT})")
    parser.add_argument("--out", type=Path, default=None,
                        help="output path (default: <pccm>/build/<manifest filename>)")
    parser.add_argument("--quiet", action="store_true", help="suppress progress output")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    say = (lambda *a: None) if args.quiet else (lambda *a: print(*a))

    say(f"PCCM Stage A builder {BUILDER_VERSION}")
    say(f"  manifest : {args.spec}")
    say(f"  contract : {args.contract}")

    try:
        spec = load_spec(args.spec)
    except SpecError as error:
        print(f"SPECIFICATION ERROR: {error}", file=sys.stderr)
        return 2
    try:
        contract = load_contract(args.contract)
    except ContractError as error:
        print(f"INPUT CONTRACT ERROR: {error}", file=sys.stderr)
        return 2

    out_path = args.out or (DEFAULT_BUILD_DIR / spec.stage_a_filename)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    say(f"  model    : {spec.model['short_name']} {spec.model['model_version']}")
    say(f"  phase    : {spec.model['build_phase']}")
    say(f"  sheets   : {len(spec.sheets)}")
    say(f"  inputs   : {len(contract.inputs)}   tables: {len(contract.all_tables)}")

    workbook, metadata = build_workbook(spec, contract)
    workbook.save(out_path)
    workbook.close()

    say(f"  built    : {out_path}")
    say(f"  stamped  : builder {metadata.builder_version}, {metadata.build_timestamp}")
    say("")
    say("Structural verification:")

    result = verify_workbook(out_path, spec, contract)
    say(result.report())

    if not result.ok:
        print("BUILD FAILED: structural verification did not pass", file=sys.stderr)
        return 1

    say("")
    say("Stage A build complete.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
