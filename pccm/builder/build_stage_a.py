#!/usr/bin/env python3
"""PCCM Stage A build entry point.

Usage:
    python pccm/builder/build_stage_a.py [--spec PATH] [--contract PATH]
                                         [--calc-contract PATH]
                                         [--out PATH] [--quiet]

Reads the structural manifest and all FOUR contracts, generates the Stage A
workbook, emits the Stage-B inputs (build/vba/modConstants.bas,
build/stage_b_manifest.json, build/phase4_scenarios.json) and the Phase-5
generated artifacts (build/vba/modCalcContract.bas, build/phase5_cases.json), then
runs structural verification against every specification. Exits non-zero on a
specification error (2) or a verification failure (1).

The calculation contract is a REQUIRED build input as of Phase 5 Gate-A Step 3: it
is loaded, validated against the other four authorities, projected into the
workbook and checked in the generated artifact. Nothing here calculates.

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
    CalcContractError,
    ContractError,
    SpecError,
    build_workbook,
    emit_calc_artifacts,
    emit_stage_b,
    load_calc_contract,
    load_contract,
    load_driver_contract,
    load_spec,
    load_structure_contract,
    validate_calc_against,
    verify_workbook,
)

PCCM_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SPEC = PCCM_ROOT / "spec" / "workbook.yaml"
DEFAULT_CONTRACT = PCCM_ROOT / "spec" / "input_contract.yaml"
DEFAULT_DRIVERS = PCCM_ROOT / "spec" / "driver_contract.yaml"
DEFAULT_STRUCTURE = PCCM_ROOT / "spec" / "structure_contract.yaml"
DEFAULT_CALC = PCCM_ROOT / "spec" / "calc_contract.yaml"
DEFAULT_BUILD_DIR = PCCM_ROOT / "build"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build the PCCM Stage A workbook.")
    parser.add_argument("--spec", type=Path, default=DEFAULT_SPEC,
                        help=f"path to the workbook manifest (default: {DEFAULT_SPEC})")
    parser.add_argument("--contract", type=Path, default=DEFAULT_CONTRACT,
                        help=f"path to the input contract (default: {DEFAULT_CONTRACT})")
    parser.add_argument("--drivers", type=Path, default=DEFAULT_DRIVERS,
                        help=f"path to the driver contract (default: {DEFAULT_DRIVERS})")
    parser.add_argument("--structure", type=Path, default=DEFAULT_STRUCTURE,
                        help=f"path to the structure contract (default: {DEFAULT_STRUCTURE})")
    parser.add_argument("--calc-contract", type=Path, default=DEFAULT_CALC,
                        help=f"path to the calculation contract (default: {DEFAULT_CALC})")
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
    say(f"  drivers  : {args.drivers}")
    say(f"  structure: {args.structure}")
    say(f"  calc     : {args.calc_contract}")

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
    try:
        drivers = load_driver_contract(args.drivers)
    except ContractError as error:
        print(f"DRIVER CONTRACT ERROR: {error}", file=sys.stderr)
        return 2
    try:
        structure = load_structure_contract(args.structure)
    except ContractError as error:
        print(f"STRUCTURE CONTRACT ERROR: {error}", file=sys.stderr)
        return 2
    # The calculation contract is loaded AND cross-validated before any emission,
    # so a contract that disagrees with the other four authorities fails the build
    # rather than producing a workbook nobody should trust. The rules themselves
    # live in the accepted Step-1 loader and are not restated here.
    try:
        calc = load_calc_contract(args.calc_contract)
        validate_calc_against(calc, spec, contract, drivers, structure)
    except CalcContractError as error:
        print(f"CALCULATION CONTRACT ERROR: {error}", file=sys.stderr)
        return 2
    except ContractError as error:
        print(f"CALCULATION CONTRACT ERROR: {error}", file=sys.stderr)
        return 2

    out_path = args.out or (DEFAULT_BUILD_DIR / spec.stage_a_filename)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    say(f"  model    : {spec.model['short_name']} {spec.model['model_version']}")
    say(f"  phase    : {spec.model['build_phase']}")
    say(f"  sheets   : {len(spec.sheets)}")
    say(f"  inputs   : {len(contract.inputs)}   tables: {len(contract.all_tables)}")
    say(f"  drivers  : {len(drivers.all_registers)} registers")
    say(f"  structure: {len(structure.all_grids)} grids, "
        f"{len(structure.counters)} ID counters, {len(structure.buttons)} buttons")
    say(f"  calc     : contract {calc.version}, sheet {calc.sheet}, "
        f"{len(calc.scalar_blocks)} scalar blocks, {len(calc.all_tables)} tables, "
        f"FP_VERSION {calc.fingerprint_version}")

    try:
        workbook, metadata = build_workbook(spec, contract, drivers, structure, calc)
    except RuntimeError as error:
        print(f"CROSS-SPECIFICATION ERROR: {error}", file=sys.stderr)
        return 2
    except ContractError as error:
        print(f"CROSS-CONTRACT ERROR: {error}", file=sys.stderr)
        return 2
    workbook.save(out_path)
    workbook.close()

    artifacts = emit_stage_b(out_path.parent, spec, contract, drivers, structure)
    calc_artifacts = emit_calc_artifacts(out_path.parent, spec, calc)

    say(f"  built    : {out_path}")
    say(f"  emitted  : {artifacts.module_path}")
    say(f"  emitted  : {artifacts.manifest_path}")
    say(f"  emitted  : {artifacts.scenario_path}")
    say(f"  emitted  : {calc_artifacts.module_path}")
    say(f"  emitted  : {calc_artifacts.cases_path}")
    say(f"  stamped  : builder {metadata.builder_version}, {metadata.build_timestamp}")
    say("")
    say("Structural verification:")

    result = verify_workbook(out_path, spec, contract, drivers, structure, calc)
    say(result.report())

    if not result.ok:
        print("BUILD FAILED: structural verification did not pass", file=sys.stderr)
        return 1

    say("")
    say("Stage A build complete.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
