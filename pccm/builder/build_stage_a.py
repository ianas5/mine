#!/usr/bin/env python3
"""PCCM Stage A build entry point.

Usage:
    python pccm/builder/build_stage_a.py [--spec PATH] [--contract PATH]
                                         [--calc-contract PATH]
                                         [--out PATH] [--quiet]

Reads the structural manifest and all SIX contracts, generates the Stage A
workbook, emits the Stage-B inputs (build/vba/modConstants.bas,
build/stage_b_manifest.json, build/phase4_scenarios.json), the Phase-5 generated
artifacts (build/vba/modCalcContract.bas, build/phase5_cases.json) and the
Phase-6 generated artifacts (build/vba/modSimContract.bas,
build/phase6_cases.json), then runs structural verification against every
specification. Exits non-zero on a specification error (2) or a verification
failure (1).

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
    emit_inspection,
    emit_sim_artifacts,
    emit_phase7_acceptance,
    emit_phase8_charts,
    emit_phase9_model_check,
    apply_protection,
    emit_protection_projection,
    emit_methodology_projection,
    emit_reset_projection,
    resolve_unlocked,
    emit_phase8_dashboard,
    emit_phase8_results,
    emit_sim_gate_b_artifacts,
    emit_stage_b,
    SimContractError,
    load_calc_contract,
    load_contract,
    load_driver_contract,
    load_sim_contract,
    load_spec,
    load_structure_contract,
    validate_calc_against,
    validate_sim_against,
    verify_workbook,
)

PCCM_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SPEC = PCCM_ROOT / "spec" / "workbook.yaml"
DEFAULT_CONTRACT = PCCM_ROOT / "spec" / "input_contract.yaml"
DEFAULT_DRIVERS = PCCM_ROOT / "spec" / "driver_contract.yaml"
DEFAULT_STRUCTURE = PCCM_ROOT / "spec" / "structure_contract.yaml"
DEFAULT_CALC = PCCM_ROOT / "spec" / "calc_contract.yaml"
DEFAULT_SIM = PCCM_ROOT / "spec" / "sim_contract.yaml"
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
    parser.add_argument("--sim-contract", type=Path, default=DEFAULT_SIM,
                        help=f"path to the simulation contract (default: {DEFAULT_SIM})")
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
    say(f"  sim      : {args.sim_contract}")

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
    # The simulation contract is a REQUIRED build input as of Phase 6 Step 1. It
    # is loaded and cross-validated so a contract that disagrees with the other
    # five authorities fails the build.
    #
    # It emits NO Phase-6 WORKBOOK content and NO executable simulation code -
    # the workbook gains no simulation result, no formula, no _SimData row and no
    # Phase-6 publication. As of Step 5 it IS projected into two external
    # generated Stage-A artefacts: build/vba/modSimContract.bas, a constants-only
    # module no Phase-6 VBA owner imports yet, and build/phase6_cases.json, the
    # conformance corpus the later implementation steps assert against.
    try:
        import yaml as _yaml

        sim = load_sim_contract(args.sim_contract)
        validate_sim_against(
            sim, spec, contract, drivers, structure,
            _yaml.safe_load(args.calc_contract.read_text(encoding="utf-8")),
        )
    except SimContractError as error:
        print(f"SIMULATION CONTRACT ERROR: {error}", file=sys.stderr)
        return 2
    except ContractError as error:
        print(f"SIMULATION CONTRACT ERROR: {error}", file=sys.stderr)
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
    say(f"  sim      : contract {sim.version}, RNG_VERSION {sim.rng_version}, "
        f"SIM_METHOD_VERSION {sim.sim_method_version}, "
        f"max iterations {sim.max_iterations_representable}")

    try:
        workbook, metadata = build_workbook(spec, contract, drivers, structure, calc, sim)
    except RuntimeError as error:
        print(f"CROSS-SPECIFICATION ERROR: {error}", file=sys.stderr)
        return 2
    except ContractError as error:
        print(f"CROSS-CONTRACT ERROR: {error}", file=sys.stderr)
        return 2
    # P10-2A. PROTECTION IS MARKED BEFORE THE FILE IS SAVED, because the
    # locked/unlocked state of a cell is part of the cell. The policy is
    # resolved from the contracts that already declare which inputs are
    # editable; nothing here spells a range.
    unlocked = resolve_unlocked(structure, contract, drivers)
    apply_protection(workbook, unlocked)
    workbook.save(out_path)
    workbook.close()
    protection = emit_protection_projection(
        out_path.parent / "phase10_protection_inspection.json", structure, unlocked)
    print(f"  protection: {len(protection['sheets'])} sheets, "
          f"{sum(s['unlocked_count'] for s in protection['sheets'])} unlocked cells, "
          f"passwordless={protection['passwordless']}")

    # P10-2B. WHERE A LATER WINDOWS RUN MUST LOOK to prove that Reset Results
    # cleared every publication and changed no input and no identity. Addresses
    # only: no expected value, no reset, and nothing here runs anything.
    reset = emit_reset_projection(
        out_path.parent / "phase10_reset_inspection.json",
        structure, contract, drivers, calc, sim)
    print(f"  reset    : "
          f"{sum(len(block['cells']) for block in reset['preserved']['editable_inputs'])} "
          f"editable input cells preserved, "
          f"{len(reset['publications']['simulation']['cleared'])} simulation "
          f"publication groups cleared")

    # P10-3. WHERE EVERY METHODOLOGY LINE LANDED, and the manifest wording that
    # landed there - including the Build Metadata rows, which is how a later run
    # confirms the workbook shows the versions the authorities actually hold.
    methodology = emit_methodology_projection(
        out_path.parent / "phase10_methodology_inspection.json", spec, metadata)
    print(f"  method   : {len(methodology['sections'])} sections, "
          f"{len(methodology['lines'])} written lines, "
          f"{len(methodology['metadata'])} build metadata rows")

    artifacts = emit_stage_b(out_path.parent, spec, contract, drivers, structure)
    calc_artifacts = emit_calc_artifacts(out_path.parent, spec, calc)
    # Phase 6 Step 5. Emitted from the accepted authorities and from the accepted
    # Step-2/3/4 reference; nothing here implements a generator, a sampler or a
    # simulation of its own.
    sim_artifacts = emit_sim_artifacts(out_path.parent, spec, sim, contract, calc)
    # Identities only, for the Gate-B Windows harness. No expected value lives
    # here; phase5_cases.json remains the sole expected-value authority.
    inspection = emit_inspection(out_path.parent, calc, contract)
    # The Phase-6 Gate-B pair, for the Windows harness only: addresses in one,
    # oracle expectations in the other. No workbook content, no VBA, no layout.
    sim_gate_b = emit_sim_gate_b_artifacts(out_path.parent, spec, sim, contract, calc)
    # THE PHASE-7 ACCEPTANCE PAIR, and it is a SEPARATE pair on purpose. The
    # Phase-6 files above are the projection and corpus Run 6 was accepted
    # against; widening their schema to carry Phase-7 geometry would change an
    # artefact whose identity is historical evidence. The Windows acceptance
    # harness reads both pairs.
    # BOTH STRUCTURAL LIMITS. The year-column limit fixes W3's duration; the
    # calendar-year window fixes where a 200-year span may sit, and W3's
    # original 2026 start put its last year beyond it.
    phase7 = emit_phase7_acceptance(
        sim, calc, structure.limits.max_generated_year_columns,
        structure.limits.min_year, structure.limits.max_year, out_path.parent)

    # THE PHASE-8 RESULTS PROJECTION, on the terms Phase 7's pair was added:
    # its own file, because the earlier ones are the artefacts earlier runs were
    # accepted against.
    phase8 = emit_phase8_results(
        spec, calc, sim.raw, structure.limits.max_generated_year_columns,
        out_path.parent)

    # P8-2. THE DASHBOARD PROJECTION, and its own file for the same reason: the
    # Results projection above is what the accepted P8-1 Windows run was
    # produced against. It carries the Results-cell each Dashboard cell mirrors,
    # which is the only claim this step makes.
    phase8_dashboard = emit_phase8_dashboard(spec, out_path.parent)

    # P8-3. THE CHART PROJECTION, and its own file for the same reason: the two
    # above are what the accepted P8-1 and P8-2 Windows runs were produced
    # against. It carries what each chart plots and the authority behind it.
    phase8_charts = emit_phase8_charts(
        spec, structure.limits.max_generated_year_columns, out_path.parent, sim)

    # PHASE 9. THE MODEL CHECK PROJECTION, and its own file for the reason every
    # projection above has its own: the earlier ones are the artefacts the
    # accepted Windows runs were produced against. It carries the addresses, the
    # vocabularies, the overflow wording and the advisory threshold - the
    # threshold being the INPUT CONTRACT's, carried rather than restated, so a
    # runner checks the advisory against the number the sheet was built from.
    phase9 = emit_phase9_model_check(spec, contract, out_path.parent)

    say(f"  built    : {out_path}")
    say(f"  emitted  : {artifacts.module_path}")
    say(f"  emitted  : {artifacts.manifest_path}")
    say(f"  emitted  : {artifacts.scenario_path}")
    say(f"  emitted  : {calc_artifacts.module_path}")
    say(f"  emitted  : {calc_artifacts.cases_path}")
    say(f"  emitted  : {sim_artifacts.module_path}")
    say(f"  emitted  : {sim_artifacts.cases_path}")
    say(f"  emitted  : {inspection.path}")
    say(f"  emitted  : {sim_gate_b.inspection_path}")
    say(f"  emitted  : {sim_gate_b.cases_path}")
    say(f"  emitted  : {sim_gate_b.oracle_path}  (host-local oracle evidence)")
    say(f"  emitted  : {phase7[0]}")
    say(f"  emitted  : {phase7[1]}")
    say(f"  emitted  : {phase8}")
    say(f"  emitted  : {phase8_dashboard}")
    say(f"  emitted  : {phase8_charts}")
    say(f"  emitted  : {phase9}")
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
