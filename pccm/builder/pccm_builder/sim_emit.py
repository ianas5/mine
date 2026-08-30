#!/usr/bin/env python3
"""Emit the Phase-6 Stage-A generated artefacts.

Two files, written beside the Stage-A workbook:

  build/vba/modSimContract.bas   the simulation literals a later VBA
                                 implementation will need
  build/phase6_cases.json        the conformance corpus those steps assert against

NEITHER IS EXECUTED HERE, and neither is embedded in the Stage-A `.xlsx`. The
workbook gains no simulation content, no `_SimData` row and no Phase-6
publication; `modSimContract.bas` is an external file, and Stage B imports the
modules its manifest names rather than everything in the directory, so an
artefact no Phase-6 owner has declared yet is simply not imported.

--------------------------------------------------------------------------------
A PROJECTION, NOT AN IMPLEMENTATION
--------------------------------------------------------------------------------
`modSimContract.bas` declares `Public Const` and nothing else. No `Sub`, no
`Function`, no `Property`, no recurrence, no jump arithmetic, no sampler, no
quantile, no digest recurrence, no worksheet access. Its whole job is to stop a
later VBA module from hardcoding a number that already has an owner.

So this module states nothing of its own. Every constant is READ from the
authority that owns it:

    RNG, seeding, components, jumps, Cheng literals,
    digest framing, labels, geometry, ceilings        ->  spec/sim_contract.yaml
    FIXED seed domain, iteration minimum,
    selectable confidence levels                      ->  spec/input_contract.yaml
    model version                                     ->  spec/workbook.yaml

The model version is a CONSTANT, not only a banner comment. Step 1 requires a
successful run to snapshot `model_version` in its metadata, and a comment cannot
be read by the code that has to write it - leaving it in prose would invite the
later VBA to declare a literal of its own, which is the single thing this module
exists to prevent.

--------------------------------------------------------------------------------
TWO THINGS THE PROJECTION DELIBERATELY OMITS
--------------------------------------------------------------------------------
D6-11 is still staged and Step 5 is not the step that activates it, so the
generated module must pass the CURRENT forbidden-construct guard with no entry
scoped and no change to `structure_contract.yaml`.

That means the RNG family name and the future command endpoint name are not
projected at all - not as identifiers, not as string values, not as commentary.
Neutral names (`SIM_RNG_*`) carry the same information without asserting the
forbidden token, and the accepted Quantile naming discipline is kept so the
globally forbidden quantile token never appears either.
"""

from __future__ import annotations

import hashlib
import json
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .contract_loader import InputContract
from .sim_cases import (
    GATE_B_CASES_FILENAME,
    GATE_B_ORACLE_FILENAME,
    build_gate_b_oracle_measurements,
    build_gate_b_pair,
    build_sim_cases,
    render_gate_b_cases_json,
)
from .sim_inspection import emit_sim_inspection
from .sim_loader import SimContract
from .sim_oracle import ITERATIONS_INPUT_KEY, business_minimum_iterations
from .spec_loader import WorkbookSpec

SIM_MODULE_NAME = "modSimContract"

GENERATED_BANNER = (
    "GENERATED FILE - DO NOT EDIT.\n"
    "Emitted from spec/sim_contract.yaml, spec/input_contract.yaml and\n"
    "spec/workbook.yaml by the Stage-A builder. Edit the authority and rebuild;\n"
    "edits made here are overwritten. This module declares CONSTANTS ONLY - no\n"
    "generator, no jump arithmetic, no sampler, no simulation, no statistic, no\n"
    "digest recurrence and no worksheet access."
)

VBA_LONG_MIN = -2147483648
VBA_LONG_MAX = 2147483647


@dataclass(frozen=True)
class SimArtifacts:
    module_path: Path
    cases_path: Path


def emit_sim_artifacts(
    build_dir: Path,
    spec: WorkbookSpec,
    sim: SimContract,
    inputs: InputContract,
    calc: Any,
) -> SimArtifacts:
    """Write modSimContract.bas and phase6_cases.json. Returns their paths."""
    build_dir = Path(build_dir)
    vba_dir = build_dir / "vba"
    vba_dir.mkdir(parents=True, exist_ok=True)

    module_path = vba_dir / f"{SIM_MODULE_NAME}.bas"
    module_path.write_text(
        render_sim_contract_module(spec, sim, inputs), encoding="utf-8"
    )

    cases_path = build_dir / "phase6_cases.json"
    cases_path.write_text(render_sim_cases_json(spec, sim, inputs, calc), encoding="utf-8")
    return SimArtifacts(module_path=module_path, cases_path=cases_path)


@dataclass(frozen=True)
class SimGateBArtifacts:
    """The two Phase-6 Gate-B evidence artefacts, emitted together.

    Deliberately separate from `SimArtifacts`: those two are consumed by the
    Linux conformance suite, these two exist only so a Windows harness can find
    cells and compare against the oracle. Neither carries workbook content.
    """

    inspection_path: Path
    cases_path: Path
    oracle_path: Path


def emit_sim_gate_b_artifacts(
    build_dir: Path,
    spec: WorkbookSpec,
    sim: SimContract,
    inputs: InputContract,
    calc: Any,
) -> SimGateBArtifacts:
    """Write the inspection projection and the two Gate-B parity artefacts.

    THE TWO ARE NOT THE SAME KIND OF THING, and D2 proved it. The case authority
    is cross-platform invariant and is frozen by hash. The oracle evidence is the
    floating output of THIS host's libm and is deliberately not: it carries its
    own provenance and is bound to the authority it was generated against.

    The authority is written first so the evidence can name its SHA-256, which is
    what lets the pre-Excel preflight refuse a pair that did not come from one
    build.
    """
    build_dir = Path(build_dir)
    build_dir.mkdir(parents=True, exist_ok=True)

    inspection = emit_sim_inspection(build_dir, sim, inputs)

    portable, measurements = build_gate_b_pair(
        sim, inputs, calc, spec.model["model_version"]
    )
    cases_path = build_dir / GATE_B_CASES_FILENAME
    cases_text = json.dumps(portable, indent=2, sort_keys=False) + "\n"
    cases_path.write_text(cases_text, encoding="utf-8")

    oracle_path = build_dir / GATE_B_ORACLE_FILENAME
    oracle_path.write_text(
        json.dumps(
            build_gate_b_oracle_measurements(
                portable,
                measurements,
                hashlib.sha256(cases_text.encode("utf-8")).hexdigest(),
                _source_revision(),
            ),
            indent=2,
            sort_keys=False,
        ) + "\n",
        encoding="utf-8",
    )
    return SimGateBArtifacts(
        inspection_path=inspection.path,
        cases_path=cases_path,
        oracle_path=oracle_path,
    )


def _source_revision() -> str:
    """The revision this evidence was generated from, or a plain refusal.

    Recorded, never guessed: evidence with no attributable source revision is
    weaker evidence, and writing "unknown" while pretending otherwise would hand
    that weakness on as though it were strength.
    """
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=str(Path(__file__).resolve().parent.parent.parent),
            capture_output=True, text=True, timeout=15, check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return "unavailable"
    revision = result.stdout.strip()
    return revision if result.returncode == 0 and revision else "unavailable"


# ---------------------------------------------------------------------------
# literal rendering
# ---------------------------------------------------------------------------
def vba_double(value: float) -> str:
    """A Double literal VBA parses back to the same bits.

    `repr` gives the shortest decimal that round-trips, and VBA's parser is
    correctly rounded, so the bits survive. An integral value gets an explicit
    `.0` so VBA cannot type it as a Long. The same discipline Phase 5 proved, and
    it is locale-free: `repr` never emits a thousands separator and always uses a
    point.
    """
    text = repr(float(value))
    if "e" in text or "E" in text or "." in text:
        return text
    return text + ".0"


def _identifier(text: str) -> str:
    out = []
    for char in str(text):
        out.append(char.upper() if (char.isalnum() or char == "_") else "_")
    return "".join(out)


class _Module:
    """The generated text, accumulated. Nothing here evaluates anything."""

    def __init__(self) -> None:
        self.lines: list[str] = []

    def raw(self, text: str = "") -> None:
        self.lines.append(text)

    def comment(self, text: str) -> None:
        self.lines.append(f"' {text}" if text else "'")

    def section(self, title: str) -> None:
        self.lines.append("' " + "-" * 72)
        self.lines.append(f"' {title}")
        self.lines.append("' " + "-" * 72)

    def const(self, name: str, value: Any, comment: str | None = None,
              force_double: bool = False) -> None:
        if isinstance(value, bool):
            raise TypeError(f"{name}: booleans are not projected as constants")
        if isinstance(value, str):
            kind, rendered = "String", '"' + value.replace('"', '""') + '"'
        elif isinstance(value, int):
            if force_double or not VBA_LONG_MIN <= value <= VBA_LONG_MAX:
                # A Long cannot hold it. `m1 = 4294967087` and several jump
                # matrix elements are the reason this branch exists: projected as
                # Long they would not compile, and silently truncating one would
                # change the generator.
                kind, rendered = "Double", vba_double(float(value))
                if float(rendered) != float(value):  # pragma: no cover - < 2**53
                    raise ValueError(f"{name}: {value} does not survive as a Double")
            else:
                kind, rendered = "Long", str(value)
        elif isinstance(value, float):
            kind, rendered = "Double", vba_double(value)
        else:
            raise TypeError(f"{name}: cannot project {value!r}")
        text = f"Public Const {name} As {kind} = {rendered}"
        if comment:
            text += f"    ' {comment}"
        self.lines.append(text)

    def text(self) -> str:
        return "\n".join(self.lines) + "\n"


# ---------------------------------------------------------------------------
# modSimContract.bas
# ---------------------------------------------------------------------------
def render_sim_contract_module(
    spec: WorkbookSpec, sim: SimContract, inputs: InputContract
) -> str:
    """The generated VBA constants module, as text.

    Deterministic: the only inputs are the three authorities. No timestamp, no
    path, no environment.
    """
    raw = sim.raw
    module = _Module()
    module.raw(f'Attribute VB_Name = "{SIM_MODULE_NAME}"')
    module.raw("Option Explicit")
    module.raw()
    for line in GENERATED_BANNER.splitlines():
        module.comment(line)
    module.comment("")
    module.comment(f"Model version      : {spec.model['model_version']}")
    module.comment(f"Simulation contract: {sim.version}")
    module.raw()

    # --- versions -----------------------------------------------------------
    module.section("Versions")
    module.const("SIM_MODEL_VERSION", str(spec.model["model_version"]),
                 "workbook.yaml: model.model_version - snapshotted by a successful run")
    module.const("SIM_CONTRACT_VERSION", sim.version)
    module.const("SIM_RNG_VERSION", sim.rng_version,
                 "bumped by generator, seeding or stream-assignment changes")
    module.const("SIM_METHOD_VERSION", sim.sim_method_version,
                 "bumped by sampling, accumulation, statistical or digest changes")
    module.raw()

    # --- the generator ------------------------------------------------------
    constants = raw["rng"]["constants"]
    state = raw["rng"]["state"]
    module.section("Generator constants (the family name is deliberately not projected)")
    module.const("SIM_RNG_M1", int(constants["m1"]), "exceeds Long; projected as Double")
    module.const("SIM_RNG_M2", int(constants["m2"]), "exceeds Long; projected as Double")
    module.const("SIM_RNG_A12", int(constants["a12"]))
    module.const("SIM_RNG_A13N", int(constants["a13n"]))
    module.const("SIM_RNG_A21", int(constants["a21"]))
    module.const("SIM_RNG_A23N", int(constants["a23n"]))
    module.const("SIM_RNG_NORM", float(constants["norm"]))
    module.const("SIM_RNG_STATE_WORDS", int(state["words"]))
    for ordinal, word in enumerate(state["order"], start=1):
        module.const(f"SIM_RNG_STATE_{ordinal}", str(word))
    module.const("SIM_RNG_STATE_ORIENTATION", str(state["orientation"]))
    module.const("SIM_RNG_MATRIX_OPERAND_ORIENTATION",
                 str(state["matrix_operand_orientation"]),
                 "reverse each triple at the matrix boundary")
    module.const("SIM_RNG_U_LOWER_EXCLUSIVE", int(raw["rng"]["output_domain"]["lower"]))
    module.const("SIM_RNG_U_UPPER_EXCLUSIVE", int(raw["rng"]["output_domain"]["upper"]))
    module.raw()

    # --- seeding ------------------------------------------------------------
    auto = raw["seeding"]["auto"]
    lifecycle = raw["seeding"]["nonce_lifecycle"]
    module.section("Seeding")
    module.const("SIM_SEED_MIN", _seed_bound(inputs, "min"),
                 "owned by input_contract.yaml")
    module.const("SIM_SEED_MAX", _seed_bound(inputs, "max"),
                 "owned by input_contract.yaml")
    module.const("SIM_AUTO_MODULUS", int(auto["modulus"]))
    module.const("SIM_AUTO_MULTIPLIER", int(auto["multiplier"]))
    module.const("SIM_AUTO_PERIOD", int(auto["period"]))
    module.const("SIM_NONCE_INITIAL", int(lifecycle["initial"]))
    module.const("SIM_NONCE_FIRST_VALID", int(lifecycle["first_valid_allocation"]))
    module.const("SIM_NONCE_LAST_VALID", int(lifecycle["last_valid_allocation"]))
    module.const("SIM_NONCE_EXHAUSTED", int(lifecycle["exhausted_value"]))
    for ordinal, mode in enumerate(raw["label_sets"]["seed_mode"], start=1):
        module.const(f"SIM_SEED_MODE_{_identifier(mode)}", str(mode))
    module.raw()

    # --- components and streams ---------------------------------------------
    module.section("Components and stream assignment")
    for ordinal, entry in enumerate(raw["components"]["kinds"], start=1):
        prefix = f"SIM_COMPONENT_{ordinal}"
        module.const(f"{prefix}_KEY", str(entry["key"]))
        module.const(f"{prefix}_DRIVER_KIND", str(entry["driver_kind"]))
        module.const(f"{prefix}_ROLE", str(entry["role"]))
        module.const(f"{prefix}_PER_DRIVER", int(entry["per_driver"]))
    module.const("SIM_COMPONENT_KIND_COUNT", len(raw["components"]["kinds"]))
    module.const("SIM_STREAM_INDEX_ORIGIN", int(raw["stream_assignment"]["index_origin"]))
    module.const("SIM_STREAM_ID_COMPARISON",
                 str(raw["stream_assignment"]["permanent_id_comparison"]))
    for ordinal, kind in enumerate(raw["accumulation"]["driver_kind_order"], start=1):
        module.const(f"SIM_ACCUMULATION_KIND_{ordinal}", str(kind))
    module.const("SIM_ACCUMULATION_WITHIN_KIND",
                 str(raw["accumulation"]["within_kind_order"]))
    module.raw()

    # --- jump ---------------------------------------------------------------
    jump = raw["jump"]
    module.section("Stream jump")
    module.const("SIM_STREAM_SPACING_EXPONENT", int(jump["stream_spacing_exponent"]))
    module.const("SIM_JUMP_DECOMPOSITION_H", int(jump["decomposition_h"]),
                 "the VBA-safe modular multiply split")
    module.comment("Every element is projected as Double: several exceed Long, and a")
    module.comment("matrix whose elements had two different VBA types would be a trap.")
    for name, matrix in (("A1", jump["a1_p127"]), ("A2", jump["a2_p127"])):
        for row_index, row in enumerate(matrix, start=1):
            for column_index, element in enumerate(row, start=1):
                module.const(
                    f"SIM_JUMP_{name}_R{row_index}C{column_index}",
                    int(element),
                    force_double=True,
                )
    module.raw()

    # --- distributions ------------------------------------------------------
    distributions = raw["distributions"]
    module.section("Distribution families")
    for ordinal, family in enumerate(distributions["families"], start=1):
        module.const(f"SIM_FAMILY_{ordinal}", str(family))
    module.const("SIM_FAMILY_COUNT", len(distributions["families"]))
    module.const("SIM_PERT_LAMBDA", int(distributions["beta_pert"]["lambda"]))
    module.const("SIM_PERT_SHAPE_LOWER", int(distributions["beta_pert"]["shape_lower"]))
    module.const("SIM_PERT_SHAPE_UPPER", int(distributions["beta_pert"]["shape_upper"]))
    module.const("SIM_PERT_ALPHA_PLUS_BETA",
                 int(distributions["beta_pert"]["alpha_plus_beta"]))
    module.const("SIM_UNIFORM_USES_MOST_LIKELY", 0,
                 "0 = a Uniform ignores Most Likely entirely (D1)")
    module.const("SIM_DEGENERATE_UNIFORMS_CONSUMED",
                 int(distributions["degenerate"]["uniforms_consumed"]))
    module.raw()

    # --- Cheng --------------------------------------------------------------
    cheng = raw["cheng"]
    module.section("Locked acceptance/rejection literals - literals, never evaluated")
    module.const("SIM_CHENG_UNIFORMS_PER_ATTEMPT",
                 int(cheng["uniforms_per_non_degenerate_proposal_attempt"]))
    for ordinal, literal in enumerate(cheng["bb"]["literals"], start=1):
        module.const(f"SIM_CHENG_BB_LITERAL_{ordinal}", float(literal))
    for ordinal, literal in enumerate(cheng["bc"]["literals"], start=1):
        module.const(f"SIM_CHENG_BC_LITERAL_{ordinal}", float(literal))
    module.const("SIM_CHENG_BB_APPLIES_WHEN", str(cheng["bb"]["applies_when"]))
    module.const("SIM_CHENG_BC_APPLIES_WHEN", str(cheng["bc"]["applies_when"]))
    module.const("SIM_CHENG_ACCEPT_OPERATOR", str(cheng["bb"]["acceptance_operator"]))
    module.raw()

    # --- risk and contribution ----------------------------------------------
    module.section("Risk and contribution rules")
    module.const("SIM_RISK_OCCURRENCE_UNIFORMS",
                 int(raw["risk"]["occurrence"]["uniforms_per_risk_per_iteration"]))
    module.const("SIM_RISK_OCCURRENCE_OPERATOR",
                 str(raw["risk"]["occurrence"]["comparison_operator"]))
    module.const("SIM_RISK_SEVERITY_INVOCATION",
                 str(raw["risk"]["severity"]["invocation_policy"]))
    module.const("SIM_RISK_SEVERITY_DEGENERATE_UNIFORMS",
                 int(raw["risk"]["severity"]["degenerate_consumption"]))
    module.const("SIM_COST_NOMINAL_RULE", str(raw["contribution"]["cost_line"]["nominal"]))
    module.const("SIM_COST_PV_RULE", str(raw["contribution"]["cost_line"]["pv"]))
    module.const("SIM_COST_QUANTITY_APPLICATIONS",
                 int(raw["contribution"]["cost_line"]["quantity_applications"]))
    module.const("SIM_RISK_NOMINAL_RULE",
                 str(raw["contribution"]["risk"]["nominal_when_occurred"]))
    module.const("SIM_RISK_PV_RULE", str(raw["contribution"]["risk"]["pv_when_occurred"]))
    module.raw()

    # --- publication banks --------------------------------------------------
    identity = raw["sim_data"]["run_identity"]
    banks = raw["publication"]["banks"]
    module.section("Publication banks (the second bank consumes COLUMNS, not rows)")
    for label in banks["labels"]:
        module.const(f"SIM_BANK_{_identifier(label)}", str(label))
    module.const("SIM_BANK_COUNT", int(banks["count"]))
    module.const("SIM_ACTIVE_BANK_ROW", int(
        next(f["row"] for f in identity["fields"] if f["key"] == "active_bank")))
    module.const("SIM_SHARED_VALUE_COLUMN", str(identity["value_column"]))
    for label, column in identity["bank_value_columns"].items():
        module.const(f"SIM_SNAPSHOT_COLUMN_{_identifier(label)}", str(column))
    records = raw["sim_data"]["iteration_records"]
    for label, columns in records["banks"].items():
        for key, column in columns.items():
            module.const(
                f"SIM_ITER_{_identifier(label)}_{_identifier(key)}_COLUMN", str(column))
    transaction = raw["publication"]["transaction"]
    module.const("SIM_FINAL_COMMIT_RANGE", str(transaction["final_commit_range"]))
    # The durable write-ahead recovery marker. Projected as ONE constant so no
    # production procedure spells the coordinate for itself.
    pending = raw["sim_data"]["pending_auto_nonce"]
    module.const("SIM_PENDING_AUTO_NONCE_CELL", str(pending["cell"]))
    module.raw()

    # --- persisted summary and contingency ----------------------------------
    summary = raw["sim_data"]["summary_statistics"]
    module.section("Persisted summary statistics (computed by modSimStats, never by a formula)")
    module.const("SIM_SUMMARY_LABEL_COLUMN", str(summary["label_column"]))
    module.const("SIM_SUMMARY_FIRST_ROW", int(summary["first_row"]))
    module.const("SIM_SUMMARY_LAST_ROW", int(summary["last_row"]))
    for label, columns in summary["bank_value_columns"].items():
        for measure, column in columns.items():
            module.const(
                f"SIM_SUMMARY_{_identifier(label)}_{_identifier(measure)}_COLUMN",
                str(column))
    for metric in summary["metrics"]:
        module.const(f"SIM_SUMMARY_ROW_{_identifier(metric['key'])}", int(metric["row"]))
    module.raw()

    contingency = raw["sim_data"]["contingency_ladder"]
    module.section("Persisted contingency ladder (the WHOLE ladder, before any commit)")
    module.const("SIM_CONTINGENCY_LABEL_COLUMN", str(contingency["label_column"]))
    module.const("SIM_CONTINGENCY_FIRST_ROW", int(contingency["first_row"]))
    module.const("SIM_CONTINGENCY_LAST_ROW", int(contingency["last_row"]))
    for label, columns in contingency["bank_value_columns"].items():
        for measure, column in columns.items():
            module.const(
                f"SIM_CONTINGENCY_{_identifier(label)}_{_identifier(measure)}_COLUMN",
                str(column))
    for rung in contingency["rungs"]:
        module.const(f"SIM_CONTINGENCY_ROW_{_identifier(rung['key'])}", int(rung["row"]))
    module.raw()

    # --- the Phase-5 bridge and the settled public surface ------------------
    bridge = raw["phase5_bridge"]
    module.section("The one Phase-5 preparation bridge, and the settled read accessors")
    module.const("SIM_PREPARE_BRIDGE", str(bridge["procedure"]))
    module.const("SIM_PREPARE_BRIDGE_OWNER", str(bridge["owner_module"]))
    module.const("SIM_PREPARE_REQUIRES_STATUS", str(bridge["requires_phase5_status"]))
    surface = raw["command_surface"]
    for ordinal, name in enumerate(surface["read_accessors"], start=1):
        module.const(f"SIM_READ_ACCESSOR_{ordinal}", str(name))
    module.const("SIM_READ_ACCESSOR_COUNT", len(surface["read_accessors"]))
    module.raw()

    # --- request fingerprint ------------------------------------------------
    # The SIM extension's FRAMING only. The stream tag, FP_VERSION and the hash
    # mathematics belong to modCalcContract and are deliberately not repeated
    # here: the extension is a section of the accepted PCCM-FP stream, not a
    # stream of its own.
    request = raw["request_fingerprint"]["sim_section"]
    module.section("Request fingerprint: the SIM extension framing")
    module.const("SIM_REQUEST_SECTION", str(request["name"]))
    module.const("SIM_REQUEST_RECORD_COUNT", int(request["record_count"]))
    for mode, shape in request["effective_records"].items():
        module.const(
            f"SIM_REQUEST_FIELD_COUNT_{_identifier(mode)}", int(shape["field_count"])
        )
    for mode, shape in request["effective_records"].items():
        for ordinal, field in enumerate(shape["fields"], start=1):
            module.const(f"SIM_REQUEST_{_identifier(mode)}_FIELD_{ordinal}", str(field))
    for field in request["fields"]:
        module.const(
            f"SIM_REQUEST_TYPE_{_identifier(field)}", str(request["field_types"][field])
        )
    module.const(
        "SIM_REQUEST_AUTO_SEED", str(request["auto_supplied_seed_representation"]),
        "AUTO has no supplied seed field at all - not zero, not blank",
    )
    module.raw()

    # --- result digest ------------------------------------------------------
    digest = raw["result_digest"]
    module.section("Result digest framing (the hash mathematics is modCalcContract's)")
    module.const("SIM_DIGEST_STREAM_TAG", str(digest["stream_tag"]))
    module.const("SIM_DIGEST_SECTION", str(digest["section_name"]))
    module.const("SIM_DIGEST_FIELD_COUNT", int(digest["record_field_count"]))
    module.const("SIM_DIGEST_INDEX_ORIGIN", int(digest["iteration_index_origin"]))
    for ordinal, field in enumerate(digest["record_fields"], start=1):
        module.const(f"SIM_DIGEST_FIELD_{ordinal}", str(field))
    for ordinal, kind in enumerate(digest["field_types"], start=1):
        module.const(f"SIM_DIGEST_FIELD_TYPE_{ordinal}", str(kind))
    module.raw()

    # --- state and attempt vocabulary ---------------------------------------
    module.section("Two orthogonal state axes")
    for label in raw["label_sets"]["sim_state"]:
        module.const(f"SIM_STATE_{_identifier(label)}", str(label))
    for label in raw["label_sets"]["attempt_result"]:
        module.const(f"SIM_ATTEMPT_{_identifier(label)}", str(label))
    module.raw()

    # --- geometry and ceilings ----------------------------------------------
    layout = sim.layout
    module.section("_SimData geometry and the ceilings it determines")
    module.const("SIM_DATA_SHEET", layout.sheet)
    module.const("SIM_DATA_VISIBILITY", layout.required_visibility)
    module.const("SIM_DATA_HEADER_ROW", layout.header_row)
    module.const("SIM_DATA_FIRST_ITERATION_ROW", layout.first_iteration_row)
    module.const("SIM_DATA_RESERVED_ROWS", layout.reserved_row_count)
    module.const("SIM_DATA_FOOTER_ROWS", layout.footer_rows)
    module.const("SIM_MAX_ITERATIONS", layout.max_iterations_representable,
                 "technical ceiling; not a business rule")
    module.const("SIM_MIN_ITERATIONS", business_minimum_iterations(inputs),
                 f"business minimum, owned by input_contract.yaml ({ITERATIONS_INPUT_KEY})")
    identity = raw["sim_data"]["run_identity"]
    module.const("SIM_IDENTITY_LABEL_COLUMN", str(identity["label_column"]))
    module.const("SIM_IDENTITY_VALUE_COLUMN", str(identity["value_column"]))
    module.const("SIM_IDENTITY_FIRST_ROW", int(identity["first_row"]))
    module.const("SIM_IDENTITY_LAST_ROW", int(identity["last_row"]))
    for field in identity["fields"]:
        module.const(f"SIM_IDENTITY_ROW_{_identifier(field['key'])}", int(field["row"]))
    module.raw()

    # --- run id -------------------------------------------------------------
    run_id = raw["run_id"]
    module.section("Run identity counter")
    module.const("SIM_RUN_ID_INITIAL", int(run_id["initial"]))
    module.const("SIM_RUN_ID_FIRST", int(run_id["first_successful_value"]))
    module.const("SIM_RUN_ID_MAXIMUM", int(run_id["maximum"]))
    module.raw()

    # --- the reported ladder ------------------------------------------------
    statistics = raw["statistics"]
    module.section("Reported quantile ladder (Quantile naming discipline retained)")
    module.const("SIM_STAT_MEAN_METHOD", str(statistics["mean"]["method"]))
    module.const("SIM_STAT_SD_METHOD", str(statistics["standard_deviation"]["method"]))
    module.const("SIM_STAT_SD_DIVISOR", str(statistics["standard_deviation"]["divisor"]))
    module.const("SIM_QUANTILE_METHOD", str(statistics["percentile"]["method"]))
    ladder = _ladder_labels(sim, inputs)
    module.const("SIM_QUANTILE_COUNT", len(ladder))
    for ordinal, label in enumerate(ladder, start=1):
        module.const(f"SIM_QUANTILE_{ordinal}", label)
    for ordinal, label in enumerate(statistics["headline_percentiles"], start=1):
        module.const(f"SIM_QUANTILE_HEADLINE_{ordinal}", str(label))
    for ordinal, label in enumerate(statistics["fixed_nonselectable_percentiles"], start=1):
        module.const(f"SIM_QUANTILE_FIXED_{ordinal}", str(label))
    module.const("SIM_CONTINGENCY_FORMULA", str(raw["contingency"]["formula"]))
    module.const("SIM_CONTINGENCY_BASELINE", str(raw["contingency"]["baseline"]))
    module.raw()

    return module.text()


def _seed_bound(inputs: InputContract, which: str) -> int:
    """The FIXED seed domain, read from `input_contract.yaml` and nowhere else."""
    from .sim_rng import _seed_domain

    minimum, maximum = _seed_domain(inputs)
    return minimum if which == "min" else maximum


def _ladder_labels(sim: SimContract, inputs: InputContract) -> list[str]:
    from .sim_oracle import resolve_percentile_ladder

    return list(resolve_percentile_ladder(sim, inputs).ordered)


# ---------------------------------------------------------------------------
# phase6_cases.json
# ---------------------------------------------------------------------------
def render_sim_cases_json(
    spec: WorkbookSpec, sim: SimContract, inputs: InputContract, calc: Any
) -> str:
    """The conformance corpus, as deterministic JSON text.

    `sort_keys=False` keeps the authored order of each record, which is stable
    because every structure is built in a fixed sequence. `allow_nan=False` makes
    a non-finite value a BUILD FAILURE rather than the non-standard `NaN` and
    `Infinity` tokens no other JSON reader accepts.
    """
    document = build_sim_cases(sim, inputs, calc, spec.model["model_version"])
    return json.dumps(
        document, indent=2, sort_keys=False, allow_nan=False, ensure_ascii=True
    ) + "\n"
