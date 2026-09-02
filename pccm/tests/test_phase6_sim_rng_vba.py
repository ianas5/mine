#!/usr/bin/env python3
"""PCCM Phase 6 Step-6 conformance tests for `src/vba/modSimRng.bas`.

The first Phase-6 source VBA: state validation, FIXED seeding, the AUTO nonce
mapping, one recurrence step, the canonical 2^127 jump and canonical
component-stream assignment.

--------------------------------------------------------------------------------
WHAT THESE TESTS PROVE, AND WHAT THEY DO NOT
--------------------------------------------------------------------------------
SOURCE CONFORMANCE, on Linux, now. The module is read as text: its purity, its
use of the projected constants, the shape of every locked formula, and the
arithmetic those formulas describe.

VBA EXECUTION CONFORMANCE is NOT proved here and is deferred to Gate B on
Windows. No VBA runtime exists in this step. Where a test evaluates the
arithmetic, it evaluates a TRANSCRIPTION of the expressions read out of the
module - which is evidence about the algorithm, not about the interpreter. No
test in this file may be read as "VBA produced this number".

The retained Step-0 evidence and `build/phase6_cases.json` groups A_rng, B_jump
and B_seed remain the execution-vector authority for that later gate.

Runs standalone or under pytest.
"""

from __future__ import annotations

import hashlib
import json
import math
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import load_contract, load_sim_contract, load_structure_contract  # noqa: E402
from pccm_builder.sim_emit import render_sim_contract_module  # noqa: E402
from pccm_builder.spec_loader import load_spec  # noqa: E402
from pccm_builder.vba_source import (  # noqa: E402
    VbaModule,
    contains_construct,
    load_modules,
    logical_statements,
    strip_comments,
    strip_strings,
)

SRC_VBA = PCCM_ROOT / "src" / "vba"
SIM_RNG_BAS = SRC_VBA / "modSimRng.bas"
SPEC = PCCM_ROOT / "spec"
EVIDENCE = PCCM_ROOT / "evidence" / "phase6_step0" / "vectors"

_CACHE: dict[str, object] = {}


def _module() -> VbaModule:
    return VbaModule(
        name="modSimRng", path=SIM_RNG_BAS,
        raw=SIM_RNG_BAS.read_text(encoding="utf-8"),
    )


def _code() -> str:
    """The module with comments and string literals removed."""
    return _module().code


def _procedure(name: str) -> str:
    """One procedure body, from its declaration to its End."""
    code = _module().code_without_string_removal
    pattern = re.compile(
        rf"^\s*(?:Public|Private)\s+(?:Function|Sub)\s+{re.escape(name)}\b", re.M
    )
    match = pattern.search(code)
    assert match, f"{name} is not declared"
    tail = code[match.start():]
    end = re.search(r"^\s*End\s+(?:Function|Sub)\s*$", tail, re.M)
    assert end, f"{name} has no End"
    return tail[: end.end()]


def _sim_contract():
    if "sim" not in _CACHE:
        _CACHE["sim"] = load_sim_contract(SPEC / "sim_contract.yaml")
    return _CACHE["sim"]


def _generated_constants() -> dict[str, tuple[str, str]]:
    """`name -> (type, literal)` from the generated projection, rendered fresh."""
    if "consts" not in _CACHE:
        text = render_sim_contract_module(
            load_spec(SPEC / "workbook.yaml"),
            _sim_contract(),
            load_contract(SPEC / "input_contract.yaml"),
        )
        out: dict[str, tuple[str, str]] = {}
        for line in text.splitlines():
            match = re.match(r"^Public Const (\w+) As (\w+) = (.*)$", line)
            if not match:
                continue
            name, kind, rest = match.groups()
            literal = rest.split("    '")[0].rstrip()
            out[name] = (kind, literal)
        _CACHE["consts"] = out
    return _CACHE["consts"]  # type: ignore[return-value]


def _const(name: str):
    kind, literal = _generated_constants()[name]
    if kind == "String":
        return literal[1:-1].replace('""', '"')
    return float(literal) if kind == "Double" else int(literal)


def _evidence(name: str) -> dict:
    return json.loads((EVIDENCE / f"{name}.json").read_text(encoding="utf-8"))


# ===========================================================================
# A. the module exists and is declared
# ===========================================================================
def test_01_the_module_exists_and_opens_correctly() -> None:
    lines = SIM_RNG_BAS.read_text(encoding="utf-8").splitlines()
    assert lines[0] == 'Attribute VB_Name = "modSimRng"'
    assert lines[1] == "Option Explicit"


def test_02_the_module_is_declared_hand_written_in_the_structure_contract() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    modules = {m.name: m for m in structure.vba_modules}
    assert "modSimRng" in modules
    assert modules["modSimRng"].generated is False
    assert "modSimContract" in modules and modules["modSimContract"].generated is True


def test_03_the_public_surface_is_small_and_step_7_ready() -> None:
    public = set(_module().public_procedures)
    assert public == {
        "SimRngValidateState",
        "SimRngStateFromFixedSeed",
        "SimRngAutoSeedFromNonce",
        "SimRngNextUniform",
        "SimRngJumpNextStream",
        "SimRngStreamInitialState",
        "SimRngBuildComponentStreams",
    }, sorted(public)
    # The numerical internals stay private: nothing downstream needs them yet.
    private = set(_module().procedures) - public
    for helper in ("SimRngReduce", "SimRngMultModM", "MRG32k3aStep", "SimRngNorm"):
        assert helper in private, helper


def test_04_no_public_procedure_accepts_an_object() -> None:
    """No worksheet, workbook or COM object crosses the module boundary."""
    code = _module().code_without_string_removal
    for statement in logical_statements(strip_comments(code)):
        text = statement[1]
        if not re.match(r"^\s*Public\s+(Function|Sub)\s", text):
            continue
        for banned in ("As Object", "As Worksheet", "As Workbook", "As Range",
                       "As ListObject", "As Variant"):
            assert banned not in text, f"{banned} in: {text}"


# ===========================================================================
# B. purity - the sweep is active from the first commit
# ===========================================================================
FORBIDDEN_IN_MODULE = (
    "Range", "Cells", "Worksheet", "Worksheets", "Workbook", "Workbooks",
    "ListObject", "Application", "ThisWorkbook", "ActiveWorkbook", "Names(",
    "Evaluate", "MsgBox", "Rnd(", "Randomize", "Open ", "Kill ", "Environ",
    "Shell", "CreateObject", "GetObject", "Date", "Now", "Timer", "DoEvents",
)


def test_05_the_module_touches_no_workbook_and_no_environment() -> None:
    code = _code()
    problems = [token for token in FORBIDDEN_IN_MODULE if token in code]
    assert not problems, problems


def test_06_the_module_holds_no_global_or_static_state() -> None:
    """No module-level generator, no Static local, no hidden singleton."""
    code = strip_comments(_module().raw)
    inside = False
    module_level = []
    for line in code.splitlines():
        stripped = line.strip()
        if re.match(r"^(Public|Private)?\s*(Function|Sub)\s", stripped):
            inside = True
        elif re.match(r"^End\s+(Function|Sub)\b", stripped):
            inside = False
        elif not inside and re.match(r"^(Public|Private|Dim|Global)\s+\w+\s+As\s", stripped):
            module_level.append(stripped)
    assert not module_level, module_level
    assert not re.search(r"^\s*Static\s", code, re.M), "a Static local exists"
    # The only module-level declarations are the two public types.
    assert code.count("Public Type ") == 2
    assert "End Type" in code


def test_07_no_sampler_or_engine_code_leaked_in() -> None:
    """Step 6 produces RAW uniforms only."""
    code = _code()
    for token in ("Triangular", "BetaPert", "Beta_", "Cheng", "Bernoulli",
                  "Occurrence", "Severity", "Percentile", "Quantile", "Contingency",
                  "Digest", "Fingerprint", "StandardDeviation", "Contribution",
                  "Iteration", "SimData", "Results"):
        assert token not in code, f"{token} appears in modSimRng"
    # The one exception is the component ROLE strings, which are constants read
    # from the projection rather than sampler code - and they are string
    # payloads, which the code view strips.
    raw = SIM_RNG_BAS.read_text(encoding="utf-8")
    assert "SIM_COMPONENT_2_ROLE" in raw and "SIM_COMPONENT_3_ROLE" in raw


# ===========================================================================
# C. D6-11 - the scoped grant is real
# ===========================================================================
def test_08_the_algorithm_token_is_in_executable_code_here() -> None:
    """A grant nobody exercises proves nothing."""
    assert contains_construct([_module()], "MRG32k3a") == ["modSimRng"]
    assert "MRG32k3aStep" in _module().procedures, (
        "the token must be a real procedure name, not only prose"
    )
    assert "MRG32k3a" in _code(), "the token was stripped as a comment or a literal"


def test_09_the_algorithm_token_is_in_no_other_module() -> None:
    others = [m for m in load_modules([SRC_VBA]) if m.name != "modSimRng"]
    assert others, "no other module was read"
    assert not contains_construct(others, "MRG32k3a")
    # And it is not in the generated projection either: the family name was
    # deliberately left out of modSimContract so the projection needs no grant.
    text = render_sim_contract_module(
        load_spec(SPEC / "workbook.yaml"), _sim_contract(),
        load_contract(SPEC / "input_contract.yaml"),
    )
    assert "MRG32k3a" not in text


def test_10_every_scoped_construct_has_exactly_one_owner() -> None:
    """TWO scoped constructs since Step 11, each to ONE module.

    RunSimulation was global while no module could legitimately contain it, and
    was scoped in the same commit that introduced its owner. What this test
    protects is unchanged: a scoped construct names exactly one owner, and every
    other module is still refused.
    """
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    scoped = [r for r in structure.forbidden_construct_rules if r.is_scoped]
    assert [(r.construct, tuple(r.allowed_in)) for r in scoped] == [
        ("MRG32k3a", ("modSimRng",)),
        ("RunSimulation", ("modSimReport",)),
    ], scoped
    for rule in scoped:
        assert len(rule.allowed_in) == 1, rule.construct
        assert "*" not in rule.allowed_in
    endpoint = next(r for r in structure.forbidden_construct_rules
                    if r.construct == "RunSimulation")
    assert endpoint.allowed_in == ("modSimReport",)
    assert endpoint.forbidden_in("modSimRng") is True
    assert endpoint.forbidden_in("modSimReport") is False


def test_11_the_globally_forbidden_constructs_still_apply_here() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    for rule in structure.forbidden_construct_rules:
        if rule.construct == "MRG32k3a":
            continue
        assert rule.forbidden_in("modSimRng") is True, rule.construct
        assert not contains_construct([_module()], rule.construct), rule.construct


def test_12_the_manifest_carries_the_structured_rules() -> None:
    from pccm_builder.stage_b_emit import build_manifest
    from pccm_builder import load_driver_contract

    manifest = build_manifest(
        load_spec(SPEC / "workbook.yaml"),
        load_contract(SPEC / "input_contract.yaml"),
        load_driver_contract(SPEC / "driver_contract.yaml"),
        load_structure_contract(SPEC / "structure_contract.yaml"),
    )
    rules = manifest["vba"]["forbidden_construct_rules"]
    by_construct = {r["construct"]: r["allowed_in"] for r in rules}
    assert by_construct["MRG32k3a"] == ["modSimRng"]
    # SCOPED SINCE STEP 11, in the commit that introduced its owner.
    assert by_construct["RunSimulation"] == ["modSimReport"]
    assert by_construct["Percentile"] == []
    # EVERY OTHER construct is still global. Exactly two are scoped, and each
    # names exactly one owner.
    for construct, owners in by_construct.items():
        if construct in ("MRG32k3a", "RunSimulation"):
            assert len(owners) == 1, (construct, owners)
            continue
        assert owners == [], (construct, owners)
    # Every flattened entry still has a structured rule, so no consumer is left
    # without an authority.
    assert set(by_construct) == set(manifest["vba"]["forbidden_constructs"])
    names = [m["name"] for m in manifest["vba"]["modules"]]
    # THE PHASE-6 BLOCK, CONTIGUOUS AND IN ORDER - see the note in the sibling
    # batteries. P7-2 landed a module after it; the block itself is unchanged.
    block = ['modSimContract', 'modSimRng', 'modSimSample', 'modSimEngine', 'modSimStats', 'modSimFingerprint', 'modSimNonce', 'modSimReport']
    at = names.index(block[0])
    assert names[at:at + len(block)] == block, names[at:at + len(block)]


# ===========================================================================
# THE TRANSCRIPTION
#
# READ THIS BEFORE READING ANY RESULT BELOW.
#
# `tests/phase6_vba_transcribe.py` compiles the STATEMENTS OF modSimRng.bas into
# Python and runs them. It is a transcription of the source, not a VBA
# interpreter: it proves that the algorithm the module WRITES DOWN reproduces
# the accepted vectors, and it fails the moment a locked expression in the .bas
# is altered, because every expression it evaluates is read out of the file at
# test time.
#
# It proves nothing about how VBA itself would execute those statements. Type
# coercion, the numeric parser, Fix on the VBA side, ByRef binding and overflow
# behaviour are the Windows runtime's business, and Gate B is where they are
# settled against these same accepted vectors.
#
# Every operational value comes from the projected constants, exactly as the
# module reads them, so a wrong constant fails here too.
#
# The engine was extracted mechanically from this file in Step 7 so the sampler
# suite could reuse it rather than grow a second transcription language. The
# names below are re-exported unchanged, and every Step-6 assertion is
# untouched.
# ===========================================================================
from phase6_vba_transcribe import (  # noqa: E402
    _Ref, _assign, _copy, _fix, _val, build as _build_transcription,
)


def _transcribe() -> dict:
    """The compiled module: `name -> callable`, plus the constants it reads."""
    if "vba" not in _CACHE:
        _CACHE["vba"] = _build_transcription(
            {"modSimRng": SIM_RNG_BAS},
            {name: _const(name) for name in _generated_constants()},
        )
    return _CACHE["vba"]  # type: ignore[return-value]


def _mk(*words: float) -> dict:
    """A SimRngState dict in the module's own oldest-first field order."""
    fields = [f for f, _ in _transcribe()["_types"]["SimRngState"]]
    assert fields == ["S10", "S11", "S12", "S20", "S21", "S22"], fields
    return dict(zip(fields, (float(w) for w in words)))


def _words(state: dict) -> list[int]:
    return [int(state[f]) for f, _ in _transcribe()["_types"]["SimRngState"]]


def _seeded(seed: int) -> dict:
    vba = _transcribe()
    state, detail = _mk(0, 0, 0, 0, 0, 0), _Ref("")
    assert vba["SimRngStateFromFixedSeed"](_Ref(seed), state, detail), detail.v
    return state


def _draw(state: dict, count: int) -> list[str]:
    vba = _transcribe()
    out = []
    for _ in range(count):
        u, detail = _Ref(0.0), _Ref("")
        assert vba["SimRngNextUniform"](state, u, detail), detail.v
        out.append(repr(u.v))
    return out


# ===========================================================================
# D. SOURCE ARITHMETIC CONFORMANCE - the transcription against the accepted
#    Step-0 vectors. NOT a VBA execution result; see the banner above.
# ===========================================================================
def test_13_the_constructed_normalisation_is_the_projected_constant() -> None:
    """The 15-digit ceiling: the construction is bound to the projection."""
    vba = _transcribe()
    assert vba["SimRngNorm"]() == _const("SIM_RNG_NORM")
    assert _const("SIM_RNG_NORM") == 1.0 / (_const("SIM_RNG_M1") + 1.0)
    # And the reason the module does not spell it: fifteen significant digits
    # name a DIFFERENT Double, four ulp away.
    spelled = float(f"{_const('SIM_RNG_NORM'):.15g}")
    assert spelled != _const("SIM_RNG_NORM")
    gap = abs(spelled - _const("SIM_RNG_NORM")) / math.ulp(_const("SIM_RNG_NORM"))
    assert gap == 4.0, gap
    # The literal in the projection itself needs sixteen.
    assert float(f"{_const('SIM_RNG_NORM'):.16g}") == _const("SIM_RNG_NORM")
    assert "SIM_RNG_NORM" not in _code(), "the module spells the constant it must construct"


def test_14_fixed_seeding_reproduces_every_accepted_example() -> None:
    for example in _evidence("seed_vectors")["examples"]:
        assert _words(_seeded(example["seed"])) == example["state"], example


def test_15_the_fixed_seed_domain_is_the_projected_one() -> None:
    vba = _transcribe()
    for seed in (_const("SIM_SEED_MIN") - 1, _const("SIM_SEED_MAX") + 1, -1, 0):
        state, detail = _mk(0, 0, 0, 0, 0, 0), _Ref("")
        assert vba["SimRngStateFromFixedSeed"](_Ref(seed), state, detail) is False
        assert "admissible domain" in detail.v
        assert _words(state) == [0] * 6, "a refused seed still wrote the state"
    for seed in (_const("SIM_SEED_MIN"), _const("SIM_SEED_MAX")):
        state, detail = _mk(0, 0, 0, 0, 0, 0), _Ref("")
        assert vba["SimRngStateFromFixedSeed"](_Ref(seed), state, detail), detail.v


def test_16_the_auto_nonce_mapping_reproduces_every_accepted_pair() -> None:
    vba = _transcribe()
    exhausted = _const("SIM_NONCE_EXHAUSTED")
    checked = 0
    for pair in _evidence("seed_vectors")["nonce_to_seed_pairs"]:
        nonce, seed, detail = pair["auto_nonce"], _Ref(0), _Ref("")
        ok = vba["SimRngAutoSeedFromNonce"](_Ref(nonce), seed, detail)
        if nonce >= exhausted:
            # The accepted mathematics wraps here; the module refuses to, and the
            # evidence records exactly the seed nonce 0 already issued.
            assert ok is False and detail.v == "auto nonce: exhausted"
            assert pair["effective_seed"] == \
                _evidence("seed_vectors")["nonce_to_seed_pairs"][0]["effective_seed"]
            continue
        assert ok, detail.v
        assert seed.v == pair["effective_seed"], pair
        checked += 1
    assert checked >= 7, checked


def test_17_the_auto_nonce_domain_is_closed_at_both_ends() -> None:
    vba = _transcribe()
    for nonce, message in ((_const("SIM_NONCE_FIRST_VALID") - 1, "below the first valid"),
                           (_const("SIM_NONCE_EXHAUSTED"), "exhausted"),
                           (_const("SIM_NONCE_EXHAUSTED") + 1, "exhausted")):
        seed, detail = _Ref(0), _Ref("")
        assert vba["SimRngAutoSeedFromNonce"](_Ref(nonce), seed, detail) is False
        assert message in detail.v
        assert seed.v == 0, "a refused nonce still wrote the seed"
    seed, detail = _Ref(0), _Ref("")
    assert vba["SimRngAutoSeedFromNonce"](_Ref(_const("SIM_NONCE_LAST_VALID")), seed, detail)


def test_18_the_first_five_uniforms_match_for_every_accepted_seed() -> None:
    per_seed = _evidence("rng_vectors")["per_seed"]
    assert per_seed, "no accepted per-seed vector"
    for seed, record in per_seed.items():
        assert _draw(_seeded(int(seed)), 5) == record["first_5"], seed


def test_19_twenty_draws_match_and_leave_the_accepted_state() -> None:
    vectors = _evidence("rng_vectors")
    state = _seeded(12345)
    assert _words(state) == vectors["seed_state_12345"]
    assert _draw(state, 20) == vectors["first_20_uniforms"]
    assert _words(state) == vectors["state_after_20"]


def test_20_the_jump_reproduces_every_accepted_stream() -> None:
    vba = _transcribe()
    streams = _evidence("jump_vectors")["streams"]
    assert streams, "no accepted stream vector"
    base = _seeded(12345)
    current, ladder = _copy(base), {0: _copy(base)}
    highest = max(int(k) for k in streams)
    for index in range(1, highest + 1):
        stepped, detail = _mk(0, 0, 0, 0, 0, 0), _Ref("")
        assert vba["SimRngJumpNextStream"](current, stepped, detail), detail.v
        current = stepped
        ladder[index] = _copy(stepped)
    for key, record in streams.items():
        state = ladder[int(key)]
        assert _words(state) == record["initial_state"], key
        assert _draw(_copy(state), 5) == record["first_5_uniforms"], key


def test_21_stream_k_is_algorithmic_beyond_the_design_target() -> None:
    """Stream 401 is in the accepted vectors so a 400-entry table cannot pass."""
    vba = _transcribe()
    streams = _evidence("jump_vectors")["streams"]
    assert "401" in streams, "the beyond-400 vector disappeared"
    for key, record in streams.items():
        state, detail = _mk(0, 0, 0, 0, 0, 0), _Ref("")
        assert vba["SimRngStreamInitialState"](_seeded(12345), _Ref(int(key)),
                                               state, detail), detail.v
        assert _words(state) == record["initial_state"], key
    state, detail = _mk(0, 0, 0, 0, 0, 0), _Ref("")
    assert vba["SimRngStreamInitialState"](_seeded(12345), _Ref(-1), state, detail) is False
    assert detail.v == "stream index: negative"


def test_22_the_component_assignment_reproduces_the_accepted_order() -> None:
    vba = _transcribe()
    accepted = _evidence("stream_assignment_vectors")
    costs = [f"CL-{i:03d}" for i in range(1, 201)]
    risks = [f"R-{i:03d}" for i in range(1, 101)]
    components, detail = [], _Ref("")
    assert vba["SimRngBuildComponentStreams"](
        costs, _Ref(len(costs)), risks, _Ref(len(risks)),
        _seeded(12345), components, detail), detail.v
    assert len(components) == accepted["total_components"] == 400

    # The Step-0 vectors label the kind by the component key's family, which is
    # the projection's own COST_SAMPLE / RISK_OCCURRENCE / RISK_SEVERITY.
    label = {_const(f"SIM_COMPONENT_{i}_DRIVER_KIND"): _const(f"SIM_COMPONENT_{i}_KEY").split("_")[0]
             for i in (1, 2, 3)}
    seen = [{"component": [label[c["DriverKind"]], c["PermanentId"], c["Role"]],
             "stream": c["StreamIndex"]} for c in components]
    assert seen[:10] == accepted["family_a_first_10"]
    assert seen[-4:] == accepted["family_a_last_4"]
    assert [c["StreamIndex"] for c in components] == list(range(400))


def test_23_each_component_carries_the_stream_its_index_names() -> None:
    """The single ladder walk gives the same states as stream k computed alone."""
    vba = _transcribe()
    costs = [f"CL-{i:03d}" for i in range(1, 4)]
    risks = [f"R-{i:03d}" for i in range(1, 3)]
    components, detail = [], _Ref("")
    base = _seeded(12345)
    assert vba["SimRngBuildComponentStreams"](
        costs, _Ref(3), risks, _Ref(2), base, components, detail), detail.v
    assert len(components) == 3 + 2 * 2
    for component in components:
        alone, detail = _mk(0, 0, 0, 0, 0, 0), _Ref("")
        assert vba["SimRngStreamInitialState"](
            base, _Ref(component["StreamIndex"]), alone, detail), detail.v
        assert _words(component["InitialState"]) == _words(alone), component["StreamIndex"]


def test_24_the_risks_interleave_and_do_not_form_two_blocks() -> None:
    vba = _transcribe()
    components, detail = [], _Ref("")
    assert vba["SimRngBuildComponentStreams"](
        ["CL-001"], _Ref(1), ["R-002", "R-001"], _Ref(2),
        _seeded(12345), components, detail), detail.v
    assert [(c["PermanentId"], c["Role"]) for c in components] == [
        ("CL-001", _const("SIM_COMPONENT_1_ROLE")),
        ("R-001", _const("SIM_COMPONENT_2_ROLE")),
        ("R-001", _const("SIM_COMPONENT_3_ROLE")),
        ("R-002", _const("SIM_COMPONENT_2_ROLE")),
        ("R-002", _const("SIM_COMPONENT_3_ROLE")),
    ]


def test_25_the_ordering_is_ordinal_and_reads_no_numeric_suffix() -> None:
    vba = _transcribe()
    components, detail = [], _Ref("")
    assert vba["SimRngBuildComponentStreams"](
        ["CL-999", "CL-1000", "CL-100"], _Ref(3), [], _Ref(0),
        _seeded(12345), components, detail), detail.v
    assert [c["PermanentId"] for c in components] == ["CL-100", "CL-1000", "CL-999"], (
        "a numeric-suffix reading would put CL-999 first"
    )


def test_26_the_caller_arrays_are_never_reordered() -> None:
    vba = _transcribe()
    costs = ["CL-003", "CL-001", "CL-002"]
    risks = ["R-002", "R-001"]
    components, detail = [], _Ref("")
    assert vba["SimRngBuildComponentStreams"](
        costs, _Ref(3), risks, _Ref(2), _seeded(12345), components, detail), detail.v
    assert costs == ["CL-003", "CL-001", "CL-002"], costs
    assert risks == ["R-002", "R-001"], risks


def test_27_row_order_does_not_change_the_assignment() -> None:
    vba = _transcribe()
    costs = [f"CL-{i:03d}" for i in range(1, 21)]
    risks = [f"R-{i:03d}" for i in range(1, 11)]
    runs = []
    for order in (list(range(20)), list(reversed(range(20)))):
        for risk_order in (list(range(10)), list(reversed(range(10)))):
            components, detail = [], _Ref("")
            assert vba["SimRngBuildComponentStreams"](
                [costs[i] for i in order], _Ref(20),
                [risks[i] for i in risk_order], _Ref(10),
                _seeded(12345), components, detail), detail.v
            runs.append([(c["PermanentId"], c["Role"], c["StreamIndex"],
                          tuple(_words(c["InitialState"]))) for c in components])
    assert all(run == runs[0] for run in runs)


def test_28_a_duplicate_identity_is_refused_not_deduplicated() -> None:
    vba = _transcribe()
    for costs, risks, message in (
        (["CL-001", "CL-001"], [], "duplicate cost line permanent id"),
        (["CL-001"], ["R-001", "R-001"], "duplicate risk permanent id"),
    ):
        components, detail = [], _Ref("")
        assert vba["SimRngBuildComponentStreams"](
            costs, _Ref(len(costs)), risks, _Ref(len(risks)),
            _seeded(12345), components, detail) is False
        assert message in detail.v
        assert components == [], "a refused model still produced components"


def test_29_a_blank_identity_and_a_negative_count_are_refused() -> None:
    vba = _transcribe()
    for costs, risks, message in (
        ([""], [], "blank permanent id"),
        (["CL-001"], [""], "blank permanent id"),
    ):
        components, detail = [], _Ref("")
        assert vba["SimRngBuildComponentStreams"](
            costs, _Ref(len(costs)), risks, _Ref(len(risks)),
            _seeded(12345), components, detail) is False
        assert message in detail.v
    for costCount, riskCount in ((-1, 0), (0, -1), (-1, -1)):
        components, detail = [], _Ref("")
        assert vba["SimRngBuildComponentStreams"](
            [], _Ref(costCount), [], _Ref(riskCount),
            _seeded(12345), components, detail) is False
        assert "negative driver count" in detail.v


def test_29a_zero_drivers_is_a_legal_model_and_is_not_refused() -> None:
    """No accepted contract requires a Cost Line or a Risk to exist.

    Phase 5 pins that an empty driver set is not refused once the model-level
    prerequisites resolve, and Phase 6 introduced no minimum of its own. An
    earlier draft of this module invented one; this is the regression that keeps
    it out.
    """
    vba = _transcribe()
    state = _seeded(12345)
    before = _words(state)
    components, detail = ["a stale result the caller must not read back"], _Ref("")
    assert vba["SimRngBuildComponentStreams"](
        [], _Ref(0), [], _Ref(0), state, components, detail) is True, detail.v
    assert detail.v == ""
    # ZERO JUMPS: the base state is the caller's, untouched.
    assert _words(state) == before
    # The logical count is zero, so no element of the carrier is present. The
    # accepted Phase-5 zero-count convention sizes it to one slot, and that slot
    # holds no component: a blank PermanentId is what SimRngOrderIds refuses.
    assert len(components) == 1
    assert components[0]["PermanentId"] == ""
    assert components[0]["DriverKind"] == "" and components[0]["Role"] == ""
    assert _words(components[0]["InitialState"]) == [0] * 6


def test_29b_an_empty_model_still_validates_the_base_state() -> None:
    """Zero work is not permission to accept a state the recurrence cannot be in."""
    vba = _transcribe()
    broken = _mk(0, 0, 0, 1, 1, 1)
    before = _words(broken)
    components, detail = [], _Ref("")
    assert vba["SimRngBuildComponentStreams"](
        [], _Ref(0), [], _Ref(0), broken, components, detail) is False
    assert "all zero" in detail.v
    assert _words(broken) == before
    assert components == [], "a refused empty model still wrote the carrier"
    # And the source validates the state BEFORE it takes the empty path.
    body = _procedure("SimRngBuildComponentStreams")
    assert body.index("SimRngValidateState(baseState, detail)") < \
        body.index("If total = 0 Then")


def test_29c_the_empty_path_reads_no_bound_from_either_driver_array() -> None:
    """Nothing is ordered, so nothing is indexed. The arrays are never touched."""
    body = _procedure("SimRngBuildComponentStreams")
    empty = body[body.index("If total = 0 Then"):body.index("If Not SimRngOrderIds")]
    for token in ("LBound", "costIds", "riskIds", "SimRngOrderIds",
                  "SimRngJumpNextStream", "costOrder", "riskOrder"):
        assert token not in empty, token
    assert "ReDim built(0 To 0)" in empty
    assert "components = built" in empty
    # And no minimum-driver rule of any spelling survives anywhere in the body.
    for invented in ("no driver", "at least one", "total < 1", "total >= 1",
                     "total < 1&", "declares no"):
        assert invented not in body, invented


def test_29d_the_empty_model_agrees_with_the_accepted_python_reference() -> None:
    """VBA source transcription against sim_rng.py, on the empty model.

    The correction exists to bring the two back into agreement; this is the test
    that says so directly. No new stochastic authority is generated here - the
    Python reference was accepted at Step 2 and is not re-derived.
    """
    from pccm_builder import RngReference
    from pccm_builder.sim_rng import RngState

    reference = RngReference.from_contracts(
        load_sim_contract(SPEC / "sim_contract.yaml"),
        load_contract(SPEC / "input_contract.yaml"),
    )
    # The accepted reference: no component, no stream state, and the base state
    # is validated and returned as supplied.
    assert reference.components_for([], []) == ()
    base = reference.fixed_seed_to_state(12345)
    assert reference.component_stream_states(base, ()) == ()
    assert reference.validate_state(base) == base

    # ...and the transcribed VBA: True, nothing emitted, nothing consumed.
    vba = _transcribe()
    state = _seeded(12345)
    assert _words(state) == list(base.words)
    components, detail = [], _Ref("")
    assert vba["SimRngBuildComponentStreams"](
        [], _Ref(0), [], _Ref(0), state, components, detail) is True, detail.v
    assert _words(state) == list(base.words), (
        "the VBA advanced a state the reference did not"
    )

    # Both refuse the same inadmissible base state on the same empty model.
    words = (0, 0, 0, 1, 1, 1)
    broken = _mk(*words)
    try:
        reference.component_stream_states(RngState.of(*words), ())
    except Exception as refusal:            # noqa: BLE001 - the reference's own error
        assert "zero" in str(refusal).lower(), str(refusal)
    else:  # pragma: no cover - the reference must refuse an absorbing state
        raise AssertionError("the Python reference accepted an all-zero component")
    components, detail = [], _Ref("")
    assert vba["SimRngBuildComponentStreams"](
        [], _Ref(0), [], _Ref(0), broken, components, detail) is False


def test_30_state_validation_refuses_every_inadmissible_word() -> None:
    vba = _transcribe()
    m1, m2 = _const("SIM_RNG_M1"), _const("SIM_RNG_M2")
    cases = [
        (_mk(m1, 1, 1, 1, 1, 1), "first word 0 is not below its modulus"),
        (_mk(1, 1, 1, m2, 1, 1), "second word 0 is not below its modulus"),
        (_mk(1, -1, 1, 1, 1, 1), "first word 1 is negative"),
        (_mk(1, 1, 1.5, 1, 1, 1), "first word 2 is not an integer"),
        (_mk(0, 0, 0, 1, 1, 1), "the first component is all zero"),
        (_mk(1, 1, 1, 0, 0, 0), "the second component is all zero"),
    ]
    for state, message in cases:
        detail = _Ref("")
        assert vba["SimRngValidateState"](state, detail) is False, message
        assert message in detail.v, (detail.v, message)
    detail = _Ref("")
    assert vba["SimRngValidateState"](_mk(m1 - 1, 0, 0, m2 - 1, 0, 0), detail), detail.v


def test_31_a_failing_draw_leaves_the_caller_state_untouched() -> None:
    vba = _transcribe()
    broken = _mk(0, 0, 0, 1, 1, 1)
    before = _words(broken)
    u, detail = _Ref(-1.0), _Ref("")
    assert vba["SimRngNextUniform"](broken, u, detail) is False
    assert _words(broken) == before, "a refused draw advanced the state"
    assert u.v == -1.0, "a refused draw wrote the uniform"
    jumped, detail = _mk(0, 0, 0, 0, 0, 0), _Ref("")
    assert vba["SimRngJumpNextStream"](broken, jumped, detail) is False
    assert _words(broken) == before
    assert _words(jumped) == [0] * 6


def test_32_every_uniform_lies_strictly_inside_the_open_interval() -> None:
    state = _seeded(1)
    values = [float(u) for u in _draw(state, 2000)]
    assert all(0.0 < u < 1.0 for u in values)
    assert len(set(values)) == len(values), "the stream repeated inside 2000 draws"


# ===========================================================================
# E. SHAPE CONFORMANCE - what the source is allowed to say
# ===========================================================================
_OWNED_VALUES = (
    "4294967087", "4294944443", "1403580", "810728", "527612", "1370589",
    "131072", "48271", "2147483647", "2147483646", "2147483645", "2328306549",
)


def test_33_no_owned_value_is_restated_as_a_literal() -> None:
    """Every operational number has an owner; a second copy is a second authority."""
    code = _code()
    for value in _OWNED_VALUES:
        assert value not in code, value
    for name, (kind, literal) in _generated_constants().items():
        if not name.startswith("SIM_JUMP_A"):
            continue
        assert literal.rstrip(".0") not in code or float(literal) in (0.0, 1.0, 2.0), name
    # And nothing bigger than a loop bound is spelled at all.
    literals = {match.group(0).rstrip("#")
                for match in re.finditer(r"\b\d+#?", code)}
    assert literals <= {"0", "1", "2"}, sorted(literals)


def test_34_the_reduction_is_fix_based_and_never_the_vba_mod() -> None:
    code = _code()
    assert not re.search(r"\bMod\b", code), "VBA Mod coerces to an integer type"
    assert not re.search(r"\bInt\(", code), "Int floors; the accepted reduction truncates"
    assert not re.search(r"\bRound\(", code)
    assert not re.search(r"\bCInt\(", code)
    reduce_body = _procedure("SimRngReduce")
    assert "k = Fix(p / m)" in reduce_body
    assert "r = p - k * m" in reduce_body
    assert "If r < 0# Then r = r + m" in reduce_body


def test_35_the_negative_remainder_correction_is_present_in_both_primitives() -> None:
    assert re.search(r"If r < 0# Then r = r \+ m", _procedure("SimRngReduce"))
    assert re.search(r"If v < 0# Then v = v \+ m", _procedure("SimRngMultModM"))


def test_36_the_safe_multiply_uses_the_projected_split_and_only_that() -> None:
    body = _procedure("SimRngMultModM")
    assert body.count("SIM_JUMP_DECOMPOSITION_H") == 3, body
    assert _const("SIM_JUMP_DECOMPOSITION_H") == 2 ** 17
    for statement in ("a1 = Fix(rest / SIM_JUMP_DECOMPOSITION_H)",
                      "rest = rest - a1 * SIM_JUMP_DECOMPOSITION_H",
                      "v = a1 * s",
                      "v = v - Fix(v / m) * m",
                      "v = v * SIM_JUMP_DECOMPOSITION_H + rest * s + c",
                      "v = rest * s + c"):
        assert statement in body, statement
    # One primitive, shared: the AUTO power and the jump both go through it.
    assert "SimRngMultModM" in _procedure("SimRngAutoSeedFromNonce")
    assert "SimRngMultModM" in _procedure("SimRngJumpRow")
    assert _code().count("Function SimRngMultModM") == 1


def test_37_the_recurrence_pairs_each_coefficient_with_its_own_word() -> None:
    body = _procedure("MRG32k3aStep")
    assert "signed = SIM_RNG_A12 * state.S11 - SIM_RNG_A13N * state.S10" in body
    assert "p1 = SimRngReduce(signed, SIM_RNG_M1)" in body
    assert "signed = SIM_RNG_A21 * state.S22 - SIM_RNG_A23N * state.S20" in body
    assert "p2 = SimRngReduce(signed, SIM_RNG_M2)" in body
    for statement in ("advanced.S10 = state.S11", "advanced.S11 = state.S12",
                      "advanced.S12 = p1", "advanced.S20 = state.S21",
                      "advanced.S21 = state.S22", "advanced.S22 = p2"):
        assert statement in body, statement


def test_38_the_combination_boundary_is_the_accepted_one() -> None:
    body = _procedure("SimRngNextUniform")
    assert "If p1 <= p2 Then" in body, "the boundary decides when the residues are equal"
    assert "candidate = (p1 - p2 + SIM_RNG_M1) * SimRngNorm()" in body
    assert "candidate = (p1 - p2) * SimRngNorm()" in body
    assert "If Not (candidate > 0# And candidate < 1#) Then" in body


def test_39_the_jump_reverses_the_triples_in_and_out() -> None:
    body = _procedure("SimRngJumpNextStream")
    for statement in ("inFirst(0) = state.S12", "inFirst(1) = state.S11",
                      "inFirst(2) = state.S10", "inSecond(0) = state.S22",
                      "inSecond(1) = state.S21", "inSecond(2) = state.S20",
                      "candidate.S10 = outFirst(2)", "candidate.S11 = outFirst(1)",
                      "candidate.S12 = outFirst(0)", "candidate.S20 = outSecond(2)",
                      "candidate.S21 = outSecond(1)", "candidate.S22 = outSecond(0)"):
        assert statement in body, statement
    # Every matrix element is read from the projection, row by row, and each row
    # goes through the safe multiply.
    for component, modulus in (("A1", "SIM_RNG_M1"), ("A2", "SIM_RNG_M2")):
        for row in (1, 2, 3):
            names = ", ".join(f"SIM_JUMP_{component}_R{row}C{col}" for col in (1, 2, 3))
            assert names in re.sub(r"\s+", " ", body), names
        assert body.count(modulus) == 3
    assert _procedure("SimRngJumpRow").count("SimRngMultModM") == 3


def test_40_there_is_no_lookup_table_and_no_substream() -> None:
    code = _code()
    for token in ("Substream", "SubStream", "2^76", "SIM_JUMP_A3", "Array(",
                  "Choose(", "Switch(", "Select Case"):
        assert token not in code, token
    ladder = _procedure("SimRngStreamInitialState")
    assert "For index = 1 To k" in ladder
    assert "SimRngJumpNextStream(current, stepped, detail)" in ladder
    assert "400" not in code and "SIM_DESIGN" not in code


def test_41_the_ordering_is_ordinal_and_reads_nothing_out_of_the_identity() -> None:
    body = _procedure("SimRngOrderIds")
    assert body.count("vbBinaryCompare") == 2
    assert "vbTextCompare" not in _code()
    for token in ("LCase", "UCase", "Trim", "Val(", "Right(", "Left(", "Mid(",
                  "InStr", "Replace(", "Split(", "Sort", "Application"):
        assert token not in _code(), token


def test_42_every_constant_the_module_reads_has_an_owner() -> None:
    referenced = {name for name in _module().referenced_upper_identifiers
                  if name.startswith("SIM_")}
    assert referenced, "the module reads no projected constant"
    missing = referenced - set(_generated_constants())
    assert not missing, sorted(missing)
    # Nothing the module reads is declared locally: modSimContract owns them all.
    assert not _module().constants, _module().constants


def test_43_every_output_is_committed_after_its_last_check() -> None:
    for procedure, guard, commit in (
        ("SimRngStateFromFixedSeed", "SimRngValidateState(candidate, detail)",
         "state = candidate"),
        ("SimRngNextUniform", "candidate > 0# And candidate < 1#", "state = advanced"),
        ("SimRngJumpNextStream", "SimRngValidateState(candidate, detail)",
         "jumped = candidate"),
        ("SimRngAutoSeedFromNonce", "left its residue class", "seed = CLng(result)"),
        ("SimRngBuildComponentStreams", "SimRngJumpNextStream(current, stepped, detail)",
         "components = built"),
    ):
        body = _procedure(procedure)
        assert guard in body and commit in body, procedure
        # rindex on both: SimRngBuildComponentStreams commits the carrier on two
        # paths - the zero-component one and the ladder - and the LADDER commit
        # is the one this guard governs.
        assert body.rindex(guard) < body.rindex(commit), procedure
    # The zero-component commit has its own guard, and it is the state check.
    body = _procedure("SimRngBuildComponentStreams")
    assert body.index("SimRngValidateState(baseState, detail)") < \
        body.index("components = built")


def test_44_caller_arrays_are_read_through_their_own_lower_bound() -> None:
    """A 1-based caller array must not silently shift the assignment."""
    code = _code()
    for name in ("costIds", "riskIds", "ids"):
        indexed = list(re.finditer(rf"\b{name}\((?!\))", code))
        assert indexed, f"{name} is never read"
        for match in indexed:
            tail = code[match.end():]
            assert tail.startswith("LBound("), (
                f"{name} indexed without LBound: {tail[:40]}"
            )


def test_45_the_transcription_read_the_whole_module() -> None:
    """No procedure was silently skipped by the compiler above."""
    vba = _transcribe()
    assert set(vba["_procs"]) == set(_module().procedures)
    for name in vba["_procs"]:
        assert callable(vba[name]), name
    assert set(vba["_types"]) == {"SimRngState", "SimRngComponent"}


def test_46_the_auto_mapping_is_logarithmic_in_the_nonce() -> None:
    """Square-and-multiply, not stepping: a linear walk is unusable near the period."""
    vba = _transcribe()
    limit = 2 * math.log2(_const("SIM_NONCE_LAST_VALID")) + 2
    real, calls = vba["SimRngMultModM"], [0]

    def counted(*args):
        calls[0] += 1
        if calls[0] > limit:
            raise AssertionError("the AUTO mapping is not O(log nonce)")
        return real(*args)

    vba["SimRngMultModM"] = counted
    try:
        seed, detail = _Ref(0), _Ref("")
        assert vba["SimRngAutoSeedFromNonce"](
            _Ref(_const("SIM_NONCE_LAST_VALID")), seed, detail), detail.v
    finally:
        vba["SimRngMultModM"] = real
    assert calls[0] <= limit, calls[0]
    assert "half = Fix(remaining / 2#)" in _procedure("SimRngAutoSeedFromNonce")


def test_47_the_jump_ladder_is_walked_once_per_component() -> None:
    """O(N) jumps for N components, not O(N^2)."""
    vba = _transcribe()
    real, calls = vba["SimRngJumpNextStream"], [0]

    def counted(*args):
        calls[0] += 1
        return real(*args)

    vba["SimRngJumpNextStream"] = counted
    try:
        components, detail = [], _Ref("")
        assert vba["SimRngBuildComponentStreams"](
            [f"CL-{i:03d}" for i in range(1, 31)], _Ref(30),
            [f"R-{i:03d}" for i in range(1, 11)], _Ref(10),
            _seeded(12345), components, detail), detail.v
    finally:
        vba["SimRngJumpNextStream"] = real
    assert len(components) == 50
    assert calls[0] == 49, calls[0]
