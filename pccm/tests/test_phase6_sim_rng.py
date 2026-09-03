#!/usr/bin/env python3
"""PCCM Phase 6 Step-2 conformance tests for the Python RNG reference.

`builder/pccm_builder/sim_rng.py` is the ORACLE for the RNG the accepted
simulation contract describes. These tests prove it reproduces the retained
Step-0 vectors EXACTLY - state word for state word, and uniform for uniform at
the binary64 value, compared through `float.hex()` rather than any tolerance.
No approximate comparison belongs to an RNG backbone.

The retained evidence is a TEST oracle only. `sim_rng.py` never reads it; a
separate test asserts that, because a reference that consulted the vectors at run
time would be marking its own homework.

NOTHING HERE SAMPLES. There is no Uniform, Triangular or Beta-PERT transform, no
Bernoulli trial, no iteration, no statistic and no digest - not even
`x = (1-u)a + ub`, which is Step 3.

Runs standalone or under pytest.
"""

from __future__ import annotations

import ast
import json
import sys
from pathlib import Path

import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    Component,
    RngReference,
    RngState,
    load_contract,
    load_sim_contract,
)
from pccm_builder.sim_rng import (  # noqa: E402
    COST_KIND,
    RISK_KIND,
    ROLE_OCCURRENCE,
    ROLE_SEVERITY,
    ROLE_VALUE,
)

CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
SIM_PATH = PCCM_ROOT / "spec" / "sim_contract.yaml"
EVIDENCE = PCCM_ROOT / "evidence" / "phase6_step0"

_REF: RngReference | None = None


def _ref() -> RngReference:
    global _REF
    if _REF is None:
        _REF = RngReference.from_contracts(
            load_sim_contract(SIM_PATH), load_contract(CONTRACT_PATH)
        )
    return _REF


def _vectors(name: str) -> dict:
    return json.loads((EVIDENCE / "vectors" / name).read_text(encoding="utf-8"))


def _raw(name: str) -> dict:
    return json.loads((EVIDENCE / "raw" / name).read_text(encoding="utf-8"))


def _same_double(actual: float, retained_repr: str) -> bool:
    """EXACT binary64 identity, not a tolerance.

    The retained vectors store `repr()` of the Double the accepted formula
    produced. Comparing through `float.hex()` compares the bit pattern, so a
    value that merely prints the same cannot pass.
    """
    return float.hex(actual) == float.hex(float(retained_repr))


# ===========================================================================
# A - D. seeding
# ===========================================================================
def test_01_the_reference_derives_its_constants_from_the_contract() -> None:
    reference = _ref()
    constants = yaml.safe_load(SIM_PATH.read_text(encoding="utf-8"))["rng"]["constants"]
    assert reference.m1 == constants["m1"] == 4294967087
    assert reference.m2 == constants["m2"] == 4294944443
    assert (reference.a12, reference.a13n) == (constants["a12"], constants["a13n"])
    assert (reference.a21, reference.a23n) == (constants["a21"], constants["a23n"])
    assert float.hex(reference.norm) == float.hex(constants["norm"])


def test_02_the_seed_domain_comes_from_the_owning_input_contract() -> None:
    """Not a second constant that could drift from D6-19a's owner."""
    reference = _ref()
    validation = load_contract(CONTRACT_PATH).inputs["random_seed"].validation
    assert reference.seed_min == int(validation["formula1"]) == 1
    assert reference.seed_max == int(validation["formula2"]) == 2147483646


def test_03_scalar_seed_repeats_into_all_six_words() -> None:
    vectors = _vectors("seed_vectors.json")
    assert vectors["seed_to_state_rule"] == "state = [seed] * 6"
    for example in vectors["examples"]:
        state = _ref().fixed_seed_to_state(example["seed"])
        assert state.as_list() == example["state"]
        assert state.as_list() == [example["seed"]] * 6


def test_04_the_four_retained_fixed_seeds_map_exactly() -> None:
    for seed in (1, 2, 12345, 2147483646):
        assert _ref().fixed_seed_to_state(seed).as_list() == [seed] * 6


def test_05_an_invalid_fixed_seed_is_refused() -> None:
    from pccm_builder import SimRngError

    for bad in (0, -1, 2147483647, 2 ** 31, 1.0, 12345.0, "12345", None, True, False):
        try:
            _ref().fixed_seed_to_state(bad)  # type: ignore[arg-type]
        except SimRngError:
            continue
        raise AssertionError(f"seed {bad!r} was accepted")


def test_06_auto_nonce_maps_to_the_retained_seeds() -> None:
    retained = _raw("seed_map.json")["nonce_to_seed_pairs"]
    for pair in retained:
        nonce = pair["auto_nonce"]
        if nonce >= _ref().nonce_exhausted:
            continue
        assert _ref().auto_seed_from_nonce(nonce) == pair["effective_seed"], nonce
    assert _ref().auto_seed_from_nonce(0) == 1
    assert _ref().auto_seed_from_nonce(1) == 48271


def test_07_the_auto_nonce_boundaries_are_exact() -> None:
    from pccm_builder import SimRngError

    reference = _ref()
    assert reference.nonce_exhausted == 2147483646
    assert reference.auto_seed_from_nonce(reference.nonce_exhausted - 1) >= 1
    for bad in (-1, reference.nonce_exhausted, reference.nonce_exhausted + 1, 1.0, True):
        try:
            reference.auto_seed_from_nonce(bad)  # type: ignore[arg-type]
        except SimRngError:
            continue
        raise AssertionError(f"nonce {bad!r} was accepted")


def test_08_auto_seeds_stay_inside_the_admissible_domain() -> None:
    """The mapping's whole point: every nonce yields a usable FIXED-domain seed."""
    reference = _ref()
    for nonce in (0, 1, 2, 3, 10, 1000, 999983, 2147483645):
        seed = reference.auto_seed_from_nonce(nonce)
        assert reference.seed_min <= seed <= reference.seed_max, nonce
        reference.fixed_seed_to_state(seed)


def test_09_the_auto_mapping_is_a_modular_power_not_stepping() -> None:
    """Agreement over a prefix, and the same answer where stepping is unusable."""
    reference = _ref()
    stepped = 1
    for nonce in range(0, 3000):
        assert reference.auto_seed_from_nonce(nonce) == stepped, nonce
        stepped = (stepped * reference.auto_multiplier) % reference.auto_modulus
    # At this nonce a stepping implementation would need 900 million
    # multiplications; the authority is a power, so it is immediate.
    assert reference.auto_seed_from_nonce(900_000_000) == pow(
        reference.auto_multiplier, 900_000_000, reference.auto_modulus
    )


# ===========================================================================
# E - G. the recurrence
# ===========================================================================
def test_10_the_retained_first_twenty_uniforms_reproduce_exactly() -> None:
    retained = _vectors("rng_vectors.json")
    state = _ref().fixed_seed_to_state(12345)
    assert state.as_list() == retained["seed_state_12345"]
    drawn, after = _ref().uniforms(state, 20)
    assert len(drawn) == len(retained["first_20_uniforms"]) == 20
    for index, (actual, expected) in enumerate(zip(drawn, retained["first_20_uniforms"])):
        assert _same_double(actual, expected), f"draw {index + 1}: {actual!r} != {expected}"
    assert after.as_list() == retained["state_after_20"]


def test_11_the_published_first_uniform_is_reproduced() -> None:
    """The RngStreams default-state conformance point, kept explicit."""
    draw = _ref().next_uniform(_ref().fixed_seed_to_state(12345))
    assert repr(draw.uniform) == "0.12701112204657714"
    assert float.hex(draw.uniform) == float.hex(0.12701112204657714)


def test_12_the_retained_per_seed_vectors_reproduce_exactly() -> None:
    per_seed = _vectors("rng_vectors.json")["per_seed"]
    for text, case in sorted(per_seed.items()):
        seed = int(text)
        state = _ref().fixed_seed_to_state(seed)
        assert state.as_list() == case["initial_state"]
        drawn, _ = _ref().uniforms(state, len(case["first_5"]))
        for index, (actual, expected) in enumerate(zip(drawn, case["first_5"])):
            assert _same_double(actual, expected), f"seed {seed} draw {index + 1}"


def test_13_the_state_advance_is_exactly_the_contracted_shift() -> None:
    reference = _ref()
    before = reference.fixed_seed_to_state(12345)
    s10, s11, s12, s20, s21, s22 = before.words
    after = reference.next_uniform(before).state
    p1 = (reference.a12 * s11 - reference.a13n * s10) % reference.m1
    p2 = (reference.a21 * s22 - reference.a23n * s20) % reference.m2
    assert after.as_list() == [s11, s12, p1, s21, s22, p2]


def test_14_every_uniform_over_the_retained_corpus_is_strictly_inside_0_1() -> None:
    reference = _ref()
    for seed in (1, 2, 12345, 2147483646):
        state = reference.fixed_seed_to_state(seed)
        drawn, _ = reference.uniforms(state, 500)
        for index, value in enumerate(drawn):
            assert 0.0 < value < 1.0, f"seed {seed} draw {index + 1} produced {value!r}"


def test_15_an_invalid_state_is_refused_at_the_boundary() -> None:
    from pccm_builder import SimRngError

    reference = _ref()
    for bad in (
        RngState((0, 0, 0, 1, 1, 1)),
        RngState((1, 1, 1, 0, 0, 0)),
        RngState((reference.m1, 1, 1, 1, 1, 1)),
        RngState((1, 1, 1, reference.m2, 1, 1)),
        RngState((-1, 1, 1, 1, 1, 1)),
    ):
        try:
            reference.next_uniform(bad)
        except SimRngError:
            continue
        raise AssertionError(f"invalid state {bad.as_list()} was accepted")


def test_16_the_state_is_immutable_and_advancing_returns_a_new_one() -> None:
    reference = _ref()
    state = reference.fixed_seed_to_state(12345)
    words = state.as_list()
    first = reference.next_uniform(state)
    second = reference.next_uniform(state)
    assert state.as_list() == words, "the input state moved"
    assert first.state == second.state, "the same input gave two different successors"
    assert float.hex(first.uniform) == float.hex(second.uniform)


# ===========================================================================
# H - I. the jump
# ===========================================================================
def test_17_the_jump_matrices_come_from_the_contract() -> None:
    declared = yaml.safe_load(SIM_PATH.read_text(encoding="utf-8"))["jump"]
    assert [list(r) for r in _ref().jump_a1] == declared["a1_p127"]
    assert [list(r) for r in _ref().jump_a2] == declared["a2_p127"]
    retained = _raw("jump.json")
    assert [list(r) for r in _ref().jump_a1] == retained["A1p127_derived_from_recurrence"]
    assert [list(r) for r in _ref().jump_a2] == retained["A2p127_derived_from_recurrence"]


def test_18_the_retained_stream_states_reproduce_exactly() -> None:
    retained = _vectors("jump_vectors.json")
    base = RngState(tuple(retained["streams"]["0"]["initial_state"]))
    for text, case in sorted(retained["streams"].items(), key=lambda kv: int(kv[0])):
        k = int(text)
        state = _ref().stream_initial_state(base, k)
        assert state.as_list() == case["initial_state"], f"stream {k}"


def test_19_the_retained_stream_uniforms_reproduce_exactly() -> None:
    retained = _vectors("jump_vectors.json")
    for text, case in sorted(retained["streams"].items(), key=lambda kv: int(kv[0])):
        state = RngState(tuple(case["initial_state"]))
        drawn, after = _ref().uniforms(state, len(case["first_5_uniforms"]))
        for index, (actual, expected) in enumerate(zip(drawn, case["first_5_uniforms"])):
            assert _same_double(actual, expected), f"stream {text} draw {index + 1}"
        assert after.as_list() == case["state_after_5"], f"stream {text}"


def test_20_stream_one_matches_the_published_second_stream_state() -> None:
    """The RngStreams conformance point that settles the jump orientation."""
    base = _ref().fixed_seed_to_state(12345)
    assert _ref().stream_initial_state(base, 1).as_list() == [
        3692455944, 1366884236, 2968912127, 335948734, 4161675175, 475798818,
    ]


def test_21_the_required_stream_indices_are_covered() -> None:
    retained = set(_vectors("jump_vectors.json")["streams"])
    assert {"0", "1", "7", "399"} <= retained
    beyond = [int(k) for k in retained if int(k) > 400]
    assert beyond, "the retained vectors carry no stream beyond 400"
    base = _ref().fixed_seed_to_state(12345)
    for k in sorted(int(x) for x in retained):
        expected = _vectors("jump_vectors.json")["streams"][str(k)]["initial_state"]
        assert _ref().stream_initial_state(base, k).as_list() == expected, k


def test_22_stream_zero_is_the_base_and_jumps_compose() -> None:
    reference = _ref()
    base = reference.fixed_seed_to_state(12345)
    assert reference.stream_initial_state(base, 0) == base
    once = reference.jump_to_next_stream(base)
    assert reference.stream_initial_state(base, 1) == once
    assert reference.stream_initial_state(base, 7) == reference.stream_initial_state(once, 6)


def test_23_a_jump_is_not_a_seek_to_an_iteration() -> None:
    """Distinct concepts: a jump moves 2^127 draws, not `k` draws."""
    reference = _ref()
    base = reference.fixed_seed_to_state(12345)
    jumped = reference.jump_to_next_stream(base)
    _, stepped = reference.uniforms(base, 1)
    assert jumped != stepped
    assert not hasattr(reference, "seek_to_iteration")


# ===========================================================================
# J - L. component streams
# ===========================================================================
def _design_target_components() -> tuple[Component, ...]:
    return _ref().components_for(
        [f"CL-{i:03d}" for i in range(1, 201)],
        [f"R-{i:03d}" for i in range(1, 101)],
    )


def test_24_the_retained_assignments_reproduce_exactly() -> None:
    retained = _vectors("stream_assignment_vectors.json")
    assignment = {
        tuple(component.as_list()): index
        for component, index in _ref().assign_component_streams(_design_target_components())
    }
    assert len(assignment) == retained["total_components"] == 400
    for entry in retained["family_a_first_10"] + retained["family_a_last_4"]:
        key = tuple(entry["component"])
        assert assignment[key] == entry["stream"], key


def test_25_risks_interleave_occurrence_and_severity_per_risk() -> None:
    """NOT three global blocks.

    Reading `component_kind` as the component KEY - `COST_SAMPLE`,
    `RISK_OCCURRENCE`, `RISK_SEVERITY` - would put every occurrence stream before
    every severity stream. The retained tail is `R-099 occ, R-099 sev, R-100 occ,
    R-100 sev`, so the kind axis is the DRIVER kind and the role is a separate,
    later key.
    """
    assignment = _ref().assign_component_streams(_design_target_components())
    tail = [(c.kind, c.permanent_id, c.role, index) for c, index in assignment[-4:]]
    assert tail == [
        (RISK_KIND, "R-099", ROLE_OCCURRENCE, 396),
        (RISK_KIND, "R-099", ROLE_SEVERITY, 397),
        (RISK_KIND, "R-100", ROLE_OCCURRENCE, 398),
        (RISK_KIND, "R-100", ROLE_SEVERITY, 399),
    ]
    kinds = [c.kind for c, _ in assignment]
    assert kinds[:200] == [COST_KIND] * 200
    assert kinds[200:] == [RISK_KIND] * 200


def test_26_the_component_count_rule_holds() -> None:
    """`C + 2R`."""
    for cost, risk in ((200, 100), (1, 1), (5, 0), (0, 3)):
        components = _ref().components_for(
            [f"CL-{i:03d}" for i in range(1, cost + 1)],
            [f"R-{i:03d}" for i in range(1, risk + 1)],
        )
        assert len(components) == cost + 2 * risk
        assert len(_ref().assign_component_streams(components)) == cost + 2 * risk


def test_27_assignment_is_a_bijection_onto_0_to_n_minus_1() -> None:
    assignment = _ref().assign_component_streams(_design_target_components())
    assert [index for _, index in assignment] == list(range(400))


def test_28_physical_row_order_changes_nothing() -> None:
    import random

    canonical = {
        tuple(c.as_list()): index
        for c, index in _ref().assign_component_streams(_design_target_components())
    }
    rnd = random.Random(20260825)
    for _ in range(5):
        shuffled = list(_design_target_components())
        rnd.shuffle(shuffled)
        assert {
            tuple(c.as_list()): index
            for c, index in _ref().assign_component_streams(shuffled)
        } == canonical


def test_29_ordering_is_ordinal_utf16_and_not_numeric() -> None:
    """`CL-1000` sorts BEFORE `CL-999`. Deterministic and portable, not numeric."""
    ids = ["CL-001", "CL-002", "CL-999", "CL-1000", "CL-0001"]
    components = _ref().components_for(ids, [])
    ordered = [c.permanent_id for c, _ in _ref().assign_component_streams(components)]
    assert ordered == ["CL-0001", "CL-001", "CL-002", "CL-1000", "CL-999"]
    assert ordered != sorted(ids, key=lambda s: int(s.split("-")[1]))


def test_30_the_id_key_is_utf16_code_units_not_code_points() -> None:
    """Explicit authority, not an accidental property of today's ASCII IDs."""
    from pccm_builder.calc_fingerprint import utf16_sort_key

    reference = _ref()
    for identifier in ("CL-001", "R-100", "CL-1000"):
        component = Component(COST_KIND, identifier, ROLE_VALUE)
        assert reference.canonical_sort_key(component)[1] == utf16_sort_key(identifier)
    # Above the BMP the two orders disagree; the accepted key follows UTF-16.
    astral, private = "\U00010000", ""
    assert utf16_sort_key(astral) < utf16_sort_key(private)
    assert astral > private, "code-point order disagrees, which is the whole point"


def test_31_component_stream_states_walk_the_ladder() -> None:
    reference = _ref()
    base = reference.fixed_seed_to_state(12345)
    components = reference.components_for(["CL-001", "CL-002"], ["R-001"])
    walked = reference.component_stream_states(base, components)
    assert [index for _, index, _ in walked] == [0, 1, 2, 3]
    for _, index, state in walked:
        assert state == reference.stream_initial_state(base, index)


def test_32_adding_a_driver_may_shift_later_streams_and_that_is_accepted() -> None:
    """D6-16 Family A's stated consequence, preserved rather than smoothed over."""
    reference = _ref()
    before = {
        tuple(c.as_list()): i
        for c, i in reference.assign_component_streams(
            reference.components_for(["CL-002", "CL-003"], [])
        )
    }
    after = {
        tuple(c.as_list()): i
        for c, i in reference.assign_component_streams(
            reference.components_for(["CL-001", "CL-002", "CL-003"], [])
        )
    }
    assert before[(COST_KIND, "CL-002", ROLE_VALUE)] == 0
    assert after[(COST_KIND, "CL-002", ROLE_VALUE)] == 1


# ===========================================================================
# M - N. scope discipline
# ===========================================================================
SIM_RNG = PCCM_ROOT / "builder" / "pccm_builder" / "sim_rng.py"


def test_33_the_reference_reads_no_evidence_at_run_time() -> None:
    """SEMANTIC, not a substring scan.

    The module's own prose says it must not read `evidence/`, which is exactly
    the sentence a substring scan would trip over. What matters is that no import
    and no file-access call reaches the evidence package.
    """
    tree = ast.parse(SIM_RNG.read_text(encoding="utf-8"))
    readers = ("open", "read_text", "read_bytes", "load", "loads", "glob", "rglob")
    for node in ast.walk(tree):
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            names = [a.name for a in node.names] + [getattr(node, "module", "") or ""]
            for name in names:
                assert "evidence" not in name, name
                assert "phase6_step0" not in name, name
        if isinstance(node, ast.Call):
            name = getattr(node.func, "attr", None) or getattr(node.func, "id", None)
            if name in readers:
                raise AssertionError(f"sim_rng.py calls {name}() - it reads nothing at all")
    # The decisive fact: the module performs NO file access of any kind, so it
    # cannot read the evidence package however a path were spelled. Prose that
    # mentions `evidence/` to say it is off limits is not a dependency.


def test_34_the_reference_uses_no_third_party_or_stdlib_rng() -> None:
    source = SIM_RNG.read_text(encoding="utf-8")
    tree = ast.parse(source)
    imported = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported.add(node.module.split(".")[0])
    for banned in ("random", "secrets", "numpy", "np", "scipy"):
        assert banned not in imported, f"sim_rng.py imports {banned}"
    for banned in ("random.", "secrets.", "np.random", "default_rng", "Randomize", "Rnd("):
        assert banned not in source, f"sim_rng.py references {banned}"


def test_35_no_sampler_or_simulation_exists_in_the_reference() -> None:
    """Step 2 produces raw MRG uniforms and stream identities. Nothing else."""
    source = SIM_RNG.read_text(encoding="utf-8")
    tree = ast.parse(source)
    defined = {
        node.name
        for node in ast.walk(tree)
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }
    for banned in (
        "sample_uniform", "sample_triangular", "sample_beta", "sample_pert",
        "cheng_bb", "cheng_bc", "bernoulli", "occurrence", "severity",
        "iterate", "run_simulation", "percentile", "result_digest", "contingency",
    ):
        assert not any(banned in name for name in defined), f"sim_rng.py defines {banned}"
    # Not even the Uniform DISTRIBUTION transform, which is Step 3.
    for banned in ("(1 - u) * a", "(1-u)*a", "MostLikely", "Quantity", "Knom", "Kpv"):
        assert banned not in source, f"sim_rng.py contains {banned}"


def test_36_the_only_phase6_vba_is_the_generator_backbone() -> None:
    """Step 2 authorised no VBA at all. Each later step authorised one module.

    The assertion is unchanged in substance - a Phase-6 module may not appear
    without a step that authorises it - and its right-hand side names the six
    that have. The algorithm token is still absent from every module except the
    one D6-11 scopes it to: modSimSample consumes randomness through the
    modSimRng public surface, and modSimReport orchestrates through it, and
    neither is granted an exception of its own.
    """
    src = PCCM_ROOT / "src" / "vba"
    names = {p.name for p in src.glob("*.bas")} | {p.name for p in src.glob("*.cls")}
    for authorised in ("modSimRng.bas", "modSimSample.bas", "modSimEngine.bas",
                       "modSimStats.bas", "modSimFingerprint.bas",
                       "modSimReport.bas"):
        assert authorised in names, authorised
    # Nothing beyond them WITHOUT A STEP THAT AUTHORISES IT - which is the
    # claim, and it is unchanged. `modSimSensitivity` arrived with P7-2, so it
    # is named here exactly as each Phase-6 module was named when its own step
    # landed, rather than admitted by loosening the comparison.
    unauthorised = {n for n in names if n.startswith("modSim")} - {
        "modSimRng.bas", "modSimSample.bas", "modSimEngine.bas", "modSimStats.bas",
        "modSimFingerprint.bas", "modSimNonce.bas", "modSimReport.bas",
        "modSimSensitivity.bas", "modSimPostReport.bas"}
    assert unauthorised == set(), sorted(unauthorised)
    for path in sorted(src.glob("*.bas")):
        if path.stem == "modSimRng":
            continue
        assert "MRG32k3a" not in path.read_text(encoding="utf-8", errors="replace"), path.name


def test_37_the_reference_holds_no_global_mutable_state() -> None:
    """No singleton, no module-level generator, no hidden seeding."""
    source = SIM_RNG.read_text(encoding="utf-8")
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    assert target.id.isupper(), (
                        f"module-level mutable binding {target.id!r}; only locked constants "
                        "belong at module scope"
                    )
    tree_again = ast.parse(source)
    assert not [n for n in ast.walk(tree_again) if isinstance(n, (ast.Global, ast.Nonlocal))], (
        "sim_rng.py rebinds module or enclosing state"
    )


def test_38_uniforms_validates_the_state_even_at_zero_count() -> None:
    """Step-2 cleanup, authorised in the Step-3 round.

    The zero-draw case previously returned whatever it was handed, because the
    loop that would have validated never ran. Harmless in isolation; a hole worth
    closing before samplers start passing states through this API.
    """
    from pccm_builder import SimRngError

    for bad in (
        RngState((0, 0, 0, 1, 1, 1)),
        RngState((1, 1, 1, 0, 0, 0)),
        RngState((_ref().m1, 1, 1, 1, 1, 1)),
    ):
        try:
            _ref().uniforms(bad, 0)
        except SimRngError:
            continue
        raise AssertionError(f"count 0 accepted the invalid state {bad.as_list()}")

    good = _ref().fixed_seed_to_state(12345)
    drawn, after = _ref().uniforms(good, 0)
    assert drawn == () and after == good


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
