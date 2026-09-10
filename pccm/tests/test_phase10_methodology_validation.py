#!/usr/bin/env python3
"""P10-3 MUTATION CONTROLS.

A WORDING CONTROL THAT CANNOT FAIL IS THE WORST KIND. The Methodology sheet is
prose, and prose passes any test that merely asks whether text exists. So every
control below breaks ONE sentence, one term, one number or one authority, reruns
the WHOLE conformance battery against the damaged copy, and requires a NAMED
control among the refusers - not just any failure somewhere.

WHAT IS MUTATED, AND WHY THESE.

  M5    profiling made to SCALE rather than distribute, and the 100% rule
        dropped. Both are the misunderstanding the section exists to prevent.

  M8    every way the two annual objects can be conflated: the ladder made to
        sum, the selected profile made not to, the interpolation identity
        corrupted, the two terms merged, and the forbidden sentence written out
        in the words the specification named.

  M10   "undefined" turned into "zero", the engine's own status label
        paraphrased, and the tornado exclusion dropped.

  M11   a fifth state added, a state dropped, a state renamed, and the REFUSED
        paragraph deleted.

  THE TWO VERSIONS  each authority moved on its own, the release string put back
        to a phase string, a historical record rewritten - and the collapse this
        batch was most likely to commit: the builder version derived from the
        model version while the two happen to agree.

Nothing here writes to the repository. Damaged manifests are dumped to a
temporary file so the real loader validates them, and the conformance module's
cache is pointed at the result for exactly one control.

Runs standalone or under pytest.
"""
from __future__ import annotations

import copy
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402
import yaml  # noqa: E402

from pccm_builder.spec_loader import SpecError, load_spec  # noqa: E402

import test_phase10_methodology as conformance  # noqa: E402

MANIFEST = PCCM_ROOT / "spec" / "workbook.yaml"


# ---------------------------------------------------------------------------
# the harness
# ---------------------------------------------------------------------------
def _raw() -> dict:
    return yaml.safe_load(MANIFEST.read_text(encoding="utf-8"))


def _section(raw: dict, key: str) -> dict:
    return next(s for s in raw["methodology"]["sections"] if s["key"] == key)


def _write(raw: dict) -> tuple:
    """Dump a damaged manifest and load it with the REAL loader.

    The loader is part of what is being tested: several mutations below are
    supposed to be refused before a control ever runs, and a harness that
    bypassed it would hide that.
    """
    directory = tempfile.mkdtemp()
    path = Path(directory) / "workbook.yaml"
    text = yaml.safe_dump(raw, sort_keys=False, allow_unicode=False, width=10000)
    path.write_text(text, encoding="utf-8")
    return load_spec(path), text


def _tests() -> list[str]:
    names = sorted(name for name in dir(conformance) if name.startswith("test_"))
    assert len(names) >= 50, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _install(damaged: dict):
    saved = dict(conformance._CACHE)
    conformance._CACHE.update(damaged)

    def restore() -> None:
        conformance._CACHE.clear()
        conformance._CACHE.update(saved)

    return restore


def _detects(expected: str, damaged: dict | None = None, **kwargs) -> None:
    """Damage what the battery reads, run all of it, and require a NAMED refuser.

    `damaged` takes cache keys that are not identifiers - the per-document ones
    are `("doc", name)` tuples - and `kwargs` is the ordinary spelling for the
    rest.
    """
    restore = _install({**(damaged or {}), **kwargs})
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _manifest_mutation(expected: str, mutate) -> None:
    """Damage the manifest, then require a NAMED control to refuse it."""
    raw = copy.deepcopy(_raw())
    mutate(raw)
    spec, text = _write(raw)
    _detects(expected, spec=spec, manifest_text=text)


def _refused_by_the_loader(mutate) -> str:
    raw = copy.deepcopy(_raw())
    mutate(raw)
    try:
        _write(raw)
    except SpecError as error:
        return str(error)
    raise AssertionError("the loader accepted a manifest it must refuse")


# ===========================================================================
# A. M5 - DISTRIBUTES, NEVER SCALES
# ===========================================================================
def test_01_profiling_made_to_scale_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M5")
        section["paragraphs"][0] = (
            "Profiling SCALES an amount across the project years, multiplying it "
            "by the weight of each year.")

    _manifest_mutation("test_17", mutate)


def test_02_dropping_the_does_not_scale_clause_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M5")
        section["paragraphs"][0] = (
            "Profiling DISTRIBUTES an amount across the project years.")

    _manifest_mutation("test_17", mutate)


def test_03_dropping_the_hundred_percent_rule_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M5")
        section["paragraphs"][1] = section["paragraphs"][1].replace(
            "must total 100%", "should usually total about 100%")

    _manifest_mutation("test_18", mutate)


def test_04_pointing_cost_profiling_at_the_wrong_register_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        term = _section(raw, "M5")["terms"][0]
        term["text"] = term["text"].replace("Cost Line ID", "Risk ID")

    _manifest_mutation("test_18", mutate)


# ===========================================================================
# B. M8 - THE TWO ANNUAL OBJECTS
# ===========================================================================
def test_10_the_selected_profile_made_not_to_sum_is_rejected() -> None:
    """THE MUTATION THE SPECIFICATION NAMED. It is wrong for the selected-Px
    profile, and it is the sentence a reader would otherwise carry away."""
    def mutate(raw: dict) -> None:
        term = _section(raw, "M8")["terms"][1]
        term["text"] = term["text"].replace(
            "It DOES sum to the selected total Px, subject only to "
            "floating-point round-off.",
            "The annual profile does not sum to the total Px.")

    _manifest_mutation("test_24", mutate)


def test_11_that_same_mutation_is_caught_as_a_forbidden_sentence() -> None:
    """AND BY THE OTHER CONTROL TOO. One says the right sentence is missing; the
    other says a wrong one is present. Both must hold."""
    def mutate(raw: dict) -> None:
        section = _section(raw, "M8")
        section["paragraphs"].append(
            "Note that the annual profile does not sum to total Px.")

    _manifest_mutation("test_26", mutate)


def test_12_the_ladder_made_to_sum_to_the_total_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        term = _section(raw, "M8")["terms"][0]
        term["text"] = term["text"].replace(
            "Adding a per-year P50 across the years does NOT give the reported "
            "total P50",
            "Adding a per-year P50 across the years gives the reported total P50")

    _manifest_mutation("test_23", mutate)


def test_13_dropping_the_different_iteration_reason_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        term = _section(raw, "M8")["terms"][0]
        term["text"] = term["text"].replace(
            ": each year's number comes from a different iteration", "")

    _manifest_mutation("test_23", mutate)


def test_14_corrupting_the_interpolation_identity_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        formula = _section(raw, "M8")["formula"]
        formula["text"] = ("Profile_Px(year) = (1 + f) x AnnualVector_lo(year) "
                           "+ f x AnnualVector_hi(year)")

    _manifest_mutation("test_25", mutate)


def test_15_removing_the_identity_altogether_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        _section(raw, "M8").pop("formula")

    _manifest_mutation("test_25", mutate)


def test_16_merging_the_two_annual_objects_into_one_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M8")
        first, second = section["terms"]
        section["terms"] = [{"term": "Annual results",
                             "text": f"{first['text']} {second['text']}"}]

    _manifest_mutation("test_22", mutate)


def test_17_dropping_the_same_interpolation_claim_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        term = _section(raw, "M8")["terms"][1]
        term["text"] = term["text"].replace(
            "using the SAME interpolation the total used",
            "using an interpolation between them")

    _manifest_mutation("test_24", mutate)


def test_18_a_reconciliation_claim_that_contradicts_the_sim_contract_is_rejected() -> None:
    """THE CROSS-CHECK, NOT THE SELF-CHECK. Methodology is measured against the
    accepted simulation contract, so wording that drifted from the engine fails
    even if the sheet is internally consistent."""
    def mutate(raw: dict) -> None:
        term = _section(raw, "M8")["terms"][0]
        term["text"] = term["text"].replace("does NOT give the reported total P50",
                                            "reconciles with the reported total P50")

    _manifest_mutation("test_27", mutate)


# ===========================================================================
# C. M10 - UNDEFINED IS NOT ZERO
# ===========================================================================
def test_20_reporting_zero_variance_as_rho_zero_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M10")
        section["paragraphs"][1] = (
            "A driver with no variance has no relationship with the total, so "
            "the model reports rho = 0 for it.")

    _manifest_mutation("test_32", mutate)


def test_21_calling_the_correlation_defined_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M10")
        section["paragraphs"][1] = section["paragraphs"][1].replace(
            "Rank correlation is UNDEFINED", "Rank correlation is zero")

    _manifest_mutation("test_32", mutate)


def test_22_paraphrasing_the_engines_own_status_label_is_rejected() -> None:
    """THE LABEL IS THE ENGINE'S. A sheet that invented a friendlier wording
    would describe a status the user will never actually see."""
    def mutate(raw: dict) -> None:
        section = _section(raw, "M10")
        section["paragraphs"][1] = section["paragraphs"][1].replace(
            "n/a - no variance", "not applicable")

    _manifest_mutation("test_33", mutate)


def test_23_letting_a_zero_variance_driver_into_the_tornado_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M10")
        section["paragraphs"][2] = section["paragraphs"][2].replace(
            "It is EXCLUDED from the Tornado ranking and from the tornado's input "
            "altogether", "It appears in the Tornado as a bar of zero length")

    _manifest_mutation("test_34", mutate)


# ===========================================================================
# D. M11 - FOUR STATES, AND REFUSED IS NOT ONE
# ===========================================================================
def test_30_a_fifth_state_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        _section(raw, "M11")["terms"].append(
            {"term": "REFUSED", "text": "The model declined to act."})

    _manifest_mutation("test_35", mutate)


def test_31_dropping_a_state_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M11")
        section["terms"] = [term for term in section["terms"]
                            if term["term"] != "NOT CALCULATED"]

    _manifest_mutation("test_35", mutate)


def test_32_renaming_a_state_is_rejected() -> None:
    """THE WORDS ARE THE ENGINE'S. A friendlier synonym on the sheet would not
    match the word the workbook actually displays."""
    def mutate(raw: dict) -> None:
        term = next(t for t in _section(raw, "M11")["terms"] if t["term"] == "STALE")
        term["term"] = "OUT OF DATE"

    _manifest_mutation("test_35", mutate)


def test_33_deleting_the_refused_paragraph_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        _section(raw, "M11").pop("closing")

    _manifest_mutation("test_36", mutate)


def test_34_calling_refused_a_persistent_state_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M11")
        section["closing"] = ("REFUSED is a fifth state the workbook can be left "
                              "in after a command declines to act.")

    _manifest_mutation("test_36", mutate)


def test_35_losing_the_live_versus_published_distinction_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M11")
        section["paragraphs"][0] = section["paragraphs"][0].replace(
            "reports the model's condition LIVE", "reports the last published result")

    _manifest_mutation("test_37", mutate)


# ===========================================================================
# E. THE EARLIER SECTIONS
# ===========================================================================
def test_40_reversing_the_three_point_ordering_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M2")
        section["paragraphs"][1] = "Maximum <= Most Likely <= Minimum"

    _manifest_mutation("test_11", mutate)


def test_41_letting_uniform_use_the_most_likely_value_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        term = next(t for t in _section(raw, "M2")["terms"] if t["term"] == "Uniform")
        term["text"] = ("Every value between the minimum and the maximum is equally "
                        "likely, and Uniform uses the most likely value as the centre "
                        "of that range.")

    _manifest_mutation("test_13", mutate)


def test_42_that_same_mutation_is_caught_as_an_implication() -> None:
    """AND BY THE NEGATIVE CONTROL. 'Do not imply that Uniform uses ML' cannot be
    proved by any positive assertion, so the wrong claim is searched for
    directly - here it is added somewhere else entirely."""
    def mutate(raw: dict) -> None:
        section = _section(raw, "M2")
        section["paragraphs"].append(
            "In practice Uniform uses the most likely value to centre the range.")

    _manifest_mutation("test_14", mutate)


def test_43_losing_the_type_7_percentile_method_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M7")
        section["paragraphs"][0] = section["paragraphs"][0].replace(
            "the TYPE 7 method - the same definition Excel's PERCENTILE.INC uses",
            "the nearest-rank method")

    _manifest_mutation("test_20", mutate)


def test_44_turning_contingency_into_a_cost_element_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M7")
        section["paragraphs"][-1] = (
            "Contingency is therefore a cost element in its own right, added to "
            "the deterministic basis and profiled with it.")

    _manifest_mutation("test_21", mutate)


def test_45_claiming_an_auto_run_is_reproducible_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        term = next(t for t in _section(raw, "M6")["terms"] if t["term"] == "AUTO seed")
        term["text"] = ("The model chooses a seed. Re-running an AUTO run "
                        "reproduces the same numbers.")

    _manifest_mutation("test_19", mutate)


def test_46_turning_correlation_into_causation_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M9")
        section["paragraphs"][2] = ("A high-ranking driver causes the total to move, "
                                    "and reducing it will reduce the total.")

    _manifest_mutation("test_31", mutate)


def test_47_ranking_by_signed_rho_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        term = next(t for t in _section(raw, "M9")["terms"] if t["term"] == "Ranking")
        term["text"] = "Drivers are ranked by rho, descending."

    _manifest_mutation("test_30", mutate)


def test_48_naming_a_vba_module_on_the_sheet_is_rejected() -> None:
    """NO MACHINE-ORIENTED NOISE. The audience is never a developer."""
    def mutate(raw: dict) -> None:
        section = _section(raw, "M11")
        section["paragraphs"].append(
            "The state words are derived by modCalcReport and published from there.")

    _manifest_mutation("test_41", mutate)


# ===========================================================================
# F. STRUCTURE - WHAT THE LOADER ITSELF MUST REFUSE
# ===========================================================================
def test_50_a_reordered_section_list_is_refused_by_the_loader() -> None:
    """REQUIRED CONTROL 2, from the other side. Order is not a convention here:
    a manifest whose sections are out of sequence never loads."""
    def mutate(raw: dict) -> None:
        sections = raw["methodology"]["sections"]
        sections[6], sections[7] = sections[7], sections[6]

    assert "must be numbered M1..M11 in order" in _refused_by_the_loader(mutate)


def test_51_a_missing_section_is_rejected() -> None:
    """AND THE LOADER IS NOT THE CONTROL HERE. M1..M10 is a perfectly well-formed
    sequence, so dropping the last section loads cleanly - which is exactly why
    the battery names the eleven it requires rather than counting what arrived.
    """
    def mutate(raw: dict) -> None:
        raw["methodology"]["sections"] = [
            section for section in raw["methodology"]["sections"]
            if section["key"] != "M11"]

    _manifest_mutation("test_02", mutate)


def test_51b_a_gap_in_the_numbering_is_refused_by_the_loader() -> None:
    """A section removed from the MIDDLE is a different thing: it breaks the
    sequence, and the loader refuses it before any control runs."""
    def mutate(raw: dict) -> None:
        raw["methodology"]["sections"] = [
            section for section in raw["methodology"]["sections"]
            if section["key"] != "M5"]

    assert "must be numbered M1..M10 in order" in _refused_by_the_loader(mutate)


def test_52_a_string_excel_would_evaluate_is_refused_by_the_loader() -> None:
    """FORMULAS ARE PRESENTATION TEXT. A line starting with `=` would stop being
    a sentence the moment the workbook opened."""
    def mutate(raw: dict) -> None:
        _section(raw, "M8")["formula"]["text"] = (
            "= (1 - f) * AnnualVector_lo + f * AnnualVector_hi")

    assert "reads as a formula" in _refused_by_the_loader(mutate)


def test_53_a_leading_minus_is_refused_too() -> None:
    def mutate(raw: dict) -> None:
        _section(raw, "M7")["paragraphs"][0] = (
            "- Contingency(Px) = Total(Px) - Deterministic basis")

    assert "reads as a formula" in _refused_by_the_loader(mutate)


def test_54_a_methodology_body_with_no_block_is_refused() -> None:
    def mutate(raw: dict) -> None:
        raw.pop("methodology")

    assert "carries no methodology block" in _refused_by_the_loader(mutate)


def test_55_a_second_term_with_the_same_name_is_refused() -> None:
    def mutate(raw: dict) -> None:
        section = _section(raw, "M11")
        section["terms"].append(dict(section["terms"][0]))

    assert "defines the same term twice" in _refused_by_the_loader(mutate)


def test_56_a_sheet_that_declares_blocks_beside_the_body_is_refused() -> None:
    def mutate(raw: dict) -> None:
        sheet = next(s for s in raw["sheets"] if s["name"] == "Methodology")
        sheet["blocks"] = [{"type": "note", "text": "Typed straight onto the sheet."}]

    assert "must not also declare blocks" in _refused_by_the_loader(mutate)


# ===========================================================================
# G. THE TWO VERSION AUTHORITIES
# ===========================================================================
def test_60_a_stale_version_file_is_rejected() -> None:
    _detects("test_60", version_file="0.5.0\n")


def test_61_a_stale_model_version_in_the_manifest_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        raw["model"]["model_version"] = "0.5.0"

    _manifest_mutation("test_60", mutate)


def test_62_a_stale_builder_version_is_rejected() -> None:
    _detects("test_61", builder_version="0.5.0")


def test_63_a_phase_string_instead_of_a_release_string_is_rejected() -> None:
    def mutate(raw: dict) -> None:
        raw["model"]["build_phase"] = "Phase 10 - Delivery (implementation)"

    _manifest_mutation("test_62", mutate)


def test_64_a_release_string_with_different_words_is_rejected() -> None:
    """THE JUSTIFIED VARIANT IS THE HYPHEN, NOT THE WORDING. A value that reads
    differently is a different stamp, however reasonable it sounds."""
    def mutate(raw: dict) -> None:
        raw["model"]["build_phase"] = "Production Release 1.0"

    _manifest_mutation("test_62", mutate)


def test_65_deriving_the_builder_version_from_the_model_version_is_rejected() -> None:
    """THE COLLAPSE THIS BATCH WAS MOST LIKELY TO COMMIT. Both authorities read
    1.0.0 today, so a line that made one read the other would pass every value
    comparison in the project - until the first builder-only change."""
    damaged = conformance._builder_source().replace(
        'BUILDER_VERSION = "1.0.0"',
        'BUILDER_VERSION = spec.model["model_version"]')
    _detects("test_63", builder_source=damaged)


def test_66_a_second_builder_version_literal_is_rejected() -> None:
    damaged = conformance._builder_source().replace(
        'BUILDER_VERSION = "1.0.0"',
        'BUILDER_VERSION = "1.0.0"\nFALLBACK_BUILDER_VERSION = "1.0.0"\n'
        'BUILDER_VERSION = FALLBACK_BUILDER_VERSION')
    _detects("test_63", builder_source=damaged)


def test_67_the_manifest_claiming_the_builder_version_is_rejected() -> None:
    """A SECOND OWNER, NOT A CONVENIENCE. If the manifest carried it too, the
    builder would have two answers and no rule for choosing."""
    def mutate(raw: dict) -> None:
        raw["model"]["builder_version"] = "1.0.0"

    _manifest_mutation("test_63", mutate)


def test_68_a_hardcoded_model_version_in_the_emitter_is_rejected() -> None:
    """THE GENERATED CONSTANT MUST STAY A PROJECTION. A literal in the emitter
    would agree with the manifest today and quietly stop agreeing on the day the
    model version next moves."""
    damaged = conformance._emitter_source().replace(
        'module.const("SIM_MODEL_VERSION", str(spec.model["model_version"]),',
        'module.const("SIM_MODEL_VERSION", "1.0.0",')
    _detects("test_64", emitter_source=damaged)


def test_69_a_rewritten_historical_record_is_rejected() -> None:
    """REQUIRED CONTROL 22, MUTATED. The easiest way to break this rule is a
    repository-wide search and replace, which would turn each closure record
    into a claim that an earlier phase shipped this release."""
    damaged = conformance._doc("phase5_gate_a_step3.md").replace("0.5.0", "1.0.0")
    _detects("test_65", {("doc", "phase5_gate_a_step3.md"): damaged})


# ===========================================================================
# H. THE RENDERER
# ===========================================================================
def test_70_a_sentence_hard_coded_into_the_renderer_is_rejected() -> None:
    """PREFER GENERATED CONTENT. A fallback literal is how a spec-driven sheet
    quietly becomes a typed one."""
    damaged = conformance._renderer_source() + (
        '\n_FALLBACK = "Profiling DISTRIBUTES an amount across the project years. '
        'It does NOT scale, multiply or inflate the amount."\n')
    _detects("test_05", renderer_source=damaged)


def test_71_a_renderer_that_stops_dispatching_is_rejected() -> None:
    damaged = conformance._builder_source().replace(
        'elif sheet_spec.body == "methodology":', 'elif False:')
    _detects("test_06", builder_source=damaged)


def test_72_a_renderer_that_writes_a_formula_is_rejected() -> None:
    damaged = conformance._renderer_source().replace(
        "cell.value = line.text", 'cell.value = "=" + line.text')
    _detects("test_42", renderer_source=damaged)


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
