#!/usr/bin/env python3
"""P10-3: the Methodology sheet, and the first production release stamp.

TWO THINGS THAT LOOK UNRELATED AND ARE NOT. The Methodology sheet is where a
user is told what the model does; the Build Metadata block at the bottom of it
is where the same user is told which model they are looking at. Both are
presentation, both are generated, and both are wrong in exactly the same way if
anybody types a value into them.

WHAT THESE CONTROLS DEFEND.

  the WORDING     Eleven sections were specified sentence by sentence because
                  several of them correct a misunderstanding that a reasonable
                  person would otherwise arrive at on their own - most of all
                  M8, where the two annual objects sum differently, and M10,
                  where "undefined" and "zero" are opposite claims. So the
                  wording is asserted here, and asserted AGAINST THE CONTRACTS
                  IT DESCRIBES wherever an accepted contract states the same
                  thing. A sheet that drifted from the engine would teach a
                  user something the model does not do.

  the SILENCE     The sheet must remain a reader. It holds no formula, no
                  macro, no defined name and no state word that anything reads
                  back. Nothing in the workbook points at it.

  the TWO VERSIONS Model version and builder version are INDEPENDENT
                  authorities that both happen to read 1.0.0 for this release.
                  The easiest possible mistake here is to make one derive from
                  the other while they agree, and discover it the first time
                  they need to differ.

  the PAST        The release stamp moved forward. The records of earlier
                  phases did not, and must not: a closure record that silently
                  claimed to have shipped 1.0.0 would be a false rewrite of
                  what happened.

Runs standalone or under pytest.
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
SPEC = PCCM_ROOT / "spec"
SRC = PCCM_ROOT / "src" / "vba"
BUILD = PCCM_ROOT / "build"
DOCS = PCCM_ROOT / "docs"
BUILDER = PCCM_ROOT / "builder"
sys.path.insert(0, str(BUILDER))

import pytest  # noqa: E402
import yaml  # noqa: E402

from pccm_builder.methodology import plan_methodology  # noqa: E402
from pccm_builder.spec_loader import load_spec  # noqa: E402
from pccm_builder.workbook_builder import BUILDER_VERSION, BuildMetadata  # noqa: E402

MANIFEST = SPEC / "workbook.yaml"
WORKBOOK = BUILD / "PCCM_stageA.xlsx"
SHEET = "Methodology"

# THE ACCEPTED TREE THIS BATCH IS MEASURED AGAINST. P10-2C is the last accepted
# commit; every implementation owner must still be byte-identical to it, because
# this batch is documentation, presentation and a version stamp.
ACCEPTED = "4f1a57d"

SECTION_KEYS = tuple(f"M{number}" for number in range(1, 12))

# THE ONE METADATA ROW THAT IS NOT A DECLARED VALUE. Every other row restates an
# authority a control can read; this one records WHEN the artifact was produced,
# so a control that compared it against a freshly computed one would be
# comparing two clocks. What is asserted about it is its shape and its position.
TIMESTAMP_LABEL = "Build Timestamp (UTC)"
TIMESTAMP_SHAPE = re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} UTC$")

# The four persistent state words, and where they are owned. Methodology is
# checked against this file rather than against a list retyped here.
STATE_OWNER = SPEC / "calc_contract.yaml"


# READ-ONLY ANSWERS THAT COST REAL TIME. Loading the built workbook, parsing the
# manifest and shelling out to git are the expensive things these controls do,
# and none of those answers can change while the battery runs. Memoised
# SEPARATELY from `_CACHE`, which is the mutation battery's injection point:
# nothing memoised here is ever damaged, and keeping the two apart is what stops
# a memo from hiding a mutation.
_MEMO: dict = {}


def _git(*args: str) -> str:
    if args not in _MEMO:
        _MEMO[args] = subprocess.run(["git", *args], cwd=REPO_ROOT, check=True,
                                     stdout=subprocess.PIPE, text=True).stdout
    return _MEMO[args]


# THE ONE INJECTION POINT. Every accessor below reads through this cache, so the
# mutation battery can replace what the WHOLE battery sees rather than what one
# control happens to re-read from disk. Empty in normal use: nothing is cached
# between calls, and each accessor falls through to the file it is about.
_CACHE: dict = {}


def _spec():
    if "spec" in _CACHE:
        return _CACHE["spec"]
    if "spec" not in _MEMO:
        _MEMO["spec"] = load_spec(MANIFEST)
    return _MEMO["spec"]


def _manifest_text() -> str:
    if "manifest_text" in _CACHE:
        return _CACHE["manifest_text"]
    return MANIFEST.read_text(encoding="utf-8")


def _version_file() -> str:
    if "version_file" in _CACHE:
        return _CACHE["version_file"]
    return (PCCM_ROOT / "VERSION").read_text(encoding="utf-8")


def _builder_version() -> str:
    if "builder_version" in _CACHE:
        return _CACHE["builder_version"]
    return BUILDER_VERSION


def _builder_source() -> str:
    if "builder_source" in _CACHE:
        return _CACHE["builder_source"]
    return (BUILDER / "pccm_builder" / "workbook_builder.py").read_text(encoding="utf-8")


def _renderer_source() -> str:
    if "renderer_source" in _CACHE:
        return _CACHE["renderer_source"]
    return (BUILDER / "pccm_builder" / "methodology.py").read_text(encoding="utf-8")


def _emitter_source() -> str:
    if "emitter_source" in _CACHE:
        return _CACHE["emitter_source"]
    return (BUILDER / "pccm_builder" / "sim_emit.py").read_text(encoding="utf-8")


def _doc(name: str) -> str:
    if ("doc", name) in _CACHE:
        return _CACHE[("doc", name)]
    return (DOCS / name).read_text(encoding="utf-8")


def _block() -> dict:
    return _spec().methodology


def _section(key: str) -> dict:
    return next(s for s in _block()["sections"] if s["key"] == key)


def _section_strings(key: str) -> list[str]:
    """Every string a section puts on the sheet, in the order it appears."""
    section = _section(key)
    strings = [section["title"], section["summary"], *section["paragraphs"]]
    for term in section["terms"]:
        strings += [term["term"], term["text"]]
    formula = section.get("formula")
    if formula is not None:
        strings += [formula["caption"], formula["text"], formula["note"]]
    if section.get("closing"):
        strings.append(section["closing"])
    return strings


def _text(key: str) -> str:
    """One section's prose, flattened. Case is preserved: several of the
    requirements are about a word being SHOUTED, and lowering the text would
    quietly satisfy the claim it was written to make."""
    return " ".join(_section_strings(key))


def _all_text() -> str:
    block = _block()
    strings = [block[key] for key in ("heading", "purpose", "note",
                                      "metadata_heading", "metadata_note", "closing")]
    for key in SECTION_KEYS:
        strings += _section_strings(key)
    return " ".join(strings)


def _metadata() -> BuildMetadata:
    """The build stamp the renderer is handed, resolved from the same
    authorities the builder resolves it from."""
    if "contracts" not in _MEMO:
        from pccm_builder.contract_loader import load_contract
        from pccm_builder.driver_loader import load_driver_contract
        from pccm_builder.structure_loader import load_structure_contract

        _MEMO["contracts"] = (
            load_contract(SPEC / "input_contract.yaml"),
            load_driver_contract(SPEC / "driver_contract.yaml"),
            load_structure_contract(SPEC / "structure_contract.yaml"),
        )
    return BuildMetadata.create(_spec(), *_MEMO["contracts"])


def _plan():
    return plan_methodology(_spec(), _metadata())


def _accepted_text(path: str) -> str:
    return _git("show", f"{ACCEPTED}:pccm/{path}")


def _workbook():
    import openpyxl

    if "workbook" not in _MEMO:
        _MEMO["workbook"] = openpyxl.load_workbook(WORKBOOK)
    return _MEMO["workbook"]


needs_build = pytest.mark.skipif(
    not WORKBOOK.is_file(), reason="Stage A has not been built into pccm/build")


# ===========================================================================
# A. THE SHEET IS GENERATED, AND ITS CONTENT IS DECLARED
# ===========================================================================
def test_01_the_sheet_is_rendered_from_a_declared_block_not_from_typed_cells() -> None:
    """REQUIRED CONTROL 1, first half. The manifest carries the content and the
    sheet carries a body rather than a list of generic sections - which is what
    makes exactly one author responsible for those rows."""
    sheet = _spec().sheet(SHEET)
    assert sheet.body == "methodology", sheet.body
    assert sheet.blocks == [], "the sheet declares blocks beside a rendered body"
    assert _block(), "the manifest carries no methodology block"
    assert _block()["sheet"] == SHEET


def test_02_all_eleven_sections_are_present() -> None:
    """REQUIRED CONTROL 1. M1-M11, no more and no fewer."""
    keys = [section["key"] for section in _block()["sections"]]
    assert keys == list(SECTION_KEYS), keys


def test_03_the_section_order_is_deterministic() -> None:
    """REQUIRED CONTROL 2. Order is the manifest's, the plan preserves it, and
    the loader refuses any numbering that is not M1..Mn in sequence - so the
    order cannot be an accident of dictionary iteration anywhere."""
    planned = [line.key for line in _plan() if line.kind == "section"]
    assert planned == list(SECTION_KEYS), planned
    rows = [line.row for line in _plan() if line.kind == "section"]
    assert rows == sorted(rows), rows


def test_04_every_section_carries_a_title_a_summary_and_prose() -> None:
    for key in SECTION_KEYS:
        section = _section(key)
        assert section["title"].strip()
        assert section["summary"].strip()
        assert section["paragraphs"], f"{key} has no prose"


def test_05_the_renderer_owns_layout_and_not_one_sentence() -> None:
    """PREFER GENERATED CONTENT. The module that writes the sheet is checked for
    the thing that would make the manifest a decoration: a sentence of its own.

    Every string the renderer can put in a cell comes from the block it was
    handed. What is asserted here is the converse - that no sentence from the
    manifest is ALSO present in the renderer, which is how a hard-coded fallback
    usually arrives.
    """
    source = _renderer_source()
    for key in SECTION_KEYS:
        for sentence in _section_strings(key):
            assert sentence not in source, f"{key}: the renderer restates {sentence[:50]!r}"
    for key in ("heading", "purpose", "note", "closing"):
        assert _block()[key] not in source


def test_06_the_builder_dispatches_the_body_to_that_renderer() -> None:
    builder = _builder_source()
    assert 'elif sheet_spec.body == "methodology":' in builder
    assert "render_methodology(worksheet, spec, styles, metadata)" in builder


# ===========================================================================
# B. THE WORDING. EVERY REQUIREMENT THAT WAS SPECIFIED AS A SENTENCE
# ===========================================================================
def test_10_m1_states_the_two_registers_and_the_reporting_currency() -> None:
    """M1. A cost line is included in full; a risk's severity IS the amount and
    its probability decides whether it is incurred; everything is SAR."""
    text = _text("M1")
    assert "Quantity multiplied by unit cost" in text
    assert "no probability of occurrence" in text
    assert "Probability multiplied by severity" in text
    assert "the severity IS the amount" in text
    assert "SAR" in text


def test_11_m2_states_the_required_ordering() -> None:
    """REQUIRED CONTROL 3."""
    assert "Minimum <= Most Likely <= Maximum" in _text("M2")


def test_12_m2_names_the_three_families() -> None:
    terms = [term["term"] for term in _section("M2")["terms"]]
    for family in ("Triangular", "PERT", "Uniform"):
        assert family in terms, terms


def test_13_uniform_does_not_use_the_most_likely_value() -> None:
    """REQUIRED CONTROL 4. Stated on the Uniform term itself, so a reader who
    reads only that term still learns it."""
    uniform = next(term for term in _section("M2")["terms"] if term["term"] == "Uniform")
    assert "MOST LIKELY VALUE IS NOT USED" in uniform["text"]
    assert "does not affect a Uniform sample" in uniform["text"]


def test_14_nothing_says_uniform_uses_the_most_likely_value() -> None:
    """AND THE OPPOSITE IS NOWHERE. The requirement was 'do not imply that
    Uniform uses ML', which no positive assertion can prove."""
    flat = re.sub(r"\s+", " ", _all_text()).lower()
    for wrong in ("uniform uses the most likely",
                  "uniform sampling uses the most likely",
                  "uniform draws from the most likely"):
        assert wrong not in flat, wrong


def test_15_m3_states_probability_times_severity() -> None:
    text = _text("M3")
    assert "probability multiplied by its severity" in text
    assert "either happens at full severity or does not happen at all" in text


def test_16_m4_covers_currency_escalation_and_present_value() -> None:
    terms = [term["term"] for term in _section("M4")["terms"]]
    assert terms == ["Currency", "Escalation", "Present value"], terms
    assert "All three rates are INPUTS" in _text("M4")


def test_17_m5_says_profiling_distributes_and_does_not_scale() -> None:
    """REQUIRED CONTROL 5, and the sentence that carries the whole section."""
    text = _text("M5")
    assert ("Profiling DISTRIBUTES an amount across the project years. "
            "It does NOT scale, multiply or inflate the amount.") in text


def test_18_m5_separates_the_two_grids_by_the_ids_they_follow() -> None:
    terms = {term["term"]: term["text"] for term in _section("M5")["terms"]}
    assert set(terms) == {"Cost Profiling", "Risk Profiling"}, sorted(terms)
    assert "Cost Line ID" in terms["Cost Profiling"]
    assert "Risk ID" in terms["Risk Profiling"]
    assert "must total 100%" in _text("M5")


def test_19_m6_separates_a_fixed_seed_from_an_auto_one() -> None:
    """AND CLAIMS NO REPRODUCIBILITY FOR AUTO. The requirement was explicit: an
    AUTO run must not be described as deterministically repeatable."""
    terms = {term["term"]: term["text"] for term in _section("M6")["terms"]}
    assert set(terms) == {"FIXED seed", "AUTO seed"}, sorted(terms)
    assert "reproduce the same run exactly" in terms["FIXED seed"]
    assert "NOT reproducible" in terms["AUTO seed"]
    assert "a similar distribution, not the same numbers" in terms["AUTO seed"]


def test_20_m7_identifies_the_type_7_percentile_method() -> None:
    """REQUIRED CONTROL 6."""
    text = _text("M7")
    assert "TYPE 7 method" in text
    assert "PERCENTILE.INC" in text
    assert "P50" in text and "P80" in text


def test_21_m7_describes_contingency_as_a_measurement() -> None:
    text = _text("M7")
    assert "Contingency(Px) = Total(Px) - Deterministic basis" in text
    assert "Contingency is therefore a MEASUREMENT, not a cost element." in text
    assert "Nothing in the model adds it, allocates it or profiles it" in text


# --- M8, the section this batch exists for -------------------------------
def test_22_m8_carries_the_two_annual_objects_as_two_separate_things() -> None:
    """REQUIRED CONTROL 7."""
    terms = [term["term"] for term in _section("M8")["terms"]]
    assert terms == ["1. Annual percentile ladder", "2. Selected-Px annual profile"], terms
    assert "They answer different questions and they do not add up the same way" \
        in _text("M8")


def test_23_the_ladder_is_not_required_to_sum_to_the_total() -> None:
    """REQUIRED CONTROL 8, and the reason - each year's number comes from a
    different iteration - is given rather than asserted."""
    ladder = _section("M8")["terms"][0]["text"]
    assert "percentiled INDEPENDENTLY" in ladder
    assert ("Adding a per-year P50 across the years does NOT give the reported "
            "total P50") in ladder
    assert "NOT a profile of any single project" in ladder
    assert "each year's number comes from a different iteration" in ladder


def test_24_the_selected_px_profile_does_sum_to_the_selected_total() -> None:
    """REQUIRED CONTROL 9."""
    profile = _section("M8")["terms"][1]["text"]
    assert ("It DOES sum to the selected total Px, subject only to "
            "floating-point round-off.") in profile
    assert "SAME interpolation the total used" in profile


def test_25_m8_carries_the_lo_hi_f_interpolation_identity() -> None:
    """REQUIRED CONTROL 10. Shown as a formula because this is the one thing a
    reader cannot be told in words without ambiguity."""
    formula = _section("M8")["formula"]
    assert formula["text"] == (
        "Profile_Px(year) = (1 - f) x AnnualVector_lo(year) + f x AnnualVector_hi(year)")
    assert "interpolated between the sorted simulated outcomes lo and hi" in formula["caption"]
    assert "When f = 0" in formula["note"]
    assert "summing the profile over the years returns the selected total Px" in formula["note"]


def test_26_the_forbidden_annual_sentence_appears_nowhere() -> None:
    """REQUIRED CONTROL 11, in the only form that can actually hold: the claim
    is not that some sentence is present, it is that a WRONG one is absent
    everywhere on the sheet.

    The variants below are the ways the same false statement is normally
    written. It is false for the selected-Px profile, which does sum.
    """
    flat = re.sub(r"\s+", " ", _all_text()).lower()
    for wrong in (
        "the annual profile does not sum to the total px",
        "the annual profile does not sum to total px",
        "the annual profile will not sum to the total px",
        "the selected-px annual profile does not sum",
        "the profile does not sum to the selected total px",
    ):
        assert wrong not in flat, f"Methodology states {wrong!r}, which is wrong"


def test_27_m8_agrees_with_the_accepted_simulation_contract() -> None:
    """AND THE CONTRACT IT DESCRIBES SAYS THE SAME. sim_contract.yaml already
    settled the two objects; Methodology is checked against it rather than
    against itself, so a sheet that drifted from the engine fails here."""
    sim = re.sub(r"\s+", " ", re.sub(r"(?m)^#[ \t]*", "",
                 (SPEC / "sim_contract.yaml").read_text(encoding="utf-8")))
    assert "Sum_y of a per-year Px does NOT equal the reported total Px" in sim
    assert "It DOES reconcile: Sum_y = reported Px" in sim
    ladder, profile = (term["text"] for term in _section("M8")["terms"])
    assert "does NOT give the reported total P50" in ladder
    assert "It DOES sum to the selected total Px" in profile


# --- M9, M10 --------------------------------------------------------------
def test_30_m9_names_spearman_and_ranks_by_absolute_rho() -> None:
    """REQUIRED CONTROL 12."""
    text = _text("M9")
    assert "SPEARMAN RANK CORRELATION" in text
    assert "ranked by the ABSOLUTE value of rho" in text
    terms = [term["term"] for term in _section("M9")["terms"]]
    assert terms == ["Sign", "Magnitude", "Ranking"], terms


def test_31_m9_does_not_turn_correlation_into_causation() -> None:
    text = _text("M9")
    assert "ASSOCIATION, not causation" in text
    flat = re.sub(r"\s+", " ", text).lower()
    for wrong in ("causes the total", "driving the total up", "because that driver"):
        assert wrong not in flat, wrong


def test_32_m10_reports_zero_variance_as_undefined_and_not_as_zero() -> None:
    """REQUIRED CONTROL 13. The distinction is spelled out, because "no
    relationship" and "nothing to measure" are different answers."""
    text = _text("M10")
    assert "Rank correlation is UNDEFINED" in text
    assert "n/a - no variance" in text
    assert "it is NOT reported as rho = 0" in text
    assert "A rho of zero says the model looked and found no relationship" in text


def test_33_the_no_variance_label_is_the_one_the_engine_publishes() -> None:
    """AND IT IS THE ENGINE'S OWN LABEL, not a paraphrase of it."""
    sim = yaml.safe_load((SPEC / "sim_contract.yaml").read_text(encoding="utf-8"))
    label = _find_status_label(sim)
    assert label == "n/a - no variance", label
    assert label in _text("M10")


def _find_status_label(node) -> str | None:
    if isinstance(node, dict):
        if "status_label" in node and isinstance(node["status_label"], str):
            return node["status_label"]
        for value in node.values():
            found = _find_status_label(value)
            if found is not None:
                return found
    elif isinstance(node, list):
        for value in node:
            found = _find_status_label(value)
            if found is not None:
                return found
    return None


def test_34_a_zero_variance_driver_is_excluded_from_the_tornado() -> None:
    """REQUIRED CONTROL 14, with the reason: a zero-length bar beside measured
    ones would read as a measurement."""
    text = _text("M10")
    assert "EXCLUDED from the Tornado ranking and from the tornado's input altogether" in text
    assert "stays visible in the Sensitivity table as a diagnostic row" in text


# --- M11 ------------------------------------------------------------------
def test_35_m11_carries_exactly_the_four_persistent_state_words() -> None:
    """REQUIRED CONTROL 15, checked against the contract that owns the
    vocabulary rather than against a list retyped here."""
    owner = yaml.safe_load(STATE_OWNER.read_text(encoding="utf-8"))["state_labels"]
    terms = [term["term"] for term in _section("M11")["terms"]]
    assert sorted(terms) == sorted(owner["derived_status"]), terms
    assert set(terms) == {"CURRENT", "STALE", "INVALID", "NOT CALCULATED"}


def test_36_refused_is_an_attempt_outcome_and_not_a_fifth_state() -> None:
    """REQUIRED CONTROL 16. It is said in the section's closing line, and it is
    NOT one of the terms - which is the structural half of the same claim."""
    section = _section("M11")
    assert "REFUSED is not a fifth state." in section["closing"]
    assert "It is the outcome of one attempt" in section["closing"]
    assert "REFUSED" not in [term["term"] for term in section["terms"]]
    owner = yaml.safe_load(STATE_OWNER.read_text(encoding="utf-8"))["state_labels"]
    assert "REFUSED" in owner["attempt_result"]
    assert "REFUSED" not in owner["derived_status"]


def test_37_m11_separates_the_live_condition_from_the_published_record() -> None:
    text = _text("M11")
    assert "reports the model's condition LIVE" in text
    assert "report what was last PUBLISHED" in text
    assert "they answer different questions" in text


# ===========================================================================
# C. THE SHEET STAYS A READER
# ===========================================================================
def test_40_no_methodology_string_can_be_read_as_a_formula() -> None:
    """REQUIRED CONTROL 17, at the source. The two identities on the sheet are
    sentences; the loader refuses anything Excel would evaluate."""
    block = _block()
    strings = [block[key] for key in ("heading", "purpose", "note", "metadata_heading",
                                      "metadata_note", "closing")]
    for key in SECTION_KEYS:
        strings += _section_strings(key)
    for text in strings:
        assert text.lstrip()[:1] not in ("=", "+", "-", "@"), text[:60]
    assert _spec().sheet(SHEET).subtitle.lstrip()[:1] not in ("=", "+", "-", "@")


def test_41_the_sheet_names_no_module_no_procedure_and_no_bas_file() -> None:
    """NO MACHINE-ORIENTED NOISE. The audience is never a developer, so the VBA
    module inventory is read from the structure contract and every declared name
    is required to be absent from the sheet."""
    structure = yaml.safe_load((SPEC / "structure_contract.yaml").read_text(encoding="utf-8"))
    flat = _all_text()
    for module in structure["vba"]["modules"]:
        assert module["name"] not in flat, f"the sheet names {module['name']}"
    for noise in (".bas", "Public Sub", "Public Function", "End Sub", "PCCM_",
                  "openpyxl", "workbook.yaml"):
        assert noise not in flat, f"the sheet carries {noise!r}"


def test_42_the_renderer_writes_no_formula_into_any_cell() -> None:
    source = _renderer_source()
    assert '= "="' not in source
    assert 'f"={' not in source
    assert "number_format" not in source, (
        "the methodology renderer is formatting values; it writes prose only")


@needs_build
def test_43_the_built_sheet_holds_no_formula_and_no_validation() -> None:
    """REQUIRED CONTROL 17, in the artifact. Text only, and nothing attached to
    it that could act."""
    sheet = _workbook()[SHEET]
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                assert not cell.value.startswith("="), f"{cell.coordinate}: {cell.value[:40]}"
            assert cell.data_type != "f", f"{cell.coordinate} holds a formula"
    assert not sheet.data_validations.dataValidation, "the sheet carries data validation"
    assert not sheet.conditional_formatting._cf_rules, "the sheet carries conditional formatting"
    assert not getattr(sheet, "tables", {}), "the sheet carries a table"


@needs_build
def test_44_nothing_in_the_workbook_points_at_the_methodology_sheet() -> None:
    """IT IS READ BY NOBODY. A formula anywhere that referenced this sheet would
    make a sentence into an input."""
    workbook = _workbook()
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    assert SHEET not in cell.value, (
                        f"{worksheet.title}!{cell.coordinate} reads Methodology")
    for name, defined in workbook.defined_names.items():
        assert SHEET not in str(defined.value), f"defined name {name} points at Methodology"


def test_45_no_vba_module_reads_the_sheet() -> None:
    for path in sorted(SRC.glob("*.bas")) + sorted(SRC.glob("*.vba")):
        code = "\n".join(line for line in path.read_text(encoding="utf-8").splitlines()
                         if not line.strip().startswith("'"))
        assert "Methodology" not in code, f"{path.name} references the Methodology sheet"
        assert "shMethodology" not in code, f"{path.name} references shMethodology"


# ===========================================================================
# D. PRESENTATION
# ===========================================================================
def test_50_the_sheet_declares_a_reading_measure_and_freezes_its_header() -> None:
    sheet = _spec().sheet(SHEET)
    assert sheet.freeze_panes == "A6"
    assert sheet.show_gridlines is False
    assert sheet.subtitle == "How PCCM works - a concise in-workbook reference"
    layout = _spec().presentation["layout"]
    assert sheet.column_widths[layout["label_column"]] >= 20, sheet.column_widths
    assert sheet.column_widths[layout["value_column"]] >= 80, sheet.column_widths


def test_51_the_plan_leaves_a_blank_row_between_sections() -> None:
    """NOT A WALL OF TEXT. Every section ends on a blank line, so the eye has
    somewhere to rest between eleven of them."""
    lines = _plan()
    by_row = {line.row: line for line in lines}
    for key in SECTION_KEYS:
        first = next(line.row for line in lines if line.kind == "section" and line.key == key)
        assert by_row[first - 1].kind == "blank", f"{key} has no blank line above it"
    assert lines[-1].kind == "closing"


def test_52_the_plan_writes_into_two_columns_and_no_others() -> None:
    """THE LEFT RAIL AND THE PROSE COLUMN. A section tag, a defined term or a
    metadata label hangs in one; everything else runs in the other."""
    labelled = {line.kind for line in _plan() if line.label is not None}
    assert labelled == {"section", "term", "metadata_heading", "metadata_value"}, labelled
    for line in _plan():
        if line.kind == "blank":
            assert line.label is None and line.text is None


def test_53_the_sheet_is_a_concise_reference_not_a_manual() -> None:
    """THE USER MANUAL COMES LATER, and this sheet must not pre-empt it. The
    cap is a bound on length, not a target: it exists so that adding the manual
    to this sheet fails rather than passes quietly."""
    written = [line for line in _plan() if line.kind != "blank"]
    assert len(written) <= 140, f"{len(written)} written lines"
    words = len(_all_text().split())
    assert words <= 3000, f"{words} words"


@needs_build
def test_54_every_string_on_the_built_sheet_was_declared_somewhere() -> None:
    """NOTHING WAS TYPED INTO A CELL. Every string in the artifact is the
    manifest's, the sheet header's, or a build metadata label or value."""
    block = _block()
    metadata = _metadata()
    allowed = {block[key] for key in ("heading", "purpose", "note", "metadata_heading",
                                      "metadata_note", "closing")}
    for key in SECTION_KEYS:
        allowed.update(_section_strings(key))
        allowed.add(key)
    sheet_spec = _spec().sheet(SHEET)
    allowed.update({sheet_spec.title, sheet_spec.subtitle})
    for label, value in metadata.as_rows():
        allowed.add(label)
        if label != TIMESTAMP_LABEL:
            allowed.add(value)

    found = [cell.value for row in _workbook()[SHEET].iter_rows() for cell in row
             if isinstance(cell.value, str) and cell.value.strip()]
    assert found, "the Methodology sheet is empty"
    undeclared = sorted(set(found) - allowed)
    assert len(undeclared) == 1 and TIMESTAMP_SHAPE.match(undeclared[0]), undeclared


@needs_build
def test_55_the_built_sheet_matches_the_plan_row_for_row() -> None:
    """AND IT LANDED WHERE THE PLAN SAID. The projection a Windows run reads is
    the same object the renderer walked, so this proves both at once."""
    sheet = _workbook()[SHEET]
    layout = _spec().presentation["layout"]
    label_column, text_column = layout["label_column"], layout["value_column"]
    for line in _plan():
        if line.label is not None:
            assert sheet[f"{label_column}{line.row}"].value == line.label, line
        if line.label == TIMESTAMP_LABEL:
            assert TIMESTAMP_SHAPE.match(sheet[f"{text_column}{line.row}"].value)
        elif line.text is not None:
            assert sheet[f"{text_column}{line.row}"].value == line.text, line
        if line.kind == "blank":
            assert sheet[f"{label_column}{line.row}"].value is None
            assert sheet[f"{text_column}{line.row}"].value is None


# ===========================================================================
# E. THE RELEASE STAMP
# ===========================================================================
def test_60_the_model_version_authority_reads_the_release() -> None:
    """REQUIRED CONTROL 18."""
    assert _spec().model["model_version"] == "1.0.0"
    assert _version_file().strip() == "1.0.0"


def test_61_the_builder_version_authority_reads_the_release() -> None:
    """REQUIRED CONTROL 19."""
    assert _builder_version() == "1.0.0"


def test_62_the_build_phase_authority_is_the_release_string() -> None:
    """REQUIRED CONTROL 20, INCLUDING THE JUSTIFIED VARIANT. The Phase-10
    contract prefers `Release 1.0 — Production` with an em dash, written in
    prose. `workbook.yaml` is ASCII throughout, and a build stamp that renders
    differently depending on encoding is a poor build stamp - so the settled
    value is the same words with the manifest's own hyphen."""
    build_phase = _spec().model["build_phase"]
    assert build_phase == "Release 1.0 - Production", build_phase
    assert build_phase.replace(" - ", " ").split() == ["Release", "1.0", "Production"]
    assert "Phase" not in build_phase, "the artifact still stamps itself with a phase"
    assert _manifest_text().isascii()
    record = _doc("phase10_step1_contract.md")
    assert "Release 1.0 — Production" in record, (
        "the contract's own wording is no longer in the record")


def test_63_the_two_version_authorities_are_declared_independently() -> None:
    """REQUIRED CONTROL 21, and the trap this batch was most likely to fall
    into. They read the same value for this release. Nothing derives one from
    the other, and no third place restates either.
    """
    builder = _builder_source()
    assert builder.count("BUILDER_VERSION = ") == 1
    assert re.search(r'^BUILDER_VERSION = "\d+\.\d+\.\d+"$', builder, re.M)
    assert "BUILDER_VERSION = spec" not in builder
    assert 'BUILDER_VERSION = str(' not in builder

    declarations = [line for line in _manifest_text().splitlines()
                    if ("BUILDER_VERSION" in line or "builder_version" in line)
                    and not line.lstrip().startswith("#")]
    assert declarations == [], declarations

    # AND THE BUILD STAMP KEEPS THEM APART. Two rows, two owners, two values -
    # even now that the values agree.
    rows = dict(_metadata().as_rows())
    assert rows["PCCM Model Version"] == _spec().model["model_version"]
    assert rows["Builder Version"] == _builder_version()
    assert "model_version" not in [f.name for f in _builder_version_sources()]


def _builder_version_sources():
    import dataclasses

    return [field for field in dataclasses.fields(BuildMetadata)
            if field.name == "builder_version"]


def test_64_the_model_version_is_projected_not_restated() -> None:
    """The model version reaches the VBA as a generated constant, from the same
    manifest key. A second literal would be a second authority."""
    emit = _emitter_source()
    assert 'module.const("SIM_MODEL_VERSION", str(spec.model["model_version"])' in emit
    assert '"1.0.0"' not in emit and '"0.5.0"' not in emit


def test_65_historical_records_still_record_their_own_versions() -> None:
    """REQUIRED CONTROL 22. NO FALSE REWRITE OF PROJECT HISTORY.

    These documents record what earlier phases shipped. `0.5.0` is a true
    statement about Phase 5 and stays a true statement about Phase 5; a
    search-and-replace across the repository would have turned each of them into
    a claim that Phase 5 delivered the 1.0.0 release.
    """
    for name, expected in (
        ("phase5_gate_a_step3.md", "`VERSION` | 0.4.0 → 0.5.0"),
        ("phase5_gate_a_step4.md", "the workbook remains model version 0.5.0"),
        ("phase6_step5.md", '"model_version": "0.5.0"'),
    ):
        text = _doc(name)
        assert expected in text, f"{name} lost its historical version record"


def test_66_no_historical_document_was_touched_at_all() -> None:
    """AND NOT ONE BYTE OF THEM MOVED. The control above names three; this one
    refuses a change to any document in the repository's record of the past."""
    # MODIFIED OR DELETED, NOT ADDED. The claim is that the record of the past
    # is not rewritten; a NEW record of a new event is how the project moves
    # forward, and forbidding one would make the control an obstacle rather than
    # a guard. `--name-status` separates the two.
    touched = [line for line in _git("diff", "--name-status", ACCEPTED, "--",
                                     "pccm/docs").splitlines() if line.strip()]
    changed = [line for line in touched if not line.startswith("A\t")]
    assert changed == [], f"historical records changed: {changed}"


@needs_build
def test_67_the_workbook_visible_metadata_agrees_with_the_authorities() -> None:
    """REQUIRED CONTROL 23. The block a user actually reads is compared with the
    files that own each value - not with the projection, which is generated from
    the same place and would agree with itself."""
    sheet = _workbook()[SHEET]
    layout = _spec().presentation["layout"]
    label_column, text_column = layout["label_column"], layout["value_column"]
    shown = {}
    for line in _plan():
        if line.kind == "metadata_value":
            assert sheet[f"{label_column}{line.row}"].value == line.label
            shown[line.label] = sheet[f"{text_column}{line.row}"].value

    assert shown["PCCM Model Version"] == \
        yaml.safe_load(_manifest_text())["model"]["model_version"]
    assert shown["Build Phase"] == \
        yaml.safe_load(_manifest_text())["model"]["build_phase"]
    assert shown["Builder Version"] == _builder_version()
    assert shown["PCCM Model Version"] == _version_file().strip()
    assert TIMESTAMP_SHAPE.match(shown[TIMESTAMP_LABEL]), shown[TIMESTAMP_LABEL]
    assert set(shown) == {label for label, _ in _metadata().as_rows()}


@needs_build
def test_68_the_document_properties_carry_the_release_too() -> None:
    properties = _workbook().properties
    assert properties.category == "Release 1.0 - Production", properties.category
    assert "1.0.0" in properties.description


# ===========================================================================
# F. NOTHING ELSE MOVED
# ===========================================================================
def test_70_no_vba_source_changed_since_the_accepted_tree() -> None:
    """REQUIRED CONTROL 24, first half. This batch is documentation,
    presentation and a version stamp; not one line of the engine belongs in
    it."""
    changed = [line for line in _git("diff", "--name-only", ACCEPTED, "--",
                                     "pccm/src").splitlines() if line.strip()]
    # P10-RP. The runtime protection reconciliation is a LATER batch with its own
    # authorisation, and this control's claim - that the methodology batch
    # touched no engine line - still holds. It is now proved by reversal: taking
    # the declared structural window back out reproduces the accepted bytes.
    import sys as _sys
    _sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_structural_window import (DECLARED_STRUCTURAL_WINDOW_CHANGES,
                                       strip_structural_window)
    declared = {f"pccm/src/vba/{name}" for name in DECLARED_STRUCTURAL_WINDOW_CHANGES}
    for path in sorted(set(changed) & declared):
        name = path.rsplit("/", 1)[1]
        current = strip_structural_window(
            name, (PCCM_ROOT / "src" / "vba" / name).read_bytes().decode("utf-8"))
        accepted = _git("show", f"{ACCEPTED}:{path}")
        assert current.replace("\r\n", "\n") == accepted.replace("\r\n", "\n"), path
    changed = [path for path in changed if path not in declared]
    assert changed == [], f"VBA source changed: {changed}"


def test_71_no_implementation_owner_in_the_builder_changed() -> None:
    """REQUIRED CONTROL 24, second half. The calculation, the simulation, the
    fingerprint, the state derivation and the profiling semantics all live in
    named modules, and none of them is a documentation file."""
    changed = {Path(line).name for line in
               _git("diff", "--name-only", ACCEPTED, "--", "pccm/builder").splitlines()
               if line.strip()}
    allowed = {
        # P10-3's own additions and the two files that dispatch to them.
        "methodology.py", "spec_loader.py", "workbook_builder.py", "styling.py",
        "__init__.py", "build_stage_a.py",
        # P10-4A. DECLARED, NOT EXEMPTED. The performance benchmark plan is a
        # later batch's addition to the same tree, and this control is measured
        # from P10-2C rather than from a moving head - so a file added after
        # P10-3 has to be NAMED here or the control would forbid the project
        # from continuing. What the control still refuses is unchanged: an edit
        # to any calculation, simulation, fingerprint, state or profiling owner.
        # `benchmark.py` is none of those - it emits a list of what to measure
        # and computes nothing - and its own battery proves that separately.
        "benchmark.py",
    }
    forbidden = sorted(changed - allowed)
    assert forbidden == [], f"implementation modules changed: {forbidden}"
    # AND THE TWO SHARED FILES CHANGED ONLY WHERE THEY HAD TO. `workbook_builder`
    # gained a dispatch arm and a version literal; nothing else in it moved.
    diff = _git("diff", "-U0", ACCEPTED, "--",
                "pccm/builder/pccm_builder/workbook_builder.py")
    removed = [line for line in diff.splitlines()
               if line.startswith("-") and not line.startswith("---")]
    assert removed == ['-BUILDER_VERSION = "0.5.0"'], removed


def test_72_no_contract_specification_changed_except_the_manifest() -> None:
    changed = {Path(line).name for line in
               _git("diff", "--name-only", ACCEPTED, "--", "pccm/spec").splitlines()
               if line.strip()}
    assert changed <= {"workbook.yaml"}, sorted(changed)


def test_73_the_manifest_changed_only_where_this_batch_was_authorised_to() -> None:
    """AND INSIDE THE MANIFEST, ONLY THREE THINGS MOVED: the two version keys,
    the Methodology sheet's body, and the new methodology block. Every other
    sheet, the phase-6 shell, the phase-9 shell and all presentation tokens
    except the two P10-3 additions must be identical to the accepted tree.
    """
    accepted = yaml.safe_load(_accepted_text("spec/workbook.yaml"))
    current = yaml.safe_load(_manifest_text())

    assert set(current) - set(accepted) == {"methodology"}
    assert set(accepted) - set(current) == set()

    for key in set(accepted) - {"model", "sheets"}:
        assert current[key] == accepted[key], f"{key} changed"

    assert {k: v for k, v in current["model"].items() if k not in
            ("model_version", "build_phase")} == \
        {k: v for k, v in accepted["model"].items() if k not in
         ("model_version", "build_phase")}

    moved = [(a["name"], c["name"]) for a, c in zip(accepted["sheets"], current["sheets"])
             if a != c]
    assert moved == [(SHEET, SHEET)], moved


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
