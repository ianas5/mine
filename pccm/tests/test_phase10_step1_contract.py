#!/usr/bin/env python3
"""The Phase-10 Step-1 contract records decisions, and holds them.

WHAT A CONTRACT RECORD IS FOR. It is read by the person who implements from it,
and the whole point of settling a contract before code is that the code cannot
quietly settle it differently. So the decisions that were actually DECIDED - the
six commands, Workbook_Open, what Reset clears and preserves, what Repair may
not force, 1.0.0, the practical benchmark matrix, verified-versus-expected
support - are pinned by name here.

WHAT THESE CONTROLS ALSO REFUSE. The two wordings this round exists to correct:
the M8 annual-Px sentence that would contradict the accepted selected-profile
identity, and a distribution claim that turns "no Win32 Declare" into a support
statement. Both are checked against the specs they are about, so a record that
drifts from its own authorities fails rather than reading plausibly.

AND ONE MORE: that implementation has not started. A contract record which
quietly grew an implementation would be the easiest thing in this project to
miss.

Runs standalone or under pytest.
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
RECORD = PCCM_ROOT / "docs" / "phase10_step1_contract.md"
SPEC = PCCM_ROOT / "spec"
SRC = PCCM_ROOT / "src" / "vba"
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402


def _text() -> str:
    return RECORD.read_text(encoding="utf-8")


def _flat() -> str:
    """The record with its line wrapping and blockquote markers taken out.

    A prose claim must not depend on WHERE a paragraph happened to wrap. The
    marker is stripped only at the START of a line, because removing every
    "> " would take the arrow out of an expression like `a -> b`.
    """
    return re.sub(r"\s+", " ", re.sub(r"(?m)^>[ \t]?", "", _text()))


# ===========================================================================
# A. THE RECORD EXISTS, AND SAYS WHAT IT IS
# ===========================================================================
def test_01_the_record_declares_the_contract_settled() -> None:
    assert RECORD.is_file()
    text = _text()
    assert "PHASE 10 — CONTRACT SETTLED" in text
    assert "Implementation is NOT started by this record." in text
    assert "Phases 7, 8 and 9 are not reopened." in text


# THE CONTRACT COMMIT ITSELF. Every claim this record makes about "the tree" is
# a claim about THIS tree, and is asked of it rather than of whatever the
# working directory holds while a later step is being built.
CONTRACT_COMMIT = "6ab8f6a"


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=REPO_ROOT, check=True,
                          stdout=subprocess.PIPE, text=True).stdout


def test_02_implementation_had_not_started_when_the_contract_was_settled() -> None:
    """CONVERTED AT P10-2A, NOT DELETED, AND NOT LOOSENED.

    As first written this scanned TODAY's tree, which was right for exactly as
    long as nothing had been implemented - and it went red the moment P10-2A
    landed the protection owner it authorises. Deleting it would retire the
    claim; relaxing it to "some Phase-10 files may exist" would retire it more
    quietly.

    SO IT ASKS THE COMMIT INSTEAD, the way the Phase-8 closure record was
    corrected. The record says implementation was not started BY IT; the check
    is that the CONTRACT COMMIT'S OWN TREE carried none. That is the claim the
    sentence actually makes, it is immune to everything committed afterwards,
    and it is stricter in the way that matters: an implementation file smuggled
    into the contract commit still fails, and now it fails forever rather than
    only until implementation legitimately began.
    """
    listing = _git("ls-tree", "-r", "--name-only", CONTRACT_COMMIT, "pccm/").split()
    for name in ("modReset.bas", "modRepair.bas", "modProtection.bas",
                 "ThisWorkbook.vba", "ThisWorkbook.cls", "ThisWorkbook.bas",
                 "protection.py"):
        assert not [p for p in listing if p.endswith("/" + name)], (
            f"the contract commit {CONTRACT_COMMIT} already carried {name}")
    structure = _git("show", f"{CONTRACT_COMMIT}:pccm/spec/structure_contract.yaml")
    for endpoint in ("PCCM_ResetResults", "PCCM_RepairProfiling", "PCCM_Calculate\""):
        assert endpoint not in structure or "entry_point" not in structure.split(
            endpoint)[0][-40:], (
            f"{endpoint} was already bound at {CONTRACT_COMMIT}")
    assert structure.count("entry_point:") == 5, (
        f"the button table had already moved at {CONTRACT_COMMIT}")


def test_02a_this_batch_is_the_one_the_contract_authorises() -> None:
    """AND THE SEPARATE HALF: what exists now is what 6ab8f6a authorised, and
    only that. Reset and Repair are LATER steps; finding either here would mean
    a batch ran ahead of its authorisation."""
    # ADVANCED AT P10-2B AND COMPLETED AT P10-2C. The record authorises SIX
    # commands and they arrived one step at a time; the fence moved with the
    # steps that were actually authorised rather than being deleted once one of
    # them landed. All six are now bound, which is the whole of what 6ab8f6a
    # permits - so a SEVENTH is what this now refuses.
    structure = (SPEC / "structure_contract.yaml").read_text(encoding="utf-8")
    assert structure.count("entry_point:") == 11, (
        "the button table is not the five Phase-4 buttons plus the six "
        "operational commands the contract authorises")
    for command in ("PCCM_Calculate", "PCCM_RunSimulation", "PCCM_RunSensitivity",
                    "PCCM_RunAnnualStochastic", "PCCM_ResetResults",
                    "PCCM_RepairProfiling"):
        assert f'entry_point: "{command}"' in structure, command
    for authorised in ("modProtection.bas", "ThisWorkbook.vba"):
        assert (SRC / authorised).is_file(), f"{authorised} is authorised and absent"


# ===========================================================================
# B. THE SIX COMMANDS
# ===========================================================================
COMMANDS = (
    ("Calculate", "PCCM_Calculate"),
    ("Run Simulation", "PCCM_RunSimulation"),
    ("Run Sensitivity", "PCCM_RunSensitivity"),
    ("Run Annual Cash Flow", "PCCM_RunAnnualStochastic"),
    ("Reset Results", "PCCM_ResetResults"),
    ("Repair Profiling", "PCCM_RepairProfiling"),
)


@pytest.mark.parametrize("caption,entry", COMMANDS)
def test_10_every_authorised_command_is_recorded(caption: str, entry: str) -> None:
    flat = _flat()
    assert f"**{caption}**" in flat, f"the record does not name the command {caption}"
    assert entry in flat, f"the record does not name the entry point {entry}"


def test_11_the_four_existing_endpoints_really_exist() -> None:
    """FOUR OF THE SIX ARE NOT NEW, and the record says so. If one of them did
    not exist, "existing endpoint, newly bound" would be false."""
    source = "\n".join(p.read_text(encoding="utf-8") for p in sorted(SRC.glob("*.bas")))
    for entry in ("PCCM_Calculate", "PCCM_RunSimulation", "PCCM_RunSensitivity",
                  "PCCM_RunAnnualStochastic"):
        assert re.search(rf"^Public (?:Sub|Function) {entry}\b", source, re.M), entry
        assert "existing endpoint, newly bound" in _text()


def test_12_binding_a_button_changes_no_phase_7_semantics() -> None:
    """THE ONE THING A BUTTON MUST NOT DO. Sensitivity and the annual step are
    separate operations by an accepted Phase-7 decision; giving them a button is
    a delivery change, not a semantic one, and the record has to say so."""
    flat = _flat()
    assert "remain **explicit separate operations**" in flat
    assert "Neither runs as part of a simulation" in flat
    assert "No Phase-7 contract is reopened by giving an existing endpoint a button" in flat


# ===========================================================================
# C. WORKBOOK_OPEN
# ===========================================================================
def test_20_the_open_handler_scope_is_minimal_and_stated() -> None:
    flat = _flat()
    assert "first** `ThisWorkbook` event handler is authorised" in flat
    assert "UserInterfaceOnly:=True" in flat
    for forbidden in ("calculate, simulate, derive state, publish, run business logic",
                      "take a destructive action, or show a dialog on success"):
        assert forbidden in flat, forbidden
    assert "Protection configuration has ONE owner." in flat
    # AND THE ALTERNATIVE IS REJECTED IN TERMS, not left as an option.
    assert "explicitly rejected" in flat


def test_21_the_open_handler_is_failure_safe() -> None:
    flat = _flat()
    for state in ("ScreenUpdating", "Calculation", "EnableEvents"):
        assert state in flat, state
    assert "usable and unprotected rather than half-protected" in flat


# ===========================================================================
# D. RESET RESULTS
# ===========================================================================
_RESET_CLEARS = (
    "deterministic calculation publication and last-successful fingerprint",
    "calculation last-attempt presentation and history",
    "simulation publications, both banks, and the active publication identity",
    "annual distribution and profile publications and their stamps",
    "sensitivity publication",
)
_RESET_PRESERVES = (
    "permanent-ID counters",
    "run-id / nonce monotonic identity state",
    "seed history and identity fields required for non-collision",
    "applied timeline inputs",
    "build metadata",
)


@pytest.mark.parametrize("cleared", _RESET_CLEARS)
def test_30_reset_clears_every_publication_named(cleared: str) -> None:
    assert cleared in _flat(), f"Reset no longer records clearing: {cleared}"


@pytest.mark.parametrize("preserved", _RESET_PRESERVES)
def test_31_reset_preserves_every_identity_named(preserved: str) -> None:
    assert preserved in _flat(), f"Reset no longer records preserving: {preserved}"


def test_32_reset_derives_states_and_forces_none() -> None:
    """THE RULE THAT KEEPS RESET OUT OF THE STATE MACHINE. Forcing a state string
    would make Reset a second opinion about state, which section 13 forbids."""
    flat = _flat()
    assert "**No state string is forced to obtain these outcomes.**" in flat
    assert "They fall out of the existing owners once the publications are gone" in flat
    for outcome in ("`NOT CALCULATED`", "`INVALID`", "`NOT PRODUCED`"):
        assert outcome in flat, outcome


def test_33_the_input_comparison_is_not_a_byte_comparison() -> None:
    """THE ACCEPTANCE DEFINITION, AND THE ONE IT REFUSES. A binary comparison
    would fail on metadata that legitimately moved and prove nothing about the
    inputs the control exists to protect."""
    flat = _flat()
    assert "compares exactly before and after" in flat
    assert "**Not** a binary workbook-byte comparison." in flat


# ===========================================================================
# E. REPAIR PROFILING
# ===========================================================================
def test_40_the_repairable_and_ambiguous_classes_are_both_present() -> None:
    text = _text()
    assert text.count("| repairable |") == 4, "the repairable class list has moved"
    assert text.count("| **ambiguous — refuse** |") == 4, "the refuse class list has moved"
    assert "weights blank" in text and "unmade assumption" in text


def test_41_a_structural_repair_may_not_force_stale() -> None:
    """THE CLARIFICATION THIS ROUND EXISTS FOR. A repair that preserves the
    driver-to-weight mapping has changed no model input, and inventing a
    semantic change would be exactly the kind of silent business-logic move
    section 13 forbids."""
    flat = _flat()
    assert "**Do not force `STALE`.**" in flat
    assert "The existing fingerprint owner decides" in flat
    assert "the calculation fingerprint must remain identical" in flat
    assert "ordinary existing fingerprint and state rules apply" in flat


# ===========================================================================
# F. M8 - THE CORRECTED ANNUAL Px WORDING
# ===========================================================================
def test_50_m8_states_both_objects_and_does_not_conflate_them() -> None:
    flat = _flat()
    assert "Annual simulated distributions" in flat
    assert "The selected-Px annual profile" in flat
    assert "Profile_Px(y) = (1 - f) * AnnualVector_lo(y) + f * AnnualVector_hi(y)" in flat
    assert "**This profile DOES sum to the selected total Px**" in flat
    assert "degenerates to a single iteration's vector when `f = 0`" in flat


def test_51_m8_matches_the_accepted_sim_contract() -> None:
    """THE RECORD'S REASONING IS CHECKED AGAINST THE CONTRACT IT REASONS ABOUT.
    sim_contract already settles the two objects; a Methodology section that
    drifted from it would teach a user something the model does not do."""
    sim = re.sub(r"\s+", " ",
                 re.sub(r"(?m)^#[ \t]?", "",
                        (SPEC / "sim_contract.yaml").read_text(encoding="utf-8")))
    assert "Sum_y of a per-year Px does NOT equal the reported total Px" in sim
    assert "It DOES reconcile: Sum_y = reported Px" in sim
    assert "sums_to_total_percentile: false" in (SPEC / "sim_contract.yaml").read_text(
        encoding="utf-8")


def test_52_the_forbidden_sentence_is_forbidden_in_terms() -> None:
    """AND THE WORDING THAT WOULD CONTRADICT IT IS NAMED. A record that merely
    said the right thing could still be reworded into the wrong one; this says
    which sentence may not be written."""
    flat = _flat()
    assert 'Methodology must **not** say "the annual profile does not sum to the total Px"' in flat
    assert "contradicts the accepted selected-profile identity" in flat


# ===========================================================================
# G. RELEASE AND VERSION
# ===========================================================================
def test_60_the_release_is_1_0_0_and_the_authorities_stay_independent() -> None:
    flat = _flat()
    assert "→ **1.0.0**" in flat
    assert "remain INDEPENDENT authorities after release" in flat
    assert "A later builder-only change need not move the model version" in flat
    assert "No version literal is duplicated." in flat


def test_61_the_build_phase_finalisation_is_recorded_and_still_stale() -> None:
    """THE ITEM, AND THE PROOF IT IS STILL AN ITEM. When implementation lands
    this control's second half changes; until then the record must not claim a
    correction it has not made."""
    flat = _flat()
    assert "Phase-10 metadata finalisation item and reopens no earlier phase" in flat
    assert "Release 1.0 — Production" in flat
    spec = (SPEC / "workbook.yaml").read_text(encoding="utf-8")
    assert "Phase 5 - Calculation Workspace (Gate A: source)" in spec, (
        "the build phase was corrected; the record's account of it is out of date")
    assert 'model_version: "0.5.0"' in spec, (
        "the model version moved before implementation was authorised")


def test_62_the_source_revision_is_an_addition_not_an_existing_row() -> None:
    builder = (PCCM_ROOT / "builder" / "pccm_builder" / "workbook_builder.py").read_text(
        encoding="utf-8")
    assert "Source Revision" not in builder, (
        "the source revision already exists; the record calls it an addition")
    assert "**ADD**" in _text()


# ===========================================================================
# H. PERFORMANCE
# ===========================================================================
def test_70_the_matrix_is_practical_and_its_omissions_are_declared() -> None:
    flat = _flat()
    assert "not a Cartesian product" in flat
    assert "**not required**" in flat, "the large 100,000 omission is not recorded"
    assert "**maximum** for Simulation, Sensitivity and annual replay" in flat
    assert "may be omitted for Sensitivity" in flat
    assert "iteration-independent" in flat
    assert "never left to read as missing evidence" in flat


def test_71_the_runtime_record_and_the_warm_cold_rule_are_required() -> None:
    flat = _flat()
    for item in ("Excel version and bitness", "Windows version", "CPU", "RAM",
                 "local versus synced", "other workbooks are open"):
        assert item in flat, item
    assert "median of three subsequent runs" in flat
    assert "**not user-visible**" in flat


def test_72_no_threshold_precedes_the_baseline() -> None:
    flat = _flat()
    assert "No absolute pass/fail threshold is contracted before the baseline exists." in flat
    assert "**1.5×**" in flat and "investigate" in flat
    assert "**2×**" in flat and "blocks delivery" in flat


def test_73_the_phase_7_timings_are_quoted_as_historical_and_not_edited() -> None:
    flat = _flat()
    assert "the operator's report, explicitly not confirmed by the repository" in flat
    assert "quoted, never edited" in flat
    # AND THEY REALLY ARE STILL IN THE PHASE-7 RECORD, UNCHANGED.
    phase7 = (PCCM_ROOT / "docs" / "phase7_closure.md").read_text(encoding="utf-8")
    assert "20 drivers × 10k ≈ 6.4–7.6 s, 100 × 10k ≈ 31.4 s, 300 × 10k ≈ 105.4 s" in phase7


# ===========================================================================
# I. DISTRIBUTION - VERIFIED VERSUS EXPECTED
# ===========================================================================
def test_80_the_three_support_tiers_are_separate() -> None:
    flat = _flat()
    assert "REQUIRED / OFFICIALLY SUPPORTED" in flat
    assert "COMPATIBLE BY INSPECTION, NOT INDEPENDENTLY VALIDATED" in flat
    assert "UNSUPPORTED" in flat
    assert "VERIFIED** support separately from **EXPECTED" in flat


def test_81_no_support_claim_is_made_from_the_absence_of_a_declare() -> None:
    """THE OVERCLAIM THIS ROUND REFUSES. "No Win32 Declare" is an inspection
    result about pointers. It is not evidence that 32-bit Office runs this
    workbook, and the record must not let the two become one sentence."""
    flat = _flat()
    assert "not claimed as accepted unless it is actually tested" in flat
    assert "expected* to be compatible" in flat
    # THE INSPECTION IS TRUE: there is no Declare in any module.
    for path in sorted(SRC.glob("*.bas")):
        body = "\n".join(line for line in path.read_text(encoding="utf-8").splitlines()
                         if not line.strip().startswith("'"))
        assert not re.search(r"^\s*(?:Public |Private )?Declare\b", body, re.M), path.name


def test_82_the_minimum_excel_claim_is_about_formulas_only() -> None:
    """AND IT IS RE-DERIVABLE. The record names the newest functions the built
    workbook uses; if a later change introduced a modern one, the claim would be
    wrong and this fails."""
    flat = _flat()
    assert "no requirement beyond Excel 2007" in flat
    assert "an inspection result about formulas, not a support claim" in flat
    assert "the supported minimum is whatever acceptance actually runs" in flat

    workbook = PCCM_ROOT / "build" / "PCCM_stageA.xlsx"
    if not workbook.is_file():
        pytest.skip("Stage A has not been built into pccm/build")
    import openpyxl

    modern = {"XLOOKUP", "XMATCH", "IFS", "SWITCH", "TEXTJOIN", "CONCAT", "LET",
              "LAMBDA", "SEQUENCE", "FILTER", "SORTBY", "UNIQUE", "MAXIFS",
              "MINIFS", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER"}
    used: set[str] = set()
    for sheet in openpyxl.load_workbook(workbook).worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    used.update(re.findall(r"(?<![A-Z0-9._])([A-Z][A-Z0-9.]*)\(",
                                           cell.value))
    assert not (used & modern), (
        f"the workbook now uses a post-2007 function: {sorted(used & modern)}")


def test_83_mac_is_not_called_supported() -> None:
    flat = _flat()
    assert "Excel for Mac: unsupported / unvalidated" in flat
    assert "Excel Online; Excel mobile; LibreOffice" in flat


# ===========================================================================
# J. THE CLAIMS THAT MUST NOT BE OVERSTATED, AND THE GUARD
# ===========================================================================
def test_90_the_transactional_claim_is_scoped_not_global() -> None:
    flat = _flat()
    assert "is **not** made" in flat
    assert "every **new** Phase-10 destructive or structural command is transactionally safe" in flat
    assert "retain their already-accepted restoration contracts" in flat
    assert "no newly introduced partial-write path is permitted" in flat


_GUARDED = ("costing mathematics", "risk mathematics", "percentile semantics",
            "sensitivity semantics", "annual profile identity", "fingerprints",
            "existing state contracts", "Model Check taxonomy")


@pytest.mark.parametrize("guarded", _GUARDED)
def test_91_every_guarded_owner_is_named(guarded: str) -> None:
    assert guarded in _flat(), f"the no-new-business-logic guard has lost: {guarded}"


def test_92_the_shared_owners_are_declared_in_advance() -> None:
    flat = _flat()
    for owner in ("modAppState", "ThisWorkbook", "the profiling structure owner",
                  "structure_contract.yaml"):
        assert owner in flat, owner
    assert "stops implementation and is surfaced before proceeding" in flat


def test_93_the_acceptance_matrix_carries_the_two_new_rows() -> None:
    text = _text()
    assert "the six commands are reachable as buttons" in text
    assert "`Workbook_Open` failure is safe" in text
    assert "a structural-only repair moves no fingerprint" in text
    assert "are not re-run** unless implementation touches a shared owner" in _flat()


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
