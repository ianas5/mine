#!/usr/bin/env python3
"""The Phase-8 closure record says only things the repository can be asked.

WHAT A CLOSURE RECORD IS FOR. It is read later, by someone deciding what a
result rested on. Every number in it is therefore either RE-DERIVED here, or
labelled as evidence from a run this repository cannot reproduce - and the
difference is stated rather than left to the reader.

WHAT THESE CONTROLS REFUSE. A commit that does not exist. An authority quoted
for a step it did not produce. A Phase-7 authority quietly moved. A correction
described but never made, or made but never described. A supplemental proof
presented as if it were the main acceptance. And the claim that Phase 9 has
started when it has not.

Runs standalone or under pytest.
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
RECORD = PCCM_ROOT / "docs" / "phase8_closure.md"
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

# THE THREE STEP AUTHORITIES AND THE SUPPLEMENTAL ONE, as the record states them.
ACCEPTANCE = {
    "P8-1": ("35bd6ce", 160),
    "P8-2": ("7ff5dc7", 124),
    "P8-3": ("9e3c141", 189),
}
SUPPLEMENTAL = ("bfae0eb", 23)
FINAL_TREE = "bfae0eb"

# PHASE 7'S TWO AUTHORITIES. This record may not move them.
PHASE7_IMPLEMENTATION = "79d4c3e"
PHASE7_ACCEPTANCE = "ad78988"

# THE FOUR LATER CORRECTIONS, by commit subject. A hash inside the row it lands
# in cannot be written by that row, so the subject is what resolves.
CORRECTIONS = {
    "2e72ddf": "the Results state path is read-only",
    "0cfa10f": "a risk publishes its risk name",
    "94c6b37": "zero variance is undefined, not zero",
    "6781ea7": "the tornado input is the ranked population",
}


def _text() -> str:
    return RECORD.read_text(encoding="utf-8")


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=REPO_ROOT, check=True,
                          stdout=subprocess.PIPE, text=True).stdout


# ===========================================================================
# A. THE RECORD EXISTS AND SAYS WHAT IT IS
# ===========================================================================
def test_01_the_record_declares_phase_8_closed() -> None:
    text = _text()
    assert RECORD.is_file()
    assert "PHASE 8 — ACCEPTED / CLOSED" in text
    assert "No outstanding Phase-8 blocker." in text


def test_02_phase_9_is_not_started_and_says_so() -> None:
    text = _text()
    assert "Phase 9 is not started by this record." in text
    assert "**Phase 9 has not started.**" in text
    # AND THE TREE AGREES. A record claiming Phase 9 is unstarted while Phase-9
    # work existed would be the one claim nobody would check.
    for pattern in ("phase9", "phase_9"):
        found = [p.name for p in (PCCM_ROOT / "tests").glob(f"*{pattern}*")]
        found += [p.name for p in (PCCM_ROOT / "docs").glob(f"*{pattern}*")]
        found += [p.name for p in (PCCM_ROOT / "bootstrap" / "windows").glob(f"*{pattern}*")]
        assert not found, f"Phase-9 work exists in the tree: {found}"


# ===========================================================================
# B. EVERY QUOTED COMMIT IS REAL, AND IS WHAT THE RECORD SAYS IT IS
# ===========================================================================
@pytest.mark.parametrize("step", sorted(ACCEPTANCE))
def test_10_each_step_authority_exists_and_is_quoted(step: str) -> None:
    commit, checks = ACCEPTANCE[step]
    resolved = _git("rev-parse", "--verify", f"{commit}^{{commit}}").strip()
    assert resolved.startswith(commit), resolved
    text = _text()
    assert f"`{commit}`" in text, f"{step}'s authority is not quoted"
    assert str(checks) in text, f"{step}'s check count is not stated"
    assert step in text


def test_11_the_supplemental_authority_exists_and_is_separate() -> None:
    commit, checks = SUPPLEMENTAL
    assert _git("rev-parse", "--verify", f"{commit}^{{commit}}").strip().startswith(commit)
    text = _text()
    assert f"`{commit}`" in text
    assert f"**{checks} / {checks}**" in text
    # IT IS NOT PRESENTED AS THE MAIN ACCEPTANCE.
    assert "SUPPLEMENTAL PROOF" in text
    assert "not a rerun of P8-3" in text
    assert "neither extend nor\nreplace the 189" in text


def test_12_the_p8_3_main_evidence_is_not_displaced() -> None:
    """THE 189/189 RUN REMAINS THE ACCEPTANCE. A later tree is a later tree."""
    text = _text()
    assert "remains the 189/189 run at" in text
    assert "Nothing later re-establishes it and nothing later replaces it." in text
    assert f"`{FINAL_TREE}` is the tree the last Windows run was produced against" in text
    assert "It is **not** a\nreplacement for the evidence commits above" in text


# ===========================================================================
# C. THE LINEAGE CLAIMS ARE RE-DERIVED, NOT ASSERTED
# ===========================================================================
@pytest.mark.parametrize("step", sorted(ACCEPTANCE))
def test_20_each_quoted_production_lineage_is_what_git_reports(step: str) -> None:
    """THE RECORD PRINTS A COMMAND AND ITS OUTPUT. Both are run here."""
    commit, _checks = ACCEPTANCE[step]
    actual = _git("log", "--format=%h %s", "-1", commit,
                  "--", "pccm/src", "pccm/spec").strip()
    assert actual, f"{step}: no production commit resolves at {commit}"
    quoted = f"$ git log --oneline -1 {commit} -- pccm/src pccm/spec"
    text = _text()
    assert quoted in text, f"the record does not show the query for {step}"
    following = text[text.index(quoted) + len(quoted):].lstrip().splitlines()[0].strip()
    assert following == actual, (
        f"{step}: the record says {following!r}; git says {actual!r}")


def test_21_no_production_byte_moved_between_p8_3_and_the_final_tree() -> None:
    """THE CLAIM THAT MAKES `bfae0eb` SAFE TO QUOTE."""
    diff = _git("diff", "--name-status", ACCEPTANCE["P8-3"][0], FINAL_TREE,
                "--", "pccm/src", "pccm/spec").strip()
    assert diff == "", f"production moved after the P8-3 acceptance:\n{diff}"
    assert "$ git diff --name-status 9e3c141 bfae0eb -- pccm/src pccm/spec" in _text()
    assert "(no output)" in _text()


def test_22_the_phase_7_authorities_are_unmoved() -> None:
    """PHASE 8 CORRECTED PHASE-7-OWNED MODULES AND STILL MAY NOT MOVE THESE."""
    text = _text()
    for authority in (PHASE7_IMPLEMENTATION, PHASE7_ACCEPTANCE):
        assert f"`{authority}`" in text, authority
        assert _git("rev-parse", "--verify", f"{authority}^{{commit}}").strip().startswith(
            authority)
    assert "does not reopen Phase 7 and does not move either\nauthority" in text
    # AND THE PHASE-7 RECORD STILL NAMES THEM ITSELF.
    phase7 = (PCCM_ROOT / "docs" / "phase7_closure.md").read_text(encoding="utf-8")
    for authority in (PHASE7_IMPLEMENTATION, PHASE7_ACCEPTANCE):
        assert f"`{authority}`" in phase7, (
            f"the Phase-7 record no longer names {authority}")


# ===========================================================================
# D. THE FOUR CORRECTIONS ARE REAL, AND EACH ONE CHANGED SOMETHING
# ===========================================================================
@pytest.mark.parametrize("commit,subject", sorted(CORRECTIONS.items()))
def test_30_each_declared_correction_resolves_to_one_commit(
        commit: str, subject: str) -> None:
    found = _git("log", "--format=%h", f"--grep={subject}", "--fixed-strings",
                 FINAL_TREE).split()
    assert len(found) == 1, (
        f"{subject!r} resolves to {len(found)} commits, not one: {found}")
    assert found[0].startswith(commit), f"{subject!r} is {found[0]}, not {commit}"
    text = _text()
    assert f"`{commit}`" in text, f"the record does not quote {commit}"


def test_31_each_declared_correction_actually_changed_something() -> None:
    """A ROW DESCRIBING A CHANGE NOBODY MADE IS DECORATION."""
    for commit in CORRECTIONS:
        touched = _git("show", "--name-only", "--format=", commit).split()
        assert touched, f"{commit} touched nothing"
    # THE TWO pccm/src ONES, AND THE TWO THAT ARE NOT.
    for commit in ("2e72ddf", "0cfa10f"):
        touched = _git("show", "--name-only", "--format=", commit)
        assert "pccm/src/vba/" in touched, (
            f"{commit} is described as a production correction and touches no VBA")
    for commit in ("94c6b37", "6781ea7"):
        touched = _git("show", "--name-only", "--format=", commit)
        assert "pccm/builder/" in touched, (
            f"{commit} is described as a builder correction and touches no builder")
        assert "pccm/src/vba/" not in touched, (
            f"{commit} is described as changing no VBA, and it does")
    assert "change no VBA and no contract" in _text()


def test_32_the_risk_label_correction_is_described_as_it_is() -> None:
    """THE ONE-LINE CLAIM IS CHECKABLE."""
    removed = [line for line in
               _git("diff", "0cfa10f~1", "0cfa10f",
                    "--", "pccm/src/vba/modSimPostReport.bas").splitlines()
               if line.startswith("-") and not line.startswith("---")]
    assert removed == ["-    column = COL_RISK_REGISTER_DESCRIPTION"], removed
    text = _text()
    assert "COL_RISK_REGISTER_RISK_NAME" in text
    assert "**One line.**" in text
    assert "The cost-line branch is untouched" in text


def test_33_the_sweep_results_are_labelled_as_run_evidence() -> None:
    """TWO NUMBERS THIS REPOSITORY CANNOT REPRODUCE ON DEMAND - a full sweep is
    over an hour - so they are quoted with the exit code that made them
    trustworthy, not presented as re-derived."""
    text = _text()
    for count in ("5044 passed", "5062 passed"):
        assert count in text, count
    assert text.count("PYTEST_EXIT=0") == 2, (
        "a sweep result is quoted without the exit code that validates it")


# ===========================================================================
# E. THE STEP TABLE IS COMPLETE AND THE VERDICTS AGREE WITH THEMSELVES
# ===========================================================================
def test_40_every_step_is_closed_and_the_summary_agrees() -> None:
    text = _text()
    summary = text[text.index("P8-1   CLOSED"):]
    for step, (commit, checks) in ACCEPTANCE.items():
        assert f"{step}   CLOSED / ACCEPTED" in summary, step
        assert commit in summary, f"{step}: {commit} missing from the summary"
        assert str(checks) in summary, f"{step}: {checks} missing from the summary"
    assert "P8-Z   SUPPLEMENTAL PROOF" in summary
    assert SUPPLEMENTAL[0] in summary
    # THE SUPPLEMENTAL ONE IS NOT MARKED CLOSED/ACCEPTED AS A STEP.
    assert "P8-Z   CLOSED / ACCEPTED" not in summary


def test_41_the_four_corrections_are_all_in_the_correction_table() -> None:
    text = _text()
    table = text[text.index("## 3. Later Phase-8 integration corrections"):
                 text.index("## 4. ")]
    for label in ("Results read-only state integration",
                  "Risk-driver label ownership",
                  "Zero-variance Sensitivity presentation",
                  "Zero-variance tornado exclusion"):
        assert label in table, f"the correction table omits {label}"
    for commit in CORRECTIONS:
        assert f"`{commit}`" in table, f"{commit} is not in the correction table"


def test_42_the_p8_z_scope_is_stated_as_narrow() -> None:
    text = _text()
    section = text[text.index("## 5. P8-Z"):text.index("## 6. ")]
    assert "supplemental, and narrow on purpose" in section
    assert "re-tests none of that suite's surface" in section
    assert "nothing fabricates a status" in section
    # AND WHAT IT ACTUALLY PROVED IS LISTED, not summarised away.
    for proof in ("4 ranked of 5 drivers", "n/a - no variance",
                  "**blank, not zero**", "exactly once",
                  "fabricate no categories"):
        assert proof in section, f"the P8-Z section omits {proof!r}"
