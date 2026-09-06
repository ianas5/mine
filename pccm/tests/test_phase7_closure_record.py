#!/usr/bin/env python3
"""The Phase-7 closure settlement, checked against the repository it settles.

A CLOSURE RECORD THAT NOBODY CHECKS IS A PRESS RELEASE. This project has
recorded every earlier phase's settlement in `docs/` and pinned the load-bearing
claims with controls, and Phase 7's record is held to the same standard: every
number, digest and commit it states is re-derived here from the tree and from
git, and a claim that drifts fails.

WHAT THESE CONTROLS DELIBERATELY DO NOT DO. They cannot confirm a Windows
result - no PowerShell and no Excel runs here - so they check that the record
ATTRIBUTES those results honestly (operator-reported, not re-derived) rather
than checking the results themselves. The same applies to the historical P7-2
through P7-5 numbers: the controls require them to be presented as historical,
because the repository cannot re-establish them.
"""
from __future__ import annotations

import hashlib
import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
RECORD = PCCM_ROOT / "docs" / "phase7_closure.md"
WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
TESTS = PCCM_ROOT / "tests"

import pytest  # noqa: E402

# THE TWO AUTHORITIES THE RECORD SEPARATES. Named here so a control can fail if
# the record starts quoting one where it means the other.
IMPLEMENTATION_AUTHORITY = "79d4c3e"
ACCEPTANCE_HEAD = "ad78988"

RUNNERS = {
    "W1": ("phase7_w1_smoke.ps1", "164113f"),
    "W2": ("phase7_w2_many_drivers.ps1", "fb1e96c"),
    "W3": ("phase7_w3_long_years.ps1", "cc78015"),
    "W4": ("phase7_w4_base_simulation.ps1", "3874496"),
    "W5": ("phase7_w5_annual_success.ps1", "df1e34d"),
    "W6": ("phase7_w6_selector_move.ps1", "e504b7b"),
    "W7": ("phase7_w7_bank_cycle.ps1", "7de7b01"),
    "W8": ("phase7_w8_refusal.ps1", "ad78988"),
}

FROZEN_HARNESS = "phase7_acceptance_scenarios.ps1"
FROZEN_SHA256 = "9744d9b7c1b4ebbc94ae48db54dd2ffb74af9a554ee43d8da8b1e06efe68c8ed"

_CACHE: dict = {}


def _text() -> str:
    return RECORD.read_text(encoding="utf-8")


def _git(*args: str) -> str:
    """git, or a skip. A tarball has no history and cannot answer these."""
    try:
        out = subprocess.run(("git", "-C", str(REPO_ROOT)) + args,
                             capture_output=True, text=True, timeout=120)
    except (OSError, subprocess.SubprocessError) as exc:  # pragma: no cover
        pytest.skip(f"git is unavailable here: {exc}")
    if out.returncode != 0:
        pytest.skip(f"git could not answer {' '.join(args)}: {out.stderr.strip()}")
    return out.stdout


def _collected(target: str) -> int:
    """How many tests pytest actually collects. The shell never sees the glob -
    a pattern handed to a subprocess arrives unexpanded and collects nothing,
    which would have made this control pass by collecting zero."""
    if target not in _CACHE:
        if any(ch in target for ch in "*?["):
            paths = sorted(str(p.relative_to(PCCM_ROOT)) for p in TESTS.glob(
                target.split("/", 1)[1]))
            assert paths, f"{target} matched no file"
        else:
            paths = [target]
        out = subprocess.run(
            [sys.executable, "-m", "pytest", *paths, "--collect-only", "-q"],
            cwd=PCCM_ROOT, capture_output=True, text=True, timeout=900)
        match = re.search(r"^(\d+) tests collected", out.stdout, re.M)
        assert match, f"could not collect {target}:\n{out.stdout[-2000:]}"
        _CACHE[target] = int(match.group(1))
    return _CACHE[target]


# ===========================================================================
# A. THE RECORD EXISTS AND SAYS WHAT IT DECIDED
# ===========================================================================

def test_01_the_record_exists_and_declares_the_verdict() -> None:
    assert RECORD.exists(), "Phase 7 has no closure record"
    text = _text()
    assert "PHASE 7 — ACCEPTED / CLOSED" in text
    assert "No outstanding Phase-7 runtime blocker." in text
    assert "no further phase-7 windows scenario is required" in text.lower()


def test_02_it_does_not_claim_the_project_is_finished() -> None:
    """PHASE 7 IS NOT PCCM. A closure record that read as a delivery notice
    would be the most expensive sentence in the repository."""
    text = _text()
    assert "This is not project completion." in text
    for remaining in ("Phase 8", "Phase 9", "Phase 10",
                      "final user documentation", "red-team"):
        assert remaining in text, f"the record does not say {remaining!r} remains"
    lowered = text.lower()
    for forbidden in ("pccm is complete", "project is complete",
                      "ready for delivery", "delivery accepted"):
        assert forbidden not in lowered, f"the record claims {forbidden!r}"


def test_03_phase_6_runtime_authority_is_kept_historical() -> None:
    text = _text()
    assert "historical Phase-6 runtime authority remains historical only" in text
    assert "supersedes it for Phase-7 functionality" in text


# ===========================================================================
# B. THE TWO AUTHORITIES, AGAINST GIT
# ===========================================================================

def test_10_the_implementation_authority_is_the_last_commit_touching_src_or_spec() -> None:
    """THE CLAIM THE WHOLE SETTLEMENT RESTS ON, re-derived rather than repeated."""
    out = _git("log", "--format=%h", "-1", "--", "pccm/src", "pccm/spec").strip()
    assert out.startswith(IMPLEMENTATION_AUTHORITY), (
        f"the last commit touching pccm/src or pccm/spec is {out!r}, but the "
        f"record names {IMPLEMENTATION_AUTHORITY}")
    assert IMPLEMENTATION_AUTHORITY in _text()


def test_11_no_production_or_spec_byte_moved_between_the_two_authorities() -> None:
    diff = _git("diff", "--stat", IMPLEMENTATION_AUTHORITY, ACCEPTANCE_HEAD,
                "--", "pccm/src", "pccm/spec")
    assert diff.strip() == "", (
        "pccm/src or pccm/spec changed between the implementation authority and "
        f"the acceptance head:\n{diff}")


def test_12_every_commit_since_the_authority_is_evidence_only() -> None:
    """A commit that touched production would make the acceptance head a second
    implementation authority, which is exactly what the record denies."""
    names = _git("log", "--name-only", "--format=",
                 f"{IMPLEMENTATION_AUTHORITY}..{ACCEPTANCE_HEAD}")
    allowed = ("pccm/bootstrap/", "pccm/builder/", "pccm/tests/", "pccm/docs/")
    offenders = sorted({
        line.strip() for line in names.splitlines()
        if line.strip() and not line.strip().startswith(allowed)})
    assert not offenders, f"non-evidence paths changed after the authority: {offenders}"


def test_13_the_acceptance_head_is_named_and_not_confused_with_production() -> None:
    text = _text()
    assert ACCEPTANCE_HEAD in text
    assert "introduced **no** production VBA" in text
    # The record must place them in different rows of the authority table.
    table = text.split("## 1.")[1].split("## 2.")[0]
    assert IMPLEMENTATION_AUTHORITY in table and ACCEPTANCE_HEAD in table


# ===========================================================================
# C. THE EIGHT RUNNERS
# ===========================================================================

def test_20_every_recorded_runner_exists_with_the_recorded_digest_and_size() -> None:
    text = _text()
    for tag, (name, _commit) in RUNNERS.items():
        path = WINDOWS / name
        assert path.exists(), f"{tag}'s runner {name} is missing"
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        lines = path.read_text(encoding="utf-8").count("\n")
        row = [l for l in text.splitlines() if l.startswith(f"| {tag} |")]
        assert len(row) == 1, f"the evidence table has {len(row)} rows for {tag}"
        assert f"`{name}` ({lines})" in row[0], (
            f"{tag}: the record's line count disagrees with the file ({lines})")
        assert f"`{digest[:16]}`" in row[0], (
            f"{tag}: the record's digest disagrees with the file ({digest[:16]})")


def test_21_every_recorded_commit_exists_and_belongs_to_the_scenario() -> None:
    """THE COLUMN RECORDS WHERE EACH SCENARIO WAS ACCEPTED, which is not always
    where its runner was written: W3's runner landed in ba0949e and was accepted
    at the fixture correction cc78015, which touched the corpus and the controls
    instead. So the control requires the commit to touch something belonging to
    that scenario, and the record footnotes the one case where they differ."""
    for tag, (name, commit) in RUNNERS.items():
        touched = _git("show", "--name-only", "--format=", commit)
        owned = (f"pccm/bootstrap/windows/{name}",
                 f"pccm/tests/test_phase7_{tag.lower()}")
        assert any(line.startswith(owned) for line in touched.splitlines()
                   if line.strip()) or any(o in touched for o in owned), (
            f"{tag}: commit {commit} touched nothing belonging to the scenario:\n{touched}")
        if f"pccm/bootstrap/windows/{name}" not in touched:
            assert f"`{commit}`†" in _text(), (
                f"{tag}: the record does not footnote that {commit} is not the "
                "commit that wrote the runner")


def test_22_the_abandoned_harness_is_frozen_and_unreachable() -> None:
    frozen = WINDOWS / FROZEN_HARNESS
    assert hashlib.sha256(frozen.read_bytes()).hexdigest() == FROZEN_SHA256, (
        "the abandoned acceptance harness was modified; it is frozen as history")
    assert FROZEN_SHA256 in _text()
    # NAMING IT IS NOT REACHING IT. W1's header explains that it replaces the
    # abandoned harness, which is exactly the sentence a reader needs; what must
    # not exist is a dot-source or an invocation of it.
    for name, _commit in RUNNERS.values():
        body = (WINDOWS / name).read_text(encoding="utf-8")
        for reached in (f". (Join-Path $scriptDir '{FROZEN_HARNESS}')",
                        f"& (Join-Path $scriptDir '{FROZEN_HARNESS}')",
                        f"Import-Module {FROZEN_HARNESS}"):
            assert reached not in body, f"{name} executes the abandoned harness"


# ===========================================================================
# D. THE NUMBERS: RE-ESTABLISHED VERSUS HISTORICAL
# ===========================================================================

def test_30_the_re_established_static_numbers_are_true_today() -> None:
    text = _text()
    suites = len(list(TESTS.glob("test_*.py")))
    assert f"**{suites} files**" in text, f"the tree has {suites} test files"
    total = _collected("tests")
    assert f"**{total:,}**" in text, f"pytest collects {total:,} tests"
    phase7 = _collected("tests/test_phase7*.py")
    phase7_suites = len(list(TESTS.glob("test_phase7*.py")))
    assert f"**{phase7_suites} suites, {phase7} tests**" in text
    w_tests = sum(sum(1 for line in (TESTS / f"test_phase7_{tag.lower()}"
                                     f"{'_smoke_source' if tag == 'W1' else '_source'}.py"
                                     ).read_text(encoding="utf-8").splitlines()
                      if line.startswith("def test"))
                  for tag in RUNNERS)
    assert f"**{w_tests} tests**" in text, f"the W-suites hold {w_tests} tests"


def test_31_the_unverifiable_numbers_are_labelled_historical() -> None:
    """THE HONESTY CONTROL. None of these is recorded anywhere in the tree, so
    the record must present them as reported rather than confirmed - and must do
    it in the section that says so, not in the re-established one."""
    text = _text()
    established, _, historical = text.partition(
        "### 3.2 Historical — reported, not re-establishable from committed evidence")
    assert historical, "the record has no historical section"
    historical = historical.split("## 4.")[0]
    for number in ("3,741", "3,780", "3,909", "4,290", "105.4", "31.4", "3.638e-12"):
        assert number in historical, f"{number} is not in the historical section"
        assert number not in established, (
            f"{number} is presented as re-established, and the repository "
            "cannot re-establish it")
    assert "not** confirmed by" in historical or "not confirmed by" in historical


def test_32_the_repository_really_cannot_re_establish_them() -> None:
    """The claim in §3.2 is itself checked: if one of these numbers were in the
    tree after all, the record would be understating its own evidence."""
    haystacks = []
    for folder in (PCCM_ROOT / "docs", PCCM_ROOT / "readiness"):
        for path in folder.rglob("*.md"):
            if path == RECORD:
                continue
            haystacks.append(path.read_text(encoding="utf-8", errors="replace"))
    blob = "\n".join(haystacks)
    for number in ("3,741", "3,780", "3,909", "3.638e-12"):
        assert number not in blob, (
            f"{number} IS recorded in the repository; §3.2 understates the evidence")


def test_33_the_two_run_caveat_on_the_full_sweep_is_stated() -> None:
    text = _text()
    assert "No single sweep has covered all" in text
    assert "across two runs" in text


# ===========================================================================
# E. PROVENANCE THE RECORD MUST NOT BLUR
# ===========================================================================

def test_40_sensitivity_is_not_attributed_to_w1_through_w8() -> None:
    """W1-W8 RAN NO SENSITIVITY, and every one of their authorisations excluded
    it. Folding it into their evidence would be the single easiest way for this
    settlement to overclaim."""
    text = _text()
    assert "deliberately ran no sensitivity" in text
    # INVOKING IT IS THE CLAIM, NOT MENTIONING IT. W1 requires the procedure to
    # EXIST in the compiled project - a public-surface check the record now
    # states explicitly - and never runs it. A control that banned the name
    # would have forced that honest sentence out of the record.
    assert "no W-runner invokes `PCCM_RunSensitivity`" in text
    assert "Existing is not running." in text
    for name, _commit in RUNNERS.values():
        body = (WINDOWS / name).read_text(encoding="utf-8")
        for invocation in ("Run('PCCM_RunSensitivity')",
                           "Operation 'PCCM_RunSensitivity'",
                           "Endpoint 'PCCM_RunSensitivity'",
                           'Run("PCCM_RunSensitivity")'):
            assert invocation not in body, (
                f"{name} invokes PCCM_RunSensitivity after all")


def test_41_the_sensitivity_evidence_the_record_does_cite_is_present() -> None:
    timing = (WINDOWS / "phase7_timing_scenarios.ps1").read_text(encoding="utf-8")
    assert "PCCM_RunSensitivity" in timing
    for count in ("DriverCount = 20", "DriverCount = 100", "DriverCount = 300"):
        assert count in timing, f"the timing harness no longer defines {count}"
    assert "Iterations = 10000" in timing
    for suite in ("test_phase7_sim_sensitivity.py",
                  "test_phase7_sim_sensitivity_validation.py",
                  "test_phase7_sim_postreport.py",
                  "test_phase7_sim_annual_replay_vba.py"):
        assert (TESTS / suite).exists(), f"the record cites a missing suite {suite}"
        assert suite in _text()


def test_42_windows_results_are_attributed_to_the_operator() -> None:
    text = _text()
    assert "produced by the operator" in text
    assert "does not hold a" in text and "transcript" in text
    assert "no independent annual Windows oracle, and none is claimed" in text


def test_43_the_harness_corrections_are_classified_not_blamed_on_production() -> None:
    text = _text()
    section = text.split("## 4.")[1].split("## 5.")[0]
    assert "None of them was a" in section and "production defect" in section
    for defect in ("Join-Path", "2200", "A D", "contingency"):
        assert defect in section, f"the correction record omits {defect}"
    assert "production validation was correct to refuse" in section


def test_44_the_limitations_section_lists_every_caveat_the_record_makes() -> None:
    text = _text()
    section = text.split("## 6.")[1].split("## 7.")[0]
    for caveat in ("no independent annual Windows oracle",
                   "operator-reported",
                   "not re-establishable here",
                   "no W-scenario",
                   "No single sweep",
                   "historical evidence only"):
        assert caveat in section, f"the limitations section omits {caveat!r}"


def test_45_phase_8_is_named_as_next_and_not_started() -> None:
    text = _text()
    assert "Phase 8 may begin" in text
    assert "Phase 8 is not started by this record." in text
    # The handoff surface Phase 8 inherits is named, and it is the real one.
    store = (PCCM_ROOT / "src" / "vba" / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    for accessor in ("PCCM_AnnualDistributionState", "PCCM_AnnualProfileState",
                     "PCCM_AnnualProfilePx", "PCCM_AnnualYearCount"):
        assert accessor in text, f"the record does not name {accessor}"
        assert f"Function {accessor}" in store, f"{accessor} is not a real accessor"
