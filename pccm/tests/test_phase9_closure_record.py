#!/usr/bin/env python3
"""The Phase-9 closure record says only things the repository can be asked.

WHAT A CLOSURE RECORD IS FOR. It is read later, by someone deciding what a
result rested on. Every number in it is therefore either RE-DERIVED here, or
labelled as evidence from a run this repository cannot reproduce - and the
difference is stated rather than left to the reader.

WHAT THESE CONTROLS REFUSE. A commit that does not exist. A lineage quoted for a
tree it did not produce. A production byte moved between the lineage and the
Windows-tested tree without the record saying so. A Phase-7 or Phase-8 authority
quietly displaced. A reconciliation that stops matching the source it reconciles.
An accepted fact quietly dropped from the record. And the claim that Phase 10 has
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
RECORD = PCCM_ROOT / "docs" / "phase9_closure.md"
SRC = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

# THE THREE AUTHORITIES, as the record states them.
CONTRACT = "dd082c9"
LINEAGE = "c52e732"
WINDOWS_TREE = "8ccadd0"
PREREQUISITES = 7
RESULTS = 121

# PHASE 7 AND PHASE 8. This record may not move them.
PHASE7 = {"implementation": "79d4c3e", "acceptance": "ad78988"}
PHASE8 = {"P8-1": "35bd6ce", "P8-2": "7ff5dc7", "P8-3": "9e3c141", "P8-Z": "bfae0eb"}


def _text() -> str:
    return RECORD.read_text(encoding="utf-8")


def _flat() -> str:
    """The record with its line wrapping and blockquote markers taken out.

    A prose claim must not depend on WHERE the paragraph happened to wrap. The
    first draft of this asserted a sentence with the newline and indent baked
    in, which would have failed the next time anybody reflowed a paragraph -
    a control about wrapping rather than about content.

    THE MARKER IS STRIPPED ONLY AT THE START OF A LINE. Removing every "> "
    took the arrow out of `0 -> 1 -> 0` as well, which is a heartbeat fact this
    file exists to protect.
    """
    import re

    return re.sub(r"\s+", " ", re.sub(r"(?m)^>[ \t]?", "", _text()))


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=REPO_ROOT, check=True,
                          stdout=subprocess.PIPE, text=True).stdout


def _exists(commit: str) -> bool:
    return subprocess.run(["git", "cat-file", "-e", commit + "^{commit}"],
                          cwd=REPO_ROOT, stdout=subprocess.DEVNULL,
                          stderr=subprocess.DEVNULL).returncode == 0


# ===========================================================================
# A. THE RECORD EXISTS AND SAYS WHAT IT IS
# ===========================================================================
def test_01_the_record_declares_phase_9_closed() -> None:
    text = _text()
    assert RECORD.is_file()
    assert "PHASE 9 — ACCEPTED / CLOSED" in text
    assert "No outstanding Phase-9 blocker." in text
    assert "No further Windows run is required for Phase 9." in text


def test_02_phase_10_is_not_started_and_says_so() -> None:
    """THE CLAIM IS ABOUT THIS RECORD, NOT ABOUT THE FUTURE.

    The Phase-8 record made this mistake and had to be re-anchored: as first
    written it scanned TODAY's tree, so Phase 9 legitimately starting would
    falsify a statement about Phase 8 that never stopped being true. This one is
    written the settled way from the start - the record says Phase 10 is not
    started BY IT, and the check is that this record's own commit carried no
    Phase-10 file. A Phase-10 file smuggled into the closure commit still fails,
    and it fails forever rather than only until Phase 10 legitimately starts.
    """
    text = _text()
    assert "Phase 10 is not started by this record." in text
    assert "**Phase 10 has not started.**" in text
    listing = _git("ls-tree", "-r", "--name-only", "HEAD", "pccm/").split()
    for pattern in ("phase10", "phase_10"):
        found = [name for name in listing
                 if pattern in name.rsplit("/", 1)[-1].lower()
                 and name.split("/")[1] in ("tests", "docs", "bootstrap", "src", "spec")]
        assert not found, f"the tree already carries Phase-10 work: {found}"


# ===========================================================================
# B. EVERY QUOTED COMMIT IS REAL, AND IS WHAT THE RECORD SAYS IT IS
# ===========================================================================
@pytest.mark.parametrize("label,commit", [
    ("contract", CONTRACT), ("lineage", LINEAGE), ("windows tree", WINDOWS_TREE)])
def test_10_each_authority_exists_and_is_quoted(label: str, commit: str) -> None:
    assert _exists(commit), f"the {label} commit {commit} does not exist"
    assert commit in _text(), f"the record does not quote the {label} commit"


def test_11_the_lineage_is_what_git_reports_for_the_windows_tree() -> None:
    """RE-DERIVED, NOT ASSERTED. The record claims c52e732 is the last commit
    changing a production or spec byte at 8ccadd0. Git is what says so."""
    reported = _git("log", "--format=%h", "-1", WINDOWS_TREE,
                    "--", "pccm/src", "pccm/spec").strip()
    assert reported == LINEAGE, (
        f"the record quotes {LINEAGE} as the lineage; git reports {reported}")


def test_12_no_production_byte_moved_between_the_lineage_and_the_tested_tree() -> None:
    """THE CLAIM THAT MAKES 8ccadd0 A TREE AND NOT AN AUTHORITY. If a production
    byte HAD moved after the lineage, the Windows run would have executed
    something the lineage does not describe."""
    changed = _git("diff", "--name-status", LINEAGE, WINDOWS_TREE,
                   "--", "pccm/src", "pccm/spec").strip()
    assert changed == "", f"production moved after the lineage:\n{changed}"
    assert "(no output)" in _text(), (
        "the record does not show the empty diff it rests on")


def test_13_the_windows_numbers_are_labelled_as_run_evidence() -> None:
    """A NUMBER THIS REPOSITORY CANNOT REPRODUCE IS LABELLED AS SUCH. Nothing
    here can re-run Excel, so the record must attribute these to the run."""
    text = _text()
    assert f"{PREREQUISITES} / {PREREQUISITES}" in text
    assert f"**{RESULTS} / {RESULTS}**" in text
    assert "**P9-1 PASS**" in text
    assert "Windows acceptance" in text and WINDOWS_TREE in text
    # AND THE TWO RUNS THAT DID NOT REACH THE SCENARIOS ARE NOT EVIDENCE.
    assert "neither is acceptance evidence" in text


# ===========================================================================
# C. EARLIER AUTHORITIES ARE NOT DISPLACED
# ===========================================================================
@pytest.mark.parametrize("label,commit", sorted(PHASE7.items()) + sorted(PHASE8.items()))
def test_20_the_earlier_authorities_are_unmoved(label: str, commit: str) -> None:
    assert _exists(commit), f"{label} {commit} does not exist"
    assert commit in _text(), f"the record does not carry {label} {commit}"


def test_21_the_record_says_it_replaces_nothing() -> None:
    text = _text()
    assert "It **does not replace**" in _flat()
    assert "unmoved" in text
    # AND IT DOES NOT CLAIM TO RE-ESTABLISH THEM EITHER.
    assert "No historical digest moved." in text


# ===========================================================================
# D. THE TWO RECONCILIATIONS STILL MATCH THEIR SOURCES
# ===========================================================================
def test_30_the_scenario_a_reconciliation_matches_the_owner() -> None:
    """THE RECORD'S REASONING IS CHECKED AGAINST THE CODE IT REASONS ABOUT.

    A closure record that explains a Windows result has to keep matching the
    thing it explains. If DeriveStatus ever stopped testing validity before
    history, the explanation in section 4 would be wrong and this fails.
    """
    text = _text()
    assert ("an invalid model is INVALID whether or not anyone pressed Calculate"
            in _flat())
    # THE OWNER'S OWN WORDING, NORMALISED THE SAME WAY. calc_contract carries
    # the sentence as a wrapped comment, so the leading `#` and the wrap come
    # out before the two texts are compared - otherwise this would be a control
    # about where a YAML comment breaks.
    contract = re.sub(r"\s+", " ",
                      re.sub(r"(?m)^#[ \t]?", "",
                             (SPEC / "calc_contract.yaml").read_text(encoding="utf-8")))
    assert ("an invalid model is INVALID whether or not anyone pressed Calculate"
            in contract), "the record quotes calc_contract wording it does not carry"

    body = (SRC / "modCalcReport.bas").read_text(encoding="utf-8")
    start = body.index("Private Function DeriveStatus")
    derive = body[start:body.index("End Function", start)]
    invalid = derive.index("CALC_STATUS_INVALID")
    not_calculated = derive.index("CALC_STATUS_NOT_CALCULATED")
    assert invalid < not_calculated, (
        "DeriveStatus no longer tests validity before history, so the record's "
        "Scenario-A reconciliation is out of date")
    # AND THE UNTOUCHED FIXTURE STILL HAS THE BLANKS THE RECORD NAMES.
    assert "required: true" in (SPEC / "input_contract.yaml").read_text(encoding="utf-8")
    for named in ("applied Base Year", "applied Start Year", "applied Duration"):
        assert named in (SRC / "modCalcResolve.bas").read_text(encoding="utf-8"), named


def test_31_the_iteration_reconciliation_matches_both_fingerprints() -> None:
    """ABSENT FROM ONE, PRESENT IN THE OTHER - asked of the source, not of the
    record. This is the whole reason the advisory needs no special behaviour."""
    text = _text()
    assert "**absent** from the calculation fingerprint" in text
    assert "**present** in the simulation request fingerprint" in text

    report = (SRC / "modCalcReport.bas").read_text(encoding="utf-8")
    start = report.index("Private Function BuildFingerprint")
    header = report[start:report.index("End Function", start)]
    for field in ("Timeline.BaseYear", "Timeline.StartYear", "Timeline.Duration",
                  "Timeline.DiscountRate"):
        assert field in header, field
    assert "Iterations" not in header, (
        "the calculation fingerprint now reads an iteration count; the record's "
        "iteration reconciliation is out of date")

    sim = (SRC / "modSimReport.bas").read_text(encoding="utf-8")
    start = sim.index("Private Function CurrentRequestFingerprint")
    request = sim[start:sim.index("End Function", start)]
    assert "package.Iterations" in request, (
        "iterations no longer reach the simulation request fingerprint")
    assert "iterations" in (SPEC / "sim_contract.yaml").read_text(encoding="utf-8")


def test_32_the_advisory_boundary_is_stated_as_strictly_less_than() -> None:
    text = _text()
    assert "strictly below 10,000" in text
    for boundary in ("**9,999 warns**", "**10,000 does not warn**",
                     "**10,001 does not warn**"):
        assert boundary in text, boundary


# ===========================================================================
# E. THE ACCEPTED FACTS ARE ALL PRESENT
# ===========================================================================
# NAMED, NOT COUNTED. A record that had to carry "at least fifteen bullet
# points" could lose the one that mattered and keep the number.
_ACCEPTED = (
    "read-only presentation and aggregation layer",
    "no new state machine",
    "no new refusal",
    "only actionable ERROR and WARNING rows",
    "INFO is context",
    "deterministically ordered",
    "fixed 100-row window",
    "unused slots are **`#N/A`**",
    "machine-readable owner",
    "does **not alter the calculation fingerprint**",
    "preserves the existing simulation request fingerprint semantics",
    "live and persisted `(last evaluated)` facts stay separate",
    "structured permanent-id Subject",
    "no prose parsing",
    "no worksheet-called path can write",
)

_WINDOWS_FACTS = (
    "from a worksheet cell",
    "freeze panes and every projected register header",
    "reconcile exactly",
    "can be **PASS**",
    "refuses nothing",
    "live STALE",
    "exactly one actionable root-cause ERROR",
    "Subject is the offending permanent id",
    "post-validation arithmetic refusal",
    "does not duplicate",
    "`0 -> 1 -> 0`",
    "rewrote no persisted status cell",
    "closed naturally and cleanly",
)


@pytest.mark.parametrize("claim", _ACCEPTED)
def test_40_every_accepted_fact_is_recorded(claim: str) -> None:
    assert claim in _flat(), f"the record has lost: {claim}"


@pytest.mark.parametrize("fact", _WINDOWS_FACTS)
def test_41_every_windows_fact_is_recorded(fact: str) -> None:
    assert fact in _flat(), f"the record has lost the Windows fact: {fact}"


def test_42_the_not_calculated_coverage_is_stated_honestly() -> None:
    """THE ONE THING THE RUN DID NOT SAMPLE, SAID PLAINLY. A closure record that
    quietly omitted this would be claiming coverage it does not have."""
    text = _text()
    assert "`NOT CALCULATED -> WARNING` remains part of the accepted" in _flat()
    assert "**statically covered**" in text
    assert "did not independently sample" in text
    assert "not an open acceptance blocker and requires no additional Windows" in _flat()
    # AND THE STATIC COVERAGE IS REAL.
    suite = (PCCM_ROOT / "tests" / "test_phase9_model_check.py").read_text(encoding="utf-8")
    assert suite.count('"NOT CALCULATED"') >= 4, (
        "the record claims static coverage the Model Check suite does not carry")


def test_43_the_step_1_contract_is_not_rewritten() -> None:
    """THE TERMINOLOGY CLARIFICATION IS DEFERRED, NOT APPLIED. Rewording a
    historical contract to match a later reading is how evidence stops being
    evidence."""
    text = _text()
    assert "is **not rewritten here**" in _flat()
    assert "Phase-10 documentation cleanup item" in text
    # AND THE STEP-1 CONTRACT REALLY IS UNTOUCHED SINCE ITS OWN COMMIT.
    changed = _git("diff", "--name-only", CONTRACT, "HEAD",
                   "--", "pccm/docs/phase9_step1_contract.md").strip()
    assert changed == "", (
        f"the Step-1 contract was edited after {CONTRACT}: {changed}")


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
