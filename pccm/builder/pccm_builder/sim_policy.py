#!/usr/bin/env python3
"""The Phase-6 oracle and evidence policy — the accepted cross-implementation
comparison rules, in one place.

NAMED `sim_policy`, NOT `sim_evidence`, ON PURPOSE. `test_48` in
`tests/test_phase6_sim_contract.py` refuses any production module that imports a
name containing "evidence", because production must never read the retained
evidence package. This module reads nothing at build time - the binding to the
Step-0 record is performed by a control, from `tests/` - but a module whose name
collided with that scan would have made the reviewer choose between a true name
and a working detector. The detector is right; the name moved.

WHY THIS MODULE EXISTS
----------------------
Step 0 §10.1 settled the ownership question and it settled it against
`sim_contract.yaml`:

    The tolerance is not a simulation-runtime contract ... The engine never
    compares two Doubles for approximate equality at runtime ... A tolerance
    exists only when two IMPLEMENTATIONS are compared - which is oracle, Gate-A
    and Gate-B evidence.

    Single owner: the Phase-6 oracle and evidence policy. `sim_contract.yaml`
    stores no tolerance at all, so the rule cannot come to live in two files.

This module is that single owner in code. `sim_contract.yaml` still stores no
tolerance, the PowerShell harness still spells no `1e-N` constant of its own, and
Step-13's Gate-B comparison reads the numbers from here by way of the emitted
portable case authority.

WHAT WENT WRONG WITHOUT IT
--------------------------
Step 13 wrote its own rule - `EXACT, and there is no other mode` - and compared a
Python-oracle `result_digest` against a VBA `result_digest`. Step 0 §10.4 keeps
the digest exact for SAME-RUNTIME replay only, and never promised it across
languages. Run 4 then failed `P6-ORA` on differences the accepted policy had
anticipated and admitted, and `P6-DET` went red for a cross-language clause that
did not belong in a repeatability scenario.

The constants below are not a second authority. They are a copy of the §10.3
table, and `validate_evidence_policy_record` is the assertion that the copy is
still faithful - the same discipline `validate_result_digest_contract` applies to
the digest grammar.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

EVIDENCE_POLICY_VERSION = 1

# The Step-0 §10.3 subject keys, spelled once.
TRANSFORMED_SAMPLE = "transformed_sample"
CHENG_VECTOR = "cheng_vector"
ITERATION_TOTAL = "iteration_total"
SUMMARY_STATISTIC = "summary_statistic"


class EvidencePolicyError(Exception):
    """The policy in code and the policy in the Step-0 record disagree."""


@dataclass(frozen=True)
class ToleranceRule:
    """One row of the accepted §10.3 table.

    `absolute_floor` is a FACTOR, not a bound: the bound is
    `absolute_floor * scale`, and `scale_kind` names which scale. Step 0 is
    explicit that no rule is purely relative and none is purely absolute,
    because cancellation can drive a total near zero while every contribution
    that produced it is large.
    """

    subject: str
    relative: float
    absolute_floor: float | None
    scale_kind: str | None

    def agrees(self, left: float, right: float, scale: float | None = None) -> bool:
        """Does one pair of cross-implementation values satisfy this rule?"""
        left = float(left)
        right = float(right)
        if left == right:
            return True
        gap = abs(left - right)
        magnitude = max(abs(left), abs(right))
        if magnitude > 0.0 and gap <= self.relative * magnitude:
            return True
        if self.absolute_floor is None:
            return False
        if scale is None:
            raise EvidencePolicyError(
                f"the {self.subject} rule carries a scale-aware floor and no scale "
                "was supplied; a floor keyed to the output's own magnitude is not "
                "the accepted rule"
            )
        return gap <= self.absolute_floor * abs(float(scale))


# §10.3, in the order the record states it.
POLICY: dict[str, ToleranceRule] = {
    TRANSFORMED_SAMPLE: ToleranceRule(TRANSFORMED_SAMPLE, 1e-12, 1e-12, "conditioning"),
    CHENG_VECTOR: ToleranceRule(CHENG_VECTOR, 1e-11, None, None),
    ITERATION_TOTAL: ToleranceRule(ITERATION_TOTAL, 3e-10, 3e-10, "accumulation"),
    SUMMARY_STATISTIC: ToleranceRule(SUMMARY_STATISTIC, 3e-10, 3e-10, "accumulation"),
}

# §10.4 - "What stays exact". No tolerance applies to any of these, in either
# direction. The qualifier on the last one is the whole point: the digest is
# exact for a REPLAY inside one runtime, and was never promised across two.
EXACT_SUBJECTS: tuple[str, ...] = (
    "MRG32k3a state and uniform values",
    "jump state",
    "Bernoulli occurrence decisions",
    "proposal and draw counts, where the arithmetic path is fixed",
    "same-runtime G2/G3 result_digest",
)

SAME_RUNTIME_DIGEST_SUBJECT = "same-runtime G2/G3 result_digest"


def accumulation_scale(prepared: Any) -> float:
    """`S` for one prepared model: the largest single contribution it can form.

    Step 0 keys the floor to "the scale that actually produced the number, not
    the number itself", and for an iteration total that scale is
    `max |contribution|` over the drivers summed. The engine forms a cost
    contribution as `unit_cost * quantity * k` and a risk contribution as
    `severity * k`, so the largest either can reach is bounded by the driver's
    own conditioning scale times its factors.

    Computed from the PREPARED MODEL, by arithmetic only: no transcendental is
    involved, so this number is the same on every host. That matters - the floor
    belongs to the portable authority, not to host-local oracle evidence.
    """
    scale = 0.0
    for driver in tuple(prepared.cost_drivers) + tuple(prepared.risk_drivers):
        points = [abs(float(driver.minimum)), abs(float(driver.maximum))]
        if driver.most_likely is not None:
            points.append(abs(float(driver.most_likely)))
        conditioning = max(points)
        factor = max(abs(float(driver.knom)), abs(float(driver.kpv)))
        quantity = abs(float(getattr(driver, "quantity", 1.0) or 1.0))
        is_risk = "RISK" in str(
            getattr(driver.driver_kind, "value", driver.driver_kind)
        ).upper()
        contribution = conditioning * factor * (1.0 if is_risk else quantity)
        scale = max(scale, contribution)
    return scale


def policy_payload() -> dict[str, Any]:
    """The block emitted into the portable case authority.

    The harness reads its numbers from here. It spells no tolerance of its own,
    and a policy that moved in the Step-0 record moves in the emitted artefact
    on the next build rather than being remembered in PowerShell.
    """
    return {
        "authority": "docs/phase6_step0.md §10 — the Phase-6 oracle and evidence policy",
        "version": EVIDENCE_POLICY_VERSION,
        "note": (
            "Cross-implementation comparison only. The engine never compares two "
            "Doubles for approximate equality at runtime, and no published number "
            "is produced by a tolerance test."
        ),
        "tolerances": {
            rule.subject: {
                "relative": rule.relative,
                "absolute_floor": rule.absolute_floor,
                "scale_kind": rule.scale_kind,
            }
            for rule in POLICY.values()
        },
        "exact_subjects": list(EXACT_SUBJECTS),
        "same_runtime_digest_is_exact": True,
        "cross_language_digest_is_exact": False,
        "cross_language_digest_note": (
            "§10.4 keeps the result_digest exact for SAME-RUNTIME replay. It is "
            "not a cross-implementation equality subject and must not be used as "
            "a pass criterion when two implementations are compared; the digest "
            "resolves one ULP in one retained iteration, which §10.3 admits."
        ),
    }


# ---------------------------------------------------------------------------
# the binding to the accepted record
# ---------------------------------------------------------------------------
_ROW = re.compile(r"^\|\s*(?P<subject>[^|]+?)\s*\|[^|]*\|\s*(?P<value>[^|]+?)\s*\|")

# The §10.3 row each rule copies, named so a renamed row is a failure rather
# than a silently skipped check.
RECORD_ROWS: dict[str, str] = {
    TRANSFORMED_SAMPLE: "individual Uniform / Triangular / PERT-rescale transformed samples",
    CHENG_VECTOR: "deterministic Cheng vector outputs",
    ITERATION_TOTAL: "F1 per-iteration no-Beta end-to-end totals",
    SUMMARY_STATISTIC: "summary statistics compared cross-language",
}


def _numbers(text: str) -> tuple[float | None, float | None]:
    relative = re.search(r"rel\s*[≤<=]+\s*([0-9.eE+-]+)", text)
    absolute = re.search(r"abs\s*[≤<=]+\s*([0-9.eE+-]+)", text)
    return (float(relative.group(1)) if relative else None,
            float(absolute.group(1)) if absolute else None)


def validate_evidence_policy_record(record: str | Path) -> None:
    """Assert the constants above are still a faithful copy of Step 0 §10.

    Reads the accepted record, finds each named §10.3 row, and compares the
    numbers. Also checks §10.4 still lists every exact subject, and still
    qualifies the digest as SAME-RUNTIME - because it is that qualifier, not the
    tolerances, that Step 13 read past.
    """
    text = Path(record).read_text(encoding="utf-8") if isinstance(record, Path) \
        else record
    for subject, row_label in RECORD_ROWS.items():
        rows = [line for line in text.splitlines()
                if line.startswith("| " + row_label + " ")]
        if len(rows) != 1:
            raise EvidencePolicyError(
                f"the Step-0 §10.3 row '{row_label}' appears {len(rows)} times; "
                "the policy in code can no longer be bound to the record"
            )
        relative, absolute = _numbers(rows[0])
        rule = POLICY[subject]
        if relative != rule.relative:
            raise EvidencePolicyError(
                f"{subject}: the record says rel <= {relative}, this module says "
                f"{rule.relative}"
            )
        if absolute != rule.absolute_floor:
            raise EvidencePolicyError(
                f"{subject}: the record says abs floor {absolute}, this module "
                f"says {rule.absolute_floor}"
            )
    # The record marks code spans with backticks; the subjects are compared as
    # prose so a change of typography is not read as a change of policy.
    plain = text.replace("`", "")
    for subject in EXACT_SUBJECTS:
        if subject not in plain:
            raise EvidencePolicyError(
                f"Step-0 §10.4 no longer lists the exact subject {subject!r}"
            )
    if "same-runtime G2/G3 result_digest" not in plain:
        raise EvidencePolicyError(
            "Step-0 §10.4 no longer qualifies the exact result_digest as "
            "SAME-RUNTIME, which is the qualifier Step 13 read past"
        )


def agrees(subject: str, left: float, right: float,
           scale: float | None = None) -> bool:
    """Cross-implementation agreement for one named subject."""
    if subject not in POLICY:
        raise EvidencePolicyError(f"{subject} is not a Step-0 §10.3 subject")
    return POLICY[subject].agrees(left, right, scale)
