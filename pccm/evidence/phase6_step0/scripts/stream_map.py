"""D6-16: component -> stream assignment. Families A and B, both exercised.

A component is one consumer of one stream. Cost Line -> 1 component.
Risk -> 2 components (occurrence, severity), per the accepted plan section 5.6.
"""
import os
import re
import sys

_BUILDER = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))))), "builder")
if _BUILDER not in sys.path:
    sys.path.insert(0, _BUILDER)

from pccm_builder.calc_fingerprint import utf16_sort_key   # noqa: E402

COST_ID = re.compile(r"^CL-[0-9]{3,}$")     # the ACCEPTED Permanent-ID pattern
RISK_ID = re.compile(r"^R-[0-9]{3,}$")      # {3,} is UNBOUNDED above

KIND_RANK = {"COST": 0, "RISK": 1}
ROLE_RANK = {"value": 0, "occurrence": 0, "severity": 1}


def components(cost_ids, risk_ids):
    """The component set, independent of any row order."""
    out = [("COST", cid, "value") for cid in cost_ids]
    for rid in risk_ids:
        out.append(("RISK", rid, "occurrence"))
        out.append(("RISK", rid, "severity"))
    return out


def is_ascii_id(s):
    """Every ID admitted by the accepted patterns is ASCII. Checked, not assumed:
    above the BMP, code-point order and UTF-16 code-unit order disagree."""
    return all(ord(ch) < 0x80 for ch in s)


def sort_key(comp):
    """Ordinal UTF-16 code-unit order, using the ACCEPTED Phase-5 sort key.

    Not Python's own string ordering (code points), not locale collation, not a
    case-insensitive comparison. Reusing `utf16_sort_key` means no new collation
    authority is created by D6-16.
    """
    kind, pid, role = comp
    return (KIND_RANK[kind], utf16_sort_key(pid), ROLE_RANK[role])


def family_a(cost_ids, risk_ids):
    """Canonical sorted order -> sequential stream indices 0..N-1."""
    ordered = sorted(components(cost_ids, risk_ids), key=sort_key)
    return {c: i for i, c in enumerate(ordered)}


def numeric_part(pid):
    return int(pid.split("-", 1)[1])


def family_b(cost_ids, risk_ids, risk_offset):
    """index = kind_offset + numeric_id. Requires a BOUNDED numeric domain."""
    m = {}
    for cid in cost_ids:
        m[("COST", cid, "value")] = numeric_part(cid)
    for rid in risk_ids:
        n = numeric_part(rid)
        m[("RISK", rid, "occurrence")] = risk_offset + 2 * n
        m[("RISK", rid, "severity")] = risk_offset + 2 * n + 1
    return m


def family_b_collision_witness(risk_offset):
    """For ANY finite risk_offset K, exhibit a REPRESENTABLE Cost-Line ID that
    collides with a Risk component. This is what an unbounded ID pattern costs.

    Returns (cost_id, risk_id, shared_index) or None if no witness exists.
    """
    # The Risk block starts at K + 2*n for the smallest admitted n. The smallest
    # numeric part admitted by ^R-[0-9]{3,}$ is 0 (written "R-000").
    target = risk_offset + 0
    cost_id = "CL-" + str(target).zfill(3)
    if not COST_ID.match(cost_id):
        return None                      # not representable -> no witness
    a = numeric_part(cost_id)
    b = risk_offset + 2 * 0
    return (cost_id, "R-000", a) if a == b else None
