#!/usr/bin/env python3
"""Take the P9-2B subject plumbing back out of a VBA module.

WHY A REVERSAL AND NOT A NEW DIGEST. Several accepted controls pin a module - or
the accepted prefix of one - to the bytes an earlier phase left behind. P9-2B
threads a structured `subject` out-parameter through the current-model
preparation so a caller can have the offending permanent id as a VALUE instead
of reading it out of the refusal sentence. Moving those pins to new opaque
numbers would record THAT something changed and stop proving WHAT.

So the plumbing is removed mechanically and the ORIGINAL digest is required
back. Delete the named comment blocks, delete every line that is nothing but a
subject assignment, take the parameter out of the signatures and the calls,
rejoin the continuations it split - and the module must be the earlier bytes, to
the byte. Nothing else can hide inside that: a changed condition, a reworded
message, a moved Boolean or a different constant all survive the reversal, and
then the digest fails exactly as it always did.

ONE DEFINITION, USED BY EVERY CONTROL THAT NEEDS IT. Four suites pin these
bytes; four copies of this rule would be four chances for one of them to be
quietly more permissive than the others.
"""

from __future__ import annotations

import re

__all__ = ["COMMENT_ADDITIONS", "reverse_subject_plumbing"]

# The prose P9-2B added, per module, matched EXACTLY. Absence is a failure: if
# the text is not what this reverses, what else changed cannot be established.
COMMENT_ADDITIONS: dict[str, tuple[str, ...]] = {
    "modCalcResolve": (
        "    ' SUBJECT IS PLUMBING, NOT A RULE. Where a refusal below is about ONE driver\n"
        "    ' the owner already holds its permanent id; this carries that id out so a\n"
        "    ' caller need not read the sentence. It stays blank for every model-wide\n"
        "    ' refusal - a missing register, an FX table, an inflation grid - because no\n"
        "    ' driver is at fault in those, and blank on success.\n",
        "    ' FROM HERE THE ROW HAS AN IDENTITY. Anything refused above this line is a\n"
        "    ' row whose Permanent ID could not be read at all, and it has no id to name.\n",
    ),
    "modCalcCheck": (
        "    '\n"
        "    ' SUBJECT IS PLUMBING, NOT A RULE. It carries the permanent id this function\n"
        "    ' ALREADY holds when a per-driver predicate refuses - the same id DriverLabel\n"
        "    ' builds its sentence from - so a caller can have it without reading prose.\n"
        "    ' No condition, no Boolean and no message below is touched by it. It is blank\n"
        "    ' for a model-level refusal, because no driver is at fault in one, and blank\n"
        "    ' again on success, so a later refusal elsewhere cannot inherit it.\n",
    ),
    # modCalcReport's plumbing added no commentary: the reporter prefix is at its
    # accepted raw-line ceiling and had none to spare.
    "modCalcReport": (),
}


def reverse_subject_plumbing(module: str, text: str) -> str:
    """*text* with every trace of the subject out-parameter removed.

    Unknown modules are returned untouched, so a control can call this over a
    whole inventory without needing to know which files the correction reached.
    """
    blocks = COMMENT_ADDITIONS.get(module)
    if blocks is None:
        return text
    for block in blocks:
        if block not in text:
            raise AssertionError(
                f"{module}: a P9-2B comment block is not the text this reversal "
                "removes, so what else changed cannot be established")
        text = text.replace(block, "", 1)

    stripped: list[str] = []
    for line in text.split("\n"):
        # A LINE THAT IS NOTHING BUT A SUBJECT ASSIGNMENT GOES ENTIRELY. Every
        # one of them is either the id the owner already holds or the clear that
        # stops it leaking; neither existed before.
        if re.match(r"^\s*subject = ", line):
            continue
        line = re.sub(r",?\s*ByRef subject As String", "", line)
        line = line.replace(", subject As String", "")
        line = line.replace(": subject = vbNullString", "")
        line = re.sub(r",\s*subject(?=\))", "", line)
        stripped.append(line)

    # AND A SIGNATURE THE PARAMETER SPLIT IS PUT BACK ON ONE LINE. Without this
    # the reversal would leave a dangling continuation and every digest would
    # fail for a reason that has nothing to do with what changed.
    joined: list[str] = []
    for line in stripped:
        if joined and joined[-1].rstrip().endswith(", _") and line.strip().startswith(")"):
            joined[-1] = joined[-1].rstrip()[:-3] + line.strip()
            continue
        joined.append(line)
    return "\n".join(joined)
