#!/usr/bin/env python3
"""P10-2B: take the Reset Results clear back out of a publication owner.

WHY A REVERSAL RATHER THAN A NEW DIGEST. Three accepted endpoint owners gained a
narrow clear/restore pair so that Reset Results could ask each of them to clear
what it wrote. The controls that pinned those files to the accepted P10-2A tree
were therefore going to be false, and the two honest ways to settle that are:

    move the pin to the new bytes, and lose the claim that nothing else moved;
    or remove the addition mechanically and require the accepted bytes back.

The second is stronger and is the settlement P9-2B already established for the
structured-subject plumbing. A pin that moves says "these are the bytes now". A
reversal says "the ONLY thing that changed is the declared block, because taking
it out reproduces the accepted file exactly" - which is the claim the control
exists to make.

THE ADDITION IS ONE APPENDED BLOCK PER OWNER, and that is not a convenience: it
is what makes the reversal exact. Nothing was inserted, re-indented, renamed or
re-ordered anywhere above it, so the accepted region is a literal prefix of the
current file.
"""
from __future__ import annotations

# The first line of the banner every P10-2B addition opens with. It is matched
# together with the rule line above it, so the phrase appearing inside a comment
# somewhere else cannot be mistaken for the start of the block.
BANNER = "' P10-2B. THE RESET CLEAR"
RULE = "' " + "=" * 74


def reset_addition_of(text: str) -> str:
    """The declared P10-2B block, or "" when the file carries none."""
    marker = _marker(text)
    if marker is None:
        return ""
    return text[text.index(marker):]


def strip_reset_addition(text: str) -> str:
    """`text` with the declared P10-2B block removed.

    A file that never gained one comes back unchanged, so the same reversal can
    be applied to every owner without asking first which ones were touched.
    """
    marker = _marker(text)
    if marker is None:
        return text
    return text[: text.index(marker)]


def _marker(text: str) -> str | None:
    """The exact separator that opens the addition, honouring line endings.

    modCalcReport.bas is CRLF and every other module is LF. An anchor written
    with one and matched against the other finds nothing - and finding nothing
    would silently mean "this file was not touched", which is the failure mode
    that makes a control useless rather than red.

    The banner line ends with the owner's own name, so only its stable prefix is
    matched; the RULE line above it is what makes that prefix unambiguous.
    """
    for eol in ("\r\n", "\n"):
        marker = eol + RULE + eol + BANNER
        found = text.count(marker)
        if found == 1:
            return marker
        if found > 1:
            raise AssertionError(
                f"the P10-2B banner appears {found} times; the reversal removes "
                "one appended block, not several")
    return None
