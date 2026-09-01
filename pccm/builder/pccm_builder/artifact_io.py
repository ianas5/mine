#!/usr/bin/env python3
"""Byte-exact emission for the artefacts whose PHYSICAL bytes are pinned.

WHY THIS EXISTS
---------------
`Path.write_text` opens the file in TEXT mode with `newline=None`, and Python
then translates every `\\n` it writes into `os.linesep`. On Linux that is `\\n`
and nothing happens. On Windows it is `\\r\\n`, and the file on disk is not the
string that was written.

Run 5's pre-Excel check measured exactly that. The accepted invariant artefacts
came out of a clean Windows Stage-A build with different SHA-256s:

    phase6_gate_b_inspection.json   raw eac55e72...  LF-normalised 83eff35f...
    phase6_gate_b_cases.json        raw dee31593...  LF-normalised 6a9d8678...

Both reproduce their accepted hashes exactly once the CRLF translation is
removed - 246 and 247 line endings respectively, and no other difference. The
CONTENT was invariant, as the D2 split intended. The BYTES were not, because the
emitter went through a platform text mode, and Linux could never see it.

    a cross-platform invariant artefact  !=  an artefact whose content is
                                            the same on every host

The second is what the project had. The first is what the pinned hashes claim,
and this module is what makes the claim literally true.

THE CONTRACT
------------
UTF-8, no BOM, LF line endings, a retained final LF, written as BYTES. It is
enforced rather than assumed: a payload that could not be invariant is refused
at the point of writing, where the diagnosis is one line, instead of becoming a
hash mismatch on another machine three steps later.

WHAT USES IT, AND WHAT DELIBERATELY DOES NOT
--------------------------------------------
The artefacts with a pinned SHA-256 use it: `phase6_gate_b_inspection.json`,
the portable `phase6_gate_b_cases.json`, and `phase6_cases.json`. The host-local
`phase6_gate_b_oracle_local.json` uses it too, for serialisation hygiene - that
is NOT a claim of invariance, and it is still never pinned by hash.

The generated `.bas` modules do NOT use it, and must not. `CodeModule.AddFromString`
consumes them on Windows, where CRLF is the separator VBA expects; emitting them
as LF bytes would be a portability "fix" that broke the import it was meant to
protect. The remaining host-local build inputs keep text mode because nothing
claims their bytes are the same on two hosts.
"""

from __future__ import annotations

import codecs
from pathlib import Path

_BOM = codecs.BOM_UTF8


class ArtifactSerialisationError(Exception):
    """A payload that could not be byte-invariant across hosts."""


def write_lf_artifact(path: Path, text: str) -> bytes:
    """Write `text` as UTF-8/LF bytes and return what is now on disk.

    Returns the bytes READ BACK, not the bytes offered. A caller that hashes the
    return value is describing the physical file rather than its own assumption
    about what writing it produced - which is the assumption that failed.
    """
    path = Path(path)
    payload = text.encode("utf-8")
    if payload.startswith(_BOM):
        raise ArtifactSerialisationError(
            f"{path.name} carries a UTF-8 BOM; the invariant artefacts are "
            "UTF-8 without one, and a BOM would move every pinned hash"
        )
    if b"\r" in payload:
        raise ArtifactSerialisationError(
            f"{path.name} contains a carriage return, so its bytes cannot be "
            "the same on Windows and Linux"
        )
    if not payload.endswith(b"\n"):
        raise ArtifactSerialisationError(
            f"{path.name} does not end with a newline; the accepted artefacts do, "
            "and dropping it would move every pinned hash"
        )
    path.write_bytes(payload)
    written = path.read_bytes()
    if written != payload:
        raise ArtifactSerialisationError(
            f"{path.name} does not hold the bytes it was given"
        )
    return written


def canonical_module_identity(data: bytes) -> str:
    """SHA-256 of a generated `.bas` module, with line endings normalised.

    THIS IS AN IDENTITY LAYER, NOT AN INVARIANCE CLAIM. The generated modules
    are written in text mode on purpose - `CodeModule.AddFromString` consumes
    them on Windows, where CRLF is the separator VBA expects - so their physical
    bytes are host-dependent by design and cannot be pinned. What IS the same on
    every host is the projection the accepted renderer produces from the
    accepted authorities, and that is what this identifies.

    NORMALISATION IS LINE ENDINGS AND NOTHING ELSE. Not whitespace, not case,
    not encoding: anything broader would let a real change to the projection
    pass as the same module. A BOM is refused rather than stripped, for the same
    reason, and the final newline is part of the identity.

    AND "LINE ENDINGS" MEANS LF AND CRLF, NOT A BARE CR. The first version wrote
    `.replace(b"\r\n", b"\n").replace(b"\r", b"\n")`, which quietly admitted a
    third representation: a module whose every LF had been replaced by a lone CR
    hashed identically to the accepted one. That is not a line-ending difference
    the accepted renderers can produce - it is a corrupted or foreign file, and
    an identity layer that maps it onto the accepted hash is answering a
    different question from the one it was asked. A bare CR is REFUSED.
    """
    if data.startswith(_BOM):
        raise ArtifactSerialisationError(
            "a generated module carries a UTF-8 BOM; the accepted projection "
            "does not, and stripping one here would hide the difference"
        )
    stray = data.replace(b"\r\n", b"")
    if b"\r" in stray:
        raise ArtifactSerialisationError(
            "a generated module carries a carriage return that is not part of a "
            "CRLF; LF and CRLF are the accepted representations and a bare CR is "
            "not a third one"
        )
    canonical = data.replace(b"\r\n", b"\n")
    if not canonical.endswith(b"\n"):
        raise ArtifactSerialisationError(
            "a generated module does not end with a newline; the accepted "
            "projection does, and the final newline is part of its identity"
        )
    import hashlib

    return hashlib.sha256(canonical).hexdigest()
