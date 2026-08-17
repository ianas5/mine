#!/usr/bin/env python3
"""Package the tracked PCCM tree for independent review, byte-for-byte.

WHY THIS EXISTS
---------------
A review package must be diffable against the repository without reporting noise.
Zipping the WORKING TREE does not achieve that, and neither does `git archive`:

  * `pccm/.gitattributes` declares `*.ps1 text eol=crlf`, because VBA's
    `CodeModule.AddFromString` needs CRLF separators on Windows;
  * git therefore stores those files with LF and converts on checkout;
  * `git archive` applies the same conversion, so the archived bytes are CRLF
    while the tracked bytes are LF;
  * and a file rewritten in place by tooling can stay LF in the working tree while
    git still reports it clean, because the attribute normalises the comparison.

The result was that two Phase-4 PowerShell files differed by raw hash between the
accepted Phase-4 package and the Step-1 package while being identical after
line-ending normalisation. No source changed; only the packaging did.

WHAT THIS DOES
--------------
Reads each blob straight out of the git object store and writes those exact bytes
into the archive. No smudge filter, no eol conversion, no working-tree read. What
ships is what is tracked, so a frozen-source diff compares like with like.

    python pccm/tools/package_review.py <output.zip> [revision]

`revision` defaults to HEAD. File modes are preserved so executable bits survive.
"""

from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path

PREFIX = "pccm/"
SUBTREE = "pccm"


def _git(*args: str) -> bytes:
    return subprocess.run(
        ["git", *args], check=True, stdout=subprocess.PIPE, cwd=_repo_root()
    ).stdout


def _repo_root() -> Path:
    return Path(__file__).resolve().parent.parent.parent


def _entries(revision: str) -> list[tuple[str, str, str]]:
    """(mode, blob sha, path) for every tracked file under the subtree."""
    listing = _git("ls-tree", "-r", "-z", f"{revision}:{SUBTREE}").decode("utf-8")
    entries = []
    for record in listing.split("\0"):
        if not record:
            continue
        meta, path = record.split("\t", 1)
        mode, kind, sha = meta.split(" ", 2)
        if kind != "blob":
            continue
        entries.append((mode, sha, path))
    return sorted(entries, key=lambda e: e[2])


def main(argv: list[str]) -> int:
    if not 2 <= len(argv) <= 3:
        print(__doc__)
        return 2
    output = Path(argv[1]).resolve()
    revision = argv[2] if len(argv) == 3 else "HEAD"

    entries = _entries(revision)
    if not entries:
        print(f"no tracked files under {SUBTREE}/ at {revision}", file=sys.stderr)
        return 1

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for mode, sha, path in entries:
            # Raw object bytes. This is the whole point: no filters are applied.
            blob = _git("cat-file", "blob", sha)
            info = zipfile.ZipInfo(PREFIX + path)
            info.date_time = (1980, 1, 1, 0, 0, 0)  # deterministic, not build-time
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (0o755 if mode == "100755" else 0o644) << 16
            archive.writestr(info, blob)

    print(f"{output}  ({len(entries)} tracked files from {revision}:{SUBTREE})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
