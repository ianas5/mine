"""Static reading of VBA source, for the Phase-4 Linux tests.

VBA cannot be compiled or executed on Linux, so the Stage-A toolchain reads it as
text. Doing that naively produces false results in both directions, and this
module exists to avoid both:

  * A comment explaining why ``Worksheet_Change`` is absent must not be read as a
    declaration of ``Worksheet_Change``. Code and commentary are separated first.
  * A constant used inside a string literal is not a reference to that constant,
    so string contents are removed before identifiers are collected.

Nothing here claims the VBA is correct. It answers narrow, mechanical questions:
which procedures exist, which identifiers are referenced, and whether a forbidden
construct appears in code rather than in prose.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

PROCEDURE_RE = re.compile(
    r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+\w+)\s+(\w+)",
    re.IGNORECASE,
)
CONST_RE = re.compile(r"^\s*(?:Public\s+|Private\s+)?Const\s+(\w+)", re.IGNORECASE)
IDENTIFIER_RE = re.compile(r"\b[A-Z][A-Z0-9_]{2,}\b")


@dataclass(frozen=True)
class VbaModule:
    name: str
    path: Path
    raw: str

    @property
    def code(self) -> str:
        """The source with comments and string literals removed."""
        return strip_strings(strip_comments(self.raw))

    @property
    def code_without_string_removal(self) -> str:
        return strip_comments(self.raw)

    @property
    def public_procedures(self) -> list[str]:
        found = []
        for line in self.code_without_string_removal.splitlines():
            match = PROCEDURE_RE.match(line)
            if match and (match.group(1) or "").lower() != "private":
                found.append(match.group(3))
        return found

    @property
    def procedures(self) -> list[str]:
        return [
            match.group(3)
            for line in self.code_without_string_removal.splitlines()
            if (match := PROCEDURE_RE.match(line))
        ]

    @property
    def constants(self) -> list[str]:
        return [
            match.group(1)
            for line in self.code_without_string_removal.splitlines()
            if (match := CONST_RE.match(line))
        ]

    @property
    def referenced_upper_identifiers(self) -> set[str]:
        """SCREAMING_CASE identifiers used in code, excluding those declared here."""
        declared = set(self.constants)
        return {name for name in IDENTIFIER_RE.findall(self.code)} - declared


def strip_comments(source: str) -> str:
    """Remove VBA comments, respecting double-quoted string literals.

    A naive split on ``'`` would truncate ``"don't"`` and would also miss that an
    apostrophe inside a string is not a comment marker.
    """
    out_lines: list[str] = []
    for line in source.splitlines():
        in_string = False
        cut = len(line)
        index = 0
        while index < len(line):
            char = line[index]
            if char == '"':
                # A doubled quote inside a string is an escaped quote, not a close.
                if in_string and index + 1 < len(line) and line[index + 1] == '"':
                    index += 2
                    continue
                in_string = not in_string
            elif char == "'" and not in_string:
                cut = index
                break
            index += 1
        out_lines.append(line[:cut])
    return "\n".join(out_lines)


def strip_strings(source: str) -> str:
    """Replace the contents of every double-quoted literal with an empty string."""
    return re.sub(r'"(?:[^"]|"")*"', '""', source)


def load_modules(directories: list[Path]) -> list[VbaModule]:
    """Every .bas module found in *directories*, sorted by name."""
    modules: list[VbaModule] = []
    for directory in directories:
        if not directory.is_dir():
            continue
        for path in sorted(directory.glob("*.bas")):
            modules.append(
                VbaModule(name=path.stem, path=path, raw=path.read_text(encoding="utf-8"))
            )
    return sorted(modules, key=lambda m: m.name)


def contains_construct(modules: list[VbaModule], construct: str) -> list[str]:
    """Modules whose CODE contains *construct*. Commentary never counts."""
    return [m.name for m in modules if construct.lower() in m.code.lower()]
