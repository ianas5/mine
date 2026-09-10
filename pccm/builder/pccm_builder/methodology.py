"""P10-3: the Methodology sheet.

WHAT THIS SHEET IS. The one place inside the workbook that explains the method
to the person reading the answers. It is written for that reader and for no one
else: no module name, no procedure name, no cell address and no VBA appears on
it, and the two identities it states in symbols are sentences in cells rather
than formulas.

WHAT IT IS NOT, AND THIS IS THE WHOLE OF IT:

  it is not an authority   Every sentence describes a decision some accepted
                           owner already made. If a sentence here and the engine
                           ever disagreed, the engine would be right and the
                           sentence would be a documentation defect. Nothing in
                           the builder or the VBA reads a word of it.

  it is not calculated     Not one cell on the sheet holds a formula. The two
                           formula-looking lines are text, and the manifest
                           loader refuses any methodology string that Excel
                           would read as a formula at all.

  it is not typed here     Every word comes from the manifest's `methodology`
                           block. This module owns the LAYOUT - which row a
                           thing lands on and which font it wears - and owns no
                           wording whatsoever. A sentence is changed by editing
                           the specification, where the controls read it.

THE LAYOUT IS PLANNED BEFORE IT IS WRITTEN. `plan_methodology` returns the
whole sheet as an ordered list of lines; the renderer walks it and the
projection artefact serialises it. That is deliberate: a Windows run that wants
to confirm the sheet says what the manifest says can read the projection rather
than re-deriving the geometry, and there is exactly one description of where
everything sits.
"""
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

__all__ = [
    "MethodologyLine",
    "plan_methodology",
    "render_methodology",
    "emit_methodology_projection",
]


@dataclass(frozen=True)
class MethodologyLine:
    """One rendered row: what it is, what stands in the two columns, whose it is.

    `label` is the left rail - a section tag, a defined term or a metadata
    label - and `text` is the prose column. Either may be absent; a line with
    neither is a deliberate blank.
    """

    row: int
    kind: str
    key: str | None = None
    label: str | None = None
    text: str | None = None

    def as_dict(self) -> dict[str, Any]:
        entry: dict[str, Any] = {"row": self.row, "kind": self.kind}
        if self.key is not None:
            entry["key"] = self.key
        if self.label is not None:
            entry["label"] = self.label
        if self.text is not None:
            entry["text"] = self.text
        return entry


# The font each kind wears. A kind that is not named here is a blank line.
# `term` and `metadata_label` are the only two that put anything in the left
# rail beside a section tag.
_LABEL_FONT = {
    "section": "section",
    "term": "term",
    "metadata_heading": "section",
    "metadata_value": "label",
}
_TEXT_FONT = {
    "heading": "section",
    "purpose": "value",
    "note": "note",
    "section": "section",
    "summary": "note",
    "paragraph": "value",
    "term": "value",
    "formula_caption": "note",
    "formula": "value_locked",
    "formula_note": "note",
    "section_closing": "note",
    "metadata_note": "note",
    "metadata_value": "value",
    "closing": "note",
}
# Rows that carry a heading get the manifest's section height. Everything else
# is left to Excel, which auto-fits a wrapped paragraph to however many lines it
# actually needs - a height typed here would clip the longest sentence on the
# sheet the first time anyone rewrote it.
_HEADING_KINDS = ("heading", "section", "metadata_heading")


def plan_methodology(spec: Any, metadata: Any) -> list[MethodologyLine]:
    """The whole sheet, in order, as rows that have not been written yet."""
    block = spec.methodology
    if not block:
        return []

    lines: list[MethodologyLine] = []
    row = int(spec.presentation["layout"]["body_start_row"])

    def emit(kind: str, *, key: str | None = None, label: str | None = None,
             text: str | None = None) -> None:
        nonlocal row
        lines.append(MethodologyLine(row=row, kind=kind, key=key, label=label,
                                     text=text))
        row += 1

    def blank() -> None:
        emit("blank")

    emit("heading", text=block["heading"])
    emit("purpose", text=block["purpose"])
    emit("note", text=block["note"])
    blank()

    for section in block["sections"]:
        key = section["key"]
        emit("section", key=key, label=key, text=section["title"])
        emit("summary", key=key, text=section["summary"])
        for paragraph in section["paragraphs"]:
            emit("paragraph", key=key, text=paragraph)
        for term in section["terms"]:
            emit("term", key=key, label=term["term"], text=term["text"])
        formula = section.get("formula")
        if formula is not None:
            emit("formula_caption", key=key, text=formula["caption"])
            emit("formula", key=key, text=formula["text"])
            emit("formula_note", key=key, text=formula["note"])
        if section.get("closing"):
            emit("section_closing", key=key, text=section["closing"])
        blank()

    # THE BUILD METADATA BLOCK, and it stays on this sheet because this is where
    # it has always been. Every value comes from `BuildMetadata`, which resolves
    # them from the manifest and the contracts; not one of them is restated here.
    emit("metadata_heading", label=block["metadata_heading"])
    emit("metadata_note", text=block["metadata_note"])
    for label, value in metadata.as_rows():
        emit("metadata_value", label=label, text=value)
    blank()

    emit("closing", text=block["closing"])
    return lines


def render_methodology(worksheet: Any, spec: Any, styles: Any, metadata: Any) -> None:
    """Write the planned sheet. No wording, no address and no formula is typed."""
    lines = plan_methodology(spec, metadata)
    if not lines:
        return

    label_column = styles.layout.label_column
    text_column = styles.layout.value_column

    for line in lines:
        if line.label is not None:
            cell = worksheet[f"{label_column}{line.row}"]
            cell.value = line.label
            cell.font = getattr(styles, _LABEL_FONT[line.kind])
            cell.alignment = styles.prose
        if line.text is not None:
            cell = worksheet[f"{text_column}{line.row}"]
            cell.value = line.text
            cell.font = getattr(styles, _TEXT_FONT[line.kind])
            cell.alignment = styles.prose
        if line.kind in _HEADING_KINDS:
            worksheet.row_dimensions[line.row].height = styles.row_height("section")


def emit_methodology_projection(path: Path, spec: Any, metadata: Any) -> dict[str, Any]:
    """Where every methodology line landed, and what it says.

    ADDRESSES AND THE TEXT THAT GOES WITH THEM. A later Windows run reads this
    to confirm that the sheet in the saved workbook carries the manifest's
    wording at the manifest's geometry. It asserts nothing itself.
    """
    lines = plan_methodology(spec, metadata)
    block = spec.methodology
    projection = {
        "sheet": block.get("sheet"),
        "label_column": spec.presentation["layout"]["label_column"],
        "text_column": spec.presentation["layout"]["value_column"],
        "sections": [
            {"key": section["key"], "title": section["title"],
             "row": next(line.row for line in lines
                         if line.kind == "section" and line.key == section["key"])}
            for section in block.get("sections", [])
        ],
        "metadata": [
            {"row": line.row, "label": line.label, "value": line.text}
            for line in lines if line.kind == "metadata_value"
        ],
        "lines": [line.as_dict() for line in lines if line.kind != "blank"],
    }
    path.write_text(json.dumps(projection, indent=2) + "\n", encoding="utf-8")
    return projection
