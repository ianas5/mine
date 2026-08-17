"""Workbook-level defined names.

Two families, both derived from the input contract so no address is duplicated:

  inp*  a single Setup input cell
  lst*  the data body of a Config list-master table

On lst* ranges: Excel data validation cannot reference a structured table
reference (``tblCurrencies[Currency]``) and cannot reference another sheet
without a defined name. A range defined name over the table's data body is the
supported mechanism, so the table remains the semantic source of truth and the
defined name is the compatibility shim the builder keeps in sync with it.

The range covers the table's full reserved data body, including rows the user has
not filled yet, so filling a blank row does not require the name to be re-pointed.
"""

from __future__ import annotations

from openpyxl.workbook import Workbook
from openpyxl.workbook.defined_name import DefinedName

from .contract_loader import InputContract


def apply_defined_names(workbook: Workbook, contract: InputContract) -> dict[str, str]:
    """Create every contract-declared defined name. Returns name -> reference."""
    created: dict[str, str] = {}

    for name, reference in sorted(contract.input_defined_names.items()):
        _add(workbook, name, reference)
        created[name] = reference

    for name, reference in sorted(contract.list_defined_names.items()):
        _add(workbook, name, reference)
        created[name] = reference

    return created


def _add(workbook: Workbook, name: str, reference: str) -> None:
    if name in workbook.defined_names:
        raise ValueError(f"defined name {name!r} already exists in the workbook")
    workbook.defined_names.add(DefinedName(name=name, attr_text=reference))
