"""PCCM workbook builder (Stage A, Linux/Python).

Generates the Stage A workbook from two source specifications:
  spec/workbook.yaml        structure and presentation
  spec/input_contract.yaml  inputs, list masters, tables, names, validation

The generated workbook is a build artifact, never a source of truth.

The public surface is deliberately limited to exactly what build_stage_a.py and
the test suites import. Internal types remain reachable through their own
modules but are not re-exported here.
"""

from .contract_loader import ContractError, load_contract
from .spec_loader import SpecError, load_spec
from .verify import structural_digest, verify_workbook
from .workbook_builder import BUILDER_VERSION, build_workbook

__all__ = [
    "BUILDER_VERSION",      # build_stage_a.py
    "ContractError",        # build_stage_a.py, contract validation tests
    "SpecError",            # build_stage_a.py, manifest validation tests
    "build_workbook",       # build_stage_a.py, structure and Phase 2 tests
    "load_contract",        # build_stage_a.py, all test suites
    "load_spec",            # build_stage_a.py, all test suites
    "structural_digest",    # reproducibility test
    "verify_workbook",      # build_stage_a.py
]
