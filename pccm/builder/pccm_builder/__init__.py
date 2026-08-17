"""PCCM workbook builder (Stage A, Linux/Python).

Generates the Stage A workbook from three source specifications:
  spec/workbook.yaml         structure and presentation
  spec/input_contract.yaml   Setup inputs and Config master lists
  spec/driver_contract.yaml  Cost Line and Risk Register schemas

The generated workbook is a build artifact, never a source of truth.

The public surface is limited to exactly what build_stage_a.py and the test
suites import. Internal types remain reachable through their own modules but are
not re-exported here.
"""

from .contract_loader import ContractError, load_contract
from .driver_loader import DriverContractError, load_driver_contract
from .spec_loader import SpecError, load_spec
from .verify import structural_digest, verify_workbook
from .workbook_builder import BUILDER_VERSION, build_workbook

__all__ = [
    "BUILDER_VERSION",        # build_stage_a.py
    "ContractError",          # build_stage_a.py, input contract tests
    "DriverContractError",    # build_stage_a.py, driver contract tests
    "SpecError",              # build_stage_a.py, manifest tests
    "build_workbook",         # build_stage_a.py, all structural test suites
    "load_contract",          # build_stage_a.py, all test suites
    "load_driver_contract",   # build_stage_a.py, Phase 3 tests
    "load_spec",              # build_stage_a.py, all test suites
    "structural_digest",      # reproducibility test
    "verify_workbook",        # build_stage_a.py
]
