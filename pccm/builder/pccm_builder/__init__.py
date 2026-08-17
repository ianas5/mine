"""PCCM workbook builder (Stage A, Linux/Python).

Generates the structural workbook skeleton from the manifest at pccm/spec/.
The generated workbook is a build artifact, never a source of truth.

The public surface is deliberately limited to exactly what build_skeleton.py and
the Phase 1 tests import. Internal types (SheetSpec, WorkbookSpec, StyleBook,
BuildMetadata, VerificationResult) remain reachable through their own modules but
are not re-exported here.
"""

from .skeleton import BUILDER_VERSION, build_workbook
from .spec_loader import SpecError, load_spec
from .verify import structural_digest, verify_workbook

__all__ = [
    "BUILDER_VERSION",      # build_skeleton.py
    "SpecError",            # build_skeleton.py, manifest validation tests
    "build_workbook",       # build_skeleton.py, structural tests
    "load_spec",            # build_skeleton.py, both test suites
    "structural_digest",    # structural tests (reproducibility)
    "verify_workbook",      # build_skeleton.py
]
