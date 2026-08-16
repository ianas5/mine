"""PCCM workbook builder (Stage A, Linux/Python).

Generates the structural workbook skeleton from the manifest at pccm/spec/.
The generated workbook is a build artifact, never a source of truth.
"""

from .skeleton import BUILDER_VERSION, BuildMetadata, build_workbook
from .spec_loader import SheetSpec, SpecError, WorkbookSpec, load_spec
from .styling import StyleBook
from .verify import VerificationResult, structural_digest, verify_workbook

__all__ = [
    "BUILDER_VERSION",
    "BuildMetadata",
    "SheetSpec",
    "SpecError",
    "StyleBook",
    "VerificationResult",
    "WorkbookSpec",
    "build_workbook",
    "load_spec",
    "structural_digest",
    "verify_workbook",
]
