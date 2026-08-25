"""PCCM workbook builder (Stage A, Linux/Python).

Generates the Stage A workbook from four source specifications:
  spec/workbook.yaml            structure and presentation
  spec/input_contract.yaml      Setup inputs and Config master lists
  spec/driver_contract.yaml     Cost Line and Risk Register schemas
  spec/structure_contract.yaml  structural runtime: applied timeline, grids, identity

and emits the two Stage-B inputs those contracts project into languages that
cannot read YAML: build/vba/modConstants.bas and build/stage_b_manifest.json.

The generated workbook is a build artifact, never a source of truth.

A fifth specification, spec/calc_contract.yaml, owns the physical shape of the
Phase-5 `_Calc` workspace. As of Phase 5 Gate-A Step 3 it is a REAL BUILD INPUT:
loaded, cross-validated, projected into the workbook as empty representation, and
checked in the generated artifact. Nothing in Stage A calculates - the analytical
oracle is called only to produce expected values for build/phase5_cases.json, and
never to populate a cell. calc_fingerprint.py is the reference implementation of
the Calculation Input Fingerprint and owns the hash mathematics outright.

The public surface is limited to exactly what build_stage_a.py and the test
suites import. Internal types remain reachable through their own modules but are
not re-exported here.
"""

from .calc_emit import emit_calc_artifacts
from .gate_b_inspection import emit_inspection
from .calc_loader import (
    CalcContractError,
    load_calc_contract,
    validate_calc_against,
)
from .contract_loader import ContractError, load_contract
from .driver_loader import DriverContractError, load_driver_contract
from .sim_loader import (
    SimContractError,
    load_sim_contract,
    validate_sim_against,
)
from .sim_rng import (
    Component,
    Draw,
    RngReference,
    RngState,
    SimRngError,
)
from .sim_sample import (
    ACCEPTED_FAMILIES,
    BernoulliResult,
    PreparedBetaPert,
    SampleResult,
    SimSampleError,
    bernoulli_occurs,
    prepare_beta_pert,
    sample_beta_pert,
    sample_distribution,
    sample_prepared_beta,
    sample_triangular,
    sample_uniform,
)
from .spec_loader import SpecError, load_spec
from .stage_b_emit import emit_stage_b
from .structure_loader import StructureContractError, load_structure_contract
from .verify import structural_digest, verify_workbook
from .workbook_builder import BUILDER_VERSION, build_workbook

__all__ = [
    "BUILDER_VERSION",           # build_stage_a.py
    "CalcContractError",         # Phase 5 Gate-A Step-1 tests
    "ContractError",             # build_stage_a.py, input contract tests
    "DriverContractError",       # build_stage_a.py, driver contract tests
    "ACCEPTED_FAMILIES",         # Phase 6 Step-3 reference surface
    "BernoulliResult",           # Phase 6 Step-3 reference surface
    "Component",                 # Phase 6 Step-2 tests
    "PreparedBetaPert",          # Phase 6 Step-3 reference surface
    "SampleResult",              # Phase 6 Step-3 reference surface
    "SimSampleError",            # Phase 6 Step-3 tests
    "bernoulli_occurs",          # Phase 6 Step-3 reference surface
    "prepare_beta_pert",         # Phase 6 Step-3 reference surface
    "sample_beta_pert",          # Phase 6 Step-3 reference surface
    "sample_distribution",       # Phase 6 Step-3 reference surface
    "sample_prepared_beta",      # Phase 6 Step-3 reference surface
    "sample_triangular",         # Phase 6 Step-3 reference surface
    "sample_uniform",            # Phase 6 Step-3 reference surface
    "Draw",                      # Phase 6 Step-2 tests
    "RngReference",              # Phase 6 Step-2 reference surface
    "RngState",                  # Phase 6 Step-2 reference surface
    "SimContractError",          # build_stage_a.py, Phase 6 Step-1 tests
    "SimRngError",               # Phase 6 Step-2 tests
    "SpecError",                 # build_stage_a.py, manifest tests
    "StructureContractError",    # build_stage_a.py, Phase 4 tests
    "build_workbook",            # build_stage_a.py, all structural test suites
    "emit_calc_artifacts",       # build_stage_a.py, Phase 5 Gate-A Step-3 tests
    "emit_inspection",           # build_stage_a.py, Phase 5 Gate-B harness
    "emit_stage_b",              # build_stage_a.py, Phase 4 Stage-B emission tests
    "load_calc_contract",        # Phase 5 Gate-A Step-1 tests
    "load_contract",             # build_stage_a.py, all test suites
    "load_driver_contract",      # build_stage_a.py, Phase 3 and 4 tests
    "load_sim_contract",         # build_stage_a.py, Phase 6 Step-1 tests
    "load_spec",                 # build_stage_a.py, all test suites
    "load_structure_contract",   # build_stage_a.py, Phase 4 tests
    "structural_digest",         # reproducibility test
    "validate_calc_against",     # build_stage_a.py, Phase 5 Gate-A tests
    "validate_sim_against",      # build_stage_a.py, Phase 6 Step-1 tests
    "verify_workbook",           # build_stage_a.py
]
