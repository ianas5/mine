#!/usr/bin/env python3
"""PCCM Phase 1 structural tests.

Runs standalone (``python pccm/tests/test_phase1_structure.py``) and is also
collectable by pytest if it is available.

The locked sheet order, visibility and intended CodeNames are duplicated here on
purpose. The manifest is the builder's structural authority; this file is the
independent architecture-conformance check, so a manifest that drifts from the
Architecture Lock fails here even though the build itself succeeded.
"""

from __future__ import annotations

import os
import sys
import tempfile
import zipfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from openpyxl import load_workbook  # noqa: E402

from pccm_builder import build_workbook, load_spec, structural_digest  # noqa: E402

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"

# --- Architecture Lock Revision B -------------------------------------------
LOCKED_SHEET_ORDER = [
    "Dashboard",
    "Setup",
    "Config",
    "Cost Lines",
    "Risk Register",
    "Inflation",
    "Cost Profiling",
    "Risk Profiling",
    "Model Check",
    "Results",
    "Sensitivity",
    "Methodology",
    "_Calc",
    "_SimData",
]

LOCKED_CODENAMES = {
    "Dashboard": "shDashboard",
    "Setup": "shSetup",
    "Config": "shConfig",
    "Cost Lines": "shCostLines",
    "Risk Register": "shRiskRegister",
    "Inflation": "shInflation",
    "Cost Profiling": "shCostProfiling",
    "Risk Profiling": "shRiskProfiling",
    "Model Check": "shModelCheck",
    "Results": "shResults",
    "Sensitivity": "shSensitivity",
    "Methodology": "shMethodology",
    "_Calc": "shCalc",
    "_SimData": "shSimData",
}

HIDDEN_SHEET = "_Calc"
VERY_HIDDEN_SHEET = "_SimData"
INTERNAL_SHEETS = {HIDDEN_SHEET, VERY_HIDDEN_SHEET}
ACTIVE_SHEET = "Dashboard"

REQUIRED_SETUP_LABELS = [
    "Project Name",
    "Project Duration (Years)",
    "Project Start Year",
    "Base Year",
    "Reporting Currency",
    "Selected Confidence Level",
    "Monte Carlo Iterations",
    "Discount Rate",
    "Random Seed",
]

REQUIRED_CONFIG_SECTIONS = [
    "Categories",
    "Currencies",
    "Units of Measure",
    "Inflation Profile Names",
    "Distribution Names",
    "Confidence Levels",
]

LOCKED_DISTRIBUTIONS = ["Triangular", "Beta-PERT", "Uniform"]
LOCKED_CONFIDENCE_LEVELS = [f"P{n}" for n in range(50, 100, 5)]

_ARTIFACT_CACHE: dict[str, Path] = {}
_TEMPDIR: tempfile.TemporaryDirectory | None = None


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _build(name: str, timestamp: str = "1970-01-01 00:00:00 UTC") -> Path:
    """Build the skeleton into a temp directory and return its path."""
    global _TEMPDIR
    if name in _ARTIFACT_CACHE:
        return _ARTIFACT_CACHE[name]
    if _TEMPDIR is None:
        _TEMPDIR = tempfile.TemporaryDirectory(prefix="pccm-phase1-")

    previous = os.environ.get("PCCM_BUILD_TIMESTAMP")
    os.environ["PCCM_BUILD_TIMESTAMP"] = timestamp
    try:
        spec = load_spec(SPEC_PATH)
        workbook, _ = build_workbook(spec)
        path = Path(_TEMPDIR.name) / f"{name}.xlsx"
        workbook.save(path)
        workbook.close()
    finally:
        if previous is None:
            os.environ.pop("PCCM_BUILD_TIMESTAMP", None)
        else:
            os.environ["PCCM_BUILD_TIMESTAMP"] = previous

    _ARTIFACT_CACHE[name] = path
    return path


def _artifact() -> Path:
    return _build("primary")


def _strings(worksheet) -> set[str]:
    values: set[str] = set()
    for row in worksheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                values.add(cell.value.strip())
    return values


# ---------------------------------------------------------------------------
# 1-3, 8. sheet inventory
# ---------------------------------------------------------------------------
def test_01_exactly_fourteen_worksheets() -> None:
    workbook = load_workbook(_artifact())
    assert len(workbook.sheetnames) == 14, f"found {len(workbook.sheetnames)}"


def test_02_sheet_names_match_locked_list() -> None:
    workbook = load_workbook(_artifact())
    assert set(workbook.sheetnames) == set(LOCKED_SHEET_ORDER)


def test_03_sheet_order_matches_locked_list() -> None:
    workbook = load_workbook(_artifact())
    assert workbook.sheetnames == LOCKED_SHEET_ORDER


def test_08_no_unexpected_worksheet() -> None:
    workbook = load_workbook(_artifact())
    unexpected = set(workbook.sheetnames) - set(LOCKED_SHEET_ORDER)
    assert not unexpected, f"unexpected sheets: {sorted(unexpected)}"
    assert "Sheet" not in workbook.sheetnames, "default openpyxl sheet was not removed"


# ---------------------------------------------------------------------------
# 4-7, 18. visibility and active sheet
# ---------------------------------------------------------------------------
def test_04_dashboard_is_active() -> None:
    workbook = load_workbook(_artifact())
    assert workbook.active is not None
    assert workbook.active.title == ACTIVE_SHEET


def test_05_calc_is_hidden() -> None:
    workbook = load_workbook(_artifact())
    assert workbook[HIDDEN_SHEET].sheet_state == "hidden"


def test_06_simdata_is_very_hidden() -> None:
    workbook = load_workbook(_artifact())
    assert workbook[VERY_HIDDEN_SHEET].sheet_state == "veryHidden"


def test_07_all_other_sheets_are_visible() -> None:
    workbook = load_workbook(_artifact())
    for name in LOCKED_SHEET_ORDER:
        if name in INTERNAL_SHEETS:
            continue
        assert workbook[name].sheet_state == "visible", (
            f"{name} is {workbook[name].sheet_state}"
        )


def test_18_internal_sheets_are_not_active() -> None:
    workbook = load_workbook(_artifact())
    assert workbook.active.title not in INTERNAL_SHEETS
    assert workbook.active.sheet_state == "visible"


# ---------------------------------------------------------------------------
# 9-10. intended CodeNames (manifest only; Stage B is the runtime authority)
# ---------------------------------------------------------------------------
def test_09_manifest_has_one_intended_codename_per_sheet() -> None:
    spec = load_spec(SPEC_PATH)
    assert len(spec.sheets) == len(LOCKED_SHEET_ORDER)
    for sheet in spec.sheets:
        assert sheet.codename, f"{sheet.name} has no intended CodeName"
        assert sheet.codename == LOCKED_CODENAMES[sheet.name], (
            f"{sheet.name}: manifest says {sheet.codename}, "
            f"lock says {LOCKED_CODENAMES[sheet.name]}"
        )


def test_10_intended_codenames_are_unique() -> None:
    spec = load_spec(SPEC_PATH)
    codenames = [s.codename for s in spec.sheets]
    assert len(codenames) == len(set(codenames)), "duplicate intended CodeNames"


# ---------------------------------------------------------------------------
# 11-12. artifact integrity
# ---------------------------------------------------------------------------
def test_11_workbook_reopens_with_openpyxl() -> None:
    workbook = load_workbook(_artifact())
    assert workbook.sheetnames == LOCKED_SHEET_ORDER
    workbook.close()


def test_12_artifact_is_a_genuine_ooxml_package() -> None:
    path = _artifact()
    assert path.suffix == ".xlsx"
    assert zipfile.is_zipfile(path), "artifact is not a ZIP container"
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
    assert "[Content_Types].xml" in names
    assert "xl/workbook.xml" in names


# ---------------------------------------------------------------------------
# 13-14. Phase 1 shells
# ---------------------------------------------------------------------------
def test_13_setup_shell_contains_required_labels() -> None:
    workbook = load_workbook(_artifact())
    values = _strings(workbook["Setup"])
    missing = [label for label in REQUIRED_SETUP_LABELS if label not in values]
    assert not missing, f"missing Setup labels: {missing}"
    assert "SAR" in values, "Setup does not show the locked reporting currency"


def test_14_config_shell_contains_reserved_sections() -> None:
    workbook = load_workbook(_artifact())
    values = _strings(workbook["Config"])
    missing = [name for name in REQUIRED_CONFIG_SECTIONS if name not in values]
    assert not missing, f"missing Config sections: {missing}"
    for item in LOCKED_DISTRIBUTIONS + LOCKED_CONFIDENCE_LEVELS:
        assert item in values, f"Config is missing locked constant {item!r}"


# ---------------------------------------------------------------------------
# 15-17. nothing from a later phase has leaked in
# ---------------------------------------------------------------------------
def test_15_no_vba_project_present() -> None:
    with zipfile.ZipFile(_artifact()) as archive:
        names = [n.lower() for n in archive.namelist()]
    assert not any(n.endswith("vbaproject.bin") for n in names)
    assert not any(n.endswith(".xlsm") for n in names)


def test_16_no_calculation_formulas_introduced() -> None:
    workbook = load_workbook(_artifact())
    offenders = []
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    offenders.append(f"{worksheet.title}!{cell.coordinate}")
    assert not offenders, f"formulas present: {offenders[:10]}"
    assert len(list(workbook.defined_names)) == 0, "defined names present"
    for worksheet in workbook.worksheets:
        assert not getattr(worksheet, "tables", {}), f"{worksheet.title} declares a table"


def test_17_no_external_links_or_connections() -> None:
    with zipfile.ZipFile(_artifact()) as archive:
        names = archive.namelist()
    assert not [n for n in names if "externalLink" in n]
    assert not [n for n in names if "connections" in n.lower()]
    assert not [n for n in names if "queryTable" in n]


# ---------------------------------------------------------------------------
# structural reproducibility (logical, not byte-identical)
# ---------------------------------------------------------------------------
def test_19_structurally_reproducible() -> None:
    first = _build("repro_a", timestamp="2000-01-01 00:00:00 UTC")
    second = _build("repro_b", timestamp="2000-01-01 00:00:00 UTC")
    assert structural_digest(first) == structural_digest(second), (
        "two builds of the same source produced different workbook structure"
    )


def test_21_fx_single_source_of_truth() -> None:
    """Setup owns FX rates; Config owns the currency master list and no rates.

    Guards the Phase 2 boundary: exactly one FX authority. Config must never
    grow a rate table, so it is asserted to contain no FX token and no numeric
    value at all.
    """
    workbook = load_workbook(_artifact())
    config_sheet = workbook["Config"]
    config = _strings(config_sheet)
    setup = _strings(workbook["Setup"])

    # Config owns the currency master list...
    assert "Currencies" in config, "Config is missing the Currencies section"
    assert "Currencies / FX" not in config, "Config still uses the ambiguous 'Currencies / FX'"

    # ...and owns no FX authority.
    fx_mentions = sorted(v for v in config if "fx" in v.lower())
    assert not fx_mentions, f"Config must not mention FX: {fx_mentions}"

    numeric = [
        f"{cell.coordinate}={cell.value!r}"
        for row in config_sheet.iter_rows()
        for cell in row
        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
    ]
    assert not numeric, f"Config contains numeric values (possible rate area): {numeric}"

    # Setup is the single source of truth for FX rates, and states the convention.
    assert "FX Rates" in setup, "Setup is missing the FX Rates section"
    assert "FX Rates" not in config, "FX Rates must not appear on Config"
    assert any("1 unit of the source currency = X SAR" in v for v in setup), (
        "Setup does not state the SAR-per-source-unit FX convention"
    )


def test_20_manifest_locked_order_matches_architecture() -> None:
    spec = load_spec(SPEC_PATH)
    assert spec.workbook["locked_sheet_order"] == LOCKED_SHEET_ORDER
    assert spec.sheet_names == LOCKED_SHEET_ORDER
    assert spec.active_sheet == ACTIVE_SHEET


# ---------------------------------------------------------------------------
# standalone runner
# ---------------------------------------------------------------------------
def _run_all() -> int:
    tests = sorted(
        (name, fn)
        for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures: list[tuple[str, str]] = []
    print("PCCM Phase 1 structural tests")
    print("=" * 66)
    for name, fn in tests:
        try:
            fn()
        except AssertionError as error:
            failures.append((name, str(error) or "assertion failed"))
            print(f"  [FAIL] {name}")
            print(f"         {error}")
        except Exception as error:  # noqa: BLE001 - report any unexpected error
            failures.append((name, f"{type(error).__name__}: {error}"))
            print(f"  [ERROR] {name}")
            print(f"          {type(error).__name__}: {error}")
        else:
            print(f"  [PASS] {name}")
    print("=" * 66)
    print(f"  {len(tests) - len(failures)} passed, {len(failures)} failed")
    if _TEMPDIR is not None:
        _TEMPDIR.cleanup()
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(_run_all())
