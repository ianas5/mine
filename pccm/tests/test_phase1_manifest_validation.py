#!/usr/bin/env python3
"""PCCM Phase 1 manifest validation tests.

The builder must fail loudly on a bad specification and must never repair one
silently. Each case here mutates a copy of the real manifest into an invalid
state and asserts that loading it raises SpecError.

Runs standalone or under pytest.
"""

from __future__ import annotations

import copy
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable

import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import SpecError, load_spec  # noqa: E402

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"


def _base() -> dict[str, Any]:
    with SPEC_PATH.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def _assert_rejected(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    data = copy.deepcopy(_base())
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-badspec-") as tmp:
        path = Path(tmp) / "broken.yaml"
        with path.open("w", encoding="utf-8") as handle:
            yaml.safe_dump(data, handle, sort_keys=False)
        try:
            load_spec(path)
        except SpecError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of SpecError"
            ) from error
    raise AssertionError(f"{reason}: invalid manifest was silently accepted")


def test_rejects_sheet_order_drift() -> None:
    _assert_rejected(
        lambda d: d["sheets"].insert(0, d["sheets"].pop(1)),
        "sheet order drifted from the architecture lock",
    )


def test_rejects_duplicate_codename() -> None:
    _assert_rejected(
        lambda d: d["sheets"][1].__setitem__("codename", d["sheets"][0]["codename"]),
        "duplicate intended CodeName",
    )


def test_rejects_hidden_sheet_as_active() -> None:
    _assert_rejected(
        lambda d: d["workbook"].__setitem__("active_sheet", "_Calc"),
        "a hidden sheet must never be the active sheet",
    )


def test_rejects_unknown_active_sheet() -> None:
    _assert_rejected(
        lambda d: d["workbook"].__setitem__("active_sheet", "Nonexistent"),
        "active sheet not present in the sheet list",
    )


def test_rejects_invalid_visibility() -> None:
    _assert_rejected(
        lambda d: d["sheets"][0].__setitem__("visibility", "sometimes"),
        "invalid visibility value",
    )


def test_rejects_malformed_codename() -> None:
    _assert_rejected(
        lambda d: d["sheets"][0].__setitem__("codename", "Dashboard"),
        "CodeName not matching the sh<PascalCase> convention",
    )


def test_rejects_missing_required_key() -> None:
    _assert_rejected(
        lambda d: d["model"].pop("model_version"),
        "missing required model key",
    )


def test_rejects_unknown_block_type() -> None:
    _assert_rejected(
        lambda d: d["sheets"][0]["blocks"].append({"type": "chart"}),
        "unknown block type",
    )


def test_rejects_missing_manifest_file() -> None:
    try:
        load_spec(PCCM_ROOT / "spec" / "does_not_exist.yaml")
    except SpecError:
        return
    raise AssertionError("a missing manifest was silently accepted")


def test_valid_manifest_loads() -> None:
    spec = load_spec(SPEC_PATH)
    assert len(spec.sheets) == 14
    assert spec.active_sheet == "Dashboard"


def _run_all() -> int:
    tests = sorted(
        (name, fn)
        for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 1 manifest validation tests")
    print("=" * 66)
    for name, fn in tests:
        try:
            fn()
        except AssertionError as error:
            failures += 1
            print(f"  [FAIL] {name}\n         {error}")
        except Exception as error:  # noqa: BLE001
            failures += 1
            print(f"  [ERROR] {name}\n          {type(error).__name__}: {error}")
        else:
            print(f"  [PASS] {name}")
    print("=" * 66)
    print(f"  {len(tests) - failures} passed, {failures} failed")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(_run_all())
