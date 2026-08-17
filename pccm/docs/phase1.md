# PCCM — Phase 1: Source Foundation + Workbook Skeleton

**Status: complete, pending review.**

Phase 1 establishes a reproducible, source-controlled foundation and generates
the structural workbook shell that every later phase builds on. It deliberately
contains **no model logic**.

The functional design is fixed by the **Architecture Lock, Revision B**, which
remains authoritative. This document does not restate it.

---

## Two-stage build architecture

| Stage | Runs on | Produces | Status |
|---|---|---|---|
| **A** | Linux + Python 3.11 + openpyxl | `PCCM_skeleton.xlsx` — sheets, order, visibility, presentation, structural shells | **This phase** |
| **B** | Windows + Excel COM (PowerShell 5.1) | the final `.xlsm` — VBA import, CodeName assignment, Excel-runtime objects, save/reopen/verify | Later phase |

Stage B viability is already proven independently: the Excel COM readiness gate
passed all 15 tests on the target machine (`pccm/readiness/windows/`). That
readiness script is a disposable diagnostic and is **not** production build code;
the production bootstrap will live in `pccm/bootstrap/windows/`.

Stage B is **not** exercised in Phase 1, and nothing here needs to be run on
Windows yet.

## Source-of-truth principle

The source of truth is **code, specification, configuration and tests**. The
workbook is a *generated build artifact*.

- `pccm/spec/workbook.yaml` is the structural authority. The builder reads it and
  hardcodes no sheet names, order, visibility or style literals.
- `pccm/build/` is git-ignored. A clean checkout regenerates the skeleton.
- An `.xlsx`/`.xlsm` edited by hand is never source. Any structural change is
  made in the manifest and rebuilt.

---

## Repository layout

    pccm/
      VERSION                         model version (matches spec model_version)
      requirements.txt                pinned Stage A dependencies
      spec/workbook.yaml              structural manifest (the authority)
      builder/
        build_skeleton.py             entry point
        pccm_builder/
          spec_loader.py              parse + validate the manifest; fails loudly
          styling.py                  centralised presentation tokens
          skeleton.py                 workbook + sheet creation, population
          verify.py                   structural verification, structural digest
          __init__.py                 public API: exactly the 6 objects the entry point and tests use
      tests/
        test_phase1_structure.py      18 required structural checks, reproducibility, FX ownership
        test_phase1_manifest_validation.py   invalid manifests must be rejected
        oracle/                       (empty) Python reference implementation
        fixtures/                     (empty) expected-value vectors
      src/vba/                        (empty) authoritative VBA source
      bootstrap/windows/              (empty) production Stage B bootstrap
      docs/                           this file
      build/                          generated artifacts (git-ignored)
      readiness/windows/              disposable Excel COM smoke test (gate closed)

## Build

    python3 -m pip install -r pccm/requirements.txt
    python3 pccm/builder/build_skeleton.py

Produces `pccm/build/PCCM_skeleton.xlsx` and runs structural verification against
the manifest that produced it. Non-zero exit on a specification error (2) or a
verification failure (1).

Options: `--spec PATH`, `--out PATH`, `--quiet`.

### Reproducible builds

Set `PCCM_BUILD_TIMESTAMP` to pin the build stamp:

    PCCM_BUILD_TIMESTAMP="2026-01-01 00:00:00 UTC" python3 pccm/builder/build_skeleton.py

Two builds of the same source then produce structurally identical workbooks.
Reproducibility is asserted **logically**, by comparing a normalised structural
digest — sheet order, visibility, gridlines, freeze panes, column widths and every
non-empty cell value. Byte-identical `.xlsx` files are *not* required, because ZIP
member ordering and archive metadata legitimately vary.

## Tests

    python3 pccm/tests/test_phase1_structure.py
    python3 pccm/tests/test_phase1_manifest_validation.py

Both run standalone (no pytest required) and are also pytest-collectable.

`test_phase1_structure.py` re-declares the locked sheet order, visibility and
intended CodeNames **independently of the manifest**. The manifest is the
builder's authority; the test is the architecture-conformance check, so manifest
drift fails even when the build itself succeeds.

---

## FX source of truth

One authority each, fixed before Phase 2 so the boundary cannot drift:

| Concern | Owner | Notes |
|---|---|---|
| Currency **master list** | **Config** → *Currencies* | user-maintainable; holds no rates |
| **FX rates** | **Setup** → *FX Rates* | convention: **1 unit of the source currency = X SAR** |

Cost Lines and Risks will later select a currency from the Config master list; the
corresponding rate is resolved from Setup. There is no FX forecasting and no FX
uncertainty in this model.

`test_21_fx_single_source_of_truth` enforces this: Config must contain the
*Currencies* section, must not mention FX at all, and must contain no numeric
value (so it can never quietly become a rate table); Setup must contain the *FX
Rates* section and state the convention.

---

## Versioning

Three version concepts are kept distinct and must not be conflated:

| Concept | Meaning | Where |
|---|---|---|
| **Model version** | version of the model design | `spec/workbook.yaml` → `model.model_version`; `pccm/VERSION` |
| **Builder version** | version of this build tooling | `pccm_builder/skeleton.py` → `BUILDER_VERSION` |
| **Manifest version** | version of the specification format | `spec/workbook.yaml` → `manifest_version` |

Separately, **run metadata** (run id, seed, iterations, timestamp of a Monte
Carlo run) belongs to a simulation run, not to a build. It does not exist yet.

Build metadata is stamped in two controlled places: the workbook document
properties, and a *Build Metadata* block on the Methodology sheet. **The build
timestamp is display-only and is never a computational input.**

---

## What Phase 1 deliberately does NOT implement

None of the following exists in the generated workbook, and their absence is
asserted by the tests:

- Monte Carlo engine, RNG, seed derivation, distribution sampling
- Cost, risk, inflation, FX, escalation, NPV or contingency calculations
- Sensitivity analysis
- Model Check rules, severities or the simulation gate
- Production VBA — the workbook is `.xlsx` with no VBA project
- Dynamic timeline logic (Apply / Update Timeline)
- Stale-result fingerprinting
- Dashboard formulas and charts
- Excel Tables, defined names, data validation, conditional formatting
- Add/Delete Cost Line and Risk macros, buttons and shapes
- Worksheet protection
- Runtime CodeName assignment

On CodeNames specifically: the manifest **records the intended CodeNames** and the
tests verify they are present, unique and match the lock. That is a specification
record only. **Stage B remains the authority for actually assigning and verifying
CodeNames in Excel**, and no claim is made that runtime CodeName persistence has
been implemented in Phase 1.

The sheet shells contain labels, section headings and reserved areas — never
fabricated numbers or placeholder formulas dressed up as working logic.
