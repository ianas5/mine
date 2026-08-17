# PCCM — Phase 2: Functional Setup & Configuration Layer

**Status: accepted and closed.** Phase 3 builds on this; see `phase3.md`.

Phase 2 turns Setup and Config from structural placeholders into controlled
model-input infrastructure. It is still a Stage A `.xlsx`; no VBA, no Stage B.

It contains **no business calculations**. Defined names and data validation are
input *infrastructure*, not computational logic.

Authority for functional design remains the **Architecture Lock, Revision B**.
Phase 1 is closed; see `phase1.md` for the foundation this builds on.

---

## Source authorities

| File | Owns |
|---|---|
| `spec/workbook.yaml` | structure and presentation: sheets, order, visibility, style tokens, non-input sheet shells |
| `spec/input_contract.yaml` | **inputs**: semantic keys, labels, cells, defined names, types, defaults, number formats, validation, editability |
| `spec/driver_contract.yaml` | added in Phase 3: Cost Line and Risk Register schemas |

Setup and Config declare `body: contract` in the manifest and therefore have
**no blocks**; the input contract is their single layout authority. The loader
rejects a sheet that declares both, so the two files cannot disagree.

No input address, list value or validation rule appears in Python source.

---

## Setup inputs

Labels in column **B**, values in column **C**, explanatory notes in column **E**.

| Input | Defined name | Cell | Default | Editable | Validation |
|---|---|---|---|---|---|
| Project Name | `inpProjectName` | C9 | blank | yes | none |
| Project Duration (Years) | `inpDurationYears` | C10 | blank | yes | whole ≥ 1, **uncapped** |
| Project Start Year | `inpProjectStartYear` | C11 | blank | yes | none |
| Base Year | `inpBaseYear` | C12 | blank | yes | none |
| Reporting Currency | `inpReportingCurrency` | C15 | **SAR** | **no** | none (no dropdown) |
| Selected Confidence Level | `inpSelectedConfidenceLevel` | C18 | **P50** | yes | list ← `lstConfidenceLevels` |
| Monte Carlo Iterations | `inpMonteCarloIterations` | C19 | **10000** | yes | whole ≥ **1000** |
| Discount Rate | `inpDiscountRate` | C20 | blank | yes | none |
| Random Seed | `inpRandomSeed` | C21 | blank | yes | none |

Deliberate restraint, each of these being a Model Check or RNG concern later:

- **Duration is not capped at 25.** 25 years is a design/benchmark target, not a
  business rule. The advisory belongs to Model Check.
- **Iterations have no upper limit.** 1000 is the locked hard minimum; the
  “below 10000” advisory belongs to Model Check.
- **Base Year ≤ Start Year is not enforced here.** It is a blocking Model Check
  rule, not a keystroke-level restriction.
- **Discount Rate and Random Seed have no range.** The seed's admissible domain
  is fixed when the RNG is implemented; inventing one now would be a guess.
- **Selected Confidence Level is a reporting selector only.** It is wired to
  nothing — not to simulation execution, not to staleness, not to any calculation.

---

## Model invariants: the SAR identity

Three values are **not user data**. They are part of the model and are declared in
`input_contract.yaml → model_invariants`:

| Invariant | Where |
|---|---|
| Reporting Currency = `SAR` | `inpReportingCurrency`, model-controlled |
| Currency master contains `SAR` | `tblCurrencies` data row 1 |
| FX identity is exactly `SAR \| 1` | `tblFXRates` data row 1 |

Row-level ownership is expressed by `locked_seed_rows: N` on a table: the first
*N* data rows are model-controlled. Those rows get the **locked** visual treatment
and are **excluded from user data validation** — attaching an input rule to a cell
the user does not own would misrepresent who controls it.

Concretely, for the FX table at rows 28–39: locked identity `B28:C28`; user
currency validation `B29:B39`; user FX validation `C29:C39`. None of those
addresses is hardcoded — all are derived from the table spec.

Users still add and manage every other currency and rate. Uniqueness of
user-entered currencies is **not** enforced here; that is Model Check's job.

The loader validates each declared identity against the named table's locked seed
rows, so removing, unlocking or altering one fails the build rather than the
review.

## Excel address bounds

A reference is not valid merely because it parses: `XFE1` and `A1048577` both
match a naive pattern and are both outside the grid. A central validator
(`check_cell`, `check_row`, `check_column`) enforces the real limits — maximum row
**1,048,576**, maximum column **XFD (16,384)** — over every contract-owned
coordinate: input label and value cells, table first and last column, table header
and last data row, and every section, note, convention and intro row.

It runs **before rendering**, so an out-of-grid address fails as `ContractError`
with a precise message rather than as an incidental openpyxl exception.

## Config ownership

Six Excel Tables, each with a range defined name for validation:

| Section | Table | Defined name | Kind | Seeded |
|---|---|---|---|---|
| Categories | `tblCategories` | `lstCategories` | user-maintainable | — |
| Currencies | `tblCurrencies` | `lstCurrencies` | user-maintainable, **row 1 locked** | `SAR` (model invariant) |
| Units of Measure | `tblUOM` | `lstUOM` | user-maintainable | — |
| Inflation Profile Names | `tblInflationProfiles` | `lstInflationProfiles` | user-maintainable | — |
| Distribution Names | `tblDistributions` | `lstDistributions` | **locked constant** | Triangular, Beta-PERT, Uniform |
| Confidence Levels | `tblConfidenceLevels` | `lstConfidenceLevels` | **locked constant** | P50 … P95 |

`SAR` is seeded into the currency master because the model's reporting currency
is SAR. No other business values are invented: no categories, no units, no
inflation profile names.

Config holds **profile names only**. Annual inflation rates belong to the
Inflation sheet in a later phase.

---

## FX source of truth

**Setup owns FX rates. Config owns the currency master list. Nothing else.**

`tblFXRates` on Setup, columns `Currency` and `FX to SAR`, twelve reserved rows,
with `SAR = 1` as a **locked model invariant** in data row 1 (see above). The convention is
printed beside the table:

> Convention: 1 unit of source currency = X SAR

`Currency` validates against `lstCurrencies`, so a rate can only be entered for a
currency that exists in the master list. There is no FX forecasting, no FX
uncertainty and no year-by-year FX profile. Resolving a rate for a cost line or
risk belongs to a later phase.

The boundary is tested semantically, not by a blanket rule: `tblCurrencies` must
have exactly one column named `Currency`, `tblFXRates` must own `FX to SAR`, and
exactly one rate-bearing table may exist anywhere in the workbook. A future
Config table carrying legitimate numeric metadata will not trip it.

---

## Defined-name convention

| Prefix | Meaning | Example |
|---|---|---|
| `inp` | one Setup input cell | `inpDiscountRate` → `'Setup'!$C$20` |
| `lst` | data body of a Config list table | `lstCurrencies` → `'Config'!$B$28:$B$37` |
| `tbl` | Excel Table | `tblFXRates` |

No name depends on a row number for its meaning, and the builder derives every
reference from the contract, so an address exists in exactly one place.

**Why `lst*` ranges exist alongside the tables.** Excel data validation cannot
reference a structured table reference (`tblCurrencies[Currency]`) and cannot
reference another sheet without a defined name. A range defined name over the
table's data body is the supported mechanism. The table remains the semantic
source of truth; the defined name is the compatibility shim, kept in sync by the
builder from the same contract entry.

The range covers the table's full reserved data body including not-yet-filled
rows, so filling a blank row needs no re-pointing. Growing a list *beyond* its
reserved rows currently requires widening `data_rows` in the contract and
rebuilding — see limitations.

---

## Visual input language

Centralised in `workbook.yaml → presentation.colors`; nothing else defines a
colour.

| Kind | Treatment |
|---|---|
| Editable input | pale warm fill `FFFBEF`, thin `D8C9A3` border |
| Model-controlled / locked constant / locked seed row | grey fill `EFF1F4`, **bold**, `C2CAD3` border |
| Table header | `E2E6EB` fill, bold, rule beneath |
| Section heading | bold, no fill |
| Note | small italic grey |

Reporting Currency `SAR` and the locked Config constants are visibly different
from editable cells. Worksheet protection is **not** applied in Phase 2 — visual
and contract-level distinction is sufficient, and protection would add friction
to later development for no verification benefit yet.

---

## Build and test

    python3 -m pip install -r pccm/requirements.txt
    python3 pccm/builder/build_stage_a.py

Produces **`pccm/build/PCCM_stageA.xlsx`** and runs 113 structural verification
checks against both specifications. Exit 2 on a specification error, 1 on a
verification failure.

    python3 pccm/tests/test_phase1_structure.py             # 21 tests
    python3 pccm/tests/test_phase1_manifest_validation.py    # 10 tests
    python3 pccm/tests/test_phase2_inputs.py                 # 30 tests
    python3 pccm/tests/test_phase2_contract_validation.py    # 20 tests

All run standalone; no pytest required.

### On the artifact name

The Stage A artifact is now `PCCM_stageA.xlsx`, replacing `PCCM_skeleton.xlsx` —
it is no longer a skeleton. There is **one** Stage A artifact and one builder;
the Phase 1 suite validates the same artifact, so its structural guarantees are
still exercised on every build with no duplicate skeleton to maintain.

The Phase 1 “absence” tests were made phase-aware rather than deleted: the
workbook must still contain **no formulas**, and now must contain **exactly** the
contract's tables and defined names — no more, no less.

---

## Deliberately not implemented

- Cost Lines and Risk Register functional tables; Inflation rates; Cost and Risk
  Profiling grids
- Deterministic cost calculations, risk EMV, escalation, FX conversion, NPV
- Monte Carlo, RNG, distributions, percentiles, sensitivity
- Model Check engine, severities, the simulation gate
- Apply / Update Timeline logic
- Stale-result fingerprinting
- Dashboard formulas and charts
- VBA, buttons, macros, worksheet event handlers
- Worksheet protection
- Runtime CodeName assignment — still specification-only; **Stage B remains the
  authority**

## Known limitations

1. **List ranges are fixed-size.** `lst*` names cover the reserved data body.
   Adding values beyond the reserved rows needs a contract change and a rebuild.
   A dynamic range is deliberately deferred: it cannot be verified without Excel,
   and Stage B is the right place to decide it.
2. **Data validation is advisory, not enforcement.** Excel validation can be
   bypassed by pasting. Model Check is the authority for input correctness.
3. **Every validation permits a blank cell.** Required-ness is a Model Check
   rule, not something to block at the keyboard.
4. **Nothing consumes these inputs yet.** The defined names exist so later phases
   have a stable contract to bind to; no formula reads them today.
