# PCCM — Phase 3: Cost Lines & Risk Register Input Layer

**Status: complete, pending review.**

Phase 3 establishes the two authoritative **driver registers** — the things that
contribute uncertainty to total project cost. Still Stage A `.xlsx`; no VBA, no
Stage B, and **no calculations**.

Authority for functional design remains the **Architecture Lock, Revision B**.
Phases 1 and 2 are closed.

---

## Source authorities

| File | Owns |
|---|---|
| `spec/workbook.yaml` | structure and presentation |
| `spec/input_contract.yaml` | Setup inputs and Config master lists |
| **`spec/driver_contract.yaml`** | **Cost Line and Risk Register schemas** |

Cost Lines and Risk Register declare `body: drivers` in the manifest and carry no
blocks; the driver contract is their single layout authority. The loader rejects a
sheet that declares both.

Column **widths** live in the driver contract rather than the manifest: a column
and its width are one thing, and splitting them across two files would guarantee
drift as the schema evolves.

### Cross-contract validation

Before anything is rendered the builder asserts:

- the manifest and the input contract agree on the reporting currency
- the manifest and the driver contract agree on which sheets are driver-bodied
- every driver list validation source exists as a list defined name in the input contract
- no driver register targets a sheet the input contract already owns

---

## Preflight consistency hardening

Three guards added before Phase 3 work, closing a gap carried from Phase 2:

1. **Cross-spec reporting currency.** `workbook.yaml → model.reporting_currency`
   must equal `input_contract.yaml → model_invariants.reporting_currency`. A
   mismatch raises before workbook generation.
2. **Semantic reporting-currency input.** The invariant now names the input:
   `reporting_currency_input: "reporting_currency"` and
   `reporting_currency_defined_name: "inpReportingCurrency"`. It is validated
   against *that* input — it exists, is model-controlled, its default equals the
   declared currency, and it uses the intended defined name. It is no longer
   satisfied by "any locked text input whose default happens to equal SAR".
3. **VERSION housekeeping.** `pccm/VERSION` must equal
   `workbook.yaml → model.model_version`. Builder Version stays deliberately
   independent — they are different version concepts.

---

## Cost Line schema — `tblCostLines` on `Cost Lines`

| # | Column | Type | Ownership | Validation |
|---|---|---|---|---|
| 1 | Cost Line ID | text | **model-controlled** | none |
| 2 | Category | text | user | list ← `lstCategories` |
| 3 | Description | text | user | none |
| 4 | UOM | text | user | list ← `lstUOM` |
| 5 | Quantity | decimal | user | **none** |
| 6 | Currency | text | user | list ← `lstCurrencies` |
| 7 | Inflation Profile | text | user | list ← `lstInflationProfiles` |
| 8 | Unit Cost Min | decimal | user | none |
| 9 | Unit Cost Most Likely | decimal | user | none |
| 10 | Unit Cost Max | decimal | user | none |
| 11 | Distribution | text | user | list ← `lstDistributions` |

## Risk Register schema — `tblRiskRegister` on `Risk Register`

| # | Column | Type | Ownership | Validation |
|---|---|---|---|---|
| 1 | Risk ID | text | **model-controlled** | none |
| 2 | Risk Name | text | user | none |
| 3 | Description | text | user | none |
| 4 | Category | text | user | list ← `lstCategories` |
| 5 | Probability | percentage | user | decimal **between 0 and 1** |
| 6 | Currency | text | user | list ← `lstCurrencies` |
| 7 | Inflation Profile | text | user | list ← `lstInflationProfiles` |
| 8 | Impact Min | decimal | user | none |
| 9 | Impact Most Likely | decimal | user | none |
| 10 | Impact Max | decimal | user | none |
| 11 | Distribution | text | user | list ← `lstDistributions` |
| 12 | Risk Owner | text | user | none |

**There is no `Included` column.** Every risk entered is simulated — that decision
is locked, and `Included` is in the contract's `forbidden_headers`, so adding one
fails the build.

Both schemas are pinned by a `locked_schema` block: headers must match exactly and
in order, so a renamed, reordered or removed column fails the build rather than
the review.

---

## Validation ownership

Data validation guides entry; it is **not** the Model Check engine. All rules
permit blanks. Deliberately **not** imposed:

| Not imposed | Belongs to |
|---|---|
| `Min <= Most Likely <= Max` | Model Check |
| Quantity positivity / non-zero | Model Check |
| Required-ness and completeness | Model Check |
| Currency uniqueness | Model Check |
| Bernoulli occurrence from Probability | the simulation engine |

---

## Why there is no Total Cost column

The locked rule is: **sample the unit cost first, then multiply by the
deterministic quantity.** A user-entered total would contradict that, and a
formula placeholder would be a calculation this phase must not contain. Any
deterministic or sampled total is derived in a later phase. `Total Cost` and its
variants are in `forbidden_headers`.

---

## Permanent IDs: present in schema, lifecycle deferred

The ID columns exist because they are part of the stable schema, and every later
relationship (profiling rows, sensitivity drivers, RNG streams) keys on them.

**No identifier is allocated in Stage A.** The columns are genuinely blank. There
is deliberately:

- no pre-numbering
- no use of worksheet row numbers
- no `ROW()` or any other formula
- no fixed maximum ID range
- no claim that ID sequencing is implemented

Allocation requires lifecycle state that only Add/Delete operations can own
correctly — including the rule that IDs are never reused after deletion. That is
Stage B VBA work. Faking it now would create exactly the row-position-derived
identity the architecture forbids.

The ID columns are model-controlled: locked visual treatment, and no data
validation. The loader rejects an editable identity column outright.

---

## Column ownership is positional, not inferred

Each register has **exactly one** model-controlled column — the leading identifier
— and **every column after it is a user-owned input**. The loader asserts that
positionally.

The earlier rule rejected only `editable: false` *combined with* a declared
validation, which meant ownership was in effect inferred from whether a column
happened to carry a rule. Every column that legitimately has no validation —
Description, Quantity, Risk Name, Risk Owner, and all six three-point parameters —
could therefore have drifted to `editable: false` and silently left the user's
control while still building cleanly. Ownership is now stated, not deduced.

The independent conformance suite asserts the same thing from the other side. It
declares each column as `(key, header, type, user_owned)` and checks all four,
rather than re-declaring headers alone and then reading ownership back out of the
contract under test. A header is the weakest part of a column to pin: `Quantity`
can keep its caption while its semantic key or type drifts underneath, and the
calculation phases will bind to the key and the type.

---

## Verifying that no validation touches a protected cell

Whether a data validation covers a model-controlled cell is a question about
**Excel rectangles**, and the verifier now answers it that way, through one shared
helper:

    data_validation_intersects(worksheet, target_range)

It reduces each area of every `DataValidation.sqref` and the target to integer
bounds (`min_col, min_row, max_col, max_row`) and tests rectangle overlap. It
compares no strings, enumerates no cells — testing a 25-row column costs the same
as testing one cell — and handles a single cell, a contiguous range, and a
multi-area sqref alike. Open references such as `B:B` widen to the sheet edge
rather than parsing to a partial bound.

It is used for **both** protected regions: the locked Setup/Config identity rows
and the two driver identity columns.

**Any** overlap fails. Attaching a user-input rule to a cell the user does not own
misrepresents ownership no matter how small the overlap, so a single ID cell, a
partial run of the column, a range that merely crosses the column horizontally,
and one offending area inside an otherwise innocent multi-area validation all fail
identically.

`test_phase3_verifier_intersection.py` proves the gate itself, not just the helper:
each test builds the real workbook, injects one specific offending validation,
saves it, and runs the actual `verify_workbook` path, asserting it fails and names
the right check — plus two tests proving a rule outside every protected range
leaves the gate green.

---

## Uniform / Most Likely status

Uniform has no Most Likely parameter. Phase 3:

- **keeps** the `Unit Cost Most Likely` and `Impact Most Likely` columns
- adds **presentation-only conditional formatting** that greys the Most Likely
  cell when that row's Distribution is `Uniform`
- does **not** disable, lock or validate the cell

The rule is declared in the driver contract, not hardcoded. It constrains nothing:
the cell remains user-editable and carries no data validation, and
`test_30_conditional_formatting_is_presentation_only` asserts exactly that
distinction. Real behavioural disabling is a Stage B UI concern.

---

## Reserved rows are capacity, not a cap

Each register is materialised with **25 reserved blank rows**, declared as
`reserved_rows` in the driver contract. This is initial capacity for structural
review because Stage A has no Add/Delete macros yet.

It is **not** a business maximum. Stage B macros will grow these tables beyond it.
The Architecture design targets (200 cost lines, 100 risks) are benchmark figures
for performance testing and are deliberately encoded nowhere;
`test_29_reserved_rows_are_capacity_not_a_declared_maximum` guards against a limit
key creeping into the contract.

---

## Build and test

    python3 pccm/builder/build_stage_a.py

Produces `pccm/build/PCCM_stageA.xlsx` — the **same** Stage A artifact, evolving
phase by phase. There is no parallel workbook.

    python3 pccm/tests/test_phase1_structure.py
    python3 pccm/tests/test_phase1_manifest_validation.py
    python3 pccm/tests/test_phase2_inputs.py
    python3 pccm/tests/test_phase2_contract_validation.py
    python3 pccm/tests/test_phase3_drivers.py
    python3 pccm/tests/test_phase3_driver_contract_validation.py
    python3 pccm/tests/test_phase3_verifier_intersection.py

---

## Deliberately not implemented

- Permanent-ID allocation; Add/Delete Cost Line and Risk macros; any VBA or
  worksheet event handler
- Cost and Risk profiling; Inflation annual rates; timeline Apply / Update
- Deterministic calculations, FX conversion, escalation, NPV, expected risk
- Monte Carlo, RNG, sensitivity
- Model Check engine; stale-result fingerprint
- Results and Dashboard logic
- Worksheet protection
- Runtime CodeName assignment — still specification-only; **Stage B is the authority**
