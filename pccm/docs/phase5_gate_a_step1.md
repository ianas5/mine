# Phase 5 — Gate A — Step 1: contract and fingerprint foundation

**Status: ready for independent review.**

The first implementation step of Phase 5. It builds the two things every later
Phase-5 step will be measured against — the physical-layout contract and the
fingerprint reference implementation — and **nothing else**. The analytical
engine, the workbook emission and all VBA remain deliberately unwritten.

---

## 0. The five statements this step must make explicitly

> **NO VBA WAS IMPLEMENTED.**
> **NO WORKBOOK PHASE-5 BLOCK WAS EMITTED.**
> **NO WINDOWS HARNESS WAS MODIFIED.**
> **NO WINDOWS TEST WAS RUN.**
> **PHASE 6 HAS NOT BEGUN.**

Nothing in `src/vba/`, `bootstrap/windows/`, `readiness/windows/` or the Phase-4
workbook artifacts was touched. The Phase-4 source freeze holds.

---

## 1. Scope implemented

| # | Delivered | Notes |
|---|---|---|
| A | `spec/calc_contract.yaml` | the narrowly scoped fifth authority |
| B | its loader and fail-loud validator | `builder/pccm_builder/calc_loader.py` |
| C | the Python fingerprint reference implementation | `builder/pccm_builder/calc_fingerprint.py` |
| D | fixed fingerprint vectors and static contract tests | two new test modules |

Plus the four locked document errata and the new Gate-B vector-coverage
requirement, applied to `docs/phase5_plan.md` in place (§4 below).

## 1.1 Scope deliberately NOT implemented

Every one of these is Gate-A Step 2 or later, and none of it exists in this
package:

`calc_oracle` analytical formulas · `_Calc` workbook emission · Stage-A Phase-5
blocks · `modCalcContract.bas` · `modCalcFactors` · `modCalcAnalytical` ·
`modCalcFingerprint` · `modCalcResolve` · `modCalcCheck` · `modCalcReport` ·
`PCCM_Calculate` · the five status/fingerprint accessors · transactional
write-back · any Windows harness change · the transient Windows vector-coverage
diagnostic module of plan §24.1.

The calculation contract is **loaded and fully validated but not projected**: no
Stage-A code reads it, and `build_stage_a.py` is unchanged. That is intentional —
the layout is locked before anything is built against it.

---

## 2. Files changed

| File | Change |
|---|---|
| `docs/phase5_plan.md` | **modified** — four errata applied in place, plus the §24.1 Gate-B requirement and a §0 errata register |
| `spec/calc_contract.yaml` | **new** — 5th authority: `_Calc` physical layout and tolerance constants |
| `builder/pccm_builder/calc_loader.py` | **new** — loader, layout validator, authority-boundary validator, cross-contract validator |
| `builder/pccm_builder/calc_fingerprint.py` | **new** — fingerprint reference implementation; sole owner of the hash mathematics |
| `builder/pccm_builder/__init__.py` | **modified** — exports `load_calc_contract` and `CalcContractError`; docstring records that the fifth contract is validated but not yet projected |
| `tests/test_phase5_calc_contract_validation.py` | **new** — 74 tests |
| `tests/test_phase5_fingerprint.py` | **new** — 42 tests |

Unchanged: `spec/workbook.yaml`, `spec/input_contract.yaml`,
`spec/driver_contract.yaml`, `spec/structure_contract.yaml`, every file under
`src/`, every file under `bootstrap/`, every file under `readiness/`,
`builder/build_stage_a.py`, and all eleven Phase 1–4 test modules.

---

## 3. What the calculation contract owns — and what it is forbidden

**Owns:** the exact `_Calc` anchors of plan §16.3 · the Phase-5 block and
ListObject names · every column schema (order, header, type, number format, unit)
· the scalar block labels, formats and seeded initial values · the profiling and
reconciliation tolerance constants and the identity conditioning coefficients ·
the two calculation-state label sets · `FP_VERSION` · the reserved Phase-4 rows
and cells.

**Forbidden, and enforced:** the hash base, either modulus, either initial state
and the recurrence · every mathematical formula · the Cost Line / Risk Register
input schemas · the distribution master list · the FX convention · permanent-ID
rules · timeline structural rules.

Where the contract needs one of those it declares an `authority_references` entry
naming the owning file and a dotted locator, and cross-validation resolves each
locator against that file. A rename upstream therefore breaks the build here
rather than leaving a stale comment behind.

### The layout it encodes

```
_Calc rows 1–11, C10:C11   Phase-4 reserved (counters). Frozen.
calc_state    B13:C20, notes E13:E20    8 rows   C13:C20 is the success commit
calc_totals   B23:C32, notes E23:E32   10 rows   C23:C32
tblCalcYears              H:J   (3)   header row 15
tblCalcInflationFactors   M:P   (4)   header row 15
tblCalcFX                 S:U   (3)   header row 15
tblCalcDrivers            X:AR (21)   header row 15
tblCalcAnnual             AU:BB (8)   header row 15
```

Two-column gutters separate every band, so a widened schema fails the overlap
assertion instead of silently overwriting its neighbour.

### Seeded initial `calc_state` values

```
C13 blank   C14 blank   C15 blank   C16 blank
C17 NONE    C18 blank   C19 NOT CALCULATED   C20 blank
```

`Fingerprint Version` (`C15`) is **blank**, not `1`. Seeding the algorithm version
at build time would make a never-calculated workbook look as though it held a
partial successful snapshot.

---

## 4. Plan errata applied

All four are **editorial / test-definition** corrections applied in place to
`docs/phase5_plan.md` Revision E. No constant, vector, anchor, schema, expected
value or locked decision changed. A new §0 registers them; each amended passage
cites its erratum number.

| # | Correction | Sections touched |
|---|---|---|
| **E1** | Case 33 post-failure state described as two moments. After rollback and before failure metadata, `C13:C20` is restored exactly; the **final** acceptance comparison is `C13:C16` + `C23:C32` + the five analytical ListObjects, never all of `C13:C20` | §0, §23 case 33 (§12.5 and §25.7 already agreed) |
| **E2** | `281,320,423,161` is **not** `131 × Long.MaxValue`. The exact statements `131 × 2,147,483,647 = 281,320,357,757` and `281,320,423,161 = 131 × Long.MaxValue + 65,404` replace it; "approximately 131 times the signed-`Long` maximum" is the permitted prose. Reduction vectors and remainders unchanged | §0, §11.5, §23 case 36 |
| **E3** | No acceptance wording may claim VBA executes or produces the fingerprint on Linux. A new §21.0 locks the proof split. Gate-A **static** validation of VBA source is explicitly not weakened | §0, §11.6, §21.0, §21, §28.5, §28.6 |
| **E4** | "six public Phase-5 callables" → **"six public `PCCM_` automation/API entry points"**. A bound on the `PCCM_` endpoint surface, not on the `Public` keyword; numerical helpers may be `Public` where cross-module VBA calls require it | §0, header, §11.13, §27, §29 |

### The new Gate-B requirement — locked, not implemented

Plan **§24.1** and acceptance criterion **20** now require that real Windows/VBA
exercise the canonical encoder and the reducer **directly** against the complete
locked vector set — the ten numeric encodings, both decimal separators, all four
reduction vectors, the `> U+7FFF` and non-BMP UTF-16 vectors, and the
366-code-unit / `50B6EB0E26857EA7` reference — rather than inferring correctness
from one golden digest.

The preferred design is a **transient, test-only VBA diagnostic module** imported
only into the disposable Windows harness working copy. It must not enter the
Stage-B production manifest, must not persist in the accepted workbook, must not
create a button, and must not add a `PCCM_` endpoint.

**It is not implemented in this step, by instruction.**

---

## 5. The fingerprint reference implementation

`builder/pccm_builder/calc_fingerprint.py` is the **single source** for the hash
mathematics: `FP_BASE = 131`, `FP_MOD_1 = 2147483647`, `FP_MOD_2 = 2147483629`,
`FP_INIT_1 = 1`, `FP_INIT_2 = 1`. They appear nowhere else — not in YAML, not in a
second Python module. `FP_VERSION` is **projected from the contract** and passed
in; the module holds no copy, and a test asserts it does not.

Implemented: UTF-16 code-unit iteration and length · signed-`AscW` code-unit
normalisation · the `<TAG><LEN>:<VALUE>` length-prefixed field encoding · text,
canonical-numeric and stream-integer fields · record and section encoding · the
canonical `HEADER`/`COST`/`RISK` stream · ordinal UTF-16 Permanent-ID ordering ·
the two-modulus digest · the fixed uppercase `HEX8`+`HEX8` result.

Two reducers, deliberately:

- `reduce_exact` — exact integer `%`. Python has arbitrary-precision integers, so
  this **is** the definition, and it is the oracle.
- `reduce_double_only` — a faithful **Python mirror** of the locked VBA reduction
  (`x`, `q = trunc(x / m)`, `r = x − q·m`, one `>= m` correction, one `< 0`
  correction), written entirely in float arithmetic. Using `int` here would
  silently repair the imprecision the corrections exist to absorb.

The mirror is asserted equal to `reduce_exact` on boundary combinations and on a
seeded 20,000-case random sweep, and it reproduces the reference digest.

**It is a Python mirror, not a VBA execution.** No test in this package claims
otherwise.

---

## 6. Tests added

### `tests/test_phase5_fingerprint.py` — 42 tests

Every expected value is an **independent literal** transcribed from the plan.
Nothing is derived from the implementation under test.

- the reference vector: the canonical stream reproduced **exactly**,
  **366** UTF-16 code units, digest **`50B6EB0E26857EA7`** — under both reducers;
  plus a per-position mutation sweep proving every code unit, tag, length and
  colon is inside the hash;
- **all ten** locked numeric canonical strings, exactly;
- **all eight** locked collision-probe digests, exactly, all distinct, under both
  reducers;
- **all four** locked reduction vectors, each `x` asserted to exceed
  `Long.MaxValue`, plus `281320423161 = 131 × 2147483647 + 65404` and
  `131 × 2147483647 = 281320357757` asserted directly, so the corrected E2 prose
  cannot regress;
- UTF-16: a code unit above `U+7FFF` and its signed-`AscW` normalisation; a
  non-BMP character contributing **two** surrogate code units; length prefixes
  counting UTF-16 units rather than Python code points; and driver ordering that
  distinguishes UTF-16 order from Python string order;
- the corrected delimiter-collision analysis (below);
- the locale test (below);
- the hash constants as literals; both moduli prime and distinct; a simulated
  `Long`-wrapped reduction proving it cannot reach the reference digest.

### `tests/test_phase5_calc_contract_validation.py` — 74 tests

Fail-loud negative tests in the established pattern — mutate a copy, assert
`CalcContractError`. Covering: every accepted anchor · band width equals schema
width · overlapping bands · the two-column gutters · missing and unexpected
ListObjects · the Phase-4 reservation and both counter cells · `calc_state` rows
exactly 13:20 and its load-bearing row order · `calc_totals` rows exactly 23:32 ·
the seeded initial values, including the blank `Fingerprint Version` ·
`FP_VERSION = 1` · both state axes exactly, with `REFUSED` and `FAILED` on the
derived-status axis rejected **by name** · the exact 21-column `tblCalcDrivers`
and 8-column `tblCalcAnnual` schemas · `applies_to` coverage · the hash-constant
scan · the tolerance constants and cancellation-safe conditioning scales · sheet
visibility · cross-contract agreement with all four accepted authorities.

### The corrected delimiter-collision regression

The test asserts the **corrected** analysis and would fail against Revision C's
wrong claim:

| Encoding | Colliding probe rows |
|---|---|
| `U+001F` join | **4 ↔ 5** only |
| colon join | **1 ↔ 2** and **1 ↔ 4** |
| length-prefixed | **none** |

A companion test asserts that rows 1–2, 3–5 and 4–8 — the three pairs Revision C
named — do **not** collide under `U+001F`, so the old claim cannot come back.

### The locale test — Python reference only

`canonical_number(value, ".")` and `canonical_number(value, ",")` are asserted
equal, on all ten locked numeric literals.

**This proves the reference normalisation semantics. It does NOT prove VBA
`Format`/`Str` runtime behaviour under a comma locale.** That proof is reserved
for Windows Gate B (§24.1). The test's own docstring says so.

---

## 7. Exact test counts

Run from a clean tree on Linux, Python 3.11.

```
python -m pytest pccm/tests/ -q        618 passed, 0 failed
python pccm/builder/build_stage_a.py   181 passed, 0 failed   (post-build verification)
```

| Module | Before | After |
|---|---|---|
| `test_phase1_manifest_validation.py` | 10 | **10** |
| `test_phase1_structure.py` | 21 | **21** |
| `test_phase2_contract_validation.py` | 42 | **42** |
| `test_phase2_inputs.py` | 40 | **40** |
| `test_phase3_driver_contract_validation.py` | 31 | **31** |
| `test_phase3_drivers.py` | 28 | **28** |
| `test_phase3_verifier_intersection.py` | 12 | **12** |
| `test_phase4_oracle.py` | 68 | **68** |
| `test_phase4_stage_b_source.py` | 155 | **155** |
| `test_phase4_structure.py` | 43 | **43** |
| `test_phase4_structure_contract_validation.py` | 52 | **52** |
| `test_phase5_calc_contract_validation.py` | — | **74** |
| `test_phase5_fingerprint.py` | — | **42** |
| **total** | **502** | **618** |

**Every Phase 1–4 count is unchanged and no Phase 1–4 assertion was weakened,
removed or relaxed.** The 502 → 618 delta is exactly the 116 new Phase-5 tests.

Both new modules also run standalone (`python tests/test_phase5_*.py`), matching
the existing convention: 42 passed / 0 failed and 74 passed / 0 failed.

**No Windows run was performed.**

---

## 8. Negative controls — the new tests were verified to fail

A test that has never failed has not been shown to test anything. Each of these
sabotages was applied to a working copy of the implementation, the suite was run,
and the source was restored.

| Sabotage | Result |
|---|---|
| length prefix counts Python code points instead of UTF-16 units | **2 failed** |
| driver records sorted by raw Python string instead of UTF-16 order | **1 failed** |
| signed-`AscW` code-unit normalisation removed | **1 failed** |
| decimal-separator normalisation removed | **1 failed** |
| the `>= modulus` correction removed from the mirror | **1 failed** |
| the `< 0` correction removed from the mirror | **1 failed** |
| `tblCalcDrivers` band widened to `X:AS` in the real contract | **27 failed** |
| `FP_VERSION` bumped to `2` in the real contract | **31 failed** |

### A note on the two reducer corrections

On IEEE-754 doubles, `Fix(x / m)` turns out never to be off by one over this
range: an exhaustive scan near every multiple of both moduli, plus two million
random values, found **zero** cases. Neither correction branch is therefore
behaviourally reachable, and a purely behavioural suite would not notice one being
deleted — as the first run of these controls demonstrated.

The corrections are still required, because the guarantee is an **error bound**,
not an observation: `x / m ≤ 131` carries a relative error of at most `2⁻⁵³`, so
`Fix` **may** be off by at most one in either direction. Two tests were added to
close the gap:

1. a test that forces `q` off by `−1` and `+1` and asserts a single correction in
   each direction recovers the exact remainder — proving the corrections are
   *correct*, not merely present;
2. a **static** source check on `reduce_double_only` asserting the locked shape
   survives: `math.trunc`, both corrections, no `%` and no `//`, and float
   arithmetic throughout — proving they are *present*.

Both now fail when either correction is removed.

---

## 9. Proof scope — what Gate A did and did not establish

Recorded here because erratum E3 exists precisely to stop this line being blurred.

| | Gate A (this step, Linux) | Gate B (later, real Windows Excel) |
|---|---|---|
| numerical / fingerprint oracle | **Python**, asserted against fixed literals | — |
| VBA | **generated / source / static conformance only. NOT EXECUTED.** | **executed** |
| canonical numeric encoding | reference semantics, separator injected as an argument | actual, under a real locale |
| UTF-16 / `AscW` | reference semantics | actual `AscW`, including sign normalisation |
| `Double`-only reducer | Python **mirror**, equal to exact integer `%` | the **real** VBA reducer |
| end-to-end fingerprint | Python produces `50B6EB0E26857EA7` | real VBA must produce it, plus the §24.1 vector set |

Gate-A static validation of VBA source is **not weakened** by this split. No VBA
source exists yet for Phase 5; when it does, every mechanical sweep applies.

---

## 10. What review should look at

1. `spec/calc_contract.yaml` — is the authority boundary drawn where §20 says?
   Is anything owned here that belongs to another contract?
2. `builder/pccm_builder/calc_loader.py` — are the locked constants really locked
   *in the loader* rather than read from the file they are meant to check?
3. `builder/pccm_builder/calc_fingerprint.py` — the mirror's fidelity to the
   locked VBA form, and the claim that the hash mathematics exists in exactly one
   place.
4. The two test modules — are any expected values derived from the implementation
   rather than transcribed?
5. `docs/phase5_plan.md` §0 and §24.1 — are the errata editorial, as claimed?

---

## 11. Next step — NOT started

Gate-A Step 2, whatever the review scopes it to be. Nothing beyond this document's
§1 has been written, and **Gate-A Step 2 was not begun**.

**PHASE 5 GATE A STEP 1 READY FOR INDEPENDENT REVIEW**
