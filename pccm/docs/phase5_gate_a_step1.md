# Phase 5 — Gate A — Step 1: contract and fingerprint foundation

**Status: PATCHED THREE TIMES after independent review — ready for re-review.**

Three review rounds, all recorded here.

**Round 1** (§11) — three blocking defects: the validator did not lock the
accepted schema, separator normalisation was textual rather than positional, and a
required authority reference could silently disappear. Plus a packaging cleanup and
an inaccurate test comment.

**Round 2** (§12) — the three round-1 blockers confirmed fixed, and one remaining
class found: **the contract document's own identity was not locked, and part of its
machine-readable content was silently ignored.** Round 2 also found two places
where the locked design had copied the first YAML draft rather than the accepted
Revision-E plan.

**No mathematical design changed in any round.** The ten numeric literals, the
eight probe digests, the four reduction vectors, `50B6EB0E26857EA7`, the
Double-only reducer, every `_Calc` anchor, the tolerance values, the
authority-reference set and the status-axis semantics are untouched.

Round 2 does change two **presentation/audit** values, and only to bring the
contract *into* line with the accepted plan: three `tblCalcDrivers` columns carried
pseudo-units, and the ten `calc_totals` labels used an ASCII hyphen where the plan
uses an em dash. §12.5 records the full parity audit.

**Round 3** (§13) — one blocking defect *below* every guard built so far: PyYAML
resolved duplicate mapping keys silently, before validation, so a contract could
declare two competing values for one field and pass every check. Fixed with a
strict safe loader at the parser boundary. `spec/calc_contract.yaml` is unchanged
by round 3 — no design value moved.

The first implementation step of Phase 5. It builds the two things every later
Phase-5 step will be measured against — the physical-layout contract and the
fingerprint reference implementation — and **nothing else**. The analytical
engine, the workbook emission and all VBA remain deliberately unwritten.

---

## 0. The statements this step must make explicitly

> **NO VBA WAS IMPLEMENTED.**
> **NO PHASE-5 WORKBOOK BLOCK WAS EMITTED.**
> **NO WINDOWS HARNESS LOGIC WAS MODIFIED.**
> **NO WINDOWS TEST WAS RUN.**
> **STEP 2 HAS NOT BEGUN.**
> **PHASE 6 HAS NOT BEGUN.**

Nothing in `src/vba/`, `bootstrap/windows/`, `readiness/windows/` or the Phase-4
workbook artifacts was touched, in this round or either previous one. No `.ps1`
file was edited at all. The Phase-4 source freeze holds.

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
| `spec/calc_contract.yaml` | **new** — 5th authority: `_Calc` physical layout and tolerance constants. Round 2: ten redundant machine-readable fields removed (now comments); Revision-E parity of three driver units and ten totals labels |
| `builder/pccm_builder/calc_loader.py` | **new** — strict YAML loader (duplicate-key rejection), full locked-design guard, authority-boundary validator, document-version lock, table-key lock, unknown-key policy, cross-contract validator |
| `builder/pccm_builder/calc_fingerprint.py` | **new** — fingerprint reference implementation; sole owner of the hash mathematics |
| `builder/pccm_builder/__init__.py` | **modified** — exports `load_calc_contract` and `CalcContractError`; docstring records that the fifth contract is validated but not yet projected |
| `tests/test_phase5_calc_contract_validation.py` | **new** — 151 tests |
| `tests/test_phase5_fingerprint.py` | **new** — 52 tests |
| `tools/package_review.py` | **new** — blob-exact review packaging (§11.4) |

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

### `tests/test_phase5_fingerprint.py` — 52 tests

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

### `tests/test_phase5_calc_contract_validation.py` — 151 tests

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
python -m pytest pccm/tests/ -q        705 passed, 0 failed
python pccm/builder/build_stage_a.py   181 passed, 0 failed   (post-build verification)
```

Stage-A verification is **still 181/181** because Step 1 still does not emit the
calculation contract into the workbook. The patch changed validation, not
emission.

| Module | Pre-Phase-5 | Step 1 | Round-1 patch | Round-2 patch | Round-3 patch |
|---|---|---|---|---|---|
| `test_phase1_manifest_validation.py` | 10 | 10 | 10 | 10 | **10** |
| `test_phase1_structure.py` | 21 | 21 | 21 | 21 | **21** |
| `test_phase2_contract_validation.py` | 42 | 42 | 42 | 42 | **42** |
| `test_phase2_inputs.py` | 40 | 40 | 40 | 40 | **40** |
| `test_phase3_driver_contract_validation.py` | 31 | 31 | 31 | 31 | **31** |
| `test_phase3_drivers.py` | 28 | 28 | 28 | 28 | **28** |
| `test_phase3_verifier_intersection.py` | 12 | 12 | 12 | 12 | **12** |
| `test_phase4_oracle.py` | 68 | 68 | 68 | 68 | **68** |
| `test_phase4_stage_b_source.py` | 155 | 155 | 155 | 155 | **155** |
| `test_phase4_structure.py` | 43 | 43 | 43 | 43 | **43** |
| `test_phase4_structure_contract_validation.py` | 52 | 52 | 52 | 52 | **52** |
| `test_phase5_calc_contract_validation.py` | — | 74 | 110 | 135 | **151** |
| `test_phase5_fingerprint.py` | — | 42 | 52 | 52 | **52** |
| **total** | **502** | **618** | **664** | **689** | **705** |

**Every Phase 1–4 count is unchanged and no Phase 1–4 assertion was weakened,
removed or relaxed.** No Step-1 assertion was weakened either: the 74 contract
tests became 110 and the 42 fingerprint tests became 52, all by addition. Two
Step-1 tests were **re-pointed, not relaxed** — the two authority-reference cases
now assert rejection at load time (where the new set lock fires first) and are
each accompanied by a new test proving the cross-validation resolver still refuses
the same input on its own terms, so both guards are covered instead of one.

Both new modules also run standalone (`python tests/test_phase5_*.py`), matching
the existing convention: **52 passed / 0 failed** and **151 passed / 0 failed**.

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

The patch adds a second, larger set covering the three review blockers — 21
contract mutations and 8 separators, run against both the pre-patch and patched
trees. See §11.6.

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

## 11. Independent review round 1 — three blockers, fixed

### 11.1 Blocker 1 — the validator did not lock the accepted schema

**The finding.** §3 of this document claimed the contract owns and locks the exact
schemas, labels, value types, number formats, units, `applies_to` semantics,
tolerance constants and conditioning definitions. The loader locked far less: the
five anchors, the two header lists, and shape rules that only asked whether a
value was *syntactically plausible*. Everything else could be edited freely and the
build still reported success.

Review demonstrated fifteen accepted mutations. Re-run against `f6a35fe` as part
of this patch, **all fifteen were confirmed accepted**, along with six more.

**The fix.** `calc_loader.py` now holds the accepted Revision-E design in full, as
locked constants in the loader — `LOCKED_TABLES` (five `TableSchema` values
carrying anchors, `row_rule` and a `ColumnSchema` per column with `key`, `header`,
`value_type`, `number_format`, `units` and `applies_to`), `LOCKED_CALC_STATE`
(eight `StateRow` values), `LOCKED_CALC_TOTALS` (ten `TotalRow` values),
`LOCKED_TOLERANCES` and `LOCKED_CONDITIONING_TERMS`. Validation compares the
contract against them **attribute by attribute**, naming every differing attribute
in the error.

`LOCKED_TABLE_ANCHORS`, `LOCKED_DRIVER_HEADERS` and `LOCKED_ANNUAL_HEADERS` are now
*derived* from `LOCKED_TABLES`, so the anchors and the schemas cannot disagree with
each other.

**This creates no second runtime source of truth.** `calc_contract.yaml` remains
the only representation that later workbook emission will consume. The loader's
copy is a **design regression guard** that never feeds emission — the same
relationship `structure_oracle.py` already has with the structure contract.

**The tests.** Fifteen named regressions, one per demonstrated mutation
(`test_r1_…` … `test_r15_…`), plus four exhaustive per-attribute sweeps:

| Sweep | Coverage |
|---|---|
| every attribute of every table column | 39 columns × 6 attributes across all five tables |
| every table `row_rule` | 5 |
| every attribute of every `calc_state` row | 8 rows × 6 attributes |
| every attribute of every `calc_totals` row | 10 rows × 7 attributes |

Each sweep alters one attribute at a time and asserts the loader refuses it,
reporting by name every attribute that was silently accepted. That is the general
guard; the fifteen named tests are the specific ones review found.

### 11.2 Blocker 2 — separator normalisation was textual, not positional

**The finding.** `canonical_number()` replaced **every** occurrence of the
separator. That is safe only while the separator happens not to occur elsewhere in
scientific notation — and `E`, `+`, `-` and every digit do occur elsewhere. The API
accepted any one-character separator, so the contract it advertised was wider than
the one it honoured.

Confirmed against `f6a35fe`:

```
canonical_number(1.23, "E")  ->  1.2300000000000000.+00
canonical_number(1.23, "+")  ->  1.2300000000000000E.00
canonical_number(1.23, "0")  ->  1.23..............E+..
canonical_number(1.23, "1")  ->  ..2300000000000000E+00
```

The digit cases are worse than the two review quoted: a global replace of a digit
destroys the significand itself.

**The fix — positional normalisation, as preferred.** The accepted form is fixed:
optional `-`, one digit, the decimal marker, sixteen fractional digits, `E`, the
exponent sign, at least two exponent digits. The marker's **index** therefore
follows from the form, and exactly one character is rewritten. Two new named
operations make the step explicit and independently testable:

- `apply_decimal_separator(text, sep)` — models what a host formatter under that
  locale would have produced;
- `normalise_decimal_separator(text, sep)` — the inverse, and the operation the VBA
  encoder must perform on whatever its formatter returned.

Both locate the marker through `_HOST_FORM_RE`, which matches the marker as *any*
single character and is unambiguous because the fraction is fixed at exactly
sixteen digits: the exponent marker is the `E` that follows those sixteen digits,
wherever else an `E` may appear.

**The accepted domain is stated, not assumed.** A separator must be exactly one
**UTF-16 code unit** — measured in code units, not Python code points, so an astral
character is refused with a clear message rather than silently mishandled. No
locale decimal separator lives outside the BMP; the constraint is checked rather
than relied upon.

**Nothing else moved.** The ten numeric literals and `50B6EB0E26857EA7` are
unchanged, and `.` and `,` behave exactly as before.

**The tests.** Six new fingerprint tests: all twelve hostile separators against
thirteen values; the two exact reproductions asserted by literal *and* asserted
**not** to equal the malformed pre-patch output; digit separators against the
mantissa; a `-` separator against the leading sign; `apply`/`normalise` proven
inverse and proven to touch at most one character; and refusal of malformed host
strings, a mismatched marker, and an out-of-domain separator.

### 11.3 Blocker 3 — a required authority reference could silently disappear

**The finding.** Cross-validation resolved the references that were present. It
never checked that the required set was **complete**. Deleting the FX-convention
reference left a contract that loaded and cross-validated cleanly while no longer
declaring one of the boundaries Step 1 exists to protect.

**The fix.** `LOCKED_AUTHORITY_REFERENCES` locks the exact six
`(concept, owner, locator)` triples, and `_validate_authority_reference_set` runs
at **load** time: none missing, none extra, none duplicated, no concept claimed
twice, no rename. The resolver still runs at cross-validation, so a locator is
checked both for being *declared* and for *resolving*.

Only boundary metadata is locked here. The referenced values stay owned by the
upstream contracts.

**The tests.** Ten new cases: remove the FX reference; remove the distribution
reference; remove each of the six in turn; duplicate one; one concept under two
owners; change an owner; change a locator; rename a concept; add an unexpected
reference. Plus two resolver-level tests reached by replacement, so the second
guard is proven independently of the first.

### 11.4 Packaging — blob-exact review archives

**The finding.** `bootstrap/windows/com_lifecycle.ps1` and
`phase4_functional_test.ps1` differed by raw hash from the accepted Phase-4
package while being identical after LF/CRLF normalisation. **Not a Phase-4 source
change** — a packaging artefact, and recorded as such.

**Root cause, confirmed.** `pccm/.gitattributes` declares `*.ps1 text eol=crlf`
(VBA's `CodeModule.AddFromString` needs CRLF). Git stores those files with LF and
converts on checkout — and `git archive`, used to build the Step-1 ZIP, applies the
same conversion. The archive therefore carried CRLF while the tracked bytes are LF.
Compounding it, a file rewritten in place by tooling can remain LF in the working
tree while git still reports it clean, because the attribute normalises the
comparison — which is why the working tree held one CRLF `.ps1` and two LF ones.

**The fix.** `tools/package_review.py` reads each blob straight from the git object
store and writes those exact bytes into the archive: no smudge filter, no eol
conversion, no working-tree read, deterministic timestamps, file modes preserved.
Verified: all 71 tracked files in the package are byte-identical to
`git cat-file blob`, and the three PowerShell files now carry their tracked
md5s — `2e6338de…`, `62bd4964…`, `c949cd91…`.

**No PowerShell logic was modified.** No `.ps1` file was edited at all.

### 11.5 Test comment correction

The `Mod` negative control described a VBA `Mod` implementation as wrapping the
intermediate into signed 32-bit range. The verified failure mode is an
overflow / coercion failure, not a guaranteed wrap. The test keeps its purpose and
now says what it actually shows: *even a hypothetical signed-`Long` wrap produces
the wrong digest, and real VBA must never reach this path because the native `Mod`
and `\` operators are prohibited in the recurrence and a static source rule
enforces that.*

### 11.6 Negative controls for the patch

Both trees were driven through the same 21 contract mutations and the same 8
separators. `f6a35fe` is the pre-patch commit; the patched tree is this one.

| | pre-patch `f6a35fe` | patched |
|---|---|---|
| contract design changes accepted undetected | **21 / 21** | **0 / 21** |
| separators mis-encoded | **4 / 8** | **0 / 8** |

Running the new test modules against `f6a35fe` fails at import, because they
reference symbols the patch introduces. That is a collection failure, not evidence,
which is why the behavioural comparison above was run instead: it imports only
symbols present in both trees and exercises the defects directly.

---

## 12. Independent review round 2 — contract identity and ignored fields

Round 2 confirmed the three round-1 blockers fixed, reproduced 664/664, 181/181,
366 code units and `50B6EB0E26857EA7`, and independently ran a further **300,000**
Double-reducer parity cases against exact integer modulus with **zero** mismatches.

One class of defect remained, in two halves: **the contract could lie about what it
is, and part of what it says was never read.**

### 12.1 The document's own identity was not locked

`calc_contract_version: "1.0.0"` was parsed as a non-empty string and then ignored.
A loader written for 1.0.0 accepted `"9.9.9"`, `"2.0.0"` and `"foo"`.

That is the same failure as the round-1 blockers, applied to the document header:
the declaration exists to say which parser the document expects, and honouring it is
the only way a future format change becomes a build failure instead of a silent
misreading.

**Fixed.** `LOCKED_CALC_CONTRACT_VERSION = "1.0.0"`, checked **first**, before any
other parsing. Seven versions are asserted rejected — `1.0.1`, `2.0.0`, `0.9.0`,
`9.9.9`, `foo`, `1.0`, empty — and `1.0.0` accepted.

**This is a different version domain from `FP_VERSION`.** One says which document
format this is; the other says which canonical encoding produced a stored digest.
They move independently, and a test asserts they are not conflated.

### 12.2 The table mapping keys were not part of the contract

The five `table_name` values were locked; the YAML mapping keys above them were not.
`tables.calc_fx` → `foo` and `tables.calc_years` → `years2` were both accepted with
the ListObject names untouched.

That matters because a consumer addresses a block by its **semantic key**, not by
the Excel object name — so a rename breaks every such consumer while the workbook
stays byte-identical.

**Fixed.** `LOCKED_TABLE_KEYS` locks the five keys **and the key → ListObject
pairing**:

```
calc_years              -> tblCalcYears
calc_inflation_factors  -> tblCalcInflationFactors
calc_fx                 -> tblCalcFX
calc_drivers            -> tblCalcDrivers
calc_annual             -> tblCalcAnnual
```

Tests: each of the five renamed, the two specific renames review demonstrated, a
missing key, and two keys swapped onto each other's ListObjects — the case that
locking either half alone would miss.

### 12.3 Machine-readable fields that were silently ignored

Ten fields were present in the YAML, absent from `CalcContract`, and never
validated. Review changed all ten and the loader accepted every one. Same failure
class as round 1: the YAML reads as authoritative while the loader discards part of
it.

**The rule now enforced.** A machine-readable field is **parsed and validated**, or
it **does not exist**. There is no third category.

**Decision: removed, not retained.** All ten were redundant — each value is either
derivable from data already present, or a boundary already declared once in
`authority_references`:

| Removed field | Why it was redundant |
|---|---|
| `calc_state.commit_range` | `value_column` + `first_row`..`last_row` → `C13:C20` |
| `calc_state.snapshot_range` / `attempt_range` / `derived_range` | follow from the locked row groups |
| `calc_totals.value_range` | `value_column` + `first_row`..`last_row` → `C23:C32` |
| `calc_totals.units` | every one of the ten fields already declares `SAR` |
| `phase4_reservation.owning_contract` | boundary already in `authority_references` |
| `fingerprint.mathematics_owner` | boundary already in `authority_references` |
| `fingerprint.version_written_by` | enforced by the locked `initial: null` on row 15 |
| `calc_fx.fx_convention_owner` | boundary already in `authority_references` |

Each explanation survives **as a YAML comment**, which cannot be mistaken for data.
The derived ranges remain available from the parsed contract
(`calc_state.value_range()` → `C13:C20`), and a test asserts so — they were removed
for being derivable, not for being unwanted.

Their mutation tests are replaced, as instructed, by tests proving **reintroduction
is refused** as an unsupported key.

### 12.4 Unknown keys are refused everywhere

`sheet: {foo: bar}` was silently accepted. Fail-loud allowed-key validation now
covers every mapping level: root, `sheet`, `phase4_reservation`, `fingerprint`,
`state_labels`, each scalar block, each scalar field (with per-block key sets, since
a `calc_state` row and a `calc_totals` row legitimately carry different fields),
each table, each table column, `tolerances`, and each authority-reference entry.

Documentary `note` keys remain permitted — **explicitly, per level**. Their presence
is a decision, not evidence that arbitrary keys are tolerated, and a test asserts a
note can still be edited freely.

A dedicated test covers the practical case: a misspelled required key
(`value_colum`) leaves the real key missing *and* adds an unknown one, and must not
pass either way.

### 12.5 Revision-E parity audit

The locked design had been built by transcribing the first YAML implementation.
That is exactly the drift a design regression guard is supposed to prevent, so the
plan was re-read column by column. **Two discrepancies, both confirmed; nothing
else in the five schemas, eight state rows or ten totals rows differed.**

**`tblCalcDrivers` pseudo-units.** Plan §16.4 records the unit of `Distribution`,
`Currency` and `Inflation Profile` as `—`. The contract said `"name"`, `"key"` and
`"key"`. These are categorical identifiers; a unit invents a meaning the plan does
not have. All three now carry the plan's own em dash, exposed as `NO_UNIT` in the
loader.

Note the distinction the plan draws and the fix preserves: `Currency` in
**`tblCalcFX`** *is* that table's key and keeps `units: key`; `Currency` in
**`tblCalcDrivers`** is a categorical attribute and carries `—`. Same for
`Inflation Profile` across `tblCalcInflationFactors` and `tblCalcDrivers`.

**`calc_totals` labels.** The plan writes `Escalated Deterministic Base — Nominal`
with an em dash (U+2014); the contract used an ASCII hyphen-minus. The contract owns
labels, so all ten were aligned to the plan's wording. `Mean-Basis` keeps its
ordinary hyphen — only the separator before *Nominal* / *PV* is an em dash.

Presentation and audit semantics only. No calculation changes.

**The parity test.** A **third independent copy** of the design now lives in the
test module — `PLAN_TABLE_KEYS`, `PLAN_YEARS`, `PLAN_INFLATION`, `PLAN_FX`,
`PLAN_DRIVERS`, `PLAN_ANNUAL`, `PLAN_DRIVER_APPLIES`, `PLAN_CALC_STATE`,
`PLAN_CALC_TOTALS`, `PLAN_TOLERANCES`, `PLAN_ANCHORS` — transcribed by hand from
the plan and reading **nothing** from `LOCKED_TABLES`. Both the contract *and* the
loader's lock are asserted against it, so the two cannot drift away from the plan
together and still confirm each other.

### 12.6 Negative controls against `d0c9bca`

Both trees driven through the same 29 mutations. `d0c9bca` is the round-1 commit.

| | `d0c9bca` | patched |
|---|---|---|
| unsupported contract versions accepted | 4 / 4 | **0 / 4** |
| table mapping keys renamable | 5 / 5 | **0 / 5** |
| ignored fields changeable / reintroducible | 10 / 10 | **0 / 10** |
| unknown keys accepted | 10 / 10 | **0 / 10** |
| **total undetected** | **29 / 29** | **0 / 29** |
| driver columns carrying a pseudo-unit | 3 | **0** |
| `calc_totals` labels not using the plan em dash | 10 / 10 | **0 / 10** |

As in round 1, the new test modules cannot simply be run against the old commit —
they import symbols the patch introduces — so the comparison above imports only
symbols present in both trees and exercises each defect directly.

### 12.7 Packaging — unchanged

The raw-git-blob packaging of §11.4 is kept. Review confirmed the PowerShell
difference is line endings only, with identical content after normalisation, and
recorded no semantic Phase-4 defect. No `.ps1` file was edited in either round.

### 12.8 What round 2 did not touch

Unchanged and re-asserted by the existing tests: the fingerprint constants,
`FP_VERSION`, the canonical grammar, the UTF-16 rules, the ten numeric vectors, the
eight collision vectors, `50B6EB0E26857EA7`, the Double-only reducer, the four
reduction vectors, positional decimal normalisation, the six-triple
authority-reference set, the tolerance values and conditioning mathematics, the
`_Calc` anchors, and the status-axis semantics.

No round-1 test was weakened or removed. The 135 contract tests contain all 110
from round 1.

---

## 13. Independent review round 3 — the parser boundary

Round 3 confirmed every earlier finding fixed and reproduced 689/689, 181/181,
135/135 contract, 52/52 fingerprint. **One blocking defect remained, below every
guard built so far.**

### 13.1 The finding — duplicate YAML keys were resolved before validation

`yaml.safe_load` accepts a mapping key declared more than once and silently keeps
the last. It does so **at parse time**, before any validator runs, so every lock
added in rounds 1 and 2 is defeated by writing the field twice:

```yaml
units: "USD"
units: "SAR per unit"
```

A human reading the contract sees `USD`. The validator only ever receives
`SAR per unit`, compares it against the locked design, and reports success.

That is worse than an unvalidated field. An unvalidated field is merely
unchecked; this one **actively misleads its reader** while passing every check.
And it is general: the same trick hides a loosened tolerance, a re-pointed
semantic table key, a changed `header_row`, or a second `calc_contract_version`.

Reproduced against `06aa1ae`: **20 of 20** duplicate-key mutations accepted, at
every nesting level from the document root to an individual table column.

### 13.2 The fix — a strict safe loader at the parser boundary

`_StrictYamlLoader` derives from `yaml.SafeLoader` and overrides
`construct_mapping` to refuse a key that has already appeared in **that same
mapping**. `load_calc_contract` parses through `_strict_safe_load` **first**,
before anything else, because every guard downstream assumes each field was
declared exactly once — and that assumption belongs to the parser.

**Recursive by construction.** PyYAML calls `construct_mapping` for every mapping
node at every depth, so one override covers the root, `sheet`,
`phase4_reservation`, `fingerprint`, `state_labels`, `scalar_blocks`, an
individual scalar block, an individual scalar field, `tables`, an individual
table, an individual table column, `tolerances`, `conditioning_terms` and an
`authority_references` item — without fourteen one-off checks that could drift
apart.

Detection runs on the **raw node**, before `flatten_mapping`, so it sees exactly
what the file says rather than a post-merge view.

**Safe loading is preserved**: no arbitrary Python object construction, no unsafe
tags, and scalar, list, null, boolean and nested-mapping semantics unchanged. A
test constructs an `!!python/object/apply` tag and asserts it is refused; another
asserts the ordinary semantics still parse identically.

**Neither first-wins nor last-wins.** Both orderings are asserted rejected, so no
implicit resolution rule exists to rely on. The error names the key, the source
file, and the line and column of **both** declarations.

### 13.3 It is an ADDITIONAL guard, not a replacement

Every round-1 and round-2 validator still runs and still fires. A dedicated test
re-asserts twelve of them after the strict loader is in place: the version lock,
the semantic table-key lock, unknown-key rejection, the table schema lock, the
`calc_state` and `calc_totals` locks, exact tolerances, conditioning terms, the
authority-reference set, hash-mathematics exclusion, the Phase-4 reservation and
the status-axis lock.

### 13.4 Negative controls against `06aa1ae`

| | `06aa1ae` | patched |
|---|---|---|
| duplicate-key mutations accepted | **20 / 20** | **0 / 20** |
| unmodified contract still loads | yes | yes |

The twenty span all fourteen required levels. Ten of them use **contradictory**
values — `USD` against `SAR per unit`, `1e-3` against `1e-9`, `header_row: 16`
against `15`, `i1: [abs_d, abs_e]` against the correct terms — because two
identical values would only prove detection exists, whereas two different values
are the actual hazard.

**A correction to my own first attempt.** Two of the twenty cases originally
duplicated a `- key:` line. That starts a **new list item**, not a duplicate
mapping key, and was caught by the existing schema-count check rather than by the
new guard — so it proved nothing about duplicate detection and would have
overstated coverage. Both were re-anchored to a key *inside* the field and column
mappings (`row:` and `header:`), which is what the requirement actually asks for.
The same correction was applied to the corresponding entries in the test suite.

### 13.5 Documentation counts corrected

§6's per-module headings still carried the counts from the first Step-1 revision
(42 and 74), and §7's standalone line carried the round-1 figures (52 and 110).
Both were true when written and misleading as current state. Every current-state
count is now the actual post-patch value; the progression table in §7 keeps the
historical columns, explicitly labelled by round.

### 13.6 What round 3 did not touch

`spec/calc_contract.yaml` is **unchanged** — no design value moved. Nor did the
fingerprint constants, `FP_VERSION`, the canonical encoding, the UTF-16 semantics,
the numeric literals, the collision vectors, `50B6EB0E26857EA7`, the Double-only
reducer, positional decimal normalisation, the authority-reference set, the
Revision-E schemas, the tolerances, the conditioning scales, the `_Calc` anchors,
the status semantics, any Phase-4 source, or any PowerShell.

---

## 14. Narrow reopening — conditioning erratum C1

**Step 1 was accepted and closed. It was reopened once, narrowly, and only
because a later regression proved it necessary.**

Gate-A Step 2 demonstrated that the accepted plan §15 conditioning definitions do
not satisfy their own stated objective. `spec/calc_contract.yaml` owns the
`conditioning_terms` representation, so the corrected operand names had to land
here. **Nothing else in Step 1 was touched.**

| Identity | Was | Now |
|---|---|---|
| `i1` | `abs_a`, `abs_b`, `abs_c` | `sum_abs_a_driver_contributions`, `sum_abs_b_driver_contributions`, `sum_abs_c_driver_contributions` |
| `i2` | `abs_c`, `abs_d`, `abs_e` | `sum_abs_c_driver_contributions`, `sum_abs_d_driver_contributions`, `sum_abs_e_driver_contributions` |
| `i3a` / `i4a` | `sum_abs_annual_base`, `abs_c` | `sum_abs_annual_base_driver_contributions`, `sum_abs_c_driver_contributions` |
| `i3b` / `i4b` | `sum_abs_annual_risk`, `abs_d` | `sum_abs_annual_risk_driver_contributions`, `sum_abs_d_driver_contributions` |
| `i3c` / `i4c` | `sum_abs_annual_total`, `abs_e` | `sum_abs_annual_total_driver_contributions`, `sum_abs_e_driver_contributions` |

The old names meant already-cancelled numbers — the headline totals, and the
annual row aggregates. The new ones mean the underlying per-driver and
per-driver-per-year contributions, before aggregation. Plan §15 erratum C1 carries
the reasoning and both reproducers.

**No tolerance number changed.** `1e-9`, `1e-6`, `1e-12` and `1` are exactly as
accepted; `LOCKED_TOLERANCES` is untouched.

Updated for the erratum, and only for it: `spec/calc_contract.yaml`
(`conditioning_terms` and its explanatory comment), `calc_loader.py`
(`LOCKED_CONDITIONING_TERMS`), and the contract regression and Revision-E parity
tests that assert those names. Every other Step-1 lock — the document version, the
semantic table keys, the unknown-key policy, the full schema locks, the
authority-reference set, the hash-mathematics exclusion, the strict YAML loader —
is unchanged, and the contract test count is unchanged at **151**.

---

## 15. Next step — NOT started

Gate-A Step 2, whatever the review scopes it to be. Nothing beyond this document's
§1 has been written, and **Gate-A Step 2 was not begun**.

All three patch rounds stayed entirely inside Gate-A Step 1. Restating it:

> **NO VBA WAS IMPLEMENTED.**
> **NO PHASE-5 WORKBOOK BLOCK WAS EMITTED.**
> **NO WINDOWS HARNESS LOGIC WAS MODIFIED** — no `.ps1` file was edited at all;
> the packaging change reads git blobs and touches no PowerShell.
> **NO WINDOWS TEST WAS RUN.**
> **STEP 2 HAS NOT BEGUN.**
> **PHASE 6 HAS NOT BEGUN.**

Stage-A verification remains **181 / 181**: Step 1 still does not emit the
calculation contract into the workbook, and `build_stage_a.py` is unchanged.

**PHASE 5 GATE A STEP 1 STRICT-YAML PATCH READY FOR INDEPENDENT REVIEW**
