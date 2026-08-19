# Phase 5 — Gate B — Step B1: the Windows harness extension

**Status: harness source, ready for independent review. NOTHING HAS BEEN RUN.**

Gate A is accepted and closed at `1968fb8`. This step authors the Windows Gate-B
harness required by `phase5_plan.md` §24, §25 and implementation-sequence item 11,
and reviews it statically on Linux. The sequence is deliberate: author, review on
Linux, review independently, and only then run real Excel on Windows.

---

## What this step does NOT claim

1. **NO WINDOWS RUN HAS BEEN MADE.** No Excel COM session was started, no
   `.xlsm` was driven, and `phase4_functional_test.ps1` was not executed.
2. **NO VBA HAS BEEN EXECUTED.** The transient diagnostic module has never been
   imported into a real VBA project, and not one locked vector has been evaluated
   by real VBA.
3. **Gate B has not passed, and Phase 5 is not accepted.** Everything below is a
   statement about SOURCE: what the harness would do, which authority each value
   comes from, and in what order the steps run.
4. **No production behaviour is demonstrated.** `tests/test_phase5_gate_b_harness_source.py`
   reads text. It starts nothing.

---

## The accepted Gate-A production head

```
1968fb86bc172d31fadc760f9e131a109fda718c
```

Frozen for this step and unmodified by it: `modCalcReport`, `modCalcResolve`,
`modCalcCheck`, `modCalcFactors`, `modCalcAnalytical`, `modCalcFingerprint` and
all Phase-4 production VBA, plus `calc_numeric.py`, `calc_oracle.py`,
`calc_fingerprint.py`, `calc_cases.py`, `calc_contract.yaml` and
`phase5_cases.json`'s content.

---

## Architecture: an extension, not a second harness

`phase5_gate_b_scenarios.ps1` is **dot-sourced into** `phase4_functional_test.ps1`.
It runs inside that script's one COM lifecycle, against the one Excel instance it
owns, the one workbook it opened and the one Stage-B bootstrap it ran, and
reports through the same `Add-Result`. It creates no Excel process, no release
ledger, no bootstrap invocation and no shutdown of its own — `test_02` refuses
each by name.

Every accepted Phase-4 discipline is reused rather than restated: caller-side
`@(...)`, one pipeline object per row through `Write-RowObject`, the
non-enumerating `New-Checklist` factory, a `catch` attached to its own `try`,
keyed-only fixtures, failure-safe cleanup, the owned-process identity, explicit
transient release through `Release-Transient`, and natural shutdown.

### The Phase-4 matrix is a prerequisite, and it is checked

`P5-P4` reads the results the Phase-4 matrix actually produced and requires
**35/35 PASS, 0 FAIL, 0 SKIP** before any Phase-5 scenario runs. If the matrix is
not intact it reports `P5-ALL` as **FAIL** — never as a SKIP, because
"Phase 5 was not attempted" must be as loud as "Phase 5 failed".

The 35 results are the nine lettered scenarios through `D0`, the ten sequential
steps `D-J.1` … `D-J.10`, then `K`, `K2`, `L`, `M`, `N`, `O`, `P`, `Q`, `R`,
`S`, `T`, `U`, `V`, `W`, `Y` and `Z`. No Phase-4 scenario semantics were
rewritten to make room.

---

## The missing projection, and how it was closed

**This is the one thing in this step that added a build output, and it is
reported here first because §4 asks for it to be.**

The Windows harness has to find things in the driven workbook: the five `_Calc`
ListObjects, the `calc_state` and `calc_totals` value cells and the meaning of
each row, the Setup scalars a fixture writes, and the Config/Setup tables it
seeds. **No existing build output projects any of that.**

* `stage_b_manifest.json` projects sheets, modules, entry points, API procedures,
  buttons, the timeline/counter defined names, the two driver registers and the
  three grids — and stops there.
* `phase5_cases.json` is an expected-VALUE corpus and carries no addresses.
* `phase4_scenarios.json` is the structural oracle's output, also value-only.

The layout does exist in `build/vba/modCalcContract.bas` and
`build/vba/modConstants.bas`, but teaching PowerShell to parse VBA constants
would put a second reader of the same authority inside the harness — which is
what "a second contract" means in practice.

**The projection**, `build/phase5_gate_b_inspection.json`, is emitted by the
Stage-A build from the same accepted authorities through their own loaders. It
carries **identities only** — names, sheets, addresses, rows, columns — and no
expected value, no tolerance and no analytical fact. `phase5_cases.json` remains
the sole expected-value authority.

It cannot become a second contract, because `test_42` and `test_43` pin every
address in it against the generated `modCalcContract.bas` and `modConstants.bas`.
If the two ever disagree the Linux build fails, rather than the harness silently
inspecting the wrong cell on Windows. `test_45` separately refuses any VBA
parsing in the harness, and `test_41` refuses any expected value in the
projection.

If the reviewer would rather this projection did not exist, the alternative is a
manifest key; the harness reads one JSON object either way and nothing else
changes.

---

## The 37-case coverage ledger

Every ID in `phase5_cases.json → plan_cases[*].id` maps to at least one Windows
scenario. The map is **data** (`Get-Phase5CoverageLedger`), validated by
`P5-PRE` **before Excel is started**: a case emitted into the corpus with no
mapping, a mapping naming a scenario the harness does not define, a ledger entry
for a case the corpus no longer emits, or a fixture that does not carry the
evidence its kind promises all abort the run.

Cases share scenarios and workbook fixtures, and one Excel process serves them
all. What may not happen is a case disappearing because several share a fixture.

| Case | Kind | Title | Windows scenario |
| --- | --- | --- | --- |
| 1 | `analytical` | SAR, no inflation, one project year | `P5-AN` |
| 2 | `analytical` | foreign currency | `P5-AN` |
| 3 | `analytical` | multi-year profiling with compounded inflation | `P5-AN` |
| 4 | `analytical` | present value across multiple years | `P5-AN` |
| 5 | `analytical` | Triangular deterministic basis versus mean | `P5-AN` |
| 6 | `analytical` | Beta-PERT deterministic basis versus mean | `P5-AN` |
| 7 | `analytical` | Uniform midpoint equals mean | `P5-AN` |
| 8 | `analytical` | risk expected value with probability below one | `P5-AN` |
| 9 | `analytical` | multi-year risk profile | `P5-AN` |
| 10 | `analytical` | Base Year equals Start Year | `P5-AN` |
| 11 | `analytical` | Base Year earlier than Start Year | `P5-AN` |
| 12 | `analytical` | zero inflation | `P5-AN` |
| 13 | `analytical` | negative but valid inflation | `P5-AN` |
| 14 | `refusal` | blank required inflation rate | `P5-RF` |
| 15 | `refusal` | profile does not sum to one hundred percent | `P5-RF` |
| 16 | `refusal` | Quantity of zero | `P5-RF` |
| 17 | `refusal` | negative Quantity | `P5-RF` |
| 18 | `refusal` | discount rate of minus one hundred percent | `P5-RF` |
| 19 | `analytical` | discount rate negative but above minus one hundred percent | `P5-AN` |
| 20 | `refusal` | inflation rate of minus one hundred percent | `P5-RF` |
| 21 | `analytical` | inflation rate negative but above minus one hundred percent | `P5-AN` |
| 22 | `analytical` | Uniform with a populated Most Likely, which is ignored | `P5-AN` |
| 23 | `refusal` | profile summing to one hundred percent but containing a blank | `P5-RF` |
| 24 | `refusal` | controlled refusal on Double overflow | `P5-RF` |
| 25 | `analytical` | unreferenced incomplete FX row does not block | `P5-AN` |
| 26 | `fingerprint` | fingerprint reference vector | `P5-D1`, `P5-D4`, `P5-D5` |
| 27 | `fingerprint` | delimiter-hostile field content | `P5-D6` |
| 28 | `statistics` | naive overflow with a representable result | `P5-D7` |
| 29 | `refusal` | discount factor underflow | `P5-RF` |
| 30 | `analytical` | cancellation-heavy reconciliation | `P5-AN`, `P5-ID` |
| 31 | `analytical` | Base-Year factor row | `P5-AN` |
| 32 | `runtime_only` | derived status reverts to CURRENT after an input is restored | `P5-RC`, `P5-S5` |
| 33 | `runtime_only` | mid-write failure and full logical rollback | `P5-FA` |
| 34 | `runtime_only` | invalid input with no Calculate attempted | `P5-S3`, `P5-S4`, `P5-KP` |
| 35 | `fingerprint` | locale separator injection | `P5-D2` |
| 36 | `fingerprint` | reduction beyond Long | `P5-D3` |
| 37 | `runtime_only` | failure at the commit boundary | `P5-FC` |

---

## The Windows scenarios

| ID | What it establishes |
| --- | --- |
| `P5-PRE` | Coverage preflight, pure PowerShell, before Excel |
| `P5-P4` | The Phase-4 matrix reached 35/35, 0 FAIL, 0 SKIP |
| `P5-M` | 15 modules **by name**, exactly 5 buttons, no `PCCM_Calculate` button, 6 `api_procedures` |
| `P5-EV` | No `Worksheet_Change` / `Workbook_SheetChange` in the real project |
| `P5-D0` | The transient diagnostic module imported, **after** A1 |
| `P5-D1` | Ten canonical numeric encodings on real VBA |
| `P5-D2` | Both decimal separators injected into the accepted encoder |
| `P5-D3` | All four Double-only reductions |
| `P5-D4` | UTF-16: signed `AscW`, unit counting, surrogates, length prefixes |
| `P5-D5` | The complete reference stream: unit count **and** digest |
| `P5-D6` | The eight delimiter-hostile collision probes |
| `P5-D7` | Convex statistics at the naive-overflow boundary |
| `P5-D8` | The diagnostic module **removed**, inventory back to 15 |
| `P5-AN` | Every analytical fixture, every emitted expected value |
| `P5-RF` | Every prerequisite refusal, no partial analytical output |
| `P5-ID` | Reconciliation identities I1–I5, cancellation-heavy included |
| `P5-S1`…`P5-S6` | The six-row status matrix |
| `P5-ST` | The primary staleness sequence |
| `P5-NS` | Four non-staleness proofs |
| `P5-KP` | A refusal preserves the prior successful snapshot |
| `P5-RC` | Revert to CURRENT without calculating |
| `P5-FA` / `P5-FC` | Rollback at both locked failpoint boundaries |
| `P5-AX` | The invocation axis and the attempt axis, read separately |

---

## The six-row status matrix

Every row asserts **all four** accessors — `PCCM_CalculationStatus()`,
`PCCM_CalculationAttemptResult()`, `PCCM_CalculationAttemptDetail()`,
`PCCM_CalculationFingerprint()` — plus `PCCM_CurrentInputFingerprint()` where it
applies, and the snapshot state the row requires.

| Scenario | Row | Status | Attempt | Detail | Snapshot |
| --- | --- | --- | --- | --- | --- |
| `P5-S1` | 1 successful calculation, unchanged inputs | `CURRENT` | `SUCCESS` | blank | new |
| `P5-S2` | 2 valid fingerprinted input changed, no Calculate | `STALE` | `SUCCESS` | blank | unchanged |
| `P5-S3` | 3 invalid current input, no Calculate | `INVALID` | `SUCCESS` | blank | unchanged |
| `P5-S4` | 4 invalid current input + `PCCM_Calculate` | `INVALID` | `REFUSED` | specific | unchanged |
| `P5-S5` | 5 exact restoration of the prior input, no Calculate | `CURRENT` | `REFUSED` | still readable | unchanged |
| `P5-S6` | 6 injected write failure on valid changed inputs | `STALE` | `FAILED` | specific | previous restored |

**Status is never derived from attempt history.** Rows 5 and 6 exist because the
two axes are allowed to disagree, and the harness does not tidy that away:
`test_17` requires row 5 to assert `CURRENT` and `REFUSED` together and to
compare the refusal detail against the one that was actually recorded.

Every status read goes through `PCCM_CalculationStatus` **first**. The status
cell is last-evaluated, not live; reading C19 without asking would report
whatever the previous scenario left there.

---

## The direct real-VBA vectors, and the transient module

`bootstrap/windows/phase5_gate_b_diagnostics.bas`, module name
`modPhase5GateBDiagnostics`. Fourteen `GBD_*` procedures, each a **thin wrapper**
over an already-Public accepted helper:

| Procedure | Production helper it calls |
| --- | --- |
| `GBD_Ping` | *(identity only; nothing)* |
| `GBD_CanonicalNumber` | `modCalcFingerprint.CalcFpCanonicalNumber` |
| `GBD_CanonicalNumberConstructed` | `CalcFpCanonicalNumber`, on a value built on target |
| `GBD_ConstructedValueText` | `CalcFpCanonicalNumber` |
| `GBD_ReduceDouble` | `modCalcFingerprint.CalcFpReduceDouble` |
| `GBD_TextFromUnits` | *(builds a String from code units; `ChrW$`)* |
| `GBD_Utf16Length` | `modCalcFingerprint.CalcFpUtf16Length` |
| `GBD_RawAscW` | *(raw `AscW`, so the SIGNED result is observable)* |
| `GBD_NormaliseCodeUnit` | `modCalcFingerprint.CalcFpNormaliseCodeUnit` |
| `GBD_CanonicalTextField` | `modCalcFingerprint.CalcFpCanonicalText` |
| `GBD_StreamLength` | `modCalcFingerprint.CalcFpUtf16Length` |
| `GBD_DigestStream` | `modCalcFingerprint.CalcFpDigestStream` |
| `GBD_ProbeDigest` | `CalcFpCanonicalText`, `CalcFpCanonicalInteger`, `CalcFpDigestStream` |
| `GBD_ConvexStatistic` | `modCalcAnalytical.TriangularMean` / `PertMean` / `UniformMean` |

`test_21` pins that call list exactly and proves every name is already `Public`
in the accepted production source: **no production visibility was reopened.**

### Lifecycle

```
A1  first Application.Run of the run  ->  the PRODUCTION project compiles
P5-P4  the Phase-4 matrix is intact
P5-D0  the diagnostic module is imported into the DISPOSABLE workbook
P5-D1 .. P5-D7  the locked vectors
P5-D8  the diagnostic module is REMOVED; inventory re-asserted at 15
P5-AN onward  the analytical acceptance work, with no test module installed
```

A1 remains the first real VBA compilation boundary and stays production-only
(`test_22`): a test module must never mask or contaminate proof that the accepted
project itself compiles. The module is not in the manifest, not in the structure
contract, not imported by `build_stage_b.ps1`, not under `src/vba`, creates no
button, declares no `PCCM_` endpoint, and no workbook is ever saved with it
installed (`test_19`, `test_20`, `test_23`, `test_47`).

### The two extremes are built on target

`MAX_DOUBLE` and the minimum subnormal are the two values a COM `Double` round
trip is most likely to disturb. Each is exercised **twice** — once marshalled
from PowerShell, once constructed inside VBA — and both must equal the fixture.
The subnormal is built by halving 1 exactly 1074 times, every intermediate a
power of two and therefore exact; `MAX_DOUBLE` is taken from the accepted kernel
constant rather than retyped. The vector is never skipped and never weakened, and
a marshalling fault is reported as itself.

### Separator injection

Both `"."` and `","` go into the **same** accepted encoder as its own argument,
on one host, in one run. `test_25` refuses `Application.International`,
`UseSystemSeparators`, `Set-Culture` and `CurrentCulture` in both the harness and
the diagnostic module: no regional setting is read or altered.

### The reference stream

Both the **366-unit count and the digest** are asserted, and the count first — a
digest asserted alone would agree with itself over a stream that arrived
truncated. Both values are read from the corpus; `test_07` refuses the literal
`50B6EB0E26857EA7` and the literal `366` anywhere in the harness or the
diagnostic module.

---

## Analytical and refusal coverage

For a successful calculation the harness asserts **every emitted expected value**:
`tblCalcYears`, `tblCalcInflationFactors`, `tblCalcFX`, `tblCalcDrivers`,
`tblCalcAnnual` and all ten `calc_totals` cells, plus `calc_state`. Row counts
are asserted first, so a table that came back short is a failure rather than a
reason to compare fewer rows. The driver and annual comparisons iterate the
**fixture's own field names**, so a field added to the corpus is asserted without
editing the harness.

**Blank is not numeric zero.** `Test-CalcValue` compares an expected `null` as a
blank and refuses a blank against a numeric expectation, which is what makes the
Base-Year blank rate and every N/A field meaningful (`test_36`).

A refusal must be `REFUSED` with a specific, non-empty detail and status
`INVALID`, and **no partial analytical output may survive**: every `_Calc` table
is checked for populated rows.

---

## Refusal snapshot semantics

Three groups with three different fates, compared **separately**:

| Group | Fate |
| --- | --- |
| C13:C16 — the last successful record | **UNCHANGED** |
| C23:C32 and the five analytical tables | **UNCHANGED**, blanks included |
| C17:C20 — the attempt and status axis | **CHANGED**, as the row expects |

Comparing all of C13:C20 as unchanged would assert that the refusal was never
recorded, which is the opposite of the requirement. `test_32` pins the two field
groups and proves they do not overlap; `test_34` requires row 4 to assert all
four C17:C20 changes explicitly; `test_nc_12` plants the whole-block comparison
and watches the detector see it.

---

## Staleness, revert and non-staleness

The primary sequence (`P5-S1` → `P5-S2` → `P5-ST`) changes **one profiling
weight pair** — a normal fingerprinted analytical input — not a timeline. Two
weights are exchanged so the profile still sums to 100% and the model stays
VALID: the row under test is STALE, not INVALID. Recalculating must change the
stored digest and return the model to CURRENT.

`P5-RC` restores the changed input **exactly** and does not calculate. Status
returns to `CURRENT` while the attempt axis still reads `REFUSED` and the refusal
detail is still readable, byte for byte. That disagreement is required and is not
cleaned up.

`P5-NS` proves four changes leave `CURRENT` / `SUCCESS` / unchanged digest:
Description, a **real `ListObject` sort** (not values copied between rows, so
Permanent-ID canonical ordering is proved on real Excel), Selected Confidence
Level, and an **unreferenced** FX assumption.

---

## Rollback at both locked boundaries

Through the accepted Phase-4 `PCCM_AutomationBegin(confirm, failpointName)` /
`FailPointCheck` mechanism and no other (`test_31`).

| Scenario | Failpoint | Where the production hook is |
| --- | --- | --- |
| `P5-FA` | `Phase5AnalyticalWrite` | `RunCalculation`, after `WriteAnalytical`, before `VerifyAnalytical` |
| `P5-FC` | `Phase5SuccessCommit` | `WriteSuccessCommit`, the statement immediately before `Range(CALC_STATE_VALUE_RANGE).Value2 = block` |

Both strings are a **checked copy**: `test_30` pins them against the accepted
`modCalcReport.bas` declarations, and `test_31` re-proves that the commit hook is
still adjacent to the C13:C20 assignment, so Gate B exercises that hook and not
an upstream one.

Each scenario establishes a successful snapshot, changes a valid fingerprinted
input so the model is genuinely STALE, arms the failpoint, calculates, and then
asserts: C17 `FAILED`, C18 specific, C19 a freshly **derived** status that is
never `FAILED`, C20 a fresh timestamp, `PCCM_CalculationStatus()` `STALE` and not
`CURRENT`, C13:C16 / C23:C32 / all five tables the previous snapshot exactly, no
mixed old/new analytical state, and `EnableEvents`, `ScreenUpdating` and
`Calculation` restored. It then disarms and calculates again, because a rollback
that left the workbook unusable would not be a rollback.

The immediate post-rollback / pre-metadata moment is **not** observed: the public
runtime does not expose it, and acceptance does not depend on an observation that
does not exist.

---

## The two axes

`P5-AX` reads `PCCM_AutomationResult()` and `PCCM_CalculationAttemptResult()`
**separately** and never reports one as the other. Reaching the second line at all
is part of the evidence: a `MsgBox` would have hung the call, not failed it.

A committed-`SUCCESS` / cleanup-`FAIL` disagreement is **not forced**. The
accepted harness has no safe way to make `FinishOperation` fail, and forcing it
would prove the forcing. What the source establishes is that both axes are
readable independently; the harness records a note saying exactly that.

---

## Expected-value authority

`build/phase5_cases.json`, loaded from the **supplied `-BuildDir`**, and nothing
else. `test_06` scans the Phase-5 harness for any numeric literal that could be a
hand-copied oracle value and allows exactly four: `0.0` and `1.0` as arithmetic
identities, and `0.99` and `3.75` as fixture INPUTS the harness writes into the
workbook. `test_06a` proves the diagnostic module states no expected value at
all. `test_08` proves the harness reads the corpus by reference.

---

## Module, button and API assertions

`P5-M` asserts the 15 production modules **by name in both directions** — every
manifest module present, and nothing outside the manifest present — because a
count alone would pass a project that gained a stray module and lost a real one.
Exactly five command buttons persist, each bound to its manifest entry point, and
**no shape has `OnAction = PCCM_Calculate`**.

`api_procedures` is consumed **as `api_procedures`**: the harness asserts that no
API procedure is also an entry point and that none is bound to a button. It is
not folded into `entry_points` to reuse button logic.

---

## COM lifecycle

Carried over unchanged. The final run must prove no owned Excel process leaks,
every transient COM reference releases, the instance shuts down naturally, a
forced stop is never reported as PASS, and no pre-existing Excel process is
terminated. **This step authors those checks only.**

---

## What remains unproven until the Windows run

Everything about behaviour, and specifically:

* that the harness extension **parses and runs at all** — Linux has no PowerShell
  host, so its syntax has been reviewed and structurally balanced, not executed
* that the transient module imports, compiles and removes cleanly, and that
  `Application.Run` reaches a `GBD_*` function
* that real VBA produces the ten canonical encodings, both separator results, the
  four remainders, the UTF-16 answers, the 366-unit count and
  `50B6EB0E26857EA7`
* that `ChrW$` round-trips a surrogate pair and that a NUL survives a BSTR
* that the minimum subnormal survives COM marshalling, or that the on-target
  construction is needed — the harness will report which
* that the fixture applier can drive every emitted model into the workbook, and
  that Apply Timeline generates the year columns each fixture needs
* that `PCCM_Calculate` produces the emitted expected values on real Excel
  arithmetic and a real locale
* that both injected failures roll back exactly as described
* that the six status rows come out as the matrix says
* that no owned Excel process leaks and the instance exits naturally

Gate A established what the source says. Gate B has not been run.
