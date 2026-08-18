# Phase 5 — Gate A — Step 5: the resolution layer

**Status: CORRECTED after independent review — ready for re-review.**

Step 4 is accepted and closed at `4b4e221`. This step adds one hand-written VBA
module — `modCalcResolve` — that reads the workbook and hands back plain typed
data, plus the static Linux suite that reads it, plus the narrow build plumbing
it needs.

---

## Correction round — five blocking defects found by independent review

Independent review of `17b1229` confirmed the architecture and most of the source
shape: reference-set-first ordering, referenced-only FX, the global reporting-
currency invariant without unconditional seeding, referenced-only inflation,
calendar-year anchoring, Permanent-ID profiling, blank ≠ zero, binary key
matching, applied-timeline use, empty-driver support, no `_Calc` write-back, no
`PCCM_Calculate`, and three unchanged numerical modules. It then found five
blocking defects.

### 1. The Phase-4 structural prerequisites were never invoked

The first submission said the `STRUCTURE CHANGE PENDING` gate belonged to the
future checker. **That ownership was wrong** — a boundary error, not an accepted
ambiguity. The locked plan assigns structural prerequisites to Phase 4, *invoked,
not duplicated*, and `modCalcCheck` is the Phase-5 **numerical** prerequisite
checker; it must not absorb the structural gate.

`ResolveModel` began directly with `ResolveAppliedTimeline`, so a workbook whose
applied cells still held numbers could resolve even while its entered structure
had drifted away from them.

A private `StructuralPrerequisites` adapter now runs **first**, before any
calculation input is read:

| Structural state | Outcome |
| --- | --- |
| `STATE_NOT_APPLIED` | controlled failure |
| `STATE_PENDING` | controlled failure |
| unreadable / unrecognised | controlled failure |
| `STATE_CURRENT` | continue only if `modStructuralCheck.ValidateStructure()` returns `""` |

A non-empty report is a controlled failure and **the Phase-4 report is preserved
verbatim** in the detail — Phase 4 already says which invariant failed and where,
and rewording it would lose the only description the user can act on.

No structural rule is copied: no ID pattern check, no duplicate-ID check, no
orphan-row check, no profiling or inflation grid matching, no counter integrity.
`test_45` refuses every one of those identifiers in the resolver and requires
exactly **one** call into `modStructuralCheck`. `modStructuralCheck` itself is
untouched.

Calling the gate first is not permission to widen anything: Phase-5 assumption
resolution below is still referenced-only.

### 2. D2 was missing from referenced inflation resolution

A rate of `-1` or lower resolved as a valid assumption. The accepted oracle
refuses it inside `resolve_inflation`: `1 + rate <= 0` collapses the price base
and no inflation factor can be built from it.

The check now sits on each referenced profile's required year, after the rate is
read and before it is stored, as `rate <= -1#` — the same condition for a finite
Double without forming the sum. It is deliberately **not** in `NumericCell`,
which also reads values where a negative number is legitimate, and it is inside
the referenced-profile loop, so a bad rate on an **unreferenced** profile still
cannot block the model. The refusal names the profile, the calendar year and the
condition.

### 3. Identifiers were type-coerced with `CStr`

`RawCellText` ended with `text = CStr(cell.Value)`, which silently turned the
number `123` into the key `"123"` and `True` into `"True"`. A driver whose
Currency cell holds a number must not match an FX row whose Currency is the
corresponding text merely because VBA can render both as Strings.

A type gate now proves `VarType(cell.Value) = vbString` **before** the value is
taken. Conversion is the mechanism the gate exists to prevent, so it cannot be
the mechanism that produces the key. Proven text is then taken exactly as it
stands — still no trim, no case fold, no default — and a whitespace-only String
is still refused in `ExactIdentifier`.

The old test asserted the presence of `text = CStr(cell.Value)`, which locked the
defect in. It now asserts the gate precedes the assignment and precedes any
successful return.

### 4. Numeric-looking text was accepted as a number

`NumericCell` and `NumericNamedCell` delegated straight to
`modWorkbook.TryReadDouble`, whose Phase-4 semantics deliberately parse a
non-empty String through `IsNumeric` and `CDbl`. So `"0.05"` typed into a rate
cell became a Double.

The accepted oracle's `_numeric` is stricter: a real number is accepted; blank,
Boolean and text are refused. **A numeric-looking String is still text.**

A private `IsRealNumber` now gates both readers on `VarType`, accepting only
`vbInteger`, `vbLong`, `vbSingle`, `vbDouble`, `vbCurrency`, `vbDecimal` and
`vbByte`. **`TryReadDouble` is not weakened** — it has legitimate Phase-4
structural callers, and `test_51` asserts it is untouched and that the Step-5
rule was not pushed into Phase 4. The rule is also deliberately **not** applied to
`YearColumn`: an Excel table header is a text label by nature, and parsing a
numeric year header is a different structural operation that keeps Phase 4's
semantics (`test_52`).

### 5. A referenced profile could escape existence checking

`ResolveInflationRates` returned early on `nameCount = 0 Or yearCount = 0`. Those
are not the same question:

* `nameCount = 0` — no profile is referenced, nothing is consulted, an empty
  result is right.
* `yearCount = 0` with `nameCount > 0` — profiles **are** referenced and there
  are simply no annual rates to read. Each referenced profile must still exist.

A one-year `Base Year = Last Year` model could therefore let a driver name a
profile that was not in the table.

The guards are now separate. `nameCount = 0` returns before the table is opened;
otherwise every referenced key is resolved and the required-year loop then
iterates zero times. No rate is fabricated, no Base-Year rate is invented, and
the rate array is allocated only where there are rates to hold.

---

## What this step does NOT claim

1. **NO VBA WAS EXECUTED.** There is no VBA interpreter on Linux, and none was
   simulated. Nothing here observed the resolver read a cell.
2. **No workbook was read.** Every assertion in
   `tests/test_phase5_resolve_source.py` is a statement about source text: which
   procedures exist, in what order they call one another, and which constructs
   appear in executable code.
3. **No resolved value has been checked for correctness.** Whether Excel returns
   what these procedures expect, whether a real `tblFXRates` resolves, whether a
   real profiling grid matches by Permanent ID at runtime — all Gate B's.
4. **The resolver has never been compiled.** `Option Explicit` is present and
   the source tests check block balance, line length and declaration shape, but a
   static reader is not a compiler.
5. **No Windows artefact changed.** The bootstrap, the COM lifecycle script and
   the Phase-4 functional harness are untouched.
6. **Nothing was written anywhere.** There is no `_Calc` write-back, no
   `PCCM_Calculate`, no status accessor, no button, no attempt state machine, no
   RNG.

---

## Responsibility

`modCalcResolve` is the worksheet-aware resolution layer, and the only one.

| Module | Worksheet access |
| --- | --- |
| `modCalcResolve` | **yes** |
| `modCalcFactors` | no |
| `modCalcAnalytical` | no |
| `modCalcFingerprint` | no |

Allowing the resolver workbook access does **not** relax the kernel's sweep.
`test_04` still refuses `Application.`, `ThisWorkbook`, `ActiveWorkbook`,
`Worksheets`, `Worksheet`, `Range`, `Cells`, `ListObject(s)`, `Names(`,
`Evaluate`, `WorksheetFunction` and `modWorkbook.` in all three numerical
modules, and `test_05` asserts the dependency runs one way — the resolver calls
`modWorkbook`, and no numerical module calls the resolver.

The resolver goes through the accepted Phase-4 primitives (`Sh`, `Lo`,
`LoExists`, `BodyRowCount`, `CellIn`, `NameExists`, `ReadValue`, `IsEmptyCell`,
`TryReadDouble`, `IsWholeInRange`, `SafeLong`) rather than reaching for
`ThisWorkbook` itself.

### It reimplements nothing

No distribution mean, no deterministic central value, no inflation compounding,
no discount arithmetic, no `Knom`/`Kpv`, no A/B/C/D/E, no fingerprint, no
reconciliation. Where a resolved factor must be materialised it **calls** the
accepted function: `ResolveInflationFactors` delegates to
`modCalcFactors.BuildInflationFactors` and contains no compounding of its own,
and every numeric read is range-checked through
`modCalcFactors.IsUsableDouble`. `test_34` refuses the owned names in the
resolver's executable code and refuses a compounding or convex-statistic
expression; `test_35` requires the delegation.

---

## Public API

```vba
Public Function ResolveModel(model As ResolvedModel, detail As String) As Boolean

Public Function ResolveAppliedTimeline(timeline As ResolvedTimeline, _
                                       detail As String) As Boolean
Public Function ResolveProjectYears(timeline As ResolvedTimeline, _
                                    projectIndexes() As Long, calendarYears() As Long, _
                                    detail As String) As Boolean
Public Function ResolveDrivers(drivers() As ResolvedDriver, driverCount As Long, _
                               detail As String) As Boolean
Public Function ReferencedCurrencies(drivers() As ResolvedDriver, driverCount As Long, _
                                     names() As String, nameCount As Long, _
                                     detail As String) As Boolean
Public Function ReferencedProfiles(drivers() As ResolvedDriver, driverCount As Long, _
                                   names() As String, nameCount As Long, _
                                   detail As String) As Boolean
Public Function ResolveFxRates(names() As String, nameCount As Long, _
                               rates() As Double, detail As String) As Boolean
Public Function ResolveInflationRates(names() As String, nameCount As Long, _
                                      timeline As ResolvedTimeline, rates() As Double, _
                                      yearCount As Long, detail As String) As Boolean
Public Function ResolveInflationFactors(rates() As Double, profileIndex As Long, _
                                        yearCount As Long, timeline As ResolvedTimeline, _
                                        factors() As Double, detail As String) As Boolean
Public Function ResolveProfileWeights(drivers() As ResolvedDriver, driverCount As Long, _
                                      timeline As ResolvedTimeline, weights() As Double, _
                                      detail As String) As Boolean
```

Ten Public procedures, whitelisted exactly in both directions by `test_39`.
`ResolveModel` is the entry point; the nine narrower ones exist because each
stage is independently testable and because later orchestration will want them
separately. Every one carries a `detail` out-parameter (`test_38`) — a
resolution failure a user cannot act on is barely better than a crash.

`ListObject`, `Range` and `Worksheet` appear only in **private** helper
signatures (`test_40`).

## The plain structures it produces

```vba
Public Type ResolvedTimeline          Public Type ResolvedDriver
    BaseYear As Long                      PermanentId As String
    StartYear As Long                     IsRisk As Boolean
    Duration As Long                      Currency As String
    LastYear As Long                      InflationProfile As String
    DiscountRate As Double                Distribution As String
End Type                                  DistKind As Long
                                          Quantity As Double
Public Type ResolvedModel                 Probability As Double
    Timeline As ResolvedTimeline          MinValue As Double
    Drivers() As ResolvedDriver           MostLikely As Double
    DriverCount As Long                   MaxValue As Double
    Currencies() As String                HasMostLikely As Boolean
    CurrencyRates() As Double          End Type
    CurrencyCount As Long
    Profiles() As String
    ProfileCount As Long
    InflationRates() As Double         ' (profileIndex, yearOffset)
    RequiredYearCount As Long
    Weights() As Double                ' (driverIndex, projectYearOffset)
    ProjectIndexes() As Long
    CalendarYears() As Long
    DriverFxRates() As Double          ' parallel to Drivers
End Type
```

Every field is `Long`, `Double`, `Boolean`, `String` or a typed array of those.
No `Range`, `Worksheet`, `ListObject`, `Object` or `Variant` appears in any
carrier (`test_06`), and no carrier stores a workbook row as identity
(`test_07`) — identity is the Permanent ID.

Every count is explicit, for the same reason the Step-4 aggregate boundaries
carry theirs: **a VBA array cannot represent a zero-element set**, so the empty
model would be unreachable if the count were derived from `LBound`/`UBound`.

`ResolvedDriver` is a resolution-only carrier, deliberately distinct from
`DriverFactors`. These are values that could be READ as the type the model
needs; they are not yet values that have been CHECKED. `modCalcCheck` is next
and owns that.

---

## The ordering rule

```
ResolveAppliedTimeline → ResolveProjectYears
  → ResolveDrivers                              identify the model
    → ReferencedCurrencies, ReferencedProfiles  derive the reference sets
      → ResolveFxRates, ResolveInflationRates   consult the assumptions
        → AttachDriverFx, ResolveProfileWeights
```

That order **is** the referenced-only rule. A Config assumption for a currency
or a profile nobody uses cannot block a valid model, because resolution never
asks about it.

`test_10` proves the ordering by locating the first call to each stage inside
`ResolveModel` and requiring the indices to increase; moving `ResolveFxRates`
above `ReferencedCurrencies` fails it. `test_11` proves the reference sets are
derived from the drivers alone — neither builder may mention `TBL_FX_RATES`,
`TBL_INFLATION` or `modWorkbook.` at all.

---

## FX: a global invariant, and a referenced set

Two questions, deliberately not conflated.

**Is the reporting currency sound?** A GLOBAL INVARIANT. It must appear exactly
once in the FX table and must equal its identity rate, in every model —
including one that references no currency and one with no drivers.
`ReportingCurrencyInvariant` runs **before** the empty-reference-set return
(`test_14`, `test_33`), so an empty driver set does not excuse a broken
reporting currency.

**Which currencies does this model resolve?** Only those a driver references.
The resolved set is sized `0 To nameCount - 1` from the REFERENCED names and is
never widened (`test_15`):

| Model | Resolved FX set |
| --- | --- |
| empty driver set | empty |
| USD only | `{USD}` — no injected SAR row |
| SAR only | `{SAR}` |
| USD + SAR | both |

Being validated globally does not make a currency referenced.

**Referenced foreign currencies** must match exactly one row, with a numeric,
strictly positive rate. `MatchingFxRows` **counts** every match rather than
returning the first, so a duplicate is reported instead of silently resolved
(`test_17`). **Unreferenced** currencies are never looked at: the loop walks the
referenced names, not the FX table, and `test_12` fails any `For … = 1 To
BodyRowCount` sweep in `ResolveFxRates`.

`"SAR"` is not written into the VBA. `REPORTING_CURRENCY` and
`REPORTING_CURRENCY_RATE` are projected from the FX table's own locked seed row
(`test_16`).

---

## Inflation: referenced profiles, anchored to calendar years

The required span is `BaseYear + 1 .. LastProjectYear`, selected **by calendar
year** (`test_19`). A Start Year shift therefore selects the rates for the new
years rather than moving the old values positionally. The column is located by
its header value, never by arithmetic on a column index (`test_20`), so a grid
that has not been regenerated for the applied timeline reports a missing column
instead of silently reading the wrong year.

A missing profile, a missing required year, a blank required rate or a
non-numeric one is a controlled failure. **A blank is not zero** — the grid is
seeded blank precisely so an assumption the user never made cannot be fabricated
as 0%. No Base-Year rate is invented to fill the array (`test_21`): the span
starts at `BaseYear + 1`, and where `BaseYear = LastYear` the required span is
legitimately empty.

An incomplete or invalid **unreferenced** profile is never consulted
(`test_13`).

---

## Applied timeline

The resolver reads `nmBaseYear_Applied`, `nmStartYear_Applied`,
`nmDuration_Applied` and `nmLastYear_Applied`, and **never** the entered aliases
(`test_18`). An entered Duration that has not been applied has not generated its
project-year columns, so calculating from it would calculate against a shape the
workbook does not have.

Project year 1 is the Start Year (`test_22`). The discount rate is read from
`inpDiscountRate` as an ordinary required Setup input; a blank one is an unmade
assumption, not zero.

The resolver **does** gate on `nmStructuralState`, through the
`StructuralPrerequisites` adapter described in the correction round above. The
first submission claimed that gate belonged to the checker; that was a boundary
error found in review, not an accepted ambiguity.

---

## Profiling: by Permanent ID

The grid row is found by matching the Permanent ID (`test_23`) and never by
walking the register and the grid in parallel (`test_24`). A driver reorder
cannot attach another driver's weights. Only the project-year columns belonging
to the applied timeline are read (`test_25`).

**A numeric `0` weight is legitimate** — a driver may genuinely spend nothing in
a year. **A blank weight is not zero**: `NumericCell` catches the blank *before*
any numeric coercion (`test_26`) and refuses it. Nothing repairs a weight.

The profiling-sum tolerance check is **not** duplicated here. It belongs to the
accepted numerical prerequisite design, and `TOL_PROFILING_SUM_ABSOLUTE` has one
authority.

---

## Identifiers are exact

Permanent ID, Currency, Inflation Profile and Distribution are used exactly as
entered. `" USD "` is not `"USD"`.

`RawCellText` reads the cell text without trimming (`test_27`) — deliberately
not `modWorkbook.TextOf`, which trims. Trimming is right for deciding whether a
row is *populated* and wrong for the key a lookup will *compare*. So row
presence reuses Phase 4's own definition (a non-blank key column) while the key
itself is read raw: presence and identity are different questions and only one
of them may trim.

`ExactIdentifier` stores `text` unchanged (`test_28`). A whitespace-only value
is **refused**, not trimmed into an empty key that would then fail a lookup for
the wrong reason. Every lookup compares with `StrComp(..., vbBinaryCompare)`;
`vbTextCompare` appears nowhere (`test_29`).

An unknown Distribution is refused, never mapped to a default: the adapter has
no `Case Else`, and its caller fails on an unmapped kind (`test_30`).

**Phase-4 ownership is not reopened.** The `CL-`/`R-` prefixes, the counter
semantics and the structural pattern rules stay Phase 4's. Resolution needs only
enough text semantics to look up, order and reference.

---

## The empty driver set

No accepted contract requires at least one Cost Line or Risk, and Step 5 does
not invent one. `ResolveDrivers` returns success with `driverCount = 0`;
`ReferencedCurrencies`, `ReferencedProfiles`, `ResolveFxRates` and
`ResolveProfileWeights` each succeed on an empty set (`test_31`); each empty
branch precedes any bound access and each refuses a negative count (`test_32`);
and the reporting-currency invariant still runs (`test_33`).

---

## What Step 5 deliberately does NOT validate

`modCalcCheck` is next and owns the numerical prerequisites. Step 5 fails only
where resolution itself cannot proceed:

Three owners, not two:

| Owner | Concern |
| --- | --- |
| **Phase 4**, invoked by `StructuralPrerequisites` | the structural state; `ValidateStructure()` — ID patterns, duplicates, orphans, profiling and inflation grid shape, counters |
| **Step 5**, `modCalcResolve` | a referenced key has no matching source; ambiguous duplicate rows; a value cannot be read as the required type; a required value is blank; the reporting-currency invariant; an unknown distribution name; D2 (`1 + rate <= 0`) on a referenced rate |
| **Step 6**, `modCalcCheck` | `Min <= ML <= Max`; `Quantity > 0`; `0 <= Probability <= 1`; the profiling-sum tolerance |

Phase-4 structural prerequisites are **invoked** by the Step-5 entry path, never
duplicated. `modCalcCheck` will own only Phase-5 **numerical** prerequisites and
must not absorb the structural gate.

The Step-5 row follows the accepted oracle: its `_identifier`, `_numeric`,
`resolve_fx` and `resolve_inflation` — including that function's own
`1 + rate <= 0` refusal — are the resolution layer, and the range and ordering
rules enforced in `calculate` are not. One deliberate exception: the oracle's
`_resolve_weights` also checks the profiling sum, and that check is **not**
duplicated here — it needs `SafeSignedSum` over the assembled vector and belongs
with the tolerance authority, in the checker. That deferral is accepted.

---

## Build plumbing

`spec/structure_contract.yaml` now declares thirteen modules — the accepted
twelve plus `modCalcResolve`, `generated: false`. The generated set remains
exactly `{modConstants, modCalcContract}` (`test_02`).

**One further change, and it is the only one beyond the inventory.** Four facts
the resolver needs had no VBA projection, because Phase 4 manages structure and
never reads a rate, a discount or a distribution name:

| Projected now | From |
| --- | --- |
| `COL_FX_RATES_CURRENCY`, `COL_FX_RATES_FX_TO_SAR` | the FX table's own columns |
| `REPORTING_CURRENCY`, `REPORTING_CURRENCY_RATE` | the FX table's locked seed row |
| `DISTRIBUTION_NAME_1..3`, `DISTRIBUTION_COUNT` | the distributions Config table |
| `NM_INPUT_*` (including `NM_INPUT_DISCOUNT_RATE`) | the input contract's defined names |

The alternative was a second copy of the contract's own coordinates hand-written
into VBA, which §2 forbids. `stage_b_emit` now projects them from the existing
authorities; `modConstants.bas` gains those constants and nothing else. The
reporting currency is read from `tblFXRates`' locked seed row rather than
written into the emitter, so a change to that row is a change to the projection.

Which internal shape each distribution NAME selects remains an **adapter**, in
the resolver — exactly as it is an adapter in the Python oracle, and explicitly
not an authority.

---

## Retargeted invariants

Three existing tests named the world before this module existed. Each was
retargeted, not deleted.

| Test | Was | Now |
| --- | --- | --- |
| `test_05_no_module_is_a_dumping_ground` | three Phase-5 modules under the code/raw limits | four; the resolver is measured by the same policy |
| `test_02_step_4_added_exactly_three_modules_and_no_fourth` | inventory = Phase 4 + the three kernel modules | + exactly one Step-5 module, asserted in both directions, and `KERNEL_MODULES` is still exactly three |
| `test_08_the_deferred_phase_6_surface_does_not_exist_yet` | `modCalcResolve` on the deferred list, scanned in `.raw` | it left the list at Step 5; the scan now reads EXECUTABLE code, because the resolver legitimately names the checker when saying which prerequisites it leaves to it |
| `test_the_stage_b_manifest_carries_the_step_4_kernel_and_nothing_later` | the resolver must be ABSENT from the manifest | it must be PRESENT; the checker, the reporter and the endpoints must be absent. Renamed to `..._carries_the_implemented_modules_and_nothing_later` |

---

## Line metrics

| Module | Raw | Blank | Comment | Code | code < 900 | raw < 1200 |
| --- | --- | --- | --- | --- | --- | --- |
| `modCalcResolve` | 948 | 49 | 234 | 665 | yes | yes |

Before the correction the figures were 835/612. The three Step-4 modules are
byte-identical and their metrics are unchanged (1088/809, 1177/868, 511/284).

---

## Tests

| Module | Tests |
| --- | --- |
| `test_phase1_manifest_validation.py` | 10 |
| `test_phase1_structure.py` | 21 |
| `test_phase2_contract_validation.py` | 42 |
| `test_phase2_inputs.py` | 40 |
| `test_phase3_driver_contract_validation.py` | 31 |
| `test_phase3_drivers.py` | 28 |
| `test_phase3_verifier_intersection.py` | 12 |
| `test_phase4_oracle.py` | 68 |
| `test_phase4_stage_b_source.py` | 155 |
| `test_phase4_structure.py` | 43 |
| `test_phase4_structure_contract_validation.py` | 55 |
| `test_phase5_calc_contract_validation.py` | 151 |
| `test_phase5_fingerprint.py` | 52 |
| `test_phase5_numeric.py` | 94 |
| `test_phase5_oracle.py` | 111 |
| `test_phase5_stage_a.py` | 57 |
| `test_phase5_vba_source.py` | 120 |
| `test_phase5_resolve_source.py` | **79 (new)** |
| **Total** | **1169** |

The Step-4 baseline was 1090 and the first Step-5 submission 1146. The
correction adds 23 more static tests, all in the Step-5 suite. No test was
removed and none was weakened.

### Mutation evidence

Fourteen regressions were planted into a scratch copy of `modCalcResolve.bas`
and the suite was run against each. **All fourteen were caught**, none silently:

| Planted regression | Result |
| --- | --- |
| seed the resolved FX set with the reporting currency | caught |
| consult FX before the reference sets exist | caught |
| attach profiling weights by row position | caught |
| trim an identifier into the model | caught |
| fabricate a blank cell as zero | caught |
| refuse the empty driver set | caught |
| skip the reporting-currency invariant when nothing is referenced | caught |
| compare identifiers case-insensitively | caught |
| read the entered duration instead of the applied one | caught |
| compound the inflation series in the resolver | caught |
| suppress workbook errors with `On Error Resume Next` | caught |
| default an unknown distribution | caught |
| locate a year column by arithmetic instead of by header | caught |
| walk every inflation row instead of the referenced set | caught |

Fifteen more were planted for the correction round, and all fifteen were caught:

| Planted regression | Result |
| --- | --- |
| skip the structural gate entirely | caught |
| accept a `STATE_PENDING` structure | caught |
| accept a `STATE_NOT_APPLIED` structure | caught |
| refuse a state without a diagnostic | caught |
| drop the unrecognised-state branch | caught |
| ignore the `ValidateStructure()` report | caught |
| drop the D2 rejection | caught |
| coerce any identifier with `CStr` | caught |
| accept numeric-looking text in `NumericCell` | caught |
| accept numeric-looking text in `NumericNamedCell` | caught |
| let `Boolean` through the numeric type gate | caught |
| restore the combined inflation guard | caught |
| apply the strict numeric rule to a year header | caught |

The first attempt at the state-branch detector **missed** the emptied
`STATE_NOT_APPLIED` case: its window ran past the next `Case` and borrowed the
refusal belonging to the branch below. The detector now bounds each branch at the
next `Case` or `End Select`, and the mutation is caught.

Fourteen further negative controls plant the same defect classes as synthetic
module text and assert the sweeps see them.

### The suite guards its own language

`test_41` reads the file back and fails if it ever contains a phrase claiming
the VBA ran or a real workbook was read. The forbidden phrases are assembled
from parts so the guard cannot trip over its own wording.

---

## Error handling

There is no `On Error Resume Next` in `modCalcResolve`, and no `On Error` at all
(`test_37`): every failure is a returned `False` with a diagnostic. Nothing
substitutes a zero, a blank string, an identity factor, a default distribution,
the reporting currency or the first matching row for a value it could not read.

No message box, no user-facing refusal handling: status and detail cross the
boundary and later orchestration owns what a user is told.
