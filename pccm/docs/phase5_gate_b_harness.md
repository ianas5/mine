# Phase 5 — Gate B — Step B1: the Windows harness extension

**Status: correction round 4 — harness source, ready for independent review.
NOTHING HAS BEEN RUN.**

Gate A is accepted and closed at `1968fb8`. This step authors the Windows Gate-B
harness required by `phase5_plan.md` §24, §25 and implementation-sequence item 11,
and reviews it statically on Linux. The sequence is deliberate: author, review on
Linux, review independently, and only then run real Excel on Windows.

The first submission (`93f306d`) had its architecture accepted and was rejected
with **eight harness defects**, recorded in full under
**[Correction round 1](#correction-round-1)**. Correction round 1 (`aa18cab`)
closed all eight and was rejected with **five more**, recorded under
**[Correction round 2](#correction-round-2)**: incomplete plan section 18
prerequisite coverage, a status row that did not compare the analytical
snapshot, a row-order probe that changed two dimensions, a missing driver-audit
reconstruction, and an application-state proof that checked Excel's defaults
rather than the caller's own state. Correction round 2 (`2a2ae86`) closed all
five and was rejected with **four more**, recorded under
**[Correction round 3](#correction-round-3)**: a fixture loader that bypassed the
Phase-4 owner of inflation profile rows, a fixture that never proved its own
structural prerequisites, three prerequisite mutations that could not reach the
predicate they claimed, and an audit cross-check required to be exact that used a
relative tolerance. Correction round 3 (`aa6611c`) closed all four and was
rejected with **two more**, recorded under
**[Correction round 4](#correction-round-4)**: a locked FX seed that was never
restored between scenarios, and a referenced-only proof that showed non-blocking
without showing no-effect. Every one was a defect in what the
harness would have PROVED, not in how it is wired: two fixture writers that
destroyed the condition they were exercising, a refusal proof that asserted the
opposite of the rule, a reconciliation block that reimplemented a rejected
oracle, a staleness proof with no oracle at all, a reorder proof that reordered
nothing, a missing end-to-end fingerprint parity, two published columns never
asserted, and semantic values in an address projection.

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

Correction round 1 removed three semantic values that had been in it and replaced
name-banning with a positive schema; see
**[correction 8](#8--the-inspection-projection-was-not-identities-only)**. The
schema is now:

```
schema_version, purpose, provenance
calc          -> sheet, required_visibility, tables, scalar_blocks
  tables[*]   -> table_name, header_row, first_column, last_column,
                 first_column_index, column_count, first_body_row,
                 row_rule, columns
  blocks[*]   -> label_column, value_column, first_row, last_row,
                 value_range, rows
inputs[*]     -> defined_name, sheet, cell, type
input_tables[*] -> table_name, sheet, header_row, first_column, columns,
                   locked_seed_rows
```

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

## Correction round 1

Independent review reproduced 76/76 and 351/351 from the submitted package,
accepted the architecture, and rejected the harness on eight defects.

### 1 — blank fixture values were coerced to zero

`Write-Phase5InflationRates` and `Write-Phase5Weights` both cast before they
branched:

```powershell
-Value ([double]$rates.$year)      # [double]$null is numeric ZERO
-Value ([double]$weight)
```

So plan case 14 — *blank required inflation rate* — was written as a rate of
**0**, and plan case 23 — *a profile summing to 100% but containing a blank* —
was written as **50% / 0% / 50%**. Both are VALID models. The refusal each case
exists to prove could never have fired: the fixture was silently destroying the
condition it was written to exercise, and both cases would have sat in the
ledger claiming coverage they did not have.

Both writers now branch on `$null` **before** the cast and write a genuine blank
through `Set-TableCell -Value $null`, which calls `ClearContents` — never `0`,
never `""`. `test_48` proves the guard precedes the conversion in both
procedures; `test_49` ties the proof to cases 14 and 23 and re-checks that the
corpus still carries their blanks and that case 23's remaining weights still sum
to one, so a zero in place of the blank could not be refused for the wrong
reason. `test_nc_29`, `test_nc_30` and `test_nc_31` plant the shipped shape and
a guard placed after the cast.

### 2 — the refusal proof asserted the opposite of the rule

P5-RF required every `_Calc` table to hold **zero populated rows** after a
refusal. That would have **failed against correct production behaviour**: P5-AN
runs first and leaves a successful snapshot, `Set-Phase5Fixture` changes the
INPUT model and never touches `_Calc`, and a pre-write refusal is required to
leave C13:C16, C23:C32 and all five tables exactly as they were.

*No partial result* means **no partial NEW snapshot survives**, not *erase the
old successful one*.

P5-RF now establishes a successful baseline from plan case 3, captures the full
snapshot, applies each refusal fixture **without clearing `_Calc`**, and after
each one asserts INVALID / REFUSED / specific detail, then C13:C16 + C23:C32 +
all five tables unchanged, then C17:C20 changed. The baseline is captured once
and every refusal is compared against it, which additionally proves that a run of
nine successive refusals never erodes the snapshot. `Get-Phase5Snapshot`,
`Add-SnapshotUnchangedChecks` and a new `Add-Phase5AttemptAxisChecks` were
factored to file scope so the refusal, rollback and status scenarios make the
SAME comparison rather than three different ones.

### 3 — the identity block relabelled §15 and reimplemented a rejected oracle

The block checked five relabelled identities and decided each with its own
PowerShell tolerance:

```powershell
max(|left|, |right|, floor) * identity_relative_coefficient
```

That is the **headline-based conditioning erratum C1 rejected**, and plan case 30
exists precisely because that shape can falsely fail a correct cancellation-heavy
calculation. Reintroducing it in PowerShell would have made a rejected oracle the
acceptance authority. It also never asserted annual Base or Risk separately, and
never asserted I5 at all.

P5-ID now carries the locked mapping — **I1** (A+B=C, nominal and PV), **I2**
(C+D=E, nominal and PV), **I3a/b/c** (annual Base / Risk / Total nominal →
C_nom / D_nom / E_nom), **I4a/b/c** (the same in present value), **I5** (profile
weights per driver) — and the evidence is:

* **production accepted the fixture.** `Reconcile` and `AllIdentitiesHold` run
  inside `PCCM_Calculate`, and a commit is unreachable unless they pass. A
  SUCCESS on case 30 is production's own statement that the identities held under
  the accepted C1 allowance.
* **every published value equals the emitted oracle**, through the same
  `Add-Phase5AnalyticalChecks` path as everywhere else.
* **each annual column is asserted separately**, row by row, against its own
  emitted value.
* **I5** is the per-driver weight vector, asserted in the profiling grid against
  the emitted corpus; the other half of I5 is the refusal matrix, where cases 15
  and 23 prove an invalid vector is refused.

No new PowerShell tolerance decides anything. `test_52` refuses
`[Math]::Max([Math]::Max(`, `identity_absolute_floor`, `$close = {` and
`[Math]::Abs` anywhere in the identity block; `test_nc_34` plants the shipped
helper.

### 4 — the staleness recalculation had no oracle

§25.2 requires the affected value to change **to the oracle value**. The harness
made the model stale by exchanging two profiling weights, which produces a model
the corpus does not describe — so after the second `PCCM_Calculate` there was
nothing to compare against, and the executable proof was an annual **row count**.

Plan cases **3** and **19** have identical applied structure — same timeline, FX,
inflation profile and rates, driver and weights — and differ only in Discount
Rate (`0.10` → `-0.05`). The sequence is now

```
establish case 3, calculate, capture digest and state
change ONE ordinary fingerprinted scalar: inpDiscountRate -> case 19's value
assert STALE / previous SUCCESS / stored digest old / current digest changed
calculate
assert CURRENT / SUCCESS / stored digest changed
Add-Phase5AnalyticalChecks against case 19's OWN emitted expected block
Add-Phase5SuccessStateChecks against case 19
```

so every published value of the recalculated model is an oracle value. No Apply
Timeline is used to create staleness, and nothing is hand-calculated. The harness
also **checks** that the two fixtures differ in exactly one scalar, so a corpus
change cannot silently turn this into a two-variable transition.

### 5 — the row-order proof reordered nothing, and the probes were cumulative

P5-NS ran on plan case 3, which has **one** Cost Line. The `ListObject.Sort` call
was real; sorting a one-row table changes nothing, so canonical Permanent-ID
ordering was never exercised. And each probe left its change in place while the
next one ran, so by the fourth probe four edits were live and no exclusion had
been isolated.

The reorder probe now runs on plan case **30**, which has three Cost Lines. It
writes distinct sort keys, captures the actual permanent-ID order **before** and
**after** the sort, and requires the order to have **changed** — `"Sort was
called"` is not accepted as evidence — while also requiring the same identifiers
to be present, only reordered.

All four probes are now independent:

```
baseline: CURRENT / SUCCESS / digest F
change ONE excluded input
assert CURRENT / SUCCESS / F
restore exactly
assert the baseline is back
```

The probe takes the digest it must hold against as an **argument**, so a later
probe cannot inherit an earlier probe's edit. `test_55` requires all four
restorations and the per-probe digest parameter; `test_nc_39` and `test_nc_40`
plant the one-row fixture and the cumulative shape.

### 6 — the end-to-end golden fingerprint parity was missing

P5-D5 proves the emitted reference STREAM digests correctly through
`CalcFpDigestStream`. That is necessary and not sufficient: P5-AN only proved
`stored == current`, which **two identically wrong production fingerprints would
satisfy**.

Plan case 1 is the model the reference stream was built from, so the complete
production path — resolution, referenced factors, record construction, canonical
section ordering, `BuildFingerprint` — must land on the emitted digest. P5-AN now
asserts, for case 1 only:

* `PCCM_CurrentInputFingerprint()` equals `fingerprint.reference.digest`
* `PCCM_CalculationFingerprint()` after the commit equals the same value

read from the corpus; the literal `50B6EB0E26857EA7` remains refused anywhere in
the harness. Both proofs are required and both survive: the direct primitive one
and the end-to-end one are different claims.

**UTF-16 fields are now compared in full.** P5-D4 checked only
`StartsWith("S<units>:")`, which passes a mangled payload of the right length —
exactly the failure a surrogate-pair vector exists to catch. The corpus already
emits `canonical_text_field` for every vector, and the complete field is now
compared ordinally, with the prefix check kept as a supplementary claim.

### 7 — the analytical audit was not actually complete

Two published columns had no emitted expectation behind them, so the harness
could only have asserted them by deriving the answer in PowerShell:

| Column | Was | Now |
| --- | --- | --- |
| `tblCalcYears.Calendar Year` | unasserted | `expected.calc_years[].calendar_year` |
| `tblCalcFX.Referenced By` | unasserted | `expected.resolved_fx_rows[].referenced_by` |

Both come from the **existing** oracle: `AppliedTimeline.project_years()` already
owns the calendar year, and the model's own driver list already owns the
reference count. No oracle algorithm changed; `evaluate()` now surfaces what
`calculate` already had in hand. The Stage-A golden-independence ledger validates
both blocks against an independently derived reference, so the new fields are
held to the same standard as every other emitted expectation.

The **successful `calc_state` record** is now asserted cell by cell — C13
non-blank, C14 exactly the digest the API returned, C15 the emitted
`FP_VERSION`, C16 the emitted `applied_timeline` text, C17 `SUCCESS`, C18 blank,
C19 `CURRENT`, C20 non-blank — with the two timestamps deliberately **not**
required to equal each other. The applied-timeline format belongs to
`modCalcReport.AppliedTimelineText`; the corpus carries a checked copy and
`test_60` pins it against that procedure.

`test_58` now derives the required coverage from the inspection projection's own
declared columns and fails if any of them has no emitted expectation, rather than
checking that each table NAME appears.

### 8 — the inspection projection was not identities-only

Its `calc` object carried `fingerprint_version`, `derived_status_labels` and
`attempt_result_labels`. A version number is an expected VALUE and the two label
lists are model SEMANTICS; none is an address. All three are removed, and the
schema version is now 2.

They were not moved into PowerShell literals: the version is read from
`phase5_cases.json → fingerprint.constants.FP_VERSION`, and the status and
attempt vocabularies are stated in the matrix rows that assert them, which is
where the matrix states them.

`test_41` and the new `test_61` no longer rely on banning names. The emitter
declares a **positive schema** — `ALLOWED_ROOT_KEYS`, `ALLOWED_CALC_KEYS`,
`ALLOWED_TABLE_KEYS`, `ALLOWED_BLOCK_KEYS`, `ALLOWED_INPUT_KEYS`,
`ALLOWED_INPUT_TABLE_KEYS` — and the tests assert set equality at every level, so
the next semantic value is refused whatever it is called. `test_nc_47` plants
both the removed values and an innocent-looking `default_precision` to show the
ban-list would have missed the second and the allowlist does not.

One nuance the tests state explicitly: `fingerprint_version` survives as a
`calc_state` **row name**, because row 15 is where the version is written and
that is an address. Banning the string everywhere would have refused an identity
along with the value.

---

## Correction round 2

Independent review reproduced 110/110 and 351/351, accepted all eight round-1
fixes as closed, and rejected the harness on five more defects.

### 1 — the Windows refusal coverage did not cover all of plan section 18

P5-RF ran the nine refusal cases the 37-case plan corpus carries. Those nine are
valid and remain — but they do not exhaust section 18. **Base Year after Start
Year, STRUCTURE CHANGE PENDING, a duplicated referenced currency, a
non-numeric Probability, an unknown Distribution and a dozen more locked
predicates had no real-Windows scenario at all.**

Most of those boundaries cannot be expressed as valid analytical models — `"abc"`
is not a Discount Rate, a blank is not a Quantity — and the typed oracle is right
to refuse them at its own boundary. So the new corpus describes **workbook
mutations** rather than pretending such values are valid cases. The Python oracle
is not weakened and no analytical expectation is invented for a refused model.

**The 37 plan cases are untouched.** The matrix lives in a separate emitted
section, `phase5_cases.json → gate_b`:

```
gate_b
  schema_version, purpose
  prerequisite_cases[]   id, section, predicate, title, base_plan_case,
                         mutation{kind, ...}, detail_tokens[],
                         expected_attempt, expected_status, snapshot_unchanged
  no_block_cases[]       same shape; expected SUCCESS / CURRENT, no tokens
  plan_refusal_tokens{}  the discriminators for the nine refusal plan cases
  audit_reconstruction   title, model, expected, relationships[]
```

The `mutation.kind` vocabulary the harness implements is
`entered_structure`, `named_number`, `named_text`, `named_blank`,
`register_cell`, `fx_row`, `fx_remove`, `inflation_cell`,
`inflation_profile_rename`, `inflation_profile_add` and `profiling_cell`.
An unknown kind throws rather than being ignored. **PowerShell holds no list of
its own**: `test_66` refuses every locked predicate name in the harness source
and requires every emitted mutation kind to be one the applier implements.

A `null` in the matrix is a **blank cell** — the same rule the fixture writers
follow, because several locked prerequisites *are* the blank. A blank Setup
scalar goes through `ClearContents`, never `''`.

**Specific detail discrimination.** P5-RF previously proved only that the detail
was non-empty, which cannot tell a refusal from the intended predicate apart from
any other refusal. Every scenario now asserts a token set drawn from the accepted
production message, and `test_65` proves each token actually appears in
`modCalcResolve.bas`, `modCalcCheck.bas`, `modCalcFactors.bas`,
`modCalcReport.bas`, `modAppState.bas`, `modStructuralCheck.bas` or the
generated constants — identifiers the fixture itself supplies (`USD`, `SAR`,
`CL-001`, `2027`, `Standard`) excepted. Whole sentences are not frozen, so a
harmless wording edit does not break the proof while a refusal from the wrong
predicate does.

**The referenced-only complement.** A harness that only proved refusals would
accept a model that refused too much, so three no-block cases prove SUCCESS,
CURRENT, a blank detail and an **unchanged digest** for an unreferenced duplicate
currency, an unreferenced currency with a blank rate, and an unreferenced
incomplete inflation profile.

### The complete prerequisite ledger

| Locked predicate | Condition | Gate-B scenario | Detail discriminator |
| --- | --- | --- | --- |
| `18.T1` | Base Year later than Start Year, through modCalcCheck directly | `DC-01` → `P5-DC` *(direct)* | `Base Year`, `Start Year` |
| `18.T2` | an entered structural value changed and the timeline was not re-applied | `PQ-02` → `P5-PQ` | `STRUCTURE CHANGE PENDING`, `structural prerequisite` |
| `18.D1` | Discount Rate blank | `PQ-03` → `P5-PQ` | `Discount Rate`, `blank` |
| `18.D2` | Discount Rate non-numeric | `PQ-04` → `P5-PQ` | `Discount Rate`, `not numeric` |
| `18.F1` | a referenced foreign currency has no tblFXRates row | `PQ-05` → `P5-PQ` | `FX`, `USD`, `rows` |
| `18.F2` | a referenced foreign currency appears twice | `PQ-06` → `P5-PQ` | `FX`, `USD`, `rows` |
| `18.F3` | a referenced FX rate is not strictly positive | `PQ-07` → `P5-PQ` | `FX`, `USD`, `strictly positive` |
| `18.F4` | a referenced FX rate is blank | `PQ-08` → `P5-PQ` | `FX rate for referenced currency`, `USD`, `blank` |
| `18.F5` | a referenced FX rate is non-numeric | `PQ-09` → `P5-PQ` | `FX rate for referenced currency`, `USD`, `not numeric` |
| `18.F6` | the reporting currency has no row | `PQ-10` → `P5-PQ` | `the reporting currency`, `SAR`, `exactly once` |
| `18.F7` | the reporting currency appears twice | `PQ-11` → `P5-PQ` | `the reporting currency`, `SAR`, `exactly once` |
| `18.F8` | the reporting currency rate is not exactly 1 | `PQ-12` → `P5-PQ` | `the reporting currency`, `SAR`, `must resolve to` |
| `18.I1` | a driver references an inflation profile that does not exist | `PQ-13` → `P5-PQ` | `inflation: profile`, `MissingGateBProfile`, `not present` |
| `18.I2` | a required referenced inflation rate is non-numeric | `PQ-14` → `P5-PQ` | `inflation profile`, `2027`, `not numeric` |
| `18.P1` | a required profiling cell is non-numeric | `PQ-15` → `P5-PQ` | `profiling for driver`, `CL-001`, `not numeric` |
| `18.X1` | Distribution is blank | `PQ-16` → `P5-PQ` | `Distribution`, `blank` |
| `18.X2` | Distribution is not one of the three accepted kinds | `PQ-17` → `P5-PQ` | `Distribution`, `not an accepted distribution` |
| `18.O1` | Triangular Min <= Most Likely <= Max is violated | `PQ-18` → `P5-PQ` | `Triangular`, `Min <= Most Likely <= Max` |
| `18.O2` | Beta-PERT Min <= Most Likely <= Max is violated | `PQ-19` → `P5-PQ` | `Beta-PERT`, `Min <= Most Likely <= Max` |
| `18.O3` | Uniform Min <= Max is violated | `PQ-20` → `P5-PQ` | `Uniform`, `Min <= Max` |
| `18.Q1` | Quantity is blank | `PQ-21` → `P5-PQ` | `Quantity`, `blank` |
| `18.Q2` | Quantity is non-numeric | `PQ-22` → `P5-PQ` | `Quantity`, `not numeric` |
| `18.R1` | Probability is blank | `PQ-23` → `P5-PQ` | `Probability`, `blank` |
| `18.R2` | Probability is non-numeric | `PQ-24` → `P5-PQ` | `Probability`, `not numeric` |
| `18.R3` | Probability below zero | `PQ-25` → `P5-PQ` | `Probability`, `fraction in [0, 1]` |
| `18.R4` | Probability above one | `PQ-26` → `P5-PQ` | `Probability`, `fraction in [0, 1]` |
| `18.N1` | an UNREFERENCED foreign currency appearing twice does not block | `PN-01` → `P5-PN` | *(SUCCESS, detail blank)* |
| `18.N2` | an UNREFERENCED foreign currency with a blank rate does not block | `PN-02` → `P5-PN` | *(SUCCESS, detail blank)* |
| `18.N3` | an UNREFERENCED inflation profile with a missing rate does not block | `PN-03` → `P5-PN` | *(SUCCESS, detail blank)* |
| plan case 14 | blank required inflation rate | `14` → `P5-RF` | `inflation profile`, `blank` |
| plan case 15 | profile does not sum to 100% | `15` → `P5-RF` | `profiling weights sum to` |
| plan case 16 | Quantity of zero | `16` → `P5-RF` | `Quantity`, `strictly positive` |
| plan case 17 | negative Quantity | `17` → `P5-RF` | `Quantity`, `strictly positive` |
| plan case 18 | discount rate of -100% | `18` → `P5-RF` | `discount rate`, `1 + r <= 0` |
| plan case 20 | inflation rate of -100% | `20` → `P5-RF` | `inflation profile`, `1 + rate <= 0` |
| plan case 23 | profile summing to 100% containing a blank | `23` → `P5-RF` | `profiling for driver`, `blank` |
| plan case 24 | controlled Double overflow | `24` → `P5-RF` | `inflation factor` |
| plan case 29 | discount factor underflow | `29` → `P5-RF` | `discount factors` |

### 2 — status row 2 did not prove the analytical snapshot unchanged

Row 2 proved the accessor axis and C13:C16. **A defect where
`PCCM_CalculationStatus()` rewrote analytical outputs while re-deriving the
status would have passed it**, because C23:C32 and the five ListObjects were
never compared.

A full `Get-Phase5Snapshot` baseline is now captured **before** the
fingerprinted input changes, and `Add-SnapshotUnchangedChecks` compares
C13:C16, C23:C32 and all five tables afterwards. C17:C20 is handled **separately**
— C19 and C20 are deliberately refreshed by the status evaluation, so asserting
the whole block unchanged would assert that asking for the status did nothing.
Row 2 now states all four: C17 still SUCCESS, C18 still blank, C19 re-derived to
STALE, C20 carrying a fresh timestamp.

### 3 — the row-order probe changed Description as well as row order

The corrected round-1 probe used a real multi-row fixture and proved the physical
order changed — that part stands. But it first **rewrote every Description** to
`sort-key-…` to force the ordering, so it changed **two** non-fingerprinted
dimensions and stopped being a row-order-only proof.

`Write-Phase5Driver` already gives every row a deterministic, distinct
Description — `GateB <PermanentId>` — so the existing values are sufficient sort
keys. The probe now edits **no cell at all**: it captures the permanent-ID order,
sorts descending on the existing Description values, captures the order again,
requires it to have changed with the same identifiers present, asserts CURRENT /
SUCCESS / unchanged digests, then sorts ascending and asserts the original order
is back. `test_70` fails if `Set-TableCell` appears anywhere in that probe.

### 4 — the driver-audit A/B/C/D reconstruction was missing

Every driver row and every headline total was compared against the Python oracle,
and that remains. But the locked plan additionally requires a **cross-check
between two parts of the real workbook**: the published audit columns must
reconstruct the published headline totals.

| Headline | `tblCalcDrivers` column | Ordinal | Partition |
| --- | --- | --- | --- |
| `a_nom` / `a_pv` | `deterministic_nominal` / `deterministic_pv` | 14 / 15 | Cost Line |
| `b_nom` / `b_pv` | `uncertainty_mean_shift_nominal` / `_pv` | 18 / 19 | Cost Line |
| `c_nom` / `c_pv` | `mean_basis_nominal` / `mean_basis_pv` | 16 / 17 | Cost Line |
| `d_nom` / `d_pv` | `expected_risk_nominal` / `expected_risk_pv` | 20 / 21 | Risk |

`P5-AR` drives a Gate-B-only multi-driver fixture — **three Cost Lines and two
Risks**, so none of A, B, C or D is trivial — reads the ACTUAL `tblCalcDrivers`
body, partitions it by Driver Kind, sums each audit column, and compares against
the ACTUAL `calc_totals` cell. The mapping is emitted in
`gate_b.audit_reconstruction.relationships`, so PowerShell never restates
"column 18". **An N/A blank is skipped, never folded in as the opposite kind's
identity 1**, and the scenario separately asserts that the opposite kind
publishes BLANK in each column. No new tolerance is invented: this is an audit
relationship over the same published Doubles, not the reconciliation allowance.

The fixture's `expected` block goes through the accepted oracle and is validated
by the Stage-A golden-independence ledger like every other emitted expectation.

### 5 — the rollback application-state proof checked defaults, not restoration

The failpoint scenarios asserted `EnableEvents = True`, `ScreenUpdating = True`
and `Calculation = Automatic` after the injected failure. That proves only that
the application happens to be in convenient defaults — which it would be even if
`FinishOperation` restored nothing. The accepted Phase-4 scenario S already
rejected this pattern.

Each failpoint scenario now establishes a **deliberately non-default caller
state** before arming the failpoint:

```
ScreenUpdating = False
EnableEvents   = False
DisplayAlerts  = False
Calculation    = xlCalculationManual (-4135)
StatusBar      = "PCCM Phase-5 rollback sentinel <failpoint name>"
```

captures those exact values, runs the failed calculation, and asserts **all five**
against the captured values — **before** normalising anything for the scenarios
that follow. The StatusBar sentinel carries the failpoint name, so a value left
by the other scenario cannot satisfy it, and a further check requires the
captured state to differ from Excel's defaults so the proof cannot be vacuous.
Both `P5-FA` and `P5-FC` go through the same shared runner, which captures
independently on each invocation — neither inherits the other's proof.

---

## Correction round 3

Independent review reproduced 134/134 and 351/351, accepted all five round-2
fixes as closed, and rejected the harness on four more.

### 1 — the fixture bypassed the Phase-4 owner of inflation profile rows

`Config!tblInflationProfiles` is the Phase-4 source of truth for inflation
profile identities, and `modInflation.SyncProfileRows` **rebuilds**
`Inflation!tblInflation` from that master inside `PCCM_ApplyTimeline`. The
loader was writing profile names straight into `tblInflation` **before** Apply —
so the very Apply the fixture depends on deleted the rows it had just planted,
and `Write-Phase5InflationRates` then searched for a row production had already
removed. A harness fixture-construction defect, not a calculation defect.

**The corrected `Set-Phase5Fixture` sequence:**

```
1  clear previous fixture drivers through the production delete endpoints,
   and reset the ID counters
2  write the Setup scalars (Base/Start/Duration entered, Discount Rate)
3  write the fixture FX assumptions above the locked reporting-currency seed
4  synchronise Config!tblInflationProfiles to EXACTLY the fixture's profile set
5  add Cost Lines and Risks through PCCM_AddCostLine / PCCM_AddRisk and write
   their register cells
6  PCCM_ApplyTimeline  -- production creates the structure:
       SetYearColumns   the profiling and inflation year bands
       SyncProfileRows  the Inflation rows, from the Config master
       profiling rows   synchronised by permanent ID
   ... and the fixture PROVES this step succeeded (see 2 below)
7  only now: blank the rate cells, write the annual inflation rates keyed by
   PROFILE NAME x CALENDAR YEAR, and write the profiling weights keyed by
   permanent ID
```

`Set-Phase5InflationProfileMaster` clears the editable master rows, writes
exactly the distinct fixture profile names as exact binary text, and writes **no
rate**. The table identity comes from
`phase5_gate_b_inspection.json → input_tables.inflation_profiles`; `test_75`
refuses the literals `'Config'` and `'tblInflationProfiles'` anywhere in the
harness, and refuses any `tblInflation` seeding in the fixture. No Phase-4
module was edited and no harness-only structural sync exists: the point is to
drive the existing production mechanism correctly.

**Rate placement is keyed on both axes.** The row was an incremented counter,
which assumed the emitted model's profile order equals the physical grid order —
it does not, because `SyncProfileRows` rebuilds in **Config-master** order.
The row is now found by **inflation profile name** through `Find-GridRow`, and
the column by **calendar-year header** as before. Every rate cell is blanked
first, because `SyncProfileRows` deliberately preserves a surviving profile's
rates by name and a value left by an earlier fixture must not be inherited as
this one's assumption. `test_nc_64` demonstrates the order-independence
positively: with the model listing `Standard, Flat` and the grid materialising
`Flat, Standard`, the keyed writer places each rate on its own profile and the
positional one puts Standard's rate on Flat's row.

### 2 — the fixture did not prove its own structural setup succeeded

The loader discarded `PCCM_AutomationResult` after `PCCM_ApplyTimeline` and
carried straight on into the rate writer, the weight writer and
`PCCM_Calculate`. A refused or failed Apply, or a failed structural
revalidation, would therefore have surfaced later as a *calculation* defect
rather than as the fixture fault it was.

`Set-Phase5Fixture` now **throws** — twice, with the offending text in the
exception — if the Apply did not return `OK|…`, or if `PCCM_StructuralReport()`
is not blank. Both gates precede the value writers; `test_76` asserts the
ordering by index.

**This gate applies to the clean base fixture only.** A later prerequisite
mutation is *meant* to make an input invalid, and `Invoke-Phase5Mutation` imposes
no such gate — except where an entry explicitly asks for one via
`require_clean_structure`, which only the unreferenced-profile no-block case
does. `test_77` asserts that no prerequisite case carries that flag.

### 3 — three prerequisite mutations could not reach their predicate

A token that appears somewhere in production source proves **vocabulary**, not
that the chosen mutation reaches the procedure that emits it. Three entries were
not semantically reachable as written.

**`base_year_after_start_year`.** `modTimeline` **prevalidates** Base > Start
and refuses the Apply without moving the applied timeline, so the workbook is
left with entered ≠ applied and the next `PCCM_Calculate` is refused by
`StructuralPrerequisites` with STRUCTURE CHANGE PENDING — a different predicate,
in a different module, with a different message. The predicate has left the
mutation matrix and is proved **directly**: `P5-DC` calls the transient
diagnostic's `GBD_CheckBaseAfterStart`, which builds a `ResolvedModel` on
target with `DriverCount = 0` and calls the already-Public
`modCalcCheck.CheckResolvedModel`. A control, `GBD_CheckTimelineAccepted`,
requires the same construction with Base ≤ Start to be **accepted**, so the
refusal cannot come from something else the harness built wrong. It runs before
the diagnostic module is removed, and no production visibility is reopened.
STRUCTURE CHANGE PENDING remains its own separate workbook scenario.

**`referenced_profile_missing`.** Renaming a `tblInflation` row breaks the
Phase-4 invariant *"Inflation rows match the Config profile master"*, so
`ResolveModel` refuses at `ValidateStructure` and the Step-5 lookup is never
reached. The mutation now moves the **driver's reference** instead —
`cost_lines.CL-001.inflation_profile → "MissingGateBProfile"` — leaving Phase-4
structure coherent. A direct COM write bypasses the cell's Data Validation, which
is the point of an invalid-input runtime test. `ResolveDrivers` reads the exact
text, `ReferencedProfiles` contains it, and `ResolveInflationRates` cannot find
it in `tblInflation`: exactly the resolver refusal the tokens name.

**`unreferenced_profile_incomplete`.** Adding a row straight to
`tblInflation` creates the same master/grid mismatch from the other side. The
profile is now declared in `Config!tblInflationProfiles` and
`PCCM_ApplyTimeline` is re-run through the accepted automation path, so
`SyncProfileRows` creates the matching Inflation row — with **blank rates by
construction**, because a profile it has not seen before gets a cleared slot.
`PCCM_StructuralReport()` is required to be blank before the calculation, which
is half the claim: the profile is structurally legitimate and numerically ignored
because nothing references it. The calculation must then be SUCCESS / CURRENT
with a blank detail and the baseline digest unchanged.

The token-vocabulary check (`test_65`) is kept — it still catches a
discriminator that names nothing in production — but it is **vocabulary
evidence, not proof of runtime reachability**, and `test_79`, `test_80` and
`test_81` now pin each of these three ownership boundaries to the route that
actually reaches it.

### 4 — the audit cross-check used a relative tolerance

`P5-AR` compared the reconstructed sum against the published headline with
`-Tolerance $Cases.tolerances.identity_relative_coefficient` — 1e-12. That is
not the production C1 allowance, but it is still an epsilon on a relationship
required to be **exact**, and it could hide a small but real mismatch between two
values the workbook publishes.

The emitted Gate-B audit fixture reconstructs to the **identical IEEE Double**
for all eight relationships — `test_82` asserts that with `==`, so exactness
is representable rather than aspirational. The comparison is now
`-Tolerance 0.0`, passed explicitly so the intent is unmistakable, and
`test_82` refuses `identity_relative_coefficient`,
`identity_absolute_floor`, `conditioning_scale_floor`,
`profiling_sum_absolute`, any non-zero `-Tolerance`, and `[Math]::Abs` /
`[Math]::Max` in that block. This is an audit relationship — *does the audit
table reconstruct the headline the workbook publishes?* — not I1–I4
reconciliation, which keeps its own production allowance untouched.

---

## Correction round 4

Independent review reproduced 151/151 and 351/351, accepted all four round-3
fixes as closed, and rejected the harness on two more.

### 1 — the locked FX seed was never restored between scenarios

`tblFXRates` row 1 is the reporting currency's own row, built by Stage A as a
**locked seed**. The fixture reset was `Clear-Phase5UserRows -KeepRows 1`, which
preserves whatever is physically in row 1 — and the Gate-B prerequisite matrix
deliberately destroys it.

| Mutation | What it leaves behind | What the next fixture inherited |
| --- | --- | --- |
| `PQ-10` removes the reporting row | the foreign `USD` row shifts **up into row 1** | row 1 preserved as though it were the seed; the model's own reporting entry skipped and a **second** `USD` row appended |
| `PQ-12` sets the reporting rate to `2` | row 1 is `SAR / 2` | every later fixture inherits `SAR = 2` and refuses on the **global reporting-currency invariant** instead of the predicate it claims to test |

Deterministic cross-scenario contamination, not a race.

**The seed is now captured once, from the real workbook.** `P5-FX` runs
immediately after the Phase-4 prerequisite is proved intact — the last moment the
workbook is guaranteed untouched by Phase 5 — and reads row 1's Currency and
FX-to-SAR out of the Stage-B workbook that passed Stage-A verification, the
Stage-B persistence checks and the 35-scenario Phase-4 matrix.

The capture helper is `Save-Phase5LockedFxSeed`; the accessor
`Get-Phase5LockedFxSeed` is the only way a fixture reaches it, and it throws when
the capture never ran.

It is **not** reconstructed. `test_83` refuses `'SAR'`, `"SAR"`, `$Model.fx`,
`$Cases.`, `REPORTING_CURRENCY` and a hard-coded rate of 1 anywhere in the
capture or the reset: rebuilding the seed as *"SAR, 1"* would make the fixture
manufacture the very invariant `PQ-10` … `PQ-12` exist to test, and if the
**built** seed is wrong the analytical scenarios must still fail rather than be
repaired into agreement. The capture is reported as a note, never asserted
against a literal. `Get-Phase5LockedFxSeed` throws if the capture never ran, so a
future reordering fails loudly instead of silently reverting to the old
behaviour.

**The reset sequence**, `Reset-Phase5FxTable`, called at the top of the FX step
in `Set-Phase5Fixture`:

```
1  read tblFXRates from the inspection projection
2  ensure at least one physical body row exists
3  remove every body row after row 1, whatever it is
4  REWRITE row 1 from the capture: Currency, then FX to SAR
5  read row 1 back and throw if it did not restore
6  only then append the current fixture's non-reporting FX rows
```

Step 4 is the point: row 1 is rewritten, never trusted. `test_85` asserts the
reset precedes the first `Add-BlankTableRow` and that `Clear-Phase5UserRows` is
gone from the fixture entirely.

**This is harness isolation, not a production repair.** No production VBA
changed, no reset endpoint was added, and the single COM/workbook lifecycle is
untouched — the helper undoes only what the harness itself did to the workbook it
keeps reusing. And it does not disarm the mutations: `PQ-10` still physically
removes the reporting row, `PQ-11` still appends a duplicate, `PQ-12` still
rewrites the rate, and the reset happens only when the **next** clean fixture is
established. `test_86` pins all three mutation shapes and the
establish → mutate → calculate ordering, and refuses `Reset-Phase5FxTable` inside
the mutation applier.

### 2 — the referenced-only proof showed non-blocking, not no-effect

`P5-PN` proved SUCCESS, CURRENT, a blank detail and an unchanged digest after an
unreferenced mutation. Those are necessary and not sufficient: **a defect that
kept the unreferenced assumption out of the fingerprint while consuming it in the
calculation would satisfy every one of them and still publish wrong numbers.**
Referenced-only means the assumption is outside the calculation model, not merely
outside the digest.

Each no-block scenario now runs:

```
Set-Phase5Fixture(base model) -> PCCM_Calculate -> capture the baseline digest
apply the unreferenced mutation
PCCM_Calculate again
assert SUCCESS / CURRENT / blank detail / digest unchanged
Add-Phase5AnalyticalChecks   against the SAME emitted base plan case
Add-Phase5SuccessStateChecks against the SAME emitted base plan case
```

so the complete analytical workspace — `tblCalcYears`,
`tblCalcInflationFactors`, `tblCalcFX`, `tblCalcDrivers`, `tblCalcAnnual` and all
ten `calc_totals` cells — is compared against the base case's own emitted oracle
**after** the mutation. No new corpus was needed: every base plan case already
carries the expected block, and `test_87` checks that each one a no-block entry
names really does.

The successful `calc_state` record is re-asserted too — C14 the stored digest,
C15 the emitted version, C16 the emitted applied timeline, C17 SUCCESS, C18
blank, C19 CURRENT. The two timestamps are **not** required to equal the first
calculation's: a recalculation may refresh them and the contract does not say
otherwise.

`test_88` requires the scenario to call the shared checker rather than a reduced
copy, and refuses any direct `Get-CalcTableRows` in that block.

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
| `P5-FX` | The locked FX seed captured from the untouched Stage-B workbook |
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
| `P5-DC` | Predicates the workbook cannot reach, through `modCalcCheck` directly |
| `P5-D8` | The diagnostic module **removed**, inventory back to 15 |
| `P5-AN` | Every analytical fixture, every emitted expected value |
| `P5-RF` | The nine refusal plan cases, each with its own detail discriminator |
| `P5-PQ` | The plan section 18 prerequisite matrix (25 workbook-reachable predicates) |
| `P5-PN` | The referenced-only complement: 3 assumptions that neither block nor affect |
| `P5-AR` | Driver-audit A/B/C/D reconstruction over a multi-driver fixture |
| `P5-ID` | Identities I1, I2, I3a–c, I4a–c, I5 — production `Reconcile` is the authority |
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
`INVALID`, and it must **preserve the prior successful snapshot**: C13:C16,
C23:C32 and all five tables exactly as they were, while C17:C20 changes. *No
partial result* means no partial NEW snapshot survives — never that the previous
successful one is erased. See
**[correction 2](#2--the-refusal-proof-asserted-the-opposite-of-the-rule)**.

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
Description, a **real multi-row `ListObject` sort** on plan case 30 with the
permanent-ID order captured before and after and required to have changed,
Selected Confidence Level, and an **unreferenced** FX assumption. Each probe
starts from a baseline and restores its change before the next begins.

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
* that each of the 25 prerequisite mutations produces the refusal its tokens
  describe, that the direct `modCalcCheck` call refuses and its control
  accepts, and that the three no-block mutations produce no refusal at all
* that `SyncProfileRows` materialises the fixture's Config-master profiles into
  `tblInflation` with the year band Apply generated, and that the keyed rate
  writer finds each row
* that row 1 of `tblFXRates` in the built workbook really is the reporting seed,
  and that restoring it returns each scenario to a clean baseline
* that an unreferenced assumption leaves every published analytical value
  identical to the base case's
* that the audit columns really do reconstruct the headline totals on real Excel
  arithmetic
* that `FinishOperation` restores a genuinely non-default caller state, including
  `DisplayAlerts` and `StatusBar`
* that `Set-TableCell -Value $null` really leaves an Excel cell BLANK, so cases
  14 and 23 present the model with the blank they describe
* that a real `ListObject.Sort` over three Cost Lines moves the rows on target
* that `PCCM_CurrentInputFingerprint()` on plan case 1 lands on the emitted
  reference digest through the whole production path
* that `PCCM_Calculate` produces the emitted expected values on real Excel
  arithmetic and a real locale
* that both injected failures roll back exactly as described
* that the six status rows come out as the matrix says
* that no owned Excel process leaks and the instance exits naturally

Gate A established what the source says. Gate B has not been run.
