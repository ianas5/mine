# Phase 5 — Gate B — Step B1: the Windows harness extension

**Status: correction round 5 — harness source, ready for independent review.
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
without showing no-effect. Correction round 4 (`56b90d9`) closed both and was
rejected with **three more**, recorded under
**[Correction round 5](#correction-round-5)**: a table reader that destroyed the
`Value2` types the analytical comparator depends on, a snapshot whose
"exactness" was a string serialisation, and an FX seed capture/restore that
could repair a defective built seed. Every one was a defect in what the
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

## Correction round 5

Independent review reproduced 164/164 and 351/351, accepted both round-4 fixes,
and rejected the harness on three more. All three are the same family: the
harness was comparing **display text** where it claimed to compare **values**.

### 1 — the table reader destroyed the `Value2` types

`Get-CalcTableRows` delegated to the accepted Phase-4 `Get-TableBody`, which
does

```powershell
$v = $cell.Value2
if ($null -eq $v) { $line += '' } else { $line += [string]$v }
```

That is right for the structural comparisons it was written for. It is wrong for
Phase-5 analytical evidence, because `Test-CalcValue` is deliberately
type-sensitive: a numeric expectation requires a numeric actual, and a blank must
never equal a numeric zero. A correct cell holding `Value2 = 1.05` arrived as the
String `"1.05"` and the comparison returned **False before it ever compared a
number**. The first successful Gate-B analytical scenario would have failed with
production behaving perfectly — a deterministic harness defect.

`Get-Phase5TypedTableBody` reads the body itself and preserves what Excel
published: blank → `$null`, text → String, numeric → the numeric scalar, Boolean
→ Boolean. Nothing formats; formatting belongs to failure diagnostics, and a
reader that formats has already decided the comparison. The accepted row-emission
idiom survives — one non-enumerated `object[]` per physical row through
`Write-RowObject`, an empty body emitting nothing — and the row is **allocated at
the column count and assigned by index**, because `$line += $null` appends
nothing in PowerShell and a blank cell would silently vanish from the row
(`test_nc_78`). Full COM ownership and release discipline, unchanged.

**The Phase-4 helper is not modified**, and `test_89` asserts its current shape
so the reasoning above cannot rot. `Get-CalcTableRows` no longer calls it, so all
five `_Calc` tables now consume the types Excel actually published. The I5
profiling-weight comparison in `P5-ID` had the same defect for the same reason
and moves to the typed reader too.

**The comparator was not widened.** `test_95` refuses any change that would let a
numeric String satisfy a numeric expectation — that would hide a workbook which
published a number as text, which is precisely what Gate B is for.

### 2 — snapshot "exactness" was a string serialisation

The rollback, refusal and status-row proofs require exact preservation of
C13:C16, C23:C32 and the five analytical ListObjects, blanks included.
`Get-Phase5Snapshot` was `Format-CalcValue`-ing every cell and joining each row
into one String. Two stages destroyed the evidence:

| Stage | What it did |
| --- | --- |
| `Get-TableBody` | every non-null cell → String, every null → `""` |
| `Format-CalcValue` + join | the row → one display String |

So a numeric `1` and the String `"1"` produced identical evidence, and a real
`Empty` and an empty String both collapsed to `<blank>`. And the scalar
comparisons used `Test-CalcValue`, which treats `$null` and `""` as the same
absence **by design** — right for "is this the value the oracle expected?",
too weak for "was this restored exactly as it was?".

Snapshot identity now has its own rule, `Test-Phase5ExactValue`:

| Expected | Requires |
| --- | --- |
| `$null` | `$null` — an empty String is a value the user entered, not an absence |
| String | String, exact case-sensitive equality |
| numeric | non-String, non-Boolean numeric, exact equality |
| Boolean | Boolean, exact value |

No tolerance, no display-text conversion. `Get-Phase5Snapshot` retains rows as
typed cell arrays; `Add-SnapshotUnchangedChecks` compares row count, then column
count, then every cell through the strict comparator, and reports the first
difference with its type. `Format-Phase5Typed` exists only to describe a failure.

`test_93` runs the four planted restorations as a decision table — numeric `1`
restored as `"1"`, a blank restored as `""`, identical display text with the
wrong type, a `calc_totals` blank replaced by `""` — and also shows the shipped
chain accepting the pair the typed snapshot rejects.

### 3 — the FX seed capture could repair a defective built seed

Round 4 fixed cross-scenario contamination and said the seed was captured "from
the real Stage-B workbook, not repaired from literals". The value flow still
performed a type-changing repair: `Save-Phase5LockedFxSeed` read through
`Get-TableBody`, so a correct numeric `1` was captured as the String `"1"` — and
an **incorrect** built seed already holding the text `"1"` was captured
identically. `Reset-Phase5FxTable` then wrote `([double]$Seed.Rate)`, converting
the capture into a number.

A workbook that had built the reporting FX rate as **text** would therefore have
been silently corrected before any analytical scenario ran, in violation of the
round-4 rule: *if the built seed itself is wrong, the analytical scenarios must
fail; the harness must not repair it into agreement.*

The capture now uses the typed reader, and the restoration writes the captured
value **as itself** through `Set-Phase5TypedCell`, which assigns `Value2`
directly and chooses nothing — a `$null` still becomes a genuine blank via
`ClearContents`. The read-back uses the typed reader and the strict comparator,
because a read-back through the analytical comparator would accept the very type
change the restoration must not make. `P5-FX` reports the captured value **and
its type** as a note and asserts neither against a literal: whether the built
rate is numeric is the production build's claim, and the analytical scenarios are
what test it. The accepted Phase-4 `Set-TableCell` is untouched and still chooses
between `[string]` and `[double]`, which is right for fixture authoring.

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
A1  first Application.Run of the run  ->  the automation surface ANSWERS
P5-P4  the Phase-4 matrix is intact
P5-CMP  the WHOLE production project compiles (VBE Compile VBAProject)
P5-D0  the diagnostic module is imported into the DISPOSABLE workbook
P5-D1 .. P5-D7  the locked vectors
P5-D8  the diagnostic module is REMOVED; inventory re-asserted at 15
P5-AN onward  the analytical acceptance work, with no test module installed
```

**A1 is the first `Application.Run` boundary, not a compilation boundary.** It
proves the automation surface answers; Runtime Run 7 passed it and then met a
VBE compile error, so the whole-project claim belongs to `P5-CMP` alone. What
`test_22` preserves is the ORDERING that matters: the diagnostic module is
imported only after the production project has been proved to compile, and it
stays production-only so a test module can never mask or contaminate that
proof. The module is not in the manifest, not in the structure
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
  that its rate was built as a number rather than as text, and that restoring it
  returns each scenario to a clean baseline
* that `Range.Value2` really does hand back a numeric scalar for a numeric cell,
  `$null` for a blank one and a String for text, so the typed reader observes
  what this source assumes it will
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

---

## Runtime Run 1: the Phase-4 prerequisite sequencing correction

Gate B has now been run once on Windows, at harness commit `35640ec`. That run
did not reach a Phase-5 result. It stopped at the prerequisite gate:

```
[FAIL] P5-P4  Phase-4 prerequisite: 35/35 PASS, 0 FAIL, 0 SKIP
    FAIL all 35 Phase-4 scenarios reported a result -- missing: Y, Z
    ok   the Phase-4 matrix has 0 FAIL
    ok   the Phase-4 matrix has 0 SKIP
    FAIL the Phase-4 matrix is 35/35 PASS -- passed 33 of 35

[FAIL] P5-ALL  Phase-5 Gate-B scenarios
    not attempted: the Phase-4 structural matrix is not intact, so no Phase-5
    result would mean anything

[PASS] Z  Excel closed naturally after the functional run
[PASS] Y  Transient COM releases

  36 passed, 2 failed, 0 skipped
  Phase-4 structural matrix: 35 of 35
```

Nothing in the workbook was wrong. The prerequisite demanded two results that
cannot exist at the point it runs.

### Why Y and Z cannot precede Phase 5

`Y` and `Z` are the only two entries in the 35-case matrix that are not Phase-4
*behaviour* cases:

* **Z** asserts the owned Excel process exited naturally. It is recorded after
  `Workbook.Close`, `Application.Quit`, the COM release ledger and
  `Wait-ExcelExit`.
* **Y** asserts every transient COM object released cleanly across the *whole*
  run, Phase-5's own transients included. It is recorded last of all, from
  `Get-TransientFailures`.

Phase 5 runs inside the live automation session — `Invoke-Phase5GateBScenarios`
is called with the same `$excel` and `$wb`, before `PCCM_AutomationEnd`. So the
gate asked for a post-session result from inside the session. It was
unsatisfiable by construction, and no workbook could have passed it.

Evaluating `Y` early would not have helped either: it would attest to Phase-4's
transients only, which is weaker evidence, not stronger.

### The correction: two gates, one matrix

The threshold was never the problem, so the threshold did not move. The matrix
is still 35 cases and 35/35 PASS with 0 FAIL and 0 SKIP is still required. What
changed is *where* each demand is made.

| | `P5-P4` (entry) | `P5-FIN` (final) |
|---|---|---|
| runs | before any Phase-5 scenario, session live | after `Y` and `Z`, session gone |
| judges | the 33 prerequisite cases | all 35 matrix cases |
| FAIL tolerated | none | none |
| SKIP tolerated | none | none |
| pass count | 33/33 | 35/35 |
| by name | — | each deferred case, exactly one record, `PASS` |
| on failure | `P5-ALL` FAIL, `return` | FAIL, and the run exits 1 |

`$script:Phase4FinalizationScenarioIds = @('Y', 'Z')` names the deferred cases.
`Get-Phase4PrerequisiteScenarioIds` **derives** the entry set as the matrix minus
that list — there is no second hand-maintained roster to drift, a case added to
the matrix becomes a prerequisite case automatically, and the only way to defer
one is to name it while it stays inside the 35.

`P5-P4` proves the partition rather than asserting it in prose: the two sets must
be disjoint, must sum to 35, and every deferred name must be a real matrix
member. It also checks that no deferred case has *already* been recorded — if one
had, it was never a post-session case and has no business being excluded. That
check is the direct regression assertion for the Run-1 topology.

`P5-FIN` is what makes the deferral safe. A bare 35/35 count would be satisfiable
by a matrix that lost `Z` and counted something else twice, so each deferred case
is checked by name for exactly one record with status `PASS`. Because the
driver's exit code is driven by the FAIL count, a `P5-FIN` FAIL fails the whole
run — including on a Phase-4 SKIP, which the summary alone would have printed
before exiting 0.

### What this correction is not

It does not lower 35 to 33. It does not exempt any case from the FAIL/SKIP rule.
It does not remove `Y` or `Z` coverage — it makes both *mandatory at a gate that
can see them*, which the accepted source never had. It touches no production VBA,
no calculation behaviour, no contract, no fingerprint and no oracle: both gates
read the in-memory result ledger and nothing else, and the tests pin that neither
reaches for `$Workbook`, `$Excel`, `.Run(`, `VBProject` or any `PCCM_` endpoint.

Runtime Run 1 stands as historical evidence of a harness sequencing defect. It
was not a production failure, and it has not been rerun.

---

## Runtime Run 2: four harness roots and one production finding

Run 2 (harness commit `cc70c37`) reached the Phase-5 scenarios and finalised
correctly — 35/35 Phase-4, `P5-FIN` PASS, clean shutdown ledger, natural Excel
exit. It then failed 24 scenarios. They are not 24 defects. Four harness roots
account for 22 of them; the remaining two are one production finding.

| Root | Scenarios | Run-2 evidence |
|---|---|---|
| R1 inventory semantics | P5-M, P5-D8 | `present 30 of 15`, `extra: ThisWorkbook, shDashboard, …` |
| R2 commentary read as code | P5-EV | `modAppState: Worksheet_Change; modAppState: NPV` |
| R3 typed-parameter shadowing | 7 | `PropertyNotFoundException: The property 'rows' cannot be found` |
| R4 shape count as button count | P5-M | `found 6` |
| **P** production canonical number | P5-D1, P5-D2 | `1.7976931348623200E+308` vs `…3157E+308` |

### R1 — a VBProject is not a module list

`P5-M` and `P5-D8` enumerated `VBProject.VBComponents` and compared every
component *name* against the manifest's 15-entry `vba.modules`. The 30 is
arithmetic, not a defect:

```
15 standard modules + 14 sheet documents + 1 ThisWorkbook = 30 components
```

`vba.modules` describes the production **standard modules**. It never described
document components and never could — Excel creates those when a sheet exists;
the bootstrap does not import them.

The inventory is now partitioned by VBIDE component type
(`Get-Phase5VbComponentInventory` → `Add-Phase5ModuleInventoryChecks`, shared by
both scenarios so they cannot drift). Nothing was weakened to "at least 15":
the standard-module set must still equal the manifest set exactly, by name, in
both directions. Document components are *counted* (`sheets + 1`) so a stray one
cannot hide there either, and class modules, UserForms and ActiveX designers are
excluded outright. `P5-D8` proves the diagnostic module's absence against the
standard-module partition — the partition it would have to reappear in.

### R2 — a comment about `Worksheet_Change` is not a `Worksheet_Change`

`P5-EV` searched each `CodeModule`'s raw text for every
`forbidden_constructs` entry. Both Run-2 hits are prose in accepted production
source:

```
modAppState.bas:7    ' … No cost, risk, escalation, FX, NPV, EMV,
modAppState.bas:78   ' … no input Worksheet_Change handler, and this
```

A comment explaining that there is no handler was read as a handler. The comment
stays — it is the reason the guarantee exists.

`Remove-VbaCommentary` strips comments (tracking string literals, so an
apostrophe inside `"it's"` does not truncate the statement, and handling
`Rem`-form) before the same manifest-driven scan runs. A second check,
`Test-VbaProcedureDeclared`, additionally reports a real
`Sub Worksheet_Change(` / `Sub Workbook_SheetChange(` **declaration** as the
declaration it is. No construct left the manifest list and no blanket text
substitution is used, so a real declaration on the line after a comment is still
caught.

The Python side already drew this line — `builder/pccm_builder/vba_source.py`
strips comments and string literals before any structural scan, for exactly this
reason. This is the same rule applied to the code Excel actually holds.

### R3 — one shadowed parameter, seven scenarios

```powershell
function Get-CalcScalar {
    param($Workbook, $Inspection, [string]$Block, [string]$FieldKey)
    $block = $Inspection.calc.scalar_blocks.$Block     # <-- $block IS $Block
    $row = [int]$block.rows.$FieldKey                  # <-- String has no .rows
```

PowerShell variable names are **case-insensitive**, and a typed parameter keeps
its constraint for the life of the variable. Assigning the block PSCustomObject
to `[string]$Block` converted it to `"@{value_column=C; rows=…}"`; the next line
asked a String for `.rows` and StrictMode did the rest — on **every** call.
P5-S2, P5-ST, P5-S3, P5-S4, P5-S5, P5-KP and P5-RC are one defect.

The inspection projection was never at fault: it carries `rows` for both
`calc_state` and `calc_totals`, as `phase5_gate_b_inspection.json` shows. The
schema is unchanged.

The local is renamed `$blockSpec`, and a source test scans **every** typed
parameter in all three PowerShell files for the same shadowing — the class is
closed, not just the instance.

### R4 — a Shape is not a command button

`P5-M` counted every `Shape` on every manifest sheet and required five, while
all five declared buttons were present with the right `OnAction` and no shape
called `PCCM_Calculate`. It reported `found 6` and did not say what the sixth
was, so the run could not be diagnosed from its own evidence.

A command button is a shape **bound to a macro**. Every shape is still
enumerated and still judged; the five-button rule now applies to the bound
shapes, every bound shape must be one of the five declared buttons, and no
undeclared shape may carry a `PCCM_` macro at all. The inventory is reported by
`sheet!name -> macro`. This is strictly stronger than the count it replaces —
an unbound decoration passes, a sixth *bound* shape does not.

### Diagnostics — Run 2 could not be located from its own output

Eleven scenarios reported one indistinguishable sentence:

```
System.InvalidCastException: Unable to cast object of type 'System.Double' to type 'System.String'.
```

The accepted Phase-4 `Format-Err` returns exception type and message and nothing
else. `Format-Phase5Err` adds the inner-exception chain, script/line/column, the
offending source line, and the `ScriptStackTrace` frame by frame — which is what
distinguishes `Set-Phase5Fixture → Reset-Phase5FxTable` from
`Set-Phase5Fixture → Write-Phase5Driver` when one shared path serves eleven
scenarios. It reads plain data only: no COM object reaches the ledger.
`Format-Err` itself is accepted Phase-4 source and is untouched.

### What Run 2 proved *works*

`P5-FX` passed and recorded, on real Windows PowerShell 5.1 and real Excel:

```
P5-FX: locked FX seed captured as String'SAR' / Double:1
```

That is `New-Object 'object[]' $colCount`, index assignment, `Write-RowObject`
and `Format-Phase5Typed` all behaving exactly as designed — a String stayed a
String and a numeric stayed a Double across the COM boundary. The round-5 typed
reader is therefore **not** the `InvalidCastException` boundary, and it was not
changed. Stringifying `Value2` would destroy the evidence architecture for
nothing.

The `Double → String` cast is **not yet root-caused**, and is deliberately not
guessed at. See the Run-2 review return for the narrowed candidate set.

---

## Review round 2A: three corrections to the Run-2 package

### Correction 1 — the component inventory emitted one nested array

`Get-Phase5VbComponentInventory` accumulated into `$out` and ended with
`return ,$out`. The unary comma exists to stop PowerShell unrolling a
collection. That is right for a function returning one **row** whose cells must
stay together — which is why `Write-RowObject` uses `-NoEnumerate` — and wrong
for a function producing a **sequence of records**: the caller's `@(...)` sees a
single nested array, and every downstream

```powershell
$Components | Where-Object { [int]$_.Type -eq ... }
```

filters one array-shaped object that has no `.Type` at all. No partition would
ever match, silently, and no textual source test can see it.

The function now emits one `PSCustomObject` per component:

```
zero components -> nothing emitted
one component   -> one PSCustomObject
N components    -> N PSCustomObjects
```

and the caller's `@(...)` is the authority that turns 0/1/N into an `Object[]`.
A `PSCustomObject` is not a collection, so there is nothing for the comma to
protect. **`Get-Phase5TypedTableBody` is untouched** — it emits one `object[]`
per row and must keep `-NoEnumerate`, or row boundaries would be lost.

### Correction 2 — the button proof now binds sheet, shape and macro together

The previous rule proved three independent global sets: five bound shapes, every
bound name declared, every declared entry point present somewhere. This passes
on a real defect:

```
btnPCCMAddCostLine    -> PCCM_DeleteCostLine
btnPCCMDeleteCostLine -> PCCM_AddCostLine
```

Every name exists, every macro exists, five shapes are bound, nothing calls
`PCCM_Calculate` — and two buttons do the opposite of what they say. Worse, the
old per-button line would have printed `ok the button btnPCCMAddCostLine calls
PCCM_AddCostLine` about a button that calls `PCCM_DeleteCostLine`. That check is
removed, not merely supplemented.

The manifest already carries the whole identity (`sheet`, `shape_name`,
`entry_point`), so the proof operates on `(Sheet, ShapeName, OnAction)` triples.
For each declared button: exactly one shape of that name exists on that sheet;
**that** shape's `OnAction` equals the declared entry point (case-sensitive); and
no second copy of the name exists on any other sheet. Then: no undeclared
macro-bound shape, no undeclared shape reaching `PCCM_*` at all, no shape calling
`PCCM_Calculate`, and the bound triple set equals the declared triple set —
which closes the count from both ends without ever counting raw shapes. The raw
`Shape.Count == 5` rule is **not** restored; a sixth *unbound* decorative shape
still passes.

### Correction 3 — string literals are data, not code

`Remove-VbaCommentary` removed comments but left string payloads, so the runtime
scanner and the static one had different semantics for the same question. The
Python authority has always done both — `VbaModule.code` is
`strip_strings(strip_comments(raw))`, and `contains_construct()` scans that.

```vba
MsgBox "NPV is not available"          ' prose, in a message
Err.Raise 5, , "Worksheet_Change"      ' prose, in an error string
```

`Remove-VbaStringLiterals` applies the same regex the Python side uses —
`"(?:[^"]|"")*"` replaced by an **empty** literal, so the surrounding statement
keeps its shape and a forbidden token after a literal is still found.
`Get-VbaExecutableCode` composes the two in the Python order: comments first
(that pass is the one that understands literals, so it must run while they are
intact), then literals. `Remove-VbaCommentary` also gained the doubled-quote
escape it was missing, so `"he said ""don't"""` no longer closes early.

See *Review round 2B* below for the `Rem` half of the same rule, which this
round got wrong.

The manifest `forbidden_constructs` list is unchanged, and `test_129` runs the
corrected rule over every frozen production module to prove the stricter scanner
does not flag accepted source.

### Correction 4 — what P5-D5 actually proves

An earlier statement of mine — that P5-D5 passed because the reference model's
values are representable within the formatter's effective significant digits —
**was wrong**, and the correction matters.

P5-D5 does not reconstruct the reference stream through
`CalcFpCanonicalNumber` at all. It reads the **pre-emitted** stream:

```powershell
$stream = [string]$reference.stream
$length = [string]$Excel.Run('GBD_StreamLength', $stream)
$digest = [string]$Excel.Run('GBD_DigestStream', $stream)
```

So P5-D5 proves that an already-canonical stream survives VBA length and digest
processing. It says nothing about whether production VBA can **generate** that
stream from its Double inputs. The 15-significant-digit ceiling in
`Format$`, which P5-D1 and P5-D2 expose, is therefore not excused by P5-D5
passing — and P5-D1/P5-D2 become **more** important, not less: they are the only
scenarios that exercise the generation step directly.

P5-D1 and P5-D2 remain a PRODUCTION finding and remain unrepaired.

---

## Review round 2B: `Rem` is a statement, not a line prefix

Round 2A recognised `Rem` commentary only as the first token of a physical line:

```powershell
if ($text -match '^\s*Rem(\s|$)') { $text = '' }
```

VBA permits `Rem` wherever a **statement** may begin, which includes after a
colon separator:

```vba
x = 1: Rem Worksheet_Change is deliberately absent
```

That line survived the round-2A stripper intact, so P5-EV could report
`Worksheet_Change` as executable code from a comment — the same class of false
positive the round-2 correction existed to remove, in a form it did not cover.

The shape was the real problem. The old rule rebuilt the line and *then* matched
a line-anchored regex over it, so it could not see whether a colon was inside a
string literal or whether one had occurred at all. The decision now happens at a
character position, inside the single pass that already tracks literals:

```powershell
if ((-not $inString) -and $atStatementStart -and
    ((($i + 3) -le $line.Length)) -and
    ($line.Substring($i, 3) -eq 'Rem') -and
    (((($i + 3) -eq $line.Length)) -or [char]::IsWhiteSpace($line[$i + 3]))) {
    break
}
```

Three conditions, all required, and one consequence:

* **outside a literal** — `x = "Rem Worksheet_Change"` is data;
* **at a statement boundary** — the line start, or after a colon that is itself
  outside a literal, so the colon in `"text : Rem NPV"` does not open one;
* **the complete keyword** — followed by whitespace or end of line, so
  `Remember` and `RemoteValue` stay identifiers;
* once it begins, the rest of the physical line is commentary.

The boundary is maintained in the same loop: whitespace leaves it alone (so
`x = 1 :   Rem …` still works), a literal closes it, and a colon outside a
literal reopens it. No broad regex was introduced; the line-anchored one is
gone rather than supplemented.

### Decision table

| Source | Verdict |
|---|---|
| `Rem Worksheet_Change is absent` | allowed |
| `x = 1: Rem Worksheet_Change is absent` | allowed |
| `x = 1 :   Rem NPV is deliberately absent` | allowed |
| `x = "Rem Worksheet_Change"` | allowed (string data) |
| `x = "text : Rem NPV" : Randomize` | **`Randomize` FAILS** |
| `Remember = 1` | not a comment |
| `RemoteValue = "NPV"` | not a comment |
| `x = 1: Randomize` | **FAIL** |
| `x = 1: Rem Randomize is deliberately absent` | allowed |
| `REM …` / `rem …` | allowed (VBA is case-insensitive) |
| `MyLabel: Rem NPV` | allowed |
| `Dim Remainder As Long: Randomize` | **FAIL** |

Everything from round 2A still holds: apostrophe comments, apostrophes inside
literals, doubled quotes, literal removal, and real forbidden tokens after a
literal or before a trailing comment.

`test_132` additionally proves that on every frozen production module the
runtime stripper and the Python authority produce **identical** executable code,
so the stricter rule did not change what the accepted source says.

---

## The canonical Double encoder: a production correction

Gate B Runtime Run 2 exposed the first defect in accepted **production** VBA.
This section records what it was, why the contract did not move, and what the
replacement actually does.

### 1. The contract stays at 17 significant digits

The reference encoder is `f"{number:.16E}"` — one digit before the point,
sixteen after — with negative zero normalised. That is not a stylistic choice:

* **17 digits is the shortest width that round-trips every binary64.** At 15,
  distinct Doubles collide. A concrete pair, from the test corpus:

  ```
  x = -1.996150245444706e-194
  y = -1.9961502454447058e-194     (the next representable value above x)

  15 digits   -1.99615024544471E-194  ==  -1.99615024544471E-194
  17 digits   -1.9961502454447061E-194 != -1.9961502454447058E-194
  ```

  Over 4 000 deterministic probes, more than 100 neighbour pairs collapse at 15
  digits. A canonical numeric FIELD that maps two different Doubles to the same
  text maps two different models to the same fingerprint — which is the one
  thing the fingerprint exists to prevent.

* Reducing the contract to 15 to suit `Format$` would therefore not be a
  relaxation, it would be a defect. The contract is confirmed, not assumed;
  `test_02` and `test_nc_02` in `tests/test_phase5_canonical_number.py` prove it.

### 2. Root cause

```vba
Private Const FP_NUMBER_FORMAT As String = "0.0000000000000000E+00"
text = Format$(number, FP_NUMBER_FORMAT)
```

Sixteen fractional **placeholders**, but VBA's numeric-to-text conversion
carries about **15 significant decimal digits**. The placeholders beyond that
are filled with zeros, not with recovered digits. Every Run-2 failure has that
shape:

| value | produced | contracted |
|---|---|---|
| `0.1` | `1.0000000000000000E-01` | `1.0000000000000001E-01` |
| `1e-20` | `1.0000000000000000E-20` | `9.9999999999999995E-21` |
| `0.1 + 0.2` | `3.0000000000000000E-01` | `3.0000000000000004E-01` |
| `MAX_DOUBLE` | `1.7976931348623**200**E+308` | `1.7976931348623**157**E+308` |
| min subnormal | `4.9406564584124**700**E-324` | `4.9406564584124**654**E-324` |

Three distinctions matter here:

* **stored precision** — a binary64 holds 53 bits, about 15.95 decimal digits;
* **display precision** — what the host's conversion will emit, about 15;
* **round-trip representation** — 17 digits, the width needed to recover the
  exact bit pattern.

`Format$` supplies the second. The contract needs the third. **Adding
placeholders cannot help**, because the digits were never produced —
`test_nc_01` models the defect and reproduces all five observations exactly.

### 3. The replacement: generated, not formatted

A binary64 is exactly `M × 2^E` with `M` an integer, so its decimal expansion is
**finite** and computable with integer arithmetic alone.

```
decompose   value  ->  M in [2^52, 2^53), E in [-1126, 971]
                       only *2 and /2, so every step is exact;
                       2^-1074 normalises to M = 2^52, E = -1126
E >= 0      digits of  M * 2^E,          point at the end
E <  0      digits of  M * 5^(-E),       point -E places from the right
round       once, to 17 significant digits, half to EVEN
format      [-]d.dddddddddddddddd E[+-]dd
```

The big integer is held as base-10⁷ limbs in Doubles. A limb-by-factor product
stays under 10¹⁴, well inside the exact-integer ceiling of 2⁵³, so **every
intermediate is exact**. Powers are consumed in chunks that fit one limb — 2²³
and 5¹⁰ — so the widest case (M × 5¹¹²⁶, 804 digits, 115 limbs) costs about 113
passes rather than 1 126.

Round-half-even is required, not decorative: a binary64's expansion terminates,
so the 18th significant digit really can be a 5 with nothing after it. Thirteen
such exact ties occur in 20 000 random Doubles.

Nothing calls `Format`, `CStr` or `Str`; digits are selected from a literal
table. No Windows API, no `Declare`, no `LongPtr`, no `Application` reference —
only `Double`, `Long`, `String` and `Boolean` appear, so 32-bit and 64-bit
Office are the same code. This module still contains no `Mod` and no `\`, even
where both operands are Longs and either would have been correct.

### 4. Separator invariance became structural

The separator stays the locked public interface and is still validated as
exactly one UTF-16 code unit. What changed is that it can no longer **reach** the
output: the encoder emits the marker itself, so there is nothing to normalise.
The locked dual injection — `.` and `,` on one host, byte-identical output — is
now true by construction rather than by repair. `CalcFpMarkerIndex` survives as
a **post-condition on this module's own output**.

### 5. Proof, and the limit of it

The shipped algorithm is transcribed routine-for-routine into
`tests/vba_canonical_port.py` and held against the Python oracle:

| corpus | result |
|---|---|
| every power of ten from 1e-323 to 1e308, plus both neighbours of each | 0 mismatches |
| 400 000 deterministic bit patterns | 0 mismatches |
| the emitted parity corpus (2 432 vectors) | 0 mismatches |
| the ten locked P5-D1 vectors and eleven P5-D2 separator vectors | 0 mismatches |

**What this does not prove** is that VBA executes the transcription faithfully.
That is Gate B's, on Windows, against the emitted corpus. The port is never an
authority: expectations come from `calc_fingerprint.py`, and `test_14` pins each
ported routine and constant against the module it mirrors so the two cannot
drift.

### 6. Fingerprint compatibility: nothing moved

The emitted corpus before and after, leaf by leaf:

```
leaves before : 3999
leaves after  : 11386
added         : 7387   all under fingerprint.canonical_parity
removed       : 0
CHANGED       : 0
```

The reference digest is still `50B6EB0E26857EA7` over 366 code units; the
reference stream, collision probes, UTF-16 vectors, the ten numeric encodings,
the separator vectors, every plan case and every regression vector are
byte-identical.

**No golden value changed, and none needed to.** The oracle was always correct —
the emitted corpus and the reference digest have only ever been produced by
Python. What was wrong was VBA's ability to *reproduce* them. Text and integer
field encoding, UTF-16 length, field framing, section and record grammar, and
the digest recurrence are all untouched.

**Compatibility with previously calculated workbook fingerprints:** a stored
fingerprint produced by the defective encoder over a model containing a number
that needs more than 15 significant digits would not match one produced now.
That is the correction working as intended — the old value was wrong — and it
surfaces as `STALE`, which is the accepted signal for "the stored fingerprint no
longer describes these inputs". No accepted workbook has ever been calculated:
Gate B has not passed, so no such fingerprint exists in the field.

### 7. P5-DP

The corrected encoder is held at Gate B against the whole emitted parity corpus,
not against ten examples — `Format$` was wrong on six of the ten and right on
four, so ten was never a width that could accept a replacement. Each probe is
rebuilt from its **IEEE-754 bit pattern**, never parsed from a decimal literal,
and the expected text is read from the corpus: the harness computes no canonical
string of its own. Nine neighbour triples double as a collision proof — three
distinct Doubles must give three distinct canonical strings.

---

## Build-artifact isolation

The first full-suite run of the canonical-encoder correction on Windows reported
`1609 passed, 2 failed`, both with `KeyError: 'canonical_parity'`. Rebuilding
Stage A by hand made the same suite report `1611 passed`. The production encoder
was never implicated.

The cause was mine, in the new suite's corpus helper:

```python
if not BUILD.is_file():
    subprocess.run([... build_stage_a.py ...])
return json.loads(BUILD.read_text(...))["fingerprint"]["canonical_parity"]
```

It rebuilt only when the file was **absent**, which treats "a file exists" as
"an artifact generated from this source". A checkout carrying `build/` from the
previous commit satisfies the first and not the second. **A suite that requires
the operator to know it must rebuild first is not a suite that can be trusted
after `git pull`.**

The corpus is now emitted by the real builder — `emit_calc_artifacts`, the same
entry point `build_stage_a.py` calls — into a test-owned temporary directory,
once per session. Nothing consults `pccm/build` at all, no subprocess is
spawned, and there is no second implementation of the corpus. This is the
pattern `tests/test_phase5_gate_b_harness_source.py::_emitted` already used,
whose docstring already said *"Never read from `build/`"*.

One other site had the same defect and is closed with it: `test_113` read
`build/phase5_gate_b_inspection.json` when it happened to exist and skipped
silently when it did not — trusting a stale file, and proving nothing when the
build is broken. It now reads the freshly emitted projection.

`test_phase4_oracle.py::test_43` also names a path under `build/`, and was
examined and **left alone**: it asserts the artifact *matches* the oracle, so a
stale file makes it fail loudly rather than pass falsely. Different failure
direction, accepted Phase-4 source, out of scope.

The regression plants a previous-schema corpus at the real repository path,
clears the session cache so the helper must fetch again, and runs the two tests
that failed on Windows. Corrupt payloads — unparseable, empty, `{}`, and one
carrying a tampered expectation — are proved unable to become an oracle. The
whole condition is reproduced end to end: with a genuine `build/` from the
previous commit in place, the accepted commit's suite fails on exactly tests 11
and 12 with `KeyError`, and the corrected suite passes.

---

## Runtime Run 3: two production compile blockers

Run 3 never reached a scenario. The VBE stopped during the production compile,
twice, and everything after that is session-loss cascade. The verification gap
it exposed matters as much as the two defects: **1616 Python tests and 351
Stage-A checks passed on a project that would not compile.**

### Blocker 1 — `Dim scale As Long`

`modCalcFingerprint.CalcFpBuildCanonical`, Syntax error, `scale` highlighted.

`scale` is **not** simply a forbidden token: seven `scale` locals in
`modCalcFactors` and `modCalcAnalytical` compiled in Runtime Run 2, which
reached P5-M with every module present and every API procedure callable. What
distinguishes the failing site is position — in all seven that compiled, `scale`
is a *later* item in its `Dim` list; only this one stood as the token
immediately after `Dim`, where the parser is looking for a fresh identifier and
`Scale` is a Visual Basic statement keyword.

Renamed to `decimalScale`. `base`, in the two new power helpers, was renamed to
`powerBase` on the same reasoning: `Base` is the keyword in `Option Base`, Run 3
stopped before reaching those lines, and an unproven identifier in the code path
that already failed once is not worth defending.

Both are identifier renames and are **provably invisible**: mapping the new
names back reproduces the accepted executable text exactly, and `test_21` keeps
it that way.

### Blocker 2 — `MAX_DOUBLE` overflow

`modCalcFactors`, Overflow, on a literal the VBE displayed as
`1.79769313486232E+308` — which is **not what the source said**. The source
carried the 17-digit form:

```vba
Public Const MAX_DOUBLE As Double = 1.7976931348623157E+308
```

The displayed value is the fifteen-significant-digit rounding of it, and that
rounding is **above** the maximum finite Double. VBA converts a numeric literal
at about fifteen digits and only then range-checks the result, so the value is
out of range before it exists. It is the same fifteen-digit ceiling Run 2 proved
on the formatting side, arriving from the other direction — and it is why the
old test passed: it compared the literal as an exact `Decimal`, which is correct
for a correctly-rounding parser and irrelevant to this one.

**No decimal spelling can be trusted here**, and a rounded-down literal would be
a different number wearing the right name: `IsUsableDouble` compares against this
bound, so a value below the true maximum would refuse the largest representable
Double — and the accepted `MAX_DOUBLE` fingerprint vector requires that value to
be usable and to encode as `1.7976931348623157E+308`. The contract is therefore
**exact DBL_MAX**, not a safe approximation.

So it is built from the identity the module already stated:

```vba
Public Function MAX_DOUBLE() As Double      ' cached
Private Function BuildMaxDouble() As Double
    result = MAX_SIGNIFICAND                ' 2^53 - 1
    For doubling = 1 To MAX_EXPONENT        ' * 2^971
        result = result * 2#
```

Doubling is exact in binary floating point and no intermediate exceeds the final
value, so nothing overflows on the way; the result is bit-for-bit DBL_MAX. It is
a Function because a `Const` initialiser cannot compute — every call site reads
`MAX_DOUBLE` unchanged, and none used it in a constant expression.
`MAX_EXPONENT`, already declared as 971, is reused rather than restated.

### Closing the static gap

Two focused classes, run over declarations parsed from **executable** code, so
prose naming a keyword is never read as a declaration:

* `test_86`/`test_87`/`test_88` — no production declaration may introduce a VBA
  statement keyword or type name. The eight pre-existing sites are grandfathered
  on Run 2's successful compile and on nothing else, checked as a **set** in both
  directions so a new one cannot slip past; a stale entry fails too.
* `test_57`/`test_89` — every Double literal must still be in range after a
  fifteen-significant-digit round trip, and every `Const` literal must fit its
  declared type. The retired `MAX_DOUBLE` literal is the negative control.
* `test_90`/`test_91` — array bounds are constant expressions, no local shadows
  its procedure, every referenced `Const` is declared, and module-level
  declarations precede the first procedure in both changed modules.

These are not a VBA compiler and are not trying to be. They stop the two
deterministic blockers now observed from passing the static gate again.

### Nothing else moved

`phase5_cases.json` is **byte-for-byte identical** to the accepted commit's,
reference digest `50B6EB0E26857EA7` included. The eleven other production
modules are byte-identical to accepted Gate-A `1968fb8`.

---

## Review round 3A: the boundaries and the exemption authority

### Blocker 1 — the construction still trusted a long decimal literal

The round-3 fix built `MAX_DOUBLE` from

```vba
Private Const MAX_SIGNIFICAND As Double = 9007199254740991#
```

which is a **sixteen**-significant-digit decimal spelling — precisely the class
of literal the same round declared unparseable. The reasoning was
self-inconsistent, and the failure mode is worse than the one it replaced: an
overflow stops the compile, whereas a significand parsed one unit low **compiles,
runs, and silently yields the Double immediately below the maximum**. Python
parsing that literal exactly proves nothing about VBA.

A sweep found five such literals, all inside the two authorised modules:

| module | literal | digits |
|---|---|---|
| `modCalcFactors` | `4503599627370496#` (`TWO_52`) | 16 |
| `modCalcFactors` | `9007199254740991#` (`MAX_SIGNIFICAND`) | 16 |
| `modCalcFingerprint` | `4503599627370496#` (`FP_TWO_52`) | 16 |
| `modCalcFingerprint` | `9007199254740992#` (`FP_TWO_53`) | 16 |
| `modCalcFingerprint` | `1000000000000000#` (10^15) | 16 |

All five are gone. What remains is bit widths as small `Long` constants, and
constructions from `1#`, `2#` and `10#`:

```vba
Private Const SIGNIFICAND_BITS As Long = 53
Private Const MANTISSA_BITS As Long = 52

Private Function ExactPowerOfTwo(ByVal bits As Long) As Double
    result = 1#
    For doubling = 1 To bits
        result = result * 2#
```

* `TWO_52` = `ExactPowerOfTwo(52)`
* `MAX_SIGNIFICAND` = `ExactPowerOfTwo(53) - 1#` — exact, because 2⁵³−1 needs
  exactly 53 bits and is representable, so the subtraction does not round
* `MAX_DOUBLE` = `MAX_SIGNIFICAND × 2^971` — 971 exact doublings, no
  intermediate above the final value
* `FP_MANTISSA_BITS` / `FP_SIGNIFICAND_BITS` build the decomposition bounds
  through the encoder's existing `CalcFpIntegerPower`
* 10¹⁵ = `CalcFpIntegerPower(10#, 15)` — every 10^k for k ≤ 15 is exactly
  representable, so each step is exact

**No parser-precision assumption remains anywhere in production VBA**: the
longest surviving Double literal is `10000000#`, eight digits, and a static rule
keeps it that way.

The one-unit-low case is pinned as a negative control: it lands exactly on
`nextafter(DBL_MAX, 0)`, is distinguishable, and is rejected.

### Blocker 2 — module-wide grandfathering was a hole

The exemption key was `(module, identifier)`. `modCalcFactors` already declares
`scale`, so that pair grandfathered **any new `scale` anywhere in the module** —
contradicting the stated rule that a new one is rejected.

The key is now the **site**, counted as a multiset:

```
(module, enclosing scope, declaration kind, identifier, normalised statement) -> count
```

Fifteen sites are recorded, each an occurrence Runtime Run 2 compiled. No line
numbers — those move whenever anything above them does. Both directions are
required: a new site fails, and a stale entry fails too, because a stale
exemption is itself a hole.

Four ways a new occurrence can appear, all rejected: a **new procedure** in a
module that already has that name; a **second reserved name** inside an already
grandfathered procedure; an **identical repeat** of a grandfathered declaration,
caught by the count; and the same identifier in a **new module**. Comments and
string literals remain non-declarations.

### A correction to the round-3 documentation

The earlier note offered a pattern — that the seven pre-existing `scale`
declarations sit later in their `Dim` lists while the failing one stood
immediately after `Dim`. **That is not asserted as a VBA language rule, and it is
not the basis of any exemption.** It is an observation about the failing site.
What the exemptions rest on is narrower and empirical: Runtime Run 2 compiled
these exact occurrences in the target Excel environment, and nothing more is
claimed for them.

---

## Runtime Run 4: R5 located, and one result per scenario

Run 4 is the first run in which the production side is largely proved. The VBA
project **compiled**, Phase-4 was 35/35, and P5-M, P5-EV, P5-D0 through P5-D8,
P5-DC, Y, Z and P5-FIN all passed. In particular:

* **P5-D1** — all ten locked canonical numeric vectors
* **P5-DP** — all **2 432** binary64 parity vectors on real VBA, every neighbour
  triple still distinct
* **P5-D2** — both decimal-separator injections

The canonical encoder and the compile-safety corrections now have real
target-runtime evidence. Nothing in production VBA is touched this round.

### R5 — located exactly

The diagnostics added after Run 2 did the job they were added for:

```
System.InvalidCastException: Unable to cast object of type 'System.Double'
to type 'System.String'.
  at phase5_gate_b_scenarios.ps1:922   source: $cell.Value2 = $Value
  at Set-Phase5TypedCell -> Reset-Phase5FxTable -> Set-Phase5Fixture
```

The locked seed is `String 'SAR'` then `Double 1`, restored through one helper.
The **String assignment succeeded; the Double assignment through the same source
line failed.** PowerShell binds a COM property setter per call site, so the site
was already bound for a String and the Double could not be marshalled through
it. The accepted Phase-4 `Set-TableCell` never hit this because it has one
assignment line per branch.

This is a harness COM-binding defect. It is not evidence that the workbook wrote
a number as text — it happened while the harness was restoring **its own
captured fixture**.

`Set-Phase5TypedCell` now dispatches on the **captured CLR type**, one COM
assignment site per branch, and refuses an unsupported type by name rather than
coercing it. That is not a retreat to inference: `Set-TableCell` asks what a
value *ought* to be, this asks what Excel *actually published*. A captured
`String '1'` is still written back as `String '1'`; a captured `Double 1` as
`Double 1`. The strict `Test-Phase5ExactValue` read-back is unchanged.

### P5-FX proves the path before anything depends on it

Eleven scenarios reported R5 as their own failure, forty lines into work that
had never started. The restoration path is a prerequisite, so it is now proved
like one: P5-FX captures the untouched seed, restores it **to its own captured
value** through the real `Reset-Phase5FxTable`, and checks value *and* type with
the strict comparator. If that fails, P5-ALL is a FAIL and the scenarios return
— leaving Y, Z and P5-FIN to run, so the lifecycle evidence still arrives.

### The ledger defect: 19 records over 17 IDs

`P5-S2` and `P5-ST` were each recorded twice. The try block recorded both on
real evidence; a **later setup step in the same try** — restoring the base
fixture — then threw, and the enclosing catch recorded both IDs again. Two
things were wrong, and both are fixed:

* **Ownership.** A step that runs after a scenario has been recorded is not part
  of that scenario. The re-establishment now sits in its own try/catch and
  reports as `P5-SU`, a setup failure.
* **Structure.** `Add-Phase5Result` records every ID it emits and refuses a
  second, turning the attempt into a visible Note. Every Phase-5 emission in
  both files goes through it, including the grouped `P5-D1..P5-DC` and
  `P5-S3..P5-RC` catches — the second of which had the same latent risk and is
  audited now rather than after another run. The ledger starts at the first
  Phase-5 entry point, so `P5-PRE` is inside it.

This is not a print-time filter: nothing downstream de-duplicates, and the
ledger genuinely contains one record per ID.

### Why the status failures are not production findings

P5-S1 is what establishes the successful baseline, and it failed **inside**
`Set-Phase5Fixture`, before its calculation. `establishedFingerprint` stayed
empty, so P5-S2 saw `NONE`/`INVALID` with blank fingerprints, P5-ST recalculated
from an unestablished state, and P5-S3 inherited `REFUSED` and
`STRUCTURE CHANGE PENDING`. None of that is evidence about production status
behaviour. Those scenarios must be rerun unchanged on a valid baseline once R5
is confirmed fixed. P5-S4 and P5-KP passed and are untouched.

---

## Review round 4A: exact captured types, and a ledger that fails closed

### Blocker 1 — "exact typed" was not exact for numeric types

The Run-4 setter accepted `Single`, `Int16`, `Int32`, `Int64`, `Byte` and
`Decimal` and wrote all of them as `Double`. That is normalisation, not
restoration: a captured `Int32 1` came back as `Double 1`. And the comparator
could not see it, because it ended in

```powershell
return ([double]$Actual -eq [double]$Expected)
```

so `Int32 1` compared **equal** to `Double 1`. The helper that exists to prove
exact restoration was proving something weaker than it claimed. A `DateTime`
branch was there too, with no runtime evidence behind it and no defined
exact-type contract.

**Supported captured types are now exactly:** an empty cell (`$null`),
`System.String`, `System.Double`, `System.Boolean`. `Double` is matched by exact
type name — `-is [double]` would be true for a boxed `Int32` under PowerShell's
numeric conversions. Everything else reaches the throw, **before any
assignment**, naming the real CLR type. The harness does not get to decide that
some other numeric type "is really" a Double.

**The comparator establishes exact CLR type identity first:**

```powershell
if ($null -eq $Expected) { return ($null -eq $Actual) }
if ($null -eq $Actual) { return $false }
if ($Actual.GetType().FullName -cne $Expected.GetType().FullName) { return $false }
```

then compares by that one already-equal type — text case-sensitively, Double
exactly with no tolerance, Boolean as Boolean. One gate subsumes the three
separate probes it replaces and also catches the numeric pairs they missed. All
nine call sites compare Excel-read against Excel-read or Excel-captured, never
against a JSON-derived expectation, so type identity is the right rule
everywhere it is used.

| captured → restored | verdict |
|---|---|
| `Int32 1` → `Double 1` | **FAIL** |
| `Int64 1` → `Double 1` | **FAIL** |
| `Single 1` → `Double 1` | **FAIL** |
| `Decimal 1` → `Double 1` | **FAIL** |
| `String "1"` → `Double 1` | **FAIL** |
| `Double 1` → `String "1"` | **FAIL** |
| `null` → `""` | **FAIL** |
| `String "1"` → `String "1"` | PASS |
| `Double 1` → `Double 1` | PASS |
| `Boolean True` → `Boolean True` | PASS |
| `null` → `null` | PASS |

### Blocker 2 — a duplicate attempt could still finish green

The one-result guard turned a duplicate into a **Note**, and the driver declares
success from the FAIL count. Notes do not contribute to it. So a future
ownership defect could record `P5-X` as PASS, attempt `P5-X` as FAIL, have the
attempt suppressed, and the run could still report ALL CHECKS PASSED. Fail-open
behaviour in an evidence harness.

A duplicate attempt is now recorded in `$script:Phase5LedgerViolations`, and
**`P5-LDG`** reports on it: PASS when there were none, FAIL naming every
duplicate when there were. It is emitted from the driver **after Y and Z** and
before the summary, so cleanup and lifecycle evidence still arrive; it goes
through `Add-Result` rather than the guard, so the ledger can never suppress its
own report; and it carries its own emitted-once flag, so many duplicate attempts
still produce exactly one `P5-LDG`.

The guard deliberately does **not** throw — that would take the shutdown ledger,
Y, Z and P5-FIN down with it. The invariants hold together:

1. one scenario result per ID — the first stands, no second is appended;
2. the attempt is visible, as a violation and as a Note;
3. any attempt forces the run to FAIL, through `P5-LDG`;
4. it cannot be reduced to a non-failing Note;
5. many attempts still give one integrity result;
6. Y, Z and P5-FIN still run;
7. nothing de-duplicates at print time.

## Runtime Run 5: R5 closed, and a fixture-establishment ordering defect

**Run 5 is VALID EVIDENCE — R5 CLOSED ON REAL WINDOWS; HARNESS
FIXTURE-ESTABLISHMENT ORDERING DEFECT DISCOVERED.** It ran on real Windows
against `a291853`, it is not rerun and not overwritten, and everything below is
read from it.

### What Run 5 closed

**R5 is closed on real Windows.** `P5-FX` PASSED: the locked FX seed was captured
from the untouched Stage-B workbook, written back through the real restoration
path, and read back through the strict typed comparator with its captured CLR
type intact. The typed COM write from Run 4 and the exact-type comparator from
review 4A therefore hold against real Excel, and the failure mode Run 4 recorded
eleven times on one line inside `Set-Phase5TypedCell` did not recur.

That is the finding this section records. The scenarios that DID fail in Run 5
failed for the new root below, which is a different fault in a different place —
it is not R5 returning under another name.

### The new root

`Set-Phase5Fixture` performed its production mutations in an order in which the
driver Adds **could not** succeed, discarded their results, and wrote fixture
data into the rows they had failed to key.

```
Set-Phase5InflationProfileMaster    rewrites Config!tblInflationProfiles
                                    -> Inflation!tblInflation still holds the
                                       PREVIOUS fixture's applied profiles
$Excel.Run('PCCM_AddCostLine')      -> RunDriverOperation
      | Out-Null                    -> AddDriver allocates CL-001, writes it
                                    -> ValidateStructure
                                    -> CheckInflationProfiles: master and grid
                                       disagree
                                    -> Err.Raise 5002
                                    -> TryRestoreDriver rolls the register, the
                                       profiling grid AND the ID counter back
                                    -> RecordResult "FAIL|..."
                                    -> the harness reads nothing
Write-Phase5Driver ... -RowIndex 1  writes description, quantity, unit costs,
                                    currency, profile and distribution into a
                                    row that carries no key
$Excel.Run('PCCM_ApplyTimeline')    -> PreMutationCheck
                                    -> [no_orphan_structural_data]
                                       tblCostLines row(s) 1 hold data but
                                       carry no key.
```

**The refusal is correct production behaviour.** `PCCM_ApplyTimeline` refused to
synchronise over unkeyed structural data because synchronisation would have
erased it silently, which is exactly what `PreMutationCheck` exists to prevent.
Production is not at fault anywhere in this chain. The orphan row was
manufactured by the harness.

### Root and cascade

| | |
|---|---|
| **Root** | `Set-Phase5Fixture` invoked `PCCM_AddCostLine` / `PCCM_AddRisk` while the Config profile master and `Inflation!tblInflation` disagreed, discarded `PCCM_AutomationResult`, and wrote driver data into the unkeyed row the failed Add left behind |
| **Production defect** | none |
| **Reported as** | `[no_orphan_structural_data] tblCostLines row(s) 1 hold data but carry no key.` from `PCCM_ApplyTimeline`, several statements after the operation that actually failed |
| **Cascade** | every scenario that establishes a fixture. Derived from the harness source — the scenarios that call `Set-Phase5Fixture` — not counted from the run transcript: `P5-AN`, `P5-RF`, `P5-PQ`, `P5-PN`, `P5-AR`, `P5-ID`, `P5-S1`, `P5-S2`, `P5-ST`, `P5-SU`, `P5-NS`, `P5-S3`, `P5-S4`, `P5-S5`, `P5-KP`, `P5-RC`, `P5-S6`, `P5-FA`, `P5-FC` |
| **Not cascade** | `P5-PRE`, `P5-P4`, `P5-FX`, `P5-M`, `P5-EV`, `P5-D0`–`P5-D8`, `P5-DP`, `P5-DC`, `P5-LDG`, `P5-FIN` — none of them establishes a fixture |
| **Scenarios that reached their own predicate** | none of the cascade |

Not one cascade scenario reached the predicate it exists to test, and that is a
consequence of the mechanism rather than a tally: `Set-Phase5Fixture` **throws**,
so the scenario's own `catch` records a FAIL before a single one of its checks is
evaluated, and the text it records is the fixture's message about `tblCostLines`
row 1. The thing that was actually broken — the order in which the fixture
performs its production mutations — had no result of its own anywhere in the
ledger.

### Why the order was wrong, and why moving one call fixes it

`Config!tblInflationProfiles` is the profile master. `modInflation.SyncProfileRows`
rebuilds `Inflation!tblInflation` from it, and it runs inside
`PCCM_ApplyTimeline`. So writing the master **creates** a disagreement that only
an Apply can close.

`modDrivers.RunDriverOperation` runs `modStructuralCheck.ValidateStructure()`
after every successful Add, and `ValidateStructure` includes
`CheckInflationProfiles`, which compares the grid against the master. An Add
attempted between "master rewritten" and "Apply" is therefore an Add attempted
over a workbook production is **required** to call incoherent. It cannot succeed.

Moving the Apply above the Adds makes that window empty. By the time the first
Add runs, the grid has been rebuilt from the master and the two agree. Nothing
about production changed: the operations are simply performed in an order where
each one's preconditions actually hold.

The Adds need no second Apply behind them. `modProfiling.SyncRows` preserves the
year columns Apply generated and only adds the new driver's profiling row, and
each Add revalidates the structure itself.

### The corrected choreography

`Set-Phase5Fixture` is now eight marked steps, and **every production mutation
begins and ends structurally coherent**:

| Step | What happens | What is proved |
|---|---|---|
| **A** | empty the registers, reset the identity counters | every delete returns `OK\|*`; the identifier is gone; no keyed row and no unkeyed data survives, checked on the way in **and** on the way out; each counter reads back through the typed reader as the exact numeric initial |
| **B** | the Setup scalars | — |
| **C** | restore the captured FX seed, append the fixture's rows | the seed round-trips with its captured type (unchanged from `P5-FX`) |
| **D** | write the Config profile master | the workbook is now **knowingly incoherent**; no production endpoint may be called until E |
| **E** | `PCCM_ApplyTimeline`, **with no driver added yet** | `PCCM_AutomationResult` is `OK\|*`; `PCCM_StructuralReport` is blank |
| **F** | the drivers, one checked Add each | per Add: result is `OK\|*`; the register holds exactly *n* keyed rows; the target row carries a permanent identifier; that identifier **is** the emitted `permanent_id`, compared binary — and only then `Write-Phase5Driver` |
| **G** | inflation rates and profiling weights | (both already keyed by name and by header) |
| **H** | — | `PCCM_StructuralReport` is blank again |

Every one of those checks **throws**. Fixture establishment fails loudly or not
at all: a postcondition that returned a diagnostic, wrote a note or continued to
the next driver would put the run back where Run 5 was.

### `PCCM_AutomationResult` is the authority — and it must be this operation's

Two rules came out of the root, and both are structural now.

**The result is the verdict.** That `Excel.Run` returned says only that VBA did
not raise across the COM boundary. `modAppState.Announce` records `OK|…` or
`FAIL|…`, and that recorded string is the operation's own judgement. Every
production mutation the Phase-5 harness makes now goes through
`Invoke-Phase5ProductionOperation`, which reads it and requires `OK|*`.

**The result is cleared first.** `gAutomationLastResult` is a single global that
survives until something overwrites it. An endpoint that failed *before* reaching
`RecordResult` would leave the previous operation's `OK|…` in place, and a reader
that did not clear first would accept a stale success as this operation's own —
fail-open, in exactly the place the correction is about.
`PCCM_AutomationBegin` calls `ClearAutomation`, so the helper arms immediately
before the call and the value that comes back is this operation's or nothing at
all. This is not an invention: the accepted Phase-4 `Set-AppliedTimeline` has
used that idiom since the matrix was written.

That reasoning also applied to two sites in `Invoke-Phase5Mutation`, and both now
go through the same helper. Requiring `OK|*` there is **not** the clean-structure
gate that `test_77` forbids: the deformation the mutation applier exists to
perform is still free to make the model refuse at calculation time. It only
requires that the operation the corpus asked for actually happened.

### The harness must expose contamination, not launder it

`Assert-Phase5NoUnkeyedRegisterData` is the harness-side mirror of
`modWorkbook.OrphanRows`, term for term: the key column's text is empty and some
other column in the row is not. It runs **before** `Clear-Phase5Registers`
empties a register as well as after.

Before, because a fixture that inherits a contaminated register from an earlier
scenario must fail naming the earlier scenario's damage — not silently delete it
and carry on. There is deliberately **no repair path**: an orphan row is not
blanked, not adopted and not deleted. It is left exactly where it is so the run
reports it rather than erasing the evidence.

`Clear-Phase5Registers` also proves each delete **took**. A production `Delete`
answers a declined confirmation with a *success* that removes nothing, so `OK|`
alone does not mean gone.

And the counters are read back through `Get-Phase5TypedNamedValue` and compared
with `Test-Phase5ExactValue`. The accepted `Get-NamedValue` stringifies, and it
cannot tell a numeric counter of `0` from a **text** counter of `"0"` —
`modDrivers.TryReadCounter` draws exactly that distinction and refuses a counter
that is not a whole number, so a reset that landed as text would make the very
first Add refuse. Same typed discipline the Run-4 correction established for
cells.

### `P5-FIX` — the fixture proves itself first

Run 5 reported an ordering defect only through the scenarios that inherited it,
none of which had reached its own predicate. A prerequisite is proved like a prerequisite:
once, at the point where it can still be reported as itself. That is what `P5-FX`
already does for the FX seed, and `P5-FIX` now does for fixture establishment.

It runs immediately after `P5-D8` and before `P5-AN` — the first fixture the run
establishes is its own — and it drives the **real** `Set-Phase5Fixture`, not a
copy. It asserts what fixture establishment is supposed to have achieved:

1. the Apply reported `OK|*`;
2. `PCCM_StructuralReport` is blank;
3. each register carries exactly the emitted identifiers, in order, compared
   binary;
4. no register row holds data without a key;
5. the Config master holds exactly the emitted profiles;
6. `SyncProfileRows` rebuilt `tblInflation` to agree with the master — the check
   that names the *reason* rather than the symptom if the order ever regresses;
7. the baseline actually calculates.

Checks 1 to 6 are claims about what the **harness** did. Check 7 can fail for a
production reason instead, so the gate below reports what was observed rather
than deciding whose fault it is, and the attempt detail is carried into the
checklist so the distinction can be made from the evidence.

A failure is a **FAIL**, never a SKIP, and it gates through `P5-ALL` and
`return` exactly as `P5-FX` does. Returning from the scenario driver leaves the
caller's shutdown, `Z`, `Y`, `P5-LDG` and `P5-FIN` untouched, so the lifecycle
evidence is still produced.

### What did not change

No production VBA. `modDrivers`, `modStructuralCheck`, `modInflation`,
`modProfiling`, `modTimeline`, `modAppState`, `modCalcResolve`, `modCalcCheck`,
`modCalcReport`, `modCalcFingerprint`, `modCalcFactors` and every other
production module are byte-identical. The orphan refusal is correct production
behaviour and `ValidateStructure` is not suppressed, weakened or worked around
anywhere: the correction is entirely an ordering and result-checking change in
the harness. The builder, the spec, the canonical oracle and the emitted corpus
are unchanged, and no status logic was touched.

### The fifteen regressions

| | Claim |
|---|---|
| R1 | the timeline is applied before the first driver Add, exactly once |
| R2 | no production mutation runs inside the D→E incoherence window |
| R3 | no production endpoint is invoked with its result discarded |
| R4 | every checked operation requires `OK\|*` and throws otherwise |
| R5 | the result is cleared before the operation it is supposed to describe |
| R6 | the row is proved keyed before any fixture data is written |
| R7 | the issued identifier is the emitted `permanent_id`, compared binary |
| R8 | `Write-Phase5Driver` has exactly one call site, inside the checked Add |
| R9 | every delete is checked, and proved to have taken |
| R10 | the registers are proved empty and free of unkeyed data, at both ends, with no repair path |
| R11 | the identity counters are proved typed and exact |
| R12 | coherence is proved at both ends of the fixture |
| R13 | a failed fixture postcondition throws — no note, no warning, no continue |
| R14 | `P5-FIX` gates the scenarios that depend on it, and the lifecycle evidence survives the gate |
| R15 | the Run-5 defect replayed as a state machine, both orders |

R15 models production from its own source contract — an Add revalidates the
structure and rolls its identifier allocation back when that fails, Apply
rebuilds the inflation grid from the Config master, and a mutation over an
unkeyed row is refused — then runs the old order and the new one and compares the
outcomes. The old order reproduces
`[no_orphan_structural_data] tblCostLines row(s) 1 hold data but carry no key`
and leaves the counter rolled back to 0. The new order ends with `CL-001` in row
1, no orphan, and a coherent structure. A third replay shows the harness half of
the correction: with the ordering regressed, the checked Add still refuses to let
fixture data reach the unkeyed row.

Twelve source mutations were planted against the corrected harness and each was
caught by at least two independent tests: the Run-5 order restored (5 tests), an
unchecked cost Add (2), an unchecked risk Add (2), the result no longer cleared
before the operation (2), a case-insensitive identifier comparison (3), the
delete postcondition dropped (3), the entry orphan scan dropped (2), the counter
read-back stringified (3), the closing coherence gate dropped (2), the driver
write hoisted above the key proof (2), the self-proof gate downgraded to a SKIP
that does not return (2), and the operation-result gate disarmed (3).

### Status

Gate B is **not** accepted. Phase 5 is **not** accepted. No runtime execution is
requested yet.

**NO WINDOWS/EXCEL RUNTIME WAS EXECUTED DURING THIS CORRECTION ROUND.**

## Runtime Run 6: `{}` is a valid rate mapping

**Run 6 is VALID EVIDENCE — RUN-5 FIXTURE CHOREOGRAPHY PROGRESSED THROUGH
STRUCTURAL APPLY AND CHECKED DRIVER CREATION; P5-FIX THEN EXPOSED AN
EMPTY-OBJECT POWERSHELL PROPERTY-ENUMERATION DEFECT.** It ran once on real
Windows against `5fd8992`, it is not rerun and not overwritten.

Result: **53 passed, 2 failed, 0 skipped** — `P5-FIX` and `P5-ALL`. `P5-ALL` is
not an independent defect: it is the dependency gate reporting that the
dependent scenarios were not attempted. **Run 6 exposes one root.**

### What the gate bought

This is what `P5-FIX` was added for. A fixture defect was reported as a fixture
defect, with the failing statement named, instead of nineteen downstream
scenarios each reporting the same inherited symptom as its own predicate
failing. Run 5 needed a review round to locate its root from the cascade; Run 6
handed it over in one result.

### The root

```
phase5_gate_b_scenarios.ps1:1847
    foreach ($year in $rates.PSObject.Properties.Name) {

System.Management.Automation.RuntimeException:
    The property 'Name' cannot be found on this object.

Write-Phase5InflationRates -> Set-Phase5Fixture -> P5-FIX
```

The golden fixture is plan case 1, *SAR, no inflation, one project year*, and
its emitted model carries

```json
"inflation": { "Standard": {} }
```

**That is not malformed data.** The inflation grid spans Base Year + 1 through
Start Year + Duration − 1. Case 1 is Base 2026, Start 2026, Duration 1, so that
span runs 2027..2026 — it is empty, and there is no calendar year a rate could
belong to. `{}` is the correct encoding of "there is nothing to inflate".

It is not a one-case curiosity either. **Eleven of the twenty-eight emitted
models carry an empty rate mapping**, and they are exactly the eleven whose
span is zero: cases 1, 2, 6, 7, 8, 16, 17, 18, 22, 25 and 30. The correlation is
proved in both directions by regression, so a rate map that were empty for any
*other* reason would fail as a corpus defect rather than pass as a harness one.

### Why the expression fails

`$collection.Name` is **member enumeration**: PowerShell projects the member
across every element. Under `Set-StrictMode -Version 2.0` an empty collection
cannot answer it and raises `The property 'Name' cannot be found on this
object.` rather than returning zero names.

`.PSObject.Properties` has no such edge. It is a real property of a real
`PSObject` whether or not the object holds any members, and `foreach` over an
empty collection is zero iterations. So the correction is to enumerate the
`PSPropertyInfo` objects and read `Name` and `Value` off each individual one:

```powershell
foreach ($profileProperty in $Model.inflation.PSObject.Properties) {
    $name  = [string]$profileProperty.Name
    $rates = $profileProperty.Value
    $rowIndex = Find-GridRow -Workbook $Workbook -Grid $grid -Key $name
    foreach ($rateProperty in $rates.PSObject.Properties) {
        $year      = [string]$rateProperty.Name
        $rateValue = $rateProperty.Value
        ...
    }
}
```

There is **no guard, no count test and no early return.** An empty collection is
simply an empty collection. A `Count -gt 0` guard would have worked and would
still have been wrong: it leaves the projection in place for every non-empty
case and puts the harness one corpus change away from the same failure
elsewhere.

The value now comes off the **same property object** the year name came from,
which also removes the `$rates.$year` dynamic lookup — the blank branch can no
longer be reached through a lookup that missed.

### `{}` versus `{"2028": null}`

These are different, and both survive:

| Shape | Meaning | Cells written | Where |
|---|---|---|---|
| `{}` | zero rate entries | **none** — the loop body never runs | cases 1, 2, 6, 7, 8, 16, 17, 18, 22, 25, 30 |
| `{"2028": null}` | calendar year 2028 **is** an entry, and its cell must be **BLANK** | **one**, blank | case 14, the blank-required-rate refusal |

Collapsing the second into the first destroys case 14: the refusal it exists to
prove could never fire, because the rate the model says is blank would simply
never be written. Collapsing the first into an error is what Run 6 did. The
`[double]$null` guard from the Run-4 round is unchanged and still branches
before the cast, so a blank never becomes a numeric zero.

### The audit: every `.PSObject.Properties.Name` site

| Function / site | Container enumerated | Can legally be empty? | Evidence | Action |
|---|---|---|---|---|
| `Write-Phase5InflationRates` | **`$rates`** (a profile's rate map) | **YES** | empty in 11 of 28 emitted models, exactly the zero-span timelines | **CORRECTED** — enumerates `PSPropertyInfo` objects |
| `Write-Phase5InflationRates` | `$Model.inflation` | no | every model has ≥1 driver, and the profile set **equals** the set of profiles those drivers reference | converted with the inner loop — same procedure, same expression class, one line apart |
| `Set-Phase5Fixture` | `$Model.inflation` | no | same derivation | left as it is — proven site, no churn; pinned by regression |
| `Invoke-Phase5GateBScenarios`, in the `P5-FIX` block | `$golden.model.inflation` | no | same derivation | left as it is; pinned by regression |
| `Get-CalcScalarBlock` | `$Inspection.calc.scalar_blocks.<block>.rows` | no | both emitted blocks carry rows | left as it is; pinned by regression |
| `Get-Phase5Snapshot` | `$Inspection.calc.tables` | no | five tables, always emitted | left as it is; pinned by regression |
| `Add-Phase5AnalyticalChecks` ×4 | `$wanted` — one emitted `calc_years` / `resolved_fx_rows` / `drivers` / `annual` row | no | 40 + 19 + 21 + 40 row objects emitted, none empty | left as it is; pinned by regression |
| `Add-Phase5AnalyticalChecks` | `$expected.resolved_fx` | no | non-empty in all 19 expected blocks | left as it is; pinned by regression |
| `Add-Phase5AnalyticalChecks`, `P5-ID` ×2 | `$expected.totals` | no | ten fields, all 19 expected blocks | left as it is; pinned by regression |

The "no churn" half of that instruction and the "nothing waiting for Run 7" half
are both honoured by making the audit **executable**. `test_183` re-derives the
projection list from the source, refuses an unclassified container outright, and
re-proves every class-B non-emptiness claim against the artifacts the builder
actually emits. A future corpus case that makes any of those containers empty
fails that test by name, here, rather than failing the next Windows run.

Note that the class-B claim for `$Model.inflation` is a *derivation*, not a
tally: it is non-empty **because** every model has at least one driver and every
driver names a profile. The test asserts that derivation, not just the current
counts.

### What Run 6 proved works

`P5-FIX` entered the real `Set-Phase5Fixture` and reached **step G**. Every
production mutation before step G is fail-closed and throws on any failed
postcondition, so reaching `Write-Phase5InflationRates` at all is real-Windows
evidence that, on the path taken:

* the registers emptied and the counters reset (step A);
* `PCCM_ApplyTimeline` returned `OK|*` with a blank `PCCM_StructuralReport`
  (step E);
* every checked `PCCM_AddCostLine` / `PCCM_AddRisk` returned `OK|*` and its row
  carried the emitted `permanent_id` before any driver data was written
  (step F).

**The Run-5 orphan/add-order root was not reproduced on the path reached.** That
is not the same as a `P5-FIX` PASS — `P5-FIX` failed before completing, and the
closing coherence proof at step H never ran.

Also PASS in Run 6: the Phase-4 matrix 35/35, `P5-FX`, `P5-M`, `P5-EV`,
`P5-D0`–`P5-D8`, `P5-DP` at 2 432/2 432, `P5-DC`, `Y`, `Z`, `P5-LDG` with zero
duplicate attempts, and `P5-FIN`. `Workbook.Close`, `Application.Quit` and the
natural PID exit all true.

### The gate itself behaved correctly

`P5-FIX` FAIL → `P5-ALL` FAIL, *not attempted*, and the analytical, status and
rollback scenarios were not allowed to emit a cascade — while `Z`, `Y`,
`P5-LDG` and `P5-FIN` all still completed. That behaviour is preserved
unchanged, and `test_182` pins it, including the exact message Run 6 recorded.

No production calculation or status finding is available from Run 6, because the
dependent Phase-5 scenarios were correctly not attempted.

### The limit of the proof in this round

Seven source mutations were planted against the corrected writer and each was
caught by at least two independent tests: the exact Run-6 projection restored
(6 tests), the profile loop projecting `.Name` again (4), a `Count` guard that
leaves the projection in place (6), a null rate skipped instead of written blank
(2), the rate value resolved dynamically again (2), the profile row located
inside the rate loop (2), and the gate message reduced to a bare "not attempted"
(2).

The regressions that model the enumeration semantics are **Python models of the
intended behaviour**. They prove what the corrected shape must do with `{}`,
with one rate, with several, and with a null-valued rate, and that the first two
shapes are distinct paths. They do **not** claim to reproduce the Windows
PowerShell property adapter. Whether the real adapter yields zero properties for
an empty `PSCustomObject` is a runtime fact, and the next Windows run is where
it is proved.

### What did not change

No production VBA — `modInflation`, `modStructuralCheck`, `modDrivers`,
`modCalcReport`, `modCalcFingerprint`, `modCalcFactors` and every other module
are byte-identical. No builder, spec, oracle or corpus change: plan case 1 still
carries `{}`, plan case 14 still carries its blank required rate, and the
canonical parity corpus is untouched. No calculation or status logic was
touched. This was a PowerShell harness enumeration defect and the fix is
confined to one procedure.

### Status

Gate B is **not** accepted. Phase 5 is **not** accepted. No runtime execution is
requested yet.

**NO WINDOWS/EXCEL RUNTIME WAS EXECUTED DURING THIS CORRECTION ROUND.**

## Runtime Run 7: callability is not compilation

**Run 7 is VALID EVIDENCE — RUN-6 EMPTY-INFLATION ENUMERATION CORRECTION PASSED
THE PREVIOUS RUNTIME FAILURE AND THE GOLDEN FIXTURE REACHED PCCM_CALCULATE;
PCCM_CALCULATE THEN EXPOSED A REAL PRODUCTION VBA COMPILE DEFECT IN THE
ANALYTICAL PATH.** It ran once on real Windows against `37f2dfd`; it is not
rerun and not overwritten.

Result: **53 passed, 2 failed, 0 skipped** — `P5-FIX` and `P5-ALL`, and `P5-ALL`
is the dependency gate reporting that the dependent scenarios were not
attempted. One root.

### The Run-6 correction is closed by real runtime evidence

Run 6 died at `$rates.PSObject.Properties.Name` inside
`Write-Phase5InflationRates`. Run 7 went straight past it and reached
`PCCM_Calculate`. The empty-object enumeration correction now has real Windows
evidence behind it and is frozen; the `{}` versus `{"year": null}` semantics are
not reopened.

### The root

```
$Excel.Run('PCCM_Calculate')   ->   HRESULT 0x800A9C68

VBE:  Compile error: Sub or Function not defined
      highlighting   Contribute(...)
      inside         modCalcAnalytical.AccumulateTotals
```

### Why a procedure that exists is reported as not defined

`Contribute` was declared, exactly once, in that same module — and its
declaration read

```vba
Private Function Contribute(ByRef terms() As Double, ByVal slot As Long, _
                            ByVal value As Double, ByRef scale As Double, ...
```

`Scale` is a VB statement keyword. The parser rejects it in a declaration
position, so the declaration never produced a procedure, so every call to
`Contribute` was a call to a name that did not exist. The VBE reports the
*symptom* at the call site; the *cause* is four parameters into the declaration
forty lines away.

This is the same class Run 3 found at `Dim scale As Long` in
`modCalcFingerprint`. What changed is that Run 3 looked like one bad line and
Run 7 proves it is a class.

### The authority that let it through

`test_phase5_vba_source.py` held a `COMPILE_PROVEN_RESERVED_SITES` map of
fifteen grandfathered declarations. Its stated authority was:

> Runtime Run 2 imported all fifteen modules, reached P5-M, and confirmed every
> API procedure callable, **which is only possible if the whole project
> compiled.**

**Run 7 disproved that inference inside a single Excel session:**

| | |
|---|---|
| `A1` | **PASS** — `PCCM_AutomationBegin` is callable |
| `P5-M` | **PASS** — fifteen modules present, and six API procedures *reported* callable under the evidence model P5-M then had. One of the six, `PCCM_Calculate`, had never crossed `Application.Run`; that borrowed claim was removed in the review of ae52bdd, and P5-M now proves **six declared, five callable** |
| `P5-FIX` | **FAIL** — `PCCM_Calculate` → VBE compile error |

VBA compiles on demand. A project answers an API call while a procedure body
nothing has reached yet still holds a fatal declaration. So callability of one
procedure — or of six — is not proof that every deferred body compiled, and
**every one of the fifteen sites rested on that inference, not just the one Run
7 happened to reach first.**

### The correction: fifteen renames, and no exemption mechanism

All fifteen were renamed from the procedure's own semantics. No numerical
change, no reordering, no new helper, no tolerance change.

| # | Module | Procedure | Kind | Was | Now |
|---|---|---|---|---|---|
| 1 | `modCalcAnalytical` | `AnnualSeries` | variable | `width` | `groupWidth` |
| 2 | `modCalcAnalytical` | `Contribute` | parameter | `scale` | `measureScale` |
| 3 | `modCalcAnalytical` | `Identity` | parameter | `scale` | `conditioningScale` |
| 4 | `modCalcAnalytical` | `Pair` | parameter | `scale` | `combinedScale` |
| 5 | `modCalcAnalytical` | `Reconcile` | variable | `scale` | `identityScale` |
| 6 | `modCalcAnalytical` | `ScaleOne` | parameter | `scale` | `groupScale` |
| 7 | `modCalcAnalytical` | `TotalIdentity` | variable | `scale` | `pairedScale` |
| 8 | `modCalcFactors` | `BuildFactor` | variable | `width` | `groupWidth` |
| 9 | `modCalcFactors` | `ExactAddShifted` | variable | `scale` | `subLimbScale` |
| 10 | `modCalcFactors` | `ExactAnyBelow` | variable | `scale` | `bitScale` |
| 11 | `modCalcFactors` | `IdentityAllowance` | parameter | `scale` | `termScale` |
| 12 | `modCalcFactors` | `RoundExact` | variable | `scale` | `scaleExponent` |
| 13 | `modCalcFingerprint` | `CalcFpEncodeSection` | parameter | `name` | `sectionName` |
| 14 | `modCalcReport` | `CountCurrencyReferences` | variable | `currency` | `currencyIndex` |
| 15 | `modCalcResolve` | `DistributionKindOf` | parameter | `name` | `distributionName` |

The names come from what each identifier *is*: `groupWidth` is the number of
factors in one project year's product group; `subLimbScale` is the power of two
for the sub-limb remainder of a shift; `scaleExponent` is a binary exponent, not
a magnitude, which is why it is `As Long`.

**The grandfather mechanism is gone, and nothing replaced it.**
`COMPILE_PROVEN_RESERVED_SITES` is now empty and the rule is: **zero
declarations using a curated reserved identifier, in any production module.**
`test_86` asserts the sweep is empty in both directions; `test_87` proves a
planted reserved declaration is rejected regardless of module, procedure,
declaration kind, or whether that spelling was there historically — including
the exact site Run 7 rejected. A rule with no exceptions cannot rot.

### Proof that the calculation did not move

Three independent proofs, because a rename that also changed an expression would
close a compile class and open a numerical one:

1. **Reversal to the base commit.** `test_86b` reverses each new identifier
   inside the procedure it belongs to and requires the result to equal
   `37f2dfd`'s blob **byte for byte**, for all five modules.
2. **The frozen fingerprint digest.** `modCalcFingerprint`'s body digest moved,
   and `test_64j` does not merely accept the new number: it reverses
   `sectionName` → `name` over the identical reduction and requires the
   *previous* digest back.
3. **The contribution structure.** `test_91` reads all twelve `Contribute` calls
   in `AccumulateTotals` and pins the array, the conditioning accumulator and
   the contributed value of each, in order — two into D, six into A/B/C, then
   four into E from two passes. `test_92` pins `IdentityAllowance`'s full
   signature and its one caller's five positional arguments.

Comments and string literals were deliberately **not** rewritten: `Contribute`'s
commentary still says "conditioning scale" in English, and `Reconcile`'s
diagnostic still ends `& " scale"` — that literal is user-facing text, and
renaming it would have been a behaviour change smuggled in as a refactor. Three
such hits were reverted after the mechanical pass, and `test_94` pins them.

### `P5-CMP` — a real whole-project compile gate

The claim "the project compiles" now has a scenario of its own, and it runs
**after `P5-P4` and before `P5-FX`, `P5-FIX` and every fixture** — before any
statement that touches the workbook.

It drives the VBE's own Compile VBAProject command:

```powershell
$control = $bars.FindControl($null, 578)
```

- **By ID, never by caption.** 578 is Compile VBAProject. A caption lookup finds
  nothing on a non-English Excel and would report success for a project that
  never compiled. `test_187` asserts `FindControl` receives only `$null, 578`
  and that `.Caption` is never read.
- **The control must exist.** Its absence is a failure of the gate, not a pass —
  a project that cannot be proved to compile here is not a project that
  compiles.
- **Enabled is the evidence.** The VBE enables the command while something is
  left to compile and disables it once the project is fully compiled. So
  `Enabled = False` before means already compiled, and `Enabled = False` after a
  successful `Execute` is the positive proof. Still enabled afterwards is a
  FAIL, not an explanation.
- **Fail closed, evidence intact.** Every step is inside try/catch and a throw
  is a FAIL carrying Excel's own message. Nothing suppresses, auto-answers or
  dismisses a compile-error dialog — `DisplayAlerts` is left exactly as the
  accepted lifecycle set it, because a dismissed dialog is a destroyed
  diagnostic. `test_188` sweeps for `-ErrorAction SilentlyContinue`, `SendKeys`,
  `$ErrorActionPreference` and the rest.
- **No false proof.** The gate never claims success from importing modules or
  from invoking a macro; `test_188` asserts `VBComponents.Import`,
  `PCCM_AutomationBegin` and `PCCM_Calculate` appear nowhere inside it.
- **One result, and the lifecycle survives.** Exactly two `Add-Phase5Result
  'P5-CMP'` sites — the success path and the catch path — and a failure gates
  through `P5-ALL` and **returns**, leaving shutdown, `Z`, `Y`, `P5-LDG` and
  `P5-FIN` reachable. All three COM transients are released.

**A caveat stated rather than hidden:** whether `Execute()` on control 578 is
reliable in every VBE build, and whether `Enabled` is a dependable
post-condition there, are runtime facts. The gate is written so that any
deviation is a **FAIL with evidence** rather than a silent pass, and it does not
fall back to the retired A1 assumption under any circumstance. If the next
Windows run shows the mechanism is unreliable, that is a finding about the gate,
which is exactly what a gate is for.

### The corrected claims

`A1`'s check label was `PCCM_AutomationBegin is callable (the VBA project
compiles)`. It is now `PCCM_AutomationBegin is callable` — what it actually
observes. `P5-M`'s API checks are unchanged by THIS round and still recorded
(the review of ae52bdd later found that one of its six callability claims was
itself borrowed, and split them into six declared and five callable);
what is removed is the conclusion drawn from them. Both files carry the Run-7
counterexample where the claim used to be, so it cannot be quietly restored, and
`test_190` asserts no check label in either file claims compilation except
`P5-CMP`'s own two.

The useful callability evidence is kept. Only the overstated conclusion is gone.

### What Run 7 positively proved

The Run-6 root did not recur; the golden fixture progressed to
`PCCM_Calculate`; `P5-FX`, `P5-M`, `P5-EV`, `P5-D0`–`P5-D8`, `P5-DP` at
2 432/2 432 and `P5-DC` all PASSED; the Phase-4 matrix was 35/35; `P5-LDG`
recorded zero duplicate attempts; `Y`, `Z` and `P5-FIN` PASSED;
`Workbook.Close`, `Application.Quit` and the natural PID exit were all true.

`P5-FIX` **did not pass** — it failed before completing. No analytical, refusal,
prerequisite, audit, status or rollback evidence exists from Run 7, because the
dependency gate correctly prevented those scenarios from running, and none of it
is interpreted here.

### Scope

Production VBA changed in exactly one way: fifteen declaration identifiers.
`modCalcAnalytical`, `modCalcFactors`, `modCalcFingerprint`, `modCalcReport` and
`modCalcResolve` are otherwise byte-identical to `37f2dfd`, proved by reversal.
No builder, spec, oracle or corpus change. The Run-6 PowerShell enumeration
correction is untouched; the only harness changes are `P5-CMP` and the two
corrected labels.

### Mutation controls

Eight mutations were planted and each was caught by at least two independent
tests:

| | Mutation | Detectors |
|---|---|---|
| M1 | `Contribute`'s parameter back to `ByRef scale As Double` | 6 |
| M2 | a different analytical `scale` declaration restored (`TotalIdentity`) | 5 |
| M3 | a `modCalcFactors` `scale` declaration restored (`ExactAnyBelow`) | 4 |
| M4 | `CalcFpEncodeSection`'s `name` restored | 6 (incl. the frozen digest) |
| M5 | `CountCurrencyReferences`' `currency` restored | 4 |
| M6 | `DistributionKindOf`'s `name` restored | 4 |
| M7 | a one-site grandfather exemption reintroduced | 2 |
| M8 | the false A1 text restored | 2 |

### Status

Gate B is **not** accepted. Phase 5 is **not** accepted. No runtime execution is
requested yet.

**NO WINDOWS/EXCEL RUNTIME WAS EXECUTED DURING THIS CORRECTION ROUND.** The
static tests above do not execute the VBA compiler and do not claim to; the
compile class is closed by construction, and `P5-CMP` is what will test it on
real Excel.

## Review of ae52bdd: one compile authority, and no borrowed evidence

Independent review accepted the Run-7 production renames provisionally and
rejected the package on two evidence-authority blockers. Both were real.

### Blocker 1 — the retired A1 authority was still standing

The Run-7 round retired "A1 proves the project compiles" from A1's own check
label and stopped there. The same claim was still asserted in eight other
places. What was there, and what is there now:

| File | Was | Now |
|---|---|---|
| `phase4_functional_test.ps1` (Phase-5 overview) | `P5-D0 … imported only AFTER A1 has proved the production project compiles` | a `P5-CMP` entry was added above it, and P5-D0 now names `P5-CMP` |  <!-- retired-authority: quoted to record what was removed -->
| `phase5_gate_b_diagnostics.bas` (module header) | `only AFTER scenario A1 has proved the production VBA project compiles` | `only AFTER scenario P5-CMP has proved …` |  <!-- retired-authority: quoted to record what was removed -->
| `phase5_gate_b_scenarios.ps1` (import commentary) | `Scenario A1 has already made the first Application.Run … so the proof that the accepted project compiles is complete and unmasked` | P5-CMP has driven Compile VBAProject; the paragraph records why A1 was demoted |
| `phase5_gate_b_scenarios.ps1` (P5-D0 result title) | `Transient diagnostic module imported AFTER the A1 production compile` | `… AFTER the P5-CMP whole-project compile` |  <!-- retired-authority: quoted to record what was removed -->
| `docs/phase5_gate_b_harness.md` (lifecycle) | `A1 … -> the PRODUCTION project compiles` | `A1 … -> the automation surface ANSWERS`, with a `P5-CMP` line added |
| `docs/phase5_gate_b_harness.md` (prose) | `A1 remains the first real VBA compilation boundary` | **A1 is the first `Application.Run` boundary, not a compilation boundary** |  <!-- retired-authority: quoted to record what was removed -->
| `docs/phase5_plan.md` (Phase-4 preservation) | `the VBA project still compiles \| A1 unchanged` | `P5-CMP`, with A1's real role stated |
| `docs/phase4_gate_b_run4.md` (run record) | `the real VBA project compiled and was callable` | `the real VBA automation surface answered`, with the correction dated |
| `docs/phase4_gate_b_final.md` | `real VBA compilation and callability, through A1` | `real VBA automation-surface callability, through A1` |
| `tests/…harness_source.py` | module docstring, `test_22`'s name, docstring and messages | all restated; `test_22` now also requires the import to follow **P5-CMP** |

The two Phase-4 run records were **corrected in place, not rewritten**. Those
runs did report A1 PASS and that observation stands; what was wrong was the
conclusion drawn from it, so each line now says what was observed and names the
round that retired the rest.

`test_191` sweeps every `.ps1`, `.bas`, `.py` and `.md` under `bootstrap`,
`docs`, `tests`, `src`, `builder` and `spec` for nine retired formulations. A
line may quote one only if it carries a `retired-authority` marker — the same
idiom the COM-lifecycle sweep already uses for its refusal list — and the test
requires at least three such markers to exist, so a sweep that had stopped
matching anything would fail rather than pass quietly.

### The evidence hierarchy

| Scenario | Claims | Does not claim |
|---|---|---|
| `A1` | the first production `Application.Run` returned; the automation surface answers | anything about compilation |
| `P5-CMP` | the **whole** VBA project compiles, via the VBE's Compile VBAProject command | — |
| `P5-D0` | the diagnostic module was imported **after** `P5-CMP` passed | — |
| `P5-M` | six API procedures are **declared**; five of them are **callable** | that `PCCM_Calculate` is callable |
| `P5-FIX` | the **first** `PCCM_Calculate` of the run, on a valid fixture | — |

### Blocker 2 — P5-M passed PCCM_Calculate as callable without calling it

The old loop branched on the name, skipped the invocation, set `$callable =
$true` anyway, and emitted **"the API procedure PCCM_Calculate is callable"** as
a PASS with the detail "exercised by the analytical scenarios below". The name
had not crossed `Application.Run`. An expected future exercise was counted as
present evidence — the same species of overclaim as A1's.

The fix is not to call it. `PCCM_Calculate` is stateful: it resolves the model,
writes the `_Calc` block and publishes a status, so running it at inventory time
against whatever the workbook happened to hold would establish a snapshot no
scenario asked for. Three kinds of evidence are now distinguished, and each of
the six gets the one it actually has:

* **Declaration**, for all six — read from the persisted project's own
  `CodeModule` text through `Get-Phase5ProjectProcedureNames`, which strips
  comments and string literals with the same `Get-VbaExecutableCode` that P5-EV
  uses. A procedure named only in a comment is not declared, and a manifest that
  names it is not the project that holds it.
* **Callability**, for the five read-only procedures — and callable means
  `Excel.Run` returned. The single `is callable` label in P5-M sits downstream
  of `$probe = $Excel.Run($name)`.
* **Execution**, for `PCCM_Calculate` — deferred, in a Note that says where it
  goes: P5-FIX is the first valid-fixture execution and P5-AN drives the corpus.

`test_194` proves the branch sets no flag and emits no check; `test_195` proves
the declaration reader reads the project and not the manifest, and that no check
in P5-M has a constant condition — the other shape a hollowed-out check takes.

### Mutation controls

| | Mutation | Detectors |
|---|---|---|
| M1 | `A1 has proved the production project compiles` restored | 2 |  <!-- retired-authority: quoted to record what was removed -->
| M2 | `A1 remains the first real VBA compilation boundary` restored | 2 |  <!-- retired-authority: quoted to record what was removed -->
| M3 | the P5-D0 title back to `AFTER the A1 production compile` | 2 |  <!-- retired-authority: quoted to record what was removed -->
| M4 | `PCCM_Calculate` marked callable without being called | 2 |
| M5 | the persisted-project declaration check removed | 2 |
| M6 | one of the five callable APIs loses its `Excel.Run` | 2 |

### Scope

Production VBA is **byte-identical to ae52bdd**: the fifteen renames were not
reopened. `P5-CMP`'s mechanism is unchanged. No builder, spec, oracle or corpus
change.

### Status

Gate B is **not** accepted. Phase 5 is **not** accepted. No runtime execution is
requested yet.

**NO WINDOWS/EXCEL RUNTIME WAS EXECUTED DURING THIS CORRECTION ROUND.**

## Review of d21e1d7: which VBProject did command 578 compile?

Independent review accepted the ae52bdd corrections and rejected d21e1d7 on one
material blocker plus two evidence-wording overclaims. All three were real.

### Blocker 1 — P5-CMP did not prove which project it compiled

Command 578 is a **VBE** command and it acts on the VBE's **active** project.
`P5-CMP` read `Enabled` and called `Execute` without ever reading
`$vbe.ActiveVBProject` or binding it to `$Workbook.VBProject`. So it proved at
most *"the VBE's active project compiled"*, and a fresh owned Excel instance can
still carry an add-in, a startup workbook or `PERSONAL.XLSB`, each with its own
`VBProject`. Gate B may not assume the right one is active.

**The identity gate.** Both projects are now read, and compared, before the
command is touched at all:

```powershell
$targetProject = $Workbook.VBProject      # checked: must be readable
$activeProject = $vbe.ActiveVBProject     # checked: must be readable

$targetFull = [System.IO.Path]::GetFullPath($targetFile)
$activeFull = [System.IO.Path]::GetFullPath($activeFile)
$sameFile   = [string]::Equals($targetFull, $activeFull,
                  [System.StringComparison]::OrdinalIgnoreCase)
$targetIsActive = $haveFiles -and $sameFile
```

**The identity is the file path.** `VBAProject` is the default name every
project gets, so two unrelated projects routinely share it; a name comparison
would report identity between them. The Stage-B workbook is saved to a concrete
`.xlsm` path before this runs, so `FileName` is available and is what
distinguishes them. `GetFullPath` normalises separators and relative segments,
and the comparison is case-insensitive because NTFS paths are. No caption is
read, and both sides must actually name a file — two empty `FileName`s are not
an identity, which is what `$haveFiles` exists to catch.

The two project **names** are recorded as diagnostic context and are never the
test.

**Fail closed rather than activate.** A mismatch does not compile something
else and report PASS. It does not activate the target project either: doing that
through the VBIDE model is possible, but it is UI manipulation this round has no
runtime evidence for, and compiling the wrong project is the failure being
corrected. So a mismatch records a Note naming both projects, leaves
`$targetIsActive` false — which is itself an `Add-Check`, so the checklist that
decides the result fails — and never reaches `FindControl`, `Enabled` or
`Execute`.

That last point matters: a **disabled** Compile command is not evidence either.
"Nothing left to compile" over somebody else's project says nothing about this
one, so the `Enabled` reads live on the same guarded branch as the `Execute`.

**COM lifecycle.** Five transients — `$control`, `$bars`, `$activeProject`,
`$targetProject`, `$vbe` — released in one `finally`, each under its own label,
each cleared to `$null`. No `VBProject` RCW outlives `P5-CMP`.

`P5-CMP`'s claim is now *the PCCM production VBProject compiled*, not *some
active VBProject compiled*.

### Blocker 2 — the stale "six API procedures callable" history  <!-- retired-authority: quoted to record what was removed -->

The Run-7 history block still read `P5-M PASS six API procedures callable`. Run 7  <!-- retired-authority: quoted to record what was removed -->
did print those six PASS lines, but the review of ae52bdd established that one of
them — `PCCM_Calculate` — had never crossed `Application.Run`. The history now
records what the harness **reported** and names the claim that was borrowed,
here and in the VBA-source suite, the harness commentary and this document.

### Blocker 3 — `test_14` claimed six exercises from name presence

`test_14_all_six_api_procedures_are_exercised` searched the scenario source for  <!-- retired-authority: quoted to record what was removed -->
each API name as a string literal. A name in source is not an exercise. It is now
`test_14_each_api_procedure_has_the_evidence_the_hierarchy_gives_it`, and it
asserts the hierarchy: **six declared** (P5-M, against the persisted project),
**five callable** (P5-M, downstream of `Excel.Run`), **`PCCM_Calculate` executed
first in P5-FIX** and again across the corpus in P5-AN.

### Mutation controls

| | Mutation | Detectors |
|---|---|---|
| M1 | the `ActiveVBProject` read deleted | 2 |
| M2 | the identity check forced `$true` | 2 |
| M3 | `Name` compared instead of `FileName` | 2 |
| M4 | `Execute` hoisted above the identity proof | 2 |
| M5 | `P5-M PASS six API procedures callable` restored | 2 |  <!-- retired-authority: quoted to record what was removed -->
| M6 | `test_14` back to the retired name | 2 |

Three of these survived or were weak on the first pass, and two of the three
exposed defects in the regressions rather than in the harness: the identity test
checked that `GetFullPath` appeared *near* the comparison instead of pinning the
comparison's **operands**, and the stale-phrase sweep compared a phrase
containing `API` against an already-lowercased line, so it matched nothing at
all. That second one is the same shape as the defect the sweep exists to catch —
a check that reports clean because it is looking at the wrong thing — and it is
recorded here rather than quietly fixed.

### Scope

Production VBA is **byte-identical to ae52bdd**: all 13 modules verified by
digest. The fifteen reserved-identifier renames, `Contribute`, the Run-6
enumeration correction, P5-M's six-declared/five-callable split, P5-FIX's
calculation ownership and command ID 578 itself are all untouched apart from
binding the command to its target project.

### Status

Gate B is **not** accepted. Phase 5 is **not** accepted. No runtime execution is
requested yet.

**NO WINDOWS/EXCEL RUNTIME WAS EXECUTED DURING THIS CORRECTION ROUND.**
