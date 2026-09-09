# PCCM — Phase 7 closure settlement

```text
PHASE 7 — ACCEPTED / CLOSED
No outstanding Phase-7 runtime blocker.
```

Phase 7 added the stochastic **annual** product, the **selected-Px profile**, the
**reporting selector** semantics, the **A/B publication** discipline and
**sensitivity**. It now has its own Windows/Excel runtime evidence for every one
of those, produced by eight dedicated minimal runners executed on the operator's
machine against a real Excel and a real Stage-B workbook.

This record settles what was proved, by which evidence, and — just as
deliberately — what was **not** proved and what remains historical.

---

## 1. Two authorities, and they are not the same commit

A settlement that quoted one commit would let a reader believe the last commit
on the branch shipped production code. It did not.

| Authority | Commit | What it is |
|---|---|---|
| **Production / spec implementation** | `79d4c3e` | the last commit that changed a byte under `pccm/src` or `pccm/spec` **at the point the Windows evidence was produced**. Every executable VBA module and every contract Phase 7 was accepted on is rooted here. |
| **Windows acceptance evidence** | `ad78988` | the branch HEAD the accepted runtime evidence was produced against. It introduced **no** production VBA and **no** spec change. |

Established from the repository, not asserted:

```text
$ git log --oneline -1 ad78988 -- pccm/src pccm/spec
79d4c3e P7-6 correction: a whole ladder of quantiles, sorting exactly once

$ git diff --stat 79d4c3e ad78988 -- pccm/src pccm/spec
(no output)
```

The query is bounded at the acceptance head on purpose. Later phases move on:
Phase 8 changes `spec/workbook.yaml`, which is the workbook's presentation
layout, on the output sheet. That does not move Phase 7's baseline, and an
unbounded query would make it look as though it had. What must never change
after `ad78988` is any MODULE these scenarios were run against. That is narrower
than "`pccm/src` never changes", and deliberately: a later phase may add a
module - P8-1 adds the Results state adapter - without touching a byte any
Windows run executed. A control asserts that no file under `pccm/src` is
modified or removed after `ad78988`; additions are allowed and are visible in
the diff, and a **declared later correction** — §1.1 — is allowed and must be
named there.

Every one of the fourteen commits between the two touches only
`pccm/bootstrap/windows`, `pccm/builder` and `pccm/tests` — Windows runners,
acceptance-corpus fixture and projection metadata, and their static controls.

### 1.1 Later corrections to shared owners — declared, not denied

A later phase can find a real defect in a module Phase 7 was accepted against.
Pretending otherwise would leave two bad options: leave the defect, or change
the bytes and let this record go on claiming `79d4c3e` was the last commit that
touched them. Neither is acceptable, so the record carries a third: **the
modification is permitted and it is named here.**

This does **not** reopen Phase 7 and it does **not** move either authority.
`79d4c3e` remains the implementation baseline the eight scenarios were built
from; `ad78988` remains the evidence head they were produced against. What a row
below says is that a byte those scenarios executed has changed **since** that
evidence was produced, which commit changed it, and why — so that any later
reading of the Phase-7 evidence knows exactly what is no longer identical to the
tree it was produced on.

A row names the correction by its **commit subject**, not by a hash. The hash of
the commit that lands a row cannot be written inside that row, and a placeholder
that a later commit fills in is a claim nobody checks. The subject is unique,
and it resolves:

```text
$ git log --format='%h %s' --grep='the Results state path is read-only' -1
$ git log --format='%h %s' --grep='a risk publishes its risk name' -1
$ git log --format='%h %s' --grep='the live calculation state is read-only' -1
$ git log --format='%h %s' --grep='Phase 9 Step 2B: the refusal names its driver, structurally' -1
$ git log --format='%h %s' --grep='Phase 9 Step 3: the last two owners name their driver too' -1
```

| Module | Correction | Found by | What changed, and why |
|---|---|---|---|
| `pccm/src/vba/modSimReport.bas` | P8-1: the Results state path is read-only | P8-1, first complete Windows run | Adds `SimReportDerivedStatus`, a public delegation returning the existing private `DeriveSimStatus()`. **No derivation, no persistence and no existing procedure changed.** `PCCM_SimulationStatus` still derives *and* writes `_SimData!D28:D29`, exactly as Phase 7 accepted it. |
| `pccm/src/vba/modSimAnnualStore.bas` | P8-1: the Results state path is read-only | P8-1, first complete Windows run | Splits the annual precondition into a command path (`SimAnnualStoreCurrentRun`, unchanged behaviour, still asks the persisting entry point, still the one `modSimAnnualRun` uses) and a read path (`SimAnnualStoreCurrentRunReadOnly`), both settling through one private `CurrentRunFor`. The two state accessors take the read path. **No state rule moved and no refusal text changed.** |
| `pccm/src/vba/modSimPostReport.bas` | P8-3 Windows run 4: a risk publishes its risk name | P8-3, fourth Windows run | `DriverNameOf` read `COL_RISK_REGISTER_DESCRIPTION` for a risk — the column `driver_contract.yaml` declares `required: false`, an optional note — and so published no name for a risk that had none. It now reads `COL_RISK_REGISTER_RISK_NAME`, which that contract declares `required: true`. **One line.** The cost-line branch is untouched and was always right: a cost line has no name column, and its description is its `required: true` label. No driver id, rank, signed rho, absolute rho, direction, status, ordering, fingerprint or replay mathematics is touched. |
| `pccm/src/vba/modCalcReport.bas` | P9-2: the live calculation state is read-only | P9-2, the Model Check surface | Adds `CalcReportDerivedStatus`, a public delegation returning the existing private `DeriveStatus()` over the existing private `PrepareCurrentCalculation()`. **No derivation, no persistence and no existing procedure changed.** `PCCM_CalculationStatus` still derives *and* writes `_Calc!C19:C20`, exactly as Phase 5 accepted it and Phase 7 ran it. The reason is the P8-1 reason one module along: Excel forbids a function a worksheet cell called to change the workbook, so a Model Check cell reaching the persisting entry point would show `#VALUE!`. The derivation was already pure and the persistence is the caller's, so this exposes the pure half - one derivation, two entry points, differing only in whether `C19:C20` is rewritten afterwards. |
| `pccm/src/vba/modCalcReport.bas` | Phase 9 Step 2B: the refusal names its driver, structurally | P9-2B, the Model Check refusal subject | The private `PrepareCurrentCalculation()` gains a `ByRef subject As String` beside the `ByRef detail As String` P9-2A gave it, and the public `CalcReportDerivedStatus` passes both through. **This is the one row in this table that changes the text of a procedure Phase 7 executed**, and it changes only its signature, the `Dim` lines the extra local sits on, the two owner calls it is handed to, and the entry clear that now clears both. Not one condition, message, Boolean, constant or arithmetic expression moved — proved mechanically, not asserted: `tests/vba_subject_plumbing.py` removes the plumbing again and the accepted reporter prefix comes back byte for byte, so the frozen digests the Phase-5 and Phase-6 controls take never moved at all. |
| `pccm/src/vba/modCalcResolve.bas` | Phase 9 Step 2B: the refusal names its driver, structurally | P9-2B, the Model Check refusal subject | `ResolveModel`, `ResolveDrivers`, `ReadRegister`, `ReadDriverRow`, `ResolveProfileWeights` and `AttachDriverFx` each gain a `ByRef subject As String`, and the permanent id each of them ALREADY holds when it refuses about one driver is assigned to it — set where the id becomes known, cleared again on success so a later model-wide refusal cannot inherit it, and left blank wherever no single driver is at fault. **Every refusal condition, every Boolean and every refusal sentence is untouched**, which the same mechanical reversal proves by restoring the `ad78988` bytes exactly. |
| `pccm/src/vba/modCalcCheck.bas` | Phase 9 Step 2B: the refusal names its driver, structurally | P9-2B, the Model Check refusal subject | `CheckResolvedModel` gains the same `ByRef subject As String`, and the per-driver loop assigns the id it already builds its sentence from, so an ordering, `Quantity`, `Probability` or profiling-sum refusal names its cost line or risk while a model-level refusal names none. **One added parameter, one assignment in the loop, two clears — no validation predicate, no ordering rule and no message text.** The reversal restores the `ad78988` bytes exactly, which is why no historical digest in the Phase-5 or Phase-6 controls moved. |
| `pccm/src/vba/modCalcAnalytical.bas` | Phase 9 Step 3: the last two owners name their driver too | P9-3, the last static contract gap | `AccumulateTotals`, `BuildAnnualSeries` and `Reconcile` each gain a `ByRef subject As String` and assign the permanent id **their own loop already holds** when they refuse about one driver — a conditioning magnitude at either contribution pass, a per-driver-per-year annual magnitude, an I5 profile sum — and clear it again where the model-wide phase begins, so a measure total, a coefficient mismatch or an annual-series failure still names nobody. **No identity, tolerance, conditioning rule, arithmetic expression or refusal sentence is touched.** The module was one raw line under its accepted ceiling and did not move off it: every assignment rides on a statement that was already there, so the correction is NET ZERO lines. The mechanical reversal in `tests/vba_subject_plumbing.py` restores the `ad78988` bytes exactly, which is why the pre-Run-7 digest did not move. |
| `pccm/src/vba/modCalcReport.bas` | Phase 9 Step 3: the last two owners name their driver too | P9-3, the last static contract gap | `BuildDriverFactors`, `BuildAudits`, `BuildAnnual` and `BuildFingerprint` gain the same parameter and carry it to the four remaining driver-specific families — an inflation profile outside the resolved reference set, a Knom or Kpv factor build, a driver-audit build, a fingerprint-record encoding. Each assignment is the id the enclosing loop already had; each is cleared where the model-wide phase begins, so fingerprint construction and the reconciliation identities still name nobody. **Signatures, call sites and those assignments — nothing else**, and again NET ZERO raw lines. The accepted reporter prefix still hashes to `8d67d3f1…` once the plumbing is reversed. |

**The defect these rows exist for.** `PCCM_AnnualDistributionState` and
`PCCM_AnnualProfileState` reached `PCCM_SimulationStatus`, which persists the
derived status pair. Excel forbids that to a function a worksheet cell called,
so from the moment a publication existed both Results state cells showed
`#VALUE!`. Phase 7 never saw it: its scenarios call the accessors through
`Application.Run`, out of cell, where the write is permitted. It is a Phase-8
integration defect in a Phase-7-owned module, and that is precisely the case
§1.1 was written for.

**What the runtime evidence still stands on.** Every W-scenario invoked these
accessors out of cell, where both paths behave identically — the read path
returns the same string from the same derivation. The one observable difference
is that `_SimData!D28:D29` is no longer rewritten as a side effect of asking,
and no accepted scenario asserted that it was.

---

## 2. The eight accepted Windows scenarios

Every runner is a **dedicated minimal file**. The large
`phase7_acceptance_scenarios.ps1` was abandoned for runtime acceptance after
Run 1 and is frozen as history at SHA-256
`9744d9b7c1b4ebbc94ae48db54dd2ffb74af9a554ee43d8da8b1e06efe68c8ed`; no accepted
scenario runs it, and every W-suite pins that digest.

| # | Scenario | Final commit | Runner (lines) | SHA-256 (first 16) | Reported result |
|---|---|---|---|---|---|
| W1 | compile / public API smoke | `164113f` | `phase7_w1_smoke.ps1` (757) | `bf41a2193dbf2ea5` | 59 checks / 0 failed |
| W2 | 300 drivers × 5 years | `fb1e96c` | `phase7_w2_many_drivers.ps1` (978) | `558bcfa0528b38c2`* | 7 prereq + 21 checks / 0 failed |
| W3 | 10 drivers × 200 years (2001–2200) | `cc78015`† | `phase7_w3_long_years.ps1` (1096) | `97af69bf836a2d79` | 7 prereq + 36 checks / 0 failed |
| W4 | no-run refusal, then FIXED baseline | `3874496` | `phase7_w4_base_simulation.ps1` (1281) | `9e1be0629d337b42` | 9 prereq + 55 checks / 0 failed |
| W5 | first successful annual run | `df1e34d` | `phase7_w5_annual_success.ps1` (1511) | `de2e3322c43e93da` | 16 prereq + 33 checks / 0 failed |
| W6 | reporting selector P80 → P50 | `e504b7b` | `phase7_w6_selector_move.ps1` (1601) | `feefceb0b7b0592a` | 16 prereq + 66 checks / 0 failed |
| W7 | A→B→A bank cycle, 20 → 4 year shrink | `7de7b01` | `phase7_w7_bank_cycle.ps1` (1658) | `6abcf0874909c21f` | 12 prereq + 60 checks / 0 failed |
| W8 | STALE and INVALID annual refusals | `ad78988` | `phase7_w8_refusal.ps1` (1896) | `cf4adee6061af837` | 17 prereq + 75 checks / 0 failed |

\* the full digests are pinned as literals inside the W-suites; the short forms
above are for reading, and the suites are the authority.

† W3's runner landed in `ba0949e`; `cc78015` is the **fixture correction** that
made the scenario accepted, and it touched the acceptance corpus and its
controls rather than the runner. The column records the commit each scenario was
accepted at, which is not always the commit that wrote its runner.

**Where these results come from.** Every Windows result in that table was
produced by the operator on a Windows host and reported back. This repository
holds the runners, their static controls and their design — it does not hold a
transcript of the runs, so the check counts are the operator's report, accepted
as such. Nothing in this record re-derives, re-runs or infers a Windows number.

**What each scenario established at runtime**

- **W1** — the whole VBAProject compiled in real Excel and every required public
  procedure resolved. Before any production, the annual accessors returned
  NOT PRODUCED / NOT PRODUCED / blank Px / year count 0.
- **W2** — 300 drivers over 5 years: every driver mapping, every annual value
  and all ten totals matched the independent Phase-5 oracle.
- **W3** — 10 drivers over exactly 200 years, timeline 2001–2200, project index
  exactly 1…200; drivers, years, inflation, annual and totals all matched.
- **W4** — the annual endpoint refused before any simulation existed and left
  the empty state untouched; then one FIXED-seed run (seed 20260905, 1,000
  iterations) published run_id 1 in bank A with no AUTO nonce consumed, and its
  deterministic base matched the Phase-5 oracle.
- **W5** — the first successful `PCCM_RunAnnualStochastic`. Four compact annual
  records, both percentile ladders complete, numeric and ordered, and the
  selected-Px profiles reconciled to the total: **nominal identity delta 0**,
  **PV identity delta 4.54747350886464E-13**. Contingency matched
  `selected_px_total − deterministic_base_estimate_a`. Simulation identity and
  iteration data unchanged by the annual pass.
- **W6** — the selector moved P80 → P50 with **no** re-simulation. The run
  identity did not move, the distribution ladders stayed value-identical and
  CURRENT, and the persisted P80 profile survived as OTHER Px. An annual-only
  rerun published P50, reconciled at P50.
- **W7** — A1 (20 years, bank A, run 1) → B (bank B, run 2, distinct FIXED seed
  and fingerprint) → A2 (4 years, bank A, run 3). Former rows 5…20 were
  completely cleared and non-authoritative; each inactive bank survived the
  other's publication value-identically; all three results reconciled.
- **W8** — from a reconciled baseline, iterations 1000 → 1001 through the
  projected input made the simulation STALE while the calculation stayed
  CURRENT; a cost-line maximum below its own minimum made the model INVALID
  through the ordinary register. The annual endpoint refused in both states and
  the published answer survived both refusals physically intact, with no
  publication appearing in the other bank. REFUSED was shown to be an attempt
  outcome and never a persistent state.

---

## 3. Static and source evidence

### 3.1 Re-established from this branch today

| Measurement | Value | How |
|---|---|---|
| Test suites in `pccm/tests` | **72 files** at the closure commit `b844915` | file count |
| Tests collected | **4,663** at the closure commit `b844915` | `pytest tests --collect-only` |
| Phase-7 suites / tests | **23 suites, 1029 tests** | `pytest tests/test_phase7*.py --collect-only` |
| W1–W8 static controls | **307 tests** | the eight W-suites |
| Stage A | **351 passed / 0 failed** | `python3 builder/build_stage_a.py` on the clean tree |
| Focused Gate-B + Phase-7 suites | **488 passed / 0 failed** | post-commit, clean tree |
| Broad sweep during W8 preparation | **4,638 passed** | pre-commit tree; the 5 remaining were the dirty-tree artefact family, cleared by committing and rebuilding Stage A, and re-run green inside the 488 |

**An honest note on the totals, because they do not line up neatly.** The broad
sweep collected 4,643 and reported 4,638 passing; the five that did not were the
dirty-tree artefact family, cleared by committing and rebuilding Stage A, and
re-run green afterwards inside the focused 488. So all 4,643 have been observed
passing at that tree — **across two runs rather than one**. The tree now
collects 4,731: this settlement added its own controls, and Phase 8 has added
suites after it. Those have been run focused and have never been part of a full
sweep, so no single sweep has covered all of them. One command would settle it
and it has not been run.

THE FIRST TWO ROWS ARE A SNAPSHOT, AND SAY SO. They describe the tree at the
closure commit, because a tree-wide count moves every time any later phase adds
a suite - and it has moved twice since, for reasons that have nothing to do with
Phase 7. Re-deriving them at HEAD would make this record need an edit whenever
somebody wrote a test, which is a record that says less each time it is touched.
The counts below them are Phase-7-scoped and closed, so those are still
re-derived from the tree on every run. The **evidence** all of them describe -
the eight scenarios, the two authorities, the harness classification - does not
move at all.

### 3.2 Historical — reported, not re-establishable from committed evidence

The repository records none of the following numbers. They are carried here as
the operator's report of earlier Phase-7 rounds and are **not** confirmed by
this branch:

| Round | Reported |
|---|---|
| P7-2 | 51 suites, 3,741 passed / 0 failed |
| P7-3 | 53 suites, 3,780 passed / 0 failed |
| P7-4 | sensitivity persistence and orchestration completed; 20 drivers × 10k ≈ 6.4–7.6 s, 100 × 10k ≈ 31.4 s, 300 × 10k ≈ 105.4 s; no subsampling or cap required |
| post-fix broad sweep | 57 suites, 3,909 passed / 0 failed / 0 errors |
| P7-5 | annual replay and reconciliation; worst nominal/PV per-iteration annual-sum delta 3.638e-12; selected-profile sum identities within Double round-off; no scaling; quantile-ladder sort optimisation |
| post-optimisation sweep | 4,290 / 0 |

Two of these are partly corroborated by the repository, and the corroboration is
worth stating precisely:

- The P7-4 **scenario shape** is in the tree. `phase7_timing_scenarios.ps1`
  defines exactly A = 20 drivers × 10,000, B = 100 × 10,000, C = 300 × 10,000
  and invokes `PCCM_RunSensitivity`. The driver and iteration counts are
  therefore re-establishable; **the elapsed seconds are not**.
- The P7-5 replay authority is in the tree as
  `tests/test_phase7_sim_annual_replay_vba.py` and passes at HEAD; the worst
  delta it produced in that round is not recorded anywhere, so 3.638e-12 stays
  historical.

---

## 4. Harness corrections — classified

Four defects were found during Phase-7 Windows acceptance. **None of them was a
PCCM production defect, and none changed production VBA or the simulation or
annual contract.**

| # | Defect | Class | Resolution |
|---|---|---|---|
| 1 | the large acceptance harness: PowerShell 5.1 `Join-Path` incompatibility, a missing transitive helper, and incorrect public-procedure detection with the lifecycle problems behind it | harness architecture | abandoned for runtime acceptance; replaced by the dedicated minimal W1–W8 runners. The file is frozen as history and never patched again. |
| 2 | the W3 fixture asked for a span extending past the structural maximum year 2200 | **fixture** defect; production validation was correct to refuse | the start year is now derived from the structural contract's own boundary; the span became 2001–2200 |
| 3 | W5 first run: a `[string]` cast applied to a whole `char[]` turned column `AD` into `"A D"` | PowerShell helper defect | corrected and boundary-tested at Z→AA, AD, AN/AO, AZ/BA; every annual reader audited for the same helper |
| 4 | W5 second run: the runner read the contingency block as the total percentile block | harness / projection **semantic** defect | the contract already owned the distinction; it is now projected as `summary_semantics` and consumed rather than guessed. The rerun reconciled. |

Defects 2 and 4 share a lesson the project has now applied three times: **a
semantic nothing projects is a semantic the harness has to guess.** `selector_semantics`,
`publication_semantics` and `model_states` were each projected *before* the
scenario that depended on them, rather than after a failed run.

---

## 5. The contracts Phase 7 now has runtime evidence for

### Annual stochastic — W4, W5, W6, W7, W8
- annual nominal and PV distributions produced in real Excel
- replay tied to the accepted simulation identity, and the stamp bound to it
  field by field
- Type-7 ladders, complete, numeric and ordered
- selected-Px profile tied to the exact total-percentile source identity
- profile sum = total Px, under the accepted identity rule, **with no scaling**
- compact persistence only — records scale with project years, and no N×Y matrix
  is materialised in the contracted annual region
- the four Phase-8 handoff accessors report correctly

### Selector — W6
- the selector is reporting-only: moving it does not stale the simulation
- the distribution stays CURRENT and value-identical
- the previously published profile becomes OTHER Px and is never relabelled
- an annual-only rerun publishes the new profile

### Banking and persistence — W4, W7
- A/B bank isolation, and the projected `"" → A → B → A` cycle
- publication identity binding, and marker-plus-year-count authority — never the
  last non-blank row
- surplus annual rows cleared on a duration shrink, across the whole former span
- the inactive bank preserved value-identically, iteration block included

### Refusal and state — W4, W8
- the no-run annual refusal, the STALE refusal and the INVALID refusal
- historical payload preservation across every refusal
- persistent state kept distinct from attempt result: REFUSED is an attempt
  outcome and appears in no persistent vocabulary
- no refusal consumes a run identity, an AUTO nonce or the pending marker

### Sensitivity — and this one is not W1–W8
**W1 through W8 deliberately ran no sensitivity**; every one of their
authorisations excluded it, and no W-runner invokes `PCCM_RunSensitivity`. The
one thing they do establish is that the entry point is there: W1's public-surface
matrix required `PCCM_RunSensitivity` to exist and resolve in the compiled
VBAProject, and it did. Existing is not running. The sensitivity contract items — one driver per cost
line and risk, full-N Spearman, mid-rank ties, zero-variance exclusion, ranking
by |rho| with the signed rho retained, deterministic ID tie-break, the full
population persisted, display capacity not being a driver ceiling, and
performance accepted without subsampling — rest on:

- the **P7-3 / P7-4 static suites**, which are in this tree and pass
  (`test_phase7_sim_sensitivity.py`, `test_phase7_sim_sensitivity_validation.py`,
  `test_phase7_sim_postreport.py` and their validation pairs); and
- the **P7-4 Windows timing round**, whose runner and scenario shape are in the
  tree but whose measured timings are historical (§3.2).

That is real evidence and it is enough to close Phase 7, but it is a different
provenance from W1–W8 and this record does not blur them.

---

## 6. Known limitations, stated rather than buried

1. **There is no independent annual Windows oracle, and none is claimed.** The
   annual answer is reconciled against the contract's own Type-7 definition over
   the published iteration column and against the accepted identity rule. That
   is a cross-check, not a second implementation. The independent Python oracle
   from P7-5/P7-6 proves the computation statically; W5–W8 prove live execution,
   persistence, read-back and reconciliation.
2. **Windows results are operator-reported.** This repository holds no run
   transcript. Every check count in §2 is accepted on the operator's report.
3. **The historical numbers in §3.2 are not re-establishable here.**
4. **Sensitivity has no W-scenario** (§5).
5. **No single sweep has covered all 4,643 tests on the final tree** (§3.1).
6. The **Phase-6 runtime authority remains historical evidence only.** It is not
   superseded as a record of Phase 6; it simply no longer stands in for Phase-7
   functionality, which now has runtime authority of its own.
7. **Two Phase-7 modules have been modified since `ad78988`** by a declared
   Phase-8 integration correction (§1.1). The accepted Windows evidence was
   produced before that change, and this record does not claim it was re-run
   afterwards.

---

## 7. Settlement

- **Phase 7 is ACCEPTED / CLOSED**, with no outstanding runtime blocker.
- **No further Phase-7 Windows scenario is required.** The eight accepted
  scenarios cover compile and public surface, deterministic scale in both
  dimensions, the refusal floor, the first successful annual run, selector
  semantics, bank cycling with a duration shrink, and both live refusing states.
- The **historical Phase-6 runtime authority remains historical only**.
- **Phase-7 runtime authority supersedes it for Phase-7 functionality.**
- **Phase 8 may begin** from the accepted Phase-7 implementation baseline
  `79d4c3e` and the evidence baseline `ad78988`.

**This is not project completion.** Phase 8, Phase 9 and Phase 10 remain, along
with final user documentation, final delivery validation and independent
red-team review.

---

## 8. Next step

**Phase 8** — presentation, from the accepted handoff surface Phase 7 leaves
behind: `PCCM_AnnualDistributionState`, `PCCM_AnnualProfileState`,
`PCCM_AnnualProfilePx` and `PCCM_AnnualYearCount`, over a persisted annual
result whose authority is the publication marker plus the stamped year count.
Phase 8 is not started by this record.
