# Phase 19 — Training Phases

**Closed:** 2026-07-09 · **Milestone:** M4 (Intelligence) — closing phase

The governing principle this phase, held above everything: **training phases are context, not
prediction.** A phase *explains* why a stretch of history looked the way it did — it never
biases or rewrites analytics. Historical results are always read against the phase that was
active at that time, so editing today's phase can never reinterpret a past block.

## Context, not prediction (the standing principle)

- **A phase is a lens over a fixed window, not a filter at query time.** The Phase Report
  computes strictly inside `[startDate, endDate]` (end = today for an ongoing phase). A
  completed block reads **identically regardless of what today's phase is** — a property
  pinned by a test that runs the same completed phase with two very different `today`s and
  asserts byte-equal reports.
- **Phases never touch the rolling-range views.** The Body / Training / Nutrition sections
  and every insight stay phase-agnostic; declaring or ending a phase changes none of them.
  Phase analytics is a peer long-term view (§5.4/§4), not a global mode.
- **Never auto-detected.** Phases are user-declared, may not overlap, and at most one is
  ongoing — both invariants enforced in `phaseRepository` inside a transaction, not guessed
  from the data.
- **Honesty over judgement.** A block shorter than 14 days or with < 2 in-phase measurements
  returns `insufficient-data` for body deltas (§5.4/§3.1); a custom or data-thin block reads
  "not enough data to judge," never a fabricated verdict.

## What was built

- **`phases` table** (migration 0010) + **`phaseRepository`**: CRUD with the no-overlap and
  single-ongoing invariants (open-ended phases are modelled with a far-future sentinel so two
  ongoing phases always collide), `getOngoingPhase`, `getPhaseForDate` (the phase active on a
  date), and an end-yesterday `endPhase`. Overlaps throw a typed `PhaseValidationError` the UI
  turns into a guided fix.
- **`PhaseAnalyticsCalculator`** (`domain/analytics/phaseAnalytics.ts`, §5.4, pure): windows
  the full-history inputs to the phase and produces a **Phase Report** — body deltas (first vs
  last in-phase snapshot via the §5.4 `compareSnapshots`, gated by the minimums), a training
  summary (workouts, weekly consistency vs target, total + by-group volume, key-exercise e1RM
  verdicts, **in-phase PRs** judged against all prior history), a nutrition summary (adherence
  over logged days) with an **intent verdict** (a cut running a surplus or gaining weight reads
  "counter"; an aligned block "on track"; a thin/custom block "unclear"), and **per-week rates**
  so different-length blocks compare fairly. It composes the Phase 17/18 calculators windowed to
  the block — a phase is aggregation over proven pieces, not new math.
- **Backup extended**: `data.json` gains `phases`, `schemaVersion` 10 → 11 with a v10→v11
  data-shape upgrader (old archives default `phases: []`), and the round-trip suite now carries
  a phase row end-to-end (export → wipe → import → byte-equal).
- **UI**: a **Phases** management screen (ongoing block with end/edit, history list, declare
  form with type chips + ISO dates so a block can be declared over historical data), the
  **Phase Report** screen (a coach's read — intent verdict first, then body change, training
  quality, nutrition), a **current-phase progress card** on the Analytics home, a Settings
  entry point, and a modest **phase-complete moment** (a celebratory dialog into the report on
  ending a block).

## What changed

New: `domain/models/phase`; `data/schema` migration 0010 + `phases` table; `phaseRepository`
(+ tests); `domain/analytics/phaseAnalytics` (+ tests); `features/analytics/{hooks/usePhases,
components/PhaseFormSheet, components/PhaseReportView, components/CurrentPhaseCard,
screens/PhasesScreen, screens/PhaseReportScreen}`; `app/(tabs)/analytics/{phases,phase/[id]}`.
Modified: backup (`backupSchema`/`collect`/`replace`/`upgraders`/`schemaVersion`/testkit),
`changeBus` (`phases` table), `AnalyticsScreen` (current-phase card), `SettingsScreen` (entry).
No frozen document changed — the `phases` table and its backup format were already in the v1
baseline (DATABASE §3.7, §6); this phase implements them.

## Screens affected

Analytics home (current-phase card), Phases management (new), Phase Report (new), Settings
(phases entry).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| No-overlap + single-ongoing enforced (closed, adjacent-ok, double-ongoing, run-into) | repo test | ✅ |
| End-yesterday frees the ongoing slot; edit re-checks overlap; end-before-start rejected | repo test | ✅ |
| `getPhaseForDate` resolves the active phase, null in a gap | repo test | ✅ |
| In-phase workouts/volume/by-group; PRs vs prior history; weekly consistency | domain test | ✅ |
| Body deltas = first-vs-last in phase; §5.4 minimums gate (span/snapshots) | domain test | ✅ |
| Intent verdict: cut-down aligned, bulk-that-lost counter, bulk-in-deficit counter, custom unclear | domain test | ✅ |
| **Completed phase reads identically regardless of `today`** (context, not prediction) | domain test | ✅ |
| Backup round-trip includes phases; schemaVersion 11 matches the migration journal | backup tests | ✅ |
| `npm run check` | typecheck + lint + format + 355 tests + db:check (16 tables) | ✅ green (3× stable) |
| On-device phase walk (declare/overlap-guard/end, historical report reconciles vs Compare + history, intent honesty, complete moment, two-phase compare, edit-doesn't-rewrite-history, backup round-trip, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **The phase UI is device-only** — the management screen, report, current-phase card,
   declare form and complete moment fold into TD-001. What *is* verified off-device is the
   whole engine: the repository invariants and the calculator (deltas, training/nutrition
   summaries, intent, per-week rates), reconciled to a hand-computed fixture.
2. **Dates are entered as `YYYY-MM-DD` text**, not a native date picker — functional and
   honest (a block can be declared over any historical range), but a date-picker polish pass
   belongs to the Phase 21 feel pass. No new debt is registered; it's a deliberate,
   documented simplicity.
3. **Historical phase analytics use current settings** (weekly target, default bodyweight)
   where the schedule that was active then isn't stored — the same accepted approximation as
   the range views; volume and body deltas, the load-bearing numbers, are exact per date.
4. **`usePhaseReport` fetches full history** and recomputes on any source change. Acceptable at
   personal scale (§3.3); the obvious first optimization only if the report is ever felt slow.

## Technical debt

None introduced. TD-001 gains the training-phases device walk. (TD-003/005/007/008/010
unchanged.)

## Retrospective

**What went well?** Framing the phase as *a lens over a fixed window* made the whole thing
fall out cleanly: the calculator composes the already-proven Phase 17/18 calculators windowed
to `[start, end]`, so the new code is aggregation plus an intent verdict, not new analytics.
The "context, not prediction" principle became a concrete, testable property — a completed
phase must read the same no matter what today is — which is now a regression test rather than
an aspiration. The no-overlap invariant modelled with a far-future sentinel for open-ended
phases collapsed "no overlap" and "at most one ongoing" into a single check.

**What was harder than expected?** Deciding how much of the report to gate. §5.4's
insufficient-data clause is easy to over-read as "gate the whole report," but the honest
reading is to gate only the body-delta comparison (which genuinely needs two points over a
real span) while still showing the training and nutrition that happened. Getting the intent
verdict to be *specific and honest* — naming weight movement or calorie skew, and hedging to
"not enough data" instead of guessing — took the most iteration; a vague "good/bad" would have
violated the phase's own principle.

**What should change before the next phase?** Nothing structural. CP-D (the M3+M4 checkpoint)
is next: a §3 review over the full analytics engine — calculators modular and pure, honesty
rules everywhere, the P1/P8 sweep of every visible metric, and a performance sanity check
against the §7 budgets. The phase engine built here is the last analytics surface it audits.

## Lessons Learned

- **What surprised you:** how much the single property "a completed block reads the same
  regardless of today" disciplined the design — it ruled out every shortcut that would have
  reached for `today` or the active phase inside a historical computation.
- **What documentation prevented mistakes:** ANALYTICS §5.4 fixed exactly what the Phase Report
  contains (body deltas, training/nutrition summaries judged against intent, rate-normalized
  comparison) and the 14-day / 2-snapshot minimums; DATABASE §3.7 fixed the `phases` schema,
  the no-overlap + single-ongoing invariants, and its place in backups; §6.3 fixed the upgrader
  discipline for the schemaVersion bump; FITNESS §5.3/§5.4 fixed the body-delta directionality
  the deltas and intent verdict rely on.
- **What should be reused:** modelling an open-ended interval with a far-future sentinel so
  overlap and single-ongoing are one check; composing a new calculator over existing windowed
  calculators rather than re-deriving; encoding a governing principle ("reads the same
  regardless of today") as an executable test.
- **What should be avoided:** letting a phase bias or filter the rolling-range views; reaching
  for `today` or the active phase inside a historical computation; fabricating an intent verdict
  when the data is thin (say "not enough data"); auto-detecting phases from the data.
- **Amendment proposals:** none — no frozen-document defect surfaced. No new debt.
