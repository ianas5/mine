# Phase 7 — Personal Records & Exercise Report

**Closed:** 2026-07-09 · **Milestone:** M1 Training

Strength history becomes meaning. Every exercise now has trustworthy personal
records — all five FITNESS_DOMAIN §3.7 types, always recomputed from history and
never cached — surfaced optimistically as you lift, tallied at finish, and laid
out on a per-exercise report. Beating a best celebrates once; a tie celebrates
nothing; deleting the record-holding workout makes the record recede.

## What was built

- **PR domain** (`domain/fitness/personalRecords.ts`, pure): `computeExerciseBests`
  (heaviest weight, best e1RM at r ≤ 12, best set volume, best session volume) and
  `detectNewPRs` (strictly-greater only — ties are never PRs, §3.7). Weight/e1RM
  use single-side `effectiveLoad` (unilateral doubling never inflates a strength
  PR); volume uses the doubled figure. A rep-at-load PR is reported **only** when
  that exact load was lifted before — a first-ever load is already a weight PR, so
  no noisy "record" for every new weight. Nothing is cached; edits/deletes make
  records recede by recomputation.
- **Exercise Report** (`domain/analytics/exerciseReport.ts`, the first
  `domain/analytics` calculator): all-time sessions, working sets, total volume,
  the four bests, average reps/working set, average effective load, and last
  performed (ANALYTICS §5.5). Trend + progression rate are the deferred time-series
  work (Phase 15) and are deliberately **absent, not faked** (P8).
- **Repository** (`workoutRepository.getExerciseSetHistory`): full working-history
  rows with per-entry unilateral counting + the exercise's name/load type — the
  raw input the domain turns into PRs and the report.
- **Reactive hooks** (`useExerciseReport`, `useExercisePrBaseline`): both recompute
  from history on every `workouts` write, so records recede live with no cache.
- **PR celebration** (delight registry): an optimistic **`PrBadge`** materializes on
  a completed set that sets a new *running* weight/e1RM best (only genuine
  record-setters light up, not every set of a first-ever exercise), with the
  `success` haptic on that ✓ instead of `light`; the finish summary shows "N new
  PRs" and the save toast reads "Workout saved · N PRs". The count is computed from
  history **before** the durable write — trustworthy, never cached.
- **Exercise Report screen** (`/workouts/exercise/[id]`): reached from the library
  exercise actions sheet ("View report"). Records / all-time / a "trend coming"
  note that explicitly says no line is shown rather than an estimated one.
- **Defect fix** (`settingsRepository.get`): the lazy single-row create is now an
  idempotent `onConflictDoNothing` + re-select, so the several bodyweight readers
  that now mount together can't race on the settings row.
- **Tests (18 new, 145 total):** PR strictness/ties, warm-up + r > 12 exclusions,
  unilateral single-side-weight vs doubled-volume, rep-at-load only-if-prior,
  first-ever-is-a-PR, recede-on-delete; report aggregation + empty; in-session
  running-best flags; and a real-SQLite suite proving session PRs vs saved
  history, the tie case, recede-on-delete, and report/​history reconciliation.

## Trustworthy by construction (the standing rule)

- **No `personal_records` table.** PRs are pure derivations of the `sets` table.
  There is nothing to go stale, nothing authoritative to cache — the source of
  truth is the logged sets, exactly as FITNESS_DOMAIN rule 7/13 demands.
- **Strictly greater, warm-ups and r > 12 excluded**, single-side load for strength
  PRs — the conservative reading, all tested against the doc's edge cases.
- **The finish count is computed from history before saving**, so "N PRs" reflects
  the same recompute the report will show afterward.
- **Recede-on-delete is free**: delete a workout → `workouts` change → every PR/​
  report hook re-queries → the record drops. Proven end-to-end.

## What changed

New: `domain/fitness/personalRecords`, `domain/analytics/{exerciseReport,index}`
(+ fitness barrel exports); `workoutRepository.getExerciseSetHistory`;
`hooks/{useExerciseReport,useExercisePrBaseline}`; `logic/{sessionPRs,
inSessionPrFlags}`; components `PrBadge`; screen `ExerciseReportScreen`; route
`app/(tabs)/workouts/exercise/[id]`. Modified: `SetRow` (PR badge + success
haptic), `ActiveExerciseCard` (per-set PR flags), `WorkoutSummarySheet` + 
`ActiveWorkoutScreen` (PRs earned + toast), `ExerciseActionsSheet` (View report),
`settingsRepository` (race-safe create). No frozen document changed.

## Screens affected

Active Workout (in-session PR badges + PRs-earned summary + toast), Exercise
Library (View report action), and the new Exercise Report screen.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| All five §3.7 PR types; strictly greater; ties are not PRs | domain test | ✅ |
| Warm-up excluded; e1RM only r ≤ 12; heavier set still a weight PR | domain test | ✅ |
| Unilateral: single-side weight, doubled volume | domain test | ✅ |
| Rep-at-load PR only when the load was lifted before | domain test | ✅ |
| First-ever performance counts as PRs (delight registry) | domain test | ✅ |
| Recede-on-delete (records recompute downward) | domain + real-SQLite tests | ✅ |
| Report aggregation (sessions/sets/volume/averages/last performed) | analytics test | ✅ |
| Report reconciles with saved history exactly | real-SQLite test | ✅ |
| Session PRs vs saved history; tie → 0 | real-SQLite test | ✅ |
| In-session running-best flags (no double-badge, warm-ups never) | logic test | ✅ |
| `npm run check` | typecheck + lint + format + 145 tests + db:check (6 tables) | ✅ green (5× stable) |
| On-device PR badge/haptic, summary/toast, report open, recede-after-delete | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** for the PR badge, summary, or report (DB-backed;
   expo-sqlite native-only). Correctness is proven by domain + repo + logic tests;
   the celebration feel + report visuals fold into the consolidated TD-001 pass.
2. **Report is all-time only** — the §5.5 selected-range metrics, e1RM trend +
   progression rate, and recent-notes are deferred to Phase 15 with the rest of the
   time-series/regression work (**TD-007**). The screen states no trend is shown
   rather than faking one (P8).
3. **PR count reflects distinct records** — one heavier top set can legitimately set
   weight + e1RM + set-volume + session-volume records at once, so "4 PRs" is
   accurate, not inflated (each is a real, independent §3.7 record).

## Technical debt introduced

- **TD-007** — Exercise Report is all-time only; §5.5's selected-range view, e1RM
  trend + progression rate, and recent-notes are deferred to Phase 15.

## Retrospective

**What went well?** "No cached PRs" made the whole phase fall out of the existing
spine: PRs and the report are pure functions of the `sets` table, so recede-on-
delete, edit-recompute, and finish-summary consistency are all the *same* recompute
— zero reconciliation code, and the real-SQLite tests confirm it end to end. The
five PR types reduced to one `computeExerciseBests` pass plus a strictly-greater
comparison, and the report reused the exact `effectiveLoad`/`setVolume` functions
the workout detail screen uses, so "reconciles exactly" was true by construction.

**What was harder than expected?** Two judgement calls. First, *what counts as a
PR*: a brand-new load is technically "more reps than the (zero) prior at that
load", which would fire a rep-PR for every new weight — noisy and untrustworthy, so
I restricted rep-at-load PRs to loads with prior history (a new load is already a
weight PR). Second, the optimistic in-session badge: naively "beats the historical
best" lights up every set of a first-ever exercise; a *running* best (baseline +
earlier session sets) makes only genuine record-setters glow. Both are documented
and tested. Separately, adding more concurrent bodyweight readers exposed a latent
select-then-insert race in `settingsRepository.get()` — fixed idempotently.

**What should change before the next phase?** Nothing structural. When Phase 15
adds the trend, it plugs into `getExerciseSetHistory` (already returns per-workout
e1RM inputs) — no new query needed. Keep deriving records from history; never add a
records table.

## Lessons Learned

- **What surprised you:** the trustworthiness rule *simplified* the design rather
  than constraining it — refusing to cache PRs removed an entire class of
  invalidation bugs, so "records recede on delete" needed no code beyond the
  change-bus that already existed. The subtle work was semantic (what is a PR),
  not mechanical.
- **What documentation prevented mistakes:** FITNESS_DOMAIN §3.7 enumerated the five
  types and the strictly-greater rule; §3.5 fixed the e1RM r ≤ 12 cap and the
  single-canonical-Epley estimator; §3.4 fixed single-side-load-for-strength vs
  doubled-volume; ANALYTICS §5.5 gave the exact report contract, and the roadmap's
  "no fake trend (P8)" made the placeholder a requirement. The delight registry
  fixed the celebration to one badge + one haptic + one toast (never confetti).
- **What should be reused:** derive-from-history + change-bus recompute as the
  pattern for every "record/best/aggregate" going forward (never a cache); the
  `ExerciseSetRow` shape as the shared input for records and analytics;
  running-best flags for any optimistic in-session hint; idempotent
  `onConflictDoNothing` single-row upserts to make lazy-create race-safe.
- **What should be avoided:** caching PRs or bests (the rule, and it pays for
  itself); celebrating first-ever-at-a-load as a rep PR (noise); select-then-insert
  for a singleton row under concurrent readers.
- **Amendment proposals:** none — no frozen-document defect surfaced. TD-007 records
  the §5.5 remainder as scheduled with the Phase 15 time-series work.
