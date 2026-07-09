# Phase 17 — Workout & Muscle Analytics

**Closed:** 2026-07-09 · **Milestone:** M4 (Intelligence) — first phase

The Training section answers *"am I training right?"* The governing principle this phase:
**training analytics must explain training quality, not just quantity.** Every metric maps
to a coaching question — am I getting stronger? consistent? balanced? which muscles am I
neglecting? where am I improving? — and the Muscle Report reads like a coach reviewing the
last months, not a spreadsheet of totals.

## Quality over quantity (the standing principle)

- **Strength, not volume, is the "am I improving?" signal.** Progress is read from the
  key-exercise **e1RM trends**; weekly **volume** is shown once, neutrally, as context
  (its trend classification is `neutral` — more volume isn't universally "better").
- **Balance and neglect are surfaced as attention, not applause.** Push:Pull (healthy
  0.8–1.25) and Upper:Lower (1.0–2.0) flag softly when out of band; Most/Least-trained
  flags a neglected group. `Other`-group and legs/core volume are excluded from the
  push/pull and upper/lower ratios (§3.3) so the balance stays honest.
- **Every data-gated metric is a `MetricResult`.** A lift below the §6.4 minimums reads
  "needs more sessions"; a never-trained muscle shows an honest zero state; missed
  workouts are `no-target-set` without a program schedule — never fabricated.
- **The engine stays pure.** SQL only joins/windows; all domain math (working sets,
  effective load, unilateral doubling from the stored marker, volume, e1RM, §3.4
  bodyweight resolution) runs in `domain/` (ANALYTICS rule 9).

## What was built

- **Shared training input** (`domain/analytics/trainingData.ts`): `TrainingWorkout`/
  `Exercise`/`Set` + `WeighIn` + the §3.4 `resolveBodyweightForDate` (latest weigh-in on/
  before the date, else settings, else null). Fed by
  `workoutRepository.getTrainingWorkoutsSince` — a join that classifies nothing (the
  calculators do).
- **`WorkoutAnalyticsCalculator`** (§5.1, pure): total workouts + vs-previous, frequency,
  consistency (this-week progress + streak, reusing the Phase 16 §3.8 functions), total
  volume + weekly-bucketed **volume trend** + vs-previous, **volume by muscle group**,
  **Push:Pull** and **Upper:Lower** balance with band flags, **Most/Least trained** over
  30 days (zeros included), **key-exercise strength summary** (top 5 lifts by session count,
  each with its e1RM trend), average session duration, and **missed workouts** (program-
  gated).
- **`MuscleAnalyticsCalculator`** (§5.6, pure): per canonical group — 30-day volume +
  working sets, current-week volume, frequency, **strongest** lift (best e1RM),
  **fastest-** and **weakest-improving** lift (by e1RM slope), volume trend, last-trained
  recency, and an honest **untrained** zero state.
- **Analytics Training section** (`TrainingSection`): consistency, a "Getting stronger?"
  summary listing the main lifts' verdicts, the two balance tiles, the most/least-trained
  pair, and a single weekly-volume `ChartFrame` — coaching answers, not a totals page — plus
  a link to the Muscle Report.
- **Muscle Report screen** (`/analytics/muscles`): a per-group coach's review (strongest,
  fastest-improving, recency, volume) with never-trained groups grouped under "Not trained
  in this range." The Analytics tab became a route stack.
- **`useTrainingAnalytics(range)`** composes both calculators from one windowed fetch,
  memoized by data version + range; the §3.4 fallback is `settings.defaultBodyweightKg`
  with weigh-ins preferred (consistent with TD-009).
- **Tests (12 new, 297 total):** a hand-computed 4-week fixture the calculators reconcile
  to exactly — volume + per-group volume, push:pull & upper:lower ratios + flags,
  most/least trained with zero groups, key-exercise trends (improving vs. insufficient),
  unilateral doubling from the stored marker, `other`/legs exclusion from balance, missed-
  workout gating, previous-period comparison, and the Muscle Report's strongest/fastest-
  improving/last-trained + never-trained zero state.

## What changed

New: `domain/analytics/{trainingData,workoutAnalytics,muscleAnalytics,trainingFixture}`
(+ tests); `features/analytics/{components/TrainingSection,hooks/useTrainingAnalytics,screens/MuscleReportScreen}`;
`app/(tabs)/analytics/{_layout,index,muscles}` (tab → stack).
Modified: `workoutRepository.getTrainingWorkoutsSince`; `AnalyticsScreen` (Training section
live; Nutrition placeholder remains). No migration, no new deps, no frozen document changed.

## Screens affected

Analytics home (Training section live), Muscle Report (new).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Total + per-group volume exact (primary group only, §3.3) | domain test (fixture) | ✅ |
| Push:Pull & Upper:Lower ratios + band flags; legs/core/other excluded | domain test | ✅ |
| Most/least trained over 30d, zero groups included | domain test | ✅ |
| Key-exercise strength: improving vs. insufficient below minimums | domain test | ✅ |
| Unilateral single-logged doubling from the stored marker (§3.5) | domain test | ✅ |
| Missed workouts gated on a weekday schedule; else no-target-set | domain test | ✅ |
| Previous equal-length period comparison | domain test | ✅ |
| Muscle Report: strongest / fastest-improving / last-trained; never-trained zero state | domain tests | ✅ |
| Consistency this-week progress (never a mid-week %) + streak | domain test | ✅ |
| `npm run check` | typecheck + lint + format + 297 tests + db:check (15 tables) | ✅ green (3× stable) |
| On-device Training + Muscle Report walk (coaching read, band flags, range recompute, reconcile vs history, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **The section/report UI is device-only** — its rendering (the volume chart, tiles,
   per-group cards) folds into TD-001. The rigor that *is* verified off-device is both
   calculators, reconciled exactly to a hand-computed fixture.
2. **Bodyweight for e1RM in the key-exercise/muscle trends uses the latest weigh-in**
   (not per-workout-date), while *volume* resolves bodyweight per workout date (§3.4). The
   e1RM approximation only affects bodyweight-load lifts, which are secondary in this
   external-dominant profile; volume — the balance/neglect input — is exact per date.
3. **The per-exercise Report screen body is still all-time** (range-scoping + recent-notes
   remain in TD-007) — Phase 17 delivered the *training-wide* range analytics, not that
   screen's own range view.

## Technical debt

None introduced. TD-001 gains the Training + Muscle Report walk. TD-007's remaining
selected-range Exercise-Report body + recent-notes are deferred to an unscheduled report
pass. (TD-003/005/008/010 unchanged.)

## Retrospective

**What went well?** Framing each metric as a coaching question kept the section from
becoming a totals dump: volume got exactly one neutral chart, while strength (e1RM trends)
and balance/neglect (flagged tiles) carry the verdicts. The calculators composed cleanly
over already-built pieces — §3.8 consistency (Phase 16), `computeExerciseTrend`/`Bests`
(Phases 15/7), `setVolumeKg` with its stored-marker doubling — so the new code is mostly
aggregation, and the hand-computed fixture made every number falsifiable.

**What was harder than expected?** Deciding volume's directionality. Volume up isn't
universally good (a cut lowers it on purpose), so classifying the volume trend `neutral`
and letting e1RM carry "am I improving?" was the honest call — it kept the section aligned
with "quality, not quantity." The other care point was keeping SQL dumb: the training query
joins and windows but the *countable*/*working-set*/*volume* decisions all happen in the
calculators via domain functions, never in SQL.

**What should change before the next phase?** Nothing structural. Phase 18 (nutrition
analytics + the Insight engine) fills the Nutrition section and lights up the dashboard's
empty insight slot; the balance/neglect and strength verdicts built here are exactly the
evidence several §6.2 insight rules will cite.

## Lessons Learned

- **What surprised you:** how much "which coaching question does this answer?" pruned the
  design — metrics that were only totals (raw set counts, per-session volume tables) simply
  didn't earn a place once every tile had to justify a decision.
- **What documentation prevented mistakes:** FITNESS_DOMAIN §3.3 fixed the push/pull &
  upper/lower mappings and the primary-group-only, `other`-excluded volume rule; §3.4/§3.5
  fixed effective load, unilateral doubling, and e1RM; ANALYTICS §5.1/§5.6 fixed exactly
  which metrics exist and their bands; §4 fixed the completed-vs-partial-week comparison
  rule the streak/consistency rely on.
- **What should be reused:** the hand-computed fixture as the calculators' contract (every
  metric reconciled by hand); classifying domain-semantic rows (countable, working set,
  volume) in `domain/` from SQL-fetched rows; composing new calculators over existing pure
  functions rather than re-deriving math.
- **What should be avoided:** treating volume as "more = better" (it's context, not a
  verdict); putting a wall of totals on an analytics section (each tile must answer a
  question); implementing working-set/volume semantics in SQL; fabricating a trend for a
  lift or muscle below the §6.4 minimums (say "needs more sessions").
- **Amendment proposals:** none — no frozen-document defect surfaced. No new debt.
