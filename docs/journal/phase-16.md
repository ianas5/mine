# Phase 16 — Dashboard v1 (the daily briefing)

**Closed:** 2026-07-09 · **Milestone:** M3 (Awareness)

The Dashboard becomes the app's home: a **daily briefing**, not an analytics report. It
answers exactly one question — *"What should I know today?"* — calm, fast, and useful,
with no charts or secondary metrics. Its content is the closed ANALYTICS §6.5 list and
nothing else; all deeper analysis stays in the Analytics tab. This also retires the last
placeholder screen (TD-002) and closes the Focus-Mode (TD-006) and bodyweight-source
(TD-009) debts.

## The briefing discipline (ANALYTICS §6.5 / UI_UX §7.2, the standing principle)

- **Closed content list, top to bottom:** header greeting (with the trend weight and the
  morning weigh-in prompt state) → today's workout card → calories & protein rings + macro
  bars → insight slot (empty-capable) → weekly streak line → quick actions. **Nothing
  else** — adding a card would require amending §6.5 first.
- **No charts on the dashboard.** The trend weight appears only in the greeting line, never
  as a card; every chart lives in Analytics. A glance answers "right now."
- **Context reorders, never adds/removes** (§7.6): mornings put the workout card above
  macros and surface the weigh-in prompt; afternoons/evenings put remaining calories &
  protein first. The layout stays recognizably the same app all day (P14).
- **Focus Mode is subtraction** (§5.1): during a live session the workout card becomes the
  live session card (first), the insight slot hides, and quick actions slim to Return · Log
  Meal · Add Weight — no second UI.

## What was built

- **Consistency & streaks (`domain/fitness/consistency.ts`, pure, §3.8):** `weekStartIso`
  (ISO-Monday), `countableWorkoutsByWeek`, `currentWeekProgress` (the in-progress week as
  **progress, never a %** — a fresh Monday reads "0 of N", not a crash), `weekConsistencyPercent`
  (completed weeks), and `weeklyStreak` (consecutive weeks meeting target; the current
  partial week counts only if already met and never breaks the streak). Phase 17's
  `WorkoutAnalyticsCalculator` will consume these same functions.
- **Bodyweight source unified (TD-009):** `useDefaultBodyweight` now prefers the latest
  `body_snapshots` weigh-in (FITNESS_DOMAIN §5.2 canonical) via the new
  `bodyRepository.getLatestWeightKg()`, falling back to `settings.defaultBodyweightKg`;
  reactive to body + settings writes.
- **Countable-workout query:** `workoutRepository.getCountableWorkoutDatesSince` fetches
  the rows and classifies countability with the domain `isWorkingSet` (which needs load
  type) — SQL stays free of domain semantics (ANALYTICS rule 9).
- **Dashboard data hook (`useDashboard`):** composes nutrition, program suggestion, streak,
  and trend weight straight from repositories + pure domain. The dashboard is a *different
  feature* from workouts/nutrition, so it reads the data layer directly rather than
  importing their hooks (ARCHITECTURE §4). Reactive to every source table.
- **Dashboard screen:** the full closed-list briefing with daypart ordering and Focus-Mode
  subtractions; a greeting that shows the trend weight, the morning weigh-in prompt, and a
  warm "back at it — first session in N weeks" after a ≥ 14-day gap. Quick actions reach
  their sheet in **one tap** via navigation intent params (`?open=meal|weight|measure`),
  driven as **derived sheet visibility** (no `setState`-in-effect) with the intent cleared
  on close so a repeat tap re-fires.
- **Primitives:** `ProgressBar` (macro/streak tracks) in `core/ui`; `Ring` (react-native-svg
  calorie/protein rings) in the dashboard feature.
- **Composition root:** `app/(tabs)/index` reads the active-session summary from the
  workouts store and passes it to the dashboard as data — the dashboard never imports the
  workouts feature.
- **Doc:** applied the ratified DATABASE §3.1 `target_weight_kg` amendment (+ changelog).
- **Tests (9 new, 285 total):** the §3.8 consistency/streak edge cases — ISO-week bucketing,
  two-per-day counting, partial-week progress, the Monday-reset, streak continuation across
  a partial week, and streak-break on a missed completed week.

## What changed

New: `domain/fitness/consistency` (+ tests); `features/dashboard/{hooks/useDashboard,components/Ring,screens/DashboardScreen}`;
`core/ui/ProgressBar`. Modified: `useDefaultBodyweight` (latest-weigh-in preference) +
`bodyRepository.getLatestWeightKg`; `workoutRepository.getCountableWorkoutDatesSince`;
`app/(tabs)/index` (session summary → dashboard); `NutritionScreen` + `MeasurementsScreen`
(quick-action intent open); DATABASE.md §3.1 (ratified amendment). No migration, no new deps.

## Screens affected

Dashboard (now the real briefing), Nutrition + Measurements (quick-action intent open).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| ISO-Monday week start; two-per-day counting | domain tests | ✅ |
| Current week = progress (Monday resets to 0 of N, no crash) | domain tests | ✅ |
| Completed-week % capped at 100; no divide-by-zero | domain test | ✅ |
| Streak counts consecutive completed weeks | domain test | ✅ |
| Partial current week never breaks the streak | domain test | ✅ |
| Streak breaks on a fully-elapsed missed week | domain test | ✅ |
| `npm run check` | typecheck + lint + format + 285 tests + db:check (15 tables) | ✅ green (3× stable) |
| On-device dashboard walk (closed list, daypart ordering, Focus Mode, 1-tap quick actions, streak across a week boundary, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **The dashboard is composition-heavy UI** — its rendering (rings, ordering, Focus-Mode
   transitions, daypart) is device-only and folds into TD-001. The rigor that *is* verified
   off-device is the §3.8 consistency/streak math (the only new logic) and the boundary-clean
   data composition (typecheck + boundary lint).
2. **The insight slot is intentionally empty** — the `InsightEngine` is Phase 18; the slot
   ships empty-capable (renders nothing) so nothing is faked before then (P8).
3. **No MMKV cold-start cache (ANALYTICS §8).** The briefing computes well inside the §7
   budget at personal scale, so per "never optimize prematurely" the disposable paint cache
   is deferred until a measured need; correctness never depends on it (rule 10).
4. **Live session card shows a static snapshot** (exercises/sets logged, not a ticking
   clock) — the app-wide session bar already carries the live elapsed time; the dashboard
   card is a glance + Return, refreshed on navigation.

## Technical debt

**TD-002, TD-006, TD-009 all resolved** this phase. No new debt introduced. TD-001 gains
the dashboard walk. (TD-003/005/008/010 unchanged; TD-007's remaining range/notes → Phase 17.)

## Retrospective

**What went well?** The closed-list discipline made the dashboard easy to *not* over-build:
every candidate widget was checked against §6.5, and "no charts here" kept it calm. The
§3.8 functions dropped out cleanly as pure date math, so the streak's trickiest rule
(a partial week neither counts nor breaks) is a two-line branch with its own test. The
composition-root pattern (app route passes the session summary) resolved the feature-boundary
tension without moving the session store or weakening the lint.

**What was harder than expected?** Making quick actions open a sheet in one tap *without* a
`setState`-in-effect (the `react-hooks/set-state-in-effect` rule bans it). The clean answer
was to make sheet visibility **derived** from the navigation intent param and clear the
param on close in the event handler — no effect, and repeat taps re-fire. The other subtlety
was reading the active session across the feature boundary; the app route (which may import
any feature) building a small summary is the right seam.

**What should change before the next phase?** Nothing structural. Phase 17 (workout & muscle
analytics) consumes the §3.8 functions built here and fills the Analytics Training section;
the dashboard's insight slot waits for the Phase 18 `InsightEngine`.

## Lessons Learned

- **What surprised you:** how much a *closed* content list simplifies design — the hardest
  UI decision (what goes on the home screen) was already made by §6.5, so the work was
  composition, not invention.
- **What documentation prevented mistakes:** ANALYTICS §6.5 fixed the exact dashboard
  contents (and "nothing else"); UI_UX §7.6 fixed daypart reordering (reorder, never
  add/remove) and §5.1 fixed the Focus-Mode subtractions; FITNESS_DOMAIN §3.8 fixed the
  weekly-not-daily streak and the partial-week progress rule that stops every Monday reading
  as a consistency crash.
- **What should be reused:** derived-visibility-from-route-intent (a lint-clean alternative
  to effect-driven sheet opening); the composition-root pattern for cross-feature state
  (pass a summary as data from the app route); classifying domain-semantic rows (countable
  workouts) in the data layer via a domain function rather than in SQL.
- **What should be avoided:** putting charts or secondary metrics on the dashboard (it's a
  briefing, not a report); `setState` inside an effect to react to navigation; importing one
  feature's store/hooks from another feature (compose at the app root instead); a daily
  streak (weekly is the honest, rest-day-respecting unit).
- **Amendment proposals:** none new. The Phase 15 DATABASE §3.1 `target_weight_kg` amendment
  was ratified and is now applied to the doc. No frozen-document defect surfaced.
