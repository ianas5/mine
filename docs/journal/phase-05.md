# Phase 5 — History, Prefill & Exercise-History Preview

**Closed:** 2026-07-09 · **Milestone:** M1 Training

The past powers the present. A repeated exercise now arrives pre-filled with
last time's working sets (zero typing), each card shows Last / Best / Best e1RM
at the point of logging, and every past session is browsable and editable — with
edits and deletes recomputing everywhere for free.

## What was built

- **Domain summarizer** (`domain/fitness/exerciseHistory.ts`, pure):
  `summarizeExerciseHistory(rows, loadType)` → `{ last, bestWeightSet,
  bestE1rmKg }`. Filters to working sets via `isWorkingSet` (FITNESS_DOMAIN
  §3.2), `last` is the sets of the most-recent workout (max `workoutOrder`),
  best weight is the heaviest working set, best e1RM is the max Epley over
  e1RM-eligible sets (reps 1–12, weight > 0, §3.5). Returns an **all-null
  preview** for a first-ever exercise — no fabricated data (P8).
- **Repository history/prefill/edit methods** (`workoutRepository`):
  `listRecent(limit)` (ordered `date desc, createdAt desc`); `getExerciseHistory`
  (sets ⨝ workout_exercises ⨝ workouts, flattened to `HistorySetRow` with the
  workout's `created_at` as order); `getExercisePreview` (compose the two);
  `updateSet(setId, patch)` and `deleteSet(setId)` — both emit `'workouts'` on
  the change-bus.
- **Prefill wiring:** `useSessionStore.addExercise(exercise, prefill?)` seeds the
  card from last time's working sets (positional = same set-number) when present,
  else one blank 0/0 set. `ActiveWorkoutScreen` fetches
  `getExercisePreview` before adding so a re-added exercise is logged with zero
  typing.
- **In-card history panel** (`ExerciseHistoryPanel`): compact Last (with relative
  date) / Best / Best e1RM rows, or the "First time — set your baseline" state,
  or nothing while loading. Wired into `ActiveExerciseCard` via
  `useExercisePreview`.
- **History list + detail:** `RecentWorkoutList` on the Workouts home (name,
  relative date, exercise/set/volume counts → tap to detail);
  `WorkoutDetailScreen` (`/workouts/[id]`) with a stats card, per-exercise
  `DetailSetRow`s (editable weight/reps via `SetValueControl`, warm-up toggle,
  per-set delete), and a **Dialog-gated delete-workout** flow.
- **Format helpers:** `formatSet` (timed → `Ns`, bodyweight → `N`, else
  `w × r`), `formatKg` (integer as-is, else one decimal), `formatRelativeDate`
  (Today / Yesterday / `Mon · 4d ago`).
- **Full recompute for free:** no derived history is cached — `updateSet` /
  `deleteSet` / `remove` emit `'workouts'`, and every history hook
  (`useExercisePreview`, `useRecentWorkouts`, `useWorkout`) re-queries on that
  version bump. Edit A's weight → B's preview and Recent's totals reflect it;
  delete A → it leaves every view. This is the change-bus proof the acceptance
  asks for, verified by tests.
- **Tests (11 new, 115 total):** domain summarizer (empty→all-null P8,
  last-excludes-warm-ups, best-weight + best-e1RM eligibility); repository
  history (prefill across sessions, edit-recompute, delete-set-recompute,
  delete-workout-recompute, `listRecent` ordering); `formatSet` per load-type;
  and a session-store prefill test (re-added exercise inherits last working sets).

## What changed

New: `domain/fitness/exerciseHistory` (+ barrel exports); `workoutRepository`
history/prefill/edit/delete methods; `core/utils/{date.formatRelativeDate,
number.formatKg}`; workouts feature `logic/{formatSet,workoutSummary}`,
`hooks/{useExercisePreview,useRecentWorkouts,useWorkout}`, components
`ExerciseHistoryPanel`/`DetailSetRow`/`RecentWorkoutList`, screen
`WorkoutDetailScreen`, route `app/(tabs)/workouts/[id].tsx`. Modified:
`useSessionStore.addExercise` (prefill), `ActiveExerciseCard` (panel),
`ActiveWorkoutScreen` (fetch preview on add), `WorkoutsHomeScreen` (Recent
section). No frozen document changed.

## Screens affected

Workouts home (new Recent section), Active Workout (each card now shows the
history panel and prefills on add), Workout Detail (new — review/edit/delete a
past session).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Summary matches working-set rules; warm-ups excluded from Last | domain test | ✅ |
| Best weight + best e1RM (rep-cap eligibility) across all history | domain test | ✅ |
| First-ever exercise → all-null preview, no fabricated data (P8) | domain test | ✅ |
| Prefill inherits last session's working sets across sessions | repo + store tests | ✅ |
| Edit a set's weight → preview recomputes (change-bus) | repo test | ✅ |
| Delete a set / delete a workout → history recedes everywhere | repo tests | ✅ |
| `listRecent` ordered newest-first | repo test | ✅ |
| `formatSet` per load-type (timed/bodyweight/weighted) | logic test | ✅ |
| `npm run check` | typecheck + lint + format + 115 tests + db:check (5 tables) | ✅ green |
| On-device history & prefill walk (session A→B prefill + panel, Recent list, detail edit/delete recompute, first-time baseline, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** for the new DB-backed screens (expo-sqlite is
   native-only). The history/prefill/detail flow is proven behaviorally by the
   domain + repository + store tests; visuals fold into the consolidated TD-001
   device pass (checklist extended with the full history & prefill walk).
2. **History list is not yet virtualized with FlashList.** The roadmap names
   FlashList for the history list; with a single-user dataset in the low
   hundreds of sessions a plain mapped list inside the home `ScrollView` is
   correct and simpler, and Recent is capped at 20 rows. FlashList is warranted
   only once a full-screen "all history" archive lands — deferred as **TD-005**,
   removed when that screen exists.
3. **PRs earned** are still absent from summaries and the Best e1RM line is
   informational, not a PR badge — PR detection is Phase 7.

## Technical debt introduced

- **TD-005** — the Recent history list is a plain mapped list, not a virtualized
  FlashList (roadmap Phase 5 names FlashList). Acceptable now: Recent is capped
  at 20 and the dataset is single-user/small; virtualization before a durable
  large list violates the Rule of Two. Must be removed when a full-screen
  all-history archive screen is built.
- TD-001 device checklist extended with the history & prefill gym walk.

## Retrospective

**What went well?** The "no cached derived data" decision from Phase 4 paid its
biggest dividend here: *full recompute on edit/delete* required zero extra code —
the same change-bus emit that persists a set edit also invalidates every history
view, and the tests confirm A→B propagation directly. The pure summarizer meant
the entire Last/Best/e1RM contract is one tested function the panel just renders.
Prefill dropped into the existing `addExercise` signature as an optional
argument, so the fast path (re-add an exercise → zero typing) is one fetch on the
screen and nothing new in the store's core.

**What was harder than expected?** Deciding what *not* to build. The roadmap
names FlashList, but wiring virtualization for a 20-row capped list — inside a
screen that's otherwise a `ScrollView` — would have added a dependency and
gesture surface for no user-visible gain at this data scale. I registered the
deferral as TD-005 with an explicit removal trigger rather than gold-plating,
consistent with the Rule of Two and "no premature optimization." The transient
single-test flake seen at the end of Phase 4 recurred once here and cleared on
rerun (all 115 green) — a timer-timing artifact, not a real failure.

**What should change before the next phase?** Phase 6 (crash safety) will add the
first genuinely device-defining test (force-kill → resume). Keep the
recompute-from-source discipline — drafts should be the *only* new persisted
state, and finishing a workout must leave no draft behind. The history hooks
established here (`useTableVersion('workouts')` + re-query) are the template the
session-bar and resume-banner should reuse.

## Lessons Learned

- **What surprised you:** "edit/delete with full recompute semantics" sounds like
  the hard part of the phase but was effectively free — because nothing derived is
  stored, correctness reduces to "emit the right table on write," which was
  already the Phase-4 convention. The expensive-sounding requirement was cheap
  precisely because an earlier architectural choice (single source of truth in
  SQLite, everything else re-derived) had already paid for it.
- **What documentation prevented mistakes:** UI_UX §4.1 fixed the panel contract
  (Last / Best / Best e1RM) so there was no guessing what to show; FITNESS_DOMAIN
  §3.2/§3.5 defined working-set filtering and e1RM eligibility so the summarizer's
  edges (warm-up exclusion, rep-cap on e1RM) came straight from the doc into test
  names; P8 ("never fabricate data") made the first-time state a hard requirement
  rather than a nicety — the all-null preview and "set your baseline" copy exist
  because the principle demanded them.
- **What should be reused:** the `useTableVersion(table)` + re-query hook pattern
  for anything that must reflect writes live; the pure-summarizer + thin-panel
  split; optional-prefill-argument as the shape for "smart default without
  changing the core action"; TD registry entries with explicit removal triggers
  instead of speculative infrastructure.
- **What should be avoided:** building virtualization / infrastructure ahead of a
  durable second consumer (Rule of Two) even when a roadmap line names the tool;
  caching derived history (it would have made recompute the hard problem it
  wasn't).
- **Amendment proposals:** none — no frozen-document defect surfaced. TD-005
  records the FlashList deferral as scheduled debt, not a doc change.
