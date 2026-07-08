# Phase 4 — Active Workout Logging

**Closed:** 2026-07-08 · **Tag:** `v0.5.0-phase4` · **Milestone:** M1 Training

The product's heart. Optimized for logging speed and gym reality: big touch
targets, minimal typing, previous values reused, clear live progress.

## What was built

- **Domain formulas** (`domain/fitness`, pure): `constants.ts` (plausibility
  ranges + e1RM rep cap — the single source, CODING_STANDARDS §6.2);
  `sets.ts` (effective load per load-type, working-set rule, volume with
  unilateral doubling, low-confidence flag, Epley e1RM with the `r=1⇒w`
  special case, e1RM eligibility); `workoutStats.ts` (working-set count, total
  volume, countable-workout). Workout domain models added.
- **Schema + migration 0002** (DATABASE §3.4): `workouts` (multi-per-date,
  date index), `workout_exercises` (FK cascade to workouts, **FK RESTRICT** to
  exercises, `unilateral_counting` CHECK), `sets` (rpe/rir range CHECKs). The
  RESTRICT FK now backs Phase 3's archive-not-delete.
- **`workoutRepository`** with `saveCompletedWorkout` in one all-or-nothing
  transaction (new driver-agnostic `runInTransaction` in `core/db`), `getById`
  (full tree with joined exercise names), `countOnDate`, `remove`.
- **Session store** (`useSessionStore`, Zustand): the in-memory active session —
  start, add/remove exercise, add set (**inherits the previous set's
  weight/reps within the session** for ≤1-tap repeats), update (clamped to
  ranges), toggle done, warm-up toggle, unilateral counting, discard. Survives
  in-app navigation (module-scoped); crash-safe SQLite drafts are Phase 6.
- **Active Workout screen** (full-screen route above the tabs): live elapsed
  timer, progress readout (exercises · sets done/total), exercise cards, the
  gym-optimized **`SetValueControl`** (big 48pt ± targets *and* tap-to-type,
  UI_UX §5.3), one-tap ✓ complete with haptic, warm-up toggle, fast exercise
  picker sheet, and a finish → summary (duration/volume/working sets) → save
  or discard (Dialog-guarded when sets were logged) flow.
- **Workouts home** restructured: the tab is now a small hub (Start empty
  workout · Resume when a session is active · Manage exercises →), with the
  Phase-3 library moved to a pushed sub-route (`ExerciseLibraryScreen`).
- **Tests (34 new, 104 total):** domain edge cases 1,5,6,7,8 (sets +
  workoutStats); repository — save/drop-empty, two-per-day (edge 3),
  transaction rollback on bad FK, and FK-RESTRICT-blocks-referenced-delete;
  session store; `SetValueControl`; and an **Active Workout screen render
  test** driving the loop (renders session, one-tap complete flips state,
  empty state).

## What changed

New: `core/utils/{scalars,date}`, `domain/fitness/{constants,sets,workoutStats}`,
`domain/models/workout`, migration 0002, `data/repositories/workoutRepository`,
`data/mappers` (via getById), `core/db/runInTransaction`, workouts feature
(`stores`, `logic/sessionMapping`, `hooks/{useElapsed,useDefaultBodyweight}`,
components `SetValueControl`/`SetRow`/`ActiveExerciseCard`/`ExercisePickerSheet`/
`WorkoutSummarySheet`, screens `ActiveWorkoutScreen`/`WorkoutsHomeScreen`),
routes (`app/(tabs)/workouts/` stack, `app/active-workout.tsx`). Renamed the
Phase-3 library screen. Deps: zustand (was mandated but not yet installed).
No frozen document changed.

## Screens affected

Workouts home (new hub), Active Workout (new, the heart), Exercise Library
(moved to sub-route). Set-value control, set row, exercise card, picker &
summary sheets.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Effective load / working-set / volume / e1RM per FITNESS_DOMAIN §3 | domain unit tests (edge 1,5,6,7,8) | ✅ |
| Empty workout not countable (edge 2) | workoutStats test | ✅ |
| Save persists tree; empty sets dropped; warm-up kept | repo test (getById round-trip) | ✅ |
| Two workouts same day (edge 3) | repo test | ✅ |
| Transaction rollback on invalid FK (no partial tree) | repo test | ✅ |
| FK RESTRICT blocks deleting a referenced exercise | repo test | ✅ |
| Session: add-set inherits previous; clamps; toggle; discard | store tests | ✅ |
| Logging loop renders; one-tap ✓ flips state; empty state | Active Workout screen render test | ✅ |
| `npm run check` | typecheck + lint + format + 104 tests + db:check (5 tables) | ✅ green |
| On-device gym walk (5 exercises/15+ sets ≤2-tap, both themes, minimize/resume, save & discard, keyboard entry, haptics) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** (DB-backed screens; expo-sqlite native-only) — the
   Active Workout visual + full gym walk fold into the consolidated TD-001
   device pass (checklist extended below). The loop is proven behaviorally by
   the screen render test.
2. **In-app session continuity only:** the store survives tab navigation, but
   surviving an app kill (crash-safe SQLite draft), the rest timer, and the
   app-wide session bar are **Phase 6** by the roadmap — not built here.
3. **No history prefill:** first set of an exercise defaults 0/0 (accepted debt,
   removed in Phase 5); within-session set inheritance already gives fast repeats.
4. **PRs earned** are not in the summary yet (PR detection is Phase 7).
5. `workouts.template_id` has no FK yet (templates table lands Phase 8) — TD-004.

## Technical debt introduced

- **TD-004** — `workouts.template_id` is a plain column; its FK to `templates`
  (DATABASE §3.4) is added when that table exists in Phase 8.
- TD-001 device checklist extended with the full active-workout gym walk.

## Retrospective

**What went well?** The pure domain split paid off — every FITNESS_DOMAIN §3
rule is a tiny tested function, and the screen/store just compose them. The
`runInTransaction` helper made all-or-nothing saves driver-agnostic and the
rollback test passes against real SQLite. The session store kept the screen
thin.

**What was harder than expected?** The shared `AppDb` union type
(`Expo | BetterSQLite3`) silently collapsed Drizzle's builder generics — join
projections typed as the raw joined shape and `.values()` reported "0 args".
Fixing it to a single production-driver type with a cast at the test-injection
boundary restored inference (both drivers share the runtime builder API). Also:
zustand was mandated since Phase-0 planning but never actually installed — the
store test caught it. And the token lint correctly rejected `marginHorizontal:
0` (raw zero), which forced removing unnecessary margin overrides.

**What should change before the next phase?** Install-and-verify stack
dependencies when first referenced, not when first tested. Treat the
single-driver `AppDb` type as settled. Keep building screens as thin
compositions over tested store + pure logic — the render test then needs only
to confirm wiring.

## Lessons Learned

- **What surprised you:** a *union* database type quietly destroys Drizzle's
  fluent-builder type inference (methods degrade to base overloads) — a single
  concrete type is required, with the alternate driver cast in at the edge.
  Epley at 1 rep needed an explicit special-case (`w`, not `w·(1+1/30)`), caught
  by a test written straight from the doc. The react-compiler lint flags a
  hoisted `function` used earlier in a value position, so effect-referenced
  helpers must be `const` declared before the effect.
- **What documentation prevented mistakes:** FITNESS_DOMAIN §3.2/§3.4/§3.5 gave
  exact working-set, effective-load, doubling, and e1RM rules — the edge-case
  list (§8) became the test names verbatim, and #6 unilateral doubling was
  unambiguous because the counting marker is stored on the data (DATABASE §3.4).
  ARCHITECTURE §7.1's "single write on finish" shaped the store-vs-SQLite split;
  CODING_STANDARDS §6.2 kept every threshold in one constants module.
- **What should be reused:** `runInTransaction` for all multi-row writes; the
  store + pure-logic + thin-screen pattern; screen render tests as the
  behavioral stand-in for un-screenshottable DB-backed screens; the
  `SetValueControl` (steppers + tap-to-type) as the numeric-entry pattern.
- **What should be avoided:** union types over driver handles; referencing
  not-yet-installed deps; raw numeric literals (even `0`) in feature styles;
  hoisted-function references inside effects.
- **Amendment proposals:** none — no frozen-document defect surfaced.
