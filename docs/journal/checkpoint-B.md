# CP-B — Architecture Verification Checkpoint (end of M1)

**Date:** 2026-07-09 · **Covers:** Phases 3–8 (Milestone 1 · Training) · **Type:** review only (no features).

Per IMPLEMENTATION_ROADMAP §3, verified against the frozen documents with them
open. Device-only items are **not blockers** (owner directive): they accumulate
into the single consolidated TD-001 on-device pass.

**Gate at review:** `npm run check` green — typecheck, boundary lint
(`--max-warnings 0`), prettier, **159 tests / 41 suites**, `db:check` (9 tables).
Suite stable across repeated runs (the cross-realm `instanceof` flake fixed in
Phase 6 has not recurred).

## 1. PRODUCT_PRINCIPLES — the Tests

- **No vanity metrics.** Every number shown is a real training figure (volume,
  e1RM, bests, sessions) derived from logged sets; the Exercise Report shows no
  fabricated trend — it states "no trend line is shown rather than an estimated
  one" (P8). ✅
- **No dark patterns.** PR celebration is exactly one badge + one `success`
  haptic + one factual toast ("Workout saved · N PRs") — no confetti, no streak
  pressure, no badge dots, no nags (P16/P18). ✅
- **Logging speed (P9).** Within-session set inheritance, last-time prefill,
  ≤ 2-tap set entry (stepper *and* tap-to-type), one-tap ✓, and ≤ 1-tap suggested
  start. Tap *feel* is device-verified → TD-001. ✅ (mechanics), ⏳ (device)
- **Instant feedback / trust (P15, "trust is sacred").** Optimistic in-session PR
  badge; crash-safe drafts make recovery invisible; edit/delete recompute
  everywhere with records receding honestly. ✅
- **Offline-first (P19).** Zero network calls; entirely local. ✅

## 2. FITNESS_DOMAIN — formula & edge-case conformance

- **Formulas cite and match their sections.** Effective load (§3.4), working-set
  rule (§3.2), volume + unilateral doubling (§3.5/§3.4), Epley e1RM with the
  `r=1⇒w` case and r ≤ 12 trust cap (§3.5), PR types (§3.7). Single source of
  truth for thresholds in `domain/fitness/constants.ts` (CODING_STANDARDS §6.2). ✅
- **Edge-case tests (§8) present for shipped math:**

  | Edge | Rule | Named test |
  |---|---|---|
  | 1 | warm-ups excluded; warm-up-only not countable | `sets`/`workoutStats` ("edge 1") |
  | 2 | empty/draft workout not countable | `workoutStats` ("edge 2") |
  | 3 | two workouts one day, each PRs independently | `workoutRepository` ("edge 3") |
  | 4 | edit/delete → recompute, records recede | `personalRecords`, `sessionPRs`, `workoutHistory` |
  | 5 | bodyweight-load, unknown bodyweight → 0 vol, low-confidence | `sets`/`workoutStats` ("edge 5") |
  | 6 | unilateral: ×2 volume, single-side PRs | `sets`/`workoutStats`/`personalRecords` |
  | 7 | r > 12 → e1RM shown, no e1RM PR | `sets` ("edge 7"), `personalRecords` |
  | 13 | plausibility clamp (weight/reps) | `useSessionStore` ("clamps to plausibility ranges") |

  Two **minor coverage gaps** noted below (F-B3): edge 14 (device-local date
  helpers) and the PR-side of edge 8 (a `timed` exercise setting no weight/e1RM
  PR) are correct in code but lack a directly-named test. Edges 9–12/15/16 and
  the consistency half of 17 are M2+/analytics — out of M1 scope.

## 3. ARCHITECTURE — boundaries & state ownership

- **Boundary lint clean** and *enforced* (not just documented); the one-way rule
  holds through the new program/template feature. ✅
- **Vertical-slice read** (manual trace): `app/(tabs)/workouts/exercise/[id].tsx`
  → `features/workouts/screens/ExerciseReportScreen` → `hooks/useExerciseReport`
  → `data/repositories/workoutRepository.getExerciseSetHistory` → `core/db`
  (`getDb`) + `data/schema`, with the pure calculator in
  `domain/analytics/exerciseReport`. UI never touches SQL; domain stays pure. ✅
- **State ownership per §6:** the active session and the rest timer live in
  Zustand (`useSessionStore`, `useRestTimerStore`) — high-frequency, ephemeral.
  SQLite is written during a session at exactly two points: `checkpointDraft`
  (recovery, `SessionKeeper`) and `saveCompletedWorkout` (finish,
  `ActiveWorkoutScreen`) — matching §7.1. MMKV (`core/storage/prefs`) holds only
  the disposable theme override. Clean. ✅

## 4. DATABASE — schema, migrations, no derived tables

- **Schema == docs.** Nine tables — settings, exercises, programs, templates,
  template_exercises, workouts, workout_exercises, sets, workout_drafts — each
  matches DATABASE §3, including the §3.4 `template_id … ON DELETE SET NULL`
  provenance FK (added this milestone). ✅
- **No derived/cache tables.** PRs, bests, and the Exercise Report are pure
  recomputations from `sets`; there is no records/analytics/cache table
  (ARCHITECTURE rule 8, FITNESS_DOMAIN rule 7/13). `workout_drafts` is recovery
  state, not history, and is excluded from backups by design. ✅
- **Migrations immutable.** 0000–0004 committed; each early migration
  (0000–0002) is touched by exactly one commit (its origin phase). 0004 rebuilt
  `workouts` via the §5.4 table-rebuild `INSERT … SELECT` pattern; `db:check`
  applies all five cleanly. ✅

## 5. Tech-debt registry audit (§2)

| ID | State | On schedule? |
|---|---|---|
| TD-001 | Consolidated on-device verification (M0 + M1 checklists + screenshots for DB-backed screens) | ⚠️ deadline text stale — see F-B1 |
| TD-002 | Placeholder tabs (Dashboard/Nutrition/Measurements/Analytics) | ✅ replaced as their phases land (M2–M4); review at CP-D |
| TD-003 | Sheet drag-to-dismiss | ✅ due Phase 9 / 21 |
| TD-004 | `workouts.template_id` FK | ✅ **Resolved (Phase 8, migration 0004)** |
| TD-005 | Recent list not FlashList | ✅ due when an all-history archive ships |
| TD-006 | Dashboard-side Focus-Mode subtractions | ✅ due Phase 16 |
| TD-007 | Report range/trend/progression/notes | ✅ due Phase 15 |

No overdue debt on a correct reading (the Phase-4 item TD-004 is resolved; the
Phase-7 item TD-007 is scheduled for Phase 15). The only registry issue is
cosmetic (F-B1).

## 6. Roadmap drift

**None.** M1 = Phases 3–8 delivered exactly as specified (catalog → session →
history/prefill → crash-safety/rest/bar → PRs/report → programs/templates).
Every reduction from a phase's full scope is recorded as a scheduled TD, not
silently dropped. `domain/analytics` was introduced in Phase 7 for the §5.5
Exercise Report — correct placement per ARCHITECTURE §9, and an explicit Phase 7
deliverable, not scope creep.

## Findings

- **F-B1 · Low (registry hygiene).** TD-001's "Must be removed by" still reads
  **CP-A**, but device verification was deferred by standing owner directive into
  one consolidated pass spanning M0 + M1. As written it reads as overdue (and §2
  says overdue blocks a checkpoint). *Recommendation:* amend TD-001's deadline to
  "the consolidated on-device pass (before the M2 daily-use gate / when a device
  is available)," recording the standing approval — a doc edit, no code.
- **F-B2 · Low (deferred, tied to TD-001).** The screenshot archive (roadmap
  §2.3, both themes at checkpoints) is populated for Phases 00–02 but absent for
  03–08, whose screens are DB-backed and cannot render on web (expo-sqlite is
  native-only). Folds into the TD-001 device pass; non-blocking.
- **F-B3 · Low (test completeness).** Two edge cases are correct in code but lack
  a directly-named test: **edge 14** — the device-local date helpers
  (`todayIso`/`isoWeekday`/`weekdayLabel`/`formatRelativeDate`) have no unit test,
  and the `isoWeekday` Mon=0 conversion underpins weekday suggestions; **edge 8
  (PR side)** — no test asserts a `timed` exercise sets no weight/e1RM/volume PR
  (the `effectiveLoad = 0` guard handles it). *Recommendation:* add two small
  tests before Phase 9.
- **F-B4 · Info (forward-looking).** `domain/analytics/exerciseReport` is pure and
  correctly placed, but predates the ANALYTICS_ENGINE conventions (MetricResult,
  interpretation triplet) that CP-D/E enforce. Flag for conformance review when
  full analytics land — not an M1 concern.

## Verdict

**CP-B PASSES.** All five checkpoint dimensions are satisfied: product principles
upheld, domain formulas conform with edge-case tests, boundaries enforced and a
vertical slice reads clean, schema matches the docs with immutable migrations and
no derived tables, and the debt registry is on schedule with TD-004 resolved. The
four findings are minor and non-blocking — one registry-text correction (F-B1),
one documented device-pass deferral (F-B2), two small recommended tests (F-B3),
and one forward note (F-B4). No correctness defects, no boundary violations, no
roadmap drift.
