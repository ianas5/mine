# Phase 8 — Programs & Templates

**Closed:** 2026-07-09 · **Milestone:** M1 Training

Planned training arrives: programs, weekday-mapped session templates with targets,
one-tap smart-default starts, and repeat-last. The guiding invariant throughout —
**a template is a plan, not history**: a workout copies its exercises and sets at
save time, so editing or deleting a template never rewrites a past session.

## What was built

- **Migrations 0004** (DATABASE §3.3/§3.4): `programs`, `templates` (program FK
  cascade, `weekday` 0–6 CHECK), `template_exercises` (targets + `rest_seconds`,
  cascade to template, RESTRICT to exercise), and the **`workouts` table rebuild**
  (SQLite table-rebuild pattern, §5.4) adding `template_id … ON DELETE SET NULL`.
  This resolves **TD-004**.
- **`programRepository`**: programs/templates CRUD; **single-active invariant** in a
  transaction (clear all, set one); `getActiveProgram`; template resolution
  (joins the catalog for names/load types); `getRecentTemplateUses` for the
  weekday-mode suggestion. Template edits rebuild only `template_exercises`;
  deletes cascade to the plan and SET NULL the provenance — never a workout.
- **Smart-default suggestion** (`domain/fitness/suggestTemplate`, pure, UI_UX §5.2):
  active program's template for today's weekday → most-frequent template on this
  weekday over 8 weeks → Repeat Last → nothing. All fallbacks, ties resolve to the
  most recent.
- **Session store extensions**: `begin(startedAt, name, prepared, templateId?)`
  starts a pre-built session; `SessionExercise` now carries a display-only
  `target` and `restSeconds`; `templateId` is remembered as provenance and written
  on save. Targets, rest, and provenance all round-trip through the crash draft.
- **Start flows** (`startPreparation`): `prepareTemplateStart` pre-loads a
  template's exercises, pre-filling from last time's working sets (padded to the
  target set count) and attaching targets; `prepareRepeatLast` reloads a past
  workout's working sets. `useStartWorkout` begins the session and opens the
  Active Workout screen — the suggested start is ≤ 1 tap.
- **Logging-time targets**: each active card shows `Target · 3 × 8–10 @ RPE 8`
  (`formatTarget`), and a working set's ✓ seeds the rest timer from the template's
  `rest_seconds` (falling back to 90 s).
- **Management UI**: Workouts home gains a smart **Start section** (suggested start
  + empty + repeat-last) and a Programs entry; `ProgramsScreen` (list + create),
  `ProgramDetailScreen` (rename, set/clear active, session list, delete),
  `TemplateEditorScreen` (name, weekday chips, per-exercise target steppers, add
  via the exercise picker, save/delete). New routes under `workouts/programs` and
  `workouts/templates`.
- **Tests (14 new, 159 total):** the suggestion chain (scheduled / weekday-mode /
  repeat-last / none / tie); `formatTarget`; `prepareRepeatLast` (working sets
  only); and a real-SQLite `programRepository` suite — single-active enforced,
  template resolution, **editing a template never rewrites a past workout**,
  **deleting a template keeps the workout and nulls its provenance (SET NULL)**,
  and recent-template-use reporting.

## The invariant, enforced (the standing rule)

- **Workouts are self-contained.** `saveCompletedWorkout` copies exercises and sets
  into `workout_exercises`/`sets`; `template_id` is the *only* link back, and it is
  provenance. The template-edit test proves a completely rewritten template leaves
  the performed workout byte-for-byte the same.
- **Delete is SET NULL, never cascade.** Removing a template drops the plan and
  clears the workout's `template_id` — the session, its exercises and its sets all
  survive. Verified against real SQLite.

## What changed

New: migration 0004; `data/repositories/programRepository`;
`domain/models/program`; `domain/fitness/templateSuggestion`;
`core/utils/{isoWeekday,weekdayLabel}`; workouts `logic/{startPreparation,
formatTarget}`, `hooks/{usePrograms,useProgram,useTemplateSuggestion,
useStartWorkout}`, components `StartSection`, screens
`ProgramsScreen`/`ProgramDetailScreen`/`TemplateEditorScreen`, routes. Modified:
`useSessionStore` (begin + target/restSeconds/templateId), `sessionDraftSchema`
(persist the new fields), `useRestTimerStore.start` (`defaultSec` seed), `SetRow`
+ `ActiveExerciseCard` (target line + rest seed), `sessionMapping` +
`workoutRepository` (write `template_id`), `WorkoutsHomeScreen` (Start section +
Programs row), `changeBus` (`programs` channel). No frozen document changed.

## Screens affected

Workouts home (smart Start section + Programs entry), Active Workout (target line
per exercise, template rest seeding), and three new program/template screens.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Suggestion chain (scheduled / weekday-mode / repeat-last / none / tie) | domain test | ✅ |
| Single active program enforced | real-SQLite test | ✅ |
| Template resolves its exercises + targets | real-SQLite test | ✅ |
| **Editing a template never rewrites a past workout** | real-SQLite test | ✅ |
| **Delete template → workout kept, `template_id` SET NULL** | real-SQLite test | ✅ |
| Repeat-last reloads working sets only | logic test | ✅ |
| `formatTarget` renders ranges / omits missing / null | logic test | ✅ |
| `npm run check` | typecheck + lint + format + 159 tests + db:check (9 tables) | ✅ green (4× stable) |
| On-device programs walk (build PPL, weekday suggestion, ≤2-tap start, targets visible, rest seed, repeat-last, single-active, template-edit-safe, delete SET NULL, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** for the program/template screens (DB-backed; expo-sqlite
   native-only). Correctness is proven by the repository + domain + logic tests;
   the full programs walk folds into the consolidated TD-001 device pass (checklist
   extended).
2. **Management UI is functional, not fancy** — no drag-reorder of templates or
   exercises (positions follow insertion order), and program archive exists in the
   repository but the UI offers delete only. Neither is a roadmap deliverable; both
   can be added when a real need appears (Rule of Two).
3. **Migration testing** follows the existing approach — `db:check` applies every
   migration to a fresh DB and the SET-NULL/edit tests exercise the resulting
   schema; there is no separate pre-0004 fixture-upgrade harness (the rebuild is
   the standard drizzle `INSERT … SELECT` copy).
4. **Missed-workout detection** (§3.8), enabled by weekday mapping, is a
   consistency-analytics feature for a later milestone — the data it needs is now
   in place.

## Technical debt

- **TD-004 resolved** (migration 0004 adds the `template_id` FK with SET NULL).
- No new debt introduced.

## Retrospective

**What went well?** The "workout copies everything at save" decision from Phase 4
made the phase's headline invariant free: templates and workouts share no rows, so
"editing a plan can't touch history" is true by construction and the test just
confirms it. The suggestion logic stayed a tiny pure function with the repository
resolving names around it, and `begin` generalized the existing session-start path
so template-start and repeat-last are the same code with different inputs. Adding
`template_id` provenance threaded cleanly through the already-there save mapping.

**What was harder than expected?** Adding a foreign key to an existing table in
SQLite requires the full table-rebuild dance; drizzle-kit generated it correctly
(create `__new_workouts`, copy, drop, rename, `foreign_keys` toggling), and
`db:check` confirmed it applies. The other care point was pre-fill semantics for a
template start: reconcile "the plan says 3 sets" with "last time you did 4" —
resolved by taking the max and padding from history so a plan never *removes* a set
you actually did, and never fabricates reps.

**What should change before the next phase?** Nothing structural. CP-B (the M1
checkpoint) is next: it should confirm the tap-budget on device and that the
session-store-vs-SQLite ownership stayed clean through programs (it did — templates
are read to build a session; only `saveCompletedWorkout` writes history).

## Lessons Learned

- **What surprised you:** the strongest guarantee in the phase needed the least
  code — because history was already self-contained, "a template edit must not
  rewrite the past" required *nothing* beyond not sharing rows, and SET NULL fell
  out of one FK clause. The design bought the invariant three phases ago.
- **What documentation prevented mistakes:** DATABASE §3.3/§3.4 fixed every table
  shape, the `ON DELETE SET NULL` provenance rule, and the §5.4 rebuild pattern;
  UI_UX §5.2 gave the exact suggestion fallback chain and the "all pre-filled,
  never locked" stance; the roadmap's ≤ 2-tap budget shaped the one-tap suggested
  start. FITNESS_DOMAIN §3.4's single-side/counting rules kept template-started
  unilateral exercises consistent with logged ones.
- **What should be reused:** `begin(prepared[])` as the single pre-built-session
  entry point (any future "start from X" builds `PreparedExercise[]`); the
  pure-suggestion + repository-resolves-names split; take-max-and-pad for
  reconciling a plan with history without deleting or fabricating; the
  transaction-guarded single-active pattern for any "only one may be active" flag.
- **What should be avoided:** ever writing derived plan state into a workout;
  cascade-deleting history when a plan is removed (SET NULL is the rule);
  fabricating target reps into pre-filled sets (pad with real history or leave 0).
- **Amendment proposals:** none — no frozen-document defect surfaced. TD-004 is
  resolved; no new debt.
