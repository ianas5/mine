# Phase 3 — Exercise Catalog

**Closed:** 2026-07-08 · **Tag:** `v0.4.0-phase3` · **Milestone:** M1 Training

## What was built

A browsable, extensible exercise library backed by a seeded catalog:

- **Domain taxonomy** (`domain/fitness/taxonomy.ts`): the 11 canonical muscle
  groups + `other`, load types, English labels, and the push/pull/legs +
  upper/lower maps — all from FITNESS_DOMAIN §3.3/§3.4, pure and framework-free.
  `Exercise` domain model added.
- **Schema + migration 0001** (`exercises`, DATABASE §3.2): stable-id catalog,
  case-insensitive unique name (`COLLATE NOCASE` index), enum CHECK constraints
  for muscle group and load type (literals inlined — see Lessons), JSON
  `secondary_muscle_groups` (stored, not used in v1), `is_custom`/`is_archived`.
- **Seed library** (`data/seed/exercises.ts`): **110** machine/cable-rich
  hypertrophy exercises across all groups, stable `ex_seed_<slug>` ids;
  `seedDatabase` is idempotent (`onConflictDoNothing` on id) and runs in the
  DB-ready gate after migrations (injected via the new `afterMigrate` prop).
- **`exerciseRepository`** (DATABASE §7): `listActive`/`listArchived` (sorted),
  `createCustom` (uuid via expo-crypto, typed `DuplicateExerciseNameError` on
  collision), `archive`/`unarchive`, and `remove` restricted to custom rows
  (seed rows can only be archived, never deleted). Emits the `exercises`
  change-bus event on every write.
- **Exercise Library screen** (Workouts tab): search, Active/Archived toggle,
  muscle-grouped `SectionList` (virtualized for the 100+ rows), per-exercise
  actions sheet (archive/unarchive; delete-with-Dialog for custom), and the
  **custom-exercise sheet** — the project's first **React Hook Form + Zod**
  form (name/group/load-type/unilateral), establishing the form pattern.
- **Tests (15 new, 70 total):** repository against real SQLite (seed
  idempotency, 100+ count, sort, custom create, case-insensitive duplicate
  rejection, archive round-trip, custom-delete vs seed-archive-only); pure
  `groupExercises`; the Zod schema; and a **seeded screen render test** that
  mounts `WorkoutsScreen` and verifies grouping + live search.

## What changed

New: `domain/fitness/*`, `domain/models/exercise.ts`, `data/seed/*`,
`data/id.ts`, `data/mappers/exerciseMapper.ts`,
`data/repositories/exerciseRepository.ts`, `features/workouts/{schemas,logic,
hooks,components}`, migration `0001`. Modified: `data/schema/tables.ts`,
`core/db` (change-bus `exercises`, gate `afterMigrate`), `app/_layout`
(seeder injection), `WorkoutsScreen`, three primitive prop types
(`error`/`title`/`cta` widened to `| undefined` for exactOptionalPropertyTypes),
jest setup (expo-crypto mock), deps (expo-crypto, zod, react-hook-form,
@hookform/resolvers). No frozen document changed (the two approved §3.3/§4
amendments landed in the preceding docs commit).

## Screens affected

Workouts tab (placeholder → full Exercise Library) + its two sheets.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Seeds 110 exercises; idempotent | repo test (run seed twice → count stable ≥100) | ✅ |
| Active excludes archived; sorted by name | repo test | ✅ |
| Create custom; hydrated correctly | repo test | ✅ |
| Case-insensitive duplicate rejected | repo test (`My Lift` vs `my lift`) | ✅ |
| Archive hides from active, appears in archived, unarchive restores | repo test | ✅ |
| Custom hard-deletes; seed only archives (archive-not-delete) | repo test | ✅ |
| Library renders grouped; search filters live | **seeded screen render test** (RNTL) | ✅ |
| Migration 0001 applies from zero | `db:check` + harness | ✅ |
| Boundary lint (domain purity) still enforced | deliberate violation → caught → removed | ✅ |
| `npm run check` | typecheck + lint + format + 70 tests + db:check | ✅ green |
| On-device visual walk (browse/search/create/archive, 100+ no jank, both themes, relaunch) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots for this (or any future DB-backed) screen:** expo-sqlite
   is native-only, so on web the DB-ready gate shows its error state instead of
   the library. Per the owner directive, visual verification of DB-backed
   screens is consolidated into the single on-device TD-001 pass (its checklist
   now includes the full Exercise Library walk). Behavioral proof is the seeded
   RNTL render test; the library's visual atoms are all Phase-1
   gallery-verified primitives.
2. Feature code can still import a `data/schema` table object without tripping
   lint (the SQL-execution ban catches drizzle/expo-sqlite, not the table const).
   Repository-only access remains a convention here; not worth a bespoke rule yet.
3. `remove()`'s "unreferenced" guard is currently just the `is_custom` check;
   the FK `RESTRICT` half becomes real when `workout_exercises` lands (Phase 4).

## Technical debt introduced

None new. TD-001's consolidated on-device checklist extended with the Exercise
Library walk + 100+ scroll performance.

## Retrospective

**What went well?** The repository + seed ran green against real SQLite on the
first full run; the RNTL seeded-screen test turned out to be a *better* proof
than a screenshot (it asserts behavior, not just appearance) and needs no device.
RHF + Zod slotted into the Sheet primitive cleanly.

**What was harder than expected?** drizzle-kit serialized string enum values in
`CHECK` constraints as `?` bind-parameters (numbers inline, strings don't),
producing broken DDL; fixed with `sql.raw` inlining the trusted enum literals,
after resetting the mis-generated migration to the committed 0000 state.
`exactOptionalPropertyTypes` also surfaced four `undefined`-to-optional-prop
errors once real conditional props appeared — the fix (widen genuinely-optional
props to `T | undefined`; omit undefined in data builders) is now a known pattern.

**What should change before the next phase?** Adopt `sql.raw` (with trusted
constants) as the standard for enum CHECKs. Treat `T | undefined` on optional
props that receive conditional values as the house convention, so it stops
appearing as a late typecheck failure. When generation misfires, reset
migrations to the last committed state before regenerating — never hand-edit.

## Lessons Learned

- **What surprised you:** drizzle-kit's `CHECK` serialization bind-parameterizes
  string literals (invalid in CHECK) while inlining numbers — silent until you
  read the generated SQL. The `@/` path alias *does* resolve inside drizzle-kit
  (via the config's TS loader), so the schema could import domain enums. zod got
  auto-upgraded 3→4 by `@hookform/resolvers`'s peer; v4's `z.enum(readonlyArray)`
  worked without change.
- **What documentation prevented mistakes:** DATABASE §3.2 gave the exact column
  set and the archive-not-delete policy, which the repo + tests encode verbatim;
  FITNESS_DOMAIN §3.3/§3.4 fixed the taxonomy so the seed and CHECK constraints
  share one source (the domain arrays feed the migration). CODING_STANDARDS'
  RHF-for-forms rule meant the first form set the durable pattern rather than an
  ad-hoc one.
- **What should be reused:** the seeded-DB RNTL screen test as the standard
  visual/behavioral proof for DB-backed screens in this environment; the
  better-sqlite3 harness; `sql.raw` enum-CHECK helper (`oneOf`).
- **What should be avoided:** hand-editing generated migrations; passing
  `undefined` to bare optional props; assuming web can render anything below the
  DB boundary.
- **Amendment proposals:** none — no frozen-document defect surfaced. (The
  screenshot deferral is an environment constraint handled by the owner's
  consolidated-device-pass directive, not a doc defect.)
