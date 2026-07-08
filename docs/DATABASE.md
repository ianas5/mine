# DATABASE.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Owner concern:** persistence — schema, relationships, indexes, migrations, backup/import/export, repository boundaries · **Depends on:** FITNESS_DOMAIN.md, ARCHITECTURE.md.
>
> This document owns **how data is durably stored and evolved**. Meaning of the data is owned by FITNESS_DOMAIN; where each kind of state lives is owned by ARCHITECTURE (§3/§6); this document realizes the SQLite side of that contract. Anything not stored here (MMKV prefs, Zustand session state) is out of scope by design.

---

## 1. Persistence Stack

| Piece | Choice | Role |
|---|---|---|
| Engine | **SQLite** via `expo-sqlite` | On-device relational store; single source of truth. |
| Schema & queries | **Drizzle ORM** | Typed schema defined in TypeScript (`src/data/schema/`), typed query builder used inside repositories. |
| Migrations | **`drizzle-kit`** generated SQL files | Checked into the repo; executed at startup by the migration runner in `core/db` (the DB-ready gate from ARCHITECTURE §8/§12). |
| Files | `expo-file-system` | Progress-photo bytes + backup archives. Not SQLite's job. |

**Connection policy:** one shared connection, opened once at startup. PRAGMAs set on open: `journal_mode = WAL`, `foreign_keys = ON`, `busy_timeout = 5000`. All multi-row writes run in transactions (ARCHITECTURE rule 11).

---

## 2. Global Schema Conventions

1. **Primary keys:** `id TEXT` — app-generated **UUID v4** (`expo-crypto` `randomUUID()`). Never auto-increment integers (stable across export/import).
2. **Domain dates** (workout date, nutrition date, snapshot date, photo date): `TEXT` in `YYYY-MM-DD` — the *device-local calendar date* per FITNESS_DOMAIN §2.3. Sorting is lexicographic = chronological.
3. **Audit timestamps** (`created_at`, `updated_at`): `INTEGER` Unix epoch **milliseconds** (UTC). Audit only — never used in domain math.
4. **Weekdays:** `INTEGER 0–6`, **0 = Monday** (ISO), matching the Monday week start.
5. **Booleans:** `INTEGER` 0/1.
6. **Enums:** lowercase `TEXT` values, validated by Zod at the edge and by `CHECK` constraints in SQL. Canonical enum values are listed per table below.
7. **Units are metric and implicit in the column** (kg, cm, ml, kcal, g per FITNESS_DOMAIN §2.1). No unit columns, ever.
8. **Numbers:** `REAL` for kg/cm/g/%/RPE; `INTEGER` for kcal, ml, reps, counts. Store raw precision; round only in the UI (FITNESS_DOMAIN §2.2).
9. **Nullability encodes "not recorded"** (FITNESS_DOMAIN §2.4): a nullable column left `NULL` means *not measured/logged*, and `0` is a real value. Columns are `NOT NULL` only when the domain guarantees presence.
10. **No derived tables.** PRs, trends, adherence, insights are never persisted in SQLite (ARCHITECTURE rule 8). Disposable caches belong to MMKV.
11. **Soft-delete via `is_archived`** for catalog entities referenced by history (exercises, foods). True `DELETE` only for entities history doesn't reference or where cascade is intended.

---

## 3. Schema

Sixteen tables in six groups. Format: `column TYPE [constraints] — meaning`.

### 3.1 Settings & configuration

#### `settings` — single-row domain configuration (ARCHITECTURE §6)

```
id                     INTEGER PK CHECK (id = 1)      — enforced single row
weekly_workout_target  INTEGER NOT NULL DEFAULT 4     — FITNESS_DOMAIN §3.8
default_bodyweight_kg  REAL NULL                      — fallback for bodyweight loads (§3.4)
height_cm              REAL NULL                      — for derived BMI (§5.2)
water_cup_ml           INTEGER NOT NULL DEFAULT 250   — water increment (§4.1)
created_at, updated_at INTEGER NOT NULL
```

One row, created by the initial migration. Upsert-only; never deleted. **In backups.** (Theme, last tab, etc. are MMKV — not here.)

#### `nutrition_targets` — time-versioned targets (FITNESS_DOMAIN §4.1)

```
id             TEXT PK
effective_from TEXT NOT NULL UNIQUE  — YYYY-MM-DD; the date this target set becomes active
kcal           INTEGER NOT NULL
protein_g      REAL NOT NULL
carb_g         REAL NOT NULL
fat_g          REAL NOT NULL
water_ml       INTEGER NULL          — optional water target
created_at, updated_at INTEGER NOT NULL
```

**Resolution rule (canonical, implemented once in `nutritionRepository`):** the active target for date *D* is the row with the greatest `effective_from ≤ D`. If none exists (dates before the first target), adherence for those days is `insufficient-data` — never a fabricated default.
**Decision:** we do **not** additionally snapshot targets onto nutrition days. Resolution via `effective_from` is deterministic and single-sourced; the snapshot option FITNESS_DOMAIN allows is unnecessary. Editing a target row rewrites what that historical target *was* — allowed, deliberate, and the user's own act; normal target *changes* are new rows.
Index: `UNIQUE(effective_from)` (doubles as the lookup index).

#### `phases` — user-defined training phases (ANALYTICS_ENGINE §5.4)

```
id         TEXT PK
name       TEXT NOT NULL
type       TEXT NOT NULL CHECK IN ('cutting','recomp','lean_bulk','maintenance','custom')
start_date TEXT NOT NULL            — YYYY-MM-DD
end_date   TEXT NULL                — NULL = ongoing; CHECK (end_date IS NULL OR end_date >= start_date)
notes      TEXT NULL
created_at, updated_at INTEGER NOT NULL
```

**No-overlap invariant** enforced in `phaseRepository` within a transaction; at most one `end_date IS NULL` row. Index: `(start_date)`. **In backups.**

### 3.2 Exercise catalog

#### `exercises`

```
id                       TEXT PK                       — seed rows use stable ids: 'ex_seed_<slug>'
name                     TEXT NOT NULL UNIQUE COLLATE NOCASE
primary_muscle_group     TEXT NOT NULL CHECK IN
   ('chest','shoulders','back','biceps','triceps','forearms',
    'core','glutes','quads','hamstrings','calves','other')   — FITNESS_DOMAIN §3.3
secondary_muscle_groups  TEXT NOT NULL DEFAULT '[]'    — JSON array of the same enum; stored for v2, NOT used in v1 volume math
load_type                TEXT NOT NULL DEFAULT 'external' CHECK IN
   ('external','bodyweight','bodyweight_plus','assisted','timed')  — §3.4
default_unilateral       INTEGER NOT NULL DEFAULT 0    — hint only; prefills the entry-level marker (§3.4 below)
is_custom                INTEGER NOT NULL DEFAULT 0    — 0 = seed library, 1 = user-created
is_archived              INTEGER NOT NULL DEFAULT 0    — soft-delete; archived rows hidden from pickers, kept for history
notes                    TEXT NULL
created_at, updated_at   INTEGER NOT NULL
```

**Deletion policy:** an exercise referenced by any `workout_exercises` or `template_exercises` row can only be **archived**. Hard delete is allowed only when unreferenced (enforced by FK `RESTRICT` + repository check).

### 3.3 Programs & templates

#### `programs`

```
id          TEXT PK
name        TEXT NOT NULL
notes       TEXT NULL
is_active   INTEGER NOT NULL DEFAULT 0   — at most one active program (enforced in repository transaction)
is_archived INTEGER NOT NULL DEFAULT 0
created_at, updated_at INTEGER NOT NULL
```

#### `templates` — a single-session blueprint

```
id          TEXT PK
program_id  TEXT NULL REFERENCES programs(id) ON DELETE CASCADE   — NULL = standalone template
name        TEXT NOT NULL
position    INTEGER NOT NULL DEFAULT 0    — order within program
weekday     INTEGER NULL CHECK (0–6)      — optional schedule day (0 = Monday); enables "missed workout" (§3.8)
notes       TEXT NULL
is_archived INTEGER NOT NULL DEFAULT 0
created_at, updated_at INTEGER NOT NULL
```

#### `template_exercises` — targets, not performance

```
id             TEXT PK
template_id    TEXT NOT NULL REFERENCES templates(id) ON DELETE CASCADE
exercise_id    TEXT NOT NULL REFERENCES exercises(id) ON DELETE RESTRICT
position       INTEGER NOT NULL
target_sets    INTEGER NULL
target_rep_min INTEGER NULL
target_rep_max INTEGER NULL
target_rpe     REAL NULL
rest_seconds   INTEGER NULL       — rest-timer default; UI convenience only (§3.6)
notes          TEXT NULL
```

Index: `(template_id, position)`.

### 3.4 Training history

#### `workouts` — completed sessions only (drafts live in `workout_drafts`)

```
id          TEXT PK
date        TEXT NOT NULL                — YYYY-MM-DD; multiple per date allowed (§2.5)
name        TEXT NOT NULL                — e.g. 'Push Day'
template_id TEXT NULL REFERENCES templates(id) ON DELETE SET NULL   — provenance only
started_at  INTEGER NULL                 — epoch ms; both present → duration
ended_at    INTEGER NULL
notes       TEXT NULL
created_at, updated_at INTEGER NOT NULL
```

Index: `(date)`. Duration is derived (`ended_at − started_at`), never stored.

#### `workout_exercises` — one exercise as performed in one workout

```
id                  TEXT PK
workout_id          TEXT NOT NULL REFERENCES workouts(id) ON DELETE CASCADE
exercise_id         TEXT NOT NULL REFERENCES exercises(id) ON DELETE RESTRICT
position            INTEGER NOT NULL
unilateral_counting TEXT NOT NULL DEFAULT 'none' CHECK IN ('none','single_doubled','per_side')
notes               TEXT NULL
```

**`unilateral_counting` is the explicit, unambiguous marker required by FITNESS_DOMAIN edge case #6:**

| Value | Meaning | Aggregation rule |
|---|---|---|
| `none` | bilateral movement | count as logged |
| `single_doubled` | one side logged once | volume & working-set stimulus **×2**; PRs use single-side load |
| `per_side` | each side logged as its own set | **no doubling**; PRs use single-side load |

The doubling decision is stored **on the data**, not inferred at computation time — analytics can never double-guess.
Indexes: `(workout_id, position)`, `(exercise_id)` (exercise-history and PR queries).

#### `sets`

```
id                  TEXT PK
workout_exercise_id TEXT NOT NULL REFERENCES workout_exercises(id) ON DELETE CASCADE
position            INTEGER NOT NULL
weight_kg           REAL NOT NULL DEFAULT 0   — meaning depends on exercise load_type (FITNESS_DOMAIN §3.4):
                                              — external: full load · bodyweight_plus: ADDED load
                                              — assisted: assist amount · bodyweight/timed: 0
reps                INTEGER NOT NULL DEFAULT 0 — for 'timed' exercises this holds SECONDS (edge case 8)
rpe                 REAL NULL CHECK (0–10)
rir                 INTEGER NULL CHECK (0–10)
is_warmup           INTEGER NOT NULL DEFAULT 0
notes               TEXT NULL
```

Index: `(workout_exercise_id, position)`. Working-set classification (FITNESS_DOMAIN §3.2) is **computed**, not stored.

#### `workout_drafts` — crash-safe active-session checkpoint (ARCHITECTURE §7.1)

```
id         INTEGER PK CHECK (id = 1)   — at most one active session
payload    TEXT NOT NULL               — JSON snapshot of the Zustand session (Zod-validated on resume)
updated_at INTEGER NOT NULL
```

Deliberately a **JSON blob, not relational**: it's a recovery checkpoint of ephemeral state, rewritten on debounce/background, deleted on finish/discard. On launch, a valid draft triggers Resume/Discard; an unparseable draft is discarded gracefully. **Excluded from backups** — it is not history.

### 3.5 Nutrition

#### `foods` — reusable foods and quick meals (FITNESS_DOMAIN §4.1)

```
id              TEXT PK              — seed rows: 'food_seed_<slug>'
name            TEXT NOT NULL
serving_amount  REAL NOT NULL        — e.g. 100
serving_unit    TEXT NOT NULL CHECK IN ('g','ml','piece','scoop','cup','serving')
kcal            INTEGER NOT NULL     — per serving
protein_g       REAL NOT NULL        — per serving, 0.1 g precision
carb_g          REAL NOT NULL
fat_g           REAL NOT NULL
is_quick_meal   INTEGER NOT NULL DEFAULT 0  — composed meals ('Chicken and rice') vs single foods
is_custom       INTEGER NOT NULL DEFAULT 1
is_archived     INTEGER NOT NULL DEFAULT 0
created_at, updated_at INTEGER NOT NULL
```

No fiber column (FITNESS_DOMAIN §4). Duplicates allowed (edge case #15); repository may surface same-name matches for reuse.

#### `meal_entries` — the log; **macros snapshotted at log time** (§4.1)

```
id            TEXT PK
date          TEXT NOT NULL          — YYYY-MM-DD
slot          TEXT NULL CHECK IN ('breakfast','lunch','dinner','snacks')
food_id       TEXT NULL REFERENCES foods(id) ON DELETE SET NULL  — provenance; history survives food deletion
food_name     TEXT NOT NULL          — snapshot
logged_amount REAL NOT NULL
logged_unit   TEXT NOT NULL          — snapshot of serving_unit
kcal          INTEGER NOT NULL       — snapshot, scaled per §4.2
protein_g     REAL NOT NULL
carb_g        REAL NOT NULL
fat_g         REAL NOT NULL
logged_at     INTEGER NOT NULL       — epoch ms (ordering within a day)
```

**Snapshot rule:** editing a food definition never rewrites past `meal_entries`. Day totals are `SUM(...) GROUP BY date` — no stored day-total table.
Indexes: `(date)`, `(food_id)`.

#### `water_days` — one row per date (§2.5 daily-keyed)

```
date       TEXT PK        — YYYY-MM-DD
ml         INTEGER NOT NULL DEFAULT 0   — cumulative; 0 is a real logged value, absent row = unlogged
updated_at INTEGER NOT NULL
```

### 3.6 Body & photos

#### `body_snapshots` — one per date, field-merged (FITNESS_DOMAIN §5.1)

```
date            TEXT PK    — YYYY-MM-DD; the domain's one-snapshot-per-date rule, in the schema itself
weight_kg       REAL NULL
body_fat_pct    REAL NULL
muscle_mass_kg  REAL NULL
visceral_fat    REAL NULL          — index, unitless
bmi             REAL NULL          — entered value; derived-from-height only when NULL (§5.2)
neck_cm, chest_cm, waist_cm, hips_cm            REAL NULL
left_arm_cm, right_arm_cm                       REAL NULL
left_forearm_cm, right_forearm_cm               REAL NULL
left_thigh_cm, right_thigh_cm                   REAL NULL
left_calf_cm, right_calf_cm                     REAL NULL
created_at, updated_at INTEGER NOT NULL
```

**Merge-upsert contract (repository-enforced):** saving a snapshot updates **only the fields present in the input**; omitted fields keep their stored values; clearing a field is an explicit `NULL` assignment through a dedicated "clear" action. Bilateral sites are always per-side columns — never collapsed (§5.1).

#### `progress_photos` — metadata only; bytes on the filesystem (ARCHITECTURE §12)

```
id         TEXT PK
date       TEXT NOT NULL
angle      TEXT NOT NULL CHECK IN ('front','side','back')
file_name  TEXT NOT NULL UNIQUE      — relative name under <documentDirectory>/photos/, e.g. '2026-07-08_front_<id>.jpg'
width, height INTEGER NULL           — cached dimensions for layout
notes      TEXT NULL
created_at INTEGER NOT NULL
```

Index: `(date)`. **File lifecycle is transactional in spirit:** write file → insert row (on insert failure, delete file); delete row → delete file. An orphan sweep (files without rows, rows without files) runs opportunistically at startup; rows whose file is missing render a "missing photo" state, never crash.

### 3.7 Entity-relationship overview

```
programs 1─* templates 1─* template_exercises *─1 exercises
workouts 1─* workout_exercises *─1 exercises
workout_exercises 1─* sets
foods 1─* meal_entries                    (SET NULL + snapshot)
nutrition_targets ── resolved by effective_from ──> meal-entry dates
phases (by date range) · body_snapshots (by date) · water_days (by date) · progress_photos (by date)
settings (single row) · workout_drafts (single row, ephemeral)
```

---

## 4. Index Summary

Beyond primary keys and stated `UNIQUE`s:

| Index | Serves |
|---|---|
| `workouts(date)` | range windows (7/30/90/180/365d), calendar, consistency |
| `workout_exercises(workout_id)` / `(exercise_id)` | session loads / exercise history & PR scans |
| `sets(workout_exercise_id)` | session + history loads |
| `template_exercises(template_id)` | template loads |
| `templates(program_id)` | program loads |
| `meal_entries(date)` / `(food_id)` | day view, nutrition ranges / food usage |
| `progress_photos(date)` | photo timeline & comparisons |
| `nutrition_targets(effective_from)` (unique) | target resolution |
| `phases(start_date)` | phase lookup & overlap checks |

Add further indexes only against a measured slow query — not speculatively.

---

## 5. Migration Strategy

1. **Toolchain:** schema lives in `src/data/schema/*.ts` (Drizzle). `drizzle-kit generate` produces numbered SQL files (`src/data/schema/migrations/0000_*.sql`, …) — **checked into git, immutable once committed.**
2. **Runtime:** the migration runner (`core/db`) applies pending migrations **inside the DB-ready gate**, in order, each in a transaction, before any feature renders. Drizzle's journal table records applied migrations.
3. **Forward-only.** No down-migrations. Recovery from a bad migration is a code fix in a *new* migration.
4. **Additive-first.** Prefer `ADD COLUMN` (nullable or defaulted) and new tables. Renames/drops/type changes use SQLite's **table-rebuild pattern** (create new → copy → drop → rename) inside one transaction, with `foreign_keys` handling per SQLite docs.
5. **User data is sacred:** a migration that would lose data is forbidden; if a column is retired, its data is migrated or deliberately, documentedly abandoned. Any destructive change requires amending this document first (AI rule 10).
6. **Seeding is not migration.** After migrations, an idempotent seeder inserts the seed exercise library (FITNESS_DOMAIN §3.3 taxonomy) and starter foods using **stable ids** (`ex_seed_*`, `food_seed_*`) with insert-if-absent semantics. Re-running never duplicates; user edits to seed rows are never overwritten. Seed content versioning: new seed items ship as new stable ids.
7. **Testing:** every migration PR runs against (a) an empty DB and (b) a fixture DB representing the previous version with realistic data (process detail in DEVELOPMENT_WORKFLOW).

---

## 6. Backup / Export / Import

**Format — one archive:** `fitness-backup-YYYY-MM-DD.zip`

```
data.json      — the versioned document
photos/        — the referenced image files (by file_name)
```

`data.json` envelope:

```jsonc
{
  "app": "personal-fitness-tracker",
  "format": 1,                  // backup-format version
  "schemaVersion": <n>,         // DB migration version at export
  "exportedAt": "<ISO-8601>",
  "data": {
    "settings": {...}, "nutritionTargets": [...], "phases": [...], "exercises": [...],
    "programs": [...], "templates": [...], "templateExercises": [...],
    "workouts": [...], "workoutExercises": [...], "sets": [...],
    "foods": [...], "mealEntries": [...], "waterDays": [...],
    "bodySnapshots": [...], "progressPhotos": [...]
  }
}
```

**Included:** every table above **except `workout_drafts`** (ephemeral recovery state) — and MMKV is never included (disposable by definition, ARCHITECTURE §3).

**Export path:** `data/backup` service → reads all tables (it lives inside the data layer and may use Drizzle directly for the full dump; the repositories-only-SQL rule applies to *features*, and `data/backup` is within the persistence boundary) → writes archive via `expo-file-system` → OS share sheet.

**Import path (all-or-nothing):**
1. Pick archive → extract to a temp directory.
2. **Zod-validate `data.json` in full** against the schema for its declared `format`/`schemaVersion` (the highest-risk edge, ARCHITECTURE §10). Any failure → abort, nothing touched.
3. If `schemaVersion` is older than current, run **data-shape upgraders** (pure functions mirroring the SQL migrations' semantics) to current shape. A *newer* version than the app supports → refuse with a clear message ("update the app first").
4. **Replace, don't merge:** within a single transaction, delete all rows from domain tables and insert the imported rows. (A single-user personal tool restoring a backup means "make my data this." Merge semantics are ambiguous and dangerous — explicitly out of scope for v1.)
5. Photos: copy `photos/` into the document directory; then reconcile (rows without files → flagged missing; files without rows → removed).
6. On any failure at any step: transaction rolls back, temp files discarded, existing data untouched.
7. **Safety export before replacement:** before replacing any data, the app **attempts** an automatic safety export of the current database (same archive format, named `pre-import-safety-…zip`). If the safety export **fails for any reason, the import pauses and the user must explicitly confirm** ("Continue without safety backup?") before replacement proceeds. Silent continuation is forbidden.

---

## 7. Repository Boundaries

One repository per aggregate (ARCHITECTURE §14); repositories are the **only** feature-reachable SQL surface. All return `domain/models` types via `data/mappers` — never raw rows.

| Repository | Tables owned | Notable contract points |
|---|---|---|
| `settingsRepository` | `settings` | get/update single row; creates row if missing. |
| `exerciseRepository` | `exercises` | list (active/archived), create custom, archive; hard-delete only when unreferenced. |
| `programRepository` | `programs`, `templates`, `template_exercises` | CRUD; enforces single `is_active` program in a transaction. |
| `workoutRepository` | `workouts`, `workout_exercises`, `sets`, `workout_drafts` | `saveCompletedWorkout(session)` — one transaction: insert workout tree + delete draft. `checkpointDraft` / `loadDraft` / `discardDraft`. History queries by date range & by exercise (feeds PR/analytics computation). |
| `foodRepository` | `foods` | CRUD, archive, same-name reuse surfacing. |
| `nutritionRepository` | `meal_entries`, `water_days`, `nutrition_targets` | Day view (entries + water + resolved target), range queries for analytics, **owns the single implementation of target resolution** (§3.1). |
| `bodyRepository` | `body_snapshots` | **Owns the merge-upsert** (§3.6) and explicit field-clear; date-range queries. |
| `photoRepository` | `progress_photos` | Metadata CRUD **coordinated with file IO** (write-file-then-row; delete-row-then-file); startup orphan sweep. |
| `phaseRepository` | `phases` | CRUD; enforces no-overlap and single-ongoing invariants in a transaction. |
| `backupService` (`data/backup`) | all (read), all (replace on import) | Export/import per §6. Not a repository; a data-layer service. |

Change-bus contract: every write emits the owning table name(s) (ARCHITECTURE §7); `saveCompletedWorkout` emits `workouts` (subscribers of workout data re-query; derived metrics recompute per ARCHITECTURE rule 8).

---

## 8. Data Integrity Rules (summary of the contracts above)

1. FK enforcement always on; cascades exactly as declared in §3 — history trees (`workout_exercises`, `sets`, `template_exercises`) cascade with their parent; catalog references (`exercises`, `foods`) never cascade into history (`RESTRICT` / `SET NULL` + snapshot).
2. Daily-keyed tables (`body_snapshots`, `water_days`) use the date as PK — duplicates are structurally impossible.
3. Meal macros and food names are snapshots; catalog edits never rewrite the log.
4. Plausibility ranges (FITNESS_DOMAIN §8.13) are enforced at the Zod edge; `CHECK` constraints carry only structural invariants (enums, 0–10 RPE, weekday 0–6).
5. No derived values in SQLite. If it can be recomputed from these tables, it is.

---

## 9. AI Decision Rules (Database)

1. **Schema changes only via migration.** Never mutate schema ad hoc, never edit a committed migration file, never renumber. New change → new migration.
2. **Never write SQL outside `data/`.** Features use repositories; `data/backup` is the only other SQL surface, for dump/restore only.
3. **Snapshots protect history.** When logging references a catalog entity (food → meal entry), copy the values the log depends on. Editing catalogs must never change what history says happened.
4. **Archive, don't delete, referenced catalog rows.** `is_archived = 1` hides; only unreferenced rows may be hard-deleted.
5. **Daily-keyed = date PK + merge-upsert.** Body snapshots and water days: update provided fields only; omission never nulls a stored value; clearing is explicit.
6. **`unilateral_counting` travels with the data.** Aggregation code reads the stored marker; it never infers doubling from exercise names or heuristics.
7. **Store raw metric values, no unit columns, no imperial.** kg/cm/ml/kcal/g as specified; rounding is a UI concern.
8. **Drafts are not history.** `workout_drafts` is a single JSON checkpoint: excluded from backups, deletable without data loss, Zod-validated before resume, discarded gracefully when invalid.
9. **No derived/cache tables in SQLite.** PRs, trends, adherence, insights: recompute (domain layer) or cache in MMKV. If you're adding a table whose rows could be recomputed, stop.
10. **Destructive migrations are forbidden** without amending this document first and providing a data-preservation (or explicit-abandonment) plan.
11. **Import is all-or-nothing, replace-not-merge, and always preceded by an attempted automatic safety export; on safety-export failure, explicit user confirmation is required before replacement.** Never partially apply an import; never merge silently.
12. **Photos: file first, row second; row gone, file gone.** Never store image bytes in SQLite; never leave the metadata pointing at nothing without the "missing" state handling it.
13. **Target resolution has one implementation** (in `nutritionRepository`). Never re-derive "active target for date D" elsewhere.
14. **New indexes require a demonstrated slow query.** Don't decorate the schema speculatively.

---

## Changelog

- 2026-07-08 — v1 baseline frozen (safety-export wording refinement applied; `phases` table + `phaseRepository` added per approved ANALYTICS_ENGINE amendment; table count corrected to sixteen).
