# CP-C — Architecture Verification Checkpoint (end of M2)

**Date:** 2026-07-09 · **Covers:** Phases 9–14 (Milestone 2 · Fuel & Body) · **Type:** review only (no features).

Per IMPLEMENTATION_ROADMAP §3, verified against the frozen documents with them open.
This checkpoint also **opens the daily-use gate**: from here the owner may begin real
logging, and every migration-bearing install follows the backup-first rule with real
stakes. Device-only items remain **non-blockers** (owner directive) and accumulate into
the single consolidated TD-001 on-device pass.

**Gate at review:** `npm run check` green — typecheck, boundary lint
(`--max-warnings 0`), prettier, **238 tests / 56 suites**, `db:check` (15 tables).
Suite stable across repeated full runs (the Phase-14 native-constraint-in-transaction
flake was resolved by driving fault injection through a deterministic JS-layer throw;
no recurrence over 3× full runs).

Roadmap CP-C special attention — **snapshot rules · target resolution single-sourced ·
photo lifecycle · no derived tables** — each verified below.

## 1. PRODUCT_PRINCIPLES — the Tests

- **No vanity metrics.** Every M2 number is real and derived: nutrition day totals are
  a live `sumMacros` reduce over logged entries; adherence/remaining resolve through
  the one target path; body deltas and comparisons come from stored snapshots. ✅
- **No fabricated data (P8).** A day before the first target reads "no target set"
  (insufficient-data, not a default); a body field on only one compared date shows
  "—", never a guessed baseline; a photo whose file is missing renders a "File
  missing" placeholder, never a broken image or silent drop. ✅
- **No dark patterns.** Meal-entry delete is immediate + a 5 s Undo toast; the
  destructive import is a plain confirm, no pressure. No streaks-as-pressure, no nags. ✅
- **Logging speed (P9).** Log Meal opens on Recent & Frequent with last-used portion +
  time-of-day slot; ≤ 3-tap repeat meal; Add Weight pre-set to last weight. Tap *feel*
  is device-verified → TD-001. ✅ (mechanics), ⏳ (device)

## 2. FITNESS_DOMAIN — formulas cite and match

- **Macros / Atwater / portion scaling** (§4.2) — `domain/nutrition/macros` with
  edge tests incl. empty-day zero. ✅
- **Time-versioned targets** (§4.1) — the active target for a date is the greatest
  `effective_from ≤ date`; resolved **once** (see §4). ✅
- **Adherence + protein floor vs calorie/carb/fat bands** (§4) — `domain/nutrition`
  with boundary tests. ✅
- **BMI entered-else-derived; per-side bilateral; directionality + §6.4 deadband**
  (§5.1–5.4) — `domain/body` snapshot + comparison with the named edge tests
  (incomparable ⇒ "—", waist-down improving, arm-down declining, weight/BMI neutral,
  %-undefined at A=0). ✅
- **Oldest-missing capture angle** (UI_UX §5.2) — `domain/photos` with tests. ✅

## 3. ARCHITECTURE — boundaries & state ownership

- **Boundary lint clean** (`eslint-plugin-boundaries`, `--max-warnings 0`), and a
  static sweep confirms **zero feature→feature imports** across `src/features`. ✅
- **Vertical slice read — nutrition day:** route `app/(tabs)/nutrition/index`
  (re-export) → `NutritionScreen` (imports the `useNutritionDay` hook only, no repo /
  drizzle / SQL) → `useNutritionDay` (calls `nutritionRepository` only) →
  `nutritionRepository` (the sole SQL surface, 28 query sites). Layering holds end to
  end. ✅
- **Cross-feature reads via the data layer, not the owning feature.** Nutrition and
  measurements read settings (height/bodyweight) through local repository-backed hooks,
  never by importing the settings feature. ✅
- **State ownership (§6).** Domain data in SQLite; theme/last-tab in MMKV (excluded
  from backups); the active session in the Zustand store checkpointed to
  `workout_drafts` (excluded from backups — not history). ✅
- **`data/backup` is a data-layer service, not a repository** — it may use Drizzle
  directly for the full dump/restore; the repositories-only rule binds features, and
  the zip/share/picker I/O is behind the injected `ArchiveStore` so `core`/`data` stay
  free of native imports (wired at the composition root, like the DB handle and photo
  store). ✅

## 4. DATABASE — schema, migrations, no derived tables

- **No derived tables crept in.** The 15 physical tables are all base data — no
  `personal_records`, no cached day totals, no adherence cache. PRs recompute from
  history (M1); day totals are a runtime reduce; targets resolve by query. ✅
- **Target resolution single-sourced.** `resolveTargetForDate` exists **only** in
  `nutritionRepository`; every consumer (day hook, targets editor) calls it; there is
  no competing date-based target math anywhere in the tree. ✅
- **Snapshot-at-write.** `meal_entries` store their own `kcal/protein/carb/fat` at log
  time and `food_id` is `SET NULL` on food delete — a food edit/delete never rewrites
  past entries. `body_snapshots` merge-upsert: the `MeasurementPatch` (numbers-only)
  **cannot express a clear**, so omitted fields keep their stored values; clearing is a
  separate explicit `SET … = NULL`. ✅
- **Photo lifecycle (§3.6).** Save writes the file **then** inserts the row (insert
  failure deletes the file); delete removes the row **then** the file; the startup
  sweep deletes rowless files and flags fileless rows missing. The FS/DB pair reconciles
  to a single consistent state. ✅
- **Migrations immutable & consistent.** `db:check` passes; journal 0000–0008 intact,
  additive; no migration edited in place. ✅
- **Schema-in-DB vs docs — one planned delta.** DATABASE.md documents **sixteen**
  tables; the device DB has **fifteen**. The difference is exactly the `phases` table,
  a documented-but-future Phase 19 deliverable. Not a defect — see Finding F-C1. ✅ (with note)

## 5. Data-safety guarantees (daily-use-gate readiness)

Verified in code and by 20 backup tests (real SQLite + in-memory doubles):

- **Full validation before any write** — `parseAndUpgradeEnvelope` (JSON → header →
  app/format → version → full Zod) is pure and runs before the safety export; a
  malformed/foreign/typewrong document throws untouched. ✅
- **Version reconciliation before replacement** — newer refused (`schema-too-new`),
  older upgraded, gaps refused (`unsupported-schema`). ✅
- **Replace-not-merge, single transaction** — every table deleted child-first, inserted
  parent-first, inside one `BEGIN/COMMIT`. ✅
- **Rollback leaves data intact** — a mid-transaction failure rolls back deletes *and*
  inserts (deterministic JS-layer fault test). ✅
- **Safety-export gate** — an automatic `pre-import-safety-…zip` is attempted first;
  on failure the import pauses for explicit confirmation (`onSafetyExportFailed`); decline
  aborts untouched, and success never asks. ✅
- **Photo reconciliation is separate from DB integrity** — photos are copied and swept
  only after the row transaction commits; a crash there leaves recoverable orphan files,
  never a corrupt DB. ✅
- **No backup coverage gap.** Schema, `collect`, and `replace` agree on the same 14
  backed-up tables; `workout_drafts` is the only exclusion (ephemeral, correct). A
  round-trip test proves export → mutate → import is byte-equivalent incl. photos. ✅

## 6. Body / nutrition ownership boundaries

- Nutrition owns `foods`, `meal_entries`, `nutrition_targets`, `water_days`; measurements
  own `body_snapshots`, `progress_photos`. Neither imports the other or the settings
  feature; shared config (height/bodyweight) is read via data-layer hooks. ✅
- **TD-009** remains the one acknowledged cross-domain seam (bodyweight-load volume still
  reads `settings.defaultBodyweightKg`, not the latest weigh-in) — deferred to Phase 16,
  the first place body + training analytics meet. On schedule. ✅

## 7. Technical-debt registry audit (§2)

Nine entries; **none due at CP-C**. TD-004 resolved (Phase 8). Future/scheduled:
TD-002 (CP-D), TD-003/008 (Phase 21), TD-005 (unscheduled), TD-006/009 (Phase 16),
TD-007 (Phase 15). **TD-001** (device verification) carries its original CP-A deadline
but was explicitly kept open by owner directive at CP-B; its stakes rise now — see F-C2.
No undocumented shortcuts surfaced in the M2 diff.

---

## Findings

**Verdict: PASS.** M2 meets the frozen documents; the data-safety guarantees and
ownership boundaries hold; no derived tables, single-sourced targets, correct snapshot
and photo-lifecycle semantics. Two non-blocking findings, both observations:

- **F-C1 — Documented `phases` table not yet in the DB (expected).** DATABASE.md lists
  sixteen tables incl. `phases`; the DB has fifteen. This is the planned Phase 19
  deliverable; the backup envelope, `schemaVersion` gate, and (empty) upgrader registry
  are pre-wired so adding it is additive and forward-compatible. **No amendment or
  action** — noting the delta explicitly for the §3.4 "schema == docs" check.

- **F-C2 — TD-001 device verification is now load-bearing for the gate.** Opening the
  daily-use gate makes real-data backup/restore correctness matter on real hardware, yet
  the device-only surfaces (zip pack/unpack, share sheet, document picker, photo bytes,
  and tap budgets) remain unverified under TD-001. The correctness-critical logic is
  proven off-device, but I recommend the owner exercise the **TD-001 backup walk**
  (export → reinstall → import round-trip, corrupt-zip error, safety-export-fail confirm)
  on a physical iPhone before trusting a real restore. Tracked; not a code defect.

Per the checkpoint contract, findings are fixed or become approved doc amendments before
M3 (Phase 15) starts. Neither finding requires a code change; F-C1 needs no amendment
(the delta is already documented as a future table), and F-C2 is a device-pass
recommendation, not a defect.

## Readiness for the daily-use gate

Ready. Backup/restore is validated, all-or-nothing, safety-gated, and covers every
persistent table; the FS/DB transaction and orphan handling are correct; no derived
state can drift because none is stored. The one caveat is F-C2: the device backup walk
should be run before the first real restore, since that path is now real-stakes and
currently proven only off-device.
