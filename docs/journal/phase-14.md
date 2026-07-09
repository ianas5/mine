# Phase 14 — Backup, Export & Import (the daily-use gate)

**Closed:** 2026-07-09 · **Milestone:** M2 (Fuel & Body) → opens the daily-use gate

Full data ownership lands: a one-file `.zip` export of every table plus the photo
bytes, shared through the OS; and an **all-or-nothing** import that validates, checks
the version, takes a safety backup, and replaces everything in a single transaction —
or touches nothing at all. This is a data-safety feature, not a convenience: the
import path is built so that **existing data is never corrupted on failure**.

## The safety contract (DATABASE §6, the standing principle)

Filesystem and database behave as one logical unit, and the import is ordered so the
only outcomes are "fully applied" or "nothing touched":

1. **Extract + validate first.** `data.json` is `JSON.parse`d and Zod-validated in
   full against the current shape. A malformed or foreign document aborts before any
   write — and before the safety export is even attempted.
2. **Version-reconcile.** A newer `schemaVersion` is refused ("update the app first");
   an older one runs data-shape upgraders up to current; a gap with no upgrader is
   refused. Never guessed.
3. **Attempted safety export.** Before replacing anything, the app auto-exports the
   *current* data as `pre-import-safety-…zip`. If that fails for any reason, the
   import **pauses** and the user must explicitly confirm continuing without it
   (§6.7). Silent continuation is impossible by construction.
4. **Single-transaction replace.** Every table is deleted child-first and re-inserted
   parent-first inside one `BEGIN/COMMIT`. Any failure rolls the whole thing back —
   deletes included — leaving the prior data exactly intact (§6.6). Replace, never
   merge.
5. **Photo reconcile last.** Only after the row replace commits are photo files copied
   in and the orphan sweep run (files without rows removed; rows without files flagged
   missing). A crash here leaves recoverable orphan files, never a corrupt DB.

## What was built

- **`backupSchema` (Zod)** — row schemas mirroring all 14 backed-up tables exactly
  (camelCase, matching Drizzle's shape); unknown keys stripped, so only known columns
  are ever inserted and any missing/mistyped field aborts the import. A lenient
  **header** schema is parsed first so version errors are precise.
- **`schemaVersion`** — `CURRENT_SCHEMA_VERSION` (= applied-migration count), guarded
  against drift by a test that asserts it equals the migration-journal length.
- **`upgraders`** — the sequential data-shape upgrade registry (empty at v1; the first
  entry arrives with `phases` in Phase 19); refuses newer/gapped versions.
- **`parseEnvelope`** — the pure validate → app/format check → version-reconcile →
  full-data-validation pipeline; every failure throws a typed `ImportError` and, by
  construction, touches nothing.
- **`collect`** — reads every table into a `BackupData` block (Drizzle directly; the
  service is inside the persistence boundary). `workout_drafts` excluded (ephemeral).
- **`replace`** — the transactional, FK-ordered, chunked delete-all + insert-all.
- **`archiveStore`** — the zip/share/pick/extract/stage I/O boundary as an injectable
  interface (like `PhotoStore`), with a real `expoArchiveStore` (fflate over the
  SDK-57 `File` byte API + expo-sharing + expo-document-picker). The orchestration is
  therefore fully testable with an in-memory double.
- **`backupService`** — `exportAndShare`, `pickArchive`, `importArchive(uri, hooks)`;
  the ordered import flow above, with `onSafetyExportFailed` as the confirmation gate.
- **`ImportError`** — a code union (`unreadable-archive`, `invalid-data`,
  `unsupported-format`, `schema-too-new`, `unsupported-schema`, `aborted-no-safety`)
  each mapped to a specific user message; a thrown `ImportError` always means untouched.
- **Settings → Backup UI** (`BackupSection`) — Export (share sheet) and Import,
  double-guarded by a destructive "Replace all data?" confirmation and, on safety
  failure, an explicit "continue without safety backup?" confirmation; a spinner
  during work; specific toasts on each outcome.
- **Composition root** wires `expoArchiveStore` alongside the photo store and DB handle.
- **Tests (20 new, 238 total):** schema-version drift guard; the full parse/validate
  matrix (valid, unknown-key stripping, non-JSON, foreign app, malformed data,
  unknown format, too-new, older-no-upgrader); transactional replace (round-trip,
  replace-not-merge, and mid-transaction failure rolls back deletes + inserts); and
  the orchestration — export→mutate→import byte-equivalence incl. photo reconcile,
  malformed/newer refused-untouched, the safety-export gate (decline aborts, confirm
  proceeds, success never asks), and replace-failure rolls back without committing photos.

## What changed

New: `data/backup/*` (schema, schemaVersion, upgraders, parseEnvelope, collect,
replace, archiveStore, expoArchiveStore, backupService, importError, index) + tests +
`backupTestKit`; `features/settings/components/BackupSection`. Modified: `app/_layout`
(archive-store wiring); `SettingsScreen` (Backup section). Added deps: `fflate`,
`expo-sharing`, `expo-document-picker`. No migration, no frozen document changed.

## Screens affected

Settings (new Backup section: Export / Import with two confirmation dialogs).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Schema version matches migration journal (drift guard) | node test | ✅ |
| Valid envelope parses; unknown keys stripped | node test | ✅ |
| Non-JSON / foreign app / malformed data ⇒ invalid-data | node tests | ✅ |
| Unknown format ⇒ unsupported-format; newer ⇒ schema-too-new; old-no-upgrader ⇒ unsupported-schema | node tests | ✅ |
| Transactional replace round-trips a full dataset | node test (real SQLite) | ✅ |
| Replace is replace-not-merge | node test | ✅ |
| Mid-transaction insert failure rolls back deletes + inserts | node test | ✅ |
| Export → mutate → import ⇒ byte-equivalent data + photos (orphan swept) | node test | ✅ |
| Malformed / newer archive ⇒ refused, data untouched | node tests | ✅ |
| Safety-export failure: decline aborts untouched; confirm proceeds; success never asks | node tests | ✅ |
| Replace failure never commits photos, data intact | node test | ✅ |
| `npm run check` | typecheck + lint + format + 238 tests + db:check (15 tables) | ✅ green (3× stable) |
| On-device backup walk (export→share, reinstall→import, corrupt-zip error, safety-fail confirm, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **Zip pack/unpack, the share sheet, and the document picker are device-only** —
   they need real files and native surfaces, so they live behind `expoArchiveStore`
   and fold into TD-001 (checklist extended with a full backup walk). The
   correctness-critical logic (validate → safety → replace → reconcile, and the
   transaction rollback) is proven off-device against real SQLite + in-memory doubles.
2. **Import is replace-only, never merge** — deliberate (§6.4): a single-user tool
   restoring a backup means "make my data exactly this." Merge semantics are ambiguous
   and dangerous and are explicitly out of scope for v1.
3. **The upgrader registry is empty** — only the current schema has ever shipped, so
   any older `schemaVersion` is honestly refused rather than guessed. The first
   upgrader (and a round-trip test for it) arrives with the `phases` table in Phase 19.
4. **No image compression in the archive** beyond the originals — a personal-scale
   gallery doesn't warrant it; a downscale pass can come later if archives grow large.

## Technical debt

None introduced. TD-001's device checklist gains the backup walk; the earlier
deferrals stand unchanged (TD-003 gestures, TD-008 keyboard, TD-009 bodyweight source).

## Retrospective

**What went well?** The `PhotoStore`-style injected boundary paid off again: putting
all zip/share/picker I/O behind `ArchiveStore` made the entire import *decision tree* —
validation order, the safety-export gate, the point-of-no-return, and the transaction
rollback — provable with an in-memory double and real SQLite, with zero device
dependency. Ordering the import so the only failure residue is a recoverable orphan
file (never a corrupt DB) fell straight out of the DATABASE §6 contract. And keeping
`parseAndUpgradeEnvelope` pure meant the "refused ⇒ untouched" guarantee is a property
of a single side-effect-free function.

**What was harder than expected?** A genuine Jest/better-sqlite3 environment flake:
the original fault-injection tests forced a failure with a `CHECK` violation *inside*
the manual `BEGIN` transaction, and under the full 56-file suite exactly one of the
two such tests would intermittently see the native constraint fail to propagate (the
same per-file-realm native-addon quirk noted in Phase 6, where the driver binds
`Error` to whichever realm loaded it first). It reproduced even single-worker and
never in isolation — a shared process-global addon artifact, not a product bug (the
constraint fires deterministically outside Jest, and the rollback is correct in
isolation). The fix was to make the fault a **deterministic JS-layer throw** (spy the
`insert` to fail at the `sets` table), which exercises the exact same rollback path —
deletes *and* inserts must be undone — without depending on native constraint
propagation across realms. Green 3× over the full suite after.

**What should change before the next phase?** Nothing structural — this opens the
daily-use gate, so CP-C reviews M2 next. Phase 19 will be the first real exercise of
the version machinery (a `phases` upgrader + schemaVersion bump + round-trip test);
the empty-but-wired upgrader registry and the drift guard are there precisely so that
change is additive and safe.

## Lessons Learned

- **What surprised you:** that a native SQL `CHECK` inside a manual transaction could
  *intermittently* fail to throw under Jest's shared better-sqlite3 realm — a reminder
  that fault-injection tests should fail through a layer you fully control (a JS throw)
  rather than one the test environment can perturb (native constraint propagation).
- **What documentation prevented mistakes:** DATABASE §6 fixed the exact import
  ordering (validate → version → **attempted** safety export → single-transaction
  replace → photo reconcile) and the non-negotiable §6.7 rule that a failed safety
  export must force explicit user confirmation — both encoded directly, with the
  confirmation as an injected async hook so "silent continuation is forbidden" holds
  by construction.
- **What should be reused:** the injected-I/O-boundary pattern for anything external
  the DB references (files, now archives) — it keeps `core`/`data` free of native
  imports and makes the whole workflow unit-testable; the pure `parse → reconcile →
  validate` pipeline that guarantees "refused ⇒ untouched"; the versioned-envelope +
  sequential-upgrader shape for any forward-compatible document.
- **What should be avoided:** merge-on-import (ambiguous, dangerous); validating after
  writing anything; letting a fault-injection test depend on native constraint timing;
  importing expo from `core`/`data` (the archive store is wired at the composition root).
- **Amendment proposals:** none — no frozen-document defect surfaced. No new debt.
