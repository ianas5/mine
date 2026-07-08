# CP-A — Architecture Verification Checkpoint (end of M0)

**Date:** 2026-07-08 · **Covers:** Phases 0–2 · **Type:** review only (no features).

Per IMPLEMENTATION_ROADMAP §3, verified against the frozen documents with them
open. Device-only items are **not blockers** (owner directive): they accumulate
into the single consolidated TD-001 on-device pass.

## 1. PRODUCT_PRINCIPLES — the Tests

- **No vanity metrics / no dark patterns shipped.** M0 shipped infrastructure +
  a Settings screen only; nothing computes or displays metrics yet. ✅
- **Logging-speed priority (P9) not yet exercised** (no logging flows in M0);
  the primitives that will serve them (Stepper ≤2-tap, Sheet) exist and are
  tested. ✅
- **P15 instant feedback:** optimistic settings update in place; skeletons over
  spinners in the Settings loading state. ✅
- **P16/P17/P18:** no engagement mechanics, no emoji chrome, single-accent
  discipline held (verified in the gallery). ✅
- **P19 offline-first:** zero network calls anywhere; the app is entirely local.
  The web "Database failed to open" state actually *demonstrates* the app has no
  fallback to anything remote. ✅

## 2. FITNESS_DOMAIN — formula conformance

- **No domain formulas have shipped yet** (they begin Phase 4). The
  `domain/fitness/constants.ts` single-source rule (CODING_STANDARDS §6.2) is
  not yet exercised. Nothing to violate. ✅ (N/A for M0)

## 3. ARCHITECTURE — boundaries & state ownership

- **Boundary lint clean** and the dependency direction is *enforced*, not
  documented: cross-feature imports, `domain`-imports-framework, SQL-outside-
  `data`/`core/db`, and MMKV-outside-`core/storage` were each proven to fail
  lint via deliberate violations (Phases 0–2). ✅
- **Vertical-slice read** (the one required manual trace):
  `app/settings.tsx` → `features/settings/screens/SettingsScreen` →
  `features/settings/hooks/useSettings` → `data/repositories/settingsRepository`
  → `core/db` (`getDb`, change-bus) + `data/schema`. The UI never touches SQL;
  it goes through the hook → repository. The migration bundle reaches the gate
  only via the `app/_layout` composition root (now documented, ARCHITECTURE §4
  amendment). ✅
- **State ownership (§6 table):** theme override → MMKV (disposable); domain
  config (weekly target, height, bodyweight) → SQLite (in backups). Matches the
  table exactly. ✅
- **Routes thin:** every `(tabs)` route and `settings.tsx` is a one-line
  delegator; only `_layout` carries wiring. ✅

## 4. DATABASE — schema, migrations, no derived tables

- **Schema == docs:** the `settings` table matches DATABASE §3.1 (columns,
  defaults, single-row `CHECK (id = 1)`), verified at the DB level by
  `db:check` and the repository tests. ✅
- **Migrations immutable & forward-only:** `0000_misty_maverick.sql` committed;
  the runner applies pending migrations in the DB-ready gate. ✅
- **No derived tables:** the only table is `settings`; nothing recomputable is
  persisted. ✅

## 5. Tech-debt registry audit

- **TD-001** (device verification) — open, on schedule; now the consolidated
  device-pass bucket. Not a blocker (owner directive).
- **TD-002** (placeholder-screen duplication) — open; dissolves as real screens
  land; not yet due (CP-D).
- **TD-003** (Sheet drag-dismiss) — open; due Phase 9 or 21. On schedule.
- No overdue entries. ✅

## Verdict

**CP-A passes** for all hardware-independent criteria. The only outstanding M0
item is the physical-iPhone verification consolidated in TD-001, which — per the
owner's directive — does not gate M1. **M1 (Phase 3) is cleared to start.**
