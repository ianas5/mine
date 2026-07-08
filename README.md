# Personal Fitness Tracker

A private, single-user mobile fitness tracking app — workouts, nutrition, body
measurements, progress photos, and analytics that explain progress over time.

## Start here

**[`PROJECT_INDEX.md`](./PROJECT_INDEX.md)** — the navigation entry point for the
entire documentation set: what each document owns, reading order, key decision
locator, and precedence.

- [`PRODUCT_PRINCIPLES.md`](./PRODUCT_PRINCIPLES.md) — the product's constitution (highest authority)
- [`PROJECT_VISION.md`](./PROJECT_VISION.md) — the product definition
- [`docs/`](./docs) — the eight technical documents (domain, architecture, database, analytics, design, UX, standards, workflow) and the implementation roadmap

## Status

**Documentation baseline: FROZEN (v1 · 2026-07-08).** The complete governing
documentation set is committed; implementation has not started and awaits
explicit approval to begin Phase 0 of
[`docs/IMPLEMENTATION_ROADMAP.md`](./docs/IMPLEMENTATION_ROADMAP.md).

Per the freeze terms: no document changes except for real, discovered defects
(via the amendment flow in `docs/DEVELOPMENT_WORKFLOW.md` §6); new features
belong to post-v1; implementation follows the frozen documentation exactly.

## Technology (baseline, detailed in `docs/ARCHITECTURE.md` §3)

Expo (latest stable SDK) · React Native · TypeScript (strict) · Expo Router ·
SQLite (`expo-sqlite`) + Drizzle ORM/drizzle-kit · Zustand · React Hook Form ·
Zod · MMKV · Victory Native XL · Inter · Lucide — local-first, offline-only,
single-user, metric units, English UI.

## Preserved assets

- `icon.png` — 1024×1024 app icon, retained for reuse in the new build.
