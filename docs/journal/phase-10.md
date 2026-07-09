# Phase 10 — Targets, Adherence Display & Water

**Closed:** 2026-07-09 · **Milestone:** M2 (Fuel & Body)

Intake gets meaning: time-versioned nutrition targets, remaining/consumed per
macro with per-day adherence, and water logging. The governing rule this phase —
**time-versioned targets are authoritative and resolved through one canonical
path**: every consumer (day view, remaining, adherence, and every future
analytic) asks `resolveTargetForDate`, and nowhere re-derives "the target for a
date."

## What was built

- **Migration 0006** (DATABASE §3.5): `nutrition_targets` (`effective_from`
  UNIQUE, macros + optional `water_ml`) and `water_days` (`date` PK, `ml`; an
  absent row = unlogged, `ml = 0` = a real logged zero).
- **The single canonical resolution** (`nutritionRepository.resolveTargetForDate`):
  the active target for a date is the row with the greatest `effective_from ≤
  date`, or `null` before the first target (insufficient-data — never a fabricated
  default). This is the *only* place that logic exists; `setTarget` (upsert on
  `effective_from` — "set new targets from <date>"), `listTargets`, and the day
  view all route through it or feed it.
- **Adherence domain** (`domain/nutrition/adherence.ts`, pure): `dayAdherence`
  (protein floor with a 90 % near-band and never "over"; calories ±10 %; carbs/fat
  ±15 %; water ≥ target only when a water goal is set — FITNESS_DOMAIN §4.3) and
  `remainingMacros` (target − consumed, negative = over). Takes an already-resolved
  target — it never resolves.
- **Water logging** (`getWater`/`addWater`/`setWater`): +/- a configurable cup
  (from Settings), floored at 0, preserving the 0-vs-unlogged distinction.
- **Day view upgraded**: `useNutritionDay` now composes entries + resolved target
  + water into totals, remaining, and adherence in one place. `MacroSummary` shows
  consumed / target and "N left / N over" per macro, coloured by adherence; the
  `WaterCard` shows logged/target with +/- cup and a "Not logged" state.
- **Targets editor** (`TargetsEditorScreen`, UI_UX §4.7): "Set new targets from
  <date>" (a date stepper, default today), the five fields seeded from today's
  active target, and the past eras listed **read-only** beneath — the
  time-versioned model is visible, never a silent "edit".
- **Tests (9 new, 191 total):** adherence bands + protein floor + water-only-when-
  targeted + negative remaining (domain); and a real-SQLite suite for the
  **canonical resolution** — before-first ⇒ null, greatest-≤-date, the **P5
  defining test** (new targets forward never change an old day), same-date upsert,
  and the water 0-vs-unlogged distinction.

## The canonical path, enforced (the standing rule)

- **One implementation.** `resolveTargetForDate` is the sole holder of the
  "greatest `effective_from ≤ date`" query. The adherence and remaining functions
  are pure over a *resolved* target; the day hook calls the resolver once per date;
  no screen or component computes a target from dates.
- **Authoritative and honest.** A date before the first target resolves to `null`
  and the UI says "no target set" — insufficient-data, never a default. Setting
  new targets writes a new era and leaves every prior day's resolution untouched
  (the P5 test proves it).

## What changed

New: migration 0006; `domain/nutrition/adherence`; `nutrition_targets` +
`water_days` schema; `domain/models` `NutritionTarget`; repository target/water
methods; nutrition `hooks/{useTargets,useWaterCupMl}`, components `WaterCard`,
screen `TargetsEditorScreen` + route. Modified: `useNutritionDay` (composes
target/water/remaining/adherence), `MacroSummary` (target/remaining/adherence),
`NutritionScreen` (water card + Targets entry). No frozen document changed.

## Screens affected

Nutrition day view (target/remaining/adherence + water card), Targets editor (new).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Adherence: protein floor + near-band, calorie/carb/fat bands, water | domain test | ✅ |
| Remaining incl. negative (over) | domain test | ✅ |
| Resolution: before-first ⇒ null (insufficient-data) | real-SQLite test | ✅ |
| Resolution: greatest `effective_from ≤ date` | real-SQLite test | ✅ |
| **P5: new targets forward never change an old day** | real-SQLite test | ✅ |
| Same-date set upserts the era in place | real-SQLite test | ✅ |
| Water: logged 0 ≠ unlogged; +/- floors at 0 | real-SQLite test | ✅ |
| `npm run check` | typecheck + lint + format + 191 tests + db:check (13 tables) | ✅ green (3× stable) |
| On-device targets/adherence/water walk (remaining updates, forward targets don't touch today, pre-first-target reads insufficient-data, water 0-vs-unlogged, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** for the nutrition screens (DB-backed). Correctness is
   proven by domain + repository tests; visuals fold into TD-001 (checklist
   extended with the targets/adherence/water walk).
2. **Range adherence % and completeness** (hit-days / logged-days over a range,
   §4.3) are **analytics**, not this phase — Phase 10 shows the *per-day* status,
   which is what the day view needs. The range aggregation lands with the
   nutrition analytics calculator (M4).
3. **Targets editor uses a date stepper, not a calendar picker** — fine for
   "today / a nearby date"; a full picker can join the Phase 21 pass if setting a
   far-past era is ever needed on device.

## Technical debt

None introduced. TD registry unchanged (TD-004 resolved earlier; TD-001/002/003/
005/006/007 all on schedule).

## Retrospective

**What went well?** Keeping resolution in exactly one function made the whole
phase compose cleanly: the domain adherence/remaining functions are pure over a
resolved target, the day hook resolves once and hands the number down, and the P5
guarantee is a property of that single query rather than something to defend in
every consumer. The 0-vs-unlogged water rule fell straight out of "absent row =
null, present row = value" with no sentinel.

**What was harder than expected?** The adherence bands needed care to match §4.3
exactly — protein is a *floor* (a near-band below, never an "over"), while
calories/carbs/fat are *symmetric bands* (both under and over matter). Encoding
that as two distinct helpers (`floorStatus` vs `bandStatus`) rather than one
generic comparator kept each rule readable and testable against the doc's numbers.
The other care point was the boundary rule: the nutrition feature can't import the
settings feature, so the water-cup size reads through a small nutrition-local hook
over `settingsRepository` — the right seam, not a cross-feature import.

**What should change before the next phase?** Nothing structural. When nutrition
analytics arrive (M4), the range adherence calculator will resolve each day's
target through this same `resolveTargetForDate` — the single path is already the
one the analytics layer must use, so there is no second implementation to reconcile.

## Lessons Learned

- **What surprised you:** "authoritative, resolved once" is as much a *simplicity*
  win as a correctness one — with a single resolver, the time-versioning guarantee
  (old days keep old targets) needs no defensive code anywhere else, and the P5
  test is a two-line property of that query.
- **What documentation prevented mistakes:** DATABASE §3.5 fixed the
  `effective_from` resolution rule, the no-snapshot decision, and the water
  0-vs-unlogged semantics; FITNESS_DOMAIN §4.3 gave the exact adherence bands
  (protein floor vs symmetric bands, water-only-when-targeted); UI_UX §4.7 fixed
  the "set new targets from …, past eras read-only" framing that keeps the model
  visible.
- **What should be reused:** the single-resolver pattern for any time-versioned or
  date-scoped authority (phases next); pure-judgement-over-resolved-input to keep
  domain functions free of lookups; the absent-row-is-null convention for any
  daily-keyed value; the feature-local repository hook when a boundary blocks a
  cross-feature settings read.
- **What should be avoided:** duplicating a "which record applies to this date"
  lookup in more than one place; fabricating a default target for pre-history days
  (insufficient-data is the honest answer, P8); collapsing distinct adherence rules
  (floor vs band) into one comparator.
- **Amendment proposals:** none — no frozen-document defect surfaced. No new debt.
