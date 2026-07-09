# Phase 9 — Nutrition Logging

**Closed:** 2026-07-09 · **Milestone:** M2 (Fuel & Body)

The meal loop opens M2: a foods catalog, a day view, and a Log Meal sheet built
for ≤ 3 taps on a repeated meal. The invariant that governed templates and PRs
holds here too — **history is self-contained**: every meal entry snapshots its
scaled macros and food name at log time, so editing (or deleting) a food never
rewrites what you already logged.

## What was built

- **Migration 0005** (DATABASE §3.5): `foods` (per-serving macros, serving-unit
  CHECK, quick-meal flag; no fiber) and `meal_entries` (macros snapshotted, slot
  CHECK, `food_id … ON DELETE SET NULL` provenance, `(date)` + `(food_id)`
  indexes). Day totals are `Σ` at read time — no stored aggregate.
- **Nutrition domain** (`domain/nutrition`, pure): `scalePortion` (logged =
  per-serving × amount/serving, kcal→int, grams→0.1 g, never ÷0), `sumMacros`
  (day/range totals), `atwaterKcal` + `isKcalImplausible` (the §4.2 macro/energy
  cross-check — a validation aid that never overrides entered kcal), the
  serving-unit / meal-slot taxonomy, and `defaultSlotForHour` (time-of-day slot
  fallback, §5.2).
- **`nutritionRepository`**: foods CRUD (delete SET-NULLs entry provenance),
  `getFoodPicks` (Recent & Frequent with quick meals pinned, each carrying its
  last-used amount and most-frequent slot as smart defaults), meal-entry
  add/list/delete/**restore** (verbatim, for Undo), and `foodHasHistory`. Starter
  foods seed (~16 foods incl. two quick meals), insert-if-absent.
- **Log Meal sheet** (UI_UX §4.3): opens on Recent & Frequent → tap a food → a
  portion panel pre-filled with the last-used amount + smart slot, live scaled
  macro preview → Save. A repeated meal is **open · food · Save = 3 taps**;
  search and "New food" sit above for the uncommon case.
- **Nutrition day view**: consumed-macros summary (targets arrive Phase 10),
  the day's entries, previous/next day navigation with a tap-to-today, and a
  Foods entry. **Undo-able delete**: a trash tap removes the entry immediately
  and raises a 5 s **Undo toast** that restores the identical row (the `Toast`
  primitive gained an optional action for this).
- **Foods catalog + editor**: list/search; create/edit a food (name, serving
  amount + unit, kcal, protein/carb/fat, quick-meal) with the live plausibility
  warning; delete is Dialog-gated (entries keep their snapshot).
- **Tests (23 new, 182 total):** portion scaling / rounding / ÷0, day totals,
  Atwater + implausibility, slot fallback, and a real-SQLite `nutritionRepository`
  suite — snapshot rule (edit food ⇒ entries unchanged), day totals reconcile,
  **Undo restores an identical row**, food delete keeps entries (SET NULL), and
  the Recent-&-Frequent ordering with last-used portion. Plus the CP-B follow-up
  date/timed-PR tests landed just before this phase.

## History stays self-contained (the standing principle)

- **Meal entries snapshot everything they need** — `food_name`, `logged_unit`,
  and the four scaled macros — so a food edit touches only the `foods` row. The
  test redefines a food completely and the prior entry stays byte-identical.
- **Delete is SET NULL, never cascade** — removing a food clears each entry's
  `food_id` and leaves the snapshot intact; the day still totals correctly.
- **Undo re-inserts verbatim** (same id, same snapshot, same `logged_at`), so an
  accidental delete costs nothing.

## What changed

New: migration 0005; `domain/nutrition/{taxonomy,macros,index}`;
`domain/models/nutrition`; `data/repositories/nutritionRepository`;
`data/seed/foods` (+ seeded in `seedDatabase`); nutrition feature
(`hooks/{useNutritionDay,useFoodPicks,useFood}`, components
`MacroSummary`/`MealEntryRow`/`LogMealSheet`, screens
`NutritionScreen`/`FoodsScreen`/`FoodEditorScreen`); nutrition converted to a
route stack (`app/(tabs)/nutrition/…`). Modified: `core/ui/Toast` (optional
action for Undo), `core/utils/date` (`addDaysIso`), `changeBus` (`nutrition`
channel), `schema/tables` (foods + meal_entries; taxonomy imported from the
domain as the single source). No frozen document changed.

## Screens affected

Nutrition day view (new), Foods catalog + editor (new), Log Meal sheet (new).
The Nutrition tab is now a drill-down stack like Workouts.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Portion scaling / rounding / zero-serving guard | domain test | ✅ |
| Day totals = Σ entries | domain + real-SQLite tests | ✅ |
| Macro/energy cross-check flags implausible, never overrides | domain test | ✅ |
| Time-of-day slot fallback | domain test | ✅ |
| **Snapshot: edit food ⇒ past entries unchanged** | real-SQLite test | ✅ |
| **Undo restores an identical row** | real-SQLite test | ✅ |
| Food delete keeps entries (SET NULL) | real-SQLite test | ✅ |
| Recent & Frequent ordering + last-used portion | real-SQLite test | ✅ |
| `npm run check` | typecheck + lint + format + 182 tests + db:check (11 tables) | ✅ green (3× stable) |
| On-device meal loop (≤ 3-tap repeat, quick meals, day nav, delete+undo, edit-food-safe, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** for the nutrition screens (DB-backed; expo-sqlite
   native-only). Correctness is proven by domain + repository tests; visuals + the
   ≤ 3-tap device count fold into the consolidated TD-001 pass (checklist extended).
2. **Delete is a visible trash tap + Undo toast, not swipe-to-reveal.** The
   forgiving Undo mechanism (UI_UX §6) is in place; the swipe gesture is deferred
   with the rest of the gesture-handler work — **TD-003** was broadened to cover
   row swipe-to-reveal alongside Sheet drag-dismiss (Phase 21 feel pass).
3. **Targets, remaining, adherence, and water are Phase 10** — the day view shows
   consumed only, by roadmap design.
4. **Portion entry uses a stepper** (last-used amount pre-fill keeps the common
   case at zero adjustment); a tap-to-type numeric path can join the Phase 21
   keyboard-first pass if the stepper proves slow on device.

## Technical debt

- **TD-003 broadened** to include destructive-row swipe-to-reveal (deletes work
  today via visible tap + Undo). No new debt IDs.

## Retrospective

**What went well?** The snapshot rule made "edit a food safely" free again — meal
entries copy what they need, so the food table is pure definition and the test
just confirms independence. The domain/data/feature split stayed clean: all the
math (`scalePortion`, `sumMacros`, the cross-check) is pure and tested without a
device, the repository owns SQL + the Recent-&-Frequent ranking, and the sheet is
a thin composition. Extending `Toast` with one optional action turned Undo into a
three-line call at every delete site.

**What was harder than expected?** Two judgement calls. The Recent-&-Frequent
ranking wants "most-used + last-used amount + most-frequent slot" without a heavy
query every keystroke — resolved by scanning a bounded recent window (400 entries)
and aggregating in memory, accurate at personal scale and cheap. And the ≤ 3-tap
budget hinges on the *slot* default being right: a naive "last slot" misfires, so
the picker computes the mode slot per food (with a time-of-day fallback), keeping
Save as the third tap without a slot correction.

**What should change before the next phase?** Nothing structural. Phase 10 layers
time-versioned targets, remaining/adherence, and water onto this day view — the
resolution logic lands in `nutritionRepository` as its single implementation
(P5's defining test), and the day view already has the slot to render remaining.

## Lessons Learned

- **What surprised you:** the same three-phase-old decision — history copies its
  own facts — paid out a third time (templates, then... nutrition). "Editing a
  definition must not rewrite the log" needed no guard code, only *not* recomputing
  from the food. Snapshotting at write time is the cheapest possible integrity.
- **What documentation prevented mistakes:** FITNESS_DOMAIN §4.1/§4.2 fixed the
  snapshot rule, the exact scaling formula, the four-macro/no-fiber scope, and the
  cross-check-as-aid-not-authority stance; DATABASE §3.5 gave the SET NULL + no-
  stored-day-total design; UI_UX §4.3/§5.2/§6 fixed the ≤ 3-tap flow, the smart
  slot/portion defaults, and the Undo-toast tier for reversible deletes.
- **What should be reused:** snapshot-at-write for any "log references a
  definition" relationship (measurements/photos next); the bounded-recent-window +
  in-memory aggregate for frequency/recency ranking; the `Toast` action for any
  reversible destructive action; the day-nav + `addDaysIso` pattern for every
  date-scoped view.
- **What should be avoided:** recomputing a logged entry from its current
  definition (breaks history); a "last value" default where a "most-frequent"
  default is what keeps the tap budget (slot); unbounded full-table scans for
  ranking.
- **Amendment proposals:** none — no frozen-document defect surfaced. TD-003 was
  broadened, not a doc change; no new debt.
