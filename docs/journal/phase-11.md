# Phase 11 — Measurements Logging

**Closed:** 2026-07-09 · **Milestone:** M2 (Fuel & Body)

Body tracking arrives: a ≤ 3-tap weigh-in, a partial-by-design tape-day sheet, and
a current-state home with a delta'd weight log. The governing rule this phase —
**body snapshots are historical records, not forms**: saving writes only the
fields you entered, an omission never clears a stored value, and clearing is a
separate, deliberate action.

## What was built

- **Migration 0007** (DATABASE §3.6): `body_snapshots` — one row per date, every
  metric optional/nullable, bilateral sites stored **per side** (never collapsed).
- **`bodyRepository`** with the **merge-upsert contract**: `saveSnapshot(date,
  patch)` writes only the fields present in the patch (insert-or-update on the date
  PK; omitted fields keep their stored values, absent fields on a new date default
  NULL). Clearing is the *separate* `clearField(date, field)` — the save API
  literally cannot express a clear, so an omission can never erase. Plus
  `getSnapshot`, `listSnapshots`, and `getWeightLog`.
- **Body domain** (`domain/body`, pure): the field taxonomy (keys, labels, units,
  the natural head-to-toe measuring order), `deriveBmi` + `resolveBmi`
  (entered-wins-else-derived, §5.2), `latestFieldValues` (most-recent non-null per
  field for the current-state + placeholders), `frequentlyLoggedFields` (co-logged
  ≥ 50 % → expanded, §5.2), and `weightLogWithDeltas`.
- **Add Weight sheet** (UI_UX §4.4): a stepper pre-set to the last weight (0.1 kg
  steps) → Save — a repeat weigh-in is Add Weight · adjust · Save ≤ 3 taps.
- **Add Measurements sheet** (UI_UX §4.5): every site with its last value as a
  placeholder; frequently co-logged fields expanded, the rest behind **More sites**;
  fill any subset → merge-upsert only the filled fields.
- **Measurements home**: current composition (with entered-else-derived BMI) +
  circumferences (latest per field) + the weight log with per-entry deltas.
- **Tests (12 new, 203 total):** BMI derive/resolve, latest-per-field, frequency
  heuristic, weight deltas (domain); and a real-SQLite `bodyRepository` suite —
  partial create, **same-date merge (omit ≠ clear)**, **bilateral per side**,
  **explicit clear removes one field only**, and the weigh-ins-only weight log.

## Snapshots are records, not forms (the standing principle)

- **The save API can't clear.** `MeasurementPatch` is `Partial<Record<field,
  number>>` — numbers only. Omitting a field writes nothing for it; there is no
  value that means "erase." The same-date-merge test proves a later partial save
  leaves untouched fields intact.
- **Clearing is deliberate.** Only `clearField` nulls a value, one field at a time
  — proven to leave every other field untouched.
- **History is per-side and honest.** Bilateral sites never collapse; a field never
  recorded reads as absent (no fabricated baseline, P8).

## What changed

New: migration 0007; `domain/body/{fields,snapshot,index}`; `domain/models`
`BodySnapshot` (re-exported from `domain/body`); `data/repositories/bodyRepository`;
measurements feature (`hooks/{useBodyData,useBodyHeightCm}`, components
`AddWeightSheet`/`AddMeasurementsSheet`, screen `MeasurementsScreen`). Modified:
`changeBus` (`body` channel). No frozen document changed.

## Screens affected

Measurements home (new — current state + weight log), Add Weight sheet, Add
Measurements sheet.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| BMI derive / entered-wins | domain test | ✅ |
| Latest non-null value per field | domain test | ✅ |
| Frequently-logged ≥ 50 % (weight-only default) | domain test | ✅ |
| Weight-log deltas from the previous weigh-in | domain test | ✅ |
| Partial create (absent ⇒ null) | real-SQLite test | ✅ |
| **Same-date merge: omit ≠ clear** | real-SQLite test | ✅ |
| **Bilateral sites stored per side** | real-SQLite test | ✅ |
| **Explicit clear removes one field only** | real-SQLite test | ✅ |
| Weight log is weigh-ins only, newest first | real-SQLite test | ✅ |
| `npm run check` | typecheck + lint + format + 203 tests + db:check (14 tables) | ✅ green (3× stable) |
| On-device measurements walk (≤ 3-tap weigh-in, partial tape day, same-date merge, explicit clear, bilateral per side, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** for the measurements screens (DB-backed). Correctness is
   proven by domain + repository tests; visuals + the ≤ 3-tap device count fold into
   TD-001 (checklist extended).
2. **Keyboard-next chaining is not fully wired** — inputs set `returnKeyType="next"`
   but focus does not yet advance across fields (needs `Input` ref-forwarding).
   Every field is fillable by tap; registered as **TD-008** for the Phase 21
   keyboard-first pass.
3. **Weight isn't yet the bodyweight source for training volume** — bodyweight-load
   exercises still read `settings.defaultBodyweightKg` rather than the latest
   weigh-in (§5.2). Both are user-entered kg; registered as **TD-009**, unified at
   the dashboard/analytics phase where body + training data meet.
4. **Comparison between two dates** (§5.4) is **Phase 12** by roadmap design — this
   phase logs and shows current state; the compare view is next.

## Technical debt

- **TD-008** — keyboard-next focus chaining in multi-field sheets (Phase 21).
- **TD-009** — unify the bodyweight source (latest weigh-in vs settings) (Phase 16
  / first body+training analytics).

## Retrospective

**What went well?** Making the *type* enforce the rule was the win: because
`MeasurementPatch` only carries numbers, "omit ≠ clear" isn't a runtime guard I
have to remember — it's structurally impossible to express a clear through save,
and clearing lives in its own method. The merge-upsert then reduced to
`onConflictDoUpdate` with only the present columns in the `set`, and the test just
confirms the property. The domain stayed a set of tiny pure functions over the
field list, so the current-state, placeholders, frequency default, and deltas are
all four-line functions.

**What was harder than expected?** Two small things. The `BodySnapshot` type is
tightly coupled to the field list, so putting it in `domain/models` and importing
`BodyField` from `domain/body` risked a cycle — resolved by making `domain/body`
the canonical home and having `domain/models` re-export it. And re-seeding the Add
Weight stepper to the latest weight on each open, without a render-phase setState,
was cleanest as a `key` at the call site (remount = fresh initial state) rather
than an effect.

**What should change before the next phase?** Nothing structural. Phase 12
(comparison) reads two snapshots and diffs them per §5.4 using the same field
taxonomy and the §5.3 directionality (which I deliberately left for Phase 12 where
it is first used) — the domain field metadata and `latestFieldValues` are already
the shape it needs.

## Lessons Learned

- **What surprised you:** the safest way to honour "omitted fields must never
  clear" was to make the save API unable to represent a clear at all — a type-level
  guarantee beats a remembered runtime check, and it made the merge-upsert trivial.
- **What documentation prevented mistakes:** DATABASE §3.6 fixed the merge-upsert
  contract, the explicit-clear action, and per-side bilateral storage; FITNESS_DOMAIN
  §5.1/§5.2 fixed the field set, the one-per-date rule, and BMI entered-else-derived;
  UI_UX §4.4/§4.5/§5.2 fixed the ≤ 3-tap weigh-in, partial-tape-day entry, and the
  co-logged-≥ 50 % expand default.
- **What should be reused:** the numbers-only patch type as the pattern for any
  "merge, never clear" upsert; a dedicated explicit-clear method rather than a
  nullable save; the field-taxonomy + pure-derivation split (labels/units/order in
  one place, math in small functions); the `key`-to-reseed trick for form defaults.
- **What should be avoided:** a nullable save patch that conflates omit and clear;
  placing a type in `domain/models` when it is really owned by a domain sub-module
  (cycle risk); render-phase setState for prop-driven form defaults.
- **Amendment proposals:** none — no frozen-document defect surfaced. Two scheduled
  debts (TD-008, TD-009), both feel/integration, not correctness.
