# Phase 12 — Measurement Comparison

**Closed:** 2026-07-09 · **Milestone:** M2 (Fuel & Body)

Any two body snapshots can now be compared field by field, with the change coloured
by which direction is good for each metric. The governing rule this phase — **never
fabricate a baseline**: a field present on only one of the two selected dates shows
"—", never a guessed change.

## What was built

- **Comparison domain** (`domain/body/comparison.ts`, pure): `compareSnapshots(a, b,
  heightCm)` returns a per-field result — absolute change (`B − A`), percentage
  change (`(B − A)/A × 100`, **null when A = 0 or the field is incomparable**, §5.4),
  and a direction. Direction is judged by the §6.4 **stability deadband** on the
  single A→B delta (sub-threshold ⇒ `stable`) mapped through §5.3 **directionality**
  (`lower`-is-better for waist/hips/body-fat/visceral; `higher`-is-better for muscle
  mass and the muscular sites; `neutral` for weight and BMI, which depend on a goal).
  A field on only one date is **`incomparable`** — no change, no fabricated baseline
  (P8). BMI is derived per date (entered-else-derived) so it compares whenever both
  dates have weight + height.
- **Directionality + deadband constants** (`BODY_DIRECTION`, `BODY_STABILITY`): the
  §5.3 good-direction and §6.4 stability thresholds per field (the doc's
  weight/circumference/body-fat/muscle values; visceral-fat and BMI given
  conservative defaults in the same spirit).
- **`bodyRepository.snapshotDates`**: the dates that have a snapshot (newest first) —
  the compare-view pickers are limited to these.
- **Compare screen** (`/measurements/compare`): two date pickers (snapshot dates
  only, defaulting to the two most recent), an A → B table of every field present on
  either date, with Δ, %Δ, and a direction glyph coloured green/red/neutral;
  incomparable fields render "—". Reached from a Compare entry on the Measurements
  home (shown once ≥ 2 measurement days exist). The Measurements tab became a stack.
- **Tests (7 new, 210 total):** incomparable ⇒ "—" (no baseline), absolute + %
  change, §5.3 direction mapping (waist-down improving, arm-down declining),
  deadband ⇒ stable, weight/BMI neutral, %-undefined at A = 0, and per-date BMI
  derivation making BMI comparable.

## No fabricated baseline (the standing principle)

- **Incomparable is a first-class result.** When a field is present on only one
  date, `compareSnapshots` returns `direction: 'incomparable'` with `deltaAbs` and
  `deltaPct` both `null`; the UI shows "—". There is no path that invents the
  missing side.
- **Percentage is honest too.** `(B − A)/A` is `null` when `A = 0` (§5.4) — shown as
  no percentage rather than infinity or a divide error.

## What changed

New: `domain/body/comparison` (+ barrel exports); `bodyRepository.snapshotDates`;
`features/measurements/screens/CompareScreen`; measurements converted to a route
stack (`app/(tabs)/measurements/…` with `index` + `compare`). Modified:
`MeasurementsScreen` (Compare entry when ≥ 2 days). No frozen document changed.

## Screens affected

Measurements home (Compare entry), Compare screen (new).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Field on only one date ⇒ incomparable, "—" (no baseline) | domain test | ✅ |
| Absolute + percentage change on both-present fields | domain test | ✅ |
| §5.3 direction (waist-down improving; arm-down declining) | domain test | ✅ |
| §6.4 deadband ⇒ stable (number still shown) | domain test | ✅ |
| Weight/BMI neutral (no fixed good direction) | domain test | ✅ |
| % undefined at A = 0 | domain test | ✅ |
| BMI derived per date ⇒ comparable | domain test | ✅ |
| `npm run check` | typecheck + lint + format + 210 tests + db:check (14 tables) | ✅ green (3× stable) |
| On-device compare walk (pickers limited to snapshot dates, direction colouring, deadband stable, present-in-one "—", both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** for the compare screen (DB-backed). Correctness is proven
   by the domain comparison tests; visuals fold into TD-001 (checklist extended).
2. **Date pickers are a snapshot-date list sheet, not a calendar** — correct by
   design (§12 limits pickers to dates that actually have data) and sufficient at
   personal scale; a scrollable/searchable variant can come with the Phase 21 pass
   if the history grows long.
3. **Weight/BMI show neutral direction** — deliberate: their "good" direction needs
   a `targetWeight` the app doesn't store yet. When a weight goal exists, the same
   `BODY_DIRECTION` map is the one place to make weight goal-aware.

## Technical debt

None introduced. The prior deferrals stand (TD-008 keyboard chaining, TD-009
bodyweight source; both unchanged and on schedule).

## Retrospective

**What went well?** The Phase-11 groundwork made this almost declarative: the field
taxonomy, `resolveBmi`, and `BodySnapshot` were already the exact shapes the compare
needed, so `compareSnapshots` is one pure pass over the field list. Encoding
directionality and the deadband as per-field constant maps kept the good/bad
colouring a domain fact (§5.3 says it must not be a per-chart setting) — the screen
only reads a `direction` and picks a colour. And "never fabricate a baseline" was
trivial to honour because the snapshots already store nullable fields: incomparable
just falls out of "either side is null."

**What was harder than expected?** Deciding weight's direction. §5.3 makes weight
"toward `targetWeight`", but there is no target-weight setting yet, so any fixed
green/red would be a guess in one direction. Marking weight and BMI `neutral` (show
the change, no good/bad colour) is the honest choice until a goal exists — and it
localizes the future change to a single map entry.

**What should change before the next phase?** Nothing structural. Phase 13 (photos)
is the last M2 logging surface before the daily-use gate; it reuses the same
per-date, merge-minded discipline, with the added filesystem-transaction concern
(write file → insert row; delete row → delete file) from DATABASE §3.6.

## Lessons Learned

- **What surprised you:** the "no fabricated baseline" rule needed no special code —
  because snapshots store nullable fields and the compare returns `incomparable`
  whenever either side is null, honesty was the *default*, not an extra guard.
- **What documentation prevented mistakes:** FITNESS_DOMAIN §5.4 fixed the exact
  comparison outputs and the A = 0 / present-in-one rules; §5.3 fixed directionality
  as a domain fact (not a UI toggle); §6.4 gave the stability thresholds so a tiny
  change reads "stable" rather than a spurious trend.
- **What should be reused:** per-field constant maps (`BODY_DIRECTION`,
  `BODY_STABILITY`) as the single source for interpretation the UI only reads; the
  `incomparable` sentinel for any two-sided diff where a side may be missing;
  deriving a computed field (BMI) into both sides before comparing so it participates
  honestly.
- **What should be avoided:** picking a fixed good-direction for a goal-dependent
  metric (weight) — `neutral` until the goal exists beats a misleading colour;
  fabricating either side of a comparison; dividing by a zero baseline.
- **Amendment proposals:** none — no frozen-document defect surfaced. No new debt.
