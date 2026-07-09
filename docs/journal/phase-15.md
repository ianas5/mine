# Phase 15 — Analytics Foundation + First Charts

**Closed:** 2026-07-09 · **Milestone:** M3 (Awareness) — first phase

The app starts to *explain* progress. This phase builds the trend machinery and chart
infrastructure and proves them on the first real surface: the **Body section is live**
(weight trend, distance-to-target, waist trend), and the **Exercise Report gains its
e1RM trend + sparkline** (closing the Phase 7 trend debt). The governing discipline —
analytics is **entirely derived**: nothing analytical is authoritative, every metric is
a `MetricResult`, every displayed number carries the interpretation triplet, and
insufficient-data propagates honestly to the UI.

## The honesty contract (ANALYTICS §2/§3, the standing principle)

- **`MetricResult<T>` is the single result type.** Every metric is `ok` (value + window
  + provenance) or `insufficient-data` (reason + a concrete `needed` sentence). No
  throws for data reasons; no `NaN`/`Infinity`/`null` can reach the UI.
- **The interpretation triplet is structural.** `StatTile` *requires* a context line and
  `ChartFrame` *requires* an interpretation sentence — a value or plot literally cannot
  be rendered without its reference + classification (DESIGN_SYSTEM §6; ANALYTICS rule 3).
- **Insufficient-data is stated, never papered over.** Below the §6.4 minimums (≥ 3
  points over ≥ 14 days) a trend returns its `needed` text; the chart is replaced by an
  honest message, never a fabricated line; distance-to-target with no goal returns
  `no-target-set`.
- **The engine is pure.** `domain/analytics` takes windowed domain models + `today` as
  inputs and returns values — no repository calls, no clocks, no MMKV (ANALYTICS rule 8).
  Repositories window/join; the engine computes; hooks fetch + memoize; the UI renders.

## What was built

**Pure engine (`domain/analytics/`, all `MetricResult`-returning):**
- `metricResult` — the `ok`/`insufficient` contract + `Range`.
- `ranges` — the six canonical rolling windows (§7); `rangeWindow(key, today, firstRecord)`.
- `timeSeries` — `SeriesPoint`, windowing, span.
- `regression` — least-squares slope vs. **day-index** (respects real gaps; §6.4).
- `trend` — the canonical trend: ≥ 3 points / ≥ 14 days minimums, `Δ = slope × span`
  deadband, classification via §5.3 directionality. Adds a `neutral` classification for
  metrics with no fixed good direction (e.g. weight with no goal) — moving but unjudged.
- `movingAverage` — the §6.2 7-day MA (≥ 2 points in window, else no value).
- `bucketing` — display downsampling to ≤ 120 points (daily 7d/30d, weekly 90d/180d/365d,
  monthly for long all-time), means vs. sums. **Trend math always uses raw points.**
- `bodyAnalytics` (`BodyAnalyticsCalculator`) — weight headline (latest + 7-day trend
  weight), weight trend classified toward/away from goal, honest distance-to-target (ETA
  only when the trend is meaningful *and* heading toward goal; none when stable/away), and
  per-site trends (waist et al.) with §5.3 directionality + §6.4 deadbands.
- `exerciseReport` gains `bestE1rmSeries` + `computeExerciseTrend` — the per-workout
  best-e1RM series (reusing `computeExerciseBests` so it matches the report exactly) and
  its regression trend (progression rate = kg/week when the minimums are met).

**UI primitives (`core/ui/`):** `StatTile` (interpretation-triplet, required context),
`ChartFrame` (required interpretation line + range slot), `Sparkline` (react-native-svg
60×20 glyph). **`TrendChart`** (`features/analytics`, Victory Native XL / Skia,
device-only) — the accent line + dashed goal reference, consuming engine-bucketed points.

**Screens & wiring:** the **Analytics home** with a range control and a live Body section
(weight ChartFrame + trend-weight/distance StatTiles + waist StatTile; sparse ranges show
the honest need-more state); Training/Nutrition shown as honest "coming" placeholders. The
**Exercise Report** replaces its "trend coming" note with the real e1RM sparkline +
progression rate (or the honest insufficient-data sentence). Memoization is by data
version + range (`useBodyAnalytics` recomputes only when body/settings change or the range
switches).

**Goal weight (approved this phase):** migration 0009 adds `settings.target_weight_kg`
(nullable), a Settings "Goal weight" field, and full backup coverage — `schemaVersion`
9 → 10 with the **first data-shape upgrader** (v9 → v10 defaults `targetWeightKg` to null),
exercising the Phase 14 upgrader framework for the first time.

**Tests (38 new, 276 total; 7 new suites):** exhaustive regression/deadband/minimums
(§6.4), moving average (§6.2), bucketing + the 120-point cap, ranges/windowing, the full
Body calculator (trend-weight, toward/away/at-goal ETA, neutral-without-goal, site
trends, insufficient sites), the e1RM series + trend, and the v9 → v10 backup upgrader.

## Amendment proposals (for ratification)

1. **DATABASE §3.1 — add `target_weight_kg REAL NULL` to the `settings` table.** Phase 15's
   distance-to-target and weight directionality (§5.3 "toward `targetWeight`") require a
   stored goal weight, which ANALYTICS §5.3 and FITNESS_DOMAIN §5.2/§5.3 already reference
   but the schema doc omitted. Approved in-session before implementation; the column,
   Settings field, backup coverage, and v9 → v10 upgrader are built. **This is the one
   schema/doc delta at the next checkpoint** — DATABASE §3.1 should gain the column line so
   "schema in DB == schema docs" holds. No math changes.
2. **ANALYTICS note — `E1RM_STABILITY_KG` (2.5 kg) and the `neutral` trend classification.**
   FITNESS_DOMAIN §6.4 fixes deadbands for body metrics only; the e1RM strength trend has
   none, so this analytics-level constant is a new metric's threshold (not a §6.4
   redefinition). The `neutral` classification represents "moving but no good/bad verdict"
   (weight with no goal), consistent with the Phase 12 body-comparison `neutral`. Recorded
   for awareness; no frozen-doc change required unless you want them enumerated in §3.1/§5.3.

## What changed

New: `domain/analytics/{metricResult,ranges,timeSeries,regression,trend,movingAverage,bucketing,bodyAnalytics}`
(+ tests) and `exerciseReport` trend additions; `core/ui/{StatTile,ChartFrame,Sparkline}`;
`features/analytics/{components/TrendChart,hooks/useBodyAnalytics,screens/AnalyticsScreen}`;
`core/utils.daysBetweenIso`. Modified: `settings` schema + model + mapper + repository +
Settings UI (goal weight); backup (`schemaVersion` 10, settings schema, first upgrader,
drift test); `useExerciseReport` + `ExerciseReportScreen` (live strength trend). Migration
0009. Added deps: `victory-native`, `@shopify/react-native-skia`.

## Screens affected

Analytics home (now live Body section), Exercise Report (live strength trend), Settings
(goal weight field).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Least-squares slope recovery + gap-respecting + null guards | domain test | ✅ |
| Trend minimums: no-data / too-few / span-too-short | domain tests | ✅ |
| Deadband → stable; §5.3 mapping (lower/higher/neutral) | domain tests | ✅ |
| 7-day MA (≥ 2 in window, older points excluded, latest) | domain tests | ✅ |
| Bucketing: daily/weekly, sums conserved, ≤ 120 cap coarsens | domain tests | ✅ |
| Body: trend weight, ETA only toward goal, none stable/away, no-target-set, neutral, site trends | domain tests | ✅ |
| e1RM series matches report bests; trend improving; insufficient below minimums | domain tests | ✅ |
| Backup v9 → v10 upgrader defaults targetWeightKg to null | domain test | ✅ |
| `npm run check` | typecheck + lint + format + 276 tests + db:check (15 tables) | ✅ green (3× stable) |
| On-device Analytics + chart walk (range control, Victory weight chart + goal line, ETA, sparse-range honest state, exercise sparkline, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **Chart rendering (Victory Native XL / Skia) is device-only** and cannot be launched or
   visually verified here — the pure engine (bucketing, regression, trend, all
   `MetricResult` paths) is fully tested off-device; the Skia rendering folds into TD-001.
2. **`TrendChart` implements the essential §6.1 rules only** (line + dashed target); area
   fill, the regression overlay, and the tap tooltip are deferred as **TD-010** (blind Skia
   styling is unverifiable; every conclusion is already carried in text, P8).
3. **Memoization is in-memory (hook-level)**, not the MMKV cold-start cache (§8) — that
   disposable, version-keyed paint cache lands with the dashboard (Phase 16), where it
   matters. Correctness never depends on it (ANALYTICS rule 10).
4. **Training & Nutrition analytics are honest placeholders** — their calculators are
   Phases 17–18; nothing is estimated before then.

## Technical debt

**TD-007** trend/progression-rate portion **resolved** (selected-range report body +
recent-notes remain → Phase 17). **TD-002** shrinks to the Dashboard placeholder only.
**TD-010** added (chart §6.1 styling completeness). One amendment proposal (DATABASE §3.1
`target_weight_kg`). TD-001 gains the Analytics walk.

## Retrospective

**What went well?** `MetricResult` as the *only* return type made honesty the path of least
resistance — every "what if there's not enough data" question has exactly one answer shape,
and the UI's insufficient-data rendering is just a `status` switch. Splitting the engine into
tiny pure modules (regression / trend / MA / bucketing) meant each §6.4 rule got its own
exhaustive test, and the BodyAnalyticsCalculator is a thin composition over them. Reusing
`computeExerciseBests` per workout for the e1RM series guarantees the trend can never
disagree with the report's headline bests.

**What was harder than expected?** Weight directionality without a stored goal. §5.3 defines
weight as "toward `targetWeight`," but the schema had no target — so distance-to-target and
the good/bad verdict were unanswerable. Rather than fake a direction, I surfaced it as a
decision, added `target_weight_kg` (approved), and gave the trend a `neutral` classification
for the no-goal case so weight still *trends* honestly (up/down/steady) without a fabricated
good/bad judgment. The other friction was Victory Native XL's generic typing (the data type
needs an index signature and stable `yKeys`), resolved without `any`.

**What should change before the next phase?** Nothing structural. Phase 16 (Dashboard) will
consume this engine and add the MMKV cold-start cache; TD-009 (bodyweight source) is due
there too, now that body + training data meet. The `target_weight_kg` amendment should be
ratified into DATABASE §3.1 so the schema/doc parity check stays clean.

## Lessons Learned

- **What surprised you:** how much the "required context line" convention did — making
  `StatTile`/`ChartFrame` refuse to render without interpretation turned ANALYTICS rule 3
  from a review checklist into a compile-time-ish guarantee.
- **What documentation prevented mistakes:** FITNESS_DOMAIN §6.4 fixed the regression
  minimums, the `Δ = slope × window` deadband, and the exact stability thresholds so the
  trend math is the doc's, verbatim; §6.2 fixed the MA rule; ANALYTICS §4 fixed the
  windowing + 120-point bucket cap and that *trend math runs on raw points*; ANALYTICS §3.1
  fixed the `MetricResult`/`Trend` shapes.
- **What should be reused:** `MetricResult` as the universal metric contract; the tiny-pure-
  module engine layout with per-rule tests; the "compute latest + trend together" shape for
  every headline; deriving a display series (bucketed) separately from the trend series (raw).
- **What should be avoided:** inventing a direction for a goal-dependent metric (use
  `neutral` until the goal exists); resampling in the component (the engine buckets once);
  letting a chart be the sole carrier of a conclusion (ChartFrame + StatTiles carry it in
  text); putting a stability threshold silently at odds with §6.4 (e1RM's is a *new* metric's
  constant, surfaced, not a redefinition).
- **Amendment proposals:** DATABASE §3.1 `target_weight_kg` (built, approved in-session,
  awaiting doc ratification); `E1RM_STABILITY_KG` + `neutral` classification recorded for
  awareness. No other frozen-document defect surfaced.
