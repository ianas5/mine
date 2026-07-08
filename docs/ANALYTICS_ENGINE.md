# ANALYTICS_ENGINE.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Owner concern:** derived metrics, trends, adherence, insight generation, interpretation, prioritization, caching, and performance of analytics · **Depends on:** FITNESS_DOMAIN.md, ARCHITECTURE.md, DATABASE.md.
>
> FITNESS_DOMAIN owns the **math** (formulas, thresholds, directionality — its §6 is canonical and this document never redefines it). This document owns **everything between that math and the screen**: which metrics exist, how they are windowed and compared, how raw numbers become worded insight cards, how insights are ranked and deduplicated, and how it all stays fast. Chart *appearance* belongs to DESIGN_SYSTEM; *where* insights appear belongs to UI_UX_GUIDELINES.

---

## 1. Purpose

The vision's core demand: *"The app should not only store data. It should explain progress."*

The analytics engine is the app's brain. Its job is to answer the user's eight standing questions (getting stronger? gaining muscle? losing fat? consistent? hitting macros? measurements improving? what changed? what to adjust?) — with **honest, interpreted, prioritized** statements, never a wall of raw numbers.

---

## 2. Honesty Principles (binding)

1. **No vanity metrics.** Banned: composite "fitness scores," gamified points, arbitrary levels, streak mechanics that punish rest days, and any metric whose main purpose is to look impressive rather than inform a decision. Every metric must answer one of the eight standing questions.
2. **No raw numbers without interpretation.** Every displayed statistic carries the **interpretation triplet**:
   - **value** (the number, correctly rounded),
   - **reference** (vs. target, vs. previous period, or trend direction — at least one),
   - **classification** (improving / stable / declining / on-track / off-track, per FITNESS_DOMAIN §5.3 directionality and §6.4 deadbands).
   A number with no reference is a rendering bug by definition.
3. **Missing data is stated, never papered over.** Below minimum data, a metric returns `insufficient-data` with what's needed ("Log 2 more weigh-ins across 2 weeks to see a trend"). No zero-filling, no interpolation, no fabricated baselines (FITNESS_DOMAIN §2.4, §6.4).
4. **Unlogged ≠ failed, unlogged ≠ succeeded.** Nutrition averages/adherence use logged days only; logging completeness is reported as its own honest number ("logged 5/7 days").
5. **Correlation is hedged.** Insights may say "may indicate," never assert causation, and never give medical advice.
6. **Negative movement is reported as plainly as positive.** The engine is a coach's eye, not a cheerleader: a declining trend gets the same prominence rules as an improving one.

---

## 3. Engine Architecture

Per ARCHITECTURE §9: the engine is **pure functions in `domain/analytics/`** — no I/O, no framework imports. Feature hooks fetch windowed rows via repositories, hand them to the engine, render the results.

### 3.1 Result contract

Every metric function returns a discriminated union — never throws for data reasons:

```
MetricResult<T> =
  | { status: 'ok'; value: T; window: Range; computedFrom: { points: number; span: days } }
  | { status: 'insufficient-data'; reason: 'no-data' | 'too-few-points' | 'span-too-short' | 'no-target-set';
      needed: string }   // human-readable, English, e.g. "at least 3 weigh-ins over 14 days"
```

Trend-shaped values standardize on:

```
Trend = { slopePerWeek: number; deltaOverWindow: number;
          classification: 'improving' | 'stable' | 'declining';
          direction: 'increasing' | 'stable' | 'decreasing' }   // §5.3 mapping applied
```

`direction` is the raw sign; `classification` is direction mapped through FITNESS_DOMAIN §5.3 directionality. Both are exposed so the UI can say "waist ↓ 1.2 cm — improving."

### 3.2 Inputs

The engine consumes **domain models for the requested window** (workouts with exercise entries + sets + their exercises' `load_type`/muscle group/`unilateral_counting`; meal entries + water days + resolved targets; body snapshots; phases) plus `settings`. Target resolution stays in `nutritionRepository` (DATABASE rule 13) — the engine receives `(date, target)` pairs, it never resolves targets itself.

### 3.3 Scale posture

This is one person's history. Five years ≈ ~1,000 workouts / ~20,000 sets / ~10,000 meal entries — trivially in-memory. Therefore: **repositories filter by date window and join; the engine computes in memory.** SQL never implements domain semantics (working-set filtering, doubling, effective load) — those live once, in `domain/` (ARCHITECTURE rule 12). No SQL `SUM` of volume, ever: SQL can't know a warm-up from a working set with unilateral doubling.

### 3.4 Calculator modules

`domain/analytics/` is organized as **independent calculator modules**, each a self-contained pure unit with its own input type and `MetricResult` outputs:

`WorkoutAnalyticsCalculator` · `NutritionAnalyticsCalculator` · `BodyAnalyticsCalculator` · `ExerciseAnalyticsCalculator` · `MuscleAnalyticsCalculator` · `PhaseAnalyticsCalculator` · `InsightEngine` (consumes the calculators' outputs; owns §6).

Calculators never import each other's internals — shared math lives in `domain/fitness`. Adding analytics = adding a calculator (or extending one's table in §5), never rewiring existing ones.

---

## 4. Time Ranges & Windowing

- Canonical ranges (FITNESS_DOMAIN §7): **7d / 30d / 90d / 180d / 365d / all-time**, rolling, ending today.
- **Comparison window:** where a metric is compared "vs. previous period," the previous period is the equal-length window immediately preceding the current one (e.g. this 7d vs. prior 7d). All-time has no comparison window.
- **Partial current week rule:** weekly consistency for the in-progress week is presented as **progress** ("2 of 4 planned"), not a percentage; percentage comparisons ("improved vs. last week") only ever compare **completed weeks**. This prevents every Monday from reporting a consistency crash.
- **Week buckets** are ISO Monday weeks (FITNESS_DOMAIN §2.3); day buckets are device-local dates.
- **Chart series preparation** (engine output, appearance owned by DESIGN_SYSTEM): daily granularity for 7d/30d; weekly buckets for 90d/180d; weekly or monthly for 365d/all-time, capped at **~120 points** per series via bucket-averaging (means for measurements, sums for volume). Bucketing is downsampling for display only — trend math always runs on raw points.

---

## 5. Metric Catalog

Formulas per FITNESS_DOMAIN §3/§4/§6 — referenced, not restated. Listed per metric: definition source, output, minimums.

### 5.1 Workout analytics (`WorkoutAnalyticsCalculator`)

| Metric | Definition | Output / interpretation reference |
|---|---|---|
| Total workouts | countable workouts (§3.8) in range | value + vs. previous period |
| Weekly consistency | §3.8 | current week as progress; completed weeks as %; trend across weeks |
| Monthly consistency | §3.8 (30d basis) | % + classification vs. `weeklyWorkoutTarget` |
| Training frequency | §3.8 | sessions/week + vs. previous period |
| Training streak | §3.8 weekly streak | weeks count; milestone insight source |
| Missed workouts | §3.8; **only when an active program has weekday schedules** | list + count; otherwise `insufficient-data: no-target-set` |
| Total volume | Σ workout volume (§3.5, with §3.4 effective load & unilateral rules) | kg + vs. previous period |
| Volume trend | regression over weekly volume buckets (§6.4) | Trend |
| Volume by muscle group | working-set volume + set counts per primary group (§3.3), range-windowed | ranked list; feeds balance & least-trained |
| Push : Pull balance | volume ratio per §3.3 pattern mapping, excluding `other`/core | ratio + flag when outside **0.8–1.25** |
| Upper : Lower balance | same, upper vs. lower | ratio + flag when outside **1.0–2.0** (user's stated priority is upper body — the wider band reflects intent; still surfaced, softly, beyond it) |
| Most / least trained | max/min working sets per canonical group in 30d, zero counts included | named groups + set counts |
| Exercise PRs | §3.7 types; recomputed from full history | per-exercise records + "new PR" events in range |
| Strength trend (per exercise) | regression over per-workout best e1RM series (§3.5) | Trend; needs ≥ 3 sessions of that exercise across ≥ 14 days |
| Key-exercise strength summary | strength trend for the **top 5 exercises by session count in the range** | dashboard-grade summary of "am I getting stronger?" |
| Avg session duration | mean over workouts with both timestamps | minutes; null-duration sessions excluded |

### 5.2 Nutrition analytics (`NutritionAnalyticsCalculator`)

| Metric | Definition | Output / interpretation reference |
|---|---|---|
| Daily totals & remaining | §4.2 vs. resolved target | today's dashboard values |
| Average daily kcal / protein / carb / fat | mean over **logged days** (§4.2) | value + vs. active target + vs. previous period |
| Calorie adherence | ±10% hit rule (§4.3) over logged days | % + under/over-eating skew (which side misses fall on) |
| Protein adherence | ≥ target rule (§4.3) | % + "hit X of Y logged days" |
| Carb / fat adherence | ±15% rule (§4.3) | % (secondary prominence) |
| Water adherence | ≥ target, when water target set | % ; else `no-target-set` |
| Logging completeness | logged days / days in range | honest companion to every nutrition metric |
| Consecutive protein-miss streak | §4.3 consecutive-miss rule | days; insight trigger at ≥ 3 |
| Weekly / monthly nutrition trend | weekly buckets of logged-day averages; a week enters comparisons only with **≥ 4 logged days** | Trend per macro + kcal |

### 5.3 Body analytics (`BodyAnalyticsCalculator`)

| Metric | Definition | Output / interpretation reference |
|---|---|---|
| Weight now | latest snapshot weight + **7-day moving average** (§6.2) as the headline value | MA shown as "trend weight," raw as "latest" |
| Weight trend | regression + deadband (§6.4) | Trend, classified toward/away from `targetWeight` (§5.3) |
| Distance to target | latest MA − target | kg to go + weekly rate → honest ETA only when trend is meaningful ("~6 weeks at current rate"); no ETA on `stable` |
| Per-site trends | waist, chest, hips, arms, forearms, thighs, calves, neck (per side; UI may average L/R for display but the engine reports per side) | Trend per §6.4 thresholds + §5.3 directionality |
| Body fat / muscle mass / visceral trend | same machinery | Trend |
| Recomposition signal | §6.5 verbatim | boolean + contributing markers; prime insight |
| Two-date comparison | §5.4 (absolute Δ, %Δ, direction per field) | Measurements-module compare view |
| Best / latest per site | §5.2 "best" semantics | shown only alongside trend context |

### 5.4 Phase analytics (`PhaseAnalyticsCalculator`)

**Phase:** a user-defined, named period of training intent: `name`, `type` (`cutting` | `recomp` | `lean_bulk` | `maintenance` | `custom`), `startDate`, optional `endDate` (`null` = ongoing), notes. **Phases are never auto-detected** — the user declares them. Phases may not overlap (one phase at a time; repository-enforced).

The **Phase Report** contains:

- **Body deltas:** for each measured field, first vs. last snapshot within the phase (absolute Δ, %Δ, direction per FITNESS_DOMAIN §5.4) — e.g. *"Lean Bulk: +3.2 kg body weight, +2.4 cm chest, +1.8 cm arms, +0.7 cm waist."*
- **Training summary:** workouts, consistency vs. target, total volume, volume by muscle group, PRs set during the phase, key-exercise e1RM change.
- **Nutrition summary:** average kcal/protein/carb/fat over logged days, adherence %, logging completeness — interpreted against the phase's intent (e.g. a cut with kcal chronically over target is flagged).
- **Phase comparison:** any two phases side-by-side (rate-normalized per week, since phases differ in length).
- Ongoing phases report progress-to-date. A phase shorter than 14 days or without ≥ 2 snapshots returns `insufficient-data` per §3.1.

Phase analytics is a first-class long-term view, peer to the date-range views (§4). Time ranges answer "what happened lately"; phases answer "did that block work."

### 5.5 Exercise Report (`ExerciseAnalyticsCalculator`)

Per exercise, over a selected range + all-time: total sessions, total working sets, total volume, best weight, best e1RM, best set volume, best session volume, average reps per working set, average effective load, e1RM Trend (§3.1), **progression rate** (e1RM regression slope as kg/week, shown only when trend minimums are met), last performed (date + days ago), and recent notes (set/entry/workout notes mentioning the exercise). This is the data contract for the exercise history page.

### 5.6 Muscle Group Report (`MuscleAnalyticsCalculator`)

Per canonical group: current-week + trailing weekly volume, 30d volume, working sets, frequency (sessions touching the group per week), **strongest exercise** (highest best e1RM), **fastest-improving exercise** (highest positive e1RM slope among exercises meeting trend minimums), **weakest-improving exercise** (lowest slope), volume Trend, last trained (date + days since). Feeds the muscle detail views and the balance/neglect insights (§6.2 #18–19).

---

## 6. Insight Cards

An **insight** is a generated, worded, prioritized statement. Insights are **derived and disposable** — never a source of truth (ARCHITECTURE rule 8).

### 6.1 Anatomy

```
Insight = {
  ruleId:    string        // stable rule identifier, e.g. 'recomp-signal'
  instanceKey: string      // ruleId + salient data (dedup / cooldown key)
  category:  'training' | 'nutrition' | 'body' | 'consistency' | 'milestone' | 'housekeeping'
  tone:      'positive' | 'attention' | 'neutral'   // tone ≠ importance; visual language owned by DESIGN_SYSTEM
  priority:  number        // computed, §6.3
  title:     string        // short English headline
  body:      string        // 1–2 sentences, numbers included, hedged where inferential
  data:      object        // the numbers used (for detail views / charts)
  window:    Range
}
```

**Wording rules:** English; specific numbers with units ("waist −1.4 cm over 8 weeks"), not vague praise; hedge inference ("may indicate"); state timeframes; never shame ("Protein was below target on 3 straight logged days" — not "you failed"); never prescribe medically ("consider more pull work" is the strongest allowed nudge, and only for training balance).

The engine speaks like an **intelligent coach, not a spreadsheet.** Every insight is a natural English sentence that explains, not a labeled delta.
Good: *"Your chest training volume has increased steadily over the last 6 weeks."* · *"Your protein consistency improved compared to last month."* · *"Your waist is decreasing while your body weight remains stable. This may indicate successful body recomposition."*
Bad: *"Trend: +4.2%."* — bare numbers with no subject, timeframe, or meaning are forbidden anywhere insight text appears.

### 6.2 Rule catalog (v1)

Base priority **B** (0–100), cooldown = min days before the same `instanceKey` may reappear.

| # | ruleId | Trigger (all thresholds from FITNESS_DOMAIN where defined) | Tone | B | Cooldown |
|---|---|---|---|---|---|
| 1 | `recomp-signal` | §6.5 fires over 8w window | positive | 95 | 14d |
| 2 | `weight-trend-to-target` | weight Trend classified improving toward target, ≥ 2 weeks | positive | 70 | 7d |
| 3 | `weight-trend-away` | weight Trend moving away from target ≥ 2 weeks | attention | 80 | 7d |
| 4 | `weight-stalled` | weight stable ≥ 4 weeks while > 2 kg from target **and** kcal adherence < 60% | attention | 75 | 14d |
| 5 | `waist-decreasing` | waist Trend improving over 30/90d | positive | 75 | 14d |
| 6 | `site-growing` | chest / arm / thigh Trend improving over 90d | positive | 65 | 14d per site |
| 7 | `site-declining` | muscular site Trend declining over 90d | attention | 70 | 14d per site |
| 8 | `protein-miss-streak` | ≥ 3 consecutive logged days below protein target | attention | 85 | 3d |
| 9 | `protein-strong-week` | protein hit ≥ 6 of last 7 logged days (≥ 5 logged) | positive | 55 | 7d |
| 10 | `kcal-skew` | ≥ 70% of last 14 logged days' misses on one side (chronic over/under) | attention | 70 | 10d |
| 11 | `logging-gap` | nutrition logged < 4 of last 7 days | neutral | 50 | 7d |
| 12 | `consistency-up` | completed-week consistency > previous week (both completed) | positive | 60 | 7d |
| 13 | `consistency-down` | completed-week consistency < previous week by ≥ 25 pts | attention | 75 | 7d |
| 14 | `streak-milestone` | weekly streak hits 4 / 8 / 12 / 26 / 52 weeks | positive | 65 | once per milestone |
| 15 | `new-pr` | any new PR event (§3.7) in last 7d | positive | 80 | per exercise+type |
| 16 | `strength-trend-up` | key-exercise e1RM Trend improving over 90d | positive | 70 | 21d per exercise |
| 17 | `strength-trend-down` | key-exercise e1RM Trend declining over 90d | attention | 75 | 21d per exercise |
| 18 | `push-pull-imbalance` | ratio outside 0.8–1.25 over 30d (≥ 8 sessions) | attention | 65 | 14d |
| 19 | `neglected-muscle` | canonical group at 0 working sets over 30d while training ≥ 8 sessions | attention | 60 | 14d per group |
| 20 | `volume-drop` | weekly volume < 60% of trailing 4-week average (completed week) | attention | 65 | 14d |
| 21 | `measurement-due` | no body snapshot in 21d | neutral (housekeeping) | 40 | 7d |
| 22 | `photo-due` | no progress photo in 35d | neutral (housekeeping) | 35 | 7d |
| 23 | `water-low-week` | water target set and hit < 3 of last 7 logged days | neutral | 45 | 7d |

Adding a rule = adding to this table first (doc amendment), then implementing.

### 6.3 Prioritization & selection

- **Score** = B + magnitude bonus (0–10, scaled by how far past threshold) + freshness bonus (0–10, decaying over 7 days since trigger).
- **Dashboard shows at most 3** insight cards, highest score first, with **at most 2 of the same category** and at most 1 `housekeeping`. The Analytics tab shows the full current list grouped by category.
- **Dedup & cooldown:** one live insight per `instanceKey`; a fired instance re-fires within its cooldown **only if its data flips classification** (e.g. improving → declining always breaks cooldown — direction changes are never suppressed).
- **Conflict guard:** mutually exclusive rules (2 vs 3; 12 vs 13; 16 vs 17 per exercise) can never co-fire for the same window — precondition of one negates the other.
- **Quiet state:** if nothing fires, the dashboard shows nothing fabricated — a calm "No new signals — keep logging" empty state (final wording in UI_UX).

### 6.4 Evaluation timing

Insights are evaluated: on app foreground, and debounced (~1 s) after any change-bus emission from a table the rules read. Evaluation is synchronous, pure, and cheap (§7); results go to the in-memory store and the MMKV snapshot (§8). Cooldown state (`instanceKey` → last-fired date) is **MMKV** — losing it merely re-shows a card, harming nothing.

### 6.5 Dashboard philosophy

The dashboard answers exactly one question: **"What should I know today?"** Its content is fixed:

- today's workout (+ completion state)
- remaining calories
- remaining protein
- today's macros
- top insight(s) (≤ 3 per §6.3)
- current streak
- quick actions
- current trend weight, rendered only in the header greeting line (with its morning unlogged prompt state) — never as a card

**Nothing else.** All deeper analytics — trends, reports, phases, charts — live in the Analytics tab (and per-entity pages). Adding a dashboard card requires amending this section first.

---

## 7. Performance Boundaries

Correctness and readability come first; **never optimize prematurely.** The budgets below are watchdogs, not invitations — optimize only when a measured breach exists, and then via windowing/memoization before anything exotic.

- **Budget:** full dashboard analytics (all §5 dashboard metrics + all insight rules) ≤ **50 ms** on a modern iPhone with 5 years of data; single chart series ≤ 16 ms. Budgets are asserted in dev via timing logs.
- **Windowed queries only:** repositories fetch by indexed date ranges (DATABASE §4); "all-time" charts fetch bucketed breadth, not every set — but PR computation, which is definitionally all-history, fetches per-exercise history via the `workout_exercises(exercise_id)` index.
- **Memoize by data version:** each table gets a monotonic in-memory version counter bumped by its change-bus emission; computed results are memoized on `(metricId, range, versions-of-tables-read)`. Unchanged data = zero recompute.
- **No background workers in v1.** Data scale doesn't justify them; if a budget is ever breached, defer via `InteractionManager` first, workers only via doc amendment.
- **Charts get pre-bucketed series** (§4) — the UI never downsamples.

---

## 8. Caching Rules

- **Cache tier:** MMKV only (ARCHITECTURE §3): last-computed dashboard snapshot + insight list + cooldown map, keyed by data versions. Purpose: instant first paint on cold start while fresh computation runs.
- **Disposable always:** deleting the MMKV cache changes nothing but a few milliseconds. Any cached value used for rendering is replaced the moment fresh computation lands.
- **Never in SQLite** (DATABASE rule 9). **Never authoritative** (ARCHITECTURE rule 8). Import/restore clears all analytics caches unconditionally.

---

## 9. Missing-Data Handling (consolidated)

| Situation | Engine behavior |
|---|---|
| Below §6.4 minimums (points/span) | `insufficient-data` + concrete `needed` text |
| No nutrition target for a date | adherence for that date = `no-target-set`; excluded from adherence %, surfaced once as setup nudge |
| No program schedule | missed-workouts = `insufficient-data`; consistency falls back to `weeklyWorkoutTarget` (§3.8) |
| Unlogged nutrition days | excluded from averages/adherence; drive `logging-gap` + completeness display |
| Sparse weeks in trends | weeks below minimums drop out of comparisons; never zero-filled |
| Bodyweight-load without bodyweight | volume contribution 0, flagged low-confidence (§3.4); volume metrics carry a footnote flag when such sets are in-window |
| Field absent in a snapshot | comparison shows "—" (§5.4); trend uses only dates where the field exists |
| Phase too short / too few snapshots | `insufficient-data` per §5.4 minimums (14 days, ≥ 2 snapshots) |

---

## 10. AI Decision Rules (Analytics Engine)

1. **Formulas come from FITNESS_DOMAIN, verbatim, once.** This engine composes them; it never redefines volume, e1RM, adherence, deadbands, directionality, or recomposition. A needed change to math is a FITNESS_DOMAIN amendment first.
2. **Every metric returns `MetricResult`.** No throws for data conditions; no `NaN`/`Infinity`/`null` leaking to the UI — structurally impossible states, not conventions.
3. **No raw number without the interpretation triplet** (value + reference + classification). If a reference genuinely doesn't exist, the metric isn't ready to display.
4. **Never fabricate:** no zero-fill, no interpolation, no default targets, no treating unlogged as anything but unlogged.
5. **New metrics/insights must trace to one of the eight standing questions** and be added to §5/§6.2 tables (doc first, code second). If it doesn't inform a decision, it's a vanity metric — reject it.
6. **Insight wording:** English, specific numbers + units + timeframe, hedged inference, no shame, no medical prescriptions. Templates live with the rule, not scattered in components.
7. **Respect cooldowns and the dashboard cap (3)** — except classification flips, which always surface.
8. **Purity is absolute:** `domain/analytics` takes data as arguments and returns values. No repository calls, no MMKV, no clocks — "now" and windowed rows are inputs.
9. **SQL never implements domain semantics.** Repositories window/join; the engine computes. No `SUM(weight*reps)` in SQL, ever.
10. **Caches are disposable, MMKV-only, version-keyed,** and cleared on import. If correctness ever depends on a cache, the design is wrong.
11. **Comparisons compare like with like:** equal-length previous windows; completed weeks vs. completed weeks; logged days vs. logged days.
12. **Mind the budgets (§7).** A metric that can't compute inside budget gets a design fix (windowing, memoization), not a spinner.
13. **Never create analytics simply because data exists.** Every metric must help the user make a better training or nutrition decision. If it doesn't influence a decision, it must not exist.

---

## Changelog

- 2026-07-08 — v1 baseline frozen (eight approved refinements applied: phase analytics, exercise report, muscle report, dashboard philosophy, coach personality, calculator modularity, no-premature-optimization, decision-value rule; F1 consistency amendment: greeting trend weight added to the §6.5 closed list).
