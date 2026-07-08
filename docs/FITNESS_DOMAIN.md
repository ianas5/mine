# FITNESS_DOMAIN.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Owner concern:** product vocabulary, formulas, units, and fitness rules · **Audience:** every other document and all future implementation.
>
> This document is the **single source of truth for what fitness concepts mean** in this app. If any other document, screen, or line of code contradicts a definition here, this document wins. Where a downstream document (DATABASE, ANALYTICS_ENGINE) needs a formula, it references the canonical definition here rather than restating it.

---

## 1. Purpose & Scope

This app is a **private, single-user, local-first fitness tracker**. There is exactly one user ("the user" / "I"). There is no concept of accounts, other athletes, sharing, or social comparison. Every rule below assumes one person's continuous history.

The domain covers five pillars:

1. **Training** — workouts, exercises, sets, programs, records.
2. **Nutrition** — daily energy, macros, water, targets.
3. **Body** — weight, composition, and circumference measurements.
4. **Progress photos** — front/side/back visual history.
5. **Derived meaning** — the metrics and signals that turn the four above into "am I progressing?"

This document defines the **concepts and math**. How those metrics are ranked, cached, worded into insight cards, and rendered belongs to ANALYTICS_ENGINE and UI_UX. Where I state a threshold (e.g. the recomposition signal), it is the *domain definition of the concept*; ANALYTICS_ENGINE owns how it becomes a card.

---

## 2. Global Conventions

### 2.1 Units (metric is the only unit system)

| Quantity | Unit | Stored precision | Display precision |
|---|---|---|---|
| Body weight, external load | kilogram (`kg`) | 0.01 | 1 decimal (0.1) |
| Load increment | `kg` | — | 0.25 kg steps allowed |
| Circumference / length | centimeter (`cm`) | 0.1 | 1 decimal |
| Water / liquids | milliliter (`ml`) | 1 | ml, shown as L above 1000 |
| Energy | kilocalorie (`kcal`) | 1 (integer) | integer |
| Macros (protein / carb / fat) | gram (`g`) | 0.1 | whole grams in summary cards; 0.1 g in detailed food/meal views |
| Body fat, body % | percentage points (`%`) | 0.1 | 1 decimal |
| Muscle mass | `kg` | 0.1 | 1 decimal |
| RPE | unitless 0–10 | 0.5 | 1 decimal |
| RIR (reps in reserve) | unitless integer | 1 | integer |

There is **no imperial support** and no unit toggle. Do not build one unless explicitly requested.

### 2.2 Numbers & rounding

- **Store raw, round late.** All persisted values keep full stored precision; rounding happens only at display time, using the table above.
- **Rounding mode:** round half up (0.05 → 0.1).
- **Aggregate math** (volume, averages, regression) uses raw stored values; never chain-round.
- **Percentages** display as integers unless a single decimal adds meaning (e.g. body-fat delta `-0.4%`).

### 2.3 Dates, "today", and weeks

- A **calendar date** is the device-local date in `YYYY-MM-DD` form, captured at the moment of logging. History is **never** retroactively shifted for timezone or travel.
- **"Today"** = the device-local calendar date now.
- A **week starts Monday 00:00 local** (ISO-8601). "This week" is Monday of the current week through now. This is fixed; there is no configurable week-start.
- A **month**, for trend windows, is treated as a **rolling day count**, not a calendar month, to keep math uniform (see §7).

### 2.4 Missing vs. zero (critical rule)

- `null` / absent = **not recorded**. `0` = **recorded as zero** (e.g. "I drank 0 ml").
- **Averages, adherence, and trends must skip `null`, never coerce it to 0.** A day with no nutrition entry is an *unlogged* day and is excluded from nutrition averages — it is not a 0-kcal day.
- Whether unlogged days count *against* consistency/adherence is defined per-metric below, and is deliberate — not an accident of coercion.

### 2.5 Identity & mutability

- Every entity has a stable, app-generated unique `id` that never changes.
- **Daily-keyed records** (a nutrition day, a body snapshot) are keyed by date: **one per calendar date**, upserted/merged — never duplicated.
- **Event records** (a workout, a photo) are *not* date-unique: multiple are allowed on the same date.
- All records are **editable and deletable**. Editing or deleting historical data **recomputes** every derived metric and PR as if the new state had always been true. There is no "locked history."

---

## 3. Training Domain

### 3.1 Entity vocabulary

- **Exercise (catalog entry):** a *definition* of a movement (e.g. "Cable Fly"). Reusable across workouts. Has a fixed identity, a muscle mapping, and a load type. Comes from the **seed library** or is **user-created (custom)**.
- **Program:** a reusable training plan — an ordered collection of **templates** (e.g. "Push / Pull / Legs"), optionally mapped to weekdays. Defines *intent* (target sets/reps per exercise).
- **Template:** a reusable single-session blueprint (e.g. "Push Day"): an ordered list of exercises with **target sets, target rep range, and optional target RPE**. A template holds no performed data.
- **Workout (session):** a *performed* training event on a date. Created blank, from a template, or by repeating a previous workout. Contains **exercise entries**.
- **Exercise entry:** one exercise *as performed within one workout* — the catalog exercise plus its ordered **sets** and an optional note.
- **Set:** one performed effort: `weight`, `reps`, optional `rpe`/`rir`, and a `warmup` flag. The atomic unit of training data.
- **Rest, duration, note:** session/set metadata (see §3.6).

### 3.2 Working set vs. warm-up (the most important training rule)

A **working set** is a set that counts toward volume, tonnage, muscle stimulus, estimated 1RM, and personal records.

A set is a **working set** if **all** hold:
1. `warmup` flag is **false**, and
2. `reps ≥ 1`, and
3. it has a valid load for its exercise's load type (see §3.4) — i.e. `effectiveLoad > 0`.

A **warm-up set** (`warmup = true`) is recorded and shown for context but is **excluded from every derived metric and PR**. Warm-ups still count toward session duration and "sets performed (raw)" displays, but never toward "working sets."

A set with `reps = 0` is a **placeholder/failed/empty** set: it is excluded from all metrics and does not count as a working set, regardless of the warm-up flag.

### 3.3 Muscle taxonomy & movement patterns

Every catalog exercise maps to exactly **one `primaryMuscleGroup`** and zero or more **`secondaryMuscleGroups`**.

**Canonical muscle groups (11):**
`Chest, Shoulders, Back, Biceps, Triceps, Forearms, Core, Glutes, Quads, Hamstrings, Calves`
Plus the fallback group **`Other`** for custom exercises with no assigned group.

**Movement-pattern classification** (derived from `primaryMuscleGroup`, used for push/pull and upper/lower balance):

| Pattern | Muscle groups |
|---|---|
| **Push** | Chest, Shoulders, Triceps |
| **Pull** | Back, Biceps, Forearms |
| **Legs** | Glutes, Quads, Hamstrings, Calves |
| **Core** | Core |

| Split | Muscle groups |
|---|---|
| **Upper** | Chest, Shoulders, Back, Biceps, Triceps, Forearms |
| **Lower** | Glutes, Quads, Hamstrings, Calves |
| **Core** | Core (counted separately; not upper or lower) |

**Rule:** *Volume by muscle group* attributes each working set's volume to the exercise's **`primaryMuscleGroup` only** (100%). Secondary muscles are **not** credited in v1 (this avoids double-counting and keeps balance metrics honest). A future weighted-secondary model may be introduced, but only via an explicit change to this document. `Other`-group volume is tracked but excluded from push/pull/upper/lower balance metrics.

### 3.4 Load types (how "weight" is interpreted)

Every exercise has a **`loadType`** that defines `effectiveLoad` (the kg used in volume and e1RM):

| `loadType` | Meaning | `effectiveLoad` |
|---|---|---|
| `external` (default) | Machine, cable, barbell, dumbbell | the logged `weight` |
| `bodyweight` | Pull-ups, dips, push-ups (no added load) | current **bodyweight** (see rule) |
| `bodyweight_plus` | Bodyweight + belt/dumbbell | bodyweight + logged added `weight` |
| `assisted` | Assisted machine (weight *reduces* load) | max(bodyweight − logged assist, 0) |
| `timed` | Plank, timed holds | no weight; `reps` field holds **seconds**; excluded from weight/e1RM PRs |

The user's profile is machine/cable-centric, so **`external` is the overwhelming default**. Bodyweight rules exist for correctness but are secondary.

**Bodyweight resolution rule:** "current bodyweight" = the user's most recent body-snapshot `weight` **on or before the workout's date**. If none exists, use `settings.defaultBodyweight` (a single configurable kg value). If that is also unset, the set still counts as a *working set for stimulus/sets* but contributes **0 to volume load**, and is flagged low-confidence.

### 3.5 Training formulas (canonical)

Let a working set have effective load `w` (kg) and reps `r`.

- **Set volume (volume load):**  `v = w × r`  → kg. (For `timed` exercises, volume is undefined/0.)
- **Exercise-entry volume:** Σ of set volumes over its working sets.
- **Workout volume:** Σ over all exercise entries.
- **Working sets (count):** number of working sets (used for weekly-sets and muscle-frequency metrics).
- **Estimated 1RM (e1RM)** — **Epley formula**, the single canonical estimator:
  `e1RM = w × (1 + r / 30)`
  - For `r = 1`, `e1RM = w`.
  - e1RM is only **trusted for PRs when `r ≤ 12`** (higher-rep estimates are noisy). Sets with `r > 12` still display an e1RM but do **not** set an e1RM PR.
  - Display e1RM rounded to nearest 0.5 kg; store raw.
- **Estimated strength progress (per exercise):** the time series of the best working-set e1RM per workout for that exercise. The "strength trend" is the regression slope of that series (§6.4). There is no single global "strength score" in v1.

### 3.6 Rest, duration, RPE/RIR, notes

- **Rest timer:** a UI convenience; a target rest in seconds may be attached to an exercise/set but is **not** required and not used in analytics.
- **Workout duration (minutes):** `end − start` when both are captured; otherwise `null`. Null durations are excluded from duration averages.
- **RPE (0–10, 0.5 steps)** and **RIR (integer)** are optional per set, for autoregulation context only. They are **never** inputs to PRs or volume. RPE and RIR are two views of the same effort (`RIR ≈ 10 − RPE`); store whichever the user enters, don't force both.
- **Note:** free English text on a set, exercise entry, or workout.

### 3.7 Personal Records (PRs)

A PR is defined **as of a date**: it is a record relative to *all working-set history strictly before, plus within, that workout's date*. Because history is editable, PRs are always recomputed from scratch, never cached as immutable.

**PR types tracked per exercise:**

| PR type | Definition |
|---|---|
| **Heaviest weight** | max `effectiveLoad` across working sets (any reps ≥ 1). |
| **Best e1RM** | max e1RM across working sets with `r ≤ 12`. *(Primary strength PR.)* |
| **Rep PR at a load** | most reps performed at a given `effectiveLoad`. |
| **Best set volume** | max `w × r` in a single working set. |
| **Best session volume** | max total volume for that exercise within one workout. |

**"New PR" event:** a working set that establishes a strictly greater value than all prior history for that PR type. Ties are **not** PRs (must be strictly greater). Warm-ups and `timed`/loadless sets cannot set weight/e1RM PRs.

### 3.8 Consistency & streaks

Consistency answers "am I training as often as I intend?" It requires a **`weeklyWorkoutTarget`** (integer, a user setting; default **4**, matching a hypertrophy-focused profile). If a program defines a weekday schedule, the **planned count for a week** is the program's sessions for that week; otherwise it is `weeklyWorkoutTarget`.

- **Countable workout:** a workout with **≥ 1 working set** (empty/draft sessions don't count).
- **Weekly consistency %** = `min(100, completedCountableWorkouts_thisWeek / plannedThisWeek × 100)`.
- **Monthly consistency %** = countable workouts in last 30 days / (planned per week × 30/7).
- **Training frequency** = countable workouts per week, averaged over the selected range.
- **Active day** = a calendar day with ≥ 1 countable workout.
- **Training streak** = consecutive weeks meeting ≥ `weeklyWorkoutTarget` countable workouts. (Weekly, not daily — daily streaks punish normal rest days and are misleading for hypertrophy training.)
- **Missed workout** = a program-scheduled session whose weekday passed with no countable workout mapped to it. Only computable when a scheduled program exists.

---

## 4. Nutrition Domain

Nutrition tracks exactly: **Calories, Protein, Carbohydrates, Fat, Water.** (Fiber is deliberately not tracked.)

### 4.1 Vocabulary

- **Food:** a reusable definition with per-serving macros (`kcal, protein, carb, fat`) and a serving definition (amount + unit, e.g. 100 g, 1 scoop). From a small seed set or user-created. Includes **quick meals / meal templates** (e.g. "Protein shake", "Chicken and rice").
- **Meal entry:** one logged food at a portion, on a date, optionally tagged to a slot (Breakfast / Lunch / Dinner / Snacks). Stores the **computed macros at the logged portion** (snapshotted, so later edits to the food definition don't rewrite history).
- **Nutrition day:** the per-date aggregate of all meal entries plus water. **One per calendar date.**
- **Water:** cumulative `ml` for the day, logged in increments (default cup = **250 ml**, configurable).
- **Targets:** daily goals for `kcal, protein, carb, fat`, and optional `water`. **Targets are time-versioned:**
  - A target set has an **`effectiveFrom`** date. The *active target* for any date is the target set with the latest `effectiveFrom ≤ that date`.
  - **Historical adherence uses the target that was active on each logged date** — never the current target.
  - A nutrition day may snapshot the active targets at log time (denormalized copy) when it simplifies implementation; if it does, that snapshot is authoritative for that day.
  - **Changing targets never rewrites historical adherence.** New targets apply from their `effectiveFrom` forward only.

### 4.2 Nutrition formulas

- **Day totals:** `kcal, protein, carb, fat` = Σ over that day's meal entries. Water = logged cumulative ml.
- **Remaining:** `target − consumed` per macro (may be negative = over).
- **Portion scaling:** logged macros = food's per-serving macros × (loggedAmount / servingAmount), computed at log time and stored.
- **Average daily X (over a range):** mean of day totals **over logged days only** (unlogged days excluded — see §2.4).
- **Macro energy cross-check (validation aid, not stored):** `kcal ≈ 4·protein + 4·carb + 9·fat`. Used only to flag implausible foods; never overrides entered kcal.

### 4.3 Adherence

Adherence answers "did I hit my targets?" It is defined **per day, then aggregated over logged days**, always against the target active on that date (§4.1).

| Target | "Hit" definition (per logged day) |
|---|---|
| **Protein** | `consumed ≥ target` (protein is a floor; more is fine). A softer "near-hit" band of `≥ 90% of target` may be shown but the canonical hit is `≥ target`. |
| **Calories** | `within ±10% of target` (both under and over matter for body-composition goals). Under −10% = under-eating; over +10% = over-eating. |
| **Carbs / Fat** | `within ±15% of target` (softer, they flex around protein and calories). |
| **Water** | `consumed ≥ target`. |

- **Adherence % (range)** = hit days / **logged days** in range × 100.
- **Unlogged days** are reported separately as a *logging-completeness* number ("logged 5/7 days") and are **not** counted as hits or misses. Honesty over flattering numbers.
- **Consecutive-miss streak** (e.g. "protein below target 3 days") counts consecutive **logged** days that missed; an unlogged day **breaks the streak's continuity but is noted**, because "no data" is not "missed."

---

## 5. Body Domain

### 5.1 Body snapshot

A **body snapshot** is the per-date record of any subset of body metrics. **One snapshot per calendar date**, upserted and **field-merged** (a new snapshot that omits a field does not erase a previously recorded value for that date; explicitly clearing a field is a separate, deliberate action).

**Fields (all optional, all metric):**

- **Composition:** `weight` (kg), `bodyFat` (%), `muscleMass` (kg), `visceralFat` (index), `bmi`.
- **Circumferences (cm):** `neck, chest, waist, hips, leftArm, rightArm, leftForearm, rightForearm, leftThigh, rightThigh, leftCalf, rightCalf`.

The measurement list is exactly the vision's list. **Bilateral sites are stored per side** (`left*`/`right*`) — never collapsed into one value. A "combined arm" display, if any, is a UI aggregation, not a stored field.

### 5.2 Body metric semantics

- **Weight** is the anchor metric and doubles as the bodyweight source for bodyweight-load exercises (§3.4).
- **BMI** may be **entered** (e.g. from a smart scale) or **derived** as `weight / (heightM²)` if `settings.height` (cm) is set. Entered value wins if present.
- **"Best" measurement** is directional and depends on the user's *intent* for that site (see §5.3): for `waist`/`bodyFat`/`visceralFat`, best = **lowest**; for muscular sites (`chest, arms, forearms, thighs, calves, muscleMass`) best = **highest**; for `weight`, best = **closest to target weight**.

### 5.3 Directionality (which way is "good")

Interpretation of a trend requires knowing the desired direction. This is fixed for the user's stated goals (improve composition, reduce waist/belly, build upper body):

| Metric | "Improving" direction |
|---|---|
| Waist, hips (belly/fat sites), body fat %, visceral fat | **decreasing** |
| Chest, arms, forearms, thighs, calves, muscle mass, neck | **increasing** |
| Weight | **toward `targetWeight`** (direction depends on whether current > or < target) |

Directionality is a **domain fact**, not a per-chart setting. ANALYTICS_ENGINE uses it to color and word trends; it must not redefine it.

### 5.4 Comparison between two dates

The Measurements module compares **any two body snapshots** (date A → date B). For each field present in both:

- **Absolute change:** `B − A` (in the field's unit).
- **Percentage change:** `(B − A) / A × 100` (undefined when `A = 0`; show "—").
- **Trend direction:** improving / stable / declining, per §5.3 and the stability deadband in §6.4.
- Fields present in only one snapshot show "—" for the change (no fabricated baseline).

---

## 6. Derived Meaning — canonical metric math

> This section defines the **math of derived metrics**. ANALYTICS_ENGINE owns *when/where they run, caching, ranking, and the wording of insight cards.* It must use these definitions verbatim.

### 6.1 Time series

A metric time series is an ordered list of `(date, value)` points, ascending by date, `null`s excluded. All windowed metrics operate on the subset of points within the active range (§7).

### 6.2 Averages & moving average

- **Simple average:** mean of in-range values (logged points only).
- **7-day moving average** (used to de-noise daily weight): for each day, mean of available points in the trailing 7 calendar days. Requires ≥ 2 points in the window, else `null` for that day.

### 6.3 Totals & rates

- **Total volume / total workouts / total sets:** straight sums over the range.
- **Rate metrics** (frequency, per-week volume): sum over range ÷ (range length in days / 7).

### 6.4 Trend (linear regression + deadband)

The canonical trend is the **least-squares linear regression** of value vs. day-index over the in-range series.

- **Slope** `m` = regression slope, in *units per day*. Report also as *units per week* (`m × 7`) for display.
- **Requires ≥ 3 points** spanning **≥ 14 days**. Fewer/shorter → trend = `insufficient-data` (show latest value only, no direction claim).
- **Classification (deadband to avoid noise):** compute total projected change over the window `Δ = m × windowDays`. Then:
  - `|Δ|` below the metric's **stability threshold** → **stable**.
  - otherwise **increasing** / **decreasing** by sign of `m`, mapped to **improving/declining** via §5.3 directionality.
- **Default stability thresholds** (the change below which we call it "stable"): weight ±0.8 kg, waist/circumferences ±0.5 cm, body fat ±0.4%, muscle mass ±0.3 kg. (These are domain defaults; ANALYTICS_ENGINE may expose them as constants but not silently change them.)

### 6.5 Body recomposition signal (domain definition)

Recomposition = *losing fat while holding or gaining muscle at roughly stable body weight.* The **signal fires** over a comparison window (default **8 weeks**, minimum **4 weeks**, using the earliest and latest snapshots in the window) when **both**:

1. **Weight is stable:** `|Δweight| ≤ 1.0 kg` **or** `|Δweight%| ≤ 1.5%`, **and**
2. **At least one fat-down or muscle-up marker:**
   - `waist` decreased by `≥ 1.0 cm`, **or**
   - `bodyFat` decreased by `≥ 0.5 %`, **or**
   - `muscleMass` increased by `≥ 0.3 kg`.

If weight *dropped meaningfully* while waist/fat dropped, that is **fat loss**, not recomposition — the signal must not fire (guarded by the stability condition). The concept lives here; the card wording and priority live in ANALYTICS_ENGINE.

---

## 7. Time Ranges

The app supports six ranges everywhere trends appear. They are **rolling day windows** ending today (uniform math; no calendar-month drift):

| Range label | Window |
|---|---|
| 7 days | last 7 days incl. today |
| 30 days | last 30 days |
| 3 months | last 90 days |
| 6 months | last 180 days |
| 1 year | last 365 days |
| All time | first record → today |

A range with too few points for a given metric yields that metric's `insufficient-data` state (§6.4), never a fabricated or zero-filled value.

---

## 8. Edge Cases (must be handled explicitly)

1. **Warm-up & placeholder sets:** excluded from all metrics/PRs (§3.2). A workout of only warm-ups/`reps=0` is **not countable**.
2. **Empty/draft workout** (no working sets): excluded from volume, consistency, and averages; may exist as a saved draft.
3. **Two workouts on one day:** both kept; day/muscle aggregates sum them; each can independently set PRs.
4. **Editing/deleting history:** triggers full recompute of PRs, trends, adherence, consistency. PRs may **recede** after a deletion — this is correct.
5. **Bodyweight-load with no known bodyweight:** counts for sets/stimulus, contributes 0 volume, flagged low-confidence (§3.4).
6. **Unilateral exercises (the rule is explicit and binding):**
   - If the entry is **marked unilateral and logged once** (`single_doubled`), volume and working-set stimulus are **multiplied by 2** for muscle/volume aggregation; **e1RM and weight PRs use the single-side load** (you don't press two arms' worth).
   - If **both sides are logged as separate sets** (`per_side`), do **not** multiply again — count as logged; PRs still use the single-side load.
   - The entry carries an explicit **`unilateralCounting` marker** so aggregation is unambiguous; the database design must make this impossible to misread (see DATABASE §3.4).
7. **High-rep sets (`r > 12`):** e1RM shown but never sets an e1RM PR (§3.5); weight/rep/volume PRs still apply.
8. **`timed` exercises:** `reps` field = seconds; no weight, volume, or weight/e1RM PR; may hold a "longest hold" record (a rep-PR analog on seconds).
9. **Missing measurement fields:** never fabricated; comparisons show "—". Merging a partial snapshot never nulls existing fields (§5.1).
10. **Unlogged nutrition days:** excluded from averages/adherence; surfaced as logging-completeness, not as misses (§4.3).
11. **Zero is real:** `0 ml` water logged is a value; only absence is `null`.
12. **Percentage change from 0 baseline:** undefined → "—", never ∞ or a division error.
13. **Implausible input (domain plausibility ranges):** weight `0–500 kg`, load `0–1000 kg`, reps `0–100`, circumference `0–300 cm`, body fat `0–75 %`, daily kcal `0–20000`, RPE `0–10`. Values outside are rejected at input (the *enforcement* mechanism is DATABASE/CODING_STANDARDS' concern; the *ranges* are domain).
14. **Timezone/travel:** logging uses device-local date at the moment; past dates are never shifted; ranges use device-local "today."
15. **Duplicate food/meal templates:** allowed (personal app); no dedup enforced, but a custom food with identical name+macros may be surfaced for reuse rather than re-created.
16. **Nutrition targets are time-versioned first-class domain data** (§4.1): the active target for a date resolves by `effectiveFrom`; days before the first target set have no target and yield `insufficient-data` for adherence. Targets belong in durable storage and in backups.
17. **Program schedule absent:** consistency falls back to `weeklyWorkoutTarget`; "missed workout" is simply not computed (no fabricated misses).

---

## 9. AI Decision Rules

These are binding rules for any AI or developer extending this project. When in doubt, obey these before improvising.

1. **This document is canonical.** When a formula, unit, or rule is needed, use the one defined here. **Never invent an alternative formula** (e.g. always Epley for e1RM, never Brzycki, unless this document is amended first).
2. **Metric only.** Default and store kg / cm / ml / kcal / g. Never add imperial units or a unit toggle unless explicitly requested.
3. **English only** for all data labels, insight text, and UI copy.
4. **Single user, forever.** Never introduce accounts, multi-user, sharing, or social concepts. There is no "other athlete."
5. **Missing ≠ zero.** Never coerce `null` to `0` in any average, trend, or adherence calculation. Skip nulls; report logging completeness separately.
6. **Working-set discipline.** Exclude warm-ups and `reps = 0` from every metric and PR. When unsure whether a set counts, apply §3.2 exactly.
7. **Strictly-greater PRs.** A PR requires beating prior history *strictly*; ties are not PRs. Always recompute PRs from full history — never trust a cached "highest ever."
8. **One record per date where date-keyed.** Body snapshots and nutrition days are upserted and **field-merged**; never create duplicates and never null an existing field by omission.
9. **Round late.** Compute on raw values; round only for display, per §2.1.
10. **Directionality is fixed.** Use §5.3 to decide "good/bad"; do not let a chart or screen redefine which direction is improvement.
11. **Don't fabricate trends.** Below the minimum points/span, return `insufficient-data` and show only the latest value. Never zero-fill or interpolate to force a line.
12. **Honesty over vanity.** Adherence and consistency must reflect real logged data; never inflate a percentage by treating unlogged days as successes.
13. **Recompute on mutation.** Any edit/delete of history recomputes all derived metrics; accept that records may recede.
14. **Concept vs. presentation boundary.** Define/compute the *meaning* here; leave wording, ranking, colors, and caching of insights to ANALYTICS_ENGINE and UI_UX. Do not hardcode insight sentences in the domain layer.
15. **Prefer the simplest correct rule.** This is a personal tool: when two interpretations are equally valid, choose the one that is simplest to maintain and hardest to misread (e.g. weekly streaks over daily streaks).

---

## Changelog

- 2026-07-08 — v1 baseline frozen (fiber removed from nutrition domain; macro display precision refined; nutrition targets time-versioned from the start; refinements 3–8 approved as drafted).
