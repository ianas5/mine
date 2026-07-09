# Checkpoint D — M3 + M4 Architecture & Honesty Review

**Closed:** 2026-07-09 · **Scope:** everything shipped since CP-C — Phases 15–19 (analytics
foundation, dashboard, workout/muscle analytics, nutrition analytics + the Insight engine,
training phases). **Type:** review only — no features, no schema, no behaviour change.

**Verdict: PASS.** The application has remained intellectually honest. Three non-blocking
findings are recorded below; none gate M5.

The one question this checkpoint asked above all others — *has the application remained
intellectually honest?* — is answered against eight verifications, each with evidence.

## The eight standing questions, and what answers them

Every visible metric across M3+M4 maps to one of ANALYTICS §2's eight standing questions
(*getting stronger? · gaining muscle? · losing fat? · consistent? · hitting macros? ·
measurements improving? · what changed? · what to adjust?*). None is a metric for its own sake:

| Surface | Metric | Standing question |
|---|---|---|
| Dashboard | trend weight (header) | losing fat? / what changed? |
| Dashboard | today's workout card (planned/done/rest) | consistent? |
| Dashboard | calorie & protein rings, carb/fat bars | hitting macros? |
| Dashboard | weekly streak "N of M this week" | consistent? |
| Dashboard | insight cards (top-3) | what to adjust? / what changed? |
| Analytics · Body | weight chart + trend-weight tile | losing fat? / what changed? |
| Analytics · Body | distance-to-target (+ETA) | losing fat? |
| Analytics · Body | waist / site tiles | measurements improving? |
| Analytics · Training | consistency tile (streak / this-week / frequency) | consistent? |
| Analytics · Training | "Getting stronger?" + per-lift e1RM | getting stronger? |
| Analytics · Training | Push:Pull / Upper:Lower balance | what to adjust? |
| Analytics · Training | most/least trained (30d) | what to adjust? |
| Analytics · Training | weekly volume chart (neutral) | what changed? (context) |
| Muscle Report | per-group volume/sets/frequency, strongest & fastest-improving lift, recency | gaining muscle? / getting stronger? / what to adjust? |
| Analytics · Nutrition | calorie & protein adherence, completeness, calorie skew | hitting macros? / what to adjust? |
| Insights (23 rules) | each worded conclusion | what to adjust? / what changed? |
| Phase Report | body deltas, training summary, adherence + intent verdict, per-week rates | measurements? / stronger? / macros? / what changed over the block? |

## Verifications

1. **Every visible metric answers one of the eight questions.** ✅ The mapping above is
   exhaustive over the M3+M4 surfaces; no metric exists that doesn't inform a decision.

2. **No metric has become a vanity metric.** ✅ No composite "fitness score," points, levels,
   or gamified rank exists anywhere. The one streak (`weeklyStreak`) is **weekly, not daily**:
   a partial current week counts only if the target is met and never *breaks* on a rest day —
   it cannot punish rest (ANALYTICS §2 ban, verified `consistency.ts` + its tests). Volume is
   classified `neutral`, never "more = better."

3. **Every displayed value includes the interpretation triplet.** ✅ Enforced *structurally*:
   `StatTile.context` (reference + classification) and `ChartFrame.interpretation` are **required
   props** — a tile or plot cannot be constructed without its interpretation line (`StatTile.tsx`,
   `ChartFrame.tsx`). Value strings are pre-formatted by callers; the primitive does no math.

4. **No screen presents raw numbers without context.** ✅ (with F-D1). The Dashboard, Body,
   Training and Nutrition surfaces route every metric through `StatTile`/`ChartFrame`/`Ring`/
   `ProgressBar`, each pairing the number with a target, reference, direction tone, or "of N"
   denominator (swept in `DashboardScreen`, `TrainingSection`, `NutritionSection`,
   `AnalyticsScreen`). The Phase Report body-delta rows and Muscle-report volume figure render
   outside `StatTile` but still carry the triplet in substance (labelled field, signed
   delta/window, direction-tone classification) — recorded as F-D1.

5. **Every calculator remains pure.** ✅ `grep` over `src/domain/` finds **no** import of
   `react`, `expo`, `react-native`, `@/data`, `@/core/db`, or `@/core/storage`, and no
   wall-clock/randomness (`Date.now()`/`Math.random()`/argless `new Date()`). This is
   CI-enforced, not merely currently-true: `eslint-plugin-boundaries` `external` rules disallow
   react/expo/zustand from the pure layers and drizzle/sqlite outside `data`+`db`, and `lint`
   runs at `--max-warnings 0`.

6. **SQL still performs retrieval only.** ✅ No `CASE`/`WHEN`/arithmetic appears in any
   repository select; `sql\`\`` usages are aggregates/ordering only. The training join
   (`getTrainingWorkoutsSince`) selects raw columns — every working-set, effective-load,
   unilateral-doubling, volume, e1RM and bodyweight decision happens in `domain/` (ANALYTICS
   rule 9). Nutrition target *resolution* is the one repository-side derivation, and it is the
   documented single canonical path (DATABASE rule 13), not analytics.

7. **Every insight is evidence-backed and traceable.** ✅ `Insight.evidence: InsightEvidence`
   is a **required** field; all 23 rules set it (25 `evidence:` assignments). `insightEvidenceHref`
   is an **exhaustive** switch over every evidence kind with no `default`, so each card taps
   through to the exact screen that proves it (muscle report, an analytics tab, an exercise, or
   measurements/photos). Insights are derived and disposable; the MMKV cooldown map is UI
   dismissal state, never an authoritative input.

8. **Every dashboard item still deserves its place.** ✅ The dashboard is a closed set — header,
   one workout card, one macros card, the insight slot, one streak line, five quick actions —
   each mapping to a standing question (row 1–5 above). Quick actions are navigation, not
   metrics. Focus Mode hides the insight slot and slims actions during a session. No charts, no
   secondary metrics, no accumulation.

## Findings (non-blocking)

- **F-D1 — a few M4 metrics render outside `StatTile`.** The Phase Report body-delta rows
  (`PhaseReportView`) and the Muscle-report 30-day volume figure (`MuscleReportScreen`) render
  numbers via bespoke rows rather than the `StatTile` primitive. Each still carries the triplet
  in substance — a labelled field, a signed delta or windowed magnitude, and a direction/tone
  classification — and the muscle volume is deliberately neutral (quantity as context, matching
  §5.1). *Recommendation:* prefer `StatTile`/a shared delta-row primitive when the feel pass
  (Phase 21) touches these, so the triplet stays structurally guaranteed rather than by
  convention. Not a blocker — no naked, unclassified number ships.

- **F-D2 — full-history fetch + double compute at personal scale.** `useInsights` and
  `usePhaseReport` fetch all history and compute the 90-/30-day workout analytics twice per
  evaluation. Acceptable at personal scale (§3.3) and already journaled (Phases 18–19). The
  ANALYTICS §7 budgets (dashboard ≤ 50 ms, chart ≤ 16 ms) are trivially met on the seed profile;
  the **5-year synthetic-dataset** measurement is the Phase 22 performance gate, where these two
  fetches are the first optimization if any budget is breached. Watch-item, not overdue.

- **F-D3 — `daysSince` duplicates `daysBetweenIso`.** `insights/rules.ts` builds `new Date(...)`
  from explicit date strings for a day diff; it is pure and deterministic but re-implements the
  `@/core/utils` helper. *Recommendation:* reuse `daysBetweenIso` for one date path. Cosmetic.

## Technical-debt registry audit

Nothing is overdue at CP-D. Resolved since earlier checkpoints: **TD-002, TD-004, TD-006,
TD-009**. Open items all carry future deadlines — TD-003/008/010 (Phase 21 feel pass), TD-007
(unscheduled Exercise-Report pass / Phase 20), TD-005 (before an all-history archive ships),
and **TD-001** (the owner's physical-iPhone verification, which has accreted the M3+M4 device
walks — analytics, dashboard, muscle report, nutrition/insights, and now training phases).
TD-010's alternate trigger ("first checkpoint after the charts are device-verified") has **not**
fired, since the TD-001 device pass remains owner-gated.

## Honesty posture — the philosophy read

M3+M4 turned raw data into understanding without ever letting a derived number pose as truth.
The interpretation triplet is not a convention but a type constraint; insufficient-data is a
first-class `MetricResult` that propagates from the calculators to the card ("needs more
sessions," "not enough logged days," "Set a goal weight"); inference is hedged and evidence-
linked; and the newest surface — training phases — is explicitly *context, not prediction*, with
a completed block that reads the same regardless of today. The engine is pure, SQL retrieves,
insights cite their proof, and every screen earns its numbers. The application is honest.

**M4 is accepted. Proceeding to M5 (Refinement) is unblocked.**
