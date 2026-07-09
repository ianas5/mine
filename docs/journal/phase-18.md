# Phase 18 — Nutrition Analytics + Insight Engine

**Closed:** 2026-07-09 · **Milestone:** M4 (Intelligence) — second phase

The governing principle this phase, above everything else: **insights are the product.**
Analytics produces numbers; the Insight Engine produces *understanding*. Every insight has
to answer *"so what?"* — actionable, specific, evidence-based, time-aware, calm, and
impossible to misinterpret. It should connect metrics into a conclusion the user can't
already read off a tile, never repeat a number they can see, and never speak in generic
coaching platitudes.

## Insights are the product (the standing principle)

- **The engine produces conclusions, not restatements.** Each rule reads the calculators'
  output and, where it can, *connects* signals — e.g. recomposition reads weight stability
  **and** a shrinking waist / rising lean marker together, not either alone; a protein-miss
  streak is framed against the strength trend it threatens, not just reported as a percent.
- **Every insight is evidence-backed and one tap from its proof.** Each carries an
  `InsightEvidence` that resolves to the exact screen that justifies it (muscle report,
  measurements, a specific exercise, the relevant analytics tab) — an insight you can't
  verify is a claim, not an insight.
- **Honesty propagates.** Nutrition adherence is computed over **logged days only** (§5.2) —
  never diluted by unlogged days — and logging *completeness* is surfaced separately so a
  high adherence over three logged days can't masquerade as a month of discipline.
  Insufficient data stays insufficient all the way to the card; a quiet week shows a calm
  "no new signals," not filler.
- **Calm by construction.** The dashboard shows **at most three** insights (≤2 per category,
  ≤1 housekeeping), dismissals start a **cooldown** so the same nudge can't nag, and Focus
  Mode (active workout) hides insights entirely. The engine stays pure; all I/O (the
  cooldown map) is injected.

## What was built

- **`NutritionAnalyticsCalculator`** (`domain/analytics/nutritionAnalytics.ts`, §5.2, pure):
  calorie & protein **adherence over logged days only** (each a `MetricResult<AdherenceStat>`
  carrying hit-days / logged-days), logging **completeness** vs. days-in-range, on-logged-day
  averages, and the trailing signals the rules cite — **protein-miss streak**, **calorie
  skew** (chronically over vs. under), and the day-level adherence series. Targets are
  resolved per day from the effective target on that date.
- **Recomposition signal** (`domain/analytics/recomp.ts`, §6.5, pure): over a 56-day window
  (min 28-day span), fires only when **weight is stable** (|Δ| ≤ 1 kg or ≤ 1.5%) *and* a body
  marker moves the right way (waist −≥1 cm, body-fat −≥0.5%, or lean/muscle +≥0.3 kg) — the
  multi-signal conclusion "you're recomping" no single tile can state.
- **The Insight Engine** (`domain/analytics/insights/`, pure): all **23 §6.2 rules** as
  independent evaluators over a single `InsightContext`, each emitting a coach-voice
  `Insight` (title, body, tone, category, classification, evidence, magnitude). Around them:
  **scoring** (base + magnitude, category & housekeeping caps), **dedup/conflict** resolution,
  a **cooldown** filter that suppresses a recently-dismissed insight *unless its direction
  flips*, and `selectDashboardInsights` (top-3 for the briefing). The engine takes the
  cooldown map as an argument and returns a new one to stamp — no storage inside `domain/`.
- **Cooldown persistence** (`core/storage/insightCooldowns.ts`): a small MMKV-backed map
  (instance key → classification + dismissed-on date), the only stateful piece, injected at
  the edge.
- **Orchestration hook** (`data/analytics/useInsights.ts`): assembles the full context from
  every repository + calculator (body / nutrition / 90- & 30-day workout / recomp / recent
  PRs / last-two-completed-weeks consistency / neglected groups / last snapshot & photo),
  evaluates against the MMKV cooldowns, and exposes `{ all, dashboard, dismiss }` plus
  `insightEvidenceHref`. It lives in the **data layer** deliberately: features can't import
  one another, and both the dashboard and the Analytics tab need it (ARCHITECTURE §4).
- **UI:** `InsightCard` (tone edge, icon, dismiss); the dashboard **InsightSlot** (top-3,
  hidden in Focus Mode); the Analytics **InsightList** (full live list, tap-through to
  evidence, calm quiet state) and the **NutritionSection** (adherence tiles + logging
  completeness, replacing the placeholder).

## What changed

New: `domain/analytics/{nutritionAnalytics,recomp}` (+ tests) and
`domain/analytics/insights/{types,rules,engine,index,insightTestKit}` (+ tests);
`core/storage/insightCooldowns`; `core/ui/InsightCard`; `data/analytics/useInsights`;
`features/analytics/{components/InsightList,components/NutritionSection,hooks/useNutritionAnalytics}`.
Modified: `nutritionRepository.getDailyNutritionSince` (per-day target resolution);
`AnalyticsScreen` (InsightList + live Nutrition section); `DashboardScreen` (InsightSlot);
barrels. No migration, no new deps, no frozen document changed.

## Screens affected

Dashboard (insight slot live), Analytics home (Insight list + Nutrition section live).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Adherence over logged days only; completeness separate (§5.2) | domain test (fixture) | ✅ |
| Protein-miss streak, calorie skew, on-logged-day averages | domain test | ✅ |
| Insufficient logged days → `insufficient-data`, not a fabricated % | domain test | ✅ |
| Recomp fires only on stable weight + a moving marker; window/span gating | domain test | ✅ |
| All 23 §6.2 rules fire on their trigger and stay silent otherwise | domain tests | ✅ |
| Scoring, category/housekeeping caps, dashboard top-3 selection | domain test | ✅ |
| Cooldown suppresses a dismissed insight but re-fires on a direction flip | domain test | ✅ |
| Quiet context → zero insights (calm "no new signals" state) | domain test | ✅ |
| `npm run check` | typecheck + lint + format + 331 tests + db:check (15 tables) | ✅ green (3× stable) |
| On-device insight/nutrition walk (adherence read, engineered signals → matching card, top-3 cap, evidence tap-through, dismiss/flip, Focus Mode, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **The card/section UI is device-only** — the dashboard slot, the Analytics insight list
   and Nutrition tiles, tone edges and tap-through all fold into TD-001. What *is* verified
   off-device is the whole engine: the calculators, the 23 rules, scoring/dedup/cooldown/
   selection, all reconciled against fixtures.
2. **The Analytics insight list is flat-sorted by score**, not grouped by category. At
   personal single-user scale the live list is short; a grouped view earns its place only if
   the list ever grows long enough to skim.
3. **`useInsights` fetches full history and computes the 90- and 30-day workout analytics
   twice per evaluation.** Acceptable at personal data scale (§3.3) — the fetch is one
   change-bus-gated pass and the double-compute is a few ms — but it's the obvious first
   optimization if the insight refresh is ever felt.

## Technical debt

None introduced. TD-001 gains the nutrition + insight-engine device walk. TD-007's remaining
selected-range Exercise-Report body + recent-notes stay deferred. (TD-003/005/008/010
unchanged.)

## Retrospective

**What went well?** Holding every rule to *"so what?"* is what kept the engine from becoming
a metric-restatement layer. The rules that connect signals — recomposition, protein-streak-
vs-strength, consistency-drop-with-a-neglected-group — are the ones that read like a coach;
the ones that would just echo a tile were cut or folded into a stronger rule. The calculators
built in Phases 15/17 meant the engine is almost pure composition: it reads their output and
concludes, it doesn't re-derive. Injecting the cooldown map kept `domain/` pure while still
giving dismissals real memory.

**What was harder than expected?** Two things. First, the cooldown *flip* rule — a dismissed
insight must stay gone, but not if its underlying direction reverses (you dismissed "protein
slipping," then it slips further the other way) — needed the classification stored alongside
the dismissal so the re-evaluation can tell "same nudge" from "new situation." Second, the
quiet state: an empty-nutrition context made the logging-gap rule fire, so a genuinely quiet
week looked noisy; the fix was an honest neutral baseline in the test kit and making sure no
rule fires on *absence* of data unless absence itself is the insight (a housekeeping nudge),
which it then rate-limits to one.

**What should change before the next phase?** Nothing structural. Phase 19 (the Phases
feature) adds the one remaining schema table and its analytics; the insight engine's context
is the natural place a future "you're N weeks into a cut" phase-aware rule would attach, but
that's out of this phase's scope and correctly left alone.

## Lessons Learned

- **What surprised you:** how much *"does this connect two things the user can't connect
  themselves?"* pruned the rule set — a rule that only restated one calculator's number
  didn't survive the question, and the survivors are the ones worth surfacing.
- **What documentation prevented mistakes:** ANALYTICS §5.2 fixed adherence to logged days
  only (and completeness as a separate honesty signal); §6.2 enumerated exactly the 23 rules
  and their triggers so the engine has a closed, reviewable rule set; §6.3 fixed the
  dashboard caps, the quiet state, and the dismiss-cooldown-with-flip behavior; §6.5 fixed
  the recomposition multi-signal gate.
- **What should be reused:** the injected cooldown map as the pattern for "stateful behavior
  without impure domain code"; storing an insight's *classification* alongside its dismissal
  so cooldowns can distinguish a repeat from a reversal; the data-layer hook as the home for
  cross-feature shared derivation (features can't import each other, both import data).
- **What should be avoided:** insights that restate a visible metric; firing a rule on the
  *absence* of data as if it were a finding (that's a rate-limited housekeeping nudge, or
  nothing); diluting adherence with unlogged days; letting the dashboard grow past a calm,
  scannable three.
- **Amendment proposals:** none — no frozen-document defect surfaced. No new debt.
