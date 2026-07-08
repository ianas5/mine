# PRODUCT_PRINCIPLES.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Position:** the highest-level document in this project — the product's constitution.
>
> Every other document (PROJECT_VISION.md and the technical documents) operates **under** this one. When any design, feature, metric, screen, or line of code conflicts with these principles, **the principles win**. When two principles appear to conflict, the one with the lower number wins.
>
> This document is deliberately opinionated. It exists so that a decision made at 11 pm eighteen months from now is the same decision that would be made today, fresh, with full context.

---

## Why this document exists

This is a personal fitness application for **one user, forever**. There is no market to chase, no competitor to match, no growth metric to satisfy, no one else to impress. That freedom is dangerous: without external constraints, personal projects die by accumulation — one more feature, one more chart, one more screen, until the app is a chore to use and a swamp to maintain.

These principles are the replacement for market discipline. They are the taste of the product, written down.

---

## I. Purpose

### P1 — The app exists to produce better decisions, not more data. *(The Prime Directive)*

The measure of this app is not how much it stores or shows. It is whether next week's training and eating are **better informed** than they would have been without it.

**In practice:**
- Every feature, metric, chart, and card must trace to one of the eight standing questions in PROJECT_VISION (stronger? more muscle? less fat? consistent? hitting macros? measurements improving? what changed? what to adjust?).
- A metric that cannot change what I do next week is a vanity metric. It does not ship (ANALYTICS_ENGINE rule 13 is this principle, enforced).
- When evaluating anything new, the first question is never "is it cool?" or "is it easy?" — it is **"what decision does this improve?"**

### P2 — Every feature must solve a problem I actually have.

Not a problem I might have someday, not a problem other lifters have, not a problem that makes the feature list longer.

**In practice:**
- The user profile in PROJECT_VISION is the whole market: hypertrophy training, machine/cable preference, body recomposition, macro tracking. Features serve that person specifically.
- "Might be useful someday" is a reason to write an idea down, not to build it. Build when the problem is felt, not foreseen.
- Speculative configurability is forbidden: no settings for behaviors I don't actually vary (this is why there is no unit toggle, no week-start setting, no configurable formulas).

### P3 — Never add a feature because another app has it.

MyFitnessPal, Strong, Hevy, Whoop — they are built for markets, engagement, and subscriptions. Their choices are answers to *their* problems.

**In practice:**
- "App X does this" carries **zero weight** in any decision here. The only valid argument is P1/P2 in this app's own terms.
- Specifically and permanently out (restating VISION with constitutional force): social anything, streak-shame mechanics, badges-as-engagement, barcode-scanning food databases with 4 million entries I'll never eat, AI meal photo scoring, readiness scores. If a future revision wants one of these, it must amend this document first and explain what decision it improves.

### P4 — Build for year five, not week one.

This app should be a pleasure to open in five years, with five years of my history inside it.

**In practice:**
- Data outlives everything: schema and backup decisions are made as if the UI will be rewritten twice (DATABASE's migration and export rules are this principle in code).
- Long-term maintainability beats short-term velocity; readable code beats clever code (binding on CODING_STANDARDS).
- Every screen is designed asking: *what does this look like with 1,000 workouts and 300 weigh-ins?* — not with demo data.
- No dependency on services that can disappear (see P19).

### P5 — The app grows with me.

My goals will change — cut, recomp, lean bulk, maintain, and things I haven't planned yet. The product adapts without ever making me start over.

**In practice:**
- Goals and targets are time-versioned data, never global rewrites (time-versioned nutrition targets; user-defined phases).
- Changing direction is a first-class action: declare a new phase, set new targets — history stays intact and *comparable*.
- Analytics interpret history against the goal *of that era* (adherence uses the target active on that date; phase reports judge a cut as a cut).
- Nothing in schema or UX may assume one permanent goal.

---

## II. Intelligence

### P6 — Explain progress; don't just display data.

The app's personality is an **intelligent coach reviewing my log**, not a spreadsheet with a theme.

**In practice:**
- Every statistic carries the interpretation triplet — value, reference, classification (ANALYTICS_ENGINE rule 3). A naked number is a defect.
- Charts exist to make one sentence visible; every chart states that sentence (the mandatory ChartFrame interpretation line). If the sentence can't be written, the chart doesn't ship.
- "Trend: +4.2%" is the canonical counter-example. "Your chest volume has grown steadily for 6 weeks" is the standard.

### P7 — Honesty over encouragement.

A coach who only flatters is useless. The app reports decline as plainly as progress, and never inflates numbers to feel good.

**In practice:**
- Unlogged ≠ success, unlogged ≠ failure; missing data is named, never papered over (FITNESS_DOMAIN §2.4).
- Insufficient data yields "insufficient data" — never a fabricated trend.
- No manipulative celebration: one toast and a haptic for a PR is right; confetti economics are not (P3).
- The tone is factual and calm even when the news is bad: *"Protein was below target on 3 straight logged days"* — stated, not scolded, not softened into invisibility.

### P8 — Trust is sacred.

Every statement the app makes must be safe to act on. Never exaggerate, never guess, never overinterpret.

**In practice:**
- Below data minimums the app says "insufficient data" — never a manufactured trend (FITNESS_DOMAIN §6.4).
- Low-confidence computations are labeled as such (e.g. bodyweight-load volume without a known bodyweight).
- Inference is hedged ("may indicate"), causation is never asserted, and thresholds are conservative by design.
- One misleading insight costs more than fifty missing ones: when in doubt, the engine stays silent.

---

## III. Daily Use

### P9 — Logging must always be faster than reviewing.

The app lives or dies at the gym and at the table. Logging is the product's heartbeat; analysis is its reward. If logging is slow, there is nothing to analyze.

**In practice:**
- Optimize the logging paths before any analytics path, always.
- Concrete budgets: log a set in **≤ 2 taps** from the active workout (previous values pre-filled, stepper-adjusted); log a repeated meal in **≤ 3 taps** from the dashboard; log weight in **≤ 3 taps**. These budgets are UI_UX acceptance criteria, not aspirations.
- Previous performance is *shown at the point of logging* (last weight×reps, best) so the decision "what do I lift now?" needs no navigation.
- Quick actions on the dashboard are sacred; they never get buried by content.

### P10 — Default to the fewest taps, the fewest screens, the fewest choices.

Respect for my time is measured in interactions removed.

**In practice:**
- Every flow is designed by counting taps, then removing one.
- Smart defaults over questions: today's date, last-used weights, the active program's next session — pre-filled, editable, never asked twice.
- No confirmation dialogs except destructive actions (DESIGN_SYSTEM: `Dialog` is destructive-only).
- No onboarding wizard, no tutorial overlay, no empty-state quiz. The app teaches by being obvious.

### P11 — Every screen has one primary purpose.

If a screen's purpose can't be stated in one sentence, it's two screens — or half of one that shouldn't exist.

**In practice:**
- Each screen gets a written one-sentence purpose in UI_UX_GUIDELINES; content that doesn't serve it moves or dies.
- One primary action per screen, visually unmistakable (the single-accent rule in DESIGN_SYSTEM).
- The five-tab structure is fixed; features fold into it rather than sprouting new top-level destinations.

### P12 — The dashboard answers exactly one question: "What should I know today?"

**In practice:**
- Its contents are a closed list (ANALYTICS_ENGINE §6.5): today's workout, remaining calories/protein, today's macros, top insights (≤ 3), current streak, quick actions, and the trend weight in the header greeting. **Nothing else, ever, without amending that list.**
- The dashboard is a briefing, not a report. Ten seconds to read, then either act (log) or leave.

### P13 — Analytics live in the Analytics section.

Depth belongs where I go to think, not where I go to act.

**In practice:**
- Trends, reports, phases, and comparisons live in the Analytics tab and entity pages (exercise report, muscle report). Other screens may carry at most a *pointer* to an interesting finding — the top insight cards — never the analysis itself.
- Logging screens show the minimum context needed to log well (last performance), and no more.

### P14 — Reduce cognitive load everywhere.

The app thinks so I don't have to. Attention is spent on training decisions, not on operating software.

**In practice:**
- Recognition over recall: pickers show recent/frequent items first (foods I eat, exercises I do — the personal app advantage).
- One number per concept: *the* weight trend, *the* consistency figure — never three variants of the same metric competing.
- Consistent placement: the same action lives in the same place on every screen (UI_UX owns the conventions).
- Progressive disclosure: summary first, detail on demand — never everything at once.

### P15 — Instant feedback.

Every interaction feels immediate. The app always feels native.

**In practice:**
- Every tap produces feedback within one frame (pressed state, haptic, or value change).
- Local-first makes real waits rare; where work takes time (import, export, photo processing) progress is shown immediately and the UI never freezes.
- Writes are optimistic — the UI reflects a logged set the instant it's tapped; persistence catches up invisibly.
- Skeletons over spinners for first paints; no interaction ever ends in silent nothing.

---

## IV. Character

### P16 — Premium through simplicity, not decoration.

Premium is what's *removed*: noise, clutter, gradients-for-no-reason, competing colors, redundant labels.

**In practice:**
- Calm dark surfaces, one earned accent, engineered typography, generous space (DESIGN_SYSTEM implements this).
- Visual noise is treated as a bug class: if an element doesn't inform or enable, it's removed.
- No decoration that pretends to be information (tinted cards by category, icon soup, emoji chrome).

### P17 — Calm, intelligent, confident.

The app never shouts, never nags, never panics. It states, suggests, and stays out of the way.

**In practice:**
- Maximum 3 insight cards on the dashboard; cooldowns prevent repetition (ANALYTICS_ENGINE §6.3).
- `danger` styling is reserved for destructive actions — being off-track is *attention*, not an emergency.
- Notifications (if ever added) would follow this personality: rare, useful, actionable — and would require a principle-level review first.
- Empty states are quiet and directive ("Log your first weigh-in"), never apologetic or cute.

### P18 — Delight over engagement.

I should *enjoy* logging because the experience is beautiful, smooth, and satisfying — never because the app manipulates me into returning. Satisfaction, not addiction.

**In practice:**
- Invest in the feel of the core loop — the stepper's tick, the set-complete haptic, the smoothness of the sheet — because those happen thousands of times.
- Forbidden forever: engagement mechanics (guilt streaks, FOMO notifications, variable rewards, daily-goal nag loops).
- The weekly streak exists as *information about consistency* (P1), never as leverage.
- If a design choice's purpose is "makes the user come back," it fails this principle; if it's "makes the visit better," it passes.

---

## V. Ownership

### P19 — Offline-first, forever.

The app must work completely — log, analyze, review, export — with airplane mode on, in a basement gym, in ten years when some service is dead.

**In practice:**
- No feature may *require* a network. Ever. A future feature that wants connectivity (e.g. cloud backup) must be an optional layer over a fully functional offline core.
- No accounts, no login, no telemetry, no analytics SDKs phoning home. The app doesn't know or care who I am.

### P20 — My data belongs entirely to me.

Five years of training history is irreplaceable personal property.

**In practice:**
- Full-fidelity export at any moment, in an open format (DATABASE §6), including photos.
- No lock-in by design: the backup is complete enough to rebuild everything.
- Destructive operations are guarded (safety export before import; explicit confirmation when it fails).
- Progress photos are private files on my device — they never leave except inside my own export.

---

## VI. Craft

### P21 — Subtract before you add. Quality over quantity.

The default answer to "should we add X?" is **no**. The bar is P1+P2, and features must also *stay* justified.

**In practice:**
- Fewer features, polished, beat many features, adequate. Ship one excellent module rather than three passable ones.
- If a shipped feature turns out not to improve decisions, **remove it** — removal is a feature. Usage honesty applies to the product itself.
- Every addition is weighed with its permanent costs: maintenance, visual space, cognitive load, schema weight.

### P22 — The codebase is a product too.

I am also the maintainer. A codebase I dread opening kills the app as surely as a bad UI.

**In practice:**
- Readability over cleverness; boring, obvious solutions over elegant surprises (binding on CODING_STANDARDS).
- The documentation set is the single source of truth; code that contradicts the docs is wrong even if it works. Docs are amended *first*, then code (DEVELOPMENT_WORKFLOW owns the process).
- No architecture astronautics: the structure in ARCHITECTURE is as elaborate as this project ever needs to be.

---

## The Tests

Fast gates for future decisions. Failing any test means no — or means amending this document first, deliberately.

**New feature test** — all six must pass:
1. What decision does it improve? (P1)
2. Do I have this problem now? (P2)
3. Would I build it if no other app had it? (P3)
4. Will I still want it in five years, with five years of data? (P4)
5. Does it keep logging fast? (P9, P10)
6. What is being removed or simplified to pay for it? (P21)

**New metric/insight test:** Which of the eight standing questions does it answer, and what would I do differently because of it? No answer → vanity metric → rejected. (P1, P6)

**New screen test:** State its purpose in one sentence; name its one primary action; say which tab it lives under. Can't → redesign. (P11, P13)

**New dashboard card test:** Is it on the closed list in ANALYTICS_ENGINE §6.5? No → it goes to Analytics, or the list is formally amended. (P12)

**New dependency/service test:** Does the app still work fully offline without it, forever? No → rejected. (P19)

---

## Precedence & Amendment

- **Precedence:** PRODUCT_PRINCIPLES → PROJECT_VISION → FITNESS_DOMAIN → ARCHITECTURE → DATABASE / ANALYTICS_ENGINE → DESIGN_SYSTEM → UI_UX_GUIDELINES → CODING_STANDARDS → DEVELOPMENT_WORKFLOW → IMPLEMENTATION_ROADMAP. Lower documents implement, never override, higher ones.
- **Amendment:** principles may change — deliberately. An amendment names the principle, the reason, and what it newly allows/forbids, and is made *before* the work that needs it. Quietly violating a principle is never an option; loudly changing one sometimes is.

## AI Decision Rules (Constitution)

1. **Cite before you build.** Any non-trivial product decision must be justifiable by a numbered principle; if you can't name the principle, stop and ask or propose an amendment.
2. **Run the Tests.** New feature/metric/screen/dashboard-card/dependency proposals go through the tests above, in writing, before design or code.
3. **Principles beat convenience, precedent, and other apps.** "It was easier," "the old app did it," and "app X has it" are non-arguments here.
4. **Conflicts resolve upward.** Between two documents, the higher one wins; between two principles, the lower number wins; between a principle and taste, the principle wins.
5. **Never weaken a principle silently in implementation** — no "temporary" features that skip the tests, no vanity metrics behind a flag, no network requirement labeled "optional" that isn't.

---

## Changelog

- 2026-07-08 — v1 baseline frozen (P1–P22, including P5, P8, P15, P18 added by approved refinement).
