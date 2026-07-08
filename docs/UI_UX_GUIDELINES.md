# UI_UX_GUIDELINES.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Owner concern:** user flows, interaction patterns, navigation behavior, information hierarchy · **Depends on:** PRODUCT_PRINCIPLES (supreme), PROJECT_VISION, FITNESS_DOMAIN, ARCHITECTURE (nav topology), ANALYTICS_ENGINE (what is shown), DESIGN_SYSTEM (how it looks).
>
> DESIGN_SYSTEM owns what components *are*; this document owns how screens are *composed* and how flows *behave*. Tap budgets here are **acceptance criteria** (P9, P10), not aspirations.

---

## 1. Experience Pillars

1. **Logging is the product** (P9): the gym flow and the meal flow get the deepest polish; everything else serves them.
2. **Briefing → act → leave** (P12): the default session is 10 seconds on the dashboard, one quick action, done.
3. **Depth on demand** (P13, P14): summary first; drill-down is always available, never forced.
4. **Instant, calm, honest** (P15, P17, P8): every tap answers immediately; nothing shouts; nothing claims more than the data supports.

---

## 2. Navigation Model

Topology fixed by ARCHITECTURE §8: **five tabs + modal logging sheets.**

- **Tabs (fixed order):** Dashboard · Workouts · Nutrition · Measurements · Analytics. Tab state is preserved when switching (scroll position, selected range).
- **Stacks:** each tab drills down via push (e.g. Workouts → Workout detail → Exercise report). Back = header back + iOS swipe-back, always.
- **Logging sheets** (`(modals)` group, DESIGN_SYSTEM `Sheet`): the Log Set flow lives *inside* the Active Workout screen (not a modal); sheets are used for **Log Meal, Add Weight, Add Measurements, Add Photo, Start Workout picker, target/phase editors**. Sheets open from anywhere (dashboard quick actions or their home tab) and return you exactly where you were.
- **Dirty-state guard:** swiping down a sheet with unsaved input asks once ("Discard entry?"); clean sheets dismiss silently.
- **The Active Workout is a persistent mode, not a page:** while a session is live, a slim **session bar** (elapsed time · current exercise · rest countdown when running) docks above the tab bar on every screen; tapping it returns to the session. Leaving the screen never pauses or loses the workout (crash-safe drafts, ARCHITECTURE §7.1).

---

## 3. Screen Inventory (one purpose, one primary action — P11)

| Screen | Purpose (one sentence) | Primary action |
|---|---|---|
| **Dashboard** | Tell me what I should know today. | Quick actions row |
| **Workouts home** | Start today's training and review recent sessions. | Start Workout |
| Active Workout | Log the current session fast. | Complete set (✓) |
| Workout detail | Review one past session. | — (read; Repeat as secondary) |
| Exercise report | Show my full history and progression for one exercise (ANALYTICS §5.5). | — (read) |
| Programs & templates | Maintain my training plans. | Edit/assign program |
| **Nutrition home (day view)** | Show today's intake vs. targets and let me log in seconds. | Log Meal |
| Food/quick-meal editor | Define a reusable food once. | Save food |
| Targets editor | Set macro targets from a date forward (time-versioned, P5). | Save targets |
| **Measurements home** | Show my current body state and log new numbers. | Add Weight / Measurements |
| Compare view | Compare any two dates' measurements (FITNESS_DOMAIN §5.4). | Pick dates |
| Photos | Review progress photos over time. | Add Photo |
| Photo compare | Put two dates side by side. | Pick dates |
| **Analytics home** | Answer "is it working?" across training, nutrition, body. | Range selector |
| Muscle report | Everything about one muscle group (ANALYTICS §5.6). | — (read) |
| Phase list & report | Judge whether a training block worked (ANALYTICS §5.4). | New Phase |
| Settings | Rarely-visited configuration, backup, phases, targets history. | — |

Settings lives behind a gear icon on the Dashboard header — it is deliberately not a tab (P10: five tabs, fixed).

---

## 4. Core Flows & Tap Budgets

Budgets measured from the flow's natural starting screen, counting taps that advance the flow (typing/stepper adjustments count as 0; they're refinement, not navigation).

### 4.1 Log a set — budget ≤ 2 taps, typical **1**

The Active Workout screen shows the current exercise as a card with its set rows. The next set row is **pre-filled from last time** (same exercise, same set number; FITNESS_DOMAIN §3.6 context: last weight×reps + previous best visible right on the card).

- Performed as planned → tap **✓**. Done (1 tap). Haptic `light`; row locks with a subtle settle animation; PRBadge appears instantly if it's a PR (optimistic, P15).
- Different weight/reps → adjust via steppers (2.5 kg / 1 rep) → **✓**.
- Warm-up toggle, RPE, and set notes are one disclosure deeper — present but never in the path of the common case (P14).
- After a working set's ✓, the **rest timer auto-starts** (exercise's `rest_seconds`, else 90 s default) in the session bar: unobtrusive countdown, tap to skip/extend, gentle haptic at zero. Never a blocking overlay (P17).

**Exercise History Preview:** every exercise card in the Active Workout carries a **compact history panel** (collapsed into the card header, no navigation needed — the athlete never leaves the screen to remember):

```
Last (Mon · 4d ago)     80 × 8 · 80 × 8 · 75 × 10
Best                    85 × 6
Best e1RM               102.1 kg
```

Working sets only; sourced from the same per-exercise history that feeds the Exercise Report (ANALYTICS §5.5). First-ever exercise shows "First time — set your baseline" (P8: no fabricated history). The full Exercise Report remains one tap away via the exercise name, for depth on demand (P13).

### 4.2 Start a workout — budget ≤ 2 taps

Dashboard shows **today's session** (from the active program's weekday mapping) as the workout card: tap **Start** (1). The session opens pre-loaded from the template with last-time values. Alternates on Workouts home: pick another template, **Repeat last workout**, or **Empty workout** (2 taps).

Finish: **Finish** → summary sheet (duration, volume, sets, PRs earned) → **Save** (2 taps; single transactional write, DATABASE §7). Discard lives in the summary sheet as a quiet secondary; discarding a session with logged sets confirms once.

A crash/kill mid-session → next launch shows a **Resume workout** banner on Dashboard and Workouts (non-blocking); Resume or Discard.

### 4.3 Log a meal — budget ≤ 3 taps for repeated meals

Quick action **Log Meal** (1) → sheet opens on **Recent & Frequent** (personal app advantage, P14: my ~15 real foods, most-used first, with quick meals pinned) → tap a food (2) → portion pre-filled with last-used amount → **Save** (3). Toast + the day's rings update optimistically.

Search/create sit above the list for the uncommon case; creating a food is a one-time cost that the Recent list then absorbs forever. Slot (breakfast/lunch/…) defaults per the smart-defaults rules (§5.2); changeable with one chip tap (counts within budget only if changed).

### 4.4 Log weight — budget ≤ 3 taps

Quick action **Add Weight** (1) → sheet with stepper pre-set to last weight (0.1 kg steps) → adjust → **Save** (2–3). Date defaults to today; merge-upsert per DATABASE §3.6.

### 4.5 Log measurements — partial by design

**Add Measurements** sheet lists all sites (composition + circumferences) with last values as placeholders; fill **any subset** and Save — omitted fields keep old values (merge-upsert, never nulled). Optimized for the "tape day" ritual: field order matches a natural top-to-bottom measuring sequence; `next` moves through fields without dismissing the keyboard.

### 4.6 Photos

**Add Photo** (1) → angle chips (front/side/back, default per §5.2) → camera/library → confirm (≤ 4 taps total). Compare: pick two dates → synchronized side-by-side pan/zoom; a **Before/After** toggle snaps to first-vs-latest for the selected angle.

### 4.7 Phases & targets (P5 flows)

- **New phase:** Settings → Phases (also linked from Analytics → Phases) → New → name, type, start date (today default) → Save. Starting a new phase while one is ongoing offers to end the current one *yesterday* (no-overlap invariant surfaces as UX, not as an error).
- **New targets:** Targets editor always writes a **new row effective from a chosen date** (default today) — the UI language is "Set new targets from …", never "edit targets", so the time-versioned model (P5) is visible and comprehensible. Past target eras are listed read-only beneath.

---

## 5. Session & Entry Patterns

### 5.1 Focus Mode (active workout)

While a session is live, the app enters a lightweight **Focus Mode** — same UI, narrowed attention (P14, P17):

- The **active workout is the primary context**: the session bar is docked app-wide; returning to the session is always one tap.
- The **dashboard defers**: insight cards are hidden for the duration (they return after Finish — nothing training-irrelevant competes for attention mid-session); the workout card becomes the live session card (elapsed time, exercises done, Return button); quick actions slim to the mid-workout-plausible set (Log Meal, Add Weight stay; Start Workout is replaced by Return to Workout).
- **No non-session toasts** while the Active Workout screen is foregrounded; only set/PR/rest feedback speaks.
- Nothing else changes — no separate theme, no locked navigation. Focus Mode is subtraction, not a second UI.

### 5.2 Smart Defaults (learned, not asked)

Defaults are computed from the user's own history with transparent frequency heuristics — personalization without AI, questions replaced by pre-fills (P10, P14). All are pre-filled and always editable; none are locked in:

- **Meal slot:** the slot this food is most often logged in; tie/no-history → time-of-day fallback.
- **Workout template:** active program's weekday mapping first; otherwise the most-frequent template for that weekday over the last 8 weeks; otherwise Repeat Last.
- **Photo angle:** the **oldest missing** angle (least-recently-captured).
- **Measurement fields:** fields co-logged in ≥ 50% of past measurement sessions start expanded; the rest sit behind "More sites."
- **Set pre-fill:** last performance for the same exercise + set number (already in §4.1).

Heuristics are pure functions over history (placed per ARCHITECTURE §9.1) with the fallback chain **history → sensible static default → ask nothing**.

### 5.3 Keyboard-first data entry

Numeric entry is a first-class path beside steppers (typing 82.4 is faster than 14 ticks):

- Numeric fields open the **decimal/number pad automatically**, current value pre-selected so typing overwrites.
- **Next advances** through the natural field order (weight → reps → next set; top-to-bottom through measurement sites); the final field's return key reads **Done/Save** and submits.
- Keyboard-avoidance keeps the focused field and its Save visible at all times; a thin accessory bar above the pad offers Next/Done and quick increments where useful.
- Rule: a flow must be completable **without ever tapping into a field manually** after the first focus.

### 5.4 Delight registry (P18)

Intentional, subtle, premium — one haptic + one visual response, ≤ 800 ms, never blocking, never confetti, never currency. The complete v1 list (additions require amending this table):

| Moment | Feedback |
|---|---|
| New PR (incl. first ever) | `success` haptic · PRBadge materializes with a soft scale-settle · toast "New PR — Bench Press 85 kg" |
| Workout completed | `success` haptic · summary sheet's stats count up briefly (duration, volume, PRs) |
| Phase completed | full-width phase report card with a quiet highlight sweep — the *report* is the reward |
| Measurement best (improving direction, FITNESS_DOMAIN §5.3) | `light` haptic · TrendArrow pulse + "best yet" caption |
| First workout after ≥ 14 days | warm, factual greeting line on Dashboard ("Back at it — first session in 3 weeks") · no guilt framing |
| Streak milestone (4/8/12/26/52 wk) | insight card styling only (ANALYTICS §6.2 #14) — information, not fanfare |

---

## 6. Interaction Standards

- **Feedback (P15):** pressed state within one frame; optimistic UI for all local writes; skeletons for first paint; determinate progress sheet for import/export; failure of an optimistic write (rare: disk full) rolls back visibly with a toast — never silently.
- **Destructive actions:** two tiers. **Undo-able** (meal entry, set row, photo): swipe-to-reveal Delete → immediate removal + 5 s **Undo toast** (fast + forgiving beats confirm dialogs, P10/P17). **Dialog-gated** (workout, program, phase, food with history, import-replace): explicit `Dialog` naming what's lost. Nothing irreversible ever hides behind a single tap.
- **Empty states:** every list/chart has one — icon, one factual line, one action ("No weigh-ins yet · Add Weight"). Never apologetic, never fake data (P8), never a dead end when logging could start.
- **Insufficient-data states:** distinct from empty — show the `needed` text from `MetricResult` verbatim ("Log 2 more weigh-ins across 2 weeks to see a trend"). Charts render their frame + message, never a fabricated flat line.
- **Errors:** infrastructure errors render a calm inline retry state at the feature boundary (ARCHITECTURE §11); form errors are inline field messages (Zod via RHF), shown on submit or blur — never toast-only, never lost.
- **Forms:** RHF + Zod resolver everywhere; numeric fields use `Stepper`/numeric keyboards with unit suffix; Save disabled only while actually invalid (not while pristine); keyboard never covers the focused field (sheets scroll).
- **Gestures:** swipe-back (iOS), swipe-to-dismiss sheets, swipe-to-reveal row actions, long-press stepper auto-repeat. No hidden-gesture-only functionality — everything reachable by visible taps too (discoverability, P14).
- **Haptics:** per DESIGN_SYSTEM tokens — `light` on set ✓/chip select/stepper tick, `success` on workout saved & new PR, `warning` on destructive confirm. Nowhere else; haptic spam is noise (P16).
- **Toasts:** one at a time, factual ("Workout saved · 3 PRs"), 2.5 s. PR celebration = `success` haptic + toast + PRBadge — that is the whole party (P18: satisfying, not manipulative).

---

## 7. Information Hierarchy Rules

1. **Level order on every screen:** primary answer (biggest number/state) → context (the interpretation triplet, always attached) → detail (lists/charts) → actions for rare cases (overflow/disclosures).
2. **Dashboard composition (closed list, ANALYTICS §6.5 / P12), top to bottom:** today's workout card (state: planned/in-progress/done/rest) → calories & protein remaining (rings) + macro bars → insight cards (≤ 3) → weekly streak line → quick actions (fixed bottom section above tab bar). The trend weight appears inside the header greeting line, never as a card (per the §6.5 closed list).
3. **Reading depth ≤ 3:** any fact is at most three levels deep (tab → screen → detail). If a fourth level seems needed, the information architecture is wrong (P14).
4. **Consistent slots:** header = title + contextual action (right); range `SegmentedControl` sits directly under the header on every analytics surface; quick actions keep fixed order everywhere they appear (Start Workout · Log Meal · Add Weight · Measure · Photo).
5. **Numbers language:** per DESIGN_SYSTEM §3 — value loud, unit quiet, context line mandatory; dates humanized ("Tue · 3 days ago") with absolute date on detail screens.
6. **Context-aware dashboard:** the dashboard's **closed content list is unchanged** (P12 stands); what adapts is **ordering and emphasis** by context, so the top of the screen answers "right now":
   - **Morning (before ~11:00):** weight/weigh-in prompt-state first (the greeting line's unlogged state rising to the top), then today's workout, then macros (naturally near-full), insights, streak.
   - **Day/Evening:** remaining calories & protein first, workout completion state, insights, streak.
   - **Active workout:** Focus Mode ordering (§5.1) — live session card first, everything else recedes.
   Cards never appear or vanish by daypart (except Focus Mode's defined subtractions) — only reorder. The layout stays recognizably the same app all day (P14). A "breakfast reminder" is expressed only as the macros card's ordinary unlogged state rising to the top in the morning — never a nag or notification (P18). *Recovery is noted as a future candidate and is not in v1.*

---

## 8. Analytics & Insight Presentation

- **Analytics home** = range selector + four fixed sections: **Training** (consistency, volume, key-exercise strength, muscle balance) · **Nutrition** (adherence, averages, completeness) · **Body** (weight + site trends, recomposition) · **Phases** (current phase progress card → phase list/reports). Each section: 2–3 StatTiles + one ChartFrame + link to its deeper report. Order fixed; sections with insufficient data show their needed-state, never vanish (P8).
- **Insight cards:** dashboard shows the top ≤ 3 (ANALYTICS §6.3); tapping a card opens its **evidence** — the relevant chart/report pre-scoped to the insight's window, so every claim is one tap from its proof (P8). Insights are dismissible (dismiss = start cooldown); no unread counters, no badge dots (P18).
- **Reports** (exercise, muscle, phase) follow one shape: identity header → headline StatTiles → trend ChartFrame → detail list. Learn one report, know them all (P14).

---

## 9. Accessibility of Interaction

44-pt targets and Dynamic Type 1.3× per DESIGN_SYSTEM §8; complete flows must remain operable with VoiceOver — logging a set, a meal, and a weigh-in are the three flows explicitly tested; charts expose their ChartFrame interpretation line as the accessible summary; Reduce Motion swaps animations for cross-fades, never removes feedback (P15 still holds).

---

## 10. AI Decision Rules (UI/UX)

1. **Tap budgets are acceptance criteria.** A change that pushes set-logging past 2 taps, meal past 3, or weight past 3 is rejected or must remove taps elsewhere first (P9, P10).
2. **Screen purpose is law:** every new screen enters the §3 table with a one-sentence purpose and one primary action before it is designed (P11); the dashboard's closed list only changes by amending ANALYTICS §6.5.
3. **Never block the athlete:** nothing modal, mandatory, or attention-demanding may interrupt an active workout; the session bar pattern is the ceiling of intrusiveness.
4. **Destructive = undo-able or dialog-gated,** per the §6 tiers; never both hidden and permanent.
5. **Every state exists:** each screen ships with empty, insufficient-data, error, and loading (skeleton) states designed — "content-only" screens are incomplete (P8, P15).
6. **No engagement dark patterns, ever:** no badges/dots, no guilt copy, no artificial delight inflation, no notification hooks (P18).
7. **Pre-fill over ask:** any value the app can know (dates, last weights, portions, slots, angles) is pre-filled and editable — never asked from zero (P10, P14).
8. **One way to do a thing:** each action has one canonical flow; don't add parallel paths that double the surface (P14, P21).
9. **Follow the fixed slots** (§7.4): headers, range control, quick-action order are conventions — breaking them needs an amendment here.
10. **Evidence within one tap:** any insight or classification shown to the user must link to the data view that justifies it (P8).
11. **Logging speed always wins.** If a UX decision makes the app more impressive but slows logging — more taps, more animation in the path, more decoration on the Active Workout — reject it. When beauty and speed conflict in a logging flow, speed is the tiebreak, every time (P9, P10, P18).

---

## Changelog

- 2026-07-08 — v1 baseline frozen (seven approved refinements applied: Focus Mode, exercise history preview, smart defaults, context-aware dashboard, keyboard-first entry, delight registry, logging-speed rule).
