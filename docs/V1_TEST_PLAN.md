# V1 Test Plan

> **Status:** the Phase 22 regression consolidation (IMPLEMENTATION_ROADMAP §5, gate 1). This
> is the single master checklist for certifying v1 on the personal iPhone. It consolidates
> every §5 phase manual checklist and the accumulated TD-001 device walks into one pass,
> grouped by area, plus the seven v1 gates. Run top to bottom on the target device model, in
> **both themes**, before tagging `v1.0.0`.

**How to use:** check each line on device. A line that fails is fixed or explicitly deferred
with rationale (never silently skipped). "Both themes" means dark and light. The automated
suite (`npm run check` — typecheck, lint, format, 375 tests, migration check) is green and is
the off-device floor; this plan is the on-device ceiling.

---

## Gate 1 — Regression (functional, per area)

### Core loop — Active Workout (P9/P10: set ≤ 2 taps)

- [ ] Start a workout; add 5 exercises; log 15+ sets at ≤ 2 taps each; warm-up toggle; keyboard entry.
- [ ] Minimize mid-session and resume; the session bar shows elapsed / current exercise / rest countdown; one tap returns.
- [ ] Rest timer auto-starts (90 s) on a working set; background ~10 min mid-rest → countdown still honest; Skip and +30 s work; gentle haptic at zero.
- [ ] Force-kill mid-session → relaunch silently restores with the recovery banner (Resume / Discard); no half-state.
- [ ] Finish → save; and the discard flow; neither leaves a draft (relaunch shows no recovery).
- [ ] **Delight:** a PR badge materializes with a soft scale-settle + success haptic; the finish summary counts up and lands; "N new PRs" toast.

### History, prefill & PRs

- [ ] Session B re-adds an exercise → prefills last-time's working sets; the in-card panel shows Last / Best / Best e1RM.
- [ ] Edit a set's weight → B's preview + Recent totals recompute; delete a set; delete a workout via the Dialog → gone everywhere.
- [ ] First-ever exercise shows the "set your baseline" state — no fabricated numbers.
- [ ] A set beating a best → PR badge + success haptic; a tie logs no PR; the Exercise Report's bests/totals/averages/last-performed reconcile with history.

### Programs & templates

- [ ] Build a PPL program (3 weekday-mapped sessions with targets); set active → Workouts home shows today's weekday session "Suggested for today"; start ≤ 2 taps, pre-loaded, targets visible; rest seconds seed the timer.
- [ ] Repeat-last reloads previous working sets; activating a second program deactivates the first (single active).
- [ ] Editing a template leaves past workouts unchanged; deleting a template leaves its past workout (provenance nulled).

### Nutrition

- [ ] Log Meal opens on Recent & Frequent, quick meals pinned; tapping a food pre-fills last-used amount + habitual slot; repeat meal ≤ 3 taps; day totals update.
- [ ] **Smart default:** a food logged mostly at dinner pre-selects Dinner; a food split evenly across slots defers to the time-of-day default; fresh install → time-of-day slot + alphabetical foods.
- [ ] Delete a meal entry → 5 s Undo toast restores the identical row; create/edit a food → past entries keep their macros; delete a food → entries survive with the snapshot.
- [ ] Targets from today show consumed / target / remaining with adherence colouring (protein floor vs calorie/carb/fat bands); new targets effective tomorrow don't change today; a day before the first target reads "no target set" (not a fabricated default).
- [ ] Water by the configured cup; a logged 0 reads differently from an unlogged day.

### Measurements & photos

- [ ] Add Weight pre-set to last weight; Save ≤ 3 taps; weight log shows the correct delta.
- [ ] Add Measurements: frequently co-logged fields expanded, rest behind "More sites"; partial subset saves; omit ≠ clear; explicit clear removes only that field; bilateral sites per side.
- [ ] **Delight:** a measurement beating its best in the good direction → gentle `light` haptic + "New best — …" toast; a routine save just confirms.
- [ ] Compare two dates: pickers list only snapshot dates; Δ / %Δ / direction glyph coloured by §5.3 (waist down = good; arm down = bad; weight/BMI neutral); sub-deadband = "stable"; one-sided field = "—" (no fabricated baseline).
- [ ] Progress photos: angle chip defaults to oldest-missing; camera + library capture copies into the app dir; timeline grouped by date; delete removes row + file; force-kill mid-save → orphan sweep cleans the stray; Compare renders side-by-side + Before/After toggle; missing file → placeholder.

### Dashboard (daily briefing, ≤ 10-second read)

- [ ] Reads top to bottom: greeting + trend weight; today's workout card (planned / done / rest); calorie & protein rings + carb/fat bars; ≤ 3 insight cards; weekly streak "N of M"; fixed quick actions (Start · Log Meal · Weight · Measure · Photo).
- [ ] Daypart reorder: morning puts workout above macros + weigh-in prompt when unweighed; day/evening puts remaining calories/protein first. Cards reorder, never appear/vanish (except Focus Mode).
- [ ] Each quick action reaches its sheet in one tap.
- [ ] Focus Mode (active session): live session card first, insight slot hidden, quick actions slim to Return · Log Meal · Weight; normal layout returns after finish.
- [ ] First session after ≥ 14 days → warm factual greeting, no guilt.

### Analytics, reports & insights

- [ ] Range control (7D/30D/3M/6M/1Y/All) recomputes every section.
- [ ] Body: weight chart in its ChartFrame with interpretation + dashed goal line; trend-weight tile with rate; distance-to-target with ETA that disappears when stable/away; sparse range → honest "not enough weigh-ins", no fabricated line.
- [ ] Training: consistency, "Getting stronger?" per-lift verdicts, Push:Pull / Upper:Lower flags, most/least-trained, one neutral volume chart — coaching answers, not a totals dump.
- [ ] Muscle report: per-group coach's review; never-trained groups honest under "Not trained in this range"; numbers reconcile with history; unilateral single-logged shows doubled volume.
- [ ] Nutrition: adherence over logged days only; completeness separate; calorie-skew note.
- [ ] Insights: engineer each rule's trigger → the matching card reads like a coach; dashboard ≤ 3 (≤ 2/category, ≤ 1 housekeeping); Analytics full list; calm "No new signals" state; tap → the pre-scoped evidence; dismiss → gone in cooldown, returns on a direction flip.

### Phases (context, not prediction)

- [ ] Declare a phase (type + start, ongoing); overlap is blocked with a guided fix; end current → declare new works.
- [ ] A phase over historical dates: report body deltas match Compare for the same dates; training/nutrition summaries reconcile; intent verdict honest (cut-with-surplus → "worth a look", aligned → "on track", thin/custom → "not enough data").
- [ ] Complete a phase → the report opens as the reward with a quiet settle; compare two phases (per-week rates); editing today's phase never changes a past phase's report.

---

## Gate 2 — Performance (ANALYTICS §7, 5-year dataset)

- [ ] Full dashboard analytics ≤ **50 ms** on the target device (dev timing log). *(Off-device Node watchdog: `src/domain/analytics/perf.test.ts` measures ~37 ms over a 1,044-workout / 20,880-set / 1,825-day synthetic set — within budget with margin.)*
- [ ] Single chart series ≤ **16 ms** (device). *(Node bucketing watchdog: ~2 ms.)*
- [ ] Cold start acceptable; scrolling the 100+ exercise library and long histories stays smooth.

## Gate 3 — Backup / data safety (DATABASE §6)

- [ ] Export produces `fitness-backup-YYYY-MM-DD.zip` (data.json + photos/); share sheet works.
- [ ] Import shows "Replace all data?"; confirm restores every screen incl. image bytes; a `pre-import-safety-…zip` is written first.
- [ ] Delete app → reinstall → import the saved backup → clean-device restore of everything incl. photos and **phases**.
- [ ] Corrupted / truncated / non-backup file → clear error, existing data untouched.
- [ ] Safety-export failure drill (no space / no share target) → "Safety backup failed — continue?"; declining leaves data untouched. *(Round-trip + refusal paths are unit-covered in `backupService.test.ts`.)*

## Gate 4 — Migrations (DATABASE §5)

- [ ] Fresh install applies 0000→current cleanly (16 tables). *(`npm run db:check` verifies on every CI run.)*
- [ ] Backup-then-migrate drill on the personal device before any migration-bearing build.
- [ ] A reverted phase's additive migration remains harmlessly in place (never down-migrate).

## Gate 5 — Accessibility (UI_UX §9, DESIGN_SYSTEM §8)

- [ ] VoiceOver completes the three critical flows: log a set, log a meal, log a weigh-in.
- [ ] Charts expose their ChartFrame interpretation line as the accessible summary.
- [ ] Dynamic Type 1.3× — no clipping on the core screens.
- [ ] Reduce Motion swaps every delight moment for an instant/cross-fade; feedback (haptic + toast + final value) preserved. *(Implemented Phase 21 via `useReducedMotion`.)*
- [ ] **Contrast:** confirm the F-E1 fix — `textTertiary` meets contrast on surface in both themes (see CP-E finding; currently 2.8 dark / 2.7 light, below AA). All other tokens pass (audited).
- [ ] 44-pt touch targets on all interactive elements.

## Gate 6 — Polish & both-themes sweep

- [ ] Every screen in dark **and** light: no clipping, no low-contrast text (pending F-E1), no visual noise.
- [ ] Empty / insufficient-data / error / loading states present on every list and chart (never a dead end, never a fabricated flat line).
- [ ] Tech-debt registry: every entry resolved or re-approved with a v1-or-later deadline (audited at CP-E — none overdue).

## Gate 7 — CP-E (final architecture + product review)

- [ ] The §3 review across the whole codebase: layered boundaries, pure calculators, SQL retrieval-only, honesty rules everywhere — all CI-enforced and re-verified at CP-E.
- [ ] The Tests (PRODUCT_PRINCIPLES) re-run against v1 as shipped: no vanity metrics, no dark patterns, tap budgets intact, every metric answers a standing question.
- [ ] First-time-user read: nothing confusing, inconsistent, noisy, or hard to discover (CP-E findings addressed).

---

**Exit:** all gates green on device → tag `v1.0.0`, update the changelog, take a fresh backup. v1 is not tagged until every gate passes.
