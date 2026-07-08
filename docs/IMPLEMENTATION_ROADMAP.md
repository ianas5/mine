# IMPLEMENTATION_ROADMAP.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Owner concern:** the ordered execution plan — phases, dependencies, acceptance, verification checkpoints, and the v1 gate · **Depends on:** the entire accepted documentation set; it plans *when*, the other documents define *what* and *how*.

---

## 1. How to read this roadmap

- The project is built in **6 milestones (M0–M6)** containing **23 small phases**. Each phase is one to a few PRs, independently shippable, and small enough to complete and verify with confidence. **If a phase feels too large during execution, it is split — never pushed through.**
- **Working software, always:** after every merged phase the app builds, launches, and everything previously shipped still works. The project never sits in a broken intermediate state. A phase that would break existing behavior mid-way is restructured until it doesn't.
- Order within a milestone is binding; later phases list explicit dependencies. Nothing starts before its dependencies are merged and checkpointed.

## 2. The Phase Contract

Every phase in §5 carries six fields, and a phase without all six is not ready to start:

**Objective** (one sentence) · **Depends** · **Deliverables** · **Acceptance criteria** (objective, checkable) · **Manual test checklist** (run on a real iPhone, per DEVELOPMENT_WORKFLOW §4.5) · **Rollback** (what happens if the phase fails after merge).

**Global rollback policy** (per-phase entries add only specifics):
- Every phase lands as squash-merged PRs → rollback = `git revert`, app returns to previous working state.
- **Migrations are additive-first**, so a reverted phase's applied migration remains harmlessly in place; never down-migrate (DATABASE §5).
- Before any migration-bearing build reaches the personal device: backup (DEVELOPMENT_WORKFLOW §7). Data damage on a dev profile → reseed; on the personal device → restore backup.

**Technical-debt policy:** any shortcut accepted inside a phase is recorded in the repo's `docs/TECH_DEBT.md` registry at merge time with three mandatory fields — **why it exists · why it's acceptable now · when it must be removed** (a named phase or checkpoint, not "later"). Checkpoints audit the registry; an overdue entry blocks the checkpoint. Undocumented shortcuts are treated as defects.

### 2.1 Phase Closure Ritual

*A phase is not complete until all five steps are done; they extend the Phase Contract and the Definition of Done from DEVELOPMENT_WORKFLOW §5.*

1. **Git tag.** Every completed phase is tagged on `main`: `v0.<N+1>.0-phase<N>` (Phase 0 → `v0.1.0-phase0`, Phase 1 → `v0.2.0-phase1`, … Phase 22's exit tag is `v1.0.0`). Every stable milestone is permanently returnable-to.
2. **Demo entry (development journal).** Each phase produces a short write-up at `docs/journal/phase-NN.md` containing exactly: **what was built · what changed · screens affected · manual tests performed (and results) · known limitations · technical debt introduced** (mirroring the `TECH_DEBT.md` entries, or "none"). The journal is the project's historical record; entries are written at closure, while the details are fresh — never reconstructed later.
3. **Screenshot archive.** At closure, screenshots of **every completed/affected screen** are saved to `docs/journal/screenshots/phase-NN/` (dark theme minimum; both themes at checkpoint phases). This builds the visual history of the app's evolution and doubles as the reference for regression eyeballing in Phase 22.
4. **Mandatory real usage.** Before a phase is declared complete, the feature is **actually used** on a real device for at least several minutes — not walked through, *used* (log the real workout, eat the real meal, take the real photo). Real usage always finds what automated tests and scripted checklists miss. Findings are fixed or journaled as known limitations before the tag.
5. **Retrospective.** The journal entry ends with three answers: **What went well? · What was harder than expected? · What should change before the next phase?** Process changes it surfaces are applied immediately (amending DEVELOPMENT_WORKFLOW/this roadmap where needed, per the amendment flow) — small continuous improvement is part of the process, not an afterthought.

## 3. Verification checkpoints

Five **architecture verification checkpoints (CP-A … CP-E)** gate milestone boundaries. A checkpoint is a dedicated review PR (no features) that verifies, with the documents open:

1. **PRODUCT_PRINCIPLES:** run the Tests against everything shipped since the last checkpoint; no vanity metrics, no dark patterns, tap budgets intact.
2. **FITNESS_DOMAIN:** every formula in code cites and matches its section; edge-case tests (§8) present for all shipped math.
3. **ARCHITECTURE:** boundary lint clean *and* a manual read of one vertical slice (route → screen → hook → repo → SQL) confirms layering; state lives per the §6 table.
4. **DATABASE:** schema in device DB == schema docs; migrations immutable; no derived tables.
5. **ANALYTICS_ENGINE** (CP-D/E): calculators pure and modular; results carry `MetricResult`; interpretation triplet everywhere.

Plus a tech-debt registry audit (§2). Findings are fixed or become approved doc amendments **before the next milestone starts**.

## 4. Phase overview

| M | Phases | Outcome when milestone completes |
|---|---|---|
| **M0 Foundation** | 0–2 + CP-A | App shell runs on device; design system + data core proven |
| **M1 Training** | 3–8 + CP-B | Complete workout logging: catalog→session→history→PRs→programs |
| **M2 Fuel & Body** | 9–13 + CP-C | Nutrition, measurements, photos, backup — **daily-use gate opens** |
| **M3 Awareness** | 14–15 (Phases 15–16) | Dashboard v1 + analytics foundation (first real charts) |
| **M4 Intelligence** | 17–19 + CP-D | Full analytics, insights, phases |
| **M5 Refinement** | 20–21 | Smart defaults + delight registry complete |
| **M6 v1 Gate** | 22 (+ CP-E inside) | Hardened, validated v1 — no new features |

**The daily-use gate (end of M2):** real personal data may begin only once backup/restore works (P20; DEVELOPMENT_WORKFLOW §7). Until then, all device installs use synthetic seed data. This is the roadmap's single most important safety rule.

---

## 5. Milestones & phases

### M0 — Foundation

**Phase 0 · Walking skeleton**
- **Objective:** a running Expo dev-build app with the five-tab shell and full toolchain, proving the whole pipeline end-to-end.
- **Depends:** — (first phase).
- **Deliverables:** Expo + TS strict project per ARCHITECTURE §5 folder structure; ESLint (boundaries, zero-warning) + Prettier + Jest configured; CI per DEVELOPMENT_WORKFLOW §3; Expo Router with 5 tabs + placeholder screens; Inter + Lucide wired; dark/light theme provider with token stubs; `npm run check` green.
- **Acceptance:** app builds and launches on a physical iPhone via dev build; tab switching works; CI passes on the PR; boundary lint demonstrably fails a deliberate violation (test then removed).
- **Manual tests:** cold start on device; switch all 5 tabs; toggle OS dark/light and see theme follow; kill & relaunch.
- **Rollback:** none needed — nothing precedes it; a failed phase 0 is restarted, not rolled back.

**Phase 1 · Design system core**
- **Objective:** the token system and first primitive batch, visible in a dev-only gallery screen.
- **Depends:** 0.
- **Deliverables:** `core/theme` complete per DESIGN_SYSTEM §2–5; primitives: Screen, Card, Section, Button, IconButton, Input, Stepper, Chip, SegmentedControl, ListRow, Toast, EmptyState, Dialog, Sheet, Skeleton; haptic/motion tokens; hidden `/gallery` route (dev builds only) rendering every primitive in every state, both themes. *Remaining DESIGN_SYSTEM §6 primitives are delivered just-in-time by the first phase that requires them.*
- **Acceptance:** all listed primitives exist with variants/states per DESIGN_SYSTEM §6; RNTL tests for each (states + a11y labels); no raw hex/size literals outside `core/theme`+`core/ui` (lint proves it); gallery renders clean in both themes.
- **Manual tests:** walk the gallery in dark & light; Dynamic Type at 1.3×; press-state latency feels ≤ 1 frame; stepper long-press auto-repeat + haptic tick.
- **Rollback:** revert PRs; gallery and primitives are additive, shell unaffected.

**Phase 2 · Data core**
- **Objective:** SQLite/Drizzle foundation: connection, migrations, DB-ready gate, change-bus, MMKV, settings.
- **Depends:** 0.
- **Deliverables:** DB connection + PRAGMAs (DATABASE §1); migration runner in the root gate; migration `0000` (settings table); `settingsRepository`; change-bus; typed MMKV wrapper; better-sqlite3 test harness (DEVELOPMENT_WORKFLOW §4.2); Settings screen behind the dashboard gear (theme override + weekly target + height + default bodyweight).
- **Acceptance:** fresh install migrates from zero; settings persist across relaunch; repo tests green against real SQLite; change-bus emission test; MMKV holds only disposable prefs (review against ARCHITECTURE §6).
- **Manual tests:** fresh install → gate → app; change weekly target, relaunch, persisted; delete app, reinstall, defaults return.
- **Rollback:** revert; migration 0000 only ships once phase is accepted (it is the schema's root).

**CP-A · Checkpoint** — §3 review over M0. Special attention: layer boundaries real, state-ownership table respected, theme enforcement working.

---

### M1 — Training (the product's heartbeat, built first — P9)

**Phase 3 · Exercise catalog**
- **Objective:** browsable, extensible exercise library.
- **Depends:** 2.
- **Deliverables:** migration (exercises); seed library (FITNESS_DOMAIN §3.3 taxonomy, machine/cable-rich per user profile); `exerciseRepository`; Workouts-tab library screen (grouped by muscle, search); custom-exercise sheet (Zod: name, group, load type, unilateral default); archive flow.
- **Acceptance:** seeder idempotent (re-run test); custom CRUD + archive-not-delete enforced by repo test; library renders with 100+ exercises without jank.
- **Manual tests:** browse/search; create custom exercise; archive one (hidden from picker); relaunch persistence.
- **Rollback:** revert; exercises migration stays (additive, unused).

**Phase 4 · Active workout logging (core loop)**
- **Objective:** start an empty workout, log sets fast, save it transactionally — the ≤ 2-tap loop.
- **Depends:** 3.
- **Deliverables:** migrations (workouts, workout_exercises, sets); `domain/fitness`: effective load, working-set rule, volume, Epley e1RM + constants module (CODING_STANDARDS §6.2); session Zustand store; Active Workout screen (add exercise → set rows, steppers, ✓, warm-up toggle, RPE/notes behind disclosure, keyboard-first per UI_UX §5.3); finish → summary sheet → single-transaction save; `workoutRepository.saveCompletedWorkout`.
- **Acceptance:** set-log budget: pre-filled set = **1 tap**, adjusted ≤ 2 (counted on device); FITNESS_DOMAIN edge tests 1–3, 5–8, 13 green; save transactionality test (failure ⇒ no partial tree); optimistic UI (✓ reflects instantly).
- **Manual tests:** full real gym-style session (5 exercises, 15+ sets) on device; warm-up excluded from summary volume; discard flow; unilateral entry both modes (`single_doubled`, `per_side`) show correct volume.
- **Rollback:** revert screens/store; migrations stay; no other feature depends yet.
- **Accepted debt:** no history prefill yet (values default 0) — removed by Phase 5.

**Phase 5 · History, prefill & exercise-history preview**
- **Objective:** the past powers the present: history list/detail, last-time prefill, in-card history panel.
- **Depends:** 4.
- **Deliverables:** workout history (FlashList) + detail screen; per-exercise history queries; set prefill from last same-exercise/same-set-number; compact history panel (Last / Best / Best e1RM — UI_UX §4.1); edit/delete workout (Dialog-gated) with full recompute semantics.
- **Acceptance:** prefill correct across sessions (repo+domain tests); preview matches FITNESS_DOMAIN working-set rules; deleting a workout updates history views everywhere (change-bus proof); first-time exercise shows baseline state, no fabricated data (P8).
- **Manual tests:** log session A, start session B → prefills + preview show A; edit A's weight → B's preview reflects it; delete A.
- **Rollback:** revert; Phase 4 loop remains fully usable.

**Phase 6 · Crash safety, rest timer, session bar**
- **Objective:** the session becomes loss-proof and app-wide.
- **Depends:** 4 (5 recommended first).
- **Deliverables:** `workout_drafts` migration + checkpointing (debounced + on-background, ARCHITECTURE §7.1); resume/discard banner; rest timer (auto-start on working ✓, per-exercise seconds, skip/extend, end haptic); persistent session bar above tabs; Focus-Mode subtractions (UI_UX §5.1).
- **Acceptance:** force-kill mid-session → relaunch → resume restores exact state (device-verified, the phase's defining test); invalid draft discarded gracefully (corrupt-payload test); timer correct through background/foreground; draft excluded from any export path.
- **Manual tests:** kill mid-set → resume; background 10 min during rest → timer honest; navigate all tabs mid-session via bar; finish → draft gone.
- **Rollback:** revert; logging still works, only without crash-safety (registered as debt if partially landed).

**Phase 7 · Personal records & Exercise Report**
- **Objective:** strength history becomes meaning: PR detection + the per-exercise report page.
- **Depends:** 5.
- **Deliverables:** PR computation (all FITNESS_DOMAIN §3.7 types, recompute-from-history); PR celebration (badge+haptic+toast per delight registry); Exercise Report screen (ANALYTICS §5.5 contract — sessions, volume, bests, averages, last performed; trend/sparkline deferred to Phase 15 and *stated on-screen as "trend coming" — no fake trend*, P8).
- **Acceptance:** PR strictness (ties are not PRs), warm-up/rep-cap exclusions, recede-on-delete — all tested; e1RM PR only for reps ≤ 12; report numbers reconcile with history detail exactly.
- **Manual tests:** beat a weight → celebration once, correctly; tie → nothing; delete PR workout → record recedes; open report for the exercise.
- **Rollback:** revert; history remains.
- **Accepted debt:** report lacks trend chart — removed in Phase 15.

**Phase 8 · Programs & templates**
- **Objective:** planned training: programs, session templates, weekday mapping, template-driven starts.
- **Depends:** 5.
- **Deliverables:** migrations (programs, templates, template_exercises); `programRepository` (single-active invariant); management UI; start-from-template (pre-loads exercises, targets shown at logging); smart default template (UI_UX §5.2: program weekday → frequency fallback); "Repeat last workout".
- **Acceptance:** start-workout budget ≤ 2 taps from Workouts home; single-active enforced (test); template edits never mutate past workouts (provenance only, SET NULL semantics).
- **Manual tests:** build PPL program with weekdays; correct suggestion on the right weekday; start from template → targets visible; repeat-last.
- **Rollback:** revert; empty-workout flow (Phase 4) unaffected.

**CP-B · Checkpoint** — §3 review over M1. Special attention: every FITNESS_DOMAIN training edge case has its named test; tap budgets device-verified; session store vs SQLite ownership clean; debt registry (Phase 4/7 items) on schedule.

---

### M2 — Fuel & Body

**Phase 9 · Nutrition logging**
- **Objective:** the ≤ 3-tap meal loop: foods, day view, meal log.
- **Depends:** 2 (parallel-safe with M1 tail).
- **Deliverables:** migrations (foods, meal_entries); starter foods seed + quick-meal flag; Nutrition day view (totals; targets appear Phase 10); Log Meal sheet: Recent & Frequent first, search, portion scaling, snapshot macros (FITNESS_DOMAIN §4.2); food editor; swipe-delete + Undo toast; day navigation.
- **Acceptance:** repeated-meal budget ≤ 3 taps (device-counted); snapshot rule tested (edit food ⇒ history unchanged); macro scaling math per FITNESS_DOMAIN §4.2; Undo restores identical row.
- **Manual tests:** log 3 meals incl. quick meal; edit a food → old entries unchanged; delete+undo; day totals correct.
- **Rollback:** revert; rest of app untouched.

**Phase 10 · Targets, adherence display & water**
- **Objective:** intake gets meaning: time-versioned targets, remaining/consumed, water.
- **Depends:** 9.
- **Deliverables:** migrations (nutrition_targets, water_days); targets editor ("Set new targets from…", history list — UI_UX §4.7); single target-resolution implementation in `nutritionRepository`; day view shows target/consumed/remaining per macro; water card (+cup, configurable size); per-day hit logic (FITNESS_DOMAIN §4.3) shown on the day.
- **Acceptance:** resolution tests (before-first-target ⇒ insufficient-data; mid-history change ⇒ old days use old target — P5's defining test); remaining math incl. negative; water 0-vs-unlogged distinction preserved.
- **Manual tests:** set targets today; log meals → remaining updates; set new targets from tomorrow → today unchanged; water increments; check a past day's adherence against its era's target.
- **Rollback:** revert; Phase 9 logging stands alone.

**Phase 11 · Measurements logging**
- **Objective:** body tracking: ≤ 3-tap weigh-in + tape-day sheet.
- **Depends:** 2.
- **Deliverables:** migration (body_snapshots); `bodyRepository` merge-upsert + explicit clear; Add Weight sheet (stepper from last value); Add Measurements sheet (co-logged fields expanded, "More sites", keyboard-next chaining); Measurements home (latest state, weight log list with deltas); BMI entered-else-derived.
- **Acceptance:** weigh-in ≤ 3 taps; merge-upsert property tests (omit ≠ clear; same-date field merge; bilateral stored per side); weight log deltas correct.
- **Manual tests:** weigh in; tape-day partial entry; same-date second entry merges; explicit clear removes one field only.
- **Rollback:** revert; standalone feature.

**Phase 12 · Measurement comparison**
- **Objective:** any-two-dates comparison per FITNESS_DOMAIN §5.4.
- **Depends:** 11.
- **Deliverables:** Compare screen (date pickers limited to snapshot dates); per-field Δ, %Δ, direction glyphs colored by §5.3 directionality; "—" for absent fields.
- **Acceptance:** zero-baseline %Δ shows "—" (edge 12); directionality coloring verified per site type (waist down = green, arm down = red).
- **Manual tests:** compare two real snapshots; compare dates with disjoint fields; waist vs arm color logic.
- **Rollback:** revert screen; logging unaffected.

**Phase 13 · Progress photos**
- **Objective:** private visual history: capture, gallery, comparison.
- **Depends:** 2.
- **Deliverables:** migration (progress_photos); file pipeline (file-first/row-second, delete-row-then-file, orphan sweep — DATABASE §3.6); Add Photo flow (angle default = oldest missing); timeline gallery; side-by-side compare + Before/After toggle.
- **Acceptance:** lifecycle tests (insert-fail deletes file; delete removes file); missing-file renders placeholder, never crashes; sweep removes orphans.
- **Manual tests:** capture all angles; compare two dates; delete a photo; kill app between photo-pick and save → no orphan after next launch.
- **Rollback:** revert; file dir ignored by rest of app.

**Phase 14 · Backup, export & import — the daily-use gate**
- **Objective:** full data ownership: zip export, validated all-or-nothing import.
- **Depends:** 9–13 (all tables exist).
- **Deliverables:** `data/backup` service; export zip (data.json + photos) via share sheet with progress UI; import: Zod-validate → version check → **attempted safety export (failure ⇒ explicit user confirmation)** → single-transaction replace → photo reconcile (DATABASE §6); Settings entries.
- **Acceptance:** round-trip test — export, wipe, import ⇒ byte-equivalent domain data + photos; malformed/newer-version archives rejected untouched; mid-import failure leaves prior data intact (fault-injection test); safety-export-failure path shows the confirmation.
- **Manual tests:** on device: export → share to Files; delete app → reinstall → import → everything back incl. photos; import a corrupted zip → clean error, data intact.
- **Rollback:** revert; **but M3 cannot start and the daily-use gate stays closed until this phase is accepted.**

**CP-C · Checkpoint** — §3 review over M2 + formally **open the daily-use gate** (user may begin real logging; from here every migration-bearing install follows the backup-first rule with *real* stakes). Verify: snapshot rules, target resolution single-sourced, photo lifecycle, no derived tables crept in.

---

### M3 — Awareness

**Phase 15 · Analytics foundation + first charts**
- **Objective:** the trend machinery and chart infrastructure, proven on body data (weight trend is the first real chart).
- **Depends:** 11, 14.
- **Deliverables:** `domain/analytics` core: time-series, 7-day MA, regression + deadbands, `MetricResult`, `Trend` (ANALYTICS §3–4); range windowing + bucketing (≤ 120 pts); Victory Native XL integration + ChartFrame (mandatory interpretation line); Analytics home skeleton with **Body section live** (weight + waist trend, distance-to-target with honest ETA); memoization by data-version; Exercise Report gains its e1RM trend + sparkline (closes Phase 7 debt).
- **Acceptance:** regression/deadband/insufficient-data tests exhaustive (FITNESS_DOMAIN §6.4 minimums); chart renders from engine-bucketed data only; insufficient-data ranges show `needed` text, never fake lines; interpretation triplet on every visible stat.
- **Manual tests:** with real/seed history: weight chart across all 6 ranges; sparse range → honest message; ETA absent when trend stable; Exercise Report trend.
- **Rollback:** revert; Analytics tab returns to placeholder; no data at risk.

**Phase 16 · Dashboard v1**
- **Objective:** the daily briefing: closed-list content, context-aware ordering, Focus Mode integration.
- **Depends:** 8, 10, 11 (15 for the greeting's trend weight).
- **Deliverables:** dashboard per UI_UX §7.2: today's-workout card (program-aware states), kcal/protein rings + macro bars, streak line, quick actions (fixed order), latest-weight greeting; context ordering morning/day/active (UI_UX §7.6); Focus-Mode behavior; insight-card *slot* ships empty-capable (insights arrive Phase 18); `domain/fitness` weekly-streak + weekly-consistency functions (FITNESS_DOMAIN §3.8) with edge-case tests — Phase 17's `WorkoutAnalyticsCalculator` consumes these same functions.
- **Acceptance:** content is exactly the closed list (review against ANALYTICS §6.5); ordering switches by daypart & session state; every quick action reaches its sheet in 1 tap; 10-second briefing test — all "what should I know" facts above the fold on a standard iPhone.
- **Manual tests:** morning vs evening ordering; during active session → Focus Mode subtractions; rest-day, done, and in-progress workout-card states; each quick action end-to-end.
- **Rollback:** revert to simple placeholder dashboard; all logging flows unaffected.

---

### M4 — Intelligence

**Phase 17 · Workout & muscle analytics**
- **Objective:** the Training section answers "am I training right?": consistency, volume, balance, muscle reports.
- **Depends:** 15.
- **Deliverables:** `WorkoutAnalyticsCalculator` + `MuscleAnalyticsCalculator` (ANALYTICS §5.1/§5.6): consistency (partial-week honesty), frequency, streak, volume + trend, per-muscle volume/sets, push:pull & upper:lower, most/least trained, key-exercise strength summary, missed workouts (program-gated); Analytics Training section UI; Muscle Report screen.
- **Acceptance:** calculator tests incl. unilateral doubling from stored marker only, `other`-exclusion in balance, completed-week comparisons; numbers reconcile against hand-computed fixture (the 6-month seed).
- **Manual tests:** training section over real history; muscle report for chest & a never-trained group (honest zero state); mid-week consistency shows progress-not-percentage.
- **Rollback:** revert section; Body section + dashboard unaffected.

**Phase 18 · Nutrition analytics + Insight engine**
- **Objective:** nutrition trends land, and the app starts speaking: the full insight pipeline.
- **Depends:** 15, 17.
- **Deliverables:** `NutritionAnalyticsCalculator` (averages over logged days, adherence + skew, completeness, weekly trend with ≥ 4-day rule); Analytics Nutrition section; `InsightEngine`: full ANALYTICS §6.2 rule catalog, scoring, cooldowns (MMKV), conflict guards, dashboard top-3 + Analytics full list, evidence tap-through (UI_UX §8); coach-voice templates per §6.1.
- **Acceptance:** per-rule trigger/boundary/cooldown/conflict tests (all 23 rules); dashboard cap + category limits enforced; every insight's evidence link lands on the right pre-scoped view; zero insights ⇒ calm empty state, never filler (P8, P17).
- **Manual tests:** engineer a protein-miss streak in seed data → card appears, reads naturally, links to evidence; dismiss → cooldown holds; recomposition fixture fires correctly; quiet state.
- **Rollback:** revert; analytics sections remain, dashboard slot returns to empty-capable.

**Phase 19 · Phases**
- **Objective:** long-term blocks: declare phases, judge them with Phase Reports.
- **Depends:** 18.
- **Deliverables:** migration (phases) + `phaseRepository` (no-overlap, end-yesterday UX); phase management (Settings + Analytics entry); `PhaseAnalyticsCalculator` (body deltas, training & nutrition summaries judged against phase intent, rate-normalized comparison — ANALYTICS §5.4); Phase Report + current-phase progress card; phase-complete delight moment; extend backup: `data.json` gains `phases`, schemaVersion bump, import data-shape upgrader, round-trip test updated.
- **Acceptance:** overlap prevention tested; report deltas match Compare view for same dates; intent-aware judgment (cut-with-surplus flagged) tested; ongoing vs completed rendering; export → wipe → import round-trip includes phases.
- **Manual tests:** declare a phase over historical data → report sane; complete it → report moment; compare two phases; attempt overlap → guided fix.
- **Rollback:** revert; phases migration stays additive; analytics otherwise intact.

**CP-D · Checkpoint** — §3 review over M3+M4 with full ANALYTICS_ENGINE conformance: calculators modular & pure, honesty rules (no naked numbers, hedged inference, insufficient-data everywhere it should be), P1/P8 sweep of every visible metric, performance sanity vs §7 budgets on the seed profile.

---

### M5 — Refinement

**Phase 20 · Smart defaults**
- **Objective:** the app learns habits: every UI_UX §5.2 heuristic live.
- **Depends:** 16, 18.
- **Deliverables:** meal-slot-by-history, template-by-weekday-frequency (beyond program mapping), measurement co-log expansion, photo-angle default (verify), recent/frequent food ordering tuning; heuristics as pure tested functions with the fallback chain.
- **Acceptance:** each heuristic unit-tested (history → expected default; empty history → static fallback); no heuristic ever blocks or asks (pre-fill only).
- **Manual tests:** a week of varied usage → defaults visibly adapt; fresh-install fallbacks sane.
- **Rollback:** revert to static defaults — pure enhancement layer.

**Phase 21 · Delight & feel pass**
- **Objective:** the UI_UX §5.4 registry complete and the thousand-times interactions tuned.
- **Depends:** 19 (phase-complete moment), 20.
- **Deliverables:** all six registry moments implemented exactly as specified; tuning pass on stepper tick, set-✓ settle, sheet physics, count-up summary; Reduce-Motion variants; remove any accumulated visual noise (P16 sweep).
- **Acceptance:** registry complete — no more, no less; every moment ≤ 800 ms, non-blocking; Reduce Motion preserves feedback; **no logging flow gained a tap or a frame of delay** (UI_UX rule 11 regression-checked).
- **Manual tests:** trigger each moment on device; feel-check the core loop repeatedly; Reduce Motion on.
- **Rollback:** revert individual moments independently; core flows untouched.

---

### M6 — The v1 Gate

**Phase 22 · Hardening (no new features — final)**
- **Objective:** prove v1 is trustworthy: regression, performance, data-safety, accessibility, polish. **Feature freeze is absolute;** anything new goes to the post-v1 backlog.
- **Depends:** everything; includes **CP-E**, the full-set final checkpoint.
- **Deliverables & acceptance (each a gate):**
  1. **Regression:** full manual pass of every §5 phase checklist, consolidated into `docs/V1_TEST_PLAN.md`; all green on the personal device model.
  2. **Performance:** ANALYTICS §7 budgets measured with a 5-year synthetic dataset (dashboard ≤ 50 ms, chart ≤ 16 ms, cold start acceptable); breaches fixed or formally amended.
  3. **Backup:** full export/import round-trip on device with the real dataset; safety-export failure drill.
  4. **Migrations:** fixture-chain test 0000→current; fresh-install path; backup-then-migrate drill on the personal device.
  5. **Accessibility:** VoiceOver on the three critical flows; Dynamic Type 1.3×; contrast audit; Reduce Motion.
  6. **Polish & bugs:** triage everything open — fix or explicitly defer with rationale; tech-debt registry emptied or re-approved with new deadlines; both-themes sweep of every screen.
  7. **CP-E:** the §3 review across the entire codebase + The Tests re-run against v1 as shipped.
- **Manual tests:** the test plan *is* this phase.
- **Rollback:** not applicable as a unit — individual fixes revert individually; v1 simply isn't tagged until every gate passes.
- **Exit:** tag `v1.0.0`, changelog, fresh backup, done.

---

## 6. AI Decision Rules (Roadmap)

1. **Respect the gates:** no phase starts before its dependencies merged and the intervening checkpoint passed; the daily-use gate and the M6 feature freeze are absolute.
2. **Split, don't push:** the moment a phase's scope feels large mid-flight, stop and split it — a long-running broken branch violates the working-software rule.
3. **Every phase ships the six contract fields** filled and honest; a phase "done" without its manual checklist run on a device is not done (DEVELOPMENT_WORKFLOW rule 4).
4. **Debt is registered at merge time** with its three fields and a named removal point — or the shortcut doesn't merge.
5. **Checkpoints block:** findings are fixed or formally amended before the next milestone; skipping a checkpoint to "keep momentum" is the failure mode this roadmap exists to prevent.
6. **Scope changes are amendments:** adding, reordering, or fattening phases happens by amending this document first (CODING_STANDARDS rule 8), never silently in a PR.
7. **Working software after every merge** — if a merged phase breaks a prior flow, fixing that outranks starting the next phase.
8. **The Closure Ritual is not optional.** No tag, no journal entry, no screenshots, no real-usage session, no retrospective → the phase is not complete, regardless of green CI. Declaring completion without the ritual is a false status report (DEVELOPMENT_WORKFLOW rule 4).

---

## Changelog

- 2026-07-08 — v1 baseline frozen (Phase Closure Ritual added per approved refinements; F3: Phase 16 gains the §3.8 streak/consistency domain functions; F4: Phase 19 extends the backup format for phases, Phase 18's dependency on 17 hardened, Phase 1 just-in-time primitive note added).
