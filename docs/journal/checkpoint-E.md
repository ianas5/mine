# Checkpoint E — The v1 Gate (Phase 22)

**Closed:** 2026-07-09 · **Scope:** the entire codebase and product, reviewed for release.
**Type:** hardening + final review — no new features (freeze is absolute).

**Verdict: RELEASE CANDIDATE.** Every gate that can be certified in this headless environment
passes; the app is code-complete and, by the standard *"would it still feel great after six
months of daily use?"*, it is built to. Two review findings and the on-device gates remain
before `v1.0.0` is tagged — detailed below. Tagged `v1.0.0-rc1`.

The question this checkpoint asked, from a first-time user's seat: *is anything confusing,
inconsistent, slow, noisy, over-built, hard to discover, or awkward?* Answered below.

## The seven gates

| # | Gate | Status |
|---|---|---|
| 1 | **Regression** — every §5 flow | ✅ Consolidated into `docs/V1_TEST_PLAN.md`; automated floor green (375 tests). On-device pass is the owner's (TD-001). |
| 2 | **Performance** — §7 budgets, 5-year data | ✅ Watchdog (`perf.test.ts`): full dashboard analytics **~37 ms** and chart bucketing **~2 ms** over a 1,044-workout / 20,880-set / 1,825-day synthetic set — within the 50 ms / 16 ms budgets with margin. Device timing-log is the authoritative check. |
| 3 | **Backup / data safety** | ✅ Round-trip + every refusal path unit-covered (`backupService.test.ts`), incl. phases. On-device real-dataset drill + safety-export failure = owner. |
| 4 | **Migrations** | ✅ `db:check` applies 0000→current on a fresh DB (16 tables) every CI run; additive-only. Backup-then-migrate device drill = owner. |
| 5 | **Accessibility** | ⚠️ Reduce Motion ✅ (Phase 21); contrast audited → **F-E1** (`textTertiary` fails AA). VoiceOver + Dynamic Type = owner device pass. |
| 6 | **Polish & both-themes** | ✅ Code-level P16 sweep clean; every state (empty/insufficient/error/loading) present. Both-themes eyeball = owner. **F-E1** is the one contrast fix. |
| 7 | **CP-E review** | ✅ Below. |

## CP-E — full-codebase & product review

**Architecture invariants (re-verified across the whole tree, CI-enforced):**

- **Layered boundaries** hold — `eslint-plugin-boundaries` at `--max-warnings 0`; features never import each other; shared cross-feature state lives in `data` (the insight + phase hooks).
- **Calculators are pure** — `grep` finds no `react`/`expo`/`@/data`/`@/core/db`/`@/core/storage` import and no wall-clock/randomness anywhere in `domain/`.
- **SQL retrieves only** — no `CASE`/arithmetic in any repository select; all classification/e1RM/volume/bodyweight/ranking logic is in `domain/` (Phase 20 moved the last stray heuristic — the nutrition picker — out of the repository).
- **Honesty holds** — the interpretation triplet is structurally required (`StatTile.context`, `ChartFrame.interpretation`); `MetricResult` insufficient-data propagates to every surface; every insight is evidence-backed with an exhaustive tap-through; MMKV confined to `core/storage`.

**The Tests (PRODUCT_PRINCIPLES) against v1 as shipped:** no vanity metrics (the sole streak is weekly and never punishes rest; no score/points/levels); no dark patterns (no badges/dots, no guilt copy, no notification hooks); tap budgets intact (set ≤ 2, meal ≤ 3, weight ≤ 3); every visible metric answers one of the eight standing questions (mapped at CP-D, unchanged); delight is the "whole party" of a haptic + one ≤ 800 ms visual, nothing more.

**First-time-user read:**

- *Confusing / inconsistent?* One inconsistency found — the Analytics home rendered its sections Body → Training → Nutrition with the phase card on top, deviating from UI_UX §8's fixed order. **Fixed** (F-E2): the home now reads Insights → range → **Training → Nutrition → Body → Phases**, matching §8 and the dashboard's workout-first hierarchy. Learn-one-report-know-them-all holds; headers, range control, and quick-action order are consistent everywhere.
- *Slower than expected?* No — the perf watchdog shows the heaviest path (5 years of data) computing in tens of milliseconds; memoization by data-version means unchanged data recomputes nothing.
- *Visually noisy?* No — the P16 sweep found shadows only on Card's raised variant and the only perpetual animation the (gated) loading Skeleton.
- *Any screen doing too much?* No — the dashboard is a closed briefing; analytics is four fixed sections; each report follows one shape.
- *Hard to discover?* No hidden-gesture-only functionality — every action is reachable by a visible tap; the pending gesture polish (TD-003) is additive, not a discovery gap.
- *Awkward workflow?* None surfaced; the core loops are ≤ 2–3 taps and pre-filled from history.

**Weekly-use / v1-scope audit** (*"impressive but not used weekly → not in v1"*): every feature was walked against "would I open this in a normal training week?" — Dashboard, workout logging, meal logging, weigh-ins/measurements (daily/weekly); programs→start-session, analytics review, insights (weekly); progress photos, muscle report, phase progress (periodic but low-footprint — a single glanceable card daily). **Nothing was found that is impressive-but-idle and should be cut.** The closest to non-weekly is Phases (declared monthly), but its daily surface is one card answering "is this block working," so it earns its place. This is the healthy outcome of purpose-gating every feature at build time — there is no cleverness to remove.

## Findings

- **F-E1 — `textTertiary` fails WCAG AA (accessibility, must-fix before v1).** Contrast on
  surface is **2.84** (dark) / **2.69** (light) — below AA (4.5) and even AA-large (3.0); used
  for small functional labels (unit suffixes, "TO GOAL"/"BODY" micro-labels, "(derived)").
  Every other token passes. *Recommended fix (DESIGN_SYSTEM §2.2 token amendment — needs your
  ratification + an on-device hierarchy eyeball):* dark `#5C6066 → #7E838A` (4.71), light
  `#9A9EA5 → #6E7278` (4.84). I did **not** change the frozen token blind. **This is the one
  code change I recommend before tagging v1.0.0.**
- **F-E2 — Analytics section order deviated from §8. FIXED this phase** (reordered to Training →
  Nutrition → Body → Phases). No amendment needed — the code now matches the frozen spec.

## Technical-debt final audit

- **Resolved:** TD-002, TD-004, TD-006, TD-009.
- **v1 device-pass (owner, gate the tag):** **TD-001** (the physical-iPhone regression — it has
  accumulated every phase's walk, now consolidated into `V1_TEST_PLAN.md`), and the three
  device-feel items re-approved from Phase 21 — **TD-003** (gesture drag/swipe), **TD-008**
  (keyboard next-field), **TD-010** (Skia chart area-fill/regression/tooltip). Each can only be
  *tuned* on hardware; their function is already covered (visible-tap deletes + Undo, tappable
  fields, the text interpretation contract), so nothing is blocked.
- **Deferred to the post-v1 backlog (additive, not defects; the Phase 22 feature freeze forbids
  building them now):** **TD-005** (FlashList — only when a full-screen all-history archive
  ships), **TD-007** (the Exercise Report's selected-range body + recent notes — the report is
  honest as-is with all-time bests).
- No overdue *defect* blocks the gate. Every open item is either owner-device work or an
  intentional, documented additive deferral.

## Release recommendation

**Tag `v1.0.0` once the owner completes, on the personal iPhone:**

1. The `V1_TEST_PLAN.md` regression pass (all seven gates green, both themes) — including
   VoiceOver on the three critical flows, Dynamic Type 1.3×, the real-dataset backup + safety
   drill, and the backup-then-migrate drill.
2. **F-E1** — apply the `textTertiary` contrast fix and confirm the hierarchy still reads well.
3. The three device-feel debts (TD-003/008/010) wired and tuned, or explicitly carried to a
   v1.0.x with rationale.

Everything else is done: the app builds, lints, typechecks, and passes 375 tests; the analytics
engine computes 5 years of data well within budget; migrations are clean and additive; backup
round-trips losslessly; the honesty and layering invariants hold across the whole codebase; and
the product is calm, consistent, discoverable, and free of anything impressive-but-idle. By the
six-month-of-daily-use standard, this is a v1 worth shipping.

## Lessons Learned

- **What went well?** The gate was mostly *confirmation*, not firefighting — because honesty,
  layering, and purpose were enforced continuously (frozen docs, boundary lint, per-phase
  checkpoints), CP-E found two small issues, not a backlog. The perf watchdog turned "is it fast
  enough?" from a hope into a number.
- **What was harder than expected?** Drawing the honest line between "fix now" and "owner must
  see it." The §8 order gap was unambiguous against a frozen spec → fixed. The contrast gap and
  the device-feel debts trade against things only visible on hardware → surfaced with concrete
  recommendations rather than changed blind. Resisting the urge to "polish" approved screens I
  couldn't see was the discipline.
- **What should be reused:** a synthetic-dataset perf watchdog as a permanent regression guard;
  consolidating the accumulated TD-001 walks into a single ordered test plan rather than
  re-deriving; auditing every feature against "would I use this weekly?" before a release.
- **What should be avoided:** tagging a release while device-only gates are unverified (hence a
  release *candidate*); changing frozen design tokens or building device-feel work blind.
- **Amendment proposals:** one — the F-E1 `textTertiary` contrast values (DESIGN_SYSTEM §2.2),
  for your ratification. No other frozen-document change.
