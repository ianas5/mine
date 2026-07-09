# Phase 22 — Hardening (the v1 gate)

**Closed:** 2026-07-09 · **Milestone:** M6 — the v1 gate. **Feature freeze: absolute.**

The governing frame: **prepare a product for release, not finish development** — verify not
just correctness but *confidence*, from a first-time user's seat, against the standard *"if I
used this every day, would it still feel great after six months?"* No new features; only
hardening, verification, and the fixes that genuinely earn their place before v1.

The full architecture + product review and the release recommendation are in
**`docs/journal/checkpoint-E.md`** (CP-E). This entry records the hardening *work*.

## What was built / done

- **Performance watchdog** (`src/domain/analytics/perf.test.ts`) — generates a synthetic
  **5-year** dataset (1,044 workouts / 20,880 sets / 261 snapshots / 1,825 nutrition days) and
  times the full dashboard-analytics path and a chart series against the §7 budgets. Measured
  **~37 ms** (budget 50) and **~2 ms** bucketing (budget 16) — within budget with margin, and
  a permanent regression guard.
- **`docs/V1_TEST_PLAN.md`** — the Phase 22 regression consolidation: every §5 flow and every
  accumulated TD-001 walk, reorganized by area, plus the seven v1 gates, as one ordered
  on-device checklist in both themes.
- **Contrast audit** (WCAG, both themes) — all tokens pass except `textTertiary` (F-E1); logged
  as **TD-011** with the proposed fix, for owner ratification (a frozen-token amendment, not
  changed blind).
- **F-E2 fixed** — the Analytics home now follows UI_UX §8's fixed section order (Insights →
  range → **Training → Nutrition → Body → Phases**), correcting a drift where Body led and the
  phase card sat on top. Aligns to the frozen spec and the dashboard's workout-first hierarchy.
- **CP-E full-codebase review** — architecture invariants (pure calculators, SQL-retrieval-only,
  layered boundaries, honesty triplet, evidence-backed insights) re-verified across the tree and
  confirmed CI-enforced; The Tests re-run against v1; a first-time-user read; and a weekly-use
  scope audit (found nothing impressive-but-idle to cut).

## What changed

New: `domain/analytics/perf.test`; `docs/V1_TEST_PLAN.md`; `docs/journal/checkpoint-E.md`.
Modified: `AnalyticsScreen` (§8 section order); `TECH_DEBT` (TD-011 added; TD-001 delight walk;
TD-003/008/010 re-approved to Phase 22 in Phase 21). No schema, no new deps, no frozen document
changed (F-E1 is a *proposed* amendment awaiting ratification).

## Screens affected

Analytics home (section reorder to §8). No other screen changed.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| 5-year dashboard analytics ≤ budget; chart bucketing ≤ budget | perf watchdog | ✅ ~37 ms / ~2 ms |
| Full-codebase invariants (purity, SQL retrieval, boundaries, MMKV, triplet, insight evidence) | grep + lint | ✅ clean |
| The Tests (no vanity metrics, no dark patterns, tap budgets, standing-question coverage) | review | ✅ |
| Contrast (all tokens, both themes) | WCAG calc | ⚠️ `textTertiary` fails → TD-011 / F-E1 |
| Analytics §8 section order | fixed + inspection | ✅ |
| `npm run check` | typecheck + lint + format + 375 tests + db:check (16 tables) | ✅ green (3× stable) |
| On-device v1 gate pass (regression both themes, VoiceOver, Dynamic Type, real-dataset backup/migrate drills, device-feel debts) | — | ⚠️ owner (V1_TEST_PLAN.md / TD-001) |

## Known limitations

1. **The device gates are the owner's** — on-device regression (both themes), VoiceOver on the
   three critical flows, Dynamic Type 1.3×, the real-dataset backup + safety drill, the
   backup-then-migrate drill, and the three device-feel debts (TD-003/008/010). This is the one
   phase whose certification is fundamentally on-hardware; the off-device floor is fully green.
2. **F-E1 (`textTertiary` contrast)** is the one recommended code change before `v1.0.0`, held as
   a token amendment (TD-011) for ratification rather than changed unseen.
3. **TD-005 / TD-007** are carried to the post-v1 backlog — additive, not defects, and barred by
   the Phase 22 feature freeze.

## Technical debt

None introduced. **TD-011** recorded (F-E1 contrast, pre-v1, owner-ratified). Registry audited:
resolved TD-002/004/006/009; owner-device TD-001/003/008/010; post-v1 additive TD-005/007. No
overdue defect.

## Retrospective

**What went well?** CP-E was confirmation, not firefighting — continuous enforcement (frozen
docs, boundary lint, per-phase checkpoints) meant the final gate surfaced two small issues, not
a backlog. Turning the perf budget into a measured number, and consolidating five phases of
device walks into one ordered plan, made "is it ready?" answerable rather than felt.

**What was harder than expected?** The judgment line: fix the §8 order (unambiguous against a
frozen spec) but *surface* the contrast and device-feel items (only verifiable on hardware)
rather than change them blind. Restraint — not polishing screens I can't see — was the hardest
and most correct discipline of the phase.

**What should change before the next phase?** There is no next phase — this is the gate. The
remaining work is the owner's on-device pass; on its completion (plus F-E1), `v1.0.0` is
warranted.

## Lessons Learned

- **What surprised you:** how little the gate found — the payoff of honesty-and-purpose enforced
  from Phase 0, not bolted on at the end.
- **What documentation prevented mistakes:** ANALYTICS §7 gave the perf budgets to measure
  against; UI_UX §8 caught the analytics-order drift; DESIGN_SYSTEM §8/§2.2 framed the contrast
  gate and the token amendment; the roadmap's seven-gate structure and absolute feature freeze
  kept the phase to hardening.
- **What should be reused:** a synthetic-dataset perf watchdog as a standing guard; a single
  consolidated `V1_TEST_PLAN`; the weekly-use scope lens before any release.
- **What should be avoided:** tagging a release with device gates unverified (hence a candidate);
  changing frozen tokens or device-feel code blind; adding "one more feature" at the gate.
- **Amendment proposals:** F-E1 `textTertiary` contrast (DESIGN_SYSTEM §2.2), for ratification.
