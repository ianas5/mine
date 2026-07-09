# Phase 21 — Delight & Feel Pass

**Closed:** 2026-07-09 · **Milestone:** M5 (Refinement) — second phase

The governing principle this phase: **the user should stop noticing the interface.** Every
delight must be *earned* — nothing animates simply because it can; each animation, haptic, or
transition has to improve understanding or reinforce confidence, or it doesn't ship. The
fastest interface is the one that feels calm. So this pass was as much about **restraint** as
addition: complete the six-moment registry, gate everything behind Reduce Motion, keep the
logging path frame-free (rule 11), and sweep out visual noise.

> **Environment note.** Animation *feel* — timing, gesture physics, Skia rendering — is
> inherently on-device and cannot be validated in this headless Linux build. This phase
> therefore implements each moment faithfully and tests the pure logic (the count-up curve,
> the "best" detection, the Reduce-Motion gating), while the subjective feel folds into TD-001.
> The genuinely device-only *tuning* debts (gesture-handler, keyboard focus, Skia chart) are
> re-approved to the Phase 22 on-device pass rather than shipped blind (see below).

## Earned delight, or none (the standing principle)

The delight registry (UI_UX §5.4) is a **closed list of exactly six moments** — one haptic +
one visual response, ≤ 800 ms, never blocking, never confetti, never currency. The work was to
make each real, add nothing beyond them, and ensure motion is never the sole carrier of
feedback.

## What was built

**Foundation (reusable, reduce-motion-first):**

- **`useReducedMotion`** (`core/theme`) — one hook every animation reads, subscribed to live
  OS changes; `Skeleton` was refactored onto it (dedup). Motion is opt-out at the source.
- **`SettleIn`** (`core/ui`) — a one-shot scale-settle + fade for a freshly-appearing element,
  RN `Animated` spring, ≤ 800 ms, non-blocking. Under Reduce Motion it appears instantly at
  full scale/opacity — feedback without motion (P15).
- **`useCountUp` + `countUpValue`** (`core/ui`) — an easeOutCubic count-up driven by
  `requestAnimationFrame`; state is set only from the async frame callback (never
  synchronously in an effect), and the number is derived purely. Under Reduce Motion or when
  disabled it returns the target immediately — the value is always correct, only the approach
  animates. The curve is unit-tested.

**The six registry moments:**

1. **New PR** — the `PrBadge` now materializes via `SettleIn` (soft scale-settle) alongside the
   existing `success` haptic + finish toast — "the whole party," unchanged in substance.
2. **Workout completed** — the summary's working-sets and volume **count up** briefly and land
   on the exact totals; the "N new PRs" banner settles in.
3. **Phase completed** — arriving straight from completing a phase (`?celebrate=1`) lands the
   report's identity header with a quiet settle — *the report is the reward* (§5.4), replacing
   nothing in the Phase 19 flow.
4. **Measurement best** — a new pure `bestFieldValues` / `isFieldBest` (`domain/body`) detects a
   value that beats the prior best in its improving direction (§5.3); on a best, a gentle
   `light` haptic + a distinct "New best — Waist 81 cm" toast fire, while a routine save just
   confirms. (This also corrected an off-§6 `success` haptic that fired on every measurement
   save — routine logging isn't a §6 haptic point.)
5. **First workout after ≥ 14 days** — verified already live (Dashboard greeting, Phase 16):
   warm, factual, no guilt.
6. **Streak milestone (4/8/12/26/52 wk)** — verified already live (insight rule #14): card
   styling only, information not fanfare.

**Reduce Motion & the noise sweep:** every new animation swaps to instant/cross-fade under
Reduce Motion. The P16 sweep audited the UI for gratuitous decoration and found it already
restrained — shadows appear only on `Card`'s raised variant, and the *only* perpetual
animation in the app is the (justified, reduce-motion-gated) loading `Skeleton`. Nothing
animates without a reason; no changes were needed.

## What changed

New: `core/theme/useReducedMotion`; `core/ui/{SettleIn,countUp}` (+ countUp test);
`domain/body` `bestFieldValues`/`isFieldBest` (+ test). Modified: `Skeleton` (shared hook);
`PrBadge`, `WorkoutSummarySheet` (settle + count-up); `AddMeasurementsSheet` + `MeasurementsScreen`
(best-yet detection, gentler haptic); `PhaseReportView`/`PhaseReportScreen`/`PhasesScreen`
(celebrate settle). No schema, no new deps (RN `Animated` + `expo-haptics`, both already in
use), no frozen document changed.

## Screens affected

Active Workout (PR badge settle), Finish summary (count-up + banner settle), Add Measurements
(best-yet moment), Phase Report (completion settle). No layout or tap-count change on any.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Count-up curve: pinned endpoints, clamped, eases past linear midpoint, monotonic | domain/ui test | ✅ |
| Measurement best: lowest/highest per direction; neutral excluded; strict beat; no prior → not a best | domain test | ✅ |
| Reduce-Motion gating returns final value/scale instantly (logic path) | inspection + test | ✅ |
| `npm run check` | typecheck + lint + format + 372 tests + db:check (16 tables) | ✅ green (3× stable) |
| On-device delight walk (each of the six moments subtle/≤800 ms/non-blocking, Reduce-Motion cross-fades preserve feedback, no logging flow gained a tap or frame, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **Animation feel is device-only** — that each moment *feels* right (timing, spring, the
   count-up's pace) and that Reduce Motion reads correctly folds into TD-001. What is verified
   off-device is the logic: the easing curve, the best-detection, the reduce-motion branches.
2. **Three feel debts are device-gated and re-approved to Phase 22** (from Phase 21):
   **TD-003** (gesture-handler drag-dismiss / swipe-to-reveal), **TD-008** (keyboard next-field
   focus chaining), **TD-010** (Skia chart area fill / regression overlay / tooltip). Each
   needs a device to *tune*; wiring or styling them blind would ship unverifiable motion or
   Skia output — contrary to their own deferral rationale and rule 11. Their function is
   already covered by visible taps (delete + Undo), tappable fields, and the text
   interpretation contract, so nothing is blocked in the meantime.

## Technical debt

None introduced. TD-001 gains the delight & feel walk. TD-003/008/010 are re-approved from
Phase 21 to **Phase 22** with rationale (device-gated feel tuning). No registry entry became
overdue.

## Retrospective

**What went well?** Reading "nothing should animate simply because it can" as the brief kept the
pass small and honest: the registry is exactly six moments, each a single haptic + a ≤ 800 ms
visual, all reduce-motion-gated and non-blocking. Building the foundation first
(`useReducedMotion`, `SettleIn`, `useCountUp`) meant the moments are three-line wirings over
tested primitives rather than bespoke animation each. The P16 sweep's best outcome was finding
nothing to remove — the app was built calm.

**What was harder than expected?** Two things. First, the lint rules that forbid synchronous
`setState` in effects and `ref.current` in render forced the count-up to be modelled as a
purely-derived value over an rAF-driven progress, and `SettleIn` to hold its `Animated.Value`s
in state initializers — both of which are the *correct* shapes and made the code cleaner.
Second, the honest scoping: this is the one phase whose deliverable is *feel*, which a headless
environment can't judge. The disciplined answer was to implement + test the logic faithfully,
fold the feel into TD-001, and re-approve the native-gesture/Skia tuning to the device pass
rather than gold-plate blind.

**What should change before the next phase?** Nothing structural. Phase 22 is the v1 gate —
the on-device hardening pass that will finally *feel* everything built here, verify the three
re-approved debts on a real iPhone, and run the full regression/perf/accessibility gates.

## Lessons Learned

- **What surprised you:** how much the reduce-motion-first foundation simplified the moments —
  once every animation reads one hook and degrades to instant, adding a moment is trivial and
  safe.
- **What documentation prevented mistakes:** UI_UX §5.4 fixed the registry to exactly six
  moments (so "no more, no less" was checkable); §9 fixed that Reduce Motion swaps, never
  removes, feedback; rule 11 fixed that no delight may enter a logging path's critical timing;
  §6 fixed the haptic vocabulary (which surfaced the off-spec measurement-save `success`).
- **What should be reused:** the `useReducedMotion` + `SettleIn` + `useCountUp` trio as the
  vocabulary for any future moment; deriving animated display values purely from an
  rAF/`Animated` source rather than mirroring them into synchronous state.
- **What should be avoided:** animating for its own sake; letting motion be the sole feedback
  (always pair with a haptic/toast/final value); building device-feel or Skia work blind when
  it can only be tuned on hardware — defer and re-approve honestly instead.
- **Amendment proposals:** none — no frozen-document defect surfaced. No new debt; three debts
  re-approved to Phase 22 with rationale.
