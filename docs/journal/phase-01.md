# Phase 1 — Design System Core

**Closed:** 2026-07-08 · **Tag:** `v0.2.0-phase1` · **Milestone:** M0 Foundation

## What was built

The complete Phase 1 primitive batch in `src/core/ui` (15 components), the
remaining theme machinery, and the dev-only gallery:

- **Theme layer:** haptic tokens (`light`/`success`/`warning` → expo-haptics,
  web no-op) and `useThemedStyles` (per-theme style caching), both exported via
  the curated `core/theme` barrel.
- **Structure:** `Screen` (safe-area + gutter + scroll), `Card`
  (default/raised/accentEdge), `Section` (the only section-header pattern),
  `ListRow` (52pt, leading/trailing/chevron), `EmptyState` (factual line + CTA),
  `Skeleton` (pulse, static under Reduce Motion).
- **Actions & input:** `Button` (4 variants × 2 sizes, pressed scale 0.98,
  width-locked loading), `IconButton` (44pt, required a11y label), `Input`
  (label/error/focus states, numeric mode with unit suffix + select-on-focus),
  `Stepper` (44pt targets, long-press auto-repeat 350/120ms, haptic tick,
  float-drift guard), `Chip` (accentSoft selected + haptic), `SegmentedControl`
  (sliding thumb at motion.fast, tab roles).
- **Overlays & feedback:** `Toast` (`showToast()` + `ToastHost` at root, one at
  a time, 2.5s, success tone = success haptic), `Dialog` (destructive-only,
  warning haptic on confirm), `Sheet` (grabber, keyboard-safe, dirty-state
  discard guard via Dialog).
- **Gallery:** `/gallery` route (dev builds only, production redirects home)
  rendering every primitive in every state, split into structure/actions/
  overlays sections.
- **Tests:** 15 RNTL test files — 46 component tests covering states, events,
  bounds, and accessibility labels/roles — plus the existing token tests
  (50 tests total).

## What changed

New: `src/core/ui/` (15 components + barrel + 15 test files),
`src/core/theme/haptics.ts`, `src/core/theme/useThemedStyles.ts`,
`src/features/gallery/` (screen + 3 section components), `app/gallery.tsx`,
`jest.setup.js`. Modified: `core/theme` barrel, `app/_layout.tsx` (ToastHost),
`eslint.config.js` (jest globals block), `package.json` (expo-haptics, RNTL,
jest setup). No frozen document was modified.

## Screens affected

The gallery (new). Tab screens unchanged. Screenshots:
`screenshots/phase-01/` — `gallery-dark.png` / `gallery-light.png` (full page)
+ `sheet-open-dark.png` / `dialog-open-dark.png` (interactive states).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Gallery renders every primitive, both themes | Dev-server (Metro, `__DEV__` true) + Chromium at iPhone viewport, dark & light | ✅ all sections render with correct tokens |
| Interactive states | Real clicks: sheet opened (grabber, unit input, Save), dialog opened, chips switched, segmented control moved, toasts fired | ✅ captured in screenshots |
| Dirty-state guard | Typed in sheet input → dismissed → "Discard entry?" dialog appeared; cancel kept sheet open | ✅ (also covered by Sheet tests) |
| Stepper auto-repeat + haptic | Unit-tested (timing + haptic call); real long-press feel needs device | ✅ / ⚠️ device feel pending (TD-001 scope) |
| `npm run check` | typecheck + lint (0 warnings) + format + 50 tests | ✅ green |
| Dynamic Type 1.3× / press latency on device | — | ⚠️ requires physical device (TD-001) |

## Known limitations

1. **Device-dependent acceptance items** (press-state latency ≤ 1 frame, haptic
   feel, Dynamic Type 1.3× walk) remain pending with TD-001 — same root cause,
   same CP-A resolution.
2. **Sheet swipe-dismiss is not a drag gesture yet** — dismissal works via
   backdrop tap, Android back, and programmatic close; grabber drag-to-dismiss
   needs gesture-handler wiring (registered as TD-003).
3. Gallery screenshots are web renders; native modals (Dialog/Sheet) may differ
   slightly in animation on device.

## Technical debt introduced

- **TD-003** — Sheet lacks gesture-based swipe-dismiss (see `TECH_DEBT.md`).

## Retrospective

**What went well?** The DESIGN_SYSTEM §6 catalog was precise enough that all 15
components were transcription; tokens-only styling survived its first real
consumer (the gallery, a `features/` module, passed the token lint untouched).

**What was harder than expected?** The test stack: RNTL v14 switched to a fully
async API (`await render`, `await fireEvent`) which no prior knowledge covered —
the pilot-test-first approach caught it before 15 files were written wrongly.
The new `react-hooks/refs` rule rejected the classic
`useRef(new Animated.Value()).current` RN pattern in three components
(fixed with lazy `useState`), and the safe-area jest mock needed its `.default`.

**What should change before the next phase?** (1) Keep the pilot-test-first
habit: validate one file through any new testing surface before batch-writing.
(2) The dev-server capture path (needed for `__DEV__`-gated screens) is slower
than static export; keep static export for product screens, dev server only for
gallery.
