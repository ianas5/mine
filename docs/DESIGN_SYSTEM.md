# DESIGN_SYSTEM.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Owner concern:** visual tokens and reusable primitive components (`core/theme`, `core/ui`) · **Depends on:** PRODUCT_PRINCIPLES, PROJECT_VISION, ARCHITECTURE. Screen composition, flows, and navigation behavior belong to UI_UX_GUIDELINES.
>
> This document implements PRODUCT_PRINCIPLES P15–P18 in visual and haptic form; on any conflict, the principles win. Every visual decision in the app resolves to a token or primitive defined here. Feature code never hardcodes a color, size, font, or radius.

---

## 1. Design Personality (P16, P17, P18)

**Premium, calm, data-confident.** The app should feel like a beautifully made instrument: dark, quiet surfaces; one energetic accent used sparingly; numbers that look engineered; generous whitespace; zero decoration that doesn't inform. The coach personality (ANALYTICS_ENGINE §6.1) extends visually: insights look like considered statements, not alert spam.

Anti-goals: clutter, gamified confetti, emoji as interface, rainbow dashboards, dense grids of tiny stats.

---

## 2. Color System

### 2.1 Principles

- **Dark theme is the primary, default theme** (gym environments, OLED, premium feel). Light theme is fully supported and token-complete from day one.
- Features use **semantic tokens only** — never raw hex, never palette names.
- The accent is **earned**: it marks primary actions, active states, and the user's progress — not decoration. Most of any screen is neutral.

### 2.2 Semantic tokens (dark ▸ light)

| Token | Dark | Light | Use |
|---|---|---|---|
| `bg` | `#0B0C0E` | `#F7F7F8` | app background |
| `surface` | `#151719` | `#FFFFFF` | cards, sheets |
| `surfaceRaised` | `#1D2023` | `#FFFFFF` (+shadow) | elevated cards, modals |
| `border` | `#26292E` | `#E4E5E9` | hairlines, card edges |
| `textPrimary` | `#F2F3F5` | `#17181A` | headings, values |
| `textSecondary` | `#9BA0A8` | `#5D6167` | labels, captions |
| `textTertiary` | `#7E838A` | `#6E7278` | placeholders, disabled |
| `accent` | `#FF6A3D` | `#E85320` | primary actions, active tab, focus |
| `accentSoft` | `#FF6A3D` @ 12% | `#E85320` @ 10% | selected chips, highlights |
| `positive` | `#3ECF8E` | `#1F9D66` | improving, hits, PRs |
| `attention` | `#FFB020` | `#B97A00` | attention-tone insights, off-track |
| `danger` | `#FF5C5C` | `#D64545` | destructive actions only |
| `chartLine` | `accent` | `accent` | primary series |
| `chartMuted` | `#3A3E44` | `#D8DADF` | axes, reference lines |

- **Tone mapping (fixed, P17):** insight `positive` → `positive`, `attention` → `attention`, `neutral` → neutral surfaces + `textSecondary`. `danger` is **never** used for insights — being off-track is attention, not an emergency.
- **Directionality coloring** obeys FITNESS_DOMAIN §5.3: *improving* is `positive` regardless of numeric sign (a shrinking waist is green).
- **Categorical chart palette** (muscle groups etc., color-blind-checked on both themes): `#FF6A3D, #3ECF8E, #4C9AFF, #C77DFF, #FFB020, #2EC4B6, #FF7DA0, #A3E635` — assigned by stable order, never randomly.

### 2.3 Contrast rules

Text ≥ 4.5:1 against its surface (≥ 3:1 for ≥ 20 pt bold); `textTertiary` never carries information alone; accent-on-surface always paired with a text label (color is never the only signal).

---

## 3. Typography

- **Typeface: Inter** (via `expo-font`), with **tabular numerals enabled for all data values** — columns of weights and macros must align like an instrument panel. Fallback: system (SF Pro / Roboto).
- Scale (pt, line-heights tuned per role):

| Token | Size / weight | Use |
|---|---|---|
| `display` | 34 / 700 | hero numbers (today's calories, current weight) |
| `title` | 24 / 700 | screen titles |
| `heading` | 18 / 600 | card titles, section headers |
| `body` | 15 / 400 | default text, insight bodies |
| `bodyStrong` | 15 / 600 | emphasized body, values in rows |
| `caption` | 13 / 500 | labels, axis text, metadata |
| `micro` | 11 / 600 uppercase +0.6 tracking | overlines, unit labels |

- **Numbers rule:** every metric value renders `bodyStrong`-or-larger, tabular, with its unit in `micro`/`caption` at `textSecondary` ("82.4 **kg**" — value loud, unit quiet).
- No font sizes outside the scale. Dynamic Type: respects OS font scaling up to 1.3× without layout breakage (test gate in UI_UX).

---

## 4. Space, Shape, Elevation, Motion

- **Spacing:** 4-pt base grid — tokens `xs 4 · sm 8 · md 12 · lg 16 · xl 20 · 2xl 24 · 3xl 32`. Screen gutter = `lg (16)`. Card padding = `lg`. Section gap = `2xl`.
- **Radius:** `sm 8` (chips, inputs) · `md 12` (buttons) · `lg 16` (cards, sheets) · `full` (pills, rings). One family, no mixing per component.
- **Elevation:** dark theme elevates by **surface color + hairline border**, not heavy shadows; light theme uses soft shadows (`y2 blur8 @ 8%`, `y8 blur24 @ 10%` for modals). Never both borders-and-heavy-shadow.
- **Motion tokens:** `fast 150ms` (presses, toggles) · `base 250ms` (sheets, transitions) · `slow 400ms` (progress rings, chart draw-in); easing `standard cubic-bezier(0.2, 0, 0, 1)`. Motion is functional only; **Reduce Motion** disables non-essential animation.
- **Haptics tokens:** `light` (set logged, chip select) · `success` (workout saved, PR — the *only* celebration channel besides a toast) · `warning` (destructive confirm). Mapped to `expo-haptics`; used via tokens so intensity policy is central.

### 4.1 Feel & Feedback (P15, P18)

- Every pressable has a **≤ 1-frame pressed state**.
- **Optimistic value updates** are the default for local writes.
- **Skeletons** (never blank screens, never full-screen spinners) for first paints.
- Long operations (import/export) show **determinate progress** in a sheet.
- The three "thousand-times" interactions — **stepper tick, set-complete, sheet dismiss** — get first-class motion + haptic tuning and are the benchmark for app feel.

---

## 5. Iconography & Imagery

- **One icon set: Lucide** (`lucide-react-native`) — consistent 1.75 stroke, sizes `16 / 20 / 24`, colored via text tokens.
- **No emoji anywhere in UI chrome** (tabs, buttons, cards, empty states). Emoji may appear only inside user-typed notes.
- Progress photos render full-bleed within `lg`-radius frames; no filters, no overlays except date/angle caption chips.

---

## 6. Primitive Component Catalog (`core/ui`)

Each primitive: purpose, variants, and non-negotiable behavior. Props contracts are finalized in code; this is the design contract. All primitives are theme-aware, Dynamic-Type-safe, and expose `accessibilityLabel`s.

**Structure**
- **`Screen`** — safe-area wrapper, `bg`, standard gutter, optional scroll; every screen uses it.
- **`Card`** — `surface`, `lg` radius, `lg` padding; variants: `default`, `raised`, `accentEdge` (3-pt accent left edge, used by InsightCard only).
- **`Section`** — heading + optional "See all" action + content gap; the only way section headers are built.
- **`ListRow`** — leading icon/text, primary/secondary text, trailing value/chevron; 52-pt min height; hairline separators inset to text.
- **`Divider`**, **`EmptyState`** (icon + one-line reason + optional CTA — wording rules in UI_UX), **`Skeleton`** (pulse placeholder for cold-start paints).

**Actions & input**
- **`Button`** — variants `primary` (accent fill), `secondary` (surface + border), `ghost` (text only), `destructive` (danger fill, confirm-gated); sizes `lg 52pt / md 44pt`; states default/pressed (scale 0.98 + `fast`)/disabled (40% fg); loading = inline spinner replacing the label, width locked.
- **`IconButton`** — 44-pt target minimum regardless of glyph size.
- **`Input`** — RHF-compatible; label (`caption`), value (`body`), error (danger text + border), focus (accent border); numeric mode uses tabular numerals + right-aligned unit suffix slot.
- **`Stepper`** — the workhorse of workout logging: value with `+ / −` targets (44 pt), configurable increment (**2.5 kg / 1 rep defaults**), long-press auto-repeat, haptic `light` per tick.
- **`Chip`** — selectable pill; selected = `accentSoft` bg + accent text; used for slots, angles, filters.
- **`SegmentedControl`** — the **single** time-range switcher (7D · 30D · 3M · 6M · 1Y · All) and any either/or toggle; sliding thumb, `fast`.

**Data display**
- **`StatTile`** — the interpretation-triplet primitive (P6, P8): value (tabular, `display`/`title`), unit, label (`micro`), and a **required context line** (reference + classification, e.g. "▲ 4% vs last week" in `positive`). *A StatTile without context is a build error by convention* — this enforces ANALYTICS_ENGINE rule 3 at the component level.
- **`ProgressRing`** — daily kcal/protein; value centered, ring in accent → `positive` at 100%; over-target renders a thin `attention` overflow arc, never a second lap.
- **`ProgressBar`** — macros, weekly consistency; track `chartMuted`, fill by tone.
- **`InsightCard`** — tone-tinted edge + icon, `heading` title, `body` text, optional tap-through to evidence chart; never stacked more than 3 (ANALYTICS_ENGINE §6.3).
- **`PRBadge`** — small `positive` pill ("PR"), used inline in set rows and history.
- **`TrendArrow`** — ▲ ▼ → glyphs colored by *classification* (not raw sign), always beside a number, never alone.
- **`ChartFrame`** — mandatory chart wrapper (P6, P8): title, **interpretation line** (the coach sentence for the chart), the plot, and range control slot. Charts never render outside a ChartFrame — a plot without interpretation violates the analytics contract.
- **`Sparkline`** — inline 60×20 trend glyph for rows (exercise/muscle reports); no axes.

**Overlays & feedback**
- **`Sheet`** — bottom sheet for all logging flows (grabber, `lg` top radius, keyboard-safe, swipe-dismiss with dirty-state guard).
- **`Dialog`** — destructive confirms only.
- **`Toast`** — single-line, bottom, 2.5 s, one at a time; success tone for saves/PRs.
- **`TabBar`** — the 5 fixed tabs; Lucide icons + `micro` labels; active = accent icon+label, 4-pt indicator dot; respects safe-area.

### 6.1 Charts

- Built on **Victory Native XL** (Skia) — approved; performant RN-native charting.
- Style rules: no vertical gridlines; ≤ 3 horizontal reference lines (`chartMuted`); axis text `caption`/`textSecondary`; primary series = `accent` 2.5 pt with soft area fill @ 8%; trend/regression overlay = dashed `positive`/`attention` by classification; target lines = dashed `chartMuted` with right-edge label; touch = single tooltip (value + date), no crosshair clutter; series data arrives pre-bucketed from the engine (ANALYTICS_ENGINE §4) — components never resample.

---

## 7. Theming Mechanics

- Tokens live in `core/theme` as typed objects (`theme.color.accent`, `theme.space.lg`, `theme.type.heading`…). A `useTheme()` hook resolves the active theme; **system-follow by default**, manual override (dark/light/system) stored in MMKV.
- Lint-enforced (rule config in CODING_STANDARDS): no hex literals, no numeric fontSize/margin/padding outside `core/theme` & `core/ui`.
- Charts, haptics, and icons consume the same tokens — one source for the entire sensory system.

---

## 8. Accessibility Baseline

44-pt minimum touch targets everywhere; contrast per §2.3; Dynamic Type to 1.3× without truncating data values; every informative color paired with text/glyph; `accessibilityLabel` on all interactive primitives; Reduce Motion respected (§4); charts accompanied by their ChartFrame interpretation line — the chart is never the sole carrier of a conclusion (this is also why ChartFrame is mandatory).

---

## 9. Do / Don't (P16)

- ✅ One accent moment per view region — ❌ accent-on-everything.
- ✅ Value big, unit small, context line always — ❌ naked numbers ("Trend: +4.2%").
- ✅ Neutral surfaces carrying quiet data — ❌ tinted cards for every category.
- ✅ Lucide icon + label — ❌ emoji tabs, emoji buttons.
- ✅ EmptyState explaining what to log — ❌ blank screens or fake placeholder data.
- ✅ `positive` for improvement per directionality — ❌ green-for-up/red-for-down regardless of goal.

---

## 10. AI Decision Rules (Design System)

1. **No raw style values in features.** Every color/size/font/radius/duration resolves to a token; new needs = new token here first.
2. **No new primitives casually.** Before creating a component, check the catalog; extend a variant before inventing a sibling. New primitives are added to §6 (doc first, code second).
3. **StatTile and ChartFrame enforce interpretation.** Never render a metric or chart outside them; never add a "compact" variant that drops the context line.
4. **Tone ≠ severity theatre.** Insights use `positive`/`attention`/neutral; `danger` is reserved for destructive actions. Never escalate visual alarm to make a card noticeable — priority ordering does that.
5. **Directionality drives color** (FITNESS_DOMAIN §5.3). Never color by raw numeric sign.
6. **Dark and light ship together.** Any styling change must be verified in both themes; a token missing a light value is incomplete.
7. **No emoji in UI chrome. One icon set (Lucide).** Ever.
8. **Accessibility is a gate, not a wishlist:** 44-pt targets, contrast, Dynamic Type 1.3×, Reduce Motion — violations are bugs.
9. **Motion and haptics only via tokens;** no ad-hoc durations, easings, or vibration calls.
10. **Charts follow §6.1 style rules and consume engine-bucketed data** — components never compute or resample analytics.
11. **Feedback is never optional:** no tap without visible response within one frame, no operation longer than 300 ms without visible progress (P15).

---

## Changelog

- 2026-07-08 — v1 baseline frozen (final decisions: accent `#FF6A3D`, Inter, Lucide, Victory Native XL; Feel & Feedback clause and principle citations added per approved refinements).
- 2026-07-09 — §2.2 `textTertiary` darkened/lightened for WCAG AA (Phase 22 CP-E finding F-E1, ratified): dark `#5C6066 → #7E838A` (2.84 → 4.71:1 on surface), light `#9A9EA5 → #6E7278` (2.69 → 4.84:1). An accessibility correction only — no other token or design change; the §2.2 contrast rule (text ≥ 4.5:1) now holds for every token.
