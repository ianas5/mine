# UI Guidelines

Design conventions used in `index.html`. These describe the existing styling so
new UI stays consistent.

## Layout

- **Mobile-first.** All content lives in `#app`, capped at `max-width: 430px` and
  centered. Design for a phone-width viewport.
- **Sticky header** (`.header`) at the top of each screen with a blurred
  translucent background; a **bottom nav** (`.bottom-nav`) and a slide-out
  **sidebar** (`.sidebar`) provide navigation.
- Respect the safe-area inset: bottom padding uses
  `calc(var(--safe-bottom) + …)` so content clears the home indicator and nav.
- Only one `.screen` is visible at a time (`.screen.active`).

## Theming & color tokens

Colors are CSS custom properties on `:root`, with a `[data-theme="light"]`
override block. Never hard-code hex values that a token already covers.

Dark theme (default):

| Token         | Value     | Use                              |
| ------------- | --------- | -------------------------------- |
| `--bg`        | `#0D0D0D` | Page background                  |
| `--surface`   | `#1A1A1A` | Raised surfaces, modals, sidebar |
| `--card`      | `#222222` | Cards, inputs                    |
| `--border`    | `#333333` | Borders / dividers               |
| `--accent`    | `#FF6B35` | Primary accent (orange)          |
| `--accent2`   | `#00D4AA` | Secondary accent (teal)          |
| `--text`      | `#F0F0F0` | Primary text                     |
| `--text-sub`  | `#888888` | Secondary text                   |
| `--text-dim`  | `#555555` | Tertiary / muted text            |
| `--danger`    | `#FF4757` | Errors, deletes, over-goal       |
| `--warn`      | `#FFD700` | PRs, highlights (gold)           |

Light theme remaps `--bg`, `--surface`, `--card`, `--border`, and the text
tokens; accents stay the same. The `<meta name="theme-color">` and the theme
toggle icon are updated in `applyTheme()`.

## Typography

- Font family: **Outfit** (Google Fonts), falling back to the system UI stack.
- Weights in active use range from 400 to 900; headings are heavy (700–800) with
  tight negative letter-spacing.
- Uppercase micro-labels (`.card-title`, `.field label`, section labels) use
  ~10–12px, weight 700, letter-spacing, and `--text-sub`.

## Components

- **Cards** — `.card` (18px radius, 1px hairline border) and `.card-sm`.
  `.grid-2` lays out two stat cards side by side.
- **Buttons** — `.btn` base with variants `.btn-primary` (accent fill),
  `.btn-secondary`, `.btn-ghost`, `.btn-danger`, plus `.btn-full` / `.btn-sm`.
  Buttons scale down slightly on `:active`.
- **FAB** — a floating `+` action button (`.fab`) anchored bottom-right within the
  430px column, with an accent glow shadow.
- **Modals** — `.modal-overlay` + `.modal`. Modals slide up from the bottom
  (`translateY(100%)` → `0`), have a rounded top and a grab handle
  (`.modal-handle`), and are opened/closed with `openModal` / `closeModal`.
- **Rings** — SVG progress rings (`.ring-*`) rotated −90° so they fill from the
  top; driven via `stroke-dashoffset` (see `docs/formulas.md`).
- **Bars** — macro bars (`.macro-bar-*`) and the weight-goal bar fill by animating
  `width`.
- **Chips / badges** — `.chip` for quick selections, `.badge` for gold accents.
- **Toast** — a single `#toast` element; call `showToast(msg)` for transient
  feedback (auto-hides after ~2.5s).

## Motion & feedback

- Transitions are short (0.15–0.35s) with ease/`cubic-bezier` curves; avoid long
  animations.
- Use `navigator.vibrate(...)` for haptic feedback on meaningful events (PRs, rest
  timer done) — always guard with `if(navigator.vibrate)`.
- `-webkit-tap-highlight-color: transparent` is used to suppress the mobile tap
  flash on interactive elements.

## Iconography & copy

- Icons are a mix of inline SVG (nav) and emoji. Recent changes deliberately
  **reduced emoji usage** — prefer restraint.
- All user-facing text should route through `t(key)` with an entry in the `I18N`
  dictionaries (English and Arabic). For Arabic, the layout switches to `dir="rtl"`.
