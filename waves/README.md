# Waves Projects Tracker

Arabic (RTL) project-cost tracking dashboard, originally built in the Claude
design tool and exported as a single self-extracting "Bundled Page". This folder
is the **unpacked** form so the design can be reviewed and refined.

## Viewing it

- **`preview.html`** — open in any browser. Loads the vendored React/ReactDOM
  first, then the dashboard, so it runs standalone.
- **`waves.bundle.html`** — the refined design re-packed into a single
  self-contained file in the original "Bundled Page" export format (manifest +
  template + unpacker). Drop this back into the Claude design tool / viewer.
  Like the original export, it expects the viewer to provide React, so it does
  **not** run by double-clicking in a plain browser — use `preview.html` for that.

Both are generated from `index.html`; edit `index.html` and regenerate.

## Structure

- `index.html` — the page template (`<x-dc>` markup) plus the `Component` logic
  (the `text/x-dc` script). **This is the file to edit when refining the UI.**
- `preview.html` — `index.html` + vendored React, runnable in a plain browser.
- `vendor/` — React/ReactDOM UMD builds (preview only).
- `assets/`
  - `lib-1.js` — `dc-runtime`, the design tool's rendering framework. Generated; do not edit.
  - `lib-2.js` — `WaveParser`: in-browser `.xlsx` reader, exposes `window.WaveParser`.
  - `lib-3.js` — `window.WAVES_DATA`: baked-in JSON snapshot of all project data.
  - `overviewFile.xlsx`, `ssFile.xlsx`, `smroFile.xlsx`, `ppFile.xlsx` — source data.
  - `image-1.png` — Waves logo.
  - `font-*.woff2` — IBM Plex Sans Arabic (weights 300–700).
- `asset-map.json` — maps the original bundle UUIDs to the extracted filenames.

## How it runs

1. `lib-1` (runtime) compiles the `<x-dc>` template + `Component` class.
2. The `WAVES_DATA` snapshot (`lib-3`) renders instantly.
3. `WaveParser` (`lib-2`) re-reads the four `.xlsx` files for live data when
   available (falls back silently to the snapshot otherwise).

## Tabs

- النظرة العامة (Overview) — project cards grouped by department (PP / SMRO / SS),
  with department + wave filters and a detail drawer (value-chain stepper + timeline).
- متابعة الإدارة (Management) — shared administrative tracker items.
- المعوقات (Issues) — blockers log across departments.

## Refinements applied

Surfaced functionality the logic already computed but never rendered, plus a
visual + accessibility pass:

- **KPI strip** on Overview (total / done / on-track / late / risk / open issues),
  with the open-issues card linking to the Issues tab.
- **Search box** (name / specialist / id) and **status filter chips** wired to the
  existing `search` and `status` state.
- **Theme switcher** in the header (هادئ / متوازن / داكن) exposing the three palettes.
- **Per-group status bar + legend** in each department header.
- **Signature element:** an animated gold wave crest drifting beneath the header —
  the "wave" the product is named for. Honors `prefers-reduced-motion`.
- **Accessibility:** keyboard-operable project cards (focusable, Enter/Space,
  `aria-label`), visible `:focus-visible` rings on all controls.
- **Responsive:** fluid gutters and wrapping toolbars/headers down to mobile.

The runtime (`lib-1`), parser (`lib-2`), and data (`lib-3`) were not modified; all
changes live in `index.html` (template + `Component`).

## Branding

Restyled onto **Anthropic's brand**: palette (dark `#141413`, light `#faf9f5`,
orange accent `#d97757`, blue/green secondaries) applied across the three themes,
and Poppins (display/numerals) + Lora (body) loaded from Google Fonts. Arabic text
stays on IBM Plex Sans Arabic via per-glyph font fallback, since Poppins/Lora don't
cover Arabic. Status colors map to brand green/blue/orange, with red kept for risk.

## Documentation

See **`DOCUMENTATION.ar.md`** (Arabic) for the end-user guide: tabs, data files and
their Excel layout, how to update data, branding, and how to re-bundle.
