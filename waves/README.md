# Waves Projects Tracker

Arabic (RTL) project-cost tracking dashboard, originally built in the Claude
design tool and exported as a single self-extracting "Bundled Page". This folder
is the **unpacked** form so the design can be reviewed and refined.

## Structure

- `index.html` — the page template (`<x-dc>` markup) plus the `Component` logic
  (the `text/x-dc` script). **This is the file to edit when refining the UI.**
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
