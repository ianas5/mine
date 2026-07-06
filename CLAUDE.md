# CLAUDE.md

Guidance for working in this repository.

## What this project is

**FitTrack Pro** — a personal fitness and health tracking app for phones. It
tracks workouts, nutrition, water, body weight, body measurements, InBody
metrics, and progress photos, and visualizes trends with charts.

The entire application is a single file: `index.html`. There is no build step,
no framework, no package manager, and no backend. It is a client-side
Progressive Web App that persists everything in the browser's `localStorage`.

## Repository layout

| Path               | Purpose                                                        |
| ------------------ | ------------------------------------------------------------- |
| `index.html`       | The whole app: markup, CSS (`<style>`), and JS (`<script>`).   |
| `icon.png`         | App / apple-touch icon referenced by `<link rel="icon">`.      |
| `font-preview.html`| Standalone page for previewing the Outfit font choice.         |
| `README.md`        | Minimal readme.                                                |
| `docs/`            | Project knowledge base (see below).                            |

## Knowledge base

- `docs/architecture.md` — how `index.html` is organized and how data flows.
- `docs/business-rules.md` — the behavioral rules the code enforces.
- `docs/formulas.md` — every calculation used (1RM, volume, goal progress…).
- `docs/ui-guidelines.md` — design tokens, layout, and component conventions.
- `docs/roadmap.md` — observable gaps and unfinished areas in the current code.
- `docs/testing-checklist.md` — manual checks to run before shipping a change.

## Tech stack

- Vanilla HTML/CSS/JavaScript, no transpiler.
- [Chart.js 4.4.0](https://cdn.jsdelivr.net/npm/chart.js@4.4.0) loaded from a CDN
  (`<script src>` in `<head>`) — used for all charts.
- [Outfit](https://fonts.googleapis.com/css2?family=Outfit) web font from Google
  Fonts.
- Persistence: `localStorage` only. No server, no external API calls for data.

## How to run

Open `index.html` in a browser, or serve the directory with any static file
server. No install or build is required. The layout is designed for a mobile
viewport (`#app` is capped at `max-width: 430px`).

## Conventions

- Everything lives in `index.html`. Keep markup, styles, and script in that one
  file unless a change specifically requires splitting.
- JavaScript is written in a terse, dependency-free style (short helpers, direct
  DOM manipulation via `innerHTML`, no modules). Match the surrounding style.
- All persisted state goes through the `load()` / `save()` helpers and the
  `STORAGE_KEYS` map — do not read or write `localStorage` keys directly.
- Rendering is pull-based: a `render*()` function reads from storage, computes,
  and rewrites a container's `innerHTML`. After mutating data, call the relevant
  `render*()` (and usually `renderHome()`).
- User-facing strings go through the `t(key)` translation helper and the `I18N`
  dictionary (English and Arabic entries exist).

## Git / branch policy

Active development branch for the current task: `claude/project-knowledge-base-bt206y`.
Commit with clear messages and push to the designated branch. Do not open a pull
request unless explicitly asked.
