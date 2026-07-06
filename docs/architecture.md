# Architecture

FitTrack Pro is a single-file client-side app. `index.html` contains three
parts in order: the `<head>` (metadata, fonts, Chart.js), the `<body>` markup
(all screens and modals), and one `<script>` block that holds all logic.

## Runtime model

- **No framework, no build.** Plain HTML, a single `<style>` block, and a single
  `<script>` block. The browser runs it directly.
- **No backend.** All data lives in the browser's `localStorage`. Nothing is
  sent to a server; the only way data leaves the device is the manual
  Export / Import feature.
- **Mobile-first single view.** Everything renders inside `#app`, which is capped
  at `max-width: 430px` and centered. It is set up as an installable PWA
  (`apple-mobile-web-app-capable`, theme-color meta, apple-touch-icon).

## Screens and navigation

The app is a set of `.screen` elements; only the one with `.active` is shown.
Navigation is handled by `switchTab(tab)`, driven by both a bottom nav bar
(`.bottom-nav`) and a slide-out sidebar (`.sidebar`).

Top-level tabs:

- **Home** (`renderHome`) — greeting, today's calories ring, macros, weekly
  training dots, weight-goal bar, last workout, today's workout.
- **Workouts** (`renderWorkouts`) — sub-tabs: History, PRs, Programs, Library.
- **Nutrition** (`renderNutrition`) — per-day meals by type, water intake.
- **Progress** (`renderProgress`) — sub-tabs: Log (weight), Analytics, Measurements, Photos.
- **Calendar** (`renderCalendar`) — month grid of workout days plus a summary.

Sub-tabs are switched with `switchWorkoutTab()` and `switchProgressTab()`.

## Data storage

All persisted keys are defined once in `STORAGE_KEYS`:

| Key (constant)     | localStorage key       | Holds                                          |
| ------------------ | ---------------------- | ---------------------------------------------- |
| `workouts`         | `ft_workouts`          | Array of logged workout sessions.              |
| `nutrition`        | `ft_nutrition`         | Array of per-date entries (`meals`, `water`).  |
| `measurements`     | `ft_measurements`      | Array of per-date weight + body measurements.  |
| `goals`            | `ft_goals`             | Single goals/settings object.                  |
| `water`            | `ft_water`             | Reserved water key (water is stored per-day in `nutrition`). |
| `programs`         | `ft_programs`          | Array of saved workout programs.               |
| `customExercises`  | `ft_custom_exercises`  | User-added exercises (name + muscle group).    |
| `customFoods`      | `ft_custom_foods`      | User-saved quick-pick foods.                   |
| `photos`           | `ft_photos`            | Progress photos (base64 data URLs).            |

Two additional standalone keys hold UI preferences: `fit_theme` (`dark` / `light`)
and `fit_lang`.

Access is always through two helpers:

```js
function load(k){try{return JSON.parse(localStorage.getItem(k))||null}catch(e){return null}}
function save(k,v){localStorage.setItem(k,JSON.stringify(v))}
```

### Key data shapes (as written by the app)

- **Workout:** `{ id, date, name, exercises: [{ name, notes?, sets: [{ reps, weight, warmup? }] }] }`
- **Nutrition entry:** `{ date, meals: [{ name, type, calories, protein, carbs, fat }], water }`
- **Measurement:** `{ date, weight, chest, waist, hips, leftArm, rightArm, leftThigh, rightThigh, bodyfat?, muscle?, visceral?, bmi? }`
- **Goals:** `{ dailyCalories, startWeight, targetWeight, proteinGoal, carbsGoal, fatGoal, startDate? }`
- **Program:** `{ id, name, exercises: [exerciseName, …] }`
- **Photo:** `{ id, date, category, dataUrl }` (category is `front` / `side` / `back`)

## Rendering pattern

Rendering is pull-based and imperative:

1. A `render*()` function calls `load()` for the keys it needs.
2. It computes derived values with the helpers in `docs/formulas.md`.
3. It rewrites a container's `innerHTML` (and updates charts via `mkChart`).

Mutations follow the inverse: read with `load()`, mutate the array/object,
`save()` it, then call the relevant `render*()` to refresh the view. `renderHome()`
is refreshed after most mutations because the home screen summarizes everything.

## Charts

Chart.js is loaded from a CDN. Chart instances are cached in `chartInstances` and
recreated through a `mkChart(id, config)` helper so re-renders destroy the prior
instance first. `renderProgress` defers if `Chart` is not yet defined
(`if(typeof Chart==='undefined'){ setTimeout(renderProgress,600); return; }`).
`chartColors()` adapts grid/tick colors to the active theme.

## Cross-cutting concerns

- **Theme** — `applyTheme()` / `toggleTheme()` set `data-theme` on `<html>`,
  swap the theme-color meta, and persist to `fit_theme`. Light overrides live in
  the `[data-theme="light"]` CSS rules.
- **Internationalization** — `I18N` holds English (`en`) and Arabic (`ar`)
  dictionaries; `t(key, ...args)` resolves a string or a function entry and falls
  back to English. `applyLang()` fills `[data-i18n]` / `[data-i18n-ph]` elements
  and sets `dir="rtl"` for Arabic.
- **Rest timer** — two implementations coexist: a modal timer with a circular
  ring (`openRestTimer` / `setRestTime` / `toggleRestTimer` / `cancelRestTimer`)
  and a bottom-bar countdown (`startRestTimer` / `stopRestTimer`), the latter
  triggered when logging a workout.
- **Backup** — `exportData()` serializes the main keys to a downloadable JSON
  file; `importData()` reads such a file and replaces stored data after a
  confirmation prompt.

## Initialization

At the end of the `<script>`, the app runs `applyTheme()`, `applyLang()`,
`renderHome()`, and `renderQuickPrograms()`, then marks the Home nav item active.
A `seedData()` function exists but is **not called**, so a fresh install starts
with no data.
