# Business Rules

These are the behavioral rules the app currently enforces in `index.html`. They
are documented from the code, not invented.

## Workouts & sets

- A **set** has `reps` and `weight`. A set may be flagged `warmup`. Warm-up sets
  are excluded from PRs, from working-set counts, and from volume totals
  (`isWorking(s)` returns `!s.warmup`).
- A set with `weight` of `0` is treated as bodyweight and displayed as **BW**.
- A **personal record (PR) marker** on a set means its weight exceeds the maximum
  working-set weight recorded for that exercise across all workouts dated *before*
  this workout's date (`isPR`).
- When a saved workout beats the prior best weight for any exercise, a
  celebration toast and vibration fire (`celebrateNewPRs`).
- The **PRs tab** aggregates, per exercise, the max weight, max reps, best single
  set volume, and best estimated 1RM, sorted by estimated 1RM (`computePRs`).

## Training streak & weekly view

- The **day streak** counts consecutive days ending today on which at least one
  workout is logged; it stops at the first day with no workout (`renderHome`).
- The home weekly strip marks each weekday (Mon–Sun) active if a workout exists on
  that date within the current week. The week starts on **Monday**
  (`(today.getDay()+6)%7`).

## Nutrition

- Meals are grouped by type: **Breakfast, Lunch, Dinner, Snacks**.
- Daily totals are the sum of each meal's calories, protein, carbs, and fat
  (`calcTotals`).
- **Calories remaining** = `max(0, dailyCalories − eaten)` — it never goes
  negative.
- The calorie ring fills to `min(1, eaten / dailyCalories)` — capped at 100%.
- Macro progress bars fill to `min(100, value / goal × 100)` — capped at 100%.
- In the Analytics calorie chart, a day's bar turns red when its calories exceed
  the daily goal; otherwise it is orange.
- Users can save a custom food to reuse it as a **quick pick** (`saveCustomFood`).

## Water

- The daily water goal is **8 glasses**.
- Tapping glass *n* sets the count to *n*; tapping the currently-filled top glass
  again decrements by one (toggle behavior). Water is stored per day inside the
  nutrition entry, and can be reset to 0.

## Weight & goals

- Weight is logged per date; logging the same date **overwrites** that date's
  weight rather than adding a duplicate (`saveWeight`).
- If no `startWeight` is set when the first weight is logged, the app initializes
  `startWeight` to the historical peak weight (max of existing weights and the new
  value) *before* applying the new weight.
- The **weight-goal bar** measures progress from an *effective start weight*, not
  the raw start:
  - When losing (target < current): effective start = the **highest** recorded
    weight (historical peak), so the bar reflects progress from the personal peak.
  - When gaining (target > current): effective start = the **lowest** recorded
    weight.
- The goal is treated as **reached** when current weight is within **0.5 kg** of
  target.

## Measurements & InBody

- Measurements are logged per date and **overwrite** an existing entry for that
  date. Body-measurement fields default to `0`; InBody fields (body fat %, muscle,
  visceral fat, BMI) default to `null` when left blank.
- BMI, body fat %, muscle mass, and visceral fat are **entered manually** (from an
  InBody scan). The app does not calculate them.
- The single "Arms" input is written to both `leftArm` and `rightArm`.

## Programs & exercise library

- A **program** is a named ordered list of exercise names, used to pre-fill a new
  workout. Programs can be created, edited, started, and deleted.
- The exercise **library** is a fixed set grouped by muscle (`EXERCISE_LIBRARY`).
  Users can add **custom exercises** with a name and muscle group; custom entries
  merge into their group and are offered alongside built-ins.

## Photos

- Progress photos are categorized **Front / Side / Back**, stored as base64 data
  URLs in `localStorage`, grouped by date, and can be viewed full-screen or
  deleted.

## Backup & restore

- **Export** writes a JSON file (`fittrack-backup-<date>.json`) containing
  workouts, nutrition, measurements, goals, programs, and custom exercises.
- **Import** replaces current data with the file's contents, but only after the
  user confirms, and only for arrays/objects that pass a basic type check. An
  invalid file shows an error and changes nothing.

## Persistence & seeding

- All state is device-local in `localStorage`; there is no account or sync.
- Demo/seed data generation exists (`seedData`) but is **disabled** — a fresh
  install starts empty.
