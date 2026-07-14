# Roadmap

This is not a product plan. It records **observable gaps and unfinished areas in
the current code** so they are visible before someone works near them. Each item
points at something real in `index.html`.

## Data lives only on one device

There is no account, server, or sync. All state is in `localStorage`, so clearing
the browser or switching devices loses everything. The only transfer mechanism is
manual **Export / Import** of a JSON file. A sync/backup story is the largest
missing capability.

## Bilingual UI is present but not switchable at runtime

Full English and Arabic dictionaries exist in `I18N`, and RTL handling is written
(`applyLang` sets `dir="rtl"`). However `lang` is declared as a constant
(`const lang='en'`), so the app effectively runs in English and the language
cannot be toggled at runtime as written. Wiring a working language switch is an
open item.

## Two parallel rest-timer implementations

The code contains two independent rest timers:

- a **modal timer** with a circular ring (`openRestTimer`, `setRestTime`,
  `toggleRestTimer`, `cancelRestTimer`), and
- a **bottom-bar countdown** (`startRestTimer`, `stopRestTimer`), which is the one
  invoked when a workout is logged.

They duplicate logic and could be consolidated.

## Reserved but unused storage key

`STORAGE_KEYS.water` (`ft_water`) is defined, but water is actually persisted
per-day inside the `nutrition` entries. The standalone key is currently unused.

## Body metrics are manual only

BMI, body fat %, muscle mass, and visceral fat are typed in by hand from an InBody
scan. There is no calculation or import, even though height/weight are available
for at least a BMI estimate.

## No automated tests

There is no test suite or build tooling. Verification is manual — see
`docs/testing-checklist.md`. Introducing even lightweight automated checks would
reduce regression risk given everything lives in one large file.

## Single-file growth

The whole app is one ~2,700-line `index.html`. This is intentional and keeps
deployment trivial, but as features grow it makes navigation and review harder.
Any future modularization should preserve the no-build, single-deploy property.
