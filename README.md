# Personal Fitness Tracker

A private, single-user mobile fitness tracking app — workouts, nutrition, body
measurements, progress photos, and analytics that explain progress over time.

See [`PROJECT_VISION.md`](./PROJECT_VISION.md) for the full product vision.

## Status

The repository has been **reset to a clean slate** for a complete greenfield
rebuild. The previous prototype (a single-file web PWA) has been removed. No
application code, architecture, or scaffolding exists yet — implementation is
intentionally paused until the dedicated architecture documents are provided.

## Intended Technology Stack

The rebuild targets a mobile-first app (primarily iPhone, with Android as a
secondary goal) on the following baseline stack:

- **Expo** (latest stable SDK)
- **React Native**
- **TypeScript** — strong typing throughout
- **Expo Router** — navigation
- **Local-first architecture**
- **SQLite** — primary database
- **Zustand** — lightweight global state
- **React Hook Form** — forms
- **Zod** — validation
- **MMKV** — lightweight settings and cached preferences
- **Feature-based project structure** with a reusable component architecture

Designed for long-term maintainability and scalability.

> Architecture documents will be provided next and take precedence over any
> assumptions above.

## Preserved Assets

- `icon.png` — 1024×1024 app icon, retained for reuse in the new build.
