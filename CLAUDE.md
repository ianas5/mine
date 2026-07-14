# CLAUDE.md

Permanent instruction set for this project. Read and follow this on every task.
Detailed business rules, formulas, and UI specifications live in `/docs` — not
here.

## Project mission

FitTrack Pro is a personal fitness and health tracking app. It helps one user
log workouts, nutrition, water, body weight, measurements, and progress photos,
and understand their trends over time. The app must stay fast, private, and
effortless to use on a phone. The user's data belongs to the user: it stays on
their device unless they explicitly export it.

## Development philosophy

- **Simplicity over cleverness.** Prefer the smallest change that fully solves
  the problem. Do not add abstraction, tooling, or dependencies the project does
  not need.
- **Preserve what works.** This is a single-file, no-build app. Keep it that way
  unless a change genuinely requires otherwise, and never break the ability to
  run it by opening the file directly.
- **Protect user data.** Persistence is local and irreplaceable. Never write
  changes that could silently corrupt, wipe, or migrate stored data without a
  clear, safe, reversible path.
- **Ship deliberately.** Small, verified, self-contained changes beat large
  speculative ones.

## General coding standards

- Match the surrounding code's style, naming, and density. Consistency with the
  existing file matters more than personal preference.
- Keep the app dependency-free and buildless. Do not introduce frameworks,
  package managers, or transpilers.
- Route all persisted state through the project's storage helpers and key map —
  never touch `localStorage` directly.
- Route all user-facing text through the translation layer; do not hard-code
  strings that should be translatable.
- Reuse existing helpers, tokens, and components before writing new ones. Avoid
  duplicating logic.
- Write clear, self-explanatory code. Comment the "why" when intent is not
  obvious; do not narrate the obvious.

## Architecture principles

- One app, one file: markup, styles, and logic stay together unless a change
  specifically demands separation.
- Rendering is pull-based: read from storage, compute, then update the view.
  After mutating data, refresh the affected view(s).
- Keep concerns separated by function and section as the file already does
  (storage, navigation, per-screen rendering, cross-cutting features).
- Derive values from stored data at render time rather than storing redundant
  computed state.
- Keep styling driven by the shared design tokens; do not scatter one-off values
  that a token already covers.

## Workflow before making changes

1. **Understand first.** Locate the relevant code and read it before editing.
   Check `/docs` for the rules or formulas that govern the area.
2. **Confirm scope.** Know exactly what should change and, just as importantly,
   what should not. If the request is ambiguous, ask before acting.
3. **Plan the smallest correct change.** Identify the specific functions,
   sections, and storage keys involved.
4. **Make the change** consistently with existing patterns.
5. **Verify** the change works and nothing adjacent regressed.

## Validation mindset

- Treat stored data as potentially empty, partial, or malformed. Handle the
  empty/first-run state and guard against missing fields.
- Never assume optional browser APIs exist — feature-check before using them.
- Preserve invariants defined in `/docs`; if a change would alter documented
  behavior, update the docs in the same change.
- Prefer failures that are visible and safe over silent data loss.

## Definition of done

A change is done only when all of the following hold:

- It fully addresses the request, with no unrelated changes bundled in.
- It loads with no console errors and the empty/first-run state still works.
- The affected flows were manually verified (see `docs/testing-checklist.md`).
- Existing behavior, data, and design conventions are preserved unless the change
  intentionally updates them.
- Any documented rule, formula, or spec affected by the change is updated in
  `/docs`.
- The commit message clearly states what changed and why.

## Maintaining consistency

- `/docs` is the source of truth for detailed behavior. When code and docs
  disagree, reconcile them — do not leave them out of sync.
- Follow the established conventions for storage access, translation, rendering,
  and styling on every task, so the project reads as if written by one author.
- When introducing a new pattern, prefer extending an existing one; if a new
  convention is truly warranted, apply it consistently and note it.
