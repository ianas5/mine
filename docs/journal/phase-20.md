# Phase 20 — Smart Defaults

**Closed:** 2026-07-09 · **Milestone:** M5 (Refinement) — first phase

The governing principle this phase: **the app must keep getting easier to use as it gets
smarter.** Every smart default has to *reduce* a decision, a tap, or some typing — and it must
save time, never surprise, always be reversible, never lock the user in, and stay transparent.
If a guess is wrong, one tap corrects it. The user should feel the app "knows how I use it"
without ever feeling it decides *for* them. So this phase deliberately **prioritized reducing
friction over adding intelligence** — no new heuristics were invented; the existing ones were
made pure, tested, honest, and less surprising.

## Reduce friction, don't add intelligence (the standing principle)

The UI_UX §5.2 heuristics were largely already built inside the feature phases. The audit
this phase asked of each: is it a **pure function over history** with the fallback chain
`history → sensible static default → ask nothing`, is it **pre-fill only** (never blocks or
asks), and is it **one-tap reversible**? Four already passed; one lived in the wrong layer with
a surprising tie.

## What was built

- **The Log Meal picker heuristic moved from the data layer to pure `domain/nutrition`**
  (`foodPicks.ts`): `mostFrequentSlot`, `aggregateFoodUsage`, and `orderFoodPicks`. The
  repository now only *retrieves* the recent window and delegates every ranking / last-used /
  most-frequent decision to these pure functions (ARCHITECTURE §9.1, ANALYTICS rule 9 — the
  same "SQL retrieves, domain decides" line CP-D verified). Behaviour is otherwise unchanged:
  quick meals pinned, then most-used, then most-recent, then alphabetical; each food carries
  its last-used portion and habitual slot as pre-fills.
- **The meal-slot default no longer surprises on a tie.** The old in-repo `mostFrequent`
  returned an *arbitrary* slot when two slots were used equally often; `mostFrequentSlot` now
  returns `null` on a tie or with no history, so the sheet falls back to the **time-of-day**
  default (§5.2's stated tie rule). A tie is not a habit — deferring is the honest, less
  surprising pre-fill, and the chips remain one-tap editable either way.
- **`aggregateFoodUsage` is order-independent.** Last-used portion is taken from the row with
  the greatest `loggedAt`, not "whichever the query returned first," so the default never
  depends on row order — a small robustness/transparency win.

## What was verified (already live, unchanged)

Each is already a pure, tested function over history, wired pre-fill-only and one-tap reversible:

- **Workout template** — `suggestTemplate` (`domain/fitness`): active program's weekday
  mapping → most-frequent template on this weekday over recent history → Repeat Last → nothing.
  A visible "Suggested for today" card (tie resolves to most-recent, a sensible non-arbitrary
  pick); the user can start any template instead.
- **Photo angle** — `oldestMissingAngle` (`domain/photos`): the oldest-missing angle first,
  else the least-recently-captured; a pre-filled, editable chip.
- **Measurement fields** — `frequentlyLoggedFields` (`domain/body`): fields co-logged in ≥ 50%
  of past sessions start expanded, the rest behind "More sites"; the user can always expand more.
- **Set pre-fill** — last performance for the same exercise + set (shipped in §4.1, Phase 7).

## What changed

New: `domain/nutrition/foodPicks` (+ tests). Modified: `nutritionRepository.getFoodPicks`
(retrieval + delegate, ~45 fewer lines of in-repo logic); nutrition barrel export. No schema,
no new deps, no UI change (the sheet already reads `pick.slot ?? defaultSlotForHour(...)`, which
now benefits from the tie fix automatically), no frozen document changed.

## Screens affected

Log Meal sheet (behaviourally: a tied slot now defers to time-of-day instead of guessing). No
visual change.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Most-frequent slot; tie → null (defer to time-of-day); no history → null; nulls ignored | domain test | ✅ |
| Usage aggregation: counts, last-used from the newest row (order-independent), habitual slot | domain test | ✅ |
| Ordering: quick-meal pin, use-count, recency tiebreak; empty history → alphabetical fallback | domain test | ✅ |
| `nutritionRepository.getFoodPicks` still returns the same shape via the pure functions | repo test | ✅ |
| Template / photo-angle / measurement-field / set heuristics remain pure, tested, wired | inspection | ✅ verified |
| `npm run check` | typecheck + lint + format + 364 tests + db:check (16 tables) | ✅ green (3× stable) |
| On-device smart-defaults walk (a week of varied logging → defaults visibly adapt; fresh install → sane static fallbacks; every default one-tap-correctable) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **Adaptation-over-time is device-only** — that defaults *visibly* adapt after a week of
   varied use, and that a fresh install shows sane static fallbacks, folds into TD-001. What
   *is* verified off-device is every heuristic as a pure function: history → expected default,
   tie/empty → static fallback.
2. **Quick meals are pinned above all regular foods regardless of use count** — an existing
   §4.3 choice, preserved deliberately (not re-litigated this phase); a rarely-used quick meal
   can outrank a daily staple. If it ever grates, it's a one-line comparator tweak, not a
   structural issue.

## Technical debt

None introduced. This phase *reduced* debt in spirit: a heuristic that had grown inside the
repository is now a pure, tested domain function, closing the gap CP-D's finding F-D-style
"SQL retrieves, domain decides" line pointed at for nutrition. No registry entry changes.

## Retrospective

**What went well?** Reading the principle literally — *reduce friction, don't add
intelligence* — turned a potentially sprawling "make the app smarter" phase into a tight,
honest one: audit the five §5.2 heuristics, find the one that wasn't pure/tested/non-surprising,
and fix exactly that. The tie-defers-to-time-of-day change is the clearest embodiment of "never
surprise": an equal split isn't a habit, so the app doesn't pretend it is.

**What was harder than expected?** Resisting scope. It's tempting to add cleverness (meal-time
clustering, portion prediction, weekday macro shaping), but every one of those risks the app
"deciding for you." The discipline was to add none, and instead make the existing guesses
trustworthy and correctable.

**What should change before the next phase?** Nothing structural. Phase 21 (the delight & feel
pass) is next; its rule 11 — *no logging flow gains a tap or a frame of delay* — is the same
friction-first value this phase served, now applied to motion and feel.

## Lessons Learned

- **What surprised you:** how much of §5.2 was already live from the feature phases — the phase
  became a verification-and-correctness pass more than a build, which is the healthy outcome of
  having built each feature with its default in place.
- **What documentation prevented mistakes:** UI_UX §5.2 fixed each heuristic and, crucially, the
  tie rule ("tie/no-history → time-of-day fallback") that exposed the in-repo bug; ARCHITECTURE
  §9.1 fixed that heuristics are pure functions, which is why the nutrition logic belonged in
  `domain/`, not the repository.
- **What should be reused:** the "pure function over history + fallback chain, repository only
  retrieves" shape for every learned default; returning `null` from a heuristic on a genuine tie
  so the caller can defer rather than guess.
- **What should be avoided:** computing a heuristic in the data layer; letting a default pick
  arbitrarily on a tie (it reads as the app deciding for you); adding new "intelligence" that
  trades transparency or reversibility for cleverness.
- **Amendment proposals:** none — no frozen-document defect surfaced. No new debt.
