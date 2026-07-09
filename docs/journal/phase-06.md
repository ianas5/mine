# Phase 6 — Crash Safety, Rest Timer & Session Bar

**Closed:** 2026-07-09 · **Milestone:** M1 Training

The active session becomes the most valuable data in the app — and loss-proof.
A crash, OS eviction, or force-close no longer costs a workout: the session is
checkpointed to SQLite and silently restored on the next launch, so recovery is
invisible. The session is also now app-wide (a docked session bar, one tap back)
and the rest between sets is timed with a gentle haptic at zero.

## What was built

- **`workout_drafts` table + migration 0003** (DATABASE §3.4): a single-row
  (`id = 1`) JSON checkpoint of the in-memory session. Deliberately a blob, not
  relational — a recovery snapshot, rewritten on debounce/background, deleted on
  finish/discard, and excluded from every backup path.
- **Draft serialization + Zod validation** (`sessionDraftSchema`):
  `serializeSession` writes a versioned payload; `parseSessionDraft` validates it
  and returns `null` for anything corrupt, out-of-shape, or a legacy version, so
  a bad draft is discarded gracefully rather than crashing. localIds are **not**
  persisted — they are regenerated on restore so a recovered session can never
  collide with ids minted afterward.
- **Repository draft methods** (`workoutRepository`): `checkpointDraft` (single-row
  upsert; high-frequency, off the change-bus), `loadDraft`, `discardDraft`. And
  the finish path now **deletes the draft inside the same transaction** as the
  workout insert (ARCHITECTURE §7.1.2) — the saved workout and the vanished draft
  commit together, so a crash can never leave both or neither.
- **Store restore + recovery flag** (`useSessionStore`): a `restore(draft)` action
  and a `recovered` flag that drives the resume banner; `acknowledgeRecovery`
  dismisses the banner without ending the session.
- **`SessionKeeper`** (mounted once at the composition root, inside the DB gate):
  the crash-safety engine. On launch it runs `recoverSession`; while a session is
  live it checkpoints **immediately on structural changes** (exercise added, set
  completed — tracked by a cheap fingerprint) and **debounced on value edits**
  (typing a weight), and **flushes on app-background**. On finish/discard it
  ensures no draft lingers.
- **Rest timer** (`useRestTimerStore` + `useRestCountdown`): auto-starts on a
  **working** set's ✓ (90 s default per UI_UX §4.2; per-exercise `rest_seconds`
  arrives with templates in Phase 8, so the store remembers in-session
  extensions per exercise instead). Wall-clock based (`endsAt`), so it stays
  **honest across background/foreground** — no decrement drift. `RestTimerBar`
  gives a slim, non-blocking countdown with **Skip** and **+30 s**; a single
  gentle haptic fires exactly once at zero.
- **Persistent session bar** (`SessionBar`, docked above the tab bar app-wide via
  a custom `tabBar` so content reflows above it, no floating overlay): elapsed ·
  current exercise · rest countdown when running (UI_UX §2.2/§5.1). One tap
  returns to the live session. Renders nothing when no workout is active.
- **Resume/discard banner** on the Workouts home: a recovered session shows
  "Workout recovered … pick up right where you left off" with **Resume** and a
  Dialog-gated **Discard**.
- **Tests (12 new, 127 total):** draft schema round-trip + corrupt/legacy → null;
  rest-timer store (default/extend-remembers-pref/skip-keeps-prefs-reset-forgets)
  and `formatCountdown`; and the phase's **defining recovery suite** over real
  SQLite — force-kill → restore *exact* state, finish deletes the draft
  transactionally, corrupt draft discarded gracefully, empty draft not recovered,
  single-row upsert.

## Resilience beyond the roadmap (per the standing directive)

The session was treated as if death can strike at any instant:

- **Structural events persist synchronously, not just on a debounce timer.** The
  highest-value moments — a set completed, an exercise added — write the draft the
  instant they happen; only low-value keystroke edits coalesce (600 ms). Worst-case
  loss is a sub-second window of *typing*, never a logged set.
- **Recovery is invisible.** A valid draft is silently rehydrated on launch — the
  session bar simply reappears and the workout is *still there*; the banner is an
  explicit offer layered on top, not a gate the user must clear.
- **Finish is atomic with draft deletion** — no window where a completed workout
  and its draft both exist (which would double-recover) or both vanish.
- **Background is a flush point**, catching the OS-eviction path the debounce
  might otherwise miss.

## What changed

New: `workout_drafts` table + migration 0003; `schemas/sessionDraftSchema`;
`logic/sessionRecovery`; `stores/useRestTimerStore`; `hooks/useRestCountdown`;
components `SessionKeeper`/`SessionBar`/`RestTimerBar`. Modified:
`useSessionStore` (restore + recovered flag), `workoutRepository` (draft methods +
transactional draft delete on finish), `SetRow` (auto-start rest on working ✓),
`ActiveWorkoutScreen` (rest bar + reset-on-leave), `WorkoutsHomeScreen` (recovery
banner), `app/_layout` (mount `SessionKeeper`), `app/(tabs)/_layout` (session bar
above the tab bar). Defect fix: `exerciseRepository` duplicate-name detection no
longer gates on `instanceof Error` (unreliable across Jest module realms for the
native driver's error type). No frozen document changed.

## Screens affected

Every tab screen (session bar now docks above the tab bar during a workout),
Active Workout (rest bar), Workouts home (recovery banner). App launch now runs
crash recovery before the first frame of content.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Force-kill mid-session → relaunch restores exact state | recovery suite (real SQLite) | ✅ |
| Finish deletes the draft transactionally (no double-recover) | recovery suite | ✅ |
| Corrupt / legacy-version draft discarded gracefully (no crash) | schema + recovery suites | ✅ |
| Empty session draft not recovered | recovery suite | ✅ |
| Checkpoint upserts a single row (id = 1) | recovery suite | ✅ |
| Draft Zod round-trip (serialize → parse) | schema test | ✅ |
| Rest auto-start default / extend-remembers / skip / reset | rest-store test | ✅ |
| `formatCountdown` m:ss rounding | rest test | ✅ |
| `npm run check` | typecheck + lint + format + 127 tests + db:check (6 tables) | ✅ green |
| On-device crash-safety & rest & session-bar walk (kill mid-set → resume; 10-min background during rest → honest timer; navigate tabs via bar; rest skip/extend/haptic; both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** for the crash-safety/rest/session-bar flows
   (expo-sqlite native-only; the timer and session bar are device feel). Behaviour
   is proven by the recovery/rest/schema suites; visuals + the kill→resume walk
   fold into the consolidated TD-001 device pass (checklist extended below).
2. **The session bar composes expo-router's internal `BottomTabBar`** (imported
   from its build path, since expo-router does not re-export it at the top level).
   This is the exact component behind `Tabs`; type-resolved and runtime-present,
   but the tab-bar-with-session-bar layout is a device-verification item in TD-001.
3. **Dashboard-side Focus-Mode subtractions deferred** — the Dashboard is a
   placeholder until Phase 16, so its insight-card hiding / live-session-card /
   slim-quick-actions cannot be built yet. The core §5.1 mechanism (app-wide
   session bar, one-tap return, no non-session toasts) is delivered. Registered as
   **TD-006**, removed when the Dashboard lands.
4. **Rest is transient across process death** (by design): a crash loses only the
   seconds-scale rest countdown, never the workout. Per-exercise persisted
   `rest_seconds` seeds from templates in Phase 8.

## Technical debt introduced

- **TD-006** — dashboard-side Focus-Mode subtractions (UI_UX §5.1) deferred until
  the Dashboard exists (Phase 16). The session bar — §5.1's essential piece — is
  built.

## Retrospective

**What went well?** The "single source of truth, everything re-derived" spine from
Phase 4/5 paid off again: crash safety is *serialize the store → one SQLite row*,
and recovery is *validate → restore*, with no derived state to reconcile.
Transactional draft-deletion on finish fell straight out of the existing
`runInTransaction` save. Making the rest timer wall-clock-based (store `endsAt`,
render-derived remaining) meant background/foreground honesty came for free and
the whole thing is a pure store + a tiny hook, fully unit-testable without a
device.

**What was harder than expected?** Two things. First, docking arbitrary content
above Expo Router's tab bar has no top-level API — the default `BottomTabBar` is
only reachable via an internal build path; I chose the custom-`tabBar` approach
(content reflows correctly) over a floating overlay (would cover screen content)
and documented the deep import. Second, a **rare Jest flake** surfaced: adding a
fifth better-sqlite3 node suite shifted module-load order and exposed that
`instanceof Error` is unreliable across Jest's per-file realms for a native
driver's error type — so the exercise-repository's duplicate-name guard
occasionally let a real UNIQUE violation escape unmapped. Root-caused it (the
error message *always* matched; the `instanceof` gate was the culprit) and fixed
detection to read message/code without `instanceof`. 15 consecutive full runs
green after.

**What should change before the next phase?** Nothing structural. The
`useTableVersion`-style reactivity and the store+pure-logic+thin-screen pattern
continue to carry their weight. Watch the deep `BottomTabBar` import across future
expo-router upgrades (it has no `exports` map today, so it resolves — a version
bump could change that).

## Lessons Learned

- **What surprised you:** the phase's scariest-sounding requirement — "never lose
  an in-progress workout" — reduced to *one JSON row upserted on a debounce, plus
  a transactional delete on finish*, because the session was already a single
  serializable store. And a *test-only* fragility (`instanceof` across Jest
  realms) masqueraded as a product flake; the fix (duck-type the driver error)
  makes the code more robust even though on-device Hermes has one realm.
- **What documentation prevented mistakes:** ARCHITECTURE §7.1 fixed the exact
  contract — checkpoint on meaningful mutations + background, single durable write
  on finish that also deletes the draft, Zod-validate before resume, discard the
  invalid gracefully — so each clause became a test. DATABASE §3.4 gave the
  single-row (`id = 1`) blob shape and the "excluded from backups" rule. UI_UX
  §2.2/§4.2/§5.1 fixed the 90 s default, wall-clock rest countdown in the session
  bar, and the docked-bar-one-tap-return mechanism. Discovering `rest_seconds`
  lives on `template_exercises` (not `exercises`) kept me from inventing a schema
  column — it is a Phase 8 seed source, not a Phase 6 gap.
- **What should be reused:** serialize-store→one-row + validate→restore as the
  crash-safety pattern for any future in-progress editor; the structural-signature
  trick (persist high-value events now, coalesce the rest) to balance durability
  against write frequency; wall-clock `endsAt` + render-derived remaining for any
  honest timer; duck-typing driver errors instead of `instanceof` in `catch`.
- **What should be avoided:** `instanceof` on native-library error types inside a
  `catch` (cross-realm-fragile under Jest); caching anything derived from the
  session (it would make recovery a reconciliation problem instead of a restore);
  floating overlays for docked chrome when a layout slot (`tabBar`) reflows content
  correctly.
- **Amendment proposals:** none — no frozen-document defect surfaced. TD-006
  records the dashboard-dependent Focus-Mode work as scheduled, not a doc change.
