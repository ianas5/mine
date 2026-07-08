# Technical Debt Registry

Per IMPLEMENTATION_ROADMAP §2: every intentionally accepted shortcut is recorded
here at merge time with **why it exists · why it's acceptable now · when it must
be removed**. Checkpoints audit this registry; an overdue entry blocks the
checkpoint.

| ID | Introduced | Description | Why it exists | Why it's acceptable now | Must be removed by |
|---|---|---|---|---|---|
| TD-001 | Phase 0 | Physical-iPhone launch of the dev build not verified (only web-build rendering at iPhone dimensions). | Phase 0 was executed in a remote Linux environment with no macOS/iOS toolchain or device. | The shell is platform-thin (routes, theme, fonts, icons); react-native-web exercised the same code paths; risk is confined to native build config. | **CP-A** — owner runs `npx expo run:ios` on a physical iPhone and confirms cold start, tab switching, theme follow, kill & relaunch. |
| TD-002 | Phase 0 | The five placeholder screens duplicate their layout markup (~40 lines × 5). | A shared placeholder component would live in `core/ui`, which is deliberately empty until Phase 1; features must not import features. | Placeholders are disposable — each is deleted when its real screen lands (Phases 3–16); abstraction before a durable second consumer violates CODING_STANDARDS rule 9. | Naturally, as each phase replaces its placeholder; any still present at **CP-D** must be reviewed. |
| TD-003 | Phase 1 | `Sheet` dismissal works via backdrop tap / Android back / programmatic close, but the grabber is not yet a drag-to-dismiss gesture (DESIGN_SYSTEM §6 specifies swipe-dismiss). | A correct drag gesture needs react-native-gesture-handler wiring and device-tuned physics that can't be validated in this environment. | All dismissal paths users need exist and the dirty-state guard covers them; the missing piece is gesture polish, not function. | **Phase 9** (first heavy logging-sheet usage) or the **Phase 21** feel pass, whichever comes first. |
