# Testing Checklist

There is no automated test suite. Verify changes manually by opening
`index.html` in a browser (a mobile-width viewport, ~390–430px, matches the
intended layout). Because all state is in `localStorage`, use a private/incognito
window or clear storage to test the empty (fresh-install) state.

## Smoke test (run for any change)

- [ ] App loads with no console errors.
- [ ] All five tabs open: Home, Workouts, Nutrition, Progress, Calendar (bottom
      nav and sidebar both switch correctly).
- [ ] Theme toggle switches dark ⇄ light; colors and the status-bar theme-color
      update.
- [ ] With empty storage, every screen shows its empty state instead of erroring.

## Workouts

- [ ] Log a workout (fresh and from a program); it appears in History, newest
      first.
- [ ] Add and delete exercises and sets; mark a set as warm-up.
- [ ] Warm-up sets are excluded from working-set counts, volume, and PRs.
- [ ] Bodyweight sets (weight 0) display as **BW**.
- [ ] Beating a prior best weight triggers the PR toast; the set shows a PR badge.
- [ ] PRs tab shows correct max weight, est. 1RM, max reps, and best volume.
- [ ] Create, edit, start, and delete a program.
- [ ] Add a custom exercise; it appears in its muscle group in the picker.
- [ ] Rest timer counts down and signals completion.

## Nutrition

- [ ] Log meals under each type; daily calorie and macro totals update.
- [ ] Calorie ring caps at 100%; "remaining" never goes negative.
- [ ] A day over the calorie goal shows a red bar in the Analytics calorie chart.
- [ ] Save a custom food and reuse it as a quick pick.
- [ ] Water: fill/tap glasses (0–8), decrement, and reset.

## Progress

- [ ] Log a weight for today; re-logging the same date overwrites (no duplicate).
- [ ] Weight-goal bar fills correctly for both a loss goal and a gain goal, using
      the historical peak/valley as the effective start.
- [ ] Goal shows "reached" within 0.5 kg of target.
- [ ] Log measurements and InBody metrics; re-logging the same date overwrites.
- [ ] Weight-trend and analytics charts render and update after new data.
- [ ] Add, view full-screen, and delete a progress photo.

## Calendar

- [ ] Workout days are marked; month navigation works; the month summary counts
      workouts and sets correctly; selecting a day shows that day's workouts.

## Data & persistence

- [ ] Reload the page — all logged data persists.
- [ ] Export produces a JSON file; Import replaces data after confirmation.
- [ ] Importing an invalid file shows an error and changes nothing.

## Responsiveness & cross-cutting

- [ ] Layout holds at ~390–430px width with no horizontal scroll.
- [ ] Modals slide up, dismiss via handle/cancel, and respect the safe-area inset.
- [ ] Charts re-color correctly after a theme switch.
