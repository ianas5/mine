# Phase 13 — Progress Photos

**Closed:** 2026-07-09 · **Milestone:** M2 (Fuel & Body)

Progress photos land as the last M2 logging surface: capture front/side/back shots,
see them on a date-grouped timeline, and compare any two of the same angle
side-by-side or with a Before/After toggle. The governing rule this phase — **the
filesystem and the database behave as a single logical transaction**: there is never
a DB row pointing at a missing file, and never an orphaned file without metadata
except as a recoverable startup state the orphan sweep reconciles (DATABASE §3.6).

## What was built

- **Photo domain** (`domain/photos`, pure): the `PHOTO_ANGLES`
  (front → side → back) tuple, labels, `oldestMissingAngle(photos)` — the smart
  capture default (a never-captured angle first, in canonical order, else the angle
  whose most-recent photo is the oldest) — and `groupPhotosByDate(photos)` for the
  newest-first timeline.
- **`ProgressPhoto` model + `progress_photos` table** (migration 0008): `id`, `date`,
  `angle` (CHECK front/side/back), `fileName` (UNIQUE), `width`/`height`/`notes`,
  `createdAt`, indexed on `date`. Fifteen tables total.
- **Photo store abstraction** (`data/photos/photoStore.ts` — interface + injectable
  registry; `expoPhotoStore.ts` — the real SDK-57 `File`/`Directory`/`Paths`
  implementation over the app's `photos/` document dir). The interface carries no
  expo import, so the repository's file/DB lifecycle is testable node-side against an
  in-memory fake store — no device required.
- **`photoRepository`** — the single-transaction lifecycle (DATABASE §3.6):
  - `savePhoto`: **write the file first, then insert the row**; if the insert throws,
    delete the just-written file and rethrow — a row never points at a missing file.
  - `deletePhoto`: **delete the row first, then the file** — a crash between the two
    leaves an orphan file, which is the recoverable state.
  - `sweepOrphans` (startup): delete files with no metadata row (e.g. a kill between
    file-write and row-insert) and count rows whose file is missing (rendered as a
    placeholder, never deleted).
  - `listPhotos`: each row annotated with its render `uri` and a `fileMissing` flag.
- **Composition-root wiring** (`app/_layout.tsx`): `setPhotoStore(expoPhotoStore)` at
  module scope (like the SQLite handle) so `core`/`data` never import expo directly;
  a `<PhotoSweeper />` runs the orphan sweep once on launch inside `DbGate`.
- **Capture flow** (`logic/pickPhoto.ts`): `pickFromLibrary()` and
  `captureFromCamera()` over `expo-image-picker` (camera requests permission first),
  each returning `{ uri, width, height }` or null on cancel.
- **Add Photo sheet**: angle chips defaulting to `oldestMissingAngle`, "Choose from
  library" / "Take photo" — save is one pass into `photoRepository.savePhoto` with
  today's date; a success toast, an error toast on failure (the file is cleaned up).
- **Photos timeline screen** (`/measurements/photos`): date-grouped gallery (newest
  first) of angle-badged thumbnails; tap opens a full-size viewer with a
  Dialog-guarded delete; a missing file renders a "File missing" placeholder, never a
  broken image.
- **Photo compare screen** (`/measurements/photos/compare`): angle chips limited to
  captured angles, Before defaulting to the earliest of that angle and After to the
  latest; a segmented **Side by side / Before / After** switch — side-by-side renders
  both, Before/After swaps a single full-width image on tap with a date label. Both
  date pickers list only that angle's photos.
- **Measurements home**: a Progress-photos entry (always visible — photos are
  independent of body-measurement data).
- **Tests (8 new, 218 total):** domain — `oldestMissingAngle` (never-captured-first
  and oldest-most-recent) and `groupPhotosByDate`; repository (node, real SQLite +
  fake store) — file & row created together, insert-failure deletes the file (no
  orphan), delete removes row + file, `fileMissing` flag when the file is gone, and
  `sweepOrphans` removing an orphan file.

## Filesystem + database as one transaction (the standing principle)

- **Write order encodes the invariant.** Create (file → row) and delete (row → file)
  are ordered so the only crash-window residue is an *orphan file* — recoverable —
  never a *dangling row*. The insert's catch actively deletes the file it just wrote,
  so even a failed save leaves nothing behind.
- **The orphan sweep is the reconciliation, not a guess.** On launch it removes files
  with no row and reports (but never deletes) rows whose file is missing — those
  render a placeholder so a lost file is visible, not silently erased.
- **The store is injected, so the lifecycle is provable off-device.** The whole
  file/DB contract is exercised against an in-memory fake store in node tests; the
  device only has to confirm the real expo store behaves like the fake.

## What changed

New: `domain/photos` (+ barrel); `domain/models/media` (`ProgressPhoto`);
`data/photos/{photoStore,expoPhotoStore}`; `data/repositories/photoRepository`;
`features/measurements/{components/AddPhotoSheet,components/PhotoSweeper,hooks/usePhotos,logic/pickPhoto,screens/PhotosScreen,screens/PhotoCompareScreen}`;
routes `app/(tabs)/measurements/photos/{index,compare}`. Migration 0008
(`progress_photos`). Modified: `changeBus` (`'photos'` table), `app/_layout.tsx`
(store wiring + sweeper), `MeasurementsScreen` (Progress-photos entry). Added deps:
`expo-file-system`, `expo-image-picker` (both `~57.0.0`). No frozen document changed.

## Screens affected

Measurements home (Progress-photos entry), Photos timeline (new), Photo compare (new).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| `oldestMissingAngle` — never-captured first, else oldest-most-recent | domain test | ✅ |
| `groupPhotosByDate` — newest date first | domain test | ✅ |
| Save writes file **and** row together | repo test (real DB + fake store) | ✅ |
| Insert failure deletes the just-written file (no orphan) | repo test | ✅ |
| Delete removes both row and file | repo test | ✅ |
| `fileMissing` flag set when the file is gone | repo test | ✅ |
| `sweepOrphans` removes a file with no row | repo test | ✅ |
| `npm run check` | typecheck + lint + format + 218 tests + db:check (15 tables) | ✅ green (3× stable) |
| On-device photos walk (capture/library, oldest-missing default, timeline, delete, kill-mid-save → sweep, compare side-by-side + Before/After, missing-file placeholder, both themes) | — | ⚠️ consolidated into TD-001 |

## Known limitations

1. **No web screenshots** — the timeline and compare are filesystem + DB backed and
   need real image URIs and native pickers. The file/DB lifecycle is proven by the
   repository tests against a real SQLite DB and a fake store; visuals fold into
   TD-001 (checklist extended with a full photos walk).
2. **No image compression/thumbnailing beyond the picker's `quality: 0.85`** — a
   single-user gallery at personal scale doesn't warrant a resize pipeline; the
   originals live in the app's document dir and are removed on delete. A downscale
   pass can come with a later storage/backup phase if the library grows large.
3. **Delete is an immediate Dialog-confirmed action, not swipe-to-reveal + Undo** —
   consistent with the still-open TD-003 gesture debt; the file removal is
   irreversible, so a confirm Dialog (not a 5 s Undo) is the correct guard here.

## Technical debt

None introduced. The prior deferrals stand unchanged and on schedule (TD-001 device
checklist extended with the photos walk; TD-003 still covers the row-swipe gesture
pass; TD-008/TD-009 untouched).

## Retrospective

**What went well?** The store abstraction paid for itself immediately: making
`PhotoStore` an injected interface with an in-memory fake meant the entire
"single logical transaction" invariant — including the failure paths (insert throws →
file deleted; orphan on disk → swept) — is provable in node tests with zero device or
expo dependency. The ordering rule (file→row on create, row→file on delete) turned a
fuzzy "keep them consistent" goal into two deterministic write sequences whose only
crash residue is the recoverable one.

**What was harder than expected?** The SDK-57 `expo-file-system` rewrite — the legacy
`writeAsStringAsync` API is gone in favour of `File`/`Directory`/`Paths` classes, so
`expoPhotoStore` had to be written against the new object API (`new File(dir, name)`,
`.copy`, `.delete`, `.exists`, `dir.list()`). Isolating that behind the store
interface kept the churn out of the repository and tests entirely.

**What should change before the next phase?** Nothing structural. Phase 13 closes M2;
Phase 14 (backup / export / import) is the daily-use gate and the natural home for a
photo-inclusive archive — the file/DB pairing built here (metadata rows + their files
under `photos/`) is exactly what a backup must bundle atomically, so the same
"one logical unit" discipline carries straight into it.

## Lessons Learned

- **What surprised you:** how completely an injected filesystem made the invariant
  testable — the kill-between-write-and-insert scenario that sounds like it needs a
  device is just "insert throws → assert the file is gone" against a `Map`-backed
  fake store.
- **What documentation prevented mistakes:** DATABASE §3.6 fixed the exact ordering
  and the orphan-vs-dangling-row asymmetry (orphan files are recoverable, dangling
  rows are never allowed), which is why create and delete run their writes in opposite
  order and the sweep only deletes files, never rows. UI_UX §5.2 fixed the
  oldest-missing capture default.
- **What should be reused:** the injected-store pattern for any external resource the
  DB references (files now; potentially exported blobs later) — it keeps `core`/`data`
  expo-free and makes the resource's lifecycle unit-testable; the "opposite write
  order + startup sweep" recipe for any row-plus-file pairing.
- **What should be avoided:** inserting the row before the file exists (would allow a
  dangling row, the one forbidden state); silently hiding a row whose file vanished
  (the placeholder makes loss visible); importing expo from `core`/`data` (the store
  registry is wired at the composition root instead).
- **Amendment proposals:** none — no frozen-document defect surfaced. No new debt.
