# DEVELOPMENT_WORKFLOW.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Owner concern:** process — environment, scripts, git, CI, testing strategy, verification, releases, and documentation governance · **Depends on:** every other document; it is the wrapper that keeps them authoritative.
>
> The team is one person plus AI assistants. The process below is deliberately lightweight everywhere except the three places a solo project actually dies: **unverified changes, migration accidents, and documentation drift.** Those get real rigor.

---

## 1. Environment & Tooling

- **Prerequisites:** Node LTS, npm (the package manager — no yarn/pnpm/bun mixing), Xcode + an iPhone for the primary device target, Android tooling as secondary.
- **Development builds, not Expo Go:** `react-native-mmkv` and Victory Native XL (Skia) require native modules, so the project runs as an **Expo development build** (`expo-dev-client`) from day one — built locally (`npx expo run:ios`) or via EAS. Expo Go is not a supported runtime for this app; discovering that mid-project is a known trap, so it's stated here.
- **Canonical scripts** (`package.json`, names fixed):

| Script | Does |
|---|---|
| `npm run dev` | start the dev server (`expo start --dev-client`) |
| `npm run ios` / `android` | build & run the dev build on device/simulator |
| `npm run typecheck` | `tsc --noEmit` |
| `npm run lint` | ESLint, `--max-warnings 0` |
| `npm run format` / `format:check` | Prettier write / verify |
| `npm run test` / `test:watch` | Jest (`jest-expo`) |
| `npm run db:generate` | `drizzle-kit generate` (new migration from schema change) |
| `npm run check` | typecheck + lint + format:check + test — the full local gate |

---

## 2. Git Workflow

- **Trunk-based, short-lived branches.** `main` is always releasable (migrations applied cleanly, `npm run check` green). Work happens on branches named `feat/…`, `fix/…`, `refactor/…`, `docs/…`, `chore/…` — one concern per branch, merged within days, not weeks.
- **Conventional Commits** (`feat:`, `fix:`, `refactor:`, `docs:`, `test:`, `chore:`), imperative mood, body explains *why* when non-obvious. Squash-merge to `main` so history reads as one clean change per PR.
- **Everything lands via PR — even solo.** The PR is where the Definition of Done (§5) is checked and where the "leave it better" rule (CODING_STANDARDS rule 11) is applied. Direct commits to `main` are allowed only for trivial docs typos.
- **No long-lived divergence:** if a branch outlives a week, split it or land it behind incompleteness (a hidden screen), never rebase-juggle for weeks.

---

## 3. CI (GitHub Actions)

Every PR runs, in order: **typecheck → lint (zero warnings) → format check → tests → migration check** (fresh DB: all migrations apply cleanly from zero; plus the fixture upgrade test, §4.3). All must pass to merge; there are no optional checks. CI runs on Node (domain, data, component tests) — no device farm; on-device verification is a manual gate (§5), honestly labeled as such.

---

## 4. Testing Strategy

What deserves tests, in priority order — mirroring where bugs cost the most:

### 4.1 Domain (exhaustive — the app's brain)

- Every function in `domain/fitness` and every calculator in `domain/analytics` is unit-tested.
- **The FITNESS_DOMAIN §8 edge-case list is the literal test checklist** — each numbered edge case appears as a named test (warm-up exclusion, `single_doubled` vs `per_side`, null-vs-zero, deadbands, recomposition guard, e1RM rep cap, insufficient-data floors…).
- Insight rules: each ANALYTICS §6.2 rule gets trigger, non-trigger, boundary, and cooldown tests.
- Threshold constants are imported from `domain/fitness/constants.ts` in tests too — a test hardcoding `1.0` instead of `RECOMP_WEIGHT_STABLE_KG` is wrong even when green.

### 4.2 Data layer (real SQLite)

- Repositories run against a **real SQLite database in Node** (Drizzle over `better-sqlite3`, same schema/migrations as the device runtime) — not mocks. Covered: merge-upsert semantics, target resolution, transactionality (a failing write leaves no partial tree), archive-vs-delete policies, change-bus emission.
- Backup: export→import round-trip equality; import validation rejection; safety-export failure path.

### 4.3 Migrations

- Every migration PR proves: (a) clean apply from empty; (b) apply over a **fixture DB of the previous version with realistic data**, then assert data survived. The fixture is regenerated at each release tag.

### 4.4 UI (selective, high-value)

- RNTL component tests for `core/ui` primitives (states, a11y labels) and the three critical flows' screen logic (log set, log meal, log weight).
- No snapshot-test farms; a failing snapshot nobody reads is noise, not coverage.

### 4.5 Manual, on-device (the honest layer)

Automated coverage ends at the simulator's edge. Each feature ships with a **manual test checklist** in its PR description (the roadmap's per-phase checklists seed these): the real flows on a real iPhone — including tap budgets with a finger, not a mouse.

---

## 5. Definition of Done (per PR)

A change is done when:

1. `npm run check` green locally and in CI.
2. **Docs consistent** — the change contradicts no accepted document; if it needed a doc change, that change was approved *first* (CODING_STANDARDS rule 8).
3. Domain/data changes carry their tests (§4.1–4.3); formulas cite their doc section (CODING_STANDARDS §7.2).
4. **Manually verified on device:** the affected flow exercised end-to-end; **both themes**; empty / insufficient-data / error states seen (UI_UX rule 5); tap budgets respected where touched (UI_UX rule 11).
5. Migration PRs additionally: fixture upgrade test passed, and a **fresh backup of the personal device's data exists before the build is installed on it** (§7).
6. One small "leave it better" improvement included where the diff touched adjacent mess (CODING_STANDARDS rule 11).

---

## 6. Documentation Governance

The documentation set is **the single source of truth**; code that contradicts it is wrong even if it works (P22).

- **Precedence** (PRODUCT_PRINCIPLES, restated as the operating rule): PRODUCT_PRINCIPLES → PROJECT_VISION → FITNESS_DOMAIN → ARCHITECTURE → DATABASE / ANALYTICS_ENGINE → DESIGN_SYSTEM → UI_UX_GUIDELINES → CODING_STANDARDS → DEVELOPMENT_WORKFLOW → IMPLEMENTATION_ROADMAP. Conflicts resolve upward; between principles, the lower number wins.
- **Amendment flow:** need identified → **stop implementation** → propose the amendment (what changes, why, which sections) → owner (you) approves → doc updated in a `docs:` commit → implementation proceeds referencing it. Amendments are edits-in-place plus a dated entry in a short *Changelog* section at the bottom of the amended document — history lives in git, the changelog is the human index.
- **Docs live in the repo** (`docs/` + root `PROJECT_INDEX.md`, `PRODUCT_PRINCIPLES.md`, `PROJECT_VISION.md`), versioned with the code they govern, PR-reviewed like code.
- **Drift audits:** any discovered code-vs-doc contradiction is filed and resolved in the *next* PR touching that area — by fixing the code, or by an approved amendment. Silent drift is the failure mode this whole section exists to prevent.

---

## 7. Releases & Personal-Data Safety

- **Versioning:** semver-ish `MAJOR.MINOR.PATCH` (MAJOR = backup-format or schema-rebuild events, MINOR = features, PATCH = fixes), tagged in git with a one-paragraph changelog. Phase tags per IMPLEMENTATION_ROADMAP §2.1.
- **Distribution:** EAS build → personal device (internal TestFlight or direct dev-build install). No store release; no OTA update pressure — updates install when *you* choose.
- **The prime directive of releases (P20):** the personal device's database is the only copy of years of history that matters. Therefore: **export a backup before installing any build containing a migration** (§5.5 makes it a gate); after install, open the app, confirm data intact, spot-check one screen per pillar.
- **Dev data separation:** development and testing run against seeded/simulator data; destructive flows (import-replace, deletions) are **never** first tested on the personal device. The seed profile includes a realistic 6-month synthetic history so analytics and performance are exercised honestly (P4: designed for year five, tested with more than three rows).

---

## 8. Working with AI Assistants

AI-authored code follows every document with no special lane: it cites principles for product judgments (PRODUCT_PRINCIPLES rule 1), runs the Tests for new features, stops for doc amendments (CODING_STANDARDS rule 8), and its PRs meet the same Definition of Done. AI may draft doc amendments but **never self-approves them** — acceptance is always yours. Session summaries or scratch notes never substitute for updating the real documents.

---

## 9. AI Decision Rules (Workflow)

1. **Never merge red or warned:** CI green + zero warnings is the floor, not the goal; never bypass, skip, or `--no-verify` the gates.
2. **Docs before code, approval before docs-change** (CODING_STANDARDS rule 8 is a workflow law, restated because it's the one that saves this project).
3. **A migration PR without a fixture-upgrade test and a pre-install personal backup does not merge.** No exceptions — this is the app's only unrecoverable failure class.
4. **Claimed = verified:** report a flow as working only after exercising it (test or device); "it should work" is not a status. Failed checks are reported with their output, not summarized away.
5. **One concern per branch/PR;** refactors ride separately from features unless inseparable.
6. **The FITNESS_DOMAIN edge-case list is the domain test checklist** — a new edge case added there requires its test in the same change.
7. **Manual checklists are deliverables:** every feature PR includes its on-device checklist and its outcome, honestly filled.
8. **When the process itself blocks good work, amend this document** — don't quietly route around it.

---

## Changelog

- 2026-07-08 — v1 baseline frozen.
