# PROJECT_INDEX.md

> **Purpose:** a map of the frozen documentation set. This document is **informational only** — it introduces no decisions, rules, scope, or requirements, and **does not override any document**. If anything here appears to conflict with another document, the other document wins, and the discrepancy is an index defect to fix here.
>
> **Baseline status:** FROZEN (v1 baseline · 2026-07-08). Changes to any governed document require the amendment flow in DEVELOPMENT_WORKFLOW §6.

---

## 1. The Document Set

| # | Document | Location | One-line purpose |
|---|---|---|---|
| 0 | **PRODUCT_PRINCIPLES.md** | root | The constitution: 22 principles (P1–P22) and the decision Tests that govern every choice. |
| 1 | **PROJECT_VISION.md** | root | What the product is: a private, single-user fitness tracker that explains progress. |
| 2 | **FITNESS_DOMAIN.md** | `docs/` | The dictionary and rulebook: every fitness concept, formula, unit, and edge case. |
| 3 | **ARCHITECTURE.md** | `docs/` | The structural spine: layers, boundaries, and each technology's lane. |
| 4 | **DATABASE.md** | `docs/` | Persistence: SQLite/Drizzle schema, migrations, backup/import, repositories. |
| 5 | **ANALYTICS_ENGINE.md** | `docs/` | The brain: derived metrics, trends, insight rules, honesty and interpretation. |
| 6 | **DESIGN_SYSTEM.md** | `docs/` | The visual language: tokens, primitive components, charts, feel. |
| 7 | **UI_UX_GUIDELINES.md** | `docs/` | The experience: flows, tap budgets, navigation, states, interaction standards. |
| 8 | **CODING_STANDARDS.md** | `docs/` | The file-level rules: TypeScript, naming, patterns, testing conventions. |
| 9 | **DEVELOPMENT_WORKFLOW.md** | `docs/` | The process: environment, git, CI, verification, doc governance, releases. |
| 10 | **IMPLEMENTATION_ROADMAP.md** | `docs/` | The execution contract: 23 phases, 5 checkpoints, closure ritual, v1 gate. |

**Living artifacts** (created during execution, not frozen): `docs/TECH_DEBT.md` (debt registry), `docs/journal/` (phase demos, retrospectives, screenshots), `docs/V1_TEST_PLAN.md` (built in Phase 22).

---

## 2. Ownership Map — what each document owns and does not own

| Document | Owns | Does NOT own |
|---|---|---|
| **PRODUCT_PRINCIPLES** | Philosophy, values, the Tests, document precedence. | Any technical or visual specifics. |
| **PROJECT_VISION** | Product scope: modules, target user, what's excluded. | How anything is built or computed. |
| **FITNESS_DOMAIN** | Vocabulary, formulas, units, thresholds, directionality, edge cases. | Storage shape (DATABASE), presentation/wording (ANALYTICS, UI_UX). |
| **ARCHITECTURE** | Layers, dependency rules, folder structure, technology lanes, state ownership, data-flow patterns. | Schema detail (DATABASE), file-level style (CODING_STANDARDS). |
| **DATABASE** | Tables, columns, indexes, migrations, backup format, repository boundaries. | What the data *means* (FITNESS_DOMAIN), where state lives (ARCHITECTURE). |
| **ANALYTICS_ENGINE** | Metric catalog, trend/adherence application, insight rules + wording + prioritization, caching, budgets, the dashboard closed list. | The formulas themselves (FITNESS_DOMAIN), chart appearance (DESIGN_SYSTEM), insight placement (UI_UX). |
| **DESIGN_SYSTEM** | Tokens (color/type/space/motion/haptics), primitive components, chart style, theming, accessibility baseline. | Screen composition and flows (UI_UX). |
| **UI_UX_GUIDELINES** | Navigation behavior, screen inventory, flows + tap budgets, interaction standards, information hierarchy, smart defaults, delight registry. | Component internals (DESIGN_SYSTEM), what metrics say (ANALYTICS). |
| **CODING_STANDARDS** | TypeScript rules, naming, file recipes, patterns, lint/format config, test conventions. | Macro structure (ARCHITECTURE), process (DEVELOPMENT_WORKFLOW). |
| **DEVELOPMENT_WORKFLOW** | Environment, scripts, git/CI, testing strategy, Definition of Done, doc governance, releases, data safety. | What to build (roadmap) or how code reads (CODING_STANDARDS). |
| **IMPLEMENTATION_ROADMAP** | Phase order, dependencies, acceptance, checkpoints, closure ritual, debt policy, v1 gate. | Any product or technical decision — it sequences them only. |

---

## 3. Recommended Reading Order

**Full onboarding (first read):** PRODUCT_PRINCIPLES → PROJECT_VISION → FITNESS_DOMAIN → ARCHITECTURE → DATABASE → ANALYTICS_ENGINE → DESIGN_SYSTEM → UI_UX_GUIDELINES → CODING_STANDARDS → DEVELOPMENT_WORKFLOW → IMPLEMENTATION_ROADMAP. *(Constitution and product first, then meaning, then structure, then data/logic, then experience, then process, then plan — each document assumes the ones before it.)*

**Task-based shortcuts:**
- *Implementing a feature phase:* IMPLEMENTATION_ROADMAP (the phase) → the owning documents it cites → CODING_STANDARDS → DEVELOPMENT_WORKFLOW §5 (Definition of Done).
- *Touching a formula or metric:* FITNESS_DOMAIN first, always; then ANALYTICS_ENGINE.
- *Building UI:* DESIGN_SYSTEM → UI_UX_GUIDELINES (in that order: primitives before composition).
- *Touching the schema:* DATABASE §5 (migrations) before anything else.
- *Unsure whether to build something at all:* PRODUCT_PRINCIPLES → The Tests.

---

## 4. Key Decision Locator

| Looking for… | Go to |
|---|---|
| Product principles & the decision Tests | PRODUCT_PRINCIPLES (P1–P22; Tests section) |
| Scope: modules, target user, exclusions | PROJECT_VISION |
| Units, rounding, dates, null-vs-zero | FITNESS_DOMAIN §2 |
| Working set / warm-up rule | FITNESS_DOMAIN §3.2 |
| Muscle taxonomy, push/pull, upper/lower | FITNESS_DOMAIN §3.3 |
| Load types & effective load | FITNESS_DOMAIN §3.4 (storage: DATABASE §3.4) |
| Volume & e1RM (Epley) formulas | FITNESS_DOMAIN §3.5 |
| PR definitions | FITNESS_DOMAIN §3.7 |
| Consistency, streaks, missed workouts | FITNESS_DOMAIN §3.8 |
| Nutrition tracking & adherence bands | FITNESS_DOMAIN §4 |
| Time-versioned nutrition targets | FITNESS_DOMAIN §4.1 (resolution: DATABASE §3.1) |
| Body measurements & directionality | FITNESS_DOMAIN §5 |
| Trend math, deadbands, insufficient-data | FITNESS_DOMAIN §6 |
| Recomposition signal | FITNESS_DOMAIN §6.5 |
| Time ranges (7/30/90/180/365/all) | FITNESS_DOMAIN §7 |
| Domain edge cases (the test checklist) | FITNESS_DOMAIN §8 |
| Layer rules & folder structure | ARCHITECTURE §4–5 |
| Technology lanes (approved stack) | ARCHITECTURE §3 |
| State ownership (SQLite/MMKV/Zustand/RHF) | ARCHITECTURE §6 |
| Crash-safe workout drafts | ARCHITECTURE §7.1 (table: DATABASE §3.4) |
| Full schema & relationships | DATABASE §3 |
| Migration strategy | DATABASE §5 |
| Backup / export / import | DATABASE §6 |
| Repository boundaries | DATABASE §7 |
| Metric catalog & calculators | ANALYTICS_ENGINE §3, §5 |
| Insight rules, priorities, cooldowns | ANALYTICS_ENGINE §6.2–6.3 |
| Dashboard closed content list | ANALYTICS_ENGINE §6.5 |
| Analytics honesty rules | ANALYTICS_ENGINE §2 |
| Color/type/spacing/motion tokens | DESIGN_SYSTEM §2–4 |
| Primitive component catalog | DESIGN_SYSTEM §6 |
| Chart style rules | DESIGN_SYSTEM §6.1 |
| Navigation model & screen inventory | UI_UX_GUIDELINES §2–3 |
| Tap budgets & core flows | UI_UX_GUIDELINES §4 |
| Focus Mode, smart defaults, keyboard-first, delight registry | UI_UX_GUIDELINES §5 |
| TypeScript & naming rules | CODING_STANDARDS §1–2 |
| Domain constants module rule | CODING_STANDARDS §6.2 |
| Testing conventions | CODING_STANDARDS §10 (strategy: DEVELOPMENT_WORKFLOW §4) |
| Scripts, git, CI | DEVELOPMENT_WORKFLOW §1–3 |
| Definition of Done | DEVELOPMENT_WORKFLOW §5 |
| Doc governance & amendment flow | DEVELOPMENT_WORKFLOW §6 |
| Release & personal-data safety | DEVELOPMENT_WORKFLOW §7 |
| Phase list & contract | IMPLEMENTATION_ROADMAP §2, §5 |
| Checkpoints & closure ritual | IMPLEMENTATION_ROADMAP §2.1, §3 |
| Daily-use gate | IMPLEMENTATION_ROADMAP §4 |
| Tech-debt policy | IMPLEMENTATION_ROADMAP §2 |

---

## 5. Precedence Order

As defined in PRODUCT_PRINCIPLES (restated here for navigation only):

**PRODUCT_PRINCIPLES → PROJECT_VISION → FITNESS_DOMAIN → ARCHITECTURE → DATABASE / ANALYTICS_ENGINE → DESIGN_SYSTEM → UI_UX_GUIDELINES → CODING_STANDARDS → DEVELOPMENT_WORKFLOW → IMPLEMENTATION_ROADMAP.**

Conflicts resolve upward; between principles, the lower number wins. Every AI or human contributor consults the owning document before acting, and amends documents (with approval) before writing code that would contradict them.

---

*PROJECT_INDEX.md is a navigation aid. It is not a source of truth, grants no permissions, and overrides nothing.*
