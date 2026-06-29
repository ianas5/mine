# Monte Carlo Cost Model

Probabilistic, **time-phased** cost estimation. Instead of one deterministic
"the project will cost X", you give each cost driver and risk a *range* of
uncertainty, run thousands of random scenarios (Monte Carlo), and read off the
expected cost, confidence levels (P50/P80/P90), required contingency, annual
cash flow, and NPV.

> ملخص: نموذج تكلفة احتمالي موزّع على السنوات بمحاكاة مونتي كارلو — تقديرات
> ثلاثية النقاط، سجل مخاطر، تضخم، وخصم/NPV، ومخرجات حسب مستوى الثقة.

Standalone project — unrelated to anything else in the repository.

---

## ⭐ Advanced workbook (primary): `AdvancedMonteCarloCostModel.xlsx`

An app-like Excel workbook. The Monte Carlo engine is built from **live
formulas** — open it and press **F9** to run / re-roll. No macros required.

**Sheets**

| Sheet | Purpose |
|-------|---------|
| **Dashboard** | Executive view: KPIs (base / expected / P50–P90 / contingency / recommended budget / NPV) + 4 charts + a checks warning. |
| **Setup** | Project name, currency, years, iterations, confidence (dropdown), inflation, discount rate, VAT. |
| **Cost Lines** | `tbl_CostLines` — 3-point unit costs, quantities, distribution & include dropdowns. Totals computed. |
| **Cost Profiling** | `tbl_CostProfile` — % of each cost line spent per year (rows must total 100%, auto-highlighted). |
| **Risk Register** | `tbl_RiskRegister` — probability + 3PE impact per risk. |
| **Risk Profiling** | `tbl_RiskProfile` — % of each risk impact per year. |
| **Inflation** | Per-year rate → compounding cumulative factor. |
| **Engine** | The formula Monte Carlo (one row per iteration). |
| **Results** | Percentiles, contingency table, annual cash flow by confidence, NPV by confidence. |
| **Sensitivity** | Tornado — P90−P10 spread of each driver. |
| **Assumptions** | Assumptions log. |
| **Checks** | Live validation (profiles = 100%, Min ≤ ML ≤ Max, valid rates…). Dashboard shows a warning if any fail. |

**How the model works**

- Each iteration samples every **cost line** from its distribution
  (Triangular / PERT / Normal) and spreads it across years by its cost profile.
- Each **risk** is *probability-weighted*: a uniform draw decides whether it
  occurs; only then is its 3PE impact sampled and spread by its risk profile.
- Annual amounts are **inflated** by the cumulative factor, summed to the total
  project cost, and discounted to **NPV**.
- `Contingency = ConfidenceTotal − Base(Most-Likely)`;
  `Recommended Budget = ConfidenceTotal`.

**Regenerate / customize**

```bash
pip install openpyxl
# args:  [iterations] [years] [start_year] [currency] [mode]   env: LINES, RISKS
python3 generate_advanced_model.py 3000 5 2025 SAR flex            # FLEX: up to 30 yrs
python3 generate_advanced_model.py 5000 20 2026 SAR               # FIXED: exactly 20 yrs
LINES=50 RISKS=20 python3 generate_advanced_model.py 5000 10 2025  # 50 cost lines, 20 risks
```

**The shipped file is a FLEX build** (structured for up to 30 years, extra years
hidden). What is *live* vs *structural*:

| Setting | How to change | |
|---------|---------------|--|
| **Currency** | edit the *Currency* cell on Setup | live |
| **Start year** | edit the *Start year* cell — calendar labels follow | live |
| **Cost-line / risk values & profiles** | edit the tables | live |
| **Add a cost line / risk** | fill a blank template row, set *Include = Yes*, add its 100% profile — Cost/Risk Profiling rows are pre-linked, the Engine slot is pre-wired | live, up to the built count |
| **Duration (year count)** | set *Number of years*, run the **`ApplySettings`** macro | macro |
| **More cost lines / risks than built, or more iterations** | regenerate (`LINES=…`, `RISKS=…`, iterations arg) | regenerate |

The simulation size (**iterations** = number of Engine rows) is **fixed when the
file is generated** — the Setup cell just reports the built count. Inflation is
set **per year on the Inflation sheet**, not on Setup.

Edit the `COST_LINES`, `RISKS`, `DISCOUNT`, `INFL_BASE`, etc. at the top of
`generate_advanced_model.py` and re-run.

> Distributions use `NORM.INV` / `BETA.INV` / `PERCENTILE.INC` — needs **Excel
> 2010+** (also works in LibreOffice Calc).

### Optional: real "Run Simulation" button (VBA)

The workbook works without macros. To add a clickable button, validation gating,
and one-click PDF export, import the modules in **`vba/`** — see
[`vba/VBA_SETUP.md`](vba/VBA_SETUP.md). Routines: `RunMonteCarlo`,
`SampleTriangular`, `ValidateInputs`, `RefreshDashboard`, `CalculatePercentiles`,
`ExportReport`.

---

## Also included

- **`index.html`** — an interactive in-browser version (no install). Add cost
  items with distributions, run up to 100k iterations, see histogram / S-curve /
  tornado, and export/import Excel & CSV.
- **`MonteCarloCostModel.xlsx`** + `generate_excel_model.py` — a lightweight,
  single-sheet in-Excel Monte Carlo (no time-phasing/risk register). Good as a
  minimal starting point.

---

## Verification

`generate_advanced_model.py` runs an **independent pure-Python** Monte Carlo
(20k samples) that re-implements the exact model logic and prints reference
numbers, so you can confirm the in-sheet results land in the same ballpark. For
the bundled example data:

| Metric | Reference |
|--------|-----------|
| Base (Most-Likely) | SAR 366,000 |
| Total expected | ≈ SAR 468,000 |
| P50 / P80 / P90 | ≈ 464k / 509k / 532k |
| Contingency @P80 | ≈ SAR 143,000 |
| NPV @P80 | ≈ SAR 405,000 |

This is an MVP — correlations between cost lines, more distributions, and a
saved-snapshot mode are natural next steps.
