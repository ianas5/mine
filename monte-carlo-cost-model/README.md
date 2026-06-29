# Monte Carlo Cost Model

A probabilistic cost-estimation tool. Instead of a single "the project will cost
X" number, you give each cost driver a **range of uncertainty**, run thousands of
random scenarios (a Monte Carlo simulation), and get a full picture of the
likely outcomes: expected cost, confidence levels (P50/P80/P90), the probability
of staying within budget, and how much contingency you actually need.

> ملخص: أداة لحساب تكلفة المشاريع باستخدام محاكاة مونتي كارلو. تعرّف لكل بند
> توزيعًا احتماليًا، تشغّل آلاف المحاكاات، وتحصل على المتوسط، نسب الثقة
> (P50/P80/P90)، احتمال البقاء ضمن الميزانية، والاحتياطي المطلوب.

This project is **standalone** and unrelated to anything else in the repository.

---

## What's inside

| File | What it is |
|------|------------|
| **`index.html`** | The interactive web app. Open it in any browser — no install, no server, runs 100% locally. |
| **`MonteCarloCostModel.xlsx`** | A self-contained Excel workbook where the simulation runs **inside the spreadsheet** with live `RAND()` formulas. Press **F9** to re-roll. |
| **`generate_excel_model.py`** | The Python (openpyxl) script that builds the `.xlsx`. Re-run it to regenerate after editing the model. |

You get the same model in two forms: a polished interactive app **and** a real,
editable Excel file.

---

## The web app (`index.html`)

Just open the file. Features:

- **Flexible line items** — add any number of cost drivers, each with its own
  probability distribution.
- **6 distributions** — Fixed, Uniform, Triangular, **PERT**, Normal, Lognormal.
- **Up to 100,000 iterations**, seeded for reproducible results.
- **Results**
  - Summary: Mean, Median (P50), Std-dev, P80, observed range.
  - **Budget & contingency**: probability of finishing under your budget, and the
    recommended budget for any confidence level (e.g. P80).
  - **Histogram** of the total-cost distribution.
  - **S-curve** (cumulative probability) with your budget marked.
  - **Tornado / sensitivity chart** — which items drive the uncertainty.
- **Excel ↔ app** — export full results to `.xlsx`, or import line items from an
  Excel/CSV file (columns: `Name, Distribution, P1, P2, P3`).
- Save/load scenarios locally, dark/light theme.

### Distribution parameters

| Distribution | P1 | P2 | P3 |
|---|---|---|---|
| Fixed | value | — | — |
| Uniform | min | max | — |
| Triangular | min | most likely | max |
| PERT | min | most likely | max |
| Normal | mean | std dev | — |
| Lognormal | mean | std dev | — |

---

## The Excel model (`MonteCarloCostModel.xlsx`)

A genuine in-sheet Monte Carlo — no macros, no add-ins.

1. Open the **Monte Carlo** sheet.
2. Edit the **yellow** input cells (P1/P2/P3) and the Budget / Confidence.
3. Press **F9** to re-roll all iterations. The summary stats, histogram and
   S-curve update live.

Each iteration is one row in the simulation block (columns to the right). Every
draw uses inverse-transform sampling on a single stored `RAND()`, so the
sampling is statistically correct. Uses `NORM.INV` / `LOGNORM.INV` / `BETA.INV`
/ `PERCENTILE.INC` / `STDEV.S` — **requires Excel 2010+** (also works in
LibreOffice Calc).

### Regenerating / customizing the Excel file

Edit the `ITEMS` list at the top of `generate_excel_model.py`, then:

```bash
pip install openpyxl
python3 generate_excel_model.py            # default 5,000 iterations
python3 generate_excel_model.py 10000      # custom iteration count
```

---

## How Monte Carlo costing works (in one paragraph)

A deterministic estimate adds up your best-guess numbers and gives one total —
which is almost always wrong, because every line item has uncertainty. Monte
Carlo instead **samples** each item from its distribution thousands of times and
sums them, building up the distribution of the *total*. From that distribution
you can read the expected cost, how bad a bad day looks (P90/P95), the chance of
beating your budget, and the contingency required to hit a target confidence.
