# PCCM — Phase 8 closure settlement

```text
PHASE 8 — ACCEPTED / CLOSED
No outstanding Phase-8 blocker.
Phase 9 is not started by this record.
```

Phase 8 added the **presentation layer**: the Results surface and its annual
cash flow and reconciliation, the Dashboard executive summary, and the
analytical charts. Each has its own Windows/Excel runtime evidence, produced by
its own dedicated minimal runner on the operator's machine against a real Excel
and a real Stage-B workbook.

This record settles what was proved, by which evidence, and — just as
deliberately — what is supplemental rather than primary.

---

## 1. Three step authorities, and they are three different commits

A single quoted commit would let a reader believe one run established the whole
phase. It did not: three steps were accepted separately, each against the tree
that existed when its evidence was produced.

| Step | Windows acceptance | Checks | Verdict |
|---|---|---|---|
| **P8-1** — Results, annual cash flow, reconciliation | `35bd6ce` | 160 / 160 scenario checks | **CLOSED / ACCEPTED** |
| **P8-2** — Dashboard executive summary | `7ff5dc7` | 124 / 124 scenario checks | **CLOSED / ACCEPTED** |
| **P8-3** — Analytical charts | `9e3c141` | 16 / 16 prerequisites, **189 / 189** scenario results | **CLOSED / ACCEPTED** |

Each step's production lineage, established from the repository rather than
asserted — the last commit changing a byte under `pccm/src` or `pccm/spec` at
the point each acceptance was produced:

```text
$ git log --oneline -1 35bd6ce -- pccm/src pccm/spec
2e72ddf P8-1: the Results state path is read-only

$ git log --oneline -1 7ff5dc7 -- pccm/src pccm/spec
a7b2669 P8-2: the Dashboard executive summary, mirrored from Results

$ git log --oneline -1 9e3c141 -- pccm/src pccm/spec
0cfa10f P8-3 Windows run 4: a risk publishes its risk name
```

**The accepted P8-3 main Windows evidence remains the 189/189 run at
`9e3c141`.** Nothing later re-establishes it and nothing later replaces it.

### 1.1 What P8-3 proved in real Excel

All seven parts — 0, A1, A2, B, C, D and E. Four real Excel chart objects, read
out of `ChartObjects` and `SeriesCollection` rather than inferred from the file
the builder wrote: their types, their geometry, their series and category ranges,
and the bridge payload behind every plotted point. No chart source reads a
machine sheet. The state vocabulary is qualified live and untranslated —
CURRENT, OTHER Px, STALE, HISTORICAL and INVALID each observed in the transition
that produces it. Risk driver names are text. COM lifecycle clean.

---

## 2. Phase-7 authorities are not rewritten

Phase 7's two authorities stand exactly where its own closure record put them:

| Authority | Commit |
|---|---|
| Production / spec implementation | `79d4c3e` |
| Windows acceptance evidence | `ad78988` |

Phase 8 found real defects in modules Phase 7 was accepted against, and
corrected them. **That does not reopen Phase 7 and does not move either
authority.** What it does is oblige this phase to say so: every such correction
is named in the Phase-7 record's §1.1 table and in the Phase-8
declared-production-corrections table, and a control refuses an undeclared one.

---

## 3. Later Phase-8 integration corrections — declared, not denied

Four corrections were found by Phase-8 Windows acceptance after the layer they
belong to was written. Each is recorded here with what it changed and why.

| Correction | Commit | What changed |
|---|---|---|
| **Results read-only state integration** | `2e72ddf` | `PCCM_SimulationStatus` derives *and persists* `_SimData!D28:D29`; Excel forbids a cell-called function to write, so both Results state cells showed `#VALUE!` live. The derivation was already pure — `SimReportDerivedStatus` exposes it read-only, and `modSimAnnualStore` splits its command and read paths. No state rule moved. |
| **Risk-driver label ownership** | `0cfa10f` | `modSimPostReport.DriverNameOf` read `COL_RISK_REGISTER_DESCRIPTION` for a risk — the column `driver_contract.yaml` declares `required: false`. It now reads `COL_RISK_REGISTER_RISK_NAME`, which that contract declares `required: true`. **One line.** The cost-line branch is untouched and was always right. No driver id, rank, signed rho, absolute rho, direction, status, ordering, fingerprint or replay mathematics is touched. |
| **Zero-variance Sensitivity presentation** | `94c6b37` | The sheet read each published field through a *bare reference*, and Excel reads an empty reference back as **zero** — so rho, \|rho\|, rank and direction, all deliberately blank for a zero-variance driver, rendered as measured zeroes. `sim_contract` says `reported_as_zero_rho: false`. The reference is now guarded with the same idiom the Dashboard mirror already used. A genuinely measured zero still shows. |
| **Zero-variance tornado exclusion** | `6781ea7` | The bridge took the first N rows *of the sheet*, which are the ranked rows followed by the diagnostic ones — so below N eligible drivers a zero-variance driver entered the category window. `sim_contract` says `excluded_from_tornado_input: true`, a clause separate from the two beside it and therefore not a restatement of them. Every bridge field is now gated on **rank**, the field `excluded_from_ranking` makes authoritative. Order is preserved positionally; nothing is re-sorted, compacted or re-ranked. |

The first two are `pccm/src` changes and carry the declaration discipline that
applies there. The last two are in `pccm/builder` — the generated presentation —
and change no VBA and no contract.

**Whole clean-tree sweeps after the shared-production corrections**, with
pytest's own exit code captured:

```text
after the risk-driver label correction      5044 passed   PYTEST_EXIT=0
after the zero-variance presentation fix    5062 passed   PYTEST_EXIT=0
```

---

## 4. `bfae0eb` — the final Windows-tested tree, not a replacement

`bfae0eb` is the tree the last Windows run was produced against. It is **not** a
replacement for the evidence commits above and does not restate their results.

Between `9e3c141` and `bfae0eb`, **no byte under `pccm/src` or `pccm/spec`
moved**:

```text
$ git diff --name-status 9e3c141 bfae0eb -- pccm/src pccm/spec
(no output)
```

The two zero-variance corrections in that span are in `pccm/builder`, and the
rest is runners, controls and this record.

---

## 5. P8-Z — a narrow supplemental closure proof

The accepted P8-3 fixture has five drivers and all five vary, so it could never
produce a diagnostic row. The one question the zero-variance settlement turned
on therefore had no live evidence, and needed a model P8-3 does not build.

`bootstrap/windows/phase8_pz_zero_variance.ps1` asks that question and nothing
else. It applies the accepted W4 model **and request**, then gives one cost line
an identical minimum, most likely and maximum, so the *accepted kernel* reaches
the zero-variance state on its own — nothing fabricates a status.

| | |
|---|---|
| Windows acceptance | `bfae0eb` |
| Prerequisites | 7 / 7 |
| Scenario results | **23 / 23** |
| Verdict | **P8-Z PASS** |

Proved in real Excel:

- the accepted kernel classified exactly one of five drivers zero-variance —
  `4 ranked of 5 drivers`;
- the diagnostic row is **retained** on Sensitivity;
- its status is `n/a - no variance`;
- its rho, \|rho\|, rank and direction display **blank, not zero**;
- the zero-variance driver is absent from the tornado **categories**;
- it is absent from the tornado **bars**;
- every eligible ranked driver is drawn **exactly once**;
- categories = bars = the eligible ranked count;
- the fixed Top-10 source range's no-data slots **fabricate no categories** —
  ten slots, four carrying identities;
- COM lifecycle clean.

**This is supplemental, and narrow on purpose.** It is not a rerun of P8-3, it
re-tests none of that suite's surface, and its 23 checks neither extend nor
replace the 189. P8-3's acceptance stands on its own run at `9e3c141`.

---

## 6. What is closed, and what is not started

```text
P8-1   CLOSED / ACCEPTED     Windows 35bd6ce, 160/160
P8-2   CLOSED / ACCEPTED     Windows 7ff5dc7, 124/124
P8-3   CLOSED / ACCEPTED     Windows 9e3c141, 16/16 prerequisites, 189/189 results
P8-Z   SUPPLEMENTAL PROOF    Windows bfae0eb, 7/7 prerequisites, 23/23 results

PHASE 8 — ACCEPTED / CLOSED
No outstanding Phase-8 blocker.
```

**Phase 9 has not started.** Nothing in this record begins it, and no Phase-9
work exists in the tree.
