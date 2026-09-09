# PCCM — Phase 9 closure settlement

```text
PHASE 9 — ACCEPTED / CLOSED
No outstanding Phase-9 blocker.
No further Windows run is required for Phase 9.
Phase 10 is not started by this record.
```

Phase 9 added **Model Check**: one sheet that shows an operator what is wrong
with the model as it stands, and what to do about it. It is a presentation
layer. Every fact on it is owned by a module that already existed, and the whole
design question was how to show those facts without becoming a second opinion
about them.

This record settles what was proved, by which evidence, and — deliberately —
which two apparent discrepancies in the Windows report are the contract working
rather than drift.

---

## 1. The three authorities, and they are three different commits

| | Commit | What it is |
|---|---|---|
| **Contract** | `dd082c9` | the Model Check contract, settled before any code |
| **Static implementation lineage** | through `c52e732` | the last commit changing a byte under `pccm/src` or `pccm/spec` |
| **Windows-tested tree** | `8ccadd0` | the tree the accepted Windows run was produced against |

Established from the repository rather than asserted:

```text
$ git log --oneline -1 8ccadd0 -- pccm/src pccm/spec
c52e732 Phase 9 Step 3: the last two owners name their driver too

$ git diff --name-status c52e732 8ccadd0 -- pccm/src pccm/spec
(no output)
```

`8ccadd0` is the tree the last Windows run executed. It is **not** a replacement
for the lineage above and restates none of the earlier phases' results. The two
commits between `c52e732` and `8ccadd0` are runner corrections and their
controls; no production, spec or builder byte moved in that span.

### 1.1 Windows acceptance

| | |
|---|---|
| Runner | `bootstrap/windows/phase9_p1_model_check.ps1` |
| Windows acceptance | `8ccadd0` |
| Prerequisites | 7 / 7 |
| Scenario results | **121 / 121** |
| Verdict | **P9-1 PASS** |
| COM lifecycle | clean — workbook closed, application quit, natural PID exit, no emergency cleanup |

Two earlier attempts at `2f15f02` and before it did not reach the scenarios.
Neither exposed a production defect: the first died on a helper block omitted
when the runner was assembled, the second on a PowerShell array-shape defect in
the runner's own block reader. Both are recorded in their commits; neither
changed production, and neither is acceptance evidence.

---

## 2. What Model Check is, and what it is not

The accepted facts, each of which a control protects:

- it is a **read-only presentation and aggregation layer** over existing
  authorities;
- it creates **no new state machine** — every state word is mirrored from the
  module that owns it;
- it creates **no new refusal** — an ERROR row displays a refusal an owner
  already made;
- **only actionable ERROR and WARNING rows** move Overall Status;
- **INFO is context** and is never counted;
- rows are **deterministically ordered** and **deduplicated** by
  `check_id` + `subject`, never by message text;
- the visible register is a **fixed 100-row window**, and the summary counts the
  **whole logical population** with explicit overflow disclosure when it is
  larger;
- unused slots are **`#N/A`**, never blank and never a numeric zero;
- the Monte Carlo recommendation is read from the **machine-readable owner**,
  `input_contract.yaml → inputs.monte_carlo_iterations.recommended_iterations`;
  no formula, module, builder, projection, runner or test owns the literal
  independently;
- the advisory warns **strictly below 10,000**;
- it does **not alter the calculation fingerprint**;
- it **preserves the existing simulation request fingerprint semantics**;
- **live and persisted `(last evaluated)` facts stay separate**, and a persisted
  row is never presented as the live answer;
- the live refusal detail and the **structured permanent-id Subject** are exposed
  through read-only adapters;
- there is **no prose parsing** of a refusal message anywhere — in VBA, in a
  worksheet formula, in the projection or in the runner;
- **no worksheet-called path can write, persist, consume a nonce or run id, call
  a command endpoint, or mutate workbook or application state.**

---

## 3. What real Excel proved

The fifteen facts the accepted run established, each observed rather than
inferred:

1. the authorised calculation adapter answered **from a worksheet cell**;
2. the Model Check geometry, freeze panes and every projected register header;
3. the summary counts **reconcile exactly** to the displayed logical rows;
4. a valid calculated model with no simulation can be **PASS** — an optional
   output that has not been produced is not a defect;
5. **9,999 warns**;
6. **10,000 does not warn**;
7. **10,001 does not warn**;
8. the advisory itself **refuses nothing**;
9. request drift separates **live STALE** from **persisted CURRENT**;
10. an invalid three-point input produces **exactly one actionable root-cause
    ERROR**, whose **Subject is the offending permanent id** and whose
    **Message is the live current-model refusal reason** — not the persisted
    last attempt;
11. a **post-validation arithmetic refusal** — raised past every validation
    check, by an owner the ordering check never reaches — also carries the
    structured permanent id;
12. simulation `INVALID` context **does not duplicate** the root-cause error;
13. the structural report heartbeat **re-evaluated on an ordinary Calculate**:
    `0 -> 1 -> 0`;
14. workbook recalculation **rewrote no persisted status cell**;
15. the COM lifecycle **closed naturally and cleanly**.

---

## 4. Scenario A — the untouched workbook, reconciled

The Windows run reported the untouched Stage-A workbook as live calculation
`INVALID`, one actionable error, Overall `ERROR`. **That is contract-correct.**

Its required current inputs are blank: the applied Base Year, Start Year and
Duration are written only by `PCCM_ApplyTimeline` and nothing has applied a
timeline, and `inpDiscountRate` is `required: true` with `default: null`. So the
current-model preparation fails, and `spec/calc_contract.yaml` settles what
follows:

> "REFUSED is an ATTEMPT result. It is never a derived status: **an invalid
> model is INVALID whether or not anyone pressed Calculate.**"

`modCalcReport.DeriveStatus` composes the two facts in that order — validity
first, history second — so `NOT CALCULATED` is unreachable while the model is
invalid.

**Three facts, kept apart:**

| | Fact | Where it lives | Untouched Stage-A |
|---|---|---|---|
| 1 | "never calculated" — publication and history | `_Calc!C19`, *"Calculation Status (last evaluated)"*, `initial: "NOT CALCULATED"` | reads `NOT CALCULATED` — INFO, and correct |
| 2 | current model validity | `ResolveModel` ∧ `CheckResolvedModel` | fails |
| 3 | **live** calculation state | `DeriveStatus`, validity tested first | **`INVALID`** |

A persisted `NOT CALCULATED` **does not override** a live current-model
`INVALID`. Both readings are true; they are different facts, and Step-1 §1.1
requires them to stay apart.

The Step-1 scenario labelled *"untouched workbook"* represented a
**valid-but-never-calculated** model, not the blank Stage-A bootstrap fixture.
The Step-1 contract at `dd082c9` is historical and is **not rewritten here**.
The terminology is carried forward as a Phase-10 documentation cleanup item.

---

## 5. Iterations — the advisory boundary, reconciled

The D/G boundary assertion reads `Readings['calculation_state']` — the **live
calculation state**, mirrored from `PCCM_ModelCheckCalculationState`. It asserts
`CURRENT` at 9,999, 10,000 and 10,001.

- Monte Carlo iterations are **absent** from the calculation fingerprint:
  `modCalcReport.BuildFingerprint`'s header is Base Year, Start Year, Duration
  and Discount Rate, plus the per-driver records.
- Monte Carlo iterations are **present** in the simulation request fingerprint:
  `sim_contract.yaml`'s `sim_section` lists `iterations` first in both the AUTO
  and FIXED records, and `modSimReport.CurrentRequestFingerprint` passes
  `package.Iterations` into `SimFpBuildRequestFingerprint`.

**Therefore** changing 10,000 to 9,999 or 10,001 leaves the calculation
`CURRENT` and applies the **normal existing simulation STALE semantics**. The
advisory introduces no special state behaviour: it creates no refusal, no new
state, and no fingerprint effect of its own. Crossing the recommendation
threshold is not treated specially — it is an ordinary request change that
happens also to cross an advisory boundary.

---

## 6. `NOT CALCULATED` coverage

`NOT CALCULATED -> WARNING` remains part of the accepted Phase-9 contract, and
is **statically covered** in `tests/test_phase9_model_check.py`, including the
Step-1 row-A shape — the warning, with the live simulation `INVALID` shown
faithfully as INFO beneath it.

The final Windows runner **did not independently sample** a
valid-but-never-calculated live workbook: reaching that state requires reading
the surface in the window between applying the fixture and pressing Calculate,
and no scenario reads there.

**This is not an open acceptance blocker and requires no additional Windows
run.** The rule is proved; what is absent is one more sampling of a state whose
severity mapping is a one-line mirror already exercised statically.

---

## 7. Earlier authorities are not rewritten

`8ccadd0` is the final Phase-9 Windows-tested tree. It **does not replace**
any earlier historical acceptance commit.

| Phase | Authority | Status |
|---|---|---|
| Phase 7 | implementation `79d4c3e`, evidence `ad78988` | unmoved |
| Phase 8 | P8-1 `35bd6ce`, P8-2 `7ff5dc7`, P8-3 `9e3c141`, P8-Z `bfae0eb` | unmoved |

Phase 9 threaded a structured refusal subject through modules those phases were
accepted against. Not one condition, message, Boolean, constant or arithmetic
expression moved, and that is proved rather than asserted: the plumbing is
removed mechanically and the accepted bytes must come back, to the byte. Every
touched module is declared in the Phase-7 §1.1 table and the Phase-8 production
table. No historical digest moved.

---

## 8. What is closed, and what is not started

```text
P9-1   CLOSED / ACCEPTED     Windows 8ccadd0, 7/7 prerequisites, 121/121 results

PHASE 9 — ACCEPTED / CLOSED
No outstanding Phase-9 blocker.
No further Windows run is required for Phase 9.
```

**Phase 10 has not started.** Nothing in this record begins it.
