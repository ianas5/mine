# PCCM Phase 6 — Step 12: transaction failure-path settlement

Independent review of `7c37b72`, while building the Step-13 runtime matrix,
found a **source-cleanable transaction defect inherited from Step 11**. This
record states what the defect was, what it could have caused, and what closes it.

**No Windows. No Excel. No simulation executed.** Step 13 is not authorised and
no Step-13 harness scenario exists.

---

## 1. The defect

Step 11 handled the *"a helper returned False"* half of failure and stopped
there. A `Range.Value2` assignment, a chunked bank write and a verification read
are **COM calls, and COM calls raise**. Nothing in `PublishCandidate` or
`FinalCommit` caught a raised error, so:

### 1.1 A raised candidate write skipped the attempt axis

A COM error in any candidate write, in `WriteIterationBank`, or in
`VerifyCandidateBank` propagated out of `PublishCandidate` and out of
`RunSimulation` to `PCCM_RunSimulation.InvocationFailed` — **without passing
through `RecordFailure`**. The announcement was honest, but no attempt record was
written. The accepted contract says otherwise:

```yaml
refusal_or_failure_after_auto_allocation:
  next_auto_nonce_advanced: true
  active_bank_changed: false
  successful_banks_changed: false
  attempt_metadata_updated: true      # <- not satisfied for a raised write
```

The AUTO nonce was already spent, and the workbook recorded nothing about the
attempt that spent it.

### 1.2 A raised final commit skipped the restoration

`FinalCommit` reached its restore **only** when the assignment returned normally
*and* `SameBlock` returned `False`. A raised assignment, or a raised verification
read, left the procedure before the restore. That contradicts:

```yaml
prior_final_commit_block_captured_before_write: true
final_commit_failure_restores_prior_block: true
final_commit_failure:
  prior_block_restored: true
  active_bank_changed: false
```

The source cannot assume *"an exception means Excel wrote nothing"*. **That
assumption is precisely why the prior block is captured**, and the old code
captured it and then declined to use it in the one case it was captured for.

### 1.3 Both later failpoints fired on the wrong side of their write

```
Phase6CandidateBank   fired in RunSimulation, BEFORE PublishCandidate
Phase6FinalCommit     fired BEFORE the D22:D30 assignment
```

An injection before the write it is named for proves only *"nothing was
written"*. It never reaches the recovery path it exists to exercise. Worse, a
Step-11 test **asserted** the wrong order (`failpoint < write`), so the defect
was pinned in place.

---

## 2. What changed in `modSimReport.bas`

Nothing numerical, nothing about identity, no public API, no bank coordinate, no
Results formula. Only the two transaction procedures and the two failpoint
positions.

### 2.1 `PublishCandidate` — a scoped envelope

```
On Error GoTo CandidateFailed
  build the three blocks
  write snapshot / summary / contingency to the INACTIVE bank
  write the iteration chunks
  [FAILPOINT Phase6CandidateBank]      <- written, not yet verified
  verify the candidate bank
On Error GoTo 0
PublishCandidate = True

CandidateFailed:
  capture Err.Description, disarm, set detail, return False
```

**Nothing is rolled back.** The active bank is not touched, the run id is not
touched, and the half-written inactive bank is left exactly as it is — it has no
semantic standing precisely because the selector still names the other bank. The
handler is scoped: one handler, one exit, no `On Error Resume Next`.

`RunSimulation` already owned `If Not PublishCandidate(...) Then RecordFailure`,
so every candidate-stage infrastructure failure now reaches the attempt axis and
the block records FAILED with the seed mode, the effective seed and the consumed
nonce.

### 2.2 `FinalCommit` — three classes, one restore

```
On Error GoTo CaptureFailed
  previous = D22:D30                    <- A. capture, before anything is written
On Error GoTo 0
BuildCommitBlock

On Error GoTo CommitFailed              <- B. from here, every exit restores
  D22:D30 = block
  [FAILPOINT Phase6FinalCommit]
  If SameBlock(block) Then  success
  cause = "the committed block did not verify"
  GoTo RestorePrevious

CommitFailed:  cause = Err.Description

RestorePrevious:
  On Error GoTo RestoreFailed
    D22:D30 = previous                  <- ONE write back
    If SameBlock(previous) Then  "restored and remains authoritative"
    otherwise                    "could not be restored ... requires recovery"
  RestoreFailed:                 "could not be restored: <err> ... requires recovery"

CaptureFailed:                           <- C. nothing written, nothing to restore
  "no final commit was attempted ... the published bank is unchanged"
```

**All four post-write failure modes reach the same restore**: a raised
assignment, an injected failpoint, a raised verification read, and a plain
verification mismatch. The raised route falls through `CommitFailed:` into
`RestorePrevious:`; the mismatch route jumps to it.

**The capture-failure path deliberately does NOT restore.** There is no captured
block, and writing an unset `Variant` over a live publication would turn a read
failure into a data-loss event.

**A failed restoration never claims the bank survived.** Only the capture handler
says the publication is unchanged, and it is the one case where that is true.

### 2.3 Both failpoints moved, and the wrong test was reversed

```
Phase6AfterNoncePersisted   unchanged
Phase6CandidateBank         now INSIDE PublishCandidate, after the inactive-bank
                            writes, before verification completes
Phase6FinalCommit           now AFTER the D22:D30 assignment, before verification,
                            INSIDE the envelope that restores
```

Exactly the three accepted names. No fourth failpoint.

`test_44` asserted `failpoint < write` for the final commit. **That assertion
protected the defect** and is now reversed: `write < failpoint < verification`,
plus `snapshot write < iteration write < candidate failpoint < verification` and
`FAILPOINT_SIM_CANDIDATE_BANK not in RunSimulation`.

---

## 3. The tests

### Conformance — `tests/test_phase6_sim_report_vba.py`

| Test | What it now guarantees |
|---|---|
| `test_31`, `test_44` | both failpoints fire after the write they are named for |
| `test_44a` | the candidate envelope exists, covers all five fallible calls, returns False, and rolls nothing back |
| `test_44b` | a candidate failure reaches `RecordFailure`, and the attempt block keeps the consumed seed |
| `test_44c` | the prior block is captured before any write, under its own handler, and a capture failure does **not** enter the restore |
| `test_44d` | the envelope covers write, failpoint and verification, stays armed through the verification read, and both routes land on one verified restore |
| `test_44e` | a failed restoration says "cannot be guaranteed / requires recovery" and never says "remains authoritative" |
| `test_44f` | a commit failure reaches `RecordFailure`, and the attempt block never touches D22 or D30 |
| `test_44g` | the exact handler set — no `On Error Resume Next`, every named handler has a label |

### Integration — `tests/test_phase6_integration_source.py`

`test_28` walks all five staged calls and requires each arm to record rather than
re-raise, and requires the four scoped handlers to exist. `test_29` maps each
accepted `failure_semantics` clause to its source guarantee, including that
nothing decrements the AUTO nonce on any path.

### Controls

23 new non-vacuous mutation controls (`test_56`–`test_78` in the Step-11
validation suite, `test_38`–`test_46` in the integration validation suite) cover
every item in the settlement list: envelope removed, blanket suppressor,
handler reports success, handler erases the bank, each failpoint moved to each
wrong position, raised-write exits without restore, mismatch exits without
restore, verification read outside the envelope, block never captured, capture
without a handler, capture path writing an unset block, restore write removed,
restore never verified, failed restoration claiming safety, restore without a
handler, both failures bypassing the attempt record, nonce rolled back, and the
attempt block reaching the publication rows.

**Four pre-existing Step-11 controls were re-anchored, not weakened** — their
anchors moved when the procedures were rewritten — and `test_44d` was
**strengthened** after `test_68` survived it: disarming the envelope between the
assignment and the verification is now refused.

---

## 4. What did not change

`modSimRng`, `modSimSample`, `modSimEngine`, `modSimStats`, `modSimFingerprint`,
`modCalcFingerprint` and `modCalcReport` are byte-identical. No contract moved —
`sim_contract.yaml` was already correct and the implementation returned to it.
No registry movement, no D6-11 movement, no `builder/**`, `evidence/**`,
`bootstrap/windows/**` or `workbook.yaml` change.

The Step-12 closures all remain closed: the Gate-B temp-directory cleanup, the
three active manifest-owned inventory descriptions, the corrected Results
presentation text, and Round 4A's `P5-FIN` → `P5-LDG` ordering and exact-type
settlement.

`modSimReport.bas` moved, as expected:

```
was  a0b9a738b8f7346efd7f5964c311861d975075786072e2ec7b7c7773afd0c363
now  de6827dfa2d6d8d20d68fdd2c03a99a103c52bb2a2f1dcfddccee564bae30e1f
```

---

## 5. Still runtime-only

This settlement makes the failure paths *reachable and correct in source*. It
does not execute them. The Step-13 matrix gains the injections that now mean
something:

- `Phase6CandidateBank` — a written but unverified inactive bank, active bank
  unmoved, attempt FAILED, candidate with no standing;
- `Phase6FinalCommit` — a D22:D30 assignment that may have landed, followed by
  restoration of the captured block and an unchanged active bank;
- a real COM failure in each of the five candidate writes;
- a real COM failure on the commit assignment and on the verification read;
- a restoration that itself fails, and the recovery-required detail it produces.

None of these ran. They are Step-13 evidence.
