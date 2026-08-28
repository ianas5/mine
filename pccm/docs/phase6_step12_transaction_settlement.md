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
Phase6AfterNoncePersisted   now INSIDE AllocateAutoNonce, after the advance has
                            persisted, verified and been marked consumed
                            (see §6 - corrected in a second round)
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


---

## 6. Second round: the after-nonce failpoint carried the same defect

The first round of this settlement corrected the candidate and commit stages and
recorded `Phase6AfterNoncePersisted` as **unchanged**. That was wrong, and
independent review found it while building the Step-13 matrix.

### 6.1 The defect

`modAppState.FailPointCheck` **raises**:

```vb
Err.Raise vbObjectError + 5001, "modAppState.FailPointCheck", _
          "Injected structural failure after stage '" & StageName & "'."
```

and it was called as a **naked statement in `RunSimulation`**, immediately after
`AllocateAutoNonce` returned `True`:

```
AUTO nonce persisted and verified -> AllocateAutoNonce returns True
  -> Phase6AfterNoncePersisted RAISES
  -> PCCM_RunSimulation.InvocationFailed
  -> announcement only, no RecordRefusal, no attempt metadata
```

The nonce was already spent and the attempt axis stayed silent — the exact
contract clause the first round had just enforced for the two later stages:

```yaml
refusal_or_failure_after_auto_allocation:
  next_auto_nonce_advanced: true
  attempt_metadata_updated: true      # <- not satisfied
```

`AllocateAutoNonce` carried the same COM-vs-`False` gap independently of the
injection: the counter write and its read-back are COM calls, and a raised one
escaped the same way.

### 6.2 Why the first round's detector missed it

`test_44` proved only that the failpoint sat between `AllocateAutoNonce` and
`RunKernels` — an adjacency, not a recovery path. Integration `test_28` walked
the five `If Not <stage>` guards and checked their recorders, but a bare
`FailPointCheck` statement is not one of those guards, so the guarantee
*"no transaction stage escapes to the invocation axis"* was **not literally
true** when it was written. That is a detector defect, not just a source one.

### 6.3 The settlement

`AllocateAutoNonce` gains a scoped envelope over the seed derivation, the counter
write, the read-back verification, the consumed mark and the injection. The AUTO
branch became an `If/Else` so both seed modes reach one common exit:

```
On Error GoTo AllocationFailed
  FIXED: EffectiveSeed = SuppliedSeed
  AUTO:  derive seed -> write counter -> read back -> match -> NonceConsumed = True
  [FAILPOINT Phase6AfterNoncePersisted]     <- persisted, verified, marked
On Error GoTo 0
AllocateAutoNonce = True

AllocationFailed:
  capture Err.Description, disarm, set detail, return False
```

The ordering the name promises is preserved and now provable: the injection is
reached **only** after the advance has been written, read back and matched, so
`package.NonceConsumed` already records the spent sequence; and **before** any
sampling, because `RunKernels` runs only if this returns `True`.

`RunSimulation`'s existing `If Not AllocateAutoNonce ... RecordRefusal` arm owns
the result, so the attempt block records the unsuccessful attempt with the seed
mode, the effective seed and the consumed nonce.

**The counter is never rolled back**, on any path. A raised write may have
landed, and reusing the sequence would let a later AUTO run reproduce this one —
the single thing the nonce exists to prevent. The failure detail says so rather
than implying the sequence is free.

No failpoint was added or renamed. There are still exactly three, and each now
fires inside the scoped envelope of the transaction it belongs to:

| Failpoint | Owner | Armed handler |
|---|---|---|
| `Phase6AfterNoncePersisted` | `AllocateAutoNonce` | `AllocationFailed` |
| `Phase6CandidateBank` | `PublishCandidate` | `CandidateFailed` |
| `Phase6FinalCommit` | `FinalCommit` | `CommitFailed` |

### 6.4 The detector, made literally true

- `test_44` now refuses **any** `FailPointCheck` in `RunSimulation`, and pins
  write < verify < mark < injection inside the allocation, with the envelope
  armed first.
- `test_44h`–`test_44k` cover the allocation envelope, the return through the
  attempt path, the no-rollback rule and the exact three-failpoint set with each
  owner arming a handler before firing.
- Integration `test_28` now **reads `modAppState.FailPointCheck` and asserts it
  still raises** — so the guarantee cannot silently become vacuous — then
  requires that no failpoint is naked in `RunSimulation` and that every owner
  arms a handler before firing.
- 12 new mutation controls (`test_79`–`test_90`) and 5 integration controls
  (`test_47`–`test_51`), including the original defect planted back verbatim.

`modSimReport.bas` moved again:

```
first round   de6827dfa2d6d8d20d68fdd2c03a99a103c52bb2a2f1dcfddccee564bae30e1f
this round    a48bddffc5c512ed30a0ab78c2cd802fea57031bb1ebbf5b09f0fed1a394f60b
```


---

## 7. Third round: a post-write verification failure lost the allocated identity

Independent review of `e574fdb` found that the after-nonce **injection** was
settled but the ordinary AUTO **verification-failure** paths were not.

### 7.1 The defect

One Boolean, `NonceConsumed`, was doing two jobs: gating sampling *and* gating
the audit identity on the attempt record. It was set only after a verified
read-back, so all three post-write failure paths —

| Path | What happened |
|---|---|
| A | counter assignment returned, verification read **raised** |
| B | counter assignment returned, `ReadMachineLong` returned **False** |
| C | read-back succeeded but **mismatched** |

— exited with `NonceConsumed` still `False`. `WriteAttemptBlock` then wrote
**blank** for both `Last Attempt Effective Seed` and `Last Attempt AUTO Nonce`,
while the persisted counter may already have advanced. That is the audit hole
the retained authority names in as many words: *"otherwise a consumed nonce
would leave no trace and the sequence would appear to skip."*

It contradicts `seeding.nonce_lifecycle`:

```yaml
failure_after_allocation_consumes_nonce: true
attempt_metadata_preserves: ["consumed_auto_nonce", "effective_seed"]
```

### 7.2 Why the detector pinned it

`test_44` required `write < verify < mark < injection`, which is correct for the
*injected* case — but it made "every verification failure happens before the
mark" a proved property, and nothing then said what the attempt record contained
on those paths. `test_44i` only proved `WriteAttemptBlock` **mentions**
`EffectiveSeed` and `ConsumedNonce`, never that the Boolean gates populate them.
Integration `test_29` checked the publication clause and the absence of
rollback, but never mapped `attempt_metadata_preserves` to the post-write path.
Presence, not path semantics — a literal detector gap, and mine.

### 7.3 The settlement: two facts, two fields

`ALLOCATION` and `CONSUMPTION` are different claims and now have different
fields:

| Field | Meaning | Set |
|---|---|---|
| `NonceAllocated` | this run has **claimed** the nonce | immediately **before** the counter write is attempted |
| `NonceConsumed` | the advance **verifiably persisted** | only after the read-back succeeded **and** matched |

- **Sampling** is gated on `NonceConsumed` — unchanged. Nothing samples on an
  unverified advance.
- **The attempt record** is gated on `NonceAllocated`, which is what
  `attempt_metadata_preserves` is about. A refusal *before* allocation still
  blanks both, satisfying `failure_before_allocation_consumes_nonce: false`.
- **Published records** (`BuildSnapshotBlock`, `BuildCommitBlock`) keep
  `NonceConsumed`: a published bank may claim only a **proven** consumed nonce.
  A control refuses each of them being switched to the weaker flag.

### 7.4 The ambiguous raised write, stated honestly

The handler no longer says one thing while implying another. It branches:

```
allocated    -> names the nonce, says persistence is INDETERMINATE
                ("a raised write is not proof that nothing was written"),
                NOT rolled back, will not be reused
not allocated-> "no AUTO nonce was allocated"
```

And the two verification-failure details no longer say the advance *"did not
persist"* — they say it was written and could not be read back, or did not read
back as written, and that the nonce is recorded as allocated. Controls
(`test_96`, `test_97`) refuse a regression to the stronger claim.

**No contract contradiction was found.** `NonceAllocated` is a run-local field,
not a new persistent recovery state: the contract's
`failure_after_allocation_consumes_nonce: true` plus `reuse_permitted: false`
already require exactly the conservative treatment implemented here, and
`attempt_metadata_preserves` already requires the identity on the attempt. No
spec change was needed or made.

### 7.5 The routing guarantee, made precise

Integration `test_28` is renamed
`test_28_every_transaction_failure_is_routed_to_the_attempt_recorder` and now
states its own limit: it proves failures **reach** the attempt writer, and
explicitly **does not** claim the attempt row can never fail to be **stored**.
`WriteAttemptBlock` ends in a single unguarded COM write; if that raises, the run
leaves through the invocation axis with no attempt row. That is a distinct
storage failure of the audit writer, the contract requires no recovery state for
it, and settling it would be scope creep — so it is enumerated for Gate B
instead, and an assertion refuses a handler appearing there silently.

### 7.6 New coverage

`test_44h2` (two facts, two fields, FIXED touches neither), `test_44h3` (the
contract clause mapped field by field, in both directions), `test_44h4` (the two
post-write exits live in the allocated-but-unconsumed window, name the nonce,
refuse reuse, and claim nothing stronger), plus 12 controls `test_91`–`test_102`
and 3 integration controls `test_52`–`test_54`. `test_91` and `test_52` restore
the `e574fdb` single-Boolean shape verbatim and require refusal.

```
second round  a48bddffc5c512ed30a0ab78c2cd802fea57031bb1ebbf5b09f0fed1a394f60b
this round    0797e307a4a69c6847cb07415d01abf7bc584539a831ce6f124bb8e74d3af1f4
```

### 7.7 A size limit met rather than raised

Three rounds of settlement commentary pushed `modSimReport.bas` to 1232 raw
lines, over the 1200-line raw ceiling in
`test_phase4_stage_b_source.py::test_05`. Its **code** count was 854 against a
900 limit — the overage was documentation, not sprawl.

The limit was **not** raised. That file is outside this round's authorised
boundary, and raising a ceiling to fit one's own code is how a control stops
being one. Instead the prose was compressed where three rounds had said the same
thing three times — the historical narration each block carried is what this
document is for — with every operative rule, every string literal and every
asserted phrase preserved. Result: **1190 raw, 854 code**, 10 lines of headroom.

Worth flagging for the record: that raw ceiling was calibrated to the
then-largest module rather than derived from a principle, and three modules now
sit within 15 lines of it. Whether it should be re-derived is a decision for
review, not something to settle inside a failure-path correction.
