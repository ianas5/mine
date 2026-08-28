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


---

## 8. Fourth round: the allocation authority itself was wrong

Round three settled the *visibility* of a failed allocation by inventing a
pre-write `NonceAllocated` flag. Independent review established that this
**contradicted the locked authority** — Step 0 defines allocation as
*read → derive → **persist***, so "allocated" cannot be claimed before the
write — and that my own new detectors had pinned the wrong definition in place.
That analysis was accepted, the gap was classified as a contract-model gap, and
this round implements the approved settlement (**Option 3R**, carrier
**Option A**, line-cap route **new module**).

### 8.1 The Phase-6 attempt-result axis gains a fifth token

```
NONE  SUCCESS  REFUSED  FAILED  AUTO_NONCE_INDETERMINATE
```

**Phase 5 is unchanged and remains exactly four values.** `calc_contract.yaml`,
`calc_loader.py` and the Phase-5 attempt-result tests have **no diff**. The
Phase-6 axis is now a **strict superset**: `_Calc!Last Attempt Result` and
`_SimData!Last Attempt Result` carry analogous labels and are no longer required
to share an enumeration, because the fifth token names a persistence/recovery
condition Phase 5 has no analogue for. Like `REFUSED` and `FAILED` it takes no
part in deriving the simulation status.

`docs/phase6_plan.md` carried the old membership in normative prose. It now
carries a **post-acceptance authority correction** stating the new membership and
reaffirming the orthogonality rule, which was always correct.

### 8.2 The builder lock

`builder/pccm_builder/sim_loader.py` was the hard blocker: `LOCKED_ATTEMPT_RESULTS`
rejected the fifth token outright and Stage A would not build. It is now the
five-value sequence — **still an exact-sequence lock**, order and membership both
load-bearing. This is an authority correction, not a relaxation.

A **second, unavoidable** `sim_loader.py` change was required and is reported
rather than buried: the contract is **closed-world**, so the new
`nonce_lifecycle` keys were rejected until the validator learned them — and the
validator's own error text says a key it merely tolerates governs nothing. So
they are not merely permitted, they are **enforced**: `_validate_allocation_states`
pins each of the three states to the exact facts an implementation may claim in
it, pins the one-observation reconciliation table, requires the next-run
reconciliation to activate on the durable token alone, pins
`attempt_metadata_preserves` by state, and requires the declared residual.

### 8.3 The three states, in the contract

| State | `advance_persisted` | `nonce_consumed` | Retry may take the same nonce |
|---|---|---|---|
| `PRE_ALLOCATION` | false | false | **true** — nothing was ever allocated, so this is not reuse |
| `CONSUMED` | true | true | false |
| `PERSISTENCE_INDETERMINATE` | `unknown` | `unknown` | — must not be called allocated **or** unconsumed |

`reuse_permitted: false` is now explicitly scoped to a **known-consumed** nonce.

### 8.4 The new module

`modSimNonce` owns the AUTO nonce transaction and nothing else. The size control
did exactly what a module-size control is for: it exposed that one coherent
responsibility should be separated. **Neither ceiling was raised.**

```
modSimNonce   raw  552   code  321
modSimReport  raw 1179   code  819
ceilings      raw 1200   code  900
```

One public entry point, `SimNonceAllocate`, taking **scalars only**:
`hasSuppliedSeed`, `suppliedSeed` in; `effectiveSeed`, `autoNonce`,
`identityKnown`, `state`, `detail` out. `SimRunPackage` **stays Private** — the
split was not paid for by exporting it — and no shared mutable context object
exists in either direction. The dependency is one-way: `modSimReport` drives
`modSimNonce`, and `modSimNonce` names the reporter nowhere, not even to borrow
a failpoint constant. The constant moved with its owner.

Executable failpoint ownership is now exactly:

```
modSimNonce   Phase6AfterNoncePersisted
modSimReport  Phase6CandidateBank, Phase6FinalCommit
```

Still three, none added or renamed.

### 8.5 Immediate reconciliation

After a raised counter write **or** a failed first verification, the original
error is captured, the handler disarmed, and **exactly one** observation taken —
no retry loop, ever:

| Observed | Classification | What the attempt says |
|---|---|---|
| `m+1` | `CONSUMED` | nonce `m` consumed, will not be reused; run still unsuccessful; no sampling |
| `m` | `PRE_ALLOCATION` | advance did not persist; `m` **not** consumed; *"a retry may take it again"* |
| neither | `RECOVERY_REQUIRED` | inconsistent counter; nothing normalised, decremented or skipped |
| unobtainable | `PERSISTENCE_INDETERMINATE` | *"neither allocated nor unconsumed"*; durable marker written |

The `PRE_ALLOCATION` arm deliberately does **not** say *"will not be reused"* —
a control refuses that phrase there, because `m` will legitimately be reissued.

> **Superseded in part by §9.8.** The *Classification* column above reads as one
> axis, and that is no longer how the source models it. `RECOVERY_REQUIRED` is
> an **action**, not an allocation classification, so the `neither` row is two
> facts: `allocationState = PERSISTENCE_INDETERMINATE` **and**
> `recoveryRequired = True`. The `unobtainable` row's *"durable marker written"*
> is also loose — the F21 marker was established **before** the counter write
> and is **retained** here, not written now. The observation table itself
> (`m+1` / `m` / neither / unobtainable) stands unchanged; only the single-axis
> presentation is withdrawn. **§9.8 is the authority.**

### 8.6 Next-run reconciliation

**Superseded by §9.** The carrier described here — `Last Attempt Result =
AUTO_NONCE_INDETERMINATE` — was rejected in static review and is no longer the
recovery authority. The *resolution table* below survived unchanged and now
reads the sidecar instead: with pending nonce `m` and counter `c`,
`c = m+1` → resolved consumed, allocate from `c`; `c = m` → resolved not
persisted, `m` may be allocated; anything else, or an unreadable counter →
recovery, marker retained. FIXED mode never enters it.

### 8.7 Attempt metadata by state

The attempt row is gated on `AutoIdentityKnown` (the diagnostic fact), so a
post-write failure can never blank the identity and make an advanced counter look
like a skipped nonce. **Published** records — snapshot and commit block — stay
gated on `NonceConsumed`: a published bank may claim only *proven* consumption.
Controls refuse each direction being swapped.

### 8.8 The compound residual — withdrawn by §9

This section used to say that indeterminate persistence **plus** a failed
attempt-row write lost the reconciliation authority. That was true only while
the attempt row **was** the authority. It no longer is, so the residual as
stated is withdrawn rather than kept because it was once accepted.

What replaces it is narrower and is now the honest statement: an attempt-row
write failure loses the **audit line**, not the safety property. The pending
marker is established before the counter is touched and lives in a cell no
attempt writer addresses, so nonce-reuse prevention does not depend on that row
storing anything. `indeterminate_marker_storage_failure` in the contract states
exactly this, and `test_44r` holds the source to it.

Integration `test_28` still states its own limit: it proves failures are
**routed** to the attempt recorder, and does not claim the row can never fail to
be **stored**. That limit is now about audit completeness only.

### 8.9 Detectors that pinned the wrong authority

`test_44h2`, `test_44h3`, `test_44h4`, `test_91`, `test_92`, integration
`test_29`, `test_52`, `test_53` and every `claimed < write` assertion were
**rewritten, not re-anchored** — a control that pins a rejected definition is
worse than no control. They now assert the contract's order,
`read < derive < persist < established < sampling`, and a control
(`test_91`/`test_94`) restores the rejected pre-write flag and the false
non-reuse promise and requires refusal.

### 8.10 Hashes, with the cause of each movement

| Artefact | Before | After | Why |
|---|---|---|---|
| `spec/sim_contract.yaml` | `715f0025…d55076` | `a1b6f35a45f5…` | fifth token + three-state nonce authority |
| `spec/structure_contract.yaml` | `cb0d03af…439b61` | `68f135d35af4…` | `modSimNonce` registered; reporter responsibility narrowed |
| `build/vba/modSimContract.bas` | `1d949be6…293a753` | `96e22c842c11…` | projects the fifth token |
| `build/phase6_cases.json` | `98f83537…78b6d1` | `8019683a0490…` | projects the attempt-result axis |
| `build/stage_b_manifest.json` | `51335e33…272049` | `58e0f1170d5a…` | module registry gained `modSimNonce` |
| `src/vba/modSimReport.bas` | `0797e307…3af1f4` | `f57b6d06abbd…` | nonce lifecycle extracted; audit gate renamed |
| `src/vba/modSimNonce.bas` | — | `a6806b05d24b…` | new |

**Unchanged, and checked:** `modSimRng`, `modSimSample`, `modSimEngine`,
`modSimStats`, `modSimFingerprint`, `modCalcFingerprint`, `modCalcReport`, and
every Phase-5 authority including `calc_contract.yaml` and `calc_loader.py`.
D6-11 did not move: `MRG32k3a → [modSimRng]`, `RunSimulation → [modSimReport]`,
`Percentile → []`.

### 8.11 One further control found by the suite

`test_128` sweeps every settled contract leaf and mutates it to a
type-compatible wrong value. My first draft wrote the four next-run resolutions
as free-form prose validated only for non-emptiness, so a wrong value passed.
They are now settled tokens — `CONSUMED`, `PRE_ALLOCATION`, `RECOVERY_REQUIRED`,
`RECOVERY_REQUIRED` — validated by exact value, with the explanation in the
comment where it belongs. The control was right and the contract was weak.

---

## 9. The durable pending-AUTO-nonce sidecar

Static review rejected §8's carrier. `Last Attempt Result` is contractually the
result of the **chronologically last** attempt, and a durable AUTO lock has to
survive attempts that have nothing to do with it. A FIXED run is the decisive
case: the contract correctly excuses it from AUTO reconciliation
(`applies_to_fixed_mode: false`), but a FIXED attempt still rewrites the whole
Last Attempt block. Preserving the token across it would make *"Last Attempt
Result"* lie; overwriting it makes the recovery authority disappear. One cell
cannot be both. That is a structural contradiction, not another flag bug.

### 9.1 The coordinate, and why it is free

`_SimData!F21`, emitted as `SIM_PENDING_AUTO_NONCE_CELL`. Verified free three
independent ways, none of them by reading a comment:

* **Contract structure.** Column F is the bank-B value column. The bank-B
  snapshot occupies the *snapshot* field group, which ends at row 20; row 21 is
  a **shared** counter row (`Next AUTO Nonce` at D21) and therefore has no
  bank-B twin. The shared final commit is `D22:D30` — column D. The bank-B
  iteration index also uses column F, but from the header row 33 downward.
* **The built workbook.** `openpyxl` on `build/PCCM_stageA.xlsx` reports `F21`
  empty, and the whole of column F carries nothing at row 21.
* **A builder validator.** `_validate_pending_auto_nonce` recomputes the
  overlap from the layout and refuses a cell that collides with the bank-B
  snapshot, any banked row, the identity block's bounds or the iteration table.
  Moving the declared cell to `F20` is refused.

`test_32` holds all three; `test_33` proves nothing shifted — header row 33,
first iteration row 34, `reserved_rows_h` 33, and
`max_iterations_representable` still 1 048 543, still recomputed rather than
restated. **No row was added, so nothing moved.**

### 9.2 Write-ahead

```
read NEXT_AUTO_NONCE = m
derive effective seed(m)
establish + VERIFY  Pending AUTO Nonce = m     <- counter untouched until here
attempt NEXT_AUTO_NONCE = m+1
reconcile
clear Pending AUTO Nonce
Phase6AfterNoncePersisted
sampling
```

The counter is not touched until the marker is durably established, so **either
physical outcome of the marker write is safe**: if it did not land the next AUTO
run sees a blank sidecar and counter `m`; if it did, it sees pending `m` and
counter `m` and resolves `PRE_ALLOCATION`. No nonce can vanish or replay because
of uncertainty in *that* write. A failed establishment is a **definite**
outcome — `PRE_ALLOCATION`, counter never touched — not an indeterminate one.

The contract states the relationship to the Step-0 coarse sequence as a
*checkable* property rather than a claim: `write_ahead_order_refines: "order"`,
and the loader proves the coarse steps survive as an ordered subsequence. A
future edit that drops or reorders one is refused, so `order` cannot quietly
become a stale second authority.

### 9.3 Clearing

`m` and `m+1` are definite resolutions and clear the marker; `neither`,
`unreadable` and `unobtainable` retain it. Clearing is treated as a **real COM
write** — a raised clear is not proof the marker survived. Both physical
outcomes are safe for the *next* run, since a surviving marker costs one extra
reconciliation, but **this** run must not sample while its own cleanup is
unresolved, and nothing rolls the counter back to tidy up.

### 9.4 What the fifth token now means

`AUTO_NONCE_INDETERMINATE` remains, as an **audit** result only: *this attempt
met an AUTO nonce persistence outcome it could not classify.* A later attempt,
FIXED included, may overwrite it freely. Phase 5's attempt axis stays four-valued
and neither `calc_contract.yaml` nor `calc_loader.py` names the token.

**THE RULE, SETTLED.** The token is emitted **only** when THIS attempt's
`allocationState` is `PERSISTENCE_INDETERMINATE`.

* `RECOVERY_REQUIRED` is a separate **action** and does not by itself earn it.
* A known `CONSUMED` or `PRE_ALLOCATION` observation plus `recoveryRequired`
  records **`REFUSED`** — the observation stands, and one result string cannot
  hold both physical classifications anyway.
* A prior-marker recovery refusal, taken before this run selects an identity,
  records **`REFUSED`**.
* **F21, not the attempt token, is the durable cross-invocation recovery
  authority.**

> **Superseded (this section, first draft).** This paragraph originally said
> *"Both unclassified states earn it — `PERSISTENCE_INDETERMINATE` and
> `RECOVERY_REQUIRED`"*. **That is the rejected conflated design and is
> withdrawn.** It treated a recovery action as though it were an allocation
> classification, which is precisely what let a cleanup failure overwrite a
> proven `CONSUMED` observation. §9.8 states the two-axis rule that governs, and
> the table there is the authority.

### 9.5 The three source defects fixed alongside

| Defect | Was | Now |
|---|---|---|
| raised first verification read | handler disarmed before the read, so a raise skipped reconciliation | the write **and** its verification read sit in one envelope; `verification_read_raised` is a named contract cause |
| effective seed lost on failure | copied on the success arm only, so refusals wrote nonce `m` beside seed `0` | out-parameters copied **unconditionally** after the call — no branch or early exit stands between |
| failpoint fired in FIXED | call sat below an `End If` the FIXED branch fell through | the FIXED branch `Exit Function`s before it; a nesting-aware block check proves containment, not text order |

### 9.6 Detectors, and the gap that made this round necessary

The thirteen focused controls all **passed** while every one of these defects
was present. That is a detector failure as much as a source failure, and the
pattern in each case was the same: they proved a construct was *present* rather
than that a *path* had a property. The replacements are structural — they walk
logical statements, extract blocks by nesting, and assert what stands between
two points rather than merely that both exist.

Nineteen properties are now controlled, each with a non-vacuous mutation, and
the four rejected shapes are restored verbatim as `test_103`, `test_105`,
`test_109` and `test_110` and required to fail.

One of my own new controls was string-blind in the opposite direction:
`test_44h3` scanned `code_without_string_removal`, which strips **comments**, so
it could never have seen the obsolete `ALLOCATED, not CONSUMED` language come
back. It now reads the raw text, and asserts the raw text still has comments in
it before concluding anything from their absence.

### 9.7 Hashes for this correction

| Artefact | Before | After | Why |
|---|---|---|---|
| `spec/sim_contract.yaml` | `a1b6f35a45f5…` | `8997ec163722…` | sidecar authority; `fixed_mode`, `attempt_result_token`, `pending_clear`, `write_ahead_order`, `recovery_action`; residual rewritten |
| `spec/structure_contract.yaml` | `68f135d35af4…` | `a5ef5d3e8cad…` | `modSimNonce`'s registered responsibility no longer names the rejected carrier |
| `builder/pccm_builder/sim_loader.py` | *(§8)* | `ef8cb9dbf070…` | `_validate_pending_auto_nonce`; the new blocks **enforced**, not merely tolerated |
| `builder/pccm_builder/sim_emit.py` | *(§8)* | `11ee506b0bc3…` | emits `SIM_PENDING_AUTO_NONCE_CELL` |
| `build/vba/modSimContract.bas` | `96e22c842c11…` | `daa4d27889c3…` | **one line added**, the sidecar constant |
| `build/stage_b_manifest.json` | `58e0f1170d5a…` | `87d0d7bed10a…` | projects the corrected responsibility text |
| `src/vba/modSimNonce.bas` | `a6806b05d24b…` | `6e0ed05c90c0…` | write-ahead transaction; attempt-row dependency removed; allocation and recovery axes split |
| `src/vba/modSimReport.bas` | `f57b6d06abbd…` | `cfc8ed0a1299…` | unconditional out-parameter copy; `RefusalResult` covers both unclassified states; obsolete comment replaced |

**Unchanged, and checked:** `build/phase6_cases.json` (`8019683a0490…`), every
Phase-5 authority including `calc_contract.yaml` and `calc_loader.py`, and every
other handwritten module. The generated module's movement was verified by
rebuilding from `HEAD` and diffing: exactly one added line.

### 9.8 The allocation axis and the recovery axis are separate

Static review found one more defect in §9 as first implemented, and it was mine
by explicit choice: I carried `RECOVERY_REQUIRED` in the same scalar as the
three allocation states, and documented the decision in a comment defending it.
The comment was wrong. **A cleanup problem is not a persistence problem.**

The failing path was concrete. `PersistAdvance` observes the counter at `m+1`
and sets `CONSUMED` — the nonce *is* spent, and that is now a proven fact. The
transaction then tries to clear F21 and cannot. The old code wrote
`state = RECOVERY_REQUIRED`, and the reporter, deriving consumption by comparing
that one string against `CONSUMED`, reported `NonceConsumed = False` for a nonce
it had just proved consumed — and recorded `AUTO_NONCE_INDETERMINATE` for a
transition it had already classified. Nothing about the counter had changed.

There are now **two scalars**, and the boundary carries both:

```
allocationState  As String   PRE_ALLOCATION | CONSUMED | PERSISTENCE_INDETERMINATE
                             (NOT_APPLICABLE in FIXED)
recoveryRequired As Boolean  the ACTION: reconcile before the next AUTO allocation
```

The rule, stated once and controlled: **a recovery or cleanup problem must never
erase an allocation fact that was already established.** By path:

| Path | `allocationState` | `recoveryRequired` | Attempt result |
|---|---|---|---|
| counter observed `m+1`, clean | `CONSUMED` | false | `SUCCESS` |
| counter observed `m+1`, clear fails | **`CONSUMED`** | true | `REFUSED` |
| counter observed `m`, clear fails | **`PRE_ALLOCATION`** | true | `REFUSED` |
| observation unavailable after the write | `PERSISTENCE_INDETERMINATE` | false | `AUTO_NONCE_INDETERMINATE` |
| counter neither `m` nor `m+1` | `PERSISTENCE_INDETERMINATE` | true | `AUTO_NONCE_INDETERMINATE` |
| prior marker unreconcilable, before identity selection | `NOT_APPLICABLE` | true | `REFUSED` |

The last row is the other half of the correction. A run that refuses while
reconciling a **prior** marker never began a transition of its own, so it has
nothing to be indeterminate *about*; manufacturing the token there would
broaden it past its accepted meaning. F21 — not the attempt string — is what
blocks the next AUTO allocation in every one of these rows.

`RECOVERY_REQUIRED` is therefore **not** a fourth allocation state, and the
loader refuses a contract that lists it as one. `NonceConsumed` is derived from
`allocationState` alone: not from `identityKnown`, not from the call's Boolean,
and never from `recoveryRequired`.

The suite had actively **pinned** the conflation — `test_44q` required a clear
failure to produce `RECOVERY_REQUIRED`, and `test_44s` required
`RECOVERY_REQUIRED` to earn the token. Both were rewritten rather than
re-anchored; `test_44q2` and `test_44q3` were added for the axis separation and
the survival rule, and `test_127`–`test_135` restore the rejected shapes and are
required to fail.

### 9.9 One observation, not a change

`F21` sits on the row whose label at `B21` reads *"Next AUTO Nonce"*, and the
note column `H21` is empty. On the sheet a pending marker will therefore appear
beside a label naming a different field. Nothing depends on this — the cell is
addressed by a generated constant, never by its label — but a reader inspecting
`_SimData` by eye has no on-sheet name for it. Adding a label or a note would
change the emitted layout, which this authorisation did not ask for, so it is
reported rather than done.


## 10. Fifth round: the active wording, closed

The two-axis model of §9.8 was accepted as **functionally** settled. Three
comment blocks in `modSimReport.bas` and one paragraph of §9.4 still taught the
design it replaced, and nothing in the suite could refuse them. Prose that
contradicts the executable authority is a live defect: the next reader deciding
whether a change is safe reads the comment, not the transaction.

**No executable statement moved in this round.** `modSimReport.bas` holds 728
logical statements before and after, in the same order, hashing identically
(`4a28c1d6af3368e9b0f0`); only comment lines changed.

### 10.1 The three source statements

| Where | Said | Now says |
|---|---|---|
| module header | the module owns *"the AUTO nonce lifecycle"* | it explicitly **does not** — the transaction, the write-ahead marker and the durable recovery protocol belong to `modSimNonce`, driven through a narrow scalar interface |
| `RecordRefusal` | *"an unresolved AUTO advance is its own durable result"*, and the token is what *"the NEXT run reads"* | the result is an **audit classification for this attempt only**; the next run does not read it; F21 with the counter is the durable recovery authority |
| `WriteAttemptBlock` | whether the nonce was consumed *"is carried by the attempt result"* | consumption is the `allocationState` fact projected as `NonceConsumed`; the fifth token identifies `PERSISTENCE_INDETERMINATE` and nothing else |

### 10.2 The document

§9.4's rule is restated as the single-axis rule the source implements, and the
withdrawn sentence is **quoted and withdrawn** rather than deleted — the history
of a defect is part of the record, but it may not stand as though it were still
the rule. §8.5's classification table gains the same note: its observation rows
stand, its single-axis presentation does not, and §9.8 governs.

### 10.3 Detectors

`test_44u` was extended past the nonce module onto the reporter. It extracts
comment prose as the difference between the raw text and the comment-blanked
code, flattens it — VBA wraps a sentence across `' ` lines, so a phrase check on
raw text would miss half of them — and holds it to six **affirmative** rejected
phrasings and two proximity pairs. Affirmative is deliberate: the accepted prose
*denies* each claim, so *"it is not the recovery lock"* cannot match *"is the
recovery lock"*, and a paraphrase that still teaches the rejected design is
caught where a full-sentence literal would not be.

`test_44v` is new, and it is the first detector to read this document at all.
That is the gap that let §9.4 contradict §9.8 indefinitely. It judges active
prose and blockquoted record by different standards: a claim outside a `>` block
is one the document still makes, a claim inside one has been withdrawn. It
requires the exclusivity rule keyed on `allocationState`, refuses the conflated
phrasings, requires every surviving active mention of `RECOVERY_REQUIRED` to
disclaim itself as an allocation classification, and requires the withdrawal to
be on the record rather than silently deleted.

Four mutation controls restore the withdrawn wording — three in the reporter,
one in this document. Each is refused by **exactly one** detector, the named one:
three by `test_44u`, the document by `test_44v`. All four are byte-faithful
restorations of the text the battery passed over at `d36d5d4`.

### 10.4 Hashes

| File | Was | Now | Cause |
|---|---|---|---|
| `src/vba/modSimReport.bas` | `cfc8ed0a…dfc60e` | `49b7602a…4d2b06` | comment blocks only; 1179 → 1197 lines, 728 logical statements unchanged |

The `modSimReport` pin in `tests/test_phase6_integration_source.py` moves with
it. That is the whole of the change outside the four files this round names, and
it is mechanical: a comment-only edit moves a byte hash.
