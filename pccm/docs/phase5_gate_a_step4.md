# Phase 5 — Gate A — Step 4: the pure VBA numerical kernel, as source

**Status: CORRECTED TWICE after independent review — ready for re-review.**

Step 3 is accepted and closed. This step adds three hand-written VBA modules —
`modCalcFactors`, `modCalcAnalytical`, `modCalcFingerprint` — that implement the
accepted Step-2 numerical semantics and the locked fingerprint encoding, plus the
static Linux test suite that reads them, plus the narrow build plumbing that lets
the contract declare them.

---

## Second correction round — runtime-capability and API defects

Independent review of `2d76d78` confirmed that the first correction genuinely
fixed the fingerprint schema, the inflation vector, the `FP_VERSION` authority,
the exact binary constants, the split C1 ownership, the conditioning
underflow/overflow split and the removal of `Variant`. It then found two
**runtime-capability** blockers and two source/API hardening items.

The common thread in the two blockers is the same mistake made twice: **VBA
cannot represent a zero-element array.** An allocated array always satisfies
`UBound >= LBound`, and an unallocated dynamic array raises Error 9 on `LBound`
before any emptiness test could run. Every branch guarded by a count derived
from array bounds is therefore dead code — including the empty-driver branch
this document previously described as fixed. **That claim was wrong**, and the
API, not the branch, was the problem.

### 1. The driver-record buffer was one field short whenever ML was present

The field ORDER was correct. The allocation was not:

```vba
ReDim fields(0 To 5 + inflationCount + weightCount)   ' 6 + n + m SLOTS
```

The record has six fixed fields — Permanent ID, Distribution, the kind-specific
scalar, Min, Max, FX — plus both vectors, plus a seventh when Most Likely is
present. The locked one-year Triangular cost record emits **nine** fields into
**eight** slots, so the source could not build the reference record at all, even
though the static schema test was green.

The capacity is now computed and the emitted count is checked against it:

```vba
fieldCount = 6 + inflationCount + weightCount
If includeMostLikely Then fieldCount = fieldCount + 1
ReDim fields(0 To fieldCount - 1)
...
If count <> fieldCount Then Exit Function
```

`count` — never the array size — remains the encoded field count. The final
guard turns a future schema edit into a controlled failure rather than another
silent buffer mismatch. `test_53` evaluates the source's own formula
symbolically for both ML cases; `test_53a` evaluates it for the locked reproducer
and requires nine.

### 2. The empty-driver fix was unreachable

`AccumulateTotals`, `BuildAnnualSeries` and `Reconcile` all derived their driver
count from the array bounds, so the accepted empty model — zero cost lines, zero
risks, zero totals, annual rows still present, ten identities holding — could
never be expressed.

The logical count is now a parameter at every aggregate boundary, and the
zero-count branch is settled **before any bound of the corresponding array is
read**. The private order helpers take it too, so the same defect cannot
reappear one level down. No dummy driver, no `Variant`, no allocation probing,
no error-handler trick.

`test_63` proves the ordering structurally: it locates the explicit
`If <count> = 0 Then` branch and the first statement that reads a bound of, or
subscripts, each possibly-empty array, and requires the branch to come first.

### 3. `SafeProduct` could not return the empty product

`safe_product([]) == 1.0` is locked behaviour, and

```vba
If UBound(factors) < LBound(factors) Then
```

has exactly the defect above: unreachable for an allocated array, raising for an
unallocated one. `SafeProduct` and `SafeSignedSum` now take an explicit logical
count, refuse a negative one, and settle the multiplicative and additive
identities before touching a bound. Every call site passes its own count —
`TripleProduct` and `ExpectedRisk` pass 3, `BuildFactor` passes `width`,
`AnnualSeries` passes `count`, the I5 profile sum passes the project-year count.
For a positive count the Tier-1 result and the exact Tier-2 rescue are unchanged.

### 4. `ExactQuotientOfSum` reached a raw division on an unsupported divisor

It is Public, it installs no error handler, and it performed no validation before
`ExactDivideSmall`. A divisor of zero would have reached a raw division. The
public contract is now **exactly `{2, 3, 6}`** — the three convex-statistic
denominators — validated before any work is done. Any other divisor returns
`False` and leaves the result unchanged. The unreachable identity path for a
divisor of `1` has been removed rather than left as dead code, and
`ExactDivideSmall` remains the locked small-divisor kernel.

### 5. Three helpers were Public without a cross-module caller

The whitelist comment claimed every Public name had an external caller. Source
reference inspection disproved that for `DistributionMean`, `CanonicalOrder` and
`CalcFpNumberField` — each is called only inside its own module, and none was
part of the requested Gate-B diagnostic surface. All three are now `Private`,
with semantics unchanged and no inlining.

`test_64h` now **scans references** rather than trusting the comment: a Public
name with no caller outside its own module must appear in an explicit exception
set with a stated reason. The exceptions are the ten fingerprint helpers that
form the Gate-B diagnostic surface, the primitives and rescues the resolver layer
will call directly, the five distribution and audit entry points the Step-4
instruction requires to be Public, and `AllIdentitiesHold`, which later
orchestration consumes.

---

## First correction round — five blocking defects found by independent review

Independent review of `7fac269` reproduced every reported static result and
accepted the three-module architecture, the line-limit retarget, the module
inventory, the worksheet-free boundary, the four scoped arithmetic handlers and
the Phase-4 source preservation. It then found five blocking source defects.
None of them was an ambiguity in the design. All five are corrected here.

### 1. The fingerprint driver-record schema was wrong

The submitted encoder emitted, for every driver, both `Quantity` and
`Probability` with the inapplicable one set to `1`, and **emitted no resolved
inflation-factor vector at all**.

That conflated two different sections of the accepted plan. `Quantity = 1 for
risks, Probability = 1 for cost lines` is the in-memory `DriverFactors` carry
convention that the calculation and the simulation share. It is not the record
schema. The locked schema encodes the driver's own kind-specific scalar and
nothing for the other kind.

The missing inflation vector was the more serious half: without it, a change in
a referenced inflation factor leaves the record unchanged, and a stale result
presents itself as current. That defeats the entire stale-results design.

**Why the 366-unit reference did not catch it.** Golden case 1 has FX = 1,
inflation factor = 1 and profile weight = 1, so its three trailing numeric
fields are `1, 1, 1`. Substituting a `Probability` of `1` for the missing
inflation factor of `1` produced the same stream. The masking was possible only
because every value in that region was the identity. `test_53c` now digests the
same record with a **non-identity** inflation factor of `1.05` and proves the
slot is distinguishable from an identity, from its absence, and from the weight
beside it.

### 2. The exact-kernel constants were not the exact constants

```
                    submitted             correct
TWO_52              4503599627370500      4503599627370496  = 2^52
MAX_SIGNIFICAND     9007199254740990      9007199254740991  = 2^53 - 1
MAX_DOUBLE          1.79769313486232E+308 1.7976931348623157E+308
```

`TWO_52` was wrong by +4 and `MAX_SIGNIFICAND` by −1. These are not
documentation values: `DecomposeDouble` multiplies by `TWO_52` to produce an
**integer** mantissa for the limb kernel, and `RoundExact` compares an exact
significand against `MAX_SIGNIFICAND` to classify the `MAX_DOUBLE` boundary. The
first fed a non-integer into the exact kernel; the second misclassified the
largest representable Double.

The `MAX_DOUBLE` literal was also mathematically **above** the true maximum, and
a floating-point literal whose mathematical value exceeds the greatest value its
type can represent is statically invalid. `test_57` compares the literal as an
exact `Decimal` against `(2^53 − 1) × 2^971`, because `float()` would round the
defect away.

`MIN_NORMAL_DOUBLE` was unused and has been removed. `test_58` now fails any
declared constant that nothing reads.

### 3. The C1 magnitude structure was cleared twice

`ReconciliationMagnitudes` carries both the headline A/B/C/D/E scales and the
annual Base/Risk/Total scales, and `Reconcile` needs all of them at once. Both
`AccumulateTotals` and `BuildAnnualSeries` began with a whole-object clear, so
whichever ran second erased what the first had captured. **No call order could
produce the structure `Reconcile` requires** — a real integration impossibility
in the public pure API, not a stylistic problem.

Ownership is now split. `PrepareMagnitudeCoefficient` sets the coefficient on an
untouched record and otherwise verifies it, failing deterministically on a
conflict rather than reinterpreting existing magnitudes against a tolerance they
were not measured for. `ClearHeadlineMagnitudes` and `ClearAnnualMagnitudes`
each clear their own half and leave the other alone. The result is independent of
call order. `test_59` asserts the two halves are disjoint and together cover
every magnitude field.

### 4. An empty driver set was refused

`Reconcile` began with an explicit `"no drivers"` refusal. The accepted Step-2
oracle has `test_an_empty_driver_set_is_not_refused`, and no accepted contract
invented a minimum-driver rule.

The refusal was removed in this round, **but the replacement was still not
reachable**: the count was derived from the array bounds, which can never be
zero. The second correction round above fixes that properly with an explicit
logical count. This round's claim that the empty model was expressible was
wrong.

### 5. Conditioning overflow was silently turned into zero

`ConditioningScaledProduct` did `If Not ExactSumOfProducts(…) Then scaled = 0#`,
treating **every** exact failure as an accepted underflow. C1's exception is
narrow: a scaled term too small to represent cannot move an allowance floored at
`coefficient × 1`, so losing it changes no answer. A magnitude **outside** Double
range is a different fact entirely, and recording it as zero understates the
conditioning scale — narrowing a tolerance by accident, which is no better than
widening one.

`ConditioningScaledMagnitude` compounded this by retrying `coefficient *
magnitude` as raw arithmetic after `SafeMultiply` had already refused it.

Both now go through `ExactSumOfProductsCore(…, underflowToZero:=True)`. One
kernel, two policies, and the policy is a parameter:

| | model arithmetic | C1 conditioning |
| --- | --- | --- |
| exact value representable | rounded result | rounded result |
| exact value below the smallest Double | **refusal** | zero, success |
| exact value above `MAX_DOUBLE` | **refusal** | **refusal** |

`test_67` asserts the three range refusals inside `RoundExact` do not consult the
underflow flag, so an overflow can never be relabelled as an accepted underflow.

### 6. `Variant` had been added to the numerical boundary

The submitted kernel used `ByRef factors As Variant`, `ByRef groups As Variant`
and `expression() As Variant`, and the static test and this document had both
quietly added `"Variant"` to the allowed type list. **That was not an approved
constraint resolution**, and widening a boundary by editing the test that
enforces it is exactly backwards.

`Variant` is gone from the three modules. `SafeProduct` and
`ConditioningScaledProduct` take `Double()` directly. Sum-of-products uses a
flat typed vector:

```vba
ExactSumOfProducts(factors() As Double, groupStarts() As Long,
                   groupLengths() As Long, groupCount As Long, result As Double)
```

`factors` holds every factor of every group end to end; group *g* occupies
`groupLengths(g)` entries beginning at offset `groupStarts(g)`. Groups may differ
in length, which the two production call sites need. `test_68` fails on any
`As Variant` in executable code in any of the three modules, and `"Variant"` has
been removed from `ALLOWED_PARAMETER_TYPES`.

---

## What this step does NOT claim

These statements are the first thing in this document because they are the ones a
reviewer must be able to rely on.

1. **NO VBA WAS EXECUTED.** There is no VBA interpreter on Linux, and none was
   simulated. Nothing here observed VBA arithmetic.
2. **No number in this step was produced by VBA.** Every expected value in the
   repository still comes from the Python oracle, which remains the semantic
   authority.
3. **Fingerprint parity with VBA is not established.** The reference digest
   `50B6EB0E26857EA7` over 366 UTF-16 code units remains a Python result. Whether
   the VBA encoder reproduces it is Gate B's question, on real Excel on Windows.
4. **The kernel has never been compiled.** `Option Explicit` is present in all
   three modules and the source tests check block balance, line length and
   declaration shape, but a static reader is not a compiler. A type error that
   only the VBA compiler can see would still be found on Windows.
5. **No Windows artefact changed.** The bootstrap, the COM lifecycle script and
   `bootstrap/windows/phase4_functional_test.ps1` are untouched.
6. **No calculation endpoint exists.** There is no `PCCM_Calculate`, no Calculate
   button, no `modCalcResolve`, `modCalcCheck` or `modCalcReport`. A static test
   asserts each of those names is absent.

---

## The three modules

The accepted split is by responsibility, and no fourth production module was
added.

| Module | Responsibility |
| --- | --- |
| `modCalcFactors` | Range-checked arithmetic primitives, the exact Double-limb rescue kernel, the iterative inflation and discount series, `Knom` / `Kpv`, and the C1 conditioning magnitudes. |
| `modCalcAnalytical` | Distribution statistics, the per-driver audit amounts, A/B/C/D/E, the six annual series, and the identities I1–I5. |
| `modCalcFingerprint` | Canonical encoding, the double-modulus reduction, and the digest. No analytical quantity is computed while encoding one. |

### Line metrics

| Module | Raw | Blank | Comment | Code | Code < 900 | Raw < 1200 |
| --- | --- | --- | --- | --- | --- | --- |
| `modCalcFactors` | 1088 | 49 | 230 | 809 | yes | yes |
| `modCalcAnalytical` | 1177 | 63 | 246 | 868 | yes | yes |
| `modCalcFingerprint` | 499 | 28 | 186 | 285 | yes | yes |

All three are below both applicable limits. The figures were 995/781, 1098/826
and 428/255 at first submission and 1065/802, 1164/869 and 485/282 after the
first correction. `modCalcAnalytical` remains the closest to a threshold, with
32 code lines of headroom; the count plumbing added signature lines and removed a
pass-through helper, netting one line fewer than before. A comment line is one whose first
non-whitespace character is the VBA apostrophe; a blank line is neither comment
nor code.

`modCalcAnalytical` first came to 1231 raw / 959 code, over both limits. Nothing
semantic was removed to bring it down. What was removed was repeated error
plumbing — a four-line `If Not … Then / detail = / Exit Function / End If` at
twenty call sites became a one-statement call whose helper sets the diagnostic —
and five hand-written `Select Case` dispatchers over the six annual series, which
became one table built once. Every comment survived, the four safe arithmetic
primitives were not merged, and no rule moved.

---

## The size limit, and why it was retargeted

The Phase-4 assertion `raw lines < 900` was **a proxy, not a defect**. The thing
worth detecting is a collapsed responsibility split, and inside Phase-4 territory
raw size tracked that faithfully. It needed a responsibility-aware Phase-5
extension once a module arrived whose responsibility is coherent but whose
contract requires it to explain, at length, why each rescue tier exists and which
evaluation orders were rejected.

`test_05_no_module_is_a_dumping_ground` is still present, under the same name,
and now carries two limits instead of one:

* the seven Phase-4 modules keep `raw lines < 900` **exactly, unrelaxed**;
* the three Phase-5 kernel modules must satisfy **both** `code lines < 900`
  **and** `raw lines < 1200`.

Neither limit alone would do. A code-only limit would let a module grow without
bound in prose; a raw-only limit is what charged documentation as sprawl.

Size is only half the pair. The other half is
`tests/test_phase5_vba_source.py`, which asserts the responsibility boundaries
directly: each module must declare what it may own, and must not declare what
another module owns — `modCalcFactors` may not hold a distribution mean,
`modCalcAnalytical` may not hold a safe primitive or the fingerprint,
`modCalcFingerprint` may not compute `Knom`, a mean or a reconciliation.

---

## The pure-numerical boundary

Executable code in all three modules — comments and string literals stripped
first — contains none of `Application.`, `ThisWorkbook`, `ActiveWorkbook`,
`Worksheets`, `Worksheet`, `Range`, `Cells`, `ListObject`, `ListObjects`,
`Names(`, `Evaluate`, `WorksheetFunction`, `modWorkbook.`, `Rnd`, `Randomize`,
`MRG32k3a`, `NPV`, `Percentile`.

No parameter has an Excel object type, and **no parameter or local is a
`Variant`**. The only declared parameter types are `Double`, `Long`, `Boolean`,
`String`, the locked carry types `DriverFactors` and `YearFactors`, the
analytical records `DriverAudit`, `AnalyticalTotals`, `AnnualRow`,
`ReconciliationMagnitudes` and `IdentityCheck`, and the exact kernel's private
`ExactNumber`. Numerical factor groups are described by typed `Double()` and
`Long()` vectors — never a `Variant`, `Collection`, `Dictionary` or `Object`.

Zero procedures across the three modules begin with `PCCM_`.

---

## Error handling

There is no `On Error Resume Next` anywhere. Exactly four procedures install a
handler — `SafeAdd`, `SafeSubtract`, `SafeMultiply`, `SafeDivide` — and each
scopes it to one arithmetic expression:

```
On Error GoTo ArithmeticFailure
tmp = a <op> b
On Error GoTo 0
<range post-checks>
Exit Function
ArithmeticFailure:
On Error GoTo 0
SafeX = False
```

The tests assert the shape line by line: the guarded statement must match
`tmp = a <op> b`, the next line must disarm, and each label must disarm and
return `False`. On failure the caller's `ByRef` result is never written.
`SafeAccumulate` installs no handler of its own; it calls `SafeAdd`.

`modCalcAnalytical` and `modCalcFingerprint` install no handler at all. A failure
there is a returned `False` with a diagnostic string.

---

## The exact kernel

`LIMB_BITS = 24`, `LIMB_BASE = 16777216#`, `TWO_52 = 4503599627370496#` (2^52),
`MAX_SIGNIFICAND = 9007199254740991#` (2^53 − 1), `MAX_DOUBLE =
1.7976931348623157E+308` ((2^53 − 1) × 2^971). A value is `(sign, base-2^24
limbs, binary shift)` held as Doubles. The module contains no `Currency`, no `Decimal`,
no `CDec`, no `CCur`, no `Eval`, no `WorksheetFunction`, no native `Mod` and no
native `\`. The only `CLng` in the three modules narrows a hex digit, which is 0
to 15; a test asserts that this is the sole narrowing.

The internal stages are all present as separate private procedures: exact
decomposition, magnitude compare / add / subtract, exact signed sum, exact
product, exact sum of exact products, bit lookup and sticky, final rounding, the
guarded small-divisor path for 2 / 3 / 6, and power-of-two scale back.

**Two tiers, and tier 2 is exact.** `SafeSignedSum` and `SafeProduct` each run
the ordinary staged path first and return its result bit for bit when it
succeeds. Only on failure do they reach `ExactSumOf` / `ExactProductOf` and
`RoundExact`. Neither the superseded rounded positive-minus-negative cancellation
nor the superseded magnitude-balanced ordering appears; a test asserts each tier-2
path reaches the exact routines.

---

## The C2 materialization boundary

`BuildKnom` and `BuildKpv` stage `w_y * infl_y` (and `* disc_y`), sum in
project-year order, then apply FX. If that complete pipeline succeeds, the result
is returned bit for bit and FX is not distributed. Only on failure is the factor
re-formed as one exact expression `SUM_y (FX * w_y * infl_y)` and rounded once.

Neither `w_y * infl_y` nor the pre-FX sum is published, so neither is a
representability boundary. `Knom` and `Kpv` are published, so they are.

**Quantity and Probability are absent** from `BuildKnom`, `BuildKpv` and their
shared `BuildFactor`, and a test asserts the absence by name. Probability is
replaced by a Bernoulli draw in Monte Carlo; Quantity is a per-driver multiplier,
not a factor of the escalation path.

The factor series are iterative and never `(1+r)^(t-1)`: a test asserts no `^`
appears in either builder, because the power can overflow as an intermediate
where the factor is representable, and it cannot say which year failed.

---

## The analytical layer

**Five independent passes.** `B` is not computed as `C − A` and `E` is not
computed as `C + D`. Each measure keeps its own named contribution list —
`aNomTerms`, `bNomTerms`, `cNomTerms`, `dNomTerms`, `eNomTerms` and their PV
counterparts — and `E` is collected in its own second pass over the same
contributions. A test looks for the derivations by pattern and fails if one
appears.

**Six annual series, six boundaries.** Each of Base / Risk / Total in nominal and
PV is produced by its own call, and the annual Total is summed over its own
contiguous contribution list rather than added from the two series above it, so
I3c and I4c stay real checks. Each series is tier 1 (every contribution as a
Double, accumulated in canonical driver order, PV as `nominal * discount`) and,
only on failure, one exact sum of exact products with the discount factor inside
each PV product.

**Three tiers per convex statistic.** A distribution with zero uncertainty
returns its single point exactly. The stable staged form runs next — `Min/3 +
ML/3 + Max/3`, never `(Min + ML + Max)/3` — and its result is returned bit for
bit when non-zero. A staged **zero** falls through to the exact numerator,
because a zero can be a true zero or an underflow hiding a small non-zero answer
and only the exact numerator distinguishes them. The Beta-PERT numerator is four
copies of Most Likely, never `4 * ML`.

**C1 conditioning** is captured while contributions are accumulated, per driver
and per driver per year, never from the headline totals or the annual row
aggregates. A per-driver-per-year contribution that has no Double of its own is
still conditioned, through `ConditioningScaledProduct`, which folds the
coefficient into the same exact factor expression.

**Canonical order** is ascending Permanent ID on UTF-16 code units, via
`StrComp(a, b, vbBinaryCompare)`, cost lines first and then risks, each group
sorted independently. `vbTextCompare` appears nowhere.

**Tolerances are never restated.** `TOL_PROFILING_SUM_ABSOLUTE`,
`TOL_IDENTITY_ABSOLUTE_FLOOR`, `TOL_IDENTITY_RELATIVE_COEFFICIENT` and
`TOL_CONDITIONING_SCALE_FLOOR` are read from the generated `modCalcContract`, and
a test asserts no tolerance literal appears in kernel code.

### One correction to the drafted kernel

`IdentityAllowance` was drafted with the inner comparison in the wrong units: it
compared the raw `scaleFloor` against an already-scaled sum and then applied the
coefficient conditionally. It now matches the accepted oracle exactly —
`scaledFloor = coefficient * scaleFloor`, then `max(scaledFloor, scale)`, then
`max(absoluteFloor, …)`. Note the two maxima: the inner one is a **maximum**, not
an addition.

---

## The fingerprint

The grammar is unchanged from the reference implementation:

```
field   ::= <TAG> <LEN> ":" <VALUE>      LEN in UTF-16 code units
record  ::= F_I(field_count) field*
section ::= F_S(name) F_I(record_count) record*
stream  ::= F_S("PCCM-FP") F_I(version) section*
```

Sections are emitted in the locked order HEADER, COST, RISK.

**No hash constant is restated.** The literals `2147483647`, `2147483629` and
`131` do not appear anywhere in `modCalcFingerprint`; a test asserts each
absence, and asserts that `FP_BASE`, `FP_MOD_1`, `FP_MOD_2`, `FP_INIT_1`,
`FP_INIT_2`, `FP_STREAM_TAG`, the three tag constants and the three section
constants are all used.

**The reduction is the locked Double-only form**, asserted statement by
statement:

```vba
x = h * FP_BASE + u
q = Fix(x / modulus)
r = x - q * modulus
If r >= modulus Then r = r - modulus
If r < 0# Then r = r + modulus
```

Neither `Mod` nor `\` appears, and neither `x` nor `q * modulus` is narrowed to a
Long.

**UTF-16 handling** uses `Len`, `Mid$` and `AscW`. `CalcFpNormaliseCodeUnit` adds
65536 to a negative `AscW` result, because `AscW` returns a signed 16-bit
`Integer` and every code unit above `U+7FFF` comes back negative.
`CalcFpUtf16Length` is `Len(text)` — a VBA `String` is already UTF-16, so a
non-BMP character is already counted as its two surrogate units.

**The decimal separator is an argument.** `Application.International` appears
nowhere. `CalcFpCanonicalNumber` takes the separator the host formatter produced
and normalises it back to `.` **positionally**, rewriting exactly one character:
`E`, `+`, `-` and every digit already occur elsewhere in scientific notation, so
a global replace would corrupt the exponent. A test asserts `Replace` is absent
from that procedure.

**The digest** is `CalcFpHex8(h1) & CalcFpHex8(h2)`, eight hex digits each, with
both accumulators starting at `FP_INIT_1` / `FP_INIT_2`. `CalcFpHex8` divides in
Double rather than calling `Hex$`, so the accumulator is never handed to a
function that narrows to Long.

**Records are sorted by Permanent ID** with `StrComp(…, vbBinaryCompare)`, on a
private index permutation so the caller's arrays are not reordered — the
fingerprint must not have a side effect on the data it describes.

**`includeMostLikely As Boolean` is a parameter** of both record builders, passed
by the later resolver layer, which owns the distribution vocabulary.

### The record field order

Cost and risk are NOT the same shape:

```
COST:  S(PermanentId) S(Distribution) N(Quantity)    N(Min) N(Max)
       [ N(MostLikely) ] N(FxToSar) N(inflation)* N(weight)*

RISK:  S(PermanentId) S(Distribution) N(Probability) N(Min) N(Max)
       [ N(MostLikely) ] N(FxToSar) N(inflation)* N(weight)*
```

A cost line encodes `Quantity` and **no** `Probability` field. A risk encodes
`Probability` and **no** `Quantity` field. The opposite kind's multiplicative
identity is not fingerprinted at all.

Both vectors are resolved and both are encoded in project-year order, inflation
first. The inflation vector is the resolved **cumulative factor** per applied
project year — not the profile name and not the annual rates — because that is
what the calculation consumed. The encoder hashes exactly the vectors it is
handed; whether their lengths match Applied Duration is the later resolver and
check layer's question, not a pure encoder's.

Most Likely appears only when `includeMostLikely` says so, and that flag is
supplied by the caller. It is never inferred from the distribution text.

---

## Build plumbing

`spec/structure_contract.yaml` now declares twelve modules in the locked order:
`modConstants` (generated), `modWorkbook`, `modAppState`, `modTimeline`,
`modDrivers`, `modProfiling`, `modInflation`, `modStructuralCheck`,
`modCalcContract` (generated), `modCalcFactors`, `modCalcAnalytical`,
`modCalcFingerprint`.

**No numerical formula, fingerprint constant, tolerance or calculation rule was
added to `structure_contract.yaml`.** Each new entry carries a name, a
`generated` flag and a responsibility sentence, exactly as the eight existing
entries do. `structure_contract_version`, `VERSION`, `model_version` and
`BUILDER_VERSION` are unchanged; the workbook remains model version 0.5.0.

`structure_loader.py` generalises **only** the deployment invariant.
`vba.generated_module` remains `modConstants` and must still be declared
generated. What was `exactly one module may be generated` is now a comparison
against a statically locked tuple in the loader itself:

```python
GENERATED_MODULES = ("modConstants", "modCalcContract")
```

A module claiming to be generated that the builder does not emit is still
refused, and so is dropping the generated flag from the primary module.

`modCalcContract.bas` is **not** generated a second time. `emit_stage_b` still
owns `modConstants` and `emit_calc_artifacts` still owns `modCalcContract`. The
Phase-4 static-test fixture previously called `emit_stage_b` alone, so its
combined artifact directory held only one generated module; it now calls
`emit_calc_artifacts` as well. That is a fixture change only — one generator per
artifact, no duplication.

---

## Phase-4 regression discipline

The seven Phase-4 VBA source files are unchanged, proved by SHA-256 against the
digests recorded when Step 3 was accepted:

| Module | SHA-256 |
| --- | --- |
| `modWorkbook` | `9cfa8f13…5405bf` |
| `modAppState` | `ef0b5c64…04672f` |
| `modTimeline` | `4a4f24d1…7b9e3f` |
| `modDrivers` | `8f947a4c…02af48` |
| `modProfiling` | `0312858d…8d9d7ca` |
| `modInflation` | `08db3280…91118c` |
| `modStructuralCheck` | `1798c56a…53ed54` |

Buttons, `vba.entry_points` and `vba.harness_procedures` are unchanged. No new
`PCCM_` endpoint exists.

### Retargeted invariants

Each was retargeted, not deleted, and each retarget states which prior invariant
it carries.

| Test | Was | Now |
| --- | --- | --- |
| `test_03_…generated…` | `generated == ["modConstants"]` | `modConstants` is still the primary and still declared generated; the generated set equals the builder's locked tuple; a hand-written copy of either is still refused. |
| `test_05_no_module_is_a_dumping_ground` | one `raw < 900` cap for every hand-written module | Phase-4 modules keep `raw < 900` unrelaxed; kernel modules must satisfy `code < 900` **and** `raw < 1200`. |
| `test_17_no_calculation_or_simulation_code_leaked_in` | one list, forbidden everywhere | Simulation (`Rnd(`, `Randomize`, `MRG32k3a`, `WorksheetFunction.Percentile`, `RunSimulation`) is still forbidden **everywhere**, kernel included. Calculation names (`ExpectedValue`, `DiscountFactor`, `EscalationFactor`) are forbidden in the Phase-4 modules and the generated constants module, which is the territory that sweep was written to protect. |
| `test_11_every_constant_the_vba_references_is_emitted` | constants come from `modConstants` or the same module | constants may also come from `modCalcContract`, or from a `Public Const` exported by another hand-written module. A `Private Const` elsewhere is still invisible, exactly as a compiler would see it. |
| `test_rejects_a_second_generated_module` | one test | three: a hand-written module marked generated is refused; dropping the generated flag from the primary module is refused; a duplicate module name is refused. Plus one that states the current generated inventory in both directions. |
| `test_no_phase_five_vba_module_was_added_to_the_stage_b_manifest` | the kernel must be absent from the manifest | the kernel and `modCalcContract` must be **present**; `modCalcResolve`, `modCalcCheck`, `modCalcReport` and the calculation endpoints must be absent. |

### One Step-3 test module was modified, and why

Four of the five Step-3 corpus test modules are **byte-identical**:
`test_phase5_numeric.py`, `test_phase5_oracle.py`,
`test_phase5_calc_contract_validation.py`, `test_phase5_fingerprint.py`.

`test_phase5_stage_a.py` changed by exactly one test, and this is the genuine
integration requirement being reported. At Step 3 it asserted *"no Phase-5 VBA
module was added to the Stage-B manifest"*, which was correct then: Step 3 emits
generated constants only, and declaring a module it had not written would have
declared a file that did not exist. Step 4 is instructed to declare exactly those
three modules, so the assertion was moved to the boundary that is still ahead —
Phase 6's resolver, checker and reporter, and the calculation endpoints — rather
than removed. The module's test count is unchanged at 57.

---

## Generated-artifact non-regression

Rebuilt from the modified contract and compared byte for byte against the Step-3
build:

| Artifact | Result |
| --- | --- |
| `build/vba/modConstants.bas` | identical |
| `build/vba/modCalcContract.bas` | identical |
| `build/phase4_scenarios.json` | identical |
| `build/phase5_cases.json` | identical |
| `build/PCCM_stageA.xlsx` | every zip member identical except the two that carry the build timestamp (`docProps/core.xml` and the Build sheet cell). No structural change. |
| `build/stage_b_manifest.json` | changed **only** by the four added module inventory entries, and **unchanged by the correction round** |

The Stage-A post-build verifier still reports **351 passed, 0 failed**.

---

## Tests

| Module | Tests |
| --- | --- |
| `test_phase1_manifest_validation.py` | 10 |
| `test_phase1_structure.py` | 21 |
| `test_phase2_contract_validation.py` | 42 |
| `test_phase2_inputs.py` | 40 |
| `test_phase3_driver_contract_validation.py` | 31 |
| `test_phase3_drivers.py` | 28 |
| `test_phase3_verifier_intersection.py` | 12 |
| `test_phase4_oracle.py` | 68 |
| `test_phase4_stage_b_source.py` | 155 |
| `test_phase4_structure.py` | 43 |
| `test_phase4_structure_contract_validation.py` | 55 |
| `test_phase5_calc_contract_validation.py` | 151 |
| `test_phase5_fingerprint.py` | 52 |
| `test_phase5_numeric.py` | 94 |
| `test_phase5_oracle.py` | 111 |
| `test_phase5_stage_a.py` | 57 |
| `test_phase5_vba_source.py` | **112** |
| **Total** | **1082** |

The pre-review baseline was 967; the first Step-4 submission was 1029; the first
correction reached 1063. The second correction adds 19 more static tests, all in
`test_phase5_vba_source.py`. **No test was removed, and no test was weakened to
make the patch pass.**

### Discrimination against the defective commits

Run unchanged against **`7fac269`** (the first submission) the suite produces 23
failures; run unchanged against **`2d76d78`** (the first correction) it produces
14, covering every defect of the second round:

| Defect | Failing tests against `2d76d78` |
| --- | --- |
| ML buffer under-allocation | `test_53`, `test_53a` |
| empty branch after a bounds read, all three producers | `test_62`, `test_63`, `test_64a` |
| count-less private order helpers | `test_64b` |
| unreachable `SafeProduct` empty product | `test_64c`, `test_64d` |
| count-less `SafeSignedSum` boundary | `test_64c`, `test_64d`, `test_64e` |
| unguarded `ExactQuotientOfSum` divisor | `test_64f` |
| `DistributionMean` / `CanonicalOrder` / `CalcFpNumberField` Public | `test_43`, `test_45`, `test_64h`, `test_64i` |

The first round's defect classes and their failing tests against `7fac269`:

| Defect | Failing tests |
| --- | --- |
| wrong cost/risk record schema | `test_48`, `test_49`, `test_50` |
| missing inflation-factor vector | `test_51`, `test_52`, `test_53` |
| caller-selected fingerprint version | `test_54`, `test_55` |
| wrong `TWO_52` / `MAX_SIGNIFICAND` / `MAX_DOUBLE` | `test_56`, `test_57` |
| unused approximate constant | `test_58` |
| whole-object magnitudes clear | `test_59`, `test_60`, `test_61` |
| conditioning failure turned into zero | `test_64`, `test_65`, `test_66` |
| `Variant` numerical containers | `test_06`, `test_68`, `test_69`, `test_70` |

`test_67` (an overflow cannot be relabelled an accepted underflow), `test_64g`
(the convex statistics pass only supported divisors) and `test_53b`/`test_53c`
(the schema positions are distinguishable) are standing invariants rather than
discriminators.

### The thirty negative controls

`tests/test_phase5_vba_source.py` plants twelve defects in synthetic module text
and asserts the sweep that exists to catch each one does catch it. A sweep that
had silently stopped working would pass every positive test and fail these.

**The original twelve:**

1. a host call (`Application.Sum`)
2. a workbook read (`ThisWorkbook.Worksheets(…).Range(…)`)
3. a random draw (`Rnd()`)
4. an Excel object parameter type (`ByRef sheet As Worksheet`)
5. `On Error Resume Next`
6. an error handler outside the four primitives
7. a `PCCM_` endpoint
8. the native `Mod` operator
9. native integer division (`\`)
10. a wide fixed-point type (`Currency` / `CDec`)
11. a hard-coded modulus literal in fingerprint code
12. accidental Public API growth

**Ten more for the review-discovered defects:**

13. `Probability` reintroduced into a cost record
14. the inflation vector dropped
15. the inflation and weight loops swapped
16. a caller-selected fingerprint version
17. a rounded `TWO_52` / `MAX_SIGNIFICAND`
18. a `MAX_DOUBLE` literal above the representable maximum
19. a whole-object magnitude clear
20. an explicit no-driver refusal
21. an exact failure turned into a zero
22. a `Variant` numerical container

**Eight more for the second round:**

23. a capacity formula that forgets Most Likely
24. a missing emitted-count guard
25. a count derived from array bounds, in each of the three producers
26. an empty branch placed after a bounds read
27. an unreachable `SafeProduct` empty-product branch
28. a count-less `SafeSignedSum`
29. an unguarded exact divisor
30. an accidentally Public helper

A further check runs the other way: a comment that mentions `Application.` and
`Worksheets` must **not** be reported, which is what proves the sweeps read code
and not commentary.

### The suite guards its own language

`test_47` reads this test file back and fails if it ever contains a phrase
claiming the VBA ran or that parity was established. The forbidden phrases are
assembled from parts so the guard cannot trip over its own wording.

---

## The public API surface

Whitelisted **exactly**, in both directions, so accidental growth is a test
failure rather than a silent new entry point. Everything else in each module is
`Private`.

**`modCalcFactors` (17)** — `IsUsableDouble`, `SafeAdd`, `SafeSubtract`,
`SafeMultiply`, `SafeDivide`, `SafeAccumulate`, `SafeSignedSum`, `SafeProduct`,
`ExactSumOfProducts`, `ExactQuotientOfSum`, `BuildInflationFactors`,
`BuildDiscountFactors`, `BuildKnom`, `BuildKpv`, `ConditioningScaledMagnitude`,
`ConditioningScaledProduct`, `IdentityAllowance`.

**`modCalcAnalytical` (10)** — `TriangularMean`, `PertMean`, `UniformMean`,
`DeterministicCentral`, `ExpectedRisk`, `BuildDriverAudit`, `AccumulateTotals`,
`BuildAnnualSeries`, `Reconcile`, `AllIdentitiesHold`.

**`modCalcFingerprint` (10)** — `CalcFpUtf16Length`, `CalcFpNormaliseCodeUnit`,
`CalcFpCanonicalText`, `CalcFpCanonicalNumber`, `CalcFpCanonicalInteger`,
`CalcFpReduceDouble`, `CalcFpDigestStream`, `CalcFpBuildCostRecord`,
`CalcFpBuildRiskRecord`, `CalcFpBuildFingerprint`.

`DistributionMean`, `CanonicalOrder` and `CalcFpNumberField` were Public in
`2d76d78` with no cross-module caller. They are now `Private`.

Two `modCalcFactors` procedures are Public specifically because
`modCalcAnalytical` calls them, and this is why:

* **`ExactQuotientOfSum`** — the convex statistics' exact tier needs
  `(sum of terms) / divisor` for divisor in `{2, 3, 6}`, rounded once. That
  rounding is the exact kernel's business. Reimplementing it in the analytical
  layer would be a second rounding rule, and two rounding rules are two chances
  to disagree.
* **`ConditioningScaledProduct`** — a per-driver-per-year annual contribution can
  have no Double of its own while `coefficient * |contribution|` is perfectly
  finite. Only the exact kernel can fold the coefficient into that factor
  expression, and C1 requires the magnitude to be recorded anyway.

`CalcFpBuildVersionedFingerprint` is **Private**: the version is injectable only
so a future migration can express the grammar for a version other than the
current one, and production has exactly one entry point, which reads
`FP_VERSION` from `modCalcContract`.

`test_64h` scans cross-module references and fails any Public name that has no
external caller and no entry in the documented exception set.

---

## Provenance

Direction of authority is unchanged: the accepted plan and the tested Python
oracle define the semantics, and the VBA implements them. The correction round
changed VBA source and static tests only. `calc_fingerprint.py`,
`calc_numeric.py`, `calc_oracle.py`, `calc_cases.py` and `calc_contract.yaml`
are untouched, and the reference digest `50B6EB0E26857EA7` over 366 UTF-16 code
units is unchanged — `test_53b` recomputes it under the corrected schema. Nothing in the Python
oracle, `calc_contract.yaml`, `calc_fingerprint.py`, `calc_numeric.py`,
`calc_oracle.py`, `calc_cases.py` or `build/phase5_cases.json` was changed to make
the VBA easier to write.
