# PCCM Phase-6 Step-0 evidence package

Retained, reproducible evidence for the Phase-6 Step-0 decisions. Every number
quoted in `pccm/docs/phase6_step0.md` is produced by the scripts in this
directory from the inputs in this directory, and is recorded here as data.

**Nothing in this package is runtime evidence.** No Windows machine and no Excel
instance was involved in producing any file here. This is not a VBA benchmark,
not a timing measurement, and it makes no claim about how fast anything runs in
Excel.

## Reproducing it

```
cd pccm/evidence/phase6_step0
python3 scripts/run_all.py
```

One command, no arguments, no network, no third-party packages. It rewrites
`raw/`, `vectors/`, `summaries/` and `manifest.json`. Two consecutive runs
produce a byte-identical `manifest.json`; that is the determinism check, and it
covers every generated file because the manifest hashes all of them.

`scripts/controls.py` is also runnable on its own and prints the same control
results that `run_all.py` records into `summaries/`.

Requires Python 3.11 (see `environment.txt` for the exact interpreter used).

## Layout

| Path | What it holds |
|---|---|
| `scripts/` | the generators. Evidence code, not production code — see below |
| `inputs/` | every input value: the parameter grid, the probe points, the seeds |
| `raw/` | measured results, one file per subject |
| `vectors/` | cross-language vectors: an implementation in any language can be checked against these |
| `summaries/` | negative-control results and the generating environment |
| `manifest.json` | SHA-256 of every file above, plus the baselines and provenance |
| `environment.txt` | the interpreter, platform and float characteristics |

## What each `raw/` file answers

| File | Subject |
|---|---|
| `rng_backbone.json` | MRG32k3a, two independent arithmetic paths |
| `cheng_formulation.json` | the LOCKED Cheng formulation, enumerated and hashed to its source |
| `degenerate_d6_18.json` | degenerate severity under D6-18b: invoked every iteration, zero uniforms, stream unchanged |
| `version_register.json` | RNG_VERSION, SIM_METHOD_VERSION, the `PCCM-RD` version field, and what bumps each |
| `tolerance_model.json` | measured expression-order and libm sensitivity, and the comparison-tolerance policy derived from it |
| `jump.json` | the `2^127` jump matrices and the safe modular arithmetic |
| `seed_map.json` | D6-03 auto-seed cycle, D6-05 scalar → six-word state |
| `cheng.json` | D6-04 candidate B: Cheng BB/BC, measured per shape |
| `candidate_a_inverse_cdf.json` | D6-04 candidate A1: naive inverse CDF |
| `candidate_a2_newton.json` | D6-04 candidate A2: the same family, implemented competently |
| `candidate_c_table.json` | D6-04 candidate C: table + linear interpolation, error/memory curve |
| `d6_04_comparison.json` | all four, one counting method, one workload |
| `d6_18_operation_model.json` | D6-18 options A and B at four probabilities |
| `stream_architecture.json` | D6-16 stream assignment, families A and B |
| `digest.json` | D6-17 `result_digest` |
| `run_log.txt` | what the last generation run did |

`vectors/cheng_vectors.json` holds the deterministic sampler vectors that define
the locked formulation: five shapes covering both dispatches, 24 samples each,
every case exercising both an immediate acceptance and at least one retry, with
per-sample attempts, uniforms, cumulative uniforms and post-sample RNG state.

## Two things this package deliberately does NOT do

**It does not become production code.** `scripts/` is evidence. It is not
imported by `pccm/builder`, `pccm/src` or `pccm/tests`, nothing in the production
trees references it, and no file in it is offered as a Phase-6 implementation.
The dependency runs one way only: `scripts/digest.py` and `scripts/stream_map.py`
*import* the accepted Phase-5 encoder from `pccm/builder/pccm_builder/`, read
only, so that D6-16 and D6-17 demonstrably reuse accepted authority instead of
reimplementing it. No production file is modified.

**It does not create production test semantics.** `scripts/controls.py` holds
negative controls for the *evidence*: each plants a defect this package is
supposed to detect and asserts the detection fires. They are not Phase-6 tests,
they live outside `pccm/tests`, and they are not run by the project's pytest
suite.
