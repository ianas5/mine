# Python reference implementation ("oracle")

Independent Python implementations of the numerical routines the VBA engine must
reproduce: the RNG, the seed derivation, Triangular / Beta-PERT / Uniform
sampling, the percentile method and mid-rank Spearman correlation.

These generate the fixture vectors in `../fixtures/`, which the in-workbook VBA
test harness asserts against. The oracle is deliberately a *separate*
implementation so the VBA is checked against something other than itself.

**Empty in Phase 1.** No numerical work has started.
