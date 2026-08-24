"""Step-0 driver. Produces raw/, summaries/ and vectors/ deterministically."""
import json, math, os, statistics, sys, hashlib, platform
sys.path.insert(0, os.path.dirname(__file__))
from mrg32k3a import (M1, M2, A1, A2, ExactMrg, DoubleMrg, mat_pow2_mod,
                      jump_state, verify_convention, stream_states,
                      mat_vec_mod_safe, mat_vec_mod_naive)
import beta_ref as BR
import transforms as TX
import seed_map as SM
import stream_map as STM
from digest import (result_digest, result_stream, canon_double,
                    FP_BASE, FP_MOD_1, FP_MOD_2, FP_INIT_1, FP_INIT_2,
                    RESULT_STREAM_TAG, RESULT_SECTION)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
def w(sub, name, obj):
    p = os.path.join(ROOT, sub, name)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2, sort_keys=True)
    return p

GRID = json.load(open(os.path.join(ROOT, "inputs", "pert_grid.json")))
SEEDS = json.load(open(os.path.join(ROOT, "inputs", "seeds.json")))
SEED_STATE = [12345] * 6

# ============================ 1. RNG backbone ============================
def rng_backbone():
    a, b = ExactMrg(list(SEED_STATE)), DoubleMrg(list(SEED_STATE))
    ua = [a.next_u() for _ in range(20)]
    ub = [b.next_u() for _ in range(20)]
    out = {"seed_state": SEED_STATE,
           "exact_equals_double_path": ua == ub,
           "first_20_uniforms": [repr(x) for x in ua],
           "state_after_20": a.s,
           "u_strictly_between_0_and_1": all(0.0 < x < 1.0 for x in ua)}
    per_seed = {}
    for s in SEEDS["fixed_seeds"]:
        g = ExactMrg([s] * 6)
        per_seed[str(s)] = {"initial_state": [s] * 6,
                            "first_5": [repr(g.next_u()) for _ in range(5)]}
    out["per_seed"] = per_seed
    return out

# ============================ 2. jump-ahead =============================
def jump_evidence():
    J1 = mat_pow2_mod(A1, 127, M1)
    J2 = mat_pow2_mod(A2, 127, M2)
    small1, small2 = mat_pow2_mod(A1, 3, M1), mat_pow2_mod(A2, 3, M2)
    g = ExactMrg(list(SEED_STATE))
    for _ in range(8):
        g.next_u()
    wanted = [0, 1, 7, 399, 401]
    states = stream_states(SEED_STATE, max(wanted) + 1, J1, J2)
    naive_ok = all(jump_state(states[k], J1, J2) ==
                   jump_state(states[k], J1, J2, safe=False) for k in wanted)
    worst = max(J1[i][j] * (M1 - 1) for i in range(3) for j in range(3))
    worst_sum = max(sum(J1[i][j] * (M1 - 1) for j in range(3)) for i in range(3))
    def mhash(m):
        text = ";".join(",".join(str(v) for v in row) for row in m)
        return {"canonical_text": text,
                "sha256": hashlib.sha256(text.encode("ascii")).hexdigest()}
    return {"provenance":
                "A1p127 and A2p127 are DERIVED here by repeated squaring of the "
                "transition matrices A1 and A2, which are written directly from "
                "the MRG32k3a recurrence in mrg32k3a.py. They are not copied "
                "from any published table and no published table is treated as "
                "evidence. Correctness is established by the three checks below, "
                "not by resemblance to remembered literature values.",
            "A1p127_derived_from_recurrence": J1,
            "A2p127_derived_from_recurrence": J2,
            "state_vector_convention": "A1/A2 act on NEWEST-FIRST [x_{n-1},x_{n-2},x_{n-3}]; "
                                       "PCCM stores OLDEST-FIRST [s10,s11,s12]",
            "one_step_matrix_equals_recurrence": verify_convention(SEED_STATE),
            "eight_step_jump_equals_eight_steps": g.s == jump_state(SEED_STATE, small1, small2),
            "safe_multmodm_equals_exact_integer": naive_ok,
            "worst_naive_term": worst,
            "worst_naive_term_over_2_53": worst / 2**53,
            "worst_naive_row_sum": worst_sum,
            "worst_naive_row_sum_over_2_53": worst_sum / 2**53,
            "naive_overflow_note":
                "The plan's 2048x is the bound for the largest REPRESENTABLE "
                "entry. These are the actual worst values for the actual derived "
                "J1 against a maximal state vector: a single product, and the "
                "three-term row sum the naive form would accumulate before "
                "reducing. Both overflow 2^53 by three orders of magnitude, so "
                "the conclusion is unchanged and now rests on measured values.",
            "A1p127_hash": mhash(J1), "A2p127_hash": mhash(J2),
            "base_state": list(SEED_STATE),
            "streams_recorded": wanted,
            "stream_initial_states": {str(k): states[k] for k in wanted}}

# ============================ 3. Cheng ==================================
def cheng_measurements(n=20000):
    rows = []
    for r in GRID["r_values"]:
        a, b = BR.pert_shape(r)
        g = ExactMrg([SEEDS["cheng_stream_seed"]] * 6)
        counter = {"attempts": 0, "uniforms": 0}
        BR.tr_reset()
        BR.margin_reset()
        xs = []
        for _ in range(n):
            before = counter["uniforms"]
            xs.append(BR.cheng_sample(a, b, g.next_u, counter))
            rows_u = counter["uniforms"] - before
            xs[-1] = (xs[-1], rows_u)
        cons = [c for _, c in xs]
        vals = [v for v, _ in xs]
        cons.sort()
        mean_t = a / (a + b)
        var_t = a * b / ((a + b) ** 2 * (a + b + 1.0))
        rows.append({
            "r": r, "alpha": a, "beta": b, "dispatch": BR.cheng_dispatch(a, b),
            "samples_requested": n, "accepted_samples": n,
            "samples": n, "proposal_attempts": counter["attempts"],
            "attempts": counter["attempts"],
            "acceptance_rate": n / counter["attempts"],
            "uniforms_total": counter["uniforms"],
            "uniforms_per_accepted_mean": counter["uniforms"] / n,
            "uniforms_p50": cons[n // 2], "uniforms_p90": cons[int(n * 0.90)],
            "uniforms_p95": cons[int(n * 0.95)], "uniforms_p99": cons[int(n * 0.99)],
            "uniforms_max": cons[-1],
            "acceptance_path_margins": BR.margin_snapshot(),
            "transcendental_calls_total": BR.tr_total(),
            "transcendental_calls_per_sample": BR.tr_total() / n,
            "transcendental_breakdown": BR.tr_snapshot(),
            "mean_theoretical": mean_t, "mean_observed": statistics.fmean(vals),
            "mean_abs_error": abs(statistics.fmean(vals) - mean_t),
            "var_theoretical": var_t, "var_observed": statistics.pvariance(vals),
            "var_abs_error": abs(statistics.pvariance(vals) - var_t)})
    return rows

# ==================== 4. candidate A: inverse CDF =======================
def candidate_a():
    """A1: the NAIVE inverse CDF - 60 unconditional bisection halvings, and a
    fresh lgamma triple inside every evaluation. Retained because it is the
    implementation the earlier desk estimate described; it is NOT the family's
    best form. See candidate_a2()."""
    rows = []
    for r in GRID["r_values"]:
        a, b = BR.pert_shape(r)
        worst_it, worst_ev, worst_tr = 0, 0, 0
        for u in GRID["u_probes"]:
            BR.tr_reset()
            _, evals, worst = BR.beta_ppf(u, a, b)
            worst_tr = max(worst_tr, BR.tr_total())
            worst_it = max(worst_it, worst)
            worst_ev = max(worst_ev, evals)
        rows.append({"r": r, "alpha": a, "beta": b,
                     "betai_evaluations_per_sample": worst_ev,
                     "worst_continued_fraction_iterations": worst_it,
                     "transcendental_calls_per_sample": worst_tr})
    return rows


def candidate_a2():
    """A2: the same family, implemented competently. Safeguarded Newton on the
    regularised incomplete beta, with log B(a,b) hoisted to per-DRIVER setup.
    Accuracy is measured against A1, which is an independent method."""
    rows = []
    for r in GRID["r_values"]:
        a, b = BR.pert_shape(r)
        BR.tr_reset()
        lbeta = BR.log_beta(a, b)
        setup_tr = BR.tr_total()
        worst_it, worst_cf, worst_tr = 0, 0, 0
        worst_err, worst_err_at = 0.0, None
        for u in GRID["u_probes"]:
            BR.tr_reset()
            x, it, cf = BR.beta_ppf_newton(u, a, b, lbeta)
            worst_tr = max(worst_tr, BR.tr_total())
            worst_it = max(worst_it, it)
            worst_cf = max(worst_cf, cf)
            err = abs(x - BR.beta_ppf(u, a, b)[0])
            if err > worst_err:
                worst_err, worst_err_at = err, u
        rows.append({"r": r, "alpha": a, "beta": b,
                     "per_driver_setup_transcendentals": setup_tr,
                     "worst_newton_iterations": worst_it,
                     "worst_continued_fraction_iterations": worst_cf,
                     "transcendental_calls_per_sample": worst_tr,
                     "max_abs_error_vs_bisection": worst_err,
                     "max_abs_error_at_u": worst_err_at,
                     "uniforms_per_sample": 1})
    return rows

# ==================== 5. candidate C: table =============================
def candidate_c():
    rows = []
    # A SUBSET of shapes, stated: the tail shapes are where a uniform node grid
    # fails, so r=0 and r=1 must be present; the interior is sampled to show the
    # error is not an artefact of the extremes alone.
    shapes = GRID["table_r_values"]
    for nodes in GRID["table_nodes"]:
        worst, where = 0.0, None
        for r in shapes:
            a, b = BR.pert_shape(r)
            grid = [i / nodes for i in range(nodes + 1)]
            xs = [BR.beta_ppf(gv, a, b)[0] if 0.0 < gv < 1.0 else (0.0 if gv == 0 else 1.0)
                  for gv in grid]
            for u in GRID["u_probes"]:
                j = min(int(u * nodes), nodes - 1)
                t = (u - grid[j]) * nodes
                approx = xs[j] + t * (xs[j + 1] - xs[j])
                err = abs(approx - BR.beta_ppf(u, a, b)[0])
                if err > worst:
                    worst, where = err, {"r": r, "u": u}
        rows.append({"nodes": nodes,
                     "transcendental_calls_per_sample": 0,
                     "uniforms_per_sample": 1,
                     "build_inverse_cdf_evaluations_per_driver": nodes + 1,
                     "max_abs_error_normalised": worst,
                     "worst_at": where, "shapes_measured": len(GRID["table_r_values"]),
                     "memory_bytes_per_driver": (nodes + 1) * 8,
                     "memory_bytes_300_drivers": (nodes + 1) * 8 * 300})
    return rows

# ========= 4b. acceptance-path fragility model (section 5.3) =============
def acceptance_margin_model(cheng_rows, n=20000):
    """How close do Cheng's branch decisions come to flipping?

    Measured: 805k+ predicate evaluations, every one recorded. Extrapolated:
    the expected rate at the design target, from the measured near-boundary
    DENSITY. The extrapolation is stated as an extrapolation. It concerns
    Python Doubles perturbed by one ULP; it is NOT a statement about VBA, and
    no cross-language conclusion follows from it.
    """
    ulp_rel = 2.0 ** -52
    beta_samples = (200 + 100) * 100000
    rows, worst_rel, worst_abs = [], float("inf"), float("inf")
    total_eval = 0
    dens = []
    for r in cheng_rows:
        m = r["acceptance_path_margins"]
        per_sample = m["evaluated"] / n
        # Empirical density of relative margins near zero, per unit margin.
        density = m["rel_lt_1e_3"] / (m["evaluated"] * 1e-3) if m["evaluated"] else 0.0
        evals_per_run = per_sample * beta_samples
        rows.append({"r": r["r"], "dispatch": r["dispatch"],
                     "predicate_evaluations": m["evaluated"],
                     "predicate_evaluations_per_sample": per_sample,
                     "min_abs_margin": m["min_abs"], "min_rel_margin": m["min_rel"],
                     "count_rel_below_1e_3": m["rel_lt_1e_3"],
                     "count_rel_below_1e_6": m["rel_lt_1e_6"],
                     "count_rel_below_1e_9": m["rel_lt_1e_9"],
                     "count_rel_below_1e_12": m["rel_lt_1e_12"],
                     "count_rel_below_1e_15": m["rel_lt_1e_15"],
                     "near_zero_density_per_unit_rel_margin": density,
                     "expected_one_ulp_flips_per_design_target_run":
                         density * ulp_rel * evals_per_run})
        worst_rel = min(worst_rel, m["min_rel"])
        worst_abs = min(worst_abs, m["min_abs"])
        total_eval += m["evaluated"]
        dens.append(density)
    return {"one_ulp_relative": ulp_rel,
            "total_predicate_evaluations_measured": total_eval,
            "closest_observed_relative_margin": worst_rel,
            "closest_observed_absolute_margin": worst_abs,
            "closest_margin_over_one_ulp": worst_rel / ulp_rel,
            "extrapolation_assumptions": [
                "the relative-margin distribution has finite, roughly constant "
                "density in a neighbourhood of zero (supported by the measured "
                "counts below 1e-3, 1e-6 and 1e-9, but not proved)",
                "the measured density transfers to the design-target workload",
                "a perturbation of exactly one ULP is what would have to flip a "
                "branch",
                "this models PYTHON Double arithmetic only"],
            "not_a_cross_language_claim":
                "Step 0 has no VBA runtime and no Phase-6 VBA implementation. "
                "Nothing here shows that a VBA Cheng path takes the same "
                "branches. Revision 6's Layer-G rule stands unchanged.",
            "per_shape": rows}

# ============== 5b. D6-04 comparison at the design target ===============
def d6_04_comparison(cheng_rows, a_rows, a2_rows, c_rows):
    """One counting method, one workload, four implementations.

    Workload = the accepted design target's WORST CASE: 200 Cost Lines and 100
    Risks, all Beta-PERT, 25 years, 100,000 iterations, every risk at p = 1.
    That is 3.0e7 Beta samples and 1.0e7 occurrence uniforms.
    """
    beta_samples = (200 + 100) * 100000
    occurrence_uniforms = 100 * 100000
    drivers = 300

    worst_cheng_tr = max(r["transcendental_calls_per_sample"] for r in cheng_rows)
    worst_cheng_u = max(r["uniforms_per_accepted_mean"] for r in cheng_rows)
    worst_a1 = max(r["transcendental_calls_per_sample"] for r in a_rows)
    worst_a2 = max(r["transcendental_calls_per_sample"] for r in a2_rows)
    a2_setup = max(r["per_driver_setup_transcendentals"] for r in a2_rows)
    c_4096 = next(r for r in c_rows if r["nodes"] == 4096)

    def row(name, tr_per_sample, u_per_sample, build_tr, mem, acc, acc_basis):
        return {"candidate": name,
                "transcendental_calls_per_sample": tr_per_sample,
                "transcendental_calls_per_run": tr_per_sample * beta_samples + build_tr,
                "one_off_setup_transcendental_calls": build_tr,
                "uniforms_per_run": u_per_sample * beta_samples + occurrence_uniforms,
                "extra_resident_bytes": mem,
                "worst_sampling_error": acc,
                "error_basis": acc_basis}

    return {
        "workload": {"beta_samples": beta_samples,
                     "occurrence_uniforms": occurrence_uniforms,
                     "drivers": drivers,
                     "assumption": "worst case: every driver Beta-PERT, every "
                                   "risk Probability = 1"},
        "counting_method": "transcendental calls counted by beta_ref.TR, the SAME "
                           "instrumented functions in every candidate. No "
                           "flops-per-iteration multiplier is applied.",
        "rows": [
            row("A1 naive inverse CDF (60 bisection halvings)",
                worst_a1, 1, 0, 0, 0.0,
                "exact to the bisection tolerance; this row IS the accuracy reference"),
            row("A2 safeguarded Newton inverse CDF",
                worst_a2, 1, a2_setup * drivers, 0,
                max(r["max_abs_error_vs_bisection"] for r in a2_rows),
                "max abs difference from A1 over the probe grid"),
            row("B Cheng BB/BC rejection",
                worst_cheng_tr, worst_cheng_u, 0, 0, None,
                "no per-sample error: the sampler is exact in distribution; "
                "correctness shown by theoretical mean/variance agreement"),
            row("C table + linear interpolation, 4096 nodes",
                0, 1,
                c_4096["build_inverse_cdf_evaluations_per_driver"] * worst_a2 * drivers,
                c_4096["memory_bytes_300_drivers"],
                c_4096["max_abs_error_normalised"],
                "max abs normalised interpolation error over the measured shapes"),
        ]}

# ==================== 6. D6-18 operation model ==========================
def d6_18_model(cheng_rows, degenerate_res):
    """Both options, at four probabilities, using the MEASURED Cheng behaviour.

    TERMINOLOGY, corrected. Under Option A the severity sampler is entered only
    when the Bernoulli occurrence test passes, so its invocation count is a
    RANDOM VARIABLE. `N * p` is its EXPECTATION, not an observed or guaranteed
    count, and it is exact only at p = 0 and p = 1 where the Bernoulli is
    deterministic. Every Option-A figure below is therefore labelled `expected_`.
    Under Option B the sampler is entered every iteration, so its invocation
    count is exactly N for every p - not an expectation.

    The measured realised counts in raw/degenerate_d6_18.json make the
    distinction concrete: at p = 0.1 over 1000 iterations the realised
    occurrence count was 93, not 100.
    """
    by_r = {row["r"]: row for row in cheng_rows}
    mid_u = by_r[0.5]["uniforms_per_accepted_mean"]
    worst_u = max(r["uniforms_per_accepted_mean"] for r in cheng_rows)
    worst_tr = max(r["transcendental_calls_per_sample"] for r in cheng_rows)
    risks, iters = 100, 100000
    stream_states_bytes = 400 * 6 * 8
    realised = {r["probability"]: r for r in degenerate_res["rows"]
                if not r["degenerate"]}
    out = []
    for p in GRID["probabilities"]:
        expected_occ = risks * iters * p
        both = risks * iters
        row = {
            "probability": p,
            "occurrence_uniforms": both,
            "occurrence_uniforms_is_exact": True,

            "option_A_expected_severity_invocations": expected_occ,
            "option_A_invocation_count_is_random": p not in (0.0, 1.0),
            "option_A_expected_severity_uniforms_mid_shape": expected_occ * mid_u,
            "option_A_expected_severity_uniforms_worst_shape": expected_occ * worst_u,
            "option_A_expected_severity_transcendentals": expected_occ * worst_tr,

            "option_B_severity_invocations_exact": both,
            "option_B_invocation_count_is_random": False,
            "option_B_severity_uniforms_est_mid_shape": both * mid_u,
            "option_B_severity_uniforms_est_worst_shape": both * worst_u,
            "option_B_severity_transcendentals_est": both * worst_tr,

            "option_A_stream_state_bytes": stream_states_bytes,
            "option_B_stream_state_bytes": stream_states_bytes,
            "memory_difference_bytes": 0,
            "uniforms_per_accepted_mid_shape": mid_u,
            "uniforms_per_accepted_worst_shape": worst_u,
            "transcendentals_per_sample_worst_shape": worst_tr}
        r = realised.get(p)
        if r:
            row["measured_realised_occurrences_at_1000_iterations"] = \
                r["realised_occurrences"]
            row["measured_expected_occurrences_at_1000_iterations"] = \
                r["expected_occurrences"]
        out.append(row)
    return {
        "terminology":
            "Option A: N * p is the EXPECTED severity-sampler invocation count. "
            "The realised count is the Bernoulli occurrence count and is random "
            "except at p = 0 and p = 1. Option B: exactly N invocations for every "
            "p, not an expectation. Every Option-A uniform and transcendental "
            "figure below is an EXPECTED work estimate derived from that "
            "expectation, not an observed count.",
        "option_B_severity_uniform_count_is_p_independent":
            "Measured: the non-degenerate rows of raw/degenerate_d6_18.json "
            "consume 2174 severity uniforms at EVERY probability, including "
            "p = 0. That is the property Option B exists to provide.",
        "rows": out}

# ==================== 7. D6-03 / D6-05 seed mapping =====================
def seed_evidence():
    order = SM.P - 1
    complete = SM.factorise_with_multiplicity(order)
    distinct = SM.distinct_prime_divisors(order)
    # The two are DIFFERENT facts and conflating them is an arithmetic error:
    # the product of the distinct divisors is 715827882, not 2147483646.
    complete_product = 1
    for q, e in complete:
        complete_product *= q ** e
    distinct_product = 1
    for q in distinct:
        distinct_product *= q
    prim = SM.is_primitive_root(SM.MULT, SM.P)
    # Full period == exactly the admissible seed domain 1..2147483646.
    period_equals_domain = (order == SM.SEED_MAX - SM.SEED_MIN + 1)
    # Cycle closure at the period: x_{period} == x_0 for any non-zero start.
    closes = pow(SM.MULT, order, SM.P) == 1
    nonces = [0, 1, 2, 3, 10, 1000, order - 1, order]
    pairs = [{"auto_nonce": n, "effective_seed": SM.nonce_to_seed(n)} for n in nonces[:6]]
    # The two endpoints computed by fast exponentiation, not by iterating.
    for n in nonces[6:]:
        pairs.append({"auto_nonce": n,
                      "effective_seed": pow(SM.MULT, n, SM.P)})
    seen = set()
    collision_free_prefix = True
    for k in range(1, 200001):
        v = pow(SM.MULT, k, SM.P)
        if v in seen:
            collision_free_prefix = False
            break
        seen.add(v)
    # D6-05: the repeated-scalar state is valid for every admissible seed.
    # One inequality decides it for the whole domain.
    d6_05 = {"seed_max": SM.SEED_MAX, "m2": M2, "m1": M1,
             "seed_max_lt_m2": SM.SEED_MAX < M2,
             "m2_lt_m1": M2 < M1,
             "therefore_every_admissible_seed_is_a_valid_residue_in_both_components": True,
             "seed_min_state_valid": SM.state_is_valid(SM.seed_to_state(SM.SEED_MIN)),
             "seed_max_state_valid": SM.state_is_valid(SM.seed_to_state(SM.SEED_MAX)),
             "seed_zero_state_valid": SM.state_is_valid(SM.seed_to_state(0))}
    # Agreement between the power authority and the stepped cycle, over a prefix.
    paths_agree = all(SM.nonce_to_seed(n) == SM.nonce_to_seed_iterative(n)
                      for n in range(0, 3000))
    lifecycle = {
        "auto_nonce_meaning": "THE NEXT NONCE TO ALLOCATE. Not the last one used.",
        "initial_persisted_value": SM.NONCE_INITIAL,
        "allocation":
            "On an AUTO run, at the point the effective seed is derived and "
            "BEFORE any sampling: read the persisted auto_nonce, compute "
            "effective_seed = 48271^auto_nonce mod 2147483647, then persist "
            "auto_nonce + 1. Read-then-advance, in that order.",
        "advance_point":
            "Immediately at allocation, not at successful commit. This is what "
            "makes an allocated seed consumed.",
        "failure_before_allocation":
            "Consumes NO nonce. Nothing was allocated, so nothing is spent - a "
            "refusal on input validation leaves auto_nonce untouched.",
        "failure_after_allocation":
            "CONSUMES the nonce. The seed is spent even though the simulation "
            "failed, so a retry receives a NEW sequence rather than replaying "
            "the failed one. This is the inherited behaviour and it is a "
            "property of the ALLOCATION ORDER, which is why the order is stated "
            "above rather than left to the implementation.",
        "attempt_metadata":
            "A failed AUTO attempt records the allocated effective_seed and the "
            "auto_nonce it consumed, alongside the failure. The prior successful "
            "result is untouched, but the failed allocation is visible - "
            "otherwise a consumed nonce would leave no trace and the sequence "
            "would appear to skip.",
        "exhaustion":
            "auto_nonce == 2147483646 must NOT be allocated: 48271^2147483646 "
            "== 1 == the seed for nonce 0, so allocating it silently reissues "
            "the first seed. Refuse the run.",
        "authority_vs_implementation":
            "The AUTHORITY is effective_seed = 48271^auto_nonce mod 2147483647 - "
            "a modular power. A later VBA implementation must use an exact "
            "O(log nonce) modular exponentiation; O(nonce) repeated "
            "multiplication is not acceptable at large nonce and is not what "
            "the authority says. No implementation is proposed here.",
        "both_paths_agree_over_first_3000_nonces": paths_agree,
    }
    return {"modulus": SM.P, "multiplier": SM.MULT,
            "order_complete_factorisation": [list(t) for t in complete],
            "order_complete_factorisation_text": "2147483646 = "
                + " * ".join(f"{q}^{e}" if e > 1 else str(q) for q, e in complete),
            "order_complete_factorisation_product": complete_product,
            "order_complete_factorisation_product_correct": complete_product == order,
            "order_distinct_prime_divisors": distinct,
            "order_distinct_prime_divisors_product": distinct_product,
            "distinct_product_is_not_the_order": distinct_product != order,
            "primitive_root_test_uses":
                "the DISTINCT prime divisors only - multiplicity is irrelevant to "
                "the test g^((p-1)/q) != 1, so the proof is unaffected by the "
                "earlier wording error",
            "multiplier_is_primitive_root": prim,
            "nonce_lifecycle": lifecycle,
            "period": order,
            "seed_domain": [SM.SEED_MIN, SM.SEED_MAX],
            "period_equals_seed_domain_size": period_equals_domain,
            "cycle_closes_at_period": closes,
            "no_repeat_in_first_200000_nonces": collision_free_prefix,
            "nonce_to_seed_pairs": pairs,
            "d6_05_repeated_scalar": d6_05}

# ==================== 8. D6-17 result digest ============================
def digest_evidence():
    base_nom = [1234567.8901234567, 0.0, -0.5, 1e300, 5e-324]
    base_pv = [1000000.0, 0.0, -0.25, 1e299, 1e-320]
    cases = []
    def add(label, nom, pv, version=1):
        cases.append({"label": label, "totals_nominal": [repr(v) for v in nom],
                      "totals_pv": [repr(v) for v in pv], "version": version,
                      "stream": result_stream(nom, pv, version),
                      "digest": result_digest(nom, pv, version)})
    add("base", base_nom, base_pv)
    add("reversed_iteration_order", base_nom[::-1], base_pv[::-1])
    add("nominal_and_pv_swapped", base_pv, base_nom)
    add("one_iteration_dropped", base_nom[:-1], base_pv[:-1])
    add("one_ulp_perturbation",
        [math.nextafter(base_nom[0], math.inf)] + base_nom[1:], base_pv)
    add("version_2", base_nom, base_pv, 2)
    add("empty", [], [])
    digs = [c["digest"] for c in cases]
    return {"hash_constants": {"FP_BASE": FP_BASE, "FP_MOD_1": FP_MOD_1,
                               "FP_MOD_2": FP_MOD_2, "FP_INIT_1": FP_INIT_1,
                               "FP_INIT_2": FP_INIT_2},
            "source_of_constants":
                "IMPORTED from pccm/builder/pccm_builder/calc_fingerprint.py, the "
                "accepted Phase-5 single source. No new hash, number format or "
                "folding is introduced by D6-17.",
            "stream_tag": RESULT_STREAM_TAG, "section_name": RESULT_SECTION,
            "grammar": 'stream ::= F_S("PCCM-RD") F_I(version) '
                       'section("RESULT", record*); '
                       "record ::= F_I(field_count) F_I(iteration_index) "
                       "F_N(nominal) F_N(pv)",
            "digest_is_16_hex_chars": all(len(d) == 16 for d in digs),
            "all_distinct": len(set(digs)) == len(digs),
            "canonical_double_examples": {repr(v): canon_double(v)
                                          for v in [0.0, -0.0, 1.0, 0.1,
                                                    1234567.8901234567,
                                                    1e300, 5e-324, -0.5]},
            "cases": cases}

# ==================== 9. D6-16 stream architecture ======================
def stream_architecture():
    import random
    n_cost, n_risk = 200, 100
    cost_ids = [f"CL-{i:03d}" for i in range(1, n_cost + 1)]
    risk_ids = [f"R-{i:03d}" for i in range(1, n_risk + 1)]
    comps = STM.components(cost_ids, risk_ids)
    a_map = STM.family_a(cost_ids, risk_ids)

    rnd = random.Random(20260824)
    shuffled_cost = cost_ids[:]; shuffled_risk = risk_ids[:]
    rnd.shuffle(shuffled_cost); rnd.shuffle(shuffled_risk)
    a_shuf = STM.family_a(shuffled_cost, shuffled_risk)

    idxs = sorted(a_map.values())
    kinds = {"cost_components": n_cost, "risk_components": 2 * n_risk,
             "total_components": len(comps)}

    # Ordinal order is NOT numeric order once IDs widen. Recorded, not hidden.
    widened = ["CL-001", "CL-002", "CL-999", "CL-1000", "CL-0001"]
    widened_sorted = sorted(widened, key=STM.utf16_sort_key)

    # Family B under the ACCEPTED, UNBOUNDED pattern.
    witnesses = []
    for k in [1000, 100000, 2 ** 31 - 1]:
        wt = STM.family_b_collision_witness(k)
        witnesses.append({"proposed_risk_offset": k,
                          "colliding_cost_id": wt[0] if wt else None,
                          "colliding_risk_id": wt[1] if wt else None,
                          "shared_stream_index": wt[2] if wt else None,
                          "collision_exists": wt is not None})

    return {
        "design_target": kinds,
        "component_rule": "Cost Line -> 1 stream; Risk -> 2 streams (occurrence, severity)",
        "family_a": {
            "assignment_is_a_bijection_onto_0_to_n_minus_1":
                idxs == list(range(len(comps))),
            "row_order_invariant": a_map == a_shuf,
            "all_ids_ascii_so_utf16_ordinal_equals_codepoint_order":
                all(STM.is_ascii_id(i) for i in cost_ids + risk_ids),
            "first_10_streams": [{"component": list(c), "stream": a_map[c]}
                                 for c in sorted(a_map, key=lambda c: a_map[c])[:10]],
            "last_4_streams": [{"component": list(c), "stream": a_map[c]}
                               for c in sorted(a_map, key=lambda c: a_map[c])[-4:]],
            "ordinal_order_is_not_numeric_order": {
                "input": widened, "ordinal_sorted": widened_sorted,
                "note": "CL-1000 sorts before CL-999 under ordinal comparison. "
                        "Deterministic and portable, but not numeric."}},
        "family_b_under_accepted_unbounded_id_pattern": {
            "accepted_cost_pattern": STM.COST_ID.pattern,
            "accepted_risk_pattern": STM.RISK_ID.pattern,
            "numeric_part_is_bounded_above": False,
            "collision_witnesses": witnesses,
            "conclusion": "No finite kind_offset is collision-free over the "
                          "representable ID domain, because the numeric part is "
                          "unbounded. Family B needs a bound the accepted "
                          "contract does not provide."},
        "stream_state_memory_bytes": len(comps) * 6 * 8}

# ====== 9a. the LOCKED Cheng formulation (settlement section 1) ==========
def cheng_formulation():
    """The exact numerical formulation, enumerated and bound to its source.

    "Cheng 1978" is not a reproducibility contract: recognised implementations
    differ in literals (1.3862944 vs log(4), 0.0138889 vs 1/72, 0.777778 vs 7/9)
    and in the logit form (log(u1/(1-u1)) vs log(u1) - log1p(-u1)). Those choices
    change results. This record names one of them and hashes the source that
    realises it.
    """
    import inspect
    src_bb = inspect.getsource(BR.cheng_bb)
    src_bc = inspect.getsource(BR.cheng_bc)
    src_dispatch = inspect.getsource(BR.cheng_dispatch)
    joined = src_dispatch + src_bb + src_bc
    return {
        "status":
            "The locked formulation is EXACTLY the one already measured in "
            "beta_ref.py. Nothing was changed to settle this item, and every "
            "Cheng, D6-04 and acceptance-margin raw file is byte-identical to "
            "the c121cc6 package. No evidence was regenerated for a formulation "
            "change, because there was none.",
        "belongs_to": "SIM_METHOD_VERSION",
        "dispatch": {
            "rule": "BB when min(alpha, beta) > 1.0; BC otherwise",
            "boundary": "equality belongs to BC",
            "operator": "strict >, on min(alpha, beta)"},
        "BB": {
            "applies_when": "min(alpha, beta) > 1",
            "parameter_orientation": "a = min(alpha0, beta0); b = max(alpha0, beta0)",
            "precomputed_once_per_driver": [
                "alpha = a + b",
                "beta  = sqrt((alpha - 2) / (2*a*b - alpha))",
                "gamma = a + 1/beta"],
            "per_proposal_attempt": [
                "u1 = next_u(); u2 = next_u()      (exactly two uniforms)",
                "vlog = log(u1 / (1 - u1))         (LOCKED logit form)",
                "v = beta * vlog",
                "w = a * exp(v)",
                "z = u1 * u1 * u2",
                "rr = gamma * v - 1.3862944",
                "s = a + rr - w",
                "accept if  s + 2.609438 >= 5.0 * z",
                "else t = log(z); accept if s >= t",
                "else accept if rr + alpha * log(alpha / (b + w)) >= t",
                "else reject and retry"],
            "literals": [
                {"value": "1.3862944", "role": "log(4) in the squeeze",
                 "kind": "LITERAL, 8 significant digits - NOT computed as log(4)"},
                {"value": "2.609438", "role": "1 + log(5) in the squeeze",
                 "kind": "LITERAL, 7 significant digits - NOT computed"},
                {"value": "5.0", "role": "squeeze coefficient on z", "kind": "LITERAL, exact"},
                {"value": "2.0", "role": "in the beta setup", "kind": "LITERAL, exact"},
                {"value": "1.0", "role": "in 1 - u1 and 1/beta", "kind": "LITERAL, exact"}],
            "return": "w / (b + w) when the caller's FIRST parameter was the min; "
                      "b / (b + w) otherwise"},
        "BC": {
            "applies_when": "min(alpha, beta) <= 1",
            "parameter_orientation": "a = max(alpha0, beta0); b = min(alpha0, beta0) "
                                     "- the OPPOSITE of BB, and inverting it is a "
                                     "silent defect that returns the mirrored "
                                     "distribution",
            "precomputed_once_per_driver": [
                "alpha = a + b",
                "beta  = 1 / b",
                "delta = 1 + a - b",
                "k1 = delta * (0.0138889 + 0.0416667 * b) / (a * beta - 0.777778)",
                "k2 = 0.25 + (0.5 + 0.25 / delta) * b"],
            "per_proposal_attempt": [
                "u1 = next_u(); u2 = next_u()      (exactly two uniforms)",
                "if u1 < 0.5:  y = u1*u2; z = u1*y; reject if 0.25*u2 + z - y >= k1",
                "else:         z = u1*u1*u2",
                "              if z <= 0.25: vlog = log(u1/(1-u1)); v = beta*vlog; "
                "w = a*exp(v); ACCEPT",
                "              if z >= k2: reject",
                "vlog = log(u1 / (1 - u1)); v = beta * vlog; w = a * exp(v)",
                "accept if alpha * (log(alpha / (b + w)) + v) - 1.3862944 >= log(z)",
                "else reject and retry"],
            "literals": [
                {"value": "0.0138889", "role": "1/72 in k1",
                 "kind": "LITERAL, 6 significant digits - NOT computed as 1/72"},
                {"value": "0.0416667", "role": "3/72 in k1",
                 "kind": "LITERAL, 6 significant digits - NOT computed"},
                {"value": "0.777778", "role": "7/9 in k1",
                 "kind": "LITERAL, 6 significant digits - NOT computed"},
                {"value": "1.3862944", "role": "log(4) in the acceptance test",
                 "kind": "LITERAL, same 8-digit value as BB"},
                {"value": "0.25", "role": "squeeze bound and k2 terms", "kind": "LITERAL, exact"},
                {"value": "0.5", "role": "the u1 branch split and k2", "kind": "LITERAL, exact"}],
            "return": "w / (b + w) when the caller's FIRST parameter was the max; "
                      "b / (b + w) otherwise"},
        "logit_form": {
            "locked": "log(u1 / (1.0 - u1))",
            "rejected_alternative": "log(u1) - log1p(-u1)",
            "why_it_matters": "the two disagree in the tails and the difference "
                              "reaches the RETURNED SAMPLE, not merely the "
                              "acceptance decision (control 18)"},
        "comparison_operators": {
            "squeeze and acceptance tests": ">= , evaluated left >= right exactly "
                                            "as written above",
            "BC branch split": "u1 < 0.5",
            "BC fast accept": "z <= 0.25",
            "BC k2 reject": "z >= k2",
            "dispatch": "min(alpha, beta) > 1.0"},
        "what_the_literals_actually_control":
            "MEASURED, not assumed: the squeeze literals never change the value "
            "of an accepted proposal - they change WHICH proposals are accepted, "
            "and therefore consumption and the whole subsequent sequence. The "
            "logit form changes the value itself. Both must be locked, for "
            "different reasons. Controls 18 and 18b demonstrate each.",
        "degenerate_a_m_b":
            "Handled BEFORE dispatch and before parameterisation, per plan "
            "section 4.0: a degenerate driver returns `a`, enters no sampler, "
            "forms no r = (m-a)/(b-a), and consumes ZERO uniforms. Cheng is "
            "never reached, so alpha and beta are never formed and 0/0 cannot "
            "arise. Evidence: raw/degenerate_d6_18.json.",
        "source_binding": {
            "functions": ["cheng_dispatch", "cheng_bb", "cheng_bc"],
            "file": "pccm/evidence/phase6_step0/scripts/beta_ref.py",
            "sha256": hashlib.sha256(joined.encode("utf-8")).hexdigest(),
            "note": "The hash covers the three function bodies as retained. It "
                    "binds this record to the code that produced every Cheng "
                    "number in the package, so the two cannot drift apart "
                    "silently."}}

# ============ 9b. Cheng authority vectors (settlement section 2) =========
CHENG_VECTOR_CASES = [
    {"label": "BB interior 2/4", "r": 0.25, "stream": 0},
    {"label": "BB symmetric 3/3", "r": 0.5, "stream": 1},
    {"label": "BB near-boundary 1.04/4.96", "r": 0.01, "stream": 2},
    {"label": "BC alpha=1 beta=5", "r": 0.0, "stream": 3},
    {"label": "BC alpha=5 beta=1", "r": 1.0, "stream": 4},
]


def cheng_vectors(jump_res, n=24):
    """Deterministic per-sample vectors. These DEFINE the locked formulation.

    Gate-A reference vectors. They do NOT imply Python/VBA sample-for-sample
    identity in a full seeded Beta run - revision 6's Layer G is unchanged.
    """
    J1 = jump_res["A1p127_derived_from_recurrence"]
    J2 = jump_res["A2p127_derived_from_recurrence"]
    states = stream_states(SEED_STATE, max(c["stream"] for c in CHENG_VECTOR_CASES) + 1,
                           J1, J2)
    cases = []
    for case in CHENG_VECTOR_CASES:
        a, b = BR.pert_shape(case["r"])
        init = states[case["stream"]]
        g = ExactMrg(list(init))
        counter = {"attempts": 0, "uniforms": 0}
        samples = []
        for i in range(n):
            at0, u0 = counter["attempts"], counter["uniforms"]
            x = BR.cheng_sample(a, b, g.next_u, counter)
            samples.append({
                "index": i + 1,
                "accepted_sample": repr(x),
                "proposal_attempts_for_this_sample": counter["attempts"] - at0,
                "uniforms_for_this_sample": counter["uniforms"] - u0,
                "cumulative_uniforms": counter["uniforms"],
                "rng_state_after_sample": list(g.s)})
        attempts = [sm["proposal_attempts_for_this_sample"] for sm in samples]
        cases.append({
            "label": case["label"], "r": case["r"], "alpha": a, "beta": b,
            "dispatch": BR.cheng_dispatch(a, b),
            "stream_index": case["stream"], "initial_state": init,
            "samples_recorded": n,
            "exercises_immediate_acceptance": 1 in attempts,
            "exercises_at_least_one_retry": any(v > 1 for v in attempts),
            "max_attempts_in_case": max(attempts),
            "final_state": list(g.s),
            "total_attempts": counter["attempts"],
            "total_uniforms": counter["uniforms"],
            "samples": samples})
    # N is chosen so EVERY case exercises both an immediate acceptance and at
    # least one retry. The earliest retry in the highest-acceptance case is
    # sample 18, so N = 24 covers the family with margin. This refuses rather
    # than silently retaining vectors that miss a required path.
    missing = [c["label"] for c in cases
               if not (c["exercises_at_least_one_retry"]
                       and c["exercises_immediate_acceptance"])]
    if missing:
        raise SystemExit("cheng vectors do not exercise required paths: "
                         + ", ".join(missing))
    return {
        "purpose": "Gate-A reference vectors that DEFINE the locked Cheng "
                   "floating-point formulation. Any implementation reproducing "
                   "these reproduces the formulation; any that does not, does not.",
        "not_a_cross_language_identity_claim":
            "These vectors do not imply Python/VBA sample-for-sample identity in "
            "a full seeded Beta run. Revision 6's Layer G is unchanged.",
        "base_state": list(SEED_STATE),
        "samples_per_case": n,
        "why_this_n":
            "N is chosen so every case exercises an immediate acceptance AND at "
            "least one retry. The earliest retry in the highest-acceptance case "
            "is sample 18, so N = 24 covers the family with margin. The "
            "generator refuses to emit vectors that miss either path.",
        "both_dispatches_covered":
            {"BB": sum(1 for c in cases if c["dispatch"] == "BB"),
             "BC": sum(1 for c in cases if c["dispatch"] == "BC")},
        "every_case_exercises_a_retry":
            all(c["exercises_at_least_one_retry"] for c in cases),
        "every_case_exercises_immediate_acceptance":
            all(c["exercises_immediate_acceptance"] for c in cases),
        "cases": cases}

# ====== 9c. degenerate severity under D6-18b (settlement section 3) ======
def degenerate_d6_18(jump_res, iters=1000):
    """One Risk, run `iters` iterations, under D6-18 option B.

    The severity SAMPLER IS INVOKED every iteration. What it consumes depends on
    the DISTRIBUTION, not on the occurrence outcome:
      non-degenerate -> the sampler consumes its contract-defined uniforms;
      degenerate     -> it returns the constant and consumes ZERO.
    """
    J1 = jump_res["A1p127_derived_from_recurrence"]
    J2 = jump_res["A2p127_derived_from_recurrence"]
    st = stream_states(SEED_STATE, 3, J1, J2)
    rows = []
    for degenerate in (True, False):
        triple = (7.0, 7.0, 7.0) if degenerate else (0.0, 50.0, 100.0)
        for p in GRID["probabilities"] + [0.0]:
            occ_state, sev_state = list(st[1]), list(st[2])
            occ_g, sev_g = ExactMrg(occ_state), ExactMrg(sev_state)
            counter = {"attempts": 0, "uniforms": 0}
            invocations = occurrences = 0
            contribution_total = 0.0
            severity_values = set()
            for _ in range(iters):
                occurred = occ_g.next_u() < p          # Bernoulli, one uniform
                sev, _used = TX.sample_driver("pert", *triple, sev_g.next_u, counter)
                invocations += 1
                severity_values.add(repr(sev))
                if occurred:
                    occurrences += 1
                    contribution_total += sev
            rows.append({
                "degenerate": degenerate, "triple": list(triple),
                "probability": p,
                "iterations": iters,
                "occurrence_uniforms_consumed": iters,
                "severity_sampler_invocations": invocations,
                "severity_sampler_invoked_every_iteration": invocations == iters,
                "severity_uniforms_consumed": counter["uniforms"],
                "severity_stream_unchanged":
                    sev_g.s == sev_state,
                "distinct_severity_values": len(severity_values),
                "realised_occurrences": occurrences,
                "expected_occurrences": iters * p,
                "contribution_total": repr(contribution_total)})
    return {
        "d6_18b_contract":
            "The severity SAMPLER IS INVOKED every Risk iteration. A "
            "non-degenerate distribution consumes its contract-defined uniforms; "
            "a degenerate distribution returns the constant and consumes zero. "
            "The sampled value is used only when occurrence is True. "
            "'The stream advances once per iteration' is NOT the contract - it is "
            "false for a degenerate severity.",
        "two_uniforms_per_attempt_scope":
            "'each Cheng proposal attempt consumes exactly two uniforms' applies "
            "to NON-DEGENERATE Cheng BB/BC proposal attempts only. A degenerate "
            "driver makes no proposal attempt at all.",
        "rows": rows}

# ============ 9d. version register (settlement section 6) ================
def version_register():
    return {
        "principle":
            "Every version is owned by exactly one authority and every listed "
            "change is classified against exactly one version. A change that "
            "would alter retained numbers without bumping some version is the "
            "defect this register exists to prevent.",
        "versions": [
            {"name": "FP_VERSION", "initial": 1, "owner": "spec/calc_contract.yaml",
             "status": "INHERITED - unchanged by Phase 6",
             "covers": "the Phase-5 calculation INPUT fingerprint algorithm"},
            {"name": "RNG_VERSION", "initial": 1, "owner": "spec/sim_contract.yaml",
             "status": "NEW - settled here",
             "covers": "everything that determines the uniform stream a component "
                       "sees: MRG constants and combination, scalar-seed mapping, "
                       "AUTO-seed mapping, stream assignment, jump semantics"},
            {"name": "SIM_METHOD_VERSION", "initial": 1,
             "owner": "spec/sim_contract.yaml", "status": "NEW - settled here",
             "covers": "everything that turns uniforms into published numbers: "
                       "the Cheng formulation and its expression order, D6-18 "
                       "advancement, the degenerate rule, accumulation order, "
                       "percentile/statistical method"},
            {"name": "result_digest version field", "initial": 1,
             "owner": "spec/sim_contract.yaml",
             "status": "NEW - settled here as option (b)",
             "covers": "the PCCM-RD stream framing only",
             "decision":
                 "The version integer inside the PCCM-RD stream IS "
                 "SIM_METHOD_VERSION. No separate RESULT_DIGEST_VERSION is "
                 "created. Rationale: the digest exists to compare runs, and two "
                 "runs are comparable exactly when the method that produced the "
                 "numbers is the same. A framing change without a method change "
                 "still bumps SIM_METHOD_VERSION, because it changes what the "
                 "digest means; a third version would add a field nobody could "
                 "state a distinct rule for."},
        ],
        "change_classification": [
            {"change": "MRG constants or uniform combination", "bumps": "RNG_VERSION"},
            {"change": "scalar-seed -> six-word state mapping", "bumps": "RNG_VERSION"},
            {"change": "AUTO-seed nonce -> effective_seed mapping", "bumps": "RNG_VERSION"},
            {"change": "stream-assignment rule (D6-16)", "bumps": "RNG_VERSION"},
            {"change": "jump algorithm semantics (matrices, exponent, arithmetic path)",
             "bumps": "RNG_VERSION"},
            {"change": "Cheng formulation, literals or expression order",
             "bumps": "SIM_METHOD_VERSION"},
            {"change": "D6-18 advancement rule", "bumps": "SIM_METHOD_VERSION"},
            {"change": "degenerate-driver rule", "bumps": "SIM_METHOD_VERSION"},
            {"change": "accumulation order", "bumps": "SIM_METHOD_VERSION"},
            {"change": "percentile or statistical method", "bumps": "SIM_METHOD_VERSION"},
            {"change": "result-digest framing or hash stream",
             "bumps": "SIM_METHOD_VERSION (the PCCM-RD version field carries it)"},
            {"change": "Phase-5 input fingerprint encoding", "bumps": "FP_VERSION"},
        ],
        "note_on_jump_exponent":
            "The 2^127 exponent is inside RNG_VERSION, not a separate constant to "
            "vary freely: changing it changes every stream after 0.",
        "note_on_ordering":
            "A change may bump BOTH. Nothing here permits a change that bumps "
            "NEITHER while altering a retained number."}

# ============ 9e. tolerance model (settlement section 7) =================
TOL_TRIPLES = [
    (0.0, 50.0, 100.0),
    (-100.0, 0.0, 100.0),
    (-1.0e308, 0.0, 1.0e308),
    (1.0e-300, 5.0e-300, 1.0e-299),
    (1.0e300, 1.5e300, 2.0e300),
    (-5.0, -5.0, 3.0),
    (2.0, 2.0, 9.0),
    (2.0, 9.0, 9.0),
]


def _ulp_gap(x, y):
    if x == y:
        return 0.0
    u = max(math.ulp(x), math.ulp(y))
    return abs(x - y) / u if u > 0 else float("inf")


def tolerance_model():
    """Measured sensitivity, then a policy derived from it.

    Two independent sources of cross-implementation difference are measured:
      (1) EXPRESSION ORDER - the accepted form against an algebraically
          equivalent one, over the accepted extreme Double domain;
      (2) LIBM - a one-ULP difference in `log`, propagated through `exp`, which
          converts an absolute argument error into a relative result error.
    """
    rows = []
    for name, acc, alt, needs_m in (
            ("uniform_transform", TX.uniform_accepted, TX.uniform_alternate, False),
            ("triangular_transform", TX.triangular_accepted, TX.triangular_alternate, True),
            ("pert_rescale", TX.pert_rescale_accepted, TX.pert_rescale_alternate, False)):
        worst_ulp, worst_rel, worst_at, nonfinite_alt = 0.0, 0.0, None, 0
        for (a, m, b) in TOL_TRIPLES:
            if a == b:
                continue
            for u in GRID["u_probes"]:
                x = acc(u, a, m, b) if needs_m else acc(u, a, b)
                y = alt(u, a, m, b) if needs_m else alt(u, a, b)
                if not math.isfinite(y):
                    nonfinite_alt += 1
                    continue
                g = _ulp_gap(x, y)
                rel = abs(x - y) / (abs(x) if x != 0.0 else 1.0)
                if g > worst_ulp:
                    worst_ulp, worst_at = g, {"triple": [a, m, b], "u": u}
                worst_rel = max(worst_rel, rel)
        rows.append({"subject": name,
                     "expression_order_max_ulp_gap": worst_ulp,
                     "expression_order_max_rel_diff": worst_rel,
                     "worst_at": worst_at,
                     "alternate_form_overflowed_cases": nonfinite_alt})

    # Cheng: one-ULP libm perturbation, acceptance path held fixed.
    sens_rows = []
    for r in GRID["r_values"]:
        a, b = BR.pert_shape(r)
        g = ExactMrg([SEEDS["cheng_stream_seed"]] * 6)
        counter = {"attempts": 0, "uniforms": 0}
        BR.sens_reset(True)
        for _ in range(5000):
            BR.cheng_sample(a, b, g.next_u, counter)
        snap = BR.sens_snapshot()
        BR.sens_reset(False)
        sens_rows.append({"r": r, "alpha": a, "beta": b, **snap})
    worst_cheng = max(x["max_rel_out"] for x in sens_rows)
    worst_v = max(x["max_abs_v"] for x in sens_rows)
    worst_transform_ulp = max(x["expression_order_max_ulp_gap"] for x in rows)

    return {
        "method":
            "Expression order: the accepted form against an algebraically "
            "equivalent one over the accepted extreme Double domain, in ULPs. "
            "Libm: a one-ULP perturbation of the `log` result inside Cheng, "
            "propagated through `exp` and the final ratio, with the acceptance "
            "path held FIXED - this measures output sensitivity, not branch "
            "stability, which is section 5.3's separate subject.",
        "transform_rows": rows,
        "cheng_sensitivity_rows": sens_rows,
        "cheng_worst_relative_output_change_from_one_ulp_log": worst_cheng,
        "cheng_worst_abs_v_amplifier": worst_v,
        "transform_worst_ulp_gap": worst_transform_ulp,
        "ownership_decision": "B",
        "ownership_rationale":
            "The tolerance is NOT a simulation-runtime contract and does not "
            "belong in sim_contract.yaml. The engine never compares two Doubles "
            "for approximate equality at runtime: replay comparison is by "
            "result_digest, which is EXACT, and no published number is produced "
            "by a tolerance test. A tolerance exists only when two "
            "IMPLEMENTATIONS are compared, which is oracle/Gate-A/Gate-B "
            "evidence. Its single owner is therefore the Phase-6 oracle and "
            "evidence policy (plan section 15.1), and sim_contract.yaml stores "
            "no tolerance at all.",
        "policy": [
            {"subject": "individual Uniform / Triangular / PERT-rescale "
                        "transformed samples",
             "rule": "relative, with an absolute floor keyed to the driver's "
                     "conditioning scale s = max(|a|,|m|,|b|)",
             "value": "rel <= 1e-12, or abs <= 1e-12 * s",
             "basis": f"measured worst expression-order gap "
                      f"{worst_transform_ulp:.0f} ULP ~ "
                      f"{worst_transform_ulp * 2 ** -52:.2e} relative; the policy "
                      f"leaves about "
                      f"{1e-12 / max(worst_transform_ulp * 2 ** -52, 1e-300):.0f}x "
                      f"headroom"},
            {"subject": "deterministic Cheng vector outputs",
             "rule": "relative", "value": "rel <= 1e-11",
             "basis": f"measured worst relative output change from a one-ULP "
                      f"`log` difference is {worst_cheng:.3e}; the amplifier is "
                      f"|v|, worst {worst_v:.3f}. The policy leaves about "
                      f"{1e-11 / worst_cheng:.0f}x headroom over a one-ULP libm "
                      f"difference, and is still far tighter than any real "
                      f"algorithmic error"},
            {"subject": "F1 per-iteration no-Beta end-to-end totals",
             "rule": "relative, with an absolute floor keyed to the iteration's "
                     "accumulation scale S = max |contribution| over the drivers "
                     "summed",
             "value": "rel <= 3e-10, or abs <= 3e-10 * S",
             "basis": "composition, not a new measurement: an iteration total "
                      "sums at most 300 driver contributions each carrying at "
                      "most the per-sample relative error above, so the absolute "
                      "error is bounded by 300 * 1e-12 * S. A purely relative "
                      "test is unusable here because cancellation can drive a "
                      "total near zero while every contribution is large - hence "
                      "the scale floor"},
            {"subject": "summary statistics compared cross-language",
             "rule": "relative, with the same accumulation-scale floor",
             "value": "rel <= 3e-10, or abs <= 3e-10 * S",
             "basis": "a percentile is one order statistic - an element of the "
                      "sorted sample, or a convex blend of two adjacent ones - so "
                      "it inherits the per-iteration bound and no more. The mean "
                      "and standard deviation inherit the same bound under the "
                      "accepted scale-aware accumulation"},
        ],
        "exact_subjects_remain_exact": [
            "MRG32k3a state and uniform values",
            "jump state",
            "Bernoulli occurrence decisions",
            "proposal and draw counts where the arithmetic path is fixed",
            "same-runtime G2/G3 result_digest",
        ]}

# ==================== 10. vectors =======================================
def write_vectors(res):
    J1 = res["jump"]["A1p127_derived_from_recurrence"]
    J2 = res["jump"]["A2p127_derived_from_recurrence"]
    vec_streams = {}
    for k, st in res["jump"]["stream_initial_states"].items():
        g = ExactMrg(list(st))
        vec_streams[k] = {"initial_state": st,
                          "first_5_uniforms": [repr(g.next_u()) for _ in range(5)],
                          "state_after_5": g.s}
    w("vectors", "jump_vectors.json", {
        "jump_exponent": "2^127",
        "A1_jump": J1, "A2_jump": J2,
        "state_order_stored": "[s10,s11,s12,s20,s21,s22] oldest-first",
        "matrix_operand_order": "newest-first; reverse each triple before and after",
        "streams": vec_streams})
    w("vectors", "rng_vectors.json", {
        "constants": {"m1": M1, "m2": M2, "a12": 1403580, "a13n": 810728,
                      "a21": 527612, "a23n": 1370589,
                      "norm": repr(2.328306549295727688e-10)},
        "seed_state_12345": res["rng"]["seed_state"],
        "first_20_uniforms": res["rng"]["first_20_uniforms"],
        "state_after_20": res["rng"]["state_after_20"],
        "per_seed": res["rng"]["per_seed"]})
    w("vectors", "seed_vectors.json", {
        "nonce_to_seed_pairs": res["seed"]["nonce_to_seed_pairs"],
        "seed_to_state_rule": "state = [seed] * 6",
        "examples": [{"seed": s, "state": SM.seed_to_state(s)}
                     for s in [1, 2, 12345, 2147483646]]})
    w("vectors", "digest_vectors.json",
      {"hash_constants": res["digest"]["hash_constants"],
       "cases": res["digest"]["cases"]})
    w("vectors", "cheng_vectors.json", res["cheng_vectors"])
    w("vectors", "stream_assignment_vectors.json",
      {"family_a_first_10": res["streams"]["family_a"]["first_10_streams"],
       "family_a_last_4": res["streams"]["family_a"]["last_4_streams"],
       "total_components": res["streams"]["design_target"]["total_components"]})

# ==================== 11. manifest ======================================
def write_manifest():
    import subprocess
    def sha(path):
        h = hashlib.sha256()
        with open(path, "rb") as f:
            h.update(f.read())
        return h.hexdigest()
    files = {}
    for sub in ("scripts", "inputs", "raw", "vectors", "summaries"):
        d = os.path.join(ROOT, sub)
        for name in sorted(os.listdir(d)):
            p = os.path.join(d, name)
            if os.path.isfile(p) and not name.endswith(".pyc"):
                files[f"{sub}/{name}"] = {"sha256": sha(p),
                                          "bytes": os.path.getsize(p)}
    try:
        head = subprocess.check_output(["git", "rev-parse", "HEAD"],
                                       cwd=ROOT, text=True).strip()
    except Exception:
        head = None
    return {
        "accepted_planning_baseline": "03aa5044cb535513976f0ec3840bc332747678c8",
        "phase5_executable_baseline": "f571154118083e569e1fb9fbf9bf72852cc2d568",
        "git_head_at_generation": head,
        "step0_commit": None,
        "step0_commit_note":
            "A commit cannot contain its own hash, so this field stays null. "
            "git_head_at_generation above is the PARENT of the Step-0 commit, "
            "and it equals the accepted planning baseline. The Step-0 commit "
            "hash itself is reported in the delivery message; `git log -1` on "
            "the branch gives it directly.",
        "generator": "pccm/evidence/phase6_step0/scripts/run_all.py",
        "python_version": sys.version,
        "python_implementation": platform.python_implementation(),
        "platform": platform.platform(),
        "seeds": SEEDS,
        "grid": GRID,
        "provenance": "All numbers in raw/, vectors/ and summaries/ are produced "
                      "by the scripts listed here from the inputs listed here. "
                      "No value was copied from prior prose. No Windows or Excel "
                      "runtime was involved; nothing here is runtime evidence.",
        "files": files}

if __name__ == "__main__":
    res = {}
    res["rng"] = rng_backbone(); w("raw", "rng_backbone.json", res["rng"])
    res["jump"] = jump_evidence(); w("raw", "jump.json", res["jump"])
    res["cheng"] = cheng_measurements(); w("raw", "cheng.json", res["cheng"])
    res["cand_a"] = candidate_a(); w("raw", "candidate_a_inverse_cdf.json", res["cand_a"])
    res["cand_a2"] = candidate_a2(); w("raw", "candidate_a2_newton.json", res["cand_a2"])
    res["cand_c"] = candidate_c(); w("raw", "candidate_c_table.json", res["cand_c"])
    res["degenerate"] = degenerate_d6_18(res["jump"])
    w("raw", "degenerate_d6_18.json", res["degenerate"])
    res["margins"] = acceptance_margin_model(res["cheng"])
    w("raw", "acceptance_margin_model.json", res["margins"])
    res["d6_04"] = d6_04_comparison(res["cheng"], res["cand_a"], res["cand_a2"],
                                    res["cand_c"])
    w("raw", "d6_04_comparison.json", res["d6_04"])
    res["d6_18"] = d6_18_model(res["cheng"], res["degenerate"])
    w("raw", "d6_18_operation_model.json", res["d6_18"])
    res["seed"] = seed_evidence(); w("raw", "seed_map.json", res["seed"])
    res["digest"] = digest_evidence(); w("raw", "digest.json", res["digest"])
    res["streams"] = stream_architecture(); w("raw", "stream_architecture.json", res["streams"])
    res["cheng_formulation"] = cheng_formulation()
    w("raw", "cheng_formulation.json", res["cheng_formulation"])
    res["cheng_vectors"] = cheng_vectors(res["jump"])
    res["versions"] = version_register(); w("raw", "version_register.json", res["versions"])
    res["tolerance"] = tolerance_model(); w("raw", "tolerance_model.json", res["tolerance"])
    write_vectors(res)

    # Controls run inside the same invocation so the manifest can never hash a
    # stale controls result. controls.py stays independently runnable.
    import controls as CTL
    w("summaries", "controls.json", CTL.RESULTS)
    lines = [("DETECTED " if c["detected"] else "MISSED   ") + c["control"]
             for c in CTL.RESULTS]
    fired = sum(c["detected"] for c in CTL.RESULTS)
    lines.append("")
    lines.append(f"{fired}/{len(CTL.RESULTS)} controls fired")
    with open(os.path.join(ROOT, "summaries", "controls.txt"), "w",
              encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    with open(os.path.join(ROOT, "raw", "run_log.txt"), "w",
              encoding="utf-8") as f:
        f.write("Step-0 evidence generation\n")
        f.write(f"python      {sys.version.splitlines()[0]}\n")
        f.write(f"platform    {platform.platform()}\n")
        f.write(f"cheng n     {len(res['cheng'])} shapes x 20000 samples\n")
        f.write(f"candidate C {len(res['cand_c'])} node counts x "
                f"{res['cand_c'][0]['shapes_measured']} shapes\n")
        f.write(f"controls    {fired}/{len(CTL.RESULTS)} fired\n")
        f.write("No Windows or Excel runtime was involved.\n")

    w("summaries", "environment.json",
      {"python": sys.version, "platform": platform.platform(),
       "implementation": platform.python_implementation()})
    w("", "manifest.json", write_manifest())
    print("raw, vectors, summaries and manifest written")
