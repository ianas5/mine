"""Step-0 driver. Produces raw/, summaries/ and vectors/ deterministically."""
import json, math, os, statistics, sys, hashlib, platform
sys.path.insert(0, os.path.dirname(__file__))
from mrg32k3a import (M1, M2, A1, A2, ExactMrg, DoubleMrg, mat_pow2_mod,
                      jump_state, verify_convention, stream_states,
                      mat_vec_mod_safe, mat_vec_mod_naive)
import beta_ref as BR
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
def d6_18_model(cheng_rows):
    """Both options, at four probabilities, using the MEASURED Cheng behaviour.

    Consumption is quoted twice: with the family's worst measured shape and with
    the symmetric r = 0.5 shape, so the figures are not silently tied to one
    arbitrary choice of driver shape.
    """
    by_r = {row["r"]: row for row in cheng_rows}
    mid_u = by_r[0.5]["uniforms_per_accepted_mean"]
    worst_u = max(r["uniforms_per_accepted_mean"] for r in cheng_rows)
    worst_tr = max(r["transcendental_calls_per_sample"] for r in cheng_rows)
    risks, iters = 100, 100000
    stream_states_bytes = 400 * 6 * 8
    out = []
    for p in GRID["probabilities"]:
        occ = int(risks * iters * p)
        both = risks * iters
        out.append({
            "probability": p,
            "occurrence_uniforms": both,
            "option_A_severity_invocations": occ,
            "option_B_severity_invocations": both,
            "option_A_severity_uniforms_est_mid_shape": round(occ * mid_u),
            "option_B_severity_uniforms_est_mid_shape": round(both * mid_u),
            "option_A_severity_uniforms_est_worst_shape": round(occ * worst_u),
            "option_B_severity_uniforms_est_worst_shape": round(both * worst_u),
            "option_A_severity_transcendentals_est": round(occ * worst_tr),
            "option_B_severity_transcendentals_est": round(both * worst_tr),
            "option_A_stream_state_bytes": stream_states_bytes,
            "option_B_stream_state_bytes": stream_states_bytes,
            "memory_difference_bytes": 0,
            "uniforms_per_accepted_mid_shape": mid_u,
            "uniforms_per_accepted_worst_shape": worst_u,
            "transcendentals_per_sample_worst_shape": worst_tr})
    return out


# ==================== 7. D6-03 / D6-05 seed mapping =====================
def seed_evidence():
    order = SM.P - 1
    facts = SM.factorise(order)
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
    return {"modulus": SM.P, "multiplier": SM.MULT,
            "order_factorisation": facts,
            "multiplier_is_primitive_root": prim,
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
    res["margins"] = acceptance_margin_model(res["cheng"])
    w("raw", "acceptance_margin_model.json", res["margins"])
    res["d6_04"] = d6_04_comparison(res["cheng"], res["cand_a"], res["cand_a2"],
                                    res["cand_c"])
    w("raw", "d6_04_comparison.json", res["d6_04"])
    res["d6_18"] = d6_18_model(res["cheng"]); w("raw", "d6_18_operation_model.json", res["d6_18"])
    res["seed"] = seed_evidence(); w("raw", "seed_map.json", res["seed"])
    res["digest"] = digest_evidence(); w("raw", "digest.json", res["digest"])
    res["streams"] = stream_architecture(); w("raw", "stream_architecture.json", res["streams"])
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
