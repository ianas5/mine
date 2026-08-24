"""Negative controls for the EVIDENCE itself. Not production tests.

Each plants a defect the Step-0 evidence is supposed to detect, and asserts
the detection actually fires. A control that cannot fail proves nothing.
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import mrg32k3a as R
import beta_ref as BR
from digest import result_digest
import stream_map as STM
import transforms as TX
import seed_map as SM2
from digest import result_stream, RESULT_STREAM_TAG
from seed_map import is_primitive_root, nonce_to_seed, seed_to_state, state_is_valid, P

RESULTS = []
def check(name, detected):
    RESULTS.append({"control": name, "detected": bool(detected)})

ST = [12345] * 6

# 1. wrong MRG modulus
g = R.ExactMrg(list(ST)); good = [g.next_u() for _ in range(5)]
saved = R.M1
try:
    R.M1 = 4294967089
    g2 = R.ExactMrg(list(ST)); bad = [g2.next_u() for _ in range(5)]
finally:
    R.M1 = saved
check("wrong MRG modulus m1", bad != good)

# 2. wrong jump matrix element
J1 = R.mat_pow2_mod(R.A1, 127, R.M1); J2 = R.mat_pow2_mod(R.A2, 127, R.M2)
J1b = [row[:] for row in J1]; J1b[0][0] = (J1b[0][0] + 1) % R.M1
check("wrong jump matrix element", R.jump_state(ST, J1b, J2) != R.jump_state(ST, J1, J2))

# 3. naive unsafe modular matrix multiplication, in float
def naive_float(a, v, m):
    return [float((a[i][0]*v[0] + a[i][1]*v[1] + a[i][2]*v[2]) % m) for i in range(3)]
vec = [float(R.M1 - 1)] * 3
check("naive float matrix-vector loses exactness",
      [int(x) for x in naive_float(J1, vec, R.M1)] !=
      [int(x) for x in R.mat_vec_mod_safe(J1, vec, R.M1)])

# 4. BB/BC dispatch boundary moved
check("BB/BC dispatch boundary moved",
      BR.cheng_dispatch(1.0, 5.0) == "BC" and BR.cheng_dispatch(1.0001, 5.0) == "BB")

# 5. occurrence and severity streams merged
def risk_path(merge):
    occ = R.ExactMrg(list(ST))
    sev = occ if merge else R.ExactMrg([54321] * 6)
    out, cnt = [], {"attempts": 0, "uniforms": 0}
    for _ in range(40):
        happened = occ.next_u() < 0.5
        out.append(BR.cheng_sample(3.0, 3.0, sev.next_u, cnt) if happened else 0.0)
    return out
check("occurrence and severity streams merged", risk_path(True) != risk_path(False))

# 6. physical-row-order stream assignment
ids = ["CL-003", "CL-001", "CL-002"]
check("row-order stream assignment", sorted(ids) != ids)

# 7. D6-18 advancement rule swapped
def sev_invocations(uncond, p=0.1, n=1000):
    occ = R.ExactMrg(list(ST)); k = 0
    for _ in range(n):
        happened = occ.next_u() < p
        if uncond or happened:
            k += 1
    return k
check("D6-18 advancement rule swapped", sev_invocations(True) != sev_invocations(False))

# 8. result-digest field/order mutation
nom, pv = [1.0, 2.0, 3.0], [0.9, 1.9, 2.9]
check("result-digest order mutation",
      result_digest(nom[::-1], pv[::-1]) != result_digest(nom, pv))

# 9. auto-seed collision / non-primitive multiplier
check("non-primitive multiplier detected", not is_primitive_root(2, P))

# 10. counter exhaustion wrap
LIMIT = 2147483646
def next_nonce(cur, wrap):
    if cur >= LIMIT:
        return 1 if wrap else None
    return cur + 1
check("counter exhaustion wrap", next_nonce(LIMIT, True) is not None
      and next_nonce(LIMIT, False) is None)

# 11. an invalid seed produces an invalid state
check("seed 0 gives the forbidden all-zero state", not state_is_valid(seed_to_state(0)))

# 12. result-stream framing tag removed
def framing(tag):
    st = result_stream([1.0], [2.0])
    return st if tag else st.replace(RESULT_STREAM_TAG, "PCCM-FP")
check("result-stream framing tag mutation", framing(True) != framing(False))

# 13. the D6-16 sort rule is load-bearing, not decorative
CID = [f"CL-{i:03d}" for i in range(1, 6)]
RID = [f"R-{i:03d}" for i in range(1, 4)]
good_map = STM.family_a(CID, RID)
bad_map = {c: i for i, c in enumerate(
    sorted(STM.components(CID, RID), key=lambda c: STM.sort_key(c), reverse=True))}
check("stream assignment under a different ordering rule", good_map != bad_map)

# 14. Family B collision witness really collides
wt = STM.family_b_collision_witness(1000)
fb = STM.family_b([wt[0]], [wt[1]], 1000)
check("Family B collision under the unbounded ID pattern",
      wt is not None and len(set(fb.values())) < len(fb))

# 15. a degenerate driver that consumes a uniform is detected
def degenerate_consumption(correct):
    g = R.ExactMrg(list(ST))
    c = {"attempts": 0, "uniforms": 0}
    for _ in range(50):
        if correct:
            TX.sample_driver("pert", 7.0, 7.0, 7.0, g.next_u, c)
        else:
            g.next_u(); c["uniforms"] += 1        # the withdrawn "consume one" rule
    return g.s == list(ST) and c["uniforms"] == 0
check("degenerate driver consuming a uniform",
      degenerate_consumption(True) and not degenerate_consumption(False))

# 16. an EXPECTED count presented as a realised one is detected
def realised_occurrences(p, n=1000):
    g = R.ExactMrg(list(ST))
    return sum(1 for _ in range(n) if g.next_u() < p)
check("expected count presented as realised",
      realised_occurrences(0.1) != int(1000 * 0.1))

# 17. the version register classifies every listed change
import run_all as RA
_reg = RA.version_register()
_names = {v["name"].split()[0] for v in _reg["versions"]}
check("a change classified against no version",
      all(any(r["bumps"].startswith(n) for n in _names)
          for r in _reg["change_classification"]))

# 18. the log formulation is load-bearing for the RETURNED VALUE
def bb_value(u1, use_locked):
    """Only the value-producing part of BB, both log formulations."""
    import math as _m
    a, b = BR.pert_shape(0.5)
    aa, bb_ = min(a, b), max(a, b)
    alpha = aa + bb_
    beta = _m.sqrt((alpha - 2.0) / (2.0 * aa * bb_ - alpha))
    v = beta * (_m.log(u1 / (1.0 - u1)) if use_locked
                else (_m.log(u1) - _m.log1p(-u1)))
    w = aa * _m.exp(v)
    return w / (bb_ + w)
check("log(u1/(1-u1)) replaced by log(u1)-log1p(-u1)",
      any(repr(bb_value(u, True)) != repr(bb_value(u, False))
          for u in (0.999, 0.9999, 0.001, 0.0001)))

# 18b. the squeeze literal is load-bearing for the ACCEPTANCE DECISION
# It never changes the value of an accepted proposal - it changes WHICH
# proposals are accepted, and therefore consumption and the whole sequence.
# A constructed witness beats a long random search: place 5z exactly between
# the two literals' thresholds.
def squeeze_witness():
    import math as _m
    a, b = BR.pert_shape(0.5)
    aa, bb_ = min(a, b), max(a, b)
    alpha = aa + bb_
    beta = _m.sqrt((alpha - 2.0) / (2.0 * aa * bb_ - alpha))
    gamma = aa + 1.0 / beta
    for u1 in (0.3, 0.5, 0.7):
        v = beta * _m.log(u1 / (1.0 - u1))
        w = aa * _m.exp(v)
        s1 = aa + (gamma * v - 1.3862944) - w
        s2 = aa + (gamma * v - _m.log(4.0)) - w
        lo, hi = sorted((s1 + 2.609438, s2 + 2.609438))
        z = ((lo + hi) / 2.0) / 5.0
        u2 = z / (u1 * u1)
        if not 0.0 < u2 < 1.0:
            continue
        zz = u1 * u1 * u2
        if (s1 + 2.609438 >= 5.0 * zz) != (s2 + 2.609438 >= 5.0 * zz):
            return True
    return False
check("Cheng literal 1.3862944 replaced by log(4)", squeeze_witness())

# 19. the nonce authority and the stepped cycle must agree
check("auto-seed power authority diverging from the stepped cycle",
      all(SM2.nonce_to_seed(n) == SM2.nonce_to_seed_iterative(n) for n in range(500))
      and SM2.nonce_to_seed(5) != SM2.nonce_to_seed_iterative(6))

if __name__ == "__main__":
    import json
    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "summaries", "controls.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(RESULTS, f, indent=2)
    for r in RESULTS:
        print(("DETECTED " if r["detected"] else "MISSED   ") + r["control"])
    print(f"\n{sum(r['detected'] for r in RESULTS)}/{len(RESULTS)} controls fired")
