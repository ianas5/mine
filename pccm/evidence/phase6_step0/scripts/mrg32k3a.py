"""MRG32k3a: two independent arithmetic paths, plus jump-ahead.

PATH 1 (`ExactMrg`)   - Python unbounded integers. The mathematical definition.
PATH 2 (`DoubleMrg`)  - the arithmetic a VBA implementation would perform:
                        every value held as a float that happens to be an
                        integer, reduced by Fix(p/m) exactly as the plan's
                        section 5.2.1 specifies.

They must agree exactly. Where they do, the plan's claim that VBA Double
arithmetic is exact for this recurrence is demonstrated rather than asserted.
"""
M1 = 4294967087
M2 = 4294944443
A12, A13N = 1403580, 810728
A21, A23N = 527612, 1370589
NORM = 2.328306549295727688e-10          # == 1.0 / (M1 + 1), same double

# --- the transition matrices, from the recurrence itself ------------------
# x1_n = a12*x1_{n-2} - a13n*x1_{n-3}   (mod m1)
# x2_n = a21*x2_{n-1} - a23n*x2_{n-3}   (mod m2)
A1 = [[0, A12, -A13N % M1], [1, 0, 0], [0, 1, 0]]
A2 = [[A21 % M2, 0, -A23N % M2], [1, 0, 0], [0, 1, 0]]


class ExactMrg:
    """Path 1. Unbounded integers."""

    def __init__(self, state):
        self.s = list(state)                      # [s10,s11,s12,s20,s21,s22]

    def next_u(self):
        s10, s11, s12, s20, s21, s22 = self.s
        p1 = (A12 * s11 - A13N * s10) % M1
        p2 = (A21 * s22 - A23N * s20) % M2
        self.s = [s11, s12, p1, s21, s22, p2]
        if p1 <= p2:
            return (p1 - p2 + M1) * NORM
        return (p1 - p2) * NORM


class DoubleMrg:
    """Path 2. The VBA-shaped arithmetic: floats holding integers, Fix()."""

    def __init__(self, state):
        self.s = [float(v) for v in state]

    @staticmethod
    def _reduce(p, m):
        k = float(int(p / m))                     # VBA Fix(): truncate to zero
        p = p - k * m
        if p < 0.0:
            p += m
        return p

    def next_u(self):
        s10, s11, s12, s20, s21, s22 = self.s
        p1 = self._reduce(A12 * s11 - A13N * s10, M1)
        p2 = self._reduce(A21 * s22 - A23N * s20, M2)
        self.s = [s11, s12, p1, s21, s22, p2]
        if p1 <= p2:
            return (p1 - p2 + M1) * NORM
        return (p1 - p2) * NORM


# --- modular matrix arithmetic -------------------------------------------
def mat_mul_mod(a, b, m):
    """Exact-integer 3x3 modular product. Path 1 for the jump."""
    return [[sum(a[i][k] * b[k][j] for k in range(3)) % m for j in range(3)]
            for i in range(3)]


def mat_pow2_mod(a, e, m):
    """a^(2^e) mod m, by repeated squaring."""
    r = [row[:] for row in a]
    for _ in range(e):
        r = mat_mul_mod(r, r, m)
    return r


H = 1 << 17


def mult_mod_m(a, s, c, m):
    """L'Ecuyer MultModM: (a*s + c) mod m without exceeding 2^53.

    Path 2 for the jump. Splits a into a1*H + a0 so no product reaches
    a*s directly. Written in float arithmetic to mirror what VBA would do.
    """
    a = float(a); s = float(s); c = float(c); m = float(m)
    a1 = float(int(a / H))
    if a1 != 0.0:
        a -= a1 * H
        v = a1 * s
        v -= float(int(v / m)) * m
        v = v * H + a * s + c
        v -= float(int(v / m)) * m
    else:
        v = a * s + c
        v -= float(int(v / m)) * m
    if v < 0.0:
        v += m
    return v


def mat_vec_mod_safe(a, v, m):
    """3x3 modular matrix-vector product using only MultModM. Path 2."""
    out = []
    for i in range(3):
        acc = 0.0
        for j in range(3):
            acc = mult_mod_m(a[i][j], v[j], acc, m)
        out.append(acc)
    return out


def mat_vec_mod_naive(a, v, m):
    """The forbidden form, kept ONLY as a negative control."""
    return [float((int(a[i][0]) * int(v[0]) + int(a[i][1]) * int(v[1])
                   + int(a[i][2]) * int(v[2])) % m) for i in range(3)]


# --- state jumping, with the vector convention stated explicitly ----------
# A1 and A2 operate on NEWEST-FIRST vectors: v = [x_{n-1}, x_{n-2}, x_{n-3}].
# PCCM stores state OLDEST-FIRST as [s10, s11, s12], so both directions are
# reversed here. This was determined empirically (see verify_convention) and
# not assumed: getting it wrong produces a plausible but wrong stream.
def jump_state(state, j1, j2, safe=True):
    mv = mat_vec_mod_safe if safe else mat_vec_mod_naive
    v1 = mv(j1, [float(x) for x in state[:3][::-1]], M1)
    v2 = mv(j2, [float(x) for x in state[3:][::-1]], M2)
    return [int(x) for x in v1][::-1] + [int(x) for x in v2][::-1]


def verify_convention(state):
    """One step by recurrence must equal one step by matrix."""
    g = ExactMrg(list(state)); g.next_u()
    return g.s == jump_state(state, A1, A2)


def stream_states(base_state, count, j1, j2):
    """Successive streams, each one 2^127 draws beyond the last."""
    out, cur = [], list(base_state)
    for _ in range(count):
        out.append(list(cur))
        cur = jump_state(cur, j1, j2)
    return out
