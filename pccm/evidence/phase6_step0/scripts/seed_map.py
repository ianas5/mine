"""D6-03 auto_nonce -> effective_seed, and D6-05 seed -> six-word state."""
P = 2147483647          # 2^31 - 1, Mersenne prime
MULT = 48271            # candidate multiplier
SEED_MIN, SEED_MAX = 1, 2147483646

# --- D6-03 nonce lifecycle constants -------------------------------------
# `auto_nonce` is THE NEXT NONCE TO ALLOCATE, not the last one used. The
# distinction is the whole lifecycle: a persisted "next" value is meaningful
# before any run has happened, and allocation is read-then-advance rather than
# advance-then-read.
NONCE_INITIAL = 0
"""Persisted value in a workbook that has never allocated an AUTO seed."""

NONCE_EXHAUSTED = 2147483646
"""The first value that must NOT be allocated: 48271^2147483646 == 48271^0 == 1,
so allocating it would silently reissue the very first seed."""


def factorise_with_multiplicity(n):
    """The COMPLETE prime factorisation, as (prime, exponent) pairs.

    Kept distinct from `distinct_prime_divisors` below because the two are
    different facts and conflating them is an arithmetic error: the product of
    the distinct divisors of 2147483646 is 715827882, not 2147483646.
    """
    fs, d = [], 2
    while d * d <= n:
        e = 0
        while n % d == 0:
            e += 1
            n //= d
        if e:
            fs.append((d, e))
        d += 1
    if n > 1:
        fs.append((n, 1))
    return fs


def distinct_prime_divisors(n):
    """The distinct primes dividing n. This is what the primitive-root test needs."""
    return [p for p, _ in factorise_with_multiplicity(n)]


def factorise(n):
    """Retained name, retained meaning: the DISTINCT prime divisors."""
    return distinct_prime_divisors(n)


def is_primitive_root(g, p):
    """g is a primitive root mod p iff g^((p-1)/q) != 1 for every DISTINCT
    prime q dividing p-1. Multiplicity is irrelevant to the test."""
    order = p - 1
    return all(pow(g, order // q, p) != 1 for q in distinct_prime_divisors(order))


def nonce_to_seed(nonce, start=1):
    """THE AUTHORITY: effective_seed = start * MULT^nonce mod P.

    Exact modular exponentiation. The mathematical definition is a power, not a
    loop, and stating it as a power is what tells a later implementation that an
    O(log nonce) square-and-multiply is required rather than O(nonce) repeated
    multiplication. See `nonce_to_seed_iterative` for the cross-check.
    """
    return (start * pow(MULT, nonce, P)) % P


def nonce_to_seed_iterative(nonce, start=1):
    """Second independent path: the Lehmer cycle stepped one multiplication at a
    time. Agrees with `nonce_to_seed` by construction; used to check it, and
    NOT proposed as an implementation - it is O(nonce)."""
    x = start
    for _ in range(nonce):
        x = (MULT * x) % P
    return x


def seed_to_state(seed):
    """D6-05 candidate A: the scalar repeated into all six state words."""
    return [seed] * 6


def state_is_valid(state, m1=4294967087, m2=4294944443):
    s1, s2 = state[:3], state[3:]
    return (all(0 <= v < m1 for v in s1) and all(0 <= v < m2 for v in s2)
            and any(s1) and any(s2))
