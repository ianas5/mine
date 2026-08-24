"""D6-03 auto_nonce -> effective_seed, and D6-05 seed -> six-word state."""
P = 2147483647          # 2^31 - 1, Mersenne prime
MULT = 48271            # candidate multiplier
SEED_MIN, SEED_MAX = 1, 2147483646


def factorise(n):
    fs, d = [], 2
    while d * d <= n:
        while n % d == 0:
            fs.append(d); n //= d
        d += 1
    if n > 1: fs.append(n)
    return sorted(set(fs))


def is_primitive_root(g, p):
    """g is a primitive root mod p iff g^((p-1)/q) != 1 for every prime q|p-1."""
    order = p - 1
    return all(pow(g, order // q, p) != 1 for q in factorise(order))


def nonce_to_seed(nonce, start=1):
    """auto_nonce -> effective_seed. Full-period Lehmer cycle over 1..P-1."""
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
