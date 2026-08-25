"""PCCM Phase-6 Step-2 reference: the deterministic random-number backbone.

This is the PYTHON ORACLE for the RNG the accepted simulation contract
describes: scalar and AUTO seeding, one exact MRG32k3a step, the canonical
uniform, the `2^127` stream jump, and canonical component-stream assignment.

WHAT THIS IS NOT. It is not the simulation engine. There is no Uniform,
Triangular or Beta-PERT sampler here, no Bernoulli trial, no iteration, no
statistic, no digest - not even the Uniform DISTRIBUTION transform
`x = (1-u)a + ub`, which is Step 3. Step 2 produces raw MRG uniforms and stream
identities and nothing else.

--------------------------------------------------------------------------------
WHY EXACT PYTHON INTEGERS
--------------------------------------------------------------------------------
The reference arithmetic is arbitrary-precision integer arithmetic, including
the jump. That is deliberate and it is the whole point of the layering:

    Python exact integers are the ORACLE.
    The VBA-safe Double / MultModM decomposition is a LATER implementation that
    must prove itself against this oracle.

Writing the oracle in the same restricted arithmetic the implementation will use
would mean the two agree because they share a technique, which proves nothing
about either. So `MultModM` deliberately does NOT appear in this module.

--------------------------------------------------------------------------------
WHERE THE NUMBERS COME FROM
--------------------------------------------------------------------------------
`spec/sim_contract.yaml`, through an already-validated `SimContract`, plus the
seed admissibility owned by `spec/input_contract.yaml`. Nothing here holds a
second, independently maintained copy of a constant that could drift from the
contract, and nothing here reads `evidence/` - the retained Step-0 vectors are a
TEST oracle, and a reference that consulted them at run time would be marking its
own homework.

Fails loudly; never repairs. No global mutable state, no singleton, no hidden
seeding, and no dependency on a workbook, a build artefact, Excel or COM.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable, Sequence

from .calc_fingerprint import utf16_sort_key
from .contract_loader import ContractError, InputContract
from .sim_loader import SimContract

COST_KIND = "COST"
RISK_KIND = "RISK"
"""The two component-kind axis labels, as the retained Step-0 vectors spell them."""

ROLE_VALUE = "value"
ROLE_OCCURRENCE = "occurrence"
ROLE_SEVERITY = "severity"


class SimRngError(ContractError):
    """Raised when the RNG reference is asked for something the contract forbids.

    Subclasses ContractError so a specification-level refusal reports the same
    way every other PCCM authority failure does.
    """


# ---------------------------------------------------------------------------
# Values
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class RngState:
    """Six exact integers, exposed OLDEST-FIRST as `[s10, s11, s12, s20, s21, s22]`.

    Immutable on purpose. Advancing returns a new state, so a caller cannot
    accidentally share one generator between two components, and a test can hold
    a state across a jump without it moving underneath.

    NOT floating point. The later VBA representation uses exact-integer Doubles;
    that is a Step-2-and-later implementation concern, not the oracle's.
    """

    words: tuple[int, int, int, int, int, int]

    def __post_init__(self) -> None:
        if len(self.words) != 6:
            raise SimRngError(f"an MRG32k3a state has exactly six words, got {len(self.words)}")
        for index, word in enumerate(self.words):
            if not isinstance(word, int) or isinstance(word, bool):
                raise SimRngError(f"state word {index} must be an exact integer, got {word!r}")

    @classmethod
    def of(cls, *words: int) -> "RngState":
        return cls(tuple(words))  # type: ignore[arg-type]

    @property
    def first(self) -> tuple[int, int, int]:
        return self.words[0], self.words[1], self.words[2]

    @property
    def second(self) -> tuple[int, int, int]:
        return self.words[3], self.words[4], self.words[5]

    def as_list(self) -> list[int]:
        """The retained vectors are JSON lists; this is what compares against them."""
        return list(self.words)


@dataclass(frozen=True)
class Draw:
    """One step: the state AFTER the advance, and the uniform it produced."""

    state: RngState
    uniform: float


@dataclass(frozen=True)
class Component:
    """One consumer of one stream.

    `kind` is the DRIVER-kind axis (`COST` / `RISK`), and `role` is separate.
    Those are two different sort keys, and collapsing them into one - treating
    `COST_SAMPLE`, `RISK_OCCURRENCE` and `RISK_SEVERITY` as a single ordered axis
    - would produce three global blocks in which every occurrence stream precedes
    every severity stream. The accepted Step-0 vectors interleave them per Risk
    (`R-099 occurrence, R-099 severity, R-100 occurrence, …`), so the two axes
    stay separate. See `_kind_rank` and `_role_rank`.
    """

    kind: str
    permanent_id: str
    role: str

    def as_list(self) -> list[str]:
        return [self.kind, self.permanent_id, self.role]


# ---------------------------------------------------------------------------
# The reference
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class RngReference:
    """Operational constants, derived from the accepted contracts.

    Built by `from_contracts`, never by hand in production code: a reference that
    could be constructed with arbitrary constants would be a second authority.
    """

    m1: int
    m2: int
    a12: int
    a13n: int
    a21: int
    a23n: int
    norm: float
    jump_a1: tuple[tuple[int, int, int], ...]
    jump_a2: tuple[tuple[int, int, int], ...]
    auto_modulus: int
    auto_multiplier: int
    nonce_exhausted: int
    seed_min: int
    seed_max: int
    kind_order: tuple[str, ...]
    role_order: dict[str, tuple[str, ...]]

    # -- construction -------------------------------------------------------
    @classmethod
    def from_contracts(cls, sim: SimContract, inputs: InputContract) -> "RngReference":
        """Derive every operational value from the two owning authorities.

        The simulation contract owns the RNG; `input_contract.yaml` owns the
        admissible seed domain (D6-19a). Reading the domain from its owner is
        what stops this module becoming a place the range can drift.
        """
        raw = sim.raw
        constants = raw["rng"]["constants"]
        auto = raw["seeding"]["auto"]
        lifecycle = raw["seeding"]["nonce_lifecycle"]

        seed_min, seed_max = _seed_domain(inputs)

        kinds = raw["components"]["kinds"]
        kind_order: list[str] = []
        role_order: dict[str, list[str]] = {}
        for entry in kinds:
            axis = COST_KIND if entry["driver_kind"] == "cost_line" else RISK_KIND
            if axis not in kind_order:
                kind_order.append(axis)
                role_order[axis] = []
            role_order[axis].append(entry["role"])

        # The kind axis is stated twice - by the order of `components.kinds` and
        # by `accumulation.driver_kind_order` - so the two must agree or one of
        # them is silently unused.
        declared = [
            COST_KIND if k == "cost_line" else RISK_KIND
            for k in raw["accumulation"]["driver_kind_order"]
        ]
        if declared != kind_order:
            raise SimRngError(
                f"the contract states the component-kind axis twice and they disagree: "
                f"components.kinds implies {kind_order}, accumulation.driver_kind_order "
                f"says {declared}"
            )

        return cls(
            m1=int(constants["m1"]),
            m2=int(constants["m2"]),
            a12=int(constants["a12"]),
            a13n=int(constants["a13n"]),
            a21=int(constants["a21"]),
            a23n=int(constants["a23n"]),
            norm=float(constants["norm"]),
            jump_a1=_matrix(raw["jump"]["a1_p127"]),
            jump_a2=_matrix(raw["jump"]["a2_p127"]),
            auto_modulus=int(auto["modulus"]),
            auto_multiplier=int(auto["multiplier"]),
            nonce_exhausted=int(lifecycle["exhausted_value"]),
            seed_min=seed_min,
            seed_max=seed_max,
            kind_order=tuple(kind_order),
            role_order={k: tuple(v) for k, v in role_order.items()},
        )

    # -- state --------------------------------------------------------------
    def validate_state(self, state: RngState) -> RngState:
        """Refuse a state the recurrence cannot legally be in.

        Checked at the public boundary rather than assumed: an out-of-range or
        all-zero component produces a stream that looks fine and is not
        MRG32k3a, and the all-zero case is absorbing.
        """
        if not isinstance(state, RngState):
            raise SimRngError(f"expected an RngState, got {type(state).__name__}")
        for label, words, modulus in (
            ("first", state.first, self.m1),
            ("second", state.second, self.m2),
        ):
            for index, word in enumerate(words):
                if not 0 <= word < modulus:
                    raise SimRngError(
                        f"{label} component word {index} is {word}, outside [0, {modulus})"
                    )
            if not any(words):
                raise SimRngError(
                    f"the {label} component is all zero, which is an absorbing state: the "
                    "recurrence can never leave it"
                )
        return state

    # -- seeding ------------------------------------------------------------
    def fixed_seed_to_state(self, seed: int) -> RngState:
        """D6-05 (a): the scalar repeated into all six words. No mixer, no hash."""
        if not isinstance(seed, int) or isinstance(seed, bool):
            raise SimRngError(
                f"a FIXED seed must be a whole integer, got {seed!r}. A bool is not an "
                "integer here even though Python says it is."
            )
        if not self.seed_min <= seed <= self.seed_max:
            raise SimRngError(
                f"FIXED seed {seed} is outside the admissible domain "
                f"[{self.seed_min}, {self.seed_max}] owned by input_contract.yaml"
            )
        return self.validate_state(RngState((seed,) * 6))  # type: ignore[arg-type]

    def auto_seed_from_nonce(self, nonce: int) -> int:
        """D6-03 (b): `effective_seed = multiplier^nonce mod modulus`.

        A modular POWER, evaluated exactly in O(log nonce). Stepping the cycle
        `nonce` times would give the same answer and is not the authority: at a
        large nonce it is unusable, and stating the mapping as a power is what
        tells a later implementation to square-and-multiply.

        PURE. It persists nothing and increments nothing; the transactional nonce
        lifecycle belongs to the later engine boundary, not here.
        """
        if not isinstance(nonce, int) or isinstance(nonce, bool):
            raise SimRngError(f"an AUTO nonce must be a whole integer, got {nonce!r}")
        if nonce < 0:
            raise SimRngError(f"an AUTO nonce must not be negative, got {nonce}")
        if nonce >= self.nonce_exhausted:
            raise SimRngError(
                f"AUTO nonce {nonce} is exhausted: the cycle has period "
                f"{self.nonce_exhausted}, so allocating it would silently reissue the seed "
                f"for nonce {nonce - self.nonce_exhausted}"
            )
        return pow(self.auto_multiplier, nonce, self.auto_modulus)

    # -- one step -----------------------------------------------------------
    def next_uniform(self, state: RngState) -> Draw:
        """One exact MRG32k3a step, returning the new state and the uniform."""
        s10, s11, s12, s20, s21, s22 = self.validate_state(state).words
        p1 = (self.a12 * s11 - self.a13n * s10) % self.m1
        p2 = (self.a21 * s22 - self.a23n * s20) % self.m2
        advanced = RngState((s11, s12, p1, s21, s22, p2))
        if p1 <= p2:
            uniform = (p1 - p2 + self.m1) * self.norm
        else:
            uniform = (p1 - p2) * self.norm
        if not 0.0 < uniform < 1.0:
            raise SimRngError(
                f"the combination produced {uniform!r}, outside the open interval (0, 1) the "
                "contract requires. Both endpoints are excluded by construction, so this "
                "cannot happen for a valid state - it means the state or a constant is wrong."
            )
        return Draw(state=advanced, uniform=uniform)

    def uniforms(self, state: RngState, count: int) -> tuple[tuple[float, ...], RngState]:
        """`count` successive uniforms and the state that follows them."""
        if not isinstance(count, int) or isinstance(count, bool) or count < 0:
            raise SimRngError(f"count must be a non-negative integer, got {count!r}")
        # Validate ONCE up front, including at count == 0. Without this the
        # zero-draw case returned whatever it was handed, because the loop that
        # would have validated never ran - a defensive hole that only widens once
        # samplers start passing states through this API.
        self.validate_state(state)
        drawn: list[float] = []
        current = state
        for _ in range(count):
            step = self.next_uniform(current)
            drawn.append(step.uniform)
            current = step.state
        return tuple(drawn), current

    # -- jumping ------------------------------------------------------------
    def jump_to_next_stream(self, state: RngState) -> RngState:
        """Advance one canonical `2^127` stream jump.

        ORIENTATION, stated because getting it wrong is silent. PCCM stores and
        exposes state OLDEST-FIRST; the jump matrices operate on NEWEST-FIRST
        triples. Each triple is therefore reversed on the way in and reversed
        back on the way out. A transpose or a dropped reversal produces a
        plausible stream that is not the canonical one, so the accepted Step-0
        vectors - not recollection - are what settle this.
        """
        self.validate_state(state)
        first = _mat_vec_mod(self.jump_a1, tuple(reversed(state.first)), self.m1)
        second = _mat_vec_mod(self.jump_a2, tuple(reversed(state.second)), self.m2)
        return self.validate_state(
            RngState(tuple(reversed(first)) + tuple(reversed(second)))  # type: ignore[arg-type]
        )

    def stream_initial_state(self, base_state: RngState, k: int) -> RngState:
        """The initial state of stream `k`: the base advanced by `k` jumps.

        Applied repeatedly, which is right for Step 2's component counts - 400 at
        the design target. No precomputed table: a table would be a second copy
        of the same fact, and it would go stale against the matrices.
        """
        if not isinstance(k, int) or isinstance(k, bool) or k < 0:
            raise SimRngError(f"a stream index must be a non-negative integer, got {k!r}")
        current = self.validate_state(base_state)
        for _ in range(k):
            current = self.jump_to_next_stream(current)
        return current

    # -- component streams --------------------------------------------------
    def _kind_rank(self, kind: str) -> int:
        try:
            return self.kind_order.index(kind)
        except ValueError:
            raise SimRngError(
                f"unknown component kind {kind!r}; the contract declares {list(self.kind_order)}"
            ) from None

    def _role_rank(self, kind: str, role: str) -> int:
        roles = self.role_order.get(kind, ())
        try:
            return roles.index(role)
        except ValueError:
            raise SimRngError(
                f"kind {kind!r} has no role {role!r}; the contract declares {list(roles)}"
            ) from None

    def canonical_sort_key(self, component: Component) -> tuple[int, tuple[int, ...], int]:
        """`(component_kind, permanent_id, role)` - the accepted D6-16 order.

        Three SEPARATE keys. The Permanent ID is compared on ordinal UTF-16 code
        units using the accepted Phase-5 sort key, not Python's own string
        ordering, not locale collation, not case-insensitively, and not by
        reading the numeric suffix - so `CL-1000` sorts BEFORE `CL-999`.
        """
        return (
            self._kind_rank(component.kind),
            utf16_sort_key(component.permanent_id),
            self._role_rank(component.kind, component.role),
        )

    def components_for(
        self, cost_line_ids: Iterable[str], risk_ids: Iterable[str]
    ) -> tuple[Component, ...]:
        """The component set. Cost Line -> 1; Risk -> occurrence + severity."""
        out: list[Component] = []
        for identifier in cost_line_ids:
            out.append(Component(COST_KIND, identifier, ROLE_VALUE))
        for identifier in risk_ids:
            for role in self.role_order.get(RISK_KIND, ()):
                out.append(Component(RISK_KIND, identifier, role))
        return tuple(out)

    def assign_component_streams(
        self, components: Sequence[Component]
    ) -> tuple[tuple[Component, int], ...]:
        """Assign stream indices `0 … N-1` in canonical order.

        IDENTITY ONLY. Nothing is sampled from these streams in Step 2.

        The order is a SORT, so the physical row order of the registers cannot
        reach it. The accepted consequence is preserved rather than smoothed
        over: inserting or removing a driver changes the canonical order, so
        later components move to different streams and their samples change.
        """
        ordered = sorted(components, key=self.canonical_sort_key)
        seen: set[tuple[str, str, str]] = set()
        for component in ordered:
            identity = (component.kind, component.permanent_id, component.role)
            if identity in seen:
                raise SimRngError(f"duplicate component {identity}")
            seen.add(identity)
        return tuple((component, index) for index, component in enumerate(ordered))

    def component_stream_states(
        self, base_state: RngState, components: Sequence[Component]
    ) -> tuple[tuple[Component, int, RngState], ...]:
        """Each component with its stream index and that stream's initial state.

        Walks the jump ladder once rather than re-deriving each stream from the
        base, which is the same states in `N` jumps instead of `N(N-1)/2`.
        """
        assignment = self.assign_component_streams(components)
        out: list[tuple[Component, int, RngState]] = []
        current = self.validate_state(base_state)
        for component, index in assignment:
            if index != len(out):  # pragma: no cover - assignment is 0..N-1 by construction
                raise SimRngError("stream indices are not contiguous from zero")
            out.append((component, index, current))
            current = self.jump_to_next_stream(current)
        return tuple(out)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _matrix(rows: Any) -> tuple[tuple[int, int, int], ...]:
    if not isinstance(rows, list) or len(rows) != 3:
        raise SimRngError(f"a jump matrix must have three rows, got {rows!r}")
    out = []
    for row in rows:
        if not isinstance(row, list) or len(row) != 3:
            raise SimRngError(f"a jump matrix row must have three elements, got {row!r}")
        out.append(tuple(int(v) for v in row))
    return tuple(out)  # type: ignore[return-value]


def _mat_vec_mod(
    matrix: tuple[tuple[int, int, int], ...], vector: tuple[int, int, int], modulus: int
) -> tuple[int, int, int]:
    """Exact-integer modular matrix-vector product.

    Arbitrary precision on purpose. The naive form overflows a Double by three
    orders of magnitude, which is exactly why the VBA implementation needs a
    decomposition - and exactly why the ORACLE must not use one.
    """
    return tuple(  # type: ignore[return-value]
        (row[0] * vector[0] + row[1] * vector[1] + row[2] * vector[2]) % modulus
        for row in matrix
    )


def _seed_domain(inputs: InputContract) -> tuple[int, int]:
    """The admissible FIXED seed domain, read from the contract that OWNS it."""
    seed = inputs.inputs.get("random_seed")
    if seed is None:
        raise SimRngError("input_contract.yaml declares no 'random_seed' input")
    validation = seed.validation
    if not isinstance(validation, dict):
        raise SimRngError(
            "input_contract.yaml leaves random_seed.validation unset, so the admissible "
            "domain has no owner"
        )
    if validation.get("kind") != "whole" or validation.get("operator") != "between":
        raise SimRngError(
            f"random_seed.validation must be a whole-number 'between' rule, got "
            f"kind={validation.get('kind')!r} operator={validation.get('operator')!r}"
        )
    try:
        low = int(str(validation["formula1"]))
        high = int(str(validation["formula2"]))
    except (KeyError, ValueError) as error:
        raise SimRngError(f"random_seed.validation bounds are not whole numbers: {error}") from error
    if low > high:
        raise SimRngError(f"random_seed.validation bounds are inverted: {low} > {high}")
    return low, high
