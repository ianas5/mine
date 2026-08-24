"""D6-17: `result_digest`, built from the ACCEPTED Phase-5 encoder.

Nothing here is a new hash, a new number format or a new grammar. The encoder,
the canonical Double form, the UTF-16 folding and both moduli are IMPORTED from
`pccm/builder/pccm_builder/calc_fingerprint.py`, which is the accepted Phase-5
single source. This module only chooses the FRAMING for a simulation result:
which fields, in which order, under which section name.

Direction of dependency matters: evidence imports production. No production code
is changed, and no evidence code is offered as production.
"""
import os
import sys

_BUILDER = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))))), "builder")
if _BUILDER not in sys.path:
    sys.path.insert(0, _BUILDER)

from pccm_builder.calc_fingerprint import (        # noqa: E402
    FP_BASE, FP_MOD_1, FP_MOD_2, FP_INIT_1, FP_INIT_2,
    canonical_number, encode_section, fingerprint, integer_field,
    number_field, text_field,
)

RESULT_STREAM_TAG = "PCCM-RD"
"""First field of a result stream. Distinct from the input fingerprint's
`PCCM-FP`, so an input stream and a result stream can never coincide."""

RESULT_SECTION = "RESULT"


def canon_double(x):
    """The accepted canonical Double text. Re-exported, not reimplemented."""
    return canonical_number(x)


def result_stream(totals_nominal, totals_pv, version=1):
    """`stream ::= F_S("PCCM-RD") F_I(version) section("RESULT", record*)`

    `record ::= F_I(field_count) F_I(iteration_index) F_N(nominal) F_N(pv)`

    The iteration count `n` is the section's record count and the iteration
    index is the record's first field, both hashed by the accepted grammar.
    Iteration order IS the record order: this digest is order-sensitive by
    construction, which is what makes G2/G3 replay comparisons meaningful.
    """
    if len(totals_nominal) != len(totals_pv):
        raise ValueError("nominal and PV totals must have the same length")
    records = [
        (integer_field(i + 1), number_field(totals_nominal[i]), number_field(totals_pv[i]))
        for i in range(len(totals_nominal))
    ]
    return (text_field(RESULT_STREAM_TAG).encode()
            + integer_field(version).encode()
            + encode_section(RESULT_SECTION, records))


def result_digest(totals_nominal, totals_pv, version=1):
    return fingerprint(result_stream(totals_nominal, totals_pv, version))
