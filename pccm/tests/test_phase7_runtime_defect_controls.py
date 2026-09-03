#!/usr/bin/env python3
"""The controls derived from the P7-4 Windows runtime defect of 0734a38.

WHAT HAPPENED
-------------
The approved intermediate timing run built the current Phase-7 workbook, ran a
successful 10,000-iteration simulation (bank A, run 1, status CURRENT), and
called the real endpoint. It failed in 70 ms with

    FAIL|Run Sensitivity|Subscript out of range

and the harness, reading the CONTRACTED sensitivity stamp, came back with
`Mean`, `Sample Standard Deviation`, `Minimum`, `P10`, `P50`, `P55`, `P60`.

TWO INDEPENDENT DEFECTS, AND NEITHER DETECTOR EXISTED
-----------------------------------------------------
1.  `modSimSensitivity.SimSensitivitySortAscending` decided which run to draw
    from with ONE expression:

        If fromLow < midPoint And (fromHigh >= highEnd Or _
           Not (series(fromHigh) < series(fromLow))) Then

    VBA's `And` and `Or` evaluate BOTH operands whatever the first answers.
    `series(fromHigh)` was therefore read even after `fromHigh >= highEnd` had
    said the right run was spent, and in the last merge of a pass `highEnd` IS
    `count`, so the read is `series(count)` on an array `0 To count - 1`:
    run-time error 9. 3,876 static tests were green because the transcriber
    compiles VBA `And`/`Or` to PYTHON's `and`/`or`, which short-circuit, so the
    ported vectors never evaluated the out-of-range operand.

2.  The P7-1 sensitivity block was allocated at J-Q and S-Z with its stamp on
    rows 8-14 of each bank's first column. J8:J14 is the summary-statistics
    LABEL column and S8:S14 is published contingency-ladder bank-B nominal
    data. The only occupancy check that existed compared a Phase-7 column
    against the six ITERATION bank columns, which J and S are not.

The controls below decide the PROPERTIES those defects violated, so neither
class can return, and each is derived from the contract or the source rather
than from a list somebody has to remember to extend.
"""

from __future__ import annotations

import glob
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pytest
import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import load_sim_contract  # noqa: E402
from pccm_builder.sim_loader import SimContractError  # noqa: E402
from pccm_builder.vba_source import logical_statements, strip_comments  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"
TESTS = PCCM_ROOT / "tests"
CONTRACT = SPEC / "sim_contract.yaml"
TRANSCRIBER = TESTS / "phase6_vba_transcribe.py"

# VBA intrinsics and type coercions. A call to one of these inside a logical
# operand cannot be an out-of-range array read, so naming them keeps the
# short-circuit rule pointed at subscripts.
_INTRINSICS = frozenset({
    "if", "then", "elseif", "and", "or", "not", "cstr", "clng", "cdbl", "cbool",
    "cint", "abs", "len", "trim", "mid", "ascw", "asc", "sqr", "log", "exp",
    "fix", "int", "sgn", "strcomp", "isempty", "isnull", "iserror", "isobject",
    "isnumeric", "isarray", "isdate", "vartype", "typename", "lbound", "ubound",
    "chr", "chrw", "left", "right", "instr", "space", "string", "val", "err",
    "set", "iif", "array", "cvar", "clngptr", "csng",
})


def _text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _statements(path: Path) -> list[str]:
    return [item[1] if isinstance(item, tuple) else item
            for item in logical_statements(strip_comments(_text(path)))]


def _raw_contract() -> dict:
    with CONTRACT.open(encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def _column_number(letter: str) -> int:
    value = 0
    for char in str(letter).strip().upper():
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def _column_letter(number: int) -> str:
    out = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        out = chr(ord("A") + remainder) + out
    return out


def _split_operands(text: str) -> list[str]:
    """Top-level `And`/`Or` operands of one VBA expression, parens respected."""
    operands: list[str] = []
    depth = 0
    current = ""
    index = 0
    while index < len(text):
        char = text[index]
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        if depth == 0:
            match = re.match(r"\b(And|Or)\b", text[index:])
            if match:
                operands.append(current)
                current = ""
                index += match.end()
                continue
        current += char
        index += 1
    operands.append(current)
    return operands


def _all_operands(text: str) -> list[str]:
    """Every operand at every nesting level, so an inner `Or` is examined too."""
    out: list[str] = []
    pending = [text]
    while pending:
        piece = pending.pop()
        parts = _split_operands(piece)
        if len(parts) > 1:
            out.extend(parts)
            for part in parts:
                stripped = part.strip()
                # Descend through a fully parenthesised operand, and through a
                # `Not (...)`, so `A And (B Or C(i))` reaches B and C(i).
                inner = re.fullmatch(r"(?:Not\s*)?\((.*)\)", stripped, re.S)
                if inner:
                    pending.append(inner.group(1))
    return out


def _subscripted_indices(operand: str) -> set[str]:
    """The simple variables used as subscripts inside `name(...)` in an operand."""
    found: set[str] = set()
    for name, args in re.findall(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\(([^()]*)\)", operand):
        if name.lower() in _INTRINSICS:
            continue
        for token in re.findall(r"\b([A-Za-z_][A-Za-z0-9_]*)\b", args):
            if token.lower() not in _INTRINSICS:
                found.add(token)
    return found


def _compared_names(operand: str) -> set[str]:
    """Variables that appear on either side of a relational comparison."""
    found: set[str] = set()
    for left, _op, right in re.findall(
            r"([A-Za-z_][A-Za-z0-9_.]*)\s*(<=|>=|<>|<|>)\s*([A-Za-z_][A-Za-z0-9_.]*)",
            operand):
        found.add(left.split(".")[0])
        found.add(right.split(".")[0])
    return found


def _short_circuit_violations() -> list[tuple[str, str, set[str]]]:
    """Statements whose correctness needs an evaluation rule VBA does not have.

    THE PROPERTY, NOT A TOKEN LIST. An expression is a violation when one
    operand SUBSCRIPTS an array with a variable that another operand of the same
    `And`/`Or` COMPARES - that is precisely the shape of "the guard bounds the
    index the other side reads", and it is the shape that only works if the
    guard can stop the read. In VBA it cannot: both operands are always
    evaluated.
    """
    violations: list[tuple[str, str, set[str]]] = []
    for path in sorted(glob.glob(str(SRC_VBA / "*.bas"))):
        for statement in _statements(Path(path)):
            if not re.search(r"\b(And|Or)\b", statement):
                continue
            operands = _all_operands(statement)
            if len(operands) < 2:
                continue
            for position, operand in enumerate(operands):
                indices = _subscripted_indices(operand)
                if not indices:
                    continue
                guarded = set()
                for other_position, other in enumerate(operands):
                    if other_position == position:
                        continue
                    guarded |= (indices & _compared_names(other))
                if guarded:
                    violations.append((Path(path).name, statement, guarded))
    return violations


# ===========================================================================
# 1. THE SHORT-CIRCUIT CLASS
# ===========================================================================
def test_01_no_vba_expression_relies_on_short_circuit_evaluation() -> None:
    violations = _short_circuit_violations()
    assert not violations, (
        "these expressions read an array with an index another operand of the "
        "same And/Or is guarding. VBA evaluates both operands unconditionally, "
        "so the guard cannot stop the read and the subscript is evaluated out "
        "of range (run-time error 9). Restructure into nested If/ElseIf, where "
        "each branch is reached only after the branch above it has excluded the "
        "out-of-range case:\n" + "\n".join(
            f"  {name}: {statement}   (guarded index: {sorted(guarded)})"
            for name, statement, guarded in violations)
    )


def test_02_the_merge_decides_exhaustion_before_it_reads_either_side() -> None:
    """The specific regression, as an ORDERING of branches.

    Checked as structure rather than as the absence of the old expression: a
    future rewrite that reintroduced the fault with different variable names
    would not contain the old text, and would still be the same defect.
    """
    statements = _statements(SRC_VBA / "modSimSensitivity.bas")
    merge = [s for s in statements if "scratch(target) = series(" in s]
    assert len(merge) == 4, (
        "the merge must draw from exactly four branches - left run spent, right "
        f"run spent, and the two comparisons - found {len(merge)}: {merge}")
    guards = [s for s in statements
              if re.fullmatch(r"(?:If|ElseIf) (from(?:Low|High) >= \w+) Then", s.strip())]
    assert len(guards) == 2, (
        f"both exhaustion tests must be their own branch; found {guards}")
    positions = [statements.index(s) for s in statements
                 if s.strip().startswith(("If fromLow >= midPoint",
                                          "ElseIf fromHigh >= highEnd",
                                          "ElseIf series(fromHigh) <"))]
    assert positions == sorted(positions) and len(positions) == 3, (
        "the exhaustion branches must precede the branch that reads both sides, "
        f"got {positions}")


def test_03_the_transcriber_cannot_catch_this_class_and_says_so() -> None:
    """The detector gap is named where the detector lives.

    The transcriber rewrites `And` to Python's `and`. That is not a bug in it -
    it is a transcription of source, not a VBA interpreter - but it does mean a
    green ported vector is NO evidence about this class, and test_01 is what
    supplies that evidence instead. Pinning the rewrite here stops anyone from
    later concluding that the ported suites already cover it.
    """
    source = _text(TRANSCRIBER)
    assert re.search(r'sub\(r"\\bAnd\\b", " and ", text\)', source), (
        "the transcriber no longer rewrites And to Python's `and`; if it now "
        "models VBA's unconditional evaluation, say so here and keep test_01 "
        "as the source-level rule")
    assert re.search(r"short[- ]circuit", source, re.I), (
        "the transcriber must record that its And/Or rewrite adopts Python's "
        "short-circuit semantics, which VBA does not have")


# ===========================================================================
# 2. THE IDENTITY DOMAINS
# ===========================================================================
def test_04_each_published_identity_field_is_read_through_its_own_domain() -> None:
    statements = _statements(SRC_VBA / "modSimPostReport.bas")
    joined = "\n".join(statements)
    assert "IsWholeInRange(raw, minimum, maximum, measured)" in joined, (
        "the snapshot reader must apply the bounds its caller names, not a "
        "ceiling of its own")
    expected = {
        "SIM_IDENTITY_ROW_RUN_ID": ("SIM_RUN_ID_FIRST", "SIM_RUN_ID_MAXIMUM"),
        "SIM_IDENTITY_ROW_EFFECTIVE_SEED": ("SIM_SEED_MIN", "SIM_SEED_MAX"),
        "SIM_IDENTITY_ROW_ITERATIONS_RUN": ("SIM_MIN_ITERATIONS", "SIM_MAX_ITERATIONS"),
    }
    calls = re.findall(r"ReadSnapshotLong\(bank, (\w+), CDbl\((\w+)\), *_?\s*CDbl\((\w+)\)",
                       joined)
    assert len(calls) == 3, f"expected three bounded reads, found {calls}"
    for row_constant, low, high in calls:
        assert row_constant in expected, row_constant
        assert (low, high) == expected[row_constant], (
            f"{row_constant} must be bounded by {expected[row_constant]}, not ({low}, {high})")


def test_05_the_auto_seed_domain_really_does_exceed_the_iteration_ceiling() -> None:
    """The old bound was not merely untidy - it was wrong, and computably so.

    The mapping is the contract's own, and the iteration ceiling is derived
    from the loader, so this test states no number of its own.
    """
    raw = _raw_contract()
    auto = raw["seeding"]["auto"]
    modulus = int(auto["modulus"])
    multiplier = int(auto["multiplier"])
    ceiling = load_sim_contract(CONTRACT).max_iterations_representable
    seeds = {nonce: pow(multiplier, nonce, modulus) for nonce in range(6)}
    assert seeds[0] <= ceiling and seeds[1] <= ceiling, (
        "the first two AUTO runs happened to fall under the iteration ceiling, "
        "which is why the timing run got as far as the sort at all")
    over = [nonce for nonce, seed in seeds.items() if seed > ceiling]
    assert over, (
        "if no AUTO nonce could produce a seed above the iteration ceiling the "
        "old bound would have been harmless; it can, and this is the evidence")
    assert min(over) == 2, (
        f"the third AUTO run is the first one the old bound would have refused; "
        f"first offending nonce is {min(over)}")


def test_06_the_persisted_iteration_count_cannot_make_a_single_cell_read() -> None:
    """`Range("C34:C34").Value2` is a SCALAR, not a 1x1 array.

    `ReadTotals` indexes what it reads as `block(index, 1)`, which is only
    valid for a multi-cell range. What makes that safe is the lower bound on
    the persisted iteration count, so the bound is the proof and there is no
    guard pretending to be one.
    """
    from pccm_builder import load_contract
    from pccm_builder.sim_emit import business_minimum_iterations
    minimum = business_minimum_iterations(load_contract(SPEC / "input_contract.yaml"))
    assert minimum > 1, (
        f"the business minimum is {minimum}; at 1 the persisted total would be a "
        "single cell, Range(...).Value2 would be a scalar, and ReadTotals would "
        "index it as an array")
    statements = _statements(SRC_VBA / "modSimPostReport.bas")
    assert any("block(index, 1)" in s for s in statements), (
        "ReadTotals no longer indexes the persisted total as an array; if it "
        "changed, this proof needs to change with it")


# ===========================================================================
# 3. THE SHEET OCCUPANCY
# ===========================================================================
def _footprints() -> list[tuple[str, str, int, int]]:
    from pccm_builder.sim_loader import _sim_data_footprints
    raw = _raw_contract()["sim_data"]
    return _sim_data_footprints(raw, "test", int(raw["iteration_records"]["header_row"]),
                               int(raw["iteration_records"]["first_iteration_row"]))


def test_07_no_two_sim_data_blocks_reserve_the_same_cell() -> None:
    by_column: dict[str, list[tuple[str, int, int]]] = {}
    for owner, column, low, high in _footprints():
        by_column.setdefault(column.upper(), []).append((owner, low, high))
    collisions = []
    for column, entries in by_column.items():
        for index, (owner, low, high) in enumerate(entries):
            for other, olow, ohigh in entries[index + 1:]:
                if owner == other:
                    continue
                if max(low, olow) <= min(high, ohigh):
                    collisions.append(f"{owner} x {other} at {column}"
                                      f"{max(low, olow)}:{column}{min(high, ohigh)}")
    assert not collisions, "\n".join(collisions)


def test_08_the_sensitivity_stamp_cells_belong_to_nothing_else() -> None:
    """Named separately from the sweep because it is the cell that actually lied.

    A sweep that stopped covering the stamp would still pass; this would not.
    """
    records = _raw_contract()["sim_data"]["sensitivity_records"]
    stamp = records["stamp"]
    rows = [int(field["row"]) for field in stamp["fields"]]
    mine = {f"sensitivity_stamp[{bank}]" for bank in ("A", "B")}
    for bank, column in stamp["bank_value_columns"].items():
        for row in rows:
            owners = {owner for owner, col, low, high in _footprints()
                      if col.upper() == column.upper() and low <= row <= high
                      and not owner.startswith(tuple(mine))}
            assert not owners, (
                f"the bank {bank} stamp cell {column}{row} is also reserved by "
                f"{sorted(owners)}; on Windows that cell answered with the other "
                "block's content")


def test_09_the_stamp_column_is_the_banks_own_first_column() -> None:
    records = _raw_contract()["sim_data"]["sensitivity_records"]
    for bank, span in records["banks"].items():
        assert records["stamp"]["bank_value_columns"][bank] == span["first_column"], (
            f"the bank {bank} stamp must sit in the block's own first column, so "
            "the block occupies one contiguous column band at every row")
    first = records["columns"][0]["column"]
    assert records["banks"]["A"]["first_column"] == first


def test_10_the_loader_refuses_the_allocation_the_windows_run_exposed() -> None:
    """The P7-1 choice, reconstructed - contract AND locked pins together.

    Pinning the chosen letters does not prove the choice was sound: the pin and
    the contract agreed perfectly while both were wrong. This puts J-Q/S-Z back
    in both places and requires the loader to refuse on the CELL, which is the
    control that did not exist when the choice was made.
    """
    columns = [("driver_id", "J"), ("driver_type", "K"), ("driver_name", "L"),
               ("rho", "M"), ("abs_rho", "N"), ("rank", "O"),
               ("direction", "P"), ("status", "Q")]
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        shutil.copytree(SPEC, root / "spec")
        broken = root / "spec" / "sim_contract.yaml"
        text = broken.read_text(encoding="utf-8")
        text = text.replace('A: {first_column: "CC", last_column: "CJ"}',
                            'A: {first_column: "J", last_column: "Q"}')
        text = text.replace('B: {first_column: "CL", last_column: "CS"}',
                            'B: {first_column: "S", last_column: "Z"}')
        text = text.replace('        A: "CC"\n        B: "CL"',
                            '        A: "J"\n        B: "S"')
        current = [c["column"] for c in _raw_contract()["sim_data"]
                   ["sensitivity_records"]["columns"]]
        for (key, letter), now in zip(columns, current):
            text = text.replace(f'column: "{now}", header:', f'column: "{letter}", header:')
        broken.write_text(text, encoding="utf-8")

        # The locked pins move with it, so the refusal cannot come from them.
        script = (
            "import sys; sys.path.insert(0, %r)\n"
            "from pccm_builder import sim_loader\n"
            "sim_loader.LOCKED_SENSITIVITY_STAMP_COLUMNS = {'A': 'J', 'B': 'S'}\n"
            "sim_loader.LOCKED_SENSITIVITY_COLUMN_LAYOUT = (%r)\n"
            "from pathlib import Path\n"
            "try:\n"
            "    sim_loader.load_sim_contract(Path(%r))\n"
            "except sim_loader.SimContractError as error:\n"
            "    print('REFUSED:', error)\n"
            "else:\n"
            "    print('LOADED')\n"
        ) % (str(PCCM_ROOT / "builder"),
             tuple((key, letter, header, kind) for (key, letter), (_k, _l, header, kind)
                   in zip(columns, _raw_contract()["sim_data"]["sensitivity_records"]
                          and [(c["key"], c["column"], c["header"], c["value_type"])
                               for c in _raw_contract()["sim_data"]
                               ["sensitivity_records"]["columns"]])),
             str(broken))
        completed = subprocess.run([sys.executable, "-c", script],
                                   capture_output=True, text=True, check=True)
        out = completed.stdout
    assert out.startswith("REFUSED:"), out
    assert "both reserve" in out and "J8" in out, out
    assert "summary_statistics" in out, (
        "the refusal must name the block that already owns the cell, so the "
        "diagnosis is the answer rather than the start of one: " + out)


# ===========================================================================
# 4. THE FAILURE SAFETY THE WINDOWS RUN PROVED
# ===========================================================================
# The endpoint failed with a RAISED run-time error rather than a refusal, and
# the workbook came back untouched: status still CURRENT, bank A run_id 1,
# result_digest 246A7BB32851E1B1, consumed_auto_nonce 0, next_auto_nonce 1, F21
# blank, last_run_id 1, no sensitivity record published and the Sensitivity
# sheet empty. Those properties were correct before the correction and must stay
# correct after it, so they are controls now rather than an observation.
def _module_statements() -> list[str]:
    return _statements(SRC_VBA / "modSimPostReport.bas")


def _procedure_text(name: str) -> str:
    source = strip_comments(_text(SRC_VBA / "modSimPostReport.bas"))
    pattern = re.compile(
        rf"^(?:Public |Private )?(?:Sub|Function) {name}\b.*?^End (?:Sub|Function)",
        re.S | re.M)
    match = pattern.search(source)
    assert match, f"{name} is not declared"
    return match.group(0)


def test_11_a_raised_runtime_error_writes_nothing_either() -> None:
    """A refusal is not the only way out, and the other way was the one taken.

    `test_16` in the P7-4 suite proves the REFUSAL paths write nothing. The
    Windows failure took neither refusal nor success: an unhandled error reached
    `InvocationFailed`. Its handlers restore application state and announce, and
    that must be all they do.
    """
    endpoint = _procedure_text("PCCM_RunSensitivity")
    tail = endpoint[endpoint.index("InvocationFailed:"):]
    for banned in ("Range", "Value2", "ClearContents", "StampCell", "ClearRecords"):
        assert banned not in tail, (
            f"the raised-error path touches {banned!r}; a failure must leave the "
            "workbook exactly as it found it")
    assert "modAppState.FinishOperation" in tail, (
        "the raised-error path must still restore the application state")


def test_12_no_cell_is_written_before_publication_is_entered() -> None:
    """The whole result is built in memory, so a failure anywhere is harmless.

    Checked by locating every write in the module and requiring each one to be
    inside the three procedures publication owns - not by trusting that the
    earlier steps happen to be read-only today.
    """
    writers = {"Publish", "ClearRecords", "StampCell", "FillRecord"}
    source = strip_comments(_text(SRC_VBA / "modSimPostReport.bas"))
    procedures = re.findall(
        r"^(?:Public |Private )?(?:Sub|Function) (\w+)\b.*?^End (?:Sub|Function)",
        source, re.S | re.M)
    for name in procedures:
        body = _procedure_text(name)
        for statement in [s.strip() for s in body.split("\n")]:
            if re.search(r"\.(Value2|Formula|FormulaR1C1)\s*=", statement) or \
                    ".ClearContents" in statement:
                assert name in writers, (
                    f"{name} writes to the workbook: {statement!r}. Only "
                    f"{sorted(writers)} may, and only after every step has "
                    "succeeded")


def test_13_the_endpoint_moves_no_simulation_state() -> None:
    """run_id, the AUTO nonce, F21 and the digest are read and never written."""
    source = strip_comments(_text(SRC_VBA / "modSimPostReport.bas"))
    for constant in ("SIM_IDENTITY_ROW_RUN_ID", "SIM_IDENTITY_ROW_EFFECTIVE_SEED",
                     "SIM_IDENTITY_ROW_ITERATIONS_RUN", "SIM_IDENTITY_ROW_ACTIVE_BANK"):
        for statement in source.split("\n"):
            if constant in statement:
                assert "=" not in statement.split(constant)[0].split("(")[0] or \
                    "ReadSnapshotLong" in statement or "SharedText" in statement, statement
                assert ".Value2 =" not in statement, (
                    f"{constant} appears on the left of a write: {statement!r}")
    for banned in ("SIM_IDENTITY_ROW_NEXT_AUTO_NONCE", "SIM_IDENTITY_ROW_CONSUMED_AUTO_NONCE",
                   "SIM_IDENTITY_ROW_RESULT_DIGEST", "SIM_PENDING_AUTO_NONCE_CELL",
                   "SIM_IDENTITY_ROW_LAST_RUN_ID"):
        assert banned not in source, (
            f"the sensitivity endpoint names {banned}; it has no business "
            "reaching the nonce, the digest cell or the run-id counter at all")


# ===========================================================================
# 5. THE ROOT CAUSE, EXECUTED UNDER VBA'S EVALUATION RULE
# ===========================================================================
class _Series(list):
    """A VBA array `0 To count - 1`, which records every read it is given.

    Python lists accept negative indices and raise only past the end; a VBA
    array does neither, so the bound is checked here rather than inherited.
    """

    def __init__(self, values) -> None:
        super().__init__(values)
        self.out_of_range: list[int] = []

    def read(self, index: int) -> float:
        if index < 0 or index >= len(self):
            self.out_of_range.append(index)
            return 0.0
        return self[index]


def _merge_pass(values, decide, eager: bool):
    """One bottom-up merge sort, with the draw decision supplied by the caller.

    `decide` receives the loop state and the array, and answers True to draw
    from the left run. `eager` says whether the operands it does not need are
    evaluated anyway - which is what VBA does and Python does not.
    """
    count = len(values)
    series = _Series(values)
    scratch = [0.0] * count
    run_length = 1
    while run_length < count:
        low_end = 0
        while low_end < count:
            mid_point = min(low_end + run_length, count)
            high_end = min(mid_point + run_length, count)
            from_low, from_high = low_end, mid_point
            for target in range(low_end, high_end):
                if decide(series, from_low, from_high, mid_point, high_end, eager):
                    scratch[target] = series.read(from_low)
                    from_low += 1
                else:
                    scratch[target] = series.read(from_high)
                    from_high += 1
            low_end += 2 * run_length
        for target in range(count):
            series[target] = scratch[target]
        run_length *= 2
    return list(series), series.out_of_range


def _decide_original(series, from_low, from_high, mid_point, high_end, eager):
    """The expression as P7-4 shipped it, quoted for the record:

        If fromLow < midPoint And (fromHigh >= highEnd Or _
           Not (series(fromHigh) < series(fromLow))) Then
    """
    left = from_low < mid_point
    if eager:
        # VBA: every operand is evaluated, so the array reads happen whatever
        # the guards answered.
        right = (from_high >= high_end) or not (series.read(from_high) < series.read(from_low))
        if from_high >= high_end:
            series.read(from_high)
            series.read(from_low)
        return left and right
    return left and ((from_high >= high_end)
                     or not (series.read(from_high) < series.read(from_low)))


def _decide_current(series, from_low, from_high, mid_point, high_end, eager):
    """The branch structure the source now carries, pinned by `test_02`."""
    if from_low >= mid_point:
        return False
    if from_high >= high_end:
        return True
    return not (series.read(from_high) < series.read(from_low))


@pytest.mark.parametrize("values", [
    [2.0, 1.0],
    [1.0, 2.0, 3.0, 5.0, 4.0],
    [float((index * 7919) % 4001) for index in range(1000)] + [9.0, 8.0],
])
def test_14_the_previous_expression_reads_out_of_range_under_vbas_rule(values) -> None:
    """The defect, reproduced: eager evaluation, and the read lands past the end.

    Under Python's rule the SAME expression never touches the offending index,
    which is exactly why 3,876 transcribed tests were green.
    """
    _, lazy = _merge_pass(list(values), _decide_original, eager=False)
    assert lazy == [], (
        "under short-circuit evaluation the old expression is harmless - that is "
        "the whole reason it survived review")
    _, greedy = _merge_pass(list(values), _decide_original, eager=True)
    assert greedy, (
        "the old expression must be shown to read out of range once both operands "
        "are always evaluated; that read is VBA run-time error 9")
    assert max(greedy) >= len(values), (
        f"the offending read must be at or past the array's upper bound "
        f"{len(values) - 1}; got {sorted(set(greedy))}")


@pytest.mark.parametrize("values", [
    [2.0, 1.0],
    [1.0],
    [1.0, 2.0, 3.0, 5.0, 4.0],
    [3.0, 3.0, 3.0, 3.0],
    [float((index * 7919) % 4001) for index in range(1000)] + [9.0, 8.0],
])
def test_15_the_corrected_branches_are_safe_under_both_rules_and_still_sort(values) -> None:
    for eager in (False, True):
        ordered, offences = _merge_pass(list(values), _decide_current, eager=eager)
        assert offences == [], (
            f"the corrected merge read out of range at {offences} with eager={eager}")
        assert ordered == sorted(values), (ordered, sorted(values))


# ===========================================================================
# 6. ONE SET OF COORDINATES, AND EVERY OWNER USES IT
# ===========================================================================
def test_16_the_orchestrator_spells_no_coordinate_of_its_own() -> None:
    """`modSimPostReport` may name a column only through the projection.

    A single letter typed into this module would be a second declaration of the
    layout, and it would go on agreeing with the contract right up to the moment
    the contract moved - which is what just happened.
    """
    literals: set[str] = set()
    for statement in _module_statements():
        literals.update(re.findall(r'"([A-Za-z]{1,3})"', statement))
    assert not literals, (
        f"these look like column letters typed into the module: {sorted(literals)}. "
        "The coordinates belong to SIM_SENSITIVITY_*, SIM_ITER_*, "
        "SIM_SNAPSHOT_COLUMN_* and SIM_SHARED_VALUE_COLUMN.")
    for required in ("SIM_SENSITIVITY_A_FIRST_COLUMN", "SIM_SENSITIVITY_B_FIRST_COLUMN",
                     "SIM_SENSITIVITY_STAMP_COLUMN_A", "SIM_SENSITIVITY_STAMP_COLUMN_B"):
        assert any(required in statement for statement in _module_statements()), (
            f"{required} is the projected coordinate and must be what the module reads")


def test_17_every_owner_of_the_layout_agrees_with_the_contract() -> None:
    """The contract, the loader's pins, and the emitted projection, compared.

    Not "all three are CC" - all three are read, and required to be EQUAL. A
    test that named the letters would need editing on the next legitimate move
    and would be exactly as blind as the one that missed J.
    """
    from pccm_builder import sim_loader

    records = _raw_contract()["sim_data"]["sensitivity_records"]
    banks = records["banks"]
    stamp = records["stamp"]["bank_value_columns"]

    # the loader's locked pins
    assert sim_loader.LOCKED_SENSITIVITY_STAMP_COLUMNS == stamp, (
        sim_loader.LOCKED_SENSITIVITY_STAMP_COLUMNS, stamp)
    pinned = tuple((c["key"], c["column"], c["header"], c["value_type"])
                   for c in records["columns"])
    assert sim_loader.LOCKED_SENSITIVITY_COLUMN_LAYOUT == pinned

    # the emitted projection, built from the contract by the real emitter
    generated = PCCM_ROOT / "build" / "vba" / "modSimContract.bas"
    if not generated.is_file():
        pytest.skip("Stage A has not been built in this tree")
    projected = dict(re.findall(
        r'Public Const (SIM_SENSITIVITY_\w*COLUMN\w*) As String = "([A-Z]+)"',
        generated.read_text(encoding="utf-8")))
    assert projected.get("SIM_SENSITIVITY_A_FIRST_COLUMN") == banks["A"]["first_column"]
    assert projected.get("SIM_SENSITIVITY_A_LAST_COLUMN") == banks["A"]["last_column"]
    assert projected.get("SIM_SENSITIVITY_B_FIRST_COLUMN") == banks["B"]["first_column"]
    assert projected.get("SIM_SENSITIVITY_B_LAST_COLUMN") == banks["B"]["last_column"]
    assert projected.get("SIM_SENSITIVITY_STAMP_COLUMN_A") == stamp["A"]
    assert projected.get("SIM_SENSITIVITY_STAMP_COLUMN_B") == stamp["B"]

    # and the timing harness, which is the fourth reader of the same layout
    harness = (PCCM_ROOT / "bootstrap" / "windows"
               / "phase7_timing_scenarios.ps1").read_text(encoding="utf-8")
    for bank in ("A", "B"):
        assert f"'{bank}' = '{stamp[bank]}'" in harness, (
            f"the timing harness must read the bank {bank} stamp at {stamp[bank]}")
        last = banks[bank]["last_column"]
        assert f"'{bank}' = '{last}'" in harness, (
            f"the timing harness must read the bank {bank} status column at {last}")
