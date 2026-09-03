#!/usr/bin/env python3
"""PCCM Phase-6 test-only VBA SOURCE TRANSCRIBER.

--------------------------------------------------------------------------------
WHAT THIS IS, AND WHAT IT IS NOT
--------------------------------------------------------------------------------
It parses the statements of a `.bas` module - signatures, `Dim`/`ReDim`,
`If`/`ElseIf`/`Else`, `Do`/`Do While`/`Loop`, `For`/`Next`, `Exit Function`,
`Exit Do` and assignments - and compiles them into Python, modelling the VBA
semantics that differ from Python's:

    a scalar passed ByRef       a one-slot box, so an out-parameter reaches the caller
    a scalar passed ByVal       re-boxed at entry, so the caller cannot be written
    a UDT assignment            a deep copy into the caller's storage
    ReDim of a ByRef array      resized in place, not rebound
    Fix                         truncation toward zero
    StrComp(a, b, vbBinaryCompare)  ordinal comparison of UTF-16 code units
    LBound(x) + i               the caller's own lower bound

It is a TRANSCRIPTION OF SOURCE, not an interpreter of VBA. It proves that the
algorithm a module WRITES DOWN reproduces the accepted vectors, and it fails the
moment a locked expression changes, because every expression it evaluates is
read out of the file at test time.

It proves NOTHING about how VBA itself would execute those statements. Type
coercion, the numeric-literal parser, `Fix` on the VBA side, `ByRef` binding and
overflow behaviour are the Windows runtime's business, and Gate B is where they
are settled against the same accepted vectors.

--------------------------------------------------------------------------------
AND `And`/`Or` ARE THE ONE PLACE THAT DIFFERENCE HAS ALREADY COST A RUN
--------------------------------------------------------------------------------
`_expr` rewrites VBA `And` to Python `and` and VBA `Or` to Python `or`. Those
are NOT the same operator: PYTHON SHORT-CIRCUITS AND VBA DOES NOT. VBA evaluates
both operands of `And` and `Or` whatever the first one answers.

So an expression whose second operand is only safe because the first one
excluded it - the ordinary `If index <= limit And thing(index) ...` shape - runs
here and raises on Windows. It did: the P7-4 Windows run of 0734a38 failed with
"Subscript out of range" inside `SimSensitivitySortAscending`, whose merge read
`series(fromHigh)` after `fromHigh >= highEnd` had already said the right run
was spent. Every ported vector for that sort was green, because here the read
never happened.

A GREEN PORTED VECTOR IS THEREFORE NO EVIDENCE ABOUT THAT CLASS. What supplies
the evidence instead is `tests/test_phase7_runtime_defect_controls.py::test_01`,
a source-level rule over every `.bas`: no operand of an `And`/`Or` may subscript
an array with an index another operand of the same expression is comparing. This
rewrite is left as it is, deliberately - making it eager would change how every
transcribed module evaluates - and the rule above is what stands in for it.

--------------------------------------------------------------------------------
THIS IS TEST INFRASTRUCTURE AND CARRIES NO AUTHORITY
--------------------------------------------------------------------------------
It lives under `tests/` and NOTHING in `builder/`, `src/`, `spec/` or
`bootstrap/` may import it. The authorities remain the contracts, the accepted
Python reference modules and the accepted evidence vectors; this file only reads
`.bas` text and evaluates it.

Extracted mechanically from the accepted Step-6 suite
(`tests/test_phase6_sim_rng_vba.py`, commit 2ec1844) so Step 7 could reuse it
rather than grow a second transcription language. The Step-6 semantics are
unchanged; the additions are `Do ... Loop`, the `Log`/`Exp`/`Sqr`/`Abs`
builtins, and the ability to compile more than one module into one namespace.
"""

from __future__ import annotations

import math
import re
from pathlib import Path

from pccm_builder.vba_source import logical_statements, strip_comments

class _Ref:
    """A VBA scalar, boxed, so a ByRef parameter behaves as VBA passes it."""

    __slots__ = ("v",)

    def __init__(self, v: object = 0.0) -> None:
        self.v = v

    def __repr__(self) -> str:  # pragma: no cover - diagnostics only
        return f"_Ref({self.v!r})"


def _val(x):
    return x.v if isinstance(x, _Ref) else x


def _fix(x):
    return float(math.trunc(x))


def _copy(x):
    if isinstance(x, dict):
        return {k: _copy(v) for k, v in x.items()}
    if isinstance(x, list):
        return [_copy(v) for v in x]
    return x


def _assign(target, value):
    value = _copy(value)
    if isinstance(target, dict):
        target.clear()
        target.update(value)
    else:
        target[:] = value


def _cstr(x):
    return str(int(x)) if float(x).is_integer() else repr(x)


# ---------------------------------------------------------------------------
# VBA strings ARE UTF-16, and the hash is defined over UTF-16 CODE UNITS. Python
# strings are sequences of CODE POINTS, so `len`, slicing and `ord` all disagree
# with VBA on any astral character - exactly the inputs the surrogate-pair
# vectors exist to cover. These three model the VBA semantics directly.
# ---------------------------------------------------------------------------
def _units(text: str) -> bytes:
    return text.encode("utf-16-le", "surrogatepass")


def _decode(raw: bytes) -> str:
    return raw.decode("utf-16-le", "surrogatepass")


def _len(value):
    """VBA `Len`: UTF-16 code units for a String, element count for an array."""
    return len(_units(value)) // 2 if isinstance(value, str) else len(value)


def _mid(text, start, length=None):
    """VBA `Mid`: 1-BASED, in code units, and taking a LENGTH not an end."""
    raw = _units(text)
    lo = (int(start) - 1) * 2
    hi = len(raw) if length is None else lo + int(length) * 2
    return _decode(raw[lo:hi])


def _ascw(text):
    """VBA `AscW`: the first UTF-16 code unit, as a SIGNED 16-bit Integer.

    Everything above U+7FFF comes back negative, which is the whole reason
    `CalcFpNormaliseCodeUnit` exists.
    """
    unit = int.from_bytes(_units(text)[:2], "little")
    return unit - 65536 if unit > 32767 else unit


def _vbfor(low, high, step):
    """VBA's counted loop with an explicit Step, including a negative one."""
    step = int(step)
    assert step != 0, "a VBA For loop may not Step by zero"
    value, limit = int(low), int(high)
    while value <= limit if step > 0 else value >= limit:
        yield value
        value += step


def _strcomp(a, b, mode):
    assert mode == 0, "only vbBinaryCompare is transcribed"
    ka, kb = a.encode("utf-16-be"), b.encode("utf-16-be")
    return (ka > kb) - (ka < kb)


_BUILTIN = {
    "Fix": "_fix", "CDbl": "float", "CLng": "int", "CStr": "_cstr",
    "Len": "_len", "LBound": "_lbound", "UBound": "_ubound", "StrComp": "_strcomp",
    # VBA `Log` is the NATURAL logarithm, not base 10. These map to DOTLESS
    # names on purpose: the member-access rewrite below would read `math.log`
    # as a UDT field access and turn it into `math["log"]`.
    "Log": "_log", "Exp": "_exp", "Sqr": "_sqrt", "Abs": "abs",
    "Asc": "_asc", "Mid": "_mid", "AscW": "_ascw",
}
_SIG = re.compile(
    r"^(Public|Private)\s+(Function|Sub)\s+(\w+)\((.*)\)(?:\s+As\s+(\w+))?$"
)
_PARAM = re.compile(r"^(ByRef|ByVal)?\s*(\w+)(\(\))?\s+As\s+(\w+)$")
_DEFAULT = {"Boolean": "False", "Double": "0e0", "Long": "0", "String": '""'}


def _match_paren(text: str, opening: int) -> int:
    depth = 0
    for i in range(opening, len(text)):
        if text[i] == "(":
            depth += 1
        elif text[i] == ")":
            depth -= 1
            if depth == 0:
                return i
    raise AssertionError(f"unbalanced parentheses: {text}")


def _split_commas(text: str) -> list[str]:
    out, depth, cur = [], 0, ""
    for ch in text:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    if cur.strip():
        out.append(cur)
    return [p.strip() for p in out]


def _protect(text: str) -> tuple[str, list[str]]:
    lits: list[str] = []

    def take(match: re.Match) -> str:
        lits.append(match.group(0))
        return f"\x00{len(lits) - 1}\x00"

    return re.sub(r'"(?:[^"]|"")*"', take, text), lits


def _bracket(text: str, names: list[str]) -> str:
    """`arr(i)` becomes `arr[i]` for every declared array."""
    if not names:
        return text
    pattern = re.compile(r"\b(" + "|".join(sorted(names, key=len, reverse=True)) + r")\(")
    while True:
        match = pattern.search(text)
        if not match:
            return text
        close = _match_paren(text, match.end() - 1)
        text = (text[: match.start()] + match.group(1) + "["
                + text[match.end(): close] + "]" + text[close + 1:])


def _mark_byref(text: str, procs: dict, env: dict) -> str:
    """Protect a bare scalar passed to a ByRef scalar parameter from unboxing."""
    for name, params in procs.items():
        positions = [i for i, p in enumerate(params) if p[0] == "ByRef" and p[4] == "scalar"]
        if not positions:
            continue
        pattern = re.compile(rf"\b{name}\(")
        start = 0
        while True:
            match = pattern.search(text, start)
            if not match:
                break
            close = _match_paren(text, match.end() - 1)
            args = _split_commas(text[match.end(): close])
            for i in positions:
                if i < len(args) and re.fullmatch(r"\w+", args[i]) \
                        and env.get(args[i], ("",))[0] == "scalar":
                    args[i] = "\x01" + args[i] + "\x01"
            text = text[: match.end()] + ", ".join(args) + text[close:]
            start = match.end()
    return text


def _expr(text: str, env: dict, procs: dict, cond: bool = False) -> str:
    text, lits = _protect(text)
    # A MODULE-QUALIFIED CALL is the same procedure: `modCalcFingerprint.CalcFpX(...)`
    # and `CalcFpX(...)` name one thing. The transcription puts every compiled
    # procedure in ONE namespace, so the qualifier is dropped mechanically here,
    # before anything else reads the name - in particular before the ByRef
    # marking, which looks the name up, and before the member-access rewrite,
    # which would otherwise read the module as a UDT.
    text = re.sub(r"\bmod[A-Za-z]\w*\.(?=\w+\()", "", text)
    text = _mark_byref(text, procs, env)
    text = text.replace("<>", "!=")
    if cond:
        text = re.sub(r"(?<![<>=!])=(?!=)", "==", text)
    text = re.sub(r"\bNot\b", " not ", text)
    text = re.sub(r"\bAnd\b", " and ", text)
    text = re.sub(r"\bOr\b", " or ", text)
    text = text.replace("&", "+")
    text = text.replace("vbNullString", '""').replace("vbBinaryCompare", "0")
    # `Mid$` and friends are the same builtins with VBA's String-returning
    # suffix. Dropping the `$` is mechanical; there is no second behaviour.
    text = re.sub(r"\b(Mid|Left|Right|Trim|UCase|LCase|Chr|Space)\$\(", r"\1(", text)
    for vba, py in _BUILTIN.items():
        text = re.sub(rf"\b{vba}\(", f"{py}(", text)
    text = _bracket(text, [n for n, k in env.items() if k[0] == "array"])
    text = re.sub(r"([\w\]])\.(\w+)", r'\1["\2"]', text)
    scalars = [n for n, k in env.items() if k[0] == "scalar"]
    if scalars:
        joined = "|".join(sorted(scalars, key=len, reverse=True))
        text = re.sub(rf"(?<!\x01)\b({joined})\b(?!\s*[\[(])(?!\x01)", r"\1.v", text)
    text = re.sub(r"(\d)#", r"\1e0", text)
    text = text.replace("\x01", "")
    return re.sub(r"\x00(\d+)\x00", lambda m: lits[int(m.group(1))], text)


def _split_assign(text: str) -> tuple[str, str] | None:
    depth, in_string = 0, False
    for i, ch in enumerate(text):
        if ch == '"':
            in_string = not in_string
        if in_string:
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "=" and depth == 0 and text[i - 1] not in "<>!" \
                and (i + 1 >= len(text) or text[i + 1] != "="):
            return text[:i].strip(), text[i + 1:].strip()
    return None


def _parse_types(code: str) -> dict[str, list[tuple[str, str]]]:
    types: dict[str, list[tuple[str, str]]] = {}
    # Private Type as well as Public: a module may keep its prepared
    # representation private precisely so no caller can write it.
    for block in re.finditer(r"^(?:Public|Private) Type (\w+)\n(.*?)^End Type",
                             code, re.M | re.S):
        fields = []
        for line in block.group(2).splitlines():
            field = re.match(r"\s*(\w+)\s+As\s+(\w+)\s*$", line)
            if field:
                fields.append((field.group(1), field.group(2)))
        types[block.group(1)] = fields
    return types


def _proto(kind: str, types: dict):
    if kind in types:
        return {f: _proto(t, types) for f, t in types[kind]}
    if kind == "Boolean":
        return False
    return {"String": "", "Long": 0}.get(kind, 0.0)


def _compile_proc(sig: re.Match, body: list[str], types: dict, procs: dict) -> str:
    name, args, ret = sig.group(3), sig.group(4), sig.group(5)
    env: dict[str, tuple[str, str]] = {}
    names = []
    for raw in _split_commas(args):
        param = _PARAM.match(raw)
        assert param, f"{name}: unreadable parameter {raw!r}"
        mode, pname, arr, ptype = (param.group(1) or "ByRef", param.group(2),
                                   bool(param.group(3)), param.group(4))
        env[pname] = ("array" if arr else ("udt" if ptype in types else "scalar"), ptype)
        names.append((mode, pname, arr, ptype))
    out = [f"def {name}({', '.join(n[1] for n in names)}):"]
    for mode, pname, arr, ptype in names:
        if mode == "ByVal" and env[pname][0] == "scalar":
            # VBA ByVal is a private copy: assigning it must not reach the caller.
            out.append(f"    {pname} = _Ref(_val({pname}))")
    out.append(f"    _r = {_DEFAULT.get(ret, 'None')}")
    indent = 1
    for text in body:
        indent = _emit(text, out, indent, env, types, procs, name)
    assert indent == 1, f"{name}: unbalanced block structure"
    out.append("    return _r")
    return "\n".join(out)


def _declare(decl: str, env: dict, types: dict, procs: dict, pad: str,
             redim: bool) -> str:
    shaped = re.match(r"^(\w+)(?:\((.*)\))?\s+As\s+(\w+)$", decl)
    if shaped is None:  # a ReDim, which restates the bounds and not the type
        again = re.match(r"^(\w+)\((.*)\)$", decl)
        assert again, f"unreadable declaration {decl!r}"
        name, bounds = again.group(1), again.group(2)
        assert env[name][0] == "array", f"ReDim of a non-array {name}"
        # ReDim RESIZES IN PLACE. A ByRef array parameter is resized in the
        # caller, so rebinding the name here would lose the result.
        return pad + (f"_assign({name}, _arr({_bounds(bounds, env, procs)}, "
                      f"{env[name][1]!r}))")
    name, bounds, kind = shaped.group(1), shaped.group(2), shaped.group(3)
    if bounds is not None:
        env[name] = ("array", kind)
        if not bounds.strip():
            return pad + f"{name} = []"
        return pad + f"{name} = _arr({_bounds(bounds, env, procs)}, {kind!r})"
    if kind in types:
        env[name] = ("udt", kind)
        return pad + f"{name} = _new({kind!r})"
    env[name] = ("scalar", kind)
    return pad + f"{name} = _Ref({_DEFAULT[kind]})"


def _bounds(bounds: str, env: dict, procs: dict) -> str:
    low, high = bounds.split(" To ")
    return f"{_expr(low, env, procs)}, {_expr(high, env, procs)}"


def _emit(text: str, out: list[str], indent: int, env: dict, types: dict,
          procs: dict, procname: str) -> int:
    pad = "    " * indent
    if re.match(r"^(Dim|ReDim)\s", text):
        for decl in _split_commas(text.split(" ", 1)[1]):
            out.append(_declare(decl, env, types, procs, pad,
                                text.startswith("ReDim ")))
        return indent
    inline = re.match(r"^If (.*) Then (.+)$", text)
    block = re.match(r"^If (.*) Then$", text)
    if block:
        out.append(pad + f"if {_expr(block.group(1), env, procs, cond=True)}:")
        return indent + 1
    if inline:
        out.append(pad + f"if {_expr(inline.group(1), env, procs, cond=True)}:")
        out.append(pad + "    " + _simple(inline.group(2), env, procs, procname))
        return indent
    elif_ = re.match(r"^ElseIf (.*) Then$", text)
    if elif_:
        out.append("    " * (indent - 1)
                   + f"elif {_expr(elif_.group(1), env, procs, cond=True)}:")
        return indent
    if text == "Else":
        out.append("    " * (indent - 1) + "else:")
        return indent
    if text == "End If":
        return indent - 1
    loop = re.match(r"^Do While (.*)$", text)
    if loop:
        out.append(pad + f"while {_expr(loop.group(1), env, procs, cond=True)}:")
        return indent + 1
    if text == "Do":
        # An unconditional Do ... Loop leaves only through Exit Do or Exit
        # Function, exactly as the rejection samplers are written.
        out.append(pad + "while True:")
        return indent + 1
    if text == "Loop":
        return indent - 1
    # `For i = a To b Step s`. Only the explicit-Step form goes through the
    # generator; the plain form keeps the exact `range` it always had.
    stepped = re.match(r"^For (\w+) = (.*?) To (.*?) Step (.*)$", text)
    if stepped:
        var, low, high, step = stepped.groups()
        out.append(pad + f"for _t{indent} in _vbfor("
                         f"{_expr(low, env, procs)}, {_expr(high, env, procs)}, "
                         f"{_expr(step, env, procs)}):")
        out.append(pad + f"    {var}.v = _t{indent}")
        return indent + 1
    counted = re.match(r"^For (\w+) = (.*) To (.*)$", text)
    if counted:
        var, low, high = counted.groups()
        out.append(pad + f"for _t{indent} in range(int({_expr(low, env, procs)}), "
                         f"int({_expr(high, env, procs)}) + 1):")
        out.append(pad + f"    {var}.v = _t{indent}")
        return indent + 1
    if re.match(r"^Next\b", text):
        return indent - 1
    out.append(pad + _simple(text, env, procs, procname))
    return indent


def _simple(text: str, env: dict, procs: dict, procname: str) -> str:
    if text in ("Exit Function", "Exit Sub"):
        return "return _r"
    if text == "Exit Do":
        return "break"
    parts = _split_assign(text)
    assert parts, f"unreadable statement {text!r}"
    lhs, rhs = parts
    value = _expr(rhs, env, procs)
    if lhs == procname:
        return f"_r = {value}"
    kind = env.get(lhs, (None,))[0]
    if kind == "scalar":
        return f"{lhs}.v = {value}"
    if kind in ("array", "udt"):
        return f"_assign({lhs}, {value})"
    return f"{_expr(lhs, env, procs)} = _copy({value})"


def build(sources, constants, only=None, extra=None, signature_only=None) -> dict:
    """Compile `.bas` modules into one namespace.

    `sources`     ordered {module name: Path}. Later modules may call earlier ones
                  and vice versa: resolution happens at call time.
    `constants`   {name: value} the modules read - the generated projection.
    `only`        optional {module name: {procedure names}} to compile a SUBSET of
                  a module. Used for borrowing an accepted Phase-5 predicate
                  without dragging in a module that does not belong to the step
                  under test.
    `extra`       optional additional namespace entries.
    `signature_only`  optional {module name: {procedure names}} whose SIGNATURES
                  are read from the source and registered, but whose bodies are
                  NOT compiled - the caller binds them through `extra`. Used for
                  accepted primitives whose VBA bodies are outside this engine's
                  reach; reading the real declaration is what keeps the ByRef /
                  ByVal call convention honest instead of retyped by hand.

    Returns the namespace, which also carries `_python_source`, `_types` and
    `_procs` for the tests that assert over them.
    """
    types: dict[str, list[tuple[str, str]]] = {}
    per_module: dict[str, list[tuple[re.Match, list[str]]]] = {}
    for name, path in sources.items():
        code = strip_comments(Path(path).read_text(encoding="utf-8"))
        types.update(_parse_types(code))
        per_module[name] = _split_procedures(code, (only or {}).get(name))

    procs: dict[str, list[tuple]] = {}
    for name, entries in per_module.items():
        for sig, _ in entries:
            procs[sig.group(3)] = _signature_params(sig, types)
    for name, wanted in (signature_only or {}).items():
        code = strip_comments(Path(sources[name]).read_text(encoding="utf-8"))
        found = set()
        for sig, _ in _split_procedures(code, wanted):
            procs[sig.group(3)] = _signature_params(sig, types)
            found.add(sig.group(3))
        assert found == set(wanted), f"missing signatures: {sorted(set(wanted) - found)}"

    source: list[str] = []
    for name, entries in per_module.items():
        for sig, body in entries:
            source.append(_compile_proc(sig, body, types, procs))

    namespace: dict = {
        "_Ref": _Ref, "_val": _val, "_fix": _fix, "_copy": _copy, "_assign": _assign,
        "_cstr": _cstr, "_strcomp": _strcomp,
        "_log": math.log, "_exp": math.exp, "_sqrt": math.sqrt,
        "_asc": lambda text: ord(text[0]),
        "_len": _len, "_mid": _mid, "_ascw": _ascw, "_vbfor": _vbfor,
        "_lbound": lambda seq: 0, "_ubound": lambda seq: len(seq) - 1,
        "_new": lambda kind: _proto(kind, types),
        "_arr": lambda low, high, kind: [
            _proto(kind, types) for _ in range(max(int(high) - int(low) + 1, 0))
        ],
    }
    namespace.update(constants)
    namespace.update(extra or {})
    text = "\n\n".join(source)
    namespace["_python_source"] = text
    namespace["_types"] = types
    namespace["_procs"] = procs
    exec(compile(text, "<vba source transcription>", "exec"), namespace)
    return namespace


def _signature_params(sig: re.Match, types: dict) -> list[tuple]:
    params = []
    for raw in _split_commas(sig.group(4)):
        param = _PARAM.match(raw)
        assert param, raw
        arr = bool(param.group(3))
        ptype = param.group(4)
        params.append((param.group(1) or "ByRef", param.group(2), arr, ptype,
                       "array" if arr else ("udt" if ptype in types else "scalar")))
    return params


def _split_procedures(code: str, only=None) -> list[tuple[re.Match, list[str]]]:
    statements = logical_statements(code)
    out, index = [], 0
    while index < len(statements):
        sig = _SIG.match(statements[index][1])
        if not sig:
            index += 1
            continue
        body, index = [], index + 1
        while not re.match(r"^End (Function|Sub)$", statements[index][1]):
            body.append(statements[index][1])
            index += 1
        index += 1
        if only is None or sig.group(3) in only:
            out.append((sig, body))
    if only is not None:
        found = {sig.group(3) for sig, _ in out}
        assert found == set(only), f"missing procedures: {sorted(set(only) - found)}"
    return out
