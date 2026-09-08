#!/usr/bin/env python3
"""A small Excel expression evaluator, for checking a sheet's OWN formulas.

WHY THIS EXISTS. A control that re-implemented the Model Check semantics in
Python would prove that two implementations agree, which is not the claim worth
making: the sheet is the implementation, and a defect in its formulas is exactly
what a control has to catch. So the controls evaluate the FORMULAS THE BUILD
WROTE, cell by cell, against readings a scenario supplies. Mutate a formula and
the answer moves.

WHAT IT IS NOT. It is not an Excel. It supports the operators and the handful of
functions the Model Check surface uses, and it REFUSES anything else by name
rather than guessing - a silent fallback would let a formula the controls cannot
really evaluate look as though it had been checked.

WHAT IT DELIBERATELY COPIES FROM EXCEL, because these are the semantics the
sheet's correctness turns on:

  * an EMPTY cell reads back as 0, which is why the Model Check candidate block
    is built never to leave one;
  * `=` on two strings is CASE-INSENSITIVE, and EXACT is not;
  * an error propagates through every operator and every function except
    IFERROR, so a broken owner reaches the register as an error rather than as a
    quietly wrong word.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Callable

__all__ = ["Evaluator", "ExcelError", "NA", "VALUE"]


@dataclass(frozen=True)
class ExcelError:
    name: str

    def __str__(self) -> str:  # pragma: no cover - diagnostics only
        return self.name


NA = ExcelError("#N/A")
VALUE = ExcelError("#VALUE!")
DIV0 = ExcelError("#DIV/0!")


_MISSING = object()


class FormulaError(Exception):
    """The evaluator itself cannot proceed - an unsupported construct, a cycle,
    a reference nobody supplied. Never an Excel error value: those are results."""


_TOKEN = re.compile(
    r"""
    (?P<ws>\s+)
  | (?P<string>"(?:[^"]|"")*")
  | (?P<sheetref>(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ .]*)![$]?[A-Za-z]{1,3}[$]?\d+)
  | (?P<ref>[$]?[A-Za-z]{1,3}[$]?\d+)
  | (?P<number>\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?)
  | (?P<func>[A-Za-z_][A-Za-z0-9_.]*\s*\()
  | (?P<name>[A-Za-z_][A-Za-z0-9_.]*)
  | (?P<op><=|>=|<>|[-+*/&^=<>(),:%])
    """,
    re.VERBOSE,
)

_COMPARISONS = {"=", "<>", "<", ">", "<=", ">="}


def _tokenize(text: str) -> list[tuple[str, str]]:
    tokens: list[tuple[str, str]] = []
    position = 0
    while position < len(text):
        match = _TOKEN.match(text, position)
        if match is None:
            raise FormulaError(f"cannot tokenize {text[position:position + 20]!r} in {text!r}")
        position = match.end()
        kind = match.lastgroup
        assert kind is not None
        if kind == "ws":
            continue
        tokens.append((kind, match.group()))
    return tokens


class _Parser:
    """Precedence, lowest first: comparison, &, +-, */, unary -, %, ^."""

    def __init__(self, tokens: list[tuple[str, str]]) -> None:
        self.tokens = tokens
        self.at = 0

    def peek(self) -> tuple[str, str] | None:
        return self.tokens[self.at] if self.at < len(self.tokens) else None

    def take(self) -> tuple[str, str]:
        token = self.peek()
        if token is None:
            raise FormulaError("formula ended unexpectedly")
        self.at += 1
        return token

    def expect(self, value: str) -> None:
        kind, text = self.take()
        if text != value:
            raise FormulaError(f"expected {value!r}, found {text!r}")

    def parse(self) -> Any:
        node = self.comparison()
        if self.peek() is not None:
            raise FormulaError(f"trailing tokens at {self.peek()!r}")
        return node

    def comparison(self) -> Any:
        node = self.concat()
        while (token := self.peek()) and token[1] in _COMPARISONS:
            operator = self.take()[1]
            node = ("binary", operator, node, self.concat())
        return node

    def concat(self) -> Any:
        node = self.additive()
        while (token := self.peek()) and token[1] == "&":
            self.take()
            node = ("binary", "&", node, self.additive())
        return node

    def additive(self) -> Any:
        node = self.multiplicative()
        while (token := self.peek()) and token[1] in ("+", "-"):
            operator = self.take()[1]
            node = ("binary", operator, node, self.multiplicative())
        return node

    def multiplicative(self) -> Any:
        node = self.unary()
        while (token := self.peek()) and token[1] in ("*", "/"):
            operator = self.take()[1]
            node = ("binary", operator, node, self.unary())
        return node

    def unary(self) -> Any:
        token = self.peek()
        if token and token[1] in ("-", "+"):
            operator = self.take()[1]
            return ("unary", operator, self.unary())
        return self.postfix()

    def postfix(self) -> Any:
        node = self.primary()
        while (token := self.peek()) and token[1] == "%":
            self.take()
            node = ("percent", node)
        return node

    def primary(self) -> Any:
        kind, text = self.take()
        if kind == "number":
            return ("literal", float(text))
        if kind == "string":
            return ("literal", text[1:-1].replace('""', '"'))
        if kind == "func":
            name = text[:-1].strip().upper()
            arguments: list[Any] = []
            if (token := self.peek()) and token[1] == ")":
                self.take()
                return ("call", name, arguments)
            while True:
                arguments.append(self.comparison())
                kind2, text2 = self.take()
                if text2 == ")":
                    break
                if text2 != ",":
                    raise FormulaError(f"expected , or ) in {name}, found {text2!r}")
            return ("call", name, arguments)
        if kind in ("ref", "sheetref"):
            if (token := self.peek()) and token[1] == ":":
                self.take()
                kind2, text2 = self.take()
                if kind2 not in ("ref", "sheetref"):
                    raise FormulaError(f"range end {text2!r} is not a reference")
                return ("range", text, text2)
            return ("ref", text)
        if kind == "name":
            if text.upper() in ("TRUE", "FALSE"):
                return ("literal", text.upper() == "TRUE")
            return ("name", text)
        if text == "(":
            node = self.comparison()
            self.expect(")")
            return node
        raise FormulaError(f"unexpected token {text!r}")


def _parse(formula: str) -> Any:
    body = formula[1:] if formula.startswith("=") else formula
    return _Parser(_tokenize(body)).parse()


def _column_index(letters: str) -> int:
    value = 0
    for character in letters.upper():
        value = value * 26 + (ord(character) - 64)
    return value


def _split(reference: str) -> tuple[str | None, str, int]:
    sheet = None
    if "!" in reference:
        sheet, reference = reference.split("!", 1)
        sheet = sheet.strip("'")
    match = re.fullmatch(r"[$]?([A-Za-z]{1,3})[$]?(\d+)", reference)
    if match is None:
        raise FormulaError(f"{reference!r} is not a cell reference")
    return sheet, match.group(1).upper(), int(match.group(2))


class Evaluator:
    """Evaluate one sheet's formulas against supplied external values.

    `cells` maps `COLUMNROW` (no dollars) to a stored value - a formula string, a
    number, a string, or None for an empty cell. `externals` answers everything
    off the sheet: other sheets' cells by `Sheet!D10`, defined names by name, and
    zero-argument worksheet functions by `NAME()`.
    """

    def __init__(self, cells: dict[str, Any], externals: dict[str, Any],
                 now: float = 45000.0) -> None:
        self.cells = cells
        self.externals = externals
        self.now = now
        self._cache: dict[str, Any] = {}
        self._active: set[str] = set()

    # -- public -------------------------------------------------------------
    def value(self, address: str) -> Any:
        sheet, column, row = _split(address)
        if sheet is not None:
            return self._external(f"{sheet}!{column}{row}")
        key = f"{column}{row}"
        if key in self._cache:
            return self._cache[key]
        if key in self._active:
            raise FormulaError(f"circular reference at {key}")
        stored = self.cells.get(key)
        self._active.add(key)
        try:
            if stored is None:
                result: Any = 0.0        # AN EMPTY CELL IS A HARD ZERO IN EXCEL.
            elif isinstance(stored, str) and stored.startswith("="):
                result = self._evaluate(_parse(stored))
            else:
                result = stored
        finally:
            self._active.discard(key)
        self._cache[key] = result
        return result

    def evaluate(self, formula: str) -> Any:
        return self._evaluate(_parse(formula))

    def invalidate(self) -> None:
        self._cache.clear()

    # -- internals ----------------------------------------------------------
    def _udf(self, name: str) -> Any:
        for key, value in self.externals.items():
            if key.endswith("()") and key[:-2].upper() == name:
                return value
        return _MISSING

    def _external(self, key: str) -> Any:
        if key not in self.externals:
            raise FormulaError(f"no external value supplied for {key!r}")
        value = self.externals[key]
        return 0.0 if value is None else value

    def _evaluate(self, node: Any) -> Any:
        kind = node[0]
        if kind == "literal":
            return node[1]
        if kind == "ref":
            return self.value(node[1])
        if kind == "range":
            return self._range(node[1], node[2])
        if kind == "name":
            return self._external(node[1])
        if kind == "percent":
            inner = self._evaluate(node[1])
            return inner if isinstance(inner, ExcelError) else _number(inner) / 100.0
        if kind == "unary":
            return _apply_unary(node[1], self._evaluate(node[2]))
        if kind == "binary":
            return _apply_binary(node[1], self._evaluate(node[2]), self._evaluate(node[3]))
        if kind == "call":
            return self._call(node[1], node[2])
        raise FormulaError(f"unknown node {kind!r}")

    def _range(self, start: str, end: str) -> list[Any]:
        sheet_a, column_a, row_a = _split(start)
        sheet_b, column_b, row_b = _split(end)
        if sheet_a != sheet_b:
            raise FormulaError(f"range {start}:{end} spans two sheets")
        columns = range(_column_index(column_a), _column_index(column_b) + 1)
        values = []
        for row in range(row_a, row_b + 1):
            for column in columns:
                reference = f"{_letters(column)}{row}"
                values.append(self.value(reference if sheet_a is None
                                         else f"{sheet_a}!{reference}"))
        return values

    def _call(self, name: str, arguments: list[Any]) -> Any:
        if name == "IFERROR":
            first = self._evaluate(arguments[0])
            return self._evaluate(arguments[1]) if isinstance(first, ExcelError) else first
        if name == "IF":
            condition = self._evaluate(arguments[0])
            if isinstance(condition, ExcelError):
                return condition
            try:
                taken = _truthy(condition)
            except _Coerce as error:
                return error.error
            branch = arguments[1] if taken else (
                arguments[2] if len(arguments) > 2 else None)
            return False if branch is None else self._evaluate(branch)
        if name == "NA":
            return NA
        if name == "NOW":
            return self.now

        if not arguments:
            # A ZERO-ARGUMENT WORKSHEET FUNCTION. These are the accepted VBA
            # accessors, and a scenario supplies what each one answered. An
            # accessor nobody supplied is a FormulaError rather than a blank:
            # a control that silently read 0 for a missing owner would be
            # testing a workbook nobody built.
            supplied = self._udf(name)
            if supplied is not _MISSING:
                return supplied
        values = [self._evaluate(argument) for argument in arguments]
        if name not in _FUNCTIONS:
            raise FormulaError(f"unsupported function {name}")
        # AN ERROR PROPAGATES. Only IFERROR above may swallow one, and the Model
        # Check register deliberately does not use it: an owner that fails must
        # reach the sheet as an error, not as "no such check".
        flat: list[Any] = []
        for value in values:
            flat.extend(value if isinstance(value, list) else [value])
        first_error = next((v for v in flat if isinstance(v, ExcelError)), None)
        if first_error is not None and name not in ("ISNA", "ISERROR"):
            return first_error
        try:
            return _FUNCTIONS[name](values)
        except _Coerce as error:
            return error.error


def _letters(index: int) -> str:
    text = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        text = chr(65 + remainder) + text
    return text


def _number(value: Any) -> float:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        if value == "":
            return 0.0
        try:
            return float(value)
        except ValueError as error:
            raise _Coerce(VALUE) from error
    raise _Coerce(VALUE)


class _Coerce(Exception):
    def __init__(self, error: ExcelError) -> None:
        super().__init__(error.name)
        self.error = error


def _text(value: Any) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _apply_unary(operator: str, value: Any) -> Any:
    if isinstance(value, list):
        return [_apply_unary(operator, item) for item in value]
    if isinstance(value, ExcelError):
        return value
    try:
        number = _number(value)
    except _Coerce as error:
        return error.error
    return -number if operator == "-" else number


def _apply_binary(operator: str, left: Any, right: Any) -> Any:
    if isinstance(left, list) or isinstance(right, list):
        size = len(left) if isinstance(left, list) else len(right)
        return [_apply_binary(operator,
                              left[index] if isinstance(left, list) else left,
                              right[index] if isinstance(right, list) else right)
                for index in range(size)]
    if isinstance(left, ExcelError):
        return left
    if isinstance(right, ExcelError):
        return right
    if operator == "&":
        return _text(left) + _text(right)
    if operator in _COMPARISONS:
        return _compare(operator, left, right)
    try:
        a, b = _number(left), _number(right)
    except _Coerce as error:
        return error.error
    if operator == "+":
        return a + b
    if operator == "-":
        return a - b
    if operator == "*":
        return a * b
    if operator == "/":
        return DIV0 if b == 0 else a / b
    raise FormulaError(f"unsupported operator {operator!r}")


def _rank(value: Any) -> tuple[int, Any]:
    """Excel's ordering across types: numbers, then text, then booleans."""
    if isinstance(value, bool):
        return (2, value)
    if isinstance(value, (int, float)):
        return (0, float(value))
    # TEXT COMPARISON IS CASE-INSENSITIVE in Excel, which is why EXACT exists at
    # all and why the duplicate rule uses EXACT rather than `=`.
    return (1, str(value).upper())


def _compare(operator: str, left: Any, right: Any) -> bool:
    a, b = _rank(left), _rank(right)
    if operator == "=":
        return a == b
    if operator == "<>":
        return a != b
    if a[0] != b[0]:
        return {"<": a[0] < b[0], "<=": a[0] <= b[0],
                ">": a[0] > b[0], ">=": a[0] >= b[0]}[operator]
    return {"<": a[1] < b[1], "<=": a[1] <= b[1],
            ">": a[1] > b[1], ">=": a[1] >= b[1]}[operator]


def _truthy(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        raise _Coerce(VALUE)
    return bool(value)


def _flatten(values: list[Any]) -> list[Any]:
    out: list[Any] = []
    for value in values:
        out.extend(value if isinstance(value, list) else [value])
    return out


def _sumproduct(values: list[Any]) -> float:
    arrays = [value if isinstance(value, list) else [value] for value in values]
    size = max(len(array) for array in arrays)
    total = 0.0
    for index in range(size):
        product = 1.0
        for array in arrays:
            product *= _number(array[index] if len(array) > 1 else array[0])
        total += product
    return total


def _match(values: list[Any]) -> Any:
    needle, haystack, kind = values[0], values[1], int(_number(values[2]))
    if kind != 0:
        raise FormulaError("only exact MATCH is supported")
    for index, candidate in enumerate(haystack, start=1):
        if _compare("=", needle, candidate):
            return float(index)
    return NA


def _index(values: list[Any]) -> Any:
    array, position = values[0], int(_number(values[1]))
    if position < 1 or position > len(array):
        return ExcelError("#REF!")
    return array[position - 1]


def _find(values: list[Any]) -> Any:
    needle, haystack = _text(values[0]), _text(values[1])
    start = int(_number(values[2])) if len(values) > 2 else 1
    position = haystack.find(needle, start - 1)
    return ExcelError("#VALUE!") if position < 0 else float(position + 1)


def _exact(values: list[Any]) -> Any:
    left, right = values[0], values[1]
    if isinstance(left, list) or isinstance(right, list):
        size = len(left) if isinstance(left, list) else len(right)
        return [_exact([left[i] if isinstance(left, list) else left,
                        right[i] if isinstance(right, list) else right])
                for i in range(size)]
    return _text(left) == _text(right)


def _mid(values: list[Any]) -> str:
    text, start, length = _text(values[0]), int(_number(values[1])), int(_number(values[2]))
    if length < 0:
        raise _Coerce(VALUE)
    return text[start - 1: start - 1 + length]


_FUNCTIONS: dict[str, Callable[[list[Any]], Any]] = {
    "AND": lambda v: all(_truthy(x) for x in _flatten(v)),
    "OR": lambda v: any(_truthy(x) for x in _flatten(v)),
    "NOT": lambda v: not _truthy(v[0]),
    "MAX": lambda v: max(_number(x) for x in _flatten(v)),
    "MIN": lambda v: min(_number(x) for x in _flatten(v)),
    "SUM": lambda v: sum(_number(x) for x in _flatten(v)
                         if not isinstance(x, (bool, str))),
    "SUMPRODUCT": _sumproduct,
    "EXACT": lambda v: _exact(v),
    "LEN": lambda v: float(len(_text(v[0]))),
    "TRIM": lambda v: " ".join(_text(v[0]).split()),
    "CHAR": lambda v: chr(int(_number(v[0]))),
    "SUBSTITUTE": lambda v: _text(v[0]).replace(_text(v[1]), _text(v[2])),
    "MID": _mid,
    "FIND": _find,
    "TEXT": lambda v: _text(v[0]),
    "T": lambda v: v[0] if isinstance(v[0], str) else "",
    "N": lambda v: _number(v[0]),
    "INDEX": _index,
    "MATCH": _match,
    "COUNTIF": lambda v: float(sum(1 for x in v[0] if _compare("=", x, v[1]))),
    "ISNA": lambda v: v[0] is NA or (isinstance(v[0], ExcelError) and v[0].name == "#N/A"),
    "ISERROR": lambda v: isinstance(v[0], ExcelError),
}


def evaluate_cell(cells: dict[str, Any], externals: dict[str, Any], address: str) -> Any:
    return Evaluator(cells, externals).value(address)
