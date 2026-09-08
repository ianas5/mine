#!/usr/bin/env python3
"""The Phase-9 Model Check plan: WHAT the sheet shows and WHERE every part sits.

MODEL CHECK OWNS NO FACT. Every value it renders is a presentation of something
an accepted owner already decided - the structural report, the live calculation
state, the live simulation state, the four annual accessors, the Sensitivity
sheet's own availability sentence, and the requested iteration count. Nothing
here derives a state, invents a word, or creates a refusal.

WHY A PLAN OBJECT AND NOT A RENDERER. Three consumers need the same answers: the
renderer that writes the cells, the projection a Windows runner reads, and the
verifier that decides which cells are allowed to hold a formula. If each worked
the geometry out for itself there would be three authorities for one layout, and
they would agree right up until one of them moved. This is the one.

THE ORDER IS FIXED AT BUILD TIME. A check's severity, group and check id are
contract text; only WHETHER IT APPLIES is a runtime question. So the register's
order - severity, then group in the declared order, then check id - is computed
here, once, and the sheet renders the k-th ACTIVE candidate in that fixed order.
Discovery order is not an input to anything below.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from .artifact_io import write_lf_artifact
from .contract_loader import InputContract
from .spec_loader import WorkbookSpec

SCHEMA_VERSION = 1

INSPECTION_FILENAME = "phase9_model_check_inspection.json"

ALLOWED_KEYS = ("schema_version", "purpose", "provenance", "sheet", "freeze_panes",
                "vocabulary", "summary", "register", "evaluation", "advisory",
                "sources", "number_formats", "worksheet_safety")

# THE ONE INPUT THAT CARRIES A RECOMMENDATION. Named here so the resolver can
# ask the input contract for it by name and fail loudly when it is gone, rather
# than searching for "an input that happens to have one".
ADVISORY_INPUT_KEY = "monte_carlo_iterations"

# A reading reference inside a condition, subject or message.
_PLACEHOLDER = re.compile(r"\{([a-z_]+)\}")

# The reading key the ADVISORY threshold resolves through. It is not a reading -
# it is a value the input contract owns - and it is spelled once, here.
_RECOMMENDATION_TOKEN = "recommended_iterations"

# The message and guidance sentences write the threshold for a reader. This is
# the token they use; the value behind it is the contract's, formatted.
_RECOMMENDATION_TEXT_TOKEN = "<recommended>"


class ModelCheckPlan:
    """Where every Model Check cell sits and what it holds.

    Built once from the manifest, the input contract and the accepted Phase-6/7/8
    blocks it mirrors. Everything a caller can ask is derived; nothing is typed.
    """

    def __init__(self, spec: WorkbookSpec, contract: InputContract,
                 results: dict[str, Any], sensitivity: dict[str, Any]) -> None:
        self._spec = spec
        shell = (spec.phase9_shell or {}).get("model_check")
        if not shell:
            raise ValueError(
                "workbook.yaml carries no phase9_shell.model_check; there is no Model "
                "Check surface to plan")
        self.block = shell
        self.sheet = str(shell["sheet"])
        self.label_column = str(shell["label_column"])
        self.value_column = str(shell["value_column"])
        self.severity_order = [str(s) for s in shell["severity_order"]]
        self.group_order = [str(g) for g in shell["group_order"]]
        self.actionable = [str(s) for s in shell["actionable_severities"]]
        self.informational = [s for s in self.severity_order if s not in self.actionable][0]
        self.row_window = int(shell["row_window"])
        self.no_data_formula = str(shell["no_data_formula"])
        self.mirror_formula = str(shell["mirror_formula"])
        self.number_formats = {str(k): str(v) for k, v in shell["number_formats"].items()}
        self.summary = shell["summary"]
        self.register = shell["register"]
        self.evaluation = shell["evaluation"]
        self.readings = self.evaluation["readings"]
        self.candidates = self.evaluation["candidates"]
        self.structural = shell["structural"]
        self.structural_slots = int(self.candidates["structural_slots"])

        # THE THRESHOLD, ASKED FOR BY NAME. `recommendation_for` refuses when the
        # field is gone, so a removed recommendation fails the BUILD instead of
        # quietly producing a Model Check that never advises.
        self.recommended_iterations = contract.recommendation_for(ADVISORY_INPUT_KEY)

        self._results = results
        self._sensitivity = sensitivity
        self._summary_rows = {str(e["key"]): int(e["row"]) for e in self.summary["rows"]}
        self._reading_rows = {str(e["key"]): int(e["row"]) for e in self.readings["rows"]}
        self._candidate_columns = {str(c["key"]): str(c["column"])
                                   for c in self.candidates["columns"]}
        self.ordered_checks = self._order(shell["checks"])

        first = int(self.candidates["first_row"])
        self.candidate_first_row = first
        self.candidate_last_row = first + self.structural_slots + len(self.ordered_checks) - 1
        # THE SLOTS SORT AHEAD OF THE DECLARED REGISTER, so the candidate block
        # IS the ordered population and no sort happens at runtime. The loader
        # proves the premise - structural severity and group are first in both
        # declared orders - and this class relies on it here and nowhere else.
        self.structural_first_row = first
        self.structural_last_row = first + self.structural_slots - 1
        self.declared_first_row = self.structural_last_row + 1

    # -- ordering -----------------------------------------------------------
    def _order(self, checks: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Severity, then group in the declared order, then check id.

        THE WHOLE POINT OF DOING IT HERE. A register sorted at runtime would
        depend on how Excel happened to evaluate it; a register sorted at build
        time cannot. The keys are POSITIONS in the declared vocabularies, so a
        word outside them has no position and the loader has already refused it.
        """
        return sorted(
            checks,
            key=lambda c: (self.severity_order.index(str(c["severity"])),
                           self.group_order.index(str(c["group"])),
                           str(c["check_id"])),
        )

    # -- addresses ----------------------------------------------------------
    def summary_row(self, key: str) -> int:
        try:
            return self._summary_rows[key]
        except KeyError as error:
            raise ValueError(f"the Model Check summary publishes no {key!r} row") from error

    def summary_cell(self, key: str) -> str:
        return f"${self.value_column}${self.summary_row(key)}"

    def reading_row(self, key: str) -> int:
        try:
            return self._reading_rows[key]
        except KeyError as error:
            raise ValueError(f"the Model Check readings publish no {key!r}") from error

    def reading_cell(self, key: str) -> str:
        return f"${self.value_column}${self.reading_row(key)}"

    def candidate_column(self, key: str) -> str:
        try:
            return self._candidate_columns[key]
        except KeyError as error:
            raise ValueError(f"the candidate block declares no {key!r} column") from error

    def candidate_range(self, key: str) -> str:
        column = self.candidate_column(key)
        return (f"${column}${self.candidate_first_row}:"
                f"${column}${self.candidate_last_row}")

    def register_rows(self) -> range:
        first = int(self.register["first_row"])
        return range(first, first + self.row_window)

    # -- readings -----------------------------------------------------------
    def reading_formula(self, entry: dict[str, Any]) -> str:
        kind = str(entry["kind"])
        if kind == "procedure":
            call = f"{entry['procedure']}()"
            if not entry.get("anchor"):
                return f"={call}"
            # THE ANCHOR, AND WHY IT IS NOT A TRICK.
            #
            # A zero-argument function that is not `Application.Volatile` has no
            # dependencies Excel can see, so a cell holding one answers once and
            # then keeps that answer however far the model moves - the defect the
            # P8-1 record names, and the reason the five accepted adapters are
            # volatile.
            #
            # These two owners are NOT ours to make volatile: `PCCM_StructuralReport`
            # and the two attempt accessors are shared with command paths that
            # have no reason to re-run on every calculation cycle. So the CALLER
            # declares the dependency instead, by making the cell a dependent of
            # the volatile `evaluated` heartbeat above it. `T()` of that timestamp
            # is the empty string, so the concatenation changes no value; it
            # changes only what Excel believes this cell depends on.
            #
            # NO LINUX TEST CAN SETTLE THAT IT RE-EVALUATES. It is the first thing
            # the Phase-9 Windows acceptance has to look at.
            return f"=T({self.summary_cell('evaluated')})&{call}"
        if kind == "defined_name":
            return f"={entry['defined_name']}"
        if kind == "mirror":
            rows = self._results_rows()
            block, key = str(entry["source_block"]), str(entry["source_key"])
            try:
                row = rows[block][key]
            except KeyError as error:
                raise ValueError(
                    f"Model Check mirrors {block}.{key}, which the Results block does "
                    "not publish") from error
            reference = f"{self._results['sheet']}!${self._results['nominal_column']}${row}"
            return self.mirror_formula.format(ref=reference)
        if kind == "sensitivity_availability":
            # THE WHOLE SENTENCE, MIRRORED, and the address resolved from the
            # accepted Phase-7 block rather than typed. P7-4 is what a typed
            # address costs: the persisted block moved, the formula did not, and
            # a sheet reported "not produced" after a run that had just succeeded.
            column = str(self._sensitivity["columns"][1]["column"])
            row = int(self._sensitivity["availability_row"])
            reference = f"{self._sensitivity['sheet']}!${column}${row}"
            return self.mirror_formula.format(ref=reference)
        if kind == "derived":
            return self._derived_formula(str(entry["derived"]))
        raise ValueError(f"the Model Check readings declare an unknown kind {kind!r}")

    def _results_rows(self) -> dict[str, dict[str, int]]:
        results = self._results
        return {
            "run_stamp": {str(f["key"]): int(f["row"])
                          for f in results["run_stamp"]["fields"]},
            "summary": {str(m["key"]): int(m["row"])
                        for m in results["summary"]["metrics"]},
        }

    def _derived_formula(self, name: str) -> str:
        report = self.reading_cell("structural_report")
        if name == "structural_fault_count":
            # ONE FAULT PER LINE, COUNTED FROM THE REPORT ITSELF. This is the
            # authoritative total and it is UNBOUNDED: it does not depend on how
            # many slots the sheet has, so a fault population larger than the
            # slot capacity is still reported in full even though it cannot all
            # be drawn. That is the difference between a windowed register and a
            # truncated count.
            return (f'=IF({report}="",0,(LEN({report})-LEN(SUBSTITUTE({report},'
                    f'CHAR(13)&CHAR(10),"")))/2)')
        if name == "structural_surplus":
            return (f"=MAX(0,{self.reading_cell('structural_fault_count')}-"
                    f"{self.structural_slots})")
        raise ValueError(f"unknown derived reading {name!r}")

    # -- candidates ---------------------------------------------------------
    def condition_formula(self, condition: str) -> str:
        """A declared condition with its reading keys resolved to addresses.

        NOT ONE ADDRESS IS TYPED IN THE MANIFEST. A reading that moves takes its
        conditions with it, and a key no reading publishes fails the build by
        name - the loader refuses it first, and this refuses it again rather
        than formatting a hole into a formula.
        """
        def replace(match: re.Match[str]) -> str:
            token = match.group(1)
            if token == _RECOMMENDATION_TOKEN:
                return str(self.recommended_iterations)
            return self.reading_cell(token)

        return _PLACEHOLDER.sub(replace, condition)

    def text_of(self, value: str) -> str:
        """A message or a guidance sentence, with the threshold written for a
        reader. The number is the contract's, formatted here; there is no
        literal in the manifest, in this file or in any test."""
        return value.replace(_RECOMMENDATION_TEXT_TOKEN,
                             f"{self.recommended_iterations:,}")

    def field_value(self, value: str) -> str:
        """A candidate's check id, subject or message.

        Exactly `{reading}` becomes a reference to that reading; anything else is
        text. The loader has already refused the mixture, so there is no third
        case to guess at.
        """
        tokens = _PLACEHOLDER.findall(value or "")
        if tokens:
            return f"={self.reading_cell(tokens[0])}"
        if not value:
            # A BLANK CELL IS NOT A BLANK VALUE. Excel reads an empty reference
            # back as a hard 0, so a register row whose subject came from a truly
            # empty candidate cell would print `0` where the honest answer is
            # "this issue is model-wide". P8-3's fourth Windows run found exactly
            # that: an empty string reached a sheet as a numeric zero and was
            # read as a driver identity. `=""` is an empty STRING and INDEX
            # returns it as one.
            return '=""'
        return self.text_of(value)

    def structural_slot_rows(self) -> range:
        return range(self.structural_first_row, self.structural_last_row + 1)

    def declared_rows(self) -> range:
        return range(self.declared_first_row, self.candidate_last_row + 1)

    # -- the formulas the renderer writes -----------------------------------
    def candidate_formulas(self) -> dict[int, dict[str, Any]]:
        """Every candidate row, by sheet row, as {column key: value}."""
        cells: dict[int, dict[str, Any]] = {}
        report = self.reading_cell("structural_report")
        col = self.candidate_column
        first = self.candidate_first_row

        for index, row in enumerate(self.structural_slot_rows(), start=1):
            previous = row - 1
            line = f"${col('line')}{row}"
            start = f"${col('line_start')}{row}"
            end = f"${col('line_end')}{row}"
            cells[row] = {
                "ordinal": index,
                "line_start": ("=1" if index == 1
                               else f"=IF(${col('line_end')}{previous}=0,0,"
                                    f"${col('line_end')}{previous}+2)"),
                "line_end": (f"=IF(OR({start}=0,{start}>LEN({report})),0,"
                             f"IFERROR(FIND(CHAR(13)&CHAR(10),{report},{start}),0))"),
                "line": f'=IF({end}=0,"",MID({report},{start},{end}-{start}))',
                # THE OWNER'S OWN IDENTITY, SPLIT WHERE THE OWNER PUT THE BRACKETS.
                # `check_id` is the invariant that failed; the remainder is how the
                # owner identified WHAT failed. Model Check parses no permanent id
                # of its own invention out of the prose.
                "check_id": (f'=IF({line}="","",MID({line},FIND("[",{line})+1,'
                             f'FIND("]",{line})-FIND("[",{line})-1))'),
                "group": str(self.structural["group"]),
                "severity": str(self.structural["severity"]),
                "subject": (f'=IF({line}="","",'
                            f'TRIM(MID({line},FIND("]",{line})+1,LEN({line}))))'),
                "message": f"=${col('subject')}{row}",
                "guidance": str(self.structural["guidance"]),
                "active": f'=IF({line}="",0,1)',
            }
        for index, check in enumerate(self.ordered_checks, start=1):
            row = self.declared_first_row + index - 1
            cells[row] = {
                "ordinal": self.structural_slots + index,
                "check_id": str(check["check_id"]),
                "group": str(check["group"]),
                "severity": str(check["severity"]),
                "subject": self.field_value(str(check.get("subject") or "")),
                "message": self.field_value(str(check["message"])),
                "guidance": self.text_of(str(check["guidance"])),
                "active": f'=IF({self.condition_formula(str(check["condition"]))},1,0)',
            }

        # NO CANDIDATE CELL MAY BE EMPTY, for the reason `field_value` states: an
        # empty source cell is read back through INDEX as a hard zero, and a zero
        # in the register would be an issue with a fabricated identity.
        blank = sorted((row, key) for row, entry in cells.items()
                       for key, value in entry.items()
                       if value is None or value == "")
        if blank:
            raise ValueError(f"the Model Check candidate block leaves cells empty: {blank}")

        for row in range(first, self.candidate_last_row + 1):
            active = f"${col('active')}{row}"
            key = f"${col('dedup_key')}{row}"
            entry = cells[row]
            entry["dedup_key"] = (f'=IF({active}=0,"",${col("check_id")}{row}&"|"&'
                                  f'${col("subject")}{row})')
            if row == first:
                entry["counted"] = f"={active}"
                entry["sequence"] = f"=${col('counted')}{row}"
            else:
                # DEDUPLICATION BY check_id + subject, AND BY NOTHING ELSE.
                #
                # EXACT rather than COUNTIF on purpose: COUNTIF reads its
                # criterion as a PATTERN, so a subject carrying * or ? - a
                # perfectly ordinary thing for an owner's fault text to carry -
                # would match rows it is not equal to and suppress a real issue.
                # EXACT compares literally, and case-sensitively, which is what a
                # permanent id deserves.
                previous_keys = f"${col('dedup_key')}{first}:${col('dedup_key')}{row - 1}"
                entry["counted"] = (f"=IF(AND({active}=1,"
                                    f"SUMPRODUCT(--EXACT({previous_keys},{key}))=0),1,0)")
                entry["sequence"] = (f"=${col('sequence')}{row - 1}+"
                                     f"${col('counted')}{row}")
        return cells

    def register_formula(self, position: int, field: str) -> str:
        """The `position`-th displayed row, for one register column.

        THE k-TH COUNTED CANDIDATE, FOUND BY ITS RUNNING SEQUENCE. `sequence` is
        non-decreasing and rises by exactly one at each counted row, so the FIRST
        row whose sequence equals k IS the k-th counted row. No sort, no array
        formula, no dependence on anything but the fixed candidate order.

        AND AN UNUSED SLOT IS #N/A, NOT BLANK. Excel reads "" back as a hard
        zero and would happily present it as an issue with no identity; NA() is
        the accepted no-data representation, exactly as it is for an empty
        tornado slot. There is deliberately no IFERROR here: a genuine error
        arriving from an owner must reach the register as an error rather than be
        laundered into "no such check".
        """
        sequence = self.candidate_range("sequence")
        source = self.candidate_range(field)
        last = f"${self.candidate_column('sequence')}${self.candidate_last_row}"
        return (f"=IF({position}>{last},{self.no_data_formula[1:]},"
                f"INDEX({source},MATCH({position},{sequence},0)))")

    def summary_formula(self, key: str) -> str:
        counted = self.candidate_range("counted")
        severity = self.candidate_range("severity")
        errors, warnings, info = (self.summary_cell(name) for name in
                                  ("error_count", "warning_count", "info_count"))
        if key == "overall_status":
            # PRECEDENCE, AND IT READS ONLY THE ACTIONABLE COUNTS. The
            # informational count is not in this formula at all, which is the
            # strongest form of "INFO never changes the Overall Status".
            passing, warning, error = (self.severity_state(name) for name in
                                       ("PASS", "WARNING", "ERROR"))
            return (f'=IF({errors}>0,"{error}",IF({warnings}>0,"{warning}",'
                    f'"{passing}"))')
        if key in ("error_count", "warning_count", "info_count"):
            word = {"error_count": "ERROR", "warning_count": "WARNING",
                    "info_count": "INFO"}[key]
            severity_word = self.severity_word(word)
            total = f'=SUMPRODUCT({counted},--({severity}="{severity_word}"))'
            # THE FAULTS THAT HAVE NO SLOT ARE STILL COUNTED. They cannot be
            # drawn - the register is a hundred rows and they sort past it - but
            # the population is the owner's, not the window's.
            if severity_word == str(self.structural["severity"]):
                total += f"+{self.reading_cell('structural_surplus')}"
            return total
        if key == "total_checks":
            # RECONCILIATION BY CONSTRUCTION. The total is the sum of the three
            # counts above it, so a reader can add up the register and get this
            # number, and no arrangement of the source block can make it drift.
            return f"={errors}+{warnings}+{info}"
        if key == "evaluated":
            return "=NOW()"
        if key == "disclosure":
            total_cell = self.summary_cell("total_checks")
            template = str(self.block["disclosure_template"]).replace(
                "{window}", str(self.row_window))
            before, after = template.split("{total}")
            return (f'=IF({total_cell}>{self.row_window},'
                    f'"{before}"&{total_cell}&"{after}","")')
        raise ValueError(f"the Model Check summary declares no formula for {key!r}")

    def severity_word(self, name: str) -> str:
        if name not in self.severity_order:
            raise ValueError(f"{name!r} is not a declared severity")
        return name

    def severity_state(self, name: str) -> str:
        states = [str(s) for s in self.block["overall_states"]]
        if name not in states:
            raise ValueError(f"{name!r} is not a declared overall state")
        return name

    # -- the cells the verifier permits -------------------------------------
    def formula_cells(self) -> set[str]:
        """Every cell this plan writes a formula into, and no other.

        ENUMERATED, NOT WAVED THROUGH. A formula anywhere else on this sheet -
        in a label column, in a row nothing declares, in the gap between the
        register and the evaluation block - is still a build failure, which is
        the only reason the sheet-wide ban is worth anything.
        """
        cells = {self.summary_cell(key).replace("$", "")
                 for key in self._summary_rows}
        for entry in self.readings["rows"]:
            cells.add(f"{self.value_column}{int(entry['row'])}")
        for row in self.register_rows():
            for column in self.register["columns"]:
                cells.add(f"{column['column']}{row}")
        for row, entry in self.candidate_formulas().items():
            for key, value in entry.items():
                if isinstance(value, str) and value.startswith("="):
                    cells.add(f"{self.candidate_column(key)}{row}")
        return cells


# ===========================================================================
# THE PROJECTION
# ===========================================================================
# IDENTITIES ONLY, on exactly the terms the Phase-5, 6, 7 and 8 projections hold
# theirs: addresses, vocabularies, procedure names and the wording this sheet
# owns. No expected value and no threshold of its own - the threshold below is
# the INPUT CONTRACT's, carried so a Windows runner can check the advisory
# against the same number the sheet was built from rather than against one typed
# into a harness.


def build_phase9_inspection(plan: ModelCheckPlan) -> dict[str, Any]:
    register = plan.register
    return {
        "schema_version": SCHEMA_VERSION,
        "purpose": (
            "Where the Phase-9 Model Check surface sits and what it aggregates. "
            "Addresses, vocabularies and wording projected from workbook.yaml and "
            "the accepted contracts; the advisory threshold is the input "
            "contract's, carried rather than restated."
        ),
        "provenance": {
            "workbook_manifest": "workbook.yaml",
            "threshold_owner": f"input_contract.yaml -> inputs.{ADVISORY_INPUT_KEY}"
                               " -> recommended_iterations",
        },
        "sheet": plan.sheet,
        "freeze_panes": _freeze_panes(plan),
        "vocabulary": {
            "severity_order": list(plan.severity_order),
            "group_order": list(plan.group_order),
            "actionable_severities": list(plan.actionable),
            "informational_severity": plan.informational,
            "overall_states": [str(s) for s in plan.block["overall_states"]],
        },
        "summary": {
            "label_column": plan.label_column,
            "value_column": plan.value_column,
            "heading_row": int(plan.summary["heading_row"]),
            "rows": {str(entry["key"]): {"row": int(entry["row"]),
                                         "label": str(entry["label"]),
                                         "format": str(entry["format"])}
                     for entry in plan.summary["rows"]},
            "precedence": ("any actionable ERROR -> ERROR; else any actionable "
                           "WARNING -> WARNING; else PASS. INFO never counts."),
        },
        "register": {
            "heading_row": int(register["heading_row"]),
            "header_row": int(register["header_row"]),
            "first_row": int(register["first_row"]),
            "row_window": plan.row_window,
            "last_row": int(register["first_row"]) + plan.row_window - 1,
            "no_data_formula": plan.no_data_formula,
            "columns": [{"key": str(c["key"]), "header": str(c["header"]),
                         "column": str(c["column"])}
                        for c in register["columns"]],
            # THE ROWS THAT REPORT AN ABSENCE RATHER THAN A FAULT, named so a
            # runner can assert they are never counted without having to know
            # which check ids they happen to be.
            "optional_publications": [str(check["check_id"])
                                      for check in plan.ordered_checks
                                      if check.get("optional_publication")],
            "disclosure_row": plan.summary_row("disclosure"),
            "disclosure_template": str(plan.block["disclosure_template"]),
            # THE OVERFLOW CLAIM, SPELLED OUT so a runner checks the right thing:
            # the counts are of the whole population and only the REGISTER is
            # windowed.
            "overflow_rule": ("the counts and the total are the full logical "
                              "population; only the register is windowed, and the "
                              "disclosure row says so whenever the total exceeds "
                              "the window"),
        },
        "evaluation": {
            "heading_row": int(plan.evaluation["heading_row"]),
            "readings": {
                "header_row": int(plan.readings["header_row"]),
                "first_row": int(plan.readings["first_row"]),
                "rows": {str(entry["key"]): {
                    "row": int(entry["row"]),
                    "label": str(entry["label"]),
                    "kind": str(entry["kind"]),
                    "source": _reading_source(entry),
                    "anchored": bool(entry.get("anchor", False)),
                } for entry in plan.readings["rows"]},
            },
            "candidates": {
                "header_row": int(plan.candidates["header_row"]),
                "first_row": plan.candidate_first_row,
                "last_row": plan.candidate_last_row,
                "structural_slots": plan.structural_slots,
                "structural_first_row": plan.structural_first_row,
                "structural_last_row": plan.structural_last_row,
                "declared_first_row": plan.declared_first_row,
                "columns": {str(c["key"]): str(c["column"])
                            for c in plan.candidates["columns"]},
                "dedup_key": "check_id + subject",
                "order": ("severity, then group in the declared order, then check "
                          "id; fixed at build time and never at runtime"),
            },
            # THE ORDERED REGISTER, so a runner can assert the k-th displayed row
            # against the plan rather than re-deriving the sort.
            "declared_checks": [
                {"ordinal": plan.structural_slots + index,
                 "check_id": str(check["check_id"]),
                 "group": str(check["group"]),
                 "severity": str(check["severity"]),
                 "row": plan.declared_first_row + index - 1}
                for index, check in enumerate(plan.ordered_checks, start=1)
            ],
        },
        "advisory": {
            "check_id": _advisory_check(plan)["check_id"],
            "input": ADVISORY_INPUT_KEY,
            "reading": "requested_iterations",
            "threshold": plan.recommended_iterations,
            "comparison": "strictly less than",
            "severity": str(_advisory_check(plan)["severity"]),
            "message": plan.text_of(str(_advisory_check(plan)["message"])),
            "guidance": plan.text_of(str(_advisory_check(plan)["guidance"])),
            "refuses": False,
        },
        "sources": {
            "calculation_state": "modResultsState.PCCM_ModelCheckCalculationState",
            "simulation_state": "modResultsState.PCCM_ResultsSimulationState",
            "structural": "modStructuralCheck.PCCM_StructuralReport",
            "annual": [str(entry["procedure"]) for entry in plan.readings["rows"]
                       if entry["kind"] == "procedure"
                       and str(entry["key"]).startswith("annual_")],
        },
        "number_formats": dict(plan.number_formats),
        # THE CLAIM A WINDOWS RUN HAS TO CHECK FIRST. Every one of these is called
        # from a cell; none of them may write, persist, consume or command.
        "worksheet_safety": {
            "cell_called": sorted({str(entry["procedure"])
                                   for entry in plan.readings["rows"]
                                   if entry["kind"] == "procedure"}),
            "forbidden": ["worksheet write", "status persistence", "nonce consumption",
                          "run id mutation", "bank publication",
                          "command endpoint invocation", "application state mutation"],
            "persisting_entry_points_not_reachable": [
                "modCalcReport.PCCM_CalculationStatus",
                "modSimReport.PCCM_SimulationStatus",
            ],
        },
    }


def _advisory_check(plan: ModelCheckPlan) -> dict[str, Any]:
    """The one declared check whose condition compares against the contract's
    recommendation. Found by the token, never by its id: an id typed here would
    be a second name for a thing the manifest already names."""
    found = [c for c in plan.ordered_checks
             if "{" + _RECOMMENDATION_TOKEN + "}" in str(c["condition"])
             and str(c["severity"]) in plan.actionable]
    if len(found) != 1:
        raise ValueError(
            f"{len(found)} actionable checks compare against the recommendation; the "
            "advisory must be exactly one")
    return found[0]


def _reading_source(entry: dict[str, Any]) -> str:
    for key in ("procedure", "defined_name", "derived"):
        if key in entry:
            return str(entry[key])
    if entry["kind"] == "mirror":
        return f"Results.{entry['source_block']}.{entry['source_key']}"
    return "Sensitivity.availability"


def _freeze_panes(plan: ModelCheckPlan) -> str:
    """The sheet's own freeze, cross-checked against the block that promises it.

    TWO DECLARATIONS OF ONE FACT IS THE FAILURE MODE this project keeps finding,
    so the sheet entry stays the authority and the block's copy exists only to be
    compared with it. A mismatch fails the build naming both.
    """
    declared = str(plan.block.get("freeze_panes") or "")
    sheet = next((s for s in plan._spec.sheets if s.name == plan.sheet), None)
    actual = str(getattr(sheet, "freeze_panes", "") or "")
    if declared != actual:
        raise ValueError(
            f"phase9_shell.model_check.freeze_panes is {declared!r} but the "
            f"{plan.sheet!r} sheet declares {actual!r}")
    return actual


def validate_phase9_inspection(inspection: dict[str, Any]) -> None:
    unexpected = sorted(set(inspection) - set(ALLOWED_KEYS))
    if unexpected:
        raise ValueError(f"{INSPECTION_FILENAME}: unexpected key(s) {unexpected}")
    for key in ALLOWED_KEYS:
        if key not in inspection:
            raise ValueError(f"{INSPECTION_FILENAME}: omits {key!r}")

    vocabulary = inspection["vocabulary"]
    if vocabulary["informational_severity"] in vocabulary["actionable_severities"]:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the informational severity is also actionable; an "
            "INFO row could then move the Overall Status")

    register = inspection["register"]
    if register["last_row"] - register["first_row"] + 1 != register["row_window"]:
        raise ValueError(f"{INSPECTION_FILENAME}: the register window does not fit its rows")
    if register["last_row"] >= inspection["evaluation"]["heading_row"]:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the register reaches the evaluation block")
    if not register["no_data_formula"].startswith("="):
        raise ValueError(f"{INSPECTION_FILENAME}: the no-data representation is not a formula")

    optional = register["optional_publications"]
    if not optional:
        raise ValueError(
            f"{INSPECTION_FILENAME}: no optional publication is declared; the "
            "sheet would have nothing to report as absent-but-not-defective")
    declared = {str(entry["check_id"]): str(entry["severity"])
                for entry in inspection["evaluation"]["declared_checks"]}
    for check_id in optional:
        severity = declared.get(check_id)
        if severity in vocabulary["actionable_severities"]:
            raise ValueError(
                f"{INSPECTION_FILENAME}: the optional publication {check_id} is "
                f"{severity}; its absence is not a defect")

    candidates = inspection["evaluation"]["candidates"]
    if candidates["structural_slots"] < register["row_window"]:
        raise ValueError(
            f"{INSPECTION_FILENAME}: {candidates['structural_slots']} structural slots "
            f"cannot fill a {register['row_window']}-row window")
    if candidates["dedup_key"] != "check_id + subject":
        raise ValueError(
            f"{INSPECTION_FILENAME}: the duplicate rule is {candidates['dedup_key']!r}, "
            "not check_id + subject")

    advisory = inspection["advisory"]
    if advisory["comparison"] != "strictly less than":
        raise ValueError(
            f"{INSPECTION_FILENAME}: the advisory compares {advisory['comparison']!r}; "
            "the threshold itself must never warn")
    if advisory["refuses"]:
        raise ValueError(f"{INSPECTION_FILENAME}: an advisory that refuses is not an advisory")
    if advisory["severity"] not in vocabulary["actionable_severities"]:
        raise ValueError(f"{INSPECTION_FILENAME}: the advisory is not actionable")
    threshold = advisory["threshold"]
    if not isinstance(threshold, int) or threshold < 1:
        raise ValueError(f"{INSPECTION_FILENAME}: the advisory threshold is {threshold!r}")
    formatted = f"{threshold:,}"
    for field_name in ("message", "guidance"):
        if formatted not in advisory[field_name]:
            raise ValueError(
                f"{INSPECTION_FILENAME}: the advisory {field_name} does not carry the "
                f"contract's threshold {formatted}")

    safety = inspection["worksheet_safety"]
    reachable = set(safety["cell_called"])
    for persisting in safety["persisting_entry_points_not_reachable"]:
        if persisting.split(".")[-1] in reachable:
            raise ValueError(
                f"{INSPECTION_FILENAME}: {persisting} is called from a cell; it persists")


def emit_phase9_model_check(spec: WorkbookSpec, contract: InputContract,
                            build_dir: Path) -> Path:
    shell = spec.phase6_shell or {}
    plan = ModelCheckPlan(spec, contract, shell["results"], shell["sensitivity"])
    inspection = build_phase9_inspection(plan)
    validate_phase9_inspection(inspection)
    path = build_dir / INSPECTION_FILENAME
    write_lf_artifact(path, json.dumps(inspection, indent=2) + "\n")
    return path
