#!/usr/bin/env python3
"""P10-UX MUTATION CONTROLS.

A presentation control that cannot fail is worse than no control at all: it
makes a batch look proved while leaving exactly the class of defect a human had
to find by opening the file. Every control below damages one of the authorities
this batch touches - `spec/workbook.yaml`, `spec/structure_contract.yaml` or the
emitted chart projection - reruns the WHOLE UX battery against the damaged copy,
and requires a NAMED detector among the refusers.

THE TWELVE ARE THE REGRESSIONS THE MANUAL REVIEW WOULD FIND AGAIN: charts put
back to the size they were reviewed at, a plot dropped, a source range moved
behind a layout change, #N/A returned to an axis, an absent value turned into a
blank a chart draws as zero, the tornado's labels unpinned, the compact format
made a lie, two plots overlapping, the timeline button back over the paragraph
it obscured, a button rebound, the Inflation message returned to fine print or
made unconditional, and a second Apply button.

Nothing here writes to the repository: damaged copies live in memory and the
conformance module is pointed at them for one control.

Runs standalone or under pytest.
"""
from __future__ import annotations

import copy
import json
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402
import yaml  # noqa: E402

import test_phase10_ux_presentation as conformance  # noqa: E402


def _tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 12, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _install(shell=None, structure=None, projection=None):
    """Point the conformance module's caches at damaged copies.

    THE CACHE *IS* THE INJECTION POINT, and deliberately so. Every accessor in
    the battery reads through it, so replacing an entry replaces what every
    control sees - including the controls that would otherwise re-read the
    manifest from disk and quietly pass.
    """
    saved = dict(conformance._CACHE)
    conformance._CACHE.clear()
    conformance._CACHE.update(saved)
    if shell is not None:
        conformance._CACHE["shell"] = shell
    if structure is not None:
        conformance._CACHE["structure"] = structure
    if projection is not None:
        conformance._CACHE["projection"] = projection

    def restore() -> None:
        conformance._CACHE.clear()
        conformance._CACHE.update(saved)

    return restore


def _control(expected: str, **damaged) -> None:
    restore = _install(**damaged)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _shell():
    conformance._CACHE.pop("shell", None)
    return copy.deepcopy(conformance._shell())


def _projection():
    conformance._CACHE.pop("projection", None)
    return copy.deepcopy(conformance._projection())


def _structure():
    """A live structure contract, rebuilt from a damaged YAML string."""
    from pccm_builder import load_structure_contract
    conformance._CACHE.pop("structure", None)
    return load_structure_contract(conformance.SPEC / "structure_contract.yaml")


def _damaged_structure(replacements: list[tuple[str, str]]):
    """Load the structure contract from a copy with `replacements` applied."""
    import tempfile
    import shutil
    from pccm_builder import load_structure_contract
    text = (conformance.SPEC / "structure_contract.yaml").read_text(encoding="utf-8")
    for old, new in replacements:
        assert text.count(old) == 1, (old[:60], text.count(old))
        text = text.replace(old, new, 1)
    temp = Path(tempfile.mkdtemp(prefix="pccm-ux-mutation-"))
    shutil.copytree(conformance.SPEC, temp / "spec")
    (temp / "spec" / "structure_contract.yaml").write_text(text, encoding="utf-8")
    return load_structure_contract(temp / "spec" / "structure_contract.yaml")


def _refused_by_the_contract(replacements: list[tuple[str, str]], expected: str) -> None:
    """Some mutations never reach the battery, because the CONTRACT refuses to
    load at all - which is a stronger refusal, not a weaker one.

    A control that pretended otherwise would have to loosen the loader to let the
    damage through, so the detector is named where it actually is.
    """
    from pccm_builder import StructureContractError
    try:
        _damaged_structure(replacements)
    except StructureContractError as error:
        assert expected in str(error), (expected, str(error))
        return
    raise AssertionError("the contract loaded a workbook it should have refused")


def test_00_the_accepted_sources_pass_every_detector() -> None:
    restore = _install()
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused == [], refused


# ===========================================================================
# UX-001. THE CHARTS
# ===========================================================================
def test_01_the_charts_go_back_to_the_size_they_were_reviewed_at() -> None:
    """The exact regression: a later change that quietly restores 9.6 x 7.2."""
    damaged = _projection()
    for chart in damaged["charts"]:
        chart["width_cm"] = conformance.REVIEWED_WIDTH_CM
        chart["height_cm"] = conformance.REVIEWED_HEIGHT_CM
    _control("test_01", projection=damaged)


def test_02_only_one_chart_is_enlarged_and_the_rest_are_left(  ) -> None:
    """Half a correction is the worst outcome: a Dashboard where one plot is
    readable and three are not looks deliberate."""
    damaged = _projection()
    damaged["charts"][1]["height_cm"] = conformance.REVIEWED_HEIGHT_CM
    _control("test_01", projection=damaged)


def test_03_a_chart_is_dropped_to_make_the_others_fit() -> None:
    damaged = _projection()
    damaged["charts"] = [c for c in damaged["charts"] if c["key"] != "tornado"]
    _control("test_02", projection=damaged)


def test_04_a_source_range_moves_behind_the_layout_change() -> None:
    """A presentation batch that repointed a series would be changing what the
    Dashboard says while claiming to change only how it looks."""
    damaged = _shell()
    damaged["charts"]["charts"][0]["series"][0]["key"] = "annual_nominal"
    _control("test_03", shell=damaged)


def test_05_na_returns_to_the_category_axis() -> None:
    """The defect itself: rows of #N/A down an empty chart's axis."""
    damaged = _projection()
    for chart in damaged["charts"]:
        chart["no_data_category"] = "NA()"
    _control("test_04", projection=damaged)


def test_06_an_absent_value_becomes_a_blank_a_chart_plots_as_zero() -> None:
    """THE ONE MISTAKE THE #N/A FIX COULD EASILY HAVE MADE. Blanking the value
    column along with the label draws a project costing nothing."""
    damaged = _projection()
    damaged["charts"][0]["no_data_value"] = '""'
    _control("test_04", projection=damaged)


def test_07_a_value_column_is_blanked_along_with_its_category() -> None:
    """The same mistake one level down, in the bridge rather than the chart."""
    damaged = _projection()
    for column in damaged["bridge"]["annual"]["columns"]:
        column["absent"] = '""'
    _control("test_04", projection=damaged)


def test_08_the_tornado_labels_are_unpinned_from_the_low_end() -> None:
    """Signed rho puts the value axis through the middle of the plot, so an
    unpinned category label lands on the bars."""
    damaged = _projection()
    damaged["axis_presentation"]["category_label_position"] = "nextTo"
    _control("test_05", projection=damaged)


def test_09_the_tornado_loses_its_extra_category_room() -> None:
    damaged = _projection()
    damaged["charts"][-1]["width_cm"] = damaged["charts"][0]["width_cm"]
    _control("test_05", projection=damaged)


def test_10_the_compact_format_becomes_a_single_fixed_scale() -> None:
    """`#,##0,,"M"` alone prints every figure of a half-million-riyal project as
    "0M", which is a compact format that lies about the order of magnitude."""
    damaged = _shell()
    damaged["charts"]["number_formats"]["money_axis"] = '#,##0,,"M"'
    _control("test_06", shell=damaged)


def test_11_the_axis_label_font_goes_back_to_the_body_size() -> None:
    damaged = _projection()
    damaged["axis_presentation"]["label_font_size"] = 1000
    _control("test_06", projection=damaged)


def test_12_two_enlarged_charts_overlap() -> None:
    """Bigger is only better if they still fit beside each other."""
    damaged = _projection()
    damaged["charts"][1]["anchor"] = "D65"
    _control("test_07", projection=damaged)


def test_13_a_chart_is_anchored_outside_the_reserved_region() -> None:
    damaged = _projection()
    damaged["charts"][0]["anchor"] = "B20"
    _control("test_07", projection=damaged)


# ===========================================================================
# UX-002. THE TIMELINE BUTTON
# ===========================================================================
def test_14_the_timeline_button_goes_back_over_the_applied_block() -> None:
    """E43 is the section heading row of the block, in that block's note
    column. This is the observation the correction came from."""
    damaged = _damaged_structure([
        ('      shape_name: "btnPCCMApplyTimeline"\n'
         '      caption: "Apply / Update Timeline"\n'
         '      entry_point: "PCCM_ApplyTimeline"\n'
         '      anchor_cell: "E59"',
         '      shape_name: "btnPCCMApplyTimeline"\n'
         '      caption: "Apply / Update Timeline"\n'
         '      entry_point: "PCCM_ApplyTimeline"\n'
         '      anchor_cell: "E43"')])
    _control("test_08", structure=damaged)


def test_15_the_timeline_button_lands_inside_the_applied_values() -> None:
    """Below the heading is not enough: it must clear the values and their
    notes too, and a two-row shape has to clear them by its whole height."""
    damaged = _damaged_structure([('      anchor_cell: "E59"', '      anchor_cell: "E49"')])
    _control("test_08", structure=damaged)


def test_16_the_timeline_button_is_appended_instead_of_leading() -> None:
    """Applying a timeline is what generates the columns every later command
    needs. Last in the list is the wrong instruction to give a new user."""
    damaged = _damaged_structure([('      anchor_cell: "E59"', '      anchor_cell: "E71"')])
    _control("test_09", structure=damaged)


def test_17_the_timeline_button_is_rebound() -> None:
    """The most expensive single-line mistake available in a layout batch."""
    _refused_by_the_contract(
        [('      entry_point: "PCCM_ApplyTimeline"', '      entry_point: "PCCM_Calculate"')],
        "are bound to more than one button")


def test_18_an_accepted_command_binding_changes_with_the_layout() -> None:
    damaged = _damaged_structure([
        ('      caption: "Run Sensitivity"', '      caption: "Sensitivity"')])
    _control("test_10", structure=damaged)


def test_19_a_register_button_is_dragged_along_by_the_shift() -> None:
    """The Cost Lines and Risk Register buttons are on other sheets and had no
    reason to move. A block shift that reached them would be a bug nobody would
    look for."""
    damaged = _damaged_structure([('      anchor_cell: "N6"', '      anchor_cell: "N8"')])
    _control("test_10", structure=damaged)


# ===========================================================================
# UX-003. THE INFLATION MESSAGE
# ===========================================================================
def test_20_the_inflation_message_goes_back_to_stating_a_fact() -> None:
    """"Timeline not yet applied" is true and was read straight past. The
    correction is that it now says where to go and what to press."""
    # BOTH SIDES ARE CHANGED TOGETHER, so the contract still loads and the
    # battery is the thing being tested. Changing only the declaration is
    # refused by the loader before the battery runs - correctly, but it would
    # prove the loader rather than these controls.
    was = ("No timeline applied. Go to Setup and select Apply / Update Timeline "
           "to generate the inflation-year columns.")
    damaged = _damaged_structure([
        (f'  inflation_not_applied: "{was}"',
         '  inflation_not_applied: "Timeline not yet applied."'),
        ('    =IF(nmYearCount_Applied="","No timeline applied. Go to Setup and select\n'
         '    Apply / Update Timeline to generate the inflation-year columns.",',
         '    =IF(nmYearCount_Applied="","Timeline not yet applied.",')])
    _control("test_11", structure=damaged)


def test_21_the_message_becomes_unconditional() -> None:
    """A message that does not clear itself is a message that lies as soon as
    the user does what it asked."""
    damaged = _damaged_structure([
        ('  inflation_formula: >-\n'
         '    =IF(nmYearCount_Applied="","No timeline applied. Go to Setup and select\n',
         '  inflation_formula: >-\n'
         '    =IF(TRUE,"No timeline applied. Go to Setup and select\n')])
    _control("test_12", structure=damaged)


def test_22_the_message_reads_a_name_the_structural_owner_does_not_publish() -> None:
    """A second state authority, introduced by reading something else."""
    damaged = _damaged_structure([
        ('    IF(nmInflFirstYear>nmInflLastYear,',
         '    IF(nmInflationReady>nmInflLastYear,')])
    _control("test_12", structure=damaged)


def test_23_a_second_apply_button_appears_on_inflation() -> None:
    """Explicitly forbidden, and for a good reason: two buttons for one command
    is two places for it to be wrong."""
    _refused_by_the_contract(
        [('    - key: "add_cost_line"',
          '    - key: "apply_timeline_inflation"\n'
          '      sheet: "Inflation"\n'
          '      shape_name: "btnPCCMApplyTimelineAgain"\n'
          '      caption: "Apply / Update Timeline"\n'
          '      entry_point: "PCCM_ApplyTimeline"\n'
          '      anchor_cell: "N6"\n\n'
          '    - key: "add_cost_line"')],
        "are bound to more than one button")


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
