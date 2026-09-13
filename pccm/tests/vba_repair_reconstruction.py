#!/usr/bin/env python3
"""P10-2C CORRECTION: take the Repair reconstruction rule back out of modRepair.

WHY A REVERSAL. Final acceptance run 7 at ee6e9fb found a genuine production
defect: Repair Profiling recreated a missing profiling row with its project-year
weights seeded at the profiling owner's initial value, a zero, where the settled
Phase-10 Step-1 contract (§5) says "id only, weights blank - a blank is an unmade
assumption, not a zero", and would have seeded added project years the same
way where the contract says "extend with blanks". The correction lives in
modRepair.bas only. Every control that pinned production to an accepted tree
keeps its claim the way P10-RP established: taking the declared change back out
must reproduce the accepted bytes exactly, so the ONLY thing that moved is the
declared correction.

THE FRAGMENTS ARE GENERATED FROM THE REAL DIFF against ee6e9fb, the last tree
Windows executed before the correction, and applying them all must reproduce
that file byte for byte. modRepair.bas is LF; the fragments carry LF.
"""
from __future__ import annotations

# The tree the reversal must reproduce: the one final acceptance run 7 executed.
ACCEPTED_BEFORE_REPAIR_RECONSTRUCTION = "ee6e9fb"

DECLARED_REPAIR_RECONSTRUCTION_CHANGES = {
    "modRepair.bas":
        "Apply records which permanent ids had a profiling row before the two "
        "owners run and afterwards blanks every project-year weight the owners "
        "had to reconstruct - a recreated row entirely, an added project year for "
        "every row - through modProfiling.SetValueFor; the recognised-empty "
        "comment and the success wording say so",
}

# (current fragment, accepted fragment), in file order.
_HUNKS: dict[str, tuple[tuple[str, str], ...]] = {
    "modRepair.bas": (
        ("'                                 with every weight following its permanent id\n'\n' plus the one rule this command adds on top of them, applied through the same\n' owner's (permanent id, project year) accessor and never through an address of\n' its own: what those two calls had to RECONSTRUCT is left blank (see Apply).\n'\n' And the decision to act is taken against the ACCEPTED structural checker, not\n' against a second opinion assembled here: modStructuralCheck.ValidateStructure\n",
         "'                                 with every weight following its permanent id\n'\n' And the decision to act is taken against the ACCEPTED structural checker, not\n' against a second opinion assembled here: modStructuralCheck.ValidateStructure\n"),
        ('\' "a profile" means. A control asserts the two literals agree.\n\'\n\' 0% is the total this command RECOGNISES as "no profile yet": a row whose\n\' weights are all zero or blank is a driver nobody has profiled, so restructuring\n\' it moves no allocation. Zero is the profiling owner\'s ordinary initial value,\n\' the one the Add and Apply / Update Timeline path seeds a new driver and a new\n\' project year with. It is NOT what this command writes: a row or a project year\n\' Repair Profiling has to reconstruct is left BLANK (see Apply), because a blank\n\' is an unmade assumption and a zero is a typed one. Neither recognised total is\n\' a semantic choice this command could get wrong. Anything between them is a\n\' half-entered profile, and restructuring one is refused rather than guessed at.\nPrivate Const REPAIR_PROFILE_SUM_TARGET As Double = 1#\nPrivate Const REPAIR_PROFILE_SUM_EMPTY As Double = 0#\n',
         '\' "a profile" means. A control asserts the two literals agree.\n\'\n\' 0% is the contract\'s EMPTY state, not an invalid one: PROFILE_INITIAL_VALUE is\n\' what SetYearColumns seeds a new project year with and what SyncRows gives a\n\' newly identified driver, so a row that totals zero is a driver nobody has\n\' profiled yet. Neither of these two is a semantic choice this command could get\n\' wrong. Anything between them is a half-entered profile, and restructuring one\n\' is refused rather than guessed at.\nPrivate Const REPAIR_PROFILE_SUM_TARGET As Double = 1#\nPrivate Const REPAIR_PROFILE_SUM_EMPTY As Double = 0#\n'),
        ('End Function\n\n\' THE REPAIR ITSELF: TWO OWNER CALLS, THEN THE ONE RULE THIS COMMAND ADDS.\n\' Columns before rows, in the order Apply / Update Timeline uses them: SyncRows\n\' preserves a weight by (permanent id, project-year index), so the index set has\n\' to be the right one before ownership is re-established over it.\n\'\n\' AND THEN WHAT THE OWNERS RECONSTRUCTED IS LEFT BLANK. The two owners are the\n\' ordinary Add / Apply Timeline path, and on that path a driver that has never\n\' been profiled and a project year that has just been applied are seeded with\n\' the profiling owner\'s initial value, which is a zero. Repair Profiling is not\n\' that path. It is the recovery command whose contract says a row it had to\n\' recreate holds "id only, weights blank - a blank is an unmade assumption, not\n\' a zero", and that a grid it had to widen is "extended with blanks"; final\n\' acceptance run 7 found a recreated row seeded at zero instead. So the ids that\n\' had a row BEFORE the owners ran are recorded first, and afterwards every\n\' project-year weight of a row the owners had to recreate, and every project-year\n\' position the owners had to add, is cleared - through the profiling owner\'s own\n\' (permanent id, project year) accessor, so this module still names no cell. A\n\' weight that existed before, at a position that existed before, is untouched:\n\' blank stays blank, zero stays zero, a number stays exactly that number. Both\n\' grids are snapshotted before Apply and restored on any failure, so a failure\n\' after the blanking rolls it back with everything else.\nPrivate Sub Apply(ByRef plan As ProfilingPlan)\n    Dim hadRow As Object\n    If Not plan.NeedsRepair Then Exit Sub\n    Set hadRow = PresentIds(plan.Kind)\n    modProfiling.SetYearColumns plan.Kind, plan.StartYear, plan.TargetYears\n    modProfiling.SyncRows plan.Kind\n    BlankReconstructed plan, hadRow\nEnd Sub\n\n\' The permanent ids that have a profiling row NOW, as the owner lists them.\nPrivate Function PresentIds(ByVal kind As String) As Object\n    Dim found As Object\n    Dim listed As String, parts() As String\n    Dim index As Long\n    Set found = CreateObject("Scripting.Dictionary")\n    listed = modProfiling.IdList(kind)\n    If Len(listed) > 0 Then\n        parts = Split(listed, ",")\n        For index = LBound(parts) To UBound(parts)\n            If Not found.Exists(parts(index)) Then found.Add parts(index), True\n        Next index\n    End If\n    Set PresentIds = found\nEnd Function\n\n\' Every weight cell this command RECONSTRUCTED, cleared, and nothing else: all\n\' project years of a row that had no row before the owners ran, and the\n\' project years beyond the previous width of every row when the grid was\n\' widened. A shrink adds no position, so it clears nothing here.\nPrivate Sub BlankReconstructed(ByRef plan As ProfilingPlan, ByVal hadRow As Object)\n    Dim listed As String, ids() As String\n    Dim index As Long, year As Long, firstYear As Long\n    listed = modProfiling.IdList(plan.Kind)\n    If Len(listed) = 0 Then Exit Sub\n    ids = Split(listed, ",")\n    For index = LBound(ids) To UBound(ids)\n        If hadRow.Exists(ids(index)) Then\n            firstYear = plan.ActualYears + 1\n        Else\n            firstYear = 1\n        End If\n        For year = firstYear To plan.TargetYears\n            modProfiling.SetValueFor plan.Kind, ids(index), year, Empty\n        Next year\n    Next index\nEnd Sub\n\n',
         "End Function\n\n' THE REPAIR ITSELF, AND IT IS TWO CALLS. Columns before rows, in the order\n' Apply / Update Timeline uses them: SyncRows preserves a weight by (permanent\n' id, project-year index), so the index set has to be the right one before\n' ownership is re-established over it.\nPrivate Sub Apply(ByRef plan As ProfilingPlan)\n    If Not plan.NeedsRepair Then Exit Sub\n    modProfiling.SetYearColumns plan.Kind, plan.StartYear, plan.TargetYears\n    modProfiling.SyncRows plan.Kind\nEnd Sub\n\n"),
        ('    Summary = out & " Every weight that could be attributed to a driver was preserved " & _\n              "exactly, by permanent identifier and project year."\n    If Reconstructed(cost) Or Reconstructed(risk) Then\n        Summary = Summary & " Rows and project years this command had to reconstruct " & _\n                  "are left blank for you to complete: a blank is an unmade assumption, " & _\n                  "not a zero."\n    End If\nEnd Function\n\n\' Whether a repair recreated a row or added a project year - the two cases whose\n\' weights this command leaves blank.\nPrivate Function Reconstructed(ByRef plan As ProfilingPlan) As Boolean\n    If Not plan.NeedsRepair Then Exit Function\n    Reconstructed = (plan.Missing > 0) Or (plan.TargetYears > plan.ActualYears)\nEnd Function\n\n',
         '    Summary = out & " Every weight that could be attributed to a driver was preserved " & _\n              "exactly, by permanent identifier and project year."\nEnd Function\n\n'),
    ),
}


def strip_repair_reconstruction(module_name: str, text: str) -> str:
    """`text` with the P10-2C correction removed.

    ALL OR NONE. A module the correction never touched, and a text that predates
    the correction (an accepted tree read back through git), carry NONE of the
    fragments and come back unchanged. A text that carries ALL of them has the
    layer taken off exactly. A text that carries SOME of them is neither: a
    fragment moved or something rode along inside it, and that raises, because
    "unchanged" must never be the answer to a partial match. Whether the layer
    is PRESENT where it must be is the Repair suite's own control (test_30).
    """
    hunks = _HUNKS.get(module_name, ())
    if not hunks:
        return text
    matched = []
    for current, accepted in hunks:
        for pair in ((current, accepted),
                     (current.replace("\r\n", "\n"), accepted.replace("\r\n", "\n"))):
            if text.count(pair[0]) == 1:
                matched.append(pair)
                break
        else:
            matched.append(None)
    if all(pair is None for pair in matched):
        return text
    if any(pair is None for pair in matched):
        missing = [hunks[i][0].splitlines()[0] for i, pair in enumerate(matched) if pair is None]
        raise AssertionError(
            f"{module_name}: the declared Repair reconstruction is partially present - "
            f"{len(hunks) - len(missing)} of {len(hunks)} fragments match, so the reversal "
            f"cannot be exact. A fragment moved or something rode along inside it:\n  {missing[0]!r}")
    for current, accepted in matched:  # type: ignore[misc]
        text = text.replace(current, accepted, 1)
    return text


def repair_reconstruction_touches(module_name: str) -> bool:
    return module_name in _HUNKS
