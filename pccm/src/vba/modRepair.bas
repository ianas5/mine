Attribute VB_Name = "modRepair"
Option Explicit

' ==========================================================================
' REPAIR PROFILING - STRUCTURE ONLY, AND ONLY WHERE IT IS RECONSTRUCTABLE
' ==========================================================================
' WHAT IT IS FOR. The Cost Profiling and Risk Profiling grids are GENERATED:
' one row per identified driver in register order, one column per applied
' project year. Add, Delete and Apply / Update Timeline keep them that way. A
' structural mishap - an interrupted operation, a hand-edited row, a workbook
' recovered by Excel - can leave a row missing, a row orphaned, the order drifted
' or the year columns the wrong width. The structural checker REPORTS all of that
' and refuses to act on it, by design, so until now the only route back was by
' hand on a protected sheet.
'
' WHAT IT IS NOT. It is not a business-input reset: it clears no weight the user
' typed and no driver they entered. It is not a normaliser: a profile that does
' not total 100% stays exactly as they left it, because 100% is a Model Check
' rule and this command owns no semantics.
'
' -------------------------------------------------------------------------
' THE ONE SENTENCE THE WHOLE MODULE IMPLEMENTS
' -------------------------------------------------------------------------
'   Repair structure when structure is reconstructable; refuse when semantic
'   intent is ambiguous.
'
' Every refusal below is a case where the workbook admits more than one honest
' answer, and choosing one would be this command inventing user intent.
'
' -------------------------------------------------------------------------
' IT OWNS NO GEOMETRY AND NO RULE
' -------------------------------------------------------------------------
' Not one grid address, column letter, register column or ID pattern appears
' here. The repair itself is two calls per grid, to the two owners Apply /
' Update Timeline already uses:
'
'   modProfiling.SetYearColumns   the project-year columns and their headers
'   modProfiling.SyncRows         one row per register ID, in register order,
'                                 with every weight following its permanent id
'
' And the decision to act is taken against the ACCEPTED structural checker, not
' against a second opinion assembled here: modStructuralCheck.ValidateStructure
' reports every fault by its contract-declared key, and this command refuses
' unless every reported key is one it is contracted to repair. A structural rule
' added later is therefore blocking by default rather than silently ignored.
'
' -------------------------------------------------------------------------
' WHY THERE IS NO CONFIRMATION PROMPT
' -------------------------------------------------------------------------
' Reset Results asks, because it destroys published answers. Repair Profiling
' does not ask, because it refuses every case that would lose a weight: a shrink
' that would trim data is a refusal, not a prompt. A command that cannot destroy
' anything has nothing to warn about, and a warning it cannot justify would
' teach the user to click through the ones that matter.
' ==========================================================================

' The stages the Windows harness may fail this command after. Between the two
' grids and after both, because "Cost repaired and Risk broken" is the outcome
' the transaction exists to make impossible.
Public Const FAILPOINT_REPAIR_COST As String = "Phase10RepairCost"
Public Const FAILPOINT_REPAIR_RISK As String = "Phase10RepairRisk"

' THE TWO PROFILE SUMS THIS COMMAND RECOGNISES.
'
' 100% is the value modCalcCheck compares a resolved profile against, and it is
' expressed here as a literal for the same reason it is expressed there: the
' calculation contract owns the TOLERANCE, which is the generated
' TOL_PROFILING_SUM_ABSOLUTE below, and nobody owns the target because it is what
' "a profile" means. A control asserts the two literals agree.
'
' 0% is the contract's EMPTY state, not an invalid one: PROFILE_INITIAL_VALUE is
' what SetYearColumns seeds a new project year with and what SyncRows gives a
' newly identified driver, so a row that totals zero is a driver nobody has
' profiled yet. Neither of these two is a semantic choice this command could get
' wrong. Anything between them is a half-entered profile, and restructuring one
' is refused rather than guessed at.
Private Const REPAIR_PROFILE_SUM_TARGET As Double = 1#
Private Const REPAIR_PROFILE_SUM_EMPTY As Double = 0#

' How many offending identifiers a refusal names before it stops listing. A
' refusal that printed four hundred ids would be as unusable as one that named
' none.
Private Const REPAIR_MAX_NAMED As Long = 5

' What one grid needs, decided before anything is touched.
Private Type ProfilingPlan
    Kind As String
    Label As String
    StartYear As Long
    TargetYears As Long
    ActualYears As Long
    Missing As Long
    Orphans As Long
    Duplicates As Long
    OrderDrift As Boolean
    WidthDrift As Boolean
    HeaderDrift As Boolean
    NeedsRepair As Boolean
End Type

' ==========================================================================
' THE ENDPOINT
' ==========================================================================
Public Sub PCCM_RepairProfiling()
    ' THE SAME ENVELOPE THE ACCEPTED ENDPOINTS USE, and for the same reason: a
    ' failure must not leave ScreenUpdating off or Calculation manual.
    Dim state As AppStateSnapshot, result As OperationResult
    Dim stateCaptured As Boolean, cleanupAttempted As Boolean
    Dim cleanup As String, failure As String

    On Error GoTo InvocationFailed
    state = modAppState.CaptureAppState()
    stateCaptured = True
    ' STRUCTURAL, AND IT SAYS SO. Repair Profiling rewrites the project-year ListColumns and re-syncs the
    ' profiling rows, and its rollback rebuilds both grids.
    ' Windows proved a protected sheet refuses all of that with Error 1004, so
    ' this command opens the protection window that FinishOperation closes on
    ' every exit path - including after the rollback below.
    modAppState.BeginStructuralOperation state
    result = RepairProfiling()
    On Error GoTo 0

    On Error GoTo NormalCleanupFailed
    cleanupAttempted = True
    cleanup = modAppState.FinishOperation(state)
    On Error GoTo 0
    stateCaptured = False

    If Len(cleanup) > 0 Then
        result = modAppState.Failed("Repair Profiling", _
            "Application state could not be restored: " & cleanup)
    End If
    modAppState.Announce result
    Exit Sub

NormalCleanupFailed:
    failure = Err.Description
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Repair Profiling", _
        "Application state could not be restored: " & failure)
    Exit Sub

InvocationFailed:
    failure = Err.Description
    On Error GoTo CleanupFailed
    If stateCaptured And Not cleanupAttempted Then
        cleanupAttempted = True
        cleanup = modAppState.FinishOperation(state)
        stateCaptured = False
        If Len(cleanup) > 0 Then failure = failure & vbCrLf & cleanup
    End If
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Repair Profiling", failure)
    Exit Sub

CleanupFailed:
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Repair Profiling", failure & vbCrLf & _
        "Application state could not be restored after the failure.")
End Sub

' ==========================================================================
' THE COMMAND
' ==========================================================================
Private Function RepairProfiling() As OperationResult
    Dim cost As ProfilingPlan, risk As ProfilingPlan
    Dim costBefore As TableSnapshot, riskBefore As TableSnapshot
    Dim detail As String, failure As String
    Dim captured As Boolean, headerDrift As Boolean

    ' --- 1. THE ACCEPTED UNKEYED-DATA GATE. Data in a row with no id cannot be
    ' attributed to anything, and synchronisation would delete it without ever
    ' being able to warn. modStructuralCheck already owns that question.
    detail = modStructuralCheck.PreMutationCheck()
    If Len(detail) > 0 Then
        RepairProfiling = Refused(detail)
        Exit Function
    End If

    ' --- 2. EVERY REPORTED FAULT MUST BE ONE THIS COMMAND REPAIRS.
    If Not OnlyRepairableFaults(headerDrift, detail) Then
        RepairProfiling = Refused(detail)
        Exit Function
    End If

    ' --- 3. ASSESS BOTH GRIDS BEFORE TOUCHING EITHER. A refusal found while
    ' looking at Risk must leave Cost alone, which is only possible if every
    ' question is asked before the first answer is written.
    If Not Assess(modProfiling.CostKind(), SH_COST_PROFILING, headerDrift, cost, detail) Then
        RepairProfiling = Refused(detail)
        Exit Function
    End If
    If Not Assess(modProfiling.RiskKind(), SH_RISK_PROFILING, headerDrift, risk, detail) Then
        RepairProfiling = Refused(detail)
        Exit Function
    End If

    If Not (cost.NeedsRepair Or risk.NeedsRepair) Then
        RepairProfiling = modAppState.Succeeded( _
            SH_COST_PROFILING & " and " & SH_RISK_PROFILING & " are already " & _
            "structurally correct. Nothing was changed.")
        Exit Function
    End If

    ' --- 4. ONE TRANSACTION OVER BOTH GRIDS.
    On Error GoTo RepairFailed
    costBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(cost.Kind))
    riskBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(risk.Kind))
    captured = True

    Apply cost
    modAppState.FailPointCheck FAILPOINT_REPAIR_COST
    Apply risk
    modAppState.FailPointCheck FAILPOINT_REPAIR_RISK

    ' --- 5. THE ACCEPTED VALIDATOR GATES THE RESULT. A repair that did not
    ' produce a structurally valid workbook is a failure, and it rolls back.
    detail = modStructuralCheck.ValidateStructure()
    If Len(detail) > 0 Then
        Err.Raise vbObjectError + 5201, "modRepair.RepairProfiling", _
                  "structural revalidation failed after repair:" & vbCrLf & detail
    End If
    On Error GoTo 0

    RepairProfiling = modAppState.Succeeded(Summary(cost, risk))
    Exit Function

RepairFailed:
    failure = Err.Description
    On Error GoTo 0
    RepairProfiling = Rollback(captured, costBefore, riskBefore, failure)
End Function

' THE REPAIR ITSELF, AND IT IS TWO CALLS. Columns before rows, in the order
' Apply / Update Timeline uses them: SyncRows preserves a weight by (permanent
' id, project-year index), so the index set has to be the right one before
' ownership is re-established over it.
Private Sub Apply(ByRef plan As ProfilingPlan)
    If Not plan.NeedsRepair Then Exit Sub
    modProfiling.SetYearColumns plan.Kind, plan.StartYear, plan.TargetYears
    modProfiling.SyncRows plan.Kind
End Sub

' BOTH GRIDS BACK, OR THE FAILURE SAYS SO. The snapshot pair is the whole of
' what this command can have changed, so restoring it is the whole rollback.
Private Function Rollback(ByVal captured As Boolean, _
                          ByRef costBefore As TableSnapshot, _
                          ByRef riskBefore As TableSnapshot, _
                          ByVal failure As String) As OperationResult
    Dim note As String

    If Not captured Then
        Rollback = modAppState.Failed("Repair Profiling", failure & vbCrLf & _
            "Nothing had been modified when the failure occurred; both profiling " & _
            "grids are exactly as they were.")
        Exit Function
    End If

    On Error GoTo RestoreFailed
    modWorkbook.RestoreTable modProfiling.ProfilingTable(modProfiling.CostKind()), costBefore
    modWorkbook.RestoreTable modProfiling.ProfilingTable(modProfiling.RiskKind()), riskBefore
    On Error GoTo 0

    Rollback = modAppState.Failed("Repair Profiling", failure & vbCrLf & _
        "Both profiling grids were restored to the state they were in before the " & _
        "command ran. No weight, row or project-year column was left changed.")
    Exit Function

RestoreFailed:
    note = Err.Description
    On Error GoTo 0
    Rollback = modAppState.Failed("Repair Profiling", failure & vbCrLf & _
        "The repair failed AND the profiling grids could not be fully restored: " & _
        note & ". They require recovery before the model is used.")
End Function

Private Function Refused(ByVal detail As String) As OperationResult
    Refused = modAppState.Failed("Repair Profiling was refused. Nothing was changed.", _
                                 detail)
End Function

' ==========================================================================
' THE GATE: EVERY REPORTED FAULT MUST BE ONE THIS COMMAND REPAIRS
' ==========================================================================
' The accepted checker reports every structural fault tagged with its
' contract-declared key. This partitions those keys rather than re-deriving the
' faults, so there is no second structural rule here and a rule added later is
' BLOCKING by default: an unrecognised key refuses.
'
' WHY DUPLICATES ARE ON THE REPAIRABLE SIDE. `no_duplicate_ids` is reported for a
' repeated REGISTER id and for a repeated GRID row alike, and those are not the
' same question: a grid row repeated with identical weights is unambiguous and
' collapses, a repeated register id is not repairable at all. The key alone
' cannot tell them apart, so it is admitted here and both cases are decided
' explicitly in Assess below - which is also the only place that can see the
' WEIGHTS the distinction turns on.
' ==========================================================================
Private Function OnlyRepairableFaults(ByRef headerDrift As Boolean, _
                                      ByRef detail As String) As Boolean
    Dim report As String, blocking As String, key As String
    Dim lines() As String
    Dim index As Long

    report = modStructuralCheck.ValidateStructure()
    If Len(report) = 0 Then
        OnlyRepairableFaults = True
        Exit Function
    End If

    lines = Split(report, vbCrLf)
    For index = LBound(lines) To UBound(lines)
        key = FaultKey(lines(index))
        If Len(key) > 0 Then
            If Not IsRepairableFault(key) Then
                blocking = blocking & lines(index) & vbCrLf
            ElseIf key = CHK_PROFILING_YEAR_HEADERS Then
                ' THE HEADER RULE IS THE CHECKER'S AND STAYS THE CHECKER'S. What
                ' a project-year header must read is decided in one place; this
                ' takes the ANSWER rather than asking the question again, which
                ' is the difference between orchestrating an owner and becoming
                ' a second one. It is not attributed to a grid because it does
                ' not need to be: both are relabelled from the one applied start
                ' year, so relabelling both is correct and idempotent.
                headerDrift = True
            End If
        End If
    Next index

    If Len(blocking) = 0 Then
        OnlyRepairableFaults = True
        Exit Function
    End If

    detail = "Repair Profiling restores STRUCTURE. The workbook reports a condition " & _
             "it cannot restore without deciding what you meant:" & vbCrLf & _
             blocking & "Resolve that first. Both profiling grids are unchanged."
End Function

Private Function IsRepairableFault(ByVal key As String) As Boolean
    Select Case key
    Case CHK_PROFILING_COLUMN_COUNT, CHK_PROFILING_YEAR_HEADERS, _
         CHK_COST_PROFILING_IDS_MATCH, CHK_RISK_PROFILING_IDS_MATCH, _
         CHK_NO_DUPLICATE_IDS
        IsRepairableFault = True
    End Select
End Function

' The key out of one reported fault line, which the checker writes as
' "  [key] text". Reading the tag it declares is not parsing prose: the tags are
' the generated CHK_ constants, and an untagged line yields nothing.
Private Function FaultKey(ByVal reported As String) As String
    Dim opened As Long, closed As Long
    opened = InStr(reported, "[")
    If opened = 0 Then Exit Function
    closed = InStr(opened, reported, "]")
    If closed <= opened + 1 Then Exit Function
    FaultKey = Mid$(reported, opened + 1, closed - opened - 1)
End Function

' ==========================================================================
' ASSESSMENT - EVERY QUESTION ASKED BEFORE THE FIRST ANSWER IS WRITTEN
' ==========================================================================
Private Function Assess(ByVal kind As String, ByVal label As String, _
                        ByVal headerDrift As Boolean, ByRef plan As ProfilingPlan, _
                        ByRef detail As String) As Boolean
    Dim grid As ListObject, register As ListObject
    Dim registerIds As Object, gridIds As Object
    Dim ordered() As String, gridOrder() As String
    Dim registerCount As Long, gridCount As Long
    Dim index As Long

    plan.Kind = kind
    plan.Label = label
    Set grid = modProfiling.ProfilingTable(kind)
    Set register = modDrivers.RegisterTable(kind)

    ' --- the applied timeline decides the width, and it is read through the
    ' bounded accessor. A triple that is not readable has already been refused
    ' by OnlyRepairableFaults; this cannot then produce a nonsense target.
    plan.TargetYears = modWorkbook.ReadLongInRange(NM_APPLIED_DURATION, 1, _
                                                   LIMIT_MAX_YEAR_COLUMNS, 0)
    plan.StartYear = modWorkbook.ReadLongInRange(NM_APPLIED_START_YEAR, _
                                                  LIMIT_MIN_YEAR, LIMIT_MAX_YEAR, 0)
    If plan.TargetYears > 0 And plan.StartYear = 0 Then
        detail = "the applied timeline is not readable, so there is no width to " & _
                 "repair " & label & " to. Apply / Update Timeline on Setup first."
        Exit Function
    End If
    plan.ActualYears = modProfiling.YearColumnCount(kind)
    plan.WidthDrift = (plan.ActualYears <> plan.TargetYears)

    ' --- the register: its ids, in its order, and no id twice -------------
    If Not ReadRegisterIds(register, kind, label, registerIds, ordered, _
                           registerCount, detail) Then Exit Function

    ' --- the grid: its ids, its weights, and what the duplicates mean -----
    If Not ReadGridIds(grid, kind, label, registerIds, gridIds, gridOrder, _
                       gridCount, plan, detail) Then Exit Function

    ' --- what is missing, orphaned and out of order ----------------------
    For index = 1 To registerCount
        If Not gridIds.Exists(ordered(index)) Then plan.Missing = plan.Missing + 1
    Next index
    For index = 1 To gridCount
        If Not registerIds.Exists(gridOrder(index)) Then plan.Orphans = plan.Orphans + 1
    Next index
    ' ORDER IS ONLY A QUESTION WHEN THE TWO SETS ALREADY AGREE. A grid missing a
    ' row or carrying an orphan is being rebuilt anyway, and calling that
    ' "out of order" as well would report one fault as two.
    If plan.Missing = 0 And plan.Orphans = 0 And plan.Duplicates = 0 Then
        If gridCount = registerCount Then
            For index = 1 To registerCount
                If StrComp(ordered(index), gridOrder(index), vbTextCompare) <> 0 Then
                    plan.OrderDrift = True
                    Exit For
                End If
            Next index
        Else
            plan.OrderDrift = True
        End If
    End If

    ' --- the headers, as the checker reported them. Only worth acting on
    ' where the width is already right: a width repair relabels every column on
    ' its way past, so reporting both would report one fault as two.
    If Not plan.WidthDrift Then plan.HeaderDrift = headerDrift

    ' --- a shrink that would trim data is a refusal, never a repair -------
    If plan.TargetYears < plan.ActualYears Then
        If Not TrimIsEmpty(kind, label, plan.TargetYears, detail) Then Exit Function
    End If

    plan.NeedsRepair = plan.WidthDrift Or plan.HeaderDrift Or plan.OrderDrift _
                       Or plan.Missing > 0 Or plan.Orphans > 0 Or plan.Duplicates > 0
    Assess = True
End Function

' ==========================================================================
' THE REGISTER SIDE
' ==========================================================================
Private Function ReadRegisterIds(ByVal register As ListObject, ByVal kind As String, _
                                 ByVal label As String, ByRef ids As Object, _
                                 ByRef ordered() As String, ByRef count As Long, _
                                 ByRef detail As String) As Boolean
    Dim rows As Long, row As Long, idColumn As Long
    Dim idText As String

    Set ids = CreateObject("Scripting.Dictionary")
    idColumn = modDrivers.IdColumn(kind)
    rows = modWorkbook.BodyRowCount(register)
    ReDim ordered(1 To IIf(rows < 1, 1, rows))
    count = 0

    For row = 1 To rows
        idText = modWorkbook.TextOf(modWorkbook.CellIn(register, row, idColumn))
        If modWorkbook.IsErrorText(idText) Then
            detail = register.Name & " row " & CStr(row) & " has an identifier that " & _
                     "cannot be read. " & label & " cannot be keyed to a driver whose " & _
                     "identifier is an error value."
            Exit Function
        End If
        If Len(idText) > 0 Then
            ' A REPEATED REGISTER ID IS NOT REPAIRABLE. Two drivers claiming one
            ' identifier means one profiling row could belong to either, and
            ' nothing in the workbook says which.
            If ids.Exists(idText) Then
                detail = register.Name & " contains the identifier " & idText & " more " & _
                         "than once, so a profiling row keyed to it could belong to " & _
                         "either driver. Give each driver its own identifier first."
                Exit Function
            End If
            ids.Add idText, True
            count = count + 1
            ordered(count) = idText
        End If
    Next row
    ReadRegisterIds = True
End Function

' ==========================================================================
' THE GRID SIDE - WHERE THE WEIGHTS ARE, AND SO WHERE THE AMBIGUITY IS
' ==========================================================================
Private Function ReadGridIds(ByVal grid As ListObject, ByVal kind As String, _
                             ByVal label As String, ByRef registerIds As Object, _
                             ByRef ids As Object, ByRef ordered() As String, _
                             ByRef count As Long, ByRef plan As ProfilingPlan, _
                             ByRef detail As String) As Boolean
    Dim rows As Long, row As Long
    Dim idText As String
    Dim weights As Variant, held As Variant

    Set ids = CreateObject("Scripting.Dictionary")
    rows = modWorkbook.BodyRowCount(grid)
    ReDim ordered(1 To IIf(rows < 1, 1, rows))
    count = 0

    For row = 1 To rows
        idText = modWorkbook.TextOf(modWorkbook.CellIn(grid, row, 1))
        If modWorkbook.IsErrorText(idText) Then
            detail = label & " row " & CStr(row) & " has an identifier that cannot be " & _
                     "read, so its weights cannot be attributed to any driver."
            Exit Function
        End If
        If Len(idText) > 0 Then
            If Not ReadRowWeights(grid, kind, label, row, idText, weights, detail) Then
                Exit Function
            End If
            If ids.Exists(idText) Then
                ' A REPEATED GRID ROW IS REPAIRABLE ONLY WHEN THE TWO ROWS SAY
                ' THE SAME THING. Identical weights collapse to one row and
                ' nothing is lost; different weights are two answers to one
                ' question, and picking either would be this command deciding
                ' which of the user's two profiles they meant.
                held = ids(idText)
                If Not SameWeights(held, weights) Then
                    detail = label & " carries " & idText & " on more than one row, and " & _
                             "those rows hold different weights. Delete the row you do " & _
                             "not want, then run Repair Profiling again."
                    Exit Function
                End If
                plan.Duplicates = plan.Duplicates + 1
            Else
                ids.Add idText, weights
                count = count + 1
                ordered(count) = idText
            End If
            ' --- the semantic gate, and it applies ONLY where restructuring
            ' would move this row's weights between two different sets of
            ' project-year positions. A half-entered profile is the user's to
            ' finish; normalising it here would be this command inventing an
            ' allocation, and moving it silently would make the result look like
            ' this command's work.
            If plan.WidthDrift And registerIds.Exists(idText) Then
                If Not RecognisedProfile(weights, label, idText, detail) Then Exit Function
            End If
        End If
    Next row
    ReadGridIds = True
End Function

' One row's project-year weights, as a 0-based Variant array of Double or Empty.
' A cell that is neither blank nor a number cannot be preserved as a weight, and
' saying so here is what stops it being carried into a rebuilt grid as text.
Private Function ReadRowWeights(ByVal grid As ListObject, ByVal kind As String, _
                                ByVal label As String, ByVal row As Long, _
                                ByVal idText As String, ByRef weights As Variant, _
                                ByRef detail As String) As Boolean
    Dim fixedCols As Long, years As Long, index As Long
    Dim cell As Range
    Dim carried() As Variant

    fixedCols = modProfiling.FixedColumnCount(kind)
    years = modProfiling.YearColumnCount(kind)
    ReDim carried(0 To IIf(years < 1, 0, years - 1))

    For index = 1 To years
        Set cell = modWorkbook.CellIn(grid, row, fixedCols + index)
        If modWorkbook.IsEmptyCell(cell) Then
            carried(index - 1) = Empty
        ElseIf modWorkbook.IsErrorText(modWorkbook.TextOf(cell)) Then
            detail = label & ": the weight for " & idText & " in project year " & _
                     CStr(index) & " is an error value and cannot be preserved."
            Exit Function
        ElseIf Not IsNumeric(cell.Value) Then
            detail = label & ": the weight for " & idText & " in project year " & _
                     CStr(index) & " is not a number, so it cannot be preserved as a " & _
                     "profiling weight."
            Exit Function
        Else
            carried(index - 1) = CDbl(cell.Value)
        End If
    Next index

    weights = carried
    ReadRowWeights = True
End Function

' Two weight vectors say the same thing when every position agrees, and a blank
' does NOT agree with a zero: the accepted grid language distinguishes "not
' entered" from "nothing allocated", and collapsing them would lose that.
Private Function SameWeights(ByRef held As Variant, ByRef found As Variant) As Boolean
    Dim index As Long
    If UBound(held) <> UBound(found) Then Exit Function
    For index = LBound(held) To UBound(held)
        If IsEmpty(held(index)) <> IsEmpty(found(index)) Then Exit Function
        If Not IsEmpty(held(index)) Then
            If held(index) <> found(index) Then Exit Function
        End If
    Next index
    SameWeights = True
End Function

' ==========================================================================
' THE SEMANTIC GATE
' ==========================================================================
' TWO RECOGNISED TOTALS AND NOTHING BETWEEN THEM. 100% is a profile; 0% is the
' absence of one, which is what the contract seeds a new driver and a new project
' year with. Anything else is half-entered, and restructuring it would hand the
' user back an allocation this command had a part in without being able to say
' what part.
'
' THE ARITHMETIC IS NOT THIS MODULE'S. The sum goes through the accepted signed
' summation primitive and the comparison through the accepted subtraction, with
' the calculation contract's own tolerance - the same three the Model Check uses
' to ask the same question.
Private Function RecognisedProfile(ByRef weights As Variant, ByVal label As String, _
                                   ByVal idText As String, ByRef detail As String) As Boolean
    Dim terms() As Double
    Dim index As Long, count As Long
    Dim total As Double

    count = UBound(weights) - LBound(weights) + 1
    ReDim terms(0 To IIf(count < 1, 0, count - 1))
    For index = 0 To count - 1
        If IsEmpty(weights(LBound(weights) + index)) Then
            terms(index) = 0#
        Else
            terms(index) = CDbl(weights(LBound(weights) + index))
        End If
    Next index

    If Not modCalcFactors.SafeSignedSum(terms, count, total) Then
        detail = label & ": the weights for " & idText & " cannot be summed, so this " & _
                 "command cannot tell whether restructuring them would change what " & _
                 "they mean."
        Exit Function
    End If
    If IsRecognisedSum(total) Then
        RecognisedProfile = True
        Exit Function
    End If

    detail = label & ": the weights for " & idText & " total " & CStr(total) & _
             ", which is neither 100% nor an empty profile, and the project-year " & _
             "columns have to change. Repair Profiling restores structure and never " & _
             "chooses weights, so it will not move a half-entered profile between two " & _
             "different sets of project years. Complete or clear that row first."
End Function

Private Function IsRecognisedSum(ByVal total As Double) As Boolean
    If WithinTolerance(total, REPAIR_PROFILE_SUM_EMPTY) Then
        IsRecognisedSum = True
        Exit Function
    End If
    IsRecognisedSum = WithinTolerance(total, REPAIR_PROFILE_SUM_TARGET)
End Function

Private Function WithinTolerance(ByVal total As Double, ByVal target As Double) As Boolean
    Dim difference As Double
    If Not modCalcFactors.SafeSubtract(total, target, difference) Then Exit Function
    WithinTolerance = (Abs(difference) <= TOL_PROFILING_SUM_ABSOLUTE)
End Function

' ==========================================================================
' THE LAST STRUCTURAL QUESTION
' ==========================================================================
' A SHRINK THAT WOULD TRIM DATA IS A REFUSAL. The count and the affected
' identifiers come from the profiling owner's own destructive assessment - the
' one Apply / Update Timeline warns from - so there is no second definition here
' of what counts as data worth keeping.
Private Function TrimIsEmpty(ByVal kind As String, ByVal label As String, _
                             ByVal targetYears As Long, ByRef detail As String) As Boolean
    Dim affected() As String
    Dim affectedCount As Long, hits As Long

    hits = modProfiling.CountDataBeyond(kind, targetYears, affected, affectedCount)
    If hits = 0 Then
        TrimIsEmpty = True
        Exit Function
    End If
    detail = label & " has " & CStr(hits) & " weight(s) in project years the applied " & _
             "timeline no longer covers, on " & Named(affected, affectedCount) & _
             ". Repairing the width would delete them. Clear those cells, or apply a " & _
             "timeline that covers them, then run Repair Profiling again."
End Function

Private Function Named(ByRef ids() As String, ByVal count As Long) As String
    Dim index As Long, out As String
    For index = 1 To count
        If index > REPAIR_MAX_NAMED Then
            Named = out & " and " & CStr(count - REPAIR_MAX_NAMED) & " more"
            Exit Function
        End If
        If Len(out) > 0 Then out = out & ", "
        out = out & ids(index)
    Next index
    Named = out
End Function

' ==========================================================================
' WHAT THE USER IS TOLD
' ==========================================================================
' ONE SENTENCE PER GRID THAT NEEDED WORK, AND THE PRESERVATION CLAIM ONCE. The
' claim is worth stating because it is the question a user actually has when a
' repair command has just rewritten a grid they typed into.
Private Function Summary(ByRef cost As ProfilingPlan, ByRef risk As ProfilingPlan) As String
    Dim out As String
    out = Describe(cost)
    If Len(Describe(risk)) > 0 Then
        If Len(out) > 0 Then out = out & " "
        out = out & Describe(risk)
    End If
    Summary = out & " Every weight that could be attributed to a driver was preserved " & _
              "exactly, by permanent identifier and project year."
End Function

Private Function Describe(ByRef plan As ProfilingPlan) As String
    Dim parts As String

    If Not plan.NeedsRepair Then Exit Function
    If plan.Missing > 0 Then parts = Joined(parts, CStr(plan.Missing) & " missing row(s) restored")
    If plan.Orphans > 0 Then parts = Joined(parts, CStr(plan.Orphans) & " orphan row(s) removed")
    If plan.Duplicates > 0 Then parts = Joined(parts, CStr(plan.Duplicates) & " duplicate row(s) collapsed")
    If plan.OrderDrift Then parts = Joined(parts, "row order restored")
    If plan.WidthDrift Then
        parts = Joined(parts, "project-year columns " & CStr(plan.ActualYears) & _
                              " -> " & CStr(plan.TargetYears))
    ElseIf plan.HeaderDrift Then
        parts = Joined(parts, "project-year headers relabelled")
    End If
    Describe = plan.Label & " repaired: " & parts & "."
End Function

Private Function Joined(ByVal existing As String, ByVal part As String) As String
    If Len(existing) = 0 Then
        Joined = part
    Else
        Joined = existing & ", " & part
    End If
End Function
