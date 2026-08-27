Attribute VB_Name = "modSimReport"
Option Explicit

' ==========================================================================
' modSimReport - the Phase-6 workbook orchestration layer, and nothing else.
'
' This module owns the ENDPOINT, the settled read accessors, the two simulation
' control reads, the AUTO nonce lifecycle, the order in which the accepted pure
' kernels are called, dual-bank publication into `_SimData`, and the attempt and
' status bookkeeping.
'
' IT OWNS NO MATHEMATICS. Not one distribution, not one uniform, not one mean,
' not one quantile, not one contingency subtraction, not one canonical field and
' not one hash step happens here. Every number it publishes was produced by the
' module the accepted architecture gave it to:
'
'     modCalcReport      the resolved model, the CURRENT analytical fingerprint,
'                        the deterministic base estimate A, the applied timeline
'     modSimEngine       the iteration totals
'     modSimStats        the moments, the ladder and every contingency
'     modSimFingerprint  the request fingerprint and the result digest
'
' --------------------------------------------------------------------------
' WHY THE PUBLICATION HAS TWO BANKS
' --------------------------------------------------------------------------
' A candidate success is written ENTIRELY into the bank that is not published,
' verified there, and published by one small final write that moves the active
' bank. A failure at any earlier point leaves the published bank physically
' untouched and the half-written candidate with no semantic standing at all,
' because `active_bank` still names the other one.
'
' That is the only design under which "the prior successful publication
' survives" is structurally true. Writing over the published rows and stamping
' the Run ID afterwards cannot survive a COM failure half way through a million
' rows, and no rollback of a million rows is a transaction anybody should
' attempt.
'
' --------------------------------------------------------------------------
' WHAT THIS MODULE NEVER TOUCHES
' --------------------------------------------------------------------------
' Results. Not one cell. The Stage-A presentation formulas already read the
' active bank, so publication is `_SimData` and only `_SimData`; a second
' written transaction is exactly the failure mode where the distribution
' commits and the sheet that shows it does not.
'
' Selected Confidence Level. It is a reporting selector: it is not read here, it
' is not part of the request identity, it allocates nothing and it decides no
' status.
'
' NOTHING IN THIS MODULE HAS BEEN EXECUTED. It is source, submitted for review.
' Behaviour on real Excel is proven at Gate B.
' ==========================================================================

' The Gate-B failpoints. Public because a later harness arms them BY NAME.
Public Const FAILPOINT_SIM_AFTER_NONCE As String = "Phase6AfterNoncePersisted"
Public Const FAILPOINT_SIM_CANDIDATE_BANK As String = "Phase6CandidateBank"
Public Const FAILPOINT_SIM_FINAL_COMMIT As String = "Phase6FinalCommit"

' Iteration rows per bulk write. A PERFORMANCE detail and nothing more: no
' contract names it, and changing it changes no published value. What it must
' not become is one COM call per iteration.
Private Const SIM_WRITE_CHUNK As Long = 16384

' Everything one invocation stages, before any cell is written. It is Private on
' purpose: a Public result type would be a caller-writable trust boundary, and
' Step 7 already paid for that lesson once.
Private Type SimRunPackage
    Drivers() As DriverFactors
    DriverCount As Long
    AnalyticalFingerprint As String
    BaseNominal As Double
    BasePv As Double
    AppliedTimeline As String
    DecimalSeparator As String

    Iterations As Long
    SeedMode As String
    HasSuppliedSeed As Boolean
    SuppliedSeed As Long
    EffectiveSeed As Long
    ConsumedNonce As Long
    NonceConsumed As Boolean

    CandidateRunId As Long
    ActiveBank As String
    TargetBank As String

    TotalNominal() As Double
    TotalPv() As Double

    NominalSummary As SimStatsMeasure
    PvSummary As SimStatsMeasure
    NominalLabels() As String
    NominalLadder() As Double
    PvLabels() As String
    PvLadder() As Double
    NominalContingency() As Double
    PvContingency() As Double

    RequestFingerprint As String
    ResultDigest As String
    Stamp As Date
End Type

' ==========================================================================
' THE SEVEN PHASE-6 PROCEDURES
' ==========================================================================
Public Sub PCCM_RunSimulation()
    ' THE WHOLE INVOCATION IS INSIDE AN ENVELOPE, installed before the first
    ' fallible operation, exactly as the accepted PCCM_Calculate is. The two
    ' Boolean facts carry the cleanup distinction rather than statement
    ' position, and cleanup is attempted at most once.
    Dim state As AppStateSnapshot, result As OperationResult
    Dim stateCaptured As Boolean, cleanupAttempted As Boolean, committed As Boolean
    Dim cleanup As String, failure As String

    On Error GoTo InvocationFailed
    state = modAppState.CaptureAppState()
    stateCaptured = True
    modAppState.BeginOperation
    result = RunSimulation(committed)
    On Error GoTo 0

    On Error GoTo NormalCleanupFailed
    cleanupAttempted = True
    cleanup = modAppState.FinishOperation(state)
    On Error GoTo 0
    stateCaptured = False

    If Len(cleanup) > 0 Then result = CleanupOutcome(result, committed, cleanup)
    modAppState.Announce result
    Exit Sub

NormalCleanupFailed:
    failure = Err.Description
    On Error GoTo 0
    modAppState.Announce CleanupOutcome(result, committed, _
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
    modAppState.Announce modAppState.Failed("Run Simulation", failure)
    Exit Sub

CleanupFailed:
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Run Simulation", failure & vbCrLf & _
        "Application state could not be restored after the failure.")
End Sub

Private Function CleanupOutcome(ByRef result As OperationResult, ByVal committed As Boolean, _
                                ByVal cleanup As String) As OperationResult
    ' A CLEANUP PROBLEM AFTER THE COMMIT IS AN INVOCATION FAILURE, NOT A FAILED
    ' SIMULATION. Once the active bank has moved, the workbook says SUCCESS and
    ' that is committed truth; nothing here rewrites it, unpublishes a bank or
    ' rolls anything back.
    If committed Then
        CleanupOutcome = modAppState.Failed("Run Simulation", _
            "The simulation COMMITTED successfully and the published bank records " & _
            "it. Application state could not be fully restored afterwards:" & _
            vbCrLf & cleanup)
        Exit Function
    End If
    If result.Ok Then
        CleanupOutcome = modAppState.Failed("Run Simulation", cleanup)
    Else
        CleanupOutcome = modAppState.Failed("Run Simulation", result.Detail & vbCrLf & cleanup)
    End If
End Function

Public Function PCCM_SimulationStatus() As String
    ' Re-evaluates the derived status and writes ONLY the two derived rows. It
    ' touches no bank, no counter and no attempt field: asking for the status is
    ' not a simulation.
    Dim status As String
    status = DeriveSimStatus()
    WriteStatusBlock status
    PCCM_SimulationStatus = status
End Function

Public Function PCCM_SimulationRequestFingerprint() As String
    ' THE STORED one, from the ACTIVE bank. It never recomputes: this is what
    ' the last successful run was asked, and a recomputation here would answer a
    ' different question and hide every staleness.
    PCCM_SimulationRequestFingerprint = _
        ActiveSnapshotText(SIM_IDENTITY_ROW_REQUEST_FINGERPRINT)
End Function

Public Function PCCM_SimulationResultDigest() As String
    ' The STORED digest of the active bank. Never recomputed, and never rebuilt
    ' from the published iteration rows.
    PCCM_SimulationResultDigest = ActiveSnapshotText(SIM_IDENTITY_ROW_RESULT_DIGEST)
End Function

Public Function PCCM_CurrentSimulationRequestFingerprint() As String
    ' The RECOMPUTED one, through the same side-effect-free path the status
    ' derivation uses. Blank when the current request cannot be formed.
    Dim fingerprint As String, detail As String
    If Not CurrentRequestFingerprint(fingerprint, detail) Then Exit Function
    PCCM_CurrentSimulationRequestFingerprint = fingerprint
End Function

Public Function PCCM_SimulationAttemptResult() As String
    PCCM_SimulationAttemptResult = SharedText(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT)
End Function

Public Function PCCM_SimulationAttemptDetail() As String
    PCCM_SimulationAttemptDetail = SharedText(SIM_IDENTITY_ROW_LAST_ATTEMPT_DETAIL)
End Function

' ==========================================================================
' The run, in the accepted Step-11A order
' ==========================================================================
Private Function RunSimulation(ByRef committed As Boolean) As OperationResult
    Dim package As SimRunPackage
    Dim detail As String

    ' 1-3. Preparation, the two controls, and every PRE-ALLOCATION prerequisite.
    '      Nothing random has happened yet and no sequence has been spent.
    If Not PrepareRun(package, detail) Then
        RunSimulation = RecordRefusal(package, detail)
        Exit Function
    End If

    ' 4. AUTO only: derive, PERSIST and VERIFY the nonce before sampling. After
    '    this point the sequence is spent whatever happens.
    If Not AllocateAutoNonce(package, detail) Then
        RunSimulation = RecordRefusal(package, detail)
        Exit Function
    End If
    modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE

    ' 5-11. The accepted kernels, in order, entirely in memory.
    If Not RunKernels(package, detail) Then
        RunSimulation = RecordRefusal(package, detail)
        Exit Function
    End If

    ' 12-18. Stage the success, choose the INACTIVE bank, write it and verify it.
    '        The candidate failpoint is NOT here: injecting before any candidate
    '        write only proves that nothing was written. It lives inside
    '        PublishCandidate, after the inactive bank has been written, so
    '        Gate B can prove a partially written candidate has no standing.
    package.Stamp = Now
    package.TargetBank = InactiveBank(package.ActiveBank)
    If Not PublishCandidate(package, detail) Then
        RunSimulation = RecordFailure(package, detail)
        Exit Function
    End If

    ' 19. The one final write. The active bank moves last, inside it.
    If Not FinalCommit(package, detail) Then
        RunSimulation = RecordFailure(package, detail)
        Exit Function
    End If

    committed = True
    RunSimulation = modAppState.Succeeded( _
        "Simulation complete. " & CStr(package.Iterations) & _
        " iterations published to bank " & package.TargetBank & ".")
End Function

Private Function PrepareRun(ByRef package As SimRunPackage, ByRef detail As String) As Boolean
    ' 1. The ONE Phase-5 bridge. It requires a CURRENT calculation, writes
    '    nothing, and hands back the CURRENT analytical fingerprint - never a
    '    stored last-successful one.
    If Not modCalcReport.CalcPrepareSimulationInputs( _
            package.Drivers, package.DriverCount, package.AnalyticalFingerprint, _
            package.BaseNominal, package.BasePv, package.AppliedTimeline, _
            package.DecimalSeparator, detail) Then
        Exit Function
    End If

    ' 2. The two simulation controls, strictly.
    If Not ResolveIterations(package.Iterations, detail) Then Exit Function
    If Not ResolveSeed(package, detail) Then Exit Function

    ' 3. Machine prerequisites, all of them, BEFORE anything is allocated.
    If Not ValidatePreAllocation(package, detail) Then Exit Function

    PrepareRun = True
End Function

Private Function ResolveIterations(ByRef iterations As Long, ByRef detail As String) As Boolean
    ' A GENUINE NUMBER. Not blank, not the text "10000", not a Boolean, not an
    ' error and not a fraction: a control that accepted numeric-looking text
    ' would let a workbook run a simulation nobody could reproduce.
    Dim raw As Variant, value As Double

    If Not modWorkbook.NameExists(NM_INPUT_MONTE_CARLO_ITERATIONS) Then
        detail = "simulation: the iteration count input is missing"
        Exit Function
    End If
    raw = modWorkbook.ReadValue(NM_INPUT_MONTE_CARLO_ITERATIONS)
    If Not modWorkbook.IsWholeInRange(raw, CDbl(SIM_MIN_ITERATIONS), _
                                      CDbl(SIM_MAX_ITERATIONS)) Then
        detail = "simulation: Monte Carlo Iterations must be a whole number between " & _
                 CStr(SIM_MIN_ITERATIONS) & " and " & CStr(SIM_MAX_ITERATIONS)
        Exit Function
    End If
    If Not modWorkbook.TryReadDouble(raw, value) Then
        detail = "simulation: Monte Carlo Iterations is not a usable number"
        Exit Function
    End If
    ' NARROWED ONLY AFTER THE BOUNDS ARE PROVEN.
    iterations = modWorkbook.SafeLong(value)
    ResolveIterations = True
End Function

Private Function ResolveSeed(ByRef package As SimRunPackage, ByRef detail As String) As Boolean
    ' A BLANK Random Seed IS the AUTO request. There is no separate seed-mode
    ' cell and no zero sentinel: absence is the mode.
    Dim raw As Variant, value As Double

    If Not modWorkbook.NameExists(NM_INPUT_RANDOM_SEED) Then
        detail = "simulation: the random seed input is missing"
        Exit Function
    End If
    If modWorkbook.IsEmptyCell(modWorkbook.NamedCell(NM_INPUT_RANDOM_SEED)) Then
        package.SeedMode = SIM_SEED_MODE_AUTO
        package.HasSuppliedSeed = False
        ResolveSeed = True
        Exit Function
    End If

    raw = modWorkbook.ReadValue(NM_INPUT_RANDOM_SEED)
    If Not modWorkbook.IsWholeInRange(raw, CDbl(SIM_SEED_MIN), CDbl(SIM_SEED_MAX)) Then
        detail = "simulation: Random Seed must be blank for AUTO, or a whole number " & _
                 "between " & CStr(SIM_SEED_MIN) & " and " & CStr(SIM_SEED_MAX)
        Exit Function
    End If
    If Not modWorkbook.TryReadDouble(raw, value) Then
        detail = "simulation: Random Seed is not a usable number"
        Exit Function
    End If
    package.SeedMode = SIM_SEED_MODE_FIXED
    package.HasSuppliedSeed = True
    package.SuppliedSeed = modWorkbook.SafeLong(value)
    ResolveSeed = True
End Function

Private Function ValidatePreAllocation(ByRef package As SimRunPackage, _
                                       ByRef detail As String) As Boolean
    ' EVERYTHING THAT CAN REFUSE, REFUSES HERE - before a random sequence is
    ' spent. A run that could never be identified or committed must not consume
    ' one.
    Dim lastRunId As Long, nonce As Long

    If Not ReadActiveBank(package.ActiveBank, detail) Then Exit Function
    If Not ReadMachineLong(SIM_IDENTITY_ROW_LAST_RUN_ID, CDbl(SIM_RUN_ID_INITIAL), _
                           CDbl(SIM_RUN_ID_MAXIMUM), lastRunId, detail) Then
        detail = "simulation: the run identity counter is not readable - " & detail
        Exit Function
    End If
    If lastRunId >= SIM_RUN_ID_MAXIMUM Then
        detail = "simulation: the run identity counter is exhausted; no further run " & _
                 "can be identified or committed"
        Exit Function
    End If
    package.CandidateRunId = lastRunId + 1

    If Len(InactiveBank(package.ActiveBank)) = 0 Then
        detail = "simulation: the target publication bank cannot be determined"
        Exit Function
    End If

    If Not package.HasSuppliedSeed Then
        If Not ReadMachineLong(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE, _
                               CDbl(SIM_NONCE_FIRST_VALID), CDbl(SIM_NONCE_EXHAUSTED), _
                               nonce, detail) Then
            detail = "simulation: the AUTO nonce counter is not readable - " & detail
            Exit Function
        End If
        If nonce >= SIM_NONCE_EXHAUSTED Then
            detail = "simulation: the AUTO nonce sequence is exhausted"
            Exit Function
        End If
        package.ConsumedNonce = nonce
    End If

    ValidatePreAllocation = True
End Function

Private Function AllocateAutoNonce(ByRef package As SimRunPackage, _
                                   ByRef detail As String) As Boolean
    ' PERSIST BEFORE SAMPLING, AND VERIFY THE PERSISTENCE.
    '
    ' If the advance cannot be written and read back, sampling must not begin:
    ' a run that drew from a sequence the workbook does not know it spent could
    ' be reproduced by a later AUTO run, which is the one thing the nonce exists
    ' to prevent. After this returns True the nonce is spent, and no later
    ' refusal, failure or restoration rolls it back.
    Dim seed As Long, stored As Long, probe As String

    If package.HasSuppliedSeed Then
        package.EffectiveSeed = package.SuppliedSeed
        AllocateAutoNonce = True
        Exit Function
    End If

    If Not modSimRng.SimRngAutoSeedFromNonce(package.ConsumedNonce, seed, detail) Then
        Exit Function
    End If
    package.EffectiveSeed = seed

    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = package.ConsumedNonce + 1
    If Not ReadMachineLong(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE, _
                           CDbl(SIM_NONCE_FIRST_VALID), CDbl(SIM_NONCE_EXHAUSTED) + 1#, _
                           stored, probe) Then
        detail = "simulation: the AUTO nonce advance could not be verified"
        Exit Function
    End If
    If stored <> package.ConsumedNonce + 1 Then
        detail = "simulation: the AUTO nonce advance did not persist"
        Exit Function
    End If
    package.NonceConsumed = True
    AllocateAutoNonce = True
End Function

Private Function RunKernels(ByRef package As SimRunPackage, ByRef detail As String) As Boolean
    ' 5. The engine, once.
    If Not modSimEngine.SimEngineRun(package.Drivers, package.DriverCount, _
                                     package.EffectiveSeed, package.Iterations, _
                                     package.TotalNominal, package.TotalPv, detail) Then
        Exit Function
    End If

    ' 6-7. The statistics, over the SAME retained arrays that will be published.
    If Not modSimStats.SimStatsDescribe(package.TotalNominal, package.Iterations, _
                                        package.NominalSummary, package.NominalLabels, _
                                        package.NominalLadder, detail) Then
        Exit Function
    End If
    If Not modSimStats.SimStatsDescribe(package.TotalPv, package.Iterations, _
                                        package.PvSummary, package.PvLabels, _
                                        package.PvLadder, detail) Then
        Exit Function
    End If

    ' 8. The two ladders must be the SAME owner ladder. They come from one
    '    projection, so a disagreement means something moved between the calls -
    '    and a summary whose two columns are different ladders is unreadable.
    If Not SameLadder(package, detail) Then Exit Function

    ' 9. EVERY rung, both measures, through the accepted primitive. Not the
    '    selected one: Selected CL may move without a rerun, and a publication
    '    holding one rung would force a rerun or a worksheet subtraction.
    If Not BuildContingencies(package, detail) Then Exit Function

    ' 10. The request fingerprint, continuing the CURRENT analytical hash state.
    If Not modSimFingerprint.SimFpBuildRequestFingerprint( _
            package.AnalyticalFingerprint, package.Iterations, package.SeedMode, _
            package.HasSuppliedSeed, package.SuppliedSeed, _
            package.RequestFingerprint, detail) Then
        Exit Function
    End If

    ' 11. The digest, over the exact arrays about to be published.
    If Not modSimFingerprint.SimFpResultDigest(package.TotalNominal, package.TotalPv, _
                                               package.Iterations, _
                                               package.DecimalSeparator, _
                                               package.ResultDigest, detail) Then
        Exit Function
    End If

    RunKernels = True
End Function

Private Function SameLadder(ByRef package As SimRunPackage, ByRef detail As String) As Boolean
    ' Provenance, checked rather than assumed. Both ladders came from
    ' SimStatsDescribe in this invocation; nothing has assigned to either since,
    ' and this proves they are still the accepted projected ladder in the
    ' accepted order.
    Dim index As Long, label As String

    If package.NominalSummary.QuantileCount <> SIM_QUANTILE_COUNT Then
        detail = "simulation: the nominal ladder is not the accepted length"
        Exit Function
    End If
    If package.PvSummary.QuantileCount <> SIM_QUANTILE_COUNT Then
        detail = "simulation: the PV ladder is not the accepted length"
        Exit Function
    End If
    For index = 0 To SIM_QUANTILE_COUNT - 1
        If Not LadderLabelAt(index, label, detail) Then Exit Function
        If StrComp(package.NominalLabels(LBound(package.NominalLabels) + index), _
                   label, vbBinaryCompare) <> 0 Then
            detail = "simulation: the nominal ladder is not the accepted projection"
            Exit Function
        End If
        If StrComp(package.PvLabels(LBound(package.PvLabels) + index), _
                   label, vbBinaryCompare) <> 0 Then
            detail = "simulation: the PV ladder is not the accepted projection"
            Exit Function
        End If
    Next index
    SameLadder = True
End Function

Private Function LadderLabelAt(ByVal position As Long, ByRef label As String, _
                               ByRef detail As String) As Boolean
    Select Case position
        Case 0
            label = SIM_QUANTILE_1
        Case 1
            label = SIM_QUANTILE_2
        Case 2
            label = SIM_QUANTILE_3
        Case 3
            label = SIM_QUANTILE_4
        Case 4
            label = SIM_QUANTILE_5
        Case 5
            label = SIM_QUANTILE_6
        Case 6
            label = SIM_QUANTILE_7
        Case 7
            label = SIM_QUANTILE_8
        Case 8
            label = SIM_QUANTILE_9
        Case 9
            label = SIM_QUANTILE_10
        Case 10
            label = SIM_QUANTILE_11
        Case Else
            detail = "simulation: the ladder has no rung at that position"
            Exit Function
    End Select
    LadderLabelAt = True
End Function

Private Function BuildContingencies(ByRef package As SimRunPackage, _
                                    ByRef detail As String) As Boolean
    Dim index As Long, value As Double

    ReDim package.NominalContingency(0 To SIM_QUANTILE_COUNT - 1)
    ReDim package.PvContingency(0 To SIM_QUANTILE_COUNT - 1)
    For index = 0 To SIM_QUANTILE_COUNT - 1
        If Not modSimStats.SimStatsContingency( _
                package.NominalLadder(LBound(package.NominalLadder) + index), _
                package.BaseNominal, value, detail) Then
            Exit Function
        End If
        package.NominalContingency(index) = value
        If Not modSimStats.SimStatsContingency( _
                package.PvLadder(LBound(package.PvLadder) + index), _
                package.BasePv, value, detail) Then
            Exit Function
        End If
        package.PvContingency(index) = value
    Next index
    BuildContingencies = True
End Function

' ==========================================================================
' Candidate publication - THE INACTIVE BANK ONLY
' ==========================================================================
Private Function PublishCandidate(ByRef package As SimRunPackage, _
                                  ByRef detail As String) As Boolean
    ' THE CANDIDATE TRANSACTION, INSIDE A SCOPED ERROR ENVELOPE.
    '
    ' A Range assignment, a chunk write and a verification read are COM calls,
    ' and COM calls raise. Without a handler here a raised write left the run
    ' through PCCM_RunSimulation's invocation handler, which reports honestly
    ' but writes NO attempt record - so an infrastructure failure after the AUTO
    ' nonce was consumed left the attempt axis silent, against the accepted
    ' `refusal_or_failure_after_auto_allocation.attempt_metadata_updated: true`.
    '
    ' The envelope converts every such failure into a plain False, so
    ' RunSimulation's existing `If Not PublishCandidate ... RecordFailure` path
    ' owns it and the attempt block records FAILED with the seed mode, the
    ' effective seed and the consumed nonce.
    '
    ' NOTHING IS ROLLED BACK. The active bank is not touched, the run id is not
    ' touched, and the half-written INACTIVE bank is left exactly as it is: it
    ' has no semantic standing precisely because the selector still names the
    ' other bank. Scoped, not `On Error Resume Next` - one handler, one exit.
    Dim snapshot As Variant, summary As Variant, contingency As Variant
    Dim failure As String

    On Error GoTo CandidateFailed

    BuildSnapshotBlock package, snapshot
    BuildSummaryBlock package, summary
    BuildContingencyBlock package, contingency

    SimSheet.Range(SnapshotRange(package.TargetBank)).Value2 = snapshot
    SimSheet.Range(SummaryRange(package.TargetBank)).Value2 = summary
    SimSheet.Range(ContingencyRange(package.TargetBank)).Value2 = contingency
    If Not WriteIterationBank(package, detail) Then Exit Function

    ' THE CANDIDATE BANK IS NOW WRITTEN AND NOT YET VERIFIED. This is the
    ' runtime boundary Gate B needs: the inactive bank holds candidate data, the
    ' active bank has not moved, and the run must still end as FAILED.
    modAppState.FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK

    If Not VerifyCandidateBank(package, snapshot, summary, contingency, detail) Then
        Exit Function
    End If

    On Error GoTo 0
    PublishCandidate = True
    Exit Function

CandidateFailed:
    failure = Err.Description
    On Error GoTo 0
    detail = "simulation: inactive-bank publication failed in bank " & _
             package.TargetBank & ": " & failure & _
             ". The candidate bank may be partially written and it has " & _
             "no semantic standing; the previously published bank remains " & _
             "authoritative."
End Function

Private Sub BuildSnapshotBlock(ByRef package As SimRunPackage, ByRef block As Variant)
    ' Rows 8..20 of the target bank, in the accepted order. ONE block, built
    ' once, so verification compares against the exact values that were written.
    Dim built(1 To 13, 1 To 1) As Variant
    built(1, 1) = package.Stamp
    built(2, 1) = package.CandidateRunId
    built(3, 1) = package.RequestFingerprint
    built(4, 1) = package.ResultDigest
    built(5, 1) = package.SeedMode
    If package.HasSuppliedSeed Then
        built(6, 1) = package.SuppliedSeed
    Else
        built(6, 1) = vbNullString
    End If
    built(7, 1) = package.EffectiveSeed
    If package.NonceConsumed Then
        built(8, 1) = package.ConsumedNonce
    Else
        built(8, 1) = vbNullString
    End If
    built(9, 1) = package.Iterations
    built(10, 1) = SIM_RNG_VERSION
    built(11, 1) = SIM_METHOD_VERSION
    built(12, 1) = SIM_MODEL_VERSION
    built(13, 1) = package.AppliedTimeline
    block = built
End Sub

Private Sub BuildSummaryBlock(ByRef package As SimRunPackage, ByRef block As Variant)
    ' Sixteen rows, two measures. Every value came from SimStatsDescribe except
    ' the deterministic base, which came from the Phase-5 bridge.
    Dim built(1 To 16, 1 To 2) As Variant
    Dim index As Long
    built(1, 1) = package.NominalSummary.Mean
    built(1, 2) = package.PvSummary.Mean
    built(2, 1) = package.NominalSummary.SampleStandardDeviation
    built(2, 2) = package.PvSummary.SampleStandardDeviation
    built(3, 1) = package.NominalSummary.Minimum
    built(3, 2) = package.PvSummary.Minimum
    For index = 0 To SIM_QUANTILE_COUNT - 1
        built(4 + index, 1) = package.NominalLadder(LBound(package.NominalLadder) + index)
        built(4 + index, 2) = package.PvLadder(LBound(package.PvLadder) + index)
    Next index
    built(15, 1) = package.NominalSummary.Maximum
    built(15, 2) = package.PvSummary.Maximum
    built(16, 1) = package.BaseNominal
    built(16, 2) = package.BasePv
    block = built
End Sub

Private Sub BuildContingencyBlock(ByRef package As SimRunPackage, ByRef block As Variant)
    Dim built(1 To 11, 1 To 2) As Variant
    Dim index As Long
    For index = 0 To SIM_QUANTILE_COUNT - 1
        built(1 + index, 1) = package.NominalContingency(index)
        built(1 + index, 2) = package.PvContingency(index)
    Next index
    block = built
End Sub

Private Function WriteIterationBank(ByRef package As SimRunPackage, _
                                    ByRef detail As String) As Boolean
    ' CHUNKED BULK WRITES. At the hundred-thousand design target a per-cell loop
    ' would be a hundred thousand COM calls; the chunk size is a performance
    ' detail and no contract names it.
    '
    ' THE INDEX IS LOGICAL. Step 8 retains iteration i at element i - 1, so
    ' element `offset` is iteration SIM_DIGEST_INDEX_ORIGIN + offset whatever the
    ' carrier's physical LBound happens to be.
    Dim block As Variant
    Dim offset As Long, rows As Long, index As Long

    offset = 0
    Do While offset < package.Iterations
        rows = package.Iterations - offset
        If rows > SIM_WRITE_CHUNK Then rows = SIM_WRITE_CHUNK
        ReDim block(1 To rows, 1 To 3)
        For index = 0 To rows - 1
            block(index + 1, 1) = SIM_DIGEST_INDEX_ORIGIN + offset + index
            block(index + 1, 2) = _
                package.TotalNominal(LBound(package.TotalNominal) + offset + index)
            block(index + 1, 3) = _
                package.TotalPv(LBound(package.TotalPv) + offset + index)
        Next index
        SimSheet.Range(IterationRange(package.TargetBank, offset, rows)).Value2 = block
        offset = offset + rows
    Loop
    WriteIterationBank = True
End Function

Private Function VerifyCandidateBank(ByRef package As SimRunPackage, _
                                     ByRef snapshot As Variant, ByRef summary As Variant, _
                                     ByRef contingency As Variant, _
                                     ByRef detail As String) As Boolean
    ' READS ONLY THE INACTIVE BANK. No tolerance, no recomputation, and no
    ' statistic or digest reconstructed from worksheet data - the question is
    ' "did this land", not "is it still right".
    Dim written As Variant
    Dim offset As Long, rows As Long, index As Long

    If Not SameBlock(SnapshotRange(package.TargetBank), snapshot, 13, 1) Then
        detail = "simulation: the candidate snapshot did not verify in bank " & _
                 package.TargetBank
        Exit Function
    End If
    If Not SameBlock(SummaryRange(package.TargetBank), summary, 16, 2) Then
        detail = "simulation: the candidate summary did not verify in bank " & _
                 package.TargetBank
        Exit Function
    End If
    If Not SameBlock(ContingencyRange(package.TargetBank), contingency, 11, 2) Then
        detail = "simulation: the candidate contingency ladder did not verify in bank " & _
                 package.TargetBank
        Exit Function
    End If

    offset = 0
    Do While offset < package.Iterations
        rows = package.Iterations - offset
        If rows > SIM_WRITE_CHUNK Then rows = SIM_WRITE_CHUNK
        written = SimSheet.Range(IterationRange(package.TargetBank, offset, rows)).Value2
        For index = 1 To rows
            If Not SameCell(written(index, 1), _
                            SIM_DIGEST_INDEX_ORIGIN + offset + index - 1) Then
                detail = "simulation: the candidate iteration index did not verify"
                Exit Function
            End If
            If Not SameCell(written(index, 2), package.TotalNominal( _
                    LBound(package.TotalNominal) + offset + index - 1)) Then
                detail = "simulation: a candidate nominal total did not verify"
                Exit Function
            End If
            If Not SameCell(written(index, 3), package.TotalPv( _
                    LBound(package.TotalPv) + offset + index - 1)) Then
                detail = "simulation: a candidate PV total did not verify"
                Exit Function
            End If
        Next index
        offset = offset + rows
    Loop
    VerifyCandidateBank = True
End Function

' ==========================================================================
' The one final write
' ==========================================================================
Private Function FinalCommit(ByRef package As SimRunPackage, ByRef detail As String) As Boolean
    ' D22:D30 IN ONE ASSIGNMENT, with the active bank as its last field. Not
    ' nine writes that could half-succeed and publish a bank with no run id, or
    ' a run id pointing at a bank that was never activated.
    '
    ' THREE FAILURE CLASSES, AND ONLY ONE OF THEM SKIPS THE RESTORE:
    '
    '   A. the capture itself fails - nothing has been written, so there is
    '      nothing to put back and no captured block to put back WITH.
    '   B. anything after the assignment was attempted - a raised COM write, an
    '      injected failpoint, a raised verification read, or a plain
    '      verification mismatch. All four restore.
    '   C. the restore itself fails - said plainly, never glossed.
    '
    ' B IS THE ONE THAT MATTERS. The source must not assume "an exception means
    ' Excel wrote nothing": a raised assignment leaves the range in an unknown
    ' state, which is exactly why the prior block was captured first.
    Dim previous As Variant, block As Variant
    Dim cause As String, failure As String

    ' A. CAPTURE FIRST, under its own handler, before anything is written.
    On Error GoTo CaptureFailed
    previous = SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2
    On Error GoTo 0

    BuildCommitBlock package, block

    ' B. THE COMMIT. From the assignment onward every exit restores.
    On Error GoTo CommitFailed
    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = block
    modAppState.FailPointCheck FAILPOINT_SIM_FINAL_COMMIT
    If SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1) Then
        On Error GoTo 0
        FinalCommit = True
        Exit Function
    End If
    On Error GoTo 0
    cause = "the committed block did not verify"
    GoTo RestorePrevious

CommitFailed:
    cause = Err.Description
    On Error GoTo 0

RestorePrevious:
    ' ONE WRITE BACK, THEN VERIFY IT. Put the captured block back, so the run id
    ' and the published bank are exactly what they were, and the half-written
    ' candidate bank stays unpublished with no semantic standing.
    On Error GoTo RestoreFailed
    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous
    If SameBlock(SIM_FINAL_COMMIT_RANGE, previous, 9, 1) Then
        On Error GoTo 0
        detail = "simulation: the final commit did not complete (" & cause & _
                 "); the previous published bank has been restored and " & _
                 "remains authoritative"
        Exit Function
    End If
    On Error GoTo 0
    detail = "simulation: the final commit did not complete (" & cause & _
             ") AND the previous shared block could not be restored. The " & _
             "publication selector cannot be guaranteed and requires recovery."
    Exit Function

RestoreFailed:
    failure = Err.Description
    On Error GoTo 0
    detail = "simulation: the final commit did not complete (" & cause & _
             ") AND the previous shared block could not be restored: " & failure & _
             ". The publication selector cannot be guaranteed and requires recovery."
    Exit Function

CaptureFailed:
    ' NO CANDIDATE WRITE WAS ATTEMPTED, and there is no captured block, so the
    ' restore path is NOT entered - entering it would write an unset Variant
    ' over a publication this run never touched.
    failure = Err.Description
    On Error GoTo 0
    detail = "simulation: the previous shared commit block could not be read, so " & _
             "no final commit was attempted: " & failure & _
             ". The published bank is unchanged and remains authoritative."
End Function

Private Sub BuildCommitBlock(ByRef package As SimRunPackage, ByRef block As Variant)
    Dim built(1 To 9, 1 To 1) As Variant
    built(1, 1) = package.CandidateRunId
    built(2, 1) = SIM_ATTEMPT_SUCCESS
    built(3, 1) = vbNullString
    built(4, 1) = package.SeedMode
    built(5, 1) = package.EffectiveSeed
    If package.NonceConsumed Then
        built(6, 1) = package.ConsumedNonce
    Else
        built(6, 1) = vbNullString
    End If
    built(7, 1) = SIM_STATE_CURRENT
    ' THE SAME captured moment as the snapshot's own stamp.
    built(8, 1) = package.Stamp
    built(9, 1) = package.TargetBank
    block = built
End Sub

' ==========================================================================
' Attempt bookkeeping - never the run id, never the active bank
' ==========================================================================
Private Function RecordRefusal(ByRef package As SimRunPackage, _
                               ByVal detail As String) As OperationResult
    WriteAttemptBlock package, SIM_ATTEMPT_REFUSED, detail
    RecordRefusal = modAppState.Failed("Run Simulation", detail)
End Function

Private Function RecordFailure(ByRef package As SimRunPackage, _
                               ByVal detail As String) As OperationResult
    WriteAttemptBlock package, SIM_ATTEMPT_FAILED, detail
    RecordFailure = modAppState.Failed("Run Simulation", detail)
End Function

Private Sub WriteAttemptBlock(ByRef package As SimRunPackage, ByVal result As String, _
                              ByVal detail As String)
    ' D23:D29 ONLY. The run identity counter and the publication selector are
    ' not in this range and are never written by an unsuccessful attempt.
    '
    ' THE STATUS IS DERIVED, not inherited from the attempt. A failed run over a
    ' request that still matches the published bank leaves the distribution
    ' CURRENT, and saying otherwise would be reporting the attempt as the state.
    Dim block(1 To 7, 1 To 1) As Variant
    block(1, 1) = result
    block(2, 1) = detail
    If Len(package.SeedMode) > 0 Then
        block(3, 1) = package.SeedMode
        If package.HasSuppliedSeed Or package.NonceConsumed Then
            block(4, 1) = package.EffectiveSeed
        Else
            block(4, 1) = vbNullString
        End If
    Else
        block(3, 1) = vbNullString
        block(4, 1) = vbNullString
    End If
    If package.NonceConsumed Then
        block(5, 1) = package.ConsumedNonce
    Else
        block(5, 1) = vbNullString
    End If
    block(6, 1) = DeriveSimStatus()
    block(7, 1) = Now
    SimSheet.Range(AttemptRange()).Value2 = block
End Sub

Private Sub WriteStatusBlock(ByVal status As String)
    Dim block(1 To 2, 1 To 1) As Variant
    block(1, 1) = status
    block(2, 1) = Now
    SimSheet.Range(StatusRange()).Value2 = block
End Sub

' ==========================================================================
' The derived status, and the side-effect-free current request
' ==========================================================================
Private Function DeriveSimStatus() As String
    ' THE CURRENT REQUEST AND THE ACTIVE BANK DECIDE. The attempt history never
    ' does, and neither does the reporting selector.
    Dim fingerprint As String, detail As String, active As String, stored As String

    If Not CurrentRequestFingerprint(fingerprint, detail) Then
        DeriveSimStatus = SIM_STATE_INVALID
        Exit Function
    End If
    If Not ReadActiveBank(active, detail) Then
        DeriveSimStatus = SIM_STATE_INVALID
        Exit Function
    End If
    ' A BLANK ACTIVE BANK IS THE ABSENCE OF A PUBLICATION, not a state. Two
    ' blanks are not a match; reporting CURRENT from them would claim a
    ' simulation that never ran.
    If Len(active) = 0 Then Exit Function
    stored = ActiveSnapshotText(SIM_IDENTITY_ROW_REQUEST_FINGERPRINT)
    If Len(stored) = 0 Then
        DeriveSimStatus = SIM_STATE_INVALID
        Exit Function
    End If
    If StrComp(fingerprint, stored, vbBinaryCompare) = 0 Then
        DeriveSimStatus = SIM_STATE_CURRENT
    Else
        DeriveSimStatus = SIM_STATE_STALE
    End If
End Function

Private Function CurrentRequestFingerprint(ByRef fingerprint As String, _
                                           ByRef detail As String) As Boolean
    ' SIDE-EFFECT FREE. It reads no counter, allocates no nonce, runs no engine,
    ' computes no statistic and writes nothing. The operational counters are not
    ' request identity: an exhausted counter prevents another RUN, and must not
    ' make an already-matching publication stale merely by being asked about.
    Dim package As SimRunPackage

    If Not modCalcReport.CalcPrepareSimulationInputs( _
            package.Drivers, package.DriverCount, package.AnalyticalFingerprint, _
            package.BaseNominal, package.BasePv, package.AppliedTimeline, _
            package.DecimalSeparator, detail) Then
        Exit Function
    End If
    If Not ResolveIterations(package.Iterations, detail) Then Exit Function
    If Not ResolveSeed(package, detail) Then Exit Function
    If Not modSimFingerprint.SimFpBuildRequestFingerprint( _
            package.AnalyticalFingerprint, package.Iterations, package.SeedMode, _
            package.HasSuppliedSeed, package.SuppliedSeed, fingerprint, detail) Then
        Exit Function
    End If
    CurrentRequestFingerprint = True
End Function

' ==========================================================================
' Machine-state access
' ==========================================================================
Private Function ReadActiveBank(ByRef bank As String, ByRef detail As String) As Boolean
    Dim value As String
    value = SharedText(SIM_IDENTITY_ROW_ACTIVE_BANK)
    If Len(value) = 0 Then
        bank = vbNullString
        ReadActiveBank = True
        Exit Function
    End If
    If StrComp(value, SIM_BANK_A, vbBinaryCompare) = 0 Then
        bank = SIM_BANK_A
        ReadActiveBank = True
        Exit Function
    End If
    If StrComp(value, SIM_BANK_B, vbBinaryCompare) = 0 Then
        bank = SIM_BANK_B
        ReadActiveBank = True
        Exit Function
    End If
    detail = "simulation: the publication selector holds a value that is not a bank"
End Function

Private Function InactiveBank(ByVal active As String) As String
    ' The first success targets A; every success after that targets whichever
    ' bank is not published.
    If Len(active) = 0 Then
        InactiveBank = SIM_BANK_A
        Exit Function
    End If
    If StrComp(active, SIM_BANK_A, vbBinaryCompare) = 0 Then
        InactiveBank = SIM_BANK_B
        Exit Function
    End If
    If StrComp(active, SIM_BANK_B, vbBinaryCompare) = 0 Then
        InactiveBank = SIM_BANK_A
    End If
End Function

Private Function ActiveSnapshotText(ByVal row As Long) As String
    Dim active As String, detail As String
    If Not ReadActiveBank(active, detail) Then Exit Function
    If Len(active) = 0 Then Exit Function
    ActiveSnapshotText = modWorkbook.TextOf( _
        SimSheet.Range(SnapshotColumn(active) & CStr(row)))
End Function

Private Function ReadMachineLong(ByVal row As Long, ByVal minValue As Double, _
                                 ByVal maxValue As Double, ByRef value As Long, _
                                 ByRef detail As String) As Boolean
    Dim raw As Variant, number As Double
    raw = SharedCell(row).Value2
    If Not modWorkbook.IsWholeInRange(raw, minValue, maxValue) Then
        detail = "the stored value is not a whole number in its accepted range"
        Exit Function
    End If
    If Not modWorkbook.TryReadDouble(raw, number) Then
        detail = "the stored value is not a usable number"
        Exit Function
    End If
    value = modWorkbook.SafeLong(number)
    ReadMachineLong = True
End Function

Private Function SameBlock(ByVal address As String, ByRef block As Variant, _
                           ByVal rows As Long, ByVal columns As Long) As Boolean
    Dim written As Variant, row As Long, column As Long
    written = SimSheet.Range(address).Value2
    For row = 1 To rows
        For column = 1 To columns
            If Not SameCell(WrittenCell(written, row, column, columns), _
                            block(row, column)) Then
                Exit Function
            End If
        Next column
    Next row
    SameBlock = True
End Function

Private Function WrittenCell(ByRef written As Variant, ByVal row As Long, _
                             ByVal column As Long, ByVal columns As Long) As Variant
    ' A single-cell Range returns a scalar rather than an array.
    If IsArray(written) Then
        WrittenCell = written(row, column)
    Else
        WrittenCell = written
    End If
End Function

Private Function SameCell(ByVal written As Variant, ByVal wanted As Variant) As Boolean
    If IsEmpty(written) Then
        SameCell = (VarType(wanted) = vbString And Len(CStr(wanted)) = 0)
        Exit Function
    End If
    If VarType(wanted) = vbString Then
        SameCell = (StrComp(CStr(written), CStr(wanted), vbBinaryCompare) = 0)
        Exit Function
    End If
    If Not IsNumeric(written) Then Exit Function
    SameCell = (CDbl(written) = CDbl(wanted))
End Function

Private Function SimSheet() As Worksheet
    Set SimSheet = modWorkbook.Sh(SIM_DATA_SHEET)
End Function

Private Function SharedCell(ByVal row As Long) As Range
    Set SharedCell = SimSheet.Range(SIM_SHARED_VALUE_COLUMN & CStr(row))
End Function

Private Function SharedText(ByVal row As Long) As String
    SharedText = modWorkbook.TextOf(SharedCell(row))
End Function

Private Function SnapshotColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        SnapshotColumn = SIM_SNAPSHOT_COLUMN_A
    Else
        SnapshotColumn = SIM_SNAPSHOT_COLUMN_B
    End If
End Function

Private Function SnapshotRange(ByVal bank As String) As String
    SnapshotRange = SnapshotColumn(bank) & CStr(SIM_IDENTITY_ROW_LAST_SUCCESSFUL_STAMP) & _
                    ":" & SnapshotColumn(bank) & CStr(SIM_IDENTITY_ROW_APPLIED_TIMELINE)
End Function

Private Function SummaryRange(ByVal bank As String) As String
    Dim first As String, last As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        first = SIM_SUMMARY_A_NOMINAL_COLUMN
        last = SIM_SUMMARY_A_PV_COLUMN
    Else
        first = SIM_SUMMARY_B_NOMINAL_COLUMN
        last = SIM_SUMMARY_B_PV_COLUMN
    End If
    SummaryRange = first & CStr(SIM_SUMMARY_FIRST_ROW) & ":" & _
                   last & CStr(SIM_SUMMARY_LAST_ROW)
End Function

Private Function ContingencyRange(ByVal bank As String) As String
    Dim first As String, last As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        first = SIM_CONTINGENCY_A_NOMINAL_COLUMN
        last = SIM_CONTINGENCY_A_PV_COLUMN
    Else
        first = SIM_CONTINGENCY_B_NOMINAL_COLUMN
        last = SIM_CONTINGENCY_B_PV_COLUMN
    End If
    ContingencyRange = first & CStr(SIM_CONTINGENCY_FIRST_ROW) & ":" & _
                       last & CStr(SIM_CONTINGENCY_LAST_ROW)
End Function

Private Function IterationRange(ByVal bank As String, ByVal offset As Long, _
                                ByVal rows As Long) As String
    Dim first As String, last As String, top As Long
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        first = SIM_ITER_A_ITERATION_INDEX_COLUMN
        last = SIM_ITER_A_TOTAL_PV_COLUMN
    Else
        first = SIM_ITER_B_ITERATION_INDEX_COLUMN
        last = SIM_ITER_B_TOTAL_PV_COLUMN
    End If
    top = SIM_DATA_FIRST_ITERATION_ROW + offset
    IterationRange = first & CStr(top) & ":" & last & CStr(top + rows - 1)
End Function

Private Function AttemptRange() As String
    AttemptRange = SIM_SHARED_VALUE_COLUMN & CStr(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT) & _
                   ":" & SIM_SHARED_VALUE_COLUMN & CStr(SIM_IDENTITY_ROW_STATUS_EVALUATED_AT)
End Function

Private Function StatusRange() As String
    StatusRange = SIM_SHARED_VALUE_COLUMN & CStr(SIM_IDENTITY_ROW_SIMULATION_STATUS) & _
                  ":" & SIM_SHARED_VALUE_COLUMN & CStr(SIM_IDENTITY_ROW_STATUS_EVALUATED_AT)
End Function
