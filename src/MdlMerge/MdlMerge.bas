Attribute VB_Name = "MdlMerge"
Option Explicit

Public Enum MergeCellOutcome
    MergeTakeRemote = 1
    MergeTakeLocal = 2
    MergeSameValue = 3
    MergeNeedsChoice = 4
End Enum

Public Type MergeResult
    MergedValues As Variant
    Conflicts As Variant
    ConflictCount As Long
    PendingCount As Long
    Resolved As Boolean
End Type

Public Const ERR_MERGE_INVALID_INPUT As Long = 1301

Public Function MergeThreeWay(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef remoteValues As Variant, ByRef declaredColumns As Variant, ByRef result As MergeResult, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim baseRowLow As Long
    Dim baseRowHigh As Long
    Dim baseColumnLow As Long
    Dim baseColumnHigh As Long
    Dim localRowLow As Long
    Dim localRowHigh As Long
    Dim localColumnLow As Long
    Dim localColumnHigh As Long
    Dim remoteRowLow As Long
    Dim remoteRowHigh As Long
    Dim remoteColumnLow As Long
    Dim remoteColumnHigh As Long
    Dim declaredLow As Long
    Dim declaredHigh As Long
    Dim probeBound As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim declaredIndex As Long
    Dim priorIndex As Long
    Dim conflictCount As Long
    Dim conflictIndex As Long
    Dim outcome As MergeCellOutcome
    Dim cellError As AppErrorInfo
    Dim mergedValues() As Variant
    Dim conflicts() As Variant

    ResetMergeResult result
    errorInfo = MakeAppError(ERR_NONE, "MdlMerge.MergeThreeWay", "")

    If IsEmpty(declaredColumns) Then
        If IsEmpty(baseValues) And IsEmpty(localValues) And IsEmpty(remoteValues) Then
            result.Resolved = True
            MergeThreeWay = True
            Exit Function
        End If
        GoTo InvalidInput
    End If
    If Not IsArray(declaredColumns) Then GoTo InvalidInput

    On Error Resume Next
    declaredLow = LBound(declaredColumns, 1)
    declaredHigh = UBound(declaredColumns, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    Err.Clear
    probeBound = LBound(declaredColumns, 2)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    Err.Clear
    On Error GoTo InvalidInput

    For declaredIndex = declaredLow To declaredHigh
        If Not IsNumeric(declaredColumns(declaredIndex)) Then GoTo InvalidInput
        columnIndex = CLng(declaredColumns(declaredIndex))
        If CDbl(declaredColumns(declaredIndex)) <> CDbl(columnIndex) Or columnIndex < 1 Then GoTo InvalidInput
        For priorIndex = declaredLow To declaredIndex - 1
            If CLng(declaredColumns(priorIndex)) = columnIndex Then GoTo InvalidInput
        Next priorIndex
    Next declaredIndex

    If IsEmpty(baseValues) And IsEmpty(localValues) And IsEmpty(remoteValues) Then
        result.Resolved = True
        MergeThreeWay = True
        Exit Function
    End If
    If IsEmpty(baseValues) Or IsEmpty(localValues) Or IsEmpty(remoteValues) Then GoTo InvalidInput
    If Not IsArray(baseValues) Or Not IsArray(localValues) Or Not IsArray(remoteValues) Then GoTo InvalidInput

    On Error Resume Next
    baseRowLow = LBound(baseValues, 1)
    baseRowHigh = UBound(baseValues, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    baseColumnLow = LBound(baseValues, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    baseColumnHigh = UBound(baseValues, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    Err.Clear
    probeBound = LBound(baseValues, 3)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    Err.Clear

    localRowLow = LBound(localValues, 1)
    localRowHigh = UBound(localValues, 1)
    localColumnLow = LBound(localValues, 2)
    localColumnHigh = UBound(localValues, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    remoteRowLow = LBound(remoteValues, 1)
    remoteRowHigh = UBound(remoteValues, 1)
    remoteColumnLow = LBound(remoteValues, 2)
    remoteColumnHigh = UBound(remoteValues, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    Err.Clear
    probeBound = LBound(localValues, 3)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    Err.Clear
    probeBound = LBound(remoteValues, 3)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    Err.Clear
    On Error GoTo InvalidInput

    If baseRowLow <> localRowLow Or baseRowLow <> remoteRowLow Or baseRowHigh <> localRowHigh Or baseRowHigh <> remoteRowHigh Then GoTo InvalidInput
    If baseColumnLow <> localColumnLow Or baseColumnLow <> remoteColumnLow Or baseColumnHigh <> localColumnHigh Or baseColumnHigh <> remoteColumnHigh Then GoTo InvalidInput
    For declaredIndex = declaredLow To declaredHigh
        columnIndex = CLng(declaredColumns(declaredIndex))
        If columnIndex < baseColumnLow Or columnIndex > baseColumnHigh Then GoTo InvalidInput
    Next declaredIndex

    ReDim mergedValues(baseRowLow To baseRowHigh, baseColumnLow To baseColumnHigh)
    For rowIndex = baseRowLow To baseRowHigh
        For columnIndex = baseColumnLow To baseColumnHigh
            mergedValues(rowIndex, columnIndex) = remoteValues(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex

    For rowIndex = baseRowLow To baseRowHigh
        For declaredIndex = declaredLow To declaredHigh
            columnIndex = CLng(declaredColumns(declaredIndex))
            outcome = ClassifyCell(baseValues(rowIndex, columnIndex), localValues(rowIndex, columnIndex), remoteValues(rowIndex, columnIndex), cellError)
            If outcome = 0 Then
                ResetMergeResult result
                errorInfo = cellError
                Exit Function
            End If
            If outcome = MergeNeedsChoice Then conflictCount = conflictCount + 1
        Next declaredIndex
    Next rowIndex

    If conflictCount > 0 Then ReDim conflicts(1 To conflictCount, 1 To 5)
    For rowIndex = baseRowLow To baseRowHigh
        For declaredIndex = declaredLow To declaredHigh
            columnIndex = CLng(declaredColumns(declaredIndex))
            outcome = ClassifyCell(baseValues(rowIndex, columnIndex), localValues(rowIndex, columnIndex), remoteValues(rowIndex, columnIndex), cellError)
            If outcome = 0 Then
                ResetMergeResult result
                errorInfo = cellError
                Exit Function
            End If
            Select Case outcome
                Case MergeTakeLocal, MergeSameValue
                    mergedValues(rowIndex, columnIndex) = localValues(rowIndex, columnIndex)
                Case MergeNeedsChoice
                    conflictIndex = conflictIndex + 1
                    conflicts(conflictIndex, 1) = rowIndex
                    conflicts(conflictIndex, 2) = columnIndex
                    conflicts(conflictIndex, 3) = baseValues(rowIndex, columnIndex)
                    conflicts(conflictIndex, 4) = localValues(rowIndex, columnIndex)
                    conflicts(conflictIndex, 5) = remoteValues(rowIndex, columnIndex)
            End Select
        Next declaredIndex
    Next rowIndex

    result.MergedValues = mergedValues
    result.ConflictCount = conflictCount
    result.PendingCount = conflictCount
    result.Resolved = (conflictCount = 0)
    If conflictCount > 0 Then
        result.Conflicts = conflicts
    Else
        result.Conflicts = Empty
    End If
    MergeThreeWay = True
    Exit Function

InvalidInput:
    ResetMergeResult result
    errorInfo = MakeAppError(ERR_MERGE_INVALID_INPUT, "MdlMerge.MergeThreeWay", "Snapshot arrays or declared column indexes are invalid or do not match.")
End Function

Public Function ApplyConflictChoices(ByRef result As MergeResult, ByRef choices As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim conflictLow As Long
    Dim conflictHigh As Long
    Dim conflictColumnLow As Long
    Dim conflictColumnHigh As Long
    Dim choiceLow As Long
    Dim choiceHigh As Long
    Dim choiceColumnLow As Long
    Dim choiceColumnHigh As Long
    Dim probeBound As Long
    Dim conflictIndex As Long
    Dim choiceIndex As Long
    Dim matchIndex As Long
    Dim matchCount As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim selection As String
    Dim selections() As Variant

    errorInfo = MakeAppError(ERR_NONE, "MdlMerge.ApplyConflictChoices", "")
    If result.Resolved Then
        If result.ConflictCount = 0 And IsEmpty(choices) Then
            ApplyConflictChoices = True
        Else
            errorInfo = MakeAppError(ERR_MERGE_INVALID_INPUT, "MdlMerge.ApplyConflictChoices", "Conflict choices have already been applied or are not expected.")
        End If
        Exit Function
    End If
    If result.ConflictCount < 1 Or result.PendingCount <> result.ConflictCount Then GoTo InvalidChoices
    If Not IsArray(result.Conflicts) Or Not IsArray(result.MergedValues) Or Not IsArray(choices) Then GoTo InvalidChoices

    On Error Resume Next
    conflictLow = LBound(result.Conflicts, 1)
    conflictHigh = UBound(result.Conflicts, 1)
    conflictColumnLow = LBound(result.Conflicts, 2)
    conflictColumnHigh = UBound(result.Conflicts, 2)
    choiceLow = LBound(choices, 1)
    choiceHigh = UBound(choices, 1)
    choiceColumnLow = LBound(choices, 2)
    choiceColumnHigh = UBound(choices, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo InvalidChoices
        GoTo InvalidChoices
    End If
    Err.Clear
    probeBound = LBound(result.Conflicts, 3)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo InvalidChoices
        GoTo InvalidChoices
    End If
    Err.Clear
    probeBound = LBound(choices, 3)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo InvalidChoices
        GoTo InvalidChoices
    End If
    Err.Clear
    On Error GoTo InvalidChoices

    If conflictHigh - conflictLow + 1 <> result.ConflictCount Or conflictColumnHigh - conflictColumnLow + 1 <> 5 Then GoTo InvalidChoices
    If choiceHigh - choiceLow + 1 <> result.ConflictCount Or choiceColumnHigh - choiceColumnLow + 1 <> 3 Then GoTo InvalidChoices
    ReDim selections(1 To result.ConflictCount)

    For choiceIndex = choiceLow To choiceHigh
        If Not IsNumeric(choices(choiceIndex, choiceColumnLow)) Or Not IsNumeric(choices(choiceIndex, choiceColumnLow + 1)) Then GoTo InvalidChoices
        rowIndex = CLng(choices(choiceIndex, choiceColumnLow))
        columnIndex = CLng(choices(choiceIndex, choiceColumnLow + 1))
        If CDbl(choices(choiceIndex, choiceColumnLow)) <> CDbl(rowIndex) Or CDbl(choices(choiceIndex, choiceColumnLow + 1)) <> CDbl(columnIndex) Then GoTo InvalidChoices
        If VarType(choices(choiceIndex, choiceColumnLow + 2)) <> vbString Then GoTo InvalidChoices
        selection = UCase$(Trim$(CStr(choices(choiceIndex, choiceColumnLow + 2))))
        If selection <> "LOCAL" And selection <> "REMOTE" Then GoTo InvalidChoices

        matchCount = 0
        For conflictIndex = conflictLow To conflictHigh
            If rowIndex = CLng(result.Conflicts(conflictIndex, conflictColumnLow)) And columnIndex = CLng(result.Conflicts(conflictIndex, conflictColumnLow + 1)) Then
                matchCount = matchCount + 1
                matchIndex = conflictIndex - conflictLow + 1
            End If
        Next conflictIndex
        If matchCount <> 1 Then GoTo InvalidChoices
        If Not IsEmpty(selections(matchIndex)) Then GoTo InvalidChoices
        selections(matchIndex) = selection
    Next choiceIndex

    For conflictIndex = 1 To result.ConflictCount
        If IsEmpty(selections(conflictIndex)) Then GoTo InvalidChoices
    Next conflictIndex

    For conflictIndex = conflictLow To conflictHigh
        rowIndex = CLng(result.Conflicts(conflictIndex, conflictColumnLow))
        columnIndex = CLng(result.Conflicts(conflictIndex, conflictColumnLow + 1))
        selection = CStr(selections(conflictIndex - conflictLow + 1))
        If selection = "LOCAL" Then
            result.MergedValues(rowIndex, columnIndex) = result.Conflicts(conflictIndex, conflictColumnLow + 3)
        Else
            result.MergedValues(rowIndex, columnIndex) = result.Conflicts(conflictIndex, conflictColumnLow + 4)
        End If
    Next conflictIndex
    result.PendingCount = 0
    result.Resolved = True
    ApplyConflictChoices = True
    Exit Function

InvalidChoices:
    errorInfo = MakeAppError(ERR_MERGE_INVALID_INPUT, "MdlMerge.ApplyConflictChoices", "Conflict choices must cover every conflict exactly once with a Local or Remote selection.")
End Function

Public Function IsMergeResolved(ByRef result As MergeResult) As Boolean
    IsMergeResolved = result.Resolved And result.PendingCount = 0
End Function

Private Function ClassifyCell(ByVal baseValue As Variant, ByVal localValue As Variant, ByVal remoteValue As Variant, ByRef errorInfo As AppErrorInfo) As MergeCellOutcome
    Dim localMatchesBase As Boolean
    Dim remoteMatchesBase As Boolean
    Dim localMatchesRemote As Boolean

    If Not CompareCellValues(baseValue, localValue, localMatchesBase, errorInfo) Then Exit Function
    If Not CompareCellValues(baseValue, remoteValue, remoteMatchesBase, errorInfo) Then Exit Function
    If Not CompareCellValues(localValue, remoteValue, localMatchesRemote, errorInfo) Then Exit Function
    If localMatchesBase Then
        ClassifyCell = MergeTakeRemote
        Exit Function
    End If
    If remoteMatchesBase Then
        ClassifyCell = MergeTakeLocal
        Exit Function
    End If
    If localMatchesRemote Then
        ClassifyCell = MergeSameValue
    Else
        ClassifyCell = MergeNeedsChoice
    End If
End Function

Private Sub ResetMergeResult(ByRef result As MergeResult)
    result.MergedValues = Empty
    result.Conflicts = Empty
    result.ConflictCount = 0
    result.PendingCount = 0
    result.Resolved = False
End Sub
