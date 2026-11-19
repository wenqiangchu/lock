Attribute VB_Name = "MdlDiff"
Option Explicit

Public Const ERR_DIFF_INVALID_INPUT As Long = 1201

Public Function CompareSnapshots(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef declaredColumns As Variant, ByRef changes As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim baseRowLow As Long
    Dim baseRowHigh As Long
    Dim baseColumnLow As Long
    Dim baseColumnHigh As Long
    Dim localRowLow As Long
    Dim localRowHigh As Long
    Dim localColumnLow As Long
    Dim localColumnHigh As Long
    Dim declaredLow As Long
    Dim declaredHigh As Long
    Dim probeBound As Long
    Dim rowIndex As Long
    Dim declaredIndex As Long
    Dim priorIndex As Long
    Dim columnIndex As Long
    Dim changeCount As Long
    Dim changeIndex As Long
    Dim equal As Boolean
    Dim comparisonError As AppErrorInfo
    Dim result() As Variant

    changes = Empty
    errorInfo = MakeAppError(ERR_NONE, "MdlDiff.CompareSnapshots", "")

    If IsEmpty(declaredColumns) Then
        If IsEmpty(baseValues) And IsEmpty(localValues) Then
            CompareSnapshots = True
            Exit Function
        End If
        errorInfo = MakeAppError(ERR_DIFF_INVALID_INPUT, "MdlDiff.CompareSnapshots", "Declared columns are required for non-empty snapshots.")
        Exit Function
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

    If IsEmpty(baseValues) And IsEmpty(localValues) Then
        CompareSnapshots = True
        Exit Function
    End If
    If IsEmpty(baseValues) Xor IsEmpty(localValues) Then GoTo InvalidInput
    If Not IsArray(baseValues) Or Not IsArray(localValues) Then GoTo InvalidInput

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
    Err.Clear
    probeBound = LBound(localValues, 3)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo InvalidInput
        GoTo InvalidInput
    End If
    Err.Clear
    On Error GoTo InvalidInput

    If baseRowLow <> localRowLow Or baseRowHigh <> localRowHigh Or baseColumnLow <> localColumnLow Or baseColumnHigh <> localColumnHigh Then GoTo InvalidInput
    For declaredIndex = declaredLow To declaredHigh
        columnIndex = CLng(declaredColumns(declaredIndex))
        If columnIndex < baseColumnLow Or columnIndex > baseColumnHigh Then GoTo InvalidInput
    Next declaredIndex

    For rowIndex = baseRowLow To baseRowHigh
        For declaredIndex = declaredLow To declaredHigh
            columnIndex = CLng(declaredColumns(declaredIndex))
            If Not CompareCellValues(baseValues(rowIndex, columnIndex), localValues(rowIndex, columnIndex), equal, comparisonError) Then
                errorInfo = comparisonError
                Exit Function
            End If
            If Not equal Then changeCount = changeCount + 1
        Next declaredIndex
    Next rowIndex

    If changeCount = 0 Then
        CompareSnapshots = True
        Exit Function
    End If

    ReDim result(1 To changeCount, 1 To 4)
    For rowIndex = baseRowLow To baseRowHigh
        For declaredIndex = declaredLow To declaredHigh
            columnIndex = CLng(declaredColumns(declaredIndex))
            If Not CompareCellValues(baseValues(rowIndex, columnIndex), localValues(rowIndex, columnIndex), equal, comparisonError) Then
                errorInfo = comparisonError
                Exit Function
            End If
            If Not equal Then
                changeIndex = changeIndex + 1
                result(changeIndex, 1) = rowIndex
                result(changeIndex, 2) = columnIndex
                result(changeIndex, 3) = baseValues(rowIndex, columnIndex)
                result(changeIndex, 4) = localValues(rowIndex, columnIndex)
            End If
        Next declaredIndex
    Next rowIndex

    changes = result
    CompareSnapshots = True
    Exit Function

InvalidInput:
    changes = Empty
    errorInfo = MakeAppError(ERR_DIFF_INVALID_INPUT, "MdlDiff.CompareSnapshots", "Snapshot arrays or declared column indexes are invalid or do not match.")
End Function
