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
    MergedFonts As Variant
    MergedFills As Variant
    Conflicts As Variant
    ConflictCount As Long
    PendingCount As Long
    Resolved As Boolean
End Type

Public Const ERR_MERGE_INVALID_INPUT As Long = 1301

Public Function MergeThreeWay(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef remoteValues As Variant, ByRef baseFonts As Variant, ByRef localFonts As Variant, ByRef remoteFonts As Variant, ByRef baseFills As Variant, ByRef localFills As Variant, ByRef remoteFills As Variant, ByRef declaredColumns As Variant, ByRef result As MergeResult, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim declaredLow As Long
    Dim declaredHigh As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim declaredIndex As Long
    Dim priorIndex As Long
    Dim conflictCount As Long
    Dim conflictIndex As Long
    Dim outcome As MergeCellOutcome
    Dim cellError As AppErrorInfo
    Dim mergedValues() As Variant
    Dim mergedFonts() As Variant
    Dim mergedFills() As Variant
    Dim conflicts() As Variant

    ResetMergeResult result
    errorInfo = MakeAppError(ERR_NONE, "MdlMerge.MergeThreeWay", "")
    If Not ValidateDeclaredColumns(declaredColumns, declaredLow, declaredHigh) Then GoTo InvalidInput
    If IsEmpty(baseValues) And IsEmpty(localValues) And IsEmpty(remoteValues) And IsEmpty(baseFonts) And IsEmpty(localFonts) And IsEmpty(remoteFonts) And IsEmpty(baseFills) And IsEmpty(localFills) And IsEmpty(remoteFills) Then
        result.Resolved = True
        MergeThreeWay = True
        Exit Function
    End If
    If IsEmpty(baseValues) Or IsEmpty(localValues) Or IsEmpty(remoteValues) Or IsEmpty(baseFonts) Or IsEmpty(localFonts) Or IsEmpty(remoteFonts) Or IsEmpty(baseFills) Or IsEmpty(localFills) Or IsEmpty(remoteFills) Then GoTo InvalidInput
    If Not GetMatrixBounds(baseValues, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not MatrixBoundsMatch(localValues, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not MatrixBoundsMatch(remoteValues, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not MatrixBoundsMatch(baseFonts, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not MatrixBoundsMatch(localFonts, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not MatrixBoundsMatch(remoteFonts, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not MatrixBoundsMatch(baseFills, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not MatrixBoundsMatch(localFills, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not MatrixBoundsMatch(remoteFills, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput

    For declaredIndex = declaredLow To declaredHigh
        If Not IsNumeric(declaredColumns(declaredIndex)) Then GoTo InvalidInput
        columnIndex = CLng(declaredColumns(declaredIndex))
        If CDbl(declaredColumns(declaredIndex)) <> CDbl(columnIndex) Or columnIndex < columnLow Or columnIndex > columnHigh Then GoTo InvalidInput
        For priorIndex = declaredLow To declaredIndex - 1
            If CLng(declaredColumns(priorIndex)) = columnIndex Then GoTo InvalidInput
        Next priorIndex
    Next declaredIndex

    ReDim mergedValues(rowLow To rowHigh, columnLow To columnHigh)
    ReDim mergedFonts(rowLow To rowHigh, columnLow To columnHigh)
    ReDim mergedFills(rowLow To rowHigh, columnLow To columnHigh)
    For rowIndex = rowLow To rowHigh
        For columnIndex = columnLow To columnHigh
            mergedValues(rowIndex, columnIndex) = remoteValues(rowIndex, columnIndex)
            mergedFonts(rowIndex, columnIndex) = remoteFonts(rowIndex, columnIndex)
            mergedFills(rowIndex, columnIndex) = remoteFills(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex

    For rowIndex = rowLow To rowHigh
        For declaredIndex = declaredLow To declaredHigh
            columnIndex = CLng(declaredColumns(declaredIndex))
            If Not ValidateCellTriplet(baseValues(rowIndex, columnIndex), localValues(rowIndex, columnIndex), remoteValues(rowIndex, columnIndex), baseFonts(rowIndex, columnIndex), localFonts(rowIndex, columnIndex), remoteFonts(rowIndex, columnIndex), baseFills(rowIndex, columnIndex), localFills(rowIndex, columnIndex), remoteFills(rowIndex, columnIndex), cellError) Then
                errorInfo = cellError
                Exit Function
            End If
            outcome = ClassifyComponent(baseFonts(rowIndex, columnIndex), localFonts(rowIndex, columnIndex), remoteFonts(rowIndex, columnIndex), cellError)
            If outcome = 0 Then
                errorInfo = cellError
                Exit Function
            End If
            If outcome = MergeNeedsChoice Then conflictCount = conflictCount + 1
            outcome = ClassifyComponent(baseFills(rowIndex, columnIndex), localFills(rowIndex, columnIndex), remoteFills(rowIndex, columnIndex), cellError)
            If outcome = 0 Then
                errorInfo = cellError
                Exit Function
            End If
            If outcome = MergeNeedsChoice Then conflictCount = conflictCount + 1
        Next declaredIndex
    Next rowIndex

    If conflictCount > 0 Then ReDim conflicts(1 To conflictCount, 1 To 6)
    For rowIndex = rowLow To rowHigh
        For declaredIndex = declaredLow To declaredHigh
            columnIndex = CLng(declaredColumns(declaredIndex))
            outcome = ClassifyComponent(baseFonts(rowIndex, columnIndex), localFonts(rowIndex, columnIndex), remoteFonts(rowIndex, columnIndex), cellError)
            Select Case outcome
                Case MergeTakeLocal, MergeSameValue
                    mergedFonts(rowIndex, columnIndex) = localFonts(rowIndex, columnIndex)
                    If Not ValueFromTextFont(localFonts(rowIndex, columnIndex), mergedValues(rowIndex, columnIndex), cellError) Then
                        errorInfo = cellError
                        Exit Function
                    End If
                Case MergeNeedsChoice
                    conflictIndex = conflictIndex + 1
                    SetConflict conflicts, conflictIndex, rowIndex, columnIndex, "TextFont", baseFonts(rowIndex, columnIndex), localFonts(rowIndex, columnIndex), remoteFonts(rowIndex, columnIndex)
            End Select

            outcome = ClassifyComponent(baseFills(rowIndex, columnIndex), localFills(rowIndex, columnIndex), remoteFills(rowIndex, columnIndex), cellError)
            Select Case outcome
                Case MergeTakeLocal, MergeSameValue
                    mergedFills(rowIndex, columnIndex) = localFills(rowIndex, columnIndex)
                Case MergeNeedsChoice
                    conflictIndex = conflictIndex + 1
                    SetConflict conflicts, conflictIndex, rowIndex, columnIndex, "Fill", baseFills(rowIndex, columnIndex), localFills(rowIndex, columnIndex), remoteFills(rowIndex, columnIndex)
            End Select
        Next declaredIndex
    Next rowIndex

    result.MergedValues = mergedValues
    result.MergedFonts = mergedFonts
    result.MergedFills = mergedFills
    result.ConflictCount = conflictCount
    result.PendingCount = conflictCount
    result.Resolved = (conflictCount = 0)
    If conflictCount > 0 Then result.Conflicts = conflicts
    MergeThreeWay = True
    Exit Function

InvalidInput:
    ResetMergeResult result
    errorInfo = MakeAppError(ERR_MERGE_INVALID_INPUT, "MdlMerge.MergeThreeWay", "Snapshot arrays, formatting records, or declared column indexes are invalid or do not match.")
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
    Dim mergedRowLow As Long
    Dim mergedRowHigh As Long
    Dim mergedColumnLow As Long
    Dim mergedColumnHigh As Long
    Dim conflictIndex As Long
    Dim choiceIndex As Long
    Dim matchIndex As Long
    Dim matchCount As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim componentName As String
    Dim selection As String
    Dim selectedComponent As Variant
    Dim selectedValue As Variant
    Dim selections() As Variant
    Dim mergedValues As Variant
    Dim mergedFonts As Variant
    Dim mergedFills As Variant

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
    If Not IsArray(result.Conflicts) Or Not IsArray(result.MergedValues) Or Not IsArray(result.MergedFonts) Or Not IsArray(result.MergedFills) Or Not IsArray(choices) Then GoTo InvalidChoices
    If Not GetMatrixBounds(result.Conflicts, conflictLow, conflictHigh, conflictColumnLow, conflictColumnHigh) Then GoTo InvalidChoices
    If Not GetMatrixBounds(choices, choiceLow, choiceHigh, choiceColumnLow, choiceColumnHigh) Then GoTo InvalidChoices
    If conflictHigh - conflictLow + 1 <> result.ConflictCount Or conflictColumnHigh - conflictColumnLow + 1 <> 6 Then GoTo InvalidChoices
    If choiceHigh - choiceLow + 1 <> result.ConflictCount Or choiceColumnHigh - choiceColumnLow + 1 <> 4 Then GoTo InvalidChoices
    If Not GetMatrixBounds(result.MergedValues, mergedRowLow, mergedRowHigh, mergedColumnLow, mergedColumnHigh) Then GoTo InvalidChoices
    If Not MatrixBoundsMatch(result.MergedFonts, mergedRowLow, mergedRowHigh, mergedColumnLow, mergedColumnHigh) Then GoTo InvalidChoices
    If Not MatrixBoundsMatch(result.MergedFills, mergedRowLow, mergedRowHigh, mergedColumnLow, mergedColumnHigh) Then GoTo InvalidChoices
    If Not ValidateConflictRecords(result.Conflicts, conflictLow, conflictHigh, conflictColumnLow, mergedRowLow, mergedRowHigh, mergedColumnLow, mergedColumnHigh, errorInfo) Then GoTo InvalidChoices
    ReDim selections(1 To result.ConflictCount)

    For choiceIndex = choiceLow To choiceHigh
        If Not TryReadLong(choices(choiceIndex, choiceColumnLow), rowIndex) Then GoTo InvalidChoices
        If Not TryReadLong(choices(choiceIndex, choiceColumnLow + 1), columnIndex) Then GoTo InvalidChoices
        If VarType(choices(choiceIndex, choiceColumnLow + 2)) <> vbString Or VarType(choices(choiceIndex, choiceColumnLow + 3)) <> vbString Then GoTo InvalidChoices
        componentName = CStr(choices(choiceIndex, choiceColumnLow + 2))
        selection = UCase$(Trim$(CStr(choices(choiceIndex, choiceColumnLow + 3))))
        If componentName <> "TextFont" And componentName <> "Fill" Then GoTo InvalidChoices
        If selection <> "LOCAL" And selection <> "REMOTE" Then GoTo InvalidChoices

        matchCount = 0
        For conflictIndex = conflictLow To conflictHigh
            If rowIndex = CLng(result.Conflicts(conflictIndex, conflictColumnLow)) And columnIndex = CLng(result.Conflicts(conflictIndex, conflictColumnLow + 1)) And componentName = CStr(result.Conflicts(conflictIndex, conflictColumnLow + 2)) Then
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

    mergedValues = result.MergedValues
    mergedFonts = result.MergedFonts
    mergedFills = result.MergedFills
    For conflictIndex = conflictLow To conflictHigh
        rowIndex = CLng(result.Conflicts(conflictIndex, conflictColumnLow))
        columnIndex = CLng(result.Conflicts(conflictIndex, conflictColumnLow + 1))
        componentName = CStr(result.Conflicts(conflictIndex, conflictColumnLow + 2))
        selection = CStr(selections(conflictIndex - conflictLow + 1))
        If selection = "LOCAL" Then
            selectedComponent = result.Conflicts(conflictIndex, conflictColumnLow + 4)
        Else
            selectedComponent = result.Conflicts(conflictIndex, conflictColumnLow + 5)
        End If
        If componentName = "TextFont" Then
            mergedFonts(rowIndex, columnIndex) = selectedComponent
            If Not ValueFromTextFont(mergedFonts(rowIndex, columnIndex), selectedValue, errorInfo) Then Exit Function
            mergedValues(rowIndex, columnIndex) = selectedValue
        Else
            mergedFills(rowIndex, columnIndex) = selectedComponent
        End If
    Next conflictIndex
    result.MergedValues = mergedValues
    result.MergedFonts = mergedFonts
    result.MergedFills = mergedFills
    result.PendingCount = 0
    result.Resolved = True
    ApplyConflictChoices = True
    Exit Function

InvalidChoices:
    errorInfo = MakeAppError(ERR_MERGE_INVALID_INPUT, "MdlMerge.ApplyConflictChoices", "Conflict choices must cover every component conflict exactly once with a Local or Remote selection.")
End Function

Private Function ValidateConflictRecords(ByRef conflicts As Variant, ByVal conflictLow As Long, ByVal conflictHigh As Long, ByVal conflictColumnLow As Long, ByVal rowLow As Long, ByVal rowHigh As Long, ByVal columnLow As Long, ByVal columnHigh As Long, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim conflictIndex As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim componentName As String

    For conflictIndex = conflictLow To conflictHigh
        If Not TryReadLong(conflicts(conflictIndex, conflictColumnLow), rowIndex) Then Exit Function
        If Not TryReadLong(conflicts(conflictIndex, conflictColumnLow + 1), columnIndex) Then Exit Function
        If rowIndex < rowLow Or rowIndex > rowHigh Or columnIndex < columnLow Or columnIndex > columnHigh Then Exit Function
        If VarType(conflicts(conflictIndex, conflictColumnLow + 2)) <> vbString Then Exit Function
        componentName = CStr(conflicts(conflictIndex, conflictColumnLow + 2))
        If Not ValidateConflictComponent(componentName, conflicts(conflictIndex, conflictColumnLow + 3), conflicts(conflictIndex, conflictColumnLow + 4), conflicts(conflictIndex, conflictColumnLow + 5), errorInfo) Then Exit Function
    Next conflictIndex
    ValidateConflictRecords = True
End Function

Private Function ValidateConflictComponent(ByVal componentName As String, ByVal baseComponent As Variant, ByVal localComponent As Variant, ByVal remoteComponent As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim equal As Boolean
    Dim ignoredValue As Variant

    If componentName = "TextFont" Then
        If Not ValueFromTextFont(baseComponent, ignoredValue, errorInfo) Then Exit Function
        If Not ValueFromTextFont(localComponent, ignoredValue, errorInfo) Then Exit Function
        If Not ValueFromTextFont(remoteComponent, ignoredValue, errorInfo) Then Exit Function
    ElseIf componentName = "Fill" Then
        If Not CompareFormattingData(baseComponent, baseComponent, equal, errorInfo) Then Exit Function
        If Not CompareFormattingData(localComponent, localComponent, equal, errorInfo) Then Exit Function
        If Not CompareFormattingData(remoteComponent, remoteComponent, equal, errorInfo) Then Exit Function
    Else
        Exit Function
    End If

    If Not CompareFormattingData(baseComponent, localComponent, equal, errorInfo) Then Exit Function
    If equal Then Exit Function
    If Not CompareFormattingData(baseComponent, remoteComponent, equal, errorInfo) Then Exit Function
    If equal Then Exit Function
    If Not CompareFormattingData(localComponent, remoteComponent, equal, errorInfo) Then Exit Function
    If equal Then Exit Function
    ValidateConflictComponent = True
End Function

Private Function TryReadLong(ByVal value As Variant, ByRef number As Long) As Boolean
    Dim numericValue As Double

    On Error GoTo InvalidValue
    If Not IsNumeric(value) Then Exit Function
    numericValue = CDbl(value)
    number = CLng(numericValue)
    If numericValue <> CDbl(number) Then Exit Function
    TryReadLong = True
    Exit Function

InvalidValue:
    Err.Clear
End Function

Public Function IsMergeResolved(ByRef result As MergeResult) As Boolean
    IsMergeResolved = result.Resolved And result.PendingCount = 0
End Function

Private Function ClassifyComponent(ByVal baseComponent As Variant, ByVal localComponent As Variant, ByVal remoteComponent As Variant, ByRef errorInfo As AppErrorInfo) As MergeCellOutcome
    Dim localMatchesBase As Boolean
    Dim remoteMatchesBase As Boolean
    Dim localMatchesRemote As Boolean

    If Not CompareFormattingData(baseComponent, localComponent, localMatchesBase, errorInfo) Then Exit Function
    If Not CompareFormattingData(baseComponent, remoteComponent, remoteMatchesBase, errorInfo) Then Exit Function
    If Not CompareFormattingData(localComponent, remoteComponent, localMatchesRemote, errorInfo) Then Exit Function
    If localMatchesBase Then
        ClassifyComponent = MergeTakeRemote
    ElseIf remoteMatchesBase Then
        ClassifyComponent = MergeTakeLocal
    ElseIf localMatchesRemote Then
        ClassifyComponent = MergeSameValue
    Else
        ClassifyComponent = MergeNeedsChoice
    End If
End Function

Private Function ValidateCellTriplet(ByVal baseValue As Variant, ByVal localValue As Variant, ByVal remoteValue As Variant, ByVal baseFont As Variant, ByVal localFont As Variant, ByVal remoteFont As Variant, ByVal baseFill As Variant, ByVal localFill As Variant, ByVal remoteFill As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim equal As Boolean
    Dim compareError As AppErrorInfo

    If Not CompareCellValues(baseValue, baseValue, equal, compareError) Then
        errorInfo = compareError
        Exit Function
    End If
    If Not CompareCellValues(localValue, localValue, equal, compareError) Then
        errorInfo = compareError
        Exit Function
    End If
    If Not CompareCellValues(remoteValue, remoteValue, equal, compareError) Then
        errorInfo = compareError
        Exit Function
    End If
    If Not TextFontMatchesValue(baseValue, baseFont, errorInfo) Then Exit Function
    If Not TextFontMatchesValue(localValue, localFont, errorInfo) Then Exit Function
    If Not TextFontMatchesValue(remoteValue, remoteFont, errorInfo) Then Exit Function
    If Not CompareFormattingData(baseFont, baseFont, equal, errorInfo) Then Exit Function
    If Not CompareFormattingData(localFont, localFont, equal, errorInfo) Then Exit Function
    If Not CompareFormattingData(remoteFont, remoteFont, equal, errorInfo) Then Exit Function
    If Not CompareFormattingData(baseFill, baseFill, equal, errorInfo) Then Exit Function
    If Not CompareFormattingData(localFill, localFill, equal, errorInfo) Then Exit Function
    If Not CompareFormattingData(remoteFill, remoteFill, equal, errorInfo) Then Exit Function
    ValidateCellTriplet = True
End Function

Private Function TextFontMatchesValue(ByVal cellValue As Variant, ByVal fontData As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim valueKind As String
    Dim textValue As String
    Dim emptyFont As String
    Dim runs As Variant

    If VarType(fontData) <> vbString Then GoTo InvalidData
    If Not ParseTextFontData(CStr(fontData), valueKind, textValue, runs, emptyFont, errorInfo) Then Exit Function
    If VarType(cellValue) = vbEmpty Then
        TextFontMatchesValue = (valueKind = "E")
    ElseIf VarType(cellValue) = vbString Then
        TextFontMatchesValue = (valueKind = "S" And StrComp(textValue, CStr(cellValue), vbBinaryCompare) = 0)
    End If
    If Not TextFontMatchesValue Then GoTo InvalidData
    Exit Function

InvalidData:
    errorInfo = MakeAppError(ERR_MERGE_INVALID_INPUT, "MdlMerge.TextFontMatchesValue", "A text/font snapshot does not match its corresponding cell value.")
End Function

Private Function ValueFromTextFont(ByVal fontData As Variant, ByRef cellValue As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim valueKind As String
    Dim textValue As String
    Dim emptyFont As String
    Dim runs As Variant

    If VarType(fontData) <> vbString Then GoTo InvalidData
    If Not ParseTextFontData(CStr(fontData), valueKind, textValue, runs, emptyFont, errorInfo) Then Exit Function
    If valueKind = "E" Then
        cellValue = Empty
    ElseIf valueKind = "S" Then
        cellValue = textValue
    Else
        GoTo InvalidData
    End If
    ValueFromTextFont = True
    Exit Function

InvalidData:
    errorInfo = MakeAppError(ERR_MERGE_INVALID_INPUT, "MdlMerge.ValueFromTextFont", "The chosen text/font component cannot supply a cell value.")
End Function

Private Sub SetConflict(ByRef conflicts As Variant, ByVal index As Long, ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal componentName As String, ByVal baseComponent As Variant, ByVal localComponent As Variant, ByVal remoteComponent As Variant)
    conflicts(index, 1) = rowIndex
    conflicts(index, 2) = columnIndex
    conflicts(index, 3) = componentName
    conflicts(index, 4) = baseComponent
    conflicts(index, 5) = localComponent
    conflicts(index, 6) = remoteComponent
End Sub

Private Sub ResetMergeResult(ByRef result As MergeResult)
    result.MergedValues = Empty
    result.MergedFonts = Empty
    result.MergedFills = Empty
    result.Conflicts = Empty
    result.ConflictCount = 0
    result.PendingCount = 0
    result.Resolved = False
End Sub

Private Function ValidateDeclaredColumns(ByRef declaredColumns As Variant, ByRef lowerBound As Long, ByRef upperBound As Long) As Boolean
    Dim probe As Long
    Dim index As Long
    Dim columnIndex As Long

    If Not IsArray(declaredColumns) Then Exit Function
    On Error Resume Next
    lowerBound = LBound(declaredColumns, 1)
    upperBound = UBound(declaredColumns, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    probe = LBound(declaredColumns, 2)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    If upperBound < lowerBound Then Exit Function
    For index = lowerBound To upperBound
        If Not IsNumeric(declaredColumns(index)) Then Exit Function
        columnIndex = CLng(declaredColumns(index))
        If CDbl(declaredColumns(index)) <> CDbl(columnIndex) Or columnIndex < 1 Then Exit Function
    Next index
    ValidateDeclaredColumns = True
End Function

Private Function MatrixBoundsMatch(ByRef values As Variant, ByVal rowLow As Long, ByVal rowHigh As Long, ByVal columnLow As Long, ByVal columnHigh As Long) As Boolean
    Dim actualRowLow As Long
    Dim actualRowHigh As Long
    Dim actualColumnLow As Long
    Dim actualColumnHigh As Long

    If Not GetMatrixBounds(values, actualRowLow, actualRowHigh, actualColumnLow, actualColumnHigh) Then Exit Function
    MatrixBoundsMatch = (actualRowLow = rowLow And actualRowHigh = rowHigh And actualColumnLow = columnLow And actualColumnHigh = columnHigh)
End Function

Private Function GetMatrixBounds(ByRef values As Variant, ByRef rowLow As Long, ByRef rowHigh As Long, ByRef columnLow As Long, ByRef columnHigh As Long) As Boolean
    Dim probe As Long

    If Not IsArray(values) Then Exit Function
    On Error Resume Next
    rowLow = LBound(values, 1)
    rowHigh = UBound(values, 1)
    columnLow = LBound(values, 2)
    columnHigh = UBound(values, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    probe = LBound(values, 3)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    GetMatrixBounds = (rowHigh >= rowLow And columnHigh >= columnLow)
End Function
