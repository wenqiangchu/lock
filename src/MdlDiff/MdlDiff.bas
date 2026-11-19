Attribute VB_Name = "MdlDiff"
Option Explicit

Public Const ERR_DIFF_INVALID_INPUT As Long = 1201

Public Function CompareSnapshots(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef baseFonts As Variant, ByRef localFonts As Variant, ByRef baseFills As Variant, ByRef localFills As Variant, ByRef declaredColumns As Variant, ByRef changes As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim localRowLow As Long
    Dim localRowHigh As Long
    Dim localColumnLow As Long
    Dim localColumnHigh As Long
    Dim fontRowLow As Long
    Dim fontRowHigh As Long
    Dim fontColumnLow As Long
    Dim fontColumnHigh As Long
    Dim localFontRowLow As Long
    Dim localFontRowHigh As Long
    Dim localFontColumnLow As Long
    Dim localFontColumnHigh As Long
    Dim fillRowLow As Long
    Dim fillRowHigh As Long
    Dim fillColumnLow As Long
    Dim fillColumnHigh As Long
    Dim localFillRowLow As Long
    Dim localFillRowHigh As Long
    Dim localFillColumnLow As Long
    Dim localFillColumnHigh As Long
    Dim declaredLow As Long
    Dim declaredHigh As Long
    Dim rowIndex As Long
    Dim declaredIndex As Long
    Dim priorIndex As Long
    Dim columnIndex As Long
    Dim changeCount As Long
    Dim changeIndex As Long
    Dim formatEqual As Boolean
    Dim cellError As AppErrorInfo
    Dim result() As Variant

    changes = Empty
    errorInfo = MakeAppError(ERR_NONE, "MdlDiff.CompareSnapshots", "")
    If Not ValidateDeclaredColumns(declaredColumns, declaredLow, declaredHigh) Then GoTo InvalidInput

    If IsEmpty(baseValues) And IsEmpty(localValues) And IsEmpty(baseFonts) And IsEmpty(localFonts) And IsEmpty(baseFills) And IsEmpty(localFills) Then
        CompareSnapshots = True
        Exit Function
    End If
    If IsEmpty(baseValues) Or IsEmpty(localValues) Or IsEmpty(baseFonts) Or IsEmpty(localFonts) Or IsEmpty(baseFills) Or IsEmpty(localFills) Then GoTo InvalidInput
    If Not GetMatrixBounds(baseValues, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not GetMatrixBounds(localValues, localRowLow, localRowHigh, localColumnLow, localColumnHigh) Then GoTo InvalidInput
    If Not GetMatrixBounds(baseFonts, fontRowLow, fontRowHigh, fontColumnLow, fontColumnHigh) Then GoTo InvalidInput
    If Not GetMatrixBounds(localFonts, localFontRowLow, localFontRowHigh, localFontColumnLow, localFontColumnHigh) Then GoTo InvalidInput
    If Not GetMatrixBounds(baseFills, fillRowLow, fillRowHigh, fillColumnLow, fillColumnHigh) Then GoTo InvalidInput
    If Not GetMatrixBounds(localFills, localFillRowLow, localFillRowHigh, localFillColumnLow, localFillColumnHigh) Then GoTo InvalidInput

    If Not SameBounds(rowLow, rowHigh, columnLow, columnHigh, localRowLow, localRowHigh, localColumnLow, localColumnHigh) Then GoTo InvalidInput
    If Not SameBounds(rowLow, rowHigh, columnLow, columnHigh, fontRowLow, fontRowHigh, fontColumnLow, fontColumnHigh) Then GoTo InvalidInput
    If Not SameBounds(rowLow, rowHigh, columnLow, columnHigh, localFontRowLow, localFontRowHigh, localFontColumnLow, localFontColumnHigh) Then GoTo InvalidInput
    If Not SameBounds(rowLow, rowHigh, columnLow, columnHigh, fillRowLow, fillRowHigh, fillColumnLow, fillColumnHigh) Then GoTo InvalidInput
    If Not SameBounds(rowLow, rowHigh, columnLow, columnHigh, localFillRowLow, localFillRowHigh, localFillColumnLow, localFillColumnHigh) Then GoTo InvalidInput
    For declaredIndex = declaredLow To declaredHigh
        If Not IsNumeric(declaredColumns(declaredIndex)) Then GoTo InvalidInput
        columnIndex = CLng(declaredColumns(declaredIndex))
        If CDbl(declaredColumns(declaredIndex)) <> CDbl(columnIndex) Or columnIndex < columnLow Or columnIndex > columnHigh Then GoTo InvalidInput
        For priorIndex = declaredLow To declaredIndex - 1
            If CLng(declaredColumns(priorIndex)) = columnIndex Then GoTo InvalidInput
        Next priorIndex
    Next declaredIndex

    For rowIndex = rowLow To rowHigh
        For declaredIndex = declaredLow To declaredHigh
            columnIndex = CLng(declaredColumns(declaredIndex))
            If Not ValidateCellPair(baseValues(rowIndex, columnIndex), localValues(rowIndex, columnIndex), baseFonts(rowIndex, columnIndex), localFonts(rowIndex, columnIndex), baseFills(rowIndex, columnIndex), localFills(rowIndex, columnIndex), cellError) Then
                errorInfo = cellError
                Exit Function
            End If
            If Not CompareFormattingData(baseFonts(rowIndex, columnIndex), localFonts(rowIndex, columnIndex), formatEqual, cellError) Then
                errorInfo = cellError
                Exit Function
            End If
            If Not formatEqual Then changeCount = changeCount + 1
            If Not CompareFormattingData(baseFills(rowIndex, columnIndex), localFills(rowIndex, columnIndex), formatEqual, cellError) Then
                errorInfo = cellError
                Exit Function
            End If
            If Not formatEqual Then changeCount = changeCount + 1
        Next declaredIndex
    Next rowIndex

    If changeCount = 0 Then
        CompareSnapshots = True
        Exit Function
    End If
    ReDim result(1 To changeCount, 1 To 5)
    For rowIndex = rowLow To rowHigh
        For declaredIndex = declaredLow To declaredHigh
            columnIndex = CLng(declaredColumns(declaredIndex))
            If Not CompareFormattingData(baseFonts(rowIndex, columnIndex), localFonts(rowIndex, columnIndex), formatEqual, cellError) Then
                errorInfo = cellError
                Exit Function
            End If
            If Not formatEqual Then
                changeIndex = changeIndex + 1
                result(changeIndex, 1) = rowIndex
                result(changeIndex, 2) = columnIndex
                result(changeIndex, 3) = "TextFont"
                result(changeIndex, 4) = baseFonts(rowIndex, columnIndex)
                result(changeIndex, 5) = localFonts(rowIndex, columnIndex)
            End If
            If Not CompareFormattingData(baseFills(rowIndex, columnIndex), localFills(rowIndex, columnIndex), formatEqual, cellError) Then
                errorInfo = cellError
                Exit Function
            End If
            If Not formatEqual Then
                changeIndex = changeIndex + 1
                result(changeIndex, 1) = rowIndex
                result(changeIndex, 2) = columnIndex
                result(changeIndex, 3) = "Fill"
                result(changeIndex, 4) = baseFills(rowIndex, columnIndex)
                result(changeIndex, 5) = localFills(rowIndex, columnIndex)
            End If
        Next declaredIndex
    Next rowIndex

    changes = result
    CompareSnapshots = True
    Exit Function

InvalidInput:
    changes = Empty
    errorInfo = MakeAppError(ERR_DIFF_INVALID_INPUT, "MdlDiff.CompareSnapshots", "Snapshot arrays, formatting records, or declared column indexes are invalid or do not match.")
End Function

Private Function ValidateCellPair(ByVal baseValue As Variant, ByVal localValue As Variant, ByVal baseFont As Variant, ByVal localFont As Variant, ByVal baseFill As Variant, ByVal localFill As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim valuesEqual As Boolean
    Dim formatEqual As Boolean
    Dim compareError As AppErrorInfo

    If Not CompareCellValues(baseValue, localValue, valuesEqual, compareError) Then
        errorInfo = compareError
        Exit Function
    End If
    If Not ValidateTextFontValue(baseValue, baseFont, errorInfo) Then Exit Function
    If Not ValidateTextFontValue(localValue, localFont, errorInfo) Then Exit Function
    If Not CompareFormattingData(baseFont, localFont, formatEqual, errorInfo) Then Exit Function
    If Not CompareFormattingData(baseFill, baseFill, formatEqual, errorInfo) Then Exit Function
    If Not CompareFormattingData(localFill, localFill, formatEqual, errorInfo) Then Exit Function
    ValidateCellPair = True
End Function

Private Function ValidateTextFontValue(ByVal cellValue As Variant, ByVal fontData As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim valueKind As String
    Dim textValue As String
    Dim emptyFont As String
    Dim runs As Variant
    Dim expectedKind As String

    If VarType(fontData) <> vbString Then GoTo InvalidData
    If Not ParseTextFontData(CStr(fontData), valueKind, textValue, runs, emptyFont, errorInfo) Then Exit Function
    If VarType(cellValue) = vbEmpty Then
        expectedKind = "E"
    ElseIf VarType(cellValue) = vbString Then
        expectedKind = "S"
    Else
        GoTo InvalidData
    End If
    If valueKind <> expectedKind Then GoTo InvalidData
    If expectedKind = "S" And StrComp(textValue, CStr(cellValue), vbBinaryCompare) <> 0 Then GoTo InvalidData
    ValidateTextFontValue = True
    Exit Function

InvalidData:
    errorInfo = MakeAppError(ERR_DIFF_INVALID_INPUT, "MdlDiff.ValidateTextFontValue", "A text/font snapshot does not match its corresponding cell value.")
End Function

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

Private Function SameBounds(ByVal expectedRowLow As Long, ByVal expectedRowHigh As Long, ByVal expectedColumnLow As Long, ByVal expectedColumnHigh As Long, ByVal actualRowLow As Long, ByVal actualRowHigh As Long, ByVal actualColumnLow As Long, ByVal actualColumnHigh As Long) As Boolean
    SameBounds = (expectedRowLow = actualRowLow And expectedRowHigh = actualRowHigh And expectedColumnLow = actualColumnLow And expectedColumnHigh = actualColumnHigh)
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
