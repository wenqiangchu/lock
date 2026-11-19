Attribute VB_Name = "MdlSchema"
Option Explicit

Public Enum SchemaColumnKind
    SchemaColumnIgnored = 0
    SchemaColumnProduct = 1
    SchemaColumnAmg = 2
End Enum

Public Type SheetSchema
    ColumnCount As Long
    HeaderNames As Variant
    NormalizedNames As Variant
    ColumnKinds As Variant
    AmgColumnIndexes As Variant
    AnchorColumnIndexes As Variant
    Fingerprint As String
End Type

Public Const ERR_SCHEMA_INVALID As Long = 1101
Public Const ERR_SCHEMA_AMG_DUPLICATE As Long = 1102
Public Const ERR_SCHEMA_NO_ANCHOR As Long = 1103

Public Function ParseHeader(ByRef headerValues As Variant, ByRef schema As SheetSchema, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim lowerFirst As Long
    Dim upperFirst As Long
    Dim lowerSecond As Long
    Dim upperSecond As Long
    Dim dimensionCount As Long
    Dim probeBound As Long
    Dim count As Long
    Dim index As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim headerNames() As Variant
    Dim normalizedNames() As Variant
    Dim columnKinds() As Variant
    Dim amgIndexes() As Variant
    Dim anchorIndexes() As Variant
    Dim amgCount As Long
    Dim anchorCount As Long
    Dim rawValue As Variant
    Dim normalizedName As String
    Dim seenAmg As Collection
    Dim amgKey As String

    On Error GoTo InvalidInput
    ResetSchema schema

    If IsArray(headerValues) Then
        On Error Resume Next
        lowerFirst = LBound(headerValues, 1)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo InvalidInput
            GoTo InvalidInput
        End If
        upperFirst = UBound(headerValues, 1)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo InvalidInput
            GoTo InvalidInput
        End If
        Err.Clear
        lowerSecond = LBound(headerValues, 2)
        If Err.Number = 9 Then
            dimensionCount = 1
        ElseIf Err.Number <> 0 Then
            Err.Clear
            On Error GoTo InvalidInput
            GoTo InvalidInput
        Else
            upperSecond = UBound(headerValues, 2)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo InvalidInput
                GoTo InvalidInput
            End If
            Err.Clear
            probeBound = LBound(headerValues, 3)
            If Err.Number = 0 Then GoTo InvalidInput
            If Err.Number <> 9 Then
                Err.Clear
                On Error GoTo InvalidInput
                GoTo InvalidInput
            End If
            dimensionCount = 2
        End If
        Err.Clear
        On Error GoTo InvalidInput
        If dimensionCount = 1 Then
            count = upperFirst - lowerFirst + 1
        ElseIf dimensionCount = 2 Then
            If upperFirst = lowerFirst Then
                count = upperSecond - lowerSecond + 1
            ElseIf upperSecond = lowerSecond Then
                count = upperFirst - lowerFirst + 1
            Else
                GoTo InvalidInput
            End If
        Else
            GoTo InvalidInput
        End If
    Else
        count = 1
    End If
    If count < 1 Then GoTo InvalidInput

    ReDim headerNames(1 To count)
    ReDim normalizedNames(1 To count)
    ReDim columnKinds(1 To count)
    ReDim amgIndexes(1 To count)
    ReDim anchorIndexes(1 To 5)
    Set seenAmg = New Collection

    For index = 1 To count
        If IsArray(headerValues) Then
            If dimensionCount = 1 Then
                rawValue = headerValues(lowerFirst + index - 1)
            ElseIf upperFirst = lowerFirst Then
                rawValue = headerValues(lowerFirst, lowerSecond + index - 1)
            Else
                rawValue = headerValues(lowerFirst + index - 1, lowerSecond)
            End If
        Else
            rawValue = headerValues
        End If

        If IsError(rawValue) Or IsNull(rawValue) Then GoTo InvalidInput
        If IsEmpty(rawValue) Then
            headerNames(index) = vbNullString
            normalizedNames(index) = vbNullString
            columnKinds(index) = SchemaColumnIgnored
        ElseIf VarType(rawValue) = vbString Then
            headerNames(index) = CStr(rawValue)
            normalizedName = UCase$(Trim$(CStr(rawValue)))
            normalizedNames(index) = normalizedName
            If Len(normalizedName) = 0 Then
                columnKinds(index) = SchemaColumnIgnored
            ElseIf IsAmgHeader(normalizedName) Then
                amgKey = "C" & normalizedName
                On Error Resume Next
                seenAmg.Add index, amgKey
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo InvalidInput
                    errorInfo = MakeAppError(ERR_SCHEMA_AMG_DUPLICATE, "MdlSchema.ParseHeader", "AMG header names are duplicated after normalization.", CStr(index))
                    Exit Function
                End If
                On Error GoTo InvalidInput
                amgCount = amgCount + 1
                amgIndexes(amgCount) = index
                columnKinds(index) = SchemaColumnAmg
            Else
                columnKinds(index) = SchemaColumnProduct
                If index <= 5 Then
                    anchorCount = anchorCount + 1
                    anchorIndexes(anchorCount) = index
                End If
            End If
        Else
            GoTo InvalidInput
        End If
    Next index

    If anchorCount = 0 Then
        errorInfo = MakeAppError(ERR_SCHEMA_NO_ANCHOR, "MdlSchema.ParseHeader", "At least one product header is required in columns A through E.")
        Exit Function
    End If

    schema.ColumnCount = count
    schema.HeaderNames = headerNames
    schema.NormalizedNames = normalizedNames
    schema.ColumnKinds = columnKinds
    If amgCount > 0 Then
        ReDim Preserve amgIndexes(1 To amgCount)
        schema.AmgColumnIndexes = amgIndexes
    Else
        schema.AmgColumnIndexes = Empty
    End If
    ReDim Preserve anchorIndexes(1 To anchorCount)
    schema.AnchorColumnIndexes = anchorIndexes
    schema.Fingerprint = BuildFullFingerprint(schema)
    ParseHeader = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSchema.ParseHeader", "")
    Exit Function

InvalidInput:
    ResetSchema schema
    errorInfo = MakeAppError(ERR_SCHEMA_INVALID, "MdlSchema.ParseHeader", "The header must contain only text or empty values in a one-dimensional row or column.")
End Function

Public Function GetEditableColumns(ByRef schema As SheetSchema) As Variant
    GetEditableColumns = schema.AmgColumnIndexes
End Function

Public Function GetAnchorColumns(ByRef schema As SheetSchema) As Variant
    GetAnchorColumns = schema.AnchorColumnIndexes
End Function

Public Function BuildColumnFingerprint(ByRef schema As SheetSchema, ByVal columnIndexes As Variant) As String
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long
    Dim columnIndex As Long
    Dim probeBound As Long
    Dim value As String
    Dim fingerprint As String

    If schema.ColumnCount < 1 Then Exit Function
    If Not IsArray(columnIndexes) Then Exit Function
    On Error Resume Next
    lowerBound = LBound(columnIndexes, 1)
    upperBound = UBound(columnIndexes, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    probeBound = LBound(columnIndexes, 2)
    If Err.Number = 0 Or Err.Number <> 9 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    For index = lowerBound To upperBound
        If Not IsNumeric(columnIndexes(index)) Then Exit Function
        columnIndex = CLng(columnIndexes(index))
        If CDbl(columnIndexes(index)) <> CDbl(columnIndex) Then Exit Function
        If columnIndex < 1 Or columnIndex > schema.ColumnCount Then Exit Function
        value = CStr(schema.NormalizedNames(columnIndex))
        fingerprint = fingerprint & CStr(columnIndex) & ":" & CStr(schema.ColumnKinds(columnIndex)) & ":" & CStr(Len(value)) & ":" & value & ";"
    Next index
    BuildColumnFingerprint = fingerprint
End Function

Private Function IsAmgHeader(ByVal normalizedName As String) As Boolean
    IsAmgHeader = (Len(normalizedName) >= 3 And Left$(normalizedName, 3) = "AMG")
End Function

Private Function BuildFullFingerprint(ByRef schema As SheetSchema) As String
    Dim columnIndex As Long
    Dim normalizedName As String
    Dim fingerprint As String

    For columnIndex = 1 To schema.ColumnCount
        normalizedName = CStr(schema.NormalizedNames(columnIndex))
        fingerprint = fingerprint & CStr(columnIndex) & ":" & CStr(schema.ColumnKinds(columnIndex)) & ":" & CStr(Len(normalizedName)) & ":" & normalizedName & ";"
    Next columnIndex
    BuildFullFingerprint = fingerprint
End Function

Private Sub ResetSchema(ByRef schema As SheetSchema)
    schema.ColumnCount = 0
    schema.HeaderNames = Empty
    schema.NormalizedNames = Empty
    schema.ColumnKinds = Empty
    schema.AmgColumnIndexes = Empty
    schema.AnchorColumnIndexes = Empty
    schema.Fingerprint = vbNullString
End Sub
