Attribute VB_Name = "TestMdlSchema"
Option Explicit

Public Sub RunMdlSchemaTests()
    RunTestCase "AnchorSelection"
    RunTestCase "EmptyHeaders"
    RunTestCase "MissingAnchor"
    RunTestCase "DuplicateAmg"
    RunTestCase "InvalidTypes"
    RunTestCase "Fingerprint"
End Sub

Public Sub TestMdlSchemaAnchorSelection()
    RunTestCase "AnchorSelection"
End Sub

Public Sub TestMdlSchemaEmptyHeaders()
    RunTestCase "EmptyHeaders"
End Sub

Public Sub TestMdlSchemaMissingAnchor()
    RunTestCase "MissingAnchor"
End Sub

Public Sub TestMdlSchemaDuplicateAmg()
    RunTestCase "DuplicateAmg"
End Sub

Public Sub TestMdlSchemaInvalidTypes()
    RunTestCase "InvalidTypes"
End Sub

Public Sub TestMdlSchemaFingerprint()
    RunTestCase "Fingerprint"
End Sub

Private Sub RunTestCase(ByVal caseName As String)
    On Error GoTo Failed
    Select Case caseName
        Case "AnchorSelection"
            TestAnchorSelectionAndEditableColumns
        Case "EmptyHeaders"
            TestEmptyHeadersPreserveColumnPositions
        Case "MissingAnchor"
            TestNoAnchorFails
        Case "DuplicateAmg"
            TestNormalizedDuplicateAmgFails
        Case "InvalidTypes"
            TestInvalidHeaderTypesFail
        Case "Fingerprint"
            TestFingerprintIsStableAndColumnSensitive
        Case Else
            Err.Raise vbObjectError + 1804, "TestMdlSchema", "Unknown test case: " & caseName
    End Select
    Exit Sub

Failed:
    TestSupport.RecordTestFailure "TestMdlSchema." & caseName & ": " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestAnchorSelectionAndEditableColumns()
    Dim headers As Variant
    Dim schema As SheetSchema
    Dim errorInfo As AppErrorInfo
    Dim anchors As Variant
    Dim editable As Variant

    ReDim headers(1 To 1, 1 To 7)
    headers(1, 1) = "product_a"
    headers(1, 2) = "AMG_First"
    headers(1, 3) = "product_b"
    headers(1, 4) = "AMG_Second"
    headers(1, 5) = "product_c"
    headers(1, 6) = "AMG_Third"
    headers(1, 7) = "product_after_e"

    AssertTrue ParseHeader(headers, schema, errorInfo), "A valid header row should parse."
    anchors = GetAnchorColumns(schema)
    editable = GetEditableColumns(schema)
    AssertEqual CStr(UBound(anchors) - LBound(anchors) + 1), "3", "Only the three product columns in A-E should anchor rows."
    AssertEqual CStr(anchors(1)), "1", "The first anchor should preserve physical column A."
    AssertEqual CStr(anchors(2)), "3", "The second anchor should preserve physical column C."
    AssertEqual CStr(anchors(3)), "5", "The third anchor should preserve physical column E."
    AssertEqual CStr(UBound(editable) - LBound(editable) + 1), "3", "All AMG columns should remain editable."
    AssertEqual CStr(editable(3)), "6", "The AMG column after E should remain editable."
    AssertEqual CStr(schema.ColumnKinds(7)), CStr(SchemaColumnProduct), "Product columns after E should be classified but not anchored."
End Sub

Private Sub TestEmptyHeadersPreserveColumnPositions()
    Dim headers As Variant
    Dim schema As SheetSchema
    Dim errorInfo As AppErrorInfo
    Dim anchors As Variant

    ReDim headers(0 To 5)
    headers(0) = "product"
    headers(1) = Empty
    headers(2) = "AMG_Value"
    headers(3) = ""
    headers(4) = "AMG_Notes"
    headers(5) = Empty

    AssertTrue ParseHeader(headers, schema, errorInfo), "Empty headers should not invalidate other columns."
    anchors = GetAnchorColumns(schema)
    AssertEqual CStr(schema.ColumnCount), "6", "Empty headers should retain physical column positions."
    AssertEqual CStr(anchors(1)), "1", "The first physical column should remain the anchor."
    AssertEqual CStr(schema.ColumnKinds(2)), CStr(SchemaColumnIgnored), "An Empty header should be ignored."
    AssertEqual CStr(schema.ColumnKinds(4)), CStr(SchemaColumnIgnored), "A zero-length header should be ignored."
    AssertEqual CStr(schema.AmgColumnIndexes(2)), "5", "An ignored column should not shift AMG physical indexes."
End Sub

Private Sub TestNoAnchorFails()
    Dim headers As Variant
    Dim schema As SheetSchema
    Dim errorInfo As AppErrorInfo

    ReDim headers(1 To 6)
    headers(1) = "AMG_One"
    headers(2) = "AMG_Two"
    headers(3) = ""
    headers(4) = Empty
    headers(5) = "AMG_Three"
    headers(6) = "product_after_e"

    AssertFalse ParseHeader(headers, schema, errorInfo), "A sheet without an A-E product anchor must fail."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_SCHEMA_NO_ANCHOR), "The missing-anchor failure should be distinguishable."
End Sub

Private Sub TestNormalizedDuplicateAmgFails()
    Dim headers As Variant
    Dim schema As SheetSchema
    Dim errorInfo As AppErrorInfo

    ReDim headers(1 To 3)
    headers(1) = "product"
    headers(2) = "AMG_One"
    headers(3) = " amg_one "

    AssertFalse ParseHeader(headers, schema, errorInfo), "AMG names that normalize to the same key must fail."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_SCHEMA_AMG_DUPLICATE), "The duplicate-AMG failure should be distinguishable."
End Sub

Private Sub TestInvalidHeaderTypesFail()
    Dim headers As Variant
    Dim schema As SheetSchema
    Dim errorInfo As AppErrorInfo

    ReDim headers(1 To 2)
    headers(1) = "product"
    headers(2) = 5
    AssertFalse ParseHeader(headers, schema, errorInfo), "Non-text non-empty headers must fail."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_SCHEMA_INVALID), "Invalid header values should return an input error."
End Sub

Private Sub TestFingerprintIsStableAndColumnSensitive()
    Dim headers As Variant
    Dim schema As SheetSchema
    Dim errorInfo As AppErrorInfo
    Dim firstFingerprint As String
    Dim secondFingerprint As String
    Dim changedFingerprint As String
    Dim indexes As Variant

    ReDim headers(1 To 5)
    ReDim indexes(1 To 2)
    headers(1) = "product"
    headers(2) = "AMG_One"
    headers(3) = "AMG_Two"
    headers(4) = "product_b"
    headers(5) = "AMG_Three"
    indexes(1) = 2
    indexes(2) = 5

    AssertTrue ParseHeader(headers, schema, errorInfo), "The fingerprint fixture should parse."
    firstFingerprint = BuildColumnFingerprint(schema, indexes)
    secondFingerprint = BuildColumnFingerprint(schema, indexes)
    AssertEqual firstFingerprint, secondFingerprint, "The same ordered columns should have the same fingerprint."

    indexes(2) = 3
    changedFingerprint = BuildColumnFingerprint(schema, indexes)
    AssertFalse (StrComp(firstFingerprint, changedFingerprint, vbBinaryCompare) = 0), "Changing a physical column should change the fingerprint."
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 1801, "TestMdlSchema", message
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, ByVal message As String)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 1802, "TestMdlSchema", message
End Sub
