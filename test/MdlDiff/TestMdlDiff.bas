Attribute VB_Name = "TestMdlDiff"
Option Explicit

Public Sub RunMdlDiffTests()
    RunTestCase "NoChanges"
    RunTestCase "DeclaredComponentsOnly"
    RunTestCase "IndependentFontAndFillChanges"
    RunTestCase "InvalidDeclaredValue"
    RunTestCase "MismatchedDimensions"
    RunTestCase "InvalidDeclaredColumns"
    RunTestCase "EmptySnapshots"
End Sub

Private Sub RunTestCase(ByVal caseName As String)
    On Error GoTo Failed
    Select Case caseName
        Case "NoChanges": TestNoChanges
        Case "DeclaredComponentsOnly": TestDeclaredComponentsOnly
        Case "IndependentFontAndFillChanges": TestIndependentFontAndFillChanges
        Case "InvalidDeclaredValue": TestInvalidDeclaredValue
        Case "MismatchedDimensions": TestMismatchedDimensions
        Case "InvalidDeclaredColumns": TestInvalidDeclaredColumns
        Case "EmptySnapshots": TestEmptySnapshots
        Case Else: Err.Raise vbObjectError + 1901, "TestMdlDiff", "Unknown test case: " & caseName
    End Select
    Exit Sub
Failed:
    TestSupport.RecordTestFailure "TestMdlDiff." & caseName & ": " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestNoChanges()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    MakeBasicSnapshots baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns
    AssertTrue CompareSnapshots(baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns, changes, errorInfo), "Equal snapshots should compare successfully."
    AssertTrue IsEmpty(changes), "Equal snapshots should return no change records."
End Sub

Private Sub TestDeclaredComponentsOnly()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo
    Dim normalFont As String
    Dim boldFont As String

    MakeBasicSnapshots baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns
    normalFont = TestFontSignature(False, 0)
    boldFont = TestFontSignature(True, 0)
    localFonts(2, 2) = TestTextFontData(localValues(2, 2), boldFont)
    localValues(2, 4) = "added"
    localFonts(2, 4) = TestTextFontData(localValues(2, 4), normalFont)
    baseValues(2, 3) = 25
    localValues(2, 3) = 99

    AssertTrue CompareSnapshots(baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns, changes, errorInfo), "Declared string and font changes should be reported."
    AssertEqual CStr(UBound(changes, 1)), "2", "Only the two declared-component changes should be included."
    AssertEqual CStr(changes(1, 1)), "2", "The first change should preserve its physical row."
    AssertEqual CStr(changes(1, 2)), "2", "The first change should preserve its physical column."
    AssertEqual CStr(changes(1, 3)), "TextFont", "Font changes should report the TextFont component."
    AssertEqual CStr(changes(2, 2)), "4", "The second change should preserve its physical column."
    AssertEqual CStr(changes(2, 3)), "TextFont", "A value change should be included in the TextFont component."
End Sub

Private Sub TestIndependentFontAndFillChanges()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    MakeBasicSnapshots baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns
    localFonts(2, 2) = TestTextFontData(localValues(2, 2), TestFontSignature(False, 255))
    localFills(2, 4) = TestFillData(65280)

    AssertTrue CompareSnapshots(baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns, changes, errorInfo), "Separate font and fill changes should compare."
    AssertEqual CStr(UBound(changes, 1)), "2", "Both independent component changes should be listed."
    AssertEqual CStr(changes(1, 3)), "TextFont", "The first component should identify font data."
    AssertEqual CStr(changes(2, 3)), "Fill", "The second component should identify fill data."
End Sub

Private Sub TestInvalidDeclaredValue()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    MakeBasicSnapshots baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns
    baseValues(2, 2) = 1
    localValues(2, 2) = 1
    AssertFalse CompareSnapshots(baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns, changes, errorInfo), "Non-string declared values must fail even when equal."
    AssertTrue IsEmpty(changes), "Invalid values must not return partial change records."
End Sub

Private Sub TestMismatchedDimensions()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    MakeBasicSnapshots baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns
    ReDim localFonts(2 To 4, 1 To 4)
    AssertFalse CompareSnapshots(baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns, changes, errorInfo), "Different physical bounds must fail."
    AssertTrue IsEmpty(changes), "A dimension error must not return partial changes."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_DIFF_INVALID_INPUT), "Dimension errors should have the diff input error code."
End Sub

Private Sub TestInvalidDeclaredColumns()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    MakeBasicSnapshots baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns
    ReDim declaredColumns(1 To 2)
    declaredColumns(1) = 2
    declaredColumns(2) = 2
    AssertFalse CompareSnapshots(baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns, changes, errorInfo), "Duplicate declared columns must fail."
    AssertTrue IsEmpty(changes), "Invalid target columns must not return partial changes."
End Sub

Private Sub TestEmptySnapshots()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    baseValues = Empty
    localValues = Empty
    baseFonts = Empty
    localFonts = Empty
    baseFills = Empty
    localFills = Empty
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 2
    AssertTrue CompareSnapshots(baseValues, localValues, baseFonts, localFonts, baseFills, localFills, declaredColumns, changes, errorInfo), "Two empty snapshots should compare successfully."
    AssertTrue IsEmpty(changes), "Two empty snapshots should produce no changes."
End Sub

Private Sub MakeBasicSnapshots(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef baseFonts As Variant, ByRef localFonts As Variant, ByRef baseFills As Variant, ByRef localFills As Variant, ByRef declaredColumns As Variant)
    Dim normalFont As String
    Dim blankValue As Variant
    Dim rowIndex As Long

    ReDim baseValues(2 To 3, 1 To 4)
    ReDim localValues(2 To 3, 1 To 4)
    ReDim baseFonts(2 To 3, 1 To 4)
    ReDim localFonts(2 To 3, 1 To 4)
    ReDim baseFills(2 To 3, 1 To 4)
    ReDim localFills(2 To 3, 1 To 4)
    ReDim declaredColumns(1 To 2)
    declaredColumns(1) = 2
    declaredColumns(2) = 4
    normalFont = TestFontSignature(False, 0)
    baseValues(2, 2) = "same"
    baseValues(2, 4) = "old"
    localValues = baseValues
    For rowIndex = 2 To 3
        baseFonts(rowIndex, 2) = TestTextFontData(baseValues(rowIndex, 2), normalFont)
        baseFonts(rowIndex, 4) = TestTextFontData(baseValues(rowIndex, 4), normalFont)
        localFonts(rowIndex, 2) = baseFonts(rowIndex, 2)
        localFonts(rowIndex, 4) = baseFonts(rowIndex, 4)
        baseFills(rowIndex, 2) = TestFillData(0)
        baseFills(rowIndex, 4) = TestFillData(0)
        localFills(rowIndex, 2) = baseFills(rowIndex, 2)
        localFills(rowIndex, 4) = baseFills(rowIndex, 4)
    Next rowIndex
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 1902, "TestMdlDiff", message
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, ByVal message As String)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 1903, "TestMdlDiff", message
End Sub
