Attribute VB_Name = "TestMdlDiff"
Option Explicit

Public Sub RunMdlDiffTests()
    RunTestCase "NoChanges"
    RunTestCase "DeclaredChangesOnly"
    RunTestCase "ExactStringAndEmptyComparison"
    RunTestCase "InvalidDeclaredValue"
    RunTestCase "MismatchedDimensions"
    RunTestCase "InvalidDeclaredColumns"
    RunTestCase "EmptySnapshots"
End Sub

Private Sub RunTestCase(ByVal caseName As String)
    On Error GoTo Failed
    Select Case caseName
        Case "NoChanges"
            TestNoChanges
        Case "DeclaredChangesOnly"
            TestDeclaredChangesOnly
        Case "ExactStringAndEmptyComparison"
            TestExactStringAndEmptyComparison
        Case "InvalidDeclaredValue"
            TestInvalidDeclaredValue
        Case "MismatchedDimensions"
            TestMismatchedDimensions
        Case "InvalidDeclaredColumns"
            TestInvalidDeclaredColumns
        Case "EmptySnapshots"
            TestEmptySnapshots
        Case Else
            Err.Raise vbObjectError + 1901, "TestMdlDiff", "Unknown test case: " & caseName
    End Select
    Exit Sub

Failed:
    TestSupport.RecordTestFailure "TestMdlDiff." & caseName & ": " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestNoChanges()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    MakeBasicSnapshots baseValues, localValues, declaredColumns
    AssertTrue CompareSnapshots(baseValues, localValues, declaredColumns, changes, errorInfo), "Equal snapshots should compare successfully."
    AssertTrue IsEmpty(changes), "Equal snapshots should return no change records."
End Sub

Private Sub TestDeclaredChangesOnly()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    ReDim baseValues(2 To 3, 1 To 4)
    ReDim localValues(2 To 3, 1 To 4)
    ReDim declaredColumns(1 To 2)
    declaredColumns(1) = 2
    declaredColumns(2) = 4
    baseValues(2, 2) = "old"
    localValues(2, 2) = "new"
    baseValues(3, 4) = Empty
    localValues(3, 4) = "added"
    baseValues(2, 3) = 25
    localValues(2, 3) = 99

    AssertTrue CompareSnapshots(baseValues, localValues, declaredColumns, changes, errorInfo), "Declared string changes should be reported."
    AssertEqual CStr(UBound(changes, 1)), "2", "Only the two declared-column changes should be included."
    AssertEqual CStr(changes(1, 1)), "2", "The first change should preserve its physical row number."
    AssertEqual CStr(changes(1, 2)), "2", "The first change should preserve its physical column number."
    AssertEqual CStr(changes(2, 1)), "3", "The second change should preserve its physical row number."
    AssertEqual CStr(changes(2, 2)), "4", "The second change should preserve its physical column number."
    AssertTrue IsEmpty(changes(2, 3)), "The Base blank should remain an Empty sentinel."
    AssertEqual CStr(changes(2, 4)), "added", "The Local value should be preserved."
End Sub

Private Sub TestExactStringAndEmptyComparison()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    ReDim baseValues(2 To 4, 1 To 1)
    ReDim localValues(2 To 4, 1 To 1)
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 1
    baseValues(2, 1) = "Case"
    localValues(2, 1) = "case"
    baseValues(3, 1) = "text "
    localValues(3, 1) = "text"
    baseValues(4, 1) = Empty
    localValues(4, 1) = ""

    AssertTrue CompareSnapshots(baseValues, localValues, declaredColumns, changes, errorInfo), "String changes should compare successfully."
    AssertEqual CStr(UBound(changes, 1)), "3", "Case, whitespace, and Empty/text differences should all be detected."
End Sub

Private Sub TestInvalidDeclaredValue()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    ReDim baseValues(2 To 2, 1 To 1)
    ReDim localValues(2 To 2, 1 To 1)
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 1
    baseValues(2, 1) = 1
    localValues(2, 1) = 1

    AssertFalse CompareSnapshots(baseValues, localValues, declaredColumns, changes, errorInfo), "Non-string declared values must fail even when equal."
    AssertTrue IsEmpty(changes), "Invalid values must not return partial change records."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_INVALID_INPUT), "Invalid cell values should retain the shared input error code."
End Sub

Private Sub TestMismatchedDimensions()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    ReDim baseValues(2 To 3, 1 To 2)
    ReDim localValues(2 To 4, 1 To 2)
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 1

    AssertFalse CompareSnapshots(baseValues, localValues, declaredColumns, changes, errorInfo), "Different physical row bounds must fail."
    AssertTrue IsEmpty(changes), "A dimension error must not return partial changes."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_DIFF_INVALID_INPUT), "Dimension errors should have the diff input error code."
End Sub

Private Sub TestInvalidDeclaredColumns()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    ReDim baseValues(2 To 2, 1 To 2)
    ReDim localValues(2 To 2, 1 To 2)
    ReDim declaredColumns(1 To 2)
    declaredColumns(1) = 1
    declaredColumns(2) = 1

    AssertFalse CompareSnapshots(baseValues, localValues, declaredColumns, changes, errorInfo), "Duplicate declared columns must fail."
    AssertTrue IsEmpty(changes), "Invalid target columns must not return partial changes."
End Sub

Private Sub TestEmptySnapshots()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim declaredColumns As Variant
    Dim changes As Variant
    Dim errorInfo As AppErrorInfo

    baseValues = Empty
    localValues = Empty
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 2

    AssertTrue CompareSnapshots(baseValues, localValues, declaredColumns, changes, errorInfo), "Two empty snapshots should compare successfully."
    AssertTrue IsEmpty(changes), "Two empty snapshots should produce no changes."
End Sub

Private Sub MakeBasicSnapshots(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef declaredColumns As Variant)
    ReDim baseValues(2 To 3, 1 To 3)
    ReDim localValues(2 To 3, 1 To 3)
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 2
    baseValues(2, 2) = "same"
    localValues(2, 2) = "same"
    baseValues(3, 2) = Empty
    localValues(3, 2) = Empty
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
