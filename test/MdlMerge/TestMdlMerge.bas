Attribute VB_Name = "TestMdlMerge"
Option Explicit

Public Sub RunMdlMergeTests()
    RunTestCase "AutoMerge"
    RunTestCase "ConflictChoices"
    RunTestCase "IncompleteChoices"
    RunTestCase "InvalidChoices"
    RunTestCase "InvalidType"
    RunTestCase "MismatchedShape"
    RunTestCase "EmptySnapshots"
End Sub

Private Sub RunTestCase(ByVal caseName As String)
    On Error GoTo Failed
    Select Case caseName
        Case "AutoMerge"
            TestAutoMerge
        Case "ConflictChoices"
            TestConflictChoices
        Case "IncompleteChoices"
            TestIncompleteChoices
        Case "InvalidChoices"
            TestInvalidChoices
        Case "InvalidType"
            TestInvalidType
        Case "MismatchedShape"
            TestMismatchedShape
        Case "EmptySnapshots"
            TestEmptySnapshots
        Case Else
            Err.Raise vbObjectError + 2001, "TestMdlMerge", "Unknown test case: " & caseName
    End Select
    Exit Sub

Failed:
    TestSupport.RecordTestFailure "TestMdlMerge." & caseName & ": " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestAutoMerge()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim declaredColumns As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    MakeSnapshots baseValues, localValues, remoteValues, declaredColumns
    localValues(2, 2) = "local"
    remoteValues(2, 3) = "remote"
    localValues(3, 2) = "same-change"
    remoteValues(3, 2) = "same-change"

    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, declaredColumns, result, errorInfo), "Independent edits and equal edits should merge."
    AssertTrue IsMergeResolved(result), "A merge without divergent same-cell edits should be resolved."
    AssertEqual CStr(result.MergedValues(2, 2)), "local", "A Local-only edit should be kept."
    AssertEqual CStr(result.MergedValues(2, 3)), "remote", "A Remote-only edit should be kept."
    AssertEqual CStr(result.MergedValues(3, 2)), "same-change", "Equal edits should be kept once."
    AssertEqual CStr(result.MergedValues(3, 3)), "base", "Unchanged declared values should remain Remote."
End Sub

Private Sub TestConflictChoices()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim declaredColumns As Variant
    Dim choices As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    MakeSnapshots baseValues, localValues, remoteValues, declaredColumns
    localValues(2, 2) = "local-one"
    remoteValues(2, 2) = "remote-one"
    localValues(3, 3) = "local-two"
    remoteValues(3, 3) = "remote-two"

    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, declaredColumns, result, errorInfo), "Divergent edits should produce a conflict result."
    AssertFalse IsMergeResolved(result), "Conflicts must remain unresolved until every choice is supplied."
    AssertEqual CStr(result.ConflictCount), "2", "Both divergent cells should be listed."
    AssertEqual CStr(result.Conflicts(1, 1)), "2", "The first conflict should preserve its physical row."
    AssertEqual CStr(result.Conflicts(1, 2)), "2", "The first conflict should preserve its physical column."

    ReDim choices(1 To 2, 1 To 3)
    choices(1, 1) = 2
    choices(1, 2) = 2
    choices(1, 3) = "Local"
    choices(2, 1) = 3
    choices(2, 2) = 3
    choices(2, 3) = "Remote"
    AssertTrue ApplyConflictChoices(result, choices, errorInfo), "A complete choice set should resolve all conflicts."
    AssertTrue IsMergeResolved(result), "All conflict choices should mark the result resolved."
    AssertEqual CStr(result.MergedValues(2, 2)), "local-one", "The Local choice should be applied to its cell."
    AssertEqual CStr(result.MergedValues(3, 3)), "remote-two", "The Remote choice should be applied to its cell."
End Sub

Private Sub TestIncompleteChoices()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim declaredColumns As Variant
    Dim choices As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo
    Dim provisionalValue As String

    MakeSnapshots baseValues, localValues, remoteValues, declaredColumns
    localValues(2, 2) = "local"
    remoteValues(2, 2) = "remote"
    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, declaredColumns, result, errorInfo), "The conflict fixture should merge to a pending result."
    provisionalValue = CStr(result.MergedValues(2, 2))
    choices = Empty

    AssertFalse ApplyConflictChoices(result, choices, errorInfo), "Missing choices must not resolve a conflict."
    AssertFalse IsMergeResolved(result), "An incomplete choice set must leave the result unresolved."
    AssertEqual CStr(result.MergedValues(2, 2)), provisionalValue, "Incomplete choices must not apply a partial value."
End Sub

Private Sub TestInvalidChoices()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim declaredColumns As Variant
    Dim choices As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    MakeSnapshots baseValues, localValues, remoteValues, declaredColumns
    localValues(2, 2) = "local"
    remoteValues(2, 2) = "remote"
    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, declaredColumns, result, errorInfo), "The invalid-choice fixture should create a conflict."
    ReDim choices(1 To 1, 1 To 3)
    choices(1, 1) = 2
    choices(1, 2) = 2
    choices(1, 3) = "Both"

    AssertFalse ApplyConflictChoices(result, choices, errorInfo), "Unknown choice values must be rejected."
    AssertFalse IsMergeResolved(result), "Invalid choices must leave the merge unresolved."
End Sub

Private Sub TestInvalidType()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim declaredColumns As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    MakeSnapshots baseValues, localValues, remoteValues, declaredColumns
    baseValues(2, 2) = "same"
    localValues(2, 2) = "same"
    remoteValues(2, 2) = 1

    AssertFalse MergeThreeWay(baseValues, localValues, remoteValues, declaredColumns, result, errorInfo), "Non-string declared values must fail."
    AssertTrue IsEmpty(result.MergedValues), "Invalid values must not produce a writable result."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_INVALID_INPUT), "Cell type failures should retain the shared input error code."
End Sub

Private Sub TestMismatchedShape()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim declaredColumns As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    ReDim baseValues(2 To 3, 1 To 3)
    ReDim localValues(2 To 3, 1 To 3)
    ReDim remoteValues(2 To 4, 1 To 3)
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 2

    AssertFalse MergeThreeWay(baseValues, localValues, remoteValues, declaredColumns, result, errorInfo), "Different snapshot row bounds must fail."
    AssertTrue IsEmpty(result.MergedValues), "Shape failures must not leave a writable partial result."
End Sub

Private Sub TestEmptySnapshots()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim declaredColumns As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    baseValues = Empty
    localValues = Empty
    remoteValues = Empty
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 2

    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, declaredColumns, result, errorInfo), "Three empty snapshots should merge."
    AssertTrue IsMergeResolved(result), "An empty merge has no pending conflicts."
    AssertTrue IsEmpty(result.MergedValues), "The empty merge should return an Empty value set."
End Sub

Private Sub MakeSnapshots(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef remoteValues As Variant, ByRef declaredColumns As Variant)
    ReDim baseValues(2 To 3, 1 To 3)
    ReDim localValues(2 To 3, 1 To 3)
    ReDim remoteValues(2 To 3, 1 To 3)
    ReDim declaredColumns(1 To 2)
    declaredColumns(1) = 2
    declaredColumns(2) = 3
    baseValues(2, 2) = "base"
    baseValues(2, 3) = "base"
    baseValues(3, 2) = "base"
    baseValues(3, 3) = "base"
    localValues = baseValues
    remoteValues = baseValues
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2002, "TestMdlMerge", message
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, ByVal message As String)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 2003, "TestMdlMerge", message
End Sub
