Attribute VB_Name = "TestMdlMerge"
Option Explicit

Public Sub RunMdlMergeTests()
    RunTestCase "AutoMergeComponents"
    RunTestCase "SingleCellConflict"
    RunTestCase "SeparateComponentConflicts"
    RunTestCase "InvalidConflictDataIsAtomic"
    RunTestCase "IncompleteChoices"
    RunTestCase "InvalidChoices"
    RunTestCase "InvalidType"
    RunTestCase "MismatchedShape"
    RunTestCase "EmptySnapshots"
End Sub

Private Sub RunTestCase(ByVal caseName As String)
    On Error GoTo Failed
    Select Case caseName
        Case "AutoMergeComponents": TestAutoMergeComponents
        Case "SingleCellConflict": TestSingleCellConflict
        Case "SeparateComponentConflicts": TestSeparateComponentConflicts
        Case "InvalidConflictDataIsAtomic": TestInvalidConflictDataIsAtomic
        Case "IncompleteChoices": TestIncompleteChoices
        Case "InvalidChoices": TestInvalidChoices
        Case "InvalidType": TestInvalidType
        Case "MismatchedShape": TestMismatchedShape
        Case "EmptySnapshots": TestEmptySnapshots
        Case Else: Err.Raise vbObjectError + 2001, "TestMdlMerge", "Unknown test case: " & caseName
    End Select
    Exit Sub
Failed:
    TestSupport.RecordTestFailure "TestMdlMerge." & caseName & ": " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestAutoMergeComponents()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim remoteFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim remoteFills As Variant
    Dim declaredColumns As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    MakeSnapshots baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns
    SetTestValue localValues, localFonts, 2, 2, "local-only", True, 0
    SetTestValue remoteValues, remoteFonts, 2, 3, "remote-only", False, 0
    localValues(3, 2) = "same-change"
    remoteValues(3, 2) = "same-change"
    localFonts(3, 2) = TestTextFontData(localValues(3, 2), TestFontSignature(False, 0))
    remoteFonts(3, 2) = localFonts(3, 2)
    localFills(2, 3) = baseFills(2, 3)
    remoteFills(2, 3) = TestFillData(65280)

    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns, result, errorInfo), "Independent value and fill edits should merge. " & errorInfo.Message
    AssertTrue IsMergeResolved(result), "A merge without divergent same-component edits should be resolved."
    AssertEqual CStr(result.MergedValues(2, 2)), "local-only", "A Local-only text edit should be kept."
    AssertEqual CStr(result.MergedValues(2, 3)), "remote-only", "A Remote-only text edit should be kept."
    AssertEqual CStr(result.MergedValues(3, 2)), "same-change", "Equal text edits should be kept once."
    AssertEqual CStr(result.MergedFills(2, 3)), CStr(remoteFills(2, 3)), "A Remote-only fill edit should be merged independently."
End Sub

Private Sub TestSingleCellConflict()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim remoteFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim remoteFills As Variant
    Dim declaredColumns As Variant
    Dim choices As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    ReDim baseValues(2 To 2, 2 To 2)
    ReDim localValues(2 To 2, 2 To 2)
    ReDim remoteValues(2 To 2, 2 To 2)
    ReDim baseFonts(2 To 2, 2 To 2)
    ReDim localFonts(2 To 2, 2 To 2)
    ReDim remoteFonts(2 To 2, 2 To 2)
    ReDim baseFills(2 To 2, 2 To 2)
    ReDim localFills(2 To 2, 2 To 2)
    ReDim remoteFills(2 To 2, 2 To 2)
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 2
    baseValues(2, 2) = "base"
    localValues(2, 2) = "local"
    remoteValues(2, 2) = "remote"
    baseFonts(2, 2) = TestTextFontData(baseValues(2, 2), TestFontSignature(False, 0))
    localFonts(2, 2) = TestTextFontData(localValues(2, 2), TestFontSignature(True, 255))
    remoteFonts(2, 2) = TestTextFontData(remoteValues(2, 2), TestFontSignature(False, 16711680))
    baseFills(2, 2) = TestFillData(0)
    localFills(2, 2) = baseFills(2, 2)
    remoteFills(2, 2) = baseFills(2, 2)

    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns, result, errorInfo), "A one-cell matrix should merge and expose its divergent text/font component."
    AssertEqual CStr(result.ConflictCount), "1", "Only the single divergent text/font component should conflict."
    ReDim choices(1 To 1, 1 To 4)
    choices(1, 1) = 2
    choices(1, 2) = 2
    choices(1, 3) = "TextFont"
    choices(1, 4) = "Remote"
    AssertTrue ApplyConflictChoices(result, choices, errorInfo), "A one-item choice matrix should resolve the conflict."
    AssertEqual CStr(result.MergedValues(2, 2)), "remote", "The selected text/font component should supply the corresponding text value."
End Sub

Private Sub TestSeparateComponentConflicts()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim remoteFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim remoteFills As Variant
    Dim declaredColumns As Variant
    Dim choices As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    MakeSnapshots baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns
    localFonts(2, 2) = TestTextFontData(localValues(2, 2), TestFontSignature(True, 255))
    remoteFonts(2, 2) = TestTextFontData(remoteValues(2, 2), TestFontSignature(False, 16711680))
    localFills(2, 2) = TestFillData(255)
    remoteFills(2, 2) = TestFillData(16711680)

    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns, result, errorInfo), "Independent font and fill conflicts should be listed."
    AssertFalse IsMergeResolved(result), "Component conflicts must remain pending before user choices."
    AssertEqual CStr(result.ConflictCount), "2", "Font and fill must create separate conflicts for the same cell."
    AssertEqual CStr(result.Conflicts(1, 3)), "TextFont", "The first conflict should identify TextFont."
    AssertEqual CStr(result.Conflicts(2, 3)), "Fill", "The second conflict should identify Fill."

    ReDim choices(1 To 2, 1 To 4)
    choices(1, 1) = 2
    choices(1, 2) = 2
    choices(1, 3) = "TextFont"
    choices(1, 4) = "Local"
    choices(2, 1) = 2
    choices(2, 2) = 2
    choices(2, 3) = "TextFont"
    choices(2, 4) = "Remote"
    AssertFalse ApplyConflictChoices(result, choices, errorInfo), "Duplicate component choices must not resolve a separate conflict."
    AssertFalse IsMergeResolved(result), "Duplicate choices must leave all conflicts pending."
    choices(2, 3) = "Fill"
    AssertTrue ApplyConflictChoices(result, choices, errorInfo), "Each component choice should resolve only its matching component."
    AssertTrue IsMergeResolved(result), "Every selected component conflict should be resolved."
    AssertEqual CStr(result.MergedValues(2, 2)), CStr(localValues(2, 2)), "The Local TextFont choice should also select its text value."
    AssertEqual CStr(result.MergedFonts(2, 2)), CStr(localFonts(2, 2)), "The Local font mapping should be kept."
    AssertEqual CStr(result.MergedFills(2, 2)), CStr(remoteFills(2, 2)), "The Remote fill should be kept independently."
End Sub

Private Sub TestInvalidConflictDataIsAtomic()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim remoteFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim remoteFills As Variant
    Dim declaredColumns As Variant
    Dim choices As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo
    Dim originalValue As String
    Dim originalFont As String
    Dim originalFill As String

    MakeSnapshots baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns
    localFills(2, 2) = TestFillData(255)
    remoteFills(2, 2) = TestFillData(16711680)
    SetTestValue localValues, localFonts, 3, 2, "local", True, 1
    SetTestValue remoteValues, remoteFonts, 3, 2, "remote", False, 2
    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns, result, errorInfo), "The fixture should create an early fill and later text/font conflict."
    AssertEqual CStr(result.Conflicts(1, 3)), "Fill", "The first conflict should be the earlier row fill."
    AssertEqual CStr(result.Conflicts(2, 3)), "TextFont", "The second conflict should be the later text/font component."

    originalValue = CStr(result.MergedValues(3, 2))
    originalFont = CStr(result.MergedFonts(3, 2))
    originalFill = CStr(result.MergedFills(2, 2))
    result.Conflicts(2, 6) = "TF1|malformed"
    ReDim choices(1 To 2, 1 To 4)
    choices(1, 1) = 2
    choices(1, 2) = 2
    choices(1, 3) = "Fill"
    choices(1, 4) = "Local"
    choices(2, 1) = 3
    choices(2, 2) = 2
    choices(2, 3) = "TextFont"
    choices(2, 4) = "Remote"

    AssertFalse ApplyConflictChoices(result, choices, errorInfo), "Malformed later conflict data must reject the entire choice set."
    AssertFalse IsMergeResolved(result), "Malformed conflict data must leave the merge pending."
    AssertEqual CStr(result.MergedValues(3, 2)), originalValue, "A rejected choice set must not change values."
    AssertEqual CStr(result.MergedFonts(3, 2)), originalFont, "A rejected choice set must not change fonts."
    AssertEqual CStr(result.MergedFills(2, 2)), originalFill, "A rejected choice set must not apply an earlier fill choice."

    result.Conflicts(2, 6) = remoteFonts(3, 2)
    result.Conflicts(2, 1) = 999
    choices(2, 1) = 999
    AssertFalse ApplyConflictChoices(result, choices, errorInfo), "An out-of-range conflict coordinate must be rejected before choices are applied."
    AssertEqual CStr(result.MergedFills(2, 2)), originalFill, "An invalid coordinate must not apply any choice."
End Sub

Private Sub TestIncompleteChoices()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim remoteFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim remoteFills As Variant
    Dim declaredColumns As Variant
    Dim choices As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo
    Dim provisionalValue As String

    MakeSnapshots baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns
    SetTestValue localValues, localFonts, 2, 2, "local", True, 1
    SetTestValue remoteValues, remoteFonts, 2, 2, "remote", False, 2
    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns, result, errorInfo), "The fixture should create two conflicts."
    provisionalValue = CStr(result.MergedValues(2, 2))
    choices = Empty

    AssertFalse ApplyConflictChoices(result, choices, errorInfo), "Missing choices must not resolve any conflict."
    AssertFalse IsMergeResolved(result), "An incomplete choice set must remain unresolved."
    AssertEqual CStr(result.MergedValues(2, 2)), provisionalValue, "Incomplete choices must not apply a partial value."
End Sub

Private Sub TestInvalidChoices()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim remoteFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim remoteFills As Variant
    Dim declaredColumns As Variant
    Dim choices As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    MakeSnapshots baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns
    SetTestValue localValues, localFonts, 2, 2, "local", True, 1
    SetTestValue remoteValues, remoteFonts, 2, 2, "remote", False, 2
    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns, result, errorInfo), "The invalid-choice fixture should create conflicts."
    ReDim choices(1 To 1, 1 To 4)
    choices(1, 1) = 2
    choices(1, 2) = 2
    choices(1, 3) = "Unknown"
    choices(1, 4) = "Local"
    AssertFalse ApplyConflictChoices(result, choices, errorInfo), "Unknown component names must be rejected."
    choices(1, 1) = 2147483648#
    choices(1, 3) = "TextFont"
    AssertFalse ApplyConflictChoices(result, choices, errorInfo), "Conflict coordinates outside the Long range must be rejected safely."
    AssertFalse IsMergeResolved(result), "Invalid choices must leave the merge unresolved."
End Sub

Private Sub TestInvalidType()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim remoteFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim remoteFills As Variant
    Dim declaredColumns As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    MakeSnapshots baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns
    remoteValues(2, 2) = 1
    AssertFalse MergeThreeWay(baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns, result, errorInfo), "Non-string declared values must fail."
    AssertTrue IsEmpty(result.MergedValues), "Invalid values must not produce writable results."
End Sub

Private Sub TestMismatchedShape()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim remoteFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim remoteFills As Variant
    Dim declaredColumns As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    MakeSnapshots baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns
    ReDim remoteFonts(2 To 4, 1 To 3)
    AssertFalse MergeThreeWay(baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns, result, errorInfo), "Different format snapshot bounds must fail."
    AssertTrue IsEmpty(result.MergedValues), "Shape failures must not leave a writable partial result."
End Sub

Private Sub TestEmptySnapshots()
    Dim baseValues As Variant
    Dim localValues As Variant
    Dim remoteValues As Variant
    Dim baseFonts As Variant
    Dim localFonts As Variant
    Dim remoteFonts As Variant
    Dim baseFills As Variant
    Dim localFills As Variant
    Dim remoteFills As Variant
    Dim declaredColumns As Variant
    Dim result As MergeResult
    Dim errorInfo As AppErrorInfo

    baseValues = Empty
    localValues = Empty
    remoteValues = Empty
    baseFonts = Empty
    localFonts = Empty
    remoteFonts = Empty
    baseFills = Empty
    localFills = Empty
    remoteFills = Empty
    ReDim declaredColumns(1 To 1)
    declaredColumns(1) = 2
    AssertTrue MergeThreeWay(baseValues, localValues, remoteValues, baseFonts, localFonts, remoteFonts, baseFills, localFills, remoteFills, declaredColumns, result, errorInfo), "Three empty snapshots should merge."
    AssertTrue IsMergeResolved(result), "An empty merge has no pending conflicts."
    AssertTrue IsEmpty(result.MergedValues), "The empty merge should return no value matrix."
End Sub

Private Sub MakeSnapshots(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef remoteValues As Variant, ByRef baseFonts As Variant, ByRef localFonts As Variant, ByRef remoteFonts As Variant, ByRef baseFills As Variant, ByRef localFills As Variant, ByRef remoteFills As Variant, ByRef declaredColumns As Variant)
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim normalFont As String

    ReDim baseValues(2 To 3, 1 To 3)
    ReDim localValues(2 To 3, 1 To 3)
    ReDim remoteValues(2 To 3, 1 To 3)
    ReDim baseFonts(2 To 3, 1 To 3)
    ReDim localFonts(2 To 3, 1 To 3)
    ReDim remoteFonts(2 To 3, 1 To 3)
    ReDim baseFills(2 To 3, 1 To 3)
    ReDim localFills(2 To 3, 1 To 3)
    ReDim remoteFills(2 To 3, 1 To 3)
    ReDim declaredColumns(1 To 2)
    declaredColumns(1) = 2
    declaredColumns(2) = 3
    normalFont = TestFontSignature(False, 0)
    For rowIndex = 2 To 3
        For columnIndex = 2 To 3
            If rowIndex = 2 Then baseValues(rowIndex, columnIndex) = "base" Else baseValues(rowIndex, columnIndex) = "base-2"
            baseFonts(rowIndex, columnIndex) = TestTextFontData(baseValues(rowIndex, columnIndex), normalFont)
            baseFills(rowIndex, columnIndex) = TestFillData(0)
        Next columnIndex
    Next rowIndex
    localValues = baseValues
    remoteValues = baseValues
    localFonts = baseFonts
    remoteFonts = baseFonts
    localFills = baseFills
    remoteFills = baseFills
End Sub

Private Sub SetTestValue(ByRef values As Variant, ByRef fonts As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal text As String, ByVal isBold As Boolean, ByVal colorValue As Long)
    values(rowIndex, columnIndex) = text
    fonts(rowIndex, columnIndex) = TestTextFontData(text, TestFontSignature(isBold, colorValue))
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
