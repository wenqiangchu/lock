Attribute VB_Name = "TestMdlUI"
Option Explicit

Public Sub RunMdlUITests()
    TestRibbonEnabledMapping
    TestConflictChoiceMatrix
    TestConflictChoiceMatrixRejectsIncompleteSelections
    TestConflictValueFormatting
End Sub

Private Sub TestRibbonEnabledMapping()
    Dim view As SessionView

    view.State = SessionIdle
    AssertTrue IsRibbonButtonEnabled("btnStartEdit", view), "Start should be enabled while Idle."
    AssertFalse IsRibbonButtonEnabled("btnSubmit", view), "Submit should be disabled while Idle."
    AssertFalse IsRibbonButtonEnabled("btnAbandon", view), "Abandon should be disabled while Idle."
    AssertFalse IsRibbonButtonEnabled("unknown", view), "Unknown Ribbon controls should be disabled."

    view.State = SessionCheckedOut
    view.CanSubmit = True
    view.CanAbandon = True
    AssertFalse IsRibbonButtonEnabled("btnStartEdit", view), "Start should be disabled during an active session."
    AssertTrue IsRibbonButtonEnabled("btnSubmit", view), "Submit should reflect the session permission."
    AssertTrue IsRibbonButtonEnabled("btnAbandon", view), "Abandon should reflect the session permission."
End Sub

Private Sub TestConflictChoiceMatrix()
    Dim conflicts(0 To 1, 1 To 6) As Variant
    Dim selections(4 To 5) As Variant
    Dim choices As Variant

    conflicts(0, 1) = "SheetA"
    conflicts(0, 2) = 4
    conflicts(0, 3) = 3
    conflicts(0, 4) = "base-a"
    conflicts(0, 5) = "local-a"
    conflicts(0, 6) = "remote-a"
    conflicts(1, 1) = "SheetB"
    conflicts(1, 2) = 9
    conflicts(1, 3) = 5
    conflicts(1, 4) = Empty
    conflicts(1, 5) = "local-b"
    conflicts(1, 6) = "remote-b"
    selections(4) = "Local"
    selections(5) = "Remote"

    AssertTrue BuildConflictChoiceMatrix(conflicts, selections, choices), "A complete conflict selection should build a session choice matrix."
    AssertEqual CStr(LBound(choices, 1)), "1", "Choice output should use one-based rows."
    AssertEqual CStr(UBound(choices, 1)), "2", "Every conflict should have one choice row."
    AssertEqual CStr(choices(1, 1)), "SheetA", "The sheet name should be preserved."
    AssertEqual CStr(choices(1, 2)), "4", "The physical row should be preserved."
    AssertEqual CStr(choices(1, 3)), "3", "The physical column should be preserved."
    AssertEqual CStr(choices(1, 4)), "Local", "The first choice should be preserved."
    AssertEqual CStr(choices(2, 4)), "Remote", "The second choice should be preserved."
End Sub

Private Sub TestConflictChoiceMatrixRejectsIncompleteSelections()
    Dim conflicts(1 To 1, 1 To 6) As Variant
    Dim selections(1 To 1) As Variant
    Dim choices As Variant

    conflicts(1, 1) = "SheetA"
    conflicts(1, 2) = 4
    conflicts(1, 3) = 3
    conflicts(1, 4) = "base"
    conflicts(1, 5) = "local"
    conflicts(1, 6) = "remote"
    selections(1) = vbNullString
    AssertFalse BuildConflictChoiceMatrix(conflicts, selections, choices), "An unanswered conflict must not create commit choices."
    AssertTrue IsEmpty(choices), "Invalid selections must not return a partial matrix."

    selections(1) = "Both"
    AssertFalse BuildConflictChoiceMatrix(conflicts, selections, choices), "Only Local or Remote is a valid choice."
    selections(1) = "Local"
    conflicts(1, 2) = 4.5
    AssertFalse BuildConflictChoiceMatrix(conflicts, selections, choices), "A non-integer physical row must be rejected."
End Sub

Private Sub TestConflictValueFormatting()
    AssertEqual FormatConflictValue(Empty), "(blank)", "Empty cells should be recognizable in the conflict view."
    AssertEqual FormatConflictValue("first" & vbCrLf & "second"), "first" & vbLf & "second", "Line endings should be normalized without changing content."
    AssertEqual FormatConflictValue(String$(5000, "x")), String$(5000, "x"), "Long values must be displayed without truncation."
    AssertEqual ExcelColumnLabel(1), "A", "The first physical column should use Excel address notation."
    AssertEqual ExcelColumnLabel(5), "E", "Anchor columns should retain their letter addresses."
    AssertEqual ExcelColumnLabel(27), "AA", "Column labels should handle multi-letter addresses."
    AssertEqual ExcelColumnLabel(0), vbNullString, "Invalid column indexes should not produce a misleading address."
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2802, "TestMdlUI", message
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, ByVal message As String)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 2803, "TestMdlUI", message
End Sub
