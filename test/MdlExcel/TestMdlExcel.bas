Attribute VB_Name = "TestMdlExcel"
Option Explicit

Public Sub RunMdlExcelTests()
    On Error GoTo Failed
    TestHeartbeatScheduling
    TestPhysicalSnapshotsAndDeclaredWrite
    TestOrphanAndInvalidEditableContent
    Exit Sub

Failed:
    TestSupport.RecordTestFailure "TestMdlExcel: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestHeartbeatScheduling()
    Dim errorInfo As AppErrorInfo

    AssertFalse ScheduleSessionHeartbeat(0, errorInfo), "A non-positive heartbeat interval must be rejected."
    AssertTrue ScheduleSessionHeartbeat(60, errorInfo), "A valid heartbeat callback should be scheduled. " & errorInfo.Message
    CancelSessionHeartbeat
End Sub

Private Sub TestPhysicalSnapshotsAndDeclaredWrite()
    Dim handleId As String
    Dim sheet As Worksheet
    Dim schema As SheetSchema
    Dim snapshot As SnapshotData
    Dim expected As SessionSchema
    Dim targets As Variant
    Dim mergedValues As Variant
    Dim invalidCells As Variant
    Dim sheetNames As Variant
    Dim errorInfo As AppErrorInfo
    Dim valid As Boolean
    Dim anchorBefore As Variant
    Dim unselectedBefore As Variant

    On Error GoTo Failed
    AssertTrue RegisterCurrentTestWorkbook(handleId, errorInfo), "The current test workbook should be registered. " & errorInfo.Message
    Set sheet = ThisWorkbook.Worksheets.Item(1)
    sheet.Name = "MdlExcelFixture"
    AssertTrue ListWorksheetNames(handleId, sheetNames, errorInfo), "Worksheet names should be exposed as strings."
    AssertEqual CStr(sheetNames(1)), "MdlExcelFixture", "The workbook adapter should return the physical worksheet name."
    FillHeaders sheet
    FillFixtureRows sheet
    targets = OneSheetTarget("MdlExcelFixture", "AMG_Value")

    AssertTrue ReadSheetSchema(handleId, "MdlExcelFixture", schema, errorInfo), "The schema should read from the worksheet."
    AssertEqual CStr(schema.AnchorColumnIndexes(1)), "1", "Column A should be the first anchor."
    AssertEqual CStr(schema.AnchorColumnIndexes(2)), "3", "Column C should remain an anchor across an empty header."
    AssertEqual CStr(schema.AnchorColumnIndexes(3)), "5", "Column E should remain an anchor."
    AssertEqual CStr(UBound(schema.AnchorColumnIndexes)), "3", "A product column after E must not extend the anchor set."

    AssertTrue ReadDeclaredSnapshot(handleId, targets, snapshot, errorInfo), "The physical snapshot should be read. " & errorInfo.Message
    AssertEqual CStr(snapshot.LastDataRow), "10", "The later row should define the boundary across an internal blank row."
    AssertTrue IsEmpty(snapshot.Values(3, 4)), "The internal blank row must retain its physical position."
    AssertEqual CStr(snapshot.Values(10, 4)), "VALUE_10", "The physical last row should remain addressable."
    CopySnapshotSchema snapshot, expected
    AssertTrue ValidateStructure(handleId, expected, valid, errorInfo), "The fixture structure should validate."
    AssertTrue valid, "The unchanged product anchors should validate."

    anchorBefore = sheet.Cells(4, 1).Value2
    unselectedBefore = sheet.Cells(4, 2).Value2
    mergedValues = snapshot.Values
    mergedValues(4, 4) = "LOCAL_UPDATE"
    AssertTrue WriteDeclaredColumns(handleId, targets, mergedValues, errorInfo), "The declared AMG column should be written. " & errorInfo.Message
    AssertEqual CStr(sheet.Cells(4, 4).Value2), "LOCAL_UPDATE", "The selected AMG cell should contain the merged value."
    AssertEqual CStr(sheet.Cells(4, 1).Value2), CStr(anchorBefore), "The product anchor must not be written."
    AssertEqual CStr(sheet.Cells(4, 2).Value2), CStr(unselectedBefore), "An undeclared AMG column must not be written."

    mergedValues(4, 4) = 7
    AssertFalse WriteDeclaredColumns(handleId, targets, mergedValues, errorInfo), "A non-string merge value must be rejected before writing."
    AssertEqual CStr(sheet.Cells(4, 4).Value2), "LOCAL_UPDATE", "Rejected output must leave the workbook value unchanged."

    sheet.Cells(4, 5).Value2 = "ANCHOR_CHANGED"
    AssertTrue ValidateStructure(handleId, expected, valid, errorInfo), "A changed anchor is a structure result, not a read failure."
    AssertFalse valid, "A changed product anchor must block a commit."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_EXCEL_STRUCTURE), "Anchor drift should have a distinct error."
    sheet.Cells(4, 5).Value2 = "PRODUCT_4"

    CloseWorkbookHandle handleId, False
    Exit Sub

Failed:
    If Len(handleId) > 0 Then CloseWorkbookHandle handleId, False
    Err.Raise Err.Number, "TestMdlExcel.TestPhysicalSnapshotsAndDeclaredWrite", Err.Description
End Sub

Private Sub TestOrphanAndInvalidEditableContent()
    Dim handleId As String
    Dim sheet As Worksheet
    Dim orphanSheet As Worksheet
    Dim targets As Variant
    Dim invalidCells As Variant
    Dim snapshot As SnapshotData
    Dim errorInfo As AppErrorInfo

    On Error GoTo Failed
    AssertTrue RegisterCurrentTestWorkbook(handleId, errorInfo), "The current test workbook should be registered."
    Set sheet = ThisWorkbook.Worksheets.Item("MdlExcelFixture")
    targets = OneSheetTarget("MdlExcelFixture", "AMG_Value")
    sheet.Cells(6, 4).Formula = "=""FORMULA"""
    AssertFalse ValidateEditableContents(handleId, targets, invalidCells, errorInfo), "Formulas in an editable column must be rejected."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_EXCEL_INVALID_CONTENTS), "The formula failure should be distinguishable."
    sheet.Cells(6, 4).ClearContents

    sheet.Cells(6, 4).Value2 = 12
    AssertFalse ValidateEditableContents(handleId, targets, invalidCells, errorInfo), "Non-string values in an editable column must be rejected."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_EXCEL_INVALID_CONTENTS), "The value-type failure should be distinguishable."
    sheet.Cells(6, 4).ClearContents

    Set orphanSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets.Item(ThisWorkbook.Worksheets.Count))
    orphanSheet.Name = "OrphanFixture"
    FillHeaders orphanSheet
    orphanSheet.Cells(2, 1).Value2 = "KEY_2"
    orphanSheet.Cells(2, 3).Value2 = "GROUP_2"
    orphanSheet.Cells(2, 5).Value2 = "PRODUCT_2"
    orphanSheet.Cells(30, 4).Value2 = "UNANCHORED"
    targets = OneSheetTarget("OrphanFixture", "AMG_Value")
    AssertFalse ReadDeclaredSnapshot(handleId, targets, snapshot, errorInfo), "An AMG-only tail row must fail closed."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_EXCEL_ORPHAN_DATA), "The orphan-row failure should be distinguishable."
    CloseWorkbookHandle handleId, False
    Exit Sub

Failed:
    If Len(handleId) > 0 Then CloseWorkbookHandle handleId, False
    Err.Raise Err.Number, "TestMdlExcel.TestOrphanAndInvalidEditableContent", Err.Description
End Sub

Private Sub FillHeaders(ByRef sheet As Worksheet)
    sheet.Cells(1, 1).Value2 = "Product_A"
    sheet.Cells(1, 2).Value2 = "AMG_Notes"
    sheet.Cells(1, 3).Value2 = "Product_C"
    sheet.Cells(1, 4).Value2 = "AMG_Value"
    sheet.Cells(1, 5).Value2 = "Product_E"
    sheet.Cells(1, 6).Value2 = "Product_F"
End Sub

Private Sub FillFixtureRows(ByRef sheet As Worksheet)
    Dim rowIndex As Variant
    For Each rowIndex In Array(2, 4, 10)
        sheet.Cells(CLng(rowIndex), 1).Value2 = "KEY_" & CStr(rowIndex)
        sheet.Cells(CLng(rowIndex), 2).Value2 = "NOTE_" & CStr(rowIndex)
        sheet.Cells(CLng(rowIndex), 3).Value2 = "GROUP_" & CStr(rowIndex)
        sheet.Cells(CLng(rowIndex), 4).Value2 = "VALUE_" & CStr(rowIndex)
        sheet.Cells(CLng(rowIndex), 5).Value2 = "PRODUCT_" & CStr(rowIndex)
    Next rowIndex
End Sub

Private Function OneSheetTarget(ByVal sheetName As String, ByVal amgHeader As String) As Variant
    Dim result(1 To 1, 1 To 2) As Variant
    result(1, 1) = sheetName
    result(1, 2) = amgHeader
    OneSheetTarget = result
End Function

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2401, "TestMdlExcel", message
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, ByVal message As String)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 2402, "TestMdlExcel", message
End Sub
