Attribute VB_Name = "TestMdlExcel"
Option Explicit

Private Const XL_OPEN_XML_WORKBOOK As Long = 51
Private Const XL_WBAT_WORKSHEET As Long = -4167

Public Sub RunMdlExcelTests()
    On Error GoTo Failed
    TestHeartbeatScheduling
    TestPhysicalSnapshotsAndDeclaredWrite
    TestOrphanAndInvalidEditableContent
    TestRealWorkbookRoundTripAndFakeLock
    Exit Sub

Failed:
    TestSupport.RecordTestFailure "TestMdlExcel: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestRealWorkbookRoundTripAndFakeLock()
    Dim root As String
    Dim workbookPath As String
    Dim lockPath As String
    Dim errorInfo As AppErrorInfo
    Dim cleanupError As AppErrorInfo
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    root = GetLocalAppDataFolder() & "\tests\excel-real-" & NewGuidString()
    AssertTrue EnsureLocalAppDirectory(root, errorInfo), "The real workbook test root should be created below LOCALAPPDATA. " & errorInfo.Message
    workbookPath = root & "\MDDTRealFixture.xlsx"
    CreateRealTestWorkbook workbookPath
    VerifyRealWorkbookRoundTrip workbookPath
    VerifyRealWorkbookLockConflict workbookPath, lockPath
    GoTo Cleanup

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
Cleanup:
    On Error Resume Next
    If Len(lockPath) > 0 Then
        If Not StorageRemoveLockDirectory(lockPath, cleanupError) Then TestSupport.RecordTestFailure "TestMdlExcel.TestRealWorkbookRoundTripAndFakeLock: lock cleanup failed: " & cleanupError.Message & " " & cleanupError.Context
    End If
    If Len(root) > 0 Then
        If Not RemoveLocalWorkspace(root, cleanupError) Then TestSupport.RecordTestFailure "TestMdlExcel.TestRealWorkbookRoundTripAndFakeLock: workspace cleanup failed: " & cleanupError.Message & " " & cleanupError.Context
    End If
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "TestMdlExcel.TestRealWorkbookRoundTripAndFakeLock", failureDescription
End Sub

Private Sub CreateRealTestWorkbook(ByVal workbookPath As String)
    Dim fixture As Workbook
    Dim sheet As Worksheet
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    Set fixture = Application.Workbooks.Add(XL_WBAT_WORKSHEET)
    Set sheet = fixture.Worksheets.Item(1)
    sheet.Name = "MdlExcelFixture"
    FillHeaders sheet
    FillFixtureRows sheet
    fixture.SaveAs Filename:=workbookPath, FileFormat:=XL_OPEN_XML_WORKBOOK
    fixture.Close SaveChanges:=False
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise failureNumber, "TestMdlExcel.CreateRealTestWorkbook", failureDescription
End Sub

Private Sub VerifyRealWorkbookRoundTrip(ByVal workbookPath As String)
    Dim handleId As String
    Dim targets As Variant
    Dim snapshot As SnapshotData
    Dim mergedValues As Variant
    Dim mergedFonts As Variant
    Dim mergedFills As Variant
    Dim parsedKind As String
    Dim parsedText As String
    Dim emptyCellFont As String
    Dim runs As Variant
    Dim errorInfo As AppErrorInfo
    Dim errorInfo2 As AppErrorInfo
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    targets = OneSheetTarget("MdlExcelFixture", "AMG_Value")
    AssertTrue OpenWorkbookHandle(workbookPath, WorkbookSharedRemote, handleId, errorInfo), "MdlExcel should import a real local XLSX through the shared-workbook path. " & errorInfo.Message
    AssertTrue ReadDeclaredSnapshot(handleId, targets, snapshot, errorInfo), "The real XLSX data should be read through MdlExcel. " & errorInfo.Message
    AssertEqual CStr(snapshot.Values(4, 4)), "VALUE_4", "The real workbook import returned an unexpected cell value."
    AssertTrue ParseTextFontData(CStr(snapshot.Fonts(4, 4)), parsedKind, parsedText, runs, emptyCellFont, errorInfo), "The imported text/font snapshot should parse. " & errorInfo.Message

    mergedValues = snapshot.Values
    mergedFonts = snapshot.Fonts
    mergedFills = snapshot.Fills
    mergedValues(4, 4) = "EDITED4"
    mergedFonts(4, 4) = BuildTextFontData(mergedValues(4, 4), runs, emptyCellFont, errorInfo2)
    AssertTrue Len(CStr(mergedFonts(4, 4))) > 0, "The value edit should retain the real workbook's font mapping. " & errorInfo2.Message
    AssertTrue WriteDeclaredColumns(handleId, targets, mergedValues, mergedFonts, mergedFills, errorInfo), "The real XLSX should accept a declared-column edit. " & errorInfo.Message
    AssertTrue SaveWorkbook(handleId, errorInfo), "MdlExcel should persist the edit with Workbook.Save. " & errorInfo.Message
    CloseWorkbookHandle handleId, False
    handleId = vbNullString

    AssertTrue OpenWorkbookHandle(workbookPath, WorkbookSharedRemote, handleId, errorInfo), "The saved XLSX should reopen through MdlExcel. " & errorInfo.Message
    AssertTrue ReadDeclaredSnapshot(handleId, targets, snapshot, errorInfo), "The reopened XLSX snapshot should read. " & errorInfo.Message
    AssertEqual CStr(snapshot.Values(4, 4)), "EDITED4", "The saved cell value did not survive closing and reopening."
    CloseWorkbookHandle handleId, False
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    If Len(handleId) > 0 Then CloseWorkbookHandle handleId, False
    Err.Raise failureNumber, "TestMdlExcel.VerifyRealWorkbookRoundTrip", failureDescription
End Sub

Private Sub VerifyRealWorkbookLockConflict(ByVal workbookPath As String, ByRef lockPath As String)
    Dim owner As CLockCoordinator
    Dim contender As CLockCoordinator
    Dim lockTargets As Variant
    Dim leases As Variant
    Dim contenderLeases As Variant
    Dim targetFields As Variant
    Dim ownerPath As String
    Dim status As LockAvailability
    Dim errorInfo As AppErrorInfo

    Set owner = New CLockCoordinator
    Set contender = New CLockCoordinator
    lockTargets = OneLockTarget(workbookPath, "MdlExcelFixture", "AMG_Value")
    targetFields = LockTargetAt(lockTargets, 1)
    AssertTrue owner.AcquireLocks(lockTargets, "excel-owner", leases, errorInfo), "A lock should be created beside the real test workbook. " & errorInfo.Message
    lockPath = CStr(leases(1, 1))
    ownerPath = lockPath & "\owner.txt"
    AssertTrue InStr(1, lockPath, "\history\mddtlock\MDDTREALFIXTURE.XLSX\", vbTextCompare) > 0, "The real lock should use the readable shared history layout."
    AssertFalse contender.AcquireLocks(lockTargets, "excel-contender", contenderLeases, errorInfo), "A second session should be blocked by the owner file on disk."
    AssertTrue contender.InspectTargetLock(targetFields, status, errorInfo), "The on-disk owner should be inspectable."
    AssertEqual CStr(status), CStr(LockHeld), "A valid owner file should report a held lock."

    AssertTrue StorageWriteText(ownerPath, "damaged fake owner", errorInfo), "The test should be able to replace the temporary owner record."
    AssertTrue contender.InspectTargetLock(targetFields, status, errorInfo), "A damaged fake owner should remain an inspectable lock state."
    AssertEqual CStr(status), CStr(LockUnknown), "A damaged owner file must fail closed as an unknown lock."
    AssertFalse contender.AcquireLocks(lockTargets, "excel-after-damage", contenderLeases, errorInfo), "A damaged fake lock must still block acquisition."
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
    Dim mergedFonts As Variant
    Dim mergedFills As Variant
    Dim parsedKind As String
    Dim parsedText As String
    Dim emptyCellFont As String
    Dim runs As Variant
    Dim errorInfo As AppErrorInfo
    Dim errorInfo2 As AppErrorInfo
    Dim invalidCells As Variant
    Dim sheetNames As Variant
    Dim valid As Boolean
    Dim anchorBefore As Variant
    Dim unselectedBefore As Variant
    Dim currentStep As String
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    currentStep = "register test workbook"
    AssertTrue RegisterCurrentTestWorkbook(handleId, errorInfo), "The current test workbook should be registered. " & errorInfo.Message
    Set sheet = ThisWorkbook.Worksheets.Item(1)
    sheet.Name = "MdlExcelFixture"
    currentStep = "prepare workbook fixture"
    AssertTrue ListWorksheetNames(handleId, sheetNames, errorInfo), "Worksheet names should be exposed as strings."
    AssertEqual CStr(sheetNames(1)), "MdlExcelFixture", "The workbook adapter should return the physical worksheet name."
    FillHeaders sheet
    FillFixtureRows sheet
    sheet.Cells(4, 4).Characters(1, 2).Font.Color = RGB(255, 0, 0)
    sheet.Cells(4, 4).Characters(4, 2).Font.Bold = True
    sheet.Cells(4, 4).Interior.Color = RGB(220, 230, 241)
    targets = OneSheetTarget("MdlExcelFixture", "AMG_Value")

    currentStep = "read sheet schema"
    AssertTrue ReadSheetSchema(handleId, "MdlExcelFixture", schema, errorInfo), "The schema should read from the worksheet."
    AssertEqual CStr(schema.AnchorColumnIndexes(1)), "1", "Column A should be the first anchor."
    AssertEqual CStr(schema.AnchorColumnIndexes(2)), "3", "Column C should remain an anchor across an empty header."
    AssertEqual CStr(schema.AnchorColumnIndexes(3)), "5", "Column E should remain an anchor."
    AssertEqual CStr(UBound(schema.AnchorColumnIndexes)), "3", "A product column after E must not extend the anchor set."

    currentStep = "read value, font, and fill snapshot"
    AssertTrue ReadDeclaredSnapshot(handleId, targets, snapshot, errorInfo), "The physical snapshot should be read. " & errorInfo.Message
    AssertEqual CStr(snapshot.LastDataRow), "10", "The later row should define the boundary across an internal blank row."
    AssertTrue IsEmpty(snapshot.Values(3, 4)), "The internal blank row must retain its physical position."
    AssertEqual CStr(snapshot.Values(10, 4)), "VALUE_10", "The physical last row should remain addressable."
    currentStep = "parse font runs"
    AssertTrue ParseTextFontData(CStr(snapshot.Fonts(4, 4)), parsedKind, parsedText, runs, emptyCellFont, errorInfo), "The font snapshot should preserve localized character runs."
    AssertEqual parsedText, "VALUE_4", "The font data must remain bound to the source string."
    AssertTrue UBound(runs, 1) >= 2, "Distinct character formatting should produce separate runs."
    AssertTrue CompareFormattingData(snapshot.Fills(4, 4), snapshot.Fills(4, 4), valid, errorInfo), "The fill snapshot should use the canonical format protocol."
    CopySnapshotSchema snapshot, expected
    currentStep = "validate snapshot structure"
    AssertTrue ValidateStructure(handleId, expected, valid, errorInfo), "The fixture structure should validate."
    AssertTrue valid, "The unchanged product anchors should validate."

    anchorBefore = sheet.Cells(4, 1).Value2
    unselectedBefore = sheet.Cells(4, 2).Value2
    mergedValues = snapshot.Values
    mergedFonts = snapshot.Fonts
    mergedFills = snapshot.Fills
    mergedValues(4, 4) = "LOCAL04"
    currentStep = "prepare merged text and format data"
    AssertTrue ParseTextFontData(CStr(mergedFonts(4, 4)), parsedKind, parsedText, runs, emptyCellFont, errorInfo), "The source font runs should parse before text replacement."
    mergedFonts(4, 4) = BuildTextFontData(mergedValues(4, 4), runs, emptyCellFont, errorInfo2)
    currentStep = "write and verify merged value and formats"
    AssertTrue WriteDeclaredColumns(handleId, targets, mergedValues, mergedFonts, mergedFills, errorInfo), "The declared AMG value and formats should be written. " & errorInfo.Message & " " & errorInfo.Context
    AssertEqual CStr(sheet.Cells(4, 4).Value2), "LOCAL04", "The selected AMG cell should contain the merged value."
    AssertEqual CStr(sheet.Cells(4, 4).Characters(1, 1).Font.Color), CStr(RGB(255, 0, 0)), "The first localized font run should round-trip."
    AssertEqual CStr(sheet.Cells(4, 4).Characters(4, 1).Font.Bold), "True", "The second localized font run should round-trip."
    AssertEqual CStr(sheet.Cells(4, 4).Interior.Color), CStr(RGB(220, 230, 241)), "The direct fill should round-trip."
    AssertEqual CStr(sheet.Cells(4, 1).Value2), CStr(anchorBefore), "The product anchor must not be written."
    AssertEqual CStr(sheet.Cells(4, 2).Value2), CStr(unselectedBefore), "An undeclared AMG column must not be written."

    currentStep = "reject invalid merged value"
    mergedValues(4, 4) = 7
    AssertFalse WriteDeclaredColumns(handleId, targets, mergedValues, mergedFonts, mergedFills, errorInfo), "A non-string merge value must be rejected before writing."
    AssertEqual CStr(sheet.Cells(4, 4).Value2), "LOCAL04", "Rejected output must leave the workbook value unchanged."

    currentStep = "validate anchor drift"
    sheet.Cells(4, 5).Value2 = "ANCHOR_CHANGED"
    AssertTrue ValidateStructure(handleId, expected, valid, errorInfo), "A changed anchor is a structure result, not a read failure."
    AssertFalse valid, "A changed product anchor must block a commit."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_EXCEL_STRUCTURE), "Anchor drift should have a distinct error."
    sheet.Cells(4, 5).Value2 = "PRODUCT_4"

    CloseWorkbookHandle handleId, False
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    If Len(handleId) > 0 Then CloseWorkbookHandle handleId, False
    Err.Raise failureNumber, "TestMdlExcel.TestPhysicalSnapshotsAndDeclaredWrite", currentStep & ": " & failureDescription
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

Private Function OneLockTarget(ByVal workbookPath As String, ByVal sheetName As String, ByVal amgHeader As String) As Variant
    Dim result(1 To 1, 1 To 3) As Variant
    result(1, 1) = workbookPath
    result(1, 2) = sheetName
    result(1, 3) = amgHeader
    OneLockTarget = result
End Function

Private Function LockTargetAt(ByRef targets As Variant, ByVal rowIndex As Long) As Variant
    Dim result(1 To 3) As Variant
    Dim columnIndex As Long

    For columnIndex = 1 To 3
        result(columnIndex) = targets(rowIndex, columnIndex)
    Next columnIndex
    LockTargetAt = result
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
