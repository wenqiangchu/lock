Attribute VB_Name = "MdlExcel"
Option Explicit
Option Private Module

Public Enum WorkbookRole
    WorkbookLocalCopy = 1
    WorkbookSharedRemote = 2
End Enum

Public Type SnapshotData
    SheetName As String
    HeaderFingerprint As String
    FirstDataRow As Long
    LastDataRow As Long
    Targets As Variant
    DeclaredColumnIndexes As Variant
    AnchorColumnIndexes As Variant
    AnchorValues As Variant
    Values As Variant
    Fonts As Variant
    Fills As Variant
End Type

Public Type SessionSchema
    SheetName As String
    HeaderFingerprint As String
    FirstDataRow As Long
    LastDataRow As Long
    Targets As Variant
    DeclaredColumnIndexes As Variant
    AnchorColumnIndexes As Variant
    AnchorValues As Variant
End Type

Public Const ERR_EXCEL_INVALID_INPUT As Long = 1601
Public Const ERR_EXCEL_HANDLE As Long = 1602
Public Const ERR_EXCEL_SHEET As Long = 1603
Public Const ERR_EXCEL_INVALID_CONTENTS As Long = 1604
Public Const ERR_EXCEL_ORPHAN_DATA As Long = 1605
Public Const ERR_EXCEL_STRUCTURE As Long = 1606
Public Const ERR_EXCEL_IO As Long = 1607
Public Const ERR_EXCEL_FILE_MISSING As Long = 1608
Public Const ERR_EXCEL_ACCESS_DENIED As Long = 1609
Public Const ERR_EXCEL_NETWORK_UNAVAILABLE As Long = 1610

Private Const XL_UP As Long = -4162
Private Const XL_CELL_TYPE_FORMULAS As Long = -4123
Private Const XL_MAX_ROWS As Long = 1048576
Private Const XL_PATTERN_NONE As Long = -4142
Private Const XL_PATTERN_SOLID As Long = 1

Private mHandles As Collection
Private mHeartbeatAt As Date
Private mHeartbeatProcedure As String
Private mHeartbeatScheduled As Boolean

Public Function ScheduleSessionHeartbeat(ByVal intervalSeconds As Long, ByRef errorInfo As AppErrorInfo) As Boolean
    If intervalSeconds < 1 Then
        errorInfo = MakeAppError(ERR_EXCEL_INVALID_INPUT, "MdlExcel.ScheduleSessionHeartbeat", "The heartbeat interval must be positive.")
        Exit Function
    End If

    On Error GoTo Failed
    CancelSessionHeartbeat
    mHeartbeatAt = DateAdd("s", intervalSeconds, Now)
    mHeartbeatProcedure = "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!MdlSession.SessionHeartbeat"
    Application.OnTime EarliestTime:=mHeartbeatAt, Procedure:=mHeartbeatProcedure
    mHeartbeatScheduled = True
    ScheduleSessionHeartbeat = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ScheduleSessionHeartbeat", "")
    Exit Function

Failed:
    mHeartbeatScheduled = False
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.ScheduleSessionHeartbeat", "The session heartbeat could not be scheduled.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Sub CancelSessionHeartbeat()
    If Not mHeartbeatScheduled Then Exit Sub
    On Error Resume Next
    Application.OnTime EarliestTime:=mHeartbeatAt, Procedure:=mHeartbeatProcedure, Schedule:=False
    On Error GoTo 0
    mHeartbeatScheduled = False
End Sub

Public Function OpenWorkbookHandle(ByVal path As String, ByVal role As WorkbookRole, ByRef handleId As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim normalizedPath As String
    Dim pathStatus As SharedPathStatus
    Dim systemError As Long
    Dim pathError As AppErrorInfo
    Dim workbook As Workbook
    Dim candidate As Workbook
    Dim entry As CWorkbookHandle
    Dim handleKey As String

    handleId = vbNullString
    On Error GoTo Failed
    Select Case role
        Case WorkbookLocalCopy
            normalizedPath = Trim$(path)
            If Not IsLocalAppDataPath(normalizedPath) Then GoTo InvalidPath
        Case WorkbookSharedRemote
            If Not ValidateWorkbookPath(path, normalizedPath, errorInfo) Then Exit Function
            If Not InspectWorkbookPath(normalizedPath, pathStatus, systemError, pathError) Then
                errorInfo = pathError
                Exit Function
            End If
            Select Case pathStatus
                Case SharedPathNotFound
                    errorInfo = MakeAppError(ERR_EXCEL_FILE_MISSING, "MdlExcel.OpenWorkbookHandle", "The configured workbook file or folder was not found.", normalizedPath)
                    Exit Function
                Case SharedPathAccessDenied
                    errorInfo = MakeAppError(ERR_EXCEL_ACCESS_DENIED, "MdlExcel.OpenWorkbookHandle", "Access to the configured workbook was denied.", normalizedPath & " (Windows error " & CStr(systemError) & ")")
                    Exit Function
                Case SharedPathNetworkUnavailable
                    errorInfo = MakeAppError(ERR_EXCEL_NETWORK_UNAVAILABLE, "MdlExcel.OpenWorkbookHandle", "The network location for the configured workbook is unavailable.", normalizedPath & " (Windows error " & CStr(systemError) & ")")
                    Exit Function
                Case SharedPathIsDirectory
                    errorInfo = MakeAppError(ERR_EXCEL_INVALID_INPUT, "MdlExcel.OpenWorkbookHandle", "The configured path identifies a folder, not a workbook file.", normalizedPath)
                    Exit Function
            End Select
        Case Else
            GoTo InvalidPath
    End Select

    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, normalizedPath, vbTextCompare) = 0 Then
            errorInfo = MakeAppError(ERR_EXCEL_HANDLE, "MdlExcel.OpenWorkbookHandle", "The workbook is already open in the host Excel instance.", normalizedPath)
            Exit Function
        End If
    Next candidate

    Set workbook = Application.Workbooks.Open(Filename:=normalizedPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    If workbook.ReadOnly Then
        workbook.Close SaveChanges:=False
        errorInfo = MakeAppError(ERR_EXCEL_HANDLE, "MdlExcel.OpenWorkbookHandle", "Excel opened the workbook read-only, so editing cannot start. It may be locked by another Excel session, read-only, or unavailable for writing under your account.", normalizedPath)
        Exit Function
    End If

    handleKey = NewGuidString()
    If Len(handleKey) = 0 Then GoTo Failed
    Set entry = New CWorkbookHandle
    entry.Initialize workbook, True, role
    EnsureHandles
    mHandles.Add entry, handleKey
    handleId = handleKey
    OpenWorkbookHandle = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.OpenWorkbookHandle", "")
    Exit Function

InvalidPath:
    errorInfo = MakeAppError(ERR_EXCEL_INVALID_INPUT, "MdlExcel.OpenWorkbookHandle", "A local copy below LOCALAPPDATA or a fully qualified workbook path is required.", path)
    Exit Function

Failed:
    If Not workbook Is Nothing Then
        On Error Resume Next
        workbook.Close SaveChanges:=False
        On Error GoTo 0
    End If
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.OpenWorkbookHandle", "The workbook could not be opened safely.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function ListWorksheetNames(ByVal handleId As String, ByRef sheetNames As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim entry As CWorkbookHandle
    Dim workbook As Workbook
    Dim worksheet As Worksheet
    Dim names() As Variant
    Dim index As Long

    sheetNames = Empty
    On Error GoTo Failed
    If Not ResolveHandle(handleId, entry, workbook, errorInfo) Then Exit Function
    If workbook.Worksheets.Count = 0 Then
        ListWorksheetNames = True
        errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ListWorksheetNames", "")
        Exit Function
    End If
    ReDim names(1 To workbook.Worksheets.Count)
    For Each worksheet In workbook.Worksheets
        index = index + 1
        names(index) = worksheet.Name
    Next worksheet
    sheetNames = names
    ListWorksheetNames = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ListWorksheetNames", "")
    Exit Function

Failed:
    sheetNames = Empty
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.ListWorksheetNames", "The worksheet list could not be read safely.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function ReadSheetSchema(ByVal handleId As String, ByVal sheetName As String, ByRef schema As SheetSchema, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim workbook As Workbook
    Dim worksheet As Worksheet
    Dim cells As Range
    Dim headerRange As Range
    Dim rawHeaders As Variant
    Dim headerValues() As Variant
    Dim lastColumn As Long
    Dim columnIndex As Long

    On Error GoTo Failed
    If Not ResolveWorksheet(handleId, sheetName, workbook, worksheet, errorInfo) Then Exit Function
    Set cells = worksheet.Cells
    Set headerRange = worksheet.Range(cells.Item(1, 1), cells.Item(1, worksheet.Columns.Count))
    rawHeaders = headerRange.Value2
    If IsArray(rawHeaders) Then
        For columnIndex = UBound(rawHeaders, 2) To LBound(rawHeaders, 2) Step -1
            If Not IsBlankHeader(rawHeaders(1, columnIndex)) Then
                lastColumn = columnIndex
                Exit For
            End If
        Next columnIndex
    ElseIf Not IsBlankHeader(rawHeaders) Then
        lastColumn = 1
    End If

    If lastColumn > 0 Then
        ReDim headerValues(1 To lastColumn)
        For columnIndex = 1 To lastColumn
            If IsArray(rawHeaders) Then
                headerValues(columnIndex) = rawHeaders(1, columnIndex)
            Else
                headerValues(columnIndex) = rawHeaders
            End If
        Next columnIndex
        If Not ParseHeader(headerValues, schema, errorInfo) Then Exit Function
    Else
        If Not ParseHeader(Empty, schema, errorInfo) Then Exit Function
    End If
    ReadSheetSchema = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ReadSheetSchema", "")
    Exit Function

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_SHEET, "MdlExcel.ReadSheetSchema", "The sheet header could not be read.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function ReadDeclaredSnapshot(ByVal handleId As String, ByRef targets As Variant, ByRef snapshot As SnapshotData, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim workbook As Workbook
    Dim worksheet As Worksheet
    Dim schema As SheetSchema
    Dim declaredIndexes As Variant
    Dim invalidCells As Variant
    Dim rawValues As Variant
    Dim values() As Variant
    Dim fonts() As Variant
    Dim fills() As Variant
    Dim anchorValues() As Variant
    Dim firstRow As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim anchorIndex As Long
    Dim anchorCount As Long
    Dim hasAnchor As Boolean
    Dim hasDeclaredValue As Boolean
    Dim textFontData As String
    Dim fillData As String

    ResetSnapshot snapshot
    If Not ResolveSnapshotInputs(handleId, targets, workbook, worksheet, schema, declaredIndexes, errorInfo) Then Exit Function
    If Not ValidateEditableContents(handleId, targets, invalidCells, errorInfo) Then Exit Function
    lastRow = FindLastDataRow(worksheet, schema.AnchorColumnIndexes, declaredIndexes)
    firstRow = 2

    snapshot.SheetName = CStr(targets(LBound(targets, 1), LBound(targets, 2)))
    snapshot.HeaderFingerprint = schema.Fingerprint
    snapshot.FirstDataRow = firstRow
    snapshot.LastDataRow = lastRow
    snapshot.Targets = targets
    snapshot.DeclaredColumnIndexes = declaredIndexes
    snapshot.AnchorColumnIndexes = schema.AnchorColumnIndexes

    If lastRow < firstRow Then
        snapshot.Values = Empty
        snapshot.Fonts = Empty
        snapshot.Fills = Empty
        snapshot.AnchorValues = Empty
        ReadDeclaredSnapshot = True
        errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ReadDeclaredSnapshot", "")
        Exit Function
    End If

    rawValues = worksheet.Range(worksheet.Cells(firstRow, 1), worksheet.Cells(lastRow, schema.ColumnCount)).Value2
    ReDim values(firstRow To lastRow, 1 To schema.ColumnCount)
    For rowIndex = firstRow To lastRow
        For columnIndex = 1 To schema.ColumnCount
            values(rowIndex, columnIndex) = MatrixOrScalarValue(rawValues, rowIndex - firstRow + 1, columnIndex)
        Next columnIndex
    Next rowIndex

    anchorCount = UBound(schema.AnchorColumnIndexes) - LBound(schema.AnchorColumnIndexes) + 1
    ReDim anchorValues(firstRow To lastRow, 1 To anchorCount)
    For rowIndex = firstRow To lastRow
        hasAnchor = False
        For anchorIndex = 1 To anchorCount
            columnIndex = CLng(schema.AnchorColumnIndexes(LBound(schema.AnchorColumnIndexes) + anchorIndex - 1))
            anchorValues(rowIndex, anchorIndex) = values(rowIndex, columnIndex)
            If IsMeaningfulValue(values(rowIndex, columnIndex)) Then hasAnchor = True
        Next anchorIndex

        hasDeclaredValue = False
        For anchorIndex = LBound(declaredIndexes) To UBound(declaredIndexes)
            columnIndex = CLng(declaredIndexes(anchorIndex))
            If IsMeaningfulValue(values(rowIndex, columnIndex)) Then hasDeclaredValue = True
        Next anchorIndex
        If Not hasAnchor And hasDeclaredValue Then
            errorInfo = MakeAppError(ERR_EXCEL_ORPHAN_DATA, "MdlExcel.ReadDeclaredSnapshot", "A declared AMG value cannot be assigned to a product anchor row.", CStr(snapshot.SheetName) & " row " & CStr(rowIndex))
            Exit Function
        End If
    Next rowIndex

    ReDim fonts(firstRow To lastRow, 1 To schema.ColumnCount)
    ReDim fills(firstRow To lastRow, 1 To schema.ColumnCount)
    For rowIndex = firstRow To lastRow
        For anchorIndex = LBound(declaredIndexes) To UBound(declaredIndexes)
            columnIndex = CLng(declaredIndexes(anchorIndex))
            If Not ReadTextFontData(worksheet.Cells(rowIndex, columnIndex), values(rowIndex, columnIndex), textFontData, errorInfo) Then Exit Function
            If Not ReadFillData(worksheet.Cells(rowIndex, columnIndex), fillData, errorInfo) Then Exit Function
            fonts(rowIndex, columnIndex) = textFontData
            fills(rowIndex, columnIndex) = fillData
        Next anchorIndex
    Next rowIndex

    snapshot.Values = values
    snapshot.Fonts = fonts
    snapshot.Fills = fills
    snapshot.AnchorValues = anchorValues
    ReadDeclaredSnapshot = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ReadDeclaredSnapshot", "")
End Function

Public Function ValidateEditableContents(ByVal handleId As String, ByRef targets As Variant, ByRef invalidCells As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim workbook As Workbook
    Dim worksheet As Worksheet
    Dim schema As SheetSchema
    Dim declaredIndexes As Variant

    invalidCells = Empty
    If Not ResolveSnapshotInputs(handleId, targets, workbook, worksheet, schema, declaredIndexes, errorInfo) Then Exit Function
    ValidateEditableContents = ValidateSheetColumns(worksheet, schema, declaredIndexes, CStr(targets(LBound(targets, 1), LBound(targets, 2))), invalidCells, errorInfo)
End Function

Public Sub CopySnapshotSchema(ByRef snapshot As SnapshotData, ByRef expected As SessionSchema)
    expected.SheetName = snapshot.SheetName
    expected.HeaderFingerprint = snapshot.HeaderFingerprint
    expected.FirstDataRow = snapshot.FirstDataRow
    expected.LastDataRow = snapshot.LastDataRow
    expected.Targets = snapshot.Targets
    expected.DeclaredColumnIndexes = snapshot.DeclaredColumnIndexes
    expected.AnchorColumnIndexes = snapshot.AnchorColumnIndexes
    expected.AnchorValues = snapshot.AnchorValues
End Sub

Public Function ValidateStructure(ByVal handleId As String, ByRef expected As SessionSchema, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim current As SnapshotData
    Dim currentSchema As SheetSchema

    valid = False
    If Not ReadSheetSchema(handleId, expected.SheetName, currentSchema, errorInfo) Then Exit Function
    If StrComp(currentSchema.Fingerprint, expected.HeaderFingerprint, vbBinaryCompare) <> 0 Then GoTo StructureChanged
    If Not ReadDeclaredSnapshot(handleId, expected.Targets, current, errorInfo) Then Exit Function
    If current.FirstDataRow <> expected.FirstDataRow Or current.LastDataRow <> expected.LastDataRow Then GoTo StructureChanged
    If Not IndexArraysEqual(current.AnchorColumnIndexes, expected.AnchorColumnIndexes) Then GoTo StructureChanged
    If Not IndexArraysEqual(current.DeclaredColumnIndexes, expected.DeclaredColumnIndexes) Then GoTo StructureChanged
    If Not AnchorMatricesEqual(current.AnchorValues, expected.AnchorValues) Then GoTo StructureChanged
    valid = True
    ValidateStructure = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ValidateStructure", "")
    Exit Function

StructureChanged:
    ValidateStructure = True
    errorInfo = MakeAppError(ERR_EXCEL_STRUCTURE, "MdlExcel.ValidateStructure", "The sheet structure or product anchor mapping changed during the session.", expected.SheetName)
End Function

Public Function WriteDeclaredColumns(ByVal handleId As String, ByRef targets As Variant, ByRef mergedValues As Variant, ByRef mergedFonts As Variant, ByRef mergedFills As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim workbook As Workbook
    Dim worksheet As Worksheet
    Dim schema As SheetSchema
    Dim declaredIndexes As Variant
    Dim invalidCells As Variant
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim outputRow As Long
    Dim equal As Boolean
    Dim comparisonError As AppErrorInfo
    Dim expectedValue As Variant
    Dim actualValue As Variant
    Dim outputValues() As Variant
    Dim rangeToWrite As Range
    Dim readBack As Variant
    Dim fontData As String
    Dim fillData As String
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim appStateChanged As Boolean

    On Error GoTo Failed
    If Not ResolveSnapshotInputs(handleId, targets, workbook, worksheet, schema, declaredIndexes, errorInfo) Then Exit Function
    If IsEmpty(mergedValues) Then
        If Not IsEmpty(mergedFonts) Or Not IsEmpty(mergedFills) Then GoTo InvalidInput
        WriteDeclaredColumns = True
        errorInfo = MakeAppError(ERR_NONE, "MdlExcel.WriteDeclaredColumns", "")
        Exit Function
    End If
    If Not GetMatrixBounds(mergedValues, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If columnLow > LBound(schema.HeaderNames) Or columnHigh < UBound(schema.HeaderNames) Then GoTo InvalidInput
    If rowLow < 2 Or rowHigh > XL_MAX_ROWS Then GoTo InvalidInput
    If Not MatrixBoundsMatch(mergedFonts, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not MatrixBoundsMatch(mergedFills, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If Not ValidateSheetColumns(worksheet, schema, declaredIndexes, CStr(targets(LBound(targets, 1), LBound(targets, 2))), invalidCells, errorInfo) Then Exit Function

    For rowIndex = rowLow To rowHigh
        For columnIndex = LBound(declaredIndexes) To UBound(declaredIndexes)
            expectedValue = mergedValues(rowIndex, CLng(declaredIndexes(columnIndex)))
            fontData = CStr(mergedFonts(rowIndex, CLng(declaredIndexes(columnIndex))))
            fillData = CStr(mergedFills(rowIndex, CLng(declaredIndexes(columnIndex))))
            If Not IsSupportedCellValue(expectedValue) Then GoTo InvalidValue
            If Not TextFontMatchesValue(expectedValue, fontData, errorInfo) Then GoTo InvalidFormat
            If Not CompareFormattingData(fillData, fillData, equal, comparisonError) Then GoTo InvalidFormat
        Next columnIndex
    Next rowIndex

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    appStateChanged = True
    rowCount = rowHigh - rowLow + 1

    For columnIndex = LBound(declaredIndexes) To UBound(declaredIndexes)
        ReDim outputValues(1 To rowCount, 1 To 1)
        For rowIndex = rowLow To rowHigh
            expectedValue = mergedValues(rowIndex, CLng(declaredIndexes(columnIndex)))
            If Not IsSupportedCellValue(expectedValue) Then GoTo InvalidValue
            outputRow = rowIndex - rowLow + 1
            outputValues(outputRow, 1) = expectedValue
        Next rowIndex

        Set rangeToWrite = worksheet.Range(worksheet.Cells(rowLow, CLng(declaredIndexes(columnIndex))), worksheet.Cells(rowHigh, CLng(declaredIndexes(columnIndex))))
        rangeToWrite.Value2 = outputValues
        readBack = rangeToWrite.Value2
        For rowIndex = rowLow To rowHigh
            outputRow = rowIndex - rowLow + 1
            expectedValue = outputValues(outputRow, 1)
            actualValue = MatrixOrScalarValue(readBack, outputRow, 1)
            If Not CompareCellValues(expectedValue, actualValue, equal, comparisonError) Then
                errorInfo = comparisonError
                GoTo Cleanup
            End If
            If Not equal Then
                errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.WriteDeclaredColumns", "The written value did not match the requested merge result.", CStr(targets(LBound(targets, 1), LBound(targets, 2))) & " row " & CStr(rowIndex) & " column " & CStr(declaredIndexes(columnIndex)))
                GoTo Cleanup
            End If
        Next rowIndex
    Next columnIndex

    For rowIndex = rowLow To rowHigh
        For columnIndex = LBound(declaredIndexes) To UBound(declaredIndexes)
            outputRow = CLng(declaredIndexes(columnIndex))
            expectedValue = mergedValues(rowIndex, outputRow)
            fontData = CStr(mergedFonts(rowIndex, outputRow))
            fillData = CStr(mergedFills(rowIndex, outputRow))
            If Not WriteTextFontData(worksheet.Cells(rowIndex, outputRow), expectedValue, fontData, errorInfo) Then GoTo Cleanup
            If Not WriteFillData(worksheet.Cells(rowIndex, outputRow), fillData, errorInfo) Then GoTo Cleanup
        Next columnIndex
    Next rowIndex

    WriteDeclaredColumns = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.WriteDeclaredColumns", "")
    GoTo Cleanup

InvalidInput:
    errorInfo = MakeAppError(ERR_EXCEL_INVALID_INPUT, "MdlExcel.WriteDeclaredColumns", "Merged values must be a physical-row matrix covering the declared columns.")
    GoTo Cleanup

InvalidValue:
    errorInfo = MakeAppError(ERR_EXCEL_INVALID_CONTENTS, "MdlExcel.WriteDeclaredColumns", "Only string or Empty values may be written to AMG cells.", CStr(rowIndex) & "," & CStr(declaredIndexes(columnIndex)))
    GoTo Cleanup

InvalidFormat:
    If errorInfo.Code = ERR_NONE Then errorInfo = MakeAppError(ERR_EXCEL_INVALID_CONTENTS, "MdlExcel.WriteDeclaredColumns", "A merged font or fill snapshot is invalid or does not match its cell value.", CStr(rowIndex) & "," & CStr(declaredIndexes(columnIndex)))
    GoTo Cleanup

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.WriteDeclaredColumns", "Declared values could not be written safely.", CStr(Err.Number) & ": " & Err.Description)

Cleanup:
    If appStateChanged Then
        On Error Resume Next
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreenUpdating
        On Error GoTo 0
    End If
End Function

Public Function SaveWorkbook(ByVal handleId As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim workbook As Workbook
    Dim entry As CWorkbookHandle

    On Error GoTo Failed
    If Not ResolveHandle(handleId, entry, workbook, errorInfo) Then Exit Function
    If workbook.ReadOnly Then
        errorInfo = MakeAppError(ERR_EXCEL_HANDLE, "MdlExcel.SaveWorkbook", "The workbook is read-only and cannot be saved.", workbook.FullName)
        Exit Function
    End If
    workbook.Save
    SaveWorkbook = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.SaveWorkbook", "")
    Exit Function

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.SaveWorkbook", "Workbook.Save failed; the save result may be uncertain.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Sub CloseWorkbookHandle(ByVal handleId As String, ByVal saveChanges As Boolean)
    Dim entry As CWorkbookHandle
    Dim workbook As Workbook
    Dim errorInfo As AppErrorInfo

    If Not ResolveHandle(handleId, entry, workbook, errorInfo) Then Exit Sub
    On Error GoTo Failed
    If entry.IsOwned Then workbook.Close SaveChanges:=saveChanges
    mHandles.Remove handleId
    entry.Detach
    Exit Sub

Failed:
    LogWrite LogError, "MdlExcel.CloseWorkbookHandle", CStr(Err.Number) & ": " & Err.Description
End Sub

Public Function RegisterCurrentTestWorkbook(ByRef handleId As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim entry As CWorkbookHandle

    handleId = NewGuidString()
    If Len(handleId) = 0 Then
        errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.RegisterCurrentTestWorkbook", "A test handle could not be generated.")
        Exit Function
    End If
    Set entry = New CWorkbookHandle
    entry.Initialize ThisWorkbook, False, WorkbookLocalCopy
    EnsureHandles
    mHandles.Add entry, handleId
    RegisterCurrentTestWorkbook = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.RegisterCurrentTestWorkbook", "")
End Function

Private Function ResolveSnapshotInputs(ByVal handleId As String, ByRef targets As Variant, ByRef workbook As Workbook, ByRef worksheet As Worksheet, ByRef schema As SheetSchema, ByRef declaredIndexes As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim targetSheet As String
    Dim headerName As String
    Dim normalizedHeader As String
    Dim foundIndex As Long
    Dim indexes() As Variant
    Dim entry As CWorkbookHandle

    If Not GetMatrixBounds(targets, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidTargets
    If columnHigh - columnLow + 1 <> 2 Then GoTo InvalidTargets
    For rowIndex = rowLow To rowHigh
        If VarType(targets(rowIndex, columnLow)) <> vbString Or VarType(targets(rowIndex, columnLow + 1)) <> vbString Then GoTo InvalidTargets
        If rowIndex = rowLow Then
            targetSheet = CStr(targets(rowIndex, columnLow))
        ElseIf StrComp(targetSheet, CStr(targets(rowIndex, columnLow)), vbTextCompare) <> 0 Then
            GoTo InvalidTargets
        End If
    Next rowIndex
    If Not ResolveHandle(handleId, entry, workbook, errorInfo) Then Exit Function
    If Not GetWorksheetFromWorkbook(workbook, targetSheet, worksheet) Then
        errorInfo = MakeAppError(ERR_EXCEL_SHEET, "MdlExcel.ResolveSnapshotInputs", "The requested sheet does not exist.", targetSheet)
        Exit Function
    End If
    If Not ReadSheetSchema(handleId, targetSheet, schema, errorInfo) Then Exit Function

    ReDim indexes(1 To rowHigh - rowLow + 1)
    For rowIndex = rowLow To rowHigh
        headerName = CStr(targets(rowIndex, columnLow + 1))
        normalizedHeader = UCase$(Trim$(headerName))
        foundIndex = 0
        For columnIndex = 1 To schema.ColumnCount
            If StrComp(CStr(schema.NormalizedNames(columnIndex)), normalizedHeader, vbBinaryCompare) = 0 Then
                If schema.ColumnKinds(columnIndex) <> SchemaColumnAmg Then GoTo InvalidTargets
                foundIndex = columnIndex
                Exit For
            End If
        Next columnIndex
        If foundIndex = 0 Then GoTo InvalidTargets
        For columnIndex = 1 To rowIndex - rowLow
            If CLng(indexes(columnIndex)) = foundIndex Then GoTo InvalidTargets
        Next columnIndex
        indexes(rowIndex - rowLow + 1) = foundIndex
    Next rowIndex
    declaredIndexes = indexes
    ResolveSnapshotInputs = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ResolveSnapshotInputs", "")
    Exit Function

InvalidTargets:
    errorInfo = MakeAppError(ERR_EXCEL_INVALID_INPUT, "MdlExcel.ResolveSnapshotInputs", "Targets must contain one sheet and unique AMG headers in a two-column matrix.")
End Function

Private Function ValidateSheetColumns(ByRef worksheet As Worksheet, ByRef schema As SheetSchema, ByRef declaredIndexes As Variant, ByVal sheetName As String, ByRef invalidCells As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim lastRow As Long
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim rangeToCheck As Range
    Dim formulaCells As Range
    Dim cell As Range
    Dim values As Variant
    Dim value As Variant
    Dim address As String

    On Error GoTo Failed
    invalidCells = Empty
    lastRow = FindLastDataRow(worksheet, schema.AnchorColumnIndexes, declaredIndexes)
    If lastRow < 2 Then
        ValidateSheetColumns = True
        errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ValidateEditableContents", "")
        Exit Function
    End If
    rowCount = lastRow - 1

    For columnIndex = LBound(declaredIndexes) To UBound(declaredIndexes)
        Set rangeToCheck = worksheet.Range(worksheet.Cells(2, CLng(declaredIndexes(columnIndex))), worksheet.Cells(lastRow, CLng(declaredIndexes(columnIndex))))
        Set formulaCells = Nothing
        On Error Resume Next
        Set formulaCells = rangeToCheck.SpecialCells(XL_CELL_TYPE_FORMULAS)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo Failed
        If Not formulaCells Is Nothing Then
            For Each cell In formulaCells.Cells
                address = cell.Address(False, False)
                invalidCells = Array(sheetName, address, "formula")
                errorInfo = MakeAppError(ERR_EXCEL_INVALID_CONTENTS, "MdlExcel.ValidateEditableContents", "AMG columns cannot contain formulas.", sheetName & "!" & address)
                Exit Function
            Next cell
        End If

        values = rangeToCheck.Value2
        For rowIndex = 1 To rowCount
            value = MatrixOrScalarValue(values, rowIndex, 1)
            If Not IsSupportedCellValue(value) Then
                address = worksheet.Cells(rowIndex + 1, CLng(declaredIndexes(columnIndex))).Address(False, False)
                invalidCells = Array(sheetName, address, "unsupported value type")
                errorInfo = MakeAppError(ERR_EXCEL_INVALID_CONTENTS, "MdlExcel.ValidateEditableContents", "AMG cells must contain only strings or blank cells.", sheetName & "!" & address)
                Exit Function
            End If
        Next rowIndex
    Next columnIndex

    ValidateSheetColumns = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ValidateEditableContents", "")
    Exit Function

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.ValidateEditableContents", "Declared cell contents could not be validated.", CStr(Err.Number) & ": " & Err.Description)
End Function

Private Function FindLastDataRow(ByRef worksheet As Worksheet, ByRef anchorIndexes As Variant, ByRef declaredIndexes As Variant) As Long
    Dim columnIndex As Variant
    Dim lastRow As Long
    Dim candidate As Long

    lastRow = 1
    For Each columnIndex In anchorIndexes
        candidate = CLng(worksheet.Cells(XL_MAX_ROWS, CLng(columnIndex)).End(XL_UP).Row)
        If candidate > lastRow Then lastRow = candidate
    Next columnIndex
    For Each columnIndex In declaredIndexes
        candidate = CLng(worksheet.Cells(XL_MAX_ROWS, CLng(columnIndex)).End(XL_UP).Row)
        If candidate > lastRow Then lastRow = candidate
    Next columnIndex
    FindLastDataRow = lastRow
End Function


Private Function GetWorksheetFromWorkbook(ByRef workbook As Workbook, ByVal sheetName As String, ByRef worksheet As Worksheet) As Boolean
    On Error GoTo Missing
    Set worksheet = workbook.Worksheets.Item(sheetName)
    GetWorksheetFromWorkbook = True
    Exit Function
Missing:
    Set worksheet = Nothing
End Function

Private Function ResolveWorksheet(ByVal handleId As String, ByVal sheetName As String, ByRef workbook As Workbook, ByRef worksheet As Worksheet, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim entry As CWorkbookHandle
    If Not ResolveHandle(handleId, entry, workbook, errorInfo) Then Exit Function
    If Not GetWorksheetFromWorkbook(workbook, sheetName, worksheet) Then
        errorInfo = MakeAppError(ERR_EXCEL_SHEET, "MdlExcel.ResolveWorksheet", "The requested sheet does not exist.", sheetName)
        Exit Function
    End If
    ResolveWorksheet = True
End Function

Private Function ResolveHandle(ByVal handleId As String, ByRef entry As CWorkbookHandle, ByRef workbook As Workbook, ByRef errorInfo As AppErrorInfo) As Boolean
    On Error GoTo Missing
    If mHandles Is Nothing Then GoTo Missing
    Set entry = mHandles.Item(handleId)
    Set workbook = entry.WorkbookObject
    If workbook Is Nothing Then GoTo Missing
    ResolveHandle = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.ResolveHandle", "")
    Exit Function
Missing:
    errorInfo = MakeAppError(ERR_EXCEL_HANDLE, "MdlExcel.ResolveHandle", "The workbook handle is invalid or closed.", handleId)
End Function

Private Sub EnsureHandles()
    If mHandles Is Nothing Then Set mHandles = New Collection
End Sub

Private Function ReadTextFontData(ByVal cell As Range, ByVal cellValue As Variant, ByRef data As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim textValue As String
    Dim currentFont As String
    Dim characterFont As String
    Dim emptyCellFont As String
    Dim emptyRuns As Variant
    Dim runBuffer() As Variant
    Dim runs() As Variant
    Dim runCount As Long
    Dim runStart As Long
    Dim runLength As Long
    Dim characterIndex As Long
    Dim index As Long
    Dim fontObject As Object

    On Error GoTo Failed
    If VarType(cellValue) = vbString Then textValue = CStr(cellValue)
    If Len(textValue) = 0 Then
        Set fontObject = cell.Font
        If Not CaptureFontSignature(fontObject, emptyCellFont, errorInfo) Then Exit Function
        emptyRuns = Empty
        data = BuildTextFontData(cellValue, emptyRuns, emptyCellFont, errorInfo)
        ReadTextFontData = (Len(data) > 0)
        Exit Function
    End If

    ReDim runBuffer(1 To Len(textValue), 1 To 3)
    ' Character-level reads are required to preserve localized font runs.
    For characterIndex = 1 To Len(textValue)
        Set fontObject = cell.Characters(characterIndex, 1).Font
        If Not CaptureFontSignature(fontObject, characterFont, errorInfo) Then Exit Function
        If characterIndex = 1 Then
            emptyCellFont = characterFont
            currentFont = characterFont
            runStart = 1
            runLength = 1
        ElseIf StrComp(characterFont, currentFont, vbBinaryCompare) = 0 Then
            runLength = runLength + 1
        Else
            runCount = runCount + 1
            runBuffer(runCount, 1) = runStart
            runBuffer(runCount, 2) = runLength
            runBuffer(runCount, 3) = currentFont
            currentFont = characterFont
            runStart = characterIndex
            runLength = 1
        End If
    Next characterIndex
    runCount = runCount + 1
    runBuffer(runCount, 1) = runStart
    runBuffer(runCount, 2) = runLength
    runBuffer(runCount, 3) = currentFont
    ReDim runs(1 To runCount, 1 To 3)
    For index = 1 To runCount
        runs(index, 1) = runBuffer(index, 1)
        runs(index, 2) = runBuffer(index, 2)
        runs(index, 3) = runBuffer(index, 3)
    Next index
    data = BuildTextFontData(cellValue, runs, emptyCellFont, errorInfo)
    ReadTextFontData = (Len(data) > 0)
    Exit Function

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.ReadTextFontData", "A declared cell font could not be read safely.", CStr(Err.Number) & ": " & Err.Description)
End Function

Private Function CaptureFontSignature(ByVal fontObject As Object, ByRef signature As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim properties As Variant
    Dim propertyNames As Variant

    propertyNames = Array("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Superscript", "Subscript", "OutlineFont", "Shadow", "ColorIndex", "Color", "ThemeColor", "TintAndShade")
    If Not CaptureObjectProperties(fontObject, propertyNames, ",13,14,", properties, errorInfo, "MdlExcel.CaptureFontSignature") Then Exit Function
    If IsNumeric(properties(13)) Then
        If CDbl(properties(13)) = 0 Then
            properties(13) = CVErr(2042)
            properties(14) = CVErr(2042)
        End If
    End If
    signature = BuildFontSignature(properties, errorInfo)
    CaptureFontSignature = (Len(signature) > 0)
End Function

Private Function ReadFillData(ByVal cell As Range, ByRef data As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim fillObject As Object
    Dim properties As Variant
    Dim propertyNames As Variant

    On Error GoTo Failed
    Set fillObject = cell.Interior
    propertyNames = Array("Pattern", "ColorIndex", "Color", "ThemeColor", "TintAndShade", "PatternColorIndex", "PatternColor", "PatternThemeColor", "PatternTintAndShade")
    If Not CaptureObjectProperties(fillObject, propertyNames, ",4,5,8,9,", properties, errorInfo, "MdlExcel.ReadFillData") Then Exit Function
    If IsNumeric(properties(4)) Then
        If CDbl(properties(4)) = 0 Then
            properties(4) = CVErr(2042)
            properties(5) = CVErr(2042)
        End If
    End If
    If IsNumeric(properties(8)) Then
        If CDbl(properties(8)) = 0 Then
            properties(8) = CVErr(2042)
            properties(9) = CVErr(2042)
        End If
    End If
    If IsNumeric(properties(1)) Then
        If CDbl(properties(1)) = XL_PATTERN_SOLID Or CDbl(properties(1)) = XL_PATTERN_NONE Then
            properties(6) = CVErr(2042)
            properties(7) = CVErr(2042)
            properties(8) = CVErr(2042)
            properties(9) = CVErr(2042)
        End If
    End If
    data = BuildFillData(properties, errorInfo)
    ReadFillData = (Len(data) > 0)
    Exit Function

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.ReadFillData", "A declared cell fill could not be read safely.", CStr(Err.Number) & ": " & Err.Description)
End Function

Private Function CaptureObjectProperties(ByVal source As Object, ByVal propertyNames As Variant, ByVal optionalIndexes As String, ByRef properties As Variant, ByRef errorInfo As AppErrorInfo, ByVal sourceName As String) As Boolean
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long
    Dim value As Variant
    Dim result() As Variant

    On Error GoTo Failed
    lowerBound = LBound(propertyNames)
    upperBound = UBound(propertyNames)
    ReDim result(1 To upperBound - lowerBound + 1)
    For index = lowerBound To upperBound
        If TryReadObjectProperty(source, CStr(propertyNames(index)), value) Then
            result(index - lowerBound + 1) = value
        ElseIf InStr(1, optionalIndexes, "," & CStr(index - lowerBound + 1) & ",", vbBinaryCompare) > 0 Then
            result(index - lowerBound + 1) = CVErr(2042)
        Else
            errorInfo = MakeAppError(ERR_EXCEL_IO, sourceName, "A required direct formatting property could not be read.", CStr(propertyNames(index)))
            Exit Function
        End If
    Next index
    properties = result
    CaptureObjectProperties = True
    errorInfo = MakeAppError(ERR_NONE, sourceName, "")
    Exit Function

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_IO, sourceName, "A direct formatting property could not be read.", CStr(Err.Number) & ": " & Err.Description)
End Function

Private Function TryReadObjectProperty(ByVal source As Object, ByVal propertyName As String, ByRef value As Variant) As Boolean
    On Error GoTo Failed
    value = CallByName(source, propertyName, VbGet)
    If IsNull(value) Or IsError(value) Or IsObject(value) Or IsArray(value) Then Exit Function
    TryReadObjectProperty = True
    Exit Function

Failed:
    Err.Clear
End Function

Private Function WriteTextFontData(ByVal cell As Range, ByVal cellValue As Variant, ByVal data As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim valueKind As String
    Dim textValue As String
    Dim emptyCellFont As String
    Dim runs As Variant
    Dim actualData As String
    Dim runIndex As Long
    Dim fontObject As Object
    Dim equal As Boolean
    Dim compareError As AppErrorInfo

    If Not TextFontMatchesValue(cellValue, data, errorInfo) Then Exit Function
    If Not ParseTextFontData(data, valueKind, textValue, runs, emptyCellFont, errorInfo) Then Exit Function
    On Error GoTo Failed
    If IsEmpty(runs) Then
        Set fontObject = cell.Font
        If Not ApplyFontSignature(fontObject, emptyCellFont, errorInfo) Then Exit Function
    Else
        For runIndex = LBound(runs, 1) To UBound(runs, 1)
            Set fontObject = cell.Characters(CLng(runs(runIndex, 1)), CLng(runs(runIndex, 2))).Font
            If Not ApplyFontSignature(fontObject, CStr(runs(runIndex, 3)), errorInfo) Then Exit Function
        Next runIndex
    End If
    If Not ReadTextFontData(cell, cellValue, actualData, errorInfo) Then Exit Function
    If Not CompareFormattingData(data, actualData, equal, compareError) Then
        errorInfo = compareError
        Exit Function
    End If
    If Not equal Then
        errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.WriteTextFontData", "The written font runs did not match the requested snapshot.")
        Exit Function
    End If
    WriteTextFontData = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.WriteTextFontData", "")
    Exit Function

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.WriteTextFontData", "A declared cell font could not be written safely.", CStr(Err.Number) & ": " & Err.Description)
End Function

Private Function ApplyFontSignature(ByVal fontObject As Object, ByVal signature As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim properties As Variant
    Dim propertyNames As Variant
    Dim index As Long
    Dim themeApplicable As Boolean

    If Not ParseFontSignature(signature, properties, errorInfo) Then Exit Function
    propertyNames = Array("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Superscript", "Subscript", "OutlineFont", "Shadow")
    For index = 0 To 9
        If Not SetNormalizedProperty(fontObject, CStr(propertyNames(index)), CStr(properties(index + 1)), errorInfo) Then Exit Function
    Next index
    themeApplicable = (CStr(properties(13)) <> "A")
    If themeApplicable Then
        If Not SetNormalizedProperty(fontObject, "ThemeColor", CStr(properties(13)), errorInfo) Then Exit Function
        If Not SetNormalizedProperty(fontObject, "TintAndShade", CStr(properties(14)), errorInfo) Then Exit Function
    Else
        If Not SetNormalizedProperty(fontObject, "ColorIndex", CStr(properties(11)), errorInfo) Then Exit Function
        If Not SetNormalizedProperty(fontObject, "Color", CStr(properties(12)), errorInfo) Then Exit Function
    End If
    ApplyFontSignature = True
End Function

Private Function WriteFillData(ByVal cell As Range, ByVal data As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim fillObject As Object
    Dim properties As Variant
    Dim actualData As String
    Dim themeApplicable As Boolean
    Dim patternThemeApplicable As Boolean
    Dim equal As Boolean
    Dim compareError As AppErrorInfo

    If Not ParseFillData(data, properties, errorInfo) Then Exit Function
    On Error GoTo Failed
    Set fillObject = cell.Interior
    If Not SetNormalizedProperty(fillObject, "Pattern", CStr(properties(1)), errorInfo) Then Exit Function
    themeApplicable = (CStr(properties(4)) <> "A")
    If themeApplicable Then
        If Not SetNormalizedProperty(fillObject, "ThemeColor", CStr(properties(4)), errorInfo) Then Exit Function
        If Not SetNormalizedProperty(fillObject, "TintAndShade", CStr(properties(5)), errorInfo) Then Exit Function
    Else
        If Not SetNormalizedProperty(fillObject, "ColorIndex", CStr(properties(2)), errorInfo) Then Exit Function
        If Not SetNormalizedProperty(fillObject, "Color", CStr(properties(3)), errorInfo) Then Exit Function
    End If
    patternThemeApplicable = (CStr(properties(8)) <> "A")
    If patternThemeApplicable Then
        If Not SetNormalizedProperty(fillObject, "PatternThemeColor", CStr(properties(8)), errorInfo) Then Exit Function
        If Not SetNormalizedProperty(fillObject, "PatternTintAndShade", CStr(properties(9)), errorInfo) Then Exit Function
    Else
        If Not SetNormalizedProperty(fillObject, "PatternColorIndex", CStr(properties(6)), errorInfo) Then Exit Function
        If Not SetNormalizedProperty(fillObject, "PatternColor", CStr(properties(7)), errorInfo) Then Exit Function
    End If
    If Not ReadFillData(cell, actualData, errorInfo) Then Exit Function
    If Not CompareFormattingData(data, actualData, equal, compareError) Then
        errorInfo = compareError
        Exit Function
    End If
    If Not equal Then
        errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.WriteFillData", "The written fill did not match the requested snapshot.", "Expected: " & FormatFormattingPreview(data) & vbCrLf & "Actual: " & FormatFormattingPreview(actualData))
        Exit Function
    End If
    WriteFillData = True
    errorInfo = MakeAppError(ERR_NONE, "MdlExcel.WriteFillData", "")
    Exit Function

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.WriteFillData", "A declared cell fill could not be written safely.", CStr(Err.Number) & ": " & Err.Description)
End Function

Private Function SetNormalizedProperty(ByVal target As Object, ByVal propertyName As String, ByVal encodedValue As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim value As Variant
    Dim applicable As Boolean
    Dim decodeError As AppErrorInfo

    If Not DecodeFormatScalar(encodedValue, value, applicable, decodeError) Then
        errorInfo = decodeError
        Exit Function
    End If
    If Not applicable Then
        SetNormalizedProperty = True
        Exit Function
    End If
    On Error GoTo Failed
    CallByName target, propertyName, VbLet, value
    SetNormalizedProperty = True
    Exit Function

Failed:
    errorInfo = MakeAppError(ERR_EXCEL_IO, "MdlExcel.SetNormalizedProperty", "A direct formatting property could not be written.", propertyName & "=" & encodedValue & "; " & CStr(Err.Number) & ": " & Err.Description)
End Function

Private Function TextFontMatchesValue(ByVal cellValue As Variant, ByVal fontData As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim valueKind As String
    Dim textValue As String
    Dim emptyCellFont As String
    Dim runs As Variant

    If Not ParseTextFontData(fontData, valueKind, textValue, runs, emptyCellFont, errorInfo) Then Exit Function
    If VarType(cellValue) = vbEmpty Then
        TextFontMatchesValue = (valueKind = "E")
    ElseIf VarType(cellValue) = vbString Then
        TextFontMatchesValue = (valueKind = "S" And StrComp(textValue, CStr(cellValue), vbBinaryCompare) = 0)
    End If
    If Not TextFontMatchesValue Then errorInfo = MakeAppError(ERR_EXCEL_INVALID_CONTENTS, "MdlExcel.TextFontMatchesValue", "The font snapshot text does not match the merged cell value.")
End Function

Private Function MatrixBoundsMatch(ByRef values As Variant, ByVal rowLow As Long, ByVal rowHigh As Long, ByVal columnLow As Long, ByVal columnHigh As Long) As Boolean
    Dim actualRowLow As Long
    Dim actualRowHigh As Long
    Dim actualColumnLow As Long
    Dim actualColumnHigh As Long

    If Not GetMatrixBounds(values, actualRowLow, actualRowHigh, actualColumnLow, actualColumnHigh) Then Exit Function
    MatrixBoundsMatch = (actualRowLow = rowLow And actualRowHigh = rowHigh And actualColumnLow = columnLow And actualColumnHigh = columnHigh)
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

Private Function MatrixOrScalarValue(ByRef values As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long) As Variant
    If IsArray(values) Then
        MatrixOrScalarValue = values(LBound(values, 1) + rowIndex - 1, LBound(values, 2) + columnIndex - 1)
    Else
        MatrixOrScalarValue = values
    End If
End Function

Private Function IsSupportedCellValue(ByVal value As Variant) As Boolean
    IsSupportedCellValue = (VarType(value) = vbEmpty Or VarType(value) = vbString)
End Function

Private Function IsBlankHeader(ByVal value As Variant) As Boolean
    If IsEmpty(value) Then
        IsBlankHeader = True
    ElseIf VarType(value) = vbString Then
        IsBlankHeader = (Len(CStr(value)) = 0)
    End If
End Function

Private Function IsMeaningfulValue(ByVal value As Variant) As Boolean
    If IsEmpty(value) Then Exit Function
    If VarType(value) = vbString Then IsMeaningfulValue = (Len(CStr(value)) > 0) Else IsMeaningfulValue = True
End Function

Private Function IndexArraysEqual(ByRef leftValues As Variant, ByRef rightValues As Variant) As Boolean
    Dim leftLow As Long
    Dim leftHigh As Long
    Dim leftColumnLow As Long
    Dim leftColumnHigh As Long
    Dim rightLow As Long
    Dim rightHigh As Long
    Dim rightColumnLow As Long
    Dim rightColumnHigh As Long
    Dim index As Long

    If IsEmpty(leftValues) Or IsEmpty(rightValues) Then
        IndexArraysEqual = (IsEmpty(leftValues) And IsEmpty(rightValues))
        Exit Function
    End If
    If Not GetMatrixBounds(leftValues, leftLow, leftHigh, leftColumnLow, leftColumnHigh) Then
        If Not IsArray(leftValues) Then Exit Function
        On Error Resume Next
        leftLow = LBound(leftValues, 1)
        leftHigh = UBound(leftValues, 1)
        If Err.Number <> 0 Then Err.Clear: Exit Function
        index = LBound(leftValues, 2)
        If Err.Number <> 9 Then Err.Clear: Exit Function
        Err.Clear
        On Error GoTo 0
        leftColumnLow = 0
        leftColumnHigh = 0
    End If
    If Not GetMatrixBounds(rightValues, rightLow, rightHigh, rightColumnLow, rightColumnHigh) Then
        If Not IsArray(rightValues) Then Exit Function
        On Error Resume Next
        rightLow = LBound(rightValues, 1)
        rightHigh = UBound(rightValues, 1)
        If Err.Number <> 0 Then Err.Clear: Exit Function
        index = LBound(rightValues, 2)
        If Err.Number <> 9 Then Err.Clear: Exit Function
        Err.Clear
        On Error GoTo 0
        rightColumnLow = 0
        rightColumnHigh = 0
    End If
    If leftHigh - leftLow <> rightHigh - rightLow Then Exit Function
    For index = 0 To leftHigh - leftLow
        If CLng(leftValues(leftLow + index)) <> CLng(rightValues(rightLow + index)) Then Exit Function
    Next index
    IndexArraysEqual = True
End Function

Private Function AnchorMatricesEqual(ByRef leftValues As Variant, ByRef rightValues As Variant) As Boolean
    Dim leftRowLow As Long
    Dim leftRowHigh As Long
    Dim leftColumnLow As Long
    Dim leftColumnHigh As Long
    Dim rightRowLow As Long
    Dim rightRowHigh As Long
    Dim rightColumnLow As Long
    Dim rightColumnHigh As Long
    Dim rowIndex As Long
    Dim columnIndex As Long

    If IsEmpty(leftValues) Or IsEmpty(rightValues) Then
        AnchorMatricesEqual = (IsEmpty(leftValues) And IsEmpty(rightValues))
        Exit Function
    End If
    If Not GetMatrixBounds(leftValues, leftRowLow, leftRowHigh, leftColumnLow, leftColumnHigh) Then Exit Function
    If Not GetMatrixBounds(rightValues, rightRowLow, rightRowHigh, rightColumnLow, rightColumnHigh) Then Exit Function
    If leftRowLow <> rightRowLow Or leftRowHigh <> rightRowHigh Or leftColumnLow <> rightColumnLow Or leftColumnHigh <> rightColumnHigh Then Exit Function
    For rowIndex = leftRowLow To leftRowHigh
        For columnIndex = leftColumnLow To leftColumnHigh
            If Not AnchorValuesEqual(leftValues(rowIndex, columnIndex), rightValues(rowIndex, columnIndex)) Then Exit Function
        Next columnIndex
    Next rowIndex
    AnchorMatricesEqual = True
End Function

Private Function AnchorValuesEqual(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If VarType(leftValue) <> VarType(rightValue) Then Exit Function
    If IsEmpty(leftValue) Then
        AnchorValuesEqual = True
    ElseIf VarType(leftValue) = vbString Then
        AnchorValuesEqual = (StrComp(CStr(leftValue), CStr(rightValue), vbBinaryCompare) = 0)
    ElseIf IsError(leftValue) Or IsNull(leftValue) Then
        AnchorValuesEqual = False
    Else
        AnchorValuesEqual = (leftValue = rightValue)
    End If
End Function

Private Sub ResetSnapshot(ByRef snapshot As SnapshotData)
    snapshot.SheetName = vbNullString
    snapshot.HeaderFingerprint = vbNullString
    snapshot.FirstDataRow = 0
    snapshot.LastDataRow = 0
    snapshot.Targets = Empty
    snapshot.DeclaredColumnIndexes = Empty
    snapshot.AnchorColumnIndexes = Empty
    snapshot.AnchorValues = Empty
    snapshot.Values = Empty
    snapshot.Fonts = Empty
    snapshot.Fills = Empty
End Sub
