Attribute VB_Name = "MdlInfra"
Option Explicit

Public Enum LogLevel
    LogDebug = 0
    LogInfo = 1
    LogWarning = 2
    LogError = 3
End Enum

Public Type AppErrorInfo
    Code As Long
    Source As String
    Message As String
    Context As String
End Type

Public Type AppConfig
    WorkbookPaths As Variant
    Version As Long
End Type

Public Enum SharedPathStatus
    SharedPathAvailable = 0
    SharedPathNotFound = 1
    SharedPathAccessDenied = 2
    SharedPathNetworkUnavailable = 3
    SharedPathIsDirectory = 4
    SharedPathUnknown = 5
End Enum

Private Type GuidData
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Const ERR_NONE As Long = 0
Public Const ERR_INVALID_INPUT As Long = 1001
Public Const ERR_CONFIG_MISSING As Long = 1002
Public Const ERR_CONFIG_INVALID As Long = 1003
Public Const ERR_IO As Long = 1004
Public Const ERR_UNC_REQUIRED As Long = 1005
Public Const ERR_NOT_FOUND As Long = 1007

Public Const FONT_FORMAT_PROPERTY_COUNT As Long = 14
Public Const FILL_FORMAT_PROPERTY_COUNT As Long = 9

Private Const LOG_MAX_BYTES As Long = 1048576
Private Const LOG_FILE_NAME As String = "app.log"
Private Const CONFIG_FILE_NAME As String = "config.ini"
Private Const INVALID_FILE_ATTRIBUTES As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef guid As GuidData) As Long
Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal fileName As LongPtr) As Long
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
#Else
Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef guid As GuidData) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal fileName As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
#End If

Public Function MakeAppError(ByVal code As Long, ByVal source As String, ByVal message As String, Optional ByVal context As String = "") As AppErrorInfo
    Dim value As AppErrorInfo
    value.Code = code
    value.Source = source
    value.Message = message
    value.Context = context
    MakeAppError = value
End Function

Public Function NewGuidString() As String
    Dim value As GuidData
    Dim tail As String
    Dim index As Long

    If CoCreateGuid(value) <> 0 Then Exit Function
    For index = 0 To 7
        tail = tail & Right$("0" & Hex$(value.Data4(index)), 2)
    Next index
    NewGuidString = Right$("00000000" & Hex$(value.Data1), 8) & "-" & _
                    Right$("0000" & Hex$(CLng(value.Data2) And &HFFFF&), 4) & "-" & _
                    Right$("0000" & Hex$(CLng(value.Data3) And &HFFFF&), 4) & "-" & _
                    Left$(tail, 4) & "-" & Mid$(tail, 5)
End Function

Public Function EnsureLocalAppDirectory(ByVal path As String, ByRef errorInfo As AppErrorInfo) As Boolean
    EnsureLocalAppDirectory = EnsureLocalDirectory(path, errorInfo)
    If EnsureLocalAppDirectory Then errorInfo = MakeAppError(ERR_NONE, "MdlInfra.EnsureLocalAppDirectory", "")
End Function

Public Function IsLocalAppDataPath(ByVal path As String) As Boolean
    IsLocalAppDataPath = IsUnderLocalAppData(path)
End Function

Public Function GetLocalAppDataFolder() As String
    Dim basePath As String
    basePath = Environ$("LOCALAPPDATA")
    If Len(basePath) = 0 Then Exit Function
    GetLocalAppDataFolder = JoinPath(basePath, "MDDTLock")
End Function

Public Function GetDefaultConfigPath() As String
    Dim appFolder As String
    appFolder = GetLocalAppDataFolder()
    If Len(appFolder) > 0 Then GetDefaultConfigPath = JoinPath(appFolder, CONFIG_FILE_NAME)
End Function

Public Function ValidateUncPath(ByVal path As String, ByRef normalizedPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim value As String
    Dim serverAndShare As String
    Dim slashAt As Long

    normalizedPath = vbNullString
    value = Replace$(Trim$(path), "/", "\")
    If Len(value) < 5 Or Left$(value, 2) <> "\\" Then
        errorInfo = MakeAppError(ERR_UNC_REQUIRED, "MdlInfra.ValidateUncPath", "A UNC path is required.", path)
        Exit Function
    End If
    If Left$(value, 4) = "\\?\" Or Left$(value, 4) = "\\.\" Then
        errorInfo = MakeAppError(ERR_UNC_REQUIRED, "MdlInfra.ValidateUncPath", "Device paths are not accepted as configuration paths.", path)
        Exit Function
    End If

    serverAndShare = Mid$(value, 3)
    slashAt = InStr(1, serverAndShare, "\", vbBinaryCompare)
    If slashAt < 2 Or slashAt = Len(serverAndShare) Then
        errorInfo = MakeAppError(ERR_UNC_REQUIRED, "MdlInfra.ValidateUncPath", "The UNC path must include a server and share name.", path)
        Exit Function
    End If
    If InStr(slashAt + 1, serverAndShare, "\", vbBinaryCompare) = slashAt + 1 Then
        errorInfo = MakeAppError(ERR_UNC_REQUIRED, "MdlInfra.ValidateUncPath", "The UNC share name cannot be empty.", path)
        Exit Function
    End If

    normalizedPath = value
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.ValidateUncPath", "")
    ValidateUncPath = True
End Function

Public Function InspectUncPath(ByVal path As String, ByRef status As SharedPathStatus, ByRef systemError As Long, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim normalized As String
    Dim attributes As Long

    status = SharedPathUnknown
    systemError = 0
    If Not ValidateUncPath(path, normalized, errorInfo) Then Exit Function

    attributes = GetFileAttributesW(StrPtr(normalized))
    If attributes <> INVALID_FILE_ATTRIBUTES Then
        If (attributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            status = SharedPathIsDirectory
        Else
            status = SharedPathAvailable
        End If
        InspectUncPath = True
        errorInfo = MakeAppError(ERR_NONE, "MdlInfra.InspectUncPath", "")
        Exit Function
    End If

    systemError = GetLastError()
    status = ClassifySharedPathError(systemError)
    InspectUncPath = True
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.InspectUncPath", "", CStr(systemError))
End Function

Public Function ClassifySharedPathError(ByVal systemError As Long) As SharedPathStatus
    Select Case systemError
        Case 2, 3
            ClassifySharedPathError = SharedPathNotFound
        Case 5
            ClassifySharedPathError = SharedPathAccessDenied
        Case 53, 64, 67, 121, 1231, 1232, 1236
            ClassifySharedPathError = SharedPathNetworkUnavailable
        Case Else
            ClassifySharedPathError = SharedPathUnknown
    End Select
End Function

Public Function ReadConfig(ByVal configPath As String, ByRef config As AppConfig, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim fileNumber As Integer
    Dim opened As Boolean
    Dim lineText As String
    Dim trimmedLine As String
    Dim separatorAt As Long
    Dim key As String
    Dim value As String
    Dim legacyPath As String
    Dim rawPaths As Collection
    Dim pathValues() As Variant
    Dim index As Long
    Dim versionSeen As Boolean
    Dim legacySeen As Boolean
    Dim invalidConfig As Boolean

    On Error GoTo Failed
    config.WorkbookPaths = Empty
    config.Version = 0
    If Not IsUnderLocalAppData(configPath) Then
        errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.ReadConfig", "Configuration must be stored below LOCALAPPDATA.", configPath)
        Exit Function
    End If
    If Len(Dir$(configPath, vbNormal Or vbHidden Or vbSystem)) = 0 Then
        errorInfo = MakeAppError(ERR_CONFIG_MISSING, "MdlInfra.ReadConfig", "Configuration file is missing. Create config.ini with Version=2 and one or more UNC WorkbookPath= entries.", configPath)
        Exit Function
    End If

    Set rawPaths = New Collection
    fileNumber = FreeFile
    Open configPath For Input Access Read As #fileNumber
    opened = True
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineText
        If Len(lineText) > 0 Then
            If Left$(lineText, 3) = Chr$(&HEF) & Chr$(&HBB) & Chr$(&HBF) Then
                lineText = Mid$(lineText, 4)
            ElseIf AscW(Left$(lineText, 1)) = -257 Then
                lineText = Mid$(lineText, 2)
            End If
        End If
        trimmedLine = Trim$(lineText)
        If Len(trimmedLine) = 0 Or Left$(trimmedLine, 1) = ";" Or Left$(trimmedLine, 1) = "#" Then GoTo NextLine
        separatorAt = InStr(1, lineText, "=", vbBinaryCompare)
        If separatorAt <= 1 Then
            invalidConfig = True
            Exit Do
        End If
        key = LCase$(Trim$(Left$(lineText, separatorAt - 1)))
        value = Trim$(Mid$(lineText, separatorAt + 1))
        Select Case key
            Case "version"
                If versionSeen Or Not IsNumeric(value) Then
                    invalidConfig = True
                    Exit Do
                End If
                If CDbl(value) <> 1 And CDbl(value) <> 2 Then
                    invalidConfig = True
                    Exit Do
                End If
                config.Version = CLng(value)
                versionSeen = True
            Case "workbookpath"
                rawPaths.Add value
            Case "mainfilepath"
                If legacySeen Then
                    invalidConfig = True
                    Exit Do
                End If
                legacyPath = value
                legacySeen = True
            Case Else
                invalidConfig = True
                Exit Do
        End Select
NextLine:
    Loop
    Close #fileNumber
    opened = False

    If invalidConfig Or Not versionSeen Then GoTo InvalidConfig
    Select Case config.Version
        Case 1
            If rawPaths.Count <> 0 Or Not legacySeen Or Len(legacyPath) = 0 Then GoTo InvalidConfig
            ReDim pathValues(1 To 1)
            pathValues(1) = legacyPath
        Case 2
            If legacySeen Or rawPaths.Count = 0 Then GoTo InvalidConfig
            ReDim pathValues(1 To rawPaths.Count)
            For index = 1 To rawPaths.Count
                pathValues(index) = CStr(rawPaths.Item(index))
            Next index
        Case Else
            GoTo InvalidConfig
    End Select

    If Not NormalizeConfigPaths(pathValues, config.WorkbookPaths, errorInfo) Then GoTo InvalidConfig
    ReadConfig = True
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.ReadConfig", "")
    Exit Function

InvalidConfig:
    config.WorkbookPaths = Empty
    errorInfo = MakeAppError(ERR_CONFIG_INVALID, "MdlInfra.ReadConfig", "Configuration is invalid. Use Version=2 with unique UNC WorkbookPath= entries.", configPath)
    Exit Function

Failed:
    If opened Then Close #fileNumber
    If Err.Number <> 0 Then errorInfo = MakeAppError(ERR_IO, "MdlInfra.ReadConfig", "Configuration could not be read.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function WriteConfig(ByVal configPath As String, ByRef config As AppConfig, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim normalizedPaths As Variant
    Dim folder As String
    Dim fileNumber As Integer
    Dim opened As Boolean
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long

    On Error GoTo Failed
    If Not NormalizeConfigPaths(config.WorkbookPaths, normalizedPaths, errorInfo) Then Exit Function
    If Not IsUnderLocalAppData(configPath) Then
        errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.WriteConfig", "Configuration must be stored below LOCALAPPDATA.", configPath)
        Exit Function
    End If

    folder = ParentPath(configPath)
    If Not EnsureLocalDirectory(folder, errorInfo) Then Exit Function
    fileNumber = FreeFile
    Open configPath For Output Access Write As #fileNumber
    opened = True
    Print #fileNumber, "Version=2"
    lowerBound = LBound(normalizedPaths)
    upperBound = UBound(normalizedPaths)
    For index = lowerBound To upperBound
        Print #fileNumber, "WorkbookPath=" & CStr(normalizedPaths(index))
    Next index
    Close #fileNumber
    opened = False
    WriteConfig = True
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.WriteConfig", "")
    Exit Function

Failed:
    If opened Then Close #fileNumber
    errorInfo = MakeAppError(ERR_IO, "MdlInfra.WriteConfig", "Configuration could not be written.", CStr(Err.Number) & ": " & Err.Description)
End Function

Private Function NormalizeConfigPaths(ByRef sourcePaths As Variant, ByRef normalizedPaths As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long
    Dim priorIndex As Long
    Dim normalized As String
    Dim result() As Variant

    normalizedPaths = Empty
    If Not IsArray(sourcePaths) Then GoTo InvalidConfigPaths
    On Error GoTo InvalidConfigPaths
    lowerBound = LBound(sourcePaths)
    upperBound = UBound(sourcePaths)
    If upperBound < lowerBound Then GoTo InvalidConfigPaths
    ReDim result(1 To upperBound - lowerBound + 1)

    For index = lowerBound To upperBound
        If VarType(sourcePaths(index)) <> vbString Then GoTo InvalidConfigPaths
        If Not ValidateUncPath(CStr(sourcePaths(index)), normalized, errorInfo) Then GoTo InvalidConfigPaths
        For priorIndex = 1 To index - lowerBound
            If StrComp(CStr(result(priorIndex)), normalized, vbTextCompare) = 0 Then GoTo InvalidConfigPaths
        Next priorIndex
        result(index - lowerBound + 1) = normalized
    Next index

    normalizedPaths = result
    NormalizeConfigPaths = True
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.NormalizeConfigPaths", "")
    Exit Function

InvalidConfigPaths:
    normalizedPaths = Empty
    errorInfo = MakeAppError(ERR_CONFIG_INVALID, "MdlInfra.NormalizeConfigPaths", "Workbook paths must be unique UNC file paths.")
End Function

Public Function ProbeUncLatency(ByVal path As String, ByRef elapsedMs As Double, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim normalized As String
    Dim startedAt As Double
    Dim attributes As Long

    On Error GoTo Failed
    If Not ValidateUncPath(path, normalized, errorInfo) Then Exit Function
    startedAt = Timer
    attributes = GetAttr(normalized)
    elapsedMs = Timer - startedAt
    If elapsedMs < 0 Then elapsedMs = elapsedMs + 86400#
    elapsedMs = elapsedMs * 1000#
    ProbeUncLatency = True
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.ProbeUncLatency", "")
    Exit Function

Failed:
    elapsedMs = 0
    errorInfo = MakeAppError(ERR_IO, "MdlInfra.ProbeUncLatency", "The existing UNC target could not be read.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function CompareCellValues(ByVal leftValue As Variant, ByVal rightValue As Variant, ByRef equal As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim leftType As VbVarType
    Dim rightType As VbVarType

    equal = False
    leftType = VarType(leftValue)
    rightType = VarType(rightValue)
    If Not IsSupportedCellType(leftType) Or Not IsSupportedCellType(rightType) Then
        errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.CompareCellValues", "Cell values must be strings or Empty.")
        Exit Function
    End If

    If leftType = vbEmpty Or rightType = vbEmpty Then
        equal = (leftType = vbEmpty And rightType = vbEmpty)
    Else
        equal = (StrComp(CStr(leftValue), CStr(rightValue), vbBinaryCompare) = 0)
    End If
    CompareCellValues = True
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.CompareCellValues", "")
End Function

Public Function EncodeFormatField(ByVal value As String) As String
    EncodeFormatField = CStr(Len(value)) & ":" & value
End Function

Public Function DecodeFormatField(ByRef data As String, ByRef cursor As Long, ByRef value As String) As Boolean
    Dim separatorAt As Long
    Dim fieldLength As Long
    Dim lengthText As String
    Dim index As Long

    value = vbNullString
    If cursor < 1 Or cursor > Len(data) + 1 Then Exit Function
    separatorAt = InStr(cursor, data, ":", vbBinaryCompare)
    If separatorAt <= cursor Then Exit Function
    lengthText = Mid$(data, cursor, separatorAt - cursor)
    For index = 1 To Len(lengthText)
        If Mid$(lengthText, index, 1) < "0" Or Mid$(lengthText, index, 1) > "9" Then Exit Function
    Next index
    On Error GoTo InvalidField
    fieldLength = CLng(lengthText)
    If fieldLength < 0 Or separatorAt + fieldLength > Len(data) + 1 Then Exit Function
    value = Mid$(data, separatorAt + 1, fieldLength)
    cursor = separatorAt + fieldLength + 1
    DecodeFormatField = True
    Exit Function

InvalidField:
    value = vbNullString
End Function

Public Function BuildFontSignature(ByRef properties As Variant, ByRef errorInfo As AppErrorInfo) As String
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long
    Dim fieldValue As String
    Dim result As String

    If Not GetVectorBounds(properties, lowerBound, upperBound) Then GoTo InvalidProperties
    If upperBound - lowerBound + 1 <> FONT_FORMAT_PROPERTY_COUNT Then GoTo InvalidProperties
    result = "FP1|"
    For index = lowerBound To upperBound
        If Not EncodeFormatScalar(properties(index), fieldValue) Then GoTo InvalidProperties
        result = result & EncodeFormatField(fieldValue)
    Next index
    BuildFontSignature = result
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.BuildFontSignature", "")
    Exit Function

InvalidProperties:
    errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.BuildFontSignature", "Font properties must contain the supported fixed-order scalar fields.")
End Function

Public Function BuildTextFontData(ByVal cellValue As Variant, ByRef runs As Variant, ByVal emptyCellFont As String, ByRef errorInfo As AppErrorInfo) As String
    Dim valueEqual As Boolean
    Dim valueError As AppErrorInfo
    Dim runLow As Long
    Dim runHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim rowIndex As Long
    Dim startValue As Long
    Dim lengthValue As Long
    Dim expectedStart As Long
    Dim fontSignature As String
    Dim encodedValue As String
    Dim result As String

    If Not CompareCellValues(cellValue, cellValue, valueEqual, valueError) Then GoTo InvalidData
    If VarType(cellValue) = vbEmpty Then
        encodedValue = "E"
    Else
        encodedValue = "S"
    End If
    If VarType(cellValue) = vbEmpty Then
        result = "TF1|" & EncodeFormatField(encodedValue) & EncodeFormatField(vbNullString)
    Else
        result = "TF1|" & EncodeFormatField(encodedValue) & EncodeFormatField(CStr(cellValue))
    End If

    If VarType(cellValue) = vbString And Len(CStr(cellValue)) > 0 Then
        If Not GetMatrixBounds(runs, runLow, runHigh, columnLow, columnHigh) Then GoTo InvalidData
        If columnHigh - columnLow + 1 <> 3 Then GoTo InvalidData
        result = result & EncodeFormatField(CStr(runHigh - runLow + 1))
        expectedStart = 1
        For rowIndex = runLow To runHigh
            If Not IsNumeric(runs(rowIndex, columnLow)) Or Not IsNumeric(runs(rowIndex, columnLow + 1)) Then GoTo InvalidData
            startValue = CLng(runs(rowIndex, columnLow))
            lengthValue = CLng(runs(rowIndex, columnLow + 1))
            If CDbl(runs(rowIndex, columnLow)) <> CDbl(startValue) Or CDbl(runs(rowIndex, columnLow + 1)) <> CDbl(lengthValue) Then GoTo InvalidData
            If startValue <> expectedStart Or lengthValue < 1 Then GoTo InvalidData
            fontSignature = CStr(runs(rowIndex, columnLow + 2))
            If Not IsValidFontSignature(fontSignature) Then GoTo InvalidData
            result = result & EncodeFormatField(CStr(startValue)) & EncodeFormatField(CStr(lengthValue)) & EncodeFormatField(fontSignature)
            expectedStart = expectedStart + lengthValue
        Next rowIndex
        If expectedStart <> Len(CStr(cellValue)) + 1 Then GoTo InvalidData
    Else
        If Not IsEmpty(runs) Then GoTo InvalidData
        If Not IsValidFontSignature(emptyCellFont) Then GoTo InvalidData
        result = result & EncodeFormatField("0") & EncodeFormatField(emptyCellFont)
    End If

    BuildTextFontData = result
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.BuildTextFontData", "")
    Exit Function

InvalidData:
    BuildTextFontData = vbNullString
    errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.BuildTextFontData", "Text and font runs are invalid or do not cover the full string.")
End Function

Public Function BuildFillData(ByRef properties As Variant, ByRef errorInfo As AppErrorInfo) As String
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long
    Dim fieldValue As String
    Dim result As String

    If Not GetVectorBounds(properties, lowerBound, upperBound) Then GoTo InvalidProperties
    If upperBound - lowerBound + 1 <> FILL_FORMAT_PROPERTY_COUNT Then GoTo InvalidProperties
    result = "FL1|"
    For index = lowerBound To upperBound
        If Not EncodeFormatScalar(properties(index), fieldValue) Then GoTo InvalidProperties
        result = result & EncodeFormatField(fieldValue)
    Next index
    BuildFillData = result
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.BuildFillData", "")
    Exit Function

InvalidProperties:
    errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.BuildFillData", "Fill properties must contain the supported fixed-order scalar fields.")
End Function

Public Function ParseTextFontData(ByVal data As String, ByRef valueKind As String, ByRef textValue As String, ByRef runs As Variant, ByRef emptyCellFont As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim cursor As Long
    Dim fieldValue As String
    Dim runCount As Long
    Dim runIndex As Long
    Dim startValue As Long
    Dim lengthValue As Long
    Dim expectedStart As Long
    Dim runValues() As Variant

    valueKind = vbNullString
    textValue = vbNullString
    runs = Empty
    emptyCellFont = vbNullString
    If Left$(data, 4) <> "TF1|" Then GoTo InvalidData
    cursor = 5
    If Not DecodeFormatField(data, cursor, valueKind) Then GoTo InvalidData
    If valueKind <> "E" And valueKind <> "S" Then GoTo InvalidData
    If Not DecodeFormatField(data, cursor, textValue) Then GoTo InvalidData
    If valueKind = "E" And Len(textValue) > 0 Then GoTo InvalidData
    If Not DecodeFormatField(data, cursor, fieldValue) Then GoTo InvalidData
    If Not IsNumeric(fieldValue) Then GoTo InvalidData
    runCount = CLng(fieldValue)
    If CStr(runCount) <> fieldValue Or runCount < 0 Then GoTo InvalidData

    If runCount = 0 Then
        If Len(textValue) > 0 Then GoTo InvalidData
        If Not DecodeFormatField(data, cursor, emptyCellFont) Then GoTo InvalidData
        If Not IsValidFontSignature(emptyCellFont) Then GoTo InvalidData
    Else
        If valueKind <> "S" Or Len(textValue) = 0 Then GoTo InvalidData
        ReDim runValues(1 To runCount, 1 To 3)
        expectedStart = 1
        For runIndex = 1 To runCount
            If Not DecodeFormatField(data, cursor, fieldValue) Then GoTo InvalidData
            If Not IsNumeric(fieldValue) Then GoTo InvalidData
            startValue = CLng(fieldValue)
            If CStr(startValue) <> fieldValue Then GoTo InvalidData
            If Not DecodeFormatField(data, cursor, fieldValue) Then GoTo InvalidData
            If Not IsNumeric(fieldValue) Then GoTo InvalidData
            lengthValue = CLng(fieldValue)
            If CStr(lengthValue) <> fieldValue Then GoTo InvalidData
            If startValue <> expectedStart Or lengthValue < 1 Then GoTo InvalidData
            If Not DecodeFormatField(data, cursor, fieldValue) Then GoTo InvalidData
            If Not IsValidFontSignature(fieldValue) Then GoTo InvalidData
            runValues(runIndex, 1) = startValue
            runValues(runIndex, 2) = lengthValue
            runValues(runIndex, 3) = fieldValue
            expectedStart = expectedStart + lengthValue
        Next runIndex
        If expectedStart <> Len(textValue) + 1 Then GoTo InvalidData
        runs = runValues
    End If
    If cursor <> Len(data) + 1 Then GoTo InvalidData
    ParseTextFontData = True
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.ParseTextFontData", "")
    Exit Function

InvalidData:
    runs = Empty
    errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.ParseTextFontData", "The text/font snapshot is malformed or unsupported.")
End Function

Public Function ParseFontSignature(ByVal data As String, ByRef properties As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    ParseFontSignature = ParseFixedProperties(data, "FP1|", FONT_FORMAT_PROPERTY_COUNT, properties, errorInfo, "MdlInfra.ParseFontSignature")
End Function

Public Function ParseFillData(ByVal data As String, ByRef properties As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    ParseFillData = ParseFixedProperties(data, "FL1|", FILL_FORMAT_PROPERTY_COUNT, properties, errorInfo, "MdlInfra.ParseFillData")
End Function

Public Function DecodeFormatScalar(ByVal encodedValue As String, ByRef value As Variant, ByRef isApplicable As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rawValue As String
    Dim cursor As Long

    value = Empty
    isApplicable = True
    Select Case Left$(encodedValue, 1)
        Case "A"
            If encodedValue <> "A" Then GoTo InvalidScalar
            isApplicable = False
        Case "E"
            If encodedValue <> "E" Then GoTo InvalidScalar
        Case "B"
            If encodedValue = "B1" Then
                value = True
            ElseIf encodedValue = "B0" Then
                value = False
            Else
                GoTo InvalidScalar
            End If
        Case "N"
            rawValue = Mid$(encodedValue, 2)
            If Len(rawValue) = 0 Or Not IsNumeric(rawValue) Then GoTo InvalidScalar
            On Error GoTo InvalidScalar
            value = CDbl(rawValue)
        Case "S"
            cursor = 2
            If Not DecodeFormatField(encodedValue, cursor, rawValue) Then GoTo InvalidScalar
            If cursor <> Len(encodedValue) + 1 Then GoTo InvalidScalar
            value = rawValue
        Case Else
            GoTo InvalidScalar
    End Select
    DecodeFormatScalar = True
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.DecodeFormatScalar", "")
    Exit Function

InvalidScalar:
    value = Empty
    isApplicable = False
    errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.DecodeFormatScalar", "A normalized format property is malformed.")
End Function

Public Function CompareFormattingData(ByVal leftValue As Variant, ByVal rightValue As Variant, ByRef equal As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim leftText As String
    Dim rightText As String

    equal = False
    If VarType(leftValue) <> vbString Or VarType(rightValue) <> vbString Then GoTo InvalidData
    leftText = CStr(leftValue)
    rightText = CStr(rightValue)
    If Not IsCanonicalFormatData(leftText) Or Not IsCanonicalFormatData(rightText) Then GoTo InvalidData
    If Left$(leftText, 4) <> Left$(rightText, 4) Then GoTo InvalidData
    equal = (StrComp(leftText, rightText, vbBinaryCompare) = 0)
    CompareFormattingData = True
    errorInfo = MakeAppError(ERR_NONE, "MdlInfra.CompareFormattingData", "")
    Exit Function

InvalidData:
    errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.CompareFormattingData", "Formatting snapshots must be valid canonical font or fill data of the same component type.")
End Function

Public Function FormatFormattingPreview(ByVal data As Variant) As String
    Dim valueKind As String
    Dim textValue As String
    Dim emptyCellFont As String
    Dim runs As Variant
    Dim properties As Variant
    Dim errorInfo As AppErrorInfo
    Dim index As Long
    Dim description As String

    If VarType(data) <> vbString Then
        FormatFormattingPreview = "(format unavailable)"
        Exit Function
    End If
    If Left$(CStr(data), 4) = "TF1|" Then
        If Not ParseTextFontData(CStr(data), valueKind, textValue, runs, emptyCellFont, errorInfo) Then
            FormatFormattingPreview = "(unsupported font data)"
            Exit Function
        End If
        If valueKind = "E" Then
            description = "Text: (blank)"
        ElseIf Len(textValue) = 0 Then
            description = "Text: (zero-length)"
        Else
            description = "Text: " & textValue
        End If
        If IsEmpty(runs) Then
            description = description & vbCrLf & "Cell font: " & DescribeFontSignature(emptyCellFont)
        Else
            For index = LBound(runs, 1) To UBound(runs, 1)
                description = description & vbCrLf & "Characters " & CStr(runs(index, 1)) & "-" & CStr(CLng(runs(index, 1)) + CLng(runs(index, 2)) - 1) & ": " & DescribeFontSignature(CStr(runs(index, 3)))
            Next index
        End If
        FormatFormattingPreview = description
        Exit Function
    End If
    If Left$(CStr(data), 4) = "FL1|" Then
        If Not ParseFillData(CStr(data), properties, errorInfo) Then
            FormatFormattingPreview = "(unsupported fill data)"
            Exit Function
        End If
        FormatFormattingPreview = "Fill: " & DescribePropertyList(properties, Array("Pattern", "ColorIndex", "Color", "ThemeColor", "TintAndShade", "PatternColorIndex", "PatternColor", "PatternThemeColor", "PatternTintAndShade"))
    Else
        FormatFormattingPreview = "(unsupported formatting data)"
    End If
End Function

Private Function DescribeFontSignature(ByVal data As String) As String
    Dim properties As Variant
    Dim errorInfo As AppErrorInfo

    If Not ParseFontSignature(data, properties, errorInfo) Then
        DescribeFontSignature = "(unsupported font data)"
    Else
        DescribeFontSignature = DescribePropertyList(properties, Array("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Superscript", "Subscript", "OutlineFont", "Shadow", "ColorIndex", "Color", "ThemeColor", "TintAndShade"))
    End If
End Function

Private Function DescribePropertyList(ByRef properties As Variant, ByRef names As Variant) As String
    Dim index As Long
    Dim value As Variant
    Dim applicable As Boolean
    Dim errorInfo As AppErrorInfo
    Dim result As String

    For index = LBound(properties) To UBound(properties)
        If index > LBound(properties) Then result = result & "; "
        result = result & CStr(names(index - LBound(properties))) & "="
        If DecodeFormatScalar(CStr(properties(index)), value, applicable, errorInfo) Then
            If applicable Then
                If IsEmpty(value) Then
                    result = result & "(empty)"
                Else
                    result = result & CStr(value)
                End If
            Else
                result = result & "(not applicable)"
            End If
        Else
            result = result & "(invalid)"
        End If
    Next index
    DescribePropertyList = result
End Function

Private Function ParseFixedProperties(ByVal data As String, ByVal signature As String, ByVal expectedCount As Long, ByRef properties As Variant, ByRef errorInfo As AppErrorInfo, ByVal source As String) As Boolean
    Dim cursor As Long
    Dim index As Long
    Dim fieldValue As String
    Dim result() As Variant

    properties = Empty
    If Left$(data, Len(signature)) <> signature Then GoTo InvalidData
    ReDim result(1 To expectedCount)
    cursor = Len(signature) + 1
    For index = 1 To expectedCount
        If Not DecodeFormatField(data, cursor, fieldValue) Then GoTo InvalidData
        If Not IsValidFormatScalar(fieldValue) Then GoTo InvalidData
        result(index) = fieldValue
    Next index
    If cursor <> Len(data) + 1 Then GoTo InvalidData
    properties = result
    ParseFixedProperties = True
    errorInfo = MakeAppError(ERR_NONE, source, "")
    Exit Function

InvalidData:
    properties = Empty
    errorInfo = MakeAppError(ERR_INVALID_INPUT, source, "The normalized format property record is malformed or unsupported.")
End Function

Private Function IsValidFontSignature(ByVal data As String) As Boolean
    Dim properties As Variant
    Dim errorInfo As AppErrorInfo
    IsValidFontSignature = ParseFontSignature(data, properties, errorInfo)
End Function

Private Function IsCanonicalFormatData(ByVal data As String) As Boolean
    Dim properties As Variant
    Dim valueKind As String
    Dim textValue As String
    Dim emptyCellFont As String
    Dim runs As Variant
    Dim errorInfo As AppErrorInfo

    Select Case Left$(data, 4)
        Case "TF1|"
            IsCanonicalFormatData = ParseTextFontData(data, valueKind, textValue, runs, emptyCellFont, errorInfo)
        Case "FL1|"
            IsCanonicalFormatData = ParseFillData(data, properties, errorInfo)
    End Select
End Function

Private Function EncodeFormatScalar(ByVal value As Variant, ByRef encodedValue As String) As Boolean
    Select Case VarType(value)
        Case vbEmpty
            encodedValue = "E"
        Case vbString
            encodedValue = "S" & EncodeFormatField(CStr(value))
        Case vbError
            encodedValue = "A"
        Case vbBoolean
            encodedValue = IIf(CBool(value), "B1", "B0")
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            encodedValue = "N" & CStr(CDbl(value))
        Case Else
            Exit Function
    End Select
    EncodeFormatScalar = True
End Function

Private Function IsValidFormatScalar(ByVal encodedValue As String) As Boolean
    Dim value As Variant
    Dim applicable As Boolean
    Dim errorInfo As AppErrorInfo
    IsValidFormatScalar = DecodeFormatScalar(encodedValue, value, applicable, errorInfo)
End Function

Private Function GetVectorBounds(ByRef values As Variant, ByRef lowerBound As Long, ByRef upperBound As Long) As Boolean
    Dim probe As Long

    If Not IsArray(values) Then Exit Function
    On Error Resume Next
    lowerBound = LBound(values, 1)
    upperBound = UBound(values, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    probe = LBound(values, 2)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    GetVectorBounds = (upperBound >= lowerBound)
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

Public Sub LogWrite(ByVal level As LogLevel, ByVal source As String, ByVal message As String)
    Dim appFolder As String
    Dim logFolder As String
    Dim logPath As String
    Dim fileNumber As Integer

    On Error Resume Next
    appFolder = GetLocalAppDataFolder()
    If Len(appFolder) = 0 Then Exit Sub
    logFolder = JoinPath(appFolder, "logs")
    If Not DirectoryExists(logFolder) Then EnsureDirectoryTree logFolder
    logPath = JoinPath(logFolder, LOG_FILE_NAME)
    If Len(Dir$(logPath, vbNormal Or vbHidden Or vbSystem)) > 0 Then
        If FileLen(logPath) >= LOG_MAX_BYTES Then
            If Len(Dir$(logPath & ".1", vbNormal Or vbHidden Or vbSystem)) > 0 Then Kill logPath & ".1"
            Name logPath As logPath & ".1"
        End If
    End If
    fileNumber = FreeFile
    Open logPath For Append Access Write As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & CStr(level) & vbTab & source & vbTab & message
    Close #fileNumber
    On Error GoTo 0
End Sub

Private Function IsSupportedCellType(ByVal valueType As VbVarType) As Boolean
    IsSupportedCellType = (valueType = vbEmpty Or valueType = vbString)
End Function

Private Function IsUnderLocalAppData(ByVal path As String) As Boolean
    Dim rootPath As String
    Dim normalized As String

    rootPath = Replace$(Trim$(Environ$("LOCALAPPDATA")), "/", "\")
    normalized = Replace$(Trim$(path), "/", "\")
    If Len(rootPath) = 0 Or PathHasParentTraversal(rootPath) Or PathHasParentTraversal(normalized) Then Exit Function
    Do While Len(rootPath) > 3 And Right$(rootPath, 1) = "\"
        rootPath = Left$(rootPath, Len(rootPath) - 1)
    Loop
    If Len(rootPath) = 0 Or Len(normalized) <= Len(rootPath) Then Exit Function
    IsUnderLocalAppData = (StrComp(Left$(normalized, Len(rootPath)), rootPath, vbTextCompare) = 0 And Mid$(normalized, Len(rootPath) + 1, 1) = "\")
End Function

Private Function PathHasParentTraversal(ByVal path As String) As Boolean
    Dim parts() As String
    Dim index As Long

    parts = Split(Replace$(path, "/", "\"), "\")
    For index = LBound(parts) To UBound(parts)
        If Trim$(parts(index)) = ".." Then
            PathHasParentTraversal = True
            Exit Function
        End If
    Next index
End Function

Private Function EnsureLocalDirectory(ByVal path As String, ByRef errorInfo As AppErrorInfo) As Boolean
    If Not IsUnderLocalAppData(path & "\sentinel") Then
        errorInfo = MakeAppError(ERR_INVALID_INPUT, "MdlInfra.EnsureLocalDirectory", "Local directories must remain below LOCALAPPDATA.", path)
        Exit Function
    End If
    EnsureLocalDirectory = EnsureDirectoryTree(path)
    If Not EnsureLocalDirectory Then errorInfo = MakeAppError(ERR_IO, "MdlInfra.EnsureLocalDirectory", "A local application directory could not be created.", path)
End Function

Private Function EnsureDirectoryTree(ByVal path As String) As Boolean
    Dim parent As String
    Dim separatorAt As Long

    On Error GoTo Failed
    If DirectoryExists(path) Then
        EnsureDirectoryTree = True
        Exit Function
    End If
    separatorAt = InStrRev(path, "\")
    If separatorAt <= 0 Then Exit Function
    parent = Left$(path, separatorAt - 1)
    If Len(parent) > 0 And Not DirectoryExists(parent) Then
        If Not EnsureDirectoryTree(parent) Then Exit Function
    End If
    MkDir path
    EnsureDirectoryTree = True
    Exit Function

Failed:
    If DirectoryExists(path) Then EnsureDirectoryTree = True
End Function

Private Function DirectoryExists(ByVal path As String) As Boolean
    Dim attributes As Long
    On Error GoTo Missing
    attributes = GetAttr(path)
    DirectoryExists = ((attributes And vbDirectory) = vbDirectory)
    Exit Function
Missing:
    DirectoryExists = False
End Function

Private Function ParentPath(ByVal path As String) As String
    Dim separatorAt As Long
    separatorAt = InStrRev(path, "\")
    If separatorAt > 0 Then ParentPath = Left$(path, separatorAt - 1)
End Function

Private Function JoinPath(ByVal basePath As String, ByVal child As String) As String
    If Len(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) = "\" Then
        JoinPath = basePath & child
    Else
        JoinPath = basePath & "\" & child
    End If
End Function
