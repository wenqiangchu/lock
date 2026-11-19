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
