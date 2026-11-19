Attribute VB_Name = "TestMdlInfra"
Option Explicit

Public Sub RunMdlInfraTests()
    On Error GoTo Failed
    TestUncValidation
    TestPathErrorClassification
    TestCellValueComparison
    TestConfigRoundTrip
    TestLegacyConfigRead
    TestConfigRejectsMappedAndDuplicatePaths
    TestMissingConfigReportsPath
    TestLocalPathGuard
    TestGuidGeneration
    Exit Sub

Failed:
    TestSupport.RecordTestFailure "TestMdlInfra: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestUncValidation()
    Dim normalized As String
    Dim errorInfo As AppErrorInfo

    AssertTrue ValidateUncPath("\\server\share\folder\book.xlsx", normalized, errorInfo), "UNC path should validate."
    AssertEqual normalized, "\\server\share\folder\book.xlsx", "UNC path should be normalized."
    AssertTrue ValidateUncPath("//server/share/book.xlsx", normalized, errorInfo), "Forward slashes should normalize to UNC separators."
    AssertEqual normalized, "\\server\share\book.xlsx", "Forward slash path normalization is incorrect."
    AssertFalse ValidateUncPath("X:\folder\book.xlsx", normalized, errorInfo), "Mapped drives must not be accepted as config paths."
    AssertFalse ValidateUncPath("\server", normalized, errorInfo), "A server without a share must be rejected."
End Sub

Private Sub TestPathErrorClassification()
    AssertEqual CStr(ClassifySharedPathError(2)), CStr(SharedPathNotFound), "A missing file error should be classified as not found."
    AssertEqual CStr(ClassifySharedPathError(3)), CStr(SharedPathNotFound), "A missing directory error should be classified as not found."
    AssertEqual CStr(ClassifySharedPathError(5)), CStr(SharedPathAccessDenied), "Access denied should retain its own category."
    AssertEqual CStr(ClassifySharedPathError(53)), CStr(SharedPathNetworkUnavailable), "A bad network path should be classified as unavailable."
    AssertEqual CStr(ClassifySharedPathError(1231)), CStr(SharedPathNetworkUnavailable), "A network-unreachable error should be classified as unavailable."
    AssertEqual CStr(ClassifySharedPathError(9999)), CStr(SharedPathUnknown), "Unrecognized system errors must remain unknown."
End Sub

Private Sub TestCellValueComparison()
    Dim equal As Boolean
    Dim errorInfo As AppErrorInfo

    AssertTrue CompareCellValues("value", "value", equal, errorInfo), "String comparison should succeed."
    AssertTrue equal, "Identical strings should compare equal."
    AssertTrue CompareCellValues("Value", "value", equal, errorInfo), "Case-sensitive comparison should succeed."
    AssertFalse equal, "String comparison must preserve case."
    AssertTrue CompareCellValues("value ", "value", equal, errorInfo), "Whitespace comparison should succeed."
    AssertFalse equal, "String comparison must preserve spaces."
    AssertTrue CompareCellValues(Empty, Empty, equal, errorInfo), "Empty comparison should succeed."
    AssertTrue equal, "Two Empty values should compare equal."
    AssertTrue CompareCellValues(Empty, vbNullString, equal, errorInfo), "Empty and empty text are comparable."
    AssertFalse equal, "Empty and zero-length text must remain distinct."
    AssertFalse CompareCellValues(1, 1, equal, errorInfo), "Numeric cell values must be rejected."
End Sub

Private Sub TestConfigRoundTrip()
    Dim folder As String
    Dim configPath As String
    Dim sourceConfig As AppConfig
    Dim loadedConfig As AppConfig
    Dim errorInfo As AppErrorInfo
    Dim failureNumber As Long
    Dim failureDescription As String

    folder = GetLocalAppDataFolder() & "\tests\infra-" & Format$(Now, "yyyymmddhhnnss")
    configPath = folder & "\config.ini"
    sourceConfig.WorkbookPaths = Array("\\server\share\study.xlsx", "\\server\share\team-b\study.xlsx")
    sourceConfig.Version = 2

    On Error GoTo Failed
    AssertTrue WriteConfig(configPath, sourceConfig, errorInfo), "Config should write below LOCALAPPDATA."
    AssertTrue ReadConfig(configPath, loadedConfig, errorInfo), "Config should read after writing."
    AssertEqual CStr(UBound(loadedConfig.WorkbookPaths) - LBound(loadedConfig.WorkbookPaths) + 1), "2", "Both configured workbook paths should round-trip."
    AssertEqual CStr(loadedConfig.WorkbookPaths(1)), "\\server\share\study.xlsx", "The first config path did not round-trip."
    AssertEqual CStr(loadedConfig.WorkbookPaths(2)), "\\server\share\team-b\study.xlsx", "The second config path did not round-trip."
    AssertEqual CStr(loadedConfig.Version), "2", "The config writer should emit version 2."

    GoTo Cleanup

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
Cleanup:
    CleanupConfigTestFile folder, configPath
    If failureNumber <> 0 Then Err.Raise failureNumber, "TestMdlInfra.TestConfigRoundTrip", failureDescription
End Sub

Private Sub TestLegacyConfigRead()
    Dim folder As String
    Dim configPath As String
    Dim fileNumber As Integer
    Dim loadedConfig As AppConfig
    Dim errorInfo As AppErrorInfo
    Dim failureNumber As Long
    Dim failureDescription As String

    folder = GetLocalAppDataFolder() & "\tests\infra-legacy-" & CStr(CLng(Timer * 1000))
    configPath = folder & "\config.ini"
    On Error GoTo Failed
    AssertTrue EnsureLocalAppDirectory(folder, errorInfo), "The legacy config test folder should be created locally."
    fileNumber = FreeFile
    Open configPath For Output Access Write As #fileNumber
    Print #fileNumber, "Version=1"
    Print #fileNumber, "MainFilePath=\\server\share\legacy.xlsx"
    Close #fileNumber

    AssertTrue ReadConfig(configPath, loadedConfig, errorInfo), "A version 1 config should remain readable."
    AssertEqual CStr(loadedConfig.Version), "1", "Legacy config should preserve its source version."
    AssertEqual CStr(UBound(loadedConfig.WorkbookPaths) - LBound(loadedConfig.WorkbookPaths) + 1), "1", "Legacy config should expose one workbook path."
    AssertEqual CStr(loadedConfig.WorkbookPaths(1)), "\\server\share\legacy.xlsx", "Legacy main file path should map into the workbook list."

    GoTo Cleanup

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
Cleanup:
    CleanupConfigTestFile folder, configPath
    If failureNumber <> 0 Then Err.Raise failureNumber, "TestMdlInfra.TestLegacyConfigRead", failureDescription
End Sub

Private Sub TestConfigRejectsMappedAndDuplicatePaths()
    Dim folder As String
    Dim configPath As String
    Dim fileNumber As Integer
    Dim config As AppConfig
    Dim loadedConfig As AppConfig
    Dim errorInfo As AppErrorInfo
    Dim failureNumber As Long
    Dim failureDescription As String

    folder = GetLocalAppDataFolder() & "\tests\infra-invalid"
    configPath = folder & "\config.ini"
    On Error GoTo Failed
    config.WorkbookPaths = Array("X:\folder\book.xlsx")
    AssertFalse WriteConfig(configPath, config, errorInfo), "Mapped drive paths must not be written to config."

    config.WorkbookPaths = Array("\\server\share\book.xlsx", "\\SERVER\SHARE\book.xlsx")
    AssertFalse WriteConfig(configPath, config, errorInfo), "Duplicate UNC paths must be rejected case-insensitively."

    AssertTrue EnsureLocalAppDirectory(folder, errorInfo), "The invalid config test folder should be created locally."
    fileNumber = FreeFile
    Open configPath For Output Access Write As #fileNumber
    Print #fileNumber, "Version=2"
    Print #fileNumber, "WorkbookPath=\\server\share\book.xlsx"
    Print #fileNumber, "WorkbookPath=\\SERVER\SHARE\book.xlsx"
    Close #fileNumber
    AssertFalse ReadConfig(configPath, loadedConfig, errorInfo), "A manually edited duplicate path must make the config invalid."
    GoTo Cleanup

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
Cleanup:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    On Error GoTo 0
    CleanupConfigTestFile folder, configPath
    If failureNumber <> 0 Then Err.Raise failureNumber, "TestMdlInfra.TestConfigRejectsMappedAndDuplicatePaths", failureDescription
End Sub

Private Sub TestMissingConfigReportsPath()
    Dim configPath As String
    Dim config As AppConfig
    Dim errorInfo As AppErrorInfo

    configPath = GetLocalAppDataFolder() & "\tests\missing-" & NewGuidString() & "\config.ini"
    AssertFalse ReadConfig(configPath, config, errorInfo), "A missing config file should prevent startup."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_CONFIG_MISSING), "A missing config should use its dedicated error code."
    AssertEqual errorInfo.Context, configPath, "The missing config error should identify the exact local path."
End Sub

Private Sub CleanupConfigTestFile(ByVal folder As String, ByVal configPath As String)
    On Error Resume Next
    If Len(Dir$(configPath, vbNormal Or vbHidden Or vbSystem)) > 0 Then Kill configPath
    If Len(Dir$(folder, vbDirectory)) > 0 Then RmDir folder
    On Error GoTo 0
End Sub

Private Sub TestLocalPathGuard()
    Dim path As String
    Dim errorInfo As AppErrorInfo

    path = Environ$("LOCALAPPDATA") & "\..\Documents\mddt-config.ini"
    AssertFalse EnsureLocalAppDirectory(path, errorInfo), "A parent traversal must not escape LOCALAPPDATA."
End Sub

Private Sub TestGuidGeneration()
    Dim firstGuid As String
    Dim secondGuid As String

    firstGuid = NewGuidString()
    secondGuid = NewGuidString()
    AssertEqual CStr(Len(firstGuid)), "36", "A generated GUID should use the standard textual form."
    AssertEqual Mid$(firstGuid, 9, 1), "-", "A GUID should contain its first separator at position 9."
    AssertEqual Mid$(firstGuid, 14, 1), "-", "A GUID should contain its second separator at position 14."
    AssertFalse (StrComp(firstGuid, secondGuid, vbBinaryCompare) = 0), "Two generated GUIDs should not match."
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 1701, "TestMdlInfra", message
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, ByVal message As String)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 1702, "TestMdlInfra", message
End Sub
