Attribute VB_Name = "TestMdlInfra"
Option Explicit

Public Sub RunMdlInfraTests()
    On Error GoTo Failed
    TestUncValidation
    TestPathResolution
    TestCellValueComparison
    TestConfigRoundTrip
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

Private Sub TestPathResolution()
    Dim uncPath As String
    Dim errorInfo As AppErrorInfo

    AssertTrue ResolvePathToUnc("\\server\share\book.xlsx", uncPath, errorInfo), "A UNC path should pass resolution unchanged."
    AssertEqual uncPath, "\\server\share\book.xlsx", "Resolved UNC path changed unexpectedly."
    AssertFalse ResolvePathToUnc("relative\book.xlsx", uncPath, errorInfo), "Relative paths must not resolve to UNC."
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

    folder = GetLocalAppDataFolder() & "\tests\infra-" & Format$(Now, "yyyymmddhhnnss")
    configPath = folder & "\config.ini"
    sourceConfig.MainFilePath = "\\server\share\study.xlsx"
    sourceConfig.Version = 1

    AssertTrue WriteConfig(configPath, sourceConfig, errorInfo), "Config should write below LOCALAPPDATA."
    AssertTrue ReadConfig(configPath, loadedConfig, errorInfo), "Config should read after writing."
    AssertEqual loadedConfig.MainFilePath, sourceConfig.MainFilePath, "Config path did not round-trip."
    AssertEqual CStr(loadedConfig.Version), "1", "Config version did not round-trip."

    Kill configPath
    RmDir folder
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
