Attribute VB_Name = "TestMdlStorage"
Option Explicit

Public Sub RunMdlStorageTests()
    RunTestCase "CreateAndReadText"
    RunTestCase "RenameAndRemoveLockDirectory"
    RunTestCase "LocalWorkspaceLifecycle"
    RunTestCase "RejectUnsafePaths"
End Sub

Private Sub RunTestCase(ByVal caseName As String)
    On Error GoTo Failed
    Select Case caseName
        Case "CreateAndReadText"
            TestCreateAndReadText
        Case "RenameAndRemoveLockDirectory"
            TestRenameAndRemoveLockDirectory
        Case "LocalWorkspaceLifecycle"
            TestLocalWorkspaceLifecycle
        Case "RejectUnsafePaths"
            TestRejectUnsafePaths
        Case Else
            Err.Raise vbObjectError + 2101, "TestMdlStorage", "Unknown test case: " & caseName
    End Select
    Exit Sub

Failed:
    TestSupport.RecordTestFailure "TestMdlStorage." & caseName & ": " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestCreateAndReadText()
    Dim root As String
    Dim path As String
    Dim created As Boolean
    Dim contents As String
    Dim lastWrite As Date
    Dim errorInfo As AppErrorInfo

    root = TestRoot()
    path = root & "\metadata"
    AssertTrue EnsureLocalAppDirectory(GetLocalAppDataFolder() & "\tests", errorInfo), "The local test parent should exist."
    AssertTrue StorageCreateDirectory(root, created, errorInfo), "A local test root should be created. " & CStr(errorInfo.Code) & ": " & errorInfo.Context
    AssertTrue created, "The first directory creation should report created."
    AssertTrue StorageCreateDirectory(root, created, errorInfo), "An existing directory should be accepted."
    AssertFalse created, "An existing directory should report not-created."
    AssertTrue StorageCreateDirectory(path, created, errorInfo), "A child directory should be created."
    AssertTrue StorageWriteText(path & "\owner.txt", "Version=1" & vbCrLf & "Token=abc", errorInfo), "Metadata text should be written."
    AssertTrue StorageReadText(path & "\owner.txt", contents, lastWrite, errorInfo), "Metadata text should be readable."
    AssertEqual contents, "Version=1" & vbCrLf & "Token=abc", "Metadata should round-trip exactly."
    AssertTrue (lastWrite > 0), "A metadata read should return a modification time."
    AssertTrue RemoveLocalWorkspace(root, errorInfo), "The temporary test tree should be removable."
End Sub

Private Sub TestRenameAndRemoveLockDirectory()
    Dim root As String
    Dim lockPath As String
    Dim quarantinePath As String
    Dim created As Boolean
    Dim errorInfo As AppErrorInfo

    root = TestRoot()
    lockPath = root & "\lock"
    quarantinePath = root & "\reclaim-test"
    AssertTrue EnsureLocalAppDirectory(GetLocalAppDataFolder() & "\tests", errorInfo), "The local test parent should exist."
    AssertTrue StorageCreateDirectory(root, created, errorInfo), "A local test root should be created. " & CStr(errorInfo.Code) & ": " & errorInfo.Context
    AssertTrue StorageCreateDirectory(lockPath, created, errorInfo), "A lock directory should be created."
    AssertTrue StorageWriteText(lockPath & "\owner.txt", "Token=test", errorInfo), "The owner record should be written."
    AssertTrue StorageRenameDirectory(lockPath, quarantinePath, errorInfo), "A lock directory should be renamable for reclamation."
    AssertTrue StorageRemoveLockDirectory(quarantinePath, errorInfo), "A quarantined lock directory should be removable."
    AssertTrue RemoveLocalWorkspace(root, errorInfo), "The temporary test root should be removable."
End Sub

Private Sub TestLocalWorkspaceLifecycle()
    Dim workspace As String
    Dim errorInfo As AppErrorInfo
    Dim sessionId As String

    sessionId = NewGuidString()
    AssertTrue CreateLocalWorkspace(sessionId, workspace, errorInfo), "A session workspace should be created."
    AssertTrue IsLocalAppDataPath(workspace), "A session workspace must be below LOCALAPPDATA."
    AssertTrue RemoveLocalWorkspace(workspace, errorInfo), "A session workspace should be removable."
End Sub

Private Sub TestRejectUnsafePaths()
    Dim workspace As String
    Dim errorInfo As AppErrorInfo

    workspace = Environ$("LOCALAPPDATA") & "\..\Documents\mddt-workspace"
    AssertFalse RemoveLocalWorkspace(workspace, errorInfo), "Workspace cleanup must reject a path that escapes LOCALAPPDATA."
End Sub

Private Function TestRoot() As String
    TestRoot = GetLocalAppDataFolder() & "\tests\storage-" & NewGuidString()
End Function

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2102, "TestMdlStorage", message
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, ByVal message As String)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 2103, "TestMdlStorage", message
End Sub
