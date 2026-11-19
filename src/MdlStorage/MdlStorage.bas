Attribute VB_Name = "MdlStorage"
Option Explicit

Private Const STORAGE_ERR_INVALID As Long = 1401
Private Const STORAGE_ERR_IO As Long = 1402
Private Const STORAGE_ERR_EXISTS As Long = 1403
Private Const STORAGE_FILE_ATTRIBUTES As Long = vbNormal Or vbHidden Or vbSystem Or vbReadOnly
Private Const STORAGE_DIRECTORY_ATTRIBUTES As Long = vbDirectory Or vbHidden Or vbSystem
Private Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400

Public Function StorageCreateDirectory(ByVal path As String, ByRef created As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim attributes As Long

    created = False
    On Error GoTo Failed
    If Len(path) = 0 Then GoTo InvalidInput
    If PathExists(path, attributes) Then
        If (attributes And vbDirectory) = vbDirectory Then
            StorageCreateDirectory = True
            errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StorageCreateDirectory", "")
            Exit Function
        End If
        errorInfo = MakeAppError(STORAGE_ERR_EXISTS, "MdlStorage.StorageCreateDirectory", "A file already uses the requested directory path.", path)
        Exit Function
    End If

    MkDir path
    created = True
    StorageCreateDirectory = True
    errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StorageCreateDirectory", "")
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(STORAGE_ERR_INVALID, "MdlStorage.StorageCreateDirectory", "A directory path is required.", path)
    Exit Function

Failed:
    If PathExists(path, attributes) And (attributes And vbDirectory) = vbDirectory Then
        created = False
        StorageCreateDirectory = True
        errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StorageCreateDirectory", "")
    Else
        errorInfo = MakeAppError(STORAGE_ERR_IO, "MdlStorage.StorageCreateDirectory", "The directory could not be created.", CStr(Err.Number) & ": " & Err.Description)
    End If
End Function

Public Function StoragePathExists(ByVal path As String, ByRef exists As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim attributes As Long
    Dim errorNumber As Long

    exists = False
    On Error GoTo Failed
    If Len(path) = 0 Then GoTo InvalidInput
    attributes = GetAttr(path)
    exists = True
    StoragePathExists = True
    errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StoragePathExists", "")
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(STORAGE_ERR_INVALID, "MdlStorage.StoragePathExists", "A path is required.")
    Exit Function

Failed:
    errorNumber = Err.Number
    Err.Clear
    If errorNumber = 53 Or errorNumber = 76 Then
        StoragePathExists = True
        errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StoragePathExists", "")
    Else
        errorInfo = MakeAppError(STORAGE_ERR_IO, "MdlStorage.StoragePathExists", "The path status could not be read safely.", CStr(errorNumber))
    End If
End Function

Public Function StorageReadText(ByVal path As String, ByRef contents As String, ByRef lastWrite As Date, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim fileNumber As Integer
    Dim opened As Boolean

    contents = vbNullString
    lastWrite = 0
    On Error GoTo Failed
    fileNumber = FreeFile
    Open path For Binary Access Read Lock Write As #fileNumber
    opened = True
    If LOF(fileNumber) > 0 Then contents = Input$(LOF(fileNumber), #fileNumber)
    lastWrite = FileDateTime(path)
    Close #fileNumber
    opened = False
    StorageReadText = True
    errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StorageReadText", "")
    Exit Function

Failed:
    If opened Then Close #fileNumber
    errorInfo = MakeAppError(STORAGE_ERR_IO, "MdlStorage.StorageReadText", "The metadata file could not be read.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function StorageWriteText(ByVal path As String, ByVal contents As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim fileNumber As Integer
    Dim opened As Boolean

    On Error GoTo Failed
    fileNumber = FreeFile
    Open path For Output Access Write Lock Read Write As #fileNumber
    opened = True
    Print #fileNumber, contents;
    Close #fileNumber
    opened = False
    StorageWriteText = True
    errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StorageWriteText", "")
    Exit Function

Failed:
    If opened Then Close #fileNumber
    errorInfo = MakeAppError(STORAGE_ERR_IO, "MdlStorage.StorageWriteText", "The metadata file could not be written.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function StorageRenameDirectory(ByVal sourcePath As String, ByVal destinationPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
    On Error GoTo Failed
    If Not IsDirectory(sourcePath) Then GoTo InvalidInput
    If PathExists(destinationPath) Then
        errorInfo = MakeAppError(STORAGE_ERR_EXISTS, "MdlStorage.StorageRenameDirectory", "The destination directory already exists.", destinationPath)
        Exit Function
    End If
    Name sourcePath As destinationPath
    StorageRenameDirectory = True
    errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StorageRenameDirectory", "")
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(STORAGE_ERR_INVALID, "MdlStorage.StorageRenameDirectory", "The source directory does not exist.", sourcePath)
    Exit Function

Failed:
    errorInfo = MakeAppError(STORAGE_ERR_IO, "MdlStorage.StorageRenameDirectory", "The directory could not be renamed.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function StorageRemoveLockDirectory(ByVal path As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim ownerPath As String

    On Error GoTo Failed
    If Not IsDirectory(path) Then
        StorageRemoveLockDirectory = True
        errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StorageRemoveLockDirectory", "")
        Exit Function
    End If
    ownerPath = JoinPath(path, "owner.txt")
    If PathExists(ownerPath) Then Kill ownerPath
    RmDir path
    StorageRemoveLockDirectory = True
    errorInfo = MakeAppError(ERR_NONE, "MdlStorage.StorageRemoveLockDirectory", "")
    Exit Function

Failed:
    errorInfo = MakeAppError(STORAGE_ERR_IO, "MdlStorage.StorageRemoveLockDirectory", "The lock directory could not be removed safely.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function CreateLocalWorkspace(ByVal sessionId As String, ByRef workspacePath As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rootPath As String
    Dim sessionsPath As String
    Dim created As Boolean

    workspacePath = vbNullString
    If Not IsSafeSessionId(sessionId) Then GoTo InvalidInput
    rootPath = GetLocalAppDataFolder()
    If Len(rootPath) = 0 Then GoTo InvalidInput
    sessionsPath = JoinPath(rootPath, "sessions")
    If Not EnsureLocalAppDirectory(rootPath, errorInfo) Then Exit Function
    If Not EnsureLocalAppDirectory(sessionsPath, errorInfo) Then Exit Function
    workspacePath = JoinPath(sessionsPath, sessionId)
    If Not StorageCreateDirectory(workspacePath, created, errorInfo) Then Exit Function
    If Not created Then
        errorInfo = MakeAppError(STORAGE_ERR_EXISTS, "MdlStorage.CreateLocalWorkspace", "A workspace already exists for this session.", workspacePath)
        workspacePath = vbNullString
        Exit Function
    End If
    CreateLocalWorkspace = True
    errorInfo = MakeAppError(ERR_NONE, "MdlStorage.CreateLocalWorkspace", "")
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(STORAGE_ERR_INVALID, "MdlStorage.CreateLocalWorkspace", "A safe session identifier and LOCALAPPDATA path are required.")
End Function

Public Function CopySharedFileToWorkspace(ByVal sourceUncPath As String, ByVal workspacePath As String, ByRef localPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim normalizedSource As String
    Dim fileName As String
    Dim slashAt As Long
    Dim attributes As Long

    localPath = vbNullString
    On Error GoTo Failed
    If Not ValidateUncPath(sourceUncPath, normalizedSource, errorInfo) Then Exit Function
    If Not IsLocalAppDataPath(workspacePath) Or Not IsDirectory(workspacePath) Then GoTo InvalidInput
    If Not PathExists(normalizedSource, attributes) Or (attributes And vbDirectory) = vbDirectory Then GoTo InvalidInput
    slashAt = InStrRev(normalizedSource, "\")
    If slashAt <= 2 Or slashAt = Len(normalizedSource) Then GoTo InvalidInput
    fileName = Mid$(normalizedSource, slashAt + 1)
    localPath = JoinPath(workspacePath, fileName)
    If PathExists(localPath) Then GoTo InvalidInput
    FileCopy normalizedSource, localPath
    CopySharedFileToWorkspace = True
    errorInfo = MakeAppError(ERR_NONE, "MdlStorage.CopySharedFileToWorkspace", "")
    Exit Function

InvalidInput:
    localPath = vbNullString
    errorInfo = MakeAppError(STORAGE_ERR_INVALID, "MdlStorage.CopySharedFileToWorkspace", "The source must be an existing UNC file and the workspace must be local.", sourceUncPath)
    Exit Function

Failed:
    localPath = vbNullString
    errorInfo = MakeAppError(STORAGE_ERR_IO, "MdlStorage.CopySharedFileToWorkspace", "The shared file could not be copied to the local workspace.", CStr(Err.Number) & ": " & Err.Description)
End Function

Public Function RemoveLocalWorkspace(ByVal workspacePath As String, ByRef errorInfo As AppErrorInfo) As Boolean
    On Error GoTo Failed
    If Not IsLocalAppDataPath(workspacePath) Then GoTo InvalidInput
    If Not IsDirectory(workspacePath) Then
        RemoveLocalWorkspace = True
        errorInfo = MakeAppError(ERR_NONE, "MdlStorage.RemoveLocalWorkspace", "")
        Exit Function
    End If
    If Not DeleteLocalTree(workspacePath) Then GoTo Failed
    RemoveLocalWorkspace = True
    errorInfo = MakeAppError(ERR_NONE, "MdlStorage.RemoveLocalWorkspace", "")
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(STORAGE_ERR_INVALID, "MdlStorage.RemoveLocalWorkspace", "Only a workspace below LOCALAPPDATA can be removed.", workspacePath)
    Exit Function

Failed:
    errorInfo = MakeAppError(STORAGE_ERR_IO, "MdlStorage.RemoveLocalWorkspace", "The local workspace could not be removed safely.", CStr(Err.Number) & ": " & Err.Description)
End Function

Private Function DeleteLocalTree(ByVal path As String) As Boolean
    Dim entries As Collection
    Dim entryName As String
    Dim childPath As String
    Dim entry As Variant
    Dim attributes As Long

    On Error GoTo Failed
    Set entries = New Collection
    entryName = Dir$(JoinPath(path, "*"), STORAGE_FILE_ATTRIBUTES Or STORAGE_DIRECTORY_ATTRIBUTES)
    Do While Len(entryName) > 0
        If entryName <> "." And entryName <> ".." Then entries.Add entryName
        entryName = Dir$()
    Loop

    For Each entry In entries
        childPath = JoinPath(path, CStr(entry))
        attributes = GetAttr(childPath)
        If (attributes And FILE_ATTRIBUTE_REPARSE_POINT) <> 0 Then Exit Function
        If (attributes And vbDirectory) = vbDirectory Then
            If Not DeleteLocalTree(childPath) Then Exit Function
        Else
            Kill childPath
        End If
    Next entry
    RmDir path
    DeleteLocalTree = True
    Exit Function

Failed:
    DeleteLocalTree = False
End Function

Private Function IsSafeSessionId(ByVal sessionId As String) As Boolean
    Dim index As Long
    Dim character As String

    If Len(sessionId) < 1 Or Len(sessionId) > 64 Then Exit Function
    For index = 1 To Len(sessionId)
        character = Mid$(sessionId, index, 1)
        If Not ((character >= "0" And character <= "9") Or (character >= "A" And character <= "Z") Or (character >= "a" And character <= "z") Or character = "-") Then Exit Function
    Next index
    IsSafeSessionId = True
End Function

Private Function IsDirectory(ByVal path As String) As Boolean
    Dim attributes As Long
    If Not PathExists(path, attributes) Then Exit Function
    IsDirectory = ((attributes And vbDirectory) = vbDirectory)
End Function

Private Function PathExists(ByVal path As String, Optional ByRef attributes As Long = 0) As Boolean
    On Error GoTo Missing
    attributes = GetAttr(path)
    PathExists = True
    Exit Function
Missing:
    attributes = 0
End Function

Private Function JoinPath(ByVal basePath As String, ByVal child As String) As String
    If Len(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) = "\" Then
        JoinPath = basePath & child
    Else
        JoinPath = basePath & "\" & child
    End If
End Function
