VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CLockCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const OWNER_FILE_NAME As String = "owner.txt"
Private Const LOCK_ROOT_NAME As String = ".mddt-locks"
Private Const OWNER_VERSION As String = "1"
Private Const DEFAULT_COLUMN_TTL_SECONDS As Long = 300
Private Const DEFAULT_GATE_TTL_SECONDS As Long = 600
Private Const ERR_LOCK_BUSY As Long = 1501
Private Const ERR_LOCK_INVALID As Long = 1502
Private Const ERR_LOCK_OWNERSHIP As Long = 1503
Private Const ERR_LOCK_EXPIRED As Long = 1504
Private Const ERR_LOCK_STORAGE As Long = 1505

Private mTestRoot As String
Private mColumnTtlSeconds As Long
Private mGateTtlSeconds As Long
Private mTestClockOffsetSeconds As Long

Private Sub Class_Initialize()
    mColumnTtlSeconds = DEFAULT_COLUMN_TTL_SECONDS
    mGateTtlSeconds = DEFAULT_GATE_TTL_SECONDS
End Sub

Friend Sub SetTestRoot(ByVal path As String)
    Dim errorInfo As AppErrorInfo

    If Not IsLocalAppDataPath(path) Then Err.Raise vbObjectError + 2201, "CLockCoordinator.SetTestRoot", "Test root must remain below LOCALAPPDATA."
    If Not EnsureLocalAppDirectory(path, errorInfo) Then Err.Raise vbObjectError + 2202, "CLockCoordinator.SetTestRoot", errorInfo.Message
    mTestRoot = path
End Sub

Friend Sub SetTestTtl(ByVal columnSeconds As Long, ByVal gateSeconds As Long)
    If columnSeconds < 1 Or gateSeconds < 1 Then Err.Raise vbObjectError + 2203, "CLockCoordinator.SetTestTtl", "Test TTL values must be positive."
    mColumnTtlSeconds = columnSeconds
    mGateTtlSeconds = gateSeconds
End Sub

Friend Sub SetTestClockOffset(ByVal seconds As Long)
    mTestClockOffsetSeconds = seconds
End Sub

Friend Function AcquireLocks(ByRef targets As Variant, ByVal sessionId As String, ByRef leases As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim targetCount As Long
    Dim rowIndex As Long
    Dim targetIndex As Long
    Dim priorIndex As Long
    Dim lockPath As String
    Dim targetKey As String
    Dim token As String
    Dim created As Boolean
    Dim acquired() As Boolean
    Dim targetKeys() As String
    Dim lockPaths() As String
    Dim tokens() As String
    Dim originalRows() As Long
    Dim order() As Long
    Dim result() As Variant
    Dim ownerError As AppErrorInfo
    Dim cleanupError As AppErrorInfo

    leases = Empty
    If Not IsSafeIdentifier(sessionId) Then GoTo InvalidInput
    If Not GetMatrixBounds(targets, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidInput
    If columnHigh - columnLow + 1 <> 3 Then GoTo InvalidInput
    targetCount = rowHigh - rowLow + 1
    If targetCount < 1 Then GoTo InvalidInput

    ReDim acquired(1 To targetCount)
    ReDim targetKeys(1 To targetCount)
    ReDim lockPaths(1 To targetCount)
    ReDim tokens(1 To targetCount)
    ReDim originalRows(1 To targetCount)
    ReDim order(1 To targetCount)

    targetIndex = 0
    For rowIndex = rowLow To rowHigh
        targetIndex = targetIndex + 1
        originalRows(targetIndex) = rowIndex
        If VarType(targets(rowIndex, columnLow)) <> vbString Or VarType(targets(rowIndex, columnLow + 1)) <> vbString Or VarType(targets(rowIndex, columnLow + 2)) <> vbString Then GoTo InvalidInput
        If Not BuildTargetKey(CStr(targets(rowIndex, columnLow)), CStr(targets(rowIndex, columnLow + 1)), CStr(targets(rowIndex, columnLow + 2)), targetKey, lockPath, errorInfo) Then Exit Function
        targetKeys(targetIndex) = targetKey
        lockPaths(targetIndex) = lockPath
        For priorIndex = 1 To targetIndex - 1
            If StrComp(targetKeys(priorIndex), targetKey, vbBinaryCompare) = 0 Then GoTo InvalidInput
        Next priorIndex
        order(targetIndex) = targetIndex
    Next rowIndex
    SortTargetOrder order, targetKeys, targetCount

    For targetIndex = 1 To targetCount
        rowIndex = order(targetIndex)
        If Not TryAcquireLockDirectory(lockPaths(rowIndex), targetKeys(rowIndex), sessionId, mColumnTtlSeconds, token, errorInfo) Then GoTo Rollback
        tokens(rowIndex) = token
        acquired(rowIndex) = True
    Next targetIndex

    ReDim result(1 To targetCount, 1 To 4)
    For targetIndex = 1 To targetCount
        result(targetIndex, 1) = lockPaths(targetIndex)
        result(targetIndex, 2) = targetKeys(targetIndex)
        result(targetIndex, 3) = sessionId
        result(targetIndex, 4) = tokens(targetIndex)
    Next targetIndex
    leases = result
    AcquireLocks = True
    errorInfo = MakeAppError(ERR_NONE, "MdlLock.AcquireLocks", "")
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.AcquireLocks", "Targets must be unique UNC file, sheet, and AMG header rows.")
    Exit Function

Rollback:
    For targetIndex = 1 To targetCount
        If acquired(targetIndex) Then StorageRemoveLockDirectory lockPaths(targetIndex), cleanupError
    Next targetIndex
    leases = Empty
End Function

Friend Function RenewLocks(ByRef leases As Variant, ByRef renewed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim rowIndex As Long
    Dim probeBound As Long
    Dim fields As Variant
    Dim currentError As AppErrorInfo

    renewed = False
    If IsEmpty(leases) Then
        renewed = True
        RenewLocks = True
        errorInfo = MakeAppError(ERR_NONE, "MdlLock.RenewLocks", "")
        Exit Function
    End If
    If Not GetMatrixBounds(leases, rowLow, rowHigh, columnLow, columnHigh) Or columnHigh - columnLow + 1 <> 4 Then GoTo InvalidInput

    renewed = True
    For rowIndex = rowLow To rowHigh
        fields = LeaseFieldsFromMatrix(leases, rowIndex, columnLow)
        If Not RenewLeaseFields(fields, mColumnTtlSeconds, currentError) Then
            renewed = False
            If errorInfo.Code = ERR_NONE Then errorInfo = currentError
        End If
    Next rowIndex
    RenewLocks = True
    If renewed Then errorInfo = MakeAppError(ERR_NONE, "MdlLock.RenewLocks", "")
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.RenewLocks", "Lease records must be an N by 4 matrix.")
End Function

Friend Function ValidateLease(ByRef leaseFields As Variant, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim firstIndex As Long
    Dim lastIndex As Long

    valid = False
    If Not GetVectorBounds(leaseFields, firstIndex, lastIndex) Or lastIndex - firstIndex + 1 <> 4 Then GoTo InvalidInput
    If Not ValidateLeaseScalars(CStr(leaseFields(firstIndex)), CStr(leaseFields(firstIndex + 1)), CStr(leaseFields(firstIndex + 2)), CStr(leaseFields(firstIndex + 3)), mColumnTtlSeconds, valid, errorInfo) Then Exit Function
    ValidateLease = True
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.ValidateLease", "A lease must contain path, target key, session ID, and fencing token.")
End Function

Friend Function InspectTargetLock(ByRef targetFields As Variant, ByRef status As LockAvailability, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim firstIndex As Long
    Dim lastIndex As Long
    Dim targetKey As String
    Dim lockPath As String
    Dim exists As Boolean
    Dim storedKey As String
    Dim storedSession As String
    Dim storedToken As String
    Dim lastWrite As Date

    status = LockUnknown
    If Not GetVectorBounds(targetFields, firstIndex, lastIndex) Or lastIndex - firstIndex + 1 <> 3 Then GoTo InvalidInput
    If VarType(targetFields(firstIndex)) <> vbString Or VarType(targetFields(firstIndex + 1)) <> vbString Or VarType(targetFields(firstIndex + 2)) <> vbString Then GoTo InvalidInput
    If Not BuildTargetKey(CStr(targetFields(firstIndex)), CStr(targetFields(firstIndex + 1)), CStr(targetFields(firstIndex + 2)), targetKey, lockPath, errorInfo) Then Exit Function

    If Not StoragePathExists(lockPath, exists, errorInfo) Then
        errorInfo = MakeAppError(ERR_NONE, "MdlLock.InspectTargetLock", "Lock status is unknown; checkout will verify it.")
        InspectTargetLock = True
        Exit Function
    End If
    If Not exists Then
        status = LockFree
        InspectTargetLock = True
        errorInfo = MakeAppError(ERR_NONE, "MdlLock.InspectTargetLock", "")
        Exit Function
    End If
    If Not ReadOwner(lockPath, storedKey, storedSession, storedToken, lastWrite, errorInfo) Then
        status = LockUnknown
        InspectTargetLock = True
        errorInfo = MakeAppError(ERR_NONE, "MdlLock.InspectTargetLock", "Lock metadata is unavailable; checkout will verify it.")
        Exit Function
    End If
    If StrComp(storedKey, targetKey, vbBinaryCompare) <> 0 Then
        status = LockUnknown
    ElseIf IsExpired(lastWrite, mColumnTtlSeconds) Then
        status = LockExpired
    Else
        status = LockHeld
    End If
    InspectTargetLock = True
    errorInfo = MakeAppError(ERR_NONE, "MdlLock.InspectTargetLock", "")
    Exit Function

InvalidInput:
    status = LockUnknown
    errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.InspectTargetLock", "A target must contain a UNC file path, sheet, and AMG header.")
End Function

Friend Function ReleaseLocks(ByRef leases As Variant, ByRef failures As Variant) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim rowIndex As Long
    Dim failureCount As Long
    Dim fields As Variant
    Dim failureRecords() As Variant
    Dim errorInfo As AppErrorInfo

    failures = Empty
    If IsEmpty(leases) Then
        ReleaseLocks = True
        Exit Function
    End If
    If Not GetMatrixBounds(leases, rowLow, rowHigh, columnLow, columnHigh) Or columnHigh - columnLow + 1 <> 4 Then Exit Function
    ReDim failureRecords(1 To rowHigh - rowLow + 1, 1 To 3)

    For rowIndex = rowLow To rowHigh
        fields = LeaseFieldsFromMatrix(leases, rowIndex, columnLow)
        If Not ReleaseLeaseFields(fields, mColumnTtlSeconds, errorInfo) Then
            failureCount = failureCount + 1
            failureRecords(failureCount, 1) = rowIndex
            failureRecords(failureCount, 2) = CStr(fields(1))
            failureRecords(failureCount, 3) = errorInfo.Message & " " & errorInfo.Context
        End If
    Next rowIndex

    If failureCount = 0 Then
        ReleaseLocks = True
    Else
        failures = failureRecords
    End If
End Function

Friend Function TryReclaimExpiredLock(ByRef targetFields As Variant, ByRef reclaimed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim firstIndex As Long
    Dim lastIndex As Long
    Dim targetKey As String
    Dim lockPath As String

    reclaimed = False
    If Not GetVectorBounds(targetFields, firstIndex, lastIndex) Or lastIndex - firstIndex + 1 <> 3 Then GoTo InvalidInput
    If VarType(targetFields(firstIndex)) <> vbString Or VarType(targetFields(firstIndex + 1)) <> vbString Or VarType(targetFields(firstIndex + 2)) <> vbString Then GoTo InvalidInput
    If Not BuildTargetKey(CStr(targetFields(firstIndex)), CStr(targetFields(firstIndex + 1)), CStr(targetFields(firstIndex + 2)), targetKey, lockPath, errorInfo) Then Exit Function
    TryReclaimExpiredLock = ReclaimLockDirectory(lockPath, targetKey, mColumnTtlSeconds, reclaimed, errorInfo)
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.TryReclaimExpiredLock", "A target must contain a UNC file path, sheet, and AMG header.")
End Function

Friend Function AcquireCommitGate(ByVal fileKey As String, ByVal sessionId As String, ByRef gateFields As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim normalizedFile As String
    Dim targetKey As String
    Dim lockPath As String
    Dim token As String
    Dim result(1 To 4) As Variant

    gateFields = Empty
    If Not IsSafeIdentifier(sessionId) Then GoTo InvalidInput
    If Not ValidateUncPath(fileKey, normalizedFile, errorInfo) Then Exit Function
    targetKey = BuildFileTargetKey(normalizedFile, "COMMIT")
    If Not BuildLockPath(normalizedFile, targetKey, lockPath, errorInfo) Then Exit Function
    If Not TryAcquireLockDirectory(lockPath, targetKey, sessionId, mGateTtlSeconds, token, errorInfo) Then Exit Function
    result(1) = lockPath
    result(2) = targetKey
    result(3) = sessionId
    result(4) = token
    gateFields = result
    AcquireCommitGate = True
    errorInfo = MakeAppError(ERR_NONE, "MdlLock.AcquireCommitGate", "")
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.AcquireCommitGate", "A UNC file key and safe session ID are required.")
End Function

Friend Function ValidateCommitGate(ByRef gateFields As Variant, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim firstIndex As Long
    Dim lastIndex As Long

    valid = False
    If Not GetVectorBounds(gateFields, firstIndex, lastIndex) Or lastIndex - firstIndex + 1 <> 4 Then GoTo InvalidInput
    If Not ValidateLeaseScalars(CStr(gateFields(firstIndex)), CStr(gateFields(firstIndex + 1)), CStr(gateFields(firstIndex + 2)), CStr(gateFields(firstIndex + 3)), mGateTtlSeconds, valid, errorInfo) Then Exit Function
    ValidateCommitGate = True
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.ValidateCommitGate", "A gate lease must contain path, target key, session ID, and fencing token.")
End Function

Friend Function ReleaseCommitGate(ByRef gateFields As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim firstIndex As Long
    Dim lastIndex As Long

    If Not GetVectorBounds(gateFields, firstIndex, lastIndex) Or lastIndex - firstIndex + 1 <> 4 Then
        errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.ReleaseCommitGate", "A gate lease must contain four fields.")
        Exit Function
    End If
    ReleaseCommitGate = ReleaseLeaseFields(gateFields, mGateTtlSeconds, errorInfo)
End Function

Private Function BuildTargetKey(ByVal fileKey As String, ByVal sheetName As String, ByVal amgHeader As String, ByRef targetKey As String, ByRef lockPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim normalizedFile As String
    Dim normalizedSheet As String
    Dim normalizedAmg As String

    targetKey = vbNullString
    lockPath = vbNullString
    If Not ValidateUncPath(fileKey, normalizedFile, errorInfo) Then Exit Function
    normalizedSheet = UCase$(Trim$(sheetName))
    normalizedAmg = UCase$(Trim$(amgHeader))
    If Len(normalizedSheet) = 0 Or Len(normalizedAmg) < 3 Or Left$(normalizedAmg, 3) <> "AMG" Then GoTo InvalidInput
    targetKey = BuildFileTargetKey(normalizedFile, normalizedSheet) & EncodePart(normalizedAmg)
    If Not BuildLockPath(normalizedFile, targetKey, lockPath, errorInfo) Then Exit Function
    BuildTargetKey = True
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.BuildTargetKey", "A non-empty sheet and AMG-prefixed header are required.", sheetName)
End Function

Private Function BuildFileTargetKey(ByVal fileKey As String, ByVal role As String) As String
    BuildFileTargetKey = EncodePart(UCase$(Replace$(fileKey, "/", "\"))) & EncodePart(UCase$(role))
End Function

Private Function EncodePart(ByVal value As String) As String
    EncodePart = CStr(Len(value)) & "#" & EncodeUtf16(value)
End Function

Private Function EncodeUtf16(ByVal value As String) As String
    Dim index As Long
    Dim code As Long
    Dim encoded As String

    For index = 1 To Len(value)
        code = AscW(Mid$(value, index, 1))
        If code < 0 Then code = code + 65536
        encoded = encoded & Right$("0000" & Hex$(code), 4)
    Next index
    EncodeUtf16 = encoded
End Function

Private Function DecodeUtf16(ByVal value As String) As String
    Dim index As Long
    Dim code As Long
    Dim decoded As String

    If Len(value) Mod 4 <> 0 Then Exit Function
    decoded = String$(Len(value) \ 4, vbNullChar)
    For index = 1 To Len(value) Step 4
        code = CLng("&H" & Mid$(value, index, 4))
        Mid$(decoded, (index + 3) \ 4, 1) = ChrW(code)
    Next index
    DecodeUtf16 = decoded
End Function

Private Function BuildLockPath(ByVal fileKey As String, ByVal targetKey As String, ByRef lockPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rootPath As String
    Dim fileRoot As String
    Dim separatorAt As Long
    Dim created As Boolean

    If Len(mTestRoot) > 0 Then
        rootPath = mTestRoot
        If Not EnsureLocalAppDirectory(rootPath, errorInfo) Then Exit Function
    Else
        separatorAt = InStrRev(fileKey, "\")
        If separatorAt <= 2 Then GoTo InvalidInput
        rootPath = Left$(fileKey, separatorAt - 1) & "\" & LOCK_ROOT_NAME
        If Not StorageCreateDirectory(rootPath, created, errorInfo) Then Exit Function
    End If

    fileRoot = JoinPath(rootPath, HashKey(UCase$(fileKey)))
    If Not StorageCreateDirectory(fileRoot, created, errorInfo) Then Exit Function
    lockPath = JoinPath(fileRoot, HashKey(targetKey))
    BuildLockPath = True
    Exit Function

InvalidInput:
    errorInfo = MakeAppError(ERR_LOCK_INVALID, "MdlLock.BuildLockPath", "The UNC file path must identify a file below a share root.", fileKey)
End Function

Private Function TryAcquireLockDirectory(ByVal lockPath As String, ByVal targetKey As String, ByVal sessionId As String, ByVal ttlSeconds As Long, ByRef fencingToken As String, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim created As Boolean
    Dim reclaimed As Boolean
    Dim ownerPath As String

    fencingToken = vbNullString
    If Not StorageCreateDirectory(lockPath, created, errorInfo) Then Exit Function
    If Not created Then
        If Not ReclaimLockDirectory(lockPath, targetKey, ttlSeconds, reclaimed, errorInfo) Then Exit Function
        If Not reclaimed Then
            errorInfo = MakeAppError(ERR_LOCK_BUSY, "MdlLock.TryAcquireLockDirectory", "The requested lock is held or cannot be safely reclaimed.", targetKey)
            Exit Function
        End If
        If Not StorageCreateDirectory(lockPath, created, errorInfo) Then Exit Function
        If Not created Then
            errorInfo = MakeAppError(ERR_LOCK_BUSY, "MdlLock.TryAcquireLockDirectory", "Another session acquired the lock during reclamation.", targetKey)
            Exit Function
        End If
    End If

    fencingToken = NewGuidString()
    If Len(fencingToken) = 0 Then
        StorageRemoveLockDirectory lockPath, errorInfo
        errorInfo = MakeAppError(ERR_LOCK_STORAGE, "MdlLock.TryAcquireLockDirectory", "A fencing token could not be generated.")
        Exit Function
    End If
    ownerPath = JoinPath(lockPath, OWNER_FILE_NAME)
    If Not StorageWriteText(ownerPath, SerializeOwner(sessionId, fencingToken, targetKey), errorInfo) Then
        StorageRemoveLockDirectory lockPath, errorInfo
        errorInfo = MakeAppError(ERR_LOCK_STORAGE, "MdlLock.TryAcquireLockDirectory", "The owner metadata could not be written.", targetKey)
        fencingToken = vbNullString
        Exit Function
    End If
    TryAcquireLockDirectory = True
    errorInfo = MakeAppError(ERR_NONE, "MdlLock.TryAcquireLockDirectory", "")
End Function

Private Function ReclaimLockDirectory(ByVal lockPath As String, ByVal expectedKey As String, ByVal ttlSeconds As Long, ByRef reclaimed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim storedKey As String
    Dim sessionId As String
    Dim fencingToken As String
    Dim lastWrite As Date
    Dim ownerPath As String
    Dim quarantinePath As String
    Dim ownerError As AppErrorInfo

    reclaimed = False
    If Not ReadOwner(lockPath, storedKey, sessionId, fencingToken, lastWrite, ownerError) Then
        errorInfo = MakeAppError(ERR_LOCK_OWNERSHIP, "MdlLock.ReclaimLockDirectory", "Lock metadata is missing or damaged; it remains occupied.", lockPath)
        ReclaimLockDirectory = True
        Exit Function
    End If
    If StrComp(storedKey, expectedKey, vbBinaryCompare) <> 0 Then
        errorInfo = MakeAppError(ERR_LOCK_OWNERSHIP, "MdlLock.ReclaimLockDirectory", "The lock directory hash collides with a different target.", lockPath)
        ReclaimLockDirectory = True
        Exit Function
    End If
    If Not IsExpired(lastWrite, ttlSeconds) Then
        errorInfo = MakeAppError(ERR_LOCK_BUSY, "MdlLock.ReclaimLockDirectory", "The lock heartbeat is still within its lease period.", lockPath)
        ReclaimLockDirectory = True
        Exit Function
    End If

    quarantinePath = JoinPath(ParentPath(lockPath), "reclaim-" & NewGuidString())
    If Not StorageRenameDirectory(lockPath, quarantinePath, errorInfo) Then Exit Function
    ownerPath = JoinPath(quarantinePath, OWNER_FILE_NAME)
    If Not StorageRemoveLockDirectory(quarantinePath, errorInfo) Then Exit Function
    reclaimed = True
    ReclaimLockDirectory = True
End Function

Private Function ReadOwner(ByVal lockPath As String, ByRef targetKey As String, ByRef sessionId As String, ByRef fencingToken As String, ByRef lastWrite As Date, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim contents As String
    If Not StorageReadText(JoinPath(lockPath, OWNER_FILE_NAME), contents, lastWrite, errorInfo) Then Exit Function
    If Not ParseOwner(contents, targetKey, sessionId, fencingToken) Then
        errorInfo = MakeAppError(ERR_LOCK_OWNERSHIP, "MdlLock.ReadOwner", "The owner metadata is malformed.", lockPath)
        Exit Function
    End If
    ReadOwner = True
End Function

Private Function SerializeOwner(ByVal sessionId As String, ByVal fencingToken As String, ByVal targetKey As String) As String
    SerializeOwner = "Version=" & OWNER_VERSION & vbCrLf & _
                     "Session=" & sessionId & vbCrLf & _
                     "Token=" & fencingToken & vbCrLf & _
                     "Target=" & EncodeUtf16(targetKey)
End Function

Private Function ParseOwner(ByVal contents As String, ByRef targetKey As String, ByRef sessionId As String, ByRef fencingToken As String) As Boolean
    Dim lines() As String
    Dim lineText As String
    Dim key As String
    Dim value As String
    Dim separatorAt As Long
    Dim version As String
    Dim encodedTarget As String
    Dim index As Long

    lines = Split(contents, vbLf)
    For index = LBound(lines) To UBound(lines)
        lineText = Replace$(lines(index), vbCr, vbNullString)
        separatorAt = InStr(1, lineText, "=", vbBinaryCompare)
        If separatorAt > 0 Then
            key = Left$(lineText, separatorAt - 1)
            value = Mid$(lineText, separatorAt + 1)
            Select Case key
                Case "Version": version = value
                Case "Session": sessionId = value
                Case "Token": fencingToken = value
                Case "Target": encodedTarget = value
            End Select
        End If
    Next index
    If version <> OWNER_VERSION Or Len(sessionId) = 0 Or Len(fencingToken) = 0 Or Len(encodedTarget) = 0 Then Exit Function
    targetKey = DecodeUtf16(encodedTarget)
    ParseOwner = (Len(targetKey) > 0)
End Function

Private Function ValidateLeaseScalars(ByVal lockPath As String, ByVal targetKey As String, ByVal sessionId As String, ByVal fencingToken As String, ByVal ttlSeconds As Long, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim storedKey As String
    Dim storedSession As String
    Dim storedToken As String
    Dim lastWrite As Date

    valid = False
    If Not ReadOwner(lockPath, storedKey, storedSession, storedToken, lastWrite, errorInfo) Then Exit Function
    If StrComp(storedKey, targetKey, vbBinaryCompare) <> 0 Or StrComp(storedSession, sessionId, vbBinaryCompare) <> 0 Or StrComp(storedToken, fencingToken, vbBinaryCompare) <> 0 Then
        errorInfo = MakeAppError(ERR_LOCK_OWNERSHIP, "MdlLock.ValidateLeaseScalars", "The owner or fencing token no longer matches.", lockPath)
        ValidateLeaseScalars = True
        Exit Function
    End If
    If IsExpired(lastWrite, ttlSeconds) Then
        errorInfo = MakeAppError(ERR_LOCK_EXPIRED, "MdlLock.ValidateLeaseScalars", "The lease has expired and cannot be renewed.", lockPath)
        ValidateLeaseScalars = True
        Exit Function
    End If
    valid = True
    ValidateLeaseScalars = True
    errorInfo = MakeAppError(ERR_NONE, "MdlLock.ValidateLeaseScalars", "")
End Function

Private Function RenewLeaseFields(ByRef leaseFields As Variant, ByVal ttlSeconds As Long, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim valid As Boolean
    Dim ownerPath As String

    If Not ValidateLeaseScalars(CStr(leaseFields(1)), CStr(leaseFields(2)), CStr(leaseFields(3)), CStr(leaseFields(4)), ttlSeconds, valid, errorInfo) Then Exit Function
    If Not valid Then Exit Function
    ownerPath = JoinPath(CStr(leaseFields(1)), OWNER_FILE_NAME)
    If Not StorageWriteText(ownerPath, SerializeOwner(CStr(leaseFields(3)), CStr(leaseFields(4)), CStr(leaseFields(2))), errorInfo) Then Exit Function
    RenewLeaseFields = True
End Function

Private Function ReleaseLeaseFields(ByRef leaseFields As Variant, ByVal ttlSeconds As Long, ByRef errorInfo As AppErrorInfo) As Boolean
    If Not RenewLeaseFields(leaseFields, ttlSeconds, errorInfo) Then Exit Function
    If Not StorageRemoveLockDirectory(CStr(leaseFields(1)), errorInfo) Then Exit Function
    ReleaseLeaseFields = True
    errorInfo = MakeAppError(ERR_NONE, "MdlLock.ReleaseLeaseFields", "")
End Function

Private Function IsExpired(ByVal lastWrite As Date, ByVal ttlSeconds As Long) As Boolean
    Dim ageSeconds As Long
    ageSeconds = DateDiff("s", lastWrite, DateAdd("s", mTestClockOffsetSeconds, Now))
    If ageSeconds < 0 Then ageSeconds = 0
    IsExpired = (ageSeconds > ttlSeconds)
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

Private Function LeaseFieldsFromMatrix(ByRef matrix As Variant, ByVal rowIndex As Long, ByVal columnLow As Long) As Variant
    Dim fields(1 To 4) As Variant
    fields(1) = matrix(rowIndex, columnLow)
    fields(2) = matrix(rowIndex, columnLow + 1)
    fields(3) = matrix(rowIndex, columnLow + 2)
    fields(4) = matrix(rowIndex, columnLow + 3)
    LeaseFieldsFromMatrix = fields
End Function

Private Function IsSafeIdentifier(ByVal value As String) As Boolean
    Dim index As Long
    Dim character As String

    If Len(value) < 1 Or Len(value) > 64 Then Exit Function
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If Not ((character >= "0" And character <= "9") Or (character >= "A" And character <= "Z") Or (character >= "a" And character <= "z") Or character = "-") Then Exit Function
    Next index
    IsSafeIdentifier = True
End Function

Private Sub SortTargetOrder(ByRef order() As Long, ByRef targetKeys() As String, ByVal targetCount As Long)
    Dim index As Long
    Dim insertionIndex As Long
    Dim value As Long

    For index = 2 To targetCount
        value = order(index)
        insertionIndex = index - 1
        Do While insertionIndex >= 1
            If StrComp(targetKeys(order(insertionIndex)), targetKeys(value), vbBinaryCompare) <= 0 Then Exit Do
            order(insertionIndex + 1) = order(insertionIndex)
            insertionIndex = insertionIndex - 1
        Loop
        order(insertionIndex + 1) = value
    Next index
End Sub

Private Function TryGetFileName(ByVal path As String, ByRef fileName As String) As Boolean
    Dim separatorAt As Long
    separatorAt = InStrRev(path, "\")
    If separatorAt <= 2 Or separatorAt = Len(path) Then Exit Function
    fileName = Mid$(path, separatorAt + 1)
    TryGetFileName = (Len(fileName) > 0)
End Function

Private Function HashKey(ByVal value As String) As String
    Dim hashValue As Double
    Dim scaledValue As Double
    Dim index As Long
    Dim code As Long
    Const MODULUS As Double = 4294967291#

    hashValue = 2166136261#
    For index = 1 To Len(value)
        code = AscW(Mid$(value, index, 1))
        If code < 0 Then code = code + 65536
        scaledValue = hashValue * 251# + CDbl(code) + 1#
        hashValue = scaledValue - Fix(scaledValue / MODULUS) * MODULUS
    Next index
    HashKey = Format$(hashValue, "0000000000")
End Function

Private Function JoinPath(ByVal basePath As String, ByVal child As String) As String
    If Len(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) = "\" Then
        JoinPath = basePath & child
    Else
        JoinPath = basePath & "\" & child
    End If
End Function

Private Function ParentPath(ByVal path As String) As String
    Dim separatorAt As Long
    separatorAt = InStrRev(path, "\")
    If separatorAt > 0 Then ParentPath = Left$(path, separatorAt - 1)
End Function
