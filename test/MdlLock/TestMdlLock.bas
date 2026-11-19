Attribute VB_Name = "TestMdlLock"
Option Explicit

Private Const TEST_FILE_PATH As String = "\\test-server\share\book.xlsx"

Public Sub RunMdlLockTests()
    RunTestCase "ContentionAndRelease"
    RunTestCase "DifferentColumnsAndBatchRollback"
    RunTestCase "ReadableLockLayoutAndPathAliases"
    RunTestCase "WorkbookRootsAreIndependent"
    RunTestCase "RenewAndOwnershipMismatch"
    RunTestCase "ExpiredAndDamagedLocks"
    RunTestCase "InspectStatus"
    RunTestCase "CommitGate"
    RunTestCase "RejectInvalidTargets"
End Sub

Public Sub RunMdlLockInspectStatusTest()
    RunTestCase "InspectStatus"
End Sub

Public Sub RunMdlLockInspectFreeTest()
    Dim root As String
    Dim coordinator As CLockCoordinator
    Dim targets As Variant
    Dim targetFields As Variant
    Dim status As LockAvailability
    Dim errorInfo As AppErrorInfo
    Dim inspectSucceeded As Boolean

    On Error GoTo Cleanup
    root = NewTestRoot()
    Set coordinator = NewTestCoordinator(root)
    targets = OneTarget("AMG1")
    targetFields = TargetAt(targets, 1)
    inspectSucceeded = coordinator.InspectTargetLock(targetFields, status, errorInfo)
    If Not inspectSucceeded Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectFree: inspection failed with code " & CStr(errorInfo.Code) & ", source " & errorInfo.Source & ", message " & errorInfo.Message & ", context " & errorInfo.Context
        GoTo Cleanup
    End If
    If status <> LockFree Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectFree: expected LockFree but received " & CStr(status) & "; error code " & CStr(errorInfo.Code) & ", source " & errorInfo.Source & ", message " & errorInfo.Message & ", context " & errorInfo.Context
        GoTo Cleanup
    End If
    GoTo Cleanup

Cleanup:
    If Len(root) > 0 Then RemoveTestRoot root
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.InspectFree: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub RunTestCase(ByVal caseName As String)
    On Error GoTo Failed
    Select Case caseName
        Case "ContentionAndRelease": TestContentionAndRelease
        Case "DifferentColumnsAndBatchRollback": TestDifferentColumnsAndBatchRollback
        Case "ReadableLockLayoutAndPathAliases": TestReadableLockLayoutAndPathAliases
        Case "WorkbookRootsAreIndependent": TestWorkbookRootsAreIndependent
        Case "RenewAndOwnershipMismatch": TestRenewAndOwnershipMismatch
        Case "ExpiredAndDamagedLocks": TestExpiredAndDamagedLocks
        Case "InspectStatus": TestInspectStatus
        Case "CommitGate": TestCommitGate
        Case "RejectInvalidTargets": TestRejectInvalidTargets
        Case Else: Err.Raise vbObjectError + 2301, "TestMdlLock", "Unknown test case: " & caseName
    End Select
    Exit Sub

Failed:
    TestSupport.RecordTestFailure "TestMdlLock." & caseName & ": " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestWorkbookRootsAreIndependent()
    Dim firstRoot As String
    Dim secondRoot As String
    Dim firstOwner As CLockCoordinator
    Dim secondOwner As CLockCoordinator
    Dim targets As Variant
    Dim firstLeases As Variant
    Dim secondLeases As Variant
    Dim failures As Variant
    Dim errorInfo As AppErrorInfo

    On Error GoTo Cleanup
    firstRoot = NewTestRoot()
    secondRoot = NewTestRoot()
    Set firstOwner = NewTestCoordinator(firstRoot)
    Set secondOwner = NewTestCoordinator(secondRoot)
    targets = OneTarget("AMG_A")

    AssertTrue firstOwner.AcquireLocks(targets, "first-root", firstLeases, errorInfo), "The first workbook root should acquire the target. " & errorInfo.Message
    AssertTrue secondOwner.AcquireLocks(targets, "second-root", secondLeases, errorInfo), "The same file identity in a different parent directory should use an independent lock root. " & errorInfo.Message
    AssertFalse StrComp(CStr(firstLeases(1, 1)), CStr(secondLeases(1, 1)), vbTextCompare) = 0, "Different parent directories must not share a lock path."
    AssertTrue firstOwner.ReleaseLocks(firstLeases, failures), "The first root lock should release."
    AssertTrue secondOwner.ReleaseLocks(secondLeases, failures), "The second root lock should release."
    GoTo Cleanup

Cleanup:
    RemoveTestRoot firstRoot
    RemoveTestRoot secondRoot
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.WorkbookRootsAreIndependent: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestReadableLockLayoutAndPathAliases()
    Dim root As String
    Dim owner As CLockCoordinator
    Dim contender As CLockCoordinator
    Dim uncTarget As Variant
    Dim mappedTarget As Variant
    Dim columnTargets As Variant
    Dim leases As Variant
    Dim contenderLeases As Variant
    Dim columnLeases As Variant
    Dim failures As Variant
    Dim errorInfo As AppErrorInfo
    Dim firstPath As String
    Dim secondPath As String

    On Error GoTo Cleanup
    root = NewTestRoot()
    Set owner = NewTestCoordinator(root)
    Set contender = NewTestCoordinator(root)
    uncTarget = OneTargetForPath(TEST_FILE_PATH, "Sheet~One", "AMG-A")
    mappedTarget = OneTargetForPath("X:\mapped\book.xlsx", "SHEET~ONE", "amg-a")

    AssertTrue owner.AcquireLocks(uncTarget, "unc-owner", leases, errorInfo), "The UNC-spelled target should acquire its lock. " & errorInfo.Message
    firstPath = CStr(leases(1, 1))
    AssertTrue InStr(1, firstPath, "\history\mddtlock\BOOK.XLSX\S-SHEET~007EONE__C-AMG-A_LOCK", vbTextCompare) > 0, "The lock path should identify its workbook, sheet, and escaped AMG target."
    AssertFalse contender.AcquireLocks(mappedTarget, "mapped-user", contenderLeases, errorInfo), "A drive-letter alias to the same workbook identity should encounter the existing lock."
    AssertTrue owner.ReleaseLocks(leases, failures), "The original alias should release the shared lock."

    columnTargets = TwoTargets("AMG_A", "AMG_B")
    AssertTrue owner.AcquireLocks(columnTargets, "two-columns", columnLeases, errorInfo), "Different columns should acquire independent lock directories. " & errorInfo.Message
    firstPath = CStr(columnLeases(1, 1))
    secondPath = CStr(columnLeases(2, 1))
    AssertFalse StrComp(firstPath, secondPath, vbTextCompare) = 0, "Different AMG columns must use different lock directories."
    AssertTrue InStr(1, firstPath, "__C-AMG_A_LOCK", vbTextCompare) > 0, "The first column name should be visible in its lock directory."
    AssertTrue InStr(1, secondPath, "__C-AMG_B_LOCK", vbTextCompare) > 0, "The second column name should be visible in its lock directory."
    AssertTrue owner.ReleaseLocks(columnLeases, failures), "Both column locks should release independently."
    GoTo Cleanup

Cleanup:
    RemoveTestRoot root
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.ReadableLockLayoutAndPathAliases: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestInspectStatus()
    Dim root As String
    Dim coordinator As CLockCoordinator
    Dim targets As Variant
    Dim targetFields As Variant
    Dim leases As Variant
    Dim damagedRoot As String
    Dim status As LockAvailability
    Dim errorInfo As AppErrorInfo
    Dim failures As Variant
    Dim created As Boolean
    Dim inspectSucceeded As Boolean

    On Error GoTo Cleanup
    root = NewTestRoot()
    Set coordinator = NewTestCoordinator(root)
    targets = OneTarget("AMG1")
    targetFields = TargetAt(targets, 1)
    inspectSucceeded = coordinator.InspectTargetLock(targetFields, status, errorInfo)
    If Not inspectSucceeded Or status <> LockFree Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: absent lock returned success=" & CStr(inspectSucceeded) & ", status=" & CStr(status) & ", error code=" & CStr(errorInfo.Code) & ", source=" & errorInfo.Source & ", message=" & errorInfo.Message & ", context=" & errorInfo.Context
        GoTo Cleanup
    End If

    If Not coordinator.AcquireLocks(targets, "InspectOwner", leases, errorInfo) Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: acquiring inspection lock failed; error code=" & CStr(errorInfo.Code) & ", source=" & errorInfo.Source & ", message=" & errorInfo.Message & ", context=" & errorInfo.Context
        GoTo Cleanup
    End If
    inspectSucceeded = coordinator.InspectTargetLock(targetFields, status, errorInfo)
    If Not inspectSucceeded Or status <> LockHeld Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: held lock returned success=" & CStr(inspectSucceeded) & ", status=" & CStr(status) & ", error code=" & CStr(errorInfo.Code) & ", source=" & errorInfo.Source & ", message=" & errorInfo.Message & ", context=" & errorInfo.Context
        GoTo Cleanup
    End If
    If Not coordinator.ReleaseLocks(leases, failures) Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: releasing inspection lock failed."
        GoTo Cleanup
    End If
    leases = Empty

    coordinator.SetTestTtl 1, 1
    If Not coordinator.AcquireLocks(targets, "ExpiredOwner", leases, errorInfo) Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: acquiring expiry lock failed; error code=" & CStr(errorInfo.Code) & ", source=" & errorInfo.Source & ", message=" & errorInfo.Message & ", context=" & errorInfo.Context
        GoTo Cleanup
    End If
    coordinator.SetTestClockOffset 2
    inspectSucceeded = coordinator.InspectTargetLock(targetFields, status, errorInfo)
    If Not inspectSucceeded Or status <> LockExpired Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: expired lock returned success=" & CStr(inspectSucceeded) & ", status=" & CStr(status) & ", error code=" & CStr(errorInfo.Code) & ", source=" & errorInfo.Source & ", message=" & errorInfo.Message & ", context=" & errorInfo.Context
        GoTo Cleanup
    End If
    coordinator.SetTestClockOffset 0
    If Not coordinator.ReleaseLocks(leases, failures) Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: releasing expiry lock failed."
        GoTo Cleanup
    End If
    leases = Empty

    targetFields = TargetAt(targets, 1)
    damagedRoot = root & "\damaged"
    If Not StorageCreateDirectory(damagedRoot, created, errorInfo) Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: creating damaged-owner fixture failed; " & errorInfo.Message
        GoTo Cleanup
    End If
    Set coordinator = NewTestCoordinator(damagedRoot)
    If Not coordinator.AcquireLocks(targets, "DamageOwner", leases, errorInfo) Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: acquiring damaged-owner fixture failed; " & errorInfo.Message
        GoTo Cleanup
    End If
    If Not StorageWriteText(CStr(leases(1, 1)) & "\owner.txt", "damaged", errorInfo) Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: corrupting owner metadata failed; " & errorInfo.Message
        GoTo Cleanup
    End If
    inspectSucceeded = coordinator.InspectTargetLock(targetFields, status, errorInfo)
    If Not inspectSucceeded Or status <> LockUnknown Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: malformed owner returned success=" & CStr(inspectSucceeded) & ", status=" & CStr(status) & ", error code=" & CStr(errorInfo.Code) & ", source=" & errorInfo.Source & ", message=" & errorInfo.Message & ", context=" & errorInfo.Context
        GoTo Cleanup
    End If
    If Not StorageRemoveLockDirectory(CStr(leases(1, 1)), errorInfo) Then
        TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: removing isolated damaged lock failed; " & errorInfo.Message
        GoTo Cleanup
    End If
    leases = Empty
    GoTo Cleanup

Cleanup:
    If Not IsEmpty(leases) Then
        If coordinator Is Nothing Then
            TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: a lease remained without its coordinator."
        ElseIf Not coordinator.ReleaseLocks(leases, failures) Then
            TestSupport.RecordTestFailure "TestMdlLock.InspectStatus: cleanup could not release remaining leases."
        End If
    End If
    If Len(root) > 0 Then RemoveTestRoot root
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.InspectStatus cleanup: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestContentionAndRelease()
    Dim root As String
    Dim first As CLockCoordinator
    Dim second As CLockCoordinator
    Dim targets As Variant
    Dim firstLeases As Variant
    Dim secondLeases As Variant
    Dim leaseFields As Variant
    Dim failures As Variant
    Dim errorInfo As AppErrorInfo
    Dim valid As Boolean

    On Error GoTo Cleanup
    root = NewTestRoot()
    Set first = NewTestCoordinator(root)
    Set second = NewTestCoordinator(root)
    targets = OneTarget("AMG_A")

    AssertTrue first.AcquireLocks(targets, "session-1", firstLeases, errorInfo), "The first owner should acquire a free lock. " & errorInfo.Message
    AssertFalse second.AcquireLocks(targets, "session-2", secondLeases, errorInfo), "The same target should be exclusive."
    AssertTrue IsEmpty(secondLeases), "A failed acquisition must return no partial lease set."
    leaseFields = LeaseAt(firstLeases, 1)
    AssertTrue first.ValidateLease(leaseFields, valid, errorInfo), "The first owner lease should be readable."
    AssertTrue valid, "The first owner lease should remain valid after contention."
    AssertTrue first.ReleaseLocks(firstLeases, failures), "The owning session should release its lease."
    AssertTrue IsEmpty(failures), "A successful release should have no failures."
    GoTo Cleanup

Cleanup:
    RemoveTestRoot root
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.ContentionAndRelease: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestDifferentColumnsAndBatchRollback()
    Dim root As String
    Dim holder As CLockCoordinator
    Dim batchOwner As CLockCoordinator
    Dim verifier As CLockCoordinator
    Dim heldTarget As Variant
    Dim batchTargets As Variant
    Dim verifyTarget As Variant
    Dim heldLeases As Variant
    Dim batchLeases As Variant
    Dim verifyLeases As Variant
    Dim failures As Variant
    Dim errorInfo As AppErrorInfo

    On Error GoTo Cleanup
    root = NewTestRoot()
    Set holder = NewTestCoordinator(root)
    Set batchOwner = NewTestCoordinator(root)
    Set verifier = NewTestCoordinator(root)
    heldTarget = OneTarget("AMG_B")
    batchTargets = TwoTargets("AMG_A", "AMG_B")

    AssertTrue holder.AcquireLocks(heldTarget, "holder", heldLeases, errorInfo), "A holder should lock the second target."
    AssertFalse batchOwner.AcquireLocks(batchTargets, "batch", batchLeases, errorInfo), "A busy member should cancel the complete batch."
    AssertTrue IsEmpty(batchLeases), "A rolled-back batch must not return partial leases."

    verifyTarget = OneTarget("AMG_A")
    AssertTrue verifier.AcquireLocks(verifyTarget, "verifier", verifyLeases, errorInfo), "A free column should remain independently lockable after rollback."
    AssertTrue verifier.ReleaseLocks(verifyLeases, failures), "The verification lease should be released."
    AssertTrue holder.ReleaseLocks(heldLeases, failures), "The held target should be released."
    GoTo Cleanup

Cleanup:
    RemoveTestRoot root
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.DifferentColumnsAndBatchRollback: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestRenewAndOwnershipMismatch()
    Dim root As String
    Dim owner As CLockCoordinator
    Dim targets As Variant
    Dim leases As Variant
    Dim leaseFields As Variant
    Dim forgedLease As Variant
    Dim forgedMatrix As Variant
    Dim failures As Variant
    Dim errorInfo As AppErrorInfo
    Dim renewed As Boolean
    Dim valid As Boolean

    On Error GoTo Cleanup
    root = NewTestRoot()
    Set owner = NewTestCoordinator(root)
    targets = OneTarget("AMG_A")
    AssertTrue owner.AcquireLocks(targets, "owner", leases, errorInfo), "A lease should be acquired."
    AssertTrue owner.RenewLocks(leases, renewed, errorInfo), "The lease batch should be processed."
    AssertTrue renewed, "A matching owner should renew."

    leaseFields = LeaseAt(leases, 1)
    forgedLease = leaseFields
    forgedLease(4) = "forged-token"
    AssertTrue owner.ValidateLease(forgedLease, valid, errorInfo), "Ownership mismatch is a validation result, not an input failure."
    AssertFalse valid, "A different fencing token must be rejected."
    forgedMatrix = leases
    forgedMatrix(1, 4) = "forged-token"
    AssertFalse owner.ReleaseLocks(forgedMatrix, failures), "A different fencing token must not release the lock."
    AssertFalse IsEmpty(failures), "Failed release must identify the remaining lock."
    AssertTrue owner.ValidateLease(leaseFields, valid, errorInfo), "The original owner should remain valid."
    AssertTrue valid, "The forged release must not remove the owner's lock."
    AssertTrue owner.ReleaseLocks(leases, failures), "The original owner should release the lock."
    GoTo Cleanup

Cleanup:
    RemoveTestRoot root
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.RenewAndOwnershipMismatch: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestExpiredAndDamagedLocks()
    Dim root As String
    Dim owner As CLockCoordinator
    Dim contender As CLockCoordinator
    Dim targets As Variant
    Dim targetFields As Variant
    Dim leases As Variant
    Dim newLeases As Variant
    Dim damagedLeases As Variant
    Dim leaseFields As Variant
    Dim failures As Variant
    Dim ownerPath As String
    Dim contents As String
    Dim lastWrite As Date
    Dim errorInfo As AppErrorInfo
    Dim reclaimed As Boolean

    On Error GoTo Cleanup
    root = NewTestRoot()
    Set owner = NewTestCoordinator(root)
    Set contender = NewTestCoordinator(root)
    owner.SetTestTtl 300, 600
    targets = OneTarget("AMG_A")
    AssertTrue owner.AcquireLocks(targets, "expired", leases, errorInfo), "The initial lock should be acquired."
    contender.SetTestClockOffset 301
    targetFields = TargetAt(targets, 1)
    AssertTrue contender.TryReclaimExpiredLock(targetFields, reclaimed, errorInfo), "An expired lock should be examinable."
    AssertTrue reclaimed, "A lock older than its TTL should be reclaimed."
    contender.SetTestClockOffset 0
    AssertTrue contender.AcquireLocks(targets, "replacement", newLeases, errorInfo), "A new owner should acquire a reclaimed lock."
    AssertTrue contender.ReleaseLocks(newLeases, failures), "The replacement lease should release cleanly."

    targets = OneTarget("AMG_DAMAGED")
    AssertTrue owner.AcquireLocks(targets, "damaged", damagedLeases, errorInfo), "A second test lock should be acquired."
    leaseFields = LeaseAt(damagedLeases, 1)
    ownerPath = CStr(leaseFields(1)) & "owner.txt"
    AssertTrue StorageWriteText(ownerPath, "damaged-owner-record", errorInfo), "The owner record should be replaceable for the corruption case."
    contender.SetTestClockOffset 1000
    targetFields = TargetAt(targets, 1)
    reclaimed = True
    AssertTrue contender.TryReclaimExpiredLock(targetFields, reclaimed, errorInfo), "A damaged owner should be handled as a lock state."
    AssertFalse reclaimed, "A damaged owner record must not be reclaimed."
    AssertTrue StorageReadText(ownerPath, contents, lastWrite, errorInfo), "The damaged lock should remain present for manual recovery."
    AssertTrue StrComp(contents, "damaged-owner-record", vbBinaryCompare) = 0, "The damaged metadata should remain untouched."
    GoTo Cleanup

Cleanup:
    RemoveTestRoot root
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.ExpiredAndDamagedLocks: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestCommitGate()
    Dim root As String
    Dim first As CLockCoordinator
    Dim second As CLockCoordinator
    Dim firstGate As Variant
    Dim secondGate As Variant
    Dim errorInfo As AppErrorInfo
    Dim valid As Boolean

    On Error GoTo Cleanup
    root = NewTestRoot()
    Set first = NewTestCoordinator(root)
    Set second = NewTestCoordinator(root)
    AssertTrue first.AcquireCommitGate(TEST_FILE_PATH, "first", firstGate, errorInfo), "The first commit gate should be acquired."
    AssertFalse second.AcquireCommitGate(TEST_FILE_PATH, "second", secondGate, errorInfo), "A second writer must wait at the file gate."
    AssertTrue first.ValidateCommitGate(firstGate, valid, errorInfo), "The first gate should validate."
    AssertTrue valid, "The first gate should remain valid."
    AssertTrue first.ReleaseCommitGate(firstGate, errorInfo), "The gate owner should release it."
    AssertTrue second.AcquireCommitGate(TEST_FILE_PATH, "second", secondGate, errorInfo), "The next writer should acquire the released gate."
    AssertTrue second.ReleaseCommitGate(secondGate, errorInfo), "The second gate owner should release it."
    GoTo Cleanup

Cleanup:
    RemoveTestRoot root
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.CommitGate: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Sub TestRejectInvalidTargets()
    Dim root As String
    Dim owner As CLockCoordinator
    Dim badTargets As Variant
    Dim targets As Variant
    Dim leases As Variant
    Dim errorInfo As AppErrorInfo

    On Error GoTo Cleanup
    root = NewTestRoot()
    Set owner = NewTestCoordinator(root)
    badTargets = OneTarget("not-an-amg")
    AssertFalse owner.AcquireLocks(badTargets, "safe-id", leases, errorInfo), "A non-AMG target must be rejected."
    AssertTrue IsEmpty(leases), "Invalid targets must not return leases."
    targets = OneTarget("AMG_A")
    AssertFalse owner.AcquireLocks(targets, "../unsafe", leases, errorInfo), "Unsafe session IDs must be rejected."
    targets = OneTarget("AMG" & String$(200, "A"))
    AssertFalse owner.AcquireLocks(targets, "long-path", leases, errorInfo), "Lock paths beyond the supported Windows path length must fail closed."
    AssertTrue IsEmpty(leases), "An unsupported long lock path must not return a lease."
    GoTo Cleanup

Cleanup:
    RemoveTestRoot root
    If Err.Number <> 0 Then TestSupport.RecordTestFailure "TestMdlLock.RejectInvalidTargets: " & Err.Description & " (" & CStr(Err.Number) & ")"
End Sub

Private Function NewTestCoordinator(ByVal root As String) As CLockCoordinator
    Dim coordinator As CLockCoordinator
    Set coordinator = New CLockCoordinator
    coordinator.SetTestRoot root
    Set NewTestCoordinator = coordinator
End Function

Private Function NewTestRoot() As String
    NewTestRoot = GetLocalAppDataFolder() & "\tests\locks-" & NewGuidString()
End Function

Private Sub RemoveTestRoot(ByVal root As String)
    Dim errorInfo As AppErrorInfo
    If Len(root) > 0 Then
        If Not RemoveLocalWorkspace(root, errorInfo) Then
            TestSupport.RecordTestFailure "TestMdlLock.RemoveTestRoot: " & errorInfo.Message & " " & errorInfo.Context
        End If
    End If
End Sub

Private Function OneTarget(ByVal amgHeader As String) As Variant
    OneTarget = OneTargetForPath(TEST_FILE_PATH, "Sheet1", amgHeader)
End Function

Private Function OneTargetForPath(ByVal filePath As String, ByVal sheetName As String, ByVal amgHeader As String) As Variant
    Dim targets(1 To 1, 1 To 3) As Variant
    targets(1, 1) = filePath
    targets(1, 2) = sheetName
    targets(1, 3) = amgHeader
    OneTargetForPath = targets
End Function

Private Function TwoTargets(ByVal firstHeader As String, ByVal secondHeader As String) As Variant
    Dim targets(1 To 2, 1 To 3) As Variant
    targets(1, 1) = TEST_FILE_PATH
    targets(1, 2) = "Sheet1"
    targets(1, 3) = firstHeader
    targets(2, 1) = TEST_FILE_PATH
    targets(2, 2) = "Sheet1"
    targets(2, 3) = secondHeader
    TwoTargets = targets
End Function

Private Function LeaseAt(ByRef leases As Variant, ByVal rowIndex As Long) As Variant
    Dim fields(1 To 4) As Variant
    Dim columnIndex As Long
    For columnIndex = 1 To 4
        fields(columnIndex) = leases(rowIndex, columnIndex)
    Next columnIndex
    LeaseAt = fields
End Function

Private Function TargetAt(ByRef targets As Variant, ByVal rowIndex As Long) As Variant
    Dim fields(1 To 3) As Variant
    Dim columnIndex As Long
    For columnIndex = 1 To 3
        fields(columnIndex) = targets(rowIndex, columnIndex)
    Next columnIndex
    TargetAt = fields
End Function

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2302, "TestMdlLock", message
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub
