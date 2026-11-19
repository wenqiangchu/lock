Attribute VB_Name = "MdlLock"
Option Explicit

Public Enum LockAvailability
    LockFree = 0
    LockHeld = 1
    LockExpired = 2
    LockUnknown = 3
End Enum

Private mCoordinator As CLockCoordinator

Public Function AcquireLocks(ByRef targets As Variant, ByVal sessionId As String, ByRef leases As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim coordinator As CLockCoordinator
    Set coordinator = GetCoordinator()
    AcquireLocks = coordinator.AcquireLocks(targets, sessionId, leases, errorInfo)
End Function

Public Function RenewLocks(ByRef leases As Variant, ByRef renewed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim coordinator As CLockCoordinator
    Set coordinator = GetCoordinator()
    RenewLocks = coordinator.RenewLocks(leases, renewed, errorInfo)
End Function

Public Function ValidateLease(ByRef leaseFields As Variant, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim coordinator As CLockCoordinator
    Set coordinator = GetCoordinator()
    ValidateLease = coordinator.ValidateLease(leaseFields, valid, errorInfo)
End Function

Public Function InspectLockStatus(ByRef targetFields As Variant, ByRef status As LockAvailability, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim coordinator As CLockCoordinator
    Set coordinator = GetCoordinator()
    InspectLockStatus = coordinator.InspectTargetLock(targetFields, status, errorInfo)
End Function

Public Function ReleaseLocks(ByRef leases As Variant, ByRef failures As Variant) As Boolean
    Dim coordinator As CLockCoordinator
    Set coordinator = GetCoordinator()
    ReleaseLocks = coordinator.ReleaseLocks(leases, failures)
End Function

Public Function TryReclaimExpiredLock(ByRef targetFields As Variant, ByRef reclaimed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim coordinator As CLockCoordinator
    Set coordinator = GetCoordinator()
    TryReclaimExpiredLock = coordinator.TryReclaimExpiredLock(targetFields, reclaimed, errorInfo)
End Function

Public Function AcquireCommitGate(ByVal fileKey As String, ByVal sessionId As String, ByRef gateFields As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim coordinator As CLockCoordinator
    Set coordinator = GetCoordinator()
    AcquireCommitGate = coordinator.AcquireCommitGate(fileKey, sessionId, gateFields, errorInfo)
End Function

Public Function ValidateCommitGate(ByRef gateFields As Variant, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim coordinator As CLockCoordinator
    Set coordinator = GetCoordinator()
    ValidateCommitGate = coordinator.ValidateCommitGate(gateFields, valid, errorInfo)
End Function

Public Function ReleaseCommitGate(ByRef gateFields As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim coordinator As CLockCoordinator
    Set coordinator = GetCoordinator()
    ReleaseCommitGate = coordinator.ReleaseCommitGate(gateFields, errorInfo)
End Function

Private Function GetCoordinator() As CLockCoordinator
    If mCoordinator Is Nothing Then Set mCoordinator = New CLockCoordinator
    Set GetCoordinator = mCoordinator
End Function
