Attribute VB_Name = "MdlSession"
Option Explicit
Option Private Module

Public Enum SessionState
    SessionIdle = 0
    SessionTargetSelected = 1
    SessionCheckingOut = 2
    SessionCheckedOut = 3
    SessionCommitting = 4
    SessionConflictReview = 5
    SessionAborting = 6
End Enum

Public Enum SubmitOutcome
    SubmitKeepSession = 0
    SubmitCompleted = 1
    SubmitNeedsConflictReview = 2
    SubmitCompletedWithCleanupWarning = 3
End Enum

Public Type SessionView
    State As SessionState
    SessionId As String
    TargetCount As Long
    ConflictCount As Long
    CanSubmit As Boolean
    CanAbandon As Boolean
    SaveResultUncertain As Boolean
    StatusText As String
End Type

Public Const ERR_SESSION_INVALID_STATE As Long = 1701
Public Const ERR_SESSION_INVALID_TARGETS As Long = 1702
Public Const ERR_SESSION_LOCK_LOST As Long = 1703
Public Const ERR_SESSION_UNDECLARED_CHANGE As Long = 1704
Public Const ERR_SESSION_SAVE_UNCERTAIN As Long = 1705

Private Const HEARTBEAT_INTERVAL_SECONDS As Long = 60

Private mState As SessionState
Private mSessionId As String
Private mMainFilePath As String
Private mWorkspacePath As String
Private mLocalFilePath As String
Private mLocalHandle As String
Private mLeases As Variant
Private mTargets As Variant
Private mSheetNames As Variant
Private mSheetTargets As Variant
Private mTargetCount As Long
Private mGroupCount As Long
Private mBaseSnapshots() As SnapshotData
Private mPendingConflicts As Variant
Private mCanSubmit As Boolean
Private mSaveResultUncertain As Boolean
Private mStatusText As String

Public Function BeginCheckout(ByRef targets As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim config As AppConfig
    Dim normalizedFile As String
    Dim normalizedTargets As Variant
    Dim groupNames As Variant
    Dim groupTargets As Variant
    Dim groupCount As Long
    Dim sharedHandle As String
    Dim validationSnapshot As SnapshotData
    Dim errorCleanup As AppErrorInfo
    Dim index As Long

    If mState <> SessionIdle Then
        errorInfo = MakeAppError(ERR_SESSION_INVALID_STATE, "MdlSession.BeginCheckout", "A session is already active.")
        Exit Function
    End If
    If Not TransitionTo(SessionTargetSelected, errorInfo) Then Exit Function
    mSessionId = NewGuidString()
    If Len(mSessionId) = 0 Then GoTo Failed
    If Not ReadConfig(GetDefaultConfigPath(), config, errorInfo) Then GoTo Failed
    If Not ValidateUncPath(config.MainFilePath, normalizedFile, errorInfo) Then GoTo Failed
    If Not NormalizeTargets(targets, normalizedFile, normalizedTargets, errorInfo) Then GoTo Failed
    If Not BuildTargetGroups(normalizedTargets, groupNames, groupTargets, groupCount, errorInfo) Then GoTo Failed
    mTargets = normalizedTargets
    mSheetNames = groupNames
    mSheetTargets = groupTargets
    mTargetCount = UBound(normalizedTargets, 1)
    mGroupCount = groupCount
    mMainFilePath = normalizedFile

    If Not TransitionTo(SessionCheckingOut, errorInfo) Then GoTo Failed
    If Not OpenWorkbookHandle(mMainFilePath, WorkbookSharedRemote, sharedHandle, errorInfo) Then GoTo Failed
    For index = 1 To mGroupCount
        If Not ReadDeclaredSnapshot(sharedHandle, mSheetTargets(index), validationSnapshot, errorInfo) Then GoTo Rollback
    Next index
    CloseWorkbookHandle sharedHandle, False
    sharedHandle = vbNullString

    If Not AcquireLocks(mTargets, mSessionId, mLeases, errorInfo) Then GoTo Failed
    If Not CreateLocalWorkspace(mSessionId, mWorkspacePath, errorInfo) Then GoTo Rollback
    If Not CopySharedFileToWorkspace(mMainFilePath, mWorkspacePath, mLocalFilePath, errorInfo) Then GoTo Rollback
    If Not OpenWorkbookHandle(mLocalFilePath, WorkbookLocalCopy, mLocalHandle, errorInfo) Then GoTo Rollback

    ReDim mBaseSnapshots(1 To mGroupCount)
    For index = 1 To mGroupCount
        If Not ReadDeclaredSnapshot(mLocalHandle, mSheetTargets(index), mBaseSnapshots(index), errorInfo) Then GoTo Rollback
    Next index

    mCanSubmit = True
    mSaveResultUncertain = False
    mStatusText = "Checked out"
    If Not TransitionTo(SessionCheckedOut, errorInfo) Then GoTo Rollback
    ScheduleHeartbeat
    BeginCheckout = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.BeginCheckout", "")
    Exit Function

Rollback:
    If Len(sharedHandle) > 0 Then CloseWorkbookHandle sharedHandle, False
    RollbackCheckout errorCleanup
    If errorCleanup.Code <> ERR_NONE Then LogWrite LogError, "MdlSession.BeginCheckout", errorCleanup.Message & " " & errorCleanup.Context
    Exit Function

Failed:
    If Len(sharedHandle) > 0 Then CloseWorkbookHandle sharedHandle, False
    If errorInfo.Code = ERR_NONE Then errorInfo = MakeAppError(ERR_SESSION_INVALID_TARGETS, "MdlSession.BeginCheckout", "The checkout could not be started.")
    RollbackCheckout errorCleanup
End Function

Public Function GetAvailableTargets(ByRef targets As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim config As AppConfig
    Dim normalizedFile As String
    Dim handleId As String
    Dim sheetNames As Variant
    Dim sheetItem As Variant
    Dim schema As SheetSchema
    Dim schemaError As AppErrorInfo
    Dim targetFields As Variant
    Dim lockStatus As LockAvailability
    Dim lockError As AppErrorInfo
    Dim entries As Collection
    Dim entry As Variant
    Dim amgIndexes As Variant
    Dim index As Long
    Dim outputIndex As Long
    Dim result() As Variant

    targets = Empty
    If Not ReadConfig(GetDefaultConfigPath(), config, errorInfo) Then Exit Function
    If Not ValidateUncPath(config.MainFilePath, normalizedFile, errorInfo) Then Exit Function
    If Not OpenWorkbookHandle(normalizedFile, WorkbookSharedRemote, handleId, errorInfo) Then Exit Function
    On Error GoTo Failed
    If Not ListWorksheetNames(handleId, sheetNames, errorInfo) Then GoTo Cleanup
    Set entries = New Collection

    If IsArray(sheetNames) Then
        For Each sheetItem In sheetNames
            If Not ReadSheetSchema(handleId, CStr(sheetItem), schema, schemaError) Then
                If schemaError.Code = ERR_SCHEMA_NO_ANCHOR Then
                    LogWrite LogWarning, "MdlSession.GetAvailableTargets", "A sheet without an A-E product anchor was skipped."
                Else
                    errorInfo = schemaError
                    GoTo Cleanup
                End If
            ElseIf Not IsEmpty(schema.AmgColumnIndexes) Then
                amgIndexes = schema.AmgColumnIndexes
                For index = LBound(amgIndexes) To UBound(amgIndexes)
                    entry = Array(normalizedFile, CStr(sheetItem), CStr(schema.HeaderNames(CLng(amgIndexes(index)))))
                    If Not InspectLockStatus(entry, lockStatus, lockError) Then lockStatus = LockUnknown
                    entry = Array(normalizedFile, CStr(sheetItem), CStr(schema.HeaderNames(CLng(amgIndexes(index)))), LockStatusText(lockStatus))
                    entries.Add entry
                Next index
            End If
        Next sheetItem
    End If

    If entries.Count > 0 Then
        ReDim result(1 To entries.Count, 1 To 4)
        For outputIndex = 1 To entries.Count
            entry = entries.Item(outputIndex)
            result(outputIndex, 1) = entry(0)
            result(outputIndex, 2) = entry(1)
            result(outputIndex, 3) = entry(2)
            result(outputIndex, 4) = entry(3)
        Next outputIndex
        targets = result
    End If
    GetAvailableTargets = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.GetAvailableTargets", "")
    GoTo Cleanup

Failed:
    targets = Empty
    errorInfo = MakeAppError(ERR_SESSION_INVALID_TARGETS, "MdlSession.GetAvailableTargets", "可签出目标无法安全读取。", CStr(Err.Number) & ": " & Err.Description)

Cleanup:
    If Len(handleId) > 0 Then CloseWorkbookHandle handleId, False
End Function

Public Function SubmitChanges(ByRef conflicts As Variant, ByRef errorInfo As AppErrorInfo) As SubmitOutcome
    Dim noChoices As Variant
    noChoices = Empty
    SubmitChanges = ExecuteCommit(noChoices, False, conflicts, errorInfo)
End Function

Public Function ResumeAfterConflictReview(ByRef choices As Variant, ByRef conflicts As Variant, ByRef errorInfo As AppErrorInfo) As SubmitOutcome
    ResumeAfterConflictReview = ExecuteCommit(choices, True, conflicts, errorInfo)
End Function

Private Function ExecuteCommit(ByRef choices As Variant, ByVal isConflictReview As Boolean, ByRef conflicts As Variant, ByRef errorInfo As AppErrorInfo) As SubmitOutcome
    Dim remoteHandle As String
    Dim gateFields As Variant
    Dim gateHeld As Boolean
    Dim saveAttempted As Boolean
    Dim commitSucceeded As Boolean
    Dim cleanupSucceeded As Boolean
    Dim gateReleaseSucceeded As Boolean
    Dim gateError As AppErrorInfo
    Dim transitionError As AppErrorInfo
    Dim cleanupError As AppErrorInfo
    Dim mergeResults() As MergeResult
    Dim currentConflicts As Variant
    Dim hasLocalChanges As Boolean
    Dim nextState As SessionState
    Dim index As Long

    conflicts = Empty
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.ExecuteCommit", "")
    ExecuteCommit = SubmitKeepSession
    On Error GoTo Failed
    If isConflictReview Then
        If mState <> SessionConflictReview Then GoTo InvalidState
    ElseIf mState <> SessionCheckedOut Then
        GoTo InvalidState
    End If
    If Not mCanSubmit Or mSaveResultUncertain Then
        errorInfo = MakeAppError(ERR_SESSION_LOCK_LOST, "MdlSession.ExecuteCommit", "提交已禁用；请保留本地副本并检查会话状态。")
        Exit Function
    End If
    If Not TransitionTo(SessionCommitting, errorInfo) Then Exit Function
    nextState = SessionCheckedOut

    If Not ValidateSessionLeases(errorInfo) Then GoTo Cleanup
    If Not AcquireCommitGate(mMainFilePath, mSessionId, gateFields, errorInfo) Then GoTo Cleanup
    gateHeld = True
    If Not OpenWorkbookHandle(mMainFilePath, WorkbookSharedRemote, remoteHandle, errorInfo) Then GoTo Cleanup
    If Not BuildCurrentMerges(remoteHandle, mergeResults, currentConflicts, hasLocalChanges, errorInfo) Then GoTo Cleanup

    conflicts = currentConflicts
    If IsArray(currentConflicts) Then
        If isConflictReview Then
            If Not ConflictViewsEqual(mPendingConflicts, currentConflicts) Then
                mPendingConflicts = currentConflicts
                nextState = SessionConflictReview
                ExecuteCommit = SubmitNeedsConflictReview
                mStatusText = "共享数据在裁决期间发生变化，请重新检查冲突"
                GoTo Cleanup
            End If
            If Not ApplySessionConflictChoices(mergeResults, choices, errorInfo) Then GoTo Cleanup
        Else
            mPendingConflicts = currentConflicts
            nextState = SessionConflictReview
            ExecuteCommit = SubmitNeedsConflictReview
            mStatusText = "需要逐格裁决冲突"
            GoTo Cleanup
        End If
    Else
        mPendingConflicts = Empty
    End If

    If Not hasLocalChanges Then
        commitSucceeded = True
        ExecuteCommit = SubmitCompleted
        GoTo Cleanup
    End If

    If Not ValidateCommitAuthority(gateFields, errorInfo) Then GoTo Cleanup
    For index = 1 To mGroupCount
        If Not WriteDeclaredColumns(remoteHandle, mSheetTargets(index), mergeResults(index).MergedValues, errorInfo) Then GoTo Cleanup
    Next index
    If Not ValidateCommitAuthority(gateFields, errorInfo) Then GoTo Cleanup

    saveAttempted = True
    If Not SaveWorkbook(remoteHandle, errorInfo) Then
        mSaveResultUncertain = True
        mCanSubmit = False
        mStatusText = "共享保存结果不确定；本地副本和列锁已保留，禁止再次提交"
        errorInfo = MakeAppError(ERR_SESSION_SAVE_UNCERTAIN, "MdlSession.ExecuteCommit", "共享保存结果不确定。请核对共享文件后再决定如何处理本地副本。", errorInfo.Message & " " & errorInfo.Context)
        GoTo Cleanup
    End If

    commitSucceeded = True
    ExecuteCommit = SubmitCompleted
    GoTo Cleanup

InvalidState:
    errorInfo = MakeAppError(ERR_SESSION_INVALID_STATE, "MdlSession.ExecuteCommit", "当前会话状态不允许提交。")
    Exit Function

Cleanup:
    If Len(remoteHandle) > 0 Then CloseWorkbookHandle remoteHandle, False
    If gateHeld Then
        gateReleaseSucceeded = ReleaseCommitGate(gateFields, gateError)
        gateHeld = False
        If Not gateReleaseSucceeded And errorInfo.Code = ERR_NONE Then errorInfo = gateError
    Else
        gateReleaseSucceeded = True
    End If

    If commitSucceeded Then
        cleanupSucceeded = ReleaseSessionResources(cleanupError)
        If Not TransitionTo(SessionIdle, transitionError) Then
            cleanupSucceeded = False
            If cleanupError.Code = ERR_NONE Then cleanupError = transitionError
        End If
        ResetSession
        If Not cleanupSucceeded Or Not gateReleaseSucceeded Then
            If cleanupError.Code = ERR_NONE And Not gateReleaseSucceeded Then cleanupError = gateError
            If cleanupError.Code = ERR_NONE Then cleanupError = MakeAppError(ERR_IO, "MdlSession.ExecuteCommit", "提交已完成，但本地会话清理未全部完成。")
            errorInfo = cleanupError
            ExecuteCommit = SubmitCompletedWithCleanupWarning
        Else
            errorInfo = MakeAppError(ERR_NONE, "MdlSession.ExecuteCommit", "")
        End If
    ElseIf mState = SessionCommitting Then
        If errorInfo.Code = ERR_NONE Then errorInfo = MakeAppError(ERR_SESSION_INVALID_STATE, "MdlSession.ExecuteCommit", "提交未完成，会话和本地副本已保留。")
        If Not TransitionTo(nextState, cleanupError) Then LogWrite LogError, "MdlSession.ExecuteCommit", cleanupError.Message
        If nextState = SessionConflictReview Then mStatusText = "需要逐格裁决冲突"
    End If
    Exit Function

Failed:
    If saveAttempted Then
        mSaveResultUncertain = True
        mCanSubmit = False
        mStatusText = "共享保存结果不确定；本地副本和列锁已保留，禁止再次提交"
        errorInfo = MakeAppError(ERR_SESSION_SAVE_UNCERTAIN, "MdlSession.ExecuteCommit", "共享保存过程中发生错误，结果不确定。请核对共享文件。", CStr(Err.Number) & ": " & Err.Description)
    Else
        errorInfo = MakeAppError(ERR_SESSION_INVALID_STATE, "MdlSession.ExecuteCommit", "提交过程发生错误；本地会话已保留。", CStr(Err.Number) & ": " & Err.Description)
    End If
    Resume Cleanup
End Function

Public Function CancelConflictReview() As Boolean
    Dim errorInfo As AppErrorInfo

    If mState <> SessionConflictReview Then Exit Function
    If Not TransitionTo(SessionCheckedOut, errorInfo) Then Exit Function
    mPendingConflicts = Empty
    mStatusText = "Conflict review canceled; local session retained"
    CancelConflictReview = True
End Function

Public Function AbandonSession(ByVal userConfirmed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    If Not userConfirmed Then
        errorInfo = MakeAppError(ERR_SESSION_INVALID_STATE, "MdlSession.AbandonSession", "Abandoning the session requires explicit user confirmation.")
        Exit Function
    End If
    If mState <> SessionCheckedOut And mState <> SessionConflictReview Then
        errorInfo = MakeAppError(ERR_SESSION_INVALID_STATE, "MdlSession.AbandonSession", "The current session cannot be abandoned in this state.")
        Exit Function
    End If
    If Not TransitionTo(SessionAborting, errorInfo) Then Exit Function
    AbandonSession = ReleaseSessionResources(errorInfo)
    ResetSession
End Function

Public Function GetSessionView() As SessionView
    Dim view As SessionView
    view.State = mState
    view.SessionId = mSessionId
    view.TargetCount = mTargetCount
    If IsArray(mPendingConflicts) Then view.ConflictCount = UBound(mPendingConflicts, 1) - LBound(mPendingConflicts, 1) + 1
    view.CanSubmit = ((mState = SessionCheckedOut) And mCanSubmit And Not mSaveResultUncertain)
    view.CanAbandon = (mState = SessionCheckedOut Or mState = SessionConflictReview)
    view.SaveResultUncertain = mSaveResultUncertain
    view.StatusText = mStatusText
    GetSessionView = view
End Function

Public Sub SessionHeartbeat()
    Dim renewed As Boolean
    Dim errorInfo As AppErrorInfo

    If mState <> SessionCheckedOut And mState <> SessionConflictReview Then Exit Sub
    If Not RenewLocks(mLeases, renewed, errorInfo) Or Not renewed Then
        mCanSubmit = False
        mStatusText = "Column lease renewal failed; submission is disabled"
        LogWrite LogError, "MdlSession.SessionHeartbeat", errorInfo.Message & " " & errorInfo.Context
        Exit Sub
    End If
    ScheduleHeartbeat
End Sub

Private Function TransitionTo(ByVal nextState As SessionState, ByRef errorInfo As AppErrorInfo) As Boolean
    If Not IsSessionTransitionAllowed(mState, nextState) Then
        errorInfo = MakeAppError(ERR_SESSION_INVALID_STATE, "MdlSession.TransitionTo", "请求的会话状态转换不合法。", CStr(mState) & " -> " & CStr(nextState))
        Exit Function
    End If
    mState = nextState
    TransitionTo = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.TransitionTo", "")
End Function

Public Function IsSessionTransitionAllowed(ByVal currentState As SessionState, ByVal nextState As SessionState) As Boolean
    Select Case currentState
        Case SessionIdle
            IsSessionTransitionAllowed = (nextState = SessionTargetSelected)
        Case SessionTargetSelected
            IsSessionTransitionAllowed = (nextState = SessionCheckingOut Or nextState = SessionIdle)
        Case SessionCheckingOut
            IsSessionTransitionAllowed = (nextState = SessionCheckedOut Or nextState = SessionIdle)
        Case SessionCheckedOut
            IsSessionTransitionAllowed = (nextState = SessionCommitting Or nextState = SessionAborting)
        Case SessionCommitting
            IsSessionTransitionAllowed = (nextState = SessionConflictReview Or nextState = SessionCheckedOut Or nextState = SessionIdle)
        Case SessionConflictReview
            IsSessionTransitionAllowed = (nextState = SessionCommitting Or nextState = SessionCheckedOut Or nextState = SessionAborting)
        Case SessionAborting
            IsSessionTransitionAllowed = (nextState = SessionIdle)
    End Select
End Function

Private Function BuildCurrentMerges(ByVal remoteHandle As String, ByRef merges() As MergeResult, ByRef conflicts As Variant, ByRef hasLocalChanges As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim index As Long
    Dim baseSnapshot As SnapshotData
    Dim localSnapshot As SnapshotData
    Dim remoteSnapshot As SnapshotData
    Dim expected As SessionSchema
    Dim valid As Boolean
    Dim localChanges As Variant
    Dim conflictCount As Long
    Dim outputIndex As Long
    Dim sourceIndex As Long
    Dim conflictRow As Long
    Dim conflictColumn As Long
    Dim conflictValues() As Variant

    ReDim merges(1 To mGroupCount)
    hasLocalChanges = False
    conflictCount = 0
    For index = 1 To mGroupCount
        baseSnapshot = mBaseSnapshots(index)
        CopySnapshotSchema baseSnapshot, expected
        If Not ValidateStructure(mLocalHandle, expected, valid, errorInfo) Then Exit Function
        If Not valid Then Exit Function
        If Not ValidateStructure(remoteHandle, expected, valid, errorInfo) Then Exit Function
        If Not valid Then Exit Function
        If Not ReadDeclaredSnapshot(mLocalHandle, mSheetTargets(index), localSnapshot, errorInfo) Then Exit Function
        If Not ReadDeclaredSnapshot(remoteHandle, mSheetTargets(index), remoteSnapshot, errorInfo) Then Exit Function
        If Not CompareSnapshots(baseSnapshot.Values, localSnapshot.Values, baseSnapshot.DeclaredColumnIndexes, localChanges, errorInfo) Then Exit Function
        If Not IsEmpty(localChanges) Then hasLocalChanges = True
        If Not MergeThreeWay(baseSnapshot.Values, localSnapshot.Values, remoteSnapshot.Values, baseSnapshot.DeclaredColumnIndexes, merges(index), errorInfo) Then Exit Function
        conflictCount = conflictCount + merges(index).ConflictCount
    Next index

    conflicts = Empty
    If conflictCount > 0 Then
        ReDim conflictValues(1 To conflictCount, 1 To 6)
        outputIndex = 0
        For index = 1 To mGroupCount
            If merges(index).ConflictCount > 0 Then
                For sourceIndex = LBound(merges(index).Conflicts, 1) To UBound(merges(index).Conflicts, 1)
                    outputIndex = outputIndex + 1
                    conflictRow = CLng(merges(index).Conflicts(sourceIndex, 1))
                    conflictColumn = CLng(merges(index).Conflicts(sourceIndex, 2))
                    conflictValues(outputIndex, 1) = CStr(mSheetNames(index))
                    conflictValues(outputIndex, 2) = conflictRow
                    conflictValues(outputIndex, 3) = conflictColumn
                    conflictValues(outputIndex, 4) = merges(index).Conflicts(sourceIndex, 3)
                    conflictValues(outputIndex, 5) = merges(index).Conflicts(sourceIndex, 4)
                    conflictValues(outputIndex, 6) = merges(index).Conflicts(sourceIndex, 5)
                Next sourceIndex
            End If
        Next index
        conflicts = conflictValues
    End If
    BuildCurrentMerges = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.BuildCurrentMerges", "")
End Function

Private Function ValidateSessionLeases(ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim rowIndex As Long
    Dim fields As Variant
    Dim valid As Boolean

    If IsEmpty(mLeases) Then GoTo LeaseLost
    If Not GetMatrixBounds(mLeases, rowLow, rowHigh, columnLow, columnHigh) Then GoTo LeaseLost
    If columnHigh - columnLow + 1 <> 4 Or rowHigh - rowLow + 1 <> mTargetCount Then GoTo LeaseLost
    For rowIndex = rowLow To rowHigh
        fields = Array(mLeases(rowIndex, columnLow), mLeases(rowIndex, columnLow + 1), mLeases(rowIndex, columnLow + 2), mLeases(rowIndex, columnLow + 3))
        If Not ValidateLease(fields, valid, errorInfo) Then
            mCanSubmit = False
            Exit Function
        End If
        If Not valid Then GoTo LeaseLost
    Next rowIndex
    ValidateSessionLeases = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.ValidateSessionLeases", "")
    Exit Function

LeaseLost:
    mCanSubmit = False
    errorInfo = MakeAppError(ERR_SESSION_LOCK_LOST, "MdlSession.ValidateSessionLeases", "列锁已失效或租约记录不完整；提交已停止。")
End Function

Private Function ValidateCommitAuthority(ByRef gateFields As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim valid As Boolean

    If Not ValidateSessionLeases(errorInfo) Then Exit Function
    If Not ValidateCommitGate(gateFields, valid, errorInfo) Then Exit Function
    If Not valid Then
        mCanSubmit = False
        errorInfo = MakeAppError(ERR_SESSION_LOCK_LOST, "MdlSession.ValidateCommitAuthority", "共享提交门闩已失效；提交已停止。")
        Exit Function
    End If
    ValidateCommitAuthority = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.ValidateCommitAuthority", "")
End Function

Private Function ConflictViewsEqual(ByRef leftConflicts As Variant, ByRef rightConflicts As Variant) As Boolean
    Dim leftRowLow As Long
    Dim leftRowHigh As Long
    Dim leftColumnLow As Long
    Dim leftColumnHigh As Long
    Dim rightRowLow As Long
    Dim rightRowHigh As Long
    Dim rightColumnLow As Long
    Dim rightColumnHigh As Long
    Dim leftIndex As Long
    Dim rightIndex As Long
    Dim valueIndex As Long
    Dim equal As Boolean
    Dim comparisonError As AppErrorInfo

    If IsEmpty(leftConflicts) Or IsEmpty(rightConflicts) Then
        ConflictViewsEqual = (IsEmpty(leftConflicts) And IsEmpty(rightConflicts))
        Exit Function
    End If
    If Not GetMatrixBounds(leftConflicts, leftRowLow, leftRowHigh, leftColumnLow, leftColumnHigh) Then Exit Function
    If Not GetMatrixBounds(rightConflicts, rightRowLow, rightRowHigh, rightColumnLow, rightColumnHigh) Then Exit Function
    If leftRowHigh - leftRowLow <> rightRowHigh - rightRowLow Then Exit Function
    If leftColumnHigh - leftColumnLow <> 5 Or rightColumnHigh - rightColumnLow <> 5 Then Exit Function
    For leftIndex = leftRowLow To leftRowHigh
        rightIndex = rightRowLow + leftIndex - leftRowLow
        If StrComp(CStr(leftConflicts(leftIndex, leftColumnLow)), CStr(rightConflicts(rightIndex, rightColumnLow)), vbBinaryCompare) <> 0 Then Exit Function
        If CLng(leftConflicts(leftIndex, leftColumnLow + 1)) <> CLng(rightConflicts(rightIndex, rightColumnLow + 1)) Then Exit Function
        If CLng(leftConflicts(leftIndex, leftColumnLow + 2)) <> CLng(rightConflicts(rightIndex, rightColumnLow + 2)) Then Exit Function
        For valueIndex = 3 To 5
            If Not CompareCellValues(leftConflicts(leftIndex, leftColumnLow + valueIndex), rightConflicts(rightIndex, rightColumnLow + valueIndex), equal, comparisonError) Then Exit Function
            If Not equal Then Exit Function
        Next valueIndex
    Next leftIndex
    ConflictViewsEqual = True
End Function

Private Function ApplySessionConflictChoices(ByRef merges() As MergeResult, ByRef choices As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim choiceIndex As Long
    Dim groupIndex As Long
    Dim matchCount As Long
    Dim outputRow As Long
    Dim assignedCount As Long
    Dim groupChoices() As Variant
    Dim errorCheck As AppErrorInfo

    If Not GetMatrixBounds(choices, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidChoices
    If columnHigh - columnLow + 1 <> 4 Then GoTo InvalidChoices
    For choiceIndex = rowLow To rowHigh
        If VarType(choices(choiceIndex, columnLow)) <> vbString Then GoTo InvalidChoices
        If VarType(choices(choiceIndex, columnLow + 1)) = vbString Or Not IsNumeric(choices(choiceIndex, columnLow + 1)) Then GoTo InvalidChoices
        If VarType(choices(choiceIndex, columnLow + 2)) = vbString Or Not IsNumeric(choices(choiceIndex, columnLow + 2)) Then GoTo InvalidChoices
        If VarType(choices(choiceIndex, columnLow + 3)) <> vbString Then GoTo InvalidChoices
    Next choiceIndex

    For groupIndex = 1 To mGroupCount
        matchCount = 0
        For choiceIndex = rowLow To rowHigh
            If StrComp(CStr(choices(choiceIndex, columnLow)), CStr(mSheetNames(groupIndex)), vbTextCompare) = 0 Then matchCount = matchCount + 1
        Next choiceIndex
        If matchCount > 0 Then
            ReDim groupChoices(1 To matchCount, 1 To 3)
            outputRow = 0
            For choiceIndex = rowLow To rowHigh
                If StrComp(CStr(choices(choiceIndex, columnLow)), CStr(mSheetNames(groupIndex)), vbTextCompare) = 0 Then
                    outputRow = outputRow + 1
                    groupChoices(outputRow, 1) = choices(choiceIndex, columnLow + 1)
                    groupChoices(outputRow, 2) = choices(choiceIndex, columnLow + 2)
                    groupChoices(outputRow, 3) = choices(choiceIndex, columnLow + 3)
                End If
            Next choiceIndex
            If Not ApplyConflictChoices(merges(groupIndex), groupChoices, errorInfo) Then Exit Function
            assignedCount = assignedCount + matchCount
        ElseIf merges(groupIndex).ConflictCount > 0 Then
            GoTo InvalidChoices
        End If
    Next groupIndex
    If assignedCount <> rowHigh - rowLow + 1 Then GoTo InvalidChoices
    ApplySessionConflictChoices = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.ApplySessionConflictChoices", "")
    Exit Function

InvalidChoices:
    errorInfo = MakeAppError(ERR_SESSION_INVALID_TARGETS, "MdlSession.ApplySessionConflictChoices", "每个冲突必须且只能对应一个 Local 或 Remote 选择。")
End Function

Private Sub ScheduleHeartbeat()
    Dim errorInfo As AppErrorInfo
    If ScheduleSessionHeartbeat(HEARTBEAT_INTERVAL_SECONDS, errorInfo) Then Exit Sub
    mCanSubmit = False
    mStatusText = "Automatic lease renewal could not be scheduled"
    LogWrite LogError, "MdlSession.ScheduleHeartbeat", errorInfo.Message & " " & errorInfo.Context
End Sub

Private Sub CancelHeartbeat()
    CancelSessionHeartbeat
End Sub

Private Function NormalizeTargets(ByRef targets As Variant, ByVal expectedFile As String, ByRef normalizedTargets As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim rowIndex As Long
    Dim outputIndex As Long
    Dim filePath As String
    Dim normalizedFile As String
    Dim result() As Variant

    If Not GetMatrixBounds(targets, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidTargets
    If columnHigh - columnLow + 1 <> 3 Or rowHigh < rowLow Then GoTo InvalidTargets
    ReDim result(1 To rowHigh - rowLow + 1, 1 To 3)
    outputIndex = 0
    For rowIndex = rowLow To rowHigh
        outputIndex = outputIndex + 1
        If VarType(targets(rowIndex, columnLow)) <> vbString Or VarType(targets(rowIndex, columnLow + 1)) <> vbString Or VarType(targets(rowIndex, columnLow + 2)) <> vbString Then GoTo InvalidTargets
        If Not ValidateUncPath(CStr(targets(rowIndex, columnLow)), normalizedFile, errorInfo) Then Exit Function
        If StrComp(normalizedFile, expectedFile, vbTextCompare) <> 0 Then GoTo InvalidTargets
        filePath = normalizedFile
        result(outputIndex, 1) = filePath
        result(outputIndex, 2) = Trim$(CStr(targets(rowIndex, columnLow + 1)))
        result(outputIndex, 3) = Trim$(CStr(targets(rowIndex, columnLow + 2)))
        If Len(CStr(result(outputIndex, 2))) = 0 Or Len(CStr(result(outputIndex, 3))) = 0 Then GoTo InvalidTargets
    Next rowIndex
    normalizedTargets = result
    NormalizeTargets = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.NormalizeTargets", "")
    Exit Function

InvalidTargets:
    normalizedTargets = Empty
    errorInfo = MakeAppError(ERR_SESSION_INVALID_TARGETS, "MdlSession.NormalizeTargets", "Targets must contain the configured UNC file, a sheet name, and an AMG header.")
End Function

Private Function BuildTargetGroups(ByRef targets As Variant, ByRef sheetNames As Variant, ByRef groupTargets As Variant, ByRef groupCount As Long, ByRef errorInfo As AppErrorInfo) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim rowIndex As Long
    Dim groupIndex As Long
    Dim targetCount As Long
    Dim currentName As String
    Dim names() As String
    Dim counts() As Long
    Dim offsets() As Long
    Dim namesResult() As Variant
    Dim groupsResult() As Variant
    Dim targetIndex As Long
    Dim columnIndex As Long
    Dim grouped() As Variant

    If Not GetMatrixBounds(targets, rowLow, rowHigh, columnLow, columnHigh) Then GoTo InvalidTargets
    targetCount = rowHigh - rowLow + 1
    ReDim names(1 To targetCount)
    ReDim counts(1 To targetCount)
    ReDim offsets(1 To targetCount)

    For rowIndex = rowLow To rowHigh
        currentName = CStr(targets(rowIndex, columnLow + 1))
        groupIndex = FindSheetIndex(names, groupCount, currentName)
        If groupIndex = 0 Then
            groupCount = groupCount + 1
            names(groupCount) = currentName
            groupIndex = groupCount
        End If
        counts(groupIndex) = counts(groupIndex) + 1
    Next rowIndex

    ReDim namesResult(1 To groupCount)
    ReDim groupsResult(1 To groupCount)
    For groupIndex = 1 To groupCount
        namesResult(groupIndex) = names(groupIndex)
        ReDim grouped(1 To counts(groupIndex), 1 To 2)
        groupsResult(groupIndex) = grouped
    Next groupIndex

    For rowIndex = rowLow To rowHigh
        currentName = CStr(targets(rowIndex, columnLow + 1))
        groupIndex = FindSheetIndex(names, groupCount, currentName)
        grouped = groupsResult(groupIndex)
        offsets(groupIndex) = offsets(groupIndex) + 1
        grouped(offsets(groupIndex), 1) = currentName
        grouped(offsets(groupIndex), 2) = CStr(targets(rowIndex, columnLow + 2))
        groupsResult(groupIndex) = grouped
    Next rowIndex

    sheetNames = namesResult
    groupTargets = groupsResult
    BuildTargetGroups = True
    errorInfo = MakeAppError(ERR_NONE, "MdlSession.BuildTargetGroups", "")
    Exit Function

InvalidTargets:
    groupCount = 0
    sheetNames = Empty
    groupTargets = Empty
    errorInfo = MakeAppError(ERR_SESSION_INVALID_TARGETS, "MdlSession.BuildTargetGroups", "The checkout target matrix is invalid.")
End Function

Private Function FindSheetIndex(ByRef names() As String, ByVal groupCount As Long, ByVal sheetName As String) As Long
    Dim index As Long
    For index = 1 To groupCount
        If StrComp(names(index), sheetName, vbTextCompare) = 0 Then
            FindSheetIndex = index
            Exit Function
        End If
    Next index
End Function

Private Function LockStatusText(ByVal status As LockAvailability) As String
    Select Case status
        Case LockFree: LockStatusText = "空闲"
        Case LockHeld: LockStatusText = "已占用"
        Case LockExpired: LockStatusText = "租约过期，可尝试回收"
        Case Else: LockStatusText = "未知，签出时验证"
    End Select
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

Private Sub RollbackCheckout(ByRef errorInfo As AppErrorInfo)
    Dim cleanupError As AppErrorInfo
    Dim releaseFailures As Variant

    CancelHeartbeat
    If Len(mLocalHandle) > 0 Then CloseWorkbookHandle mLocalHandle, False
    mLocalHandle = vbNullString
    If Len(mWorkspacePath) > 0 Then
        If Not RemoveLocalWorkspace(mWorkspacePath, cleanupError) Then errorInfo = cleanupError
    End If
    If Not IsEmpty(mLeases) Then
        If Not ReleaseLocks(mLeases, releaseFailures) Then errorInfo = MakeAppError(ERR_IO, "MdlSession.RollbackCheckout", "One or more acquired column locks could not be released.")
    End If
    ResetSession
End Sub

Private Function ReleaseSessionResources(ByRef errorInfo As AppErrorInfo) As Boolean
    Dim cleanupError As AppErrorInfo
    Dim releaseFailures As Variant
    Dim successful As Boolean

    successful = True
    CancelHeartbeat
    If Len(mLocalHandle) > 0 Then CloseWorkbookHandle mLocalHandle, False
    mLocalHandle = vbNullString
    If Len(mWorkspacePath) > 0 Then
        If Not RemoveLocalWorkspace(mWorkspacePath, cleanupError) Then
            successful = False
            errorInfo = cleanupError
        End If
    End If
    If Not IsEmpty(mLeases) Then
        If Not ReleaseLocks(mLeases, releaseFailures) Then
            successful = False
            errorInfo = MakeAppError(ERR_IO, "MdlSession.ReleaseSessionResources", "One or more column locks could not be released; they will remain until expiry.")
        End If
    End If
    ReleaseSessionResources = successful
End Function

Private Sub ResetSession()
    CancelHeartbeat
    mState = SessionIdle
    mSessionId = vbNullString
    mMainFilePath = vbNullString
    mWorkspacePath = vbNullString
    mLocalFilePath = vbNullString
    mLocalHandle = vbNullString
    mLeases = Empty
    mTargets = Empty
    mSheetNames = Empty
    mSheetTargets = Empty
    mTargetCount = 0
    mGroupCount = 0
    Erase mBaseSnapshots
    mPendingConflicts = Empty
    mCanSubmit = False
    mSaveResultUncertain = False
    mStatusText = vbNullString
End Sub
