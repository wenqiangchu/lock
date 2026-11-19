Attribute VB_Name = "TestMdlSession"
Option Explicit

Public Sub RunMdlSessionTests()
    TestTransitionRules
    TestIdleSessionRejectsActions
    TestIdleSessionView
End Sub

Private Sub TestTransitionRules()
    AssertTrue IsSessionTransitionAllowed(SessionIdle, SessionTargetSelected), "Idle should allow target selection."
    AssertFalse IsSessionTransitionAllowed(SessionIdle, SessionCommitting), "Idle must not allow commit."
    AssertTrue IsSessionTransitionAllowed(SessionCheckingOut, SessionCheckedOut), "Checkout should finish in CheckedOut."
    AssertTrue IsSessionTransitionAllowed(SessionCheckedOut, SessionCommitting), "CheckedOut should allow commit."
    AssertTrue IsSessionTransitionAllowed(SessionCommitting, SessionConflictReview), "Commit should allow conflict review."
    AssertTrue IsSessionTransitionAllowed(SessionConflictReview, SessionCommitting), "Conflict review should allow revalidation."
    AssertTrue IsSessionTransitionAllowed(SessionAborting, SessionIdle), "Aborting should finish in Idle."
    AssertFalse IsSessionTransitionAllowed(SessionAborting, SessionCheckedOut), "Aborting must not resume editing."
End Sub

Private Sub TestIdleSessionRejectsActions()
    Dim conflicts As Variant
    Dim choices As Variant
    Dim errorInfo As AppErrorInfo
    Dim outcome As SubmitOutcome

    outcome = SubmitChanges(conflicts, errorInfo)
    AssertEqual CStr(outcome), CStr(SubmitKeepSession), "Submitting while Idle should not complete."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_SESSION_INVALID_STATE), "Idle submit should return a state error."
    AssertFalse CancelConflictReview(), "Idle cannot cancel conflict review."
    AssertFalse AbandonSession(False, errorInfo), "Idle cannot be abandoned without confirmation."

    choices = Empty
    outcome = ResumeAfterConflictReview(choices, conflicts, errorInfo)
    AssertEqual CStr(outcome), CStr(SubmitKeepSession), "Idle cannot resume conflict review."
    AssertEqual CStr(errorInfo.Code), CStr(ERR_SESSION_INVALID_STATE), "Idle resume should return a state error."
End Sub

Private Sub TestIdleSessionView()
    Dim view As SessionView

    view = GetSessionView()
    AssertEqual CStr(view.State), CStr(SessionIdle), "A reset session should report Idle."
    AssertEqual CStr(view.TargetCount), "0", "Idle should have no targets."
    AssertEqual CStr(view.ConflictCount), "0", "Idle should have no conflicts."
    AssertFalse view.CanSubmit, "Idle should not allow submit."
    AssertFalse view.CanAbandon, "Idle should not allow abandon."
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2501, "TestMdlSession", message
End Sub

Private Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, ByVal message As String)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 2502, "TestMdlSession", message
End Sub
