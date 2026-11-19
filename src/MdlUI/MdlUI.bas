Attribute VB_Name = "MdlUI"
Option Explicit

Private Const TARGET_FORM_WIDTH As Double = 680#
Private Const TARGET_FORM_HEIGHT As Double = 430#
Private Const CONFLICT_FORM_WIDTH As Double = 680#
Private Const CONFLICT_FORM_HEIGHT As Double = 520#
Private mRibbon As IRibbonUI

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    On Error GoTo Failed
    Set mRibbon = ribbon
    Exit Sub

Failed:
    LogWrite LogError, "MdlUI.RibbonOnLoad", CStr(Err.Number) & ": " & Err.Description
End Sub

Public Sub RibbonGetEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
    Dim view As SessionView

    On Error GoTo Failed
    view = GetSessionView()
    enabled = IsRibbonButtonEnabled(control.Id, view)
    Exit Sub

Failed:
    enabled = False
    LogWrite LogError, "MdlUI.RibbonGetEnabled", CStr(Err.Number) & ": " & Err.Description
End Sub

Public Function IsRibbonButtonEnabled(ByVal buttonId As String, ByRef view As SessionView) As Boolean
    Select Case buttonId
        Case "btnStartEdit"
            IsRibbonButtonEnabled = (view.State = SessionIdle)
        Case "btnSubmit"
            IsRibbonButtonEnabled = view.CanSubmit
        Case "btnAbandon"
            IsRibbonButtonEnabled = view.CanAbandon
    End Select
End Function

Public Sub RibbonStartEdit(ByVal control As IRibbonControl)
    Dim workbookPaths As Variant
    Dim workbookPath As String
    Dim availableTargets As Variant
    Dim selectedTargets As Variant
    Dim errorInfo As AppErrorInfo

    On Error GoTo Failed
    If Not GetConfiguredWorkbookPaths(workbookPaths, errorInfo) Then GoTo ShowError
    If Not ShowWorkbookSelector(workbookPaths, workbookPath) Then Exit Sub
    If Not GetAvailableTargets(workbookPath, availableTargets, errorInfo) Then GoTo ShowError
    If IsEmpty(availableTargets) Then
        MsgBox "No editable AMG columns are available.", vbInformation, "MDDT Collaboration"
        Exit Sub
    End If
    If Not ShowTargetSelector(availableTargets, selectedTargets) Then Exit Sub
    If Not BeginCheckout(workbookPath, selectedTargets, errorInfo) Then GoTo ShowError
    InvalidateRibbon
    MsgBox "The local working copy is open. You can start editing.", vbInformation, "MDDT Collaboration"
    Exit Sub

ShowError:
    If Len(errorInfo.Context) > 0 Then
        MsgBox errorInfo.Message & vbCrLf & errorInfo.Context, vbExclamation, "Unable to Start Editing"
    Else
        MsgBox errorInfo.Message, vbExclamation, "Unable to Start Editing"
    End If
    Exit Sub

Failed:
    LogWrite LogError, "MdlUI.RibbonStartEdit", CStr(Err.Number) & ": " & Err.Description
    MsgBox "An error occurred while starting the edit session. Check the log and session state.", vbExclamation, "Unable to Start Editing"
End Sub

Private Function ShowWorkbookSelector(ByRef paths As Variant, ByRef selectedPath As String) As Boolean
    Dim form As Object
    Dim listBox As Object
    Dim heading As Object
    Dim acceptButton As Object
    Dim cancelButton As Object
    Dim handlers As CTargetSelectorEvents
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long

    selectedPath = vbNullString
    On Error GoTo Failed
    If Not GetVectorBounds(paths, lowerBound, upperBound) Then Exit Function

    Set form = New frmMddtDialog
    With form
        .Caption = "Select Workbook"
        .Width = TARGET_FORM_WIDTH
        .Height = TARGET_FORM_HEIGHT
        .StartUpPosition = 1
    End With
    Set heading = form.Controls.Add("Forms.Label.1", "lblWorkbookHeading", True)
    With heading
        .Caption = "Select one configured shared workbook (UNC path)."
        .Left = 12
        .Top = 10
        .Width = 640
        .Height = 20
    End With
    Set listBox = form.Controls.Add("Forms.ListBox.1", "lstWorkbooks", True)
    With listBox
        .Left = 12
        .Top = 34
        .Width = 640
        .Height = 320
        .ColumnCount = 1
        .ColumnWidths = "620 pt"
        .MultiSelect = 0
    End With
    For index = lowerBound To upperBound
        listBox.AddItem CStr(paths(index))
    Next index

    Set acceptButton = form.Controls.Add("Forms.CommandButton.1", "cmdWorkbookAccept", True)
    With acceptButton
        .Caption = "Continue"
        .Left = 424
        .Top = 366
        .Width = 105
        .Height = 28
        .Default = True
    End With
    Set cancelButton = form.Controls.Add("Forms.CommandButton.1", "cmdWorkbookCancel", True)
    With cancelButton
        .Caption = "Cancel"
        .Left = 544
        .Top = 366
        .Width = 105
        .Height = 28
        .Cancel = True
    End With
    Set handlers = New CTargetSelectorEvents
    handlers.Bind form, listBox, acceptButton, cancelButton, "Select one workbook from the list."
    form.Show 1
    If Not handlers.Accepted Then GoTo Cleanup
    selectedPath = CStr(listBox.List(handlers.SelectedIndex))
    ShowWorkbookSelector = (Len(selectedPath) > 0)
    GoTo Cleanup

Failed:
    selectedPath = vbNullString
    LogWrite LogError, "MdlUI.ShowWorkbookSelector", CStr(Err.Number) & ": " & Err.Description
    MsgBox "Unable to open the workbook selector. Checkout was not started.", vbExclamation, "Unable to Select Workbook"

Cleanup:
    On Error Resume Next
    Set handlers = Nothing
    If Not form Is Nothing Then form.Hide
    Set listBox = Nothing
    Set form = Nothing
    On Error GoTo 0
End Function

Public Sub RibbonSubmit(ByVal control As IRibbonControl)
    Dim outcome As SubmitOutcome
    Dim conflicts As Variant
    Dim choices As Variant
    Dim errorInfo As AppErrorInfo

    On Error GoTo Failed
    outcome = SubmitChanges(conflicts, errorInfo)
    Do While outcome = SubmitNeedsConflictReview
        If Not ShowConflictReview(conflicts, choices) Then
            Call CancelConflictReview
            InvalidateRibbon
            MsgBox "Conflict review was canceled. The local copy and column locks are retained; you can submit later.", vbInformation, "Submission Canceled"
            Exit Sub
        End If
        outcome = ResumeAfterConflictReview(choices, conflicts, errorInfo)
    Loop

    InvalidateRibbon
    Select Case outcome
        Case SubmitCompleted
            MsgBox "Submission complete.", vbInformation, "MDDT Collaboration"
        Case SubmitCompletedWithCleanupWarning
            MsgBox "Shared data was submitted, but local cleanup is incomplete." & vbCrLf & errorInfo.Message, vbExclamation, "Submission Complete"
        Case Else
            MsgBox errorInfo.Message, vbExclamation, "Submission Incomplete"
    End Select
    Exit Sub

Failed:
    LogWrite LogError, "MdlUI.RibbonSubmit", CStr(Err.Number) & ": " & Err.Description
    InvalidateRibbon
    MsgBox "An error occurred during submission. The local session has been retained. Check the session state and log.", vbExclamation, "Submission Incomplete"
End Sub

Public Sub RibbonAbandon(ByVal control As IRibbonControl)
    Dim errorInfo As AppErrorInfo
    Dim answer As VbMsgBoxResult

    On Error GoTo Failed
    answer = MsgBox("Abandon this edit session and close the local copy? Unsubmitted local changes will be discarded.", vbQuestion Or vbYesNo, "Confirm Abandon")
    If answer <> vbYes Then Exit Sub
    If Not AbandonSession(True, errorInfo) Then
        MsgBox errorInfo.Message, vbExclamation, "Unable to Abandon Session"
    Else
        MsgBox "This edit session was abandoned.", vbInformation, "MDDT Collaboration"
    End If
    InvalidateRibbon
    Exit Sub

Failed:
    LogWrite LogError, "MdlUI.RibbonAbandon", CStr(Err.Number) & ": " & Err.Description
    InvalidateRibbon
    MsgBox "An error occurred while abandoning the session. Check the session state and log.", vbExclamation, "Unable to Abandon Session"
End Sub

Public Function BuildConflictChoiceMatrix(ByRef conflicts As Variant, ByRef selections As Variant, ByRef choices As Variant) As Boolean
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim selectionLow As Long
    Dim selectionHigh As Long
    Dim rowIndex As Long
    Dim selectionIndex As Long
    Dim result() As Variant
    Dim rowCount As Long

    choices = Empty
    On Error GoTo InvalidInput
    If Not GetMatrixBounds(conflicts, rowLow, rowHigh, columnLow, columnHigh) Then Exit Function
    If columnHigh - columnLow + 1 <> 7 Then Exit Function
    If Not GetVectorBounds(selections, selectionLow, selectionHigh) Then Exit Function
    If rowHigh - rowLow <> selectionHigh - selectionLow Then Exit Function

    rowCount = rowHigh - rowLow + 1
    ReDim result(1 To rowCount, 1 To 5)
    For rowIndex = rowLow To rowHigh
        selectionIndex = selectionLow + rowIndex - rowLow
        If VarType(conflicts(rowIndex, columnLow)) <> vbString Then Exit Function
        If VarType(conflicts(rowIndex, columnLow + 1)) = vbString Or Not IsNumeric(conflicts(rowIndex, columnLow + 1)) Then Exit Function
        If VarType(conflicts(rowIndex, columnLow + 2)) = vbString Or Not IsNumeric(conflicts(rowIndex, columnLow + 2)) Then Exit Function
        If CDbl(conflicts(rowIndex, columnLow + 1)) <> CDbl(CLng(conflicts(rowIndex, columnLow + 1))) Then Exit Function
        If CDbl(conflicts(rowIndex, columnLow + 2)) <> CDbl(CLng(conflicts(rowIndex, columnLow + 2))) Then Exit Function
        If VarType(conflicts(rowIndex, columnLow + 3)) <> vbString Then Exit Function
        If CStr(conflicts(rowIndex, columnLow + 3)) <> "TextFont" And CStr(conflicts(rowIndex, columnLow + 3)) <> "Fill" Then Exit Function
        If VarType(selections(selectionIndex)) <> vbString Then Exit Function
        If StrComp(CStr(selections(selectionIndex)), "Local", vbBinaryCompare) <> 0 And StrComp(CStr(selections(selectionIndex)), "Remote", vbBinaryCompare) <> 0 Then Exit Function
        result(rowIndex - rowLow + 1, 1) = CStr(conflicts(rowIndex, columnLow))
        result(rowIndex - rowLow + 1, 2) = conflicts(rowIndex, columnLow + 1)
        result(rowIndex - rowLow + 1, 3) = conflicts(rowIndex, columnLow + 2)
        result(rowIndex - rowLow + 1, 4) = CStr(conflicts(rowIndex, columnLow + 3))
        result(rowIndex - rowLow + 1, 5) = CStr(selections(selectionIndex))
    Next rowIndex

    choices = result
    BuildConflictChoiceMatrix = True
    On Error GoTo 0
    Exit Function

InvalidInput:
    choices = Empty
    Err.Clear
    On Error GoTo 0
End Function

Public Function FormatConflictValue(ByVal value As Variant) As String
    Dim formatted As String

    If IsEmpty(value) Then
        FormatConflictValue = "(blank)"
        Exit Function
    End If
    If IsNull(value) Then
        FormatConflictValue = "(invalid value)"
        Exit Function
    End If
    If IsError(value) Then
        FormatConflictValue = "(error value)"
        Exit Function
    End If
    If VarType(value) <> vbString Then
        FormatConflictValue = "(non-string value)"
        Exit Function
    End If

    formatted = Replace$(CStr(value), vbCrLf, vbLf)
    formatted = Replace$(formatted, vbCr, vbLf)
    FormatConflictValue = formatted
End Function

Public Function FormatConflictComponent(ByVal componentName As String, ByVal value As Variant) As String
    If componentName = "TextFont" Or componentName = "Fill" Then
        FormatConflictComponent = FormatFormattingPreview(value)
    Else
        FormatConflictComponent = FormatConflictValue(value)
    End If
End Function

Public Function ExcelColumnLabel(ByVal columnIndex As Long) As String
    Dim remaining As Long

    If columnIndex < 1 Then Exit Function
    remaining = columnIndex
    Do While remaining > 0
        remaining = remaining - 1
        ExcelColumnLabel = Chr$((remaining Mod 26) + Asc("A")) & ExcelColumnLabel
        remaining = remaining \ 26
    Loop
End Function

Private Function ShowTargetSelector(ByRef targets As Variant, ByRef selectedTargets As Variant) As Boolean
    Dim form As Object
    Dim listBox As Object
    Dim heading As Object
    Dim acceptButton As Object
    Dim cancelButton As Object
    Dim handlers As CTargetSelectorEvents
    Dim rowLow As Long
    Dim rowHigh As Long
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim rowIndex As Long
    Dim listIndex As Long
    Dim selectedCount As Long
    Dim outputIndex As Long
    Dim result() As Variant

    selectedTargets = Empty
    On Error GoTo Failed
    If Not GetMatrixBounds(targets, rowLow, rowHigh, columnLow, columnHigh) Then Exit Function
    If columnHigh - columnLow + 1 <> 4 Then Exit Function

    Set form = New frmMddtDialog
    With form
        .Caption = "Select AMG Columns to Check Out"
        .Width = TARGET_FORM_WIDTH
        .Height = TARGET_FORM_HEIGHT
        .StartUpPosition = 1
    End With
    Set heading = form.Controls.Add("Forms.Label.1", "lblTargetHeading", True)
    With heading
        .Caption = "Worksheet / AMG Column / Lock Status"
        .Left = 12
        .Top = 10
        .Width = 630
        .Height = 20
    End With
    Set listBox = form.Controls.Add("Forms.ListBox.1", "lstTargets", True)
    With listBox
        .Left = 12
        .Top = 34
        .Width = 640
        .Height = 320
        .ColumnCount = 4
        .ColumnWidths = "170 pt;190 pt;150 pt;0 pt"
        .MultiSelect = 1
    End With
    For rowIndex = rowLow To rowHigh
        listBox.AddItem CStr(targets(rowIndex, columnLow + 1))
        listIndex = listBox.ListCount - 1
        listBox.List(listIndex, 1) = CStr(targets(rowIndex, columnLow + 2))
        listBox.List(listIndex, 2) = CStr(targets(rowIndex, columnLow + 3))
        listBox.List(listIndex, 3) = CStr(targets(rowIndex, columnLow))
    Next rowIndex

    Set acceptButton = form.Controls.Add("Forms.CommandButton.1", "cmdTargetAccept", True)
    With acceptButton
        .Caption = "Start Editing"
        .Left = 424
        .Top = 366
        .Width = 105
        .Height = 28
        .Default = True
    End With
    Set cancelButton = form.Controls.Add("Forms.CommandButton.1", "cmdTargetCancel", True)
    With cancelButton
        .Caption = "Cancel"
        .Left = 544
        .Top = 366
        .Width = 105
        .Height = 28
        .Cancel = True
    End With
    Set handlers = New CTargetSelectorEvents
    handlers.Bind form, listBox, acceptButton, cancelButton, "Select at least one AMG column."
    form.Show 1
    If Not handlers.Accepted Then GoTo Cleanup

    For listIndex = 0 To listBox.ListCount - 1
        If listBox.Selected(listIndex) Then selectedCount = selectedCount + 1
    Next listIndex
    If selectedCount = 0 Then GoTo Cleanup
    ReDim result(1 To selectedCount, 1 To 3)
    For listIndex = 0 To listBox.ListCount - 1
        If listBox.Selected(listIndex) Then
            outputIndex = outputIndex + 1
            result(outputIndex, 1) = CStr(listBox.List(listIndex, 3))
            result(outputIndex, 2) = CStr(listBox.List(listIndex, 0))
            result(outputIndex, 3) = CStr(listBox.List(listIndex, 1))
        End If
    Next listIndex
    selectedTargets = result
    ShowTargetSelector = True
    GoTo Cleanup

Failed:
    ShowTargetSelector = False
    LogWrite LogError, "MdlUI.ShowTargetSelector", CStr(Err.Number) & ": " & Err.Description
    MsgBox "Unable to open the target selector. Checkout was not started.", vbExclamation, "Unable to Select Targets"

Cleanup:
    On Error Resume Next
    Set handlers = Nothing
    If Not form Is Nothing Then form.Hide
    Set listBox = Nothing
    Set form = Nothing
    On Error GoTo 0
End Function

Private Function ShowConflictReview(ByRef conflicts As Variant, ByRef choices As Variant) As Boolean
    Dim form As Object
    Dim locationLabel As Object
    Dim countLabel As Object
    Dim baseValue As Object
    Dim localValue As Object
    Dim remoteValue As Object
    Dim previousButton As Object
    Dim nextButton As Object
    Dim localButton As Object
    Dim remoteButton As Object
    Dim finishButton As Object
    Dim cancelButton As Object
    Dim handlers As CConflictReviewEvents

    choices = Empty
    On Error GoTo Failed
    If Not IsArray(conflicts) Then Exit Function
    Set form = New frmMddtDialog
    With form
        .Caption = "Review Cell Conflicts"
        .Width = CONFLICT_FORM_WIDTH
        .Height = CONFLICT_FORM_HEIGHT
        .StartUpPosition = 1
    End With
    Set locationLabel = AddLabel(form, "lblConflictLocation", 12, 10, 620, 24, vbNullString)
    Set countLabel = AddLabel(form, "lblConflictCount", 12, 34, 620, 20, vbNullString)
    Set baseValue = AddValueBox(form, "txtBaseValue", "Base", 12, 62)
    Set localValue = AddValueBox(form, "txtLocalValue", "Local", 12, 160)
    Set remoteValue = AddValueBox(form, "txtRemoteValue", "Remote", 12, 258)

    Set previousButton = AddButton(form, "cmdPrevious", "Previous", 12, 384, 92)
    Set nextButton = AddButton(form, "cmdNext", "Next", 116, 384, 92)
    Set localButton = AddButton(form, "cmdChooseLocal", "Keep Local", 224, 384, 130)
    Set remoteButton = AddButton(form, "cmdChooseRemote", "Keep Remote", 370, 384, 130)
    Set finishButton = AddButton(form, "cmdFinish", "Apply Choices", 516, 384, 130)
    finishButton.Default = True
    Set cancelButton = AddButton(form, "cmdCancel", "Cancel", 516, 422, 130)
    cancelButton.Cancel = True

    Set handlers = New CConflictReviewEvents
    handlers.Bind form, conflicts, locationLabel, countLabel, baseValue, localValue, remoteValue, previousButton, nextButton, localButton, remoteButton, finishButton, cancelButton
    form.Show 1
    If Not handlers.Confirmed Then GoTo Cleanup
    choices = handlers.ChoiceMatrix
    ShowConflictReview = Not IsEmpty(choices)
    GoTo Cleanup

Failed:
    ShowConflictReview = False
    LogWrite LogError, "MdlUI.ShowConflictReview", CStr(Err.Number) & ": " & Err.Description
    MsgBox "Unable to open conflict review. No shared data was written; the local session is retained.", vbExclamation, "Conflict Review Unavailable"

Cleanup:
    On Error Resume Next
    Set handlers = Nothing
    If Not form Is Nothing Then form.Hide
    Set form = Nothing
    On Error GoTo 0
End Function

Private Function AddLabel(ByRef form As Object, ByVal name As String, ByVal left As Double, ByVal top As Double, ByVal width As Double, ByVal height As Double, ByVal caption As String) As Object
    Set AddLabel = form.Controls.Add("Forms.Label.1", name, True)
    With AddLabel
        .Caption = caption
        .Left = left
        .Top = top
        .Width = width
        .Height = height
        .WordWrap = True
    End With
End Function

Private Function AddValueBox(ByRef form As Object, ByVal name As String, ByVal labelText As String, ByVal left As Double, ByVal top As Double) As Object
    Dim title As Object
    Set title = AddLabel(form, "lbl" & name, left, top, 620, 18, labelText)
    Set AddValueBox = form.Controls.Add("Forms.TextBox.1", name, True)
    With AddValueBox
        .Left = left
        .Top = top + 19
        .Width = 620
        .Height = 74
        .MultiLine = True
        .ScrollBars = 2
        .WordWrap = True
        .Locked = True
        .MaxLength = 0
        .TabStop = False
    End With
End Function

Private Function AddButton(ByRef form As Object, ByVal name As String, ByVal caption As String, ByVal left As Double, ByVal top As Double, ByVal width As Double) As Object
    Set AddButton = form.Controls.Add("Forms.CommandButton.1", name, True)
    With AddButton
        .Caption = caption
        .Left = left
        .Top = top
        .Width = width
        .Height = 28
    End With
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

Private Sub InvalidateRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
    On Error GoTo 0
End Sub
