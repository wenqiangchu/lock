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
    Dim availableTargets As Variant
    Dim selectedTargets As Variant
    Dim errorInfo As AppErrorInfo

    On Error GoTo Failed
    If Not GetAvailableTargets(availableTargets, errorInfo) Then GoTo ShowError
    If IsEmpty(availableTargets) Then
        MsgBox "没有可签出的 AMG 列。", vbInformation, "MDDT 并发编辑"
        Exit Sub
    End If
    If Not ShowTargetSelector(availableTargets, selectedTargets) Then Exit Sub
    If Not BeginCheckout(selectedTargets, errorInfo) Then GoTo ShowError
    InvalidateRibbon
    MsgBox "本地工作副本已打开，可以开始编辑。", vbInformation, "MDDT 并发编辑"
    Exit Sub

ShowError:
    MsgBox errorInfo.Message, vbExclamation, "无法开始编辑"
    Exit Sub

Failed:
    LogWrite LogError, "MdlUI.RibbonStartEdit", CStr(Err.Number) & ": " & Err.Description
    MsgBox "开始编辑时发生错误；请检查日志和会话状态。", vbExclamation, "无法开始编辑"
End Sub

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
            MsgBox "已取消冲突裁决。本地副本和列锁仍保留，可以稍后继续提交。", vbInformation, "提交已取消"
            Exit Sub
        End If
        outcome = ResumeAfterConflictReview(choices, conflicts, errorInfo)
    Loop

    InvalidateRibbon
    Select Case outcome
        Case SubmitCompleted
            MsgBox "提交完成。", vbInformation, "MDDT 并发编辑"
        Case SubmitCompletedWithCleanupWarning
            MsgBox "共享数据已提交，但本地清理未全部完成。" & vbCrLf & errorInfo.Message, vbExclamation, "提交完成"
        Case Else
            MsgBox errorInfo.Message, vbExclamation, "提交未完成"
    End Select
    Exit Sub

Failed:
    LogWrite LogError, "MdlUI.RibbonSubmit", CStr(Err.Number) & ": " & Err.Description
    InvalidateRibbon
    MsgBox "提交过程中发生错误；本地会话状态已保留，请检查会话状态和日志。", vbExclamation, "提交未完成"
End Sub

Public Sub RibbonAbandon(ByVal control As IRibbonControl)
    Dim errorInfo As AppErrorInfo
    Dim answer As VbMsgBoxResult

    On Error GoTo Failed
    answer = MsgBox("放弃本次编辑并关闭本地副本？未提交的本地改动将被丢弃。", vbQuestion Or vbYesNo, "确认放弃")
    If answer <> vbYes Then Exit Sub
    If Not AbandonSession(True, errorInfo) Then
        MsgBox errorInfo.Message, vbExclamation, "无法放弃会话"
    Else
        MsgBox "本次会话已放弃。", vbInformation, "MDDT 并发编辑"
    End If
    InvalidateRibbon
    Exit Sub

Failed:
    LogWrite LogError, "MdlUI.RibbonAbandon", CStr(Err.Number) & ": " & Err.Description
    InvalidateRibbon
    MsgBox "放弃会话时发生错误；请检查会话状态和日志。", vbExclamation, "无法放弃会话"
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
    If columnHigh - columnLow + 1 <> 6 Then Exit Function
    If Not GetVectorBounds(selections, selectionLow, selectionHigh) Then Exit Function
    If rowHigh - rowLow <> selectionHigh - selectionLow Then Exit Function

    rowCount = rowHigh - rowLow + 1
    ReDim result(1 To rowCount, 1 To 4)
    For rowIndex = rowLow To rowHigh
        selectionIndex = selectionLow + rowIndex - rowLow
        If VarType(conflicts(rowIndex, columnLow)) <> vbString Then Exit Function
        If VarType(conflicts(rowIndex, columnLow + 1)) = vbString Or Not IsNumeric(conflicts(rowIndex, columnLow + 1)) Then Exit Function
        If VarType(conflicts(rowIndex, columnLow + 2)) = vbString Or Not IsNumeric(conflicts(rowIndex, columnLow + 2)) Then Exit Function
        If CDbl(conflicts(rowIndex, columnLow + 1)) <> CDbl(CLng(conflicts(rowIndex, columnLow + 1))) Then Exit Function
        If CDbl(conflicts(rowIndex, columnLow + 2)) <> CDbl(CLng(conflicts(rowIndex, columnLow + 2))) Then Exit Function
        If VarType(selections(selectionIndex)) <> vbString Then Exit Function
        If StrComp(CStr(selections(selectionIndex)), "Local", vbBinaryCompare) <> 0 And StrComp(CStr(selections(selectionIndex)), "Remote", vbBinaryCompare) <> 0 Then Exit Function
        result(rowIndex - rowLow + 1, 1) = CStr(conflicts(rowIndex, columnLow))
        result(rowIndex - rowLow + 1, 2) = conflicts(rowIndex, columnLow + 1)
        result(rowIndex - rowLow + 1, 3) = conflicts(rowIndex, columnLow + 2)
        result(rowIndex - rowLow + 1, 4) = CStr(selections(selectionIndex))
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
        FormatConflictValue = "（空白）"
        Exit Function
    End If
    If IsNull(value) Then
        FormatConflictValue = "（无效值）"
        Exit Function
    End If
    If IsError(value) Then
        FormatConflictValue = "（错误值）"
        Exit Function
    End If
    If VarType(value) <> vbString Then
        FormatConflictValue = "（非字符串值）"
        Exit Function
    End If

    formatted = Replace$(CStr(value), vbCrLf, vbLf)
    formatted = Replace$(formatted, vbCr, vbLf)
    FormatConflictValue = formatted
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

    Set form = CreateObject("Forms.UserForm.1")
    With form
        .Caption = "选择要签出的 AMG 列"
        .Width = TARGET_FORM_WIDTH
        .Height = TARGET_FORM_HEIGHT
        .StartUpPosition = 1
    End With
    Set heading = form.Controls.Add("Forms.Label.1", "lblTargetHeading", True)
    With heading
        .Caption = "工作表 / AMG 列 / 锁状态"
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
        .Caption = "开始编辑"
        .Left = 424
        .Top = 366
        .Width = 105
        .Height = 28
        .Default = True
    End With
    Set cancelButton = form.Controls.Add("Forms.CommandButton.1", "cmdTargetCancel", True)
    With cancelButton
        .Caption = "取消"
        .Left = 544
        .Top = 366
        .Width = 105
        .Height = 28
        .Cancel = True
    End With
    Set handlers = New CTargetSelectorEvents
    handlers.Bind form, listBox, acceptButton, cancelButton
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
    MsgBox "目标选择窗体无法打开。没有开始签出。", vbExclamation, "无法选择目标"

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
    Set form = CreateObject("Forms.UserForm.1")
    With form
        .Caption = "逐格裁决冲突"
        .Width = CONFLICT_FORM_WIDTH
        .Height = CONFLICT_FORM_HEIGHT
        .StartUpPosition = 1
    End With
    Set locationLabel = AddLabel(form, "lblConflictLocation", 12, 10, 620, 24, vbNullString)
    Set countLabel = AddLabel(form, "lblConflictCount", 12, 34, 620, 20, vbNullString)
    Set baseValue = AddValueBox(form, "txtBaseValue", "Base", 12, 62)
    Set localValue = AddValueBox(form, "txtLocalValue", "Local", 12, 160)
    Set remoteValue = AddValueBox(form, "txtRemoteValue", "Remote", 12, 258)

    Set previousButton = AddButton(form, "cmdPrevious", "上一项", 12, 384, 92)
    Set nextButton = AddButton(form, "cmdNext", "下一项", 116, 384, 92)
    Set localButton = AddButton(form, "cmdChooseLocal", "保留 Local", 224, 384, 130)
    Set remoteButton = AddButton(form, "cmdChooseRemote", "保留 Remote", 370, 384, 130)
    Set finishButton = AddButton(form, "cmdFinish", "完成裁决", 516, 384, 130)
    finishButton.Default = True
    Set cancelButton = AddButton(form, "cmdCancel", "取消", 516, 422, 130)
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
    MsgBox "冲突裁决窗体无法打开。共享文件尚未写入，本地会话仍保留。", vbExclamation, "无法裁决冲突"

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
