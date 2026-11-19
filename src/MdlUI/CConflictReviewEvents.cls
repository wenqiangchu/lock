VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CConflictReviewEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mPreviousButton As MSForms.CommandButton
Private WithEvents mNextButton As MSForms.CommandButton
Private WithEvents mLocalButton As MSForms.CommandButton
Private WithEvents mRemoteButton As MSForms.CommandButton
Private WithEvents mFinishButton As MSForms.CommandButton
Private WithEvents mCancelButton As MSForms.CommandButton
Private mForm As Object
Private mLocationLabel As MSForms.Label
Private mCountLabel As MSForms.Label
Private mBaseValue As MSForms.TextBox
Private mLocalValue As MSForms.TextBox
Private mRemoteValue As MSForms.TextBox
Private mConflicts As Variant
Private mRowLow As Long
Private mRowHigh As Long
Private mCurrentRow As Long
Private mSelections() As String
Private mConfirmed As Boolean

Public Sub Bind(ByVal formObject As Object, ByRef conflicts As Variant, ByVal locationLabel As Object, ByVal countLabel As Object, ByVal baseValue As Object, ByVal localValue As Object, ByVal remoteValue As Object, ByVal previousButton As Object, ByVal nextButton As Object, ByVal localButton As Object, ByVal remoteButton As Object, ByVal finishButton As Object, ByVal cancelButton As Object)
    If Not ReadConflictBounds(conflicts, mRowLow, mRowHigh) Then Err.Raise vbObjectError + 2801, "CConflictReviewEvents.Bind", "A component conflict matrix with seven columns is required."
    mConflicts = conflicts
    mCurrentRow = mRowLow
    ReDim mSelections(1 To mRowHigh - mRowLow + 1)
    Set mForm = formObject
    Set mLocationLabel = locationLabel
    Set mCountLabel = countLabel
    Set mBaseValue = baseValue
    Set mLocalValue = localValue
    Set mRemoteValue = remoteValue
    Set mPreviousButton = previousButton
    Set mNextButton = nextButton
    Set mLocalButton = localButton
    Set mRemoteButton = remoteButton
    Set mFinishButton = finishButton
    Set mCancelButton = cancelButton
    mConfirmed = False
    RenderCurrent
End Sub

Public Property Get Confirmed() As Boolean
    Confirmed = mConfirmed
End Property

Public Property Get ChoiceMatrix() As Variant
    Dim selections() As Variant
    Dim rowIndex As Long
    Dim selectionIndex As Long
    Dim result As Variant

    ReDim selections(1 To UBound(mSelections) - LBound(mSelections) + 1)
    For rowIndex = mRowLow To mRowHigh
        selectionIndex = rowIndex - mRowLow + 1
        selections(selectionIndex) = mSelections(selectionIndex)
    Next rowIndex
    If Not BuildConflictChoiceMatrix(mConflicts, selections, result) Then Exit Property
    ChoiceMatrix = result
End Property

Private Sub mPreviousButton_Click()
    If mCurrentRow > mRowLow Then
        mCurrentRow = mCurrentRow - 1
        RenderCurrent
    End If
End Sub

Private Sub mNextButton_Click()
    If mCurrentRow < mRowHigh Then
        mCurrentRow = mCurrentRow + 1
        RenderCurrent
    End If
End Sub

Private Sub mLocalButton_Click()
    SelectChoice "Local"
End Sub

Private Sub mRemoteButton_Click()
    SelectChoice "Remote"
End Sub

Private Sub mFinishButton_Click()
    If Not AllChoicesMade Then
        MsgBox "Choose Local or Remote for every conflicting cell.", vbInformation, "Unresolved Conflicts"
        Exit Sub
    End If
    mConfirmed = True
    mForm.Hide
End Sub

Private Sub mCancelButton_Click()
    mConfirmed = False
    mForm.Hide
End Sub

Private Sub SelectChoice(ByVal choice As String)
    mSelections(mCurrentRow - mRowLow + 1) = choice
    If mCurrentRow < mRowHigh Then mCurrentRow = mCurrentRow + 1
    RenderCurrent
End Sub

Private Sub RenderCurrent()
    Dim selectionIndex As Long
    Dim rowValue As Long
    Dim columnValue As Long
    Dim componentName As String

    selectionIndex = mCurrentRow - mRowLow + 1
    rowValue = CLng(mConflicts(mCurrentRow, 2))
    columnValue = CLng(mConflicts(mCurrentRow, 3))
    componentName = CStr(mConflicts(mCurrentRow, 4))
    mLocationLabel.Caption = CStr(mConflicts(mCurrentRow, 1)) & "!" & ExcelColumnLabel(columnValue) & CStr(rowValue) & " - " & componentName
    mCountLabel.Caption = "Conflict " & CStr(selectionIndex) & " / " & CStr(mRowHigh - mRowLow + 1) & "; resolved " & CStr(SelectedCount())
    mBaseValue.Text = FormatConflictComponent(componentName, mConflicts(mCurrentRow, 5))
    mLocalValue.Text = FormatConflictComponent(componentName, mConflicts(mCurrentRow, 6))
    mRemoteValue.Text = FormatConflictComponent(componentName, mConflicts(mCurrentRow, 7))
    mPreviousButton.Enabled = (mCurrentRow > mRowLow)
    mNextButton.Enabled = (mCurrentRow < mRowHigh)
    mLocalButton.Caption = ChoiceCaption("Keep Local", mSelections(selectionIndex) = "Local")
    mRemoteButton.Caption = ChoiceCaption("Keep Remote", mSelections(selectionIndex) = "Remote")
    mFinishButton.Enabled = AllChoicesMade
End Sub

Private Function ChoiceCaption(ByVal caption As String, ByVal selected As Boolean) As String
    If selected Then
        ChoiceCaption = caption & " (selected)"
    Else
        ChoiceCaption = caption
    End If
End Function

Private Function SelectedCount() As Long
    Dim rowIndex As Long
    For rowIndex = LBound(mSelections) To UBound(mSelections)
        If Len(mSelections(rowIndex)) > 0 Then SelectedCount = SelectedCount + 1
    Next rowIndex
End Function

Private Function AllChoicesMade() As Boolean
    AllChoicesMade = (SelectedCount() = UBound(mSelections) - LBound(mSelections) + 1)
End Function

Private Function ReadConflictBounds(ByRef conflicts As Variant, ByRef rowLow As Long, ByRef rowHigh As Long) As Boolean
    Dim columnLow As Long
    Dim columnHigh As Long
    Dim probe As Long

    If Not IsArray(conflicts) Then Exit Function
    On Error Resume Next
    rowLow = LBound(conflicts, 1)
    rowHigh = UBound(conflicts, 1)
    columnLow = LBound(conflicts, 2)
    columnHigh = UBound(conflicts, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    probe = LBound(conflicts, 3)
    If Err.Number <> 9 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    ReadConflictBounds = (columnHigh - columnLow + 1 = 7 And rowHigh >= rowLow)
End Function
