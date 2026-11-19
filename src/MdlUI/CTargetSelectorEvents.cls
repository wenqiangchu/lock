VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CTargetSelectorEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mAcceptButton As MSForms.CommandButton
Private WithEvents mCancelButton As MSForms.CommandButton
Private mForm As Object
Private mListBox As MSForms.ListBox
Private mAccepted As Boolean

Public Sub Bind(ByVal formObject As Object, ByVal listBox As Object, ByVal acceptButton As Object, ByVal cancelButton As Object)
    Set mForm = formObject
    Set mListBox = listBox
    Set mAcceptButton = acceptButton
    Set mCancelButton = cancelButton
    mAccepted = False
End Sub

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Private Sub mAcceptButton_Click()
    Dim itemIndex As Long
    Dim anySelected As Boolean

    For itemIndex = 0 To mListBox.ListCount - 1
        If mListBox.Selected(itemIndex) Then
            anySelected = True
            Exit For
        End If
    Next itemIndex
    If Not anySelected Then
        MsgBox "请至少选择一个 AMG 列。", vbInformation, "未选择目标"
        Exit Sub
    End If
    mAccepted = True
    mForm.Hide
End Sub

Private Sub mCancelButton_Click()
    mAccepted = False
    mForm.Hide
End Sub
