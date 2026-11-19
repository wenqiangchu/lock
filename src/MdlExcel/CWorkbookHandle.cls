VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CWorkbookHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWorkbook As Workbook
Private mOwned As Boolean
Private mRole As WorkbookRole

Friend Sub Initialize(ByRef workbook As Workbook, ByVal owned As Boolean, ByVal role As WorkbookRole)
    Set mWorkbook = workbook
    mOwned = owned
    mRole = role
End Sub

Friend Property Get WorkbookObject() As Workbook
    Set WorkbookObject = mWorkbook
End Property

Friend Property Get IsOwned() As Boolean
    IsOwned = mOwned
End Property

Friend Property Get Role() As WorkbookRole
    Role = mRole
End Property

Friend Sub Detach()
    Set mWorkbook = Nothing
End Sub
