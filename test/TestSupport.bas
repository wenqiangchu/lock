Attribute VB_Name = "TestSupport"
Option Explicit

Private mLastFailure As String

Public Function RunSelectedTest(ByVal macroName As String) As Boolean
    Dim qualifiedMacro As String

    On Error GoTo Failed
    mLastFailure = vbNullString
    qualifiedMacro = "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!" & macroName
    Application.Run qualifiedMacro
    RunSelectedTest = (Len(mLastFailure) = 0)
    Exit Function

Failed:
    mLastFailure = Err.Source & ": " & Err.Description & " (" & CStr(Err.Number) & ")"
End Function

Public Function GetLastTestFailure() As String
    GetLastTestFailure = mLastFailure
End Function

Public Sub RecordTestFailure(ByVal message As String)
    If Len(mLastFailure) = 0 Then mLastFailure = message
End Sub
