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

Public Function TestFontSignature(ByVal isBold As Boolean, ByVal colorValue As Long) As String
    Dim properties(1 To FONT_FORMAT_PROPERTY_COUNT) As Variant
    Dim errorInfo As AppErrorInfo

    properties(1) = "Arial"
    properties(2) = 10
    properties(3) = isBold
    properties(4) = False
    properties(5) = 0
    properties(6) = False
    properties(7) = False
    properties(8) = False
    properties(9) = False
    properties(10) = False
    properties(11) = -4105
    properties(12) = colorValue
    properties(13) = CVErr(2042)
    properties(14) = CVErr(2042)
    TestFontSignature = BuildFontSignature(properties, errorInfo)
End Function

Public Function TestTextFontData(ByVal cellValue As Variant, ByVal fontSignature As String) As String
    Dim runs As Variant
    Dim oneRun(1 To 1, 1 To 3) As Variant
    Dim errorInfo As AppErrorInfo

    If VarType(cellValue) = vbString And Len(CStr(cellValue)) > 0 Then
        oneRun(1, 1) = 1
        oneRun(1, 2) = Len(CStr(cellValue))
        oneRun(1, 3) = fontSignature
        runs = oneRun
    Else
        runs = Empty
    End If
    TestTextFontData = BuildTextFontData(cellValue, runs, fontSignature, errorInfo)
End Function

Public Function TestFillData(ByVal colorValue As Long) As String
    Dim properties(1 To FILL_FORMAT_PROPERTY_COUNT) As Variant
    Dim errorInfo As AppErrorInfo

    properties(1) = 1
    properties(2) = -4105
    properties(3) = colorValue
    properties(4) = CVErr(2042)
    properties(5) = CVErr(2042)
    properties(6) = -4105
    properties(7) = 0
    properties(8) = CVErr(2042)
    properties(9) = CVErr(2042)
    TestFillData = BuildFillData(properties, errorInfo)
End Function
