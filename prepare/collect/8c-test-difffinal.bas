Attribute VB_Name = "ZZTestDiffFinal"
'==============================================================================
'  8c-test-difffinal.bas  --  the three questions that decide the design
'
'  Pure ASCII. Excel VBA7.
'
'  E1b : mixed-range property reads, printed unconditionally (no Select Case)
'  E6  : is StrComp on a multi-megabyte string faster than hashing it?
'  E7  : is Range.Value(11) DETERMINISTIC?  (same range, read twice, and
'        after a save/reopen)  -> decides whether it can be a fingerprint
'
'  Entry point: ZZRUNFINAL
'  Report     : %TEMP%\mddt_final.txt
'==============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" _
        (ByRef q As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" _
        (ByRef q As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" _
        (ByRef q As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" _
        (ByRef q As Currency) As Long
#End If

Private m_log As String
Private m_freq As Double
Private m_f As Currency

Public Sub ZZCOMPILEPROBE()
    Dim n As Long
    n = 1
    Debug.Print "compile ok " & n
End Sub

Public Sub ZZRUNFINAL()
    On Error GoTo Fail

    m_log = ""
    QueryPerformanceFrequency m_f
    m_freq = CDbl(m_f)

    Say "==================================================================="
    Say " FINAL QUESTIONS  --  Excel " & Application.Version & " build " & Application.Build
    Say "==================================================================="

    Dim a As Boolean, b As Boolean, c As Boolean
    a = Application.EnableEvents
    b = Application.ScreenUpdating
    c = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    TestMixed2
    TestStrComp
    TestValue11Determinism

    Application.EnableEvents = a
    Application.ScreenUpdating = b
    Application.DisplayAlerts = c

    Say ""
    Say "ALL DONE."
    WriteReport
    Exit Sub

Fail:
    Say "*** ABORTED: " & Err.Number & " " & Err.Description
    WriteReport
End Sub

'==============================================================================
' E1b -- mixed range properties, no conditional logic at all
'==============================================================================
Private Sub TestMixed2()
    Dim wb As Workbook, ws As Worksheet
    Dim r As Range
    Dim v As Variant

    Say ""
    Say "---- E1b: MIXED range reads (must not be trusted) --------------"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ws.Range("A1").Value2 = 111
    ws.Range("A2").Value2 = 222
    ws.Range("A1").Interior.Color = RGB(255, 0, 0)
    ws.Range("A2").Interior.Color = RGB(0, 0, 255)
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Font.Bold = False
    ws.Range("A1").Font.Color = RGB(0, 128, 0)
    ws.Range("A2").Font.Color = RGB(128, 0, 128)
    ws.Range("A1").NumberFormat = "0.00"
    ws.Range("A2").NumberFormat = "General"
    ws.Range("A1").HorizontalAlignment = -4131     ' xlLeft
    ws.Range("A2").HorizontalAlignment = -4152     ' xlRight
    ws.Range("A1").Borders(7).LineStyle = 1
    ws.Range("A2").Borders(7).LineStyle = -4142

    Set r = ws.Range("A1:A2")
    Say "   mixed range A1:A2  (A1 red/bold/0.00/left/border, A2 blue/plain)"

    v = r.Interior.Color
    Say "   Interior.Color        Null=" & IsNull(v) & "  val=" & Str(v)
    v = r.Interior.ColorIndex
    Say "   Interior.ColorIndex   Null=" & IsNull(v) & "  val=" & Str(v)
    v = r.Font.Bold
    Say "   Font.Bold             Null=" & IsNull(v) & "  val=" & Str(v)
    v = r.Font.Color
    Say "   Font.Color            Null=" & IsNull(v) & "  val=" & Str(v)
    v = r.NumberFormat
    Say "   NumberFormat          Null=" & IsNull(v) & "  val=" & Str(v)
    v = r.HorizontalAlignment
    Say "   HorizontalAlignment   Null=" & IsNull(v) & "  val=" & Str(v)
    v = r.Borders(7).LineStyle
    Say "   Borders(7).LineStyle  Null=" & IsNull(v) & "  val=" & Str(v)

    Say "   (single-cell A1 for reference) Color=" & ws.Range("A1").Interior.Color & _
        "  Bold=" & ws.Range("A1").Font.Bold

    ' ---- and the same after making the range uniform -------------------
    ws.Range("A1:A2").Interior.Color = RGB(9, 9, 9)
    Set r = ws.Range("A1:A2")
    v = r.Interior.Color
    Say "   UNIFORM Interior.Color  Null=" & IsNull(v) & "  val=" & Str(v) & _
        "   (expected 591366 = RGB(9,9,9))"

    wb.Close False
End Sub

Private Function Str(ByVal v As Variant) As String
    If IsNull(v) Then
        Str = "<NULL>"
    ElseIf IsEmpty(v) Then
        Str = "<EMPTY>"
    Else
        Str = CStr(v)
    End If
End Function

'==============================================================================
' E6 -- StrComp vs hash on a large string
'==============================================================================
Private Sub TestStrComp()
    Dim s1 As String, s2 As String
    Dim n As Long, i As Long
    Dim t As Double
    Dim r As Long
    Dim rep As Long

    Say ""
    Say "---- E6: comparing vs hashing a large string -------------------"

    n = 4431693          ' same size as the 60k-cell XML measured in T9
    s1 = String$(n, "A")
    s2 = s1

    t = NowSec()
    r = StrComp(s1, s2, vbBinaryCompare)
    Say " StrComp 4.2M chars (equal)   : " & Fmt(NowSec() - t) & "  result=" & r

    s2 = s2 & "B"
    t = NowSec()
    r = StrComp(s1, s2, vbBinaryCompare)
    Say " StrComp 4.2M chars (differ, same length prefix) : " & _
        Fmt(NowSec() - t) & "  result=" & r

    Mid$(s2, 1, 1) = "B"
    t = NowSec()
    r = StrComp(s1, s2, vbBinaryCompare)
    Say " StrComp 4.2M chars (differ at byte 1) : " & Fmt(NowSec() - t) & "  result=" & r

    ' hash for comparison
    t = NowSec()
    r = CLng(FnvBytes(s1, 200000))
    Say " hash first 200k chars (byte) : " & Fmt(NowSec() - t) & _
        " -> extrapolated " & Fmt((NowSec() - t) * n / 200000) & " for 4.2M"

    ' how expensive is a full byte conversion?
    t = NowSec()
    Dim bb() As Byte
    bb = StrConv(s1, vbFromUnicode)
    Say " StrConv 4.2M -> bytes        : " & Fmt(NowSec() - t) & "  ubound=" & UBound(bb)
End Sub

'==============================================================================
' E7 -- is Value(11) deterministic?
'==============================================================================
Private Sub TestValue11Determinism()
    Dim wb As Workbook, ws As Worksheet
    Dim r As Range
    Dim s1 As String, s2 As String, s3 As String, s4 As String
    Dim p As String
    Dim t As Double

    Say ""
    Say "---- E7: is Range.Value(11) DETERMINISTIC? ---------------------"

    p = Environ$("TEMP") & "\mddt_e7_probe.xlsx"
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    On Error GoTo 0

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ws.Range("A1").Value2 = "text"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Interior.Color = RGB(200, 100, 50)
    ws.Range("B2").Value2 = 3.14159
    ws.Range("B2").NumberFormat = "0.000"
    ws.Range("C3").Formula = "=1+2"

    Set r = ws.Range("A1:C3")

    t = NowSec()
    s1 = CStr(r.Value(11))
    Say " read #1                : " & Fmt(NowSec() - t) & "  len=" & Len(s1)

    t = NowSec()
    s2 = CStr(r.Value(11))
    Say " read #2 (same object)  : " & Fmt(NowSec() - t) & "  len=" & Len(s2)

    Say " #1 == #2 ?             : " & (StrComp(s1, s2, vbBinaryCompare) = 0)

    ' write to a file, save, close, reopen
    wb.SaveAs p, 51
    wb.Close False
    Set wb = Application.Workbooks.Open(p, 0, False)
    Set ws = wb.Worksheets(1)
    Set r = ws.Range("A1:C3")
    s3 = CStr(r.Value(11))
    Say " after save+reopen      : len=" & Len(s3) & _
        "   same as #1 ? " & (StrComp(s1, s3, vbBinaryCompare) = 0)

    ' modify then undo-equivalent: recolor back to identical value
    ws.Range("A1").Interior.Color = RGB(1, 2, 3)
    ws.Range("A1").Interior.Color = RGB(200, 100, 50)
    s4 = CStr(r.Value(11))
    Say " after color round-trip : len=" & Len(s4) & _
        "   same as #1 ? " & (StrComp(s1, s4, vbBinaryCompare) = 0)

    ' where do they diverge, if anywhere?
    If StrComp(s1, s3, vbBinaryCompare) <> 0 Then
        Say " DIFF vs s3:"
        ShowDiff s1, s3
    End If
    If StrComp(s1, s4, vbBinaryCompare) <> 0 Then
        Say " DIFF vs s4:"
        ShowDiff s1, s4
    End If

    wb.Close False
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    On Error GoTo 0
End Sub

Private Sub ShowDiff(ByVal a As String, ByVal b As String)
    Dim n As Long, i As Long
    n = Len(a)
    If Len(b) < n Then n = Len(b)
    For i = 1 To n
        If Mid$(a, i, 1) <> Mid$(b, i, 1) Then
            Say "   first difference at char " & i
            Say "   A: [" & Mid$(a, i, 80) & "]"
            Say "   B: [" & Mid$(b, i, 80) & "]"
            Exit Sub
        End If
    Next i
    Say "   identical up to " & n & " chars, lengths " & Len(a) & " vs " & Len(b)
End Sub

'==============================================================================
' Hashing (byte based, for the comparison above)
'==============================================================================
Private Function FnvBytes(ByVal s As String, ByVal n As Long) As Long
    Dim b() As Byte, i As Long
    Dim h As Long
    Dim part As String

    If n > Len(s) Then n = Len(s)
    part = Left$(s, n)
    b = StrConv(part, vbFromUnicode)

    h = &H811C9DC5
    For i = 0 To UBound(b)
        h = FnvStep(h, CLng(b(i)))
    Next i
    FnvBytes = h
End Function

Private Function FnvStep(ByVal h As Long, ByVal c As Long) As Long
    Const P As Long = 16777619
    Dim u As Double, hi As Double, lo As Double, r As Double

    h = h Xor (c And &HFF&)
    If h < 0 Then
        u = CDbl(h) + 4294967296#
    Else
        u = CDbl(h)
    End If

    hi = Int(u / 65536#)
    lo = u - hi * 65536#

    r = (hi * P - Int(hi * P / 65536#) * 65536#) * 65536#
    r = r + (lo * P - Int(lo * P / 4294967296#) * 4294967296#)
    r = r - Int(r / 4294967296#) * 4294967296#

    If r > 2147483647# Then r = r - 4294967296#
    FnvStep = CLng(r)
End Function

'==============================================================================
' Utility
'==============================================================================
Private Function NowSec() As Double
    Dim cc As Currency
    QueryPerformanceCounter cc
    NowSec = CDbl(cc) / m_freq
End Function

Private Function Fmt(ByVal s As Double) As String
    Fmt = Format$(s, "0.000") & " s"
End Function

Private Sub Say(ByVal s As String)
    m_log = m_log & s & vbCrLf
    Debug.Print s
End Sub

Private Sub WriteReport()
    Dim p As String, f As Integer
    p = Environ$("TEMP") & "\mddt_final.txt"
    On Error Resume Next
    f = FreeFile
    Open p For Output As #f
    Print #f, m_log
    Close #f
    On Error GoTo 0
    Debug.Print "report: " & p
End Sub