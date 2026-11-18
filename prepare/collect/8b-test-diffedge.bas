Attribute VB_Name = "ZZTestDiffEdge"
'==============================================================================
'  8b-test-diffedge.bas  --  edge cases that decide the merge implementation
'
'  Pure ASCII. Excel VBA7.
'
'  E1  : does a mixed range return Null for Interior.Color / Font.Bold /
'        NumberFormat / Borders?  (the "uniform fast path" depends on it)
'  E2  : non-contiguous Union - what do .Value2 / .Formula / .Copy give?
'  E3  : does Value(11) cover merged cells, hyperlinks, comments?
'  E4  : realistic shape - 30 scattered study columns x 5000 rows,
'        read column-by-column and write back column-by-column
'  E5  : Range.Copy between two disjoint column sets (write-back path)
'
'  Entry point: ZZRUNEDGE
'  Report     : %TEMP%\mddt_edge.txt
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
Private r3 As Range

Public Sub ZZCOMPILEPROBE()
    ' Touching this sub forces the whole module to compile.
    Dim n As Long
    n = 1
    Debug.Print "compile ok " & n
End Sub

Public Sub ZZRUNEDGE()
    On Error GoTo Fail

    m_log = ""
    QueryPerformanceFrequency m_f
    m_freq = CDbl(m_f)

    Say "==================================================================="
    Say " EDGE CASES  --  Excel " & Application.Version & "  build " & Application.Build
    Say "==================================================================="

    Dim e1 As Boolean, e2 As Boolean, e3 As Boolean
    e1 = Application.EnableEvents
    e2 = Application.ScreenUpdating
    e3 = Application.DisplayAlerts

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    TestMixed
    TestUnion
    TestValue11Extras
    TestRealisticShape
    TestCopyDisjoint

    Application.EnableEvents = e1
    Application.ScreenUpdating = e2
    Application.DisplayAlerts = e3

    Say ""
    Say "ALL DONE."
    WriteReport
    Exit Sub

Fail:
    Say "*** ABORTED: " & Err.Number & " " & Err.Description
    WriteReport
End Sub

'==============================================================================
' E1 -- mixed-range property behaviour
'==============================================================================
Private Sub TestMixed()
    Dim wb As Workbook, ws As Worksheet
    Dim r As Range
    Dim colA As Long, colB As Long

    Say ""
    Say "---- E1: does a MIXED range return Null? -----------------------"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ws.Range("A1").Value2 = "one"
    ws.Range("A2").Value2 = "two"

    ' make A1 and A2 differ in each property
    ws.Range("A1").Interior.Color = RGB(255, 0, 0)      ' red
    ws.Range("A2").Interior.Color = RGB(0, 0, 255)      ' blue
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Font.Bold = False
    ws.Range("A1").NumberFormat = "0.00"
    ws.Range("A2").NumberFormat = "General"
    ws.Range("A1").Borders(7).LineStyle = 1             ' xlContinuous
    ws.Range("A2").Borders(7).LineStyle = -4142         ' xlNone

    Set r = ws.Range("A1:A2")

    Report "Interior.Color        ", r
    Report "Interior.ColorIndex   ", r
    Report "Font.Bold             ", r
    Report "NumberFormat          ", r
    Report "Borders(7).LineStyle  ", r

    ' also the same properties on each cell, for comparison
    Say "   (per cell: A1 Color=" & ws.Range("A1").Interior.Color & _
        " Bold=" & ws.Range("A1").Font.Bold & _
        " Fmt=" & ws.Range("A1").NumberFormat & ")"
    Say "   (per cell: A2 Color=" & ws.Range("A2").Interior.Color & _
        " Bold=" & ws.Range("A2").Font.Bold & _
        " Fmt=" & ws.Range("A2").NumberFormat & ")"

    ' uniform range must still report a value
    ws.Range("A1:A2").Interior.Color = RGB(0, 255, 0)
    Set r = ws.Range("A1:A2")
    Report "UNIFORM Interior.Color", r

    wb.Close False
End Sub

Private Sub Report(ByVal label As String, ByVal r As Range)
    Dim v As Variant
    On Error Resume Next
    Select Case True
        Case label Like "Interior.Color*"
            If label Like "*ColorIndex*" Then
                v = r.Interior.ColorIndex
            Else
                v = r.Interior.Color
            End If
        Case label Like "Font.Bold*"
            v = r.Font.Bold
        Case label Like "NumberFormat*"
            v = r.NumberFormat
        Case label Like "Borders*"
            v = r.Borders(7).LineStyle
    End Select
    If Err.Number <> 0 Then
        Say "  " & label & " : ERROR " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        Say "  " & label & " : Null? " & IsNull(v) & _
            "   value=" & IIf(IsNull(v), "<null>", CStr(v))
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' E2 -- non-contiguous Union semantics
'==============================================================================
Private Sub TestUnion()
    Dim wb As Workbook, ws As Worksheet
    Dim u As Range, col As Range
    Dim pick As Variant, k As Long, i As Long
    Dim v As Variant, n As Long
    Dim t As Double

    Say ""
    Say "---- E2: non-contiguous Union semantics ------------------------"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    For i = 1 To 100
        ws.Cells(i, 2).Value2 = "B" & i
        ws.Cells(i, 5).Value2 = "E" & i
        ws.Cells(i, 7).Value2 = "G" & i
    Next i

    pick = Array(2, 5, 7)

    Set u = Nothing
    For k = 0 To UBound(pick)
        Set col = ws.Range(ws.Cells(1, pick(k)), ws.Cells(100, pick(k)))
        If u Is Nothing Then Set u = col Else Set u = Union(u, col)
    Next k

    Say " Union areas      : " & u.Areas.Count
    Say " Union Cells.Count: " & u.Cells.Count
    Say " Union address    : " & u.Address(0, 0)

    t = NowSec()
    v = u.Formula
    Say " u.Formula        : " & Fmt(NowSec() - t)
    If IsArray(v) Then
        Say "   UBound1=" & UBound(v, 1) & " UBound2=" & UBound(v, 2) & _
            "   v(1,1)=" & CStr(v(1, 1))
        n = 0
        If UBound(v, 2) >= 2 Then n = UBound(v, 2)
        If n >= 1 Then Say "   v(1,last col)=" & CStr(v(1, n))
    Else
        Say "   not an array: " & CStr(v)
    End If

    t = NowSec()
    v = u.Value2
    Say " u.Value2         : " & Fmt(NowSec() - t)
    If IsArray(v) Then
        Say "   UBound1=" & UBound(v, 1) & " UBound2=" & UBound(v, 2)
    Else
        Say "   not an array: " & CStr(v)
    End If

    ' per-area read
    Say " per-area loop:"
    t = NowSec()
    For k = 1 To u.Areas.Count
        v = u.Areas(k).Formula
        Say "   area " & k & " -> " & UBound(v, 1) & "x" & UBound(v, 2) & _
            "  first=" & CStr(v(1, 1))
    Next k
    Say "   total " & Fmt(NowSec() - t)

    wb.Close False
End Sub

'==============================================================================
' E3 -- does Value(11) cover merged cells / hyperlinks / comments?
'==============================================================================
Private Sub TestValue11Extras()
    Dim wb As Workbook, ws As Worksheet
    Dim s As String

    Say ""
    Say "---- E3: Value(11) coverage of merged / link / comment ----------"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ws.Range("A1:B1").Merge
    ws.Range("A1").Value2 = "merged"

    ws.Hyperlinks.Add Anchor:=ws.Range("A3"), _
        Address:="https://example.com/x", TextToDisplay:="click"

    On Error Resume Next
    ws.Range("A5").AddComment "a note"
    On Error GoTo 0

    Set r3 = ws.Range("A1:B5")
    s = CStr(r3.Value(11))
    Say " length: " & Len(s)
    Say " contains 'MergeAcross' : " & (InStr(1, s, "MergeAcross", vbTextCompare) > 0)
    Say " contains 'MergeDown'   : " & (InStr(1, s, "MergeDown", vbTextCompare) > 0)
    Say " contains 'HRef'        : " & (InStr(1, s, "HRef", vbTextCompare) > 0)
    Say " contains 'Comment'     : " & (InStr(1, s, "Comment", vbTextCompare) > 0)
    Say " contains 'a note'      : " & (InStr(1, s, "a note", vbTextCompare) > 0)
    Say " contains 'example.com' : " & (InStr(1, s, "example.com", vbTextCompare) > 0)

    ' merged cell: what does Formula give?
    Say " A1 Formula: [" & CStr(ws.Range("A1").Formula) & "]"
    Say " B1 Formula: [" & CStr(ws.Range("B1").Formula) & "]  (merged partner)"

    wb.Close False
End Sub

'==============================================================================
' E4 -- realistic shape: 30 scattered columns x 5000 rows
'==============================================================================
Private Sub TestRealisticShape()
    Dim wb As Workbook, ws As Worksheet, wd As Worksheet
    Dim rows As Long, cols As Long
    Dim a As Variant, v As Variant, outArr As Variant
    Dim pickCols() As Long
    Dim i As Long, j As Long, k As Long
    Dim t As Double
    Dim sheetCol As Long

    Say ""
    Say "---- E4: realistic shape, 30 scattered study columns -----------"

    rows = 5000
    cols = 30

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set wd = wb.Worksheets.Add
    wd.Name = "dst"

    ' columns 1..90, study columns are every 3rd
    ReDim a(1 To rows, 1 To 90)
    For i = 1 To rows
        For j = 1 To 90
            a(i, j) = "R" & i & "C" & j
        Next j
    Next i
    t = NowSec()
    ws.Range(ws.Cells(1, 1), ws.Cells(rows, 90)).Value2 = a
    Say " seed 5000x90 grid      : " & Fmt(NowSec() - t)

    ReDim pickCols(1 To cols)
    For k = 1 To cols
        pickCols(k) = (k - 1) * 3 + 1
    Next k

    ' ---- read: 30 separate column reads --------------------------------
    t = NowSec()
    ReDim outArr(1 To rows, 1 To cols)
    For k = 1 To cols
        sheetCol = pickCols(k)
        v = ws.Range(ws.Cells(1, sheetCol), ws.Cells(rows, sheetCol)).Formula
        For i = 1 To rows
            outArr(i, k) = v(i, 1)
        Next i
    Next k
    Say " read 30 cols 1-by-1    : " & Fmt(NowSec() - t) & _
        "   (" & Format$(rows * cols, "#,##0") & " cells)"

    ' ---- read: whole bounding block then pick -------------------------
    t = NowSec()
    v = ws.Range(ws.Cells(1, 1), ws.Cells(rows, 90)).Formula
    Say " read bounding 5000x90  : " & Fmt(NowSec() - t) & _
        "   (" & Format$(rows * 90, "#,##0") & " cells)"

    ' ---- diff two 5000x30 arrays --------------------------------------
    Dim outArr2 As Variant
    ReDim outArr2(1 To rows, 1 To cols)
    For i = 1 To rows
        For k = 1 To cols
            outArr2(i, k) = outArr(i, k)
        Next k
    Next i
    outArr2(2500, 15) = "CHANGED"

    t = NowSec()
    Dim diffs As Long
    diffs = 0
    For i = 1 To rows
        For k = 1 To cols
            If CStr(outArr(i, k)) <> CStr(outArr2(i, k)) Then
                diffs = diffs + 1
                If diffs > 100000 Then Exit For
            End If
        Next k
    Next i
    Say " diff 5000x30 arrays    : " & Fmt(NowSec() - t) & "   diffs=" & diffs

    ' ---- write back: 30 separate column copy --------------------------
    t = NowSec()
    For k = 1 To cols
        sheetCol = pickCols(k)
        ws.Range(ws.Cells(1, sheetCol), ws.Cells(rows, sheetCol)).Copy _
            Destination:=wd.Range(wd.Cells(1, sheetCol), wd.Cells(rows, sheetCol))
    Next k
    Say " copy 30 cols 1-by-1    : " & Fmt(NowSec() - t)

    wb.Close False
End Sub

'==============================================================================
' E5 -- Copy between two disjoint column sets
'==============================================================================
Private Sub TestCopyDisjoint()
    Dim wb As Workbook, ws As Worksheet, wd As Worksheet
    Dim i As Long, k As Long
    Dim t As Double
    Dim srcSet As Range, dstSet As Range, c As Range

    Say ""
    Say "---- E5: Range.Copy from a multi-area source -------------------"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set wd = wb.Worksheets.Add
    wd.Name = "dst"

    For i = 1 To 50
        ws.Cells(i, 2).Value2 = "B" & i
        ws.Cells(i, 5).Value2 = "E" & i
        ws.Cells(i, 7).Value2 = "G" & i
    Next i

    Dim pk As Variant
    pk = Array(2, 5, 7)

    Set srcSet = Nothing
    For k = 0 To UBound(pk)
        Set c = ws.Range(ws.Cells(1, pk(k)), ws.Cells(50, pk(k)))
        If srcSet Is Nothing Then
            Set srcSet = c
        Else
            Set srcSet = Union(srcSet, c)
        End If
    Next k

    t = NowSec()
    On Error Resume Next
    srcSet.Copy Destination:=wd.Range("B1")
    Say " multi-area Copy -> B1  : " & Fmt(NowSec() - t)
    If Err.Number <> 0 Then
        Say "   FAILED " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        Say "   dst B1=" & CStr(wd.Range("B1").Value2) & _
            "  E1=" & CStr(wd.Range("E1").Value2) & _
            "  G1=" & CStr(wd.Range("G1").Value2)
    End If
    On Error GoTo 0

    wb.Close False
End Sub

'==============================================================================
' Utility
'==============================================================================
Private Function NowSec() As Double
    Dim c As Currency
    QueryPerformanceCounter c
    NowSec = CDbl(c) / m_freq
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
    p = Environ$("TEMP") & "\mddt_edge.txt"
    On Error Resume Next
    f = FreeFile
    Open p For Output As #f
    Print #f, m_log
    Close #f
    On Error GoTo 0
    Debug.Print "report: " & p
End Sub