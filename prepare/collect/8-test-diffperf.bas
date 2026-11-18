Attribute VB_Name = "ZZTestDiffPerf"
'==============================================================================
'  8-test-diffperf.bas   --  T8 / T9 / T15 / T16
'
'  Pure ASCII. Excel VBA7 (x86 and x64 safe).
'
'  T8  : how long does it really take to read and diff a study-column block?
'  T9  : is Range.Value(xlRangeValueXMLSpreadsheet) usable (content + speed)?
'  T15 : does Application.EnableEvents = False suppress a target workbook's
'        own Workbook_Open macro, and do we get our App state back intact?
'  T16 : what happens when the target file is already open?
'
'  Entry point: ZZRUNALL
'  Report     : %TEMP%\mddt_t89.txt  (and Desktop if writable)
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
Private m_freq0 As Currency
Private m_cells As Long
Private m_tmp As String

'==============================================================================
' Entry point
'==============================================================================
Public Sub ZZCOMPILEPROBE()
    Dim n As Long
    n = 1
    Debug.Print "compile ok " & n
End Sub

Public Sub ZZRUNALL()
    On Error GoTo Fail

    m_log = ""
    m_tmp = Environ$("TEMP")

    QueryPerformanceFrequency m_freq0
    m_freq = CDbl(m_freq0)

    Say "==================================================================="
    Say " T8 / T9 / T15 / T16"
    Say " Excel      : " & Application.Version & "  build " & Application.Build
    Say " OS         : " & Application.OperatingSystem
    Say " 64-bit     : " & Is64()
    Say " QPC freq   : " & Format$(m_freq, "0")
    Say " TEMP       : " & m_tmp
    Say "==================================================================="

    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim oldAlerts As Boolean
    Dim oldSecurity As Long

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldAlerts = Application.DisplayAlerts
    oldSecurity = Application.AutomationSecurity

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AutomationSecurity = 1          ' msoAutomationSecurityLow
    Application.Calculation = -4135             ' xlCalculationManual

    TestT9Content
    TestBlock 500, 10
    TestBlock 2000, 30
    TestBlock 5000, 30
    TestMultiArea 2000, 20
    TestT15Events
    TestT16AlreadyOpen

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    Application.AutomationSecurity = oldSecurity

    Say ""
    Say "ALL DONE."
    WriteReport
    Exit Sub

Fail:
    Say "*** ABORTED: " & Err.Number & " " & Err.Description
    WriteReport
End Sub

'==============================================================================
' T9 part 1 -- what does Value(11) actually contain?
'==============================================================================
Private Sub TestT9Content()
    Dim wb As Workbook, ws As Worksheet
    Dim rng As Range, s As String

    Say ""
    Say "---- T9a: what Value(11) contains ------------------------------"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ws.Range("A1").Value2 = "bold-fill-font"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(0, 0, 255)
    ws.Range("A1").Interior.Color = RGB(255, 255, 0)
    ws.Range("A1").NumberFormat = "0.00"

    ws.Range("B1").Value2 = 1234.5
    ws.Range("B1").Font.Italic = True
    ws.Range("C1").Value2 = "plain"

    Set rng = ws.Range("A1:C1")
    s = CStr(rng.Value(11))

    Say " Value(11) length for A1:C1 : " & Len(s)
    Say " contains 'Bold'      : " & (InStr(1, s, "Bold", vbTextCompare) > 0)
    Say " contains 'Italic'    : " & (InStr(1, s, "Italic", vbTextCompare) > 0)
    Say " contains 'Interior'  : " & (InStr(1, s, "Interior", vbTextCompare) > 0)
    Say " contains 'Font'      : " & (InStr(1, s, "Font", vbTextCompare) > 0)
    Say " contains 'NumberFormat': " & (InStr(1, s, "NumberFormat", vbTextCompare) > 0)
    Say " contains 'Border'    : " & (InStr(1, s, "Border", vbTextCompare) > 0)
    Say " contains 'Alignment' : " & (InStr(1, s, "Alignment", vbTextCompare) > 0)
    Say " contains 'MergeAcross': " & (InStr(1, s, "MergeAcross", vbTextCompare) > 0)
    Say " first 200 chars:"
    Say "   [" & Left$(Replace(s, vbCrLf, " "), 200) & "]"

    wb.Close False
End Sub

'==============================================================================
' T8 / T9 part 2 -- one block size
'==============================================================================
Private Sub TestBlock(ByVal rows As Long, ByVal cols As Long)
    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet
    Dim src As Range, dst As Range
    Dim t As Double, dt As Double
    Dim a As Variant, b As Variant
    Dim s As String
    Dim i As Long, j As Long, n As Long
    Dim diffs As Long
    Dim uniformColor As Variant, uniformBold As Variant
    Dim sampleN As Long
    Dim hashHex As String

    m_cells = rows * cols

    Say ""
    Say "---- T8/T9b: block " & rows & " rows x " & cols & " cols = " & _
        Format$(m_cells, "#,##0") & " cells ----"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set ws2 = wb.Worksheets.Add
    ws2.Name = "dst"

    ' ---- build synthetic data -------------------------------------------
    t = NowSec()
    ReDim a(1 To rows, 1 To cols)
    For i = 1 To rows
        For j = 1 To cols
            a(i, j) = "R" & CStr(i) & "C" & CStr(j)
        Next j
    Next i
    dt = NowSec() - t
    Say " build array            : " & Fmt(dt)

    Set src = ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols))
    Set dst = ws2.Range(ws2.Cells(1, 1), ws2.Cells(rows, cols))

    t = NowSec()
    src.Value2 = a
    dt = NowSec() - t
    Say " write array to sheet   : " & Fmt(dt)

    ' uniform style, to test the "one call tells us everything" case
    src.Interior.Color = RGB(255, 255, 204)
    src.Font.Color = RGB(0, 0, 0)
    src.Font.Bold = False

    ' ---- 1. read Formula array (1 COM call) ------------------------------
    t = NowSec()
    a = src.Formula
    dt = NowSec() - t
    Say " read .Formula array    : " & Fmt(dt)

    t = NowSec()
    b = src.Value2
    dt = NowSec() - t
    Say " read .Value2 array     : " & Fmt(dt)

    ' ---- 2. pure VBA array diff -----------------------------------------
    t = NowSec()
    diffs = 0
    For i = 1 To rows
        For j = 1 To cols
            If CStr(a(i, j)) <> CStr(b(i, j)) Then diffs = diffs + 1
        Next j
    Next i
    dt = NowSec() - t
    Say " diff two arrays        : " & Fmt(dt) & "   (diffs=" & diffs & ")"

    ' ---- 3. hash the Formula array --------------------------------------
    t = NowSec()
    hashHex = HashArray(a, rows, cols)
    dt = NowSec() - t
    Say " hash .Formula array    : " & Fmt(dt) & "   (" & hashHex & ")"

    ' ---- 4. per-cell style reads (sampled, then extrapolated) -----------
    sampleN = 5000
    If sampleN > m_cells Then sampleN = m_cells
    t = NowSec()
    uniformColor = 0
    For n = 1 To sampleN
        uniformColor = src.Cells(n, 1).Interior.Color
    Next n
    dt = NowSec() - t
    Say " per-cell Interior.Color: " & Fmt(dt) & " for " & sampleN & _
        " cells -> extrapolated " & Fmt(dt * m_cells / sampleN) & " for block"

    t = NowSec()
    uniformBold = 0
    For n = 1 To sampleN
        uniformBold = src.Cells(n, 1).Font.Bold
    Next n
    dt = NowSec() - t
    Say " per-cell Font.Bold     : " & Fmt(dt) & " for " & sampleN & _
        " cells -> extrapolated " & Fmt(dt * m_cells / sampleN) & " for block"

    ' ---- 5. whole-range style read (uniform) ----------------------------
    t = NowSec()
    uniformColor = src.Interior.Color
    dt = NowSec() - t
    Say " whole-range Interior.Color : " & Fmt(dt) & "   Null? " & IsNull(uniformColor)

    t = NowSec()
    uniformBold = src.Font.Bold
    dt = NowSec() - t
    Say " whole-range Font.Bold      : " & Fmt(dt) & "   Null? " & IsNull(uniformBold)

    ' ---- 6. whole-range style read (mixed) ------------------------------
    src.Cells(1, 1).Interior.Color = RGB(255, 0, 0)
    t = NowSec()
    uniformColor = src.Interior.Color
    dt = NowSec() - t
    Say " whole-range Interior.Color (mixed) : " & Fmt(dt) & "   Null? " & IsNull(uniformColor)
    src.Cells(1, 1).Interior.Color = RGB(255, 255, 204)

    ' ---- 7. Range.Copy write-back ---------------------------------------
    t = NowSec()
    src.Copy Destination:=dst
    dt = NowSec() - t
    Say " Range.Copy to other sheet : " & Fmt(dt)

    ' ---- 8. Value(11) XML ------------------------------------------------
    ' Skipped for very large blocks: the XML can reach tens of MB and is not
    ' representative. We measure it only where it is realistic.
    Dim prefixN As Long
    If m_cells <= 60000 Then
        On Error Resume Next
        t = NowSec()
        s = CStr(src.Value(11))
        dt = NowSec() - t
        If Err.Number <> 0 Then
            Say " read Value(11) XML     : FAILED " & Err.Number & " " & Err.Description
            Err.Clear
            s = ""
        Else
            Say " read Value(11) XML     : " & Fmt(dt) & "   length=" & _
                Format$(Len(s), "#,##0") & " chars (" & _
                Format$(Len(s) / 1048576#, "0.00") & " M chars)"
        End If
        On Error GoTo 0

        If Len(s) > 0 Then
            prefixN = 200000
            If prefixN > Len(s) Then prefixN = Len(s)

            t = NowSec()
            hashHex = HashStrMid(s, prefixN)
            dt = NowSec() - t
            Say " hash XML via Mid$      : " & Fmt(dt) & " for " & prefixN & _
                " chars -> extrapolated " & Fmt(dt * Len(s) / prefixN) & " full"

            t = NowSec()
            hashHex = HashStrBytes(s, prefixN)
            dt = NowSec() - t
            Say " hash XML via byte()    : " & Fmt(dt) & " for " & prefixN & _
                " chars -> extrapolated " & Fmt(dt * Len(s) / prefixN) & " full"
        End If
    Else
        Say " read Value(11) XML     : skipped (block too large for XML)"
    End If

    wb.Close False
    Set wb = Nothing
End Sub

'==============================================================================
' T15 -- does EnableEvents=False suppress the target's own Workbook_Open?
'==============================================================================
Private Sub TestT15Events()
    Dim wb As Workbook
    Dim p As String, mk As String
    Dim code As String

    Say ""
    Say "---- T15: EnableEvents vs target Workbook_Open ------------------"

    p = m_tmp & "\mddt_t15_probe.xlsm"
    mk = m_tmp & "\mddt_t15_marker.txt"
    KillQuiet mk
    KillQuiet p

    ' Build a macro workbook whose Workbook_Open writes a marker file.
    Set wb = Application.Workbooks.Add
    code = "Private Sub Workbook_Open()" & vbCrLf & _
           "    Dim f As Integer" & vbCrLf & _
           "    f = FreeFile" & vbCrLf & _
           "    Open """ & mk & """ For Output As #f" & vbCrLf & _
           "    Print #f, ""opened at "" & Now" & vbCrLf & _
           "    Close #f" & vbCrLf & _
           "End Sub"
    On Error Resume Next
    wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString code
    If Err.Number <> 0 Then
        Say " cannot inject code (VBProject access denied): " & Err.Description
        On Error GoTo 0
        wb.Close False
        Exit Sub
    End If
    On Error GoTo 0
    wb.SaveAs p, 52          ' xlOpenXMLWorkbookMacroEnabled
    wb.Close False
    Say " probe macro workbook created: " & p

    ' Case 1: EnableEvents = False -> marker must NOT appear
    KillQuiet mk
    Application.EnableEvents = False
    Set wb = Application.Workbooks.Open(p, 0, True)
    wb.Close False
    Set wb = Nothing
    Say " EnableEvents=False  -> marker exists? " & FileExists(mk) & _
        "   (expected: False)"

    ' Case 2: EnableEvents = True -> marker must appear (control)
    KillQuiet mk
    Application.EnableEvents = True
    Set wb = Application.Workbooks.Open(p, 0, True)
    wb.Close False
    Set wb = Nothing
    Say " EnableEvents=True   -> marker exists? " & FileExists(mk) & _
        "   (expected: True)"

    ' Case 3: does our Application state survive the Open?
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = -4135
    Set wb = Application.Workbooks.Open(p, 0, True)
    Say " state after Open    -> EnableEvents=" & Application.EnableEvents & _
        "  ScreenUpdating=" & Application.ScreenUpdating & _
        "  Calculation=" & Application.Calculation
    wb.Close False
    Set wb = Nothing
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = -4135

    KillQuiet mk
    KillQuiet p
End Sub

'==============================================================================
' T16 -- target file is already open
'==============================================================================
Private Sub TestT16AlreadyOpen()
    Dim wb1 As Workbook, wb2 As Workbook
    Dim p As String
    Dim t As Double, dt As Double

    Say ""
    Say "---- T16: opening a workbook that is already open ----------------"

    p = m_tmp & "\mddt_t16_probe.xlsx"
    KillQuiet p

    Set wb1 = Application.Workbooks.Add
    wb1.Worksheets(1).Range("A1").Value2 = "already-open-probe"
    wb1.SaveAs p, 51
    Set wb1 = Application.Workbooks.Open(p, 0, False)
    Say " first open : name=" & wb1.Name & "  windows=" & wb1.Windows.Count

    t = NowSec()
    Set wb2 = Application.Workbooks.Open(p, 0, False)
    dt = NowSec() - t

    Say " second Open: " & Fmt(dt)
    Say "   same object? " & (wb1 Is wb2)
    Say "   name=" & wb2.Name & "  windows=" & wb2.Windows.Count
    Say "   Workbooks.Count=" & Application.Workbooks.Count

    ' read-only attempt while it is open read-write
    t = NowSec()
    On Error Resume Next
    Dim wb3 As Workbook
    Set wb3 = Application.Workbooks.Open(p, 0, True)
    dt = NowSec() - t
    If Err.Number <> 0 Then
        Say " read-only open while r/w open: FAILED " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        Say " read-only open while r/w open: OK in " & Fmt(dt) & _
            "  same object as wb1? " & (wb3 Is wb1)
        wb3.Close False
    End If
    On Error GoTo 0

    On Error Resume Next
    wb2.Close False
    wb1.Close False
    On Error GoTo 0
    KillQuiet p
End Sub

'==============================================================================
' T8c -- non-contiguous study columns (the REAL shape)
'==============================================================================
Private Sub TestMultiArea(ByVal rows As Long, ByVal cols As Long)
    Dim wb As Workbook, ws As Worksheet
    Dim a As Variant, v As Variant, u As Range, col As Range
    Dim out() As Variant
    Dim i As Long, j As Long, k As Long
    Dim t As Double, dt As Double
    Dim pick As Variant
    Dim ub1 As Long, ub2 As Long
    Dim nulls As Long

    Say ""
    Say "---- T8c: non-contiguous columns (study cols are scattered) -----"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ReDim a(1 To rows, 1 To cols)
    For i = 1 To rows
        For j = 1 To cols
            a(i, j) = "R" & i & "C" & j
        Next j
    Next i
    ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols)).Value2 = a

    pick = Array(2, 5, 7, 11, 13, 17)

    t = NowSec()
    Set u = Nothing
    For k = 0 To UBound(pick)
        Set col = ws.Range(ws.Cells(1, pick(k)), ws.Cells(rows, pick(k)))
        If u Is Nothing Then
            Set u = col
        Else
            Set u = Union(u, col)
        End If
    Next k
    dt = NowSec() - t
    Say " build Union of 6 cols  : " & Fmt(dt) & "   areas=" & u.Areas.Count

    t = NowSec()
    On Error Resume Next
    v = u.Formula
    dt = NowSec() - t
    ub1 = -1: ub2 = -1
    ub1 = UBound(v, 1)
    ub2 = UBound(v, 2)
    If Err.Number <> 0 Then
        Say " read Union.Formula      : FAILED " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    Say " read Union.Formula      : " & Fmt(dt)
    Say "   UBound(1)=" & ub1 & "  UBound(2)=" & ub2 & _
        "   (sheet rows=" & rows & ", 6 columns picked)"

    If ub1 > 0 And ub2 > 0 Then
        On Error Resume Next
        Say "   v(1,1)=" & CStr(v(1, 1)) & "   v(1,last)=" & CStr(v(1, ub2))
        If rows >= 2 Then Say "   v(2,1)=" & CStr(v(2, 1))
        On Error GoTo 0

        nulls = 0
        On Error Resume Next
        For i = 1 To ub1
            For j = 1 To ub2
                If IsNull(v(i, j)) Then nulls = nulls + 1
            Next j
        Next i
        On Error GoTo 0
        Say "   Null cells in result: " & nulls & " of " & (ub1 * ub2)
    End If

    t = NowSec()
    ReDim out(1 To rows, 1 To 6)
    For k = 0 To UBound(pick)
        v = ws.Range(ws.Cells(1, pick(k)), ws.Cells(rows, pick(k))).Formula
        For i = 1 To rows
            out(i, k + 1) = v(i, 1)
        Next i
    Next k
    dt = NowSec() - t
    Say " read 6 cols one-by-one : " & Fmt(dt) & "  (assembled in VBA)"

    wb.Close False
End Sub

'==============================================================================
' Hashing helpers (same algorithm as the production design)
'==============================================================================
Private Function HashArray(ByRef a As Variant, _
                           ByVal rows As Long, ByVal cols As Long) As String
    Dim h1 As Long, h2 As Long, i As Long, j As Long, c As Long, k As Long
    Dim s As String

    h1 = &H811C9DC5
    h2 = &H9DC5811C

    For i = 1 To rows
        For j = 1 To cols
            s = CStr(a(i, j))
            c = Len(s)
            h1 = FnvStep(h1, c)
            h2 = FnvStep(h2, c Xor &H5A)
            For k = 1 To c
                h1 = FnvStep(h1, AscW(Mid$(s, k, 1)))
                h2 = FnvStep(h2, AscW(Mid$(s, k, 1)) Xor &H5A)
            Next k
            h1 = FnvStep(h1, 30)
            h2 = FnvStep(h2, 30 Xor &H5A)
        Next j
    Next i

    HashArray = LngToHex(h1) & LngToHex(h2)
End Function

Private Function HashStrMid(ByVal s As String, ByVal n As Long) As String
    Dim h1 As Long, h2 As Long, i As Long, c As Long
    h1 = &H811C9DC5
    h2 = &H9DC5811C
    For i = 1 To n
        c = AscW(Mid$(s, i, 1))
        h1 = FnvStep(h1, c)
        h2 = FnvStep(h2, c Xor &H5A)
    Next i
    HashStrMid = LngToHex(h1) & LngToHex(h2)
End Function

Private Function HashStrBytes(ByVal s As String, ByVal n As Long) As String
    Dim b() As Byte, i As Long
    Dim h1 As Long, h2 As Long
    Dim part As String

    part = Left$(s, n)
    b = StrConv(part, vbFromUnicode)

    h1 = &H811C9DC5
    h2 = &H9DC5811C
    For i = 0 To UBound(b)
        h1 = FnvStep(h1, CLng(b(i)))
        h2 = FnvStep(h2, CLng(b(i)) Xor &H5A)
    Next i
    HashStrBytes = LngToHex(h1) & LngToHex(h2)
End Function

' FNV-1a single step, 32-bit modular multiply.
' Do NOT use the Mod operator: VBA Long overflows because h*P is about 7.2e16.
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

Private Function LngToHex(ByVal x As Long) As String
    LngToHex = Right$("00000000" & Hex$(x And &HFFFFFFFF), 8)
End Function

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

Private Function Is64() As Boolean
    #If Win64 Then
        Is64 = True
    #Else
        Is64 = False
    #End If
End Function

Private Function FileExists(ByVal p As String) As Boolean
    FileExists = (Len(Dir$(p)) > 0)
End Function

Private Sub KillQuiet(ByVal p As String)
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    On Error GoTo 0
End Sub

Private Sub Say(ByVal s As String)
    m_log = m_log & s & vbCrLf
    Debug.Print s
End Sub

Private Sub WriteReport()
    Dim p As String, f As Integer

    p = m_tmp & "\mddt_t89.txt"
    On Error Resume Next
    f = FreeFile
    Open p For Output As #f
    Print #f, m_log
    Close #f
    If Err.Number <> 0 Then
        Err.Clear
        p = Environ$("USERPROFILE") & "\Desktop\mddt_t89.txt"
        f = FreeFile
        Open p For Output As #f
        Print #f, m_log
        Close #f
    End If
    On Error GoTo 0
    Debug.Print "report: " & p
End Sub