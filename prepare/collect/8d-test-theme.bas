Attribute VB_Name = "ZZTestTheme"
'==============================================================================
'  8d-test-theme.bas  --  T7: does a colour survive normalising/round-trips?
'
'  Pure ASCII. Excel VBA7.
'
'  Why this matters
'  ----------------
'  Fill colour and font colour are part of the change whitelist. If the SAME
'  visible colour reports a DIFFERENT Interior.Color value depending on how it
'  was applied (palette vs theme vs explicit RGB) or on the workbook's theme,
'  then the three-way merge will report phantom differences and produce fake
'  conflicts.
'
'  Questions
'  ---------
'  T7a  Does ThemeColor-based fill report a different .Color than explicit RGB?
'  T7b  Does changing the workbook theme change an already-set .Color?
'  T7c  Does a save / close / reopen change .Color?
'  T7d  Does .Color survive Range.Copy across workbooks with different themes?
'  T7e  What do the "no fill" states report (none / automatic / white)?
'  T7f  Does Font.Color behave the same way as Interior.Color?
'
'  Entry point: ZZRUNTHEME
'  Report     : %TEMP%\mddt_theme.txt
'==============================================================================
Option Explicit

Private m_log As String
Private Const THEME_DIR As String = _
    "C:\Program Files\Microsoft Office\root\Document Themes 16\"

Public Sub ZZCOMPILEPROBE()
    Dim n As Long
    n = 1
    Debug.Print "compile ok " & n
End Sub

Public Sub ZZRUNTHEME()
    On Error GoTo Fail

    m_log = ""

    Say "==================================================================="
    Say " T7  COLOUR STABILITY  --  Excel " & Application.Version & _
        " build " & Application.Build
    Say "==================================================================="

    Dim a As Boolean, b As Boolean, c As Boolean
    a = Application.EnableEvents
    b = Application.ScreenUpdating
    c = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next

    TestThemeColorVsRgb
    If Err.Number <> 0 Then
        Say "  [T7a aborted: " & Err.Number & " " & Err.Description & "]"
        Err.Clear
    End If

    TestNoFillStates
    If Err.Number <> 0 Then
        Say "  [T7e aborted: " & Err.Number & " " & Err.Description & "]"
        Err.Clear
    End If

    TestFontColor
    If Err.Number <> 0 Then
        Say "  [T7f aborted: " & Err.Number & " " & Err.Description & "]"
        Err.Clear
    End If

    TestThemeSwap
    If Err.Number <> 0 Then
        Say "  [T7b aborted: " & Err.Number & " " & Err.Description & "]"
        Err.Clear
    End If

    TestSaveReopen
    If Err.Number <> 0 Then
        Say "  [T7c aborted: " & Err.Number & " " & Err.Description & "]"
        Err.Clear
    End If

    TestCrossWorkbookCopy
    If Err.Number <> 0 Then
        Say "  [T7d aborted: " & Err.Number & " " & Err.Description & "]"
        Err.Clear
    End If

    On Error GoTo 0

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
' T7a -- theme colour vs explicit RGB for the same visible colour
'==============================================================================
Private Sub TestThemeColorVsRgb()
    Dim wb As Workbook, ws As Worksheet

    Say ""
    Say "---- T7a: theme colour vs explicit RGB -------------------------"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' xlThemeColorAccent1 = 5. TintAndShade 0 = the pure theme colour.
    On Error Resume Next
    ws.Range("A1").Interior.ThemeColor = 5
    ws.Range("A1").Interior.TintAndShade = 0
    If Err.Number <> 0 Then
        Say " ThemeColor not supported: " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    ' read back the RGB that Excel resolved it to, then apply that RGB plainly
    Dim resolved As Long
    resolved = ws.Range("A1").Interior.Color
    Say " A1 set via ThemeColor=5/Tint=0  -> .Color = " & resolved & _
        "  hex=" & Hex$(resolved And &HFFFFFF)
    Say " A1 .ThemeColor = " & ws.Range("A1").Interior.ThemeColor & _
        "  .TintAndShade = " & ws.Range("A1").Interior.TintAndShade

    ws.Range("A2").Interior.Color = resolved
    Say " A2 set via explicit RGB         -> .Color = " & ws.Range("A2").Interior.Color & _
        "  .ThemeColor = " & ws.Range("A2").Interior.ThemeColor
    Say " A1.Color == A2.Color ? " & (ws.Range("A1").Interior.Color = ws.Range("A2").Interior.Color)

    ' a themed tint, which is the common user case (Theme colour, lighter 40%)
    On Error Resume Next
    ws.Range("A3").Interior.ThemeColor = 5
    ws.Range("A3").Interior.TintAndShade = 0.4
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    Say " A3 ThemeColor=5/Tint=0.4        -> .Color = " & _
        ws.Range("A3").Interior.Color & "  .TintAndShade = " & _
        ws.Range("A3").Interior.TintAndShade

    Dim tinted As Long
    tinted = ws.Range("A3").Interior.Color
    ws.Range("A4").Interior.Color = tinted
    Say " A4 same RGB, set explicitly     -> .Color = " & ws.Range("A4").Interior.Color & _
        "  .TintAndShade = " & ws.Range("A4").Interior.TintAndShade
    Say "   (tint is lost on explicit set? A3.Tint=" & _
        ws.Range("A3").Interior.TintAndShade & "  A4.Tint=" & _
        ws.Range("A4").Interior.TintAndShade & ")"

    wb.Close False
End Sub

'==============================================================================
' T7b/T7e -- the "no fill" states
'==============================================================================
Private Sub TestNoFillStates()
    Dim wb As Workbook, ws As Worksheet

    Say ""
    Say "---- T7e: no-fill states ---------------------------------------"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    Say " pristine cell      .Color=" & ws.Range("A1").Interior.Color & _
        "  ColorIndex=" & ws.Range("A1").Interior.ColorIndex & _
        "  Pattern=" & ws.Range("A1").Interior.Pattern & _
        "  ThemeColor=" & ws.Range("A1").Interior.ThemeColor & _
        "  Tint=" & ws.Range("A1").Interior.TintAndShade

    ws.Range("A2").Interior.Pattern = -4142         ' xlNone
    Say " Pattern=xlNone     .Color=" & ws.Range("A2").Interior.Color & _
        "  ColorIndex=" & ws.Range("A2").Interior.ColorIndex & _
        "  Pattern=" & ws.Range("A2").Interior.Pattern

    ws.Range("A3").Interior.ColorIndex = -4142      ' xlColorIndexNone
    Say " ColorIndex=None    .Color=" & ws.Range("A3").Interior.Color & _
        "  ColorIndex=" & ws.Range("A3").Interior.ColorIndex

    ws.Range("A4").Interior.ColorIndex = 0          ' xlColorIndexAutomatic
    Say " ColorIndex=Auto    .Color=" & ws.Range("A4").Interior.Color & _
        "  ColorIndex=" & ws.Range("A4").Interior.ColorIndex

    ws.Range("A5").Interior.Color = RGB(255, 255, 255)
    Say " explicit white     .Color=" & ws.Range("A5").Interior.Color & _
        "  ColorIndex=" & ws.Range("A5").Interior.ColorIndex

    ws.Range("A6").Interior.Color = 16777215        ' xlNone as Color
    Say " Color=16777215     .Color=" & ws.Range("A6").Interior.Color & _
        "  ColorIndex=" & ws.Range("A6").Interior.ColorIndex & _
        "  Pattern=" & ws.Range("A6").Interior.Pattern

    wb.Close False
End Sub

'==============================================================================
' T7f -- font colour
'==============================================================================
Private Sub TestFontColor()
    Dim wb As Workbook, ws As Worksheet

    Say ""
    Say "---- T7f: font colour behaviour --------------------------------"

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ws.Range("A1").Value2 = "x"
    Say " pristine          .Font.Color=" & ws.Range("A1").Font.Color & _
        "  ColorIndex=" & ws.Range("A1").Font.ColorIndex & _
        "  ThemeColor=" & ws.Range("A1").Font.ThemeColor & _
        "  Tint=" & ws.Range("A1").Font.TintAndShade

    ws.Range("A2").Font.Color = RGB(255, 0, 0)
    Say " RGB(255,0,0)      .Font.Color=" & ws.Range("A2").Font.Color & _
        "  ColorIndex=" & ws.Range("A2").Font.ColorIndex

    On Error Resume Next
    ws.Range("A3").Font.ThemeColor = 5
    ws.Range("A3").Font.TintAndShade = 0
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    Say " ThemeColor=5      .Font.Color=" & ws.Range("A3").Font.Color & _
        "  ThemeColor=" & ws.Range("A3").Font.ThemeColor

    Dim themed As Long
    themed = ws.Range("A3").Font.Color
    On Error Resume Next
    ws.Range("A4").Font.Color = themed
    If Err.Number <> 0 Then
        Say " setting Font.Color=" & themed & " FAILED " & Err.Number & _
            " " & Err.Description
        Err.Clear
    Else
        Say " same RGB explicit .Font.Color=" & ws.Range("A4").Font.Color & _
            "  ThemeColor=" & ws.Range("A4").Font.ThemeColor
        Say "   (ThemeColor dropped? A3=" & ws.Range("A3").Font.ThemeColor & _
            " A4=" & ws.Range("A4").Font.ThemeColor & ")"
    End If
    On Error GoTo 0

    ' automatic font colour
    On Error Resume Next
    ws.Range("A5").Font.ColorIndex = -4105          ' xlAutomatic
    If Err.Number <> 0 Then
        Say " ColorIndex=Automatic FAILED " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        Say " ColorIndex=Automatic .Font.Color=" & ws.Range("A5").Font.Color & _
            "  ColorIndex=" & ws.Range("A5").Font.ColorIndex
    End If
    On Error GoTo 0

    wb.Close False
End Sub

'==============================================================================
' T7b -- swap the workbook theme and see if existing colours change
'==============================================================================
Private Sub TestThemeSwap()
    Dim wb As Workbook, ws As Worksheet
    Dim before As Long, after As Long
    Dim logoBefore As Long, logoAfter As Long
    Dim themeA As String, themeB As String

    Say ""
    Say "---- T7b: does swapping the theme change .Color? ---------------"

    themeA = THEME_DIR & "Office Theme.thmx"
    themeB = THEME_DIR & "Facet.thmx"

    If Dir$(themeA) = "" Or Dir$(themeB) = "" Then
        Say " theme files not found, skipping"
        Say "   looked for: " & themeA
        Say "               " & themeB
        Exit Sub
    End If

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' one themed fill, one explicit RGB fill, one wordart-ish logo fill
    On Error Resume Next
    ws.Range("A1").Interior.ThemeColor = 5
    ws.Range("A1").Interior.TintAndShade = 0
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    ws.Range("A2").Interior.Color = RGB(12, 34, 56)

    before = ws.Range("A1").Interior.Color
    logoBefore = ws.Range("A2").Interior.Color
    Say " BEFORE  themed A1=" & before & "   explicit A2=" & logoBefore & _
        "   A1.ThemeColor=" & ws.Range("A1").Interior.ThemeColor

    On Error Resume Next
    wb.ApplyTheme themeB
    If Err.Number <> 0 Then
        Say " ApplyTheme failed: " & Err.Number & " " & Err.Description
        Err.Clear
        wb.Close False
        Exit Sub
    End If
    On Error GoTo 0

    after = ws.Range("A1").Interior.Color
    logoAfter = ws.Range("A2").Interior.Color
    Say " AFTER   themed A1=" & after & "   explicit A2=" & logoAfter & _
        "   A1.ThemeColor=" & ws.Range("A1").Interior.ThemeColor
    Say " themed fill changed by theme swap? " & (before <> after)
    Say " explicit fill changed by theme swap? " & (logoBefore <> logoAfter)
    If before <> after Then
        Say "   >>> PHANTOM DIFF RISK: same cell, value changed without any edit"
    End If

    ' restore
    On Error Resume Next
    wb.ApplyTheme themeA
    On Error GoTo 0
    Say " after restoring theme, A1=" & ws.Range("A1").Interior.Color & _
        "  (original " & before & ", back? " & _
        (ws.Range("A1").Interior.Color = before) & ")"

    wb.Close False
End Sub

'==============================================================================
' T7c -- save / close / reopen
'==============================================================================
Private Sub TestSaveReopen()
    Dim wb As Workbook, ws As Worksheet
    Dim p As String
    Dim c1 As Long, c2 As Long, f1 As Long, f2 As Long, t1 As Long, t2 As Long
    Dim ci1 As Long, ci2 As Long

    Say ""
    Say "---- T7c: does .Color survive save + reopen? -------------------"

    p = Environ$("TEMP") & "\mddt_t7c.xlsx"
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    On Error GoTo 0

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ws.Range("A1").Value2 = "probe"
    ws.Range("A1").Interior.Color = RGB(200, 100, 50)
    ws.Range("A1").Font.Color = RGB(10, 120, 240)
    ws.Range("A1").Font.Bold = True

    On Error Resume Next
    ws.Range("A2").Interior.ThemeColor = 5
    ws.Range("A2").Interior.TintAndShade = 0.4
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    c1 = ws.Range("A1").Interior.Color
    f1 = ws.Range("A1").Font.Color
    t1 = ws.Range("A2").Interior.Color
    ci1 = ws.Range("A1").Interior.ColorIndex

    wb.SaveAs p, 51
    wb.Close False

    Set wb = Application.Workbooks.Open(p, 0, False)
    Set ws = wb.Worksheets(1)

    c2 = ws.Range("A1").Interior.Color
    f2 = ws.Range("A1").Font.Color
    t2 = ws.Range("A2").Interior.Color
    ci2 = ws.Range("A1").Interior.ColorIndex

    Say " A1 fill : before=" & c1 & "  after=" & c2 & "  same? " & (c1 = c2)
    Say " A1 font : before=" & f1 & "  after=" & f2 & "  same? " & (f1 = f2)
    Say " A1 ColorIndex: before=" & ci1 & "  after=" & ci2 & "  same? " & (ci1 = ci2)
    Say " A2 themed : before=" & t1 & "  after=" & t2 & "  same? " & (t1 = t2)
    Say " A2 ThemeColor after reopen = " & ws.Range("A2").Interior.ThemeColor & _
        "  Tint=" & ws.Range("A2").Interior.TintAndShade

    wb.Close False
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    On Error GoTo 0
End Sub

'==============================================================================
' T7d -- copy across workbooks with different themes
'==============================================================================
Private Sub TestCrossWorkbookCopy()
    Dim wbSrc As Workbook, wbDst As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim srcColor As Long, dstColor As Long
    Dim themeA As String, themeB As String

    Say ""
    Say "---- T7d: Copy into a workbook with a DIFFERENT theme ----------"

    themeA = THEME_DIR & "Office Theme.thmx"
    themeB = THEME_DIR & "Facet.thmx"

    Set wbSrc = Application.Workbooks.Add
    Set ws1 = wbSrc.Worksheets(1)
    Set wbDst = Application.Workbooks.Add
    Set ws2 = wbDst.Worksheets(1)

    If Dir$(themeA) <> "" And Dir$(themeB) <> "" Then
        On Error Resume Next
        wbSrc.ApplyTheme themeA
        wbDst.ApplyTheme themeB
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        Say " source theme = Office Theme, dest theme = Facet"
    Else
        Say " theme files missing; both workbooks keep the default theme"
    End If

    ws1.Range("A1").Value2 = "copied"
    ws1.Range("A1").Interior.Color = RGB(200, 100, 50)
    ws1.Range("A1").Font.Color = RGB(10, 120, 240)
    On Error Resume Next
    ws1.Range("A2").Interior.ThemeColor = 5
    ws1.Range("A2").Interior.TintAndShade = 0.4
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    srcColor = ws1.Range("A1").Interior.Color

    ws1.Range("A1:A2").Copy Destination:=ws2.Range("A1")
    Application.CutCopyMode = False

    dstColor = ws2.Range("A1").Interior.Color
    Say " explicit RGB  : src=" & srcColor & "  dst=" & dstColor & _
        "  same? " & (srcColor = dstColor)
    Say "  dst Font.Color = " & ws2.Range("A1").Font.Color & _
        "  (src " & ws1.Range("A1").Font.Color & ")"

    Say " themed cell   : src=" & ws1.Range("A2").Interior.Color & _
        "  dst=" & ws2.Range("A2").Interior.Color & _
        "  same? " & (ws1.Range("A2").Interior.Color = ws2.Range("A2").Interior.Color)
    Say "  dst ThemeColor=" & ws2.Range("A2").Interior.ThemeColor & _
        "  Tint=" & ws2.Range("A2").Interior.TintAndShade
    If ws1.Range("A2").Interior.Color <> ws2.Range("A2").Interior.Color Then
        Say "   >>> PHANTOM DIFF RISK on themed fills across themes"
    End If

    wbSrc.Close False
    wbDst.Close False
End Sub

'==============================================================================
' Utility
'==============================================================================
Private Sub Say(ByVal s As String)
    m_log = m_log & s & vbCrLf
    Debug.Print s
End Sub

Private Sub WriteReport()
    Dim p As String, f As Integer
    p = Environ$("TEMP") & "\mddt_theme.txt"
    On Error Resume Next
    f = FreeFile
    Open p For Output As #f
    Print #f, m_log
    Close #f
    On Error GoTo 0
    Debug.Print "report: " & p
End Sub