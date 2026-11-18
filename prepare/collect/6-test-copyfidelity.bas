Attribute VB_Name = "TestCopyFidelity"
' ============================================================
'  6-test-copyfidelity.bas  -  cross-workbook Range.Copy fidelity test
'
'  This is the technical premise of the whole design: write-back uses
'  Range.Copy instead of a Value2 array, and the only reason is that Copy
'  carries styles, hyperlinks, comments and validations along with the values.
'  This module verifies that premise.
'
'  Usage:
'    1. Open a BLANK workbook
'    2. Alt+F11 -> Insert -> Module -> paste this whole file
'    3. Put the cursor inside TestCopyFidelity and press F5
'    4. The report is written to Desktop\exs_copyfidelity.txt
'    5. Takes about 10-20 seconds
'
'  Already verified on Excel 16.0 x64. See VERIFIED.md for results.
'
'  NOTE: keep this file pure ASCII.
' ============================================================
Option Explicit

Private m_log As String
Private m_fail As Long
Private m_warn As Long

' Set to False to suppress MsgBox (for unattended runs).
Private Const SHOW_MSGBOX As Boolean = True

' Safe delete: no error if the file is absent (avoids error 53).
Private Sub SafeKill(ByVal p As String)
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    Err.Clear
    On Error GoTo 0
End Sub

' Safe mkdir: no error if the folder exists (avoids error 75).
Private Sub SafeMkDir(ByVal p As String)
    On Error Resume Next
    If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub TestCopyFidelity()
    Dim srcPath As String, dstPath As String
    Dim wbSrc As Workbook, wbDst As Workbook
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim t0 As Single

    m_log = "": m_fail = 0: m_warn = 0
    t0 = Timer

    LogLine "ExcelSync cross-workbook Copy fidelity test"
    LogLine "time : " & Format$(Now, "yyyy-mm-dd hh:mm:ss")
    LogLine "excel: " & Application.Version & "  " & Application.OperatingSystem & _
            "  bitness=" & IIf(Is64(), "x64", "x86")
    LogLine "user : " & Environ$("USERDOMAIN") & "\" & Environ$("USERNAME")
    LogLine ""

    Dim tmpDir As String
    tmpDir = Environ$("TEMP") & "\exs_fid_" & Format$(Now, "hhnnss")
    SafeMkDir tmpDir
    srcPath = tmpDir & "\src.xlsx"
    dstPath = tmpDir & "\dst.xlsx"

    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---------- build source workbook ----------
    LogSection "A. Build the source workbook"
    Set wbSrc = Workbooks.Add
    Set wsSrc = wbSrc.Worksheets(1)
    wsSrc.Name = "SrcData"

    ' second sheet used as an internal hyperlink target
    Dim wsTarget As Worksheet
    Set wsTarget = wbSrc.Worksheets.Add(After:=wsSrc)
    wsTarget.Name = "TargetSheet"
    wsTarget.Range("A1").Value = "jump target"

    ' values
    wsSrc.Range("A1").Value = "Key"
    wsSrc.Range("B1").Value = "Text"
    wsSrc.Range("C1").Value = "Number"
    wsSrc.Range("A2").Value = "K001"
    wsSrc.Range("B2").Value = "plain string"
    wsSrc.Range("C2").Value = 12345.6789
    wsSrc.Range("A3").Value = "K002"
    wsSrc.Range("B3").Value = "leading zero test"

    ' styles
    wsSrc.Range("A1:C1").Interior.Color = RGB(68, 114, 196)
    wsSrc.Range("A1:C1").Font.Color = RGB(255, 255, 255)
    wsSrc.Range("A1:C1").Font.Bold = True
    wsSrc.Range("A1:C1").Font.Size = 12
    wsSrc.Range("A1:C1").Font.Name = "Arial"
    wsSrc.Range("B2").Interior.Color = RGB(255, 242, 204)
    wsSrc.Range("B2").Font.Color = RGB(192, 0, 0)
    wsSrc.Range("B2").Font.Italic = True
    wsSrc.Range("C2").NumberFormat = "#,##0.00"
    ' IMPORTANT: set text format BEFORE assigning the value, otherwise Excel
    ' stores "00123" as the number 123 immediately.
    wsSrc.Range("C3").NumberFormat = "@"
    wsSrc.Range("C3").Value = "00123"
    wsSrc.Range("A3").HorizontalAlignment = xlCenter
    wsSrc.Range("A2:C3").VerticalAlignment = xlCenter
    wsSrc.Range("A2:C3").Borders.LineStyle = xlContinuous
    wsSrc.Range("A2:C3").Borders.Weight = xlThin

    ' merged cells
    wsSrc.Range("E1:F2").Merge
    wsSrc.Range("E1").Value = "merged area"

    ' hyperlinks (the key test items)
    wsSrc.Hyperlinks.Add Anchor:=wsSrc.Range("A4"), _
        Address:="", SubAddress:="'TargetSheet'!A1", TextToDisplay:="jump to target"
    wsSrc.Hyperlinks.Add Anchor:=wsSrc.Range("A5"), _
        Address:="https://example.com", TextToDisplay:="external link"

    ' comment
    On Error Resume Next
    wsSrc.Range("B4").AddComment "this is a comment"
    Err.Clear
    On Error GoTo EH

    ' data validation
    With wsSrc.Range("D2:D4").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="A,B,C"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' conditional formatting
    wsSrc.Range("C2:C4").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="1000"
    wsSrc.Range("C2:C4").FormatConditions(1).Interior.Color = RGB(198, 239, 206)

    ' row height / column width
    wsSrc.Rows(1).RowHeight = 28
    wsSrc.Columns(2).ColumnWidth = 25

    ' freeze panes (needs an activated visible window; ignore failure)
    On Error Resume Next
    wsSrc.Activate
    wsSrc.Range("A2").Select
    ActiveWindow.FreezePanes = True
    Err.Clear
    On Error GoTo EH

    Dim hBefore As String
    hBefore = SnapshotHash(wsSrc)
    LogLine "  source snapshot hash: " & hBefore

    wbSrc.SaveAs srcPath, xlOpenXMLWorkbook
    wbSrc.Close False
    LogLine "  source workbook saved: " & srcPath

    ' ---------- build destination and copy ----------
    LogSection "B. Build destination workbook and run Copy"
    Set wbDst = Workbooks.Add
    Set wsDst = wbDst.Worksheets(1)
    wsDst.Name = "DstData"
    wbDst.Worksheets.Add(After:=wsDst).Name = "TargetSheet"
    wbDst.SaveAs dstPath, xlOpenXMLWorkbook

    Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)
    Set wsSrc = wbSrc.Worksheets("SrcData")

    Dim usedRng As Range
    Set usedRng = wsSrc.Range("A1:F6")
    LogLine "  copying range: " & usedRng.Address(False, False) & "  (" & usedRng.Cells.Count & " cells)"

    Dim tCopy As Single: tCopy = Timer
    usedRng.Copy Destination:=wsDst.Range("A1")
    LogLine "  Copy elapsed: " & Format$((Timer - tCopy) * 1000, "0") & " ms"

    ' ---------- control: Value2 array write ----------
    LogSection "B2. Control: Value2 array write (expected to lose hyperlinks/styles)"
    Dim wsV As Worksheet
    Set wsV = wbDst.Worksheets.Add(After:=wbDst.Sheets(wbDst.Sheets.Count))
    wsV.Name = "Value2Test"
    Dim v As Variant
    v = usedRng.Value2
    wsV.Range("A1").Resize(UBound(v, 1), UBound(v, 2)).Value2 = v
    LogLine "  Value2 written hyperlink count = " & wsV.Range("A1:F6").Hyperlinks.Count & _
            "  (source=" & usedRng.Hyperlinks.Count & ")"
    LogLine "  Value2 written A1 fill color    = &H" & Hex(wsV.Range("A1").Interior.Color) & _
            "  (source=&H" & Hex(usedRng.Cells(1, 1).Interior.Color) & ")"
    If wsV.Range("A1:F6").Hyperlinks.Count = 0 Then
        LogLine "  CONFIRMED: Value2 loses all hyperlinks -> write-back must NOT use Value2"
    Else
        LogLine "  UNEXPECTED: Value2 kept the hyperlinks, re-check this"
    End If

    ' ---------- item by item comparison ----------
    LogSection "C. Item by item fidelity comparison"

    Dim same As Boolean
    same = ValuesEqual(usedRng, wsDst.Range("A1").Resize(usedRng.Rows.Count, usedRng.Columns.Count))
    Check "values (Value2)", same, "mismatch"

    Dim hlSrc As Long, hlDst As Long
    hlSrc = usedRng.Hyperlinks.Count
    hlDst = wsDst.Range("A1").Resize(usedRng.Rows.Count, usedRng.Columns.Count).Hyperlinks.Count
    Check "hyperlink count", (hlSrc = hlDst), "source=" & hlSrc & " dest=" & hlDst

    If hlDst >= 2 Then
        Dim subAddr As String, extAddr As String
        Dim h As Hyperlink
        subAddr = "": extAddr = ""
        For Each h In wsDst.Range("A1:F6").Hyperlinks
            If Len(h.SubAddress) > 0 Then subAddr = h.SubAddress
            If InStr(1, h.Address, "example.com", vbTextCompare) > 0 Then extAddr = h.Address
        Next
        Check "internal SubAddress preserved", (InStr(1, subAddr, "TargetSheet", vbTextCompare) > 0), "actual='" & subAddr & "'"
        Check "external Address preserved", (Len(extAddr) > 0), "actual='" & extAddr & "'"
        Check "Hyperlink TextToDisplay preserved", _
              (wsDst.Range("A4").Value = "jump to target"), "actual='" & CStr(wsDst.Range("A4").Value) & "'"
    End If

    Check "fill color A1", _
          (wsDst.Range("A1").Interior.Color = RGB(68, 114, 196)), _
          "source=&H" & Hex(RGB(68, 114, 196)) & " dest=&H" & Hex(wsDst.Range("A1").Interior.Color)
    Check "fill color B2", _
          (wsDst.Range("B2").Interior.Color = RGB(255, 242, 204)), _
          "dest=&H" & Hex(wsDst.Range("B2").Interior.Color)

    Check "font color B2", (wsDst.Range("B2").Font.Color = RGB(192, 0, 0)), "dest=&H" & Hex(wsDst.Range("B2").Font.Color)
    Check "bold A1", (wsDst.Range("A1").Font.Bold = True), "dest=" & wsDst.Range("A1").Font.Bold
    Check "italic B2", (wsDst.Range("B2").Font.Italic = True), "dest=" & wsDst.Range("B2").Font.Italic
    Check "font size A1 = 12", (wsDst.Range("A1").Font.Size = 12), "dest=" & wsDst.Range("A1").Font.Size
    Check "font name A1 = Arial", (wsDst.Range("A1").Font.Name = "Arial"), "dest=" & wsDst.Range("A1").Font.Name

    Check "number format C2", (wsDst.Range("C2").NumberFormat = "#,##0.00"), "dest=" & wsDst.Range("C2").NumberFormat
    Check "text format C3", (wsDst.Range("C3").NumberFormat = "@"), "dest=" & wsDst.Range("C3").NumberFormat
    Check "leading zero kept C3", (CStr(wsDst.Range("C3").Value) = "00123"), "dest='" & CStr(wsDst.Range("C3").Value) & "'"

    Check "horizontal center A3", (wsDst.Range("A3").HorizontalAlignment = xlCenter), "dest=" & wsDst.Range("A3").HorizontalAlignment
    Check "border A2", (wsDst.Range("A2").Borders(xlEdgeLeft).LineStyle = xlContinuous), "dest=" & wsDst.Range("A2").Borders(xlEdgeLeft).LineStyle
    Check "merged E1:F2", (wsDst.Range("E1:F2").MergeCells = True), "dest MergeCells=" & wsDst.Range("E1:F2").MergeCells

    Dim cmtDst As String
    On Error Resume Next
    cmtDst = wsDst.Range("B4").Comment.Text
    Err.Clear
    On Error GoTo EH
    Check "comment B4", (Len(cmtDst) > 0), "dest='" & cmtDst & "'"

    Dim dvDst As String
    On Error Resume Next
    dvDst = wsDst.Range("D2").Validation.Formula1
    Err.Clear
    On Error GoTo EH
    Check "data validation D2", (dvDst = "A,B,C"), "dest Formula1='" & dvDst & "'"

    Check "conditional format C2:C4", (wsDst.Range("C2:C4").FormatConditions.Count = 1), _
          "dest count=" & wsDst.Range("C2:C4").FormatConditions.Count

    ' Row height and column width are Row/Column properties, not Cell properties:
    ' Range.Copy(Destination) does not carry them. Known limitation, not a bug.
    Dim rhOk As Boolean, cwOk As Boolean
    rhOk = (Abs(wsDst.Rows(1).RowHeight - 28) < 0.6)
    cwOk = (Abs(wsDst.Columns(2).ColumnWidth - 25) < 0.6)
    LogLine "  row height Row1 after Copy = " & wsDst.Rows(1).RowHeight & " (source=28)"
    LogLine "  col width  ColB after Copy = " & wsDst.Columns(2).ColumnWidth & " (source=25)"
    If (Not rhOk) Or (Not cwOk) Then
        LogLine "  LIMITATION: Range.Copy does not carry row height / column width"
        LogLine "              copy them explicitly after Copy (demonstrated below)"
        wsDst.Rows(1).RowHeight = wsSrc.Rows(1).RowHeight
        wsDst.Columns(2).ColumnWidth = wsSrc.Columns(2).ColumnWidth
        LogLine "              after explicit copy: rowHeight=" & wsDst.Rows(1).RowHeight & _
                "  colWidth=" & wsDst.Columns(2).ColumnWidth
        If Abs(wsDst.Rows(1).RowHeight - 28) < 0.6 And Abs(wsDst.Columns(2).ColumnWidth - 25) < 0.6 Then
            LogLine "              CONFIRMED: explicit copy works, 3 lines of code in the implementation"
        Else
            LogLine "              explicit copy also failed, needs investigation"
        End If
    Else
        LogLine "  row height / column width were carried (unexpectedly good)"
    End If
    LogLine "  note: freeze panes is a window-level setting, never carried by Copy"

    ' ---------- the shrink problem ----------
    LogSection "D. The shrink problem (why Clear is mandatory before write-back)"
    Dim wsT As Worksheet
    Set wsT = wbDst.Worksheets.Add(After:=wbDst.Sheets(wbDst.Sheets.Count))
    wsT.Name = "ShrinkTest"

    Dim m As Long
    For m = 1 To 10
        wsT.Cells(m, 1).Value2 = "OLD" & m
        wsT.Cells(m, 2).Value2 = "old data " & m
    Next m
    LogLine "  step 1: wrote 10 rows of old data into A1:B10"
    LogLine "    A10 = '" & CStr(wsT.Cells(10, 1).Value2) & "'"

    Dim newData(1 To 4, 1 To 2) As Variant
    For m = 1 To 4
        newData(m, 1) = "NEW" & m
        newData(m, 2) = "new data " & m
    Next m

    ' wrong way: overwrite only A1:B4
    wsT.Range("A1:B4").Value2 = newData
    LogLine ""
    LogLine "  step 2 (WRONG): overwrite A1:B4 only, no Clear below"
    LogLine "    A4 = '" & CStr(wsT.Cells(4, 1).Value2) & "'   (expect NEW4)"
    LogLine "    A5 = '" & CStr(wsT.Cells(5, 1).Value2) & "'   (stale OLD5)"
    LogLine "    A10= '" & CStr(wsT.Cells(10, 1).Value2) & "'   (stale OLD10)"
    Dim hadResidue As Boolean
    hadResidue = (Len(CStr(wsT.Cells(5, 1).Value2)) > 0)
    If hadResidue Then
        LogLine "    CONFIRMED residue: old rows survive and look like a failed delete"
    End If

    ' right way: Clear the union first
    wsT.Range("A1:B10").ClearContents
    wsT.Range("A1:B10").ClearFormats
    wsT.Range("A1:B4").Value2 = newData
    LogLine ""
    LogLine "  step 3 (RIGHT): Clear A1:B10 (the union) then overwrite A1:B4"
    LogLine "    A4 = '" & CStr(wsT.Cells(4, 1).Value2) & "'   (expect NEW4)"
    LogLine "    A5 = '" & CStr(wsT.Cells(5, 1).Value2) & "'   (expect empty)"
    LogLine "    A10= '" & CStr(wsT.Cells(10, 1).Value2) & "'   (expect empty)"
    If Len(CStr(wsT.Cells(5, 1).Value2)) = 0 And Len(CStr(wsT.Cells(10, 1).Value2)) = 0 Then
        LogLine "    CONFIRMED: clearing the union leaves no residue"
        Check "no residue after clearing the union", True, ""
    Else
        Check "no residue after clearing the union", False, "residue remains"
    End If
    LogLine ""
    LogLine "  conclusion: clear Union(serverRange, localRange), otherwise deletes look lost"

    ' ---------- performance ----------
    LogSection "E. Performance baseline"
    Dim tFill As Single, i As Long, j As Long
    Dim big() As Variant
    ReDim big(1 To 500, 1 To 50)
    For i = 1 To 500
        For j = 1 To 50
            big(i, j) = "R" & i & "C" & j
        Next j
    Next i
    tFill = Timer
    wsDst.Range("H1").Resize(500, 50).Value2 = big
    LogLine "  write 500x50 = 25,000 cells (Value2): " & Format$((Timer - tFill) * 1000, "0") & " ms"

    tCopy = Timer
    wsDst.Range("H1").Resize(500, 50).Copy Destination:=wsDst.Range("BA1")
    LogLine "  copy  500x50 (Copy + styles)        : " & Format$((Timer - tCopy) * 1000, "0") & " ms"
    LogLine "  your real scale is about 250x25 = 6,250 cells, so expect well under 500 ms"

    ' ---------- conclusion ----------
    LogSection "F. Conclusion"
    If m_fail = 0 Then
        LogLine "  ALL KEY CHECKS PASSED -> Range.Copy is safe for write-back."
    Else
        LogLine "  " & m_fail & " check(s) failed. Send this report back."
        LogLine "  If the failure is about hyperlinks, write-back must become:"
        LogLine "  Clear + Value2 + explicitly rebuild Hyperlinks."
    End If
    LogLine ""
    LogLine "  total elapsed: " & Format$(Timer - t0, "0.0") & " s"

    wbSrc.Close False
    wbDst.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    WriteReport

    On Error Resume Next
    SafeKill srcPath
    SafeKill dstPath
    RmDir tmpDir
    Err.Clear
    On Error GoTo 0

    If SHOW_MSGBOX Then
        MsgBox "Test finished." & vbCrLf & vbCrLf & _
               "failed checks: " & m_fail & vbCrLf & _
               "report written to Desktop\exs_copyfidelity.txt" & vbCrLf & vbCrLf & _
               "Please send that file back.", vbInformation, "ExcelSync fidelity test"
    End If
    Exit Sub

EH:
    LogLine ""
    LogLine "RUNTIME ERROR: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not wbSrc Is Nothing Then wbSrc.Close False
    If Not wbDst Is Nothing Then wbDst.Close False
    WriteReport
    If SHOW_MSGBOX Then
        MsgBox "Test aborted: " & Err.Description & vbCrLf & "A partial report was written to the Desktop.", vbExclamation
    End If
End Sub

' ============================================================
Private Sub LogLine(s As String)
    m_log = m_log & s & vbCrLf
    Debug.Print s
End Sub

Private Sub LogSection(s As String)
    LogLine ""
    LogLine String(78, "=")
    LogLine "  " & s
    LogLine String(78, "=")
End Sub

Private Sub Check(name As String, ok As Boolean, detail As String)
    If ok Then
        LogLine "  PASS " & name
    Else
        m_fail = m_fail + 1
        LogLine "  FAIL " & name & "   <- " & detail
    End If
End Sub

Private Function ValuesEqual(r1 As Range, r2 As Range) As Boolean
    Dim a As Variant, b As Variant, i As Long, j As Long
    If r1.Rows.Count <> r2.Rows.Count Or r1.Columns.Count <> r2.Columns.Count Then
        ValuesEqual = False: Exit Function
    End If
    a = r1.Value2
    b = r2.Value2
    For i = 1 To UBound(a, 1)
        For j = 1 To UBound(a, 2)
            If CStr(a(i, j)) <> CStr(b(i, j)) Then ValuesEqual = False: Exit Function
        Next j
    Next i
    ValuesEqual = True
End Function

Private Function SnapshotHash(ws As Worksheet) As String
    ' Two-lane FNV-1a 32-bit, concatenated into 64 bits.
    ' This is the same algorithm the real implementation will use.
    Dim buf As String, v As Variant, i As Long, j As Long
    Dim lr As Long, lc As Long
    Dim h1 As Long, h2 As Long, c As Long
    lr = 6: lc = 6
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value2
    For i = 1 To lr
        For j = 1 To lc
            buf = buf & CStr(v(i, j)) & Chr$(30)
        Next j
    Next i
    h1 = &H811C9DC5
    h2 = &H9DC5811C
    For i = 1 To Len(buf)
        c = AscW(Mid$(buf, i, 1))
        h1 = FnvStep(h1, c)
        h2 = FnvStep(h2, c Xor &H5A)
    Next i
    SnapshotHash = LngToHex(h1) & LngToHex(h2)
End Function

' FNV-1a single step, 32-bit modular multiply.
' Do NOT use the Mod operator: VBA Mod coerces operands to Long, and h*P is
' about 7.2e16, which always overflows. Use Double segment arithmetic instead.
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
    ' hi*P and lo*P are both below 1.1e12, exactly representable in Double
    r = (hi * P - Int(hi * P / 65536#) * 65536#) * 65536#
    r = r + (lo * P - Int(lo * P / 4294967296#) * 4294967296#)
    r = r - Int(r / 4294967296#) * 4294967296#
    If r > 2147483647# Then r = r - 4294967296#
    FnvStep = CLng(r)
End Function

Private Function LngToHex(ByVal x As Long) As String
    LngToHex = Right$("00000000" & Hex$(x And &HFFFFFFFF), 8)
End Function

Private Function Is64() As Boolean
    #If Win64 Then
        Is64 = True
    #Else
        Is64 = False
    #End If
End Function

Private Sub WriteReport()
    Dim p As String
    Dim f As Integer

    p = Environ$("USERPROFILE") & "\Desktop\exs_copyfidelity.txt"
    On Error Resume Next
    f = FreeFile
    Open p For Output As #f
    Print #f, m_log
    Close #f
    If Err.Number <> 0 Then
        p = Environ$("TEMP") & "\exs_copyfidelity.txt"
        Err.Clear
        f = FreeFile
        Open p For Output As #f
        Print #f, m_log
        Close #f
        Debug.Print "report written to: " & p
    End If
    On Error GoTo 0
    Debug.Print "report path: " & p
End Sub
