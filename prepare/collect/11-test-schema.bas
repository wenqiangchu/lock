Attribute VB_Name = "ZZTestSchema"
'==============================================================================
'  11-test-schema.bas  --  T21 + T22
'
'  Runs against a REAL mddt workbook opened read-only.  Writes nothing.
'
'  T21  AMG header case collision
'       The AMG prefix case is NOT stable. UCase$() is used everywhere to
'       normalise it. If a sheet contains both AMG01 and amg01, normalisation
'       collapses two real columns into ONE column lock, and we would silently
'       lock only one of them. Measure whether it can happen; if it does,
'       refuse checkout on that sheet.
'
'  T22  Anchor = product columns, capped at 5
'       PRODUCT columns are the ones to the LEFT of the first AMG column.
'       They are the only columns the tool never writes, so they are the only
'       safe anchor. The cap is 5, but a sheet with 3 product columns uses 3.
'
'       Guards measured here:
'         G1  uniqueness of the concatenated anchor key
'         G2  all-blank anchor rows must be skipped
'         G3  a sheet with ZERO product columns has no anchor at all
'             -> cannot identify rows -> must refuse checkout
'         G4  a very narrow anchor (1-2 product columns) is weak
'             -> uniqueness must be checked, flag for review
'
'  ENTRY POINT: ZZRUNSCHEMA([path])
'  REPORT     : %TEMP%\mddt_schema.txt
'
'  Pure ASCII. Excel VBA7.
'==============================================================================
Option Explicit

Private m_log As String
Private m_tmp As String

Public Sub ZZCOMPILEPROBE()
    Dim n As Long
    n = 1
    Debug.Print "compile ok " & n
End Sub

'==============================================================================
Public Sub ZZRUNSCHEMA(Optional ByVal TargetPath As String = "")
    On Error GoTo Fatal

    m_log = ""
    m_tmp = Environ$("TEMP")

    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open m_tmp & "\mddt_schema.txt" For Output As #ff
    Close #ff
    On Error GoTo 0

    Dim p As String
    If Len(TargetPath) > 0 Then
        p = TargetPath
    Else
        p = Environ$("ZZ_TARGET")
    End If

    Say "==================================================================="
    Say " T21 (AMG case collision) + T22 (anchor = product cols, max 5)"
    Say " Excel " & Application.Version & " build " & Application.Build
    If Len(p) = 0 Then
        Say " ERROR: no target. Pass a path, or set ZZ_TARGET."
        GoTo Done
    End If
    Say " target: " & p
    Say "==================================================================="

    If Len(p) > 0 And Len(Dir$(p)) = 0 Then
        Say " ERROR: file not found."
        GoTo Done
    End If

    Dim a As Boolean, b As Boolean, c As Boolean
    a = Application.EnableEvents
    b = Application.ScreenUpdating
    c = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(Filename:=p, ReadOnly:=True, UpdateLinks:=0)

    AggSheets wb
    ShowDetail wb

    wb.Close SaveChanges:=False

    Application.EnableEvents = a
    Application.ScreenUpdating = b
    Application.DisplayAlerts = c

Done:
    Say ""
    Say "ALL DONE."
    WriteReport
    Exit Sub

Fatal:
    Say "*** FATAL: " & Err.Number & " " & Err.Description
    On Error Resume Next
    WriteReport
End Sub

'==============================================================================
Private Sub AggSheets(ByVal wb As Workbook)
    Dim i As Long, c As Long
    Dim ws As Worksheet

    Dim tSheets As Long, tStudy As Long, tProduct As Long
    Dim tAMG As Long, tCollide As Long, tSheetsCollide As Long
    Dim tData As Long, tDup As Long, tBlank As Long
    Dim tNoAnchor As Long, tNarrow As Long
    Dim worstPct As Double, worstName As String
    Dim prodHist1 As Long, prodHist2 As Long, prodHist3 As Long
    Dim prodHist4 As Long, prodHist5 As Long, prodHist6 As Long

    Say ""
    Say "---- per-sheet -----------------------------------------------------"
    Say PadR("sheet", 22) & PadL("cols", 6) & PadL("prod", 5) & _
        PadL("anch", 5) & PadL("amg", 6) & PadL("rows", 9) & _
        PadL("dup", 6) & PadL("dup%", 8) & PadL("blank", 6) & _
        PadL("s", 3) & "  flags"

    For i = 1 To wb.Worksheets.Count
        Set ws = wb.Worksheets(i)

        Dim lastCol As Long, lastRow As Long
        Dim ghostNote As String
        Dim firstAMG As Long, nProduct As Long, nAnchor As Long, nAMG As Long
        Dim nCollide As Long, dupRows As Long, blankRows As Long, dataRows As Long
        Dim dupPct As Double

        ' VBA hoists Dim to procedure scope, so a Dim inside this loop is NOT
        ' re-initialised per iteration. Without these explicit resets the
        ' counters accumulate across sheets and every percentage is wrong.
        ghostNote = ""
        firstAMG = 0
        nProduct = 0
        nAnchor = 0
        nAMG = 0
        nCollide = 0
        dupRows = 0
        blankRows = 0
        dataRows = 0
        dupPct = 0

        lastCol = HeaderLastCol(ws, ghostNote)
        If lastCol < 1 Then
            Say PadR(ws.Name, 22) & "  <no header>"
            GoTo NextSheet
        End If

        ' the row bound must be scanned on a REAL data column. Column 1 is a
        ' product column by construction on every study sheet.
        lastRow = ColLastRow(ws, 1, ghostNote)
        If lastRow < 1 Then
            Say PadR(ws.Name, 22) & "  <no rows>"
            GoTo NextSheet
        End If

        '------ classify headers ------------------------------------------
        Dim hdr As String, key As String
        Dim seen As Object
        Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = 1

        firstAMG = 0
        nAMG = 0
        nCollide = 0
        For c = 1 To lastCol
            hdr = SafeText(ws.Cells(1, c))
            If IsAMG(hdr) Then
                If firstAMG = 0 Then firstAMG = c
                nAMG = nAMG + 1
                key = UCase$(Trim$(hdr))
                If seen.Exists(key) Then
                    nCollide = nCollide + 1
                Else
                    seen.Add key, c
                End If
            End If
        Next c

        If firstAMG > 0 Then
            nProduct = firstAMG - 1
        Else
            nProduct = lastCol          ' pure product sheet, no AMG at all
        End If

        nAnchor = nProduct
        If nAnchor > 5 Then nAnchor = 5
        If nAnchor < 0 Then nAnchor = 0

        '------ G1 + G2 : anchor uniqueness -------------------------------
        If nAnchor >= 1 And lastRow >= 2 Then
            Dim arrA As Variant
            arrA = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, nAnchor)).Value2

            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")

            Dim r As Long, cc As Long
            Dim parts() As String
            Dim isBlank As Boolean
            ReDim parts(1 To nAnchor)

            For r = 1 To UBound(arrA, 1)
                isBlank = True
                For cc = 1 To nAnchor
                    parts(cc) = Norm(arrA(r, cc))
                    If Len(parts(cc)) > 0 Then isBlank = False
                Next cc

                If isBlank Then
                    blankRows = blankRows + 1
                Else
                    dataRows = dataRows + 1
                    Dim sb As String
                    sb = Join(parts, vbTab)
                    If dict.Exists(sb) Then
                        dict(sb) = dict(sb) + 1
                    Else
                        dict.Add sb, 1
                    End If
                End If
            Next r

            Dim k As Variant
            For Each k In dict.Keys
                If dict(k) > 1 Then dupRows = dupRows + (dict(k) - 1)
            Next k
        End If

        If dataRows > 0 Then dupPct = dupRows / dataRows Else dupPct = 0

        '------ flags ------------------------------------------------------
        Dim flags As String
        flags = ""
        If nCollide > 0 Then flags = flags & "CASE-COLLIDE(" & nCollide & ") "
        If nAMG > 0 And nProduct = 0 Then flags = flags & "NO-PRODUCT-COLS "
        If nAMG > 0 And nProduct >= 1 And nProduct <= 2 Then flags = flags & "NARROW-ANCHOR "
        If nProduct > 5 Then flags = flags & "PROD>5 "
        If dupPct > 0.05 Then flags = flags & "DUP>5% "
        If blankRows > 0 Then flags = flags & "BLANK(" & blankRows & ") "

        Say PadR(ws.Name, 22) & PadL(CStr(lastCol), 6) & _
            PadL(CStr(nProduct), 5) & PadL(CStr(nAnchor), 5) & _
            PadL(CStr(nAMG), 6) & PadL(CStr(dataRows), 9) & _
            PadL(CStr(dupRows), 6) & PadL(Format$(dupPct * 100, "0.0"), 8) & _
            PadL(CStr(blankRows), 6) & _
            PadL(iif(nAMG > 0, "Y", "-"), 3) & "  " & flags & ghostNote

        '------ totals -----------------------------------------------------
        tSheets = tSheets + 1
        If nAMG > 0 Then
            tStudy = tStudy + 1
        Else
            tProduct = tProduct + 1
        End If
        tAMG = tAMG + nAMG
        tCollide = tCollide + nCollide
        If nCollide > 0 Then tSheetsCollide = tSheetsCollide + 1
        tData = tData + dataRows
        tDup = tDup + dupRows
        tBlank = tBlank + blankRows
        If nAMG > 0 And nProduct = 0 Then tNoAnchor = tNoAnchor + 1
        If nAMG > 0 And nProduct >= 1 And nProduct <= 2 Then tNarrow = tNarrow + 1
        If dupPct > worstPct Then
            worstPct = dupPct
            worstName = ws.Name
        End If

        Select Case nProduct
            Case 0: prodHist1 = prodHist1 + 1
            Case 1: prodHist2 = prodHist2 + 1
            Case 2: prodHist3 = prodHist3 + 1
            Case 3: prodHist4 = prodHist4 + 1
            Case 4: prodHist5 = prodHist5 + 1
            Case Else: prodHist6 = prodHist6 + 1
        End Select

NextSheet:
    Next i

    Say ""
    Say "---- TOTALS ---------------------------------------------------------"
    Say " sheets scanned          : " & tSheets
    Say "   study sheets (has AMG): " & tStudy
    Say "   product sheets        : " & tProduct
    Say " AMG columns total       : " & tAMG
    Say " sheets w/ case collision: " & tSheetsCollide & "  (collisions=" & tCollide & ")"
    Say " data rows total         : " & tData
    Say " duplicate anchor rows   : " & tDup
    Say " all-blank anchor rows   : " & tBlank
    Say " sheets w/ NO product col: " & tNoAnchor
    Say " sheets w/ NARROW anchor : " & tNarrow & "   (1-2 product cols)"
    Say " worst duplicate sheet   : " & worstName & "  (" & Format$(worstPct * 100, "0.00") & "%)"
    Say ""
    Say " product-column count histogram:"
    Say "   0 prod cols : " & prodHist1
    Say "   1 prod col  : " & prodHist2
    Say "   2 prod cols : " & prodHist3
    Say "   3 prod cols : " & prodHist4
    Say "   4 prod cols : " & prodHist5
    Say "   5+ prod cols: " & prodHist6
    Say ""

    Say "---- VERDICT --------------------------------------------------------"
    If tCollide > 0 Then
        Say " T21: CASE COLLISION EXISTS on " & tSheetsCollide & " sheet(s)."
        Say "      UCase$() collapses two real columns into one lock."
        Say "      Action: runtime check is MANDATORY, refuse checkout there."
    Else
        Say " T21: no case collision in this file."
        Say "      Keep the runtime check anyway (case normalisation is required"
        Say "      for matching, so the hazard exists by construction)."
    End If
    Say ""

    If tNoAnchor > 0 Then
        Say " T22-G3: " & tNoAnchor & " study sheet(s) have ZERO product columns."
        Say "      No anchor possible -> those sheets can never be checked out."
        Say "      They must be excluded from the tool (or given an anchor rule)."
    Else
        Say " T22-G3: every study sheet has at least 1 product column. OK."
    End If
    Say ""

    If tNarrow > 0 Then
        Say " T22-G4: " & tNarrow & " sheet(s) have only 1-2 product columns."
        Say "      Anchor is weak there. Uniqueness check must be enforced."
    End If
    Say ""

    If worstPct > 0.05 Then
        Say " T22-G1: DUPLICATE RATE TOO HIGH on '" & worstName & "' = " & _
            Format$(worstPct * 100, "0.00") & "%."
        Say "      Product columns are NOT sufficient there -> refuse checkout."
    ElseIf tDup > 0 Then
        Say " T22-G1: " & tDup & " duplicate anchor row(s) overall."
        Say "      Plan: list them, let the user confirm, then use a composite"
        Say "      key with an occurrence index."
    Else
        Say " T22-G1: anchor key is unique everywhere. Guard passes."
    End If
    Say ""
End Sub

'==============================================================================
Private Sub ShowDetail(ByVal wb As Workbook)
    Dim i As Long, c As Long
    Dim ws As Worksheet
    Dim hdr As String, key As String
    Dim seen As Object
    Dim anyShown As Boolean

    Say ""
    Say "---- T21 detail: case-insensitive duplicate AMG headers --------------"
    For i = 1 To wb.Worksheets.Count
        Set ws = wb.Worksheets(i)
        Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = 1

        Dim hits As String
        hits = ""
        Dim zz As String
        For c = 1 To HeaderLastCol(ws, zz)
            hdr = SafeText(ws.Cells(1, c))
            If IsAMG(hdr) Then
                key = UCase$(Trim$(hdr))
                If seen.Exists(key) Then
                    hits = hits & "   col " & seen(key) & "=[" & seen(key & "!o") & "]" & _
                                   "  vs  col " & c & "=[" & hdr & "]" & vbCrLf
                Else
                    seen.Add key, c
                    seen.Add key & "!o", hdr
                End If
            End If
        Next c

        If Len(hits) > 0 Then
            anyShown = True
            Say " sheet: " & ws.Name
            Say hits
        End If
    Next i
    If Not anyShown Then Say " (none)"

    Say ""
    Say "---- header / anchor sample (first 4 study sheets) ------------------"
    Dim shown As Long
    For i = 1 To wb.Worksheets.Count
        If shown >= 4 Then Exit For
        Set ws = wb.Worksheets(i)

        Dim firstAMG As Long, c2 As Long
        Dim zz2 As String
        firstAMG = 0
        For c2 = 1 To HeaderLastCol(ws, zz2)
            If IsAMG(SafeText(ws.Cells(1, c2))) Then
                firstAMG = c2
                Exit For
            End If
        Next c2

        If firstAMG > 0 Then
            shown = shown + 1
            Dim nAnchor As Long
            nAnchor = firstAMG - 1
            If nAnchor > 5 Then nAnchor = 5

            Dim anchorStr As String, amgStr As String
            anchorStr = ""
            For c2 = 1 To nAnchor
                anchorStr = anchorStr & "[" & SafeText(ws.Cells(1, c2)) & "]"
            Next c2
            amgStr = ""
            For c2 = firstAMG To HeaderLastCol(ws, zz2)
                If IsAMG(SafeText(ws.Cells(1, c2))) Then
                    If Len(amgStr) < 120 Then amgStr = amgStr & "[" & SafeText(ws.Cells(1, c2)) & "]"
                End If
            Next c2

            Say " " & ws.Name
            Say "   anchor (" & nAnchor & " cols): " & anchorStr
            Say "   AMG starts at col " & firstAMG & " : " & amgStr
        End If
    Next i
End Sub

'==============================================================================
Private Function IsAMG(ByVal hdr As String) As Boolean
    Dim t As String
    t = UCase$(Trim$(hdr))
    If Len(t) < 3 Then Exit Function
    IsAMG = (Left$(t, 3) = "AMG")
End Function

Private Function iif(ByVal cond As Boolean, ByVal a As String, ByVal b As String) As String
    If cond Then iif = a Else iif = b
End Function

Private Function PadR(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then
        PadR = Left$(s, n)
    Else
        PadR = s & Space$(n - Len(s))
    End If
End Function

Private Function PadL(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then
        PadL = Left$(s, n)
    Else
        PadL = Space$(n - Len(s)) & s
    End If
End Function

' Normalise a cell for key comparison.
' Numbers: force a locale-independent general form, otherwise 1 vs 1.0
' could produce different strings depending on the regional decimal mark.
Private Function Norm(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then
        Norm = "#ERR"
        Exit Function
    End If
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    On Error Resume Next
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency
            s = Trim$(CStr(v))
            If InStr(s, ",") > 0 Then s = Replace(s, ",", ".")
        Case vbDate
            s = Format$(v, "yyyy-mm-dd hh:nn:ss")
        Case Else
            s = Trim$(CStr(v))
    End Select
    If Err.Number <> 0 Then
        s = ""
        Err.Clear
    End If
    On Error GoTo 0

    Norm = s
End Function

Private Function SafeText(ByVal r As Range) As String
    Dim s As String
    On Error Resume Next
    s = CStr(r.Value2)
    If Err.Number <> 0 Then
        s = ""
        Err.Clear
    End If
    On Error GoTo 0
    SafeText = s
End Function

'==============================================================================
'  Bounds: NEVER use UsedRange or Find.
'
'  UsedRange cannot be trusted here: one sheet in the real file is inflated to
'  a 1.1M x 2M dimension by stray formatting, so UsedRange would report a
'  million rows that contain no data.
'
'  Find cannot be trusted either: measured on Excel 16.0 build 19127,
'  Range.Find(What:="*") returned $B$1 for BOTH xlByColumns/xlPrevious and
'  xlByRows/xlPrevious on every sheet of a test file, i.e. it silently gave
'  the wrong answer instead of failing.
'
'  So: read a bounded block as an array in ONE COM call and scan for the last
'  non-empty entry. One call, no ambiguity, and empty cells stay Empty even
'  inside an inflated dimension.
'==============================================================================

' last column that has a non-empty value in ROW 1 (the header row)
Private Function HeaderLastCol(ByVal ws As Worksheet, ByRef ghostNote As String) As Long
    Const MAX_COLS As Long = 4096

    Dim arr As Variant
    Dim i As Long, n As Long

    On Error Resume Next
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, MAX_COLS)).Value2
    If Err.Number <> 0 Then
        Err.Clear
        HeaderLastCol = 0
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    n = 0
    For i = 1 To MAX_COLS
        If Not IsCellEmpty(arr(1, i)) Then n = i
    Next i

    If n >= MAX_COLS Then ghostNote = ghostNote & "HDR-AT-SCAN-LIMIT "
    HeaderLastCol = n
End Function

' last row that has a non-empty value in the given column.
' Doubles the window until the data ends before the window edge, so an
' inflated dimension costs a little time but never truncates real data.
Private Function ColLastRow(ByVal ws As Worksheet, ByVal col As Long, _
                            ByRef ghostNote As String) As Long
    Const MAX_ROWS As Long = 400000

    Dim win As Long, arr As Variant
    Dim r As Long, n As Long
    Dim hitEdge As Boolean

    If col < 1 Then Exit Function

    win = 8192
    Do
        If win > MAX_ROWS Then win = MAX_ROWS

        On Error Resume Next
        arr = ws.Range(ws.Cells(1, col), ws.Cells(win, col)).Value2
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            ghostNote = ghostNote & "ROW-READ-FAIL "
            Exit Function
        End If
        On Error GoTo 0

        n = 0
        For r = win To 1 Step -1
            If Not IsCellEmpty(arr(r, 1)) Then
                n = r
                Exit For
            End If
        Next r

        hitEdge = (n = win)
        If Not hitEdge Then
            ColLastRow = n
            Exit Function
        End If

        If win >= MAX_ROWS Then
            ghostNote = ghostNote & "ROW-AT-SCAN-LIMIT "
            ColLastRow = n
            Exit Function
        End If

        win = win * 2
    Loop
End Function

Private Function IsCellEmpty(ByVal cell As Variant) As Boolean
    If IsEmpty(cell) Then
        IsCellEmpty = True
    ElseIf IsNull(cell) Then
        IsCellEmpty = True
    ElseIf IsError(cell) Then
        IsCellEmpty = False
    ElseIf VarType(cell) = vbString Then
        IsCellEmpty = (Len(Trim$(CStr(cell))) = 0)
    Else
        IsCellEmpty = False
    End If
End Function

Private Sub Say(ByVal s As String)
    Dim f As Integer
    m_log = m_log & s & vbCrLf
    Debug.Print s
    ' flush immediately: if the run dies inside a COM call we still want to
    ' know the last step that succeeded
    On Error Resume Next
    f = FreeFile
    Open m_tmp & "\mddt_schema.txt" For Append As #f
    Print #f, s
    Close #f
    On Error GoTo 0
End Sub

Private Sub WriteReport()
    Dim p As String, f As Integer
    p = m_tmp & "\mddt_schema.txt"
    On Error Resume Next
    f = FreeFile
    Open p For Output As #f
    Print #f, m_log
    Close #f
    On Error GoTo 0
    Debug.Print "report: " & p
End Sub