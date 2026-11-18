Attribute VB_Name = "ZZTestPropDirty"
'==============================================================================
'  8e-test-propdirty.bas  --  T18
'
'  Question: does adding / changing / reading / deleting a
'  CustomDocumentProperty make the workbook dirty?
'
'  Why it matters
'  --------------
'  The Import flow wants to stamp the local copy with a role marker
'  (MDDT_ROLE = LOCAL). If that stamp makes the workbook look dirty, the user
'  sees "unsaved changes" for a file they have not touched, and our own
'  "is the local copy modified?" logic gets a false positive.
'
'  Pure ASCII. Excel VBA7.
'  Entry point: ZZRUNPROP
'  Report     : %TEMP%\mddt_propdirty.txt
'==============================================================================
Option Explicit

Private m_log As String
Private m_tmp As String

Public Sub ZZCOMPILEPROBE()
    Dim n As Long
    n = 1
    Debug.Print "compile ok " & n
End Sub

Public Sub ZZRUNPROP()
    On Error GoTo Fail

    m_log = ""
    m_tmp = Environ$("TEMP")

    Say "==================================================================="
    Say " T18  document-property dirty behaviour"
    Say " Excel " & Application.Version & " build " & Application.Build
    Say "==================================================================="

    Dim a As Boolean, b As Boolean, c As Boolean
    a = Application.EnableEvents
    b = Application.ScreenUpdating
    c = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next

    T18a_add_on_clean
    If Err.Number <> 0 Then Say "  [T18a aborted " & Err.Number & " " & Err.Description & "]": Err.Clear

    T18b_change_value
    If Err.Number <> 0 Then Say "  [T18b aborted " & Err.Number & " " & Err.Description & "]": Err.Clear

    T18c_read_only
    If Err.Number <> 0 Then Say "  [T18c aborted " & Err.Number & " " & Err.Description & "]": Err.Clear

    T18d_delete
    If Err.Number <> 0 Then Say "  [T18d aborted " & Err.Number & " " & Err.Description & "]": Err.Clear

    T18e_persist
    If Err.Number <> 0 Then Say "  [T18e aborted " & Err.Number & " " & Err.Description & "]": Err.Clear

    T18f_no_disk_write
    If Err.Number <> 0 Then Say "  [T18f aborted " & Err.Number & " " & Err.Description & "]": Err.Clear

    T18g_readonly_file
    If Err.Number <> 0 Then Say "  [T18g aborted " & Err.Number & " " & Err.Description & "]": Err.Clear

    T18h_builtin
    If Err.Number <> 0 Then Say "  [T18h aborted " & Err.Number & " " & Err.Description & "]": Err.Clear

    T18i_type_custom
    If Err.Number <> 0 Then Say "  [T18i aborted " & Err.Number & " " & Err.Description & "]": Err.Clear

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
' helper: make a clean, saved workbook at path p
'==============================================================================
Private Function MakeClean(ByVal p As String) As Workbook
    Dim wb As Workbook
    KillQ p
    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Range("A1").Value2 = "data"
    wb.SaveAs p, 51
    Set MakeClean = wb          ' object return values REQUIRE Set
End Function

'==============================================================================
' T18a -- add a custom property to a CLEAN workbook
'==============================================================================
Private Sub T18a_add_on_clean()
    Dim wb As Workbook
    Dim p As String
    Dim d1 As Boolean, d2 As Boolean

    Say ""
    Say "---- T18a: Add on a clean workbook -----------------------------"

    p = m_tmp & "\mddt_t18a.xlsx"
    Set wb = MakeClean(p)

    d1 = wb.Saved
    Say " Saved after SaveAs            : " & d1 & "   (clean expected True)"

    On Error Resume Next
    wb.CustomDocumentProperties.Add "MDDT_ROLE", False, 4, "LOCAL"
    If Err.Number <> 0 Then
        Say " Add FAILED " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    d2 = wb.Saved
    Say " Saved after Add               : " & d2
    Say " >>> ADD MAKES IT DIRTY?       : " & (Not d2)

    Say " read back                     : [" & ReadProp(wb) & "]"
    Say " Saved after read              : " & wb.Saved

    wb.Close False
    KillQ p
End Sub

'==============================================================================
' T18b -- change the VALUE of an existing property
'==============================================================================
Private Sub T18b_change_value()
    Dim wb As Workbook
    Dim p As String
    Dim d0 As Boolean, d1 As Boolean

    Say ""
    Say "---- T18b: change value of existing property on clean wb --------"

    p = m_tmp & "\mddt_t18b.xlsx"
    Set wb = MakeClean(p)

    On Error Resume Next
    wb.CustomDocumentProperties.Add "MDDT_ROLE", False, 4, "LOCAL"
    On Error GoTo 0
    wb.Save
    d0 = wb.Saved
    Say " clean after save              : " & d0

    On Error Resume Next
    wb.CustomDocumentProperties("MDDT_ROLE").Value = "LOCAL2"
    If Err.Number <> 0 Then
        Say " set value FAILED " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    d1 = wb.Saved
    Say " Saved after changing value    : " & d1
    Say " >>> CHANGE MAKES IT DIRTY?    : " & (Not d1)
    Say " read back                     : [" & ReadProp(wb) & "]"

    wb.Close False
    KillQ p
End Sub

'==============================================================================
' T18c -- reading a property must be inert
'==============================================================================
Private Sub T18c_read_only()
    Dim wb As Workbook
    Dim p As String
    Dim i As Long
    Dim s As String

    Say ""
    Say "---- T18c: reading properties is inert -------------------------"

    p = m_tmp & "\mddt_t18c.xlsx"
    Set wb = MakeClean(p)

    On Error Resume Next
    wb.CustomDocumentProperties.Add "MDDT_ROLE", False, 4, "LOCAL"
    On Error GoTo 0
    wb.Save
    Say " clean before reads            : " & wb.Saved

    On Error Resume Next
    For i = 1 To 200
        s = CStr(wb.CustomDocumentProperties("MDDT_ROLE").Value)
    Next i
    Err.Clear
    On Error GoTo 0

    Say " Saved after 200 reads         : " & wb.Saved
    Say " >>> READ MAKES IT DIRTY?      : " & (Not wb.Saved)

    ' also enumerate the collection
    On Error Resume Next
    Dim n As Long
    n = wb.CustomDocumentProperties.Count
    Say " CustomDocumentProperties count: " & n
    Err.Clear
    On Error GoTo 0
    Say " Saved after Count             : " & wb.Saved

    wb.Close False
    KillQ p
End Sub

'==============================================================================
' T18d -- delete a property
'==============================================================================
Private Sub T18d_delete()
    Dim wb As Workbook
    Dim p As String

    Say ""
    Say "---- T18d: delete a property -----------------------------------"

    p = m_tmp & "\mddt_t18d.xlsx"
    Set wb = MakeClean(p)

    On Error Resume Next
    wb.CustomDocumentProperties.Add "MDDT_ROLE", False, 4, "LOCAL"
    On Error GoTo 0
    wb.Save
    Say " clean before delete           : " & wb.Saved

    On Error Resume Next
    wb.CustomDocumentProperties("MDDT_ROLE").Delete
    If Err.Number <> 0 Then
        Say " delete FAILED " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    Say " Saved after delete            : " & wb.Saved
    Say " >>> DELETE MAKES IT DIRTY?    : " & (Not wb.Saved)

    ' Add when it already exists?
    On Error Resume Next
    wb.CustomDocumentProperties.Add "DUP", False, 4, "1"
    Err.Clear
    wb.CustomDocumentProperties.Add "DUP", False, 4, "2"
    If Err.Number <> 0 Then
        Say " Add duplicate FAILED          : " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        Say " Add duplicate SUCCEEDED       : (value now " & _
            CStr(wb.CustomDocumentProperties("DUP").Value) & ")"
    End If
    On Error GoTo 0

    wb.Close False
    KillQ p
End Sub

'==============================================================================
' T18e -- persistence through SaveAs / SaveCopyAs / reopen
'==============================================================================
Private Sub T18e_persist()
    Dim wb As Workbook, wb2 As Workbook
    Dim p As String, p2 As String, p3 As String

    Say ""
    Say "---- T18e: persistence -----------------------------------------"

    p = m_tmp & "\mddt_t18e.xlsx"
    p2 = m_tmp & "\mddt_t18e_sa.xlsx"
    p3 = m_tmp & "\mddt_t18e_sc.xlsx"
    KillQ p2
    KillQ p3

    Set wb = MakeClean(p)
    On Error Resume Next
    wb.CustomDocumentProperties.Add "MDDT_ROLE", False, 4, "LOCAL"
    On Error GoTo 0
    wb.Save

    Say " after save                    : [" & ReadProp(wb) & "]"
    Say " Saved                         : " & wb.Saved

    wb.SaveAs p2, 51
    Say " after SaveAs                  : [" & ReadProp(wb) & "]"
    Say " Saved after SaveAs            : " & wb.Saved

    wb.SaveCopyAs p3
    Say " SaveCopyAs done; Saved        : " & wb.Saved

    Set wb2 = Application.Workbooks.Open(p3, 0, True)
    Say " reopened SaveCopyAs copy      : [" & ReadProp(wb2) & "]"
    wb2.Close False

    wb.Close False

    Set wb2 = Application.Workbooks.Open(p2, 0, False)
    Say " reopened SaveAs result        : [" & ReadProp(wb2) & "]"
    Say " Saved on open (r/w)           : " & wb2.Saved
    wb2.Close False

    KillQ p
    KillQ p2
    KillQ p3
End Sub

'==============================================================================
' T18f -- does Add touch the file on disk before Save?
'==============================================================================
Private Sub T18f_no_disk_write()
    Dim wb As Workbook
    Dim p As String
    Dim sz1 As Double, sz2 As Double
    Dim t1 As Date, t2 As Date
    Dim f As Integer

    Say ""
    Say "---- T18f: does Add write to disk before Save? ------------------"

    p = m_tmp & "\mddt_t18f.xlsx"
    Set wb = MakeClean(p)

    With CreateObject("Scripting.FileSystemObject").GetFile(p)
        sz1 = .Size
        t1 = .DateLastModified
    End With
    Say " before Add: size=" & sz1 & "  mtime=" & Format$(t1, "hh:nn:ss")

    On Error Resume Next
    wb.CustomDocumentProperties.Add "MDDT_ROLE", False, 4, "LOCAL"
    Err.Clear
    On Error GoTo 0

    ' force the file system to report fresh metadata
    DoEvents
    With CreateObject("Scripting.FileSystemObject").GetFile(p)
        sz2 = .Size
        t2 = .DateLastModified
    End With
    Say " after Add : size=" & sz2 & "  mtime=" & Format$(t2, "hh:nn:ss")
    Say " >>> FILE CHANGED ON DISK?     : " & ((sz1 <> sz2) Or (t1 <> t2))
    Say "     (expected False: nothing is written until Save)"

    wb.Close False
    KillQ p
End Sub

'==============================================================================
' T18g -- adding a property to a workbook opened read-only
'==============================================================================
Private Sub T18g_readonly_file()
    Dim wb As Workbook
    Dim p As String

    Say ""
    Say "---- T18g: Add on a read-only opened workbook ------------------"

    p = m_tmp & "\mddt_t18g.xlsx"
    Set wb = MakeClean(p)
    wb.Close False

    Set wb = Application.Workbooks.Open(p, 0, True)    ' ReadOnly:=True
    Say " opened read-only, ReadOnly=" & wb.ReadOnly

    On Error Resume Next
    wb.CustomDocumentProperties.Add "MDDT_ROLE", False, 4, "LOCAL"
    If Err.Number <> 0 Then
        Say " Add on read-only FAILED       : " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        Say " Add on read-only SUCCEEDED    : (in-memory only)"
    End If
    On Error GoTo 0

    Say " Saved after Add (read-only)   : " & wb.Saved

    wb.Close False
    KillQ p
End Sub

'==============================================================================
' T18h -- compare with BUILTIN document properties
'==============================================================================
Private Sub T18h_builtin()
    Dim wb As Workbook
    Dim p As String

    Say ""
    Say "---- T18h: builtin property (used by earlier save tests) --------"

    p = m_tmp & "\mddt_t18h.xlsx"
    Set wb = MakeClean(p)
    Say " Saved after SaveAs            : " & wb.Saved

    On Error Resume Next
    wb.BuiltinDocumentProperties.Item("Comments").Value = "probe " & Now
    If Err.Number <> 0 Then
        Say " set Comments FAILED           : " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    Say " Saved after setting Comments  : " & wb.Saved
    Say " >>> BUILTIN MAKES IT DIRTY?   : " & (Not wb.Saved)

    wb.Close False
    KillQ p
End Sub

'==============================================================================
' T18i -- which custom property types are usable
'==============================================================================
Private Sub T18i_type_custom()
    Dim wb As Workbook
    Dim p As String
    Dim i As Long

    Say ""
    Say "---- T18i: custom property value types -------------------------"

    p = m_tmp & "\mddt_t18i.xlsx"
    Set wb = MakeClean(p)

    Dim types As Variant, names As Variant
    types = Array(2, 3, 4, 5, 1)      ' 2=Integer 3=Date 4=Text 5=Number 1=msoPropertyTypeNumber
    names = Array("P_INT", "P_DATE", "P_TEXT", "P_NUM", "P_NUM2")

    For i = 0 To UBound(types)
        On Error Resume Next
        Select Case types(i)
            Case 2: wb.CustomDocumentProperties.Add names(i), False, 2, 42
            Case 3: wb.CustomDocumentProperties.Add names(i), False, 3, Now
            Case 4: wb.CustomDocumentProperties.Add names(i), False, 4, "hello"
            Case 5: wb.CustomDocumentProperties.Add names(i), False, 5, 3.5
            Case 1: wb.CustomDocumentProperties.Add names(i), False, 1, 7
        End Select
        If Err.Number <> 0 Then
            Say " " & names(i) & " (type " & types(i) & ") FAILED " & Err.Number & _
                " " & Err.Description
            Err.Clear
        Else
            Say " " & names(i) & " (type " & types(i) & ") OK -> " & _
                CStr(wb.CustomDocumentProperties(names(i)).Value)
        End If
        On Error GoTo 0
    Next i

    wb.Save
    Dim wb2 As Workbook
    wb.Close False
    Set wb2 = Application.Workbooks.Open(p, 0, True)
    Say " after reopen:"
    For i = 0 To UBound(names)
        On Error Resume Next
        Say "   " & names(i) & " = " & CStr(wb2.CustomDocumentProperties(names(i)).Value)
        If Err.Number <> 0 Then
            Say "   " & names(i) & " = <read failed " & Err.Number & ">"
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    wb2.Close False
    KillQ p
End Sub

'==============================================================================
' helpers
'==============================================================================
Private Function ReadProp(ByVal wb As Workbook) As String
    Dim s As String
    On Error Resume Next
    s = CStr(wb.CustomDocumentProperties("MDDT_ROLE").Value)
    If Err.Number <> 0 Then
        s = "<absent>"
        Err.Clear
    End If
    On Error GoTo 0
    ReadProp = s
End Function

Private Sub KillQ(ByVal p As String)
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
    p = m_tmp & "\mddt_propdirty.txt"
    On Error Resume Next
    f = FreeFile
    Open p For Output As #f
    Print #f, m_log
    Close #f
    On Error GoTo 0
    Debug.Print "report: " & p
End Sub