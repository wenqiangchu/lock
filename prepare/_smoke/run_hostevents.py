"""Host-instance integration experiment.

Builds a REAL installed .xlam add-in (IsAddin=True), hooks Application-level
events from inside it, then makes the add-in itself open/close a candidate
workbook. Every event that fires is written to a log so we can see exactly
which handlers are triggered, and what happens to focus/visibility.

Outputs: %TEMP%\\mddt_hostevents.log
"""

import os
import sys
import time
import shutil
import subprocess

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(HERE, "hostev")
XLSM = os.path.join(WORK, "harness_build.xlsm")
XLAM = os.path.join(WORK, "mddt_addin.xlam")
LOG = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "mddt_hostevents.log")

# ---------------------------------------------------------------------------
# VBA sources (pure ASCII)
# ---------------------------------------------------------------------------

CLS_APP = """
Public WithEvents App As Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    LogEv "APP:WorkbookOpen", Wb.Name
End Sub
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    LogEv "APP:WorkbookBeforeClose", Wb.Name
End Sub
Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    LogEv "APP:WorkbookActivate", Wb.Name
End Sub
Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    LogEv "APP:WorkbookDeactivate", Wb.Name
End Sub
Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    LogEv "APP:WorkbookBeforeSave", Wb.Name
End Sub
Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    LogEv "APP:NewWorkbook", Wb.Name
End Sub
Private Sub App_SheetActivate(ByVal Sh As Object)
    LogEv "APP:SheetActivate", Sh.Name
End Sub
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    LogEv "APP:SheetChange", Sh.Name
End Sub
Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    LogEv "APP:SheetSelectionChange", Sh.Name
End Sub
Private Sub App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    LogEv "APP:WindowActivate", Wb.Name
End Sub
Private Sub App_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    LogEv "APP:WindowDeactivate", Wb.Name
End Sub
Private Sub App_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    LogEv "APP:WindowResize", Wb.Name
End Sub
"""

THISWORKBOOK = """
Private Sub Workbook_Open()
    LogEv "THISWB:Workbook_Open", Me.Name
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    LogEv "THISWB:Workbook_BeforeClose", Me.Name
End Sub
Private Sub Workbook_Activate()
    LogEv "THISWB:Workbook_Activate", Me.Name
End Sub
Private Sub Workbook_Deactivate()
    LogEv "THISWB:Workbook_Deactivate", Me.Name
End Sub
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    LogEv "THISWB:Workbook_WindowActivate", Me.Name
End Sub
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    LogEv "THISWB:Workbook_SheetActivate", Sh.Name
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    LogEv "THISWB:Workbook_BeforeSave", Me.Name
End Sub
"""

STD = """
Option Explicit

Public gApp As clsAppEvents
Private mSeq As Long
Private mLog As String
Public gCand As String

' ---- A Workbook_Open in a STANDARD module of an add-in. Kept on purpose:
' ---- we want to see whether Excel calls it when other workbooks open.
Public Sub Workbook_Open()
    LogEv "STD:Workbook_Open", "standard module"
End Sub

Public Sub LogSetPath(ByVal p As String)
    mLog = p
End Sub

Public Sub LogEv(ByVal kind As String, ByVal detail As String)
    Dim f As Integer
    Dim line As String
    mSeq = mSeq + 1
    On Error Resume Next
    f = FreeFile
    Open mLog For Append As #f
    line = Format$(mSeq, "0000") & "|" & kind & "|" & detail & _
        "|EE=" & Application.EnableEvents & _
        "|SU=" & Application.ScreenUpdating
    Print #f, line
    Close #f
    On Error GoTo 0
End Sub

Public Sub ZZCOMPILEPROBE()
    Dim n As Long
    n = 1
    Debug.Print "compile ok " & n
End Sub

Public Sub ZZTEST_Run()
    If Len(mLog) = 0 Then mLog = Environ$("TEMP") & "\mddt_hostevents.log"
    LogEv "===", "ZZTEST_Run BEGIN"
    LogEv "TEST", "excel=" & Application.Version & " build " & Application.Build
    LogEv "TEST", "addin_isaddin=" & ThisWorkbook.IsAddin
    LogEv "TEST", "workbooks_now=" & Workbooks.Count

    gCand = Environ$("TEMP") & "\\mddt_hostev_candidate.xlsx"
    On Error Resume Next
    If Len(Dir$(gCand)) > 0 Then Kill gCand
    On Error GoTo 0

    Set gApp = New clsAppEvents
    Set gApp.App = Application
    LogEv "TEST", "app events HOOKED"

    MakeCandidate

    ScenarioA_open_events_on
    ScenarioB_open_events_off
    ScenarioC_hidden_and_copy
    ScenarioD_active_state
    ScenarioE_interactive

    Set gApp = Nothing
    LogEv "TEST", "app events UNHOOKED"
    LogEv "===", "ZZTEST_Run END"
End Sub

Private Sub MakeCandidate()
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Name = "AE"
    wb.Worksheets(1).Range("A1").Value2 = "col1"
    wb.Worksheets(1).Range("A2").Value2 = "r2"
    wb.Worksheets(1).Range("A2").Interior.Color = RGB(255, 255, 153)
    wb.Worksheets.Add After:=wb.Worksheets(1)
    wb.Worksheets(2).Name = "LB"
    wb.SaveAs gCand, 51
    wb.Close False
    LogEv "SETUP", "candidate created: " & gCand
End Sub

' --- A: plain open/close with events ON -------------------------------
Private Sub ScenarioA_open_events_on()
    Dim wb As Workbook
    LogEv "A", "--- open candidate with EnableEvents=TRUE ---"
    Set wb = Application.Workbooks.Open(gCand, 0, False)
    LogEv "A", "opened; windows=" & wb.Windows.Count & " visible=" & wb.Windows(1).Visible
    LogEv "A", "activewb_now=" & Application.ActiveWorkbook.Name
    wb.Close False
    Set wb = Nothing
    LogEv "A", "closed"
End Sub

' --- B: open/close with events OFF ------------------------------------
Private Sub ScenarioB_open_events_off()
    Dim wb As Workbook
    LogEv "B", "--- open candidate with EnableEvents=FALSE ---"
    Application.EnableEvents = False
    Set wb = Application.Workbooks.Open(gCand, 0, False)
    LogEv "B", "opened (events off)"
    wb.Close False
    Set wb = Nothing
    Application.EnableEvents = True
    LogEv "B", "closed, events restored"
End Sub

' --- C: hidden windows + Range.Copy -----------------------------------
Private Sub ScenarioC_hidden_and_copy()
    Dim loc As Workbook, cand As Workbook
    Dim r As String

    LogEv "C", "--- hidden windows + copy ---"
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set loc = Application.Workbooks.Add
    loc.Worksheets(1).Range("A1").Value2 = "LOCAL1"
    loc.Worksheets(1).Range("A2").Value2 = "LOCAL2"
    loc.Worksheets(1).Range("A2").Interior.Color = RGB(153, 204, 255)
    loc.Windows(1).Visible = False
    LogEv "C", "local copy hidden; windows=" & loc.Windows.Count & _
        " visible=" & loc.Windows(1).Visible

    Set cand = Application.Workbooks.Open(gCand, 0, False)
    LogEv "C", "candidate opened; visible=" & cand.Windows(1).Visible
    cand.Windows(1).Visible = False
    LogEv "C", "candidate hidden; visible=" & cand.Windows(1).Visible

    ' copy hidden -> hidden
    On Error Resume Next
    loc.Worksheets(1).Range("A1:A2").Copy Destination:=cand.Worksheets(1).Range("B1")
    If Err.Number <> 0 Then
        r = "FAIL " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        r = "OK"
    End If
    On Error GoTo 0
    LogEv "C", "Copy hidden->hidden: " & r & _
        "   B1=" & CStr(cand.Worksheets(1).Range("B1").Value2) & _
        "   B2=" & CStr(cand.Worksheets(1).Range("B2").Value2) & _
        "   B2fill=" & cand.Worksheets(1).Range("B2").Interior.Color

    ' save the hidden target
    On Error Resume Next
    cand.Save
    If Err.Number <> 0 Then
        LogEv "C", "Save on hidden workbook: FAIL " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        LogEv "C", "Save on hidden workbook: OK"
    End If
    On Error GoTo 0

    ' clear while hidden
    On Error Resume Next
    cand.Worksheets(1).Range("B1:B3").Clear
    If Err.Number <> 0 Then
        LogEv "C", "Clear on hidden workbook: FAIL " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        LogEv "C", "Clear on hidden workbook: OK"
    End If
    On Error GoTo 0

    cand.Close False
    loc.Close False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    LogEv "C", "--- done ---"
End Sub

' --- D: does the user's active workbook / sheet / selection survive? ---
Private Sub ScenarioD_active_state()
    Dim userWb As Workbook, cand As Workbook
    Dim beforeWb As String, beforeSh As String, beforeSel As String
    Dim midWb As String, midSh As String

    LogEv "D", "--- user state preservation ---"
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set userWb = Application.Workbooks.Add
    userWb.Worksheets(1).Range("D5").Value2 = "user cell"
    userWb.Activate
    userWb.Worksheets(1).Range("D5").Select
    beforeWb = Application.ActiveWorkbook.Name
    beforeSh = Application.ActiveSheet.Name
    beforeSel = Application.Selection.Address(False, False)
    LogEv "D", "BEFORE wb=" & beforeWb & " sheet=" & beforeSh & " sel=" & beforeSel

    Set cand = Application.Workbooks.Open(gCand, 0, False)
    cand.Windows(1).Visible = False
    midWb = "?"
    midSh = "?"
    On Error Resume Next
    midWb = Application.ActiveWorkbook.Name
    midSh = Application.ActiveSheet.Name
    On Error GoTo 0
    LogEv "D", "MID (after open+hide) wb=" & midWb & " sheet=" & midSh

    cand.Close False
    LogEv "D", "AFTER close: activewb=" & Application.ActiveWorkbook.Name & _
        " activesheet=" & Application.ActiveSheet.Name

    userWb.Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    LogEv "D", "restored: wb=" & Application.ActiveWorkbook.Name & _
        " sheet=" & Application.ActiveSheet.Name

    userWb.Close False
    LogEv "D", "--- done ---"
End Sub

' --- E: Application.Interactive ---------------------------------------
Private Sub ScenarioE_interactive()
    Dim cand As Workbook
    LogEv "E", "--- Application.Interactive probe ---"
    LogEv "E", "interactive_before=" & Application.Interactive
    Application.Interactive = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error Resume Next
    Set cand = Application.Workbooks.Open(gCand, 0, False)
    If Err.Number <> 0 Then
        LogEv "E", "open with Interactive=False FAILED " & Err.Number & " " & Err.Description
        Err.Clear
        GoTo Done
    End If
    LogEv "E", "opened with Interactive=False; visible=" & cand.Windows(1).Visible
    cand.Windows(1).Visible = False
    cand.Close False
Done:
    On Error GoTo 0
    Application.Interactive = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    LogEv "E", "interactive_after=" & Application.Interactive
End Sub
"""


def sh(*args):
    subprocess.run(args, capture_output=True)


def kill_excel():
    sh("powershell", "-NoProfile", "-Command",
       "Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force")


def main():
    if os.path.isdir(WORK):
        for _ in range(5):
            try:
                shutil.rmtree(WORK)
                break
            except Exception:
                time.sleep(1)
    os.makedirs(WORK, exist_ok=True)

    kill_excel()
    time.sleep(1)
    if os.path.exists(LOG):
        os.remove(LOG)

    print("building add-in ...", flush=True)
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.EnableEvents = False
    app.ScreenUpdating = False
    app.AutomationSecurity = 1

    ai = None
    try:
        wb = app.Workbooks.Add()
        comps = wb.VBProject.VBComponents

        cls = comps.Add(2)          # vbext_ct_ClassModule
        cls.Name = "clsAppEvents"
        cls.CodeModule.AddFromString(CLS_APP)

        std = comps.Add(1)          # vbext_ct_StdModule
        std.Name = "MdlHostTest"
        std.CodeModule.AddFromString(STD)

        comps("ThisWorkbook").CodeModule.AddFromString(THISWORKBOOK)

        wb.SaveAs(XLSM, 52)
        wb.SaveAs(XLAM, 55)         # xlAddIn
        wb.Close(False)
        print("add-in built: %s" % XLAM, flush=True)

        # Open the .xlam directly. Excel loads it as a real add-in
        # (IsAddin = True) and fires its ThisWorkbook Workbook_Open, which is
        # exactly the situation a warm host instance has in production.
        ai = app.Workbooks.Open(XLAM)
        print("add-in loaded: %s  IsAddin=%s" % (ai.Name, ai.IsAddin), flush=True)
        time.sleep(1)

        print("running ZZTEST_Run inside the add-in ...", flush=True)
        ok = False
        for name in ("mddt_addin.xlam!ZZTEST_Run", ai.Name + "!ZZTEST_Run", "ZZTEST_Run"):
            try:
                app.Run(name)
                print("ran as: %s" % name, flush=True)
                ok = True
                break
            except Exception as e:
                print("  %s -> %s" % (name, e), flush=True)
        if not ok:
            print("COULD NOT RUN THE TEST", flush=True)
    finally:
        try:
            if ai is not None:
                ai.Close(False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        del app
        kill_excel()
        time.sleep(1)

    if os.path.exists(LOG):
        print("\n================ EVENT LOG ================", flush=True)
        with open(LOG, "r", encoding="utf-8", errors="replace") as f:
            sys.stdout.write(f.read())
    else:
        print("no log at %s" % LOG, flush=True)


if __name__ == "__main__":
    main()