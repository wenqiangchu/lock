"""Host-instance experiment, round 2 - event enumeration done correctly.

Round 1 bugs fixed:
  * Scenario A claimed EnableEvents=TRUE but never set it -> no events recorded.
  * The .xlam was loaded with EnableEvents=False, so its own Workbook_Open was
    suppressed and we could not tell whether a standard-module Workbook_Open is
    ever called.

This round answers, precisely:
  Q1  When a real installed .xlam is loaded with events ON, which of its own
      entry points run?  (ThisWorkbook.Workbook_Open vs a standard-module
      Workbook_Open vs Auto_Open)
  Q2  When the add-in opens a candidate workbook with events ON, exactly which
      Application-level events fire, in what order?
  Q3  With EnableEvents=False, do they all disappear?
  Q4  Does Application.Interactive=False auto-hide newly opened workbooks?
  Q5  Can the whole commit run without disturbing the user's active
      workbook / window / sheet / selection?
"""

import os
import sys
import time
import shutil
import subprocess

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(HERE, "hostev2")
XLSM = os.path.join(WORK, "build.xlsm")
XLAM = os.path.join(WORK, "mddt2.xlam")
LOG = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "mddt_hostev2.log")

CLS_APP = r"""
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
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    LogEv "APP:WorkbookAfterSave", Wb.Name
End Sub
Private Sub App_SheetActivate(ByVal Sh As Object)
    LogEv "APP:SheetActivate", Sh.Name
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
"""

THISWORKBOOK = r"""
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
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    LogEv "THISWB:Workbook_BeforeSave", Me.Name
End Sub
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    LogEv "THISWB:Workbook_WindowActivate", Me.Name
End Sub
"""

STD = r"""
Option Explicit

Public gApp As clsAppEvents
Private mSeq As Long
Private mLog As String
Public gCand As String

' ---------------------------------------------------------------------------
' These are the three entry points people mistakenly put in standard modules.
' Excel is known to call some of them for add-ins. We want the truth.
' ---------------------------------------------------------------------------
Public Sub Workbook_Open()
    LogEv "STD:Workbook_Open", "standard module"
End Sub

Public Sub Auto_Open()
    LogEv "STD:Auto_Open", "standard module"
End Sub

Public Sub Auto_Close()
    LogEv "STD:Auto_Close", "standard module"
End Sub

Public Sub LogEv(ByVal kind As String, ByVal detail As String)
    Dim f As Integer
    Dim line As String
    If Len(mLog) = 0 Then mLog = Environ$("TEMP") & "\mddt_hostev2.log"
    mSeq = mSeq + 1
    On Error Resume Next
    f = FreeFile
    Open mLog For Append As #f
    line = Format$(mSeq, "0000") & "|" & kind & "|" & detail & _
        "|EE=" & Application.EnableEvents & _
        "|SU=" & Application.ScreenUpdating & _
        "|IA=" & Application.Interactive
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
    If Len(mLog) = 0 Then mLog = Environ$("TEMP") & "\mddt_hostev2.log"
    LogEv "===", "BEGIN"

    gCand = Environ$("TEMP") & "\mddt_h2_candidate.xlsx"
    On Error Resume Next
    If Len(Dir$(gCand)) > 0 Then Kill gCand
    On Error GoTo 0

    Set gApp = New clsAppEvents
    Set gApp.App = Application

    MakeCandidate
    Q2_events_on
    Q3_events_off
    Q4_interactive
    Q5_user_state
    Q5b_aggressive

    Set gApp = Nothing
    LogEv "===", "END"
End Sub

Private Sub MakeCandidate()
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Name = "AE"
    wb.Worksheets(1).Range("A1").Value2 = "c1"
    wb.Worksheets(1).Range("A2").Value2 = "c2"
    wb.Worksheets(1).Range("A2").Interior.Color = RGB(255, 255, 153)
    wb.SaveAs gCand, 51
    wb.Close False
End Sub

' ---- Q2: events ON ---------------------------------------------------
Private Sub Q2_events_on()
    Dim wb As Workbook

    LogEv "---", "Q2: marker: events will be turned ON now"
    Application.EnableEvents = True
    LogEv "Q2", "EnableEvents set to TRUE"

    Set wb = Application.Workbooks.Open(gCand, 0, False)
    LogEv "Q2", "OPEN returned; activewb=" & SafeStr("wb") & _
        "  activewindow=" & SafeStr("win") & "  wbc=" & Application.Workbooks.Count
    wb.Close False
    LogEv "Q2", "CLOSE returned; wbc=" & Application.Workbooks.Count & _
        "  activewb=" & SafeStr("wb")

    Application.EnableEvents = False
    LogEv "Q2", "EnableEvents back to FALSE"
End Sub

' ---- Q3: events OFF --------------------------------------------------
Private Sub Q3_events_off()
    Dim wb As Workbook
    LogEv "---", "Q3: same cycle with EnableEvents=FALSE"
    Application.EnableEvents = False
    Set wb = Application.Workbooks.Open(gCand, 0, False)
    LogEv "Q3", "OPEN returned"
    wb.Close False
    LogEv "Q3", "CLOSE returned"
End Sub

' ---- Q4: Application.Interactive -------------------------------------
Private Sub Q4_interactive()
    Dim wb As Workbook

    LogEv "---", "Q4: does Interactive=False auto-hide a new window?"
    Application.EnableEvents = False
    Application.ScreenUpdating = True

    Application.Interactive = False
    Set wb = Application.Workbooks.Open(gCand, 0, False)
    LogEv "Q4", "Interactive=False -> windows=" & wb.Windows.Count & _
        "  visible=" & wb.Windows(1).Visible & _
        "  interactive=" & Application.Interactive

    Application.Interactive = True
    LogEv "Q4", "Interactive back to TRUE -> visible=" & wb.Windows(1).Visible

    Application.ScreenUpdating = False
    wb.Close False
End Sub

' ---- Q5: user state, gentle ------------------------------------------
Private Sub Q5_user_state()
    Dim userWb As Workbook, cand As Workbook
    Dim sBefore As String

    LogEv "---", "Q5: user state with ScreenUpdating=False + hide"
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set userWb = Application.Workbooks.Add
    userWb.Worksheets(1).Range("D5").Value2 = "usercell"
    userWb.Activate
    userWb.Worksheets(1).Range("D5").Select
    sBefore = Snap("BEFORE")

    Set cand = Application.Workbooks.Open(gCand, 0, False)
    Snap "MID-open"

    cand.Windows(1).Visible = False
    Snap "MID-hidden"

    cand.Close False
    Snap "AFTER-close"

    ' put the user exactly back
    On Error Resume Next
    userWb.Activate
    userWb.Worksheets(1).Range("D5").Select
    On Error GoTo 0
    Snap "RESTORED"

    userWb.Close False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ---- Q5b: the aggressive variant -------------------------------------
Private Sub Q5b_aggressive()
    Dim userWb As Workbook, cand As Workbook

    LogEv "---", "Q5b: Interactive=False + ScreenUpdating=False + hide"
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set userWb = Application.Workbooks.Add
    userWb.Worksheets(1).Range("D5").Value2 = "usercell2"
    userWb.Activate
    userWb.Worksheets(1).Range("D5").Select
    Snap "BEFORE"

    Application.Interactive = False
    Set cand = Application.Workbooks.Open(gCand, 0, False)
    Snap "MID-open"
    LogEv "Q5b", "visible=" & cand.Windows(1).Visible
    cand.Windows(1).Visible = False
    Snap "MID-hidden"
    cand.Close False
    Snap "AFTER-close"
    Application.Interactive = True

    On Error Resume Next
    userWb.Activate
    userWb.Worksheets(1).Range("D5").Select
    On Error GoTo 0
    Snap "RESTORED"

    userWb.Close False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Function SafeStr(ByVal what As String) As String
    Dim r As String
    On Error Resume Next
    Select Case what
        Case "wb":  r = Application.ActiveWorkbook.Name
        Case "win": r = Application.ActiveWindow.Caption
        Case "sh":  r = Application.ActiveSheet.Name
        Case "sel": r = CStr(Application.Selection.Address(False, False))
    End Select
    If Err.Number <> 0 Or Len(r) = 0 Then
        r = "<none>"
        Err.Clear
    End If
    On Error GoTo 0
    SafeStr = r
End Function

Private Function Snap(ByVal tag As String) As String
    Dim s As String
    On Error Resume Next
    s = "wb=" & SafeStr("wb") & _
        " sh=" & SafeStr("sh") & _
        " win=" & SafeStr("win") & _
        " sel=" & SafeStr("sel") & _
        " wbc=" & Application.Workbooks.Count & _
        " EE=" & Application.EnableEvents & _
        " SU=" & Application.ScreenUpdating & _
        " IA=" & Application.Interactive
    If Err.Number <> 0 Then
        s = "wb=<err " & Err.Number & ">"
        Err.Clear
    End If
    On Error GoTo 0
    LogEv "SNAP:" & tag, s
    Snap = s
End Function
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
    app.AutomationSecurity = 1
    # IMPORTANT: events ON, so the add-in's own Workbook_Open can fire
    app.EnableEvents = True
    app.ScreenUpdating = True

    ai = None
    try:
        wb = app.Workbooks.Add()
        comps = wb.VBProject.VBComponents

        cls = comps.Add(2)
        cls.Name = "clsAppEvents"
        cls.CodeModule.AddFromString(CLS_APP)

        std = comps.Add(1)
        std.Name = "MdlHostTest"
        std.CodeModule.AddFromString(STD)

        comps("ThisWorkbook").CodeModule.AddFromString(THISWORKBOOK)

        wb.SaveAs(XLSM, 52)
        wb.SaveAs(XLAM, 55)
        wb.Close(False)
        print("add-in built: %s" % XLAM, flush=True)

        # Seed the log so that anything that fires during add-in load is caught.
        with open(LOG, "w", encoding="utf-8") as f:
            f.write("0000|PHASE|before opening the .xlam\n")

        ai = app.Workbooks.Open(XLAM)
        print("add-in loaded: %s  IsAddin=%s" % (ai.Name, ai.IsAddin), flush=True)

        with open(LOG, "a", encoding="utf-8") as f:
            f.write("0000|PHASE|after opening the .xlam\n")

        time.sleep(1)
        print("running ZZTEST_Run ...", flush=True)
        app.Run("mddt2.xlam!ZZTEST_Run")
        print("ZZTEST_Run returned", flush=True)
    finally:
        try:
            app.Run("mddt2.xlam!LogEv", "PHASE", "closing the .xlam")
        except Exception:
            pass
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