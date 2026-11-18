"""Role identification + role-scoped event handling.

Answers, by experiment:

  R1  Can we tell "this is the add-in itself" from inside an Application handler?
  R2  Is object identity (Wb Is mLocal) reliable, and role-discriminating?
  R3  Does a WithEvents class module attached to ONE workbook fire only for that
      workbook, or for all of them?
  R4  What happens when we compare against a reference whose workbook was closed?
      (stale reference -> error? silently False?)
  R5  Can a custom document property act as a role marker?
      - does it dirty the workbook?
      - does it survive SaveAs?
      - does it survive SaveCopyAs?
  R6  Can the attached class module Cancel a close?

Run: python _smoke/run_roles3.py
Log: %TEMP%\\mddt_roles3.log
"""

import os
import sys
import time
import shutil
import subprocess

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(HERE, "roles3")
XLSM = os.path.join(WORK, "build.xlsm")
XLAM = os.path.join(WORK, "mddt3.xlam")
LOG = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "mddt_roles3.log")

CLS_APP = r"""
Public WithEvents App As Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    LogEv "APP:Open", Dsc(Wb)
End Sub
Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    LogEv "APP:Activate", Dsc(Wb)
End Sub
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    LogEv "APP:BeforeClose", Dsc(Wb)
End Sub
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    LogEv "APP:SheetChange", Sh.Parent.Name & "/" & Sh.Name
End Sub
"""

CLS_WB = r"""
Public WithEvents Wb As Workbook
Public Tag As String

Private Sub Wb_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    LogEv "WB[" & Tag & "]:SheetChange", Sh.Name
End Sub
Private Sub Wb_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    LogEv "WB[" & Tag & "]:BeforeSave", Tag
End Sub
Private Sub Wb_BeforeClose(Cancel As Boolean)
    LogEv "WB[" & Tag & "]:BeforeClose", Tag
    If gCancelClose Then
        Cancel = True
        LogEv "WB[" & Tag & "]:BeforeClose", "Cancel set to True"
    End If
End Sub
Private Sub Wb_Activate()
    LogEv "WB[" & Tag & "]:Activate", Tag
End Sub
Private Sub Wb_Deactivate()
    LogEv "WB[" & Tag & "]:Deactivate", Tag
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
"""

STD = r"""
Option Explicit

Public gApp As clsAppEvents
Public mLocal As Workbook
Public mOther As Workbook
Public mEvLocal As clsWbEvents
Public mEvOther As clsWbEvents
Public gCancelClose As Boolean
Public gStale As Workbook

Private mSeq As Long
Private mLog As String

Public Sub LogEv(ByVal kind As String, ByVal detail As String)
    Dim f As Integer
    Dim line As String
    If Len(mLog) = 0 Then mLog = Environ$("TEMP") & "\mddt_roles3.log"
    mSeq = mSeq + 1
    On Error Resume Next
    f = FreeFile
    Open mLog For Append As #f
    line = Format$(mSeq, "0000") & "|" & kind & "|" & detail & _
        "|EE=" & Application.EnableEvents
    Print #f, line
    Close #f
    On Error GoTo 0
End Sub

' Describe a workbook for logging: role guess + identity tests
Public Function Dsc(ByVal wb As Workbook) As String
    Dim s As String
    On Error Resume Next
    s = wb.Name
    If Err.Number <> 0 Then
        Dsc = "<invalid reference>"
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    On Error Resume Next
    If wb Is ThisWorkbook Then s = s & " [IsAddinSelf=TRUE]"
    If Not mLocal Is Nothing Then
        If wb Is mLocal Then s = s & " [IsLocal=TRUE]"
    End If
    If Not mOther Is Nothing Then
        If wb Is mOther Then s = s & " [IsOther=TRUE]"
    End If
    If Err.Number <> 0 Then
        s = s & " [Is-compare ERROR " & Err.Number & "]"
        Err.Clear
    End If
    On Error GoTo 0
    Dsc = s
End Function

Public Sub ZZCOMPILEPROBE()
    Dim n As Long
    n = 1
    Debug.Print "compile ok " & n
End Sub

Public Sub ZZTEST_Run()
    If Len(mLog) = 0 Then mLog = Environ$("TEMP") & "\mddt_roles3.log"
    LogEv "===", "BEGIN"
    LogEv "SELF", "ThisWorkbook.Name=" & ThisWorkbook.Name & _
        " IsAddin=" & ThisWorkbook.IsAddin

    Set gApp = New clsAppEvents
    Set gApp.App = Application

    R1_self_identity
    R2_R3_scoped_events
    R4_stale_reference
    R5_custom_property
    R6_cancel_close

    Set gApp = Nothing
    LogEv "===", "END"
End Sub

' ---- R1: can we identify the add-in from inside an Application handler? ----
Private Sub R1_self_identity()
    Dim n As Long, hit As Long
    LogEv "---", "R1: add-in self identity"
    ' The add-in workbook is visible in Workbooks collection
    n = Application.Workbooks.Count
    LogEv "R1", "workbooks=" & n
    For n = 1 To Application.Workbooks.Count
        If Application.Workbooks(n) Is ThisWorkbook Then
            hit = hit + 1
            LogEv "R1", "Workbooks(" & n & ") Is ThisWorkbook -> TRUE (" & _
                Application.Workbooks(n).Name & ")"
        End If
    Next n
    LogEv "R1", "matches found = " & hit
End Sub

' ---- R2/R3: does WithEvents on one workbook fire for the other? ----
Private Sub R2_R3_scoped_events()
    Dim pA As String, pB As String

    LogEv "---", "R2/R3: role-scoped WithEvents"

    pA = Environ$("TEMP") & "\mddt_r3_local.xlsx"
    pB = Environ$("TEMP") & "\mddt_r3_other.xlsx"
    KillQ pA
    KillQ pB

    Application.EnableEvents = True

    Set mLocal = Application.Workbooks.Add
    mLocal.SaveAs pA, 51

    Set mOther = Application.Workbooks.Add
    mOther.SaveAs pB, 51
    LogEv "R2", "mLocal=" & Dsc(mLocal)
    LogEv "R2", "mOther=" & Dsc(mOther)

    Set mEvLocal = New clsWbEvents
    mEvLocal.Tag = "LOCAL"
    Set mEvLocal.Wb = mLocal

    Set mEvOther = New clsWbEvents
    mEvOther.Tag = "OTHER"
    Set mEvOther.Wb = mOther

    LogEv "R3", ">>> write to mLocal (expect WB[LOCAL] only)"
    mLocal.Worksheets(1).Range("A1").Value2 = "x"

    LogEv "R3", ">>> write to mOther (expect WB[OTHER] only)"
    mOther.Worksheets(1).Range("A1").Value2 = "y"

    LogEv "R3", ">>> activate mOther"
    mOther.Activate
    LogEv "R3", ">>> activate mLocal"
    mLocal.Activate
End Sub

' ---- R4: stale reference behaviour ----
Private Sub R4_stale_reference()
    Dim wb3 As Workbook
    Dim pC As String
    Dim r As String

    LogEv "---", "R4: stale (closed) workbook reference"

    pC = Environ$("TEMP") & "\mddt_r3_stale.xlsx"
    KillQ pC

    Set wb3 = Application.Workbooks.Add
    wb3.SaveAs pC, 51
    Set gStale = wb3

    LogEv "R4", "before close: gStale.Name=" & SafeName(gStale)
    wb3.Close False

    ' reference is now stale. What do the usual tests do?
    LogEv "R4", "after close: gStale Is Nothing -> " & (gStale Is Nothing)
    LogEv "R4", "after close: SafeName(gStale)=" & SafeName(gStale)

    On Error Resume Next
    Dim b As Boolean
    b = (gStale Is mOther)
    If Err.Number <> 0 Then
        LogEv "R4", "gStale Is mOther -> ERROR " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        LogEv "R4", "gStale Is mOther -> " & b
    End If
    On Error GoTo 0

    ' the practical question: after a workbook closes, can we compare a LIVE
    ' incoming reference against our stored slot without an error?
    On Error Resume Next
    b = (mOther Is gStale)
    If Err.Number <> 0 Then
        LogEv "R4", "mOther Is gStale -> ERROR " & Err.Number
        Err.Clear
    Else
        LogEv "R4", "mOther Is gStale -> " & b
    End If
    On Error GoTo 0

    ' and can we still set it to Nothing safely?
    On Error Resume Next
    Set gStale = Nothing
    If Err.Number <> 0 Then
        LogEv "R4", "Set gStale = Nothing -> ERROR " & Err.Number
        Err.Clear
    Else
        LogEv "R4", "Set gStale = Nothing -> OK"
    End If
    On Error GoTo 0

    KillQ pC
End Sub

Private Function SafeName(ByVal wb As Workbook) As String
    Dim s As String
    On Error Resume Next
    s = wb.Name
    If Err.Number <> 0 Then
        s = "<ERROR " & Err.Number & ">"
        Err.Clear
    End If
    On Error GoTo 0
    SafeName = s
End Function

' ---- R5: custom document property as a role marker ----
Private Sub R5_custom_property()
    Dim p1 As String, p2 As String, p3 As String
    Dim v As String
    Dim wasSaved As Boolean

    LogEv "---", "R5: custom document property as role marker"

    p1 = Environ$("TEMP") & "\mddt_r5_a.xlsx"
    p2 = Environ$("TEMP") & "\mddt_r5_b.xlsx"     ' after SaveAs
    p3 = Environ$("TEMP") & "\mddt_r5_c.xlsx"     ' after SaveCopyAs
    KillQ p1
    KillQ p2
    KillQ p3

    Set mLocal = Application.Workbooks.Add
    mLocal.SaveAs p1, 51
    mLocal.Worksheets(1).Range("A1").Value2 = "data"

    wasSaved = mLocal.Saved
    LogEv "R5", "Saved before adding property = " & wasSaved

    On Error Resume Next
    mLocal.CustomDocumentProperties.Add "MDDT_ROLE", False, 4, "LOCAL"
    If Err.Number <> 0 Then
        LogEv "R5", "CustomDocumentProperties.Add FAILED " & Err.Number & _
            " " & Err.Description
        Err.Clear
    Else
        LogEv "R5", "CustomDocumentProperties.Add OK"
    End If
    On Error GoTo 0

    LogEv "R5", "Saved after adding property = " & mLocal.Saved & _
        "  (dirty if False)"

    v = ReadProp(mLocal)
    LogEv "R5", "read back = [" & v & "]"

    ' SaveAs -> does the property survive?
    mLocal.SaveAs p2, 51
    LogEv "R5", "after SaveAs: read = [" & ReadProp(mLocal) & "]"

    ' SaveCopyAs -> does the property survive in the copy?
    mLocal.SaveCopyAs p3
    Dim wbc As Workbook
    Set wbc = Application.Workbooks.Open(p3, 0, True)
    LogEv "R5", "after SaveCopyAs, reopened copy: read = [" & ReadProp(wbc) & "]"
    wbc.Close False

    ' plain copy of a range does NOT create a workbook, so skip that.
    ' Does a NEW workbook have the property? (must be absent)
    LogEv "R5", "fresh workbook   : read = [" & ReadProp(Application.Workbooks.Add) & "]"
    Application.Workbooks(Application.Workbooks.Count).Close False

    mLocal.Close False
    KillQ p1
    KillQ p2
    KillQ p3
End Sub

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

' ---- R6: can the attached class module cancel a close? ----
Private Sub R6_cancel_close()
    Dim pD As String
    Dim wb4 As Workbook

    LogEv "---", "R6: cancel a close through the attached class"

    pD = Environ$("TEMP") & "\mddt_r6.xlsx"
    KillQ pD

    Set wb4 = Application.Workbooks.Add
    wb4.SaveAs pD, 51

    Dim ev As clsWbEvents
    Set ev = New clsWbEvents
    ev.Tag = "R6"
    Set ev.Wb = wb4

    Application.EnableEvents = True
    gCancelClose = True
    On Error Resume Next
    wb4.Close False
    If Err.Number <> 0 Then
        LogEv "R6", "Close with Cancel=True raised " & Err.Number & " " & _
            Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    LogEv "R6", "after attempted close: workbooks count = " & _
        Application.Workbooks.Count
    LogEv "R6", "is wb4 still open? " & NameStillOpen(pD)

    gCancelClose = False
    On Error Resume Next
    wb4.Close False
    On Error GoTo 0
    LogEv "R6", "after real close: workbooks count = " & _
        Application.Workbooks.Count
    Set ev = Nothing
    KillQ pD
End Sub

Private Function NameStillOpen(ByVal p As String) As Boolean
    Dim i As Long
    NameStillOpen = False
    For i = 1 To Application.Workbooks.Count
        If StrComp(Application.Workbooks(i).FullName, p, vbTextCompare) = 0 Then
            NameStillOpen = True
            Exit Function
        End If
    Next i
End Function

Private Sub KillQ(ByVal p As String)
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    On Error GoTo 0
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
    app.AutomationSecurity = 1
    app.EnableEvents = True
    app.ScreenUpdating = True

    ai = None
    try:
        wb = app.Workbooks.Add()
        comps = wb.VBProject.VBComponents

        cls = comps.Add(2)
        cls.Name = "clsAppEvents"
        cls.CodeModule.AddFromString(CLS_APP)

        cls2 = comps.Add(2)
        cls2.Name = "clsWbEvents"
        cls2.CodeModule.AddFromString(CLS_WB)

        std = comps.Add(1)
        std.Name = "MdlRoles"
        std.CodeModule.AddFromString(STD)

        comps("ThisWorkbook").CodeModule.AddFromString(THISWORKBOOK)

        wb.SaveAs(XLSM, 52)
        wb.SaveAs(XLAM, 55)
        wb.Close(False)
        print("add-in built", flush=True)

        with open(LOG, "w", encoding="utf-8") as f:
            f.write("0000|PHASE|before opening the .xlam\n")

        ai = app.Workbooks.Open(XLAM)
        print("add-in loaded: %s IsAddin=%s" % (ai.Name, ai.IsAddin), flush=True)

        time.sleep(1)
        print("running ZZTEST_Run ...", flush=True)
        app.Run("mddt3.xlam!ZZTEST_Run")
        print("returned", flush=True)
    finally:
        try:
            app.Run("mddt3.xlam!LogEv", "PHASE", "closing add-in")
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
        print("\n================ LOG ================", flush=True)
        with open(LOG, "r", encoding="utf-8", errors="replace") as f:
            sys.stdout.write(f.read())
    else:
        print("no log", flush=True)


if __name__ == "__main__":
    main()