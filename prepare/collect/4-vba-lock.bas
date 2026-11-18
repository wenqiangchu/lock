Attribute VB_Name = "TestSmbLock"
' ============================================================
'  4-vba-lock.bas  -  does VBA "Open ... Lock" participate in SMB exclusion?
'
'  PowerShell side already proved the server behaviour with Win32 CreateFile.
'  This module proves the native VBA Open/Lock statement is perceived by the
'  other machine too, which is what owners.tsv locking relies on.
'
'  Usage:
'    Side A (you): import this module into any workbook, put the cursor in
'                  HoldLockVBA, press F5. It holds exclusive access for 90 s.
'    Side B (colleague): on another machine run
'                  .\4-test-lock.ps1 -Path "\\SERVER\SHARE\_exs_probe_me\vba.lock" -Mode Try
'
'  Expected: side B reports SHARING_VIOLATION(32).
'
'  SelfTestMutex runs on one machine only and checks local semantics.
'
'  NOTE: keep this file pure ASCII.
' ============================================================
Option Explicit

' Change this to your probe path before using HoldLockVBA / TryLockVBA.
Private Const LOCK_PATH As String = "\\SERVER\SHARE\_exs_probe_me\vba.lock"

' Set to False to suppress MsgBox (useful for unattended runs).
Private Const SHOW_MSGBOX As Boolean = True

' ------------------------------------------------------------
' Hold exclusive access for N seconds.
' ------------------------------------------------------------
Public Sub HoldLockVBA(Optional ByVal seconds As Long = 90, _
                       Optional ByVal path As String = "")
    Dim p As String
    p = IIf(Len(path) > 0, path, LOCK_PATH)

    Debug.Print "target: " & p
    Debug.Print "time  : " & Format$(Now, "yyyy-mm-dd hh:mm:ss")

    Dim f As Integer

    ' Create the file first (harmless if it already exists).
    On Error Resume Next
    f = FreeFile
    Open p For Output As #f
    If Err.Number = 0 Then
        Print #f, "HELD_BY " & Environ$("USERNAME") & " @ " & Environ$("COMPUTERNAME")
        Close #f
    End If
    Err.Clear
    On Error GoTo EH

    ' Take exclusive access: write access + deny read and write to others.
    f = FreeFile
    On Error GoTo LockFailed
    Open p For Binary Access Write Lock Read Write As #f
    On Error GoTo EH

    Debug.Print "OK: holding exclusive access (handle " & f & ")"
    Debug.Print ""
    Debug.Print ">>> On the other machine run:"
    Debug.Print "    .\4-test-lock.ps1 -Path """ & p & """ -Mode Try"
    Debug.Print ""
    Debug.Print "holding for " & seconds & " seconds..."

    Dim i As Long
    For i = seconds To 1 Step -1
        If i Mod 10 = 0 Then Debug.Print "  remaining " & i & "s"
        SleepMs 1000
        DoEvents
    Next i

    Close #f
    Debug.Print "released."
    If SHOW_MSGBOX Then
        MsgBox "Lock released." & vbCrLf & vbCrLf & _
               "If the colleague side reported SHARING_VIOLATION(32)," & vbCrLf & _
               "the VBA lock approach holds.", vbInformation
    End If
    Exit Sub

LockFailed:
    Debug.Print "FAIL: cannot take exclusive access: " & Err.Number & " - " & Err.Description
    Debug.Print "  Err 70 (Permission denied) also means mutual exclusion is working."
    On Error GoTo 0
    If SHOW_MSGBOX Then
        MsgBox "Exclusive access refused: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
               "Err 70 means the other side holds it, which is the expected result.", vbInformation
    End If
    Exit Sub

EH:
    On Error Resume Next
    Close #f
    If SHOW_MSGBOX Then
        MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation
    End If
End Sub

' ------------------------------------------------------------
' Try to take exclusive access (the side B counterpart, in VBA).
' ------------------------------------------------------------
Public Sub TryLockVBA(Optional ByVal path As String = "")
    Dim p As String
    p = IIf(Len(path) > 0, path, LOCK_PATH)

    If Dir(p) = "" Then
        MsgBox "File not found: " & p & vbCrLf & _
               "Let side A run HoldLockVBA first.", vbExclamation
        Exit Sub
    End If

    Dim f As Integer
    f = FreeFile
    On Error GoTo Locked
    Open p For Binary Access Write Lock Read Write As #f
    On Error GoTo 0

    ' Unexpected success.
    Close #f
    Debug.Print "SEVERE: I also obtained exclusive access."
    If SHOW_MSGBOX Then
        MsgBox "I also obtained exclusive access!" & vbCrLf & vbCrLf & _
               "This share does not enforce file sharing violations," & vbCrLf & _
               "so the VBA Open Lock approach is unreliable here." & vbCrLf & _
               "Use a mkdir-based atomic lock instead.", vbCritical
    End If
    Exit Sub

Locked:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    Debug.Print "OK: refused with " & en & " - " & ed
    If SHOW_MSGBOX Then
        If en = 70 Then
            MsgBox "Refused with Err 70 (Permission denied)." & vbCrLf & vbCrLf & _
                   "Mutual exclusion works, so the VBA lock approach holds.", vbInformation
        Else
            MsgBox "Refused, but with a non-typical code: " & en & " - " & ed, vbExclamation
        End If
    End If
End Sub

' ------------------------------------------------------------
' Single-machine self test. Verifies VBA semantics only.
' Real SMB exclusion must be tested with two machines.
' ------------------------------------------------------------
Public Sub SelfTestMutex(Optional ByVal path As String = "")
    Dim log As String
    Dim f1 As Integer, f2 As Integer
    Dim p As String

    If Len(path) = 0 Then
        p = Environ$("TEMP") & "\exs_mutex_test.lock"
    Else
        p = path
    End If

    log = "mutex self test " & Format$(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf & String(60, "-") & vbCrLf
    log = log & "path: " & p & vbCrLf & String(60, "-") & vbCrLf

    ' prepare
    SafeKill p
    On Error Resume Next
    MkDir Left$(Environ$("TEMP"), Len(Environ$("TEMP")))
    Err.Clear
    On Error GoTo 0

    f1 = FreeFile
    Open p For Output As #f1
    Print #f1, "x"
    Close #f1

    ' 1) first exclusive open
    f1 = FreeFile
    On Error Resume Next
    Err.Clear
    Open p For Binary Access Write Lock Read Write As #f1
    If Err.Number = 0 Then
        log = log & "1) first exclusive open : OK (handle " & f1 & ")" & vbCrLf
    Else
        log = log & "1) first exclusive open : FAILED " & Err.Number & " " & Err.Description & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0

    ' 2) second exclusive open from the same process (must fail)
    f2 = FreeFile
    On Error Resume Next
    Err.Clear
    Open p For Binary Access Write Lock Read Write As #f2
    If Err.Number = 0 Then
        log = log & "2) second exclusive open: UNEXPECTED SUCCESS" & vbCrLf
        Close #f2
    Else
        log = log & "2) second exclusive open: refused " & Err.Number & " (" & Err.Description & ")" & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0

    ' 3) shared read while write is locked
    f2 = FreeFile
    On Error Resume Next
    Err.Clear
    Open p For Binary Lock Write As #f2
    If Err.Number = 0 Then
        log = log & "3) shared read          : allowed (as expected)" & vbCrLf
        Close #f2
    Else
        log = log & "3) shared read          : refused " & Err.Number & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0

    ' 4) after release, exclusive open works again
    Close #f1
    f1 = FreeFile
    On Error Resume Next
    Err.Clear
    Open p For Binary Access Write Lock Read Write As #f1
    If Err.Number = 0 Then
        log = log & "4) open after release   : OK" & vbCrLf
        Close #f1
    Else
        log = log & "4) open after release   : FAILED " & Err.Number & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0

    ' 5) Name statement cannot overwrite an existing file (important for design)
    Dim p2 As String: p2 = p & ".new"
    On Error Resume Next
    Err.Clear
    f1 = FreeFile
    Open p2 For Output As #f1
    Print #f1, "new content"
    Close #f1
    Name p2 As p
    If Err.Number = 0 Then
        log = log & "5) Name over existing   : OK" & vbCrLf
    Else
        log = log & "5) Name over existing   : FAILED " & Err.Number & " " & Err.Description & vbCrLf
        log = log & "   (expected error 58: VBA Name cannot replace an existing file)" & vbCrLf
        log = log & "   -> use MoveFileExW(MOVEFILE_REPLACE_EXISTING) or FSO MoveFile(OverWriteFiles)" & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
    SafeKill p2

    ' 6) lock/unlock cost
    Dim t As Single, i As Long
    t = Timer
    For i = 1 To 100
        f1 = FreeFile
        Open p For Binary Access Write Lock Read Write As #f1
        Close #f1
    Next i
    log = log & "6) 100x lock+unlock      : " & Format$((Timer - t) * 1000, "0") & " ms" & vbCrLf

    log = log & String(60, "-") & vbCrLf
    log = log & "This self test on a local disk verifies VBA semantics only." & vbCrLf
    log = log & "Real SMB exclusion must be tested from two machines." & vbCrLf

    Debug.Print log

    Dim rp As String
    rp = Environ$("USERPROFILE") & "\Desktop\exs_vbalock.txt"
    f1 = FreeFile
    On Error Resume Next
    Open rp For Output As #f1
    Print #f1, log
    Close #f1
    On Error GoTo 0

    SafeKill p
    If SHOW_MSGBOX Then MsgBox "Self test finished. Report: " & rp, vbInformation
End Sub

' ------------------------------------------------------------
Private Sub SleepMs(ByVal ms As Long)
    Dim t As Single
    t = Timer
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
End Sub

' Safe delete: no error when the file does not exist (avoids error 53).
Private Sub SafeKill(ByVal p As String)
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    Err.Clear
    On Error GoTo 0
End Sub
