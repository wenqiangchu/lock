# =============================================================================
#  9-t3-t5-driver.ps1   --   T3 (save atomicity) + T5 (concurrent reader)
#
#  Pure ASCII. Windows PowerShell 5.1. Requires Excel 16.0.
#
#  WHAT IT DOES
#    Works ONLY on a throwaway probe copy. Never touches a production file.
#
#      P0  baseline          copy probe, record hash, open/close sanity
#      P1  NORMAL_SAVE       dirty + Save while the Linux observer watches
#      P2  READER_HOLDS      Linux holds the file open while Excel saves
#      P3..P6  KILL_xS       kill EXCEL at 0.5 / 1.5 / 3.0 / 5.0 s into Save
#      P7  KILL_AFTER_OK     control: kill after Save already returned
#
#    After every phase the probe file is re-validated:
#      - does it still open in Excel?
#      - is the zip container intact?
#      - is the content OLD, NEW, or something else?
#      - are there leftover temp / ~$ files?
#
#  WHY
#    Save() is the ONLY moment we write to the server. It takes ~6.8 s for an
#    11 MB workbook, so every commit has a ~7 s window where the file on the
#    share is being rewritten. We must know what survives if that window is
#    interrupted, and what a Linux/SAS reader sees during it.
#
#  USAGE
#    .\9-t3-t5-driver.ps1 -ProbeDir '\\srv\share\_mddt_probe' `
#                         -SourceFile '\\srv\share\book.xlsx' [-SkipKillTests]
#
#  Then run the Linux observer in PuTTY, pointed at:
#      <ProbeDir>/book_probe.xlsx   and   <ProbeDir>/_phase.txt
#
#  SAFETY
#    - ProbeDir must contain "_probe" in its name (hard refusal otherwise)
#    - the probe file is a COPY; the source is only ever read
#    - every phase recreates the probe from the pristine local copy
# =============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ProbeDir,
    [Parameter(Mandatory = $true)][string]$SourceFile,
    [int]$KillDelaysMs = 0,          # 0 = use the built-in list
    [switch]$SkipKillTests,
    [switch]$SkipReaderHold,         # skip P2 (it needs a human on the Linux box)
    [int]$ReaderHoldWaitSec = 30,    # how long to wait for the manual hold
    [switch]$KeepProbe
)

$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------------- helpers
function Now-Ms {
    [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
}

function Say([string]$s) {
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss.fff'), $s
    Write-Host $line
    Add-Content -LiteralPath $script:LogFile -Value $line -Encoding ASCII
}

function Set-Phase([string]$name) {
    $script:PhaseName = $name
    Set-Content -LiteralPath $script:PhaseFile -Value $name -Encoding ASCII
    Say "PHASE >>> $name"
}

function Get-FileHashSafe([string]$p) {
    if (-not (Test-Path -LiteralPath $p)) { return 'MISSING' }
    try { (Get-FileHash -LiteralPath $p -Algorithm MD5).Hash.ToLower() } catch { 'ERR' }
}

function Get-SizeSafe([string]$p) {
    if (-not (Test-Path -LiteralPath $p)) { return -1 }
    try { (Get-Item -LiteralPath $p).Length } catch { -1 }
}

# Zip container integrity without needing Excel
function Test-ZipIntact([string]$p) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    try {
        $z = [System.IO.Compression.ZipFile]::OpenRead($p)
        $n = $z.Entries.Count
        $z.Dispose()
        return @{ Ok = $true; Entries = $n; Error = '' }
    }
    catch {
        return @{ Ok = $false; Entries = 0; Error = $_.Exception.Message }
    }
}

function Get-StrayFiles([string]$dir, [string]$stem) {
    $out = @()
    try {
        $out += Get-ChildItem -LiteralPath $dir -Force -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like '~$*' -or
                               $_.Name -like '*.tmp' -or
                               $_.Name -like '~DF*' -or
                               ($_.Name -like "$stem*" -and $_.Name -ne "$stem.xlsx") } |
                Select-Object -ExpandProperty Name
    } catch { }
    return $out
}

# Verify with Excel itself: can the file be opened, and what is in A1?
function Test-OpenWithExcel([string]$p) {
    $r = @{ Opens = $false; Content = ''; Error = ''; RepairPrompt = $false }
    $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $xl = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.EnableEvents = $false
        $xl.AskToUpdateLinks = $false
        $wb = $xl.Workbooks.Open($p, 0, $true)
        $r.Opens = $true
        try { $r.Content = [string]$wb.Worksheets(1).Range('A1').Value2 } catch { $r.Content = '<read fail>' }
        $wb.Close($false)
    }
    catch {
        $r.Error = $_.Exception.Message
        if ($r.Error -match 'repair|corrupt|damaged') { $r.RepairPrompt = $true }
    }
    finally {
        if ($xl) { try { $xl.Quit() } catch { }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { } }
        Start-Sleep -Milliseconds 300
        $after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        ($after | Where-Object { $before -notcontains $_ }) | ForEach-Object {
            try { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue } catch { }
        }
    }
    return $r
}

# Start a hidden Excel that owns exactly one workbook and return handle + pid
function Start-ExcelWith([string]$p) {
    $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $xl.ScreenUpdating = $false
    $xl.AskToUpdateLinks = $false
    $wb = $xl.Workbooks.Open($p, 0, $false)
    Start-Sleep -Milliseconds 200
    $after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $new = @($after | Where-Object { $before -notcontains $_ })
    return @{ Xl = $xl; Wb = $wb; Pid = ($new | Select-Object -First 1) }
}

# Kill helper launched as a SEPARATE process so the kill is genuinely external.
function Start-Killer([int]$TargetPid, [int]$DelayMs, [string]$GoFile, [string]$DoneFile) {
    $killer = Join-Path $env:TEMP ("mddt_killer_{0}.ps1" -f ([Guid]::NewGuid().ToString('N').Substring(0, 8)))
    $body = @'
param($TargetPid,$DelayMs,$GoFile,$DoneFile)
$ErrorActionPreference = 'SilentlyContinue'
"waiting" | Set-Content -LiteralPath $DoneFile -Encoding ASCII
$spin = 0
while (-not (Test-Path -LiteralPath $GoFile)) {
    Start-Sleep -Milliseconds 5
    $spin++
    if ($spin -gt 120000) { exit 9 }
}
$go = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
Start-Sleep -Milliseconds $DelayMs
$kill = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
Stop-Process -Id $TargetPid -Force
"go=$go kill=$kill delay=$DelayMs" | Set-Content -LiteralPath $DoneFile -Encoding ASCII
'@
    Set-Content -LiteralPath $killer -Value $body -Encoding ASCII
    $p = Start-Process -FilePath 'powershell.exe' `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $killer,
                        '-TargetPid', $TargetPid, '-DelayMs', $DelayMs,
                        '-GoFile', $GoFile, '-DoneFile', $DoneFile) `
        -PassThru -WindowStyle Hidden
    return @{ Proc = $p; Script = $killer }
}

function Invoke-KillDuringSave([string]$phase, [int]$delayMs, [string]$pristineCopy) {
    Say "---- $phase : kill at ${delayMs} ms into Save ----"
    Set-Phase $phase

    # fresh probe from the pristine local copy
    Copy-Item -LiteralPath $pristineCopy -Destination $script:ProbeFile -Force
    $goFile = Join-Path $env:TEMP 'mddt_go.txt'
    $doneFile = Join-Path $env:TEMP 'mddt_killer_done.txt'
    Remove-Item -LiteralPath $goFile, $doneFile -Force -ErrorAction SilentlyContinue

    $h = Start-ExcelWith $script:ProbeFile
    if (-not $h.Pid) { Say "  !! could not identify EXCEL pid, skipping"; try { $h.Xl.Quit() } catch { }; return }

    $k = Start-Killer -TargetPid $h.Pid -DelayMs $delayMs -GoFile $goFile -DoneFile $doneFile
    Start-Sleep -Milliseconds 400   # let the killer arm itself

    # make it dirty
    try { $h.Wb.Worksheets(1).Range('A1').Value2 = "$phase dirty" } catch { Say "  dirty failed: $($_.Exception.Message)" }

    Say "  Save() starting, killer armed at ${delayMs} ms ..."
    $tSave0 = Now-Ms
    New-Item -ItemType File -Path $goFile -Force | Out-Null

    $saveErr = ''
    try { $h.Wb.Save() } catch { $saveErr = $_.Exception.Message }

    $tSave1 = Now-Ms
    Set-Content -LiteralPath $script:PhaseFile -Value 'idle' -Encoding ASCII

    try { $k.Proc.WaitForExit(15000) | Out-Null } catch { }
    $detail = ''
    if (Test-Path -LiteralPath $doneFile) { $detail = (Get-Content -LiteralPath $doneFile -Raw).Trim() }

    $saveMs = $tSave1 - $tSave0
    Say ("  Save() returned after {0} ms; err=[{1}]" -f $saveMs, $saveErr)
    Say "  killer: $detail"

    # did the kill land while Save() was still running?
    $landed = $false
    if ($detail -match 'kill=(\d+)') {
        $killMs = [int64]$Matches[1]
        $landed = ($killMs -ge $tSave0) -and ($killMs -le $tSave1)
    }
    if (-not $landed) {
        Say "  !! KILL WAS TOO LATE - it fired after Save() had already returned."
        Say "  !! This phase proves nothing about mid-save corruption."
    } else {
        Say "  kill landed DURING Save() - this phase is meaningful."
    }
    Remove-Item -LiteralPath $k.Script -Force -ErrorAction SilentlyContinue

    # if the process is somehow still alive, kill it now
    try { Stop-Process -Id $h.Pid -Force -ErrorAction SilentlyContinue } catch { }
    Start-Sleep -Milliseconds 800
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($h.Wb) } catch { }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($h.Xl) } catch { }

    Report-Probe $phase $landed
}

function Report-Probe([string]$phase, [bool]$killLanded = $false) {
    Start-Sleep -Milliseconds 400
    $size = Get-SizeSafe $script:ProbeFile
    $hash = Get-FileHashSafe $script:ProbeFile
    $zip = if ($size -gt 0) { Test-ZipIntact $script:ProbeFile } else { @{ Ok = $false; Entries = 0; Error = 'file missing' } }
    $stray = Get-StrayFiles $script:ProbeDir $script:ProbeStem
    $op = Test-OpenWithExcel $script:ProbeFile

    # NOTE: we do NOT compare against a single "NEW" hash. Each phase writes
    # different content, so a third hash is EXPECTED and is not a fault.
    # What matters: does the file still exist, is the zip intact, does Excel open it.
    $state = 'UNKNOWN'
    if ($hash -eq 'MISSING') { $state = 'GONE' }
    elseif ($hash -eq 'ERR') { $state = 'UNREADABLE' }
    elseif ($hash -eq $script:HashClean) { $state = 'OLD_UNCHANGED' }
    else { $state = 'CHANGED' }

    $verdict = 'ok'
    if (-not $zip.Ok) { $verdict = 'ZIP BROKEN' }
    elseif (-not $op.Opens) { $verdict = 'EXCEL CANNOT OPEN' }
    elseif ($state -eq 'GONE') { $verdict = 'FILE GONE' }
    elseif ($state -eq 'UNREADABLE') { $verdict = 'UNREADABLE' }

    Say "  size=$size zipOk=$($zip.Ok) entries=$($zip.Entries) excelOpens=$($op.Opens) content=$state"
    Say "  md5=$hash"
    Say "  stray=[$($stray -join ', ')]"
    if ($op.Error) { Say "  excel error: $($op.Error)" }
    Say "  >>> VERDICT: $verdict"

    $row = [pscustomobject]@{
        phase = $phase; size = $size; md5 = $hash; content = $state
        zipOk = $zip.Ok; entries = $zip.Entries
        excelOpens = $op.Opens; excelError = $op.Error; a1 = $op.Content
        stray = ($stray -join '|'); verdict = $verdict; killInWindow = $killLanded
    }
    $script:Results += $row
}

# ----------------------------------------------------------------------------- main
$script:LogFile = Join-Path $env:TEMP ("mddt_t3t5_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$script:Results = @()
$script:PhaseName = 'idle'

Say "=============================================================="
Say " T3 (save atomicity) + T5 (concurrent reader)"
Say "=============================================================="
Say "probe dir   : $ProbeDir"
Say "source file : $SourceFile"
Say "log         : $script:LogFile"

# --- hard safety gate
if ($ProbeDir -notmatch '_probe') {
    throw "REFUSED: ProbeDir must contain '_probe' in its name. Got: $ProbeDir"
}
if (-not (Test-Path -LiteralPath $ProbeDir)) {
    throw "ProbeDir does not exist: $ProbeDir"
}
if (-not (Test-Path -LiteralPath $SourceFile)) {
    throw "SourceFile does not exist: $SourceFile"
}
if ((Resolve-Path -LiteralPath $ProbeDir).Path -eq (Split-Path -Parent $SourceFile)) {
    throw "REFUSED: ProbeDir is the same directory as SourceFile."
}

$srcLen = (Get-Item -LiteralPath $SourceFile).Length
Say ("source size : {0:N0} bytes ({1:N2} MB)" -f $srcLen, ($srcLen / 1MB))

# --- the REAL probe file lives on the share so the Linux observer can see it
$script:ProbeStem = [IO.Path]::GetFileNameWithoutExtension($SourceFile) + '_probe'
$script:ProbeFile = Join-Path $ProbeDir ($script:ProbeStem + '.xlsx')
$script:PhaseFile = Join-Path $ProbeDir '_phase.txt'

Say "probe file  : $script:ProbeFile"
Say "phase file  : $script:PhaseFile"
Say ""

# --- P0: pristine local copy + baseline hashes
Set-Phase 'P0_BASELINE'
$pristine = Join-Path $env:TEMP ($script:ProbeStem + '_pristine.xlsx')
Say "copying source -> local pristine ..."
Copy-Item -LiteralPath $SourceFile -Destination $pristine -Force
Say ("pristine local copy: {0:N0} bytes" -f (Get-Item -LiteralPath $pristine).Length)

Say "copying pristine -> probe on share ..."
Copy-Item -LiteralPath $pristine -Destination $script:ProbeFile -Force
$script:HashClean = Get-FileHashSafe $script:ProbeFile
Say "clean hash : $($script:HashClean)"

# produce a "dirty saved" reference hash so we can tell OLD from NEW
Say "producing the NEW-content reference hash ..."
$h = Start-ExcelWith $pristine
try {
    $h.Wb.Worksheets(1).Range('A1').Value2 = 'REFERENCE_DIRTY'
    $h.Wb.Save()
} catch { Say "  !! reference save failed: $($_.Exception.Message)" }
Start-Sleep -Milliseconds 500
try { $h.Wb.Close($false) } catch { }
try { $h.Xl.Quit() } catch { }
Start-Sleep -Milliseconds 600
try { Stop-Process -Id $h.Pid -Force -ErrorAction SilentlyContinue } catch { }
$script:HashDirty = Get-FileHashSafe $pristine
Say "dirty hash : $($script:HashDirty)"
Say "  (OLD = $($script:HashClean.Substring(0,8))... , NEW = $($script:HashDirty.Substring(0,8))...)"

if (-not $SkipReaderHold) {
    Say ""
    Say "################################################################"
    Say " NOW start the Linux observer in PuTTY:"
    Say "   ./9-linux-observer.sh '$($script:ProbeFile)' 600 200 '$($script:PhaseFile)' /tmp/t3t5.csv"
    Say " You have 20 seconds to do that. Continuing automatically."
    Say "################################################################"
    for ($i = 20; $i -gt 0; $i--) {
        Write-Host "`r  observer countdown: $i s   " -NoNewline
        Start-Sleep -Seconds 1
    }
    Write-Host "`r  proceeding                  "
}

# --- P1: normal save, reader watching
Say ""
Say "==== P1 NORMAL_SAVE ===="
Set-Phase 'P1_NORMAL_SAVE'
Copy-Item -LiteralPath $pristine -Destination $script:ProbeFile -Force
$h = Start-ExcelWith $script:ProbeFile
$t0 = Now-Ms
try { $h.Wb.Worksheets(1).Range('A1').Value2 = 'P1' ; $h.Wb.Save() }
catch { Say "  P1 save error: $($_.Exception.Message)" }
$t1 = Now-Ms
Say ("  P1 Save() took {0} ms" -f ($t1 - $t0))
try { $h.Wb.Close($false) } catch { }
try { $h.Xl.Quit() } catch { }
Start-Sleep -Milliseconds 600
try { Stop-Process -Id $h.Pid -Force -ErrorAction SilentlyContinue } catch { }
Set-Phase 'idle'
Report-Probe 'P1_NORMAL_SAVE'

# --- P2: a Linux reader holds the file while Excel saves
if (-not $SkipReaderHold) {
Say ""
Say "==== P2 READER_HOLDS (needs you to hold the file on Linux) ===="
Say "  In PuTTY, run this NOW:"
Say "     exec 9< '$($script:ProbeFile)'; sleep 60; exec 9<&-"
Say ("  Waiting {0} s so the hold is definitely active ..." -f $ReaderHoldWaitSec)
for ($i = $ReaderHoldWaitSec; $i -gt 0; $i--) {
    Write-Host "`r  hold countdown: $i s   " -NoNewline
    Start-Sleep -Seconds 1
}
Write-Host "`r  starting P2                  "

Set-Phase 'P2_READER_HOLDS'
Copy-Item -LiteralPath $pristine -Destination $script:ProbeFile -Force
$h = Start-ExcelWith $script:ProbeFile
$t0 = Now-Ms
$p2err = ''
try { $h.Wb.Worksheets(1).Range('A1').Value2 = 'P2'; $h.Wb.Save() }
catch { $p2err = $_.Exception.Message }
$t1 = Now-Ms
Say ("  P2 Save() took {0} ms  err=[{1}]" -f ($t1 - $t0), $p2err)
if ($p2err) { Say "  >>> EXCEL COULD NOT SAVE WHILE LINUX HELD THE FILE" }
else { Say "  >>> EXCEL SAVED FINE WHILE LINUX HELD THE FILE" }
try { $h.Wb.Close($false) } catch { }
try { $h.Xl.Quit() } catch { }
Start-Sleep -Milliseconds 600
try { Stop-Process -Id $h.Pid -Force -ErrorAction SilentlyContinue } catch { }
Set-Phase 'idle'
Report-Probe 'P2_READER_HOLDS'
}   # end P2

# --- P3..P6: kill mid-save
if (-not $SkipKillTests) {
    $delays = if ($KillDelaysMs -gt 0) { @($KillDelaysMs) } else { @(500, 1500, 3000, 5000) }
    $i = 3
    foreach ($d in $delays) {
        Say ""
        Say "==== KILL TEST ===="
        Invoke-KillDuringSave ("P{0}_KILL_{1}MS" -f $i, $d) $d $pristine
        $i++
    }

    # --- control: kill AFTER save returned
    Say ""
    Say "==== P7 KILL_AFTER_OK (control) ===="
    Set-Phase 'P7_KILL_AFTER_OK'
    Copy-Item -LiteralPath $pristine -Destination $script:ProbeFile -Force
    $h = Start-ExcelWith $script:ProbeFile
    try { $h.Wb.Worksheets(1).Range('A1').Value2 = 'P7'; $h.Wb.Save() } catch { }
    Say "  Save() completed, now killing EXCEL"
    try { Stop-Process -Id $h.Pid -Force -ErrorAction SilentlyContinue } catch { }
    Start-Sleep -Milliseconds 800
    Set-Phase 'idle'
    Report-Probe 'P7_KILL_AFTER_OK'
}

# --- summary
Say ""
Say "=============================================================="
Say " SUMMARY"
Say "=============================================================="
$script:Results | Format-Table phase, size, content, zipOk, excelOpens, killInWindow, verdict -AutoSize |
    Out-String -Width 200 | ForEach-Object { $_.TrimEnd() } | ForEach-Object { Say $_ }

$bad = @($script:Results | Where-Object { $_.verdict -ne 'ok' })
Say ""
if ($bad.Count -eq 0) {
    Say ">>> ALL PHASES OK: the server file was never left damaged."
} else {
    Say ">>> $($bad.Count) PHASE(S) LEFT A BAD FILE:"
    $bad | ForEach-Object { Say "      $($_.phase) : $($_.verdict)" }
    Say ""
    Say ">>> THIS MEANS: a crash during Save can destroy the file."
    Say ">>> DESIGN CHANGE REQUIRED: back up before Save."
}

$csv = Join-Path $env:TEMP ("mddt_t3t5_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$script:Results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding ASCII
Say ""
Say "csv : $csv"
Say "log : $script:LogFile"
Say "pristine kept at: $pristine"

if (-not $KeepProbe) {
    Set-Phase 'idle'
    Remove-Item -LiteralPath $script:ProbeFile -Force -ErrorAction SilentlyContinue
    Say "probe removed from the share."
}
Say "done."