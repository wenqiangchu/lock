<#
    3b-save-test.ps1

    Measure the REAL cost of Excel saving a workbook back to the SMB share.

    Why this test exists
    --------------------
    Test 3 established:
        create a NEW file on the share   ~106 - 245 s
        overwrite an EXISTING file       ~4 - 8 s
        rename within the share          ~0.4 s

    Question: when Excel saves, does it overwrite in place (~5 s) or does it
    write a temp file and rename it, thereby paying the new-file penalty
    (~110 s) on every single save?

    Safety
    ------
    Writes ONLY inside -ShareProbeDir. The production file is used strictly
    as a read-only copy source. Probe files are listed at the end and can be
    removed with -Cleanup.

    Pure ASCII. PowerShell 5.1 compatible.

    Paths
    -----
    Both UNC (\\server\share\...) and mapped drives (X:\...) are accepted.
    A mapped drive is reported together with the UNC it resolves to, because
    the two are NOT equivalent for measurement or for the final tool.

    Usage
    -----
        .\3b-save-test.ps1 `
            -SourceFile     "X:\path\mddt_sdtm_xxx.xlsx" `
            -ShareProbeDir  "X:\_exs_probe_YOURNAME" `
            -Repeat 3

        # then, to delete the probe files
        .\3b-save-test.ps1 -SourceFile ... -ShareProbeDir ... -Repeat 0 -Cleanup
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SourceFile,
    [Parameter(Mandatory = $true)][string]$ShareProbeDir,
    [int]$Repeat = 3,
    [switch]$Cleanup
)

$ErrorActionPreference = 'Stop'

function Show([string]$Label, [double]$Seconds) {
    Write-Host ("  {0,-46} {1,9:N2} s" -f $Label, $Seconds)
}

# ----------------------------------------------------------------- preflight

# Resolve a mapped drive (X:\...) to its UNC target. Returns $null when the
# path is UNC already, is a local disk, or the mapping cannot be resolved.
function Get-UncPath([string]$Path) {
    if ($Path.StartsWith('\\')) { return $Path }
    if ($Path -notmatch '^[A-Za-z]:') { return $null }
    try {
        $drv = Get-PSDrive -Name $Path.Substring(0, 1) -ErrorAction Stop
        if ($drv -and $drv.DisplayRoot) {
            return ($drv.DisplayRoot.TrimEnd('\') + $Path.Substring(2))
        }
    }
    catch { }
    return $null
}

# True when the path lives on a network share (UNC or a mapped network drive).
function Test-IsNetworkPath([string]$Path) {
    if ($Path.StartsWith('\\')) { return $true }
    if ((Get-UncPath $Path) -ne $null) { return $true }
    return $false
}

Write-Host ''
Write-Host '=== preflight ==='
Write-Host ("  Documents     : " + [Environment]::GetFolderPath('MyDocuments'))
Write-Host ("  LOCALAPPDATA  : " + $env:LOCALAPPDATA)
Write-Host ("  SourceFile    : " + $SourceFile)
Write-Host ("  ShareProbeDir : " + $ShareProbeDir)

$uncSource = Get-UncPath $SourceFile
$uncProbe  = Get-UncPath $ShareProbeDir

if ($null -ne $uncSource) { Write-Host ("  SourceFile  UNC = " + $uncSource) }
if ($null -ne $uncProbe)  { Write-Host ("  ProbeDir    UNC = " + $uncProbe)  }

if (-not (Test-IsNetworkPath $ShareProbeDir)) {
    throw ('ShareProbeDir must be on the network share (UNC or a mapped drive), ' +
           'not a local disk. Got: ' + $ShareProbeDir)
}

# --- OneDrive guard --------------------------------------------------------
# A local working copy or a probe copy must never land inside OneDrive:
# that would upload the data to the cloud, which the data policy forbids.
$odRoots = @()
foreach ($v in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
    if ($v -and $v.Trim().Length -gt 0) { $odRoots += $v.Trim().TrimEnd('\') }
}
foreach ($known in @('MyDocuments', 'Desktop')) {
    $p = ''
    try { $p = [Environment]::GetFolderPath($known) } catch { }
    if ($p -and ($p -match 'OneDrive')) { $odRoots += $p.TrimEnd('\') }
}

if ($odRoots.Count -gt 0) {
    Write-Host ('  OneDrive roots: ' + ($odRoots -join ' | '))
} else {
    Write-Host '  OneDrive roots: none detected'
}

# Running from inside a OneDrive-redirected folder is a trap: any RELATIVE path
# (including a bare probe folder name) then resolves into OneDrive.
$cwd = (Get-Location).Path
Write-Host ('  CurrentDir    : ' + $cwd)
foreach ($r in $odRoots) {
    if ($cwd.ToLower().StartsWith($r.ToLower())) {
        Write-Host ''
        Write-Host '  WARNING: the current directory is inside OneDrive.'
        Write-Host '           Any relative path will resolve into OneDrive.'
        Write-Host '           Use absolute X:\ paths only.'
        Write-Host ''
        break
    }
}

$checkTargets = @($ShareProbeDir)
if ($null -ne $uncProbe) { $checkTargets += $uncProbe }
if ($null -ne $uncSource) { $checkTargets += $uncSource }

foreach ($t in $checkTargets) {
    foreach ($r in $odRoots) {
        if ($t.ToLower().TrimEnd('\').StartsWith($r.ToLower())) {
            throw ('Refusing to write: "' + $t + '" is inside OneDrive root "' + $r + '". ' +
                   'The measurement would be meaningless and the data would be synced to the cloud.')
        }
    }
    if ($t.ToLower().Contains('onedrive')) {
        throw ('Refusing to write: "' + $t + '" looks like a OneDrive path.')
    }
}

if (-not (Test-Path -LiteralPath $SourceFile)) {
    throw ('SourceFile not found: ' + $SourceFile)
}
if (-not (Test-Path -LiteralPath $ShareProbeDir)) {
    New-Item -ItemType Directory -Force -Path $ShareProbeDir | Out-Null
    Write-Host '  (created ShareProbeDir)'
}

if ($null -eq $uncProbe) {
    Write-Host '  NOTE: mapped drive could not be resolved to UNC.'
    Write-Host '        Run "net use" or "Get-PSDrive X" to inspect the mapping.'
}

$stamp  = Get-Date -Format 'HHmmss'
$uname  = $env:USERNAME
$base   = [IO.Path]::GetFileNameWithoutExtension($SourceFile)
$dst    = Join-Path $ShareProbeDir ($base + '_savetest_' + $uname + '.xlsx')

$touched = @()

Write-Host ''
Write-Host ('  >>> ALL probe files will be created under: ' + $ShareProbeDir)
Write-Host ('  >>> main probe file                     : ' + $dst)
Write-Host ''

# ---------------------------------------------------------------- hygiene
# Orphaned EXCEL.EXE processes and leftover ~$ owner files hold share-mode
# locks. They turn an open into a wait and produce meaningless timings.
Write-Host '=== environment hygiene ==='
$xlProcs = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($xlProcs.Count -eq 0) {
    Write-Host '  EXCEL.EXE running : 0'
} else {
    Write-Host ('  EXCEL.EXE running : ' + $xlProcs.Count)
    foreach ($p in $xlProcs) {
        Write-Host ("    pid={0}  title='{1}'" -f $p.Id, $p.MainWindowTitle)
    }
    Write-Host '  WARNING: any Excel instance can hold a share lock on the probe file.'
}

$lockFiles = @(Get-ChildItem -LiteralPath $ShareProbeDir -Filter '~$*' -Force -ErrorAction SilentlyContinue)
if ($lockFiles.Count -eq 0) {
    Write-Host '  ~$ lock files     : 0'
} else {
    Write-Host ('  ~$ lock files     : ' + $lockFiles.Count + '  <-- stale lock, delete before trusting timings')
    foreach ($f in $lockFiles) { Write-Host ('    ' + $f.Name) }
}

# ------------------------------------ phase 0: small-operation latency
# This is the real diagnostic. Bulk throughput says little; per-request
# latency multiplies by every small I/O that Excel performs.
Write-Host ''
Write-Host '=== phase 0: small-operation latency (50 x create + delete) ==='

$lat = @()
for ($i = 1; $i -le 50; $i++) {
    $lp = Join-Path $ShareProbeDir ('lat_' + $uname + '_' + $stamp + '_' + $i + '.tmp')
    $sw = [Diagnostics.Stopwatch]::StartNew()
    [IO.File]::WriteAllBytes($lp, (New-Object byte[] 0))
    Remove-Item -LiteralPath $lp -Force -ErrorAction SilentlyContinue
    $sw.Stop()
    $lat += $sw.Elapsed.TotalMilliseconds
}

$sortedLat = @($lat | Sort-Object)
$avgLat    = ($lat | Measure-Object -Average).Average
$p90idx    = [int][math]::Min($sortedLat.Count - 1, [math]::Floor($sortedLat.Count * 0.9))
$maxLat    = $sortedLat[$sortedLat.Count - 1]

Write-Host ("  min {0,7:N0} ms   avg {1,7:N0} ms   p90 {2,7:N0} ms   max {3,7:N0} ms" -f $sortedLat[0], $avgLat, $sortedLat[$p90idx], $maxLat)

if ($maxLat -gt ($avgLat * 5) -and $maxLat -gt 200) {
    Write-Host '  >> latency SPIKES: intermittent blocking, not bandwidth.'
}
if ($avgLat -gt 50) {
    Write-Host '  >> average latency is HIGH: every small I/O costs real time.'
}
if ($avgLat -le 15 -and $maxLat -lt 200) {
    Write-Host '  >> latency looks healthy for a LAN share.'
}

# --------------------------------------------------- phase A: raw SMB writes
Write-Host ''
Write-Host '=== phase A: raw SMB write cost (reference numbers) ==='

$zero = Join-Path $ShareProbeDir ('zero_' + $uname + '_' + $stamp + '.bin')

$sw = [Diagnostics.Stopwatch]::StartNew()
[IO.File]::WriteAllBytes($zero, (New-Object byte[] 0))
$sw.Stop()
Show 'create NEW 0-byte file' $sw.Elapsed.TotalSeconds
$touched += $zero

$sw = [Diagnostics.Stopwatch]::StartNew()
[IO.File]::WriteAllBytes($zero, (New-Object byte[] 0))
$sw.Stop()
Show 'overwrite SAME 0-byte file' $sw.Elapsed.TotalSeconds

$isNew = -not (Test-Path -LiteralPath $dst)
$sw = [Diagnostics.Stopwatch]::StartNew()
Copy-Item -LiteralPath $SourceFile -Destination $dst -Force
$sw.Stop()
if ($isNew) {
    Show 'copy workbook -> NEW name (first time)' $sw.Elapsed.TotalSeconds
} else {
    Show 'copy workbook -> EXISTING name' $sw.Elapsed.TotalSeconds
}
$touched += $dst

# --------------------------------------- phase B/C: Excel open + save timing
Write-Host ''
Write-Host '=== phase B: Excel open + Save (in place) ==='

$xl = $null
$wb = $null

if ($Repeat -gt 0) {

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible        = $false
    $xl.DisplayAlerts  = $false
    $xl.AskToUpdateLinks = $false
    $xl.EnableEvents   = $false
    $xl.ScreenUpdating = $false

    try {
        $fi0 = Get-Item -LiteralPath $dst
    Write-Host ("  probe copy : {0:N2} MB" -f ($fi0.Length / 1MB))

    for ($i = 1; $i -le $Repeat; $i++) {

            $sw = [Diagnostics.Stopwatch]::StartNew()
            $wb = $xl.Workbooks.Open($dst, 0, $false)
            $sw.Stop()
            $tOpen = $sw.Elapsed.TotalSeconds

            if ($i -eq 1) {
                Write-Host ("  sheets     : " + $wb.Worksheets.Count)
                Write-Host ("  legacy shared workbook (MultiUserEditing): " + $wb.MultiUserEditing)
            }

            # Mark the workbook dirty without touching any cell.
            # Sheet protection does not block document properties or defined names.
            try {
                $wb.BuiltinDocumentProperties.Item('Comments').Value = 'probe ' + $stamp + ' ' + $i
            }
            catch {
                $wb.Names.Add('ZZ_SAVE_PROBE', '=' + $stamp) | Out-Null
            }

            # Was the workbook genuinely dirty? If the dirty flag never got set,
            # Save() is a no-op and the elapsed time is meaningless.
            # $wb.Saved is authoritative and is NOT affected by the SMB
            # directory cache, which makes LastWriteTime look stale.
            $wasDirty = -not $wb.Saved

            $sw = [Diagnostics.Stopwatch]::StartNew()
            $wb.Save()
            $sw.Stop()
            $tSave = $sw.Elapsed.TotalSeconds

            $stillDirty = $wb.Saved

            $sw = [Diagnostics.Stopwatch]::StartNew()
            $wb.Close($false)
            $sw.Stop()
            $tClose = $sw.Elapsed.TotalSeconds

            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
            $wb = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()

            $flag = 'written'
            if (-not $wasDirty) { $flag = '*** NOT DIRTY BEFORE SAVE - number INVALID ***' }
            elseif ($stillDirty) { $flag = '*** STILL DIRTY AFTER SAVE ***' }

            Write-Host ("  run {0}: open {1,8:N2} s    SAVE() {2,8:N2} s    close {3,6:N2} s    {4}" -f $i, $tOpen, $tSave, $tClose, $flag)
        }

        # ------------------------ phase D: pure open/close, no save ----------
        # Isolates one question: does a previous SAVE make later opens slower?
        # Compare these numbers with phase B 'open' numbers.
        Write-Host ''
        Write-Host '=== phase D: open/close only, no save (3x) ==='

        for ($i = 1; $i -le 3; $i++) {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $wb = $xl.Workbooks.Open($dst, 0, $false)
            $sw.Stop()
            $tOpen = $sw.Elapsed.TotalSeconds

            $sw = [Diagnostics.Stopwatch]::StartNew()
            $wb.Close($false)
            $sw.Stop()
            $tClose = $sw.Elapsed.TotalSeconds

            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
            $wb = $null

            Write-Host ("  open-only {0}: open {1,8:N2} s    close {2,8:N2} s" -f $i, $tOpen, $tClose)
        }

        # ------------------------------------------------ phase C: SaveAs path
        Write-Host ''
        Write-Host '=== phase C: Excel SaveAs to a NEW file name ==='

        for ($i = 1; $i -le 2; $i++) {
            $newName = Join-Path $ShareProbeDir ($base + '_saveas_' + $uname + '_' + $stamp + '_' + $i + '.xlsx')

            $wb = $xl.Workbooks.Open($dst, 0, $false)
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $wb.SaveAs($newName, 51)      # 51 = xlOpenXMLWorkbook
            $sw.Stop()
            Write-Host ("  saveas {0}: {1,8:N2} s  -> {2}" -f $i, $sw.Elapsed.TotalSeconds, $newName)

            $wb.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
            $wb = $null
            $touched += $newName
        }
    }
    finally {
        if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
        if ($null -ne $xl) {
            try { $xl.Quit() } catch { }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ------------------------------------------------------------------- verdict
Write-Host ''
Write-Host '=== how to read the numbers ==='
Write-Host '  phase 0 avg under 15 ms        -> the share is healthy; look elsewhere.'
Write-Host '  phase 0 avg over 50 ms         -> every small I/O is expensive.'
Write-Host '  phase 0 max spikes             -> intermittent blocking / lock contention.'
Write-Host '  phase B SAVE() stable          -> the write path is predictable (good news).'
Write-Host '  phase B open wildly variable   -> lock contention on open. Check for'
Write-Host '                                    leftover EXCEL.EXE, ~$ files, or another user.'
Write-Host '  phase D open fast, B open slow -> each SAVE makes the next open slower.'

# ------------------------------------------------------------------- cleanup
Write-Host ''
if ($Cleanup) {
    Write-Host '=== cleanup ==='
    foreach ($p in $touched) {
        if (Test-Path -LiteralPath $p) {
            Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
            Write-Host ('  removed ' + $p)
        }
    }
} else {
    Write-Host '=== probe files kept (re-run with -Cleanup to delete) ==='
    foreach ($p in $touched) { Write-Host ('  ' + $p) }
}

Write-Host ''
Write-Host 'Done.'
