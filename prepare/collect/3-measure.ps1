# ============================================================
#  3-measure.ps1  -  network and Excel timing measurements
#
#  SAFETY: by default this script writes NOTHING to the network share.
#          Share-side writes only happen inside a probe folder, and only when
#          you pass -WorkDir explicitly.
#
#  Usage:
#     # network-only, zero writes to the share
#     .\3-measure.ps1 -File "\\SERVER\SHARE\book.xlsx"
#
#     # full run incl. Excel open/save against a scratch folder on the same share
#     .\3-measure.ps1 -File "\\SERVER\SHARE\book.xlsx" -WorkDir "\\SERVER\SHARE\_exs_probe_you"
#
#     # skip everything that needs Excel
#     .\3-measure.ps1 -File "..." -SkipExcel
#
#  NOTE: keep this file pure ASCII (PowerShell 5.1 reads BOM-less .ps1 as ANSI).
# ============================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$File,
    [string]$WorkDir = '',
    [string]$OutFile = "$([Environment]::GetFolderPath('Desktop'))\exs_measure.txt",
    [switch]$SkipExcel,
    [switch]$KeepProbe,
    [int]$Repeat = 3
)

$ErrorActionPreference = 'Continue'

function ResolveFull([string]$p) {
    try { return (Resolve-Path -LiteralPath $p -ErrorAction Stop).ProviderPath }
    catch {
        $parent = Split-Path $p -Parent
        if (-not $parent) { $parent = '.' }
        try { $absParent = (Resolve-Path -LiteralPath $parent -ErrorAction Stop).ProviderPath }
        catch { $absParent = [IO.Path]::GetFullPath((Join-Path (Get-Location).Path $parent)) }
        return [IO.Path]::Combine($absParent, (Split-Path $p -Leaf))
    }
}

$report = New-Object System.Collections.Generic.List[string]
function W($s = '') { $report.Add([string]$s) }
function Section($t) { W ''; W ('=' * 78); W "  $t"; W ('=' * 78) }
function Stat($label, [double[]]$sec, $unit = 's') {
    if (-not $sec -or $sec.Count -eq 0) { W ("  {0,-46} <no data>" -f $label); return }
    $min = ($sec | Measure-Object -Minimum).Minimum
    $max = ($sec | Measure-Object -Maximum).Maximum
    $avg = ($sec | Measure-Object -Average).Average
    W ("  {0,-46} min={1,8:N3}{4}  avg={2,8:N3}{4}  max={3,8:N3}{4}" -f $label, $min, $avg, $max, $unit)
}

W 'ExcelSync network and timing measurements'
W ("Generated: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
W "Test file: $File"
W ("Computer: {0}   User: {1}\{2}   Session: {3}" -f $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME, $env:SESSIONNAME)

if (-not (Test-Path $File)) {
    Write-Host "File not accessible: $File" -ForegroundColor Red
    W 'ERROR: file not accessible, aborting.'
    $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8
    return
}

$fi = Get-Item $File
$sizeMB = $fi.Length / 1MB
$dir = Split-Path $File -Parent
$baseName = [IO.Path]::GetFileNameWithoutExtension($File)
$ext = [IO.Path]::GetExtension($File)

# local scratch (never on the share)
$localDir = Join-Path $env:LOCALAPPDATA 'ExcelSync\_measure'
New-Item -ItemType Directory -Path $localDir -Force | Out-Null
$localFile = Join-Path $localDir ($baseName + $ext)

# share-side probe folder, only if requested
$probeDir = ''
if ($WorkDir -ne '') {
    $WorkDir = ResolveFull $WorkDir
    $probeDir = Join-Path $WorkDir ('_exs_probe_' + ($env:USERNAME -replace '[^\w]', ''))
    try {
        if (-not (Test-Path $probeDir)) { New-Item -ItemType Directory -Path $probeDir -Force -ErrorAction Stop | Out-Null }
        W ''
        W "Probe folder: $probeDir"
    } catch {
        W ''
        W "WARNING: cannot create probe folder: $($_.Exception.Message)"
        W '         Share-side write tests will be skipped.'
        $probeDir = ''
    }
}

# ------------------------------------------------------------
Section 'A. Preconditions'
# ------------------------------------------------------------
W ("File size     : {0:N2} MB ({1:N0} bytes)" -f $sizeMB, $fi.Length)
W ("File mtime    : {0}" -f $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))
W ("Share path    : $dir")
W ("Local scratch : $localDir")
W ("Probe folder  : {0}" -f $(if ($probeDir) { $probeDir } else { '<not requested, share writes disabled>' }))
$dirKind = if ($dir -like '\\*') { 'UNC path (SMB share)' }
           elseif ($dir -match '^[A-Za-z]:') { 'local drive ' + $dir.Substring(0, 2) }
           else { 'relative path (test only)' }
W ("Path kind     : $dirKind")

try {
    $conn = Get-SmbConnection -ErrorAction SilentlyContinue | Where-Object { $File -like "\\$($_.ServerName)\*" } | Select-Object -First 1
    if ($conn) { W ("SMB           : server=$($conn.ServerName) dialect=$($conn.Dialect) share=$($conn.ShareName)") }
    else { W 'SMB           : <no connection detail available>' }
} catch { W 'SMB           : <query failed>' }

# ------------------------------------------------------------
Section 'B. Pure network: file copy timing'
# ------------------------------------------------------------
$copyTimes = @()
$downloadedFile = $null
for ($i = 1; $i -le $Repeat; $i++) {
    $dst = Join-Path $localDir ("dl$i" + $ext)
    try {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        Copy-Item $File $dst -Force -ErrorAction Stop
        $sw.Stop()
        $copyTimes += $sw.Elapsed.TotalSeconds
        $rate = if ($sizeMB -ge 1) { ("  -> {0:N2} MB/s" -f ($sizeMB / [math]::Max($sw.Elapsed.TotalSeconds, 0.001))) } else { '' }
        W ("  run {0}: {1:N3} s{2}" -f $i, $sw.Elapsed.TotalSeconds, $rate)
        if ($i -eq 1) { $downloadedFile = $dst } else { Remove-Item $dst -Force -ErrorAction SilentlyContinue }
    } catch { W ("  run {0}: FAILED {1}" -f $i, $_.Exception.Message) }
}
W ''
Stat 'download (share -> local)' $copyTimes 's'
if ($copyTimes.Count -gt 1 -and $sizeMB -ge 1) {
    $first = $copyTimes[0]
    $rest = @($copyTimes | Select-Object -Skip 1)
    $restAvg = ($rest | Measure-Object -Average).Average
    W ''
    W ("  first {0:N3} s vs later avg {1:N3} s -> delta {2:N3} s (SMB connect, AV first scan, cold server cache)" -f $first, $restAvg, ($first - $restAvg))
}

if ($probeDir) {
    $upTimes = @()
    $upFile = Join-Path $probeDir ('up_' + $baseName + $ext)
    for ($i = 1; $i -le [math]::Min(2, $Repeat); $i++) {
        try {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            Copy-Item $File $upFile -Force -ErrorAction Stop
            $sw.Stop()
            $upTimes += $sw.Elapsed.TotalSeconds
            $rate = if ($sizeMB -ge 1) { (" -> {0:N2} MB/s" -f ($sizeMB / [math]::Max($sw.Elapsed.TotalSeconds, 0.001))) } else { '' }
            W ("  upload run {0}: {1:N3} s{2}" -f $i, $sw.Elapsed.TotalSeconds, $rate)
        } catch { W ("  upload run {0}: FAILED {1}" -f $i, $_.Exception.Message) }
    }
    W ''
    Stat 'upload (local -> share)' $upTimes 's'
} else {
    W ''
    W '  upload test skipped (no -WorkDir given)'
}

# ------------------------------------------------------------
Section 'C. Metadata operations (per-call cost of locks / state files)'
# ------------------------------------------------------------
$metaTimes = @()
for ($i = 1; $i -le 20; $i++) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $null = Test-Path $File
    $null = (Get-Item $File -ErrorAction SilentlyContinue).Length
    $null = (Get-Item $File -ErrorAction SilentlyContinue).LastWriteTime
    $sw.Stop()
    $metaTimes += $sw.Elapsed.TotalMilliseconds
}
Stat 'stat x20 (exists + length + mtime)' $metaTimes 'ms'

if ($probeDir) {
    $smallTimes = @()
    $smallPath = Join-Path $probeDir 'small.tst'
    for ($i = 1; $i -le 10; $i++) {
        try {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            Set-Content -Path $smallPath -Value ('x' * 512) -Encoding ASCII -NoNewline -ErrorAction Stop
            $null = Get-Content $smallPath -Raw -ErrorAction Stop
            Remove-Item $smallPath -Force -ErrorAction Stop
            $sw.Stop()
            $smallTimes += $sw.Elapsed.TotalMilliseconds
        } catch { break }
    }
    Stat 'small file create+write+read+delete x10' $smallTimes 'ms'

    $dirTimes = @()
    for ($i = 1; $i -le 5; $i++) {
        try {
            $d = Join-Path $probeDir ("mk$i")
            $sw = [Diagnostics.Stopwatch]::StartNew()
            New-Item -ItemType Directory -Path $d -ErrorAction Stop | Out-Null
            Remove-Item $d -Force -ErrorAction Stop
            $sw.Stop()
            $dirTimes += $sw.Elapsed.TotalMilliseconds
        } catch { break }
    }
    Stat 'mkdir+rmdir x5' $dirTimes 'ms'
} else {
    W '  share-side metadata tests skipped (no -WorkDir given)'
}

$listTimes = @()
for ($i = 1; $i -le 5; $i++) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $null = Get-ChildItem $dir -Force -ErrorAction SilentlyContinue
    $sw.Stop()
    $listTimes += $sw.Elapsed.TotalMilliseconds
}
Stat 'directory enumeration' $listTimes 'ms'

# ------------------------------------------------------------
Section 'D. Excel open / save timings (this is the commit window)'
# ------------------------------------------------------------
if ($SkipExcel) {
    W '  skipped (-SkipExcel)'
} else {
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.EnableEvents = $false
        W ("  Excel COM started, version={0}" -f $excel.Version)
    } catch {
        W "  ERROR: cannot start Excel COM: $($_.Exception.Message)"
        W '  Excel may be missing, or automation is blocked by policy.'
    }

    if ($excel) {
        # --- open local copy ---
        $openLocal = @()
        for ($i = 1; $i -le [math]::Min(2, $Repeat); $i++) {
            $wb = $null
            try {
                $sw = [Diagnostics.Stopwatch]::StartNew()
                $wb = $excel.Workbooks.Open($downloadedFile, 0, $true)   # ReadOnly
                $sw.Stop()
                $openLocal += $sw.Elapsed.TotalSeconds
                W ("  open local (read-only) run {0}: {1:N2} s   sheets={2}" -f $i, $sw.Elapsed.TotalSeconds, $wb.Worksheets.Count)
            } catch { W ("  open local failed: {0}" -f $_.Exception.Message) }
            finally {
                if ($wb) { try { $wb.Close($false) } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
            }
        }
        W ''
        Stat 'Excel OPEN local copy (read-only) = Excel parsing cost' $openLocal 's'

        # --- open share copy (read-only, inside probe folder, never the production file) ---
        $openShare = @()
        if ($probeDir) {
            $shareCopy = Join-Path $probeDir ('share_' + $baseName + $ext)
            try {
                Copy-Item $File $shareCopy -Force -ErrorAction Stop
                for ($i = 1; $i -le [math]::Min(2, $Repeat); $i++) {
                    $wb = $null
                    try {
                        $sw = [Diagnostics.Stopwatch]::StartNew()
                        $wb = $excel.Workbooks.Open($shareCopy, 0, $true)
                        $sw.Stop()
                        $openShare += $sw.Elapsed.TotalSeconds
                        W ("  open share copy (read-only) run {0}: {1:N2} s" -f $i, $sw.Elapsed.TotalSeconds)
                    } catch { W ("  open share copy failed: {0}" -f $_.Exception.Message) }
                    finally {
                        if ($wb) { try { $wb.Close($false) } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
                    }
                }
            } catch { W "  preparing share copy failed: $($_.Exception.Message)" }
        } else {
            W '  share-copy open skipped (no -WorkDir given)'
        }
        W ''
        Stat 'Excel OPEN share copy (read-only) = network + parsing' $openShare 's'
        if ($openLocal.Count -gt 0 -and $openShare.Count -gt 0) {
            $diff = (($openShare | Measure-Object -Average).Average) - (($openLocal | Measure-Object -Average).Average)
            W ''
            W ("  delta {0:N2} s = extra cost of opening from the share" -f $diff)
        }

        # --- save to share (the key metric) ---
        $saveShare = @()
        if ($probeDir) {
            $tmpSave = Join-Path $probeDir ('save_' + $baseName + $ext)
            for ($i = 1; $i -le [math]::Min(3, $Repeat); $i++) {
                $wb = $null
                try {
                    Copy-Item $File $tmpSave -Force -ErrorAction Stop
                    $wb = $excel.Workbooks.Open($tmpSave)
                    $wb.Worksheets.Item(1).Cells.Item(1, 1).Value2 = 'MEASURE_' + $i
                    $sw = [Diagnostics.Stopwatch]::StartNew()
                    $wb.Save()
                    $sw.Stop()
                    $saveShare += $sw.Elapsed.TotalSeconds
                    W ("  save to share run {0}: {1:N2} s" -f $i, $sw.Elapsed.TotalSeconds)
                } catch { W ("  save-to-share failed: {0}" -f $_.Exception.Message) }
                finally {
                    if ($wb) { try { $wb.Close($false) } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
                }
            }
        } else {
            W '  save-to-share test skipped (no -WorkDir given)'
        }
        W ''
        Stat 'Excel SAVE to share = main part of the commit window' $saveShare 's'

        # --- save to local (reference point) ---
        $saveLocal = @()
        for ($i = 1; $i -le 2; $i++) {
            $wb = $null
            try {
                $tl = Join-Path $localDir ("savelocal$i" + $ext)
                Copy-Item $File $tl -Force -ErrorAction Stop
                $wb = $excel.Workbooks.Open($tl)
                $wb.Worksheets.Item(1).Cells.Item(1, 1).Value2 = 'MEASURE_L' + $i
                $sw = [Diagnostics.Stopwatch]::StartNew()
                $wb.Save()
                $sw.Stop()
                $saveLocal += $sw.Elapsed.TotalSeconds
                W ("  save local run {0}: {1:N2} s" -f $i, $sw.Elapsed.TotalSeconds)
            } catch { W ("  save-local failed: {0}" -f $_.Exception.Message) }
            finally {
                if ($wb) { try { $wb.Close($false) } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
            }
        }
        W ''
        Stat 'Excel SAVE local (reference)' $saveLocal 's'
        if ($saveLocal.Count -gt 0 -and $saveShare.Count -gt 0) {
            $nw = (($saveShare | Measure-Object -Average).Average) - (($saveLocal | Measure-Object -Average).Average)
            W ''
            W ("  network overhead on Save ~= {0:N2} s (this is the tax paid on every commit)" -f $nw)
        }

        # --- ~$ lock file behaviour ---
        if ($probeDir) {
            W ''
            W '  ~$ lock file behaviour:'
            $lockName = '~$' + [IO.Path]::GetFileName($tmpSave)
            $lockFull = Join-Path $probeDir $lockName
            $wb = $null
            try {
                $wb = $excel.Workbooks.Open($tmpSave)              # read-write
                Start-Sleep -Milliseconds 400
                W ("    opened read-write : lock file present = {0}  (expected True)" -f (Test-Path $lockFull))
                $wb.Close($false); $wb = $null
                Start-Sleep -Milliseconds 400
                $after = Test-Path $lockFull
                W ("    after close       : lock file present = {0}  (expected False; True means orphan-lock risk)" -f $after)
                if ($after) { W '    WARNING: lock file survived Close - orphan lock scenario.' }

                $wb = $excel.Workbooks.Open($tmpSave, 0, $true)    # read-only
                Start-Sleep -Milliseconds 400
                $ro = Test-Path $lockFull
                W ("    opened read-only  : lock file present = {0}  (False means a View button truly avoids locking)" -f $ro)
            } catch { W ("    ~`$ test failed: {0}" -f $_.Exception.Message) }
            finally {
                if ($wb) { try { $wb.Close($false) } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
            }

            # --- does a read-only open block someone else's write? ---
            W ''
            W '  Does a read-only open block another writer? (decides if a View button is safe)'
            $wb1 = $null; $excel2 = $null; $wb2 = $null
            try {
                $wb1 = $excel.Workbooks.Open($tmpSave, 0, $true)
                $excel2 = New-Object -ComObject Excel.Application
                $excel2.Visible = $false; $excel2.DisplayAlerts = $false
                try {
                    $wb2 = $excel2.Workbooks.Open($tmpSave, 0, $false)
                    W '    OK: a second instance can open read-write while the first is read-only'
                } catch {
                    W ("    BLOCKED: second instance cannot open read-write: {0}" -f $_.Exception.Message)
                    W '    -> a read-only open also blocks commits; readers must download a local copy instead.'
                }
            } catch { W ("    test failed: {0}" -f $_.Exception.Message) }
            finally {
                if ($wb2) { try { $wb2.Saved = $true; $wb2.Close($false) } catch {} }
                if ($excel2) { try { $excel2.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel2) } catch {} }
                if ($wb1) { try { $wb1.Close($false) } catch {} }
            }
        }

        try { $excel.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 500
        $leftover = Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' }
        if ($leftover) {
            W ''
            W ("  NOTE: {0} hidden EXCEL process(es) left over (normal for COM). To clean:" -f $leftover.Count)
            W '        Get-Process EXCEL | Where-Object {$_.MainWindowTitle -eq ""} | Stop-Process'
        }
    }
}

# ------------------------------------------------------------
Section 'F. Verdict'
# ------------------------------------------------------------
if ($copyTimes.Count -gt 0) {
    $netDown = ($copyTimes | Measure-Object -Average).Average
    if ($sizeMB -lt 1) {
        W ("  NOTE: test file is only {0:N2} MB, timings are not representative." -f $sizeMB)
        W '        Rerun with the real production file.'
    }
    W ("  download (extrapolated to 10 MB): {0:N2} s" -f ($netDown / [math]::Max($sizeMB, 0.01) * 10))
}
W ''
W '  Commit window ~= open share file + Copy region + Save to share'
W '  Acceptance: Save to share <= 10 s -> viable (today the wait is hours).'
W '              Save to share > 30 s  -> slim the file first.'

# ------------------------------------------------------------
Section 'G. Cleanup'
# ------------------------------------------------------------
if ($probeDir) {
    W "Probe folder: $probeDir"
    if ($KeepProbe) {
        W 'Kept (-KeepProbe). Delete it manually when done.'
    } else {
        $ans = Read-Host 'Delete the probe folder on the share now? [y/N]'
        if ($ans -match '^[Yy]') {
            try { Remove-Item $probeDir -Recurse -Force -ErrorAction Stop; W 'Probe folder deleted.' }
            catch { W "Delete failed: $($_.Exception.Message)" }
        } else { W 'Kept. Delete it manually when done.' }
    }
} else {
    W 'Nothing was written to the share.'
}
try { Remove-Item $localDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}

$report -join "`r`n" | Set-Content -Path $OutFile -Encoding UTF8
Write-Host ''
Write-Host "Report saved: $OutFile" -ForegroundColor Green
Write-Host 'Please send this file back.' -ForegroundColor Yellow