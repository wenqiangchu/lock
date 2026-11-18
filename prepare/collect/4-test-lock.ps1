# ============================================================
#  4-test-lock.ps1  -  shared-share mutual exclusion test (needs two machines)
#
#  Verifies whether the SMB server really enforces file sharing violations and
#  whether directory creation is atomic. This decides if owners.tsv locking is
#  trustworthy.
#
#  SAFETY: writes exactly one lock file and one lock directory at -Path.
#          The script refuses to write into a directory that looks like
#          production (contains .xlsx but no "_exs_probe" in the path)
#          unless you pass -Force.
#
#  Usage (A = you, B = colleague, SAME path on both):
#     A: .\4-test-lock.ps1 -Path "\\SERVER\SHARE\_exs_probe_me\lock.tst" -Mode Hold -Seconds 90
#     B: .\4-test-lock.ps1 -Path "\\SERVER\SHARE\_exs_probe_me\lock.tst" -Mode Try
#
#  Directory atomicity variant:
#     A: .\4-test-lock.ps1 -Path "\\SERVER\SHARE\_exs_probe_me" -Mode HoldDir -Seconds 90
#     B: .\4-test-lock.ps1 -Path "\\SERVER\SHARE\_exs_probe_me" -Mode TryDir
#
#  Cleanup:
#     .\4-test-lock.ps1 -Path "\\SERVER\SHARE\_exs_probe_me\lock.tst" -Mode Cleanup
#
#  NOTE: keep this file pure ASCII (PowerShell 5.1 reads BOM-less .ps1 as ANSI).
# ============================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [ValidateSet('Hold', 'Try', 'HoldDir', 'TryDir', 'Cleanup')][string]$Mode = 'Try',
    [int]$Seconds = 90,
    [switch]$Force,
    [string]$OutFile = "$([Environment]::GetFolderPath('Desktop'))\exs_lock_$(($env:USERNAME) -replace '[^\w]','').txt"
)

$ErrorActionPreference = 'Continue'
$report = New-Object System.Collections.Generic.List[string]
function W($s = '') { $report.Add([string]$s); Write-Host $s }
function Section($t) { W ''; W ('=' * 78); W "  $t"; W ('=' * 78) }

W 'ExcelSync mutual exclusion test'
W ("Time: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
W ("User: {0}\{1}   Machine: {2}   Session: {3}" -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME, $env:SESSIONNAME)
W "Mode: $Mode"
W "Target: $Path"

$results = New-Object System.Collections.Generic.List[object]
function Record($name, $ok, $detail) {
    $results.Add([pscustomobject]@{ Name = $name; OK = $ok; Detail = $detail })
    $mark = if ($ok -eq $true) { 'PASS' } elseif ($ok -eq $false) { 'FAIL' } else { 'INFO' }
    W ("  [{0}] {1}: {2}" -f $mark, $name, $detail)
}

# ------------------------------------------------------------
# Production-directory guard
# ------------------------------------------------------------
function Test-LooksLikeProduction([string]$p) {
    if ($p -like '*_exs_probe*' -or $p -like '*_exs_test*') { return $false }
    $d = Split-Path $p -Parent
    if (-not $d) { $d = $p }
    if (-not (Test-Path $d)) { return $false }
    $xlsx = Get-ChildItem $d -Filter *.xlsx -ErrorAction SilentlyContinue | Select-Object -First 1
    return ($null -ne $xlsx)
}

if (-not $Force -and (Test-LooksLikeProduction $Path)) {
    W ''
    W 'REFUSED: the target directory contains .xlsx files and does not look like a probe folder.'
    W '         This script creates and deletes files. Pick a scratch directory, e.g.'
    W ('           -Path "...\_exs_probe_<you>\lock.tst"')
    W '         Or pass -Force if you are certain.'
    $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8
    return
}

# ------------------------------------------------------------
# Win32 helpers (closest to real SMB semantics)
# ------------------------------------------------------------
$sig = @'
using System;
using System.Runtime.InteropServices;
public class SmbLock {
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern IntPtr CreateFileW(
        string lpFileName, uint dwDesiredAccess, uint dwShareMode,
        IntPtr lpSecurityAttributes, uint dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool CreateDirectoryW(string path, IntPtr sa);
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool RemoveDirectoryW(string path);

    public const uint GENERIC_READ  = 0x80000000;
    public const uint GENERIC_WRITE = 0x40000000;
    public const uint OPEN_EXISTING = 3;
    public const uint CREATE_NEW    = 1;
    public const uint FILE_ATTRIBUTE_NORMAL = 0x80;

    // write access requested, share mode 0 = exclusive
    public static IntPtr OpenExclusive(string path) {
        return CreateFileW(path, GENERIC_WRITE, 0, IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero);
    }
    // read access, share read+write+delete
    public static IntPtr OpenSharedRead(string path) {
        return CreateFileW(path, GENERIC_READ, 7, IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero);
    }
    public static IntPtr CreateNewExclusive(string path) {
        return CreateFileW(path, GENERIC_WRITE, 0, IntPtr.Zero, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero);
    }
    public static bool IsValid(IntPtr h) { return h != IntPtr.Zero && h.ToInt64() != -1; }
    public static int LastError() { return Marshal.GetLastWin32Error(); }
}
'@
$hasWin32 = $true
try {
    if (-not ('SmbLock' -as [type])) { Add-Type -TypeDefinition $sig -Language CSharp -ErrorAction Stop }
} catch {
    $hasWin32 = $false
    W ''
    W "WARNING: Add-Type failed: $($_.Exception.Message)"
    W '         Possibly Constrained Language Mode. Falling back to plain PowerShell tests.'
}

function ErrName([int]$e) {
    switch ($e) {
        0   { 'SUCCESS(0)' }
        5   { 'ERROR_ACCESS_DENIED(5)' }
        32  { 'ERROR_SHARING_VIOLATION(32) <- expected: already exclusively held' }
        80  { 'ERROR_FILE_EXISTS(80)' }
        183 { 'ERROR_ALREADY_EXISTS(183)' }
        2   { 'ERROR_FILE_NOT_FOUND(2)' }
        3   { 'ERROR_PATH_NOT_FOUND(3)' }
        53  { 'ERROR_BAD_NETPATH(53)' }
        64  { 'ERROR_NETNAME_DELETED(64)' }
        67  { 'ERROR_BAD_NET_NAME(67)' }
        default { "Win32Error=$e" }
    }
}

# ------------------------------------------------------------
if ($Mode -in 'Hold', 'Try') {
    Section "File exclusion test ($Mode)"

    $parent = Split-Path $Path -Parent
    if (-not (Test-Path $parent)) {
        W "Parent directory does not exist: $parent"
        W 'Create it first, e.g.:'
        W "  New-Item -ItemType Directory -Path '$parent' -Force"
        $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8
        return
    }

    if (-not $hasWin32) {
        W 'Cannot run without the Win32 helper. Aborting.'
        $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8
        return
    }

    if ($Mode -eq 'Hold') {
        Set-Content -Path $Path -Value "HELD_BY $env:USERNAME @ $env:COMPUTERNAME $(Get-Date -Format o)" -Encoding UTF8 -ErrorAction SilentlyContinue
        $h = [SmbLock]::OpenExclusive($Path)
        if (-not [SmbLock]::IsValid($h)) {
            W ("FAIL: cannot take exclusive access: {0}" -f (ErrName ([SmbLock]::LastError())))
            $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8
            return
        }
        W 'OK: I now hold exclusive access.'
        W ('  handle=0x{0:X}' -f $h.ToInt64())
        W ''
        W '  >>> On the OTHER machine, run:'
        W ('      .\4-test-lock.ps1 -Path "{0}" -Mode Try' -f $Path)
        W ''
        W ("  Holding for {0} seconds, do not close this window..." -f $Seconds)
        for ($i = $Seconds; $i -gt 0; $i--) {
            if ($i % 10 -eq 0) { Write-Host ("    remaining {0}s" -f $i) -ForegroundColor Cyan }
            Start-Sleep -Seconds 1
        }
        [void][SmbLock]::CloseHandle($h)
        W 'Lock released.'
        W ''
        W '  >>> If the colleague side reported SHARING VIOLATION, the lock plan holds.'
        W ('  >>> Cleanup: .\4-test-lock.ps1 -Path "{0}" -Mode Cleanup' -f $Path)
    }
    else {
        if (-not (Test-Path $Path)) {
            W "File does not exist: $Path"
            W 'Let side A run with -Mode Hold first (it creates the file).'
            $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8
            return
        }
        $h = [SmbLock]::OpenExclusive($Path)
        if ([SmbLock]::IsValid($h)) {
            W 'SEVERE: I also obtained exclusive access.'
            W '  -> this share does NOT enforce file sharing violations.'
            W '  -> the SMB byte-lock plan is unreliable; use a mkdir-based atomic lock.'
            Record 'file exclusion' $false 'side B also got exclusive access'
            [void][SmbLock]::CloseHandle($h)
        } else {
            $e = [SmbLock]::LastError()
            W ("OK: refused - {0}" -f (ErrName $e))
            if ($e -eq 32) {
                Record 'file exclusion' $true 'SHARING_VIOLATION(32): server enforces it, lock plan holds'
            } else {
                Record 'file exclusion' $null ("refused but with a non-typical code: " + (ErrName $e))
            }
        }

        # distinguish "exclusive" from "completely inaccessible"
        $hr = [SmbLock]::OpenSharedRead($Path)
        if ([SmbLock]::IsValid($hr)) {
            W 'INFO: a shared read is still allowed - normal SMB semantics (write excluded, read allowed)'
            [void][SmbLock]::CloseHandle($hr)
        } else {
            W ("INFO: even a shared read is refused: {0}" -f (ErrName ([SmbLock]::LastError())))
        }

        # CreateNew as an alternative primitive
        $cn = Join-Path $parent ('new_' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.tst')
        $hc = [SmbLock]::CreateNewExclusive($cn)
        if ([SmbLock]::IsValid($hc)) {
            W 'INFO: CreateNew (exclusive) works - usable as an alternative lock primitive'
            [void][SmbLock]::CloseHandle($hc)
            Remove-Item $cn -Force -ErrorAction SilentlyContinue
        } else {
            W ("INFO: CreateNew failed: {0}" -f (ErrName ([SmbLock]::LastError())))
        }
    }
}

# ------------------------------------------------------------
if ($Mode -in 'HoldDir', 'TryDir') {
    Section "Directory atomic-creation test ($Mode)"
    if (-not $hasWin32) {
        W 'Cannot run without the Win32 helper. Aborting.'
        $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8
        return
    }
    $probe = Join-Path $Path ('.lock_' + [guid]::NewGuid().ToString('N').Substring(0, 6))

    if ($Mode -eq 'HoldDir') {
        $heldName = $probe + '_HELD'
        $ok = [SmbLock]::CreateDirectoryW($heldName, [IntPtr]::Zero)
        if ($ok) {
            W "OK: created directory lock: $heldName"
            Set-Content (Join-Path $Path '.dirlockname') $heldName -ErrorAction SilentlyContinue
            W ''
            W '  >>> On the OTHER machine, run:'
            W ('      .\4-test-lock.ps1 -Path "{0}" -Mode TryDir' -f $Path)
            W ''
            W ("  Holding for {0} seconds..." -f $Seconds)
            Start-Sleep -Seconds $Seconds
            [void][SmbLock]::RemoveDirectoryW($heldName)
            W 'Released.'
        } else {
            W ("FAIL: cannot create directory: {0}" -f (ErrName ([SmbLock]::LastError())))
        }
    }
    else {
        $nameFile = Join-Path $Path '.dirlockname'
        if (-not (Test-Path $nameFile)) {
            W 'Cannot find .dirlockname. Let side A run -Mode HoldDir first.'
            $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8
            return
        }
        $target = (Get-Content $nameFile -Raw).Trim()
        W "Side A holds: $target"
        $ok = [SmbLock]::CreateDirectoryW($target, [IntPtr]::Zero)
        if ($ok) {
            W 'SEVERE: I could create it too -> mkdir is NOT atomic here, unusable as a lock'
            Record 'mkdir atomic lock' $false 'side B also succeeded'
            [void][SmbLock]::RemoveDirectoryW($target)
        } else {
            $e = [SmbLock]::LastError()
            W ("OK: refused - {0}" -f (ErrName $e))
            if ($e -eq 183 -or $e -eq 80) {
                Record 'mkdir atomic lock' $true 'ERROR_ALREADY_EXISTS: mkdir is a usable atomic lock'
            } else {
                Record 'mkdir atomic lock' $null ("refused but with a non-typical code: " + (ErrName $e))
            }
        }
    }
}

# ------------------------------------------------------------
if ($Mode -eq 'Cleanup') {
    Section 'Cleanup'
    foreach ($f in @($Path, (Join-Path (Split-Path $Path -Parent) '.dirlockname'))) {
        if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue; W "removed $f" }
    }
    Get-ChildItem (Split-Path $Path -Parent) -Filter '.lock_*' -Force -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue; W "removed $($_.FullName)" }
    Get-ChildItem (Split-Path $Path -Parent) -Filter 'new_*.tst' -Force -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue; W "removed $($_.FullName)" }
}

# ------------------------------------------------------------
Section 'Summary'
if ($results.Count -eq 0) { W '  (no verdict in this mode)' }
foreach ($r in $results) {
    $m = if ($r.OK -eq $true) { 'PASS' } elseif ($r.OK -eq $false) { 'FAIL' } else { 'INFO' }
    W ("  {0,-5} {1,-20} {2}" -f $m, $r.Name, $r.Detail)
}
W ''
W 'Please send this file back (one from each machine).'

$report -join "`r`n" | Set-Content -Path $OutFile -Encoding UTF8
Write-Host ''
Write-Host "Report saved: $OutFile" -ForegroundColor Green