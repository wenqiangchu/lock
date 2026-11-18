# ============================================================
#  5-test-replace.ps1  -  can we do "write temp + rename swap"?
#
#  This is the architectural upgrade. If it works, the commit no longer locks
#  the file at all: readers may keep it open forever and the commit window
#  drops from ~20 s to ~2 s.
#
#  Principle: POSIX rename() semantics - replacing a directory entry does not
#  affect an inode already open by someone else. Samba's "strict rename"
#  setting influences this (default "no" normally allows it).
#
#  SAFETY: never writes directly into -Dir. It creates and works only inside
#          <Dir>\_exs_probe_<user>\.
#
#  Usage:
#     1) .\5-test-replace.ps1 -Mode Prepare -Dir "\\SERVER\SHARE\_exs_probe_me"
#     2) open a.xlsx in Excel from the probe folder, keep it OPEN (read-write)
#     3) .\5-test-replace.ps1 -Mode Replace -Dir "\\SERVER\SHARE\_exs_probe_me"
#     4) close Excel, then
#        .\5-test-replace.ps1 -Mode Verify  -Dir "\\SERVER\SHARE\_exs_probe_me"
#        .\5-test-replace.ps1 -Mode Cleanup -Dir "\\SERVER\SHARE\_exs_probe_me"
#
#  NOTE: keep this file pure ASCII (PowerShell 5.1 reads BOM-less .ps1 as ANSI).
# ============================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [ValidateSet('Prepare', 'Replace', 'Verify', 'Cleanup')][string]$Mode = 'Prepare',
    [string]$OutFile = "$([Environment]::GetFolderPath('Desktop'))\exs_replace.txt"
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
function W($s = '') { $report.Add([string]$s); Write-Host $s }
function Section($t) { W ''; W ('=' * 78); W "  $t"; W ('=' * 78) }

if (-not (Test-Path $Dir)) {
    Write-Host "Directory does not exist: $Dir" -ForegroundColor Red
    return
}

# always work inside a dedicated probe subfolder
$Dir = ResolveFull $Dir
$probe = Join-Path $Dir ('_exs_probe_' + ($env:USERNAME -replace '[^\w]', '') + '_swap')

$a = Join-Path $probe 'a.xlsx'
$b = Join-Path $probe 'b.xlsx'

W 'ExcelSync rename-swap feasibility test'
W ("Time: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
W ("User: {0}\{1}   Machine: {2}" -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME)
W "Mode: $Mode"
W "Probe folder: $probe"

# ------------------------------------------------------------
# Win32 MoveFileEx
# ------------------------------------------------------------
$sig = @'
using System;
using System.Runtime.InteropServices;
public class Mv {
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool MoveFileExW(string existing, string newName, uint flags);

    public const uint MOVEFILE_REPLACE_EXISTING = 0x1;
    public const uint MOVEFILE_WRITE_THROUGH    = 0x8;
    public const uint MOVEFILE_COPY_ALLOWED     = 0x2;

    // Do NOT p/invoke GetLastError() here: the managed call would overwrite it.
    // Marshal.GetLastWin32Error() must be read immediately after the call.
    public static int LastError;

    public static bool Replace(string src, string dst) {
        bool ok = MoveFileExW(src, dst,
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH | MOVEFILE_COPY_ALLOWED);
        LastError = ok ? 0 : Marshal.GetLastWin32Error();
        return ok;
    }

    public static string Err(int e) {
        switch (e) {
            case 0:   return "SUCCESS(0)";
            case 5:   return "ACCESS_DENIED(5) <- may be busy, or a real permission problem";
            case 32:  return "SHARING_VIOLATION(32) <- target is in use, replace refused";
            case 2:   return "FILE_NOT_FOUND(2)";
            case 3:   return "PATH_NOT_FOUND(3)";
            case 183: return "ALREADY_EXISTS(183)";
            case 123: return "INVALID_NAME(123)";
            case 17:  return "NOT_SAME_DEVICE(17) <- cross-volume, cannot rename";
            default:  return "Win32Error=" + e;
        }
    }
}
'@
$hasWin32 = $true
try {
    if (-not ('Mv' -as [type])) { Add-Type -TypeDefinition $sig -Language CSharp -ErrorAction Stop }
} catch {
    $hasWin32 = $false
    W ''
    W "WARNING: Add-Type failed: $($_.Exception.Message)"
    W '         The rename test cannot run. Report this.'
}

function NewTestBook([string]$path, [string]$marker) {
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false; $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Cells.Item(1, 1).Value2 = 'MARKER'
        $ws.Cells.Item(1, 2).Value2 = $marker
        $ws.Cells.Item(2, 1).Value2 = 'CREATED'
        $ws.Cells.Item(2, 2).Value2 = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')
        $wb.SaveAs($path, 51)   # 51 = xlOpenXMLWorkbook
        $wb.Close($false)
        $xl.Quit()
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        return $true
    } catch {
        W ("  WARNING: Excel COM unavailable ({0}); using a plain text placeholder" -f $_.Exception.Message)
        Set-Content -Path $path -Value ("MARKER=$marker`nCREATED=$(Get-Date -Format o)") -Encoding UTF8
        return $false
    }
}

# ------------------------------------------------------------
if ($Mode -eq 'Prepare') {
    Section 'Prepare'
    if (-not (Test-Path $probe)) { New-Item -ItemType Directory -Path $probe -Force | Out-Null }
    $e1 = NewTestBook $a 'OLD'
    W ("  created {0}  (marker=OLD, excel={{0}})" -f $a) -replace '\{0\}', $e1
    $e2 = NewTestBook $b 'NEW'
    W ("  created {0}  (marker=NEW, excel={1})" -f $b, $e2)
    W ''
    W '  >>> Now open a.xlsx in Excel FROM THE PROBE FOLDER and keep it OPEN'
    W '      (read-write, not read-only).'
    W ''
    W '  >>> With Excel still open, run:'
    W ('      .\5-test-replace.ps1 -Mode Replace -Dir "{0}"' -f $Dir)
    W ''
    W '  Expected: either MoveFileEx reports a sharing/access error (regex swap not supported)'
    W '            or it succeeds (rename-swap architecture is available).'
}

if ($Mode -eq 'Replace') {
    Section 'Replace: try to replace a file that is open in Excel'
    if (-not (Test-Path $a)) { W "Missing $a - run -Mode Prepare first"; $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8; return }
    if (-not (Test-Path $b)) { W "Missing $b - run -Mode Prepare first"; $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8; return }
    if (-not $hasWin32) { W 'Win32 helper unavailable, cannot test.'; $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8; return }

    $lockName = '~$' + [IO.Path]::GetFileName($a)
    $lockFull = Join-Path $probe $lockName
    $lockExists = Test-Path $lockFull
    W ("  lock file {0}: {1}" -f $lockName, $(if ($lockExists) { 'present (a.xlsx is open)' } else { 'absent (WARNING: is a.xlsx really open?)' }))
    if (-not $lockExists) {
        W ''
        W '  WARNING: cannot confirm a.xlsx is open. If you are sure it is, Excel may not'
        W '           create a ~$ file in this configuration. The test can still run but'
        W '           interpret the result carefully.'
    }

    $tmpUpload = Join-Path $probe ('upload_' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.xlsx')
    W ''
    W '  Step 1: copy b.xlsx to a temp upload file (simulates uploading a new version)'
    Copy-Item $b $tmpUpload -Force -ErrorAction SilentlyContinue
    W ("    -> {0}  exists={1}" -f $tmpUpload, (Test-Path $tmpUpload))

    W ''
    W '  Step 2: MoveFileEx(REPLACE_EXISTING) onto a.xlsx'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $ok = [Mv]::Replace($tmpUpload, $a)
    $sw.Stop()
    $err = [Mv]::Err([Mv]::LastError)
    W ("    result: {0}   elapsed={1:N3} s   error={2}" -f $(if ($ok) { 'SUCCESS' } else { 'FAILED' }), $sw.Elapsed.TotalSeconds, $err)

    if ($ok) {
        W ''
        $fi = Get-Item $a
        W ("    a.xlsx now: size={0}  mtime={1}" -f $fi.Length, $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))
        W ''
        W '  >>> Look at the Excel window that has a.xlsx open:'
        W '      - no error and content still OLD -> POSIX rename semantics, as expected'
        W '      - "file changed / reload" prompt   -> also acceptable'
        W '      - "file corrupt" or Excel crash    -> this architecture is NOT safe'
        W ''
        W '  >>> Close Excel WITHOUT saving, then run:'
        W ('      .\5-test-replace.ps1 -Mode Verify -Dir "{0}"' -f $Dir)
    } else {
        W ''
        W '  Two cases must be distinguished:'
        W '    A) code 32 (SHARING_VIOLATION) -> clearly "in use"; conservative design only.'
        W '    B) code  5 (ACCESS_DENIED)     -> may be "in use", or no write permission.'
        W ''
        W '  >>> To distinguish: close Excel to release a.xlsx, then rerun:'
        W ('      .\5-test-replace.ps1 -Mode Replace -Dir "{0}"' -f $Dir)
        W '      - succeeds now -> it really was "in use" -> conservative design'
        W '      - still fails  -> permission problem or Samba forbids rename'
        W ''
        W '  -> Conservative design: Workbook.Open + Save (commit window 10-20 s).'
        W '     Must detect who holds the file and show a clear message.'
        Remove-Item $tmpUpload -Force -ErrorAction SilentlyContinue
    }
}

if ($Mode -eq 'Verify') {
    Section 'Verify: confirm the replacement'
    if (-not (Test-Path $a)) { W 'File does not exist'; $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8; return }
    W "  file: $a"
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false; $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open($a, 0, $true)
        $ws = $wb.Worksheets.Item(1)
        $m1 = $ws.Cells.Item(1, 1).Value2
        $m2 = $ws.Cells.Item(1, 2).Value2
        W ("    A1='{0}'  B1='{1}'" -f $m1, $m2)
        if ($m2 -eq 'NEW') {
            W ''
            W '    RESULT: content is NEW -> rename swap worked.'
            W '            Adopt it: no lock on book.xlsx, commit window ~2 s.'
        } elseif ($m2 -eq 'OLD') {
            W ''
            W '    RESULT: content is still OLD -> replacement did not take effect.'
        } else {
            W ("    INFO: unexpected content: {0}" -f $m2)
        }
        $wb.Close($false); $xl.Quit()
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    } catch { W ("    read failed: {0}" -f $_.Exception.Message) }

    W ''
    W '  Probe folder contents:'
    Get-ChildItem $probe -Force -ErrorAction SilentlyContinue | ForEach-Object {
        W ("    {0,-42} {1,12:N0}  {2}" -f $_.Name, $_.Length, $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))
    }
}

if ($Mode -eq 'Cleanup') {
    Section 'Cleanup'
    if (Test-Path $probe) {
        $locks = Get-ChildItem $probe -Filter '~$*' -Force -ErrorAction SilentlyContinue
        if ($locks) {
            foreach ($l in $locks) { W ("  WARNING: lock file still present: {0} (close Excel first)" -f $l.Name) }
        }
        $ans = Read-Host "Delete the probe folder?" + " ($probe) [y/N]"
        if ($ans -match '^[Yy]') {
            Remove-Item $probe -Recurse -Force -ErrorAction SilentlyContinue
            W "  removed $probe"
        } else { W '  kept.' }
    } else { W '  probe folder does not exist.' }
}

# ------------------------------------------------------------
Section 'How to interpret'
W 'If MoveFileEx succeeds while the file is open:'
W '  -> architecture A (rename swap):'
W '       1. download book.xlsx locally'
W '       2. patch locally, save as a new file'
W '       3. upload to _sync/upload/<uuid>.xlsx'
W '       4. MoveFileEx(upload, book.xlsx)'
W '     Pro: no lock on the data file, readers never block you, ~2 s commit'
W '     Con: a reader who SAVES will overwrite your commit (readers must be read-only)'
W ''
W 'If MoveFileEx fails:'
W '  -> architecture B (conservative):'
W '       1. acquire commit.lock'
W '       2. Workbook.Open(UNC)   <- failure means someone holds it -> prompt retry'
W '       3. Copy region -> Save -> Close'
W '       4. release lock'
W '     Commit window 10-20 s; must handle the "reader blocks commit" experience.'

$report -join "`r`n" | Set-Content -Path $OutFile -Encoding UTF8
Write-Host ''
Write-Host "Report saved: $OutFile" -ForegroundColor Green