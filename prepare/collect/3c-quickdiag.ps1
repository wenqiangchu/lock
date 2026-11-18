<#
    quickdiag.ps1  --  one-page diagnostic. Pure ASCII.

    Answers four things in about 30 seconds:
      1. are there leftover EXCEL.EXE processes holding locks?
      2. are there stale ~$ owner files on the share?
      3. what is the per-operation latency of the share?
      4. how long does one Excel open/save/close actually take?

    Run it with a CLEAN environment (no Excel running, probe dir empty).

    Usage:
        .\quickdiag.ps1 -File "X:\path\small_file.xlsx" -Probe "X:\_exs_probe_me"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$File,
    [Parameter(Mandatory = $true)][string]$Probe
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Probe)) {
    New-Item -ItemType Directory -Force -Path $Probe | Out-Null
}
Write-Host ('probe dir : ' + $Probe)
Write-Host ('source    : ' + $File)
Write-Host ('cwd       : ' + (Get-Location).Path)

Write-Host ''
Write-Host '=== 1. leftover Excel processes ==='
$procs = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
Write-Host ('  EXCEL.EXE count : ' + $procs.Count)
foreach ($p in $procs) {
    Write-Host ("    pid=" + $p.Id + "  title='" + $p.MainWindowTitle + "'")
}

Write-Host ''
Write-Host '=== 2. stale lock files ==='
$locks = @(Get-ChildItem -LiteralPath $Probe -Filter '~$*' -Force -ErrorAction SilentlyContinue)
Write-Host ('  ~$ count : ' + $locks.Count)
foreach ($l in $locks) { Write-Host ('    ' + $l.Name) }

Write-Host ''
Write-Host '=== 3. latency: 50 x create + delete ==='
$lat = @()
for ($i = 1; $i -le 50; $i++) {
    $f  = Join-Path $Probe ('lat_' + $i + '.tmp')
    $sw = [Diagnostics.Stopwatch]::StartNew()
    [IO.File]::WriteAllBytes($f, (New-Object byte[] 0))
    Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
    $sw.Stop()
    $lat += $sw.Elapsed.TotalMilliseconds
}
$sorted = @($lat | Sort-Object)
$avg    = ($lat | Measure-Object -Average).Average
Write-Host ("  min " + [int]$sorted[0] + " ms   avg " + [int]$avg + " ms   max " + [int]$sorted[$sorted.Count - 1] + " ms")
if ($avg -lt 15) { Write-Host '  -> latency healthy for a LAN share.' }
elseif ($avg -lt 50) { Write-Host '  -> latency elevated.' }
else { Write-Host '  -> latency HIGH: every small IO costs real time.' }

Write-Host ''
Write-Host '=== 4. Excel open / save / close (1 run) ==='
$dst = Join-Path $Probe ('qd_' + $env:USERNAME + '.xlsx')
if (-not (Test-Path -LiteralPath $dst)) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Copy-Item -LiteralPath $File -Destination $dst -Force
    $sw.Stop()
    Write-Host ('  upload to NEW name : ' + [int]$sw.Elapsed.TotalSeconds + ' s')
} else {
    Write-Host '  probe file already exists, skipping upload'
}
Write-Host ('  probe size : ' + [int]((Get-Item -LiteralPath $dst).Length / 1KB) + ' KB')

$xl = New-Object -ComObject Excel.Application
$xl.Visible       = $false
$xl.DisplayAlerts = $false
try {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $wb = $xl.Workbooks.Open($dst, 0, $false)
    $sw.Stop()
    Write-Host ('  open  : ' + [int]$sw.Elapsed.TotalSeconds + ' s')

    try {
        $wb.BuiltinDocumentProperties.Item('Comments').Value = 'probe ' + (Get-Date -Format 'HHmmss')
    }
    catch {
        Write-Host '  (could not set document property)'
    }
    Write-Host ('  was dirty before save : ' + (-not $wb.Saved))

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $wb.Save()
    $sw.Stop()
    Write-Host ('  SAVE  : ' + [int]$sw.Elapsed.TotalSeconds + ' s')
    Write-Host ('  clean after save      : ' + $wb.Saved)

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $wb.Close($false)
    $sw.Stop()
    Write-Host ('  close : ' + [int]$sw.Elapsed.TotalSeconds + ' s')
}
finally {
    try { $xl.Quit() } catch { }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ''
Write-Host '=== leftover after run ==='
$left = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
Write-Host ('  EXCEL.EXE count : ' + $left.Count)
$locks2 = @(Get-ChildItem -LiteralPath $Probe -Filter '~$*' -Force -ErrorAction SilentlyContinue)
Write-Host ('  ~$ count        : ' + $locks2.Count)
Write-Host ''
Write-Host 'done.'
