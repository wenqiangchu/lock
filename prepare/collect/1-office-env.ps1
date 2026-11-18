# ============================================================
#  1-office-env.ps1  -  Office / policy / environment collection
#
#  Read-only. Does not touch any network share.
#  Writes only: Desktop\exs_office.txt and a profile marker in %LOCALAPPDATA%.
#
#  Usage:
#     .\1-office-env.ps1
#     .\1-office-env.ps1 -CheckProfile     # after logoff + logon
#
#  NOTE: keep this file pure ASCII (PowerShell 5.1 reads BOM-less .ps1 as ANSI).
# ============================================================
[CmdletBinding()]
param(
    [switch]$CheckProfile,
    [string]$OutFile = "$([Environment]::GetFolderPath('Desktop'))\exs_office.txt"
)

$ErrorActionPreference = 'Continue'
$report = New-Object System.Collections.Generic.List[string]
function W($s = '') { $report.Add([string]$s) }
function Section($t) { W ''; W ('=' * 78); W "  $t"; W ('=' * 78) }
function Try_($label, $block) {
    try { $r = & $block; W ("{0}: {1}" -f $label, $r) }
    catch { W ("{0}: <ERR> {1}" -f $label, $_.Exception.Message) }
}

W 'ExcelSync environment report'
W ("Generated: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
W ("Computer: {0}    User: {1}\{2}" -f $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME)
W ("PowerShell: {0}" -f $PSVersionTable.PSVersion.ToString())

# ------------------------------------------------------------
Section 'A. Session / user profile (can a local working copy survive?)'
# ------------------------------------------------------------
W "SESSIONNAME        : $env:SESSIONNAME      (prefix RDP- means remote desktop session)"
W "CLIENTNAME         : $env:CLIENTNAME"
W "LOCALAPPDATA       : $env:LOCALAPPDATA"
W "APPDATA            : $env:APPDATA"
W "USERPROFILE        : $env:USERPROFILE"
Try_ 'PROCESSOR_ARCHITECTURE' { $env:PROCESSOR_ARCHITECTURE }
Try_ 'OS' { (Get-CimInstance Win32_OperatingSystem).OperatingSystemProperties = $null; (Get-CimInstance Win32_OperatingSystem).Caption }
Try_ 'SessionId' { (Get-Process -Id $PID).SessionId }

Section 'A2. Profile type (roaming / FSLogix / mandatory / local)'
Try_ 'USERPROFILE from Volatile Environment' {
    $va = Get-ItemProperty 'HKCU:\Volatile Environment' -ErrorAction SilentlyContinue
    "$($va.USERPROFILE)   HOMESHARE=$($va.HOMESHARE)"
}
Try_ 'ProfileImagePath' {
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction SilentlyContinue).ProfileImagePath
}
Try_ 'FSLogix installed' {
    $p = Get-Service -Name 'frxsvc' -ErrorAction SilentlyContinue
    if ($p) { "yes (status=$($p.Status))" } else { 'no' }
}
Try_ 'Mandatory profile hint' {
    $k = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    $m = (Get-ChildItem $k -ErrorAction SilentlyContinue | ForEach-Object {
        (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).State } | Where-Object { $_ -band 1 })
    if ($m) { 'possible mandatory profile (State bit 1 set)' } else { 'no mandatory flag seen' }
}

$marker = Join-Path $env:LOCALAPPDATA 'ExcelSync\_profile_marker.txt'
if ($CheckProfile) {
    Section 'A3. Profile persistence check (-CheckProfile mode)'
    if (Test-Path $marker) {
        W 'Marker file exists -> LOCALAPPDATA survived logoff/logon (persistent profile)'
        W ("  content: {0}" -f (Get-Content $marker -Raw))
    } else {
        W 'Marker file missing -> LOCALAPPDATA was cleared by logoff (NOT persistent)'
        W '  Impact: the local working copy would be lost and uncommitted edits too.'
    }
} else {
    Section 'A3. Profile persistence check (step 1: write marker)'
    try {
        New-Item -ItemType Directory -Path (Split-Path $marker) -Force | Out-Null
        "marker created $(Get-Date -Format o) by $env:USERNAME" | Set-Content $marker -Encoding UTF8
        W "Marker written: $marker"
        W 'Now LOG OFF and LOG BACK ON, then run:  .\1-office-env.ps1 -CheckProfile'
    } catch { W "Failed to write marker: $($_.Exception.Message)" }
}

# ------------------------------------------------------------
Section 'B. Office / Excel version and bitness'
# ------------------------------------------------------------
Try_ 'EXCEL.EXE path' {
    $paths = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\EXCEL.EXE"
    )
    $found = $paths | Where-Object { Test-Path $_ }
    if ($found) { $found -join '; ' } else {
        (Get-ChildItem 'C:\Program Files\Microsoft Office','C:\Program Files (x86)\Microsoft Office' -Filter EXCEL.EXE -Recurse -ErrorAction SilentlyContinue | Select-Object -First 3).FullName -join '; '
    }
}
Try_ 'EXCEL.EXE file version' {
    $f = Get-ChildItem 'C:\Program Files\Microsoft Office','C:\Program Files (x86)\Microsoft Office' -Filter EXCEL.EXE -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($f) { $f.VersionInfo.FileVersion } else { 'not found' }
}
Try_ 'Office ClickToRun (bitness)' {
    $c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path $c2r) {
        $p = Get-ItemProperty $c2r
        "Platform=$($p.Platform)  ProductReleaseIds=$($p.ProductReleaseIds)  VersionToReport=$($p.VersionToReport)"
    } else { 'not a ClickToRun installation' }
}
Try_ 'Registered Office versions (HKCU)' {
    $v = Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
         Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | Select-Object -ExpandProperty PSChildName
    if ($v) { $v -join ', ' } else { 'none' }
}
Try_ 'OS is 64-bit' { [Environment]::Is64BitOperatingSystem }
Try_ 'PowerShell is 64-bit' { [Environment]::Is64BitProcess }

# ------------------------------------------------------------
Section 'C. Macro security policy (decides whether deployment is possible)'
# ------------------------------------------------------------
$explain = @{
    '1' = 'enable all macros (usable, but lax)'
    '2' = 'disable with notification  <-- needs trusted location to avoid prompts'
    '3' = 'only signed macros  <-- SEVERE: certificate required'
    '4' = 'disable all, no notification  <-- SEVERE: plan not deployable'
}
foreach ($k in @(
    @{ N='VBAWarnings (HKCU 16.0)'; P='HKCU:\Software\Microsoft\Office\16.0\Excel\Security' },
    @{ N='VBAWarnings (HKLM 16.0)'; P='HKLM:\Software\Microsoft\Office\16.0\Excel\Security' },
    @{ N='VBAWarnings (HKCU 15.0)'; P='HKCU:\Software\Microsoft\Office\15.0\Excel\Security' }
)) {
    if (Test-Path $k.P) {
        $val = (Get-ItemProperty $k.P -ErrorAction SilentlyContinue).VBAWarnings
        $e = if ($null -ne $val -and $explain.ContainsKey("$val")) { $explain["$val"] } else { '(unset or unknown)' }
        W ("{0}: {1}  {2}" -f $k.N, $(if ($null -ne $val) { $val } else { '<unset>' }), $e)
    } else {
        W ("{0}: <key not present>" -f $k.N)
    }
}
Try_ 'AccessVBOM (programmatic access to VBA project)' {
    $p = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    if (Test-Path $p) { (Get-ItemProperty $p -ErrorAction SilentlyContinue).AccessVBOM } else { '<none>' }
}
Try_ 'Trusted Publishers count' {
    $p = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Publishers'
    if (Test-Path $p) { (Get-ChildItem $p).Count } else { 0 }
}
Try_ 'ActiveX / Forms controls policy' {
    $p = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    $r = Get-ItemProperty $p -ErrorAction SilentlyContinue
    "DisableAllActiveX=$($r.DisableAllActiveX)  LoadControlsInForms=$($r.LoadControlsInForms)"
}
Try_ 'BlockContentExecutionFromInternet' {
    $p = 'HKCU:\Software\Microsoft\Office\16.0\Common\Security'
    if (Test-Path $p) { (Get-ItemProperty $p -ErrorAction SilentlyContinue).BlockContentExecutionFromInternet } else { '<none>' }
}
Try_ 'Group policy macro keys (HKLM Policies)' {
    $p = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\excel\security'
    if (Test-Path $p) { (Get-ItemProperty $p | Format-List | Out-String) } else { '<no policy key>' }
}

# ------------------------------------------------------------
Section 'D. Trusted locations'
# ------------------------------------------------------------
$tlRoot = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations'
if (-not (Test-Path $tlRoot)) { $tlRoot = 'HKCU:\Software\Microsoft\Office\15.0\Excel\Security\Trusted Locations' }
if (Test-Path $tlRoot) {
    Try_ 'AllowNetworkLocations (1 = UNC paths may be trusted)' {
        $v = (Get-ItemProperty $tlRoot -ErrorAction SilentlyContinue).AllowNetworkLocations
        if ($null -eq $v) {
            '<unset> (default: UNC paths cannot be added as trusted locations)'
        } else { "$v $(if ($v -eq 1) { 'allowed' } else { 'not allowed' })" }
    }
    $locs = Get-ChildItem $tlRoot -ErrorAction SilentlyContinue
    W ''
    W "Configured trusted locations: $($locs.Count)"
    foreach ($l in $locs) {
        $p = Get-ItemProperty $l.PSPath -ErrorAction SilentlyContinue
        W ("  [{0}] {1}" -f $l.PSChildName, $p.Path)
        W ("       AllowSubfolders={0}" -f $p.AllowSubfolders)
    }
    W ''
    W 'Check manually: does the list contain your share path?'
    W ''
    W 'Note: files on a network share do not automatically open in Protected View.'
    W '      It depends on Mark-of-the-Web and the Protected View switches.'
    W '      3-measure.ps1 measures whether the share copy can be opened read-write.'
} else {
    W 'Trusted Locations key not found (never configured).'
}
Try_ 'Protected View switches' {
    $p = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView'
    if (Test-Path $p) {
        $r = Get-ItemProperty $p
        "DisableInternetFilesInPV=$($r.DisableInternetFilesInPV)  DisableUnsafeLocationsInPV=$($r.DisableUnsafeLocationsInPV)  DisableAttachmentsInPV=$($r.DisableAttachmentsInPV)"
    } else { '<none>' }
}

# ------------------------------------------------------------
Section 'E. Antivirus / realtime scanning (can cause intermittent Save failures)'
# ------------------------------------------------------------
Try_ 'Registered antivirus products' {
    $av = Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
    if ($av) { ($av | ForEach-Object { "$($_.displayName) [state=$($_.productState)]" }) -join ' | ' } else { 'none detected (may be hidden by policy)' }
}
Try_ 'Defender realtime protection' {
    $st = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($st) { "RealTimeProtectionEnabled=$($st.RealTimeProtectionEnabled)" } else { 'Get-MpComputerStatus unavailable' }
}
Try_ 'Defender exclusion paths' {
    $m = Get-MpPreference -ErrorAction SilentlyContinue
    if ($m) { ($m.ExclusionPath -join '; ') } else { '<unavailable>' }
}
Try_ 'Defender network protection' {
    $m = Get-MpPreference -ErrorAction SilentlyContinue
    if ($m) { $m.EnableNetworkProtection } else { '<unavailable>' }
}

# ------------------------------------------------------------
Section 'F. Network share / SMB'
# ------------------------------------------------------------
Try_ 'Mapped drives' {
    $d = Get-SmbMapping -ErrorAction SilentlyContinue
    if ($d) { ($d | ForEach-Object { "$($_.LocalPath) -> $($_.RemotePath) [$($_.Status)]" }) -join ' | ' } else { 'none (UNC used directly?)' }
}
Try_ 'net use' { (net use) -join "`n" }
Try_ 'SMB client configuration' {
    $c = Get-SmbClientConfiguration -ErrorAction SilentlyContinue
    if ($c) { "EnableSecuritySignature=$($c.EnableSecuritySignature)  RequireSecuritySignature=$($c.RequireSecuritySignature)  EnableOplocks=$($c.EnableOplocks)" } else { '<unavailable>' }
}
Try_ 'SMB connections (visit the share first)' {
    $c = Get-SmbConnection -ErrorAction SilentlyContinue
    if ($c) { ($c | ForEach-Object { "$($_.ServerName) dialect=$($_.Dialect) share=$($_.ShareName)" }) -join ' | ' } else { 'none' }
}
Try_ 'Offline Files / CSC' {
    $svc = Get-Service CSC -ErrorAction SilentlyContinue
    $p = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\CSC\Parameters' -ErrorAction SilentlyContinue
    "CSC service=$($svc.Status)"
}
Try_ 'Interfaces with a gateway (context only)' {
    $ip = Get-NetIPConfiguration -ErrorAction SilentlyContinue | Where-Object { $_.IPv4DefaultGateway } | Select-Object -First 1
    if ($ip) { "$($ip.InterfaceAlias) gateway=$($ip.IPv4DefaultGateway.NextHop)" } else { '<none>' }
}

# ------------------------------------------------------------
Section 'G. Excel add-ins currently present'
# ------------------------------------------------------------
Try_ 'HKCU Excel Addins' {
    $p = 'HKCU:\Software\Microsoft\Office\Excel\Addins'
    if (Test-Path $p) {
        (Get-ChildItem $p | ForEach-Object {
            $v = Get-ItemProperty $_.PSPath
            "$($_.PSChildName): LoadBehavior=$($v.LoadBehavior) Desc=$($v.FriendlyName)"
        }) -join "`n"
    } else { '<none>' }
}
Try_ 'XLSTART contents' {
    $p = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
    if (Test-Path $p) { (Get-ChildItem $p).Name -join ', ' } else { '<missing>' }
}
Try_ 'Existing .xlam files' {
    $p = Join-Path $env:APPDATA 'Microsoft\AddIns'
    if (Test-Path $p) { (Get-ChildItem $p -Filter *.xla* -ErrorAction SilentlyContinue).Name -join ', ' } else { '<AddIns folder missing>' }
}

# ------------------------------------------------------------
Section 'H. Disk space (local working copy needs 2-3x the file size)'
# ------------------------------------------------------------
Try_ 'Drive holding LOCALAPPDATA' {
    $d = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $env:LOCALAPPDATA -like "$($_.Root)*" }
    if ($d) { "drive=$($d.Name) free=$([math]::Round($d.Free/1GB,2)) GB" } else { 'not located' }
}
Try_ 'Temp folder' {
    $q = Split-Path $env:TEMP -Qualifier
    $drv = $q.TrimEnd(':')
    $free = (Get-PSDrive -Name $drv -ErrorAction SilentlyContinue).Free
    if ($null -ne $free) { "$env:TEMP  free=$([math]::Round($free/1GB,2)) GB" } else { "$env:TEMP  (could not read free space)" }
}

# ------------------------------------------------------------
Section 'I. Script execution policy'
# ------------------------------------------------------------
Try_ 'ExecutionPolicy' { (Get-ExecutionPolicy -List | Out-String).Trim() }

# Output
$report -join "`r`n" | Set-Content -Path $OutFile -Encoding UTF8
Write-Host ''
Write-Host "Report saved: $OutFile" -ForegroundColor Green
if (-not $CheckProfile) {
    Write-Host 'Next: log off, log back on, then run  .\1-office-env.ps1 -CheckProfile' -ForegroundColor Yellow
}
Write-Host 'Please send this file back.' -ForegroundColor Yellow