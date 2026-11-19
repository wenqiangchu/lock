[CmdletBinding()]
param(
    [switch]$Worker,
    [string]$ResultPath,
    [string]$BaselineIds = "",
    [string]$TestMacro = "TestSuite.RunAllVbaTests",
    [int]$TimeoutSeconds = 120
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Read-TextIfPresent {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return "" }
    $text = Get-Content -LiteralPath $Path -Raw
    if ($null -eq $text) { return "" }
    return $text.Trim()
}

function Invoke-TestWorker {
    param([string]$OutputPath, [string]$KnownPids, [string]$MacroName)

    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class VbaTestNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

    $baseline = @()
    if (-not [string]::IsNullOrWhiteSpace($KnownPids)) {
        $baseline = @($KnownPids.Split(',') | ForEach-Object { [int]$_ })
    }
    $rootPath = Split-Path -Parent $PSScriptRoot
    $excel = $null
    $book = $null
    $components = $null
    $project = $null
    $ownedExcel = $false
    $excelPid = [uint32]0

    try {
        $excel = New-Object -ComObject Excel.Application
        [void][VbaTestNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
        if ($baseline -contains [int]$excelPid) {
            throw "Excel automation resolved to a pre-existing user process; tests stopped before workbook access."
        }
        $ownedExcel = $true
        $excel.AutomationSecurity = 1
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.ScreenUpdating = $false
        $book = $excel.Workbooks.Add()
        $project = $book.VBProject
        $components = $project.VBComponents
        $formsGuid = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
        $formsReference = $null
        foreach ($reference in $project.References) {
            if ([string]$reference.Guid -eq $formsGuid) { $formsReference = $reference; break }
        }
        if ($null -eq $formsReference) { [void]$project.References.AddFromGuid($formsGuid, 2, 0) }

        $sourceRoot = Join-Path $rootPath "src"
        $testRoot = Join-Path $rootPath "test"
        $files = @()
        foreach ($extension in @(".bas", ".cls", ".frm")) {
            $files += Get-ChildItem -LiteralPath $sourceRoot -Recurse -File -Filter "*$extension"
            $files += Get-ChildItem -LiteralPath $testRoot -Recurse -File -Filter "*$extension"
        }
        foreach ($file in ($files | Sort-Object FullName -Unique)) {
            if ($file.Extension -eq ".cls") {
                $component = $components.Add(2)
                $component.Name = [IO.Path]::GetFileNameWithoutExtension($file.Name)
                $source = Get-Content -LiteralPath $file.FullName -Raw
                $codeStart = $source.IndexOf("Option Explicit", [StringComparison]::OrdinalIgnoreCase)
                if ($codeStart -lt 0) { throw "Class module has no Option Explicit declaration: $($file.Name)" }
                $component.CodeModule.AddFromString($source.Substring($codeStart))
            }
            else {
                [void]$components.Import($file.FullName)
            }
        }

        $entryMacro = "'" + ([string]$book.Name).Replace("'", "''") + "'!TestSupport.RunSelectedTest"
        $testPassed = [bool]$excel.Run($entryMacro, $MacroName)
        if (-not $testPassed) {
            $failureMacro = "'" + ([string]$book.Name).Replace("'", "''") + "'!TestSupport.GetLastTestFailure"
            $testFailure = [string]$excel.Run($failureMacro)
            if ([string]::IsNullOrWhiteSpace($testFailure)) { $testFailure = "The test harness returned False without recording a failure." }
            throw "VBA test macro '$MacroName' failed. $testFailure"
        }
        $result = [ordered]@{
            status = "passed"
            excel_version = [string]$excel.Version
            excel_build = [string]$excel.Build
            excel_pid = [int]$excelPid
            imported_component_count = $components.Count
            test_macro = $MacroName
            test_status = "passed"
            workbook_saved = $false
        }
        [IO.File]::WriteAllText($OutputPath, ($result | ConvertTo-Json -Depth 4))
    }
    finally {
        Release-ComObject $components
        Release-ComObject $project
        if ($null -ne $book) {
            try { $book.Close($false) } catch { }
            Release-ComObject $book
        }
        if ($null -ne $excel) {
            if ($ownedExcel) {
                try { $excel.Quit() } catch { }
            }
            Release-ComObject $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if ($Worker) {
    try {
        Invoke-TestWorker -OutputPath $ResultPath -KnownPids $BaselineIds -MacroName $TestMacro
        exit 0
    }
    catch {
        $failureText = "{0}: {1}" -f $_.Exception.GetType().FullName, $_.Exception.Message
        if ([string]::IsNullOrWhiteSpace($failureText)) { $failureText = "The VBA test worker failed without an exception message." }
        try {
            $failureResult = [ordered]@{ status = "failed"; error = $failureText }
            [IO.File]::WriteAllText($ResultPath, ($failureResult | ConvertTo-Json -Depth 4))
        } catch { }
        [Console]::Error.WriteLine($failureText)
        exit 1
    }
}

$baselineProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
$baselineIdsArray = @($baselineProcesses | ForEach-Object { $_.Id })
$baselineIdsText = $baselineIdsArray -join ","
$tempRoot = [IO.Path]::GetTempPath()
$runId = [Guid]::NewGuid().ToString("N")
$resultPath = Join-Path $tempRoot ("mddt-vbatest-{0}.json" -f $runId)
$stderrPath = Join-Path $tempRoot ("mddt-vbatest-{0}.err" -f $runId)
$stdoutPath = Join-Path $tempRoot ("mddt-vbatest-{0}.out" -f $runId)
$workerProcess = $null
$workerResult = $null
$failure = $null
    $workerProcessId = $null

try {
    $powershellCommand = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($null -eq $powershellCommand) { $powershellCommand = Get-Command pwsh.exe -ErrorAction Stop }
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Worker -ResultPath `"$resultPath`" -BaselineIds `"$baselineIdsText`" -TestMacro `"$TestMacro`""
    $workerProcess = Start-Process -FilePath $powershellCommand.Source -ArgumentList $arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    if ($null -eq $workerProcess) { throw "The VBA test worker did not return a process handle." }
    $workerProcessId = $workerProcess.Id
    if (-not $workerProcess.WaitForExit($TimeoutSeconds * 1000)) {
        $failure = "VBA test execution exceeded the $TimeoutSeconds second hard timeout (worker PID $workerProcessId)."
        try { Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue } catch { }
        [void]$workerProcess.WaitForExit(5000)
    }
    else {
        $workerProcess.Refresh()
        $workerExitCode = $workerProcess.ExitCode
        $workerReportedPassed = $false
        if (Test-Path -LiteralPath $resultPath) {
            try {
                $workerResult = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
                $workerReportedPassed = ($workerResult.status -eq "passed")
            } catch { }
        }
        if (($null -eq $workerExitCode -and -not $workerReportedPassed) -or ($null -ne $workerExitCode -and $workerExitCode -ne 0)) {
            $detail = Read-TextIfPresent $stderrPath
            if (-not $detail) { $detail = Read-TextIfPresent $stdoutPath }
            if (-not $detail -and (Test-Path -LiteralPath $resultPath)) {
                try {
                    $workerResult = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
                    if ($workerResult.status -eq "failed") { $detail = [string]$workerResult.error }
                } catch { }
            }
            if (-not $detail) {
                $hasExited = if ($null -eq $workerProcess) { "missing" } else { [string]$workerProcess.HasExited }
                $resultExists = Test-Path -LiteralPath $resultPath
                $stdoutBytes = if (Test-Path -LiteralPath $stdoutPath) { (Get-Item -LiteralPath $stdoutPath).Length } else { 0 }
                $stderrBytes = if (Test-Path -LiteralPath $stderrPath) { (Get-Item -LiteralPath $stderrPath).Length } else { 0 }
                $detail = "Worker PID $workerProcessId exited=$hasExited code=$workerExitCode; result=$resultExists stdoutBytes=$stdoutBytes stderrBytes=$stderrBytes."
            }
            $failure = "VBA tests failed. $detail"
        }
    }

    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    do {
        $newExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $baselineIdsArray -notcontains $_.Id })
        if ($newExcel.Count -eq 0) { break }
        Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $deadline)

    $newExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $baselineIdsArray -notcontains $_.Id })
    if ($newExcel.Count -gt 0) {
        foreach ($processId in ($newExcel | ForEach-Object { $_.Id })) {
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
        $resultSummary = ""
        if ($null -ne $workerResult) {
            $resultSummary = " Worker status=$($workerResult.status), macro=$($workerResult.test_macro), test status=$($workerResult.test_status)."
        }
        if ($null -eq $failure) {
            $failure = "The isolated Excel process did not exit cleanly within 20 seconds and was terminated by the watchdog.$resultSummary"
        }
        else {
            $failure = "$failure The isolated Excel process also required watchdog termination.$resultSummary"
        }
    }

    if ($null -ne $failure) {
        $detail = Read-TextIfPresent $stderrPath
        if ($detail) { Write-Output $detail }
        throw $failure
    }
    if (-not (Test-Path -LiteralPath $resultPath)) { throw "VBA test process exited without producing a result." }
    Get-Content -LiteralPath $resultPath -Raw
}
finally {
    if ($null -ne $workerProcess) { $workerProcess.Dispose() }
    Remove-Item -LiteralPath $resultPath, $stderrPath, $stdoutPath -Force -ErrorAction SilentlyContinue
}
