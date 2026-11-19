[CmdletBinding()]
param(
    [switch]$Worker,
    [string]$ResultPath,
    [string]$BaselineIds = "",
    [int]$TimeoutSeconds = 90
)

$ErrorActionPreference = "Stop"
$xlUp = -4162
$maxExcelRows = 1048576

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Set-SyntheticCell {
    param([object]$Sheet, [int]$Row, [int]$Column, [object]$Value)
    $cells = $null
    $cell = $null
    try {
        $cells = $Sheet.Cells
        $cell = $cells.Item($Row, $Column)
        $cell.Value2 = $Value
    }
    finally {
        Release-ComObject $cell
        Release-ComObject $cells
    }
}

function Get-CellValue {
    param([object]$Sheet, [int]$Row, [int]$Column)
    $cells = $null
    $cell = $null
    try {
        $cells = $Sheet.Cells
        $cell = $cells.Item($Row, $Column)
        return $cell.Value2
    }
    finally {
        Release-ComObject $cell
        Release-ComObject $cells
    }
}

function Get-LastDataRow {
    param([object]$Sheet, [int]$Column, [int]$BottomRow)
    $cells = $null
    $bottom = $null
    $last = $null
    try {
        $cells = $Sheet.Cells
        $bottom = $cells.Item($BottomRow, $Column)
        $last = $bottom.End($xlUp)
        return [int]$last.Row
    }
    finally {
        Release-ComObject $last
        Release-ComObject $bottom
        Release-ComObject $cells
    }
}

function Set-StyleOnlyCell {
    param([object]$Sheet, [int]$Row, [int]$Column)
    $cells = $null
    $cell = $null
    $interior = $null
    try {
        $cells = $Sheet.Cells
        $cell = $cells.Item($Row, $Column)
        $interior = $cell.Interior
        $interior.Color = 65535
    }
    finally {
        Release-ComObject $interior
        Release-ComObject $cell
        Release-ComObject $cells
    }
}

function Get-WorksheetRowCount {
    param([object]$Sheet)
    $rows = $null
    try {
        $rows = $Sheet.Rows
        return [int]$rows.Count
    }
    finally {
        Release-ComObject $rows
    }
}

function Test-BlankValue {
    param([object]$Value)
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return $Value.Length -eq 0 }
    return $false
}

function Assert-Condition {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Invoke-ProbeWorker {
    param([string]$OutputPath, [string]$KnownPids)

    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class BoundaryProbeNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

    $baseline = @()
    if (-not [string]::IsNullOrWhiteSpace($KnownPids)) {
        $baseline = @($KnownPids.Split(',') | ForEach-Object { [int]$_ })
    }

    $excel = $null
    $book = $null
    $sheetValid = $null
    $sheetOrphan = $null
    $ownedExcel = $false
    $excelPid = [uint32]0

    try {
        $excel = New-Object -ComObject Excel.Application
        [void][BoundaryProbeNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
        if ($baseline -contains [int]$excelPid) {
            throw "Excel automation resolved to a pre-existing user process; probe stopped before workbook access."
        }
        $ownedExcel = $true
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.ScreenUpdating = $false
        $version = [string]$excel.Version
        $build = [string]$excel.Build

        $book = $excel.Workbooks.Add()
        $sheetValid = $book.Worksheets.Item(1)
        $sheetOrphan = $book.Worksheets.Add()

        $headers = @{
            1 = "PRODUCT_A"
            3 = "Product_C"
            4 = "aMg001"
            5 = "Product_E"
            6 = "Product_F"
        }
        foreach ($column in $headers.Keys) {
            Set-SyntheticCell $sheetValid 1 ([int]$column) $headers[$column]
            Set-SyntheticCell $sheetOrphan 1 ([int]$column) $headers[$column]
        }

        foreach ($row in @(2, 4, 10)) {
            Set-SyntheticCell $sheetValid $row 1 "KEY_$row"
            Set-SyntheticCell $sheetValid $row 3 "GROUP_$row"
            Set-SyntheticCell $sheetValid $row 5 "PRODUCT_$row"
            Set-SyntheticCell $sheetValid $row 4 "VALUE_$row"
            Set-SyntheticCell $sheetOrphan $row 1 "KEY_$row"
            Set-SyntheticCell $sheetOrphan $row 3 "GROUP_$row"
            Set-SyntheticCell $sheetOrphan $row 5 "PRODUCT_$row"
            Set-SyntheticCell $sheetOrphan $row 4 "VALUE_$row"
        }
        Set-SyntheticCell $sheetOrphan 30 4 "ORPHAN_VALUE"
        Set-StyleOnlyCell $sheetValid $maxExcelRows 1
        Set-StyleOnlyCell $sheetValid $maxExcelRows 4

        $anchorColumns = @(1, 3, 5)
        $headerAnchorColumns = @()
        $productAfterE = 0
        for ($column = 1; $column -le 6; $column++) {
            $headerValue = Get-CellValue $sheetValid 1 $column
            if (Test-BlankValue $headerValue) { continue }
            if ([string]$headerValue -match '^(?i:AMG)') {
                continue
            }
            if ($column -le 5) { $headerAnchorColumns += $column }
            else { $productAfterE++ }
        }

        $validLastRows = @{}
        $columnsToScan = @()
        foreach ($column in $anchorColumns) { $columnsToScan += [int]$column }
        if ($columnsToScan -notcontains 4) { $columnsToScan += 4 }
        foreach ($column in $columnsToScan) {
            $validLastRows[[int]$column] = Get-LastDataRow $sheetValid ([int]$column) $maxExcelRows
        }
        $validMaxRow = ($validLastRows.Values | Measure-Object -Maximum).Maximum
        $orphanAnchorRows = @{}
        foreach ($column in $anchorColumns) {
            $orphanAnchorRows[[int]$column] = Get-LastDataRow $sheetOrphan ([int]$column) $maxExcelRows
        }
        $orphanEditLast = Get-LastDataRow $sheetOrphan 4 $maxExcelRows
        $orphanAnchorMax = ($orphanAnchorRows.Values | Measure-Object -Maximum).Maximum

        $internalBlankRows = 0
        for ($row = 2; $row -le 10; $row++) {
            $allRelevantBlank = $true
            foreach ($column in $columnsToScan) {
                $value = Get-CellValue $sheetValid $row ([int]$column)
                if (-not (Test-BlankValue $value)) { $allRelevantBlank = $false; break }
            }
            if ($allRelevantBlank) { $internalBlankRows++ }
        }

        $anchorlessDeclaredRows = @()
        for ($row = 2; $row -le $orphanEditLast; $row++) {
            $hasAnchor = $false
            foreach ($column in $anchorColumns) {
                $value = Get-CellValue $sheetOrphan $row ([int]$column)
                if (-not (Test-BlankValue $value)) { $hasAnchor = $true; break }
            }
            $editValue = Get-CellValue $sheetOrphan $row 4
            if (-not $hasAnchor -and -not (Test-BlankValue $editValue)) {
                $anchorlessDeclaredRows += $row
            }
        }

        Assert-Condition (($headerAnchorColumns -join ',') -eq '1,3,5') "A-E header classification did not match the controlled contract."
        Assert-Condition ($productAfterE -eq 1) "A product header after column E was not excluded from anchors."
        Assert-Condition ($validMaxRow -eq 10) "The styled worksheet tail changed the valid data end row."
        Assert-Condition ($internalBlankRows -eq 6) "Blank rows between records were not preserved in the physical row span."
        Assert-Condition ($orphanAnchorMax -eq 10 -and $orphanEditLast -eq 30) "The orphan-row fixture did not produce distinct anchor and edit boundaries."
        Assert-Condition (($anchorlessDeclaredRows -join ',') -eq '30') "A declared value without any anchor was not detected."

        $validLastRowsReport = [ordered]@{}
        foreach ($column in ($validLastRows.Keys | Sort-Object)) {
            $validLastRowsReport[[string]$column] = [int]$validLastRows[$column]
        }
        $report = [ordered]@{
            status = "passed"
            excel_version = $version
            excel_build = $build
            excel_pid = [int]$excelPid
            anchor_columns_a_to_e = $headerAnchorColumns
            product_columns_after_e = $productAfterE
            valid_last_row_by_anchor_and_edit_column = $validLastRowsReport
            valid_effective_last_row = [int]$validMaxRow
            preserved_internal_blank_rows = $internalBlankRows
            orphan_fixture_last_anchor_row = [int]$orphanAnchorMax
            orphan_fixture_last_declared_row = [int]$orphanEditLast
            orphan_declared_value_rows_rejected = $anchorlessDeclaredRows
            style_only_excel_max_row_ignored = $true
            workbook_saved = $false
        }
        [IO.File]::WriteAllText($OutputPath, ($report | ConvertTo-Json -Depth 5))
    }
    finally {
        Release-ComObject $sheetOrphan
        Release-ComObject $sheetValid
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
        Invoke-ProbeWorker -OutputPath $ResultPath -KnownPids $BaselineIds
        exit 0
    }
    catch {
        $failureText = "{0}: {1}" -f $_.Exception.GetType().FullName, $_.Exception.Message
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
$resultPath = Join-Path $tempRoot ("mddt-boundary-{0}.json" -f $runId)
$stderrPath = Join-Path $tempRoot ("mddt-boundary-{0}.err" -f $runId)
$stdoutPath = Join-Path $tempRoot ("mddt-boundary-{0}.out" -f $runId)
$workerProcess = $null
$failure = $null

try {
    $powershellCommand = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($null -eq $powershellCommand) { $powershellCommand = Get-Command pwsh.exe -ErrorAction Stop }
    $powershellPath = $powershellCommand.Source
    if ([string]::IsNullOrWhiteSpace($powershellPath)) { throw "No PowerShell executable path was resolved for the boundary worker." }
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Worker -ResultPath `"$resultPath`" -BaselineIds `"$baselineIdsText`""
    $workerProcess = Start-Process -FilePath $powershellPath -ArgumentList $arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    if ($null -eq $workerProcess) { throw "The boundary worker did not return a process handle." }
    if (-not $workerProcess.WaitForExit($TimeoutSeconds * 1000)) {
        $failure = "Excel boundary probe exceeded the $TimeoutSeconds second hard timeout."
        try { Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue } catch { }
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
            $detail = if (Test-Path -LiteralPath $stderrPath) { (Get-Content -LiteralPath $stderrPath -Raw).Trim() } else { "" }
            if (-not $detail -and (Test-Path -LiteralPath $stdoutPath)) { $detail = (Get-Content -LiteralPath $stdoutPath -Raw).Trim() }
            if (-not $detail -and (Test-Path -LiteralPath $resultPath)) {
                try {
                    $workerResult = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
                    if ($workerResult.status -eq "failed") { $detail = [string]$workerResult.error }
                } catch { }
            }
            if (-not $detail) { $detail = "Worker exited with code $workerExitCode and no diagnostic output." }
            $failure = "Excel boundary probe failed. $detail"
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
        $newIds = $newExcel | ForEach-Object { $_.Id }
        foreach ($processId in $newIds) {
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
        if ($null -eq $failure) { $failure = "The isolated Excel process did not exit cleanly and was terminated by the watchdog." }
    }

    if ($null -ne $failure) {
        if (Test-Path -LiteralPath $stderrPath) {
            $detail = (Get-Content -LiteralPath $stderrPath -Raw).Trim()
            if ($detail) { Write-Output $detail }
        }
        throw $failure
    }

    if (-not (Test-Path -LiteralPath $resultPath)) { throw "Excel probe exited without producing a result." }
    Get-Content -LiteralPath $resultPath -Raw
}
finally {
    if ($null -ne $workerProcess) { $workerProcess.Dispose() }
    Remove-Item -LiteralPath $resultPath, $stderrPath, $stdoutPath -Force -ErrorAction SilentlyContinue
}
