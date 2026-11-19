[CmdletBinding()]
param(
    [string]$OutputPath = "",
    [int]$TimeoutSeconds = 120,
    [switch]$Worker,
    [string]$ResultPath = "",
    [string]$BaselineIds = ""
)

$ErrorActionPreference = "Stop"
$script:RibbonContentType = "application/xml"

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-XmlEntry {
    param([IO.Compression.ZipArchive]$Archive, [string]$EntryName)

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) { throw "Required package part is missing: $EntryName" }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { $content = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $document = [System.Xml.XmlDocument]::new()
    $document.PreserveWhitespace = $true
    $document.LoadXml($content)
    return $document
}

function Set-XmlEntry {
    param([IO.Compression.ZipArchive]$Archive, [string]$EntryName, [System.Xml.XmlDocument]$Document)

    $existing = $Archive.GetEntry($EntryName)
    if ($null -ne $existing) { $existing.Delete() }
    $entry = $Archive.CreateEntry($EntryName, [IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    try { $Document.Save($writer) } finally { $writer.Dispose() }
}

function Add-RibbonPart {
    param([string]$PackagePath, [string]$RibbonPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $ribbonDocument = [System.Xml.XmlDocument]::new()
    $ribbonDocument.PreserveWhitespace = $true
    $ribbonDocument.Load($RibbonPath)
    $archive = [IO.Compression.ZipFile]::Open($PackagePath, [IO.Compression.ZipArchiveMode]::Update)
    try {
        Set-XmlEntry $archive "customUI/customUI14.xml" $ribbonDocument

        $contentTypes = Get-XmlEntry $archive "[Content_Types].xml"
        $override = $contentTypes.SelectSingleNode("/*[local-name()='Types']/*[local-name()='Override' and @PartName='/customUI/customUI14.xml']")
        if ($null -eq $override) {
            $override = $contentTypes.CreateElement("Override", $contentTypes.DocumentElement.NamespaceURI)
            [void]$override.SetAttribute("PartName", "/customUI/customUI14.xml")
            [void]$override.SetAttribute("ContentType", $script:RibbonContentType)
            [void]$contentTypes.DocumentElement.AppendChild($override)
        }
        else {
            [void]$override.SetAttribute("ContentType", $script:RibbonContentType)
        }
        Set-XmlEntry $archive "[Content_Types].xml" $contentTypes

        $relationships = Get-XmlEntry $archive "_rels/.rels"
        $relationshipType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
        $relationship = $relationships.SelectSingleNode("/*[local-name()='Relationships']/*[local-name()='Relationship' and @Type='$relationshipType']")
        if ($null -eq $relationship) {
            $usedIds = @($relationships.SelectNodes("/*[local-name()='Relationships']/*[local-name()='Relationship']") | ForEach-Object { $_.GetAttribute("Id") })
            $suffix = 1
            do { $relationshipId = "rIdMddtRibbon$suffix"; $suffix++ } while ($usedIds -contains $relationshipId)
            $relationship = $relationships.CreateElement("Relationship", $relationships.DocumentElement.NamespaceURI)
            [void]$relationship.SetAttribute("Id", $relationshipId)
            [void]$relationship.SetAttribute("Type", $relationshipType)
            [void]$relationships.DocumentElement.AppendChild($relationship)
        }
        [void]$relationship.SetAttribute("Target", "customUI/customUI14.xml")
        Set-XmlEntry $archive "_rels/.rels" $relationships
    }
    finally {
        $archive.Dispose()
    }
}

function Test-AddinPackage {
    param([string]$PackagePath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($PackagePath)
    try {
        if ($null -eq $archive.GetEntry("xl/vbaProject.bin")) { throw "The add-in package has no VBA project." }
        $ui = Get-XmlEntry $archive "customUI/customUI14.xml"
        $buttons = @($ui.SelectNodes("//*[local-name()='button']"))
        if ($buttons.Count -ne 3) { throw "The Ribbon definition must contain exactly three buttons." }
        $actions = @($buttons | ForEach-Object { $_.GetAttribute("onAction") })
        foreach ($required in @("RibbonStartEdit", "RibbonSubmit", "RibbonAbandon")) {
            if ($actions -notcontains $required) { throw "Ribbon callback is missing: $required" }
        }

        $contentTypes = Get-XmlEntry $archive "[Content_Types].xml"
        $override = $contentTypes.SelectSingleNode("/*[local-name()='Types']/*[local-name()='Override' and @PartName='/customUI/customUI14.xml']")
        if ($null -eq $override -or $override.GetAttribute("ContentType") -ne $script:RibbonContentType) {
            throw "The Ribbon content type is missing or invalid."
        }
        $relationships = Get-XmlEntry $archive "_rels/.rels"
        $relationship = $relationships.SelectSingleNode("/*[local-name()='Relationships']/*[local-name()='Relationship' and @Type='http://schemas.microsoft.com/office/2007/relationships/ui/extensibility']")
        if ($null -eq $relationship -or $relationship.GetAttribute("Target") -ne "customUI/customUI14.xml") {
            throw "The package Ribbon relationship is missing or invalid."
        }
        return [ordered]@{ verified = $true; ribbon_button_count = $buttons.Count; callbacks = $actions }
    }
    finally {
        $archive.Dispose()
    }
}

function Build-Workbook {
    param([string]$PackagePath, [string]$KnownPids)

    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class MddtAddinBuildNative {
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
    $project = $null
    $components = $null
    $ownedExcel = $false
    $excelPid = [uint32]0

    try {
        $excel = New-Object -ComObject Excel.Application
        [void][MddtAddinBuildNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
        if ($baseline -contains [int]$excelPid) { throw "Excel automation resolved to a pre-existing user process." }
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
        $files = @()
        foreach ($extension in @(".bas", ".cls", ".frm")) {
            $files += Get-ChildItem -LiteralPath $sourceRoot -Recurse -File -Filter "*$extension"
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

        $book.IsAddin = $true
        $book.SaveAs($PackagePath, 55)
        $componentCount = $components.Count
        $version = [string]$excel.Version
        $build = [string]$excel.Build
        $book.Close($false)
        Release-ComObject $book
        $book = $null
        return [ordered]@{
            status = "created"
            excel_version = $version
            excel_build = $build
            excel_pid = [int]$excelPid
            imported_component_count = $componentCount
            workbook_saved = $true
        }
    }
    finally {
        Release-ComObject $components
        Release-ComObject $project
        if ($null -ne $book) {
            try { $book.Close($false) } catch { }
            Release-ComObject $book
        }
        if ($null -ne $excel) {
            if ($ownedExcel) { try { $excel.Quit() } catch { } }
            Release-ComObject $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if ($Worker) {
    try {
        if (Test-Path -LiteralPath $OutputPath) { throw "Refusing to overwrite an existing add-in: $OutputPath" }
        $result = Build-Workbook -PackagePath $OutputPath -KnownPids $BaselineIds
        [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json -Depth 5))
        exit 0
    }
    catch {
        $failure = [ordered]@{ status = "failed"; error = $_.Exception.Message }
        try { [IO.File]::WriteAllText($ResultPath, ($failure | ConvertTo-Json -Depth 4)) } catch { }
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
}

$rootPath = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $rootPath "build\MDDTLockIssue.xlam" }
$absoluteOutputPath = [IO.Path]::GetFullPath($OutputPath)
if ($absoluteOutputPath.StartsWith("\\")) { throw "Add-in output must be local; UNC output is not allowed." }
$buildRoot = [IO.Path]::GetFullPath((Join-Path $rootPath "build"))
$buildRootPrefix = $buildRoot.TrimEnd([char[]]@('\', '/')) + [IO.Path]::DirectorySeparatorChar
if (-not $absoluteOutputPath.StartsWith($buildRootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Add-in output must remain inside the repository build directory."
}
if ([IO.Path]::GetExtension($absoluteOutputPath) -ne ".xlam") { throw "Add-in output must use the .xlam extension." }
if (Test-Path -LiteralPath $absoluteOutputPath) { throw "Refusing to overwrite an existing add-in: $absoluteOutputPath" }
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($absoluteOutputPath))

$baselineExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
$baselineIds = @($baselineExcel | ForEach-Object { $_.Id })
$baselineIdsText = $baselineIds -join ","
$tempRoot = [IO.Path]::GetTempPath()
$runId = [Guid]::NewGuid().ToString("N")
$resultPath = Join-Path $tempRoot "mddt-addin-build-$runId.json"
$stderrPath = Join-Path $tempRoot "mddt-addin-build-$runId.err"
$stdoutPath = Join-Path $tempRoot "mddt-addin-build-$runId.out"
$workerProcess = $null
$failure = $null
$workerProcessId = $null
$buildSucceeded = $false

try {
    $powershellCommand = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($null -eq $powershellCommand) { $powershellCommand = Get-Command pwsh.exe -ErrorAction Stop }
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Worker -OutputPath `"$absoluteOutputPath`" -ResultPath `"$resultPath`" -BaselineIds `"$baselineIdsText`""
    $workerProcess = Start-Process -FilePath $powershellCommand.Source -ArgumentList $arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    if ($null -eq $workerProcess) { throw "The build worker did not return a process handle." }
    $workerProcessId = $workerProcess.Id
    if (-not $workerProcess.WaitForExit($TimeoutSeconds * 1000)) {
        $failure = "Add-in build exceeded the $TimeoutSeconds second hard timeout (worker PID $workerProcessId)."
        try { Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue } catch { }
        [void]$workerProcess.WaitForExit(5000)
    }
    else {
        $workerProcess.Refresh()
        if ($workerProcess.ExitCode -ne 0) {
            $detail = if (Test-Path -LiteralPath $stderrPath) { (Get-Content -LiteralPath $stderrPath -Raw).Trim() } else { "" }
            if (-not $detail -and (Test-Path -LiteralPath $resultPath)) {
                $failedResult = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
                $detail = [string]$failedResult.error
            }
            $failure = "Add-in build failed. $detail"
        }
        elseif (-not (Test-Path -LiteralPath $resultPath)) {
            $failure = "Build worker exited without producing a result."
        }
    }

    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    do {
        $newExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $baselineIds -notcontains $_.Id })
        if ($newExcel.Count -eq 0) { break }
        Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $deadline)
    $newExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $baselineIds -notcontains $_.Id })
    if ($newExcel.Count -gt 0) {
        foreach ($processId in ($newExcel | ForEach-Object { $_.Id })) { Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue }
        if ($null -eq $failure) { $failure = "The isolated Excel process did not exit cleanly and was terminated by the watchdog." }
    }
    if ($null -ne $failure) { throw $failure }

    $workerResult = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
    if ($workerResult.status -ne "created") { throw "Build worker did not report a created package." }
    Add-RibbonPart -PackagePath $absoluteOutputPath -RibbonPath (Join-Path $rootPath "src\MdlUI\ribbon.xml")
    $packageResult = Test-AddinPackage -PackagePath $absoluteOutputPath
    $hash = (Get-FileHash -LiteralPath $absoluteOutputPath -Algorithm SHA256).Hash
    $finalResult = [ordered]@{
        status = "verified"
        output_path = $absoluteOutputPath
        excel_version = $workerResult.excel_version
        excel_build = $workerResult.excel_build
        imported_component_count = $workerResult.imported_component_count
        ribbon_button_count = $packageResult.ribbon_button_count
        callbacks = $packageResult.callbacks
        sha256 = $hash
    }
    $buildSucceeded = $true
    $finalResult | ConvertTo-Json -Depth 5
}
finally {
    if ($null -ne $workerProcess) { $workerProcess.Dispose() }
    if (-not $buildSucceeded -and (Test-Path -LiteralPath $absoluteOutputPath)) {
        Remove-Item -LiteralPath $absoluteOutputPath -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $resultPath, $stderrPath, $stdoutPath -Force -ErrorAction SilentlyContinue
}
