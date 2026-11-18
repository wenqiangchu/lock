# ============================================================
#  2-file-profile.ps1  -  xlsx structural profile (read-only)
#
#  Copies the file to %TEMP%, unzips it there, parses the XML.
#  Never opens Excel, so it creates no ~$ file and does not lock the share.
#  Writes only: a report on the Desktop (and a temp folder that is removed).
#
#  Usage:
#     .\2-file-profile.ps1 -File "\\SERVER\SHARE\book.xlsx"
#     .\2-file-profile.ps1 -File "..." -KeepTemp
#
#  NOTE: keep this file pure ASCII (PowerShell 5.1 reads BOM-less .ps1 as ANSI).
# ============================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$File,
    [string]$OutFile = "$([Environment]::GetFolderPath('Desktop'))\exs_fileprofile.txt",
    [switch]$KeepTemp,
    [int]$MaxSheetsDetail = 40
)

Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$ErrorActionPreference = 'Continue'
$report = New-Object System.Collections.Generic.List[string]
function W($s = '') { $report.Add([string]$s) }

$warn = New-Object System.Collections.Generic.List[string]
function WARN($s) { $warn.Add($s); W "  WARN: $s" }

# ---------- helpers ----------
function ConvertDim([string]$ref) {
    if (-not $ref) { return @(0, 0) }
    $parts = $ref.Split(':')
    $last = $parts[$parts.Count - 1]
    $m = [regex]::Match($last, '^([A-Za-z]+)(\d+)$')
    if (-not $m.Success) { return @(0, 0) }
    $letters = $m.Groups[1].Value.ToUpper()
    $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return @([int]$m.Groups[2].Value, $n)
}

# ---------- basic info ----------
W 'ExcelSync file structure profile'
W ("Generated: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
W "Source: $File"

if (-not (Test-Path $File)) {
    Write-Host "File not found: $File" -ForegroundColor Red
    return
}

$fi = Get-Item $File
W ''
W ('=' * 78); W '  A. File basics'; W ('=' * 78)
W ("Size          : {0:N0} bytes ({1:N2} MB)" -f $fi.Length, ($fi.Length / 1MB))
W ("LastWriteTime : {0}" -f $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))
W ("CreationTime  : {0}" -f $fi.CreationTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))
W ("Extension     : {0}" -f $fi.Extension)
W ("ReadOnly attr : {0}" -f $fi.IsReadOnly)
W ("Directory     : {0}" -f $fi.DirectoryName)
try { W ("SHA256        : {0}" -f (Get-FileHash $File -Algorithm SHA256 -ErrorAction Stop).Hash) }
catch { W "SHA256        : <failed> $($_.Exception.Message)" }

W ''
W 'Related files in the same directory:'
try {
    Get-ChildItem $fi.DirectoryName -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like '~$*' -or $_.Extension -in '.xlsx', '.xlsm', '.xlsb', '.xls' } |
        ForEach-Object {
            $tag = if ($_.Name -like '~$*') { '[Excel lock file]' } else { '' }
            W ("  {0,-40} {1,12:N0}  {2}  {3}" -f $_.Name, $_.Length, $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'), $tag)
        }
    $owner = Get-ChildItem $fi.DirectoryName -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "~`$$($fi.Name)" }
    if ($owner) {
        W ''
        W "  Someone (or a stale session) currently has this file open."
        W "  Lock file: $($owner.Name)  modified $($owner.LastWriteTime)"
        WARN "a ~`$ lock file exists: $($owner.Name)"
    }
} catch { W "  <enumeration failed> $($_.Exception.Message)" }

# ---------- unpack ----------
$tmpRoot = Join-Path $env:TEMP ('exs_prof_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$tmpFile = Join-Path $tmpRoot $fi.Name
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
W ''
W "Local temp copy: $tmpFile"
try {
    Copy-Item $File $tmpFile -Force -ErrorAction Stop
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tmpFile)
} catch {
    W "Copy/unpack failed: $($_.Exception.Message)"
    W 'If the share is unreadable, copy the file locally first and point -File at the copy.'
    $report -join "`r`n" | Set-Content $OutFile -Encoding UTF8
    Write-Host "Report saved: $OutFile" -ForegroundColor Yellow
    return
}

# ---------- B. parts ----------
W ''
W ('=' * 78); W '  B. Package parts (uncompressed size)'; W ('=' * 78)
$entries = $zip.Entries | Sort-Object Length -Descending
W ("{0,10} {1,10}  {2}" -f 'raw MB', 'zip MB', 'part')
foreach ($e in ($entries | Select-Object -First 20)) {
    W ("{0,10:N2} {1,10:N2}  {2}" -f ($e.Length / 1MB), ($e.CompressedLength / 1MB), $e.FullName)
}
W ''
$totalUncompressed = ($entries | Measure-Object Length -Sum).Sum
W ("Parts: {0}   total uncompressed: {1:N2} MB   compression ratio: {2:P1}" -f `
        $entries.Count, ($totalUncompressed / 1MB), (1 - ($fi.Length / [math]::Max($totalUncompressed, 1))))

function SumPart([string]$pattern) {
    $m = $entries | Where-Object { $_.FullName -like $pattern }
    if ($m) {
        [pscustomobject]@{
            Count = ($m | Measure-Object).Count
            Bytes = ($m | Measure-Object Length -Sum).Sum
            Names = ($m | Select-Object -First 5 -ExpandProperty FullName) -join ', '
        }
    } else { [pscustomobject]@{ Count = 0; Bytes = 0; Names = '' } }
}

W ''
W 'Key parts:'
foreach ($p in @(
    @{ N = 'styles.xml'; P = 'xl/styles.xml' },
    @{ N = 'sharedStrings'; P = 'xl/sharedStrings.xml' },
    @{ N = 'worksheets'; P = 'xl/worksheets/*.xml' },
    @{ N = 'media (images)'; P = 'xl/media/*' },
    @{ N = 'charts'; P = 'xl/charts/*' },
    @{ N = 'pivotCache'; P = 'xl/pivotCache/*' },
    @{ N = 'pivotTables'; P = 'xl/pivotTables/*' },
    @{ N = 'drawings'; P = 'xl/drawings/*' },
    @{ N = 'externalLinks'; P = 'xl/externalLinks/*' },
    @{ N = 'connections'; P = 'xl/connections.xml' },
    @{ N = 'vbaProject'; P = 'xl/vbaProject.bin' },
    @{ N = 'tables (ListObject)'; P = 'xl/tables/*' },
    @{ N = 'comments'; P = 'xl/comments*.xml' },
    @{ N = 'customXml'; P = 'customXml/*' }
)) {
    $r = SumPart $p.P
    if ($r.Count -gt 0) {
        W ("  {0,-20} count={1,-4} size={2,8:N2} MB   {3}" -f $p.N, $r.Count, ($r.Bytes / 1MB), $r.Names)
    } else {
        W ("  {0,-20} none" -f $p.N)
    }
}

if ((SumPart 'xl/vbaProject.bin').Count -gt 0 -or $fi.Extension -in '.xlsm', '.xlsb') {
    WARN 'file contains VBA or is a macro-enabled format; the design assumed plain .xlsx'
}
if ((SumPart 'xl/charts/*').Count -gt 0) { WARN 'contains charts: ClearFormats / whole-range overwrite would break them' }
if ((SumPart 'xl/pivotTables/*').Count -gt 0) { WARN 'contains pivot tables: overwrite the data body only' }
if ((SumPart 'xl/tables/*').Count -gt 0) { WARN 'contains ListObject (Excel Table): do NOT Clear the whole range' }

# ---------- read key XML ----------
function ReadEntryText($entry) {
    if (-not $entry) { return '' }
    $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $t = $sr.ReadToEnd(); $sr.Dispose(); return $t
}

$wbXml  = ReadEntryText ($entries | Where-Object { $_.FullName -eq 'xl/workbook.xml' } | Select-Object -First 1)
$wbRels = ReadEntryText ($entries | Where-Object { $_.FullName -eq 'xl/_rels/workbook.xml.rels' } | Select-Object -First 1)
$styles = ReadEntryText ($entries | Where-Object { $_.FullName -eq 'xl/styles.xml' } | Select-Object -First 1)
$ssXml  = ReadEntryText ($entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' } | Select-Object -First 1)

W ''
W ('=' * 78); W '  C. Workbook-level properties'; W ('=' * 78)
foreach ($pat in @('calcPr', 'workbookPr', 'fileSharing', 'workbookProtection', 'definedNames')) {
    $m = [regex]::Match($wbXml, "<$pat\b[^>]*/?>")
    if ($m.Success) { W ("  {0,-20} {1}" -f $pat, $m.Value) } else { W ("  {0,-20} <absent>" -f $pat) }
}
if ($wbXml -match '<fileSharing\b') {
    WARN 'legacy shared workbook (fileSharing) present: convert to a normal workbook first'
}
if ($wbXml -match 'workbookProtection[^>]*workbookPassword') {
    WARN 'workbook structure protection present (may block adding/removing sheets)'
}

$dnNodes = [regex]::Matches($wbXml, '<definedName\b([^>]*)>(.*?)</definedName>', 'Singleline')
if ($dnNodes.Count -gt 0) {
    W ("  defined names: {0}" -f $dnNodes.Count)
    foreach ($dn in ($dnNodes | Select-Object -First 15)) {
        $nm = [regex]::Match($dn.Groups[1].Value, 'name="([^"]*)"').Groups[1].Value
        $ref = $dn.Groups[2].Value -replace '<[^>]+>', ''
        $hid = if ($dn.Groups[1].Value -match 'hidden="1"') { ' [hidden]' } else { '' }
        W ("    - {0}{1} = {2}" -f $nm, $hid, $ref)
    }
}

$ssCount = 0; $ssUnique = 0
if ($ssXml) {
    $m = [regex]::Match($ssXml, 'uniqueCount="(\d+)"'); if ($m.Success) { $ssUnique = [int]$m.Groups[1].Value }
    $m2 = [regex]::Match($ssXml, 'count="(\d+)"'); if ($m2.Success) { $ssCount = [int]$m2.Groups[1].Value }
    W ("  sharedStrings: count={0} uniqueCount={1}  ({2:N2} MB)" -f $ssCount, $ssUnique,
        ((($entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' }).Length) / 1MB))
}

W ''
W ('=' * 78); W '  D. Style fragmentation'; W ('=' * 78)
$cellXfsCount = 0
$xfNumFmtIds = @()
if ($styles) {
    $m = [regex]::Match($styles, '<cellXfs\b[^>]*count="(\d+)"')
    if ($m.Success) { $cellXfsCount = [int]$m.Groups[1].Value }
    W ("  cellXfs count (distinct cell format combinations): {0}" -f $cellXfsCount)
    W ("  styles.xml size: {0:N2} MB" -f ((($entries | Where-Object { $_.FullName -eq 'xl/styles.xml' }).Length) / 1MB))
    if ($cellXfsCount -gt 200) { WARN "cellXfs=$cellXfsCount is high (style fragmentation), a common cause of file bloat" }
    if ($cellXfsCount -gt 1000) { WARN "cellXfs=$cellXfsCount is very high, normalise to named styles" }

    $xfsBlock = [regex]::Match($styles, '<cellXfs\b[^>]*>(.*?)</cellXfs>', 'Singleline')
    if ($xfsBlock.Success) {
        foreach ($x in [regex]::Matches($xfsBlock.Groups[1].Value, '<xf\b[^>]*>')) {
            $n = [regex]::Match($x.Value, 'numFmtId="(\d+)"')
            $xfNumFmtIds += $(if ($n.Success) { [int]$n.Groups[1].Value } else { 0 })
        }
        W ("  parsed cellXfs entries: {0}" -f $xfNumFmtIds.Count)
    }
    W ("  custom numFmt count: {0}" -f ([regex]::Matches($styles, '<numFmt\b[^>]*/>').Count))
} else { W '  <no styles.xml>' }

# ---------- sheet list ----------
W ''
W ('=' * 78); W '  E. Worksheet list'; W ('=' * 78)
$relMap = @{}
foreach ($m in [regex]::Matches($wbRels, '<Relationship\b[^>]*/>')) {
    $id = [regex]::Match($m.Value, 'Id="([^"]+)"').Groups[1].Value
    $tg = [regex]::Match($m.Value, 'Target="([^"]+)"').Groups[1].Value
    if ($id) { $relMap[$id] = $tg }
}
$sheets = @()
foreach ($n in [regex]::Matches($wbXml, '<sheet\b[^>]*/>')) {
    $name = [regex]::Match($n.Value, 'name="([^"]*)"').Groups[1].Value
    $rid  = [regex]::Match($n.Value, 'r:id="([^"]*)"').Groups[1].Value
    $state = [regex]::Match($n.Value, 'state="([^"]*)"').Groups[1].Value
    $tgt = $relMap[$rid]
    if ($tgt -and $tgt -notlike 'xl/*') { $tgt = 'xl/' + $tgt.TrimStart('/') }
    $sheets += [pscustomobject]@{ Name = $name; State = $state; Target = $tgt; Rid = $rid }
}
W ("Sheets: {0}   (visible: {1}   hidden/veryHidden: {2})" -f $sheets.Count,
    ($sheets | Where-Object { -not $_.State }).Count,
    ($sheets | Where-Object { $_.State }).Count)

# ---------- sharedStrings array (loaded once) ----------
$script:ssArr = @()
if ($ssXml) {
    foreach ($si in [regex]::Matches($ssXml, '<si>(.*?)</si>', 'Singleline')) {
        $txt = $si.Groups[1].Value
        $txt = $txt -replace '<phoneticPr[^>]*/>', ''
        $txt = $txt -replace '<rPh\b.*?</rPh>', ''
        $txt = $txt -replace '<[^>]+>', ''
        $txt = $txt -replace '&lt;', '<' -replace '&gt;', '>' -replace '&amp;', '&' -replace '&quot;', '"' -replace '&apos;', "'"
        $script:ssArr += $txt
    }
}

# ---------- per-sheet analysis ----------
function AnalyzeSheet($sh) {
    $entry = $entries | Where-Object { $_.FullName -eq $sh.Target } | Select-Object -First 1
    if (-not $entry) { return $null }

    $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $xml = $sr.ReadToEnd(); $sr.Dispose()

    $dim = [regex]::Match($xml, '<dimension\b[^>]*ref="([^"]+)"').Groups[1].Value
    $dc = ConvertDim $dim

    $rowCount  = [regex]::Matches($xml, '<row\b').Count
    $cellCount = [regex]::Matches($xml, '<c\b').Count
    $isCount   = [regex]::Matches($xml, '<is\b').Count
    $fCount    = [regex]::Matches($xml, '<f\b').Count
    $hlCount   = [regex]::Matches($xml, '<hyperlink\b').Count
    $hlInternal = [regex]::Matches($xml, '<hyperlink\b[^>]*location=').Count
    $mergeCount = [regex]::Matches($xml, '<mergeCell\b').Count
    $dvCount   = [regex]::Matches($xml, '<dataValidation\b').Count
    $cfCount   = [regex]::Matches($xml, '<conditionalFormatting\b').Count
    $tblMatch  = [regex]::Match($xml, '<tableParts\b[^>]*count="(\d+)"')
    $tblN = if ($tblMatch.Success) { [int]$tblMatch.Groups[1].Value } else { [regex]::Matches($xml, '<tablePart\b').Count }
    $rowHt = [regex]::Matches($xml, '<row\b[^>]*ht=').Count
    $hasProt = $xml -match '<sheetProtection\b'
    $protAttrs = if ($hasProt) { [regex]::Match($xml, '<sheetProtection\b[^>]*/>').Value } else { '' }
    $hasCols = $xml -match '<cols\b'
    $arrF = [regex]::Matches($xml, '<f\b[^>]*t="array"').Count
    $sharedF = [regex]::Matches($xml, '<f\b[^>]*t="shared"').Count

    $volatile = @()
    foreach ($fn in 'TODAY', 'NOW', 'RAND', 'RANDBETWEEN', 'INDIRECT', 'OFFSET', 'INFO', 'CELL') {
        if ($xml -match ">\s*=\s*[^<]*\b$fn\s*\(") { $volatile += $fn }
    }
    $hasHyperlinkFn = $xml -match '=\s*HYPERLINK\s*\('

    # ---- column A (candidate key column) ----
    $keyStat = $null
    $aCells = [regex]::Matches($xml, '<c\b[^>]*r="A\d+"[^>]*/>|<c\b[^>]*r="A\d+"[^>]*>.*?</c>', 'Singleline')
    if ($aCells.Count -gt 0) {
        $vals = New-Object System.Collections.Generic.List[string]
        $types = @{ num = 0; str = 0; inlineStr = 0; bool = 0; err = 0; empty = 0 }
        $fmtIds = @{}
        foreach ($c in $aCells) {
            $tag = [regex]::Match($c.Value, '^<c\b[^>]*>').Value
            $t = [regex]::Match($tag, '\bt="([^"]+)"').Groups[1].Value
            $sIdx = [regex]::Match($tag, '\bs="(\d+)"').Groups[1].Value
            if ($sIdx -ne '' -and $xfNumFmtIds.Count -gt [int]$sIdx) {
                $fid = $xfNumFmtIds[[int]$sIdx]
                if ($fmtIds.ContainsKey($fid)) { $fmtIds[$fid]++ } else { $fmtIds[$fid] = 1 }
            }
            $v = [regex]::Match($c.Value, '<v>(.*?)</v>', 'Singleline').Groups[1].Value
            if ($t -eq 's' -and $v -match '^\d+$') {
                $vi = [int]$v
                if ($vi -lt $script:ssArr.Count) { $vals.Add($script:ssArr[$vi]) } else { $vals.Add('<BAD_SS_IDX>') }
                $types.str++
            } elseif ($t -eq 'inlineStr') {
                $it = [regex]::Match($c.Value, '<is>.*?<t[^>]*>(.*?)</t>', 'Singleline').Groups[1].Value
                $vals.Add($it); $types.inlineStr++
            } elseif ($v -ne '') {
                $vals.Add($v)
                if ($t -eq 'b') { $types.bool++ } elseif ($t -eq 'e') { $types.err++ } else { $types.num++ }
            } else {
                $types.empty++
            }
        }
        $nonEmpty = @($vals | Where-Object { $_ -ne '' -and $_.Trim() -ne '' })
        # first row is usually the header; exclude it from uniqueness checks
        $dataRows = if ($nonEmpty.Count -gt 1) { @($nonEmpty | Select-Object -Skip 1) } else { @() }
        # NOTE: must wrap in @() -- a single GroupInfo has its own .Count that would shadow array length
        $dupGroups = @($dataRows | Group-Object | Where-Object { $_.Count -gt 1 })
        $leadingZero = @($dataRows | Where-Object { $_ -match '^0\d+$' })
        $textFmtCount = 0
        if ($fmtIds.ContainsKey(49)) { $textFmtCount = $fmtIds[49] }

        $keyStat = [pscustomobject]@{
            ACells = $aCells.Count
            NonEmpty = $nonEmpty.Count
            DataRows = $dataRows.Count
            HeaderValue = $(if ($nonEmpty.Count -gt 0) { $nonEmpty[0] } else { '' })
            DupGroups = $dupGroups.Count
            DupSamples = (($dupGroups | Select-Object -First 10 | ForEach-Object { "$($_.Name) x$($_.Count)" }) -join ', ')
            LeadingZero = $leadingZero.Count
            LZSamples = (($leadingZero | Select-Object -First 8) -join ', ')
            TypeNum = $types.num; TypeStr = $types.str; TypeInline = $types.inlineStr
            TypeBool = $types.bool; TypeErr = $types.err; TypeEmpty = $types.empty
            Fmt49 = $textFmtCount
            FmtIds = (($fmtIds.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5 | ForEach-Object { "numFmtId=$($_.Key) x$($_.Value)" }) -join ' ')
        }
    }

    return [pscustomobject]@{
        Name = $sh.Name; State = $sh.State; Target = $sh.Target
        Dim = $dim; DimRows = $dc[0]; DimCols = $dc[1]
        Rows = $rowCount; Cells = $cellCount; InlineStr = $isCount
        Formulas = $fCount; ArrayF = $arrF; SharedF = $sharedF
        Hyperlinks = $hlCount; HyperInternal = $hlInternal; HyperlinkFn = $hasHyperlinkFn
        Merges = $mergeCount; DataValidations = $dvCount; CondFormats = $cfCount
        Tables = $tblN; RowHeights = $rowHt; HasCols = $hasCols
        HasProtection = $hasProt; ProtAttrs = $protAttrs
        Volatile = ($volatile -join ',')
        SizeMB = [math]::Round($entry.Length / 1MB, 3)
        KeyStat = $keyStat
    }
}

W ''
W ('=' * 78); W '  F. Per-sheet detail'; W ('=' * 78)
$results = @()
foreach ($sh in $sheets) {
    $r = AnalyzeSheet $sh
    if ($r) { $results += $r }
}

W ''
W ("{0,-28} {1,7} {2,7} {3,16} {4,8} {5,7}" -f 'Sheet', 'rows', 'cells', 'dimension', 'hyperlnk', 'MB')
W ('-' * 82)
foreach ($r in $results) {
    $st = if ($r.State) { " ($($r.State))" } else { '' }
    W ("{0,-28} {1,7} {2,7} {3,16} {4,8} {5,7}" -f ($r.Name + $st), $r.Rows, $r.Cells, $r.Dim, $r.Hyperlinks, $r.SizeMB)
}
$sumRows  = ($results | Measure-Object Rows -Sum).Sum
$sumCells = ($results | Measure-Object Cells -Sum).Sum
$sumHL    = ($results | Measure-Object Hyperlinks -Sum).Sum
$sumMB    = ($results | Measure-Object SizeMB -Sum).Sum
W ('-' * 82)
W ("{0,-28} {1,7} {2,7} {3,16} {4,8} {5,7}" -f 'TOTAL', $sumRows, $sumCells, '', $sumHL, [math]::Round($sumMB, 3))

W ''
W 'Per-sheet attributes (only non-empty attributes are listed):'
foreach ($r in $results) {
    $flags = @()
    if ($r.Formulas) { $flags += "formulas=$($r.Formulas)" + $(if ($r.ArrayF) { "(array $($r.ArrayF))" } else { '' }) + $(if ($r.SharedF) { "(shared $($r.SharedF))" } else { '' }) }
    if ($r.Hyperlinks) { $flags += "hyperlinks=$($r.Hyperlinks)" + $(if ($r.HyperInternal) { "(internal $($r.HyperInternal))" } else { '' }) + $(if ($r.HyperlinkFn) { "+HYPERLINK() formula" } else { '' }) }
    if ($r.Merges) { $flags += "merged=$($r.Merges)" }
    if ($r.DataValidations) { $flags += "dataValidation=$($r.DataValidations)" }
    if ($r.CondFormats) { $flags += "condFormat=$($r.CondFormats)" }
    if ($r.Tables) { $flags += "ListObject=$($r.Tables)" }
    if ($r.RowHeights) { $flags += "customRowHeight=$($r.RowHeights)" }
    if ($r.HasCols) { $flags += "customColWidth" }
    if ($r.HasProtection) { $flags += "PROTECTED" }
    if ($r.Volatile) { $flags += "volatile=$($r.Volatile)" }
    if ($r.InlineStr) { $flags += "inlineStr=$($r.InlineStr)" }
    W ("  {0,-26} {1}" -f $r.Name, ($flags -join '  '))

    if ($r.HasProtection) {
        W "      -> $($r.ProtAttrs)"
        WARN "sheet [$($r.Name)] is protected: write-back needs Unprotect (password) or a design change"
    }
    if ($r.Tables) { WARN "sheet [$($r.Name)] has ListObject x$($r.Tables): must not Clear the whole range" }
    if ($r.Hyperlinks) {
        W ("      -> hyperlinks: {0}" -f $(if ($r.HyperInternal -gt 0) { 'contains internal jumps (TOC), Copy must preserve them' } else { 'external links only' }))
    }
}

W ''
W 'dimension vs actual cell count (bloat indicator):'
foreach ($r in $results) {
    if ($r.DimRows -le 0) { continue }
    $dimArea = $r.DimRows * $r.DimCols
    $density = if ($dimArea -gt 0) { $r.Cells / $dimArea } else { 0 }
    $flag = if ($density -lt 0.5 -and $dimArea -gt 20000) { '  <- sparse (suspected format garbage)' } else { '' }
    W ("  {0,-26} dim={1,-14} area={2,8:N0}  cells={3,7:N0}  density={4,5:P0}{5}" -f `
            $r.Name, $r.Dim, $dimArea, $r.Cells, $density, $flag)
    if ($density -lt 0.3 -and $dimArea -gt 100000) {
        WARN "sheet [$($r.Name)] dimension=$($r.Dim) but only $($r.Cells) cells (density $('{0:P0}' -f $density)) -> typical format garbage, cleaning would slim the file"
    }
}

# ---------- key column ----------
W ''
W ('=' * 78); W '  G. Column A (key column) usability'; W ('=' * 78)
foreach ($r in $results) {
    if (-not $r.KeyStat) { W "  $($r.Name): <no column A cells>"; continue }
    $k = $r.KeyStat
    W ''
    W "  [$($r.Name)]"
    W ("    column A cells={0}  non-empty={1}  (first row treated as header: '{2}', data rows={3})" -f $k.ACells, $k.NonEmpty, $k.HeaderValue, $k.DataRows)
    W ("    types: num={0} sharedStr={1} inlineStr={2} bool={3} err={4} empty={5}" -f $k.TypeNum, $k.TypeStr, $k.TypeInline, $k.TypeBool, $k.TypeErr, $k.TypeEmpty)
    W ("    numFmtId histogram: {0}" -f $k.FmtIds)
    W ("    text-format (numFmtId=49) cells: {0}" -f $k.Fmt49)
    W ("    duplicate key groups={0}  {1}" -f $k.DupGroups, $k.DupSamples)
    W ("    leading-zero keys={0}  {1}" -f $k.LeadingZero, $k.LZSamples)

    if ($k.DupGroups -gt 0) {
        if ($r.Hyperlinks -gt 0 -and $k.ACells -le 10) {
            W '    note: column A looks like a TOC description column, not a unique key (has hyperlinks, few rows); duplicates are expected'
        } elseif ($k.NonEmpty -le 3) {
            W '    note: too little data to judge key uniqueness'
        } else {
            WARN "sheet [$($r.Name)] column A has $($k.DupGroups) duplicate key groups -> key-based guard cannot align rows, fix data first"
        }
    }
    if ($k.LeadingZero -gt 0 -and $k.Fmt49 -eq 0 -and $k.DupGroups -eq 0) {
        WARN "sheet [$($r.Name)] has $($k.LeadingZero) leading-zero keys but column A is not text-formatted -> Excel may convert them to numbers on edit"
    }
    if ($k.TypeNum -gt 0 -and $k.TypeStr -gt 0) {
        WARN "sheet [$($r.Name)] column A mixes numeric and text values -> normalisation must be careful"
    }
    if ($k.ACells -gt 10 -and $k.TypeEmpty -gt $k.ACells * 0.2) {
        WARN "sheet [$($r.Name)] column A has many empty cells ($($k.TypeEmpty)/$($k.ACells)) -> possible merged cells or sub-heading rows"
    }
}

# ---------- summary ----------
W ''
W ('=' * 78); W '  H. Key findings'; W ('=' * 78)
$nonEmptySheets = $results | Where-Object { $_.Cells -gt 0 }
W ("Sheets with data: {0} / {1}" -f $nonEmptySheets.Count, $results.Count)
W ("Total rows: $sumRows   total cells: $sumCells   total raw size: $([math]::Round($sumMB, 2)) MB")
W ''
$est = 0
foreach ($r in $nonEmptySheets) { $est += ($r.Cells * 45) }
W ("Estimated pure XML volume at ~45 bytes/cell: {0:N2} MB (about {1:N2} MB after compression)" -f ($est / 1MB), ($est * 0.15 / 1MB))
W ("Actual file size: {0:N2} MB" -f ($fi.Length / 1MB))
if ($fi.Length -gt 1MB -and $fi.Length -gt $est * 2.5) {
    W ''
    W '  WARN: actual size far exceeds the pure-data estimate -> bloat comes from styles/images/empty formats; slimming has good payoff'
    $warn.Add("file bloat: actual $([math]::Round($fi.Length / 1MB, 2)) MB vs data estimate $([math]::Round($est / 1MB, 2)) MB")
} elseif ($fi.Length -le 1MB) {
    W '  (file is under 1 MB, bloat check skipped)'
} else {
    W '  OK: size roughly matches the data volume'
}

W ''
if ($warn.Count -eq 0) {
    W '  No blocking issues found.'
} else {
    W ("  {0} item(s) need attention:" -f $warn.Count)
    $n = 0
    foreach ($x in $warn) { $n++; W "   $n. $x" }
}

W ''
W ('=' * 78); W '  I. Commit-window estimate'; W ('=' * 78)
$avgCells = if ($nonEmptySheets.Count -gt 0) { [int](($nonEmptySheets | Measure-Object Cells -Average).Average) } else { 0 }
W ("Average cells per sheet: $avgCells")
W 'Assuming 1-3 declared sheets (calibrate with 3-measure.ps1):'
W ("  download whole file : {0:N1} s (at 10 MB/s)" -f ($fi.Length / 1MB / 10))
W ("  region Copy         : {0:N2} s (measured: 25000 cells took 62 ms)" -f ([math]::Max(0.05, $avgCells * 0.0025)))
W '  Save                : see 3-measure.ps1'
W '  -> conservative estimate: server held 8-20 s'

$zip.Dispose()
if (-not $KeepTemp) {
    try { Remove-Item $tmpRoot -Recurse -Force -ErrorAction Stop; W ''; W "Temp folder removed: $tmpRoot" }
    catch { W "Could not remove temp folder, delete manually: $tmpRoot" }
} else {
    W ''
    W "Temp folder kept for manual inspection: $tmpRoot"
}

$report -join "`r`n" | Set-Content -Path $OutFile -Encoding UTF8
Write-Host ''
Write-Host "Report saved: $OutFile" -ForegroundColor Green
if ($warn.Count -gt 0) {
    Write-Host "$($warn.Count) item(s) need attention, please send the report back." -ForegroundColor Yellow
} else {
    Write-Host 'No blocking issues found.' -ForegroundColor Green
}