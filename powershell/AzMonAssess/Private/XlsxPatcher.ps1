#requires -Version 7.0
# Surgical in-place editor for an already-generated report.xlsx — used by
# Invoke-AzMonVerification to flip review-status cells after re-checking
# findings against the customer's current environment, without touching
# any other cell, style, or sheet. Complements XlsxBuilder.ps1 (which only
# ever builds a workbook from scratch); this module never regenerates a
# sheet, it only rewrites the text of specific already-existing cells.
#
# Safe by construction because New-AzMonExcelReport.ps1 writes every
# string cell as an inline string (t="inlineStr", see XlsxBuilder.ps1) —
# there is no shared-strings table, so changing one cell's <t> text can
# never change the text shown in any other cell.

function Read-AzMonZipEntryText {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Entry)
    $stream = $Entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        return $reader.ReadToEnd()
    } finally {
        $stream.Dispose()
    }
}

function Resolve-AzMonXlsxSheetPartPath {
    <#
    .SYNOPSIS
        Resolves a worksheet's display name (e.g. '4.ImpactedResourcesAnalysis')
        to its package part path (e.g. 'xl/worksheets/sheet3.xml') by reading
        xl/workbook.xml + xl/_rels/workbook.xml.rels, rather than assuming a
        fixed sheet order.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Zip, [Parameter(Mandatory)] [string] $SheetName)

    $r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $wbEntry = $Zip.GetEntry('xl/workbook.xml')
    if (-not $wbEntry) { throw 'xl/workbook.xml not found - not a valid .xlsx package.' }
    [xml]$wbXml = Read-AzMonZipEntryText -Entry $wbEntry
    $sheetNode = $wbXml.SelectSingleNode("//*[local-name()='sheet' and @name='$SheetName']")
    if (-not $sheetNode) { throw "Sheet '$SheetName' not found in workbook.xml." }
    $rId = $sheetNode.GetAttribute('id', $r)

    $relsEntry = $Zip.GetEntry('xl/_rels/workbook.xml.rels')
    if (-not $relsEntry) { throw 'xl/_rels/workbook.xml.rels not found.' }
    [xml]$relsXml = Read-AzMonZipEntryText -Entry $relsEntry
    $relNode = $relsXml.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rId']")
    if (-not $relNode) { throw "Relationship '$rId' for sheet '$SheetName' not found." }
    return "xl/$($relNode.Target)"
}

function Get-AzMonXlsxHeaderColumnMap {
    <#
    .SYNOPSIS
        Maps column letter -> header text from a worksheet XML DOM's first
        row, so callers locate columns by name instead of a hardcoded index.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $SheetXml)

    $headerRow = $SheetXml.SelectSingleNode("//*[local-name()='sheetData']/*[local-name()='row'][1]")
    $map = @{}
    if (-not $headerRow) { return $map }
    foreach ($cell in $headerRow.SelectNodes("*[local-name()='c']")) {
        $colLetter = ($cell.GetAttribute('r') -replace '\d+$', '')
        $t = $cell.SelectSingleNode("*[local-name()='is']/*[local-name()='t']")
        if ($t) { $map[$colLetter] = $t.InnerText }
    }
    return $map
}

function Get-AzMonExcelReviewStatus {
    <#
    .SYNOPSIS
        Read-only pass: returns the current "REQUIRED ACTIONS / REVIEW
        STATUS" text for every (checkName, id) row in the given sheet, keyed
        as "findingId|resourceid-lowercased". Used to make verification
        idempotent (skip no-op rewrites, detect regressions).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $SheetName,
        [string] $KeyColumnHeader = 'checkName',
        [string] $ResourceColumnHeader = 'id',
        [string] $StatusColumnHeader = 'REQUIRED ACTIONS / REVIEW STATUS'
    )
    if (-not (Test-Path -LiteralPath $Path)) { throw "Report not found: $Path" }
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $result = @{}
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $partPath = Resolve-AzMonXlsxSheetPartPath -Zip $zip -SheetName $SheetName
            $sheetEntry = $zip.GetEntry($partPath)
            if (-not $sheetEntry) { throw "Worksheet part '$partPath' not found." }
            [xml]$sheetXml = Read-AzMonZipEntryText -Entry $sheetEntry
            $colMap = Get-AzMonXlsxHeaderColumnMap -SheetXml $sheetXml
            $keyCol = ($colMap.Keys | Where-Object { $colMap[$_] -eq $KeyColumnHeader } | Select-Object -First 1)
            $resourceCol = ($colMap.Keys | Where-Object { $colMap[$_] -eq $ResourceColumnHeader } | Select-Object -First 1)
            $statusCol = ($colMap.Keys | Where-Object { $colMap[$_] -eq $StatusColumnHeader } | Select-Object -First 1)
            if (-not $keyCol -or -not $statusCol) { throw "Required columns not found in header row of '$SheetName'." }

            $rows = $sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")
            for ($i = 1; $i -lt $rows.Count; $i++) {
                $cells = @{}
                foreach ($cell in $rows[$i].SelectNodes("*[local-name()='c']")) {
                    $cells[($cell.GetAttribute('r') -replace '\d+$', '')] = $cell
                }
                if (-not $cells.ContainsKey($keyCol)) { continue }
                $findingIdNode = $cells[$keyCol].SelectSingleNode("*[local-name()='is']/*[local-name()='t']")
                if (-not $findingIdNode -or -not $findingIdNode.InnerText) { continue }
                $findingId = $findingIdNode.InnerText
                $resourceId = ''
                if ($resourceCol -and $cells.ContainsKey($resourceCol)) {
                    $rNode = $cells[$resourceCol].SelectSingleNode("*[local-name()='is']/*[local-name()='t']")
                    if ($rNode) { $resourceId = $rNode.InnerText }
                }
                $statusText = ''
                if ($cells.ContainsKey($statusCol)) {
                    $sNode = $cells[$statusCol].SelectSingleNode("*[local-name()='is']/*[local-name()='t']")
                    if ($sNode) { $statusText = $sNode.InnerText }
                }
                $result["$findingId|$(([string]$resourceId).ToLowerInvariant())"] = $statusText
            }
        } finally {
            $zip.Dispose()
        }
    } finally {
        $fs.Dispose()
    }
    return $result
}

function Update-AzMonExcelReviewStatus {
    <#
    .SYNOPSIS
        Rewrites the "REQUIRED ACTIONS / REVIEW STATUS" cell text for the
        given (FindingId, ResourceId) rows in an existing report.xlsx,
        in place. Only the target cells' text changes - every other cell,
        style, sheet, and package part is byte-for-byte untouched, so all
        existing workbook formatting is preserved.
    .PARAMETER StatusUpdate
        Array of @{ FindingId; ResourceId; NewStatus } hashtables.
    .OUTPUTS
        Array of @{ FindingId; ResourceId; PreviousStatus; NewStatus }
        actually applied (rows already at NewStatus are skipped as no-ops).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $SheetName,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $StatusUpdate,
        [string] $KeyColumnHeader = 'checkName',
        [string] $ResourceColumnHeader = 'id',
        [string] $StatusColumnHeader = 'REQUIRED ACTIONS / REVIEW STATUS'
    )
    $applied = [System.Collections.Generic.List[hashtable]]::new()
    if (-not $StatusUpdate -or $StatusUpdate.Count -eq 0) { return $applied.ToArray() }
    if (-not (Test-Path -LiteralPath $Path)) { throw "Report not found: $Path" }
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $updateByKey = @{}
    foreach ($u in $StatusUpdate) {
        $updateByKey["$($u.FindingId)|$(([string]$u.ResourceId).ToLowerInvariant())"] = $u
    }

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $partPath = Resolve-AzMonXlsxSheetPartPath -Zip $zip -SheetName $SheetName
            $sheetEntry = $zip.GetEntry($partPath)
            if (-not $sheetEntry) { throw "Worksheet part '$partPath' not found." }
            [xml]$sheetXml = Read-AzMonZipEntryText -Entry $sheetEntry
            $colMap = Get-AzMonXlsxHeaderColumnMap -SheetXml $sheetXml
            $keyCol = ($colMap.Keys | Where-Object { $colMap[$_] -eq $KeyColumnHeader } | Select-Object -First 1)
            $resourceCol = ($colMap.Keys | Where-Object { $colMap[$_] -eq $ResourceColumnHeader } | Select-Object -First 1)
            $statusCol = ($colMap.Keys | Where-Object { $colMap[$_] -eq $StatusColumnHeader } | Select-Object -First 1)
            if (-not $keyCol -or -not $statusCol) { throw "Required columns not found in header row of '$SheetName'." }

            $rows = $sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")
            for ($i = 1; $i -lt $rows.Count; $i++) {
                $cells = @{}
                foreach ($cell in $rows[$i].SelectNodes("*[local-name()='c']")) {
                    $cells[($cell.GetAttribute('r') -replace '\d+$', '')] = $cell
                }
                if (-not $cells.ContainsKey($keyCol)) { continue }
                $findingIdNode = $cells[$keyCol].SelectSingleNode("*[local-name()='is']/*[local-name()='t']")
                if (-not $findingIdNode -or -not $findingIdNode.InnerText) { continue }
                $findingId = $findingIdNode.InnerText
                $resourceId = ''
                if ($resourceCol -and $cells.ContainsKey($resourceCol)) {
                    $rNode = $cells[$resourceCol].SelectSingleNode("*[local-name()='is']/*[local-name()='t']")
                    if ($rNode) { $resourceId = $rNode.InnerText }
                }
                $key = "$findingId|$(([string]$resourceId).ToLowerInvariant())"
                if (-not $updateByKey.ContainsKey($key)) { continue }

                if (-not $cells.ContainsKey($statusCol)) { continue }
                $tNode = $cells[$statusCol].SelectSingleNode("*[local-name()='is']/*[local-name()='t']")
                if (-not $tNode) { continue }
                $newStatus = $updateByKey[$key].NewStatus
                $oldStatus = $tNode.InnerText
                if ($oldStatus -eq $newStatus) { continue }
                $tNode.InnerText = $newStatus
                $applied.Add(@{ FindingId = $findingId; ResourceId = $resourceId; PreviousStatus = $oldStatus; NewStatus = $newStatus })
            }

            if ($applied.Count -gt 0) {
                $newXmlText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $sheetXml.DocumentElement.OuterXml
                $sheetEntry.Delete()
                $newEntry = $zip.CreateEntry($partPath, [System.IO.Compression.CompressionLevel]::Optimal)
                $stream = $newEntry.Open()
                try {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($newXmlText)
                    $stream.Write($bytes, 0, $bytes.Length)
                } finally {
                    $stream.Dispose()
                }
            }
        } finally {
            $zip.Dispose()
        }
    } finally {
        $fs.Dispose()
    }
    return $applied.ToArray()
}
