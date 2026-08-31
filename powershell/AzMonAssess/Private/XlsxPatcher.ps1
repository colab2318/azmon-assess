#requires -Version 7.0
# Surgical in-place editor for an already-generated report.xlsx — used by
# Invoke-AzMonVerification to flip review-status cells after re-checking
# findings against the customer's current environment, without touching
# any other cell, style, or sheet. Complements XlsxBuilder.ps1 (which only
# ever builds a workbook from scratch); this module never regenerates a
# sheet, it only rewrites the text of specific already-existing cells.
#
# New-AzMonExcelReport.ps1 writes every string cell as an inline string
# (t="inlineStr", see XlsxBuilder.ps1). But once a human opens the file in
# real Excel and saves it - the whole point of the finalize workflow's
# manual review step - Excel rewrites every string cell to ITS preferred
# shared-strings format (t="s", an index into xl/sharedStrings.xml)
# regardless of how the file looked when opened. Every read below handles
# BOTH encodings; writes always rewrite the touched cell back to a fresh
# inline string so later reads never depend on a possibly-stale
# sharedStrings.xml.

function Assert-AzMonXlsxPackage {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Path)

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $signature = [byte[]]::new(8)
        $bytesRead = $stream.Read($signature, 0, $signature.Length)
    } finally {
        $stream.Dispose()
    }

    $signatureHex = (($signature | ForEach-Object { $_.ToString('X2') }) -join '')
    if ($bytesRead -ge 8 -and $signatureHex -eq 'D0CF11E0A1B11AE1') {
        throw "Workbook '$Path' is an encrypted or legacy binary Office file, not a readable .xlsx package. In Excel, remove password or sensitivity-label encryption if permitted, then use Save As > Excel Workbook (*.xlsx) and retry. Renaming the extension is not sufficient."
    }
    if ($bytesRead -lt 2 -or $signature[0] -ne 0x50 -or $signature[1] -ne 0x4B) {
        throw "Workbook '$Path' is not a valid .xlsx package. Save it from Excel as Excel Workbook (*.xlsx), ensure the save completes, and retry."
    }
}

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

function Get-AzMonXlsxSharedStrings {
    <#
    .SYNOPSIS
        Loads xl/sharedStrings.xml (if the package has one) into an
        ordered array of plain text, indexed the same way a t="s" cell's
        <v> value references it. Returns an empty array for a workbook
        with no shared-strings part (e.g. a freshly-generated report.xlsx,
        which uses only inline strings until a human opens/saves it).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Zip)

    $entry = $Zip.GetEntry('xl/sharedStrings.xml')
    if (-not $entry) { return @() }
    [xml]$xml = Read-AzMonZipEntryText -Entry $entry
    $siNodes = $xml.SelectNodes("//*[local-name()='sst']/*[local-name()='si']")
    return @(foreach ($si in $siNodes) {
            ($si.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join ''
        })
}

function Get-AzMonXlsxCellText {
    <#
    .SYNOPSIS
        Reads a cell's text regardless of whether it's stored as an
        inline string (t="inlineStr", written by this tool), a shared
        string (t="s", the format Excel converts cells to on save - looks
        up SharedStrings by index), a formula string result (t="str"), or
        a bare value.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Cell, [string[]] $SharedStrings = @())

    $type = $Cell.GetAttribute('t')
    if ($type -eq 'inlineStr') {
        $isNode = $Cell.SelectSingleNode("*[local-name()='is']")
        if (-not $isNode) { return $null }
        return (($isNode.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join '')
    }
    if ($type -eq 's') {
        $vNode = $Cell.SelectSingleNode("*[local-name()='v']")
        if (-not $vNode) { return $null }
        $idx = 0
        if (-not [int]::TryParse($vNode.InnerText, [ref] $idx)) { return $null }
        if ($idx -lt 0 -or $idx -ge $SharedStrings.Count) { return $null }
        return $SharedStrings[$idx]
    }
    $vNode = $Cell.SelectSingleNode("*[local-name()='v']")
    if ($vNode) { return $vNode.InnerText }
    return $null
}

function Set-AzMonXlsxCellInlineText {
    <#
    .SYNOPSIS
        Rewrites a cell to a fresh inline string (t="inlineStr"),
        replacing whatever it held before (shared string, formula result,
        number, or an existing inline string) - so the cell keeps reading
        correctly via Get-AzMonXlsxCellText no matter how many times Excel
        has round-tripped the file since. Leaves every other attribute on
        the cell (r, s/style index, ...) untouched.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Cell, [Parameter(Mandatory)] [string] $Text)

    while ($Cell.HasChildNodes) { [void]$Cell.RemoveChild($Cell.FirstChild) }
    $Cell.SetAttribute('t', 'inlineStr')
    $isEl = $Cell.OwnerDocument.CreateElement('is', $Cell.NamespaceURI)
    $tEl = $Cell.OwnerDocument.CreateElement('t', $Cell.NamespaceURI)
    $tEl.InnerText = $Text
    [void]$isEl.AppendChild($tEl)
    [void]$Cell.AppendChild($isEl)
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
    param([Parameter(Mandatory)] $SheetXml, [string[]] $SharedStrings = @())

    $headerRow = $SheetXml.SelectSingleNode("//*[local-name()='sheetData']/*[local-name()='row'][1]")
    $map = @{}
    if (-not $headerRow) { return $map }
    foreach ($cell in $headerRow.SelectNodes("*[local-name()='c']")) {
        $colLetter = ($cell.GetAttribute('r') -replace '\d+$', '')
        $text = Get-AzMonXlsxCellText -Cell $cell -SharedStrings $SharedStrings
        if ($text) { $map[$colLetter] = $text }
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
    Assert-AzMonXlsxPackage -Path $Path
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
            $sharedStrings = Get-AzMonXlsxSharedStrings -Zip $zip
            $colMap = Get-AzMonXlsxHeaderColumnMap -SheetXml $sheetXml -SharedStrings $sharedStrings
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
                $findingId = Get-AzMonXlsxCellText -Cell $cells[$keyCol] -SharedStrings $sharedStrings
                if (-not $findingId) { continue }
                $resourceId = ''
                if ($resourceCol -and $cells.ContainsKey($resourceCol)) {
                    $resourceId = Get-AzMonXlsxCellText -Cell $cells[$resourceCol] -SharedStrings $sharedStrings
                }
                $statusText = ''
                if ($cells.ContainsKey($statusCol)) {
                    $statusText = Get-AzMonXlsxCellText -Cell $cells[$statusCol] -SharedStrings $sharedStrings
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
    Assert-AzMonXlsxPackage -Path $Path
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
            $sharedStrings = Get-AzMonXlsxSharedStrings -Zip $zip
            $colMap = Get-AzMonXlsxHeaderColumnMap -SheetXml $sheetXml -SharedStrings $sharedStrings
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
                $findingId = Get-AzMonXlsxCellText -Cell $cells[$keyCol] -SharedStrings $sharedStrings
                if (-not $findingId) { continue }
                $resourceId = ''
                if ($resourceCol -and $cells.ContainsKey($resourceCol)) {
                    $resourceId = Get-AzMonXlsxCellText -Cell $cells[$resourceCol] -SharedStrings $sharedStrings
                }
                $key = "$findingId|$(([string]$resourceId).ToLowerInvariant())"
                if (-not $updateByKey.ContainsKey($key)) { continue }

                if (-not $cells.ContainsKey($statusCol)) { continue }
                $cellNode = $cells[$statusCol]
                $oldStatus = Get-AzMonXlsxCellText -Cell $cellNode -SharedStrings $sharedStrings
                $newStatus = $updateByKey[$key].NewStatus
                if ($oldStatus -eq $newStatus) { continue }
                Set-AzMonXlsxCellInlineText -Cell $cellNode -Text $newStatus
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
