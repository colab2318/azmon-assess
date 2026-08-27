#requires -Version 7.0
# Zero-dependency Excel (.xlsx) builder. Produces plain OOXML using only
# System.IO.Compression — no Excel install, no COM, no ImportExcel/EPPlus.
# Mirrors the approach already used by Private/PptxBuilder.ps1 (mutable
# hashtable objects -> serialized to XML -> zipped at Save time). Strings
# are written as inline strings (t="inlineStr") so no shared-strings table
# is needed; styles are dynamically registered and deduped by descriptor key.

function New-AzMonXlsxWorkbook {
    [CmdletBinding()]
    param()
    $wb = @{
        Kind         = 'XlsxWorkbook'
        Sheets       = [System.Collections.Generic.List[hashtable]]::new()
        Fonts        = [System.Collections.Generic.List[hashtable]]::new()
        Fills        = [System.Collections.Generic.List[hashtable]]::new()
        Xfs          = [System.Collections.Generic.List[hashtable]]::new()
        StyleLookup  = @{}
    }
    # Reserved defaults every xlsx styles.xml is expected to carry.
    [void]$wb.Fonts.Add(@{ Bold = $false; ColorHex = $null; Underline = $false; SizePt = 11 })
    [void]$wb.Fills.Add(@{ PatternType = 'none' })
    [void]$wb.Fills.Add(@{ PatternType = 'gray125' })
    [void]$wb.Xfs.Add(@{ FontId = 0; FillId = 0; WrapText = $false; VerticalTop = $false; HorizontalAlign = $null })
    $wb.StyleLookup['default'] = 0
    return $wb
}

function Add-AzMonXlsxSheet {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Workbook, [Parameter(Mandatory)] [string] $Name)
    $sheet = @{
        Kind             = 'XlsxSheet'
        Name             = $Name
        Rows             = [System.Collections.Generic.List[hashtable]]::new()
        ColumnWidths     = @{}
        FreezeHeaderRows = 0
        AutoFilterRange  = $null
        MergedCells      = [System.Collections.Generic.List[string]]::new()
        Hyperlinks       = [System.Collections.Generic.List[hashtable]]::new()
        MaxColumnCount   = 0
    }
    $Workbook.Sheets.Add($sheet)
    return $sheet
}

function ConvertTo-AzMonXlsxColumnLetter {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [int] $Index)
    $n = $Index
    $letters = ''
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        $letters = [string][char](65 + $rem) + $letters
        # NOTE: PowerShell's / between ints yields a double when not evenly
        # divisible, and [Math]::Floor(double) returns a double too — cast
        # back to [int] or the next [char] cast on $rem fails with an
        # "Invalid cast from Decimal/Double to Char" error.
        $n = [int][Math]::Floor(($n - 1) / 26)
    }
    return $letters
}

function Get-AzMonXlsxStyleId {
    <#
    .SYNOPSIS
        Returns (creating if necessary) a cellXfs index for the given style
        intent. Dedups fonts/fills/xfs by a composite descriptor key.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Workbook,
        [switch] $Bold,
        [string] $FontColorHex,
        [switch] $Underline,
        [int] $FontSizePt = 11,
        [string] $FillHex,
        [switch] $WrapText,
        [switch] $VerticalTop,
        [ValidateSet('left', 'center', 'right', '')] [string] $HorizontalAlign
    )
    $key = "b=$($Bold.IsPresent);fc=$FontColorHex;u=$($Underline.IsPresent);sz=$FontSizePt;fill=$FillHex;wrap=$($WrapText.IsPresent);vt=$($VerticalTop.IsPresent);ha=$HorizontalAlign"
    if ($Workbook.StyleLookup.ContainsKey($key)) { return $Workbook.StyleLookup[$key] }

    $fontId = -1
    for ($i = 0; $i -lt $Workbook.Fonts.Count; $i++) {
        $f = $Workbook.Fonts[$i]
        if ($f.Bold -eq $Bold.IsPresent -and $f.ColorHex -eq $FontColorHex -and $f.Underline -eq $Underline.IsPresent -and $f.SizePt -eq $FontSizePt) { $fontId = $i; break }
    }
    if ($fontId -lt 0) {
        [void]$Workbook.Fonts.Add(@{ Bold = $Bold.IsPresent; ColorHex = $FontColorHex; Underline = $Underline.IsPresent; SizePt = $FontSizePt })
        $fontId = $Workbook.Fonts.Count - 1
    }

    $fillId = 0
    if ($FillHex) {
        for ($i = 0; $i -lt $Workbook.Fills.Count; $i++) {
            $fl = $Workbook.Fills[$i]
            if ($fl.PatternType -eq 'solid' -and $fl.FgColorHex -eq $FillHex) { $fillId = $i; break }
        }
        if ($fillId -eq 0) {
            [void]$Workbook.Fills.Add(@{ PatternType = 'solid'; FgColorHex = $FillHex })
            $fillId = $Workbook.Fills.Count - 1
        }
    }

    [void]$Workbook.Xfs.Add(@{
            FontId          = $fontId
            FillId          = $fillId
            WrapText        = $WrapText.IsPresent
            VerticalTop     = $VerticalTop.IsPresent
            HorizontalAlign = $HorizontalAlign
        })
    $xfId = $Workbook.Xfs.Count - 1
    $Workbook.StyleLookup[$key] = $xfId
    return $xfId
}

function New-AzMonXlsxCell {
    <#
    .SYNOPSIS
        Builds one cell descriptor. Value type (string vs number) is
        auto-detected unless -AsString is forced.
    #>
    [CmdletBinding()]
    param(
        $Value,
        [switch] $AsString,
        [switch] $Bold,
        [string] $FillHex,
        [switch] $WrapText,
        [switch] $VerticalTop,
        [ValidateSet('left', 'center', 'right', '')] [string] $HorizontalAlign,
        [string] $FontColorHex,
        [switch] $Underline,
        [int] $FontSizePt = 11
    )
    $isNumeric = (-not $AsString) -and ($null -ne $Value) -and ($Value -is [double] -or $Value -is [int] -or $Value -is [int64] -or $Value -is [float] -or $Value -is [decimal])
    return @{
        Value           = $Value
        IsNumeric       = $isNumeric
        Bold            = $Bold.IsPresent
        FillHex         = $FillHex
        WrapText        = $WrapText.IsPresent
        VerticalTop     = $VerticalTop.IsPresent
        HorizontalAlign = $HorizontalAlign
        FontColorHex    = $FontColorHex
        Underline       = $Underline.IsPresent
        FontSizePt      = $FontSizePt
    }
}

function Add-AzMonXlsxRow {
    <#
    .SYNOPSIS
        Appends a row of plain values (auto-boxed to cells) or pre-built
        cell descriptors (from New-AzMonXlsxCell). Returns the 1-based row
        number just written.
    .NOTES
        -Cell is intentionally NOT [Parameter(Mandatory)]: PowerShell's
        Mandatory-parameter validation rejects an array argument if ANY
        element is $null (throws "Cannot bind argument ... because it is
        null"), and cell arrays legitimately contain $null for unset
        nullable fields (e.g. DailyQuotaGb). AllowEmptyCollection() plus a
        default of @() covers the "nothing passed" case instead.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Sheet,
        [AllowEmptyCollection()] [array] $Cell = @(),
        [double] $HeightPt
    )
    $cells = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($c in $Cell) {
        if ($c -is [hashtable] -and $c.ContainsKey('IsNumeric')) { $cells.Add($c) }
        else { $cells.Add((New-AzMonXlsxCell -Value $c)) }
    }
    $row = @{ Cells = $cells; HeightPt = $HeightPt }
    $Sheet.Rows.Add($row)
    if ($cells.Count -gt $Sheet.MaxColumnCount) { $Sheet.MaxColumnCount = $cells.Count }
    return $Sheet.Rows.Count
}

function Add-AzMonXlsxHeaderRow {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Sheet, [Parameter(Mandatory)] [string[]] $Header)
    $cells = @($Header | ForEach-Object {
            New-AzMonXlsxCell -Value $_ -AsString -Bold -FillHex $script:AzMonXlsxHeaderFillHex -FontColorHex 'FFFFFF' -HorizontalAlign center
        })
    return (Add-AzMonXlsxRow -Sheet $Sheet -Cell $cells)
}

function Set-AzMonXlsxRowFill {
    <#
    .NOTES
        Local var is $rowObj, not $row — PowerShell variables are
        case-insensitive, so a local $row would alias the type-constrained
        [int] $Row parameter and blow up on assigning it a hashtable.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Sheet, [Parameter(Mandatory)] [int] $Row, [int] $ColumnCount, [string] $Hex)
    if (-not $Hex) { return }
    $rowObj = $Sheet.Rows[$Row - 1]
    if (-not $ColumnCount) { $ColumnCount = $rowObj.Cells.Count }
    for ($i = 0; $i -lt $ColumnCount; $i++) {
        if ($i -ge $rowObj.Cells.Count) { $rowObj.Cells.Add((New-AzMonXlsxCell -Value $null -AsString)) }
        $rowObj.Cells[$i].FillHex = $Hex
    }
}

function Set-AzMonXlsxCellWrap {
    <#
    .NOTES
        Local var is $rowObj — see note in Set-AzMonXlsxRowFill. Grows the
        Cells list first (mirrors Set-AzMonXlsxHyperlink) so a Column past
        the row's populated cell count doesn't throw an out-of-range error.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Sheet, [Parameter(Mandatory)] [int] $Row, [Parameter(Mandatory)] [int] $Column)
    $rowObj = $Sheet.Rows[$Row - 1]
    while ($rowObj.Cells.Count -lt $Column) { $rowObj.Cells.Add((New-AzMonXlsxCell -Value $null -AsString)) }
    $rowObj.Cells[$Column - 1].WrapText = $true
    $rowObj.Cells[$Column - 1].VerticalTop = $true
}

function Set-AzMonXlsxHyperlink {
    <#
    .NOTES
        Local var is $rowObj — see note in Set-AzMonXlsxRowFill.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Sheet,
        [Parameter(Mandatory)] [int] $Row,
        [Parameter(Mandatory)] [int] $Column,
        [Parameter(Mandatory)] [string] $Url,
        [string] $Text = 'Open'
    )
    $rowObj = $Sheet.Rows[$Row - 1]
    while ($rowObj.Cells.Count -lt $Column) { $rowObj.Cells.Add((New-AzMonXlsxCell -Value $null -AsString)) }
    $rowObj.Cells[$Column - 1] = New-AzMonXlsxCell -Value $Text -AsString -FontColorHex '0563C1' -Underline
    $Sheet.Hyperlinks.Add(@{ Row = $Row; Column = $Column; Url = $Url })
}

function Set-AzMonXlsxColumnWidths {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Sheet, [Parameter(Mandatory)] [double[]] $Width)
    for ($i = 0; $i -lt $Width.Count; $i++) { $Sheet.ColumnWidths[$i + 1] = $Width[$i] }
}

function Set-AzMonXlsxFreezeHeader {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Sheet, [int] $Rows = 1)
    $Sheet.FreezeHeaderRows = $Rows
}

function Set-AzMonXlsxAutoFilter {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Sheet, [int] $StartRow = 1, [int] $RowCount, [int] $ColumnCount)
    if (-not $RowCount) { $RowCount = $Sheet.Rows.Count }
    if (-not $ColumnCount) { $ColumnCount = $Sheet.MaxColumnCount }
    if ($RowCount -le 0 -or $ColumnCount -le 0) { return }
    $lastCol = ConvertTo-AzMonXlsxColumnLetter $ColumnCount
    $Sheet.AutoFilterRange = "A${StartRow}:$lastCol$RowCount"
}

function Merge-AzMonXlsxCells {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Sheet, [Parameter(Mandatory)] [string] $Range)
    $Sheet.MergedCells.Add($Range)
}

function Set-AzMonXlsxRowHeight {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Sheet, [Parameter(Mandatory)] [int] $Row, [Parameter(Mandatory)] [double] $HeightPt)
    $Sheet.Rows[$Row - 1].HeightPt = $HeightPt
}

# ---- style palette (shared with the report builders) -------------------

$script:AzMonXlsxHeaderFillHex = '1F4E78'
$script:AzMonXlsxSevFillHex = @{
    critical = 'FFC7CE'
    high     = 'FFEB9C'
    medium   = 'FCE4D6'
    low      = 'C6EFCE'
    info     = 'D9E1F2'
}

# ---- XML serialization ---------------------------------------------------

function Get-AzMonXlsxCellXml {
    param([Parameter(Mandatory)] [hashtable] $Workbook, [Parameter(Mandatory)] [hashtable] $Cell, [Parameter(Mandatory)] [string] $CellRef)
    $value = $Cell.Value
    if ($null -eq $value -or ([string]$value) -eq '') {
        if (-not $Cell.FillHex) { return '' }
        $styleId = Get-AzMonXlsxStyleId -Workbook $Workbook -FillHex $Cell.FillHex
        return "<c r=`"$CellRef`" s=`"$styleId`"/>"
    }
    $styleId = Get-AzMonXlsxStyleId -Workbook $Workbook -Bold:$Cell.Bold -FillHex $Cell.FillHex -WrapText:$Cell.WrapText `
        -VerticalTop:$Cell.VerticalTop -HorizontalAlign $Cell.HorizontalAlign -FontColorHex $Cell.FontColorHex -Underline:$Cell.Underline -FontSizePt ($Cell.FontSizePt ?? 11)
    if ($Cell.IsNumeric) {
        $num = [System.Convert]::ToString([double]$value, [System.Globalization.CultureInfo]::InvariantCulture)
        return "<c r=`"$CellRef`" s=`"$styleId`"><v>$num</v></c>"
    }
    $escaped = ConvertTo-AzMonXmlText ([string]$value)
    return "<c r=`"$CellRef`" t=`"inlineStr`" s=`"$styleId`"><is><t xml:space=`"preserve`">$escaped</t></is></c>"
}

function Get-AzMonXlsxSheetXml {
    param([Parameter(Mandatory)] [hashtable] $Workbook, [Parameter(Mandatory)] [hashtable] $Sheet)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')

    $lastCol = [Math]::Max(1, $Sheet.MaxColumnCount)
    $lastColLetter = ConvertTo-AzMonXlsxColumnLetter $lastCol
    $lastRow = [Math]::Max(1, $Sheet.Rows.Count)
    [void]$sb.Append("<dimension ref=`"A1:$lastColLetter$lastRow`"/>")

    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0">')
    if ($Sheet.FreezeHeaderRows -gt 0) {
        $topRow = $Sheet.FreezeHeaderRows + 1
        [void]$sb.Append("<pane ySplit=`"$($Sheet.FreezeHeaderRows)`" topLeftCell=`"A$topRow`" activePane=`"bottomLeft`" state=`"frozen`"/><selection pane=`"bottomLeft`" activeCell=`"A$topRow`" sqref=`"A$topRow`"/>")
    }
    [void]$sb.Append('</sheetView></sheetViews>')

    if ($Sheet.ColumnWidths.Count -gt 0) {
        [void]$sb.Append('<cols>')
        foreach ($colIdx in ($Sheet.ColumnWidths.Keys | Sort-Object)) {
            $w = $Sheet.ColumnWidths[$colIdx]
            [void]$sb.Append("<col min=`"$colIdx`" max=`"$colIdx`" width=`"$w`" customWidth=`"1`"/>")
        }
        [void]$sb.Append('</cols>')
    }

    [void]$sb.Append('<sheetData>')
    for ($r = 0; $r -lt $Sheet.Rows.Count; $r++) {
        $rowNum = $r + 1
        $row = $Sheet.Rows[$r]
        $heightAttr = if ($row.HeightPt) { " ht=`"$($row.HeightPt)`" customHeight=`"1`"" } else { '' }
        [void]$sb.Append("<row r=`"$rowNum`"$heightAttr>")
        for ($c = 0; $c -lt $row.Cells.Count; $c++) {
            $colLetter = ConvertTo-AzMonXlsxColumnLetter ($c + 1)
            [void]$sb.Append((Get-AzMonXlsxCellXml -Workbook $Workbook -Cell $row.Cells[$c] -CellRef "$colLetter$rowNum"))
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')

    if ($Sheet.AutoFilterRange) { [void]$sb.Append("<autoFilter ref=`"$($Sheet.AutoFilterRange)`"/>") }

    if ($Sheet.MergedCells.Count -gt 0) {
        [void]$sb.Append("<mergeCells count=`"$($Sheet.MergedCells.Count)`">")
        foreach ($m in $Sheet.MergedCells) { [void]$sb.Append("<mergeCell ref=`"$m`"/>") }
        [void]$sb.Append('</mergeCells>')
    }

    if ($Sheet.Hyperlinks.Count -gt 0) {
        [void]$sb.Append('<hyperlinks>')
        for ($i = 0; $i -lt $Sheet.Hyperlinks.Count; $i++) {
            $h = $Sheet.Hyperlinks[$i]
            $colLetter = ConvertTo-AzMonXlsxColumnLetter $h.Column
            [void]$sb.Append("<hyperlink ref=`"$colLetter$($h.Row)`" r:id=`"rId$($i + 1)`"/>")
        }
        [void]$sb.Append('</hyperlinks>')
    }

    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function Get-AzMonXlsxSheetRelsXml {
    param([Parameter(Mandatory)] [hashtable] $Sheet)
    if ($Sheet.Hyperlinks.Count -eq 0) { return $null }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for ($i = 0; $i -lt $Sheet.Hyperlinks.Count; $i++) {
        $h = $Sheet.Hyperlinks[$i]
        $url = ($h.Url -replace '&', '&amp;')
        [void]$sb.Append("<Relationship Id=`"rId$($i + 1)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`" Target=`"$url`" TargetMode=`"External`"/>")
    }
    [void]$sb.Append('</Relationships>')
    return $sb.ToString()
}

function Get-AzMonXlsxStylesXml {
    param([Parameter(Mandatory)] [hashtable] $Workbook)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

    [void]$sb.Append("<fonts count=`"$($Workbook.Fonts.Count)`">")
    foreach ($f in $Workbook.Fonts) {
        [void]$sb.Append('<font>')
        [void]$sb.Append("<sz val=`"$($f.SizePt)`"/><name val=`"Calibri`"/>")
        if ($f.Bold) { [void]$sb.Append('<b/>') }
        if ($f.Underline) { [void]$sb.Append('<u/>') }
        if ($f.ColorHex) { [void]$sb.Append("<color rgb=`"FF$($f.ColorHex)`"/>") }
        [void]$sb.Append('</font>')
    }
    [void]$sb.Append('</fonts>')

    [void]$sb.Append("<fills count=`"$($Workbook.Fills.Count)`">")
    foreach ($fl in $Workbook.Fills) {
        if ($fl.PatternType -eq 'solid') {
            [void]$sb.Append("<fill><patternFill patternType=`"solid`"><fgColor rgb=`"FF$($fl.FgColorHex)`"/><bgColor indexed=`"64`"/></patternFill></fill>")
        } else {
            [void]$sb.Append("<fill><patternFill patternType=`"$($fl.PatternType)`"/></fill>")
        }
    }
    [void]$sb.Append('</fills>')

    [void]$sb.Append('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>')
    [void]$sb.Append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')

    [void]$sb.Append("<cellXfs count=`"$($Workbook.Xfs.Count)`">")
    foreach ($xf in $Workbook.Xfs) {
        $needsAlign = $xf.WrapText -or $xf.VerticalTop -or $xf.HorizontalAlign
        if ($needsAlign) {
            [void]$sb.Append("<xf numFmtId=`"0`" fontId=`"$($xf.FontId)`" fillId=`"$($xf.FillId)`" borderId=`"0`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`" applyAlignment=`"1`">")
            $wrapAttr = if ($xf.WrapText) { ' wrapText="1"' } else { '' }
            $vAttr = if ($xf.VerticalTop) { ' vertical="top"' } else { '' }
            $hAttr = if ($xf.HorizontalAlign) { " horizontal=`"$($xf.HorizontalAlign)`"" } else { '' }
            [void]$sb.Append("<alignment$wrapAttr$vAttr$hAttr/>")
            [void]$sb.Append('</xf>')
        } else {
            [void]$sb.Append("<xf numFmtId=`"0`" fontId=`"$($xf.FontId)`" fillId=`"$($xf.FillId)`" borderId=`"0`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`"/>")
        }
    }
    [void]$sb.Append('</cellXfs>')

    [void]$sb.Append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
    [void]$sb.Append('</styleSheet>')
    return $sb.ToString()
}

function Get-AzMonXlsxWorkbookXml {
    param([Parameter(Mandatory)] [hashtable] $Workbook)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.Append('<sheets>')
    for ($i = 0; $i -lt $Workbook.Sheets.Count; $i++) {
        $n = $i + 1
        $escaped = ConvertTo-AzMonXmlText $Workbook.Sheets[$i].Name
        [void]$sb.Append("<sheet name=`"$escaped`" sheetId=`"$n`" r:id=`"rId$n`"/>")
    }
    [void]$sb.Append('</sheets></workbook>')
    return $sb.ToString()
}

function Get-AzMonXlsxWorkbookRelsXml {
    param([Parameter(Mandatory)] [int] $SheetCount)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for ($i = 0; $i -lt $SheetCount; $i++) {
        $n = $i + 1
        [void]$sb.Append("<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$n.xml`"/>")
    }
    $stylesId = $SheetCount + 1
    [void]$sb.Append("<Relationship Id=`"rId$stylesId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>")
    [void]$sb.Append('</Relationships>')
    return $sb.ToString()
}

function Get-AzMonXlsxContentTypesXml {
    param([Parameter(Mandatory)] [int] $SheetCount)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
    [void]$sb.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    [void]$sb.Append('<Default Extension="xml" ContentType="application/xml"/>')
    [void]$sb.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    [void]$sb.Append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    [void]$sb.Append('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>')
    [void]$sb.Append('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
    for ($i = 1; $i -le $SheetCount; $i++) {
        [void]$sb.Append("<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")
    }
    [void]$sb.Append('</Types>')
    return $sb.ToString()
}

function Get-AzMonXlsxCoreXml {
    param([string] $Title = 'Azure Monitoring Assessment')
    $escaped = ConvertTo-AzMonXmlText $Title
    $now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>$escaped</dc:title>
<dc:creator>azmon-assess</dc:creator>
<cp:lastModifiedBy>azmon-assess</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified>
</cp:coreProperties>
"@
}

function Get-AzMonXlsxAppXml {
    param([Parameter(Mandatory)] [string[]] $SheetName)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">')
    [void]$sb.Append('<Application>azmon-assess</Application>')
    [void]$sb.Append("<HeadingPairs><vt:vector size=`"2`" baseType=`"variant`"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>$($SheetName.Count)</vt:i4></vt:variant></vt:vector></HeadingPairs>")
    [void]$sb.Append("<TitlesOfParts><vt:vector size=`"$($SheetName.Count)`" baseType=`"lpstr`">")
    foreach ($n in $SheetName) { [void]$sb.Append("<vt:lpstr>$(ConvertTo-AzMonXmlText $n)</vt:lpstr>") }
    [void]$sb.Append('</vt:vector></TitlesOfParts>')
    [void]$sb.Append('</Properties>')
    return $sb.ToString()
}

$script:AzMonXlsxRootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

function Save-AzMonXlsxWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Workbook,
        [Parameter(Mandatory)] [string] $Path,
        [string] $Title = 'Azure Monitoring Assessment'
    )
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $dir = Split-Path -Parent $Path
    if ($dir) { New-AzMonOutputDirectory -Path $dir | Out-Null }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }

    $sheetCount = $Workbook.Sheets.Count
    # Pre-render sheet XML first so any lazily-registered styles (fills,
    # fonts, hyperlink font) land in styles.xml before it is serialized.
    $sheetXmls = [System.Collections.Generic.List[string]]::new()
    $sheetRelXmls = [System.Collections.Generic.List[string]]::new()
    foreach ($sheet in $Workbook.Sheets) {
        $sheetXmls.Add((Get-AzMonXlsxSheetXml -Workbook $Workbook -Sheet $sheet))
        $sheetRelXmls.Add((Get-AzMonXlsxSheetRelsXml -Sheet $sheet))
    }

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            Write-AzMonZipEntry -Zip $zip -EntryName '[Content_Types].xml' -Content (Get-AzMonXlsxContentTypesXml -SheetCount $sheetCount)
            Write-AzMonZipEntry -Zip $zip -EntryName '_rels/.rels' -Content $script:AzMonXlsxRootRels
            Write-AzMonZipEntry -Zip $zip -EntryName 'docProps/core.xml' -Content (Get-AzMonXlsxCoreXml -Title $Title)
            Write-AzMonZipEntry -Zip $zip -EntryName 'docProps/app.xml' -Content (Get-AzMonXlsxAppXml -SheetName @($Workbook.Sheets | ForEach-Object { $_.Name }))
            Write-AzMonZipEntry -Zip $zip -EntryName 'xl/workbook.xml' -Content (Get-AzMonXlsxWorkbookXml -Workbook $Workbook)
            Write-AzMonZipEntry -Zip $zip -EntryName 'xl/_rels/workbook.xml.rels' -Content (Get-AzMonXlsxWorkbookRelsXml -SheetCount $sheetCount)
            Write-AzMonZipEntry -Zip $zip -EntryName 'xl/styles.xml' -Content (Get-AzMonXlsxStylesXml -Workbook $Workbook)

            for ($i = 0; $i -lt $sheetCount; $i++) {
                $n = $i + 1
                Write-AzMonZipEntry -Zip $zip -EntryName "xl/worksheets/sheet$n.xml" -Content $sheetXmls[$i]
                if ($sheetRelXmls[$i]) {
                    Write-AzMonZipEntry -Zip $zip -EntryName "xl/worksheets/_rels/sheet$n.xml.rels" -Content $sheetRelXmls[$i]
                }
            }
        } finally {
            $zip.Dispose()
        }
    } finally {
        $fs.Dispose()
    }
    return (Resolve-Path -LiteralPath $Path).Path
}
