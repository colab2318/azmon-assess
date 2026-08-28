#requires -Version 7.0
# Zero-dependency PowerPoint (.pptx) builder. Produces plain OOXML using only
# System.IO.Compression — no Office install, no COM, no external module.
# Charts are approximated with plain rectangles/text (proportional bars,
# stacked "share" bars) instead of native chart parts, which keeps the XML
# surface small and avoids embedding a workbook part.

$script:AzMonEmuPerInch = 914400
$script:AzMonSlideWidthIn = 13.333
$script:AzMonSlideHeightIn = 7.5
$script:AzMonSlideWidthEmu = 12192000
$script:AzMonSlideHeightEmu = 6858000

# Microsoft logo (white wordmark), extracted from the customer's own
# Executive Summary brand template - used on the title slide only.
$script:AzMonPptxMsLogoEmfBase64 = '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'

function Convert-AzMonInchesToEmu {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [double] $Inches)
    return [int64][Math]::Round($Inches * $script:AzMonEmuPerInch)
}

function ConvertTo-AzMonXmlText {
    [CmdletBinding()]
    param([string] $Text)
    if ($null -eq $Text) { return '' }
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;')
}

function New-AzMonPptxDeck {
    [CmdletBinding()]
    param()
    return @{
        Kind   = 'PptxDeck'
        Slides = [System.Collections.Generic.List[hashtable]]::new()
    }
}

function New-AzMonPptxSlide {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Deck)
    $slide = @{ ShapeXml = [System.Text.StringBuilder]::new(); NextId = 2 }
    $Deck.Slides.Add($slide)
    return $slide
}

function Get-AzMonPptxNextId {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Slide)
    $id = $Slide.NextId
    $Slide.NextId = $id + 1
    return $id
}

# ---- shape primitives ---------------------------------------------------

function Add-AzMonPptxRect {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [double] $X, [Parameter(Mandatory)] [double] $Y,
        [Parameter(Mandatory)] [double] $Width, [Parameter(Mandatory)] [double] $Height,
        [Parameter(Mandatory)] [string] $ColorHex
    )
    $id = Get-AzMonPptxNextId -Slide $Slide
    $xEmu = Convert-AzMonInchesToEmu $X; $yEmu = Convert-AzMonInchesToEmu $Y
    $cxEmu = Convert-AzMonInchesToEmu $Width; $cyEmu = Convert-AzMonInchesToEmu $Height
    [void]$Slide.ShapeXml.Append(@"
<p:sp><p:nvSpPr><p:cNvPr id="$id" name="Rect$id"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="$xEmu" y="$yEmu"/><a:ext cx="$cxEmu" cy="$cyEmu"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="$ColorHex"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>
"@)
}

function Add-AzMonPptxText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [double] $X, [Parameter(Mandatory)] [double] $Y,
        [Parameter(Mandatory)] [double] $Width, [Parameter(Mandatory)] [double] $Height,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Text,
        [int] $Size = 18,
        [switch] $Bold,
        [string] $ColorHex = '091F2C',
        [ValidateSet('l', 'ctr', 'r')] [string] $Align = 'l'
    )
    $id = Get-AzMonPptxNextId -Slide $Slide
    $xEmu = Convert-AzMonInchesToEmu $X; $yEmu = Convert-AzMonInchesToEmu $Y
    $cxEmu = Convert-AzMonInchesToEmu $Width; $cyEmu = Convert-AzMonInchesToEmu $Height
    $szHundredths = $Size * 100
    $boldAttr = if ($Bold) { ' b="1"' } else { '' }
    $escaped = ConvertTo-AzMonXmlText $Text
    [void]$Slide.ShapeXml.Append(@"
<p:sp><p:nvSpPr><p:cNvPr id="$id" name="TextBox$id"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="$xEmu" y="$yEmu"/><a:ext cx="$cxEmu" cy="$cyEmu"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="$Align"/><a:r><a:rPr lang="en-US" sz="$szHundredths"$boldAttr dirty="0"><a:solidFill><a:srgbClr val="$ColorHex"/></a:solidFill></a:rPr><a:t>$escaped</a:t></a:r></a:p></p:txBody></p:sp>
"@)
}

function Add-AzMonPptxBullets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [double] $X, [Parameter(Mandatory)] [double] $Y,
        [Parameter(Mandatory)] [double] $Width, [Parameter(Mandatory)] [double] $Height,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [AllowEmptyString()] [string[]] $Bullet,
        [int] $Size = 14,
        [string] $ColorHex = '091F2C'
    )
    $id = Get-AzMonPptxNextId -Slide $Slide
    $xEmu = Convert-AzMonInchesToEmu $X; $yEmu = Convert-AzMonInchesToEmu $Y
    $cxEmu = Convert-AzMonInchesToEmu $Width; $cyEmu = Convert-AzMonInchesToEmu $Height
    $szHundredths = $Size * 100
    $bulletChar = [string][char]0x2022
    $paras = [System.Text.StringBuilder]::new()
    foreach ($line in $Bullet) {
        $prefix = if ([string]::IsNullOrEmpty($line)) { '' } else { "$bulletChar " }
        $escaped = ConvertTo-AzMonXmlText "$prefix$line"
        [void]$paras.Append(@"
<a:p><a:r><a:rPr lang="en-US" sz="$szHundredths" dirty="0"><a:solidFill><a:srgbClr val="$ColorHex"/></a:solidFill></a:rPr><a:t>$escaped</a:t></a:r></a:p>
"@)
    }
    $body = $paras.ToString()
    [void]$Slide.ShapeXml.Append(@"
<p:sp><p:nvSpPr><p:cNvPr id="$id" name="Bullets$id"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="$xEmu" y="$yEmu"/><a:ext cx="$cxEmu" cy="$cyEmu"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/>$body</p:txBody></p:sp>
"@)
}

function Add-AzMonPptxSlideHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [string] $Title,
        [string] $Subtitle,
        [string] $NavyHex = '091F2C', [string] $GreyHex = '454142', [string] $WhiteHex = 'FFFFFF'
    )
    Add-AzMonPptxRect -Slide $Slide -X 0 -Y 0 -Width $script:AzMonSlideWidthIn -Height 0.9 -ColorHex $NavyHex
    Add-AzMonPptxText -Slide $Slide -X 0.5 -Y 0.2 -Width 12 -Height 0.6 -Text $Title -Size 26 -Bold -ColorHex $WhiteHex
    if ($Subtitle) {
        Add-AzMonPptxText -Slide $Slide -X 0.5 -Y 1.0 -Width 12 -Height 0.4 -Text $Subtitle -Size 13 -ColorHex $GreyHex
    }
}

function Add-AzMonPptxKpiCard {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [double] $X, [Parameter(Mandatory)] [double] $Y,
        [double] $Width = 4.0, [double] $Height = 1.6,
        [Parameter(Mandatory)] [string] $Label,
        [Parameter(Mandatory)] [string] $Value,
        [string] $AccentColorHex = '0078D4'
    )
    Add-AzMonPptxRect -Slide $Slide -X $X -Y $Y -Width $Width -Height $Height -ColorHex 'FFFFFF'
    Add-AzMonPptxRect -Slide $Slide -X $X -Y $Y -Width 0.15 -Height $Height -ColorHex $AccentColorHex
    Add-AzMonPptxText -Slide $Slide -X ($X + 0.4) -Y ($Y + 0.2) -Width ($Width - 0.6) -Height 0.4 -Text $Label.ToUpperInvariant() -Size 11 -Bold -ColorHex '454142'
    Add-AzMonPptxText -Slide $Slide -X ($X + 0.4) -Y ($Y + 0.5) -Width ($Width - 0.6) -Height 1.0 -Text $Value -Size 32 -Bold -ColorHex '091F2C'
}

# ---- pseudo-chart primitives (rectangles, no native chart parts) -------

function Add-AzMonPptxHBarChart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [string[]] $Category,
        [Parameter(Mandatory)] [double[]] $Value,
        [string[]] $ColorHex,
        [double] $X = 0.7, [double] $Y = 1.4, [double] $Width = 8.0, [double] $RowHeight = 0.7,
        [double] $LabelWidth = 1.6
    )
    $max = ($Value | Measure-Object -Maximum).Maximum
    if (-not $max -or $max -le 0) { $max = 1 }
    $barAreaWidth = $Width - $LabelWidth - 0.9
    for ($i = 0; $i -lt $Category.Count; $i++) {
        $rowY = $Y + ($i * $RowHeight)
        $color = if ($ColorHex -and $ColorHex[$i]) { $ColorHex[$i] } else { '0078D4' }
        Add-AzMonPptxText -Slide $Slide -X $X -Y ($rowY + 0.05) -Width $LabelWidth -Height ($RowHeight - 0.1) -Text $Category[$i] -Size 13 -Bold -ColorHex '091F2C'
        $barWidth = [Math]::Max(0.05, ($Value[$i] / $max) * $barAreaWidth)
        Add-AzMonPptxRect -Slide $Slide -X ($X + $LabelWidth) -Y ($rowY + 0.1) -Width $barWidth -Height ($RowHeight - 0.3) -ColorHex $color
        Add-AzMonPptxText -Slide $Slide -X ($X + $LabelWidth + $barWidth + 0.1) -Y ($rowY + 0.05) -Width 0.9 -Height ($RowHeight - 0.1) -Text ([string]$Value[$i]) -Size 13 -Bold -ColorHex '091F2C'
    }
}

function Add-AzMonPptxVBarChart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [string[]] $Category,
        [Parameter(Mandatory)] [double[]] $Value,
        [string] $ColorHex = '0078D4',
        [double] $X = 0.7, [double] $Y = 1.4, [double] $Width = 11.9, [double] $Height = 4.6,
        [string[]] $ValueLabel
    )
    $max = ($Value | Measure-Object -Maximum).Maximum
    if (-not $max -or $max -le 0) { $max = 1 }
    $n = $Category.Count
    if ($n -eq 0) { return }
    $colWidth = $Width / $n
    $barWidth = [Math]::Min($colWidth * 0.55, 1.4)
    for ($i = 0; $i -lt $n; $i++) {
        $colX = $X + ($i * $colWidth) + (($colWidth - $barWidth) / 2)
        $barHeight = [Math]::Max(0.05, ($Value[$i] / $max) * $Height)
        $barY = $Y + ($Height - $barHeight)
        Add-AzMonPptxRect -Slide $Slide -X $colX -Y $barY -Width $barWidth -Height $barHeight -ColorHex $ColorHex
        $label = if ($ValueLabel -and $ValueLabel[$i]) { $ValueLabel[$i] } else { [string]$Value[$i] }
        Add-AzMonPptxText -Slide $Slide -X ($X + ($i * $colWidth)) -Y ($barY - 0.35) -Width $colWidth -Height 0.3 -Text $label -Size 12 -Bold -ColorHex '091F2C' -Align 'ctr'
        Add-AzMonPptxText -Slide $Slide -X ($X + ($i * $colWidth)) -Y ($Y + $Height + 0.05) -Width $colWidth -Height 0.7 -Text $Category[$i] -Size 11 -ColorHex '454142' -Align 'ctr'
    }
}

function Add-AzMonPptxShareBar {
    <#
    .SYNOPSIS
        100%-stacked single bar + legend — a low-risk stand-in for a
        doughnut/pie chart (arc geometry is not worth the added XML risk).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [string[]] $Category,
        [Parameter(Mandatory)] [double[]] $Value,
        [Parameter(Mandatory)] [string[]] $ColorHex,
        [double] $X = 0.7, [double] $Y = 1.6, [double] $Width = 11.5, [double] $BarHeight = 0.9
    )
    $total = ($Value | Measure-Object -Sum).Sum
    if (-not $total -or $total -le 0) { $total = 1 }
    $cursor = $X
    for ($i = 0; $i -lt $Category.Count; $i++) {
        $segWidth = ($Value[$i] / $total) * $Width
        if ($segWidth -gt 0.01) {
            Add-AzMonPptxRect -Slide $Slide -X $cursor -Y $Y -Width $segWidth -Height $BarHeight -ColorHex $ColorHex[$i]
        }
        $cursor += $segWidth
    }
    $legendY = $Y + $BarHeight + 0.3
    for ($i = 0; $i -lt $Category.Count; $i++) {
        $rowY = $legendY + ($i * 0.45)
        Add-AzMonPptxRect -Slide $Slide -X $X -Y ($rowY + 0.05) -Width 0.25 -Height 0.25 -ColorHex $ColorHex[$i]
        $pct = [Math]::Round(($Value[$i] / $total) * 100, 1)
        Add-AzMonPptxText -Slide $Slide -X ($X + 0.35) -Y $rowY -Width 8 -Height 0.4 -Text "$($Category[$i]): $($Value[$i])  ($pct%)" -Size 14 -ColorHex '091F2C'
    }
}

function Add-AzMonPptxPicture {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [double] $X, [Parameter(Mandatory)] [double] $Y,
        [Parameter(Mandatory)] [double] $Width, [Parameter(Mandatory)] [double] $Height,
        [Parameter(Mandatory)] [string] $RelId,
        [string] $Name = 'Picture',
        [string] $Description = ''
    )
    $id = Get-AzMonPptxNextId -Slide $Slide
    $xEmu = Convert-AzMonInchesToEmu $X; $yEmu = Convert-AzMonInchesToEmu $Y
    $cxEmu = Convert-AzMonInchesToEmu $Width; $cyEmu = Convert-AzMonInchesToEmu $Height
    $descAttr = if ($Description) { " descr=`"$(ConvertTo-AzMonXmlText $Description)`"" } else { '' }
    [void]$Slide.ShapeXml.Append(@"
<p:pic><p:nvPicPr><p:cNvPr id="$id" name="$Name"$descAttr/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="$RelId"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="$xEmu" y="$yEmu"/><a:ext cx="$cxEmu" cy="$cyEmu"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>
"@)
}

function Add-AzMonPptxTable {
    <#
    .SYNOPSIS
        Simple grid table (header row + data rows) built from plain
        rectangles/text boxes, since the builder has no native table part.
        Long cell text is the caller's responsibility to pre-truncate —
        a fixed-height row does not auto-expand for overflow text.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Slide,
        [Parameter(Mandatory)] [double] $X, [Parameter(Mandatory)] [double] $Y,
        [Parameter(Mandatory)] [string[]] $Header,
        [Parameter(Mandatory)] [double[]] $ColumnWidth,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Row,
        [double] $HeaderHeight = 0.4,
        [double] $RowHeight = 0.42,
        [int] $HeaderSize = 12,
        [int] $RowSize = 11,
        [string] $HeaderFillHex = '091F2C',
        [string] $HeaderFontHex = 'FFFFFF',
        [string] $AltRowFillHex = 'E8E6DF'
    )
    $totalWidth = ($ColumnWidth | Measure-Object -Sum).Sum
    Add-AzMonPptxRect -Slide $Slide -X $X -Y $Y -Width $totalWidth -Height $HeaderHeight -ColorHex $HeaderFillHex
    $cx = $X
    for ($c = 0; $c -lt $Header.Count; $c++) {
        Add-AzMonPptxText -Slide $Slide -X ($cx + 0.06) -Y ($Y + 0.04) -Width ($ColumnWidth[$c] - 0.12) -Height ($HeaderHeight - 0.06) -Text $Header[$c] -Size $HeaderSize -Bold -ColorHex $HeaderFontHex
        $cx += $ColumnWidth[$c]
    }
    $rowY = $Y + $HeaderHeight
    for ($r = 0; $r -lt $Row.Count; $r++) {
        if ($r % 2 -eq 1) {
            Add-AzMonPptxRect -Slide $Slide -X $X -Y $rowY -Width $totalWidth -Height $RowHeight -ColorHex $AltRowFillHex
        }
        $cx = $X
        $cells = @($Row[$r])
        for ($c = 0; $c -lt $Header.Count; $c++) {
            $val = if ($c -lt $cells.Count) { [string]$cells[$c] } else { '' }
            Add-AzMonPptxText -Slide $Slide -X ($cx + 0.06) -Y ($rowY + 0.03) -Width ($ColumnWidth[$c] - 0.12) -Height ($RowHeight - 0.05) -Text $val -Size $RowSize -ColorHex '091F2C'
            $cx += $ColumnWidth[$c]
        }
        $rowY += $RowHeight
    }
    return $rowY
}

# ---- static OOXML template parts ---------------------------------------

$script:AzMonPptxRootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

$script:AzMonPptxSlideMasterXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>
'@

$script:AzMonPptxSlideMasterRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
'@

$script:AzMonPptxSlideLayoutXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>
</p:sldLayout>
'@

$script:AzMonPptxSlideLayoutRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
'@

$script:AzMonPptxThemeXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="AzMonTheme">
<a:themeElements>
<a:clrScheme name="AzMon">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="091F2C"/></a:dk2>
<a:lt2><a:srgbClr val="E8E6DF"/></a:lt2>
<a:accent1><a:srgbClr val="0078D4"/></a:accent1>
<a:accent2><a:srgbClr val="F4364F"/></a:accent2>
<a:accent3><a:srgbClr val="FF5C39"/></a:accent3>
<a:accent4><a:srgbClr val="49C5B1"/></a:accent4>
<a:accent5><a:srgbClr val="07641D"/></a:accent5>
<a:accent6><a:srgbClr val="454142"/></a:accent6>
<a:hlink><a:srgbClr val="091F2C"/></a:hlink>
<a:folHlink><a:srgbClr val="091F2C"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="AzMon">
<a:majorFont><a:latin typeface="Segoe Sans Text Semibold"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Segoe Sans Text"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="AzMon">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>
'@

$script:AzMonPptxSlideRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>
'@

# Title slide only: adds the rId2 relationship Add-AzMonPptxPicture's logo call expects.
$script:AzMonPptxTitleSlideRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/ms-logo-white.emf"/>
</Relationships>
'@

function Get-AzMonPptxCoreXml {
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

function Get-AzMonPptxAppXml {
    param([int] $SlideCount)
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>azmon-assess</Application>
<Slides>$SlideCount</Slides>
<Company></Company>
</Properties>
"@
}

function Get-AzMonPptxContentTypesXml {
    param([Parameter(Mandatory)] [int] $SlideCount)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append(@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="emf" ContentType="image/x-emf"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
'@)
    for ($i = 1; $i -le $SlideCount; $i++) {
        [void]$sb.Append('<Override PartName="/ppt/slides/slide' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
    }
    [void]$sb.Append('</Types>')
    return $sb.ToString()
}

function Get-AzMonPptxPresentationXml {
    param([Parameter(Mandatory)] [int] $SlideCount)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append(@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>
'@)
    for ($i = 0; $i -lt $SlideCount; $i++) {
        $sldId = 256 + $i
        $rId = 2 + $i
        [void]$sb.Append('<p:sldId id="' + $sldId + '" r:id="rId' + $rId + '"/>')
    }
    [void]$sb.Append(@"
</p:sldIdLst>
<p:sldSz cx="$script:AzMonSlideWidthEmu" cy="$script:AzMonSlideHeightEmu" type="screen16x9"/>
<p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
"@)
    return $sb.ToString()
}

function Get-AzMonPptxPresentationRels {
    param([Parameter(Mandatory)] [int] $SlideCount)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append(@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
'@)
    for ($i = 0; $i -lt $SlideCount; $i++) {
        $rId = 2 + $i
        $n = $i + 1
        [void]$sb.Append('<Relationship Id="rId' + $rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + $n + '.xml"/>')
    }
    [void]$sb.Append('</Relationships>')
    return $sb.ToString()
}

function Get-AzMonPptxSlideXml {
    param([Parameter(Mandatory)] [AllowEmptyString()] [string] $ShapeXml)
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
$ShapeXml
</p:spTree></p:cSld>
<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>
</p:sld>
"@
}

function Write-AzMonZipEntry {
    param([Parameter(Mandatory)] $Zip, [Parameter(Mandatory)] [string] $EntryName, [Parameter(Mandatory)] [AllowEmptyString()] [string] $Content)
    $entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Content)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        $stream.Dispose()
    }
}

function Write-AzMonZipEntryBytes {
    param([Parameter(Mandatory)] $Zip, [Parameter(Mandatory)] [string] $EntryName, [Parameter(Mandatory)] [byte[]] $Bytes)
    $entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $stream.Write($Bytes, 0, $Bytes.Length)
    } finally {
        $stream.Dispose()
    }
}

function Save-AzMonPptxDeck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Deck,
        [Parameter(Mandatory)] [string] $Path,
        [string] $Title = 'Azure Monitoring Assessment'
    )
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $dir = Split-Path -Parent $Path
    if ($dir) { New-AzMonOutputDirectory -Path $dir | Out-Null }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }

    $slideCount = $Deck.Slides.Count
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            Write-AzMonZipEntry -Zip $zip -EntryName '[Content_Types].xml' -Content (Get-AzMonPptxContentTypesXml -SlideCount $slideCount)
            Write-AzMonZipEntry -Zip $zip -EntryName '_rels/.rels' -Content $script:AzMonPptxRootRels
            Write-AzMonZipEntry -Zip $zip -EntryName 'docProps/core.xml' -Content (Get-AzMonPptxCoreXml -Title $Title)
            Write-AzMonZipEntry -Zip $zip -EntryName 'docProps/app.xml' -Content (Get-AzMonPptxAppXml -SlideCount $slideCount)
            Write-AzMonZipEntry -Zip $zip -EntryName 'ppt/presentation.xml' -Content (Get-AzMonPptxPresentationXml -SlideCount $slideCount)
            Write-AzMonZipEntry -Zip $zip -EntryName 'ppt/_rels/presentation.xml.rels' -Content (Get-AzMonPptxPresentationRels -SlideCount $slideCount)
            Write-AzMonZipEntry -Zip $zip -EntryName 'ppt/slideMasters/slideMaster1.xml' -Content $script:AzMonPptxSlideMasterXml
            Write-AzMonZipEntry -Zip $zip -EntryName 'ppt/slideMasters/_rels/slideMaster1.xml.rels' -Content $script:AzMonPptxSlideMasterRels
            Write-AzMonZipEntry -Zip $zip -EntryName 'ppt/slideLayouts/slideLayout1.xml' -Content $script:AzMonPptxSlideLayoutXml
            Write-AzMonZipEntry -Zip $zip -EntryName 'ppt/slideLayouts/_rels/slideLayout1.xml.rels' -Content $script:AzMonPptxSlideLayoutRels
            Write-AzMonZipEntry -Zip $zip -EntryName 'ppt/theme/theme1.xml' -Content $script:AzMonPptxThemeXml
            Write-AzMonZipEntryBytes -Zip $zip -EntryName 'ppt/media/ms-logo-white.emf' -Bytes ([Convert]::FromBase64String($script:AzMonPptxMsLogoEmfBase64))

            for ($i = 0; $i -lt $slideCount; $i++) {
                $n = $i + 1
                $slideXml = Get-AzMonPptxSlideXml -ShapeXml $Deck.Slides[$i].ShapeXml.ToString()
                Write-AzMonZipEntry -Zip $zip -EntryName "ppt/slides/slide$n.xml" -Content $slideXml
                $slideRels = if ($n -eq 1) { $script:AzMonPptxTitleSlideRels } else { $script:AzMonPptxSlideRels }
                Write-AzMonZipEntry -Zip $zip -EntryName "ppt/slides/_rels/slide$n.xml.rels" -Content $slideRels
            }
        } finally {
            $zip.Dispose()
        }
    } finally {
        $fs.Dispose()
    }
}
