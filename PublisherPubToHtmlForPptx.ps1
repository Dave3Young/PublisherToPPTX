<#
Disclaimer: This script is provided for educational and informational purposes only and is offered "as is" with no warranty or guarantee of any kind.
    By using this script, you accept full responsibility for any consequences, including but not limited to data loss, system instability, or unintended results.
    Use at your own risk.

    To run:
   
    for current folder only
    .\PublisherPubToHtmlForPptx.ps1 -Filter "*.pub"
   
    for recursive folder processing
    .\PublisherPubToHtmlForPptx.ps1 -Filter "*.pub" -Recurse

    to rewrite the text layout files (<name>_text.json), which are otherwise kept
    .\PublisherPubToHtmlForPptx.ps1 -Filter "*.pub" -RefreshTextLayout

- scans each Publisher page for pbPicture and pbLinkedPicture shapes,
- recursively scans pbGroup shapes for pictures inside groups,
- exports only those individual picture objects as PNG files,
   does nothing—and does not treat it as an error—when a publication contains no pictures

Microsoft documents pbGroup = 6, pbLinkedPicture = 11, and pbPicture = 13, and Shape.GroupItems is the supported way to inspect shapes inside groups.
Shape.SaveAsPicture() saves a shape as an image, with the output format determined by the filename extension.

Change log - 2026-09-02:
- Added a scan of the publication's scratch area (Document.ScratchArea.Shapes). Pictures parked
  off the page on the pasteboard are not part of Page.Shapes or MasterPage.Shapes, so they were
  silently omitted before. They now export with the page label "scratch".
- Added optional support for pbOLEObject shape types via $INCLUDE_OLE_OBJECTS.
- Added scratch-area TEXT capture: any shape on the pasteboard carrying text (text boxes, and
  autoshapes/WordArt with text) is written to "<name>_scratch_text.txt" and, optionally, also
  rendered to PNG as "<name>_scratch_000_text_NNN.png". Scratch-area content never reaches the
  HTML export, so without this the text is lost. Toggles: $EXPORT_SCRATCH_TEXT and
  $EXPORT_SCRATCH_TEXT_PNG.
- Re-enabled the final "PNG image exports" summary block.

Change log - 2026-08-30:
- Added a separate SkipCounter/$pngSkipCount so PNGs that already existed on disk (and were
  not re-exported) are tallied as "Skipped" instead of being counted as "Successful" exports.
- Stopped incrementing the PNG FailCounter/$pngFailCount for errors that aren't actual
  SaveAsPicture() failures (per-shape inspection errors, group inspection errors, page/master
  page scan errors). Those still mark the file as having had a failure, but no longer inflate
  the PNG-specific failure count.
- Restored and updated the final "PNG image exports" summary block (previously commented out)
  to report Successful / Skipped / Failed counts that now reflect actual export attempts.

  source: www.david-e-young.com with the assistance of Mr. Claude 
#>
param
(
    [ValidateNotNullOrEmpty()]
    [string]
    $Filter,

    [switch]
    $Recurse,

    [switch]
    $RefreshTextLayout
)

if (-not $PSBoundParameters.ContainsKey('Filter')) {
    Write-Error "The -Filter parameter is required."
    exit 1
}

if (-not ($Filter -like "*.pub")) {
    Write-Error "The filter must specify .pub files (e.g., '*.pub' or 'file.pub')."
    exit 1
}

$app = $null

# ------------------------------------------------------------
# Publisher constants
# ------------------------------------------------------------

$HTML_FILTERED_FORMAT = 7

# PbShapeType
$PB_GROUP = 6
$PB_LINKED_PICTURE = 11
$PB_PICTURE = 13
$PB_OLE_OBJECT = 12

# Pasted/embedded artwork sometimes lands as an OLE object rather than a true
# picture shape. Set this to $true to attempt SaveAsPicture() on those as well.
$INCLUDE_OLE_OBJECTS = $false

# Write the text of any text-bearing shape in the scratch area to a .txt file.
$EXPORT_SCRATCH_TEXT = $true

# Also render those scratch-area text shapes to PNG, so the formatting/layout is
# preserved as well as the words.
$EXPORT_SCRATCH_TEXT_PNG = $true

# PbPictureResolution
# 3 = 300 DPI
$PNG_RESOLUTION = 3


# ------------------------------------------------------------
# Recursive picture export function
# ------------------------------------------------------------

function Export-PublisherPictures
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Shapes,

        [Parameter(Mandatory = $true)]
        [string]
        $OutputFolder,

        [Parameter(Mandatory = $true)]
        [string]
        $BaseName,

        [Parameter(Mandatory = $true)]
        [int]
        $PageNumber,

        [Parameter(Mandatory = $false)]
        [string]
        $PageLabel = "page",

        [Parameter(Mandatory = $true)]
        [ref]
        $PictureCounter,

        [Parameter(Mandatory = $true)]
        [ref]
        $SuccessCounter,

        [Parameter(Mandatory = $true)]
        [ref]
        $SkipCounter,

        [Parameter(Mandatory = $true)]
        [ref]
        $FailCounter,

        [Parameter(Mandatory = $true)]
        [ref]
        $FileHadFailure
    )

    $shapeCount = $Shapes.Count

    for ($shapeNumber = 1;
         $shapeNumber -le $shapeCount;
         $shapeNumber++) {

        $shape = $null

        try {

            $shape = $Shapes.Item($shapeNumber)
            $shapeType = $shape.Type

            # ----------------------------------------------------
            # Picture or linked picture
            # ----------------------------------------------------

            if (
                $shapeType -eq $PB_PICTURE -or
                $shapeType -eq $PB_LINKED_PICTURE -or
                ($INCLUDE_OLE_OBJECTS -and $shapeType -eq $PB_OLE_OBJECT)
            ) {

                $PictureCounter.Value++

                $pngFileName = "{0}_{1}_{2:D3}_image_{3:D3}.png" -f `
                    $BaseName,
                    $PageLabel,
                    $PageNumber,
                    $PictureCounter.Value

                $pngFilePath = Join-Path `
                    $OutputFolder `
                    $pngFileName

                if (Test-Path $pngFilePath) {

                    Write-Warning "PNG already exists; skipping: $pngFilePath"

                    $SkipCounter.Value++

                    continue
                }

                try {

                    Write-Output `
                        "Exporting image $($PictureCounter.Value) from $PageLabel $PageNumber..."

                    $shape.SaveAsPicture(
                        [string]$pngFilePath,
                        $PNG_RESOLUTION
                    )

                    if (Test-Path $pngFilePath) {

                        Write-Output "Exported PNG: $pngFilePath"

                        $SuccessCounter.Value++
                    }
                    else {

                        Write-Error "Failed to create PNG: $pngFilePath"

                        $FailCounter.Value++
                        $FileHadFailure.Value = $true
                    }

                }
                catch {

                    Write-Error `
                        "Error exporting image $($PictureCounter.Value) from $PageLabel $PageNumber."

                    Write-Error $_

                    $FailCounter.Value++
                    $FileHadFailure.Value = $true
                }
            }

            # ----------------------------------------------------
            # Grouped shape
            # ----------------------------------------------------

            elseif ($shapeType -eq $PB_GROUP) {

                try {

                    $groupItems = $shape.GroupItems

                    Export-PublisherPictures `
                        -Shapes $groupItems `
                        -OutputFolder $OutputFolder `
                        -BaseName $BaseName `
                        -PageNumber $PageNumber `
                        -PageLabel $PageLabel `
                        -PictureCounter $PictureCounter `
                        -SuccessCounter $SuccessCounter `
                        -SkipCounter $SkipCounter `
                        -FailCounter $FailCounter `
                        -FileHadFailure $FileHadFailure
                }
                catch {

                    Write-Warning `
                        "Unable to inspect grouped shape on $PageLabel $PageNumber."

                    Write-Warning $_

                    $FileHadFailure.Value = $true
                }
            }

        }
        catch {

            Write-Warning `
                "Unable to inspect shape $shapeNumber on $PageLabel $PageNumber."

            Write-Warning $_

            $FileHadFailure.Value = $true
        }
        finally {

            $shape = $null
        }
    }
}


# ------------------------------------------------------------
# Text layout export
#
# Publisher's HTML export writes tabs as runs of spaces and leaves
# out tab stops, and some of its spacing is wrong. This records each
# paragraph's real text and formatting, in points, for stage 2 to
# match to the HTML by its words.
# ------------------------------------------------------------

function Get-PublisherParagraphs
{
    param
    (
        [Parameter(Mandatory = $true)]
        $TextRange
    )

    $paragraphs = New-Object System.Collections.Generic.List[object]

    # Paragraphs end in a carriage return; the text after the last one is empty.
    $texts = ([string]$TextRange.Text) -split "`r"

    for ($i = 1; $i -le $texts.Count; $i++) {

        if ($texts[$i - 1].Length -eq 0) {
            continue
        }

        try {

            $para = $TextRange.Paragraphs($i, 1)
            $format = $para.ParagraphFormat

            $tabs = New-Object System.Collections.Generic.List[object]
            $tabStops = $format.Tabs

            for ($t = 1; $t -le $tabStops.Count; $t++) {

                $stop = $tabStops.Item($t)

                $tabs.Add([ordered]@{
                    pos   = [double]$stop.Position
                    align = [int]$stop.Alignment
                })
            }

            # Where Publisher puts the paragraph on the page, and the top
            # of each of its lines. Lines() runs on past the paragraph, and
            # past the end of the text it keeps returning the last line, so
            # stop at a line that starts after the paragraph or doesn't move on.
            # lineStarts holds where each line starts in the paragraph's text,
            # so stage 2 can break lines where Publisher does, and lineLefts
            # and lineWidths where each line sits across the page, which
            # moves where text wraps around a picture or another frame.
            $lineTops = New-Object System.Collections.Generic.List[double]
            $lineStarts = New-Object System.Collections.Generic.List[int]
            $lineLefts = New-Object System.Collections.Generic.List[double]
            $lineWidths = New-Object System.Collections.Generic.List[double]
            $top = $null
            $left = $null
            $height = $null

            try {

                $top = [double]$para.BoundTop
                $left = [double]$para.BoundLeft
                $height = [double]$para.BoundHeight
                $paraEnd = $para.Start + $para.Length
                $prevStart = -1

                for ($k = 1; $k -le 500; $k++) {

                    $line = $para.Lines($k, 1)

                    $lineTop = [double]$line.BoundTop

                    if ($line.Length -eq 0 -or
                        $line.Start -ge $paraEnd -or
                        $line.Start -le $prevStart -or
                        ($lineTops.Count -gt 0 -and $lineTop -le $lineTops[$lineTops.Count - 1])) {
                        break
                    }

                    $prevStart = $line.Start
                    $lineTops.Add($lineTop)
                    $lineStarts.Add([int]($line.Start - $para.Start))
                    $lineLefts.Add([double]$line.BoundLeft)
                    $lineWidths.Add([double]$line.BoundWidth)
                }
            }
            catch { }

            $paragraphs.Add([ordered]@{
                top         = $top
                left        = $left
                height      = $height
                lineTops    = $lineTops
                lineStarts  = $lineStarts
                lineLefts   = $lineLefts
                lineWidths  = $lineWidths
                text        = $texts[$i - 1]
                align       = [int]$format.Alignment
                firstIndent = [double]$format.FirstLineIndent
                leftIndent  = [double]$format.LeftIndent
                rightIndent = [double]$format.RightIndent
                spaceBefore = [double]$format.SpaceBefore
                spaceAfter  = [double]$format.SpaceAfter
                lineSpacing = [double]$format.LineSpacing
                lineRule    = [int]$format.LineSpacingRule
                font        = [string]$para.Font.Name
                size        = [double]$para.Font.Size
                tabs        = $tabs
            })
        }
        catch {

            Write-Warning "Unable to read paragraph $i of a text frame."
            Write-Warning $_
        }
    }

    return ,$paragraphs
}

function Add-PublisherTextLayout
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Shapes,

        [Parameter(Mandatory = $true)]
        [int]
        $PageNumber,

        [Parameter(Mandatory = $true)]
        $Records
    )

    $shapeCount = $Shapes.Count

    for ($shapeNumber = 1;
         $shapeNumber -le $shapeCount;
         $shapeNumber++) {

        try {

            $shape = $Shapes.Item($shapeNumber)

            if ($shape.Type -eq $PB_GROUP) {

                Add-PublisherTextLayout `
                    -Shapes $shape.GroupItems `
                    -PageNumber $PageNumber `
                    -Records $Records

                continue
            }

            $record = [ordered]@{
                page   = $PageNumber
                name   = [string]$shape.Name
                left   = [double]$shape.Left
                top    = [double]$shape.Top
                width  = [double]$shape.Width
                height = [double]$shape.Height
            }

            # Shapes without a table or text frame raise here; that is normal.
            $hasTable = $false
            try { $hasTable = ($shape.HasTable -eq -1) } catch { }

            if ($hasTable) {

                $cells = New-Object System.Collections.Generic.List[object]
                $rows = $shape.Table.Rows

                for ($r = 1; $r -le $rows.Count; $r++) {

                    $rowCells = $rows.Item($r).Cells

                    for ($c = 1; $c -le $rowCells.Count; $c++) {

                        # Cells covered by a merge raise; skip them.
                        try {

                            $cell = $rowCells.Item($c)

                            $cells.Add([ordered]@{
                                row        = $r
                                col        = $c
                                valign     = [int]$cell.VerticalTextAlignment
                                margins    = @([double]$cell.MarginLeft, [double]$cell.MarginTop,
                                               [double]$cell.MarginRight, [double]$cell.MarginBottom)
                                paragraphs = (Get-PublisherParagraphs -TextRange $cell.TextRange)
                            })
                        }
                        catch { }
                    }
                }

                $record.cells = $cells
                $Records.Add($record)

                continue
            }

            $textRange = $null
            try {
                if ($shape.HasTextFrame -eq -1) {
                    $textRange = $shape.TextFrame.TextRange
                }
            }
            catch { }

            if ($textRange -and $textRange.Length -gt 0) {

                $frame = $shape.TextFrame

                # pbTextFrame (17) or a shape with text, such as an autoshape
                $record.type = [int]$shape.Type
                $record.valign = [int]$frame.VerticalTextAlignment
                $record.margins = @([double]$frame.MarginLeft, [double]$frame.MarginTop,
                                    [double]$frame.MarginRight, [double]$frame.MarginBottom)
                $record.paragraphs = (Get-PublisherParagraphs -TextRange $textRange)

                # Text that doesn't fit is hidden; the export still writes it.
                $overflowing = $false
                try { $overflowing = [bool]$frame.Overflowing } catch { }
                $record.overflowing = $overflowing

                $Records.Add($record)
            }
        }
        catch {

            Write-Warning "Unable to read text layout of a shape on page $PageNumber."
            Write-Warning $_
        }
    }
}


# ------------------------------------------------------------
# Recursive TEXT export function
#
# Used for the scratch area, whose contents never appear in the
# HTML export. Rather than testing for a specific shape type, this
# asks every shape whether it has a TextFrame with text in it, so
# text boxes, autoshapes with captions and WordArt are all caught.
# ------------------------------------------------------------

function Export-PublisherTextShapes
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Shapes,

        [Parameter(Mandatory = $true)]
        [string]
        $OutputFolder,

        [Parameter(Mandatory = $true)]
        [string]
        $BaseName,

        [Parameter(Mandatory = $true)]
        [int]
        $PageNumber,

        [Parameter(Mandatory = $false)]
        [string]
        $PageLabel = "scratch",

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]
        $TextFilePath = "",

        [Parameter(Mandatory = $true)]
        [ref]
        $TextCounter,

        [Parameter(Mandatory = $true)]
        [ref]
        $PictureCounter,

        [Parameter(Mandatory = $true)]
        [ref]
        $SuccessCounter,

        [Parameter(Mandatory = $true)]
        [ref]
        $SkipCounter,

        [Parameter(Mandatory = $true)]
        [ref]
        $FailCounter,

        [Parameter(Mandatory = $true)]
        [ref]
        $FileHadFailure
    )

    $shapeCount = $Shapes.Count

    for ($shapeNumber = 1;
         $shapeNumber -le $shapeCount;
         $shapeNumber++) {

        $shape = $null

        try {

            $shape = $Shapes.Item($shapeNumber)
            $shapeType = $shape.Type

            # ----------------------------------------------------
            # Recurse into groups
            # ----------------------------------------------------

            if ($shapeType -eq $PB_GROUP) {

                try {

                    Export-PublisherTextShapes `
                        -Shapes $shape.GroupItems `
                        -OutputFolder $OutputFolder `
                        -BaseName $BaseName `
                        -PageNumber $PageNumber `
                        -PageLabel $PageLabel `
                        -TextFilePath $TextFilePath `
                        -TextCounter $TextCounter `
                        -PictureCounter $PictureCounter `
                        -SuccessCounter $SuccessCounter `
                        -SkipCounter $SkipCounter `
                        -FailCounter $FailCounter `
                        -FileHadFailure $FileHadFailure
                }
                catch {

                    Write-Warning `
                        "Unable to inspect grouped shape for text on $PageLabel $PageNumber."

                    Write-Warning $_

                    $FileHadFailure.Value = $true
                }

                continue
            }

            # ----------------------------------------------------
            # Does this shape carry text?
            # Shapes without a text frame raise here; that is normal.
            # ----------------------------------------------------

            $shapeText = $null

            try {
                $shapeText = $shape.TextFrame.TextRange.Text
            }
            catch {
                $shapeText = $null
            }

            if ([string]::IsNullOrWhiteSpace($shapeText)) {
                continue
            }

            $TextCounter.Value++

            # Publisher separates paragraphs with CR; normalise for
            # Windows text editors.

            $normalisedText = $shapeText -replace "`r`n", "`n"
            $normalisedText = $normalisedText -replace "`r", "`n"
            $normalisedText = $normalisedText -replace "`n", [Environment]::NewLine

            $shapeName = ""

            try {
                $shapeName = $shape.Name
            }
            catch {
                $shapeName = ""
            }

            $label = "text block $($TextCounter.Value)"

            if ($shapeName) {
                $label = "$label - $shapeName"
            }

            Write-Output "Found $label on $PageLabel $PageNumber."

            # ----------------------------------------------------
            # Append to the text file
            # ----------------------------------------------------

            if ($TextFilePath) {

                try {

                    $block = @()
                    $block += "==================================================="
                    $block += "$PageLabel $PageNumber - $label (shape type $shapeType)"
                    $block += "==================================================="
                    $block += $normalisedText
                    $block += ""

                    Add-Content `
                        -Path $TextFilePath `
                        -Value ($block -join [Environment]::NewLine) `
                        -Encoding UTF8 `
                        -ErrorAction Stop
                }
                catch {

                    Write-Error "Unable to write text to: $TextFilePath"
                    Write-Error $_

                    $FileHadFailure.Value = $true
                }
            }

            # ----------------------------------------------------
            # Optionally render the text shape to PNG as well
            # ----------------------------------------------------

            if ($EXPORT_SCRATCH_TEXT_PNG) {

                $PictureCounter.Value++

                $pngFileName = "{0}_{1}_{2:D3}_text_{3:D3}.png" -f `
                    $BaseName,
                    $PageLabel,
                    $PageNumber,
                    $PictureCounter.Value

                $pngFilePath = Join-Path `
                    $OutputFolder `
                    $pngFileName

                if (Test-Path $pngFilePath) {

                    Write-Warning "PNG already exists; skipping: $pngFilePath"

                    $SkipCounter.Value++
                }
                else {

                    try {

                        $shape.SaveAsPicture(
                            [string]$pngFilePath,
                            $PNG_RESOLUTION
                        )

                        if (Test-Path $pngFilePath) {

                            Write-Output "Exported PNG: $pngFilePath"

                            $SuccessCounter.Value++
                        }
                        else {

                            Write-Error "Failed to create PNG: $pngFilePath"

                            $FailCounter.Value++
                            $FileHadFailure.Value = $true
                        }

                    }
                    catch {

                        Write-Error `
                            "Error rendering $label on $PageLabel $PageNumber."

                        Write-Error $_

                        $FailCounter.Value++
                        $FileHadFailure.Value = $true
                    }
                }
            }

        }
        catch {

            Write-Warning `
                "Unable to inspect shape $shapeNumber for text on $PageLabel $PageNumber."

            Write-Warning $_

            $FileHadFailure.Value = $true
        }
        finally {

            $shape = $null
        }
    }
}


try {

    $files = Get-ChildItem $Filter -File -Recurse:$Recurse

    if (-not $files) {
        Write-Error "No Publisher files found for the filter: $Filter"
        exit 1
    }

    Write-Output ""
    Write-Output "Starting Publisher conversion..."
    Write-Output ""

    try {
        $app = New-Object -ComObject Publisher.Application
    }
    catch {
        Write-Error "Microsoft Publisher is not installed or accessible."
        exit 1
    }

    # ------------------------------------------------------------
    # Counters
    # ------------------------------------------------------------

    $fileSuccessCount = 0
    $fileFailCount = 0

    $htmlSuccessCount = 0
    $htmlFailCount = 0

    $pngSuccessCount = 0
    $pngSkipCount = 0
    $pngFailCount = 0

    $scratchTextCount = 0

#   $pictureCount = 0

    foreach ($file in $files) {

        if ($file.Extension -ine ".pub") {
            continue
        }

        $doc = $null
        $page = $null

        $fileFullName = $file.FullName
        # Trimmed: Windows can't hold a folder or file name ending in a
        # space, so "My Flyer .pub" would otherwise fail to export.
        $baseName = $file.BaseName.Trim()
        $sourceDirectory = $file.DirectoryName

        # --------------------------------------------------------
        # Create output folder
        # --------------------------------------------------------

        $outputFolder = Join-Path `
            $sourceDirectory `
            $baseName

        Write-Output ""
        Write-Output "=============================================="
        Write-Output "Source: $fileFullName"
        Write-Output "Output: $outputFolder"
        Write-Output "=============================================="

        try {

            if (-not (Test-Path $outputFolder)) {

                New-Item `
                    -ItemType Directory `
                    -Path $outputFolder `
                    -Force `
                    -ErrorAction Stop |
                    Out-Null

                Write-Output "Created output folder."
            }
            else {

                Write-Output "Output folder already exists."
            }

        }
        catch {

            Write-Error "Unable to create output folder: $outputFolder"
            Write-Error $_

            $fileFailCount++

            continue
        }

        # --------------------------------------------------------
        # Publisher 2010 crashes (MSVCR90 invalid parameter) when
        # it saves to a path of 260 characters or more, which loses
        # the rest of the publication. Skip it with a clear error.
        # --------------------------------------------------------

        $longestPath = Join-Path `
            $outputFolder `
            ($baseName + "_scratch_000_image_000.png")

        if ($longestPath.Length -gt 259) {

            Write-Error ("Output paths would be too long for Publisher " +
                "($($longestPath.Length) characters; the limit is 259). " +
                "Move or rename the publication so the folder path is " +
                "$($longestPath.Length - 259) characters shorter: $fileFullName")

            $fileFailCount++

            continue
        }

        # --------------------------------------------------------
        # Open Publisher document
        # --------------------------------------------------------

        try {

            $doc = $app.Open($fileFullName)

        }
        catch {

            Write-Error "Error opening file: $fileFullName"
            Write-Error $_

            $fileFailCount++

            continue
        }

        if (-not $doc) {

            Write-Error "Failed to open file: $fileFullName"

            $fileFailCount++

            continue
        }

        $thisFileHadFailure = $false

        try {

            # ====================================================
            # INDIVIDUAL PICTURE EXPORT
            # ====================================================

            Write-Output ""
            Write-Output "Searching publication for pictures..."

           $pictureCount = 0

            try {

                $pageCount = $doc.Pages.Count

                for (
                    $pageNumber = 1;
                    $pageNumber -le $pageCount;
                    $pageNumber++
                ) {

                    try {

                        $page = $doc.Pages.Item($pageNumber)

                        Export-PublisherPictures `
                            -Shapes $page.Shapes `
                            -OutputFolder $outputFolder `
                            -BaseName $baseName `
                            -PageNumber $pageNumber `
                            -PictureCounter ([ref]$pictureCount) `
                            -SuccessCounter ([ref]$pngSuccessCount) `
                            -SkipCounter ([ref]$pngSkipCount) `
                            -FailCounter ([ref]$pngFailCount) `
                            -FileHadFailure ([ref]$thisFileHadFailure)

                    }
                    catch {

                        Write-Error `
                            "Error examining page $pageNumber."

                        Write-Error $_

                        $thisFileHadFailure = $true
                    }
                    finally {

                        $page = $null
                    }
                }

                # ------------------------------------------------
                # Master pages (backgrounds, logos, etc. placed on
                # a master page are not part of Page.Shapes above)
                # ------------------------------------------------

                $masterPageCount = $doc.MasterPages.Count

                for (
                    $masterPageNumber = 1;
                    $masterPageNumber -le $masterPageCount;
                    $masterPageNumber++
                ) {

                    try {

                        $masterPage = $doc.MasterPages.Item($masterPageNumber)

                        Export-PublisherPictures `
                            -Shapes $masterPage.Shapes `
                            -OutputFolder $outputFolder `
                            -BaseName $baseName `
                            -PageNumber $masterPageNumber `
                            -PageLabel "masterpage" `
                            -PictureCounter ([ref]$pictureCount) `
                            -SuccessCounter ([ref]$pngSuccessCount) `
                            -SkipCounter ([ref]$pngSkipCount) `
                            -FailCounter ([ref]$pngFailCount) `
                            -FileHadFailure ([ref]$thisFileHadFailure)

                    }
                    catch {

                        Write-Error `
                            "Error examining master page $masterPageNumber."

                        Write-Error $_

                        $thisFileHadFailure = $true
                    }
                    finally {

                        $masterPage = $null
                    }
                }

                # ------------------------------------------------
                # Scratch area (the pasteboard beside the pages).
                # Draft / unused objects parked off the canvas live
                # here and are NOT part of Page.Shapes or
                # MasterPage.Shapes, nor of the HTML export, so they
                # must be scanned separately. There is one scratch
                # area per publication.
                # ------------------------------------------------

                try {

                    $scratchArea = $doc.ScratchArea

                    if ($scratchArea) {

                        $scratchShapes = $scratchArea.Shapes

                        if ($scratchShapes -and $scratchShapes.Count -gt 0) {

                            Write-Output ""
                            Write-Output "Scanning scratch area (off-canvas / draft objects)..."

                            Export-PublisherPictures `
                                -Shapes $scratchShapes `
                                -OutputFolder $outputFolder `
                                -BaseName $baseName `
                                -PageNumber 0 `
                                -PageLabel "scratch" `
                                -PictureCounter ([ref]$pictureCount) `
                                -SuccessCounter ([ref]$pngSuccessCount) `
                                -SkipCounter ([ref]$pngSkipCount) `
                                -FailCounter ([ref]$pngFailCount) `
                                -FileHadFailure ([ref]$thisFileHadFailure)

                            # --- text boxes on the pasteboard -------------

                            if ($EXPORT_SCRATCH_TEXT) {

                                $scratchTextPath = Join-Path `
                                    $outputFolder `
                                    ($baseName + "_scratch_text.txt")

                                if (Test-Path $scratchTextPath) {

                                    Write-Warning `
                                        "Scratch text file already exists; not rewriting: $scratchTextPath"

                                    $scratchTextPath = ""
                                }

                                $thisFileTextCount = 0

                                Export-PublisherTextShapes `
                                    -Shapes $scratchShapes `
                                    -OutputFolder $outputFolder `
                                    -BaseName $baseName `
                                    -PageNumber 0 `
                                    -PageLabel "scratch" `
                                    -TextFilePath $scratchTextPath `
                                    -TextCounter ([ref]$thisFileTextCount) `
                                    -PictureCounter ([ref]$pictureCount) `
                                    -SuccessCounter ([ref]$pngSuccessCount) `
                                    -SkipCounter ([ref]$pngSkipCount) `
                                    -FailCounter ([ref]$pngFailCount) `
                                    -FileHadFailure ([ref]$thisFileHadFailure)

                                $scratchTextCount += $thisFileTextCount

                                if ($thisFileTextCount -eq 0) {

                                    Write-Output "No text found in the scratch area."
                                }
                                elseif ($scratchTextPath) {

                                    Write-Output `
                                        "Exported scratch text ($thisFileTextCount block(s)): $scratchTextPath"
                                }
                            }
                        }
                        else {

                            Write-Output "Scratch area contains no shapes."
                        }
                    }

                }
                catch {

                    Write-Warning "Unable to inspect the scratch area for this publication."
                    Write-Warning $_

                    $thisFileHadFailure = $true
                }
                finally {

                    $scratchShapes = $null
                    $scratchArea = $null
                }

                if ($pictureCount -eq 0) {

                    if ($thisFileHadFailure) {

                        Write-Output `
                            "No pictures were identified, but errors occurred while scanning the publication; results may be incomplete."
                    }
                    else {

                        Write-Output `
                            "No picture objects found in publication."
                    }
                }
                else {

                    Write-Output `
                        "Pictures found: $pictureCount"
                }

            }
            catch {

                Write-Error `
                    "Error while scanning publication for pictures."

                Write-Error $_

                $thisFileHadFailure = $true
            }


            # ====================================================
            # PAGE SIZE
            # ====================================================
            # Publisher's HTML sizes its page wrapper to the content, not
            # the page, and sometimes writes a malformed height. Record the
            # real page size, in inches, for stage 2.

            $pageSizePath = Join-Path `
                $outputFolder `
                ($baseName + "_pagesize.txt")

            if (Test-Path $pageSizePath) {

                Write-Output "Page size file already exists; not rewriting: $pageSizePath"
            }
            else {

                try {

                    $inv = [System.Globalization.CultureInfo]::InvariantCulture
                    $pageWidthIn = ([double]$doc.PageSetup.PageWidth / 72).ToString("0.####", $inv)
                    $pageHeightIn = ([double]$doc.PageSetup.PageHeight / 72).ToString("0.####", $inv)

                    Set-Content `
                        -Path $pageSizePath `
                        -Value "$pageWidthIn $pageHeightIn" `
                        -Encoding ASCII `
                        -ErrorAction Stop

                    Write-Output "Page size: $pageWidthIn x $pageHeightIn in"
                }
                catch {

                    Write-Warning "Unable to record the page size."
                    Write-Warning $_
                }
            }


            # ====================================================
            # TEXT LAYOUT
            # ====================================================
            # Each paragraph's real text, with its tabs, and its
            # formatting, for stage 2. Master pages are page 0.

            $textLayoutPath = Join-Path `
                $outputFolder `
                ($baseName + "_text.json")

            if ((Test-Path $textLayoutPath) -and -not $RefreshTextLayout) {

                Write-Output "Text layout file already exists; not rewriting: $textLayoutPath"
            }
            else {

                try {

                    # Bounds snap to screen pixels at the current zoom: about
                    # 1.8pt at the default 37%, 0.2pt at 400%.
                    try { $doc.ActiveView.Zoom = 400 } catch { }

                    $layoutRecords = New-Object System.Collections.Generic.List[object]

                    for ($pageNumber = 1; $pageNumber -le $doc.Pages.Count; $pageNumber++) {

                        Add-PublisherTextLayout `
                            -Shapes $doc.Pages.Item($pageNumber).Shapes `
                            -PageNumber $pageNumber `
                            -Records $layoutRecords
                    }

                    for ($masterNumber = 1; $masterNumber -le $doc.MasterPages.Count; $masterNumber++) {

                        Add-PublisherTextLayout `
                            -Shapes $doc.MasterPages.Item($masterNumber).Shapes `
                            -PageNumber 0 `
                            -Records $layoutRecords
                    }

                    $layout = [ordered]@{
                        version = 1
                        units   = "pt"
                        shapes  = $layoutRecords
                    }

                    $json = ConvertTo-Json -InputObject $layout -Depth 10

                    [System.IO.File]::WriteAllText(
                        [string]$textLayoutPath,
                        $json,
                        (New-Object System.Text.UTF8Encoding $false))

                    Write-Output "Text layout: $($layoutRecords.Count) text frame(s) and table(s)"
                }
                catch {

                    Write-Warning "Unable to record the text layout."
                    Write-Warning $_
                }
            }


            # ====================================================
            # HTML EXPORT
            # ====================================================

            Write-Output ""
            Write-Output "Exporting filtered HTML..."

            $htmlBasePath = Join-Path `
                $outputFolder `
                $baseName

            $htmlFilePath = $htmlBasePath + ".htm"

            if (Test-Path $htmlFilePath) {

                Write-Warning `
                    "HTML already exists; skipping: $htmlFilePath"

            }
            else {

                try {

                    # [string]: Publisher 2010's SaveAs hangs when handed
                    # Join-Path's wrapped PowerShell object instead of a string.
                    $doc.SaveAs(
                        [string]$htmlBasePath,
                        $HTML_FILTERED_FORMAT,
                        $false
                    )

                    if (Test-Path $htmlFilePath) {

                        Write-Output `
                            "Exported HTML: $htmlFilePath"

                        $htmlSuccessCount++
                    }
                    else {

                        Write-Error `
                            "Expected HTML file was not created: $htmlFilePath"

                        $htmlFailCount++
                        $thisFileHadFailure = $true
                    }

                }
                catch {

                    Write-Error `
                        "Error exporting HTML: $fileFullName"

                    Write-Error $_

                    $htmlFailCount++
                    $thisFileHadFailure = $true
                }
            }


            # ====================================================
            # Per-file result
            # ====================================================

            if ($thisFileHadFailure) {

                $fileFailCount++

                Write-Warning `
                    "Completed with one or more errors: $fileFullName"
            }
            else {

                $fileSuccessCount++

                Write-Output `
                    "Successfully processed: $fileFullName"
            }

        }
        finally {

            try {

                if ($doc) {

                    $doc.Close()

                    $doc = $null
                }

            }
            catch {
            }
        }
    }


    # ============================================================
    # FINAL SUMMARY
    # ============================================================

    Write-Output ""
    Write-Output "=============================================="
    Write-Output "Publisher Conversion Complete"
    Write-Output "=============================================="
    Write-Output ""
    Write-Output "Publisher files:"
    Write-Output "  Successful: $fileSuccessCount"
    Write-Output "  Failed:     $fileFailCount"
    Write-Output ""
    Write-Output "HTML exports:"
    Write-Output "  Successful: $htmlSuccessCount"
    Write-Output "  Failed:     $htmlFailCount"
    Write-Output ""
    Write-Output "PNG image exports:"
    Write-Output "  Successful: $pngSuccessCount"
    Write-Output "  Skipped (already existed): $pngSkipCount"
    Write-Output "  Failed:     $pngFailCount"
    Write-Output ""
    Write-Output "Scratch-area text blocks found: $scratchTextCount"
    Write-Output ""
    Write-Output "=============================================="

}
catch {

    Write-Error $_

}
finally {

    try {

        if ($app) {

            $app.Quit()

            $app = $null
        }

    }
    catch {
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}