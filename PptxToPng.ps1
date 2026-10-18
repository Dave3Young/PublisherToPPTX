# Downloaded this script? Unblock it first to avoid Windows security errors: right-click it, Properties, tick Unblock, OK. Or run: Unblock-File .\PptxToPng.ps1
<#
Disclaimer: This script is provided "as is"
    with no warranty or guarantee of any kind. By using this script, you accept full responsibility for any
    consequences, including but not limited to data loss, system instability, or unintended results.
    Use at your own risk.

    Saves each slide of a .pptx as a PNG, the way File > Save As > PNG does in PowerPoint
    when "All Slides" is chosen. PowerPoint itself draws the slides, through COM, so fonts,
    line breaks and shapes come out as PowerPoint shows them.

    Each Deck.pptx gets a Deck\ folder holding Slide1.PNG, Slide2.PNG, ... named as
    PowerPoint names them. Files already in that folder with the same names are replaced.

    Stand-alone: it needs only PowerPoint, not the rest of this repository. It is handy for
    checking a converted .pptx against the original publication page by page.

    Requires Microsoft PowerPoint installed. Works in Windows PowerShell 5.1 or PowerShell 7.

    To run:

    for one presentation
    .\PptxToPng.ps1 -Path "C:\Pubs\Example\Example.pptx"

    for every .pptx in a folder, or in a whole tree
    .\PptxToPng.ps1 -Path "C:\Pubs"
    .\PptxToPng.ps1 -Path "C:\Pubs" -Recurse

    at a set resolution instead of PowerPoint's default (150 dpi: a letter page is 1275 x 1650)
    .\PptxToPng.ps1 -Path "C:\Pubs" -Recurse -Dpi 150

    source: www.david-e-young.com
#>
[CmdletBinding()]
param
(
    # A .pptx file, or a folder of them.
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Path,

    # With a folder: also the .pptx files in its subfolders.
    [switch]
    $Recurse,

    # Where each presentation's PNG folder is made. Default: beside the .pptx.
    [string]
    $OutputFolder,

    # Pixels per inch of slide. Leave out for Save As's own size, which follows
    # PowerPoint's export resolution setting (96 dpi unless changed).
    [ValidateRange(1, 1200)]
    [int]
    $Dpi = 0
)

$ErrorActionPreference = "Stop"

# PpSaveAsFileType.ppSaveAsPNG
$ppSaveAsPNG = 18

# MsoTriState
$msoTrue = -1
$msoFalse = 0

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "Path not found: $Path" -ErrorAction Continue
    exit 1
}

$item = Get-Item -LiteralPath $Path

if ($item.PSIsContainer) {
    $files = @(Get-ChildItem -LiteralPath $item.FullName -Filter "*.pptx" -File -Recurse:$Recurse |
        Where-Object { $_.Name -notlike "~$*" } |
        Sort-Object FullName)
}
elseif ($item.Extension -eq ".pptx") {
    $files = @($item)
}
else {
    Write-Error "Not a .pptx file or a folder: $Path" -ErrorAction Continue
    exit 1
}

if ($files.Count -eq 0) {
    Write-Output "No .pptx files found in $Path"
    exit 0
}

if ($OutputFolder) {
    New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
    $OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath
}

# PowerPoint runs as one shared instance: one the user already has open must be
# left running when this script is done.
$wasRunning = [bool](Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue)

$app = $null
$doneCount = 0
$failCount = 0
$slideCount = 0

try {

    $app = New-Object -ComObject PowerPoint.Application

    foreach ($file in $files) {

        $folder = if ($OutputFolder) { $OutputFolder } else { $file.DirectoryName }
        $target = Join-Path $folder $file.BaseName
        $pres = $null

        try {

            # read-only, keep its name, no window
            $pres = $app.Presentations.Open($file.FullName, $msoTrue, $msoFalse, $msoFalse)

            if ($Dpi -gt 0) {
                # Save As has no size of its own; Export writes the same folder
                # and names at the size asked for.
                $width = [int][math]::Round($pres.PageSetup.SlideWidth / 72 * $Dpi)
                $height = [int][math]::Round($pres.PageSetup.SlideHeight / 72 * $Dpi)
                $pres.Export($target, "PNG", $width, $height)
            }
            else {
                $pres.SaveAs($target, $ppSaveAsPNG)
            }

            $slides = $pres.Slides.Count
            $slideCount += $slides
            $doneCount++
            Write-Output "Saved $slides slide(s): $($file.FullName) -> $target"
        }
        catch {
            $failCount++
            Write-Warning "Failed: $($file.FullName): $($_.Exception.Message)"
        }
        finally {
            if ($pres) {
                $pres.Close()
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres)
            }
        }
    }

    Write-Output ""
    Write-Output "Presentations saved: $doneCount    Failed: $failCount    Slides: $slideCount"
}
catch {

    Write-Error $_

}
finally {

    try {

        if ($app) {

            if (-not $wasRunning -and $app.Presentations.Count -eq 0) {
                $app.Quit()
            }

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
            $app = $null
        }

    }
    catch {
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

if ($failCount -gt 0) {
    exit 1
}
