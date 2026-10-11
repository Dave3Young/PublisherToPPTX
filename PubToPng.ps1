<#
Disclaimer: This script is provided for educational and informational purposes only and is offered "as is"
    with no warranty or guarantee of any kind. By using this script, you accept full responsibility for any
    consequences, including but not limited to data loss, system instability, or unintended results.
    Use at your own risk.

    Saves each page of a Publisher .pub as a PNG, as File > Save As > PNG does in Publisher
    for the page showing, but for every page in turn. Publisher itself draws the pages,
    through COM (Page.SaveAsPicture), so they come out as Publisher shows them.

    Each Flyer.pub gets a Flyer_pages\ folder holding Page1.png, Page2.png, ... Files
    already in that folder with the same names are replaced. (The folder is not Flyer\,
    which is where PublisherPubToHtmlForPptx.ps1 writes its export.)

    Stand-alone: it needs only Publisher, not the rest of this repository. It is handy for
    checking a converted .pptx against the original publication page by page.

    Requires Microsoft Publisher installed. Works in Windows PowerShell 5.1 and PowerShell 7.

    To run:

    for one publication
    .\PubToPng.ps1 -Path "C:\Pubs\Example.pub"

    for every .pub in a folder, or in a whole tree
    .\PubToPng.ps1 -Path "C:\Pubs"
    .\PubToPng.ps1 -Path "C:\Pubs" -Recurse

    at 300 dpi instead of Publisher's default 150 dpi (a letter page is 2550 x 3300)
    .\PubToPng.ps1 -Path "C:\Pubs" -Recurse -Dpi 300

    source: www.david-e-young.com
#>
[CmdletBinding()]
param
(
    # A .pub file, or a folder of them.
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Path,

    # With a folder: also the .pub files in its subfolders.
    [switch]
    $Recurse,

    # Where each publication's PNG folder is made. Default: beside the .pub.
    [string]
    $OutputFolder,

    # Pixels per inch of page: Publisher's own picture resolutions only.
    [ValidateSet(96, 150, 300)]
    [int]
    $Dpi = 150
)

$ErrorActionPreference = "Stop"

# PbPictureResolution: web 96 dpi, desktop print 150 dpi, commercial print 300 dpi
$resolutions = @{ 96 = 1; 150 = 2; 300 = 3 }

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "Path not found: $Path" -ErrorAction Continue
    exit 1
}

$item = Get-Item -LiteralPath $Path

if ($item.PSIsContainer) {
    $files = @(Get-ChildItem -LiteralPath $item.FullName -Filter "*.pub" -File -Recurse:$Recurse |
        Where-Object { $_.Name -notlike "~$*" } |
        Sort-Object FullName)
}
elseif ($item.Extension -eq ".pub") {
    $files = @($item)
}
else {
    Write-Error "Not a .pub file or a folder: $Path" -ErrorAction Continue
    exit 1
}

if ($files.Count -eq 0) {
    Write-Output "No .pub files found in $Path"
    exit 0
}

if ($OutputFolder) {
    New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
    $OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath
}

# Publisher runs as one shared instance: one the user already has open must be
# left running when this script is done.
$wasRunning = [bool](Get-Process -Name "MSPUB" -ErrorAction SilentlyContinue)

$app = $null
$doneCount = 0
$failCount = 0
$pageCount = 0

try {

    $app = New-Object -ComObject Publisher.Application

    foreach ($file in $files) {

        $folder = if ($OutputFolder) { $OutputFolder } else { $file.DirectoryName }
        # a name ending in a space before .pub would make a folder Windows can't open
        $target = Join-Path $folder ($file.BaseName.Trim() + "_pages")
        $doc = $null

        try {

            # read-only, not added to the recent files
            $doc = $app.Open($file.FullName, $true, $false)

            New-Item -ItemType Directory -Force -Path $target | Out-Null

            $pages = $doc.Pages.Count

            for ($i = 1; $i -le $pages; $i++) {
                $png = Join-Path $target ("Page{0}.png" -f $i)
                $doc.Pages.Item($i).SaveAsPicture($png, $resolutions[$Dpi])
            }

            $pageCount += $pages
            $doneCount++
            Write-Output "Saved $pages page(s): $($file.FullName) -> $target"
        }
        catch {
            $failCount++
            Write-Warning "Failed: $($file.FullName): $($_.Exception.Message)"
        }
        finally {
            if ($doc) {
                $doc.Close()
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
            }
        }
    }

    Write-Output ""
    Write-Output "Publications saved: $doneCount    Failed: $failCount    Pages: $pageCount"
}
catch {

    Write-Error $_

}
finally {

    try {

        if ($app) {

            if (-not $wasRunning -and $app.Documents.Count -eq 0) {
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
