<#
Builds the side-by-side samples in Publisher:
  sidebyside-row        three images in a row
  sidebyside-grid       a 2x2 grid of images
  sidebyside-imagetext  images beside text boxes
Each gets a .pub and a Publisher PDF showing how it should look. The test
images are written to sidebyside-images\. Existing samples are not rewritten.
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$samples = Split-Path -Parent $MyInvocation.MyCommand.Path
$imageDir = Join-Path $samples 'sidebyside-images'
New-Item -ItemType Directory -Force $imageDir | Out-Null

# Labelled, distinctly coloured test images so placement and order are easy to check.
$specs = @(
    @{ Name = 'A'; Color = [System.Drawing.Color]::FromArgb(214, 69, 65) },
    @{ Name = 'B'; Color = [System.Drawing.Color]::FromArgb(52, 120, 198) },
    @{ Name = 'C'; Color = [System.Drawing.Color]::FromArgb(60, 160, 90) },
    @{ Name = 'D'; Color = [System.Drawing.Color]::FromArgb(230, 160, 40) }
)
$img = @{}
foreach ($s in $specs) {
    $bmp = New-Object System.Drawing.Bitmap 400, 400
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'AntiAlias'
    $g.Clear($s.Color)
    $g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 12), 6, 6, 388, 388)
    $font = New-Object System.Drawing.Font 'Arial', 180, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Alignment = 'Center'; $fmt.LineAlignment = 'Center'
    $g.DrawString($s.Name, $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF 0, 0, 400, 400), $fmt)
    $path = Join-Path $imageDir "img_$($s.Name).png"
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    $img[$s.Name] = [string]$path
}

function Add-Pic($doc, $name, $left, $top, $size) {
    # LinkToFile = msoFalse, SaveWithDocument = msoTrue
    [void]$doc.Pages(1).Shapes.AddPicture($img[$name], 0, -1, $left, $top, $size, $size)
}
function Add-Text($doc, $text, $left, $top, $width, $height, $pt) {
    $tb = $doc.Pages(1).Shapes.AddTextbox(1, $left, $top, $width, $height)
    $tb.TextFrame.TextRange.Text = $text
    $tb.TextFrame.TextRange.Font.Size = $pt
}

$app = New-Object -ComObject Publisher.Application
try {
    $builds = [ordered]@{
        'sidebyside-row' = {
            param($d)
            Add-Text $d 'Three images in a row: A, B, C from left to right.' 72 72 468 40 18
            Add-Pic $d 'A' 72 130 144
            Add-Pic $d 'B' 234 130 144
            Add-Pic $d 'C' 396 130 144
        }
        'sidebyside-grid' = {
            param($d)
            Add-Text $d 'A 2x2 grid: A B on top, C D below.' 72 72 468 40 18
            Add-Pic $d 'A' 126 130 170
            Add-Pic $d 'B' 316 130 170
            Add-Pic $d 'C' 126 320 170
            Add-Pic $d 'D' 316 320 170
        }
        'sidebyside-imagetext' = {
            param($d)
            Add-Pic $d 'A' 72 72 200
            Add-Text $d ("This text box sits to the right of image A, top-aligned with it.`r" +
                         "Image B sits below, with a text box to its left.") 292 72 248 200 16
            Add-Text $d 'This text box sits to the left of image B.' 72 300 248 200 16
            Add-Pic $d 'B' 340 300 200
        }
    }
    foreach ($name in $builds.Keys) {
        $dir = Join-Path $samples $name
        if (Test-Path (Join-Path $dir "$name.pub")) {
            Write-Output "$name.pub already exists; not rewriting"
            continue
        }
        New-Item -ItemType Directory -Force $dir | Out-Null
        $doc = $app.NewDocument()
        & $builds[$name] $doc
        $pub = [string](Join-Path $dir "$name.pub")
        $pdf = [string](Join-Path $dir "$name.pdf")
        $doc.SaveAs($pub)
        $doc.ExportAsFixedFormat(2, $pdf)   # pbFixedFormatTypePDF
        $doc.Close()
        Write-Output "created $pub and $pdf"
    }
}
finally {
    $app.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
