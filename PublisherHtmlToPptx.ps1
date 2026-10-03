<#
Disclaimer: This script is provided for educational and informational purposes only and is offered "as is"
    with no warranty or guarantee of any kind. By using this script, you accept full responsibility for any
    consequences, including but not limited to data loss, system instability, or unintended results.
    Use at your own risk.

    Companion to PublisherPubToHTMLPNGfilesFinal.ps1.

    That script leaves one export subfolder per .pub file, containing <BaseName>.htm, the
    <BaseName>_files folder, and the 300 dpi PNGs pulled out of the pages, master pages and
    scratch area. This script turns each of those folders into a .pptx that reproduces the
    original page layout, so it can be imported into Canva with the text boxes and pictures
    arriving as separate editable elements instead of one flattened page image.

    Requires Python 3.9+ on PATH. The first run installs the Python packages it needs.
    Works in Windows PowerShell 5.1 and PowerShell 7.

    To run:

    for one export folder
    .\PublisherHtmlToPptx.ps1 -Path "C:\Pubs\Newsletter"

    for recursive folder processing
    .\PublisherHtmlToPptx.ps1 -Path "C:\Pubs" -Recurse

    to substitute the 300 dpi exported PNGs for Publisher's web-quality images
    .\PublisherHtmlToPptx.ps1 -Path "C:\Pubs" -Recurse -HiRes -Report

    source: www.david-e-young.com
#>
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Path,

    [switch]
    $Recurse,

    # Replace the images Publisher wrote for the web with the 300 dpi PNGs the
    # export script pulled out of the publication, keeping the HTML geometry.
    [switch]
    $HiRes,

    # Skip the extra slides carrying scratch-area art and text.
    [switch]
    $NoExtras,

    # Let text wrap freely instead of breaking each line where Publisher does.
    [switch]
    $Reflow,

    # Write <BaseName>_conversion_report.txt beside each .pptx.
    [switch]
    $Report,

    # Where the .pptx files are written. Default: beside each export folder's HTML.
    [string]
    $OutputFolder,

    # CSS pixels per inch. Leave at 96 unless pages come out the wrong size.
    [double]
    $Dpi = 96,

    # Skip the dependency check (faster on repeat runs).
    [switch]
    $SkipInstall
)

# Note: -Verbose is supplied by [CmdletBinding()] as a common parameter, so it
# must NOT be declared above. It is read back out of $PSBoundParameters below.

$ErrorActionPreference = "Stop"

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$converter = Join-Path $scriptDirectory "pubhtml2pptx.py"

if (-not (Test-Path -LiteralPath $converter)) {
    Write-Error "Cannot find pubhtml2pptx.py next to this script: $converter" -ErrorAction Continue
    exit 1
}

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "Path not found: $Path" -ErrorAction Continue
    exit 1
}

# ------------------------------------------------------------
# Locate Python
#
# $pythonExe is the executable; $pythonPrefix holds any arguments
# that must come first (the py launcher needs -3).
# ------------------------------------------------------------

$pythonExe = $null
$pythonPrefix = @()
$pythonVersion = ""

$previousPreference = $ErrorActionPreference
$ErrorActionPreference = "Continue"

foreach ($candidate in @("python", "python3", "py")) {

    $found = Get-Command $candidate -CommandType Application -ErrorAction SilentlyContinue

    if (-not $found) {
        continue
    }

    if ($found -is [array]) {
        $found = $found[0]
    }

    $candidateExe = $found.Source
    $candidatePrefix = @()

    if ($candidate -eq "py") {
        $candidatePrefix = @("-3")
    }

    # On Windows, "python" often resolves to the Microsoft Store app execution
    # alias: a zero-byte stub that exists on PATH but opens the Store instead of
    # running anything. So ask each candidate for its version and only accept one
    # that actually answers "Python 3.x".

    $versionText = ""

    try {
        $versionText = (& $candidateExe @($candidatePrefix + @("--version")) 2>&1 | Out-String).Trim()
    }
    catch {
        $versionText = ""
    }

    if ($LASTEXITCODE -ne 0 -or $versionText -notmatch "^Python 3\.(\d+)") {
        Write-Verbose "Skipping '$candidate' ($candidateExe): returned '$versionText'"
        continue
    }

    if ([int]$Matches[1] -lt 9) {
        Write-Warning "Skipping '$candidate': $versionText is older than Python 3.9."
        continue
    }

    $pythonExe = $candidateExe
    $pythonPrefix = $candidatePrefix
    $pythonVersion = $versionText

    break
}

$ErrorActionPreference = $previousPreference

if (-not $pythonExe) {
    Write-Error ("A working Python 3.9+ was not found on PATH. Install it from python.org and " +
        "tick 'Add python.exe to PATH'. If 'python' on this machine opens the Microsoft Store, " +
        "turn off the Python app execution aliases in Settings > Apps > App execution aliases.") -ErrorAction Continue
    exit 1
}

Write-Output ""
Write-Output "Using Python: $pythonExe  ($pythonVersion)"

# ------------------------------------------------------------
# Dependencies
#
# Native stderr must not be treated as a terminating error here,
# so the preference is relaxed around the probe.
# ------------------------------------------------------------

if (-not $SkipInstall) {

    Write-Output "Checking Python packages..."

    $previousPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    $null = & $pythonExe @($pythonPrefix + @("-c", "import pptx, bs4, lxml, PIL")) 2>&1
    $probeExitCode = $LASTEXITCODE

    $ErrorActionPreference = $previousPreference

    if ($probeExitCode -ne 0) {

        Write-Output "Installing python-pptx, beautifulsoup4, lxml and Pillow..."

        $ErrorActionPreference = "Continue"

        & $pythonExe @($pythonPrefix + @("-m", "pip", "install", "--quiet", "--upgrade",
            "python-pptx", "beautifulsoup4", "lxml", "Pillow"))

        $installExitCode = $LASTEXITCODE

        $ErrorActionPreference = $previousPreference

        if ($installExitCode -ne 0) {
            Write-Error ("Package installation failed. Try running: " +
                "`"$pythonExe`" -m pip install python-pptx beautifulsoup4 lxml Pillow")
            exit 1
        }
    }
    else {

        Write-Output "Packages already present."
    }
}

# ------------------------------------------------------------
# Build the argument list
# ------------------------------------------------------------

$arguments = @()
$arguments += $pythonPrefix
$arguments += $converter
$arguments += $Path

if ($Recurse)  { $arguments += "--recurse" }
if ($HiRes)    { $arguments += "--hires" }
if ($NoExtras) { $arguments += "--no-extras" }
if ($Reflow)   { $arguments += "--reflow" }
if ($Report)   { $arguments += "--report" }

if ($PSBoundParameters.ContainsKey("Verbose")) {
    $arguments += "--verbose"
}

if ($PSBoundParameters.ContainsKey("Dpi")) {
    $arguments += @("--dpi", $Dpi.ToString([System.Globalization.CultureInfo]::InvariantCulture))
}

if ($OutputFolder) {
    $arguments += @("--output", $OutputFolder)
}

# The call operator quotes each element itself, so paths containing
# spaces need no extra handling here.

$ErrorActionPreference = "Continue"

& $pythonExe @arguments

$exitCode = $LASTEXITCODE

$ErrorActionPreference = "Stop"

if ($exitCode -eq 0) {
    Write-Output "Conversion complete."
}
else {
    Write-Warning "Conversion finished with errors (exit code $exitCode)."
}

exit $exitCode