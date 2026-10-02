# PublisherToPptx
Converts a **Microsoft Publisher `.pub` file** into a **PowerPoint `.pptx`** that imports
cleanly into [Canva](https://www.canva.com/). Each Publisher page becomes a slide laid out
like the original, and the text boxes and pictures arrive in Canva as separate editable
elements instead of one flattened page image.

The conversion runs in two stages:

| Script | Stage |
| --- | --- |
| `PublisherPubToHTMLPNGfilesFinal.ps1` | 1: has Publisher save each `.pub` as filtered HTML and export its pictures as 300 dpi PNGs |
| `PublisherHtmlToPptx.ps1` | 2: finds Python, installs the packages it needs, and runs `pubhtml2pptx.py` |
| `pubhtml2pptx.py` | 2: rebuilds each page of the HTML export as a slide |

## Requirements
- Windows 10/11 with Windows PowerShell 5.1. Stage 2 also runs in PowerShell 7.
- Stage 1: Microsoft Publisher installed, because the script drives it through COM.
- Stage 2: Python 3.9 or later on PATH. The first run installs `python-pptx`,
  `beautifulsoup4`, `lxml` and `Pillow`. The Microsoft Store `python` alias is detected and
  skipped. `pubhtml2pptx.py` must stay in the same folder as `PublisherHtmlToPptx.ps1`.

## Usage
```powershell
# Stage 1: run in the folder holding the .pub files.
# Each MyFlyer.pub gets a MyFlyer\ export folder beside it.
.\PublisherPubToHTMLPNGfilesFinal.ps1 -Filter "*.pub"
.\PublisherPubToHTMLPNGfilesFinal.ps1 -Filter "*.pub" -Recurse

# Stage 2, one export folder: writes MyFlyer\MyFlyer.pptx
.\PublisherHtmlToPptx.ps1 -Path .\MyFlyer

# Stage 2, every export folder in a tree, with the 300 dpi pictures and a report
.\PublisherHtmlToPptx.ps1 -Path C:\Pubs -Recurse -HiRes -Report
```

In Canva, choose **Create a design > Import file**, or drag the `.pptx` onto the Projects page.

### Stage 1 parameters
| Parameter | Default | Meaning |
| --- | --- | --- |
| `-Filter` | (required) | The `.pub` files to convert, such as `"*.pub"` or `"MyFlyer.pub"`. |
| `-Recurse` | | Also convert `.pub` files in subfolders. |

These settings are at the top of the script:

| Setting | Default | Meaning |
| --- | --- | --- |
| `$INCLUDE_OLE_OBJECTS` | `$false` | Also try to export embedded OLE objects as pictures. |
| `$EXPORT_SCRATCH_TEXT` | `$true` | Write text from the scratch area to `<name>_scratch_text.txt`. |
| `$EXPORT_SCRATCH_TEXT_PNG` | `$true` | Also render that scratch-area text to PNG. |
| `$PNG_RESOLUTION` | `3` | Picture export resolution. `3` is 300 dpi. |

### Stage 2 parameters
| Parameter | Default | Meaning |
| --- | --- | --- |
| `-Path` | (required) | An export folder, an `.htm` file, or a tree of export folders with `-Recurse`. |
| `-Recurse` | | Convert every export folder below `-Path`. Any folder holding an `.htm` file counts. |
| `-HiRes` | | Replace Publisher's web-quality images with stage 1's 300 dpi PNGs, keeping the HTML layout. |
| `-NoExtras` | | Leave out the extra slides that carry scratch-area pictures and text. |
| `-Report` | | Write `<name>_conversion_report.txt` beside each `.pptx`. |
| `-OutputFolder` | the export folder | Where the `.pptx` files are written. |
| `-Dpi` | 96 | CSS pixels per inch. Change it only if pages come out the wrong size. |
| `-SkipInstall` | | Skip the Python package check, which makes repeat runs faster. |
| `-Verbose` | | Log every shape as it is placed. |

## What gets converted
| Publisher content | In the `.pptx` |
| --- | --- |
| Page | One slide, at the page's size. When the pages differ in size, the deck uses the largest. |
| Text box | An editable text box, keeping font, size, colour, bold, italic, underline and line breaks. |
| Picture | A picture at the same position and size. With `-HiRes`, stage 1's 300 dpi PNG is used. |
| Pictures inside groups | Exported one by one in stage 1. |
| Master page pictures | Exported in stage 1. |
| Filled box | A filled rectangle. |
| Table | A rectangle for each cell's fill, plus a text box for each cell's text. |
| Scratch area | Pictures and text left off the page never reach the HTML export, so they go on extra slides at the end. Turn these off with `-NoExtras`. |

BMP, TIFF, WMF, EMF, WebP and ICO images are converted to PNG in an `_pptx_converted`
folder inside the export folder. A table becomes separate boxes, not a PowerPoint table.

## Limits
Warnings are counted in the run summary, and each one is listed with `-Report` or `-Verbose`.
- **Fonts:** Canva replaces any font it doesn't have, so check line breaks on every page
  before printing.
- **Image types:** images in other formats (such as Publisher's `.wmz` files) are skipped,
  with a warning.
- **`-HiRes` matching:** a 300 dpi PNG replaces a web image only when the counts on the page
  agree and the aspect ratios match. Otherwise the web image is kept, with a warning.
- **Shapes and WordArt:** only what Publisher's HTML export writes out comes through. A
  shape or WordArt the export doesn't write out is missing from the slide.

## Samples
`samples\` holds six test publications, one folder each. Every folder holds the source
`.pub` and, where stage 1 has been run, its HTML export. Converted `.pptx` output is not kept.

| Sample | Contents |
| --- | --- |
| `bizcard` | `.pub`, plus the stage 1 HTML and `bizcard_files\` |
| `Publication3` | `.pub` and the stage 1 HTML (it has no pictures) |
| `Publication4` | `.pub` only |
| `petlexia publication`, `petlexia publication2`, `petlexia publication3` | `.pub` only |

The second and third petlexia files are named with a space before `.pub`.

To convert the samples that already have an export:

```powershell
.\PublisherHtmlToPptx.ps1 -Path .\samples -Recurse -Report
```
