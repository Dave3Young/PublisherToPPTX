# PublisherToPptx
Converts a **Microsoft Publisher `.pub` file** into a **PowerPoint `.pptx`** that imports
cleanly into [Canva](https://www.canva.com/). Each Publisher page becomes a slide laid out
like the original, and the text boxes and pictures arrive in Canva as separate editable
elements instead of one flattened page image.

The conversion runs in two stages:

| Script | Stage |
| --- | --- |
| `PublisherPubToHTMLPNGfilesFinal.ps1` | 1: has Publisher save each `.pub` as filtered HTML and export its pictures as 300 dpi PNGs, and records the page size and each paragraph's text, formatting and position |
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
| `-Reflow` | | Let text wrap freely in PowerPoint instead of breaking each line where Publisher does. Easier to edit, but lines can break differently from Publisher. |
| `-Report` | | Write `<name>_conversion_report.txt` beside each `.pptx`. |
| `-OutputFolder` | the export folder | Where the `.pptx` files are written. |
| `-Dpi` | 96 | CSS pixels per inch. Change it only if pages come out the wrong size. |
| `-SkipInstall` | | Skip the Python package check, which makes repeat runs faster. |
| `-Verbose` | | Log every shape as it is placed. |

## What stage 1 records beyond the HTML
Publisher's HTML export loses some layout, so stage 1 also asks Publisher for it directly.

| File | Holds | Stage 2 uses it to |
| --- | --- | --- |
| `<name>_pagesize.txt` | The page width and height in inches. | Size the slides. The export sizes its page to the content, sometimes with a broken height. |
| `<name>_text.json` | Every paragraph's real text (with its tabs), alignment, indents, spacing, line-spacing rule, font, tab stops, where Publisher puts it on the page, and where each of its lines starts. | Put tabs and tab stops back exactly, place each paragraph, and each line of inline pictures, where Publisher does, and break lines where Publisher breaks them. |

Stage 2 matches a paragraph in the HTML to `_text.json` by its words and, for repeated text,
by its position. Exports made before these files existed still convert, using estimates.
Stage 1 doesn't rewrite an existing `_text.json`. A `_text.json` written before stage 1
recorded line starts converts with the text wrapping freely; to get Publisher's line breaks,
move the old file aside and run stage 1 again.

## What gets converted
| Publisher content | In the `.pptx` |
| --- | --- |
| Page | One slide, at the page size stage 1 records in `<name>_pagesize.txt`. Without that file, the size comes from the HTML, grown to fit the content. When the pages differ in size, the deck uses the largest. |
| Text box | An editable text box, keeping font, size, colour, bold, italic, underline, all caps, small caps, letter spacing, line breaks, indents (hanging ones too, from `_text.json`), tabs and blank lines. With `_text.json`, each line ends where Publisher ends it, with a hyphen where Publisher hyphenates a word, so text wraps as in Publisher even where it flows around a picture. Use `-Reflow` to let PowerPoint wrap the text instead. Publisher exports some text boxes, filled ones for instance, as a picture of the text. Stage 2 rebuilds these as editable text from the copy Publisher keeps in the HTML's VML markup, with the box's outline behind it: a rectangle, rounded rectangle, ellipse or callout, solid or dashed. |
| Text box overflow | Left out, as Publisher hides it: paragraphs past the bottom of a text box, and inline pictures too wide for it. Each is listed as a warning. |
| Picture | A picture at the same position and size, cropped to the same shape when Publisher crops it to a rounded rectangle or ellipse, and stacked in Publisher's order. With `-HiRes`, stage 1's 300 dpi PNG is used. |
| Pictures inside groups | Exported one by one in stage 1. |
| Master page pictures | Exported in stage 1. |
| Filled box | A filled rectangle. |
| Table | A rectangle for each cell's fill and for each cell border, plus a text box for each cell's text. Cell content is centred vertically unless the cell says otherwise, as in Publisher. |
| Pictures in a line of text | Placed side by side as in Publisher, with the spaces and tabs between them kept. Any words in the line become small text boxes beside the pictures. |
| Scratch area | Pictures and text left off the page never reach the HTML export, so they go on extra slides at the end. Turn these off with `-NoExtras`. |

BMP, TIFF, WMF, EMF, WebP and ICO images are converted to PNG in an `_pptx_converted`
folder inside the export folder. A table becomes separate boxes, not a PowerPoint table.

## Limits
Warnings are counted in the run summary, and each one is listed with `-Report` or `-Verbose`.
- **Fonts:** Canva replaces any font it doesn't have, so check line breaks on every page
  before printing. A font not installed on the converting computer is written as Calibri,
  which is what Publisher draws in its place. Abadi becomes Gill Sans MT, Elephant Pro
  becomes Elephant, and FangSong and KaiTi become SimSun; add others to `FONT_SUBSTITUTES`
  in `pubhtml2pptx.py`.
- **Line breaks:** each line of a paragraph ends in a line break where Publisher ends it.
  Editing the text in Canva or PowerPoint doesn't rewrap those lines, so convert with
  `-Reflow` if the text will be rewritten. PowerPoint sets some fonts up to about 2% wider
  than Publisher, so a line that nearly fills its box is set slightly tighter to keep it on
  one line. A line Publisher indents to clear a picture on its left starts at the box edge.
- **Text wrapping and hyphenation with `-Reflow`, or without `_text.json`:** PowerPoint
  can't wrap text around a picture, and doesn't hyphenate. Where Publisher does either,
  lines break differently, and the paragraphs that follow in the same text box can sit
  higher or lower than in Publisher.
- **Image types:** images in other formats (such as Publisher's `.wmz` files) are skipped,
  with a warning.
- **`-HiRes` matching:** a 300 dpi PNG replaces a web image only when the counts on the page
  agree and the aspect ratios match. Otherwise the web image is kept, with a warning.
- **Tabs without `_text.json`:** Publisher's HTML export writes tabs as runs of spaces.
  A run that ends within two spaces of a default tab stop (every half inch) becomes a tab
  again, in left-aligned paragraphs only, and custom tab stops are lost. Re-run stage 1
  to write `_text.json` and get exact tabs.
- **Text position:** with `_text.json`, text lands within about 2pt of Publisher. PowerPoint
  places a line's extra spacing differently from Publisher, and stage 2 corrects for it
  with one factor for all fonts. PowerPoint rounds spacing given in points to whole points,
  so line spacing is written as a multiple where it can be, and each paragraph's rounded
  space after is made up in the next. Without the file, positions are estimated from the HTML.
- **Shapes and WordArt:** only what Publisher's HTML export writes out comes through. A
  shape or WordArt the export doesn't write out is missing from the slide.
- **Text boxes exported as pictures:** a text box inside a group, a rotated one, WordArt, or
  one whose outline PowerPoint has no preset for (a star or cloud, for instance) stays a
  picture. Without `_text.json`, a rebuilt text box that Publisher aligns to the
  bottom or middle starts at the top instead.

## Samples
`samples\` holds nine test publications, one folder each. Every folder holds the source
`.pub` and, where stage 1 has been run, its HTML export. Converted `.pptx` output is not kept.

| Sample | Contents |
| --- | --- |
| `bizcard` | `.pub`, plus the stage 1 HTML and `bizcard_files\` |
| `Publication3` | `.pub` and the stage 1 HTML (it has no pictures) |
| `Publication4` | `.pub`, plus the stage 1 HTML in `Publication4\` (a blank page) |
| `petlexia publication`, `petlexia publication2`, `petlexia publication3` | `.pub`, plus the stage 1 HTML and `_files\` in a subfolder of the same name |
| `sidebyside-row` | `.pub`, PDF and stage 1 export: three pictures in a row |
| `sidebyside-grid` | `.pub`, PDF and stage 1 export: a 2x2 grid of pictures |
| `sidebyside-imagetext` | `.pub`, PDF and stage 1 export: pictures beside text boxes |

The second and third petlexia files are named with a space before `.pub`. Stage 1 trims
that space, so their exports are named `petlexia publication2` and `petlexia publication3`.

`bizcard` and `Publication3` keep their export next to the `.pub`. The others were exported
later and keep stage 1's own layout, a subfolder named after the publication. `bizcard` and
`Publication3` were exported before stage 1 recorded page sizes and text layout, so they
have no `_pagesize.txt` or `_text.json`.

`make_sidebyside.ps1` builds the three `sidebyside` samples in Publisher, including a PDF
showing how each should look. Its test pictures are lettered squares in `sidebyside-images\`,
so a picture in the wrong place is easy to spot. It doesn't rewrite a sample that exists.

To convert the samples:

```powershell
.\PublisherHtmlToPptx.ps1 -Path .\samples -Recurse -Report
```
