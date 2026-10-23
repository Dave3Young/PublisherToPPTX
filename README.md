# PublisherToPptx
Converts a **Microsoft Publisher `.pub` file** into a **PowerPoint `.pptx`** that imports
cleanly into [Canva](https://www.canva.com/). Each Publisher page becomes a slide laid out
like the original, and the text boxes and pictures arrive in Canva as separate editable
elements instead of one flattened page image.

The conversion runs in two stages:

| Script | Stage |
| --- | --- |
| `PublisherPubToHtmlForPptx.ps1` | 1: has Publisher save each `.pub` as filtered HTML and export its pictures as 300 dpi PNGs, and records the page size and each paragraph's text, formatting and position |
| `PublisherHtmlToPptx.ps1` | 2: finds Python, installs the packages it needs, and runs `pubhtml2pptx.py` |
| `pubhtml2pptx.py` | 2: rebuilds each page of the HTML export as a slide |

Use this repository's stage 1 script, not `PublisherPubToHTMLPNGfilesFinal.ps1` from
[Publisher_Exports](https://github.com/Dave3Young/Publisher_Exports). That script does not
record the page size or the text layout files that stage 2 needs to match Publisher's layout.

## Unblock the scripts after downloading
Windows marks files downloaded from the internet as blocked, and PowerShell may then refuse to
run them with an error such as "is not digitally signed" or "cannot be loaded". Before running
a script for the first time, unblock it:
1. In File Explorer, right-click the script and choose **Properties**.
2. On the **General** tab, tick **Unblock**, then click **Apply** and **OK**.

![The Unblock checkbox on the General tab of a file's Properties](images/unblock.png)

Or unblock every script in the folder at once from PowerShell:

```powershell
Get-ChildItem *.ps1, *.py | Unblock-File
```

If you downloaded the repository as a ZIP, unblock the ZIP the same way before extracting it,
and every file inside comes out unblocked. If there is no **Unblock** checkbox, the file is not
blocked.

## Requirements
- Windows 10/11 with Windows PowerShell 5.1, which comes with Windows.
- Stage 1: Microsoft Publisher installed, because the script drives it through COM.
- Stage 2: Python 3.9 or later on PATH. The first run installs `python-pptx`,
  `beautifulsoup4`, `lxml` and `Pillow`. The Microsoft Store `python` alias is detected and
  skipped. `pubhtml2pptx.py` must stay in the same folder as `PublisherHtmlToPptx.ps1`.

## Usage
Both stages are usually run on a whole tree of folders at once with `-Recurse`. Leave
`-Recurse` off and name one file or folder to convert only that, so that other files in the
same parent folder are left alone.

```powershell
# Stage 1 gives each MyFlyer.pub a MyFlyer\ export folder beside it.

# Every .pub file in the current folder and all of its subfolders
.\PublisherPubToHtmlForPptx.ps1 -Filter "*.pub" -Recurse

# Every .pub file in the current folder only, not its subfolders
.\PublisherPubToHtmlForPptx.ps1 -Filter "*.pub"

# One .pub file only, leaving the other .pub files in the folder alone
.\PublisherPubToHtmlForPptx.ps1 -Filter "MyFlyer.pub"

# One file, or one folder's .pub files, somewhere else
.\PublisherPubToHtmlForPptx.ps1 -Filter "C:\Pubs\Spring\MyFlyer.pub"
.\PublisherPubToHtmlForPptx.ps1 -Filter "C:\Pubs\Spring\*.pub"
```

With `-Recurse`, a file name such as `"MyFlyer.pub"` matches every file of that name in the
subfolders too, so leave `-Recurse` off when converting one file.

```powershell
# Stage 2 writes each export folder's .pptx inside it: MyFlyer\MyFlyer.pptx

# Every export folder in the current folder and all of its subfolders
.\PublisherHtmlToPptx.ps1 -Path . -Recurse

# Every export folder in a tree elsewhere, with the 300 dpi pictures and a report
.\PublisherHtmlToPptx.ps1 -Path C:\Pubs -Recurse -HiRes -Report

# One export folder only, leaving the other export folders alone
.\PublisherHtmlToPptx.ps1 -Path .\MyFlyer

# One export folder somewhere else, named by its folder or its .htm file
.\PublisherHtmlToPptx.ps1 -Path C:\Pubs\Spring\MyFlyer -HiRes
.\PublisherHtmlToPptx.ps1 -Path C:\Pubs\Spring\MyFlyer\MyFlyer.htm
```

In Canva, choose **Create a design > Import file**, or drag the `.pptx` onto the Projects page.

### Stage 1 parameters
| Parameter | Default | Meaning |
| --- | --- | --- |
| `-Filter` | (required) | The `.pub` files to convert, such as `"*.pub"`, `"MyFlyer.pub"` or `"C:\Pubs\Spring\*.pub"`. |
| `-Recurse` | | Also convert matching `.pub` files in subfolders. Leave it off to convert only the files `-Filter` names. |
| `-RefreshTextLayout` | | Rewrite `<name>_text.json` even if it exists. |

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
| `-Recurse` | | Convert every export folder below `-Path`. Any folder holding an `.htm` file counts. Leave it off to convert only the one folder or file `-Path` names. |
| `-HiRes` | | Replace Publisher's web-quality images with stage 1's 300 dpi PNGs, keeping the HTML layout. |
| `-NoExtras` | | Leave out the extra slides that carry scratch-area pictures and text. |
| `-Reflow` | | Let text wrap freely in PowerPoint instead of breaking each line where Publisher does. Easier to edit, but lines can break differently from Publisher. |
| `-Report` | | Write `<name>_conversion_report.txt` beside each `.pptx`. |
| `-OutputFolder` | the export folder | Where the `.pptx` files are written. |
| `-Dpi` | found from the export | CSS pixels per inch. Leave it out: it is worked out from stage 1's measurements, else 96 (see "Publisher versions and display scaling" below). Set it only if pages come out the wrong size. |
| `-SkipInstall` | | Skip the Python package check, which makes repeat runs faster. |
| `-Verbose` | | Log every shape as it is placed. |

## Saving slides as PNG
`PptxToPng.ps1` is a separate tool and needs only PowerPoint. It saves every slide of a
`.pptx` as a PNG, as **File > Save As > PNG** does in PowerPoint: each `Deck.pptx` gets a
`Deck\` folder holding `Slide1.PNG`, `Slide2.PNG` and so on. Use it to compare a converted
`.pptx` with the original publication, page by page.

```powershell
.\PptxToPng.ps1 -Path .\MyFlyer\MyFlyer.pptx
.\PptxToPng.ps1 -Path C:\Pubs -Recurse -Dpi 150
```

| Parameter | Default | Meaning |
| --- | --- | --- |
| `-Path` | (required) | A `.pptx` file, or a folder of them. |
| `-Recurse` | | With a folder, also convert the `.pptx` files in its subfolders. |
| `-OutputFolder` | beside each `.pptx` | Where each presentation's PNG folder is made. |
| `-Dpi` | Save As's own size | Pixels per inch of slide. Save As follows PowerPoint's export resolution setting, 96 dpi unless changed. At 150 dpi a letter page is 1275 x 1650. |

`PubToPng.ps1` does the same for Publisher, and needs only Publisher. It saves every page
of a `.pub` as a PNG, as **File > Save As > PNG** does for the page showing. Each `Flyer.pub`
gets a `Flyer_pages\` folder holding `Page1.png`, `Page2.png` and so on. The folder is not
`Flyer\`, which is where stage 1 writes its export. At the same `-Dpi`, its pages and
`PptxToPng.ps1`'s slides are the same size, so they can be compared pixel for pixel.

```powershell
.\PubToPng.ps1 -Path .\MyFlyer.pub
.\PubToPng.ps1 -Path C:\Pubs -Recurse -Dpi 300
```

| Parameter | Default | Meaning |
| --- | --- | --- |
| `-Path` | (required) | A `.pub` file, or a folder of them. |
| `-Recurse` | | With a folder, also convert the `.pub` files in its subfolders. |
| `-OutputFolder` | beside each `.pub` | Where each publication's PNG folder is made. |
| `-Dpi` | 150 | Pixels per inch of page: 96, 150 or 300, Publisher's own picture resolutions. |

## What stage 1 records beyond the HTML
Publisher's HTML export loses some layout, so stage 1 also asks Publisher for it directly.

| File | Holds | Stage 2 uses it to |
| --- | --- | --- |
| `<name>_pictures.json` | Where each exported picture sits on its page. | Swap each 300 dpi PNG in for the HTML picture in the same place with `-HiRes`, and split a picture grouped with text boxes. |
| `<name>_pagesize.txt` | The page width and height in inches. | Size the slides. The export sizes its page to the content, sometimes with a broken height. |
| `<name>_text.json` | Every paragraph's real text (with its tabs), alignment, indents, spacing, line-spacing rule, font, tab stops, where Publisher puts it on the page, and where each of its lines starts. Also each text box placed in another's text, with where Publisher puts it. | Put tabs and tab stops back exactly, place each paragraph, each line of inline pictures and each inline text box where Publisher does, and break lines where Publisher breaks them. |

Stage 2 matches a paragraph in the HTML to `_text.json` by its words and, for repeated text,
by its position. Exports made before these files existed still convert, using estimates.
Stage 1 doesn't rewrite an existing `_text.json` unless run with `-RefreshTextLayout`.
A `_text.json` written before stage 1 recorded line starts converts with the text wrapping
freely, and one written before it recorded inline text boxes puts them where the HTML export
puts their pictures; run stage 1 with `-RefreshTextLayout` to fix either.

## What gets converted
| Publisher content | In the `.pptx` |
| --- | --- |
| Page | One slide, at the page size stage 1 records in `<name>_pagesize.txt`. Without that file, the size comes from the HTML, grown to fit the content. When the pages differ in size, the deck uses the largest. |
| Text box | An editable text box, keeping font, size, colour, bold, italic, underline, all caps, small caps, letter spacing (Publisher's tracking too, from `_text.json`), line breaks, indents (hanging ones too, from `_text.json`), tabs and blank lines. With `_text.json`, each line ends where Publisher ends it, with a hyphen where Publisher hyphenates a word, so text wraps as in Publisher even where it flows around a picture. Use `-Reflow` to let PowerPoint wrap the text instead. Publisher exports some text boxes, filled ones for instance, as a picture of the text. Stage 2 rebuilds these as editable text from the copy Publisher keeps in the HTML's VML markup, with the box's outline behind it: a rectangle, rounded rectangle, ellipse or callout, solid or dashed. |
| Text box overflow | Left out, as Publisher hides it: paragraphs past the bottom of a text box, and inline pictures too wide for it. Each is listed as a warning. |
| Picture | A picture at the same position and size, cropped to the same shape when Publisher crops it to a rounded rectangle or ellipse, and stacked in Publisher's order. With `-HiRes`, stage 1's 300 dpi PNG is used. |
| Pictures inside groups | Exported one by one in stage 1. Publisher exports a group, a picture with its caption for instance, as one picture of the lot. Where the group holds only pictures and text boxes and stage 1 recorded where each picture sits, stage 2 rebuilds it as stage 1's PNG of each picture and editable text, each in its place. |
| Master page pictures | Exported in stage 1. |
| Filled box | A filled rectangle. |
| Table | A rectangle for each cell's fill and for each cell border, plus a text box for each cell's text. Cell content is centred vertically unless the cell says otherwise, as in Publisher. |
| Pictures in a line of text | Placed side by side as in Publisher, with the spaces and tabs between them kept. Any words in the line become small text boxes beside the pictures. |
| Text box in a line of text | Publisher exports it as a picture of the text. Stage 2 rebuilds it as an editable text box from the copy Publisher keeps in the HTML's VML markup, with its fill and outline, where stage 1 saw Publisher put it, and with Publisher's line breaks. One wider than its frame runs past the frame's edge, as in Publisher. Text boxes inside it are rebuilt the same way. One that Publisher hides, in a text box's overflow for instance, is left out, with a warning. With a `_text.json` written before stage 1 recorded inline text boxes, or none, it goes where the export puts the picture. |
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
  one line.
- **Text wrapping and hyphenation with `-Reflow`, or without `_text.json`:** PowerPoint
  can't wrap text around a picture, and doesn't hyphenate. Where Publisher does either,
  lines break differently, and the paragraphs that follow in the same text box can sit
  higher or lower than in Publisher.
- **Publisher versions and display scaling:** Publisher 2010 writes the HTML's positions at
  96 pixels per inch. Publisher for Microsoft 365, on a display scaled to 125%, writes them
  at 120 and draws every box a few pixels larger. Sizes in points are unaffected. Stage 2
  matches the HTML's shapes to stage 1's measurements to find the scale and the extra size,
  and corrects both. The report notes when it does. Without `_text.json` and `_pictures.json` it can't tell, and assumes 96.
- **Image types:** images in other formats (such as Publisher's `.wmz` files) are skipped,
  with a warning.
- **`-HiRes` matching:** with `_pictures.json`, a 300 dpi PNG replaces the web image in the
  same place and of the same size and shape, and an image with none (WordArt, or a shape
  Publisher exports as a picture) stays web quality. Without it, the PNGs replace the web
  images in order, only when the counts on the page agree and the aspect ratios match;
  otherwise the web images are kept, with a warning. Re-run stage 1 on an existing export
  to add the file.
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
- **Text boxes exported as pictures:** a text box in a group that also holds a shape, a
  rotated one, WordArt, or
  one whose outline PowerPoint has no preset for (a star or cloud, for instance) stays a
  picture. Without `_text.json`, a rebuilt text box that Publisher aligns to the
  bottom or middle starts at the top instead.

## Samples
`samples\` holds four test publications, one folder each. Every folder holds the source
`.pub` and, where stage 1 has been run, its HTML export. Converted `.pptx` output is not kept.

| Sample | Contents |
| --- | --- |
| `bizcard` | `.pub`, plus the stage 1 HTML and `bizcard_files\` in `bizcard\`, and an older export beside the `.pub` |
| `sidebyside-row` | `.pub`, PDF and stage 1 export: three pictures in a row |
| `sidebyside-grid` | `.pub`, PDF and stage 1 export: a 2x2 grid of pictures |
| `sidebyside-imagetext` | `.pub`, PDF and stage 1 export: pictures beside text boxes |

Each export is in stage 1's own layout, a subfolder named after the publication, with its
`_pagesize.txt` and `_text.json`. `bizcard` also keeps an older export next to the `.pub`,
made before stage 1 recorded page sizes and text layout. It shows how stage 2 converts an
export without those files. Converting `samples` with `-Recurse` converts both copies.

`make_sidebyside.ps1` builds the three `sidebyside` samples in Publisher, including a PDF
showing how each should look. Its test pictures are lettered squares in `sidebyside-images\`,
so a picture in the wrong place is easy to spot. It doesn't rewrite a sample that exists.

To convert the samples:

```powershell
.\PublisherHtmlToPptx.ps1 -Path .\samples -Recurse -Report
```
