#!/usr/bin/env python3
"""
pubhtml2pptx.py
===============

Converts the filtered-HTML + PNG output of PublisherPubToHTMLPNGfilesFinal.ps1
into a .pptx whose slides reproduce the original Publisher page layout, so the
result can be imported into Canva (or PowerPoint/Google Slides) with the text
boxes and pictures arriving as SEPARATE EDITABLE ELEMENTS rather than one
flattened page image.

Expected input (exactly what the PowerShell export script produces):

    <source folder>\\
        Newsletter.pub
        Newsletter\\                                  <-- the export subfolder
            Newsletter.htm                            <-- filtered HTML
            Newsletter_files\\                        <-- Publisher's web assets
            Newsletter_page_001_image_001.png         <-- 300dpi COM exports
            Newsletter_masterpage_001_image_004.png
            Newsletter_scratch_000_image_007.png
            Newsletter_scratch_000_text_008.png
            Newsletter_scratch_text.txt

Usage
-----
    # one export folder (or point straight at the .htm)
    python pubhtml2pptx.py "C:\\Pubs\\Newsletter"

    # every export folder beneath a tree
    python pubhtml2pptx.py "C:\\Pubs" --recurse

    # use the 300dpi PNGs in place of Publisher's web-quality JPEGs
    python pubhtml2pptx.py "C:\\Pubs" --recurse --hires

Options
-------
    -o, --output PATH   output .pptx (single input) or output folder (--recurse)
    --recurse           find every *_export folder beneath the input folder
    --dpi N             CSS pixels per inch, default 96
    --hires             swap web images for the matching 300dpi COM PNG
    --no-extras         do not append slides for scratch-area PNGs/text
    --reflow            let text wrap freely instead of breaking lines where
                        Publisher does
    --page-size WxH     fallback page size in inches when stage 1 recorded none,
                        default 8.5x11
    --report            write <basename>_conversion_report.txt beside the pptx
    -v, --verbose       per-shape logging

Disclaimer: provided as is, without warranty. Test on copies.
Written for the Publisher retirement toolkit at www.david-e-young.com.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
import glob as globmod
import dataclasses
from dataclasses import dataclass, field
from typing import Optional
from urllib.parse import unquote

try:
    from bs4 import BeautifulSoup, Comment, NavigableString, Tag
    # Base of Comment, Doctype, CData etc.: strings that are not page text.
    # Publisher 2010 puts its VML shape markup in <!--[if gte vml 1]> comments.
    from bs4.element import PreformattedString
except ImportError:
    sys.exit("Missing dependency. Run:  pip install beautifulsoup4 lxml python-pptx Pillow")

try:
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    from pptx.oxml.ns import qn
except ImportError:
    sys.exit("Missing dependency. Run:  pip install python-pptx")

try:
    from PIL import Image
    HAVE_PIL = True
except ImportError:
    HAVE_PIL = False

EMU_PER_INCH = 914400
PPTX_MAX_IN = 56.0

# Outline shapes drawn behind rebuilt text boxes
SHAPE_GEOMS = {
    "rect": MSO_SHAPE.RECTANGLE,
    "roundRect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL,
    "wedgeRoundRectCallout": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
    "wedgeRectCallout": MSO_SHAPE.RECTANGULAR_CALLOUT,
    "wedgeEllipseCallout": MSO_SHAPE.OVAL_CALLOUT,
}
# VML dashstyle -> PowerPoint dash
LINE_DASHES = {
    "dot": MSO_LINE_DASH_STYLE.ROUND_DOT, "1 1": MSO_LINE_DASH_STYLE.SQUARE_DOT,
    "shortdot": MSO_LINE_DASH_STYLE.SQUARE_DOT, "dash": MSO_LINE_DASH_STYLE.DASH,
    "shortdash": MSO_LINE_DASH_STYLE.DASH, "dashdot": MSO_LINE_DASH_STYLE.DASH_DOT,
    "shortdashdot": MSO_LINE_DASH_STYLE.DASH_DOT, "longdash": MSO_LINE_DASH_STYLE.LONG_DASH,
    "longdashdot": MSO_LINE_DASH_STYLE.LONG_DASH_DOT,
    "longdashdotdot": MSO_LINE_DASH_STYLE.DASH_DOT_DOT,
    "shortdashdotdot": MSO_LINE_DASH_STYLE.DASH_DOT_DOT,
}

# ----------------------------------------------------------------------------
# Small CSS engine
# ----------------------------------------------------------------------------

INHERITED_PROPS = {
    "font-family", "font-size", "font-weight", "font-style", "font-variant",
    "color", "text-align", "line-height", "letter-spacing", "text-indent",
    "text-transform", "direction", "white-space",
}

NAMED_COLORS = {
    "black": "000000", "silver": "C0C0C0", "gray": "808080", "grey": "808080",
    "white": "FFFFFF", "maroon": "800000", "red": "FF0000", "purple": "800080",
    "fuchsia": "FF00FF", "magenta": "FF00FF", "green": "008000", "lime": "00FF00",
    "olive": "808000", "yellow": "FFFF00", "navy": "000080", "blue": "0000FF",
    "teal": "008080", "aqua": "00FFFF", "cyan": "00FFFF", "orange": "FFA500",
    "darkblue": "00008B", "darkred": "8B0000", "darkgreen": "006400",
    "lightgray": "D3D3D3", "lightgrey": "D3D3D3", "gold": "FFD700",
}

_COMMENT_RE = re.compile(r"/\*.*?\*/", re.S)
_RULE_RE = re.compile(r"([^{}]+)\{([^{}]*)\}", re.S)
_ATRULE_RE = re.compile(r"@[a-zA-Z-]+[^{;]*(\{(?:[^{}]|\{[^{}]*\})*\}|;)", re.S)
_LEN_RE = re.compile(r"^\s*(-?[\d.]+)\s*(px|pt|in|cm|mm|pc|em|ex|%)?\s*$", re.I)


def parse_decls(text: str) -> dict:
    """Turn 'color:red; font-size:12pt' into a dict."""
    out = {}
    for chunk in text.split(";"):
        if ":" not in chunk:
            continue
        name, _, value = chunk.partition(":")
        name = name.strip().lower()
        value = value.strip()
        if name in ("padding", "margin") and value:
            # 1-4 values: top, right, bottom, left as CSS repeats them
            v = value.split()
            v = (v * 4)[:4] if len(v) == 1 else (v + v)[:4] if len(v) == 2 else (
                v + [v[1]] if len(v) == 3 else v[:4])
            for side, val in zip(("top", "right", "bottom", "left"), v):
                out[f"{name}-{side}"] = val
        elif name and value:
            out[name] = value
    return out


class StyleSheet:
    """Minimal cascade: tag (1) < class (10) < id (100) < inline (1000)."""

    def __init__(self):
        self.rules = []  # (specificity, order, selector_parts, decls)
        self._order = 0

    def add_css(self, css: str) -> None:
        css = _COMMENT_RE.sub(" ", css)
        css = _ATRULE_RE.sub(lambda m: m.group(1) if m.group(1).startswith("{") else " ", css)
        for selectors, body in _RULE_RE.findall(css):
            decls = parse_decls(body)
            if not decls:
                continue
            for sel in selectors.split(","):
                sel = sel.strip().lower()
                if not sel or " " in sel or ">" in sel:
                    # descendant/child selectors: approximate using the rightmost part
                    sel = re.split(r"[\s>+~]+", sel)[-1]
                    if not sel:
                        continue
                parsed = self._parse_simple(sel)
                if parsed is None:
                    continue
                spec, parts = parsed
                self._order += 1
                self.rules.append((spec, self._order, parts, decls))

    @staticmethod
    def _parse_simple(sel: str):
        m = re.match(r"^([a-z0-9]+)?((?:[.#][\w-]+)*)(?::[\w-]+)?$", sel)
        if not m:
            return None
        tag = m.group(1)
        rest = m.group(2) or ""
        classes = re.findall(r"\.([\w-]+)", rest)
        ids = re.findall(r"#([\w-]+)", rest)
        spec = (1 if tag else 0) + 10 * len(classes) + 100 * len(ids)
        return spec, (tag, classes, ids)

    def declarations_for(self, el: Tag) -> dict:
        tag = el.name.lower()
        el_classes = {c.lower() for c in (el.get("class") or [])}
        el_id = (el.get("id") or "").lower()
        matched = []
        for spec, order, parts, decls in self.rules:
            sel_tag, sel_classes, sel_ids = parts
            if sel_tag and sel_tag != tag:
                continue
            if any(c.lower() not in el_classes for c in sel_classes):
                continue
            if any(i.lower() != el_id for i in sel_ids):
                continue
            matched.append((spec, order, decls))
        matched.sort(key=lambda t: (t[0], t[1]))
        out = {}
        for _, _, decls in matched:
            out.update(decls)
        return out


def computed_style(el: Tag, sheet: StyleSheet, inherited: dict) -> dict:
    """Inherited props + matched rules + presentational attrs + inline style."""
    style = {k: v for k, v in inherited.items() if k in INHERITED_PROPS}
    style.update(sheet.declarations_for(el))

    # legacy presentational attributes that Publisher still emits
    if el.name == "font":
        if el.get("color"):
            style["color"] = el["color"]
        if el.get("face"):
            style["font-family"] = el["face"]
        if el.get("size"):
            try:
                style["font-size"] = {1: "8pt", 2: "10pt", 3: "12pt", 4: "14pt",
                                      5: "18pt", 6: "24pt", 7: "36pt"}[int(el["size"])]
            except (ValueError, KeyError):
                pass
    if el.get("align"):
        style["text-align"] = el["align"]
    if el.get("bgcolor"):
        style["background-color"] = el["bgcolor"]
    if el.name in ("b", "strong"):
        style["font-weight"] = "bold"
    if el.name in ("i", "em"):
        style["font-style"] = "italic"
    if el.name == "u":
        style["text-decoration"] = "underline"
    if el.name in ("s", "strike", "del"):
        style["text-decoration"] = "line-through"

    inline = el.get("style")
    if inline:
        style.update(parse_decls(inline))
    return style


def to_px(value, dpi: float, pct_basis: Optional[float] = None) -> Optional[float]:
    """CSS length -> pixels (the unit everything else is derived from)."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    m = _LEN_RE.match(str(value))
    if not m:
        return None
    num = float(m.group(1))
    unit = (m.group(2) or "px").lower()
    if unit == "px":
        return num
    if unit == "pt":
        return num * dpi / 72.0
    if unit == "in":
        return num * dpi
    if unit == "cm":
        return num * dpi / 2.54
    if unit == "mm":
        return num * dpi / 25.4
    if unit == "pc":
        return num * dpi / 6.0
    if unit in ("em", "ex"):
        return num * 12.0 * dpi / 72.0 * (0.5 if unit == "ex" else 1.0)
    if unit == "%":
        return None if pct_basis is None else num * pct_basis / 100.0
    return None


def font_pt(value, dpi: float) -> Optional[float]:
    px = to_px(value, dpi)
    return None if px is None else px * 72.0 / dpi


def parse_color(value) -> Optional[RGBColor]:
    if not value:
        return None
    v = str(value).strip().lower()
    if v in ("transparent", "none", "inherit", "initial", "auto"):
        return None
    v = v.replace("window", "").strip() or "x"
    m = re.match(r"^#([0-9a-f]{3})$", v)
    if m:
        h = m.group(1)
        return RGBColor.from_string("".join(c * 2 for c in h).upper())
    m = re.match(r"^#([0-9a-f]{6})", v)
    if m:
        return RGBColor.from_string(m.group(1).upper())
    m = re.match(r"^rgba?\(([^)]+)\)$", v)
    if m:
        parts = [p.strip() for p in m.group(1).split(",")]
        try:
            if len(parts) >= 4 and float(parts[3]) == 0:
                return None
            r, g, b = (int(round(float(p.rstrip('%')) * (2.55 if p.endswith('%') else 1)))
                       for p in parts[:3])
            return RGBColor(max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))
        except ValueError:
            return None
    if v in NAMED_COLORS:
        return RGBColor.from_string(NAMED_COLORS[v])
    m = re.match(r"^([0-9a-f]{6})$", v)
    if m:
        return RGBColor.from_string(m.group(1).upper())
    return None


def background_color(style: dict) -> Optional[RGBColor]:
    c = parse_color(style.get("background-color"))
    if c is not None:
        return c
    bg = style.get("background")
    if bg and "url(" not in bg:
        for token in bg.split():
            c = parse_color(token)
            if c is not None:
                return c
    return None


# ----------------------------------------------------------------------------
# Layout model
# ----------------------------------------------------------------------------

BLOCK_TAGS = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote",
              "pre", "dd", "dt", "tr", "address", "center"}
SKIP_TAGS = {"script", "style", "head", "meta", "link", "title", "noscript",
             "object", "param", "v:shapetype", "o:shapedefaults"}

ALIGN_MAP = {
    "left": PP_ALIGN.LEFT, "right": PP_ALIGN.RIGHT, "center": PP_ALIGN.CENTER,
    "middle": PP_ALIGN.CENTER, "justify": PP_ALIGN.JUSTIFY,
}


@dataclass
class Run:
    text: str
    font: Optional[str] = None
    size_pt: Optional[float] = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Optional[RGBColor] = None
    caps: Optional[str] = None        # "all" | "small"
    spacing_pt: Optional[float] = None    # letter spacing


@dataclass
class Para:
    runs: list = field(default_factory=list)
    align: Optional[object] = None
    space_after_pt: float = 0.0
    line_spacing: Optional[float] = None
    indent_pt: float = 0.0            # first line, from the left indent; may be negative
    left_indent_pt: float = 0.0
    keep_lines: bool = False          # lines broken where Publisher breaks them
    tab_stops: list = field(default_factory=list)    # [(pos_pt, PbTabAlignmentType)]
    pub: Optional[dict] = None        # stage 1's record for this paragraph
    pub_placed: bool = False          # pub was found at this paragraph's place
    exact_line_pt: Optional[float] = None

    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class Box:
    """One shape destined for the slide, in CSS pixels."""
    kind: str                      # "text" | "image" | "rect"
    x: float
    y: float
    w: float
    h: float
    paras: list = field(default_factory=list)
    src: Optional[str] = None
    fill: Optional[RGBColor] = None
    line: Optional[RGBColor] = None
    line_w_px: float = 0.0
    order: int = 0
    note: str = ""
    anchor: str = "top"            # text: "top" | "middle" | "bottom"
    fixed: bool = False            # at Publisher's own position: never shifted
    geom: str = "rect"             # rect: a SHAPE_GEOMS key
    adj: tuple = ()                # rect: the preset's adjustment values
    dash: Optional[str] = None     # rect: a LINE_DASHES key
    z: Optional[int] = None        # z-index of the positioned shape it came from


# Font files for measuring inline text. ImageFont.truetype finds these in
# C:\Windows\Fonts by file name. Fonts not listed are looked up in the
# registry's list of installed fonts.
FONT_FILES = {
    "arial": "arial.ttf", "verdana": "verdana.ttf", "calibri": "calibri.ttf",
    "cambria": "cambria.ttc", "times new roman": "times.ttf", "georgia": "georgia.ttf",
    "tahoma": "tahoma.ttf", "trebuchet ms": "trebuc.ttf", "segoe ui": "segoeui.ttf",
    "courier new": "cour.ttf", "garamond": "gara.ttf", "century gothic": "gothic.ttf",
}
# Fonts that are not free to install, replaced by an installed font close in
# style and width when the real one is missing. Keys are lower-case.
FONT_SUBSTITUTES = {
    "abadi": "Gill Sans MT", "abadi extra light": "Gill Sans MT",
    "abadi mt": "Gill Sans MT", "abadi mt condensed": "Gill Sans MT Condensed",
    "abadi mt condensed light": "Gill Sans MT Condensed",
    "abadi mt condensed extra bold": "Gill Sans MT Condensed",
    # Chinese fonts: Publisher draws these with SimSun's Latin letters
    "fangsong": "SimSun", "kaiti": "SimSun",
}
MISSING_FONT = "Calibri"
_font_cache: dict = {}
_installed_fonts: Optional[dict] = None


def installed_fonts() -> dict:
    """Lower-case face name -> font file, from the Windows font registry."""
    global _installed_fonts
    if _installed_fonts is None:
        _installed_fonts = {}
        try:
            import winreg
        except ImportError:
            return _installed_fonts
        key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, key_path) as key:
                    i = 0
                    while True:
                        try:
                            name, value, _ = winreg.EnumValue(key, i)
                        except OSError:
                            break
                        i += 1
                        name = re.sub(r"\s*\((TrueType|OpenType)\)\s*$", "", name, flags=re.I)
                        # a collection lists its faces as "A & B"
                        for face in name.split(" & "):
                            _installed_fonts.setdefault(face.strip().lower(), str(value))
            except OSError:
                pass
    return _installed_fonts


_family_installed: dict = {}


def font_installed(family: str) -> bool:
    """True if a face of this family is installed, or if the font list is
    unavailable (not Windows)."""
    key = family.lower()
    if key not in _family_installed:
        inst = installed_fonts()
        _family_installed[key] = (not inst or key in inst or key in FONT_FILES
                                  or any(k.startswith(key + " ") for k in inst))
    return _family_installed[key]


def font_name(family: Optional[str]) -> Optional[str]:
    """The family to use. A missing font becomes its substitute, else Calibri,
    which is what Publisher draws in place of any font it does not have."""
    if not family or font_installed(family):
        return family
    sub = FONT_SUBSTITUTES.get(family.lower())
    return sub if sub and font_installed(sub) else MISSING_FONT

# Collapses HTML whitespace but keeps non-breaking spaces, which Publisher
# uses to space out text and inline pictures.
_HTML_WS_RE = re.compile(r"[ \t\r\n\f]+")


# Publisher's HTML export writes a tab as a run of non-breaking spaces sized
# to roughly reach the next tab stop, then a space. Several typed spaces come
# out the same way, so a run counts as tabs only when it ends close to a stop.
_NBSP_RUN_RE = re.compile("(\xa0{2,} ?)")
# Any whitespace run holding a non-breaking space, for laying out inline lines.
_WS_RUN_RE = re.compile("([ \xa0]*\xa0[ \xa0]*)")
TAB_STOP_PT = 36.0               # Publisher's default tab stops: every half inch
INLINE_PICTURE_GAP_PT = 2.88     # Publisher's default spacing around a picture
# Publisher puts a line's extra spacing (above single spacing) below the text,
# or above it for exact spacing; PowerPoint puts about this share of it above.
PPT_EXTRA_ABOVE = 0.66


# ----------------------------------------------------------------------------
# Stage 1's text layout: each paragraph's real text and formatting
# ----------------------------------------------------------------------------

OBJECT_CHAR = "\ufffc"           # Publisher's placeholder for an inline picture
_PUB_TOKEN_RE = re.compile(r"(\s+|\ufffc)")
TAB_ALIGN = {0: "l", 1: "ctr", 2: "r", 3: "dec"}       # PbTabAlignmentType


def pub_key(text: str) -> str:
    """Matches a paragraph across the HTML and stage 1's layout: its words,
    with every kind of whitespace removed."""
    return re.sub(r"\s+", "", text)


# A bullet or list number the export wrote as text (Symbol, Wingdings or
# Unicode glyphs, "1." or "a)"); Publisher's own text leaves them out
_BULLET_RE = re.compile("^[ \xa0]*(?:[•·▪■●○◦§Ø"
                        "ü–‐-]"
                        r"|\(?(?:\d{1,3}|[a-zA-Z]|[ivxIVX]{1,5})[.)])[ \xa0\t]+")


def _keep_key_chars(runs: list, n: int) -> None:
    """Cuts a paragraph's runs after its first n non-whitespace characters."""
    for i, run in enumerate(runs):
        if n <= 0:
            del runs[i:]
            return
        seen = 0
        for j, ch in enumerate(run.text):
            if not ch.isspace():
                seen += 1
                if seen == n:
                    run.text = run.text[:j + 1]
                    del runs[i + 1:]
                    return
        n -= seen


def _drop_prefix(runs: list, n: int) -> list:
    """Removes the first n characters of a paragraph's runs and returns them
    as runs of their own, styled as they were."""
    removed = []
    for run in runs:
        if n <= 0:
            break
        cut = min(n, len(run.text))
        removed.append(dataclasses.replace(run, text=run.text[:cut]))
        run.text = run.text[cut:]
        n -= cut
    runs[:] = [r for r in runs if r.text] or runs[:1]
    return removed


def load_pub_text(export_dir: str, base: str) -> dict:
    """Stage 1's <base>_text.json: every paragraph, listed by pub_key. A line
    break splits a paragraph in the HTML, so each line is indexed too."""
    path = os.path.join(export_dir, base + "_text.json")
    try:
        with open(path, encoding="utf-8-sig") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        return {}
    index: dict = {}

    def add(par):
        lines = re.split("[\v\n]", str(par.get("text", "")))
        tops = sorted(set(par.get("lineTops") or []))
        # each line on a row of its own: its row's top places it
        rows = len(lines) > 1 and par.get("top") is not None and len(tops) == len(lines)
        # with each line's start, every wrapped line's top is known too
        starts = par.get("lineStarts") or []
        all_tops = par.get("lineTops") or []
        wrapped = (len(lines) > 1 and not rows and par.get("top") is not None
                   and len(starts) == len(all_tops))
        offset = 0                             # the line's start, in UTF-16 units
        for i, line in enumerate(lines):
            size = len(line.encode("utf-16-le")) // 2
            key = pub_key(line)
            if key:
                rec = {**par, "text": line,
                       "firstIndent": par.get("firstIndent", 0) if i == 0 else 0}
                ks = [k for k, s in enumerate(starts) if offset <= s < offset + max(size, 1)]
                if len(lines) > 1:
                    # this line's own line starts, from its start
                    rec["lineStarts"] = [starts[k] - offset for k in ks]
                if wrapped and ks and starts[ks[0]] == offset:
                    end = all_tops[ks[-1] + 1] if ks[-1] + 1 < len(all_tops) else (
                        par["top"] + (par.get("height") or 0.0))
                    first = all_tops[ks[0]]
                    rec.update(top=first, height=end - first,
                               lineTops=[all_tops[k] for k in ks],
                               spaceBefore=par.get("spaceBefore", 0) if i == 0 else 0,
                               spaceAfter=par.get("spaceAfter", 0) if i + 1 == len(lines) else 0)
                elif wrapped:
                    rec["top"] = None
                elif rows:
                    end = tops[i + 1] if i + 1 < len(tops) else (
                        par["top"] + (par.get("height") or 0.0))
                    rec.update(top=tops[i], height=end - tops[i], lineTops=[tops[i]],
                               spaceBefore=par.get("spaceBefore", 0) if i == 0 else 0,
                               spaceAfter=par.get("spaceAfter", 0) if i + 1 == len(lines) else 0)
                elif len(lines) > 1:
                    rec["top"] = None          # a line's own position isn't recorded
                index.setdefault(key, []).append(rec)
            offset += size + 1

    for shape in data.get("shapes", []):
        for par in shape.get("paragraphs") or []:
            add(par)
        for cell in shape.get("cells") or []:
            for par in cell.get("paragraphs") or []:
                add(par)
    return index


def pub_words_and_gaps(text: str):
    """A paragraph's words (pictures count as words) and the whitespace
    before each word, keyed by word index; the key len(words) is trailing."""
    words, gaps = [], {}
    for tok in _PUB_TOKEN_RE.split(text):
        if not tok:
            continue
        if tok.isspace():
            gaps[len(words)] = gaps.get(len(words), "") + tok
        else:
            words.append(tok)
    return words, gaps


def apply_pub_whitespace(runs: list, pub_text: str) -> bool:
    """Swaps the HTML's whitespace for Publisher's real whitespace, tabs
    included, keeping each run's formatting. False if the words differ."""
    html = "".join(r.text for r in runs)
    h_words, _ = pub_words_and_gaps(html)
    p_words, p_gaps = pub_words_and_gaps(pub_text)
    if h_words != p_words:
        return False
    spans, pos, wi = [], 0, 0                  # (start, end, replacement)
    for tok in _PUB_TOKEN_RE.split(html):
        if not tok:
            continue
        if tok.isspace():
            # Publisher has no space before its first word or after its last
            edge = wi == 0 or wi == len(p_words)
            spans.append((pos, pos + len(tok), p_gaps.get(wi, "" if edge else tok)))
        else:
            wi += 1
        pos += len(tok)
    si, cpos = 0, 0
    for run in runs:
        out = []
        for k, ch in enumerate(run.text):
            g = cpos + k
            while si < len(spans) and spans[si][1] <= g:
                si += 1
            if si < len(spans) and spans[si][0] <= g:
                if g == spans[si][0]:
                    out.append(spans[si][2])
                continue
            out.append(ch)
        cpos += len(run.text)
        run.text = "".join(out)
    return True


_SOFT_HYPHENS = "\x1f\xad"


def _utf16_index(text: str, offset: int) -> int:
    """The str index of a UTF-16 offset, as Publisher counts characters."""
    n = 0
    for i, ch in enumerate(text):
        if n >= offset:
            return i
        n += 2 if ord(ch) > 0xFFFF else 1
    return len(text)


def _pub_line_breaks(para, rec: dict) -> bool:
    """Breaks the paragraph's lines where Publisher does, from stage 1's
    lineStarts, so text wraps the same as in Publisher even where it
    hyphenates a word or wraps around a picture. A word Publisher hyphenates
    gets a hyphen. True if the paragraph is now on Publisher's lines (a
    one-line paragraph already is); False, changing nothing, if the text isn't
    Publisher's or stage 1 didn't record its lines."""
    text = str(rec.get("text", ""))
    starts = rec.get("lineStarts") or []
    full = "".join(r.text for r in para.runs)
    if not starts or full.rstrip() != text.rstrip():
        return False
    breaks = {}                                # str index -> (cut from, insert)
    for s in starts:
        i = _utf16_index(text, s)
        if not 0 < i < len(full.rstrip()):
            continue
        j = i
        while j > 0 and full[j - 1] in " \xa0":
            j -= 1                            # the spaces at the wrap go
        if j == 0 or full[j - 1] in "\v\n\t":
            continue
        if j < i:
            breaks[j] = (i, "\v")
        elif full[j - 1] in _SOFT_HYPHENS:
            breaks[j - 1] = (i, "-\v")
        elif full[j - 1].isalnum() and full[i].isalnum():
            breaks[i] = (i, "-\v")            # Publisher hyphenated the word
        else:
            breaks[i] = (i, "\v")
    if not breaks:
        return True
    pos, skip_to = 0, 0
    for run in para.runs:
        out = []
        for k, ch in enumerate(run.text):
            g = pos + k
            if g in breaks:
                out.append(breaks[g][1])
                skip_to = breaks[g][0]
                if skip_to > g:
                    continue
            if g < skip_to:
                continue
            out.append(ch)
        pos += len(run.text)
        run.text = "".join(out)
    return True


# PowerPoint can set text up to about 2% wider than Publisher (it varies with
# the font and size), so a line Publisher just fits would wrap again before its
# break. Lines measuring within this share of the width are tightened to fit,
# by at most KEEP_LINES_MAX_SPC of the font size per character.
KEEP_LINES_SLACK = 0.025
KEEP_LINES_MAX_SPC = 0.05

_SPACE_RUN_RE = re.compile(" {8,}")


def _space_breaks(para, rec: dict) -> None:
    """Some authors end lines with a run of spaces long enough to wrap.
    Publisher drops the spaces at the wrap; PowerPoint carries them onto the
    next line. When the runs split the text into exactly Publisher's lines,
    they become line breaks."""
    tops = set(rec.get("lineTops") or [])
    pieces = _SPACE_RUN_RE.split(str(rec.get("text", "")).rstrip())
    if len(tops) < 2 or len(pieces) != len(tops):
        return
    for run in para.runs:
        run.text = _SPACE_RUN_RE.sub("\v", run.text)
    for run in reversed(para.runs):
        run.text = run.text.rstrip(" \v")
        if run.text:
            break


def next_tab_px(x: float, stops_px: list, default_px: float) -> float:
    """Where a tab at x lands: the next custom stop, else the next default one."""
    for stop in stops_px:
        if stop > x + 0.01:
            return stop
    return (int(x // default_px) + 1) * default_px


def pub_line_multiple(rec: dict) -> Optional[float]:
    """A paragraph's line spacing as a multiple, from PbLineSpacingRule."""
    rule = rec.get("lineRule")
    if rule == 0:
        return 1.0
    if rule == 1:
        return 1.5
    if rule == 2:
        return 2.0
    if rule == 5:
        return float(rec.get("lineSpacing") or 1.0)
    return None                  # exact or at-least spacing: keep the HTML's


def snap_to_tab(pos: float, width: float, nbsp_w: float, stop: float) -> Optional[int]:
    """Tabs a whitespace run at pos stands for, or None if it is just spaces."""
    end = pos + width
    k = round(end / stop)
    if k * stop > pos + 0.5 and abs(k * stop - end) <= 2 * nbsp_w:
        return k - int(pos // stop)
    return None


def text_width_px(text: str, family: Optional[str], size_px: float,
                  bold: bool = False, italic: bool = False) -> float:
    """Width of a run of text, measured with the real font where possible."""
    if not text:
        return 0.0
    if HAVE_PIL:
        name = (font_name(family) or "").lower()
        style = " ".join(s for s, on in (("bold", bold), ("italic", italic)) if on)
        key = (name, style, round(size_px * 4))
        if key not in _font_cache:
            font = None
            fname = ((installed_fonts().get(f"{name} {style}") if style else None)
                     or FONT_FILES.get(name) or installed_fonts().get(name))
            if fname:
                try:
                    from PIL import ImageFont
                    font = ImageFont.truetype(fname, max(key[2], 1))
                except (OSError, ImportError):
                    font = None
            _font_cache[key] = font
        font = _font_cache[key]
        if font is not None:
            return font.getlength(text) / 4.0
    # rough fallback: spaces are narrow, other characters about half an em
    spaces = sum(1 for ch in text if ch in " \xa0")
    return (spaces * 0.28 + (len(text) - spaces) * 0.5) * size_px


def is_positioned(style: dict) -> bool:
    return style.get("position", "static").lower() in ("absolute", "relative", "fixed")


def has_coords(style: dict) -> bool:
    return style.get("left") is not None or style.get("top") is not None


class Converter:
    def __init__(self, args, warn):
        self.dpi = float(args.dpi)
        self.args = args
        self.warn = warn
        # break lines where Publisher does, unless asked to let text reflow
        self.keep_lines = not getattr(args, "reflow", False)
        self.order = 0
        # space after the last block _emit_node placed
        self.trailing = 0.0
        # stage 1's paragraphs, by pub_key; empty for older exports
        self.pub_text: dict = {}
        self.pub_claimed: set = set()     # ids of records already placed

    # -- runs and paragraphs ------------------------------------------------

    def _make_run(self, text: str, style: dict) -> Optional[Run]:
        if not text:
            return None
        fam = style.get("font-family")
        if fam:
            fam = font_name(fam.split(",")[0].strip().strip("'\""))
        weight = str(style.get("font-weight", "")).lower()
        bold = weight in ("bold", "bolder") or (weight.isdigit() and int(weight) >= 600)
        deco = str(style.get("text-decoration", "")).lower()
        return Run(
            text=text,
            font=fam or None,
            size_pt=font_pt(style.get("font-size"), self.dpi),
            bold=bold,
            italic=str(style.get("font-style", "")).lower() in ("italic", "oblique"),
            underline="underline" in deco,
            color=parse_color(style.get("color")),
            caps=("all" if str(style.get("text-transform", "")).lower() == "uppercase"
                  else "small" if "small-caps" in str(style.get("font-variant", "")).lower()
                  else None),
            spacing_pt=(lambda v: v * 72.0 / self.dpi if v else None)(
                to_px(style.get("letter-spacing"), self.dpi)),
        )

    def collect_paras(self, nodes, sheet: StyleSheet, inherited: dict,
                      stop_nodes: set, region=None) -> list:
        """Paragraphs for a list of sibling nodes (or a single element's children)."""
        if isinstance(nodes, Tag):
            nodes = list(nodes.children)
        paras: list = []
        current = Para()

        def flush():
            nonlocal current
            # Publisher writes an empty paragraph as <p>&nbsp;</p>
            if current.runs and (current.text().strip() or "\xa0" in current.text()):
                paras.append(current)
            current = Para()

        def emit(node, style):
            nonlocal current
            if isinstance(node, PreformattedString):
                return
            if isinstance(node, NavigableString):
                raw = str(node)
                if str(style.get("white-space", "")).lower().startswith("pre"):
                    text = raw.replace("\r\n", "\n")
                else:
                    text = _HTML_WS_RE.sub(" ", raw)
                if not text or (not text.strip() and not current.runs and "\xa0" not in text):
                    return
                run = self._make_run(text, style)
                if run:
                    current.runs.append(run)
                return
            if not isinstance(node, Tag):
                return
            name = node.name.lower()
            if name in SKIP_TAGS or id(node) in stop_nodes:
                return
            if name == "br":
                if current.runs:
                    paras.append(current)
                    current = Para()
                return
            if name in ("img", "table"):
                return
            child_style = computed_style(node, sheet, style)
            if name in BLOCK_TAGS:
                flush()
                for c in node.children:
                    emit(c, child_style)
                if current.runs:
                    current.align = ALIGN_MAP.get(str(child_style.get("text-align", "")).lower())
                    lh = to_px(child_style.get("line-height"), self.dpi)
                    fs = to_px(child_style.get("font-size", "12pt"), self.dpi)
                    if lh and fs:
                        current.line_spacing = round(min(max(lh / fs, 0.5), 3.0), 3)
                    mb = to_px(child_style.get("margin-bottom"), self.dpi)
                    current.space_after_pt = (mb or 0) * 72.0 / self.dpi
                    ti = to_px(child_style.get("text-indent"), self.dpi)
                    current.indent_pt = (ti or 0) * 72.0 / self.dpi
                    ml = to_px(child_style.get("margin-left"), self.dpi)
                    current.left_indent_pt = max(ml or 0.0, 0.0) * 72.0 / self.dpi
                flush()
            else:
                for c in node.children:
                    emit(c, child_style)

        for n in nodes:
            emit(n, inherited)
        flush()
        base_align = ALIGN_MAP.get(str(inherited.get("text-align", "")).lower())
        for p in paras:
            if p.align is None:
                p.align = base_align
            rec, placed = self.pub_find(pub_key(p.text()), region)
            # the export writes a bullet as text; Publisher's text has none
            bullet = _BULLET_RE.match(p.text()) if rec is None else None
            bullet_runs = []
            if bullet:
                rec, placed = self.pub_find(pub_key(p.text()[bullet.end():]), region)
                if rec:
                    bullet_runs = _drop_prefix(p.runs, bullet.end())
            if rec is None and region is not None:
                # a paragraph running into the overflow: Publisher's text is
                # only its visible start, so keep just that
                rec, placed = self.pub_find_start(pub_key(p.text()), region)
                if rec:
                    _keep_key_chars(p.runs, len(pub_key(rec["text"])))
            if rec and apply_pub_whitespace(p.runs, rec["text"]):
                p.tab_stops = [(t.get("pos", 0.0), t.get("align", 0))
                               for t in rec.get("tabs") or []]
                p.pub, p.pub_placed = rec, placed
                p.keep_lines = self.keep_lines and _pub_line_breaks(p, rec)
                if not p.keep_lines:
                    _space_breaks(p, rec)
                # Publisher's indents: the HTML leaves out hanging ones
                p.left_indent_pt = max(float(rec.get("leftIndent") or 0.0), 0.0)
                p.indent_pt = max(float(rec.get("firstIndent") or 0.0), -p.left_indent_pt)
            elif p.align in (None, PP_ALIGN.LEFT):
                self._restore_tabs(p)
            if bullet_runs:
                if p.indent_pt < 0:
                    # a hanging bullet: a tab takes the text to the indent
                    bullet_runs[-1].text = bullet_runs[-1].text.rstrip(" \xa0\t") + "\t"
                    bullet_runs = [r for r in bullet_runs if r.text]
                p.runs[:0] = bullet_runs
        text_paras = [p for p in paras if p.text().strip()]
        if not text_paras or all(p.pub_placed for p in text_paras):
            # _pub_place folds the empty paragraphs into the space after
            return text_paras
        # otherwise an empty paragraph is a blank line, except at the end
        while paras and not paras[-1].text().strip():
            paras.pop()
        return paras

    def pub_find(self, key: str, region=None):
        """Stage 1's record for a paragraph: the first unclaimed one inside
        region (x0, y0, x1, y1 in px) if there is one, so repeated text matches
        the right copy, and claims it. Returns (record, found_in_region)."""
        recs = self.pub_text.get(key)
        if not recs:
            return None, False
        if region is not None:
            x0, y0, x1, y1 = region
            tol = 3.0
            for rec in recs:
                if rec.get("top") is None or rec.get("left") is None:
                    continue
                if id(rec) in self.pub_claimed:
                    continue
                rx = rec["left"] * self.dpi / 72.0
                ry = rec["top"] * self.dpi / 72.0
                if x0 - tol <= rx <= x1 + tol and y0 - tol <= ry <= y1 + tol:
                    self.pub_claimed.add(id(rec))
                    return rec, True
        return recs[0], False

    def pub_find_start(self, key: str, region):
        """An unclaimed record inside region whose text is the start of key,
        for a paragraph cut short by its text box's overflow. Claims it."""
        if len(key) < 20:
            return None, False
        x0, y0, x1, y1 = region
        tol = 3.0
        for k, recs in self.pub_text.items():
            if len(k) < 20 or len(k) >= len(key) or not key.startswith(k):
                continue
            for rec in recs:
                if rec.get("top") is None or id(rec) in self.pub_claimed:
                    continue
                rx = rec["left"] * self.dpi / 72.0
                ry = rec["top"] * self.dpi / 72.0
                if x0 - tol <= rx <= x1 + tol and y0 - tol <= ry <= y1 + tol:
                    self.pub_claimed.add(id(rec))
                    return rec, True
        return None, False

    def _pub_place(self, paras: list, bottom: Optional[float] = None):
        """Puts paragraphs where Publisher puts them: sets each one's exact line
        spacing and space after so it starts at Publisher's top. Returns (top,
        height) in px, or None unless every paragraph was found in place.

        Empty paragraphs are left out (their space falls into the space after),
        and so is a text box's overflow, which Publisher hides: paragraphs at
        the end that stage 1 did not see in this box, below text that reaches
        the frame's bottom (px). Both are removed from paras."""
        text = [p for p in paras if p.text().strip()]
        if bottom is not None:
            n = len(text)
            while n and not text[n - 1].pub_placed:
                n -= 1
            if 0 < n < len(text) and text[n - 1].pub_placed:
                last = text[n - 1].pub
                lines = max(len(last.get("lineTops") or []), 1)
                end = (last["top"] + (last.get("height") or 0.0)) * self.dpi / 72.0
                if end >= bottom - 2.5 * (last.get("height") or 0.0) / lines * self.dpi / 72.0:
                    self.warn(f"dropped {len(text) - n} paragraph(s) in a text box's "
                              f"overflow, which Publisher hides: {text[n].text()[:40]!r}")
                    text = text[:n]
        if not text or not all(p.pub_placed for p in text):
            return None
        paras[:] = text
        settings = []
        for i, p in enumerate(paras):
            rec = p.pub
            # older layout files can repeat the last line
            tops = sorted(set(rec.get("lineTops") or [])) or [rec["top"]]
            height = rec.get("height") or 0.0
            # the next paragraph's top; an empty paragraph between them isn't
            # in the HTML, so its space ends up in this one's space after
            nxt = paras[i + 1].pub["top"] if i + 1 < len(paras) else rec["top"] + height
            if len(tops) >= 2:
                line = (tops[-1] - tops[0]) / (len(tops) - 1)
            else:
                line = height - (rec.get("spaceAfter") or 0.0) - (rec.get("spaceBefore") or 0.0)
            if line <= 0:
                return None
            # how far PowerPoint would put this paragraph's text below Publisher's
            multiple = pub_line_multiple(rec) if rec.get("lineRule") in (0, 1, 2, 5) else None
            lift = PPT_EXTRA_ABOVE * line * (1.0 - 1.0 / multiple) if multiple and multiple > 1 else 0.0
            sizes = [r.size_pt for r in p.runs if r.size_pt and r.text]
            if rec.get("lineRule") == 3 and sizes:
                # exact spacing: Publisher puts all the extra above the text.
                # (rec's size is -9999999 when the paragraph mixes sizes.)
                lift = -(1.0 - PPT_EXTRA_ABOVE) * (line - 1.2 * max(sizes))
            settings.append([line, nxt - tops[-1] - line, lift])
        # raise each paragraph by its lift: the box by the first one's, and
        # each later one through the space after the paragraph before it
        for i in range(1, len(settings)):
            settings[i - 1][1] -= settings[i][2] - settings[i - 1][2]
        for p, (line, after, _) in zip(paras, settings):
            p.exact_line_pt = line
            p.space_after_pt = max(after, 0.0)
        first, last = paras[0].pub, paras[-1].pub
        k = self.dpi / 72.0
        top = first["top"] - settings[0][2]
        return top * k, (last["top"] + (last.get("height") or 0.0) - top) * k

    def _restore_tabs(self, para: Para) -> None:
        """Turns the export's tab runs back into tabs, measuring from the frame
        edge as Publisher does. Assumes the runs fall on the first line."""
        stop = TAB_STOP_PT * self.dpi / 72.0
        pos = para.indent_pt * self.dpi / 72.0
        for run in para.runs:
            size_px = (run.size_pt or 12.0) * self.dpi / 72.0
            nbsp_w = text_width_px("\xa0", run.font, size_px)
            out = []
            for i, piece in enumerate(_NBSP_RUN_RE.split(run.text)):
                width = text_width_px(piece, run.font, size_px)
                tabs = snap_to_tab(pos, width, nbsp_w, stop) if i % 2 else None
                if tabs:
                    out.append("\t" * tabs)
                    pos = (int(pos // stop) + tabs) * stop
                else:
                    out.append(piece)
                    pos += width
            run.text = "".join(out)

    def estimate_height(self, paras: list, width_px: float) -> float:
        """Rough laid-out height, used only to stack sibling content sensibly."""
        total = 0.0
        width_px = max(width_px, 20.0)
        for p in paras:
            size_px = max((r.size_pt or 12.0) for r in p.runs) * self.dpi / 72.0
            chars = len(p.text())
            per_line = max(int(width_px / max(size_px * 0.5, 1.0)), 1)
            lines = max(1, -(-chars // per_line))
            lead = (p.line_spacing or 1.18)
            total += lines * size_px * lead + p.space_after_pt * self.dpi / 72.0
        return total

    # -- geometry -----------------------------------------------------------

    def _box_geometry(self, el: Tag, style: dict, origin, cb_w, cb_h):
        left = to_px(style.get("left"), self.dpi, cb_w)
        top = to_px(style.get("top"), self.dpi, cb_h)
        w = to_px(style.get("width"), self.dpi, cb_w)
        h = to_px(style.get("height"), self.dpi, cb_h)

        if w is None and el.get("width"):
            w = to_px(el["width"], self.dpi, cb_w)
        if h is None and el.get("height"):
            h = to_px(el["height"], self.dpi, cb_h)

        right = to_px(style.get("right"), self.dpi, cb_w)
        bottom = to_px(style.get("bottom"), self.dpi, cb_h)
        if left is None and right is not None and w is not None:
            left = cb_w - right - w
        if top is None and bottom is not None and h is not None:
            top = cb_h - bottom - h

        return origin[0] + (left or 0.0), origin[1] + (top or 0.0), w, h

    # -- tables -------------------------------------------------------------

    def _table_boxes(self, table: Tag, sheet: StyleSheet, inherited: dict,
                     x, y, w, h, stop_nodes) -> tuple:
        """Returns (boxes, consumed_height). Cells become individual text boxes."""
        rows = [tr for tr in table.find_all("tr") if tr.find(["td", "th"])]
        if not rows:
            return [], 0.0

        pad = to_px(table.get("cellpadding"), self.dpi)
        pad = 2.0 if pad is None else pad
        grid = [r.find_all(["td", "th"]) for r in rows]
        ncols = max(sum(max(1, int(c.get("colspan", 1) or 1)) for c in cells) for cells in grid)
        ncols = max(ncols, 1)
        w = w or to_px(table.get("width"), self.dpi) or 400.0

        # column widths from the first row with a cell per column, since a
        # cell spanning columns doesn't say how its width divides
        col_w = [w / ncols] * ncols
        for cells in grid:
            if len(cells) != ncols:
                continue
            declared = [to_px(computed_style(c, sheet, inherited).get("width") or c.get("width"),
                              self.dpi, w) for c in cells]
            if all(d for d in declared):
                total = sum(declared)
                col_w = [d * w / total for d in declared]
                break

        # row heights from declared values, else from estimated content
        row_h, cell_paras = [], []
        for ri, cells in enumerate(grid):
            rs = computed_style(rows[ri], sheet, inherited)
            declared_h = to_px(rs.get("height") or rows[ri].get("height"), self.dpi)
            per_row, tallest = [], 0.0
            ci = 0
            for cell in cells:
                span = max(1, int(cell.get("colspan", 1) or 1))
                cw = sum(col_w[ci:ci + span]) if ci < ncols else col_w[-1]
                cs = computed_style(cell, sheet, inherited)
                # Publisher sets the height on the cells, not the row
                cell_h = to_px(cs.get("height") or cell.get("height"), self.dpi)
                if cell_h and int(cell.get("rowspan", 1) or 1) == 1:
                    declared_h = max(declared_h or 0.0, cell_h)
                if cell.name == "th":
                    cs.setdefault("font-weight", "bold")
                paras = self.collect_paras(cell, sheet, cs, stop_nodes)
                est = self.estimate_height(paras, cw - 2 * pad) + 2 * pad
                tallest = max(tallest, est)
                per_row.append((cell, cs, paras, ci, cw))
                ci += span
            cell_paras.append(per_row)
            row_h.append(declared_h or max(tallest, 14.0))

        if h and sum(row_h) < h * 0.92:
            pass  # declared container taller than the content: leave rows compact
        boxes = []
        cy = y
        for ri, per_row in enumerate(cell_paras):
            for cell, cs, paras, ci, cw in per_row:
                cx = x + sum(col_w[:ci])
                fill = background_color(cs)
                if fill is not None:
                    self.order += 1
                    boxes.append(Box(kind="rect", x=cx, y=cy, w=cw, h=row_h[ri],
                                     fill=fill, order=self.order, note="table cell fill"))
                boxes.extend(self._cell_borders(cs, cx, cy, cw, row_h[ri]))
                # HTML centres cell content vertically unless told otherwise,
                # and Publisher's export relies on that
                valign = str(cs.get("vertical-align") or cell.get("valign")
                             or "middle").lower()
                if valign not in ("top", "bottom"):
                    valign = "middle"
                if cell.find(["img", "table"]) is not None:
                    # pictures or a nested table: lay the cell out like a container
                    cell_style = {k: v for k, v in cs.items()
                                  if not k.startswith(("background", "border"))}
                    placed: list = []
                    used = self._emit_node(cell, sheet, cell_style, (cx, cy), cw, row_h[ri],
                                           placed, stop_nodes)
                    pad_b = to_px(cs.get("padding-bottom"), self.dpi) or 0.0
                    # like Publisher, centre without the last paragraph's space after
                    free = max(row_h[ri] - (used - self.trailing) - pad_b, 0.0)
                    dy = {"top": 0.0, "middle": free / 2, "bottom": free}[valign]
                    for b in placed:
                        if not b.fixed:
                            b.y += dy
                    # Publisher hides a text box's overflow: whatever starts
                    # below its bottom, and inline pictures too wide to fit
                    bottom = cy + row_h[ri]
                    kept = [b for b in placed
                            if b.y < bottom - 1.0 and not (b.kind == "image" and b.w > cw + 2.0)]
                    if len(kept) < len(placed):
                        self.warn(f"dropped {len(placed) - len(kept)} shape(s) in a text box's "
                                  f"overflow, which Publisher hides")
                    boxes.extend(kept)
                elif paras:
                    paras = self.collect_paras(cell, sheet, cs, stop_nodes,
                                               region=(cx, cy, cx + cw, cy + row_h[ri]))
                    where = self._pub_place(paras, bottom=cy + row_h[ri])
                    # the cell's own padding, else the table's cellpadding,
                    # plus that of the text box's <div class=shape> inside it
                    pl, pt_, pr, pb = (
                        (pad if v is None else v) + inner for v, inner in zip(
                            (to_px(cs.get("padding-" + side), self.dpi, cw)
                             for side in ("left", "top", "right", "bottom")),
                            self._wrapper_padding(cell, sheet, cs, cw)))
                    self.order += 1
                    if where:
                        top, _ = where
                        boxes.append(Box(kind="text", x=cx + pl, y=top,
                                         w=max(cw - pl - pr, 8.0),
                                         h=max(cy + row_h[ri] - top, 8.0),
                                         paras=paras, order=self.order, note="table cell",
                                         fixed=True))
                    else:
                        boxes.append(Box(kind="text", x=cx + pl, y=cy + pt_,
                                         w=max(cw - pl - pr, 8.0),
                                         h=max(row_h[ri] - pt_ - pb, 8.0),
                                         paras=paras, order=self.order, note="table cell",
                                         anchor=valign))
            cy += row_h[ri]
        return boxes, cy - y

    def _cell_borders(self, style: dict, x, y, w, h) -> list:
        """A table cell's borders, each a thin filled rectangle centred on
        the cell's edge, as Publisher's collapsed borders are."""
        out = []
        for side in ("top", "bottom", "left", "right"):
            val = style.get(f"border-{side}") or style.get("border")
            if not val:
                continue
            toks = str(val).lower().split()
            if any(t in ("none", "hidden") for t in toks):
                continue
            color, width = None, None
            for t in toks:
                c = parse_color(t)
                if c is not None and color is None:
                    color = c
                bw = to_px(t, self.dpi)
                if bw is not None and width is None:
                    width = bw
            width = 1.0 if width is None else width
            if width <= 0:
                continue
            color = color if color is not None else RGBColor(0, 0, 0)
            if side in ("top", "bottom"):
                ey = y if side == "top" else y + h
                rect = (x - width / 2, ey - width / 2, w + width, width)
            else:
                ex = x if side == "left" else x + w
                rect = (ex - width / 2, y - width / 2, width, h + width)
            self.order += 1
            out.append(Box(kind="rect", x=rect[0], y=rect[1], w=rect[2], h=rect[3],
                           fill=color, order=self.order, note="table cell border"))
        return out

    def _wrapper_padding(self, cell: Tag, sheet: StyleSheet, style: dict, cw: float) -> list:
        """Padding (left, top, right, bottom px) of the divs that wrap all of a
        cell's content, one inside the other."""
        total = [0.0, 0.0, 0.0, 0.0]
        el = cell
        while True:
            kids = [c for c in el.children
                    if isinstance(c, Tag) or (isinstance(c, NavigableString)
                                              and not isinstance(c, PreformattedString)
                                              and c.strip())]
            if len(kids) != 1 or not isinstance(kids[0], Tag) or kids[0].name != "div":
                return total
            el = kids[0]
            style = computed_style(el, sheet, style)
            for i, side in enumerate(("left", "top", "right", "bottom")):
                total[i] += to_px(style.get("padding-" + side), self.dpi, cw) or 0.0

    # -- segmentation: split a node's own content into document-order pieces --

    def _segments(self, el: Tag, sheet: StyleSheet, style: dict, stop_nodes: set) -> list:
        """[('text', [nodes], style) | ('img', node, style) | ('table', node, style)]"""
        segs: list = []
        buf: list = []
        buf_style = style

        def flush_buf():
            nonlocal buf
            if buf:
                segs.append(("text", buf, buf_style))
                buf = []

        def scan(node, cur_style):
            nonlocal buf_style
            for child in node.children:
                if isinstance(child, PreformattedString):
                    continue
                if isinstance(child, NavigableString):
                    if str(child).strip():
                        if not buf:
                            buf_style = cur_style
                        buf.append(child)
                    continue
                if not isinstance(child, Tag):
                    continue
                name = child.name.lower()
                if name in SKIP_TAGS or id(child) in stop_nodes:
                    continue
                cs = computed_style(child, sheet, cur_style)
                if name == "img":
                    flush_buf()
                    segs.append(("img", child, cs))
                    continue
                if name == "table":
                    flush_buf()
                    segs.append(("table", child, cs))
                    continue
                if (name in BLOCK_TAGS and child.find("img") is not None
                        and child.find(["table", *BLOCK_TAGS]) is None):
                    # a paragraph mixing pictures and text: lay it out as a line
                    flush_buf()
                    segs.append(("line", child, cs))
                    continue
                if child.find(["img", "table"]) is not None:
                    scan(child, cs)
                    continue
                if not buf:
                    buf_style = cur_style
                buf.append(child)
        scan(el, style)
        flush_buf()
        return segs

    # -- inline lines: pictures and text side by side -------------------------

    def _inline_items(self, node: Tag, sheet: StyleSheet, style: dict, out: list):
        for child in node.children:
            if isinstance(child, PreformattedString):
                continue
            if isinstance(child, NavigableString):
                text = _HTML_WS_RE.sub(" ", str(child))
                if text:
                    out.append(("text", text, style))
                continue
            if not isinstance(child, Tag):
                continue
            name = child.name.lower()
            if name in SKIP_TAGS:
                continue
            cs = computed_style(child, sheet, style)
            if name in ("img", "br"):
                out.append((name, child, cs))
                continue
            self._inline_items(child, sheet, cs, out)

    def _line_boxes(self, para: Tag, sheet: StyleSheet, style: dict,
                    x, y, w, h, boxes: list, region=None) -> float:
        """Places a paragraph's pictures and text left to right, wrapping at the
        frame edge, as Publisher lays out a line with inline pictures. Spaces
        keep their width, and the export's tab runs snap to tab stops.
        Returns the height used."""
        items: list = []
        self._inline_items(para, sheet, style, items)
        gap = INLINE_PICTURE_GAP_PT * self.dpi / 72.0

        measured = []          # [kind, payload, style, width, height]
        after_space = True
        for kind, payload, st in items:
            if kind == "img":
                _, _, iw, ih = self._box_geometry(payload, st, (0.0, 0.0), w, h)
                iw = iw or 50.0
                ih = ih or iw
                measured.append(["img", payload, st, iw + 2 * gap, ih + 2 * gap])
                after_space = False
            elif kind == "br":
                measured.append(["br", None, st, 0.0, 0.0])
                after_space = True
            else:
                text = payload.lstrip(" ") if after_space else payload
                if not text:
                    continue
                after_space = text.endswith(" ")
                size_px = to_px(st.get("font-size", "12pt"), self.dpi) or 16.0
                fam = (st.get("font-family") or "").split(",")[0].strip().strip("'\"")
                for i, piece in enumerate(_WS_RUN_RE.split(text)):
                    if piece:
                        measured.append(["space" if i % 2 else "text", piece, st,
                                         text_width_px(piece, fam, size_px), size_px * 1.2])

        # Publisher's own text for this line: real tabs and tab stops
        rec, placed = self.pub_find(pub_key("".join(
            OBJECT_CHAR if m[0] == "img" else (m[1] or "") for m in measured)), region)
        pub_gaps, stops_px = None, []
        if rec:
            h_words = []
            for m in measured:
                if m[0] == "img":
                    h_words.append(OBJECT_CHAR)
                elif m[0] == "text":
                    h_words.extend(m[1].split())
                elif m[0] == "space":
                    m.append(len(h_words))         # index of the word it precedes
            p_words, p_gaps = pub_words_and_gaps(rec["text"])
            if p_words == h_words:
                pub_gaps = p_gaps
            else:
                placed = False
                stops_px = sorted(t.get("pos", 0.0) * self.dpi / 72.0
                                  for t in rec.get("tabs") or [])

        align = str(style.get("text-align", "left")).lower()
        left_aligned = align not in ("center", "middle", "right")
        stop = TAB_STOP_PT * self.dpi / 72.0
        indent = to_px(style.get("text-indent"), self.dpi) or 0.0
        lines, cur, pos = [], [], indent
        for m in measured:
            if m[0] == "br":
                lines.append(cur)
                cur, pos = [], 0.0
                continue
            if m[0] == "space":
                size_px = m[4] / 1.2
                fam = (m[2].get("font-family") or "").split(",")[0].strip().strip("'\"")
                if pub_gaps is not None and m[5] in pub_gaps:
                    # Publisher's real whitespace: tabs go to their stops
                    space_w = text_width_px(" ", fam, size_px)
                    end = pos
                    for ch in pub_gaps[m[5]]:
                        if ch == "\t":
                            end = next_tab_px(end, stops_px, stop)
                        elif ch not in "\v\n\r":
                            end += space_w
                    m[3] = end - pos
                elif left_aligned and m[1].count("\xa0") >= 2:
                    tabs = snap_to_tab(pos, m[3], text_width_px("\xa0", fam, size_px), stop)
                    if tabs:
                        m[3] = (int(pos // stop) + tabs) * stop - pos
            if cur and pos + m[3] > w + 0.5:
                lines.append(cur)
                cur, pos = [], 0.0
                if m[0] == "space":
                    continue
            cur.append(m)
            pos += m[3]
        lines.append(cur)

        # Without stage 1's position, share the line's extra spacing above and
        # below it. With it, Publisher's line top is known and the extra
        # spacing falls below; the tallest picture sits its gap below the top.
        spacing = (pub_line_multiple(rec) if rec and pub_gaps is not None else None) \
            or self._line_multiple(style)
        # where Publisher puts the line, when stage 1 recorded it
        cy = rec["top"] * self.dpi / 72.0 if placed else y
        for li, line in enumerate(lines):
            if not line:
                continue
            lead = indent if li == 0 else 0.0
            total = sum(m[3] for m in line)
            if align in ("center", "middle"):
                cx = x + lead + max(w - lead - total, 0.0) / 2
            elif align == "right":
                cx = x + max(w - total, lead)
            else:
                cx = x + lead
            content_h = max(m[4] for m in line)
            extra = content_h * (spacing - 1.0)
            top = cy if placed else cy + extra / 2
            for kind, payload, st, iw, ih, *_ in line:
                if kind == "img":
                    src = unquote(payload.get("src") or "")
                    if src:
                        self.order += 1
                        boxes.append(Box(kind="image", x=cx + gap,
                                         y=top + content_h - ih + gap,
                                         w=iw - 2 * gap, h=ih - 2 * gap,
                                         src=src, order=self.order,
                                         note=payload.get("alt") or "", fixed=placed))
                elif kind == "text" and payload.strip():
                    run = self._make_run(payload.strip(), st)
                    size_px = ih / 1.2
                    fam = run.font if run else None
                    skip = text_width_px(payload[:len(payload) - len(payload.lstrip())],
                                         fam, size_px)
                    vis = text_width_px(payload.strip(), fam, size_px)
                    self.order += 1
                    # slack so PowerPoint's own metrics don't wrap the text
                    boxes.append(Box(kind="text", x=cx + skip, y=top + content_h - ih,
                                     w=vis * 1.15 + 4.0, h=ih,
                                     paras=[Para(runs=[run])], order=self.order,
                                     note="inline text", fixed=placed))
                cx += iw
            cy += content_h + extra

        mb = to_px(style.get("margin-bottom"), self.dpi) or 0.0
        self.trailing = mb
        return cy - y + mb

    def _line_multiple(self, style: dict) -> float:
        """A paragraph's line-height as a multiple, such as 1.14 for 114%."""
        value = str(style.get("line-height") or "").strip()
        if value.endswith("%"):
            try:
                ratio = float(value[:-1]) / 100.0
            except ValueError:
                ratio = 1.0
        else:
            lh = to_px(value, self.dpi)
            fs = to_px(style.get("font-size", "12pt"), self.dpi)
            ratio = lh / fs if lh and fs else 1.0
        return min(max(ratio, 1.0), 3.0)

    # -- tree walk ----------------------------------------------------------

    def _shape_children(self, el: Tag, sheet: StyleSheet) -> list:
        out = []
        for child in el.find_all(True):
            if child.name in SKIP_TAGS:
                continue
            st = computed_style(child, sheet, {})
            if not (is_positioned(st) and (has_coords(st) or st.get("width"))):
                continue
            parent, nested = child.parent, False
            while parent is not None and parent is not el:
                if isinstance(parent, Tag):
                    ps = computed_style(parent, sheet, {})
                    if is_positioned(ps) and (has_coords(ps) or ps.get("width")):
                        nested = True
                        break
                parent = parent.parent
            if not nested:
                out.append(child)
        return out

    def walk(self, el: Tag, sheet: StyleSheet, inherited: dict,
             origin, cb_w, cb_h, boxes: list, depth=0):
        shape_children = self._shape_children(el, sheet)
        stop_nodes = {id(c) for c in shape_children}
        own_style = inherited if el.name == "body" else computed_style(el, sheet, inherited)

        self._emit_node(el, sheet, own_style, origin, cb_w, cb_h, boxes, stop_nodes)

        for child in shape_children:
            cstyle = computed_style(child, sheet, own_style)
            x, y, w, h = self._box_geometry(child, cstyle, origin, cb_w, cb_h)
            new_w = w if w is not None else max(cb_w - (x - origin[0]), 1.0)
            new_h = h if h is not None else max(cb_h - (y - origin[1]), 1.0)
            first = len(boxes)
            if child.name.lower() == "img":
                src = unquote(child.get("src") or "")
                if src:
                    self.order += 1
                    adj = tuple(float(v) for v in str(cstyle.get("x-adj") or "").split())
                    boxes.append(Box(kind="image", x=x, y=y, w=new_w, h=new_h,
                                     src=src, order=self.order,
                                     note=child.get("alt") or "",
                                     geom=cstyle.get("x-geom") or "rect", adj=adj))
            else:
                self.walk(child, sheet, cstyle, (x, y), new_w, new_h, boxes, depth + 1)
            # Publisher's stacking order is the z-index, not the order in the file
            try:
                z = int(str(cstyle.get("z-index") or "").strip())
            except ValueError:
                continue
            for b in boxes[first:]:
                if b.z is None:
                    b.z = z

    def _emit_node(self, el: Tag, sheet: StyleSheet, style: dict,
                   origin, cb_w, cb_h, boxes: list, stop_nodes: set) -> float:
        """Adds the node's boxes and returns the height its content used."""
        x, y = origin
        w = cb_w if cb_w else 200.0
        h = cb_h if cb_h else 40.0

        # background / border of the container itself
        fill = background_color(style)
        border, border_w = None, 0.0
        for prop in ("border", "border-top", "border-left", "border-color", "border-width"):
            val = style.get(prop)
            if not val:
                continue
            for tok in str(val).split():
                c = parse_color(tok)
                if c is not None and border is None:
                    border = c
                bw = to_px(tok, self.dpi)
                if bw and bw > 0 and border_w == 0.0:
                    border_w = bw
        if border is not None and border_w == 0.0:
            border_w = 1.0
        # a rebuilt VML text box or picture says which outline shape it had
        adj = tuple(float(v) for v in str(style.get("x-adj") or "").split())
        if (fill is not None or border is not None) and w and h:
            self.order += 1
            boxes.append(Box(kind="rect", x=x, y=y, w=w, h=h, fill=fill,
                             line=border, line_w_px=border_w, order=self.order,
                             geom=style.get("x-geom") or "rect", adj=adj,
                             dash=style.get("x-dash")))

        if el.name.lower() == "img":
            src = unquote(el.get("src") or "")
            if src:
                self.order += 1
                boxes.append(Box(kind="image", x=x, y=y, w=w, h=h, src=src,
                                 order=self.order, note=el.get("alt") or "",
                                 geom=style.get("x-geom") or "rect", adj=adj))
            return h

        if el.name.lower() == "table":
            # a positioned table: keep its cell grid instead of flowing the cells
            tb, used = self._table_boxes(el, sheet, style, x, y, w, h, stop_nodes)
            boxes.extend(tb)
            return used

        pad_l = to_px(style.get("padding-left"), self.dpi, w) or 0.0
        pad_t = to_px(style.get("padding-top"), self.dpi, h) or 0.0
        pad_r = to_px(style.get("padding-right"), self.dpi, w) or 0.0
        content_x = x + pad_l
        content_w = max(w - pad_l - pad_r, 10.0)
        cursor = y + pad_t

        segs = self._segments(el, sheet, style, stop_nodes)
        trailing = 0.0
        text_segs = [s for s in segs if s[0] == "text"]
        only_text = len(segs) == len(text_segs)

        for kind, payload, seg_style in segs:
            if kind == "text":
                paras = self.collect_paras(payload, sheet, seg_style, stop_nodes,
                                           region=(x, y, x + w, y + h))
                if not paras:
                    continue
                where = self._pub_place(paras, bottom=y + h if cb_h else None)
                if where:
                    top, used = where
                    self.order += 1
                    boxes.append(Box(kind="text", x=content_x, y=top, w=content_w,
                                     h=max(used, y + h - top if only_text else used, 8.0),
                                     paras=paras, order=self.order, fixed=True))
                    cursor = top + used
                    trailing = 0.0
                    continue
                est = self.estimate_height(paras, content_w)
                if only_text and len(text_segs) == 1:
                    box_h = max(h - pad_t, 8.0)      # honour the declared frame
                else:
                    box_h = max(est, 8.0)
                self.order += 1
                boxes.append(Box(kind="text", x=content_x, y=cursor,
                                 w=content_w, h=box_h, paras=paras, order=self.order))
                cursor += est
                trailing = paras[-1].space_after_pt * self.dpi / 72.0
                continue

            if kind == "img":
                ix, iy, iw, ih = self._box_geometry(payload, seg_style,
                                                    (content_x, cursor), content_w, h)
                if not has_coords(seg_style):
                    ix, iy = content_x, cursor
                iw = iw or content_w
                ih = ih or (iw * 0.75)
                src = unquote(payload.get("src") or "")
                if src:
                    self.order += 1
                    boxes.append(Box(kind="image", x=ix, y=iy, w=iw, h=ih, src=src,
                                     order=self.order, note=payload.get("alt") or ""))
                if not has_coords(seg_style):
                    cursor += ih
                    trailing = 0.0
                continue

            if kind == "line":
                cursor += self._line_boxes(payload, sheet, seg_style, content_x, cursor,
                                           content_w, h, boxes, region=(x, y, x + w, y + h))
                trailing = self.trailing
                continue

            # table
            tx, ty, tw, th = self._box_geometry(payload, seg_style,
                                                (content_x, cursor), content_w, h)
            if not has_coords(seg_style):
                tx, ty = content_x, cursor
            tb, used = self._table_boxes(payload, sheet, seg_style, tx, ty,
                                         tw or content_w, th, stop_nodes)
            boxes.extend(tb)
            if not has_coords(seg_style):
                cursor += used
                trailing = 0.0
        self.trailing = trailing
        return cursor - y


# ----------------------------------------------------------------------------
# Page detection
# ----------------------------------------------------------------------------

def _vml_pt(value, default=0.0) -> float:
    """A VML length in pt: '34.84pt', '0' or '8.68mm'."""
    m = re.match(r"\s*(-?[\d.]+)\s*(pt|in|px|mm|cm)?", str(value or ""))
    if not m:
        return default
    v = float(m.group(1))
    return {"in": v * 72.0, "px": v * 0.75, "mm": v * 72.0 / 25.4,
            "cm": v * 72.0 / 2.54}.get(m.group(2), v)


def _vml_fraction(value, default: float) -> float:
    """A VML fraction: '0.25', '25%' or '16384f' (65536ths)."""
    v = str(value or "").strip()
    try:
        if v.endswith("f"):
            return float(v[:-1]) / 65536.0
        if v.endswith("%"):
            return float(v[:-1]) / 100.0
        return float(v) if v else default
    except ValueError:
        return default


# VML shape type (#_x0000_tNN) -> outline: rectangles, rounded ones, ellipses
# and their callouts. Other shapes keep Publisher's picture.
_VML_TYPES = {"202": "rect", "1": "rect", "2": "roundRect", "3": "ellipse",
              "176": "roundRect", "61": "wedgeRectCallout",
              "62": "wedgeRoundRectCallout", "63": "wedgeEllipseCallout"}


def _vml_outline(shape: Tag) -> Optional[tuple]:
    """(geom, adjustments) for a VML shape's outline, or None if PowerPoint
    has no matching preset."""
    if shape.name in ("v:rect", "v:oval"):
        return ("rect" if shape.name == "v:rect" else "ellipse"), ()
    if shape.name == "v:roundrect":
        # arcsize: the corner radius as a share of the shorter side
        return "roundRect", (_vml_fraction(shape.get("arcsize"), 0.2),)
    m = re.match(r"#_x0000_t(\d+)$", str(shape.get("type") or ""))
    geom = _VML_TYPES.get(m.group(1)) if m else None
    if geom is None:
        return None
    adj = [float(v) for v in re.findall(r"-?\d+(?:\.\d+)?", str(shape.get("adj") or ""))]
    if m.group(1) == "2":
        return geom, ((adj[0] if adj else 3600.0) / 21600.0,)
    if m.group(1) == "176":
        return geom, (1 / 6,)
    if geom.endswith("Callout"):
        tx, ty = (adj + [1350.0, 25920.0][len(adj):])[:2]
        if 0 <= tx <= 21600 and 0 <= ty <= 21600:
            # the pointer ends inside the shape, so none shows
            return {"wedgeRectCallout": "rect", "wedgeEllipseCallout": "ellipse"}.get(
                geom, "roundRect"), ((1 / 6,) if geom == "wedgeRoundRectCallout" else ())
        return geom, (tx / 21600.0 - 0.5, ty / 21600.0 - 0.5) + (
            (1 / 6,) if geom == "wedgeRoundRectCallout" else ())
    return geom, ()


def _vml_text_area(shape: Tag, shapetype: Optional[Tag], geom: str, adj: tuple,
                   width: float, height: float) -> tuple:
    """How far (left, top, right, bottom pt) a shape's text area sits inside
    its outline: the shape type's textboxrect, in its coordsize. Formula
    rectangles of rounded shapes are worked out from the corner radius, and
    the ellipse's from its preset."""
    path = (shape.find("v:path") or (shapetype.find("v:path") if shapetype else None))
    rect = str(path.get("textboxrect") or "") if path else ""
    nums = rect.split(";")[0].split(",")
    if len(nums) == 4 and all(re.fullmatch(r"-?\d+", n.strip()) for n in nums):
        size = str((shapetype or shape).get("coordsize") or "21600,21600").split(",")
        try:
            cw, ch = float(size[0]), float(size[-1])
        except ValueError:
            cw = ch = 21600.0
        l, t, r, b = (float(n) for n in nums)
        return (l / cw * width, t / ch * height, (cw - r) / cw * width, (ch - b) / ch * height)
    if geom == "roundRect" and adj:
        k = (1 - 0.5 ** 0.5) * adj[0] * min(width, height)   # where the corner arc is at 45 degrees
        return (k, k, k, k)
    if geom.startswith(("ellipse", "wedgeEllipse")):
        return (0.1464 * width, 0.1464 * height, 0.1464 * width, 0.1464 * height)
    return (0.0, 0.0, 0.0, 0.0)


def restore_vml_pictures(soup: BeautifulSoup) -> int:
    """A picture cropped to a shape (a rounded rectangle, say) can come only in
    the VML, with no <img> for other browsers. Adds a positioned <img> for each,
    saying which shape crops it. Returns how many were added."""
    shown = {el.get("v:shapes") for el in soup.find_all(attrs={"v:shapes": True})}
    count = 0
    for c in soup.find_all(string=lambda t: isinstance(t, Comment) and "vml" in t[:30]):
        vml = BeautifulSoup(str(c), "html.parser")
        for shape in vml.find_all(["v:shape", "v:rect", "v:roundrect", "v:oval"]):
            data = shape.find("v:imagedata", recursive=False)
            decl = parse_decls(shape.get("style", ""))
            if (data is None or not data.get("src") or shape.get("id") in shown
                    or shape.find_parent("v:group") or decl.get("rotation")
                    or str(decl.get("position", "")).lower() != "absolute"):
                continue
            geom, adj = _vml_outline(shape) or ("rect", ())
            styles = ["position:absolute"] + [
                f"{k}:{_vml_pt(decl.get(k)):.2f}pt" for k in ("left", "top", "width", "height")]
            if decl.get("z-index"):
                styles.append(f"z-index:{decl['z-index']}")
            if geom != "rect":
                styles.append(f"x-geom:{geom}")
                if adj:
                    styles.append("x-adj:" + " ".join(f"{v:.5f}" for v in adj))
            img = soup.new_tag("img", src=data["src"], style=";".join(styles),
                               alt=shape.get("alt") or "")
            img["v:shapes"] = shape.get("id") or ""
            c.insert_after(img)
            shown.add(shape.get("id"))
            count += 1
    return count


def restore_vml_text_boxes(soup: BeautifulSoup) -> int:
    """Publisher exports some text boxes (filled ones, for instance) as a
    picture of the text, keeping the real text only in the VML inside an
    <!--[if gte vml 1]> comment. Swaps each such picture for a positioned div
    holding that text, so it stays editable, with the shape's outline behind
    it. Shapes inside a VML group, rotated or WordArt shapes, inline ones and
    outlines PowerPoint has no preset for keep their picture. Returns how many
    were swapped."""
    found, text_rects = {}, {}
    for c in soup.find_all(string=lambda t: isinstance(t, Comment) and "vml" in t[:30]):
        vml = BeautifulSoup(str(c), "html.parser")
        for shape in vml.find_all(["v:shape", "v:rect", "v:roundrect", "v:oval"]):
            box = shape.find("v:textbox", recursive=False)
            if (box is None or not box.get_text(strip=True) or shape.find_parent("v:group")
                    or shape.find("v:textpath")):
                continue
            decl = parse_decls(shape.get("style", ""))
            if (decl.get("rotation") or "layout-flow" in str(box.get("style", ""))
                    or str(decl.get("position", "")).lower() != "absolute"):
                continue
            outline = _vml_outline(shape)
            if outline is None:
                continue
            found[shape.get("id")] = (shape, decl, box, outline)
        for st in vml.find_all("v:shapetype"):
            text_rects[st.get("id")] = st
    count = 0
    for img in soup.find_all("img"):
        hit = found.get(img.get("v:shapes"))
        if not hit:
            continue
        shape, decl, box, (geom, adj) = hit
        inset = [_vml_pt(v, 2.88) for v in (box.get("inset") or "").split(",")]
        inset += [2.88] * (4 - len(inset))          # left, top, right, bottom
        inner = box.find("div")
        pad = parse_decls(inner.get("style", "")) if inner else {}
        width, height = _vml_pt(decl.get("width")), _vml_pt(decl.get("height"))
        # the shape's own text area inside its outline
        area = _vml_text_area(shape, text_rects.get(str(shape.get("type") or "").lstrip("#")),
                              geom, adj, width, height)
        styles = [
            "position:absolute",
            f"left:{_vml_pt(decl.get('left')):.2f}pt",
            f"top:{_vml_pt(decl.get('top')):.2f}pt",
            f"width:{width:.2f}pt",
            f"height:{height:.2f}pt",
            f"padding-left:{area[0] + inset[0] + _vml_pt(pad.get('padding-left')):.2f}pt",
            f"padding-top:{area[1] + inset[1] + _vml_pt(pad.get('padding-top')):.2f}pt",
            f"padding-right:{area[2] + inset[2] + _vml_pt(pad.get('padding-right')):.2f}pt",
            f"padding-bottom:{area[3] + inset[3] + _vml_pt(pad.get('padding-bottom')):.2f}pt",
        ]
        fill = str(shape.get("fillcolor") or "").split()
        if fill and str(shape.get("filled", "t")).lower() not in ("f", "false"):
            styles.append(f"background-color:{fill[0]}")
        stroke = shape.find("v:stroke")
        if (str(shape.get("stroked", "t")).lower() not in ("f", "false")
                and str((stroke or {}).get("on", "t")).lower() not in ("f", "false")):
            color = (str(shape.get("strokecolor") or "black").split() or ["black"])[0]
            styles.append(f"border:{_vml_pt(shape.get('strokeweight'), 0.75):.2f}pt solid {color}")
            dash = str((stroke or {}).get("dashstyle") or "").lower()
            if dash and dash != "solid":
                styles.append(f"x-dash:{dash}")
        if geom != "rect":
            styles.append(f"x-geom:{geom}")
        if adj:
            styles.append("x-adj:" + " ".join(f"{v:.5f}" for v in adj))
        if decl.get("z-index"):
            styles.append(f"z-index:{decl['z-index']}")
        div = soup.new_tag("div", style=";".join(styles))
        for child in list((inner or box).children):
            div.append(child.extract())
        # the picture sits in a positioned span of its own
        target = img.parent if (img.parent and img.parent.name == "span"
                                and len(img.parent.find_all(True)) == 1) else img
        target.replace_with(div)
        count += 1
    return count


def find_page_containers(soup: BeautifulSoup, sheet: StyleSheet, dpi: float,
                         default_h: Optional[float] = None):
    """Publisher wraps each page in a sized, positioned div. Find them.
    Publisher sometimes writes a malformed height ('10.-1737in') on that
    div; default_h (px) stands in for it."""
    body = soup.body or soup
    candidates = []
    for el in body.find_all(["div", "table", "section"]):
        st = computed_style(el, sheet, {})
        w = to_px(st.get("width"), dpi)
        h = to_px(st.get("height"), dpi)
        # Publisher's own wrapper: a positioned div straight under <body>
        wrapper = el.name == "div" and el.parent is body and is_positioned(st)
        if not h and default_h and wrapper and st.get("height"):
            h = default_h
        if not w or not h or ((w < 200 or h < 200) and not wrapper):
            continue
        # must not be nested inside another candidate
        candidates.append((el, w, h, st))

    top = []
    for el, w, h, st in candidates:
        if not any(other is not el and other in el.parents for other, _, _, _ in candidates):
            top.append((el, w, h, st))
    return top


def read_page_size(export_dir: str, base: str) -> Optional[tuple]:
    """The real page size in inches, recorded by stage 1 as <base>_pagesize.txt."""
    path = os.path.join(export_dir, base + "_pagesize.txt")
    try:
        with open(path, encoding="ascii") as fh:
            w, h = (float(v) for v in fh.read().split()[:2])
    except (OSError, ValueError):
        return None
    return (w, h) if w > 0 and h > 0 else None


def bounding_page(boxes: list, dpi: float, fallback_in):
    if not boxes:
        return fallback_in[0] * dpi, fallback_in[1] * dpi
    max_x = max(b.x + (b.w or 0) for b in boxes)
    max_y = max(b.y + (b.h or 0) for b in boxes)
    # never smaller than the fallback page: an unpositioned publication is
    # almost always a normal page that Publisher exported as flowed HTML
    return max(max_x, fallback_in[0] * dpi), max(max_y, fallback_in[1] * dpi)


# ----------------------------------------------------------------------------
# Image resolution
# ----------------------------------------------------------------------------

class ImageResolver:
    def __init__(self, html_dir: str, export_dir: str, warn):
        self.html_dir = html_dir
        self.export_dir = export_dir
        self.warn = warn
        self.index = {}
        self.used = set()
        self.converted = {}
        for root, _dirs, files in os.walk(export_dir):
            for f in files:
                self.index.setdefault(f.lower(), os.path.join(root, f))

    def resolve(self, src: str) -> Optional[str]:
        if not src or src.lower().startswith(("http://", "https://", "data:")):
            return None
        src = src.replace("\\", "/").split("?")[0].split("#")[0]
        cand = os.path.normpath(os.path.join(self.html_dir, src))
        if os.path.isfile(cand):
            return self._usable(cand)
        base = os.path.basename(src).lower()
        hit = self.index.get(base)
        if hit:
            return self._usable(hit)
        self.warn(f"image not found on disk: {src}")
        return None

    def _usable(self, path: str) -> Optional[str]:
        ext = os.path.splitext(path)[1].lower()
        if ext in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"):
            if ext in (".bmp", ".tif", ".tiff"):
                return self._to_png(path)
            self.used.add(os.path.normcase(os.path.abspath(path)))
            return path
        if ext in (".wmf", ".emf", ".webp", ".ico"):
            return self._to_png(path)
        self.warn(f"unsupported image type skipped: {os.path.basename(path)}")
        return None

    def _to_png(self, path: str) -> Optional[str]:
        if path in self.converted:
            return self.converted[path]
        if not HAVE_PIL:
            self.warn(f"cannot convert {os.path.basename(path)} without Pillow")
            return None
        out = os.path.join(self.export_dir, "_pptx_converted")
        os.makedirs(out, exist_ok=True)
        dest = os.path.join(out, os.path.splitext(os.path.basename(path))[0] + ".png")
        try:
            with Image.open(path) as im:
                im.convert("RGBA" if im.mode in ("RGBA", "LA", "P") else "RGB").save(dest)
            self.converted[path] = dest
            self.used.add(os.path.normcase(os.path.abspath(path)))
            return dest
        except Exception as exc:                      # noqa: BLE001
            self.warn(f"could not convert {os.path.basename(path)}: {exc}")
            self.converted[path] = None
            return None


def aspect_of(path: str) -> Optional[float]:
    if not HAVE_PIL:
        return None
    try:
        with Image.open(path) as im:
            return im.width / im.height if im.height else None
    except Exception:                                  # noqa: BLE001
        return None


def hires_candidates(export_dir: str, base: str, page_no: int) -> list:
    pat = os.path.join(export_dir, f"{base}_page_{page_no:03d}_image_*.png")
    return sorted(globmod.glob(pat))


# ----------------------------------------------------------------------------
# PPTX emission
# ----------------------------------------------------------------------------

def px_to_emu(px: float, dpi: float) -> int:
    return int(round(px / dpi * EMU_PER_INCH))


def _line_sizes(para) -> set:
    """The largest font size on each line of a paragraph, spaces included;
    None for a run with no size."""
    sizes, line = set(), []
    for run in para.runs:
        for k, piece in enumerate(run.text.split("\v")):
            if k and line:
                sizes.add(max(line, key=lambda s: s or 0.0) if None not in line else None)
                line = []
            if piece:
                line.append(run.size_pt)
    if line:
        sizes.add(max(line, key=lambda s: s or 0.0) if None not in line else None)
    return sizes


def _line_tightening(para, width_pt: float) -> dict:
    """Letter spacing (pt) to take off each line of a paragraph kept on
    Publisher's lines, keyed by line number, for lines that only just fit."""
    lines = [[]]
    for run in para.runs:
        for k, piece in enumerate(run.text.split("\v")):
            if k:
                lines.append([])
            lines[-1].append((run, piece))
    stops = sorted(pos for pos, _ in para.tab_stops)
    out = {}
    for i, pieces in enumerate(lines):
        # where the line ends, with PowerPoint's extra width, following tabs
        # to their stops; only the text after the last tab can be squeezed
        x = para.left_indent_pt + (para.indent_pt if i == 0 else 0.0)
        chars = 0
        for k, (run, piece) in enumerate(pieces):
            if k == len(pieces) - 1:
                piece = piece.rstrip()
            for j, seg in enumerate(piece.split("\t")):
                if j:
                    x = next_tab_px(x, stops, TAB_STOP_PT)
                    chars = 0
                x += (text_width_px(seg, run.font, run.size_pt or 12.0, bool(run.bold),
                                    bool(run.italic)) * (1.0 + KEEP_LINES_SLACK)
                      + (run.spacing_pt or 0.0) * len(seg))
                chars += len(seg)
        need = x - width_pt
        if need > 0 and chars:
            size = max((run.size_pt or 12.0) for run, _ in pieces)
            # PowerPoint's letter spacing is in hundredths of a point
            out[i] = min(math.ceil(need / chars * 100) / 100, KEEP_LINES_MAX_SPC * size)
    return out


def build_slide(prs, boxes, resolver, dpi, warn, verbose):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sw_px = prs.slide_width / EMU_PER_INCH * dpi
    sh_px = prs.slide_height / EMU_PER_INCH * dpi

    for b in sorted(boxes, key=lambda z: (z.z or 0, z.order)):
        w = max(b.w or 1.0, 1.0)
        h = max(b.h or 1.0, 1.0)
        # keep shapes from running miles off the canvas
        if b.x > sw_px * 1.5 or b.y > sh_px * 1.5:
            warn(f"shape far outside the page, kept anyway ({int(b.x)},{int(b.y)}px)")
        left, top = px_to_emu(b.x, dpi), px_to_emu(b.y, dpi)
        width, height = px_to_emu(w, dpi), px_to_emu(h, dpi)

        if b.kind == "image":
            path = resolver.resolve(b.src)
            if not path:
                continue
            try:
                pic = slide.shapes.add_picture(path, left, top, width, height)
            except Exception as exc:                   # noqa: BLE001
                warn(f"could not place image {os.path.basename(path)}: {exc}")
                continue
            if b.geom in SHAPE_GEOMS and b.geom != "rect":
                # cropped to a shape, as in Publisher
                pic.auto_shape_type = SHAPE_GEOMS[b.geom]
                av = pic._element.spPr.find(qn("a:prstGeom")).find(qn("a:avLst"))
                for i, v in enumerate(b.adj):
                    av.append(av.makeelement(qn("a:gd"), {
                        "name": "adj" if i == 0 else f"adj{i + 1}",
                        "fmla": f"val {int(round(v * 100000))}"}))
            continue

        if b.kind == "rect":
            shp = slide.shapes.add_shape(SHAPE_GEOMS.get(b.geom, MSO_SHAPE.RECTANGLE),
                                         left, top, width, height)
            for i, v in enumerate(b.adj[:len(shp.adjustments)]):
                shp.adjustments[i] = v
            if b.fill is not None:
                shp.fill.solid()
                shp.fill.fore_color.rgb = b.fill
            else:
                shp.fill.background()
            if b.line is not None:
                shp.line.color.rgb = b.line
                shp.line.width = Pt(max(b.line_w_px * 72.0 / dpi, 0.5))
                if b.dash in LINE_DASHES:
                    shp.line.dash_style = LINE_DASHES[b.dash]
            else:
                shp.line.fill.background()
            shp.shadow.inherit = False
            continue

        # text
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"middle": MSO_ANCHOR.MIDDLE,
                              "bottom": MSO_ANCHOR.BOTTOM}.get(b.anchor, MSO_ANCHOR.TOP)
        # PowerPoint rounds spacing in points to whole points, so each
        # paragraph's rounding is carried into the next one's space after
        carry = 0.0
        for pi, para in enumerate(b.paras):
            p = tf.paragraphs[0] if pi == 0 else tf.add_paragraph()
            if para.align is not None:
                p.alignment = para.align
            sizes = _line_sizes(para)
            if para.exact_line_pt and len(sizes) == 1 and None not in sizes:
                # a multiple isn't rounded: PowerPoint's single spacing is
                # 1.2 times the largest font size on the line, whatever the font
                p.line_spacing = para.exact_line_pt / (1.2 * sizes.pop())
            elif para.exact_line_pt:
                p.line_spacing = Pt(para.exact_line_pt)
            elif para.line_spacing:
                p.line_spacing = para.line_spacing
            if "\t" in para.text():
                ppr = p._p.get_or_add_pPr()
                ppr.set("defTabSz", str(int(Pt(TAB_STOP_PT))))
                if para.tab_stops:
                    tab_lst = ppr.makeelement(qn("a:tabLst"), {})
                    for pos_pt, algn in sorted(para.tab_stops):
                        tab_lst.append(ppr.makeelement(qn("a:tab"), {
                            "pos": str(int(Pt(pos_pt))), "algn": TAB_ALIGN.get(algn, "l")}))
                    # schema order: after spacing, which is already set
                    ppr.append(tab_lst)
            if para.left_indent_pt:
                p._p.get_or_add_pPr().set("marL", str(int(Pt(min(para.left_indent_pt, 288)))))
            if para.indent_pt:
                # first-line indent, negative for a hanging one: python-pptx
                # has no API for either
                p._p.get_or_add_pPr().set("indent", str(int(Pt(
                    max(min(para.indent_pt, 144), -para.left_indent_pt)))))
            if para.space_after_pt or carry:
                want = min(para.space_after_pt, 48) + carry
                after = max(round(want), 0)
                carry = want - after
                if after:
                    p.space_after = Pt(after)
            tighten = _line_tightening(para, w * 72.0 / dpi) if para.keep_lines else {}
            line = 0
            for run in para.runs:
                # "\v" is a line break inside the paragraph
                for k, piece in enumerate(run.text.split("\v")):
                    if k:
                        line += 1
                        p.add_line_break()
                        if run.size_pt:
                            # sized like its line: a bare break is 18pt
                            p._p[-1].get_or_add_rPr().set("sz", str(int(run.size_pt * 100)))
                    if not piece:
                        continue
                    r = p.add_run()
                    r.text = piece
                    f = r.font
                    if run.font:
                        f.name = run.font
                    f.size = Pt(max(min(run.size_pt or 12.0, 400.0), 1.0))
                    f.bold = run.bold
                    f.italic = run.italic
                    f.underline = run.underline
                    if run.color is not None:
                        f.color.rgb = run.color
                    rpr = r._r.get_or_add_rPr()
                    if run.caps:
                        rpr.set("cap", run.caps)
                    spacing = (run.spacing_pt or 0.0) - tighten.get(line, 0.0)
                    if round(spacing * 100):
                        rpr.set("spc", str(int(round(max(min(spacing, 100.0), -100.0) * 100))))
        if verbose:
            print(f"    text  {int(b.x):>5},{int(b.y):>5}  {int(w):>4}x{int(h):<4}  "
                  f"{b.paras[0].text()[:40]!r}" if b.paras else "")
    return slide


def add_image_slide(prs, path, dpi, caption=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sw, sh = prs.slide_width, prs.slide_height
    ar = aspect_of(path)
    margin = int(0.4 * EMU_PER_INCH)
    avail_w, avail_h = sw - 2 * margin, sh - 2 * margin
    if ar:
        if avail_w / avail_h > ar:
            h = avail_h
            w = int(h * ar)
        else:
            w = avail_w
            h = int(w / ar)
    else:
        w, h = avail_w, avail_h
    slide.shapes.add_picture(path, int((sw - w) / 2), int((sh - h) / 2), w, h)
    if caption:
        tb = slide.shapes.add_textbox(margin, int(sh - 0.38 * EMU_PER_INCH),
                                      sw - 2 * margin, int(0.28 * EMU_PER_INCH))
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = caption
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor.from_string("808080")
    return slide


def add_text_slide(prs, title, body):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    margin = int(0.5 * EMU_PER_INCH)
    tb = slide.shapes.add_textbox(margin, margin,
                                  prs.slide_width - 2 * margin,
                                  prs.slide_height - 2 * margin)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    r.font.size = Pt(16)
    r.font.bold = True
    for line in body.splitlines():
        p = tf.add_paragraph()
        r = p.add_run()
        r.text = line
        r.font.size = Pt(11)
    return slide


# ----------------------------------------------------------------------------
# Driver
# ----------------------------------------------------------------------------

def natural_key(path: str):
    name = os.path.basename(path).lower()
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", name)]


def html_files_for(export_dir: str, base: str) -> list:
    """Publisher writes <base>.htm for page 1 plus page*.htm for the rest."""
    found = []
    for pattern in ("*.htm", "*.html"):
        found.extend(globmod.glob(os.path.join(export_dir, pattern)))
        found.extend(globmod.glob(os.path.join(export_dir, f"{base}_files", pattern)))
    seen, out = set(), []
    primary = os.path.join(export_dir, base + ".htm")
    if os.path.isfile(primary):
        out.append(primary)
        seen.add(os.path.normcase(primary))
    for f in sorted(found, key=natural_key):
        nc = os.path.normcase(f)
        if nc in seen:
            continue
        if os.path.basename(f).lower() in ("index.htm", "index.html") and out:
            continue
        seen.add(nc)
        out.append(f)
    return out


def convert_export(export_dir: str, out_path: str, args) -> dict:
    base = os.path.basename(os.path.normpath(export_dir))
    warnings: list = []

    def warn(msg):
        warnings.append(msg)
        if args.verbose:
            print(f"    ! {msg}")

    htmls = html_files_for(export_dir, base)
    if not htmls:
        raise FileNotFoundError(f"no .htm found in {export_dir}")

    fallback = tuple(float(v) for v in args.page_size.lower().split("x"))
    # Publisher sizes its page wrapper to the content, not the page, and
    # sometimes writes a malformed height, so prefer stage 1's real size.
    real_page = read_page_size(export_dir, base)
    dpi = float(args.dpi)
    resolver = ImageResolver(export_dir, export_dir, warn)
    conv = Converter(args, warn)
    conv.pub_text = load_pub_text(export_dir, base)

    prs = Presentation()
    slide_pages = []   # (page_w_px, page_h_px, boxes)

    for html_path in htmls:
        with open(html_path, "rb") as fh:
            raw = fh.read()
        soup = BeautifulSoup(raw, "lxml")
        restore_vml_text_boxes(soup)
        restore_vml_pictures(soup)
        sheet = StyleSheet()
        for st in soup.find_all("style"):
            sheet.add_css(st.get_text())
        for link in soup.find_all("link", rel=lambda v: v and "stylesheet" in " ".join(v).lower()):
            href = unquote(link.get("href") or "")
            cand = os.path.normpath(os.path.join(os.path.dirname(html_path), href))
            if os.path.isfile(cand):
                try:
                    with open(cand, "rb") as fh:
                        sheet.add_css(fh.read().decode("utf-8", "replace"))
                except OSError:
                    warn(f"could not read stylesheet {href}")

        body = soup.body or soup
        body_style = computed_style(body, sheet, {}) if isinstance(body, Tag) else {}
        pages = find_page_containers(soup, sheet, dpi,
                                     (real_page or fallback)[1] * dpi)

        resolver.html_dir = os.path.dirname(html_path)

        if pages:
            for el, w, h, st in pages:
                conv.order = 0
                boxes = []
                conv.walk(el, sheet, {**body_style, **st}, (0.0, 0.0), w, h, boxes)
                if real_page:
                    w, h = real_page[0] * dpi, real_page[1] * dpi
                else:
                    # the wrapper can be smaller than its content
                    w, h = bounding_page(boxes, dpi, (w / dpi, h / dpi))
                slide_pages.append((w, h, boxes))
        else:
            conv.order = 0
            boxes = []
            page_in = real_page or fallback
            pw = page_in[0] * dpi
            ph = page_in[1] * dpi
            conv.walk(body, sheet, body_style, (0.0, 0.0), pw, ph, boxes)
            if not real_page:
                pw, ph = bounding_page(boxes, dpi, fallback)
            slide_pages.append((pw, ph, boxes))

    if not slide_pages:
        raise ValueError("no page content found in the HTML")

    # one deck-wide slide size: PowerPoint/Canva allow only one
    page_w = max(p[0] for p in slide_pages)
    page_h = max(p[1] for p in slide_pages)
    w_in = min(max(page_w / dpi, 1.0), PPTX_MAX_IN)
    h_in = min(max(page_h / dpi, 1.0), PPTX_MAX_IN)
    prs.slide_width = Emu(int(round(w_in * EMU_PER_INCH)))
    prs.slide_height = Emu(int(round(h_in * EMU_PER_INCH)))
    sizes = {(round(p[0]), round(p[1])) for p in slide_pages}
    if len(sizes) > 1:
        warn(f"pages have {len(sizes)} different sizes; the deck uses the largest "
             f"({w_in:.2f} x {h_in:.2f} in) and smaller pages sit top-left")

    for idx, (pw, ph, boxes) in enumerate(slide_pages, start=1):
        if args.hires:
            apply_hires(boxes, export_dir, base, idx, resolver, warn)
        if args.verbose:
            print(f"  page {idx}: {len(boxes)} shapes  ({pw/dpi:.2f} x {ph/dpi:.2f} in)")
        build_slide(prs, boxes, resolver, dpi, warn, args.verbose)

    extra = 0
    if not args.no_extras:
        extra = append_extras(prs, export_dir, base, resolver, dpi, warn)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    prs.save(out_path)

    stats = {
        "pptx": out_path,
        "slides": len(slide_pages) + extra,
        "pages": len(slide_pages),
        "extras": extra,
        "shapes": sum(len(p[2]) for p in slide_pages),
        "warnings": warnings,
        "size_in": (w_in, h_in),
    }

    if args.report:
        report = os.path.splitext(out_path)[0] + "_conversion_report.txt"
        with open(report, "w", encoding="utf-8") as fh:
            fh.write(f"Source export folder : {export_dir}\n")
            fh.write(f"PowerPoint written   : {out_path}\n")
            fh.write(f"Slide size           : {w_in:.2f} x {h_in:.2f} in\n")
            fh.write(f"Layout slides        : {len(slide_pages)}\n")
            fh.write(f"Extra slides         : {extra}\n")
            fh.write(f"Shapes placed        : {stats['shapes']}\n")
            fh.write(f"HTML files read      : {len(htmls)}\n\n")
            if warnings:
                fh.write("Warnings:\n")
                for w in warnings:
                    fh.write(f"  - {w}\n")
            else:
                fh.write("No warnings.\n")
        stats["report"] = report
    return stats


def apply_hires(boxes, export_dir, base, page_no, resolver, warn):
    """Swap Publisher's web-quality images for the 300dpi COM PNGs, in order."""
    cands = hires_candidates(export_dir, base, page_no)
    if not cands:
        return
    img_boxes = [b for b in sorted(boxes, key=lambda z: z.order) if b.kind == "image"]
    if not img_boxes:
        return
    if len(cands) != len(img_boxes):
        warn(f"page {page_no}: {len(img_boxes)} images in HTML but {len(cands)} "
             f"exported PNGs; high-res swap skipped for this page")
        return
    for box, png in zip(img_boxes, cands):
        box_ar = (box.w / box.h) if box.h else None
        png_ar = aspect_of(png)
        if box_ar and png_ar and abs(box_ar - png_ar) / max(box_ar, png_ar) > 0.12:
            warn(f"page {page_no}: aspect mismatch for {os.path.basename(png)}; kept web image")
            continue
        box.src = os.path.basename(png)
        resolver.used.add(os.path.normcase(os.path.abspath(png)))


def append_extras(prs, export_dir, base, resolver, dpi, warn) -> int:
    """Scratch-area art and text never reach the HTML; carry them in on their own slides."""
    added = 0
    patterns = [f"{base}_scratch_*_image_*.png", f"{base}_scratch_*_text_*.png"]
    files = []
    for pat in patterns:
        files.extend(sorted(globmod.glob(os.path.join(export_dir, pat)), key=natural_key))
    for f in files:
        if os.path.normcase(os.path.abspath(f)) in resolver.used:
            continue
        try:
            add_image_slide(prs, f, dpi, caption=f"Scratch area (off-canvas): {os.path.basename(f)}")
            added += 1
        except Exception as exc:                       # noqa: BLE001
            warn(f"could not add scratch image {os.path.basename(f)}: {exc}")

    txt = os.path.join(export_dir, base + "_scratch_text.txt")
    if os.path.isfile(txt):
        try:
            with open(txt, "r", encoding="utf-8", errors="replace") as fh:
                content = fh.read().strip()
            if content:
                lines = content.splitlines()
                per = 34
                for i in range(0, len(lines), per):
                    add_text_slide(prs, "Scratch-area text (not in the HTML export)",
                                   "\n".join(lines[i:i + per]))
                    added += 1
        except OSError as exc:
            warn(f"could not read {os.path.basename(txt)}: {exc}")
    return added


def scan_tree(root: str) -> dict:
    """Walk the tree once and report what is actually there.

    An export folder is any folder holding at least one .htm/.html file. The
    earlier rule demanded <FolderName>.htm, which missed exports where Publisher
    named the page differently (index.htm and friends) or where the folder had
    been renamed after the export.
    """
    export_dirs, pub_files, companion_only = [], [], []
    for dirpath, dirnames, filenames in os.walk(root):
        base = os.path.basename(os.path.normpath(dirpath))

        # Publisher's own asset folder: its pages are read via the parent export
        # folder, so it is never an export root in its own right.
        dirnames[:] = [d for d in dirnames if not d.lower().endswith("_files")]
        if base.lower().endswith("_files"):
            continue

        pub_files.extend(os.path.join(dirpath, f) for f in filenames
                         if f.lower().endswith(".pub"))

        htms = [f for f in filenames if f.lower().endswith((".htm", ".html"))]
        if htms:
            export_dirs.append(dirpath)
        elif any(os.path.isdir(os.path.join(dirpath, d)) and d.lower().endswith("_files")
                 for d in os.listdir(dirpath) if os.path.isdir(os.path.join(dirpath, d))):
            companion_only.append(dirpath)

    return {
        "exports": sorted(set(export_dirs)),
        "pubs": sorted(pub_files),
        "companion_only": sorted(companion_only),
    }


def find_export_dirs(root: str) -> list:
    return scan_tree(root)["exports"]


def explain_no_exports(root: str, scan: dict) -> str:
    """Say what was found instead, so the failure diagnoses itself."""
    lines = [f"No Publisher export folders found beneath {root}"]
    lines.append("  (an export folder is one containing a .htm or .html file)")

    pubs = scan["pubs"]
    if pubs:
        lines.append("")
        lines.append(f"  Found {len(pubs)} Publisher file(s) here but no HTML export beside them:")
        for p in pubs[:5]:
            lines.append(f"    {p}")
        if len(pubs) > 5:
            lines.append(f"    ... and {len(pubs) - 5} more")
        lines.append("")
        lines.append("  Run the export script first, in the folder holding the .pub files:")
        lines.append('    .\\PublisherPubToHTMLPNGfilesFinal.ps1 -Filter "*.pub" -Recurse')
        lines.append("")
        lines.append("  That leaves one subfolder per publication containing <BaseName>.htm,")
        lines.append("  which is what this converter reads.")
    elif scan["companion_only"]:
        lines.append("")
        lines.append("  Found Publisher _files asset folder(s) with no .htm beside them:")
        for d in scan["companion_only"][:5]:
            lines.append(f"    {d}")
        lines.append("")
        lines.append("  The .htm was probably moved or deleted. Re-run the export script,")
        lines.append("  or delete the export subfolder so it will be rebuilt (the export")
        lines.append("  script skips a publication whose .htm already exists).")
    else:
        lines.append("")
        lines.append("  Nothing that looks like a Publisher export is in that tree.")
        lines.append("  Check the path, and that you pointed at the folder holding the")
        lines.append("  .pub files or their export subfolders.")
    return "\n".join(lines)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Convert Publisher filtered-HTML exports into layout-preserving .pptx for Canva.")
    ap.add_argument("input", help="export folder, a .htm file, or a tree with --recurse")
    ap.add_argument("-o", "--output", help="output .pptx, or output folder with --recurse")
    ap.add_argument("--recurse", action="store_true", help="process every export folder below input")
    ap.add_argument("--dpi", type=float, default=96.0, help="CSS pixels per inch (default 96)")
    ap.add_argument("--hires", action="store_true", help="use the 300dpi COM-exported PNGs")
    ap.add_argument("--no-extras", action="store_true", help="skip scratch-area slides")
    ap.add_argument("--reflow", action="store_true",
                    help="let text wrap freely instead of breaking lines where Publisher does")
    ap.add_argument("--page-size", default="8.5x11", help="fallback page size in inches")
    ap.add_argument("--report", action="store_true", help="write a per-file conversion report")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    inp = os.path.abspath(args.input)
    if not os.path.exists(inp):
        print(f"Input not found: {inp}", file=sys.stderr)
        return 1

    if os.path.isfile(inp):
        targets = [os.path.dirname(inp)]
    elif args.recurse:
        scan = scan_tree(inp)
        targets = scan["exports"]
        if not targets:
            print(explain_no_exports(inp, scan), file=sys.stderr)
            return 1
    else:
        targets = [inp]

    print("")
    print("Publisher HTML export -> PowerPoint (Canva-ready)")
    print("=" * 52)

    ok = failed = 0
    total_warn = 0
    for d in targets:
        base = os.path.basename(os.path.normpath(d))
        if args.output and not args.recurse and os.path.splitext(args.output)[1].lower() == ".pptx":
            out = os.path.abspath(args.output)
        else:
            out_dir = os.path.abspath(args.output) if args.output else d
            out = os.path.join(out_dir, base + ".pptx")
        print(f"\n{base}")
        try:
            stats = convert_export(d, out, args)
        except Exception as exc:                        # noqa: BLE001
            print(f"  FAILED: {exc}")
            if args.verbose:
                import traceback
                traceback.print_exc()
            failed += 1
            continue
        ok += 1
        total_warn += len(stats["warnings"])
        print(f"  {stats['pages']} page slide(s) + {stats['extras']} extra, "
              f"{stats['shapes']} shapes, {stats['size_in'][0]:.2f} x {stats['size_in'][1]:.2f} in")
        print(f"  -> {out}")
        if stats["warnings"] and not args.verbose:
            print(f"  {len(stats['warnings'])} warning(s)"
                  + (f"; see {os.path.basename(stats['report'])}" if args.report else
                     " (re-run with --report or -v for detail)"))

    print("")
    print("=" * 52)
    print(f"Converted: {ok}    Failed: {failed}    Warnings: {total_warn}")
    print("")
    print("In Canva: Create a design > Import file, or drag the .pptx onto the")
    print("Projects page. Fonts Canva does not have are substituted, so check")
    print("line breaks on each page before sending anything to print.")
    print("")
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())