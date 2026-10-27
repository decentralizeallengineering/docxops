"""header_footer, section_break, page_setup."""

from __future__ import annotations

import copy
import re

from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml.ns import qn
from docx.section import Section
from docx.shared import Twips

from ._core import OUT, PATH, TEXT, Ctx, OpError, finish, load, obj, operation
from ._xml import (
    PPR_ORDER, Styles, field_runs, find_heading, get_or_add_pPr, insert_ordered, is_w,
    make, run, section_end_block, top_level_block,
)

SECTION = {"type": "integer", "minimum": 1}
TOKEN = re.compile(r"\{(page|pages)\}")
FIELD_FOR = {"page": "PAGE", "pages": "NUMPAGES"}
PAGE_NUMBER_TEXT = {"page": "{page}", "page_x_of_y": "Page {page} of {pages}"}
JC = {"left": "left", "center": "center", "right": "right"}


def _content_paragraph(text: str, style_id: str | None, align: str):
    p = make("w:p")
    pPr = make("w:pPr")
    if style_id:
        pPr.append(make("w:pStyle", val=style_id))
    insert_ordered(pPr, make("w:jc", val=JC[align]), PPR_ORDER)
    p.append(pPr)
    pos = 0
    for m in TOKEN.finditer(text):
        if m.start() > pos:
            p.append(run(text[pos:m.start()]))
        # Result left blank: Word and LibreOffice compute PAGE/NUMPAGES at layout.
        for r in field_runs(f"{FIELD_FOR[m.group(1)]} \\* MERGEFORMAT", ""):
            p.append(r)
        pos = m.end()
    if pos < len(text):
        p.append(run(text[pos:]))
    return p


def _fill(hf, paragraphs: list) -> None:
    el = hf._element
    for child in list(el):
        el.remove(child)
    for p in paragraphs or [make("w:p")]:
        el.append(p)


@operation(obj(
    "Set header and/or footer text (replaces existing). {page} and {pages} become fields; "
    "page_numbers adds Page X of Y or X. section: 1-based, default all.",
    {"path": PATH, "header": {"type": "string"}, "footer": {"type": "string"},
     "page_numbers": {"enum": ["none", "page", "page_x_of_y"]},
     "page_numbers_in": {"enum": ["header", "footer"]},
     "align": {"enum": ["left", "center", "right"]}, "section": SECTION, "out": OUT},
))
def header_footer(ctx: Ctx, path: str, header: str | None = None, footer: str | None = None,
                  page_numbers: str = "none", page_numbers_in: str = "footer",
                  align: str = "center", section: int | None = None,
                  out: str | None = None) -> None:
    if header is None and footer is None and page_numbers == "none":
        raise OpError("nothing to do: give header, footer or page_numbers")
    doc = load(path)
    styles = Styles(doc)
    sections = list(doc.sections)
    if section is not None and section > len(sections):
        raise OpError(f"section {section} out of range; document has {len(sections)}")

    content = {"header": header, "footer": footer}
    numbers = PAGE_NUMBER_TEXT.get(page_numbers)
    touched = []
    for i, s in enumerate(sections, 1):
        if section is not None and i != section:
            continue
        for kind in ("header", "footer"):
            text = content[kind]
            with_numbers = numbers if page_numbers_in == kind else None
            if text is None and with_numbers is None:
                continue
            hf = s.header if kind == "header" else s.footer
            if section is None and i > 1 and hf.is_linked_to_previous:
                continue  # inherits from the section before
            hf.is_linked_to_previous = False
            sid = styles.find(kind.capitalize())
            sid = sid.get(qn("w:styleId")) if sid is not None else None
            paras = []
            if text:
                paras += [_content_paragraph(line, sid, align) for line in text.split("\n")]
            if with_numbers:
                paras.append(_content_paragraph(with_numbers, sid, align))
            _fill(hf, paras)
            touched.append({"section": i, "part": kind})
    ctx.details.update({"updated": touched, "page_numbers": page_numbers})
    ctx.changed = bool(touched)
    finish(ctx, doc, path, out)


# ----------------------------------------------------------------- sections

def governing_sectPr(body, block):
    """The sectPr that defines the section containing a body-level block."""
    node = block
    while node is not None:
        if is_w(node, "p"):
            sp = node.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
            if sp is not None:
                return sp
        if is_w(node, "sectPr"):
            return node
        node = node.getnext()
    sp = body.find(qn("w:sectPr"))
    if sp is None:
        sp = make("w:sectPr")
        body.append(sp)
    return sp


def break_before(body, block):
    """Start a new section at block. Returns the sectPr of the section that starts there."""
    prev = block.getprevious()
    if prev is None:
        raise OpError("cannot insert a section break before the first block of the document")
    if is_w(prev, "p") and prev.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
        return governing_sectPr(body, block)  # already a section boundary
    if not is_w(prev, "p"):
        prev = make("w:p")
        block.addprevious(prev)
    gov = governing_sectPr(body, block)
    ended = copy.deepcopy(gov)
    # The new (earlier) section keeps the header/footer references; the later one
    # inherits them, which is what Word itself does on "Insert > Section break".
    for ref in gov.findall(qn("w:headerReference")) + gov.findall(qn("w:footerReference")):
        gov.remove(ref)
    insert_ordered(get_or_add_pPr(prev), ended, PPR_ORDER)
    return gov


def set_orientation(sec: Section, orientation: str) -> None:
    want = WD_ORIENT.LANDSCAPE if orientation == "landscape" else WD_ORIENT.PORTRAIT
    w, h = sec.page_width, sec.page_height
    if w is not None and h is not None:
        if (want == WD_ORIENT.LANDSCAPE) != (w > h):
            sec.page_width, sec.page_height = h, w
    sec.orientation = want


START = {"new_page": WD_SECTION.NEW_PAGE, "continuous": WD_SECTION.CONTINUOUS,
         "odd_page": WD_SECTION.ODD_PAGE, "even_page": WD_SECTION.EVEN_PAGE}


@operation(obj(
    "New section starting before a heading (default: at the end), optionally with its own "
    "orientation. isolate=true also ends it after that heading's content.",
    {"path": PATH, "before_heading": TEXT, "occurrence": SECTION,
     "orientation": {"enum": ["portrait", "landscape"]},
     "start": {"enum": list(START)}, "isolate": {"type": "boolean"}, "out": OUT},
))
def section_break(ctx: Ctx, path: str, before_heading: str | None = None,
                  occurrence: int = 1, orientation: str | None = None,
                  start: str = "new_page", isolate: bool = False,
                  out: str | None = None) -> None:
    doc = load(path)
    body = doc.element.body
    if before_heading is None:
        if isolate:
            raise OpError("isolate needs before_heading")
        sec = doc.add_section(START[start])
        sectPr = sec._sectPr
    else:
        styles = Styles(doc)
        h = find_heading(doc, before_heading, occurrence, ctx, styles)
        block = top_level_block(h.p, body)
        if isolate:
            after = section_end_block(doc, h, styles).getnext()
            if after is not None and not is_w(after, "sectPr"):
                following = break_before(body, after)
                if orientation:
                    Section(following, doc.part).start_type = WD_SECTION.NEW_PAGE
            else:
                ctx.warn("heading's section runs to the end of the document; nothing to isolate")
        sectPr = break_before(body, block)
        sec = Section(sectPr, doc.part)
        sec.start_type = START[start]
    if orientation:
        if start == "continuous":
            ctx.warn("orientation changes need a new page; Word will start one anyway")
        set_orientation(sec, orientation)
    sections = [s._sectPr for s in doc.sections]
    number = next(i for i, sp in enumerate(sections, 1) if sp is sectPr)
    ctx.details.update({"section": number, "sections": len(sections),
                        "orientation": "landscape" if sec.orientation == WD_ORIENT.LANDSCAPE
                        else "portrait", "start": start})
    ctx.changed = True
    finish(ctx, doc, path, out)


PAPER_TWIPS = {"A3": (16838, 23811), "A4": (11906, 16838), "A5": (8391, 11906),
               "Letter": (12240, 15840), "Legal": (12240, 20160)}
UNIT_TWIPS = {"mm": 1440 / 25.4, "cm": 1440 / 2.54, "in": 1440.0, "pt": 20.0}
MARGIN_KEYS = ("top", "bottom", "left", "right", "header", "footer", "gutter")


@operation(obj(
    "Set paper, orientation, margins (in unit) for a section (1-based) or all.",
    {"path": PATH, "section": SECTION, "orientation": {"enum": ["portrait", "landscape"]},
     "paper": {"enum": list(PAPER_TWIPS)}, "unit": {"enum": list(UNIT_TWIPS)},
     "margins": {"type": "object", "propertyNames": {"enum": list(MARGIN_KEYS)},
                 "additionalProperties": {"type": "number", "minimum": 0}}, "out": OUT},
))
def page_setup(ctx: Ctx, path: str, section: int | None = None,
               orientation: str | None = None, paper: str | None = None,
               unit: str = "mm", margins: dict | None = None, out: str | None = None) -> None:
    if not (orientation or paper or margins):
        raise OpError("nothing to do: give orientation, paper or margins")
    doc = load(path)
    sections = list(doc.sections)
    if section is not None and section > len(sections):
        raise OpError(f"section {section} out of range; document has {len(sections)}")
    applied = []
    for i, s in enumerate(sections, 1):
        if section is not None and i != section:
            continue
        if paper:
            w, h = PAPER_TWIPS[paper]
            landscape = (orientation or ("landscape" if s.orientation == WD_ORIENT.LANDSCAPE
                                         else "portrait")) == "landscape"
            s.page_width, s.page_height = (Twips(h), Twips(w)) if landscape else (Twips(w), Twips(h))
        if orientation:
            set_orientation(s, orientation)
        for k, v in (margins or {}).items():
            twips = Twips(round(v * UNIT_TWIPS[unit]))
            attr = {"header": "header_distance", "footer": "footer_distance"}.get(k, k + "_margin")
            if k == "gutter":
                attr = "gutter"
            setattr(s, attr, twips)
        if s.page_width is not None and s.left_margin is not None and s.right_margin is not None:
            if s.left_margin + s.right_margin >= s.page_width:
                raise OpError(f"section {i}: left+right margins exceed the page width")
        applied.append({"section": i,
                        "page_mm": [round(s.page_width.mm, 1), round(s.page_height.mm, 1)],
                        "orientation": "landscape" if s.orientation == WD_ORIENT.LANDSCAPE
                        else "portrait"})
    ctx.details["sections"] = applied
    ctx.changed = True
    finish(ctx, doc, path, out)
