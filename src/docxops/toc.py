"""add_toc: a real TOC field whose result is pre-filled with hyperlinked entries."""

from __future__ import annotations

import re

from docx.oxml.ns import qn

from ._core import OUT, PATH, Ctx, OpError, finish, load, obj, operation
from ._xml import (
    PPR_ORDER, Bookmarks, Styles, field_runs, get_or_add_pPr, headings,
    insert_ordered, is_w, iter_fields, instr_has_switch, make, paragraph,
    paragraph_bookmarks, run, set_update_fields, suppress_inherited_numbering,
    text_width_twips,
)

TOC_GALLERY = "Table of Contents"
SDT_TAG = "docxops-toc"


def parse_levels(levels) -> tuple[int, int]:
    if isinstance(levels, int):
        return 1, levels
    m = re.fullmatch(r"([1-9])(?:-([1-9]))?", str(levels).strip())
    if not m:
        raise OpError(f"levels must look like '1-3' or be 1..9, got {levels!r}")
    lo = int(m.group(1))
    hi = int(m.group(2) or m.group(1))
    if hi < lo:
        raise OpError(f"levels range is reversed: {levels!r}")
    return lo, hi


def is_toc_instr(instr: str) -> bool:
    s = instr.strip()
    if not re.match(r"(?i)TOC\b", s):
        return False
    # \c and \a build tables of figures/tables, not a contents list.
    return not (instr_has_switch(s, "c") or instr_has_switch(s, "a"))


def _is_toc_sdt(sdt) -> bool:
    sdtPr = sdt.find(qn("w:sdtPr"))
    if sdtPr is None:
        return False
    tag = sdtPr.find(qn("w:tag"))
    if tag is not None and tag.get(qn("w:val")) == SDT_TAG:
        return True
    gal = sdtPr.find(f"{qn('w:docPartObj')}/{qn('w:docPartGallery')}")
    if gal is None or gal.get(qn("w:val")) != TOC_GALLERY:
        return False
    content = sdt.find(qn("w:sdtContent"))
    return content is not None and any(is_toc_instr(f.instr) for f in iter_fields(content))


def _enclosing_p(el):
    while el is not None and not is_w(el, "p"):
        el = el.getparent()
    return el


def find_existing_tocs(body, styles: Styles) -> list[list]:
    """Groups of sibling elements that together form an existing TOC."""
    groups: list[list] = [[sdt] for sdt in body.iter(qn("w:sdt")) if _is_toc_sdt(sdt)]
    for f in iter_fields(body):
        if not is_toc_instr(f.instr) or _inside_toc_sdt(f.begin):
            continue
        first = _enclosing_p(f.begin)
        last = _enclosing_p(f.end)
        if first is None or last is None or first.getparent() is not last.getparent():
            continue
        group, node = [], first
        while node is not None:
            group.append(node)
            if node is last:
                break
            node = node.getnext()
        if not group or group[-1] is not last:
            continue
        prev = first.getprevious()
        if prev is not None and is_w(prev, "p") and styles.para_style_id(prev) == "TOCHeading":
            group.insert(0, prev)
        groups.append(group)
    return groups


def _inside_toc_sdt(el) -> bool:
    a = el.getparent()
    while a is not None:
        if is_w(a, "sdt") and _is_toc_sdt(a):
            return True
        a = a.getparent()
    return False


def _next_sdt_id(body) -> int:
    ids = [0]
    for el in body.iter(qn("w:sdtPr")):
        i = el.find(qn("w:id"))
        if i is not None:
            try:
                ids.append(abs(int(i.get(qn("w:val")))))
            except (TypeError, ValueError):
                pass
    return (max(ids) + 1) % 2147483647 or 1


def _entry_paragraph(style_id: str, tab_pos: int, text: str, bookmark: str):
    p = paragraph(style_id=style_id)
    pPr = get_or_add_pPr(p)
    tabs = make("w:tabs")
    tabs.append(make("w:tab", val="right", leader="dot", pos=tab_pos))
    insert_ordered(pPr, tabs, PPR_ORDER)
    link = make("w:hyperlink", anchor=bookmark, history=1)
    link.append(run(text))
    hidden = make("w:rPr")
    hidden.append(make("w:webHidden"))
    tab_run = run(rPr=hidden)
    tab_run.append(make("w:tab"))
    link.append(tab_run)
    # Page number: a PAGEREF whose result is left empty until Word updates fields.
    for r in field_runs(f"PAGEREF {bookmark} \\h", "", rPr=hidden):
        link.append(r)
    p.append(link)
    return p


@operation(obj(
    "Insert or replace a real TOC field, pre-filled with hyperlinked heading entries "
    "(page numbers blank until Word updates fields on open).",
    {
        "path": PATH,
        "levels": {"anyOf": [{"type": "string", "pattern": "^[1-9](-[1-9])?$"},
                             {"type": "integer", "minimum": 1, "maximum": 9}]},
        "title": {"type": "string"},
        "position": {"enum": ["auto", "start"]},
        "out": OUT,
    },
))
def add_toc(ctx: Ctx, path: str, levels: str | int = "1-3", title: str = "Contents",
            position: str = "auto", out: str | None = None) -> None:
    lo, hi = parse_levels(levels)
    doc = load(path)
    body = doc.element.body
    styles = Styles(doc)

    # 1. Remove any existing TOC, remembering where it was.
    anchor_parent, anchor_index = None, None
    groups = find_existing_tocs(body, styles)
    for g in groups:
        parent = g[0].getparent()
        if anchor_parent is None:
            anchor_parent, anchor_index = parent, parent.index(g[0])
        for el in g:
            el.getparent().remove(el)
    ctx.details["replaced_existing"] = len(groups)

    # 2. Bookmark each heading in range (reusing _Toc bookmarks from a prior run).
    bms = Bookmarks(doc.element)
    entries, added, reused = [], 0, 0
    for h in headings(doc, styles):
        if not lo <= h.level <= hi:
            continue
        text = re.sub(r"\s+", " ", h.text).strip()
        if not text:
            ctx.warn(f"skipped an empty heading (paragraph {h.index})")
            continue
        existing = paragraph_bookmarks(h.p, "_Toc")
        if existing:
            name = existing[0]
            reused += 1
        else:
            name = bms.wrap(h.p, bms.new_name("_Toc"))
            added += 1
        entries.append({"level": h.level, "text": text, "bookmark": name})

    # 3. Build the content control: title + one paragraph per entry.
    instr = f'TOC \\o "{lo}-{hi}" \\h \\z \\u'
    tab_pos = text_width_twips(doc)
    sdt = make("w:sdt")
    sdtPr = make("w:sdtPr")
    sdtPr.append(make("w:tag", val=SDT_TAG))
    sdtPr.append(make("w:id", val=_next_sdt_id(body)))
    dpo = make("w:docPartObj")
    dpo.append(make("w:docPartGallery", val=TOC_GALLERY))
    dpo.append(make("w:docPartUnique"))
    sdtPr.append(dpo)
    sdt.append(sdtPr)
    content = make("w:sdtContent")
    sdt.append(content)

    if title:
        content.append(paragraph(title, styles.ensure("TOC Heading", ctx)))
        suppress_inherited_numbering(styles, "TOCHeading")

    paras = []
    for e in entries:
        sid = styles.ensure(f"TOC {e['level']}", ctx)
        paras.append(_entry_paragraph(sid, tab_pos, e["text"], e["bookmark"]))
    if not paras:
        ctx.warn(f"no headings found at levels {lo}-{hi}; inserted an empty TOC field")
        paras.append(paragraph(style_id=styles.ensure("TOC 1", ctx)))
    *opening, end = field_runs(instr)
    first = paras[0]
    pos = 1 if first.find(qn("w:pPr")) is not None else 0
    for i, r in enumerate(opening):
        first.insert(pos + i, r)
    paras[-1].append(end)
    for p in paras:
        content.append(p)

    # 4. Place it: where the old TOC was, else after a leading Title/Subtitle.
    if position == "auto" and anchor_parent is not None:
        anchor_parent.insert(anchor_index, sdt)
        where = "replaced previous TOC in place"
    else:
        idx = 0
        if position == "auto":
            for child in body:
                if is_w(child, "p") and styles.para_style_name(child) in ("Title", "Subtitle"):
                    idx += 1
                    continue
                break
        body.insert(idx, sdt)
        where = "start of document" if idx == 0 else "after title"

    set_update_fields(doc)
    ctx.changed = True
    ctx.details.update({
        "field_code": instr,
        "position": where,
        "entries": entries,
        "bookmarks_added": added,
        "bookmarks_reused": reused,
        "page_numbers": "blank until fields are updated (w:updateFields is set)",
    })
    finish(ctx, doc, path, out)
