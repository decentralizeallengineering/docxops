"""add_caption (SEQ fields) and cross_reference (REF/PAGEREF to bookmarks)."""

from __future__ import annotations

import copy
import re

from docx.oxml.ns import qn

from ._core import OUT, PATH, TEXT, Ctx, OpError, finish, load, obj, operation
from ._text import apply, cleanup_empty, compile_pattern, find_matches, insert_at
from ._xml import (
    PPR_ORDER, V_NS, Bookmarks, Field, Styles, bookmark_text, body_paragraphs,
    field_result_runs, field_runs, find_heading, get_or_add_pPr, has_ancestor, insert_ordered, is_w,
    iter_fields, make, norm, para_text, paragraph, paragraph_bookmarks, run,
    set_field_result, set_update_fields,
)


def _roman(n: int) -> str:
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def _alpha(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _format_number(n: int, instr: str) -> str:
    m = re.search(r"\\\*\s*(\w+)", instr)
    fmt = m.group(1) if m else "ARABIC"
    if fmt == "ROMAN":
        return _roman(n)
    if fmt == "roman":
        return _roman(n).lower()
    if fmt == "ALPHABETIC":
        return _alpha(n)
    if fmt == "alphabetic":
        return _alpha(n).lower()
    return str(n)


def seq_identifier(f: Field) -> str | None:
    parts = f.instr.split()
    if len(parts) >= 2 and parts[0].upper() == "SEQ":
        return parts[1]
    return None


def renumber_seq(root) -> dict[str, list[tuple[Field, int]]]:
    """Recompute SEQ results in document order. Returns identifier -> [(field, n)]."""
    counters: dict[str, int] = {}
    seen: dict[str, list] = {}
    for f in iter_fields(root):
        ident = seq_identifier(f)
        if ident is None:
            continue
        key = ident.casefold()
        n = counters.get(key, 0)
        m = re.search(r"\\r\s*(\d+)", f.instr)
        if m:
            n = int(m.group(1))
        elif not re.search(r"\\c\b", f.instr):
            n += 1
        counters[key] = n
        set_field_result(f, "" if re.search(r"\\h\b", f.instr) else _format_number(n, f.instr))
        seen.setdefault(key, []).append((f, n))
    return seen


def refresh_refs(body) -> int:
    """Refresh REF results (text and \\p above/below) from their bookmarks."""
    order = {el: i for i, el in enumerate(body.iter())}
    starts = {b.get(qn("w:name")): b for b in body.iter(qn("w:bookmarkStart"))}
    n = 0
    for f in iter_fields(body):
        parts = f.instr.split()
        if len(parts) < 2 or parts[0].upper() != "REF" or parts[1] not in starts:
            continue
        if re.search(r"\\[nrw]\b", f.instr):
            continue  # number-only variants are left for Word to compute
        if re.search(r"\\p\b", f.instr):
            here = order.get(f.begin, 0)
            there = order.get(starts[parts[1]], 0)
            set_field_result(f, "above" if there < here else "below")
        else:
            text = bookmark_text(body, parts[1])
            if text is None:
                continue
            set_field_result(f, text)
        n += 1
    return n


def _enclosing_p(el):
    while el is not None and not is_w(el, "p"):
        el = el.getparent()
    return el


def _figure_paragraphs(body) -> list:
    out = []
    for p in body_paragraphs(body):
        if has_ancestor(p, "txbxContent"):
            continue
        if next(p.iter(qn("w:drawing")), None) is not None or \
                next(p.iter(f"{{{V_NS}}}imagedata"), None) is not None:
            out.append(p)
    return out


def _caption_field(p, label: str) -> Field | None:
    for f in iter_fields(p):
        ident = seq_identifier(f)
        if ident is not None and ident.casefold() == label.casefold():
            return f
    return None


def ensure_caption_bookmark(p, f: Field, bms: Bookmarks, name: str | None = None) -> str:
    """Bookmark the caption's label and number (start of paragraph to end of SEQ)."""
    existing = paragraph_bookmarks(p, "_Ref")
    if existing and name is None:
        return existing[0]
    last = f.end if f.simple else f.end.getparent()
    while last.getparent() is not p:
        last = last.getparent()
    first = None
    for child in p:
        if child.tag not in (qn("w:pPr"), qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
            first = child
            break
    return bms.wrap(p, name or bms.new_name("_Ref"), first, last)


@operation(obj(
    "Caption the Nth table or figure (target 1-based; -1 = last) with 'Label {SEQ}: text'. "
    "Tables above, figures below by default. Replaces an adjacent existing caption. "
    "Renumbers all captions.",
    {"path": PATH, "kind": {"enum": ["Table", "Figure"]}, "text": {"type": "string"},
     "target": {"type": "integer", "not": {"const": 0}},
     "position": {"enum": ["above", "below"]}, "separator": {"type": "string"}, "out": OUT},
    ["path", "kind"],
))
def add_caption(ctx: Ctx, path: str, kind: str, text: str = "", target: int = -1,
                position: str | None = None, separator: str = ": ",
                out: str | None = None) -> None:
    doc = load(path)
    body = doc.element.body
    styles = Styles(doc)
    if kind == "Table":
        items = [t for t in body.iter(qn("w:tbl")) if not has_ancestor(t, "tbl")]
    else:
        items = _figure_paragraphs(body)
    if not items:
        raise OpError(f"document has no {kind.lower()}s to caption")
    if not -len(items) <= target <= len(items):
        raise OpError(f"target {target} out of range; document has {len(items)} {kind.lower()}(s)")
    obj_el = items[target - 1 if target > 0 else target]
    number = items.index(obj_el) + 1
    position = position or ("above" if kind == "Table" else "below")

    neighbor = obj_el.getprevious() if position == "above" else obj_el.getnext()
    reuse_name, replaced = None, False
    if neighbor is not None and is_w(neighbor, "p") and _caption_field(neighbor, kind):
        names = paragraph_bookmarks(neighbor, "_Ref")
        reuse_name = names[0] if names else None
        neighbor.getparent().remove(neighbor)
        replaced = True

    p = paragraph(style_id=styles.ensure("Caption", ctx))
    if position == "above":
        insert_ordered(get_or_add_pPr(p), make("w:keepNext"), PPR_ORDER)
    p.append(run(kind + " "))
    for r in field_runs(f"SEQ {kind} \\* ARABIC", "0"):
        p.append(r)
    if text:
        p.append(run(separator + text))
    (obj_el.addprevious if position == "above" else obj_el.addnext)(p)

    seqs = renumber_seq(body)
    f = _caption_field(p, kind)
    bms = Bookmarks(doc.element)
    if reuse_name:
        bms.names.discard(reuse_name)
    name = ensure_caption_bookmark(p, f, bms, reuse_name or bms.new_name("_Ref"))
    refresh_refs(body)
    n = next(num for fld, num in seqs[kind.casefold()] if fld.begin is f.begin)
    ctx.details.update({"label": f"{kind} {n}", "number": n, "bookmark": name,
                        "target": number, "position": position, "replaced": replaced,
                        "caption_text": para_text(p)})
    ctx.changed = True
    finish(ctx, doc, path, out)


def _find_caption(body, target: str):
    m = re.fullmatch(r"\s*(\w+)\s+(\d+)\s*", target)
    seqs = renumber_seq(body)
    if m:
        for f, n in seqs.get(m.group(1).casefold(), []):
            if n == int(m.group(2)):
                return _enclosing_p(f.begin), f
    key = norm(target)
    for fields in seqs.values():
        for f, _ in fields:
            p = _enclosing_p(f.begin)
            if p is not None and key in norm(para_text(p)):
                return p, f
    return None, None


@operation(obj(
    "Replace placeholder text with a REF/PAGEREF field to a caption ('Table 2' or caption "
    "text) or heading. display: text, page (blank until update) or above_below.",
    {"path": PATH, "target": TEXT, "placeholder": TEXT,
     "target_type": {"enum": ["auto", "caption", "heading"]},
     "display": {"enum": ["text", "page", "above_below"]}, "out": OUT},
    ["path", "target", "placeholder"],
))
def cross_reference(ctx: Ctx, path: str, target: str, placeholder: str,
                    target_type: str = "auto", display: str = "text",
                    out: str | None = None) -> None:
    doc = load(path)
    body = doc.element.body
    styles = Styles(doc)
    bms = Bookmarks(doc.element)

    tp, cap_field = None, None
    if target_type in ("auto", "caption"):
        tp, cap_field = _find_caption(body, target)
    if tp is not None:
        kind = "caption"
        bookmark = ensure_caption_bookmark(tp, cap_field, bms)
    elif target_type == "caption":
        raise OpError(f"caption not found: {target!r}")
    else:
        h = find_heading(doc, target, 1, ctx, styles)
        kind, tp = "heading", h.p
        existing = paragraph_bookmarks(tp, "_Ref")
        bookmark = existing[0] if existing else bms.wrap(tp, bms.new_name("_Ref"))

    pattern = compile_pattern(placeholder, True, False)
    hits = []
    for p in body_paragraphs(body):
        _, ms = find_matches(p, pattern)
        hits.extend((p, m) for m in ms)
    safe = [(p, m) for p, m in hits if not m.unsafe]
    if not hits:
        raise OpError(f"placeholder not found: {placeholder!r}")
    if not safe:
        raise OpError(f"placeholder {placeholder!r} found but {hits[0][1].unsafe}")
    if len(hits) > 1:
        ctx.warn(f"placeholder occurs {len(hits)} times; only the first was replaced")
    p, m = safe[0]
    r, offset = apply(m, "")
    rPr = r.find(qn("w:rPr"))
    t = m.atoms[0].t
    instr = {"text": f"REF {bookmark} \\h", "page": f"PAGEREF {bookmark} \\h",
             "above_below": f"REF {bookmark} \\p \\h"}[display]
    new_runs = field_runs(instr, "", rPr=copy.deepcopy(rPr) if rPr is not None else None)
    insert_at(r, t, offset, new_runs)
    cleanup_empty(m.atoms[1:])

    begin_fc = new_runs[0].find(qn("w:fldChar"))
    if display == "page":
        set_update_fields(doc)
        ctx.details["page_number"] = "blank until fields are updated (w:updateFields is set)"
    else:
        refresh_refs(body)
    f = next(f for f in iter_fields(p) if f.begin is begin_fc)
    result = "".join(x.text or "" for rr in field_result_runs(f) for x in rr.iter(qn("w:t")))
    ctx.details.update({"target_kind": kind, "bookmark": bookmark, "field": instr,
                        "result": result})
    ctx.changed = True
    finish(ctx, doc, path, out)
