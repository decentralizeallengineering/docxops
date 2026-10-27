"""Low-level OOXML helpers: element order, fields, bookmarks, styles, headings."""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from ._core import OpError

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
V_NS = "urn:schemas-microsoft-com:vml"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Schema sequence orders (ECMA-376) for children we insert into.
PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
    "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
    "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
    "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
    "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
    "em", "lang", "eastAsianLayout", "specVanish", "oMath",
]
STYLE_ORDER = [
    "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
    "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal",
    "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr",
    "tcPr", "tblStylePr",
]
# Elements that may follow w:updateFields inside w:settings.
SETTINGS_AFTER_UPDATEFIELDS = [
    "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids",
    "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
    "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade",
    "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary",
    "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
]


def local(el) -> str:
    tag = el.tag
    return tag.split("}", 1)[1] if isinstance(tag, str) and "}" in tag else str(tag)


def is_w(el, name: str) -> bool:
    return el.tag == qn("w:" + name)


def make(tag: str, **attrs) -> object:
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn("w:" + k), str(v))
    return el


def insert_ordered(parent, child, order: list[str]):
    """Insert child into parent respecting a schema sequence; replaces a same-tag child."""
    name = local(child)
    existing = parent.find(qn("w:" + name))
    if existing is not None:
        existing.addprevious(child)
        parent.remove(existing)
        return child
    idx = order.index(name)
    later = {qn("w:" + n) for n in order[idx + 1:]}
    for c in parent:
        if c.tag in later:
            c.addprevious(child)
            return child
    parent.append(child)
    return child


def get_or_add_pPr(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = make("w:pPr")
        p.insert(0, pPr)
    return pPr


def get_or_add_rPr(r):
    rPr = r.find(qn("w:rPr"))
    if rPr is None:
        rPr = make("w:rPr")
        r.insert(0, rPr)
    return rPr


def set_text(t, text: str) -> None:
    t.text = text
    if text != text.strip() or "  " in text:
        t.set(XML_SPACE, "preserve")


def run(text: str = "", rPr=None, bold: bool = False):
    r = make("w:r")
    if rPr is not None:
        r.append(copy.deepcopy(rPr))
    if bold:
        insert_ordered(get_or_add_rPr(r), make("w:b"), RPR_ORDER)
    if text:
        t = make("w:t")
        set_text(t, text)
        r.append(t)
    return r


def paragraph(text: str = "", style_id: str | None = None):
    p = make("w:p")
    if style_id:
        pPr = make("w:pPr")
        pPr.append(make("w:pStyle", val=style_id))
        p.append(pPr)
    if text:
        p.append(run(text))
    return p


# ---------------------------------------------------------------- fields

def field_runs(instr: str, result: str = "", rPr=None) -> list:
    """Runs for a complex field: begin, instr, separate, [result], end."""
    begin = run(rPr=rPr)
    begin.append(make("w:fldChar", fldCharType="begin"))
    ins = run(rPr=rPr)
    it = make("w:instrText")
    it.text = f" {instr.strip()} "
    it.set(XML_SPACE, "preserve")
    ins.append(it)
    sep = run(rPr=rPr)
    sep.append(make("w:fldChar", fldCharType="separate"))
    runs = [begin, ins, sep]
    if result:
        runs.append(run(result, rPr=rPr))
    end = run(rPr=rPr)
    end.append(make("w:fldChar", fldCharType="end"))
    runs.append(end)
    return runs


@dataclass
class Field:
    instr: str = ""
    begin: object = None  # w:fldChar begin (or w:fldSimple)
    sep: object = None
    end: object = None
    simple: bool = False

    @property
    def code(self) -> str:
        return self.instr.strip().split(" ", 1)[0].upper() if self.instr.strip() else ""


def iter_fields(root) -> list[Field]:
    """All complex and simple fields under root, in document order (by begin)."""
    found: list[tuple[int, Field]] = []
    stack: list[tuple[int, Field]] = []
    order = 0
    tags = (qn("w:fldChar"), qn("w:instrText"), qn("w:fldSimple"))
    for el in root.iter(*tags):
        order += 1
        if el.tag == qn("w:fldSimple"):
            found.append((order, Field(el.get(qn("w:instr")) or "", el, None, el, True)))
            continue
        if el.tag == qn("w:instrText"):
            if stack:
                stack[-1][1].instr += el.text or ""
            continue
        kind = el.get(qn("w:fldCharType"))
        if kind == "begin":
            stack.append((order, Field(begin=el)))
        elif kind == "separate" and stack:
            stack[-1][1].sep = el
        elif kind == "end" and stack:
            pos, f = stack.pop()
            f.end = el
            found.append((pos, f))
    found.sort(key=lambda x: x[0])
    return [f for _, f in found]


def field_result_runs(f: Field) -> list:
    """w:r elements forming the result of a complex field within one paragraph."""
    if f.simple or f.sep is None:
        return [] if not f.simple else list(f.begin.iter(qn("w:r")))
    sep_run = f.sep.getparent()
    end_run = f.end.getparent()
    out = []
    node = sep_run.getnext()
    while node is not None and node is not end_run:
        if is_w(node, "r"):
            out.append(node)
        else:
            out.extend(node.iter(qn("w:r")))
        node = node.getnext()
    return out


def set_field_result(f: Field, text: str) -> None:
    """Replace a field's displayed result with text, keeping the first run's format."""
    if f.simple:
        runs = list(f.begin.iter(qn("w:r")))
        rPr = runs[0].find(qn("w:rPr")) if runs else None
        for r in runs:
            r.getparent().remove(r)
        f.begin.append(run(text, rPr=rPr))
        return
    if f.sep is None:
        return
    runs = field_result_runs(f)
    ts = [t for r in runs for t in r.iter(qn("w:t"))]
    if ts:
        set_text(ts[0], text)
        for t in ts[1:]:
            t.getparent().remove(t)
    else:
        sep_run = f.sep.getparent()
        rPr = sep_run.find(qn("w:rPr"))
        sep_run.addnext(run(text, rPr=rPr))


def instr_has_switch(instr: str, switch: str) -> bool:
    return re.search(r"(?i)\\" + re.escape(switch) + r"\b", instr) is not None


# ------------------------------------------------------------- bookmarks

def bookmark_names(root) -> set[str]:
    return {b.get(qn("w:name")) for b in root.iter(qn("w:bookmarkStart"))}


class Bookmarks:
    """Allocates unique bookmark ids/names for a document part."""

    def __init__(self, root) -> None:
        ids = [0]
        for tag in ("w:bookmarkStart", "w:bookmarkEnd"):
            for b in root.iter(qn(tag)):
                try:
                    ids.append(int(b.get(qn("w:id"))))
                except (TypeError, ValueError):
                    pass
        self.next_id = max(ids) + 1
        self.names = bookmark_names(root)

    def new_name(self, prefix: str) -> str:
        nums = [int(n[len(prefix):]) for n in self.names
                if n and n.startswith(prefix) and n[len(prefix):].isdigit()]
        n = max(nums) + 1 if nums else 100000001
        name = f"{prefix}{n}"
        self.names.add(name)
        return name

    def wrap(self, p, name: str, first=None, last=None) -> str:
        """Bookmark the paragraph content from first to last (default: all of it)."""
        bid = str(self.next_id)
        self.next_id += 1
        start = make("w:bookmarkStart", id=bid, name=name)
        end = make("w:bookmarkEnd", id=bid)
        if first is None:
            pPr = p.find(qn("w:pPr"))
            if pPr is not None:
                pPr.addnext(start)
            else:
                p.insert(0, start)
        else:
            first.addprevious(start)
        if last is None:
            p.append(end)
        else:
            last.addnext(end)
        return name


def paragraph_bookmarks(p, prefix: str = "") -> list[str]:
    return [b.get(qn("w:name")) for b in p.findall(qn("w:bookmarkStart"))
            if (b.get(qn("w:name")) or "").startswith(prefix)]


def bookmark_text(root, name: str) -> str | None:
    """Visible text inside a bookmark (single paragraph scope)."""
    for b in root.iter(qn("w:bookmarkStart")):
        if b.get(qn("w:name")) != name:
            continue
        bid = b.get(qn("w:id"))
        p = b.getparent()
        while p is not None and not is_w(p, "p"):
            p = p.getparent()
        if p is None:
            return None
        parts, inside = [], False
        for el in p.iter():
            if el is b:
                inside = True
            elif el.tag == qn("w:bookmarkEnd") and el.get(qn("w:id")) == bid:
                break
            elif inside and el.tag == qn("w:t"):
                parts.append(el.text or "")
        return "".join(parts)
    return None


# ----------------------------------------------------------------- text

def para_text(p) -> str:
    """Visible text of a paragraph: w:t and tabs, skipping deleted text and nested paragraphs."""
    parts = []
    for el in p.iter(qn("w:t"), qn("w:tab"), qn("w:br")):
        anc = el.getparent()
        skip = False
        while anc is not None and anc is not p:
            if anc.tag in (qn("w:del"), qn("w:txbxContent"), qn("w:moveFrom")):
                skip = True
                break
            anc = anc.getparent()
        if skip:
            continue
        if el.tag == qn("w:t"):
            parts.append(el.text or "")
        elif el.tag == qn("w:tab") and is_w(el.getparent(), "r"):
            parts.append("\t")
        elif el.tag == qn("w:br"):
            parts.append("\n")
    return "".join(parts)


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().casefold()


def has_ancestor(el, *names: str) -> bool:
    tags = {qn("w:" + n) for n in names}
    a = el.getparent()
    while a is not None:
        if a.tag in tags:
            return True
        a = a.getparent()
    return False


def body_paragraphs(body) -> list:
    """All paragraphs in body order, including tables and content controls, not text boxes."""
    return [p for p in body.iter(qn("w:p")) if not has_ancestor(p, "txbxContent")]


# --------------------------------------------------------------- styles

class Styles:
    def __init__(self, doc) -> None:
        self.doc = doc
        self.root = doc.styles.element
        self.refresh()

    def refresh(self) -> None:
        self.by_id: dict[str, object] = {}
        self.by_name: dict[str, object] = {}
        self.default_para = None
        for s in self.root.findall(qn("w:style")):
            sid = s.get(qn("w:styleId"))
            self.by_id[sid] = s
            n = s.find(qn("w:name"))
            if n is not None:
                self.by_name.setdefault((n.get(qn("w:val")) or "").casefold(), s)
            if s.get(qn("w:type")) == "paragraph" and s.get(qn("w:default")) in ("1", "true"):
                self.default_para = s

    def name_of(self, style_id: str | None) -> str:
        s = self.by_id.get(style_id) if style_id else self.default_para
        if s is None:
            return style_id or "Normal"
        n = s.find(qn("w:name"))
        return ui_name(n.get(qn("w:val")) if n is not None else style_id)

    def find(self, name: str, stype: str = "paragraph"):
        """Find a style by UI name, internal name, or style id."""
        key = name.casefold()
        for cand in (self.by_name.get(key), self.by_id.get(name),
                     self.by_name.get(INTERNAL_NAMES.get(key, key))):
            if cand is not None and cand.get(qn("w:type")) == stype:
                return cand
        return None

    def para_style_id(self, p) -> str | None:
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            return None
        ps = pPr.find(qn("w:pStyle"))
        return ps.get(qn("w:val")) if ps is not None else None

    def para_style_name(self, p) -> str:
        return self.name_of(self.para_style_id(p))

    def outline_level(self, p) -> int | None:
        """1..9 for heading-like paragraphs, None for body text."""
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            ol = pPr.find(qn("w:outlineLvl"))
            if ol is not None:
                return _lvl(ol)
        sid = self.para_style_id(p)
        s = self.by_id.get(sid) if sid else self.default_para
        seen = set()
        while s is not None and id(s) not in seen:
            seen.add(id(s))
            spPr = s.find(qn("w:pPr"))
            if spPr is not None:
                ol = spPr.find(qn("w:outlineLvl"))
                if ol is not None:
                    return _lvl(ol)
            n = s.find(qn("w:name"))
            m = re.fullmatch(r"(?i)heading ([1-9])", (n.get(qn("w:val")) if n is not None else "") or "")
            if m:
                return int(m.group(1))
            b = s.find(qn("w:basedOn"))
            s = self.by_id.get(b.get(qn("w:val"))) if b is not None else None
        return None

    def ensure(self, name: str, ctx=None) -> str:
        """Return the style id for a paragraph style, creating known built-ins."""
        s = self.find(name)
        if s is not None:
            return s.get(qn("w:styleId"))
        spec = builtin_spec(name)
        if spec is None:
            raise OpError(f"paragraph style not found: {name!r}")
        if spec.get("based") and spec["based"] != "Normal":
            self.ensure(ui_name_from_id(spec["based"]))
        el = _style_xml(spec, self.find("Normal") is not None)
        self.root.append(el)
        self.refresh()
        if ctx is not None:
            ctx.warn(f"style {spec['ui']!r} was missing; added a built-in definition")
        return spec["id"]


def suppress_inherited_numbering(styles: "Styles", style_id: str) -> bool:
    """Give a style numId=0 when it would inherit list numbering from its base
    style (TOC Heading is based on Heading 1, which may be numbered)."""
    s = styles.by_id.get(style_id)
    if s is None:
        return False
    own = s.find(f"{qn('w:pPr')}/{qn('w:numPr')}")
    if own is not None:
        return False
    base, seen, inherited = s, set(), False
    while base is not None and id(base) not in seen:
        seen.add(id(base))
        b = base.find(qn("w:basedOn"))
        base = styles.by_id.get(b.get(qn("w:val"))) if b is not None else None
        if base is not None and base.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is not None:
            inherited = True
            break
    if not inherited:
        return False
    pPr = s.find(qn("w:pPr"))
    if pPr is None:
        pPr = insert_ordered(s, make("w:pPr"), STYLE_ORDER)
    numPr = make("w:numPr")
    numPr.append(make("w:numId", val=0))
    insert_ordered(pPr, numPr, PPR_ORDER)
    return True


def _lvl(ol) -> int | None:
    try:
        v = int(ol.get(qn("w:val")))
    except (TypeError, ValueError):
        return None
    return v + 1 if 0 <= v < 9 else None


# Word stores some built-in style names in lower case.
INTERNAL_NAMES = {f"heading {i}": f"heading {i}" for i in range(1, 10)}
INTERNAL_NAMES.update({f"toc {i}": f"toc {i}" for i in range(1, 10)})
INTERNAL_NAMES.update({"caption": "caption", "header": "header", "footer": "footer"})


def ui_name(internal: str) -> str:
    m = re.fullmatch(r"(?i)(heading|toc) ([1-9])", internal or "")
    if m:
        return ("Heading " if m.group(1).lower() == "heading" else "TOC ") + m.group(2)
    if internal in ("caption", "header", "footer"):
        return internal.capitalize()
    return internal


def ui_name_from_id(sid: str) -> str:
    m = re.fullmatch(r"(Heading|TOC)([1-9])", sid)
    if m:
        return f"{m.group(1)} {m.group(2)}"
    return {"TOCHeading": "TOC Heading"}.get(sid, sid)


BUILTIN_SIMPLE = {
    "title": ("Title", "Title", {"sz": 56}),
    "subtitle": ("Subtitle", "Subtitle", {"i": True, "sz": 30}),
    "quote": ("Quote", "Quote", {"i": True}),
    "intense quote": ("IntenseQuote", "Intense Quote", {"i": True, "b": True}),
    "no spacing": ("NoSpacing", "No Spacing", {}),
    "list paragraph": ("ListParagraph", "List Paragraph", {"ind": 720}),
    "list bullet": ("ListBullet", "List Bullet", {"ind": 360}),
    "list number": ("ListNumber", "List Number", {"ind": 360}),
    "body text": ("BodyText", "Body Text", {"after": 120}),
    "normal": ("Normal", "Normal", {}),
}

BUILTIN_STYLE_NAMES = sorted(
    [f"Heading {i}" for i in range(1, 10)] + ["Caption", "TOC Heading"]
    + [v[1] for v in BUILTIN_SIMPLE.values()]
)


def builtin_spec(name: str) -> dict | None:
    key = name.casefold().strip()
    m = re.fullmatch(r"heading ?([1-9])", key)
    if m:
        n = int(m.group(1))
        return {"id": f"Heading{n}", "ui": f"Heading {n}", "name": f"heading {n}",
                "based": "Normal", "outline": n - 1, "b": True,
                "sz": [32, 26, 24, 22, 22, 22, 22, 22, 22][n - 1],
                "before": 240 if n == 1 else 200, "keepNext": True, "prio": 9}
    m = re.fullmatch(r"toc ?([1-9])", key)
    if m:
        n = int(m.group(1))
        return {"id": f"TOC{n}", "ui": f"TOC {n}", "name": f"toc {n}", "based": "Normal",
                "after": 100, "ind": 220 * (n - 1), "prio": 39}
    if key == "toc heading":
        return {"id": "TOCHeading", "ui": "TOC Heading", "name": "TOC Heading",
                "based": "Heading1", "outline": 9, "prio": 39}
    if key == "caption":
        return {"id": "Caption", "ui": "Caption", "name": "caption", "based": "Normal",
                "b": True, "sz": 18, "after": 200, "prio": 35}
    if key in BUILTIN_SIMPLE:
        sid, ui, extra = BUILTIN_SIMPLE[key]
        spec = {"id": sid, "ui": ui, "name": ui, "based": None if sid == "Normal" else "Normal"}
        spec.update(extra)
        return spec
    return None


def _style_xml(spec: dict, has_normal: bool):
    s = make("w:style", type="paragraph", styleId=spec["id"])
    s.append(make("w:name", val=spec["name"]))
    if spec.get("based") and (has_normal or spec["based"] != "Normal"):
        s.append(make("w:basedOn", val=spec["based"]))
    s.append(make("w:next", val="Normal"))
    if "prio" in spec:
        s.append(make("w:uiPriority", val=spec["prio"]))
    s.append(make("w:qFormat"))
    pPr = make("w:pPr")
    if spec.get("keepNext"):
        insert_ordered(pPr, make("w:keepNext"), PPR_ORDER)
    if "before" in spec or "after" in spec:
        sp = make("w:spacing")
        if "before" in spec:
            sp.set(qn("w:before"), str(spec["before"]))
        if "after" in spec:
            sp.set(qn("w:after"), str(spec["after"]))
        insert_ordered(pPr, sp, PPR_ORDER)
    if spec.get("ind"):
        insert_ordered(pPr, make("w:ind", left=spec["ind"]), PPR_ORDER)
    if "outline" in spec:
        insert_ordered(pPr, make("w:outlineLvl", val=spec["outline"]), PPR_ORDER)
    if len(pPr):
        s.append(pPr)
    rPr = make("w:rPr")
    if spec.get("b"):
        insert_ordered(rPr, make("w:b"), RPR_ORDER)
    if spec.get("i"):
        insert_ordered(rPr, make("w:i"), RPR_ORDER)
    if spec.get("sz"):
        insert_ordered(rPr, make("w:sz", val=spec["sz"]), RPR_ORDER)
    if len(rPr):
        s.append(rPr)
    return s


# ------------------------------------------------------------- headings

@dataclass
class Heading:
    p: object
    level: int
    text: str
    style: str
    index: int  # position among body paragraphs
    children: list = field(default_factory=list)


def headings(doc, styles: Styles | None = None, skip=None) -> list[Heading]:
    styles = styles or Styles(doc)
    out = []
    for i, p in enumerate(body_paragraphs(doc.element.body)):
        if skip is not None and skip(p):
            continue
        lvl = styles.outline_level(p)
        if lvl is None:
            continue
        out.append(Heading(p, lvl, para_text(p).strip(), styles.para_style_name(p), i))
    return out


def find_heading(doc, text: str, occurrence: int = 1, ctx=None, styles=None) -> Heading:
    hs = headings(doc, styles)
    key = norm(text)
    if not key:
        raise OpError("heading text is empty")
    matches = [h for h in hs if norm(h.text) == key]
    if not matches:
        matches = [h for h in hs if key in norm(h.text)]
        if matches and ctx is not None:
            ctx.warn(f"no exact heading {text!r}; matched by substring")
    if not matches:
        sample = [h.text for h in hs[:15]]
        raise OpError(f"heading not found: {text!r}. Headings: {sample}")
    if occurrence > len(matches):
        raise OpError(f"heading {text!r} occurs {len(matches)} time(s); occurrence {occurrence} requested")
    if len(matches) > 1 and ctx is not None:
        ctx.warn(f"heading {text!r} matches {len(matches)} paragraphs; used occurrence {occurrence}")
    return matches[occurrence - 1]


def top_level_block(el, body):
    """The direct child of body that contains el."""
    node = el
    while node is not None and node.getparent() is not body:
        node = node.getparent()
    return node


def section_end_block(doc, h: Heading, styles: Styles):
    """Last body-level block belonging to heading h's section (before the next
    heading of the same or higher level, or before the body sectPr)."""
    body = doc.element.body
    start = top_level_block(h.p, body)
    last = start
    node = start.getnext()
    while node is not None and not is_w(node, "sectPr"):
        if is_w(node, "p"):
            lvl = styles.outline_level(node)
            if lvl is not None and lvl <= h.level:
                break
        last = node
        node = node.getnext()
    return last


def body_end_anchor(body):
    """Insert new body content before this element (the trailing sectPr) or append."""
    sect = body.find(qn("w:sectPr"))
    return sect


def append_to_body(body, el) -> None:
    sect = body_end_anchor(body)
    if sect is not None:
        sect.addprevious(el)
    else:
        body.append(el)


def set_update_fields(doc) -> bool:
    settings = doc.settings.element
    uf = settings.find(qn("w:updateFields"))
    if uf is not None:
        if uf.get(qn("w:val")) in ("true", "1", None):
            return False
        uf.set(qn("w:val"), "true")
        return True
    uf = make("w:updateFields", val="true")
    after = {qn("w:" + n) for n in SETTINGS_AFTER_UPDATEFIELDS}
    for c in settings:
        if c.tag in after:
            c.addprevious(uf)
            return True
    settings.append(uf)
    return True


def text_width_twips(doc) -> int:
    try:
        s = doc.sections[0]
        w = s.page_width.twips - s.left_margin.twips - s.right_margin.twips
        gutter = s.gutter.twips if s.gutter is not None else 0
        return max(1440, w - gutter)
    except Exception:
        return 9360
