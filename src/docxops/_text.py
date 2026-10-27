"""Cross-run text matching that only edits w:t content it can change safely."""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass

from docx.oxml.ns import qn
from lxml import etree

from ._xml import set_text

# Containers we descend into while reading a paragraph's text.
_DESCEND = {qn("w:" + n) for n in (
    "hyperlink", "ins", "smartTag", "sdt", "sdtContent", "fldSimple", "customXml",
    "moveTo", "dir", "bdo",
)}
# Inert markers that do not break a match.
_IGNORE = {qn("w:" + n) for n in (
    "pPr", "bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd",
    "commentRangeStart", "commentRangeEnd", "sdtPr", "sdtEndPr", "customXmlPr",
    "smartTagPr", "lastRenderedPageBreak", "rPr",
)}


@dataclass
class Atom:
    start: int
    end: int
    t: object
    run: object


@dataclass
class Match:
    start: int
    end: int
    atoms: list
    unsafe: str | None
    mixed: bool


def paragraph_layout(p) -> tuple[str, list[Atom], set[int], set[int]]:
    """Paragraph text plus where it cannot be edited across.

    Returns (text, atoms, char_barriers, cuts): tabs and breaks appear in text as
    \t / \n and are listed in char_barriers; fields, drawings and other objects
    are zero-width cuts between two positions.
    """
    parts: list[str] = []
    atoms: list[Atom] = []
    chars: set[int] = set()
    cuts: set[int] = set()
    pos = 0

    def barrier(el=None) -> None:
        nonlocal pos
        name = el.tag.rsplit("}", 1)[-1] if el is not None else ""
        if name in ("tab", "ptab"):
            ch = "\t"
        elif name in ("br", "cr"):
            ch = "\n"
        else:
            cuts.add(pos)
            return
        chars.add(pos)
        parts.append(ch)
        pos += 1

    def walk(el) -> None:
        nonlocal pos
        for child in el:
            tag = child.tag
            if not isinstance(tag, str) or tag in _IGNORE:
                continue
            if tag == qn("w:r"):
                for rc in child:
                    if rc.tag == qn("w:t"):
                        text = rc.text or ""
                        atoms.append(Atom(pos, pos + len(text), rc, child))
                        parts.append(text)
                        pos += len(text)
                    elif rc.tag in _IGNORE:
                        continue
                    else:
                        barrier(rc)
            elif tag == qn("w:fldSimple"):
                cuts.add(pos)  # its result is generated text; don't edit across it
                walk(child)
                cuts.add(pos)
            elif tag in _DESCEND:
                walk(child)
            else:
                barrier(child)

    walk(p)
    return "".join(parts), atoms, chars, cuts


def paragraph_atoms(p) -> tuple[str, list[Atom]]:
    text, atoms, _, _ = paragraph_layout(p)
    return text, atoms


def compile_pattern(find: str, match_case: bool, whole_word: bool):
    pat = re.escape(find)
    if whole_word:
        pat = r"(?<!\w)" + pat + r"(?!\w)"
    return re.compile(pat, 0 if match_case else re.IGNORECASE)


def _rpr_key(run) -> bytes:
    rPr = run.find(qn("w:rPr"))
    if rPr is None:
        return b""
    rPr = copy.deepcopy(rPr)
    for el in rPr.iter():
        for a in list(el.attrib):
            if "rsid" in a:
                del el.attrib[a]
    return etree.tostring(rPr, method="c14n")


def find_matches(p, pattern) -> tuple[str, list[Match]]:
    text, atoms, chars, cuts = paragraph_layout(p)
    out = []
    for m in pattern.finditer(text):
        s, e = m.span()
        if s == e:
            continue
        span = [a for a in atoms if a.start < e and a.end > s]
        unsafe = None
        if any(s <= c < e for c in chars):
            unsafe = "crosses a tab or line break"
        elif any(s < c < e for c in cuts):
            unsafe = "crosses a field, drawing or other object"
        elif len({id(a.run.getparent()) for a in span}) > 1:
            unsafe = "crosses a hyperlink, revision or field boundary"
        mixed = len({_rpr_key(a.run) for a in span}) > 1
        out.append(Match(s, e, span, unsafe, mixed))
    return text, out


def apply(match: Match, replacement: str) -> tuple:
    """Replace a safe match. Returns (run, offset) where the replacement starts."""
    first = match.atoms[0]
    for i, a in enumerate(match.atoms):
        t = a.t
        cur = t.text or ""
        lo = max(match.start - a.start, 0)
        hi = min(match.end - a.start, len(cur))
        if i == 0:
            set_text(t, cur[:lo] + replacement + cur[hi:])
        else:
            set_text(t, cur[:lo] + cur[hi:])
    return first.run, match.start - first.start


def cleanup_empty(atoms: list[Atom]) -> None:
    """Drop w:t left empty by replacements, and runs left with nothing but rPr."""
    for a in atoms:
        t, r = a.t, a.run
        if (t.text or "") == "" and t.getparent() is r:
            r.remove(t)
            if all(c.tag == qn("w:rPr") for c in r) and r.getparent() is not None:
                r.getparent().remove(r)


def insert_at(run, t, offset: int, new_elements: list) -> None:
    """Insert elements (runs) into the paragraph at character offset within t."""
    text = t.text or ""
    before, after = text[:offset], text[offset:]
    children = list(run)
    tail = children[children.index(t) + 1:]
    right = None
    if after or tail:
        right = copy.deepcopy(run)
        for c in list(right):
            if c.tag != qn("w:rPr"):
                right.remove(c)
        if after:
            nt = copy.deepcopy(t)
            set_text(nt, after)
            right.append(nt)
        for c in tail:
            right.append(c)
    set_text(t, before)
    if not before:
        run.remove(t)
    anchor = run
    for el in new_elements:
        anchor.addnext(el)
        anchor = el
    if right is not None:
        anchor.addnext(right)
    if all(c.tag == qn("w:rPr") for c in run):
        run.getparent().remove(run)
