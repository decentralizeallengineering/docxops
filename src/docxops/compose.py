"""compose: append documents with docxcompose (styles, numbering, images carried over)."""

from __future__ import annotations

from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn

from ._core import OUT, PATH, Ctx, OpError, load, obj, operation, save


def _counts(doc) -> dict:
    numbering = 0
    try:
        numbering = len(doc.part.numbering_part.element.findall(qn("w:num")))
    except Exception:
        pass
    return {"paragraphs": len(doc.paragraphs), "tables": len(doc.tables),
            "styles": len(doc.styles.element.findall(qn("w:style"))), "lists": numbering}


@operation(obj(
    "Append documents to path (docxcompose keeps their styles, numbering and images). "
    "separator between documents: page_break, section_break or none.",
    {"path": PATH, "append": {"type": "array", "minItems": 1, "items": PATH},
     "separator": {"enum": ["page_break", "section_break", "none"]}, "out": OUT},
    ["path", "append"],
))
def compose(ctx: Ctx, path: str, append: list[str], separator: str = "page_break",
            out: str | None = None) -> None:
    from docxcompose.composer import Composer

    master = load(path)
    others = []
    for p in append:
        try:
            others.append((p, load(p)))
        except OpError as exc:
            raise OpError(f"cannot append {p}: {exc}") from None
    before = _counts(master)
    composer = Composer(master)
    appended = []
    for p, d in others:
        if separator == "page_break":
            composer.doc.add_page_break()
        elif separator == "section_break":
            composer.doc.add_section(WD_SECTION.NEW_PAGE)
        composer.append(d)
        appended.append({"path": p, **_counts(d)})
    ctx.details.update({"appended": appended, "before": before,
                        "after": _counts(composer.doc), "separator": separator})
    ctx.changed = True
    save(ctx, composer.doc, path, out)
