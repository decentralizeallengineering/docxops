"""comments_list: comments with author, date, anchored text, replies and resolved state."""

from __future__ import annotations

from docx.oxml.ns import qn
from lxml import etree

from ._core import PATH, Ctx, load, obj, operation
from ._xml import para_text

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"


def _part_xml(doc, suffix: str):
    for rel in doc.part.rels.values():
        if rel.is_external or not rel.reltype.endswith(suffix):
            continue
        part = rel.target_part
        el = getattr(part, "element", None)
        if el is not None:
            return el
        return etree.fromstring(part.blob)
    return None


def _anchors(body) -> dict[str, str]:
    """Text between commentRangeStart and commentRangeEnd for each comment id."""
    active: dict[str, list[str]] = {}
    done: dict[str, str] = {}
    for el in body.iter(qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:t"),
                        qn("w:p")):
        if el.tag == qn("w:commentRangeStart"):
            active[el.get(qn("w:id"))] = []
        elif el.tag == qn("w:commentRangeEnd"):
            cid = el.get(qn("w:id"))
            if cid in active:
                done[cid] = "".join(active.pop(cid)).strip()
        elif el.tag == qn("w:p"):
            for parts in active.values():
                if parts:
                    parts.append("\n")
        else:
            for parts in active.values():
                parts.append(el.text or "")
    return done


@operation(obj(
    "List comments: id, author, initials, date, text, the commented text, reply_to and "
    "resolved (when the file records it).",
    {"path": PATH},
))
def comments_list(ctx: Ctx, path: str) -> None:
    doc = load(path)
    root = _part_xml(doc, "/comments")
    if root is None:
        ctx.details.update({"comments": [], "count": 0})
        return
    ext = _part_xml(doc, "/commentsExtended")
    state: dict[str, dict] = {}
    if ext is not None:
        for c in ext.iter(f"{{{W15}}}commentEx"):
            state[c.get(f"{{{W15}}}paraId")] = {
                "done": c.get(f"{{{W15}}}done") in ("1", "true"),
                "parent": c.get(f"{{{W15}}}paraIdParent"),
            }
    anchors = _anchors(doc.element.body)
    by_para: dict[str, str] = {}
    out = []
    for c in root.findall(qn("w:comment")):
        cid = c.get(qn("w:id"))
        paras = c.findall(qn("w:p"))
        last_para_id = paras[-1].get(f"{{{W14}}}paraId") if paras else None
        if last_para_id:
            by_para[last_para_id] = cid
        st = state.get(last_para_id, {})
        out.append({
            "id": cid,
            "author": c.get(qn("w:author")),
            "initials": c.get(qn("w:initials")),
            "date": c.get(qn("w:date")),
            "text": "\n".join(para_text(p) for p in paras).strip(),
            "anchor_text": anchors.get(cid),
            "resolved": st.get("done") if st else None,
            "_parent": st.get("parent"),
        })
    for item in out:
        parent = item.pop("_parent")
        item["reply_to"] = by_para.get(parent) if parent else None
    ctx.details.update({"comments": out, "count": len(out),
                        "resolved_known": ext is not None})
