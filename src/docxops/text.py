"""replace_text: find/replace across runs without disturbing untouched formatting."""

from __future__ import annotations

from docx.oxml.ns import qn

from ._core import OUT, PATH, Ctx, finish, load, obj, operation
from ._text import apply, cleanup_empty, compile_pattern, find_matches, paragraph_atoms

HEADER_FOOTER = ("/header", "/footer")


def text_roots(doc, scope: str) -> list[tuple[str, object]]:
    roots = [("body", doc.element.body)]
    if scope == "all":
        seen = set()
        for rel in doc.part.rels.values():
            if rel.is_external or not rel.reltype.endswith(HEADER_FOOTER):
                continue
            part = rel.target_part
            if part.partname in seen:
                continue
            seen.add(part.partname)
            roots.append((str(part.partname).rsplit("/", 1)[-1], part.element))
    return roots


@operation(obj(
    "Replace literal text, also across runs; untouched runs keep their format. Unsafe "
    "matches (across tabs/fields/links/revisions) are skipped and reported. all=+headers/footers.",
    {"path": PATH, "find": {"type": "string", "minLength": 1}, "replace": {"type": "string"},
     "match_case": {"type": "boolean"}, "whole_word": {"type": "boolean"},
     "scope": {"enum": ["body", "all"]}, "max_replacements": {"type": "integer", "minimum": 1},
     "out": OUT},
    ["path", "find", "replace"],
))
def replace_text(ctx: Ctx, path: str, find: str, replace: str, match_case: bool = True,
                 whole_word: bool = False, scope: str = "body",
                 max_replacements: int | None = None, out: str | None = None) -> None:
    doc = load(path)
    pattern = compile_pattern(find, match_case, whole_word)
    done = mixed = 0
    skipped: list[dict] = []
    limit_hit = False
    for where, root in text_roots(doc, scope):
        for p in root.iter(qn("w:p")):
            text, matches = find_matches(p, pattern)
            if not matches:
                continue
            todo = []
            for m in matches:
                if m.unsafe:
                    skipped.append({"where": where, "match": text[m.start:m.end],
                                                    "context": text[max(0, m.start - 30):m.end + 30],
                                    "reason": m.unsafe})
                    continue
                if max_replacements is not None and done + len(todo) >= max_replacements:
                    limit_hit = True
                    break
                todo.append(m)
            _, atoms = paragraph_atoms(p)
            for m in reversed(todo):
                apply(m, replace)
                mixed += m.mixed
            done += len(todo)
            cleanup_empty(atoms)
    ctx.details.update({"replaced": done, "skipped": skipped[:100],
                        "skipped_count": len(skipped), "mixed_formatting": mixed})
    if skipped:
        ctx.warn(f"{len(skipped)} match(es) not replaced because they could not be done "
                 "safely; see details.skipped")
    if mixed:
        ctx.warn(f"{mixed} replacement(s) spanned runs with different formatting and took "
                 "the formatting of the first run")
    if limit_hit:
        ctx.warn(f"stopped at max_replacements={max_replacements}")
    if done == 0 and not skipped:
        ctx.warn(f"no matches for {find!r}")
    ctx.changed = done > 0
    finish(ctx, doc, path, out)
