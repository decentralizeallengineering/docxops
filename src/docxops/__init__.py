"""docxops: Word (.docx) structure operations for LLM agents.

Every public operation is a plain function returning a JSON-serializable report
``{"ok", "changed", "warnings", "details"}``; none raises for user errors.
``SCHEMAS[name]`` holds the JSON Schema of each operation's arguments.
"""

from ._core import OPERATIONS, SCHEMAS
from .comments import comments_list
from .compose import compose
from .layout import header_footer, page_setup, section_break
from .refs import add_caption, cross_reference
from .structure import add_heading, add_paragraph_after, outline, set_heading_numbering, styles_apply
from .tables import add_table, table_from_csv
from .text import replace_text
from .toc import add_toc

__version__ = "0.1.0"


def call(name: str, args: dict | None = None) -> dict:
    """Dispatch a tool call by name, e.g. from an LLM tool-use block."""
    fn = OPERATIONS.get(name)
    if fn is None:
        return {"ok": False, "changed": False, "warnings": [],
                "details": {"error": f"unknown operation {name!r}",
                            "operations": sorted(OPERATIONS)}}
    if not isinstance(args, dict):
        args = {} if args is None else args
        if not isinstance(args, dict):
            return {"ok": False, "changed": False, "warnings": [],
                    "details": {"error": "args must be an object"}}
    return fn(**args)


def tool_definitions() -> list[dict]:
    """Operations as tool definitions: name, description, input_schema."""
    return [{"name": n, "description": s.get("description", ""),
             "input_schema": {k: v for k, v in s.items() if k != "description"}}
            for n, s in sorted(SCHEMAS.items())]


__all__ = [
    "OPERATIONS", "SCHEMAS", "call", "tool_definitions",
    "add_toc", "outline", "add_heading", "set_heading_numbering", "add_paragraph_after",
    "replace_text", "add_table", "table_from_csv", "add_caption", "cross_reference",
    "header_footer", "section_break", "page_setup", "compose", "styles_apply",
    "comments_list",
]
