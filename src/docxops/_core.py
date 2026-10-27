"""Report shape, the @operation decorator, and load/save helpers."""

from __future__ import annotations

import functools
import inspect
import json
import os
import tempfile
from pathlib import Path
from typing import Any, Callable

import jsonschema
from docx import Document

SCHEMAS: dict[str, dict] = {}
OPERATIONS: dict[str, Callable[..., dict]] = {}

MAX_SCHEMA_CHARS = 600


class OpError(Exception):
    """A user error: bad input, missing target, unreadable file. Becomes ok=False."""


class Ctx:
    """Collects warnings and details while an operation runs."""

    def __init__(self) -> None:
        self.warnings: list[str] = []
        self.details: dict[str, Any] = {}
        self.changed = False

    def warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)


def report(ok: bool, changed: bool, warnings: list[str], details: dict) -> dict:
    return {"ok": ok, "changed": changed, "warnings": list(warnings), "details": details}


def _jsonable(value: Any) -> Any:
    # Round-trip through json so callers always get plain types.
    return json.loads(json.dumps(value, default=str))


def operation(schema: dict) -> Callable[[Callable[..., None]], Callable[..., dict]]:
    """Register a public operation.

    The wrapped function receives a Ctx as its first argument and fills it in.
    Arguments are validated against the schema (None values are treated as
    omitted). Any exception becomes an ok=False report; nothing escapes.
    """

    def deco(fn: Callable[..., None]) -> Callable[..., dict]:
        name = fn.__name__
        size = len(json.dumps(schema, separators=(",", ":")))
        if size > MAX_SCHEMA_CHARS:  # pragma: no cover - guarded by tests too
            raise RuntimeError(f"schema for {name} is {size} chars")
        sig = inspect.signature(fn)
        public_params = list(sig.parameters.values())[1:]
        public_sig = sig.replace(parameters=public_params)

        @functools.wraps(fn)
        def wrapper(*args: Any, **kwargs: Any) -> dict:
            ctx = Ctx()
            try:
                try:
                    bound = public_sig.bind(*args, **kwargs)
                except TypeError as exc:
                    raise OpError(f"bad arguments: {exc}") from None
                bound.apply_defaults()
                payload = {k: v for k, v in bound.arguments.items() if v is not None}
                payload = _jsonable(payload)
                errors = sorted(
                    jsonschema.Draft202012Validator(schema).iter_errors(payload),
                    key=lambda e: list(e.path),
                )
                if errors:
                    msgs = []
                    for e in errors[:5]:
                        where = "/".join(str(p) for p in e.path) or "(root)"
                        msgs.append(f"{where}: {e.message}")
                    raise OpError("invalid arguments: " + "; ".join(msgs))
                fn(ctx, *bound.args, **bound.kwargs)
                return _jsonable(report(True, ctx.changed, ctx.warnings, ctx.details))
            except OpError as exc:
                details = dict(ctx.details)
                details["error"] = str(exc)
                return _jsonable(report(False, False, ctx.warnings, details))
            except Exception as exc:  # never raise to the agent
                details = dict(ctx.details)
                details["error"] = f"{type(exc).__name__}: {exc}"
                details["error_type"] = "internal"
                return _jsonable(report(False, False, ctx.warnings, details))

        wrapper.__signature__ = public_sig  # type: ignore[attr-defined]
        SCHEMAS[name] = schema
        OPERATIONS[name] = wrapper
        return wrapper

    return deco


def load(path: str):
    p = Path(path)
    if not p.exists():
        raise OpError(f"file not found: {path}")
    if not p.is_file():
        raise OpError(f"not a file: {path}")
    if p.suffix.lower() not in (".docx", ".docm", ".dotx", ".dotm"):
        raise OpError(f"not a Word file (expected .docx): {path}")
    try:
        return Document(str(p))
    except Exception as exc:
        raise OpError(f"cannot open as .docx: {type(exc).__name__}: {exc}") from None


def save(ctx: Ctx, doc, path: str, out: str | None) -> None:
    """Save atomically to out (or back to path) and record where."""
    target = Path(out) if out else Path(path)
    if not target.parent.exists():
        raise OpError(f"output folder does not exist: {target.parent}")
    fd, tmp = tempfile.mkstemp(prefix=".docxops-", suffix=".docx", dir=str(target.parent))
    os.close(fd)
    try:
        doc.save(tmp)
        try:
            os.replace(tmp, target)
        except PermissionError:
            raise OpError(f"cannot write {target} (is it open in Word?)") from None
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    ctx.details["saved_to"] = str(target)


def finish(ctx: Ctx, doc, path: str, out: str | None) -> None:
    """Save when something changed, or when an explicit out was requested."""
    if ctx.changed or out:
        save(ctx, doc, path, out)


# Shared schema fragments. Keep these small: every schema must fit 600 chars.
PATH = {"type": "string", "minLength": 1}
OUT = {"type": "string"}  # save-as path; default overwrites path
TEXT = {"type": "string"}


def obj(desc: str, props: dict, required: list[str] | None = None) -> dict:
    return {
        "description": desc,
        "type": "object",
        "properties": props,
        "required": required or ["path"],
        "additionalProperties": False,
    }
