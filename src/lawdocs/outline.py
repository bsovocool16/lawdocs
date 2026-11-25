"""A structured outline of a document: the table of contents a lawyer keeps in their head.

The outline lists every clause and heading with its full reference
("1.1(a)"), its caption or opening words, and the block id that holds it. An
agent reads the outline first and then jumps straight to the clause it needs,
instead of grepping through boilerplate.
"""

from __future__ import annotations

from .model import Document
from .numbering import apply


def _gist(text: str, words: int = 12) -> str:
    parts = text.split()
    return " ".join(parts[:words]) + (" …" if len(parts) > words else "")


def build(doc: Document) -> list[dict]:
    refs = apply(doc)
    entries: list[dict] = []
    for b in doc.blocks:
        if b.kind == "clause":
            entries.append({
                "block": b.id,
                "ref": refs[b.id],
                "level": b.level,
                "caption": b.caption,
                "gist": _gist(b.text) if b.text else None,
            })
        elif b.kind in ("heading", "title"):
            entries.append({"block": b.id, "ref": None, "level": 0, "caption": b.text, "gist": None, "kind": b.kind})
    return entries


def to_markdown(doc: Document, max_level: int = 6) -> str:
    lines = []
    for e in build(doc):
        if e["ref"] is None:
            lines.append(f"{e['caption']}  `{e['block']}`")
            continue
        if e["level"] > max_level:
            continue
        indent = "  " * (e["level"] - 1)
        label = f"**{e['caption']}**" + (f" — {e['gist']}" if e["gist"] else "") if e["caption"] else (e["gist"] or "")
        lines.append(f"{indent}- {e['ref']} {label}  `{e['block']}`")
    return "\n".join(lines) + "\n"
