"""Render the block model as Markdown that Claude Docs accepts.

Claude Docs has headings, paragraphs, lists and tables, but its ordered lists
only count 1, 2, 3 at every depth, so legal numbers are written into the text
of each clause. They are always regenerated from block levels by
``numbering.apply``, never edited by hand.

Presentation rules (the whole of them):

* title            ``# Title``
* level-1 clause   ``## 1. Caption`` and its text as a paragraph beneath
* deeper clause    one paragraph: ``1.1 **Caption.** Text`` / ``(a) Text``
* heading          ``## HEADING``
* text             a plain paragraph
* table            a pipe table
"""

from __future__ import annotations

import re

from .model import Block, Document, Run, note_marker

# Brackets are everywhere in legal drafts ("[Initial] Closing"), so they are
# not backslash-escaped: Claude Docs reads backslash-bracket as display math. Only
# the spots where brackets would form a link or footnote are broken up.
_SPECIAL = re.compile(r"([\\`*_<>|$~])")
_LINKISH = re.compile(r"\](?=[(\[:])")
_LEADING = re.compile(r"^(\s*)([#>+\-=]|\d+[.)](?=\s|$))")


def _escape(text: str) -> str:
    text = _SPECIAL.sub(r"\\\1", text)
    text = _LINKISH.sub(r"\\]", text)
    return text.replace("[^", "[\\^")


def _escape_start(md: str) -> str:
    """Stop a paragraph's first characters from turning it into a list, quote or heading."""
    m = _LEADING.match(md)
    if not m:
        return md
    lead, tok = m.group(1), m.group(2)
    if tok[-1] in ".)" and tok[:-1].isdigit():
        return lead + tok[:-1] + "\\" + tok[-1] + md[m.end():]
    return lead + "\\" + tok + md[m.end():]


def runs_md(runs: list[Run]) -> str:
    out = []
    for r in runs:
        t = _escape(r.text)
        if not t.strip():
            out.append(t)
            continue
        # Keep surrounding spaces and punctuation outside the markers: "**" next
        # to a bracket or quote often fails to parse as emphasis ("**[**initial").
        m = re.match(r"^([\s\[\]().,;:\"'\u201c\u201d\u2018\u2019]*)(.*?)([\s\[\]().,;:\"'\u201c\u201d\u2018\u2019]*)$", t, re.S)
        lead, core, trail = m.group(1), m.group(2), m.group(3)
        if not core:
            out.append(t)
            continue
        if r.bold and r.italic:
            core = f"***{core}***"
        elif r.bold:
            core = f"**{core}**"
        elif r.italic:
            core = f"*{core}*"
        out.append(lead + core + trail)
    return "".join(out)


def _table(rows: list[list[str]]) -> str:
    width = max(len(r) for r in rows)
    norm = [[_escape(c).replace("\n", " ") for c in r] + [""] * (width - len(r)) for r in rows]
    lines = ["| " + " | ".join(norm[0]) + " |", "| " + " | ".join(["---"] * width) + " |"]
    lines += ["| " + " | ".join(r) + " |" for r in norm[1:]]
    return "\n".join(lines)


def block_md(b: Block) -> list[str]:
    """Markdown paragraphs for one block (a level-1 clause can be two: heading + text)."""
    body = runs_md(b.runs)
    # Headings carry no inline emphasis: their level is their formatting.
    if b.kind == "title":
        return [f"# {_escape(b.text)}"]
    if b.kind == "heading":
        return [f"## {_escape(b.text)}"]
    if b.kind == "table":
        return [_table(b.rows)] if b.rows else []
    if b.kind == "clause":
        number = b.number or ""
        if b.level == 1:
            head = f"## {number} {_escape(b.caption)}" if b.caption else None
            if head and body:
                return [head, _escape_start(body)]
            if head:
                return [head]
            return [_escape_start(f"{number} {body}")]
        cap = ""
        if b.caption:
            cap = f"**{_escape(b.caption)}.**" + ("" if body.startswith("]") else " ")
        return [_escape_start(f"{number} {cap}{body}".rstrip())]
    return [_escape_start(body)] if body else []


def to_markdown(doc: Document) -> str:
    parts: list[str] = []
    for b in doc.blocks:
        parts.extend(block_md(b))
    if doc.notes:
        parts.append("## Footnotes")
        parts.extend(f"{note_marker(n['n'])} {_escape(n['text'])}" for n in doc.notes)
    return "\n\n".join(parts) + "\n"


def block_map(doc: Document) -> list[dict]:
    """Which Markdown paragraphs each block produced, in order.

    After the Markdown is written into a Claude Doc, the doc's own block ids
    come back in the same order, so this map is how a lawdocs block is bound
    to its Claude Docs block (phase 2: authorship and versioning use it).
    """
    out = []
    i = 0
    for b in doc.blocks:
        n = len(block_md(b))
        out.append({"block": b.id, "paragraphs": list(range(i, i + n)), "number": b.number})
        i += n
    return out
