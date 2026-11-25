"""Read a Claude Doc (as the Markdown its export produces) back into the block model.

This is the return leg of the round trip: .docx -> lawdocs -> Claude Docs ->
(edits by people and Claude) -> lawdocs -> .docx. Numbers in the text are
read only to recover each clause's level; they are then regenerated, so a
clause someone typed as "(c)" in the wrong place simply takes its correct
number.

``pull(markdown, base)`` lines the doc's paragraphs up with the blocks of the
model the doc was made from, so every block that survived keeps its Word
formatting data, id and author; new blocks get the template's house style on
export.
"""

from __future__ import annotations

import difflib
import html
import re

from . import numbering as num
from .model import SUPERSCRIPT_DIGITS, Block, Document, Run, note_number

_UNESCAPE = re.compile(r"\\([\\`*_{}\[\]()#+\-.!|$~<>\"'^])")


def _unescape(s: str) -> str:
    return _UNESCAPE.sub(r"\1", s)


def parse_inline(md: str) -> list[Run]:
    """**bold**, *italic*, ***both*** and backslash escapes -> runs."""
    runs: list[Run] = []
    bold = italic = False
    buf: list[str] = []

    def flush() -> None:
        if buf:
            text = "".join(buf)
            if runs and runs[-1].bold == bold and runs[-1].italic == italic:
                runs[-1].text += text
            else:
                runs.append(Run(text, bold, italic))
            buf.clear()

    i = 0
    while i < len(md):
        c = md[i]
        if c == "\\" and i + 1 < len(md):
            buf.append(md[i + 1])
            i += 2
            continue
        if c in "*_" and (c == "*" or (i == 0 or not md[i - 1].isalnum())):
            n = len(md[i:]) - len(md[i:].lstrip(c))
            n = min(n, 3)
            flush()
            if n == 3:
                bold, italic = not bold, not italic
            elif n == 2:
                bold = not bold
            else:
                italic = not italic
            i += n
            continue
        buf.append(c)
        i += 1
    flush()
    return [r for r in runs if r.text]


def _plain(md: str) -> str:
    return "".join(r.text for r in parse_inline(md))


def _split_md_paragraphs(md: str) -> list[str]:
    # The Claude Docs exporter writes some characters as entities ("&#91;" for a
    # leading "[").
    return [html.unescape(p.strip("\n")) for p in re.split(r"\n\s*\n", md) if p.strip()]


def _note_runs(runs: list[Run], note_ns: set[int]) -> list[Run]:
    """Turn superscript numbers ("¹²") in the text back into footnote markers."""
    out: list[Run] = []
    for r in runs:
        pos = 0
        for m in re.finditer(f"[{SUPERSCRIPT_DIGITS}]+", r.text):
            if note_number(m.group(0)) not in note_ns:
                continue
            if m.start() > pos:
                out.append(Run(r.text[pos:m.start()], r.bold, r.italic))
            out.append(Run(m.group(0), note=True))
            pos = m.end()
        if pos < len(r.text):
            out.append(Run(r.text[pos:], r.bold, r.italic))
    return out


def parse(md: str) -> Document:
    paras = _split_md_paragraphs(md)
    notes: list[dict] = []
    if "## Footnotes" in paras:
        cut = paras.index("## Footnotes")
        for p in paras[cut + 1:]:
            m = re.match(f"^([{SUPERSCRIPT_DIGITS}]+)\\s*(.*)$", p, re.S)
            if m:
                notes.append({"n": note_number(m.group(1)), "text": _plain(m.group(2))})
        paras = paras[:cut]
    note_ns = {n["n"] for n in notes}

    blocks: list[Block] = []
    level = 0
    prev_lower: str | None = None

    def add(kind: str, lvl: int, runs: list[Run], **kw) -> Block:
        b = Block(id=f"m{len(blocks) + 1:04d}", kind=kind, level=lvl, runs=_note_runs(runs, note_ns), **kw)
        blocks.append(b)
        return b

    i = 0
    while i < len(paras):
        p = paras[i]
        if p.startswith("|"):
            lines = [ln for ln in p.splitlines() if ln.strip().startswith("|")]
            rows = [[_plain(c.strip()) for c in re.split(r"(?<!\\)\|", ln.strip())[1:-1]] for ln in lines
                    if not re.fullmatch(r"\|[\s:|-]+\|", ln.strip())]
            add("table", level, [], rows=rows)
        elif p.startswith("# "):
            add("title", 0, parse_inline(p[2:]))
        elif p.startswith("## "):
            text = _plain(p[3:])
            m = re.match(r"^(\d+)\.\s+(.*)$", text)
            if m:
                cap = m.group(2).strip()
                body: list[Run] = []
                # A level-1 clause's text sits in the paragraph under its heading.
                nxt = paras[i + 1] if i + 1 < len(paras) else ""
                if nxt and not nxt.startswith(("#", "|")) and num.split_typed_number(_plain(nxt)) is None:
                    body = parse_inline(nxt)
                    i += 1
                add("clause", 1, body, caption=cap.rstrip("."), source_number=m.group(1) + ".")
                level, prev_lower = 1, None
            else:
                add("heading", 0, parse_inline(p[3:]))
                level = 0
        else:
            runs = parse_inline(p)
            text = "".join(r.text for r in runs)
            typed = num.split_typed_number(text)
            lvl = num.classify_label(typed[0], prev_lower) if typed else None
            if typed and lvl is not None and level > 0:
                label, rest = typed
                runs = _strip(runs, len(text) - len(rest))
                caption = None
                if runs and runs[0].bold:
                    m = re.match(r"^(.{1,80}?)[.:]$", runs[0].text.strip())
                    if m:
                        caption = m.group(1)
                        runs = _strip(runs, len(runs[0].text))
                add("clause", lvl, runs, caption=caption, source_number=label)
                level = lvl
                m = re.fullmatch(r"\(([a-z]+)\)", label)
                if lvl == 3 and m:
                    prev_lower = m.group(1)
                if lvl <= 2:
                    prev_lower = None
            else:
                add("text", level, runs)
        i += 1
    doc = Document(blocks=blocks, notes=notes)
    num.apply(doc)
    return doc


def _strip(runs: list[Run], n: int) -> list[Run]:
    out: list[Run] = []
    for r in runs:
        if n <= 0:
            out.append(r)
        elif len(r.text) <= n:
            n -= len(r.text)
        else:
            out.append(Run(r.text[n:], r.bold, r.italic, note=r.note))
            n = 0
    if out:
        out[0].text = out[0].text.lstrip()
    return [r for r in out if r.text]


def _key(b: Block) -> str:
    text = re.sub(r"\s+", " ", (b.caption or "") + "|" + "".join(r.text for r in b.runs if not r.note)).strip()
    return f"{b.kind}:{text}"


def pull(md: str, base: Document, editor: str = "docs") -> Document:
    """Parse a doc's Markdown and carry over ids, authors and Word data from ``base``."""
    new = parse(md)
    a = [_key(b) for b in base.blocks]
    b_keys = [_key(b) for b in new.blocks]
    sm = difflib.SequenceMatcher(a=a, b=b_keys, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for old, nb in zip(base.blocks[i1:i2], new.blocks[j1:j2]):
                _carry(old, nb, edited=None)
        elif tag == "replace":
            # Same position, changed text: keep identity and formatting, mark as edited.
            for old, nb in zip(base.blocks[i1:i2], new.blocks[j1:j2]):
                if old.kind == nb.kind:
                    _carry(old, nb, edited=editor)
            for nb in new.blocks[j1 + (i2 - i1):j2]:
                nb.author = editor
        elif tag == "insert":
            for nb in new.blocks[j1:j2]:
                nb.author = editor
    # Footnotes keep their Word formatting when their text is unchanged.
    old_notes = {n["text"]: n for n in base.notes}
    for n in new.notes:
        if n["text"] in old_notes and old_notes[n["text"]].get("word"):
            n["word"] = old_notes[n["text"]]["word"]
    new.template = base.template
    new.source = base.source
    num.apply(new)
    return new


def _carry(old: Block, nb: Block, edited: str | None) -> None:
    nb.id = old.id
    nb.word = old.word
    nb.source_number = old.source_number
    nb.author = edited or old.author
    if old.kind == "table" and old.rows and nb.rows and len(old.rows) == len(nb.rows):
        nb.rows = nb.rows  # cell text from the doc wins
