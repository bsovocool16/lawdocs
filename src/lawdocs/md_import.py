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


_SLOT_RX = {"decimal": r"[0-9]+", "decimalZero": r"[0-9]{2,}", "lowerLetter": r"[a-z]{1,3}",
            "upperLetter": r"[A-Z]{1,3}", "lowerRoman": r"[ivxl]{1,8}", "upperRoman": r"[IVXL]{1,8}",
            "ordinal": r"[0-9]+(?:st|nd|rd|th)"}
for _fmt in ("ordinalText", "cardinalText"):
    # Only the number words themselves ("Fifth", "Twenty-One"), longest first.
    _SLOT_RX[_fmt] = "(?:" + "|".join(sorted((re.escape(num.format_counter(k, _fmt)) for k in range(1, 100)),
                                              key=len, reverse=True)) + ")"
_DISTINCT = {"decimal", "ordinal", "ordinalText", "cardinalText"}


def _style_patterns(styles: list[str], allow_bare: bool = False) -> list[re.Pattern]:
    """Regexes that match a number written in one of the document's own styles
    ("Article 4", "1.1.1", "a."), so a clause in that form is read as a clause.
    A style that is nothing but a letter or numeral ("A", "iv") is left out:
    too many sentences start that way."""
    pats = []
    for st in sorted(set(styles), key=len, reverse=True):
        parts = [p for p in re.split(r"(\{(?:n|[1-9]):[A-Za-z]+\})", st) if p]
        literal = "".join(p for p in parts if not p.startswith("{") or p.startswith("{{"))
        if (not allow_bare and not literal.strip()
                and not all(re.fullmatch(r"\{(?:n|[1-9]):(%s)\}" % "|".join(_DISTINCT), p) for p in parts)):
            continue
        rx = ""
        for p in parts:
            m = re.fullmatch(r"\{(?:n|[1-9]):([A-Za-z]+)\}", p)
            rx += _SLOT_RX.get(m.group(1), r"[0-9A-Za-z]+") if m else re.escape(p.replace("{{", "{").replace("}}", "}"))
        pats.append(re.compile(r"^(" + rx + r")(?=\s)\s*"))
    return pats


def _split_label(text: str, pats: list[re.Pattern]) -> tuple[str, str] | None:
    for pat in pats:
        m = pat.match(text)
        if m and text[m.end():].strip():
            return m.group(1), text[m.end():]
    return num.split_typed_number(text)


def _norm(text: str) -> str:
    return " ".join(text.split())


def parse(md: str, styles: dict[int, list[str]] | None = None, known: dict[str, str] | None = None) -> Document:
    """Blocks from a doc's Markdown. ``styles`` are the number forms the
    document uses at each level (from the model it was made from), so
    "Article 4" or "a." read as numbers in that document. ``known`` maps the
    text of that model's unnumbered paragraphs and headings to their kind: a
    paragraph still word for word the same stays what it was, even if it
    starts with something that looks like a number."""
    styles = styles or {}
    known = known or {}
    pats = _style_patterns([st for sts in styles.values() for st in sts])
    # A "## " line is a heading or a level-1 clause, so even a bare "A" can be a number there.
    top_pats = _style_patterns(styles.get(1, []), allow_bare=True)
    reader = num.LevelReader()
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
        elif p.startswith("## ") and known.get(_norm(_plain(p[3:]))) == "heading":
            add("heading", 0, parse_inline(p[3:]))
            reader.reset()
            level = 0
        elif p.startswith("## "):
            text = _plain(p[3:])
            m = re.match(r"^(\d+\.)\s+(.*)$", text)
            sl = _split_label(text, top_pats) if not m else None
            if sl and (num.classify_label(sl[0]) or 0) > 1 and not num._relative(sl[0]):
                sl = None  # a level-1 heading never carries "1.1"
            if m or sl:
                label, cap = (m.group(1), m.group(2).strip()) if m else (sl[0], sl[1].strip())
                body: list[Run] = []
                # A level-1 clause's text sits in the paragraph under its heading.
                nxt = paras[i + 1] if i + 1 < len(paras) else ""
                if nxt and not nxt.startswith(("#", "|")) and _split_label(_plain(nxt), pats) is None:
                    body = parse_inline(nxt)
                    i += 1
                reader.read(label, 1)
                reader.stack = {1: reader.stack.get(1, "")}
                add("clause", 1, body, caption=cap.rstrip(".") or None, source_number=label)
                level = 1
            else:
                add("heading", 0, parse_inline(p[3:]))
                reader.reset()
                level = 0
        else:
            runs = parse_inline(p)
            text = "".join(r.text for r in runs)
            typed = _split_label(text, pats) if known.get(_norm(text)) not in ("text", "heading") else None
            # Outside any clause only a number that says where it sits ("1.",
            # "2.3.2", "Article 2") starts one, or one in a form this document
            # uses for clauses ("(A)" recitals); otherwise "(i)" is just text.
            top = typed and (not num._relative(typed[0]) or any(pt.match(text) for pt in pats))
            lvl = reader.read(typed[0], fallback=level + 1) if typed and (level > 0 or top) else None
            if lvl is not None:
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
    num.apply(base)
    styles: dict[int, list[str]] = {}
    for b in base.clauses():
        if b.style:
            styles.setdefault(b.level, []).append(b.style)
    known = {_norm(b.text): b.kind for b in base.blocks if b.kind in ("text", "heading", "title") and b.text.strip()}
    new = parse(md, styles=styles, known=known)
    _split_lost_bodies(new, base)
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
    # Aligned in order, because the same note text can appear more than once.
    nm = difflib.SequenceMatcher(a=[n["text"] for n in base.notes], b=[n["text"] for n in new.notes], autojunk=False)
    for blk in nm.get_matching_blocks():
        for k in range(blk.size):
            old_note = base.notes[blk.a + k]
            if old_note.get("word"):
                new.notes[blk.b + k]["word"] = old_note["word"]
    new.template = base.template
    new.source = base.source
    num.apply(new)
    return new


def _split_lost_bodies(new: Document, base: Document) -> None:
    """A level-1 clause with only a caption, followed by a plain paragraph,
    renders exactly like one whose text is that paragraph. Where the model the
    doc came from had the caption-only clause, split it back."""
    bare = set()
    for b, nxt in zip(base.blocks, base.blocks[1:]):
        if b.kind == "clause" and b.level == 1 and not b.text.strip() and nxt.kind == "text":
            bare.add(((b.caption or "").strip(), _norm(nxt.text)))
    if not bare:
        return
    out: list[Block] = []
    for b in new.blocks:
        if (b.kind == "clause" and b.level == 1 and b.runs
                and ((b.caption or "").strip(), _norm(b.text)) in bare):
            body = Block(id=b.id + "t", kind="text", level=1, runs=b.runs)
            b.runs = []
            out += [b, body]
        else:
            out.append(b)
    new.blocks = out


def _carry(old: Block, nb: Block, edited: str | None) -> None:
    if old.kind == nb.kind == "clause":
        if (nb.source_number or "").strip() == (old.number or "").strip():
            # The number is the one lawdocs showed, so the clause wasn't moved:
            # keep its known level rather than the one read back from its label.
            nb.level = old.level
        if nb.level == old.level:
            nb.style = old.style
            nb.restart = old.restart
            nb.alternative = old.alternative
            nb.continues = old.continues
    nb.id = old.id
    nb.word = old.word
    nb.source_number = old.source_number
    nb.author = edited or old.author
    if old.kind == nb.kind and nb.kind in ("text", "table"):
        # Markdown can't say which clause a plain paragraph belongs to, so the
        # level guessed from its neighbours must not override the known one.
        nb.level = old.level
    if old.kind == "table" and old.rows and nb.rows and len(old.rows) == len(nb.rows):
        nb.rows = nb.rows  # cell text from the doc wins
