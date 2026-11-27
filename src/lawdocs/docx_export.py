"""Export the block model back to .docx, formatted by a Word template.

This is the LaTeX half of lawdocs: the blocks are the content, the template
is the document class. The template is normally the .docx the document was
imported from, so its styles, multilevel list definitions, headers, footers,
page setup and footnote formatting come back untouched.

How each block is written:

* A block that still has the kind and level it was imported with gets its
  original paragraph properties back (the raw ``pPr`` saved on import),
  including section breaks and the spacer paragraphs that followed it.
* A new or moved block takes the properties most often used in the template
  for its kind and level (its "house style"), so a clause Claude adds at
  level 3 looks like every other level-3 clause.
* Clause numbers are never written as text. They come from Word's own
  numbering: the template's list if its styles carry one, otherwise a legal
  list (1. / 1.1 / (a) / (i) / (A) / (1)) that lawdocs adds to numbering.xml.
* Captions are rebuilt in the template's form (underlined or bold, and on
  the same line through a style separator where the template used one).
* Footnotes return as real Word footnotes.
"""

from __future__ import annotations

import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.sax.saxutils import escape

from .docx_import import signature
from .model import Block, Document, Run, note_number
from .numbering import apply

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

FALLBACK_ABSTRACT_ID = "990"
FALLBACK_NUM_ID = "990"


def _attr(s: str) -> str:
    return escape(s, {'"': "&quot;"})


# --------------------------------------------------------------------------
# Template
# --------------------------------------------------------------------------


class Template:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
        doc = self.parts["word/document.xml"].decode("utf8")
        body_open = re.search(r"<w:body\b[^>]*>", doc)
        body_close = doc.rindex("</w:body>")
        self.head = doc[: body_open.end()]
        self.tail = doc[body_close:]
        body = doc[body_open.end(): body_close]
        # The section properties of the last section are the body's final child.
        last = body.rfind("<w:sectPr")
        self.final_sect = body[last:] if last >= 0 and body.rstrip().endswith("</w:sectPr>") else ""
        styles = self.parts.get("word/styles.xml", b"").decode("utf8")
        self.styles_with_numbering = self._numbered_styles(styles)
        self.has_style = lambda sid: f'w:styleId="{sid}"' in styles

    @staticmethod
    def _numbered_styles(styles: str) -> set[str]:
        """Paragraph styles whose definition (or a base style's) carries numbering."""
        numbered, based = set(), {}
        for m in re.finditer(r"<w:style\b[^>]*w:styleId=\"([^\"]+)\"[^>]*>(.*?)</w:style>", styles, re.S):
            sid, body = m.group(1), m.group(2)
            b = re.search(r'<w:basedOn w:val="([^"]+)"', body)
            if b:
                based[sid] = b.group(1)
            if re.search(r'<w:numId w:val="(?!0")', body):
                numbered.add(sid)
        out = set()
        for sid in list(based) + list(numbered):
            cur, seen = sid, set()
            while cur and cur not in seen:
                seen.add(cur)
                if cur in numbered:
                    out.add(sid)
                    break
                cur = based.get(cur)
        return out | numbered


def _ppr_is_numbered(ppr: str | None, tpl: Template) -> bool:
    if not ppr:
        return False
    m = re.search(r'<w:numId w:val="([^"]+)"', ppr)
    if m:
        return m.group(1) != "0"
    st = re.search(r'<w:pStyle w:val="([^"]+)"', ppr)
    return bool(st and st.group(1) in tpl.styles_with_numbering)


def _with_numbering(ppr: str | None, level: int) -> str:
    """Add lawdocs' fallback legal list to a paragraph's properties."""
    numpr = f'<w:numPr><w:ilvl w:val="{level - 1}"/><w:numId w:val="{FALLBACK_NUM_ID}"/></w:numPr>'
    if not ppr:
        return f"<w:pPr>{numpr}</w:pPr>"
    ppr = re.sub(r"<w:numPr>.*?</w:numPr>", "", ppr, flags=re.S)
    # numPr follows pStyle, keepNext, keepLines, pageBreakBefore, framePr, widowControl.
    m = re.match(r"(<w:pPr\b[^>]*>(?:\s*<w:(?:pStyle|keepNext|keepLines|pageBreakBefore|widowControl)\b[^>]*/>|"
                 r"\s*<w:framePr\b[^>]*/>)*)", ppr)
    if m and m.group(1):
        return ppr[: m.end()] + numpr + ppr[m.end():]
    if ppr.endswith("/>"):
        return ppr[:-2] + ">" + numpr + "</w:pPr>"
    return ppr


def _norm(label: str | None) -> str:
    return (label or "").replace("\u00a0", " ").strip().lstrip("[").rstrip(".:")


def _table_text(rows: list[list[str]] | None) -> list[list[str]]:
    """Cell text without blank rows (Claude Docs drops empty rows, such as the
    blank line a schedule leaves for filling in), spacing differences or empty
    cells at the end of a row (Markdown pads a row with merged cells out to
    the table's width)."""
    out = [[" ".join(c.split()) for c in r] for r in rows or []]
    out = [r[:max((i + 1 for i, c in enumerate(r) if c), default=0)] for r in out]
    return [r for r in out if r]


def _unhide_mark(ppr: str) -> str:
    return re.sub(r"<w:(?:spec)?[vV]anish\s*/>", "", ppr)


def _fallback_numbering_xml() -> str:
    levels = [("decimal", "%1."), ("decimal", "%1.%2"), ("lowerLetter", "(%3)"),
              ("lowerRoman", "(%4)"), ("upperLetter", "(%5)"), ("decimal", "(%6)")]
    lvls = "".join(
        f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{txt}"/>'
        f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{720 * (i + 1)}" w:hanging="720"/></w:pPr></w:lvl>'
        for i, (fmt, txt) in enumerate(levels)
    )
    return (f'<w:abstractNum w:abstractNumId="{FALLBACK_ABSTRACT_ID}"><w:multiLevelType w:val="multilevel"/>'
            f"{lvls}</w:abstractNum>")


# --------------------------------------------------------------------------
# Runs
# --------------------------------------------------------------------------


class _Writer:
    def __init__(self, doc: Document, tpl: Template):
        self.doc = doc
        self.tpl = tpl
        self.needs_fallback = False
        self.note_ids: dict[int, int] = {}
        self.source_note_ids: dict[str, int] = {}  # footnote id in the source -> id in the output
        self.fn_style = tpl.has_style("FootnoteReference")
        self.next_bookmark = 900000
        self.verbatim = self.rebuilt = 0
        self.profile = self._learn_profile()
        self.index = {id(b): i for i, b in enumerate(doc.blocks)}
        self.typed: dict | None = None

    # The template's house style for each kind and level, learned from the
    # blocks that were imported from it.
    def _learn_profile(self) -> dict:
        ppr: dict[tuple, Counter] = {}
        cap: dict[int, Counter] = {}
        for b in self.doc.blocks:
            w = b.word or {}
            if "kind" not in w or w.get("label_only"):
                continue
            key = (w["kind"], w["level"])
            ppr.setdefault(key, Counter())[w.get("pPr") or ""] += 1
            if w["kind"] == "clause" and "caption" in w:
                c = w["caption"]
                cap.setdefault(w["level"], Counter())[(c["form"], c["punct"], w.get("sep_pPr"))] += 1
        return {
            "ppr": {k: c.most_common(1)[0][0] for k, c in ppr.items()},
            "caption": {k: c.most_common(1)[0][0] for k, c in cap.items()},
        }

    def _run(self, r: Run, underline: bool = False) -> str:
        if r.note:
            n = note_number(r.text)
            if n in self.note_ids:
                style = '<w:rStyle w:val="FootnoteReference"/>' if self.fn_style else '<w:vertAlign w:val="superscript"/>'
                return f'<w:r><w:rPr>{style}</w:rPr><w:footnoteReference w:id="{self.note_ids[n]}"/></w:r>'
        props = ("<w:b/>" if r.bold else "") + ("<w:i/>" if r.italic else "") + ('<w:u w:val="single"/>' if underline else "")
        rpr = f"<w:rPr>{props}</w:rPr>" if props else ""
        return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(r.text)}</w:t></w:r>'

    def _runs(self, runs: list[Run]) -> str:
        return "".join(self._run(r) for r in runs)

    # ------------------------------------------------------------------

    def _sibling(self, b: Block) -> dict | None:
        """Word data of the nearest imported clause beside ``b`` at its level and
        under the same parent (the one before it, else the one after it). A
        clause added or moved there is numbered the way that sibling is: in its
        Word list, so Word counts them together, or typed like it."""
        if b.kind != "clause":
            return None
        blocks = self.doc.blocks
        i = self.index.get(id(b))
        if i is None:
            return None
        for step in (-1, 1):
            j = i + step
            while 0 <= j < len(blocks):
                o = blocks[j]
                if (o.kind == "clause" and o.level < b.level) or o.kind in ("heading", "title"):
                    break
                if step == 1 and o.kind == "clause" and o.level == b.level and o.restart:
                    break
                ow = o.word or {}
                if (o is not b and o.kind == "clause" and o.level == b.level and ow.get("kind") == "clause"
                        and ow.get("level") == b.level and not ow.get("label_only")):
                    return ow
                j += step
        return None

    def _ppr_for(self, b: Block) -> str | None:
        """Paragraph properties for ``b``; sets ``self.typed`` when its number
        is to be written as text, as the source typed it."""
        w = b.word or {}
        self.typed = None
        same = w.get("kind") == b.kind and w.get("level") == b.level
        if same and "pPr" in w:
            ppr = w["pPr"]
        elif same:
            ppr = None
        else:
            sib = self._sibling(b)
            ppr = (sib.get("pPr") if sib else None) or self.profile["ppr"].get((b.kind, b.level)) or None
        if b.kind != "clause" or _ppr_is_numbered(ppr, self.tpl):
            return ppr
        if same and w.get("typed"):
            self.typed = w["typed"]  # typed in the source: stays typed
            return ppr
        sib = self._sibling(b)
        if sib and _ppr_is_numbered(sib.get("pPr"), self.tpl):
            return sib["pPr"]  # join the siblings' Word list
        typed = (sib or {}).get("typed")
        if typed:
            self.typed = typed
            return ppr
        self.needs_fallback = True
        return _with_numbering(ppr, b.level)

    def _caption_form(self, b: Block) -> tuple[str, str, str | None]:
        w = b.word or {}
        if w.get("kind") == b.kind and w.get("level") == b.level and "caption" in w:
            return w["caption"]["form"], w["caption"]["punct"], w.get("sep_pPr")
        return self.profile["caption"].get(b.level, ("u", ".", None))

    def _bookmarks(self, b: Block) -> str:
        out = []
        for name in (b.word or {}).get("bookmarks", []):
            self.next_bookmark += 1
            i = self.next_bookmark
            out.append(f'<w:bookmarkStart w:id="{i}" w:name="{_attr(name)}"/><w:bookmarkEnd w:id="{i}"/>')
        return "".join(out)

    def _verbatim(self, b: Block, ppr: str | None) -> str | None:
        """The source paragraph(s) as they were, when the block is unchanged:
        same kind, level, properties and text."""
        w = b.word or {}
        if "content" not in w or w.get("sig") != signature(b.kind, b.level, b.caption, b.runs):
            return None
        if (ppr or None) != (w.get("pPr") or None):
            return None  # numbering had to be added: rebuild
        if w.get("typed") and _norm(w["typed"]["label"]) != _norm(b.number):
            return None  # the typed number is out of date

        ok = True

        def note_id(m: re.Match) -> str:
            nonlocal ok
            new = self.source_note_ids.get(m.group(1))
            if new is None:
                ok = False
                return m.group(0)
            return f'<w:footnoteReference w:id="{new}"'

        content = re.sub(r'<w:footnoteReference w:id="(-?\d+)"', note_id, w["content"])
        sep = re.sub(r'<w:footnoteReference w:id="(-?\d+)"', note_id, w.get("sep_content", ""))
        if not ok:
            return None
        ppr = ppr or ""
        if "sep_content" in w:
            return f"<w:p>{ppr}{content}</w:p><w:p>{w.get('sep_pPr') or ''}{sep}</w:p>"
        return f"<w:p>{ppr}{content}</w:p>"

    def paragraph(self, b: Block) -> str:
        w = b.word or {}
        base_ppr = self._ppr_for(b)
        if not w.get("label_only"):
            raw = self._verbatim(b, base_ppr)
            if raw is not None:
                self.verbatim += 1
                return raw
        ppr = (base_ppr or "") + self._bookmarks(b)
        if w.get("label_only") and w.get("kind") == b.kind:
            self.verbatim += 1  # Word draws the whole paragraph from its numbering
            return f"<w:p>{ppr}</w:p>"
        self.rebuilt += 1
        if w.get("bullet") and w.get("kind") == b.kind and w.get("level") == b.level and b.runs:
            b = Block(b.id, b.kind, b.level, [Run(b.runs[0].text.lstrip("\u2022 "), b.runs[0].bold, b.runs[0].italic)] + b.runs[1:])
        number = ""
        if self.typed and b.kind == "clause":
            number = self._run(Run((b.number or "") + self.typed["sep"]))
        if b.kind != "clause" or not b.caption:
            # Without a caption there is no second paragraph to run into, so the
            # paragraph mark must not stay hidden.
            return f"<w:p>{_unhide_mark(ppr)}{number}{self._runs(b.runs)}</w:p>"
        form, punct, sep_ppr = self._caption_form(b)
        cap = b.caption
        bracket = cap.startswith("[")
        core = Run(cap.lstrip("["), bold=(form == "b"))
        head = number + (self._run(Run("[")) if bracket else "") + self._run(core, underline=(form == "u"))
        head += self._run(Run(punct)) if punct else ""
        body = self._runs(b.runs)
        space = self._run(Run(" ")) if b.runs and not b.runs[0].text.startswith(" ") and not b.runs[0].note else ""
        if sep_ppr is not None:
            # Style separator: the caption paragraph's mark is hidden, so Word
            # shows the caption and the text on one line.
            return f"<w:p>{ppr}{head}</w:p><w:p>{sep_ppr}{space}{body}</w:p>"
        return f"<w:p>{ppr}{head}{space}{body}</w:p>"

    def table(self, b: Block) -> str:
        w = b.word or {}
        raw = w.get("tbl")
        if raw and _table_text(w.get("rows")) == _table_text(b.rows):
            return re.sub(r'<w:footnoteReference w:id="(-?\d+)"',
                          lambda m: f'<w:footnoteReference w:id="{self.source_note_ids.get(m.group(1), m.group(1))}"', raw)
        tblpr = re.search(r"<w:tblPr>.*?</w:tblPr>", raw or "", re.S)
        tblpr = tblpr.group(0) if tblpr else (
            '<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
            + "".join(f'<w:{e} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                      for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
            + "</w:tblBorders></w:tblPr>")
        rows = "".join(
            "<w:tr>" + "".join(f'<w:tc><w:p><w:r><w:t xml:space="preserve">{escape(c)}</w:t></w:r></w:p></w:tc>' for c in r) + "</w:tr>"
            for r in (b.rows or []))
        return f"<w:tbl>{tblpr}{rows}</w:tbl>"

    def block(self, b: Block) -> str:
        out = self.table(b) if b.kind == "table" else self.paragraph(b)
        w = b.word or {}
        if w.get("kind") == b.kind:
            out += "".join(w.get("after", []))
        return out

    # ------------------------------------------------------------------

    def footnotes_xml(self) -> str | None:
        raw = self.tpl.parts.get("word/footnotes.xml")
        if raw is None:
            return None
        xml = raw.decode("utf8")
        first = xml.find("<w:footnote ")
        head = xml[:first] if first >= 0 else xml[: xml.rindex("</w:footnotes>")]
        specials = re.findall(r'<w:footnote\b[^>]*w:type="(?:separator|continuationSeparator|continuationNotice)"[^>]*>.*?</w:footnote>', xml, re.S)
        used = [int(i) for s in specials for i in re.findall(r'w:id="(-?\d+)"', s[:200])]
        next_id = max(used + [0]) + 1
        body = []
        for note in self.doc.notes:
            fid = next_id
            next_id += 1
            self.note_ids[note["n"]] = fid
            raw_note = note.get("word")
            if raw_note:
                src = re.search(r'w:id="(-?\d+)"', raw_note)
                if src:
                    self.source_note_ids[src.group(1)] = fid
                body.append(re.sub(r'w:id="-?\d+"', f'w:id="{fid}"', raw_note, count=1))
            else:
                ref = '<w:rStyle w:val="FootnoteReference"/>' if self.fn_style else '<w:vertAlign w:val="superscript"/>'
                ppr = '<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>' if self.tpl.has_style("FootnoteText") else ""
                body.append(f'<w:footnote w:id="{fid}"><w:p>{ppr}<w:r><w:rPr>{ref}</w:rPr><w:footnoteRef/></w:r>'
                            f'<w:r><w:t xml:space="preserve"> {escape(note["text"])}</w:t></w:r></w:p></w:footnote>')
        return head + "".join(specials) + "".join(body) + "</w:footnotes>"


def export_docx(doc: Document, template: str | Path, out: str | Path, stats: dict | None = None) -> Path:
    apply(doc)
    tpl = Template(template)
    w = _Writer(doc, tpl)
    footnotes = w.footnotes_xml() if doc.notes else None  # assigns footnote ids first
    front = "".join((doc.template or {}).get("front", []))
    body = front + "".join(w.block(b) for b in doc.blocks) + tpl.final_sect
    parts = dict(tpl.parts)
    parts["word/document.xml"] = (tpl.head + body + tpl.tail).encode("utf8")
    if footnotes is not None:
        parts["word/footnotes.xml"] = footnotes.encode("utf8")
    if w.needs_fallback:
        parts["word/numbering.xml"] = _add_fallback_numbering(parts.get("word/numbering.xml"))
    if stats is not None:
        stats.update(verbatim=w.verbatim, rebuilt=w.rebuilt)
    out = Path(out)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return out


def _add_fallback_numbering(raw: bytes | None) -> bytes:
    if raw is None:
        raise ValueError("The template has no numbering part; use a template that has one.")
    xml = raw.decode("utf8")
    if f'w:abstractNumId="{FALLBACK_ABSTRACT_ID}"' in xml:
        return raw
    # abstractNum elements must all come before num elements.
    first_num = xml.find("<w:num ")
    insert_at = first_num if first_num >= 0 else xml.rindex("</w:numbering>")
    xml = xml[:insert_at] + _fallback_numbering_xml() + xml[insert_at:]
    end = xml.rindex("</w:numbering>")
    xml = xml[:end] + f'<w:num w:numId="{FALLBACK_NUM_ID}"><w:abstractNumId w:val="{FALLBACK_ABSTRACT_ID}"/></w:num>' + xml[end:]
    return xml.encode("utf8")
