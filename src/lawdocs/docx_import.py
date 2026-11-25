"""Import a .docx into the lawdocs block model.

Word stores what a reader sees as a pile of layered formatting: direct run
formatting, paragraph properties, style chains, list definitions in
numbering.xml, overrides of those definitions, and numbers people typed by
hand. This importer reads all of it once and keeps only structure:

* which paragraphs are numbered clauses, and at what *legal* level
  (1 / 1.1 / (a) / (i) / (A) / (1)), whether Word auto-numbered them or a
  person typed the number;
* which paragraphs are body text continuing a clause, headings, or the title;
* bold and italic inside the text (defined terms are usually bold);
* tables, as rows of cell text.

Everything else (indents, fonts, spacing, colours, tab stops, keep-with-next)
is dropped on purpose. The renderer derives presentation from structure.

The legal level of a numbered paragraph is read from the number Word would
display ("1.1" is level 2, "(a)" is level 3, ...) rather than from Word's list
indent level, because real documents routinely mix lists whose ilvl values
mean different things. ilvl is only the fallback.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from . import numbering as num
from .model import Block, Document, Run, note_marker

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _val(el: ET.Element | None, attr: str = "val") -> str | None:
    if el is None:
        return None
    return el.get(q(attr))


def _on(el: ET.Element | None) -> bool:
    """A w:b / w:i style toggle: present and not explicitly off."""
    if el is None:
        return False
    v = el.get(q("val"))
    return v not in ("0", "false", "off", "none")


# --------------------------------------------------------------------------
# styles.xml
# --------------------------------------------------------------------------


@dataclass
class Style:
    id: str
    name: str = ""
    based_on: str | None = None
    num_id: str | None = None
    ilvl: int | None = None
    outline_lvl: int | None = None


def _read_styles(root: ET.Element | None) -> dict[str, Style]:
    styles: dict[str, Style] = {}
    if root is None:
        return styles
    for s in root.findall("w:style", NS):
        if s.get(q("type")) != "paragraph":
            continue
        sid = s.get(q("styleId")) or ""
        st = Style(id=sid, name=(_val(s.find("w:name", NS)) or sid))
        st.based_on = _val(s.find("w:basedOn", NS))
        ppr = s.find("w:pPr", NS)
        if ppr is not None:
            npr = ppr.find("w:numPr", NS)
            if npr is not None:
                st.num_id = _val(npr.find("w:numId", NS))
                il = _val(npr.find("w:ilvl", NS))
                st.ilvl = int(il) if il is not None else None
            ol = _val(ppr.find("w:outlineLvl", NS))
            if ol is not None:
                st.outline_lvl = int(ol)
        styles[sid] = st
    return styles


def _style_chain(styles: dict[str, Style], sid: str | None):
    seen = set()
    while sid and sid in styles and sid not in seen:
        seen.add(sid)
        yield styles[sid]
        sid = styles[sid].based_on


def _heading_level(styles: dict[str, Style], sid: str | None) -> int | None:
    for st in _style_chain(styles, sid):
        if st.outline_lvl is not None and st.outline_lvl < 9:
            return st.outline_lvl + 1
        m = re.fullmatch(r"heading\s*([1-9])", st.name.strip(), re.I)
        if m:
            return int(m.group(1))
    return None


def _is_toc_style(styles: dict[str, Style], sid: str | None) -> bool:
    """A table-of-contents entry ("TOC 1", "toc 2"): it repeats headings with
    their numbers typed in, and must not be read as clauses."""
    return any(re.fullmatch(r"toc\s*[1-9]|table of contents|toc heading", st.name.strip(), re.I)
               or re.fullmatch(r"TOC[1-9]", st.id) for st in _style_chain(styles, sid))


def _is_title_style(styles: dict[str, Style], sid: str | None) -> bool:
    return any(st.name.strip().lower() == "title" for st in _style_chain(styles, sid))


# --------------------------------------------------------------------------
# numbering.xml
# --------------------------------------------------------------------------


@dataclass
class Lvl:
    start: int = 1
    fmt: str = "decimal"
    text: str = ""
    is_lgl: bool = False
    p_style: str | None = None


@dataclass
class Numbering:
    abstract: dict[str, dict[int, Lvl]] = field(default_factory=dict)
    nums: dict[str, str] = field(default_factory=dict)  # numId -> abstractNumId
    overrides: dict[str, dict[int, int]] = field(default_factory=dict)  # numId -> {ilvl: start}
    # Running counters per abstractNum; None = level not used since last reset.
    _counters: dict[str, list[int | None]] = field(default_factory=dict)
    _seen_nums: set[str] = field(default_factory=set)
    _starts: dict[str, dict[int, int]] = field(default_factory=dict)  # abstractNumId -> active start overrides

    def lvl(self, num_id: str, ilvl: int) -> Lvl | None:
        aid = self.nums.get(num_id)
        if aid is None:
            return None
        return self.abstract.get(aid, {}).get(ilvl)

    def ilvl_for_style(self, num_id: str, style_id: str) -> int | None:
        aid = self.nums.get(num_id)
        for il, lv in self.abstract.get(aid or "", {}).items():
            if lv.p_style == style_id:
                return il
        return None

    def next_label(self, num_id: str, ilvl: int) -> str | None:
        """Advance the counter for (numId, ilvl) and return Word's displayed label."""
        aid = self.nums.get(num_id)
        if aid is None or aid not in self.abstract:
            return None
        levels = self.abstract[aid]
        counters = self._counters.setdefault(aid, [None] * 9)
        starts = self._starts.setdefault(aid, {})
        if num_id not in self._seen_nums:
            self._seen_nums.add(num_id)
            ov = self.overrides.get(num_id)
            if ov:
                # A list instance with start overrides begins afresh: each level
                # counts from its override, and a level shown but not yet used
                # in this instance ("%1" in "%1.%2") shows its start value.
                for i in range(9):
                    counters[i] = None
                starts.clear()
                starts.update(ov)

        def start(k: int) -> int:
            return starts.get(k, levels.get(k, Lvl()).start)

        lv = levels.get(ilvl, Lvl())
        if counters[ilvl] is None:
            counters[ilvl] = start(ilvl) - 1
        counters[ilvl] += 1
        for i in range(ilvl + 1, 9):
            counters[i] = None
            starts.pop(i, None)  # restarted by its parent: back to its own start
        if lv.fmt == "bullet":
            return "•"

        def sub(m: re.Match) -> str:
            k = int(m.group(1)) - 1
            klv = levels.get(k, Lvl())
            if counters[k] is None:
                counters[k] = start(k)  # shown, so it now counts from there
            val = counters[k]
            fmt = "decimal" if (lv.is_lgl and klv.fmt != "none") else klv.fmt
            return num.format_counter(val, fmt)

        return re.sub(r"%([1-9])", sub, lv.text).strip()


def _read_numbering(root: ET.Element | None) -> Numbering:
    n = Numbering()
    if root is None:
        return n
    for a in root.findall("w:abstractNum", NS):
        aid = a.get(q("abstractNumId")) or ""
        levels: dict[int, Lvl] = {}
        for lvl in a.findall("w:lvl", NS):
            il = int(lvl.get(q("ilvl")) or 0)
            levels[il] = Lvl(
                start=int(_val(lvl.find("w:start", NS)) or 1),
                fmt=_val(lvl.find("w:numFmt", NS)) or "decimal",
                text=_val(lvl.find("w:lvlText", NS)) or "",
                is_lgl=lvl.find("w:isLgl", NS) is not None,
                p_style=_val(lvl.find("w:pStyle", NS)),
            )
        n.abstract[aid] = levels
    for m in root.findall("w:num", NS):
        nid = m.get(q("numId")) or ""
        n.nums[nid] = _val(m.find("w:abstractNumId", NS)) or ""
        for ov in m.findall("w:lvlOverride", NS):
            so = ov.find("w:startOverride", NS)
            if so is not None:
                n.overrides.setdefault(nid, {})[int(ov.get(q("ilvl")) or 0)] = int(_val(so) or 1)
    return n


# --------------------------------------------------------------------------
# Paragraph text
# --------------------------------------------------------------------------

_SKIP_SUBTREES = {q("del"), q("moveFrom"), q("instrText"), q("delText"), q("rPr"), q("pPr"),
                  q("endnoteReference"), q("commentReference")}


def _runs(p: ET.Element, footnote=None) -> list[Run]:
    """Text runs of a paragraph. ``footnote(id) -> marker`` is called for each footnote reference."""
    runs: list[Run] = []

    def emit(text: str, bold: bool, italic: bool, underline: bool = False, note: bool = False) -> None:
        if not text:
            return
        r = Run(text, bold, italic, underline, note)
        if runs and runs[-1].same_format(r):
            runs[-1].text += text
        else:
            runs.append(r)

    def walk(el: ET.Element) -> None:
        for child in el:
            if child.tag in _SKIP_SUBTREES:
                continue
            if child.tag == q("r"):
                rpr = child.find("w:rPr", NS)
                bold = _on(rpr.find("w:b", NS)) if rpr is not None else False
                italic = _on(rpr.find("w:i", NS)) if rpr is not None else False
                u = rpr.find("w:u", NS) if rpr is not None else None
                underline = u is not None and _val(u) != "none"
                if rpr is not None and _on(rpr.find("w:vanish", NS)):
                    continue  # hidden text
                for part in child:
                    if part.tag == q("t"):
                        emit(part.text or "", bold, italic, underline)
                    elif part.tag in (q("tab"), q("ptab"), q("br"), q("cr")):
                        emit(" ", bold, italic, underline)
                    elif part.tag == q("noBreakHyphen"):
                        emit("-", bold, italic, underline)
                    elif part.tag == q("footnoteReference") and footnote is not None:
                        marker = footnote(part.get(q("id")))
                        if marker:
                            emit(marker, False, False, note=True)
            else:
                walk(child)

    walk(p)
    return _normalise(runs)


def _normalise(runs: list[Run]) -> list[Run]:
    out: list[Run] = []
    for r in runs:
        r.text = re.sub(r"[ \t\u00a0]+", " ", r.text)
        if out and out[-1].same_format(r):
            out[-1].text += r.text
        else:
            out.append(r)
    # Collapse spaces across run boundaries and trim the ends.
    for i in range(1, len(out)):
        if out[i - 1].text.endswith(" ") and out[i].text.startswith(" "):
            out[i].text = out[i].text[1:]
    if out:
        out[0].text = out[0].text.lstrip()
        out[-1].text = out[-1].text.rstrip()
    return [r for r in out if r.text]


def _strip_prefix(runs: list[Run], n_chars: int) -> list[Run]:
    out: list[Run] = []
    for r in runs:
        if n_chars <= 0:
            out.append(r)
        elif len(r.text) <= n_chars:
            n_chars -= len(r.text)
        else:
            out.append(Run(r.text[n_chars:], r.bold, r.italic, r.underline))
            n_chars = 0
    return _normalise(out)


_ABBREVIATIONS = {"Corp", "Inc", "Ltd", "Co", "No", "Nos", "Mr", "Ms", "Mrs", "Dr", "St", "Art", "Sec", "U.S", "L.P", "N.A"}


def _split_caption(runs: list[Run], level: int) -> tuple[str | None, list[Run]]:
    """Pull a clause caption ("Payment." / "DEFINITIONS") off the front of a clause."""
    if not runs:
        return None, runs
    # An optional caption in brackets: "[" + underlined "Waiver of Conflicts" + ". ..."
    if len(runs) > 1 and runs[0].text.strip() == "[" and (runs[1].bold or runs[1].underline):
        cap, rest = _split_caption(runs[1:], level)
        if cap is not None:
            return "[" + cap, rest
        return None, runs
    text = "".join(r.text for r in runs)
    # A leading bold phrase that ends in a period or colon: "Payment. The Buyer shall..."
    first = runs[0]
    if (first.bold or first.underline) and re.match(r"^\[?[A-Z0-9]", first.text):
        m = re.match(r"^([^.:]{1,80})[.:](?![0-9])\s*", first.text)
        if m and len(m.group(1).split()) <= 10 and not m.group(1).startswith(('"', "\u201c")):
            rest = [Run(first.text[m.end():], first.bold, first.italic)] + runs[1:]
            return m.group(1).strip(), _normalise(rest)
        # A footnote marker between the caption and its period: "Caption[80]. Text"
        note = len(runs) > 2 and runs[1].note
        if note and runs[2].text.lstrip()[:1] in (".", ":") and len(first.text.split()) <= 10:
            rest = [Run(runs[2].text.lstrip()[1:].lstrip(), runs[2].bold, runs[2].italic)] + runs[3:]
            body = _normalise(rest)
            return first.text.strip(), ([runs[1], Run(" ")] + body) if body else [runs[1]]
        if len(runs) > 1 and len(first.text.split()) <= 10 and runs[1].text.lstrip()[:1] in (".", ":"):
            rest = [Run(runs[1].text.lstrip()[1:], runs[1].bold, runs[1].italic)] + runs[2:]
            return first.text.strip(), _normalise(rest)
    # An unformatted Title Case lead-in, as typed: "Late Payment. Late amounts bear..."
    m = re.match(r"^((?:[A-Z][\w'-]*|of|and|or|the|to|for|in|on)(?: (?:[A-Z][\w'-]*|of|and|or|the|to|for|in|on)){0,5})\.\s+(?=\S)", text)
    if m and m.group(1)[0].isupper() and len(m.group(1).split()) <= 6 and not first.bold:
        words = m.group(1).split()
        if words[-1] not in _ABBREVIATIONS and all(w[0].isupper() for w in (words[0], words[-1])):
            return m.group(1), _strip_prefix(runs, m.end())
    # A short top-level clause with no sentence in it is a heading-style caption.
    if level == 1 and len(text.split()) <= 10 and not text.rstrip().endswith((".", ";", ",")):
        return text.strip(), []
    return None, runs


def _is_caps_heading(text: str) -> bool:
    letters = [c for c in text if c.isalpha()]
    return (
        bool(letters)
        and all(c.isupper() for c in letters)
        and len(text.split()) <= 10
        and not text.rstrip().endswith((".", ";", ","))
    )


# --------------------------------------------------------------------------
# Import
# --------------------------------------------------------------------------


def _bookmarks(p: ET.Element) -> list[str]:
    """Bookmark names in a paragraph. Word cross-references (REF fields, often in
    footnotes) point at them, so they are re-created on export."""
    return [n for b in p.iter(q("bookmarkStart")) if (n := b.get(q("name"))) and n != "_GoBack"]


def _is_style_separator(p: ET.Element) -> bool:
    """A paragraph whose mark is hidden (Word's "style separator"): it runs on
    into the next paragraph, typically a heading-styled caption followed by body text."""
    mark = p.find("w:pPr/w:rPr", NS)
    return mark is not None and (mark.find("w:specVanish", NS) is not None or _on(mark.find("w:vanish", NS)))


def _read_footnotes(root: ET.Element | None) -> tuple[dict[str, str], dict[str, str]]:
    """Footnote text and raw XML by footnote id."""
    text: dict[str, str] = {}
    raw: dict[str, str] = {}
    if root is None:
        return text, raw
    for f in root.findall("w:footnote", NS):
        if f.get(q("type")) in ("separator", "continuationSeparator", "continuationNotice"):
            continue
        fid = f.get(q("id")) or ""
        parts = ["".join(r.text for r in _runs(p)) for p in f.findall(".//w:p", NS)]
        text[fid] = " ".join(t for t in parts if t)
        raw[fid] = xml_str(f)
    return text, raw


# --------------------------------------------------------------------------
# Raw XML for the round trip
# --------------------------------------------------------------------------

_NS_DECL = re.compile(r'\sxmlns:([A-Za-z0-9]+)="([^"]*)"')
_KNOWN_NS: dict[str, str] = {}
# Prefixes declared on the root element of the part being read. Only these can
# be left off a saved fragment: a prefix declared further down (pictures declare
# a:, pic:, a14: on the drawing itself) must stay on the fragment.
_ROOT_NS: dict[str, str] = {}


def register_namespaces(xml: bytes) -> dict[str, str]:
    """Make ElementTree write the source package's own prefixes (w:, w14:, ...).
    Returns the prefixes the part's root element declares."""
    text = xml.decode("utf8", "replace")
    root = re.search(r"<(?![?!])[^>]*>", text)
    root_ns = dict(_NS_DECL.findall(root.group(0))) if root else {}
    # Root declarations last, so they win if a prefix is rebound further down.
    for prefix, uri in _NS_DECL.findall(text) + list(root_ns.items()):
        if prefix != "xml" and _KNOWN_NS.get(prefix) != uri:
            try:
                ET.register_namespace(prefix, uri)
                _KNOWN_NS[prefix] = uri
            except ValueError:
                pass
    return root_ns


def xml_str(el: ET.Element) -> str:
    """Serialise an element without the namespace declarations the part's root already makes."""
    s = ET.tostring(el, encoding="unicode")
    return _NS_DECL.sub(lambda m: "" if _ROOT_NS.get(m.group(1)) == m.group(2) else m.group(0), s)


_TRACKED = {q("ins"), q("del"), q("moveFrom"), q("moveTo")}


def _raw_content(p: ET.Element) -> str | None:
    """A paragraph's children other than its properties, as XML. None when it
    holds tracked changes, which the model has accepted and the raw XML hasn't."""
    if any(el.tag in _TRACKED for el in p.iter()):
        return None
    return "".join(xml_str(c) for c in p if c.tag != q("pPr"))


def signature(kind: str, level: int, caption: str | None, runs) -> str:
    """What a block says, ignoring run splits and spacing, so an untouched
    paragraph is recognised after a trip through Claude Docs."""
    text = " ".join("".join(r.text for r in runs).split())
    return f"{kind}|{level}|{' '.join((caption or '').split())}|{text}"


class _Importer:
    def __init__(self, styles: dict[str, Style], numbering: Numbering, source: str):
        self.styles = styles
        self.numbering = numbering
        self.source = source
        self.blocks: list[Block] = []
        self.warnings: list[str] = []
        self.current_level = 0
        self.prev_alpha_lower: str | None = None
        self.prev_alpha_upper: str | None = None
        self.levels = num.LevelReader()
        self.typed_numbers = 0
        self.bullets = 0
        self.style_separators = 0
        self.held: ET.Element | None = None
        self.footnote_texts: dict[str, str] = {}
        self.footnote_xml: dict[str, str] = {}
        self.notes: list[dict] = []
        # Word round trip: raw properties of the paragraph being imported, and
        # spacer paragraphs seen before the first block.
        self._word: dict = {}
        self.front: list[str] = []

    def _new_id(self) -> str:
        return f"b{len(self.blocks) + 1:04d}"

    def _add(self, kind: str, level: int, runs: list[Run], **kw) -> Block:
        b = Block(id=self._new_id(), kind=kind, level=level, runs=runs,
                  author=f"import:{self.source}", **kw)
        word = {k: v for k, v in self._word.items() if v is not None}
        word.update(kind=kind, level=level)
        if "content" in word:
            word["sig"] = signature(kind, level, kw.get("caption"), runs)
        b.word = word
        self._word = {}
        self.blocks.append(b)
        return b

    def _spacer(self, p: ET.Element) -> None:
        """An empty paragraph: nothing in Claude Docs, but kept for the Word round trip
        (it may hold a page break, a section break or deliberate spacing)."""
        raw = xml_str(p)
        if self.blocks:
            self.blocks[-1].word.setdefault("after", []).append(raw)
        else:
            self.front.append(raw)

    def _num_props(self, ppr: ET.Element | None, style_id: str | None) -> tuple[str | None, int]:
        num_id: str | None = None
        ilvl: int | None = None
        if ppr is not None:
            npr = ppr.find("w:numPr", NS)
            if npr is not None:
                num_id = _val(npr.find("w:numId", NS))
                il = _val(npr.find("w:ilvl", NS))
                ilvl = int(il) if il is not None else None
        if num_id is None:
            for st in _style_chain(self.styles, style_id):
                if st.num_id is not None:
                    num_id = st.num_id
                    if ilvl is None:
                        ilvl = st.ilvl
                    break
        if num_id is not None and ilvl is None and style_id:
            ilvl = self.numbering.ilvl_for_style(num_id, style_id)
        if num_id in (None, "0"):
            return None, 0
        return num_id, ilvl or 0

    def _track_alpha(self, label: str, level: int) -> None:
        m = re.fullmatch(r"\(([a-zA-Z]+)\)", label.strip())
        if not m:
            return
        tok = m.group(1)
        if level == 3 and tok.islower():
            self.prev_alpha_lower = tok
        elif level == 5 and tok.isupper():
            self.prev_alpha_upper = tok

    def _clause(self, label: str, fallback_level: int, runs: list[Run], heading: bool,
                fmt: str | None = None) -> None:
        level = self.levels.read(label, fallback_level, fmt) or fallback_level
        self.prev_alpha_lower, self.prev_alpha_upper = self.levels.prev_lower, self.levels.prev_upper
        caption, body = _split_caption(runs, level)
        text = "".join(r.text for r in runs).strip()
        if (heading and caption is None and level <= 2 and not any(r.note for r in runs)
                and len(text.split()) <= 12 and not text.endswith((".", ":", ";", ",", "]"))):
            caption, body = text, []
        if caption is not None:
            lead = runs[1] if runs and runs[0].text.strip() == "[" and len(runs) > 1 else (runs[0] if runs else None)
            form = "u" if lead and lead.underline else "b" if lead and lead.bold else "plain"
            core = caption.lstrip("[")
            after = re.sub(r"^\s*\[\d+\]", "", text[text.find(core) + len(core):]).lstrip()
            self._word["caption"] = {"form": form, "punct": after[:1] if after[:1] in (".", ":") else ""}
        b = self._add("clause", level, body, source_number=label, caption=caption)
        b.style = num.label_style(label, level, fmt)
        b.word["source_fmt"] = fmt
        self.current_level = level

    def footnote(self, fid: str | None) -> str | None:
        text = self.footnote_texts.get(fid or "")
        if text is None:
            return None
        n = len(self.notes) + 1
        self.notes.append({"n": n, "text": text, "word": self.footnote_xml.get(fid or "")})
        return note_marker(n)

    def paragraph(self, p: ET.Element, sep: ET.Element | None = None) -> None:
        """One block from a paragraph. ``sep`` is the paragraph that follows a
        style separator, which Word displays on the same line."""
        ppr = p.find("w:pPr", NS)
        style_id = _val(ppr.find("w:pStyle", NS)) if ppr is not None else None
        runs = _runs(p, self.footnote)
        self._word = {"pPr": xml_str(ppr) if ppr is not None else None}
        # The paragraph's own runs, fields, hyperlinks and bookmarks, written back
        # verbatim while its text is unchanged (so underlining, small caps,
        # live cross-references and the like survive the round trip).
        content = _raw_content(p)
        sep_content = _raw_content(sep) if sep is not None else ""
        if content is not None and sep_content is not None:
            self._word["content"] = content
            if sep is not None:
                self._word["sep_content"] = sep_content
        marks = _bookmarks(p) + (_bookmarks(sep) if sep is not None else [])
        if marks:
            self._word["bookmarks"] = marks
        if sep is not None:
            sep_ppr = sep.find("w:pPr", NS)
            self._word["sep_pPr"] = xml_str(sep_ppr) if sep_ppr is not None else ""
            runs = _normalise(runs + _runs(sep, self.footnote))
        text = "".join(r.text for r in runs)
        num_id, ilvl = self._num_props(ppr, style_id)
        label = self.numbering.next_label(num_id, ilvl) if num_id else None
        if not text.strip():
            if label and label != "\u2022" and num.classify_label(label) is None:
                # A numbered paragraph whose whole content is its label ("Exhibit A").
                self._word["label_only"] = True
                self._add("heading", 0, [Run(label.replace("\u00a0", " "))])
                self.current_level = 0
            else:
                self._word = {}
                self._spacer(p)
            return

        heading_lvl = _heading_level(self.styles, style_id)
        if _is_toc_style(self.styles, style_id) and re.search(r"\s[0-9ivxlc]{1,5}\s*$", text):
            # Drafters do use TOC styles for body text, so only a line that ends
            # in a page number counts as a contents entry.
            self._add("text", 0, runs)
            return

        if label == "•":
            self.bullets += 1
            self._word["bullet"] = True  # Word draws the bullet; the text copy is for Claude Docs
            self._add("text", self.current_level, [Run("\u2022 ")] + runs)
            return
        if label:
            lvl_def = self.numbering.lvl(num_id, ilvl)
            self._clause(label, ilvl + 1, runs, heading=heading_lvl is not None,
                         fmt=lvl_def.fmt if lvl_def else None)
            return

        typed = num.split_typed_number(text)
        if typed:
            t_label, rest = typed
            level = num.classify_label(t_label, self.prev_alpha_lower, self.prev_alpha_upper)
            # A bare "(i)" with no clause open is more likely an inline list than a clause.
            if level is not None and (level <= 2 or self.current_level > 0):
                self.typed_numbers += 1
                prefix_len = len(text) - len(rest)
                # Kept so an untouched paragraph goes back as typed, and a new
                # clause in a typed document is typed the same way.
                lead = text[:prefix_len]
                self._word["typed"] = {"label": t_label, "sep": lead[lead.index(t_label) + len(t_label):] or " "}
                self._clause(t_label, level, _strip_prefix(runs, prefix_len), heading=heading_lvl is not None)
                return

        has_clause = any(b.kind == "clause" for b in self.blocks)
        has_title = any(b.kind == "title" for b in self.blocks)
        if not has_title and not has_clause and (_is_title_style(self.styles, style_id) or _is_caps_heading(text)):
            self._add("title", 0, runs)
            return
        if re.match(r"^\[?(IN WITNESS WHEREOF|\[?Signature Page)", text, re.I):
            self.current_level = 0
        if _is_title_style(self.styles, style_id):
            self._add("title" if not any(b.kind == "title" for b in self.blocks) else "heading", 0, runs)
            return
        if heading_lvl is not None or (_is_caps_heading(text) and len(text) > 2):
            self._add("heading", 0, runs)
            self.current_level = 0
            self.levels.reset()
            return
        self._add("text", self.current_level if has_clause else 0, runs)

    def table(self, tbl: ET.Element) -> None:
        rows: list[list[str]] = []
        numbered = False
        for tr in tbl.findall("w:tr", NS):
            row: list[str] = []
            for tc in tr.findall("w:tc", NS):
                parts = []
                for p in tc.iter(q("p")):
                    t = "".join(r.text for r in _runs(p, self.footnote))
                    if t:
                        parts.append(t)
                    if p.find("w:pPr/w:numPr", NS) is not None:
                        numbered = True
                row.append(" ".join(parts))
            rows.append(row)
        if any(any(c for c in r) for r in rows):
            self._word = {"tbl": xml_str(tbl), "rows": rows}
            self._add("table", self.current_level, [], rows=rows)
        else:
            # An empty table (a box to fill in): nothing for Claude Docs, but it
            # goes back to Word where it was, like an empty paragraph.
            self._spacer(tbl)
        if numbered:
            self.warnings.append("A table contains numbered paragraphs; their numbers were not imported.")

    def body(self, body: ET.Element) -> None:
        for child in body:
            if child.tag == q("p"):
                if self.held is not None:
                    held, self.held = self.held, None
                    self.paragraph(held, sep=child)
                elif _is_style_separator(child):
                    self.held = child
                    self.style_separators += 1
                else:
                    self.paragraph(child)
            elif child.tag == q("tbl"):
                self.table(child)
            elif child.tag == q("sdt"):
                content = child.find("w:sdtContent", NS)
                if content is not None:
                    self.body(content)


def _load(z: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(z.read(name))
    except KeyError:
        return None


def import_docx(path: str | Path) -> Document:
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        try:
            raw_doc = z.read("word/document.xml")
        except KeyError:
            raise ValueError(f"{path} has no word/document.xml; is it a .docx?") from None
        doc_ns = register_namespaces(raw_doc)
        document = ET.fromstring(raw_doc)
        styles = _read_styles(_load(z, "word/styles.xml"))
        numbering = _read_numbering(_load(z, "word/numbering.xml"))
        try:
            raw_fn = z.read("word/footnotes.xml")
            fn_ns = register_namespaces(raw_fn)
            footnotes = ET.fromstring(raw_fn)
        except KeyError:
            footnotes, fn_ns = None, {}

    imp = _Importer(styles, numbering, path.name)
    _ROOT_NS.clear()
    _ROOT_NS.update(fn_ns)
    imp.footnote_texts, imp.footnote_xml = _read_footnotes(footnotes)
    _ROOT_NS.clear()
    _ROOT_NS.update(doc_ns)
    body = document.find("w:body", NS)
    if body is not None:
        imp.body(body)
        if imp.held is not None:
            imp.paragraph(imp.held)

    warnings = imp.warnings
    ins = len(document.findall(".//w:ins", NS))
    dels = len(document.findall(".//w:del", NS))
    if ins or dels:
        warnings.append(f"Accepted {ins} tracked insertions and {dels} tracked deletions as shown.")
    comments = len(document.findall(".//w:commentRangeStart", NS))
    if comments:
        warnings.append(f"{comments} Word comments were not imported.")
    if imp.notes:
        warnings.append(f"Moved {len(imp.notes)} footnotes to a Footnotes section at the end "
                        "(Claude Docs has no footnotes); each keeps a superscript marker in the text.")
    if imp.style_separators:
        warnings.append(f"Joined {imp.style_separators} heading/text pairs split by Word style separators.")
    if imp.typed_numbers:
        warnings.append(f"Converted {imp.typed_numbers} hand-typed numbers into structure.")
    if imp.bullets:
        warnings.append(f"Kept {imp.bullets} bulleted paragraphs as text.")

    doc = Document(blocks=imp.blocks, source=path.name, warnings=warnings, notes=imp.notes,
                   template={"front": imp.front})
    num.mark_restarts(doc)
    return doc
