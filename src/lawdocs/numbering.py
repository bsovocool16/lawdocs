"""Legal numbering: 1 / 1.1 / (a) / (i) / (A) / (1).

Numbers are never stored as text the author types. They are computed from
block levels every time the document is rendered, so inserting, deleting or
moving a clause can't leave a stale number behind.

The scheme is configurable, but the default is the one most US and UK
transactional lawyers already use::

    Level 1   1.        Article / section
    Level 2   1.1       Sub-section
    Level 3   (a)       Paragraph
    Level 4   (i)       Sub-paragraph
    Level 5   (A)       Clause
    Level 6   (1)       Sub-clause

Levels 3+ are referred to in full in cross-references: "Section 1.1(a)(i)".
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .model import MAX_LEVEL, Block, Document, signature

# --------------------------------------------------------------------------
# Counter formatting
# --------------------------------------------------------------------------

_ROMAN = [
    (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
    (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
]


def to_roman(n: int) -> str:
    if n <= 0:
        return str(n)
    out = []
    for value, sym in _ROMAN:
        while n >= value:
            out.append(sym)
            n -= value
    return "".join(out)


def from_roman(s: str) -> int | None:
    s = s.lower()
    if not s or not re.fullmatch(r"m*(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})", s):
        return None
    vals = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}
    total = 0
    for i, ch in enumerate(s):
        v = vals[ch]
        if i + 1 < len(s) and vals[s[i + 1]] > v:
            total -= v
        else:
            total += v
    return total


def to_alpha(n: int) -> str:
    """1 -> a, 26 -> z, 27 -> aa, 28 -> bb (the legal convention, also Word's)."""
    if n <= 0:
        return str(n)
    letter = chr(ord("a") + (n - 1) % 26)
    return letter * ((n - 1) // 26 + 1)


def from_alpha(s: str) -> int | None:
    s = s.lower()
    if not s or not s.isalpha() or len(set(s)) != 1:
        return None
    return (len(s) - 1) * 26 + (ord(s[0]) - ord("a") + 1)


def format_counter(n: int, fmt: str) -> str:
    """Format a counter with a Word-style numFmt name."""
    if fmt in ("decimal", "decimalZero") or fmt is None:
        return f"{n:02d}" if fmt == "decimalZero" and n < 10 else str(n)
    if fmt == "lowerLetter":
        return to_alpha(n)
    if fmt == "upperLetter":
        return to_alpha(n).upper()
    if fmt == "lowerRoman":
        return to_roman(n)
    if fmt == "upperRoman":
        return to_roman(n).upper()
    if fmt == "none":
        return ""
    if fmt == "ordinalText":
        return _ordinal_text(n)
    if fmt == "cardinalText":
        return _cardinal_text(n)
    if fmt == "ordinal":
        suffix = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
        return f"{n}{suffix}"
    return str(n)


_ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve",
         "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
_ORD = {"One": "First", "Two": "Second", "Three": "Third", "Five": "Fifth", "Eight": "Eighth", "Nine": "Ninth",
        "Twelve": "Twelfth"}


def _cardinal_text(n: int) -> str:
    """Word's cardinalText: 21 -> "Twenty-One" (to 99; larger counters stay digits)."""
    if n <= 0 or n >= 100:
        return str(n)
    if n < 20:
        return _ONES[n]
    return _TENS[n // 10] + ("-" + _ONES[n % 10] if n % 10 else "")


def _ordinal_text(n: int) -> str:
    """Word's ordinalText ("FIRST:" articles in a charter): 4 -> "Fourth", 21 -> "Twenty-First"."""
    words = _cardinal_text(n)
    if words.isdigit():
        return words
    head, _, last = words.rpartition("-")
    if last in _ORD:
        last = _ORD[last]
    elif last.endswith("y"):
        last = last[:-1] + "ieth"
    else:
        last += "th"
    return (head + "-" if head else "") + last


# --------------------------------------------------------------------------
# Scheme
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Level:
    fmt: str  # numFmt name
    # How the level shows on its own line. "{n}" is this level's counter,
    # "{1}".."{9}" are ancestors' counters in their own level's format, and
    # "{1:decimal}" an ancestor's counter in a given format (so "Article I"
    # can have sections "1.01").
    label: str


_SLOT = re.compile(r"\{(n|[1-9])(?::([A-Za-z]+))?\}")


@dataclass(frozen=True)
class Scheme:
    levels: tuple[Level, ...]

    def level(self, level: int) -> Level:
        if level <= len(self.levels):
            return self.levels[level - 1]
        return LEGAL_LEVELS[min(level, len(LEGAL_LEVELS)) - 1]

    def label(self, counters: list[int], level: int) -> str:
        return render_style(self.level(level).label, counters, level, self)


LEGAL_LEVELS = (
    Level("decimal", "{1}."),
    Level("decimal", "{1}.{2}"),
    Level("lowerLetter", "({n})"),
    Level("lowerRoman", "({n})"),
    Level("upperLetter", "({n})"),
    Level("decimal", "({n})"),
    Level("lowerLetter", "({n})"),
    Level("lowerRoman", "({n})"),
    Level("upperLetter", "({n})"),
)
LEGAL = Scheme(levels=LEGAL_LEVELS[:6])


def render_style(style: str, counters: list[int], level: int, scheme: Scheme = LEGAL) -> str:
    """Fill a label template ("{1:decimal}.{2:decimal}.{n:decimal}") with counters."""
    def sub(m: re.Match) -> str:
        k = level if m.group(1) == "n" else int(m.group(1))
        fmt = m.group(2) or scheme.level(k).fmt
        return format_counter(counters[k - 1], fmt)

    return _SLOT.sub(sub, style)


def _cites_ancestors(style: str, level: int) -> bool:
    return any(m.group(1) != "n" and int(m.group(1)) < level for m in _SLOT.finditer(style))


def apply(doc: Document, scheme: Scheme = LEGAL) -> dict[str, str]:
    """Number every clause in ``doc`` from its level. Returns {block id: reference}.

    The reference is the full citation form ("1.1(a)(i)"), which is what
    cross-references use; ``block.number`` is the short on-the-line label.

    Only the counters come from structure. How a clause's number is written is
    its ``style``: a clause imported from Word keeps the form its source used
    ("1.1.1", "Article 4", "(1)"), and a clause without one takes the form of
    the clause before it at the same level, or the scheme's. A clause marked
    ``restart`` starts its level again at 1, as a schedule or part with its
    own numbering does.

    Skipped levels (a level-3 clause directly under a level-1 clause) are
    tolerated: the missing counter stays at 0 and is left out of citations.
    """
    counters = [0] * MAX_LEVEL
    last = [0] * MAX_LEVEL  # each level's latest value, kept across parents
    styles: dict[int, str] = {}
    refs_at: dict[int, str] = {}
    refs: dict[str, str] = {}
    for block in doc.blocks:
        if block.kind != "clause":
            continue
        level = max(1, min(block.level, MAX_LEVEL))
        block.level = level
        if block.restart:
            counters[level - 1] = 0
        elif block.continues:
            counters[level - 1] = last[level - 1]
        if not (block.alternative and counters[level - 1]):
            counters[level - 1] += 1
        last[level - 1] = counters[level - 1]
        for i in range(level, MAX_LEVEL):
            counters[i] = 0
        style = block.style or styles.get(level) or scheme.level(level).label
        styles[level] = style
        block.number = render_style(style, counters, level, scheme)
        kept = _kept_typed(block, level)
        if kept and _norm_label(kept) != _norm_label(block.number):
            # Typed by hand in the source and not edited since: shown as typed,
            # even where the counting disagrees (a typo such as 23.4 inside
            # clause 20). Editing the clause gives it its computed number.
            block.number = kept
            ref = _bare(kept) if len(_counters_in(kept)) > 1 or level == 1 else refs_at.get(level - 1, "") + _cited_part(kept)
            refs_at[level] = ref
            refs[block.id] = ref
            continue
        if level == 1 or _cites_ancestors(style, level):
            ref = _bare(block.number)
        else:
            # A skipped level (a paragraph (a) sitting directly under Section 1)
            # is left out of the citation: "1(a)", not "1.0(a)".
            parent = level - 1
            while parent > 1 and not counters[parent - 1]:
                parent -= 1
            ref = refs_at.get(parent, "") + _cited_part(block.number)
        refs_at[level] = ref
        refs[block.id] = ref
    return refs


def _kept_typed(block: Block, level: int) -> str | None:
    """The number typed in the source paragraph, when the block is unedited."""
    w = block.word or {}
    typed = w.get("typed")
    if not typed or not typed.get("kept") or w.get("sig") != signature(block.kind, level, block.caption, block.runs):
        return None
    return typed["label"]


def keep_typed_quirks(doc: Document) -> int:
    """Mark typed numbers that disagree with the count (a source typo such as
    23.4 inside clause 20), so they are shown as typed until the clause is
    edited. Typed numbers that agree keep following the structure, so deleting
    3.1 still renumbers an untouched 3.2. Returns how many were marked."""
    apply(doc)
    n = 0
    for b in doc.blocks:
        typed = (b.word or {}).get("typed")
        if b.kind == "clause" and typed and _norm_label(typed["label"]) != _norm_label(b.number or ""):
            typed["kept"] = True
            n += 1
    if n:
        apply(doc)
        doc.warnings.append(f"Kept {n} typed numbers that don't follow the count as typed.")
    return n


def _norm_label(label: str) -> str:
    return re.sub(r"\s+", " ", label).strip().rstrip(".:").lower()


def _bare(label: str) -> str:
    """A label as cited: no leading word ("Article 4" -> "4") or trailing dot."""
    label = label.strip()
    if re.match(r"^[A-Za-z]{2,}\s", label):
        label = re.sub(r"^[A-Za-z]+\s+", "", label)
    return label.rstrip(".:")


def _cited_part(label: str) -> str:
    """How a sub-clause's own label reads inside a citation: "(a)" for "(a)",
    "a." or "a)"; "1.1" stays "1.1"."""
    bare = _bare(label).rstrip(")")
    if re.fullmatch(r"\(?[A-Za-z]{1,8}|\(?[0-9]+", bare) and not re.fullmatch(r"[0-9]+", bare):
        return "(" + bare.lstrip("(") + ")"
    if re.fullmatch(r"\([0-9]+", bare):
        return bare + ")"
    return _bare(label)


def reference(counters: list[int], level: int, scheme: Scheme = LEGAL) -> str:
    """Full citation for a clause numbered by ``scheme``: 1 / 1.1 / 1.1(a) / 1.1(a)(i)."""
    style = scheme.level(level).label
    label = render_style(style, counters, level, scheme)
    if level == 1 or _cites_ancestors(style, level):
        return _bare(label)
    parent = level - 1
    while parent > 1 and not counters[parent - 1]:
        parent -= 1
    return reference(counters, parent, scheme) + _bare(label)


# --------------------------------------------------------------------------
# Reading numbers that already exist (Word list numbers, typed numbers)
# --------------------------------------------------------------------------

# Leading-number patterns, most specific first. Each maps to a structural level.
_ARTICLE = re.compile(r"^(?:article|section)\s+([0-9]+|[ivxlcdm]+)\b(?!\.[0-9])[.:]?\s*", re.I)
_DOTTED = re.compile(r"^([0-9]+(?:\.[0-9]+)+)\.?(?=\s|$)\s*")
_SINGLE = re.compile(r"^([0-9]+)[.)](?=\s|$)\s*")
# Letters run to three characters ("(aaa)"); roman numerals run longer ("(viii)", "(xxxviii)").
_PAREN = re.compile(r"^\(([0-9]+|[ivx]{1,8}|[IVX]{1,8}|[a-z]{1,3}|[A-Z]{1,3})\)\s*")

# "a." / "iv)" style list labels: only trusted when Word generated them, never
# split off typed text (too easily a person's initial: "A. Smith").
_LETTER_DOT = re.compile(r"([0-9]+|[ivx]{1,8}|[IVX]{1,8}|[a-z]{1,3}|[A-Z]{1,3})[.)]")

# Contract sub-paragraphs never run past (xxxix), so only i, v and x make a
# label roman; "(c)" or "(d)" is always a letter.
_ROMAN_LIKE = set("ivx")


def classify_label(label: str, prev_alpha_lower: str | None = None,
                   prev_alpha_upper: str | None = None) -> int | None:
    """Return the structural level a source number implies, or None.

    ``prev_alpha_*`` is the last lower/upper letter label seen at the letter
    level, used to settle "(i)": after "(h)" it is a letter, otherwise roman.
    """
    label = label.strip()
    if _ARTICLE.match(label + " "):
        return 1
    m = _DOTTED.match(label + " ")
    if m:
        return min(m.group(1).count(".") + 1, MAX_LEVEL)
    if _SINGLE.match(label + " ") or re.fullmatch(r"[0-9]+", label):
        return 1
    m = _PAREN.match(label) or _LETTER_DOT.fullmatch(label)
    if m:
        tok = m.group(1)
        if tok.isdigit():
            return 6
        if tok.islower():
            if set(tok) <= _ROMAN_LIKE and from_roman(tok) is not None:
                n_alpha = from_alpha(tok)
                if prev_alpha_lower is not None and n_alpha is not None:
                    prev = from_alpha(prev_alpha_lower)
                    if prev is not None and n_alpha == prev + 1:
                        return 3
                return 4
            return 3
        # (A) and the rare upper-roman (I) both sit at level 5 in the legal scheme.
        return 5
    return None


def split_typed_number(text: str) -> tuple[str, str] | None:
    """Split a typed leading number off paragraph text: "1.1 Foo" -> ("1.1", "Foo")."""
    for pat in (_ARTICLE, _DOTTED, _SINGLE, _PAREN):
        m = pat.match(text)
        if m:
            label = text[: m.end()].strip()
            rest = text[m.end():]
            if not rest.strip():
                return None
            if pat is _ARTICLE and not label.endswith((".", ":")) and re.match(r"\s*(?:[a-z(]|[-\u2013\u2014]\s)", rest):
                # "Section 2115 of the California Corporations Code ...", "Article 5(1) - This
                # provides ...", "Article 2 - Insert the name": a citation or a note about
                # that article, not the article itself.
                return None
            return label, rest
    return None


def renumber(blocks: list[Block], scheme: Scheme = LEGAL) -> Document:
    doc = Document(blocks=blocks)
    apply(doc, scheme)
    return doc


# --------------------------------------------------------------------------
# Levels and label styles from existing numbers
# --------------------------------------------------------------------------

_COUNTER = re.compile(r"[0-9]+|[A-Za-z]+")
_WORD_FMTS = {"decimal", "decimalZero", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman",
              "ordinalText", "cardinalText", "ordinal"}


def _token_fmt(tok: str, roman: bool) -> str | None:
    if tok.isdigit():
        return "decimalZero" if len(tok) == 2 and tok.startswith("0") else "decimal"
    if roman and from_roman(tok) is not None:
        return "lowerRoman" if tok.islower() else "upperRoman"
    if from_alpha(tok) is not None and len(tok) <= 3:
        return "lowerLetter" if tok.islower() else "upperLetter"
    if from_roman(tok) is not None:
        return "lowerRoman" if tok.islower() else "upperRoman"
    return None  # a word ("Article", "Schedule"): literal text


def _counters_in(label: str) -> list[re.Match]:
    out = []
    for m in _COUNTER.finditer(label):
        tok = m.group(0)
        if not tok.isdigit() and m.start() and label[m.start() - 1].isdigit():
            continue  # "5A.": the letter is part of the text, not a counter
        if tok.isdigit() or len(tok) <= 3 or (set(tok.lower()) <= set("ivx") and len(tok) <= 8):
            out.append(m)
    return out


def _roman_label(label: str, fmt: str | None) -> bool:
    if fmt:
        return fmt.endswith("Roman")
    if classify_label(label) == 4:
        return True
    # "Article IV", "Section II": a roman numeral after a word.
    return bool(re.match(r"^[A-Za-z]{2,}\s+[IVXivx]+\b", label))


def label_style(label: str, level: int, fmt: str | None = None) -> str | None:
    """The template a source label follows, for ``Block.style``.

    The last counter in the label is the clause's own (in ``fmt``, Word's numFmt
    when Word drew the label); counters before it are its ancestors', nearest
    first: "2.3.1" at level 3 is "{1:decimal}.{2:decimal}.{n:decimal}". Words and
    punctuation stay as they are ("Article {n:upperRoman}", "5A.{2:decimal}.{n:decimal}").
    """
    label = label.strip().lstrip("[").replace(" ", " ")
    if fmt in ("ordinalText", "cardinalText", "ordinal"):
        # "Fifth:", "Twenty-One", "3rd": the number is a word (or carries a suffix).
        m = re.search(r"[0-9]+(?:st|nd|rd|th)|[A-Za-z]+(?:-[A-Za-z]+)?", label)
        if not m:
            return None
        lit = lambda t: t.replace("{", "{{").replace("}", "}}")  # noqa: E731
        return lit(label[:m.start()]) + "{n:%s}" % fmt + lit(label[m.end():])
    toks = _counters_in(label)
    if not toks:
        return None
    roman_level = _roman_label(label, fmt)
    out, pos = [], 0
    for i, m in enumerate(toks):
        own = i == len(toks) - 1
        f = fmt if own and fmt in _WORD_FMTS else _token_fmt(m.group(0), roman=own and roman_level)
        if f is None:
            continue
        slot_level = level - (len(toks) - 1 - i)
        if slot_level < 1:
            continue  # more counters than levels: keep that part as written
        out.append(label[pos:m.start()].replace("{", "{{").replace("}", "}}"))
        out.append("{n:%s}" % f if own else "{%d:%s}" % (slot_level, f))
        pos = m.end()
    if pos == 0:
        return None
    out.append(label[pos:])
    style = "".join(out)
    return style


def counter_value(label: str, fmt: str | None = None) -> int | None:
    """The clause's own counter in a source label: "2.3.1" -> 1, "(c)" -> 3, "(iv)" -> 4."""
    label = label.strip().lstrip("[")
    if fmt in ("ordinalText", "cardinalText", "ordinal"):
        m = re.search(r"([0-9]+)(?:st|nd|rd|th)|[A-Za-z]+(?:-[A-Za-z]+)?", label)
        if m and m.group(1):
            return int(m.group(1))
        word = m.group(0).lower() if m else ""
        return next((k for k in range(1, 100) if format_counter(k, fmt).lower() == word), None)
    toks = _counters_in(label)
    if not toks:
        return None
    tok = toks[-1].group(0)
    if tok.isdigit():
        return int(tok)
    if _roman_label(label, fmt) and from_roman(tok) is not None:
        return from_roman(tok)
    return from_alpha(tok) or from_roman(tok)


def _relative(label: str) -> bool:
    """Labels like "(a)" or "A." say nothing about depth on their own; "1.1" and
    "Article 2" do."""
    label = label.strip().lstrip("[")
    return bool(_PAREN.match(label) or _LETTER_DOT.fullmatch(label)) and not re.fullmatch(r"[0-9]+[.)]", label)


class LevelReader:
    """Works out clause levels from the labels a document shows, in order.

    Labels that spell out their depth ("1.1.1", "Article 2") take it. Labels
    that don't ("(a)", "(1)", "(i)") nest by context: one that matches the form
    of a clause still open is its sibling; a new form goes one level below the
    deepest open clause unless its usual level is deeper still. So a "(1)"
    under clause "1" is its sub-clause, and an "(a)" under "1.1.1" goes below it
    rather than beside it.
    """

    def __init__(self) -> None:
        self.stack: dict[int, str] = {}
        self.prev_lower: str | None = None
        self.prev_upper: str | None = None

    def reset(self) -> None:
        self.stack.clear()
        self.prev_lower = self.prev_upper = None

    def read(self, label: str, fallback: int | None = None, fmt: str | None = None) -> int | None:
        core = label.strip().lstrip("[")
        # ".1.1": a list whose top level shows nothing still nests "1.1" at level 2.
        cl = classify_label(core.lstrip("."), self.prev_lower, self.prev_upper)
        if cl == 3 and fmt == "lowerRoman":
            cl = 4
        elif cl == 4 and fmt == "lowerLetter":
            cl = 3
        if cl is None:
            cl = fallback
        if cl is None:
            return None
        shape = _shape(core, cl)
        if _relative(core):
            hits = [k for k, s in self.stack.items() if s == shape]
            deepest = max(self.stack, default=0)
            if hits:
                level = hits[-1]
            else:
                level = deepest + 1 if cl <= deepest else cl
        else:
            level = cl
        level = max(1, min(level, MAX_LEVEL))
        self.stack = {k: v for k, v in self.stack.items() if k < level}
        self.stack[level] = shape
        m = _PAREN.match(core) or _LETTER_DOT.fullmatch(core)
        if m and m.group(1).isalpha():
            tok = m.group(1)
            if cl == 3 and tok.islower():
                self.prev_lower = tok
            elif cl == 5 and tok.isupper():
                self.prev_upper = tok
        if level <= 1:
            self.prev_lower = self.prev_upper = None
        return level


def _shape(label: str, cl: int) -> str:
    """The form of a label without its values: "(c)" and "(d)" share one."""
    def tok(m: re.Match) -> str:
        t = m.group(0)
        if t.isdigit():
            return "1"
        if len(t) > 3 and not set(t.lower()) <= set("ivx"):
            return t  # a word
        if t.islower():
            return "i" if cl == 4 else "a"
        return "A"

    return _COUNTER.sub(tok, label.rstrip())


def mark_restarts(doc: Document) -> int:
    """Flag clauses whose source number breaks the count, so the numbers
    lawdocs computes match the source. Returns how many were flagged.

    * ``restart``: the number starts its level again ("1." after "4.").
    * ``alternative``: the number repeats the clause before it at that level
      (a model form's "[Alternative] 4.4.4" after 4.4.4, or "5A" inserted
      after "5"), so it doesn't count.
    * ``continues``: the count runs on from the last clause at that level
      under an earlier parent, as a Word list does unless restarted.
    """
    apply(doc)
    n = 0
    counters = [0] * MAX_LEVEL
    last = [0] * MAX_LEVEL
    last_label: dict[int, str] = {}
    for b in doc.blocks:
        if b.kind != "clause":
            continue
        lv = b.level
        computed = counters[lv - 1] + 1
        own = re.search(r"\{n:([A-Za-z]+)\}", b.style or "")
        value = counter_value(b.source_number, own.group(1) if own else None) if b.source_number else None
        label = " ".join((b.source_number or "").replace("[", "").split()).rstrip(".")
        flagged = b.restart or b.alternative or b.continues
        if value is not None and not flagged and computed != 1:
            # Repeats the number before it: an alternative version ("[Alternative]
            # 4.4.4") or an inserted clause ("5A" after "5").
            if lv in last_label and value == computed - 1:
                b.alternative = True
                n += 1
            elif value == 1:
                b.restart = True
                n += 1
        elif value is not None and not flagged and computed == 1 and value > 1 and value == last[lv - 1] + 1:
            # A Word list that runs on under the next parent: "(5)" under
            # article 6 after "(1)"-"(4)" under article 5.
            b.continues = True
            n += 1
        if b.restart:
            computed = 1
        elif b.alternative:
            computed = max(1, counters[lv - 1])
        elif b.continues:
            computed = last[lv - 1] + 1
        counters[lv - 1] = computed
        last[lv - 1] = computed
        last_label[lv] = label
        for i in range(lv, MAX_LEVEL):
            counters[i] = 0
        for k in [k for k in last_label if k > lv]:
            del last_label[k]
    apply(doc)
    return n
