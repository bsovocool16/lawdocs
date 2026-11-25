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

from .model import MAX_LEVEL, Block, Document

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
    return str(n)


# --------------------------------------------------------------------------
# Scheme
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Level:
    fmt: str  # numFmt name
    # How the level shows on its own line. "{n}" is this level's counter,
    # "{1}".."{6}" are ancestors' counters (decimal levels only).
    label: str


@dataclass(frozen=True)
class Scheme:
    levels: tuple[Level, ...]

    def label(self, counters: list[int], level: int) -> str:
        lv = self.levels[level - 1]
        out = lv.label.replace("{n}", format_counter(counters[level - 1], lv.fmt))
        for i in range(level):
            out = out.replace("{%d}" % (i + 1), format_counter(counters[i], self.levels[i].fmt))
        return out


LEGAL = Scheme(
    levels=(
        Level("decimal", "{1}."),
        Level("decimal", "{1}.{2}"),
        Level("lowerLetter", "({n})"),
        Level("lowerRoman", "({n})"),
        Level("upperLetter", "({n})"),
        Level("decimal", "({n})"),
    )
)


def apply(doc: Document, scheme: Scheme = LEGAL) -> dict[str, str]:
    """Number every clause in ``doc`` from its level. Returns {block id: reference}.

    The reference is the full citation form ("1.1(a)(i)"), which is what
    cross-references use; ``block.number`` is the short on-the-line label.
    Skipped levels (a level-3 clause directly under a level-1 clause) are
    tolerated: the missing counter stays at 0 and renders as such, which makes
    the structural problem visible instead of hiding it.
    """
    counters = [0] * MAX_LEVEL
    refs: dict[str, str] = {}
    for block in doc.blocks:
        if block.kind != "clause":
            continue
        level = max(1, min(block.level, MAX_LEVEL))
        block.level = level
        counters[level - 1] += 1
        for i in range(level, MAX_LEVEL):
            counters[i] = 0
        block.number = scheme.label(counters, level)
        refs[block.id] = reference(counters, level, scheme)
    return refs


def reference(counters: list[int], level: int, scheme: Scheme = LEGAL) -> str:
    """Full citation for a clause: 1 / 1.1 / 1.1(a) / 1.1(a)(i)."""
    first = format_counter(counters[0], scheme.levels[0].fmt)
    if level == 1:
        return first
    # A skipped level (a paragraph (a) sitting directly under Section 1) has a
    # zero counter and is left out of the citation: "1(a)", not "1.0(a)".
    base = scheme.label(counters, 2).rstrip(".") if counters[1] else first
    tail = "".join(scheme.label(counters, i) for i in range(3, level + 1) if counters[i - 1])
    return base + tail


# --------------------------------------------------------------------------
# Reading numbers that already exist (Word list numbers, typed numbers)
# --------------------------------------------------------------------------

# Leading-number patterns, most specific first. Each maps to a structural level.
_ARTICLE = re.compile(r"^(?:article|section)\s+([0-9]+|[ivxlcdm]+)\b(?!\.[0-9])[.:]?\s*", re.I)
_DOTTED = re.compile(r"^([0-9]+(?:\.[0-9]+)+)\.?(?=\s|$)\s*")
_SINGLE = re.compile(r"^([0-9]+)[.)](?=\s|$)\s*")
_PAREN = re.compile(r"^\(([0-9]+|[a-z]{1,3}|[A-Z]{1,3})\)\s*")

# "a." / "iv)" style list labels: only trusted when Word generated them, never
# split off typed text (too easily a person's initial: "A. Smith").
_LETTER_DOT = re.compile(r"([0-9]+|[a-z]{1,3}|[A-Z]{1,3})[.)]")

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
            return label, rest
    return None


def renumber(blocks: list[Block], scheme: Scheme = LEGAL) -> Document:
    doc = Document(blocks=blocks)
    apply(doc, scheme)
    return doc
