"""The lawdocs block model.

A document is a flat, ordered list of blocks. Structure lives in two fields
only: ``kind`` and ``level``. There is no per-run formatting beyond emphasis,
no indents, no spacing, no styles. Everything a reader sees (numbers,
indentation, heading treatment) is *derived* from kind + level by a renderer,
so it can never drift out of sync the way Word's direct formatting does.

Kinds
-----
``title``    the document title (at most one, at the top)
``clause``   a numbered provision; ``level`` 1..9. By default 1 / 1.1 / (a) / (i) / (A) / (1);
             a document imported from Word keeps its own scheme (e.g. 1 / 1.1 / 1.1.1 / (a))
``text``     an unnumbered paragraph; ``level`` is the level of the clause it continues
             (0 = not inside any clause, e.g. recitals or a preamble)
``heading``  an unnumbered heading (e.g. "RECITALS", "Schedule 1")
``table``    a table; ``rows`` holds cell text

Inline text is a list of ``Run`` objects so bold/italic survive (defined terms
are usually bold). Underline, fonts, colours and sizes are deliberately dropped.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any

KINDS = ("title", "clause", "text", "heading", "table")
MAX_LEVEL = 9


@dataclass
class Run:
    text: str
    bold: bool = False
    italic: bool = False
    # Read on import only, as a caption signal ("Capitalization." underlined);
    # never stored or rendered.
    underline: bool = field(default=False, compare=False, repr=False)
    # A footnote marker, written as superscript digits ("³"), kept as its own run.
    note: bool = field(default=False, compare=False, repr=False)

    def same_format(self, other: "Run") -> bool:
        if self.note or other.note:
            return False
        return (self.bold, self.italic, self.underline) == (other.bold, other.italic, other.underline)

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"text": self.text}
        if self.bold:
            d["bold"] = True
        if self.italic:
            d["italic"] = True
        if self.note:
            d["note"] = True
        return d


@dataclass
class Block:
    id: str
    kind: str
    level: int = 0
    runs: list[Run] = field(default_factory=list)
    # Rendered legal number, e.g. "1.1" or "(a)". Filled in by numbering.apply().
    number: str | None = None
    # Number as it appeared in the source (Word's computed list number or a
    # typed number). Kept so cross-references in the source can be remapped.
    source_number: str | None = None
    # Clause heading, when the clause starts with a short title ("Payment.").
    caption: str | None = None
    rows: list[list[str]] | None = None
    # Who wrote this block: "import:<file>", "user", "claude". Phase 2 keeps this current.
    author: str | None = None
    # Word round-trip data captured on import: the source paragraph's raw
    # properties (pPr XML), caption form, spacer paragraphs that followed it,
    # raw table XML. Never shown in Claude Docs; used only by docx_export to
    # rebuild the paragraph exactly as the template formats it.
    word: dict[str, Any] | None = None
    # The clause starts its level's count again at 1 (a schedule or part with
    # its own numbering). Structural, so it survives edits around it.
    restart: bool = False
    # An alternative version of the clause before it at the same level (model
    # forms offer "[Alternative]" clauses): it takes that clause's number.
    alternative: bool = False
    # The count runs on from the last clause at this level under an earlier
    # parent instead of starting again at 1 (a Word list that isn't restarted).
    continues: bool = False
    # How the clause's number is written, with its counters as slots:
    # "{1:decimal}.{2:decimal}.{n:decimal}" for "1.1.1", "Article {n:upperRoman}".
    # None = like the clause before it at the same level. Set on import so a
    # Word document keeps its own forms; the counters are always recomputed.
    style: str | None = None

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"id": self.id, "kind": self.kind, "level": self.level}
        if self.runs:
            d["runs"] = [r.to_dict() for r in self.runs]
        for key in ("number", "source_number", "style", "caption", "rows", "author", "word"):
            val = getattr(self, key)
            if val is not None:
                d[key] = val
        if self.restart:
            d["restart"] = True
        if self.alternative:
            d["alternative"] = True
        if self.continues:
            d["continues"] = True
        return d

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "Block":
        return cls(
            id=d["id"],
            kind=d["kind"],
            level=d.get("level", 0),
            runs=[Run(**r) for r in d.get("runs", [])],
            number=d.get("number"),
            source_number=d.get("source_number"),
            caption=d.get("caption"),
            rows=d.get("rows"),
            author=d.get("author"),
            word=d.get("word"),
            restart=bool(d.get("restart", False)),
            alternative=bool(d.get("alternative", False)),
            continues=bool(d.get("continues", False)),
            style=d.get("style"),
        )


@dataclass
class Document:
    blocks: list[Block]
    source: str | None = None
    warnings: list[str] = field(default_factory=list)
    # Footnotes, in order of first reference: [{"n": 1, "text": "..."}].
    # Claude Docs has no footnotes, so they render as superscript markers plus a
    # Footnotes section at the end.
    notes: list[dict[str, Any]] = field(default_factory=list)
    # Word round trip: where the source package is kept, and spacer paragraphs
    # that came before the first block.
    template: dict[str, Any] | None = None

    def to_dict(self) -> dict[str, Any]:
        d = {
            "format": "lawdocs/1",
            "source": self.source,
            "warnings": self.warnings,
            "blocks": [b.to_dict() for b in self.blocks],
        }
        if self.notes:
            d["notes"] = self.notes
        if self.template:
            d["template"] = self.template
        return d

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "Document":
        return cls(
            blocks=[Block.from_dict(b) for b in d["blocks"]],
            source=d.get("source"),
            warnings=list(d.get("warnings", [])),
            notes=list(d.get("notes", [])),
            template=d.get("template"),
        )

    def clauses(self) -> list[Block]:
        return [b for b in self.blocks if b.kind == "clause"]


_SUP = str.maketrans("0123456789", "\u2070\u00b9\u00b2\u00b3\u2074\u2075\u2076\u2077\u2078\u2079")
_UNSUP = {v: k for k, v in zip("0123456789", "\u2070\u00b9\u00b2\u00b3\u2074\u2075\u2076\u2077\u2078\u2079")}
SUPERSCRIPT_DIGITS = "".join(_UNSUP)


def note_marker(n: int) -> str:
    """Footnote 12 is written "¹²": unmistakable in text full of "[10]"-style blanks."""
    return str(n).translate(_SUP)


def note_number(marker: str) -> int | None:
    digits = "".join(_UNSUP.get(c, "") for c in marker.strip())
    return int(digits) if digits and len(digits) == len(marker.strip()) else None


__all__ = ["Block", "Document", "Run", "KINDS", "MAX_LEVEL", "asdict", "note_marker", "note_number"]


def signature(kind: str, level: int, caption: str | None, runs) -> str:
    """What a block says, ignoring run splits and spacing, so an untouched
    paragraph is recognised after a trip through Claude Docs."""
    text = " ".join("".join(r.text for r in runs).split())
    return f"{kind}|{level}|{' '.join((caption or '').split())}|{text}"
