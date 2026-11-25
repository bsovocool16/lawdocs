"""The lawdocs block model.

A document is a flat, ordered list of blocks. Structure lives in two fields
only: ``kind`` and ``level``. There is no per-run formatting beyond emphasis,
no indents, no spacing, no styles. Everything a reader sees (numbers,
indentation, heading treatment) is *derived* from kind + level by a renderer,
so it can never drift out of sync the way Word's direct formatting does.

Kinds
-----
``title``    the document title (at most one, at the top)
``clause``   a numbered provision; ``level`` 1..6 maps to 1 / 1.1 / (a) / (i) / (A) / (1)
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
MAX_LEVEL = 6


@dataclass
class Run:
    text: str
    bold: bool = False
    italic: bool = False

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"text": self.text}
        if self.bold:
            d["bold"] = True
        if self.italic:
            d["italic"] = True
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

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"id": self.id, "kind": self.kind, "level": self.level}
        if self.runs:
            d["runs"] = [r.to_dict() for r in self.runs]
        for key in ("number", "source_number", "caption", "rows", "author"):
            val = getattr(self, key)
            if val is not None:
                d[key] = val
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
        )


@dataclass
class Document:
    blocks: list[Block]
    source: str | None = None
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "format": "lawdocs/1",
            "source": self.source,
            "warnings": self.warnings,
            "blocks": [b.to_dict() for b in self.blocks],
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "Document":
        return cls(
            blocks=[Block.from_dict(b) for b in d["blocks"]],
            source=d.get("source"),
            warnings=list(d.get("warnings", [])),
        )

    def clauses(self) -> list[Block]:
        return [b for b in self.blocks if b.kind == "clause"]


__all__ = ["Block", "Document", "Run", "KINDS", "MAX_LEVEL", "asdict"]
