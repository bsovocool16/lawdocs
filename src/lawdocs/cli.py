"""lawdocs command line.

    lawdocs import contract.docx -o out/     .docx -> model, Claude Docs markdown, outline
    lawdocs render out/model.json            renumber a model and print its markdown
    lawdocs outline out/model.json           print the outline
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import outline, render
from .docx_import import import_docx
from .model import Document
from .numbering import apply


def _load(path: str) -> Document:
    doc = Document.from_dict(json.loads(Path(path).read_text()))
    apply(doc)
    return doc


def cmd_import(args: argparse.Namespace) -> int:
    doc = import_docx(args.docx)
    out = Path(args.out or Path(args.docx).with_suffix(""))
    out.mkdir(parents=True, exist_ok=True)
    (out / "model.json").write_text(json.dumps(doc.to_dict(), indent=2, ensure_ascii=False))
    (out / "doc.md").write_text(render.to_markdown(doc))
    (out / "outline.md").write_text(outline.to_markdown(doc))
    (out / "outline.json").write_text(json.dumps(outline.build(doc), indent=2, ensure_ascii=False))
    (out / "blockmap.json").write_text(json.dumps(render.block_map(doc), indent=2))
    n_clauses = len(doc.clauses())
    print(f"Imported {args.docx}: {len(doc.blocks)} blocks, {n_clauses} numbered clauses -> {out}/")
    for w in doc.warnings:
        print(f"  note: {w}")
    return 0


def cmd_render(args: argparse.Namespace) -> int:
    sys.stdout.write(render.to_markdown(_load(args.model)))
    return 0


def cmd_outline(args: argparse.Namespace) -> int:
    doc = _load(args.model)
    if args.json:
        print(json.dumps(outline.build(doc), indent=2, ensure_ascii=False))
    else:
        sys.stdout.write(outline.to_markdown(doc, max_level=args.depth))
    return 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="lawdocs", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    pi = sub.add_parser("import", help="import a .docx")
    pi.add_argument("docx")
    pi.add_argument("-o", "--out", help="output folder (default: next to the .docx)")
    pi.set_defaults(func=cmd_import)

    pr = sub.add_parser("render", help="print Claude Docs markdown for a model")
    pr.add_argument("model")
    pr.set_defaults(func=cmd_render)

    po = sub.add_parser("outline", help="print the document outline")
    po.add_argument("model")
    po.add_argument("--json", action="store_true")
    po.add_argument("--depth", type=int, default=6)
    po.set_defaults(func=cmd_outline)

    args = p.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
