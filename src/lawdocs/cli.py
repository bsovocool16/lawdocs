"""lawdocs command line.

    lawdocs import contract.docx -o out/     .docx -> model, Claude Docs markdown, outline
    lawdocs render out/model.json            renumber a model and print its markdown
    lawdocs outline out/model.json           print the outline
    lawdocs pull doc.md --base out/model.json -o out/model.json
                                             a Claude Doc's Markdown export -> model (keeps Word data)
    lawdocs export out/model.json -o x.docx  model -> .docx, formatted by the source's own template
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

from . import outline, render
from .docx_export import export_docx
from .docx_import import import_docx
from .md_import import pull
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
    # Keep the source package: it is the template for exporting back to Word.
    shutil.copyfile(args.docx, out / "source.docx")
    doc.template = dict(doc.template or {}, package="source.docx")
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


def cmd_export(args: argparse.Namespace) -> int:
    doc = _load(args.model)
    template = args.template
    if template is None:
        pkg = (doc.template or {}).get("package")
        if not pkg:
            print("No template: pass --template <file.docx>.", file=sys.stderr)
            return 2
        template = Path(args.model).parent / pkg
    out = export_docx(doc, template, args.out)
    print(f"Wrote {out} using {template} as the template")
    return 0


def cmd_pull(args: argparse.Namespace) -> int:
    base = _load(args.base)
    doc = pull(Path(args.markdown).read_text(), base, editor=args.editor)
    out = Path(args.out or args.base)
    if out.resolve() != Path(args.base).resolve() and (doc.template or {}).get("package"):
        # Keep the template reachable from the new model's folder.
        src = Path(args.base).parent / doc.template["package"]
        if src.exists() and not (out.parent / doc.template["package"]).exists():
            shutil.copyfile(src, out.parent / doc.template["package"])
    out.write_text(json.dumps(doc.to_dict(), indent=2, ensure_ascii=False))
    kept = sum(1 for b in doc.blocks if b.word)
    new = sum(1 for b in doc.blocks if not b.word)
    gone = len({b.id for b in base.blocks} - {b.id for b in doc.blocks})
    print(f"Pulled {args.markdown}: {len(doc.blocks)} blocks ({kept} matched, {new} new, {gone} removed) -> {out}")
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

    pp = sub.add_parser("pull", help="read a Claude Doc's Markdown export back into a model")
    pp.add_argument("markdown")
    pp.add_argument("--base", required=True, help="the model the doc was made from")
    pp.add_argument("-o", "--out", help="where to write the new model (default: overwrite --base)")
    pp.add_argument("--editor", default="docs", help="author to record on new and changed blocks")
    pp.set_defaults(func=cmd_pull)

    pe = sub.add_parser("export", help="write a model back to .docx")
    pe.add_argument("model")
    pe.add_argument("-o", "--out", required=True)
    pe.add_argument("--template", help="a .docx whose styles and numbering to use (default: the imported source)")
    pe.set_defaults(func=cmd_export)

    args = p.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
