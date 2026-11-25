"""Run lawdocs over a folder of .docx files and report how well each survives.

    python tools/corpus_check.py corpus/ -o corpus-report/

For each document:

1. import it (structure, numbering, warnings);
2. export it straight back to Word with no edits;
3. re-import the export and compare clause references and text;
4. if LibreOffice and pdftoppm are installed, render the source and the
   export to PDF and compare every page pixel by pixel.

It also lists signs that the structure was misread: clauses whose numbers
skip a level, paragraphs that look numbered but stayed plain text, and
paragraphs that had to be rebuilt instead of kept as they were.

Writes report.md and report.json to the output folder. Needs Pillow only for
the page comparison.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
from pathlib import Path

from lawdocs import outline
from lawdocs.docx_export import export_docx
from lawdocs.docx_import import import_docx
from lawdocs.numbering import split_typed_number

_LOOKS_NUMBERED = re.compile(r"^\s*(?:\(?[0-9]+(?:\.[0-9]+)*[.)]?|\([a-zA-Z]{1,4}\)|(?:Section|Article)\s+[0-9IVX]+)\s")


def _refs(doc):
    return [e["ref"] for e in outline.build(doc) if e["ref"]]


def _texts(doc):
    return [(b.kind, " ".join((b.caption or "").split()), " ".join(b.text.split())) for b in doc.blocks]


def _render(docx: Path, out_dir: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx)],
                   capture_output=True, timeout=600)
    pdf = out_dir / (docx.stem + ".pdf")
    return pdf if pdf.exists() else None


def _compare_pdfs(a: Path, b: Path, work: Path) -> dict:
    try:
        from PIL import Image, ImageChops
    except ImportError:
        return {"skipped": "Pillow not installed"}
    if not shutil.which("pdftoppm"):
        return {"skipped": "pdftoppm not installed"}
    for pdf, prefix in ((a, "a"), (b, "b")):
        subprocess.run(["pdftoppm", "-r", "50", "-gray", "-png", str(pdf), str(work / prefix)], check=True)
    pa = sorted(work.glob("a-*.png"))
    pb = sorted(work.glob("b-*.png"))
    differing = []
    for i, (x, y) in enumerate(zip(pa, pb), 1):
        ix, iy = Image.open(x), Image.open(y)
        if ix.size != iy.size or ImageChops.difference(ix, iy).point(lambda v: 255 if v > 40 else 0).getbbox():
            differing.append(i)
    return {"pages_source": len(pa), "pages_export": len(pb), "differing_pages": differing}


def check(path: Path, out: Path, render: bool) -> dict:
    r: dict = {"file": path.name}
    try:
        doc = import_docx(path)
    except Exception as e:  # noqa: BLE001 - the report records any failure
        r["error"] = f"import failed: {e}"
        r["trace"] = traceback.format_exc()
        return r
    clauses = doc.clauses()
    r["blocks"] = len(doc.blocks)
    r["clauses"] = len(clauses)
    r["footnotes"] = len(doc.notes)
    r["warnings"] = doc.warnings
    r["skipped_levels"] = [c.number for c in clauses if c.number and re.search(r"\b0\b|\(0\)", c.number)]
    r["plain_but_numbered"] = [
        b.text[:80] for b in doc.blocks
        if b.kind == "text" and (_LOOKS_NUMBERED.match(b.text) or split_typed_number(b.text))
    ][:20]
    r["level_histogram"] = {str(lv): sum(1 for c in clauses if c.level == lv) for lv in range(1, 7)}

    work = Path(tempfile.mkdtemp(prefix="lawdocs-"))
    try:
        exported = out / f"{path.stem}.roundtrip.docx"
        stats: dict = {}
        try:
            export_docx(doc, path, exported, stats)
        except Exception as e:  # noqa: BLE001
            r["error"] = f"export failed: {e}"
            r["trace"] = traceback.format_exc()
            return r
        r["kept_verbatim"] = stats.get("verbatim")
        r["rebuilt"] = stats.get("rebuilt")
        try:
            again = import_docx(exported)
            r["refs_match"] = _refs(again) == _refs(doc)
            r["text_match"] = _texts(again) == _texts(doc)
        except Exception as e:  # noqa: BLE001
            r["error"] = f"re-import of export failed: {e}"
            return r
        if render:
            src_copy = work / ("source-" + path.name)
            shutil.copyfile(path, src_copy)
            a = _render(src_copy, work)
            b = _render(exported, work)
            r["pages"] = _compare_pdfs(a, b, work) if a and b else {"skipped": "LibreOffice not available or failed"}
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return r


def verdict(r: dict) -> str:
    if "error" in r:
        return "FAIL"
    pages = r.get("pages") or {}
    if not r.get("refs_match") or not r.get("text_match"):
        return "FAIL"
    if pages.get("differing_pages") or pages.get("pages_source") != pages.get("pages_export") and "skipped" not in pages:
        # Rebuilt paragraphs (typed numbers turned into Word numbering) and
        # accepted tracked changes are meant to change the page.
        intended = r.get("rebuilt") or any("tracked" in w for w in r.get("warnings", []))
        return "NORMALISED" if intended else "LAYOUT"
    if r["skipped_levels"] or r["plain_but_numbered"]:
        return "CHECK"
    return "PASS"


def report_md(results: list[dict]) -> str:
    lines = ["# lawdocs corpus report", ""]
    counts: dict[str, int] = {}
    for r in results:
        counts[verdict(r)] = counts.get(verdict(r), 0) + 1
    lines.append(", ".join(f"{k}: {v}" for k, v in sorted(counts.items())) + f" (of {len(results)})")
    lines += ["", "PASS: identical text, structure and pages. CHECK: round trip fine but the structure may be "
              "misread. NORMALISED: pages differ because lawdocs fixed something (typed numbers, tracked changes). "
              "LAYOUT: pages differ for no intended reason. FAIL: error or text/structure lost.", "",
              "| File | Verdict | Clauses | Kept / rebuilt | Pages differing | Notes |", "| --- | --- | --- | --- | --- | --- |"]
    for r in results:
        pages = r.get("pages") or {}
        pd = pages.get("skipped") or (f"{len(pages.get('differing_pages', []))} of {pages.get('pages_source')}"
                                      + (f" (export has {pages.get('pages_export')})"
                                         if pages.get("pages_export") != pages.get("pages_source") else ""))
        notes = r.get("error") or "; ".join(
            x for x in [f"{len(r['skipped_levels'])} skipped levels" if r.get("skipped_levels") else "",
                        f"{len(r['plain_but_numbered'])} numbered-looking plain paragraphs" if r.get("plain_but_numbered") else ""]
            if x)
        lines.append(f"| {r['file']} | {verdict(r)} | {r.get('clauses', '')} | "
                     f"{r.get('kept_verbatim', '')} / {r.get('rebuilt', '')} | {pd if pages else ''} | {notes} |")
    for r in results:
        if verdict(r) == "PASS":
            continue
        lines += ["", f"## {r['file']}", ""]
        if "error" in r:
            lines.append(f"Error: {r['error']}")
            continue
        if r["skipped_levels"]:
            lines.append("Skipped levels: " + ", ".join(r["skipped_levels"][:20]))
        for t in r["plain_but_numbered"]:
            lines.append(f"- looks numbered, kept as text: {t}")
        if (r.get("pages") or {}).get("differing_pages"):
            lines.append("Differing pages: " + ", ".join(map(str, r["pages"]["differing_pages"][:40])))
        for w in r["warnings"]:
            lines.append(f"- note: {w}")
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("folder")
    p.add_argument("-o", "--out", default="corpus-report")
    p.add_argument("--no-render", action="store_true", help="skip the LibreOffice page comparison")
    args = p.parse_args(argv)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    files = sorted(f for f in Path(args.folder).rglob("*.docx") if not f.name.startswith("~$"))
    results = []
    for f in files:
        print(f"checking {f.name} ...", file=sys.stderr)
        results.append(check(f, out, render=not args.no_render))
        print(f"  {verdict(results[-1])}", file=sys.stderr)
    (out / "report.json").write_text(json.dumps(results, indent=2, ensure_ascii=False))
    (out / "report.md").write_text(report_md(results))
    print(report_md(results))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
