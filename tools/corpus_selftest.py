"""Check that tools/corpus_check.py catches broken round trips.

    python tools/corpus_selftest.py corpus/ -o selftest-report/

corpus_check.py says a document survived when nothing differs between the
source and the export. That only means something if it would have noticed a
difference. This tool breaks each document on purpose, the way a bug in
lawdocs could, and runs the same comparisons on the broken copy:

    drop-clause     a long paragraph is removed from the export
    swap-clauses    two neighbouring paragraphs trade places in the export
    change-figure   one digit in the text changes (30 days -> 31 days)
    drop-word       one word goes missing from a paragraph
    renumber        a numbered paragraph moves one list level down
    nudge-margin    the left margin grows by 0.05 inch (text unchanged)
    md-drop         a paragraph is lost on the way back from Claude Docs
    md-figure       a digit changes on the way back from Claude Docs
    misnumber       lawdocs reads one clause's number wrong (checked
                    against LibreOffice's independent reading)

The unbroken export is the control. A break counts as caught when a check
that agreed for the control disagrees for the broken copy. It is MISSED when
every check still agrees, and n/a when the document has nothing to break that
way (no numbered paragraph with its own list settings, say) or the check that
should catch it already disagrees for the control.

Exits 1 if any break was missed. Page checks need LibreOffice, pdftoppm and
Pillow; with --no-render the margin and misnumber breaks are n/a.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import corpus_check as cc  # noqa: E402

from lawdocs import md_import  # noqa: E402
from lawdocs import render as md_render  # noqa: E402
from lawdocs.docx_export import export_docx  # noqa: E402
from lawdocs.docx_import import import_docx  # noqa: E402

DOC_XML = "word/document.xml"


# --- Paragraphs in document.xml -------------------------------------------------

_P_TAG = re.compile(r"<w:p\b[^>]*?(/?)>|</w:p>")
_T_TAG = re.compile(r"(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)")


@dataclass
class Para:
    start: int
    end: int
    xml: str

    @property
    def text(self) -> str:
        return "".join(m.group(2) for m in _T_TAG.finditer(self.xml))


def paragraphs(xml: str) -> list[Para]:
    """Outermost w:p elements, in order (a text box's paragraphs stay inside theirs)."""
    out, depth, start = [], 0, 0
    for m in _P_TAG.finditer(xml):
        if m.group(0) == "</w:p>":
            depth -= 1
            if depth == 0:
                out.append(Para(start, m.end(), xml[start:m.end()]))
        elif m.group(1):  # <w:p/>
            if depth == 0:
                out.append(Para(m.start(), m.end(), m.group(0)))
        else:
            if depth == 0:
                start = m.start()
            depth += 1
    return out


def _plain(p: Para) -> bool:
    """A paragraph safe to move or delete: real text, no nested paragraphs,
    fields, notes, bookmarks spanning elsewhere or section breaks."""
    x = p.xml
    opens = sum(1 for m in _P_TAG.finditer(x) if m.group(0) != "</w:p>")
    return (len(p.text.strip()) >= 60 and opens == 1
            and not any(t in x for t in ("<w:sectPr", "<w:fldChar", "<w:instrText", "Reference ", "<w:bookmarkStart",
                                         "<w:bookmarkEnd", "<w:commentRange", "<w:ins ", "<w:del ")))


def _middle(items: list):
    return items[len(items) // 2] if items else None


def _splice(xml: str, edits: list[tuple[int, int, str]]) -> str:
    for s, e, new in sorted(edits, reverse=True):
        xml = xml[:s] + new + xml[e:]
    return xml


# --- Breaks to the export ------------------------------------------------------
# Each takes document.xml and returns the broken XML, or None when the document
# has nothing to break that way.

def drop_clause(xml: str) -> str | None:
    p = _middle([p for p in paragraphs(xml) if _plain(p)])
    return _splice(xml, [(p.start, p.end, "")]) if p else None


def swap_clauses(xml: str) -> str | None:
    ps = paragraphs(xml)
    pairs = [(a, b) for a, b in zip(ps, ps[1:])
             if _plain(a) and _plain(b) and a.text != b.text and xml[a.end:b.start].strip() == ""]
    pair = _middle(pairs)
    if not pair:
        return None
    a, b = pair
    return _splice(xml, [(a.start, a.end, b.xml), (b.start, b.end, a.xml)])


def _edit_text(xml: str, pattern: str, repl) -> str | None:
    """Apply ``repl`` to the first match of ``pattern`` in one w:t of a paragraph
    near the middle of the document."""
    rx = re.compile(pattern)
    cands = [(p, m) for p in paragraphs(xml) if len(p.text.strip()) >= 40
             for m in _T_TAG.finditer(p.xml) if rx.search(m.group(2))]
    hit = _middle(cands)
    if not hit:
        return None
    p, m = hit
    new_t = m.group(1) + rx.sub(repl, m.group(2), count=1) + m.group(3)
    s = p.start + m.start()
    return xml[:s] + new_t + xml[s + len(m.group(0)):]


def change_figure(xml: str) -> str | None:
    return _edit_text(xml, r"\d", lambda m: str((int(m.group(0)) + 1) % 10))


def drop_word(xml: str) -> str | None:
    return _edit_text(xml, r"(?<=[a-z] )[a-z]{4,} (?=[a-z])", "")


def renumber(xml: str) -> str | None:
    rx = re.compile(r'(<w:numPr>\s*<w:ilvl w:val=")([0-7])("\s*/>\s*<w:numId w:val="(?!0")\d+"\s*/>)')
    cands = [(p, m) for p in paragraphs(xml) if len(p.text.strip()) >= 20 for m in [rx.search(p.xml)] if m]
    hit = _middle(cands)
    if not hit:
        return None
    p, m = hit
    s = p.start + m.start(2)
    return xml[:s] + str(int(m.group(2)) + 1) + xml[s + 1:]


def nudge_margin(xml: str) -> str | None:
    rx = re.compile(r'(<w:pgMar\b[^>]*\bw:left=")(\d+)(")')
    if not rx.search(xml):
        return None
    return rx.sub(lambda m: f"{m.group(1)}{int(m.group(2)) + 72}{m.group(3)}", xml)


DOCX_BREAKS = {
    "drop-clause": drop_clause,
    "swap-clauses": swap_clauses,
    "change-figure": change_figure,
    "drop-word": drop_word,
    "renumber": renumber,
    "nudge-margin": nudge_margin,
}


def rewrite_docx(src: Path, dst: Path, xml: str) -> None:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = xml.encode("utf-8") if item.filename == DOC_XML else zin.read(item.filename)
            zout.writestr(item, data)


# --- Breaks to the Claude Docs leg and to lawdocs' reading of numbers ---------

def md_drop(md: str) -> str | None:
    paras = md.split("\n\n")
    cands = [i for i, p in enumerate(paras) if len(p) >= 60 and not p.startswith(("|", "#"))]
    i = _middle(cands)
    if i is None:
        return None
    return "\n\n".join(paras[:i] + paras[i + 1:])


def md_figure(md: str) -> str | None:
    paras = md.split("\n\n")
    cands = [i for i, p in enumerate(paras) if len(p) >= 40 and re.search(r"[a-z] \S*\d", p)]
    i = _middle(cands)
    if i is None:
        return None
    paras[i] = re.sub(r"(?<=[a-z] )(\S*?)(\d)", lambda m: m.group(1) + str((int(m.group(2)) + 1) % 10), paras[i], count=1)
    return "\n\n".join(paras)


MD_BREAKS = {"md-drop": md_drop, "md-figure": md_figure}


def misnumber(doc, checked: list[int], differ: set[int]):
    """A copy of ``doc`` with one clause's number read wrong. ``checked`` are the
    clauses (positions in ``doc.clauses()``) that LibreOffice's reading was
    compared with, ``differ`` those whose numbers already disagree; the
    wrong number goes on one that agrees. How many clauses were compared is
    reported separately as coverage."""
    i = _middle([i for i in checked if i not in differ and re.search(r"\d|\([a-z]+\)", doc.clauses()[i].number or "")])
    if i is None:
        return None
    bad = copy.deepcopy(doc)
    target = bad.clauses()[i]
    if re.search(r"\d", target.number):
        target.number = re.sub(r"(\d+)(?!.*\d)", lambda m: str(int(m.group(1)) + 7), target.number)
    else:
        target.number = re.sub(r"\(([a-z]+)\)", lambda m: f"({m.group(1)}z)", target.number, count=1)
    return bad


# --- Running a document ---------------------------------------------------------

def _roundtrip_signals(doc, exported: Path) -> dict:
    again = import_docx(exported)
    return {"refs_match": cc._refs(again) == cc._refs(doc), "text_match": cc._texts(again) == cc._texts(doc)}


def _page_signal(src_pdf: Path | None, exported: Path, work: Path) -> dict | None:
    if src_pdf is None:
        return None
    sub = Path(tempfile.mkdtemp(dir=work))
    pdf = cc._render(exported, sub)
    if pdf is None:
        return None
    cmp_dir = Path(tempfile.mkdtemp(dir=work))
    pages = cc._compare_pdfs(src_pdf, pdf, cmp_dir)
    if "skipped" in pages:
        return None
    return {"pages": (pages["pages_export"], tuple(pages["differing_pages"]))}


def _differs(control: dict, broken: dict) -> list[str]:
    """Checks that agreed for the control and disagree for the broken copy."""
    caught = []
    for k, v in broken.items():
        if k not in control:
            continue
        if k.endswith("_match"):
            if control[k] and not v:
                caught.append(k)
        elif v != control[k]:
            caught.append(k)
    return caught


def selftest(path: Path, out: Path, render: bool) -> dict:
    r: dict = {"file": path.name, "breaks": {}}
    doc = import_docx(path)
    work = Path(tempfile.mkdtemp(prefix="lawdocs-selftest-"))
    try:
        exported = work / "control.docx"
        export_docx(doc, path, exported)
        control = _roundtrip_signals(doc, exported)
        md = md_render.to_markdown(doc)
        pulled = md_import.pull(md, doc)
        control["docs_refs_match"] = cc._refs(pulled) == cc._refs(doc)
        control["docs_text_match"] = cc._texts(pulled) == cc._texts(doc)

        src_pdf, checked, differ = None, [], set()
        if render:
            src_copy = work / ("source-" + path.name)
            shutil.copyfile(path, src_copy)
            src_pdf = cc._render(src_copy, Path(tempfile.mkdtemp(dir=work)))
            control.update(_page_signal(src_pdf, exported, work) or {})
            office = cc._office_labels(path, doc, Path(tempfile.mkdtemp(dir=work)))
            if office and office["compared"]:
                control["office_differ"] = len(office["differ"])
                r["office_coverage"] = f"{office['compared']} of {len(doc.clauses())} clauses"
                checked, differ = office["checked"], set(office["differ_at"])
        r["control"] = {k: list(v) if isinstance(v, tuple) else v for k, v in control.items()}

        xml = zipfile.ZipFile(exported).read(DOC_XML).decode("utf-8")
        for name, brk in DOCX_BREAKS.items():
            broken_xml = brk(xml)
            if broken_xml is None or broken_xml == xml:
                r["breaks"][name] = {"result": "n/a", "why": "nothing to break this way"}
                continue
            broken = work / f"{name}.docx"
            rewrite_docx(exported, broken, broken_xml)
            try:
                sig = _roundtrip_signals(doc, broken)
            except Exception as e:  # noqa: BLE001 - a broken file that will not import is caught
                r["breaks"][name] = {"result": "caught", "by": [f"import error: {e}"]}
                continue
            by = _differs(control, sig)
            # The page comparison is slow, so it runs only when nothing cheaper caught the break.
            if not by and "pages" in control:
                by = _differs(control, _page_signal(src_pdf, broken, work) or {})
            if name == "nudge-margin" and "pages" not in control:
                r["breaks"][name] = {"result": "n/a", "why": "pages not compared"}
                continue
            r["breaks"][name] = {"result": "caught" if by else "MISSED", "by": by}

        for name, brk in MD_BREAKS.items():
            broken_md = brk(md)
            if broken_md is None or broken_md == md:
                r["breaks"][name] = {"result": "n/a", "why": "nothing to break this way"}
                continue
            if not (control["docs_refs_match"] and control["docs_text_match"]):
                r["breaks"][name] = {"result": "n/a", "why": "the control already differs"}
                continue
            p2 = md_import.pull(broken_md, import_docx(path))
            sig = {"docs_refs_match": cc._refs(p2) == cc._refs(doc), "docs_text_match": cc._texts(p2) == cc._texts(doc)}
            by = _differs(control, sig)
            r["breaks"][name] = {"result": "caught" if by else "MISSED", "by": by}

        bad = misnumber(doc, checked, differ) if "office_differ" in control else None
        if "office_differ" not in control:
            r["breaks"]["misnumber"] = {"result": "n/a", "why": "LibreOffice's numbers not compared"}
        elif bad is None:
            r["breaks"]["misnumber"] = {"result": "n/a", "why": "no numbered clause"}
        else:
            office = cc._office_labels(path, bad, Path(tempfile.mkdtemp(dir=work))) or {}
            by = ["office_differ"] if set(office.get("differ_at", [])) - differ else []
            r["breaks"]["misnumber"] = {"result": "caught" if by else "MISSED", "by": by}
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return r


BREAKS = list(DOCX_BREAKS) + list(MD_BREAKS) + ["misnumber"]


def report_md(results: list[dict]) -> str:
    missed = sum(1 for r in results for b in r.get("breaks", {}).values() if b["result"] == "MISSED")
    caught = sum(1 for r in results for b in r.get("breaks", {}).values() if b["result"] == "caught")
    lines = ["# lawdocs corpus self-test", "",
             f"{caught} breaks caught, {missed} missed, across {len(results)} documents.", "",
             "Each column is one deliberate break (see tools/corpus_selftest.py). A cell names the check that caught "
             "it; MISSED means corpus_check would have reported the broken copy as fine.", "",
             "Numbers checked against LibreOffice: how many clauses LibreOffice's text could be matched to. A wrong "
             "number in any other clause would go unnoticed.", "",
             "| File | " + " | ".join(BREAKS) + " | Numbers checked against LibreOffice |",
             "| --- |" + " --- |" * (len(BREAKS) + 1)]
    for r in results:
        if "error" in r:
            lines.append(f"| {r['file']} | error: {r['error']} |" + " |" * len(BREAKS))
            continue
        cells = []
        for name in BREAKS:
            b = r["breaks"].get(name, {"result": "n/a"})
            cells.append(", ".join(x.removesuffix("_match") for x in b.get("by", [])) if b["result"] == "caught"
                         else b["result"])
        lines.append(f"| {r['file']} | " + " | ".join(cells) + f" | {r.get('office_coverage', 'n/a')} |")
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("folder")
    p.add_argument("-o", "--out", default="selftest-report")
    p.add_argument("--no-render", action="store_true", help="skip the LibreOffice page and number checks")
    args = p.parse_args(argv)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    files = sorted(f for f in Path(args.folder).rglob("*.docx") if not f.name.startswith("~$"))
    results = []
    for f in files:
        print(f"breaking {f.name} ...", file=sys.stderr)
        try:
            results.append(selftest(f, out, render=not args.no_render))
        except Exception as e:  # noqa: BLE001 - the report records any failure
            results.append({"file": f.name, "error": str(e)})
        brks = results[-1].get("breaks", {})
        print("  " + ", ".join(f"{k}: {v['result']}" for k, v in brks.items()), file=sys.stderr)
    (out / "selftest.json").write_text(json.dumps(results, indent=2, ensure_ascii=False))
    (out / "selftest.md").write_text(report_md(results))
    print(report_md(results))
    failed = any("error" in r or any(b["result"] == "MISSED" for b in r["breaks"].values()) for r in results)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
