from docxgen import sample_agreement, write_docx
from lawdocs import md_import, outline, render
from lawdocs.docx_export import export_docx
from lawdocs.docx_import import import_docx


def refs(doc):
    return [e["ref"] for e in outline.build(doc) if e["ref"]]


def test_parse_inline_escapes_and_emphasis():
    runs = md_import.parse_inline(r"the “**Affiliate**” and \[\_\_\] *specify* \$5")
    assert "".join(r.text for r in runs) == "the “Affiliate” and [__] specify $5"
    assert [r.text for r in runs if r.bold] == ["Affiliate"]
    assert [r.text for r in runs if r.italic] == ["specify"]


def test_pull_of_unchanged_render_is_identity(tmp_path):
    base = import_docx(write_docx(tmp_path / "a.docx", sample_agreement()))
    new = md_import.pull(render.to_markdown(base), base)
    assert [md_import._key(b) for b in new.blocks] == [md_import._key(b) for b in base.blocks]
    assert all(b.word for b in new.blocks)
    assert [b.id for b in new.blocks] == [b.id for b in base.blocks]


def test_edit_in_docs_then_export_renumbers(tmp_path):
    src = write_docx(tmp_path / "a.docx", sample_agreement())
    base = import_docx(src)
    md = render.to_markdown(base)
    # As if edited in Claude Docs: delete 1.2, add a clause typed with a stale number.
    md = md.replace("1.2 **Interpretation.** Headings are for convenience only.\n\n", "")
    md = md.replace("(c) on the timeline agreed by the parties.",
                    "(c) on the timeline agreed by the parties.\n\n(c) in a secure manner.")
    new = md_import.pull(md, base, editor="claude")
    added = next(b for b in new.blocks if b.text == "in a secure manner.")
    assert added.author == "claude" and added.word is None and added.number == "(d)"
    again = import_docx(export_docx(new, src, tmp_path / "out.docx"))
    assert "1.2" not in refs(again)
    assert "2.1(d)" in refs(again)
