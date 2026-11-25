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


def test_unedited_doc_pulls_back_to_a_verbatim_export(tmp_path):
    src = write_docx(tmp_path / "a.docx", sample_agreement())
    base = import_docx(src)
    new = md_import.pull(render.to_markdown(base), base)
    assert [b.level for b in new.blocks] == [b.level for b in base.blocks]
    direct: dict = {}
    export_docx(base, src, tmp_path / "direct.docx", direct)
    pulled: dict = {}
    export_docx(new, src, tmp_path / "pulled.docx", pulled)
    # A trip through Claude Docs with no edits rebuilds nothing extra.
    assert pulled == direct and direct["verbatim"] > 0


def test_repeated_footnote_text_keeps_each_notes_word_data():
    from lawdocs.model import Document
    base = Document(blocks=[], notes=[{"n": 1, "text": "Same.", "word": '<w:footnote w:id="7"/>'},
                                      {"n": 2, "text": "Same.", "word": '<w:footnote w:id="8"/>'}])
    new = md_import.pull("## Footnotes\n\n¹ Same.\n\n² Same.\n", base)
    assert [n.get("word") for n in new.notes] == ['<w:footnote w:id="7"/>', '<w:footnote w:id="8"/>']
