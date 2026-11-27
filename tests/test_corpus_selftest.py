"""tools/corpus_selftest.py: every deliberate break of the sample agreement is
caught by the corpus check's text and structure comparisons (no LibreOffice)."""

import sys
from pathlib import Path

from docxgen import para, sample_agreement, write_docx

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import corpus_selftest as st  # noqa: E402


def test_text_and_structure_breaks_are_caught(tmp_path):
    src = write_docx(tmp_path / "msa.docx", sample_agreement())
    r = st.selftest(src, tmp_path, render=False)
    assert r["control"]["text_match"] and r["control"]["refs_match"]
    for name in ("drop-clause", "swap-clauses", "change-figure", "drop-word", "renumber", "md-drop", "md-figure"):
        assert r["breaks"][name]["result"] == "caught", (name, r["breaks"][name])
    # These need LibreOffice.
    assert r["breaks"]["nudge-margin"]["result"] == "n/a"
    assert r["breaks"]["misnumber"]["result"] == "n/a"


def test_paragraphs_keep_text_box_paragraphs_inside_their_own():
    xml = ('<w:body><w:p><w:r><w:t>one</w:t></w:r></w:p>'
           '<w:p><w:r><w:txbxContent><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:txbxContent></w:r></w:p>'
           '<w:p/></w:body>')
    ps = st.paragraphs(xml)
    assert [p.text for p in ps] == ["one", "inner", ""]
    assert len(ps) == 3


def test_a_document_with_nothing_to_break_reports_n_a(tmp_path):
    src = write_docx(tmp_path / "short.docx", [para("Short.")])
    r = st.selftest(src, tmp_path, render=False)
    assert r["breaks"]["drop-clause"]["result"] == "n/a"
    assert r["breaks"]["swap-clauses"]["result"] == "n/a"


# The first corpus run found that a changed table cell went unnoticed: a
# table's own text is empty, and the comparison did not look at its rows
# (oneNDA keeps all its terms in tables).
def test_a_changed_table_cell_is_caught(tmp_path):
    import zipfile

    from docxgen import table

    src = write_docx(tmp_path / "t.docx", [para("Fees are set out below and are payable monthly in arrears."),
                                           table([["Service", "Fee"], ["Support", "5,000"]])])
    doc = st.import_docx(src)
    out = st.export_docx(doc, src, tmp_path / "out.docx")
    xml = zipfile.ZipFile(out).read(st.DOC_XML).decode("utf-8")
    broken = tmp_path / "broken.docx"
    st.rewrite_docx(out, broken, xml.replace("5,000", "6,000"))
    assert st._roundtrip_signals(doc, out)["text_match"]
    assert not st._roundtrip_signals(doc, broken)["text_match"]
