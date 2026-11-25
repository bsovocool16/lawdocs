import re
import zipfile

import pytest

from docxgen import footnote_ref, para, sample_agreement, write_docx
from lawdocs import outline
from lawdocs.docx_export import export_docx
from lawdocs.docx_import import import_docx
from lawdocs.model import Block, Run


def refs(doc):
    return [e["ref"] for e in outline.build(doc) if e["ref"]]


def texts(doc):
    return [(b.kind, b.caption, b.text) for b in doc.blocks]


@pytest.fixture
def src(tmp_path):
    return write_docx(tmp_path / "msa.docx", sample_agreement())


def test_round_trip_keeps_structure_and_text(src, tmp_path):
    doc = import_docx(src)
    out = export_docx(doc, src, tmp_path / "out.docx")
    again = import_docx(out)
    assert refs(again) == refs(doc)
    assert texts(again) == texts(doc)


def test_typed_numbers_stay_typed_and_follow_edits(src, tmp_path):
    doc = import_docx(src)
    out = export_docx(doc, src, tmp_path / "out.docx")
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    # Section 3 was typed by hand in the source: untouched, it goes back as it was.
    assert ">3.1 Payment. Customer shall pay" in xml and ">(a) interest" in xml
    assert 'w:numId w:val="990"' not in xml
    # Delete 3.1: the typed numbers after it are rewritten, still as text.
    doc.blocks = [b for b in doc.blocks if not b.text.startswith("Customer shall pay")]
    out = export_docx(doc, src, tmp_path / "edited.docx")
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert ">3.2 " not in xml and ">3.1 " in xml
    again = import_docx(out)
    late = next(b for b in again.clauses() if b.text.startswith("Late amounts"))
    assert late.number == "3.1" and late.source_number == "3.1"
    assert [c.number for c in again.clauses() if c.text.startswith(("interest shall", "no interest"))] == ["(a)", "(b)"]


def test_edits_renumber_and_take_house_style(src, tmp_path):
    doc = import_docx(src)
    # Delete 2.1(b)(i) and add a new (d) under 2.1, written as Claude would: no Word data.
    doc.blocks = [b for b in doc.blocks if not b.text.startswith("data protection")]
    c = next(i for i, b in enumerate(doc.blocks) if b.text.startswith("on the timeline"))
    doc.blocks.insert(c + 1, Block(id="new1", kind="clause", level=3, runs=[Run("in a secure manner.")]))
    out = export_docx(doc, src, tmp_path / "out.docx")
    again = import_docx(out)
    assert "2.1(b)(i)" in refs(again) and "2.1(b)(ii)" not in refs(again)
    new = next(b for b in again.clauses() if b.text == "in a secure manner.")
    assert new.number == "(d)"


def test_footnotes_round_trip(tmp_path):
    src = write_docx(tmp_path / "fn.docx", [
        para({"text": "Scope", "style": "Heading1"}),
        para({"style": "Heading2", "runs": [("Term", "b"), (". The term is one year.", "")], "tail": footnote_ref(5)}),
    ], footnotes={5: "Adjust as needed."})
    doc = import_docx(src)
    assert doc.notes[0]["text"] == "Adjust as needed."
    out = export_docx(doc, src, tmp_path / "out.docx")
    again = import_docx(out)
    assert [n["text"] for n in again.notes] == ["Adjust as needed."]
    assert again.clauses()[1].text.endswith("\u00b9")


def test_spacers_and_section_properties_survive(src, tmp_path):
    doc = import_docx(src)
    out = export_docx(doc, src, tmp_path / "out.docx")
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert re.search(r"<w:sectPr>.*</w:sectPr></w:body>", xml)


FIELD = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref1 \\r \\h </w:instrText></w:r>'
         '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>1.1</w:t></w:r>'
         '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def _fielded(tmp_path):
    return write_docx(tmp_path / "f.docx", [
        para({"text": "Scope", "style": "Heading1"}),
        para({"style": "Heading2", "runs": [("Term. See Section ", "")],
              "tail": FIELD + '<w:r><w:rPr><w:smallCaps/></w:rPr><w:t xml:space="preserve"> below</w:t></w:r>' + footnote_ref(5)}),
        para({"style": "Heading2", "runs": [("Fees. Paid monthly.", "")]}),
    ], footnotes={5: "Adjust as needed."})


def test_untouched_paragraphs_keep_their_word_content(tmp_path):
    src = _fielded(tmp_path)
    doc = import_docx(src)
    stats: dict = {}
    out = export_docx(doc, src, tmp_path / "out.docx", stats)
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert "REF _Ref1" in xml and "<w:smallCaps" in xml   # field and run formatting survive
    assert stats["rebuilt"] == 0
    fid = re.search(r'footnoteReference w:id="(\d+)"', xml).group(1)
    notes = zipfile.ZipFile(out).read("word/footnotes.xml").decode()
    assert f'w:id="{fid}"' in notes  # reference points at the rewritten footnote


def test_edited_paragraph_is_rebuilt(tmp_path):
    src = _fielded(tmp_path)
    doc = import_docx(src)
    fees = next(b for b in doc.blocks if (b.caption or "").startswith("Fees") or b.text.startswith("Fees"))
    fees.runs = [Run("Paid quarterly.")]
    stats: dict = {}
    out = export_docx(doc, src, tmp_path / "out.docx", stats)
    assert stats["rebuilt"] == 1
    again = import_docx(out)
    assert any("Paid quarterly." in b.text for b in again.blocks)
    assert "REF _Ref1" in zipfile.ZipFile(out).read("word/document.xml").decode()
