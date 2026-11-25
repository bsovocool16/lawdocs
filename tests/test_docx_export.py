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


def test_typed_numbers_become_word_numbering(src, tmp_path):
    doc = import_docx(src)
    out = export_docx(doc, src, tmp_path / "out.docx")
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    # Section 3 was typed by hand in the source; on export Word numbers it.
    assert ">3.1 " not in xml and ">(a) interest" not in xml
    assert 'w:numId w:val="990"' in xml
    assert "abstractNumId=\"990\"" in zipfile.ZipFile(out).read("word/numbering.xml").decode()
    again = import_docx(out)
    assert not any("hand-typed" in w for w in again.warnings)


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
