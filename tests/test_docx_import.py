import shutil
import subprocess

import pytest

from docxgen import para, sample_agreement, table, write_docx
from lawdocs import outline, render
from lawdocs.docx_import import import_docx


@pytest.fixture
def sample(tmp_path):
    return import_docx(write_docx(tmp_path / "msa.docx", sample_agreement()))


def _numbered(doc):
    return [(b.number, b.caption, b.text) for b in doc.clauses()]


def test_word_multilevel_list_becomes_legal_numbering(sample):
    refs = [e["ref"] for e in outline.build(sample) if e["ref"]]
    assert refs == [
        "1", "1.1", "1.1(a)", "1.1(b)", "1.1(c)", "1.2",
        "2", "2.1", "2.1(a)", "2.1(b)", "2.1(b)(i)", "2.1(b)(ii)", "2.1(c)", "2.2",
        "3", "3.1", "3.2", "3.2(a)", "3.2(b)",
        "4", "4.1", "4.2", "4.2(a)", "4.2(b)",
    ]


def test_title_headings_and_body_text(sample):
    kinds = [(b.kind, b.text[:20]) for b in sample.blocks if b.kind != "clause"]
    assert kinds[0] == ("title", "MASTER SERVICES AGRE")
    assert ("heading", "RECITALS") in kinds
    assert ("heading", "SCHEDULE 1") in kinds
    cont = next(b for b in sample.blocks if b.text.startswith("This sentence continues"))
    assert cont.kind == "text" and cont.level == 2


def test_captions_from_bold_lead_ins_and_typed_title_case(sample):
    caps = {b.number: b.caption for b in sample.clauses() if b.caption}
    assert caps["1."] == "Definitions"
    assert caps["1.1"] == "Defined Terms"
    assert caps["3.2"] == "Late Payment"  # typed, not bold
    by_text = {b.text: b for b in sample.clauses()}
    assert "This Agreement continues until terminated under Section 4.2." in by_text


def test_tracked_changes_accepted_and_reported(sample):
    personnel = next(b for b in sample.clauses() if b.caption == "Personnel")
    assert "may subcontract with Customer's consent" in personnel.text
    assert "may not subcontract" not in personnel.text
    assert any("tracked" in w for w in sample.warnings)


def test_typed_numbers_become_structure(sample):
    fees = [b for b in sample.clauses() if b.source_number in ("3.1", "3.2", "(a)", "(b)") and b.number in ("3.1", "3.2")]
    assert len(fees) == 2
    assert any("hand-typed" in w for w in sample.warnings)


def test_list_restart_override(sample):
    term = [b.number for b in sample.clauses()][-2:]
    assert term == ["(a)", "(b)"]


def test_bold_defined_terms_survive(sample):
    affiliate = next(b for b in sample.clauses() if "Affiliate" in b.text)
    assert affiliate.runs[0].bold and affiliate.runs[0].text == '"Affiliate"'


def test_table(sample):
    t = next(b for b in sample.blocks if b.kind == "table")
    assert t.rows[0] == ["Service", "Fee (USD)"]
    md = render.to_markdown(sample)
    assert "| Implementation | 50,000 |" in md


def test_markdown_never_turns_numbers_into_lists(tmp_path):
    doc = import_docx(write_docx(tmp_path / "d.docx", [
        para({"text": "Scope", "style": "Heading1"}),
        para("1. something typed that looks like a list."),
        para("- a dash at the start"),
    ]))
    md = render.to_markdown(doc)
    assert "## 1. Scope" in md
    assert "\n2\\. something typed" in md
    assert "\n\\- a dash" in md


def test_block_map_counts_paragraphs(sample):
    m = render.block_map(sample)
    total = sum(len(e["paragraphs"]) for e in m)
    assert total == render.to_markdown(sample).count("\n\n") + 1


def test_empty_paragraphs_dropped(tmp_path):
    doc = import_docx(write_docx(tmp_path / "e.docx", [para(""), para("Hello"), para("   ")]))
    assert [b.text for b in doc.blocks] == ["Hello"]


@pytest.mark.skipif(shutil.which("pandoc") is None, reason="pandoc not installed")
def test_pandoc_generated_docx(tmp_path):
    src = tmp_path / "t.md"
    src.write_text("1. First clause text here.\n    a. alpha\n    b. beta\n        i. roman\n2. Second.\n")
    out = tmp_path / "t.docx"
    subprocess.run(["pandoc", str(src), "-o", str(out)], check=True)
    refs = [e["ref"] for e in outline.build(import_docx(out)) if e["ref"]]
    assert refs == ["1", "1(a)", "1(b)", "1(b)(i)", "2"]
