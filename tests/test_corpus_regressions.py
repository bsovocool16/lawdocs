"""One test per problem found by running tools/corpus_check.py over public
model contracts (NVCA, gov.uk Model Services Contract, Lambert toolkit,
Charity Commission models, UKSA grant terms, oneNDA)."""

import zipfile
import xml.etree.ElementTree as ET

from docxgen import W, para, sample_agreement, write_docx
from lawdocs import md_import, numbering as num, outline, render
from lawdocs.docx_export import export_docx
from lawdocs.docx_import import import_docx
from lawdocs.model import Block, Run


def refs(doc):
    return [e["ref"] for e in outline.build(doc) if e["ref"]]


def numbers(doc):
    return [b.number for b in doc.clauses()]


def _numbering(*abstracts: str, nums: str) -> str:
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="{W}">'
            + "".join(abstracts) + nums + "</w:numbering>")


def _lvl(ilvl: int, fmt: str, text: str) -> str:
    return f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/></w:lvl>'


# A picture declares its own namespaces (a:, pic:) on the drawing, not on the
# document root. They were stripped from the saved paragraph, so the exported
# document.xml was not well-formed (Charity Commission models, MSC).
def test_namespaces_declared_inside_a_paragraph_survive_export(tmp_path):
    drawing = ('<w:r><w:drawing><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
               '<a:graphicData uri="x"/></a:graphic></w:drawing></w:r>')
    src = write_docx(tmp_path / "pic.docx", [para({"text": "Scope", "style": "Heading1"}),
                                              para({"text": "Logo: ", "tail": drawing})])
    out = export_docx(import_docx(src), src, tmp_path / "out.docx")
    xml = zipfile.ZipFile(out).read("word/document.xml")
    ET.fromstring(xml)  # well-formed
    assert b'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' in xml


# "(viii)" has four letters and fell outside the label pattern (NVCA charter).
def test_long_roman_numerals_are_labels():
    assert num.classify_label("(viii)") == 4
    assert num.classify_label("(xxviii)") == 4
    assert num.split_typed_number("(viii) shares issued") == ("(viii)", "shares issued")


# Citations and notes that start like an article heading (NVCA charter notes,
# Charity Commission guidance notes).
def test_citations_and_notes_are_not_headings():
    assert num.split_typed_number("Section 2115 of the California Corporations Code provides") is None
    assert num.split_typed_number("Article 5(1) - This provides a general power") is None
    assert num.split_typed_number("Article 2 - The articles include") is None
    assert num.split_typed_number("Section 3. Fees") == ("Section 3.", "Fees")


# A contents line has its number typed in; drafters also use TOC styles for
# body text, which must still be read (MSC, Lambert heads of terms).
def test_contents_entries_are_text_but_toc_styled_clauses_are_not(tmp_path):
    toc = '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style>'
    src = write_docx(tmp_path / "toc.docx", [
        para({"text": "Section C: Payment Provisions\t12", "style": "TOC1"}),
        para({"text": "Fees", "style": "Heading1"}),
        para({"text": "1.1 The Institution will pay:", "style": "TOC1"}),
    ], styles=toc)
    doc = import_docx(src)
    assert doc.blocks[0].kind == "text"
    assert refs(doc) == ["1", "1.1"]


# Each document keeps its own number forms: "1.1.1" stays "1.1.1", and an
# "(a)" under it is a level below it, not beside it (Lambert, MSC, COI).
def test_documents_keep_their_own_number_forms(tmp_path):
    numbering = _numbering(
        '<w:abstractNum w:abstractNumId="1">' + _lvl(0, "decimal", "%1.") + _lvl(1, "decimal", "%1.%2")
        + _lvl(2, "decimal", "%1.%2.%3") + _lvl(3, "lowerLetter", "(%4)") + "</w:abstractNum>",
        nums='<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>')
    src = write_docx(tmp_path / "uk.docx", [
        para({"text": "Definitions", "num": (1, 0)}),
        para({"text": "In this Agreement:", "num": (1, 1)}),
        para({"text": "the first rule applies;", "num": (1, 2)}),
        para({"text": "including this;", "num": (1, 3)}),
        para({"text": "and this;", "num": (1, 3)}),
        para({"text": "the second rule applies.", "num": (1, 2)}),
    ], numbering=numbering)
    doc = import_docx(src)
    assert numbers(doc) == ["1.", "1.1", "1.1.1", "(a)", "(b)", "1.1.2"]
    assert [b.level for b in doc.clauses()] == [1, 2, 3, 4, 4, 3]
    assert refs(doc)[3] == "1.1.1(a)"
    # A clause added under 1.1 is written the same way.
    doc.blocks.insert(len(doc.blocks), Block("new", "clause", 3, [Run("a third rule.")]))
    num.apply(doc)
    assert doc.blocks[-1].number == "1.1.3"


# Articles numbered "1" with sub-clauses "(1)" then "(a)" (Charity Commission
# articles): "(1)" is the sub-clause of "1", and "(a)" sits under "(1)".
def test_paren_number_under_a_plain_number_nests(tmp_path):
    numbering = _numbering(
        '<w:abstractNum w:abstractNumId="1">' + _lvl(0, "decimal", "%1") + "</w:abstractNum>",
        '<w:abstractNum w:abstractNumId="2">' + _lvl(0, "decimal", "(%1)") + _lvl(1, "lowerLetter", "(%2)") + "</w:abstractNum>",
        nums='<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>')
    src = write_docx(tmp_path / "art.docx", [
        para({"text": "The charity's powers are:", "num": (1, 0)}),
        para({"text": "to raise funds, including:", "num": (2, 0)}),
        para({"text": "by appeals;", "num": (2, 1)}),
        para({"text": "to buy property.", "num": (2, 0)}),
        para({"text": "The trustees may delegate.", "num": (1, 0)}),
    ], numbering=numbering)
    doc = import_docx(src)
    assert numbers(doc) == ["1", "(1)", "(a)", "(2)", "2"]
    assert refs(doc) == ["1", "1(1)", "1(1)(a)", "1(2)", "2"]


# A list that starts again ("1." after a heading, in a schedule or a second
# part) keeps the numbers Word shows (NVCA charter, Lambert schedules).
def test_restarted_lists_keep_their_numbers(tmp_path):
    src = write_docx(tmp_path / "restart.docx", [
        para({"text": "Scope", "style": "Heading1"}),
        para({"text": "Fees", "style": "Heading1"}),
        para({"text": "SCHEDULE 1", "style": "PlainHeading"}),
        para("1. Services. The services are consulting."),
        para("2. Hours. Up to ten hours a week."),
    ])
    doc = import_docx(src)
    assert numbers(doc) == ["1.", "2.", "1.", "2."]
    assert doc.clauses()[2].restart
    # The restart is structural: deleting the clause before it changes nothing after it.
    doc.blocks = [b for b in doc.blocks if b.text != "Fees" and b.caption != "Fees"]
    num.apply(doc)
    assert numbers(doc) == ["1.", "1.", "2."]


# Word list instances with start overrides on several levels show the start
# value of a level they haven't used yet ("2.12"), and later levels restart
# normally (NVCA IRA: section 2.12 was read as 1.12, and 3.1 as 3.13).
def test_start_overrides_on_several_levels(tmp_path):
    numbering = _numbering(
        '<w:abstractNum w:abstractNumId="1">' + _lvl(0, "decimal", "%1.") + _lvl(1, "decimal", "%1.%2") + "</w:abstractNum>",
        nums='<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>'
             '<w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="2"/>'
             '</w:lvlOverride><w:lvlOverride w:ilvl="1"><w:startOverride w:val="12"/></w:lvlOverride></w:num>')
    src = write_docx(tmp_path / "ov.docx", [
        para({"text": "Registration Rights", "num": (1, 0)}),
        para({"text": "Demand registration is available.", "num": (1, 1)}),
        para({"text": "Restrictions on Transfer apply.", "num": (2, 1)}),
        para({"text": "Information Rights", "num": (1, 0)}),
        para({"text": "Financial statements are delivered.", "num": (1, 1)}),
    ], numbering=numbering)
    doc = import_docx(src)
    assert [b.source_number for b in doc.clauses()] == ["1.", "1.1", "2.12", "3.", "3.1"]


# Model forms offer alternative versions of a clause with the same number
# (NVCA charter: two 4.4.4s); the alternative must not shift what follows.
def test_alternative_clauses_share_a_number(tmp_path):
    src = write_docx(tmp_path / "alt.docx", [
        para({"text": "Conversion", "style": "Heading1"}),
        para("1.1 Adjustment. The price adjusts on a broad-based basis."),
        para("[Use the following Section 1.1 if the term sheet calls for full ratchet.]"),
        para("1.1 Adjustment. The price adjusts to the new price."),
        para("1.2 Determination. Consideration is determined as follows."),
    ])
    doc = import_docx(src)
    assert numbers(doc) == ["1.", "1.1", "1.1", "1.2"]
    assert doc.clauses()[2].alternative


# A clause added beside existing ones joins their Word list, so Word counts it
# with them (it used to take the most common list for its level, and Word
# showed "(a)" where lawdocs showed "(d)").
def test_added_clause_joins_its_siblings_word_list(tmp_path):
    src = write_docx(tmp_path / "msa.docx", sample_agreement())
    doc = import_docx(src)
    c = next(i for i, b in enumerate(doc.blocks) if b.text.startswith("on the timeline"))
    doc.blocks.insert(c + 1, Block("new", "clause", 3, [Run("in a secure manner.")]))
    out = export_docx(doc, src, tmp_path / "out.docx")
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    added = xml[xml.rfind("<w:p>", 0, xml.find("in a secure manner")):xml.find("in a secure manner")]
    assert '<w:numId w:val="2" />' in added or '<w:numId w:val="2"/>' in added
    again = import_docx(out)
    assert next(b for b in again.clauses() if b.text == "in a secure manner.").source_number == "(d)"


# Reading a doc back from Claude Docs must not change its structure.
def test_docs_read_back_keeps_structure(tmp_path):
    numbering = _numbering(
        '<w:abstractNum w:abstractNumId="1">' + _lvl(0, "decimal", "%1.") + "</w:abstractNum>",
        '<w:abstractNum w:abstractNumId="2">' + _lvl(0, "upperLetter", "(%1)") + "</w:abstractNum>",
        '<w:abstractNum w:abstractNumId="3">' + _lvl(0, "decimal", "%1.") + _lvl(1, "lowerLetter", "%2.") + "</w:abstractNum>",
        nums='<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>'
             '<w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>')
    src = write_docx(tmp_path / "rb.docx", [
        para({"text": "BACKGROUND", "style": "PlainHeading"}),
        para({"text": "The parties wish to collaborate.", "num": (2, 0)}),   # recital (A)
        para({"text": "(2) [INSERT NAME]", "style": "PlainHeading"}),       # a heading that starts like a number
        para({"text": "Voting", "num": (1, 0)}),                            # caption-only clause ...
        para("Each share has one vote."),                                     # ... followed by a paragraph
        para({"text": "Powers", "num": (3, 0)}),
        para({"text": "to raise funds;", "num": (3, 1)}),                    # "a." style
    ], numbering=numbering)
    doc = import_docx(src)
    pulled = md_import.pull(render.to_markdown(doc), doc)
    shape = lambda d: [(b.kind, b.level, b.number, b.caption, b.text) for b in d.blocks]  # noqa: E731
    assert shape(pulled) == shape(doc)


# A Word list that isn't restarted runs on under the next parent: "(5)" under
# article 6 after "(1)"-"(4)" under article 5 (Charity Commission models).
def test_list_running_on_under_the_next_parent(tmp_path):
    numbering = _numbering(
        '<w:abstractNum w:abstractNumId="1">' + _lvl(0, "decimal", "%1") + "</w:abstractNum>",
        '<w:abstractNum w:abstractNumId="2">' + _lvl(0, "decimal", "(%1)") + "</w:abstractNum>",
        nums='<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>')
    src = write_docx(tmp_path / "run.docx", [
        para({"text": "The charity may:", "num": (1, 0)}),
        para({"text": "raise funds;", "num": (2, 0)}),
        para({"text": "buy property;", "num": (2, 0)}),
        para({"text": "The charity may also:", "num": (1, 0)}),
        para({"text": "employ staff.", "num": (2, 0)}),
    ], numbering=numbering)
    doc = import_docx(src)
    assert numbers(doc) == ["1", "(1)", "(2)", "2", "(3)"]
    assert doc.clauses()[4].continues


# "5A" inserted after "5" repeats its counter instead of taking the next one
# (NVCA charter's Section 5A).
def test_inserted_clause_with_a_letter_suffix(tmp_path):
    numbering = _numbering(
        '<w:abstractNum w:abstractNumId="1">' + _lvl(0, "decimal", "%1.") + "</w:abstractNum>",
        '<w:abstractNum w:abstractNumId="2">' + _lvl(0, "decimal", "%1A.") + "</w:abstractNum>",
        nums='<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>')
    src = write_docx(tmp_path / "5a.docx", [
        para({"text": "Voting. Each share votes.", "num": (1, 0)}),
        para({"text": "Protective Provisions. Consent is needed.", "num": (2, 0)}),
        para({"text": "Redemption. Shares are not redeemable.", "num": (1, 0)}),
    ], numbering=numbering)
    doc = import_docx(src)
    assert numbers(doc) == ["1.", "1A.", "2."]
    assert doc.clauses()[1].alternative


# Charter articles numbered in words ("FIRST:"), Word's ordinalText format.
def test_ordinal_text_numbers():
    assert [num.format_counter(n, "ordinalText") for n in (1, 2, 3, 4, 5, 12, 20, 21)] == [
        "First", "Second", "Third", "Fourth", "Fifth", "Twelfth", "Twentieth", "Twenty-First"]
    style = num.label_style("Fifth:", 1, "ordinalText")
    assert num.render_style(style, [6] + [0] * 8, 1) == "Sixth:"
    assert num.counter_value("Fifth:", "ordinalText") == 5


# An empty table (a box to fill in on a questionnaire) has nothing to show in
# Claude Docs but must go back to Word (NVCA PIPE registration rights
# agreement: the investor questionnaire's answer boxes were dropped).
def test_empty_fill_in_tables_survive_export(tmp_path):
    box = "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"
    src = write_docx(tmp_path / "q.docx", [
        para({"text": "Questionnaire", "style": "Heading1"}),
        para("1.1 Name. Full legal name of the investor:"),
        box,
        para("1.2 Address. Address for notices:"),
    ])
    doc = import_docx(src)
    assert not any(b.kind == "table" for b in doc.blocks)
    out = export_docx(doc, src, tmp_path / "out.docx")
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert xml.count("<w:tbl>") == 1 and xml.index("<w:tbl>") < xml.index("Address for notices")
