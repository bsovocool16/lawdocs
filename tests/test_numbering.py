import pytest

from lawdocs import numbering as num
from lawdocs.model import Block, Document, Run


@pytest.mark.parametrize("n,roman", [(1, "i"), (4, "iv"), (9, "ix"), (14, "xiv"), (40, "xl"), (1994, "mcmxciv")])
def test_roman_round_trip(n, roman):
    assert num.to_roman(n) == roman
    assert num.from_roman(roman) == n


@pytest.mark.parametrize("n,alpha", [(1, "a"), (26, "z"), (27, "aa"), (28, "bb"), (53, "aaa")])
def test_alpha_round_trip(n, alpha):
    assert num.to_alpha(n) == alpha
    assert num.from_alpha(alpha) == n


def _doc(levels):
    return Document(blocks=[Block(id=f"b{i}", kind="clause", level=l, runs=[Run("x")]) for i, l in enumerate(levels)])


def test_legal_scheme_labels_and_references():
    doc = _doc([1, 2, 3, 4, 4, 5, 6, 3, 2, 1, 2])
    refs = num.apply(doc)
    assert [b.number for b in doc.blocks] == ["1.", "1.1", "(a)", "(i)", "(ii)", "(A)", "(1)", "(b)", "1.2", "2.", "2.1"]
    assert refs["b4"] == "1.1(a)(ii)"
    assert refs["b6"] == "1.1(a)(ii)(A)(1)"
    assert refs["b7"] == "1.1(b)"
    assert refs["b10"] == "2.1"


def test_renumbering_is_derived_from_structure():
    doc = _doc([1, 2, 2, 2])
    num.apply(doc)
    del doc.blocks[1]  # delete 1.1; the rest must move up
    refs = num.apply(doc)
    assert list(refs.values()) == ["1", "1.1", "1.2"]


def test_skipped_level_reference_omits_zero_counter():
    doc = _doc([1, 3, 3])
    refs = num.apply(doc)
    assert refs["b1"] == "1(a)"
    assert refs["b2"] == "1(b)"


@pytest.mark.parametrize("label,prev,level", [
    ("1.", None, 1), ("Article 4", None, 1), ("Section 2", None, 1), ("2.3", None, 2), ("2.3.1", None, 3),
    ("(a)", None, 3), ("(i)", None, 4), ("(i)", "h", 3), ("(ii)", "h", 4), ("(v)", "u", 3), ("(v)", "b", 4),
    ("(A)", None, 5), ("(1)", None, 6), ("a.", None, 3), ("iv.", None, 4),
])
def test_classify_label(label, prev, level):
    assert num.classify_label(label, prev_alpha_lower=prev) == level


@pytest.mark.parametrize("text,expected", [
    ("1.1 Payment. Customer pays.", ("1.1", "Payment. Customer pays.")),
    ("(a) the first thing", ("(a)", "the first thing")),
    ("Section 3. Fees", ("Section 3.", "Fees")),
    ("ARTICLE IV DEFINITIONS", ("ARTICLE IV", "DEFINITIONS")),
    ("1999 was a year", None),
    ("A. Smith signed", None),
])
def test_split_typed_number(text, expected):
    assert num.split_typed_number(text) == expected
