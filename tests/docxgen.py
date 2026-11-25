"""Build small .docx files the way Word writes them, for tests and samples.

Paragraphs are given as dicts:
    {"text": "...", "num": (numId, ilvl), "style": "Heading1", "runs": [("text", "b"), ...]}
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
  <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
  <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="PlainHeading"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>
  <w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>
</w:styles>"""

# numId 1 -> a typical law-firm multilevel list: 1. / 1.1 / (a) / (i) / (A) / (1)
# numId 2 -> a separate (a)/(i) list Word users often start for a sub-list (ilvl 0 = (a))
# numId 3 -> bullets
NUMBERING = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="{W}">
<w:abstractNum w:abstractNumId="10">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:lvlText w:val="%1."/></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:lvlText w:val="%1.%2"/></w:lvl>
  <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/></w:lvl>
  <w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="(%4)"/></w:lvl>
  <w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="(%5)"/></w:lvl>
  <w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="(%6)"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="20">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%1)"/></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="(%2)"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="30">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="20"/></w:num>
<w:num w:numId="3"><w:abstractNumId w:val="30"/></w:num>
<w:num w:numId="4"><w:abstractNumId w:val="20"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
</w:numbering>"""


def _run(text: str, fmt: str = "") -> str:
    fmt = set(fmt.split(",")) if fmt else set()
    rpr = ""
    if "b" in fmt or "i" in fmt:
        rpr = "<w:rPr>" + ("<w:b/>" if "b" in fmt else "") + ("<w:i/>" if "i" in fmt else "") + "</w:rPr>"
    if "del" in fmt:
        return f'<w:del w:id="9" w:author="X"><w:r>{rpr}<w:delText xml:space="preserve">{escape(text)}</w:delText></w:r></w:del>'
    r = f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
    if "ins" in fmt:
        return f'<w:ins w:id="8" w:author="X">{r}</w:ins>'
    return r


def para(p: dict | str) -> str:
    if isinstance(p, str):
        p = {"text": p}
    ppr = ""
    inner = ""
    if p.get("style"):
        inner += f'<w:pStyle w:val="{p["style"]}"/>'
    if p.get("num"):
        nid, il = p["num"]
        inner += f'<w:numPr><w:ilvl w:val="{il}"/><w:numId w:val="{nid}"/></w:numPr>'
    if inner:
        ppr = f"<w:pPr>{inner}</w:pPr>"
    runs = p.get("runs") or [(p.get("text", ""), "")]
    return f"<w:p>{ppr}{''.join(_run(t, f) for t, f in runs)}{p.get('tail', '')}</w:p>"


def table(rows: list[list[str]]) -> str:
    trs = "".join("<w:tr>" + "".join(f"<w:tc>{para(c)}</w:tc>" for c in r) + "</w:tr>" for r in rows)
    return f"<w:tbl>{trs}</w:tbl>"


SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'


def footnote_ref(fid: int) -> str:
    return f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{fid}"/></w:r>'


def write_docx(path: str | Path, body: list[str], footnotes: dict[int, str] | None = None,
               numbering: str | None = None, styles: str | None = None) -> Path:
    """``numbering``/``styles`` replace the default parts (pass whole XML);
    ``styles`` may also be extra <w:style> elements to add to the defaults."""
    path = Path(path)
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W}"><w:body>{"".join(body)}{SECT}</w:body></w:document>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        ct, rels = CONTENT_TYPES, DOC_RELS
        if footnotes:
            ct = ct.replace("</Types>", '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')
            rels = rels.replace("</Relationships>", '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>')
            notes = "".join(
                f'<w:footnote w:id="{i}"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
                f'<w:r><w:t xml:space="preserve"> {escape(t)}</w:t></w:r></w:p></w:footnote>' for i, t in footnotes.items())
            z.writestr("word/footnotes.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="{W}">'
                       '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
                       '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
                       f"{notes}</w:footnotes>")
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/document.xml", doc)
        if styles and not styles.lstrip().startswith("<?xml"):
            styles = STYLES.replace("</w:styles>", styles + "</w:styles>")
        z.writestr("word/styles.xml", styles or STYLES)
        z.writestr("word/numbering.xml", numbering or NUMBERING)
    return path


def sample_agreement() -> list[str]:
    """A short services agreement with the usual Word mess: style-linked headings,
    a restarted sub-list, hand-typed numbers, a tracked change, bullets and a table."""
    return [
        para({"text": "MASTER SERVICES AGREEMENT", "style": "Title"}),
        para('This Master Services Agreement (this "Agreement") is entered into by Acme Corp. ("Customer") and Widget LLC ("Provider").'),
        para({"text": "RECITALS", "style": "PlainHeading"}),
        para("Customer wishes to engage Provider to perform the Services."),
        para({"text": "Definitions", "style": "Heading1"}),
        para({"style": "Heading2", "runs": [("Defined Terms", "b"), (". In this Agreement:", "")]}),
        para({"num": (1, 2), "runs": [('"Affiliate"', "b"), (" means any entity that controls, is controlled by, or is under common control with a party;", "")]}),
        para({"num": (1, 2), "runs": [('"Services"', "b"), (" means the services described in Section 2.1; and", "")]}),
        para({"num": (1, 2), "runs": [('"Fees"', "b"), (" means the amounts set out in Schedule 1.", "")]}),
        para({"text": "Interpretation. Headings are for convenience only.", "style": "Heading2"}),
        para({"text": "Services", "style": "Heading1"}),
        para({"style": "Heading2", "runs": [("Performance", "b"), (". Provider shall perform the Services:", "")]}),
        para({"text": "in a professional manner;", "num": (2, 0)}),
        para({"text": "in accordance with applicable law, including:", "num": (2, 0)}),
        para({"text": "data protection law; and", "num": (2, 1)}),
        para({"text": "export control law; and", "num": (2, 1)}),
        para({"text": "on the timeline agreed by the parties.", "num": (2, 0)}),
        para({"style": "Heading2", "runs": [("Personnel", "b"), (". Provider shall assign qualified personnel. ", ""), ("Provider may not subcontract.", "del"), ("Provider may subcontract with Customer's consent.", "ins")]}),
        para("This sentence continues Section 2.2 as an unnumbered paragraph."),
        para({"text": "Fees", "style": "Heading1"}),
        para("3.1 Payment. Customer shall pay the Fees within 30 days of invoice."),
        para("3.2 Late Payment. Late amounts bear interest at 1% per month, provided that:"),
        para("(a) interest shall not exceed the maximum rate permitted by law; and"),
        para("(b) no interest accrues on amounts disputed under Section 3.3."),
        para({"text": "Invoices are delivered:", "num": None}),
        para({"text": "by email; or", "num": (3, 0)}),
        para({"text": "by courier.", "num": (3, 0)}),
        para({"text": "Term", "style": "Heading1"}),
        para({"text": "This Agreement continues until terminated under Section 4.2.", "style": "Heading2"}),
        para({"text": "Either party may terminate on 30 days' notice, in which case:", "style": "Heading2"}),
        para({"text": "Customer pays for Services performed;", "num": (4, 0)}),
        para({"text": "Provider returns Customer materials.", "num": (4, 0)}),
        para({"text": "SCHEDULE 1", "style": "PlainHeading"}),
        table([["Service", "Fee (USD)"], ["Implementation", "50,000"], ["Support (per month)", "5,000"]]),
    ]
