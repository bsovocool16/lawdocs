# lawdocs

Legal documents with block-level structure instead of Word's text-level
formatting, using Claude Docs as the editor.

A clause is a block with a level. Its number (`1.1(a)(i)`), its heading
treatment and its place in the outline are computed from that level every
time, so numbering can't drift, and there is no direct formatting to fight.

## What works today

- **.docx import.** Reads Word's styles and multilevel lists (including
  style-linked headings, list restarts and hand-typed numbers), accepts
  tracked changes, keeps bold and italic, and drops everything else.
- **Legal numbering.** 1. / 1.1 / (a) / (i) / (A) / (1), recomputed from
  structure. Full references like `2.1(b)(ii)` for every clause.
- **Outline.** Every clause's reference, caption and opening words, with
  block ids, so Claude can jump to a clause instead of grepping.
- **Claude Docs output.** Markdown that writes cleanly into a Claude Doc,
  plus a map from each block to the doc paragraphs it produced. Footnotes,
  which Claude Docs lacks, become superscript markers and a Footnotes section.
- **Round trip back to Word.** `lawdocs pull` reads the edited Claude Doc
  back; `lawdocs export` writes a .docx using the source document as its
  template (LaTeX-style: blocks are content, the source is the document
  class). Untouched paragraphs get their original Word properties back;
  new or moved ones take the house style for their level; numbering is
  Word's own auto-numbering; footnotes, section breaks, headers and footers
  survive. Tested on the NVCA model SPA (210 clauses, 108 footnotes).

## Use

```
pip install -e .
lawdocs import contract.docx -o out/contract
lawdocs outline out/contract/model.json --depth 2

# after editing in Claude Docs: export the doc's Markdown, then
lawdocs pull edited.md --base out/contract/model.json
lawdocs export out/contract/model.json -o contract-edited.docx
```

`out/contract/doc.md` is what goes into Claude Docs. In Claude Code, the
`lawdocs-import` skill (`.claude/skills/lawdocs-import`) does the import and
creates the doc.

`samples/services-agreement.docx` is a small test contract with the usual Word
problems in it. `samples/services-agreement/` holds what the importer makes of it.

## Plan

| Phase | Scope |
| --- | --- |
| 1 (done) | .docx import, legal numbering, outline, Claude Docs output, round trip back to .docx |
| 2 | Versions at every Claude pass with titles, cached redlines, PDF export, block authorship |
| 3 | Defined-term checks, cross-reference checks, references that update on renumbering |
| 4 | Numbering presets (Article I / Section 1.01), schedules and exhibits, footnotes, Word comments and tracked changes as Docs suggestions |

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the design and what
Claude Docs supports today.

## Tests

```
pip install -e '.[dev]'
pytest
```
