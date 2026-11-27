# lawdocs architecture

lawdocs is a block-level engine for legal documents that uses Claude Docs as
its editor. The engine owns structure (what is a clause, at what level, who
wrote it, which version it belongs to). Claude Docs owns the editing surface,
sharing, comments and history.

## The one rule

**Presentation is derived from structure, never stored.** A block has a
`kind` (title, heading, clause, text, table) and a `level` (1 to 6). The clause
number, its heading treatment and its place in the outline are all computed
from those two fields. Nothing in the model can say "this paragraph is
indented 0.5 inch with a hanging indent and the number is 3.2(b)". That is
the Word failure mode this project exists to remove.

## Components

```
.docx ──► docx_import ──► Document (blocks + Word data) ──► render ──► Claude Docs
  │                              ▲                                      │ edits by people and Claude
  │ (kept as template)           └──────────── md_import.pull ◄─────────┘ Markdown export
  └──────────────────────────────► docx_export ──► .docx
```

| Module | Job |
| --- | --- |
| `model.py` | `Block` / `Document` / `Run`. The only structural fields are `kind` and `level`. |
| `docx_import.py` | Reads document.xml, styles.xml and numbering.xml. Resolves list numbers the way Word displays them (style-linked lists, overrides, restarts), converts hand-typed numbers to structure, accepts tracked changes, keeps bold/italic, drops all other formatting. |
| `numbering.py` | Counters, number styles, full references (`1.1(a)(i)`), the default legal scheme (1. / 1.1 / (a) / (i) / (A) / (1)), and reading existing numbers into levels, styles, restarts and alternatives. |
| `render.py` | Markdown for Claude Docs, plus the block map that ties each lawdocs block to the Claude Docs paragraphs it produced. |
| `outline.py` | The document's structured outline for people and agents. |
| `md_import.py` | Reads a Claude Doc's Markdown export back into blocks and lines them up with the previous model, so surviving blocks keep their id, author and Word data. |
| `docx_export.py` | Writes .docx using the source as the template (see below). |
| `cli.py` | `lawdocs import / render / outline / pull / export`. |
| `.claude/skills/lawdocs-import` | How Claude runs the import and writes the result into Claude Docs. |

### Why the level comes from the displayed number

Word's list indent level (`ilvl`) does not mean the same thing across lists.
A firm template's `(a)` might be ilvl 2 of the main list, while someone's
ad-hoc `(a)` sub-list is ilvl 0 of a second list. The importer computes the
number Word would display and classifies that (`1.1` is level 2, `(a)` level
3, `(i)` level 4, with `(i)` after `(h)` treated as a letter). `ilvl` is only
the fallback.

## What Claude Docs gives us (checked 2026-11-25)

| Capability | Status | Used for |
| --- | --- | --- |
| Block tree with stable ids and per-block `rev` | Yes | Binding lawdocs blocks to doc blocks; detecting which blocks changed |
| Read any past revision (`atRev`) and changes since a revision (`sinceRev`) | Yes | Versions and redlines (phase 2) |
| Suggestions (tracked changes) with who proposed them | Yes | Laying Claude passes as reviewable changes |
| Comments with author | Yes | Review workflows |
| Export a tab to PDF, Word, HTML, Markdown | Yes | Export (redline PDFs need our own renderer) |
| Ordered lists | 1, 2, 3 at every depth only | Not used; numbers are written as text by `render` |
| Paragraph indentation or styles | No | Hierarchy is shown by numbering only |
| Import of .docx | No | `docx_import` |
| Named versions | No | Our version ledger (phase 2) |
| Per-block author in reads | No | Our authorship ledger (phase 2) |

Three Docs features would make lawdocs simpler, worth raising with the
Claude Docs team: per-level ordered-list formats (1 / 1.1 / (a) / (i)),
named revisions, and a block's last writer in reads.

## Word round trip

The source .docx is kept as the template, and each block carries the raw
Word data it came from (paragraph properties, caption form, style-separator
partner, bookmarks, raw table and footnote XML, spacer paragraphs after it).
Claude Docs never sees any of it. On export:

- a block whose kind, level and text are unchanged is written back as the
  source paragraph itself: its properties and its runs, fields, bookmarks
  and comment anchors, verbatim (footnote ids remapped). Untouched text keeps
  every bit of Word formatting, including underlined defined terms, small
  caps and live REF cross-references, which then follow renumbering in Word;
- an edited block with its original kind and level gets its original
  paragraph properties back and its text rebuilt (bold, italic, footnotes);
- a new or moved clause is numbered the way its nearest sibling is: in the
  sibling's Word list (so Word counts them together) or typed like it;
  without a sibling it takes the properties most used for its kind and
  level (house style);
- a clause the source numbered by hand stays typed, with lawdocs writing the
  current number; untouched, it goes back verbatim. A clause with no
  numbering at all gets a legal list lawdocs adds to numbering.xml;
- footnotes are rebuilt as Word footnotes (original XML when unchanged), and
  bookmarks are re-created so REF fields in footnotes still resolve.

Checked on the NVCA model SPA: an unedited round trip renders identically
to the source on all 107 pages (LibreOffice, pixel comparison). For a wider
check, `tools/corpus_check.py` runs import, the Claude Docs read-back and
export over a folder of documents and compares numbers with what
LibreOffice shows and pages pixel by pixel.

`tools/corpus_selftest.py` checks the checker. It breaks each document on
purpose (drops or swaps a paragraph, changes a digit or a word, moves a list
item down a level, widens a margin, loses a paragraph on the way back from
Claude Docs, misreads one clause number) and confirms that one of
corpus_check's comparisons notices, using the unbroken export as the
control. Any break that goes unnoticed is a blind spot in the harness.

## Numbers

Only counters come from structure. How a clause's number is written is its
`style`, a template with the counters as slots, e.g.
`{1:decimal}.{2:decimal}.{n:decimal}` for "1.1.1" or `Article {n:upperRoman}`.
Import sets each clause's style from the number Word showed, so a document
keeps its own forms (1.1.1, Article I, (1), a.); a clause without a style takes
the style of the clause before it at the same level, else the default legal
scheme. Two flags keep imported numbers equal to the source's: `restart` (a
schedule or part numbers from 1 again) and `alternative` (a model form's
"[Alternative] 4.4.4", which repeats the number before it).

Levels are read in context: numbers that spell out their depth ("1.1.1",
"Article 2") take it; "(a)", "(1)", "(i)" nest under the deepest open clause
unless they match the form of one still open.

Known gaps: an edited paragraph keeps only bold, italic and footnotes from
its runs (underlining and fields inside it are lost); cross-references typed
as plain text ("Section 2.15") do not follow renumbering yet (phase 3); Word
comments survive only on untouched paragraphs and are not shown in Claude
Docs.

## State outside the doc

Phase 2 adds a ledger per document: the binding (lawdocs id to Docs ids),
the version list (Docs rev, title, who made the pass) and block authorship.
It lives in a "Versions" tab inside the doc, so it travels with the doc and is
visible to everyone working on it, with a machine-readable copy beside it.
