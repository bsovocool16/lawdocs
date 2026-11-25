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
| `numbering.py` | The legal scheme (1. / 1.1 / (a) / (i) / (A) / (1)), counters, full references (`1.1(a)(i)`), and classification of existing numbers into levels. |
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

- a block with its original kind and level gets its original paragraph
  properties back, so untouched text is laid out exactly as before;
- a new or moved block takes the properties most used for its kind and level
  in the template (house style);
- clause numbers come from Word's numbering, never typed text: the
  template's multilevel list where its styles carry one, otherwise a legal
  list lawdocs adds to numbering.xml;
- footnotes are rebuilt as Word footnotes (original XML when unchanged), and
  bookmarks are re-created so REF fields in footnotes still resolve.

Known gaps: run formatting beyond bold/italic (small caps, colour), static
cross-reference text ("Section 2.15") does not follow renumbering yet
(phase 3), and Word comments are not carried.

## State outside the doc

Phase 2 adds a ledger per document: the binding (lawdocs id to Docs ids),
the version list (Docs rev, title, who made the pass) and block authorship.
It lives in a "Versions" tab inside the doc, so it travels with the doc and is
visible to everyone working on it, with a machine-readable copy beside it.
