---
name: lawdocs-import
description: Import a Word (.docx) contract into Claude Docs with clean block structure and legal numbering (1.1(a)(i)). Use when the user hands over a .docx to bring into Claude Docs, or asks for an outline of one.
---

# Import a .docx into Claude Docs

This repository's `lawdocs` package turns a .docx into a block model, renders
it as Markdown that Claude Docs accepts, and builds an outline of the document.
Claude Docs has no import of its own, so this skill is the bridge.

## Steps

1. Install once: `pip install -e .` from the repository root.
2. Run `lawdocs import <file.docx> -o out/<name>`. It writes:
   - `model.json`: the block model (source of truth for numbering)
   - `doc.md`: Markdown for Claude Docs
   - `outline.md` / `outline.json`: clause references, captions, block ids
   - `blockmap.json`: which Markdown paragraph each block produced
3. Tell the user any `note:` lines the command printed (tracked changes
   accepted, comments or footnotes not imported, typed numbers converted).
4. Create the doc in ONE Claude Docs `batch` call, with `container.create.doc.markdown`
   set to the full contents of `doc.md` and `batch: []`. Name it after the
   document title. Do not edit numbers by hand: they come from `model.json`.
5. The ack's `xml` lists the new doc's paragraph ids in order. Paragraph *k*
   in that list is paragraph *k* in `blockmap.json`; save the pairing as
   `out/<name>/binding.json` (`{"doc": <doc id>, "node": <body id>, "blocks": {<lawdocs id>: [<docs ids>]}}`)
   so later passes can find a clause's Claude Docs block directly.
6. Open the doc for the user and give them the outline (`outline.md`) if they
   asked for one.

## Navigating with the outline

Before reading a long imported doc in full, read `outline.json`. Each entry has
the clause's full reference (`2.1(b)(i)`), caption, opening words and block id;
with `binding.json` that goes straight to the Claude Docs block id for a
`{"kind":"view","parentId": ...}` read or a targeted edit.

## Renumbering

After clauses are added, removed or moved, update `model.json` (change a
block's `level`, insert or delete blocks) and run `lawdocs render model.json`.
Every number is recomputed from levels; never type a clause number.
