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
4. Create the doc with a single `pending` block, then fill it from a blob so
   the (often 100+ KB) Markdown never passes through your context:
   - `batch` create with `markdown: "<?claude block s1?>"` and a pending `s1`;
   - Artifact publish with `asset: true`, `url` = the doc's link,
     `file_path` = `doc.md` → an asset id;
   - Docs `create` object `blob` with `{"asset": "<asset id>"}` → `blob/<id>`;
   - `update` the body: `replace` the pending block with
     `{"from": {"kind": "blob", "id": "blob/<id>"}, "as": "markdown"}`.
   Do not edit numbers by hand: they come from `model.json`.
5. The ack's `xml` lists the new doc's paragraph ids in order. Paragraph *k*
   in that list is paragraph *k* in `blockmap.json`; save the pairing as
   `out/<name>/binding.json` (`{"doc": <doc id>, "node": <body id>, "blocks": {<lawdocs id>: [<docs ids>]}}`)
   so later passes can find a clause's Claude Docs block directly.
6. Open the doc for the user and give them the outline (`outline.md`) if they
   asked for one.

## Back to Word

1. Docs `create` object `blob` with `{"from": {"object": "file", "id": "<tab id>"}, "format": "markdown"}`
   → an asset id; Artifact `read` with `path` = that asset id saves the
   Markdown locally.
2. `lawdocs pull <that.md> --base out/<name>/model.json --editor claude`
3. `lawdocs export out/<name>/model.json -o <name>-edited.docx`

Word numbers the clauses itself on export, so a clause added or deleted in
the doc renumbers everything after it. Cross-reference text does not update
yet: say so when clauses moved.

## Navigating with the outline

Before reading a long imported doc in full, read `outline.json`. Each entry has
the clause's full reference (`2.1(b)(i)`), caption, opening words and block id;
with `binding.json` that goes straight to the Claude Docs block id for a
`{"kind":"view","parentId": ...}` read or a targeted edit.

## Renumbering

After clauses are added, removed or moved, update `model.json` (change a
block's `level`, insert or delete blocks) and run `lawdocs render model.json`.
Every number is recomputed from levels; never type a clause number.
