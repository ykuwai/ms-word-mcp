# Roadmap

What ms-word-mcp is for, what it does today, and what comes next. The order
follows what the owner asked for in ykuwai/ppt-mcp#352 and since.

## Principles

- **Working beside the user.** Writes go through ranges, never the selection,
  so the user keeps typing where they are. The view follows the AI's edits
  (`WORD_FOLLOW_EDITS`). Each tool call is one undo step.
- **Read back after every write**, and report what did not take as a warning.
- **Official sources first.** Every member and constant is checked against
  Microsoft's documentation ([reference/word-object-model.md](reference/word-object-model.md))
  and, where it matters, measured on a real Word ([PROBES.md](PROBES.md)).
- **Catalog mode** keeps the tool list small in the client's context.
- **Both platforms**, Windows over COM and macOS over Apple Events, with the
  same tools and arguments. A tool that cannot do something on one platform
  says so and names what to use instead.

## Done

- Server: catalog mode, per-call `document`, unknown-argument refusals,
  failures marked as errors, the single Word worker with busy handling and
  recovery from a stuck call.
- Documents: info, list, choose, open, create, save, save as (and PDF), close.
- Text: outline, paragraphs, insert paragraphs with styles, rewrite, delete,
  find and replace.
- Formatting: fonts per script (Latin, East Asian, complex script), size,
  bold, italic, underline, color; paragraph alignment, line spacing (lines or
  points), spacing, indents, grid snapping; apply and list styles.
- Undo and redo.
- Styles: read a style whole (word_get_style) and define one (word_set_style):
  type, base and next style, gallery settings, and every character and
  paragraph property of tools/_props.py, which word_format_text and
  word_set_paragraph_format share. On the Mac a style can be changed but not
  created.
- Lists: bulleted and numbered lists with each level's bullet character or
  number format (full-width digits, circled numbers, kanji, aiueo, iroha...),
  the bullet's own font, color and size, positions, levels; each list gets a
  template of its own so no other list changes. Read a list back, change
  levels, remove a list.

- Equations: display equations from LaTeX (converted with Word's own Math
  AutoCorrect symbols) or UnicodeMath, built by Word, numbered (1), (2)... at
  the right margin by a SEQ field in the UnicodeMath `#` number, so they
  renumber themselves; listed, and read back as UnicodeMath.

- Fields and tables of contents: a table of contents for the whole document
  or for part of it (a chapter, through a bookmark and \b), any field by its
  code (PAGE, DATE, STYLEREF, SEQ, REF...), listing them, and updating all
  fields, the tables of contents (whole or page numbers only) or the SEQ
  numbers, in the whole document or in some paragraphs.

## Next

First, a fix to what is there: **read in bulk.** word_get_outline and
word_get_paragraphs ask Word one paragraph at a time, which the reference
branch measured at 15-20 s for 1,000 paragraphs. Read the whole text in one
call and the styles from the document's XML instead (see PROBES.md).

Then, ordered roughly by what the owner stressed.

1. **Styles, the rest.** Borders and shading, tab stops, list linkage, the
   language of a style; creating a style on the Mac (through VBA).
2. **Lists, the rest.** Numbered headings (levels linked to heading styles),
   restarting numbering in the middle of a list, picture bullets.
3. **Equations, the rest.** Inline equations inside a paragraph,
   chapter-style numbers (2.1) with STYLEREF, cross-references to equations,
   editing an equation in place.
4. **Figures and tables.** Pictures, tables with styles, captions with
   numbering, cross-references, list of figures.
5. **Cross-references.** Bookmarks, REF fields to headings, figures, tables and
   equations.
6. **Japanese typography.** Vertical text (section text direction, document
   grid), ruby (phonetic guide), emphasis marks, East Asian line breaking
   rules, character-unit indents, mixed-script spacing.
7. **View.** Zoom (percentage, page width, whole page), view type, follow
   edits in a second window.
8. Headers, footers and page numbers; page setup; sections; table of
   contents.
9. Comments and tracked changes; spelling and a document health check.
10. Page previews.

## Platforms

Every tool has a Mac half, measured on Word for Mac (`scripts/live_check.py`
passes on both platforms). A tool added later gets its Mac half in the same
change, or refuses on macOS and says so until it has one.
