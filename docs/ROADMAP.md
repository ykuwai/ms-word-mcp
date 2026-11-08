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

## Next

First, a fix to what is there: **read in bulk.** word_get_outline and
word_get_paragraphs ask Word one paragraph at a time, which the reference
branch measured at 15-20 s for 1,000 paragraphs. Read the whole text in one
call and the styles from the document's XML instead (see PROBES.md).

Then, ordered roughly by what the owner stressed.

1. **Styles, every property.** Create and modify styles: base and next
   style, font (all three scripts and every character property), paragraph
   format (spacing before and after in points and in lines, line spacing,
   indents in characters, East Asian line breaking, grid), borders, shading,
   list linkage, Quick Style gallery. Read a style back whole.
2. **Lists and bullets.** Bullet character, its font and color, numbering
   format and style, indents per level, linking levels to heading styles,
   multilevel lists, restart and continue.
3. **Equations.** Insert from UnicodeMath and LaTeX (Word converts both),
   display and inline, equation numbering with SEQ fields and right-aligned
   numbers, read back as linear text.
4. **Figures and tables.** Pictures, tables with styles, captions with
   numbering, cross-references, list of figures.
5. **Fields.** SEQ, REF, PAGE, TOC and the rest; insert, update, read codes
   and results.
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
