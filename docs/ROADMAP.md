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
  word_set_paragraph_format share, paragraph borders and background color
  among them. On the Mac a style can be changed but not
  created.
- Lists: bulleted and numbered lists with each level's bullet character or
  number format (full-width digits, circled numbers, kanji, aiueo, iroha...),
  the bullet's own font, color and size, positions, levels; each list gets a
  template of its own so no other list changes. Read a list back, change
  levels, remove a list. Numbered headings: levels linked to the heading
  styles number every heading, including ones added later. Restarting the
  numbering in the middle of a list (Windows; Word for Mac's scripting gives
  numbers that change from run to run, docs/PROBES.md).

- Equations: display equations from LaTeX (converted with Word's own Math
  AutoCorrect symbols) or UnicodeMath, built by Word, numbered (1), (2)... at
  the right margin by a SEQ field in the UnicodeMath `#` number, so they
  renumber themselves; listed, and read back as UnicodeMath.

- Fields and tables of contents: a table of contents for the whole document
  or for part of it (a chapter, through a bookmark and \b), any field by its
  code (PAGE, DATE, STYLEREF, SEQ, REF...), listing them, and updating all
  fields, the tables of contents (whole or page numbers only) or the SEQ
  numbers, in the whole document or in some paragraphs.

- Footnotes and endnotes: add one at the end of a paragraph or right after
  some text, list them with the paragraph of each mark, change the text, delete.

- Hyperlinks to an address or a bookmark, over some text or as new text; list,
  change and remove them. The user's own bookmarks: add, list, delete.

- Floating shapes: the shapes of Word's gallery, text boxes (vertical text too),
  lines and arrows, anchored to a paragraph; placed from the page, the margins or
  the paragraph, by points or by name (center, right...); text wrap including
  behind and in front of the text, order among shapes, fill, outline, rotation,
  and text inside with the same font settings as the body. Listed in reading
  order with their text. Rewriting a paragraph keeps a shape anchored at its start.

- Pictures in line with the text, numbered captions under any label in any
  language, live cross-references to captions, equations and headings.
- The page: paper, margins, columns, document grid, vertical text; headers,
  footers and page numbers; ruby.

- Tables: insert one filled with its cells (style, heading rows, grid
  lines, column widths and alignments, autofit, position), list and read
  them, write and format blocks of cells (shading, bold, color, alignment),
  add and delete rows and columns, merge cells, delete a table. Adding
  columns works on Windows only.

## Next

First, fixes to what is there, for long documents (PROBES.md, Reading):

- **word_get_outline** now reads Word's heading list and checks each
  heading (1,000 paragraphs: 2.5 s instead of 28 s on Windows, 3.4 s
  instead of 20 s on the Mac).
- **word_find_replace** on Windows finds about 40 matches a second, match by
  match, so thousands take several calls (docs/PROBES.md, Replacing many
  matches).
- **word_get_paragraphs** still asks Word three or more times a paragraph
  (up to 200 a call; 200 with `detail` took 22 s on Windows and 55 s on the
  Mac). A read now stops after 40 s and says where to go on. The styles could
  come from the document's XML on Windows (Content.WordOpenXML) and through
  VBA on the Mac.
- **word_insert_paragraphs** now sets a style on each run of paragraphs
  that share it and reads the text back in one call (1,000 paragraphs:
  4.5 s instead of 49 s on Windows, 6.5 s instead of 57 s on the Mac).

Then, ordered roughly by what the owner stressed.

1. **Styles, the rest.** Tab stops on the Mac, list linkage, the language
   of a style; creating a style on the Mac (through VBA).
2. **Lists, the rest.** Picture bullets.
3. **Equations, the rest.** Inline equations inside a paragraph,
   chapter-style numbers (2.1) with STYLEREF, cross-references to equations,
   editing an equation in place.
4. **Figures and tables, the rest.** A list of
   figures, numbers with the chapter (2.1); splitting cells, sorting, table
   borders one side at a time.
5. **Japanese typography, the rest.** Removing ruby outright, tate-chu-yoko.
6. **View.** Zoom (percentage, page width, whole page), view type, follow
   edits in a second window.
7. Shapes, the rest. Pictures that float (a picture out of the text and back),
   cropping; grouping; shapes in headers and footers; moving a shape to another
   paragraph.
8. Adding sections (section breaks); a header of several paragraphs.
9. Comments and tracked changes; spelling and a document health check.
10. Page previews.

## Platforms

Every tool has a Mac half, measured on Word for Mac (`scripts/live_check.py`
passes on both platforms). A tool added later gets its Mac half in the same
change, or refuses on macOS and says so until it has one.
