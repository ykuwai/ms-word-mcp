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
- Documents: info, list, choose, open, create, save, save as (and PDF, of some pages, with or without the markup), close.
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
  Pictures float too: put on the page directly, or taken out of the text and back
  (`word_set_picture` and `word_set_shape`), behind or in front of the text; a crop
  in points as the picture is shown, in the text or floating; brightness and
  contrast to fade one behind the text.

- Tracked changes: turn tracking on and off, list the changes, accept or reject
  them by number or all at once. While tracking is on, the text tools read and
  report the text as Word shows it, deletions apart.

- Comments: add one on a paragraph or on some text in it, list them with the
  text each is about, delete them by number or all at once, reply and mark one
  done. Word for Mac's scripting has no replies or done state; on the Mac they
  go through the VBA bridge (backend/bridge.py).

- Checks: Word's spelling and grammar findings and the language the text is
  checked in, marking text with a language, and a health check before sending
  (headings, placeholders, alt text, links, comments and changes left, text
  Word never checks, readability).

- Pictures in line with the text, numbered captions under any label in any
  language, live cross-references to captions, equations and headings.
- The page: paper, margins, columns, document grid, vertical text; section
  breaks; headers and footers, of one paragraph or several, and page numbers;
  ruby.
- A picture of a page as it will print, on both platforms: with comments and
  tracked changes too, which Word for Mac's PDF would draw beside a page shrunk
  to two thirds.

- Tables: insert one filled with its cells (style, heading rows, grid
  lines, column widths and alignments, autofit, position), list and read
  them, write and format blocks of cells (shading, bold, color, alignment,
  borders one side at a time), add and delete rows and columns, merge and
  split cells, sort, delete a table. On the Mac adding columns goes through
  the VBA bridge.

## Next

First, fixes to what is there, for long documents (PROBES.md, Reading):

- **word_get_outline** now reads Word's heading list and checks each
  heading (1,000 paragraphs: 2.5 s instead of 28 s on Windows, 3.4 s
  instead of 20 s on the Mac).
- **word_find_replace** now searches and replaces through the VBA bridge,
  inside Word: 3,000 matches replaced in one call (2.4 s on Windows, 7.4 s on
  the Mac), where a call replaced 740 on Windows and 1,249 on the Mac. On the
  Mac this also gives wildcards and one undo step for the whole replacement.
  Without the bridge it works match by match, as before.
- **word_get_paragraphs** now reads through the VBA bridge, in one call to
  Word (200 paragraphs with `detail`: 0.30 s instead of 11.7 s on Windows,
  0.53 s instead of 27 s on the Mac). Without the bridge it asks Word three or
  more times a paragraph, as before, and stops after 40 s.
- **word_insert_paragraphs** now sets a style on each run of paragraphs
  that share it and reads the text back in one call (1,000 paragraphs:
  4.5 s instead of 49 s on Windows, 6.5 s instead of 57 s on the Mac).

Then, ordered roughly by what the owner stressed.

1. **Styles, the rest.** Done: list linkage and a style's language; tab
   stops and creating a style on the Mac go through the VBA bridge.
2. **Lists, the rest.** Picture bullets.
3. **Equations, the rest.** Chapter-style numbers (2.1) with STYLEREF.
   (Editing an equation in place, cross-references to equations and inline
   equations inside a paragraph are done.)
4. **Figures and tables, the rest.** Numbers with the chapter (2.1). (A list
   of figures, splitting cells, sorting, and table borders one side at a time
   are done.)
5. **Japanese typography, the rest.** Removing ruby outright, tate-chu-yoko.
6. **View.** Zoom (percentage, page width, whole page), view type, follow
   edits in a second window.
7. Shapes, the rest. Moving a shape to another paragraph. (Grouping, and
   shapes in headers and footers, are done.)

## Platforms

Every tool has a Mac half, measured on Word for Mac (`scripts/live_check.py`
passes on both platforms). A tool added later gets its Mac half in the same
change, or refuses on macOS and says so until it has one.
