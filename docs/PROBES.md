# Measured behaviour

What Word was seen to do, on which version, before code relied on it. A line
here is a measurement, not a reading of the documentation; the documentation
is collected in [reference/word-object-model.md](reference/word-object-model.md).

## Word for Mac 16.114.4 (macOS, Apple Events through appscript)

Measured 2026-11-08 with `appscript.app(id="com.microsoft.Word", terms="sdef")`.

| What | Seen |
|---|---|
| Dictionary | `Word.sdef` in the app bundle is complete and follows the VBA object model closely: `text range`, `paragraph`, `paragraph format`, `font`, `Word style`, `list format`, `list level`, `list template`, `create range`, `undo`, `redo` and the math classes are all there. |
| No document open | `documents.get()` answers missing value, not an empty list. |
| A document past the last one | `documents[i].name.get()` answers missing value, not an error. A walk that waits for an error never ends; stop at the first missing name. |
| Counting paragraphs | `document.paragraphs.get()` and `count(each=k.paragraph)` both fail to decode the answer (`unpack requires a buffer of 4 bytes`). Asking for `paragraphs[i]` past the last one is error -1728. Count another way. |
| Making a range | `document.create_range(start=0, end_=0)`: appscript renames `end` to `end_`. |
| Writing text | Setting `content` of a range made with `create_range` inserts the text; `\r` starts a new paragraph. |
| Reading text | `paragraphs[i].text_object.content.get()` answers the paragraph's text with its mark (`'Appended\r'`). |
| Latin font | `font_object.ascii_name` and `other_name` set and read back. |
| East Asian font | `east_asian_name` set to "Hiragino Sans" took. Set to "Arial" (no East Asian glyphs) it is refused with -1708 and keeps the old value. |
| Complex-script font | `complex_script_name` reads back as a theme reference, `+本文のフォント - コンプレックス スクリプト ()`, until a font is set. |
| Line spacing | A new paragraph reads `line_spacing_rule` = `line_space_single`, `line_spacing` = 12.0. Setting the multiple rule with 18 pt reads back as `line_space1_pt5`, 18.0: Word turns 1.5 lines into its own rule. |
| Integers | `count(each=k.paragraph)` and `compute statistics` answer an integer typed as 32-bit but 8 bytes long, which appscript fails to decode. Decoding 8 bytes as 64-bit fixes both (backend/mac.py). `paragraphs.get()` still fails, in another decoder. |
| Counting paragraphs, again | Probing `paragraphs[i]` (doubling, then bisecting) counts them in a few calls. |
| Ranges are re-resolved | A range reached through a reference (`paragraphs[3].text_object`) is asked for again each time: after `execute find` it still reads as the whole paragraph, so a found range cannot be read through it. Search the text in Python and make ranges with `create range`. |
| Selection | `execute find` on a range left the selection at (0, 0). |
| Style | Setting `style` to `k.style_heading1` applies 見出し 1; `outline_level` reads `outline_level1`, and `outline_level_body_text` for body text. `Word_styles["見出し 2"]` finds a style by its local name. |
| Zoom | `active_window.view.zoom.percentage` reads 100. |
| Font color | Word for Mac takes 16-bit channels: setting `[192, 0, 0]` reads back black, `[49344, 0, 0]` reads back as #C00000. |
| Spacing in lines | Setting `line_unit_before` to 0.5 made `space_before` read 9.0: the measure in lines wins while it is set (the same for indents in characters). |
| Creating a style | `make new Word style` answers -1708 and makes nothing, with or without a type, whatever the name. Changing an existing style (base style, next style, font, paragraph format) works. A built-in style is reached with `Word_styles[k.style_heading1]`; `Word_styles["Heading 1"]` fails in a Japanese Word. |
| Normal's base style | Answers a style named `""`. |
| Text before an equation | Setting `content` of an empty range at the start of a paragraph that holds an equation is refused (-1708). Text set at the end of the paragraph before, as "\r" + text, works. |
| List templates | `make new list template at` the document works (without `at`, -1708). Changing its levels leaves the bullet gallery and the document's other lists as they were. `apply list format template` on a range's `list format` applies it. |
| A level's number style and format | Setting `number style` to bullet is refused until `number format` is already the bullet character; a `number format` holding `%1` is refused until the style is a numbering one. |
| Decoding, again | Patching appscript's decoder class is not enough: each codecs object keeps its own table, made when it is created, so `get list gallery` still failed until the tables of the existing objects were patched too. |
| Style | `paragraphs[1].style.get()` answers a `Word_styles['標準']` reference; `style.name_local.get()` answers `標準` in a Japanese Word. |
| End of the text | `text_object.end_of_content` answered 28 for 27 characters of text (the last paragraph mark counts). |
| Closing | `close(saving=k.no)` closes without asking. |
| New document name | `create_new_document()` names it `文書 1`, `文書 3`: the number is not reused. |

## Word for Windows (Microsoft 365, 16.0.20527 x64, Japanese UI)

Measured 2026-11-08 with `scripts/live_check.py` (17 of 20 steps passed on
the first run; the three failures are below and were fixed).

| What | Seen |
|---|---|
| Style names | Built-in styles read back in the UI language: 標準, 見出し 1, 表題. Applying them by WdBuiltinStyle constant works. |
| Theme fonts | A style's fonts read back as theme references: `+見出しのフォント`, `+本文のフォント - 日本語`, `+本文のフォント - コンプレックス スクリプト`. |
| East Asian font | Setting `NameFarEast` to Segoe UI (no East Asian glyphs) raises `0x800A16D4` and keeps the old font. |
| Line spacing | `LineSpacingRule` = multiple with `LineSpacing` = 18 reads back as `wdLineSpace1pt5`, 18.0. |
| Replace all | `Find.Execute(..., ReplaceWith="dog", Replace=wdReplaceAll)` on `Document.Content` answered without error and replaced nothing; the match was still there. Cause not known. Replacing each match through its own range works. |
| Undo record | One `UndoRecord` custom record per call: a single `Document.Undo(1)` took back a whole rewrite. |
| Font.Bold | Reads back -1 for bold. |
| Page | A new document from Normal: A4 (595.3 × 841.9 pt), margins 99.25 / 85.05 / 85.05 / 85.05 pt. |

What the documentation leaves unstated, measured with `scripts/probes/windows_checks.py`:

| What | Seen |
|---|---|
| `DisableLineHeightGrid` | True writes `<w:snapToGrid w:val="0"/>`: the paragraph does **not** snap to the grid. The property's name is right and the documentation's sentence is inverted. Reads back -1 / 0. |
| `Window.ScrollIntoView` | Leaves the selection where it was (0, 0 before and after). |
| `Range.Paragraphs` | Counts a paragraph the range covers in part: `Range(5, 6)` counts 1; `Range(0, start of paragraph 3)` counts 2, and one character further counts 3. |
| `Chr(11)` | A line break inside the paragraph: the paragraph count stays the same. |
| The last paragraph mark | Deleting it changes nothing (67 characters before and after). |
| `Font.Name` | Times New Roman set NameAscii, NameOther and NameBi and left NameFarEast; 游ゴシック set NameAscii and NameFarEast and left the other two. Which slots it writes depends on the font, so the tools set each slot by name. |
| Mixed ranges | SpaceBefore, Alignment, SizeBi and Color all read 9999999 (wdUndefined) over two paragraphs that differ. |
| `Font.Bold` | Reads as the Python int 0 or -1. |
| `Document.Path` | An unsaved document's is `''`. |
| `Styles.Add` | Takes the paragraph style and the manual formatting of wherever the user's cursor is: a style added with the cursor in a bold Title paragraph came out based on Title, 28 pt, bold, condensed 0.5 pt, kerning from 14 pt. A character style took the bold too. `Font.Reset` on the new style raised 0x800A174B. Setting its base style and assigning the base style's `Font.Duplicate` and `ParagraphFormat.Duplicate` gives a style that differs from its base in nothing (`scripts/probes/windows_style_steps.py`: setting BaseStyle alone leaves the font as it was; the Font copy is what cleans it). word_set_style first set the base again after the copy, and the new style still came out with the Title's font, spacing and kerning; setting the base once, before the copy, fixed it (live check, 34 of 34). |
| `Style.Linked` | Readable on paragraph styles only (40 of 93 linked); character, table and list styles raise "this property cannot be used for this object". |
| `Font.Name` again | Set to 游ゴシック, `NameOther` read back as Times New Roman in one run and as 游ゴシック in the next, with the same code. Not reliable: set each slot by name. |
| `Document.Undo(50)` | With one edit to undo, answers False rather than raising. |
| UndoRecord, target active | Three edits inside one record were undone by one `Undo(1)`. `IsRecordingCustomRecord` was True while editing. |
| UndoRecord, target **not** active | `IsRecordingCustomRecord` was False right after the edits: writing to a document that is not the active one ends the record. One `Undo(1)` took back only the last edit, a second the one before. The server does not switch the user's window to keep it; a call on such a document says in its result (`undo`) that it is not one undo step. |

## Earlier measurements by AlonSoko

AlonSoko measured Word 16 on both platforms while writing an earlier version of
this server. What bears on this one, in short:

- **Reading one paragraph at a time is slow.** 1,000 paragraphs with their
  outline level took 20 s through COM and 15 s through Apple Events. The whole
  text comes in one call (13 ms), the headings in one call
  (`GetCrossReferenceItems`), and styles with the document's XML
  (`Content.WordOpenXML`, 84 ms). word_get_paragraphs here still walks
  paragraph by paragraph; see ROADMAP.
- **Starting Word for Mac:** with `open -g -j` (hidden) Word loses track of
  its documents once the front one closes; `open -g` does not.
- **Commands that make their document the active one on the Mac:**
  `make new table` (fails once, then works) and `create new field`. The
  user's active document has to be put back afterwards.
- **The Mac sandbox:** saving outside Word's container raises an access
  prompt; inside it, none. Files are written in the container and moved.
- **Undo on the Mac** is one step per Apple Event. VBA's `UndoRecord` groups
  edits inside one macro run only.
- **Word for Mac's PDF export** always draws comments and tracked changes in a
  margin.
- **Language:** text marked with the custom locale `0x0C00` is never
  spell-checked.
- **Tables of contents:** a `TOC` field made with `create new field` gives the
  same table as Word's own command, on both platforms.

## Reading a long document (Word 16, both platforms)

Measured with `scripts/probes/mac_bulk_reads.py` and
`scripts/probes/outline_speed.py` on 2026-11-08 (1,000 paragraphs).

| What | Windows | Mac |
|---|---|---|
| One paragraph's outline level | | 17 ms a paragraph (200 in 3.4 s) |
| A property of every paragraph in one Apple Event (`outline level of every paragraph`, `content`, `style`) | | 22-25 s each: slower than one at a time. Paragraph formatting, fonts, a slice of paragraphs, or their starts: no answer within 60 s (-1712). |
| The whole text | | One call, 0.00-0.02 s |
| The whole text split at `\r` | | 31 pieces for 19 paragraphs with two equations (each equation's text holds marks of its own) and a table of contents: the split cannot number paragraphs by itself |
| Word's heading list (`get cross reference items` with `reference type heading`) | | One call, 9 ms. Paragraphs in a heading style only: a Normal paragraph given outline level 2 directly is not in it. Each item is the heading's text indented two spaces a level; a tab or line break becomes a space, the heading's own leading spaces are dropped, and a long heading is cut at about 100 characters. |
| `go to next` a heading item from the start | | Stopped moving after three headings: not used |
| word_get_outline, 1,000 paragraphs, 40 headings, 10 numbered equations | Every paragraph: 28.2 s and 37.0 s in two runs. Heading list, each heading read back: 2.5 s and 6.7 s, the same 40 headings | Every paragraph: 19.9 s. Heading list, each heading read back: 3.4 s, the same 40 headings |
| word_get_paragraphs, 200 of 1,000 paragraphs (`scripts/probes/paragraph_read_speed.py`) | 8.7 s and 12.9 s; with `detail` 19.8 s and 22.0 s | 11.9 s and 11.3 s; with `detail` 55.1 s and 55.3 s, close to the 69 s a call may run, so a read stops itself after 40 s |
| word_insert_paragraphs, 1,000 paragraphs | 48.6 s, a style set and read back a paragraph at a time; 4.5 s by runs of a style, read back in one call | 57 s the same way; 6.5 s with a style set on each run of paragraphs that share it and the text read back in one call |

## Paragraph borders and background color (Word for Mac 16)

| What | Measured |
|---|---|
| `get border` of a paragraph or a style | Works: line style, line width and color can be read and set |
| `get border` of a paragraph format | Answers a reference, and every read or write of it fails with -1728 |
| Color | 16-bit channels, as for fonts: `[200, 220, 255]` is black, `[51400, 56540, 65535]` is #C8DCFF |
| `shading` of a paragraph, its format, its text, or a style | All work and agree |
| A background taken off | `background pattern color index` set to `no highlight`; the color then reads `[0, 0, 0]`, so the index is what tells |

## Restarting a list's numbering (Word for Mac 16)

A list numbered 1 to 5 alone in a document restarts at the third paragraph
when `apply list template` with `list apply to this point forward` is given a
range collapsed at that paragraph's start (1, 2, 1, 2, 3). Given the whole
paragraph, it restarts the next one: the point is the range's end.

In a document that also has numbered headings, the same steps gave different
numbers from one identical run to the next, and they stayed that way when read
again later. A list of four, restarting at the third:

| Method | Results |
|---|---|
| Collapsed range, this point forward | 1, 2, 1, 2 in 3 of 8 runs; 3, 4, 1, 2 in 5 of 8 |
| The rest of the list, `selection` | 1, 2, 1, 2 in 2 of 5 runs; unchanged in 3 of 5 |
| A new list over the rest (`word_apply_list`) | 1, 2, 1, 2 once; 3, 4, 1, 2 once; 1, 2, 3, 1 three times out of 5 |

So word_restart_numbering refuses on macOS. The last row also means
`word_apply_list` over part of an existing list is unreliable on the Mac.

## Numbered headings (Word 16, both platforms)

`scripts/probes/numbered_headings.py`: `word_apply_list` on the first heading
alone, with level 1 linked to Heading 1 and level 2 to Heading 2, numbers every
heading in the document (1, 1.1, 1.2, 2, 2.1). Body text stays unnumbered, and a
Heading 2 added afterwards reads 2.2. Word for Mac and Word for Windows gave the
same numbers, on a Japanese Word, from the English style names.

## Replacing many matches (Word 16, both platforms)

`scripts/probes/find_replace_speed.py`: 1,000 paragraphs holding 3,000 matches.
One call stops itself after 20 s of searching and 45 s in all, inside the 60 s
(Windows) and 84 s (Mac) a call may run, and answers `complete: false` with how
far it got. Calling again does the rest.

| | Count the matches | Replace, one call |
|---|---|---|
| Windows (`Find.Execute`, match by match) | 827 in 20.1 s, then out of time | 740 in 31.7 s, 20 s of it searching |
| Mac (the whole text, in Python) | all 3,000 in 3.6 s | 1,249 in 45.2 s |

One paragraph holding 520 matches is slower on Windows: 433 replaced in a call.

## Pictures, captions, references, the page, ruby (Word 16, both platforms)

Measured with `scripts/probes/windows_layout.py` and `scripts/probes/mac_layout.py`
on 2026-11-08, then through the live check.

| What | Windows | Mac |
|---|---|---|
| A picture from a file | `InlineShapes.AddPicture(path, False, True, range)`, no prompt; in a document that is not active | `make new inline picture` in a document that is not active. **A file outside Word's container makes Word ask the user, and the call times out (-1712) while it waits**, per file, even in a folder where Word has saved documents. From a copy inside the container: no prompt. |
| Its natural size, 40 x 20 pixels | 30 x 15 pt | 40 x 20 pt |
| Setting one side | Width alone leaves the height; with `LockAspectRatio` on the height follows | The width alone leaves the height, so both are set |
| A caption | Paragraph in Caption, `SEQ Figure \* ARABIC` over a placeholder: "Figure 1: A red box". `GetCrossReferenceItems("Figure")` lists it. | The same. `get cross reference items` takes only the enumerated types (-1708 for a label). |
| A cross-reference | Hidden bookmark `_Ref100000001` over "Figure 1" (start of the paragraph to the field's result end + 1) and `REF _Ref100000001 \h`: "Figure 1". Not in `Bookmarks` unless `ShowHidden` | The same, `make new bookmark` `at` the document. A bookmark round a numbered equation's SEQ field works. |
| Which calls change the active document | `Range.PhoneticGuide` | `create new field`, `merge cell`, `phonetic guide`: the target is made active and the user's document given back |
| Headers and footers | `Sections(n).Headers(i).Range.Text`; `doc.Fields.Add(range.Characters(n), -1, "PAGE", False)` | `get header`/`get footer` of a section; `create new field` over `characters n` of its `text object`; `create range` is the main text's |
| Page numbers | `PageNumbers.NumberStyle`, `RestartNumberingAtSection`, `StartingNumber` | `page number options` of the footer: the same |
| Vertical text | `Section.Range.Orientation = 1` (wdTextOrientationVerticalFarEast): the page turns (595 x 842 becomes 842 x 595) | `orientation` of the section's `text object`, `text orientation vertical east asian`: the same |
| Tables when the direction changes | Not measured apart from the live check, which passes | They float (`wrap around text` true) and go left; in vertical text `wrap around text` cannot be turned off. Back in horizontal text, wrap off first: setting it resets a table's alignment. |
| Ruby | `Range.PhoneticGuide(text)` makes an EQ field (`\o\ad(\s\up 10(かんじ),漢字)`), and the text it covers reads as one `\x15` | `phonetic guide` with every parameter (`raise` 0 places it): the same field, `\x15` too |
| A position counted from a paragraph's text | Runs behind Word's once a field is in the paragraph: the text shows a field's result and hides its code | The same. A second cross-reference placed that way went inside the first's code (-1708, and the first was damaged; seen in a Sonnet trial), so places are read back from Word. |
| Changing text that has a field or ruby | A range that stops inside a field | The same: `content` of a range stopping inside one is refused (-1708). To the paragraph's real end it is taken. |
| Changing ruby | `Field.Delete`, then the words inserted at the field's place | `delete` of the field, the same |
| Document grid | `LinesPage = 30` or `CharsLine = 40` refused in vertical text ("the value is out of range") | Taken |

## Tables (Word 16, both platforms)

Measured with `scripts/probes/windows_tables.py` (Microsoft 365 16.0.20527,
Japanese UI) and `scripts/probes/mac_tables.py` (Word 16.114.4) on
2026-11-08, in scratch documents closed unsaved.

| What | Windows | Mac |
|---|---|---|
| Making a table | `Tables.Add(range, rows, columns)` at a collapsed range at the start of an empty paragraph; works in a document that is not the active one | `make new table at (create range start s end s) with properties {number of rows, number of columns}`; **-1708 unless the document is the active one**. `activate object` on the document makes it active, and on the other document gives the user's back (both measured). |
| Style of a new table | Normal Table (標準の表), `Borders.Enable` 0, `DefaultTableStyle` "" | Normal Table (標準の表), `enable borders` false, `default table style` empty |
| The paragraph it was made in | Left after the table, empty; `Range.Delete` on it removes it (answers 1). At the end of the document it stays. | The same; `delete` on its range removes it |
| An English style name | `Table.Style = "Table Grid"` and `Styles("Table Grid")`: "the requested member does not exist" on a Japanese Word | `set style to "Table Grid"`: no error, and the style stays Normal Table |
| A WdBuiltinStyle constant | `Table.Style = -161` (Light Grid) gives 表 (モノトーン) 3 | `style table light grid` gives 表 (モノトーン) 3; `style table medium shading1 accent1` gives 表 (青) 4. A style object (`Word styles "表 (格子)"`) works too. |
| Heading rows after a style | Not measured (the tools set them again either way) | **Cleared**: `heading format` of row 1 reads false after `style` is set |
| Adding rows | `Rows.Add()` at the end, `Rows.Add(Rows(1))` before the first | `make new row at end of rows`: -1708. `at after row n` and `at before row 1` work. `insert rows` works on the selection only. |
| Adding columns | `Columns.Add()`, `Columns.Add(Columns(1))` | `make new column` (after or before a column): -1708. `insert columns` works on the selection only. |
| Deleting rows and columns | `Rows(i).Delete`, `Columns(i).Delete` | `delete row i`, `delete column i` |
| Merging | `Cell(1,1).Merge(Cell(1,2))`: the cell holds both texts, a paragraph each; `Uniform` False; `Columns(1)` can no longer be reached ("cell widths differ") | `merge cell ... with` the same; `columns` can no longer be reached (-1728) |
| `Uniform` | Read False in the live check for a table none of whose cells had been merged (rows added at the bottom, fitted to the window, cells shaded); why is not measured, so the tools do not ask it before reaching a column | True until cells are merged |
| With another document active | Not measured beyond `Tables.Add` | Writing cells, `make new row`, `delete row`/`column`, style, heading format, borders, autofit, widths, row alignment, shading and fonts all work. **`merge cell` makes the table's document the active one.** (`scripts/probes/mac_tables_inactive.py`) |
| A cell merged down a column | Not measured with the probe; the live check reads and restyles such a table | `rows` can no longer be reached (-1728) and `make new row` is refused (-1708); `get cell from table` for the cell under the merged one answers -1708; the table's text keeps an empty place for it (`r1c1\rr2c1\r\x07r1c2...\r\x07\r\x07` then an empty cell in row 2) |
| A cell never shaded | | `background pattern color` reads [0, 0, 0] with `background pattern color index` `no highlight`; setting that index takes shading off |
| Cells of a row with merged cells | `Rows(1).Cells.Count` 3 of 4 | `count rows 1 each cell` 2 of 3; `get cell from table row 1 column 2` is the cell after the merged one |
| A table's text | Each cell ends in `\r\x07`, each row too after its cells | The same |
| Shading | `Cell.Shading.BackgroundPatternColor` (BGR) | `background pattern color` of the cell's `shading`, 16-bit channels |
| In a table? | `Range.Information(12)` (wdWithInTable) | `get range information ... information type with in table` answers the **text** "true" or "false" |
| Deleting a table | `Table.Delete`: the paragraphs around it stay | `delete table`: the same |

## Equations (Word 16, both platforms)

Measured with `scripts/probes/windows_equations.py` and `scripts/probes/mac_equations.py`;
what Microsoft documents is in [reference/equations.md](reference/equations.md).

| What | Windows | Mac |
|---|---|---|
| Making an equation | `OMaths.Add(range)` then `OMath.BuildUp()` | `create new equation from range r in document d`: both arguments are needed (`from range` alone, `in range`, `in selection` and `make new math object` answer -1708). Then `build up` of the document's `math object`. |
| UnicodeMath | `a/b`, `E=mc^2+√(x^2+1)`, `∫_0^∞ e^(-x^2) dx=√π/2` build up as they should | The same |
| TeX control words | `\alpha`, `\frac{a}{b}`, `\sqrt{x}` stay as text (Math AutoCorrect does not act on text put in by code) | The same |
| `#(1)` | Builds up into an equation array holding `#(1)` (its text has no visible number; to be checked on the page) | Builds up into an equation array, and **on the page** the equation is centred at display size with (1) at the right margin: the text read back does not show it, a PDF of the page does. |
| A SEQ field as the number | To be measured | `create new field` over the `1` of `#(1)` after building up, with `field sequence` "Equation", numbers three equations (1), (2), (3), right-aligned, display size ([picture](images/equation-numbers-mac.png)). |
| Reading back | `Range.Text` of a built-up equation is its pieces split by `\r`; after `Linearize`, `Range.Text` is UnicodeMath (`∫_0^∞▒〖𝑒^(−𝑥^2 ) 𝑑𝑥〗=√𝜋/2`), letters in math italic | The same, through the paragraph's content |
| Display or inline | `Type` 0 (display) for an equation alone in its paragraph | `display type` = `equation display professional` |
| Numbering | A three-cell table, equation in the middle cell, `(` SEQ Equation `)` in the right cell: (1), then (2) | `create new field` with `field type field sequence` gives (1) |
| MathML through `InsertXML` | With `MML2OMML.XSL` (in `C:\Program Files\Microsoft Office\Root\Office16`) as the transform: "XML markup cannot be inserted at this location" (25625). Plain MathML: no error, no equation. | No `InsertXML` in the dictionary |
| Math AutoCorrect list | `Application.OMathAutoCorrect.Entries` | Each entry's `value` answers missing value |

## Footnotes and endnotes (Word 16)

Measured on the Mac with `scripts/probes/mac_notes_links.py`; on Windows through
`scripts/live_check.py`.

| What | Windows | Mac |
|---|---|---|
| Making a note | `Footnotes.Add(range)`, `Endnotes.Add(range)`, then `Range.Text` | `make new footnote at document with properties {text object: range}`. Made `at` a range or the end of a paragraph, the note went to the start of the document. |
| A new note's text | Empty | The literal text "No Reference" until it is written |
| The mark in the paragraph | `\x02`, one position | The same |
| With another document active | Not measured | Works, and the active document stays the same |
| A note emptied | | `content` answers missing value |
| Styles | | The note takes Footnote Text, its mark Footnote Reference |

## Hyperlinks and bookmarks (Word 16)

Measured on the Mac with `scripts/probes/mac_notes_links.py`; on Windows through
`scripts/live_check.py` and `scripts/probes/hyperlink_removal.py`.

| What | Windows | Mac |
|---|---|---|
| Making a link | `Hyperlinks.Add(anchor, address, sub_address)`, `""` for the one not used | `make new hyperlink object at document with properties {text object, hyperlink address or sub address}` |
| The Hyperlink style | | The visible text takes it |
| Order of the links | | By place in the text |
| Counting links in a range | | `count range each hyperlink object` answers 0 whatever is in the range |
| Emptying the address | `Address = ""` leaves it as it was | Setting `""` leaves it as it was. So a link to an address that is to go to a bookmark is taken off and made again. |
| Removing a link | `Hyperlink.Delete`: the text stays, and goes from Hyperlink back to Default Paragraph Font, with no underline and the automatic color | `delete`: the text stays, and goes from Hyperlink back to Normal |
| A bookmark's name again | `Bookmarks.Add` moves it | `make new bookmark` moves it |
| A name with a space | | -1708 (refused by the tool before it reaches Word) |
| Hidden bookmarks (`_...`) | Not in `Bookmarks` unless `ShowHidden`, so the tools leave them out by name | Not counted |
| Order of the bookmarks | By name unless asked; the tool sorts by place | The tool sorts by place |

## Floating shapes (Word 16, both platforms)

`scripts/probes/shapes_windows.py`, `scripts/probes/mac_shapes.py`,
`scripts/probes/mac_floating_pictures.py`, `scripts/probes/shape_kinds.py`,
`scripts/probes/shape_look.py` (a picture of the page to look at).

| What | Windows | Mac |
|---|---|---|
| The collection | Runs in z-order; `ZOrder` and leaving in-line change it. | The same. |
| Two shapes with one name | Allowed; `Shapes.Item(name)` answers the first. | The same. |
| Making one | `AddShape(type, left, top, w, h, Anchor)`, `AddTextbox`, `AddLine`, the same `Anchor` at the end, which the page omits | `make new shape at document with properties {..., anchor: range}`, likewise `text box` and `line shape`. Made `at` a range or paragraph: -1708. The answer can point at another shape (`text_boxes[4]`, `shapes[1]`): found by the name that was not there. |
| The numbers given when making | Taken as measured from the page, turned into offsets from column and paragraph | Shape and line keep the numbers as given, a text box is converted. So: make, then set what it is measured from, then the position. |
| Default place, wrap | Relative to column and paragraph, `wrap` 3 (in front) | The same |
| Changing what it is measured from | The shape stays where it is on the page, its numbers change | The same |
| An anchor | **Takes one position in the text that `Range.Text` does not show** (a paragraph of 22 characters spans 23 positions; by offset, one off). At the start of the paragraph for a paragraph's range; where given for a collapsed one. | The same. |
| Rewriting a paragraph | Not measured apart from the live check, which passes from past the anchor | The whole range takes the anchor and its shape with it, silently; from past the anchor, it stays. |
| Deleting its paragraph | The shape goes | The same |
| Wrap in line, a drawn shape | `WrapFormat.Type = 7` holds; stays in `Shapes` | The same. `ConvertToInlineShape` answers an object and leaves it in `Shapes`. |
| A line drawn backwards | `AddLine(300, 400, 50, 350)`: box at 50, 350 with both flips | Box at 300, 400 with both flips, wrong; setting left and top puts it right, the arrow still pointing the right way |
| Colors | BGR integer | 16-bit lists; 8-bit values read back as near black |
| Dash | `Line.DashStyle` | `dash style`, in appscript `dash_style_` |
| Rotation | `Rotation` | In appscript `rotation_` |
| Alternative text | Yes | `alternative text` is not a property of `shape`; read answers missing value, set -1708 |
| Vertical anchor, auto size, word wrap of a text frame | Yes | Unknown property or missing value |
| Vertical text | `TextFrame.Orientation = 4` | `vertical east asian`, code 4; upright kana, Latin on its side ([shape-look](probes/shape_look.py)) |
| Whether the text fits | `TextFrame.Overflowing` answered True for an empty shape and for two lines of 9 pt in a 60 pt high shape, so it is not used | `overflowing` agrees with the page; `has text` is true for an empty shape, so the text is read instead |
| A line's text frame | Touching `TextRange` is refused, even to read | Read answers nothing |
| Undo | One undo takes a new shape back (live check) | Two undos did not take a new shape back |
| Document not active | Works, active document unchanged | Works, active document unchanged |
| A table cell, a later page | Works; `LayoutInCell` -1 | Works |
| Kinds | 177 of 181 read back as themselves. `line callout two` reads as `one`; `line callout one no border` and `two no border` read as `one accent bar`; `callout one border and accent bar` reads as `line callout one`. `balloon` reads back as `rounded rectangular callout`, so it is taken as that name. | The same 177 of 181, the same four. |

### Pictures that float, and cropping (`scripts/probes/picture_crop.py`)

| What | Windows | Mac |
|---|---|---|
| Made floating directly | `Shapes.AddPicture` with the `Anchor` argument: anchored at 0, not where it was given | `make new picture`: likewise anchored at 0 |
| So | Inline first, at the paragraph's start, then `ConvertToShape`: anchored there | The same; `convert to shape` answers a reference to another shape, so found by the new name |
| Position after converting | Named: left and top (-999998, -999999) | -999998 and 0.75 |
| Alt text | Kept both ways | Set while in line; kept through the round trip, and cannot be read while it floats |
| `wrap` 7 on a floating picture | Becomes an InlineShape | The same; `convert to inline shape` too |
| A crop | Counts the picture's own size: 5 on a picture shown at twice its size takes 10 points off the frame, in line and floating (90 to 80) | The same (120 to 110) |
| A crop taken off | Width comes back 0.05 short | The same |
| Changing the wrap | Moves the shape about 0.1 pt (100 to 99.9) | Did not move it here |
| Brightness and contrast | `PictureFormat.Brightness`, `.Contrast`; live check | `brightness`, `contrast`, 0.5 by default; set in line, kept when it floats |

## Fields inside fields, and macros (Word 16, both platforms)

`scripts/probes/nested_fields.py`, and the safety steps of the live check.

| What | Windows | Mac |
|---|---|---|
| A field inside another's code | The outer code reads `QUOTE "base" \\x13 QUOTE "inner" \\x15`; `Fields` lists both | The same |
| Updating the outer one | Evaluates the inner one (`baseinner`) | The same |
| A field's name built by another field | The code starts with `\\x13` | Not tried |
| `\\x13 ... \\x15` written as text | Plain characters: no field | The same |
| `Fields.Update` over a locked field | Passes over it and answers 0 | Fields are updated one at a time |
| `AutomationSecurity` | Default low, by the page | Measured low; `automation security` reads and writes |
| Updating an INCLUDETEXT field | Not seen to ask | Word can stop on a dialog of its own ("contains fields that can share data with external files or websites"), Yes or No, and the call waits. Seen in one run of two, so the live check no longer makes such a field. |
| Whether a real macro stays off | Not measured: making a macro-enabled test document needs access to the VBA project | Not measured |


## Tracked changes and comments (Word 16, both platforms)

`scripts/probes/windows_comments_revisions.py` and
`scripts/probes/mac_comments_revisions.py`, measured 2026-11-09. Each runs only
while no other document is open.

| What | Windows | Mac |
|---|---|---|
| A paragraph with a tracked deletion, markup shown | `Range.Text` holds the deleted text | The same |
| The same with the markup hidden | The text Word shows; `Range.End` does not move | The same |
| Taking the deletions' spans out of the range | Equals the text with the markup hidden | The same |
| A replacement through a range | The new text goes in front of the text marked deleted | The same |
| A whole paragraph deleted | It stays, and `Paragraphs.Count` does not change | The same |
| A ruby word marked deleted | Its field stays in `Range.Fields` and its `\x15` in `Range.Text`; the code lies inside the deletion | The same, through `field code` |
| Accepting or rejecting one change | The ones after it renumber | The same |
| `For Each` over changes that include a format change | Two format changes both answered `Index` 1 | Not seen |
| `Range.Revisions.AcceptAll` | Works | Refused (-1708), so a change is accepted by number |
| Making a comment | `Comments.Add`, with `Replies.Add`, `Done`, `DeleteRecursively` | `make new Word comment` ended the Word process, both `at` a range and `at` the document with `scope` as a property |
| Opening a file Word did not make | Opens | Word asks for access in a dialog of its own, and the Apple event times out (-1712) until someone answers |

Because a search matches text marked deleted, a replacement run a second time
would replace it again. The search leaves those matches out.
