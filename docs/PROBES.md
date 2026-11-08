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

## From the reference branch

AlonSoko measured Word 16 on both platforms before writing his own version, and
pushed it with the measurements to the `reference` branch of this repository
(`docs/PROBES.md` there, with the probe scripts and their output under
`probes/`). What bears on this server, in short:

- **Reading one paragraph at a time is slow.** 1,000 paragraphs with their
  outline level took 20 s through COM and 15 s through Apple Events. The whole
  text comes in one call (13 ms), the headings in one call
  (`GetCrossReferenceItems`), and styles with the document's XML
  (`Content.WordOpenXML`, 84 ms). word_get_outline and word_get_paragraphs
  here still walk paragraph by paragraph; see ROADMAP.
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
