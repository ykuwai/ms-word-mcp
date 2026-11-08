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
