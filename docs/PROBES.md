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
| Reading text | `paragraphs[i].text_object.content.get()` answers the paragraph's text without its mark. |
| Latin font | `font_object.ascii_name` and `other_name` set and read back. |
| East Asian font | `east_asian_name` set to "Hiragino Sans" took. Set to "Arial" (no East Asian glyphs) it is refused with -1708 and keeps the old value. Windows is expected to keep the old value silently instead; to be measured. |
| Complex-script font | `complex_script_name` reads back as a theme reference, `+本文のフォント - コンプレックス スクリプト ()`, until a font is set. |
| Line spacing | A new paragraph reads `line_spacing_rule` = `line_space_single`, `line_spacing` = 12.0. |
| Style | `paragraphs[1].style.get()` answers a `Word_styles['標準']` reference; `style.name_local.get()` answers `標準` in a Japanese Word. |
| End of the text | `text_object.end_of_content` answered 28 for 27 characters of text (the last paragraph mark counts). |
| Closing | `close(saving=k.no)` closes without asking. |
| New document name | `create_new_document()` names it `文書 1`, `文書 3`: the number is not reused. |

## Word for Windows

Not measured yet. `scripts/live_check.py` runs every tool against a scratch
document and prints PASS or FAIL for each step.
