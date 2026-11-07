# Similar projects (surveyed 2026-11-08)

Ideas and lessons only; no code is taken from these projects. Star counts are
as seen in November 2026.

| Project | Approach | What it sells | What its users complain about |
|---|---|---|---|
| [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) (~2.1k, archived) | .docx files through python-docx | Breadth of tools, easy setup | Find/replace misses text split across runs ([#80](https://github.com/GongRzhe/Office-Word-MCP-Server/issues/80)); inserted text loses the surrounding formatting ([#84](https://github.com/GongRzhe/Office-Word-MCP-Server/issues/84)); no track changes ([#64](https://github.com/GongRzhe/Office-Word-MCP-Server/issues/64)); no LaTeX equations ([#69](https://github.com/GongRzhe/Office-Word-MCP-Server/issues/69)); "where is my file" ([#21](https://github.com/GongRzhe/Office-Word-MCP-Server/issues/21)); tool names over 60 characters ([#54](https://github.com/GongRzhe/Office-Word-MCP-Server/issues/54)) |
| [ykarapazar/word-mcp-live](https://github.com/ykarapazar/word-mcp-live) (259) | Fork of the above, live Word over COM and JXA | Edits the open document; tracked changes with an author; one Ctrl+Z per action; layout diagnosis; UnicodeMath equations; cross-references; snapshot and diff | Mac parity ([#12](https://github.com/ykarapazar/word-mcp-live/issues/12), [#22](https://github.com/ykarapazar/word-mcp-live/issues/22)); a replace that loops forever when the new text contains the old ([#16](https://github.com/ykarapazar/word-mcp-live/issues/16)); path traversal ([#15](https://github.com/ykarapazar/word-mcp-live/issues/15)); multilevel lists ([#13](https://github.com/ykarapazar/word-mcp-live/issues/13)); explicit null refused ([#26](https://github.com/ykarapazar/word-mcp-live/issues/26)) |
| [flyfish-dev/word-ai](https://github.com/flyfish-dev/word-ai) (10) | Office.js add-in plus a local bridge | Writes with preconditions (expected old text), dry run, audit and rollback | Low traction |
| [SecurityRonin/docx-mcp](https://github.com/SecurityRonin/docx-mcp) (67) | Edits the OOXML directly | Real redlines, comments, footnotes, validation | Comment replies Word hides ([#21](https://github.com/SecurityRonin/docx-mcp/issues/21)); nested revisions dropped ([#19](https://github.com/SecurityRonin/docx-mcp/issues/19)) |
| Microsoft Copilot in Word, Agent Mode | Built in | Built-in styles and heading hierarchy, TOC; in preview: track changes, comments, headers, fields | Tracked edits that rewrite whole paragraphs |
| [Claude for Word](https://www.claude.com/docs/office-agents/word) | Sidebar add-in | Edits inherit the surrounding style and numbering; tracked changes; replies in comment threads | — |

## What nobody does

- Japanese typography: vertical text, ruby, East Asian fonts per script,
  line-breaking rules (kinsoku).
- Captions with SEQ numbering, and LaTeX input for equations built into Word's
  own math (OMML). word-mcp-live takes UnicodeMath only.
- View control such as zoom.

## Ideas worth taking

1. Edits take the surrounding style and numbering unless told otherwise.
2. One tool call is one undo step (done), with an undo history.
3. Tracked changes with a configurable author, on both platforms.
4. Snapshot and diff: answer only the paragraphs changed since a snapshot.
5. Preconditions on writes: refuse if the paragraph no longer reads as
   expected, since the user may be editing it at the same moment.
6. Layout diagnosis and page images for the model to check its own work,
   including Japanese checks (half- and full-width mix, kinsoku).
7. Cross-reference targets listed, REF fields inserted, alongside SEQ captions.
8. Equations from LaTeX and UnicodeMath through Word's own BuildUp.
9. Replies in comment threads, anchored so Word shows them.

## Things to avoid

- Searching run by run: use Word's Find, which crosses runs.
- A replacement that can match its own output again.
- Unrestricted file paths.
- Tool names over 60 characters; refusing an explicit null.
- Rewriting a whole paragraph for a word-level change.
- Claims such as "the only"; tool counts typed by hand.
