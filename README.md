# ms-word-mcp

An MCP server that drives a live Microsoft Word on Windows and macOS, so an AI
can write and format a document while you work in it.

> **Status:** pre-release, in development on `develop`. Not on PyPI yet.

## What it is for

- **Working beside you.** Every write goes through a range, never through your
  cursor or selection, so you can keep typing while the AI writes elsewhere in
  the same document. The window scrolls to what it changed (`WORD_FOLLOW_EDITS`),
  and each tool call is one Undo in Word.
- **Read back after every write.** A result says what Word actually holds now,
  and lists as `warnings` whatever did not take.
- **Japanese typography.** Latin, East Asian and complex-script text each have
  their own font. Spacing in lines (行) and indents in characters (字), East
  Asian line breaking (禁則), hanging punctuation, emphasis dots (傍点), and
  numbering in full-width digits, ①, 一, ア or イ.
- **Styles first.** Every character and paragraph property can be set on a
  style as well as on text, so a document's look is defined once.
- **Equations.** From LaTeX or UnicodeMath, built by Word's own equation
  editor, numbered (1), (2)... at the right margin by fields that renumber
  themselves.
- **Official sources.** Every member and constant is checked against
  Microsoft's documentation ([docs/reference](docs/reference)), and what the
  documentation leaves unstated is measured on a real Word
  ([docs/PROBES.md](docs/PROBES.md)).

## Tools

The server lists three tools by default (catalog mode): `word_describe` lists
the rest a line each and gives their arguments, `word_run` runs one and
`word_run_steps` runs several in order. Set `WORD_TOOL_MODE=full` to list
every tool instead.

| Area | Tools |
|---|---|
| Word and documents | `word_get_app_info`, `word_list_documents`, `word_activate_document`, `word_open_document`, `word_create_document`, `word_save_document`, `word_save_document_as` (also PDF), `word_close_document` |
| Reading | `word_get_document_info`, `word_get_outline`, `word_get_paragraphs` |
| Writing | `word_insert_paragraphs`, `word_set_paragraph_text`, `word_delete_paragraphs`, `word_find_replace` |
| Formatting | `word_format_text`, `word_set_paragraph_format` |
| Styles | `word_apply_style`, `word_list_styles`, `word_get_style`, `word_set_style` |
| Lists | `word_apply_list`, `word_set_list_level`, `word_remove_list`, `word_get_list` |
| Equations | `word_insert_equation`, `word_list_equations`, `word_get_equation` |
| Fields and tables of contents | `word_insert_table_of_contents`, `word_insert_field`, `word_list_fields`, `word_update_fields` |
| View | `word_get_view`, `word_set_view` |
| Undo | `word_undo`, `word_redo` |

Every tool that works on a document takes an optional `document` argument (a
full path or file name), so a call lands in the document it names whatever
another conversation has chosen.

## Setting it up

Requirements: Microsoft Word (Microsoft 365 or 2021+) on Windows 10/11 or
macOS, and [uv](https://docs.astral.sh/uv/).

From source:

```bash
git clone https://github.com/ykuwai/ms-word-mcp.git
cd ms-word-mcp
uv sync
```

Claude Code:

```bash
claude mcp add word -- uv --directory /path/to/ms-word-mcp run ms-word-mcp
```

Other clients (Claude Desktop, Cursor, `.mcp.json`):

```json
{
  "mcpServers": {
    "word": {
      "command": "uv",
      "args": ["--directory", "/path/to/ms-word-mcp", "run", "ms-word-mcp"]
    }
  }
}
```

## Settings

| Variable | Default | What it does |
|---|---|---|
| `WORD_TOOL_MODE` | `catalog` | `catalog` lists three tools that reach the rest; `full` lists every tool. |
| `WORD_FOLLOW_EDITS` | `on` | Scroll the document's window to each change. With a second window open on the document (View > New Window), the last one follows and the first stays where you are. |

## On macOS

- The first call asks for permission to control Microsoft Word (System
  Settings > Privacy & Security > Automation).
- Word for Mac is sandboxed: it writes without asking only inside
  `~/Library/Containers/com.microsoft.Word/Data/Documents`; saving elsewhere
  may make Word ask for access.
- Each change is its own Undo step (Word for Mac's scripting cannot group
  them), and a style can be changed but not created. Tools say so in their
  results.

## Development

```bash
uv run pytest -q                    # no Word needed
uv run python scripts/live_check.py # runs every tool on a real Word, in a scratch document
```

The live check creates a document, runs the tools on it, prints PASS or FAIL
for each step, exports a PDF and closes the document. It touches no other
document.

## Credits

Built with [ppt-mcp](https://github.com/ykuwai/ppt-mcp)'s design. AlonSoko
measured Word on both platforms for his own Word server and shared the
measurements and code with this project ([ykuwai/ppt-mcp#352](https://github.com/ykuwai/ppt-mcp/issues/352));
what this project took from them is noted in [docs/PROBES.md](docs/PROBES.md).

## License

MIT
