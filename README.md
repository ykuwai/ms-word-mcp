<h1 align="center">
  <img src="https://raw.githubusercontent.com/ykuwai/ms-word-mcp/main/assets/word-mcp-logo-letter.svg" alt="Word MCP (ms-word-mcp)" width="400">
</h1>

<p align="center">
  <strong>English (en)</strong> · <a href="https://github.com/ykuwai/ms-word-mcp/blob/main/i18n/README.ja.md">日本語 (ja)</a>
</p>

<p align="center">
  <a href="https://pypi.org/project/ms-word-mcp/"><img src="https://img.shields.io/pypi/v/ms-word-mcp.svg" alt="PyPI"></a>
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue.svg" alt="Python"></a>
  <a href="https://github.com/ykuwai/ms-word-mcp/blob/main/LICENSE"><img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License"></a>
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-0078d4.svg" alt="Platform">
</p>

<p align="center">
  <strong>Real-time Microsoft Word control on Windows and macOS —<br>an MCP server with <!--tools-->84<!--/tools--> tools that works in the document beside you.</strong>
</p>

---

An MCP (Model Context Protocol) server that gives AI agents control of a live Microsoft Word, over COM on Windows and Apple Events on macOS. Unlike file-based libraries such as python-docx, it works in the Word you have open, and you can keep typing while it writes elsewhere in the same document.

> **Status:** early release (0.x). Tool names and arguments may still change.

## ✨ Key Features

- **Working beside you** — Every write goes through a range, never your cursor or selection. The window scrolls to what changed
- **Read back after every write** — A result says what Word actually holds now, and lists under `warnings` whatever did not take
- **Every script** — Latin, East Asian (Chinese, Japanese, Korean) and complex-script (Arabic, Hebrew...) text each have their own font; grid spacing, ruby, emphasis marks, vertical text and full-width numbering are built in
- **Styles first** — Every character and paragraph property can be set on a style as well as on text, so a document's look is defined once
- **Equations** — From LaTeX or UnicodeMath, built by Word's own equation editor, numbered (1), (2)... at the right margin by fields that renumber themselves
- **Sees the page** — `word_get_page_image` gives the model a picture of a page as Word lays it out, without capturing your screen
- **Safe by default** — Macros stay off while a document opens, and fields that reach outside the document are refused unless you allow them
- **Windows and macOS** — The same tools, arguments and answers; where Word for Mac cannot do something, the tool says so

## 📋 Requirements

- Windows 10 or 11, or macOS
- Microsoft Word (Microsoft 365 or 2021 and later)
- [uv](https://docs.astral.sh/uv/getting-started/installation/)

On macOS the first tool call raises the system's automation prompt, and Word has to be allowed there once. See [macOS support](#-macos-support) for what differs.

## 🚀 Getting Started

Standard config — works in Claude Desktop, Cursor, `.mcp.json`, and most other MCP clients:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["ms-word-mcp"]
    }
  }
}
```

### Claude Code

**User-scoped** (available in all projects)
```bash
claude mcp add --scope user word uvx ms-word-mcp
```

**Project-scoped** (stored in `.mcp.json`, shared with your team)
```bash
claude mcp add --scope project word uvx ms-word-mcp
```

### Cursor

[![Install in Cursor](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=ms-word-mcp&config=eyJ0eXBlIjoic3RkaW8iLCJjb21tYW5kIjoidXZ4IiwiYXJncyI6WyJtcy13b3JkLW1jcCJdfQ==)

Or add manually to `~/.cursor/mcp.json` using the standard config above.

### Claude Desktop

Edit `claude_desktop_config.json` using the standard config above. It is at `%APPDATA%\Claude\` on Windows and `~/Library/Application Support/Claude/` on macOS.

### Codex

Edit `~/.codex/config.toml`:

```toml
[mcp_servers.ms-word-mcp]
command = "uvx"
args = ["ms-word-mcp"]
```

### VS Code

```bash
code --add-mcp '{"name":"word","command":"uvx","args":["ms-word-mcp"]}'
```

### From source

```bash
git clone https://github.com/ykuwai/ms-word-mcp.git
cd ms-word-mcp
uv sync
```

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

## 🛠️ Tool Categories

Documents, reading, writing, formatting, styles, lists, tables, pictures, captions and cross-references, the page, equations, floating shapes and pictures, footnotes and endnotes, hyperlinks and bookmarks, fields and tables of contents, the view, a picture of a page, and undo.

The full list of tools, by category, is in [TOOLS.md](https://github.com/ykuwai/ms-word-mcp/blob/main/TOOLS.md). Comments are not here yet.

## 💡 Example Prompts

Just describe what you want in plain language — no code required.

**Simple** — topic only

> *"Write a one-page memo announcing the move to a four-day work week."*

**With a look** — topic and style

> *"Write a two-page proposal for a new office library. Serif body text, navy headings, and a table of the costs."*

**Fixing what exists** — an open document

> *"Open report.docx. The headings are only bold text and the bullets are typed by hand: turn them into real headings and lists, then add a table of contents."*

**Japanese layout**

> *"Set this notice in vertical Japanese on A4, 40 characters by 30 lines, with ruby on the difficult names."*

**Mathematics**

> *"Add the normal distribution's density as a numbered equation after the second paragraph, and refer to it by its number in the text."*

**Keywords that help:**

| Aspect | Keywords / Phrases | What it does |
|---|---|---|
| **Structure** | `real headings` · `table of contents` · `numbered sections` | Uses Word's styles and fields, so navigation and page numbers work |
| **Look** | `serif body` · `navy headings` · `Yu Mincho for Japanese` | Defines the look on styles, once for the document |
| **Page** | `A4` · `two columns` · `40 characters by 30 lines` · `page numbers at the foot` | Sets paper, margins, columns and the document grid |
| **Tables** | `table of the costs` · `merge the header` · `shade the first row` | Builds the table with all its cells at once, then formats it |
| **Pictures and shapes** | `caption each figure` · `wrap the text around it` · `behind the text` | Adds captions and cross-references, or floats pictures and shapes |
| **Checking** | `check how each page looks` | Has the model look at the page images and correct the layout |

## 🔍 Features in Detail

### 🎯 Document Targeting

`word_activate_document` chooses the document every later call works on, whichever window is in front. The choice is one value for the whole server, and a client such as Claude Desktop routes every conversation through one server, so pass `document` on each call when several conversations work at once. It applies to that call only and takes a full path or a file name. A document that is not open is an error, never a fallback to another one.

### ✍️ Working Beside You

Writes go through ranges, so your cursor and selection stay where they are. Each write scrolls the document's window to what it changed; with a second window on the document (View > New Window), the last one follows and the first stays where you are. To keep every window still, set `WORD_FOLLOW_EDITS=off`.

Each tool call is one Undo step in Word on Windows when the document is the active one; otherwise the result says so under `undo`.

### ✅ Answers You Can Trust

Each write reads back what Word now holds. A result with `warnings` landed only in part, and the warnings say what did not. A result with `error` changed nothing unless it says otherwise.

Every member and constant is checked against Microsoft's documentation ([docs/reference](https://github.com/ykuwai/ms-word-mcp/tree/main/docs/reference)), and what the documentation leaves unstated is measured on a real Word ([docs/PROBES.md](https://github.com/ykuwai/ms-word-mcp/blob/main/docs/PROBES.md)).

### 🌏 Every Script

- **Three fonts on all text** — Latin, East Asian and complex script are set separately
- **East Asian layout** — spacing in grid lines, indents in character widths, line breaking rules, hanging punctuation, emphasis marks, vertical text and ruby (furigana) with `word_add_ruby`
- **List numbering** — full-width digits, ①, 一, ア, 가, א, α, а and more

### 🎨 Styles

`word_set_style` defines any character and paragraph property on a style; `word_format_text` and `word_set_paragraph_format` take the same properties for the exceptions. The English built-in names ("Normal", "Heading 1", "TOC 1"...) work in every language of Word.

### 🧮 Equations

`word_insert_equation` takes LaTeX or UnicodeMath and has Word's own equation editor build it. With `numbered`, the number sits at the right margin and is a field, so inserting an equation renumbers the rest.

### 🖼️ Shapes and Floating Pictures

`word_insert_shape`, `word_insert_text_box`, `word_insert_line` and `word_insert_floating_picture` add objects with text wrapping, position, layer order, fill, line and text. `word_set_shape` changes them afterwards, crops and fades pictures, and moves a picture between floating and in line.

### 📄 Seeing a Page

`word_get_page_image` returns a page as Word lays it out, so the model can check a layout that reading the text back cannot show. On Windows the page comes from Word directly. On macOS Word saves a PDF for a moment and the page is drawn from it. Neither captures the screen.

## ⚙️ Advanced Configuration

| Variable | Default | What it does |
|---|---|---|
| `WORD_TOOL_MODE` | `catalog` | `catalog` lists three tools that reach the rest; `full` lists every tool. |
| `WORD_FOLLOW_EDITS` | `on` | Scroll the document's window to each change. |
| `WORD_ALLOW_ACTIVE_FIELDS` | `off` | Allow fields that reach outside the document: INCLUDETEXT, INCLUDEPICTURE, INCLUDE, LINK, IMPORT, RD, EMBED, DATABASE, DDE, DDEAUTO, MACROBUTTON, PRINT, ASK, FILLIN. Off, `word_insert_field` refuses them and `word_update_fields` leaves the ones already in a document as they are. On, Word may stop and ask you about them in a dialog of its own, and the tool waits until you answer. |
| `WORD_ALLOW_NETWORK_PATHS` | `off` | Allow files named by a network path (`\\server\share\...`). Off, they are refused: saving to one sends the document to that server, and even checking for the file connects to it with your credentials. |
| `WORD_MAC_FILES_DIR` | the document's folder | macOS: where Word writes the PDF a page picture is drawn from, for a moment. By default the document's own folder, or Word's own folder for a document never saved. |

Set them in the server's `env`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["ms-word-mcp"],
      "env": { "WORD_FOLLOW_EDITS": "off" }
    }
  }
}
```

### The Tool List: Catalog Mode

A client sends the description of every tool it lists with every request, and the descriptions of these tools come to about 180,000 characters. By default the server lists three tools instead, about 2,500 characters, and the model reaches the rest through them:

| Tool | What it does |
|---|---|
| `word_describe` | With no names, the list of all the tools, one line each. With names, their full description and arguments. |
| `word_run` | Runs one tool: `{"tool": "word_insert_paragraphs", "params": {...}}`. |
| `word_run_steps` | Runs several in order, in one call. A step can use what an earlier one answered: `"$prev.key"` is the previous step's `key`, `"$2.inserted.0.index"` the second step's. The first step that fails ends the run and says why. |

The tools themselves are the same. What catalog mode gives up is per-tool control in the client: approving or blocking a tool by its name applies to the three tools above rather than to each one. A client that relies on that sets `WORD_TOOL_MODE=full`.

### When Word Is Busy

Calls reach Word one at a time, in the order they were asked for. When Word has a dialog or menu open, a call waits up to 15 seconds for it to close and then answers that Word is busy. A call that is still running after 45 seconds on Windows or 69 on macOS is given up, and the server carries on with the next one. Tools that may take longer, a large find and replace for instance, stop in time and answer `complete: false` with a warning, so they can be called again.

If Word is not running, `word_create_document` and `word_open_document` start it; other tools answer that it is not running.

### 🔒 Security

The server runs on your computer with your permissions and talks to one local client over stdio; it opens no port and makes no network request of its own.

An AI using these tools can read and change any document open in Word, open any file Word can open, and save or embed files wherever your account can write or read. Text inside a document, a web page or another tool's answer can try to steer the AI, so check what it does with documents you received from someone else, and approve its tool calls where that matters.

What the server does about it:

- Macros are disabled while it opens a document or creates one from a template. Word would otherwise run them. The document stays open with macros off.
- Fields that reach outside the document are refused and left alone unless you set `WORD_ALLOW_ACTIVE_FIELDS`.
- Network paths are refused unless you set `WORD_ALLOW_NETWORK_PATHS`.
- A new link goes to http, https, mailto or tel, or to a path relative to the document. A link the document already has is left as it is.

What it does not do: a document opened through these tools does not get Word's Protected View, so open only documents you would open yourself.

To report a vulnerability, see [SECURITY.md](https://github.com/ykuwai/ms-word-mcp/blob/main/SECURITY.md).

## 🍎 macOS Support

The same tools, the same arguments, the same answers. What differs is what Word for Mac's scripting can reach, and where it cannot reach something the tool says so.

Not available on macOS:

- `word_insert_table_columns`: Word for Mac's scripting cannot add columns
- Paragraph direction (right to left) and tab stops
- The size of complex-script text and the color of an underline
- Alt text on a floating shape, and where text sits from top to bottom in one

Smaller differences:

- Each change is its own Undo step, because Word for Mac's scripting cannot group them, and a style can be changed but not created. Tools say so in their results.
- Word for Mac is sandboxed: it writes without asking only inside `~/Library/Containers/com.microsoft.Word/Data/Documents`, and saving elsewhere may make Word ask for access.
- Word asks you about each file outside its own folder, so `word_insert_picture` puts a copy in Word's folder for a moment and reads it there. The first time, macOS may ask once whether the app the server runs in may use Word's data.
- `word_get_page_image` has Word save a PDF next to the document for a moment. For a document never saved, the PDF goes in Word's own folder, and macOS may ask the same question.
- Word for Mac makes a table only in the active document, so inserting a table brings its document forward for a moment and then gives the previously active one back.

## 🤝 Contributing

Pull requests are welcome. Base them on `develop`, one topic each, with tests.

```bash
uv run pytest -q                    # no Word needed
uv run python scripts/live_check.py # runs every tool on a real Word, in a scratch document
```

The live check creates a document, runs the tools on it, prints PASS or FAIL for each step, exports a PDF and closes the document. It touches no other document.

On Windows, while a client has the server running, `uv sync` cannot replace `.venv\Scripts\ms-word-mcp.exe`; run the tests with `uv run --no-sync`, or stop the server first.

## 📄 License

MIT

## 🙏 Credits

- [ppt-mcp](https://github.com/ykuwai/ppt-mcp) — the design this server is built on
- [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk) — MCP server framework
- [pywin32](https://github.com/mhammond/pywin32) — Windows COM automation
- [appscript](https://github.com/hhas/appscript) — macOS Apple Event bridge
- [Model Context Protocol](https://modelcontextprotocol.io/) — by Anthropic
