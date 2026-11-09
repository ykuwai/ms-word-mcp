"""ms-word-mcp: an MCP server that drives a live Microsoft Word.

Word on Windows over COM, and Word for Mac over Apple Events. The tools write
through ranges, never through the user's selection, and scroll the window to
what they changed, so the user can watch the work land while they write
elsewhere in the same document.
"""

from __future__ import annotations

import difflib
import json
import logging
import os
import sys
from contextlib import asynccontextmanager

# mcp 2.0 renamed mcp.server.fastmcp to mcp.server.mcpserver, and FastMCP to
# MCPServer, without a compatibility shim. ModuleNotFoundError rather than
# ImportError: a broken 2.x install must show its own traceback.
try:
    from mcp.server.mcpserver import MCPServer  # mcp >= 2.0
    from mcp.server.mcpserver.exceptions import ToolError
except ModuleNotFoundError:
    from mcp.server.fastmcp import FastMCP as MCPServer  # mcp 1.x
    from mcp.server.fastmcp.exceptions import ToolError
from pydantic import ValidationError

from . import backend
from .backend import call_document

logging.basicConfig(level=logging.INFO, stream=sys.stderr,
                    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s")
logger = logging.getLogger("ms-word-mcp")


@asynccontextmanager
async def app_lifespan(server):
    # Word is not started here: connecting happens on the first tool call, so
    # a client that never uses these tools never starts Word.
    backend.word.start()
    try:
        yield {}
    finally:
        backend.word.stop()


# The argument every document-bound tool takes next to `params`, and the tools
# that act on Word itself or choose the document in the first place.
DOCUMENT_ARG = "document"
TOOLS_WITHOUT_DOCUMENT_ARG = frozenset({
    "word_get_app_info", "word_list_documents", "word_activate_document",
    "word_open_document", "word_create_document", "word_describe",
})
DOCUMENT_PASSING_TOOLS = frozenset({"word_run", "word_run_steps"})


class WordServer(MCPServer):
    """MCPServer with a per-call `document` argument and honest failures.

    The session target set by word_activate_document is one value for the
    process, and one process can serve several conversations: a call that
    names its document is immune to another conversation switching it.
    `document` is taken out here, before the SDK validates `params`, and
    held in backend.call_document while the tool runs.

    A key next to `params` that the tool does not take is refused here (the
    SDK ignores it), and a tool's answer with a top-level "error" is marked
    isError true with its text left as the tool wrote it.
    """

    async def call_tool(self, name, arguments, *args, **kwargs):
        wanted = None
        # word_run and word_run_steps take `document` as their own argument
        # and pass it on to the tool they run, so it is left in place for
        # them. word_describe has none, and is dropped like any tool's that
        # does not work on a document: a client told to pass `document` on
        # every call passes it there too (ppt-mcp #367).
        if (name not in DOCUMENT_PASSING_TOOLS and isinstance(arguments, dict)
                and DOCUMENT_ARG in arguments):
            arguments = dict(arguments)
            wanted = arguments.pop(DOCUMENT_ARG)
            if wanted is not None and not isinstance(wanted, str):
                raise ToolError(f"'{DOCUMENT_ARG}' must be a string (a file name or "
                                f"full path), got {type(wanted).__name__}.")
            if name in TOOLS_WITHOUT_DOCUMENT_ARG:
                wanted = None
            wanted = (wanted or "").strip() or None
        tool = self.registered_tool(name)
        if tool is not None and isinstance(arguments, dict):
            refuse_unknown_arguments(name, tool, arguments)
        token = call_document.set(wanted)
        try:
            return flag_failure(await super().call_tool(name, arguments, *args, **kwargs))
        except ToolError as exc:
            readable = readable_validation_error(name, exc)
            if readable is None:
                raise
            raise ToolError(readable) from exc.__cause__
        finally:
            call_document.reset(token)

    async def list_tools(self, *args, **kwargs):
        tools = await super().list_tools(*args, **kwargs)
        if CATALOG_MODE:
            return [t for t in tools if t.name in CATALOG_TOOLS]
        return [t for t in tools if t.name not in CATALOG_TOOLS]

    def registered_tool(self, name):
        """The SDK's Tool for `name`, or None. _tool_manager is an SDK internal."""
        try:
            return self._tool_manager.get_tool(name)
        except Exception:  # noqa: BLE001
            return None

    def all_tools(self):
        return self._tool_manager.list_tools()


def failure_payload(text):
    """The JSON object `text` holds if it is a failure (a top-level "error")."""
    if not text or not text.lstrip().startswith("{"):
        return None
    try:
        payload = json.loads(text)
    except ValueError:
        return None
    return payload if isinstance(payload, dict) and "error" in payload else None


def text_of(blocks):
    return "".join(getattr(b, "text", "") or "" for b in blocks
                   if getattr(b, "type", "") == "text")


def _sdk_passes_call_tool_result():
    """mcp 1.19 was the first to send a CallToolResult from call_tool as it is."""
    from importlib.metadata import PackageNotFoundError, version

    try:
        major, minor = (int(part) for part in version("mcp").split(".")[:2])
    except (PackageNotFoundError, ValueError):
        return True
    return (major, minor) >= (1, 19)


SDK_PASSES_CALL_TOOL_RESULT = _sdk_passes_call_tool_result()


def flag_failure(result):
    """`result` as a tool error (isError true) if the tool answered a failure."""
    if hasattr(result, "content") and not isinstance(result, (list, tuple, dict)):
        if getattr(result, "is_error", False) or failure_payload(text_of(result.content)) is None:
            return result
        return result.model_copy(update={"is_error": True, "structured_content": None})
    blocks = result[0] if isinstance(result, tuple) else result
    if (not SDK_PASSES_CALL_TOOL_RESULT or isinstance(result, dict)
            or failure_payload(text_of(blocks or [])) is None):
        return result
    from mcp.types import CallToolResult

    return CallToolResult(content=list(blocks), isError=True)


def refuse_unknown_arguments(name, tool, arguments):
    declared = list((tool.parameters or {}).get("properties") or {})
    unknown = [key for key, value in arguments.items()
               if key not in declared and not (key == "params" and value in (None, {}))]
    if not unknown:
        return
    if name not in TOOLS_WITHOUT_DOCUMENT_ARG and DOCUMENT_ARG not in declared:
        declared.append(DOCUMENT_ARG)
    parts = []
    for key in unknown:
        close = difflib.get_close_matches(key, declared, n=1, cutoff=0.6)
        parts.append(f"'{key}'" + (f" (did you mean '{close[0]}'?)" if close else ""))
    raise ToolError(
        f"{name} does not take {', '.join(parts)} next to its arguments. It takes: "
        f"{', '.join(declared) or 'nothing'}. Tool arguments go inside `params`. "
        "Nothing was applied.")


def readable_validation_error(name, exc):
    """One sentence per problem for a ToolError caused by argument validation."""
    cause = exc.__cause__
    if not isinstance(cause, ValidationError):
        return None
    problems = []
    for err in cause.errors():
        loc = ".".join(str(part) for part in err.get("loc", ()))
        message = err.get("msg", "invalid")
        if err.get("type") == "extra_forbidden":
            message = "is not an argument of this tool"
        problems.append(f"{loc}: {message}" if loc else message)
    return (f"{name} was called with arguments it cannot take: " + "; ".join(problems)
            + ". Nothing was applied. word_describe gives its arguments.")


INSTRUCTIONS = """
## Working in Word beside the user

The user may be writing in the same document while you work. These tools
write through ranges, never through the user's cursor or selection, and
scroll the window to what they changed so the user can follow. Do not ask the
user to stop typing.

## A good order of work

1. Find the document: `word_list_documents`, then `word_activate_document`,
   or make one with `word_create_document` (pass `save_as` to save it at
   once) or open one with `word_open_document`. The chosen document is
   shared by every conversation this server serves, so also pass its
   `full_name` as `document` on every further call.
2. Read before writing: `word_get_outline` for the headings, then
   `word_get_paragraphs` for the parts you need.
3. Define the look with styles before writing much text (the fonts of a
   document are in Normal, the heading styles, Title and TOC 1-9 separately:
   set each one to change the fonts throughout): `word_set_style` on
   "Normal" for the body font (Latin and East Asian separately), and on
   "Heading 1", "Heading 2" for headings. The English built-in names work in
   every language of Word ("Normal", "Heading 1", "Title", "Body Text",
   "TOC 1"...). Most styles are based on Normal and take what Normal does not
   override: an indent or spacing for body text belongs on a body style such
   as "Body Text", or headings, the title and the table of contents move too. Then write with `word_insert_paragraphs`,
   giving each paragraph its style. Formatting text directly
   (`word_format_text`, `word_set_paragraph_format`) is for the exceptions.
4. Lists (`word_apply_list`), tables (`word_insert_table` with all the
   cells at once), footnotes and endnotes (`word_insert_note`), links
   (`word_insert_hyperlink`), equations (`word_insert_equation`, LaTeX or UnicodeMath,
   `numbered` for (1), (2)...), fields and tables of contents.
   Insert a table of contents last, since it adds paragraphs above the rest.
5. Look at the result: `word_get_page_image` shows a page as Word lays it
   out. Use it after layout changes, equations and lists; reading the text
   back cannot show a layout that broke.
6. Save with `word_save_document`.

## Fixing an existing document

Read it first (`word_get_paragraphs` with `detail` shows each paragraph's
style, fonts, bold and color). Paragraphs made to look like headings by
formatting alone: `word_apply_style` with a heading style and
`clear_direct_formatting`. Typed bullets ("・", "-", "•"): take the typed
mark out of the text (`word_find_replace`), then `word_apply_list`. Insert a
table of contents last; a paragraph that holds its end cannot be deleted
while it is there.

## Paragraph numbers

Paragraphs are numbered from 1, as Word counts them. Inserting or deleting
moves the paragraphs after it: every writing call answers the indices it
used (`inserted`, `index`, `paragraphs`), so take the next index from the
last answer rather than counting by hand. Several calls in one
`word_run_steps` can pass results on ("$1.inserted.0.index").

A table's cells are paragraphs too, and each row ends in an empty one (the
row's end mark): a table result gives the paragraphs it covers
(`paragraphs.first`, `paragraphs.last`). Work on a table's text with the
table tools, by row and column.

## Answers

A result with `warnings` landed only in part; the warnings say what did not.
A result with `error` changed nothing unless it says otherwise. Each writing
call is one undo step in Word on Windows when the document is Word's active
one; otherwise the result says so in `undo`.

## Scripts and languages

Word keeps three fonts on all text: Latin, East Asian (Chinese, Japanese,
Korean) and complex script (Arabic, Hebrew and other right-to-left text).
Setting one leaves the others; set the ones you mean. An East Asian font
must have East Asian glyphs: Word refuses Segoe UI or Arial there, and the
warning says so. bold, italic and size cover right-to-left text too unless
its own value is given. A right-to-left paragraph also needs
`direction: "right_to_left"` (Windows). A document from a Japanese or
Chinese template justifies Normal and fits it to a grid, which stretches
Latin text: give Normal `alignment: "left"` and `snap_to_grid: false` for a
mainly Latin document. East Asian layouts measure spacing in grid lines
(`space_before_lines`) and indents in character widths
(`first_line_indent_chars`); both are there as well as points.
"""

mcp = WordServer("ms_word_mcp", lifespan=app_lifespan, instructions=INSTRUCTIONS)

from .tools import (  # noqa: E402,F401 - mac registers the Mac halves
    documents, edit, equations, fields, formatting, links, lists, mac, notes, page, references, screen,
    tables, text, view)

for module in (documents, text, formatting, lists, tables, references, notes, links, equations,
               fields, page, view, screen, edit):
    module.register(mcp)


# -- tool schemas --------------------------------------------------------------

def inline_schema_defs(server) -> int:
    """Inline `$defs` references: MCP clients do not dereference `$ref`."""

    def resolve(node, defs, stack):
        if isinstance(node, list):
            return [resolve(item, defs, stack) for item in node]
        if not isinstance(node, dict):
            return node
        ref = node.get("$ref")
        if isinstance(ref, str) and ref.startswith("#/$defs/"):
            target_name = ref[len("#/$defs/"):]
            target = defs.get(target_name)
            if target is not None and target_name not in stack:
                merged = dict(target)
                merged.update({k: v for k, v in node.items() if k != "$ref"})
                return resolve(merged, defs, stack + [target_name])
        return {key: resolve(value, defs, stack) for key, value in node.items()}

    rewritten = 0
    for tool in server.all_tools():
        schema = tool.parameters
        if not isinstance(schema, dict) or "$defs" not in schema:
            continue
        defs = schema["$defs"]
        inlined = resolve({k: v for k, v in schema.items() if k != "$defs"}, defs, [])
        if "$ref" in json.dumps(inlined):
            inlined["$defs"] = defs
        tool.parameters = inlined
        rewritten += 1
    return rewritten


def slim(node):
    """A schema without the titles and null branches pydantic adds and no model reads."""
    if isinstance(node, list):
        return [slim(item) for item in node]
    if not isinstance(node, dict):
        return node
    out = {}
    for key, value in node.items():
        if key == "title" and isinstance(value, str):
            continue
        if key in ("properties", "$defs") and isinstance(value, dict):
            out[key] = {field: slim(sub) for field, sub in value.items()}
        else:
            out[key] = slim(value)
    branches = out.get("anyOf")
    if (isinstance(branches, list) and "default" in out and out["default"] is None
            and {"type": "null"} in branches):
        rest = [b for b in branches if b != {"type": "null"}]
        if len(rest) == 1 and isinstance(rest[0], dict) and not set(rest[0]) & set(out):
            out.pop("anyOf")
            out = {**rest[0], **out}
    return out


def add_document_arg(server) -> int:
    added = 0
    for tool in server.all_tools():
        if tool.name in TOOLS_WITHOUT_DOCUMENT_ARG or tool.name in CATALOG_TOOLS:
            continue
        schema = dict(tool.parameters)
        properties = dict(schema.get("properties") or {})
        properties[DOCUMENT_ARG] = {"type": "string"}
        schema["properties"] = properties
        tool.parameters = schema
        added += 1
    return added


# -- catalog mode, the default; WORD_TOOL_MODE=full lists every tool ----------------
# A client sends every listed tool's schema with every request. In catalog mode
# the server lists three tools instead: one that lists the rest, a line each,
# and gives full schemas on request; one that runs a tool; and one that runs a
# list of calls. The tools themselves run through the same call_tool.

def tool_mode(value: str) -> str:
    value = value.strip().lower()
    if value in ("catalog", ""):
        return "catalog"
    if value in ("full", "all"):
        return "full"
    logger.warning("WORD_TOOL_MODE=%r is not a setting; use 'catalog' (the default) "
                   "or 'full'. Using 'full'.", value)
    return "full"


CATALOG_MODE = tool_mode(os.getenv("WORD_TOOL_MODE", "catalog")) == "catalog"
CATALOG_TOOLS = frozenset({"word_describe", "word_run", "word_run_steps"})


def summary(description):
    text = " ".join((description or "").split())
    for end in (". ", "! ", "? "):
        if end in text:
            text = text.split(end, 1)[0] + end.strip()
            break
    return text[:120]


AREAS = {"documents": "Documents", "text": "Reading and writing text",
         "formatting": "Formatting and styles", "lists": "Lists", "tables": "Tables",
         "references": "Pictures, captions and cross-references",
         "notes": "Footnotes and endnotes", "links": "Hyperlinks and bookmarks",
         "page": "The page: setup, vertical text, headers, footers, ruby",
         "equations": "Equations", "fields": "Fields and tables of contents",
         "view": "View", "screen": "Seeing a page", "edit": "Undo"}


def catalog_lines():
    """One line per tool, under a heading for each area, in the order of work."""
    lines, area = [], None
    for tool in mcp.all_tools():
        if tool.name in CATALOG_TOOLS:
            continue
        module = getattr(getattr(tool, "fn", None), "__module__", "") or ""
        here = AREAS.get(module.rsplit(".", 1)[-1])
        if here and here != area:
            lines.append(f"\n# {here}")
            area = here
        lines.append(f"{tool.name}{_argument_names(tool)}: {summary(tool.description)}")
    return "\n".join(lines).strip()


def _argument_names(tool) -> str:
    """"(a, b, c)" for the tool's arguments, so the list alone shows what each
    takes (an agent called word_get_list with `paragraph` for `index`)."""
    params = ((tool.parameters or {}).get("properties") or {}).get("params") or {}
    names = list((params.get("properties") or {}).keys())
    if not names:
        return "()"
    shown = names[:6]
    return "(" + ", ".join(shown) + (", …" if len(names) > 6 else "") + ")"


def content_of(result):
    """What call_tool answered, as a list of content blocks under either SDK major."""
    if hasattr(result, "content") and not isinstance(result, (list, tuple, dict)):
        if getattr(result, "isError", False) or getattr(result, "is_error", False):
            text = text_of(result.content)
            if failure_payload(text) is None:
                raise ToolError(text or "The tool failed.")
        result = result.content
    if isinstance(result, tuple):
        result = result[0]
    if isinstance(result, dict):
        return [json.dumps(result, ensure_ascii=False)]
    return list(result or [])


async def run_one(tool, params, document):
    if not isinstance(tool, str) or not tool:
        raise ToolError("A step needs a 'tool': the name of the tool to run, such as "
                        "\"word_insert_paragraphs\".")
    target = mcp.registered_tool(tool)
    if target is None or tool in CATALOG_TOOLS:
        names = [t.name for t in mcp.all_tools() if t.name not in CATALOG_TOOLS]
        close = difflib.get_close_matches(tool, names, n=3)
        hint = f" Did you mean {', '.join(close)}?" if close else ""
        raise ToolError(f"No tool called '{tool}'.{hint} word_describe() lists them.")
    arguments = {}
    declared = (target.parameters or {}).get("properties") or {}
    if "params" in declared:
        arguments["params"] = params or {}
    elif params:
        raise ToolError(f"{tool} takes no params.")
    if document is not None:
        arguments[DOCUMENT_ARG] = document
    return content_of(await mcp.call_tool(tool, arguments))


def resolve_refs(value, results):
    """Replace "$prev.key" and "$N.key" (N counts steps from 1) with that step's answer."""
    if isinstance(value, dict):
        return {k: resolve_refs(v, results) for k, v in value.items()}
    if isinstance(value, list):
        return [resolve_refs(v, results) for v in value]
    if not (isinstance(value, str) and value.startswith("$") and "." in value):
        return value
    which, key = value[1:].split(".", 1)
    if which == "prev":
        index = len(results) - 1
    elif which.isdigit():
        index = int(which) - 1
    else:
        return value
    if not 0 <= index < len(results) or results[index] is None:
        raise ToolError(f"'{value}' refers to a step that has no result.")
    node = results[index]
    for part in key.split("."):
        if isinstance(node, dict) and part in node:
            node = node[part]
        elif isinstance(node, list) and part.isdigit() and int(part) < len(node):
            node = node[int(part)]
        else:
            raise ToolError(f"'{value}': the step's result has no '{key}'.")
    return node


@mcp.tool(name="word_describe")
async def word_describe(names: list[str] | None = None) -> str:
    """The Word tools: call with no names first for the list of all of them,
    one line each; then with the names you are about to use, once each, for
    their full description and argument schema.

    Every Word tool is run through word_run (one call) or word_run_steps
    (several), with exactly the arguments its schema gives. Start a task
    with word_describe().
    """
    if not names:
        return (catalog_lines() + "\n\nNot here yet: comments and tracked changes, footnotes. "
                "Paragraph direction (right to left), tab stops and "
                "adding table columns work on Windows only."
                "\n\nNext: word_describe(names=[...]) for the arguments of the tools "
                "you will use, then word_run or word_run_steps.")
    found, missing = {}, []
    for name in names:
        tool = mcp.registered_tool(name)
        if tool is None or name in CATALOG_TOOLS:
            missing.append(name)
            continue
        schema = dict(tool.parameters or {})
        params = (schema.get("properties") or {}).get("params")
        found[name] = {"description": tool.description, "params": params,
                       "takes_document": name not in TOOLS_WITHOUT_DOCUMENT_ARG}
    result = {"tools": found}
    if missing:
        result["unknown"] = missing
    return json.dumps(result, ensure_ascii=False)


@mcp.tool(name="word_run")
async def word_run(tool: str, params: dict | None = None, document: str | None = None):
    """Run one Word tool: `tool` is its name, `params` its arguments as
    word_describe gives them, `document` the document it works on (optional,
    as on every tool). Answers exactly what the tool answers."""
    return await run_one(tool, params, document)


@mcp.tool(name="word_run_steps")
async def word_run_steps(steps: list[dict], document: str | None = None,
                         stop_on_error: bool = True):
    """Run several Word tools in order, in one call.

    Each step is {"tool": name, "params": {...}}. A value "$prev.key" is the
    previous step's result field, and "$2.inserted.0.index" the second
    step's. A step fails when its tool raises or answers {"error": ...}; with
    stop_on_error, the first failing step ends the run and says why. Page
    pictures from word_get_page_image follow the report.
    """
    report, results, failed, images = [], [], [], []
    for number, step in enumerate(steps, start=1):
        step = step if isinstance(step, dict) else {}
        tool = step.get("tool")
        try:
            params = resolve_refs(step.get("params") or {}, results)
            blocks = await run_one(tool, params, step.get("document", document))
        except Exception as exc:  # noqa: BLE001 - reported per step
            report.append({"step": number, "tool": tool, "error": str(exc)})
            results.append(None)
            failed.append((number, tool, str(exc)))
            if stop_on_error:
                break
            continue
        text = text_of(blocks)
        try:
            data = json.loads(text) if text else {}
        except ValueError:
            data = {"text": text}
        if isinstance(data, dict) and data.get("error"):
            entry = {"step": number, "tool": tool, "error": data["error"]}
            details = {k: v for k, v in data.items() if k != "error"}
            if details:
                entry["details"] = details
            report.append(entry)
            results.append(None)
            failed.append((number, tool, str(data["error"])))
            if stop_on_error:
                break
            continue
        report.append({"step": number, "tool": tool, "result": data})
        results.append(data)
        # A page picture (word_get_page_image) comes after the report.
        images += [b for b in blocks if getattr(b, "type", "") == "image"]
    out = {"steps": report, "ran": len(report), "of": len(steps)}
    if failed:
        number, tool, message = failed[0]
        more = f" ({len(failed)} steps failed)" if len(failed) > 1 else ""
        out["error"] = f"Step {number} ({tool}) failed: {message}{more}"
    text = json.dumps(out, ensure_ascii=False)
    return [text, *images] if images else text


# The schema rewrites reach into SDK internals that have moved between majors
# before, so each degrades rather than stopping the server from starting.
for rewrite in (inline_schema_defs,
                lambda s: sum(1 for t in s.all_tools()
                              if setattr(t, "parameters", slim(t.parameters)) is None),
                add_document_arg):
    try:
        rewrite(mcp)
    except Exception:  # noqa: BLE001
        logger.warning("Could not rewrite the tool schemas", exc_info=True)

if CATALOG_MODE:
    try:
        lowlevel = getattr(mcp, "_lowlevel_server", None) or mcp._mcp_server
        lowlevel.instructions = (
            "Catalog mode: call word_describe() for the list of Word tools, "
            "word_describe(names) for their arguments, and run them with word_run or "
            "word_run_steps. Wherever the instructions say to call a tool, do it that "
            "way.\n" + (mcp.instructions or ""))
    except Exception:  # noqa: BLE001
        logger.warning("Could not add the catalog note to the instructions", exc_info=True)


def main():
    """Entry point for the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
