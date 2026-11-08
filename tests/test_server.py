"""The server around the tools: catalog mode, the `document` argument,
refusals of unknown arguments, and failures marked as errors."""

import asyncio
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

from ms_word_mcp import backend, server
from ms_word_mcp.tools import _common

ROOT = Path(__file__).parent.parent


def call(name, arguments):
    async def go():
        result = await server.mcp.call_tool(name, arguments)
        return server.content_of(result)

    return asyncio.run(go())


def text(blocks):
    return json.loads(server.text_of(blocks))


def test_tool_names_and_annotations():
    tools = server.mcp.all_tools()
    names = [t.name for t in tools]
    assert len(names) == len(set(names))
    for tool in tools:
        assert tool.name.startswith("word_")
        if tool.name in server.CATALOG_TOOLS:
            continue
        annotations = tool.annotations
        for key in ("readOnlyHint", "destructiveHint", "idempotentHint", "openWorldHint"):
            assert getattr(annotations, key, getattr(annotations, _snake(key), None)) \
                is not None, (tool.name, key)


def _snake(name):
    return "".join("_" + c.lower() if c.isupper() else c for c in name)


def test_every_input_model_forbids_unknown_keys():
    from pydantic import BaseModel

    from ms_word_mcp.tools import documents, edit, formatting, text as text_tools

    for module in (documents, edit, formatting, text_tools):
        for value in vars(module).values():
            if isinstance(value, type) and issubclass(value, BaseModel) \
                    and value.__module__ == module.__name__:
                assert value.model_config.get("extra") == "forbid", value


def test_document_argument_reaches_the_worker(word):
    word.doc.Name = "report.docx"
    seen = []
    word.document = lambda: (seen.append(backend.call_document.get()), word.doc)[1]

    original = word.execute

    def execute(func, *args, **kwargs):
        seen.append(("on worker", backend.call_document.get()))
        return original(func, *args, **kwargs)

    word.execute = execute
    call("word_get_paragraphs", {"params": {"start": 1}, "document": "report.docx"})
    assert ("on worker", "report.docx") in seen


def test_document_argument_passes_through_word_run(word):
    seen = []
    original = word.execute

    def execute(func, *args, **kwargs):
        seen.append(backend.call_document.get())
        return original(func, *args, **kwargs)

    word.execute = execute
    call("word_run", {"tool": "word_get_outline", "document": "notes.docx"})
    assert seen == ["notes.docx"]


def test_unknown_argument_next_to_params_is_refused(word):
    with pytest.raises(Exception, match="did you mean 'document'"):
        call("word_get_outline", {"documnet": "x"})


def test_validation_error_is_one_readable_sentence(word):
    with pytest.raises(Exception) as info:
        call("word_get_paragraphs", {"params": {"start": 0}})
    assert "Nothing was applied" in str(info.value)
    assert "params.start" in str(info.value)


def test_a_failure_answer_is_marked_as_an_error(word):
    async def go():
        return await server.mcp.call_tool(
            "word_get_paragraphs", {"params": {"start": 99}})

    result = asyncio.run(go())
    if hasattr(result, "is_error"):
        assert result.is_error
    else:
        blocks = result[0] if isinstance(result, tuple) else result
        assert "error" in json.loads(server.text_of(blocks))


def test_run_steps_passes_results_on_and_stops_at_a_failure(word, make_doc):
    make_doc("one")
    out = text(call("word_run_steps", {"steps": [
        {"tool": "word_insert_paragraphs",
         "params": {"paragraphs": [{"text": "two"}]}},
        {"tool": "word_set_paragraph_text",
         "params": {"index": "$1.inserted.0.index", "text": "TWO"}},
        {"tool": "word_get_paragraphs", "params": {"start": 50}},
        {"tool": "word_get_outline"},
    ]}))
    assert word.doc.texts() == ["one", "TWO"]
    assert out["ran"] == 3 and out["of"] == 4
    assert out["error"].startswith("Step 3 (word_get_paragraphs) failed")


def test_unknown_tool_gets_a_suggestion(word):
    with pytest.raises(Exception, match="word_get_outline"):
        call("word_run", {"tool": "word_get_outlines"})


def test_writing_calls_name_an_undo_step_and_reads_do_not(word, make_doc):
    make_doc("one")
    call("word_get_outline", {})
    call("word_insert_paragraphs", {"params": {"paragraphs": [{"text": "x"}]}})
    assert word.calls == [("_get_outline_impl", None),
                          ("_insert_paragraphs_impl", "Insert paragraphs")]


def test_macos_refuses_tools_without_a_mac_half(word, monkeypatch):
    monkeypatch.setattr(backend, "IS_MACOS", True)
    out = json.loads(_common.run(lambda: {"success": True}))
    assert "does not work on macOS yet" in out["error"]


@pytest.mark.parametrize("mode, expected", [
    (None, "catalog"), ("catalog", "catalog"), ("full", "full"), ("ful", "full")])
def test_catalog_mode_is_the_default(mode, expected):
    env = {k: v for k, v in os.environ.items() if k != "WORD_TOOL_MODE"}
    if mode is not None:
        env["WORD_TOOL_MODE"] = mode
    script = (
        "import asyncio\n"
        "from ms_word_mcp.server import mcp\n"
        "print(','.join(sorted(t.name for t in asyncio.run(mcp.list_tools()))))\n")
    out = subprocess.run([sys.executable, "-c", script], env=env, cwd=ROOT,
                         capture_output=True, text=True, encoding="utf-8", check=True)
    listed = out.stdout.strip().split(",")
    if expected == "catalog":
        assert listed == ["word_describe", "word_run", "word_run_steps"]
    else:
        assert "word_describe" not in listed and len(listed) > 15


def test_describe_lists_every_tool_a_line_each():
    async def go():
        return await server.word_describe()

    lines = [line for line in asyncio.run(go()).splitlines()
             if line.startswith("word_")]
    names = {t.name for t in server.mcp.all_tools()} - server.CATALOG_TOOLS
    assert {line.split(":")[0] for line in lines} == names


def test_describe_takes_the_document_argument_every_call_carries(word):
    """The instructions say to pass `document` on every call; word_describe
    has nothing to do with it and must not refuse it (ppt-mcp #367)."""
    lines = server.text_of(call("word_describe", {"document": "report.docx"}))
    assert "word_get_outline" in lines
