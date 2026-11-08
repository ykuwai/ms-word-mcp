"""The Mac halves of the tools, over Apple Events (appscript).

Each function here takes the same arguments as the COM implementation it is
registered for, and answers the same shape. What Word for Mac does
differently was measured first; see docs/PROBES.md. In short:

- A reference such as `paragraphs[3].text_object` is asked for again every
  time it is used, so a range found by `execute find` cannot be read through
  it. Text is searched in Python and ranges are made with `create range`.
- Counting asks `count each paragraph`, which needs the 8-byte integer fix in
  backend/mac.py.
- Word positions count UTF-16 units, so offsets into Python strings are
  converted (an emoji is one Python character and two Word characters).

appscript is imported inside the functions, so this module loads anywhere.
"""

from __future__ import annotations

import os

from .. import backend
from . import documents, edit, formatting
from . import text as text_tools
from ._common import mac_impl
from ._word import BUILTIN_STYLES, spacing_in_lines

MAX_PARAGRAPHS_PER_READ = text_tools.MAX_PARAGRAPHS_PER_READ

ALIGN_K = {"left": "align_paragraph_left", "center": "align_paragraph_center",
           "right": "align_paragraph_right", "justify": "align_paragraph_justify",
           "distribute": "align_paragraph_distribute"}
LINE_K = {"single": "line_space_single", "one_and_half": "line_space1_pt5",
          "double": "line_space_double", "at_least": "line_space_at_least",
          "exactly": "line_space_exactly", "multiple": "line_space_multiple"}
LINE_CODES = {"single": 0, "one_and_half": 1, "double": 2, "at_least": 3, "exactly": 4,
              "multiple": 5}


def _k():
    from appscript import k

    return k


def _name(value) -> str:
    """An appscript keyword's name (k.line_space_single -> line_space_single)."""
    return getattr(value, "name", None) or str(value).removeprefix("k.")


def _from_k(mapping: dict, value):
    name = _name(value)
    for key, kname in mapping.items():
        if kname == name:
            return key
    return name


def _u16(s: str, i: int) -> int:
    """Word's position for Python index `i` into `s` (UTF-16 units)."""
    return len(s[:i].encode("utf-16-le")) // 2


def _font_name(value):
    """A font name as Windows gives it: Word for Mac adds " ()" to theme fonts."""
    if not isinstance(value, str) or not value:
        return None
    return value[:-3] if value.endswith(" ()") else value


def _clean(content) -> str:
    return (content or "").rstrip("\r\x07") if isinstance(content, str) else ""


def _count(doc) -> int:
    return int(doc.count(each=_k().paragraph))


def _para(doc, index: int):
    count = _count(doc)
    if not 1 <= index <= count:
        raise ValueError(f"There is no paragraph {index}: the document has {count} "
                         "paragraphs (numbered from 1).")
    return doc.paragraphs[index]


def _span(doc, start: int, end: int | None) -> tuple[int, int]:
    count = _count(doc)
    end = start if end is None else end
    if end < start:
        raise ValueError(f"end ({end}) comes before start ({start}).")
    if not 1 <= start <= count or end > count:
        raise ValueError(f"Paragraphs {start} to {end} are not all in the document: it "
                         f"has {count} paragraphs (numbered from 1).")
    return start, end


def _bounds(doc, index: int) -> tuple[int, int]:
    rng = doc.paragraphs[index].text_object
    return rng.start_of_content.get(), rng.end_of_content.get()


def _style_value(doc, name: str):
    """What to set `style` to: a built-in enumerator, or the document's own style."""
    key = " ".join(name.lower().split())
    if key in BUILTIN_STYLES:
        enum = getattr(_k(), "style_" + key.replace(" ", "_").replace("heading_", "heading"),
                       None)
        if enum is not None:
            return enum
    style = doc.Word_styles[name]
    try:
        style.name_local.get()
    except Exception:  # noqa: BLE001
        raise ValueError(
            f"No style called '{name}' in this document. Use a built-in English name "
            f"({', '.join(sorted(k.title() for k in BUILTIN_STYLES))}) or one that "
            "word_list_styles gives.") from None
    return style


def _info(doc, index: int, detail: bool = False) -> dict:
    p = doc.paragraphs[index]
    out = {"index": index, "text": _clean(p.text_object.content.get()),
           "style": p.style.name_local.get()}
    level = _name(p.outline_level.get())
    if level != "outline_level_body_text":
        out["outline_level"] = int(level.rsplit("level", 1)[1])
    if detail:
        fmt = p.paragraph_format
        font = p.text_object.font_object
        rule = _from_k(LINE_K, fmt.line_spacing_rule.get())
        spacing = fmt.line_spacing.get()
        out["alignment"] = _from_k(ALIGN_K, fmt.alignment.get())
        out["fonts"] = {"latin": _font_name(font.ascii_name.get()),
                        "east_asian": _font_name(font.east_asian_name.get()),
                        "complex_script": _font_name(font.complex_script_name.get()),
                        "size": font.font_size.get()}
        out["line_spacing_rule"] = rule
        lines = spacing_in_lines(LINE_CODES.get(rule), spacing)
        if lines is not None:
            out["line_spacing_lines"] = lines
        else:
            out["line_spacing_pt"] = spacing
        out["space_before"] = fmt.space_before.get()
        out["space_after"] = fmt.space_after.get()
    return out


# -- documents ------------------------------------------------------------------

def _full(doc) -> str:
    """The document's POSIX path, or its name if it has never been saved.

    Word for Mac's `full name` is an HFS path ("Macintosh HD:Users:…").
    """
    full = doc.posix_full_name.get()
    return full if isinstance(full, str) and full else doc.name.get()


def _summary(doc) -> dict:
    return {"name": doc.name.get(), "full_name": _full(doc),
            "saved": bool(doc.saved.get()), "read_only": bool(doc.read_only.get())}


@mac_impl(documents._get_app_info_impl)
def get_app_info():
    word = backend.word
    app = word.get_app()
    docs = word.open_documents()
    return {"success": True, "version": app.version.get(), "build": None,
            "platform": backend.PLATFORM_NAME,
            "documents": [{"index": d.index, "name": d.name, "full_name": d.full_name}
                          for d in docs],
            "active_document": docs[0].full_name if docs else None,
            "target_document": word.target_full_name,
            "started_by_server": word.launched_by_us}


@mac_impl(documents._list_documents_impl)
def list_documents():
    word = backend.word
    app = word.get_app()
    return {"success": True,
            "documents": [{"index": d.index, **_summary(app.documents[d.index])}
                          for d in word.open_documents()],
            "target_document": word.target_full_name}


@mac_impl(documents._activate_document_impl)
def activate_document(document):
    return documents._activate_document_impl(document)


@mac_impl(documents._create_document_impl)
def create_document(template, save_as):
    if template:
        raise ValueError("A template is not supported on macOS yet. Leave `template` "
                         "out, or open the template with word_open_document.")
    if save_as and os.path.exists(save_as):
        raise ValueError(f"'{save_as}' already exists. Choose another name; nothing "
                         "was created.")
    word = backend.word
    app = word.get_app(allow_launch=True)
    doc = app.create_new_document()
    name = doc.name.get()
    if save_as:
        app.documents[name].save_as(file_name=save_as, file_format=_k().format_document)
        if not os.path.exists(save_as):
            raise RuntimeError(f"Word answered without error, but there is no file at "
                               f"'{save_as}'.")
        name = os.path.basename(save_as)
    target = app.documents[name]
    word.target_full_name = _full(target)
    return {"success": True, **_summary(target),
            "note": "This document is now the target of calls without `document`."}


@mac_impl(documents._save_document_impl)
def save_document():
    doc = backend.word.document()
    if not doc.path.get():
        raise ValueError("This document has never been saved, so it has no file to save "
                         "to. Use word_save_document_as with a full path.")
    doc.save()
    return {"success": True, **_summary(doc)}


@mac_impl(documents._close_document_impl)
def close_document(save):
    word = backend.word
    doc = word.document()
    name, full_name = doc.name.get(), _full(doc)
    if save is None and not doc.saved.get():
        raise ValueError(f"'{name}' has unsaved changes. Pass save=true to save them or "
                         "save=false to discard them.")
    doc.close(saving=_k().yes if save else _k().no)
    if word.target_full_name == full_name:
        word.target_full_name = None
    return {"success": True, "closed": full_name}


# -- reading ---------------------------------------------------------------------

@mac_impl(text_tools._get_document_info_impl)
def get_document_info():
    k = _k()
    doc = backend.word.document()
    setup = doc.page_setup
    return {"success": True, **_summary(doc), "paragraphs": _count(doc),
            "pages": doc.compute_statistics(statistic=k.statistic_pages),
            "words": doc.compute_statistics(statistic=k.statistic_words),
            "characters": doc.compute_statistics(statistic=k.statistic_characters),
            "page": {"width_pt": setup.page_width.get(),
                     "height_pt": setup.page_height.get(),
                     "margins_pt": {"top": setup.top_margin.get(),
                                    "bottom": setup.bottom_margin.get(),
                                    "left": setup.left_margin.get(),
                                    "right": setup.right_margin.get()}},
            "track_revisions": bool(doc.track_revisions.get())}


@mac_impl(text_tools._get_outline_impl)
def get_outline():
    doc = backend.word.document()
    count = _count(doc)
    headings = []
    for i in range(1, count + 1):
        level = _name(doc.paragraphs[i].outline_level.get())
        if level != "outline_level_body_text":
            info = _info(doc, i)
            headings.append({"index": i, "level": info.get("outline_level"),
                             "text": info["text"], "style": info["style"]})
    return {"success": True, "paragraphs": count, "headings": headings}


@mac_impl(text_tools._get_paragraphs_impl)
def get_paragraphs(start, end, detail):
    doc = backend.word.document()
    count = _count(doc)
    if start > count:
        raise ValueError(f"The document has {count} paragraphs; start is {start}.")
    last = min(count, end if end is not None else start + MAX_PARAGRAPHS_PER_READ - 1,
               start + MAX_PARAGRAPHS_PER_READ - 1)
    out = {"success": True,
           "paragraphs": [_info(doc, i, detail) for i in range(start, last + 1)],
           "count": count}
    if last < count:
        out["more"] = f"Paragraphs {last + 1} to {count} follow; pass start={last + 1}."
    return out


# -- writing ---------------------------------------------------------------------

def _scroll(doc, start: int, end: int) -> None:
    from ._common import FOLLOW_EDITS

    if not FOLLOW_EDITS:
        return
    try:
        doc.active_window.scroll_into_view(doc.create_range(start=start, end_=end))
    except Exception:  # noqa: BLE001 - the edit landed either way
        pass


@mac_impl(text_tools._insert_paragraphs_impl)
def insert_paragraphs(paragraphs, after, before):
    doc = backend.word.document()
    count = _count(doc)
    texts = [p["text"].replace("\n", "\v") for p in paragraphs]
    if before is not None:
        _para(doc, before)
        at = before
    elif after is not None:
        if after > count:
            raise ValueError(f"There is no paragraph {after}: the document has {count}.")
        at = after + 1
    else:
        at = count + 1
    if at > count:
        start, end = _bounds(doc, count)
        if count == 1 and _clean(doc.paragraphs[1].text_object.content.get()) == "":
            first = 1
            doc.create_range(start=start, end_=start).content.set("\r".join(texts))
        else:
            first = count + 1
            doc.create_range(start=end - 1, end_=end - 1).content.set(
                "\r" + "\r".join(texts))
    else:
        first = at
        start, _ = _bounds(doc, at)
        doc.create_range(start=start, end_=start).content.set("\r".join(texts) + "\r")
    written, warnings = [], []
    for offset, item in enumerate(paragraphs):
        index = first + offset
        if item.get("style"):
            p = doc.paragraphs[index]
            p.style.set(_style_value(doc, item["style"]))
            p.text_object.font_object.reset()
        info = _info(doc, index)
        if info["text"] != texts[offset]:
            warnings.append(f"Paragraph {index} reads back as {info['text']!r}, not "
                            f"{texts[offset]!r}.")
        written.append(info)
    if written:
        _scroll(doc, *_bounds(doc, written[-1]["index"]))
    out = {"success": True, "inserted": written, "paragraphs": _count(doc)}
    if warnings:
        out["warnings"] = warnings
    return out


@mac_impl(text_tools._set_paragraph_text_impl)
def set_paragraph_text(index, text):
    doc = backend.word.document()
    _para(doc, index)
    start, _ = _bounds(doc, index)
    old = _clean(doc.paragraphs[index].text_object.content.get())
    value = text.replace("\n", "\v")
    doc.create_range(start=start, end_=start + _u16(old, len(old))).content.set(value)
    info = _info(doc, index)
    _scroll(doc, *_bounds(doc, index))
    out = {"success": True, "paragraph": info, "was": old}
    if info["text"] != value:
        out["warnings"] = [f"It reads back as {info['text']!r}."]
    return out


@mac_impl(text_tools._delete_paragraphs_impl)
def delete_paragraphs(start, end):
    doc = backend.word.document()
    start, end = _span(doc, start, end)
    before = _count(doc)
    removed = [_clean(doc.paragraphs[i].text_object.content.get())
               for i in range(start, end + 1)]
    first, _ = _bounds(doc, start)
    _, last = _bounds(doc, end)
    if end == before and start > 1:
        first -= 1
    if end == before and start == 1:
        last -= 1  # the document's last paragraph mark stays
    doc.create_range(start=first, end_=last).content.set("")
    after = _count(doc)
    out = {"success": True, "deleted": end - start + 1, "texts": removed,
           "paragraphs": after}
    if before - after != end - start + 1 and not (start == 1 and end == before):
        out["warnings"] = [f"The document went from {before} to {after} paragraphs."]
    return out


def _matches(doc, find, match_case, whole_word):
    """(start, end) of each match, in Word positions, found in Python."""
    import re

    content = doc.text_object.content.get() or ""
    pattern = re.escape(find)
    if whole_word:
        pattern = rf"(?<!\w){pattern}(?!\w)"
    flags = 0 if match_case else re.IGNORECASE
    return [(_u16(content, m.start()), _u16(content, m.end()))
            for m in re.finditer(pattern, content, flags)]


def _index_at(doc, position: int) -> int:
    return int(doc.create_range(start=0, end_=position + 1).count(each=_k().paragraph))


@mac_impl(text_tools._find_replace_impl)
def find_replace(find, replace, match_case, whole_word, wildcards):
    if wildcards:
        raise ValueError("Wildcards are not supported on macOS yet. Search for plain "
                         "text, or use Windows for a wildcard search.")
    doc = backend.word.document()
    found = _matches(doc, find, match_case, whole_word)
    out = {"success": True, "matches": len(found),
           "paragraphs": sorted({_index_at(doc, s) for s, _ in found})}
    if replace is None or not found:
        return out
    value = replace.replace("\n", "\v")
    for start, end in reversed(found):
        doc.create_range(start=start, end_=end).content.set(value)
    out["replaced"] = len(found)
    left = _matches(doc, find, match_case, whole_word)
    if left and find.lower() not in value.lower():
        out["warnings"] = [f"{len(left)} matches are still there after replacing."]
    return out


# -- formatting ------------------------------------------------------------------

@mac_impl(formatting._format_text_impl)
def format_text(start, end, match, latin_font, east_asian_font, complex_script_font,
                size, size_complex_script, bold, italic, underline, color):
    k = _k()
    doc = backend.word.document()
    start, end = _span(doc, start, end)
    span_start, _ = _bounds(doc, start)
    _, span_end = _bounds(doc, end)
    if match is None:
        spans = [(span_start, span_end)]
    else:
        spans = [(s, e) for s, e in _matches(doc, match, True, False)
                 if s >= span_start and e <= span_end]
        if not spans:
            raise ValueError(f"'{match}' does not occur in paragraphs {start} to {end}. "
                             "Nothing was changed.")
    refused, warnings, readback = {}, [], []
    for s, e in spans:
        font = doc.create_range(start=s, end_=e).font_object
        for script, props, wanted in (
                ("latin", ("ascii_name", "other_name"), latin_font),
                ("east_asian", ("east_asian_name",), east_asian_font),
                ("complex_script", ("complex_script_name",), complex_script_font)):
            if wanted is None:
                continue
            for prop in props:
                try:
                    getattr(font, prop).set(wanted)
                except Exception as exc:  # noqa: BLE001 - reported below
                    refused[script] = str(exc).splitlines()[0]
        if size is not None:
            font.font_size.set(size)
        if bold is not None:
            font.bold.set(bold)
            font.bold_bi.set(bold)
        if italic is not None:
            font.italic.set(italic)
            font.italic_bi.set(italic)
        if underline is not None:
            font.underline.set(k.underline_single if underline else k.underline_none)
        if color is not None:
            value = color.lstrip("#")
            font.color.set([int(value[i:i + 2], 16) for i in (0, 2, 4)])
        got = {"text": _clean(doc.create_range(start=s, end_=e).content.get())[:80],
               "latin": _font_name(font.ascii_name.get()),
               "east_asian": _font_name(font.east_asian_name.get()),
               "complex_script": _font_name(font.complex_script_name.get()),
               "size": font.font_size.get(), "bold": font.bold.get()}
        readback.append(got)
        for script, wanted in (("latin", latin_font), ("east_asian", east_asian_font),
                               ("complex_script", complex_script_font)):
            if wanted is not None and got[script] != wanted:
                why = (" Word refused it" + (": an East Asian font needs East Asian "
                                             "glyphs" if script == "east_asian" else "")
                       + f" ({refused[script]})." if script in refused else "")
                warnings.append(f"The {script.replace('_', ' ')} font reads back as "
                                f"{got[script]!r}, not {wanted!r}.{why}")
    _scroll(doc, *spans[0])
    out = {"success": True, "ranges": len(spans), "read_back": readback[:20]}
    if size_complex_script is not None:
        warnings.append("size_complex_script is not set on macOS yet.")
    if warnings:
        out["warnings"] = sorted(set(warnings))
    return out


@mac_impl(formatting._set_paragraph_format_impl)
def set_paragraph_format(start, end, alignment, line_spacing_rule, line_spacing,
                         space_before, space_after, first_line_indent, left_indent,
                         right_indent, snap_to_grid, keep_with_next, page_break_before):
    k = _k()
    doc = backend.word.document()
    start, end = _span(doc, start, end)
    for i in range(start, end + 1):
        fmt = doc.paragraphs[i].paragraph_format
        if alignment is not None:
            fmt.alignment.set(getattr(k, ALIGN_K[alignment]))
        if line_spacing_rule is not None:
            fmt.line_spacing_rule.set(getattr(k, LINE_K[line_spacing_rule]))
            if line_spacing is not None:
                fmt.line_spacing.set(line_spacing * 12.0 if line_spacing_rule == "multiple"
                                     else line_spacing)
        for prop, value in (("space_before", space_before), ("space_after", space_after),
                            ("first_line_indent", first_line_indent),
                            ("paragraph_format_left_indent", left_indent),
                            ("paragraph_format_right_indent", right_indent),
                            ("keep_with_next", keep_with_next),
                            ("page_break_before", page_break_before)):
            if value is not None:
                getattr(fmt, prop).set(value)
        if snap_to_grid is not None:
            fmt.disable_line_height_grid.set(not snap_to_grid)
    infos = [_info(doc, i, detail=True) for i in range(start, min(end, start + 49) + 1)]
    out = {"success": True, "paragraphs": infos}
    if line_spacing_rule is not None and infos:
        wanted = spacing_in_lines(LINE_CODES[line_spacing_rule],
                                  line_spacing * 12.0 if line_spacing_rule == "multiple"
                                  else line_spacing)
        got = infos[0]
        if wanted is not None and got.get("line_spacing_lines") != wanted:
            out["warnings"] = [f"Line spacing reads back as {got['line_spacing_rule']}."]
    _scroll(doc, _bounds(doc, start)[0], _bounds(doc, end)[1])
    return out


@mac_impl(formatting._apply_style_impl)
def apply_style(start, end, style, clear_direct_formatting):
    doc = backend.word.document()
    start, end = _span(doc, start, end)
    value = _style_value(doc, style)
    for i in range(start, end + 1):
        p = doc.paragraphs[i]
        p.style.set(value)
        if clear_direct_formatting:
            p.text_object.font_object.reset()
    infos = [_info(doc, i) for i in range(start, min(end, start + 49) + 1)]
    wanted = infos[0]["style"] if infos else style
    out = {"success": True, "style": wanted, "paragraphs": infos}
    off = [p["index"] for p in infos if p["style"] != wanted]
    if off:
        out["warnings"] = [f"Paragraphs {off} read back with another style."]
    _scroll(doc, _bounds(doc, start)[0], _bounds(doc, end)[1])
    return out


# -- undo ------------------------------------------------------------------------

@mac_impl(edit._undo_impl)
def undo(times):
    doc = backend.word.document()
    done = sum(1 for _ in range(times) if doc.undo(times=1))
    out = {"success": True, "undone": done,
           "note": ("On macOS each change a tool made is its own undo step, so one tool "
                    "call may take several.")}
    if done < times:
        out["warnings"] = [f"Only {done} of {times} steps could be undone."]
    return out


@mac_impl(edit._redo_impl)
def redo(times):
    doc = backend.word.document()
    done = sum(1 for _ in range(times) if doc.redo(times=1))
    out = {"success": True, "redone": done}
    if done < times:
        out["warnings"] = [f"Only {done} of {times} steps could be redone."]
    return out


# -- open, save as ----------------------------------------------------------------

SAVE_K = {".docx": "format_document", ".docm": "format_documentME",
          ".dotx": "format_template", ".doc": "format_document97", ".rtf": "format_rtf",
          ".txt": "format_text", ".pdf": "format_PDF", ".odt": "format_ODT",
          ".htm": "format_HTML", ".html": "format_HTML"}


@mac_impl(documents._open_document_impl)
def open_document(path, read_only, visible):
    from ..backend.documents import full_name_key

    if not os.path.exists(path):
        raise ValueError(f"There is no file at '{path}'.")
    word = backend.word
    app = word.get_app(allow_launch=True)
    for d in word.open_documents():
        if full_name_key(d.full_name) == full_name_key(path):
            word.target_full_name = d.full_name
            return {"success": True, **_summary(app.documents[d.index]),
                    "already_open": True}
    before = {d.full_name for d in word.open_documents()}
    app.open(file_name=path, read_only=read_only, add_to_recent_files=False)
    opened = [d for d in word.open_documents() if d.full_name not in before]
    if len(opened) != 1:
        raise RuntimeError(f"Word answered without error, but '{path}' is not among the "
                           "open documents.")
    word.target_full_name = opened[0].full_name
    doc = app.documents[opened[0].index]
    return {"success": True, **_summary(doc), "paragraphs": _count(doc),
            "note": "This document is now the target of calls without `document`."}


@mac_impl(documents._save_document_as_impl)
def save_document_as(path, overwrite):
    if os.path.exists(path) and not overwrite:
        raise ValueError(f"'{path}' already exists. Pass overwrite=true to replace it.")
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        raise ValueError(f"The folder '{folder}' does not exist.")
    k = _k()
    word = backend.word
    doc = word.document()
    ext = os.path.splitext(path)[1].lower()
    own = _full(doc)
    doc.save_as(file_name=path, file_format=getattr(k, SAVE_K[ext]))
    if not os.path.exists(path):
        raise RuntimeError(f"Word answered without error, but there is no file at '{path}'.")
    if ext != ".pdf":
        word.target_full_name = None
        for d in word.open_documents():
            if os.path.basename(d.full_name) == os.path.basename(path):
                word.target_full_name = d.full_name
        own = word.target_full_name or own
    return {"success": True, "path": path, "bytes": os.path.getsize(path), "document": own}


STYLE_TYPES = {"style_type_paragraph": "paragraph", "style_type_character": "character",
               "style_type_table": "table", "style_type_list": "list",
               "style_type_linked": "linked", "style_type_paragraph_only": "paragraph"}


@mac_impl(formatting._list_styles_impl)
def list_styles(in_use_only):
    doc = backend.word.document()
    count = int(doc.count(each=_k().Word_style))
    out = []
    for i in range(1, count + 1):
        s = doc.Word_styles[i]
        if in_use_only and not s.in_use.get():
            continue
        kind = STYLE_TYPES.get(_name(s.style_type.get()), _name(s.style_type.get()))
        if kind in ("table", "list"):
            continue
        entry = {"name": s.name_local.get(), "type": kind, "builtin": bool(s.built_in.get())}
        if kind in ("paragraph", "linked"):
            font = s.font_object
            entry["fonts"] = {"latin": _font_name(font.ascii_name.get()),
                              "east_asian": _font_name(font.east_asian_name.get()),
                              "size": font.font_size.get()}
        out.append(entry)
    return {"success": True, "styles": out}
