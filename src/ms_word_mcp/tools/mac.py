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
from . import _props as P
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
    """Scroll the document's window to about where [start, end) is.

    Word for Mac has no `scroll into view`. Setting the window's `vertical
    percent scrolled` moves the view (measured by pictures of the window:
    50 showed page 4 of 17), although reading it back answers 0. The place is
    the edit's share of the document's characters, so it is near the edit
    rather than on it. The user's selection is not touched.
    """
    from ._common import FOLLOW_EDITS

    if not FOLLOW_EDITS:
        return
    try:
        total = doc.text_object.end_of_content.get()
        if total > 0:
            doc.active_window.vertical_percent_scrolled.set(
                max(0, min(100, round(100 * start / total))))
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
        try:
            doc.create_range(start=start, end_=start).content.set("\r".join(texts) + "\r")
        except Exception:  # noqa: BLE001 - measured below
            # Measured on Word 16.114: text set at the start of a paragraph that
            # holds an equation is refused (-1708). The new paragraphs go in
            # at the end of the paragraph before instead; they then take that
            # paragraph's formatting, so a style is set on each below.
            if at == 1:
                raise
            _, prev_end = _bounds(doc, at - 1)
            doc.create_range(start=prev_end - 1, end_=prev_end - 1).content.set(
                "\r" + "\r".join(texts))
            paragraphs = [{**item, "style": item.get("style") or "Normal"}
                          for item in paragraphs]
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
def format_text(start, end, match, font):
    doc = backend.word.document()
    spec = P.with_latin_extended(P.given(font, P.FONT_NAMES))
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
    warnings, readback = [], []
    for s, e in spans:
        rng = doc.create_range(start=s, end_=e)
        refused = P.mac_apply(rng.font_object, P.FONT_PROPS, spec)
        got = P.mac_read(rng.font_object, P.FONT_PROPS)
        warnings += P.mismatches(spec, got, refused)
        readback.append({"text": _clean(rng.content.get())[:80],
                         **formatting.font_readback(got, spec)})
    _scroll(doc, *spans[0])
    out = {"success": True, "ranges": len(spans), "read_back": readback[:20]}
    if warnings:
        out["warnings"] = sorted(set(warnings))
    return out


@mac_impl(formatting._set_paragraph_format_impl)
def set_paragraph_format(start, end, paragraph):
    doc = backend.word.document()
    spec = P.given(paragraph, P.PARAGRAPH_NAMES)
    start, end = _span(doc, start, end)
    refused = {}
    for i in range(start, end + 1):
        fmt = doc.paragraphs[i].paragraph_format
        refused.update(P.mac_apply(fmt, P.PARAGRAPH_PROPS, spec))
        P.mac_apply_line_spacing(fmt, spec.get("line_spacing_rule"),
                                 spec.get("line_spacing"))
    out, warnings = [], []
    for i in range(start, min(end, start + 49) + 1):
        fmt = doc.paragraphs[i].paragraph_format
        got = {**P.mac_read(fmt, P.PARAGRAPH_PROPS), **P.mac_read_line_spacing(fmt)}
        if not warnings:
            warnings = P.mismatches(spec, got, refused)
        info = _info(doc, i)
        out.append({"index": i, "text": info["text"][:80], "style": info["style"],
                    **formatting.paragraph_readback(got, spec)})
    result = {"success": True, "paragraphs": out}
    if warnings:
        result["warnings"] = warnings
    _scroll(doc, _bounds(doc, start)[0], _bounds(doc, end)[1])
    return result


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
               "style_type_linked": "paragraph", "style_type_paragraph_only": "paragraph"}


@mac_impl(formatting._list_styles_impl)
def list_styles(in_use_only):
    doc = backend.word.document()
    count = int(doc.count(each=_k().Word_style))
    out = []
    for i in range(1, count + 1):
        s = doc.Word_styles[i]
        if in_use_only and not s.in_use.get():
            continue
        raw = _name(s.style_type.get())
        kind = STYLE_TYPES.get(raw, raw)
        if kind in ("table", "list"):
            continue
        entry = {"name": s.name_local.get(), "type": kind, "builtin": bool(s.built_in.get())}
        if raw == "style_type_linked":
            entry["linked"] = True
        if kind == "paragraph":
            font = s.font_object
            entry["fonts"] = {"latin": _font_name(font.ascii_name.get()),
                              "east_asian": _font_name(font.east_asian_name.get()),
                              "size": font.font_size.get()}
        out.append(entry)
    return {"success": True, "styles": out}


# -- reading and defining styles ---------------------------------------------------

def _style_object(doc, name: str):
    """The style called `name`: a built-in English name, or the document's own.

    Measured: `Word_styles[k.style_heading1]` finds 見出し 1 in a Japanese Word,
    and `Word_styles[name]` finds a style by its local name.
    """
    value = _style_value(doc, name)
    return doc.Word_styles[value] if not hasattr(value, "name_local") else value


def _mac_style_name(ref):
    try:
        name = ref.name_local.get()
    except Exception:  # noqa: BLE001
        return None
    return name or None  # Normal's base style answers a style named ""


def _describe_mac_style(s) -> dict:
    raw = _name(s.style_type.get())
    kind = STYLE_TYPES.get(raw, raw)
    out = {"name": s.name_local.get(), "type": kind, "builtin": bool(s.built_in.get()),
           "in_use": bool(s.in_use.get())}
    try:
        out["description"] = s.description.get()
    except Exception:  # noqa: BLE001
        pass
    if kind == "paragraph":
        out["linked"] = raw == "style_type_linked"
        out["base_style"] = _mac_style_name(s.base_style)
        out["next_style"] = _mac_style_name(s.next_paragraph_style)
        out["automatically_update"] = bool(s.automatically_update.get())
        out["no_space_between_same_style"] = bool(s.no_space_between_same.get())
    elif kind == "character":
        out["base_style"] = _mac_style_name(s.base_style)
    if kind in ("paragraph", "character"):
        out["font"] = P.mac_read(s.font_object, P.FONT_PROPS)
    if kind == "paragraph":
        fmt = s.paragraph_format
        out["paragraph"] = {**P.mac_read(fmt, P.PARAGRAPH_PROPS),
                            **P.mac_read_line_spacing(fmt)}
    return out


@mac_impl(formatting._get_style_impl)
def get_style(style):
    doc = backend.word.document()
    return {"success": True, "style": _describe_mac_style(_style_object(doc, style))}


@mac_impl(formatting._set_style_impl)
def set_style(style, create, base_style, next_style, font, paragraph, quick_style,
              priority, automatically_update, no_space_between_same_style):
    k = _k()
    doc = backend.word.document()
    exists = True
    try:
        _style_object(doc, style).name_local.get()
    except Exception:  # noqa: BLE001
        exists = False
    if create:
        if exists:
            raise ValueError(f"A style called '{style}' already exists. Leave out "
                             "`create` to change it.")
        # Measured on Word 16.114: `make new Word style`, with or without a
        # type, answers -1708 and makes nothing, whatever the name.
        raise ValueError(
            "Creating a style does not work through Word for Mac's scripting (Word "
            "answers -1708 and makes nothing). Create it in Word's Styles pane, or on "
            "Windows, then change it here; or base the look on a built-in style such "
            "as \"Body Text\" and change that.")
    else:
        s = _style_object(doc, style)
        s.name_local.get()  # raises if it does not exist
    kind = STYLE_TYPES.get(_name(s.style_type.get()), "paragraph")
    if paragraph and kind != "paragraph":
        raise ValueError(f"'{style}' is a {kind} style, which has no paragraph formatting.")

    refused = {}
    if base_style is not None:
        try:
            s.base_style.set(_style_value(doc, base_style))
        except Exception as exc:  # noqa: BLE001
            refused["base_style"] = str(exc).splitlines()[0]
    if next_style is not None:
        try:
            s.next_paragraph_style.set(_style_value(doc, next_style))
        except Exception as exc:  # noqa: BLE001
            refused["next_style"] = str(exc).splitlines()[0]
    for key, prop, value in (("automatically_update", "automatically_update",
                              automatically_update),
                             ("no_space_between_same_style", "no_space_between_same",
                              no_space_between_same_style)):
        if value is not None:
            try:
                getattr(s, prop).set(value)
            except Exception as exc:  # noqa: BLE001
                refused[key] = str(exc).splitlines()[0]
    for key, value in (("quick_style", quick_style), ("priority", priority)):
        if value is not None:
            refused[key] = "Word for Mac's scripting has no such property; it works on Windows."
    font_spec = P.with_latin_extended(P.given(font, P.FONT_NAMES))
    para_spec = P.given(paragraph, P.PARAGRAPH_NAMES)
    font_refused = P.mac_apply(s.font_object, P.FONT_PROPS, font_spec) if font_spec else {}
    para_refused = {}
    if para_spec:
        fmt = s.paragraph_format
        para_refused = P.mac_apply(fmt, P.PARAGRAPH_PROPS, para_spec)
        P.mac_apply_line_spacing(fmt, para_spec.get("line_spacing_rule"),
                                 para_spec.get("line_spacing"))

    described = _describe_mac_style(s)
    warnings = [f"{key} was refused: {why}" for key, why in refused.items()]
    for key, wanted in (("base_style", base_style), ("next_style", next_style)):
        if wanted is not None and key not in refused:
            want = _mac_style_name(_style_object(doc, wanted))
            if described.get(key) != want:
                warnings.append(f"{key} reads back as {described.get(key)!r}, not {want!r}.")
    for key, value in (("automatically_update", automatically_update),
                       ("no_space_between_same_style", no_space_between_same_style)):
        if value is not None and key not in refused and described.get(key) != value:
            warnings.append(f"{key} reads back as {described.get(key)!r}, not {value!r}.")
    warnings += P.mismatches(font_spec, described.get("font", {}), font_refused)
    warnings += P.mismatches(para_spec, described.get("paragraph", {}), para_refused)
    out = {"success": True, "created": bool(create), "style": described}
    if warnings:
        out["warnings"] = warnings
    return out


# -- lists -------------------------------------------------------------------------

from . import lists  # noqa: E402

NUMBER_STYLES_K = {
    "bullet": "bullet", "arabic": "arabic", "upper_roman": "uppercase_roman",
    "lower_roman": "lowercase_roman", "upper_letter": "uppercase_letter",
    "lower_letter": "lowercase_letter", "ordinal": "ordinal",
    "cardinal_text": "cardinal_text", "ordinal_text": "ordinal_text",
    "arabic_full_width": "arabic_full_width", "arabic_leading_zero": "arabic_lz",
    "number_in_circle": "number_in_circle", "kanji": "kanji", "kanji_digit": "kanji_digit",
    "kanji_traditional": "kanji_traditional", "aiueo": "aiueo",
    "aiueo_half_width": "aiueo_half_width", "iroha": "iroha",
    "iroha_half_width": "iroha_half_width", "legal": "legal", "none": "none",
}
LIST_TYPES_K = {"list_no_numbering": "none", "list_listnum_only": "listnum_field",
                "list_bullet": "bullet", "list_simple_numbering": "simple_numbering",
                "list_outline_numbering": "outline_numbering",
                "list_mixed_numbering": "mixed_numbering",
                "list_picture_bullet": "picture_bullet"}


def _mac_paragraph_list(doc, index) -> dict:
    fmt = doc.paragraphs[index].text_object.list_format
    kind = LIST_TYPES_K.get(_name(fmt.list_type.get()), _name(fmt.list_type.get()))
    out = {"index": index,
           "text": _clean(doc.paragraphs[index].text_object.content.get())[:80],
           "list": kind}
    if kind != "none":
        level = fmt.list_level_number.get()
        try:
            bullet = _name(fmt.list_template.list_levels[level].number_style.get()) \
                == "list_number_style_bullet"
        except Exception:  # noqa: BLE001
            bullet = False
        out["list"] = lists.list_kind(kind, bullet)
        out["level"] = level
        out["list_string"] = fmt.list_string.get()
    return out


def _mac_describe_level(lv, index) -> dict:
    raw = _name(lv.number_style.get()).removeprefix("list_number_style_")
    style = next((k for k, v in NUMBER_STYLES_K.items() if v == raw), raw)
    out = {"level": index, "number_style": style, "number_format": lv.number_format.get()}
    if style == "bullet":
        out["bullet"] = out["number_format"]
    font = P.mac_read(lv.font_object, P.FONT_PROPS)
    out["font"] = {"name": font.get("latin_font"), "color": font.get("color"),
                   "size": font.get("size"), "bold": font.get("bold")}
    out["number_position"] = lv.number_position.get()
    out["text_position"] = lv.text_position.get()
    out["start_at"] = lv.start_at.get()
    tab = lv.tab_position.get()
    out["tab_position"] = None if tab in (999999, P.WD_UNDEFINED) else tab
    out["trailing"] = _name(lv.trailing_character.get()).removeprefix("trailing_")
    out["alignment"] = _name(lv.list_level_alignment.get()).removeprefix("list_level_align_")
    out["linked_style"] = lv.linked_style.get() or None
    return out


def _mac_set_level(doc, lv, spec, refused) -> None:
    k = _k()

    def put(key, prop, value):
        try:
            getattr(lv, prop).set(value)
        except Exception as exc:  # noqa: BLE001 - reported to the caller
            refused[f"level {spec['level']} {key}"] = str(exc).splitlines()[0]

    # Measured on Word 16.114: setting a level's number style to bullet is
    # refused (-1708) until its number format is the bullet character, and a
    # number format holding %1 is refused until the style is a numbering one.
    # So a bullet's format goes first, and a number's style.
    style = spec.get("number_style")
    steps = [("number_format", "number_format", spec.get("number_format")),
             ("number_style", "number_style",
              None if style is None
              else getattr(k, "list_number_style_" + NUMBER_STYLES_K[style]))]
    if style not in (None, "bullet"):
        steps.reverse()
    for key, prop, value in steps:
        if value is not None:
            put(key, prop, value)
    for key in ("number_position", "text_position", "tab_position", "start_at"):
        if spec.get(key) is not None:
            put(key, key, spec[key])
    if spec.get("trailing") is not None:
        put("trailing", "trailing_character", getattr(k, "trailing_" + spec["trailing"]))
    if spec.get("alignment") is not None:
        put("alignment", "list_level_alignment",
            getattr(k, "list_level_align_" + spec["alignment"]))
    if spec.get("linked_style") is not None:
        put("linked_style", "linked_style",
            _mac_style_name(_style_object(doc, spec["linked_style"])))
    font = spec.get("font") or {}
    font_spec = {k_: v for k_, v in {"latin_font": font.get("name"),
                                     "latin_font_extended": font.get("name"),
                                     "color": font.get("color"), "size": font.get("size"),
                                     "bold": font.get("bold")}.items() if v is not None}
    for name, why in P.mac_apply(lv.font_object, P.FONT_PROPS, font_spec).items():
        refused[f"level {spec['level']} font {name}"] = why


@mac_impl(lists._apply_list_impl)
def apply_list(start, end, levels, level, continue_previous):
    k = _k()
    doc = backend.word.document()
    start, end = _span(doc, start, end)
    specs = [lists._level_spec_defaults(dict(s)) for s in levels]
    outline = len(specs) > 1 or any(s["level"] > 1 for s in specs) or level not in (None, 1)
    # Measured on Word 16.114: `make new list template at` the document works
    # (without `at` it is -1708); changing its levels leaves the galleries and
    # the document's other lists as they were.
    template = doc.make(new=k.list_template, at=doc,
                        with_properties={k.outline_numbered: outline})
    refused: dict = {}
    for spec in specs:
        _mac_set_level(doc, template.list_levels[spec["level"]], spec, refused)
    first, _ = _bounds(doc, start)
    _, last = _bounds(doc, end)
    doc.create_range(start=first, end_=last).list_format.apply_list_format_template(
        list_template=template, continue_previous_list=continue_previous,
        apply_to=k.list_apply_to_whole_list, default_list_behavior=k.word10_list_behavior)
    if level is not None:
        for i in range(start, end + 1):
            doc.paragraphs[i].text_object.list_format.list_level_number.set(level)
    out = [_mac_paragraph_list(doc, i) for i in range(start, end + 1)]
    used = doc.paragraphs[start].text_object.list_format.list_template
    described = [_mac_describe_level(used.list_levels[s["level"]], s["level"]) for s in specs]
    warnings = [f"{key} was refused: {why}" for key, why in refused.items()]
    for spec, got in zip(specs, described):
        for key in ("bullet", "number_format", "number_position", "text_position",
                    "start_at"):
            if spec.get(key) is not None and got.get(key) != spec[key]:
                warnings.append(f"Level {spec['level']} {key} reads back as "
                                f"{got.get(key)!r}, not {spec[key]!r}.")
        wanted_color = (spec.get("font") or {}).get("color")
        if wanted_color and (got["font"].get("color") or "").upper() != \
                "#" + wanted_color.lstrip("#").upper():
            warnings.append(f"Level {spec['level']} font color reads back as "
                            f"{got['font'].get('color')!r}, not {wanted_color!r}.")
    off = [p["index"] for p in out if p["list"] == "none"]
    if off:
        warnings.append(f"Paragraphs {off} read back without a list.")
    _scroll(doc, first, last)
    result = {"success": True, "paragraphs": out, "levels": described}
    if warnings:
        result["warnings"] = warnings
    return result


@mac_impl(lists._set_list_level_impl)
def set_list_level(start, end, level):
    doc = backend.word.document()
    start, end = _span(doc, start, end)
    for i in range(start, end + 1):
        if _mac_paragraph_list(doc, i)["list"] == "none":
            raise ValueError(f"Paragraph {i} is not in a list. Make it one with "
                             "word_apply_list first.")
    for i in range(start, end + 1):
        doc.paragraphs[i].text_object.list_format.list_level_number.set(level)
    out = [_mac_paragraph_list(doc, i) for i in range(start, end + 1)]
    result = {"success": True, "paragraphs": out}
    off = [p["index"] for p in out if p.get("level") != level]
    if off:
        result["warnings"] = [f"Paragraphs {off} read back at another level."]
    return result


@mac_impl(lists._remove_list_impl)
def remove_list(start, end):
    doc = backend.word.document()
    start, end = _span(doc, start, end)
    for i in range(start, end + 1):
        doc.paragraphs[i].text_object.list_format.remove_numbers()
    out = [_mac_paragraph_list(doc, i) for i in range(start, end + 1)]
    result = {"success": True, "paragraphs": out}
    left = [p["index"] for p in out if p["list"] != "none"]
    if left:
        result["warnings"] = [f"Paragraphs {left} are still in a list."]
    return result


@mac_impl(lists._get_list_impl)
def get_list(index):
    doc = backend.word.document()
    _para(doc, index)
    info = _mac_paragraph_list(doc, index)
    if info["list"] == "none":
        return {"success": True, "paragraph": info}
    template = doc.paragraphs[index].text_object.list_format.list_template
    return {"success": True, "paragraph": info,
            "levels": [_mac_describe_level(template.list_levels[i], i) for i in range(1, 10)]}
