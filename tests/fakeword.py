"""A small stand-in for Word's object model, enough to run the COM implementations.

The document is one string of characters, with "\\r" ending each paragraph as
in Word, and one dict of attributes per character. Paragraph formatting and
the style live on a paragraph's mark, as they do in Word, and a character
inserted at a position copies the attributes of the character it lands in
front of. It does not imitate everything Word does; it imitates what the
tools rely on, so a test can check what a tool wrote and what it read back.
"""

from __future__ import annotations

from types import SimpleNamespace

UNDEFINED = 9999999

BUILTIN_NAMES = {-1: "Normal", -2: "Heading 1", -3: "Heading 2", -4: "Heading 3",
                 -5: "Heading 4", -63: "Title", -75: "Subtitle", -180: "List Paragraph"}

DEFAULT_ATTRS = {
    "NameAscii": "Calibri", "NameOther": "Calibri", "NameFarEast": "游明朝",
    "NameBi": "Arial", "Size": 10.5, "SizeBi": 10.5, "Bold": 0, "BoldBi": 0,
    "Italic": 0, "ItalicBi": 0, "Underline": 0, "Color": -16777216,
    "style": "Normal", "Alignment": 0, "LineSpacingRule": 0, "LineSpacing": 12.0,
    "SpaceBefore": 0.0, "SpaceAfter": 0.0, "FirstLineIndent": 0.0, "LeftIndent": 0.0,
    "RightIndent": 0.0, "DisableLineHeightGrid": False, "KeepWithNext": False,
    "PageBreakBefore": False,
}

CHAR_KEYS = ("NameAscii", "NameOther", "NameFarEast", "NameBi", "Size", "SizeBi", "Bold",
             "BoldBi", "Italic", "ItalicBi", "Underline", "Color")
PARA_KEYS = ("Alignment", "LineSpacingRule", "LineSpacing", "SpaceBefore", "SpaceAfter",
             "FirstLineIndent", "LeftIndent", "RightIndent", "DisableLineHeightGrid",
             "KeepWithNext", "PageBreakBefore")

# Fonts this fake refuses as an East Asian font, the way Word keeps the old
# one when the new font has no East Asian glyphs.
LATIN_ONLY_FONTS = {"Segoe UI", "Arial", "Calibri", "Helvetica Neue"}


class FakeStyle:
    def __init__(self, name):
        self.NameLocal = name

    def __eq__(self, other):
        return isinstance(other, FakeStyle) and other.NameLocal == self.NameLocal


class FakeStyles:
    def __init__(self, doc):
        self.doc = doc

    def Item(self, key):
        if isinstance(key, int):
            if key not in BUILTIN_NAMES:
                raise KeyError(key)
            return FakeStyle(BUILTIN_NAMES[key])
        if key in self.doc.custom_styles or key in BUILTIN_NAMES.values():
            return FakeStyle(key)
        raise Exception("The requested member of the collection does not exist.")


def outline_of(style_name):
    if style_name.startswith("Heading "):
        return int(style_name.split()[1])
    return 10


class FakeFont:
    def __init__(self, rng):
        object.__setattr__(self, "_rng", rng)

    def _chars(self):
        rng = self._rng
        attrs = rng.doc.attrs[rng.Start:rng.End]
        return attrs or [rng.doc.attrs[min(rng.Start, len(rng.doc.attrs) - 1)]]

    def __getattr__(self, name):
        if name not in CHAR_KEYS:
            raise AttributeError(name)
        values = {a[name] for a in self._chars()}
        if len(values) == 1:
            return values.pop()
        return "" if name.startswith("Name") else UNDEFINED

    def __setattr__(self, name, value):
        if name not in CHAR_KEYS:
            raise AttributeError(name)
        if name == "NameFarEast" and value in LATIN_ONLY_FONTS:
            return
        if isinstance(value, bool):
            value = -1 if value else 0
        for a in self._chars():
            a[name] = value

    def Reset(self):
        for a in self._chars():
            for key in CHAR_KEYS:
                a[key] = DEFAULT_ATTRS[key]


class FakeParagraphFormat:
    def __init__(self, rng):
        object.__setattr__(self, "_rng", rng)

    def _marks(self):
        return [self._rng.doc.attrs[p.Range.End - 1] for p in self._rng.paragraph_list()]

    def __getattr__(self, name):
        if name not in PARA_KEYS:
            raise AttributeError(name)
        values = {m[name] for m in self._marks()}
        return values.pop() if len(values) == 1 else UNDEFINED

    def __setattr__(self, name, value):
        if name not in PARA_KEYS:
            raise AttributeError(name)
        for m in self._marks():
            m[name] = value


class FakeFind:
    def __init__(self, rng):
        self.rng = rng
        self.Replacement = SimpleNamespace(ClearFormatting=lambda: None)

    def ClearFormatting(self):
        pass

    def Execute(self, FindText, MatchCase=False, MatchWholeWord=False, MatchWildcards=False,
                Forward=True, Wrap=0, Format=False, ReplaceWith=None, Replace=0):
        doc = self.rng.doc
        if Replace == 2:
            start, end = self.rng.Start, self.rng.End
            hit = False
            pos = start
            while True:
                at = doc.text.find(FindText, pos, end)
                if at < 0:
                    break
                doc.replace(at, at + len(FindText), ReplaceWith)
                end += len(ReplaceWith) - len(FindText)
                pos = at + len(ReplaceWith)
                hit = True
            return hit
        at = doc.text.find(FindText, self.rng.Start, self.rng.search_end)
        if at < 0:
            return False
        self.rng.Start, self.rng.End = at, at + len(FindText)
        return True


class FakeRange:
    def __init__(self, doc, start, end):
        self.doc = doc
        self.Start = start
        self.End = end
        self.search_end = doc.end_of_content if end == start else end

    @property
    def Text(self):
        return self.doc.text[self.Start:self.End]

    @Text.setter
    def Text(self, value):
        self.doc.replace(self.Start, self.End, value)
        self.End = self.Start + len(value)

    def InsertAfter(self, value):
        self.doc.insert(self.End, value)
        self.End += len(value)

    def InsertBefore(self, value):
        self.doc.insert(self.Start, value)
        self.End += len(value)

    def Delete(self):
        end = min(self.End, len(self.doc.text) - 1)
        self.doc.replace(self.Start, end, "")
        self.End = self.Start

    def Collapse(self, direction=1):
        if direction == 0:
            self.Start = self.End
        else:
            self.End = self.Start
        self.search_end = self.doc.end_of_content

    @property
    def Font(self):
        return FakeFont(self)

    @property
    def ParagraphFormat(self):
        return FakeParagraphFormat(self)

    @property
    def Find(self):
        if not hasattr(self, "_find"):
            self._find = FakeFind(self)
        return self._find

    def paragraph_list(self):
        return [p for p in self.doc.paragraph_objects()
                if p.Range.End > self.Start and p.Range.Start < max(self.End, self.Start + 1)]

    @property
    def Paragraphs(self):
        items = self.paragraph_list()
        return SimpleNamespace(Count=len(items), Item=lambda i: items[i - 1])

    @property
    def Style(self):
        return FakeStyle(self.doc.attrs[self.paragraph_list()[0].Range.End - 1]["style"])

    @Style.setter
    def Style(self, style):
        for p in self.paragraph_list():
            p.Style = style


class FakeParagraph:
    def __init__(self, doc, start, end, index):
        self.doc, self._start, self._end, self.index = doc, start, end, index

    @property
    def Range(self):
        return FakeRange(self.doc, self._start, self._end)

    @property
    def Style(self):
        return FakeStyle(self.doc.attrs[self._end - 1]["style"])

    @Style.setter
    def Style(self, style):
        self.doc.attrs[self._end - 1]["style"] = style.NameLocal

    @property
    def OutlineLevel(self):
        return outline_of(self.doc.attrs[self._end - 1]["style"])

    @property
    def Format(self):
        return FakeParagraphFormat(self.Range)

    def Next(self, Count=1):
        items = self.doc.paragraph_objects()
        return items[self.index] if self.index < len(items) else None


class FakeWindow:
    def __init__(self):
        self.scrolled_to = []

    def ScrollIntoView(self, rng, start=True):
        self.scrolled_to.append((rng.Start, rng.End))


class FakeDocument:
    def __init__(self, paragraphs=("",), name="Doc1.docx", path="C:\\docs",
                 styles=None):
        self.Name = name
        self.Path = path
        self.FullName = f"{path}\\{name}" if path else name
        self.Saved = True
        self.ReadOnly = False
        self.TrackRevisions = False
        self.text = ""
        self.attrs = []
        for p in paragraphs:
            text, style = (p, "Normal") if isinstance(p, str) else p
            for ch in text + "\r":
                self.text += ch
                self.attrs.append({**DEFAULT_ATTRS, "style": style})
        self.custom_styles = set(styles or ())
        self.Styles = FakeStyles(self)
        self.window = FakeWindow()
        self.undo_log = []

    # -- edits --
    def insert(self, pos, value):
        template = self.attrs[min(pos, len(self.attrs) - 1)]
        self.text = self.text[:pos] + value + self.text[pos:]
        self.attrs[pos:pos] = [dict(template) for _ in value]
        self.Saved = False

    def replace(self, start, end, value):
        template = self.attrs[min(start, len(self.attrs) - 1)]
        self.text = self.text[:start] + value + self.text[end:]
        self.attrs[start:end] = [dict(template) for _ in value]
        self.Saved = False

    @property
    def end_of_content(self):
        return len(self.text)

    # -- object model --
    def paragraph_objects(self):
        out, start = [], 0
        for i, ch in enumerate(self.text):
            if ch == "\r":
                out.append(FakeParagraph(self, start, i + 1, len(out) + 1))
                start = i + 1
        return out

    @property
    def Paragraphs(self):
        items = self.paragraph_objects()

        def item(i):
            if not 1 <= i <= len(items):
                raise IndexError(i)
            return items[i - 1]

        return SimpleNamespace(Count=len(items), Item=item)

    def Range(self, start=0, end=None):
        return FakeRange(self, start, len(self.text) if end is None else end)

    @property
    def Content(self):
        return self.Range(0, len(self.text))

    @property
    def Windows(self):
        return SimpleNamespace(Count=1, Item=lambda i: self.window)

    def texts(self):
        return [p.Range.Text.rstrip("\r") for p in self.paragraph_objects()]

    def styles(self):
        return [p.Style.NameLocal for p in self.paragraph_objects()]


class FakeWord:
    """Stands in for backend.word: runs calls inline, on one document."""

    def __init__(self, doc=None):
        self.doc = doc or FakeDocument()
        self.target_full_name = None
        self.launched_by_us = False
        self.wanted_document = None
        self.calls = []

    def execute(self, func, *args, undo=None, idempotent=False, **kwargs):
        self.calls.append((func.__name__, undo))
        return func(*args, **kwargs)

    def document(self):
        return self.doc
