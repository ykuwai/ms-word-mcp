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

BUILTIN_NAMES = {-1: "Normal", -66: "Default Paragraph Font", -2: "Heading 1", -3: "Heading 2", -4: "Heading 3",
                 -5: "Heading 4", -6: "Heading 5", -7: "Heading 6", -8: "Heading 7",
                 -9: "Heading 8", -10: "Heading 9", -63: "Title", -75: "Subtitle", -180: "List Paragraph"}

DEFAULT_ATTRS = {
    "NameAscii": "Calibri", "NameOther": "Calibri", "NameFarEast": "游明朝",
    "NameBi": "Arial", "Size": 10.5, "SizeBi": 10.5, "Bold": 0, "BoldBi": 0,
    "Italic": 0, "ItalicBi": 0, "Underline": 0, "Color": -16777216,
    "style": "Normal", "Alignment": 0, "LineSpacingRule": 0, "LineSpacing": 12.0,
    "SpaceBefore": 0.0, "SpaceAfter": 0.0, "FirstLineIndent": 0.0, "LeftIndent": 0.0,
    "RightIndent": 0.0, "DisableLineHeightGrid": False, "KeepWithNext": False,
    "PageBreakBefore": False,
    "UnderlineColor": -16777216, "StrikeThrough": 0, "DoubleStrikeThrough": 0,
    "Superscript": 0, "Subscript": 0, "SmallCaps": 0, "AllCaps": 0, "Hidden": 0,
    "Spacing": 0.0, "Scaling": 100, "Kerning": 0.0, "Position": 0.0, "EmphasisMark": 0,
    "DisableCharacterSpaceGrid": False,
    "LineUnitBefore": 0.0, "LineUnitAfter": 0.0, "SpaceBeforeAuto": 0, "SpaceAfterAuto": 0,
    "CharacterUnitFirstLineIndent": 0.0, "CharacterUnitLeftIndent": 0.0,
    "CharacterUnitRightIndent": 0.0, "KeepTogether": 0, "WidowControl": -1,
    "FarEastLineBreakControl": -1, "HangingPunctuation": -1,
    "HalfWidthPunctuationOnTopOfLine": 0, "AddSpaceBetweenFarEastAndAlpha": -1,
    "AddSpaceBetweenFarEastAndDigit": -1, "AutoAdjustRightIndent": -1, "WordWrap": -1,
    "BaseLineAlignment": 4, "ReadingOrder": 1,
}

CHAR_KEYS = ("NameAscii", "NameOther", "NameFarEast", "NameBi", "Size", "SizeBi", "Bold",
             "BoldBi", "Italic", "ItalicBi", "Underline", "Color", "UnderlineColor",
             "StrikeThrough", "DoubleStrikeThrough", "Superscript", "Subscript",
             "SmallCaps", "AllCaps", "Hidden", "Spacing", "Scaling", "Kerning", "Position",
             "EmphasisMark", "DisableCharacterSpaceGrid")
PARA_KEYS = ("Alignment", "LineSpacingRule", "LineSpacing", "SpaceBefore", "SpaceAfter",
             "FirstLineIndent", "LeftIndent", "RightIndent", "DisableLineHeightGrid",
             "KeepWithNext", "PageBreakBefore", "LineUnitBefore", "LineUnitAfter",
             "SpaceBeforeAuto", "SpaceAfterAuto", "CharacterUnitFirstLineIndent",
             "CharacterUnitLeftIndent", "CharacterUnitRightIndent", "KeepTogether",
             "WidowControl", "FarEastLineBreakControl", "HangingPunctuation",
             "HalfWidthPunctuationOnTopOfLine", "AddSpaceBetweenFarEastAndAlpha",
             "AddSpaceBetweenFarEastAndDigit", "AutoAdjustRightIndent", "WordWrap",
             "BaseLineAlignment", "ReadingOrder")

# Fonts this fake refuses as an East Asian font. Word refuses a font without
# East Asian glyphs there and keeps the old one: 0x800A16D4 on Windows, -1708
# on the Mac (both measured on Word 16).
LATIN_ONLY_FONTS = {"Segoe UI", "Arial", "Calibri", "Helvetica Neue"}


class FakeBorder:
    """One side's line, held flat in each paragraph mark or style it covers."""

    DEFAULTS = {"LineStyle": 0, "LineWidth": 8, "Color": -16777216}

    def __init__(self, stores, side):
        object.__setattr__(self, "_stores", stores)
        object.__setattr__(self, "_side", side)

    def __getattr__(self, name):
        return self._stores[0].get(f"border {self._side} {name}", self.DEFAULTS[name])

    def __setattr__(self, name, value):
        for store in self._stores:
            store[f"border {self._side} {name}"] = value


class FakeBorders:
    def __init__(self, stores):
        self._stores = stores

    def Item(self, side):
        return FakeBorder(self._stores, side)


class FakeShading:
    def __init__(self, stores):
        object.__setattr__(self, "_stores", stores)

    @property
    def BackgroundPatternColor(self):
        return self._stores[0].get("shading", -16777216)

    @BackgroundPatternColor.setter
    def BackgroundPatternColor(self, value):
        for store in self._stores:
            store["shading"] = value


class AttrProxy:
    """A Font or ParagraphFormat over one dict, as a style holds them."""

    def __init__(self, data, keys):
        object.__setattr__(self, "_data", data)
        object.__setattr__(self, "_keys", keys)

    def __getattr__(self, name):
        if name == "Duplicate":
            return AttrProxy(dict(self._data), self._keys)
        if name == "Borders":
            return FakeBorders([self._data])
        if name == "Shading":
            return FakeShading([self._data])
        if name not in self._keys:
            raise AttributeError(name)
        return self._data[name]

    def __setattr__(self, name, value):
        if name not in self._keys:
            raise AttributeError(name)
        if name == "NameFarEast" and value in LATIN_ONLY_FONTS:
            raise Exception("OLE error 0x800a16d4")
        if isinstance(value, bool) and name not in ("DisableLineHeightGrid",
                                                    "DisableCharacterSpaceGrid"):
            value = -1 if value else 0
        self._data[name] = value


class FakeStyle:
    def __init__(self, name, type_=1, builtin=False):
        self.NameLocal = name
        self.Type = type_
        self.BuiltIn = builtin
        self.InUse = True
        self.Description = f"{name} description"
        self.QuickStyle = False
        self.Priority = 99
        self.Locked = False
        self.BaseStyle = "" if name == "Normal" else "Normal"
        self.NextParagraphStyle = name
        self.AutomaticallyUpdate = False
        self.NoSpaceBetweenParagraphsOfSameStyle = False
        self._data = dict(DEFAULT_ATTRS)
        self.Font = AttrProxy(self._data, CHAR_KEYS)
        if type_ == 1:
            self.ParagraphFormat = AttrProxy(self._data, PARA_KEYS)

    @property
    def BaseStyle(self):
        return self._base

    @BaseStyle.setter
    def BaseStyle(self, value):
        # Word reads a base style back as a style, however it was set.
        if isinstance(value, int):
            value = BUILTIN_NAMES[value]
        self._base = getattr(value, "NameLocal", value)

    @property
    def Linked(self):
        if self.Type != 1:
            raise Exception("This property cannot be used for this object.")
        return False

    def __eq__(self, other):
        return isinstance(other, FakeStyle) and other.NameLocal == self.NameLocal

    def __hash__(self):
        return hash(self.NameLocal)


class FakeStyles:
    def __init__(self, doc):
        self.doc = doc
        self.defs = {name: FakeStyle(name, builtin=True) for name in BUILTIN_NAMES.values()}
        for name in doc.custom_styles:
            self.defs[name] = FakeStyle(name)

    @property
    def Count(self):
        return len(self.defs)

    def Item(self, key):
        if isinstance(key, int) and key < 0:
            if key not in BUILTIN_NAMES:
                raise KeyError(key)
            return self.defs[BUILTIN_NAMES[key]]
        if isinstance(key, int):
            return list(self.defs.values())[key - 1]
        if key in self.defs:
            return self.defs[key]
        raise Exception("The requested member of the collection does not exist.")

    def Add(self, name, type_=1):
        if name in self.defs:
            raise Exception("That style name already exists.")
        style = FakeStyle(name, type_)
        # Word takes the style and manual formatting of where the cursor is
        # (measured); the fake's cursor is at the start of the document.
        at = self.doc.attrs[0]
        style._data.update({k: at[k] for k in CHAR_KEYS + PARA_KEYS})
        style.BaseStyle = at["style"]
        self.defs[name] = style
        return style


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
            raise Exception("OLE error 0x800a16d4")
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

    @property
    def Borders(self):
        return FakeBorders(self._marks())

    @property
    def Shading(self):
        return FakeShading(self._marks())

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
            # Word turns a multiple of exactly 1, 1.5 or 2 lines into the
            # single, one-and-a-half or double rule (measured on Word 16).
            if (name == "LineSpacing" and m["LineSpacingRule"] == 5
                    and m["LineSpacing"] in (12.0, 18.0, 24.0)):
                m["LineSpacingRule"] = {12.0: 0, 18.0: 1, 24.0: 2}[m["LineSpacing"]]


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
        self.View = SimpleNamespace(Type=3, Zoom=SimpleNamespace(Percentage=100, PageFit=0))

    def ScrollIntoView(self, rng, start=True):
        self.scrolled_to.append((rng.Start, rng.End))


class FakeNote:
    """A footnote or endnote: a one-character mark in the text, and its own text."""

    def __init__(self, notes, at):
        self.notes = notes
        self.at = at
        self.Range = SimpleNamespace(Text="")

    @property
    def Reference(self):
        return FakeRange(self.notes.doc, self.at, self.at + 1)

    @property
    def Index(self):
        return self.notes.ordered().index(self) + 1

    def Delete(self):
        doc = self.notes.doc
        doc.replace(self.at, self.at + 1, "")
        self.notes.items.remove(self)
        for other in doc.Footnotes.items + doc.Endnotes.items:
            if other.at > self.at:
                other.at -= 1


class FakeNotes:
    def __init__(self, doc):
        self.doc = doc
        self.items = []

    def ordered(self):
        return sorted(self.items, key=lambda n: n.at)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, number):
        return self.ordered()[number - 1]

    def Add(self, rng):
        for other in self.doc.Footnotes.items + self.doc.Endnotes.items:
            if other.at >= rng.Start:
                other.at += 1
        self.doc.insert(rng.Start, "\x02")
        note = FakeNote(self, rng.Start)
        self.items.append(note)
        return note


class FakeHyperlink:
    def __init__(self, links, rng, address, sub):
        self.links, self.Range, self._address, self.SubAddress = links, rng, address, sub
        self.ScreenTip = ""

    @property
    def Address(self):
        return self._address

    @Address.setter
    def Address(self, value):
        if value:  # Word keeps the address when it is set to "" (measured)
            self._address = value

    @property
    def TextToDisplay(self):
        return self.Range.Text

    @TextToDisplay.setter
    def TextToDisplay(self, value):
        self.Range.Text = value

    def Delete(self):
        self.links.items.remove(self)


class FakeHyperlinks:
    def __init__(self):
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def Item(self, number):
        return sorted(self.items, key=lambda h: h.Range.Start)[number - 1]

    def Add(self, anchor, address, sub):
        link = FakeHyperlink(self, anchor, address, sub)
        self.items.append(link)
        return link


class FakeBookmark:
    def __init__(self, marks, name, rng):
        self.marks, self.Name, self.Range = marks, name, rng

    @property
    def Start(self):
        return self.Range.Start

    @property
    def Empty(self):
        return self.Range.Start == self.Range.End

    def Delete(self):
        del self.marks.items[self.Name]


class FakeBookmarks:
    """Sorted by name, as Word's own collection is by default."""

    def __init__(self):
        self.items = {}

    @property
    def Count(self):
        return len(self.items)

    def Exists(self, name):
        return name in self.items

    def Item(self, key):
        return self.items[key] if isinstance(key, str) else self.items[sorted(self.items)[key - 1]]

    def Add(self, name, rng):
        self.items[name] = FakeBookmark(self, name, rng)


class FakeDocument:
    def __init__(self, paragraphs=("",), name="Doc1.docx", path="C:\\docs",
                 styles=None):
        self.Footnotes = FakeNotes(self)
        self.Endnotes = FakeNotes(self)
        self.Hyperlinks = FakeHyperlinks()
        self.Bookmarks = FakeBookmarks()
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


# -- lists ---------------------------------------------------------------------

LEVEL_KEYS = ("NumberStyle", "NumberFormat", "NumberPosition", "TextPosition",
              "TabPosition", "TrailingCharacter", "Alignment", "StartAt", "LinkedStyle")


class FakeListLevel:
    def __init__(self, index):
        object.__setattr__(self, "Index", index)
        object.__setattr__(self, "_data", {
            "NumberStyle": 0, "NumberFormat": f"%{index}.", "NumberPosition": 18.0 * (index - 1),
            "TextPosition": 18.0 * index, "TabPosition": 999999, "TrailingCharacter": 0,
            "Alignment": 0, "StartAt": 1, "LinkedStyle": ""})
        object.__setattr__(self, "Font", AttrProxy(dict(DEFAULT_ATTRS), CHAR_KEYS))

    def __getattr__(self, name):
        if name in LEVEL_KEYS:
            return self._data[name]
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name not in LEVEL_KEYS:
            raise AttributeError(name)
        if name == "NumberFormat" and self._data["NumberStyle"] == 23 and len(value) != 1:
            raise Exception("A bullet's format is one character.")
        self._data[name] = value


class FakeListTemplate:
    def __init__(self, outline):
        self.OutlineNumbered = outline
        levels = [FakeListLevel(i) for i in range(1, 10)]
        self.ListLevels = SimpleNamespace(Count=9, Item=lambda i: levels[i - 1])


class FakeListTemplates:
    def __init__(self):
        self.items = []

    def Add(self, outline=False, name=None):
        template = FakeListTemplate(outline)
        self.items.append(template)
        return template


class FakeListFormat:
    """A paragraph's list membership, kept on its mark as (template, level)."""

    def __init__(self, rng):
        self.rng = rng

    def _marks(self):
        return [self.rng.doc.attrs[p.Range.End - 1] for p in self.rng.paragraph_list()]

    @property
    def ListType(self):
        mark = self._marks()[0]
        if mark.get("list") is None:
            return 0
        template, level = mark["list"]
        return 2 if template.ListLevels.Item(level).NumberStyle == 23 else 3

    @property
    def ListTemplate(self):
        mark = self._marks()[0]
        return mark["list"][0] if mark.get("list") else None

    @property
    def ListLevelNumber(self):
        mark = self._marks()[0]
        return mark["list"][1] if mark.get("list") else 1

    @ListLevelNumber.setter
    def ListLevelNumber(self, level):
        for mark in self._marks():
            if mark.get("list"):
                mark["list"] = (mark["list"][0], level)

    @property
    def ListString(self):
        mark = self._marks()[0]
        if not mark.get("list"):
            return ""
        template, level = mark["list"]
        lv = template.ListLevels.Item(level)
        if lv.NumberStyle == 23:
            return lv.NumberFormat
        return lv.NumberFormat.replace(f"%{level}", str(self.ListValue))

    @property
    def ListValue(self):
        mark = self._marks()[0]
        template, level = mark["list"]
        same = [m for m in self.rng.doc.attrs if m.get("list") == mark["list"]]
        at = next(i for i, m in enumerate(same) if m is mark)
        since = next(i for i in range(at, -1, -1) if i == 0 or same[i].get("restart"))
        return template.ListLevels.Item(level).StartAt + at - since

    def ApplyListTemplate(self, ListTemplate, ContinuePreviousList=False, ApplyTo=0,
                          DefaultListBehavior=2):
        if ApplyTo == 1:  # this point forward: the list stays, and counts from here
            self._marks()[0]["restart"] = True
            return
        for mark in self._marks():
            mark["list"] = (ListTemplate, 1)

    def RemoveNumbers(self, NumberType=None):
        for mark in self._marks():
            mark["list"] = None


FakeRange.ListFormat = property(lambda self: FakeListFormat(self))
FakeDocument.ListTemplates = property(
    lambda self: self.__dict__.setdefault("_list_templates", FakeListTemplates()))

# The fake has no equations.
FakeRange.OMaths = property(lambda self: SimpleNamespace(Count=0))
# No tables in a fake document.
FakeDocument.Tables = property(lambda self: SimpleNamespace(Count=0))


def _cross_reference_items(self, kind):
    """Word's heading list: each heading's text, two spaces a level below 1."""
    out = []
    for p in self.paragraph_objects():
        level = p.OutlineLevel
        if level != 10:
            out.append("  " * (level - 1) + p.Range.Text.rstrip("\r"))
    return out


FakeDocument.GetCrossReferenceItems = _cross_reference_items

# A fake document has no tables, so nothing is in one.
FakeRange.Information = lambda self, what: False
