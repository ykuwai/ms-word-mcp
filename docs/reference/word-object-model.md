# Word VBA object model — verified reference for ms-word-mcp

Every fact in this file comes from a page on learn.microsoft.com that was downloaded while writing it (Office VBA reference > Word, plus a few concept, VBA-language and Word interop pages where the VBA reference is silent or missing). Each row names the URL it was taken from. Where a page does not say something, the row says **not stated**. Nothing here is filled from memory.

**Method.** Pages were fetched with `curl` from the exact learn.microsoft.com URLs cited (HTTP 200 checked for each; two 404s are noted where they matter) and their text extracted verbatim, rather than through WebFetch, whose answers are a model's summary of the page. The member lists for Font, ParagraphFormat, Style, Styles and the List* objects were taken from the reference's own table of contents (https://learn.microsoft.com/en-us/office/vba/api/toc.json), so every documented member is listed. Generic members (Application, Creator, Parent) are left out of the per-object tables. Fetched on 2026-11-08.

Conventions: **R/W** = the page says Read/write; **R** = Read-only; **not stated** = the page gives no access mode. In method rows, `Name?` marks an optional parameter. "True" in the VBA pages is the VBA Boolean True, which converts to -1 as a number (https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/boolean-data-type).

## Discrepancies

Where `src/ms_word_mcp` disagrees with the documentation, or relies on something the documentation does not say.

### Wrong according to the docs

1. **`BUILTIN_STYLES["no spacing"] = -158` is wrong.** WdBuiltinStyle gives -158 as `wdStyleNormalObject` ("Normal (applied to an object)."). The page has no member for the "No Spacing" style at all, so no WdBuiltinStyle value can stand in for it. Today `resolve_style(doc, "No Spacing")` asks for `Styles.Item(-158)`, i.e. wdStyleNormalObject, not No Spacing. Look "No Spacing" up by name instead (its name is localized, so it needs another way to find it in a non-English Word). Source: https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle
2. **`_STYLE_TYPES[5] = "linked"` (formatting.py) is wrong.** The VBA WdStyleType page documents only 1 = wdStyleTypeParagraph, 2 = wdStyleTypeCharacter, 3 = wdStyleTypeTable and 4 = wdStyleTypeList (https://learn.microsoft.com/en-us/office/vba/api/word.wdstyletype). The Word interop page for the same enumeration adds 5 = `wdStyleTypeParagraphOnly` and 6 = `wdStyleTypeLinked`, both "Reserved for internal use" (https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdstyletype?view=word-pia). So 5 means paragraph-only, not linked. The documented way to tell whether a style is linked is `Style.Linked` (read-only Boolean: "whether a style is a linked style that can be used for both paragraph and character formatting", https://learn.microsoft.com/en-us/office/vba/api/word.style.linked). As written, `word_list_styles` labels type-5 styles "linked". If Word returns 6 for a style, the tool reports it as the bare number 6, and because 6 is neither "paragraph" nor "linked" in the code's map, the style also loses its `fonts` entry. Whether Style.Type ever returns 5 or 6 is **not stated** (https://learn.microsoft.com/en-us/office/vba/api/word.style.type: "Returns the style type. Read-only WdStyleType.").
3. **`ParagraphFormat.DisableLineHeightGrid`: the docs say the opposite of what the code assumes.** The code sets `DisableLineHeightGrid = not snap_to_grid`, so it treats True as "do not snap to the grid". The VBA page says: "True if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified" (https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.disablelineheightgrid). The interop page uses the same sentence (https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._paragraphformat.disablelineheightgrid?view=word-pia). The property's name agrees with the code, and the page's sentence looks inverted, but the docs cannot settle which is right. **Test it in Word** (set True on a paragraph and look at its grid setting in the Paragraph dialog) before relying on either. Note too that the type is Long, which can be wdUndefined, not Boolean.
4. **Bold and italic read back as -1/0, not true/false.** `Font.Bold` and `Font.Italic` are "Read/write **Long**" and "Returns True, False or wdUndefined" (https://learn.microsoft.com/en-us/office/vba/api/word.font.bold, https://learn.microsoft.com/en-us/office/vba/api/word.font.italic). VBA's True becomes -1 when converted to a number (https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/boolean-data-type). `_format_text_impl` puts `undefined(font.Bold)` and `undefined(font.Italic)` straight into `read_back`, so a caller gets `-1` (bold), `0` (not bold) or `null` (mixed). It never gets `true`/`false`. (The docs do not say how pywin32 turns the Long into a Python value; -1 is what the Long holds.) Convert them with `None if v is None else bool(v)`. All the other True/False/wdUndefined properties behave the same way (Hidden, AllCaps, KeepWithNext, PageBreakBefore, WidowControl, …).

### Relies on something the docs do not state

5. **`follow()` says ScrollIntoView "moves the view only" and leaves the selection alone.** The page says only "Scrolls through the document window so the specified range or shape is displayed in the document window". Whether the selection moves is **not stated** (https://learn.microsoft.com/en-us/office/vba/api/word.window.scrollintoview). The signature is `ScrollIntoView(Obj, Start)`. Start is an optional Boolean, default True, meaning the upper-left of the range is placed at the upper-left of the window. The code's positional `True` matches that. The page also says it "cannot be used with outline view".
6. **`undefined()` assumes SizeBi, Color, Alignment, LineSpacingRule and the spacing values return wdUndefined on mixed ranges.** Font.Size is covered by its page's example: `If mySel = wdUndefined Then MsgBox "there is a mix of font sizes in the selection."` (https://learn.microsoft.com/en-us/office/vba/api/word.font.size). The pages for Font.SizeBi, ParagraphFormat.Alignment, LineSpacingRule, LineSpacing, SpaceBefore and SpaceAfter, examples included, do **not state** what they return for a mixed range (https://learn.microsoft.com/en-us/office/vba/api/word.font.sizebi, https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.alignment, https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.linespacingrule). Neither does the archived Font.Color page. The properties that do document wdUndefined are the True/False Long ones: Bold, Italic, Hidden, AllCaps and the rest on Font, and KeepWithNext, KeepTogether, PageBreakBefore, WidowControl, Hyphenation, MirrorIndents, SpaceBeforeAuto/SpaceAfterAuto and the East Asian ones on ParagraphFormat. wdUndefined itself is 9999999 (WdConstants, https://learn.microsoft.com/en-us/office/vba/api/word.wdconstants). The page has no separate WdUndefined enumeration. The code's extra checks for `-9999999` and `""` are not documented values.
7. **NameOther is set but never read back.** Font.NameAscii covers "Latin text (characters with character codes from 0 (zero) through 127)" (https://learn.microsoft.com/en-us/office/vba/api/word.font.nameascii). Font.NameOther covers "characters with character codes from 128 through 255" (https://learn.microsoft.com/en-us/office/vba/api/word.font.nameother). The code sets both to `latin_font`, which is a sensible reading of "Latin". But `fonts_of()` reads back only NameAscii, so if NameOther does not take, no warning is raised. Font.Name is described only as "Returns or sets the name of the specified object" (https://learn.microsoft.com/en-us/office/vba/api/word.font.name). The NameAscii and NameOther pages add only "Use the Name property to change the font that's applied to all text". Which of the four slots Name writes is **not stated**.
8. **Font.Color has no page in the current Word VBA reference.** `https://learn.microsoft.com/en-us/office/vba/api/word.font.color` returns 404 and the table of contents has no entry for it. The archived Word 2007 page (https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb238769(v=office.12)) and the interop page (https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._font.color?view=word-pia) say it "Returns or sets the 24-bit color for the specified Font object", taking a WdColor constant or RGB(). The WdColor table confirms the code's byte order: wdColorRed = 255 and wdColorBlue = 16711680, i.e. value = R + G×256 + B×65536 (https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor). So `_bgr`/`_hex` are right. wdColorAutomatic = -16777216, which `_hex` turns into `None`, the same answer it gives for a mixed range (wdUndefined), so a read-back cannot tell "automatic" from "mixed". What Font.Color returns for a theme colour is **not stated**.
9. **Undo grouping when the target is not the active document.** The UndoRecord concept page says: "Avoid switching documents inside a custom undo record … Word terminates the custom undo record when the code begins to write to the second document". It also says that without EndCustomRecord, "Word will attempt to determine where to end the record" (https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-the-undorecord-object). The server writes to a document that may not be ActiveDocument. Whether that counts as switching documents is **not stated**. The rest of `run_call` matches the docs: the name is limited to 64 characters and truncated by Word anyway (https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.startcustomrecord). Nested StartCustomRecord calls are allowed, with CustomRecordLevel counting them and CustomRecordName giving the first one's name (https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.customrecordlevel, https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.customrecordname). So the `IsRecordingCustomRecord` guard is a choice, not a requirement. No page says whether UndoRecord exists on Word for Mac: **not stated**.
10. **`index_of()` assumes Range.Paragraphs counts a paragraph the range covers only partly.** The page says only "all the paragraphs in the specified range". Partly covered paragraphs are **not stated** (https://learn.microsoft.com/en-us/office/vba/api/word.range.paragraphs).
11. **`_save_document_impl` takes an empty `Document.Path` to mean "never saved".** The Path page does not say what Path returns for an unsaved document (**not stated**, https://learn.microsoft.com/en-us/office/vba/api/word.document.path). The guard itself is well-founded: Document.Save "If a document has not been saved before, the Save As dialog box prompts the user" (https://learn.microsoft.com/en-us/office/vba/api/word.document.save).
12. **`Styles.Item(WdBuiltinStyle)` is shown in an example, not listed on the Item page.** Styles.Item lists the index as "a Long indicating the ordinal position or a String representing the name" (https://learn.microsoft.com/en-us/office/vba/api/word.styles.item). Passing a negative WdBuiltinStyle constant is documented by the Style.BuiltIn page's example `ActiveDocument.Styles(wdStyleHeading1)` (https://learn.microsoft.com/en-us/office/vba/api/word.style.builtin), and by Paragraph.Style and Range.Style, which accept "a WdBuiltinStyle constant".
13. **Three more assumptions with no supporting page.** (a) `text.py` writes `\v` (Chr(11)) for a line break inside a paragraph. No fetched page documents Chr(11) as a manual line break. The InsertAfter page lists only "vbCr, vbLf, vbCrLf and vbTab" (https://learn.microsoft.com/en-us/office/vba/api/word.range.insertafter). (b) `_delete_paragraphs_impl` says "The document's last paragraph mark cannot go". The only related statement is InsertAfter's: text added to a range ending at the document's final mark goes before that mark. That the mark cannot be deleted is **not stated**. (c) `max_length=255` on find/replace text: no length limit is stated on the Find.Execute, Find.Text or Replacement pages (https://learn.microsoft.com/en-us/office/vba/api/word.find.execute, https://learn.microsoft.com/en-us/office/vba/api/word.find.text).

### Checked and correct

- **12 points per line for wdLineSpaceMultiple is correct.** LineSpacing is "in points", and "wdLineSpaceMultiple a LineSpacing property value must be specified, in points" (https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.linespacing). LinesToPoints: "1 line = 12 points" (https://learn.microsoft.com/en-us/office/vba/api/word.application.linestopoints). So 1.5 lines = 18 pt, as `_lines_to_points` computes.
- Every value in `ALIGNMENTS`, `LINE_SPACING`, `SAVE_FORMATS` and the `WD_*` constants, the other 19 entries of `BUILTIN_STYLES`, the Close SaveChanges values -1/0 and `ExportFormat=17` match the documented values. See section B.1.
- The keyword names in `Find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, Forward, Wrap, Format, ReplaceWith, Replace)`, `Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, Visible)`, `Documents.Add(Template)`, `SaveAs2(FileName, FileFormat)`, `ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport)` and `Close(SaveChanges)` all match the pages exactly. See section A.
- `font.Underline = 1 / 0`: wdUnderlineSingle = 1 and wdUnderlineNone = 0 (https://learn.microsoft.com/en-us/office/vba/api/word.wdunderline).
- Searching through a Range leaves the selection alone and redefines the Range to the match, as `_find_all` assumes: "If you've gotten to the Find object from the Range object, the selection isn't changed … but the Range object is redefined" (https://learn.microsoft.com/en-us/office/vba/api/word.find).

## A. Members the code uses

Every Word member and constant that `tools/_word.py`, `tools/text.py`, `tools/formatting.py`, `tools/documents.py`, `tools/edit.py`, `tools/_common.py` and `backend/com.py` touch. The constants are in B.1.

### A.1 Application, Documents

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Application.Name | String | R | "Returns the name of the specified object." com.py reads it to check that Word responds. | https://learn.microsoft.com/en-us/office/vba/api/word.application.name |
| Application.Version | String | R | "Returns the Microsoft Word version number." | https://learn.microsoft.com/en-us/office/vba/api/word.application.version |
| Application.Build | String | R | "Returns the version and build number of the Word application." | https://learn.microsoft.com/en-us/office/vba/api/word.application.build |
| Application.Visible | Boolean | R/W | "determines whether the object is visible". | https://learn.microsoft.com/en-us/office/vba/api/word.application.visible |
| Application.ActiveDocument | Document | R | "If there are no documents open, an error occurs." A document in Protected View cannot be reached this way. `document()` checks Documents.Count first, which avoids the error. | https://learn.microsoft.com/en-us/office/vba/api/word.application.activedocument |
| Application.Documents | Documents | R | "all the open documents". A document in a Protected View window is not a member. | https://learn.microsoft.com/en-us/office/vba/api/word.application.documents |
| Application.UndoRecord | UndoRecord | R | "Returns an UndoRecord object that provides a custom entry point into the undo stack." Platform availability: not stated. | https://learn.microsoft.com/en-us/office/vba/api/word.application.undorecord |
| Documents.Count | Long | R | Number of documents in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.documents.count |
| Documents.Item(Index) | Index: Long or String → Document | method | "One-based index of the object to be returned (Long), or the name of the object (String)." The code loops 1..Count, which matches. | https://learn.microsoft.com/en-us/office/vba/api/word.documents.item |
| Documents.Open(...) | (FileName, ConfirmConversions?, ReadOnly?, AddToRecentFiles?, PasswordDocument?, PasswordTemplate?, Revert?, WritePasswordDocument?, WritePasswordTemplate?, Format?, Encoding?, Visible?, OpenConflictDocument?, OpenAndRepair?, DocumentDirection?, NoEncodingDialog?) → Document | method | The code's keywords FileName, ConfirmConversions, ReadOnly, AddToRecentFiles and Visible match. ReadOnly "doesn't override the read-only recommended setting". Visible: "True if the document is opened in a visible window. The default value is True." Revert: if FileName is already open, False activates the open document. The code avoids this case by checking open documents first. | https://learn.microsoft.com/en-us/office/vba/api/word.documents.open |
| Documents.Add(...) | (Template?, NewTemplate?, DocumentType?, Visible?) → Document | method | Template: "If this argument is omitted, the Normal template is used." The code's keyword Template matches. | https://learn.microsoft.com/en-us/office/vba/api/word.documents.add |

### A.2 Document

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Name | String | R | "Returns the name of the specified object." | https://learn.microsoft.com/en-us/office/vba/api/word.document.name |
| FullName | String | R | "name of a document, including the path"; "equivalent to using the Path, PathSeparator, and Name properties in sequence". | https://learn.microsoft.com/en-us/office/vba/api/word.document.fullname |
| Path | String | R | "disk or Web path to the document", with no trailing separator except at a network-drive root. The value for a never-saved document is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.document.path |
| Saved | Boolean | R/W | "True if the specified document or template has not changed since it was last saved." | https://learn.microsoft.com/en-us/office/vba/api/word.document.saved |
| ReadOnly | Boolean | R | "True if changes to the document cannot be saved to the original document." | https://learn.microsoft.com/en-us/office/vba/api/word.document.readonly |
| TrackRevisions | Boolean | R/W | "True if changes are tracked in the specified document." | https://learn.microsoft.com/en-us/office/vba/api/word.document.trackrevisions |
| Content | Range | R | "a Range object that represents the main document story"; equivalent to StoryRanges(wdMainTextStory). | https://learn.microsoft.com/en-us/office/vba/api/word.document.content |
| Range(Start, End) | (Start?: Variant, End?: Variant) → Range | method | "Returns a Range object by using the specified starting and ending character positions." | https://learn.microsoft.com/en-us/office/vba/api/word.document.range |
| Paragraphs | Paragraphs | R | All the paragraphs in the document. The Paragraphs object page adds: "The Count property for this collection in a document returns the number of items in the main story only." | https://learn.microsoft.com/en-us/office/vba/api/word.document.paragraphs |
| Paragraphs.Count | Long | R | Number of paragraphs in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphs.count |
| Paragraphs.Item(Index) | Index: Long → Paragraph | method | "a Long indicating the ordinal position". The 1-based start is shown by the examples (`Paragraphs(1)`) on the Paragraphs page. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphs.item |
| Tables | Tables | R | All the tables in the document. Tables.Count is a Long, read-only (https://learn.microsoft.com/en-us/office/vba/api/word.tables.count). | https://learn.microsoft.com/en-us/office/vba/api/word.document.tables |
| Sections | Sections | R | The sections in the document. Sections.Count is a Long, read-only (https://learn.microsoft.com/en-us/office/vba/api/word.sections.count). | https://learn.microsoft.com/en-us/office/vba/api/word.document.sections |
| PageSetup | PageSetup | not stated | "Returns a PageSetup object that is associated with the specified document." | https://learn.microsoft.com/en-us/office/vba/api/word.document.pagesetup |
| Styles | Styles | R | "Returns a Styles collection for the specified document." | https://learn.microsoft.com/en-us/office/vba/api/word.document.styles |
| Windows | Windows | R | "all windows for the specified document". Windows.Count: Long, R (https://learn.microsoft.com/en-us/office/vba/api/word.windows.count). Windows.Item(Index): "a Long indicating the ordinal position or a String representing the name" (https://learn.microsoft.com/en-us/office/vba/api/word.windows.item). | https://learn.microsoft.com/en-us/office/vba/api/word.document.windows |
| ComputeStatistics(Statistic, IncludeFootnotesAndEndnotes?) | Statistic: WdStatistic → Long | method | IncludeFootnotesAndEndnotes defaults to False. The code uses wdStatisticWords 0, wdStatisticPages 2 and wdStatisticCharacters 3, all correct. | https://learn.microsoft.com/en-us/office/vba/api/word.document.computestatistics |
| Undo(Times?) | Times?: Variant → Boolean | method | "Undoes the last action or a sequence of actions … Returns True if the actions were successfully undone." What it returns when fewer than Times steps exist is **not stated**. The code calls Undo(1) in a loop, which avoids the question. | https://learn.microsoft.com/en-us/office/vba/api/word.document.undo |
| Redo(Times?) | Times?: Variant → Boolean | method | "Returns True if the actions were redone successfully." | https://learn.microsoft.com/en-us/office/vba/api/word.document.redo |
| UndoClear() | no parameters | method | "Clears the list of actions that can be undone". Not used by the code; listed here for reference. | https://learn.microsoft.com/en-us/office/vba/api/word.document.undoclear |
| Save() | no parameters | method | "If a document has not been saved before, the Save As dialog box prompts the user for a file name." The code refuses an unsaved document, which avoids this. | https://learn.microsoft.com/en-us/office/vba/api/word.document.save |
| SaveAs2(...) | (FileName?, FileFormat?, LockComments?, Password?, AddToRecentFiles?, WritePassword?, ReadOnlyRecommended?, EmbedTrueTypeFonts?, SaveNativePictureFormat?, SaveFormsData?, SaveAsAOCELetter?, Encoding?, InsertLineBreaks?, AllowSubstitutions?, LineEnding?, AddBiDiMarks?, CompatibilityMode?) → Nothing | method | FileFormat: "Can be any WdSaveFormat constant." FileName: "If a document with the specified file name already exists, the document is overwritten without prompting the user." The code checks `overwrite` itself, which is needed. CompatibilityMode defaults to 0, keeping the current mode. The code's keywords FileName and FileFormat match. | https://learn.microsoft.com/en-us/office/vba/api/word.saveas2 |
| ExportAsFixedFormat(...) | (OutputFileName: String, ExportFormat: WdExportFormat, OpenAfterExport?: Boolean, OptimizeFor?, Range?, From?, To?, Item?, IncludeDocProps?, KeepIRM?, CreateBookmarks?, DocStructureTags?, BitmapMissingFonts?, UseISO19005_1?, FixedFormatExtClassPtr?) | method | "Saves a document in PDF or XPS format." The code's keywords OutputFileName, ExportFormat (17 = wdExportFormatPDF) and OpenAfterExport match. | https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat |
| Close(SaveChanges?, OriginalFormat?, RouteDocument?) | SaveChanges: WdSaveOptions | method | "Can be one of the following WdSaveOptions constants: wdDoNotSaveChanges, wdPromptToSaveChanges, or wdSaveChanges." The code's -1 / 0 = wdSaveChanges / wdDoNotSaveChanges is correct. | https://learn.microsoft.com/en-us/office/vba/api/word.document.close(method) |
| PageSetup.PageWidth / PageHeight | Single (points) | R/W | Setting either one changes PaperSize to wdPaperCustom. | https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.pagewidth , https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.pageheight |
| PageSetup.TopMargin / BottomMargin / LeftMargin / RightMargin | Single (points) | R/W | Distance from the page edge to the body text. With MirrorMargins, Left is the inside margin and Right the outside. | https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.topmargin , https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.bottommargin , https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.leftmargin , https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.rightmargin |

### A.3 Window, UndoRecord

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Window.ScrollIntoView(Obj, Start?) | Obj: Range or Shape (Object, required); Start: Boolean, default True | method | "Scrolls through the document window so the specified range or shape is displayed in the document window." Start=True puts the upper-left corner of the range at the upper-left of the window; False puts the lower-right corner at the lower-right. "This method cannot be used with outline view." Whether it moves the selection is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.window.scrollintoview |
| UndoRecord.StartCustomRecord(Name?) | Name?: String | method | "This string is limited to 64 characters. If a longer string is supplied, the string is truncated to 64 characters." If Name is omitted or empty, "Word uses the name of the first command executed". It records "all actions done to the application while it is active". | https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.startcustomrecord |
| UndoRecord.EndCustomRecord() | no parameters | method | "Completes the creation of a custom undo record." | https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.endcustomrecord |
| UndoRecord.IsRecordingCustomRecord | Boolean | R | "whether a custom undo action is being recorded". | https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.isrecordingcustomrecord |
| UndoRecord.CustomRecordLevel | Long | R | "the number of custom undo action calls that are currently active"; 0 when none is active. | https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.customrecordlevel |
| UndoRecord.CustomRecordName | String | R | With nested StartCustomRecord calls, "the string specified by the first call will be returned"; an empty string when none is active. | https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.customrecordname |
| UndoRecord (concept page) | — | — | "Close each custom undo record with a call to EndCustomRecord." "Avoid switching documents inside a custom undo record." Calling Document.Undo inside a record undoes the record's earlier actions. Breakpoints in Debug mode end all records. Platform (Mac): **not stated**. | https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-the-undorecord-object |

### A.4 Range, Paragraph, Find

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Range.Start | Long | R/W | Character position from the start of the story. The main text story starts at 0. Setting Start past End also moves End. | https://learn.microsoft.com/en-us/office/vba/api/word.range.start |
| Range.End | Long | R/W | The ending position. Setting End before Start moves Start to the same value. | https://learn.microsoft.com/en-us/office/vba/api/word.range.end |
| Range.Text | String | R/W | "returns the plain, unformatted text of the range. When you set this property, the existing text in the range is replaced." | https://learn.microsoft.com/en-us/office/vba/api/word.range.text |
| Range.InsertAfter(Text) | Text: String | method | "After this method is applied, the range expands to include the new text." A range covering a whole paragraph gets the text after its paragraph mark. If the range ends at the document's final paragraph mark, the text goes before that mark. A bookmark range does not expand. | https://learn.microsoft.com/en-us/office/vba/api/word.range.insertafter |
| Range.Delete(Unit?, Count?) | Unit?: WdUnits; Count?: Variant → Long | method | "returns a Long value that indicates the number of items deleted, or it returns 0 (zero) if the deletion was unsuccessful." The code ignores the return value. | https://learn.microsoft.com/en-us/office/vba/api/word.range.delete |
| Range.Collapse(Direction?) | Direction?: WdCollapseDirection, default wdCollapseStart | method | Collapsing a whole-paragraph range to its end lands after the paragraph mark (the start of the next paragraph). | https://learn.microsoft.com/en-us/office/vba/api/word.range.collapse |
| Range.Find | Find | R | "Returns a Find object that contains the criteria for a find operation." | https://learn.microsoft.com/en-us/office/vba/api/word.range.find |
| Range.Paragraphs | Paragraphs | R | "all the paragraphs in the specified range". Partly covered paragraphs: **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.range.paragraphs |
| Range.Style | Variant | R/W | Set with "the local name of the style, an integer, a WdBuiltinStyle constant, or an object that represents the style." "When you return the style for a range that includes more than one style, only the first character or paragraph style is returned." | https://learn.microsoft.com/en-us/office/vba/api/word.range.style |
| Range.Font | Font | R/W | "the character formatting of the specified object". | https://learn.microsoft.com/en-us/office/vba/api/word.range.font |
| Range.ParagraphFormat | ParagraphFormat | R/W | "the paragraph settings for the specified range". | https://learn.microsoft.com/en-us/office/vba/api/word.range.paragraphformat |
| Paragraph.Range | Range | not stated | "the portion of a document that is contained within the specified paragraph". The Paragraph.Range page does not say whether this includes the paragraph mark (see E). | https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.range |
| Paragraph.Format | ParagraphFormat | not stated | "Returns or sets a ParagraphFormat object that represents the formatting of the specified paragraph". | https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.format |
| Paragraph.Style | Variant | R/W | Set with a local name, an integer, a WdBuiltinStyle constant or a Style object. Unlike Range.Style, the page says nothing about what it returns. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.style |
| Paragraph.OutlineLevel | WdOutlineLevel | R/W | "If a paragraph has a heading style applied to it (Heading 1 through Heading 9), the outline level is the same as the heading style and cannot be changed." Body text = 10. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.outlinelevel |
| Paragraph.Next(Count?) | Count?: Variant, default 1 → Paragraph | method | "Returns a Paragraph object that represents the next paragraph." What it returns after the last paragraph is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.next |
| Find.ClearFormatting() | no parameters | method | "Removes text and paragraph formatting from the text specified in a find or replace operation." | https://learn.microsoft.com/en-us/office/vba/api/word.find.clearformatting |
| Find.Replacement | Replacement | not stated | "the criteria for a replace operation". | https://learn.microsoft.com/en-us/office/vba/api/word.find.replacement |
| Replacement.ClearFormatting() | no parameters | method | Removes formatting from the replacement criteria. | https://learn.microsoft.com/en-us/office/vba/api/word.replacement.clearformatting |
| Find.Execute(...) | (FindText?, MatchCase?, MatchWholeWord?, MatchWildcards?, MatchSoundsLike?, MatchAllWordForms?, Forward?, Wrap?, Format?, ReplaceWith?, Replace?, MatchKashida?, MatchDiacritics?, MatchAlefHamza?, MatchControl?, MatchPrefix?, MatchSuffix?, MatchPhrase?, IgnoreSpace?, IgnorePunct?) → Boolean | method | "Returns True if the find operation is successful." Wrap is a WdFindWrap and also governs what happens when the text is not found in a range. Replace is a WdReplace. Replacement text "inherits the formatting of the text it replaces". With MatchCase False, uppercase matches get an uppercase replacement. The code's keyword names all match. | https://learn.microsoft.com/en-us/office/vba/api/word.find.execute |
| Find (object page) | — | — | "If you've gotten to the Find object from the Range object, the selection isn't changed when text matching the find criteria is found, but the Range object is redefined." | https://learn.microsoft.com/en-us/office/vba/api/word.find |

### A.5 Font (members the code uses)

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Name | String | R/W | Returns or sets the name of the specified object. | https://learn.microsoft.com/en-us/office/vba/api/word.font.name |
| NameAscii | String | R/W | Returns or sets the font used for Latin text (characters with character codes from 0 (zero) through 127). — Remarks: In the U.S. English version of Microsoft Word, the default value of this property is Times New Roman. Use the Name property to change the font that's applied to all text and that appears in the Font box on the Formatting toolbar. | https://learn.microsoft.com/en-us/office/vba/api/word.font.nameascii |
| NameOther | String | R/W | Returns or sets the font used for characters with character codes from 128 through 255. — Remarks: In the U.S. English version of Microsoft Word, the default value of this property is Times New Roman. Use the Name property to change the font that's applied to all text and that appears in the Font box on the Formatting toolbar. | https://learn.microsoft.com/en-us/office/vba/api/word.font.nameother |
| NameFarEast | String | R/W | Returns or sets an East Asian font name. — Remarks: In the U.S. English version of Microsoft Word, the default value of this property is Times New Roman. This is the recommended way to return or set the font for Asian text in a document created in an Asian version of Word. | https://learn.microsoft.com/en-us/office/vba/api/word.font.namefareast |
| NameBi | String | R/W | Returns or sets the name of the font in a right-to-left language document. | https://learn.microsoft.com/en-us/office/vba/api/word.font.namebi |
| Size | Single | R/W | Returns or sets the font size, in points. | https://learn.microsoft.com/en-us/office/vba/api/word.font.size |
| SizeBi | Single | R/W | Returns or sets the font size in points. — Remarks: The SizeBi property applies to text in a right-to-left language. | https://learn.microsoft.com/en-us/office/vba/api/word.font.sizebi |
| Bold | Long | R/W | True if the font is formatted as bold. — Remarks: Returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. | https://learn.microsoft.com/en-us/office/vba/api/word.font.bold |
| BoldBi | Long | R/W | True if the font is formatted as bold. — Remarks: This property returns True, False or wdUndefined (for a mixture of bold and non-bold text). Can be set to True, False, or wdToggle. The BoldBi property applies to text in a right-to-left language. | https://learn.microsoft.com/en-us/office/vba/api/word.font.boldbi |
| Italic | Long | R/W | True if the font or range is formatted as italic. — Remarks: This property returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. | https://learn.microsoft.com/en-us/office/vba/api/word.font.italic |
| ItalicBi | Long | R/W | True if the font or range is formatted as italic. — Remarks: This property returns True, False or wdUndefined (for a mixture of italic and non-italic text) and can be set to True, False, or wdToggle. Use the ItalicBi property for right-to-left languages. | https://learn.microsoft.com/en-us/office/vba/api/word.font.italicbi |
| Underline | WdUnderline | R/W | Returns or sets the type of underline applied to the font. | https://learn.microsoft.com/en-us/office/vba/api/word.font.underline |
| Reset() | no parameters | method | Removes manual character formatting (formatting not applied using a style). For example, if you manually format a word as bold and the underlying style is plain text (not bold), the Reset method removes the bold format. | https://learn.microsoft.com/en-us/office/vba/api/word.font.reset |
| Color | WdColor (24-bit) | R/W (interop: get; set) | No current VBA page: word.font.color returns 404. "Returns or sets the 24-bit color for the specified Font object. Can be any valid WdColor constant or a value returned by Visual Basic's RGB function." (archived Word 2007 page). Interop: `WdColor Color { get; set; }`. | https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb238769(v=office.12) , https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._font.color?view=word-pia |

Font.Size on a mixed range: the page's example tests `If mySel = wdUndefined Then` for "a mix of font sizes" (https://learn.microsoft.com/en-us/office/vba/api/word.font.size).

### A.6 ParagraphFormat (members the code uses)

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Alignment | WdParagraphAlignment | R/W | Returns or sets a WdParagraphAlignment constant that represents the alignment for the specified paragraphs. — Remarks: Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.alignment |
| LineSpacingRule | WdLineSpacing | R/W | Returns or sets the line spacing for the specified paragraph formatting. — Remarks: Use wdLineSpaceSingle, wdLineSpace1pt5, or wdLineSpaceDouble to set the line spacing to one of these values. To set the line spacing to an exact number of points or to a multiple number of lines, you must also set the LineSpacing property. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.linespacingrule |
| LineSpacing | Single | R/W | Returns or sets the line spacing (in points) for the specified paragraphs. — Remarks: Use the LinesToPoints method to convert a number of lines to the corresponding value in points. For example, LinesToPoints(2) returns the value 24. The LineSpacing property can be set after the LineSpacingRule property has been set to: wdLineSpaceAtLeast the line spacing can be greater than or equal to, but never less than, the specified LineSpacing value. wdLineSpaceExactly the line spacing never changes from the specified LineSpacing value, even if a larger font is used within the paragraph. wdLineSpaceMultiple a LineSpacing property value must be specified, in points. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.linespacing |
| SpaceBefore | Single | R/W | Returns or sets the spacing (in points) before the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spacebefore |
| SpaceAfter | Single | R/W | Returns or sets the amount of spacing (in points) after the specified paragraph or text column. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spaceafter |
| FirstLineIndent | Single | R/W | Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.firstlineindent |
| LeftIndent | Single | R/W | Returns or sets a Single that represents the left indent value (in points) for the specified paragraph formatting. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.leftindent |
| RightIndent | Single | R/W | Returns or sets the right indent (in points) for the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.rightindent |
| DisableLineHeightGrid | Long | R/W | True if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. Returns wdUndefined if the DisableLineHeightGrid property is set to True for only some of the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.disablelineheightgrid |
| KeepWithNext | Long | R/W | True if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. — Remarks: This property can be True, False, or wdUndefined. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.keepwithnext |
| PageBreakBefore | Long | R/W | True if a page break is forced before the specified paragraphs. Can be True, False, or wdUndefined. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.pagebreakbefore |
| Application.LinesToPoints(Lines) | Lines: Single → Single | method | "Converts a measurement from lines to points (1 line = 12 points)." | https://learn.microsoft.com/en-us/office/vba/api/word.application.linestopoints |

### A.7 Style, Styles (members the code uses)

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| NameLocal | String | R/W | Returns the name of a built-in style in the language of the user. — Remarks: Setting this property renames a user-defined style or adds an alias to a built-in style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.namelocal |
| InUse | Boolean | R | True if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document. — Remarks: The InUse property doesn't necessarily indicate whether the style is currently applied to any text in the document. For instance, if text that's been formatted with a style is deleted, the InUse property of the style remains True. For built-in styles that hav… | https://learn.microsoft.com/en-us/office/vba/api/word.style.inuse |
| Type | WdStyleType | R | Returns the style type. | https://learn.microsoft.com/en-us/office/vba/api/word.style.type |
| BuiltIn | Boolean | R | True if the specified object is one of the built-in styles or caption labels in Word. — Remarks: You can specify built-in styles across all languages by using the WdBuiltinStyle constants or within a language by using the style name for the language version of Word. For example, if you specify U.S. English in your Microsoft Office language settings, the… | https://learn.microsoft.com/en-us/office/vba/api/word.style.builtin |
| Font | Font | R/W | Returns or sets a Font object that represents the character formatting of the specified style. — Remarks: To set this property, specify an expression that returns a Font object. | https://learn.microsoft.com/en-us/office/vba/api/word.style.font |
| Item() | (Index: Variant) → Style | method | Returns an individual Style object in a collection. — Params: Index: The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual obj… | https://learn.microsoft.com/en-us/office/vba/api/word.styles.item |
| Count | Long | R | Returns a Long that represents the number of styles in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.styles.count |

## B. Constants

### B.1 The code's constants against the documentation

| Code | Code value | Documented member | Documented value | Match | Source URL |
|---|---|---|---|---|---|
| `WD_UNDEFINED` | 9999999 | WdConstants.wdUndefined | 9999999 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdconstants |
| `WD_CHARACTER` | 1 | WdUnits.wdCharacter | 1 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdunits |
| `WD_COLLAPSE_END` | 0 | WdCollapseDirection.wdCollapseEnd | 0 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdcollapsedirection |
| `WD_FIND_STOP` | 0 | WdFindWrap.wdFindStop | 0 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdfindwrap |
| `WD_REPLACE_ALL` | 2 | WdReplace.wdReplaceAll | 2 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdreplace |
| `WD_OUTLINE_BODY` | 10 | WdOutlineLevel.wdOutlineLevelBodyText | 10 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdoutlinelevel |
| `WD_STATISTIC_WORDS` | 0 | WdStatistic.wdStatisticWords | 0 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdstatistic |
| `WD_STATISTIC_PAGES` | 2 | WdStatistic.wdStatisticPages | 2 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdstatistic |
| `WD_STATISTIC_CHARACTERS` | 3 | WdStatistic.wdStatisticCharacters | 3 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdstatistic |
| `BUILTIN_STYLES["normal"]` | -1 | WdBuiltinStyle.wdStyleNormal | -1 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["heading 1"]` | -2 | WdBuiltinStyle.wdStyleHeading1 | -2 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["heading 2"]` | -3 | WdBuiltinStyle.wdStyleHeading2 | -3 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["heading 3"]` | -4 | WdBuiltinStyle.wdStyleHeading3 | -4 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["heading 4"]` | -5 | WdBuiltinStyle.wdStyleHeading4 | -5 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["heading 5"]` | -6 | WdBuiltinStyle.wdStyleHeading5 | -6 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["heading 6"]` | -7 | WdBuiltinStyle.wdStyleHeading6 | -7 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["heading 7"]` | -8 | WdBuiltinStyle.wdStyleHeading7 | -8 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["heading 8"]` | -9 | WdBuiltinStyle.wdStyleHeading8 | -9 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["heading 9"]` | -10 | WdBuiltinStyle.wdStyleHeading9 | -10 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["caption"]` | -35 | WdBuiltinStyle.wdStyleCaption | -35 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["title"]` | -63 | WdBuiltinStyle.wdStyleTitle | -63 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["body text"]` | -67 | WdBuiltinStyle.wdStyleBodyText | -67 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["subtitle"]` | -75 | WdBuiltinStyle.wdStyleSubtitle | -75 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["list paragraph"]` | -180 | WdBuiltinStyle.wdStyleListParagraph | -180 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["quote"]` | -181 | WdBuiltinStyle.wdStyleQuote | -181 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["intense quote"]` | -182 | WdBuiltinStyle.wdStyleIntenseQuote | -182 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["list bullet"]` | -49 | WdBuiltinStyle.wdStyleListBullet | -49 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["list number"]` | -50 | WdBuiltinStyle.wdStyleListNumber | -50 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `BUILTIN_STYLES["no spacing"]` | -158 | (no "No Spacing" member exists) | -158 is wdStyleNormalObject ("Normal (applied to an object).") | **NO** | https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle |
| `ALIGNMENTS["left"]` | 0 | WdParagraphAlignment.wdAlignParagraphLeft | 0 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdparagraphalignment |
| `ALIGNMENTS["center"]` | 1 | WdParagraphAlignment.wdAlignParagraphCenter | 1 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdparagraphalignment |
| `ALIGNMENTS["right"]` | 2 | WdParagraphAlignment.wdAlignParagraphRight | 2 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdparagraphalignment |
| `ALIGNMENTS["justify"]` | 3 | WdParagraphAlignment.wdAlignParagraphJustify | 3 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdparagraphalignment |
| `ALIGNMENTS["distribute"]` | 4 | WdParagraphAlignment.wdAlignParagraphDistribute | 4 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdparagraphalignment |
| `LINE_SPACING["single"]` | 0 | WdLineSpacing.wdLineSpaceSingle | 0 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdlinespacing |
| `LINE_SPACING["one_and_half"]` | 1 | WdLineSpacing.wdLineSpace1pt5 | 1 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdlinespacing |
| `LINE_SPACING["double"]` | 2 | WdLineSpacing.wdLineSpaceDouble | 2 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdlinespacing |
| `LINE_SPACING["at_least"]` | 3 | WdLineSpacing.wdLineSpaceAtLeast | 3 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdlinespacing |
| `LINE_SPACING["exactly"]` | 4 | WdLineSpacing.wdLineSpaceExactly | 4 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdlinespacing |
| `LINE_SPACING["multiple"]` | 5 | WdLineSpacing.wdLineSpaceMultiple | 5 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdlinespacing |
| `SAVE_FORMATS[".docx"]` | 16 | WdSaveFormat.wdFormatDocumentDefault | 16 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `SAVE_FORMATS[".docm"]` | 13 | WdSaveFormat.wdFormatXMLDocumentMacroEnabled | 13 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `SAVE_FORMATS[".dotx"]` | 14 | WdSaveFormat.wdFormatXMLTemplate | 14 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `SAVE_FORMATS[".doc"]` | 0 | WdSaveFormat.wdFormatDocument | 0 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `SAVE_FORMATS[".rtf"]` | 6 | WdSaveFormat.wdFormatRTF | 6 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `SAVE_FORMATS[".txt"]` | 2 | WdSaveFormat.wdFormatText | 2 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `SAVE_FORMATS[".pdf"]` | 17 | WdSaveFormat.wdFormatPDF | 17 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `SAVE_FORMATS[".odt"]` | 23 | WdSaveFormat.wdFormatOpenDocumentText | 23 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `SAVE_FORMATS[".htm"]` | 8 | WdSaveFormat.wdFormatHTML | 8 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `SAVE_FORMATS[".html"]` | 8 | WdSaveFormat.wdFormatHTML | 8 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat |
| `ExportAsFixedFormat(ExportFormat=17)` | 17 | WdExportFormat.wdExportFormatPDF | 17 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdexportformat |
| `Close(SaveChanges=-1)` | -1 | WdSaveOptions.wdSaveChanges | -1 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveoptions |
| `Close(SaveChanges=0)` | 0 | WdSaveOptions.wdDoNotSaveChanges | 0 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveoptions |
| `Font.Underline = 1` | 1 | WdUnderline.wdUnderlineSingle | 1 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdunderline |
| `Font.Underline = 0` | 0 | WdUnderline.wdUnderlineNone | 0 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdunderline |
| `_STYLE_TYPES[1]` | 1 | WdStyleType.wdStyleTypeParagraph | 1 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdstyletype |
| `_STYLE_TYPES[2]` | 2 | WdStyleType.wdStyleTypeCharacter | 2 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdstyletype |
| `_STYLE_TYPES[3]` | 3 | WdStyleType.wdStyleTypeTable | 3 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdstyletype |
| `_STYLE_TYPES[4]` | 4 | WdStyleType.wdStyleTypeList | 4 | yes | https://learn.microsoft.com/en-us/office/vba/api/word.wdstyletype |
| `_STYLE_TYPES[5] = "linked"` | 5 | not on the VBA page; interop: 5 = wdStyleTypeParagraphOnly, 6 = wdStyleTypeLinked | — | **NO** | https://learn.microsoft.com/en-us/office/vba/api/word.wdstyletype , https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdstyletype?view=word-pia |

`.htm`/`.html` → 8 is wdFormatHTML ("Standard HTML format"). The filtered variant, wdFormatFilteredHTML, is 10. `.pdf` → 17 is used only in `SAVE_FORMATS` for validation; the PDF itself is written with ExportAsFixedFormat.

### B.2 Enumerations

Each table was copied from the enumeration's page and is sorted by value.

#### WdBuiltinStyle

The page lists 132 members, from -1 to -265, with gaps in the values (-92…-94, -107…-157, -179, -183…-260). Values in the gaps are not documented. There is no member for "No Spacing".

**WdBuiltinStyle** — https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle

| Name | Value | Description |
|---|---|---|
| wdStyleBookTitle | -265 | Book Title. |
| wdStyleIntenseReference | -264 | Intense Reference. |
| wdStyleSubtleReference | -263 | Subtle Reference. |
| wdStyleIntenseEmphasis | -262 | Intense Emphasis. |
| wdStyleSubtleEmphasis | -261 | Subtle Emphasis. |
| wdStyleIntenseQuote | -182 | Intense Quote. |
| wdStyleQuote | -181 | Quote. |
| wdStyleListParagraph | -180 | List Paragraph. |
| wdStyleTableMediumList1Accent1 | -178 | Medium List 1 Accent 1. |
| wdStyleTableMediumShading2Accent1 | -177 | Medium Shading 2 Accent 1. |
| wdStyleTableMediumShading1Accent1 | -176 | Medium Shading 1 Accent 1. |
| wdStyleTableLightGridAccent1 | -175 | Light Grid Accent 1. |
| wdStyleTableLightListAccent1 | -174 | Light List Accent 1. |
| wdStyleTableLightShadingAccent1 | -173 | Light Shading Accent 1. |
| wdStyleTableColorfulGrid | -172 | Colorful Grid. |
| wdStyleTableColorfulList | -171 | Colorful List. |
| wdStyleTableColorfulShading | -170 | Colorful Shading. |
| wdStyleTableDarkList | -169 | Dark List. |
| wdStyleTableMediumGrid3 | -168 | Medium Grid 3. |
| wdStyleTableMediumGrid2 | -167 | Medium Grid 2. |
| wdStyleTableMediumGrid1 | -166 | Medium Grid 1. |
| wdStyleTableMediumList2 | -165 | Medium List 2. |
| wdStyleTableMediumList1 | -164 | Medium List 1. |
| wdStyleTableMediumShading2 | -163 | Medium Shading 2. |
| wdStyleTableMediumShading1 | -162 | Medium Shading 1. |
| wdStyleTableLightGrid | -161 | Light Grid. |
| wdStyleTableLightList | -160 | Light List. |
| wdStyleTableLightShading | -159 | Light Shading. |
| wdStyleNormalObject | -158 | Normal (applied to an object). |
| wdStyleNormalTable | -106 | Normal (applied within a table). |
| wdStyleHtmlVar | -105 | HTML Variable. |
| wdStyleHtmlTt | -104 | HTML Typewriter. |
| wdStyleHtmlSamp | -103 | HTML Sample. |
| wdStyleHtmlPre | -102 | HTML Preformatted. |
| wdStyleHtmlKbd | -101 | HTML Keyboard. |
| wdStyleHtmlDfn | -100 | HTML Definition. |
| wdStyleHtmlCode | -99 | HTML Code. |
| wdStyleHtmlCite | -98 | HTML Cite. |
| wdStyleHtmlAddress | -97 | HTML Address. |
| wdStyleHtmlAcronym | -96 | HTML Acronym. |
| wdStyleHtmlNormal | -95 | Normal (Web). |
| wdStylePlainText | -91 | Plain Text. |
| wdStyleNavPane | -90 | Document Map. |
| wdStyleEmphasis | -89 | Emphasis. |
| wdStyleStrong | -88 | Strong. |
| wdStyleHyperlinkFollowed | -87 | Followed Hyperlink. |
| wdStyleHyperlink | -86 | Hyperlink. |
| wdStyleBlockQuotation | -85 | Block Text. |
| wdStyleBodyTextIndent3 | -84 | Body Text Indent 3. |
| wdStyleBodyTextIndent2 | -83 | Body Text Indent 2. |
| wdStyleBodyText3 | -82 | Body Text 3. |
| wdStyleBodyText2 | -81 | Body Text 2. |
| wdStyleNoteHeading | -80 | Note Heading. |
| wdStyleBodyTextFirstIndent2 | -79 | Body Text First Indent 2. |
| wdStyleBodyTextFirstIndent | -78 | Body Text First Indent. |
| wdStyleDate | -77 | Date. |
| wdStyleSalutation | -76 | Salutation. |
| wdStyleSubtitle | -75 | Subtitle. |
| wdStyleMessageHeader | -74 | Message Header. |
| wdStyleListContinue5 | -73 | List Continue 5. |
| wdStyleListContinue4 | -72 | List Continue 4. |
| wdStyleListContinue3 | -71 | List Continue 3. |
| wdStyleListContinue2 | -70 | List Continue 2. |
| wdStyleListContinue | -69 | List Continue. |
| wdStyleBodyTextIndent | -68 | Body Text Indent. |
| wdStyleBodyText | -67 | Body Text. |
| wdStyleDefaultParagraphFont | -66 | Default Paragraph Font. |
| wdStyleSignature | -65 | Signature. |
| wdStyleClosing | -64 | Closing. |
| wdStyleTitle | -63 | Title. |
| wdStyleListNumber5 | -62 | List Number 5. |
| wdStyleListNumber4 | -61 | List Number 4. |
| wdStyleListNumber3 | -60 | List Number 3. |
| wdStyleListNumber2 | -59 | List Number 2. |
| wdStyleListBullet5 | -58 | List Bullet 5. |
| wdStyleListBullet4 | -57 | List Bullet 4. |
| wdStyleListBullet3 | -56 | List Bullet 3. |
| wdStyleListBullet2 | -55 | List Bullet 2. |
| wdStyleList5 | -54 | List 5. |
| wdStyleList4 | -53 | List 4. |
| wdStyleList3 | -52 | List 3. |
| wdStyleList2 | -51 | List 2. |
| wdStyleListNumber | -50 | List Number. |
| wdStyleListBullet | -49 | List Bullet. |
| wdStyleList | -48 | List. |
| wdStyleTOAHeading | -47 | TOA Heading. |
| wdStyleMacroText | -46 | Macro Text. |
| wdStyleTableOfAuthorities | -45 | Table of Authorities. |
| wdStyleEndnoteText | -44 | Endnote Text. |
| wdStyleEndnoteReference | -43 | Endnote Reference. |
| wdStylePageNumber | -42 | Page Number. |
| wdStyleLineNumber | -41 | Line Number. |
| wdStyleCommentReference | -40 | Comment Reference. |
| wdStyleFootnoteReference | -39 | Footnote Reference. |
| wdStyleEnvelopeReturn | -38 | Envelope Return. |
| wdStyleEnvelopeAddress | -37 | Envelope Address. |
| wdStyleTableOfFigures | -36 | Table of Figures. |
| wdStyleCaption | -35 | Caption. |
| wdStyleIndexHeading | -34 | Index Heading. |
| wdStyleFooter | -33 | Footer. |
| wdStyleHeader | -32 | Header. |
| wdStyleCommentText | -31 | Comment Text. |
| wdStyleFootnoteText | -30 | Footnote Text. |
| wdStyleNormalIndent | -29 | Normal Indent. |
| wdStyleTOC9 | -28 | TOC 9. |
| wdStyleTOC8 | -27 | TOC 8. |
| wdStyleTOC7 | -26 | TOC 7. |
| wdStyleTOC6 | -25 | TOC 6. |
| wdStyleTOC5 | -24 | TOC 5. |
| wdStyleTOC4 | -23 | TOC 4. |
| wdStyleTOC3 | -22 | TOC 3. |
| wdStyleTOC2 | -21 | TOC 2. |
| wdStyleTOC1 | -20 | TOC 1. |
| wdStyleIndex9 | -19 | Index 9. |
| wdStyleIndex8 | -18 | Index 8. |
| wdStyleIndex7 | -17 | Index 7. |
| wdStyleIndex6 | -16 | Index 6. |
| wdStyleIndex5 | -15 | Index 5. |
| wdStyleIndex4 | -14 | Index 4. |
| wdStyleIndex3 | -13 | Index 3. |
| wdStyleIndex2 | -12 | Index 2. |
| wdStyleIndex1 | -11 | Index 1. |
| wdStyleHeading9 | -10 | Heading 9. |
| wdStyleHeading8 | -9 | Heading 8. |
| wdStyleHeading7 | -8 | Heading 7. |
| wdStyleHeading6 | -7 | Heading 6. |
| wdStyleHeading5 | -6 | Heading 5. |
| wdStyleHeading4 | -5 | Heading 4. |
| wdStyleHeading3 | -4 | Heading 3. |
| wdStyleHeading2 | -3 | Heading 2. |
| wdStyleHeading1 | -2 | Heading 1. |
| wdStyleNormal | -1 | Normal. |

#### WdLineSpacing

**WdLineSpacing** — https://learn.microsoft.com/en-us/office/vba/api/word.wdlinespacing

| Name | Value | Description |
|---|---|---|
| wdLineSpaceSingle | 0 | Single spaced. default |
| wdLineSpace1pt5 | 1 | Space-and-a-half line spacing. Spacing is equivalent to the current font size plus 6 points. |
| wdLineSpaceDouble | 2 | Double spaced. |
| wdLineSpaceAtLeast | 3 | Line spacing is always at least a specified amount. The amount is specified separately. |
| wdLineSpaceExactly | 4 | Line spacing is only the exact maximum amount of space required. This setting commonly uses less space than single spacing. |
| wdLineSpaceMultiple | 5 | Line spacing determined by the number of lines indicated. |

#### WdParagraphAlignment

**WdParagraphAlignment** — https://learn.microsoft.com/en-us/office/vba/api/word.wdparagraphalignment

| Name | Value | Description |
|---|---|---|
| wdAlignParagraphLeft | 0 | Left-aligned. |
| wdAlignParagraphCenter | 1 | Center-aligned. |
| wdAlignParagraphRight | 2 | Right-aligned. |
| wdAlignParagraphJustify | 3 | Fully justified. |
| wdAlignParagraphDistribute | 4 | Paragraph characters are distributed to fill the entire width of the paragraph. |
| wdAlignParagraphJustifyMed | 5 | Justified with a medium character compression ratio. |
| wdAlignParagraphJustifyHi | 7 | Justified with a high character compression ratio. |
| wdAlignParagraphJustifyLow | 8 | Justified with a low character compression ratio. |
| wdAlignParagraphThaiJustify | 9 | Justified according to Thai formatting layout. |

#### WdSaveFormat

**WdSaveFormat** — https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat

| Name | Value | Description |
|---|---|---|
| wdFormatDocument | 0 | Microsoft Office Word 97 - 2003 binary file format. |
| wdFormatDocument97 | 0 | Microsoft Word 97 document format. |
| wdFormatTemplate | 1 | Word template format. |
| wdFormatTemplate97 | 1 | Word 97 template format. |
| wdFormatText | 2 | Microsoft Windows text format. |
| wdFormatTextLineBreaks | 3 | Windows text format with line breaks preserved. |
| wdFormatDOSText | 4 | Microsoft DOS text format. |
| wdFormatDOSTextLineBreaks | 5 | Microsoft DOS text with line breaks preserved. |
| wdFormatRTF | 6 | Rich text format (RTF). |
| wdFormatEncodedText | 7 | Encoded text format. |
| wdFormatUnicodeText | 7 | Unicode text format. |
| wdFormatHTML | 8 | Standard HTML format. |
| wdFormatWebArchive | 9 | Web archive format. |
| wdFormatFilteredHTML | 10 | Filtered HTML format. |
| wdFormatXML | 11 | Extensible Markup Language (XML) format. |
| wdFormatXMLDocument | 12 | XML document format. |
| wdFormatXMLDocumentMacroEnabled | 13 | XML document format with macros enabled. |
| wdFormatXMLTemplate | 14 | XML template format. |
| wdFormatXMLTemplateMacroEnabled | 15 | XML template format with macros enabled. |
| wdFormatDocumentDefault | 16 | Word default document file format. For Word, this is the DOCX format. |
| wdFormatPDF | 17 | PDF format. |
| wdFormatXPS | 18 | XPS format. |
| wdFormatFlatXML | 19 | Open XML file format saved as a single XML file. |
| wdFormatFlatXMLMacroEnabled | 20 | Open XML file format with macros enabled saved as a single XML file. |
| wdFormatFlatXMLTemplate | 21 | Open XML template format saved as a XML single file. |
| wdFormatFlatXMLTemplateMacroEnabled | 22 | Open XML template format with macros enabled saved as a single XML file. |
| wdFormatOpenDocumentText | 23 | OpenDocument Text format. |
| wdFormatStrictOpenXMLDocument | 24 | Strict Open XML document format. |

#### WdExportFormat

**WdExportFormat** — https://learn.microsoft.com/en-us/office/vba/api/word.wdexportformat

| Name | Value | Description |
|---|---|---|
| wdExportFormatPDF | 17 | Export document into PDF format. |
| wdExportFormatXPS | 18 | Export document into XML Paper Specification (XPS) format. |

#### WdStatistic

**WdStatistic** — https://learn.microsoft.com/en-us/office/vba/api/word.wdstatistic

| Name | Value | Description |
|---|---|---|
| wdStatisticWords | 0 | Count of words. |
| wdStatisticLines | 1 | Count of lines. |
| wdStatisticPages | 2 | Count of pages. |
| wdStatisticCharacters | 3 | Count of characters. |
| wdStatisticParagraphs | 4 | Count of paragraphs. |
| wdStatisticCharactersWithSpaces | 5 | Count of characters including spaces. |
| wdStatisticFarEastCharacters | 6 | Count of characters for Asian languages. |

#### WdUnderline

**WdUnderline** — https://learn.microsoft.com/en-us/office/vba/api/word.wdunderline

| Name | Value | Description |
|---|---|---|
| wdUnderlineNone | 0 | No underline. |
| wdUnderlineSingle | 1 | A single line. default. |
| wdUnderlineWords | 2 | Underline individual words only. |
| wdUnderlineDouble | 3 | A double line. |
| wdUnderlineDotted | 4 | Dots. |
| wdUnderlineThick | 6 | A single thick line. |
| wdUnderlineDash | 7 | Dashes. |
| wdUnderlineDotDash | 9 | Alternating dots and dashes. |
| wdUnderlineDotDotDash | 10 | An alternating dot-dot-dash pattern. |
| wdUnderlineWavy | 11 | A single wavy line. |
| wdUnderlineDottedHeavy | 20 | Heavy dots. |
| wdUnderlineDashHeavy | 23 | Heavy dashes. |
| wdUnderlineDotDashHeavy | 25 | Alternating heavy dots and heavy dashes. |
| wdUnderlineDotDotDashHeavy | 26 | An alternating heavy dot-dot-dash pattern. |
| wdUnderlineWavyHeavy | 27 | A heavy wavy line. |
| wdUnderlineDashLong | 39 | Long dashes. |
| wdUnderlineWavyDouble | 43 | A double wavy line. |
| wdUnderlineDashLongHeavy | 55 | Long heavy dashes. |

#### WdColorIndex

**WdColorIndex** — https://learn.microsoft.com/en-us/office/vba/api/word.wdcolorindex

| Name | Value | Description |
|---|---|---|
| wdByAuthor | -1 | Color defined by document author. |
| wdAuto | 0 | Automatic color. Default; usually black. |
| wdNoHighlight | 0 | Removes highlighting that has been applied. |
| wdBlack | 1 | Black color. |
| wdBlue | 2 | Blue color. |
| wdTurquoise | 3 | Turquoise color. |
| wdBrightGreen | 4 | Bright green color. |
| wdPink | 5 | Pink color. |
| wdRed | 6 | Red color. |
| wdYellow | 7 | Yellow color. |
| wdWhite | 8 | White color. |
| wdDarkBlue | 9 | Dark blue color. |
| wdTeal | 10 | Teal color. |
| wdGreen | 11 | Green color. |
| wdViolet | 12 | Violet color. |
| wdDarkRed | 13 | Dark red color. |
| wdDarkYellow | 14 | Dark yellow color. |
| wdGray50 | 15 | Shade 50 of gray color. |
| wdGray25 | 16 | Shade 25 of gray color. |

#### WdStyleType

The VBA page lists four members. The interop page (https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdstyletype?view=word-pia) adds wdStyleTypeParagraphOnly = 5 and wdStyleTypeLinked = 6, both "Reserved for internal use."

**WdStyleType** — https://learn.microsoft.com/en-us/office/vba/api/word.wdstyletype

| Name | Value | Description |
|---|---|---|
| wdStyleTypeParagraph | 1 | Paragraph style. |
| wdStyleTypeCharacter | 2 | Body character style. |
| wdStyleTypeTable | 3 | Table style. |
| wdStyleTypeList | 4 | List style. |

#### WdOutlineLevel

**WdOutlineLevel** — https://learn.microsoft.com/en-us/office/vba/api/word.wdoutlinelevel

| Name | Value | Description |
|---|---|---|
| wdOutlineLevel1 | 1 | Outline level 1. |
| wdOutlineLevel2 | 2 | Outline level 2. |
| wdOutlineLevel3 | 3 | Outline level 3. |
| wdOutlineLevel4 | 4 | Outline level 4. |
| wdOutlineLevel5 | 5 | Outline level 5. |
| wdOutlineLevel6 | 6 | Outline level 6. |
| wdOutlineLevel7 | 7 | Outline level 7. |
| wdOutlineLevel8 | 8 | Outline level 8. |
| wdOutlineLevel9 | 9 | Outline level 9. |
| wdOutlineLevelBodyText | 10 | No outline level. |

#### WdFindWrap

**WdFindWrap** — https://learn.microsoft.com/en-us/office/vba/api/word.wdfindwrap

| Name | Value | Description |
|---|---|---|
| wdFindStop | 0 | The find operation ends if the beginning or end of the search range is reached. |
| wdFindContinue | 1 | The find operation continues if the beginning or end of the search range is reached. |
| wdFindAsk | 2 | After searching the selection or range, Microsoft Word displays a message asking whether to search the remainder of the document. |

#### WdReplace

**WdReplace** — https://learn.microsoft.com/en-us/office/vba/api/word.wdreplace

| Name | Value | Description |
|---|---|---|
| wdReplaceNone | 0 | Replace no occurrences. |
| wdReplaceOne | 1 | Replace the first occurrence encountered. |
| wdReplaceAll | 2 | Replace all occurrences. |

#### WdCollapseDirection

**WdCollapseDirection** — https://learn.microsoft.com/en-us/office/vba/api/word.wdcollapsedirection

| Name | Value | Description |
|---|---|---|
| wdCollapseEnd | 0 | Collapse the range to the ending point. |
| wdCollapseStart | 1 | Collapse the range to the starting point. |

#### WdSaveOptions

**WdSaveOptions** — https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveoptions

| Name | Value | Description |
|---|---|---|
| wdPromptToSaveChanges | -2 | Prompt the user to save pending changes. |
| wdSaveChanges | -1 | Save pending changes automatically without prompting the user. |
| wdDoNotSaveChanges | 0 | Do not save pending changes. |

#### WdConstants

WdUndefined has no enumeration page of its own; `wdUndefined` = 9999999 is a member of WdConstants.

**WdConstants** — https://learn.microsoft.com/en-us/office/vba/api/word.wdconstants

| Name | Value | Description |
|---|---|---|
| wdBackward | -1073741823 | Indicates that selection will be extended backward using the MoveStartUntil or MoveStartWhile method of the Range or Selection object. |
| wdAutoPosition | 0 | Represents the Auto value for the specified setting. |
| wdFirst | 1 | Represents the first item in a collection. |
| wdToggle | 9999998 | Toggles a property's value. |
| wdUndefined | 9999999 | Represents an undefined value. |
| wdForward | 1073741823 | Indicates that selection will be extended forward using the MoveStartUntil or MoveStartWhile method of the Range or Selection object. |
| wdCreatorCode | 1297307460 | Represents the creator code for objects created by Microsoft Word. |

#### WdUnits

**WdUnits** — https://learn.microsoft.com/en-us/office/vba/api/word.wdunits

| Name | Value | Description |
|---|---|---|
| wdCharacter | 1 | A character. |
| wdWord | 2 | A word. |
| wdSentence | 3 | A sentence. |
| wdParagraph | 4 | A paragraph. |
| wdLine | 5 | A line. |
| wdStory | 6 | A story. |
| wdScreen | 7 | The screen dimensions. |
| wdSection | 8 | A section. |
| wdColumn | 9 | A column. |
| wdRow | 10 | A row. |
| wdWindow | 11 | A window. |
| wdCell | 12 | A cell. |
| wdCharacterFormatting | 13 | Character formatting. |
| wdParagraphFormatting | 14 | Paragraph formatting. |
| wdTable | 15 | A table. |
| wdItem | 16 | The selected item. |

#### WdListNumberStyle

ListLevel.NumberStyle takes these. For a bullet, wdListNumberStyleBullet = 23.

**WdListNumberStyle** — https://learn.microsoft.com/en-us/office/vba/api/word.wdlistnumberstyle

| Name | Value | Description |
|---|---|---|
| wdListNumberStyleArabic | 0 | Arabic numeric style. |
| wdListNumberStyleUppercaseRoman | 1 | Uppercase Roman style. |
| wdListNumberStyleLowercaseRoman | 2 | Lowercase Roman style. |
| wdListNumberStyleUppercaseLetter | 3 | Uppercase letter style. |
| wdListNumberStyleLowercaseLetter | 4 | Lowercase letter style. |
| wdListNumberStyleOrdinal | 5 | Ordinal style. |
| wdListNumberStyleCardinalText | 6 | Cardinal text style. |
| wdListNumberStyleOrdinalText | 7 | Ordinal text style. |
| wdListNumberStyleKanji | 10 | Kanji style. |
| wdListNumberStyleKanjiDigit | 11 | Kanji Digit style. |
| wdListNumberStyleAiueoHalfWidth | 12 | Aiueo half-width numeric style. |
| wdListNumberStyleIrohaHalfWidth | 13 | Iroha half width style. |
| wdListNumberStyleArabicFullWidth | 14 | Arabic full-width numeric style. |
| wdListNumberStyleKanjiTraditional | 16 | Kanji traditional style. |
| wdListNumberStyleKanjiTraditional2 | 17 | Kanji traditional 2 style. |
| wdListNumberStyleNumberInCircle | 18 | Number in circle style. |
| wdListNumberStyleAiueo | 20 | Aiueo numeric style. |
| wdListNumberStyleIroha | 21 | Iroha style. |
| wdListNumberStyleArabicLZ | 22 | Arabic LZ numeric style. |
| wdListNumberStyleBullet | 23 | Bullet style. |
| wdListNumberStyleGanada | 24 | Ganada style. |
| wdListNumberStyleChosung | 25 | Chosung style. |
| wdListNumberStyleGBNum1 | 26 | GB numeric 1 style. |
| wdListNumberStyleGBNum2 | 27 | GB numeric 2 style. |
| wdListNumberStyleGBNum3 | 28 | GB numeric 3 style. |
| wdListNumberStyleGBNum4 | 29 | GB numeric 4 style. |
| wdListNumberStyleZodiac1 | 30 | Zodiac 1 style. |
| wdListNumberStyleZodiac2 | 31 | Zodiac 2 style. |
| wdListNumberStyleZodiac3 | 32 | Zodiac 3 style. |
| wdListNumberStyleTradChinNum1 | 33 | Traditional Chinese numeric 1 style. |
| wdListNumberStyleTradChinNum2 | 34 | Traditional Chinese numeric 2 style. |
| wdListNumberStyleTradChinNum3 | 35 | Traditional Chinese numeric 3 style. |
| wdListNumberStyleTradChinNum4 | 36 | Traditional Chinese numeric 4 style. |
| wdListNumberStyleSimpChinNum1 | 37 | Simplified Chinese numeric 1 style. |
| wdListNumberStyleSimpChinNum2 | 38 | Simplified Chinese numeric 2 style. |
| wdListNumberStyleSimpChinNum3 | 39 | Simplified Chinese numeric 3 style. |
| wdListNumberStyleSimpChinNum4 | 40 | Simplified Chinese numeric 4 style. |
| wdListNumberStyleHanjaRead | 41 | Hanja Read style. |
| wdListNumberStyleHanjaReadDigit | 42 | Hanja Read Digit style. |
| wdListNumberStyleHangul | 43 | Hanqul style. |
| wdListNumberStyleHanja | 44 | Hanja style. |
| wdListNumberStyleHebrew1 | 45 | Hebrew 1 style. |
| wdListNumberStyleArabic1 | 46 | Arabic 1 numeric style. |
| wdListNumberStyleHebrew2 | 47 | Hebrew 2 style. |
| wdListNumberStyleArabic2 | 48 | Arabic 2 numeric style. |
| wdListNumberStyleHindiLetter1 | 49 | Hindi letter 1 style. |
| wdListNumberStyleHindiLetter2 | 50 | Hindi letter 2 style. |
| wdListNumberStyleHindiArabic | 51 | Hindi Arabic style. |
| wdListNumberStyleHindiCardinalText | 52 | Hindi Cardinal text style. |
| wdListNumberStyleThaiLetter | 53 | Thai letter style. |
| wdListNumberStyleThaiArabic | 54 | Thai Arabic style. |
| wdListNumberStyleThaiCardinalText | 55 | Thai Cardinal text style. |
| wdListNumberStyleVietCardinalText | 56 | Vietnamese Cardinal text style. |
| wdListNumberStyleLowercaseRussian | 58 | Lowercase Russian style. |
| wdListNumberStyleUppercaseRussian | 59 | Uppercase Russian style. |
| wdListNumberStyleLowercaseGreek | 60 | Lowercase Greek style. |
| wdListNumberStyleUppercaseGreek | 61 | Uppercase Greek style. |
| wdListNumberStyleArabicLZ2 | 62 | Arabic LZ2 numeric style. |
| wdListNumberStyleArabicLZ3 | 63 | Arabic LZ3 numeric style. |
| wdListNumberStyleArabicLZ4 | 64 | Arabic LZ4 numeric style. |
| wdListNumberStyleLowercaseTurkish | 65 | Lowercase Turkish style. |
| wdListNumberStyleUppercaseTurkish | 66 | Uppercase Turkish style. |
| wdListNumberStyleLowercaseBulgarian | 67 | Lowercase Bulgarian style. |
| wdListNumberStyleUppercaseBulgarian | 68 | Uppercase Bulgarian style. |
| wdListNumberStylePictureBullet | 249 | Picture bullet style. |
| wdListNumberStyleLegal | 253 | Legal style. |
| wdListNumberStyleLegalLZ | 254 | Legal LZ style. |
| wdListNumberStyleNone | 255 | No style applied. |

#### WdListGalleryType

**WdListGalleryType** — https://learn.microsoft.com/en-us/office/vba/api/word.wdlistgallerytype

| Name | Value | Description |
|---|---|---|
| wdBulletGallery | 1 | Bulleted list. |
| wdNumberGallery | 2 | Numbered list. |
| wdOutlineNumberGallery | 3 | Outline numbered list. |

#### WdDefaultListBehavior

**WdDefaultListBehavior** — https://learn.microsoft.com/en-us/office/vba/api/word.wddefaultlistbehavior

| Name | Value | Description |
|---|---|---|
| wdWord8ListBehavior | 0 | Use formatting compatible with Microsoft Word 97. |
| wdWord9ListBehavior | 1 | Use Web-oriented formatting as introduced in Microsoft Word 2000. |
| wdWord10ListBehavior | 2 | Use formatting compatible with Microsoft Word 2002. |

#### WdTrailingCharacter

**WdTrailingCharacter** — https://learn.microsoft.com/en-us/office/vba/api/word.wdtrailingcharacter

| Name | Value | Description |
|---|---|---|
| wdTrailingTab | 0 | A tab is inserted. |
| wdTrailingSpace | 1 | A space is inserted. Default. |
| wdTrailingNone | 2 | No character is inserted. |

#### WdListLevelAlignment

**WdListLevelAlignment** — https://learn.microsoft.com/en-us/office/vba/api/word.wdlistlevelalignment

| Name | Value | Description |
|---|---|---|
| wdListLevelAlignLeft | 0 | Left-aligned. |
| wdListLevelAlignCenter | 1 | Center-aligned. |
| wdListLevelAlignRight | 2 | Right-aligned. |

#### WdListType

**WdListType** — https://learn.microsoft.com/en-us/office/vba/api/word.wdlisttype

| Name | Value | Description |
|---|---|---|
| wdListNoNumbering | 0 | List with no bullets, numbering, or outlining. |
| wdListListNumOnly | 1 | ListNum fields that can be used in the body of a paragraph. |
| wdListBullet | 2 | Bulleted list. |
| wdListSimpleNumbering | 3 | Simple numeric list. |
| wdListOutlineNumbering | 4 | Outlined list. |
| wdListMixedNumbering | 5 | Mixed numeric list. |
| wdListPictureBullet | 6 | Picture bulleted list. |

#### WdListApplyTo

**WdListApplyTo** — https://learn.microsoft.com/en-us/office/vba/api/word.wdlistapplyto

| Name | Value | Description |
|---|---|---|
| wdListApplyToWholeList | 0 | Entire list. |
| wdListApplyToThisPointForward | 1 | From cursor insertion point to end of list. |
| wdListApplyToSelection | 2 | Selection. |

#### WdContinue

**WdContinue** — https://learn.microsoft.com/en-us/office/vba/api/word.wdcontinue

| Name | Value | Description |
|---|---|---|
| wdContinueDisabled | 0 | Formatting cannot continue from the previous list. |
| wdResetList | 1 | Numbering can be restarted. |
| wdContinueList | 2 | Formatting can continue from the previous list. |

#### WdEmphasisMark

**WdEmphasisMark** — https://learn.microsoft.com/en-us/office/vba/api/word.wdemphasismark

| Name | Value | Description |
|---|---|---|
| wdEmphasisMarkNone | 0 | No emphasis mark. |
| wdEmphasisMarkOverSolidCircle | 1 | A solid black circle. |
| wdEmphasisMarkOverComma | 2 | A comma. |
| wdEmphasisMarkOverWhiteCircle | 3 | An empty white circle. |
| wdEmphasisMarkUnderSolidCircle | 4 | A solid black circle. |

#### WdBaselineAlignment

**WdBaselineAlignment** — https://learn.microsoft.com/en-us/office/vba/api/word.wdbaselinealignment

| Name | Value | Description |
|---|---|---|
| wdBaselineAlignTop | 0 | Align along top of each font. |
| wdBaselineAlignCenter | 1 | Align center points of each font. |
| wdBaselineAlignBaseline | 2 | Align to a baseline for the paragraph. |
| wdBaselineAlignFarEast50 | 3 | Align using the baseline for Asian language font standards. |
| wdBaselineAlignAuto | 4 | Microsoft Word automatically adjusts the baseline font alignment. |

#### WdThemeColorIndex

**WdThemeColorIndex** — https://learn.microsoft.com/en-us/office/vba/api/word.wdthemecolorindex

| Name | Value | Description |
|---|---|---|
| wdNotThemeColor | -1 | No color. |
| wdThemeColorMainDark1 | 0 | Dark main color 1. |
| wdThemeColorMainLight1 | 1 | Light main color 1. |
| wdThemeColorMainDark2 | 2 | Dark main color 2. |
| wdThemeColorMainLight2 | 3 | Light main color 2. |
| wdThemeColorAccent1 | 4 | Accent color 1. |
| wdThemeColorAccent2 | 5 | Accent color 2. |
| wdThemeColorAccent3 | 6 | Accent color 3. |
| wdThemeColorAccent4 | 7 | Accent color 4. |
| wdThemeColorAccent5 | 8 | Accent color 5. |
| wdThemeColorAccent6 | 9 | Accent color 6. |
| wdThemeColorHyperlink | 10 | Hyperlink color. |
| wdThemeColorHyperlinkFollowed | 11 | Followed hyperlink color. |
| wdThemeColorBackground1 | 12 | Background color 1. |
| wdThemeColorText1 | 13 | Text color 1. |
| wdThemeColorBackground2 | 14 | Background color 2. |
| wdThemeColorText2 | 15 | Text color 2. |

#### WdReadingOrder

**WdReadingOrder** — https://learn.microsoft.com/en-us/office/vba/api/word.wdreadingorder

| Name | Value | Description |
|---|---|---|
| wdReadingOrderRtl | 0 | Right-to-left reading order. |
| wdReadingOrderLtr | 1 | Left-to-right reading order. |

#### WdLigatures

**WdLigatures** — https://learn.microsoft.com/en-us/office/vba/api/word.wdligatures

| Name | Value | Description |
|---|---|---|
| wdLigaturesNone | 0 | Does not apply any ligatures to the font. |
| wdLigaturesStandard | 1 | Applies standard ligatures to the font. Standard ligatures are designed to enhance readability. Standard ligatures in Latin languages include "fi", "fl", and "ff", for example. |
| wdLigaturesContextual | 2 | Applies contextual ligatures to the font. Contextual ligatures are often designed to enhance readability, but may also be solely ornamental. Contextual ligatures may also be contextual alternates. |
| wdLigaturesStandardContextual | 3 | Applies standard and contextual ligatures to the font. |
| wdLigaturesHistorical | 4 | Applies historical ligatures to the font. Historical ligatures are similar to standard ligatures in that they were originally intended to improve the readability of the font, but may look archaic to the modern reader. |
| wdLigaturesStandardHistorical | 5 | Applies standard and historical ligatures to the font. |
| wdLigaturesContextualHistorical | 6 | Applies contextual and historical ligatures to the font. |
| wdLigaturesStandardContextualHistorical | 7 | Applies standard, contextual, and historical ligatures to the font. |
| wdLigaturesDiscretional | 8 | Applies discretional ligatures to the font. Discretional ligatures are most often designed to be ornamental at the discretion of the type developer. |
| wdLigaturesStandardDiscretional | 9 | Applies standard and discretional ligatures to the font. |
| wdLigaturesContextualDiscretional | 10 | Applies contextual and discretional ligatures to the font. |
| wdLigaturesStandardContextualDiscretional | 11 | Applies standard, contextual and discretional ligatures to the font. |
| wdLigaturesHistoricalDiscretional | 12 | Applies historical and discretional ligatures to the font. |
| wdLigaturesStandardHistoricalDiscretional | 13 | Applies standard historical and discretional ligatures to the font. |
| wdLigaturesContextualHistoricalDiscretional | 14 | Applies contextual, historical, and discretional ligatures to a font. |
| wdLigaturesAll | 15 | Applies all types of ligatures to the font. |

#### WdNumberForm

**WdNumberForm** — https://learn.microsoft.com/en-us/office/vba/api/word.wdnumberform

| Name | Value | Description |
|---|---|---|
| wdNumberFormDefault | 0 | Applies the default number form for the font. |
| wdNumberFormLining | 1 | Applies the lining number form to the font. |
| wdNumberFormOldstyle | 2 | Applies the "old-style" number form to the font. |

#### WdNumberSpacing

**WdNumberSpacing** — https://learn.microsoft.com/en-us/office/vba/api/word.wdnumberspacing

| Name | Value | Description |
|---|---|---|
| wdNumberSpacingDefault | 0 | Applies the default number spacing for the font. |
| wdNumberSpacingProportional | 1 | Applies proportional number spacing to the font. |
| wdNumberSpacingTabular | 2 | Applies tabular number spacing to the font. |

#### WdStylisticSet

**WdStylisticSet** — https://learn.microsoft.com/en-us/office/vba/api/word.wdstylisticset

| Name | Value | Description |
|---|---|---|
| wdStylisticSetDefault | 0 | Default stylistic set for the specified font. |
| wdStylisticSet01 | 1 | First stylistic set for the specified font. |
| wdStylisticSet02 | 2 | Second stylistic set for the specified font. |
| wdStylisticSet03 | 4 | Third stylistic set for the specified font. |
| wdStylisticSet04 | 8 | Fourth stylistic set for the specified font. |
| wdStylisticSet05 | 16 | Fifth stylistic set for the specified font. |
| wdStylisticSet06 | 32 | Sixth stylistic set for the specified font. |
| wdStylisticSet07 | 64 | Seventh stylistic set for the specified font. |
| wdStylisticSet08 | 128 | Eighth stylistic set for the specified font. |
| wdStylisticSet09 | 256 | Ninth stylistic set for the specified font. |
| wdStylisticSet10 | 512 | Tenth stylistic set for the specified font. |
| wdStylisticSet11 | 1024 | Eleventh stylistic set for the specified font. |
| wdStylisticSet12 | 2048 | Twelfth stylistic set for the specified font. |
| wdStylisticSet13 | 4096 | Thirteenth stylistic set for the specified font. |
| wdStylisticSet14 | 8192 | Fourtheenth stylistic set for the specified font. |
| wdStylisticSet15 | 16384 | Fifthteenth stylistic set for the specified font. |
| wdStylisticSet16 | 32768 | Sixteenth stylistic set for the specified font. |
| wdStylisticSet17 | 65536 | Seventeenth stylistic set for the specified font. |
| wdStylisticSet18 | 131072 | Eighteenth stylistic set for the specified font. |
| wdStylisticSet19 | 262144 | Nineteenth stylistic set for the specified font. |
| wdStylisticSet20 | 524288 | Twentieth stylistic set for the specified font. |

#### WdTabAlignment

**WdTabAlignment** — https://learn.microsoft.com/en-us/office/vba/api/word.wdtabalignment

| Name | Value | Description |
|---|---|---|
| wdAlignTabLeft | 0 | Left-aligned. |
| wdAlignTabCenter | 1 | Center-aligned. |
| wdAlignTabRight | 2 | Right-aligned. |
| wdAlignTabDecimal | 3 | Decimal-aligned. |
| wdAlignTabBar | 4 | Bar-aligned. |
| wdAlignTabList | 6 | List-aligned. |

#### WdTabLeader

**WdTabLeader** — https://learn.microsoft.com/en-us/office/vba/api/word.wdtableader

| Name | Value | Description |
|---|---|---|
| wdTabLeaderSpaces | 0 | Spaces. Default. |
| wdTabLeaderDots | 1 | Dots. |
| wdTabLeaderDashes | 2 | Dashes. |
| wdTabLeaderLines | 3 | Double lines. |
| wdTabLeaderHeavy | 4 | A heavy line. |
| wdTabLeaderMiddleDot | 5 | A middle dot. |

#### WdExportOptimizeFor

**WdExportOptimizeFor** — https://learn.microsoft.com/en-us/office/vba/api/word.wdexportoptimizefor

| Name | Value | Description |
|---|---|---|
| wdExportOptimizeForPrint | 0 | Export for print, which is higher quality and results in a larger file size. |
| wdExportOptimizeForOnScreen | 1 | Export for screen, which is a lower quality and results in a smaller file size. |

#### WdExportRange

**WdExportRange** — https://learn.microsoft.com/en-us/office/vba/api/word.wdexportrange

| Name | Value | Description |
|---|---|---|
| wdExportAllDocument | 0 | Exports the entire document. |
| wdExportSelection | 1 | Exports the contents of the current selection. |
| wdExportCurrentPage | 2 | Exports the current page. |
| wdExportFromTo | 3 | Exports the contents of a range using the starting and ending positions. |

#### WdExportItem

**WdExportItem** — https://learn.microsoft.com/en-us/office/vba/api/word.wdexportitem

| Name | Value | Description |
|---|---|---|
| wdExportDocumentContent | 0 | Exports the document without markup. |
| wdExportDocumentWithMarkup | 7 | Exports the document with markup. |

#### WdExportCreateBookmarks

**WdExportCreateBookmarks** — https://learn.microsoft.com/en-us/office/vba/api/word.wdexportcreatebookmarks

| Name | Value | Description |
|---|---|---|
| wdExportCreateNoBookmarks | 0 | Do not create bookmarks in the exported document. |
| wdExportCreateHeadingBookmarks | 1 | Create a bookmark in the exported document for each Microsoft Word heading, which includes only headings within the main document and text boxes not within headers, footers, endnotes, footnotes, or comments. |
| wdExportCreateWordBookmarks | 2 | Create a bookmark in the exported document for each Word bookmark, which includes all bookmarks except those contained within headers and footers. |

#### WdOriginalFormat

**WdOriginalFormat** — https://learn.microsoft.com/en-us/office/vba/api/word.wdoriginalformat

| Name | Value | Description |
|---|---|---|
| wdWordDocument | 0 | Microsoft Word document format. |
| wdOriginalDocumentFormat | 1 | Original document format. |
| wdPromptUser | 2 | Prompt user to select a document format. |

#### WdColor

Source: https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor. Each value is the decimal form of a BGR triple: value = R + G×256 + B×65536. For example, wdColorRed = 255 (#FF0000) and wdColorBlue = 16711680 (#0000FF). `wdColorAutomatic` = -16777216 ("Automatic color; default; usually black"). The page lists 60 named colours; a few are below.

| Name | Value | HEX on the page |
|---|---|---|
| wdColorAqua | 13421619 | #33CCCC |
| wdColorAutomatic | -16777216 | #000000 |
| wdColorBlack | 0 | #000000 |
| wdColorBlue | 16711680 | #0000FF |
| wdColorBrightGreen | 65280 | #00FF00 |
| wdColorRed | 255 | #FF0000 |
| wdColorWhite | 16777215 | #FFFFFF |

## C. Styles, in full

### C.1 Style object: every documented member

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.style.

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| AutomaticallyUpdate | Boolean | R/W | True if the style is automatically redefined based on the selection. — Remarks: If the AutomaticallyUpdate property is set to False, Microsoft Word prompts for confirmation before redefining the style based on the selection. A style can be redefined when it is applied to a selection that has the same style but different manual formatting. The AutomaticallyUpdate property applies to paragraph styles only. | https://learn.microsoft.com/en-us/office/vba/api/word.style.automaticallyupdate |
| BaseStyle | Variant | R/W | Returns or sets an existing style on which you can base the formatting of another style. — Remarks: To set the BaseStyle property, specify either the local name of the base style, an integer or a wdBuiltinStyle constant, or an object that represents the base style. For a list of the wdBuiltinStyle constants, see the Style property for the object that you want to set. | https://learn.microsoft.com/en-us/office/vba/api/word.style.basestyle |
| Borders | Borders | not stated | Returns a Borders collection that represents all the borders for the specified style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.borders |
| BuiltIn | Boolean | R | True if the specified object is one of the built-in styles or caption labels in Word. — Remarks: You can specify built-in styles across all languages by using the WdBuiltinStyle constants or within a language by using the style name for the language version of Word. For example, if you specify U.S. English in your Microsoft Office language settings, the following statements are equivalent: | https://learn.microsoft.com/en-us/office/vba/api/word.style.builtin |
| Delete() | no parameters | method | Deletes the specified style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.delete |
| Description | String | R | Returns the description of the specified style. — Remarks: A typical example of a description for a style might be "Normal + Font: Arial, 12 pt, Bold, Italic, Space Before 12 pt After 3 pt, KeepWithNext, Level 2." | https://learn.microsoft.com/en-us/office/vba/api/word.style.description |
| Font | Font | R/W | Returns or sets a Font object that represents the character formatting of the specified style. — Remarks: To set this property, specify an expression that returns a Font object. | https://learn.microsoft.com/en-us/office/vba/api/word.style.font |
| Frame | Frame | R | Returns a Frame object that represents the frame formatting for the specified style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.frame |
| InUse | Boolean | R | True if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document. — Remarks: The InUse property doesn't necessarily indicate whether the style is currently applied to any text in the document. For instance, if text that's been formatted with a style is deleted, the InUse property of the style remains True. For built-in styles that have never been used in the document, this property returns False. | https://learn.microsoft.com/en-us/office/vba/api/word.style.inuse |
| LanguageID | WdLanguageID | R/W | Returns or sets a WdLanguageID constant that represents the language for the specified range. — Remarks: Some of the WdLanguageID constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed. | https://learn.microsoft.com/en-us/office/vba/api/word.style.languageid |
| LanguageIDFarEast | WdLanguageID | R/W | Returns or sets an East Asian language for the specified object. — Remarks: This is the recommended way to return or set the language of East Asian text in a document created in an East Asian version of Microsoft Word. | https://learn.microsoft.com/en-us/office/vba/api/word.style.languageidfareast |
| Linked | Boolean | R | Returns a Boolean that represents whether a style is a linked style that can be used for both paragraph and character formatting. — Remarks: True indicates that the Style Type option in the Modify Style dialog box is set to " Linked (paragraph and character)". | https://learn.microsoft.com/en-us/office/vba/api/word.style.linked |
| LinkStyle | Variant | R/W | Sets or returns a Variant that represents a link between a paragraph and a character style. — Remarks: When a character style and a paragraph style are linked, the two styles take on the same character formatting. | https://learn.microsoft.com/en-us/office/vba/api/word.style.linkstyle |
| LinkToListTemplate() | (ListTemplate: ListTemplate object, ListLevelNumber?: Variant) | method | Links the specified style to a list template so that the style's formatting can be applied to lists. — Params: ListTemplate: The list template that the style is to be linked to.; ListLevelNumber: An integer corresponding to the list level that the style is to be linked to. If this argument is omitted, then the level of the style is u… | https://learn.microsoft.com/en-us/office/vba/api/word.style.linktolisttemplate |
| ListLevelNumber | Long | R | Returns the list level for the specified style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.listlevelnumber |
| ListTemplate | ListTemplate | not stated | Returns a ListTemplate object that represents the list formatting for the specified Style object. — Remarks: A list template includes all the formatting that defines a particular list. Each of the seven formats (excluding None) found on each of the tabs in the Bullets and Numbering dialog box (Format menu) corresponds to a list template. Documents and templates can also contain collections of list templates. | https://learn.microsoft.com/en-us/office/vba/api/word.style.listtemplate |
| Locked | Boolean | R/W | True if a style cannot be changed or edited. | https://learn.microsoft.com/en-us/office/vba/api/word.style.locked |
| NameLocal | String | R/W | Returns the name of a built-in style in the language of the user. — Remarks: Setting this property renames a user-defined style or adds an alias to a built-in style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.namelocal |
| NextParagraphStyle | Variant | R/W | Returns or sets the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style. — Remarks: You can set the NextParagraphStyle property by using the local name of the style, an integer or a WdBuiltinStyle constant, or an object that represents the next style. For a list of the WdBuiltinStyle constants, see the Style property for the object that you want to set. | https://learn.microsoft.com/en-us/office/vba/api/word.style.nextparagraphstyle |
| NoProofing | Long | R/W | True if the spelling and grammar checker ignores text formatted with this style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.noproofing |
| NoSpaceBetweenParagraphsOfSameStyle | Boolean | R/W | True for Microsoft Word to remove spacing between paragraphs that are formatted using the same style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.nospacebetweenparagraphsofsamestyle |
| ParagraphFormat | ParagraphFormat | R/W | Returns or sets a ParagraphFormat object that represents the paragraph settings for the specified style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.paragraphformat |
| Priority | Long | R/W | Returns or sets a Long that represents the priority for sorting styles in the Styles task pane. | https://learn.microsoft.com/en-us/office/vba/api/word.style.priority |
| QuickStyle | Boolean | R/W | Returns or sets a Boolean that represents whether the style corresponds to an available quick style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.quickstyle |
| Shading | Shading | not stated | Returns a Shading object that refers to the shading formatting for the specified object. | https://learn.microsoft.com/en-us/office/vba/api/word.style.shading |
| Table | TableStyle | not stated | Returns a TableStyle object representing properties that can be applied to a table using a table style. | https://learn.microsoft.com/en-us/office/vba/api/word.style.table |
| Type | WdStyleType | R | Returns the style type. | https://learn.microsoft.com/en-us/office/vba/api/word.style.type |
| UnhideWhenUsed | not stated | R/W | True if the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Word after it is used in the document. | https://learn.microsoft.com/en-us/office/vba/api/word.style.unhidewhenused |
| Visibility | not stated | R/W | True if the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane. | https://learn.microsoft.com/en-us/office/vba/api/word.style.visibility |

Notes:
- `Description` is **read-only** ("Returns the description of the specified style. Read-only String."), so a tool cannot set it (https://learn.microsoft.com/en-us/office/vba/api/word.style.description).
- `Linked` is read-only. `LinkStyle` (R/W Variant) is how a paragraph style and a character style are linked; "When a character style and a paragraph style are linked, the two styles take on the same character formatting" (https://learn.microsoft.com/en-us/office/vba/api/word.style.linkstyle).
- `ListLevelNumber` is read-only, and `ListTemplate` returns the style's list formatting. To attach a list template to a style, call `LinkToListTemplate(ListTemplate, ListLevelNumber)` (https://learn.microsoft.com/en-us/office/vba/api/word.style.linktolisttemplate).
- `BaseStyle` and `NextParagraphStyle` take "the local name of the … style, an integer or a WdBuiltinStyle constant, or an object" (https://learn.microsoft.com/en-us/office/vba/api/word.style.basestyle, https://learn.microsoft.com/en-us/office/vba/api/word.style.nextparagraphstyle).
- `AutomaticallyUpdate` "applies to paragraph styles only" (https://learn.microsoft.com/en-us/office/vba/api/word.style.automaticallyupdate).
- `Font` and `ParagraphFormat` are R/W: you can assign a whole Font/ParagraphFormat object, e.g. a `.Duplicate` (https://learn.microsoft.com/en-us/office/vba/api/word.style.font, https://learn.microsoft.com/en-us/office/vba/api/word.style.paragraphformat, https://learn.microsoft.com/en-us/office/vba/api/word.font.duplicate). `Borders` and `Shading` have no access mode stated. `Frame` is "Read-only".

### C.2 Styles collection, Document.Styles, UpdateStyles, Range.Style vs Paragraph.Style

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Add() | (Name: String, Type?: WdStyleType) → Style | method | Creates a new user-defined style and adds it to the Styles collection. — Params: Name: The new style name.; Type: Can be one of the WdStyleType constants. The default is paragraph style. | https://learn.microsoft.com/en-us/office/vba/api/word.styles.add |
| Count | Long | R | Returns a Long that represents the number of styles in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.styles.count |
| Item() | (Index: Variant) → Style | method | Returns an individual Style object in a collection. — Params: Index: The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual obj… | https://learn.microsoft.com/en-us/office/vba/api/word.styles.item |
| Styles | Styles | R | Returns a Styles collection for the specified document. | https://learn.microsoft.com/en-us/office/vba/api/word.document.styles |
| UpdateStyles() | no parameters | method | Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name. | https://learn.microsoft.com/en-us/office/vba/api/word.document.updatestyles |
| CopyStylesFromTemplate() | (Template: String) | method | Copies styles from the specified template to a document. — Remarks: When styles are copied from a template to a document, like-named styles in the document are redefined to match the style descriptions in the template. Unique styles from the template are copied to the document. Unique styles in the document remain intact. — Params: Template: The template file name. | https://learn.microsoft.com/en-us/office/vba/api/word.document.copystylesfromtemplate |
| Style | Variant | R/W | Returns or sets the style for the specified object. — Remarks: To set this property, specify the local name of the style, an integer, a WdBuiltinStyle constant, or an object that represents the style. When you return the style for a range that includes more than one style, only the first character or paragraph style is returned. | https://learn.microsoft.com/en-us/office/vba/api/word.range.style |
| Style | Variant | R/W | Returns or sets the style for the specified object. — Remarks: To set this property, specify the local name of the style, an integer, a WdBuiltinStyle constant, or an object that represents the style. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.style |
| Style | Variant | R/W | Returns or sets the style for the specified object. — Remarks: To set this property, specify the local name of the style, an integer, a WdBuiltinStyle constant, or an object that represents the style. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.style |
| CharacterStyle | Variant | R | Returns a Variant that represents the style used to format one or more characters. | https://learn.microsoft.com/en-us/office/vba/api/word.range.characterstyle |
| ParagraphStyle | Variant | R | Returns a Variant that represents the style used to format a paragraph. | https://learn.microsoft.com/en-us/office/vba/api/word.range.paragraphstyle |

- **Styles.Add(Name, Type?)**: Type is a WdStyleType and "The default is paragraph style." It returns the new Style (https://learn.microsoft.com/en-us/office/vba/api/word.styles.add). The page does not mention a BaseStyle argument; set `Style.BaseStyle` afterwards.
- **Range.Style vs Paragraph.Style**: both take a local name, an integer, a WdBuiltinStyle constant or a Style object. Only the Range.Style page says that for a range with several styles "only the first character or paragraph style is returned". To read a range's styles separately, use `Range.CharacterStyle` and `Range.ParagraphStyle`, both read-only Variants (https://learn.microsoft.com/en-us/office/vba/api/word.range.characterstyle, https://learn.microsoft.com/en-us/office/vba/api/word.range.paragraphstyle).
- **Document.UpdateStyles** "Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name." It takes no parameters (https://learn.microsoft.com/en-us/office/vba/api/word.document.updatestyles).

### C.3 ParagraphFormat: every documented member

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat. Units are as the pages state them: points unless it says "in characters" (CharacterUnit*) or "in gridlines" (LineUnit*).

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| AddSpaceBetweenFarEastAndAlpha | Long | R/W | True if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. This property returns wdUndefined if it's set to True for only some of the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.addspacebetweenfareastandalpha |
| AddSpaceBetweenFarEastAndDigit | Long | R/W | True if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. This property returns wdUndefined if it's set to True for only some of the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.addspacebetweenfareastanddigit |
| Alignment | WdParagraphAlignment | R/W | Returns or sets a WdParagraphAlignment constant that represents the alignment for the specified paragraphs. — Remarks: Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.alignment |
| AutoAdjustRightIndent | Long | R/W | True if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you've specified a set number of characters per line. Returns wdUndefined if the AutoAdjustRightIndent property is set to True for only some of the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.autoadjustrightindent |
| BaseLineAlignment | WdBaselineAlignment | R/W | Returns or sets a WdBaselineAlignment constant that represents the vertical position of fonts on a line. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.baselinealignment |
| Borders | Borders | not stated | Returns a Borders collection that represents all the borders for the specified object. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.borders |
| CharacterUnitFirstLineIndent | Single | R/W | Returns or sets the value (in characters) for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.characterunitfirstlineindent |
| CharacterUnitLeftIndent | Single | R/W | Returns or sets the left indent value (in characters) for the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.characterunitleftindent |
| CharacterUnitRightIndent | Single | R/W | Returns or sets the right indent value (in characters) for the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.characterunitrightindent |
| CloseUp() | no parameters | method | Removes any spacing before paragraphs in the specified paragraph format. — Remarks: The following two statements are equivalent: ParagraphFormat.CloseUp ParagraphFormat.SpaceBefore = 0 | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.closeup |
| CollapsedByDefault | Long | R/W | Returns or sets whether the specified paragraph format is collapsed by default. — Remarks: The CollapsedByDefault property setting corresponds to the setting of the Collapsed by Default box under general on the Indents and Spacing tab of the Paragraph dialog box. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.collapsedbydefault |
| DisableLineHeightGrid | Long | R/W | True if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. Returns wdUndefined if the DisableLineHeightGrid property is set to True for only some of the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.disablelineheightgrid |
| Duplicate | ParagraphFormat | not stated | Returns a read-only ParagraphFormat object that represents the paragraph formatting of the specified paragraph. — Remarks: Use the Duplicate property to pick up the settings of all the properties of a duplicated ParagraphFormat object. You can assign the object returned by the Duplicate property to another object of the same type to apply those settings all at once. Before assigning the duplicate object to another object, you can change any of the properties of the duplicate object without affecting the original. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.duplicate |
| FarEastLineBreakControl | Long | R/W | True if Microsoft Word applies East Asian line-breaking rules to the specified paragraphs. Returns wdUndefined if the FarEastLineBreakControl property is set to True for only some of the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.fareastlinebreakcontrol |
| FirstLineIndent | Single | R/W | Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.firstlineindent |
| HalfWidthPunctuationOnTopOfLine | Long | R/W | True if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. This property returns wdUndefined if it's set to True for only some of the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.halfwidthpunctuationontopofline |
| HangingPunctuation | Long | R/W | True if hanging punctuation is enabled for the specified paragraphs. This property returns wdUndefined if it's set to True for only some of the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.hangingpunctuation |
| Hyphenation | Long | R/W | True if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation. — Remarks: This property can be True, False or wdUndefined. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.hyphenation |
| IndentCharWidth() | (Count: Integer) | method | Indents one or more paragraphs by a specified number of characters. — Remarks: Using this method is equivalent to clicking the Increase Indent button on the Formatting toolbar. — Params: Count: The number of characters by which the specified paragraphs are to be indented. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.indentcharwidth |
| IndentFirstLineCharWidth() | (Count: Integer) | method | Indents the first line of one or more paragraphs by a specified number of characters. — Params: Count: The number of characters by which the first line of each specified paragraph is to be indented. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.indentfirstlinecharwidth |
| KeepTogether | Long | R/W | True if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document. — Remarks: This property can be True, False, or wdUndefined. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.keeptogether |
| KeepWithNext | Long | R/W | True if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. — Remarks: This property can be True, False, or wdUndefined. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.keepwithnext |
| LeftIndent | Single | R/W | Returns or sets a Single that represents the left indent value (in points) for the specified paragraph formatting. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.leftindent |
| LineSpacing | Single | R/W | Returns or sets the line spacing (in points) for the specified paragraphs. — Remarks: Use the LinesToPoints method to convert a number of lines to the corresponding value in points. For example, LinesToPoints(2) returns the value 24. The LineSpacing property can be set after the LineSpacingRule property has been set to: wdLineSpaceAtLeast the line spacing can be greater than or equal to, but never less than, the specified LineSpacing value. wdLineSpaceExactly the line spacing never changes from the s… | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.linespacing |
| LineSpacingRule | WdLineSpacing | R/W | Returns or sets the line spacing for the specified paragraph formatting. — Remarks: Use wdLineSpaceSingle, wdLineSpace1pt5, or wdLineSpaceDouble to set the line spacing to one of these values. To set the line spacing to an exact number of points or to a multiple number of lines, you must also set the LineSpacing property. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.linespacingrule |
| LineUnitAfter | Single | R/W | Returns or sets the amount of spacing (in gridlines) after the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.lineunitafter |
| LineUnitBefore | Single | R/W | Returns or sets the amount of spacing (in gridlines) before the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.lineunitbefore |
| MirrorIndents | Long | R/W | Returns or sets a Long that represents whether left and right indents are the same width. Can be True, False, or wdUndefined. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.mirrorindents |
| NoLineNumber | Long | R/W | True if line numbers are repressed for the specified paragraphs. Can be True, False, or wdUndefined. — Remarks: Use the LineNumbering property of the PageSetup object to turn on line numbering. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.nolinenumber |
| OpenOrCloseUp() | no parameters | method | Toggles the spacing before the specified paragraphs. — Remarks: If the spacing before the specified paragraphs is 0 (zero), this method sets spacing to 12 points. If the spacing before the paragraphs is greater than 0 (zero), this method sets spacing to 0 (zero). | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.openorcloseup |
| OpenUp() | no parameters | method | Sets spacing before the specified paragraphs to 12 points. — Remarks: You can also use the SpaceBefore property to set the spacing of paragraphs. The following two statements are equivalent: | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.openup |
| OutlineLevel | WdOutlineLevel | R/W | Returns or sets the outline level for the specified paragraphs. — Remarks: If a paragraph has a heading style applied to it (Heading 1 through Heading 9), the outline level is the same as the heading style and cannot be changed. Outline levels are visible only in outline view or the document map pane. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.outlinelevel |
| PageBreakBefore | Long | R/W | True if a page break is forced before the specified paragraphs. Can be True, False, or wdUndefined. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.pagebreakbefore |
| ReadingOrder | WdReadingOrder | R/W | Returns or sets the reading order of the specified paragraphs without changing their alignment. — Remarks: Use the LtrPara, LtrRun, RtlPara, and RtlRun methods of the Selection object to change the paragraph alignment along with the reading order. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.readingorder |
| Reset() | no parameters | method | Removes manual paragraph formatting (formatting not applied using a style). — Remarks: If you manually right align a paragraph and the underlying style has a different alignment, the Reset method changes the alignment to match the formatting of the underlying style. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.reset |
| RightIndent | Single | R/W | Returns or sets the right indent (in points) for the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.rightindent |
| Shading | Shading | not stated | Returns a Shading object that refers to the shading formatting for the specified object. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.shading |
| Space1() | no parameters | method | Single-spaces the specified paragraphs. — Remarks: The exact spacing is determined by the font size of the largest characters in each paragraph. You can also use the LineSpacingRule property to set the spacing of paragraphs. The following two statements are equivalent: | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.space1 |
| Space15() | no parameters | method | Formats the specified paragraphs with 1.5-line spacing. — Remarks: The exact spacing is determined by adding 6 points to the font size of the largest character in each paragraph. You can also use the LineSpacingRule property to set the spacing of paragraphs. The following two statements are equivalent: | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.space15 |
| Space2() | no parameters | method | Double-spaces the specified paragraphs. — Remarks: The exact spacing is determined by adding 12 points to the font size of the largest character in each paragraph. You can also use the LineSpacingRule property to set the spacing of paragraphs. For example, the following two statements are equivalent: | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.space2 |
| SpaceAfter | Single | R/W | Returns or sets the amount of spacing (in points) after the specified paragraph or text column. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spaceafter |
| SpaceAfterAuto | Long | R/W | True if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. — Remarks: Returns wdUndefined if the SpaceAfterAuto property is set to True for only some of the specified paragraphs. Can be set to either True or False. If SpaceAfterAuto is set to True, the SpaceAfter property is ignored. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spaceafterauto |
| SpaceBefore | Single | R/W | Returns or sets the spacing (in points) before the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spacebefore |
| SpaceBeforeAuto | Long | R/W | True if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. — Remarks: Returns wdUndefined if the SpaceBeforeAuto property is set to True for only some of the specified paragraphs. Can be set to either True or False. If SpaceBeforeAuto is set to True, the SpaceBefore property is ignored. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spacebeforeauto |
| Style | Variant | R/W | Returns or sets the style for the specified object. — Remarks: To set this property, specify the local name of the style, an integer, a WdBuiltinStyle constant, or an object that represents the style. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.style |
| TabHangingIndent() | (Count: Integer) | method | Sets a hanging indent to a specified number of tab stops. . — Remarks: You can also use this method to remove tab stops from a hanging indent if the value of Count is a negative number. — Params: Count: The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative). | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.tabhangingindent |
| TabIndent() | (Count: Integer) | method | Sets the left indent for the specified paragraphs to a specified number of tab stops. — Remarks: You can also use this method to remove an indent if the value of Count is a negative number. — Params: Count: The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative). | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.tabindent |
| TabStops | TabStops | R/W | Returns or sets a TabStops collection that represents all the custom tab stops for the specified paragraphs. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.tabstops |
| TextboxTightWrap | WdTextboxTightWrap | R/W | Returns or sets a WdTextboxTightWrap constant that represents how tightly text wraps around shapes or text boxes. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.textboxtightwrap |
| WidowControl | Long | R/W | True if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Can be True, False or wdUndefined. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.widowcontrol |
| WordWrap | Long | R/W | True if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames. — Remarks: This property returns wdUndefined if it is set to True for only some of the specified paragraphs or text frames. This usage may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.wordwrap |

East Asian notes, from the pages above:
- `LineUnitBefore` / `LineUnitAfter`: spacing "in gridlines", Single, R/W. `CharacterUnitFirstLineIndent` / `CharacterUnitLeftIndent` / `CharacterUnitRightIndent`: "in characters", Single, R/W. How the character units interact with the point-based `FirstLineIndent` / `LeftIndent` / `RightIndent` is **not stated**.
- `AutoAdjustRightIndent`, `DisableLineHeightGrid`, `WordWrap`, `HangingPunctuation`, `FarEastLineBreakControl`, `HalfWidthPunctuationOnTopOfLine`, `AddSpaceBetweenFarEastAndAlpha`, `AddSpaceBetweenFarEastAndDigit`: all Long, R/W, True/False or wdUndefined when mixed. See Discrepancy 3 for the documented sense of `DisableLineHeightGrid`.
- `SpaceBeforeAuto` / `SpaceAfterAuto`: "If SpaceBeforeAuto is set to True, the SpaceBefore property is ignored" (and likewise for After).
- `OutlineLevel` cannot be changed on paragraphs with Heading 1–9 applied.
- The same properties exist on the Paragraphs collection and the Paragraph object (e.g. https://learn.microsoft.com/en-us/office/vba/api/word.paragraphs.addspacebetweenfareastandalpha).

### C.4 Font: every documented member

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.font. The Word interop `_Font` member list was checked as well (https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._font?view=word-pia).

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| AllCaps | Long | R/W | True if the font is formatted as all capital letters. — Remarks: Returns True, False, or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle (reverses the current setting). Setting AllCaps to True sets SmallCaps to False, and vice versa. | https://learn.microsoft.com/en-us/office/vba/api/word.font.allcaps |
| Bold | Long | R/W | True if the font is formatted as bold. — Remarks: Returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. | https://learn.microsoft.com/en-us/office/vba/api/word.font.bold |
| BoldBi | Long | R/W | True if the font is formatted as bold. — Remarks: This property returns True, False or wdUndefined (for a mixture of bold and non-bold text). Can be set to True, False, or wdToggle. The BoldBi property applies to text in a right-to-left language. | https://learn.microsoft.com/en-us/office/vba/api/word.font.boldbi |
| Borders | Borders | not stated | Returns a Borders collection that represents all the borders for the specified font. | https://learn.microsoft.com/en-us/office/vba/api/word.font.borders |
| ColorIndex | WdColorIndex | R/W | Returns or sets a WdColorIndex constant that represents the color for the specified font. — Remarks: The wdByAuthor constant is not a valid color for fonts. | https://learn.microsoft.com/en-us/office/vba/api/word.font.colorindex |
| ColorIndexBi | WdColorIndex | R/W | Returns or sets the color for the specified Font object in a right-to-left language document. — Remarks: The wdByAuthor constant is not valid for Font objects. | https://learn.microsoft.com/en-us/office/vba/api/word.font.colorindexbi |
| ContextualAlternates | Long | R/W | Specifies whether or not contextual alternates are enabled for the specified font. — Remarks: Contextual alternates are ligatures that are applied to individual characters based on the letters around them (their context). Contextual alternates can also be applied to entire words in certain contexts, for example, words frequently used in titles (such as "of" and "the"). When contextual alternates are enabled for a font, they are used instead of the standard ligatures in those contexts defined by the font desi… | https://learn.microsoft.com/en-us/office/vba/api/word.font.contextualalternates |
| DiacriticColor | WdColor | R/W | Returns or sets the 24-bit color to be used for diacritics for the specified Font object. — Remarks: This property can be any valid WdColor constant or a value returned by Visual Basic's RGB function. The value of the UseDiffDiacColor property must be True to use this property. | https://learn.microsoft.com/en-us/office/vba/api/word.font.diacriticcolor |
| DisableCharacterSpaceGrid | Boolean | R/W | True if Microsoft Word ignores the number of characters per line for the corresponding Font object. — Remarks: This property returns wdUndefined if the DisableCharacterSpaceGrid property is set to True for only some of the specified text. | https://learn.microsoft.com/en-us/office/vba/api/word.font.disablecharacterspacegrid |
| DoubleStrikeThrough | Long | R/W | True if the specified font is formatted as double strikethrough text. . — Remarks: Returns True, False, or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. Read/write Long.To set or return single-line strikethrough formatting, use the StrikeThrough property. Setting DoubleStrikeThrough to True sets StrikeThrough to False, and vice versa. | https://learn.microsoft.com/en-us/office/vba/api/word.font.doublestrikethrough |
| Duplicate | Font | not stated | Returns a copy of a Font object that represents the character formatting of the specified font. — Remarks: Use the Duplicate property to pick up the settings of all the properties of a duplicated Font object. You can assign the object returned by the Duplicate property to another Font object to apply those settings all at once. Before assigning the duplicate object to another object, you can change any of the properties of the duplicate object without affecting the original. | https://learn.microsoft.com/en-us/office/vba/api/word.font.duplicate |
| Emboss | Long | R/W | True if the specified font is formatted as embossed. — Remarks: Returns True, False, or wdUndefined. Can be set to True, False, or wdToggle. Setting Emboss to True sets Engrave to False, and vice versa. | https://learn.microsoft.com/en-us/office/vba/api/word.font.emboss |
| EmphasisMark | WdEmphasisMark | R/W | Returns or sets a WdEmphasisMark constant that represents the emphasis mark for a character or designated character string. | https://learn.microsoft.com/en-us/office/vba/api/word.font.emphasismark |
| Engrave | Long | R/W | True if the font is formatted as engraved. — Remarks: Returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. Setting Engrave to True sets Emboss to False, and vice versa. | https://learn.microsoft.com/en-us/office/vba/api/word.font.engrave |
| Fill | FillFormat | R | Returns a FillFormat object that contains fill formatting properties for the font used by the specified range of text. | https://learn.microsoft.com/en-us/office/vba/api/word.font.fill |
| Glow | GlowFormat | R | Returns a GlowFormat object that represents the glow formatting for the font used by the specified range of text. | https://learn.microsoft.com/en-us/office/vba/api/word.font.glow |
| Grow() | no parameters | method | Increases the font size to the next available size. — Remarks: If the selection or range contains more than one font size, each size is increased to the next available setting. | https://learn.microsoft.com/en-us/office/vba/api/word.font.grow |
| Hidden | Long | R/W | True if the font is formatted as hidden text. — Remarks: This property returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. To control the display of hidden text, use the ShowHiddenText property of the View object. To control whether properties and methods that return Range objects include or exclude hidden text when hidden text isn't displayed, use the IncludeHiddenText property of the TextRetrievalMode object. | https://learn.microsoft.com/en-us/office/vba/api/word.font.hidden |
| Italic | Long | R/W | True if the font or range is formatted as italic. — Remarks: This property returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. | https://learn.microsoft.com/en-us/office/vba/api/word.font.italic |
| ItalicBi | Long | R/W | True if the font or range is formatted as italic. — Remarks: This property returns True, False or wdUndefined (for a mixture of italic and non-italic text) and can be set to True, False, or wdToggle. Use the ItalicBi property for right-to-left languages. | https://learn.microsoft.com/en-us/office/vba/api/word.font.italicbi |
| Kerning | Single | R/W | Returns or sets the minimum font size for which Microsoft Word will adjust kerning automatically. | https://learn.microsoft.com/en-us/office/vba/api/word.font.kerning |
| Ligatures | WdLigatures | R/W | Returns or sets the ligatures setting for the specified Font object. — Remarks: Open Type fonts support the use of ligatures. The Ligatures property specifies the Word ligatures setting to apply to an Open Type font. The following table lists the four basic values for ligatures. \| Value \| Description \| Standard \| Designed to enhance readability and attractiveness. Standard ligatures in Latin languages include "fi", "fl", and "ff". \| Contextual \| Designed to enhance readability and attractivenes… | https://learn.microsoft.com/en-us/office/vba/api/word.font.ligatures |
| Line | LineFormat | R/W | Returns a LineFormat object that specifies the formatting for a line. | https://learn.microsoft.com/en-us/office/vba/api/word.font.line |
| Name | String | R/W | Returns or sets the name of the specified object. | https://learn.microsoft.com/en-us/office/vba/api/word.font.name |
| NameAscii | String | R/W | Returns or sets the font used for Latin text (characters with character codes from 0 (zero) through 127). — Remarks: In the U.S. English version of Microsoft Word, the default value of this property is Times New Roman. Use the Name property to change the font that's applied to all text and that appears in the Font box on the Formatting toolbar. | https://learn.microsoft.com/en-us/office/vba/api/word.font.nameascii |
| NameBi | String | R/W | Returns or sets the name of the font in a right-to-left language document. | https://learn.microsoft.com/en-us/office/vba/api/word.font.namebi |
| NameFarEast | String | R/W | Returns or sets an East Asian font name. — Remarks: In the U.S. English version of Microsoft Word, the default value of this property is Times New Roman. This is the recommended way to return or set the font for Asian text in a document created in an Asian version of Word. | https://learn.microsoft.com/en-us/office/vba/api/word.font.namefareast |
| NameOther | String | R/W | Returns or sets the font used for characters with character codes from 128 through 255. — Remarks: In the U.S. English version of Microsoft Word, the default value of this property is Times New Roman. Use the Name property to change the font that's applied to all text and that appears in the Font box on the Formatting toolbar. | https://learn.microsoft.com/en-us/office/vba/api/word.font.nameother |
| NumberForm | WdNumberForm | R/W | Returns or sets the number form setting for an OpenType font. — Remarks: Numbers in OpenType fonts can be displayed either with consistent heights along the baseline of the text (called "lining"), or with varying heights (called "hanging" or "old style") where numbers are displayed above or below the baseline of the text. Use the NumberForm property to specify whether numbers are displayed using lining or old-style. Setting this property has the same effect as selecting an item in the dr… | https://learn.microsoft.com/en-us/office/vba/api/word.font.numberform |
| NumberSpacing | WdNumberSpacing | R/W | Returns or sets the number spacing setting for a font. — Remarks: OpenType fonts support a proportional and tabular figure feature to control number spacing. Proportional number spacing handles each number as having a different width. For example, "1" is displayed as narrower than "5". Tabular number spacing handles numbers as equal in width so that they align vertically, which increases the readability, especially for financial information. | https://learn.microsoft.com/en-us/office/vba/api/word.font.numberspacing |
| Outline | Long | R/W | True if the font is formatted as outline. — Remarks: Returns True, False, or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. | https://learn.microsoft.com/en-us/office/vba/api/word.font.outline |
| Position | Long | R/W | Returns or sets the position of text (in points) relative to the base line. — Remarks: A positive number raises the text, and a negative number lowers it. | https://learn.microsoft.com/en-us/office/vba/api/word.font.position |
| Reflection | ReflectionFormat | R | Returns a ReflectionFormat object that represents the reflection formatting for a shape. | https://learn.microsoft.com/en-us/office/vba/api/word.font.reflection |
| Reset() | no parameters | method | Removes manual character formatting (formatting not applied using a style). For example, if you manually format a word as bold and the underlying style is plain text (not bold), the Reset method removes the bold format. | https://learn.microsoft.com/en-us/office/vba/api/word.font.reset |
| Scaling | Long | R/W | Returns or sets the scaling percentage applied to the font. — Remarks: This property stretches or compresses text horizontally as a percentage of the current size (the scaling range is from 1 through 600). | https://learn.microsoft.com/en-us/office/vba/api/word.font.scaling |
| SetAsTemplateDefault() | no parameters | method | Sets the specified font formatting as the default for the active document and all new documents based on the active template. — Remarks: The default font formatting is stored in the Normal style. | https://learn.microsoft.com/en-us/office/vba/api/word.font.setastemplatedefault |
| Shading | Shading | not stated | Returns a Shading object that refers to the shading formatting for the specified font. | https://learn.microsoft.com/en-us/office/vba/api/word.font.shading |
| Shadow | Long | R/W | True if the specified font is formatted as shadowed. — Remarks: This property can be True, False, or wdUndefined. | https://learn.microsoft.com/en-us/office/vba/api/word.font.shadow |
| Shrink() | no parameters | method | Decreases the font size to the next available size. — Remarks: If the selection or range contains more than one font size, each size is decreased to the next available setting. | https://learn.microsoft.com/en-us/office/vba/api/word.font.shrink |
| Size | Single | R/W | Returns or sets the font size, in points. | https://learn.microsoft.com/en-us/office/vba/api/word.font.size |
| SizeBi | Single | R/W | Returns or sets the font size in points. — Remarks: The SizeBi property applies to text in a right-to-left language. | https://learn.microsoft.com/en-us/office/vba/api/word.font.sizebi |
| SmallCaps | Long | R/W | True if the font is formatted as small capital letters. — Remarks: Returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. Setting the SmallCaps property to True sets the AllCaps property to False, and vice versa. | https://learn.microsoft.com/en-us/office/vba/api/word.font.smallcaps |
| Spacing | Single | R/W | Returns or sets the spacing (in points) between characters. | https://learn.microsoft.com/en-us/office/vba/api/word.font.spacing |
| StrikeThrough | Long | R/W | True if the font is formatted as strikethrough text. — Remarks: The StrikeThrough property returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. To set or return double strikethrough formatting, use the DoubleStrikeThrough property. | https://learn.microsoft.com/en-us/office/vba/api/word.font.strikethrough |
| StylisticSet | WdStylisticSet | R/W | Specifies the stylistic set for the specified font. — Remarks: Some OpenType fonts provide stylistic sets. A stylistic set defines a set of characters within the font that are intended to be used together, usually for the purpose of visual harmony, such as in headings. | https://learn.microsoft.com/en-us/office/vba/api/word.font.stylisticset |
| Subscript | Long | R/W | True if the font is formatted as subscript. — Remarks: Returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. Setting the Subscript property to True sets the Superscript property to False, and vice versa. | https://learn.microsoft.com/en-us/office/vba/api/word.font.subscript |
| Superscript | Long | R/W | True if the font is formatted as superscript. — Remarks: Returns True, False, or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle. Setting the Superscript property to True sets the Subscript property to False, and vice versa. | https://learn.microsoft.com/en-us/office/vba/api/word.font.superscript |
| TextColor | ColorFormat | R | Returns a ColorFormat object that represents the color for the specified font. | https://learn.microsoft.com/en-us/office/vba/api/word.font.textcolor |
| TextShadow | ShadowFormat | not stated | Returns a ShadowFormat object that specifies the shadow formatting for the specified font. | https://learn.microsoft.com/en-us/office/vba/api/word.font.textshadow |
| ThreeD | ThreeDFormat | R | Returns a ThreeDFormat object that contains 3D effect formatting properties for the specified font. | https://learn.microsoft.com/en-us/office/vba/api/word.font.threed |
| Underline | WdUnderline | R/W | Returns or sets the type of underline applied to the font. | https://learn.microsoft.com/en-us/office/vba/api/word.font.underline |
| UnderlineColor | WdColor | R/W ("returns or sets"; no mode word) | Returns or sets the 24-bit color of the underline for the specified Font object. . — Remarks: This property can be any valid WdColor constant or a value returned by Visual Basic's RGB function. Setting the UnderlineColor property to wdColorAutomatic resets the color of the underline to the color of the text above it. | https://learn.microsoft.com/en-us/office/vba/api/word.font.underlinecolor |
| Color | WdColor (24-bit, BGR) | R/W | Not in the current VBA reference (word.font.color is 404). See A.5 and Discrepancy 8. | https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb238769(v=office.12) , https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._font.color?view=word-pia |

Related objects and members:

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Brightness | Single | R/W | Returns a Single that represents the brightness of a specified shape color. — Remarks: You can enter a number from -1 (darkest) to 1 (lightest) for the Brightness property, 0 (zero) being neutral. | https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.brightness |
| ObjectThemeColor | WdThemeColorIndex | R/W | Returns or sets a WdThemeColorIndex constant that represents the theme color for a color format. | https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.objectthemecolor |
| RGB | Long | R/W | Returns or sets the red-green-blue (RGB) value of the specified color. | https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.rgb |
| TintAndShade | Single | R/W | Returns a Single that represents the lightening or darkening of a specified shape's color. — Remarks: You can enter a number from -1 (darkest) to 1 (lightest) for the TintAndShade property, 0 (zero) being neutral. | https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.tintandshade |
| Type | MsoColorType | R | Returns or sets the shape color type. | https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.type |
| HighlightColorIndex | WdColorIndex | R/W | Returns or sets the highlight color for the specified range. | https://learn.microsoft.com/en-us/office/vba/api/word.range.highlightcolorindex |
| BackgroundPatternColor | WdColor | R/W | Returns or sets the 24-bit color that's applied to the background of the Shading object. — Remarks: This property can be any valid WdColor constant or a value returned by Visual Basic's RGB function. | https://learn.microsoft.com/en-us/office/vba/api/word.shading.backgroundpatterncolor |
| BackgroundPatternColorIndex | WdColorIndex | R/W | Returns or sets the color that's applied to the background of the Shading object. | https://learn.microsoft.com/en-us/office/vba/api/word.shading.backgroundpatterncolorindex |
| ForegroundPatternColor | WdColor | R/W | Returns or sets the 24-bit color that's applied to the foreground of the Shading object. This color is applied to the dots and lines in the shading pattern. — Remarks: This property can be any valid WdColor constant or a value returned by Visual Basic's RGB function. | https://learn.microsoft.com/en-us/office/vba/api/word.shading.foregroundpatterncolor |
| ForegroundPatternColorIndex | WdColorIndex | R/W | Returns or sets the color that's applied to the foreground of the Shading object. This color is applied to the dots and lines in the shading pattern. | https://learn.microsoft.com/en-us/office/vba/api/word.shading.foregroundpatterncolorindex |
| Texture | WdTextureIndex | R/W | Returns or sets the shading texture for the specified object. | https://learn.microsoft.com/en-us/office/vba/api/word.shading.texture |

- **Text colour, two ways.** `Font.Color` takes a 24-bit WdColor/RGB value. `Font.TextColor` returns a ColorFormat, which is read-only as an object, with `RGB` (R/W Long), `ObjectThemeColor` (R/W WdThemeColorIndex) and `TintAndShade` (R/W Single, -1 darkest to 1 lightest). How `ColorFormat.RGB` orders its bytes is **not stated** on its page (https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.rgb).
- **Highlight** is `Range.HighlightColorIndex` (R/W WdColorIndex; wdNoHighlight = 0). Font has no highlight member (https://learn.microsoft.com/en-us/office/vba/api/word.range.highlightcolorindex, https://learn.microsoft.com/en-us/office/vba/api/word.wdcolorindex). The Font.ColorIndex page says "The wdByAuthor constant is not a valid color for fonts".
- **Mutually exclusive pairs** (from the pages): AllCaps/SmallCaps, Superscript/Subscript, StrikeThrough/DoubleStrikeThrough, Emboss/Engrave. Setting one to True sets the other to False.
- **Scaling**: 1 to 600 (%). **Position**: points, positive raises. **Spacing**: points between characters. **Kerning**: "the minimum font size for which Microsoft Word will adjust kerning automatically".

### C.5 TabStops, Borders

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Add() | (Position: Single, Alignment?: Variant, Leader?: Variant) → TabStop | method | Returns a TabStop object that represents a custom tab stop added to a document. — Params: Position: The position of the tab stop (in points) relative to the left margin.; Alignment: The alignment of the tab stop. Can be one of the WdTabAlignment constants.; Leader: The type of leader for the tab stop. Can be one of the WdTabLeader constants. If this argument is omitted, wdTabLeaderSpaces is used. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.add |
| After() | (Position: Single) | method | Returns the next TabStop object to the right of Position. — Params: Position: A location on the ruler, in points. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.after |
| Before() | (Position: Single) | method | Returns the next TabStop object to the left of Position. — Params: Position: A location on the ruler, in points. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.before |
| ClearAll() | no parameters | method | Clears all the custom tab stops from the specified paragraphs. — Remarks: To clear an individual tab stop, use the Clear method of the TabStop object. The ClearAll method doesn't clear the default tab stops. To manipulate the default tab stops, use the DefaultTabStop property for the document. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.clearall |
| Count | Long | R | Returns a Long that represents the number of tab stops in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.count |
| Item() | (Index: Variant) → TabStop | method | Returns an individual TabStop object in a collection. — Params: Index: The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual obj… | https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.item |
| Alignment | WdTabAlignment | R/W | Returns or sets a WdTabAlignment constant that represents the alignment for the specified tab stop. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.alignment |
| Clear() | no parameters | method | Removes the specified custom tab stop. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.clear |
| CustomTab | Boolean | R | True if the specified tab stop is a custom tab stop. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.customtab |
| Leader | WdTabLeader | R/W | Returns or sets the leader for the specified TabStop object. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.leader |
| Next | not stated | R | Returns the next tabstop in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.next |
| Position | Single | R/W | Returns or sets the position of a tab stop relative to the left margin. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.position |
| Previous | not stated | R | Returns the previous tab stop in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.previous |

Borders collection members (from the table of contents): alwaysinfront, application, applypageborderstoallsections, count, creator, distancefrom, distancefrombottom, distancefromleft, distancefromright, distancefromtop, enable, enablefirstpageinsection, enableotherpagesinsection, hashorizontal, hasvertical, insidecolor, insidecolorindex, insidelinestyle, insidelinewidth, item, joinborders, outsidecolor, outsidecolorindex, outsidelinestyle, outsidelinewidth, parent, shadow, surroundfooter, surroundheader. Object page: https://learn.microsoft.com/en-us/office/vba/api/word.borders. Style.Borders, ParagraphFormat.Borders and Font.Borders all return this collection.

## D. Lists and bullets

### ListFormat (Range.ListFormat)

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.listformat.

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| ApplyBulletDefault() | (DefaultListBehavior?: Variant) | method | Adds bullets and formatting to the paragraphs in the range for the specified ListFormat object. — Remarks: If the paragraphs are already formatted with bullets, this method removes the bullets and formatting. — Params: DefaultListBehavior: Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following… | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applybulletdefault |
| ApplyListTemplate() | (ListTemplate: ListTemplate, ContinuePreviousList?: Variant, ApplyTo?: Variant, DefaultListBehavior?: Variant) | method | Applies a set of list-formatting characteristics to the specified ListFormat object. — Params: ListTemplate: The list template to be applied.; ContinuePreviousList: True to continue the numbering from the previous list; False to start a new list.; ApplyTo: The portion of the list that the list template is to be applied to. Can be one of the following WdListApplyTo constants: wdListSelection, w…; DefaultListBehavior: Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following… | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applylisttemplate |
| ApplyListTemplateWithLevel() | (ListTemplate: ListTemplate, ContinuePreviousList?: Variant, ApplyTo?: Variant, DefaultListBehavior?: Variant, ApplyLevel?: Variant) | method | Applies a set of list-formatting characteristics, optionally for a specified level. — Params: ListTemplate: The list template to be applied.; ContinuePreviousList: True to continue the numbering from the previous list; False to start a new list.; ApplyTo: One of the WdListApplyTo constants that specifies the portion of the list that the list template will be applied to.; DefaultListBehavior: Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following…; ApplyLevel: The level to which the list template is to be applied. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applylisttemplatewithlevel |
| ApplyNumberDefault() | (DefaultListBehavior?: Variant) | method | Adds the default numbering scheme to the paragraphs in the range for the specified ListFormat object. — Remarks: If the paragraphs are already formatted as a numbered list, this method removes the numbers and formatting. — Params: DefaultListBehavior: Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following… | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applynumberdefault |
| ApplyOutlineNumberDefault() | (DefaultListBehavior?: Variant) | method | Adds the default outline-numbering scheme to the paragraphs in the range for the specified ListFormat object. — Remarks: If the paragraphs are already formatted as an outline-numbered list, this method removes the numbers and formatting. This method doesn't remove built-in heading styles that have been applied to paragraphs. — Params: DefaultListBehavior: Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following… | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applyoutlinenumberdefault |
| CanContinuePreviousList() | (ListTemplate: ListTemplate) | method | Returns a WdContinue constant (wdContinueDisabled, wdResetList, or wdContinueList) that indicates whether the formatting from the previous list can be continued. — Remarks: This method returns the state of the Continue previous list and Restart numbering options in the Bullets and Numbering dialog box for a specified list format. To change the settings of these options, set the ContinuePreviousList argument of the ApplyListTemplate method. — Params: ListTemplate: A list template that's been applied to previous paragraphs in the document. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.cancontinuepreviouslist |
| ConvertNumbersToText() | no parameters | method | Changes the list numbers and LISTNUM fields in the specified ListFormat object to text. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.convertnumberstotext |
| CountNumberedItems() | no parameters | method | Returns the number of bulleted or numbered items and LISTNUM fields in the specified ListFormat object. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.countnumbereditems |
| List | List | not stated | Returns a List object that represents the first formatted list contained in the specified ListFormat object. — Remarks: If the first paragraph in the range for the ListFormat object is not formatted as a list, the List property returns nothing. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.list |
| ListIndent() | no parameters | method | Increases the list level of the paragraphs in the range for the specified ListFormat object, in increments of one level. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listindent |
| ListLevelNumber | Long | R/W | Returns or sets the list level for the first paragraph in the specified ListFormat object. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listlevelnumber |
| ListOutdent() | no parameters | method | Decreases the list level of the paragraphs in the range for the specified ListFormat object, in increments of one level. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listoutdent |
| ListPictureBullet | InlineShape | not stated | Returns the InlineShape object that represents the picture used as a bullet in a picture bulleted list. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listpicturebullet |
| ListString | String | R | Returns a String that represents the appearance of the list value of the first paragraph in the range for the specified ListFormat object. For example, the second paragraph in an alphabetical list would return B. — Remarks: For a bulleted list, you will need to apply the correct font to see the string. Most bullets use the Symbol or Wingdings font. Use the ListValue property to return the numeric value of the paragraph. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.liststring |
| ListTemplate | ListTemplate | not stated | Returns a ListTemplate object that represents the list formatting for the specified ListFormat object. — Remarks: A list template includes all the formatting that defines a particular list. Each of the seven formats (excluding None) found on each of the tabs in the Bullets and Numbering dialog box corresponds to a list template. Documents and templates can also contain collections of list templates. If the first paragraph in the range for the ListFormat object is not formatted as a list, the ListTemplate property returns Nothing. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listtemplate |
| ListType | WdListType | R | Returns the type of lists that are contained in the range for the specified ListFormat object. — Remarks: The constant wdListListNumOnly refers to LISTNUM fields, which are fields that can be added within the text of a paragraph. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listtype |
| ListValue | Long | R | Returns the numeric value of the first paragraph in the range for the specified ListFormat object. — Remarks: Use the ListString property to return a string that represents the list value. If the ListFormat object applies to a bulleted list, the ListValue property returns 1. If the ListFormat object applies to an outline-numbered list, the ListValue property returns the numeric value of the first paragraph as it occurs in the sequence of paragraphs at the same level. For example, if the first paragraph for a specified ListFormat object were numbered "A.2," the ListValue property would return 2. This pr… | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listvalue |
| RemoveNumbers() | (NumberType?: WdNumberType) | method | Removes numbers or bullets from the specified list. — Params: NumberType: The type of number to be removed. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.removenumbers |
| SingleList | Boolean | R | True if the specified ListFormat object contains only one list. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.singlelist |
| SingleListTemplate | Boolean | R | True if the entire ListFormat object uses the same list template. | https://learn.microsoft.com/en-us/office/vba/api/word.listformat.singlelisttemplate |

### ListTemplates

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.listtemplates.

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Add() | (OutlineNumbered?: Variant, Name?: Variant) → ListTemplate | method | Returns a ListTemplate object that represents a new list template. — Remarks: You cannot use the Add method on ListTemplates objects returned from a ListGallery object. You can, however, modify the existing list templates in the galleries. — Params: OutlineNumbered: True to apply outline numbering to the new list template.; Name: An optional name used for linking the list template to a LISTNUM field. Use this name to index the list template in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.listtemplates.add |
| Count | Long | R | Returns a Long that represents the number of list templates in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.listtemplates.count |
| Item() | (Index: Variant) → ListTemplate | method | Returns an individual ListTemplate object in a collection. — Params: Index: The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual obj… | https://learn.microsoft.com/en-us/office/vba/api/word.listtemplates.item |

### ListTemplate

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate.

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Convert() | (Level?: Variant) | method | Converts a multiple-level list to a single-level list, or vice versa. — Remarks: You cannot use the Convert method on a list template that is derived from the ListGalleries collection. — Params: Level: The level to use for formatting the new list. When converting a multiple-level list to a single-level list, this argument can be a number f… | https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate.convert |
| ListLevels | ListLevels | not stated | Returns a ListLevels collection that represents all the levels for the specified ListTemplate. | https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate.listlevels |
| Name | String | R/W | Returns or sets the name of the specified object. | https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate.name |
| OutlineNumbered | Boolean | R/W | True if the specified ListTemplate object is outline numbered. — Remarks: Setting this property to False converts the list template to a single-level list that uses the formatting of the first level. You cannot set this property for a ListTemplate object returned from a ListGallery object. | https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate.outlinenumbered |

### ListLevels

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.listlevels.

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Count | Long | R | Returns a Long that represents the number of list levels in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevels.count |
| Item() | (Index: Long) → ListLevel | method | Returns an individual ListLevel object in a collection. — Params: Index: The individual object to be returned. Can be a Long indicating the ordinal position of the individual object. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevels.item |

### ListLevel

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Alignment | WdListLevelAlignment | R/W | Returns or sets a WdListLevelAlignment constant that represents the alignment for the list level of the list template. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.alignment |
| ApplyPictureBullet() | (FileName: String) | method | Formats a paragraph or range of paragraphs with a picture bullet. — Params: FileName: The path and file name of the picture file. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.applypicturebullet |
| Font | Font | R/W | Returns or sets a Font object that represents the character formatting of the specified object. — Remarks: To set this property, specify an expression that returns a Font object. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.font |
| Index | Long | R | Returns a Long that represents the position of an item in a collection. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.index |
| LinkedStyle | String | R/W | Returns or sets the name of the style that's linked to the specified ListLevel object. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.linkedstyle |
| NumberFormat | String | R/W | Returns or sets the number format for the specified list level. — Remarks: The percent sign (%) followed by any number from 1 through 9 represents the number style from the respective list level. For example, if you wanted the format for the first level to be "Article I," "Article II," and so on, the string for the NumberFormat property would be "Article %1" and the NumberStyle property would be set to wdListNumberStyleUpperCaseRoman. If the NumberStyle property is set to wdListNumberStyleBullet, the string for the NumberFormat property can only contain one character. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.numberformat |
| NumberPosition | Single | R/W | Returns or sets the position (in points) of the number or bullet for the specified ListLevel object. — Remarks: For each list level, you can set the position of the number or bullet, the position of the tab, and the position of the text that wraps. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.numberposition |
| NumberStyle | WdListNumberStyle | R/W | Returns or sets the number style for the ListLevel object. — Remarks: Some of the WdListNumberStyle constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.numberstyle |
| PictureBullet | InlineShape | not stated | Returns an InlineShape object that represents a picture bullet. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.picturebullet |
| ResetOnHigher | Long | R/W | Sets or returns the list level that must appear before the specified list level restarts numbering at 1. — Remarks: The ResetOnHigher property returns False if the numbering continues sequentially each time the list level appears. This feature allows lists to be interleaved, maintaining numeric sequence. You cannot set the ResetOnHigher property of a list level to a value greater than or equal to its index in the ListLevels collection. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.resetonhigher |
| StartAt | Long | R/W | Returns or sets the starting number for the specified ListLevel object. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.startat |
| TabPosition | Single | R/W | Returns or sets the tab position for the specified ListLevel object. — Remarks: Because the ListLevel object does not have a default tab setting, the TabPosition property always returns a value of 999999 or wdUndefined, unless you set the property to a value. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.tabposition |
| TextPosition | Single | R/W | Returns or sets the position (in points) for the second line of wrapping text for the specified ListLevel object. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.textposition |
| TrailingCharacter | WdTrailingCharacter | R/W | Returns or sets the character inserted after the number for the specified list level. | https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.trailingcharacter |

### ListGalleries

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.listgalleries.

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| Count | Long | R | Returns a Long that represents the number of list galleries in the collection. | https://learn.microsoft.com/en-us/office/vba/api/word.listgalleries.count |
| Item() | (Index: WDLISTGALLERYTYPE) → ListGallery | method | Returns an individual ListGalleries object in a collection. — Params: Index: The individual object to be returned. | https://learn.microsoft.com/en-us/office/vba/api/word.listgalleries.item |

### ListGallery

Object page: https://learn.microsoft.com/en-us/office/vba/api/word.listgallery.

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| ListTemplates | ListTemplates | R | Returns a ListTemplates collection that represents all the list formats for the specified list gallery. — Remarks: The ListTemplates property is a member of the Document, ListGallery, and Template objects. | https://learn.microsoft.com/en-us/office/vba/api/word.listgallery.listtemplates |
| Modified | Boolean | R | True if the specified list template is not the built-in list template for that position in the list gallery. — Remarks: Use the Reset method to set a list template in a list gallery back to the built-in list template. | https://learn.microsoft.com/en-us/office/vba/api/word.listgallery.modified |
| Reset() | (Index: Long) | method | Resets the list template specified by Index for the specified list gallery to the built-in list template format. — Params: Index: The template to reset. | https://learn.microsoft.com/en-us/office/vba/api/word.listgallery.reset |

### List object, and the properties that lead to the above

| Member | Type / values | R/W | Notes (from the page) | Source URL |
|---|---|---|---|---|
| List object (Word) | not stated | not stated | Represents a single list format that's been applied to specified paragraphs in a document. The List object is a member of the Lists collection. — Remarks: Use Lists (Index), where Index is the index number, to return a single List object. The following example returns the number of items in list one in the active document. mycount = ActiveDocument.Lists(1).CountNumberedItems To return all the paragraphs that have list formatting, use the ListParagraphs property. To return them as a range, use the Range property. To apply a different list format to an existing list, use the ApplyListTemplate method with the List object. To add a new list to a docu… | https://learn.microsoft.com/en-us/office/vba/api/word.list |
| ApplyListTemplate() | (ListTemplate: ListTemplate, ContinuePreviousList?: Variant, ApplyTo?: Variant, DefaultListBehavior?: Variant) | method | Applies a set of list-formatting characteristics to the specified ListFormat object. — Params: ListTemplate: The list template to be applied.; ContinuePreviousList: True to continue the numbering from the previous list; False to start a new list.; ApplyTo: The portion of the list that the list template is to be applied to. Can be one of the following WdListApplyTo constants: wdListApplyToSelec…; DefaultListBehavior: Sets a value that specifies whether Microsoft Word uses new web-oriented formatting for better list display. Can be either of the following… | https://learn.microsoft.com/en-us/office/vba/api/word.list.applylisttemplate |
| ApplyListTemplateWithLevel() | (ListTemplate: ListTemplate, ContinuePreviousList?: Variant, DefaultListBehavior?: Variant, ApplyLevel?: Variant) | method | Applies a set of list-formatting characteristics, optionally for a specified level. — Params: ListTemplate: The list template to be applied.; ContinuePreviousList: True to continue the numbering from the previous list; False to start a new list.; DefaultListBehavior: Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following…; ApplyLevel: The level to which the list template is to be applied. | https://learn.microsoft.com/en-us/office/vba/api/word.list.applylisttemplatewithlevel |
| Range | Range | not stated | Returns a Range object that represents the portion of a document that is contained in the specified object. | https://learn.microsoft.com/en-us/office/vba/api/word.list.range |
| Lists | Lists | R | Returns a Lists collection that contains all the formatted lists in the specified document. | https://learn.microsoft.com/en-us/office/vba/api/word.document.lists |
| ListTemplates | ListTemplates | R | Returns a ListTemplates collection that represents all the list formats for the specified document. — Remarks: The ListTemplates property is a member of the Document, ListGallery, and Template objects. | https://learn.microsoft.com/en-us/office/vba/api/word.document.listtemplates |
| ListGalleries | ListGalleries | not stated | Returns a ListGalleries collection that represents the three list template galleries. . — Remarks: Each template gallery (Bulleted, Numbered, and Outline Numbered) corresponds to a tab in the Bullets and Numbering dialog box (Format menu). | https://learn.microsoft.com/en-us/office/vba/api/word.application.listgalleries |
| ListFormat | ListFormat | R | Returns a ListFormat object that represents all the list formatting characteristics of a range. | https://learn.microsoft.com/en-us/office/vba/api/word.range.listformat |
| ListParagraphs | ListParagraphs | R | Returns a ListParagraphs collection that represents all the numbered paragraphs in the range. | https://learn.microsoft.com/en-us/office/vba/api/word.range.listparagraphs |

### D.1 Changing one list's bullet character and colour without touching other lists

What the documentation supports, step by step:

1. **Do not edit a gallery template if only one list should change.** `ListGalleries(wdBulletGallery).ListTemplates(1..7)` are the seven formats on the Bulleted tab of Bullets and Numbering (https://learn.microsoft.com/en-us/office/vba/api/word.listgallery). The pages say "You can modify these list templates". `ListGallery.Modified` then reports that a template "is not the built-in list template for that position", and `ListGallery.Reset(Index)` restores it (https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate, https://learn.microsoft.com/en-us/office/vba/api/word.listgallery.modified, https://learn.microsoft.com/en-us/office/vba/api/word.listgallery.reset). The galleries are reached from the Application object (`Application.ListGalleries`, "the three list template galleries", https://learn.microsoft.com/en-us/office/vba/api/word.application.listgalleries). Whether editing a gallery template changes lists that were already formatted with it, or lists in other documents, is **not stated**.
2. **Create a template in the document:** `Set lt = doc.ListTemplates.Add(OutlineNumbered:=False)` ("Returns a ListTemplate object that represents a new list template"; Add does not work on a gallery's ListTemplates) (https://learn.microsoft.com/en-us/office/vba/api/word.listtemplates.add).
3. **Set the bullet on the level:** `lt.ListLevels(1)` (https://learn.microsoft.com/en-us/office/vba/api/word.listlevels.item), then:
   - `.NumberStyle = wdListNumberStyleBullet` (23) (https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.numberstyle, https://learn.microsoft.com/en-us/office/vba/api/word.wdlistnumberstyle);
   - `.NumberFormat = ChrW(code)`. "If the NumberStyle property is set to wdListNumberStyleBullet, the string for the NumberFormat property can only contain one character" (https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.numberformat);
   - `.Font.Name = "Wingdings"` / `"Symbol"` etc. The ListFormat.ListString page says "For a bulleted list, you will need to apply the correct font to see the string. Most bullets use the Symbol or Wingdings font" (https://learn.microsoft.com/en-us/office/vba/api/word.listformat.liststring);
   - **Colour and size:** `.Font` is "a Font object that represents the character formatting of the specified object. Read/write Font" (https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.font). The ListLevel and ListTemplate pages list Font among "all the formatting properties for the specified list level" (https://learn.microsoft.com/en-us/office/vba/api/word.listlevel). So `.Font.Color` and `.Font.Size` are where the bullet's colour and size live. The pages do not mention bullet colour by name: they document a Font on the level, and Font has Color and Size.
   - Positions: `.NumberPosition` (points, "position … of the number or bullet"), `.TextPosition` (points, "second line of wrapping text"), `.TabPosition` ("always returns a value of 999999 or wdUndefined, unless you set the property"), `.TrailingCharacter` (WdTrailingCharacter: tab 0, space 1, none 2), `.Alignment` (WdListLevelAlignment).
4. **Apply it to just the range:** `rng.ListFormat.ApplyListTemplate ListTemplate:=lt, ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:=wdWord9ListBehavior` (https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applylisttemplate). ContinuePreviousList False is documented as "start a new list". ApplyTo values: wdListApplyToWholeList 0, wdListApplyToThisPointForward 1, wdListApplyToSelection 2 (https://learn.microsoft.com/en-us/office/vba/api/word.wdlistapplyto). DefaultListBehavior: the page recommends wdWord9ListBehavior over the Word-97-compatible default. WdDefaultListBehavior also documents wdWord10ListBehavior = 2 ("compatible with Microsoft Word 2002") (https://learn.microsoft.com/en-us/office/vba/api/word.wddefaultlistbehavior). Use `ApplyListTemplateWithLevel(..., ApplyLevel)` to target a level (https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applylisttemplatewithlevel).
5. **To restyle an existing list**, the List page says: "To apply a different list format to an existing list, use the ApplyListTemplate method with the List object. To add a new list to a document, use the ApplyListTemplate method with the ListFormat object for a specified range" (https://learn.microsoft.com/en-us/office/vba/api/word.list). Get the list from `rng.ListFormat.List`, which is the first list in the range, or Nothing (https://learn.microsoft.com/en-us/office/vba/api/word.listformat.list).
6. **Editing the template a list already uses** (`rng.ListFormat.ListTemplate.ListLevels(n)…`): the ListTemplate it returns is "the list formatting for the specified ListFormat object" (https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listtemplate). Whether other lists that use the same template change too is **not stated**. Creating a new template (step 2) is the approach that leaves every existing template unchanged.

Reading a list back: `ListFormat.ListType` (WdListType: wdListBullet 2, wdListPictureBullet 6, …), `ListFormat.ListString` (the bullet character or number text), `ListFormat.ListLevelNumber` (R/W Long), `ListFormat.ListValue` (1 for a bulleted list).

### D.2 How a list ties to a style

- `Style.LinkToListTemplate(ListTemplate, ListLevelNumber?)`: "Links the specified style to a list template so that the style's formatting can be applied to lists." If ListLevelNumber is omitted, "the level of the style is used". The example links Heading 1–9 to levels 1–9 of a new outline-numbered template, after which "Any paragraphs formatted as heading styles will assume the numbering from the list template" (https://learn.microsoft.com/en-us/office/vba/api/word.style.linktolisttemplate).
- `ListLevel.LinkedStyle` (R/W String): "the name of the style that's linked to the specified ListLevel object". The example sets `mylevel.LinkedStyle = "Heading " & mylevel.index` on a gallery template (https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.linkedstyle). It takes a name, so in a localized Word use the style's NameLocal.
- `Style.ListTemplate` returns the style's list formatting, and `Style.ListLevelNumber` (read-only Long) its level (https://learn.microsoft.com/en-us/office/vba/api/word.style.listtemplate, https://learn.microsoft.com/en-us/office/vba/api/word.style.listlevelnumber).
- ListFormat.ApplyOutlineNumberDefault "doesn't remove built-in heading styles that have been applied to paragraphs" (https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applyoutlinenumberdefault).
- Toggles: ApplyBulletDefault and ApplyNumberDefault **remove** the list if the paragraphs already have one ("If the paragraphs are already formatted with bullets, this method removes the bullets and formatting", https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applybulletdefault). Use `RemoveNumbers(NumberType?)` to remove a list explicitly (https://learn.microsoft.com/en-us/office/vba/api/word.listformat.removenumbers).

## E. Behaviour questions

| Question | What the documentation says | Source URL |
|---|---|---|
| Does Window.ScrollIntoView move the selection? | **Not stated.** The page says only that it "Scrolls through the document window so the specified range or shape is displayed". Its example passes `Selection.Range`. | https://learn.microsoft.com/en-us/office/vba/api/word.window.scrollintoview |
| Does Range.InsertAfter expand the range? | **Yes:** "After this method is applied, the range expands to include the new text." Exception: "if the range is a bookmark, Word inserts the specified text but does not extend the range or the bookmark". InsertBefore also expands the range, and the bookmark with it. | https://learn.microsoft.com/en-us/office/vba/api/word.range.insertafter , https://learn.microsoft.com/en-us/office/vba/api/word.range.insertbefore |
| Does Range.Paragraphs include paragraphs only partly in the range? | **Not stated.** The page says "all the paragraphs in the specified range". | https://learn.microsoft.com/en-us/office/vba/api/word.range.paragraphs |
| What does Document.Undo return? | Boolean: "Returns True if the actions were successfully undone." Its parameter is Times (Variant, optional). What happens when fewer than Times actions can be undone is **not stated**. Redo behaves the same way. | https://learn.microsoft.com/en-us/office/vba/api/word.document.undo , https://learn.microsoft.com/en-us/office/vba/api/word.document.redo |
| Is UndoRecord (StartCustomRecord / EndCustomRecord / IsRecordingCustomRecord) available on Word for Mac? | **Not stated.** None of the UndoRecord pages or the concept page mentions a platform or a version. | https://learn.microsoft.com/en-us/office/vba/api/word.undorecord , https://learn.microsoft.com/en-us/office/vba/api/word.application.undorecord , https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-the-undorecord-object |
| What does Font.Reset do? | "Removes manual character formatting (formatting not applied using a style). For example, if you manually format a word as bold and the underlying style is plain text (not bold), the Reset method removes the bold format." ParagraphFormat.Reset does the same for paragraph formatting. | https://learn.microsoft.com/en-us/office/vba/api/word.font.reset , https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.reset |
| What does Font.Name set, versus NameAscii / NameOther / NameFarEast / NameBi? | Name: "Returns or sets the name of the specified object." Which script slots it writes is **not stated**. NameAscii: Latin text, character codes 0–127. NameOther: character codes 128–255. Both pages say "Use the Name property to change the font that's applied to all text and that appears in the Font box". NameFarEast: "an East Asian font name", "the recommended way to return or set the font for Asian text in a document created in an Asian version of Word". NameBi: "the name of the font in a right-to-left language document". | https://learn.microsoft.com/en-us/office/vba/api/word.font.name , https://learn.microsoft.com/en-us/office/vba/api/word.font.nameascii , https://learn.microsoft.com/en-us/office/vba/api/word.font.nameother , https://learn.microsoft.com/en-us/office/vba/api/word.font.namefareast , https://learn.microsoft.com/en-us/office/vba/api/word.font.namebi |
| Does Paragraph.Range.Text include the paragraph mark? | The Paragraph.Range page does not say so directly. The Range.InsertAfter page implies it: "If you use this method with a range that refers to an entire paragraph, the text is inserted after the ending paragraph mark", and to insert at the end of a paragraph, "subtract 1 from this location (the paragraph mark is one character)". Range.Collapse agrees: collapsing a whole-paragraph range to its end lands "after the ending paragraph mark". Which character the mark is (Chr(13)) is **not stated** on these pages. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.range , https://learn.microsoft.com/en-us/office/vba/api/word.range.insertafter , https://learn.microsoft.com/en-us/office/vba/api/word.range.collapse |
| What does Range.Text give at the end of a table cell (Chr(13) & Chr(7))? | The concept page has a section titled "Returning text from a table cell without returning the end of cell marker". It drops the marker with `Range(Start:=cell.Range.Start, End:=cell.Range.End - 1)` or `MoveEnd wdCharacter, -1`. So the docs confirm there is an end-of-cell marker and that removing one character from the end drops it. That the marker is Chr(13) & Chr(7) is **not stated**. | https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-tables , https://learn.microsoft.com/en-us/office/vba/api/word.cell.range |
| Does Font.Bold read back as True / -1? | Font.Bold is "Read/write Long" and "Returns True, False or wdUndefined (a mixture of True and False). Can be set to True, False, or wdToggle." A VBA Boolean True converts to -1 as a number. wdToggle = 9999998 and wdUndefined = 9999999. | https://learn.microsoft.com/en-us/office/vba/api/word.font.bold , https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/boolean-data-type , https://learn.microsoft.com/en-us/office/vba/api/word.wdconstants |
| Is LineSpacing for wdLineSpaceMultiple 12 points per line? | **Yes.** LineSpacing is in points. For wdLineSpaceMultiple "a LineSpacing property value must be specified, in points". LinesToPoints: "1 line = 12 points", and LinesToPoints(2) returns 24. | https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.linespacing , https://learn.microsoft.com/en-us/office/vba/api/word.application.linestopoints |
| Does Find through a Range move the selection? | No: "If you've gotten to the Find object from the Range object, the selection isn't changed when text matching the find criteria is found, but the Range object is redefined." | https://learn.microsoft.com/en-us/office/vba/api/word.find |
| Does SaveAs2 overwrite an existing file? | Yes: "If a document with the specified file name already exists, the document is overwritten without prompting the user." | https://learn.microsoft.com/en-us/office/vba/api/word.saveas2 |

## Appendix: pages fetched and cited

404 while writing this file: `https://learn.microsoft.com/en-us/office/vba/api/word.font.color` and `https://learn.microsoft.com/en-us/office/vba/api/word.font.usediffdiaccolor`, which Font.DiacriticColor refers to.

- https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._font.color?view=word-pia
- https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._font?view=word-pia
- https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._paragraphformat.disablelineheightgrid?view=word-pia
- https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdstyletype?view=word-pia
- https://learn.microsoft.com/en-us/office/vba/api/word.application.activedocument
- https://learn.microsoft.com/en-us/office/vba/api/word.application.build
- https://learn.microsoft.com/en-us/office/vba/api/word.application.documents
- https://learn.microsoft.com/en-us/office/vba/api/word.application.linestopoints
- https://learn.microsoft.com/en-us/office/vba/api/word.application.listgalleries
- https://learn.microsoft.com/en-us/office/vba/api/word.application.name
- https://learn.microsoft.com/en-us/office/vba/api/word.application.undorecord
- https://learn.microsoft.com/en-us/office/vba/api/word.application.version
- https://learn.microsoft.com/en-us/office/vba/api/word.application.visible
- https://learn.microsoft.com/en-us/office/vba/api/word.borders
- https://learn.microsoft.com/en-us/office/vba/api/word.cell.range
- https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.brightness
- https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.objectthemecolor
- https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.rgb
- https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.tintandshade
- https://learn.microsoft.com/en-us/office/vba/api/word.colorformat.type
- https://learn.microsoft.com/en-us/office/vba/api/word.document.close(method)
- https://learn.microsoft.com/en-us/office/vba/api/word.document.computestatistics
- https://learn.microsoft.com/en-us/office/vba/api/word.document.content
- https://learn.microsoft.com/en-us/office/vba/api/word.document.copystylesfromtemplate
- https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat
- https://learn.microsoft.com/en-us/office/vba/api/word.document.fullname
- https://learn.microsoft.com/en-us/office/vba/api/word.document.lists
- https://learn.microsoft.com/en-us/office/vba/api/word.document.listtemplates
- https://learn.microsoft.com/en-us/office/vba/api/word.document.name
- https://learn.microsoft.com/en-us/office/vba/api/word.document.pagesetup
- https://learn.microsoft.com/en-us/office/vba/api/word.document.paragraphs
- https://learn.microsoft.com/en-us/office/vba/api/word.document.path
- https://learn.microsoft.com/en-us/office/vba/api/word.document.range
- https://learn.microsoft.com/en-us/office/vba/api/word.document.readonly
- https://learn.microsoft.com/en-us/office/vba/api/word.document.redo
- https://learn.microsoft.com/en-us/office/vba/api/word.document.save
- https://learn.microsoft.com/en-us/office/vba/api/word.document.saved
- https://learn.microsoft.com/en-us/office/vba/api/word.document.sections
- https://learn.microsoft.com/en-us/office/vba/api/word.document.styles
- https://learn.microsoft.com/en-us/office/vba/api/word.document.tables
- https://learn.microsoft.com/en-us/office/vba/api/word.document.trackrevisions
- https://learn.microsoft.com/en-us/office/vba/api/word.document.undo
- https://learn.microsoft.com/en-us/office/vba/api/word.document.undoclear
- https://learn.microsoft.com/en-us/office/vba/api/word.document.updatestyles
- https://learn.microsoft.com/en-us/office/vba/api/word.document.windows
- https://learn.microsoft.com/en-us/office/vba/api/word.documents.add
- https://learn.microsoft.com/en-us/office/vba/api/word.documents.count
- https://learn.microsoft.com/en-us/office/vba/api/word.documents.item
- https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
- https://learn.microsoft.com/en-us/office/vba/api/word.find
- https://learn.microsoft.com/en-us/office/vba/api/word.find.clearformatting
- https://learn.microsoft.com/en-us/office/vba/api/word.find.execute
- https://learn.microsoft.com/en-us/office/vba/api/word.find.replacement
- https://learn.microsoft.com/en-us/office/vba/api/word.find.text
- https://learn.microsoft.com/en-us/office/vba/api/word.font
- https://learn.microsoft.com/en-us/office/vba/api/word.font.allcaps
- https://learn.microsoft.com/en-us/office/vba/api/word.font.bold
- https://learn.microsoft.com/en-us/office/vba/api/word.font.boldbi
- https://learn.microsoft.com/en-us/office/vba/api/word.font.borders
- https://learn.microsoft.com/en-us/office/vba/api/word.font.colorindex
- https://learn.microsoft.com/en-us/office/vba/api/word.font.colorindexbi
- https://learn.microsoft.com/en-us/office/vba/api/word.font.contextualalternates
- https://learn.microsoft.com/en-us/office/vba/api/word.font.diacriticcolor
- https://learn.microsoft.com/en-us/office/vba/api/word.font.disablecharacterspacegrid
- https://learn.microsoft.com/en-us/office/vba/api/word.font.doublestrikethrough
- https://learn.microsoft.com/en-us/office/vba/api/word.font.duplicate
- https://learn.microsoft.com/en-us/office/vba/api/word.font.emboss
- https://learn.microsoft.com/en-us/office/vba/api/word.font.emphasismark
- https://learn.microsoft.com/en-us/office/vba/api/word.font.engrave
- https://learn.microsoft.com/en-us/office/vba/api/word.font.fill
- https://learn.microsoft.com/en-us/office/vba/api/word.font.glow
- https://learn.microsoft.com/en-us/office/vba/api/word.font.grow
- https://learn.microsoft.com/en-us/office/vba/api/word.font.hidden
- https://learn.microsoft.com/en-us/office/vba/api/word.font.italic
- https://learn.microsoft.com/en-us/office/vba/api/word.font.italicbi
- https://learn.microsoft.com/en-us/office/vba/api/word.font.kerning
- https://learn.microsoft.com/en-us/office/vba/api/word.font.ligatures
- https://learn.microsoft.com/en-us/office/vba/api/word.font.line
- https://learn.microsoft.com/en-us/office/vba/api/word.font.name
- https://learn.microsoft.com/en-us/office/vba/api/word.font.nameascii
- https://learn.microsoft.com/en-us/office/vba/api/word.font.namebi
- https://learn.microsoft.com/en-us/office/vba/api/word.font.namefareast
- https://learn.microsoft.com/en-us/office/vba/api/word.font.nameother
- https://learn.microsoft.com/en-us/office/vba/api/word.font.numberform
- https://learn.microsoft.com/en-us/office/vba/api/word.font.numberspacing
- https://learn.microsoft.com/en-us/office/vba/api/word.font.outline
- https://learn.microsoft.com/en-us/office/vba/api/word.font.position
- https://learn.microsoft.com/en-us/office/vba/api/word.font.reflection
- https://learn.microsoft.com/en-us/office/vba/api/word.font.reset
- https://learn.microsoft.com/en-us/office/vba/api/word.font.scaling
- https://learn.microsoft.com/en-us/office/vba/api/word.font.setastemplatedefault
- https://learn.microsoft.com/en-us/office/vba/api/word.font.shading
- https://learn.microsoft.com/en-us/office/vba/api/word.font.shadow
- https://learn.microsoft.com/en-us/office/vba/api/word.font.shrink
- https://learn.microsoft.com/en-us/office/vba/api/word.font.size
- https://learn.microsoft.com/en-us/office/vba/api/word.font.sizebi
- https://learn.microsoft.com/en-us/office/vba/api/word.font.smallcaps
- https://learn.microsoft.com/en-us/office/vba/api/word.font.spacing
- https://learn.microsoft.com/en-us/office/vba/api/word.font.strikethrough
- https://learn.microsoft.com/en-us/office/vba/api/word.font.stylisticset
- https://learn.microsoft.com/en-us/office/vba/api/word.font.subscript
- https://learn.microsoft.com/en-us/office/vba/api/word.font.superscript
- https://learn.microsoft.com/en-us/office/vba/api/word.font.textcolor
- https://learn.microsoft.com/en-us/office/vba/api/word.font.textshadow
- https://learn.microsoft.com/en-us/office/vba/api/word.font.threed
- https://learn.microsoft.com/en-us/office/vba/api/word.font.underline
- https://learn.microsoft.com/en-us/office/vba/api/word.font.underlinecolor
- https://learn.microsoft.com/en-us/office/vba/api/word.list
- https://learn.microsoft.com/en-us/office/vba/api/word.list.applylisttemplate
- https://learn.microsoft.com/en-us/office/vba/api/word.list.applylisttemplatewithlevel
- https://learn.microsoft.com/en-us/office/vba/api/word.list.range
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applybulletdefault
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applylisttemplate
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applylisttemplatewithlevel
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applynumberdefault
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.applyoutlinenumberdefault
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.cancontinuepreviouslist
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.convertnumberstotext
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.countnumbereditems
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.list
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listindent
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listlevelnumber
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listoutdent
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listpicturebullet
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.liststring
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listtemplate
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listtype
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.listvalue
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.removenumbers
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.singlelist
- https://learn.microsoft.com/en-us/office/vba/api/word.listformat.singlelisttemplate
- https://learn.microsoft.com/en-us/office/vba/api/word.listgalleries
- https://learn.microsoft.com/en-us/office/vba/api/word.listgalleries.count
- https://learn.microsoft.com/en-us/office/vba/api/word.listgalleries.item
- https://learn.microsoft.com/en-us/office/vba/api/word.listgallery
- https://learn.microsoft.com/en-us/office/vba/api/word.listgallery.listtemplates
- https://learn.microsoft.com/en-us/office/vba/api/word.listgallery.modified
- https://learn.microsoft.com/en-us/office/vba/api/word.listgallery.reset
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.alignment
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.applypicturebullet
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.font
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.index
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.linkedstyle
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.numberformat
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.numberposition
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.numberstyle
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.picturebullet
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.resetonhigher
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.startat
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.tabposition
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.textposition
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevel.trailingcharacter
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevels
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevels.count
- https://learn.microsoft.com/en-us/office/vba/api/word.listlevels.item
- https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate
- https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate.convert
- https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate.listlevels
- https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate.name
- https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate.outlinenumbered
- https://learn.microsoft.com/en-us/office/vba/api/word.listtemplates
- https://learn.microsoft.com/en-us/office/vba/api/word.listtemplates.add
- https://learn.microsoft.com/en-us/office/vba/api/word.listtemplates.count
- https://learn.microsoft.com/en-us/office/vba/api/word.listtemplates.item
- https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.bottommargin
- https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.leftmargin
- https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.pageheight
- https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.pagewidth
- https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.rightmargin
- https://learn.microsoft.com/en-us/office/vba/api/word.pagesetup.topmargin
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.format
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.next
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.outlinelevel
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.range
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraph.style
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.addspacebetweenfareastandalpha
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.addspacebetweenfareastanddigit
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.alignment
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.autoadjustrightindent
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.baselinealignment
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.borders
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.characterunitfirstlineindent
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.characterunitleftindent
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.characterunitrightindent
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.closeup
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.collapsedbydefault
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.disablelineheightgrid
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.duplicate
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.fareastlinebreakcontrol
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.firstlineindent
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.halfwidthpunctuationontopofline
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.hangingpunctuation
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.hyphenation
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.indentcharwidth
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.indentfirstlinecharwidth
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.keeptogether
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.keepwithnext
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.leftindent
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.linespacing
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.linespacingrule
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.lineunitafter
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.lineunitbefore
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.mirrorindents
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.nolinenumber
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.openorcloseup
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.openup
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.outlinelevel
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.pagebreakbefore
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.readingorder
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.reset
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.rightindent
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.shading
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.space1
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.space15
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.space2
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spaceafter
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spaceafterauto
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spacebefore
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spacebeforeauto
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.style
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.tabhangingindent
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.tabindent
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.tabstops
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.textboxtightwrap
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.widowcontrol
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.wordwrap
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphs.addspacebetweenfareastandalpha
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphs.count
- https://learn.microsoft.com/en-us/office/vba/api/word.paragraphs.item
- https://learn.microsoft.com/en-us/office/vba/api/word.range.characterstyle
- https://learn.microsoft.com/en-us/office/vba/api/word.range.collapse
- https://learn.microsoft.com/en-us/office/vba/api/word.range.delete
- https://learn.microsoft.com/en-us/office/vba/api/word.range.end
- https://learn.microsoft.com/en-us/office/vba/api/word.range.find
- https://learn.microsoft.com/en-us/office/vba/api/word.range.font
- https://learn.microsoft.com/en-us/office/vba/api/word.range.highlightcolorindex
- https://learn.microsoft.com/en-us/office/vba/api/word.range.insertafter
- https://learn.microsoft.com/en-us/office/vba/api/word.range.insertbefore
- https://learn.microsoft.com/en-us/office/vba/api/word.range.listformat
- https://learn.microsoft.com/en-us/office/vba/api/word.range.listparagraphs
- https://learn.microsoft.com/en-us/office/vba/api/word.range.paragraphformat
- https://learn.microsoft.com/en-us/office/vba/api/word.range.paragraphs
- https://learn.microsoft.com/en-us/office/vba/api/word.range.paragraphstyle
- https://learn.microsoft.com/en-us/office/vba/api/word.range.start
- https://learn.microsoft.com/en-us/office/vba/api/word.range.style
- https://learn.microsoft.com/en-us/office/vba/api/word.range.text
- https://learn.microsoft.com/en-us/office/vba/api/word.replacement.clearformatting
- https://learn.microsoft.com/en-us/office/vba/api/word.saveas2
- https://learn.microsoft.com/en-us/office/vba/api/word.sections.count
- https://learn.microsoft.com/en-us/office/vba/api/word.shading.backgroundpatterncolor
- https://learn.microsoft.com/en-us/office/vba/api/word.shading.backgroundpatterncolorindex
- https://learn.microsoft.com/en-us/office/vba/api/word.shading.foregroundpatterncolor
- https://learn.microsoft.com/en-us/office/vba/api/word.shading.foregroundpatterncolorindex
- https://learn.microsoft.com/en-us/office/vba/api/word.shading.texture
- https://learn.microsoft.com/en-us/office/vba/api/word.style
- https://learn.microsoft.com/en-us/office/vba/api/word.style.automaticallyupdate
- https://learn.microsoft.com/en-us/office/vba/api/word.style.basestyle
- https://learn.microsoft.com/en-us/office/vba/api/word.style.borders
- https://learn.microsoft.com/en-us/office/vba/api/word.style.builtin
- https://learn.microsoft.com/en-us/office/vba/api/word.style.delete
- https://learn.microsoft.com/en-us/office/vba/api/word.style.description
- https://learn.microsoft.com/en-us/office/vba/api/word.style.font
- https://learn.microsoft.com/en-us/office/vba/api/word.style.frame
- https://learn.microsoft.com/en-us/office/vba/api/word.style.inuse
- https://learn.microsoft.com/en-us/office/vba/api/word.style.languageid
- https://learn.microsoft.com/en-us/office/vba/api/word.style.languageidfareast
- https://learn.microsoft.com/en-us/office/vba/api/word.style.linked
- https://learn.microsoft.com/en-us/office/vba/api/word.style.linkstyle
- https://learn.microsoft.com/en-us/office/vba/api/word.style.linktolisttemplate
- https://learn.microsoft.com/en-us/office/vba/api/word.style.listlevelnumber
- https://learn.microsoft.com/en-us/office/vba/api/word.style.listtemplate
- https://learn.microsoft.com/en-us/office/vba/api/word.style.locked
- https://learn.microsoft.com/en-us/office/vba/api/word.style.namelocal
- https://learn.microsoft.com/en-us/office/vba/api/word.style.nextparagraphstyle
- https://learn.microsoft.com/en-us/office/vba/api/word.style.noproofing
- https://learn.microsoft.com/en-us/office/vba/api/word.style.nospacebetweenparagraphsofsamestyle
- https://learn.microsoft.com/en-us/office/vba/api/word.style.paragraphformat
- https://learn.microsoft.com/en-us/office/vba/api/word.style.priority
- https://learn.microsoft.com/en-us/office/vba/api/word.style.quickstyle
- https://learn.microsoft.com/en-us/office/vba/api/word.style.shading
- https://learn.microsoft.com/en-us/office/vba/api/word.style.table
- https://learn.microsoft.com/en-us/office/vba/api/word.style.type
- https://learn.microsoft.com/en-us/office/vba/api/word.style.unhidewhenused
- https://learn.microsoft.com/en-us/office/vba/api/word.style.visibility
- https://learn.microsoft.com/en-us/office/vba/api/word.styles.add
- https://learn.microsoft.com/en-us/office/vba/api/word.styles.count
- https://learn.microsoft.com/en-us/office/vba/api/word.styles.item
- https://learn.microsoft.com/en-us/office/vba/api/word.tables.count
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.alignment
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.clear
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.customtab
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.leader
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.next
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.position
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstop.previous
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.add
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.after
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.before
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.clearall
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.count
- https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.item
- https://learn.microsoft.com/en-us/office/vba/api/word.undorecord
- https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.customrecordlevel
- https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.customrecordname
- https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.endcustomrecord
- https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.isrecordingcustomrecord
- https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.startcustomrecord
- https://learn.microsoft.com/en-us/office/vba/api/word.wdbaselinealignment
- https://learn.microsoft.com/en-us/office/vba/api/word.wdbuiltinstyle
- https://learn.microsoft.com/en-us/office/vba/api/word.wdcollapsedirection
- https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor
- https://learn.microsoft.com/en-us/office/vba/api/word.wdcolorindex
- https://learn.microsoft.com/en-us/office/vba/api/word.wdconstants
- https://learn.microsoft.com/en-us/office/vba/api/word.wdcontinue
- https://learn.microsoft.com/en-us/office/vba/api/word.wddefaultlistbehavior
- https://learn.microsoft.com/en-us/office/vba/api/word.wdemphasismark
- https://learn.microsoft.com/en-us/office/vba/api/word.wdexportcreatebookmarks
- https://learn.microsoft.com/en-us/office/vba/api/word.wdexportformat
- https://learn.microsoft.com/en-us/office/vba/api/word.wdexportitem
- https://learn.microsoft.com/en-us/office/vba/api/word.wdexportoptimizefor
- https://learn.microsoft.com/en-us/office/vba/api/word.wdexportrange
- https://learn.microsoft.com/en-us/office/vba/api/word.wdfindwrap
- https://learn.microsoft.com/en-us/office/vba/api/word.wdligatures
- https://learn.microsoft.com/en-us/office/vba/api/word.wdlinespacing
- https://learn.microsoft.com/en-us/office/vba/api/word.wdlistapplyto
- https://learn.microsoft.com/en-us/office/vba/api/word.wdlistgallerytype
- https://learn.microsoft.com/en-us/office/vba/api/word.wdlistlevelalignment
- https://learn.microsoft.com/en-us/office/vba/api/word.wdlistnumberstyle
- https://learn.microsoft.com/en-us/office/vba/api/word.wdlisttype
- https://learn.microsoft.com/en-us/office/vba/api/word.wdnumberform
- https://learn.microsoft.com/en-us/office/vba/api/word.wdnumberspacing
- https://learn.microsoft.com/en-us/office/vba/api/word.wdoriginalformat
- https://learn.microsoft.com/en-us/office/vba/api/word.wdoutlinelevel
- https://learn.microsoft.com/en-us/office/vba/api/word.wdparagraphalignment
- https://learn.microsoft.com/en-us/office/vba/api/word.wdreadingorder
- https://learn.microsoft.com/en-us/office/vba/api/word.wdreplace
- https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat
- https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveoptions
- https://learn.microsoft.com/en-us/office/vba/api/word.wdstatistic
- https://learn.microsoft.com/en-us/office/vba/api/word.wdstyletype
- https://learn.microsoft.com/en-us/office/vba/api/word.wdstylisticset
- https://learn.microsoft.com/en-us/office/vba/api/word.wdtabalignment
- https://learn.microsoft.com/en-us/office/vba/api/word.wdtableader
- https://learn.microsoft.com/en-us/office/vba/api/word.wdthemecolorindex
- https://learn.microsoft.com/en-us/office/vba/api/word.wdtrailingcharacter
- https://learn.microsoft.com/en-us/office/vba/api/word.wdunderline
- https://learn.microsoft.com/en-us/office/vba/api/word.wdunits
- https://learn.microsoft.com/en-us/office/vba/api/word.window.scrollintoview
- https://learn.microsoft.com/en-us/office/vba/api/word.windows.count
- https://learn.microsoft.com/en-us/office/vba/api/word.windows.item
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/boolean-data-type
- https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-tables
- https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-the-undorecord-object
- https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb238769(v=office.12)
