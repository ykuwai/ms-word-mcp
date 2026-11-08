# Equations in Word — verified reference for ms-word-mcp

Everything in this file comes from Microsoft's own documentation, downloaded while writing it. The sources are:

- the Office VBA reference for Word and its table of contents (https://learn.microsoft.com/en-us/office/vba/api/toc.json);
- Microsoft's Word interop and Open XML SDK reference pages on learn.microsoft.com;
- Office support articles on support.microsoft.com;
- Murray Sargent's "Math in Office" blog, in both its archived form (learn.microsoft.com/en-us/archive/blogs/murrays/...) and its current form (devblogs.microsoft.com/math-in-office/...). Sargent is the Microsoft engineer who designed Office math.
- Unicode Technical Note 28 (UnicodeMath, by the same author), which the Office support page names as the linear format Office uses.

For Word for Mac, the source is the Apple Event dictionary that ships with Word: `/Applications/Microsoft Word.app/Contents/Resources/Word.sdef`, Word 16.114.4 (CFBundleShortVersionString), parsed with Python's `xml.etree`. Rows cite it as `Word.sdef:<line>`.

Where a page does not say something, the row says **not stated**. Nothing here is filled from memory. A few rows are marked **inferred**; each of those says which two documented facts it combines.

**Method.** Each page was fetched with `curl -sL` from the URL cited, the HTTP status was checked (all 200), and the text was extracted from the page's `<main>` element. For the devblogs posts, the post body came from the blog's WordPress JSON endpoint (`/wp-json/wp/v2/posts?slug=…`). The search for a LaTeX/linear-format member covered three things:

- every `href` in the VBA toc.json (35,053 entries, 6,214 of them Word);
- the member lists of the Word interop `Options` and `OMath` pages;
- the whole Word.sdef.

Fetched on 2026-11-08.

**Conventions.**

- **R/W** means the page says Read/write.
- **R** means the page says Read-only.
- **not stated** means the page gives no access mode or no answer.
- In method rows, `Name?` marks an optional parameter.
- VBA `True` converts to -1 as a number (see word-object-model.md).

---

## 0. Answers in brief

1. **No documented way to choose LaTeX or UnicodeMath input through the object model.** Neither Windows COM nor Mac Apple Events has one.
   - `OMath.BuildUp` and `OMaths.BuildUp` take no arguments.
   - No `Options`, `OMathAutoCorrect`, `Application` or `Document` member mentions LaTeX or an input format. The toc.json has no href containing "latex", "linearformat", "inputformat", "unicodemath", "mathml" or "omml".
   - The interop `Options` page lists only `OMathAutoBuildUp` and `OMathCopyLF`.
   - Word.sdef has 0 occurrences of "latex" or "unicodemath".

   The LaTeX switch is documented only as a UI choice ("by selecting the format from the Equations tab"). Section 2 lists the documented alternatives.
2. **Which linear format `BuildUp` parses, and which one `Linearize` produces, is not stated.** The OM pages say only "professional format" and "linear format".
3. **Equation numbers: two documented routes, neither documented for the object model.**
   - UnicodeMath `#`: `E=mc^2#(30)` followed by Enter flushes "(30)" to the right margin. This works through an equation array.
   - Captions: Insert Caption with the Equation label, which is a `SEQ Equation` field.

   Whether `OMath.BuildUp` honours `#` when called from code is **not stated**. The SEQ field itself is fully documented (§3.3–3.5). So are cross-references to bookmarks, and the three-cell table and tab layouts (§3.7).
4. **Reading back.**
   - `Range.Text` is documented only as "the plain, unformatted text of the range".
   - `Linearize` followed by `Range.Text` gives the linear text, in a format that is **not stated**.
   - `Range.WordOpenXML` returns the range's Word Open XML, which includes the OMML (Windows only). Whether this is lossless is **not stated**.
5. **Mac.**
   - Word.sdef has equivalents for almost every OMath member: `create new equation`, `build up`, `linearize`, the three `convert to … text` commands, the `math object` class, the document math properties, and `math autocorrect`.
   - It has **no** equivalent of `Range.InsertXML`, `Range.WordOpenXML` or `Range.XML`, so there is no Apple Events path to OMML.
   - It has no equivalent of `OMath.Remove`, and no bookmark-creation command.

---

## Discrepancies and traps

1. **"Display" means two different things.** WdOMathType: `wdOMathDisplay` = 0 is described as "Professional format." and `wdOMathInline` = 1 as "Inline." (https://learn.microsoft.com/en-us/office/vba/api/word.wdomathtype). Yet `OMath.Type` is about layout: "whether an equation is displayed inline with the text around it or displayed on its own line" (https://learn.microsoft.com/en-us/office/vba/api/word.omath.type). The Mac dictionary describes the same two values as "professional/built up format" and "linear/built down style" (Word.sdef:6871). Professional vs. linear is what BuildUp/Linearize control; display vs. inline is OMath.Type. Do not mix them up when reading the descriptions.
2. **Mac document math properties are marked read-only but described as settable.** Every one of them has `access="r"` together with a description that begins "Gets or sets…" (Word.sdef:3774–3823): `math binary operator break`, `math default justification`, `math font name`, `math left margin`, `math right margin`, `math subtraction operator`, `math wrap`, `integrals uses subscripts`, `nary uses subscripts`, `use default math settings`, `use small fractions`. Whether they can be set is **not stated**. On Windows the same properties are R/W.
3. **Mac `use default math settings` has no VBA counterpart.** The VBA toc has no Document property for it (Word.sdef:3822).
4. **WdOMathFunctionType descriptions.** VBA describes `wdOMathFunctionBar` (2) as "Equation fraction bar." and `wdOMathFunctionRad` (16) as "Equation base expression." (https://learn.microsoft.com/en-us/office/vba/api/word.wdomathfunctiontype). The Mac names for the same values are "math function bar type" and "math radical type". OMathBar is the over/underbar object (https://learn.microsoft.com/en-us/office/vba/api/word.omathbar), so the VBA descriptions look wrong; the names are right. The Mac enum also has a value VBA lacks: `math literal text type` = 0x02cf0016, i.e. 22 (Word.sdef:6824).
5. **WdOMathFracType value 1.** VBA calls it `wdOMathFracNoBar` ("No fraction bar."). Mac calls it `math fraction type stack` (Word.sdef:6858).
6. **WdOMathBreakBin names describe where the break falls relative to the operator.** `wdOMathBreakBinAfter` = 1 "Places the operator before a line break, at the end of the line." `wdOMathBreakBinBefore` = 0 "Places the operator after a line break, at the beginning of the following line." (https://learn.microsoft.com/en-us/office/vba/api/word.wdomathbreakbin). "After" means a break after the operator, not an operator after the break. The default is a break before the operator: "By default, a line break occurs before the binary operator" (https://devblogs.microsoft.com/math-in-office/default-math-properties/). Mac uses the same order: `break before` = 0, `break after` = 1, `repeat` = 2 (Word.sdef:6886). WdOMathBreakSub agrees on both platforms: PlusMinus = 1 puts plus at the line end and minus at the next line's start, and MinusPlus = 2 does the reverse.
7. **`OMath.ArgSize`, not `ArgumentSize`.** The VBA property is `ArgSize`: "the script size of an argument, for example, text, script, or script-script" (https://learn.microsoft.com/en-us/office/vba/api/word.omath.argsize). The page does not give the numeric values: **not stated**. Mac calls it `script size`.
8. **`GetCrossReferenceItems` documents only WdReferenceType.** `wdRefTypeNumberedItem` = 0 means numbered *paragraphs*, not equations. Whether `GetCrossReferenceItems` accepts a caption label such as `wdCaptionEquation` or "Equation" is **not stated**. Its page says "Can be any WdReferenceType constant" (https://learn.microsoft.com/en-us/office/vba/api/word.document.getcrossreferenceitems). In contrast, `InsertCrossReference` accepts "any WdReferenceType or WdCaptionLabelID constant or a user defined caption label" (https://learn.microsoft.com/en-us/office/vba/api/word.range.insertcrossreference). The GetCrossReferenceItems Remarks also name a `ReferenceWhich` argument of InsertCrossReference that does not exist; the parameter is `ReferenceItem`.
9. **`Field.Update` example.** The page's own example uses `ActiveDocument.Fields(0)` and compares the result to `1`, describing it as "1 (True)" (https://learn.microsoft.com/en-us/office/vba/api/word.field.update). Comparing with 1 conflicts with VBA True being -1, so do not copy that comparison. `Fields.Update` returns 0 on success. Its page also warns: "When using this with Word 2016 and possibly with others, the number may be incorrect." (https://learn.microsoft.com/en-us/office/vba/api/word.fields.update).
10. **Mac caption and cross-reference commands are narrower than VBA's.**
    - `insert caption` takes only a `WdCaptionLabelID`, with no custom label, no `TitleAutoText` and no `ExcludeLabel` (Word.sdef:2516).
    - `insert cross reference` takes a `WdReferenceType` only, so you cannot pass a caption label, and it has no `SeparateNumbers` or `SeparatorString` (Word.sdef:2530).
    - `get cross reference items` takes `WdReferenceType` (Word.sdef:2383).
11. **Mac `create new equation`.** Its description says it "returns a text range object", but the command has no `<result>` element. Its `in range` parameter is typed `selection object`, not `text range` (Word.sdef:2214).
12. **Localization.** The UI label "Equation" is localized; Japanese Word shows 数式, for example. Use `wdCaptionEquation` (-3) rather than the string "Equation". Two things are **not stated**: which SEQ identifier `InsertCaption` writes in a localized Word, and whether `STYLEREF` accepts English built-in style names there.

---

## 1. The OMath object model (Windows, COM / VBA)

### 1.1 Entry points

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Range.OMaths | OMaths | R | "Returns an OMaths collection that represents the OMath objects within the specified range." | https://learn.microsoft.com/en-us/office/vba/api/word.range.omaths |
| Selection.OMaths | OMaths | R | Same, for the current selection. | https://learn.microsoft.com/en-us/office/vba/api/word.selection.omaths |
| Document.OMaths | OMaths | R | The page says "within the specified range" (copied text). | https://learn.microsoft.com/en-us/office/vba/api/word.document.omaths |

### 1.2 OMaths collection

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| OMaths.Add(Range) | Range (Required, Range) → **Range** | method | "Creates an equation, from the text equation contained within the specified range, and returns a Range object that contains the new equation." The page's example: set `objRange.Text = "Celsius = (5/9)(Fahrenheit ? 32)"`, then `Selection.OMaths.Add(objRange)`, then `objRange.OMaths(1).BuildUp`. It returns a Range, not an OMath. Get the OMath with `returned.OMaths(1)`. Which linear syntax "text equation" means is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.omaths.add |
| OMaths.BuildUp | → Nothing | method | "Converts all equations in the collection to professional format." No parameters. | https://learn.microsoft.com/en-us/office/vba/api/word.omaths.buildup |
| OMaths.Linearize | → Nothing | method | "Converts all equations in the collection to linear format." No parameters. | https://learn.microsoft.com/en-us/office/vba/api/word.omaths.linearize |
| OMaths.Count | Long | R | | https://learn.microsoft.com/en-us/office/vba/api/word.omaths.count |
| OMaths.Item(Index) | Index: Long → OMath | method | "ordinal position of the object within the collection". | https://learn.microsoft.com/en-us/office/vba/api/word.omaths.item |

### 1.3 OMath (one equation)

The object page says it "Represents an equation. OMath objects are members of the OMaths collection." It lists these methods: BuildUp, ConvertToLiteralText, ConvertToMathText, ConvertToNormalText, Linearize, Remove. Source: https://learn.microsoft.com/en-us/office/vba/api/word.omath.

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| BuildUp | → Nothing | method | "Converts an equation to professional format." No parameters, so there is no input-format argument. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.buildup |
| Linearize | → Nothing | method | "Converts an equation to linear format." No parameters. Which linear format (UnicodeMath or LaTeX) it produces is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.linearize |
| ConvertToMathText | — | method | "Converts an equation to math text." Nothing more is stated. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.converttomathtext |
| ConvertToNormalText | — | method | "Converts an equation to normal text." Nothing more is stated. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.converttonormaltext |
| ConvertToLiteralText | — | method | "Converts an equation to literal text." Nothing more is stated. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.converttoliteraltext |
| Remove | — | method | "Removes an equation from the collection of equations in a document, range, or selection." Whether the text stays as plain text is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.remove |
| Range | Range | R | "the portion of a document that is contained in the specified object". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.range |
| Type | WdOMathType | R/W | "whether an equation is displayed inline with the text around it or displayed on its own line". The example sets `objEq.Type = wdOMathInline`. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.type |
| Justification | WdOMathJc | R/W | "the justification for an equation". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.justification |
| AlignPoint | Long | R/W | "the character position of the alignment point in the equation". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.alignpoint |
| ArgSize | Long | R/W | "the script size of an argument, for example, text, script, or script-script". The values are **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.argsize |
| ArgIndex | Long | R | "argument index of this component relative to the containing math object". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.argindex |
| NestingLevel | Long | R | "the nesting level for an OMath object". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.nestinglevel |
| Functions | OMathFunctions | R | "the functions contained within an equation". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.functions |
| Breaks | OMathBreaks | R | "the line breaks in an equation". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.breaks |
| Parent | Object | not stated | "the parent object of the specified OMath object". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.parent |
| ParentOMath | OMath | R | "the parent element". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.parentomath |
| ParentArg | OMath | R | "the parent, or containing, argument". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.parentarg |
| ParentFunction | OMathFunction | R | "the parent, or containing, function". | https://learn.microsoft.com/en-us/office/vba/api/word.omath.parentfunction |
| ParentCol / ParentRow | OMathMatCol / OMathMatRow | R | The parent column or row in a matrix. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.parentcol, https://learn.microsoft.com/en-us/office/vba/api/word.omath.parentrow |

### 1.4 OMathFunctions / OMathFunction (structures inside an equation)

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| OMathFunctions.Add(Range, Type, NumArgs?, NumCols?) | Range: Range (req); Type: WdOMathFunctionType (req); NumArgs, NumCols: Variant → **OMathFunction** | method | "Inserts a new structure, such as a fraction, into an equation at the specified position". | https://learn.microsoft.com/en-us/office/vba/api/word.omathfunctions.add |
| OMathFunctions.Item(Index) / Count | Long → OMathFunction / Long | method / R | | https://learn.microsoft.com/en-us/office/vba/api/word.omathfunctions.item, https://learn.microsoft.com/en-us/office/vba/api/word.omathfunctions.count |
| OMathFunction | object | — | "a mathematical function or structure that Microsoft Word supports, such as fractions, integrals, sums, and radicals". | https://learn.microsoft.com/en-us/office/vba/api/word.omathfunction |
| OMathFunction.Type | WdOMathFunctionType | R | | https://learn.microsoft.com/en-us/office/vba/api/word.omathfunction.type |
| OMathFunction.Range | Range | R | | https://learn.microsoft.com/en-us/office/vba/api/word.omathfunction.range |
| OMathFunction.OMath | OMath | R | "the equation". | https://learn.microsoft.com/en-us/office/vba/api/word.omathfunction.omath |
| OMathFunction.Args | OMathArgs | R | "the arguments for an equation". | https://learn.microsoft.com/en-us/office/vba/api/word.omathfunction.args |
| OMathFunction.Remove | → OMathFunction | method | "Removes a function from the collection of functions contained within an equation". | https://learn.microsoft.com/en-us/office/vba/api/word.omathfunction.remove |
| OMathFunction typed accessors | Acc, Bar, BorderBox, Box, Delim, EqArray, Frac, Func, GroupChar, LimLow, LimUpp, Mat, Nary, Phantom, Rad, ScrPre, ScrSub, ScrSubSup, ScrSup | R | Each returns the matching object: OMathAcc, OMathBar, OMathBorderBox, OMathBox, OMathDelim, OMathEqArray, OMathFrac, OMathFunc, OMathGroupChar, OMathLimLow, OMathLimUpp, OMathMat, OMathNary, OMathPhantom, OMathRad, OMathScrPre, OMathScrSub, OMathScrSubSup, OMathScrSup. The list comes from the toc (word.omathfunction.*). Examples of per-object members: OMathFrac.Type is a WdOMathFracType, R/W; OMathEqArray is "a mathematical equation array object". | toc.json; https://learn.microsoft.com/en-us/office/vba/api/word.omathfrac.type; https://learn.microsoft.com/en-us/office/vba/api/word.omatheqarray |

### 1.5 OMathBreaks / OMathBreak

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| OMathBreaks.Add(Range) | Range (req) → OMathBreak | method | "Inserts a break into an equation". | https://learn.microsoft.com/en-us/office/vba/api/word.omathbreaks.add |
| OMathBreak.AlignAt | Long | R/W | "the operator in one line, to which Microsoft Word aligns consecutive lines in an equation". | https://learn.microsoft.com/en-us/office/vba/api/word.omathbreak.alignat |
| OMathBreak.Range | Range | R | | https://learn.microsoft.com/en-us/office/vba/api/word.omathbreak.range |
| OMathBreak.Delete | — | method | "Deletes the specified equation break." | https://learn.microsoft.com/en-us/office/vba/api/word.omathbreak.delete |

### 1.6 Document-level math settings

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Document.OMathBreakBin | WdOMathBreakBin | R/W | Where binary operators go when an equation breaks across lines. Remarks: when it is `wdOMathBreakBinRepeat`, "use the OMathBreakSub property". | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathbreakbin |
| Document.OMathBreakSub | WdOMathBreakSub | R/W | "This property is used only when the OMathBreakBin property is set to wdOMathBreakBinRepeat." | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathbreaksub |
| Document.OMathFontName | String | R/W | "the name of the font used in a document to display equations". | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathfontname |
| Document.OMathIntSubSupLim | Boolean | R/W | "the default location of limits for integrals". Which of True/False means sub/sup position is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathintsubsuplim |
| Document.OMathNarySupSubLim | Boolean | R/W | "the default location of limits for n-ary objects other than integrals". True/False meaning **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathnarysupsublim |
| Document.OMathJc | WdOMathJc | R/W | "the default justification—left, right, centered, or centered as a group—of a group of equations". | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathjc |
| Document.OMathLeftMargin | Single | R/W | "the left margin for equations". The unit is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathleftmargin |
| Document.OMathRightMargin | Single | R/W | "the right margin for equations". The unit is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathrightmargin |
| Document.OMathSmallFrac | Boolean | R/W | "whether to use small fractions in equations". | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathsmallfrac |
| Document.OMathWrap | Single | R/W | "A value of -1 specifies that the continuation of the equation is right-aligned. All positive values indicate the indent from the left margin." The unit is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.document.omathwrap |

These are the document defaults shown in the Equation Options dialog. Sargent describes them with their RTF and OMML (`<mathPr>`) equivalents (https://devblogs.microsoft.com/math-in-office/default-math-properties/). The ISO 29500 text quoted on the Open XML SDK page says that if `m:jc` is omitted from a math paragraph, "the math paragraph is Centered as Group" (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.justification?view=openxml-3.0.1).

### 1.7 Options (application-wide)

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Options.OMathAutoBuildUp | Boolean | R/W | "whether Microsoft Word automatically converts equations to professional format." | https://learn.microsoft.com/en-us/office/vba/api/word.options.omathautobuildup |
| Options.OMathCopyLF | Boolean | R/W | "how equations are represented in plain text. True indicates equations are represented in Linear Format. False indicates equations are represented in MathML." **Inferred:** this is the Equation Options dialog checkbox "Copy MathML … / Copy Linear Format to the clipboard as plain text". The VBA page does not name the checkbox, and the blog that names it ("stored in the system registry rather than in the document", default-math-properties post) does not name the property. It does not choose an input format. | https://learn.microsoft.com/en-us/office/vba/api/word.options.omathcopylf |

The interop `Options` page lists only these two `OMath*` members, and no member containing "LaTeX" (https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.options?view=word-pia).

### 1.8 Application.OMathAutoCorrect

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Application.OMathAutoCorrect | OMathAutoCorrect | R | "the auto correct entries for equations". | https://learn.microsoft.com/en-us/office/vba/api/word.application.omathautocorrect |
| OMathAutoCorrect.Entries | OMathAutoCorrectEntries | R | "the collection of equation autocorrect entries". | https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrect.entries |
| OMathAutoCorrect.Functions | OMathRecognizedFunctions | R | "the recognized functions that are automatically corrected". | https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrect.functions |
| OMathAutoCorrect.ReplaceText | Boolean | R/W | "whether Microsoft Word automatically replaces strings in equations with the corresponding math AutoCorrect definitions". | https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrect.replacetext |
| OMathAutoCorrect.UseOutsideOMath | Boolean | R/W | "whether Microsoft Word uses math autocorrect rules outside equations". | https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrect.useoutsideomath |
| OMathAutoCorrectEntries.Add(Name, Value) | String, String → OMathAutoCorrectEntry | method | "Creates an equation auto correct entry". | https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrectentries.add |
| OMathAutoCorrectEntries.Item(Index) | String or Integer → entry | method | "name or ordinal position". | https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrectentries.item |
| OMathAutoCorrectEntry.Name / .Value | String / String | R/W / R/W | The name (for example `\alpha`) and the replacement. | https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrectentry.name, https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrectentry.value |
| OMathAutoCorrectEntry.Index / .Delete | Long / method | R / — | | https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrectentry.index, https://learn.microsoft.com/en-us/office/vba/api/word.omathautocorrectentry.delete |
| OMathRecognizedFunctions.Add(Name) | String → OMathRecognizedFunction | method | "Creates a new recognized function". | https://learn.microsoft.com/en-us/office/vba/api/word.omathrecognizedfunctions.add |
| OMathRecognizedFunction.Name / .Delete | String (R) / method | R / — | | https://learn.microsoft.com/en-us/office/vba/api/word.omathrecognizedfunction.name, https://learn.microsoft.com/en-us/office/vba/api/word.omathrecognizedfunction.delete |

Whether math autocorrect (for example `\alpha` → α) is applied to text passed through `OMaths.Add` and `BuildUp` from code is **not stated**. The support page describes autocorrect only for typing. The entry table can be read through `Entries` and applied by the caller.

### 1.9 Enumerations

| Enum | Values (name = value, description) | Source URL |
|---|---|---|
| WdOMathType | wdOMathDisplay = 0 "Professional format."; wdOMathInline = 1 "Inline." | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathtype |
| WdOMathJc | wdOMathJcCenterGroup = 1; wdOMathJcCenter = 2; wdOMathJcLeft = 3; wdOMathJcRight = 4; wdOMathJcInline = 7 | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathjc |
| WdOMathBreakBin | wdOMathBreakBinBefore = 0; wdOMathBreakBinAfter = 1; wdOMathBreakBinRepeat = 2 (see Discrepancy 6; break before the operator is the default) | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathbreakbin |
| WdOMathBreakSub | wdOMathBreakSubMinusMinus = 0 "Default."; wdOMathBreakSubPlusMinus = 1; wdOMathBreakSubMinusPlus = 2 | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathbreaksub |
| WdOMathFunctionType | Acc = 1; Bar = 2; Box = 3; BorderBox = 4; Delim = 5; EqArray = 6; Frac = 7; Func = 8; GroupChar = 9; LimLow = 10; LimUpp = 11; Mat = 12; Nary = 13; Phantom = 14; ScrPre = 15; Rad = 16; ScrSub = 17; ScrSubSup = 18; ScrSup = 19; Text = 20; NormalText = 21. Each name is prefixed `wdOMathFunction`. | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathfunctiontype |
| WdOMathFracType | wdOMathFracBar = 0; wdOMathFracNoBar = 1; wdOMathFracSkw = 2; wdOMathFracLin = 3 | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathfractype |
| WdOMathHorizAlignType | Center = 0; Left = 1; Right = 2 | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathhorizaligntype |
| WdOMathVertAlignType | Center = 0; Top = 1; Bottom = 2 | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathvertaligntype |
| WdOMathShapeType | wdOMathShapeCentered = 0; wdOMathShapeMatch = 1 | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathshapetype |
| WdOMathSpacingRule | Single = 0; 1pt5 = 1; Double = 2; Exactly = 3; Multiple = 4 | https://learn.microsoft.com/en-us/office/vba/api/word.wdomathspacingrule |

---

## 2. LaTeX vs UnicodeMath input

### 2.1 What Microsoft documents (UI only)

| Fact | Quote | Source URL |
|---|---|---|
| Word has two linear formats | "There are two linear formats for math that Word supports: … Unicode math … LaTeX math". You choose one "by selecting the format from the Equations tab". "All the other Office applications support only UnicodeMath linear format." | https://support.microsoft.com/en-us/office/linear-format-equations-using-unicodemath-and-latex-in-word-2e00618d-b1fd-49d8-8cb4-8d17f25754f8 |
| Hot keys | Alt+= inserts an equation. Ctrl+= converts to Professional. Ctrl+Shift+= converts back to linear. | same |
| LaTeX limits | `\begin{}` and `\end{}` are not supported; use `\matrix{a & b \\ c & d}`. Listed as unsupported: `\eqarray`, `\Middle`, `\ldiv`, `\dsmash`. | same |
| Linear vs Professional in Word | "The linear option will display the equation in either UnicodeMath format, or LaTeX format, which can be set in the Conversions chunk." | https://support.microsoft.com/en-us/office/write-an-equation-or-formula-1d01cabc-ceb1-458d-bc70-7f9737722702 |
| LaTeX mode needs an explicit build-up | "In Word's LaTeX mode, you must request build up. Enter Ctrl+= to build up … and Shift+Ctrl+= to build the math zone down into the current linear format (UnicodeMath or LaTeX)." | https://devblogs.microsoft.com/math-in-office/officemath/ |
| Autobuildup is disabled in LaTeX mode | "Formula autobuildup is disabled for LaTeX in Word". The LaTeX option "supports all TeX control words appearing in Appendix B of the UnicodeMath spec". "\displaymath is implied if the math zone fills the hard/soft paragraph". | https://learn.microsoft.com/en-us/archive/blogs/murrays/latex-math-in-office |
| Switching mode in PowerPoint and OneNote | Define a math autocorrect entry `\TeX` → Ⓣ (U+24C9); typing `\TeX<space>` in a math zone switches to LaTeX. `\LF` → Ⓛ (U+24C1) switches back. This is documented for PowerPoint and OneNote only. Whether inserting Ⓣ through the object model switches Word's mode is **not stated**. | same |
| `/` in LaTeX mode | A `/` entered in LaTeX mode is marked "nobuildup", so it will not build up if you later switch to UnicodeMath. | https://devblogs.microsoft.com/math-in-office/switching-from-latex-to-unicodemath-input-mode/ |

### 2.2 Is there an object-model switch? No documented one.

Places searched, with what each showed:

| Place searched | Result | Source |
|---|---|---|
| `OMath.BuildUp`, `OMaths.BuildUp` signatures | No parameters. | https://learn.microsoft.com/en-us/office/vba/api/word.omath.buildup, https://learn.microsoft.com/en-us/office/vba/api/word.omaths.buildup |
| VBA toc.json, all 35,053 hrefs | No href contains latex, linearformat, inputformat, unicodemath, mathml or omml. All `OMath*` members are Word ones and are listed in §1. | https://learn.microsoft.com/en-us/office/vba/api/toc.json |
| Interop `Options` and `OMath` member lists | `OMathAutoBuildUp`, `OMathCopyLF`; `BuildUp`, `Linearize`, etc. Nothing about LaTeX. | https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.options?view=word-pia, https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.omath?view=word-pia |
| `Application.OMathAutoCorrect` | Entries, Functions, ReplaceText, UseOutsideOMath only. | §1.8 |
| Word.sdef (Mac) | 0 hits for "latex" or "unicodemath". | Word.sdef |
| Sargent's blog | The only programmatic LaTeX switch described is in **RichEdit's** TOM API, not Word's. `tomDocMathBuild` takes the flags `tomMathAutoCorrect`, `tomTeX` and `tomMathAlphabetics`: "The flag tomTeX uses [La]TeX build-up/down rules instead of UnicodeMath rules … These values are used by math UI but are not persisted in file formats." `ITextRange2::SetText2/GetText2` takes `tomConvertTeX` (0x00200000), `tomConvertUnicodeMath` (0x00040000), `tomConvertOMML` (0x00080000) and `tomConvertMathML` (0x00010000). These "work in the current Microsoft Office RichEdit (riched20.dll in an Office subdirectory)". The Win32 SetText2 page does not list the convert flags. | https://devblogs.microsoft.com/math-in-office/default-math-properties/, https://devblogs.microsoft.com/math-in-office/setting-and-getting-text-in-various-formats/, https://learn.microsoft.com/en-us/windows/win32/api/tom/nf-tom-itextrange2-settext2 |

**Conclusion:** Microsoft documents no Word object-model member, on Windows or Mac, that selects LaTeX or UnicodeMath for `BuildUp`, `Linearize` or `OMaths.Add`. Whether `BuildUp` called from code follows the Equations-tab setting the user last chose is **not stated**.

### 2.3 Documented building blocks for LaTeX input

Each step below is documented. Doing the LaTeX conversion yourself is the caller's job, not Word's.

| Route | Documented pieces | Not stated |
|---|---|---|
| A. LaTeX → UnicodeMath (own converter) → `OMaths.Add` + `BuildUp` | UnicodeMath syntax is in Unicode Technical Note 28; the Office support page says "Microsoft Office uses the linear format described in Unicode Technical Note 28". Appendix B maps TeX control words to characters. The math autocorrect table can be read via `OMathAutoCorrect.Entries` (§1.8). | That BuildUp from code uses UnicodeMath rules when the user's UI is in LaTeX mode. |
| A′. LaTeX `\frac`/`\binom` inside UnicodeMath mode | "You can enter the common LaTeX expressions \frac{a}{b} and \binom{n}{m} in UnicodeMath input mode provided you have added math autocorrect entries to convert \frac to ⍁ (U+2341) and \binom to ⒝ (U+249D)". `OMathAutoCorrectEntries.Add(Name, Value)` is the documented way to add entries (§1.8). Source: https://devblogs.microsoft.com/math-in-office/some-unicodemath-enhancements/ | Whether autocorrect entries apply to text passed through `OMaths.Add`/`BuildUp` from code (as opposed to typing). |
| B. LaTeX → MathML (own converter) → `MML2OMML.XSL` → OMML → Flat OPC → `Range.InsertXML` | `InsertXML` takes an optional XSLT `Transform` (§2.4). Office ships `mml2omml.xsl` (§2.5). The Flat OPC + InsertXML pattern is shown in Microsoft's 2010 article (§2.4). | Whether `Transform` is a path or XSLT text. Whether InsertXML accepts a bare `m:oMath` fragment. |
| C. LaTeX → OMML (own converter) → Flat OPC → `Range.InsertXML` | OMML element names (`m:oMathPara`, `m:oMath`, `m:eqArr`, `m:jc`) are in ISO/IEC 29500, quoted on the Open XML SDK pages (§2.4). | That InsertXML preserves display vs inline exactly as given. |
| D. LaTeX → OMML through Office RichEdit (Windows only) | `ITextRange2::SetText2(tomConvertTeX, …)`, then `GetText2(tomConvertOMML, …)`, in Office's riched20.dll. | Documented only in Sargent's blog, not on Learn. This is not part of Word's object model. |

### 2.4 XML insertion and extraction

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Range.InsertXML(XML, Transform?) | XML: String (req); Transform: Variant → Nothing | method | "Inserts the specified XML into the document at the specified range, replacing any text contained within the range." XML: "This can be any valid custom XML." Transform: "Specifies the XML Transformation (XSLT) used to transform the XML. If omitted, the XML is inserted as custom XML without applying a transform." The page's example uses custom (non-Word) XML. Whether Transform is a file path or XSLT text is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.range.insertxml |
| Selection.InsertXML(XML, Transform?) | same | method | "at the cursor, replacing any selected text". | https://learn.microsoft.com/en-us/office/vba/api/word.selection.insertxml |
| Range.WordOpenXML | String | R | "the XML contained within the range in the Microsoft Word Open XML format … returns only the XML in the document that is needed to represent the specified range." | https://learn.microsoft.com/en-us/office/vba/api/word.range.wordopenxml |
| Document.WordOpenXML | String | R | "the flat XML format for the Word Open XML contents of the document". | https://learn.microsoft.com/en-us/office/vba/api/word.document.wordopenxml |
| Range.XML(DataOnly?) | DataOnly: Boolean, default False → String | method/property | "the XML text in the specified object". "True returns the text of the XML without the Word XML markup." Which schema it uses is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.range.xml |
| Flat OPC round trip | — | — | "We use the WordOpenXml property of the Range object to return a Flat OPC document", then `range.InsertXML(openxml, ref missing)`. "To insert content in a Word document, you must prepare an XML document in the flat OPC format." (Word 2007/2010 article.) | https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff191178(v=office.14) |
| Minimal Flat OPC package | — | — | A `pkg:package` with a `/_rels/.rels` part and a `/word/document.xml` part, `w:document/w:body/w:p`. This is documented for Office.js `setSelectedDataAsync` coercion (a different API), and is useful as a template. | https://learn.microsoft.com/en-us/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml |
| OMML `m:oMathPara` | — | — | "a math paragraph, or display math zone, that contains one or more oMath elements that are in display mode" (ISO/IEC 29500 §22.1.2.78). | https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.paragraph?view=openxml-3.0.1 |
| OMML `m:oMath` | — | — | "When used independently (not inside an oMathPara) with non-mathematical text preceding and/or following it, an independent oMath is interpreted as an inline math zone." Parent elements include `p` (§22.1.2.77). | https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.officemath?view=openxml-3.0.1 |
| OMML `m:eqArr`, `m:oMathParaPr`, `m:jc` | — | — | Array object (§22.1.2.34); math paragraph properties; justification (default centerGroup). | https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.equationarray?view=openxml-3.0.1 and siblings |

Namespace for `m:`: `http://schemas.openxmlformats.org/officeDocument/2006/math`. This appears in the add-in article's `w:document` declaration and in the Mac `mathml2omml.xsl`.

### 2.5 Where MML2OMML.XSL / OMML2MML.XSL live

| Platform | What is documented | Source |
|---|---|---|
| Windows | "These XSLT files are responsible for converting between Word's native math format OMML and MathML 2.0 … (omml2mml.xsl and mml2omml.xsl) … copy them to C:\Program Files\Microsoft Office\Office12 subdirectory, or **wherever winword.exe is**." This is from 2007 (Office 2007). A Microsoft page giving the path for current Office (for example `root\Office16`) was **not found**; that path appears only in third-party pages. Rule documented by Microsoft: the directory of winword.exe. | https://learn.microsoft.com/en-us/archive/blogs/murrays/improved-mathml-support-in-word-2007 |
| Windows (use) | "These XSLTs are used, for example, by Word for MathML clipboard support." This statement is quoted in a reader comment on an archived Sargent post, not in the post body. | https://learn.microsoft.com/en-us/archive/blogs/murrays/mathml-and-ecma-math-omml |
| Mac | **Not stated** in any Microsoft page found. Observed in the installed bundle (Word 16.114.4): `/Applications/Microsoft Word.app/Contents/Resources/mathml2omml.xsl` (189,705 bytes) and `…/omml2mathml.xsl` (113,277 bytes). The file names differ from the Windows ones. `mathml2omml.xsl` declares `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`. | local file listing |

---

## 3. Equation numbering

### 3.1 UnicodeMath `#` (equation-array number, flush right)

| Fact | Quote / detail | Source URL |
|---|---|---|
| Syntax | "To enter an equation number using the linear format …, type the equation followed by a # (U+0023) followed by the desired equation number text and hit Enter. For example, E=mc^2#(30) ⏎". Available in Word 2016 and later. | https://learn.microsoft.com/en-us/archive/blogs/murrays/equation-numbering-in-office-2016 |
| Mechanism | "Internally this layout is created with an equation array in which the # character acts as a marker … to flush what follows the # to the right margin." | same |
| Spec | UnicodeMath 3.2 §3.21: "To represent an equation number flushed right of the equation in UnicodeMath, enter the equation followed by a # (U+0023) followed by the desired equation number text. For example █(E=mc^2#(30)) or more simply just E=mc^2#(30)". Appendix B: `\eqno` → `#` U+0023. | https://www.unicode.org/notes/tn28/UTN28-PlainTextMath-v3.2.pdf |
| LaTeX mode | **Inferred:** LaTeX mode supports all Appendix B control words (2017 post), and Appendix B maps `\eqno` to `#` (UTN28 v3.2, 2023). So `\eqno` *may* produce the same marker. No page states this for Word. | the two sources above |
| Automatic numbering | "The book [Creating Research and Scientific Documents using Microsoft Word] approach can be updated to use the equation array # option instead of the flush-right tab." The number text after `#` can be the Equation caption / SEQ field; you copy that number to the next equations. "Word automatically numbers all such entries sequentially." Ctrl+A, F9 updates. | equation-numbering-in-office-2016; https://learn.microsoft.com/en-us/archive/blogs/murrays/more-on-equation-numbering |
| Status | "The math paragraph has elegant equation-number support, but it hasn't been exposed beyond prototyping." The `#` equation array is the shipped approach. | https://devblogs.microsoft.com/math-in-office/officemath/ |
| Object model | Not stated: whether `OMaths.Add` followed by `OMath.BuildUp` on text such as `E=mc^2#(1)` produces the flush-right layout. Not stated: how to put a field inside the number part of a built-up equation from code. | — |
| Mac | Not stated by Microsoft; only reader comments on the 2016 post discuss it. | — |

### 3.2 Caption approach (Insert Caption, Equation label, SEQ field)

Sargent's procedure (UI): References › Insert Caption › Equation label › check "Exclude label from caption" › OK, then type "(" and ")" around the number. To refer to the number, bookmark it (a name starting with "eq"), then insert Cross-reference › Bookmark › "Bookmark text". Update with Ctrl+A, F9. For chapter numbers: Insert Caption › Numbering… › "Include chapter number". Source: https://learn.microsoft.com/en-us/archive/blogs/murrays/equation-numbering-in-office-2016. Office support says the Caption command is "the easiest way to insert Seq fields" (Seq field page, below).

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Range.InsertCaption(Label, Title?, TitleAutoText?, Position?, ExcludeLabel?) | Label: String or WdCaptionLabelID; Position: WdCaptionPosition; ExcludeLabel: "True does not include the text label" | method | "Inserts a caption immediately preceding or following the specified range." "If the label has not yet been defined, an error occurs." Whether it creates a separate paragraph (and so cannot give a same-line "(1)") is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.range.insertcaption |
| CaptionLabels.Add(Name) | String → CaptionLabel | method | "a custom caption label". | https://learn.microsoft.com/en-us/office/vba/api/word.captionlabels.add |
| Application.CaptionLabels | CaptionLabels | R | Index by name or `WdCaptionLabelID`, e.g. `CaptionLabels(wdCaptionTable)`. | https://learn.microsoft.com/en-us/office/vba/api/word.application.captionlabels, https://learn.microsoft.com/en-us/office/vba/api/word.captionlabel |
| CaptionLabel.IncludeChapterNumber | Boolean | R/W | The example sets it with `.ChapterStyleLevel = 2`; "The document should already contain a Heading 2 style with numbering." | https://learn.microsoft.com/en-us/office/vba/api/word.captionlabel.includechapternumber |
| CaptionLabel.ChapterStyleLevel | Long | R/W | "1 corresponds to Heading 1 …" | https://learn.microsoft.com/en-us/office/vba/api/word.captionlabel.chapterstylelevel |
| CaptionLabel.Separator | WdSeparatorType | R/W | The character "between the chapter number and the sequence number". | https://learn.microsoft.com/en-us/office/vba/api/word.captionlabel.separator |
| CaptionLabel.NumberStyle | WdCaptionNumberStyle | R/W | | https://learn.microsoft.com/en-us/office/vba/api/word.captionlabel.numberstyle |
| CaptionLabel.Position | WdCaptionPosition | R/W | | https://learn.microsoft.com/en-us/office/vba/api/word.captionlabel.position |
| WdCaptionLabelID | wdCaptionFigure = -1; wdCaptionTable = -2; **wdCaptionEquation = -3** | enum | | https://learn.microsoft.com/en-us/office/vba/api/word.wdcaptionlabelid |
| WdCaptionPosition | wdCaptionPositionAbove = 0; wdCaptionPositionBelow = 1 | enum | | https://learn.microsoft.com/en-us/office/vba/api/word.wdcaptionposition |
| WdSeparatorType | Hyphen = 0; Period = 1; Colon = 2; EmDash = 3; EnDash = 4 (each prefixed `wdSeparator`) | enum | | https://learn.microsoft.com/en-us/office/vba/api/word.wdseparatortype |
| WdCaptionNumberStyle | Arabic = 0; UppercaseRoman = 1; LowercaseRoman = 2; UppercaseLetter = 3; LowercaseLetter = 4; … (34 values, prefixed `wdCaptionNumberStyle`) | enum | | https://learn.microsoft.com/en-us/office/vba/api/word.wdcaptionnumberstyle |
| Chapter numbers (UI) | The chapter-heading style must be unique, with multilevel list numbering applied. The separator gives e.g. "Figure 2-1". | — | "you must use a unique heading style for chapter headings". | https://support.microsoft.com/en-us/office/add-chapter-numbers-to-captions-in-word-76aa8706-3453-4a72-ad5c-b87d90ce1698 |

### 3.3 SEQ field (Office support "Field codes: Seq (Sequence) field")

Source: https://support.microsoft.com/en-us/office/field-codes-seq-sequence-field-062a387b-dfc9-4ef8-8235-29ee113d59be. It applies to desktop Word on Windows and Mac: "Word for the web doesn't support field codes."

| Part | Meaning (from the page) |
|---|---|
| Syntax | `{ SEQ Identifier [Bookmark ] [Switches ] }` |
| Identifier | "The name you assign to a series of items to number. The name must start with a letter and is limited to 40 characters (letters, numbers, and underscores)." |
| Bookmark | "Include a bookmark name to refer to an item elsewhere in the document", e.g. `{ SEQ illust Pic2 }`. |
| `\c` | "Repeats the closest preceding sequence number." |
| `\h` | "Hides the field result." Not hidden "if a Format (\*) switch is also present". |
| `\n` | "Inserts the next sequence number for the specified items. This is the default." |
| `\r n` | "Resets the sequence number to the specified number n", e.g. `{ SEQ figure \r 3 }`. |
| `\s` | "Resets the sequence number at the heading level following the "s"", e.g. `{ SEQ figure \s 2 }`. |
| Scope note | "A Seq field in a header, footer, annotation, or footnote doesn't affect the sequence numbering that results from Seq fields in the document text." |

### 3.4 STYLEREF field (for chapter-style numbers such as (2.1))

Source: https://support.microsoft.com/en-us/word/field-codes-styleref-field (Windows and Mac desktop).

| Part | Meaning (from the page) |
|---|---|
| Syntax | `{ STYLEREF StyleIdentifier [Switches ] }`. StyleIdentifier: "The name of the style … If the style name contains spaces, enclose it in quotation marks." |
| `\n` | "display the entire paragraph number for a referenced paragraph without trailing periods". |
| `\r`, `\w`, `\t`, `\l`, `\p` | Relative-context number; full-context number; suppress non-numerical text (with \n, \r or \w); last instance on page; "above"/"below". |
| Location rule | In document text, "Word searches backward from the StyleRef field for the specified style. If the style isn't found, Word searches forward". |

A chapter-style number composed from documented switches is `({ STYLEREF "Heading 1" \n }.{ SEQ Equation \s 1 })`. That this combination yields "2.1" is **not stated** on either page; it composes documented switches. `STYLEREF … \s` is not on the fetched StyleRef page. A numeric heading level as StyleIdentifier is **not stated**.

### 3.5 Creating and updating fields from code

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Fields.Add(Range, Type?, Text?, PreserveFormatting?) | Range (req); Type: WdFieldType, default **wdFieldEmpty**; Text: "Additional text needed for the field. For example, if you want to specify a switch for the field, you would add it here."; PreserveFormatting → Field | method | "If the range isn't collapsed, the field replaces the range." The example `Fields.Add Range:=…, Type:=wdFieldListNum, Text:="\s 3"` shows Text appended after the field keyword. So `Type:=wdFieldSequence, Text:="Equation \* ARABIC"` is the analogous call (**inferred**). Whether `Type:=wdFieldEmpty, Text:="SEQ Equation"` makes Text the whole code is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.fields.add |
| WdFieldType | wdFieldEmpty = -1 ("A field added by pressing Ctrl+F9 in the user interface is an Empty field."); **wdFieldSequence = 12** ("Seq (Sequence) field."); wdFieldStyleRef = 10; wdFieldRef = 3 | enum | | https://learn.microsoft.com/en-us/office/vba/api/word.wdfieldtype |
| Field.Update | → Boolean | method | "Returns True if the field is updated successfully." See Discrepancy 9. | https://learn.microsoft.com/en-us/office/vba/api/word.field.update |
| Fields.Update | → Long | method | "Returns 0 (zero) if no errors occur …, or returns a Long that represents the index of the first field that contains an error." It also says "the number may be incorrect". | https://learn.microsoft.com/en-us/office/vba/api/word.fields.update |
| Field.Code | Range | R/W | "everything that's enclosed by the field characters ({ }) including the leading space and trailing space". The example rewrites `.Code.Text` and then calls `.Update`. | https://learn.microsoft.com/en-us/office/vba/api/word.field.code |
| Field.Result | Range | R/W | "a field's result". | https://learn.microsoft.com/en-us/office/vba/api/word.field.result |
| Range.Fields | Fields | R | | https://learn.microsoft.com/en-us/office/vba/api/word.range.fields |

### 3.6 Cross-references to equation numbers

| Member | Type / values | R/W | Notes | Source URL |
|---|---|---|---|---|
| Bookmarks.Add(Name, Range?) | Name: String, "cannot be more than 40 characters or include more than one word" → Bookmark | method | Sargent bookmarks the equation number (names starting "eq") and cross-references the bookmark. | https://learn.microsoft.com/en-us/office/vba/api/word.bookmarks.add |
| Range.InsertCrossReference(ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink?, IncludePosition?, SeparateNumbers?, SeparatorString?) | ReferenceType: WdReferenceType, WdCaptionLabelID or a user caption label; ReferenceKind: WdReferenceKind; ReferenceItem: a bookmark name for wdRefTypeBookmark, otherwise "the item number or name in the Reference type option in the Cross-reference dialog box" | method | "Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined (for example, an equation, figure, or table)." | https://learn.microsoft.com/en-us/office/vba/api/word.range.insertcrossreference |
| Document.GetCrossReferenceItems(ReferenceType) | WdReferenceType → array | method | "corresponds to the items listed in the For which box in the Cross-reference dialog box". Caption labels as input: **not stated** (Discrepancy 8). | https://learn.microsoft.com/en-us/office/vba/api/word.document.getcrossreferenceitems |
| WdReferenceType | wdRefTypeNumberedItem = 0; wdRefTypeHeading = 1; wdRefTypeBookmark = 2; wdRefTypeFootnote = 3; wdRefTypeEndnote = 4 | enum | | https://learn.microsoft.com/en-us/office/vba/api/word.wdreferencetype |
| WdReferenceKind | wdContentText = -1; wdNumberRelativeContext = -2; wdNumberNoContext = -3; wdNumberFullContext = -4; wdEntireCaption = 2 ("label, number, and any additional caption of specified equation, figure, or table"); wdOnlyLabelAndNumber = 3; wdOnlyCaptionText = 4; wdFootnoteNumber = 5; wdEndnoteNumber = 6; wdPageNumber = 7; wdPosition = 15; wdFootnoteNumberFormatted = 16; wdEndnoteNumberFormatted = 17 | enum | For a bookmark, "Bookmark text" in the dialog corresponds to wdContentText ("Insert text value of the specified item"). That mapping is **inferred** from the description. | https://learn.microsoft.com/en-us/office/vba/api/word.wdreferencekind |

### 3.7 Classic layouts: tabs, or a three-cell table

| Approach | What Microsoft says | Members | Source URL |
|---|---|---|---|
| Center tab before the equation, right tab before the number | "This works perfectly for simple equations, but … the presence of tabs on the line outside the math zone makes Word think the math zone is an inline math zone and therefore uses inline math typography" (small limits, script-size fractions). The tabs also do not follow margin changes. | `Paragraph(s).TabStops.Add(Position, Alignment?, Leader?)`: Position in points "relative to the left margin"; WdTabAlignment: Left = 0, Center = 1, Right = 2, Decimal = 3, Bar = 4, List = 6 | https://learn.microsoft.com/en-us/archive/blogs/murrays/more-on-equation-numbering; https://learn.microsoft.com/en-us/office/vba/api/word.tabstops.add; https://learn.microsoft.com/en-us/office/vba/api/word.wdtabalignment |
| Three-cell table (Dong Yu) | "If you use the three-cell table approach … these problems don't occur and the approach works with Word 2007 on." "The first and third cells have the identical, small widths. The second cell is wide and contains the equation. The third cell contains the equation number." Sargent's stated preference: "I like using the three-cell table approach with the formatting and referencing methods described in the book." | `Tables.Add(Range, NumRows, NumColumns, DefaultTableBehavior?, AutoFitBehavior?)`; `Table.Cell(Row, Column)`; `Cell.Range` | same post; https://learn.microsoft.com/en-us/office/vba/api/word.tables.add; https://learn.microsoft.com/en-us/office/vba/api/word.table.cell; https://learn.microsoft.com/en-us/office/vba/api/word.cell.range |

---

## 4. Reading equations back

| Question | What the docs say | Source URL |
|---|---|---|
| `OMath.Range.Text` of a built-up equation | `Range.Text` "returns the plain, unformatted text of the range". What this contains for a built-up (Professional) equation is **not stated**. The structure delimiters U+FDD0/U+FDEE/U+FDEF that Sargent describes belong to **RichEdit's** in-memory format, not Word's Range.Text. | https://learn.microsoft.com/en-us/office/vba/api/word.range.text; https://devblogs.microsoft.com/math-in-office/officemath/ |
| Linear text | `OMath.Linearize` ("Converts an equation to linear format"), then `OMath.Range.Text`, then `OMath.BuildUp` to restore. Each call is documented. That the round trip is lossless, and which linear format is produced, are **not stated**. It edits the document (one undo step per call is **not stated**). | https://learn.microsoft.com/en-us/office/vba/api/word.omath.linearize, https://learn.microsoft.com/en-us/office/vba/api/word.omath.buildup |
| Structure as XML (Windows) | `OMath.Range.WordOpenXML` returns the Flat OPC XML "needed to represent the specified range". Math appears as `m:oMathPara` / `m:oMath` (ISO 29500 §22.1.2.77–78). The element names are documented. That WordOpenXML of an OMath range contains them follows only from the general description. Whether the result is lossless is **not stated**. | https://learn.microsoft.com/en-us/office/vba/api/word.range.wordopenxml; Open XML SDK pages in §2.4 |
| Structure walk | `OMath.Functions` → `OMathFunction.Type` and the typed accessors (§1.4); `OMath.Breaks` (§1.5). | §1.3–1.5 |
| Clipboard plain text | `Options.OMathCopyLF` chooses Linear Format vs MathML for equations "in plain text" (clipboard). | https://learn.microsoft.com/en-us/office/vba/api/word.options.omathcopylf |
| Mac | `linearize` the `math object`, then read `content` of its `text range` (Word.sdef:9156, "Returns or sets the text in the text range."), then `build up`. Mac has no WordOpenXML, so there is no lossless read path. | Word.sdef |

---

## 5. Word for Mac (Apple Events, Word.sdef)

Source: `/Applications/Microsoft Word.app/Contents/Resources/Word.sdef`, Word 16.114.4. All math items are in the "Microsoft Word Suite" (code `sWRD`). The standard `make`, `delete`, `count` and `exists` come from the included `CocoaStandard.sdef` (Word.sdef:6); whether they work on math classes is **not stated**.

### 5.1 Math commands

| Command (code) | Direct parameter | Parameters | Result | Description (sdef) | Line |
|---|---|---|---|---|---|
| create new equation (sWRD14BB) | — | from range: text range; in document: document?; in range: selection object?; in selection: selection object? | (none declared) | "Creates an equation, from the text equation contained within the specified range, and returns a text range object that contains the new equation." | 2214 |
| build up (sWRD14AD) | math object | — | — | "Converts an equation to professional/built up format." | 1993 |
| linearize (sWRD14AC) | math object | — | — | "Converts an equation to linear/built down format." | 2631 |
| convert to math text (sWRD14AE) | math object | — | — | "Converts an equation to math text." | 2129 |
| convert to normal text (sWRD14AF) | math object | — | — | "Converts an equation to normal text." | 2132 |
| convert to literal text (sWRD14B8) | math object | — | — | "Converts an equation to literal text." | 2126 |
| insert math break (sWRD14a2) | math object | at range: text range | math break | "Inserts a break into an equation and returns a math break object" | 2596 |
| add math function (sWRD14AG) | math object | in location: text range; math function type: WdOMathFunctionType; number of arguments: integer?; number of columns: integer? | (none declared) | "Inserts a new structure, such as a fraction, into an equation at the specified position." | 1910 |
| add math argument (sWRD14AL) | specifier | before arg: integer?; to function: math function?; to matrix row?; to matrix column? | math object | "You must specify only one function, row, or argument." | 1902 |
| add matrix row / add matrix column | math matrix | before row? / before column? | math matrix row / column | | 1926 / 1921 |
| add math ac entry (sWRD14aF) | math autocorrect | name: text; value: text | — | "Creates an equation auto correct entry." | 1897 |
| add math recognized function (sWRD14aH) | math autocorrect | name: text | — | "Creates a new recognized function." | 1917 |
| remove subscript / remove superscript | math sub and super script | — | math function | | 2972 / 2976 |
| convert sub and super scripts to left scripts / convert left scripts to sub and super scripts | math sub and super script / math left scripts | — | math function | | 2122 / 2114 |
| lower limit to upper limit / upper limit to lower limit | math lower limit / math upper limit | — | math function | (no description) | 2647 / 3233 |

### 5.2 Math classes

| Class (code) | Elements | Properties (type, access) | Line |
|---|---|---|---|
| math object (o112), plural "math objects" — "Represents an equation" | math function, math break | align point (integer, rw); argument index (integer, r); containing function (math function, r); **display type (WdOMathType, rw)**; **justification (WdOMathJc, rw)**; nesting level (integer, r); parent argument (math object, r); parent column (math matrix column, r); parent object (math object, r); parent row (math matrix row, r); script size (integer, rw); **text range (text range, r)** | 4284 |
| math function (o113) | math object | function type (WdOMathFunctionType, r); text range (r); math obj (math object, r); typed accessors, all r: accent, border box, box, delimiter, equation array, fraction, func, group char, left scripts, lower limit, matrix, nary, overbar, phantom, radical, sub and super script, subscript, superscript, upper limit | 4214 |
| math break (o11P) | — | align at (integer, rw); text range (r) | 4183 |
| math autocorrect (o11T) | math autocorrect entry, math recognized function | replace text (boolean, rw); use outside equations (boolean, rw) | 4156 |
| math autocorrect entry (o11U) | — | autocorrect value (text, rw); entry_index (integer, r); name (text, rw) | 4151 |
| math recognized function (o11V) | — | entry_index (r); name (text, r) | 4314 |
| Structure classes | — | math accent, math bar (is top bar), math border box (hide top/bottom/left/right, strikethroughs), math box (is differential, is operator emulator, non breaking), math delimiter (beginning/ending/separator character, delimiter shape, hide left/right delimiter, should grow), math equation array (distribute equally, row spacing, row spacing rule, use max width, vertical alignment), math fraction (numerator, denominator, fraction type), math func (func name), math group char, math left scripts, math lower/upper limit, math matrix (+ row/column), math nary (operator character, hides lower/upper limit, should grow, use sub and super script positioning), math phantom, math radical (degree, hide degree), math sub and super script (align scripts), math subscript, math superscript | 4147–4335 |

Containers that have `math object` as an element: document (Word.sdef:3742), selection object (4459), text range (9145), math delimiter, math equation array, math function, math matrix row/column.

### 5.3 Math-related properties on other classes

| Class.property | Type | Access | Description (sdef) | Line |
|---|---|---|---|---|
| application.math ac | math autocorrect | r | "the auto correct entries for equations" | 3476 |
| application.settings | Word options | r | "Returns the word options object." | 3486 |
| Word options.automatically build up equations | boolean | rw | "whether Microsoft Word automatically converts equations to professional format" | 3314 |
| Word options.plain equations use plain text | boolean | rw | "Returns or sets if equations are represented in plain text; false uses MathML" | 3385 |
| document.math binary operator break | WdOMathBreakBin | r (see Discrepancy 2) | | 3779 |
| document.math default justification | WdOMathJc | r | | 3780 |
| document.math font name | text | r | | 3781 |
| document.math left margin / math right margin | real | r | | 3782 / 3783 |
| document.math subtraction operator | WdOMathBreakSub | r | | 3784 |
| document.math wrap | real | r | | 3785 |
| document.integrals uses subscripts | boolean | r | "default location of limits for integrals" | 3774 |
| document.nary uses subscripts | boolean | r | "… n-ary objects other than integrals" | 3787 |
| document.use default math settings | boolean | r | "whether to use the default math settings when creating new equations" (no VBA counterpart) | 3822 |
| document.use small fractions | boolean | r | | 3823 |

### 5.4 Math enumerations (Mac)

The low 16 bits of each enumerator code match the VBA numeric value (for example 0x02d60007 ↔ wdOMathJcInline = 7). This is an observation, not a documented rule. appscript uses the names.

| Enum | Enumerators (code) | Line |
|---|---|---|
| WdOMathType | equation display professional (0x02d40000) "professional/built up format"; equation display inline (0x02d40001) "linear/built down style" | 6871 |
| WdOMathJc | equation justification center group (…01), center (…02), left (…03), right (…04), inline (…07) | 6879 |
| WdOMathBreakBin | math binary operator break before (…00), break after (…01), repeat (…02) | 6886 |
| WdOMathBreakSub | minus minus (…00), plus minus (…01), minus plus (…02) | 6891 |
| WdOMathFunctionType | math accent type (1), math function bar type (2), math box type (3), math boxed formula type (4), math delimiters type (5), math equation array type (6), math fraction type (7), math function apply type (8), math stretch stack type (9), math lower limit type (10), math upper limit type (11), math matrix type (12), math nary operator type (13), math phantom type (14), math left sub superscript type (15), math radical type (16), math subscript type (17), math sub superscript type (18), math superscript type (19), math text type (20), math normal text type (21), math literal text type (22) | 6824 |
| WdOMathFracType | bar (0), stack (1), slashed (2), linear (3) | 6858 |
| WdOMathHorizAlignType / WdOMathVertAlignType / WdOMathSpacingRule / WdOMathShapeType | center/left/right; center/top/bottom; single/one and a half/double/exactly/multiple; center vertically/match content size | 6848 / 6853 / 6864 / 6875 |

### 5.5 Fields, captions, cross-references and bookmarks (Mac)

| Item | Detail | Line |
|---|---|---|
| create new field | text range: text range; field type: WdFieldType?; field text: text? ("The field code for the new field."); preserve formatting: boolean? | 2220 |
| WdFieldType | field empty (0x0246ffff = -1), field style ref (0x0247000a = 10), field sequence (0x0247000c = 12) | 5926–5937 |
| update field | direct: field → boolean | 3217 |
| field class | field code (text range), result range (text range), field type (r), show codes | 3876 |
| insert caption | at: text range; caption label: **WdCaptionLabelID**; title: text?; caption position: WdCaptionPosition? (no ExcludeLabel, no custom label) | 2516 |
| caption label class | built in (r), caption label id (r), caption label position, chapter style level, include chapter number, name (r), number style, separator. application has element caption label (3436). | 3582 |
| WdCaptionLabelID | caption figure (0x0261ffff = -1), caption table (-2), caption equation (0x0261fffd = -3) | 6666 |
| insert cross reference | at; reference type: **WdReferenceType**; reference kind: WdReferenceKind; reference item: text; insert as hyperlink?; include position? | 2530 |
| get cross reference items | direct: document; reference type: WdReferenceType → list of text | 2383 |
| bookmark class | name (r), start of bookmark (rw), end of bookmark (rw), text object (r). There is **no bookmark-creation command** (`copy bookmark` exists, 2135). Whether `make new bookmark` works is **not stated**. | 3505 |
| run VB macro | macro name: text. This runs a VBA macro in Word for Mac. Whether Mac VBA has `Range.InsertXML` is **not stated**. | 3009 |
| XML | No `insert xml`, `word open xml` or `xml` member anywhere in Word.sdef. | — |

### 5.6 Windows → Mac equivalence

| Windows (VBA) | Mac (Word.sdef) | Equivalent? |
|---|---|---|
| Range/Selection/Document.OMaths | `math objects` of text range / selection object / document | yes |
| OMaths.Add(Range) | `create new equation from range r` | yes (no declared result) |
| OMaths.BuildUp / Linearize (whole collection) | `build up` / `linearize` take one `math object`; passing a list is **not stated** | partial |
| OMath.BuildUp / Linearize / ConvertToMathText / ConvertToNormalText / ConvertToLiteralText | `build up` / `linearize` / `convert to math text` / `convert to normal text` / `convert to literal text` | yes |
| OMath.Remove | none (standard `delete` **not stated**) | no |
| OMath.Range / Type / Justification / AlignPoint / ArgSize / ArgIndex / NestingLevel / ParentArg / ParentFunction / ParentOMath / ParentCol / ParentRow | text range / display type / justification / align point / script size / argument index / nesting level / parent argument / containing function / parent object / parent column / parent row | yes |
| OMath.Functions / Breaks | elements `math function` / `math break` | yes |
| OMathFunctions.Add | `add math function` | yes |
| OMathBreaks.Add / OMathBreak.AlignAt / Delete | `insert math break` / `align at` / none | partial |
| Document.OMathBreakBin, BreakSub, FontName, IntSubSupLim, NarySupSubLim, Jc, LeftMargin, RightMargin, SmallFrac, Wrap | math binary operator break, math subtraction operator, math font name, integrals uses subscripts, nary uses subscripts, math default justification, math left/right margin, use small fractions, math wrap | yes, but all marked `r` |
| Options.OMathAutoBuildUp / OMathCopyLF | settings: automatically build up equations / plain equations use plain text | yes |
| Application.OMathAutoCorrect (+ Entries, Functions, ReplaceText, UseOutsideOMath, Entries.Add, Functions.Add) | math ac (+ elements, replace text, use outside equations, add math ac entry, add math recognized function) | yes (entry delete **not stated**) |
| Range.InsertXML / WordOpenXML / XML, Document.WordOpenXML | none | **no** |
| Fields.Add / Field.Update / Code / Result | create new field / update field / field code / result range | yes |
| Range.InsertCaption | insert caption (WdCaptionLabelID only, no ExcludeLabel) | partial |
| CaptionLabels.Add | element of application; creation via `make` **not stated** | unclear |
| Range.InsertCrossReference / GetCrossReferenceItems | insert cross reference / get cross reference items (WdReferenceType only) | partial |
| Bookmarks.Add | none | no |
| TabStops.Add, Tables.Add | `tab stop` class (4518) and the Table Suite exist; not examined for this file | not checked |
| LaTeX / UnicodeMath selection | none | no (none on Windows either) |

---

## 6. Recipes built only from documented steps

Each step is marked **[doc]** (documented, with the source in the sections above) or **[not stated]**.

### 6.1 Windows: numbered display equation from UnicodeMath

1. Insert a new empty paragraph and get its Range `r`, collapsed before the paragraph mark. **[doc]** (word-object-model.md)
2. Choose a layout. Two options are documented:
   - **a. `#` route.** Set `r.Text = "E=mc^2#(1)"` **[doc: Range.Text]**. Call `eq_range = r.OMaths.Add(r)` **[doc]**, then `eq_range.OMaths(1).BuildUp()` **[doc]**. That BuildUp from code yields the flush-right number like the UI's Enter does is **[not stated]**. That the parser treats the text as UnicodeMath when the user's ribbon is set to LaTeX is **[not stated]**. For automatic numbering, the "1" would have to become a SEQ field inside the math zone; doing that from code is **[not stated]**.
   - **b. Three-cell table route.** `t = doc.Tables.Add(r, 1, 3)` **[doc]**. Put the UnicodeMath text into `t.Cell(1,2).Range`, then OMaths.Add and BuildUp **[doc]**. Optionally set `Type = wdOMathDisplay` (0) and `Justification = wdOMathJcCenter` (2) **[doc]**. In `t.Cell(1,3).Range`, insert "(", then `Fields.Add(range, wdFieldSequence (12), "Equation")` **[doc; Text composition inferred from the LISTNUM example]**, then ")". Set the cell widths and right-align the third cell **[doc elsewhere: Table/ParagraphFormat]**. Sargent says this layout keeps display typography **[doc: blog]**.
3. Bookmark the SEQ field's result range: `doc.Bookmarks.Add("eq_einstein", field.Result)` **[doc]**. Name: at most 40 characters, one word.
4. Cross-reference it: `r2.InsertCrossReference(wdRefTypeBookmark (2), wdContentText (-1), "eq_einstein", True)` **[doc; "Bookmark text" ↔ wdContentText inferred]**.
5. Renumber: `doc.Fields.Update()` or `Field.Update()` **[doc]**.
6. Chapter style (2.1): the number cell becomes `({ STYLEREF "Heading 1" \n }.{ SEQ Equation \s 1 })` **[each switch doc; the combination's result not stated]**. Alternatively set `CaptionLabels(wdCaptionEquation).IncludeChapterNumber = True` and `ChapterStyleLevel = 1` with `InsertCaption(..., ExcludeLabel=True)` **[doc]**; whether that caption stays on the equation's line is **[not stated]**.

### 6.2 Windows: from LaTeX

- There is no documented way to make `BuildUp` parse LaTeX **[not stated / absent]**.
- Documented alternatives:
  1. Convert LaTeX to UnicodeMath yourself, then follow 6.1 **[UnicodeMath syntax doc: UTN28]**.
  2. Convert LaTeX to MathML yourself; transform it with the Office `mml2omml.xsl` from the winword.exe directory **[doc: 2007 blog]**; wrap the `m:oMath`, inside `m:oMathPara` for display, in a Flat OPC `pkg:package/…/w:body/w:p` with `xmlns:m` **[element semantics doc: ISO 29500; Flat OPC + InsertXML doc: 2010 article]**; call `r.InsertXML(flat_opc)` **[doc]**. That this specific combination renders as an equation is **[not stated]**. So is whether `InsertXML(mathml, Transform)` works directly.
  3. Convert through Office RichEdit `SetText2(tomConvertTeX)` and `GetText2(tomConvertOMML)` **[doc: blog only]**, then step 2's InsertXML.
- Numbering then follows 6.1 steps 2b–6, with the OMML in the middle cell.

### 6.3 Mac (Apple Events): UnicodeMath

1. Put the linear text into a text range, e.g. `set content of r to "E=mc^2#(1)"` **[doc: text range content rw]**.
2. `create new equation from range r` **[doc]**.
3. Get the `math object` from the range's `math objects` **[doc: element]**, then `build up` it **[doc]**. Whether the `#` layout results is **[not stated]**.
4. For a numbered layout without `#`: create a 1×3 table (Table Suite, not examined), then put the SEQ field with `create new field text range rr field type field sequence field text "Equation"` **[doc]** and `update field` **[doc]**.
5. Cross-reference: `insert cross reference … reference type reference type bookmark` needs a bookmark, but there is no bookmark-creation command **[not stated]**. `insert caption` cannot exclude the label **[doc: no parameter]**.

### 6.4 Mac: from LaTeX

There is no Apple Events path: no LaTeX switch and no XML insertion in Word.sdef. The only documented route is converting LaTeX to UnicodeMath yourself and following 6.3. `run VB macro` exists; whether Mac VBA's `Range.InsertXML` works is **[not stated]**.

---

## 7. Sources

Office VBA reference (Word), all fetched with HTTP 200 on 2026-11-08. Each name below is a page at https://learn.microsoft.com/en-us/office/vba/api/<name>, e.g. https://learn.microsoft.com/en-us/office/vba/api/word.omath.buildup:
- OMath: word.omath, .buildup, .linearize, .converttomathtext, .converttonormaltext, .converttoliteraltext, .remove, .range, .type, .justification, .alignpoint, .argsize, .argindex, .nestinglevel, .functions, .breaks, .parent, .parentarg, .parentfunction, .parentomath, .parentcol, .parentrow.
- Collections: word.omaths (.add, .buildup, .linearize, .count, .item); word.range.omaths, word.selection.omaths, word.document.omaths.
- Structures: word.omathfunction (.type, .range, .omath, .args, .remove), word.omathfunctions (.add, .item, .count), word.omathbreaks.add, word.omathbreak (.alignat, .range, .delete), word.omathfrac.type, word.omatheqarray, word.omathbar.
- Document math settings: word.document.omathbreakbin, omathbreaksub, omathfontname, omathintsubsuplim, omathjc, omathleftmargin, omathnarysupsublim, omathrightmargin, omathsmallfrac, omathwrap; word.options.omathautobuildup, omathcopylf.
- Autocorrect: word.application.omathautocorrect, word.omathautocorrect (.entries, .functions, .replacetext, .useoutsideomath), word.omathautocorrectentries (.add, .item), word.omathautocorrectentry (.name, .value, .index, .delete), word.omathrecognizedfunctions.add, word.omathrecognizedfunction (.name, .delete).
- Math enums: word.wdomathtype, wdomathjc, wdomathbreakbin, wdomathbreaksub, wdomathfunctiontype, wdomathfractype, wdomathhorizaligntype, wdomathvertaligntype, wdomathshapetype, wdomathspacingrule.
- XML: word.range.insertxml, word.selection.insertxml, word.range.wordopenxml, word.document.wordopenxml, word.range.xml.
- Fields: word.fields.add, word.wdfieldtype, word.field.update, word.fields.update, word.field.code, word.field.result, word.range.fields.
- Captions and cross-references: word.range.insertcaption, word.captionlabels.add, word.captionlabel (.includechapternumber, .chapterstylelevel, .separator, .numberstyle, .position), word.application.captionlabels, word.wdcaptionlabelid, word.wdcaptionposition, word.wdseparatortype, word.wdcaptionnumberstyle, word.range.insertcrossreference, word.document.getcrossreferenceitems, word.wdreferencetype, word.wdreferencekind, word.bookmarks.add.
- Layout: word.tabstops.add, word.wdtabalignment, word.tables.add, word.table.cell, word.cell.range, word.range.text.

Other Microsoft pages:
- https://learn.microsoft.com/en-us/office/vba/api/toc.json
- https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.options?view=word-pia
- https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.omath?view=word-pia
- https://support.microsoft.com/en-us/office/linear-format-equations-using-unicodemath-and-latex-in-word-2e00618d-b1fd-49d8-8cb4-8d17f25754f8
- https://support.microsoft.com/en-us/office/write-an-equation-or-formula-1d01cabc-ceb1-458d-bc70-7f9737722702
- https://support.microsoft.com/en-us/office/field-codes-seq-sequence-field-062a387b-dfc9-4ef8-8235-29ee113d59be
- https://support.microsoft.com/en-us/word/field-codes-styleref-field
- https://support.microsoft.com/en-us/office/add-chapter-numbers-to-captions-in-word-76aa8706-3453-4a72-ad5c-b87d90ce1698
- https://learn.microsoft.com/en-us/archive/blogs/murrays/equation-numbering-in-office-2016
- https://learn.microsoft.com/en-us/archive/blogs/murrays/more-on-equation-numbering
- https://learn.microsoft.com/en-us/archive/blogs/murrays/latex-math-in-office
- https://learn.microsoft.com/en-us/archive/blogs/murrays/improved-mathml-support-in-word-2007
- https://learn.microsoft.com/en-us/archive/blogs/murrays/mathml-and-ecma-math-omml
- https://devblogs.microsoft.com/math-in-office/officemath/
- https://devblogs.microsoft.com/math-in-office/default-math-properties/
- https://devblogs.microsoft.com/math-in-office/setting-and-getting-text-in-various-formats/
- https://devblogs.microsoft.com/math-in-office/switching-from-latex-to-unicodemath-input-mode/
- https://devblogs.microsoft.com/math-in-office/some-unicodemath-enhancements/
- https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff191178(v=office.14)
- https://learn.microsoft.com/en-us/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.officemath?view=openxml-3.0.1 (and .paragraph, .equationarray, .paragraphproperties, .justification)
- https://learn.microsoft.com/en-us/windows/win32/api/tom/nf-tom-itextrange2-settext2
- https://www.unicode.org/notes/tn28/UTN28-PlainTextMath-v3.2.pdf (UnicodeMath 3.2, Murray Sargent, 13-Apr-23)

Local:
- `/Applications/Microsoft Word.app/Contents/Resources/Word.sdef` (Word 16.114.4)
- `/Applications/Microsoft Word.app/Contents/Resources/mathml2omml.xsl` and `omml2mathml.xsl`
