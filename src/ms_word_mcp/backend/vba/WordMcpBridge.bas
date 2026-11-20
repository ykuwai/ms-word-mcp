Attribute VB_Name = "WordMcpBridge"
' ms-word-mcp's VBA bridge (backend/bridge.py). Reading a paragraph over COM
' or Apple Events asks Word once a property, a few milliseconds each time;
' VBA runs inside Word and reads many paragraphs in one call. Word for Mac's
' AppleScript also has no comment replies or done state, which VBA has.
'
' Windows calls WordMcpRun through Application.Run and gets the answer back
' as its value. The Mac's `run VB macro` passes nothing and returns nothing,
' so ms-word-mcp writes the request in a file next to this template, runs
' WordMcpCall, and reads the answer file it writes. Both go through Dispatch.
'
' Strings travel as their UTF-16 code units in decimal, comma-separated
' ("72,105" is "Hi"), since VBA's file I/O on the Mac does not keep Unicode.
' Numbers are written with Str, which always uses a period. Every entry point
' traps its errors: an untrapped one would stop Word on a dialog that waits
' for a person.
Option Explicit

' Raise it, and VERSION in backend/bridge.py, with any change to this module:
' a server that finds another version loaded in Word replaces it.
Public Const BRIDGE_VERSION As String = "5"
Private Const REQUEST_FILE As String = "ms-word-mcp-bridge.request"
Private Const ANSWER_FILE As String = "ms-word-mcp-bridge.answer"
Private Const TEMPLATE_FILE As String = "ms-word-mcp-bridge.dotm"

Private Function Folder() As String
    Folder = ThisDocument.Path & Application.PathSeparator
End Function

Private Function Decode(s As String) As String
    Dim parts() As String, i As Long, out As String
    If Len(s) = 0 Then
        Decode = ""
        Exit Function
    End If
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        out = out & ChrW(CLng(parts(i)))
    Next i
    Decode = out
End Function

Private Function Encode(ByVal s As String) As String
    Dim i As Long, n As Long, parts() As String
    n = Len(s)
    If n = 0 Then
        Encode = ""
        Exit Function
    End If
    ReDim parts(n - 1)
    For i = 1 To n
        parts(i - 1) = CStr(AscW(Mid$(s, i, 1)) And &HFFFF&)
    Next i
    Encode = Join(parts, ",")
End Function

Private Function Num(ByVal x As Variant) As String
    Num = Trim$(Str$(x))
End Function

Private Function Request() As Variant
    Dim f As Integer, line As String, lines() As String, n As Long
    ReDim lines(0)
    f = FreeFile
    Open Folder() & REQUEST_FILE For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        ReDim Preserve lines(n)
        lines(n) = line
        n = n + 1
    Loop
    Close #f
    Request = lines
End Function

Private Sub Answer(text As String)
    Dim f As Integer
    f = FreeFile
    Open Folder() & ANSWER_FILE For Output As #f
    Print #f, text
    Close #f
End Sub

Private Function FindDocument(wanted As String) As Document
    Dim d As Document
    For Each d In Documents
        If d.FullName = wanted Then
            Set FindDocument = d
            Exit Function
        End If
    Next d
    For Each d In Documents
        If d.Name = wanted Then
            Set FindDocument = d
            Exit Function
        End If
    Next d
    Err.Raise 5, , "No open document is called " & wanted
End Function

Private Function Seconds(began As Single) As Single
    Dim now As Single
    now = Timer
    If now < began Then now = now + 86400
    Seconds = now - began
End Function

Private Function Dispatch(what As String, r As Variant) As String
    On Error GoTo Trap
    Select Case what
        Case "Ping": Dispatch = "ok" & vbLf & BRIDGE_VERSION
        Case "Paragraphs": Dispatch = ReadParagraphs(r)
        Case "Reply": Dispatch = Reply(r)
        Case "SetDone": Dispatch = SetDone(r)
        Case "CommentState": Dispatch = CommentState(r)
        Case "DeleteComments": Dispatch = DeleteComments(r)
        Case "FindReplace": Dispatch = FindReplace(r)
        Case "Member": Dispatch = Member(r)
        Case "AddStyle": Dispatch = AddStyle(r)
        Case "StyleList": Dispatch = StyleList(r)
        Case "GroupShapes": Dispatch = GroupShapes(r)
        Case "UngroupShape": Dispatch = UngroupShape(r)
        Case Else: Err.Raise 5, , "The bridge has no request called " & what
    End Select
    Exit Function
Trap:
    Dispatch = "error" & vbLf & Err.Number & vbLf & Encode(Err.Description)
End Function

' Windows: Application.Run "WordMcpRun", what, the request's lines joined by vbLf.
Public Function WordMcpRun(what As String, request As String) As String
    WordMcpRun = Dispatch(what, Split(request, vbLf))
End Function

' The Mac: the request file's first line says what, the rest is the request.
Public Sub WordMcpCall()
    On Error GoTo Trap
    Dim r As Variant, rest() As String, i As Long
    r = Request()
    ReDim rest(UBound(r))
    For i = 1 To UBound(r)
        rest(i - 1) = r(i)
    Next i
    Answer Dispatch(CStr(r(0)), rest)
    Exit Sub
Trap:
    Answer "error" & vbLf & Err.Number & vbLf & Encode(Err.Description)
End Sub

' The Mac, in a Word that was already running: run from a document made from
' this template, makes the template a global one of that Word until it quits
' or ms-word-mcp takes it out, as if Word had found it in its Startup folder.
Public Sub WordMcpInstall()
    On Error GoTo Trap
    AddIns.Add FileName:=Folder() & TEMPLATE_FILE, Install:=True
    Answer "ok" & vbLf & BRIDGE_VERSION
    Exit Sub
Trap:
    Answer "error" & vbLf & Err.Number & vbLf & Encode(Err.Description)
End Sub

' request: document, first, last, detail (1 or 0), revisions (1: count each
' paragraph's), seconds. answer: the document's paragraph count, then a line a
' paragraph read, until `last` or the time is up:
' index|revisions|outline level|style|text
' and with detail: |alignment|ascii font|east asian font|complex script font|
' size|complex script size|bold|italic|color|line spacing rule|line spacing|
' space before|space after
Private Function ReadParagraphs(r As Variant) As String
    Dim d As Document, p As Paragraph, rg As Range, f As Font
    Dim first As Long, last As Long, detail As Boolean, revisions As Boolean
    Dim limit As Single, began As Single, i As Long, n As Long, line As String
    Dim lines() As String, styleName As String, count As Long, changes As Long
    began = Timer
    Set d = FindDocument(Decode(CStr(r(0))))
    first = CLng(r(1))
    last = CLng(r(2))
    detail = (r(3) = "1")
    revisions = (r(4) = "1")
    limit = Val(r(5))
    count = d.Paragraphs.Count
    If last > count Then last = count
    ReDim lines(last - first + 1)
    lines(0) = "ok" & vbLf & count
    n = 1
    Set p = d.Paragraphs(first)
    For i = first To last
        Set rg = p.Range
        styleName = ""
        On Error Resume Next
        styleName = p.Style.NameLocal
        On Error GoTo 0
        changes = 0
        If revisions Then changes = rg.Revisions.Count
        line = i & "|" & changes & "|" & p.OutlineLevel & "|" _
            & Encode(styleName) & "|" & Encode(rg.Text)
        If detail Then
            Set f = rg.Font
            line = line & "|" & p.Format.Alignment & "|" & Encode(f.NameAscii) & "|" _
                & Encode(f.NameFarEast) & "|" & Encode(f.NameBi) & "|" & Num(f.Size) & "|" _
                & Num(f.SizeBi) & "|" & f.Bold & "|" & f.Italic & "|" & f.Color & "|" _
                & p.Format.LineSpacingRule & "|" & Num(p.Format.LineSpacing) & "|" _
                & Num(p.Format.SpaceBefore) & "|" & Num(p.Format.SpaceAfter)
        End If
        lines(n) = line
        n = n + 1
        If i = last Then Exit For
        If Seconds(began) > limit Then Exit For
        Set p = p.Next
    Next i
    ReDim Preserve lines(n - 1)
    ReadParagraphs = Join(lines, vbLf)
End Function

' request: document, the numbers to delete, highest first, comma-separated.
' A comment goes with its replies, as on Windows. answer: how many are left,
' and how many of the numbers were still there when their turn came.
Private Function DeleteComments(r As Variant) As String
    Dim d As Document, numbers() As String, i As Long, n As Long, asked As Long
    Set d = FindDocument(Decode(CStr(r(0))))
    numbers = Split(CStr(r(1)), ",")
    For i = 0 To UBound(numbers)
        n = CLng(numbers(i))
        If n <= d.Comments.Count Then
            d.Comments(n).DeleteRecursively
            asked = asked + 1
        End If
    Next i
    DeleteComments = "ok" & vbLf & d.Comments.Count & vbLf & asked
End Function

' The comment a reply answers; Nothing for a comment of its own (which may
' raise rather than answer Nothing).
Private Function AncestorOf(c As Comment) As Comment
    On Error Resume Next
    Set AncestorOf = c.Ancestor
End Function

' -- members of Word's object model, for what Word for Mac's scripting lacks --
'
' request: document, then an operation a line: how|path|member|argument|...
' how: get, let (one argument) or call. The path goes from the document,
' "/" between members, each Name or Name(argument;argument). An argument is
' n<number> (Str's period), s<UTF-16 units>, b1 or b0, or p<path> for an
' object. answer: a line an operation: n<number>, s<units>, b1/b0, o for an
' object, or e for nothing.
Private Function Member(r As Variant) As String
    Dim d As Document, i As Long, parts() As String, target As Object, args As Variant
    Dim lines() As String, n As Long, got As Variant, j As Long
    Set d = FindDocument(Decode(CStr(r(0))))
    ReDim lines(UBound(r))
    lines(0) = "ok"
    n = 1
    For i = 1 To UBound(r)
        If Len(r(i)) > 0 Then
            parts = Split(CStr(r(i)), "|")
            Set target = Resolve(d, CStr(parts(1)))
            If UBound(parts) >= 3 Then
                ReDim args(UBound(parts) - 3)
                For j = 3 To UBound(parts)
                    Arg d, CStr(parts(j)), args, j - 3
                Next j
            Else
                args = Array()
            End If
            On Error GoTo Failed
            Select Case parts(0)
                Case "get": got = Invoke(target, CStr(parts(2)), VbGet, args)
                Case "call": got = Invoke(target, CStr(parts(2)), VbMethod, args)
                Case "let"
                    Assign target, CStr(parts(2)), args(0)
                    got = Array(Empty)
                Case Else: Err.Raise 5, , "No such operation: " & parts(0)
            End Select
            On Error GoTo 0
            lines(n) = Typed(got)
            n = n + 1
        End If
    Next i
    ReDim Preserve lines(n - 1)
    Member = Join(lines, vbLf)
    Exit Function
Failed:
    Err.Raise Err.Number, , Err.Description & " (" & parts(0) & " " & parts(2) & ")"
End Function

' A property set. Word for Mac's VBA does not pass a value through
' CallByName (measured), so the ones the server sets are assigned by name.
Private Sub Assign(target As Object, name As String, value As Variant)
    Select Case LCase$(name)
        Case "sizebi": target.SizeBi = value
        Case "underlinecolor": target.UnderlineColor = value
        Case "readingorder": target.ReadingOrder = value
        Case "quickstyle": target.QuickStyle = value
        Case "priority": target.Priority = value
        Case "alternativetext": target.AlternativeText = value
        Case "verticalanchor": target.VerticalAnchor = value
        Case Else
            If IsObject(value) Then
                CallByName target, name, VbSet, value
            Else
                CallByName target, name, VbLet, value
            End If
    End Select
End Sub

' A member's value, in a one-item array: an object stays an object there,
' where assigning it would read its default member.
Private Function Invoke(target As Object, name As String, how As VbCallType, args As Variant) As Variant
    ' Word for Mac's VBA does not pass arguments through CallByName (measured:
    ' Range(0, 5) answered error 1004, Paragraphs.Item(1) an overflow), so the
    ' members called with arguments are called by name.
    Select Case LCase$(name) & "/" & (UBound(args) + 1)
        Case "range/2": Invoke = Array(target.Range(args(0), args(1)))
        Case "item/1": Invoke = Array(target.Item(args(0)))
        Case "add/0": Invoke = Array(target.Add())
        Case "add/1": Invoke = Array(target.Add(args(0)))
        Case "add/3": Invoke = Array(target.Add(args(0), args(1), args(2)))
        Case Else: Invoke = Called(target, name, how, args)
    End Select
End Function

Private Function Called(target As Object, name As String, how As VbCallType, args As Variant) As Variant
    Select Case UBound(args)
        Case -1: Called = Array(CallByName(target, name, how))
        Case 0: Called = Array(CallByName(target, name, how, args(0)))
        Case 1: Called = Array(CallByName(target, name, how, args(0), args(1)))
        Case 2: Called = Array(CallByName(target, name, how, args(0), args(1), args(2)))
        Case 3: Called = Array(CallByName(target, name, how, args(0), args(1), args(2), args(3)))
        Case Else: Err.Raise 5, , "Too many arguments for " & name
    End Select
End Function

Private Function Typed(got As Variant) As String
    If IsObject(got(0)) Then
        If got(0) Is Nothing Then Typed = "e" Else Typed = "o"
    ElseIf IsEmpty(got(0)) Or IsNull(got(0)) Then
        Typed = "e"
    ElseIf VarType(got(0)) = vbBoolean Then
        Typed = IIf(got(0), "b1", "b0")
    ElseIf VarType(got(0)) = vbString Then
        Typed = "s" & Encode(CStr(got(0)))
    Else
        Typed = "n" & Num(got(0))
    End If
End Function

Private Sub Arg(d As Document, code As String, args As Variant, at As Long)
    Select Case Left$(code, 1)
        Case "n"
            If InStr(code, ".") = 0 And InStr(UCase$(code), "E") = 0 Then
                args(at) = CLng(Mid$(code, 2))
            Else
                args(at) = Val(Mid$(code, 2))
            End If
        Case "s": args(at) = Decode(Mid$(code, 2))
        Case "b": args(at) = (Mid$(code, 2) = "1")
        Case "p": Set args(at) = Resolve(d, Mid$(code, 2))
        Case Else: Err.Raise 5, , "No such argument: " & code
    End Select
End Sub

Private Function Resolve(d As Document, path As String) As Object
    Dim steps() As String, i As Long, name As String, inside As String, at As Long
    Dim args As Variant, codes() As String, j As Long, here As Object, got As Variant
    Set here = d
    If Len(path) > 0 Then
        steps = Split(path, "/")
        For i = 0 To UBound(steps)
            at = InStr(steps(i), "(")
            If at > 0 Then
                name = Left$(steps(i), at - 1)
                inside = Mid$(steps(i), at + 1, Len(steps(i)) - at - 1)
                codes = Split(inside, ";")
                ReDim args(UBound(codes))
                For j = 0 To UBound(codes)
                    Arg d, CStr(codes(j)), args, j
                Next j
                got = Invoke(here, name, VbMethod, args)
            Else
                got = Invoke(here, CStr(steps(i)), VbGet, Array())
            End If
            Set here = got(0)
        Next i
    End If
    Set Resolve = here
End Function

' request: document, name, type (1 paragraph, 2 character), base style (its
' local name, or nothing for Normal or Default Paragraph Font). A new style
' takes the formatting at the user's cursor, so it is given its base style's
' font and paragraph format, as tools/formatting.py does on Windows.
' answer: the new style's local name.
Private Function AddStyle(r As Variant) As String
    Dim d As Document, s As Style, base As Style, kind As Long, stage As String
    On Error GoTo Failed
    stage = "find the document"
    Set d = FindDocument(Decode(CStr(r(0))))
    kind = CLng(r(2))
    stage = "find the base style"
    If Len(r(3)) > 0 Then
        Set base = d.Styles(Decode(CStr(r(3))))
    ElseIf kind = 1 Then
        Set base = d.Styles(wdStyleNormal)
    Else
        Set base = d.Styles(wdStyleDefaultParagraphFont)
    End If
    stage = "add the style"
    Set s = d.Styles.Add(Decode(CStr(r(1))), kind)
    stage = "set its base style"
    s.BaseStyle = base.NameLocal
    stage = "give it its base style's font"
    s.Font = base.Font.Duplicate
    If kind = 1 Then
        stage = "give it its base style's paragraph format"
        s.ParagraphFormat = base.ParagraphFormat.Duplicate
    End If
    AddStyle = "ok" & vbLf & Encode(s.NameLocal)
    Exit Function
Failed:
    Err.Raise Err.Number, , Err.Description & " (AddStyle: " & stage & ")"
End Function

' request: document, style (local name), the style whose list it joins (its
' local name; nothing to take the style's numbering off), the level. Word for
' Mac's `link to list template` is refused (-1708, measured). answer: the
' style's list level after it, 0 when it has no list.
Private Function StyleList(r As Variant) As String
    Dim d As Document, s As Style, lt As ListTemplate
    Set d = FindDocument(Decode(CStr(r(0))))
    Set s = d.Styles(Decode(CStr(r(1))))
    If Len(r(2)) = 0 Then
        s.LinkToListTemplate Nothing
    Else
        Set lt = d.Styles(Decode(CStr(r(2)))).ListTemplate
        If lt Is Nothing Then Err.Raise 5, , Decode(CStr(r(2))) & " has no numbering to join."
        s.LinkToListTemplate lt, CLng(r(3))
    End If
    If s.ListTemplate Is Nothing Then
        StyleList = "ok" & vbLf & "0"
    Else
        StyleList = "ok" & vbLf & s.ListLevelNumber
    End If
End Function

' request: document, the group's name (nothing: Word's), then a line a shape's
' name. Word for Mac's scripting has no grouping. The shapes go to
' Shapes.Range by their places, in an array of Variants, which both platforms
' take. An empty line is no shape: on the Mac the request ends with one.
' answer: the group's name.
Private Function GroupShapes(r As Variant) As String
    Dim d As Document, places() As Variant, i As Long, j As Long, n As Long
    Dim g As Shape, wanted As String
    Set d = FindDocument(Decode(CStr(r(0))))
    ReDim places(UBound(r))
    n = 0
    For i = 2 To UBound(r)
        If Len(r(i)) > 0 Then
            wanted = Decode(CStr(r(i)))
            For j = 1 To d.Shapes.Count
                If d.Shapes(j).Name = wanted Then places(n) = j
            Next j
            If IsEmpty(places(n)) Then Err.Raise 5, , "There is no shape named " & wanted & "."
            n = n + 1
        End If
    Next i
    ReDim Preserve places(n - 1)
    Set g = d.Shapes.Range(places).Group
    If Len(r(1)) > 0 Then g.Name = Decode(CStr(r(1)))
    GroupShapes = "ok" & vbLf & Encode(g.Name)
End Function

' request: document, the group's name. answer: a line a shape it held. The
' names are read before: the ShapeRange Ungroup answers refuses to be read
' (measured on Windows); the shapes keep the names they had in the group.
Private Function UngroupShape(r As Variant) As String
    Dim d As Document, g As Shape, i As Long, out As String
    Set d = FindDocument(Decode(CStr(r(0))))
    For i = 1 To d.Shapes.Count
        If d.Shapes(i).Name = Decode(CStr(r(1))) Then Set g = d.Shapes(i)
    Next i
    If g Is Nothing Then Err.Raise 5, , "There is no shape named " & Decode(CStr(r(1))) & "."
    out = "ok"
    For i = 1 To g.GroupItems.Count
        out = out & vbLf & Encode(g.GroupItems(i).Name)
    Next i
    g.Ungroup
    UngroupShape = out
End Function

' Where the document's tracked deletions are, in order of their starts.
Private Sub Deletions(d As Document, starts() As Long, ends() As Long, n As Long)
    Dim rv As Revision, i As Long, j As Long, s As Long, e As Long
    n = 0
    ReDim starts(0)
    ReDim ends(0)
    If d.Revisions.Count = 0 Then Exit Sub
    ReDim starts(d.Revisions.Count - 1)
    ReDim ends(d.Revisions.Count - 1)
    For Each rv In d.Revisions
        If rv.Type = 2 Or rv.Type = 14 Then ' wdRevisionDelete, wdRevisionMovedFrom
            starts(n) = rv.Range.Start
            ends(n) = rv.Range.End
            n = n + 1
        End If
    Next rv
    For i = 1 To n - 1 ' by their starts
        s = starts(i)
        e = ends(i)
        j = i - 1
        Do While j >= 0
            If starts(j) <= s Then Exit Do
            starts(j + 1) = starts(j)
            ends(j + 1) = ends(j)
            j = j - 1
        Loop
        starts(j + 1) = s
        ends(j + 1) = e
    Next i
End Sub

' Whether [s, e) lies within the deletions taken together: two side by side
' cover what they cover between them (tools/_word.py, inside_spans).
Private Function Deleted(s As Long, e As Long, starts() As Long, ends() As Long, n As Long) As Boolean
    Dim covered As Long, i As Long
    covered = s
    For i = 0 To n - 1
        If ends(i) > covered Then
            If starts(i) > covered Then Exit Function
            covered = ends(i)
            If covered >= e Then
                Deleted = True
                Exit Function
            End If
        End If
    Next i
End Function

' word_find_replace, as tools/text.py does it over COM, inside Word. request:
' document, what to find (a caret already doubled outside wildcards), the
' new text, replace (1 or 0), match case, whole word, wildcards, how many
' matches get a paragraph number, seconds to search, seconds in all.
' Matches in tracked deletions are left out. They are replaced from the last
' back, so the ones before keep their places, and each is read back. Inside
' one undo record unless one is open already (the server's, on Windows).
' answer: searched all (1 or 0), replaced, read back differently, out of
' time (1 or 0), then a line a match: start,end,paragraph (0 past the listed)
Private Function FindReplace(r As Variant) As String
    Dim d As Document, rg As Range, began As Single, newText As String
    Dim listed As Long, searchLimit As Single, workLimit As Single
    Dim starts() As Long, ends() As Long, paras() As Long, n As Long, reached As Long
    Dim delStarts() As Long, delEnds() As Long, nDel As Long, i As Long
    Dim searchedAll As Boolean, done As Long, missed As Long, ranOut As Boolean
    Dim recording As Boolean, lines() As String
    began = Timer
    Set d = FindDocument(Decode(CStr(r(0))))
    newText = Decode(CStr(r(2)))
    listed = CLng(r(7))
    searchLimit = Val(r(8))
    workLimit = Val(r(9))
    Deletions d, delStarts, delEnds, nDel
    ReDim starts(255)
    ReDim ends(255)
    ReDim paras(255)
    searchedAll = True
    reached = -1
    Set rg = d.Content
    rg.Find.ClearFormatting
    Do While rg.Find.Execute(FindText:=Decode(CStr(r(1))), MatchCase:=(r(4) = "1"), _
            MatchWholeWord:=(r(5) = "1"), MatchWildcards:=(r(6) = "1"), Forward:=True, _
            Wrap:=0, Format:=False)
        If rg.End <= reached Then
            Err.Raise 5, , "Word's search stopped moving forward at position " & rg.End _
                & ", after " & n & " matches. Nothing was changed. Try a narrower search."
        End If
        reached = rg.End
        If Not Deleted(rg.Start, rg.End, delStarts, delEnds, nDel) Then
            If n > UBound(starts) Then
                ReDim Preserve starts(2 * n)
                ReDim Preserve ends(2 * n)
                ReDim Preserve paras(2 * n)
            End If
            starts(n) = rg.Start
            ends(n) = rg.End
            If n < listed Then paras(n) = d.Range(0, rg.Start + 1).Paragraphs.Count
            n = n + 1
        End If
        If Seconds(began) > searchLimit Then
            searchedAll = False
            Exit Do
        End If
        rg.Collapse 0 ' wdCollapseEnd
    Loop
    If r(3) = "1" And n > 0 Then
        recording = Application.UndoRecord.IsRecordingCustomRecord
        If Not recording Then Application.UndoRecord.StartCustomRecord "Replace"
        On Error GoTo Undone
        For i = n - 1 To 0 Step -1
            If Seconds(began) > workLimit Then
                ranOut = True
                Exit For
            End If
            d.Range(starts(i), ends(i)).Text = newText
            done = done + 1
            If d.Range(starts(i), starts(i) + Len(newText)).Text <> newText Then missed = missed + 1
        Next i
        On Error GoTo 0
        If Not recording Then Application.UndoRecord.EndCustomRecord
    End If
    ReDim lines(n)
    lines(0) = "ok" & vbLf & IIf(searchedAll, "1", "0") & vbLf & done & vbLf & missed & vbLf _
        & IIf(ranOut, "1", "0")
    For i = 0 To n - 1
        lines(i + 1) = starts(i) & "," & ends(i) & "," & paras(i)
    Next i
    FindReplace = Join(lines, vbLf)
    Exit Function
Undone:
    Dim number As Long, description As String
    number = Err.Number
    description = Err.Description
    If Not recording Then Application.UndoRecord.EndCustomRecord
    Err.Raise number, , description & " (after " & done & " of " & n & " matches were replaced)"
End Function

' request: document, comment number, text. A reply to a reply goes to the
' comment it answers, as on Windows. answer: the reply's number, the count.
Private Function Reply(r As Variant) As String
    Dim d As Document, c As Comment, made As Comment, i As Long
    Set d = FindDocument(Decode(CStr(r(0))))
    Set c = d.Comments(CLng(r(1)))
    Do While Not AncestorOf(c) Is Nothing
        Set c = AncestorOf(c)
    Loop
    Set made = c.Replies.Add(c.Scope, Decode(CStr(r(2))))
    For i = 1 To d.Comments.Count
        If d.Comments(i).Range.Start = made.Range.Start Then Exit For
    Next i
    Reply = "ok" & vbLf & i & vbLf & d.Comments.Count
End Function

' request: document, comment number, 1 or 0
Private Function SetDone(r As Variant) As String
    Dim d As Document
    Set d = FindDocument(Decode(CStr(r(0))))
    d.Comments(CLng(r(1))).Done = (r(2) = "1")
    SetDone = "ok" & vbLf & IIf(d.Comments(CLng(r(1))).Done, "1", "0")
End Function

' request: document. answer: a line a comment, in Word's order: done (1 or
' 0),where the comment it replies to starts in the comments (-1: none),where
' its own text starts there. A reply's scope is its parent's, so comments are
' told apart by their own text.
Private Function CommentState(r As Variant) As String
    Dim d As Document, i As Long, out As String, c As Comment, a As Comment, at As Long
    Set d = FindDocument(Decode(CStr(r(0))))
    out = "ok"
    For i = 1 To d.Comments.Count
        Set c = d.Comments(i)
        Set a = AncestorOf(c)
        at = -1
        If Not a Is Nothing Then at = a.Range.Start
        out = out & vbLf & IIf(c.Done, "1", "0") & "," & at & "," & c.Range.Start
    Next i
    CommentState = out
End Function
