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
Public Const BRIDGE_VERSION As String = "1"
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
