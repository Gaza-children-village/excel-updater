Attribute VB_Name = "UpdJson"
Option Explicit

' ============================================================================
' GCV Excel updater - minimal JSON reader.
' Objects -> Scripting.Dictionary, arrays -> Collection, strings -> String,
' numbers -> Double, true/false -> Boolean, null -> Null.
' ============================================================================

Private Const JSON_ERR As Long = vbObjectError + 2100

' Parses JSON text into result (use this: the result may be an object or a value).
Public Sub JsonParseInto(ByVal text As String, ByRef result As Variant)
    Dim pos As Long

    pos = 1
    ParseValue text, pos, result
    SkipWhitespace text, pos
    If pos <= Len(text) Then RaiseJson "Unexpected text after the JSON value", pos
End Sub

' Convenience for text that must be an object or an array.
Public Function JsonParse(ByVal text As String) As Object
    Dim value As Variant

    JsonParseInto text, value
    If Not IsObject(value) Then RaiseJson "Expected an object or an array", 1
    Set JsonParse = value
End Function

' ---- accessors (missing keys and null give "", False or Nothing) -----------

Public Function JsonStr(ByVal obj As Object, ByVal key As String) As String
    If obj Is Nothing Then Exit Function
    If TypeName(obj) <> "Dictionary" Then Exit Function
    If Not obj.Exists(key) Then Exit Function
    If IsObject(obj(key)) Then Exit Function
    If IsNull(obj(key)) Then Exit Function
    JsonStr = CStr(obj(key))
End Function

Public Function JsonBool(ByVal obj As Object, ByVal key As String) As Boolean
    If obj Is Nothing Then Exit Function
    If TypeName(obj) <> "Dictionary" Then Exit Function
    If Not obj.Exists(key) Then Exit Function
    If IsObject(obj(key)) Then Exit Function
    If VarType(obj(key)) = vbBoolean Then JsonBool = obj(key)
End Function

Public Function JsonObj(ByVal obj As Object, ByVal key As String) As Object
    If obj Is Nothing Then Exit Function
    If TypeName(obj) <> "Dictionary" Then Exit Function
    If Not obj.Exists(key) Then Exit Function
    If IsObject(obj(key)) Then Set JsonObj = obj(key)
End Function

' ---- parser -----------------------------------------------------------------

Private Sub ParseValue(ByRef text As String, ByRef pos As Long, ByRef result As Variant)
    SkipWhitespace text, pos
    If pos > Len(text) Then RaiseJson "Unexpected end of JSON", pos

    Select Case Mid$(text, pos, 1)
        Case "{"
            Set result = ParseObject(text, pos)
        Case "["
            Set result = ParseArray(text, pos)
        Case """"
            result = ParseString(text, pos)
        Case "t"
            ExpectWord text, pos, "true"
            result = True
        Case "f"
            ExpectWord text, pos, "false"
            result = False
        Case "n"
            ExpectWord text, pos, "null"
            result = Null
        Case "-", "0" To "9"
            result = ParseNumber(text, pos)
        Case Else
            RaiseJson "Unexpected character '" & Mid$(text, pos, 1) & "'", pos
    End Select
End Sub

Private Function ParseObject(ByRef text As String, ByRef pos As Long) As Object
    Dim dict As Object
    Dim key As String
    Dim value As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    pos = pos + 1
    SkipWhitespace text, pos
    If Mid$(text, pos, 1) = "}" Then
        pos = pos + 1
        Set ParseObject = dict
        Exit Function
    End If

    Do
        SkipWhitespace text, pos
        If Mid$(text, pos, 1) <> """" Then RaiseJson "Expected a key", pos
        key = ParseString(text, pos)
        SkipWhitespace text, pos
        If Mid$(text, pos, 1) <> ":" Then RaiseJson "Expected ':'", pos
        pos = pos + 1
        value = Empty
        ParseValue text, pos, value
        If IsObject(value) Then
            Set dict(key) = value
        Else
            dict(key) = value
        End If
        SkipWhitespace text, pos
        Select Case Mid$(text, pos, 1)
            Case ","
                pos = pos + 1
            Case "}"
                pos = pos + 1
                Exit Do
            Case Else
                RaiseJson "Expected ',' or '}'", pos
        End Select
    Loop
    Set ParseObject = dict
End Function

Private Function ParseArray(ByRef text As String, ByRef pos As Long) As Collection
    Dim items As Collection
    Dim value As Variant

    Set items = New Collection
    pos = pos + 1
    SkipWhitespace text, pos
    If Mid$(text, pos, 1) = "]" Then
        pos = pos + 1
        Set ParseArray = items
        Exit Function
    End If

    Do
        value = Empty
        ParseValue text, pos, value
        items.Add value
        SkipWhitespace text, pos
        Select Case Mid$(text, pos, 1)
            Case ","
                pos = pos + 1
            Case "]"
                pos = pos + 1
                Exit Do
            Case Else
                RaiseJson "Expected ',' or ']'", pos
        End Select
    Loop
    Set ParseArray = items
End Function

Private Function ParseString(ByRef text As String, ByRef pos As Long) As String
    Dim result As String
    Dim quotePos As Long
    Dim slashPos As Long

    pos = pos + 1
    Do
        quotePos = InStr(pos, text, """")
        If quotePos = 0 Then RaiseJson "Unterminated string", pos
        slashPos = InStr(pos, text, "\")
        If slashPos = 0 Or slashPos > quotePos Then
            ParseString = result & Mid$(text, pos, quotePos - pos)
            pos = quotePos + 1
            Exit Function
        End If

        result = result & Mid$(text, pos, slashPos - pos)
        Select Case Mid$(text, slashPos + 1, 1)
            Case """", "\", "/"
                result = result & Mid$(text, slashPos + 1, 1)
                pos = slashPos + 2
            Case "b"
                result = result & Chr$(8)
                pos = slashPos + 2
            Case "f"
                result = result & Chr$(12)
                pos = slashPos + 2
            Case "n"
                result = result & vbLf
                pos = slashPos + 2
            Case "r"
                result = result & vbCr
                pos = slashPos + 2
            Case "t"
                result = result & vbTab
                pos = slashPos + 2
            Case "u"
                If Not Mid$(text, slashPos + 2, 4) Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then _
                    RaiseJson "Bad \u escape", slashPos
                ' Val("&Hxxxx&") gives the unsigned value (0..65535) on every PC.
                result = result & ChrW$(Val("&H" & Mid$(text, slashPos + 2, 4) & "&"))
                pos = slashPos + 6
            Case Else
                RaiseJson "Bad escape", slashPos
        End Select
    Loop
End Function

Private Function ParseNumber(ByRef text As String, ByRef pos As Long) As Double
    Dim start As Long

    start = pos
    Do While pos <= Len(text)
        If InStr("+-0123456789.eE", Mid$(text, pos, 1)) = 0 Then Exit Do
        pos = pos + 1
    Loop
    ' Val always uses "." as the decimal separator, whatever the PC's region.
    ParseNumber = Val(Mid$(text, start, pos - start))
End Function

Private Sub ExpectWord(ByRef text As String, ByRef pos As Long, ByVal word As String)
    If Mid$(text, pos, Len(word)) <> word Then RaiseJson "Expected '" & word & "'", pos
    pos = pos + Len(word)
End Sub

Private Sub SkipWhitespace(ByRef text As String, ByRef pos As Long)
    Do While pos <= Len(text)
        Select Case Mid$(text, pos, 1)
            Case " ", vbTab, vbCr, vbLf
                pos = pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Sub RaiseJson(ByVal message As String, ByVal pos As Long)
    Err.Raise JSON_ERR, "UpdJson", message & " (position " & pos & ")."
End Sub
