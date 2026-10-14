Attribute VB_Name = "modJson"
Option Explicit

Public Type ToolCallItem
    id As String
    toolName As String
    argumentsJson As String
End Type

Public Function JsonEscape(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case """": out = out & "\"""
            Case "\": out = out & "\\"
            Case vbCr: out = out & "\r"
            Case vbLf: out = out & "\n"
            Case vbTab: out = out & "\t"
            Case Else
                ' AscW returns a signed Integer: chars >= U+8000 (CJK) come
                ' back negative and must be normalized before the <32 check,
                ' otherwise kanji get wrongly \u-escaped.
                Dim code As Long
                code = AscW(ch)
                If code < 0 Then code = code + 65536
                If code < 32 Then
                    out = out & "\u" & Right$("0000" & Hex(code), 4)
                Else
                    out = out & ch
                End If
        End Select
    Next i
    JsonEscape = out
End Function

Public Function JsonString(ByVal s As String) As String
    JsonString = """" & JsonEscape(s) & """"
End Function

Public Function JsonExtractString(ByVal json As String, ByVal key As String) As String
    Dim keyPos As Long
    Dim i As Long
    Dim ch As String
    Dim escaped As Boolean
    Dim buf As String

    keyPos = InStr(1, json, """" & key & """", vbTextCompare)
    Do While keyPos > 0
        i = keyPos + Len(key) + 2
        JsonSkipWhitespace json, i
        If i <= Len(json) And Mid$(json, i, 1) = ":" Then Exit Do
        keyPos = InStr(keyPos + 1, json, """" & key & """", vbTextCompare)
    Loop
    If keyPos = 0 Then Exit Function

    i = i + 1
    JsonSkipWhitespace json, i
    If i > Len(json) Then Exit Function
    If Mid$(json, i, 1) <> """" Then
        ' Accept bare JSON literals (models emit "recursive":true unquoted)
        If LCase$(Left$(Mid$(json, i, 4), 4)) = "true" Then
            JsonExtractString = "true"
        ElseIf LCase$(Left$(Mid$(json, i, 5), 5)) = "false" Then
            JsonExtractString = "false"
        End If
        Exit Function
    End If

    i = i + 1
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        If escaped Then
            Select Case ch
                Case "n": buf = buf & vbLf
                Case "r": buf = buf & vbCr
                Case "t": buf = buf & vbTab
                Case """": buf = buf & """"
                Case "\": buf = buf & "\"
                Case "u"
                    If i + 4 <= Len(json) Then
                        On Error Resume Next
                        buf = buf & ChrW$(CLng("&H" & Mid$(json, i + 1, 4)))
                        If Err.Number <> 0 Then
                            Err.Clear
                            buf = buf & "u"
                        Else
                            i = i + 4
                        End If
                        On Error GoTo 0
                    Else
                        buf = buf & "u"
                    End If
                Case Else: buf = buf & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = """" Then
            JsonExtractString = buf
            Exit Function
        Else
            buf = buf & ch
        End If
        i = i + 1
    Loop
End Function

Private Sub JsonSkipWhitespace(ByVal json As String, ByRef pos As Long)
    Do While pos <= Len(json) And Mid$(json, pos, 1) <= " "
        pos = pos + 1
    Loop
End Sub

Public Function JsonExtractFirstContent(ByVal json As String) As String
    JsonExtractFirstContent = JsonExtractString(json, "content")
End Function

Public Function JsonParseToolCalls(ByVal json As String, ByRef calls() As ToolCallItem) As Long
    Dim count As Long
    Dim arrayStart As Long
    Dim arrayEnd As Long
    Dim pos As Long
    Dim blockStart As Long
    Dim blockEnd As Long
    Dim block As String
    Dim idVal As String
    Dim nameVal As String
    Dim argsVal As String

    Erase calls
    arrayStart = InStr(1, json, """tool_calls""", vbTextCompare)
    If arrayStart = 0 Then Exit Function
    arrayStart = InStr(arrayStart, json, "[")
    If arrayStart = 0 Then Exit Function
    arrayEnd = JsonFindClosingDelimiter(json, arrayStart, "[", "]")
    If arrayEnd = 0 Then Exit Function

    pos = arrayStart + 1
    Do While pos < arrayEnd
        blockStart = InStr(pos, json, "{")
        If blockStart = 0 Or blockStart >= arrayEnd Then Exit Do
        blockEnd = JsonFindClosingDelimiter(json, blockStart, "{", "}")
        If blockEnd = 0 Or blockEnd > arrayEnd Then Exit Do

        block = Mid$(json, blockStart, blockEnd - blockStart + 1)
        If InStr(1, block, """function""", vbTextCompare) > 0 Then
            idVal = JsonExtractString(block, "id")
            nameVal = JsonExtractString(block, "name")
            argsVal = JsonExtractString(block, "arguments")
            If Len(nameVal) > 0 Then
                count = count + 1
                If count = 1 Then
                    ReDim calls(1 To 1)
                Else
                    ReDim Preserve calls(1 To count)
                End If
                calls(count).id = idVal
                calls(count).toolName = nameVal
                calls(count).argumentsJson = argsVal
            End If
        End If
        pos = blockEnd + 1
    Loop

    JsonParseToolCalls = count
End Function

Public Function JsonFindClosingDelimiter(ByVal json As String, ByVal startPos As Long, ByVal openChar As String, ByVal closeChar As String) As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim inString As Boolean
    Dim escaped As Boolean

    For i = startPos To Len(json)
        ch = Mid$(json, i, 1)
        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
            End If
        ElseIf ch = """" Then
            inString = True
        ElseIf ch = openChar Then
            depth = depth + 1
        ElseIf ch = closeChar Then
            depth = depth - 1
            If depth = 0 Then
                JsonFindClosingDelimiter = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Function JsonExtractObject(ByVal json As String, ByVal key As String) As String
    Dim keyPos As Long
    Dim i As Long
    Dim endPos As Long

    keyPos = InStr(1, json, """" & key & """", vbTextCompare)
    If keyPos = 0 Then Exit Function
    i = keyPos + Len(key) + 2
    JsonSkipWhitespace json, i
    If i > Len(json) Or Mid$(json, i, 1) <> ":" Then Exit Function
    i = i + 1
    JsonSkipWhitespace json, i
    If i > Len(json) Or Mid$(json, i, 1) <> "{" Then Exit Function
    endPos = JsonFindClosingDelimiter(json, i, "{", "}")
    If endPos > i Then JsonExtractObject = Mid$(json, i, endPos - i + 1)
End Function

Public Function JsonGetArg(ByVal argsJson As String, ByVal key As String) As String
    JsonGetArg = JsonExtractString(argsJson, key)
End Function

' Removes "key":<value> from an object string for any value type.
' Providers like the AMD sglang router reject echoed assistant messages whose
' reasoning_content is null, so drop it (and it wastes context tokens anyway).
Public Function JsonRemoveKey(ByVal json As String, ByVal key As String) As String
    Dim keyPos As Long
    Dim i As Long
    Dim c As String
    Dim valEnd As Long
    Dim cutStart As Long
    Dim cutEnd As Long

    JsonRemoveKey = json
    On Error Resume Next
    keyPos = InStr(1, json, """" & key & """", vbTextCompare)
    If keyPos = 0 Then Exit Function

    i = keyPos + Len(key) + 2
    Do While i <= Len(json) And Mid$(json, i, 1) <= " "
        i = i + 1
    Loop
    If i > Len(json) Or Mid$(json, i, 1) <> ":" Then Exit Function
    i = i + 1
    Do While i <= Len(json) And Mid$(json, i, 1) <= " "
        i = i + 1
    Loop
    c = Mid$(json, i, 1)

    If c = """" Then
        Dim esc As Boolean
        Dim j As Long
        esc = False
        For j = i + 1 To Len(json)
            c = Mid$(json, j, 1)
            If esc Then
                esc = False
            ElseIf c = "\" Then
                esc = True
            ElseIf c = """" Then
                Exit For
            End If
        Next j
        valEnd = j
    ElseIf c = "{" Or c = "[" Then
        valEnd = JsonFindClosingDelimiter(json, i, c, IIf(c = "{", "}", "]"))
        If valEnd = 0 Then Exit Function
    Else
        valEnd = i - 1
        Do While valEnd + 1 <= Len(json)
            c = Mid$(json, valEnd + 1, 1)
            If c = "," Or c = "}" Or c = "]" Or c = " " Or c = vbCr Or c = vbLf Then Exit Do
            valEnd = valEnd + 1
        Loop
    End If

    cutStart = keyPos
    cutEnd = valEnd
    If Mid$(json, cutEnd + 1, 1) = "," Then
        cutEnd = cutEnd + 1
    Else
        ' last member of an object: also swallow the comma before the key
        Dim b As Long
        b = keyPos - 1
        Do While b >= 1 And Mid$(json, b, 1) = " "
            b = b - 1
        Loop
        If Mid$(json, b, 1) = "," Then cutStart = b
    End If
    JsonRemoveKey = Left$(json, cutStart - 1) & Mid$(json, cutEnd + 1)
End Function

' Drops "key":null entirely (bare null only) — some strict routers 400 on nulls.
Public Function JsonRemoveNullKey(ByVal json As String, ByVal key As String) As String
    Dim keyPos As Long
    Dim i As Long
    JsonRemoveNullKey = json
    keyPos = InStr(1, json, """" & key & """", vbTextCompare)
    If keyPos = 0 Then Exit Function
    i = keyPos + Len(key) + 2
    Do While i <= Len(json) And Mid$(json, i, 1) <= " "
        i = i + 1
    Loop
    If i > Len(json) Or Mid$(json, i, 1) <> ":" Then Exit Function
    i = i + 1
    Do While i <= Len(json) And Mid$(json, i, 1) <= " "
        i = i + 1
    Loop
    If LCase$(Left$(Mid$(json, i, 4), 4)) = "null" Then JsonRemoveNullKey = JsonRemoveKey(json, key)
End Function

' True only when "tool_calls" is present with an ARRAY value. sglang-style
' routers return "tool_calls":null on plain text answers; a bare substring
' match there made the agent think every answer was a tool call and loop.
Public Function JsonHasToolCalls(ByVal json As String) As Boolean
    Dim pos As Long
    Dim i As Long
    pos = 1
    Do
        pos = InStr(pos, json, """tool_calls""", vbTextCompare)
        If pos = 0 Then Exit Function
        i = pos + 12
        Do While i <= Len(json) And Mid$(json, i, 1) <= " "
            i = i + 1
        Loop
        If i <= Len(json) And Mid$(json, i, 1) = ":" Then
            i = i + 1
            Do While i <= Len(json) And Mid$(json, i, 1) <= " "
                i = i + 1
            Loop
            If i <= Len(json) And Mid$(json, i, 1) = "[" Then
                JsonHasToolCalls = True
                Exit Function
            End If
        End If
        pos = pos + 12
    Loop
End Function

Public Function JsonBuildToolResult(ByVal toolCallId As String, ByVal content As String) As String
    JsonBuildToolResult = "{""role"":""tool"",""tool_call_id"":" & JsonString(toolCallId) & _
        ",""content"":" & JsonString(content) & "}"
End Function

Public Function JsonBuildUserMessage(ByVal content As String) As String
    JsonBuildUserMessage = "{""role"":""user"",""content"":" & JsonString(content) & "}"
End Function

Public Function JsonBuildAssistantToolCallsMessage(ByVal rawAssistantJson As String) As String
    JsonBuildAssistantToolCallsMessage = rawAssistantJson
End Function
