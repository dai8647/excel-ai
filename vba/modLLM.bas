Attribute VB_Name = "modLLM"
Option Explicit

Private Const API_TIMEOUT_MS As Long = 15000

Public Function LLM_TestConnection() As String
    Dim body As String
    Dim resp As String

    On Error GoTo ErrHandler

    If IsAnthropicProvider() Then
        body = "{""model"":" & JsonString(GetModel()) & ",""max_tokens"":5,""stream"":false" & AnthropicThinkingOption() & ",""messages"":[{""role"":""user"",""content"":""ping""}]}"
        resp = HttpPostJson(AnthropicApiRoot() & "/messages", body, True)
    Else
        body = "{""model"":" & JsonString(GetModel()) & ",""messages"":[{""role"":""user"",""content"":""ping""}],""max_tokens"":5}"
        resp = HttpPostJson(GetBaseUrl() & "/chat/completions", body, False)
    End If
    LLM_TestConnection = "接続OK" & vbCrLf & "方式: " & GetProvider() & vbCrLf & "モデル: " & GetModel()
    Exit Function
ErrHandler:
    LLM_TestConnection = "接続エラー: " & Err.Description
End Function

Public Function LLM_ChatWithTools(ByVal history As Collection) As String
    Dim body As String
    If IsAnthropicProvider() Then
        body = BuildAnthropicRequest(history)
        LLM_ChatWithTools = NormalizeAnthropicResponse(HttpPostJson(AnthropicApiRoot() & "/messages", body, True))
    Else
        body = "{""model"":" & JsonString(GetModel()) & ",""messages"":" & BuildMessagesArray(history) & _
            ",""tools"":" & Tools_GetDefinitionsJson() & ",""tool_choice"":""auto""}"
        LLM_ChatWithTools = HttpPostJson(GetBaseUrl() & "/chat/completions", body, False)
    End If
End Function

Public Sub LLM_RefreshModelList()
    Dim json As String
    Dim ws As Worksheet
    Dim pos As Long
    Dim modelId As String
    Dim rowNum As Long

    On Error GoTo ErrHandler
    If IsAnthropicProvider() Then
        json = HttpGetJson(AnthropicApiRoot() & "/models", True)
    Else
        json = HttpGetJson(GetBaseUrl() & "/models", False)
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_MODELS)
    ws.Rows("6:" & ws.Rows.Count).ClearContents
    rowNum = 6
    pos = 1
    Do
        pos = InStr(pos, json, """id""", vbTextCompare)
        If pos = 0 Then Exit Do
        modelId = JsonExtractString(Mid$(json, pos), "id")
        If Len(modelId) > 0 Then
            ws.Cells(rowNum, 1).Value = modelId
            ws.Cells(rowNum, 2).Value = GetProvider()
            rowNum = rowNum + 1
        End If
        pos = pos + 4
    Loop
    If rowNum = 6 Then Err.Raise vbObjectError + 2003, "Models", "モデルが取得できませんでした"
    Ops_WriteStatus (rowNum - 6) & " 件のモデルを更新しました", False
    MsgBox (rowNum - 6) & " 件のモデルを取得しました。設定シートのモデル欄から選択してください。", vbInformation, "モデル一覧"
    Exit Sub
ErrHandler:
    Ops_WriteStatus "モデル取得失敗: " & Err.Description, True
    MsgBox "モデル一覧の取得に失敗しました: " & Err.Description, vbCritical, "モデル一覧"
End Sub

Public Function LLM_ExtractAssistantMessage(ByVal responseJson As String) As String
    LLM_ExtractAssistantMessage = JsonExtractFirstContent(responseJson)
End Function

Public Function LLM_ExtractAssistantRawMessageObject(ByVal responseJson As String) As String
    Dim marker As String, startPos As Long, endPos As Long
    marker = """message"":"
    startPos = InStr(1, responseJson, marker, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos, responseJson, "{")
    If startPos = 0 Then Exit Function
    endPos = JsonFindClosingDelimiter(responseJson, startPos, "{", "}")
    If endPos > startPos Then LLM_ExtractAssistantRawMessageObject = Mid$(responseJson, startPos, endPos - startPos + 1)
End Function

Private Function HttpPostJson(ByVal url As String, ByVal jsonBody As String, ByVal useAnthropicHeaders As Boolean) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.SetTimeouts API_TIMEOUT_MS, API_TIMEOUT_MS, API_TIMEOUT_MS, API_TIMEOUT_MS
    ApplyHeaders http, useAnthropicHeaders
    http.SetRequestHeader "Connection", "close"
    http.Send jsonBody
    If http.Status >= 400 Then Err.Raise vbObjectError + 2001, "LLM", "HTTP " & http.Status & ": " & http.ResponseText
    HttpPostJson = http.ResponseText
End Function

Private Function HttpGetJson(ByVal url As String, ByVal useAnthropicHeaders As Boolean) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.SetTimeouts API_TIMEOUT_MS, API_TIMEOUT_MS, API_TIMEOUT_MS, API_TIMEOUT_MS
    ApplyHeaders http, useAnthropicHeaders
    http.SetRequestHeader "Connection", "close"
    http.Send
    If http.Status >= 400 Then Err.Raise vbObjectError + 2002, "LLM", "HTTP " & http.Status & ": " & http.ResponseText
    HttpGetJson = http.ResponseText
End Function

Private Sub ApplyHeaders(ByVal http As Object, ByVal useAnthropicHeaders As Boolean)
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    If useAnthropicHeaders Then
        http.SetRequestHeader "anthropic-version", GetAnthropicVersion()
        If Len(GetApiKey()) > 0 Then http.SetRequestHeader "x-api-key", GetApiKey()
    ElseIf Len(GetApiKey()) > 0 Then
        http.SetRequestHeader "Authorization", "Bearer " & GetApiKey()
    End If
End Sub

Private Function AnthropicApiRoot() As String
    Dim root As String
    root = GetBaseUrl()
    If LCase$(Right$(root, 3)) = "/v1" Then root = Left$(root, Len(root) - 3)
    AnthropicApiRoot = root & "/v1"
End Function

Public Function BuildMessagesArray(ByVal history As Collection) As String
    Dim i As Long, parts As String
    For i = 1 To history.Count
        If i > 1 Then parts = parts & ","
        parts = parts & CStr(history(i))
    Next i
    BuildMessagesArray = "[" & parts & "]"
End Function

Private Function BuildAnthropicRequest(ByVal history As Collection) As String
    Dim i As Long, item As String, role As String, content As String, parts As String, systemText As String
    Dim calls() As ToolCallItem, callCount As Long, j As Long
    For i = 1 To history.Count
        item = CStr(history(i))
        role = JsonExtractString(item, "role")
        content = JsonExtractString(item, "content")
        If role = "system" Then
            systemText = content
        ElseIf role = "tool" Then
            parts = AddJsonPart(parts, "{""role"":""user"",""content"":[{""type"":""tool_result"",""tool_use_id"":" & JsonString(JsonExtractString(item, "tool_call_id")) & ",""content"":" & JsonString(content) & "}]}")
        ElseIf role = "assistant" Then
            callCount = JsonParseToolCalls(item, calls)
            If callCount > 0 Then
                Dim toolParts As String
                For j = 1 To callCount
                    toolParts = AddJsonPart(toolParts, "{""type"":""tool_use"",""id"":" & JsonString(calls(j).id) & ",""name"":" & JsonString(calls(j).toolName) & ",""input"":" & calls(j).argumentsJson & "}")
                Next j
                parts = AddJsonPart(parts, "{""role"":""assistant"",""content"": [" & toolParts & "]}")
            Else
                parts = AddJsonPart(parts, "{""role"":""assistant"",""content"":" & JsonString(content) & "}")
            End If
        Else
            parts = AddJsonPart(parts, "{""role"":""user"",""content"":" & JsonString(content) & "}")
        End If
    Next i
    BuildAnthropicRequest = "{""model"":" & JsonString(GetModel()) & ",""max_tokens"":2048,""stream"":false" & AnthropicThinkingOption() & ",""system"":" & JsonString(systemText) & ",""messages"": [" & parts & "],""tools"":" & Tools_GetAnthropicDefinitionsJson() & "}"
End Function

Private Function AnthropicThinkingOption() As String
    Dim root As String
    root = LCase$(GetBaseUrl())
    If InStr(1, root, "127.0.0.1:20128", vbTextCompare) > 0 Or InStr(1, root, "localhost:20128", vbTextCompare) > 0 Then
        AnthropicThinkingOption = ",""thinking"":{""type"":""disabled""}"
    End If
End Function

Private Function NormalizeAnthropicResponse(ByVal responseJson As String) As String
    Dim pos As Long, blockStart As Long, blockEnd As Long, block As String, calls As String
    Dim id As String, name As String, inputJson As String
    pos = InStr(1, responseJson, """tool_use""", vbTextCompare)
    Do While pos > 0
        blockStart = InStrRev(responseJson, "{", pos)
        If blockStart = 0 Then Exit Do
        blockEnd = JsonFindClosingDelimiter(responseJson, blockStart, "{", "}")
        If blockEnd <= pos Or blockEnd <= blockStart Then Exit Do
        block = Mid$(responseJson, blockStart, blockEnd - blockStart + 1)
        id = JsonExtractString(block, "id")
        name = JsonExtractString(block, "name")
        inputJson = JsonExtractObject(block, "input")
        calls = AddJsonPart(calls, "{""id"":" & JsonString(id) & ",""type"":""function"",""function"":{""name"":" & JsonString(name) & ",""arguments"":" & JsonString(inputJson) & "}}")
        pos = InStr(blockEnd + 1, responseJson, """tool_use""", vbTextCompare)
    Loop
    If Len(calls) > 0 Then
        NormalizeAnthropicResponse = "{""choices"":[{""message"":{""role"":""assistant"",""content"":null,""tool_calls"":[" & calls & "]}}]}"
    Else
        NormalizeAnthropicResponse = "{""choices"":[{""message"":{""role"":""assistant"",""content"":" & JsonString(JsonExtractString(responseJson, "text")) & "}}]}"
    End If
End Function

Private Function AddJsonPart(ByVal parts As String, ByVal item As String) As String
    If Len(parts) > 0 Then parts = parts & ","
    AddJsonPart = parts & item
End Function

Public Function BuildSystemPrompt() As String
    BuildSystemPrompt = "You are ExcelAI, a file operations assistant. All file paths MUST be inside the workspace: " & GetWorkspacePath() & ". Use relative paths when possible. Never use .. or absolute paths outside workspace. Respond in Japanese. When a task is complete, summarize what was done."
End Function

Public Function BuildInitialMessages(ByVal userPrompt As String) As Collection
    Dim msgs As New Collection
    msgs.Add "{""role"":""system"",""content"":" & JsonString(BuildSystemPrompt()) & "}"
    msgs.Add JsonBuildUserMessage(userPrompt)
    Set BuildInitialMessages = msgs
End Function
