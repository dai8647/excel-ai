Attribute VB_Name = "modAgentCore"
Option Explicit

Public g_StopRequested As Boolean
Public g_AgentRunning As Boolean

Public Sub Agent_RunPrompt(ByVal userPrompt As String, ByVal statusCallback As String)
    Dim messages As Collection
    Dim stepNum As Long
    Dim maxSteps As Long
    Dim responseJson As String
    Dim assistantContent As String
    Dim rawMsg As String
    Dim finishReason As String
    Dim calls() As ToolCallItem
    Dim callCount As Long
    Dim i As Long
    Dim toolResult As String
    Dim autoApproved As Boolean
    Dim contextResult As String
    Dim request As Collection
    Dim prior As Collection

    On Error GoTo ErrHandler

    InitDefaultSettings
    EnsureWorkspaceExists
    g_StopRequested = False
    g_AgentRunning = True
    ChatLog_Write "user", userPrompt
    Ops_WriteProgress "AIが依頼内容を整理しています...", False

    ' Multi-turn: system prompt plus the persisted conversation, then append
    ' this request. Everything remembered is also written to the file so a
    ' stop, step-limit or crash still leaves context for "続き".
    Set messages = New Collection
    messages.Add "{""role"":""system"",""content"":" & JsonString(BuildSystemPrompt()) & "}"
    Set prior = Session_LoadMessages()
    For i = 1 To prior.Count
        messages.Add prior(i)
    Next i
    Session_Remember JsonBuildUserMessage(userPrompt), messages

    maxSteps = GetMaxSteps()
    autoApproved = IsAutoRunEnabled()

    For stepNum = 1 To maxSteps
        If g_StopRequested Then
            ChatLog_Write "system", "ユーザーにより停止"
            Ops_WriteProgress "停止しました。履歴は保存済みなので「続き」で再開できます", False
            Agent_ReportStatus statusCallback, "停止しました"
            Agent_NotifyComplete statusCallback, "ユーザーの要求で停止しました。会話は保存されているので、続けて指示できます。"
            g_AgentRunning = False
            Exit Sub
        End If

        Agent_ReportStatus statusCallback, "AI思考中... (step " & stepNum & "/" & maxSteps & ")"
        DoEvents

        Set request = SessionTrimForModel(messages)
        responseJson = LLM_ChatWithTools(request)

        If JsonHasToolCalls(responseJson) Then
            rawMsg = LLM_ExtractAssistantRawMessageObject(responseJson)
            If Len(rawMsg) > 0 Then Session_Remember rawMsg, messages

            callCount = JsonParseToolCalls(responseJson, calls)
            For i = 1 To callCount
                Agent_ReportStatus statusCallback, "実行中: " & calls(i).toolName
                Ops_WriteProgress "実行中: " & calls(i).toolName & " (step " & stepNum & "/" & maxSteps & ")", False
                ChatLog_Write "progress", "step " & stepNum & ": " & calls(i).toolName
                ChatLog_Write "tool", calls(i).toolName & " " & calls(i).argumentsJson

                toolResult = Tools_Execute(calls(i).toolName, calls(i).argumentsJson, autoApproved)
                ChatLog_Write "tool_result", Left$(toolResult, 2000)
                Ops_WriteProgress "完了: " & calls(i).toolName, InStr(1, toolResult, "エラー", vbTextCompare) > 0

                If Len(calls(i).id) > 0 Then
                    contextResult = LimitToolResultForModel(toolResult)
                    Session_Remember JsonBuildToolResult(calls(i).id, contextResult), messages
                Else
                    Session_Remember JsonBuildUserMessage("[tool:" & calls(i).toolName & "] " & LimitToolResultForModel(toolResult)), messages
                End If
            Next i
            If callCount = 0 Then
                Session_Remember JsonBuildUserMessage("(ツール呼び出しの引数を解析できませんでした。ツール呼び出しをやり直してください。ファイル書き込みが途中で切れた場合は内容分割して append_file で行ってください。)"), messages
            End If
        Else
            assistantContent = LLM_ExtractAssistantMessage(responseJson)
            finishReason = JsonExtractString(responseJson, "finish_reason")
            If finishReason = "length" Then
                ' Output was cut at the token cap mid-answer: keep the partial
                ' reply, tell the model to continue, and loop on.
                If Len(assistantContent) > 0 Then _
                    Session_Remember "{""role"":""assistant"",""content"":" & JsonString(assistantContent) & "}", messages
                Session_Remember JsonBuildUserMessage("(応答が出力トークン上限で途切れました。同じ内容を最初から繰り返さず、続きを続けてください。ツール呼び出しが途切れていた場合は、より小さい単位で呼び直してください。)"), messages
                ChatLog_Write "system", "出力上限で途中切断 → 続きを要求 (step " & stepNum & ")"
            ElseIf Len(assistantContent) = 0 Then
                ChatLog_Write "assistant", "（応答を取得できませんでした）"
                Agent_ReportStatus statusCallback, "完了"
                Agent_NotifyComplete statusCallback, "（応答を取得できませんでした）"
                g_AgentRunning = False
                Exit Sub
            Else
                ChatLog_Write "assistant", assistantContent
                Agent_ReportStatus statusCallback, "完了"
                Agent_NotifyComplete statusCallback, assistantContent
                g_AgentRunning = False
                Exit Sub
            End If
        End If
    Next stepNum

    ChatLog_Write "error", "最大ステップ数 (" & maxSteps & ") に達しました"
    Ops_WriteStatus "最大ステップ数 (" & maxSteps & ") に達しました", True
    Agent_ReportStatus statusCallback, "ステップ上限に達しました"
    Agent_NotifyComplete statusCallback, "最大ステップ数に達したため停止しました。会話は保存されているので「続けて」で再開できます。"
    g_AgentRunning = False
    Exit Sub

ErrHandler:
    g_AgentRunning = False
    ChatLog_Write "error", "step=" & stepNum & ": " & Err.Description
    Ops_WriteStatus "step=" & stepNum & ": " & Err.Description, True
    Agent_ReportStatus statusCallback, "エラー"
    Agent_NotifyComplete statusCallback, "処理を完了できませんでした: " & Err.Description
End Sub

Public Sub Agent_RequestStop()
    If Not g_AgentRunning Then
        Ops_WriteProgress "実行中の処理はありません", False
        Exit Sub
    End If
    g_StopRequested = True
    Ops_WriteProgress "停止リクエスト受付: 現在のツール実行完了後に停止します", False
End Sub

Private Function LimitToolResultForModel(ByVal text As String) As String
    Const MAX_TOOL_CONTEXT As Long = 12000
    If Len(text) > MAX_TOOL_CONTEXT Then
        LimitToolResultForModel = Left$(text, MAX_TOOL_CONTEXT) & vbCrLf & "[tool result truncated for context]"
    Else
        LimitToolResultForModel = text
    End If
End Function

Public Sub RunSheetPrompt()
    Dim ws As Worksheet
    Dim prompt As String

    If g_AgentRunning Then
        Ops_WriteProgress "実行中です。完了を待つか「停止」を押してください", True
        Exit Sub
    End If

    Set ws = GetOpsSheet()
    prompt = Trim$(CStr(ws.Range("B9").Value))
    If Len(prompt) = 0 Then
        Ops_WriteProgress "依頼を入力してください", True
        Exit Sub
    End If

    ws.Range("B21").Value = ""    ' merged B21:H24: ClearContents raises 1004 "この操作は結合したセルには行えません"
    Ops_WriteProgress "依頼を送信しています...（履歴 " & Session_MessageCount() & " 件を継続）", False
    Agent_RunPrompt prompt, ""
    ws.Range("B21").Value = Left$(ChatLog_GetRecentAssistantText(), 30000)
    ws.Range("B21").Font.Color = RGB(0, 0, 0)
    Ops_WriteProgress "完了", False
End Sub

Private Sub Agent_ReportStatus(ByVal callback As String, ByVal msg As String)
    If Len(callback) = 0 Then Exit Sub
    On Error Resume Next
    Application.Run callback, msg
End Sub

Private Sub Agent_NotifyComplete(ByVal callback As String, ByVal msg As String)
    If Len(callback) = 0 Then Exit Sub
    On Error Resume Next
    Application.Run callback & "_Complete", msg
End Sub
