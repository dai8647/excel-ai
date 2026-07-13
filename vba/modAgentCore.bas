Attribute VB_Name = "modAgentCore"
Option Explicit

Public Sub Agent_RunPrompt(ByVal userPrompt As String, ByVal statusCallback As String)
    Dim messages As Collection
    Dim stepNum As Long
    Dim maxSteps As Long
    Dim responseJson As String
    Dim assistantContent As String
    Dim rawMsg As String
    Dim calls() As ToolCallItem
    Dim callCount As Long
    Dim i As Long
    Dim toolResult As String
    Dim autoApproved As Boolean
    Dim contextResult As String

    On Error GoTo ErrHandler

    InitDefaultSettings
    EnsureWorkspaceExists
    ChatLog_Write "user", userPrompt
    Ops_WriteProgress "AIが依頼内容を整理しています...", False

    Set messages = BuildInitialMessages(userPrompt)
    maxSteps = GetMaxSteps()
    autoApproved = IsAutoRunEnabled()

    For stepNum = 1 To maxSteps
        Agent_ReportStatus statusCallback, "AI思考中... (step " & stepNum & "/" & maxSteps & ")"
        DoEvents

        responseJson = LLM_ChatWithTools(messages)

        If JsonHasToolCalls(responseJson) Then
            rawMsg = LLM_ExtractAssistantRawMessageObject(responseJson)
            If Len(rawMsg) > 0 Then messages.Add rawMsg

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
                    messages.Add JsonBuildToolResult(calls(i).id, contextResult)
                Else
                    messages.Add JsonBuildUserMessage("[tool:" & calls(i).toolName & "] " & LimitToolResultForModel(toolResult))
                End If
            Next i
        Else
            assistantContent = LLM_ExtractAssistantMessage(responseJson)
            If Len(assistantContent) = 0 Then assistantContent = "（応答を取得できませんでした）"
            ChatLog_Write "assistant", assistantContent
            Agent_ReportStatus statusCallback, "完了"
            Agent_NotifyComplete statusCallback, assistantContent
            Exit Sub
        End If
    Next stepNum

    ChatLog_Write "error", "最大ステップ数 (" & maxSteps & ") に達しました"
    Ops_WriteStatus "最大ステップ数 (" & maxSteps & ") に達しました", True
    Agent_ReportStatus statusCallback, "ステップ上限に達しました"
    Agent_NotifyComplete statusCallback, "最大ステップ数に達したため停止しました。続きを入力してください。"
    Exit Sub

ErrHandler:
    ChatLog_Write "error", "step=" & stepNum & ": " & Err.Description
    Ops_WriteStatus "step=" & stepNum & ": " & Err.Description, True
    Agent_ReportStatus statusCallback, "エラー"
    Agent_NotifyComplete statusCallback, "処理を完了できませんでした: " & Err.Description
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

    Set ws = GetOpsSheet()
    prompt = Trim$(CStr(ws.Range("B9").Value))
    If Len(prompt) = 0 Then
        Ops_WriteProgress "依頼を入力してください", True
        Exit Sub
    End If

    ws.Range("B21").ClearContents
    Ops_WriteProgress "依頼を送信しています...", False
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
