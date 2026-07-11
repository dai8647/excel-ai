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

    On Error GoTo ErrHandler

    InitDefaultSettings
    EnsureWorkspaceExists
    ChatLog_Write "user", userPrompt

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
                Agent_ReportStatus statusCallback, "ツール実行: " & calls(i).toolName
                ChatLog_Write "tool", calls(i).toolName & " " & calls(i).argumentsJson

                toolResult = Tools_Execute(calls(i).toolName, calls(i).argumentsJson, autoApproved)
                ChatLog_Write "tool_result", Left$(toolResult, 2000)

                If Len(calls(i).id) > 0 Then
                    messages.Add JsonBuildToolResult(calls(i).id, toolResult)
                Else
                    messages.Add JsonBuildUserMessage("[tool:" & calls(i).toolName & "] " & toolResult)
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
    Agent_ReportStatus statusCallback, "ステップ上限に達しました"
    Agent_NotifyComplete statusCallback, "最大ステップ数に達したため停止しました。続きを入力してください。"
    Exit Sub

ErrHandler:
    ChatLog_Write "error", Err.Description
    Agent_ReportStatus statusCallback, "エラー"
    Agent_NotifyComplete statusCallback, "処理を完了できませんでした: " & Err.Description
End Sub

Public Sub RunSheetPrompt()
    Dim ws As Worksheet
    Dim prompt As String

    Set ws = ThisWorkbook.Worksheets(SHEET_OPS)
    prompt = Trim$(CStr(ws.Range("B7").Value))
    If Len(prompt) = 0 Then
        MsgBox "Request 欄に依頼を入力してください。", vbExclamation, "ExcelAI"
        Exit Sub
    End If

    ws.Range("B9").Value = "処理中..."
    ws.Range("B9").Font.Color = RGB(31, 78, 121)
    Agent_RunPrompt prompt, ""
    ws.Range("B9").Value = ChatLog_GetRecentAssistantText()
    ws.Range("B9").Font.Color = RGB(0, 0, 0)
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
