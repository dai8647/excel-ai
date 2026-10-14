Attribute VB_Name = "modSelfTest"
Option Explicit

' Headless acceptance harness. From PowerShell/macros:
'   Application.Run "SelfTest_Run"
' expects a string ending in "0 FAIL". Covers guard, JSON, session, trim,
' tool definitions, settings, and module-wide compilation. No network calls.

Private m_Pass As Long
Private m_Fail As Long
Private m_Fails As String

Public Function SelfTest_Run() As String
    m_Pass = 0
    m_Fail = 0
    m_Fails = ""

    Test_PathGuard
    Test_Json
    Test_Session
    Test_Trim
    Test_Settings
    Test_ToolDefs
    Test_ToolExec
    Test_AllModulesCompile

    ChatLog_Write "system", "SelfTest: " & m_Pass & " pass / " & m_Fail & " fail"
    If m_Fail = 0 Then
        SelfTest_Run = "SelfTest: " & m_Pass & " PASS / 0 FAIL"
    Else
        SelfTest_Run = "SelfTest: " & m_Pass & " PASS / " & m_Fail & " FAIL -> " & m_Fails
    End If
End Function

Private Sub Check(ByVal name As String, ByVal ok As Boolean)
    If ok Then
        m_Pass = m_Pass + 1
    Else
        m_Fail = m_Fail + 1
        m_Fails = m_Fails & IIf(Len(m_Fails) = 0, "", "; ") & name
    End If
End Sub

Private Function Rejects(ByVal inputPath As String) As Boolean
    On Error Resume Next
    Err.Clear
    ResolveSafePath inputPath
    Rejects = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub Test_PathGuard()
    Dim resolved As String
    resolved = ResolveSafePath("README.txt")
    Check "guard-inside", IsPathInWorkspace(resolved)
    Check "guard-parent-blocked", Rejects("..\evil.txt")
    Check "guard-outside-blocked", Rejects("C:\Windows")
End Sub

Private Sub Test_Json()
    Dim j As String, back As String
    j = JsonBuildUserMessage("hello ""quote"" 自己テスト" & vbLf & "line2")
    back = JsonExtractString(j, "content")
    Check "json-roundtrip", InStr(back, "自己テスト") > 0 And InStr(back, "line2") > 0
    Check "json-kanji-raw", InStr(j, "自己") > 0
    Check "json-unescape", JsonExtractString("{""a"":""\u81EA\u5DF1""}", "a") = "自己"

    Dim tc As String
    tc = "{""choices"":[{""message"":{""tool_calls"":[" & _
        "{""id"":""c1"",""type"":""function"",""function"":{""name"":""list_dir"",""arguments"":""{}""}}," & _
        "{""id"":""c2"",""type"":""function"",""function"":{""name"":""grep_files"",""arguments"":""{}""}}]}}]}"
    Dim calls() As ToolCallItem
    Dim n As Long
    n = JsonParseToolCalls(tc, calls)
    Check "json-toolcalls", n = 2 And calls(1).toolName = "list_dir" And calls(2).id = "c2"

    Check "json-finish", JsonExtractString("{""choices"":[{""finish_reason"":""length""}]}", "finish_reason") = "length"

    Dim stripped As String
    stripped = JsonRemoveKey("{""role"":""assistant"",""content"":null,""reasoning_content"":null,""tool_calls"":[]}", "reasoning_content")
    Check "json-removekey-null-mid", InStr(stripped, "reasoning_content") = 0 And JsonExtractString(stripped, "role") = "assistant" And InStr(stripped, """content"":null,") > 0
    ' reasoning value is a quoted string containing comma and closing brace
    Dim raw As String
    raw = "{""role"":""assistant"",""reasoning_content"":""a, b } c"",""content"":""ok""}"
    stripped = JsonRemoveKey(raw, "reasoning_content")
    Check "json-removekey-string-with-punct", InStr(stripped, "a, b } c") = 0 And JsonExtractString(stripped, "content") = "ok"
    stripped = JsonRemoveKey("{""a"":1,""reasoning_content"":null}", "reasoning_content")
    Check "json-removekey-last", stripped = "{""a"":1}"
    Check "json-removenullkey", JsonRemoveNullKey("{""tool_calls"":null}", "tool_calls") = "{}" And InStr(JsonRemoveNullKey("{""tool_calls"":[],""x"":1}", "tool_calls"), "tool_calls") > 0

    ' router-style null tool_calls must NOT count as a tool call
    Check "json-hasnull-false", Not JsonHasToolCalls("{""choices"":[{""message"":{""role"":""assistant"",""content"":""done"",""reasoning_content"":""x"",""tool_calls"":null,""logprobs"":null}}]}")
    Check "json-hasnull-true", JsonHasToolCalls("{""choices"":[{""message"":{""content"":null,""tool_calls"": [{""id"":""a""}]}}]}")
End Sub

Private Sub Test_Session()
    Dim items As Collection

    Session_Reset
    Check "session-empty", Session_MessageCount() = 0

    Session_AddMessage JsonBuildUserMessage("自己テストA")
    Session_AddMessage "{""role"":""assistant"",""content"":""ok""}"
    Set items = Session_LoadMessages()
    Check "session-count", items.Count = 2
    Check "session-content", InStr(CStr(items(1)), "自己テストA") > 0

    Session_Reset
    Dim big As String
    big = String$(50000, "あ")
    Session_AddMessage JsonBuildUserMessage(big)
    Set items = Session_LoadMessages()
    Check "session-oversize", items.Count = 1 And Len(JsonExtractString(CStr(items(1)), "content")) = 50000

    Session_Reset
    Session_AddMessage JsonBuildUserMessage("q")
    Session_AddMessage "{""role"":""assistant"",""content"":null,""tool_calls"":[{""id"":""z"",""type"":""function"",""function"":{""name"":""list_dir"",""arguments"":""{}""}}]}"
    Set items = Session_LoadMessages()
    Check "session-drop-pending", items.Count = 1

    Session_Reset
    Dim working As New Collection
    Session_Remember JsonBuildUserMessage("w"), working
    Check "session-remember", working.Count = 1 And Session_MessageCount() = 1
    Session_Reset
End Sub

Private Sub Test_Trim()
    Dim hist As New Collection
    Dim i As Long
    Dim role As String
    hist.Add "{""role"":""system"",""content"":""s""}"
    For i = 1 To 100
        If (i Mod 2) = 1 Then role = "user" Else role = "assistant"
        hist.Add "{""role"":""" & role & """,""content"":""m" & i & """}"
    Next i
    Dim tr As Collection
    Set tr = SessionTrimForModel(hist)
    Check "trim-smaller", tr.Count < hist.Count And tr.Count >= SESSION_KEEP_MESSAGES
    Check "trim-keeps-system", JsonExtractString(CStr(tr(1)), "role") = "system"
    Check "trim-safe-boundary", JsonExtractString(CStr(tr(2)), "role") = "user"
End Sub

Private Sub Test_Settings()
    InitDefaultSettings
    Check "settings-maxtokens", GetMaxTokens() >= 256
End Sub

Private Sub Test_ToolDefs()
    Dim oai As String, ant As String
    oai = Tools_GetDefinitionsJson()
    ant = Tools_GetAnthropicDefinitionsJson()
    Check "tooldefs-openai", Len(oai) > 800 And InStr(oai, "web_fetch") > 0 And InStr(oai, "update_excel") > 0
    Check "tooldefs-anthropic", InStr(ant, "create_chart") > 0 And InStr(ant, "update_excel") > 0
End Sub

Private Sub Test_ToolExec()
    Dim auto As Boolean
    Dim res As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    auto = True

    res = Tools_Execute("make_dir", "{""path"":""selftest_tmp""}", auto)
    Check "tool-make_dir", fso.FolderExists(GetWorkspacePath() & "\selftest_tmp")

    res = Tools_Execute("write_file", "{""path"":""selftest_tmp/a.txt"",""content"":""row1,row2" & Chr(92) & "n1,2""}", auto)
    Check "tool-write_file", InStr(res, "書き込み完了") > 0

    res = Tools_Execute("copy_file", "{""source"":""selftest_tmp/a.txt"",""destination"":""selftest_tmp/b.txt""}", auto)
    Check "tool-copy_file", fso.FileExists(GetWorkspacePath() & "\selftest_tmp\b.txt")

    res = Tools_Execute("list_dir", "{""path"":""."",""recursive"":true}", auto)
    Check "tool-list_dir-recursive", InStr(res, "selftest_tmp") > 0 And InStr(res, "b.txt") > 0

    ' Full Excel roundtrip: create -> update with formula -> inspect sees formula
    res = Tools_Execute("create_excel", "{""path"":""selftest_tmp/wb.xlsx"",""sheet_name"":""Data"",""csv_content"":""x,y" & Chr(92) & "n1,2""}", auto)
    Check "tool-create_excel", InStr(res, "Excel作成完了") > 0

    res = Tools_Execute("update_excel", "{""path"":""selftest_tmp/wb.xlsx"",""new_sheet"":""Calc"",""edits"":""A1=5" & Chr(92) & "nA2=7" & Chr(92) & "nA3==SUM(A1:A2)""}", auto)
    Check "tool-update_excel", InStr(res, "Excel更新完了") > 0

    res = Tools_Execute("inspect_excel", "{""path"":""selftest_tmp/wb.xlsx"",""sheet_name"":""Calc"",""include_formulas"":true}", auto)
    Check "tool-update-formula-live", InStr(res, "SUM(A1:A2)") > 0 And InStr(res, "Calc") > 0

    On Error Resume Next
    fso.DeleteFile GetWorkspacePath() & "\selftest_tmp\wb.xlsx", True
    fso.DeleteFolder GetWorkspacePath() & "\selftest_tmp", True
    On Error GoTo 0
End Sub

Private Sub Test_AllModulesCompile()
    Dim s As String
    s = GetProvider()                  ' modSettings
    s = s & BuildSystemPrompt()        ' modLLM
    s = s & JsonString("x")            ' modJson
    s = s & WorkspaceSafetyReport()    ' modPathGuard
    s = s & ChatLog_GetRecentAssistantText() ' modChatLog
    s = s & Tools_GetDefinitionsJson() ' modTools
    RefreshHomeStatus                  ' modSheetUI
    Agent_RequestStop                  ' modAgentCore (idle: no-op)

    ' RunSheetPrompt path touches merged B21:H24 — assigning "" must not raise 1004
    Dim ws As Worksheet
    Dim e As Long
    Set ws = GetOpsSheet()
    On Error Resume Next
    ws.Range("B21").Value = "selftest"
    ws.Range("B21").Value = ""
    e = Err.Number
    On Error GoTo 0
    Check "ui-merged-cell-write", e = 0 And CStr(ws.Range("B21").Value) = ""
    Check "compile-modules", True
End Sub
