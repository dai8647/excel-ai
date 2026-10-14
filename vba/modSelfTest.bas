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

    SafeRun "Test_PathGuard"
    SafeRun "Test_Json"
    SafeRun "Test_Session"
    SafeRun "Test_Trim"
    SafeRun "Test_Settings"
    SafeRun "Test_ToolDefs"
    SafeRun "Test_ToolExec"
    SafeRun "Test_PluginsSkills"
    SafeRun "Test_AllModulesCompile"

    ChatLog_Write "system", "SelfTest: " & m_Pass & " pass / " & m_Fail & " fail"
    If m_Fail = 0 Then
        SelfTest_Run = "SelfTest: " & m_Pass & " PASS / 0 FAIL"
    Else
        SelfTest_Run = "SelfTest: " & m_Pass & " PASS / " & m_Fail & " FAIL -> " & m_Fails
    End If
End Function

' Run a test sub without ever letting a runtime error surface as a modal
' VBA dialog (which would hang COM automation forever). A crash becomes a FAIL.
Private Sub SafeRun(ByVal testSubName As String)
    On Error GoTo Handler
    Application.Run "'" & ThisWorkbook.Name & "'!" & testSubName
    Exit Sub
Handler:
    Check "crash@" & testSubName & " (" & Err.Number & ": " & Err.Description & ")", False
    Err.Clear
End Sub

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

Public Sub Test_PathGuard()
    Dim resolved As String
    resolved = ResolveSafePath("README.txt")
    Check "guard-inside", IsPathInWorkspace(resolved)
    Check "guard-parent-blocked", Rejects("..\evil.txt")
    Check "guard-outside-blocked", Rejects("C:\Windows")
End Sub

Public Sub Test_Json()
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

Public Sub Test_Session()
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

Public Sub Test_Trim()
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

Public Sub Test_Settings()
    InitDefaultSettings
    Check "settings-maxtokens", GetMaxTokens() >= 256
End Sub

Public Sub Test_ToolDefs()
    Dim oai As String, ant As String
    oai = Tools_GetDefinitionsJson()
    ant = Tools_GetAnthropicDefinitionsJson()
    Check "tooldefs-openai", Len(oai) > 800 And InStr(oai, "web_fetch") > 0 And InStr(oai, "update_excel") > 0
    Check "tooldefs-anthropic", InStr(ant, "create_chart") > 0 And InStr(ant, "update_excel") > 0
End Sub

Public Sub Test_ToolExec()
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

Public Function SelfTest_SkillTarget(ByVal argsJson As String) As String
    SelfTest_SkillTarget = "skill-ran:" & Left$(argsJson, 40)
End Function

Public Sub Test_PluginsSkills()
    Dim defs As String
    Dim aa As Boolean
    Dim handled As Boolean
    Dim r As String
    Dim ws As Worksheet
    Dim rTest As Long

    ' ---- plugin contract via the reference Plugin_Echo ----
    defs = Tools_GetDefinitionsJson()
    Check "plugin-defs-openai", InStr(defs, "echo_last_write") > 0
    Check "plugin-defs-anthropic", InStr(Tools_GetAnthropicDefinitionsJson(), "echo_last_write") > 0
    Check "plugin-registered", InStr(Plugins_List(), "Plugin_Echo") > 0

    r = Plugins_TryExecute("echo", "{""text"":""ping-selftest""}", aa, handled)
    Check "plugin-exec", handled And r = "ping-selftest"
    aa = False
    r = Tools_Execute("echo", "{""text"":""via-dispatch""}", aa)
    Check "plugin-dispatch", InStr(r, "via-dispatch") > 0
    aa = False
    r = Tools_Execute("no_such_tool_xyz", "{}", aa)
    Check "unknown-tool-still-errors", InStr(r, "unknown") > 0 Or InStr(r, "未知") > 0

    ' Every enabled plugin may expose Plugin_SelfTest(): returning "OK" passes,
    ' returning "" is skipped, anything else fails the harness with that text.
    Dim mods As Collection
    Dim mi As Variant
    Dim plugTest As String
    Set mods = Plugins_EnabledModules()
    For Each mi In mods
        plugTest = ""
        On Error Resume Next
        plugTest = Application.Run(CStr(mi) & ".Plugin_SelfTest")
        If Err.Number <> 0 Then plugTest = ""
        Err.Clear
        On Error GoTo 0
        If plugTest = "OK" Then
            Check "plugin-selftest:" & CStr(mi), True
        ElseIf Len(plugTest) > 0 Then
            Check "plugin-selftest:" & CStr(mi) & " -> " & Left$(plugTest, 120), False
        End If
    Next mi

    ' OnBeforeWrite hook: an approved write_file must notify Plugin_Echo
    Plugin_Echo.LastWritePath = ""
    aa = True
    r = Tools_Execute("write_file", "{""path"":""selftest_plugin.txt"",""content"":""x""}", aa)
    Check "plugin-before-write-hook", InStr(Plugin_Echo.LastWritePath, "selftest_plugin.txt") > 0
    On Error Resume Next
    Kill GetWorkspacePath() & "\selftest_plugin.txt"
    On Error GoTo 0

    ' ---- user skills on the skills sheet ----
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_SKILLS)
    On Error GoTo 0
    Check "skills-sheet-exists", Not ws Is Nothing
    If ws Is Nothing Then Exit Sub

    rTest = 8
    ws.Cells(rTest, 1).Value = "selftest-skill"
    ws.Cells(rTest, 2).Value = "vba"
    ws.Cells(rTest, 3).Value = "selftest vba skill"
    ws.Cells(rTest, 4).Value = "SelfTest_SkillTarget"
    ws.Cells(rTest, 6).Value = "no"
    ws.Cells(rTest, 7).Value = "yes"
    Check "skill-def-in-both-providers", _
        InStr(Tools_GetDefinitionsJson(), "vba_selftest_skill") > 0 And _
        InStr(Tools_GetAnthropicDefinitionsJson(), "vba_selftest_skill") > 0

    handled = False
    r = Skills_TryExecute("vba_selftest_skill", "{""a"":1}", aa, handled)
    Check "skill-vba-exec", handled And InStr(r, "skill-ran:") > 0

    ' prompt skills are injected into the system prompt only while enabled
    ws.Cells(rTest + 1, 1).Value = "selftest-prompt"
    ws.Cells(rTest + 1, 2).Value = "prompt"
    ws.Cells(rTest + 1, 3).Value = "selftest prompt skill"
    ws.Cells(rTest + 1, 4).Value = "SAY_TESTPHRASE_XYZZY"
    ws.Cells(rTest + 1, 7).Value = "yes"
    Check "skill-prompt-injected", InStr(BuildSystemPrompt(), "SAY_TESTPHRASE_XYZZY") > 0

    ws.Cells(rTest, 7).Value = "no"
    ws.Cells(rTest + 1, 7).Value = "no"
    Check "skill-disable-hides-prompt", InStr(BuildSystemPrompt(), "SAY_TESTPHRASE_XYZZY") = 0

    Dim rr As Long
    For rr = rTest To rTest + 1
        ws.Rows(rr).ClearContents
    Next rr
End Sub

Public Sub Test_AllModulesCompile()
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
