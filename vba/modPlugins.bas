Attribute VB_Name = "modPlugins"
Option Explicit

' Plugin registry. A plugin is one standard .bas module named "Plugin_Xyz"
' placed in vba\plugins\ before build (auto-imported and registered there),
' or imported manually and registered with Plugins_Register.
'
' Required public members (all probed safely; a broken plugin is skipped):
'   Plugin_Name()       As String   short id, e.g. "echo"
'   Plugin_ToolNames()  As String   comma-separated lowercase tool names it owns
'   Plugin_ToolsOpenAI() As String  array body (no [ ]): comma-separated ToolDef(...) results
'   Plugin_ToolsAnthropic() As String  same shape with AnthropicToolDef(...)
'   Plugin_Execute(toolName, argsJson) As String  result text shown to the model
' Optional:
'   Plugin_WriteTools() As String   tool names that must pass the write gate
'   Plugin_OnBeforeWrite(targetPath)  fired before any gated write executes
'
' Tool names must be lowercase and unique across builtin/skill/plugin tools.

Private Const HOOK_BEFORE_WRITE As String = "Plugin_OnBeforeWrite"

Public Function Plugins_ToolsOpenAI() As String
    Plugins_ToolsOpenAI = Collect("Plugin_ToolsOpenAI")
End Function

Public Function Plugins_ToolsAnthropic() As String
    Plugins_ToolsAnthropic = Collect("Plugin_ToolsAnthropic")
End Function

Private Function Collect(ByVal fn As String) As String
    Dim mods As Collection
    Dim m As Variant
    Dim part As String
    Dim result As String

    Set mods = Plugins_EnabledModules()
    For Each m In mods
        part = CallString(CStr(m), fn)
        If Len(part) > 0 Then result = result & IIf(Len(result) = 0, "", ",") & part
    Next m
    Collect = result
End Function

' Returns CSV of tool names that must be gated by the write confirmation.
Public Function Plugins_WriteToolNames() As String
    Dim mods As Collection
    Dim m As Variant
    Dim part As String
    Dim result As String

    Set mods = Plugins_EnabledModules()
    For Each m In mods
        part = CallString(CStr(m), "Plugin_WriteTools")
        If Len(part) > 0 Then result = result & IIf(Len(result) = 0, "", ",") & part
    Next m
    Plugins_WriteToolNames = LCase$(result)
End Function

Public Function Plugins_TryExecute(ByVal toolName As String, ByVal argsJson As String, _
                                   ByRef autoApproved As Boolean, ByRef handled As Boolean) As String
    Dim mods As Collection
    Dim m As Variant
    Dim owner As String
    Dim names As String
    Dim t As Variant

    handled = False
    toolName = LCase$(toolName)

    Set mods = Plugins_EnabledModules()
    For Each m In mods
        names = LCase$(CallString(CStr(m), "Plugin_ToolNames"))
        For Each t In Split(names, ",")
            If Trim$(CStr(t)) = toolName And Len(toolName) > 0 Then
                owner = CStr(m)
                Exit For
            End If
        Next t
        If Len(owner) > 0 Then Exit For
    Next m

    If Len(owner) = 0 Then Exit Function

    handled = True
    If InStr(1, "," & Plugins_WriteToolNames() & ",", "," & toolName & ",", vbTextCompare) > 0 Then
        If Not ApproveWriteTool(toolName, argsJson, autoApproved) Then
            Plugins_TryExecute = "ユーザーが実行をキャンセルしました"
            Exit Function
        End If
    End If

    On Error GoTo PluginErr
    Plugins_TryExecute = Application.Run(owner & ".Plugin_Execute", toolName, argsJson)
    Exit Function
PluginErr:
    Plugins_TryExecute = "プラグイン '" & owner & "' でエラー: " & Err.Description
End Function

' Fire the optional OnBeforeWrite hook on every enabled plugin. Missing hook = skip.
Public Sub Plugins_FireBeforeWrite(ByVal targetPath As String)
    Dim mods As Collection
    Dim m As Variant

    If Len(targetPath) = 0 Then Exit Sub
    Set mods = Plugins_EnabledModules()
    For Each m In mods
        On Error Resume Next
        Application.Run CStr(m) & "." & HOOK_BEFORE_WRITE, targetPath
        On Error GoTo 0
    Next m
End Sub

Public Function Plugins_EnabledModules() As Collection
    Dim ws As Worksheet
    Dim modules As Collection
    Dim r As Long
    Dim lastRow As Long
    Dim moduleName As String
    Dim flag As String

    Set modules = New Collection
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PLUGINS)
    On Error GoTo 0
    If ws Is Nothing Then Set Plugins_EnabledModules = modules: Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        moduleName = Trim$(CStr(ws.Cells(r, 1).Value))
        flag = LCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        If Len(moduleName) > 0 And (flag = "yes" Or flag = "はい" Or Len(flag) = 0) Then
            If Plugins_HasModule(moduleName) Then modules.Add moduleName
        End If
    Next r
    Set Plugins_EnabledModules = modules
End Function

Public Function Plugins_HasModule(ByVal moduleName As String) As Boolean
    Dim probe As String
    On Error Resume Next
    Err.Clear
    probe = Application.Run(moduleName & ".Plugin_Name")
    Plugins_HasModule = (Err.Number = 0 And Len(probe) > 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Function Plugins_List() As String
    Dim mods As Collection
    Dim m As Variant

    Set mods = Plugins_EnabledModules()
    If mods.Count = 0 Then
        Plugins_List = "(有効なプラグインはありません)"
        Exit Function
    End If
    For Each m In mods
        Plugins_List = Plugins_List & IIf(Len(Plugins_List) = 0, "", ", ") & m & " (" & CallString(CStr(m), "Plugin_Name") & ")"
    Next m
End Function

' Manual registration after importing a plugin .bas at runtime.
Public Sub Plugins_Register()
    Dim ws As Worksheet
    Dim moduleName As String
    Dim r As Long
    Dim existing As Boolean

    moduleName = Trim$(InputBox("プラグインモジュール名 (例: Plugin_Snapshot):", "プラグイン登録"))
    If Len(moduleName) = 0 Then Exit Sub

    If Not Plugins_HasModule(moduleName) Then
        MsgBox "モジュール '" & moduleName & "' が見つからないか規約(Plugin_Name)を満たしていません。", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_PLUGINS)
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), moduleName, vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value = "yes"
            MsgBox moduleName & " を有効にしました。", vbInformation
            Exit Sub
        End If
    Next r
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = moduleName
    ws.Cells(r, 2).Value = "yes"
    MsgBox moduleName & " を登録しました。", vbInformation
End Sub

Public Sub Plugins_Unregister()
    Dim ws As Worksheet
    Dim moduleName As String
    Dim r As Long

    moduleName = Trim$(InputBox("無効化するモジュール名:", "プラグイン解除"))
    If Len(moduleName) = 0 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SHEET_PLUGINS)
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), moduleName, vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value = "no"
            MsgBox moduleName & " を無効化しました。", vbInformation
            Exit Sub
        End If
    Next r
End Sub

' Safe Application.Run for a String-returning function; returns "" on any error.
Private Function CallString(ByVal moduleFn As String, ByVal fn As String) As String
    On Error GoTo Done
    CallString = Application.Run(moduleFn & "." & fn)
Done:
End Function
