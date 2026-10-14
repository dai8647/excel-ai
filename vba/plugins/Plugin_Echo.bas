Attribute VB_Name = "Plugin_Echo"
Option Explicit

' Reference plugin + SelfTest target. Demonstrates the full contract:
' tool registration, write gating, and the OnBeforeWrite hook.

Public LastWritePath As String

Public Function Plugin_Name() As String
    Plugin_Name = "echo"
End Function

Public Function Plugin_ToolNames() As String
    Plugin_ToolNames = "echo,echo_last_write"
End Function

Public Function Plugin_WriteTools() As String
    Plugin_WriteTools = ""
End Function

Public Function Plugin_ToolsOpenAI() As String
    Plugin_ToolsOpenAI = ToolDef("echo", "Return the given text unchanged (plugin self-check)", _
        "{""text"":{""type"":""string""}}") & "," & _
        ToolDef("echo_last_write", "Report the file path the last gated write targeted (plugin hook self-check)", "{}")
End Function

Public Function Plugin_ToolsAnthropic() As String
    Plugin_ToolsAnthropic = AnthropicToolDef("echo", "Return the given text unchanged (plugin self-check)", _
        "{""text"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("echo_last_write", "Report the file path the last gated write targeted (plugin hook self-check)", "{}")
End Function

Public Function Plugin_Execute(ByVal toolName As String, ByVal argsJson As String) As String
    Select Case LCase$(toolName)
        Case "echo"
            Plugin_Execute = JsonGetArg(argsJson, "text")
        Case "echo_last_write"
            If Len(LastWritePath) = 0 Then
                Plugin_Execute = "(まだ書き込みはありません)"
            Else
                Plugin_Execute = LastWritePath
            End If
        Case Else
            Plugin_Execute = "Plugin_Echo: 未知のツール " & toolName
    End Select
End Function

Public Sub Plugin_OnBeforeWrite(ByVal targetPath As String)
    LastWritePath = targetPath
End Sub

Public Function Plugin_SelfTest() As String
    Dim before As String
    before = LastWritePath
    Plugin_Execute "echo", "{""text"":""x""}"
    Plugin_SelfTest = "OK"
End Function
