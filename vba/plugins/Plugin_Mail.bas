Attribute VB_Name = "Plugin_Mail"
Option Explicit

' Outlook 下書き作成プラグイン。
' draft_email ツール: 宛先・件名・本文のメールを作成し、下書きとして表示する(.Send はしない)。
' 添付はワークスペース内の相対パス → ResolveSafePath で絶対パス化して添付する。
' データを外部(Outlook/送信)へ出すため、書き込み確認ゲートを通過する書き込み系ツールとして登録する。

Public Function Plugin_Name() As String
    Plugin_Name = "mail"
End Function

Public Function Plugin_ToolNames() As String
    Plugin_ToolNames = "draft_email"
End Function

Public Function Plugin_WriteTools() As String
    Plugin_WriteTools = "draft_email"
End Function

Public Function Plugin_ToolsOpenAI() As String
    Plugin_ToolsOpenAI = ToolDef("draft_email", "Outlook の下書きメールを作成して表示する (送信はしない)。" & _
        "to=宛先, subject=件名, body=本文, attachment=任意の添付ファイル(ワークスペース内の相対パス)", _
        "{""to"":{""type"":""string"",""description"":""宛先メールアドレス""}," & _
        """subject"":{""type"":""string"",""description"":""件名""}," & _
        """body"":{""type"":""string"",""description"":""本文""}," & _
        """attachment"":{""type"":""string"",""description"":""添付ファイルのワークスペース内相対パス(任意)""}}")
End Function

Public Function Plugin_ToolsAnthropic() As String
    Plugin_ToolsAnthropic = AnthropicToolDef("draft_email", "Outlook の下書きメールを作成して表示する (送信はしない)。" & _
        "to=宛先, subject=件名, body=本文, attachment=任意の添付ファイル(ワークスペース内の相対パス)", _
        "{""to"":{""type"":""string"",""description"":""宛先メールアドレス""}," & _
        """subject"":{""type"":""string"",""description"":""件名""}," & _
        """body"":{""type"":""string"",""description"":""本文""}," & _
        """attachment"":{""type"":""string"",""description"":""添付ファイルのワークスペース内相対パス(任意)""}}")
End Function

Public Function Plugin_Execute(ByVal toolName As String, ByVal argsJson As String) As String
    On Error GoTo HandleErr
    Select Case LCase$(toolName)
        Case "draft_email"
            Plugin_Execute = MailDraftEmail(argsJson)
        Case Else
            Plugin_Execute = "Plugin_Mail: 未知のツール " & toolName
    End Select
    Exit Function
HandleErr:
    Plugin_Execute = "Plugin_Mail エラー: " & Err.Description
End Function

Private Function MailDraftEmail(ByVal argsJson As String) As String
    Dim outlook As Object
    Dim mail As Object
    Dim toAddr As String
    Dim subject As String
    Dim body As String
    Dim attachRel As String
    Dim attachPath As String

    toAddr = Trim$(JsonGetArg(argsJson, "to"))
    subject = JsonGetArg(argsJson, "subject")
    body = JsonGetArg(argsJson, "body")
    attachRel = Trim$(JsonGetArg(argsJson, "attachment"))

    If Len(toAddr) = 0 Then
        MailDraftEmail = "エラー: 宛先(to)を指定してください。"
        Exit Function
    End If

    Set outlook = CreateObject("Outlook.Application")
    Set mail = outlook.CreateItem(0)   ' olMailItem

    mail.To = toAddr
    mail.Subject = subject
    mail.Body = body

    If Len(attachRel) > 0 Then
        On Error GoTo PathErr
        attachPath = ResolveSafePath(attachRel)
        On Error GoTo HandleErr
        mail.Attachments.Add attachPath
    End If

    mail.Display
    Set mail = Nothing
    Set outlook = Nothing
    MailDraftEmail = "下書きを表示しました: " & toAddr
    Exit Function

PathErr:
    MailDraftEmail = "エラー: 添付ファイルを解決できませんでした: " & Err.Description
    Exit Function

HandleErr:
    MailDraftEmail = "メール下書き作成でエラーが発生しました: " & Err.Description
End Function

Public Function Plugin_SelfTest() As String
    Dim outlook As Object
    On Error GoTo NoOutlook
    Set outlook = CreateObject("Outlook.Application")
    Set outlook = Nothing
    Plugin_SelfTest = "OK"
    Exit Function
NoOutlook:
    Plugin_SelfTest = ""
End Function