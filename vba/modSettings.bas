Attribute VB_Name = "modSettings"
Option Explicit

Public Const APP_VERSION As String = "1.1.0"
Public Const SHEET_SETTINGS As String = "設定"
Public Const SHEET_CHATLOG As String = "チャットログ"
Public Const SHEET_OPS As String = "操作"
Public Const SHEET_MODELS As String = "モデル一覧"

Private Const ROW_PROVIDER As Long = 2
Private Const ROW_API_KEY As Long = 3
Private Const ROW_BASE_URL As Long = 4
Private Const ROW_MODEL As Long = 5
Private Const ROW_WORKSPACE As Long = 6
Private Const ROW_AUTO_RUN As Long = 7
Private Const ROW_CONFIRM_DANGER As Long = 8
Private Const ROW_MAX_STEPS As Long = 9
Private Const ROW_ANTHROPIC_VERSION As Long = 10

Public Function GetSettingValue(ByVal labelRow As Long) As String
    GetSettingValue = Trim$(CStr(ThisWorkbook.Worksheets(SHEET_SETTINGS).Cells(labelRow, 2).Value))
End Function

Public Sub SetSettingValue(ByVal labelRow As Long, ByVal value As String)
    ThisWorkbook.Worksheets(SHEET_SETTINGS).Cells(labelRow, 2).Value = value
End Sub

Public Function GetProvider() As String
    If LCase$(GetSettingValue(ROW_PROVIDER)) = "anthropic" Then
        GetProvider = "Anthropic"
    Else
        GetProvider = "OpenAI"
    End If
End Function

Public Function IsAnthropicProvider() As Boolean
    IsAnthropicProvider = (GetProvider() = "Anthropic")
End Function

Public Function GetApiKey() As String
    GetApiKey = GetSettingValue(ROW_API_KEY)
End Function

Public Function GetBaseUrl() As String
    Dim url As String
    url = GetSettingValue(ROW_BASE_URL)
    If Len(url) = 0 Then
        If IsAnthropicProvider() Then
            url = "https://api.anthropic.com"
        Else
            url = "https://api.openai.com/v1"
        End If
    End If
    GetBaseUrl = TrimTrailingSlash(url)
End Function

Public Function GetModel() As String
    GetModel = GetSettingValue(ROW_MODEL)
End Function

Public Function GetWorkspacePath() As String
    Dim p As String
    p = GetSettingValue(ROW_WORKSPACE)
    If Len(p) = 0 Then p = ThisWorkbook.Path & "\workspace"
    GetWorkspacePath = NormalizePath(p)
End Function

Public Function IsAutoRunEnabled() As Boolean
    IsAutoRunEnabled = ParseYesNo(GetSettingValue(ROW_AUTO_RUN), False)
End Function

Public Function IsConfirmDangerEnabled() As Boolean
    IsConfirmDangerEnabled = ParseYesNo(GetSettingValue(ROW_CONFIRM_DANGER), True)
End Function

Public Function GetMaxSteps() As Long
    GetMaxSteps = Val(GetSettingValue(ROW_MAX_STEPS))
    If GetMaxSteps <= 0 Then GetMaxSteps = 12
End Function

Public Function GetAnthropicVersion() As String
    GetAnthropicVersion = GetSettingValue(ROW_ANTHROPIC_VERSION)
    If Len(GetAnthropicVersion) = 0 Then GetAnthropicVersion = "2023-06-01"
End Function

Private Function ParseYesNo(ByVal s As String, ByVal defaultYes As Boolean) As Boolean
    Select Case LCase$(Trim$(s))
        Case "はい", "yes", "y", "true", "1", "on": ParseYesNo = True
        Case "いいえ", "no", "n", "false", "0", "off": ParseYesNo = False
        Case Else: ParseYesNo = defaultYes
    End Select
End Function

Private Function TrimTrailingSlash(ByVal s As String) As String
    Do While Len(s) > 0 And Right$(s, 1) = "/"
        s = Left$(s, Len(s) - 1)
    Loop
    TrimTrailingSlash = s
End Function

Public Function NormalizePath(ByVal p As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo Fallback
    NormalizePath = fso.GetAbsolutePathName(p)
    Exit Function
Fallback:
    NormalizePath = Replace(p, "/", "\")
End Function

Public Sub EnsureWorkspaceExists()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(GetWorkspacePath()) Then fso.CreateFolder GetWorkspacePath()
End Sub

Public Sub InitDefaultSettings()
    If Len(GetSettingValue(ROW_PROVIDER)) = 0 Then SetSettingValue ROW_PROVIDER, "OpenAI"
    If Len(GetSettingValue(ROW_BASE_URL)) = 0 Then SetSettingValue ROW_BASE_URL, "http://127.0.0.1:20128/v1"
    If Len(GetSettingValue(ROW_WORKSPACE)) = 0 Then SetSettingValue ROW_WORKSPACE, ThisWorkbook.Path & "\workspace"
    If Len(GetSettingValue(ROW_AUTO_RUN)) = 0 Then SetSettingValue ROW_AUTO_RUN, "いいえ"
    If Len(GetSettingValue(ROW_CONFIRM_DANGER)) = 0 Then SetSettingValue ROW_CONFIRM_DANGER, "はい"
    If Len(GetSettingValue(ROW_MAX_STEPS)) = 0 Then SetSettingValue ROW_MAX_STEPS, "12"
    If Len(GetSettingValue(ROW_ANTHROPIC_VERSION)) = 0 Then SetSettingValue ROW_ANTHROPIC_VERSION, "2023-06-01"
    EnsureWorkspaceExists
End Sub

Public Sub OpenWorkspaceFolder()
    EnsureWorkspaceExists
    Shell "explorer.exe """ & GetWorkspacePath() & """", vbNormalFocus
End Sub

Public Sub TestConnection()
    On Error GoTo ErrHandler
    Dim result As String
    result = LLM_TestConnection()
    Ops_WriteStatus result, False
    MsgBox result, vbInformation, "接続テスト"
    Exit Sub
ErrHandler:
    Ops_WriteStatus "接続失敗: " & Err.Description, True
    MsgBox "接続テスト失敗: " & Err.Description, vbCritical
End Sub

Public Sub RefreshModels()
    LLM_RefreshModelList
End Sub

Public Sub ShowAgent()
    InitDefaultSettings
    ShowAgentOrFallback
End Sub

Public Sub ClearChatLog()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CHATLOG)
    If MsgBox("チャットログをクリアしますか？", vbYesNo + vbQuestion) = vbYes Then
        ws.Rows("2:" & ws.Rows.Count).ClearContents
    End If
End Sub

Public Sub Ops_WriteStatus(ByVal message As String, ByVal isError As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_OPS)
    ws.Range("B14").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("B15").Value = message
    If isError Then
        ws.Range("B15").Font.Color = RGB(192, 0, 0)
    Else
        ws.Range("B15").Font.Color = RGB(0, 97, 0)
    End If
End Sub
