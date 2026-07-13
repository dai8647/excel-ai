Attribute VB_Name = "modSheetUI"
Option Explicit

Public Sub ShowAgentFallback()
    InitDefaultSettings
    ShowHome
End Sub

Public Sub ShowAgentOrFallback()
    ShowAgentFallback
End Sub

Public Sub ShowHome()
    Dim ws As Worksheet
    InitDefaultSettings
    Set ws = GetOpsSheet()
    RefreshHomeStatus
    ws.Activate
    ws.Range("B9").Select
End Sub

Public Sub RefreshHomeStatus()
    Dim ws As Worksheet
    On Error GoTo Done
    Set ws = GetOpsSheet()
    ws.Range("B4").Value = GetProvider()
    ws.Range("B5").Value = GetModel()
    ws.Range("B6").Value = GetWorkspacePath()
    ws.Range("F4").Value = IIf(IsAutoRunEnabled(), "AutoRun ON", "AutoRun OFF")
    ws.Range("F5").Value = IIf(IsConfirmDangerEnabled(), "確認 ON", "確認 OFF")
Done:
End Sub

Public Sub OpenSettings()
    ThisWorkbook.Worksheets(SHEET_SETTINGS).Activate
    ThisWorkbook.Worksheets(SHEET_SETTINGS).Range("B2").Select
End Sub

Public Sub OpenChatLog()
    ThisWorkbook.Worksheets(SHEET_CHATLOG).Activate
End Sub
