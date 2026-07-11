Attribute VB_Name = "modChatLog"
Option Explicit

Public Sub ChatLog_Write(ByVal entryType As String, ByVal content As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_CHATLOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = entryType
    ws.Cells(nextRow, 3).Value = content
End Sub

Public Function ChatLog_GetRecentAssistantText() As String
    Dim ws As Worksheet
    Dim r As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_CHATLOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    For r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If ws.Cells(r, 2).Value = "assistant" Then
            ChatLog_GetRecentAssistantText = CStr(ws.Cells(r, 3).Value)
            Exit Function
        End If
    Next r
End Function
