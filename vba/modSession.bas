Attribute VB_Name = "modSession"
Option Explicit

' Multi-turn conversation history, persisted next to the workbook as
' .excelai-session.json so follow-up requests ("さっきの直して", ステップ上限後の
' 「続き」) can see prior tool results. Reset with the "新しい会話" button.

Private Const SESSION_FILE_NAME As String = ".excelai-session.json"
Public Const SESSION_KEEP_MESSAGES As Long = 40
Private Const SESSION_HARD_MAX As Long = 120

Public Function SessionFilePath() As String
    If Len(ThisWorkbook.Path) = 0 Then
        SessionFilePath = ""
    Else
        SessionFilePath = ThisWorkbook.Path & Application.PathSeparator & SESSION_FILE_NAME
    End If
End Function

Public Function Session_LoadMessages() As Collection
    Dim items As New Collection
    Dim fso As Object
    Dim stream As Object
    Dim raw As String
    Dim openPos As Long, closePos As Long
    Dim pos As Long, blockStart As Long, blockEnd As Long
    Dim depthFix As Boolean

    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(SessionFilePath()) = 0 Then
        Set Session_LoadMessages = items
        Exit Function
    End If
    If Not fso.FileExists(SessionFilePath()) Then
        Set Session_LoadMessages = items
        Exit Function
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile SessionFilePath()
    raw = stream.ReadText
    stream.Close

    openPos = InStr(raw, "[")
    closePos = InStrRev(raw, "]")
    If openPos = 0 Or closePos <= openPos Then GoTo Done

    pos = openPos
    Do
        blockStart = InStr(pos + 1, raw, "{")
        If blockStart = 0 Or blockStart >= closePos Then Exit Do
        blockEnd = JsonFindClosingDelimiter(raw, blockStart, "{", "}")
        If blockEnd = 0 Or blockEnd > closePos Then Exit Do
        items.Add Mid$(raw, blockStart, blockEnd - blockStart + 1)
        pos = blockEnd
    Loop

    ' Never restore into a state that ends mid tool-call sequence: the last
    ' restored message must not be an assistant turn with pending tool_calls.
    Do While items.Count > 0
        If JsonHasToolCalls(CStr(items(items.Count))) Then
            items.Remove items.Count
            depthFix = True
        Else
            Exit Do
        End If
    Loop

    If items.Count > SESSION_HARD_MAX Then
        Dim trimmed As New Collection
        Dim startIdx As Long
        startIdx = items.Count - SESSION_KEEP_MESSAGES + 1
        If SafeBoundaryIndex(items, startIdx) > 0 Then
            For pos = SafeBoundaryIndex(items, startIdx) To items.Count
                trimmed.Add items(pos)
            Next pos
            Set items = trimmed
        End If
    End If
Done:
    Set Session_LoadMessages = items
End Function

Public Sub Session_AddMessage(ByVal msgJson As String)
    Dim items As Collection
    Dim stream As Object

    On Error Resume Next
    If Len(msgJson) = 0 Then Exit Sub
    Set items = Session_LoadMessages()
    items.Add msgJson
    If items.Count > SESSION_HARD_MAX Then
        Dim startIdx As Long
        startIdx = 2
        If SafeBoundaryIndex(items, items.Count - SESSION_KEEP_MESSAGES + 1) > 0 Then _
            startIdx = SafeBoundaryIndex(items, items.Count - SESSION_KEEP_MESSAGES + 1)
        Dim trimmed As New Collection
        Dim i As Long
        For i = startIdx To items.Count
            trimmed.Add items(i)
        Next i
        Set items = trimmed
    End If

    If Len(SessionFilePath()) = 0 Then Exit Sub
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText SessionSerialize(items)
    stream.SaveToFile SessionFilePath(), 2
    stream.Close
End Sub

' Adds to the in-memory working history and persists it together.
Public Sub Session_Remember(ByVal msgJson As String, ByRef working As Collection)
    If Len(msgJson) = 0 Then Exit Sub
    working.Add msgJson
    Session_AddMessage msgJson
End Sub

Public Function Session_MessageCount() As Long
    Session_MessageCount = Session_LoadMessages().Count
End Function

Public Function SessionSerialize(ByVal messages As Collection) As String
    Dim i As Long
    Dim parts As String
    For i = 1 To messages.Count
        If i > 1 Then parts = parts & ","
        parts = parts & CStr(messages(i))
    Next i
    SessionSerialize = "[" & parts & "]"
End Function

' Trims long histories to a recent suffix. Index 1 of the working history is
' always the system prompt and is kept. Trimming only starts at a "user"
' message so no assistant tool_calls loses its tool results.
Public Function SessionTrimForModel(ByVal history As Collection) As Collection
    If history.Count <= SESSION_KEEP_MESSAGES Then
        Set SessionTrimForModel = history
        Exit Function
    End If

    Dim startIdx As Long
    startIdx = SafeBoundaryIndex(history, history.Count - SESSION_KEEP_MESSAGES + 1)
    If startIdx <= 0 Then
        Set SessionTrimForModel = history
        Exit Function
    End If

    Dim result As New Collection
    Dim i As Long
    For i = 1 To startIdx - 1
        If JsonExtractString(CStr(history(i)), "role") = "system" Then result.Add history(i)
    Next i
    For i = startIdx To history.Count
        result.Add history(i)
    Next i
    Set SessionTrimForModel = result
End Function

' First index >= suggestedStart pointing at a "user" message that does not
' break a tool-call sequence. Returns False via result 0 when none exists here
' (caller should fall back to sending the whole history).
Public Function SafeBoundaryIndex(ByVal messages As Collection, ByVal suggestedStart As Long) As Long
    Dim j As Long
    Dim prev As String
    If suggestedStart < 2 Then suggestedStart = 2
    For j = suggestedStart To messages.Count
        If JsonExtractString(CStr(messages(j)), "role") = "user" Then
            prev = JsonExtractString(CStr(messages(j - 1)), "role")
            If prev <> "tool" And prev <> "" Then
                If Not (prev = "assistant" And JsonHasToolCalls(CStr(messages(j - 1)))) Then
                    SafeBoundaryIndex = j
                    Exit Function
                End If
            End If
        End If
    Next j
End Function

Public Sub Session_Reset()
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(SessionFilePath()) > 0 Then
        If fso.FileExists(SessionFilePath()) Then fso.DeleteFile SessionFilePath(), True
    End If
    On Error GoTo 0
    Ops_WriteProgress "新しい会話を開始しました（履歴をクリア）", False
End Sub
