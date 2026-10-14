Attribute VB_Name = "modSkills"
Option Explicit

' User-editable skills on the "スキル" sheet (no rebuild needed).
' Columns: A=スキル名 B=種別(prompt|vba) C=説明 D=本文(prompt)またはVBA関数名(vba)
'          E=パラメータJSONスキーマ(vba用, 省略可) F=書込確認(yes/no) G=有効(yes/no)
'
' prompt skill: injected into the system prompt so the agent follows its steps.
' vba skill: exposed as tool "vba_<スキル名>" calling Application.Run(関数名, argsJson).
' Skill names become tool names: use full-width-safe simple names; they are
' lowercased and non-word characters are replaced with "_".

Private Const MAX_SKILLS As Long = 20
Private Const MAX_BODY_CHARS As Long = 1200

Public Function Skills_PromptSection() As String
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim section As String
    Dim count As Long
    Dim kind As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_SKILLS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If count >= MAX_SKILLS Then Exit For
        If Not Enabled(ws, r) Then GoTo NextRow
        kind = LCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        If kind = "prompt" Then
            count = count + 1
            section = section & "- スキル「" & SkillToolName(CStr(ws.Cells(r, 1).Value)) & "」: " _
                & Trim$(CStr(ws.Cells(r, 3).Value)) & " | 手順: " _
                & Left$(Trim$(CStr(ws.Cells(r, 4).Value)), MAX_BODY_CHARS) & vbCrLf
        End If
NextRow:
    Next r

    If Len(section) > 0 Then
        Skills_PromptSection = vbCrLf & "ユーザー定義スキル: ユーザーの依頼がいずれかに該当する場合、その手順を必ず従ってください。" & vbCrLf & section
    End If
End Function

Public Function Skills_ToolDefsOpenAI() As String
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim name As String
    Dim kind As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_SKILLS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Not Enabled(ws, r) Then GoTo NextRow
        kind = LCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        If kind = "vba" Then
            name = SkillToolName(CStr(ws.Cells(r, 1).Value))
            If Len(Skills_ToolDefsOpenAI) > 0 Then Skills_ToolDefsOpenAI = Skills_ToolDefsOpenAI & ","
            Skills_ToolDefsOpenAI = Skills_ToolDefsOpenAI & ToolDef(name, _
                "ユーザー定義スキル「" & CStr(ws.Cells(r, 1).Value) & "」: " & CStr(ws.Cells(r, 3).Value) & " (VBA関数 " & CStr(ws.Cells(r, 4).Value) & " を実行)", _
                ParamsOrDefault(ws.Cells(r, 5).Value))
        End If
NextRow:
    Next r
End Function

Public Function Skills_ToolDefsAnthropic() As String
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim name As String
    Dim kind As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_SKILLS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Not Enabled(ws, r) Then GoTo NextRow
        kind = LCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        If kind = "vba" Then
            name = SkillToolName(CStr(ws.Cells(r, 1).Value))
            If Len(Skills_ToolDefsAnthropic) > 0 Then Skills_ToolDefsAnthropic = Skills_ToolDefsAnthropic & ","
            Skills_ToolDefsAnthropic = Skills_ToolDefsAnthropic & AnthropicToolDef(name, _
                "ユーザー定義スキル「" & CStr(ws.Cells(r, 1).Value) & "」: " & CStr(ws.Cells(r, 3).Value), _
                ParamsOrDefault(ws.Cells(r, 5).Value))
        End If
NextRow:
    Next r
End Function

Public Function Skills_WriteToolNames() As String
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim kind As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_SKILLS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Not Enabled(ws, r) Then GoTo NextRow
        kind = LCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        If kind = "vba" And IsYes(ws.Cells(r, 6).Value) Then
            If Len(Skills_WriteToolNames) > 0 Then Skills_WriteToolNames = Skills_WriteToolNames & ","
            Skills_WriteToolNames = Skills_WriteToolNames & SkillToolName(CStr(ws.Cells(r, 1).Value))
        End If
NextRow:
    Next r
    Skills_WriteToolNames = LCase$(Skills_WriteToolNames)
End Function

Public Function Skills_TryExecute(ByVal toolName As String, ByVal argsJson As String, _
                                  ByRef autoApproved As Boolean, ByRef handled As Boolean) As String
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim target As String
    Dim name As String

    handled = False
    toolName = LCase$(toolName)
    If Left$(toolName, 4) <> "vba_" Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_SKILLS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Not Enabled(ws, r) Then GoTo NextRow
        If LCase$(Trim$(CStr(ws.Cells(r, 2).Value))) <> "vba" Then GoTo NextRow
        name = SkillToolName(CStr(ws.Cells(r, 1).Value))
        If name = toolName Then
            handled = True
            target = Trim$(CStr(ws.Cells(r, 4).Value))
            If IsYes(ws.Cells(r, 6).Value) Then
                If Not ApproveWriteTool(toolName, argsJson, autoApproved) Then
                    Skills_TryExecute = "ユーザーが実行をキャンセルしました"
                    Exit Function
                End If
            End If
            On Error GoTo SkillErr
            Skills_TryExecute = Application.Run(target, argsJson)
            Exit Function
        End If
NextRow:
    Next r
    Exit Function
SkillErr:
    Skills_TryExecute = "スキル '" & toolName & "' (関数 " & target & ") でエラー: " & Err.Description
End Function

' User macros for the sheet toolbar.
Public Sub Skills_New()
    Dim ws As Worksheet
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_SKILLS)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = "new-skill"
    ws.Cells(r, 2).Value = "prompt"
    ws.Cells(r, 3).Value = "説明を書く"
    ws.Cells(r, 4).Value = "手順を書く (種別prompt) または VBA関数名を書く (種別vba)"
    ws.Cells(r, 7).Value = "no"
    ws.Cells(r, 1).Select
End Sub

Private Function Enabled(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Enabled = IsYes(ws.Cells(r, 7).Value)
End Function

Private Function IsYes(ByVal v As Variant) As Boolean
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    IsYes = (s = "yes" Or s = "はい" Or s = "y" Or s = "true" Or s = "1")
End Function

Private Function ParamsOrDefault(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) > 0 Then
        ParamsOrDefault = s
    Else
        ParamsOrDefault = "{""input"":{""type"":""string"",""description"":""JSON-formatted arguments for the VBA function""}}"
    End If
End Function

Public Function SkillToolName(ByVal rawName As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    For i = 1 To Len(rawName)
        ch = Mid$(rawName, i, 1)
        Select Case ch
            Case "a" To "z", "0" To "9", "_": out = out & ch
            Case "A" To "Z": out = out & LCase$(ch)
            Case " ": out = out & "_"
            Case Else: out = out & "_"
        End Select
    Next i
    Do While InStr(out, "__") > 0
        out = Replace(out, "__", "_")
    Loop
    out = TrimChar(out, "_")
    If Len(out) = 0 Then out = "skill"
    SkillToolName = "vba_" & out
End Function

Private Function TrimChar(ByVal s As String, ByVal ch As String) As String
    Do While Left$(s, 1) = ch
        s = Mid$(s, 2)
    Loop
    Do While Right$(s, 1) = ch
        s = Left$(s, Len(s) - 1)
    Loop
    TrimChar = s
End Function
