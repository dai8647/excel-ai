Attribute VB_Name = "modTools"
Option Explicit

Public Function Tools_GetDefinitionsJson() As String
    Dim extra As String
    extra = Skills_ToolDefsOpenAI() & Plugins_ToolsOpenAI()
    If Len(extra) > 0 Then extra = "," & extra
    Tools_GetDefinitionsJson = "[" & _
        ToolDef("list_dir", "List files and folders in a directory", "{""path"":{""type"":""string"",""description"":""Relative or absolute path within workspace""},""recursive"":{""type"":""boolean"",""description"":""Recurse into subfolders up to 3 levels (capped at 400 lines)""}}") & "," & _
        ToolDef("read_file", "Read a text file in pages (default 100KB; use max_chars and start_chars for larger files)", "{""path"":{""type"":""string""},""max_chars"":{""type"":""integer"",""description"":""1000-200000 characters""},""start_chars"":{""type"":""integer"",""description"":""Character offset for paging""}}") & "," & _
        ToolDef("write_file", "Write or overwrite a text file", "{""path"":{""type"":""string""},""content"":{""type"":""string""}}") & "," & _
        ToolDef("append_file", "Append text to a file", "{""path"":{""type"":""string""},""content"":{""type"":""string""}}") & "," & _
        ToolDef("move_file", "Move or rename a file/folder", "{""source"":{""type"":""string""},""destination"":{""type"":""string""}}") & "," & _
        ToolDef("delete_file", "Delete a file (requires confirmation setting)", "{""path"":{""type"":""string""}}") & "," & _
        ToolDef("merge_csv", "Merge CSV files matching a pattern into one file", "{""source_pattern"":{""type"":""string"",""description"":""e.g. data\\*.csv""},""output_path"":{""type"":""string""}}") & "," & _
        ToolDef("create_excel", "Create a new Excel workbook from CSV-like data", "{""path"":{""type"":""string"",""description"":""Output .xlsx path""},""sheet_name"":{""type"":""string""},""csv_content"":{""type"":""string"",""description"":""Comma-separated rows, newline-separated lines""}}") & "," & _
        ToolDef("create_chart", "Add a chart to an existing Excel workbook", "{""path"":{""type"":""string"",""description"":""Existing .xlsx or .xlsm path""},""sheet_name"":{""type"":""string""},""source_range"":{""type"":""string"",""description"":""Source range such as A1:D4""},""chart_type"":{""type"":""string"",""description"":""column, bar, line, or pie""},""title"":{""type"":""string""}}") & "," & _
        ToolDef("inspect_excel", "Inspect an .xlsx/.xlsm workbook without reading its binary contents", "{""path"":{""type"":""string""},""sheet_name"":{""type"":""string"",""description"":""Optional sheet name""},""include_formulas"":{""type"":""boolean"",""description"":""Include formulas in the sample""}}") & "," & _
        ToolDef("grep_files", "Search for text in files under a folder", "{""folder"":{""type"":""string""},""pattern"":{""type"":""string""}}") & "," & _
        ToolDef("update_excel", "Add a sheet to an existing .xlsx/.xlsm and write values or formulas. Either csv_content (CSV block starting at start_cell) or edits (newline-separated CELL=VALUE lines). Values starting with = become formulas.", "{""path"":{""type"":""string""},""sheet_name"":{""type"":""string""},""new_sheet"":{""type"":""string"",""description"":""Optional: create this sheet if missing and write into it""},""start_cell"":{""type"":""string"",""description"":""Top-left cell for csv_content, default A1""},""csv_content"":{""type"":""string""},""edits"":{""type"":""string"",""description"":""Lines like B2==SUM(B2:B9) or Name=apple, max 500""}}") & "," & _
        ToolDef("make_dir", "Create a folder inside the workspace (parents created too)", "{""path"":{""type"":""string""}}") & "," & _
        ToolDef("copy_file", "Deep-copy a file or folder to a new path inside the workspace", "{""source"":{""type"":""string""},""destination"":{""type"":""string""}}") & "," & _
        ToolDef("open_file", "Open a workspace file in its default application so the user can see the result", "{""path"":{""type"":""string""}}") & "," & _
        ToolDef("web_fetch", "Fetch a public web page or API response as text (GET only; private and local addresses are blocked)", "{""url"":{""type"":""string"",""description"":""Public http(s) URL""},""max_chars"":{""type"":""integer"",""description"":""Maximum returned characters, 1000-50000""}}") & _
        extra & "]"
End Function

Public Function Tools_GetAnthropicDefinitionsJson() As String
    Dim extra As String
    extra = Skills_ToolDefsAnthropic() & Plugins_ToolsAnthropic()
    If Len(extra) > 0 Then extra = "," & extra
    Tools_GetAnthropicDefinitionsJson = "[" & _
        AnthropicToolDef("list_dir", "List files and folders in a directory", "{""path"":{""type"":""string"",""description"":""Relative or absolute path within workspace""}}") & "," & _
        AnthropicToolDef("read_file", "Read a text file in pages (default 100KB; use max_chars and start_chars for larger files)", "{""path"":{""type"":""string""},""max_chars"":{""type"":""integer"",""description"":""1000-200000 characters""},""start_chars"":{""type"":""integer"",""description"":""Character offset for paging""}}") & "," & _
        AnthropicToolDef("write_file", "Write or overwrite a text file", "{""path"":{""type"":""string""},""content"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("append_file", "Append text to a file", "{""path"":{""type"":""string""},""content"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("move_file", "Move or rename a file/folder", "{""source"":{""type"":""string""},""destination"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("delete_file", "Delete a file (requires confirmation setting)", "{""path"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("merge_csv", "Merge CSV files matching a pattern into one file", "{""source_pattern"":{""type"":""string""},""output_path"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("create_excel", "Create a new Excel workbook from CSV-like data", "{""path"":{""type"":""string""},""sheet_name"":{""type"":""string""},""csv_content"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("create_chart", "Add a chart to an existing Excel workbook", "{""path"":{""type"":""string""},""sheet_name"":{""type"":""string""},""source_range"":{""type"":""string""},""chart_type"":{""type"":""string""},""title"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("inspect_excel", "Inspect an .xlsx/.xlsm workbook without reading its binary contents", "{""path"":{""type"":""string""},""sheet_name"":{""type"":""string""},""include_formulas"":{""type"":""boolean""}}") & "," & _
        AnthropicToolDef("grep_files", "Search for text in files under a folder", "{""folder"":{""type"":""string""},""pattern"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("update_excel", "Add a sheet to an existing .xlsx/.xlsm and write values or formulas. Either csv_content (CSV block starting at start_cell) or edits (newline-separated CELL=VALUE lines). Values starting with = become formulas.", "{""path"":{""type"":""string""},""sheet_name"":{""type"":""string""},""new_sheet"":{""type"":""string""},""start_cell"":{""type"":""string""},""csv_content"":{""type"":""string""},""edits"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("make_dir", "Create a folder inside the workspace (parents created too)", "{""path"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("copy_file", "Deep-copy a file or folder to a new path inside the workspace", "{""source"":{""type"":""string""},""destination"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("open_file", "Open a safe file in its default application", "{""path"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("web_fetch", "Fetch a public web page or API response as text (GET only; private and local addresses are blocked)", "{""url"":{""type"":""string"",""description"":""Public http(s) URL""},""max_chars"":{""type"":""integer"",""description"":""Maximum returned characters, 1000-50000""}}") & extra & "]"
End Function

Public Function ToolDef(ByVal name As String, ByVal description As String, ByVal params As String) As String
    ToolDef = "{""type"":""function"",""function"":{" & _
        """name"":" & JsonString(name) & "," & _
        """description"":" & JsonString(description) & "," & _
        """parameters"":{""type"":""object"",""properties"":" & params & "}}}"
End Function

Public Function AnthropicToolDef(ByVal name As String, ByVal description As String, ByVal params As String) As String
    AnthropicToolDef = "{""name"":" & JsonString(name) & ",""description"":" & JsonString(description) & _
        ",""input_schema"":{""type"":""object"",""properties"":" & params & "}}"
End Function

Public Function Tools_Execute(ByVal toolName As String, ByVal argsJson As String, ByRef autoApproved As Boolean) As String
    On Error GoTo ErrHandler

    Select Case LCase$(toolName)
        Case "list_dir"
            Tools_Execute = Tool_ListDir(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "recursive"))
        Case "read_file"
            Tools_Execute = Tool_ReadFile(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "max_chars"), JsonGetArg(argsJson, "start_chars"))
        Case "write_file"
            If Not ApproveWriteTool("write_file", argsJson, autoApproved) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
            End If
            Tools_Execute = Tool_WriteFile(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "content"))
        Case "append_file"
            If Not ApproveWriteTool("append_file", argsJson, autoApproved) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
            End If
            Tools_Execute = Tool_AppendFile(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "content"))
        Case "move_file"
            If Not ApproveWriteTool("move_file", argsJson, autoApproved) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
            End If
            Tools_Execute = Tool_MoveFile(JsonGetArg(argsJson, "source"), JsonGetArg(argsJson, "destination"))
        Case "delete_file"
            If IsConfirmDangerEnabled() And Not autoApproved Then
                If Not ConfirmToolAction("delete_file", argsJson) Then
                    Tools_Execute = "ユーザーが削除をキャンセルしました"
                    Exit Function
                End If
                autoApproved = True
            ElseIf Not ApproveWriteTool("delete_file", argsJson, autoApproved) Then
                Tools_Execute = "ユーザーが削除をキャンセルしました"
                Exit Function
            End If
            Tools_Execute = Tool_DeleteFile(JsonGetArg(argsJson, "path"))
        Case "merge_csv"
            If Not ApproveWriteTool("merge_csv", argsJson, autoApproved) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
            End If
            Tools_Execute = Tool_MergeCsv(JsonGetArg(argsJson, "source_pattern"), JsonGetArg(argsJson, "output_path"))
        Case "create_excel"
            If Not ApproveWriteTool("create_excel", argsJson, autoApproved) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
            End If
            Tools_Execute = Tool_CreateExcel(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "sheet_name"), JsonGetArg(argsJson, "csv_content"))
        Case "create_chart"
            If Not ApproveWriteTool("create_chart", argsJson, autoApproved) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
            End If
            Tools_Execute = Tool_CreateChart(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "sheet_name"), JsonGetArg(argsJson, "source_range"), JsonGetArg(argsJson, "chart_type"), JsonGetArg(argsJson, "title"))
        Case "inspect_excel"
            Tools_Execute = Tool_InspectExcel(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "sheet_name"), JsonGetArg(argsJson, "include_formulas"))
        Case "update_excel"
            If Not ApproveWriteTool("update_excel", argsJson, autoApproved) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
            End If
            Tools_Execute = Tool_UpdateExcel(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "sheet_name"), JsonGetArg(argsJson, "new_sheet"), JsonGetArg(argsJson, "start_cell"), JsonGetArg(argsJson, "csv_content"), JsonGetArg(argsJson, "edits"))
        Case "make_dir"
            If Not ApproveWriteTool("make_dir", argsJson, autoApproved) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
            End If
            Tools_Execute = Tool_MakeDir(JsonGetArg(argsJson, "path"))
        Case "copy_file"
            If Not ApproveWriteTool("copy_file", argsJson, autoApproved) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
            End If
            Tools_Execute = Tool_CopyFile(JsonGetArg(argsJson, "source"), JsonGetArg(argsJson, "destination"))
        Case "open_file"
            Tools_Execute = Tool_OpenFile(JsonGetArg(argsJson, "path"))
        Case "grep_files"
            Tools_Execute = Tool_GrepFiles(JsonGetArg(argsJson, "folder"), JsonGetArg(argsJson, "pattern"))
        Case "web_fetch"
            Tools_Execute = Tool_WebFetch(JsonGetArg(argsJson, "url"), JsonGetArg(argsJson, "max_chars"))
        Case Else
            Dim handled As Boolean
            Tools_Execute = Skills_TryExecute(toolName, argsJson, autoApproved, handled)
            If handled Then Exit Function
            Tools_Execute = Plugins_TryExecute(toolName, argsJson, autoApproved, handled)
            If handled Then Exit Function
            Tools_Execute = "未知のツール: " & toolName
    End Select
    Exit Function
ErrHandler:
    Tools_Execute = "エラー: " & Err.Description
End Function

Private Function ConfirmToolAction(ByVal toolName As String, ByVal argsJson As String) As Boolean
    ConfirmToolAction = (MsgBox("ツール実行の確認" & vbCrLf & vbCrLf & _
        "ツール: " & toolName & vbCrLf & "引数: " & Left$(argsJson, 500) & vbCrLf & vbCrLf & _
        "実行しますか？", vbYesNo + vbQuestion, "ExcelAI") = vbYes)
End Function

Public Function ApproveWriteTool(ByVal toolName As String, ByVal argsJson As String, ByRef autoApproved As Boolean) As Boolean
    If autoApproved Or IsAutoRunEnabled() Then
        autoApproved = True
        ApproveWriteTool = True
        FireBeforeWriteHook argsJson
        Exit Function
    End If

    If ConfirmToolAction(toolName, argsJson) Then
        autoApproved = True
        ApproveWriteTool = True
        FireBeforeWriteHook argsJson
    Else
        ApproveWriteTool = False
    End If
End Function

' Optional plugin hook: notify plugins before a gated write touches a path.
' Extracts the write target from the tool arguments; unresolvable paths are skipped.
Public Sub FireBeforeWriteHook(ByVal argsJson As String)
    Dim candidate As String
    Dim resolved As String

    candidate = JsonGetArg(argsJson, "path")
    If Len(candidate) = 0 Then candidate = JsonGetArg(argsJson, "source")
    If Len(candidate) = 0 Then Exit Sub
    On Error Resume Next
    resolved = ResolveSafePath(candidate)
    On Error GoTo 0
    If Len(resolved) = 0 Then Exit Sub
    Plugins_FireBeforeWrite resolved
End Sub

Private Function Tool_ListDir(ByVal path As String, ByVal recursiveText As String) As String
    Dim fso As Object
    Dim safePath As String
    Dim result As String
    Dim count As Long
    Dim recursive As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(path) = 0 Then path = "."
    safePath = ResolveSafePath(path)

    If Not fso.FolderExists(safePath) Then
        Tool_ListDir = "フォルダが存在しません: " & safePath
        Exit Function
    End If

    recursive = (LCase$(Trim$(recursiveText)) = "true" Or LCase$(Trim$(recursiveText)) = "1" Or Trim$(recursiveText) = "はい")
    If recursive Then
        ' FSO-based walk: Dir$ is a single shared cursor and cannot nest.
        ListFolderRecursive fso, safePath, "", result, count, 1
        Tool_ListDir = result
        Exit Function
    End If

    ListFolderFlat fso, safePath, "", result, count

    If Len(result) = 0 Then result = "(空のフォルダ)"
    Tool_ListDir = result
End Function

Private Sub ListFolderFlat(ByVal fso As Object, ByVal folderPath As String, ByVal indent As String, ByRef result As String, ByRef count As Long)
    Dim sub_ As Object
    Dim f As Object
    For Each sub_ In fso.GetFolder(folderPath).SubFolders
        result = result & indent & "[DIR]  " & sub_.Name & vbCrLf
        count = count + 1
    Next sub_
    For Each f In fso.GetFolder(folderPath).Files
        result = result & indent & "[FILE] " & f.Name & " (" & f.Size & " bytes)" & vbCrLf
        count = count + 1
    Next f
End Sub

Private Sub ListFolderRecursive(ByVal fso As Object, ByVal folderPath As String, ByVal indent As String, ByRef result As String, ByRef count As Long, ByVal depth As Long)
    Dim sub_ As Object
    Dim f As Object

    If count >= 400 Then Exit Sub
    For Each sub_ In fso.GetFolder(folderPath).SubFolders
        If count >= 400 Then Exit Sub
        result = result & indent & "[DIR]  " & sub_.Name & vbCrLf
        count = count + 1
        If depth < 3 Then ListFolderRecursive fso, sub_.Path, indent & "  ", result, count, depth + 1
    Next sub_
    For Each f In fso.GetFolder(folderPath).Files
        If count >= 400 Then Exit Sub
        result = result & indent & "[FILE] " & f.Name & " (" & f.Size & " bytes)" & vbCrLf
        count = count + 1
    Next f
End Sub

Private Function Tool_MakeDir(ByVal path As String) As String
    Dim fso As Object
    Dim safePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    safePath = ResolveSafePath(path)
    If fso.FolderExists(safePath) Then
        Tool_MakeDir = "既に存在します: " & safePath
        Exit Function
    End If
    EnsureFolder fso, safePath
    Tool_MakeDir = "フォルダ作成完了: " & safePath
End Function

Private Sub EnsureFolder(ByVal fso As Object, ByVal folderPath As String)
    Dim parent As String
    If fso.FolderExists(folderPath) Then Exit Sub
    parent = fso.GetParentFolderName(folderPath)
    If Len(parent) > 0 And parent <> folderPath And Not fso.FolderExists(parent) Then
        EnsureFolder fso, parent
    End If
    fso.CreateFolder folderPath
End Sub

Private Function Tool_CopyFile(ByVal source As String, ByVal destination As String) As String
    Dim fso As Object
    Dim srcPath As String
    Dim dstPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    srcPath = ResolveSafePath(source)
    dstPath = ResolveSafePath(destination)

    If fso.FileExists(srcPath) Then
        fso.CopyFile srcPath, dstPath, True
        Tool_CopyFile = "コピー完了: " & srcPath & " -> " & dstPath
    ElseIf fso.FolderExists(srcPath) Then
        fso.CopyFolder srcPath, dstPath, True
        Tool_CopyFile = "フォルダコピー完了: " & srcPath & " -> " & dstPath
    Else
        Tool_CopyFile = "コピー元が存在しません: " & srcPath
    End If
End Function

Private Function Tool_OpenFile(ByVal path As String) As String
    Dim fso As Object
    Dim safePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    safePath = ResolveSafePath(path)
    If Not fso.FileExists(safePath) And Not fso.FolderExists(safePath) Then
        Tool_OpenFile = "存在しません: " & safePath
        Exit Function
    End If
    Shell "explorer.exe """ & safePath & """", vbNormalFocus
    Tool_OpenFile = "開きました: " & safePath
End Function

Private Function Tool_UpdateExcel(ByVal path As String, ByVal sheetName As String, ByVal newSheet As String, ByVal startCell As String, ByVal csvContent As String, ByVal edits As String) As String
    Dim safePath As String
    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim targetCell As Object
    Dim lines() As String
    Dim cells() As String
    Dim eqPos As Long
    Dim r As Long, c As Long, i As Long
    Dim editCount As Long
    Dim rowOff As Long
    Dim resultText As String

    On Error GoTo ErrHandler
    safePath = ResolveSafePath(path)
    If LCase$(Right$(safePath, 5)) <> ".xlsx" And LCase$(Right$(safePath, 5)) <> ".xlsm" Then
        Err.Raise vbObjectError + 2102, "update_excel", "対象は .xlsx/.xlsm を指定してください"
    End If
    If Len(csvContent) = 0 And Len(edits) = 0 Then
        Err.Raise vbObjectError + 2103, "update_excel", "csv_content か edits のどちらかを指定してください"
    End If
    If Len(startCell) = 0 Then startCell = "A1"

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(safePath)

    If Len(newSheet) > 0 Then
        On Error Resume Next
        Set ws = wb.Worksheets(newSheet)
        On Error GoTo ErrHandler
        If ws Is Nothing Then
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = Left$(newSheet, 31)
        End If
    ElseIf Len(sheetName) > 0 Then
        Set ws = wb.Worksheets(sheetName)
    Else
        Set ws = wb.Worksheets(1)
    End If

    If Len(csvContent) > 0 Then
        Set targetCell = ws.Range(startCell)
        rowOff = targetCell.Row - 1
        Dim colOff As Long
        colOff = targetCell.Column - 1
        lines = Split(csvContent, vbLf)
        r = 0
        For i = LBound(lines) To UBound(lines)
            Dim line As String
            line = Replace(lines(i), vbCr, "")
            If Len(line) > 0 Then
                cells = Split(line, ",")
                For c = LBound(cells) To UBound(cells)
                    ws.Cells(rowOff + r + 1, colOff + c + 1).Value = cells(c)
                    editCount = editCount + 1
                Next c
                r = r + 1
            End If
            If editCount >= 50000 Then Exit For
        Next i
    End If

    If Len(edits) > 0 Then
        lines = Split(edits, vbLf)
        For i = LBound(lines) To UBound(lines)
            line = Replace(lines(i), vbCr, "")
            eqPos = InStr(line, "=")
            If eqPos > 1 Then
                ws.Range(Trim$(Left$(line, eqPos - 1))).Value = Mid$(line, eqPos + 1)
                editCount = editCount + 1
            End If
            If editCount >= 50000 Then Exit For
        Next i
    End If

    wb.Save
    resultText = "Excel更新完了: " & safePath & " / " & ws.Name & " / " & editCount & " セル"
    GoTo Cleanup

ErrHandler:
    resultText = "update_excelエラー: " & Err.Description

Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set targetCell = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing
    Tool_UpdateExcel = resultText
End Function

Private Function Tool_ReadFile(ByVal path As String, Optional ByVal maxCharsText As String = "", Optional ByVal startCharsText As String = "") As String
    Const DEFAULT_MAX_CHARS As Long = 100000
    Const ABSOLUTE_MAX_CHARS As Long = 200000
    Const MAX_TEXT_FILE_BYTES As Long = 5242880
    Dim fso As Object
    Dim stream As Object
    Dim safePath As String
    Dim content As String
    Dim extension As String
    Dim maxChars As Long
    Dim startChars As Long
    Dim chunk As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    safePath = ResolveSafePath(path)

    If Not fso.FileExists(safePath) Then
        Tool_ReadFile = "ファイルが存在しません: " & safePath
        Exit Function
    End If

    extension = LCase$(fso.GetExtensionName(safePath))
    If extension = "xlsx" Or extension = "xlsm" Or extension = "xlsb" Or extension = "xls" Then
        Tool_ReadFile = "Excelバイナリファイルです。read_fileではなくinspect_excelを使ってください: " & safePath
        Exit Function
    ElseIf extension = "pdf" Or extension = "docx" Or extension = "zip" Then
        Tool_ReadFile = "バイナリファイル形式はread_fileでは直接読めません: " & safePath
        Exit Function
    End If

    If fso.GetFile(safePath).Size > MAX_TEXT_FILE_BYTES Then
        Tool_ReadFile = "テキストファイルが大きすぎます (5MB上限): " & safePath
        Exit Function
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile safePath
    content = stream.ReadText
    stream.Close

    maxChars = Val(maxCharsText)
    If maxChars <= 0 Then maxChars = DEFAULT_MAX_CHARS
    If maxChars < 1000 Then maxChars = 1000
    If maxChars > ABSOLUTE_MAX_CHARS Then maxChars = ABSOLUTE_MAX_CHARS
    startChars = Val(startCharsText)
    If startChars < 0 Then startChars = 0
    If startChars >= Len(content) Then
        Tool_ReadFile = "ファイル末尾です。start_chars=" & startChars & ", length=" & Len(content)
        Exit Function
    End If

    chunk = Mid$(content, startChars + 1, maxChars)
    Tool_ReadFile = "FILE: " & safePath & vbCrLf & _
        "CHARS: " & Len(content) & ", start_chars=" & startChars & ", returned=" & Len(chunk) & vbCrLf & vbCrLf & chunk
    If startChars + Len(chunk) < Len(content) Then
        Tool_ReadFile = Tool_ReadFile & vbCrLf & vbCrLf & "[続きあり: start_chars=" & (startChars + Len(chunk)) & "]"
    End If
End Function

Private Function Tool_WriteFile(ByVal path As String, ByVal content As String) As String
    Dim fso As Object
    Dim stream As Object
    Dim safePath As String
    Dim parent As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    safePath = ResolveSafePath(path)
    parent = fso.GetParentFolderName(safePath)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then fso.CreateFolder parent

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile safePath, 2
    stream.Close

    Tool_WriteFile = "書き込み完了: " & safePath
End Function

Private Function Tool_AppendFile(ByVal path As String, ByVal content As String) As String
    Dim existing As String
    If Len(path) = 0 Then
        Tool_AppendFile = "パスが必要です"
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(ResolveSafePath(path)) Then
        existing = Tool_ReadFile(path)
    End If
    Tool_AppendFile = Tool_WriteFile(path, existing & content)
End Function

Private Function Tool_MoveFile(ByVal source As String, ByVal destination As String) As String
    Dim fso As Object
    Dim srcPath As String
    Dim dstPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    srcPath = ResolveSafePath(source)
    dstPath = ResolveSafePath(destination)

    If fso.FileExists(srcPath) Then
        fso.MoveFile srcPath, dstPath
        Tool_MoveFile = "移動完了: " & srcPath & " -> " & dstPath
    ElseIf fso.FolderExists(srcPath) Then
        fso.MoveFolder srcPath, dstPath
        Tool_MoveFile = "フォルダ移動完了: " & srcPath & " -> " & dstPath
    Else
        Tool_MoveFile = "ソースが存在しません: " & srcPath
    End If
End Function

Private Function Tool_DeleteFile(ByVal path As String) As String
    Dim fso As Object
    Dim safePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    safePath = ResolveSafePath(path)

    If fso.FileExists(safePath) Then
        fso.DeleteFile safePath, True
        Tool_DeleteFile = "削除完了: " & safePath
    Else
        Tool_DeleteFile = "ファイルが存在しません: " & safePath
    End If
End Function

Private Function Tool_MergeCsv(ByVal sourcePattern As String, ByVal outputPath As String) As String
    Dim fso As Object
    Dim folderPath As String
    Dim fileSpec As String
    Dim folder As Object
    Dim f As Object
    Dim outPath As String
    Dim header As String
    Dim headerWritten As Boolean
    Dim line As String
    Dim stream As Object
    Dim count As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    If InStr(sourcePattern, "\") > 0 Then
        folderPath = ResolveSafePath(Left$(sourcePattern, InStrRev(sourcePattern, "\") - 1))
        fileSpec = Mid$(sourcePattern, InStrRev(sourcePattern, "\") + 1)
    Else
        folderPath = GetWorkspacePath()
        fileSpec = sourcePattern
    End If

    outPath = ResolveSafePath(outputPath)
    If Not fso.FolderExists(folderPath) Then
        Tool_MergeCsv = "フォルダが存在しません: " & folderPath
        Exit Function
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open

    Set folder = fso.GetFolder(folderPath)
    headerWritten = False
    count = 0

    For Each f In folder.Files
        If f.Name Like fileSpec Then
            Dim content As String
            content = Tool_ReadFile(fso.BuildPath(folderPath, f.Name))
            If Len(content) > 0 Then
                Dim lines() As String
                lines = Split(content, vbLf)
                Dim i As Long
                For i = LBound(lines) To UBound(lines)
                    line = Replace(lines(i), vbCr, "")
                    If Len(line) = 0 Then GoTo NextLine
                    If Not headerWritten Then
                        header = line
                        stream.WriteText line & vbCrLf
                        headerWritten = True
                    ElseIf i = LBound(lines) And line = header Then
                        ' skip duplicate header
                    Else
                        stream.WriteText line & vbCrLf
                    End If
NextLine:
                Next i
                count = count + 1
            End If
        End If
    Next f

    stream.SaveToFile outPath, 2
    stream.Close

    Tool_MergeCsv = "CSV結合完了: " & count & " ファイル -> " & outPath
End Function

Private Function Tool_CreateExcel(ByVal path As String, ByVal sheetName As String, ByVal csvContent As String) As String
    Dim safePath As String
    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim lines() As String
    Dim cells() As String
    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    Dim resultText As String

    On Error GoTo ErrHandler
    safePath = ResolveSafePath(path)
    If LCase$(Right$(safePath, 5)) <> ".xlsx" And LCase$(Right$(safePath, 5)) <> ".xlsm" Then
        safePath = safePath & ".xlsx"
    End If

    If Len(sheetName) = 0 Then sheetName = "Sheet1"

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = Left$(sheetName, 31)

    lines = Split(csvContent, vbLf)
    r = 1
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Replace(lines(i), vbCr, "")
        If Len(line) > 0 Then
            cells = Split(line, ",")
            For c = LBound(cells) To UBound(cells)
                ws.Cells(r, c + 1).Value = cells(c)
                If c + 1 > lastCol Then lastCol = c + 1
            Next c
            r = r + 1
        End If
    Next i

    lastRow = r - 1
    If lastRow > 0 And lastCol > 0 Then
        With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .RowHeight = 24
        End With
        With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
            .Borders.LineStyle = 1
            .VerticalAlignment = -4108
            .Columns.AutoFit
        End With
        For c = 1 To lastCol
            If ws.Columns(c).ColumnWidth > 32 Then ws.Columns(c).ColumnWidth = 32
        Next c
    End If

    wb.SaveAs safePath
    resultText = "Excel作成完了: " & safePath
    GoTo Cleanup

ErrHandler:
    resultText = "create_excelエラー: " & Err.Description

Cleanup:
    ' Always close the hidden Excel instance so a mid-way failure cannot
    ' leave a zombie EXCEL.EXE behind.
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing
    Tool_CreateExcel = resultText
End Function

Private Function Tool_CreateChart(ByVal path As String, ByVal sheetName As String, ByVal sourceRange As String, ByVal chartType As String, ByVal title As String) As String
    Dim safePath As String
    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim chartObject As Object
    Dim chart As Object
    Dim resultText As String

    On Error GoTo ErrHandler
    safePath = ResolveSafePath(path)
    If LCase$(Right$(safePath, 5)) <> ".xlsx" And LCase$(Right$(safePath, 5)) <> ".xlsm" Then
        Err.Raise vbObjectError + 2101, "create_chart", "Excelファイル(.xlsx/.xlsm)を指定してください"
    End If
    If Len(sheetName) = 0 Then sheetName = "Sheet1"
    If Len(sourceRange) = 0 Then sourceRange = "A1:B10"

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(safePath)
    Set ws = wb.Worksheets(sheetName)
    xl.CalculateFull

    Set chartObject = ws.ChartObjects.Add(320, 20, 480, 280)
    Set chart = chartObject.Chart
    chart.SetSourceData ws.Range(sourceRange)
    chart.ChartType = ChartTypeValue(chartType)
    If Len(title) > 0 Then
        chart.HasTitle = True
        chart.ChartTitle.Text = title
    End If
    chart.HasLegend = True
    chart.Legend.Position = -4107

    wb.Save
    resultText = "グラフ作成完了: " & safePath
    GoTo Cleanup

ErrHandler:
    resultText = "create_chartエラー: " & Err.Description

Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set chart = Nothing
    Set chartObject = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing
    Tool_CreateChart = resultText
End Function

Private Function ChartTypeValue(ByVal chartType As String) As Long
    Select Case LCase$(Trim$(chartType))
        Case "bar", "横棒": ChartTypeValue = 57
        Case "line", "折れ線": ChartTypeValue = 4
        Case "pie", "円": ChartTypeValue = 5
        Case Else: ChartTypeValue = 51
    End Select
End Function

Private Function Tool_InspectExcel(ByVal path As String, ByVal sheetName As String, ByVal includeFormulasText As String) As String
    Dim safePath As String
    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim used As Object
    Dim sh As Object
    Dim result As String
    Dim sampleRows As Long, sampleCols As Long
    Dim r As Long, c As Long
    Dim valueText As String
    Dim includeFormulas As Boolean

    On Error GoTo ErrHandler
    safePath = ResolveSafePath(path)
    If LCase$(Right$(safePath, 5)) <> ".xlsx" And LCase$(Right$(safePath, 5)) <> ".xlsm" And LCase$(Right$(safePath, 5)) <> ".xlsb" And LCase$(Right$(safePath, 4)) <> ".xls" Then
        Err.Raise vbObjectError + 2301, "inspect_excel", "Excelファイル(.xlsx/.xlsm/.xlsb/.xls)を指定してください"
    End If

    includeFormulas = (LCase$(Trim$(includeFormulasText)) = "true" Or LCase$(Trim$(includeFormulasText)) = "yes" Or Trim$(includeFormulasText) = "1")
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(safePath, 0, True)
    result = "WORKBOOK: " & safePath & vbCrLf & "SHEETS:" & vbCrLf
    For Each sh In wb.Worksheets
        Set used = sh.UsedRange
        result = result & "- " & sh.Name & " | " & used.Address(False, False) & " | rows=" & used.Rows.Count & " cols=" & used.Columns.Count & vbCrLf
    Next sh

    If Len(sheetName) > 0 Then
        Set ws = wb.Worksheets(sheetName)
    Else
        Set ws = wb.Worksheets(1)
    End If
    Set used = ws.UsedRange
    result = result & vbCrLf & "SAMPLE SHEET: " & ws.Name & " (" & used.Address(False, False) & ")" & vbCrLf
    sampleRows = used.Rows.Count
    If sampleRows > 20 Then sampleRows = 20
    sampleCols = used.Columns.Count
    If sampleCols > 10 Then sampleCols = 10
    For r = 1 To sampleRows
        For c = 1 To sampleCols
            If includeFormulas Then
                valueText = CStr(ws.Cells(used.Row + r - 1, used.Column + c - 1).Formula)
            Else
                valueText = CStr(ws.Cells(used.Row + r - 1, used.Column + c - 1).Value2)
            End If
            valueText = Replace(valueText, vbCr, " ")
            valueText = Replace(valueText, vbLf, " ")
            If Len(valueText) > 300 Then valueText = Left$(valueText, 300) & "..."
            If Len(valueText) > 0 Then result = result & ws.Cells(used.Row + r - 1, used.Column + c - 1).Address(False, False) & "=" & valueText & vbCrLf
        Next c
    Next r
    GoTo Cleanup

ErrHandler:
    result = "inspect_excelエラー: " & Err.Description

Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set ws = Nothing
    Set used = Nothing
    Set wb = Nothing
    Set xl = Nothing
    Tool_InspectExcel = result
End Function

Private Function Tool_GrepFiles(ByVal folder As String, ByVal pattern As String) As String
    Dim fso As Object
    Dim folderPath As String
    Dim f As Object
    Dim content As String
    Dim result As String
    Dim lines() As String
    Dim i As Long
    Dim matchCount As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(folder) = 0 Then folder = "."
    folderPath = ResolveSafePath(folder)

    If Not fso.FolderExists(folderPath) Then
        Tool_GrepFiles = "フォルダが存在しません"
        Exit Function
    End If

    For Each f In fso.GetFolder(folderPath).Files
        If f.Size <= 51200 Then
            On Error Resume Next
            content = Tool_ReadFile(fso.BuildPath(folderPath, f.Name))
            On Error GoTo 0
            If Len(content) > 0 Then
                lines = Split(content, vbLf)
                For i = LBound(lines) To UBound(lines)
                    If InStr(1, lines(i), pattern, vbTextCompare) > 0 Then
                        result = result & f.Name & ":" & (i + 1) & ": " & Trim$(Replace(lines(i), vbCr, "")) & vbCrLf
                        matchCount = matchCount + 1
                        If matchCount >= 50 Then
                            Tool_GrepFiles = result & "... (50件上限)"
                            Exit Function
                        End If
                    End If
                Next i
            End If
        End If
    Next f

    If Len(result) = 0 Then result = "一致なし"
    Tool_GrepFiles = result
End Function

Private Function Tool_WebFetch(ByVal url As String, ByVal maxCharsText As String) As String
    Const FETCH_TIMEOUT_MS As Long = 30000
    Const MAX_RESPONSE_BYTES As Long = 500000
    Const DEFAULT_MAX_CHARS As Long = 20000
    Const ABSOLUTE_MAX_CHARS As Long = 50000
    Dim http As Object
    Dim contentType As String
    Dim body As String
    Dim maxChars As Long
    Dim contentLength As String

    url = Trim$(url)
    ValidateWebFetchUrl url

    maxChars = Val(maxCharsText)
    If maxChars <= 0 Then maxChars = DEFAULT_MAX_CHARS
    If maxChars < 1000 Then maxChars = 1000
    If maxChars > ABSOLUTE_MAX_CHARS Then maxChars = ABSOLUTE_MAX_CHARS

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.SetTimeouts FETCH_TIMEOUT_MS, FETCH_TIMEOUT_MS, FETCH_TIMEOUT_MS, FETCH_TIMEOUT_MS
    On Error Resume Next
    http.setOption 5, MAX_RESPONSE_BYTES
    http.setOption 6, False
    On Error GoTo 0
    http.SetRequestHeader "Accept", "text/html, application/xhtml+xml, application/json, text/plain, text/*;q=0.8, */*;q=0.1"
    http.SetRequestHeader "User-Agent", "ExcelAI/" & APP_VERSION & " (web_fetch)"
    http.SetRequestHeader "Connection", "close"
    http.Send

    If http.Status >= 300 Then
        Err.Raise vbObjectError + 2201, "web_fetch", "HTTP " & http.Status & " のため取得できません"
    End If

    contentType = http.GetResponseHeader("Content-Type")
    contentLength = http.GetResponseHeader("Content-Length")
    If Val(contentLength) > MAX_RESPONSE_BYTES Then
        Err.Raise vbObjectError + 2202, "web_fetch", "応答が大きすぎます (500KB上限)"
    End If

    body = WebFetchToText(http.ResponseText, contentType)
    If Len(body) > maxChars Then body = Left$(body, maxChars) & vbCrLf & "[省略: max_chars=" & maxChars & "]"
    Tool_WebFetch = "URL: " & url & vbCrLf & _
        "HTTP: " & http.Status & vbCrLf & _
        "Content-Type: " & contentType & vbCrLf & vbCrLf & body
End Function

Private Sub ValidateWebFetchUrl(ByVal url As String)
    Dim schemePos As Long
    Dim authorityStart As Long
    Dim authorityEnd As Long
    Dim authority As String
    Dim host As String

    If Len(url) = 0 Or Len(url) > 2048 Then Err.Raise vbObjectError + 2203, "web_fetch", "URLが空か長すぎます"
    If InStr(1, url, vbCr, vbBinaryCompare) > 0 Or InStr(1, url, vbLf, vbBinaryCompare) > 0 Or InStr(1, url, " ", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 2204, "web_fetch", "URLに空白または改行は使えません"
    End If

    schemePos = InStr(1, url, "://", vbTextCompare)
    If schemePos <= 1 Then Err.Raise vbObjectError + 2205, "web_fetch", "http:// または https:// のURLを指定してください"
    If LCase$(Left$(url, schemePos - 1)) <> "http" And LCase$(Left$(url, schemePos - 1)) <> "https" Then
        Err.Raise vbObjectError + 2206, "web_fetch", "許可されている方式は http/https だけです"
    End If

    authorityStart = schemePos + 3
    authorityEnd = UrlAuthorityEnd(url, authorityStart)
    If authorityEnd = 0 Then authorityEnd = Len(url) + 1
    authority = Mid$(url, authorityStart, authorityEnd - authorityStart)
    If Len(authority) = 0 Or InStr(1, authority, "@", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 2207, "web_fetch", "ホスト名またはユーザー情報が不正です"
    End If

    host = authority
    If Left$(host, 1) = "[" Then
        If InStr(host, "]") > 0 Then host = Mid$(host, 2, InStr(host, "]") - 2)
    ElseIf InStr(host, ":") > 0 Then
        host = Left$(host, InStr(host, ":") - 1)
    End If
    host = LCase$(Trim$(host))
    If Len(host) = 0 Or host = "localhost" Or host = "::1" Or host = "0.0.0.0" Or host = "127.0.0.1" Or host = "169.254.169.254" Or Right$(host, 6) = ".local" Then
        Err.Raise vbObjectError + 2208, "web_fetch", "ローカル・リンクローカル・内部向けホストは取得できません"
    End If
    If IsPrivateIpv4(host) Then Err.Raise vbObjectError + 2209, "web_fetch", "プライベートIPアドレスは取得できません"
End Sub

Private Function UrlAuthorityEnd(ByVal url As String, ByVal startPos As Long) As Long
    Dim slashPos As Long, queryPos As Long, hashPos As Long, result As Long
    slashPos = InStr(startPos, url, "/")
    queryPos = InStr(startPos, url, "?")
    hashPos = InStr(startPos, url, "#")
    result = Len(url) + 1
    If slashPos > 0 And slashPos < result Then result = slashPos
    If queryPos > 0 And queryPos < result Then result = queryPos
    If hashPos > 0 And hashPos < result Then result = hashPos
    If result <= Len(url) Then UrlAuthorityEnd = result
End Function

Private Function IsPrivateIpv4(ByVal host As String) As Boolean
    Dim parts() As String
    Dim firstPart As Long, secondPart As Long
    If InStr(host, ".") = 0 Then Exit Function
    parts = Split(host, ".")
    If UBound(parts) <> 3 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then Exit Function
    firstPart = CLng(parts(0))
    secondPart = CLng(parts(1))
    IsPrivateIpv4 = (firstPart = 10) Or _
        (firstPart = 192 And secondPart = 168) Or _
        (firstPart = 172 And secondPart >= 16 And secondPart <= 31) Or _
        (firstPart = 127) Or (firstPart = 169 And secondPart = 254)
End Function

Private Function WebFetchToText(ByVal body As String, ByVal contentType As String) As String
    Dim doc As Object
    If InStr(1, LCase$(contentType), "html", vbTextCompare) = 0 Then
        WebFetchToText = body
        Exit Function
    End If

    On Error GoTo Fallback
    Set doc = CreateObject("HTMLFile")
    doc.Open
    doc.Write body
    doc.Close
    WebFetchToText = doc.body.innerText
    Exit Function

Fallback:
    Err.Clear
    WebFetchToText = body
End Function
