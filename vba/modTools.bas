Attribute VB_Name = "modTools"
Option Explicit

Public Function Tools_GetDefinitionsJson() As String
    Tools_GetDefinitionsJson = "[" & _
        ToolDef("list_dir", "List files and folders in a directory", "{""path"":{""type"":""string"",""description"":""Relative or absolute path within workspace""}}") & "," & _
        ToolDef("read_file", "Read a text or CSV file (max 50KB)", "{""path"":{""type"":""string""}}") & "," & _
        ToolDef("write_file", "Write or overwrite a text file", "{""path"":{""type"":""string""},""content"":{""type"":""string""}}") & "," & _
        ToolDef("append_file", "Append text to a file", "{""path"":{""type"":""string""},""content"":{""type"":""string""}}") & "," & _
        ToolDef("move_file", "Move or rename a file/folder", "{""source"":{""type"":""string""},""destination"":{""type"":""string""}}") & "," & _
        ToolDef("delete_file", "Delete a file (requires confirmation setting)", "{""path"":{""type"":""string""}}") & "," & _
        ToolDef("merge_csv", "Merge CSV files matching a pattern into one file", "{""source_pattern"":{""type"":""string"",""description"":""e.g. data\\*.csv""},""output_path"":{""type"":""string""}}") & "," & _
        ToolDef("create_excel", "Create a new Excel workbook from CSV-like data", "{""path"":{""type"":""string"",""description"":""Output .xlsx path""},""sheet_name"":{""type"":""string""},""csv_content"":{""type"":""string"",""description"":""Comma-separated rows, newline-separated lines""}}") & "," & _
        ToolDef("grep_files", "Search for text in files under a folder", "{""folder"":{""type"":""string""},""pattern"":{""type"":""string""}}") & _
        "]"
End Function

Public Function Tools_GetAnthropicDefinitionsJson() As String
    Tools_GetAnthropicDefinitionsJson = "[" & _
        AnthropicToolDef("list_dir", "List files and folders in a directory", "{""path"":{""type"":""string"",""description"":""Relative or absolute path within workspace""}}") & "," & _
        AnthropicToolDef("read_file", "Read a text or CSV file (max 50KB)", "{""path"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("write_file", "Write or overwrite a text file", "{""path"":{""type"":""string""},""content"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("append_file", "Append text to a file", "{""path"":{""type"":""string""},""content"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("move_file", "Move or rename a file/folder", "{""source"":{""type"":""string""},""destination"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("delete_file", "Delete a file (requires confirmation setting)", "{""path"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("merge_csv", "Merge CSV files matching a pattern into one file", "{""source_pattern"":{""type"":""string""},""output_path"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("create_excel", "Create a new Excel workbook from CSV-like data", "{""path"":{""type"":""string""},""sheet_name"":{""type"":""string""},""csv_content"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("grep_files", "Search for text in files under a folder", "{""folder"":{""type"":""string""},""pattern"":{""type"":""string""}}") & "]"
End Function

Private Function ToolDef(ByVal name As String, ByVal description As String, ByVal params As String) As String
    ToolDef = "{""type"":""function"",""function"":{" & _
        """name"":" & JsonString(name) & "," & _
        """description"":" & JsonString(description) & "," & _
        """parameters"":{""type"":""object"",""properties"":" & params & "}}}"
End Function

Private Function AnthropicToolDef(ByVal name As String, ByVal description As String, ByVal params As String) As String
    AnthropicToolDef = "{""name"":" & JsonString(name) & ",""description"":" & JsonString(description) & _
        ",""input_schema"":{""type"":""object"",""properties"":" & params & "}}"
End Function

Public Function Tools_Execute(ByVal toolName As String, ByVal argsJson As String, ByVal autoApproved As Boolean) As String
    On Error GoTo ErrHandler

    Select Case LCase$(toolName)
        Case "list_dir"
            Tools_Execute = Tool_ListDir(JsonGetArg(argsJson, "path"))
        Case "read_file"
            Tools_Execute = Tool_ReadFile(JsonGetArg(argsJson, "path"))
        Case "write_file"
            If Not autoApproved And Not IsAutoRunEnabled() Then
                If Not ConfirmToolAction("write_file", argsJson) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
                End If
            End If
            Tools_Execute = Tool_WriteFile(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "content"))
        Case "append_file"
            If Not autoApproved And Not IsAutoRunEnabled() Then
                If Not ConfirmToolAction("append_file", argsJson) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
                End If
            End If
            Tools_Execute = Tool_AppendFile(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "content"))
        Case "move_file"
            If Not autoApproved And Not IsAutoRunEnabled() Then
                If Not ConfirmToolAction("move_file", argsJson) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
                End If
            End If
            Tools_Execute = Tool_MoveFile(JsonGetArg(argsJson, "source"), JsonGetArg(argsJson, "destination"))
        Case "delete_file"
            If IsConfirmDangerEnabled() Then
                If MsgBox("削除を実行しますか？" & vbCrLf & JsonGetArg(argsJson, "path"), vbYesNo + vbExclamation) <> vbYes Then
                    Tools_Execute = "ユーザーが削除をキャンセルしました"
                    Exit Function
                End If
            ElseIf Not autoApproved And Not IsAutoRunEnabled() Then
                If Not ConfirmToolAction("delete_file", argsJson) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
                End If
            End If
            Tools_Execute = Tool_DeleteFile(JsonGetArg(argsJson, "path"))
        Case "merge_csv"
            If Not autoApproved And Not IsAutoRunEnabled() Then
                If Not ConfirmToolAction("merge_csv", argsJson) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
                End If
            End If
            Tools_Execute = Tool_MergeCsv(JsonGetArg(argsJson, "source_pattern"), JsonGetArg(argsJson, "output_path"))
        Case "create_excel"
            If Not autoApproved And Not IsAutoRunEnabled() Then
                If Not ConfirmToolAction("create_excel", argsJson) Then
                    Tools_Execute = "ユーザーが実行をキャンセルしました"
                    Exit Function
                End If
            End If
            Tools_Execute = Tool_CreateExcel(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "sheet_name"), JsonGetArg(argsJson, "csv_content"))
        Case "grep_files"
            Tools_Execute = Tool_GrepFiles(JsonGetArg(argsJson, "folder"), JsonGetArg(argsJson, "pattern"))
        Case Else
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

Private Function Tool_ListDir(ByVal path As String) As String
    Dim fso As Object
    Dim safePath As String
    Dim entry As String
    Dim fullPath As String
    Dim result As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(path) = 0 Then path = "."
    safePath = ResolveSafePath(path)

    If Not fso.FolderExists(safePath) Then
        Tool_ListDir = "フォルダが存在しません: " & safePath
        Exit Function
    End If

    entry = Dir$(safePath & "\*", vbNormal Or vbHidden Or vbSystem Or vbDirectory)
    Do While Len(entry) > 0
        If entry <> "." And entry <> ".." Then
            fullPath = safePath & "\" & entry
            If (GetAttr(fullPath) And vbDirectory) = vbDirectory Then
                result = result & "[DIR]  " & entry & vbCrLf
            Else
                result = result & "[FILE] " & entry & " (" & FileLen(fullPath) & " bytes)" & vbCrLf
            End If
        End If
        entry = Dir$()
    Loop

    If Len(result) = 0 Then result = "(空のフォルダ)"
    Tool_ListDir = result
End Function

Private Function Tool_ReadFile(ByVal path As String) As String
    Dim fso As Object
    Dim stream As Object
    Dim safePath As String
    Dim content As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    safePath = ResolveSafePath(path)

    If Not fso.FileExists(safePath) Then
        Tool_ReadFile = "ファイルが存在しません: " & safePath
        Exit Function
    End If

    If fso.GetFile(safePath).Size > 51200 Then
        Tool_ReadFile = "ファイルが大きすぎます (50KB上限): " & safePath
        Exit Function
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile safePath
    content = stream.ReadText
    stream.Close

    Tool_ReadFile = content
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
    Dim i As Long

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
            Next c
            r = r + 1
        End If
    Next i

    wb.SaveAs safePath
    wb.Close False
    xl.Quit

    Tool_CreateExcel = "Excel作成完了: " & safePath
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
