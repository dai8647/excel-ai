Attribute VB_Name = "Plugin_Snapshot"
Option Explicit

' Plugin_Snapshot: automatic pre-write backups.
' Every approved write that targets an existing workspace file is snapshotted
' into workspace\_backups before it is overwritten. You can list and restore.
'
' Tools:
'   list_backups(path[, name])   list snapshots of one file
'   restore_backup(path[, name]) copy a snapshot back over the original

Private Const BK_SUBDIR As String = "_backups"
Private Const BK_EXT As String = ".bak"
Private Const KEEP_COUNT As Long = 5

Public Function Plugin_Name() As String
    Plugin_Name = "snapshot"
End Function

Public Function Plugin_ToolNames() As String
    Plugin_ToolNames = "list_backups,restore_backup"
End Function

Public Function Plugin_WriteTools() As String
    Plugin_WriteTools = "restore_backup"
End Function

Public Function Plugin_ToolsOpenAI() As String
    Plugin_ToolsOpenAI = _
        ToolDef("list_backups", "List snapshots of a workspace file (each line: name, time, size)", _
            "{""path"":{""type"":""string""},""name"":{""type"":""string""}}") & "," & _
        ToolDef("restore_backup", "Restore a snapshot back over the original file. name optional (default latest)", _
            "{""path"":{""type"":""string""},""name"":{""type"":""string""}}")
End Function

Public Function Plugin_ToolsAnthropic() As String
    Plugin_ToolsAnthropic = _
        AnthropicToolDef("list_backups", "List snapshots of a workspace file (each line: name, time, size)", _
            "{""path"":{""type"":""string""},""name"":{""type"":""string""}}") & "," & _
        AnthropicToolDef("restore_backup", "Restore a snapshot back over the original file. name optional (default latest)", _
            "{""path"":{""type"":""string""},""name"":{""type"":""string""}}")
End Function

' Fired before any gated write touches targetPath (already resolved, absolute).
Public Sub Plugin_OnBeforeWrite(ByVal targetPath As String)
    On Error GoTo Done
    Dim fso As Object
    Dim bkDir As String
    Dim baseName As String
    Dim destName As String

    If Len(targetPath) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    bkDir = GetWorkspacePath() & "\" & BK_SUBDIR

    ' recursion guard: never snapshot files that live in the backups folder
    If Len(GetWorkspacePath()) = 0 Then Exit Sub
    If StrComp(Left$(targetPath, Len(bkDir)), bkDir, vbTextCompare) = 0 Then Exit Sub
    If Not IsPathInWorkspace(targetPath) Then Exit Sub

    ' nothing to snapshot if the file does not exist yet
    If Not fso.FileExists(targetPath) Then Exit Sub

    If Not fso.FolderExists(bkDir) Then fso.CreateFolder bkDir
    baseName = fso.GetBaseName(targetPath)
    destName = baseName & "." & Format$(Now, "yyyyMMdd-HHmmss") & "." & Millis3() & BK_EXT
    fso.CopyFile targetPath, fso.BuildPath(bkDir, destName), True
    PruneBackups fso, bkDir, baseName
Done:
End Sub

' Resolve which snapshot(s) belong to one original file, newest first.
' Returns a 1-based collection of full backup paths.
Private Function SnapshotsFor(ByRef fso As Object, ByVal originalFull As String) As Collection
    Dim col As Collection
    Dim bkDir As String
    Dim f As Object
    Dim prefix As String
    Dim arr() As String
    Dim i As Long, j As Long
    Dim tmp As String

    Set col = New Collection
    bkDir = GetWorkspacePath() & "\" & BK_SUBDIR
    If Not fso.FolderExists(bkDir) Then Set SnapshotsFor = col: Exit Function

    prefix = fso.GetBaseName(originalFull) & "."
    For Each f In fso.GetFolder(bkDir).Files
        If fso.GetExtensionName(f.Name) = "bak" Then
            If Left$(f.Name, Len(prefix)) = prefix Then col.Add f.Path
        End If
    Next f
    If col.Count < 2 Then Set SnapshotsFor = col: Exit Function

    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next i
    For i = 1 To col.Count - 1
        For j = i + 1 To col.Count
            If StrComp(arr(i), arr(j), vbBinaryCompare) < 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    Set col = New Collection
    For i = 1 To UBound(arr)
        col.Add arr(i)
    Next i
    Set SnapshotsFor = col
End Function

Private Sub PruneBackups(ByRef fso As Object, ByVal bkDir As String, ByVal baseName As String)
    Dim f As Object
    Dim prefix As String
    Dim arr() As String
    Dim i As Long, j As Long
    Dim n As Long
    Dim tmp As String

    prefix = baseName & "."
    ReDim arr(0 To 0)
    For Each f In fso.GetFolder(bkDir).Files
        If fso.GetExtensionName(f.Name) = "bak" Then
            If Left$(f.Name, Len(prefix)) = prefix Then
                n = n + 1
                ReDim Preserve arr(1 To n)
                arr(n) = f.Path
            End If
        End If
    Next f
    If n <= KEEP_COUNT Then Exit Sub
    For i = 1 To n - 1
        For j = i + 1 To n
            If StrComp(arr(i), arr(j), vbBinaryCompare) > 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To n - KEEP_COUNT
        On Error Resume Next
        fso.DeleteFile arr(i), True
        On Error GoTo 0
    Next i
End Sub

Public Function Plugin_Execute(ByVal toolName As String, ByVal argsJson As String) As String
    Dim path As String
    Dim name As String
    Select Case LCase$(toolName)
        Case "list_backups"
            Plugin_Execute = DoListBackups(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "name"))
        Case "restore_backup"
            Plugin_Execute = DoRestore(JsonGetArg(argsJson, "path"), JsonGetArg(argsJson, "name"))
        Case Else
            Plugin_Execute = "Plugin_Snapshot: 未知のツール " & toolName
    End Select
End Function

Private Function DoListBackups(ByVal path As String, ByVal name As String) As String
    On Error GoTo ErrHandler
    Dim fso As Object
    Dim original As String
    Dim col As Collection
    Dim i As Long
    Dim result As String
    Dim f As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    original = ResolveSafePath(path)
    Set col = SnapshotsFor(fso, original)
    If col.Count = 0 Then
        DoListBackups = "バックアップはありません: " & original
        Exit Function
    End If
    For i = 1 To col.Count
        Set f = fso.GetFile(CStr(col(i)))
        result = result & f.Name & "  " & Format$(f.DateLastModified, "yyyy/mm/dd hh:nn:ss") & _
            "  " & f.Size & " bytes" & vbCrLf
    Next i
    DoListBackups = Trim$(result)
    Exit Function
ErrHandler:
    DoListBackups = "list_backups エラー: " & Err.Description
End Function

Private Function DoRestore(ByVal path As String, ByVal name As String) As String
    On Error GoTo ErrHandler
    Dim fso As Object
    Dim original As String
    Dim col As Collection
    Dim pick As String
    Dim i As Long
    Dim content As String
    Dim auto As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    original = ResolveSafePath(path)
    If Not fso.FileExists(original) Then
        DoRestore = "元ファイルが存在しません: " & original
        Exit Function
    End If
    Set col = SnapshotsFor(fso, original)
    If col.Count = 0 Then
        DoRestore = "復元できるバックアップがありません: " & original
        Exit Function
    End If
    pick = ""
    If Len(name) > 0 Then
        For i = 1 To col.Count
            If StrComp(fso.GetFile(CStr(col(i))).Name, name, vbTextCompare) = 0 Then pick = CStr(col(i)): Exit For
        Next i
        If Len(pick) = 0 Then
            DoRestore = "指定バックアップが見つかりません: " & name
            Exit Function
        End If
    Else
        pick = CStr(col(1))
    End If

    content = ReadFileUtf8(fso, pick)
    auto = True
    ' go through write_file so the before-write hook snapshots the current file too
    DoRestore = Tools_Execute("write_file", _
        "{""path"":""" & JsonEscape(original) & """,""content"":""" & JsonEscape(content) & """}", auto)
    If Len(DoRestore) > 0 And InStr(DoRestore, "書き込み完了") > 0 Then
        DoRestore = "復元完了: " & fso.GetFile(pick).Name & " -> " & original
    End If
    Exit Function
ErrHandler:
    DoRestore = "restore_backup エラー: " & Err.Description
End Function

Private Function ReadFileUtf8(ByRef fso As Object, ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadFileUtf8 = stream.ReadText
    stream.Close
End Function

Private Function JsonEscape(ByVal s As String) As String
    JsonEscape = Replace$(Replace$(Replace$(Replace$(s, "\", "\\"), """", "\"""), vbCrLf, "\n"), vbLf, "\n")
End Function

Private Function Millis3() As String
    Millis3 = Right$("000" & CStr(CLng(Timer * 1000) Mod 1000), 3)
End Function

Public Function Plugin_SelfTest() As String
    On Error GoTo Fail
    Dim fso As Object
    Dim ws As String
    Dim tmpPath As String
    Dim bkDir As String
    Dim auto As Boolean
    Dim r As String
    Dim col As Collection

    Set fso = CreateObject("Scripting.FileSystemObject")
    ws = GetWorkspacePath()
    bkDir = ws & "\" & BK_SUBDIR
    tmpPath = ws & "\snapshot_selftest.txt"
    auto = True

    ' first write creates the file (no backup), second and later create snapshots
    r = Tools_Execute("write_file", "{""path"":""snapshot_selftest.txt"",""content"":""v1""}", auto)
    r = Tools_Execute("write_file", "{""path"":""snapshot_selftest.txt"",""content"":""v2""}", auto)
    r = Tools_Execute("write_file", "{""path"":""snapshot_selftest.txt"",""content"":""v3""}", auto)

    Set col = SnapshotsFor(fso, tmpPath)
    If col.Count < 2 Then
        Plugin_SelfTest = "FAIL: バックアップが " & col.Count & " 個 (2以上必要)"
        Exit Function
    End If

    r = Plugin_Execute("list_backups", "{""path"":""snapshot_selftest.txt""}")
    If InStr(r, "snapshot_selftest") = 0 Then
        Plugin_SelfTest = "FAIL: list_backups にファイル名なし"
        Exit Function
    End If

    r = Plugin_Execute("restore_backup", "{""path"":""snapshot_selftest.txt""}")
    If InStr(r, "復元完了") = 0 Then
        Plugin_SelfTest = "FAIL: restore_backup 失敗 -> " & Left$(r, 80)
        Exit Function
    End If

    r = DoRestore("snapshot_selftest.txt", "")
    If InStr(r, "復元完了") = 0 Then
        Plugin_SelfTest = "FAIL: restore (latest) 失敗 -> " & Left$(r, 80)
        Exit Function
    End If

    On Error Resume Next
    fso.DeleteFile tmpPath, True
    If fso.FolderExists(bkDir) Then
        Dim f As Object
        For Each f In fso.GetFolder(bkDir).Files
            If fso.GetBaseName(f.Name) = "snapshot_selftest" Then fso.DeleteFile f.Path, True
        Next f
        fso.DeleteFolder bkDir, True
    End If
    On Error GoTo 0

    Plugin_SelfTest = "OK"
    Exit Function
Fail:
    Plugin_SelfTest = "FAIL: " & Err.Description
End Function