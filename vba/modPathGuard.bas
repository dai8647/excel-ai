Attribute VB_Name = "modPathGuard"
Option Explicit

Public Function ResolveSafePath(ByVal inputPath As String) As String
    Dim fso As Object
    Dim workspace As String
    Dim candidate As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    workspace = GetWorkspacePath()
    EnsureWorkspaceExists

    inputPath = Trim$(inputPath)
    If Len(inputPath) = 0 Then Err.Raise vbObjectError + 1001, "PathGuard", "パスが空です"

    If InStr(inputPath, "..") > 0 Then
        Err.Raise vbObjectError + 1002, "PathGuard", "親ディレクトリ参照 (..) は禁止です"
    End If

    If (Len(inputPath) >= 2 And Mid$(inputPath, 2, 1) = ":") Or Left$(inputPath, 2) = "\\" Then
        candidate = NormalizePath(inputPath)
    Else
        candidate = NormalizePath(fso.BuildPath(workspace, inputPath))
    End If

    If Not IsPathInWorkspace(candidate) Then
        Err.Raise vbObjectError + 1003, "PathGuard", "作業場所外のパスは使用できません: " & candidate
    End If

    ResolveSafePath = candidate
End Function

Public Function IsPathInWorkspace(ByVal fullPath As String) As Boolean
    Dim workspace As String
    Dim normalized As String

    workspace = GetWorkspacePath()
    normalized = NormalizePath(fullPath)

    If Len(normalized) < Len(workspace) Then
        IsPathInWorkspace = False
        Exit Function
    End If

    If StrComp(Left$(normalized, Len(workspace)), workspace, vbTextCompare) <> 0 Then
        IsPathInWorkspace = False
        Exit Function
    End If

    If Len(normalized) = Len(workspace) Then
        IsPathInWorkspace = True
        Exit Function
    End If

    Dim nextChar As String
    nextChar = Mid$(normalized, Len(workspace) + 1, 1)
    IsPathInWorkspace = (nextChar = "\" Or nextChar = "/")
End Function

Public Function RequireSafePath(ByVal inputPath As String) As String
    RequireSafePath = ResolveSafePath(inputPath)
End Function

Public Function WorkspaceSafetyReport() As String
    Dim workspace As String
    Dim safePath As String
    Dim escapePath As String
    Dim insideOk As Boolean
    Dim parentBlocked As Boolean
    Dim outsideBlocked As Boolean

    On Error GoTo ErrHandler
    workspace = GetWorkspacePath()
    safePath = ResolveSafePath("README.txt")
    insideOk = IsPathInWorkspace(safePath)
    parentBlocked = PathGuardRejects("..\workspace-escape-test.txt")
    escapePath = NormalizePath(workspace & "\..\workspace-escape-test.txt")
    outsideBlocked = PathGuardRejects(escapePath)

    If insideOk And parentBlocked And outsideBlocked Then
        WorkspaceSafetyReport = "Workspace safety: OK" & vbCrLf & _
            "Folder: " & workspace & vbCrLf & _
            "Inside path: OK" & vbCrLf & _
            "Parent traversal: BLOCKED" & vbCrLf & _
            "Outside absolute path: BLOCKED"
    Else
        WorkspaceSafetyReport = "Workspace safety: REVIEW" & vbCrLf & _
            "Folder: " & workspace & vbCrLf & _
            "Inside path: " & IIf(insideOk, "OK", "REVIEW") & vbCrLf & _
            "Parent traversal: " & IIf(parentBlocked, "BLOCKED", "REVIEW") & vbCrLf & _
            "Outside absolute path: " & IIf(outsideBlocked, "BLOCKED", "REVIEW")
    End If
    Exit Function

ErrHandler:
    WorkspaceSafetyReport = "Workspace safety: ERROR" & vbCrLf & Err.Description
End Function

Public Sub WorkspaceSafetyCheck()
    Dim report As String
    report = WorkspaceSafetyReport()
    Ops_WriteStatus report, (InStr(1, report, "Workspace safety: OK", vbTextCompare) = 0)
    MsgBox report, vbInformation, "Workspace safety"
End Sub

Private Function PathGuardRejects(ByVal inputPath As String) As Boolean
    Dim resolved As String
    On Error Resume Next
    Err.Clear
    resolved = ResolveSafePath(inputPath)
    PathGuardRejects = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
End Function
