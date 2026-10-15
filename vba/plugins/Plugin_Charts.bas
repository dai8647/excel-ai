Attribute VB_Name = "Plugin_Charts"
Option Explicit

' Plugin: charts - export a chart in an Excel workbook to a PNG image.
' Exposes tool "chart_to_image". Requires a workbook path; optionally a
' sheet_name and a 1-based chart_index (default 1). The chart is exported
' with Chart.Export to workspace\charts\<base>_<index>_<hhnnss>.png and
' the relative path of the result is returned.

Public Function Plugin_Name() As String
    Plugin_Name = "charts"
End Function

Public Function Plugin_ToolNames() As String
    Plugin_ToolNames = "chart_to_image"
End Function

Public Function Plugin_WriteTools() As String
    Plugin_WriteTools = ""
End Function

Public Function Plugin_ToolsOpenAI() As String
    Plugin_ToolsOpenAI = ToolDef("chart_to_image", _
        "Export a chart from an Excel workbook to a PNG image saved under workspace\charts. " & _
        "Provide the workbook path; optionally the sheet name and a 1-based chart index (default 1). " & _
        "Returns the relative path of the saved PNG.", _
        "{""path"":{""type"":""string"",""description"":""Workbook path (.xlsx/.xlsm)""}," & _
        """sheet_name"":{""type"":""string"",""description"":""Optional sheet name; default first sheet""}," & _
        """chart_index"":{""type"":""integer"",""description"":""Optional 1-based chart index on the sheet; default 1""}}")
End Function

Public Function Plugin_ToolsAnthropic() As String
    Plugin_ToolsAnthropic = AnthropicToolDef("chart_to_image", _
        "Export a chart from an Excel workbook to a PNG image saved under workspace\charts. " & _
        "Provide the workbook path; optionally the sheet name and a 1-based chart index (default 1). " & _
        "Returns the relative path of the saved PNG.", _
        "{""path"":{""type"":""string""},""sheet_name"":{""type"":""string""},""chart_index"":{""type"":""integer""}}")
End Function

Public Function Plugin_Execute(ByVal toolName As String, ByVal argsJson As String) As String
    Select Case LCase$(toolName)
        Case "chart_to_image"
            Plugin_Execute = ExportChart(JsonGetArg(argsJson, "path"), _
                JsonGetArg(argsJson, "sheet_name"), JsonGetArg(argsJson, "chart_index"))
        Case Else
            Plugin_Execute = "Plugin_Charts: 未知のツール " & toolName
    End Select
End Function

Private Function ExportChart(ByVal path As String, ByVal sheetName As String, ByVal chartIndexText As String) As String
    Dim safePath As String
    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim chartObject As Object
    Dim chart As Object
    Dim fso As Object
    Dim outFolder As String
    Dim outFile As String
    Dim baseName As String
    Dim chartIndex As Long
    Dim stamp As String
    Dim relative As String
    Dim resultText As String

    On Error GoTo ErrHandler
    safePath = ResolveSafePath(path)
    If LCase$(Right$(safePath, 5)) <> ".xlsx" And LCase$(Right$(safePath, 5)) <> ".xlsm" Then
        Err.Raise vbObjectError + 2401, "chart_to_image", "Excelファイル(.xlsx/.xlsm)を指定してください"
    End If

    chartIndex = Val(chartIndexText)
    If chartIndex < 1 Then chartIndex = 1

    Set fso = CreateObject("Scripting.FileSystemObject")
    outFolder = GetWorkspacePath() & "\charts"
    If Not fso.FolderExists(outFolder) Then fso.CreateFolder outFolder

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(safePath)

    If Len(sheetName) > 0 Then
        Set ws = wb.Worksheets(sheetName)
    Else
        Set ws = wb.Worksheets(1)
    End If

    If ws.ChartObjects.Count < chartIndex Then
        Err.Raise vbObjectError + 2402, "chart_to_image", _
            "シート '" & ws.Name & "' に chart_index " & chartIndex & " のグラフがありません (存在: " & ws.ChartObjects.Count & ")"
    End If

    Set chartObject = ws.ChartObjects(chartIndex)
    Set chart = chartObject.Chart

    baseName = fso.GetBaseName(safePath)
    stamp = Format$(Now, "hhnnss")
    outFile = outFolder & "\" & baseName & "_" & chartIndex & "_" & stamp & ".png"

    ' Chart.Export requires an absolute path.
    chart.Export outFile, "PNG"

    relative = RelativeFromWorkspace(outFile)
    resultText = "グラフPNG書き出し完了: " & relative
    GoTo Cleanup

ErrHandler:
    resultText = "chart_to_imageエラー: " & Err.Description

Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set chart = Nothing
    Set chartObject = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing
    ExportChart = resultText
End Function

' Turn an absolute path inside the workspace into a workspace-relative path
' with forward slashes so the model can reference it easily.
Private Function RelativeFromWorkspace(ByVal fullPath As String) As String
    Dim workspace As String
    workspace = GetWorkspacePath()
    If Len(fullPath) > Len(workspace) Then
        If StrComp(Left$(fullPath, Len(workspace)), workspace, vbTextCompare) = 0 Then
            RelativeFromWorkspace = Mid$(fullPath, Len(workspace) + 2)
        Else
            RelativeFromWorkspace = fullPath
        End If
    Else
        RelativeFromWorkspace = fullPath
    End If
    RelativeFromWorkspace = Replace(RelativeFromWorkspace, "\", "/")
End Function

Public Function Plugin_SelfTest() As String
    Dim fso As Object
    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim chartObject As Object
    Dim chart As Object
    Dim testWb As String
    Dim testChartFolder As String
    Dim testOut As String
    Dim ok As Boolean

    On Error GoTo Skip
    Set fso = CreateObject("Scripting.FileSystemObject")
    testWb = GetWorkspacePath() & "\charts\selftest_chart.xlsx"
    testChartFolder = GetWorkspacePath() & "\charts"
    If Not fso.FolderExists(testChartFolder) Then fso.CreateFolder testChartFolder

    ' Build a small workbook with one chart so Chart.Export can run.
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Cells(1, 1).Value = "A"
    ws.Cells(1, 2).Value = 10
    ws.Cells(2, 1).Value = "B"
    ws.Cells(2, 2).Value = 20
    Set chartObject = ws.ChartObjects.Add(320, 20, 300, 220)
    Set chart = chartObject.Chart
    chart.SetSourceData ws.Range("A1:B2")
    wb.SaveAs testWb, 51

    ' Export to PNG and confirm the file exists.
    testOut = testChartFolder & "\selftest_chart_1_" & Format$(Now, "hhnnss") & ".png"
    chart.Export testOut, "PNG"
    ok = fso.FileExists(testOut)

    On Error Resume Next
    wb.Close False
    xl.Quit
    fso.DeleteFile testWb, True
    fso.DeleteFile testOut, True
    If fso.GetFolder(testChartFolder).Files.Count = 0 And fso.GetFolder(testChartFolder).SubFolders.Count = 0 Then
        fso.DeleteFolder testChartFolder, True
    End If
    On Error GoTo 0

    Set chart = Nothing
    Set chartObject = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing

    If ok Then
        Plugin_SelfTest = "OK"
    Else
        Plugin_SelfTest = "PNG出力を確認できませんでした"
    End If
    Exit Function

Skip:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    On Error GoTo 0
    Plugin_SelfTest = ""
End Function