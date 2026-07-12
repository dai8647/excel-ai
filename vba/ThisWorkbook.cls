Option Explicit

Private Sub Workbook_Open()
    Dim safetyReport As String
    InitDefaultSettings
    On Error Resume Next
    safetyReport = WorkspaceSafetyReport()
    If Len(safetyReport) > 0 Then
        Ops_WriteStatus safetyReport, (InStr(1, safetyReport, "Workspace safety: OK", vbTextCompare) = 0)
        ThisWorkbook.Saved = True
    End If
    On Error GoTo 0
End Sub
