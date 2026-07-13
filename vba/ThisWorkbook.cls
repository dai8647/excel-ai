Option Explicit

Private Sub Workbook_Open()
    Dim safetyReport As String
    InitDefaultSettings
    RegisterShortcuts
    On Error Resume Next
    safetyReport = WorkspaceSafetyReport()
    If Len(safetyReport) > 0 Then
        Ops_WriteProgress "Workspace ready", (InStr(1, safetyReport, "Workspace safety: OK", vbTextCompare) = 0)
        ThisWorkbook.Saved = True
    End If
    On Error GoTo 0
    ShowHome
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    UnregisterShortcuts
End Sub
