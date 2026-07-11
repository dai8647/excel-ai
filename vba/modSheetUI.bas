Attribute VB_Name = "modSheetUI"
Option Explicit

Public Sub ShowAgentFallback()
    InitDefaultSettings
    ThisWorkbook.Worksheets(SHEET_OPS).Activate
    MsgBox "操作シートの Request 欄に依頼を入力して、Run request を押してください。", vbInformation, "ExcelAI"
End Sub

Public Sub ShowAgentOrFallback()
    ShowAgentFallback
End Sub
