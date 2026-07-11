#Requires -Version 5.1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlsmPath = Join-Path $root "ExcelAI.xlsm"
$vbaDir = Join-Path $root "vba"
$routerKey = ""
$routerDb = Join-Path $env:APPDATA "9router\db\data.sqlite"
$sqliteModule = Join-Path $env:APPDATA "9router\runtime\node_modules\better-sqlite3"
if ((Test-Path $routerDb) -and (Test-Path $sqliteModule)) {
    try {
        $keyScript = "const Database=require(process.argv[1]);const db=new Database(process.argv[2],{readonly:true});const row=db.prepare('select key from apiKeys where isActive=1 order by createdAt desc limit 1').get();if(row)process.stdout.write(row.key);"
        $routerKey = (& node -e $keyScript $sqliteModule $routerDb).Trim()
    }
    catch {
        $routerKey = ""
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

try {
    if (Test-Path $xlsmPath) {
        $backupPath = Join-Path $root ("ExcelAI.backup-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".xlsm")
        Copy-Item $xlsmPath $backupPath
        Remove-Item $xlsmPath -Force
    }
    $wb = $excel.Workbooks.Add()

    while (($wb.Worksheets.Count) -lt 4) {
        $wb.Worksheets.Add() | Out-Null
    }

    $wb.Worksheets.Item(1).Name = [char]0x8A2D + [char]0x5B9A
    $wb.Worksheets.Item(2).Name = [char]0x64CD + [char]0x4F5C
    $wb.Worksheets.Item(3).Name = -join @([char]0x30C1, [char]0x30E3, [char]0x30C3, [char]0x30C8, [char]0x30ED, [char]0x30B0)
    $wb.Worksheets.Item(4).Name = -join @([char]0x30E2, [char]0x30C7, [char]0x30EB, [char]0x4E00, [char]0x89A7)

    $ws = $wb.Worksheets.Item(1)
    $ws.Range("A1:C1").Merge()
    $ws.Cells.Item(1, 1).Value2 = "ExcelAI | Connection settings"
    $labels = @("Connection type", "API Key", "Base URL", "Model", "Workspace", "AutoRun", "Confirm danger", "Max steps", "Anthropic version")
    $hints = @("OpenAI or Anthropic", "Enter your 9router or provider key", "OpenAI includes /v1. Anthropic does not.", "Refresh models, then select one", "AI actions stay inside this folder", "No is recommended", "Yes is recommended", "12 is a safe start", "Native Anthropic API version")
    for ($i = 0; $i -lt $labels.Count; $i++) { $ws.Cells.Item($i + 2, 1).Value2 = $labels[$i]; $ws.Cells.Item($i + 2, 3).Value2 = $hints[$i] }
    $ws.Cells.Item(2, 2).Value2 = "OpenAI"
    $ws.Cells.Item(3, 2).NumberFormat = "@"
    $ws.Cells.Item(3, 2).Value2 = $routerKey
    $ws.Cells.Item(4, 2).Value2 = "http://127.0.0.1:20128/v1"
    $ws.Cells.Item(6, 2).Value2 = (Join-Path $root "workspace")
    $ws.Cells.Item(7, 2).Value2 = "no"
    $ws.Cells.Item(8, 2).Value2 = "yes"
    $ws.Cells.Item(9, 2).Value2 = 12
    $ws.Cells.Item(10, 2).NumberFormat = "@"
    $ws.Cells.Item(10, 2).Value2 = "2023-06-01"
    $ws.Columns.Item("A").ColumnWidth = 22; $ws.Columns.Item("B").ColumnWidth = 42; $ws.Columns.Item("C").ColumnWidth = 52
    $ws.Range("A1:C1").Interior.Color = 3158064; $ws.Range("A1:C1").Font.Color = 16777215; $ws.Range("A1:C1").Font.Bold = $true; $ws.Range("A1:C1").Font.Size = 16
    $ws.Range("A2:A10").Font.Bold = $true; $ws.Range("B2:B10").Interior.Color = 13431551; $ws.Range("C2:C10").Font.Color = 8421504; $ws.Range("A2:C10").Borders.LineStyle = 1
    $ws.Range("B2").Validation.Add(3, 1, 1, "OpenAI,Anthropic")
    $ws.Range("B7:B8").Validation.Add(3, 1, 1, "yes,no")

    $ops = $wb.Worksheets.Item(2)
    $ops.Range("A1:B1").Merge(); $ops.Cells.Item(1, 1).Value2 = "ExcelAI | Start here"
    $ops.Cells.Item(2, 1).Value2 = "1. Enter an API key.  2. Refresh models.  3. Test connection.  4. Open agent."
    $ops.Range("A1:B1").Interior.Color = 3158064; $ops.Range("A1:B1").Font.Color = 16777215; $ops.Range("A1:B1").Font.Bold = $true; $ops.Range("A1:B1").Font.Size = 16
    $ops.Columns.Item("A").ColumnWidth = 28; $ops.Columns.Item("B").ColumnWidth = 72
    $buttonSpecs = @(
        @("1. Refresh models", "RefreshModels"),
        @("2. Test connection", "TestConnection"),
        @("3. Open request panel", "ShowAgent"),
        @("4. Run request", "RunSheetPrompt"),
        @("Open workspace", "OpenWorkspaceFolder"),
        @("Clear chat log", "ClearChatLog")
    )
    $top = 60
    foreach ($spec in $buttonSpecs) {
        $button = $ops.Buttons().Add(20, $top, 220, 30)
        $button.Characters().Text = $spec[0]
        $button.OnAction = $spec[1]
        $top += 42
    }
    $ops.Cells.Item(13, 1).Value2 = "Latest check"; $ops.Cells.Item(14, 1).Value2 = "Time"; $ops.Cells.Item(15, 1).Value2 = "Result"
    $ops.Range("A13:B15").Borders.LineStyle = 1; $ops.Range("A13:A15").Font.Bold = $true
    $ops.Cells.Item(7, 1).Value2 = "Request"
    $ops.Cells.Item(7, 2).Value2 = ""
    $ops.Cells.Item(9, 1).Value2 = "Result"
    $ops.Range("A7:B7").Borders.LineStyle = 1; $ops.Range("A9:B9").Borders.LineStyle = 1
    $ops.Range("A7:A9").Font.Bold = $true
    $ops.Range("B7").Interior.Color = 13431551
    $ops.Range("B9").WrapText = $true
    $ops.Rows.Item(9).RowHeight = 90

    $log = $wb.Worksheets.Item(3)
    $log.Cells.Item(1, 1).Value2 = "Time"; $log.Cells.Item(1, 2).Value2 = "Type"; $log.Cells.Item(1, 3).Value2 = "Content"
    $log.Range("A1:C1").Interior.Color = 3158064; $log.Range("A1:C1").Font.Color = 16777215; $log.Range("A1:C1").Font.Bold = $true
    $log.Columns.Item("A").ColumnWidth = 22; $log.Columns.Item("B").ColumnWidth = 16; $log.Columns.Item("C").ColumnWidth = 90; $log.Rows.Item(1).AutoFilter()

    $models = $wb.Worksheets.Item(4)
    $models.Range("A1:B1").Merge(); $models.Cells.Item(1, 1).Value2 = "Model list"
    $models.Cells.Item(3, 1).Value2 = "Use Settings > Refresh models to update this list."
    $models.Cells.Item(5, 1).Value2 = "Model ID"; $models.Cells.Item(5, 2).Value2 = "Connection type"
    $models.Range("A1:B1").Interior.Color = 3158064; $models.Range("A1:B1").Font.Color = 16777215; $models.Range("A1:B1").Font.Bold = $true
    $models.Range("A5:B5").Interior.Color = 12632256; $models.Range("A5:B5").Font.Bold = $true
    $models.Columns.Item("A").ColumnWidth = 54; $models.Columns.Item("B").ColumnWidth = 20
    $ws.Range("B5").Validation.Add(3, 1, 1, "='" + $models.Name + "'!`$A`$6:`$A`$500")
    try {
        $modelHeaders = @{}
        if (-not [string]::IsNullOrWhiteSpace($routerKey)) { $modelHeaders["Authorization"] = "Bearer " + $routerKey }
        $routerModels = Invoke-RestMethod -Uri "http://127.0.0.1:20128/v1/models" -Headers $modelHeaders -TimeoutSec 10
        $modelRow = 6
        foreach ($routerModel in @($routerModels.data)) {
            if ($routerModel.id) {
                $models.Cells.Item($modelRow, 1).Value2 = $routerModel.id
                $models.Cells.Item($modelRow, 2).Value2 = "OpenAI"
                $modelRow += 1
            }
        }
        $defaultModel = "groq/qwen/qwen3-32b"
        if (@($routerModels.data.id) -contains $defaultModel) {
            $ws.Cells.Item(5, 2).Value2 = $defaultModel
        }
        elseif ($modelRow -gt 6) {
            $ws.Cells.Item(5, 2).Value2 = $models.Cells.Item(6, 1).Value2
        }
    }
    catch { }

    $vbProj = $wb.VBProject
    $importRoot = Join-Path $env:TEMP "ExcelAI-vba-import"
    if (Test-Path $importRoot) { Remove-Item $importRoot -Recurse -Force }
    Copy-Item $vbaDir $importRoot -Recurse

    Get-ChildItem (Join-Path $importRoot "*.bas") | ForEach-Object {
        $moduleCode = Get-Content $_.FullName -Raw -Encoding UTF8
        $moduleCode = $moduleCode -replace "(?m)^Attribute[^\r\n]*\r?\n", ""
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $moduleComponent = $vbProj.VBComponents.Add(1)
        $moduleComponent.Name = $moduleName
        $moduleComponent.CodeModule.AddFromString($moduleCode)
    }
    $frm = Join-Path $importRoot "frmAgent.frm"
    if ($false -and (Test-Path $frm)) {
        # Importing a .frm file depends on the local FM20 registration. Create the
        # standard controls through Excel instead, then inject the form's code.
        $formComponent = $vbProj.VBComponents.Add(3)
        $formComponent.Name = "frmAgent"
        $designer = $formComponent.Designer
        $formComponent.Properties.Item("Width").Value = 480
        $formComponent.Properties.Item("Height").Value = 360

        $txtChat = $designer.Controls.Add("Forms.TextBox.1", "txtChat", $true)
        $txtChat.Left = 12; $txtChat.Top = 12; $txtChat.Width = 456; $txtChat.Height = 240
        $txtChat.Locked = $true; $txtChat.MultiLine = $true; $txtChat.ScrollBars = 3; $txtChat.TabStop = $false

        $txtInput = $designer.Controls.Add("Forms.TextBox.1", "txtInput", $true)
        $txtInput.Left = 12; $txtInput.Top = 264; $txtInput.Width = 384; $txtInput.Height = 30
        $txtInput.MultiLine = $true

        $cmdSend = $designer.Controls.Add("Forms.CommandButton.1", "cmdSend", $true)
        $cmdSend.Left = 402; $cmdSend.Top = 264; $cmdSend.Width = 66; $cmdSend.Height = 30; $cmdSend.Caption = "Send"

        $lblStatus = $designer.Controls.Add("Forms.Label.1", "lblStatus", $true)
        $lblStatus.Left = 12; $lblStatus.Top = 306; $lblStatus.Width = 456; $lblStatus.Caption = "Ready"
        $lblAutoRun = $designer.Controls.Add("Forms.Label.1", "lblAutoRun", $true)
        $lblAutoRun.Left = 12; $lblAutoRun.Top = 324; $lblAutoRun.Width = 228; $lblAutoRun.Caption = "AutoRun"
        $lblWorkspace = $designer.Controls.Add("Forms.Label.1", "lblWorkspace", $true)
        $lblWorkspace.Left = 12; $lblWorkspace.Top = 342; $lblWorkspace.Width = 456; $lblWorkspace.Caption = "Workspace"

        $formCode = Get-Content $frm -Raw -Encoding UTF8
        $codeStart = $formCode.IndexOf("Option Explicit")
        if ($codeStart -lt 0) { throw "frmAgent code was not found." }
        $formComponent.CodeModule.AddFromString($formCode.Substring($codeStart))
    }

    $codePath = Join-Path $vbaDir "ThisWorkbook.cls"
    if (Test-Path $codePath) {
        $code = Get-Content $codePath -Raw -Encoding UTF8
        $code = $code -replace "(?s)^Attribute[^\r\n]*\r?\n", ""
        while ($code -match "^(Attribute[^\r\n]*\r?\n)") {
            $code = $code -replace "^(Attribute[^\r\n]*\r?\n)", ""
        }
        $mod = $wb.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
        if ($mod.CountOfLines -gt 0) { $mod.DeleteLines(1, $mod.CountOfLines) }
        $mod.AddFromString($code)
    }

    $wb.SaveAs($xlsmPath, 52)
    Write-Host "Created: $xlsmPath"
}
catch {
    Write-Host "Build failed: $_"
    Write-Host "Manual import: see docs/SETUP.md"
    exit 1
}
finally {
    if ($null -ne $wb) { $wb.Close($false) | Out-Null }
    $excel.Quit() | Out-Null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
