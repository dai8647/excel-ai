#Requires -Version 5.1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlsmPath = Join-Path $root "ExcelAI.xlsm"
$vbaDir = Join-Path $root "vba"
$navy = 5058583
$deepNavy = 3342091
$blue = 11561003
$teal = 7366159
$green = 4879126
$amber = 423385
$lightBlue = 16317162
$pale = 16513783
$inputFill = 15132671
$line = 14735307
$muted = 9122404
$white = 16777215
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

    while (($wb.Worksheets.Count) -lt 5) {
        $wb.Worksheets.Add() | Out-Null
    }

    $wb.Worksheets.Item(1).Name = "README"
    $wb.Worksheets.Item(2).Name = [char]0x8A2D + [char]0x5B9A
    $wb.Worksheets.Item(3).Name = [char]0x64CD + [char]0x4F5C
    $wb.Worksheets.Item(4).Name = -join @([char]0x30C1, [char]0x30E3, [char]0x30C3, [char]0x30C8, [char]0x30ED, [char]0x30B0)
    $wb.Worksheets.Item(5).Name = -join @([char]0x30E2, [char]0x30C7, [char]0x30EB, [char]0x4E00, [char]0x89A7)

    $readme = $wb.Worksheets.Item(1)
    $readme.Tab.Color = 5296274
    $readme.Range("A1:H1").Merge()
    $readme.Cells.Item(1, 1).Value2 = "ExcelAI | README"
    $readme.Range("A1:H1").Interior.Color = $navy
    $readme.Range("A1:H1").Font.Color = $white
    $readme.Range("A1:H1").Font.Bold = $true
    $readme.Range("A1:H1").Font.Size = 18
    $readme.Range("A2:H2").Merge()
    $readme.Cells.Item(2, 1).Value2 = "Start here: set your provider, test the connection, then run a request."
    $readme.Range("A2:H2").Interior.Color = $pale
    $readme.Range("A2:H2").Font.Color = $muted
    $readme.Range("A2:H2").Font.Size = 11
    $readme.Columns.Item("A").ColumnWidth = 18
    $readme.Columns.Item("B").ColumnWidth = 24
    foreach ($column in @("C", "D", "E", "F", "G", "H")) { $readme.Columns.Item($column).ColumnWidth = 16 }
    $readme.Range("A1:H60").VerticalAlignment = -4160
    $readme.Range("A1:H60").WrapText = $true
    $readmeLinesPath = Join-Path $root "docs\README_SHEET.txt"
    if (-not (Test-Path $readmeLinesPath)) { throw "README sheet content was not found: $readmeLinesPath" }
    $readmeRow = 4
    foreach ($line in ((Get-Content $readmeLinesPath -Raw -Encoding UTF8) -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($line)) { $readmeRow += 1; continue }
        $parts = $line.Split('|')
        switch ($parts[0]) {
            "SECTION" {
                $readme.Range("A${readmeRow}:H${readmeRow}").Merge()
                $readme.Cells.Item($readmeRow, 1).Value2 = $parts[1]
                $readme.Range("A${readmeRow}:H${readmeRow}").Interior.Color = $blue
                $readme.Range("A${readmeRow}:H${readmeRow}").Font.Color = $white
                $readme.Range("A${readmeRow}:H${readmeRow}").Font.Bold = $true
                $readme.Rows.Item($readmeRow).RowHeight = 24
                $readmeRow += 1
            }
            "STEP" {
                $readme.Cells.Item($readmeRow, 1).Value2 = "STEP " + $parts[1]
                $readme.Cells.Item($readmeRow, 2).Value2 = $parts[2]
                $readme.Range("C${readmeRow}:H${readmeRow}").Merge()
                $readme.Cells.Item($readmeRow, 3).Value2 = $parts[3]
                $readme.Range("A${readmeRow}:B${readmeRow}").Interior.Color = $lightBlue
                $readme.Range("C${readmeRow}:H${readmeRow}").Interior.Color = $pale
                $readme.Range("A${readmeRow}:B${readmeRow}").Font.Bold = $true
                $readme.Rows.Item($readmeRow).RowHeight = 38
                $readmeRow += 1
            }
            "ITEM" {
                $readme.Cells.Item($readmeRow, 1).Value2 = $parts[1]
                $readme.Cells.Item($readmeRow, 2).Value2 = $parts[2]
                $readme.Range("C${readmeRow}:H${readmeRow}").Merge()
                $readme.Cells.Item($readmeRow, 3).Value2 = $parts[3]
                $readme.Range("A${readmeRow}").Interior.Color = $lightBlue
                $readme.Range("B${readmeRow}").Interior.Color = $inputFill
                $readme.Range("C${readmeRow}:H${readmeRow}").Interior.Color = $pale
                $readme.Cells.Item($readmeRow, 1).Font.Bold = $true
                $readme.Rows.Item($readmeRow).RowHeight = 38
                $readmeRow += 1
            }
            "NOTE" {
                $readme.Range("A${readmeRow}:H${readmeRow}").Merge()
                $readme.Cells.Item($readmeRow, 1).Value2 = $parts[1]
                $readme.Range("A${readmeRow}:H${readmeRow}").Interior.Color = $pale
                $readme.Range("A${readmeRow}:H${readmeRow}").Font.Color = $muted
                $readme.Rows.Item($readmeRow).RowHeight = 30
                $readmeRow += 1
            }
        }
    }
    $readme.Range("A4:H$($readmeRow - 1)").Borders.LineStyle = 1
    $readme.Rows.Item(1).RowHeight = 30
    $readme.Rows.Item(2).RowHeight = 24

    $ws = $wb.Worksheets.Item(2)
    $ws.Tab.Color = $blue
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
    $ws.Columns.Item("A").ColumnWidth = 24; $ws.Columns.Item("B").ColumnWidth = 48; $ws.Columns.Item("C").ColumnWidth = 56
    $ws.Range("A1:C1").Interior.Color = $navy; $ws.Range("A1:C1").Font.Color = $white; $ws.Range("A1:C1").Font.Bold = $true; $ws.Range("A1:C1").Font.Size = 16
    $ws.Range("A2:A10").Interior.Color = $lightBlue; $ws.Range("A2:A10").Font.Bold = $true; $ws.Range("B2:B10").Interior.Color = $inputFill; $ws.Range("B2:B10").Font.Color = $deepNavy; $ws.Range("C2:C10").Interior.Color = $pale; $ws.Range("C2:C10").Font.Color = $muted; $ws.Range("A2:C10").Borders.LineStyle = 1
    $ws.Rows("1:10").RowHeight = 23; $ws.Rows.Item(1).RowHeight = 30
    $ws.Range("A12:C12").Merge(); $ws.Cells.Item(12, 1).Value2 = "Input cells are pale yellow. Test the connection after changing settings."; $ws.Range("A12:C12").Interior.Color = $pale; $ws.Range("A12:C12").Font.Color = $muted; $ws.Rows.Item(12).RowHeight = 24
    $ws.Range("B2").Validation.Add(3, 1, 1, "OpenAI,Anthropic")
    $ws.Range("B7:B8").Validation.Add(3, 1, 1, "yes,no")

    $ops = $wb.Worksheets.Item(3)
    $ops.Tab.Color = $teal
    $ops.Range("A1:H1").Merge(); $ops.Cells.Item(1, 1).Value2 = "ExcelAI | Operation Center"
    $ops.Range("A2:H2").Merge(); $ops.Cells.Item(2, 1).Value2 = "Check settings -> test connection -> enter Request -> run"
    $ops.Range("A1:H1").Interior.Color = $navy; $ops.Range("A1:H1").Font.Color = $white; $ops.Range("A1:H1").Font.Bold = $true; $ops.Range("A1:H1").Font.Size = 18
    $ops.Range("A2:H2").Interior.Color = $pale; $ops.Range("A2:H2").Font.Color = $muted; $ops.Range("A2:H2").Font.Size = 11
    $ops.Columns.Item("A").ColumnWidth = 26; $ops.Columns.Item("B").ColumnWidth = 68
    foreach ($column in @("C", "D", "E", "F", "G", "H")) { $ops.Columns.Item($column).ColumnWidth = 14 }
    $ops.Rows("1:2").RowHeight = 24; $ops.Rows.Item(1).RowHeight = 32
    $buttonSpecs = @(
        @("1. Refresh models", "RefreshModels", $blue),
        @("2. Test connection", "TestConnection", $teal),
        @("3. Open request panel", "ShowAgent", $navy),
        @("4. Run request", "RunSheetPrompt", $green),
        @("Open workspace", "OpenWorkspaceFolder", $amber),
        @("Clear chat log", "ClearChatLog", $muted),
        @("Workspace safety", "WorkspaceSafetyCheck", $teal)
    )
    $buttonIndex = 0
    foreach ($spec in $buttonSpecs) {
        $buttonRow = [math]::Floor($buttonIndex / 4)
        $buttonColumn = $buttonIndex % 4
        $button = $ops.Buttons().Add((20 + ($buttonColumn * 240)), (60 + ($buttonRow * 42)), 220, 30)
        $button.Characters().Text = $spec[0]
        $button.OnAction = $spec[1]
        try {
            $button.ShapeRange.Fill.ForeColor.RGB = [int]$spec[2]
            $button.ShapeRange.Line.ForeColor.RGB = [int]$spec[2]
            $button.Characters().Font.Color = $deepNavy
            $button.Characters().Font.Bold = $true
            $button.Characters().Font.Size = 11
        }
        catch { }
        try { $button.Font.ColorIndex = 1; $button.Characters().Font.ColorIndex = 1; $button.Characters().Font.Bold = $true; $button.Characters().Font.Size = 11 } catch { }
        $buttonIndex += 1
    }
    $ops.Range("A13:B13").Merge(); $ops.Cells.Item(13, 1).Value2 = "LATEST CHECK"
    $ops.Cells.Item(14, 1).Value2 = "Time"; $ops.Cells.Item(15, 1).Value2 = "Result"
    $ops.Range("A13:B13").Interior.Color = $navy; $ops.Range("A13:B13").Font.Color = $white; $ops.Range("A13:B13").Font.Bold = $true
    $ops.Range("A14:A15").Interior.Color = $lightBlue; $ops.Range("A14:A15").Font.Bold = $true; $ops.Range("B14:B15").Interior.Color = $pale
    $ops.Range("A13:B15").Borders.LineStyle = 1
    $ops.Rows.Item(13).RowHeight = 24; $ops.Rows.Item(14).RowHeight = 24; $ops.Rows.Item(15).RowHeight = 54
    $ops.Cells.Item(7, 1).Value2 = "Request"
    $ops.Cells.Item(7, 2).Value2 = ""
    $ops.Cells.Item(9, 1).Value2 = "Result"
    $ops.Range("A7:A7").Interior.Color = $blue; $ops.Range("A7:A7").Font.Color = $white; $ops.Range("A7:A7").Font.Bold = $true
    $ops.Range("A9:A10").Merge(); $ops.Range("A9:A10").Interior.Color = $blue; $ops.Range("A9:A10").Font.Color = $white; $ops.Range("A9:A10").Font.Bold = $true
    $ops.Range("B7:H7").Merge(); $ops.Range("B9:H10").Merge()
    $ops.Range("B7:H7").Interior.Color = $inputFill; $ops.Range("B9:H10").Interior.Color = $pale
    $ops.Range("A7:H10").Borders.LineStyle = 1
    $ops.Range("B7:H7").WrapText = $true; $ops.Range("B9:H10").WrapText = $true
    $ops.Rows.Item(7).RowHeight = 30; $ops.Rows("9:10").RowHeight = 28

    $log = $wb.Worksheets.Item(4)
    $log.Tab.Color = $muted
    $log.Cells.Item(1, 1).Value2 = "Time"; $log.Cells.Item(1, 2).Value2 = "Type"; $log.Cells.Item(1, 3).Value2 = "Content"
    $log.Range("A1:C1").Interior.Color = $navy; $log.Range("A1:C1").Font.Color = $white; $log.Range("A1:C1").Font.Bold = $true
    $log.Columns.Item("A").ColumnWidth = 22; $log.Columns.Item("B").ColumnWidth = 16; $log.Columns.Item("C").ColumnWidth = 96; $log.Rows.Item(1).RowHeight = 24; $log.Range("A1:C1").AutoFilter()

    $models = $wb.Worksheets.Item(5)
    $models.Tab.Color = $amber
    $models.Range("A1:C1").Merge(); $models.Cells.Item(1, 1).Value2 = "Model catalog"
    $models.Range("A3:C3").Merge(); $models.Cells.Item(3, 1).Value2 = "Refresh from the Settings sheet, then choose a model there."
    $models.Cells.Item(5, 1).Value2 = "Model ID"; $models.Cells.Item(5, 2).Value2 = "Connection type"
    $models.Range("A1:C1").Interior.Color = $navy; $models.Range("A1:C1").Font.Color = $white; $models.Range("A1:C1").Font.Bold = $true; $models.Range("A1:C1").Font.Size = 16
    $models.Range("A3:C3").Interior.Color = $pale; $models.Range("A3:C3").Font.Color = $muted
    $models.Range("A5:B5").Interior.Color = $blue; $models.Range("A5:B5").Font.Color = $white; $models.Range("A5:B5").Font.Bold = $true
    $models.Columns.Item("A").ColumnWidth = 58; $models.Columns.Item("B").ColumnWidth = 22; $models.Columns.Item("C").ColumnWidth = 14
    $models.Rows.Item(1).RowHeight = 30; $models.Rows.Item(3).RowHeight = 24; $models.Rows.Item(5).RowHeight = 24; $models.Range("A5:B5").AutoFilter()
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

    foreach ($sheet in $wb.Worksheets) {
        $sheet.Activate()
        try { $excel.ActiveWindow.DisplayGridlines = $false; $excel.ActiveWindow.Zoom = 90 } catch { }
    }
    $readme.Activate()
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
