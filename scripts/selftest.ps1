# Runs the in-workbook acceptance harness (SelfTest_Run) headlessly.
# Expect: "SelfTest: N PASS / 0 FAIL"
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlsm = Join-Path $root 'ExcelAI.xlsm'
if (-not (Test-Path $xlsm)) { throw "ExcelAI.xlsm not found - run scripts/build-xlsm.ps1 first" }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try { $excel.AutomationSecurity = 1 } catch {}
try {
    $wb = $excel.Workbooks.Open($xlsm)
    $excel.Run('Session_Reset') | Out-Null
    $result = $excel.Run('SelfTest_Run')
    Write-Output $result
    $wb.Close($false)
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
if ($result -notmatch '0 FAIL') { exit 1 }
