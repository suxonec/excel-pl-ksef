# Windows only: build excel/KSeF_IZ_Tool.xlsm from the template and the .bas modules.
# Needs Excel 365 and "Trust access to the VBA project object model"
# (File > Options > Trust Center > Trust Center Settings > Macro Settings).
param(
    [string]$Template = "excel\KSeF_IZ_Tool_template.xlsx",
    [string]$Out = "excel\KSeF_IZ_Tool.xlsm"
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open((Join-Path $root $Template))
    try {
        $components = $wb.VBProject.VBComponents
    } catch {
        throw "Excel blocks VBA project access. Enable 'Trust access to the VBA project object model'."
    }
    foreach ($f in Get-ChildItem (Join-Path $root "excel\vba\*.bas")) {
        $components.Import($f.FullName) | Out-Null
    }
    $excel.Run("SetupWorkbookQuiet")
    $wb.SaveAs((Join-Path $root $Out), 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
    Write-Host "wrote $Out"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
