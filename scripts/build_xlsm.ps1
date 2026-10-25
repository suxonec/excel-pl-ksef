# Windows only: build the release workbook excel/dist/KSeF_IZ_Tool.xlsm (no data) from the
# template and the VBA sources. The Windows installer bundles this file.
#
#   powershell -ExecutionPolicy Bypass -File scripts\build_xlsm.ps1
#
# Needs Excel 365 and "Trust access to the VBA project object model"
# (File > Options > Trust Center > Trust Center Settings > Macro Settings) while building.
# Afterwards: python scripts\verify_xlsm_vba.py excel\dist\KSeF_IZ_Tool.xlsm
param(
    [string]$Template = "excel\KSeF_IZ_Tool_template.xlsx",
    [string]$Out = "excel\dist\KSeF_IZ_Tool.xlsm"
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$outPath = Join-Path $root $Out
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $outPath) | Out-Null
if (Test-Path $outPath) { Remove-Item $outPath }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false          # do not run Workbook_Open while building
try {
    $wb = $excel.Workbooks.Open((Join-Path $root $Template))
    try {
        $components = $wb.VBProject.VBComponents
    } catch {
        throw "Excel blocks VBA project access. Enable 'Trust access to the VBA project object model'."
    }
    foreach ($f in Get-ChildItem (Join-Path $root "excel\vba\*.bas") | Sort-Object Name) {
        $components.Import($f.FullName) | Out-Null
        Write-Host "imported $($f.Name)"
    }
    # Workbook events must live in the ThisWorkbook document module
    $doc = $wb.VBProject.VBComponents | Where-Object { $_.Type -eq 100 -and $_.Name -eq $wb.CodeName }
    $code = Get-Content -Raw (Join-Path $root "excel\vba\ThisWorkbook.vba")
    $doc.CodeModule.DeleteLines(1, $doc.CodeModule.CountOfLines)
    $doc.CodeModule.AddFromString($code)
    Write-Host "injected ThisWorkbook.vba into $($doc.Name)"

    $excel.Run("SetupWorkbookQuiet")
    $wb.SaveAs($outPath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
    Write-Host "wrote $outPath"
} finally {
    $excel.EnableEvents = $true
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
