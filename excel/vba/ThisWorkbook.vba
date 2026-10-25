' ThisWorkbook document module. scripts/build_xlsm.ps1 injects this code (an imported
' .bas module cannot hold workbook events). Keep it minimal: everything else is in modSetup.
Option Explicit

Private Sub Workbook_Open()
    AutoConfigure
End Sub
