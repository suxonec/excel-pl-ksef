Attribute VB_Name = "modSetup"
' One-time setup after importing the .bas modules: run SetupWorkbook, then save as .xlsm.
Option Explicit

Public Sub SetupWorkbook()
    AddAllButtons
    MsgBox "Buttons added. Save as Excel Macro-Enabled Workbook (.xlsm).", vbInformation
End Sub

' No dialog, for scripts/build_xlsm.ps1.
Public Sub SetupWorkbookQuiet()
    AddAllButtons
End Sub

Private Sub AddAllButtons()
    AddButton "README", "B8", "Run self-test", "RunSelfTest"
    AddButton "README", "B10", "Detect IDs (1_Detect)", "RunDetect"
    AddButton "README", "B12", "Match open items (3_Match)", "RunMatch"
    AddButton "README", "B14", "Decode IZ (2_Decode)", "RunDecode"
    AddButton "1_Detect", "P2", "Detect IDs", "RunDetect"
    AddButton "2_Decode", "M2", "Take IZ from 1_Detect", "TakeIzFromDetect"
    AddButton "2_Decode", "M4", "Decode (KSeF)", "RunDecode"
    AddButton "3_Match", "U2", "Match open items", "RunMatch"
    AddButton "4_Generate", "N2", "Check (no API)", "RunGenerateCheck"
    AddButton "4_Generate", "N4", "Generate IZ (KSeF)", "RunGenerate"
    AddButton "5_Lookup", "J2", "Look up (KSeF)", "RunLookup"
    AddButton "6_Query", "F1", "Run query (KSeF)", "RunQuery"
End Sub

Private Sub AddButton(ByVal sheetName As String, ByVal anchor As String, _
                      ByVal caption As String, ByVal macroName As String)
    Dim ws As Worksheet, c As Range, b As Object, i As Long
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For i = ws.Buttons.Count To 1 Step -1
        If ws.Buttons(i).Caption = caption Then ws.Buttons(i).Delete
    Next i
    Set c = ws.Range(anchor)
    Set b = ws.Buttons.Add(c.Left, c.Top, 180, 26)
    b.Caption = caption
    b.OnAction = macroName
End Sub
