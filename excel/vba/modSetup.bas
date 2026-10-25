Attribute VB_Name = "modSetup"
' One-time setup after importing the .bas modules: run SetupWorkbook, then save as .xlsm.
' AutoConfigure runs on every open (ThisWorkbook.Workbook_Open, see ThisWorkbook.vba):
' Windows settings from the setup assistant -> Config, buttons, quiet self-test.
Option Explicit

Public Sub SetupWorkbook()
    AddAllButtons
    MsgBox "Buttons added. Save as Excel Macro-Enabled Workbook (.xlsm).", vbInformation
End Sub

' No dialog, for scripts/build_xlsm.ps1.
Public Sub SetupWorkbookQuiet()
    AddAllButtons
End Sub

Public Sub AutoConfigure()
    Dim wasSaved As Boolean, pass As Long, total As Long, fails As String, ws As Worksheet
    On Error GoTo Failed
    wasSaved = ThisWorkbook.Saved
    ApplyUserSettings
    If ThisWorkbook.Worksheets("README").Buttons.Count = 0 Then AddAllButtons
    SelfTestRun pass, total, fails
    Set ws = ThisWorkbook.Worksheets("README")
    ws.Range("A5").Value = "Self-test " & pass & "/" & total & IIf(fails = "", " passed", " FAILED") & _
                           " on open, " & Format$(Now, "yyyy-mm-dd hh:nn")
    ws.Range("A5").Font.Color = IIf(fails = "", RGB(55, 86, 35), RGB(192, 0, 0))
    ThisWorkbook.Saved = wasSaved   ' opening alone should not prompt "save changes?"
    If fails <> "" Then
        MsgBox "KSeF IZ Tool self-test FAILED (" & pass & "/" & total & "). The VBA modules " & _
               "do not match this version - reinstall or re-import them.", vbExclamation
    End If
    Exit Sub
Failed:
    Application.StatusBar = "KSeF IZ Tool: automatic setup skipped (" & Err.Description & ")"
End Sub

' Windows: copy environment / NIP / profile written by the setup assistant
' (%USERPROFILE%\.ksefiz\excel.ini) into Config, and default helper_path to the installed
' ksefiz.exe. On Mac (sandboxed Excel cannot read the home folder) Config is edited by hand.
Private Sub ApplyUserSettings()
#If Mac Then
    Exit Sub
#Else
    Dim path As String, f As Integer, txt As String, p As Long, key As String
    path = Environ("USERPROFILE") & "\.ksefiz\excel.ini"
    If FileExists(path) Then
        f = FreeFile
        Open path For Input As #f
        Do While Not EOF(f)
            Line Input #f, txt
            p = InStr(txt, "=")
            If p > 1 And Left$(txt, 1) <> ";" Then
                key = LCase$(Trim$(Left$(txt, p - 1)))
                If key = "environment" Or key = "context_nip" Or key = "profile" Then
                    SetCfgValue key, Trim$(Mid$(txt, p + 1))
                End If
            End If
        Loop
        Close #f
    End If
    If CfgValue("helper_path") = "" Or CfgValue("helper_path") = "ksefiz" Then
        If FileExists(InstalledHelperPath()) Then SetCfgValue "helper_path", InstalledHelperPath()
    End If
#End If
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
