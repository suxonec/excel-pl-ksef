Attribute VB_Name = "modSample"
' "Load sample data": copies the synthetic examples from the hidden sheet _Sample
' (written by scripts/build_workbook.py, see SAMPLE_BLOCKS there) into the input sheets,
' so a new user can try Detect, Check (no API) and Match without real data. Offline only.
'   _Sample A:F  -> 1_Detect A:F     (statement lines)
'   _Sample H:M  -> 2_Decode C:H     (decoded rows, input for 3_Match)
'   _Sample O:T  -> 3_Match A:F      (ERP open items)
'   _Sample V:AA -> 4_Generate A:F   (invoice groups showing every pre-check outcome)
' Row counts are in _Sample AC1:AC4.
Option Explicit

Public Sub LoadSampleData()
    Dim src As Worksheet, msg As String
    Set src = ThisWorkbook.Worksheets("_Sample")

    If HasInputData() Then
        If MsgBox("Load the sample data? It replaces the current contents of 1_Detect, " & _
                  "2_Decode, 3_Match and 4_Generate (inputs and results).", _
                  vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    End If

    Application.ScreenUpdating = False
    ClearSheet "1_Detect", "A2:N"
    ClearSheet "2_Decode", "A2:K"
    ClearSheet "3_Match", "A2:S"
    ClearSheet "4_Generate", "A2:L"
    With ThisWorkbook.Worksheets("1_Detect").Range("A2:F" & ThisWorkbook.Worksheets("1_Detect").Rows.Count)
        .Interior.Pattern = xlNone            ' old red highlight
        .Font.ColorIndex = xlAutomatic
    End With
    CopyBlock src, 1, "AC1", "1_Detect", "A2"
    CopyBlock src, 8, "AC2", "2_Decode", "C2"
    CopyBlock src, 15, "AC3", "3_Match", "A2"
    CopyBlock src, 22, "AC4", "4_Generate", "A2"
    Application.ScreenUpdating = True

    LogOperation "sample-data", 0, "sample data loaded"
    msg = "Sample data loaded (synthetic, no real companies):" & vbLf & vbLf & _
          "- 1_Detect: " & src.Range("AC1").Value & " bank statement lines -> click Detect IDs" & vbLf & _
          "- 2_Decode: " & src.Range("AC2").Value & " decoded invoice rows (example KSeF output)" & vbLf & _
          "- 3_Match: " & src.Range("AC3").Value & " ERP open items -> click Match open items" & vbLf & _
          "- 4_Generate: " & src.Range("AC4").Value & " invoices -> click Check (no API)" & vbLf & vbLf & _
          "The KSeF buttons (Decode, Generate IZ, Look up) need real invoices in KSeF; " & _
          "these sample numbers do not exist there."
    MsgBox msg, vbInformation
End Sub

Private Function HasInputData() As Boolean
    HasInputData = LastUsedRow(ThisWorkbook.Worksheets("1_Detect"), 1, 6) > 1 Or _
                   LastUsedRow(ThisWorkbook.Worksheets("2_Decode"), 1, 11) > 1 Or _
                   LastUsedRow(ThisWorkbook.Worksheets("3_Match"), 1, 6) > 1 Or _
                   LastUsedRow(ThisWorkbook.Worksheets("4_Generate"), 1, 6) > 1
End Function

Private Sub ClearSheet(ByVal sheetName As String, ByVal cols As String)
    Dim ws As Worksheet, parts() As String
    Set ws = ThisWorkbook.Worksheets(sheetName)
    parts = Split(cols, ":")
    ws.Range(parts(0) & ":" & parts(1) & ws.Rows.Count).ClearContents
End Sub

' Copy a 6-column block (header in row 1) of n rows to target, as values.
Private Sub CopyBlock(ByVal src As Worksheet, ByVal firstCol As Long, ByVal countCell As String, _
                      ByVal sheetName As String, ByVal anchor As String)
    Dim n As Long
    n = CLng(Val(CellText(src.Range(countCell).Value)))
    If n < 1 Then Exit Sub
    ThisWorkbook.Worksheets(sheetName).Range(anchor).Resize(n, 6).Value = _
        src.Range(src.Cells(2, firstCol), src.Cells(n + 1, firstCol + 5)).Value
End Sub
