Attribute VB_Name = "modSelfTest"
' Parity check against the Python implementation: runs tests/vectors.csv and
' tests/vectors_detect.csv, embedded by scripts/build_workbook.py in the hidden sheet
' _Vectors (A:F check vectors, H:I detect vectors, counts in K1 and K2).
Option Explicit

Public Sub RunSelfTest()
    Dim pass As Long, total As Long, fails As String
    SelfTestRun pass, total, fails
    If fails = "" Then
        MsgBox "Self-test passed: " & pass & "/" & total & " vectors.", vbInformation
    Else
        Debug.Print fails
        MsgBox "Self-test FAILED: " & pass & "/" & total & " passed." & vbLf & vbLf & _
               Left$(fails, 900), vbExclamation
    End If
End Sub

' Runs all vectors without dialogs (also used on workbook open). fails = "" when all pass.
Public Sub SelfTestRun(ByRef pass As Long, ByRef total As Long, ByRef fails As String)
    Dim ws As Worksheet, n1 As Long, n2 As Long, r As Long
    Dim kind As String, norm As String, crcOk As Boolean, ok As Boolean, reason As String
    Dim got As String, expected As String, hits As Collection, h As Variant

    pass = 0: fails = ""
    Set ws = ThisWorkbook.Worksheets("_Vectors")
    n1 = CLng(ws.Range("K1").Value)
    n2 = CLng(ws.Range("K2").Value)

    For r = 2 To n1 + 1
        CheckId CellText(ws.Cells(r, 1).Value), kind, norm, crcOk, ok, reason
        got = kind & "|" & norm & "|" & IIf(crcOk, "Y", "N") & "|" & IIf(ok, "Y", "N") & "|" & reason
        expected = CellText(ws.Cells(r, 2).Value) & "|" & CellText(ws.Cells(r, 3).Value) & "|" & _
                   CellText(ws.Cells(r, 4).Value) & "|" & CellText(ws.Cells(r, 5).Value) & "|" & _
                   CellText(ws.Cells(r, 6).Value)
        If got = expected Then
            pass = pass + 1
        Else
            fails = fails & "check row " & r & ": got " & got & " expected " & expected & vbLf
        End If
    Next r

    For r = 2 To n2 + 1
        Set hits = DetectIds(CellText(ws.Cells(r, 8).Value))
        got = ""
        For Each h In hits
            CheckId CStr(h(0)), kind, norm, crcOk, ok, reason
            If got <> "" Then got = got & "|"
            got = got & norm & ":" & IIf(ok, "Y", "N")
        Next h
        expected = CellText(ws.Cells(r, 9).Value)
        If got = expected Then
            pass = pass + 1
        Else
            fails = fails & "detect row " & r & ": got " & got & " expected " & expected & vbLf
        End If
    Next r

    total = n1 + n2
    ws.Range("K3").Value = pass & "/" & total & " passed " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    LogOperation "self-test", 0, pass & "/" & total & " passed"
End Sub
