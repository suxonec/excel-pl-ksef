Attribute VB_Name = "modDetect"
' M1: find IZ and KSeF numbers in pasted bank statement lines (sheet 1_Detect).
' Input  A:F  line_id, date, amount, currency, payer, free_text   (row 2 down)
' Output H:N  line_id, kind, raw, normalised, crc_ok, valid, reason (values only)
' Same output as `ksefiz detect`.
Option Explicit

Public Sub RunDetect()
    Dim ws As Worksheet, lastRow As Long, data As Variant, r As Long
    Dim hits As Collection, h As Variant, out() As Variant, cnt As Long, valid As Long
    Dim kind As String, norm As String, crcOk As Boolean, ok As Boolean, reason As String

    Set ws = ThisWorkbook.Worksheets("1_Detect")
    ws.Range("H2:N" & ws.Rows.Count).ClearContents
    lastRow = LastUsedRow(ws, 1, 6)
    If lastRow < 2 Then
        MsgBox "Paste statement lines into 1_Detect!A2:F first.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    data = ws.Range("A2:F" & lastRow).Value
    ReDim out(1 To 7, 1 To 1)
    For r = 1 To UBound(data, 1)
        Set hits = DetectIds(CellText(data(r, 6)))
        For Each h In hits
            CheckId CStr(h(0)), kind, norm, crcOk, ok, reason
            cnt = cnt + 1
            ReDim Preserve out(1 To 7, 1 To cnt)
            out(1, cnt) = CellText(data(r, 1))
            out(2, cnt) = kind
            out(3, cnt) = Replace(Replace(CStr(h(0)), vbCr, " "), vbLf, " ")
            out(4, cnt) = norm
            out(5, cnt) = IIf(crcOk, "Y", "N")
            out(6, cnt) = IIf(ok, "Y", "N")
            out(7, cnt) = reason
            If ok Then valid = valid + 1
        Next h
    Next r

    If cnt > 0 Then
        ws.Range("H2:N" & cnt + 1).NumberFormat = "@"
        ws.Range("H2:N" & cnt + 1).Value = TransposeRows(out, cnt, 7)
    End If
    Application.ScreenUpdating = True
    LogOperation "detect", 0, (lastRow - 1) & " lines, " & cnt & " candidates, " & valid & " valid"
    Application.StatusBar = "Detect: " & cnt & " candidates, " & valid & " valid"
End Sub

' ---------- shared helpers (also used by modMatch) ----------

Public Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function

' Last row with any value in columns firstCol..lastCol (1 if only the header).
Public Function LastUsedRow(ByVal ws As Worksheet, ByVal firstCol As Long, _
                            ByVal lastCol As Long) As Long
    Dim c As Long, r As Long, best As Long
    best = 1
    For c = firstCol To lastCol
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > best Then best = r
    Next c
    LastUsedRow = best
End Function

' out is built column-major (1 To cols, 1 To rows) so ReDim Preserve works; flip it.
Public Function TransposeRows(ByRef out() As Variant, ByVal rows As Long, _
                              ByVal cols As Long) As Variant
    Dim a() As Variant, r As Long, c As Long
    ReDim a(1 To rows, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols
            a(r, c) = out(c, r)
        Next c
    Next r
    TransposeRows = a
End Function

' Append-only audit row on the Log sheet. Offline operations make 0 API requests.
Public Sub LogOperation(ByVal operation As String, ByVal requestCount As Long, _
                        ByVal result As String, Optional ByVal errorCode As String = "", _
                        Optional ByVal inputHash As String = "offline", _
                        Optional ByVal correlationId As String = "")
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Log")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).Value = Application.UserName
    ws.Cells(r, 3).Value = CfgValue("environment")
    ws.Cells(r, 4).Value = operation
    ws.Cells(r, 5).Value = requestCount
    ws.Cells(r, 6).Value = inputHash
    ws.Cells(r, 7).Value = result
    ws.Cells(r, 8).Value = errorCode
    ws.Cells(r, 9).Value = correlationId
End Sub
