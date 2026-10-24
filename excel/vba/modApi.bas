Attribute VB_Name = "modApi"
' Buttons for the API modules (M3 decode, M5 generate, M6 lookup, M7 query) via modHelper.
' Layouts match scripts/build_workbook.py and the *_FIELDS lists in src/ksefiz/operations.py:
'   2_Decode    A iz_to_decode                        -> C:K  decode rows
'   4_Generate  A:F seller_nip..group_key             -> H:L  group_key, iz, count, status, code
'   5_Lookup    A ksef_number                         -> C:H  lookup rows
'   6_Query     B1 from, B2 to, B3 own_only (TRUE)    -> A6:D query rows (header in row 5)
' Results are written as values only.
Option Explicit

Public Sub TakeIzFromDetect()
    Dim src As Worksheet, dst As Worksheet, last As Long, r As Long, n As Long
    Dim seen As New Collection, iz As String, out() As Variant
    Set src = ThisWorkbook.Worksheets("1_Detect")
    Set dst = ThisWorkbook.Worksheets("2_Decode")
    last = src.Cells(src.Rows.Count, 8).End(xlUp).Row
    ReDim out(1 To 1, 1 To 1)
    For r = 2 To last
        If CellText(src.Cells(r, 9).Value) = "IZ" And CellText(src.Cells(r, 13).Value) = "Y" Then
            iz = CellText(src.Cells(r, 11).Value)
            On Error Resume Next
            seen.Add iz, iz
            If Err.Number = 0 Then
                n = n + 1
                ReDim Preserve out(1 To 1, 1 To n)
                out(1, n) = iz
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next r
    dst.Range("A2:A" & dst.Rows.Count).ClearContents
    If n > 0 Then
        dst.Range("A2:A" & n + 1).NumberFormat = "@"
        dst.Range("A2:A" & n + 1).Value = TransposeRows(out, n, 1)
    End If
    Application.StatusBar = n & " valid IZ copied from 1_Detect to 2_Decode!A"
End Sub

Public Sub RunDecode()
    Dim ws As Worksheet, data As Variant
    Set ws = ThisWorkbook.Worksheets("2_Decode")
    data = ColumnData(ws, 1)
    If IsEmpty(data) Then
        MsgBox "Paste IZ into 2_Decode!A2 (or click 'Take IZ from 1_Detect').", vbInformation
        Exit Sub
    End If
    RunAndWrite "decode", Array("iz"), data, Empty, ws, "C2", 9, Array(3), "C:D"
End Sub

Public Sub RunGenerateCheck()
    RunGenerateOp "check"
End Sub

Public Sub RunGenerate()
    If MsgBox("Generate collective identifiers in KSeF for all groups that pass the checks?", _
              vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    RunGenerateOp "generate"
End Sub

Private Sub RunGenerateOp(ByVal op As String)
    Dim ws As Worksheet, last As Long
    Set ws = ThisWorkbook.Worksheets("4_Generate")
    last = LastUsedRow(ws, 1, 6)
    If last < 2 Then
        MsgBox "Paste invoices into 4_Generate!A2:F (seller_nip, ksef_number, amount, " & _
               "currency, description, group_key).", vbInformation
        Exit Sub
    End If
    RunAndWrite op, Array("seller_nip", "ksef_number", "amount", "currency", "description", _
                "group_key"), ws.Range("A2:F" & last).Value, Empty, ws, "H2", 5, Array(3), "H:I"
End Sub

Public Sub RunLookup()
    Dim ws As Worksheet, data As Variant
    Set ws = ThisWorkbook.Worksheets("5_Lookup")
    data = ColumnData(ws, 1)
    If IsEmpty(data) Then
        MsgBox "Paste KSeF numbers into 5_Lookup!A2.", vbInformation
        Exit Sub
    End If
    RunAndWrite "lookup", Array("ksef_number"), data, Empty, ws, "C2", 6, Empty, "C:D"
End Sub

Public Sub RunQuery()
    Dim ws As Worksheet, d1 As String, d2 As String, extra As Variant
    Set ws = ThisWorkbook.Worksheets("6_Query")
    d1 = IsoDate(ws.Range("B1").Value)
    d2 = IsoDate(ws.Range("B2").Value)
    If d1 = "" Or d2 = "" Then
        MsgBox "Enter dates in 6_Query!B1 (from) and B2 (to), e.g. 2026-10-01.", vbInformation
        Exit Sub
    End If
    If IsTrue(ws.Range("B3").Value) Then
        extra = Array("--from", d1, "--to", d2, "--own-only")
    Else
        extra = Array("--from", d1, "--to", d2)
    End If
    RunAndWrite "query", Empty, Empty, extra, ws, "A6", 4, Array(3), "A:A"
End Sub

' ---------- shared ----------

Private Sub RunAndWrite(ByVal op As String, ByVal header As Variant, ByVal data As Variant, _
                        ByVal extra As Variant, ByVal ws As Worksheet, ByVal anchor As String, _
                        ByVal nCols As Long, ByVal numericCols As Variant, ByVal textCols As String)
    Dim outValues As Variant, res As Collection, ok As Boolean, a As Range, n As Long
    Dim status As String, msg As String

    ok = RunHelperJob(op, header, data, extra, outValues, res)
    If res Is Nothing Then Exit Sub
    status = ResultValue(res, "status")
    If status = "" Then Exit Sub  ' not started (message already shown)

    Set a = ws.Range(anchor)
    ws.Range(a, ws.Cells(ws.Rows.Count, a.Column + nCols - 1)).ClearContents
    If Not IsEmpty(outValues) Then
        n = UBound(outValues, 1)
        ToNumbers outValues, numericCols
        ws.Range(textCols).NumberFormat = "@"
        a.Resize(n, nCols).Value = outValues
    End If

    LogOperation "excel-" & op, CLng(Val(ResultValue(res, "requests"))), _
                 status & ": " & ResultValue(res, "message"), ResultValue(res, "error_code"), _
                 ResultValue(res, "input_hash"), ResultValue(res, "trace_id")
    If ResultValue(res, "budget_hour") <> "" Then
        SetCfgValue "budget_left", ResultValue(res, "budget_minute") & "/min, " & _
                    ResultValue(res, "budget_hour") & "/hour at " & Format$(Now, "hh:nn")
        Application.StatusBar = "KSeF " & op & ": " & status & " - collectiveIdentifier budget " & _
                                ResultValue(res, "budget_hour") & "/hour left"
    End If

    msg = "KSeF " & op & ": " & status & vbLf & ResultValue(res, "message")
    If ResultValue(res, "error_code") <> "" Then msg = msg & vbLf & "Error code: " & ResultValue(res, "error_code")
    If ResultValue(res, "trace_id") <> "" Then msg = msg & vbLf & "Trace id: " & ResultValue(res, "trace_id")
    If ok And status = "ok" Then
        Application.StatusBar = "KSeF " & op & ": " & ResultValue(res, "message")
    Else
        MsgBox msg, IIf(status = "cancelled", vbInformation, vbExclamation)
    End If
End Sub

' Column c from row 2 down as a 2D (n x 1) array, or Empty.
Private Function ColumnData(ByVal ws As Worksheet, ByVal c As Long) As Variant
    Dim last As Long, v As Variant
    last = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    If last < 2 Then Exit Function
    If last = 2 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = ws.Cells(2, c).Value
        ColumnData = v
    Else
        ColumnData = ws.Range(ws.Cells(2, c), ws.Cells(last, c)).Value
    End If
End Function

' Turn text amounts / counts in the given 1-based columns into numbers.
Private Sub ToNumbers(ByRef v As Variant, ByVal cols As Variant)
    Dim r As Long, c As Variant, t As String
    If IsEmpty(cols) Then Exit Sub
    For r = LBound(v, 1) To UBound(v, 1)
        For Each c In cols
            t = CellText(v(r, c))
            If t <> "" And IsNumeric(Replace(t, ".", Application.International(xlDecimalSeparator))) Then
                v(r, c) = Val(t)
            End If
        Next c
    Next r
End Sub

' TRUE cell or the text "TRUE" (CStr of a Boolean is localised, e.g. "Prawda").
Private Function IsTrue(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        IsTrue = v
    Else
        IsTrue = (UCase$(Trim$(CellText(v))) = "TRUE")
    End If
End Function

Private Function IsoDate(ByVal v As Variant) As String
    If IsDate(v) Then
        IsoDate = Format$(CDate(v), "yyyy-mm-dd")
    ElseIf CellText(v) Like "####-##-##" Then
        IsoDate = CellText(v)
    End If
End Function
