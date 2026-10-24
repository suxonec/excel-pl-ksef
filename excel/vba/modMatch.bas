Attribute VB_Name = "modMatch"
' M4: join decoded IZ rows (sheet 2_Decode C:H) to ERP open items (sheet 3_Match A:F)
' on KSeF number. Mirror of src/ksefiz/match.py - keep statuses identical.
' 2_Decode  C iz, D ksef_number, E amount, F currency, G description, H details_hidden
' 3_Match   A doc_no, B customer, C ksef_number, D amount, E currency, F due_date
' Output    3_Match H:S, one row per decoded row (values only)
Option Explicit

Private Const TOLERANCE As Double = 0.005

Public Sub RunMatch()
    Dim wsD As Worksheet, wsM As Worksheet, dec As Variant, erp As Variant
    Dim nD As Long, nE As Long, i As Long, j As Long, idx As Long, k As String
    Dim erpMap As New Collection, erpFirst() As Long, erpCount() As Long, nKeys As Long
    Dim decMap As New Collection, refs() As Long, total() As Double, totalNull() As Boolean
    Dim decKey() As String, nDecKeys As Long, out() As Variant, cnt As Long
    Dim paid As Variant, openAmt As Variant, diff As Variant, tot As Variant
    Dim cur As String, openCur As String, status As String, hit As Long

    Set wsD = ThisWorkbook.Worksheets("2_Decode")
    Set wsM = ThisWorkbook.Worksheets("3_Match")
    wsM.Range("H2:S" & wsM.Rows.Count).ClearContents
    nD = LastUsedRow(wsD, 3, 8) - 1
    nE = LastUsedRow(wsM, 1, 6) - 1
    If nD < 1 Then
        MsgBox "No decoded rows in 2_Decode!C2:H (run Decode first).", vbInformation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    dec = wsD.Range("C2:H" & nD + 1).Value
    If nE >= 1 Then erp = wsM.Range("A2:F" & nE + 1).Value

    ' ERP index: key -> first row, count
    ReDim erpFirst(1 To nE + 1): ReDim erpCount(1 To nE + 1)
    For i = 1 To nE
        If Not RowBlank(erp, i) Then
            k = KeyOf(CellText(erp(i, 3)))
            idx = KeyIndex(erpMap, k)
            If idx = 0 Then
                nKeys = nKeys + 1
                erpMap.Add nKeys, "k" & k
                erpFirst(nKeys) = i
                erpCount(nKeys) = 1
            Else
                erpCount(idx) = erpCount(idx) + 1
            End If
        End If
    Next i

    ' Decoded: references and total paid per KSeF number
    ReDim refs(1 To nD): ReDim total(1 To nD): ReDim totalNull(1 To nD): ReDim decKey(1 To nD)
    For j = 1 To nD
        If Not RowBlank(dec, j) Then
            decKey(j) = KeyOf(CellText(dec(j, 2)))
            paid = ParseAmount(dec(j, 3))
            idx = KeyIndex(decMap, decKey(j))
            If idx = 0 Then
                nDecKeys = nDecKeys + 1
                idx = nDecKeys
                decMap.Add idx, "k" & decKey(j)
            End If
            refs(idx) = refs(idx) + 1
            If IsNull(paid) Or totalNull(idx) Then
                totalNull(idx) = True
            Else
                total(idx) = total(idx) + CDbl(paid)
            End If
        End If
    Next j

    ReDim out(1 To 12, 1 To 1)
    For j = 1 To nD
        If Not RowBlank(dec, j) Then
            idx = KeyIndex(decMap, decKey(j))
            paid = ParseAmount(dec(j, 3))
            cur = UCase$(Trim$(CellText(dec(j, 4))))
            hit = KeyIndex(erpMap, decKey(j))
            openAmt = Null: openCur = ""
            If hit > 0 Then
                openAmt = ParseAmount(erp(erpFirst(hit), 4))
                openCur = UCase$(Trim$(CellText(erp(erpFirst(hit), 5))))
            End If
            If totalNull(idx) Then tot = Null Else tot = total(idx)
            If IsNull(tot) Or IsNull(openAmt) Then
                diff = Null
            Else
                diff = Round(CDbl(tot) - CDbl(openAmt), 2)
            End If

            If hit = 0 Then
                status = "Unmatched"
            ElseIf erpCount(hit) > 1 Then
                status = "Duplicate in ERP"
            ElseIf IsNull(tot) Or IsNull(openAmt) Then
                status = "Matched (no amount)"
            ElseIf cur <> "" And openCur <> "" And cur <> openCur Then
                status = "Currency diff"
            ElseIf Abs(CDbl(diff)) > TOLERANCE Then
                status = "Amount diff"
            Else
                status = "Matched"
            End If

            cnt = cnt + 1
            ReDim Preserve out(1 To 12, 1 To cnt)
            out(1, cnt) = Trim$(CellText(dec(j, 1)))
            out(2, cnt) = decKey(j)
            out(3, cnt) = NullToBlank(paid)
            out(4, cnt) = cur
            If hit > 0 Then
                out(5, cnt) = erp(erpFirst(hit), 1)
                out(6, cnt) = erp(erpFirst(hit), 2)
            End If
            out(7, cnt) = NullToBlank(openAmt)
            out(8, cnt) = openCur
            out(9, cnt) = NullToBlank(diff)
            out(10, cnt) = refs(idx)
            out(11, cnt) = NullToBlank(tot)
            out(12, cnt) = status
        End If
    Next j

    If cnt > 0 Then
        wsM.Range("H2:I" & cnt + 1).NumberFormat = "@"
        wsM.Range("H2:S" & cnt + 1).Value = TransposeRows(out, cnt, 12)
    End If
    Application.ScreenUpdating = True
    LogOperation "match", 0, cnt & " decoded rows vs " & nE & " open items"
    Application.StatusBar = "Match: " & cnt & " rows written to 3_Match!H:S"
End Sub

' Same as Python: normalised KSeF number, else trimmed uppercase input.
Public Function KeyOf(ByVal s As String) As String
    Dim n As String
    n = IZ_NORMALISE(s)
    If n = "" Then n = UCase$(Trim$(s))
    KeyOf = n
End Function

' Mirror of io_csv.parse_amount: numbers pass through; text like "1 234,56",
' "1.234,56", "1,234.56" is parsed; blank or junk -> Null.
Public Function ParseAmount(ByVal v As Variant) As Variant
    Dim t As String
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong, vbDecimal, vbByte
            ParseAmount = CDbl(v)
            Exit Function
    End Select
    t = Replace(Replace(CellText(v), " ", ""), ChrW(160), "")
    t = Trim$(t)
    If t = "" Then
        ParseAmount = Null
        Exit Function
    End If
    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then
        If InStrRev(t, ",") > InStrRev(t, ".") Then
            t = Replace(Replace(t, ".", ""), ",", ".")
        Else
            t = Replace(t, ",", "")
        End If
    Else
        t = Replace(t, ",", ".")
    End If
    If PlainDecimal(t) Then ParseAmount = Val(t) Else ParseAmount = Null
End Function

' ^[+-]?(\d+\.?\d*|\.\d+)$
Private Function PlainDecimal(ByVal t As String) As Boolean
    Dim i As Long, ch As String, digitsBefore As Long, digitsAfter As Long, dot As Boolean
    i = 1
    If Left$(t, 1) = "+" Or Left$(t, 1) = "-" Then i = 2
    For i = i To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "#" Then
            If dot Then digitsAfter = digitsAfter + 1 Else digitsBefore = digitsBefore + 1
        ElseIf ch = "." And Not dot Then
            dot = True
        Else
            PlainDecimal = False
            Exit Function
        End If
    Next i
    PlainDecimal = (digitsBefore > 0 Or digitsAfter > 0)
End Function

Private Function KeyIndex(ByVal col As Collection, ByVal key As String) As Long
    On Error Resume Next
    KeyIndex = col("k" & key)
    If Err.Number <> 0 Then KeyIndex = 0
    Err.Clear
End Function

Private Function RowBlank(ByVal a As Variant, ByVal r As Long) As Boolean
    Dim c As Long
    For c = 1 To 6
        If Trim$(CellText(a(r, c))) <> "" Then
            RowBlank = False
            Exit Function
        End If
    Next c
    RowBlank = True
End Function

Private Function NullToBlank(ByVal v As Variant) As Variant
    If IsNull(v) Then NullToBlank = "" Else NullToBlank = v
End Function
