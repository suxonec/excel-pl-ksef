Attribute VB_Name = "modIZ"
' Offline IZ / KSeF number detection, normalisation and CRC-8 validation.
' Mirror of src/ksefiz/iz.py - change both together; tests/vectors*.csv (sheet _Vectors,
' macro RunSelfTest) keep them in step. Works on Excel for Windows and Mac (no
' VBScript.RegExp, no Scripting.Dictionary).
Option Explicit

Public Const KIND_IZ As String = "IZ"
Public Const KIND_KSEF As String = "KSEF"

Private Const HEX2 As String = "[0-9A-F][0-9A-F]"
Private Const HEX6 As String = "[0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F]"

' ---------- worksheet functions ----------

' CRC-8 (poly 0x07, init 0x00) of the first 32 characters of s, as 2 hex chars.
Public Function IZ_CRC8(ByVal s As String) As String
    Dim c As Long, i As Long, b As Long, data As String
    data = Left$(s, 32)
    For i = 1 To Len(data)
        c = c Xor (AscW(Mid$(data, i, 1)) And &HFF&)
        For b = 1 To 8
            If (c And &H80&) <> 0 Then
                c = ((c * 2) Xor &H7&) And &HFF&
            Else
                c = (c * 2) And &HFF&
            End If
        Next b
    Next i
    IZ_CRC8 = Right$("0" & Hex$(c), 2)
End Function

' Canonical 35-char form, or "" if s does not hold exactly 32 significant characters.
Public Function IZ_NORMALISE(ByVal s As String) As String
    Dim c As String
    c = CompactId(s)
    If Len(c) <> 32 Then
        IZ_NORMALISE = ""
    Else
        IZ_NORMALISE = Left$(c, 10) & "-" & Mid$(c, 11, 8) & "-" & Mid$(c, 19, 12) & "-" & Right$(c, 2)
    End If
End Function

Public Function IZ_VALID(ByVal s As String) As Boolean
    Dim kind As String, norm As String, crcOk As Boolean, valid As Boolean, reason As String
    CheckId s, kind, norm, crcOk, valid, reason
    IZ_VALID = valid And kind = KIND_IZ
End Function

Public Function KSEF_VALID(ByVal s As String) As Boolean
    Dim kind As String, norm As String, crcOk As Boolean, valid As Boolean, reason As String
    CheckId s, kind, norm, crcOk, valid, reason
    KSEF_VALID = valid And kind = KIND_KSEF
End Function

' "OK", "format", "nip", "date" or "crc".
Public Function ID_REASON(ByVal s As String) As String
    Dim kind As String, norm As String, crcOk As Boolean, valid As Boolean, reason As String
    CheckId s, kind, norm, crcOk, valid, reason
    ID_REASON = reason
End Function

' "IZ", "KSEF" or "".
Public Function ID_KIND(ByVal s As String) As String
    Dim kind As String, norm As String, crcOk As Boolean, valid As Boolean, reason As String
    CheckId s, kind, norm, crcOk, valid, reason
    ID_KIND = kind
End Function

' ---------- validation ----------

Public Sub CheckId(ByVal s As String, ByRef kind As String, ByRef norm As String, _
                   ByRef crcOk As Boolean, ByRef valid As Boolean, ByRef reason As String)
    Dim fmt As String
    kind = "": crcOk = False: valid = False
    norm = IZ_NORMALISE(s)
    If norm = "" Then
        reason = "format"
        Exit Sub
    End If
    If Mid$(norm, 12, 2) = "IZ" Then
        kind = KIND_IZ
        fmt = "##########-IZ######-" & HEX6 & HEX6 & "-" & HEX2
    Else
        kind = KIND_KSEF
        fmt = "##########-########-" & HEX6 & HEX6 & "-" & HEX2
    End If
    If Not (norm Like fmt) Then
        reason = "format"
        Exit Sub
    End If
    crcOk = (IZ_CRC8(norm) = Right$(norm, 2))
    If Left$(norm, 1) = "0" Or Mid$(norm, 2, 2) = "00" Then
        reason = "nip"
    ElseIf Not DateOk(kind, norm) Then
        reason = "date"
    ElseIf Not crcOk Then
        reason = "crc"
    Else
        reason = "OK"
    End If
    valid = (reason = "OK")
End Sub

Private Function DateOk(ByVal kind As String, ByVal n As String) As Boolean
    Dim y As Long, m As Long, d As Long
    If kind = KIND_IZ Then
        y = CLng(Mid$(n, 14, 4)): m = CLng(Mid$(n, 18, 2))
    Else
        y = CLng(Mid$(n, 12, 4)): m = CLng(Mid$(n, 16, 2))
    End If
    If y < 2020 Or m < 1 Or m > 12 Then
        DateOk = False
    ElseIf kind = KIND_IZ Then
        DateOk = True
    Else
        d = CLng(Mid$(n, 18, 2))
        DateOk = (d >= 1 And d <= 31)
    End If
End Function

' Drop separators, uppercase ASCII letters only (same as Python _COMPACT_TABLE).
Public Function CompactId(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, code As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not IsSepChar(ch) Then
            code = AscW(ch)
            If code >= 97 And code <= 122 Then ch = ChrW(code - 32)
            out = out & ch
        End If
    Next i
    CompactId = out
End Function

Private Function IsSepChar(ByVal ch As String) As Boolean
    IsSepChar = (ch = "-" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = ChrW(160))
End Function

' ---------- detection ----------

' Returns a Collection of Array(rawText, startPos, endPos) for every IZ/KSeF-shaped
' token in t (1-based start, exclusive end). Same rules as iz.DETECT_RE.
Public Function DetectIds(ByVal t As String) As Collection
    Dim res As New Collection, n As Long, p As Long, e As Long
    n = Len(t)
    p = 1
    Do While p <= n
        e = 0
        If p = 1 Then
            e = TryMatch(t, p, n)
        ElseIf Not (Mid$(t, p - 1, 1) Like "#") Then
            e = TryMatch(t, p, n)
        End If
        If e > 0 Then
            res.Add Array(Mid$(t, p, e - p), p, e)
            p = e
        Else
            p = p + 1
        End If
    Loop
    Set DetectIds = res
End Function

' End position (exclusive) of an ID starting at p, or 0.
Private Function TryMatch(ByVal t As String, ByVal p As Long, ByVal n As Long) As Long
    Dim e As Long
    e = MatchTemplate(t, p, n, Array("D10", "IZD6", "X12", "X2"))
    If e = 0 Then e = MatchTemplate(t, p, n, Array("D10", "D8", "X6", "X6", "X2"))
    If e > 0 Then
        If Mid$(t, e, 1) Like "[0-9A-Fa-f]" Then e = 0
    End If
    TryMatch = e
End Function

Private Function MatchTemplate(ByVal t As String, ByVal p As Long, ByVal n As Long, _
                               ByVal segs As Variant) As Long
    Dim i As Long, q As Long
    q = p
    For i = LBound(segs) To UBound(segs)
        If i > LBound(segs) Then q = SkipGap(t, q, n)
        q = MatchSegment(t, q, n, CStr(segs(i)))
        If q = 0 Then
            MatchTemplate = 0
            Exit Function
        End If
    Next i
    MatchTemplate = q
End Function

' Segment spec: "D10" = 10 digits, "X12" = 12 hex, "IZD6" = I, Z then 6 digits.
' A line break may sit between any two characters of a segment.
Private Function MatchSegment(ByVal t As String, ByVal q As Long, ByVal n As Long, _
                              ByVal spec As String) As Long
    Dim classes As String, cnt As Long, k As Long, ch As String, cls As String
    If Left$(spec, 2) = "IZ" Then
        classes = "IZ" & String$(CLng(Mid$(spec, 4)), "D")
    Else
        classes = String$(CLng(Mid$(spec, 2)), Left$(spec, 1))
    End If
    cnt = Len(classes)
    For k = 1 To cnt
        If k > 1 Then q = SkipBreak(t, q, n)
        ch = Mid$(t, q, 1)
        cls = Mid$(classes, k, 1)
        If Not CharFits(ch, cls) Then
            MatchSegment = 0
            Exit Function
        End If
        q = q + 1
    Next k
    MatchSegment = q
End Function

Private Function CharFits(ByVal ch As String, ByVal cls As String) As Boolean
    Select Case cls
        Case "D": CharFits = (ch Like "#")
        Case "X": CharFits = (ch Like "[0-9A-Fa-f]")
        Case "I": CharFits = (ch = "I" Or ch = "i")
        Case "Z": CharFits = (ch = "Z" Or ch = "z")
        Case Else: CharFits = False
    End Select
End Function

Private Function IsInlineSep(ByVal ch As String) As Boolean
    IsInlineSep = (ch = "-" Or ch = " " Or ch = vbTab Or ch = ChrW(160))
End Function

Private Function IsIndent(ByVal ch As String) As Boolean
    IsIndent = (ch = " " Or ch = vbTab Or ch = ChrW(160))
End Function

' CRLF / CR / LF followed by indentation; returns q unchanged if there is no break.
Private Function SkipBreak(ByVal t As String, ByVal q As Long, ByVal n As Long) As Long
    Dim ch As String
    ch = Mid$(t, q, 1)
    If ch = vbCr Then
        q = q + 1
        If Mid$(t, q, 1) = vbLf Then q = q + 1
    ElseIf ch = vbLf Then
        q = q + 1
    Else
        SkipBreak = q
        Exit Function
    End If
    Do While q <= n
        If Not IsIndent(Mid$(t, q, 1)) Then Exit Do
        q = q + 1
    Loop
    SkipBreak = q
End Function

' Up to 3 inline separators, optionally a line break and up to 3 more.
Private Function SkipGap(ByVal t As String, ByVal q As Long, ByVal n As Long) As Long
    Dim k As Long, q2 As Long
    For k = 1 To 3
        If Not IsInlineSep(Mid$(t, q, 1)) Then Exit For
        q = q + 1
    Next k
    q2 = SkipBreak(t, q, n)
    If q2 <> q Then
        q = q2
        For k = 1 To 3
            If Not IsInlineSep(Mid$(t, q, 1)) Then Exit For
            q = q + 1
        Next k
    End If
    SkipGap = q
End Function
