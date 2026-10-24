Attribute VB_Name = "modHelper"
' Runs the ksefiz helper for the API buttons (file protocol: src/ksefiz/excel_job.py).
' A job folder gets in.xlsx, the helper runs in the background, progress.txt is polled for
' the status bar, Esc creates "cancel", and out.xlsx is read back as values.
' Windows: WScript.Shell. Mac: AppleScriptTask + ksefiz.scpt (scripts/install_mac_helper.sh).
Option Explicit

Private Const POLL_SECONDS As Double = 0.4
Private Const STALL_SECONDS As Double = 900

' Config sheet value by setting name in column A ("" if absent).
Public Function CfgValue(ByVal name As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Config")
    For r = 2 To 40
        If LCase$(Trim$(CellText(ws.Cells(r, 1).Value))) = LCase$(name) Then
            CfgValue = Trim$(CellText(ws.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
End Function

Public Sub SetCfgValue(ByVal name As String, ByVal v As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Config")
    For r = 2 To 40
        If LCase$(Trim$(CellText(ws.Cells(r, 1).Value))) = LCase$(name) Then
            ws.Cells(r, 2).Value = v
            Exit Sub
        End If
    Next r
End Sub

' Runs one helper job. header/data may be Empty (query). extraArgs: Array of tokens.
' On return outValues = 2D array of result rows (no header) or Empty; res = keyed Collection
' with status, message, error_code, trace_id, requests, input_hash, budget_minute, budget_hour.
Public Function RunHelperJob(ByVal op As String, ByVal header As Variant, ByVal data As Variant, _
                             ByVal extraArgs As Variant, ByRef outValues As Variant, _
                             ByRef res As Collection) As Boolean
    Dim job As String, sep As String, helper As String, env As String, args As Collection
    Dim a As Variant, done As Boolean, lastText As String, lastChange As Double, t As String

    Set res = New Collection
    outValues = Empty
    sep = Application.PathSeparator
    helper = CfgValue("helper_path")
    env = UCase$(CfgValue("environment"))
    If env = "" Then env = "TEST"
    If helper = "" Or helper = "ksefiz" Then
        MsgBox "Set Config > helper_path to the full path of the ksefiz executable.", vbExclamation
        Exit Function
    End If
    If env = "PROD" Then
        If MsgBox("Environment is PROD: real data with legal effect." & vbLf & vbLf & _
                  "Run '" & op & "' against PROD?", vbYesNo + vbExclamation + vbDefaultButton2) <> vbYes Then
            Exit Function
        End If
    End If

    Set args = New Collection
    args.Add "--op": args.Add op
    args.Add "--env": args.Add LCase$(env)
    If CfgValue("profile") <> "" Then args.Add "--profile": args.Add CfgValue("profile")
    If CfgValue("log_folder") <> "" And IsAbsolutePath(CfgValue("log_folder")) Then
        args.Add "--log-dir": args.Add CfgValue("log_folder")
    End If
    If env = "PROD" Then args.Add "--confirm-prod"
    If Not IsEmpty(extraArgs) Then
        For Each a In extraArgs
            args.Add CStr(a)
        Next a
    End If
    For Each a In args
        If Not SafeToken(CStr(a)) Then
            MsgBox "Unsupported character in setting/argument: " & CStr(a), vbExclamation
            Exit Function
        End If
    Next a

    job = NewJobFolder(op)
    If Not IsEmpty(data) Then WriteInput job & sep & "in.xlsx", header, data

    If Not StartHelper(helper, job, args) Then
        RemoveJob job
        Exit Function
    End If

    ' Poll until the launcher writes "exit"; Esc requests cancellation.
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Interrupted
    lastChange = Timer
    Do
        If FileExists(job & sep & "exit") Then done = True
        t = ReadText(job & sep & "progress.txt")
        If t <> lastText Then
            lastText = t
            lastChange = Timer
            Application.StatusBar = "KSeF " & op & ": " & ProgressLabel(t) & "   (Esc = cancel)"
        ElseIf ElapsedSince(lastChange) > STALL_SECONDS Then
            Application.EnableCancelKey = xlInterrupt
            Application.StatusBar = False
            MsgBox "ksefiz did not report progress for " & STALL_SECONDS & " s; giving up." & vbLf & _
                   "Job folder kept for diagnosis: " & job, vbExclamation
            Exit Function
        End If
        If Not done Then PauseSeconds POLL_SECONDS
    Loop Until done
    On Error GoTo 0
    Application.EnableCancelKey = xlInterrupt
    Application.StatusBar = False

    If FileExists(job & sep & "out.xlsx") Then
        ReadOutput job & sep & "out.xlsx", outValues, res
        RunHelperJob = (ResultValue(res, "status") <> "error")
    Else
        res.Add "error", "status"
        res.Add "ksefiz failed: " & Left$(ReadText(job & sep & "log.txt"), 600), "message"
    End If
    RemoveJob job
    Exit Function

Interrupted:
    If Err.Number = 18 Then  ' Esc / Cmd+.
        WriteText job & sep & "cancel", "cancel"
        Application.StatusBar = "KSeF " & op & ": cancelling after the current request..."
        Resume Next
    End If
    Application.EnableCancelKey = xlInterrupt
    Application.StatusBar = False
    MsgBox "Error while waiting for ksefiz: " & Err.Description, vbExclamation
End Function

Public Function ResultValue(ByVal res As Collection, ByVal key As String) As String
    On Error Resume Next
    ResultValue = CStr(res(key))
    If Err.Number <> 0 Then ResultValue = ""
    Err.Clear
End Function

' ---------- launching ----------

Private Function StartHelper(ByVal helper As String, ByVal job As String, ByVal args As Collection) As Boolean
    Dim a As Variant
#If Mac Then
    Dim param As String, reply As String
    param = helper & vbTab & job
    For Each a In args
        param = param & vbTab & CStr(a)
    Next a
    On Error GoTo NoLauncher
    reply = AppleScriptTask("ksefiz.scpt", "StartJob", param)
    StartHelper = (reply = "started")
    Exit Function
NoLauncher:
    MsgBox "Could not start ksefiz through AppleScript:" & vbLf & Err.Description & vbLf & vbLf & _
           "Run scripts/install_mac_helper.sh once to install ksefiz.scpt.", vbExclamation
#Else
    Dim cmd As String
    cmd = """" & helper & """ excel-run --job """ & job & """"
    For Each a In args
        cmd = cmd & " """ & CStr(a) & """"   ' SafeToken: no quotes inside
    Next a
    cmd = "cmd /v:on /c """ & cmd & " > """ & job & "\log.txt"" 2>&1 & echo !errorlevel! > """ & job & "\exit"""""
    On Error GoTo NoShell
    CreateObject("WScript.Shell").Run cmd, 0, False
    StartHelper = True
    Exit Function
NoShell:
    MsgBox "Could not start ksefiz: " & Err.Description, vbExclamation
#End If
End Function

' Letters, digits and - _ . : / \ space only (dates, profile names, paths from Config).
Private Function SafeToken(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not (ch Like "[A-Za-z0-9]" Or InStr("-_.:/\ ", ch) > 0) Then Exit Function
    Next i
    SafeToken = True
End Function

Private Function IsAbsolutePath(ByVal p As String) As Boolean
    IsAbsolutePath = (Left$(p, 1) = "/" Or Mid$(p, 2, 1) = ":" Or Left$(p, 2) = "\\")
End Function

' ---------- job folder and files ----------

Private Function JobsRoot() As String
#If Mac Then
    JobsRoot = Environ("HOME") & "/ksefiz-jobs"   ' Excel's sandbox container
#Else
    JobsRoot = Environ("TEMP") & "\ksefiz-jobs"
#End If
End Function

Private Function NewJobFolder(ByVal op As String) As String
    Dim root As String, p As String
    root = JobsRoot()
    On Error Resume Next
    MkDir root
    On Error GoTo 0
    Randomize
    p = root & Application.PathSeparator & Format$(Now, "yyyymmdd-hhnnss") & "-" & op & "-" & _
        CStr(Int(Rnd * 1000000))
    MkDir p
    NewJobFolder = p
End Function

Private Sub RemoveJob(ByVal job As String)
    Dim f As String, sep As String, names As New Collection, n As Variant
    sep = Application.PathSeparator
    On Error Resume Next
    f = Dir(job & sep & "*")
    Do While f <> ""   ' collect first: deleting while Dir() enumerates can skip files
        names.Add f
        f = Dir()
    Loop
    For Each n In names
        Kill job & sep & CStr(n)
    Next n
    RmDir job
    On Error GoTo 0
End Sub

Private Sub WriteInput(ByVal path As String, ByVal header As Variant, ByVal data As Variant)
    Dim wb As Workbook, ws As Worksheet, nCols As Long, nRows As Long
    nCols = UBound(header) - LBound(header) + 1
    nRows = UBound(data, 1) - LBound(data, 1) + 1
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Resize(1, nCols).Value = header
    ws.Range("A2").Resize(nRows, nCols).NumberFormat = "@"
    ws.Range("A2").Resize(nRows, nCols).Value = data
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=51  ' xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Sub ReadOutput(ByVal path As String, ByRef outValues As Variant, ByRef res As Collection)
    Dim wb As Workbook, ws As Worksheet, last As Long, cols As Long, r As Long
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True)
    Set ws = wb.Worksheets("out")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    cols = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If last >= 2 And CellText(ws.Cells(1, 1).Value) <> "" Then
        outValues = ws.Range(ws.Cells(2, 1), ws.Cells(last, cols)).Value
    End If
    Set ws = wb.Worksheets("_result")
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        On Error Resume Next
        res.Add CellText(ws.Cells(r, 2).Value), CellText(ws.Cells(r, 1).Value)
        On Error GoTo 0
    Next r
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Sub

Public Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(p) <> "")
End Function

Private Function ReadText(ByVal p As String) As String
    Dim f As Integer, s As String
    If Not FileExists(p) Then Exit Function
    On Error GoTo Done
    f = FreeFile
    Open p For Binary Access Read As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    ReadText = Trim$(Replace(Replace(s, vbCr, " "), vbLf, " "))
    Exit Function
Done:
    On Error Resume Next
    Close #f
End Function

Private Sub WriteText(ByVal p As String, ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, s
    Close #f
End Sub

' "done total message" -> "3/10 decode" (message only while starting)
Private Function ProgressLabel(ByVal t As String) As String
    Dim parts() As String
    If t = "" Then
        ProgressLabel = "starting..."
        Exit Function
    End If
    parts = Split(t, " ", 3)
    If UBound(parts) < 2 Then
        ProgressLabel = t
    ElseIf parts(1) = "0" Then
        ProgressLabel = parts(2)
    Else
        ProgressLabel = parts(0) & "/" & parts(1)
    End If
End Function

Private Function ElapsedSince(ByVal t0 As Double) As Double
    ElapsedSince = Timer - t0
    If ElapsedSince < 0 Then ElapsedSince = ElapsedSince + 86400#  ' midnight wrap
End Function

Private Sub PauseSeconds(ByVal s As Double)
    Dim t0 As Double
    t0 = Timer
    Do While ElapsedSince(t0) < s
        DoEvents
    Loop
End Sub
