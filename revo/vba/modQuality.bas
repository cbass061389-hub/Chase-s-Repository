Attribute VB_Name = "modQuality"
Option Explicit

'======================================================================
' modQuality  -  REVO quality capture, analysis and RCA
'
' Depends on: modConfig
'
' Public entry points:
'   QualityBuild            - run once. Creates the four sheets.
'   QualityRefresh          - rebuilds Quality_Analysis.
'   QualityLogDefect(...)   - called BY THE RELEASE FORMS. This is the
'                             one and only write path for defect data.
'   QualityDefectCodes()    - returns the active code list for combos.
'
' Design note: defects are captured at the moment of release, in the
' form the operator is already using. There is no separate after-the-fact
' entry step, because that is what produced 543 log rows with an empty
' reason column.
'
' Changes from the first draft:
'   - Worksheets.Add(After:=GetSheet(...)) raised 1004 whenever the
'     anchor sheet did not exist yet. Now goes through AddSheetAfter.
'   - A blank "Active" cell on Defect_Master silently hid the code from
'     the operator's combo box, because Empty <> False evaluates False.
'     Blank now means active; only an explicit FALSE deactivates.
'   - BuildAnalysisShell wiped the From/To dates on every rebuild.
'   - Defect_Master gets TRUE/FALSE validation on Active and a Cause
'     Category list, so the master data cannot drift into free text.
'======================================================================

Private Const PARETO_WATCH As Long = 3
Private Const CAUSE_CATS As String = "Method,Material,Setup,Operator,Supplier,Spec,Machine"

'======================================================================
' PUBLIC API - called by the release forms
'======================================================================

' Writes one defect row. Returns True on success.
' Called once per disposition bucket (Reject / Rework / B-Grade).
Public Function QualityLogDefect(ByVal dDate As Date, _
                                 ByVal sShift As String, _
                                 ByVal sCart As String, _
                                 ByVal sWO As String, _
                                 ByVal sSKU As String, _
                                 ByVal lQty As Long, _
                                 ByVal sDisposition As String, _
                                 ByVal sOpDetected As String, _
                                 ByVal sOpCaused As String, _
                                 ByVal sDefectCode As String, _
                                 ByVal sNotes As String, _
                                 ByVal sEnteredBy As String) As Boolean

    Dim ws As Worksheet
    Dim r As Long

    On Error GoTo Fail
    If lQty <= 0 Then Exit Function

    Set ws = GetSheet(SH_QLOG)
    If ws Is Nothing Then Exit Function          ' quality layer not built yet - never block a release

    r = LastRow(ws) + 1
    If r < 2 Then r = 2

    ws.Cells(r, 1).Value = dDate
    ws.Cells(r, 2).Value = sShift
    ws.Cells(r, 3).Value = sCart
    ws.Cells(r, 4).Value = sWO
    ws.Cells(r, 5).Value = sSKU
    ws.Cells(r, 6).Value = lQty
    ws.Cells(r, 7).Value = sDisposition
    ws.Cells(r, 8).Value = sOpDetected
    ws.Cells(r, 9).Value = sOpCaused
    ws.Cells(r, 10).Value = sDefectCode
    ws.Cells(r, 11).Value = LookupCategory(sDefectCode)
    ws.Cells(r, 12).Value = EscapeFlag(sOpDetected, sOpCaused)
    ws.Cells(r, 13).Value = sNotes
    ws.Cells(r, 14).Value = sEnteredBy
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd"

    QualityLogDefect = True
    Exit Function
Fail:
    LogLine "*** QualityLogDefect failed for cart " & sCart & ": " & Err.Description
    QualityLogDefect = False
End Function

' Active defect codes, formatted "CODE - Description" for a combo box.
' A blank Active cell counts as active; only an explicit FALSE hides it.
Public Function QualityDefectCodes() As Variant
    Dim ws As Worksheet, src As Variant
    Dim out() As String
    Dim i As Long, n As Long, lr As Long

    Set ws = GetSheet(SH_DEFECTS)
    If ws Is Nothing Then Exit Function
    lr = LastRow(ws)
    If lr < 2 Then Exit Function

    src = ws.Range("A2:E" & lr).Value
    ReDim out(1 To UBound(src, 1))
    For i = 1 To UBound(src, 1)
        If Len(Trim$(CStr(src(i, 1) & ""))) > 0 Then
            If IsActive(src(i, 5)) Then
                n = n + 1
                out(n) = CStr(src(i, 1)) & " - " & CStr(src(i, 2))
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve out(1 To n)
    QualityDefectCodes = out
End Function

Private Function IsActive(ByVal v As Variant) As Boolean
    Dim s As String
    If IsEmpty(v) Then IsActive = True: Exit Function
    s = UCase$(Trim$(CStr(v & "")))
    If Len(s) = 0 Then IsActive = True: Exit Function
    IsActive = Not (s = "FALSE" Or s = "0" Or s = "N" Or s = "NO")
End Function

' Strips "CODE - Description" back to "CODE".
Public Function QualityBareCode(ByVal sDisplay As String) As String
    Dim p As Long
    p = InStr(1, sDisplay, " - ")
    If p > 0 Then QualityBareCode = Left$(sDisplay, p - 1) Else QualityBareCode = sDisplay
End Function

'======================================================================
' BUILD
'======================================================================

Public Sub QualityBuild()
    On Error GoTo Fail
    FastOn

    BuildDefectMaster
    BuildLog
    BuildRCA
    BuildAnalysisShell

    FastOff
    LogLine "QualityBuild: " & SH_DEFECTS & ", " & SH_QLOG & ", " & SH_RCA & ", " & SH_QANALYSIS & " ready"
    Say "Quality layer built (v" & APP_VERSION & ")." & vbCrLf & vbCrLf & _
        "Next: review " & SH_DEFECTS & " and finalize the code list.", "REVO Quality"
    Exit Sub
Fail:
    FastReset
    LogLine "*** QualityBuild failed: " & Err.Description
    Say "Build failed: " & Err.Description, "REVO Quality", vbCritical
End Sub

Private Sub BuildDefectMaster()
    Dim ws As Worksheet, seed As Variant

    Set ws = AddSheetAfter(SH_DEFECTS, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count).Name)

    If Len(ws.Range("A1").Value) = 0 Then
        ws.Range("A1:E1").Value = Array("Defect Code", "Description", "Cause Category", "Typical Op Caused", "Active")
        StyleHeader ws.Range("A1:E1")
    End If

    ' Seeded from the existing Holds reasons so the floor's own vocabulary carries over.
    If Len(ws.Range("A2").Value) = 0 Then
        seed = Array( _
            Array("VP-GAP", "Vacuum plug gap at joint", "Method", "100", True), _
            Array("BP-PULL", "Bore plugs pulling out", "Material", "30", True), _
            Array("BP-FIT", "Bore plug fit / seating", "Setup", "30", True), _
            Array("TIP-GAP", "Tip gap", "Method", "130", True), _
            Array("TIP-ADH", "Tip adhesion failure", "Material", "130", True), _
            Array("JNT-FIT", "Joint pin fit out of spec", "Setup", "50", True), _
            Array("JNT-THD", "Joint thread damage", "Operator", "50", True), _
            Array("DIA-OOT", "Diameter out of tolerance", "Setup", "60", True), _
            Array("TPR-OOT", "Taper out of tolerance", "Setup", "60", True), _
            Array("FIN-SCR", "Finish scratch / handling damage", "Operator", "120", True), _
            Array("FIN-BLEM", "Finish blemish", "Method", "120", True), _
            Array("TUBE-DEL", "Carbon tube delamination", "Supplier", "IC", True), _
            Array("TUBE-COS", "Carbon tube cosmetic defect", "Supplier", "IC", True), _
            Array("CURE-SHT", "Short cure / premature handling", "Method", "40", True), _
            Array("SPEC-AMB", "Spec unclear or conflicting", "Spec", "", True), _
            Array("OTHER", "Not covered by an existing code", "Method", "", True))
        WriteJagged ws.Range("A2"), seed
    End If

    ' Master data stays master data: no free text in Cause Category or Active.
    AddListValidation ws.Range("C2:C" & VALIDATION_ROWS), CAUSE_CATS
    AddListValidation ws.Range("D2:D" & VALIDATION_ROWS), OPS_CSV
    AddListValidation ws.Range("E2:E" & VALIDATION_ROWS), "TRUE,FALSE"

    ws.Columns("A:E").AutoFit
End Sub

Private Sub BuildLog()
    Dim ws As Worksheet

    Set ws = AddSheetAfter(SH_QLOG, SH_DEFECTS)

    If Len(ws.Range("A1").Value) = 0 Then
        ws.Range("A1:N1").Value = Array("Date", "Shift", "Cart", "WO#", "SKU", "Qty", "Disposition", _
                                        "Op Detected", "Op Caused", "Defect Code", "Cause Category", _
                                        "Escape?", "Notes", "Entered By")
        StyleHeader ws.Range("A1:N1")
        ws.Columns("A").NumberFormat = "yyyy-mm-dd"
        If ws.AutoFilterMode = False Then ws.Rows(1).AutoFilter
    End If
    ws.Columns("A:N").AutoFit
End Sub

Private Sub BuildRCA()
    Dim ws As Worksheet

    Set ws = AddSheetAfter(SH_RCA, SH_QLOG)

    If Len(ws.Range("A1").Value) = 0 Then
        ws.Range("A1:R1").Value = Array("RCA ID", "Opened", "Defect Code", "Trigger Qty", "Owner", _
                                        "Why 1", "Why 2", "Why 3", "Why 4", "Why 5", _
                                        "Root Cause", "Cause Category", "Containment", "Corrective Action", _
                                        "Verify By", "Status", "Closed", "Effective?")
        StyleHeader ws.Range("A1:R1")
        AddListValidation ws.Range("P2:P" & VALIDATION_ROWS), _
            "Open,Containment In Place,Corrective Action In Place,Verifying,Closed"
        AddListValidation ws.Range("L2:L" & VALIDATION_ROWS), CAUSE_CATS
        ws.Range("B:B,O:O,Q:Q").NumberFormat = "yyyy-mm-dd"
        If ws.AutoFilterMode = False Then ws.Rows(1).AutoFilter
    End If
    ws.Columns("A:R").AutoFit
End Sub

Private Sub BuildAnalysisShell()
    Dim ws As Worksheet
    Dim keepFrom As Variant, keepTo As Variant

    Set ws = AddSheetAfter(SH_QANALYSIS, SH_RCA)

    ' Preserve whatever window the user last looked at.
    keepFrom = ws.Range("C3").Value
    keepTo = ws.Range("E3").Value

    ws.Cells.Clear
    ws.Range("B1").Value = "REVO QUALITY ANALYSIS"
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Bold = True
    ws.Range("B3").Value = "From"
    ws.Range("D3").Value = "To"
    ws.Range("G3").Value = "Last refresh"
    ws.Range("B3:G3").Font.Bold = True
    ws.Range("C3,E3").Interior.Color = RGB(255, 242, 204)
    ws.Range("C3").NumberFormat = "yyyy-mm-dd"
    ws.Range("E3").NumberFormat = "yyyy-mm-dd"
    ws.Range("H3").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Range("C3").Value = IIf(IsDate(keepFrom), keepFrom, Date - 90)
    ws.Range("E3").Value = IIf(IsDate(keepTo), keepTo, Date)
End Sub

'======================================================================
' ANALYSIS
'======================================================================

Public Sub QualityRefresh()
    Dim ws As Worksheet
    Dim data As Variant
    Dim r As Long

    On Error GoTo Fail
    FastOn

    Set ws = RequireSheet(SH_QANALYSIS)
    data = LoadLog(NzDate(ws.Range("C3").Value, DateSerial(1900, 1, 1)), _
                   NzDate(ws.Range("E3").Value, DateSerial(2999, 1, 1)))

    ws.Range("B6:S500").Clear

    r = 6
    r = WriteHeadline(ws, r, data)
    r = WritePareto(ws, r, data)
    r = WriteByOperation(ws, r, data)
    r = WriteRCACoverage(ws, r, data)

    ws.Range("H3").Value = Now
    ws.Columns("B:S").AutoFit

    FastOff
    Application.Calculate
    LogLine "QualityRefresh: " & IIf(IsEmpty(data), 0, UBound(data, 1)) & " events in range"
    Exit Sub
Fail:
    FastReset
    LogLine "*** QualityRefresh failed: " & Err.Description
    Say "Refresh failed: " & Err.Description, "REVO Quality", vbCritical
End Sub

Private Function LoadLog(ByVal dFrom As Date, ByVal dTo As Date) As Variant
    Dim ws As Worksheet, src As Variant, out() As Variant
    Dim i As Long, j As Long, n As Long, lr As Long

    Set ws = GetSheet(SH_QLOG)
    If ws Is Nothing Then Exit Function
    lr = LastRow(ws)
    If lr < 2 Then Exit Function

    src = ws.Range("A2:L" & lr).Value
    ReDim out(1 To UBound(src, 1), 1 To 12)
    For i = 1 To UBound(src, 1)
        If IsDate(src(i, 1)) Then
            If CDate(src(i, 1)) >= dFrom And CDate(src(i, 1)) <= dTo Then
                n = n + 1
                For j = 1 To 12
                    out(n, j) = src(i, j)
                Next j
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    LoadLog = TrimRows(out, n, 12)
End Function

Private Function WriteHeadline(ByVal ws As Worksheet, ByVal r As Long, ByVal data As Variant) As Long
    Dim i As Long, qty As Double, esc As Double
    Dim rej As Double, rwk As Double, bgr As Double

    ws.Cells(r, 2).Value = "HEADLINE"
    StyleTitle ws.Cells(r, 2)
    r = r + 1

    If IsEmpty(data) Then
        ws.Cells(r, 2).Value = "No records in range."
        WriteHeadline = r + 2
        Exit Function
    End If

    For i = 1 To UBound(data, 1)
        qty = qty + Nz(data(i, 6))
        If UCase$(Trim$(CStr(data(i, 12) & ""))) = "ESCAPE" Then esc = esc + Nz(data(i, 6))
        Select Case UCase$(Trim$(CStr(data(i, 7) & "")))
            Case "REJECT": rej = rej + Nz(data(i, 6))
            Case "REWORK": rwk = rwk + Nz(data(i, 6))
            Case "B-GRADE": bgr = bgr + Nz(data(i, 6))
        End Select
    Next i

    ws.Cells(r, 2).Resize(1, 6).Value = Array("Events", "Total Qty", "Reject", "Rework", "B-Grade", "Escape %")
    StyleHeader ws.Cells(r, 2).Resize(1, 6)
    r = r + 1
    ws.Cells(r, 2).Resize(1, 6).Value = Array(UBound(data, 1), qty, rej, rwk, bgr, IIf(qty = 0, 0, esc / qty))
    ws.Cells(r, 7).NumberFormat = "0.0%"
    WriteHeadline = r + 2
End Function

Private Function WritePareto(ByVal ws As Worksheet, ByVal r As Long, ByVal data As Variant) As Long
    Dim dic As Object, cat As Object
    Dim keys As Variant, vals() As Double
    Dim i As Long, k As Long, total As Double, run As Double
    Dim code As String

    ws.Cells(r, 2).Value = "PARETO BY DEFECT CODE"
    StyleTitle ws.Cells(r, 2)
    r = r + 1
    If IsEmpty(data) Then WritePareto = r + 1: Exit Function

    Set dic = CreateObject("Scripting.Dictionary")
    Set cat = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        code = Trim$(CStr(data(i, 10) & ""))
        If Len(code) > 0 Then
            If dic.Exists(code) Then
                dic(code) = dic(code) + Nz(data(i, 6))
            Else
                dic.Add code, Nz(data(i, 6))
            End If
            cat(code) = CStr(data(i, 11) & "")
            total = total + Nz(data(i, 6))
        End If
    Next i
    If dic.Count = 0 Then WritePareto = r + 1: Exit Function

    keys = dic.keys
    ReDim vals(0 To dic.Count - 1)
    For i = 0 To dic.Count - 1
        vals(i) = dic(keys(i))
    Next i
    SortDesc keys, vals

    ws.Cells(r, 2).Resize(1, 5).Value = Array("Defect Code", "Cause Category", "Qty", "% of Total", "Cumulative %")
    StyleHeader ws.Cells(r, 2).Resize(1, 5)
    r = r + 1
    For k = 0 To UBound(keys)
        run = run + vals(k)
        ws.Cells(r, 2).Value = keys(k)
        ws.Cells(r, 3).Value = cat(keys(k))
        ws.Cells(r, 4).Value = vals(k)
        ws.Cells(r, 5).Value = IIf(total = 0, 0, vals(k) / total)
        ws.Cells(r, 6).Value = IIf(total = 0, 0, run / total)
        ws.Cells(r, 5).Resize(1, 2).NumberFormat = "0.0%"
        If k < PARETO_WATCH Then ws.Cells(r, 2).Resize(1, 5).Interior.Color = RGB(255, 235, 156)
        r = r + 1
    Next k
    WritePareto = r + 1
End Function

Private Function WriteByOperation(ByVal ws As Worksheet, ByVal r As Long, ByVal data As Variant) As Long
    Dim caused As Object, esc As Object
    Dim ops As Variant
    Dim i As Long, k As Long
    Dim op As String, c As Double, e As Double

    ws.Cells(r, 2).Value = "BY OPERATION - WHERE CAUSED vs WHERE CAUGHT"
    StyleTitle ws.Cells(r, 2)
    r = r + 1
    If IsEmpty(data) Then WriteByOperation = r + 1: Exit Function

    Set caused = CreateObject("Scripting.Dictionary")
    Set esc = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        op = Trim$(CStr(data(i, 9) & ""))
        If Len(op) > 0 Then
            If caused.Exists(op) Then caused(op) = caused(op) + Nz(data(i, 6)) Else caused.Add op, Nz(data(i, 6))
            If UCase$(Trim$(CStr(data(i, 12) & ""))) = "ESCAPE" Then
                If esc.Exists(op) Then esc(op) = esc(op) + Nz(data(i, 6)) Else esc.Add op, Nz(data(i, 6))
            End If
        End If
    Next i

    ws.Cells(r, 2).Resize(1, 4).Value = Array("Op Caused", "Qty Caused", "Qty Escaped", "Escape %")
    StyleHeader ws.Cells(r, 2).Resize(1, 4)
    r = r + 1
    ops = OpsArray()
    For k = LBound(ops) To UBound(ops)
        op = ops(k)
        If caused.Exists(op) Then
            c = caused(op)
            e = IIf(esc.Exists(op), esc(op), 0)
            ws.Cells(r, 2).Value = op
            ws.Cells(r, 3).Value = c
            ws.Cells(r, 4).Value = e
            ws.Cells(r, 5).Value = IIf(c = 0, 0, e / c)
            ws.Cells(r, 5).NumberFormat = "0.0%"
            If c > 0 Then If e / c >= 0.5 Then ws.Cells(r, 5).Interior.Color = RGB(255, 199, 206)
            r = r + 1
        End If
    Next k
    WriteByOperation = r + 1
End Function

' The block that turns reporting into action: any top-N code without an
' open RCA gets flagged. Without this the Pareto is just a picture.
Private Function WriteRCACoverage(ByVal ws As Worksheet, ByVal r As Long, ByVal data As Variant) As Long
    Dim dic As Object, openRCA As Object
    Dim rcaWs As Worksheet, src As Variant
    Dim keys As Variant, vals() As Double
    Dim i As Long, k As Long, lr As Long
    Dim code As String, st As String

    ws.Cells(r, 2).Value = "RCA COVERAGE - TOP " & PARETO_WATCH
    StyleTitle ws.Cells(r, 2)
    r = r + 1
    If IsEmpty(data) Then WriteRCACoverage = r + 1: Exit Function

    Set dic = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        code = Trim$(CStr(data(i, 10) & ""))
        If Len(code) > 0 Then
            If dic.Exists(code) Then dic(code) = dic(code) + Nz(data(i, 6)) Else dic.Add code, Nz(data(i, 6))
        End If
    Next i
    If dic.Count = 0 Then WriteRCACoverage = r + 1: Exit Function

    keys = dic.keys
    ReDim vals(0 To dic.Count - 1)
    For i = 0 To dic.Count - 1
        vals(i) = dic(keys(i))
    Next i
    SortDesc keys, vals

    Set openRCA = CreateObject("Scripting.Dictionary")
    Set rcaWs = GetSheet(SH_RCA)
    If Not rcaWs Is Nothing Then
        lr = LastRow(rcaWs)
        If lr >= 2 Then
            src = rcaWs.Range("C2:P" & lr).Value
            For i = 1 To UBound(src, 1)
                code = Trim$(CStr(src(i, 1) & ""))
                st = UCase$(Trim$(CStr(src(i, 14) & "")))
                If Len(code) > 0 And Len(st) > 0 And st <> "CLOSED" Then openRCA(code) = True
            Next i
        End If
    End If

    ws.Cells(r, 2).Resize(1, 3).Value = Array("Defect Code", "Qty", "RCA Status")
    StyleHeader ws.Cells(r, 2).Resize(1, 3)
    r = r + 1
    For k = 0 To Application.Min(PARETO_WATCH - 1, UBound(keys))
        ws.Cells(r, 2).Value = keys(k)
        ws.Cells(r, 3).Value = vals(k)
        If openRCA.Exists(CStr(keys(k))) Then
            ws.Cells(r, 4).Value = "Open RCA"
        Else
            ws.Cells(r, 4).Value = "*** NO OPEN RCA - ASSIGN AN OWNER ***"
            ws.Cells(r, 2).Resize(1, 3).Interior.Color = RGB(255, 199, 206)
        End If
        r = r + 1
    Next k
    WriteRCACoverage = r + 1
End Function

'======================================================================
' HELPERS
'======================================================================

Private Function LookupCategory(ByVal code As String) As String
    Dim ws As Worksheet, src As Variant, i As Long, lr As Long
    Set ws = GetSheet(SH_DEFECTS)
    If ws Is Nothing Then Exit Function
    lr = LastRow(ws)
    If lr < 2 Then Exit Function
    src = ws.Range("A2:C" & lr).Value
    For i = 1 To UBound(src, 1)
        If StrComp(Trim$(CStr(src(i, 1) & "")), code, vbTextCompare) = 0 Then
            LookupCategory = CStr(src(i, 3) & "")
            Exit Function
        End If
    Next i
    LookupCategory = "UNMAPPED"
End Function

Private Function EscapeFlag(ByVal opDet As String, ByVal opCause As String) As String
    If Len(opDet) = 0 Or Len(opCause) = 0 Then Exit Function
    If StrComp(opDet, opCause, vbTextCompare) <> 0 Then
        EscapeFlag = "ESCAPE"
    Else
        EscapeFlag = "CAUGHT AT SOURCE"
    End If
End Function

Private Sub AddListValidation(ByVal rng As Range, ByVal csv As String)
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=csv
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
    On Error GoTo 0
End Sub

Private Sub StyleHeader(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
End Sub

Private Sub StyleTitle(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 84, 106)
    End With
End Sub

Private Sub WriteJagged(ByVal topLeft As Range, ByVal rows As Variant)
    Dim i As Long, j As Long, cols As Long
    Dim out() As Variant
    cols = UBound(rows(LBound(rows))) - LBound(rows(LBound(rows))) + 1
    ReDim out(1 To UBound(rows) - LBound(rows) + 1, 1 To cols)
    For i = LBound(rows) To UBound(rows)
        For j = 0 To cols - 1
            out(i - LBound(rows) + 1, j + 1) = rows(i)(j)
        Next j
    Next i
    topLeft.Resize(UBound(out, 1), cols).Value = out
End Sub

Private Function TrimRows(ByVal src As Variant, ByVal n As Long, ByVal cols As Long) As Variant
    Dim out() As Variant, i As Long, j As Long
    ReDim out(1 To n, 1 To cols)
    For i = 1 To n
        For j = 1 To cols
            out(i, j) = src(i, j)
        Next j
    Next i
    TrimRows = out
End Function

' Insertion sort, descending. Code counts are in the tens - nothing faster is needed.
Private Sub SortDesc(ByRef keys As Variant, ByRef vals() As Double)
    Dim i As Long, j As Long, tv As Double, tk As Variant
    For i = LBound(vals) + 1 To UBound(vals)
        tv = vals(i): tk = keys(i): j = i - 1
        Do While j >= LBound(vals)
            If vals(j) >= tv Then Exit Do
            vals(j + 1) = vals(j): keys(j + 1) = keys(j): j = j - 1
        Loop
        vals(j + 1) = tv: keys(j + 1) = tk
    Next i
End Sub
