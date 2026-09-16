Attribute VB_Name = "REVO_QualityDash"
'==============================================================================
' REVO_QualityDash  -  what is actually going wrong, and where
'
' Entry point: REVO_BuildQualityDashboard
'
' Built entirely from Quality Log, so it works the moment the backfill has run
' and gets better every time someone releases a cart.
'
' The point of this sheet is a Pareto, not a scoreboard. Reject percent tells
' you there is a problem; defect type by root cause by op tells you whose
' problem it is and what to change. Everything here is ranked by units so the
' top row is always the one worth the meeting.
'==============================================================================
Option Explicit

Private Const TOP_N As Long = 10

Private Type Period
    name  As String
    d1    As Date
    d2    As Date
End Type

'==============================================================================
Public Sub REVO_BuildQualityDashboard()
    BuildQualityDashboard True
End Sub

Public Sub BuildQualityDashboard(Optional ByVal interactive As Boolean = True)
    Dim ws As Worksheet
    Dim r As Long
    Dim wtd As Period, mtd As Period, ytd As Period, prior As Period
    Dim today As Date

    On Error GoTo Failed
    REVO_Core.PushAppState

    REVO_Quality.EnsureQualityLog
    today = Date

    wtd.name = "Week to date":  wtd.d1 = WeekStartOf(today): wtd.d2 = today
    mtd.name = "Month to date": mtd.d1 = DateSerial(Year(today), Month(today), 1): mtd.d2 = today
    ytd.name = "Year to date":  ytd.d1 = DateSerial(Year(today), 1, 1): ytd.d2 = today
    prior.name = "Prior week":  prior.d1 = wtd.d1 - 7: prior.d2 = wtd.d1 - 1

    Set ws = REVO_Core.GetOrCreateSheet(SH_QUAL_DASH)
    ws.Cells.Clear
    ws.Cells.Font.name = "Segoe UI"
    ws.Cells.Font.Size = 10

    ws.Range("A1").Value = "REVO QUALITY DASHBOARD"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.bold = True
    ws.Range("A2").Value = "Rebuilt " & Format$(Now, "ddd d mmm yyyy hh:nn") & _
                           " from " & SH_QUAL_LOG & ". Derived sheet - do not type here."
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = RGB(90, 90, 90)

    r = 4

    '======================================================== headline ======
    Section ws, r, "HEADLINE": r = r + 1
    r = WriteRateTable(ws, r, wtd, mtd, ytd, prior)
    r = r + 1

    '======================================================== paretos =======
    Section ws, r, "DEFECT PARETO  -  month to date": r = r + 1
    r = WritePareto(ws, r, mtd, QLC_DEFECT, "Defect type")
    r = r + 1

    Section ws, r, "ROOT CAUSE PARETO  -  month to date": r = r + 1
    r = WritePareto(ws, r, mtd, QLC_ROOT, "Root cause")
    r = r + 1

    Section ws, r, "BY OPERATION  -  month to date": r = r + 1
    r = WritePareto(ws, r, mtd, QLC_OP, "Detected at op")
    r = r + 1

    Section ws, r, "BY DIAMETER AND JOINT  -  month to date": r = r + 1
    r = WriteFamilyTable(ws, r, mtd)
    r = r + 1

    '======================================================== movers ========
    Section ws, r, "MOVERS  -  this week against last": r = r + 1
    r = WriteMovers(ws, r, wtd, prior)
    r = r + 1

    '======================================================== hygiene =======
    Section ws, r, "DATA HYGIENE": r = r + 1
    r = WriteHygiene(ws, r, ytd)

    ws.Columns("A:A").ColumnWidth = 32
    ws.Columns("B:F").ColumnWidth = 15
    ws.Columns("G:G").ColumnWidth = 58
    ws.Rows(1).RowHeight = 26

    REVO_Core.PopAppState
    REVO_Core.Audit "REVO_QualityDash", "BUILD", SH_QUAL_DASH, "Rebuilt"

    If interactive Then
        ws.Activate
        ws.Range("A1").Select
    End If
    Exit Sub

Failed:
    Dim eNum As Long, eDesc As String
    eNum = Err.Number
    eDesc = Err.Description
    REVO_Core.ResetAppState
    REVO_Core.Audit "REVO_QualityDash", "ERROR", SH_QUAL_DASH, eNum & ": " & eDesc
    If interactive Then
        MsgBox "Quality dashboard build failed." & vbCrLf & vbCrLf & _
               "Error " & eNum & ": " & eDesc, vbExclamation, "Quality Dashboard"
    End If
    Resume CleanExit
CleanExit:
End Sub

'==============================================================================
Private Function WriteRateTable(ByVal ws As Worksheet, ByVal r As Long, _
                                ByRef wtd As Period, ByRef mtd As Period, _
                                ByRef ytd As Period, ByRef prior As Period) As Long
    Dim i As Long, c As Long
    Dim p As Period
    Dim rwk As Double, rej As Double, bg As Double, rel As Double, insp As Double

    ws.Cells(r, 1).Value = "Measure"
    ws.Cells(r, 2).Value = wtd.name
    ws.Cells(r, 3).Value = prior.name
    ws.Cells(r, 4).Value = mtd.name
    ws.Cells(r, 5).Value = ytd.name
    HeaderBand ws, r, 7
    r = r + 1

    Dim labels As Variant
    labels = Array("Released to inventory", "Rework", "Reject", "B grade", _
                   "Inspected", "Reject rate", "Rework rate", "First-pass yield")

    For i = LBound(labels) To UBound(labels)
        ws.Cells(r + i - LBound(labels), 1).Value = labels(i)
    Next i

    For c = 2 To 5
        Select Case c
            Case 2: p = wtd
            Case 3: p = prior
            Case 4: p = mtd
            Case 5: p = ytd
        End Select

        SumPeriod p.d1, p.d2, rwk, rej, bg
        rel = REVO_Core.ReleasedBetween(p.d1, p.d2)
        insp = rel + rwk + rej + bg

        ws.Cells(r, c).Value = rel:      ws.Cells(r, c).NumberFormat = "#,##0"
        ws.Cells(r + 1, c).Value = rwk:  ws.Cells(r + 1, c).NumberFormat = "#,##0"
        ws.Cells(r + 2, c).Value = rej:  ws.Cells(r + 2, c).NumberFormat = "#,##0"
        ws.Cells(r + 3, c).Value = bg:   ws.Cells(r + 3, c).NumberFormat = "#,##0"
        ws.Cells(r + 4, c).Value = insp: ws.Cells(r + 4, c).NumberFormat = "#,##0"

        If insp > 0 Then
            ws.Cells(r + 5, c).Value = rej / insp
            ws.Cells(r + 6, c).Value = rwk / insp
            ws.Cells(r + 7, c).Value = rel / insp
            ws.Range(ws.Cells(r + 5, c), ws.Cells(r + 7, c)).NumberFormat = "0.0%"
        End If
    Next c

    ws.Cells(r + 4, 7).Value = "Released plus every disposition - the denominator for the rates below."
    ws.Cells(r + 7, 7).Value = "Share released straight to inventory with no disposition at all."
    ws.Range(ws.Cells(r + 4, 7), ws.Cells(r + 7, 7)).Font.Color = RGB(120, 120, 120)
    ws.Range(ws.Cells(r + 5, 1), ws.Cells(r + 7, 5)).Font.bold = True

    WriteRateTable = r + 8
End Function

'==============================================================================
Private Function WritePareto(ByVal ws As Worksheet, ByVal r As Long, ByRef p As Period, _
                             ByVal col As Long, ByVal caption As String) As Long
    Dim d As Object
    Dim keys() As String, vals() As Double, n As Long
    Dim i As Long, total As Double, running As Double
    Dim prevCum As Double
    Dim lineDrawn As Boolean

    Set d = Tally(p.d1, p.d2, col)
    n = DictToArrays(d, keys, vals)

    ws.Cells(r, 1).Value = caption
    ws.Cells(r, 2).Value = "Units"
    ws.Cells(r, 3).Value = "Share"
    ws.Cells(r, 4).Value = "Cumulative"
    ws.Cells(r, 5).Value = "Carts"
    HeaderBand ws, r, 7
    r = r + 1

    If n = 0 Then
        ws.Cells(r, 1).Value = "No events in this period."
        ws.Cells(r, 1).Font.Italic = True
        WritePareto = r + 1
        Exit Function
    End If

    For i = 1 To n
        total = total + vals(i)
    Next i

    For i = 1 To n
        If i > TOP_N Then Exit For
        running = running + vals(i)
        ws.Cells(r, 1).Value = keys(i)
        ws.Cells(r, 2).Value = vals(i)
        ws.Cells(r, 2).NumberFormat = "#,##0"
        ws.Cells(r, 3).Value = vals(i) / total
        ws.Cells(r, 3).NumberFormat = "0.0%"
        ws.Cells(r, 4).Value = running / total
        ws.Cells(r, 4).NumberFormat = "0.0%"
        ws.Cells(r, 5).Value = vals(i) / DEFAULT_SHAFTS_PER_CART
        ws.Cells(r, 5).NumberFormat = "0.0"

        ' The 80% line is where the meeting should stop.
        If running / total <= 0.8 Then
            ws.Cells(r, 1).Font.bold = True
        ElseIf Not lineDrawn And prevCum < 0.8 Then
            ws.Cells(r, 7).Value = "Everything above this line is 80% of the problem."
            ws.Cells(r, 7).Font.Color = RGB(150, 75, 0)
            lineDrawn = True
        End If
        prevCum = running / total
        r = r + 1
    Next i

    If n > TOP_N Then
        ws.Cells(r, 1).Value = "... and " & (n - TOP_N) & " more"
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    End If

    WritePareto = r
End Function

'==============================================================================
Private Function WriteFamilyTable(ByVal ws As Worksheet, ByVal r As Long, ByRef p As Period) As Long
    Dim wsQ As Worksheet, i As Long, lastR As Long
    Dim d As Object, rel As Object
    Dim dt As Date, key As String, q As Double
    Dim keys() As String, vals() As Double, n As Long

    Set wsQ = REVO_Core.GetSheet(SH_QUAL_LOG)
    If wsQ Is Nothing Then WriteFamilyTable = r: Exit Function

    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare

    lastR = wsQ.Cells(wsQ.Rows.Count, QLC_EVENTID).End(xlUp).row
    For i = QL_HEADER_ROW + 1 To lastR
        dt = REVO_Core.SafeDate(wsQ.Cells(i, QLC_DATE).Value, 0)
        If dt >= p.d1 And dt <= p.d2 Then
            key = REVO_Core.SafeS(wsQ.Cells(i, QLC_DIA).Value) & "  " & _
                  REVO_Core.SafeS(wsQ.Cells(i, QLC_JOINT).Value)
            q = REVO_Core.SafeD(wsQ.Cells(i, QLC_QTY).Value)
            If d.exists(key) Then d(key) = d(key) + q Else d.Add key, q
        End If
    Next i

    n = DictToArrays(d, keys, vals)

    ws.Cells(r, 1).Value = "Diameter / joint"
    ws.Cells(r, 2).Value = "Units dispositioned"
    ws.Cells(r, 3).Value = "Share"
    HeaderBand ws, r, 7
    r = r + 1

    If n = 0 Then
        ws.Cells(r, 1).Value = "No events in this period."
        ws.Cells(r, 1).Font.Italic = True
        WriteFamilyTable = r + 1
        Exit Function
    End If

    Dim total As Double
    For i = 1 To n
        total = total + vals(i)
    Next i

    For i = 1 To n
        If i > TOP_N Then Exit For
        ws.Cells(r, 1).Value = keys(i)
        ws.Cells(r, 2).Value = vals(i)
        ws.Cells(r, 2).NumberFormat = "#,##0"
        ws.Cells(r, 3).Value = vals(i) / total
        ws.Cells(r, 3).NumberFormat = "0.0%"
        r = r + 1
    Next i

    WriteFamilyTable = r
End Function

'==============================================================================
' MOVERS  -  what changed, which is the only part worth reading weekly
'==============================================================================
Private Function WriteMovers(ByVal ws As Worksheet, ByVal r As Long, _
                             ByRef wtd As Period, ByRef prior As Period) As Long
    Dim dNow As Object, dWas As Object
    Dim k As Variant
    Dim keys() As String, deltas() As Double, n As Long
    Dim i As Long, j As Long
    Dim tk As String, td As Double

    Set dNow = Tally(wtd.d1, wtd.d2, QLC_DEFECT)
    Set dWas = Tally(prior.d1, prior.d2, QLC_DEFECT)

    ReDim keys(1 To dNow.Count + dWas.Count + 1)
    ReDim deltas(1 To dNow.Count + dWas.Count + 1)

    For Each k In dNow.keys
        n = n + 1
        keys(n) = CStr(k)
        deltas(n) = dNow(k) - IIf(dWas.exists(k), dWas(k), 0)
    Next k
    For Each k In dWas.keys
        If Not dNow.exists(k) Then
            n = n + 1
            keys(n) = CStr(k)
            deltas(n) = -dWas(k)
        End If
    Next k

    ws.Cells(r, 1).Value = "Defect type"
    ws.Cells(r, 2).Value = "This week"
    ws.Cells(r, 3).Value = "Last week"
    ws.Cells(r, 4).Value = "Change"
    ws.Cells(r, 5).Value = "Read"
    HeaderBand ws, r, 7
    r = r + 1

    If n = 0 Then
        ws.Cells(r, 1).Value = "No events in either week."
        ws.Cells(r, 1).Font.Italic = True
        WriteMovers = r + 1
        Exit Function
    End If

    ' Biggest absolute move first - a big drop matters as much as a big rise.
    For i = 1 To n - 1
        For j = i + 1 To n
            If Abs(deltas(j)) > Abs(deltas(i)) Then
                td = deltas(i): deltas(i) = deltas(j): deltas(j) = td
                tk = keys(i): keys(i) = keys(j): keys(j) = tk
            End If
        Next j
    Next i

    For i = 1 To n
        If i > TOP_N Then Exit For
        ws.Cells(r, 1).Value = keys(i)
        ws.Cells(r, 2).Value = IIf(dNow.exists(keys(i)), dNow(keys(i)), 0)
        ws.Cells(r, 3).Value = IIf(dWas.exists(keys(i)), dWas(keys(i)), 0)
        ws.Cells(r, 4).Value = deltas(i)
        ws.Cells(r, 4).NumberFormat = "+#,##0;-#,##0;0"
        If deltas(i) > 0 Then
            ws.Cells(r, 4).Font.Color = RGB(192, 0, 0)
            ws.Cells(r, 5).Value = "Worse"
        ElseIf deltas(i) < 0 Then
            ws.Cells(r, 4).Font.Color = RGB(0, 97, 0)
            ws.Cells(r, 5).Value = "Better"
        Else
            ws.Cells(r, 5).Value = "Flat"
        End If
        r = r + 1
    Next i

    WriteMovers = r
End Function

'==============================================================================
' HYGIENE  -  the dashboard is only as good as what gets coded
'==============================================================================
Private Function WriteHygiene(ByVal ws As Worksheet, ByVal r As Long, ByRef ytd As Period) As Long
    Dim wsQ As Worksheet, i As Long, lastR As Long
    Dim total As Long, unclassified As Long, noRoot As Long, noOp As Long
    Dim dt As Date

    Set wsQ = REVO_Core.GetSheet(SH_QUAL_LOG)
    If wsQ Is Nothing Then WriteHygiene = r: Exit Function

    lastR = wsQ.Cells(wsQ.Rows.Count, QLC_EVENTID).End(xlUp).row
    For i = QL_HEADER_ROW + 1 To lastR
        dt = REVO_Core.SafeDate(wsQ.Cells(i, QLC_DATE).Value, 0)
        If dt >= ytd.d1 And dt <= ytd.d2 Then
            total = total + 1
            If UCase$(REVO_Core.SafeS(wsQ.Cells(i, QLC_DEFECT).Value)) = "UNCLASSIFIED" Then unclassified = unclassified + 1
            If UCase$(REVO_Core.SafeS(wsQ.Cells(i, QLC_ROOT).Value)) = "NOT DETERMINED" Then noRoot = noRoot + 1
            If UCase$(REVO_Core.SafeS(wsQ.Cells(i, QLC_OP).Value)) = "UNKNOWN" Then noOp = noOp + 1
        End If
    Next i

    ws.Cells(r, 1).Value = "Events year to date"
    ws.Cells(r, 2).Value = total
    r = r + 1

    If total = 0 Then
        ws.Cells(r, 1).Value = "Quality Log is empty. Run REVO_BackfillQuality to load history."
        ws.Cells(r, 1).Font.Italic = True
        WriteHygiene = r + 1
        Exit Function
    End If

    ws.Cells(r, 1).Value = "Defect type unclassified"
    ws.Cells(r, 2).Value = unclassified
    ws.Cells(r, 3).Value = unclassified / total
    ws.Cells(r, 3).NumberFormat = "0%"
    ws.Cells(r, 7).Value = "Historic free text the classifier could not read. Filter " & SH_QUAL_LOG & _
                           " on Unclassified and recode - each one you fix improves every Pareto above."
    If unclassified / total > 0.2 Then ws.Cells(r, 3).Font.Color = RGB(192, 0, 0)
    r = r + 1

    ws.Cells(r, 1).Value = "Root cause not determined"
    ws.Cells(r, 2).Value = noRoot
    ws.Cells(r, 3).Value = noRoot / total
    ws.Cells(r, 3).NumberFormat = "0%"
    ws.Cells(r, 7).Value = "Root cause is what turns a defect count into a fix. Anything over 30% " & _
                           "here means the Pareto cannot tell you what to change."
    If noRoot / total > 0.3 Then ws.Cells(r, 3).Font.Color = RGB(192, 0, 0)
    r = r + 1

    ws.Cells(r, 1).Value = "Detected-at op unknown"
    ws.Cells(r, 2).Value = noOp
    ws.Cells(r, 3).Value = noOp / total
    ws.Cells(r, 3).NumberFormat = "0%"
    ws.Cells(r, 7).Value = "Without the op you cannot aim the fix at a station. Historic rows " & _
                           "mostly lack it; rows captured on the form will not."
    r = r + 1

    WriteHygiene = r
End Function

'==============================================================================
' SUPPORT
'==============================================================================
Private Function Tally(ByVal d1 As Date, ByVal d2 As Date, ByVal col As Long) As Object
    Dim ws As Worksheet, i As Long, lastR As Long
    Dim d As Object
    Dim dt As Date, k As String, q As Double

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set Tally = d

    Set ws = REVO_Core.GetSheet(SH_QUAL_LOG)
    If ws Is Nothing Then Exit Function

    lastR = ws.Cells(ws.Rows.Count, QLC_EVENTID).End(xlUp).row
    For i = QL_HEADER_ROW + 1 To lastR
        dt = REVO_Core.SafeDate(ws.Cells(i, QLC_DATE).Value, 0)
        If dt >= d1 And dt <= d2 Then
            k = REVO_Core.SafeS(ws.Cells(i, col).Value)
            If Len(k) = 0 Then k = "(blank)"
            q = REVO_Core.SafeD(ws.Cells(i, QLC_QTY).Value)
            If d.exists(k) Then d(k) = d(k) + q Else d.Add k, q
        End If
    Next i
End Function

' Dictionary to parallel arrays, sorted descending by value.
Private Function DictToArrays(ByVal d As Object, ByRef keys() As String, _
                              ByRef vals() As Double) As Long
    Dim k As Variant, n As Long, i As Long, j As Long
    Dim tk As String, tv As Double

    If d.Count = 0 Then DictToArrays = 0: Exit Function

    ReDim keys(1 To d.Count)
    ReDim vals(1 To d.Count)

    For Each k In d.keys
        n = n + 1
        keys(n) = CStr(k)
        vals(n) = CDbl(d(k))
    Next k

    For i = 1 To n - 1
        For j = i + 1 To n
            If vals(j) > vals(i) Then
                tv = vals(i): vals(i) = vals(j): vals(j) = tv
                tk = keys(i): keys(i) = keys(j): keys(j) = tk
            End If
        Next j
    Next i

    DictToArrays = n
End Function

Private Sub SumPeriod(ByVal d1 As Date, ByVal d2 As Date, _
                      ByRef rwk As Double, ByRef rej As Double, ByRef bg As Double)
    Dim ws As Worksheet, i As Long, lastR As Long
    Dim dt As Date, disp As String, q As Double

    rwk = 0: rej = 0: bg = 0

    Set ws = REVO_Core.GetSheet(SH_QUAL_LOG)
    If ws Is Nothing Then Exit Sub

    lastR = ws.Cells(ws.Rows.Count, QLC_EVENTID).End(xlUp).row
    For i = QL_HEADER_ROW + 1 To lastR
        dt = REVO_Core.SafeDate(ws.Cells(i, QLC_DATE).Value, 0)
        If dt >= d1 And dt <= d2 Then
            disp = UCase$(REVO_Core.SafeS(ws.Cells(i, QLC_DISP).Value))
            q = REVO_Core.SafeD(ws.Cells(i, QLC_QTY).Value)
            Select Case disp
                Case "REWORK":  rwk = rwk + q
                Case "REJECT":  rej = rej + q
                Case "B GRADE": bg = bg + q
            End Select
        End If
    Next i
End Sub


Private Function WeekStartOf(ByVal d As Date) As Date
    ' Production week runs Sunday to Friday.
    WeekStartOf = d - (Weekday(d, vbSunday) - 1)
End Function

Private Sub Section(ByVal ws As Worksheet, ByVal r As Long, ByVal t As String)
    ws.Cells(r, 1).Value = t
    ws.Cells(r, 1).Font.Size = 12
    ws.Cells(r, 1).Font.bold = True
    ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(31, 56, 100)
        .Weight = xlThin
    End With
End Sub

Private Sub HeaderBand(ByVal ws As Worksheet, ByVal r As Long, ByVal nCols As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nCols))
        .Font.bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With
End Sub
