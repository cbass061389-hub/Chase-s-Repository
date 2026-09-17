Attribute VB_Name = "REVO_CartVelocity"
'==============================================================================
' REVO_CartVelocity  -  measured cart movement, by diameter and joint type
'
' Entry point: REVO_UpdateCartVelocity   (called automatically by the plan
'              update and the scorecard; safe to run on its own)
'
' WHAT IT MEASURES
'   Floor Log is a snapshot table - one row per cart per LOG FLOOR run. Pairing
'   consecutive snapshots for the same cart gives how many ops that cart
'   advanced, and over how many production days. That is the observed velocity,
'   grouped by diameter and joint type.
'
'   The two capacity inputs on Cart Tracker are measured directly rather than
'   inferred: count the carts whose "IC Done" flag flipped between snapshots,
'   divide by production days, and that is the IC rate. Same for Op 10.
'
' OUTLIERS
'   A cart that sat on hold for nine days is not a slow cart, it is a stopped
'   one, and averaging it in drags the whole rate down. Three filters, in this
'   order:
'     1. structural - negative op deltas (cart re-loaded), zero-length
'        intervals, intervals over 14 days (a missed snapshot, not a slow cart)
'     2. explicit   - any cart on Holds during the window is dropped outright
'     3. statistical - median absolute deviation at 3 robust sigma
'   The median of what survives is the run rate. Median, not mean, so one
'   survivor at the edge cannot move it.
'
' GUARDRAILS ON WRITE-BACK  (this is the "auto-apply with guardrails" setting)
'   A measured rate is applied to Cart Tracker only when BOTH hold:
'     - at least MIN_SAMPLE clean intervals contributed
'     - the measured rate is within TOLERANCE of the value already there
'   Outside the band the sheet keeps the human number, the proposal is written
'   next to it, and the Action column says why it was held. A capacity
'   assumption never moves silently by more than a quarter.
'==============================================================================
Option Explicit

Public Const MIN_SAMPLE  As Long = 8       ' clean intervals needed before applying
Public Const TOLERANCE   As Double = 0.25  ' +/- 25% of the current value
Public Const MAX_GAP_DAYS As Long = 14     ' longer gap = missed snapshot, not data
Public Const MAD_SIGMAS  As Double = 3#

Private Type GateResult
    Rate       As Double
    Samples    As Long
    Dropped    As Long
    Days       As Double
    HasData    As Boolean
End Type

'==============================================================================
Public Sub REVO_UpdateCartVelocity(Optional ByVal silent As Boolean = False)
    Dim ws As Worksheet
    Dim ic As GateResult, op10 As GateResult
    Dim fams As Object
    Dim r As Long, k As Variant
    Dim applied As String, held As String
    Dim totalIntervals As Long

    On Error GoTo Failed
    REVO_Core.PushAppState

    Set fams = CreateObject("Scripting.Dictionary")
    fams.CompareMode = vbTextCompare

    ic = MeasureGate("IC Done")
    op10 = MeasureGate("Op 10 Done")
    totalIntervals = MeasureFamilies(fams)

    Set ws = REVO_Core.GetOrCreateSheet(SH_VELOCITY)
    ws.Cells.Clear

    '--------------------------------------------------------------- header
    ws.Range("A1").Value = "CART VELOCITY  -  measured from Floor Log snapshots"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.bold = True
    ws.Range("A2").Value = "Rebuilt " & Format$(Now, "ddd d mmm yyyy hh:nn") & _
                           " by " & REVO_Core.CurrentUser & ". Derived sheet - do not type here."
    ws.Range("A2").Font.Italic = True

    ws.Range("A4").Value = "GATE RATES  (these feed Cart Tracker)"
    ws.Range("A4").Font.bold = True

    ws.Range("A5").Value = "Snapshot pairs available"
    ws.Range("B5").Value = totalIntervals
    ws.Range("C5").Value = "Each LOG FLOOR run adds one snapshot. Rates firm up as this grows."

    ws.Range("A6").Value = "Minimum sample before auto-apply"
    ws.Range("B6").Value = MIN_SAMPLE
    ws.Range("C6").Value = "INPUT - raise it if you want the rate to settle further before it moves the plan."

    ws.Range("A7").Value = "Tolerance band"
    ws.Range("B7").Value = TOLERANCE
    ws.Range("B7").NumberFormat = "0%"
    ws.Range("C7").Value = "A measured rate further than this from the current input is flagged, not applied."

    '------------------------------------------------------- write the gates
    ws.Range("A9").Value = "Gate"
    ws.Range("B9").Value = "Current input"
    ws.Range("C9").Value = "Measured carts/day"
    ws.Range("D9").Value = "Clean samples"
    ws.Range("E9").Value = "Outliers dropped"
    ws.Range("F9").Value = "Variance"
    ws.Range("G9").Value = "Action"
    WriteHeaderBand ws, 9, 7

    ApplyGate ws, 10, "IC", 5, ic, applied, held
    ApplyGate ws, 11, "Op 10", 6, op10, applied, held

    '------------------------------------------------- per-family velocities
    ws.Range("A13").Value = "VELOCITY BY DIAMETER AND JOINT TYPE"
    ws.Range("A13").Font.bold = True
    ws.Range("A14").Value = "Ops advanced per production day, median of clean intervals. " & _
                            "Total ops in routing: " & OpsInRouting() & "."
    ws.Range("A14").Font.Italic = True

    ws.Range("A16").Value = "Diameter"
    ws.Range("B16").Value = "Joint"
    ws.Range("C16").Value = "Clean intervals"
    ws.Range("D16").Value = "Dropped"
    ws.Range("E16").Value = "Median ops/day"
    ws.Range("F16").Value = "Days per cart (full routing)"
    ws.Range("G16").Value = "Confidence"
    WriteHeaderBand ws, 16, 7

    r = 17
    For Each k In fams.keys
        Dim fam As Variant
        fam = fams(k)                 ' Array(rates(), n, dropped)
        Dim rates() As Double, n As Long, dropped As Long
        rates = fam(0): n = fam(1): dropped = fam(2)

        ws.Cells(r, 1).Value = Split(CStr(k), "|")(0)
        ws.Cells(r, 2).Value = Split(CStr(k), "|")(1)
        ws.Cells(r, 3).Value = n
        ws.Cells(r, 4).Value = dropped
        If n > 0 Then
            Dim med As Double
            med = REVO_Core.Median(rates, n)
            ws.Cells(r, 5).Value = med
            ws.Cells(r, 5).NumberFormat = "0.00"
            If med > 0 Then
                ws.Cells(r, 6).Value = OpsInRouting() / med
                ws.Cells(r, 6).NumberFormat = "0.0"
            End If
        End If
        ws.Cells(r, 7).Value = ConfidenceLabel(n)
        If n < 3 Then ws.Cells(r, 7).Font.Color = RGB(192, 0, 0)
        r = r + 1
    Next k

    If fams.Count = 0 Then
        ws.Cells(17, 1).Value = "No usable snapshot pairs yet."
        ws.Cells(17, 1).Font.Italic = True
    End If

    '------------------------------------------------------------- footnote
    r = r + 2
    ws.Cells(r, 1).Value = "WHAT HAPPENED THIS RUN"
    ws.Cells(r, 1).Font.bold = True
    r = r + 1
    If Len(applied) > 0 Then
        ws.Cells(r, 1).Value = "Applied:"
        ws.Cells(r, 2).Value = applied
        r = r + 1
    End If
    If Len(held) > 0 Then
        ws.Cells(r, 1).Value = "Held:"
        ws.Cells(r, 2).Value = held
        ws.Cells(r, 2).Font.Color = RGB(150, 75, 0)
        r = r + 1
    End If
    If Len(applied) = 0 And Len(held) = 0 Then
        ws.Cells(r, 1).Value = "Nothing measurable yet - Floor Log needs at least two snapshots " & _
                               "with carts in common. Run LOG FLOOR daily."
        ws.Cells(r, 1).Font.Italic = True
    End If

    ws.Columns("A:B").ColumnWidth = 26
    ws.Columns("C:F").ColumnWidth = 18
    ws.Columns("G:G").ColumnWidth = 60

    REVO_Core.PopAppState
    REVO_Core.Audit "REVO_CartVelocity", "MEASURE", SH_CART_TRK, _
                    totalIntervals & " intervals; applied[" & applied & "] held[" & held & "]"

    If Not silent Then
        MsgBox "Cart velocity updated." & vbCrLf & vbCrLf & _
               "Snapshot pairs: " & totalIntervals & vbCrLf & _
               IIf(Len(applied) > 0, "Applied: " & applied & vbCrLf, "") & _
               IIf(Len(held) > 0, "Held: " & held & vbCrLf, "") & vbCrLf & _
               "Detail on the '" & SH_VELOCITY & "' sheet.", vbInformation, "Cart Velocity"
    End If
    Exit Sub

Failed:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    REVO_Core.ResetAppState
    REVO_Core.Audit "REVO_CartVelocity", "ERROR", SH_VELOCITY, errNum & ": " & errDesc
    If Not silent Then
        MsgBox "Cart velocity update failed." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errDesc, vbExclamation, "Cart Velocity"
    End If
    Resume CleanExit
CleanExit:
End Sub

'==============================================================================
' GATE MEASUREMENT
'
' Counts carts whose boolean gate flag flipped from FALSE to TRUE between two
' consecutive snapshots, over the production days between them.
'==============================================================================
Private Function MeasureGate(ByVal flagCaption As String) As GateResult
    Dim ws As Worksheet, hdr As Long, lastR As Long, r As Long
    Dim colDate As Long, cCart As Long, cFlag As Long
    Dim snaps As Object, cartState As Object
    Dim res As GateResult
    Dim rates() As Double, n As Long
    Dim i As Long, j As Long
    Dim snapList As Variant
    Dim d1 As Date, d2 As Date
    Dim k As Variant, cartKey As String
    Dim pd As Long, flips As Long

    Set ws = REVO_Core.GetSheet(SH_FLOOR_LOG)
    If ws Is Nothing Then MeasureGate = res: Exit Function

    hdr = REVO_Core.FindHeaderRow(ws, Array("Snapshot Date", "Cart"), 10, 20)
    If hdr = 0 Then MeasureGate = res: Exit Function

    colDate = REVO_Core.FindCol(ws, hdr, Array("Snapshot Date"))
    cCart = REVO_Core.FindCol(ws, hdr, Array("Cart"))
    cFlag = REVO_Core.FindCol(ws, hdr, Array(flagCaption))
    If colDate = 0 Or cCart = 0 Or cFlag = 0 Then MeasureGate = res: Exit Function

    ' snapshot date -> dictionary of cart -> flag
    Set snaps = CreateObject("Scripting.Dictionary")
    lastR = REVO_Core.LastDataRow(ws, cCart, hdr)

    For r = hdr + 1 To lastR
        d1 = REVO_Core.SafeDate(ws.Cells(r, colDate).Value, 0)
        cartKey = UCase$(REVO_Core.SafeS(ws.Cells(r, cCart).Value))
        If d1 > 0 And Len(cartKey) > 0 Then
            If Not snaps.exists(CLng(d1)) Then
                snaps.Add CLng(d1), CreateObject("Scripting.Dictionary")
            End If
            Set cartState = snaps(CLng(d1))
            If Not cartState.exists(cartKey) Then
                cartState.Add cartKey, REVO_Core.Boolish(ws.Cells(r, cFlag).Value)
            End If
        End If
    Next r

    If snaps.Count < 2 Then MeasureGate = res: Exit Function

    snapList = snaps.keys
    SortLongArray snapList

    ReDim rates(1 To snaps.Count)

    For i = LBound(snapList) To UBound(snapList) - 1
        d1 = CDate(CLng(snapList(i)))
        d2 = CDate(CLng(snapList(i + 1)))
        pd = REVO_Core.ProductionDaysBetween(d1 + 1, d2)
        If pd <= 0 Then pd = 1
        If (d2 - d1) <= MAX_GAP_DAYS Then
            flips = 0
            Dim a As Object, b As Object
            Set a = snaps(CLng(snapList(i)))
            Set b = snaps(CLng(snapList(i + 1)))
            For Each k In b.keys
                If a.exists(k) Then
                    If (Not CBool(a(k))) And CBool(b(k)) Then
                        If Not CartOnHoldDuring(CStr(k), d1, d2) Then flips = flips + 1
                    End If
                End If
            Next k
            n = n + 1
            rates(n) = flips / pd
            res.Days = res.Days + pd
        Else
            res.Dropped = res.Dropped + 1
        End If
    Next i

    If n = 0 Then MeasureGate = res: Exit Function

    Dim kept As Long
    kept = REVO_Core.RejectOutliers(rates, n, MAD_SIGMAS)
    res.Dropped = res.Dropped + (n - kept)
    res.Samples = kept
    res.Rate = REVO_Core.Median(rates, kept)
    res.HasData = (kept > 0)

    MeasureGate = res
End Function

'==============================================================================
' PER-FAMILY VELOCITY  (ops advanced per production day)
'==============================================================================
Private Function MeasureFamilies(ByRef fams As Object) As Long
    Dim ws As Worksheet, hdr As Long, lastR As Long, r As Long
    Dim colDate As Long, cCart As Long, cSku As Long, cOps As Long, cWO As Long
    Dim hist As Object
    Dim key As String, cartKey As String
    Dim d As Date, ops As Double
    Dim total As Long

    Set ws = REVO_Core.GetSheet(SH_FLOOR_LOG)
    If ws Is Nothing Then Exit Function

    hdr = REVO_Core.FindHeaderRow(ws, Array("Snapshot Date", "Cart"), 10, 20)
    If hdr = 0 Then Exit Function

    colDate = REVO_Core.FindCol(ws, hdr, Array("Snapshot Date"))
    cCart = REVO_Core.FindCol(ws, hdr, Array("Cart"))
    cSku = REVO_Core.FindCol(ws, hdr, Array("SKU"))
    cOps = REVO_Core.FindCol(ws, hdr, Array("Ops Done"))
    cWO = REVO_Core.FindCol(ws, hdr, Array("WO", "WO#"))
    If colDate = 0 Or cCart = 0 Or cSku = 0 Or cOps = 0 Then Exit Function

    ' cart|WO -> collection of "date|opsDone|sku"
    Set hist = CreateObject("Scripting.Dictionary")
    lastR = REVO_Core.LastDataRow(ws, cCart, hdr)

    For r = hdr + 1 To lastR
        d = REVO_Core.SafeDate(ws.Cells(r, colDate).Value, 0)
        cartKey = UCase$(REVO_Core.SafeS(ws.Cells(r, cCart).Value))
        If d > 0 And Len(cartKey) > 0 Then
            key = cartKey & "|" & UCase$(IIf(cWO > 0, REVO_Core.SafeS(ws.Cells(r, cWO).Value), ""))
            If Not hist.exists(key) Then hist.Add key, New Collection
            hist(key).Add Array(CLng(d), REVO_Core.SafeD(ws.Cells(r, cOps).Value), _
                                REVO_Core.SafeS(ws.Cells(r, cSku).Value), cartKey)
        End If
    Next r

    '--- pair consecutive observations per cart -----------------------------
    Dim kk As Variant, col As Collection
    Dim i As Long, j As Long
    Dim rec1 As Variant, rec2 As Variant
    Dim pd As Long, delta As Double, rate As Double
    Dim a As SkuAttrs, famKey As String
    Dim buf As Object
    Set buf = CreateObject("Scripting.Dictionary")
    buf.CompareMode = vbTextCompare

    For Each kk In hist.keys
        Set col = hist(kk)
        ' order by date (collections are small - insertion sort is fine)
        Dim arr() As Variant, nn As Long
        nn = col.Count
        If nn >= 2 Then
            ReDim arr(1 To nn)
            For i = 1 To nn
                arr(i) = col(i)
            Next i
            For i = 1 To nn - 1
                For j = i + 1 To nn
                    If CLng(arr(j)(0)) < CLng(arr(i)(0)) Then
                        Dim tmp As Variant
                        tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                    End If
                Next j
            Next i

            For i = 1 To nn - 1
                rec1 = arr(i): rec2 = arr(i + 1)
                delta = CDbl(rec2(1)) - CDbl(rec1(1))
                pd = REVO_Core.ProductionDaysBetween(CDate(CLng(rec1(0))) + 1, CDate(CLng(rec2(0))))

                ' structural filters
                If delta >= 0 _
                   And pd > 0 _
                   And (CLng(rec2(0)) - CLng(rec1(0))) <= MAX_GAP_DAYS _
                   And Not CartOnHoldDuring(CStr(rec1(3)), CDate(CLng(rec1(0))), CDate(CLng(rec2(0)))) Then

                    rate = delta / pd
                    a = REVO_Core.ParseSku(CStr(rec2(2)))
                    famKey = a.Diameter & "|" & a.Joint
                    If Not buf.exists(famKey) Then buf.Add famKey, New Collection
                    buf(famKey).Add rate
                    total = total + 1
                End If
            Next i
        End If
    Next kk

    '--- outlier-reject each family ----------------------------------------
    Dim rates() As Double, n As Long, kept As Long
    For Each kk In buf.keys
        Set col = buf(kk)
        n = col.Count
        If n > 0 Then
            ReDim rates(1 To n)
            For i = 1 To n
                rates(i) = CDbl(col(i))
            Next i
            kept = REVO_Core.RejectOutliers(rates, n, MAD_SIGMAS)
            fams.Add kk, Array(rates, kept, n - kept)
        End If
    Next kk

    MeasureFamilies = total
End Function

'==============================================================================
' WRITE-BACK WITH GUARDRAILS
'==============================================================================
Private Sub ApplyGate(ByVal ws As Worksheet, ByVal row As Long, ByVal label As String, _
                      ByVal trackerRow As Long, ByRef g As GateResult, _
                      ByRef applied As String, ByRef held As String)
    Dim wsT As Worksheet
    Dim cur As Double, variance As Double
    Dim action As String

    Set wsT = REVO_Core.GetSheet(SH_CART_TRK)
    If Not wsT Is Nothing Then cur = REVO_Core.SafeD(wsT.Cells(trackerRow, 2).Value)

    ws.Cells(row, 1).Value = label
    ws.Cells(row, 2).Value = cur
    ws.Cells(row, 2).NumberFormat = "0.00"

    If Not g.HasData Then
        ws.Cells(row, 7).Value = "No measurement yet - Floor Log needs two snapshots sharing carts."
        ws.Cells(row, 7).Font.Italic = True
        held = held & label & " (no data); "
        Exit Sub
    End If

    ws.Cells(row, 3).Value = g.Rate
    ws.Cells(row, 3).NumberFormat = "0.00"
    ws.Cells(row, 4).Value = g.Samples
    ws.Cells(row, 5).Value = g.Dropped

    If cur > 0 Then
        variance = (g.Rate - cur) / cur
        ws.Cells(row, 6).Value = variance
        ws.Cells(row, 6).NumberFormat = "+0%;-0%"
    End If

    '--- guardrail 1: sample size ------------------------------------------
    If g.Samples < MIN_SAMPLE Then
        action = "HELD - " & g.Samples & " clean sample(s), need " & MIN_SAMPLE & _
                 ". Measured rate shown for information only."
        ws.Cells(row, 7).Value = action
        ws.Cells(row, 7).Font.Color = RGB(150, 75, 0)
        held = held & label & " (sample " & g.Samples & "/" & MIN_SAMPLE & "); "
        Exit Sub
    End If

    '--- guardrail 2: tolerance band ---------------------------------------
    If cur > 0 And Abs(variance) > TOLERANCE Then
        action = "HELD - measured rate is " & Format$(variance, "+0%;-0%") & " from the current " & _
                 "input, outside the " & Format$(TOLERANCE, "0%") & " band. Review before accepting."
        ws.Cells(row, 7).Value = action
        ws.Cells(row, 7).Font.Color = RGB(192, 0, 0)
        held = held & label & " (" & Format$(variance, "+0%;-0%") & " swing); "
        Exit Sub
    End If

    '--- apply --------------------------------------------------------------
    If Not wsT Is Nothing Then
        wsT.Cells(trackerRow, 2).Value = Round(g.Rate, 2)
        action = "APPLIED - Cart Tracker B" & trackerRow & " set to " & Format$(g.Rate, "0.00") & _
                 " from " & g.Samples & " clean samples."
        ws.Cells(row, 7).Value = action
        ws.Cells(row, 7).Font.Color = RGB(0, 97, 0)
        applied = applied & label & "=" & Format$(g.Rate, "0.00") & "; "
        REVO_Core.Audit "REVO_CartVelocity", "APPLY", SH_CART_TRK & "!B" & trackerRow, _
                        "was " & cur & ", now " & Round(g.Rate, 2) & _
                        " (" & g.Samples & " samples)"
    End If
End Sub

'==============================================================================
' SUPPORT
'==============================================================================
' A cart sitting on an unreleased hold during the window is stopped, not slow.
Private Function CartOnHoldDuring(ByVal cart As String, ByVal d1 As Date, ByVal d2 As Date) As Boolean
    Dim ws As Worksheet, hdr As Long, lastR As Long, r As Long
    Dim cCart As Long, cHeld As Long, cRel As Long
    Dim hd As Date

    Set ws = REVO_Core.GetSheet(SH_HOLDS)
    If ws Is Nothing Then Exit Function

    hdr = REVO_Core.FindHeaderRow(ws, Array("Cart #", "Cart"), 10, 10)
    If hdr = 0 Then Exit Function

    cCart = REVO_Core.FindCol(ws, hdr, Array("Cart #", "Cart"))
    cHeld = REVO_Core.FindCol(ws, hdr, Array("Date Held"))
    cRel = REVO_Core.FindCol(ws, hdr, Array("Released?", "Released"))
    If cCart = 0 Then Exit Function

    lastR = REVO_Core.LastDataRow(ws, cCart, hdr)
    For r = hdr + 1 To lastR
        If UCase$(REVO_Core.SafeS(ws.Cells(r, cCart).Value)) = UCase$(Trim$(cart)) Then
            hd = REVO_Core.SafeDate(ws.Cells(r, cHeld).Value, 0)
            If hd > 0 And hd <= d2 Then
                ' Still held, or released after the window opened.
                If cRel = 0 Then
                    CartOnHoldDuring = True: Exit Function
                ElseIf Not REVO_Core.Boolish(ws.Cells(r, cRel).Value) Then
                    CartOnHoldDuring = True: Exit Function
                ElseIf hd >= d1 Then
                    CartOnHoldDuring = True: Exit Function
                End If
            End If
        End If
    Next r
End Function

Public Function OpsInRouting() As Long
    Dim ws As Worksheet, hdr As Long, r As Long, lastR As Long, n As Long
    Set ws = REVO_Core.GetSheet(SH_PLAN_CFG)
    If ws Is Nothing Then OpsInRouting = DEFAULT_OPS_IN_ROUTING: Exit Function

    hdr = REVO_Core.FindHeaderRow(ws, Array("Op"), 10, 6)
    If hdr = 0 Then OpsInRouting = DEFAULT_OPS_IN_ROUTING: Exit Function

    lastR = REVO_Core.LastDataRow(ws, 1, hdr)
    For r = hdr + 1 To lastR
        If Len(REVO_Core.SafeS(ws.Cells(r, 1).Value)) > 0 Then
            If REVO_Core.SafeD(ws.Cells(r, 2).Value) > 0 Then n = n + 1
        End If
    Next r
    If n = 0 Then n = DEFAULT_OPS_IN_ROUTING
    OpsInRouting = n
End Function

Private Function ConfidenceLabel(ByVal n As Long) As String
    Select Case True
        Case n >= MIN_SAMPLE:  ConfidenceLabel = "Good - " & n & " clean intervals"
        Case n >= 3:           ConfidenceLabel = "Thin - " & n & " intervals, treat as indicative"
        Case n >= 1:           ConfidenceLabel = "Not enough data - " & n & " interval(s)"
        Case Else:             ConfidenceLabel = "No data"
    End Select
End Function

Private Sub WriteHeaderBand(ByVal ws As Worksheet, ByVal r As Long, ByVal nCols As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nCols))
        .Font.bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

Private Sub SortLongArray(ByRef a As Variant)
    Dim i As Long, j As Long, t As Variant
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If CLng(a(j)) < CLng(a(i)) Then
                t = a(i): a(i) = a(j): a(j) = t
            End If
        Next j
    Next i
End Sub
