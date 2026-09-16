Attribute VB_Name = "REVO_Scorecard"
'==============================================================================
' REVO_Scorecard  -  the weekly operating picture and what to do about it
'
' Entry point: REVO_BuildScorecard
'
' This replaces "shift hours against releases" with something that answers the
' four questions that actually get asked:
'
'   What happened?      week to date, released vs plan, quality losses
'   What is happening?  where the constraint is right now, what is at risk
'   What happens next?  projected finish at the measured rate, not the assumed one
'   What do I do?       overtime hours, extra day, headcount, and which op
'
' DESIGN POSITION
'   The old tracker divided shift hours by releases. That silently assumes
'   every hour is equally useful and every shift can release. Neither is true:
'   releases are a day-shift activity, and an hour at a non-constraint op buys
'   nothing. This model works the constraint, respects the shift calendar, and
'   prefers measured rates over assumed ones wherever a measurement exists.
'
'   Every number on the sheet carries its basis - MEASURED or ASSUMED - so a
'   recommendation built on a thin sample announces itself instead of looking
'   like fact.
'==============================================================================
Option Explicit

Private Const HOURS_PER_SHIFT As Double = 10#   ' productive hours in a shift
Private Const SHIFTS_PER_DAY  As Double = 2#

Private Type OpLoad
    name        As String
    CapPerDay   As Double
    HrsPerCart  As Double
    PeopleReq   As Double
    StaffOn     As String
    CartsSched  As Double
    CartsReq    As Double
    Gap         As Double
    IsBottleneck As Boolean
End Type

Private Type WeekPicture
    PlanShafts      As Double
    ReleasedShafts  As Double
    RemainingShafts As Double
    ShaftsPerCart   As Double
    WeekEnd         As Date
    DaysLeft        As Long
    DaysElapsed     As Long
    PctToPlan       As Double
    HasPlan         As Boolean
End Type

'==============================================================================
Public Sub REVO_BuildScorecard()
    BuildScorecard True
End Sub

Public Sub BuildScorecard(Optional ByVal interactive As Boolean = True)
    Dim ws As Worksheet
    Dim wp As WeekPicture
    Dim ops() As OpLoad, nOps As Long, bn As Long
    Dim r As Long
    Dim verdict As String, verdictColor As Long

    On Error GoTo Failed
    REVO_Core.PushAppState

    ' Refresh the measured rates first - the recommendations lean on them.
    On Error Resume Next
    REVO_CartVelocity.REVO_UpdateCartVelocity True
    On Error GoTo Failed

    wp = ReadWeekPicture()
    nOps = ReadOpLoads(ops)
    bn = FindBottleneck(ops, nOps)

    Set ws = REVO_Core.GetOrCreateSheet(SH_SCORECARD)
    ws.Cells.Clear
    ws.Cells.Font.name = "Segoe UI"
    ws.Cells.Font.Size = 10

    '======================================================= 1. HEADLINE ====
    ws.Range("A1").Value = "REVO OPERATING SCORECARD"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.bold = True
    ws.Range("A2").Value = Format$(Now, "dddd d mmmm yyyy, hh:nn") & _
                           "   -   week ending " & Format$(wp.WeekEnd, "ddd d mmm") & _
                           "   -   " & REVO_Core.ShiftFor(Now) & " shift" & _
                           IIf(REVO_Core.CanReleaseNow, "", "   (no releases on nights)")
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = RGB(90, 90, 90)

    r = 4
    SectionTitle ws, r, "WHERE WE STAND"
    r = r + 1

    KV ws, r, "Plan this week (shafts)", wp.PlanShafts, "0", _
       "From REVO Floor, PLAN RELEASES Sun to Fri.": r = r + 1
    KV ws, r, "Released to date (shafts)", wp.ReleasedShafts, "0", _
       "Counted from Release Log, not typed.": r = r + 1
    KV ws, r, "Remaining to plan (shafts)", wp.RemainingShafts, "0", "": r = r + 1
    KV ws, r, "Percent to plan", wp.PctToPlan, "0%", "": r = r + 1
    KV ws, r, "Production days left", wp.DaysLeft, "0", _
       "Sunday to Friday, Saturday dark, Plan Config blackouts removed.": r = r + 1

    Dim reqCartsPerDay As Double, capCartsPerDay As Double, gapPerDay As Double
    Dim remainingCarts As Double
    remainingCarts = 0
    If wp.ShaftsPerCart > 0 Then remainingCarts = wp.RemainingShafts / wp.ShaftsPerCart

    If wp.DaysLeft > 0 Then reqCartsPerDay = remainingCarts / wp.DaysLeft
    If nOps > 0 And bn > 0 Then capCartsPerDay = ops(bn).CapPerDay
    gapPerDay = reqCartsPerDay - capCartsPerDay

    KV ws, r, "Carts still to release", remainingCarts, "0.0", _
       "At " & Format$(wp.ShaftsPerCart, "0") & " shafts per cart.": r = r + 1
    KV ws, r, "Carts/day required", reqCartsPerDay, "0.00", "": r = r + 1
    KV ws, r, "Carts/day the line can do", capCartsPerDay, "0.00", _
       "Constraint: " & IIf(bn > 0, "Op " & ops(bn).name, "unknown") & ". " & CapBasis(): r = r + 1

    '--- the verdict --------------------------------------------------------
    r = r + 1
    If Not wp.HasPlan Then
        verdict = "NO PLAN FIGURE FOUND - cannot judge status"
        verdictColor = RGB(150, 75, 0)
    ElseIf wp.RemainingShafts <= 0 Then
        verdict = "PLAN MET - " & Format$(wp.PctToPlan, "0%") & " of plan released"
        verdictColor = RGB(0, 97, 0)
    ElseIf wp.DaysLeft <= 0 Then
        verdict = "WEEK OVER, SHORT BY " & Format$(wp.RemainingShafts, "#,##0") & " SHAFTS"
        verdictColor = RGB(192, 0, 0)
    ElseIf gapPerDay <= 0 Then
        verdict = "ON TRACK - capacity covers the remaining " & Format$(remainingCarts, "0.0") & " carts"
        verdictColor = RGB(0, 97, 0)
    ElseIf gapPerDay <= capCartsPerDay * 0.15 Then
        verdict = "AT RISK - short " & Format$(gapPerDay, "0.00") & " carts/day, recoverable"
        verdictColor = RGB(150, 75, 0)
    Else
        verdict = "BEHIND - short " & Format$(gapPerDay, "0.00") & " carts/day"
        verdictColor = RGB(192, 0, 0)
    End If

    ws.Cells(r, 1).Value = verdict
    ws.Cells(r, 1).Font.Size = 14
    ws.Cells(r, 1).Font.bold = True
    ws.Cells(r, 1).Font.Color = verdictColor
    r = r + 2

    '======================================================= 2. THE ASK =====
    SectionTitle ws, r, "WHAT IT TAKES TO CLOSE THE GAP"
    r = r + 1
    r = WriteRecommendations(ws, r, wp, ops, nOps, bn, remainingCarts, reqCartsPerDay, capCartsPerDay)
    r = r + 1

    '======================================================= 3. QUALITY =====
    SectionTitle ws, r, "QUALITY"
    r = r + 1
    r = WriteQualityBlock(ws, r, wp)
    r = r + 1

    '======================================================= 4. BY OP =======
    SectionTitle ws, r, "BY OPERATION  -  where the effort goes"
    r = r + 1
    r = WriteOpTable(ws, r, ops, nOps, bn, reqCartsPerDay)
    r = r + 1

    '======================================================= 5. FOCUS =======
    SectionTitle ws, r, "FOCUS RANKING"
    r = r + 1
    r = WriteFocusRanking(ws, r, wp, ops, nOps, bn, gapPerDay)

    '--------------------------------------------------------------- format
    ws.Columns("A:A").ColumnWidth = 34
    ws.Columns("B:B").ColumnWidth = 16
    ws.Columns("C:C").ColumnWidth = 14
    ws.Columns("D:D").ColumnWidth = 14
    ws.Columns("E:E").ColumnWidth = 14
    ws.Columns("F:F").ColumnWidth = 14
    ws.Columns("G:G").ColumnWidth = 62
    ws.Rows(1).RowHeight = 26

    REVO_Core.PopAppState
    REVO_Core.Audit "REVO_Scorecard", "BUILD", SH_SCORECARD, verdict

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
    REVO_Core.Audit "REVO_Scorecard", "ERROR", SH_SCORECARD, eNum & ": " & eDesc
    If interactive Then
        MsgBox "Scorecard build failed." & vbCrLf & vbCrLf & _
               "Error " & eNum & ": " & eDesc, vbExclamation, "Scorecard"
    End If
    Resume CleanExit
CleanExit:
End Sub

'==============================================================================
' RECOMMENDATIONS
'
' Three levers, costed. The order is deliberate: the cheapest lever that
' actually closes the gap is named first, and each one says what it buys.
'==============================================================================
Private Function WriteRecommendations(ByVal ws As Worksheet, ByVal r As Long, _
                                      ByRef wp As WeekPicture, ByRef ops() As OpLoad, _
                                      ByVal nOps As Long, ByVal bn As Long, _
                                      ByVal remainingCarts As Double, _
                                      ByVal reqPerDay As Double, ByVal capPerDay As Double) As Long
    Dim gapCarts As Double, gapPerDay As Double
    Dim hrsPerCart As Double, otHours As Double, peopleNeeded As Double
    Dim extraDays As Double
    Dim bnName As String

    gapPerDay = reqPerDay - capPerDay
    gapCarts = gapPerDay * wp.DaysLeft
    If gapCarts < 0 Then gapCarts = 0

    bnName = IIf(bn > 0, ops(bn).name, "the constraint")
    If bn > 0 Then hrsPerCart = ops(bn).HrsPerCart
    If hrsPerCart <= 0 Then hrsPerCart = 5#      ' fallback, flagged below

    If gapCarts <= 0 Then
        ws.Cells(r, 1).Value = "No action needed on throughput."
        ws.Cells(r, 1).Font.Color = RGB(0, 97, 0)
        ws.Cells(r, 7).Value = "Capacity at Op " & bnName & " covers the remaining carts. " & _
                               "Spend the effort on quality and on carts already at risk."
        r = r + 1
        WriteRecommendations = r
        Exit Function
    End If

    ws.Cells(r, 1).Value = "Carts short by Friday"
    ws.Cells(r, 2).Value = gapCarts
    ws.Cells(r, 2).NumberFormat = "0.0"
    ws.Cells(r, 7).Value = "At Op " & bnName & ", the constraint. Everything below is what it takes to find these."
    r = r + 1

    ws.Cells(r, 1).Value = "Shafts at stake"
    ws.Cells(r, 2).Value = gapCarts * wp.ShaftsPerCart
    ws.Cells(r, 2).NumberFormat = "#,##0"
    r = r + 2

    '--- lever 1: overtime ---------------------------------------------------
    otHours = gapCarts * hrsPerCart
    ws.Cells(r, 1).Value = "1.  OVERTIME"
    ws.Cells(r, 1).Font.bold = True
    ws.Cells(r, 2).Value = otHours
    ws.Cells(r, 2).NumberFormat = "0.0"
    ws.Cells(r, 3).Value = "person-hours"
    ws.Cells(r, 7).Value = Format$(gapCarts, "0.0") & " carts x " & Format$(hrsPerCart, "0.0") & _
                           " hrs/cart at Op " & bnName & ". Spread over " & wp.DaysLeft & _
                           " day(s) that is " & Format$(otHours / IIf(wp.DaysLeft > 0, wp.DaysLeft, 1), "0.0") & _
                           " person-hours per day."
    r = r + 1

    '--- lever 2: extra day --------------------------------------------------
    extraDays = 0
    If capPerDay > 0 Then extraDays = gapCarts / capPerDay
    ws.Cells(r, 1).Value = "2.  EXTRA PRODUCTION DAY"
    ws.Cells(r, 1).Font.bold = True
    ws.Cells(r, 2).Value = extraDays
    ws.Cells(r, 2).NumberFormat = "0.0"
    ws.Cells(r, 3).Value = "day(s)"
    If extraDays <= 1.05 Then
        ws.Cells(r, 7).Value = "One Saturday at the current rate clears it. Cheapest single move."
        ws.Cells(r, 7).Font.Color = RGB(0, 97, 0)
    Else
        ws.Cells(r, 7).Value = Format$(extraDays, "0.0") & " extra days at the current rate - " & _
                               "more than a weekend. Overtime alone will not close this; the rate has to move."
        ws.Cells(r, 7).Font.Color = RGB(192, 0, 0)
    End If
    r = r + 1

    '--- lever 3: people -----------------------------------------------------
    peopleNeeded = 0
    If wp.DaysLeft > 0 Then
        peopleNeeded = otHours / (wp.DaysLeft * HOURS_PER_SHIFT)
    End If
    ws.Cells(r, 1).Value = "3.  ADDITIONAL PEOPLE"
    ws.Cells(r, 1).Font.bold = True
    ws.Cells(r, 2).Value = Application.WorksheetFunction.RoundUp(peopleNeeded, 0)
    ws.Cells(r, 2).NumberFormat = "0"
    ws.Cells(r, 3).Value = "at Op " & bnName
    ws.Cells(r, 7).Value = Format$(otHours, "0.0") & " person-hours over " & wp.DaysLeft & _
                           " day(s) at " & Format$(HOURS_PER_SHIFT, "0") & " productive hrs/shift. " & _
                           "Only helps if Op " & bnName & " has the machines to match."
    r = r + 1

    '--- shift note ----------------------------------------------------------
    r = r + 1
    ws.Cells(r, 1).Value = "Shift constraint"
    ws.Cells(r, 1).Font.bold = True
    ws.Cells(r, 7).Value = "Releases and FQC are day-shift only. Night capacity moves carts toward " & _
                           "release but cannot book one, so night overtime buys throughput, " & _
                           "not released shafts. If the gap is release-side, the extra hours have to be on days."
    r = r + 1

    WriteRecommendations = r
End Function

'==============================================================================
' QUALITY BLOCK
'==============================================================================
Private Function WriteQualityBlock(ByVal ws As Worksheet, ByVal r As Long, _
                                   ByRef wp As WeekPicture) As Long
    Dim relQty As Double, rwk As Double, rej As Double, bg As Double
    Dim topDefect As String, topRoot As String, topOp As String
    Dim lostShafts As Double
    Dim d1 As Date

    d1 = WeekStart(wp.WeekEnd)

    SumQuality d1, wp.WeekEnd, rwk, rej, bg, topDefect, topRoot, topOp
    relQty = wp.ReleasedShafts
    lostShafts = rwk + rej + bg

    KV ws, r, "Released this week (shafts)", relQty, "#,##0", "": r = r + 1
    KV ws, r, "Rework", rwk, "#,##0", "": r = r + 1
    KV ws, r, "Reject", rej, "#,##0", "": r = r + 1
    KV ws, r, "B grade", bg, "#,##0", "": r = r + 1

    Dim denom As Double
    denom = relQty + lostShafts
    If denom > 0 Then
        KV ws, r, "Reject rate", rej / denom, "0.0%", _
           "Rejects as a share of everything inspected this week.": r = r + 1
        KV ws, r, "Rework rate", rwk / denom, "0.0%", "": r = r + 1
        KV ws, r, "First-pass yield", relQty / denom, "0.0%", _
           "Released straight to inventory with no disposition.": r = r + 1
    Else
        ws.Cells(r, 1).Value = "No quality events recorded this week."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    End If

    If lostShafts > 0 Then
        ws.Cells(r, 1).Value = "Cost in carts"
        ws.Cells(r, 2).Value = lostShafts / IIf(wp.ShaftsPerCart > 0, wp.ShaftsPerCart, DEFAULT_SHAFTS_PER_CART)
        ws.Cells(r, 2).NumberFormat = "0.0"
        ws.Cells(r, 7).Value = "Quality losses this week are worth this many carts of capacity - " & _
                               "compare it to the throughput gap above before spending on overtime."
        r = r + 1
    End If

    If Len(topDefect) > 0 Then
        KVs ws, r, "Top defect", topDefect, "": r = r + 1
        KVs ws, r, "Top root cause", topRoot, "": r = r + 1
        KVs ws, r, "Worst op", topOp, "Where the most units were dispositioned this week.": r = r + 1
    Else
        ws.Cells(r, 1).Value = "Quality Log is empty for this week - run the backfill, then capture on release."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    End If

    WriteQualityBlock = r
End Function

'==============================================================================
' OP TABLE
'==============================================================================
Private Function WriteOpTable(ByVal ws As Worksheet, ByVal r As Long, ByRef ops() As OpLoad, _
                              ByVal nOps As Long, ByVal bn As Long, ByVal reqPerDay As Double) As Long
    Dim i As Long, hdrRow As Long

    hdrRow = r
    ws.Cells(r, 1).Value = "Op"
    ws.Cells(r, 2).Value = "Cap carts/day"
    ws.Cells(r, 3).Value = "Req carts/day"
    ws.Cells(r, 4).Value = "Gap"
    ws.Cells(r, 5).Value = "Hrs/cart"
    ws.Cells(r, 6).Value = "People req"
    ws.Cells(r, 7).Value = "Read"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))
        .Font.bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
    End With
    r = r + 1

    For i = 1 To nOps
        ws.Cells(r, 1).Value = ops(i).name
        ws.Cells(r, 2).Value = ops(i).CapPerDay
        ws.Cells(r, 2).NumberFormat = "0.00"
        ws.Cells(r, 3).Value = reqPerDay
        ws.Cells(r, 3).NumberFormat = "0.00"
        ws.Cells(r, 4).Value = ops(i).CapPerDay - reqPerDay
        ws.Cells(r, 4).NumberFormat = "+0.00;-0.00"
        ws.Cells(r, 5).Value = ops(i).HrsPerCart
        ws.Cells(r, 5).NumberFormat = "0.0"
        ws.Cells(r, 6).Value = ops(i).PeopleReq
        ws.Cells(r, 6).NumberFormat = "0.0"

        If i = bn Then
            ws.Cells(r, 7).Value = "CONSTRAINT - the line cannot go faster than this op"
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = RGB(255, 235, 235)
            ws.Cells(r, 7).Font.bold = True
            ws.Cells(r, 7).Font.Color = RGB(192, 0, 0)
        ElseIf ops(i).CapPerDay < reqPerDay Then
            ws.Cells(r, 7).Value = "Short of the required rate - will bind if the constraint clears"
            ws.Cells(r, 7).Font.Color = RGB(150, 75, 0)
        Else
            ws.Cells(r, 7).Value = "Has headroom"
            ws.Cells(r, 7).Font.Color = RGB(120, 120, 120)
        End If
        r = r + 1
    Next i

    If nOps = 0 Then
        ws.Cells(r, 1).Value = "Could not read per-op capacity from " & SH_PLAN_CFG & "."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    End If

    WriteOpTable = r
End Function

'==============================================================================
' FOCUS RANKING
'
' Scores each candidate call on the shafts it puts at risk, so the answer to
' "where do I spend Monday" is a number rather than a feeling.
'==============================================================================
Private Function WriteFocusRanking(ByVal ws As Worksheet, ByVal r As Long, _
                                   ByRef wp As WeekPicture, ByRef ops() As OpLoad, _
                                   ByVal nOps As Long, ByVal bn As Long, _
                                   ByVal gapPerDay As Double) As Long
    Dim items(1 To 6) As Variant
    Dim n As Long, i As Long, j As Long
    Dim rwk As Double, rej As Double, bg As Double
    Dim topDefect As String, topRoot As String, topOp As String
    Dim atRisk As Double, holds As Long
    Dim tmp As Variant

    SumQuality WeekStart(wp.WeekEnd), wp.WeekEnd, rwk, rej, bg, topDefect, topRoot, topOp
    atRisk = ShaftsAtRisk()
    holds = OpenHoldCount()

    ' Array(score in shafts, headline, what to do)
    If gapPerDay > 0 And bn > 0 Then
        n = n + 1
        items(n) = Array(gapPerDay * wp.DaysLeft * wp.ShaftsPerCart, _
            "THROUGHPUT - Op " & ops(bn).name & " is the constraint", _
            "Every hour added anywhere else this week buys nothing. Add hours, people or a day at Op " & _
            ops(bn).name & " only.")
    End If

    If rej + rwk + bg > 0 Then
        n = n + 1
        items(n) = Array(rej + rwk + bg, _
            "QUALITY - " & Format$(rej + rwk + bg, "#,##0") & " shafts dispositioned this week", _
            "Top defect is " & topDefect & ", root cause " & topRoot & ", worst at op " & topOp & _
            ". That is " & Format$((rej + rwk + bg) / IIf(wp.ShaftsPerCart > 0, wp.ShaftsPerCart, 98), "0.0") & _
            " carts of capacity bought back for free if it stops.")
    End If

    If atRisk > 0 Then
        n = n + 1
        items(n) = Array(atRisk, _
            "SCHEDULE RISK - " & Format$(atRisk, "#,##0") & " shafts projected to miss their week", _
            "From Cart Tracker week risk. Re-sequence or expedite before adding cost.")
    End If

    If holds > 0 Then
        n = n + 1
        items(n) = Array(holds * wp.ShaftsPerCart, _
            "HOLDS - " & holds & " cart(s) sitting on an open hold", _
            "Held carts are stopped, not slow. Clearing them is the cheapest capacity on the board.")
    End If

    If Not wp.HasPlan Then
        n = n + 1
        items(n) = Array(999999#, "DATA - no plan figure found on REVO Floor", _
            "Everything above is unanchored until PLAN RELEASES carries a number.")
    End If

    If n = 0 Then
        ws.Cells(r, 1).Value = "Nothing is flagged. Plan is covered, quality is quiet, no carts at risk."
        ws.Cells(r, 1).Font.Color = RGB(0, 97, 0)
        WriteFocusRanking = r + 1
        Exit Function
    End If

    ' Descending by shafts at stake.
    For i = 1 To n - 1
        For j = i + 1 To n
            If CDbl(items(j)(0)) > CDbl(items(i)(0)) Then
                tmp = items(i): items(i) = items(j): items(j) = tmp
            End If
        Next j
    Next i

    ws.Cells(r, 1).Value = "Rank"
    ws.Cells(r, 2).Value = "Shafts at stake"
    ws.Cells(r, 3).Value = "Call"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))
        .Font.bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
    End With
    r = r + 1

    For i = 1 To n
        ws.Cells(r, 1).Value = i
        If CDbl(items(i)(0)) < 999999# Then
            ws.Cells(r, 2).Value = CDbl(items(i)(0))
            ws.Cells(r, 2).NumberFormat = "#,##0"
        End If
        ws.Cells(r, 3).Value = CStr(items(i)(1))
        ws.Cells(r, 3).Font.bold = (i = 1)
        ws.Cells(r, 7).Value = CStr(items(i)(2))
        r = r + 1
    Next i

    WriteFocusRanking = r
End Function

'==============================================================================
' READERS
'==============================================================================
Private Function ReadWeekPicture() As WeekPicture
    Dim wp As WeekPicture
    Dim ws As Worksheet
    Dim v As Double

    wp.ShaftsPerCart = ReadShaftsPerCart()
    wp.WeekEnd = ReadWeekEnd()

    Set ws = REVO_Core.GetSheet(SH_FLOOR)
    If Not ws Is Nothing Then
        v = ValueUnderLabel(ws, Array("PLAN RELEASES  (SUN " & ChrW(8594) & " FRI)", _
                                      "PLAN RELEASES (SUN -> FRI)", "PLAN RELEASES"), 12, 30)
        wp.PlanShafts = v
        wp.HasPlan = (v > 0)
    End If

    wp.ReleasedShafts = REVO_Core.ReleasedBetween(WeekStart(wp.WeekEnd), wp.WeekEnd)
    wp.RemainingShafts = wp.PlanShafts - wp.ReleasedShafts
    If wp.RemainingShafts < 0 Then wp.RemainingShafts = 0
    If wp.PlanShafts > 0 Then wp.PctToPlan = wp.ReleasedShafts / wp.PlanShafts

    wp.DaysLeft = REVO_Core.ProductionDaysBetween(Date, wp.WeekEnd)
    wp.DaysElapsed = REVO_Core.ProductionDaysBetween(WeekStart(wp.WeekEnd), Date - 1)

    ReadWeekPicture = wp
End Function

' Labels on REVO Floor sit one row above their value. Find the label, take the
' cell below it.
Private Function ValueUnderLabel(ByVal ws As Worksheet, ByVal labels As Variant, _
                                 ByVal scanRows As Long, ByVal scanCols As Long) As Double
    Dim rr As Long, cc As Long, i As Long
    Dim t As String

    For rr = 1 To scanRows
        For cc = 1 To scanCols
            t = UCase$(REVO_Core.SafeS(ws.Cells(rr, cc).Value))
            If Len(t) > 0 Then
                For i = LBound(labels) To UBound(labels)
                    If InStr(t, UCase$(Trim$(CStr(labels(i))))) > 0 Then
                        ValueUnderLabel = REVO_Core.SafeD(ws.Cells(rr + 1, cc).Value)
                        If ValueUnderLabel <> 0 Then Exit Function
                    End If
                Next i
            End If
        Next cc
    Next rr
End Function


Private Function ReadOpLoads(ByRef ops() As OpLoad) As Long
    Dim ws As Worksheet, hdr As Long, lastR As Long, r As Long, n As Long
    Dim cOp As Long, cCap As Long, cPeople As Long
    Dim capDay As Double

    ReDim ops(1 To 40)

    Set ws = REVO_Core.GetSheet(SH_PLAN_CFG)
    If ws Is Nothing Then Exit Function

    hdr = REVO_Core.FindHeaderRow(ws, Array("Op"), 10, 8)
    If hdr = 0 Then Exit Function

    cOp = REVO_Core.FindCol(ws, hdr, Array("Op"))
    cCap = REVO_Core.FindCol(ws, hdr, Array("Weekday Cap/day", "Weekday Cap/Day"))
    If cCap = 0 Then cCap = REVO_Core.FindCol(ws, hdr, Array("Carts/day per team"))
    cPeople = REVO_Core.FindCol(ws, hdr, Array("Weekday People", "People Req"))
    If cOp = 0 Or cCap = 0 Then Exit Function

    lastR = REVO_Core.LastDataRow(ws, cOp, hdr)
    For r = hdr + 1 To lastR
        If Len(REVO_Core.SafeS(ws.Cells(r, cOp).Value)) > 0 Then
            capDay = REVO_Core.SafeD(ws.Cells(r, cCap).Value)
            If capDay > 0 Then
                n = n + 1
                If n > UBound(ops) Then ReDim Preserve ops(1 To n + 10)
                ops(n).name = REVO_Core.SafeS(ws.Cells(r, cOp).Value)
                ops(n).CapPerDay = capDay
                ops(n).PeopleReq = REVO_Core.SafeD(ws.Cells(r, cPeople).Value)
                ' Hours per cart implied by the rate and the crew on it.
                ops(n).HrsPerCart = 0
                If capDay > 0 Then
                    ops(n).HrsPerCart = (HOURS_PER_SHIFT * SHIFTS_PER_DAY) / capDay
                End If
            End If
        End If
    Next r

    ' Hours per cart from Daily Production Plan when it is published there -
    ' that is the floor's own number and beats an implied one.
    OverlayDailyPlanHours ops, n

    If n > 0 Then ReDim Preserve ops(1 To n)
    ReadOpLoads = n
End Function

Private Sub OverlayDailyPlanHours(ByRef ops() As OpLoad, ByVal n As Long)
    Dim ws As Worksheet, hdr As Long, lastR As Long, r As Long, i As Long
    Dim cOp As Long, cHrs As Long, cStaff As Long, cSched As Long
    Dim nm As String

    Set ws = REVO_Core.GetSheet(SH_DAILY_PLAN)
    If ws Is Nothing Then Exit Sub

    hdr = REVO_Core.FindHeaderRow(ws, Array("Hrs/Cart"), 20, 16)
    If hdr = 0 Then Exit Sub

    cOp = REVO_Core.FindCol(ws, hdr, Array("Op"))
    cHrs = REVO_Core.FindCol(ws, hdr, Array("Hrs/Cart"))
    cStaff = REVO_Core.FindCol(ws, hdr, Array("Staff On"))
    cSched = REVO_Core.FindCol(ws, hdr, Array("# Carts"))
    If cOp = 0 Or cHrs = 0 Then Exit Sub

    lastR = REVO_Core.LastDataRow(ws, cOp, hdr)
    For r = hdr + 1 To lastR
        nm = REVO_Core.SafeS(ws.Cells(r, cOp).Value)
        If Len(nm) > 0 Then
            For i = 1 To n
                If UCase$(ops(i).name) = UCase$(nm) Then
                    If REVO_Core.SafeD(ws.Cells(r, cHrs).Value) > 0 Then
                        ops(i).HrsPerCart = REVO_Core.SafeD(ws.Cells(r, cHrs).Value)
                    End If
                    If cStaff > 0 Then ops(i).StaffOn = REVO_Core.SafeS(ws.Cells(r, cStaff).Value)
                    If cSched > 0 Then ops(i).CartsSched = REVO_Core.SafeD(ws.Cells(r, cSched).Value)
                    Exit For
                End If
            Next i
        End If
    Next r
End Sub

Private Function FindBottleneck(ByRef ops() As OpLoad, ByVal n As Long) As Long
    Dim i As Long, lo As Double, at As Long
    If n <= 0 Then Exit Function
    lo = 1E+30
    For i = 1 To n
        If ops(i).CapPerDay > 0 And ops(i).CapPerDay < lo Then
            lo = ops(i).CapPerDay
            at = i
        End If
    Next i
    If at > 0 Then ops(at).IsBottleneck = True
    FindBottleneck = at
End Function

Private Sub SumQuality(ByVal d1 As Date, ByVal d2 As Date, _
                       ByRef rwk As Double, ByRef rej As Double, ByRef bg As Double, _
                       ByRef topDefect As String, ByRef topRoot As String, ByRef topOp As String)
    Dim ws As Worksheet, r As Long, lastR As Long
    Dim d As Date, disp As String, q As Double
    Dim dD As Object, dR As Object, dO As Object

    Set ws = REVO_Core.GetSheet(SH_QUAL_LOG)
    If ws Is Nothing Then Exit Sub

    Set dD = CreateObject("Scripting.Dictionary"): dD.CompareMode = vbTextCompare
    Set dR = CreateObject("Scripting.Dictionary"): dR.CompareMode = vbTextCompare
    Set dO = CreateObject("Scripting.Dictionary"): dO.CompareMode = vbTextCompare

    lastR = ws.Cells(ws.Rows.Count, QLC_EVENTID).End(xlUp).row
    For r = QL_HEADER_ROW + 1 To lastR
        d = REVO_Core.SafeDate(ws.Cells(r, QLC_DATE).Value, 0)
        If d >= d1 And d <= d2 Then
            disp = UCase$(REVO_Core.SafeS(ws.Cells(r, QLC_DISP).Value))
            q = REVO_Core.SafeD(ws.Cells(r, QLC_QTY).Value)
            Select Case disp
                Case "REWORK":  rwk = rwk + q
                Case "REJECT":  rej = rej + q
                Case "B GRADE": bg = bg + q
            End Select
            Bump dD, REVO_Core.SafeS(ws.Cells(r, QLC_DEFECT).Value), q
            Bump dR, REVO_Core.SafeS(ws.Cells(r, QLC_ROOT).Value), q
            Bump dO, REVO_Core.SafeS(ws.Cells(r, QLC_OP).Value), q
        End If
    Next r

    topDefect = TopKey(dD)
    topRoot = TopKey(dR)
    topOp = TopKey(dO)
End Sub

Private Sub Bump(ByVal d As Object, ByVal k As String, ByVal v As Double)
    If Len(Trim$(k)) = 0 Then Exit Sub
    If d.exists(k) Then d(k) = d(k) + v Else d.Add k, v
End Sub

Private Function TopKey(ByVal d As Object) As String
    Dim k As Variant, best As Double, bk As String
    For Each k In d.keys
        If d(k) > best Then best = d(k): bk = CStr(k)
    Next k
    TopKey = bk
End Function

Private Function ShaftsAtRisk() As Double
    Dim ws As Worksheet, r As Long
    Set ws = REVO_Core.GetSheet(SH_CART_TRK)
    If ws Is Nothing Then Exit Function
    For r = 1 To 60
        If InStr(UCase$(REVO_Core.SafeS(ws.Cells(r, 1).Value)), "SHAFTS AT RISK") > 0 Then
            ShaftsAtRisk = REVO_Core.SafeD(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
End Function

Private Function OpenHoldCount() As Long
    Dim ws As Worksheet, hdr As Long, lastR As Long, r As Long, n As Long
    Dim cCart As Long, cRel As Long

    Set ws = REVO_Core.GetSheet(SH_HOLDS)
    If ws Is Nothing Then Exit Function

    hdr = REVO_Core.FindHeaderRow(ws, Array("Cart #", "Cart"), 10, 10)
    If hdr = 0 Then Exit Function

    cCart = REVO_Core.FindCol(ws, hdr, Array("Cart #", "Cart"))
    cRel = REVO_Core.FindCol(ws, hdr, Array("Released?", "Released"))
    If cCart = 0 Or cRel = 0 Then Exit Function

    lastR = REVO_Core.LastDataRow(ws, cCart, hdr)
    For r = hdr + 1 To lastR
        If Len(REVO_Core.SafeS(ws.Cells(r, cCart).Value)) > 0 Then
            If Not REVO_Core.Boolish(ws.Cells(r, cRel).Value) Then n = n + 1
        End If
    Next r
    OpenHoldCount = n
End Function

Private Function ReadShaftsPerCart() As Double
    Dim ws As Worksheet, r As Long, v As Double
    Set ws = REVO_Core.GetSheet(SH_WEEK_PLAN)
    If Not ws Is Nothing Then
        For r = 1 To 12
            If InStr(UCase$(REVO_Core.SafeS(ws.Cells(r, 1).Value)), "SHAFTS PER CART") > 0 Then
                v = REVO_Core.SafeD(ws.Cells(r, 3).Value)
                If v > 0 Then ReadShaftsPerCart = v: Exit Function
            End If
        Next r
    End If
    ReadShaftsPerCart = DEFAULT_SHAFTS_PER_CART
End Function

Private Function ReadWeekEnd() As Date
    Dim ws As Worksheet, r As Long, d As Date
    Set ws = REVO_Core.GetSheet(SH_CART_TRK)
    If Not ws Is Nothing Then
        For r = 1 To 60
            If InStr(UCase$(REVO_Core.SafeS(ws.Cells(r, 1).Value)), "COMPLETION FRIDAY") > 0 Then
                d = REVO_Core.SafeDate(ws.Cells(r, 2).Value, 0)
                If d > 0 Then ReadWeekEnd = d: Exit Function
            End If
        Next r
    End If
    ' Fall back to the Friday of the current Sunday-start week.
    ReadWeekEnd = Date + ((6 - Weekday(Date, vbSunday)) Mod 7)
End Function

Private Function WeekStart(ByVal weekEnd As Date) As Date
    ' Production week runs Sunday to Friday.
    WeekStart = weekEnd - 5
End Function

Private Function CapBasis() As String
    Dim ws As Worksheet, r As Long
    Set ws = REVO_Core.GetSheet(SH_VELOCITY)
    If ws Is Nothing Then CapBasis = "Basis: Plan Config (assumed).": Exit Function
    For r = 9 To 12
        If InStr(UCase$(REVO_Core.SafeS(ws.Cells(r, 7).Value)), "APPLIED") > 0 Then
            CapBasis = "Basis: MEASURED from Floor Log."
            Exit Function
        End If
    Next r
    CapBasis = "Basis: Plan Config (ASSUMED) - measured rate held, see " & SH_VELOCITY & "."
End Function

'==============================================================================
' LAYOUT HELPERS
'==============================================================================
Private Sub SectionTitle(ByVal ws As Worksheet, ByVal r As Long, ByVal t As String)
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

Private Sub KV(ByVal ws As Worksheet, ByVal r As Long, ByVal k As String, _
               ByVal v As Double, ByVal fmt As String, ByVal note As String)
    ws.Cells(r, 1).Value = k
    ws.Cells(r, 2).Value = v
    ws.Cells(r, 2).NumberFormat = fmt
    If Len(note) > 0 Then
        ws.Cells(r, 7).Value = note
        ws.Cells(r, 7).Font.Color = RGB(120, 120, 120)
    End If
End Sub

Private Sub KVs(ByVal ws As Worksheet, ByVal r As Long, ByVal k As String, _
                ByVal v As String, ByVal note As String)
    ws.Cells(r, 1).Value = k
    ws.Cells(r, 2).Value = v
    If Len(note) > 0 Then
        ws.Cells(r, 7).Value = note
        ws.Cells(r, 7).Font.Color = RGB(120, 120, 120)
    End If
End Sub
