Option Explicit

'=========================================================
' FULLY WIRED: Reject Dashboard Refresh
' - Inputs: C2 (start), E2 (end), G2 (lane), stamps I2
' - Tables (Bucket + Daily Trend) use dtStart..dtEnd (user-selected)
' - KPI Blocks are ALWAYS:
'     WTD = Monday..dtEnd
'     MTD = 1st of month..dtEnd
'     YTD = Jan 1..dtEnd
' - KPI Blocks write to named cells:
'     WTD_Released, WTD_Reject, WTD_RejectPct, WTD_Rework, WTD_ReworkPct, WTD_BGrade
'     MTD_*, YTD_*
' - Label lines:
'     B6 = WTD label
'     B13 = MTD label
'     B18 = YTD label
'=========================================================

'-------------------------
' Dashboard targets
'-------------------------
Private Const SH_DASH As String = "Reject_Dashboard"
Private Const LO_BUCKET As String = "tbl_BucketSummary"
Private Const LO_TREND  As String = "tbl_DailyTrend"

' Inputs
Private Const ADDR_START As String = "C2"
Private Const ADDR_END   As String = "E2"
Private Const ADDR_LANE  As String = "G2"
Private Const ADDR_REFRESH As String = "I2"

' Existing Top KPI named ranges (keep working)
Private Const NM_RELEASED      As String = "CELL_RELEASED"
Private Const NM_REJECT        As String = "CELL_REJECT"
Private Const NM_REJECT_PERC   As String = "CELL_REJECTPERC"
Private Const NM_REWORK        As String = "CELL_REWORK"
Private Const NM_REWORK_PERC   As String = "CELL_REWORKPERC"
Private Const NM_BGRADE        As String = "CELL_BGRADE"
Private Const NM_REPORTEDDATES As String = "CELL_REPORTEDDATES"
Private Const NM_WTD_LANES         As String = "CELL_LANES"
Private Const NM_MTD_LANES As String = "MTD_Lanes"
Private Const NM_YTD_LANES As String = "YTD_Lanes"

' B-Grade % cells (Column T)
Private Const ADDR_WTD_BG_PCT As String = "T6"
Private Const ADDR_MTD_BG_PCT As String = "T11"
Private Const ADDR_YTD_BG_PCT As String = "T16"


' KPI block named ranges
Private Const NM_WTD_REL As String = "WTD_Released"
Private Const NM_WTD_REJ As String = "WTD_Reject"
Private Const NM_WTD_REJP As String = "WTD_RejectPct"
Private Const NM_WTD_REW As String = "WTD_Rework"
Private Const NM_WTD_REWP As String = "WTD_ReworkPct"
Private Const NM_WTD_BG  As String = "WTD_BGrade"

Private Const NM_MTD_REL As String = "MTD_Released"
Private Const NM_MTD_REJ As String = "MTD_Reject"
Private Const NM_MTD_REJP As String = "MTD_RejectPct"
Private Const NM_MTD_REW As String = "MTD_Rework"
Private Const NM_MTD_REWP As String = "MTD_ReworkPct"
Private Const NM_MTD_BG  As String = "MTD_BGrade"

Private Const NM_YTD_REL As String = "YTD_Released"
Private Const NM_YTD_REJ As String = "YTD_Reject"
Private Const NM_YTD_REJP As String = "YTD_RejectPct"
Private Const NM_YTD_REW As String = "YTD_Rework"
Private Const NM_YTD_REWP As String = "YTD_ReworkPct"
Private Const NM_YTD_BG  As String = "YTD_BGrade"

' Label cells (lane | range)
Private Const ADDR_WTD_LABEL As String = "B8"
Private Const ADDR_MTD_LABEL As String = "B13"
Private Const ADDR_YTD_LABEL As String = "B18"

' Stage marker
Private gStage As String

'=========================================================
' PUBLIC ENTRYPOINT
'=========================================================
Public Sub RefreshRejectDashboard_Full()

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsD As Worksheet, wsL As Worksheet

    Dim dtStart As Date, dtEnd As Date, laneFilter As String

    Dim colDate As Long, colSKU As Long, colRel As Long, colRej As Long, colRew As Long, colBG As Long
    Dim lastRow As Long, lastCol As Long
    Dim arr As Variant

    ' Selected range totals (for top KPIs + tables)
    Dim selRel As Double, selRej As Double, selRew As Double, selBG As Double
    Dim selRel_BVP As Double, selRej_BVP As Double
    Dim selRel_WVP As Double, selRej_WVP As Double

    Dim dBuckets As Object   ' key = "LANE|BUCKET" -> Make4(rel, rej, rew, bg)
    Dim dTrend As Object     ' key = CLng(date) -> Make4(rel, rej, rew, bg)
    Dim ordered As Variant

    ' KPI block ranges (based on dtEnd)
    Dim wtdStart As Date, mtdStart As Date, ytdStart As Date

    ' KPI block totals
    Dim wRel As Double, wRej As Double, wRew As Double, wBG As Double
    Dim mRel As Double, mRej As Double, mRew As Double, mBG As Double
    Dim yRel As Double, yRej As Double, yRew As Double, yBG As Double

    ordered = OrderedBuckets()

    gStage = "INIT"
    On Error GoTo CleanFail

    Set wsD = GetSheetSafe(wb, SH_DASH)
    If wsD Is Nothing Then Err.Raise vbObjectError + 10, , "Missing sheet: " & SH_DASH

    gStage = "FindLogSheet"
    Set wsL = FindLogSheet(wb)
    If wsL Is Nothing Then Err.Raise vbObjectError + 11, , "Could not find the Release Log sheet (common names)."

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    gStage = "ReadInputs"
    If Not ReadDashInputs(wsD, dtStart, dtEnd, laneFilter) Then GoTo CleanExit

    ' KPI blocks are anchored to dtEnd
    wtdStart = dtEnd - Weekday(dtEnd, vbMonday) + 1
    mtdStart = DateSerial(Year(dtEnd), Month(dtEnd), 1)
    ytdStart = DateSerial(Year(dtEnd), 1, 1)

    gStage = "FindLogColumns"
    FindLogColumns wsL, colDate, colSKU, colRel, colRej, colRew, colBG

    gStage = "LoadLogArray"
    lastRow = wsL.Cells(wsL.Rows.Count, colDate).End(xlUp).row
    If lastRow < 2 Then
        arr = Empty
    Else
        lastCol = Max6(colDate, colSKU, colRel, colRej, colRew, colBG)
        arr = wsL.Range(wsL.Cells(2, 1), wsL.Cells(lastRow, lastCol)).Value2
    End If

    gStage = "InitDictionaries"
    Set dBuckets = CreateObject("Scripting.Dictionary"): dBuckets.CompareMode = vbTextCompare
    Set dTrend = CreateObject("Scripting.Dictionary"): dTrend.CompareMode = vbTextCompare

    gStage = "PrimeBuckets"
    PrimeLaneBuckets dBuckets, Array("BVP", "WVP"), ordered

    gStage = "Aggregate_SelectedRange"
    If Not IsEmpty(arr) Then
        AggregateFromArray_SelectedRange arr, colDate, colSKU, colRel, colRej, colRew, colBG, _
                                         dtStart, dtEnd, laneFilter, _
                                         selRel, selRej, selRew, selBG, _
                                         selRel_BVP, selRej_BVP, selRel_WVP, selRej_WVP, _
                                         dBuckets, dTrend
    End If

    gStage = "Aggregate_KPIBlocks"
    If Not IsEmpty(arr) Then
        AggregateKPIBlocks_OnePass arr, colDate, colSKU, colRel, colRej, colRew, colBG, _
                                   wtdStart, mtdStart, ytdStart, dtEnd, laneFilter, _
                                   wRel, wRej, wRew, wBG, _
                                   mRel, mRej, mRew, mBG, _
                                   yRel, yRej, yRew, yBG
    End If

    gStage = "Render"
    RenderDashboard_AllWired wb, wsD, _
                             dtStart, dtEnd, laneFilter, _
                             selRel, selRej, selRew, selBG, _
                             selRel_BVP, selRej_BVP, selRel_WVP, selRej_WVP, _
                             dBuckets, dTrend, ordered, _
                             wtdStart, mtdStart, ytdStart, _
                             wRel, wRej, wRew, wBG, _
                             mRel, mRej, mRew, mBG, _
                             yRel, yRej, yRew, yBG

    gStage = "ApplyCF"
    ApplyRejectDashboardCF wb, wsD

    gStage = "StampRefresh"
    wsD.Range(ADDR_REFRESH).NumberFormat = "m/d/yyyy h:mm AM/PM"
    wsD.Range(ADDR_REFRESH).Value = Now

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "RefreshRejectDashboard_Full failed (Stage: " & gStage & "): Err " & Err.Number & " - " & Err.Description, vbExclamation

End Sub

'=========================================================
' INPUTS
'=========================================================
Private Function ReadDashInputs(ByVal ws As Worksheet, ByRef dtStart As Date, ByRef dtEnd As Date, ByRef laneFilter As String) As Boolean
    ReadDashInputs = False

    If Not IsDate(ws.Range(ADDR_START).Value) Then
        MsgBox "Start Date (C2) is not a valid date.", vbExclamation
        Exit Function
    End If
    If Not IsDate(ws.Range(ADDR_END).Value) Then
        MsgBox "End Date (E2) is not a valid date.", vbExclamation
        Exit Function
    End If

    dtStart = DateValue(CDate(ws.Range(ADDR_START).Value))
    dtEnd = DateValue(CDate(ws.Range(ADDR_END).Value))
    If dtEnd < dtStart Then
        MsgBox "End Date cannot be before Start Date.", vbExclamation
        Exit Function
    End If

    laneFilter = UCase$(Trim$(CStr(ws.Range(ADDR_LANE).Value)))
    If laneFilter = "" Then laneFilter = "ALL"
    If laneFilter <> "ALL" And laneFilter <> "BVP" And laneFilter <> "WVP" Then
        MsgBox "Lane (G2) must be ALL, BVP, or WVP.", vbExclamation
        Exit Function
    End If

    ReadDashInputs = True
End Function

Private Function LaneLabel(ByVal laneFilter As String) As String
    Dim s As String
    s = UCase$(Trim$(laneFilter))
    If s = "" Or s = "ALL" Then LaneLabel = "All" Else LaneLabel = s
End Function

'=========================================================
' FIND LOG SHEET + COLUMNS
'=========================================================
Public Function FindLogSheet(ByVal wb As Workbook) As Worksheet
    Dim names As Variant, n As Variant
    names = Array("Release_Log", "Release Log", "ReleaseLog", "Release", "Reject Log", "Releases", "ReleaseLog_Data", "Release_Log_Data")

    For Each n In names
        Set FindLogSheet = GetSheetSafe(wb, CStr(n))
        If Not FindLogSheet Is Nothing Then Exit Function
    Next n
End Function

Public Sub FindLogColumns(ByVal wsL As Worksheet, _
                           ByRef colDate As Long, ByRef colSKU As Long, ByRef colRel As Long, _
                           ByRef colRej As Long, ByRef colRew As Long, ByRef colBG As Long)

    Dim hdrRow As Long: hdrRow = 1

    colDate = FindHeaderCol(wsL, hdrRow, Array("Date", "Release Date", "Completed Date"))
    colSKU = FindHeaderCol(wsL, hdrRow, Array("SKU", "Item", "SKU / Item", "SKU Item", "Product", "Part"))
    colRel = FindHeaderCol(wsL, hdrRow, Array("Qty Released", "QtyRelease", "Quantity Released", "Released", "Release Qty", "Qty Released "))
    colRew = FindHeaderCol(wsL, hdrRow, Array("Qty Rework", "QtyRework", "Quantity Rework", "Rework", "Rework Qty", "Qty Reworked", "Qty Reworked "))
    colBG = FindHeaderCol(wsL, hdrRow, Array("Qty B Grade", "Qty B-Grade", "Qty BGrade", "Quantity B Grade", "B Grade", "B-Grade", "BGrade", "Bgrade"))
    colRej = FindHeaderCol(wsL, hdrRow, Array("Qty Reject", "QtyReject", "Quantity Reject", "Reject", "Rejected", "Reject Qty", "Qty Rejected"))

    If colDate = 0 Or colSKU = 0 Or colRel = 0 Or colRej = 0 Or colRew = 0 Or colBG = 0 Then
        Err.Raise vbObjectError + 200, , _
            "Could not detect required columns on log sheet '" & wsL.name & "'." & vbCrLf & _
            "Found: Date=" & colDate & ", SKU=" & colSKU & ", Released=" & colRel & ", Reject=" & colRej & ", Rework=" & colRew & ", B Grade=" & colBG & vbCrLf & _
            "Check header spelling on row 1."
    End If
End Sub

Public Function FindHeaderCol(ByVal ws As Worksheet, ByVal hdrRow As Long, ByVal aliases As Variant) As Long
    Dim lastCol As Long: lastCol = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, a As Variant
    Dim h As String, want As String

    For c = 1 To lastCol
        h = CStr(ws.Cells(hdrRow, c).Value2)
        h = UCase$(Trim$(Replace(Replace(h, vbCr, " "), vbLf, " ")))
        h = WorksheetFunction.Trim(h)

        If h <> "" Then
            For Each a In aliases
                want = UCase$(Trim$(CStr(a)))
                If h = want Then
                    FindHeaderCol = c
                    Exit Function
                End If
            Next a
        End If
    Next c
End Function

'=========================================================
' AGGREGATION
'=========================================================
Private Sub PrimeLaneBuckets(ByVal d As Object, ByVal lanes As Variant, ByVal buckets As Variant)
    Dim ln As Variant, b As Variant
    For Each ln In lanes
        For Each b In buckets
            d(CStr(ln) & "|" & CStr(b)) = Make4(0, 0, 0, 0)
        Next b
    Next ln
End Sub

' Selected-range aggregation (tables + top KPIs + lane reject% line)
Private Sub AggregateFromArray_SelectedRange( _
    ByVal arr As Variant, _
    ByVal colDate As Long, ByVal colSKU As Long, ByVal colRel As Long, ByVal colRej As Long, ByVal colRew As Long, ByVal colBG As Long, _
    ByVal dtStart As Date, ByVal dtEnd As Date, ByVal laneFilter As String, _
    ByRef totRel As Double, ByRef totRej As Double, ByRef totRew As Double, ByRef totBG As Double, _
    ByRef totRel_BVP As Double, ByRef totRej_BVP As Double, _
    ByRef totRel_WVP As Double, ByRef totRej_WVP As Double, _
    ByVal dBuckets As Object, ByVal dTrend As Object _
)
    Dim r As Long, vDate As Variant, dd As Date
    Dim sku As String, lane As String, bucket As String
    Dim qRel As Double, qRej As Double, qRew As Double, qBG As Double
    Dim keyLB As String, keyDay As String
    Dim aTmp As Variant

    For r = 1 To UBound(arr, 1)

        vDate = arr(r, colDate)
        If Not GetRowDate(vDate, dd) Then GoTo NextR
        If dd < dtStart Or dd > dtEnd Then GoTo NextR

        sku = Trim$(CStr(arr(r, colSKU)))
        If sku = "" Then GoTo NextR

        lane = LaneFromSKU(sku)
        bucket = RejectBucketFromSKU(sku)
        If bucket = "" Then GoTo NextR

        qRel = NzD(arr(r, colRel))
        qRej = NzD(arr(r, colRej))
        qRew = NzD(arr(r, colRew))
        qBG = NzD(arr(r, colBG))

        ' Totals obey lane filter
        If laneFilter = "ALL" Or lane = laneFilter Then
            totRel = totRel + qRel
            totRej = totRej + qRej
            totRew = totRew + qRew
            totBG = totBG + qBG

            keyDay = CStr(CLng(dd))
            If Not dTrend.exists(keyDay) Then dTrend(keyDay) = Make4(0, 0, 0, 0)

            aTmp = dTrend(keyDay)
            AddTo4 aTmp, qRel, qRej, qRew, qBG
            dTrend(keyDay) = aTmp
        End If

        ' Lane totals (always) for lane line
        If lane = "BVP" Then
            totRel_BVP = totRel_BVP + qRel
            totRej_BVP = totRej_BVP + qRej
        ElseIf lane = "WVP" Then
            totRel_WVP = totRel_WVP + qRel
            totRej_WVP = totRej_WVP + qRej
        End If

        ' Bucket Summary by Lane (always BVP/WVP)
        If lane = "BVP" Or lane = "WVP" Then
            keyLB = lane & "|" & bucket
            If Not dBuckets.exists(keyLB) Then dBuckets(keyLB) = Make4(0, 0, 0, 0)

            aTmp = dBuckets(keyLB)
            AddTo4 aTmp, qRel, qRej, qRew, qBG
            dBuckets(keyLB) = aTmp
        End If

NextR:
    Next r
End Sub

' One-pass KPI block aggregator (WTD/MTD/YTD) anchored to dtEnd
Private Sub AggregateKPIBlocks_OnePass( _
    ByVal arr As Variant, _
    ByVal colDate As Long, ByVal colSKU As Long, ByVal colRel As Long, ByVal colRej As Long, ByVal colRew As Long, ByVal colBG As Long, _
    ByVal wtdStart As Date, ByVal mtdStart As Date, ByVal ytdStart As Date, ByVal dtEnd As Date, _
    ByVal laneFilter As String, _
    ByRef wRel As Double, ByRef wRej As Double, ByRef wRew As Double, ByRef wBG As Double, _
    ByRef mRel As Double, ByRef mRej As Double, ByRef mRew As Double, ByRef mBG As Double, _
    ByRef yRel As Double, ByRef yRej As Double, ByRef yRew As Double, ByRef yBG As Double _
)
    Dim r As Long, vDate As Variant, dd As Date
    Dim sku As String, lane As String
    Dim qRel As Double, qRej As Double, qRew As Double, qBG As Double

    For r = 1 To UBound(arr, 1)

        vDate = arr(r, colDate)
        If Not GetRowDate(vDate, dd) Then GoTo NextR
        If dd > dtEnd Then GoTo NextR  ' guard

        sku = Trim$(CStr(arr(r, colSKU)))
        If sku = "" Then GoTo NextR

        lane = LaneFromSKU(sku)

        ' KPI blocks obey lane filter
        If Not (laneFilter = "ALL" Or lane = laneFilter) Then GoTo NextR

        qRel = NzD(arr(r, colRel))
        qRej = NzD(arr(r, colRej))
        qRew = NzD(arr(r, colRew))
        qBG = NzD(arr(r, colBG))

        If dd >= wtdStart Then
            wRel = wRel + qRel: wRej = wRej + qRej: wRew = wRew + qRew: wBG = wBG + qBG
        End If
        If dd >= mtdStart Then
            mRel = mRel + qRel: mRej = mRej + qRej: mRew = mRew + qRew: mBG = mBG + qBG
        End If
        If dd >= ytdStart Then
            yRel = yRel + qRel: yRej = yRej + qRej: yRew = yRew + qRew: yBG = yBG + qBG
        End If

NextR:
    Next r
End Sub

Private Function GetRowDate(ByVal v As Variant, ByRef outDate As Date) As Boolean
    On Error GoTo EH
    If IsNumeric(v) Then
        outDate = DateValue(CDate(CDbl(v)))
        GetRowDate = True
        Exit Function
    End If
    If IsDate(v) Then
        outDate = DateValue(CDate(v))
        GetRowDate = True
        Exit Function
    End If
EH:
    GetRowDate = False
End Function

'=========================================================
' RENDER (FULLY WIRED)
'=========================================================
Private Sub RenderDashboard_AllWired( _
    ByVal wb As Workbook, ByVal wsD As Worksheet, _
    ByVal dtStart As Date, ByVal dtEnd As Date, ByVal laneFilter As String, _
    ByVal selRel As Double, ByVal selRej As Double, ByVal selRew As Double, ByVal selBG As Double, _
    ByVal selRel_BVP As Double, ByVal selRej_BVP As Double, _
    ByVal selRel_WVP As Double, ByVal selRej_WVP As Double, _
    ByVal dBuckets As Object, ByVal dTrend As Object, _
    ByVal OrderedBuckets As Variant, _
    ByVal wtdStart As Date, ByVal mtdStart As Date, ByVal ytdStart As Date, _
    ByVal wRel As Double, ByVal wRej As Double, ByVal wRew As Double, ByVal wBG As Double, _
    ByVal mRel As Double, ByVal mRej As Double, ByVal mRew As Double, ByVal mBG As Double, _
    ByVal yRel As Double, ByVal yRej As Double, ByVal yRew As Double, ByVal yBG As Double _
)
    '----------------------------
    ' Top KPIs (selected range)
    '----------------------------
    PutNamed wb, NM_RELEASED, selRel
    PutNamed wb, NM_REJECT, selRej
    PutNamed wb, NM_REWORK, selRew
    PutNamed wb, NM_BGRADE, selBG
    PutNamed wb, NM_REJECT_PERC, SafePct(selRej, selRel)
    PutNamed wb, NM_REWORK_PERC, SafePct(selRew, selRel)
    PutNamed wb, NM_REPORTEDDATES, LaneLabel(laneFilter) & " | " & Format$(dtStart, "m/d/yyyy") & " - " & Format$(dtEnd, "m/d/yyyy")

    Dim bvpPct As Double, wvpPct As Double
    bvpPct = SafePct(selRej_BVP, selRel_BVP)
    wvpPct = SafePct(selRej_WVP, selRel_WVP)


    '----------------------------
    ' KPI Blocks (WTD/MTD/YTD)
    '----------------------------
    PutNamed wb, NM_WTD_REL, wRel
    PutNamed wb, NM_WTD_REJ, wRej
    PutNamed wb, NM_WTD_REJP, SafePct(wRej, wRel)
    PutNamed wb, NM_WTD_REW, wRew
    PutNamed wb, NM_WTD_REWP, SafePct(wRew, wRel)
    PutNamed wb, NM_WTD_BG, wBG

    PutNamed wb, NM_MTD_REL, mRel
    PutNamed wb, NM_MTD_REJ, mRej
    PutNamed wb, NM_MTD_REJP, SafePct(mRej, mRel)
    PutNamed wb, NM_MTD_REW, mRew
    PutNamed wb, NM_MTD_REWP, SafePct(mRew, mRel)
    PutNamed wb, NM_MTD_BG, mBG

    PutNamed wb, NM_YTD_REL, yRel
    PutNamed wb, NM_YTD_REJ, yRej
    PutNamed wb, NM_YTD_REJP, SafePct(yRej, yRel)
    PutNamed wb, NM_YTD_REW, yRew
    PutNamed wb, NM_YTD_REWP, SafePct(yRew, yRel)
    PutNamed wb, NM_YTD_BG, yBG
    
        '----------------------------
    ' B-Grade % cells (Column T)
    '----------------------------
    wsD.Range(ADDR_WTD_BG_PCT).Value = SafePct(wBG, wRel)
    wsD.Range(ADDR_MTD_BG_PCT).Value = SafePct(mBG, mRel)
    wsD.Range(ADDR_YTD_BG_PCT).Value = SafePct(yBG, yRel)

    wsD.Range(ADDR_WTD_BG_PCT).NumberFormat = "0.00%"
    wsD.Range(ADDR_MTD_BG_PCT).NumberFormat = "0.00%"
    wsD.Range(ADDR_YTD_BG_PCT).NumberFormat = "0.00%"

    ' Labels in cells (per your layout)
    Dim laneTxt As String: laneTxt = LaneLabel(laneFilter)
    wsD.Range(ADDR_WTD_LABEL).NumberFormat = "@"
    wsD.Range(ADDR_MTD_LABEL).NumberFormat = "@"
    wsD.Range(ADDR_YTD_LABEL).NumberFormat = "@"

    wsD.Range(ADDR_WTD_LABEL).Value = laneTxt & " | " & Format$(wtdStart, "m/d/yyyy") & " - " & Format$(dtEnd, "m/d/yyyy")
    wsD.Range(ADDR_MTD_LABEL).Value = laneTxt & " | " & Format$(mtdStart, "m/d/yyyy") & " - " & Format$(dtEnd, "m/d/yyyy")
    wsD.Range(ADDR_YTD_LABEL).Value = laneTxt & " | " & Format$(ytdStart, "m/d/yyyy") & " - " & Format$(dtEnd, "m/d/yyyy")

    ' Number formats for named KPIs
    ApplyNamedFormats wb

    ' Tables (selected range)
    RenderBucketSummary_ByLane wsD, LO_BUCKET, Array("BVP", "WVP"), OrderedBuckets, dBuckets
    RenderDailyTrend wsD, LO_TREND, dtStart, dtEnd, dTrend
End Sub

'=========================================================
' TABLE RENDERERS
'=========================================================
Private Sub RenderBucketSummary_ByLane(ByVal ws As Worksheet, ByVal loName As String, ByVal lanes As Variant, ByVal OrderedBuckets As Variant, ByVal dBuckets As Object)
    Dim lo As ListObject
    Set lo = GetListObjectSafe(ws, loName)
    If lo Is Nothing Then Err.Raise vbObjectError + 310, , "Missing table: " & loName

    Dim hasLaneCol As Boolean
    hasLaneCol = HasColumn(lo, "Lane")

    Dim nBuckets As Long: nBuckets = UBound(OrderedBuckets) - LBound(OrderedBuckets) + 1
    Dim nLanes As Long: nLanes = UBound(lanes) - LBound(lanes) + 1
    Dim nRows As Long: nRows = nBuckets * nLanes

    EnsureTableRows lo, nRows

    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To lo.ListColumns.Count)

    Dim r As Long: r = 0
    Dim iLane As Long, iB As Long
    Dim lane As String, bucket As String, keyLB As String, a4 As Variant

    For iLane = LBound(lanes) To UBound(lanes)
        lane = CStr(lanes(iLane))
        For iB = LBound(OrderedBuckets) To UBound(OrderedBuckets)
            bucket = CStr(OrderedBuckets(iB))
            keyLB = lane & "|" & bucket

            If dBuckets.exists(keyLB) Then
                a4 = dBuckets(keyLB)
            Else
                a4 = Make4(0, 0, 0, 0)
            End If

            r = r + 1

            If hasLaneCol Then
                outArr(r, ColIndex(lo, "Lane")) = lane
                outArr(r, ColIndex(lo, "Bucket")) = bucket
            Else
                outArr(r, ColIndex(lo, "Bucket")) = lane & " | " & bucket
            End If

            outArr(r, ColIndex(lo, "Released")) = CDbl(a4(0))
            outArr(r, ColIndex(lo, "Reject")) = CDbl(a4(1))
            outArr(r, ColIndex(lo, "Reject %")) = SafePct(CDbl(a4(1)), CDbl(a4(0)))
            outArr(r, ColIndex(lo, "Rework")) = CDbl(a4(2))
            outArr(r, ColIndex(lo, "Rework %")) = SafePct(CDbl(a4(2)), CDbl(a4(0)))
            If HasColumn(lo, "B-Grade") Then outArr(r, ColIndex(lo, "B-Grade")) = CDbl(a4(3))
        Next iB
    Next iLane

    lo.DataBodyRange.Value = outArr

    On Error Resume Next
    lo.ListColumns("Reject %").DataBodyRange.NumberFormat = "0.0%"
    lo.ListColumns("Rework %").DataBodyRange.NumberFormat = "0.0%"
    lo.ListColumns("Released").DataBodyRange.NumberFormat = "0"
    lo.ListColumns("Reject").DataBodyRange.NumberFormat = "0"
    lo.ListColumns("Rework").DataBodyRange.NumberFormat = "0"
    If HasColumn(lo, "B-Grade") Then lo.ListColumns("B-Grade").DataBodyRange.NumberFormat = "0"
    On Error GoTo 0
End Sub

Private Sub RenderDailyTrend(ByVal ws As Worksheet, ByVal loName As String, ByVal dtStart As Date, ByVal dtEnd As Date, ByVal dTrend As Object)
    Dim lo As ListObject
    Set lo = GetListObjectSafe(ws, loName)
    If lo Is Nothing Then Err.Raise vbObjectError + 320, , "Missing table: " & loName

    Dim n As Long: n = DateDiff("d", dtStart, dtEnd) + 1
    If n < 1 Then Exit Sub

    EnsureTableRows lo, n

    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To lo.ListColumns.Count)

    Dim i As Long, dd As Date, keyDay As String, a4 As Variant
    For i = 1 To n
        dd = DateAdd("d", i - 1, dtStart)
        keyDay = CStr(CLng(dd))

        If dTrend.exists(keyDay) Then
            a4 = dTrend(keyDay)
        Else
            a4 = Make4(0, 0, 0, 0)
        End If

        outArr(i, ColIndex(lo, "Date")) = dd
        outArr(i, ColIndex(lo, "Released")) = CDbl(a4(0))
        outArr(i, ColIndex(lo, "Reject")) = CDbl(a4(1))
        outArr(i, ColIndex(lo, "Reject %")) = SafePct(CDbl(a4(1)), CDbl(a4(0)))
        outArr(i, ColIndex(lo, "Rework")) = CDbl(a4(2))
        outArr(i, ColIndex(lo, "Rework %")) = SafePct(CDbl(a4(2)), CDbl(a4(0)))
        If HasColumn(lo, "B-Grade") Then outArr(i, ColIndex(lo, "B-Grade")) = CDbl(a4(3))
    Next i

    lo.DataBodyRange.Value = outArr

    On Error Resume Next
    lo.ListColumns("Date").DataBodyRange.NumberFormat = "m/d/yyyy"
    lo.ListColumns("Reject %").DataBodyRange.NumberFormat = "0.0%"
    lo.ListColumns("Rework %").DataBodyRange.NumberFormat = "0.0%"
    lo.ListColumns("Released").DataBodyRange.NumberFormat = "0"
    lo.ListColumns("Reject").DataBodyRange.NumberFormat = "0"
    lo.ListColumns("Rework").DataBodyRange.NumberFormat = "0"
    If HasColumn(lo, "B-Grade") Then lo.ListColumns("B-Grade").DataBodyRange.NumberFormat = "0"
    On Error GoTo 0
End Sub

'=========================================================
' CONDITIONAL FORMATTING (Reject % only)
'=========================================================
Private Sub ApplyRejectDashboardCF(ByVal wb As Workbook, ByVal ws As Worksheet)

    ' KPI blocks reject %
    ApplyNamedRejectPctCF wb, NM_WTD_REJP
    ApplyNamedRejectPctCF wb, NM_MTD_REJP
    ApplyNamedRejectPctCF wb, NM_YTD_REJP

    ' Optional: top KPI reject % too
    ApplyNamedRejectPctCF wb, NM_REJECT_PERC

    ' Bucket Summary Reject % column
    Dim lo As ListObject
    Set lo = GetListObjectSafe(ws, LO_BUCKET)
    If Not lo Is Nothing Then ApplyRejectPctToListObject lo, "Reject %"

    ' Daily Trend Reject % column
    Set lo = GetListObjectSafe(ws, LO_TREND)
    If Not lo Is Nothing Then ApplyRejectPctToListObject lo, "Reject %"
End Sub

Private Sub ApplyNamedRejectPctCF(ByVal wb As Workbook, ByVal nm As String)
    Dim r As Range
    Set r = ResolveNameRangeWB(wb, nm)
    If Not r Is Nothing Then
        r.NumberFormat = "0.00%"
        ApplyRejectPctThresholdCF r
    End If
End Sub

Private Sub ApplyRejectPctToListObject(ByVal lo As ListObject, ByVal colName As String)
    Dim r As Range
    On Error Resume Next
    Set r = lo.ListColumns(colName).DataBodyRange
    On Error GoTo 0

    If Not r Is Nothing Then
        r.NumberFormat = "0.0%"
        ApplyRejectPctThresholdCF r
    End If
End Sub

Private Sub ApplyRejectPctThresholdCF(ByVal rng As Range)
    On Error Resume Next
    rng.FormatConditions.Delete
    On Error GoTo 0

    Dim fc As FormatCondition

    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=0", Formula2:="=0.009999999")
    With fc
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
        .Font.Bold = False
    End With

    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=0.01", Formula2:="=0.019999999")
    With fc
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
        .Font.Bold = False
    End With

    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=0.02", Formula2:="=0.05")
    With fc
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = False
    End With

    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0.05")
    With fc
        .Interior.Color = RGB(112, 48, 160)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub

'=========================================================
' NAMED RANGE WRITE / RESOLUTION (workbook + sheet scoped)
'=========================================================
Private Sub PutNamed(ByVal wb As Workbook, ByVal nm As String, ByVal v As Variant)
    Dim r As Range
    Set r = ResolveNameRangeWB(wb, nm)
    If r Is Nothing Then Err.Raise vbObjectError + 400, , "Missing named range: " & nm
    r.Value = v
End Sub

Private Sub ApplyNamedFormats(ByVal wb As Workbook)
    ' Top KPIs
    SetNF wb, NM_RELEASED, "0"
    SetNF wb, NM_REJECT, "0"
    SetNF wb, NM_REWORK, "0"
    SetNF wb, NM_BGRADE, "0"
    SetNF wb, NM_REJECT_PERC, "0.00%"
    SetNF wb, NM_REWORK_PERC, "0.00%"

    ' WTD/MTD/YTD blocks
    SetNF wb, NM_WTD_REL, "0": SetNF wb, NM_WTD_REJ, "0": SetNF wb, NM_WTD_REW, "0": SetNF wb, NM_WTD_BG, "0"
    SetNF wb, NM_WTD_REJP, "0.00%": SetNF wb, NM_WTD_REWP, "0.00%"

    SetNF wb, NM_MTD_REL, "0": SetNF wb, NM_MTD_REJ, "0": SetNF wb, NM_MTD_REW, "0": SetNF wb, NM_MTD_BG, "0"
    SetNF wb, NM_MTD_REJP, "0.00%": SetNF wb, NM_MTD_REWP, "0.00%"

    SetNF wb, NM_YTD_REL, "0": SetNF wb, NM_YTD_REJ, "0": SetNF wb, NM_YTD_REW, "0": SetNF wb, NM_YTD_BG, "0"
    SetNF wb, NM_YTD_REJP, "0.00%": SetNF wb, NM_YTD_REWP, "0.00%"
End Sub

Private Sub SetNF(ByVal wb As Workbook, ByVal nm As String, ByVal fmt As String)
    Dim r As Range
    Set r = ResolveNameRangeWB(wb, nm)
    If Not r Is Nothing Then r.NumberFormat = fmt
End Sub

Private Function ResolveNameRangeWB(ByVal wb As Workbook, ByVal nm As String) As Range
    Dim n As name
    Dim target As String: target = LCase$(nm)

    ' workbook-scoped
    For Each n In wb.names
        If LCase$(n.name) = target Or LCase$(n.name) Like "*!" & target Then
            On Error Resume Next
            Set ResolveNameRangeWB = n.RefersToRange
            On Error GoTo 0
            If Not ResolveNameRangeWB Is Nothing Then Exit Function
        End If
    Next n
End Function

'=========================================================
' LISTOBJECT SAFE GETTERS
'=========================================================
Private Function GetListObjectSafe(ByVal ws As Worksheet, ByVal loName As String) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If StrComp(lo.name, loName, vbTextCompare) = 0 Then
            Set GetListObjectSafe = lo
            Exit Function
        End If
    Next lo
End Function

Private Sub EnsureTableRows(ByVal lo As ListObject, ByVal targetRows As Long)
    Dim curRows As Long
    If lo.DataBodyRange Is Nothing Then curRows = 0 Else curRows = lo.DataBodyRange.Rows.Count

    Application.EnableEvents = False
    Do While curRows < targetRows
        lo.ListRows.Add
        curRows = curRows + 1
    Loop
    Do While curRows > targetRows
        lo.ListRows(curRows).Delete
        curRows = curRows - 1
    Loop
    Application.EnableEvents = True
End Sub

Private Function HasColumn(ByVal lo As ListObject, ByVal header As String) As Boolean
    On Error Resume Next
    HasColumn = Not lo.ListColumns(header) Is Nothing
    On Error GoTo 0
End Function

Private Function ColIndex(ByVal lo As ListObject, ByVal header As String) As Long
    On Error GoTo EH
    ColIndex = lo.ListColumns(header).Index
    Exit Function
EH:
    Err.Raise vbObjectError + 601, "ColIndex", "Table '" & lo.name & "' is missing required column '" & header & "'."
End Function

'=========================================================
' SAFE SHEET GETTER
'=========================================================
Private Function GetSheetSafe(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.name, shName, vbTextCompare) = 0 Then
            Set GetSheetSafe = ws
            Exit Function
        End If
    Next ws
End Function

'=========================================================
' BUCKET ORDER
'=========================================================
Private Function OrderedBuckets() As Variant
    OrderedBuckets = Array("BK", "BK 30", "AIR", "CRM", "11.8", "12.4", "12.9", "12.4 30", "12.9 30", "12")
End Function

'=========================================================
' SKU PARSING
'=========================================================
Private Function LaneFromSKU(ByVal sku As String) As String
    Dim s As String: s = UCase$(sku)
    If InStr(1, s, "BVP", vbTextCompare) > 0 Then
        LaneFromSKU = "BVP"
    ElseIf InStr(1, s, "WVP", vbTextCompare) > 0 Then
        LaneFromSKU = "WVP"
    Else
        LaneFromSKU = "UNK"
    End If
End Function

Public Function RejectBucketFromSKU(ByVal sku As String) As String
    Dim s As String: s = UCase$(sku)
    Dim t As String
    t = Replace(Replace(s, ",", "."), "_", ".")

    If InStr(1, t, " AIR", vbTextCompare) > 0 Or InStr(1, t, "AIR", vbTextCompare) > 0 Then RejectBucketFromSKU = "AIR": Exit Function
    If InStr(1, t, " CRM", vbTextCompare) > 0 Or InStr(1, t, "CRM", vbTextCompare) > 0 Then RejectBucketFromSKU = "CRM": Exit Function

    If InStr(1, t, "BK", vbTextCompare) > 0 Then
        If InStr(1, t, "30", vbTextCompare) > 0 Then RejectBucketFromSKU = "BK 30" Else RejectBucketFromSKU = "BK"
        Exit Function
    End If

    If InStr(1, t, "11.8", vbTextCompare) > 0 Then RejectBucketFromSKU = "11.8": Exit Function
    If InStr(1, t, "12.4", vbTextCompare) > 0 Then
        If InStr(1, t, "30", vbTextCompare) > 0 Then RejectBucketFromSKU = "12.4 30" Else RejectBucketFromSKU = "12.4"
        Exit Function
    End If
    If InStr(1, t, "12.9", vbTextCompare) > 0 Then
        If InStr(1, t, "30", vbTextCompare) > 0 Then RejectBucketFromSKU = "12.9 30" Else RejectBucketFromSKU = "12.9"
        Exit Function
    End If

    If InStr(1, t, " 12 ", vbTextCompare) > 0 Or InStr(1, t, "12MM", vbTextCompare) > 0 Or InStr(1, t, "-12", vbTextCompare) > 0 Then
        RejectBucketFromSKU = "12"
        Exit Function
    End If

    RejectBucketFromSKU = ""
End Function

'=========================================================
' SMALL HELPERS
'=========================================================
Private Function NzD(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then
        NzD = 0
    ElseIf IsNumeric(v) Then
        NzD = CDbl(v)
    Else
        NzD = 0
    End If
End Function

Private Function SafePct(ByVal num As Double, ByVal den As Double) As Double
    If den = 0 Then SafePct = 0 Else SafePct = num / den
End Function

Private Function Make4(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Variant
    Dim v4(0 To 3) As Double
    v4(0) = a: v4(1) = b: v4(2) = c: v4(3) = d
    Make4 = v4
End Function

Private Sub AddTo4(ByRef a4 As Variant, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double)
    a4(0) = CDbl(a4(0)) + a
    a4(1) = CDbl(a4(1)) + b
    a4(2) = CDbl(a4(2)) + c
    a4(3) = CDbl(a4(3)) + d
End Sub

Private Function Max6(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal e As Long, ByVal f As Long) As Long
    Max6 = a
    If b > Max6 Then Max6 = b
    If c > Max6 Then Max6 = c
    If d > Max6 Then Max6 = d
    If e > Max6 Then Max6 = e
    If f > Max6 Then Max6 = f
End Function
