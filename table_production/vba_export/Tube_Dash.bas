Option Explicit

' =========================================================
'   TUBE INVENTORY MODEL + DASHBOARD (Complete Module)
' =========================================================

' =========================
' CONFIG - EDIT THESE ONLY
' =========================

' Sheet names
Private Const SH_MODEL As String = "Tube_Model"
Private Const SH_DASH  As String = "Tube_Dashboard"
Private Const SH_BIN   As String = "Bin_Inv"
Private Const SH_TUBE  As String = "Tube Schedule"
Private Const SH_PROD  As String = "Production Forecast"
Private Const SH_BOX_UTIL As String = "Box_Utilization"
Private Const SH_BOX_PLAN As String = "Box_Plan"
' ---- add a new dashboard sheet name (separate from the ramp dashboard) ----
Private Const SH_TUBE_DRILL As String = "Tube SKU Detail Dashboard"
Private Const SH_BOXED_COMP As String = "Boxed_By_Component"
Private Const SH_BOX_PLAN_ARCH As String = "Box_Plan_Archive"




' Bin_Inv layout
' Name = Column B, Qty = Column C (per your instruction)
Private Const BIN_COL_NAME As Long = 2
Private Const BIN_COL_QTY  As Long = 3
Private Const BIN_FIRST_ROW As Long = 2

' Tube Schedule layout (ASSUMPTION: same structure as Production Forecast weekly layout)
' Row 1 = dates across columns
' Col A = SKU
' Data starts row 2
Private Const TUBE_ROW_DATES As Long = 1
Private Const TUBE_COL_SKU   As Long = 1
Private Const TUBE_FIRST_DATA_ROW As Long = 4

' Production Forecast layout (ASSUMPTION: weekly dates in row 1, SKU in col A, data from row 2)
Private Const PROD_ROW_DATES As Long = 1
Private Const PROD_COL_SKU   As Long = 1
Private Const PROD_FIRST_DATA_ROW As Long = 5

' Timeline horizon: current month start through Dec 2026
Private Const END_YEAR As Long = 2026
Private Const END_MONTH As Long = 12

' Business rules
Private Const PROD_SHIFT_WEEKS As Long = 0  ' demand pulled earlier by 8 weeks
Private Const RECEIPT_LAG_WEEKS As Long = 4 ' tube ship date receipts apply next week



' Safety stock: 3-month average demand (monthly buckets from shifted demand)
Private Const SAFETY_MONTHS As Long = 3
Private gTubeDemandDrivers As Object   ' key: TubeSKU|WeekLong  value: dict(prodSKU -> qty)



' Tube SKUs in scope
Private Function TubeSKUList() As Variant
    TubeSKUList = Array( _
        "P CRM TUBE CARBON", _
        "P PRE TUBE CARBON 11.8MM Z", _
        "P PRE TUBE CARBON 12.4 30IN", _
        "P PRE TUBE CARBON 12.4MM", _
        "P PRE TUBE CARBON 12.9 30IN", _
        "P PRE TUBE CARBON 12.9MM", _
        "P PRE TUBE CARBON REVO JMP" _
    )
End Function

' Map belor -> tube SKU (1:1 mapping per your rule)
Private Function TubeSKUFromBelor(ByVal belor As String) As String
    belor = UCase$(Trim$(belor))

    ' Normalize 30" shafts regardless of pin
    If InStr(1, belor, "12.4", vbTextCompare) > 0 _
       And InStr(1, belor, "30", vbTextCompare) > 0 Then
        belor = "12.4 30"
    ElseIf InStr(1, belor, "12.9", vbTextCompare) > 0 _
       And InStr(1, belor, "30", vbTextCompare) > 0 Then
        belor = "12.9 30"
    End If

    Select Case belor
        Case "CRM"
            TubeSKUFromBelor = "P CRM TUBE CARBON"

        Case "11.8"
            TubeSKUFromBelor = "P PRE TUBE CARBON 11.8MM Z"

        Case "12.4"
            TubeSKUFromBelor = "P PRE TUBE CARBON 12.4MM"

        Case "12.9"
            TubeSKUFromBelor = "P PRE TUBE CARBON 12.9MM"

        Case "12.4 30"
            TubeSKUFromBelor = "P PRE TUBE CARBON 12.4 30IN"

        Case "12.9 30"
            TubeSKUFromBelor = "P PRE TUBE CARBON 12.9 30IN"

        Case "AIR"
            TubeSKUFromBelor = "P PRE TUBE CARBON REVO JMP"

        Case Else
            TubeSKUFromBelor = vbNullString
    End Select
End Function


' =========================
' MASTER
' =========================
Public Sub BuildTubeInventoryModel()
    Dim scrn As Boolean, evt As Boolean
    Dim calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    BuildTubeModelTable
    BuildTubeDashboard

CleanExit:
    Application.ScreenUpdating = scrn
    Application.EnableEvents = evt
    Application.Calculation = calc
    Exit Sub

CleanFail:
    MsgBox "BuildTubeInventoryModel failed in: " & Erl & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit

End Sub

' =========================
' MODEL BUILD
' =========================
Private Sub BuildTubeModelTable()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsM As Worksheet: Set wsM = EnsureSheet(wb, SH_MODEL)
    wsM.Cells.Clear
   

    ' Headers
    wsM.Range("A1:K1").Value = Array( _
        "Week", "SKU", "Begin Inv", "Receipts", _
        "Shifted Demand", "End Inv", _
        "Month", "Monthly Demand", "3M Avg Demand", _
        "Status", "Shortfall" _
    )
    wsM.Rows(1).Font.Bold = True

    ' Datasets
    Dim dInv As Object, dReceipts As Object, dDemand As Object
    Set dInv = CreateObject("Scripting.Dictionary")      ' sku -> onhand
    Set dReceipts = CreateObject("Scripting.Dictionary") ' sku|week -> receipts
    Set dDemand = CreateObject("Scripting.Dictionary")   ' sku|week -> shifted demand
    Set gTubeDemandDrivers = CreateObject("Scripting.Dictionary")

    LoadTubeInventory wb, dInv
    LoadTubeSchedule wb, dReceipts
    LoadShiftedTubeDemand wb, dDemand, gTubeDemandDrivers
    LoadTubeBoxDemand wb, dDemand, gTubeDemandDrivers, True, True



    ' Build week timeline (from production sheet headers, constrained to horizon)
    Dim weeks() As Date
    weeks = GetProductionWeeks(wb.Worksheets(SH_PROD))

    Dim rOut As Long: rOut = 2
    Dim sku As Variant, w As Variant
    Dim beginInv As Double, endInv As Double
    Dim rec As Double, dem As Double
    Dim key As String

    ' Ensure every tube SKU exists in inventory dict (0 if missing)
    Dim arr As Variant: arr = TubeSKUList()
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Not dInv.exists(CStr(arr(i))) Then dInv.Add CStr(arr(i)), 0#
    Next i

    For Each sku In dInv.keys
        beginInv = CDbl(dInv(sku))

        For Each w In weeks
            key = CStr(sku) & "|" & CLng(w)
            rec = GetDictDbl(dReceipts, key)
            dem = GetDictDbl(dDemand, key)

            endInv = beginInv + rec - dem

            wsM.Cells(rOut, 1).Value = w
            wsM.Cells(rOut, 1).NumberFormat = "m/d/yyyy"
            wsM.Cells(rOut, 2).Value = CStr(sku)
            wsM.Cells(rOut, 3).Value = beginInv
            wsM.Cells(rOut, 4).Value = rec
            wsM.Cells(rOut, 5).Value = dem
            wsM.Cells(rOut, 6).Value = endInv

            ' Month bucket
            wsM.Cells(rOut, 7).Value = DateSerial(Year(w), Month(w), 1)
            wsM.Cells(rOut, 7).NumberFormat = "mmm yyyy"

            beginInv = endInv
            rOut = rOut + 1
        Next w
    Next sku

    ' Monthly demand and 3M average demand (per SKU)
    ApplyMonthlyAnd3MAvg wsM

    ' Status / Shortfall
    ApplyTubeStatus wsM

    wsM.Columns.AutoFit
End Sub

' =========================
' LOADERS
' =========================

Private Sub LoadTubeInventory(ByVal wb As Workbook, ByVal dInv As Object)
    Dim ws As Worksheet: Set ws = wb.Worksheets(SH_BIN)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, BIN_COL_NAME).End(xlUp).row
    If lastRow < BIN_FIRST_ROW Then Exit Sub

    Dim allowed As Object: Set allowed = BuildAllowedTubeSKUSet()

    ' 1) Load raw Bin_Inv totals for tube SKUs
    Dim r As Long, sku As String, qty As Double
    For r = BIN_FIRST_ROW To lastRow
        sku = Trim$(CStr(ws.Cells(r, BIN_COL_NAME).Value))
        If Len(sku) > 0 And allowed.exists(sku) Then
            qty = SafeDbl(ws.Cells(r, BIN_COL_QTY).Value)
            If dInv.exists(sku) Then
                dInv(sku) = CDbl(dInv(sku)) + qty
            Else
                dInv.Add sku, qty
            End If
        End If
    Next r

    ' 2) Subtract qty currently sitting inside boxes (Boxed_By_Component)
    Dim dBoxed As Object: Set dBoxed = CreateObject("Scripting.Dictionary") ' tubeSKU -> qty boxed
    LoadBoxedByComponent_TubesOnly wb, dBoxed, allowed

    Dim k As Variant
    For Each k In dBoxed.keys
        If Not dInv.exists(CStr(k)) Then dInv.Add CStr(k), 0#
        dInv(CStr(k)) = CDbl(dInv(CStr(k))) - CDbl(dBoxed(CStr(k)))
    Next k

    ' 3) Add back qty for boxes that are planned for production (Box_Plan_Archive)
    Dim dAddBack As Object: Set dAddBack = CreateObject("Scripting.Dictionary") ' tubeSKU -> qty addback
    LoadBoxPlanArchive_AddBackTubes wb, dAddBack, allowed

    For Each k In dAddBack.keys
        If Not dInv.exists(CStr(k)) Then dInv.Add CStr(k), 0#
        dInv(CStr(k)) = CDbl(dInv(CStr(k))) + CDbl(dAddBack(CStr(k)))
    Next k
End Sub
Private Sub LoadBoxedByComponent_TubesOnly(ByVal wb As Workbook, ByVal dBoxed As Object, ByVal allowed As Object)
    On Error GoTo SafeExit

    Dim ws As Worksheet: Set ws = wb.Worksheets(SH_BOXED_COMP)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim r As Long, sku As String, q As Double
    For r = 2 To lastRow
        sku = Trim$(CStr(ws.Cells(r, 1).Value)) ' Component SKU
        If Len(sku) > 0 And allowed.exists(sku) Then
            q = SafeDbl(ws.Cells(r, 2).Value)   ' Boxed Unfinished Qty
            If q <> 0 Then
                If dBoxed.exists(sku) Then
                    dBoxed(sku) = CDbl(dBoxed(sku)) + q
                Else
                    dBoxed.Add sku, q
                End If
            End If
        End If
    Next r

SafeExit:
End Sub
Private Sub LoadBoxPlanArchive_AddBackTubes(ByVal wb As Workbook, ByVal dAddBack As Object, ByVal allowed As Object)
    On Error GoTo SafeExit

    Dim ws As Worksheet: Set ws = wb.Worksheets(SH_BOX_PLAN_ARCH)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim hdrRow As Long: hdrRow = 1

    Dim colQty As Long
    colQty = FindHeaderCol(ws, hdrRow, Array("Qty", "QTY", "Quantity"))
    If colQty = 0 Then Exit Sub

    Dim colTubeFlag As Long
    colTubeFlag = FindHeaderCol(ws, hdrRow, Array("TUBE"))
    If colTubeFlag = 0 Then Exit Sub

    Dim r As Long
    For r = 2 To lastRow
        If Not CBoolSafe(ws.Cells(r, colTubeFlag).Value) Then GoTo NextR

        Dim q As Double
        q = SafeDbl(ws.Cells(r, colQty).Value)
        If q = 0 Then GoTo NextR

        Dim tubeSKU As String
        tubeSKU = ExtractTubeSkuFromRow(ws, r, allowed)
        If Len(tubeSKU) = 0 Then GoTo NextR

        If dAddBack.exists(tubeSKU) Then
            dAddBack(tubeSKU) = CDbl(dAddBack(tubeSKU)) + q
        Else
            dAddBack.Add tubeSKU, q
        End If

NextR:
    Next r

SafeExit:
End Sub

Private Function CBoolSafe(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        CBoolSafe = CBool(v)
    Else
        Dim s As String: s = UCase$(Trim$(CStr(v)))
        CBoolSafe = (s = "TRUE" Or s = "YES" Or s = "1" Or s = "Y")
    End If
End Function

Private Function ExtractTubeSkuFromRow(ByVal ws As Worksheet, ByVal r As Long, ByVal allowed As Object) As String
    ' Searches the row for a string that matches / contains a tube SKU
    Dim c As Long
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function

    ' Prefer early columns (where your sheet stores identifiers)
    Dim scanMax As Long: scanMax = IIf(lastCol > 12, 12, lastCol)

    For c = 1 To scanMax
        Dim s As String: s = Trim$(CStr(ws.Cells(r, c).Value))
        If Len(s) > 0 Then
            ' exact match
            If allowed.exists(s) Then
                ExtractTubeSkuFromRow = s
                Exit Function
            End If

            ' substring match (handles truncated/embedded strings)
            Dim k As Variant
            For Each k In allowed.keys
                If InStr(1, s, CStr(k), vbTextCompare) > 0 Then
                    ExtractTubeSkuFromRow = CStr(k)
                    Exit Function
                End If
            Next k
        End If
    Next c
End Function



' Tube Schedule: ship dates across row 1; receipts apply next week (+1 week)
Private Sub LoadTubeSchedule(ByVal wb As Workbook, ByVal dReceipts As Object)
    Dim ws As Worksheet: Set ws = wb.Worksheets(SH_TUBE)

    Dim allowed As Object: Set allowed = BuildAllowedTubeSKUSet()

    ' Detect layout
    Dim dateRow As Long: dateRow = FindDateHeaderRow(ws, 25)
    If dateRow = 0 Then Exit Sub

    Dim skuCol As Long: skuCol = FindSKUCol(ws, dateRow)

    Dim firstRow As Long: firstRow = FindFirstSkuRow(ws, skuCol, dateRow, allowed)
    If firstRow = 0 Then Exit Sub

    Dim lastCol As Long: lastCol = ws.Cells(dateRow, ws.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, skuCol).End(xlUp).row
    If lastCol <= skuCol Or lastRow < firstRow Then Exit Sub

    Dim c As Long, r As Long
    Dim vDate As Variant, shipWeek As Date, recvWeek As Date
    Dim sku As String, qty As Double, key As String

    For c = skuCol + 1 To lastCol
        vDate = ws.Cells(dateRow, c).Value
        If IsDate(vDate) Then
            shipWeek = CDate(vDate)
            recvWeek = DateAdd("ww", RECEIPT_LAG_WEEKS, shipWeek)

            For r = firstRow To lastRow
                sku = Trim$(CStr(ws.Cells(r, skuCol).Value))
                If Len(sku) > 0 And allowed.exists(sku) Then
                    qty = SafeDbl(ws.Cells(r, c).Value)
                    If qty <> 0 Then
                        key = sku & "|" & CLng(recvWeek)
                        If dReceipts.exists(key) Then
                            dReceipts(key) = CDbl(dReceipts(key)) + qty
                        Else
                            dReceipts.Add key, qty
                        End If
                    End If
                End If
            Next r
        End If
    Next c
End Sub


' Demand: reuse consolidated logic scan (weekly) but SHIFT -8 weeks to floor date,
' then map belor -> tube SKU and store demand at floor week.
Private Sub LoadShiftedTubeDemand(ByVal wb As Workbook, ByVal dDemand As Object, ByVal dDrivers As Object)
    Dim wsPF As Worksheet: Set wsPF = wb.Worksheets(SH_PROD)

    ' Detect layout
    Dim dateRow As Long: dateRow = FindDateHeaderRow(wsPF, 25)
    If dateRow = 0 Then Exit Sub

    Dim skuCol As Long: skuCol = FindSKUCol(wsPF, dateRow)
    Dim firstRow As Long: firstRow = FindFirstSkuRow(wsPF, skuCol, dateRow, Nothing)
    If firstRow = 0 Then Exit Sub

    Dim lastCol As Long: lastCol = wsPF.Cells(dateRow, wsPF.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = wsPF.Cells(wsPF.Rows.Count, skuCol).End(xlUp).row
    If lastCol <= skuCol Or lastRow < firstRow Then Exit Sub

    Dim c As Long, r As Long
    Dim vDate As Variant, completeWeek As Date, floorWeek As Date
    Dim prodSKU As String, qty As Double
    Dim belor As String, tubeSKU As String
    Dim key As String

    For c = skuCol + 1 To lastCol
        vDate = wsPF.Cells(dateRow, c).Value
        If IsDate(vDate) Then
            completeWeek = CDate(vDate)
            floorWeek = DateAdd("ww", -PROD_SHIFT_WEEKS, completeWeek)

            For r = firstRow To lastRow
                prodSKU = Trim$(CStr(wsPF.Cells(r, skuCol).Value))
                If Len(prodSKU) > 0 Then
                    qty = SafeDbl(wsPF.Cells(r, c).Value)
                    If qty <> 0 Then
                        belor = GetBelorFromSKU(prodSKU)
                        If Len(belor) > 0 Then
                            tubeSKU = TubeSKUFromBelor(belor)
                            If Len(tubeSKU) > 0 Then
                                key = tubeSKU & "|" & CLng(floorWeek)

                                ' total shifted demand
                                If dDemand.exists(key) Then
                                    dDemand(key) = CDbl(dDemand(key)) + qty
                                Else
                                    dDemand.Add key, qty
                                End If

                                ' driver detail (prodSKU -> qty)
                                Dim dd As Object
                                If dDrivers.exists(key) Then
                                    Set dd = dDrivers(key)
                                Else
                                    Set dd = CreateObject("Scripting.Dictionary")
                                    dDrivers.Add key, dd
                                End If

                                If dd.exists(prodSKU) Then
                                    dd(prodSKU) = CDbl(dd(prodSKU)) + qty
                                Else
                                    dd.Add prodSKU, qty
                                End If
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    Next c
End Sub



' =========================
' MONTHLY + 3M AVERAGE DEMAND
' =========================
Private Sub ApplyMonthlyAnd3MAvg(ByVal wsM As Worksheet)
    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    ' Build monthly demand per SKU|Month from Shifted Demand column
    Dim dMon As Object: Set dMon = CreateObject("Scripting.Dictionary")
    Dim r As Long, sku As String
    Dim mStart As Date, mKey As String, key As String
    Dim dem As Double

    For r = 2 To lastRow
        sku = CStr(wsM.Cells(r, 2).Value)
        mStart = DateSerial(Year(wsM.Cells(r, 1).Value), Month(wsM.Cells(r, 1).Value), 1)
        mKey = CStr(mStart)
        dem = SafeDbl(wsM.Cells(r, 5).Value)
        If dem <> 0 Then
            key = sku & "|" & mKey
            If dMon.exists(key) Then
                dMon(key) = CDbl(dMon(key)) + dem
            Else
                dMon.Add key, dem
            End If
        End If
    Next r

    ' Write Monthly Demand and 3M Avg per row
    Dim dLast3 As Object: Set dLast3 = CreateObject("Scripting.Dictionary")
    ' We'll maintain rolling last-3 monthly totals per SKU as we move row order.
    ' Because rows are grouped by SKU then by week (from model build), this works.

    Dim curSku As String: curSku = vbNullString
    Dim curMonth As Date
    Dim q1 As Double, q2 As Double, q3 As Double ' oldest -> newest
    q1 = 0: q2 = 0: q3 = 0
    Dim seenMonths As Object
    Set seenMonths = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        sku = CStr(wsM.Cells(r, 2).Value)
        curMonth = DateSerial(Year(wsM.Cells(r, 1).Value), Month(wsM.Cells(r, 1).Value), 1)

        If sku <> curSku Then
            ' reset for new SKU
            curSku = sku
            q1 = 0: q2 = 0: q3 = 0
            seenMonths.RemoveAll
        End If

        ' monthly demand value for this SKU/month
        key = sku & "|" & CStr(curMonth)
        Dim monDem As Double: monDem = GetDictDbl(dMon, key)
        wsM.Cells(r, 8).Value = monDem

        ' update rolling months only once per month (first row encountered for that month)
        If Not seenMonths.exists(CStr(curMonth)) Then
            seenMonths.Add CStr(curMonth), True
            ' shift window
            q1 = q2
            q2 = q3
            q3 = monDem
        End If

        ' 3-month average: average of available months in window (avoid dividing by 3 when early)
        Dim cnt As Long: cnt = 0
        Dim sum3 As Double: sum3 = 0
        If q1 <> 0 Or seenMonths.Count >= 3 Then cnt = cnt + 1: sum3 = sum3 + q1
        If q2 <> 0 Or seenMonths.Count >= 2 Then cnt = cnt + 1: sum3 = sum3 + q2
        If q3 <> 0 Or seenMonths.Count >= 1 Then cnt = cnt + 1: sum3 = sum3 + q3

        If cnt = 0 Then
            wsM.Cells(r, 9).Value = 0
        Else
            wsM.Cells(r, 9).Value = sum3 / cnt
        End If
    Next r

    wsM.Columns(8).NumberFormat = "0"
    wsM.Columns(9).NumberFormat = "0"
End Sub

' =========================
' STATUS
' =========================
Private Sub ApplyTubeStatus(ByVal wsM As Worksheet)
    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlToLeft).row
    lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim r As Long
    Dim endInv As Double, safety As Double, shortfall As Double
    Dim status As String

    For r = 2 To lastRow
        endInv = SafeDbl(wsM.Cells(r, 6).Value)
        safety = SafeDbl(wsM.Cells(r, 9).Value)

        If endInv < 0 Then
            status = "STOCKOUT"
            shortfall = -endInv
        ElseIf endInv < safety Then
            status = "AT RISK"
            shortfall = safety - endInv
        Else
            status = "OK"
            shortfall = 0
        End If

        wsM.Cells(r, 10).Value = status
        wsM.Cells(r, 11).Value = shortfall
    Next r

    wsM.Columns(11).NumberFormat = "0"
End Sub

' =========================
' DASHBOARD (Simple + useful)
' =========================
Private Sub BuildTubeDashboard()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsD As Worksheet: Set wsD = EnsureSheet(wb, SH_DASH)

    ' HARD RESET (prevents merge leftovers + rerun issues)
    With wsD
        On Error Resume Next
        .Cells.UnMerge
        On Error GoTo 0
        .Cells.Clear
        .Cells.FormatConditions.Delete
        .Cells.Font.name = "Segoe UI"
        .Cells.Font.Size = 10
    End With

    ' Predator theme
    Const PRED_YELLOW As Long = 13041663
    Const DARK_HEADER As Long = 9211020
    Const LIGHT_GREY  As Long = 15132390

    ' Title bar (NO MERGE)
    With wsD.Range("A1:H1")
        .Value = "Tube Inventory Ramp Validation"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = PRED_YELLOW
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .RowHeight = 28
        .Borders.LineStyle = xlContinuous
        .ClearContents
        wsD.Range("A1").Value = "Tube Inventory Ramp Validation"
    End With

    ' KPI cards + values
    BuildTubeKPICards_NoMerge wsD, DARK_HEADER, LIGHT_GREY
    ComputeDashboardKPIsPretty wsD

    ' Status summary (now includes driver SKU)
BuildStatusSummaryPretty wsD, DARK_HEADER, LIGHT_GREY

' Clear the old right-table area so it disappears on rerun
wsD.Range("F8:Z2000").Clear



    wsD.Columns.AutoFit
End Sub

Private Function TopDriverSKUOnly(ByVal dd As Object) As String
    ' Returns the single highest-qty driver SKU from a driver dict (prodSKU -> qty)
    If dd Is Nothing Then Exit Function
    If dd.Count = 0 Then Exit Function

    Dim k As Variant, bestSKU As String
    Dim bestQty As Double: bestQty = -1#

    For Each k In dd.keys
        If CDbl(dd(k)) > bestQty Then
            bestQty = CDbl(dd(k))
            bestSKU = CStr(k)
        End If
    Next k

    TopDriverSKUOnly = bestSKU
End Function


Private Sub ComputeDashboardKPIs(ByVal wsD As Worksheet)
    Dim wsM As Worksheet: Set wsM = ThisWorkbook.Worksheets(SH_MODEL)

    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then
        wsD.Range("B4:B7").Value = "n/a"
        Exit Sub
    End If

    Dim firstStockout As Date: firstStockout = 0
    Dim worstEndInv As Double: worstEndInv = 9.9E+307

    Dim dRisk As Object: Set dRisk = CreateObject("Scripting.Dictionary")
    Dim dOut As Object: Set dOut = CreateObject("Scripting.Dictionary")

    Dim r As Long, sku As String, st As String
    Dim w As Date, endInv As Double

    For r = 2 To lastRow
        sku = CStr(wsM.Cells(r, 2).Value)
        w = CDate(wsM.Cells(r, 1).Value)
        endInv = SafeDbl(wsM.Cells(r, 6).Value)
        st = CStr(wsM.Cells(r, 10).Value)

        If endInv < worstEndInv Then worstEndInv = endInv

        If st = "STOCKOUT" Then
            dOut(sku) = True
            If firstStockout = 0 Or w < firstStockout Then firstStockout = w
        ElseIf st = "AT RISK" Then
            dRisk(sku) = True
        End If
    Next r

    If firstStockout = 0 Then
        wsD.Range("B4").Value = "No stockouts"
    Else
        wsD.Range("B4").Value = firstStockout
        wsD.Range("B4").NumberFormat = "m/d/yyyy"
    End If

    wsD.Range("B5").Value = worstEndInv
    wsD.Range("B5").NumberFormat = "0"

    wsD.Range("B6").Value = dRisk.Count
    wsD.Range("B7").Value = dOut.Count
End Sub

Private Sub BuildStatusSummary(ByVal wsD As Worksheet)
    Dim wsM As Worksheet: Set wsM = ThisWorkbook.Worksheets(SH_MODEL)
    Dim lastRow As Long: lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim dWorst As Object: Set dWorst = CreateObject("Scripting.Dictionary") ' sku -> status rank
    Dim dMinInv As Object: Set dMinInv = CreateObject("Scripting.Dictionary") ' sku -> min end inv

    Dim r As Long, sku As String, st As String
    Dim endInv As Double

    For r = 2 To lastRow
        sku = CStr(wsM.Cells(r, 2).Value)
        st = CStr(wsM.Cells(r, 10).Value)
        endInv = SafeDbl(wsM.Cells(r, 6).Value)

        Dim rank As Long: rank = StatusRank(st)

        If Not dWorst.exists(sku) Then
            dWorst.Add sku, rank
            dMinInv.Add sku, endInv
        Else
            If rank > CLng(dWorst(sku)) Then dWorst(sku) = rank
            If endInv < CDbl(dMinInv(sku)) Then dMinInv(sku) = endInv
        End If
    Next r

    ' Output table
    Dim startRow As Long: startRow = 10
    wsD.Cells(startRow, 1).Value = "Tube SKU"
    wsD.Cells(startRow, 2).Value = "Worst Status"
    wsD.Cells(startRow, 3).Value = "Min End Inv"
    wsD.Range(wsD.Cells(startRow, 1), wsD.Cells(startRow, 3)).Font.Bold = True

    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In dWorst.keys
        i = i + 1
        wsD.Cells(startRow + i, 1).Value = CStr(k)
        wsD.Cells(startRow + i, 2).Value = StatusFromRank(CLng(dWorst(k)))
        wsD.Cells(startRow + i, 3).Value = CDbl(dMinInv(k))
    Next k

    wsD.Columns(3).NumberFormat = "0"
End Sub

Private Function StatusRank(ByVal st As String) As Long
    st = UCase$(Trim$(st))
    Select Case st
        Case "OK": StatusRank = 1
        Case "AT RISK": StatusRank = 2
        Case "STOCKOUT": StatusRank = 3
        Case Else: StatusRank = 0
    End Select
End Function

Private Function StatusFromRank(ByVal rank As Long) As String
    Select Case rank
        Case 3: StatusFromRank = "STOCKOUT"
        Case 2: StatusFromRank = "AT RISK"
        Case 1: StatusFromRank = "OK"
        Case Else: StatusFromRank = "n/a"
    End Select
End Function

' =========================
' WEEK TIMELINE
' =========================
Private Function GetProductionWeeks(ByVal wsPF As Worksheet) As Date()
    Dim startMonth As Date, endMonth As Date
    startMonth = DateSerial(Year(Date), Month(Date), 1)
    endMonth = DateSerial(END_YEAR, END_MONTH, 1)

    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim dateRow As Long: dateRow = FindDateHeaderRow(wsPF, 20)
    If dateRow = 0 Then
        Dim fallback(0 To 0) As Date
        fallback(0) = startMonth
        GetProductionWeeks = fallback
        Exit Function
    End If

    Dim lastCol As Long: lastCol = wsPF.Cells(dateRow, wsPF.Columns.Count).End(xlToLeft).Column

    Dim c As Long, vDate As Variant, w As Date
    For c = 1 To lastCol
        vDate = wsPF.Cells(dateRow, c).Value
        If IsDate(vDate) Then
            w = CDate(vDate)
            If w >= startMonth And w <= DateAdd("m", 1, endMonth) Then
                If Not d.exists(CLng(w)) Then d.Add CLng(w), w
            End If
        End If
    Next c

    Dim keys() As Variant: keys = d.keys
    If d.Count = 0 Then
        Dim arr0(0 To 0) As Date
        arr0(0) = startMonth
        GetProductionWeeks = arr0
        Exit Function
    End If

    QuickSortVariantLong keys, LBound(keys), UBound(keys)

    Dim arr() As Date
    ReDim arr(0 To UBound(keys))
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        arr(i) = d(keys(i))
    Next i
    GetProductionWeeks = arr
End Function


Private Sub QuickSortVariantLong(ByRef a As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Double, tmp As Variant
    i = lo: j = hi
    pivot = CDbl(a((lo + hi) \ 2))
    Do While i <= j
        Do While CDbl(a(i)) < pivot: i = i + 1: Loop
        Do While CDbl(a(j)) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortVariantLong a, lo, j
    If i < hi Then QuickSortVariantLong a, i, hi
End Sub

' =========================
' HELPERS
' =========================
Private Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureSheet.name = sheetName
    End If
End Function

Private Function BuildAllowedTubeSKUSet() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = TubeSKUList()
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        d(CStr(arr(i))) = True
    Next i
    Set BuildAllowedTubeSKUSet = d
End Function

Private Function GetDictDbl(ByVal d As Object, ByVal key As String) As Double
    If d.exists(key) Then
        GetDictDbl = CDbl(d(key))
    Else
        GetDictDbl = 0#
    End If
End Function

Private Function SafeDbl(ByVal v As Variant) As Double
    If IsError(v) Then
        SafeDbl = 0#
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        SafeDbl = 0#
    ElseIf IsNumeric(v) Then
        SafeDbl = CDbl(v)
    Else
        SafeDbl = 0#
    End If
End Function

' Minimal belor extraction helper.
' IMPORTANT: If you already have GetBelorFromSKU in your REVO module, delete this and use yours.
Private Function GetBelorFromSKU(ByVal sku As String) As String
    ' Fallback classifier for tube selection
    Dim s As String: s = UCase$(Trim$(sku))

    ' --- Helper-style checks done inline for speed/clarity ---
    Dim has30 As Boolean
    has30 = (InStr(1, s, " 30", vbTextCompare) > 0) Or (InStr(1, s, "30IN", vbTextCompare) > 0) Or (InStr(1, s, "30""", vbTextCompare) > 0)

    ' 11.8
    If InStr(1, s, "11.8", vbTextCompare) > 0 Then GetBelorFromSKU = "11.8": Exit Function

    ' 12.4 30" (handle 12.4 UNI 30 / 12.4 RAD 30 / 12.4-30 / 12.4 30IN)
    If (InStr(1, s, "12.4", vbTextCompare) > 0) And (has30 Or InStr(1, s, "12.4-30", vbTextCompare) > 0) Then
        GetBelorFromSKU = "12.4 30"
        Exit Function
    End If

    ' 12.9 30" (also allow BK 30)
    If ((InStr(1, s, "12.9", vbTextCompare) > 0) And (has30 Or InStr(1, s, "12.9-30", vbTextCompare) > 0)) _
        Or (InStr(1, s, "BK", vbTextCompare) > 0 And has30) Then
        GetBelorFromSKU = "12.9 30"
        Exit Function
    End If

    ' 12.4 (non-30)
    If InStr(1, s, "12.4", vbTextCompare) > 0 Then GetBelorFromSKU = "12.4": Exit Function

    ' 12.9 (non-30) or BK (non-30)
    If InStr(1, s, "12.9", vbTextCompare) > 0 Or InStr(1, s, "BK", vbTextCompare) > 0 Then GetBelorFromSKU = "12.9": Exit Function

    ' CRM
    If InStr(1, s, "CRM", vbTextCompare) > 0 Then GetBelorFromSKU = "CRM": Exit Function

    ' AIR / JMP
    If InStr(1, s, "AIR", vbTextCompare) > 0 Or InStr(1, s, "JMP", vbTextCompare) > 0 Then GetBelorFromSKU = "AIR": Exit Function

    GetBelorFromSKU = vbNullString
End Function


Private Function FindDateHeaderRow(ByVal ws As Worksheet, Optional ByVal scanTopRows As Long = 15) As Long
    Dim r As Long, c As Long, lastCol As Long, hits As Long
    For r = 1 To scanTopRows
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        hits = 0
        For c = 1 To lastCol
            If IsDate(ws.Cells(r, c).Value) Then hits = hits + 1
            If hits >= 3 Then
                FindDateHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
    FindDateHeaderRow = 0
End Function

Private Function FindSKUCol(ByVal ws As Worksheet, ByVal dateRow As Long) As Long
    ' SKU column is usually the first non-date text column on the dateRow
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(dateRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Not IsDate(ws.Cells(dateRow, c).Value) Then
            FindSKUCol = c
            Exit Function
        End If
    Next c
    FindSKUCol = 1
End Function

Private Function FindFirstSkuRow(ByVal ws As Worksheet, ByVal skuCol As Long, ByVal afterRow As Long, ByVal allowed As Object) As Long
    Dim r As Long, lastRow As Long, v As String
    lastRow = ws.Cells(ws.Rows.Count, skuCol).End(xlUp).row
    For r = afterRow + 1 To lastRow
        v = Trim$(CStr(ws.Cells(r, skuCol).Value))
        If Len(v) > 0 Then
            If allowed Is Nothing Then
                FindFirstSkuRow = r
                Exit Function
            ElseIf allowed.exists(v) Then
                FindFirstSkuRow = r
                Exit Function
            End If
        End If
    Next r
    FindFirstSkuRow = 0
End Function


Private Sub BuildTubeKPICards_NoMerge(ByVal wsD As Worksheet, ByVal headerColor As Long, ByVal bandColor As Long)
    Dim labels As Variant
    labels = Array("First Stockout Date", "Worst Ending Inventory", "# SKUs At Risk", "# SKUs Stockout")

    Dim i As Long, c0 As Long
    For i = 0 To 3
        c0 = 1 + (i * 2) ' 1,3,5,7

        ' Header strip (no merge)
        With wsD.Range(wsD.Cells(3, c0), wsD.Cells(3, c0 + 1))
            .ClearContents
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = headerColor
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        wsD.Cells(3, c0).Value = labels(i)

        ' Value box (no merge)
        With wsD.Range(wsD.Cells(4, c0), wsD.Cells(6, c0 + 1))
            .ClearContents
            .Font.Bold = True
            .Font.Size = 14
            .Interior.Color = bandColor
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
    Next i
End Sub


Private Sub ComputeDashboardKPIsPretty(ByVal wsD As Worksheet)
    Dim wsM As Worksheet: Set wsM = ThisWorkbook.Worksheets(SH_MODEL)
    Dim lastRow As Long: lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim firstStockout As Date: firstStockout = 0
    Dim worstEndInv As Double: worstEndInv = 9.9E+307
    Dim dRisk As Object: Set dRisk = CreateObject("Scripting.Dictionary")
    Dim dOut As Object: Set dOut = CreateObject("Scripting.Dictionary")

    Dim r As Long, sku As String, st As String, w As Date, endInv As Double
    For r = 2 To lastRow
        sku = CStr(wsM.Cells(r, 2).Value)
        w = CDate(wsM.Cells(r, 1).Value)
        endInv = SafeDbl(wsM.Cells(r, 6).Value)
        st = CStr(wsM.Cells(r, 10).Value)

        If endInv < worstEndInv Then worstEndInv = endInv

        If st = "STOCKOUT" Then
            dOut(sku) = True
            If firstStockout = 0 Or w < firstStockout Then firstStockout = w
        ElseIf st = "AT RISK" Then
            dRisk(sku) = True
        End If
    Next r

    ' Write values into card bodies (row 4 col 1,3,5,7)
    If firstStockout = 0 Then
        wsD.Cells(4, 1).Value = "No stockouts"
    Else
        wsD.Cells(4, 1).Value = firstStockout
        wsD.Cells(4, 1).NumberFormat = "m/d/yyyy"
    End If

    wsD.Cells(4, 3).Value = worstEndInv
    wsD.Cells(4, 3).NumberFormat = "0"

    wsD.Cells(4, 5).Value = dRisk.Count
    wsD.Cells(4, 7).Value = dOut.Count
End Sub

Private Sub BuildStatusSummaryPretty(ByVal wsD As Worksheet, ByVal headerColor As Long, ByVal bandColor As Long)
    ' Summary starts at A8
    Dim topRow As Long: topRow = 8

    wsD.Range("A" & topRow & ":E" & topRow).Value = _
        Array("Tube SKU", "Worst Status", "Min End Inv", "First Breach Week", "Driver Prod SKU (Breach Week)")

    With wsD.Range("A" & topRow & ":E" & topRow)
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = headerColor
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With

    Dim wsM As Worksheet: Set wsM = ThisWorkbook.Worksheets(SH_MODEL)
    Dim lastRow As Long: lastRow = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim dWorst As Object: Set dWorst = CreateObject("Scripting.Dictionary")
    Dim dMinInv As Object: Set dMinInv = CreateObject("Scripting.Dictionary")
    Dim dFirstBreach As Object: Set dFirstBreach = CreateObject("Scripting.Dictionary")

    Dim r As Long, sku As String, st As String
    Dim endInv As Double, w As Date, rank As Long

    ' Build summary dictionaries
    For r = 2 To lastRow
        sku = Trim$(CStr(wsM.Cells(r, 2).Value))
        If Len(sku) = 0 Then GoTo NextR

        If IsDate(wsM.Cells(r, 1).Value) Then
            w = CDate(wsM.Cells(r, 1).Value)
        Else
            GoTo NextR
        End If

        st = UCase$(Trim$(CStr(wsM.Cells(r, 10).Value)))
        st = Replace$(st, "STOCK OUT", "STOCKOUT")

        endInv = SafeDbl(wsM.Cells(r, 6).Value)
        rank = StatusRank(st)

        If Not dWorst.exists(sku) Then
            dWorst.Add sku, rank
            dMinInv.Add sku, endInv
            If st <> "OK" Then dFirstBreach.Add sku, w
        Else
            If rank > CLng(dWorst(sku)) Then dWorst(sku) = rank
            If endInv < CDbl(dMinInv(sku)) Then dMinInv(sku) = endInv

            If st <> "OK" Then
                If Not dFirstBreach.exists(sku) Then
                    dFirstBreach.Add sku, w
                ElseIf w < CDate(dFirstBreach(sku)) Then
                    dFirstBreach(sku) = w
                End If
            End If
        End If

NextR:
    Next r

    ' Output
    Dim outRow As Long: outRow = topRow + 1
    Dim k As Variant

    For Each k In dWorst.keys
        wsD.Cells(outRow, 1).Value = CStr(k)
        wsD.Cells(outRow, 2).Value = StatusFromRank(CLng(dWorst(k)))

        wsD.Cells(outRow, 3).Value = CDbl(dMinInv(k))
        wsD.Cells(outRow, 3).NumberFormat = "#,##0;[Red]-#,##0"

        If dFirstBreach.exists(k) Then
            Dim bw As Date
            bw = CDate(dFirstBreach(k))

            wsD.Cells(outRow, 4).Value = bw
            wsD.Cells(outRow, 4).NumberFormat = "m/d/yyyy"

            ' Driver at breach week
            Dim drvKey As String
            drvKey = CStr(k) & "|" & CLng(bw)

            Dim dd As Object
            If Not (gTubeDemandDrivers Is Nothing) Then
                If gTubeDemandDrivers.exists(drvKey) Then
                    Set dd = gTubeDemandDrivers(drvKey)
                    wsD.Cells(outRow, 5).Value = TopDriverSKUOnly(dd)
                Else
                    wsD.Cells(outRow, 5).Value = vbNullString
                End If
            Else
                wsD.Cells(outRow, 5).Value = vbNullString
            End If
        Else
            wsD.Cells(outRow, 4).Value = vbNullString
            wsD.Cells(outRow, 5).Value = vbNullString
        End If

        outRow = outRow + 1
    Next k

    Dim lastOut As Long: lastOut = outRow - 1
    If lastOut < topRow + 1 Then Exit Sub

    ' Table borders + banding
    With wsD.Range("A" & topRow & ":E" & lastOut)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With

    Dim rr As Long
    For rr = topRow + 1 To lastOut
        If rr Mod 2 = 0 Then
            wsD.Range("A" & rr & ":E" & rr).Interior.Color = bandColor
        End If
    Next rr

    ' Conditional format by status
    With wsD.Range("B" & (topRow + 1) & ":B" & lastOut)
        .FormatConditions.Delete
        Dim fc As FormatCondition

        Set fc = .FormatConditions.Add(Type:=xlTextString, String:="STOCKOUT", TextOperator:=xlContains)
        fc.Interior.Color = RGB(255, 180, 180)

        Set fc = .FormatConditions.Add(Type:=xlTextString, String:="AT RISK", TextOperator:=xlContains)
        fc.Interior.Color = RGB(255, 230, 180)

        Set fc = .FormatConditions.Add(Type:=xlTextString, String:="OK", TextOperator:=xlContains)
        fc.Interior.Color = RGB(200, 240, 200)
    End With
End Sub


Public Sub BuildNegativeWeeksWithDrivers(wsD As Worksheet)
    Const START_ROW As Long = 10
    Const START_COL As Long = 6 ' Column F

    Dim wsM As Worksheet, wsPF As Worksheet
    Set wsM = ThisWorkbook.Worksheets("Tube_Model")
    Set wsPF = ThisWorkbook.Worksheets("Production Forecast")

    Dim lastRowM As Long
    lastRowM = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).row
    If lastRowM < 2 Then Exit Sub

    ' Clear old table
    wsD.Range(wsD.Cells(START_ROW - 2, START_COL), wsD.Cells(wsD.Rows.Count, START_COL + 6)).Clear

    ' Header
    wsD.Cells(START_ROW - 2, START_COL).Resize(1, 7).Value = _
    Array("Week", "Tube SKU", "Status", "End Inv", "Safety (3M Avg)", "Driver Prod SKU", "Driver Complete Week")


    With wsD.Cells(START_ROW - 2, START_COL).Resize(1, 7)
        .Font.Bold = True
        .Interior.Color = RGB(120, 120, 120)
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    Dim r As Long, outRow As Long
    outRow = START_ROW

    ' Track last driver to avoid spam rows
    Dim lastDriverKey As String
    lastDriverKey = ""

    For r = 2 To lastRowM
        Dim endInv As Double
        endInv = val(wsM.Cells(r, 6).Value)

        Dim wk As Date, tubeSKU As String, st As String
        wk = wsM.Cells(r, 1).Value
        tubeSKU = Trim$(CStr(wsM.Cells(r, 2).Value))

        ' Status column (your model uses Col 10 per KPI routine)
        st = UCase$(Trim$(CStr(wsM.Cells(r, 10).Value)))
        st = Replace$(st, "STOCK OUT", "STOCKOUT")

        ' Include by status (not only negative endInv)
        If (st = "STOCKOUT") Or (st = "AT RISK") Or (endInv < 0) Then

            ' Find production driver
            Dim drvSKU As String, drvWeek As Variant
            drvSKU = vbNullString
            drvWeek = vbNullString
            FindProductionDriver wsPF, tubeSKU, wk, drvSKU, drvWeek

            ' Optional: keep your crowding rule (skip rows with no driver)
            If drvSKU <> "" Then
                Dim driverKey As String
                driverKey = tubeSKU & "|" & st & "|" & drvSKU & "|" & CStr(drvWeek)

                If driverKey <> lastDriverKey Then
                    Dim safety As Double, shortfall As Double
safety = SafeDbl(wsM.Cells(r, 9).Value)
shortfall = SafeDbl(wsM.Cells(r, 11).Value)

wsD.Cells(outRow, START_COL).Value = wk
wsD.Cells(outRow, START_COL).NumberFormat = "m/d/yyyy"

wsD.Cells(outRow, START_COL + 1).Value = tubeSKU
wsD.Cells(outRow, START_COL + 2).Value = st

wsD.Cells(outRow, START_COL + 3).Value = endInv
wsD.Cells(outRow, START_COL + 3).NumberFormat = "#,##0;[Red]-#,##0"

wsD.Cells(outRow, START_COL + 4).Value = safety
wsD.Cells(outRow, START_COL + 4).NumberFormat = "#,##0;[Red]-#,##0"

wsD.Cells(outRow, START_COL + 5).Value = drvSKU
wsD.Cells(outRow, START_COL + 6).Value = drvWeek
wsD.Cells(outRow, START_COL + 6).NumberFormat = "m/d/yyyy"


                    outRow = outRow + 1
                    lastDriverKey = driverKey
                End If
            End If

        End If
    Next r

    ' Formatting
    If outRow > START_ROW Then
        With wsD.Range(wsD.Cells(START_ROW - 2, START_COL), wsD.Cells(outRow - 1, START_COL + 6))
            .Borders.LineStyle = xlContinuous
            .Columns.AutoFit
        End With
    End If
End Sub

Private Sub FindProductionDriver(wsPF As Worksheet, tubeSKU As String, _
                                 floorWeek As Date, _
                                 ByRef outSKU As String, _
                                 ByRef outCompleteWeek As Variant)

    outSKU = ""
    outCompleteWeek = ""

    Dim belor As String
    belor = GetBelorFromSKU(tubeSKU)
    If belor = "" Then Exit Sub

    Dim dateRow As Long: dateRow = FindDateHeaderRow(wsPF, 25)
    If dateRow = 0 Then Exit Sub

    Dim skuCol As Long: skuCol = FindSKUCol(wsPF, dateRow)
    Dim firstRow As Long: firstRow = FindFirstSkuRow(wsPF, skuCol, dateRow, Nothing)

    Dim lastCol As Long
    lastCol = wsPF.Cells(dateRow, wsPF.Columns.Count).End(xlToLeft).Column

    Dim c As Long, r As Long
    For c = skuCol + 1 To lastCol
        Dim completeWeek As Date
        If IsDate(wsPF.Cells(dateRow, c).Value) Then
            completeWeek = wsPF.Cells(dateRow, c).Value

            ' Production completion ? floor week = -8 weeks
            If DateAdd("ww", -PROD_SHIFT_WEEKS, completeWeek) = floorWeek Then

                For r = firstRow To wsPF.Cells(wsPF.Rows.Count, skuCol).End(xlUp).row
                    Dim prodSKU As String
                    prodSKU = wsPF.Cells(r, skuCol).Value
                    If prodSKU <> "" Then
                        If GetBelorFromSKU(prodSKU) = belor Then
                            If wsPF.Cells(r, c).Value <> 0 Then
                                outSKU = prodSKU
                                outCompleteWeek = completeWeek
                                Exit Sub
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    Next c
End Sub


Private Sub BuildTubeDriverMap(ByVal dDriverSKU As Object, ByVal dDriverWeek As Object)

    Dim wsPF As Worksheet: Set wsPF = ThisWorkbook.Worksheets(SH_PROD)

    Dim dateRow As Long: dateRow = FindDateHeaderRow(wsPF, 25)
    If dateRow = 0 Then Exit Sub

    Dim skuCol As Long: skuCol = FindSKUCol(wsPF, dateRow)

    Dim firstRow As Long: firstRow = FindFirstSkuRow(wsPF, skuCol, dateRow, Nothing)
    If firstRow = 0 Then Exit Sub

    Dim lastCol As Long: lastCol = wsPF.Cells(dateRow, wsPF.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = wsPF.Cells(wsPF.Rows.Count, skuCol).End(xlUp).row
    If lastCol <= skuCol Or lastRow < firstRow Then Exit Sub

    ' track best qty per (TubeSKU|FloorWeek)
    Dim dBestQty As Object: Set dBestQty = CreateObject("Scripting.Dictionary")

    Dim c As Long, r As Long
    Dim vDate As Variant, completeWeek As Date, floorWeek As Date
    Dim prodSKU As String, belor As String, tubeSKU As String
    Dim qty As Double, key As String

    For c = skuCol + 1 To lastCol
        vDate = wsPF.Cells(dateRow, c).Value
        If IsDate(vDate) Then
            completeWeek = CDate(vDate)
            floorWeek = DateAdd("ww", -PROD_SHIFT_WEEKS, completeWeek)

            For r = firstRow To lastRow
                prodSKU = Trim$(CStr(wsPF.Cells(r, skuCol).Value))
                If Len(prodSKU) > 0 Then
                    qty = SafeDbl(wsPF.Cells(r, c).Value)
                    If qty <> 0 Then
                        belor = GetBelorFromSKU(prodSKU)
                        If Len(belor) > 0 Then
                            tubeSKU = TubeSKUFromBelor(belor)
                            If Len(tubeSKU) > 0 Then
                                key = tubeSKU & "|" & CLng(floorWeek)

                                If (Not dBestQty.exists(key)) Or qty > CDbl(dBestQty(key)) Then
                                    dBestQty(key) = qty
                                    dDriverSKU(key) = prodSKU
                                    dDriverWeek(key) = completeWeek
                                End If
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    Next c

End Sub




Private Sub Sort2DByWeekThenSku_5Cols(ByRef a As Variant, ByVal n As Long)
    Dim i As Long, j As Long
    Dim t1 As Variant, t2 As Variant, t3 As Variant, t4 As Variant, t5 As Variant
    Dim w As Date, sku As String

    For i = 2 To n
        t1 = a(i, 1): t2 = a(i, 2): t3 = a(i, 3): t4 = a(i, 4): t5 = a(i, 5)
        w = t1: sku = CStr(t2)

        j = i - 1
        Do While j >= 1
            If a(j, 1) > w Then
                a(j + 1, 1) = a(j, 1): a(j + 1, 2) = a(j, 2): a(j + 1, 3) = a(j, 3): a(j + 1, 4) = a(j, 4): a(j + 1, 5) = a(j, 5)
            ElseIf a(j, 1) = w And CStr(a(j, 2)) > sku Then
                a(j + 1, 1) = a(j, 1): a(j + 1, 2) = a(j, 2): a(j + 1, 3) = a(j, 3): a(j + 1, 4) = a(j, 4): a(j + 1, 5) = a(j, 5)
            Else
                Exit Do
            End If
            j = j - 1
        Loop

        a(j + 1, 1) = t1: a(j + 1, 2) = t2: a(j + 1, 3) = t3: a(j + 1, 4) = t4: a(j + 1, 5) = t5
    Next i
End Sub


Private Function TopDriversString(ByVal dd As Object, Optional ByVal topN As Long = 3) As String
    If dd Is Nothing Then Exit Function
    If dd.Count = 0 Then Exit Function

    Dim keys() As Variant: keys = dd.keys
    Dim i As Long, j As Long
    Dim tmp As Variant

    ' sort keys descending by qty
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CDbl(dd(keys(j))) > CDbl(dd(keys(i))) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i

    Dim n As Long: n = topN
    If n > dd.Count Then n = dd.Count

    Dim parts() As String
    ReDim parts(1 To n)

    For i = 1 To n
        parts(i) = CStr(keys(i - 1)) & ": " & Format$(CDbl(dd(keys(i - 1))), "0")
    Next i

    TopDriversString = Join(parts, " | ")
End Function

Private Sub BuildDriverLookup_FromProductionForecast(ByVal wsPF As Worksheet, ByVal dDriver As Object, ByVal scanTopRows As Long)
    Dim dateRow As Long: dateRow = FindDateHeaderRow(wsPF, scanTopRows)
    If dateRow = 0 Then Exit Sub

    Dim skuCol As Long: skuCol = FindSKUCol(wsPF, dateRow)

    Dim firstRow As Long: firstRow = FindFirstSkuRow(wsPF, skuCol, dateRow, Nothing)
    If firstRow = 0 Then Exit Sub

    Dim lastCol As Long: lastCol = wsPF.Cells(dateRow, wsPF.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = wsPF.Cells(wsPF.Rows.Count, skuCol).End(xlUp).row
    If lastCol <= skuCol Or lastRow < firstRow Then Exit Sub

    Dim c As Long, r As Long
    Dim vDate As Variant, completeWeek As Date, floorWeek As Date
    Dim prodSKU As String, qty As Double
    Dim belor As String, tubeSKU As String
    Dim key As String
    Dim cur As Variant

    For c = skuCol + 1 To lastCol
        vDate = wsPF.Cells(dateRow, c).Value
        If IsDate(vDate) Then
            completeWeek = CDate(vDate)
            floorWeek = DateAdd("ww", -PROD_SHIFT_WEEKS, completeWeek)

            For r = firstRow To lastRow
                prodSKU = Trim$(CStr(wsPF.Cells(r, skuCol).Value))
                If Len(prodSKU) > 0 Then
                    qty = SafeDbl(wsPF.Cells(r, c).Value)
                    If qty <> 0 Then
                        belor = GetBelorFromSKU(prodSKU)
                        If Len(belor) > 0 Then
                            tubeSKU = TubeSKUFromBelor(belor)
                            If Len(tubeSKU) > 0 Then
                                key = tubeSKU & "|" & CLng(floorWeek)

                                ' keep the largest single contributor for that tube week
                                If dDriver.exists(key) Then
                                    cur = dDriver(key) ' Array(prodSKU, completeWeek, qty)
                                    If qty > CDbl(cur(2)) Then
                                        dDriver(key) = Array(prodSKU, completeWeek, qty)
                                    End If
                                Else
                                    dDriver.Add key, Array(prodSKU, completeWeek, qty)
                                End If
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    Next c
End Sub

'==========================================================
' ADD THIS “BOX LOGIC” PATCH TO YOUR TUBE MODULE
' - Treats Box_Plan as EXTRA tube demand (consumption)
' - Optionally treats Box_Utilization as “already consumed” (past consumption)
' - Uses Finished SKU -> Belor -> Tube SKU mapping you already have
'
' ASSUMPTIONS (safe defaults):
'   • Box_Plan has columns: Finished SKU | Box | Installed List | Production Date
'   • One row = one tube consumed (unless a Qty column exists)
'   • Consumption week = week of Production Date (no lead/lag)
'==========================================================

'-----------------------------
' 1) ADD THESE SHEET CONSTANTS
'-----------------------------


'-----------------------------
' 2) IN BuildTubeModelTable(), AFTER LoadShiftedTubeDemand, ADD:
'-----------------------------
'    LoadTubeBoxDemand wb, dDemand, gTubeDemandDrivers, True, True
'
'    (the two booleans are: include Plan, include Utilization)
'
' Example:
'    LoadShiftedTubeDemand wb, dDemand, gTubeDemandDrivers
'    LoadTubeBoxDemand wb, dDemand, gTubeDemandDrivers, True, True
'-----------------------------

'==========================================================
' 3) BOX DEMAND LOADER (ADD TO MODULE)
'==========================================================
Private Sub LoadTubeBoxDemand(ByVal wb As Workbook, _
                             ByVal dDemand As Object, _
                             ByVal dDrivers As Object, _
                             Optional ByVal includePlan As Boolean = True, _
                             Optional ByVal includeUtil As Boolean = True)

    Dim ws As Worksheet

    If includePlan Then
        On Error Resume Next
        Set ws = wb.Worksheets(SH_BOX_PLAN)
        On Error GoTo 0
        If Not ws Is Nothing Then
            LoadTubeBoxDemand_FromSheet ws, dDemand, dDrivers, "BOX_PLAN"
        End If
        Set ws = Nothing
    End If

    If includeUtil Then
        On Error Resume Next
        Set ws = wb.Worksheets(SH_BOX_UTIL)
        On Error GoTo 0
        If Not ws Is Nothing Then
            LoadTubeBoxDemand_FromSheet ws, dDemand, dDrivers, "BOX_UTIL"
        End If
        Set ws = Nothing
    End If
End Sub

'==========================================================
' 4) SHARED PARSER (reads either Box_Plan or Box_Utilization)
'==========================================================
Private Sub LoadTubeBoxDemand_FromSheet(ByVal ws As Worksheet, _
                                       ByVal dDemand As Object, _
                                       ByVal dDrivers As Object, _
                                       ByVal sourceTag As String)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    ' Find key columns by header (row 1), fallback to assumed positions
    Dim colSKU As Long, colDate As Long, colQty As Long
    colSKU = FindHeaderCol(ws, 1, Array("Finished SKU", "SKU", "Item"))
    colDate = FindHeaderCol(ws, 1, Array("Production Date", "Prod Date", "Date"))
    colQty = FindHeaderCol(ws, 1, Array("Qty", "QTY", "Quantity"))

    If colSKU = 0 Then colSKU = 1
    If colDate = 0 Then colDate = 4   ' your described format: 4th field is Production Date

    Dim r As Long
    For r = 2 To lastRow
        Dim finishedSku As String
        finishedSku = Trim$(CStr(ws.Cells(r, colSKU).Value))
        If Len(finishedSku) = 0 Then GoTo NextR

        Dim vDate As Variant
        vDate = ws.Cells(r, colDate).Value
        If Not IsDate(vDate) Then GoTo NextR

        Dim wk As Date
        wk = CDate(vDate) ' treat as week bucket date already (or any date; you’re keying by CLng(Date))

        ' Map Finished SKU -> Belor -> Tube SKU
        Dim belor As String, tubeSKU As String
        belor = GetBelorFromSKU(finishedSku)
        If Len(belor) = 0 Then GoTo NextR

        tubeSKU = TubeSKUFromBelor(belor)
        If Len(tubeSKU) = 0 Then GoTo NextR

        Dim qty As Double
        qty = 1#
        If colQty > 0 Then
            Dim qv As Variant
            qv = ws.Cells(r, colQty).Value
            If IsNumeric(qv) Then
                If CDbl(qv) <> 0 Then qty = CDbl(qv)
            End If
        End If
        If qty = 0 Then GoTo NextR

        Dim key As String
        key = tubeSKU & "|" & CLng(wk)

        ' Add to shifted demand (tube consumption)
        If dDemand.exists(key) Then
            dDemand(key) = CDbl(dDemand(key)) + qty
        Else
            dDemand.Add key, qty
        End If

        ' Add driver detail so dashboard can show it (tagged source)
        Dim dd As Object
        If dDrivers.exists(key) Then
            Set dd = dDrivers(key)
        Else
            Set dd = CreateObject("Scripting.Dictionary")
            dDrivers.Add key, dd
        End If

        Dim drvLabel As String
        drvLabel = sourceTag & ": " & finishedSku

        If dd.exists(drvLabel) Then
            dd(drvLabel) = CDbl(dd(drvLabel)) + qty
        Else
            dd.Add drvLabel, qty
        End If

NextR:
    Next r
End Sub

'==========================================================
' 5) HEADER FINDER (ADD TO MODULE)
'==========================================================
Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal names As Variant) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long, i As Long
    Dim h As String
    For c = 1 To lastCol
        h = UCase$(Trim$(CStr(ws.Cells(headerRow, c).Value)))
        If Len(h) > 0 Then
            For i = LBound(names) To UBound(names)
                If h = UCase$(Trim$(CStr(names(i)))) Then
                    FindHeaderCol = c
                    Exit Function
                End If
            Next i
        End If
    Next c
    FindHeaderCol = 0
End Function



'==========================================================
' TUBE SKU DRILLDOWN DASHBOARD
' Works like your SKU drilldown, but:
'   • Inventory = Bin_Inv (Tube SKUs only)
'   • Incoming = Tube Schedule (NOT POs)
'   • Demand = Production Forecast mapped via Belor -> Tube SKU
'   • Optional: adds Box_Plan / Box_Utilization as extra tube demand if you’ve enabled that logic
'
' Drop this into the SAME module as your Tube Inventory Model
' (so it can reuse: TubeSKUList, TubeSKUFromBelor, GetBelorFromSKU,
'  FindDateHeaderRow, FindSKUCol, FindFirstSkuRow, SafeDbl, etc.)
'==========================================================



'==========================================================
' PUBLIC ENTRY POINT (put this on a button)
'==========================================================
Public Sub ShowTubeSkuDrilldown()
    On Error GoTo fail

    Dim sku As String
    sku = ""

    ' Prefer active row sku if user is on Tube_Model or another list
    If Not ActiveCell Is Nothing Then
        If ActiveCell.row >= 2 Then
            sku = Trim$(CStr(Cells(ActiveCell.row, 1).Value2))
            If Len(sku) = 0 Then sku = Trim$(CStr(Cells(ActiveCell.row, 2).Value2)) ' sometimes SKU is in col B
        End If
    End If

    If Len(sku) = 0 Then
        sku = InputBox("Enter TUBE SKU to drill down:", "Tube SKU Drilldown")
    End If

    sku = Trim$(sku)
    If Len(sku) = 0 Then Exit Sub

    ' Validate that it’s a tube SKU in scope (or allow any and just show zeros)
    Dim allowed As Object: Set allowed = BuildAllowedTubeSKUSet()
    If Not allowed.exists(sku) Then
        If MsgBox("SKU is not in TubeSKUList(). Continue anyway?", vbQuestion + vbYesNo, "Not a Tube SKU") = vbNo Then Exit Sub
    End If

    BuildTubeSkuDrilldownDashboard sku
    Exit Sub

fail:
    MsgBox "ShowTubeSkuDrilldown failed: " & Err.Description, vbExclamation
End Sub

'==========================================================
' CORE BUILDER
'==========================================================
Private Sub BuildTubeSkuDrilldownDashboard(ByVal tubeSKU As String)
    Dim wb As Workbook: Set wb = ThisWorkbook

    Dim wsD As Worksheet: Set wsD = EnsureSheet(wb, SH_TUBE_DRILL)
    Dim wsBin As Worksheet: Set wsBin = wb.Worksheets(SH_BIN)
    Dim wsTube As Worksheet: Set wsTube = wb.Worksheets(SH_TUBE)
    Dim wsPF As Worksheet: Set wsPF = wb.Worksheets(SH_PROD)

    ' Reset
    With wsD
        On Error Resume Next
        .Cells.UnMerge
        On Error GoTo 0
        .Cells.Clear
        .Cells.FormatConditions.Delete
        .Cells.Font.name = "Segoe UI"
        .Cells.Font.Size = 10
    End With

    ' Title
    With wsD.Range("A1:H1")
        .Value = "Tube SKU Drilldown"
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = 13041663 ' Predator yellow
        .HorizontalAlignment = xlCenterAcrossSelection
        .RowHeight = 28
        .Borders.LineStyle = xlContinuous
        .ClearContents
        wsD.Range("A1").Value = "Tube SKU Drilldown: " & tubeSKU
    End With

    ' Snapshot cards (top)
    wsD.Range("A3").Value = "Tube SKU"
    wsD.Range("B3").Value = tubeSKU
    wsD.Range("A4").Value = "On-Hand (Bin_Inv)"
    wsD.Range("A5").Value = "First Breach"
    wsD.Range("A6").Value = "Worst End Inv"
    wsD.Range("A7").Value = "Safety (3M Avg est.)"

    wsD.Range("A3:A7").Font.Bold = True
    wsD.Range("A3:B7").Borders.LineStyle = xlContinuous

    ' Build week list from Production Forecast header row (same as your model)
    Dim weeks() As Date
    weeks = GetProductionWeeks(wsPF)

    ' Load on-hand for this tube SKU
    Dim onHand As Double
    onHand = TubeOnHand_FromBinInv(wsBin, tubeSKU)
    wsD.Range("B4").Value = onHand
    wsD.Range("B4").NumberFormat = "0"

    ' Build receipts dict for this tube SKU only (from Tube Schedule)
    Dim dRec As Object: Set dRec = CreateObject("Scripting.Dictionary")  ' key: CLng(week) -> qty
    LoadTubeReceipts_ForOneSKU wsTube, tubeSKU, dRec

    ' Build demand dict for this tube SKU only (from Production Forecast mapped)
    Dim dDem As Object: Set dDem = CreateObject("Scripting.Dictionary")  ' key: CLng(week) -> qty
    Dim dDrv As Object: Set dDrv = CreateObject("Scripting.Dictionary")  ' key: CLng(week) -> dict(prodSKU->qty)
    LoadTubeDemand_ForOneSKU wsPF, tubeSKU, dDem, dDrv

    ' Optional: add box demand (plan + utilization) if those sheets exist
    AddBoxDemand_ForOneTubeSKU wb, tubeSKU, dDem, dDrv

    ' Output table headers
    Dim topRow As Long: topRow = 10
    wsD.Range("A" & topRow & ":I" & topRow).Value = Array( _
        "Week", "Begin Inv", "Incoming (Tube Sched)", "Demand (Prod/Box)", "End Inv", _
        "Safety", "Status", "Top Driver(s)", "Notes" _
    )

    With wsD.Range("A" & topRow & ":I" & topRow)
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = 9211020 ' dark header
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' Compute a simple safety number for the drilldown:
    ' 3-month average monthly demand based on the demand we’re seeing in the horizon.
    Dim safety As Double
    safety = EstimateSafety3M_FromWeeklyDemand(weeks, dDem)
    wsD.Range("B7").Value = safety
    wsD.Range("B7").NumberFormat = "0"

    ' Build rows
    Dim rOut As Long: rOut = topRow + 1
    Dim i As Long
    Dim beginInv As Double: beginInv = onHand
    Dim worstEnd As Double: worstEnd = 9.9E+307
    Dim firstBreach As Date: firstBreach = 0

    For i = LBound(weeks) To UBound(weeks)
        Dim w As Date: w = weeks(i)
        Dim k As Long: k = CLng(w)

        Dim rec As Double: rec = 0#
        If dRec.exists(k) Then rec = CDbl(dRec(k))

        Dim dem As Double: dem = 0#
        If dDem.exists(k) Then dem = CDbl(dDem(k))

        Dim endInv As Double
        endInv = beginInv + rec - dem

        Dim st As String
        If endInv < 0 Then
            st = "STOCKOUT"
            If firstBreach = 0 Then firstBreach = w
        ElseIf endInv < safety Then
            st = "AT RISK"
            If firstBreach = 0 Then firstBreach = w
        Else
            st = "OK"
        End If

        If endInv < worstEnd Then worstEnd = endInv

        wsD.Cells(rOut, 1).Value = w
        wsD.Cells(rOut, 1).NumberFormat = "m/d/yyyy"

        wsD.Cells(rOut, 2).Value = beginInv
        wsD.Cells(rOut, 3).Value = rec
        wsD.Cells(rOut, 4).Value = dem
        wsD.Cells(rOut, 5).Value = endInv

        wsD.Cells(rOut, 6).Value = safety
        wsD.Cells(rOut, 7).Value = st

        ' Top drivers string
        If dDrv.exists(k) Then
            Dim dd As Object: Set dd = dDrv(k)
            wsD.Cells(rOut, 8).Value = TopDriversString(dd, 3)
        Else
            wsD.Cells(rOut, 8).Value = vbNullString
        End If

        rOut = rOut + 1
        beginInv = endInv
    Next i

    ' Summary fields
    If firstBreach = 0 Then
        wsD.Range("B5").Value = "No breach"
    Else
        wsD.Range("B5").Value = firstBreach
        wsD.Range("B5").NumberFormat = "m/d/yyyy"
    End If

    wsD.Range("B6").Value = worstEnd
    wsD.Range("B6").NumberFormat = "#,##0;[Red]-#,##0"

    ' Formatting table
    Dim lastOut As Long: lastOut = rOut - 1
    If lastOut >= topRow + 1 Then
        With wsD.Range("A" & topRow & ":I" & lastOut)
            .Borders.LineStyle = xlContinuous
            .VerticalAlignment = xlCenter
        End With

        wsD.Range("B" & (topRow + 1) & ":F" & lastOut).NumberFormat = "#,##0;[Red]-#,##0"

        ' Banding
        Dim rr As Long
        For rr = topRow + 1 To lastOut
            If rr Mod 2 = 0 Then wsD.Range("A" & rr & ":I" & rr).Interior.Color = 15132390 ' light grey
        Next rr

        ' Status conditional format
        With wsD.Range("G" & (topRow + 1) & ":G" & lastOut)
            .FormatConditions.Delete
            Dim fc As FormatCondition
            Set fc = .FormatConditions.Add(Type:=xlTextString, String:="STOCKOUT", TextOperator:=xlContains)
            fc.Interior.Color = RGB(255, 180, 180)
            Set fc = .FormatConditions.Add(Type:=xlTextString, String:="AT RISK", TextOperator:=xlContains)
            fc.Interior.Color = RGB(255, 230, 180)
            Set fc = .FormatConditions.Add(Type:=xlTextString, String:="OK", TextOperator:=xlContains)
            fc.Interior.Color = RGB(200, 240, 200)
        End With
    End If

    wsD.Columns.AutoFit
    wsD.Activate
    wsD.Range("A1").Select
End Sub

'==========================================================
' INVENTORY: Bin_Inv (tube SKU rows)
'==========================================================
Private Function TubeOnHand_FromBinInv(ByVal wsBin As Worksheet, ByVal tubeSKU As String) As Double
    Dim lastRow As Long
    lastRow = wsBin.Cells(wsBin.Rows.Count, BIN_COL_NAME).End(xlUp).row
    If lastRow < BIN_FIRST_ROW Then Exit Function

    Dim r As Long, sku As String
    Dim qty As Double

    For r = BIN_FIRST_ROW To lastRow
        sku = Trim$(CStr(wsBin.Cells(r, BIN_COL_NAME).Value))
        If StrComp(sku, tubeSKU, vbTextCompare) = 0 Then
            qty = SafeDbl(wsBin.Cells(r, BIN_COL_QTY).Value)
            TubeOnHand_FromBinInv = TubeOnHand_FromBinInv + qty
        End If
    Next r
End Function

'==========================================================
' RECEIPTS: Tube Schedule (incoming week = shipWeek + RECEIPT_LAG_WEEKS)
'==========================================================
Private Sub LoadTubeReceipts_ForOneSKU(ByVal wsTube As Worksheet, ByVal tubeSKU As String, ByVal dRec As Object)
    Dim dateRow As Long: dateRow = FindDateHeaderRow(wsTube, 25)
    If dateRow = 0 Then Exit Sub

    Dim skuCol As Long: skuCol = FindSKUCol(wsTube, dateRow)
    Dim lastCol As Long: lastCol = wsTube.Cells(dateRow, wsTube.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = wsTube.Cells(wsTube.Rows.Count, skuCol).End(xlUp).row

    Dim r As Long, sku As String
    Dim rowSku As Long: rowSku = 0

    For r = dateRow + 1 To lastRow
        sku = Trim$(CStr(wsTube.Cells(r, skuCol).Value))
        If StrComp(sku, tubeSKU, vbTextCompare) = 0 Then
            rowSku = r
            Exit For
        End If
    Next r
    If rowSku = 0 Then Exit Sub

    Dim c As Long, vDate As Variant
    Dim shipWeek As Date, recvWeek As Date
    Dim qty As Double, k As Long

    For c = skuCol + 1 To lastCol
        vDate = wsTube.Cells(dateRow, c).Value
        If IsDate(vDate) Then
            shipWeek = CDate(vDate)
            recvWeek = DateAdd("ww", RECEIPT_LAG_WEEKS, shipWeek)

            qty = SafeDbl(wsTube.Cells(rowSku, c).Value)
            If qty <> 0 Then
                k = CLng(recvWeek)
                If dRec.exists(k) Then
                    dRec(k) = CDbl(dRec(k)) + qty
                Else
                    dRec.Add k, qty
                End If
            End If
        End If
    Next c
End Sub

'==========================================================
' DEMAND: Production Forecast -> Belor -> Tube SKU, bucketed to floorWeek
'==========================================================
Private Sub LoadTubeDemand_ForOneSKU(ByVal wsPF As Worksheet, _
                                     ByVal tubeSKU As String, _
                                     ByVal dDem As Object, _
                                     ByVal dDrv As Object)

    Dim dateRow As Long: dateRow = FindDateHeaderRow(wsPF, 25)
    If dateRow = 0 Then Exit Sub

    Dim skuCol As Long: skuCol = FindSKUCol(wsPF, dateRow)
    Dim firstRow As Long: firstRow = FindFirstSkuRow(wsPF, skuCol, dateRow, Nothing)
    If firstRow = 0 Then Exit Sub

    Dim lastCol As Long: lastCol = wsPF.Cells(dateRow, wsPF.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = wsPF.Cells(wsPF.Rows.Count, skuCol).End(xlUp).row
    If lastCol <= skuCol Or lastRow < firstRow Then Exit Sub

    Dim c As Long, r As Long
    Dim vDate As Variant, completeWeek As Date, floorWeek As Date
    Dim prodSKU As String, qty As Double
    Dim belor As String, mappedTube As String
    Dim k As Long

    For c = skuCol + 1 To lastCol
        vDate = wsPF.Cells(dateRow, c).Value
        If IsDate(vDate) Then
            completeWeek = CDate(vDate)
            floorWeek = DateAdd("ww", -PROD_SHIFT_WEEKS, completeWeek)
            k = CLng(floorWeek)

            For r = firstRow To lastRow
                prodSKU = Trim$(CStr(wsPF.Cells(r, skuCol).Value))
                If Len(prodSKU) > 0 Then
                    qty = SafeDbl(wsPF.Cells(r, c).Value)
                    If qty <> 0 Then
                        belor = GetBelorFromSKU(prodSKU)
                        If Len(belor) > 0 Then
                            mappedTube = TubeSKUFromBelor(belor)
                            If StrComp(mappedTube, tubeSKU, vbTextCompare) = 0 Then
                                ' demand
                                If dDem.exists(k) Then
                                    dDem(k) = CDbl(dDem(k)) + qty
                                Else
                                    dDem.Add k, qty
                                End If

                                ' drivers
                                Dim dd As Object
                                If dDrv.exists(k) Then
                                    Set dd = dDrv(k)
                                Else
                                    Set dd = CreateObject("Scripting.Dictionary")
                                    dDrv.Add k, dd
                                End If

                                If dd.exists(prodSKU) Then
                                    dd(prodSKU) = CDbl(dd(prodSKU)) + qty
                                Else
                                    dd.Add prodSKU, qty
                                End If
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    Next c
End Sub

'==========================================================
' OPTIONAL: Add Box demand into the drilldown demand dicts
'==========================================================
Private Sub AddBoxDemand_ForOneTubeSKU(ByVal wb As Workbook, _
                                      ByVal tubeSKU As String, _
                                      ByVal dDem As Object, _
                                      ByVal dDrv As Object)

    ' If you don’t want box logic here, just delete this whole sub.

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(SH_BOX_PLAN)
    On Error GoTo 0
    If Not ws Is Nothing Then
        AddBoxDemand_FromSheet ws, tubeSKU, dDem, dDrv, "BOX_PLAN"
    End If
    Set ws = Nothing

    On Error Resume Next
    Set ws = wb.Worksheets(SH_BOX_UTIL)
    On Error GoTo 0
    If Not ws Is Nothing Then
        AddBoxDemand_FromSheet ws, tubeSKU, dDem, dDrv, "BOX_UTIL"
    End If
End Sub

Private Sub AddBoxDemand_FromSheet(ByVal ws As Worksheet, _
                                  ByVal tubeSKU As String, _
                                  ByVal dDem As Object, _
                                  ByVal dDrv As Object, _
                                  ByVal tag As String)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim colSKU As Long, colDate As Long, colQty As Long
    colSKU = FindHeaderCol(ws, 1, Array("Finished SKU", "SKU", "Item"))
    colDate = FindHeaderCol(ws, 1, Array("Production Date", "Prod Date", "Date"))
    colQty = FindHeaderCol(ws, 1, Array("Qty", "QTY", "Quantity"))

    If colSKU = 0 Then colSKU = 1
    If colDate = 0 Then colDate = 4

    Dim r As Long
    For r = 2 To lastRow
        Dim finishedSku As String: finishedSku = Trim$(CStr(ws.Cells(r, colSKU).Value))
        If Len(finishedSku) = 0 Then GoTo NextR

        Dim vDate As Variant: vDate = ws.Cells(r, colDate).Value
        If Not IsDate(vDate) Then GoTo NextR

        Dim wk As Date: wk = CDate(vDate)
        Dim k As Long: k = CLng(wk)

        Dim belor As String: belor = GetBelorFromSKU(finishedSku)
        If Len(belor) = 0 Then GoTo NextR

        Dim mappedTube As String: mappedTube = TubeSKUFromBelor(belor)
        If StrComp(mappedTube, tubeSKU, vbTextCompare) <> 0 Then GoTo NextR

        Dim qty As Double: qty = 1#
        If colQty > 0 Then
            If IsNumeric(ws.Cells(r, colQty).Value) Then
                If CDbl(ws.Cells(r, colQty).Value) <> 0 Then qty = CDbl(ws.Cells(r, colQty).Value)
            End If
        End If
        If qty = 0 Then GoTo NextR

        If dDem.exists(k) Then
            dDem(k) = CDbl(dDem(k)) + qty
        Else
            dDem.Add k, qty
        End If

        Dim dd As Object
        If dDrv.exists(k) Then
            Set dd = dDrv(k)
        Else
            Set dd = CreateObject("Scripting.Dictionary")
            dDrv.Add k, dd
        End If

        Dim drvLabel As String: drvLabel = tag & ": " & finishedSku
        If dd.exists(drvLabel) Then
            dd(drvLabel) = CDbl(dd(drvLabel)) + qty
        Else
            dd.Add drvLabel, qty
        End If

NextR:
    Next r
End Sub



'==========================================================
' SAFETY ESTIMATE: simple 3-month avg from weekly demand
'==========================================================
Private Function EstimateSafety3M_FromWeeklyDemand(ByRef weeks() As Date, ByVal dDem As Object) As Double
    ' Convert weekly demand into monthly totals then average last 3 non-zero months we have.
    Dim dMon As Object: Set dMon = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(weeks) To UBound(weeks)
        Dim w As Date: w = weeks(i)
        Dim k As Long: k = CLng(w)

        Dim dem As Double: dem = 0#
        If dDem.exists(k) Then dem = CDbl(dDem(k))
        If dem <> 0 Then
            Dim mStart As Date: mStart = DateSerial(Year(w), Month(w), 1)
            Dim mk As String: mk = CStr(mStart)
            If dMon.exists(mk) Then
                dMon(mk) = CDbl(dMon(mk)) + dem
            Else
                dMon.Add mk, dem
            End If
        End If
    Next i

    If dMon.Count = 0 Then
        EstimateSafety3M_FromWeeklyDemand = 0#
        Exit Function
    End If

    ' Take up to 3 most recent months in weeks() order
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim vals() As Double
    ReDim vals(1 To 3)
    Dim n As Long: n = 0

    For i = UBound(weeks) To LBound(weeks) Step -1
        Dim m As Date: m = DateSerial(Year(weeks(i)), Month(weeks(i)), 1)
        Dim mk2 As String: mk2 = CStr(m)
        If Not seen.exists(mk2) Then
            seen.Add mk2, True
            If dMon.exists(mk2) Then
                n = n + 1
                vals(n) = CDbl(dMon(mk2))
                If n = 3 Then Exit For
            End If
        End If
    Next i

    If n = 0 Then
        EstimateSafety3M_FromWeeklyDemand = 0#
    Else
        Dim s As Double: s = 0#
        For i = 1 To n: s = s + vals(i): Next i
        EstimateSafety3M_FromWeeklyDemand = s / n
    End If
End Function
