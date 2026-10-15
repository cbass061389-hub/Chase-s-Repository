Option Explicit

'===========================
' Sheet names (change here)
'===========================
Private Const SH_INPUT             As String = "Production Plan"
Private Const SH_PROD_TABLE        As String = "Production Table"
Private Const SH_PROD_TABLE_BACKUP As String = "Production Table Backup"
Private Const SH_ANALYSIS          As String = "Production Analysis"
Private Const SH_JAX               As String = "JAX"
Private Const SH_8W_VIEW           As String = "8 weeks view"
Private Const SH_INTERNAL_FORECAST As String = "Internal Forecast"
Private Const SH_FLOOR             As String = "REVO Floor"
Private Const SH_RELEASE_LOG        As String = "Release Log"
' Release Log columns (based on your current layout)
Private Const RL_COL_DATE As Long = 1      'A
Private Const RL_COL_SKU As Long = 2       'B
Private Const RL_COL_QTY_RELEASED As Long = 4 'D
Private Const RL_COL_QTY_REWORK As Long = 6   'F
Private Const RL_COL_QTY_BGRADE As Long = 7   'G
Private Const RL_COL_QTY_REJECT As Long = 8   'H

'========================
' CONFIG  change if names differ
'========================
Private Const SH_SRC As String = "Production Plan"  ' 3-column list
' Stores the latest forecast horizon cutoff date detected from Production Forecast row 2
Public gProdCutoff As Date


' Source table columns (3-col list)
Private Const SRC_COL_SKU As Long = 1      ' A = Shaft
Private Const SRC_COL_DATE As Long = 2     ' B = Production Week (date)
Private Const SRC_COL_QTY As Long = 3      ' C = Forecast Qty

' Destination (matrix) layout
Private Const DST_COL_SKU As Long = 1        ' A = Shaft/Item names
Private Const DST_ROW_DATES As Long = 2      ' row with actual dates
Private Const DST_ROW_TOTAL As Long = 3      ' "Weekly Totals"
Private Const DST_FIRST_DATA_ROW As Long = 4 ' first SKU row

'========================
' CONFIG  sheet/column mapping
'========================
Private Const SH_BOM As String = "BOM"                      ' SKU -> Component map
Private Const BOM_COL_SKU As Long = 1                       ' A
Private Const BOM_COL_COMP As Long = 2                      ' B
Private Const BOM_COL_QTY As Long = 3                       ' C (Qty per); default 1 if blank/0


Private Const MX_ROW_DATES As Long = 2                      ' dates in row 2
Private Const MX_ROW_FIRST As Long = 4                      ' first SKU row
Private Const MX_COL_SKU As Long = 1                        ' A

Private Const SH_INV As String = "Current Inventory"        ' component O/H
Private Const INV_COL_SKU As Long = 1                       ' A
Private Const INV_COL_QTY As Long = 2                       ' B  (O/H)

Private Const SH_PO As String = "Items on PO"               ' receipts (by component SKU)
Private Const PO_COL_SKU As Long = 1                        ' A
Private Const PO_COL_DATE As Long = 2                       ' B  (Expected Receipt Date)
Private Const PO_COL_QTY As Long = 3                        ' C  (Quantity Remaining)

Private Const SH_OUT As String = "Component Health"         ' output

Public Sub RunWeeklyREVOUpdate_Safe()
    Dim s As String
    Dim why As String
    Dim didAppendV As Variant, wasError As Boolean

    '---------------------------
    ' STEP: AppendProductionTable
    '---------------------------
    s = "AppendProductionTable"
    On Error GoTo StepError
    ' Try as Function first (to avoid type mismatch if it's actually a Sub)
    On Error Resume Next
    didAppendV = AppendProductionTable()
    wasError = (Err.Number <> 0)
    Err.Clear
    On Error GoTo StepError

    If wasError Then
        ' It's a Sub  just call it directly
        AppendProductionTable
    Else
        ' It returned something  if Boolean and False, log the skip
        If VarType(didAppendV) = vbBoolean Then
            If Not CBool(didAppendV) Then Debug.Print "Skipped Append"
        End If
    End If
    On Error GoTo 0
On Error GoTo 0

' Normalize ONLY the horizon dates (>= forecast cutoff we just detected)
If gProdCutoff <> 0 Then
    ProdTable_NormalizeDates_FromCutoff gProdCutoff
Else
    Debug.Print "!! gProdCutoff was not set (AppendProductionTable may have exited early). Skipping normalization."
End If



    '---------------------------
    ' STEP: BuildProdAnalysisBVP
    '---------------------------
    s = "BuildProdAnalysisBVP"
    On Error GoTo StepError
    BuildProdAnalysisBVP
    On Error GoTo 0

    '---------------------------
    ' STEP: BuildProdAnalysisWVP
    '---------------------------
    s = "BuildProdAnalysisWVP"
    On Error GoTo StepError
    BuildProdAnalysisWVP
    On Error GoTo 0

    '---------------------------
    ' STEP: BuildProdActuals
    '---------------------------
    s = "BuildProdActuals"
    On Error GoTo StepError
    BuildProdActuals
    On Error GoTo 0

    '---------------------------
    ' STEP: BuildProdRejects
    '---------------------------
    s = "BuildProdRejects"
    On Error GoTo StepError
    BuildProdRejects
    On Error GoTo 0

 



  
    
  

    Exit Sub

StepError:
    ' Show exactly which step failed and the error details
    MsgBox "Error in step: " & s & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, _
           vbCritical, "Weekly REVO  Step Failure"
    ' Optional: also print to Immediate Window
    Debug.Print "!! Failure at step: "; s; " | Err"; Err.Number; Err.Description
End Sub




'========================================================
' STEP 1  Append Production Table from Production Forecast matrix
'   - Cutoff = earliest date on Production Forecast row 2 (B:...)
'   - Backs up + purges Production Table horizon (Prod Date >= cutoff)
'   - Appends non-zero qty from Production Forecast matrix to Production Table A:E
'   - Safe if sheets missing / no data
'========================================================
Public Function AppendProductionTable() As Boolean
    On Error GoTo ErrHandler
    Debug.Print "=== AppendProductionTable START ==="

    ' --- Source is Production Forecast ---
    If Not SheetExists(SH_SRC) Then
        Debug.Print "!! Missing Production Forecast sheet"
        AppendProductionTable = False
        Exit Function
    End If

    ' --- Ensure Production Table exists ---
    If Not SheetExists(SH_PROD_TABLE) Then
        Debug.Print "!! Missing Production Table sheet - creating."
        Dim wsNew As Worksheet
        Set wsNew = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsNew.name = SH_PROD_TABLE
        wsNew.Range("A1:E1").Value = Array("SKU", "Production Date", "Forecast Qty", "Month", "Year")
    End If

    ' --- Ensure backup sheet exists (append-only safety net) ---
    If Not SheetExists(SH_PROD_TABLE_BACKUP) Then
        Dim wsB As Worksheet
        Set wsB = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsB.name = SH_PROD_TABLE_BACKUP
        wsB.Range("A1:F1").Value = Array("SnapshotTS", "SKU", "Production Date", "Forecast Qty", "Month", "Year")
    End If

    Dim wsSrc As Worksheet: Set wsSrc = GetSheet(SH_SRC)                 ' Production Forecast
    Dim wsOut As Worksheet: Set wsOut = GetSheet(SH_PROD_TABLE)          ' Production Table
    Dim wsBak As Worksheet: Set wsBak = GetSheet(SH_PROD_TABLE_BACKUP)   ' Backup

    '=========================================================
    ' Determine last date column on row 2 (headers)
    '=========================================================
    Dim lastColSrc As Long
    lastColSrc = wsSrc.Cells(2, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastColSrc < 2 Then
        Debug.Print "!! Production Forecast has no date columns."
        AppendProductionTable = False
        Exit Function
    End If

    '=========================================================
    ' CUT OFF = earliest valid date on Production Forecast row 2 (B:...lastCol)
    '=========================================================
    Dim hdr As Variant
    hdr = wsSrc.Range(wsSrc.Cells(2, 2), wsSrc.Cells(2, lastColSrc)).Value2   ' 1 x N

    Dim cutoff As Date, foundDate As Boolean
    foundDate = False

    Dim c As Long
    Dim tmpD As Date

    For c = 1 To UBound(hdr, 2)
        If CoerceDate(hdr(1, c), tmpD) Then
            If Not foundDate Then
                cutoff = tmpD
                foundDate = True
            ElseIf tmpD < cutoff Then
                cutoff = tmpD
            End If
        End If
    Next c

    If Not foundDate Then
        Debug.Print "!! Could not find any valid date headers on Production Forecast row 2."
        AppendProductionTable = False
        Exit Function
    End If

    Debug.Print "Cutoff date (earliest forecast week) = " & Format$(cutoff, "mm/dd/yyyy")
    gProdCutoff = cutoff

    '=========================================================
    ' BACKUP: append current horizon rows (Prod Date >= cutoff)
    '=========================================================
    BackupProdTableHorizon wsOut, wsBak, cutoff

    '=========================================================
    ' PURGE: remove rows in Production Table where Prod Date >= cutoff
    ' (Preserves history before cutoff forever.)
    ' Fast approach: read A:E into array, keep only rows with date < cutoff, rewrite.
    '=========================================================
    Dim lastRowOut As Long
    lastRowOut = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).row

    If lastRowOut >= 2 Then
        Dim arrOut As Variant
        arrOut = wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(lastRowOut, 5)).Value2  ' rows x 5

        Dim keepCount As Long, r As Long
        keepCount = 0

        ' First pass count keepers
        For r = 1 To UBound(arrOut, 1)
            If CoerceDate(arrOut(r, 2), tmpD) Then
                If tmpD < cutoff Then keepCount = keepCount + 1
            Else
                ' If Production Date isn't a date, keep it (safer than deleting unknowns)
                keepCount = keepCount + 1
            End If
        Next r

        ' Clear old data (rows 2 down)
        wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(lastRowOut, 5)).ClearContents

        If keepCount > 0 Then
            Dim keepArr() As Variant
            ReDim keepArr(1 To keepCount, 1 To 5)

            Dim k As Long: k = 0
            For r = 1 To UBound(arrOut, 1)
                Dim keepIt As Boolean
                keepIt = True

                If CoerceDate(arrOut(r, 2), tmpD) Then
                    If tmpD >= cutoff Then keepIt = False
                End If

                If keepIt Then
                    k = k + 1
                    keepArr(k, 1) = arrOut(r, 1)
                    keepArr(k, 2) = arrOut(r, 2)
                    keepArr(k, 3) = arrOut(r, 3)
                    keepArr(k, 4) = arrOut(r, 4)
                    keepArr(k, 5) = arrOut(r, 5)
                End If
            Next r

            wsOut.Cells(2, 1).Resize(keepCount, 5).Value = keepArr
        End If
    End If

    '=========================================================
    ' BUILD BUFFER from Production Forecast grid
    '   - SKU in col A starting row 5
    '   - Week dates in row 2 starting col B
    '=========================================================
    Dim lastRowSrc As Long
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row
    Debug.Print "Prod Forecast lastRow=" & lastRowSrc & " lastCol=" & lastColSrc

    If lastRowSrc < 5 Then
        Debug.Print "!! No SKU rows found on Production Forecast."
        AppendProductionTable = True
        Exit Function
    End If

    ' Load SKUs + grid in one shot
    Dim arrSku As Variant, arrGrid As Variant
    arrSku = wsSrc.Range(wsSrc.Cells(5, 1), wsSrc.Cells(lastRowSrc, 1)).Value2              ' rows x 1
    arrGrid = wsSrc.Range(wsSrc.Cells(5, 2), wsSrc.Cells(lastRowSrc, lastColSrc)).Value2   ' rows x cols

    ' Pre-coerce header dates into an array (so we don’t re-convert repeatedly)
    Dim hdrDates() As Date, hdrIsDate() As Boolean
    ReDim hdrDates(1 To UBound(hdr, 2))
    ReDim hdrIsDate(1 To UBound(hdr, 2))

    For c = 1 To UBound(hdr, 2)
        hdrIsDate(c) = CoerceDate(hdr(1, c), hdrDates(c))
    Next c

    ' First pass: count how many rows we will append
    Dim addCount As Long
    addCount = 0

    Dim sku As String
    Dim qty As Variant
    Dim rr As Long

    For rr = 1 To UBound(arrSku, 1)
        sku = Trim$(CStr(arrSku(rr, 1)))
        If Len(sku) > 0 Then
            For c = 1 To UBound(arrGrid, 2)
                If hdrIsDate(c) Then
                    If hdrDates(c) >= cutoff Then
                        qty = arrGrid(rr, c)
                        If IsNumeric(qty) Then
                            If NzD(qty) <> 0 Then addCount = addCount + 1
                        End If
                    End If
                End If
            Next c
        End If
    Next rr

    Debug.Print "Buffer count=" & addCount

    If addCount = 0 Then
        Debug.Print "   Nothing new to append."
        AppendProductionTable = True
        Exit Function
    End If

    ' Second pass: build outData
    Dim outData() As Variant
    ReDim outData(1 To addCount, 1 To 5)

    Dim i As Long: i = 0
    For rr = 1 To UBound(arrSku, 1)
        sku = Trim$(CStr(arrSku(rr, 1)))
        If Len(sku) > 0 Then
            For c = 1 To UBound(arrGrid, 2)
                If hdrIsDate(c) Then
                    tmpD = hdrDates(c)
                    If tmpD >= cutoff Then
                        qty = arrGrid(rr, c)
                        If IsNumeric(qty) Then
                            Dim q As Double
                            q = NzD(qty)
                            If q <> 0 Then
                                i = i + 1
                                outData(i, 1) = sku
                                outData(i, 2) = tmpD
                                outData(i, 3) = q
                                outData(i, 4) = Month(tmpD)
                                outData(i, 5) = Year(tmpD)
                            End If
                        End If
                    End If
                End If
            Next c
        End If
    Next rr

    '=========================================================
    ' APPEND to Production Table (A:E)
    '=========================================================
    lastRowOut = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).row
    Dim writeRow As Long
    If lastRowOut < 2 Then
        writeRow = 2
    Else
        writeRow = lastRowOut + 1
    End If

    wsOut.Cells(writeRow, 1).Resize(addCount, 5).Value = outData
    Debug.Print "Wrote " & addCount & " rows to Production Table"

    ' Deduplicate horizon (keep last occurrence per SKU|Date) and sort
    DedupHorizon_ProdTable cutoff

    AppendProductionTable = True
    Debug.Print "=== AppendProductionTable COMPLETE ==="
    Exit Function

ErrHandler:
    Debug.Print "!! ERROR in AppendProductionTable: " & Err.Number & " - " & Err.Description
    AppendProductionTable = False
End Function

'------------------------------------------------------------
' CoerceDate: robustly interprets Excel date serials and date-like values
'------------------------------------------------------------
Private Function CoerceDate(ByVal v As Variant, ByRef outD As Date) As Boolean
    On Error GoTo fail
    If IsError(v) Then GoTo fail

    If IsDate(v) Then
        outD = DateValue(CDate(v))
        CoerceDate = True
        Exit Function
    End If

    If IsNumeric(v) Then
        outD = DateValue(CDate(CDbl(v)))
        CoerceDate = True
        Exit Function
    End If

fail:
    CoerceDate = False
End Function

Private Sub BackupProdTableHorizon(ByVal wsOut As Worksheet, ByVal wsBak As Worksheet, ByVal cutoff As Date)
    On Error GoTo EH

    Dim lastRow As Long: lastRow = wsOut.Cells(wsOut.Rows.Count, "B").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim r As Long, snap As Date
    snap = Now

    ' Count rows to back up
    Dim cnt As Long: cnt = 0
    For r = 2 To lastRow
        If IsDate(wsOut.Cells(r, "B").Value) Then
            If CDate(wsOut.Cells(r, "B").Value) >= cutoff Then cnt = cnt + 1
        End If
    Next r
    If cnt = 0 Then Exit Sub

    Dim outArr() As Variant: ReDim outArr(1 To cnt, 1 To 6)
    Dim i As Long: i = 0
    For r = 2 To lastRow
        If IsDate(wsOut.Cells(r, "B").Value) Then
            Dim dd As Date: dd = CDate(wsOut.Cells(r, "B").Value)
            If dd >= cutoff Then
                i = i + 1
                outArr(i, 1) = snap
                outArr(i, 2) = wsOut.Cells(r, "A").Value
                outArr(i, 3) = dd
                outArr(i, 4) = wsOut.Cells(r, "C").Value
                outArr(i, 5) = Month(dd)
                outArr(i, 6) = Year(dd)
            End If
        End If
    Next r

    Dim lr As Long: lr = wsBak.Cells(wsBak.Rows.Count, "A").End(xlUp).row
    If lr < 2 Then lr = 1
    wsBak.Cells(lr + 1, 1).Resize(cnt, 6).Value = outArr
    Exit Sub

EH:
    Debug.Print "!! BackupProdTableHorizon failed: " & Err.Description
End Sub

Public Sub GetDateRange(ws As Worksheet, cellRef As String, ByRef startDate As Date, ByRef endDate As Date)
    Dim raw As String: raw = Trim(ws.Range(cellRef).Value)
    If Len(raw) = 0 Then startDate = 0: endDate = 0: Exit Sub

    Dim parts() As String
    parts = Split(raw, "-")
    
    If UBound(parts) < 1 Then startDate = 0: endDate = 0: Exit Sub
    
    On Error Resume Next
    startDate = CDate(Trim(parts(0)))
    endDate = CDate(Trim(parts(1)))
    On Error GoTo 0
End Sub
'========================================================
' STEP 2c  Build Actuals (from JAX H:J)
'   H = Date, I = SKU, J = Qty Released
'   Writes to Production Analysis:
'     BVP rows  520 -> cols D (WTD), J (MTD), P (YTD)
'     WVP rows 2540 -> cols D (WTD), J (MTD), P (YTD)
'========================================================
Public Sub BuildProdActuals()
    ' Actuals (Qty Released) from Release Log, split BVP/WVP by Belor bucket
    If Not (SheetExists(SH_RELEASE_LOG) And SheetExists(SH_ANALYSIS)) Then Exit Sub

    Dim wsR As Worksheet: Set wsR = GetSheet(SH_RELEASE_LOG)
    Dim wsO As Worksheet: Set wsO = GetSheet(SH_ANALYSIS)

    Dim wkStart As Date, wkEnd As Date
    If Not SafeGetDates(wsO, wkStart, wkEnd) Then Exit Sub

    Dim mStart As Date: mStart = DateSerial(Year(wkEnd), Month(wkEnd), 1)
    Dim yStart As Date: yStart = DateSerial(Year(wkEnd), 1, 1)

    ' Keys displayed (and their row)
    Dim keyRowB As Object: Set keyRowB = CreateObject("Scripting.Dictionary") ' rows 5..20
    Dim keyRowW As Object: Set keyRowW = CreateObject("Scripting.Dictionary") ' rows 25..40
    keyRowB.CompareMode = vbTextCompare
    keyRowW.CompareMode = vbTextCompare

    Dim i As Long, key As String
    For i = 5 To 20
        key = Trim$(CStr(wsO.Cells(i, 2).Value))
        If key <> "" Then keyRowB(key) = i
    Next i
    For i = 25 To 40
        key = Trim$(CStr(wsO.Cells(i, 2).Value))
        If key <> "" Then keyRowW(key) = i
    Next i

    ' Clear actuals cells (D/J/P for both blocks)
    wsO.Range("D5:D20,J5:J20,P5:P20,D25:D40,J25:J40,P25:P40").ClearContents

    Dim sumW_B() As Double, sumM_B() As Double, sumY_B() As Double
    Dim sumW_W() As Double, sumM_W() As Double, sumY_W() As Double
    ReDim sumW_B(5 To 20), sumM_B(5 To 20), sumY_B(5 To 20)
    ReDim sumW_W(25 To 40), sumM_W(25 To 40), sumY_W(25 To 40)

    Dim lastR As Long: lastR = wsR.Cells(wsR.Rows.Count, RL_COL_DATE).End(xlUp).row
    If lastR < 2 Then GoTo writeOut

    Dim arr As Variant: arr = wsR.Range(wsR.Cells(2, 1), wsR.Cells(lastR, RL_COL_QTY_REJECT)).Value
    Dim r As Long, sku As String, d As Variant, dd As Date, q As Double, rowIx As Long, bucket As String, lane As String

    For r = LBound(arr, 1) To UBound(arr, 1)
        d = arr(r, RL_COL_DATE)
        sku = Trim$(CStr(arr(r, RL_COL_SKU)))
        q = NzD(arr(r, RL_COL_QTY_RELEASED))

        If sku <> "" And IsDate(d) And q <> 0 Then
            dd = CDate(d)
            lane = LaneFromSKU(sku)
            bucket = AnalysisKeyFromSKU(sku)

            If bucket = "" Then GoTo NextR

            If lane = "BVP" Then
                If keyRowB.exists(bucket) Then
                    rowIx = CLng(keyRowB(bucket))
                    If dd >= wkStart And dd <= wkEnd Then sumW_B(rowIx) = sumW_B(rowIx) + q
                    If dd >= mStart And dd <= wkEnd Then sumM_B(rowIx) = sumM_B(rowIx) + q
                    If dd >= yStart And dd <= wkEnd Then sumY_B(rowIx) = sumY_B(rowIx) + q
                End If
            ElseIf lane = "WVP" Then
                If keyRowW.exists(bucket) Then
                    rowIx = CLng(keyRowW(bucket))
                    If dd >= wkStart And dd <= wkEnd Then sumW_W(rowIx) = sumW_W(rowIx) + q
                    If dd >= mStart And dd <= wkEnd Then sumM_W(rowIx) = sumM_W(rowIx) + q
                    If dd >= yStart And dd <= wkEnd Then sumY_W(rowIx) = sumY_W(rowIx) + q
                End If
            End If
        End If
NextR:
    Next r

writeOut:
    For i = 5 To 20
        If Not IsEmpty(wsO.Cells(i, 2).Value) Then
            wsO.Cells(i, "D").Value = sumW_B(i)
            wsO.Cells(i, "J").Value = sumM_B(i)
            wsO.Cells(i, "P").Value = sumY_B(i)
        End If
    Next i

    For i = 25 To 40
        If Not IsEmpty(wsO.Cells(i, 2).Value) Then
            wsO.Cells(i, "D").Value = sumW_W(i)
            wsO.Cells(i, "J").Value = sumM_W(i)
            wsO.Cells(i, "P").Value = sumY_W(i)
        End If
    Next i
End Sub


Public Sub BuildProdRejects()
    ' Quality (Rework + B-Grade + Reject) from Release Log, split BVP/WVP by Belor bucket
    Const COL_G As Long = 7     ' G  (WTD Quality)
    Const COL_M As Long = 13    ' M  (MTD Quality)
    Const COL_S As Long = 19    ' S  (YTD Quality)

    If Not (SheetExists(SH_RELEASE_LOG) And SheetExists(SH_ANALYSIS)) Then Exit Sub

    Dim wsR As Worksheet: Set wsR = GetSheet(SH_RELEASE_LOG)
    Dim wsO As Worksheet: Set wsO = GetSheet(SH_ANALYSIS)

    Dim wkStart As Date, wkEnd As Date
    If Not SafeGetDates(wsO, wkStart, wkEnd) Then Exit Sub

    Dim mStart As Date: mStart = DateSerial(Year(wkEnd), Month(wkEnd), 1)
    Dim yStart As Date: yStart = DateSerial(Year(wkEnd), 1, 1)

    ' Keys displayed (and their row)
    Dim keyRowB As Object: Set keyRowB = CreateObject("Scripting.Dictionary") ' rows 5..20
    Dim keyRowW As Object: Set keyRowW = CreateObject("Scripting.Dictionary") ' rows 25..40
    keyRowB.CompareMode = vbTextCompare
    keyRowW.CompareMode = vbTextCompare

    Dim i As Long, key As String
    For i = 5 To 20
        key = Trim$(CStr(wsO.Cells(i, 2).Value))
        If key <> "" Then keyRowB(key) = i
    Next i
    For i = 25 To 40
        key = Trim$(CStr(wsO.Cells(i, 2).Value))
        If key <> "" Then keyRowW(key) = i
    Next i

    ' Clear quality cells (G/M/S for both blocks)
    wsO.Range("G5:G20,M5:M20,S5:S20,G25:G40,M25:M40,S25:S40").ClearContents

    Dim sumW_B() As Double, sumM_B() As Double, sumY_B() As Double
    Dim sumW_W() As Double, sumM_W() As Double, sumY_W() As Double
    ReDim sumW_B(5 To 20), sumM_B(5 To 20), sumY_B(5 To 20)
    ReDim sumW_W(25 To 40), sumM_W(25 To 40), sumY_W(25 To 40)

    Dim lastR As Long: lastR = wsR.Cells(wsR.Rows.Count, RL_COL_DATE).End(xlUp).row
    If lastR < 2 Then GoTo writeOut

    Dim arr As Variant: arr = wsR.Range(wsR.Cells(2, 1), wsR.Cells(lastR, RL_COL_QTY_REJECT)).Value
    Dim r As Long, sku As String, d As Variant, dd As Date, q As Double, rowIx As Long, bucket As String, lane As String

    For r = LBound(arr, 1) To UBound(arr, 1)
        d = arr(r, RL_COL_DATE)
        sku = Trim$(CStr(arr(r, RL_COL_SKU)))
        If sku = "" Or Not IsDate(d) Then GoTo NextR

        q = NzD(arr(r, RL_COL_QTY_REWORK)) + NzD(arr(r, RL_COL_QTY_BGRADE)) + NzD(arr(r, RL_COL_QTY_REJECT))
        If q = 0 Then GoTo NextR

        dd = CDate(d)
        lane = LaneFromSKU(sku)
       bucket = AnalysisKeyFromSKU(sku)

        If bucket = "" Then GoTo NextR

        If lane = "BVP" Then
            If keyRowB.exists(bucket) Then
                rowIx = CLng(keyRowB(bucket))
                If dd >= wkStart And dd <= wkEnd Then sumW_B(rowIx) = sumW_B(rowIx) + q
                If dd >= mStart And dd <= wkEnd Then sumM_B(rowIx) = sumM_B(rowIx) + q
                If dd >= yStart And dd <= wkEnd Then sumY_B(rowIx) = sumY_B(rowIx) + q
            End If
        ElseIf lane = "WVP" Then
            If keyRowW.exists(bucket) Then
                rowIx = CLng(keyRowW(bucket))
                If dd >= wkStart And dd <= wkEnd Then sumW_W(rowIx) = sumW_W(rowIx) + q
                If dd >= mStart And dd <= wkEnd Then sumM_W(rowIx) = sumM_W(rowIx) + q
                If dd >= yStart And dd <= wkEnd Then sumY_W(rowIx) = sumY_W(rowIx) + q
            End If
        End If
NextR:
    Next r

writeOut:
    For i = 5 To 20
        If Not IsEmpty(wsO.Cells(i, 2).Value) Then
            wsO.Cells(i, COL_G).Value = sumW_B(i)
            wsO.Cells(i, COL_M).Value = sumM_B(i)
            wsO.Cells(i, COL_S).Value = sumY_B(i)
        End If
    Next i

    For i = 25 To 40
        If Not IsEmpty(wsO.Cells(i, 2).Value) Then
            wsO.Cells(i, COL_G).Value = sumW_W(i)
            wsO.Cells(i, COL_M).Value = sumM_W(i)
            wsO.Cells(i, COL_S).Value = sumY_W(i)
        End If
    Next i
End Sub





Public Sub BuildProdAnalysisBVP()
    If Not (SheetExists(SH_PROD_TABLE) And SheetExists(SH_ANALYSIS) And SheetExists(SH_SRC)) Then Exit Sub

    Dim wsP As Worksheet: Set wsP = GetSheet(SH_PROD_TABLE)
    Dim wsO As Worksheet: Set wsO = GetSheet(SH_ANALYSIS)

    ' FORECAST window: 4 weeks from Production Forecast headers
   Dim wkStart As Date, wkEnd As Date
If Not SafeGetDates(wsO, wkStart, wkEnd) Then Exit Sub


    Dim mStart As Date: mStart = DateSerial(Year(wkEnd), Month(wkEnd), 1)
    Dim yStart As Date: yStart = DateSerial(Year(wkEnd), 1, 1)
    Dim yEnd   As Date: yEnd = DateSerial(Year(wkEnd), 12, 31)

    Dim liveKeys As Object: Set liveKeys = CreateObject("Scripting.Dictionary")
    liveKeys.CompareMode = vbTextCompare

    Dim i As Long, key As String
    For i = 5 To 20
        key = Trim$(CStr(wsO.Cells(i, 2).Value))
        If key <> "" Then liveKeys(key) = True
    Next i

    Dim sumW As Object, sumM As Object, sumY As Object, sumFY As Object
    Set sumW = CreateObject("Scripting.Dictionary")
    Set sumM = CreateObject("Scripting.Dictionary")
    Set sumY = CreateObject("Scripting.Dictionary")
    Set sumFY = CreateObject("Scripting.Dictionary")
    sumW.CompareMode = vbTextCompare
    sumM.CompareMode = vbTextCompare
    sumY.CompareMode = vbTextCompare
    sumFY.CompareMode = vbTextCompare

    Dim k As Variant
    For Each k In liveKeys.keys
        sumW(k) = 0#: sumM(k) = 0#: sumY(k) = 0#: sumFY(k) = 0#
    Next k

    Dim lastRow As Long: lastRow = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then GoTo writeOut

    Dim arr As Variant: arr = wsP.Range("A2:C" & lastRow).Value
    Dim r As Long, sku As String, d As Variant, dd As Date, q As Double, bucket As String

    For r = LBound(arr, 1) To UBound(arr, 1)
        sku = Trim$(CStr(arr(r, 1)))
        d = arr(r, 2)
        q = NzD(arr(r, 3))

        If sku <> "" And IsDate(d) And q <> 0 Then
            If LaneFromSKU(sku) = "BVP" Then
                bucket = AnalysisKeyFromSKU(sku)

                If bucket <> "" And liveKeys.exists(bucket) Then
                    dd = CDate(d)
                    If dd >= wkStart And dd <= wkEnd Then sumW(bucket) = sumW(bucket) + q
                    If dd >= mStart And dd <= wkEnd Then sumM(bucket) = sumM(bucket) + q
                    If dd >= yStart And dd <= wkEnd Then sumY(bucket) = sumY(bucket) + q
                    If dd >= yStart And dd <= yEnd Then sumFY(bucket) = sumFY(bucket) + q
                End If
            End If
        End If
    Next r

writeOut:
    For i = 5 To 20
        key = Trim$(CStr(wsO.Cells(i, 2).Value))
        If key <> "" Then
            wsO.Cells(i, "C").Value = NzD(sumW(key))
            wsO.Cells(i, "I").Value = NzD(sumM(key))
            wsO.Cells(i, "O").Value = NzD(sumY(key))
            wsO.Cells(i, "U").Value = NzD(sumFY(key))
        End If
    Next i
End Sub




'===============================
' Local Belor + Lane helpers
'   (Standalone: no Sheet3 dependency)
'===============================
Public Function LaneFromSKU(ByVal sku As String) As String
    Dim s As String: s = UCase$(sku)
    If InStr(s, "WVP") > 0 Then
        LaneFromSKU = "WVP"
    ElseIf InStr(s, "AIR") > 0 Then
        LaneFromSKU = "WVP"
    ElseIf InStr(s, "CRM") > 0 Then
        LaneFromSKU = "WVP"
    ElseIf InStr(s, "BVP") > 0 Then
        LaneFromSKU = "BVP"
    Else
        LaneFromSKU = ""
    End If
End Function

Public Function BelorBucketFromSKU(ByVal sku As String) As String
    Dim s As String
    s = UCase$(sku)

    ' AIR / CRM overrides
    If InStr(s, "AIR") > 0 Then
        BelorBucketFromSKU = "AIR"
        Exit Function
    End If
    If InStr(s, "CRM") > 0 Then
        BelorBucketFromSKU = "CRM"
        Exit Function
    End If

    ' 30-inch variants
    If InStr(s, "12.4") > 0 And InStr(s, "30") > 0 Then
        BelorBucketFromSKU = "12.4 30"
        Exit Function
    End If

    ' 12.9 30 bucket includes BK 30
    If (InStr(s, "12.9") > 0 Or InStr(s, "BK") > 0) And InStr(s, "30") > 0 Then
        BelorBucketFromSKU = "12.9 30"
        Exit Function
    End If

    ' Standard diameters
    If InStr(s, "11.8") > 0 Then
        BelorBucketFromSKU = "11.8"
        Exit Function
    End If
    If InStr(s, "12.4") > 0 Then
        BelorBucketFromSKU = "12.4"
        Exit Function
    End If
    If InStr(s, "12.9") > 0 Or InStr(s, "BK") > 0 Then
        BelorBucketFromSKU = "12.9"
        Exit Function
    End If

    BelorBucketFromSKU = ""
End Function

Private Function BelorLaneKey(ByVal sku As String) As String
    Dim b As String: b = BelorBucketFromSKU(sku)
    If b = "" Then
        BelorLaneKey = ""
        Exit Function
    End If

    Dim ln As String: ln = LaneFromSKU(sku)
    If ln <> "" Then
        BelorLaneKey = b & " " & ln
    Else
        BelorLaneKey = b
    End If
End Function

Public Function NzDbl(val As Variant, Optional fallback As Double = 0#) As Double
    If IsError(val) Then
        NzDbl = fallback
    ElseIf IsMissing(val) Then
        NzDbl = fallback
    ElseIf IsNumeric(val) Then
        NzDbl = CDbl(val)
    ElseIf Len(Trim$(CStr(val))) = 0 Then
        NzDbl = fallback
    Else
        NzDbl = CDbl(val)
    End If
End Function

Public Function NormSKU(ByVal raw As String) As String
    raw = Trim$(raw)
    If Len(raw) = 0 Then NormSKU = "": Exit Function
    raw = UCase$(raw)
    raw = Replace(raw, " ", "")
    raw = Replace(raw, "-", "")
    NormSKU = raw
End Function


Public Sub BuildProdAnalysisWVP()
    If Not (SheetExists(SH_PROD_TABLE) And SheetExists(SH_ANALYSIS) And SheetExists(SH_SRC)) Then Exit Sub

    Dim wsP As Worksheet: Set wsP = GetSheet(SH_PROD_TABLE)
    Dim wsO As Worksheet: Set wsO = GetSheet(SH_ANALYSIS)

    ' FORECAST window: 4 weeks from Production Forecast headers
       Dim wkStart As Date, wkEnd As Date
If Not SafeGetDates(wsO, wkStart, wkEnd) Then Exit Sub

    Dim mStart As Date: mStart = DateSerial(Year(wkEnd), Month(wkEnd), 1)
    Dim yStart As Date: yStart = DateSerial(Year(wkEnd), 1, 1)
    Dim yEnd   As Date: yEnd = DateSerial(Year(wkEnd), 12, 31)

    Dim liveKeys As Object: Set liveKeys = CreateObject("Scripting.Dictionary")
    liveKeys.CompareMode = vbTextCompare

    Dim i As Long, key As String
    For i = 25 To 40
        key = Trim$(CStr(wsO.Cells(i, 2).Value))
        If key <> "" Then liveKeys(key) = True
    Next i

    Dim sumW As Object, sumM As Object, sumY As Object, sumFY As Object
    Set sumW = CreateObject("Scripting.Dictionary")
    Set sumM = CreateObject("Scripting.Dictionary")
    Set sumY = CreateObject("Scripting.Dictionary")
    Set sumFY = CreateObject("Scripting.Dictionary")
    sumW.CompareMode = vbTextCompare
    sumM.CompareMode = vbTextCompare
    sumY.CompareMode = vbTextCompare
    sumFY.CompareMode = vbTextCompare

    Dim k As Variant
    For Each k In liveKeys.keys
        sumW(k) = 0#: sumM(k) = 0#: sumY(k) = 0#: sumFY(k) = 0#
    Next k

    Dim lastRow As Long: lastRow = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then GoTo writeOut

    Dim arr As Variant: arr = wsP.Range("A2:C" & lastRow).Value
    Dim r As Long, sku As String, d As Variant, dd As Date, q As Double, bucket As String

    For r = LBound(arr, 1) To UBound(arr, 1)
        sku = Trim$(CStr(arr(r, 1)))
        d = arr(r, 2)
        q = NzD(arr(r, 3))

        If sku <> "" And IsDate(d) And q <> 0 Then
            If LaneFromSKU(sku) = "WVP" Then
                bucket = AnalysisKeyFromSKU(sku)

                If bucket <> "" And liveKeys.exists(bucket) Then
                    dd = CDate(d)
                    If dd >= wkStart And dd <= wkEnd Then sumW(bucket) = sumW(bucket) + q
                    If dd >= mStart And dd <= wkEnd Then sumM(bucket) = sumM(bucket) + q
                    If dd >= yStart And dd <= wkEnd Then sumY(bucket) = sumY(bucket) + q
                    If dd >= yStart And dd <= yEnd Then sumFY(bucket) = sumFY(bucket) + q
                End If
            End If
        End If
    Next r

writeOut:
    For i = 25 To 40
        key = Trim$(CStr(wsO.Cells(i, 2).Value))
        If key <> "" Then
            wsO.Cells(i, "C").Value = NzD(sumW(key))
            wsO.Cells(i, "I").Value = NzD(sumM(key))
            wsO.Cells(i, "O").Value = NzD(sumY(key))
            wsO.Cells(i, "U").Value = NzD(sumFY(key))
        End If
    Next i
End Sub






Private Sub WriteActualsBlock(wsOut As Worksheet, arrAct As Variant, dict As Object, _
                              rowStart As Long, rowEnd As Long, wkStart As Date, wkEnd As Date, _
                              tgtMonth As Long, tgtYear As Long)

    Dim mStart As Date: mStart = DateSerial(Year(wkEnd), Month(wkEnd), 1)
    Dim yStart As Date: yStart = DateSerial(Year(wkEnd), 1, 1)

    Dim rowIdx As Long, keyVal As String
    For rowIdx = rowStart To rowEnd
        keyVal = Trim$(CStr(wsOut.Cells(rowIdx, 2).text))
        If keyVal <> "" And dict.exists(keyVal) Then
            Dim skus As Variant: skus = dict(keyVal)
            Dim sumW As Double, sumM As Double, sumY As Double
            Dim r As Long, d As Variant, s As String, q As Double, dd As Date

            For r = LBound(arrAct, 1) To UBound(arrAct, 1)
                d = arrAct(r, 1): s = Trim$(CStr(arrAct(r, 2))): q = NzD(arrAct(r, 3))
                If IsInArray(s, skus) And IsDate(d) Then
                    dd = CDate(d)
                    If dd >= wkStart And dd <= wkEnd Then sumW = sumW + q
                    If dd >= mStart And dd <= wkEnd Then sumM = sumM + q
                    If dd >= yStart And dd <= wkEnd Then sumY = sumY + q
                End If
            Next r

            wsOut.Cells(rowIdx, "D").Value = sumW
            wsOut.Cells(rowIdx, "J").Value = sumM
            wsOut.Cells(rowIdx, "P").Value = sumY
        End If
    Next rowIdx
End Sub

Private Sub WriteRejectsBlock(wsOut As Worksheet, arrRej As Variant, dict As Object, _
                              rowStart As Long, rowEnd As Long, wkStart As Date, wkEnd As Date, _
                              tgtMonth As Long, tgtYear As Long)

    Dim mStart As Date: mStart = DateSerial(Year(wkEnd), Month(wkEnd), 1)
    Dim yStart As Date: yStart = DateSerial(Year(wkEnd), 1, 1)

    Dim rowIdx As Long, keyVal As String
    For rowIdx = rowStart To rowEnd
        keyVal = Trim$(CStr(wsOut.Cells(rowIdx, 2).text))
        If keyVal <> "" And dict.exists(keyVal) Then
            Dim skus As Variant: skus = dict(keyVal)
            Dim sumW As Double, sumM As Double, sumY As Double
            Dim i As Long, d As Variant, s As String, q As Double, dd As Date

            For i = LBound(arrRej, 1) To UBound(arrRej, 1)
                d = arrRej(i, 1): s = Trim$(CStr(arrRej(i, 2))): q = NzD(arrRej(i, 5))  ' S is 5th (O:S)
                If IsInArray(s, skus) And IsDate(d) Then
                    dd = CDate(d)
                    If dd >= wkStart And dd <= wkEnd Then sumW = sumW + q
                    If dd >= mStart And dd <= wkEnd Then sumM = sumM + q
                    If dd >= yStart And dd <= wkEnd Then sumY = sumY + q
                End If
            Next i

            wsOut.Cells(rowIdx, "G").Value = sumW
            wsOut.Cells(rowIdx, "M").Value = sumM
            wsOut.Cells(rowIdx, "S").Value = sumY
        End If
    Next rowIdx
End Sub




'====================
' Analysis helpers
'====================
Private Sub LoadKeySKUMap(ws As Worksheet, startRow As Long, keyCol As Long, skuCol As Long, endRow As Long, dict As Object)
    Dim i As Long, curKey As String, sku As String
    For i = startRow To endRow
        If Trim$(CStr(ws.Cells(i, keyCol).Value)) <> "" Then curKey = Trim$(CStr(ws.Cells(i, keyCol).Value))
        sku = Trim$(CStr(ws.Cells(i, skuCol).Value))
        If curKey <> "" And sku <> "" Then AddKeySku dict, curKey, sku
    Next i
End Sub

Private Sub AddKeySku(dict As Object, k As String, s As String)
    If Not dict.exists(k) Then
        dict.Add k, Array(s)
    Else
        Dim t As Variant: t = dict(k)
        ReDim Preserve t(UBound(t) + 1)
        t(UBound(t)) = s
        dict(k) = t
    End If
End Sub





'========================================================
' STEP 2e  8 Weeks View (Actuals first 4, Forecast next 4)
'========================================================
Public Sub Build8WeekView()
    ' 4 weeks ACTUALS (Release Log) + 4 weeks FORECAST (Production Table)
    If Not (SheetExists(SH_8W_VIEW) And SheetExists(SH_RELEASE_LOG) And SheetExists(SH_PROD_TABLE)) Then
        Debug.Print "!! Build8WeekView: required sheet missing."
        Exit Sub
    End If

    Dim wsView As Worksheet: Set wsView = GetSheet(SH_8W_VIEW)
    Dim wsR As Worksheet: Set wsR = GetSheet(SH_RELEASE_LOG)
    Dim wsProd As Worksheet: Set wsProd = GetSheet(SH_PROD_TABLE)

    ' Build aggregate dictionaries for fast lookups: key = CLng(WeekStart)|BelorLaneKey
    Dim dictAct As Object: Set dictAct = CreateObject("Scripting.Dictionary")
    Dim dictFc As Object: Set dictFc = CreateObject("Scripting.Dictionary")
    dictAct.CompareMode = vbTextCompare
    dictFc.CompareMode = vbTextCompare

    ' --- Actuals from Release Log (Qty Released)
    Dim lastR As Long: lastR = wsR.Cells(wsR.Rows.Count, RL_COL_DATE).End(xlUp).row
    If lastR >= 2 Then
        Dim arrR As Variant: arrR = wsR.Range(wsR.Cells(2, 1), wsR.Cells(lastR, RL_COL_QTY_REJECT)).Value
        Dim i As Long, dt As Variant, sku As String, qty As Double, wk As Date, k As String, bl As String
        For i = LBound(arrR, 1) To UBound(arrR, 1)
            dt = arrR(i, RL_COL_DATE)
            sku = Trim$(CStr(arrR(i, RL_COL_SKU)))
            qty = NzD(arrR(i, RL_COL_QTY_RELEASED))
            If sku <> "" And IsDate(dt) And qty <> 0 Then
                bl = BelorLaneKey(sku)
                If bl <> "" Then
                    wk = WeekStart(CDate(dt))
                    k = CStr(CLng(wk)) & "|" & bl
                    dictAct(k) = NzD(dictAct(k)) + qty
                End If
            End If
        Next i
    End If

    ' --- Forecast from Production Table
    Dim lastP As Long: lastP = wsProd.Cells(wsProd.Rows.Count, "A").End(xlUp).row
    If lastP >= 2 Then
        Dim arrP As Variant: arrP = wsProd.Range("A2:C" & lastP).Value
        Dim j As Long, d As Variant, q As Double, keySku As String
        For j = LBound(arrP, 1) To UBound(arrP, 1)
            keySku = Trim$(CStr(arrP(j, 1)))
            d = arrP(j, 2)
            q = NzD(arrP(j, 3))
            If keySku <> "" And IsDate(d) And q <> 0 Then
                bl = BelorLaneKey(keySku)
                If bl <> "" Then
                    wk = CDate(d) ' production table dates are already week starts
                    k = CStr(CLng(wk)) & "|" & bl
                    dictFc(k) = NzD(dictFc(k)) + q
                End If
            End If
        Next j
    End If

    Dim r As Long, c As Long, startDate As Date, headerKey As String, sumVal As Double, lookupKey As String

    Application.ScreenUpdating = False

    ' Rows 4-7 = actuals, rows 8-11 = forecast (per your existing layout)
    For r = 4 To 11
        If IsDate(wsView.Cells(r, "B").Value) Then
            startDate = CDate(wsView.Cells(r, "B").Value)
        Else
            wsView.Range(wsView.Cells(r, "C"), wsView.Cells(r, "R")).Value = 0
            wsView.Cells(r, "S").Formula = "=SUM(C" & r & ":R" & r & ")"
            GoTo NextRow
        End If

        For c = 3 To 18 ' C:R
            headerKey = Trim$(CStr(wsView.Cells(3, c).Value))
            If headerKey = "" Then
                wsView.Cells(r, c).Value = 0
            Else
                lookupKey = CStr(CLng(startDate)) & "|" & headerKey
                If r <= 7 Then
                    sumVal = NzD(dictAct(lookupKey))
                Else
                    sumVal = NzD(dictFc(lookupKey))
                End If
                wsView.Cells(r, c).Value = sumVal
            End If
        Next c

        wsView.Cells(r, "S").Formula = "=SUM(C" & r & ":R" & r & ")"

NextRow:
    Next r

    Application.ScreenUpdating = True
End Sub







'==================== styling helpers (no diagonal borders) ====================

Private Sub BoxBorders(rng As Range, Optional ByVal heavyBottom As Boolean = False, Optional ByVal gray As Long = 0)
    Dim edgeClr As Long, insideClr As Long
    If gray = 0 Then
        edgeClr = RGB(210, 210, 210)
        insideClr = RGB(225, 225, 225)
    Else
        edgeClr = gray
        insideClr = gray
    End If
    rng.Borders.LineStyle = xlNone            ' clears diagonals too

    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = edgeClr
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = edgeClr
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = edgeClr
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = IIf(heavyBottom, xlMedium, xlThin)
        .Color = edgeClr
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = insideClr
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = insideClr
    End With
End Sub

Public Sub StyleHeader(rng As Range)
    rng.Interior.Color = RGB(245, 247, 250)
    rng.Font.Bold = True
    rng.Font.Color = RGB(40, 40, 40)
    rng.RowHeight = 20
    rng.HorizontalAlignment = xlLeft
    rng.VerticalAlignment = xlCenter
    BoxBorders rng, True
End Sub

Public Sub StyleSubtotal(rng As Range)
    rng.Interior.Color = RGB(245, 245, 245)
    rng.Font.Bold = True
    BoxBorders rng
    On Error Resume Next
    rng.Columns(rng.Columns.Count).NumberFormat = "#,##0"
    On Error GoTo 0
End Sub

Public Sub StyleGrand(rng As Range)
    rng.Interior.Color = RGB(224, 224, 224)
    rng.Font.Bold = True
    BoxBorders rng, True
    On Error Resume Next
    rng.Columns(rng.Columns.Count).NumberFormat = "#,##0"
    On Error GoTo 0
End Sub

Private Function NzColor(colors As Variant, idx As Long) As Long
    On Error GoTo fallback
    Dim n As Long: n = UBound(colors) - LBound(colors) + 1
    If n <= 0 Then GoTo fallback
    NzColor = colors(LBound(colors) + ((idx - 1) Mod n))
    Exit Function
fallback:
    NzColor = RGB(230, 230, 230)
End Function

Public Function SafeDbl(ByVal v As Variant, Optional ByVal fallback As Double = 0#) As Double
    If IsError(v) Then
        SafeDbl = fallback
    ElseIf IsEmpty(v) Or v = "" Then
        SafeDbl = fallback
    ElseIf IsNumeric(v) Then
        SafeDbl = CDbl(v)
    Else
        SafeDbl = fallback
    End If
End Function

Public Function SafeText(ByVal v As Variant, Optional ByVal fallback As String = "") As String
    If IsError(v) Then
        SafeText = fallback
    ElseIf IsEmpty(v) Then
        SafeText = fallback
    Else
        SafeText = CStr(v)
    End If
End Function



'========================================================
' STEP 3  Email (with HTML Floor View embedded)
'========================================================
Public Sub SendWeeklyREVOUpdate()
    Const SH_SUMMARY As String = "Summary"

    '--- require Summary sheet
    If Not SheetExists(SH_SUMMARY) Then
        Debug.Print "!! SendWeeklyREVOUpdate: missing 'Summary'  email skipped."
        Exit Sub
    End If

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = GetSheet(SH_SUMMARY)

    ' ---------- read metrics safely ----------
    Dim wtdPct As Double, mtdPct As Double, ytdPct As Double
    Dim wtdAct As Double, wtdFc As Double
    Dim mtdAct As Double, mtdFc As Double
    Dim ytdAct As Double, ytdFc As Double
    Dim wtdReject As Double, mtdReject As Double, ytdReject As Double

    wtdPct = SafeDbl(ws.Range("F31").Value)
    mtdPct = SafeDbl(ws.Range("L31").Value)
    ytdPct = SafeDbl(ws.Range("R31").Value)

    wtdAct = SafeDbl(ws.Range("D31").Value)
    wtdFc = SafeDbl(ws.Range("C31").Value)

    mtdAct = SafeDbl(ws.Range("J31").Value)
    mtdFc = SafeDbl(ws.Range("I31").Value)

    ytdAct = SafeDbl(ws.Range("P31").Value)
    ytdFc = SafeDbl(ws.Range("O31").Value)

    wtdReject = SafeDbl(ws.Range("H31").Value)
    mtdReject = SafeDbl(ws.Range("N31").Value)
    ytdReject = SafeDbl(ws.Range("T31").Value)

    ' ---------- compose summary text ----------
    Dim wtdText As String, mtdText As String, ytdText As String
    wtdText = "<b style='font-size:11pt;color:" & IIf(wtdPct < 0, "red", "black") & "'>" & _
              Format(wtdPct, "0.00%") & "</b> / " & _
              Format(wtdAct, "#,##0") & " actual v. " & _
              Format(wtdFc, "#,##0") & " forecasted"

    mtdText = "<b style='font-size:11pt;color:" & IIf(mtdPct < 0, "red", "black") & "'>" & _
              Format(mtdPct, "0.00%") & "</b> / " & _
              Format(mtdAct, "#,##0") & " actual v. " & _
              Format(mtdFc, "#,##0") & " forecasted"

    ytdText = "<b style='font-size:11pt;color:" & IIf(ytdPct < 0, "red", "black") & "'>" & _
              Format(ytdPct, "0.00%") & "</b> / " & _
              Format(ytdAct, "#,##0") & " actual v. " & _
              Format(ytdFc, "#,##0") & " forecasted"

    Dim colorWTD As String, colorMTD As String, colorYTD As String
    colorWTD = IIf(wtdReject > 0.02, "red", "green")
    colorMTD = IIf(mtdReject > 0.02, "red", "green")
    colorYTD = IIf(ytdReject > 0.02, "red", "green")

    ' ---------- body (guard BuildFloorViewHtml return to avoid type mismatch) ----------
    Dim body As String
    body = "Good Afternoon,<br><br>" & _
           "<b>LastWTD:</b> " & wtdText & "<br>" & _
           "<b>ThisMTD:</b> " & mtdText & "<br>" & _
           "<b>ThisYTD:</b> " & ytdText & "<br><br>" & _
           "<span style='font-size:14pt;font-weight:bold;text-decoration:underline;'>Forecast v. Actual Summary</span><br><br>" & _
           "<span style='font-size:14pt;font-weight:bold;text-decoration:underline;'>Reject Information</span><br>" & _
           "WTD: <span style='color:" & colorWTD & ";font-size:11pt'>" & Format(wtdReject, "0.00%") & "</span><br>" & _
           "MTD: <span style='color:" & colorMTD & ";font-size:11pt'>" & Format(mtdReject, "0.00%") & "</span><br>" & _
           "YTD: <span style='color:" & colorYTD & ";font-size:11pt'>" & Format(ytdReject, "0.00%") & "</span><br><br>" & _
           "<span style='font-size:14pt;font-weight:bold;text-decoration:underline;'>4 Week Actual & 4 Week Forecast</span><br><br>" & _
           "<span style='font-size:14pt;font-weight:bold;text-decoration:underline;'>REVO Weekly Floor Update</span><br><br>" & _
           NzStr(BuildFloorViewHtml(4)) & "<br>"

    ' ---------- attachment: save a copy next to workbook (fallbacks if needed) ----------
    Dim weekLabel As String
    weekLabel = NzStr(ws.Range("C1").Value)
    If Len(weekLabel) = 0 Then weekLabel = Format(Date, "yyyy-mm-dd")
    weekLabel = Replace$(weekLabel, "/", "-")
    weekLabel = Replace$(weekLabel, ":", "-")
    weekLabel = Replace$(weekLabel, "\", "-")

    Dim ext As String, p As Long
    p = InStrRev(wb.name, ".")
    If p > 0 Then ext = Mid$(wb.name, p) Else ext = ".xlsm"

    Dim basePath As String
    basePath = wb.path
    If Len(basePath) = 0 Or InStr(1, basePath, "http", vbTextCompare) > 0 Then
        basePath = Application.DefaultFilePath
        If Len(Dir$(basePath, vbDirectory)) = 0 Then
            basePath = Environ$("USERPROFILE") & "\Documents"
        End If
    End If

    Dim attachFile As String
    attachFile = basePath & "\REVO Update for week " & weekLabel & ext
    On Error Resume Next
    wb.SaveCopyAs attachFile
    On Error GoTo 0

    ' ---------- Outlook send (fallback: save HTML file) ----------
    Dim OutApp As Object, OutMail As Object
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then
        Err.Clear: Set OutApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    If OutApp Is Nothing Then
        Dim htmlPath As String
        htmlPath = basePath & "\REVO_Update_" & Format(Date, "yyyymmdd") & ".html"
        PutTextFile htmlPath, body
        MsgBox "Outlook not available. HTML saved to:" & vbCrLf & htmlPath & vbCrLf & _
               "Attach this file as needed:" & vbCrLf & attachFile, vbExclamation, "Weekly REVO Update"
        Exit Sub
    End If

    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .To = " cfournier@predatorgroup.com"
        
        .Subject = "REVO Update  " & Format(Date, "mm.dd.yy")
        .Importance = 2
        .HTMLBody = body
        If Len(Dir$(attachFile)) > 0 Then .Attachments.Add attachFile
        .Display
    End With
End Sub

Private Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or VarType(v) = vbEmpty Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

'========================================================
' Floor View ? HTML (for email body)
'========================================================
Public Function BuildFloorViewHtml(Optional ByVal maxWeeks As Long = 0) As String
    Const COL_LABEL As Long = 6 ' labels in column F
    Const COL_TOTAL As Long = 7 ' totals in column G

    If Not SheetExists(SH_FLOOR) Then
        BuildFloorViewHtml = "<i>(Floor View not found)</i>"
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = GetSheet(SH_FLOOR)

    ' --- Robust last row:
    ' A) last used anywhere
    Dim lastAny As Range, lastAnyRow As Long
    Set lastAny = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastAny Is Nothing Then
        BuildFloorViewHtml = "<i>(No Floor View data)</i>"
        Exit Function
    End If
    lastAnyRow = lastAny.row

    ' B) last "Week of" header in column A (covers weeks where G is blank)
    Dim lastWeekHdr As Range, lastWeekHdrRow As Long
    Set lastWeekHdr = ws.Columns(1).Find(What:="Week of*", LookIn:=xlValues, LookAt:=xlPart, _
                                         SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not lastWeekHdr Is Nothing Then lastWeekHdrRow = lastWeekHdr.row

    ' C) last row in G (not strictly needed anymore, but keep as a sanity candidate)
    Dim lastG As Range, lastGRow As Long
    Set lastG = ws.Columns(COL_TOTAL).Find(What:="*", LookIn:=xlFormulas, _
                                           SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not lastG Is Nothing Then lastGRow = lastG.row

    ' Final last row takes the max of all candidates so we never truncate future weeks
    Dim lastRow As Long
    lastRow = Application.WorksheetFunction.Max(lastAnyRow, lastWeekHdrRow, lastGRow)

    Dim StyleTable$, styleTh$, styleTd$, styleWeek$, styleSub$, styleGrandRow$
    StyleTable = "border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;font-size:10pt;"
    styleTh = "background:#E6E6E6;padding:6px 8px;border:1px solid #DDDDDD;text-align:left;"
    styleTd = "padding:6px 8px;border:1px solid #EEEEEE;"
    styleWeek = "background:#202020;color:#FFFFFF;font-weight:bold;padding:8px 10px;margin-top:6px;margin-bottom:4px;"
    styleSub = "background:#F5F5F5;font-weight:bold;"
    styleGrandRow = "background:#E0E0E0;font-weight:bold;"

    Dim html As String: html = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:10pt'>"
    Dim r As Long: r = 2
    Dim openTable As Boolean, weekCount As Long, c As Long

    Do While r <= lastRow
        Dim valA As String
        valA = CStr(ws.Cells(r, 1).Value2)

        ' ---- Start of a week section ("Week of ...")
        If Len(valA) > 0 And LCase$(LEFT$(Trim$(valA), 7)) = "week of" Then
            If openTable Then
                html = html & "</tbody></table><br>"
                openTable = False
            End If

            weekCount = weekCount + 1
'            If maxWeeks > 0 And weekCount > maxWeeks Then Exit Do

            html = html & "<div style='" & styleWeek & "'>" & HtmlEncode(valA) & "</div>"

            ' Column headers on next row
            If r + 1 <= lastRow Then
                html = html & "<table style='" & StyleTable & "'><thead><tr>"
                For c = 1 To 7
                    html = html & "<th style='" & styleTh & "'>" & _
                                   HtmlEncode(CStr(ws.Cells(r + 1, c).Value)) & "</th>"
                Next c
                html = html & "</tr></thead><tbody>"
                openTable = True
                r = r + 2
                GoTo NextRow
            End If
        End If

        ' ---- Subtotal / Grand Total row (label in column F)
        Dim lbl As String
        lbl = UCase$(Trim$(CStr(ws.Cells(r, COL_LABEL).Value2)))
        If Right$(lbl, 1) = ":" Then lbl = LEFT$(lbl, Len(lbl) - 1)

        If openTable And (lbl = "SUBTOTAL" Or lbl = "GRAND TOTAL") Then
            Dim styleRow As String
            styleRow = IIf(lbl = "GRAND TOTAL", styleGrandRow, styleSub)

            html = html & "<tr>" & _
                          "<td style='" & styleTd & "' colspan='5'>&nbsp;</td>" & _
                          "<td style='" & styleTd & styleRow & "'>" & HtmlEncode( _
                              IIf(lbl = "GRAND TOTAL", "Grand Total:", "Subtotal:")) & "</td>" & _
                          "<td style='" & styleTd & styleRow & "'>" & _
                              HtmlEncode(FormatNumber(NzD(ws.Cells(r, COL_TOTAL).Value), 0)) & _
                          "</td></tr>"

            If lbl = "GRAND TOTAL" Then
                html = html & "</tbody></table>"
                openTable = False
            End If

            r = r + 1
            GoTo NextRow
        End If

        ' ---- Regular data row
        If openTable Then
            html = html & "<tr>"
            For c = 1 To 7
                Dim cellVal As Variant, outTxt As String
                cellVal = ws.Cells(r, c).Value
                Select Case c
                    Case 5
                        outTxt = IIf(IsDate(cellVal), Format$(CDate(cellVal), "m/d/yyyy"), CStr(cellVal))
                    Case 3
                        outTxt = FormatNumber(NzD(cellVal), 0)
                    Case Else
                        outTxt = CStr(cellVal)
                End Select
                html = html & "<td style='" & styleTd & "'>" & HtmlEncode(outTxt) & "</td>"
            Next c
            html = html & "</tr>"
        End If

        r = r + 1
NextRow:
    Loop

    If openTable Then html = html & "</tbody></table>"
    html = html & "</div>"
    BuildFloorViewHtml = html
End Function



'===========================
' Utilities & styling
'===========================
Public Function SheetExists(ByVal sheetName As String, _
                            Optional ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function GetSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Private Function SafeGetDates(ws As Worksheet, ByRef d1 As Date, ByRef d2 As Date) As Boolean
    Dim v1 As Variant, v2 As Variant
    Dim s1 As Date, s2 As Date

    ' 1) Try the canonical layout: D1 and D2 are dates (or serials)
    v1 = ws.Range("D1").Value2
    v2 = ws.Range("D2").Value2
    If (Not IsError(v1)) And (Not IsError(v2)) _
       And (IsDate(v1) Or IsNumeric(v1)) _
       And (IsDate(v2) Or IsNumeric(v2)) Then
        d1 = CDate(v1)
        d2 = CDate(v2)
        SafeGetDates = True
        Exit Function
    End If

    ' 2) Try a single cell with "mm/dd - mm/dd" (allow either C1 or D1)
    d1 = 0: d2 = 0
    GetDateRange ws, "C1", s1, s2
    If s1 <> 0 And s2 <> 0 Then
        d1 = s1: d2 = s2
        SafeGetDates = True
        Exit Function
    End If

    GetDateRange ws, "D1", s1, s2
    If s1 <> 0 And s2 <> 0 Then
        d1 = s1: d2 = s2
        SafeGetDates = True
        Exit Function
    End If

    ' 3) Last resort: if D1 is a valid date, infer a 7-day window
    If (Not IsError(v1)) And (IsDate(v1) Or IsNumeric(v1)) Then
        d1 = CDate(v1)
        d2 = d1 + 6
        SafeGetDates = True
        Exit Function
    End If

    Debug.Print "!! Invalid date window on '" & ws.name & _
                "'. D1=[" & SafeText(v1) & "], D2=[" & SafeText(v2) & _
                "], C1=[" & SafeText(ws.Range("C1").Value) & "]"
    SafeGetDates = False
End Function



Private Function IsArray2D(v As Variant) As Boolean
    On Error GoTo nope
    If IsArray(v) Then
        Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
        lb1 = LBound(v, 1): ub1 = UBound(v, 1)
        lb2 = LBound(v, 2): ub2 = UBound(v, 2)
        IsArray2D = (ub1 >= lb1 And ub2 >= lb2)
    End If
    Exit Function
nope:
    IsArray2D = False
End Function

Private Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(val, CStr(arr(i)), vbTextCompare) = 0 Then IsInArray = True: Exit Function
    Next i
    IsInArray = False
End Function

Private Function WeekStart(ByVal d As Date) As Date
    WeekStart = d - (Weekday(d, vbMonday) - 1) ' Monday start
End Function

Private Function HtmlEncode(ByVal s As String) As String
    If LenB(s) = 0 Then HtmlEncode = "": Exit Function
    s = Replace$(s, "&", "&amp;")
    s = Replace$(s, "<", "&lt;")
    s = Replace$(s, ">", "&gt;")
    s = Replace$(s, """", "&quot;")
    HtmlEncode = s
End Function

Private Function NzD(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then NzD = 0# Else NzD = CDbl(v)
End Function





Private Sub PutTextFile(ByVal path As String, ByVal text As String)
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, text
    Close #f
End Sub


Public Sub CleanProductionTableDuplicates()
    Const COL_SKU As Long = 1
    Const COL_DATE As Long = 2
    Const COL_QTY As Long = 3

    If Not SheetExists(SH_PROD_TABLE) Then
        MsgBox "Production Table sheet not found.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = GetSheet(SH_PROD_TABLE)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim arr As Variant: arr = ws.Range("A2:E" & lastRow).Value  ' A: SKU, B: Date, C: Qty, D: Month, E: Year
    If Not IsArray2D(arr) Then Exit Sub

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long, k As String, sku As String, d As Variant, q As Double
    Dim dd As Date

    ' Keep the last occurrence seen for each (SKU, Date)
    For r = LBound(arr, 1) To UBound(arr, 1)
        sku = Trim$(CStr(arr(r, COL_SKU)))
        d = arr(r, COL_DATE)
        q = NzD(arr(r, COL_QTY))
        If Len(sku) > 0 And IsDate(d) Then
            dd = CDate(d)
            k = sku & "|" & CLng(dd) ' CLng(date) is robust key
            dict(k) = Array(sku, dd, q) ' overwrite ? keeps the last seen value
        End If
    Next r

    ' Write back unique rows
    Dim out() As Variant, i As Long
    ReDim out(1 To dict.Count, 1 To 5)

    i = 1
    Dim v As Variant
    For Each v In dict.Items
        out(i, 1) = v(0)                ' SKU
        out(i, 2) = v(1)                ' Date
        out(i, 3) = v(2)                ' Qty
        out(i, 4) = Month(v(1))         ' Month
        out(i, 5) = Year(v(1))          ' Year
        i = i + 1
    Next v

    ' Clear and rewrite
    ws.Range("A2:E" & lastRow).ClearContents
    ws.Range("A2").Resize(dict.Count, 5).Value = out

    ' Sort by date ascending
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row), _
                        SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1:E" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
        .header = xlYes
        .Apply
    End With

    MsgBox "Production Table de-duplicated: " & dict.Count & " unique rows kept.", vbInformation
End Sub


Private Sub DedupHorizon_ProdTable(ByVal cutoff As Date)
    Dim ws As Worksheet: Set ws = GetSheet(SH_PROD_TABLE)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim arr As Variant: arr = ws.Range("A2:E" & lastRow).Value
    If Not IsArray2D(arr) Then Exit Sub

    Dim head() As Variant: head = ws.Range("A1:E1").Value

    Dim dictPast As Object:  Set dictPast = CreateObject("Scripting.Dictionary")
    Dim dictHzn  As Object:  Set dictHzn = CreateObject("Scripting.Dictionary")
    dictPast.CompareMode = vbTextCompare
    dictHzn.CompareMode = vbTextCompare

    Dim r As Long, sku As String, d As Variant, dd As Date, q As Double, k As String

    ' Separate past vs horizon; keep last in the horizon
    For r = LBound(arr, 1) To UBound(arr, 1)
        sku = Trim$(CStr(arr(r, 1)))
        d = arr(r, 2): q = NzD(arr(r, 3))
        If Len(sku) > 0 And IsDate(d) Then
            dd = CDate(d): k = sku & "|" & CLng(dd)
            If dd < cutoff Then
                dictPast(k) = Array(sku, dd, q)
            Else
                dictHzn(k) = Array(sku, dd, q)  ' overwrite ? last occurrence
            End If
        End If
    Next r

    ' Stitch back together
    Dim n As Long: n = dictPast.Count + dictHzn.Count
    Dim out() As Variant: ReDim out(1 To n, 1 To 5)
    Dim i As Long: i = 1
    Dim v As Variant
    For Each v In dictPast.Items
        out(i, 1) = v(0): out(i, 2) = v(1): out(i, 3) = v(2)
        out(i, 4) = Month(v(1)): out(i, 5) = Year(v(1)): i = i + 1
    Next v
    For Each v In dictHzn.Items
        out(i, 1) = v(0): out(i, 2) = v(1): out(i, 3) = v(2)
        out(i, 4) = Month(v(1)): out(i, 5) = Year(v(1)): i = i + 1
    Next v

    ws.Range("A2:E" & lastRow).ClearContents
    ws.Range("A2").Resize(n, 5).Value = out

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row), _
                        SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1:E" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
        .header = xlYes
        .Apply
    End With
End Sub


Sub RefreshAllQueries_Status()
    Dim cn As WorkbookConnection
    Dim i As Long, total As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = True

    total = ThisWorkbook.Connections.Count
    i = 0

    For Each cn In ThisWorkbook.Connections
        If InStr(1, cn.name, "Query -", vbTextCompare) > 0 Then
            i = i + 1
            Application.StatusBar = "Refreshing query " & i & " of " & total & ": " & cn.name
            cn.Refresh
            DoEvents ' allow Excel UI to update
        End If
    Next cn

    Application.StatusBar = "All queries refreshed successfully."
    Application.Wait (Now + TimeValue("0:00:02"))
    Application.StatusBar = False

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub





'========================
' ENTRY POINT
'========================
Public Sub UpdateInputFromProdForecast()
    Dim scrn&, evt&, calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual
'    On Error GoTo CleanFail

    Dim wsS As Worksheet, wsD As Worksheet
    Set wsS = ThisWorkbook.Worksheets(SH_SRC)
    Set wsD = ThisWorkbook.Worksheets(SH_INPUT)

    Dim lastSrc As Long, lastDstRow As Long, lastDstCol As Long
    lastSrc = wsS.Cells(wsS.Rows.Count, SRC_COL_SKU).End(xlUp).row
    lastDstRow = wsD.Cells(wsD.Rows.Count, DST_COL_SKU).End(xlUp).row
    lastDstCol = wsD.Cells(DST_ROW_DATES, wsD.Columns.Count).End(xlToLeft).Column

    '--- destination: map Date -> Column (reads row 2)
    Dim colByDate As Object: Set colByDate = CreateObject("Scripting.Dictionary")
    Dim c As Long, d As Variant
    For c = 1 To lastDstCol
        d = wsD.Cells(DST_ROW_DATES, c).Value
        If IsDate(d) Then
            If Not colByDate.exists(CLng(CDate(d))) Then colByDate.Add CLng(CDate(d)), c
        End If
    Next c

    '--- source: build sums per (SKU|Date) and set of touched dates
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim touched As Object: Set touched = CreateObject("Scripting.Dictionary")

    Dim r As Long, sku As String, dt As Variant, qty As Double, key As Variant
    For r = 2 To lastSrc
        sku = SafeText(wsS.Cells(r, SRC_COL_SKU).Value)
        dt = wsS.Cells(r, SRC_COL_DATE).Value
        qty = SafeDbl(wsS.Cells(r, SRC_COL_QTY).Value)

        If Len(sku) > 0 And IsDate(dt) And qty <> 0 Then
            If colByDate.exists(CLng(CDate(dt))) Then
                key = sku & "|" & CLng(CDate(dt))
                If sums.exists(key) Then
                    sums(key) = sums(key) + qty
                Else
                    sums.Add key, qty
                End If
                If Not touched.exists(CLng(CDate(dt))) Then touched.Add CLng(CDate(dt)), True
            End If
        End If
    Next r

    '--- clear touched date-columns on the Input matrix
    Dim t As Variant, col As Long
    For Each t In touched.keys
        col = colByDate(CLng(t))
        wsD.Cells(DST_ROW_TOTAL, col).ClearContents
        wsD.Range(wsD.Cells(DST_FIRST_DATA_ROW, col), wsD.Cells(lastDstRow, col)).ClearContents
    Next t

    '--- build SKU -> row map on Input (col A)
    Dim rowBySku As Object: Set rowBySku = CreateObject("Scripting.Dictionary")
    Dim rr As Long, a As String
    For rr = DST_FIRST_DATA_ROW To lastDstRow
        a = SafeText(wsD.Cells(rr, DST_COL_SKU).Value)
        If Len(a) > 0 Then If Not rowBySku.exists(a) Then rowBySku.Add a, rr
    Next rr

    '--- write values
    Dim parts() As String, sk As String, dKey As Long, outRow As Long
    For Each key In sums.keys
        parts = Split(key, "|")
        sk = parts(0): dKey = CLng(parts(1))
        If rowBySku.exists(sk) Then
            outRow = rowBySku(sk)
            col = colByDate(dKey)
            wsD.Cells(outRow, col).Value = sums(key)
        End If
    Next key

    '--- recompute Weekly Totals on touched columns
    Dim rngCol As Range
    For Each t In touched.keys
        col = colByDate(CLng(t))
        Set rngCol = wsD.Range(wsD.Cells(DST_FIRST_DATA_ROW, col), wsD.Cells(lastDstRow, col))
        wsD.Cells(DST_ROW_TOTAL, col).Value = Application.WorksheetFunction.Sum(rngCol)
    Next t

CleanExit:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Exit Sub

CleanFail:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Err.Raise Err.Number, "UpdateInputFromProdForecast", Err.Description
End Sub

Public Sub BuildComponentHealth()
    Dim scrn&, evt&, calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Dim wsB As Worksheet, wsM As Worksheet, wsInv As Worksheet, wsPO As Worksheet, wsOut As Worksheet
    Set wsB = ThisWorkbook.Worksheets(SH_BOM)
    Set wsM = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsInv = ThisWorkbook.Worksheets(SH_INV)
    Set wsPO = ThisWorkbook.Worksheets(SH_PO)

    ' (Re)create output sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(SH_OUT).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = ThisWorkbook.Worksheets.Add(After:=wsPO)
    wsOut.name = SH_OUT
    wsOut.Range("A1:H1").Value = Array( _
        "Component ID", _
        "Inventory", _
        "Lead Time (Months)", _
        "Safety Stock", _
        "Safety Stock Covered?", _
        "Safety Stock Whole Date", _
        "Weeks of Production Available", _
        "Inv Dead Date")
    wsOut.Rows(1).Font.Bold = True

    ' --- Load BOM: parents/children + component meta (Critical, LT days)
    Dim parents As Object: Set parents = CreateObject("Scripting.Dictionary")   ' parent SKU -> True
    Dim children As Object: Set children = CreateObject("Scripting.Dictionary") ' parent SKU -> Collection(childID, qtyPer)

    Dim compCrit As Object: Set compCrit = CreateObject("Scripting.Dictionary") ' compID -> True if CRITICAL
    Dim compLTDays As Object: Set compLTDays = CreateObject("Scripting.Dictionary") ' compID -> LT in days (first non-zero)

    Dim lastB As Long, r As Long
    Dim parent As String, comp As String
    Dim qPer As Double
    Dim crit As String, ltDays As Double

    lastB = wsB.Cells(wsB.Rows.Count, BOM_COL_SKU).End(xlUp).row

    For r = 2 To lastB
        parent = KeyNorm(wsB.Cells(r, BOM_COL_SKU).Value)
        comp = KeyNorm(wsB.Cells(r, BOM_COL_COMP).Value)
        qPer = SafeDbl(wsB.Cells(r, BOM_COL_QTY).Value)
        If qPer = 0 Then qPer = 1

        ' New meta columns on BOM:
        '   Col D = Critical_Status
        '   Col E = Lead_Time_By_Item.LT (days)
        crit = UCase$(Trim$(CStr(wsB.Cells(r, 4).Value)))
        ltDays = SafeDbl(wsB.Cells(r, 5).Value)

        If Len(parent) > 0 And Len(comp) > 0 Then
            parents(parent) = True
            If Not children.exists(parent) Then children.Add parent, New Collection
            children(parent).Add Array(comp, qPer)

            ' Capture per-component Critical + LT (use first non-zero LT encountered)
            If Not compCrit.exists(comp) Then
                compCrit.Add comp, (crit = "CRITICAL")
            ElseIf crit = "CRITICAL" Then
                ' If any row marks it critical, treat as critical
                compCrit(comp) = True
            End If

            If ltDays > 0 Then
                If Not compLTDays.exists(comp) Then
                    compLTDays.Add comp, ltDays
                End If
            End If
        End If
    Next r

    ' --- Read planning horizon (dates row)
    Dim lastCol As Long, c As Long
    lastCol = wsM.Cells(MX_ROW_DATES, wsM.Columns.Count).End(xlToLeft).Column

    Dim weekCols As New Collection
    Dim weekKeys As New Collection ' date keys (CLng)

    For c = 1 To lastCol
        If IsDate(wsM.Cells(MX_ROW_DATES, c).Value) Then
            weekCols.Add c
            weekKeys.Add CLng(CDate(wsM.Cells(MX_ROW_DATES, c).Value))
        End If
    Next c

    If weekCols.Count = 0 Then GoTo CleanExit

    ' Horizon span for monthly average calc
    Dim horizonStart As Long, horizonEnd As Long, daysSpan As Long
    horizonStart = CLng(weekKeys(1))
    horizonEnd = CLng(weekKeys(weekCols.Count)) + 7 ' assume 7-day buckets
    daysSpan = horizonEnd - horizonStart
    If daysSpan < 1 Then daysSpan = 1

    ' --- Flattened BOM cache (memoize)
    Dim flatCache As Object: Set flatCache = CreateObject("Scripting.Dictionary")

    ' --- Build weekly leaf-component needs + total need per component
    Dim compNeeds As Object: Set compNeeds = CreateObject("Scripting.Dictionary") ' key = comp|dateKey
    Dim totalNeed As Object: Set totalNeed = CreateObject("Scripting.Dictionary") ' comp -> total qty needed over horizon

    Dim lastRow As Long, rr As Long
    Dim sku As String, wkCol As Long, wkKey As Long, skuQty As Double
    Dim perUnit As Object
    Dim i As Long, k As Variant, addKey As String, compKey As String

    lastRow = wsM.Cells(wsM.Rows.Count, MX_COL_SKU).End(xlUp).row

    For rr = MX_ROW_FIRST To lastRow
        sku = KeyNorm(wsM.Cells(rr, MX_COL_SKU).Value)
        If Len(sku) = 0 Then GoTo NextRow

        Set perUnit = FlattenBOM(sku, children, parents, flatCache) ' Dict(leaf -> qty per unit)
        If Not perUnit Is Nothing And perUnit.Count > 0 Then
            For i = 1 To weekCols.Count
                wkCol = weekCols(i): wkKey = CLng(weekKeys(i))
                skuQty = SafeDbl(wsM.Cells(rr, wkCol).Value)
                If skuQty <> 0 Then
                    For Each k In perUnit.keys
                        compKey = CStr(k)
                        addKey = compKey & "|" & wkKey

                        ' Weekly requirements
                        If compNeeds.exists(addKey) Then
                            compNeeds(addKey) = compNeeds(addKey) + skuQty * SafeDbl(perUnit(k))
                        Else
                            compNeeds.Add addKey, skuQty * SafeDbl(perUnit(k))
                        End If

                        ' Total requirements for monthly average
                        If totalNeed.exists(compKey) Then
                            totalNeed(compKey) = totalNeed(compKey) + skuQty * SafeDbl(perUnit(k))
                        Else
                            totalNeed.Add compKey, skuQty * SafeDbl(perUnit(k))
                        End If
                    Next k
                End If
            Next i
        End If
NextRow:
    Next rr

    ' --- Inventory: component -> O/H
    Dim inv As Object: Set inv = CreateObject("Scripting.Dictionary")
    Dim lastInv As Long
    lastInv = wsInv.Cells(wsInv.Rows.Count, INV_COL_SKU).End(xlUp).row

    For r = 2 To lastInv
        comp = KeyNorm(wsInv.Cells(r, INV_COL_SKU).Value)
        If Len(comp) > 0 Then inv(comp) = SafeDbl(wsInv.Cells(r, INV_COL_QTY).Value)
    Next r

    ' --- PO: component -> sorted Collection of (dateKey, qty)
    Dim po As Object: Set po = CreateObject("Scripting.Dictionary")
    Dim lastPO As Long
    Dim dt As Variant, qPO As Double

    lastPO = wsPO.Cells(wsPO.Rows.Count, PO_COL_SKU).End(xlUp).row

    For r = 2 To lastPO
        comp = KeyNorm(wsPO.Cells(r, PO_COL_SKU).Value)
        dt = wsPO.Cells(r, PO_COL_DATE).Value
        qPO = SafeDbl(wsPO.Cells(r, PO_COL_QTY).Value)

        If Len(comp) > 0 And IsDate(dt) And qPO <> 0 Then
            If Not po.exists(comp) Then po.Add comp, New Collection
            po(comp).Add Array(CLng(CDate(dt)), qPO)
        End If
    Next r

    ' sort each comp's receipts
    Dim kComp As Variant
    For Each kComp In po.keys
        SortPOCollectionByDate po(kComp)
    Next kComp

' --- Build full set of components to report (everything from BOM, Inventory, and Needs)
Dim compSet As Object: Set compSet = CreateObject("Scripting.Dictionary")
Dim key As Variant, compOnly As Variant

' 1. From BOM Component_ID column
For r = 2 To lastB
    comp = KeyNorm(wsB.Cells(r, BOM_COL_COMP).Value)
    If Len(comp) > 0 Then compSet(comp) = True
Next r

' 2. From inventory (ensure any extras also included)
For Each compOnly In inv.keys
    compSet(compOnly) = True
Next compOnly

' 3. From production needs (future requirements)
For Each key In compNeeds.keys
    compOnly = Split(CStr(key), "|")(0)
    compSet(compOnly) = True
Next key


    ' --- Simulate per component
    Dim outRow As Long: outRow = 2
    Dim startInv As Double, bal As Double, w As Long, need As Double
    Dim weeksOK As Long
    Dim deadDate As Variant, dateKey As Long
    Dim colPO As Collection, hasPO As Boolean

    Dim avgDailyNeed As Double, avgMonthlyNeed As Double
    Dim leadTimeMonths As Long

    Dim safetyStock As Double
    Dim ssCovered As String
    Dim ssWholeDate As Variant
    Dim isCritical As Boolean
    Dim ltCompDays As Double

    For Each kComp In compSet.keys
        comp = CStr(kComp)

        ' --- starting inventory
        startInv = 0
        If inv.exists(comp) Then startInv = SafeDbl(inv(comp))

        ' --- lead time (months) from BOM meta (days -> months)
        ltCompDays = 0
        If compLTDays.exists(comp) Then ltCompDays = SafeDbl(compLTDays(comp))

        If ltCompDays > 0 Then
            leadTimeMonths = Round(ltCompDays / 30, 0)
            If leadTimeMonths < 0 Then leadTimeMonths = 0
        Else
            leadTimeMonths = 0
        End If

        ' --- avg monthly need
        If totalNeed.exists(comp) And daysSpan > 0 Then
            avgDailyNeed = SafeDbl(totalNeed(comp)) / daysSpan
            avgMonthlyNeed = avgDailyNeed * 30#
        Else
            avgDailyNeed = 0#
            avgMonthlyNeed = 0#
        End If

        ' --- safety stock based on Critical flag & lead time (months)
        isCritical = False
        If compCrit.exists(comp) Then isCritical = CBool(compCrit(comp))

        If leadTimeMonths > 0 And avgMonthlyNeed > 0 Then
            If isCritical Then
                safetyStock = avgMonthlyNeed * leadTimeMonths * 2#
            Else
                safetyStock = avgMonthlyNeed * leadTimeMonths * 1.5
            End If
        Else
            safetyStock = 0#
        End If

        If safetyStock <> 0 Then
            safetyStock = WorksheetFunction.Round(safetyStock, 0)
        End If

        ' --- Safety Stock Covered? (inventory only)
        If safetyStock = 0 Then
            ssCovered = "Yes"
        ElseIf startInv >= safetyStock Then
            ssCovered = "Yes"
        Else
            ssCovered = "No"
        End If

        ' --- simulate week by week (same semantics as original for WeeksOK & DeadDate)
        bal = startInv
        weeksOK = 0
        deadDate = vbNullString
        ssWholeDate = vbNullString

        hasPO = po.exists(comp)
        If hasPO Then
            Set colPO = ClonePOCollection(po(comp))
        Else
            Set colPO = Nothing
        End If

        For w = 1 To weekCols.Count
            dateKey = CLng(weekKeys(w))
            key = comp & "|" & dateKey

            need = 0
            If compNeeds.exists(key) Then
                need = SafeDbl(compNeeds(key))
            End If

            ' 1) consume this week's need
            bal = bal - need
            If bal <= 0 Then
                deadDate = CDate(dateKey)
                Exit For ' preserve original behavior
            End If
            weeksOK = weeksOK + 1

            ' 2) add receipts with ETA <= this week's date (available for subsequent weeks)
            If hasPO And Not colPO Is Nothing Then
                Do While colPO.Count > 0
                    If CDbl(colPO(1)(0)) <= dateKey Then
                        bal = bal + SafeDbl(colPO(1)(1))
                        colPO.Remove 1
                    Else
                        Exit Do
                    End If
                Loop
            End If

            ' 3) If not covered at start and not yet whole, check when SS becomes whole
            If ssWholeDate = vbNullString And ssCovered = "No" And safetyStock > 0 Then
                If bal >= safetyStock Then
                    ssWholeDate = CDate(dateKey)
                    ' do not exit; still need proper WeeksOK / DeadDate
                End If
            End If
        Next w

        ' --- write row
        wsOut.Cells(outRow, 1).Value = comp
        wsOut.Cells(outRow, 2).Value = startInv
        wsOut.Cells(outRow, 3).Value = leadTimeMonths
        wsOut.Cells(outRow, 4).Value = safetyStock

        If safetyStock > 0 Then
            wsOut.Cells(outRow, 5).Value = ssCovered
        Else
            wsOut.Cells(outRow, 5).Value = "Yes" ' trivial case; no SS required
        End If

        If LenB(ssWholeDate) <> 0 Then
            wsOut.Cells(outRow, 6).Value = ssWholeDate
            wsOut.Cells(outRow, 6).NumberFormat = "m/d/yyyy"
        End If

        wsOut.Cells(outRow, 7).Value = weeksOK

        If LenB(deadDate) <> 0 Then
            wsOut.Cells(outRow, 8).Value = deadDate
            wsOut.Cells(outRow, 8).NumberFormat = "m/d/yyyy"
        End If

        outRow = outRow + 1
    Next kComp

    ' Style as a table
    With wsOut.ListObjects.Add(xlSrcRange, wsOut.Range("A1").CurrentRegion, , xlYes)
        .name = "tblComponentHealth"
        .TableStyle = "TableStyleMedium2"
    End With
    wsOut.Columns("A:H").AutoFit

CleanExit:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Exit Sub

CleanFail:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Err.Raise Err.Number, "BuildComponentHealth", Err.Description
End Sub
Private Function ClonePOCollection(ByVal src As Collection) As Collection
    Dim i As Long
    Dim newCol As New Collection

    If Not src Is Nothing Then
        For i = 1 To src.Count
            If IsArray(src(i)) And UBound(src(i)) >= 1 Then
                newCol.Add Array(src(i)(0), src(i)(1))
            End If
        Next i
    End If

    Set ClonePOCollection = newCol
End Function


' Returns Dict(leafComponent -> qtyPerUnit) for any SKU/pre-assembly
Private Function FlattenBOM(ByVal sku As String, _
                            ByVal children As Object, _
                            ByVal parents As Object, _
                            ByVal cache As Object, _
                            Optional ByVal path As Object = Nothing) As Object
    Dim result As Object, k As Variant, pair As Variant
    Dim childID As String, mult As Double, sk As String

    sk = KeyNorm(sku)

    ' memoize
    If cache.exists(sk) Then
        Set FlattenBOM = cache(sk)
        Exit Function
    End If

    ' cycle guard
    If path Is Nothing Then Set path = CreateObject("Scripting.Dictionary")
    If path.exists(sk) Then
        Set FlattenBOM = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    path(sk) = True

    Set result = CreateObject("Scripting.Dictionary")

    If parents.exists(sk) Then
        ' This SKU is a parent/pre-assembly ? expand its children
        If children.exists(sk) Then
            For Each pair In children(sk)
                childID = KeyNorm(pair(0))
                mult = SafeDbl(pair(1)): If mult = 0 Then mult = 1

                If parents.exists(childID) Then
                    ' child is itself a pre-assembly ? recurse
                    Dim flatChild As Object, fk As Variant
                    Set flatChild = FlattenBOM(childID, children, parents, cache, path)
                    For Each fk In flatChild.keys
                        AddToDict result, CStr(fk), mult * SafeDbl(flatChild(fk))
                    Next fk
                Else
                    ' true leaf component
                    AddToDict result, childID, mult
                End If
            Next pair
        End If
    Else
        ' Not a parent anywhere ? this is already a leaf part
        AddToDict result, sk, 1
    End If

    path.Remove sk
    Set cache(sk) = result
    Set FlattenBOM = result
End Function


'========================
' UTILITIES
'========================
Private Sub AddToDict(ByVal d As Object, ByVal key As String, ByVal addVal As Double)
    If d.exists(key) Then
        d(key) = SafeDbl(d(key)) + addVal
    Else
        d.Add key, addVal
    End If
End Sub
' Sorts a Collection of PO records where each item is Array(dateKey As Double, qty As Double)
Private Sub SortPOCollectionByDate(ByRef col As Collection)
    Dim n As Long, i As Long, j As Long
    Dim dk() As Double, q() As Double
    Dim tmpD As Double, tmpQ As Double

    If col Is Nothing Then Exit Sub
    n = col.Count
    If n <= 1 Then Exit Sub

    ReDim dk(1 To n)
    ReDim q(1 To n)

    ' Copy out to parallel arrays
    For i = 1 To n
        ' Guard against malformed items
        If IsArray(col(i)) And UBound(col(i)) >= 1 Then
            dk(i) = CDbl(col(i)(0))
            q(i) = CDbl(col(i)(1))
        Else
            ' If an item isn't the expected 2-element array, treat as zero-date, zero-qty
            dk(i) = 0#
            q(i) = 0#
        End If
    Next i

    ' Insertion sort by date key
 For i = 2 To n
    tmpD = dk(i): tmpQ = q(i)
    j = i - 1

    Do While j >= 1
        If dk(j) > tmpD Then
            dk(j + 1) = dk(j)
            q(j + 1) = q(j)
            j = j - 1
        Else
            Exit Do
        End If
    Loop

    dk(j + 1) = tmpD
    q(j + 1) = tmpQ
Next i

    ' Rebuild collection in sorted order
    Dim newCol As New Collection
    For i = 1 To n
        newCol.Add Array(dk(i), q(i))
    Next i
    Set col = newCol
End Sub

Private Function KeyNorm(ByVal s As Variant) As String
    Dim t As String
    If IsError(s) Or IsNull(s) Then
        KeyNorm = ""
        Exit Function
    End If
    t = Trim$(CStr(s))
    ' collapse multiple spaces
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    KeyNorm = UCase$(t)
End Function
Public Sub BuildComponentWeeklyInventoryAudit()
    Const SH_BOM As String = "BOM"
    Const BOM_COL_SKU As Long = 1
    Const BOM_COL_COMP As Long = 2
    Const BOM_COL_QTY As Long = 3

    Const SH_INPUT As String = "Input"
    Const MX_ROW_DATES As Long = 2
    Const MX_ROW_FIRST As Long = 4
    Const MX_COL_SKU As Long = 1

    Const SH_INV As String = "Current Inventory"
    Const INV_COL_SKU As Long = 1
    Const INV_COL_QTY As Long = 2

    Const SH_PO As String = "Items on PO"
    Const PO_COL_SKU As Long = 1
    Const PO_COL_DATE As Long = 2
    Const PO_COL_QTY As Long = 3

    Const SH_AUDIT As String = "Component Weekly Inv"

    Dim scrn&, evt&, calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Dim wsB As Worksheet, wsM As Worksheet, wsInv As Worksheet, wsPO As Worksheet, wsA As Worksheet
    Set wsB = ThisWorkbook.Worksheets(SH_BOM)
    Set wsM = ThisWorkbook.Worksheets(SH_INPUT)
    Set wsInv = ThisWorkbook.Worksheets(SH_INV)
    Set wsPO = ThisWorkbook.Worksheets(SH_PO)

    ' (Re)create audit sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(SH_AUDIT).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsA = ThisWorkbook.Worksheets.Add(After:=wsPO)
    wsA.name = SH_AUDIT

    ' ---------- BOM: parents + children ----------
    Dim parents As Object: Set parents = CreateObject("Scripting.Dictionary")
    Dim children As Object: Set children = CreateObject("Scripting.Dictionary")
    Dim r As Long, lastB As Long, parent As String, comp As String, qPer As Double
    lastB = wsB.Cells(wsB.Rows.Count, BOM_COL_SKU).End(xlUp).row
    For r = 2 To lastB
        parent = KeyNorm(wsB.Cells(r, BOM_COL_SKU).Value)
        comp = KeyNorm(wsB.Cells(r, BOM_COL_COMP).Value)
        qPer = SafeDbl(wsB.Cells(r, BOM_COL_QTY).Value): If qPer = 0 Then qPer = 1
        If Len(parent) > 0 And Len(comp) > 0 Then
            parents(parent) = True
            If Not children.exists(parent) Then children.Add parent, New Collection
            children(parent).Add Array(comp, qPer)
        End If
    Next r

    ' ---------- Planning horizon ----------
    Dim lastCol As Long, c As Long
    Dim weekCols As New Collection, weekKeys As New Collection
    lastCol = wsM.Cells(MX_ROW_DATES, wsM.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If IsDate(wsM.Cells(MX_ROW_DATES, c).Value) Then
            weekCols.Add c
            weekKeys.Add CLng(CDate(wsM.Cells(MX_ROW_DATES, c).Value))
        End If
    Next c
    If weekCols.Count = 0 Then GoTo CleanExit

    ' Headers
    wsA.Cells(1, 1).Value = "Component ID"
    wsA.Cells(1, 2).Value = "Start Inv"
    For c = 1 To weekCols.Count
        wsA.Cells(1, 2 + c).Value = CDate(weekKeys(c))
        wsA.Cells(1, 2 + c).NumberFormat = "m/d/yyyy"
    Next c
    wsA.Rows(1).Font.Bold = True

    ' ---------- Flatten cache and weekly leaf needs ----------
    Dim flatCache As Object: Set flatCache = CreateObject("Scripting.Dictionary")
    Dim compNeeds As Object: Set compNeeds = CreateObject("Scripting.Dictionary") ' key=comp|dateKey

    Dim lastRow As Long, rr As Long, sku As String, wkCol As Long, wkKey As Long, skuQty As Double, addKey As String
    lastRow = wsM.Cells(wsM.Rows.Count, MX_COL_SKU).End(xlUp).row

    For rr = MX_ROW_FIRST To lastRow
        sku = KeyNorm(wsM.Cells(rr, MX_COL_SKU).Value)
        If Len(sku) = 0 Then GoTo NextRow

        Dim perUnit As Object, k As Variant, i As Long
        Set perUnit = FlattenBOM(sku, children, parents, flatCache)
        If Not perUnit Is Nothing And perUnit.Count > 0 Then
            For i = 1 To weekCols.Count
                wkCol = weekCols(i): wkKey = CLng(weekKeys(i))
                skuQty = SafeDbl(wsM.Cells(rr, wkCol).Value)
                If skuQty <> 0 Then
                    For Each k In perUnit.keys
                        addKey = CStr(k) & "|" & wkKey
                        If compNeeds.exists(addKey) Then
                            compNeeds(addKey) = compNeeds(addKey) + skuQty * SafeDbl(perUnit(k))
                        Else
                            compNeeds.Add addKey, skuQty * SafeDbl(perUnit(k))
                        End If
                    Next k
                End If
            Next i
        End If
NextRow:
    Next rr

    ' ---------- Inventory and PO ----------
    Dim inv As Object: Set inv = CreateObject("Scripting.Dictionary")
    Dim lastInv As Long: lastInv = wsInv.Cells(wsInv.Rows.Count, INV_COL_SKU).End(xlUp).row
    For r = 2 To lastInv
        comp = KeyNorm(wsInv.Cells(r, INV_COL_SKU).Value)
        If Len(comp) > 0 Then inv(comp) = SafeDbl(wsInv.Cells(r, INV_COL_QTY).Value)
    Next r

    Dim po As Object: Set po = CreateObject("Scripting.Dictionary") ' comp -> Collection(dateKey, qty)
    Dim lastPO As Long, dt As Variant, qPO As Double
    lastPO = wsPO.Cells(wsPO.Rows.Count, PO_COL_SKU).End(xlUp).row
    For r = 2 To lastPO
        comp = KeyNorm(wsPO.Cells(r, PO_COL_SKU).Value)
        dt = wsPO.Cells(r, PO_COL_DATE).Value
        qPO = SafeDbl(wsPO.Cells(r, PO_COL_QTY).Value)
        If Len(comp) > 0 And IsDate(dt) And qPO <> 0 Then
            If Not po.exists(comp) Then po.Add comp, New Collection
            po(comp).Add Array(CLng(CDate(dt)), qPO)
        End If
    Next r
    Dim kComp As Variant
    For Each kComp In po.keys
        SortPOCollectionByDate po(kComp)
    Next kComp

    ' ---------- Component set to report (leaf only) ----------
    Dim compSet As Object: Set compSet = CreateObject("Scripting.Dictionary")
    Dim key As Variant, compOnly As Variant
    For Each key In compNeeds.keys
        compOnly = Split(CStr(key), "|")(0)
        If Not parents.exists(compOnly) Then compSet(compOnly) = True
    Next key
    For Each compOnly In inv.keys
        If Not parents.exists(compOnly) Then compSet(compOnly) = True
    Next compOnly

    ' ---------- Simulate and write weekly ending balance ----------
    Dim outRow As Long: outRow = 2
    Dim startInv As Double, bal As Double, w As Long, need As Double, dateKey As Long
    Dim colPO As Collection, hasPO As Boolean

    For Each kComp In compSet.keys
        comp = CStr(kComp)
        startInv = 0: If inv.exists(comp) Then startInv = SafeDbl(inv(comp))
        bal = startInv

        wsA.Cells(outRow, 1).Value = comp
        wsA.Cells(outRow, 2).Value = startInv

        hasPO = po.exists(comp)
        If hasPO Then Set colPO = po(comp)

        For w = 1 To weekCols.Count
            dateKey = CLng(weekKeys(w))
            key = comp & "|" & dateKey
            need = 0: If compNeeds.exists(key) Then need = SafeDbl(compNeeds(key))

            ' consume this week
            bal = bal - need

            ' add receipts with ETA <= this week's date (available before next week)
            If hasPO Then
                Do While colPO.Count > 0
                    If CDbl(colPO(1)(0)) <= dateKey Then
                        bal = bal + SafeDbl(colPO(1)(1))
                        colPO.Remove 1
                    Else
                        Exit Do
                    End If
                Loop
            End If

            ' record ENDING balance for the week
            wsA.Cells(outRow, 2 + w).Value = bal
        Next w

        outRow = outRow + 1
    Next kComp

    ' Format as table
    With wsA.Range("A1").CurrentRegion
        .Columns.AutoFit
    End With
    On Error Resume Next
    wsA.ListObjects.Add(xlSrcRange, wsA.Range("A1").CurrentRegion, , xlYes).name = "tblComponentWeeklyInv"
    On Error GoTo 0
    If wsA.ListObjects.Count > 0 Then wsA.ListObjects(1).TableStyle = "TableStyleMedium2"

CleanExit:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Exit Sub

CleanFail:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Err.Raise Err.Number, "BuildComponentWeeklyInventoryAudit", Err.Description
End Sub


Public Sub SendRevoComponentHealthReport()
    Const SH_HEALTH As String = "Component Health"
    Const SH_WEEKLY As String = "Component Weekly Inv"

    Dim ws As Worksheet, wsWeekly As Worksheet
    On Error GoTo fail

    Set ws = ThisWorkbook.Worksheets(SH_HEALTH)
    Set wsWeekly = ThisWorkbook.Worksheets(SH_WEEKLY)

    '--- grab and sort the data in-memory (A1:Hlast)
    Dim lastRow As Long, arr, hdr(), data()
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No rows found on '" & SH_HEALTH & "'.", vbExclamation
        Exit Sub
    End If

    ' Full Component Health table
    arr = ws.Range("A1:H" & lastRow).Value2   ' 1-based 2D array

    ' Headers + extra Comments column
 hdr = Array( _
    arr(1, 1), _
    arr(1, 2), _
    arr(1, 3), _
    arr(1, 4), _
    arr(1, 5), _
    arr(1, 6), _
    arr(1, 7), _
    arr(1, 8), _
    "Comments" _
)


    ' slice data rows (2..n) into a new array for sorting
    Dim n As Long, i As Long
    n = lastRow - 1
    ReDim data(1 To n, 1 To 8)

    For i = 1 To n
        data(i, 1) = arr(i + 1, 1) ' Component ID
        data(i, 2) = arr(i + 1, 2) ' Inventory
        data(i, 3) = arr(i + 1, 3) ' Lead Time (Months)
        data(i, 4) = arr(i + 1, 4) ' Safety Stock
        data(i, 5) = arr(i + 1, 5) ' Safety Stock Covered?
        data(i, 6) = arr(i + 1, 6) ' Safety Stock Whole Date
        data(i, 7) = arr(i + 1, 7) ' Weeks of Production Available
        data(i, 8) = arr(i + 1, 8) ' Inv Dead Date
    Next i

    ' sort by Inv Dead Date (col 8) ascending; blanks to bottom
    SortDataByDateCol data, 8, True

    ' build HTML table string
    Dim htmlTable As String
    htmlTable = BuildHealthHtmlTable(hdr, data)

    '--- create temp workbook with the two sheets
    Dim wbTemp As Workbook
    ws.Copy    ' creates new workbook with just Component Health
    Set wbTemp = ActiveWorkbook
    wsWeekly.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)

    ' --- save to temp path (unique file name)
    Dim tmpPath As String, filePath As String, ts As String
    tmpPath = Environ$("TEMP")
    If Right$(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
    ts = Format(Now, "yyyy-mm-dd_hhmmss")
    filePath = tmpPath & "Component Health Report (" & ts & ").xlsx"

    Application.DisplayAlerts = False
    wbTemp.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    '--- build and display Outlook email
    Dim olApp As Object, olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0) ' olMailItem

    Dim intro As String, subj As String
    subj = "REVO Component Health Report - " & Format(Date, "mmm d, yyyy")

    intro = _
        "<p>Hello Cyrille,</p>" & _
        "<p>Please see the below table of REVO Components with lead times, safety stock coverage, weeks of inventory " & _
        "based on current plus inbound supply, and the projected inventory dead dates. " & _
        "The attached workbook includes this summary and the weekly inventory view for each component.</p>" & _
        htmlTable & _
        "<p>Please let me know if you have any questions or concerns about this report.</p>"

    With olMail
        .To = "cfournier@predatorgroup.com"
        .CC = "amccoy@predatorgroup.com"
        .Subject = subj
        .HTMLBody = intro
        .Attachments.Add filePath
        .Display
    End With

    Exit Sub

fail:
    Application.DisplayAlerts = True
    MsgBox "SendRevoComponentHealthReport failed: " & Err.Description, vbCritical
End Sub


'========================
' Helpers
'========================

' Build a simple HTML table from headers + data (adds an empty Comments column at the end, auto-fit width)
' Build HTML table from headers + data (adds a blank Comments col at the end)
Private Function BuildHealthHtmlTable(ByVal headers As Variant, ByRef data As Variant) As String
    Dim sb As String, i As Long, n As Long
    n = UBound(data, 1)

    sb = sb & "<table border='1' cellpadding='4' cellspacing='0' " & _
              "style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;" & _
              "font-size:10pt;width:auto;white-space:nowrap;'>"

    ' header row
    sb = sb & "<tr style='background:#e6f0ff;font-weight:bold;text-align:center'>"
    Dim j As Long
    For j = LBound(headers) To UBound(headers)
        sb = sb & "<td style='padding:4px 8px;'>" & HtmlEncode(CStr(headers(j))) & "</td>"
    Next j
    sb = sb & "</tr>"

    ' data rows
    For i = 1 To n
        sb = sb & "<tr>"

        ' 1: Component ID
        sb = sb & "<td style='padding:4px 8px;'>" & HtmlEncode(NzText(data(i, 1))) & "</td>"

        ' 2: Inventory
        sb = sb & "<td style='text-align:right;padding:4px 8px;'>" & HtmlEncode(FormatNumberCustom(data(i, 2))) & "</td>"

        ' 3: Lead Time (Months)
        sb = sb & "<td style='text-align:right;padding:4px 8px;'>" & HtmlEncode(FormatNumberCustom(data(i, 3))) & "</td>"

        ' 4: Safety Stock
        sb = sb & "<td style='text-align:right;padding:4px 8px;'>" & HtmlEncode(FormatNumberCustom(data(i, 4))) & "</td>"

        ' 5: Safety Stock Covered? (Yes/No)
        sb = sb & "<td style='text-align:center;padding:4px 8px;'>" & HtmlEncode(NzText(data(i, 5))) & "</td>"

        ' 6: Safety Stock Whole Date
        If IsEmpty(data(i, 6)) Or Len(Trim$(CStr(data(i, 6)))) = 0 Then
            sb = sb & "<td style='text-align:right;padding:4px 8px;'></td>"
        Else
            sb = sb & "<td style='text-align:right;padding:4px 8px;'>" & _
                      HtmlEncode(Format(CDate(data(i, 6)), "m/d/yyyy")) & "</td>"
        End If

        ' 7: Weeks of Production Available
        sb = sb & "<td style='text-align:right;padding:4px 8px;'>" & HtmlEncode(FormatNumberCustom(data(i, 7))) & "</td>"

        ' 8: Inv Dead Date
        If IsEmpty(data(i, 8)) Or Len(Trim$(CStr(data(i, 8)))) = 0 Then
            sb = sb & "<td style='text-align:right;padding:4px 8px;'></td>"
        Else
            sb = sb & "<td style='text-align:right;padding:4px 8px;'>" & _
                      HtmlEncode(Format(CDate(data(i, 8)), "m/d/yyyy")) & "</td>"
        End If

        ' 9: Comments (blank)
        sb = sb & "<td style='min-width:160px;padding:4px 8px;'>&nbsp;</td>"

        sb = sb & "</tr>"
    Next i

    sb = sb & "</table>"
    BuildHealthHtmlTable = sb
End Function




' Sort a 2D (1..N, 1..M) array by a date column; blanks go to bottom
' Sort a 2D Variant array (1..N, 1..M) by a "date" column.
' Oldest ? newest when ascending:=True. Blanks/non-dates go last.
Private Sub SortDataByDateCol(ByRef data As Variant, ByVal dateCol As Long, _
                              Optional ByVal ascending As Boolean = True)
    Dim n As Long, m As Long, i As Long, j As Long
    n = UBound(data, 1)
    If n <= 1 Then Exit Sub
    m = UBound(data, 2)

    ' Build keys and an index array so we don't shuffle rows during comparisons
    Dim keys() As Double, idx() As Long
    ReDim keys(1 To n)
    ReDim idx(1 To n)
    For i = 1 To n
        keys(i) = dateKey(data(i, dateCol))  ' blanks -> huge sentinel
        idx(i) = i
    Next i

    ' Insertion sort the index array by the keys
    Dim ins As Long, keyI As Double
    For i = 2 To n
        ins = idx(i)
        keyI = keys(ins)
        j = i - 1
        Do While j >= 1
            If ascending Then
                If keys(idx(j)) > keyI Then
                    idx(j + 1) = idx(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            Else
                If keys(idx(j)) < keyI Then
                    idx(j + 1) = idx(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            End If
        Loop
        idx(j + 1) = ins
    Next i

    ' Rebuild data in sorted order
    Dim sorted() As Variant, c As Long, r As Long
    ReDim sorted(1 To n, 1 To m)
    For i = 1 To n
        r = idx(i)
        For c = 1 To m
            sorted(i, c) = data(r, c)
        Next c
    Next i
    data = sorted
End Sub

' Convert a cell value to a sortable numeric date; blanks/non-dates -> very large number
Private Function dateKey(ByVal v As Variant) As Double
    If IsDate(v) Then
        dateKey = CDbl(CDate(v))
    Else
        dateKey = 1E+308   ' sentinal so blank/non-date sorts last
    End If
End Function




Private Function NzText(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        NzText = ""
    Else
        NzText = CStr(v)
    End If
End Function



Private Function FormatNumberCustom(ByVal v As Variant) As String
    If IsNumeric(v) Then
        FormatNumberCustom = Format$(CDbl(v), "#,##0") ' no decimals; adjust if needed
    Else
        FormatNumberCustom = ""
    End If
End Function

Public Sub BuildFlattenedBOMMatrix()

    Const SH_MATRIX As String = "BOM Matrix" ' output sheet name

    Dim scrn&, evt&, calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim wsB As Worksheet, wsM As Worksheet, wsOut As Worksheet

    Set wb = ThisWorkbook
    Set wsB = wb.Worksheets(SH_BOM)
    Set wsM = wb.Worksheets(SH_INPUT)

    ' --- Rebuild BOM parents/children exactly like BuildComponentHealth ---

    Dim parents As Object: Set parents = CreateObject("Scripting.Dictionary")        ' parent SKU -> True
    Dim children As Object: Set children = CreateObject("Scripting.Dictionary")      ' parent SKU -> Collection(childID, qtyPer)

    Dim lastB As Long, r As Long
    Dim parent As String, comp As String
    Dim qPer As Double

    lastB = wsB.Cells(wsB.Rows.Count, BOM_COL_SKU).End(xlUp).row

    For r = 2 To lastB
        parent = KeyNorm(wsB.Cells(r, BOM_COL_SKU).Value)
        comp = KeyNorm(wsB.Cells(r, BOM_COL_COMP).Value)
        qPer = SafeDbl(wsB.Cells(r, BOM_COL_QTY).Value)
        If qPer = 0 Then qPer = 1

        If Len(parent) > 0 And Len(comp) > 0 Then
            parents(parent) = True
            If Not children.exists(parent) Then children.Add parent, New Collection
            children(parent).Add Array(comp, qPer)
        End If
    Next r

    ' --- Collect unique SKUs from the input matrix (same drivers as component health) ---

    Dim lastM As Long, rr As Long
    Dim sku As String
    Dim skuSet As Object: Set skuSet = CreateObject("Scripting.Dictionary")
    skuSet.CompareMode = vbTextCompare

    lastM = wsM.Cells(wsM.Rows.Count, MX_COL_SKU).End(xlUp).row

    For rr = MX_ROW_FIRST To lastM
        sku = KeyNorm(wsM.Cells(rr, MX_COL_SKU).Value)
        If Len(sku) > 0 Then
            If Not skuSet.exists(sku) Then skuSet.Add sku, True
        End If
    Next rr

    If skuSet.Count = 0 Then GoTo CleanExit

    ' --- Flatten cache + collect all leaf components used by these SKUs ---

    Dim flatCache As Object: Set flatCache = CreateObject("Scripting.Dictionary")
    Dim allComps As Object: Set allComps = CreateObject("Scripting.Dictionary")
    allComps.CompareMode = vbTextCompare

    Dim perUnit As Object
    Dim k As Variant

    For Each k In skuSet.keys
        Set perUnit = FlattenBOM(CStr(k), children, parents, flatCache)
        If Not perUnit Is Nothing Then
            Dim leaf As Variant
            For Each leaf In perUnit.keys
                If Not allComps.exists(CStr(leaf)) Then
                    allComps.Add CStr(leaf), allComps.Count + 1 ' store index placeholder
                End If
            Next leaf
        End If
    Next k

    If allComps.Count = 0 Then GoTo CleanExit

    ' --- Create / reset output sheet ---

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(SH_MATRIX).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = wb.Worksheets.Add(After:=wsM)
    wsOut.name = SH_MATRIX

    ' --- Build header row: SKU + one col per component ---

    Dim compIndex As Object: Set compIndex = CreateObject("Scripting.Dictionary")
    compIndex.CompareMode = vbTextCompare

    Dim col As Long
    col = 1
    wsOut.Cells(1, col).Value = "SKU"
    col = col + 1

    ' Optional: sort component IDs alphabetically for readability
    Dim compArr() As String
    ReDim compArr(1 To allComps.Count)
    Dim i As Long
    i = 1
    For Each k In allComps.keys
        compArr(i) = CStr(k)
        i = i + 1
    Next k
    Call QuickSortStrings(compArr, 1, UBound(compArr))

    For i = 1 To UBound(compArr)
        wsOut.Cells(1, col).Value = compArr(i)
        compIndex(compArr(i)) = col
        col = col + 1
    Next i

    Dim lastCol As Long
    lastCol = col - 1

    wsOut.Rows(1).Font.Bold = True

    ' --- Fill matrix: each SKU row, flattened qty-per-unit in component columns ---

    Dim rowOut As Long
    rowOut = 2

    For Each k In skuSet.keys
        sku = CStr(k)
        Set perUnit = FlattenBOM(sku, children, parents, flatCache)

        ' Only write rows that have at least some mapping
        If Not perUnit Is Nothing And perUnit.Count > 0 Then
            wsOut.Cells(rowOut, 1).Value = sku

            Dim compID As Variant, qty As Double, compCol As Long

            For Each compID In perUnit.keys
                qty = SafeDbl(perUnit(compID))
                If qty <> 0 Then
                    If compIndex.exists(CStr(compID)) Then
                        compCol = compIndex(CStr(compID))
                        wsOut.Cells(rowOut, compCol).Value = qty
                    End If
                End If
            Next compID

            rowOut = rowOut + 1
        End If
    Next k

    ' --- Format nicely ---
    With wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(rowOut - 1, lastCol))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    wsOut.Columns(1).HorizontalAlignment = xlLeft
    wsOut.Columns("A").EntireColumn.AutoFit

    ' auto-fit component columns but keep them readable
    For col = 2 To lastCol
        wsOut.Columns(col).ColumnWidth = Application.WorksheetFunction.Max( _
            6, Len(wsOut.Cells(1, col).Value2) + 1)
    Next col

CleanExit:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Exit Sub

CleanFail:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    MsgBox "BuildFlattenedBOMMatrix failed: " & Err.Description, vbCritical
End Sub

Private Sub QuickSortStrings(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As String

    i = first: j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub
Public Sub BuildBOMHierarchyView()
    Const OUT_SHEET As String = "BOM Hierarchy View"
    Const COL_PROD_TYPE As Long = 6   ' F: Product_Type (applies to SKU/parent only)

    Dim scrn&, evt&, calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim wsB As Worksheet, wsOut As Worksheet
    Set wb = ThisWorkbook
    Set wsB = wb.Worksheets(SH_BOM)

    Dim lastB As Long
    lastB = wsB.Cells(wsB.Rows.Count, BOM_COL_SKU).End(xlUp).row
    If lastB < 2 Then GoTo CleanExit

    '========================
    ' 1) Build parents -> children map from BOM
    '========================
    Dim children As Object
    Set children = CreateObject("Scripting.Dictionary")
    children.CompareMode = vbTextCompare

    Dim parents As Object
    Set parents = CreateObject("Scripting.Dictionary")
    parents.CompareMode = vbTextCompare

    Dim r As Long
    Dim parentID As String, compID As String
    Dim qPer As Double

    For r = 2 To lastB
        parentID = KeyNorm(wsB.Cells(r, BOM_COL_SKU).Value)
        compID = KeyNorm(wsB.Cells(r, BOM_COL_COMP).Value)
        qPer = SafeDbl(wsB.Cells(r, BOM_COL_QTY).Value)
        If qPer = 0 Then qPer = 1

        If Len(parentID) > 0 And Len(compID) > 0 Then
            If Not parents.exists(parentID) Then parents.Add parentID, True
            If Not children.exists(parentID) Then children.Add parentID, New Collection
            children(parentID).Add Array(compID, qPer)
        End If
    Next r

    '========================
    ' 2) Product_Type map FROM SKU COLUMN ONLY
    '========================
    Dim prodType As Object
    Set prodType = CreateObject("Scripting.Dictionary")
    prodType.CompareMode = vbTextCompare

    Dim pt As String

    For r = 2 To lastB
        parentID = KeyNorm(wsB.Cells(r, BOM_COL_SKU).Value)
        If Len(parentID) > 0 Then
            pt = UCase$(Trim$(CStr(wsB.Cells(r, COL_PROD_TYPE).Value)))
            If pt <> "" Then
                ' First non-blank wins; do NOT assign based on Component_ID
                If Not prodType.exists(parentID) Then
                    prodType.Add parentID, pt
                End If
            End If
        End If
    Next r

    ' Build classification sets
    Dim preAsm As Object, finished As Object
    Set preAsm = CreateObject("Scripting.Dictionary")
    Set finished = CreateObject("Scripting.Dictionary")
    preAsm.CompareMode = vbTextCompare
    finished.CompareMode = vbTextCompare

    Dim k As Variant
    For Each k In prodType.keys
        Select Case prodType(k)
            Case "PREASSEMBLY"
                preAsm(k) = True
            Case "FINISHED SHAFT"
                finished(k) = True
        End Select
    Next k

    If finished.Count = 0 Then
        MsgBox "No SKUs marked as FINISHED SHAFT in BOM Product_Type column.", vbInformation
        GoTo CleanExit
    End If

    '========================
    ' 3) Create / reset output sheet
    '========================
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(OUT_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = wb.Worksheets.Add(After:=wsB)
    wsOut.name = OUT_SHEET

    ' Headers
    wsOut.Cells(1, 1).Value = "Root SKU"
    wsOut.Cells(1, 2).Value = "Item"
    wsOut.Cells(1, 3).Value = "Qty per Root"
    wsOut.Cells(1, 4).Value = "Level"
    wsOut.Cells(1, 5).Value = "Type"
    wsOut.Rows(1).Font.Bold = True

    Dim rowOut As Long: rowOut = 2

    ' Sorted list of finished shafts (Loop 1)
    Dim fsArr() As String, i As Long
    ReDim fsArr(1 To finished.Count)
    i = 1
    For Each k In finished.keys
        fsArr(i) = CStr(k)
        i = i + 1
    Next k
    QuickSortStrings fsArr, 1, UBound(fsArr)

    Dim compItem As Variant, subItem As Variant
    Dim childCol As Collection, subCol As Collection
    Dim qtyFSComp As Double, qtyPreSub As Double, qtyRootSub As Double
    Dim fs As String

    '========================
    ' Loop 1 + 2 + 3 hierarchy
    '========================
    For i = 1 To UBound(fsArr)
        fs = fsArr(i)

        ' Level 1: FINISHED SHAFT
        wsOut.Cells(rowOut, 1).Value = fs
        wsOut.Cells(rowOut, 2).Value = fs
        wsOut.Cells(rowOut, 3).Value = 1
        wsOut.Cells(rowOut, 4).Value = 1
        wsOut.Cells(rowOut, 5).Value = "FINISHED SHAFT"
        wsOut.Rows(rowOut).Font.Bold = True
        rowOut = rowOut + 1

        ' Children of FS
        If children.exists(fs) Then
            Set childCol = children(fs)

            ' First pass under FS: PREASSEMBLY children (Loop 2)
            For Each compItem In childCol
                compID = CStr(compItem(0))
                qtyFSComp = SafeDbl(compItem(1))
                If qtyFSComp = 0 Then qtyFSComp = 1

                If preAsm.exists(compID) Then
                    ' Level 2 PREASSEMBLY
                    wsOut.Cells(rowOut, 1).Value = fs
                    wsOut.Cells(rowOut, 2).Value = compID
                    wsOut.Cells(rowOut, 3).Value = qtyFSComp
                    wsOut.Cells(rowOut, 4).Value = 2
                    wsOut.Cells(rowOut, 5).Value = "PREASSEMBLY"
                    rowOut = rowOut + 1

                    ' Level 3: SUBCOMPONENTs of this PREASSEMBLY (Loop 3 partial)
                    If children.exists(compID) Then
                        Set subCol = children(compID)
                        For Each subItem In subCol
                            Dim subID As String
                            subID = CStr(subItem(0))
                            qtyPreSub = SafeDbl(subItem(1))
                            If qtyPreSub = 0 Then qtyPreSub = 1

                            qtyRootSub = qtyFSComp * qtyPreSub

                            wsOut.Cells(rowOut, 1).Value = fs
                            wsOut.Cells(rowOut, 2).Value = "    " & subID
                            wsOut.Cells(rowOut, 3).Value = qtyRootSub
                            wsOut.Cells(rowOut, 4).Value = 3
                            wsOut.Cells(rowOut, 5).Value = "SUBCOMPONENT"
                            rowOut = rowOut + 1
                        Next subItem
                    End If
                End If
            Next compItem

            ' Second pass under FS: non-preassembly children as direct Level 2 components (Loop 3 remainder)
            For Each compItem In childCol
                compID = CStr(compItem(0))
                qtyFSComp = SafeDbl(compItem(1))
                If qtyFSComp = 0 Then qtyFSComp = 1

                ' Only if NOT a preassembly SKU
                If Not preAsm.exists(compID) Then
                    wsOut.Cells(rowOut, 1).Value = fs
                    wsOut.Cells(rowOut, 2).Value = compID
                    wsOut.Cells(rowOut, 3).Value = qtyFSComp
                    wsOut.Cells(rowOut, 4).Value = 2
                    wsOut.Cells(rowOut, 5).Value = "COMPONENT"
                    rowOut = rowOut + 1
                End If
            Next compItem
        End If

        ' Blank line between Finished Shafts
        rowOut = rowOut + 1
    Next i

    '========================
    ' Format
    '========================
    With wsOut.Range("A1:E" & rowOut - 1)
        .Borders.LineStyle = xlContinuous
        .Columns(1).HorizontalAlignment = xlLeft
        .Columns(2).HorizontalAlignment = xlLeft
        .Columns(3).HorizontalAlignment = xlRight
        .Columns(4).HorizontalAlignment = xlCenter
        .Columns(5).HorizontalAlignment = xlCenter
        .EntireColumn.AutoFit
    End With

CleanExit:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Exit Sub

CleanFail:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    MsgBox "BuildBOMHierarchyView failed: " & Err.Description, vbCritical
    Resume CleanExit
End Sub


Private Sub GetForecastWindow_FromProdForecast(ByRef fStart As Date, ByRef fEnd As Date)
    Dim wsF As Worksheet: Set wsF = GetSheet(SH_SRC) ' "Production Forecast"

    Dim lastCol As Long: lastCol = wsF.Cells(2, wsF.Columns.Count).End(xlToLeft).Column
    Dim c As Long, d As Variant

    Dim firstDate As Date: firstDate = 0
    Dim chosen As Date: chosen = 0

    For c = 2 To lastCol
        d = wsF.Cells(2, c).Value
        If IsDate(d) Then
            If firstDate = 0 Then firstDate = CDate(d)
            If chosen = 0 Then
                If CDate(d) >= Date Then chosen = CDate(d)
            End If
        End If
    Next c

    If chosen = 0 Then chosen = firstDate
    If chosen = 0 Then
        ' No dates found; fall back to current week
        chosen = DateAdd("d", 1 - Weekday(Date, vbMonday), Date)
    End If

    fStart = chosen
    fEnd = DateAdd("d", 27, fStart) ' 4 weeks window
End Sub


Public Sub ProdTable_NormalizeDates_FromCutoff(ByVal cutoff As Date)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Production Table")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Dim r As Long, v As Variant
    For r = 2 To lastRow
        v = ws.Cells(r, "B").Value
        If IsDate(v) Then
            If CDate(v) >= cutoff Then
                ws.Cells(r, "B").Value = CDbl(CDate(v))
            End If
        End If
    Next r
    ws.Columns("B").NumberFormat = "m/d/yyyy"
End Sub

Private Function AnalysisKeyFromSKU(ByVal sku As String) As String
    Dim s As String: s = UCase$(Trim$(sku))

    ' AIR / CRM override
    If InStr(s, "AIR") > 0 Then AnalysisKeyFromSKU = "AIR": Exit Function
    If InStr(s, "CRM") > 0 Then AnalysisKeyFromSKU = "CRM": Exit Function

    ' BK buckets
    If InStr(s, "BK") > 0 Then
        If InStr(s, "30") > 0 Then
            AnalysisKeyFromSKU = "BK 30"""
        Else
            AnalysisKeyFromSKU = "BK"
        End If
        Exit Function
    End If

    ' 30" buckets (non-BK)
    If InStr(s, "30") > 0 Then
        If InStr(s, "11.8") > 0 Then AnalysisKeyFromSKU = "11.8 30""": Exit Function
        If InStr(s, "12.4") > 0 Then AnalysisKeyFromSKU = "12.4 30""": Exit Function
        If InStr(s, "12.9") > 0 Then AnalysisKeyFromSKU = "12.9 30""": Exit Function
    End If

    ' T buckets (10T/14T/18T variants)
    If InStr(s, "10T") > 0 Or InStr(s, "14T") > 0 Or InStr(s, "18T") > 0 Then
        If InStr(s, "11.8") > 0 Then AnalysisKeyFromSKU = "11.8T": Exit Function
        If InStr(s, "12.4") > 0 Then AnalysisKeyFromSKU = "12.4 T": Exit Function
        If InStr(s, "12.9") > 0 Then AnalysisKeyFromSKU = "12.9 T": Exit Function
    End If

    ' UNI / RAD buckets
    If InStr(s, "UNI") > 0 Then
        If InStr(s, "11.8") > 0 Then AnalysisKeyFromSKU = "11.8 UNI": Exit Function
        If InStr(s, "12.4") > 0 Then AnalysisKeyFromSKU = "12.4UNI": Exit Function
        If InStr(s, "12.9") > 0 Then AnalysisKeyFromSKU = "12.9 UNI": Exit Function
    End If

    If InStr(s, "RAD") > 0 Then
        If InStr(s, "11.8") > 0 Then AnalysisKeyFromSKU = "11.8 RAD": Exit Function
        If InStr(s, "12.4") > 0 Then AnalysisKeyFromSKU = "12.4 RAD": Exit Function
        If InStr(s, "12.9") > 0 Then AnalysisKeyFromSKU = "12.9 RAD": Exit Function
    End If

    ' Fallback (if anything slips through)
    AnalysisKeyFromSKU = ""
End Function
