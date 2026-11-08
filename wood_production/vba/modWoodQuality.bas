Attribute VB_Name = "modWoodQuality"
Option Explicit

'=====================================================================
' modWoodQuality — QC / release flow for WOOD Floor (Panthera cue butts)
'
' Same design as modTableQuality in the TABLE workbook; the differences are
' the sheet name, that stations are read from Plan Config (WS01..WSnn) instead
' of being fixed at five, and the wording. One row on WOOD Floor = one
' serialised cue butt. The "Wood QC" button runs WoodQC_Process:
'
'   * a data row is selected            -> QC that one cue
'   * anything else is selected         -> QC every cue whose Release? = TRUE
'                                          and whose QC Status is still blank
'
' For each cue the frmWoodQuality dialog captures: serial number, outcome
' (Release / B-Grade / Rework / Reject / Hold), station, reason (free text, or
' pick from the list it learns), notes. On submit this module:
'
'   1. writes the serial to WOOD Floor!Serial, and the outcome + date to
'      QC Status / QC Date (the sheet's conditional format greys the row)
'   2. appends one row to QC Log       — every event, every field (the audit)
'   3. appends to the existing tables  — tbl_ReleaseLog (always, with the
'      1-unit split), tbl_RejectLog (reject), tbl_ReworkTracker (rework),
'      the Holds sheet (hold; WOOD Floor!Hold? picks it up by formula)
'   4. learns the reason into QC Reasons so the list grows with use
'
' Columns are found by header text, not position. Dates are written as real
' dates because Summary / Weekly Reporting compare them with SUMPRODUCT.
'=====================================================================

' ---- sheets -----------------------------------------------------------------
Private Const SH_FLOOR As String = "WOOD Floor"
Private Const SH_PLAN_CONFIG As String = "Plan Config"
Private Const SH_HOLDS As String = "Holds"
Private Const SH_QC_LOG As String = "QC Log"
Private Const SH_QC_REASONS As String = "QC Reasons"
Private Const SH_RELEASE_LOG As String = "Release Log"
Private Const SH_REJECT_LOG As String = "Reject Log"
Private Const SH_REWORK As String = "Rework Tracker"

' ---- tables (ListObjects) ---------------------------------------------------
Private Const TBL_RELEASE As String = "tbl_ReleaseLog"
Private Const TBL_REJECT As String = "tbl_RejectLog"
Private Const TBL_REWORK As String = "tbl_ReworkTracker"
Private Const TBL_QC_LOG As String = "tbl_QCLog"
Private Const TBL_QC_REASONS As String = "tbl_QCReasons"

' ---- WOOD Floor headers (row 3) ---------------------------------------------
Private Const FLOOR_HEADER_ROW As Long = 3
Private Const H_WO As String = "WO#"
Private Const H_SKU As String = "SKU"
Private Const H_UNIT As String = "Unit ID"
Private Const H_SERIAL As String = "Serial"
Private Const H_ZONE As String = "Current Zone"
Private Const H_TARGET As String = "Target Date"
Private Const H_RELEASE As String = "Release?"
Private Const H_HOLD As String = "Hold?"
Private Const H_NEXT As String = "Next Station"
Private Const H_QC_STATUS As String = "QC Status"
Private Const H_QC_DATE As String = "QC Date"

' ---- station ids live on Plan Config, first column of the station table -----
Private Const STATION_RANGE As String = "A6:A12"

' ---- outcomes (also the QC Status text that drives the grey-out) -------------
Public Const OUT_RELEASE As String = "RELEASED"
Public Const OUT_BGRADE As String = "B-GRADE"
Public Const OUT_REWORK As String = "REWORK"
Public Const OUT_REJECT As String = "REJECTED"
Public Const OUT_HOLD As String = "ON HOLD"

Private Type FloorCols
    WO As Long
    sku As Long
    Unit As Long
    Serial As Long
    Zone As Long
    Target As Long
    Release As Long
    Hold As Long
    NextStation As Long
    QcStatus As Long
    QcDate As Long
    StationIds() As String      ' WS01..WSnn, from Plan Config
    StationCols() As Long       ' matching tick columns on the floor
End Type

'=====================================================================
' ENTRY POINTS (bind the sheet button to WoodQC_Process)
'=====================================================================
Public Sub WoodQC_Process()
    Dim ws As Worksheet, cols As FloorCols
    Dim r As Long, lastRow As Long, done As Long, skipped As Long, pending As Long
    Dim batch As Boolean

    On Error GoTo FailFast
    Set ws = ThisWorkbook.Worksheets(SH_FLOOR)
    cols = ResolveFloorColumns(ws)

    If ActiveSheet Is ws Then
        r = ActiveCell.Row
        If r > FLOOR_HEADER_ROW And Len(NzStr(ws.Cells(r, cols.Unit).Value)) > 0 Then
            If ProcessUnitEx(ws, cols, r, False) = 1 Then done = 1
            GoTo Report
        End If
    End If

    batch = True
    lastRow = ws.Cells(ws.Rows.Count, cols.Unit).End(xlUp).Row
    For r = FLOOR_HEADER_ROW + 1 To lastRow
        If IsReadyForQC(ws, cols, r) Then pending = pending + 1
    Next r
    If pending = 0 Then
        MsgBox "Nothing to QC." & vbCrLf & vbCrLf & _
               "Either click a cue's row and press the button to QC that one cue, " & _
               "or tick its stations so Release? = TRUE.", vbInformation, "Wood QC"
        Exit Sub
    End If
    If MsgBox(pending & " cue(s) have every station ticked and no QC status yet." & vbCrLf & vbCrLf & _
              "QC them one at a time now?", vbYesNo + vbQuestion, "Wood QC") <> vbYes Then Exit Sub

    For r = FLOOR_HEADER_ROW + 1 To lastRow
        If IsReadyForQC(ws, cols, r) Then
            Select Case ProcessUnitEx(ws, cols, r, True)
                Case 1: done = done + 1
                Case 0: skipped = skipped + 1
                Case -1: GoTo Report
            End Select
        End If
    Next r

Report:
    If done + skipped > 0 Or batch Then
        MsgBox done & " cue(s) processed" & IIf(skipped > 0, ", " & skipped & " skipped", "") & ".", _
               vbInformation, "Wood QC"
    End If
    Exit Sub

FailFast:
    RestoreApp
    MsgBox "WoodQC_Process failed: " & Err.Description, vbExclamation, "Wood QC"
End Sub

' Clears QC Status / QC Date on the selected row(s) so the row un-greys.
' Log rows are deliberately left alone: delete them by hand if the entry was wrong.
Public Sub WoodQC_ClearStatus()
    Dim ws As Worksheet, cols As FloorCols, cell As Range, r As Long, n As Long
    On Error GoTo FailFast
    Set ws = ThisWorkbook.Worksheets(SH_FLOOR)
    If Not ActiveSheet Is ws Or TypeName(Selection) <> "Range" Then
        MsgBox "Select the row(s) on " & SH_FLOOR & " first.", vbInformation, "Wood QC"
        Exit Sub
    End If
    cols = ResolveFloorColumns(ws)
    For Each cell In Selection.Rows
        r = cell.Row
        If r > FLOOR_HEADER_ROW Then
            If Len(NzStr(ws.Cells(r, cols.QcStatus).Value)) > 0 Then n = n + 1
        End If
    Next cell
    If n = 0 Then
        MsgBox "No QC status on the selected row(s).", vbInformation, "Wood QC": Exit Sub
    End If
    If MsgBox("Clear the QC status on " & n & " row(s)?" & vbCrLf & vbCrLf & _
              "The QC Log / Release Log / Reject Log entries stay — remove those by hand if they were wrong.", _
              vbYesNo + vbQuestion, "Wood QC") <> vbYes Then Exit Sub
    For Each cell In Selection.Rows
        r = cell.Row
        If r > FLOOR_HEADER_ROW Then
            ws.Cells(r, cols.QcStatus).ClearContents
            ws.Cells(r, cols.QcDate).ClearContents
        End If
    Next cell
    Exit Sub
FailFast:
    MsgBox "WoodQC_ClearStatus failed: " & Err.Description, vbExclamation, "Wood QC"
End Sub

'=====================================================================
' PER-UNIT FLOW   (1 = written, 0 = skipped / cancelled this cue, -1 = cancel the batch)
'=====================================================================
Private Function ProcessUnitEx(ws As Worksheet, cols As FloorCols, ByVal r As Long, ByVal inBatch As Boolean) As Long
    Dim frm As frmWoodQuality
    Dim unitId As String, sku As String, WO As String, zone As String, nextSt As String
    Dim existing As String, target As Variant

    unitId = NzStr(ws.Cells(r, cols.Unit).Value)
    sku = NzStr(ws.Cells(r, cols.sku).Value)
    WO = NzStr(ws.Cells(r, cols.WO).Value)
    zone = NzStr(ws.Cells(r, cols.Zone).Value)
    nextSt = NzStr(ws.Cells(r, cols.NextStation).Value)
    target = ws.Cells(r, cols.Target).Value
    existing = NzStr(ws.Cells(r, cols.QcStatus).Value)

    If Len(existing) > 0 Then
        If MsgBox(unitId & " already has QC status " & existing & " (" & _
                  Format$(ws.Cells(r, cols.QcDate).Value, "m/d/yyyy") & ")." & vbCrLf & vbCrLf & _
                  "Record another QC event for it?", vbYesNo + vbQuestion, "Wood QC") <> vbYes Then
            ProcessUnitEx = 0: Exit Function
        End If
    End If

    Set frm = New frmWoodQuality
    frm.Prime unitId, sku, WO, target, zone, nextSt, NzStr(ws.Cells(r, cols.Serial).Value), _
              StationList(cols), ReasonList(), inBatch
    frm.Show

    If frm.CancelledAll Then ProcessUnitEx = -1: Unload frm: Exit Function
    If frm.UserCancelled Then ProcessUnitEx = 0: Unload frm: Exit Function

    WriteOutcome ws, cols, r, unitId, sku, WO, target, zone, _
                 frm.Outcome, frm.Serial, frm.Station, frm.Reason, frm.Notes
    Unload frm
    ProcessUnitEx = 1
End Function

Private Sub WriteOutcome(ws As Worksheet, cols As FloorCols, ByVal r As Long, _
                         ByVal unitId As String, ByVal sku As String, ByVal WO As String, _
                         ByVal target As Variant, ByVal zone As String, _
                         ByVal outcome As String, ByVal serial As String, ByVal station As String, _
                         ByVal reason As String, ByVal notes As String)
    Dim today As Date, stamp As Date
    today = Date: stamp = Now

    On Error GoTo FailFast
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' 1. WOOD Floor
    If Len(serial) > 0 Then ws.Cells(r, cols.Serial).Value = serial
    If outcome <> OUT_HOLD Then
        ws.Cells(r, cols.QcStatus).Value = outcome
        ws.Cells(r, cols.QcDate).Value = today
        ws.Cells(r, cols.QcDate).NumberFormat = "m/d/yyyy"
    End If

    ' 2. QC Log
    Dim lo As ListObject
    Set lo = EnsureQcLog()
    AppendRow lo, Array( _
        "Timestamp", stamp, "Date", today, "User", Environ$("USERNAME"), _
        "WO#", WO, "SKU", sku, "Unit ID", unitId, "Serial", serial, _
        "Outcome", outcome, "Station", station, "Reason", reason, "Notes", notes, _
        "Zone at QC", zone, "Target Date", target, "Floor Row", r)

    ' 3. Existing logs the reports already read
    If outcome <> OUT_HOLD Then
        AppendRow GetTable(SH_RELEASE_LOG, TBL_RELEASE), Array( _
            "Date", today, "SKU", sku, "Unit ID", unitId, _
            "Qty Released", IIf(outcome = OUT_RELEASE, 1, 0), "Work Order", WO, _
            "Qty Rework", IIf(outcome = OUT_REWORK, 1, 0), "Qty B Grade", IIf(outcome = OUT_BGRADE, 1, 0), _
            "Qty Reject", IIf(outcome = OUT_REJECT, 1, 0), "Recalled", False, "Serial", serial)
    End If

    Select Case outcome
        Case OUT_REJECT
            AppendRow GetTable(SH_REJECT_LOG, TBL_REJECT), Array( _
                "Date", today, "SKU", sku, "WO#", WO, "Unit ID", unitId, "Quantity", 1, _
                "Description", JoinReason(reason, notes), "Serial", serial, "Station", station)
        Case OUT_REWORK
            AppendRow GetTable(SH_REWORK, TBL_REWORK), ReworkPairs(ws, cols, r, today, WO, sku, unitId, zone, serial)
        Case OUT_HOLD
            AppendHold unitId, station, JoinReason(reason, notes), today
    End Select

    ' 4. Learn the reason
    If Len(reason) > 0 Then LearnReason reason, today

    RestoreApp
    Exit Sub
FailFast:
    Dim errNo As Long, errText As String
    errNo = Err.Number: errText = Err.Description      ' RestoreApp's On Error clears Err
    RestoreApp
    Err.Raise errNo, "WriteOutcome", errText
End Sub

' Header/value pairs for the Rework Tracker row, including every station tick as it stood.
Private Function ReworkPairs(ws As Worksheet, cols As FloorCols, ByVal r As Long, ByVal today As Date, _
                             ByVal WO As String, ByVal sku As String, ByVal unitId As String, _
                             ByVal zone As String, ByVal serial As String) As Variant
    Dim pairs() As Variant, n As Long, i As Long
    n = 16 + 2 * (UBound(cols.StationIds) - LBound(cols.StationIds) + 1)
    ReDim pairs(0 To n - 1)
    pairs(0) = "Date": pairs(1) = today
    pairs(2) = "WO": pairs(3) = WO
    pairs(4) = "SKU": pairs(5) = sku
    pairs(6) = "Unit ID": pairs(7) = unitId
    pairs(8) = "Qty": pairs(9) = 1
    pairs(10) = "Status": pairs(11) = "In Progress"
    pairs(12) = "Previous Release Status": pairs(13) = zone
    pairs(14) = "Serial": pairs(15) = serial
    n = 16
    For i = LBound(cols.StationIds) To UBound(cols.StationIds)
        pairs(n) = cols.StationIds(i)
        pairs(n + 1) = ws.Cells(r, cols.StationCols(i)).Value
        n = n + 2
    Next i
    ReworkPairs = pairs
End Function

'=====================================================================
' WOOD Floor helpers
'=====================================================================
Private Function ResolveFloorColumns(ws As Worksheet) As FloorCols
    Dim c As FloorCols, ids As Variant, i As Long, n As Long
    c.WO = HeaderCol(ws, H_WO)
    c.sku = HeaderCol(ws, H_SKU)
    c.Unit = HeaderCol(ws, H_UNIT)
    c.Serial = HeaderCol(ws, H_SERIAL)
    c.Zone = HeaderCol(ws, H_ZONE)
    c.Target = HeaderCol(ws, H_TARGET)
    c.Release = HeaderCol(ws, H_RELEASE)
    c.Hold = HeaderCol(ws, H_HOLD)
    c.NextStation = HeaderCol(ws, H_NEXT)
    c.QcStatus = HeaderCol(ws, H_QC_STATUS, True)
    c.QcDate = HeaderCol(ws, H_QC_DATE, True)

    ids = PlanConfigStations()
    n = UBound(ids) - LBound(ids) + 1
    ReDim c.StationIds(0 To n - 1)
    ReDim c.StationCols(0 To n - 1)
    For i = 0 To n - 1
        c.StationIds(i) = CStr(ids(LBound(ids) + i))
        c.StationCols(i) = HeaderCol(ws, c.StationIds(i))
    Next i
    ResolveFloorColumns = c
End Function

' Station ids from Plan Config (A6:A12), blanks skipped. Raises if none.
Private Function PlanConfigStations() As Variant
    Dim ws As Worksheet, cell As Range, out As Collection, v As Variant, i As Long
    Set ws = ThisWorkbook.Worksheets(SH_PLAN_CONFIG)
    Set out = New Collection
    For Each cell In ws.Range(STATION_RANGE).Cells
        If Len(NzStr(cell.Value)) > 0 Then out.Add NzStr(cell.Value)
    Next cell
    If out.Count = 0 Then
        Err.Raise vbObjectError + 515, "PlanConfigStations", _
                  "No station ids in " & SH_PLAN_CONFIG & "!" & STATION_RANGE & "."
    End If
    ReDim v(0 To out.Count - 1)
    For i = 1 To out.Count: v(i - 1) = out(i): Next i
    PlanConfigStations = v
End Function

' Column index of a header on WOOD Floor row 3. Optional headers are created
' in the first free column so an older copy of the sheet still works.
Private Function HeaderCol(ws As Worksheet, ByVal header As String, Optional ByVal createIfMissing As Boolean = False) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(FLOOR_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim$(NzStr(ws.Cells(FLOOR_HEADER_ROW, c).Value)), header, vbTextCompare) = 0 Then
            HeaderCol = c: Exit Function
        End If
    Next c
    If Not createIfMissing Then
        Err.Raise vbObjectError + 513, "HeaderCol", _
                  "Header '" & header & "' not found on " & ws.Name & " row " & FLOOR_HEADER_ROW & "."
    End If
    ws.Cells(FLOOR_HEADER_ROW, lastCol + 1).Value = header
    HeaderCol = lastCol + 1
End Function

Private Function IsReadyForQC(ws As Worksheet, cols As FloorCols, ByVal r As Long) As Boolean
    If Len(NzStr(ws.Cells(r, cols.Unit).Value)) = 0 Then Exit Function
    If Len(NzStr(ws.Cells(r, cols.QcStatus).Value)) > 0 Then Exit Function
    IsReadyForQC = Boolish(ws.Cells(r, cols.Release).Value)
End Function

Private Function StationList(cols As FloorCols) As Variant
    Dim v() As String, i As Long
    ReDim v(LBound(cols.StationIds) To UBound(cols.StationIds))
    For i = LBound(cols.StationIds) To UBound(cols.StationIds)
        v(i) = cols.StationIds(i)
    Next i
    StationList = v
End Function

'=====================================================================
' QC Log / QC Reasons (created on first use if missing)
'=====================================================================
Private Function EnsureQcLog() As ListObject
    Set EnsureQcLog = EnsureTable(SH_QC_LOG, TBL_QC_LOG, Array( _
        "Timestamp", "Date", "User", "WO#", "SKU", "Unit ID", "Serial", "Outcome", _
        "Station", "Reason", "Notes", "Zone at QC", "Target Date", "Floor Row"), SH_REJECT_LOG)
End Function

Private Function EnsureQcReasons() As ListObject
    Set EnsureQcReasons = EnsureTable(SH_QC_REASONS, TBL_QC_REASONS, Array( _
        "Reason", "Times Used", "First Seen", "Last Used", "Active", "Category (later)"), SH_QC_LOG)
End Function

Private Function ReasonList() As Variant
    Dim lo As ListObject, i As Long, n As Long, v() As String, active As Variant
    Set lo = EnsureQcReasons()
    If lo.ListRows.Count = 0 Then ReasonList = Array(): Exit Function
    ReDim v(0 To lo.ListRows.Count - 1)
    For i = 1 To lo.ListRows.Count
        active = lo.ListRows(i).Range.Cells(1, lo.ListColumns("Active").Index).Value
        If Len(NzStr(lo.ListRows(i).Range.Cells(1, 1).Value)) > 0 And (IsEmpty(active) Or Boolish(active)) Then
            v(n) = NzStr(lo.ListRows(i).Range.Cells(1, 1).Value): n = n + 1
        End If
    Next i
    If n = 0 Then ReasonList = Array(): Exit Function
    ReDim Preserve v(0 To n - 1)
    ReasonList = v
End Function

Private Sub LearnReason(ByVal reason As String, ByVal today As Date)
    Dim lo As ListObject, i As Long, cUsed As Long, cLast As Long, cFirst As Long
    Set lo = EnsureQcReasons()
    cUsed = lo.ListColumns("Times Used").Index
    cLast = lo.ListColumns("Last Used").Index
    cFirst = lo.ListColumns("First Seen").Index
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(NzStr(lo.ListRows(i).Range.Cells(1, 1).Value)), reason, vbTextCompare) = 0 Then
            With lo.ListRows(i).Range
                .Cells(1, cUsed).Value = Val(NzStr(.Cells(1, cUsed).Value)) + 1
                .Cells(1, cLast).Value = today
                If IsEmpty(.Cells(1, cFirst).Value) Then .Cells(1, cFirst).Value = today
            End With
            Exit Sub
        End If
    Next i
    AppendRow lo, Array("Reason", reason, "Times Used", 1, "First Seen", today, "Last Used", today, "Active", True)
End Sub

'=====================================================================
' Holds sheet (plain range, read by WOOD Floor!Hold? formulas)
'=====================================================================
Private Sub AppendHold(ByVal unitId As String, ByVal station As String, ByVal reason As String, ByVal today As Date)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_HOLDS)
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = unitId
    ws.Cells(r, 2).Value = station
    ws.Cells(r, 3).Value = reason
    ws.Cells(r, 4).Value = today
    ws.Cells(r, 5).Value = False
    ws.Cells(r, 6).Value = "HELD"
End Sub

'=====================================================================
' ListObject helpers
'=====================================================================
Private Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set GetTable = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    On Error GoTo 0
    If GetTable Is Nothing Then
        Err.Raise vbObjectError + 514, "GetTable", "Table " & tableName & " not found on sheet " & sheetName & "."
    End If
End Function

Private Function EnsureTable(ByVal sheetName As String, ByVal tableName As String, ByVal headers As Variant, _
                             ByVal placeAfter As String) As ListObject
    Dim ws As Worksheet, lo As ListObject, i As Long, anchor As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        On Error Resume Next
        Set anchor = ThisWorkbook.Worksheets(placeAfter)
        On Error GoTo 0
        If anchor Is Nothing Then Set anchor = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set ws = ThisWorkbook.Worksheets.Add(After:=anchor)
        ws.Name = sheetName
    End If
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i - LBound(headers) + 1).Value = headers(i)
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) - LBound(headers) + 1)), , xlYes)
        lo.Name = tableName
        lo.TableStyle = "TableStyleMedium2"
        ws.Rows(1).Font.Bold = True
        ws.Columns.AutoFit
    End If
    Set EnsureTable = lo
End Function

' pairs = Array("Header", value, ...). Headers absent from the table are skipped.
' A table whose only row is empty (fresh template) gets that row, not a new one.
Private Sub AppendRow(lo As ListObject, ByVal pairs As Variant)
    Dim lr As ListRow, i As Long, col As ListColumn, header As String
    If lo.ListRows.Count = 1 And IsEmpty(lo.ListRows(1).Range.Cells(1, 1).Value) Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    For i = LBound(pairs) To UBound(pairs) - 1 Step 2
        header = CStr(pairs(i))
        Set col = Nothing
        On Error Resume Next
        Set col = lo.ListColumns(header)
        On Error GoTo 0
        If Not col Is Nothing Then
            lr.Range.Cells(1, col.Index).Value = pairs(i + 1)
            If VarType(pairs(i + 1)) = vbDate Then
                lr.Range.Cells(1, col.Index).NumberFormat = IIf(header = "Timestamp", "m/d/yyyy h:mm", "m/d/yyyy")
            End If
        End If
    Next i
End Sub

'=====================================================================
' small utilities
'=====================================================================
Private Function JoinReason(ByVal reason As String, ByVal notes As String) As String
    If Len(notes) = 0 Then
        JoinReason = reason
    ElseIf Len(reason) = 0 Then
        JoinReason = notes
    Else
        JoinReason = reason & " — " & notes
    End If
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = Trim$(CStr(v))
End Function

Private Function Boolish(ByVal v As Variant) As Boolean
    On Error Resume Next
    Select Case True
        Case IsError(v), IsEmpty(v), IsNull(v): Boolish = False
        Case VarType(v) = vbBoolean: Boolish = CBool(v)
        Case IsNumeric(v): Boolish = (CDbl(v) <> 0)
        Case Else: Boolish = (UCase$(Trim$(CStr(v))) = "TRUE" Or Trim$(CStr(v)) = "1" Or UCase$(Trim$(CStr(v))) = "X")
    End Select
End Function

Private Sub RestoreApp()
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
