Attribute VB_Name = "modTableQuality"
Option Explicit

'=====================================================================
' modTableQuality — QC / release flow for TABLE Floor
'
' One row on TABLE Floor = one serialised table. The "QC / RELEASE TABLE"
' button runs TableQC_Process:
'
'   * a data row is selected            -> QC that one table
'   * anything else is selected         -> QC every table whose Release? = TRUE
'                                          and whose QC Status is still blank
'
' For each table the frmTableQuality dialog captures: serial number, outcome
' (Release / B-Grade / Rework / Reject / Hold), station, reason (free text, or
' pick from the list it learns), notes. On submit this module:
'
'   1. writes the serial to TABLE Floor!Serial, and the outcome + date to
'      QC Status / QC Date (the sheet's conditional format greys the row)
'   2. appends one row to QC Log       — every event, every field (the audit)
'   3. appends to the existing tables  — tbl_ReleaseLog (always, with the
'      1-unit split), tbl_RejectLog (reject), tbl_ReworkTracker (rework),
'      the Holds sheet (hold; TABLE Floor!Hold? picks it up by formula)
'   4. learns the reason into QC Reasons so the list grows with use
'
' Columns are found by header text, not position: insert a column on TABLE
' Floor and nothing here breaks. Dates are written as real dates because the
' Summary / Weekly Reporting SUMPRODUCTs compare tbl_ReleaseLog[Date] to dates.
'=====================================================================

' ---- sheets -----------------------------------------------------------------
Private Const SH_FLOOR As String = "TABLE Floor"
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

' ---- TABLE Floor headers (row 3) --------------------------------------------
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

' ---- outcomes (also the QC Status text that drives the grey-out) -------------
Public Const OUT_RELEASE As String = "RELEASED"
Public Const OUT_BGRADE As String = "B-GRADE"
Public Const OUT_REWORK As String = "REWORK"
Public Const OUT_REJECT As String = "REJECTED"
Public Const OUT_HOLD As String = "ON HOLD"

' Stations: Plan Config!A6:A10 (ST01..ST05). PreAssym is the floor's own label.
Private Const STATION_PREASM As String = "PreAssym"

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
    ST(1 To 5) As Long    ' station tick columns ST01..ST05
End Type

'=====================================================================
' ENTRY POINTS (bound to the sheet buttons)
'=====================================================================
Public Sub TableQC_Process()
    Dim ws As Worksheet, cols As FloorCols
    Dim r As Long, lastRow As Long, done As Long, skipped As Long, pending As Long
    Dim batch As Boolean

    On Error GoTo FailFast
    Set ws = ThisWorkbook.Worksheets(SH_FLOOR)
    cols = ResolveFloorColumns(ws)

    ' Single-row mode: the user clicked on a table's row.
    If ActiveSheet Is ws Then
        r = ActiveCell.Row
        If r > FLOOR_HEADER_ROW And Len(NzStr(ws.Cells(r, cols.Unit).Value)) > 0 Then
            If ProcessUnit(ws, cols, r, False) Then done = 1
            GoTo Report
        End If
    End If

    ' Batch mode: every table flagged Release? and not yet QC'd.
    batch = True
    lastRow = ws.Cells(ws.Rows.Count, cols.Unit).End(xlUp).Row
    For r = FLOOR_HEADER_ROW + 1 To lastRow
        If IsReadyForQC(ws, cols, r) Then pending = pending + 1
    Next r
    If pending = 0 Then
        MsgBox "Nothing to QC." & vbCrLf & vbCrLf & _
               "Either click a table's row and press the button to QC that one table, " & _
               "or tick its stations so Release? = TRUE.", vbInformation, "Table QC"
        Exit Sub
    End If
    If MsgBox(pending & " table(s) have every station ticked and no QC status yet." & vbCrLf & vbCrLf & _
              "QC them one at a time now?", vbYesNo + vbQuestion, "Table QC") <> vbYes Then Exit Sub

    For r = FLOOR_HEADER_ROW + 1 To lastRow
        If IsReadyForQC(ws, cols, r) Then
            Select Case ProcessUnitEx(ws, cols, r, True)
                Case 1: done = done + 1
                Case 0: skipped = skipped + 1
                Case -1: GoTo Report              ' cancelled the whole batch
            End Select
        End If
    Next r

Report:
    If done + skipped > 0 Or batch Then
        MsgBox done & " table(s) processed" & IIf(skipped > 0, ", " & skipped & " skipped", "") & ".", _
               vbInformation, "Table QC"
    End If
    Exit Sub

FailFast:
    RestoreApp
    MsgBox "TableQC_Process failed: " & Err.Description, vbExclamation, "Table QC"
End Sub

' Clears QC Status / QC Date on the selected row(s) so the row un-greys.
' Log rows are deliberately left alone: delete them by hand if the entry was wrong.
Public Sub TableQC_ClearStatus()
    Dim ws As Worksheet, cols As FloorCols, cell As Range, r As Long, n As Long
    On Error GoTo FailFast
    Set ws = ThisWorkbook.Worksheets(SH_FLOOR)
    If Not ActiveSheet Is ws Or TypeName(Selection) <> "Range" Then
        MsgBox "Select the row(s) on " & SH_FLOOR & " first.", vbInformation, "Table QC"
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
        MsgBox "No QC status on the selected row(s).", vbInformation, "Table QC": Exit Sub
    End If
    If MsgBox("Clear the QC status on " & n & " row(s)?" & vbCrLf & vbCrLf & _
              "The QC Log / Release Log / Reject Log entries stay — remove those by hand if they were wrong.", _
              vbYesNo + vbQuestion, "Table QC") <> vbYes Then Exit Sub
    For Each cell In Selection.Rows
        r = cell.Row
        If r > FLOOR_HEADER_ROW Then
            ws.Cells(r, cols.QcStatus).ClearContents
            ws.Cells(r, cols.QcDate).ClearContents
        End If
    Next cell
    Exit Sub
FailFast:
    MsgBox "TableQC_ClearStatus failed: " & Err.Description, vbExclamation, "Table QC"
End Sub

'=====================================================================
' PER-UNIT FLOW
'=====================================================================
Private Function ProcessUnit(ws As Worksheet, cols As FloorCols, ByVal r As Long, ByVal inBatch As Boolean) As Boolean
    ProcessUnit = (ProcessUnitEx(ws, cols, r, inBatch) = 1)
End Function

' Returns 1 = written, 0 = skipped/cancelled this unit, -1 = cancel the batch.
Private Function ProcessUnitEx(ws As Worksheet, cols As FloorCols, ByVal r As Long, ByVal inBatch As Boolean) As Long
    Dim frm As frmTableQuality
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
                  "Record another QC event for it?", vbYesNo + vbQuestion, "Table QC") <> vbYes Then
            ProcessUnitEx = 0: Exit Function
        End If
    End If

    Set frm = New frmTableQuality
    frm.Prime unitId, sku, WO, target, zone, nextSt, NzStr(ws.Cells(r, cols.Serial).Value), _
              StationList(), ReasonList(), inBatch
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

    ' 1. TABLE Floor
    If Len(serial) > 0 Then ws.Cells(r, cols.Serial).Value = serial
    If outcome <> OUT_HOLD Then
        ws.Cells(r, cols.QcStatus).Value = outcome          ' conditional format greys the row
        ws.Cells(r, cols.QcDate).Value = today
        ws.Cells(r, cols.QcDate).NumberFormat = "m/d/yyyy"
    End If

    ' 2. QC Log — the audit trail, one row per event
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
            AppendRow GetTable(SH_REWORK, TBL_REWORK), Array( _
                "Date", today, "WO", WO, "SKU", sku, "Unit ID", unitId, "Qty", 1, _
                "Status", "In Progress", "Previous Release Status", zone, _
                "ST01", ws.Cells(r, cols.ST(1)).Value, "ST02", ws.Cells(r, cols.ST(2)).Value, _
                "ST03", ws.Cells(r, cols.ST(3)).Value, "ST04", ws.Cells(r, cols.ST(4)).Value, _
                "ST05", ws.Cells(r, cols.ST(5)).Value, "Serial", serial)
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

'=====================================================================
' TABLE Floor helpers
'=====================================================================
Private Function ResolveFloorColumns(ws As Worksheet) As FloorCols
    Dim c As FloorCols
    c.WO = HeaderCol(ws, H_WO)
    c.sku = HeaderCol(ws, H_SKU)
    c.Unit = HeaderCol(ws, H_UNIT)
    c.Serial = HeaderCol(ws, H_SERIAL)
    c.Zone = HeaderCol(ws, H_ZONE)
    c.Target = HeaderCol(ws, H_TARGET)
    c.Release = HeaderCol(ws, H_RELEASE)
    c.Hold = HeaderCol(ws, H_HOLD)
    c.NextStation = HeaderCol(ws, H_NEXT)
    Dim k As Long
    For k = 1 To 5
        c.ST(k) = HeaderCol(ws, "ST0" & k)
    Next k
    c.QcStatus = HeaderCol(ws, H_QC_STATUS, True)
    c.QcDate = HeaderCol(ws, H_QC_DATE, True)
    ResolveFloorColumns = c
End Function

' Column index of a header on TABLE Floor row 3. Optional headers are created
' in the first free column so an un-patched copy of the sheet still works.
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

Private Function StationList() As Variant
    Dim ws As Worksheet, rng As Range, cell As Range, out As Collection, v As Variant, i As Long
    Set out = New Collection
    out.Add STATION_PREASM
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_PLAN_CONFIG)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set rng = ws.Range("A6:A10")                         ' station ids, matches the floor formulas
        For Each cell In rng.Cells
            If Len(NzStr(cell.Value)) > 0 Then out.Add NzStr(cell.Value)
        Next cell
    End If
    ReDim v(0 To out.Count - 1)
    For i = 1 To out.Count: v(i - 1) = out(i): Next i
    StationList = v
End Function

'=====================================================================
' QC Log / QC Reasons (created on first use)
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
    Dim lo As ListObject, i As Long, cUsed As Long, cLast As Long
    Set lo = EnsureQcReasons()
    cUsed = lo.ListColumns("Times Used").Index
    cLast = lo.ListColumns("Last Used").Index
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(NzStr(lo.ListRows(i).Range.Cells(1, 1).Value)), reason, vbTextCompare) = 0 Then
            With lo.ListRows(i).Range
                .Cells(1, cUsed).Value = Val(NzStr(.Cells(1, cUsed).Value)) + 1
                .Cells(1, cLast).Value = today
            End With
            Exit Sub
        End If
    Next i
    AppendRow lo, Array("Reason", reason, "Times Used", 1, "First Seen", today, "Last Used", today, "Active", True)
End Sub

'=====================================================================
' Holds sheet (plain range, read by TABLE Floor!Hold? formulas)
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
    ws.Cells(r, 5).Value = False                    ' Released? — set TRUE to return it to the floor
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

' pairs = Array("Header", value, "Header", value, ...). Headers absent from the
' table are skipped silently so a table someone reshaped still gets its rows.
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
