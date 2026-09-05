Attribute VB_Name = "modMigrate"
Option Explicit

'======================================================================
' modMigrate  -  one-time lean-out, run INSIDE the workbook
'
' Run this on a COPY. It is destructive by design. The PowerShell driver
' makes the copy for you; if you are running it by hand, make one first.
'
' Why a macro instead of a rebuilt file: your retained sheets carry
' drawings AND legacy VML form controls (the macro buttons on REVO Floor,
' Command Center, Plan Config, Holds, Remaining by Shift and others).
' Any external tool that rewrites this workbook drops them. Excel deletes
' a sheet and repairs the VBA document-module binding in the same
' operation. Nothing else does both.
'
' Order of operations:
'   1. SnapshotShapes       - records every shape on every retained sheet
'   2. LeanOutWorkbook      - deletes the 36 sheets
'   3. DeleteCutQueries     - removes the 8 orphaned Power Query
'                             connections and their query definitions
'   4. TruncateReleaseLog   - drops ~1,045,800 junk rows
'   5. ResetAllUsedRanges   - recovers formatted-to-the-bottom bloat
'   6. AuditRefErrors       - proves no retained formula broke
'   7. VerifyShapes         - proves every macro button survived
'
' MigrateAll runs the lot and writes REVO_Migration_Log.txt beside the
' workbook. That log is the acceptance evidence.
'======================================================================

Private Const CUT_LIST As String = _
    "REVO Component Tracker|REVO Component Forecast|REVO Component Inventory|" & _
    "REVO Component_Model|REVO Subcomponent_Model|REVO SubComp_Model|REVO PreAsm_Model|" & _
    "REVO SubComp Dashboard|REVO PreAsm Dashboard|REVO SKU Detail Dashboard|" & _
    "Bin_Inv|REVO_INV|EST_REVO_INV|STOCK ROOM COUNT|On_PO|Shipments|Lead_Time_By_Item|" & _
    "BOM|BOM Hierarchy View|Butt Classification|" & _
    "Tube Schedule|Tube_Model|Tube_Dashboard|Tube SKU Detail Dashboard|" & _
    "Boxed_By_Component|Box_Utilization|Box_Plan|Box_Plan_Archive|" & _
    "Production Table|Production Table Backup|RCF_Step5_VirtualBuilds|" & _
    "HIEProdFlat|HIEInv|Launch Import|Shaft Balance|not used "

' Deliberately NOT in the list - these are demand forecast, not inventory:
'   REVO_SALES_FRCST
'   Upside Production Forecast

' Power Query connections whose load target is on the cut list. Verified
' against xl/connections.xml. "Query - REVO Sales Forecast" is NOT here:
' it lands on REVO_SALES_FRCST, which is retained.
Private Const CONN_CUT As String = _
    "Query - Bin_Inv|Query - BOM|Query - HIEInv|Query - HIEProdFlat|" & _
    "Query - Lead_Time_By_Item|Query - On_PO|Query - REVO INVENTORY|" & _
    "Query - STOCK ROOM COUNT"

' The underlying WorkbookQuery names, minus the "Query - " prefix.
Private Const QUERY_CUT As String = _
    "Bin_Inv|BOM|HIEInv|HIEProdFlat|Lead_Time_By_Item|On_PO|" & _
    "REVO INVENTORY|STOCK ROOM COUNT"

Private Const CONN_PROTECT As String = _
    "Query - REVO Sales Forecast|ThisWorkbookDataModel"

Private mShapeBefore As Object          ' "Sheet|ShapeName" -> True

'======================================================================

Public Sub MigrateAll()

    Dim t As Double

    If Not Confirm("This permanently deletes 36 sheets, 8 Power Query connections" & vbCrLf & _
                   "and truncates Release Log." & vbCrLf & vbCrLf & _
                   "Are you running this on a COPY?", "REVO Migration") Then Exit Sub

    On Error GoTo Fail
    t = Timer
    LogLine String$(70, "=")
    LogLine "MigrateAll starting. modConfig v" & APP_VERSION
    LogLine "Workbook: " & ThisWorkbook.FullName
    ReportWorkbookSize

    FastOn
    SnapshotShapes
    LeanOutWorkbook
    DeleteCutQueries
    TruncateReleaseLog
    ResetAllUsedRanges
    CleanBrokenNames
    FastOff

    Application.CalculateFullRebuild
    AuditRefErrors
    VerifyShapes

    ThisWorkbook.Save
    ReportWorkbookSize
    LogLine "MigrateAll finished in " & Format$(Timer - t, "0.0") & "s"
    Say "Migration complete in " & Format$(Timer - t, "0.0") & "s." & vbCrLf & vbCrLf & _
        "Sheets: " & ThisWorkbook.Worksheets.Count & vbCrLf & _
        "See " & LOG_FILE & " beside the workbook for the full audit trail.", "REVO Migration"
    Exit Sub

Fail:
    FastReset
    LogLine "*** MigrateAll FAILED: " & Err.Number & " " & Err.Description
    Say "Migration failed: " & Err.Description & vbCrLf & vbCrLf & _
        "Restore from the phase backup. Nothing was saved.", "REVO Migration", vbCritical
End Sub

'----------------------------------------------------------------------
' 1. Record every shape on every sheet that will survive, so we can
'    prove afterwards that no macro button was lost. This is the
'    acceptance criterion the brief cares most about.
'----------------------------------------------------------------------
Public Sub SnapshotShapes()

    Dim ws As Worksheet, sh As Shape
    Dim n As Long

    Set mShapeBefore = CreateObject("Scripting.Dictionary")
    For Each ws In ThisWorkbook.Worksheets
        If Not IsCutSheet(ws.Name) Then
            On Error Resume Next
            For Each sh In ws.Shapes
                mShapeBefore(ws.Name & "|" & sh.Name) = True
                n = n + 1
            Next sh
            On Error GoTo 0
        End If
    Next ws
    LogLine "SnapshotShapes: " & n & " shapes recorded across retained sheets"

End Sub

Public Sub VerifyShapes()

    Dim ws As Worksheet, sh As Shape
    Dim have As Object, k As Variant
    Dim lost As String, n As Long, missing As Long

    If mShapeBefore Is Nothing Then
        LogLine "VerifyShapes: no snapshot taken, skipping"
        Exit Sub
    End If

    Set have = CreateObject("Scripting.Dictionary")
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        For Each sh In ws.Shapes
            have(ws.Name & "|" & sh.Name) = True
            n = n + 1
        Next sh
        On Error GoTo 0
    Next ws

    For Each k In mShapeBefore.keys
        If Not have.Exists(k) Then
            missing = missing + 1
            If missing <= 25 Then lost = lost & vbCrLf & "  " & k
        End If
    Next k

    If missing = 0 Then
        LogLine "VerifyShapes: PASS - all " & mShapeBefore.Count & " shapes intact (" & n & " now present)"
    Else
        LogLine "*** VerifyShapes: FAIL - " & missing & " shapes lost:" & Replace(lost, vbCrLf, " / ")
        Say missing & " shape(s) were lost during migration:" & lost & vbCrLf & vbCrLf & _
            "Restore from backup and investigate before saving.", "REVO Migration", vbCritical
    End If

End Sub

'----------------------------------------------------------------------
' 2. Delete the cut sheets
'----------------------------------------------------------------------
Public Sub LeanOutWorkbook()

    Dim names As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim gone As Long, missing As String

    names = Split(CUT_LIST, "|")
    FastOn

    For i = LBound(names) To UBound(names)
        Set ws = GetSheet(CStr(names(i)))
        If ws Is Nothing Then
            missing = missing & vbCrLf & "  " & names(i)
        Else
            On Error Resume Next
            ws.Delete
            If Err.Number = 0 Then
                gone = gone + 1
            Else
                LogLine "  could not delete '" & names(i) & "': " & Err.Description
            End If
            Err.Clear
            On Error GoTo 0
        End If
        Set ws = Nothing
    Next i

    FastOff

    LogLine "LeanOutWorkbook: deleted " & gone & " of " & (UBound(names) - LBound(names) + 1) & _
            " sheets; " & ThisWorkbook.Worksheets.Count & " remain" & _
            IIf(Len(missing) > 0, "; not found:" & Replace(missing, vbCrLf, " "), "")

    Say "Deleted " & gone & " sheets. " & ThisWorkbook.Worksheets.Count & " remain." & _
        IIf(Len(missing) > 0, vbCrLf & vbCrLf & "Not found (already gone or renamed):" & missing, ""), _
        "REVO Migration"

End Sub

Private Function IsCutSheet(ByVal nm As String) As Boolean
    IsCutSheet = (InStr(1, "|" & CUT_LIST & "|", "|" & nm & "|", vbTextCompare) > 0)
End Function

'----------------------------------------------------------------------
' 3. The 8 Power Query connections that loaded into deleted sheets.
'    Deleting the sheet leaves the connection alive but landing nowhere.
'    Chase's call: remove them outright.
'
'    Two objects to kill per query: the WorkbookConnection (the refresh
'    plumbing) and the WorkbookQuery (the M definition in the Queries
'    pane). Deleting only the connection leaves the query behind.
'----------------------------------------------------------------------
Public Sub DeleteCutQueries()

    Dim wantConn As Variant, wantQry As Variant
    Dim i As Long, killedC As Long, killedQ As Long
    Dim c As Object, q As Object
    Dim nm As String

    wantConn = Split(CONN_CUT, "|")
    wantQry = Split(QUERY_CUT, "|")

    ' --- connections -------------------------------------------------
    For i = LBound(wantConn) To UBound(wantConn)
        nm = CStr(wantConn(i))
        If InStr(1, "|" & CONN_PROTECT & "|", "|" & nm & "|", vbTextCompare) > 0 Then
            LogLine "  PROTECTED, not deleting connection: " & nm
        Else
            On Error Resume Next
            Set c = Nothing
            Set c = ThisWorkbook.Connections(nm)
            If Not c Is Nothing Then
                c.Delete
                If Err.Number = 0 Then
                    killedC = killedC + 1
                Else
                    LogLine "  connection delete failed '" & nm & "': " & Err.Description
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next i

    ' --- query definitions -------------------------------------------
    For i = LBound(wantQry) To UBound(wantQry)
        nm = CStr(wantQry(i))
        On Error Resume Next
        Set q = Nothing
        Set q = ThisWorkbook.Queries(nm)
        If Not q Is Nothing Then
            q.Delete
            If Err.Number = 0 Then killedQ = killedQ + 1
        End If
        Err.Clear
        On Error GoTo 0
    Next i

    LogLine "DeleteCutQueries: removed " & killedC & " connections and " & killedQ & " query definitions"
    LogLine "  remaining connections: " & ConnectionNames()

    Say "Removed " & killedC & " orphaned connections and " & killedQ & " query definitions." & vbCrLf & vbCrLf & _
        "Still present: " & ConnectionNames(), "REVO Migration"

End Sub

Private Function ConnectionNames() As String
    Dim c As Object, s As String
    On Error Resume Next
    For Each c In ThisWorkbook.Connections
        s = s & IIf(Len(s) > 0, ", ", "") & c.Name
    Next c
    On Error GoTo 0
    If Len(s) = 0 Then s = "(none)"
    ConnectionNames = s
End Function

'----------------------------------------------------------------------
' 4. Release Log.
'
'    CORRECTION to the brief: this sheet holds exactly ONE formula, not
'    1,048,552 of them. The bloat is 1,048,552 row elements each carrying
'    a single FALSE in column I ("Recalled") - a checkbox column dragged
'    to the bottom of the grid. Real data ends at row 765. There is no
'    recalc problem here, only a size problem, and truncation fixes it.
'----------------------------------------------------------------------
Public Sub TruncateReleaseLog()

    Dim ws As Worksheet
    Dim src As Variant
    Dim i As Long, lastReal As Long, firstCut As Long
    Const BUFFER As Long = 2000

    Set ws = GetSheet(SH_RELEASELOG)
    If ws Is Nothing Then
        LogLine "TruncateReleaseLog: " & SH_RELEASELOG & " not found, skipped"
        Say SH_RELEASELOG & " not found.", "REVO Migration", vbExclamation
        Exit Sub
    End If

    FastOn

    ' Fast path: End(xlUp) from the bottom of column A.
    lastReal = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastReal > 1 Then
        If Not IsDate(ws.Cells(lastReal, "A").Value) Then lastReal = 0
    Else
        lastReal = 0
    End If

    ' Slow path: only if the fast path landed somewhere that is not a date.
    If lastReal = 0 Then
        src = ws.Range("A1:A" & ws.Rows.Count).Value
        For i = UBound(src, 1) To 1 Step -1
            If IsDate(src(i, 1)) Then
                lastReal = i
                Exit For
            End If
        Next i
        Erase src
    End If

    If lastReal = 0 Then
        FastOff
        LogLine "*** TruncateReleaseLog: no dated rows found, nothing truncated"
        Say "No dated rows found in " & SH_RELEASELOG & ". Nothing truncated.", _
            "REVO Migration", vbExclamation
        Exit Sub
    End If

    firstCut = lastReal + BUFFER + 1
    If firstCut < ws.Rows.Count Then
        ws.Range(ws.Rows(firstCut), ws.Rows(ws.Rows.Count)).Delete
    End If

    FastOff

    LogLine "TruncateReleaseLog: last dated row " & lastReal & ", buffer " & BUFFER & _
            ", deleted rows " & firstCut & "+"

    Say SH_RELEASELOG & " truncated." & vbCrLf & vbCrLf & _
        "Last real dated row: " & lastReal & vbCrLf & _
        "Formula buffer kept: " & BUFFER & " rows" & vbCrLf & vbCrLf & _
        "Convert this range to a Table (Ctrl+T) so it grows on its own " & _
        "instead of being filled to the bottom of the grid again.", "REVO Migration"

End Sub

'----------------------------------------------------------------------
' 5. Reset used range on every remaining sheet.
'
'    The original version deleted every row below the last used cell.
'    On REVO Floor, Command Center, Plan Config and Rework Tracker the
'    macro buttons are anchored BELOW and to the RIGHT of the data, so
'    that would have deleted the buttons - the exact failure the brief
'    forbids. ShapeFloor() computes the lowest row and rightmost column
'    occupied by any shape and refuses to cut above it.
'----------------------------------------------------------------------
Public Sub ResetAllUsedRanges()

    Dim ws As Worksheet
    Dim lastCell As Range
    Dim lastRow As Long, lastCol As Long
    Dim shRow As Long, shCol As Long
    Dim touched As Long
    Dim dummy As Long

    FastOn

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        lastRow = 0: lastCol = 0

        Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                     LookIn:=xlFormulas, LookAt:=xlPart, _
                                     SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If Not lastCell Is Nothing Then lastRow = lastCell.Row

        Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                     LookIn:=xlFormulas, LookAt:=xlPart, _
                                     SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        If Not lastCell Is Nothing Then lastCol = lastCell.Column

        ShapeFloor ws, shRow, shCol
        If shRow > lastRow Then lastRow = shRow
        If shCol > lastCol Then lastCol = shCol

        If lastRow > 0 And lastRow < ws.Rows.Count Then
            ws.Range(ws.Rows(lastRow + 1), ws.Rows(ws.Rows.Count)).Delete
            touched = touched + 1
        End If
        If lastCol > 0 And lastCol < ws.Columns.Count Then
            ws.Range(ws.Columns(lastCol + 1), ws.Columns(ws.Columns.Count)).Delete
        End If

        dummy = ws.UsedRange.Rows.Count      ' forces Excel to rebuild the used range
        Err.Clear
        On Error GoTo 0
        Set lastCell = Nothing
    Next ws

    FastOff

    LogLine "ResetAllUsedRanges: used range reset on " & touched & " sheets (shape-protected)"
    Say "Used range reset on " & touched & " sheets." & vbCrLf & vbCrLf & _
        "Save and reopen to see the file size drop.", "REVO Migration"

End Sub

' Lowest row / rightmost column occupied by any shape on the sheet.
' Shapes that are not cell-anchored are ignored rather than fatal.
Private Sub ShapeFloor(ByVal ws As Worksheet, ByRef outRow As Long, ByRef outCol As Long)
    Dim sh As Shape, r As Long, c As Long
    outRow = 0: outCol = 0
    On Error Resume Next
    For Each sh In ws.Shapes
        r = 0: c = 0
        r = sh.BottomRightCell.Row
        c = sh.BottomRightCell.Column
        If r > outRow Then outRow = r
        If c > outCol Then outCol = c
        Err.Clear
    Next sh
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------
' 6. Deleting sheets turns any workbook-scoped defined name that pointed
'    at them into #REF!. Nineteen names referenced cut sheets, mostly
'    _FilterDatabase and the ExternalData_n query landing ranges. Sheet-
'    scoped ones go with the sheet; the rest are swept here.
'----------------------------------------------------------------------
Public Sub CleanBrokenNames()

    Dim nm As Name
    Dim i As Long, killed As Long
    Dim s As String

    On Error Resume Next
    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        s = ""
        s = nm.RefersTo
        If InStr(1, s, "#REF!", vbTextCompare) > 0 Then
            LogLine "  deleting broken name: " & nm.Name & " = " & s
            nm.Delete
            If Err.Number = 0 Then killed = killed + 1
            Err.Clear
        End If
    Next i
    On Error GoTo 0

    LogLine "CleanBrokenNames: removed " & killed & " #REF! names; " & _
            ThisWorkbook.Names.Count & " remain"

End Sub

'----------------------------------------------------------------------
' 7. Prove no retained formula broke. The pre-flight scan said the keep
'    set never references the cut set; this confirms it after the fact
'    on the live workbook rather than trusting the scan.
'----------------------------------------------------------------------
Public Sub AuditRefErrors()

    Dim ws As Worksheet, rng As Range, c As Range
    Dim bad As Long, detail As String, perSheet As Long

    For Each ws In ThisWorkbook.Worksheets
        perSheet = 0
        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.Cells.SpecialCells(xlCellTypeFormulas, xlErrors)
        On Error GoTo 0
        If Not rng Is Nothing Then
            For Each c In rng
                If InStr(1, CStr(c.Text), "#REF!", vbTextCompare) > 0 Then
                    perSheet = perSheet + 1
                    If perSheet = 1 Then
                        detail = detail & vbCrLf & "  " & ws.Name & " first at " & c.Address(False, False)
                    End If
                End If
            Next c
        End If
        bad = bad + perSheet
        If perSheet > 0 Then LogLine "  #REF! on " & ws.Name & ": " & perSheet & " cells"
    Next ws

    If bad = 0 Then
        LogLine "AuditRefErrors: PASS - no #REF! formulas on any retained sheet"
    Else
        LogLine "*** AuditRefErrors: FAIL - " & bad & " #REF! formula cells"
        Say bad & " formulas broke during migration:" & detail & vbCrLf & vbCrLf & _
            "Restore from backup - the dependency scan missed something.", _
            "REVO Migration", vbCritical
    End If

End Sub

'----------------------------------------------------------------------
' Size report
'----------------------------------------------------------------------
Public Sub ReportWorkbookSize()

    Dim mb As Double
    Dim p As String

    p = ThisWorkbook.FullName
    On Error Resume Next
    mb = FileLen(p) / 1048576
    On Error GoTo 0

    LogLine "Size: " & ThisWorkbook.Worksheets.Count & " sheets, " & _
            Format$(mb, "0.00") & " MB on disk (last save)"

    Say "Sheets: " & ThisWorkbook.Worksheets.Count & vbCrLf & _
        "File size on disk: " & Format$(mb, "0.0") & " MB" & vbCrLf & vbCrLf & _
        "(Size reflects the last save, not unsaved changes.)", "REVO Migration"

End Sub

'----------------------------------------------------------------------
' Removes Tube_Dash - 2,251 lines, the only module that deletes cleanly.
' Needs "Trust access to the VBA project object model".
'----------------------------------------------------------------------
Public Sub RemoveTubeDashModule()

    Dim vbc As Object

    On Error GoTo Fail
    Set vbc = ThisWorkbook.VBProject.VBComponents("Tube_Dash")
    ThisWorkbook.VBProject.VBComponents.Remove vbc
    LogLine "RemoveTubeDashModule: Tube_Dash removed (2,251 lines)"
    Say "Tube_Dash removed (2,251 lines).", "REVO Migration"
    Exit Sub
Fail:
    LogLine "*** RemoveTubeDashModule failed: " & Err.Description
    Say "Could not remove Tube_Dash: " & Err.Description & vbCrLf & vbCrLf & _
        "Enable File > Options > Trust Center > Trust Center Settings > " & _
        "Macro Settings > Trust access to the VBA project object model.", _
        "REVO Migration", vbExclamation
End Sub
