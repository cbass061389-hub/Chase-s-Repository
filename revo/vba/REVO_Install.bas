Attribute VB_Name = "REVO_Install"
'==============================================================================
' REVO_Install  -  setup, daily driver, and the self test
'
' RUN THIS FIRST:  REVO_Install
'
' It is idempotent. Run it again after any change and it re-seeds what is
' missing without touching what is already there.
'
' WHY THERE IS A SELF TEST
'   None of this could be executed where it was written - there was no Excel in
'   the loop. So rather than assume the sheet layouts hold, REVO_SelfTest walks
'   every assumption the code makes and reports what it actually found. Run it
'   before you trust a number. It is also the fastest way to see what a layout
'   change broke six months from now.
'==============================================================================
Option Explicit

Private Const RESULT_PASS As String = "PASS"
Private Const RESULT_WARN As String = "WARN"
Private Const RESULT_FAIL As String = "FAIL"

'==============================================================================
' INSTALL
'==============================================================================
Public Sub REVO_Install()
    Dim msg As String
    Dim built As Boolean
    Dim answer As VbMsgBoxResult

    answer = MsgBox("REVO system install." & vbCrLf & vbCrLf & _
                    "This will:" & vbCrLf & _
                    "  - create Quality Log, Quality_Codes, Quality Dashboard," & vbCrLf & _
                    "    REVO Scorecard, Cart Velocity and REVO Audit sheets" & vbCrLf & _
                    "  - seed the defect / root-cause taxonomy from your own history" & vbCrLf & _
                    "  - code the 439 historical reject descriptions into Quality Log" & vbCrLf & _
                    "  - rebuild frmReleaseDetails with the quality capture fields" & vbCrLf & _
                    "  - re-point the REVO Floor button to the patched release macro" & vbCrLf & _
                    "  - add a 'Release Status' column to REVO Floor" & vbCrLf & vbCrLf & _
                    "It does not delete anything or change Release Log, Reject Log" & vbCrLf & _
                    "or Rework Tracker history." & vbCrLf & vbCrLf & _
                    "Save a copy of the workbook first. Continue?", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "REVO Install")
    If answer <> vbYes Then Exit Sub

    On Error GoTo Failed

    '--- 1. sheets and taxonomy --------------------------------------------
    REVO_Quality.EnsureQualityCodes
    REVO_Quality.EnsureQualityLog
    REVO_Release.EnsureReleaseLogHeaders
    REVO_Release.EnsureShipmentsHeaders
    REVO_Core.GetOrCreateSheet SH_AUDIT
    msg = msg & "Sheets and taxonomy seeded." & vbCrLf

    '--- 2. history ---------------------------------------------------------
    Dim added As Long
    added = REVO_Quality.BackfillQualityLog(False)
    msg = msg & "Quality history coded: " & added & " events." & vbCrLf

    '--- 3. the form --------------------------------------------------------
    ' A workbook opened from OneDrive or SharePoint reports an https:// path, so
    ' the form code cannot be found beside it. Say so before the file picker
    ' appears, otherwise the prompt looks like a fault.
    If REVO_FormBuilder.IsUrlPath(ThisWorkbook.path) Then
        msg = msg & "Workbook is on OneDrive/SharePoint - form code cannot be " & _
                    "located beside it; searched Downloads and Desktop instead." & vbCrLf
    End If

    ' Deliberately NOT fatal. The form is the most fragile step, and everything
    ' after it is useful without it. A failure here is reported and the install
    ' carries on.
    If REVO_FormBuilder.VBAccessOK() Then
        REVO_FormBuilder.REVO_BuildReleaseForm
        If Len(REVO_FormBuilder.LastBuildError) = 0 Then
            built = True
            msg = msg & "Release form rebuilt with quality capture." & vbCrLf
        Else
            msg = msg & "Release form NOT built - " & REVO_FormBuilder.LastBuildError & vbCrLf
        End If
    Else
        msg = msg & "Release form NOT rebuilt - VBA project access is blocked. " & _
                    "See revo/docs/MANUAL_FORM_BUILD.md." & vbCrLf
    End If

    '--- 4. buttons ---------------------------------------------------------
    REVO_FormBuilder.REVO_RewireButtons
    msg = msg & "Buttons re-pointed." & vbCrLf

    '--- 5. first build -----------------------------------------------------
    REVO_CartVelocity.REVO_UpdateCartVelocity True
    REVO_QualityDash.BuildQualityDashboard False
    REVO_Scorecard.BuildScorecard False
    msg = msg & "Dashboards built." & vbCrLf

    '--- 6. prove it --------------------------------------------------------
    REVO_SelfTest False
    msg = msg & vbCrLf & "Self test written to '" & SH_SELFTEST & "'." & vbCrLf & vbCrLf & _
          "READ THAT SHEET BEFORE YOU TRUST ANY NUMBER." & vbCrLf & vbCrLf & _
          "Then run REVO_ReconcileFloor - there are rows where REVO Floor and " & _
          "Release Log disagree about released-to-date, and that has to be settled " & _
          "before the scorecard means anything."

    REVO_Core.Audit "REVO_Install", "INSTALL", "Workbook", "Completed; form built=" & built

    MsgBox msg, vbInformation, "REVO Install complete"
    Exit Sub

Failed:
    ' Err FIRST. Calling anything here - ResetAppState included - can clear it,
    ' and reporting "Error 0" tells nobody anything.
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    REVO_Core.ResetAppState
    REVO_Core.Audit "REVO_Install", "ERROR", "Workbook", errNum & ": " & errDesc

    MsgBox "Install failed." & vbCrLf & vbCrLf & _
           "Error " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
           "Completed so far:" & vbCrLf & msg, vbCritical, "REVO Install"
    Resume CleanExit
CleanExit:
End Sub

'==============================================================================
' DAILY DRIVER
'
' Point the plan-update button here. Order matters: velocity feeds capacity,
' capacity feeds the scorecard.
'==============================================================================
Public Sub REVO_DailyUpdate()
    On Error GoTo Failed
    REVO_CartVelocity.REVO_UpdateCartVelocity True
    REVO_QualityDash.BuildQualityDashboard False
    REVO_Scorecard.BuildScorecard False
    REVO_Core.Audit "REVO_Install", "DAILY", "Workbook", "Velocity, quality, scorecard rebuilt"
    MsgBox "Cart velocity, quality dashboard and scorecard rebuilt.", vbInformation, "REVO"
    Exit Sub
Failed:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    REVO_Core.ResetAppState
    MsgBox "Daily update failed." & vbCrLf & vbCrLf & _
           "Error " & errNum & ": " & errDesc, vbExclamation, "REVO"
    Resume CleanExit
CleanExit:
End Sub

Public Sub REVO_BackfillQuality()
    REVO_Quality.BackfillQualityLog True
    REVO_QualityDash.BuildQualityDashboard False
End Sub

'==============================================================================
' SELF TEST
'==============================================================================
Public Sub REVO_RunSelfTest()
    REVO_SelfTest True
End Sub

Public Sub REVO_SelfTest(Optional ByVal interactive As Boolean = True)
    Dim ws As Worksheet
    Dim r As Long
    Dim nFail As Long, nWarn As Long

    On Error GoTo Failed
    REVO_Core.PushAppState

    Set ws = REVO_Core.GetOrCreateSheet(SH_SELFTEST)
    ws.Cells.Clear
    ws.Cells.Font.name = "Segoe UI"
    ws.Cells.Font.Size = 10

    ws.Range("A1").Value = "REVO SELF TEST"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.bold = True
    ws.Range("A2").Value = "Run " & Format$(Now, "ddd d mmm yyyy hh:nn") & " by " & REVO_Core.CurrentUser & _
                           ". Every assumption the REVO_* code makes, checked against this workbook."
    ws.Range("A2").Font.Italic = True

    r = 4
    ws.Cells(r, 1).Value = "Check"
    ws.Cells(r, 2).Value = "Result"
    ws.Cells(r, 3).Value = "Found"
    ws.Cells(r, 4).Value = "What it means"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Font.bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
    End With
    r = r + 1

    r = CheckSheets(ws, r, nFail, nWarn)
    r = CheckFloorLayout(ws, r, nFail, nWarn)
    r = CheckCheckboxIntegrity(ws, r, nFail, nWarn)
    r = CheckReleaseLog(ws, r, nFail, nWarn)
    r = CheckReconciliation(ws, r, nFail, nWarn)
    r = CheckFloorLog(ws, r, nFail, nWarn)
    r = CheckQuality(ws, r, nFail, nWarn)
    r = CheckForm(ws, r, nFail, nWarn)

    '--- verdict ------------------------------------------------------------
    r = r + 1
    ws.Cells(r, 1).Value = "VERDICT"
    ws.Cells(r, 1).Font.Size = 14
    ws.Cells(r, 1).Font.bold = True
    If nFail > 0 Then
        ws.Cells(r, 2).Value = nFail & " FAIL, " & nWarn & " WARN"
        ws.Cells(r, 2).Font.Color = RGB(192, 0, 0)
        ws.Cells(r, 4).Value = "Fix the failures before relying on the scorecard. Each row says what to do."
    ElseIf nWarn > 0 Then
        ws.Cells(r, 2).Value = nWarn & " WARN"
        ws.Cells(r, 2).Font.Color = RGB(150, 75, 0)
        ws.Cells(r, 4).Value = "Nothing broken. The warnings are mostly thin data that fills in as you run it."
    Else
        ws.Cells(r, 2).Value = "ALL PASS"
        ws.Cells(r, 2).Font.Color = RGB(0, 97, 0)
        ws.Cells(r, 4).Value = "Every assumption holds against this workbook."
    End If
    ws.Cells(r, 2).Font.Size = 14
    ws.Cells(r, 2).Font.bold = True

    ws.Columns("A:A").ColumnWidth = 44
    ws.Columns("B:B").ColumnWidth = 10
    ws.Columns("C:C").ColumnWidth = 30
    ws.Columns("D:D").ColumnWidth = 84

    REVO_Core.PopAppState

    If interactive Then
        ws.Activate
        ws.Range("A1").Select
        MsgBox "Self test complete." & vbCrLf & vbCrLf & _
               nFail & " failure(s), " & nWarn & " warning(s)." & vbCrLf & vbCrLf & _
               "Detail on the '" & SH_SELFTEST & "' sheet.", _
               IIf(nFail > 0, vbExclamation, vbInformation), "REVO Self Test"
    End If
    Exit Sub

Failed:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    REVO_Core.ResetAppState
    MsgBox "Self test failed to run." & vbCrLf & vbCrLf & _
           "Error " & errNum & ": " & errDesc, vbCritical, "REVO Self Test"
    Resume CleanExit
CleanExit:
End Sub

'------------------------------------------------------------------- checks
Private Function CheckSheets(ByVal ws As Worksheet, ByVal r As Long, _
                             ByRef nFail As Long, ByRef nWarn As Long) As Long
    Dim required As Variant, optionalS As Variant
    Dim i As Long

    required = Array(SH_FLOOR, SH_RELEASE_LOG, SH_REJECT_LOG, SH_REWORK, SH_PLAN_CFG)
    optionalS = Array(SH_FLOOR_LOG, SH_HOLDS, SH_CART_TRK, SH_DAILY_PLAN, SH_WEEK_PLAN, SH_SHIPMENTS)

    For i = LBound(required) To UBound(required)
        If REVO_Core.SheetExists(CStr(required(i))) Then
            r = TestRow(ws, r, "Sheet present: " & required(i), RESULT_PASS, "found", "", nFail, nWarn)
        Else
            r = TestRow(ws, r, "Sheet present: " & required(i), RESULT_FAIL, "missing", _
                    "Required. The release flow and the scorecard both read it.", nFail, nWarn)
        End If
    Next i

    For i = LBound(optionalS) To UBound(optionalS)
        If REVO_Core.SheetExists(CStr(optionalS(i))) Then
            r = TestRow(ws, r, "Sheet present: " & optionalS(i), RESULT_PASS, "found", "", nFail, nWarn)
        Else
            r = TestRow(ws, r, "Sheet present: " & optionalS(i), RESULT_WARN, "missing", _
                    "Optional. Features reading it degrade rather than fail.", nFail, nWarn)
        End If
    Next i

    CheckSheets = r
End Function

Private Function CheckFloorLayout(ByVal ws As Worksheet, ByVal r As Long, _
                                  ByRef nFail As Long, ByRef nWarn As Long) As Long
    Dim wsF As Worksheet, hdr As Long
    Dim names As Variant
    Dim i As Long, c As Long

    Set wsF = REVO_Core.GetSheet(SH_FLOOR)
    If wsF Is Nothing Then CheckFloorLayout = r: Exit Function

    hdr = REVO_Core.FindHeaderRow(wsF, Array("Release?", "Release"), 30, 40)
    If hdr = 0 Then
        r = TestRow(ws, r, "REVO Floor header row", RESULT_FAIL, "not found", _
                "No cell reading 'Release?' in the first 30 rows. The release macro cannot run.", nFail, nWarn)
        CheckFloorLayout = r
        Exit Function
    End If

    r = TestRow(ws, r, "REVO Floor header row", RESULT_PASS, "row " & hdr, _
            "Located by caption, not hardcoded - a moved header will not break it.", nFail, nWarn)

    names = Array("WO#", "SKU", "CART", "Qty", "Release?", "Released To Date", "Remaining")
    For i = LBound(names) To UBound(names)
        c = REVO_Core.FindCol(wsF, hdr, Array(CStr(names(i))))
        If c > 0 Then
            r = TestRow(ws, r, "REVO Floor column: " & names(i), RESULT_PASS, "col " & ColLetter(c), "", nFail, nWarn)
        Else
            r = TestRow(ws, r, "REVO Floor column: " & names(i), RESULT_FAIL, "missing", _
                    "The release macro needs this column on row " & hdr & ".", nFail, nWarn)
        End If
    Next i

    c = REVO_Core.FindCol(wsF, hdr, Array("Release Status"))
    If c > 0 Then
        r = TestRow(ws, r, "REVO Floor: Release Status column", RESULT_PASS, "col " & ColLetter(c), _
                "Status text goes here so it never lands in the checkbox column.", nFail, nWarn)
    Else
        r = TestRow(ws, r, "REVO Floor: Release Status column", RESULT_WARN, "not yet added", _
                "Run REVO_Install, or it is added the first time the release macro runs.", nFail, nWarn)
    End If

    CheckFloorLayout = r
End Function

' The original defect: status text written into a cell-checkbox column kills
' the checkbox. This finds any survivors.
Private Function CheckCheckboxIntegrity(ByVal ws As Worksheet, ByVal r As Long, _
                                        ByRef nFail As Long, ByRef nWarn As Long) As Long
    Dim wsF As Worksheet, hdr As Long, cRel As Long, lastR As Long, i As Long
    Dim v As Variant, n As Long, rows_ As String

    Set wsF = REVO_Core.GetSheet(SH_FLOOR)
    If wsF Is Nothing Then CheckCheckboxIntegrity = r: Exit Function

    hdr = REVO_Core.FindHeaderRow(wsF, Array("Release?", "Release"), 30, 40)
    If hdr = 0 Then CheckCheckboxIntegrity = r: Exit Function
    cRel = REVO_Core.FindCol(wsF, hdr, Array("Release?", "Release"))
    If cRel = 0 Then CheckCheckboxIntegrity = r: Exit Function

    lastR = REVO_Core.LastDataRow(wsF, cRel, hdr)
    For i = hdr + 1 To lastR
        v = wsF.Cells(i, cRel).Value
        If Not IsEmpty(v) Then
            If VarType(v) = vbString Then
                If Len(Trim$(CStr(v))) > 0 Then
                    n = n + 1
                    If n <= 12 Then rows_ = rows_ & i & " (" & CStr(v) & "), "
                End If
            End If
        End If
    Next i

    If n = 0 Then
        r = TestRow(ws, r, "Release? column holds only checkboxes", RESULT_PASS, "clean", _
                "No text in the checkbox column. Every row is clickable.", nFail, nWarn)
    Else
        r = TestRow(ws, r, "Release? column holds only checkboxes", RESULT_FAIL, n & " text cell(s)", _
                "Rows " & rows_ & "- the old macro wrote status text here and killed the checkbox. " & _
                "Run REVO_ResetCart for each, or clear the cell contents (the checkbox format returns).", _
                nFail, nWarn)
    End If

    CheckCheckboxIntegrity = r
End Function

Private Function CheckReleaseLog(ByVal ws As Worksheet, ByVal r As Long, _
                                 ByRef nFail As Long, ByRef nWarn As Long) As Long
    Dim wsL As Worksheet, hdr As Long, i As Long, c As Long, lastR As Long
    Dim names As Variant

    Set wsL = REVO_Core.GetSheet(SH_RELEASE_LOG)
    If wsL Is Nothing Then CheckReleaseLog = r: Exit Function

    hdr = REVO_Core.FindHeaderRow(wsL, Array("Qty Released"), 10, 15)
    If hdr = 0 Then
        r = TestRow(ws, r, "Release Log header", RESULT_FAIL, "not found", _
                "Released-to-date cannot be verified without it.", nFail, nWarn)
        CheckReleaseLog = r
        Exit Function
    End If

    names = Array("Date", "SKU", "Cart", "Qty Released", "Work Order")
    For i = LBound(names) To UBound(names)
        c = REVO_Core.FindCol(wsL, hdr, Array(CStr(names(i))))
        If c = 0 Then
            r = TestRow(ws, r, "Release Log column: " & names(i), RESULT_FAIL, "missing", _
                    "Needed to tie a release back to a cart.", nFail, nWarn)
        End If
    Next i

    c = REVO_Core.FindCol(wsL, hdr, Array("Date"))
    lastR = REVO_Core.LastDataRow(wsL, c, hdr)
    r = TestRow(ws, r, "Release Log rows", RESULT_PASS, CStr(lastR - hdr), _
            "History available to the scorecard and the quality rates.", nFail, nWarn)

    CheckReleaseLog = r
End Function

Private Function CheckReconciliation(ByVal ws As Worksheet, ByVal r As Long, _
                                     ByRef nFail As Long, ByRef nWarn As Long) As Long
    Dim wsF As Worksheet, hdr As Long, lastR As Long, i As Long
    Dim cWO As Long, cCart As Long, cSku As Long, cRTD As Long, cQty As Long
    Dim sheetVal As Double, logVal As Double, n As Long, nGhost As Long

    Set wsF = REVO_Core.GetSheet(SH_FLOOR)
    If wsF Is Nothing Then CheckReconciliation = r: Exit Function

    hdr = REVO_Core.FindHeaderRow(wsF, Array("Release?"), 30, 40)
    If hdr = 0 Then CheckReconciliation = r: Exit Function

    cWO = REVO_Core.FindCol(wsF, hdr, Array("WO#", "WO"))
    cCart = REVO_Core.FindCol(wsF, hdr, Array("CART", "Cart"))
    cSku = REVO_Core.FindCol(wsF, hdr, Array("SKU"))
    cRTD = REVO_Core.FindCol(wsF, hdr, Array("Released To Date"))
    cQty = REVO_Core.FindCol(wsF, hdr, Array("Qty"))
    If cWO = 0 Or cCart = 0 Or cRTD = 0 Then CheckReconciliation = r: Exit Function

    lastR = REVO_Core.LastDataRow(wsF, cCart, hdr)
    For i = hdr + 1 To lastR
        If Len(REVO_Core.SafeS(wsF.Cells(i, cSku).Value)) > 0 Then
            sheetVal = REVO_Core.SafeD(wsF.Cells(i, cRTD).Value)
            logVal = REVO_Release.ReleasedToDateFromLog( _
                        REVO_Core.SafeS(wsF.Cells(i, cWO).Value), _
                        REVO_Core.SafeS(wsF.Cells(i, cCart).Value))
            If Abs(sheetVal - logVal) > 0.001 Then
                n = n + 1
                ' The specific failure that hides the form: sheet claims fully
                ' released, log has nothing.
                If logVal = 0 And sheetVal > 0 And sheetVal >= REVO_Core.SafeD(wsF.Cells(i, cQty).Value) Then
                    nGhost = nGhost + 1
                End If
            End If
        End If
    Next i

    If n = 0 Then
        r = TestRow(ws, r, "REVO Floor agrees with Release Log", RESULT_PASS, "0 mismatches", _
                "Released-to-date is trustworthy.", nFail, nWarn)
    Else
        r = TestRow(ws, r, "REVO Floor agrees with Release Log", RESULT_FAIL, n & " mismatches", _
                "Run REVO_ReconcileFloor. Until this is settled the plan numbers are guesses.", nFail, nWarn)
    End If

    If nGhost > 0 Then
        r = TestRow(ws, r, "Rows that would skip the form silently", RESULT_FAIL, CStr(nGhost), _
                "These show fully released on the sheet with nothing in the log. Tick one and " & _
                "the old macro showed no form at all - the original complaint. " & _
                "REVO_ReconcileFloor clears them.", nFail, nWarn)
    Else
        r = TestRow(ws, r, "Rows that would skip the form silently", RESULT_PASS, "none", "", nFail, nWarn)
    End If

    CheckReconciliation = r
End Function

Private Function CheckFloorLog(ByVal ws As Worksheet, ByVal r As Long, _
                               ByRef nFail As Long, ByRef nWarn As Long) As Long
    Dim wsL As Worksheet, hdr As Long, lastR As Long, i As Long
    Dim colDate As Long
    Dim d As Object
    Dim dt As Date

    Set wsL = REVO_Core.GetSheet(SH_FLOOR_LOG)
    If wsL Is Nothing Then
        r = TestRow(ws, r, "Floor Log snapshots", RESULT_WARN, "sheet missing", _
                "Cart velocity has nothing to measure. Capacity stays on the Plan Config assumption.", nFail, nWarn)
        CheckFloorLog = r
        Exit Function
    End If

    hdr = REVO_Core.FindHeaderRow(wsL, Array("Snapshot Date"), 10, 20)
    If hdr = 0 Then
        r = TestRow(ws, r, "Floor Log snapshots", RESULT_WARN, "header not found", _
                "Cart velocity cannot read it.", nFail, nWarn)
        CheckFloorLog = r
        Exit Function
    End If

    colDate = REVO_Core.FindCol(wsL, hdr, Array("Snapshot Date"))
    Set d = CreateObject("Scripting.Dictionary")
    lastR = REVO_Core.LastDataRow(wsL, colDate, hdr)
    For i = hdr + 1 To lastR
        dt = REVO_Core.SafeDate(wsL.Cells(i, colDate).Value, 0)
        If dt > 0 Then
            If Not d.exists(CLng(dt)) Then d.Add CLng(dt), 1
        End If
    Next i

    If d.Count >= REVO_CartVelocity.MIN_SAMPLE + 1 Then
        r = TestRow(ws, r, "Floor Log snapshots", RESULT_PASS, d.Count & " snapshots", _
                "Enough history for a measured run rate.", nFail, nWarn)
    ElseIf d.Count >= 2 Then
        r = TestRow(ws, r, "Floor Log snapshots", RESULT_WARN, d.Count & " snapshots", _
                "Only " & (d.Count - 1) & " interval(s). Cart velocity will measure but will NOT " & _
                "auto-apply - it needs " & REVO_CartVelocity.MIN_SAMPLE & " clean intervals. " & _
                "Run LOG FLOOR daily and it fills in on its own.", nFail, nWarn)
    Else
        r = TestRow(ws, r, "Floor Log snapshots", RESULT_WARN, d.Count & " snapshot(s)", _
                "Two or more are needed before anything can be measured.", nFail, nWarn)
    End If

    CheckFloorLog = r
End Function

Private Function CheckQuality(ByVal ws As Worksheet, ByVal r As Long, _
                              ByRef nFail As Long, ByRef nWarn As Long) As Long
    Dim wsQ As Worksheet, lastR As Long
    Dim codes As Worksheet
    Dim nDefect As Long

    Set codes = REVO_Core.GetSheet(SH_QUAL_CODES)
    If codes Is Nothing Then
        r = TestRow(ws, r, "Quality_Codes seeded", RESULT_FAIL, "missing", "Run REVO_Install.", nFail, nWarn)
    Else
        nDefect = UBound(REVO_Quality.DefectTypeList())
        If nDefect > 1 Then
            r = TestRow(ws, r, "Quality_Codes seeded", RESULT_PASS, nDefect & " defect types", _
                    "Edit the sheet to change what the form offers - no code change needed.", nFail, nWarn)
        Else
            r = TestRow(ws, r, "Quality_Codes seeded", RESULT_FAIL, "empty", "Run REVO_Install.", nFail, nWarn)
        End If
    End If

    Set wsQ = REVO_Core.GetSheet(SH_QUAL_LOG)
    If wsQ Is Nothing Then
        r = TestRow(ws, r, "Quality Log populated", RESULT_FAIL, "missing", "Run REVO_Install.", nFail, nWarn)
    Else
        lastR = wsQ.Cells(wsQ.Rows.Count, QLC_EVENTID).End(xlUp).row
        If lastR > QL_HEADER_ROW Then
            r = TestRow(ws, r, "Quality Log populated", RESULT_PASS, (lastR - QL_HEADER_ROW) & " events", _
                    "Dashboard has history to work with.", nFail, nWarn)
        Else
            r = TestRow(ws, r, "Quality Log populated", RESULT_WARN, "empty", _
                    "Run REVO_BackfillQuality to code the existing reject history.", nFail, nWarn)
        End If
    End If

    CheckQuality = r
End Function

Private Function CheckForm(ByVal ws As Worksheet, ByVal r As Long, _
                           ByRef nFail As Long, ByRef nWarn As Long) As Long
    Dim vbc As Object, ctl As Object
    Dim needed As Variant, i As Long
    Dim missing As String, nMissing As Long
    Dim found As Boolean

    On Error GoTo NoAccess

    Set vbc = ThisWorkbook.VBProject.VBComponents("frmReleaseDetails")

    needed = Array("txtCartQtyRelease", "txtRework", "txtBGrade", "txtReject", _
                   "cboDisposition", "txtLineQty", "cboDefect", "cboLocation", _
                   "cboOp", "cboRoot", "cboAction", "txtLineNotes", _
                   "btnAddLine", "btnRemoveLine", "lstLines", _
                   "lblToInventory", "lblAllocation", "btnSubmit", "btnCancel")

    For i = LBound(needed) To UBound(needed)
        found = False
        On Error Resume Next
        Set ctl = vbc.Designer.Controls(CStr(needed(i)))
        found = Not (ctl Is Nothing)
        Set ctl = Nothing
        On Error GoTo NoAccess
        If Not found Then
            nMissing = nMissing + 1
            missing = missing & needed(i) & ", "
        End If
    Next i

    If nMissing = 0 Then
        r = TestRow(ws, r, "Release form has its controls", RESULT_PASS, "all present", _
                "The release macro can prime and read every field.", nFail, nWarn)
    Else
        r = TestRow(ws, r, "Release form has its controls", RESULT_FAIL, nMissing & " missing", _
                "Missing: " & missing & "Run REVO_BuildReleaseForm.", nFail, nWarn)
    End If

    CheckForm = r
    Exit Function

NoAccess:
    r = TestRow(ws, r, "Release form has its controls", RESULT_WARN, "cannot inspect", _
            "VBA project access is blocked, or the form does not exist yet. " & _
            "Tick Trust access to the VBA project object model and run REVO_BuildReleaseForm.", nFail, nWarn)
    CheckForm = r
    Resume Done
Done:
End Function

'------------------------------------------------------------------ helpers
Private Function TestRow(ByVal ws As Worksheet, ByVal r As Long, ByVal check As String, _
                     ByVal result As String, ByVal found As String, ByVal meaning As String, _
                     ByRef nFail As Long, ByRef nWarn As Long) As Long
    ws.Cells(r, 1).Value = check
    ws.Cells(r, 2).Value = result
    ws.Cells(r, 3).Value = found
    ws.Cells(r, 4).Value = meaning

    Select Case result
        Case RESULT_PASS
            ws.Cells(r, 2).Font.Color = RGB(0, 97, 0)
        Case RESULT_WARN
            ws.Cells(r, 2).Font.Color = RGB(150, 75, 0)
            nWarn = nWarn + 1
        Case RESULT_FAIL
            ws.Cells(r, 2).Font.Color = RGB(192, 0, 0)
            ws.Cells(r, 2).Font.bold = True
            nFail = nFail + 1
    End Select

    TestRow = r + 1
End Function

Private Function ColLetter(ByVal c As Long) As String
    Dim n As Long, s As String
    n = c
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function
