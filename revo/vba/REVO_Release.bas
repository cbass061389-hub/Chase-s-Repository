Attribute VB_Name = "REVO_Release"
'==============================================================================
' REVO_Release  -  patched cart release flow
'
' Entry point: REVO_ReleaseCarts   (the installer re-points the REVO Floor
' button to this; the legacy ReleaseToShippingAndReceiving in REVO_Ops is left
' untouched so you can compare, then delete it when you are happy)
'
' WHAT CHANGED, AND WHY
'
' 1. The checkbox column is never written to with text.
'    The old code wrote "RELEASED" / "PARTIAL" / "NO DATA" into column W. W is
'    an Excel cell-checkbox column, and a checkbox cell holding text stops
'    being a checkbox - permanently. That is what stranded cart 34 (row 13)
'    with 48 units still on it. Status now goes to its own "Release Status"
'    column and W is reset to FALSE after a successful pass.
'
' 2. Nothing is skipped silently.
'    The old code stamped RELEASED and moved on whenever released-to-date >=
'    cart qty, without showing the form. Ten rows on REVO Floor are in exactly
'    that state right now. Every skip is collected and reported at the end,
'    with the reason.
'
' 3. Released-to-date is cross-checked against Release Log.
'    Column Y was both a formula and a VBA write target, so the two fought and
'    VBA won on some rows and not others. When the sheet and the log disagree,
'    the operator is told and offered the log's number - the log is the record
'    of what actually shipped.
'
' 4. ScreenUpdating is restored around the form.
'    A modal UserForm shown while ScreenUpdating is False may paint blank or
'    open behind Excel.
'
' 5. The header row is found, not assumed.
'    The old code started at row 2. The header is row 9 and data starts at 11.
'==============================================================================
Option Explicit

' Recipients for the release notification. Edit here or set to "" to suppress.
Private Const MAIL_TO  As String = "mbyron@predatorgroup.com"
Private Const MAIL_CC  As String = "gkeys@predatorgroup.com; jausler@predatorgroup.com; " & _
                                   "cbass@predatorgroup.com; osantiago@predatorgroup.com; " & _
                                   "cfournier@predatorgroup.com; abirmingham@predatorgroup.com; " & _
                                   "rorr@predatorgroup.com; cnewman@predatorgroup.com"

' Resolved column positions on REVO Floor.
Private Type FloorMap
    ws        As Worksheet
    headerRow As Long
    firstRow  As Long
    lastRow   As Long
    cWO       As Long
    cSku      As Long
    cCart     As Long
    cQty      As Long
    cRelease  As Long   ' the checkbox column
    cRelToDate As Long
    cRemain   As Long
    cStatus   As Long   ' new - status text lives here, never in cRelease
    ok        As Boolean
    problem   As String
End Type

'==============================================================================
' MAIN ENTRY POINT
'==============================================================================
Public Sub REVO_ReleaseCarts()
    Dim fm As FloorMap
    Dim frm As frmReleaseDetails
    Dim r As Long
    Dim WO As String, sku As String, cart As String
    Dim qty As Double, priorRel As Double, logRel As Double, remaining As Double
    Dim reqQty As Double, newRel As Double
    Dim toInv As Double, reworkQty As Double, bGradeQty As Double, rejectQty As Double
    Dim summary As Object
    Dim detail As String, skips As String, warnings As String
    Dim nReleased As Long, nSkipped As Long, nCancelled As Long
    Dim releaseID As String
    Dim today As Date
    Dim answer As VbMsgBoxResult

    On Error GoTo FailFast

    '--- shift gate: the floor cannot release carts at night -----------------
    If Not REVO_Core.CanReleaseNow Then
        answer = MsgBox("It is currently the night shift, and releases are a day-shift activity." & vbCrLf & vbCrLf & _
                        "Carry on anyway?", vbQuestion + vbYesNo + vbDefaultButton2, "REVO Release")
        If answer = vbNo Then Exit Sub
    End If

    fm = MapFloorSheet()
    If Not fm.ok Then
        MsgBox "Cannot read REVO Floor." & vbCrLf & vbCrLf & fm.problem, vbExclamation, "REVO Release"
        Exit Sub
    End If

    REVO_Quality.EnsureQualityCodes
    REVO_Quality.EnsureQualityLog
    EnsureReleaseLogHeaders
    EnsureShipmentsHeaders

    Set summary = CreateObject("Scripting.Dictionary")
    summary.CompareMode = vbTextCompare
    today = Date

    REVO_Core.PushAppState

    For r = fm.firstRow To fm.lastRow

        If Not REVO_Core.Boolish(fm.ws.Cells(r, fm.cRelease).Value) Then GoTo NextRow

        WO = REVO_Core.SafeS(fm.ws.Cells(r, fm.cWO).Value)
        sku = REVO_Core.SafeS(fm.ws.Cells(r, fm.cSku).Value)
        cart = REVO_Core.SafeS(fm.ws.Cells(r, fm.cCart).Value)
        qty = REVO_Core.SafeD(fm.ws.Cells(r, fm.cQty).Value)

        '--- incomplete row: say so, do not pretend it was released ---------
        If Len(sku) = 0 Or qty <= 0 Then
            fm.ws.Cells(r, fm.cStatus).Value = "NO DATA"
            fm.ws.Cells(r, fm.cRelease).Value = False
            skips = skips & "  Row " & r & " (cart " & cart & "): no SKU or qty on the row." & vbCrLf
            nSkipped = nSkipped + 1
            GoTo NextRow
        End If

        '--- what the sheet thinks vs what the log says ----------------------
        priorRel = REVO_Core.SafeD(fm.ws.Cells(r, fm.cRelToDate).Value)
        logRel = ReleasedToDateFromLog(WO, cart)

        If Abs(priorRel - logRel) > 0.001 Then
            REVO_Core.BeginModalDialog
            answer = MsgBox( _
                "Cart " & cart & " / " & WO & " (" & sku & ")" & vbCrLf & vbCrLf & _
                "REVO Floor says released to date:  " & priorRel & vbCrLf & _
                "Release Log actually records:      " & logRel & vbCrLf & vbCrLf & _
                "Release Log is the record of what shipped." & vbCrLf & vbCrLf & _
                "Use the Release Log figure (" & logRel & ")?" & vbCrLf & _
                "Choosing No keeps the sheet value and carries on.", _
                vbQuestion + vbYesNoCancel + vbDefaultButton1, "Released-to-date mismatch")
            REVO_Core.EndModalDialog

            If answer = vbCancel Then GoTo Finish
            If answer = vbYes Then
                priorRel = logRel
                fm.ws.Cells(r, fm.cRelToDate).Value = priorRel
                REVO_Core.Audit "REVO_Release", "CORRECT", fm.ws.name & "!R" & r, _
                                "Released-to-date reset to Release Log value " & logRel
            End If
            warnings = warnings & "  Cart " & cart & " (" & WO & "): sheet " & _
                       priorRel & " vs log " & logRel & vbCrLf
        End If

        remaining = qty - priorRel

        '--- nothing left on the cart: report it, never stamp it silently ----
        If remaining <= 0 Then
            skips = skips & "  Cart " & cart & " (" & WO & ", " & sku & "): " & _
                    "already shows " & priorRel & " of " & qty & " released - nothing left." & vbCrLf
            fm.ws.Cells(r, fm.cStatus).Value = "RELEASED"
            fm.ws.Cells(r, fm.cRemain).Value = 0
            fm.ws.Cells(r, fm.cRelease).Value = False
            nSkipped = nSkipped + 1
            GoTo NextRow
        End If

        '=================== the form ======================================
        Set frm = New frmReleaseDetails
        frm.Prime sku, cart, WO, CLng(remaining), CLng(qty), CLng(priorRel)

        REVO_Core.BeginModalDialog
        frm.Show
        REVO_Core.EndModalDialog

        If frm.UserCancelled Then
            Unload frm
            Set frm = Nothing
            nCancelled = nCancelled + 1
            GoTo NextRow                    ' checkbox stays ticked on purpose
        End If

        reqQty = frm.RequestedQty
        toInv = frm.ReleasedToInventory
        reworkQty = frm.Rework
        bGradeQty = frm.BGrade
        rejectQty = frm.Reject

        releaseID = REVO_Core.NewEventID("REL")

        '--- quality events, one per disposition ----------------------------
        If reworkQty > 0 Then
            WriteQE "REWORK", reworkQty, today, WO, sku, cart, releaseID, _
                    frm.ReworkDefect, frm.ReworkLocation, frm.ReworkOp, _
                    frm.ReworkRootCause, frm.ReworkAction, frm.ReworkNotes
        End If
        If rejectQty > 0 Then
            WriteQE "REJECT", rejectQty, today, WO, sku, cart, releaseID, _
                    frm.RejectDefect, frm.RejectLocation, frm.RejectOp, _
                    frm.RejectRootCause, frm.RejectAction, frm.RejectNotes
        End If
        If bGradeQty > 0 Then
            WriteQE "B GRADE", bGradeQty, today, WO, sku, cart, releaseID, _
                    frm.BGradeDefect, frm.BGradeLocation, "", _
                    frm.BGradeRootCause, "Downgraded to B", frm.BGradeNotes
        End If

        Unload frm
        Set frm = Nothing

        '--- legacy writes, so existing reporting keeps working --------------
        If toInv > 0 Then
            AppendShipments today, sku, toInv, cart
            Accumulate summary, sku, toInv
        End If
        If reworkQty > 0 Then AppendRework today, WO, sku, cart, reworkQty
        If rejectQty > 0 Then AppendReject today, WO, sku, cart, rejectQty, RejectSummaryText(frm)
        If bGradeQty > 0 Then UpdateBGrade sku, bGradeQty

        newRel = priorRel + reqQty

        AppendReleaseLog today, sku, cart, toInv, WO, reworkQty, bGradeQty, _
                         rejectQty, reqQty, qty - newRel

        '--- cart state -----------------------------------------------------
        fm.ws.Cells(r, fm.cRelToDate).Value = newRel
        fm.ws.Cells(r, fm.cRemain).Value = qty - newRel
        fm.ws.Cells(r, fm.cStatus).Value = IIf(newRel >= qty, "RELEASED", "PARTIAL")
        fm.ws.Cells(r, fm.cRelease).Value = False     ' checkbox stays a checkbox

        nReleased = nReleased + 1

        detail = detail & _
            "Cart: " & cart & vbCrLf & _
            "SKU: " & sku & vbCrLf & _
            "Work Order: " & WO & vbCrLf & _
            "Released this pass: " & reqQty & " of " & qty & vbCrLf & _
            "To inventory: " & toInv & vbCrLf & _
            "B Grade: " & bGradeQty & "   Rework: " & reworkQty & "   Reject: " & rejectQty & vbCrLf
        If rejectQty > 0 Then detail = detail & "Reject: " & RejectSummaryText(frm) & vbCrLf
        If newRel < qty Then
            detail = detail & "*** PARTIAL - " & (qty - newRel) & " remaining on cart ***" & vbCrLf
        End If
        detail = detail & vbCrLf

NextRow:
    Next r

Finish:
    REVO_Core.PopAppState

    If nReleased > 0 Then
        ' Late-bound on purpose: these live in the legacy modules, and deleting
        ' those must not stop this one compiling.
        On Error Resume Next
        Application.Run "RunWeeklyREVOUpdate_Safe"
        Application.Run "RefreshRejectDashboard_Full"
        On Error GoTo 0
        REVO_QualityDash.BuildQualityDashboard False
        REVO_Scorecard.BuildScorecard False
    End If

    If summary.Count > 0 And Len(MAIL_TO) > 0 Then
        On Error Resume Next
        SendReleaseMail summary, detail, today
        On Error GoTo 0
    End If

    ReportOutcome nReleased, nSkipped, nCancelled, skips, warnings
    REVO_Core.Audit "REVO_Release", "RUN", SH_FLOOR, _
                    nReleased & " released, " & nSkipped & " skipped, " & nCancelled & " cancelled"
    Exit Sub

FailFast:
    Dim msg As String
    msg = "REVO_ReleaseCarts failed at row " & r & "." & vbCrLf & vbCrLf & _
          "Error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    REVO_Core.ResetAppState
    REVO_Core.Audit "REVO_Release", "ERROR", SH_FLOOR, msg
    MsgBox msg, vbCritical, "REVO Release"
End Sub

'==============================================================================
' OUTCOME REPORT  -  the operator always finds out what happened
'==============================================================================
Private Sub ReportOutcome(ByVal nReleased As Long, ByVal nSkipped As Long, _
                          ByVal nCancelled As Long, ByVal skips As String, _
                          ByVal warnings As String)
    Dim m As String
    Dim icon As VbMsgBoxStyle

    m = "Released:   " & nReleased & vbCrLf & _
        "Skipped:    " & nSkipped & vbCrLf & _
        "Cancelled:  " & nCancelled & vbCrLf

    If Len(skips) > 0 Then
        m = m & vbCrLf & "SKIPPED ROWS" & vbCrLf & skips
    End If
    If Len(warnings) > 0 Then
        m = m & vbCrLf & "RELEASED-TO-DATE MISMATCHES" & vbCrLf & warnings & vbCrLf & _
            "Run REVO_ReconcileFloor to see every row where REVO Floor and Release Log disagree."
    End If

    If nReleased = 0 And nSkipped = 0 And nCancelled = 0 Then
        m = "Nothing was ticked in the Release? column, so there was nothing to release." & vbCrLf & vbCrLf & _
            "Tick the checkbox in the Release? column on the carts you want to release, then run this again."
        icon = vbInformation
    ElseIf nSkipped > 0 Or Len(warnings) > 0 Then
        icon = vbExclamation
    Else
        icon = vbInformation
    End If

    MsgBox m, icon, "REVO Release"
End Sub

'==============================================================================
' SHEET MAPPING
'==============================================================================
Private Function MapFloorSheet() As FloorMap
    Dim fm As FloorMap

    Set fm.ws = REVO_Core.GetSheet(SH_FLOOR)
    If fm.ws Is Nothing Then
        fm.problem = "Sheet '" & SH_FLOOR & "' not found."
        MapFloorSheet = fm
        Exit Function
    End If

    fm.headerRow = REVO_Core.FindHeaderRow(fm.ws, Array("Release?", "Release"), 30, 40)
    If fm.headerRow = 0 Then
        fm.problem = "Could not find the header row - no cell reading 'Release?' in the first 30 rows."
        MapFloorSheet = fm
        Exit Function
    End If

    fm.cWO = REVO_Core.FindCol(fm.ws, fm.headerRow, Array("WO#", "WO", "Work Order"))
    fm.cSku = REVO_Core.FindCol(fm.ws, fm.headerRow, Array("SKU"))
    fm.cCart = REVO_Core.FindCol(fm.ws, fm.headerRow, Array("CART", "Cart"))
    fm.cQty = REVO_Core.FindCol(fm.ws, fm.headerRow, Array("Qty", "Quantity"))
    fm.cRelease = REVO_Core.FindCol(fm.ws, fm.headerRow, Array("Release?", "Release"))
    fm.cRelToDate = REVO_Core.FindCol(fm.ws, fm.headerRow, Array("Released To Date", "Released to date"))
    fm.cRemain = REVO_Core.FindCol(fm.ws, fm.headerRow, Array("Remaining"))

    If fm.cWO = 0 Or fm.cSku = 0 Or fm.cCart = 0 Or fm.cQty = 0 Or fm.cRelease = 0 Then
        fm.problem = "Header row " & fm.headerRow & " is missing one of: WO#, SKU, CART, Qty, Release?"
        MapFloorSheet = fm
        Exit Function
    End If
    If fm.cRelToDate = 0 Or fm.cRemain = 0 Then
        fm.problem = "Header row " & fm.headerRow & " is missing 'Released To Date' or 'Remaining'."
        MapFloorSheet = fm
        Exit Function
    End If

    ' Only now that the sheet has proved readable is it safe to add a column.
    fm.cStatus = REVO_Core.FindOrAddCol(fm.ws, fm.headerRow, "Release Status")

    fm.firstRow = fm.headerRow + 1
    fm.lastRow = REVO_Core.LastDataRow(fm.ws, fm.cRelease, fm.headerRow)
    If fm.lastRow < fm.firstRow Then fm.lastRow = fm.firstRow

    fm.ok = True
    MapFloorSheet = fm
End Function

'==============================================================================
' RELEASE LOG AS THE RECORD OF TRUTH
'==============================================================================
' Total pulled off a given WO + cart, ignoring recalled rows.
Public Function ReleasedToDateFromLog(ByVal WO As String, ByVal cart As String) As Double
    Dim ws As Worksheet, hdr As Long, r As Long, lastR As Long
    Dim cCart As Long, cWO As Long, cCartQty As Long, cQtyRel As Long, cRecall As Long
    Dim t As Double
    Dim rowWO As String, rowCart As String
    Dim pulled As Double

    Set ws = REVO_Core.GetSheet(SH_RELEASE_LOG)
    If ws Is Nothing Then Exit Function

    hdr = REVO_Core.FindHeaderRow(ws, Array("Qty Released", "Cart"), 10, 15)
    If hdr = 0 Then Exit Function

    cCart = REVO_Core.FindCol(ws, hdr, Array("Cart"))
    cWO = REVO_Core.FindCol(ws, hdr, Array("Work Order", "WO", "WO#"))
    cQtyRel = REVO_Core.FindCol(ws, hdr, Array("Qty Released"))
    cCartQty = REVO_Core.FindCol(ws, hdr, Array("Cart Qty Released"))
    cRecall = REVO_Core.FindCol(ws, hdr, Array("Recalled"))
    If cCart = 0 Or cWO = 0 Then Exit Function

    lastR = REVO_Core.LastDataRow(ws, cCart, hdr)

    For r = hdr + 1 To lastR
        rowWO = UCase$(REVO_Core.SafeS(ws.Cells(r, cWO).Value))
        rowCart = UCase$(REVO_Core.SafeS(ws.Cells(r, cCart).Value))
        If rowWO = UCase$(Trim$(WO)) And rowCart = UCase$(Trim$(cart)) Then
            If cRecall = 0 Or Not REVO_Core.Boolish(ws.Cells(r, cRecall).Value) Then
                ' "Cart Qty Released" is what came off the cart this pass.
                ' Older rows predate that column - fall back to Qty Released.
                pulled = 0
                If cCartQty > 0 Then pulled = REVO_Core.SafeD(ws.Cells(r, cCartQty).Value)
                If pulled <= 0 And cQtyRel > 0 Then pulled = REVO_Core.SafeD(ws.Cells(r, cQtyRel).Value)
                t = t + pulled
            End If
        End If
    Next r

    ReleasedToDateFromLog = t
End Function

'==============================================================================
' RECONCILIATION  -  run this once before you trust the board
'==============================================================================
Public Sub REVO_ReconcileFloor()
    Dim fm As FloorMap
    Dim r As Long, n As Long, nFix As Long
    Dim WO As String, cart As String, sku As String
    Dim sheetVal As Double, logVal As Double
    Dim rpt As String
    Dim answer As VbMsgBoxResult

    fm = MapFloorSheet()
    If Not fm.ok Then
        MsgBox "Cannot read REVO Floor." & vbCrLf & vbCrLf & fm.problem, vbExclamation
        Exit Sub
    End If

    REVO_Core.PushAppState
    For r = fm.firstRow To fm.lastRow
        sku = REVO_Core.SafeS(fm.ws.Cells(r, fm.cSku).Value)
        If Len(sku) > 0 Then
            WO = REVO_Core.SafeS(fm.ws.Cells(r, fm.cWO).Value)
            cart = REVO_Core.SafeS(fm.ws.Cells(r, fm.cCart).Value)
            sheetVal = REVO_Core.SafeD(fm.ws.Cells(r, fm.cRelToDate).Value)
            logVal = ReleasedToDateFromLog(WO, cart)
            If Abs(sheetVal - logVal) > 0.001 Then
                n = n + 1
                If n <= 30 Then
                    rpt = rpt & "  Row " & r & "  cart " & cart & " (" & WO & ")  " & _
                          "sheet " & sheetVal & "  ->  log " & logVal & vbCrLf
                End If
            End If
        End If
    Next r
    REVO_Core.PopAppState

    If n = 0 Then
        MsgBox "REVO Floor and Release Log agree on every row." & vbCrLf & vbCrLf & _
               "Released-to-date is trustworthy.", vbInformation, "Reconcile"
        Exit Sub
    End If

    If n > 30 Then rpt = rpt & "  ... and " & (n - 30) & " more." & vbCrLf

    answer = MsgBox(n & " row(s) where REVO Floor disagrees with Release Log:" & vbCrLf & vbCrLf & _
                    rpt & vbCrLf & _
                    "Release Log is the record of what actually shipped." & vbCrLf & vbCrLf & _
                    "Overwrite Released To Date and Remaining on those rows with the log figures?", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "Reconcile REVO Floor")
    If answer <> vbYes Then Exit Sub

    REVO_Core.PushAppState
    For r = fm.firstRow To fm.lastRow
        sku = REVO_Core.SafeS(fm.ws.Cells(r, fm.cSku).Value)
        If Len(sku) > 0 Then
            WO = REVO_Core.SafeS(fm.ws.Cells(r, fm.cWO).Value)
            cart = REVO_Core.SafeS(fm.ws.Cells(r, fm.cCart).Value)
            sheetVal = REVO_Core.SafeD(fm.ws.Cells(r, fm.cRelToDate).Value)
            logVal = ReleasedToDateFromLog(WO, cart)
            If Abs(sheetVal - logVal) > 0.001 Then
                fm.ws.Cells(r, fm.cRelToDate).Value = logVal
                fm.ws.Cells(r, fm.cRemain).Value = _
                    REVO_Core.SafeD(fm.ws.Cells(r, fm.cQty).Value) - logVal
                If logVal <= 0 Then
                    fm.ws.Cells(r, fm.cStatus).Value = ""
                ElseIf logVal >= REVO_Core.SafeD(fm.ws.Cells(r, fm.cQty).Value) Then
                    fm.ws.Cells(r, fm.cStatus).Value = "RELEASED"
                Else
                    fm.ws.Cells(r, fm.cStatus).Value = "PARTIAL"
                End If
                nFix = nFix + 1
            End If
        End If
    Next r
    REVO_Core.PopAppState

    REVO_Core.Audit "REVO_Release", "RECONCILE", SH_FLOOR, nFix & " rows corrected from Release Log"
    MsgBox nFix & " row(s) corrected from Release Log.", vbInformation, "Reconcile"
End Sub

'==============================================================================
' RESET A CART  -  replaces the dead UndoReleaseForCart button
'
' Clears the stranded state the old macro created: text sitting in the
' checkbox column, or a released-to-date that no longer matches the log.
'==============================================================================
Public Sub REVO_ResetCart()
    Dim fm As FloorMap
    Dim cart As String
    Dim r As Long, n As Long
    Dim found As String
    Dim WO As String, logVal As Double, qty As Double

    fm = MapFloorSheet()
    If Not fm.ok Then
        MsgBox "Cannot read REVO Floor." & vbCrLf & vbCrLf & fm.problem, vbExclamation
        Exit Sub
    End If

    cart = InputBox("Cart number to reset:" & vbCrLf & vbCrLf & _
                    "This clears the Release? checkbox and the status text, and resets " & _
                    "Released To Date from Release Log. It does not delete any log rows.", _
                    "Reset Cart")
    If Len(Trim$(cart)) = 0 Then Exit Sub

    REVO_Core.PushAppState
    For r = fm.firstRow To fm.lastRow
        If UCase$(REVO_Core.SafeS(fm.ws.Cells(r, fm.cCart).Value)) = UCase$(Trim$(cart)) Then
            WO = REVO_Core.SafeS(fm.ws.Cells(r, fm.cWO).Value)
            qty = REVO_Core.SafeD(fm.ws.Cells(r, fm.cQty).Value)
            logVal = ReleasedToDateFromLog(WO, cart)

            fm.ws.Cells(r, fm.cRelease).Value = False
            fm.ws.Cells(r, fm.cRelToDate).Value = logVal
            fm.ws.Cells(r, fm.cRemain).Value = qty - logVal
            If logVal <= 0 Then
                fm.ws.Cells(r, fm.cStatus).Value = ""
            ElseIf logVal >= qty Then
                fm.ws.Cells(r, fm.cStatus).Value = "RELEASED"
            Else
                fm.ws.Cells(r, fm.cStatus).Value = "PARTIAL"
            End If

            n = n + 1
            found = found & "  Row " & r & " (" & WO & "): released to date now " & logVal & _
                    " of " & qty & vbCrLf
        End If
    Next r
    REVO_Core.PopAppState

    If n = 0 Then
        MsgBox "No row on REVO Floor carries cart " & cart & ".", vbExclamation, "Reset Cart"
    Else
        REVO_Core.Audit "REVO_Release", "RESET CART", SH_FLOOR, "Cart " & cart & ", " & n & " row(s)"
        MsgBox n & " row(s) reset for cart " & cart & ":" & vbCrLf & vbCrLf & found & vbCrLf & _
               "The Release? checkbox is clear and clickable again.", vbInformation, "Reset Cart"
    End If
End Sub

'==============================================================================
' LOG WRITERS
'==============================================================================
Private Sub WriteQE(ByVal disp As String, ByVal qty As Double, ByVal d As Date, _
                    ByVal WO As String, ByVal sku As String, ByVal cart As String, _
                    ByVal releaseID As String, ByVal defect As String, _
                    ByVal location As String, ByVal op As String, _
                    ByVal root As String, ByVal action As String, ByVal notes As String)
    Dim ev As QualityEvent
    ev.EventDate = d
    ev.source = "RELEASE"
    ev.WO = WO
    ev.sku = sku
    ev.cart = cart
    ev.Disposition = disp
    ev.Qty = qty
    ev.DefectType = defect
    ev.Location = location
    ev.DetectedOp = op
    ev.RootCause = root
    ev.Action = action
    ev.MultiDefect = False
    ev.Notes = notes
    ev.ReleaseID = releaseID
    REVO_Quality.WriteQualityEvent ev
End Sub

Private Function RejectSummaryText(ByVal frm As frmReleaseDetails) As String
    Dim s As String
    s = frm.RejectDefect
    If Len(frm.RejectLocation) > 0 And frm.RejectLocation <> "Not Specified" Then
        s = s & " (" & frm.RejectLocation & ")"
    End If
    If Len(frm.RejectRootCause) > 0 Then s = s & " - " & frm.RejectRootCause
    If Len(frm.RejectNotes) > 0 Then s = s & "; " & frm.RejectNotes
    RejectSummaryText = s
End Function

Public Sub EnsureShipmentsHeaders()
    Dim ws As Worksheet
    Set ws = REVO_Core.GetOrCreateSheet(SH_SHIPMENTS)
    If Len(REVO_Core.SafeS(ws.Cells(1, 1).Value)) = 0 Then
        ws.Range("A1:D1").Value = Array("Date", "SKU", "Qty", "Cart")
        ws.Range("A1:D1").Font.Bold = True
    End If
End Sub

Public Sub EnsureReleaseLogHeaders()
    Dim ws As Worksheet
    Set ws = REVO_Core.GetOrCreateSheet(SH_RELEASE_LOG)
    If Len(REVO_Core.SafeS(ws.Cells(1, 1).Value)) = 0 Then
        ws.Range("A1:K1").Value = Array("Date", "SKU", "Cart", "Qty Released", "Work Order", _
                                        "Qty Rework", "Qty B Grade", "Qty Reject", "Recalled", _
                                        "Cart Qty Released", "Cart Remaining")
        ws.Range("A1:K1").Font.Bold = True
    End If
End Sub

Private Sub AppendShipments(ByVal d As Date, ByVal sku As String, ByVal qty As Double, ByVal cart As String)
    Dim ws As Worksheet, r As Long
    Set ws = REVO_Core.GetSheet(SH_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = d
    ws.Cells(r, 1).NumberFormat = "m/d/yyyy"
    ws.Cells(r, 2).Value = sku
    ws.Cells(r, 3).Value = qty
    ws.Cells(r, 4).Value = cart
End Sub

' NOTE: Release Log carries a pivot table from column L rightwards. Rows are
' appended using column A only, which keeps clear of it.
Private Sub AppendReleaseLog(ByVal d As Date, ByVal sku As String, ByVal cart As String, _
                             ByVal toInv As Double, ByVal WO As String, _
                             ByVal reworkQty As Double, ByVal bGradeQty As Double, _
                             ByVal rejectQty As Double, ByVal cartQtyRel As Double, _
                             ByVal cartRemain As Double)
    Dim ws As Worksheet, r As Long
    Set ws = REVO_Core.GetSheet(SH_RELEASE_LOG)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = d
    ws.Cells(r, 1).NumberFormat = "m/d/yyyy"
    ws.Cells(r, 2).Value = sku
    ws.Cells(r, 3).Value = cart
    ws.Cells(r, 4).Value = toInv
    ws.Cells(r, 5).Value = WO
    ws.Cells(r, 6).Value = reworkQty
    ws.Cells(r, 7).Value = bGradeQty
    ws.Cells(r, 8).Value = rejectQty
    ws.Cells(r, 9).Value = False
    ws.Cells(r, 10).Value = cartQtyRel
    ws.Cells(r, 11).Value = cartRemain
End Sub

Private Sub AppendReject(ByVal d As Date, ByVal WO As String, ByVal sku As String, _
                         ByVal cart As String, ByVal qty As Double, ByVal reason As String)
    Dim ws As Worksheet, r As Long
    Set ws = REVO_Core.GetOrCreateSheet(SH_REJECT_LOG)
    If Len(REVO_Core.SafeS(ws.Cells(1, 1).Value)) = 0 Then
        ws.Range("A1:F1").Value = Array("Date", "SKU", "WO#", "Cart", "Quantity", "Description")
        ws.Range("A1:F1").Font.Bold = True
    End If
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = d
    ws.Cells(r, 1).NumberFormat = "m/d/yyyy"
    ws.Cells(r, 2).Value = sku
    ws.Cells(r, 3).Value = WO
    ws.Cells(r, 4).Value = cart
    ws.Cells(r, 5).Value = qty
    ws.Cells(r, 6).Value = reason
End Sub

Private Sub AppendRework(ByVal d As Date, ByVal WO As String, ByVal sku As String, _
                         ByVal cart As String, ByVal qty As Double)
    Dim ws As Worksheet, r As Long
    Set ws = REVO_Core.GetOrCreateSheet(SH_REWORK)
    If Len(REVO_Core.SafeS(ws.Cells(1, 1).Value)) = 0 Then
        ws.Range("A1:G1").Value = Array("Date", "WO", "SKU", "Cart", "Qty", "Status", "Previous Release Status")
        ws.Range("A1:G1").Font.Bold = True
    End If
    r = ws.Cells(ws.Rows.Count, 3).End(xlUp).row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = d
    ws.Cells(r, 1).NumberFormat = "m/d/yyyy"
    ws.Cells(r, 2).Value = WO
    ws.Cells(r, 3).Value = sku
    ws.Cells(r, 4).Value = cart
    ws.Cells(r, 5).Value = qty
    ws.Cells(r, 6).Value = "In Progress"
End Sub

Private Sub UpdateBGrade(ByVal sku As String, ByVal qty As Double)
    ' B-grade inventory lives in the legacy REVO_Ops routine. Call it when
    ' present; never fail the release because it is not.
    On Error Resume Next
    Application.Run "UpdateBGradeInventory", sku, qty
    On Error GoTo 0
End Sub

Private Sub Accumulate(ByVal d As Object, ByVal k As String, ByVal v As Double)
    If Len(Trim$(k)) = 0 Then Exit Sub
    If d.exists(k) Then
        d(k) = d(k) + v
    Else
        d.Add k, v
    End If
End Sub

'==============================================================================
' MAIL
'==============================================================================
Private Sub SendReleaseMail(ByVal summary As Object, ByVal detail As String, ByVal d As Date)
    Dim ol As Object, mi As Object
    Dim body As String, k As Variant
    Dim total As Double

    body = "REVO Production Release - " & Format$(d, "m/d/yyyy") & vbCrLf & vbCrLf & _
           "RELEASED TO INVENTORY" & vbCrLf
    For Each k In summary.keys
        body = body & "  " & k & ": " & summary(k) & vbCrLf
        total = total + summary(k)
    Next k
    body = body & "  ----" & vbCrLf & "  Total: " & total & vbCrLf & vbCrLf & _
           "DETAIL" & vbCrLf & detail & vbCrLf & _
           "Quality dispositions are coded in the Quality Log and summarised on the Quality Dashboard."

    On Error GoTo NoMail
    Set ol = CreateObject("Outlook.Application")
    Set mi = ol.CreateItem(0)
    mi.To = MAIL_TO
    mi.CC = MAIL_CC
    mi.Subject = "REVO Production Release Notification - " & Format$(d, "m/d/yyyy")
    mi.body = body
    mi.Display
    Exit Sub
NoMail:
    REVO_Core.Audit "REVO_Release", "MAIL FAILED", "Outlook", Err.Description
End Sub
