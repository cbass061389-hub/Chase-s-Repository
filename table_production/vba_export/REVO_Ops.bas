Option Explicit

'======================== CONFIG (new layout) ========================
Private Const SRC_SHEET As String = "REVO Floor"
Private Const DEST_SHIP As String = "Shipments"
Private Const LOG_RELEASE As String = "Release Log"
Private Const LOG_REJECT As String = "Reject Log"
Private Const LOG_REWORK As String = "Rework Tracker"
Private Const SH_REVO_INV As String = "REVO_INV"
Private Const SH_REVO_FCST As String = "REVO_SALES_FRCST"
Private Const SH_EST As String = "EST_REVO_INV"
Private Const SH_BOM As String = "BOM"
Private Const SH_PROD As String = "Production Plan"
Private Const SH_BIN As String = "Bin_Inv"
Private Const SH_OUT As String = "REVO Component Forecast"
Const SH_PO_LOCAL As String = "On_PO"

Private Const EMAIL_FONT As String = "Aptos, Calibri, Arial, sans-serif"



Private Const COL_WO As Long = 2      ' B: WO#
Private Const COL_SKU As Long = 3     ' C: SKU
Private Const COL_CART As Long = 4    ' D: CART
Private Const COL_QTY As Long = 5     ' E: FORECAST (qty)
Private Const COL_DATE As Long = 6    ' F: DATE (not required for flow below)
Private Const COL_RELEASE As Long = 21 ' Release flag/marker column
Private Const HEADER_ROW As Long = 2
Private Const DATA_FIRST_ROW As Long = 3

Private Const SRC_COL_WO As Long = 2        ' B (not used by this routine, kept for ref)
Private Const SRC_COL_SKU As Long = 3       ' C
Private Const SRC_COL_CART As Long = 4      ' D (not used here)
Private Const SRC_COL_FORECAST As Long = 5  ' E
Private Const SRC_COL_DATE As Long = 6      ' F
Private Const SRC_FIRST_DATA_ROW As Long = 2

'==== DESTINATION (tweak if your sheet differs) ====
Private Const SH_DST As String = "Production Plan"
Private Const DST_COL_SKU As Long = 1         ' A
Private Const DST_ROW_DATES As Long = 2       ' dates row
Private Const DST_FIRST_DATA_ROW As Long = 5  ' first SKU data row
Private Const DST_ROW_TOTAL As Long = 4       ' weekly total row
Private Const DST_ROW_CAP As Long = 3
Private Const DST_ROW_PROD_WEEK As Long = 1
' === REVO_SALES_FRCST layout (normalized) ===
' A: SKU
' F: Forecast Mnth (date)
' G: Qty
Private Const FC_COL_SKU As Long = 1   ' A
Private Const FC_COL_DATE As Long = 6  ' F
Private Const FC_COL_QTY As Long = 7   ' G

'===== REVO Component Forecast config (no overlap) =====
Private Const RCF_SH_BOM As String = "BOM"
Private Const RCF_SH_PROD As String = "Production Plan"
Private Const RCF_SH_BIN As String = "Bin_Inv"
Private Const RCF_SH_OUT As String = "REVO Component Forecast"

' OUT sheet layout
Private Const RCF_HDR_ROW As Long = 4
Private Const RCF_FIRST_ROW As Long = 5
Private Const RCF_COL_SKU As Long = 1
Private Const RCF_COL_SS_LEVEL As Long = 2
Private Const RCF_COL_STOCK_INV As Long = 3
Private Const RCF_COL_WIP As Long = 4
Private Const RCF_COL_SS_TO_PROD As Long = 5
Private Const RCF_COL_FLOW_NEED As Long = 6
Private Const RCF_FIRST_WEEK_COL As Long = 7

' BOM layout (SKU column only for Product_Type / Critical)
Private Const RCF_BOM_COL_SKU As Long = 1
Private Const RCF_BOM_COL_COMP As Long = 2
Private Const RCF_BOM_COL_QTY As Long = 3
Private Const RCF_BOM_COL_CRIT As Long = 4      ' Critical_Status
Private Const RCF_BOM_COL_PT As Long = 6        ' Product_Type (PREASSEMBLY / FINISHED SHAFT)

' Production Forecast layout
Private Const RCF_PF_ROW_DATES As Long = 2      ' date headers
Private Const RCF_PF_COL_SKU As Long = 1        ' SKU col
Private Const RCF_PF_FIRST_DATA_ROW As Long = 5 ' first SKU row
Private Const PF_ROW_DATES As Long = 2
Private Const PF_COL_SKU As Long = 1
Private Const PF_FIRST_DATA_ROW As Long = 5

' BOM columns
Private Const BOM_COL_SKU As Long = 1    ' parent
Private Const BOM_COL_COMP As Long = 2   ' child
Private Const BOM_COL_QTY As Long = 3
Private Const BOM_COL_CRIT As Long = 4   ' Critical_Status on SKU
Private Const BOM_COL_PT As Long = 6     ' Product_Type on SKU (FINISHED SHAFT / PREASSEMBLY / etc)
Private Const RCF_COL_TOTAL_AVAIL As Long = 5
Private Const RCF_COL_P1 As Long = 6
Private Const RCF_COL_P2 As Long = 7
Private Const RCF_COL_P3 As Long = 8
Private Const RCF_COL_PRIORITY As Long = 9

'==== ON_PO LAYOUT ====
Const PO_COL_ITEM As Long = 1
Const PO_COL_DATE As Long = 2
Const PO_COL_LOC As Long = 3
Const PO_COL_QTY As Long = 4

Const ENABLE_HARD_CAP As Boolean = False   ' True = enforce constraints, False = flag only









'============================== MAIN ================================
Public Sub ReleaseToShippingAndReceiving()
    Dim wsSource As Worksheet, wsShip As Worksheet, wsRelLog As Worksheet, wsRej As Worksheet, wsRew As Worksheet
    Dim r As Long, lastRow As Long, firstDataRow As Long
    Dim currentDate As String
    Dim WO As String, sku As String, Cart As String
    Dim qty As Double, releasedToInv As Double, reworkQty As Double, bGradeQty As Double, rejectQty As Double
    Dim rejectReason As String
    Dim releaseSummary As Object ' Scripting.Dictionary
    Dim detailedSummary As String
    Dim frm As Object
    
    On Error GoTo FailFast
    
    Set wsSource = EnsureSheetExists(SRC_SHEET)
    Set wsShip = EnsureSheetExists(DEST_SHIP)
    Set wsRelLog = EnsureSheetExists(LOG_RELEASE)
    Set wsRej = EnsureSheetExists(LOG_REJECT)
    Set wsRew = EnsureSheetExists(LOG_REWORK)
    
    EnsureShipmentsHeaders wsShip
    EnsureReleaseLogHeaders wsRelLog
    EnsureRejectLogHeaders wsRej
    EnsureReworkHeaders wsRew
    
    Set releaseSummary = CreateObject("Scripting.Dictionary")
    detailedSummary = vbNullString
    currentDate = Format(Date, "m/d/yyyy")
    
    firstDataRow = 2
    lastRow = wsSource.Cells(wsSource.Rows.Count, COL_RELEASE).End(xlUp).row
    If lastRow < firstDataRow Then GoTo SendMailAndExit
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    For r = firstDataRow To lastRow
        If Boolish(wsSource.Cells(r, COL_RELEASE).Value) Then
            WO = NzStr(wsSource.Cells(r, COL_WO).Value)
            sku = NzStr(wsSource.Cells(r, COL_SKU).Value)
            Cart = NzStr(wsSource.Cells(r, COL_CART).Value)
            qty = SafeDbl(wsSource.Cells(r, COL_QTY).Value)
            
            If Len(sku) = 0 Or qty <= 0 Then
                ' Mark processed to avoid repeat passes on a bad line
                wsSource.Cells(r, COL_RELEASE).Value = "RELEASED"
                GoTo NextRow
            End If
            
            ' Defaults: all to released
            releasedToInv = qty: reworkQty = 0: bGradeQty = 0: rejectQty = 0: rejectReason = ""
            
            ' Try userform (late bound; safe if not present)
            On Error Resume Next
            Set frm = VBA.UserForms.Add("frmReleaseDetails")
            On Error GoTo 0
            
            If Not frm Is Nothing Then
                On Error Resume Next
                CallByName frm, "sku", VbLet, sku
                CallByName frm, "Cart", VbLet, Cart
                CallByName frm, "TotalQty", VbLet, qty
                If HasControl(frm, "lblSKU") Then frm.Controls("lblSKU").Caption = sku
                If HasControl(frm, "lblCart") Then frm.Controls("lblCart").Caption = Cart
                If HasControl(frm, "lblQty") Then frm.Controls("lblQty").Caption = CStr(qty)
                On Error GoTo FailFast
                
                frm.Show
                
                On Error Resume Next
                If CBool(CallByName(frm, "UserCancelled", VbGet)) Then
                    Set frm = Nothing
                    GoTo NextRow ' leave release flag as-is so user can revisit
                End If
                
                releasedToInv = GetIfExistsDbl(frm, "ReleasedToInventory", qty)
                reworkQty = GetIfExistsDbl(frm, "Rework", 0)
                bGradeQty = GetIfExistsDbl(frm, "BGrade", 0)
                rejectQty = GetIfExistsDbl(frm, "Reject", 0)
                rejectReason = GetIfExistsStr(frm, "txtRejectReason", "")
                On Error GoTo FailFast
                Set frm = Nothing
            End If
            
            ' Coherence guard (cap totals at qty)
            If releasedToInv + reworkQty + bGradeQty + rejectQty > qty Then
                Dim over As Double
                over = releasedToInv + reworkQty + bGradeQty + rejectQty - qty
                If over > 0 Then releasedToInv = Application.Max(0, releasedToInv - over)
            End If
            
            ' === Writes ===
            If releasedToInv > 0 Then
                AppendShipments wsShip, currentDate, sku, releasedToInv, Cart
                Accumulate releaseSummary, sku, releasedToInv
            End If
            
            If reworkQty > 0 Then LogRework WO, sku, Cart, reworkQty, currentDate
            If bGradeQty > 0 Then UpdateBGradeInventory sku, bGradeQty
            If rejectQty > 0 Then LogReject WO, sku, Cart, rejectQty, currentDate, rejectReason
            
            LogReleaseToLog sku, Cart, releasedToInv, WO, reworkQty, bGradeQty, rejectQty, currentDate
            
            ' === Mark processed (no deletion per your rule) ===
            wsSource.Cells(r, COL_RELEASE).Value = "RELEASED"
            
            ' === Detail line for email ===
            detailedSummary = detailedSummary & _
                "Cart: " & Cart & vbCrLf & _
                "SKU: " & sku & vbCrLf & _
                "Qty Released: " & releasedToInv & vbCrLf & _
                "Work Order: " & WO & vbCrLf & _
                "B Grade: " & bGradeQty & vbCrLf & _
                "Rework: " & reworkQty & vbCrLf & _
                "Reject: " & rejectQty & vbCrLf & _
                IIf(rejectQty > 0 And Len(rejectReason) > 0, "Reject Description: " & rejectReason & vbCrLf, "") & _
                vbCrLf
        End If
NextRow:
    Next r
    
    RunWeeklyREVOUpdate_Safe
    RefreshRejectDashboard_Full

SendMailAndExit:
    If Not releaseSummary Is Nothing Then
        If releaseSummary.Count > 0 Then
            Dim body As String
            body = BuildEmailBody(releaseSummary, detailedSummary, currentDate)
            SendOutlookMail _
                "mbyron@predatorgroup.com", _
                "gkeys@predatorgroup.com; jausler@predatorgroup.com; cbass@predatorgroup.com;osantiago@predatorgroup.com; cfournier@predatorgroup.com; abirmingham@predatorgroup.com; rorr@predatorgroup.com", _
                "REVO Production Release Notification - " & currentDate, _
                body
        End If
    End If
    
CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

FailFast:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "ReleaseToShippingAndReceiving failed: " & Err.Description, vbExclamation
End Sub

Public Sub WeeklyReportUpdate()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    RunWeeklyREVOUpdate_Safe
    RefreshRejectDashboard_Full
    
    Sheets("Reject_Dashboard").Select
    Range("B4:V18").Select
    Selection.Copy
    Sheets("Weekly Reporting").Select
    Range("A20").Select
    ActiveSheet.Paste
    
CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
End Sub

Public Sub ReleaseReworksToInventory()
    Const SH_QUEUE As String = "Rework Tracker"
    Const SH_DEST As String = "Shipments"
    Const SH_LOG  As String = "Release Log"

    ' === Column layout for REVO Floor / Rework Tracker ===
    Const COL_WO As Long = 2          ' B
    Const COL_SKU As Long = 3         ' C
    Const COL_CART As Long = 4        ' D
    Const COL_QTY As Long = 5         ' E
    Const COL_RELEASE As String = "H" ' or your actual release flag column

    Dim wsQueue As Worksheet, wsDest As Worksheet, wsLog As Worksheet
    Dim lastRow As Long, i As Long, destRow As Long
    Dim currentDate As String
    Dim releaseSummary As Object
    Dim detailedSummary As String

    Dim sku As String, Cart As String, WO As String
    Dim rowQty As Long, bGradeQty As Long, rejectQty As Long, rejectReason As String
    Dim cartRework As String, key As String

    Set wsQueue = ThisWorkbook.Sheets(SH_QUEUE)
    Set wsDest = ThisWorkbook.Sheets(SH_DEST)
    Set wsLog = EnsureSheetExists(SH_LOG)

    Set releaseSummary = CreateObject("Scripting.Dictionary")
    currentDate = Format(Date, "m/d/yyyy")

    Application.ScreenUpdating = False
    detailedSummary = vbNullString

    lastRow = wsQueue.Cells(wsQueue.Rows.Count, COL_SKU).End(xlUp).row
    If lastRow < 2 Then GoTo CleanExit

    ' === Loop backward so row deletions don't break iteration ===
    For i = lastRow To 2 Step -1
        If Boolish(wsQueue.Cells(i, COL_RELEASE).Value) Then
            WO = SafeText(wsQueue.Cells(i, COL_WO).Value)
            sku = SafeText(wsQueue.Cells(i, COL_SKU).Value)
            Cart = SafeText(wsQueue.Cells(i, COL_CART).Value)
            rowQty = SafeLng(wsQueue.Cells(i, COL_QTY).Value)

            ' --- Launch the release dialog ---
            Dim frm As frmReworkRelease
            Set frm = New frmReworkRelease
            frm.Prime sku, Cart, rowQty
            frm.Show

            If frm.UserCancelled = False Then
                bGradeQty = frm.BGrade
                rejectQty = frm.Reject
                rejectReason = frm.txtRejectReason.Value
                cartRework = Cart & " REWORK"

                ' === Write Released to Shipments ===
                If frm.ReleasedToInventory > 0 Then
                    destRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).row + 1
                    wsDest.Cells(destRow, "A").Value = currentDate
                    wsDest.Cells(destRow, "B").Value = sku
                    wsDest.Cells(destRow, "C").Value = frm.ReleasedToInventory
                    wsDest.Cells(destRow, "D").Value = cartRework
                End If

                ' === B Grade ===
                If bGradeQty > 0 Then
                    UpdateBGradeInventory sku, bGradeQty
                End If

                ' === Reject ===
                If rejectQty > 0 Then
                    LogReject WO, sku, cartRework, rejectQty, currentDate, rejectReason
                End If

                ' === Log release ===
                destRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
                wsLog.Cells(destRow, 1).Value = currentDate
                wsLog.Cells(destRow, 2).Value = sku
                wsLog.Cells(destRow, 3).Value = cartRework
                wsLog.Cells(destRow, 4).Value = frm.ReleasedToInventory
                wsLog.Cells(destRow, 5).Value = WO
                wsLog.Cells(destRow, 6).Value = 0
                wsLog.Cells(destRow, 7).Value = bGradeQty
                wsLog.Cells(destRow, 8).Value = rejectQty
                wsLog.Cells(destRow, 9).Value = False

                ' === Summary for email ===
                key = sku
                If Not releaseSummary.exists(key) Then
                    releaseSummary.Add key, CLng(frm.ReleasedToInventory)
                Else
                    releaseSummary(key) = releaseSummary(key) + CLng(frm.ReleasedToInventory)
                End If

                detailedSummary = detailedSummary & _
                    "Cart: " & cartRework & vbCrLf & _
                    "SKU: " & sku & vbCrLf & _
                    "Qty Released: " & frm.ReleasedToInventory & vbCrLf & _
                    "Work Order: " & WO & vbCrLf & _
                    "B Grade: " & bGradeQty & vbCrLf & _
                    "Reject: " & rejectQty & vbCrLf
                If rejectQty > 0 Then
                    detailedSummary = detailedSummary & "Reject Description: " & rejectReason & vbCrLf
                End If
                detailedSummary = detailedSummary & vbCrLf

                ' === Adjust Rework Tracker row (partial vs full) ===
                Dim newQty As Long
                newQty = rowQty - frm.RequestedQty

                If frm.RequestedQty >= rowQty Then
                    ' --- full release ---
                    wsQueue.Rows(i).Delete
                Else
                    ' --- partial release ---
                    wsQueue.Cells(i, COL_QTY).Value = newQty
                    wsQueue.Cells(i, COL_RELEASE).Value = frm.RequestedQty & " RELEASED"
                End If
            End If
        End If
    Next i
    
    RunWeeklyREVOUpdate_Safe
    RefreshRejectDashboard_Full

    ' === Email notification ===
    If releaseSummary.Count > 0 Then
        Dim emailBody As String, SkuKey As Variant
        emailBody = "Hello Team," & vbCrLf & vbCrLf & _
                    "The following rework shafts are completed and released for inventory:" & vbCrLf & vbCrLf

        For Each SkuKey In releaseSummary.keys
            emailBody = emailBody & "SKU: " & CStr(SkuKey) & vbTab & "Qty: " & CStr(releaseSummary(SkuKey)) & vbCrLf
        Next SkuKey

        emailBody = emailBody & vbCrLf & "Details:" & vbCrLf & vbCrLf & _
                    detailedSummary & vbCrLf & _
                    "Thank you and have a great day!" & vbCrLf & "REVO Team"

        Dim olApp As Object, olMail As Object
        On Error Resume Next
        Set olApp = GetObject(, "Outlook.Application")
        If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
        On Error GoTo 0

        If Not olApp Is Nothing Then
            Set olMail = olApp.CreateItem(0)
            With olMail
                .To = "mbyron@predatorgroup.com"
                .CC = "gkeys@predatorgroup.com; cbass@predatorgroup.com; jausler@predatorgroup.com; osantiago@predatorgroup.com; cfournier@predatorgroup.com; abirmingham@predatorgroup.com; rorr@predatorgroup.com"
                .Subject = "REVO Rework Release to Inventory - " & currentDate
                .body = emailBody
                .Display
            End With
            Set olMail = Nothing
        End If
        Set olApp = Nothing
    End If

CleanExit:
    Application.ScreenUpdating = True
    MsgBox "Rework release complete.", vbInformation
End Sub


'=========================== LOG WRITERS ===========================
Private Sub LogReleaseToLog( _
    ByVal sku As String, ByVal Cart As String, ByVal releasedQty As Double, _
    ByVal WO As String, ByVal reworkQty As Double, ByVal bGradeQty As Double, _
    ByVal rejectQty As Double, ByVal ReleaseDate As String)

    Dim wsLog As Worksheet, NextRow As Long
    Set wsLog = EnsureSheetExists(LOG_RELEASE)
    EnsureReleaseLogHeaders wsLog
    
    NextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
    With wsLog
        .Cells(NextRow, "A").Value = ReleaseDate
        .Cells(NextRow, "B").Value = sku
        .Cells(NextRow, "C").Value = Cart
        .Cells(NextRow, "D").Value = releasedQty
        .Cells(NextRow, "E").Value = WO
        .Cells(NextRow, "F").Value = reworkQty
        .Cells(NextRow, "G").Value = bGradeQty
        .Cells(NextRow, "H").Value = rejectQty
        .Cells(NextRow, "I").Value = False ' recall/verify checkbox
    End With
End Sub

Private Sub LogRework( _
    ByVal WO As String, ByVal sku As String, ByVal Cart As String, _
    ByVal qty As Double, Optional ByVal theDate As String = "")

    Dim ws As Worksheet, NextRow As Long
    Set ws = EnsureSheetExists(LOG_REWORK)
    EnsureReworkHeaders ws
    If Len(theDate) = 0 Then theDate = Format(Date, "m/d/yyyy")
    
    NextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    With ws
        .Cells(NextRow, "A").Value = theDate
        .Cells(NextRow, "B").Value = WO
        .Cells(NextRow, "C").Value = sku
        .Cells(NextRow, "D").Value = Cart
        .Cells(NextRow, "E").Value = qty
        .Cells(NextRow, "F").Value = "In Progress"
    End With
End Sub

Private Sub LogReject( _
    ByVal WO As String, ByVal sku As String, ByVal Cart As String, _
    ByVal qty As Double, ByVal theDate As String, ByVal reason As String)

    Dim ws As Worksheet, r As Long
    Set ws = EnsureSheetExists(LOG_REJECT)
    EnsureRejectLogHeaders ws
    
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    With ws
        .Cells(r, 1).Value = theDate     ' Date
        .Cells(r, 2).Value = sku          ' WO
        .Cells(r, 3).Value = WO         ' SKU
        .Cells(r, 4).Value = Cart        ' Cart
        .Cells(r, 5).Value = qty         ' RejectQty
        .Cells(r, 6).Value = reason      ' Reason
        
    End With
End Sub

Private Sub UpdateBGradeInventory(ByVal sku As String, ByVal qty As Double)
    ' TODO: hook this into your B-Grade inventory sheet/system as needed.
    ' Provided here as a stub to keep flow intact.
End Sub

'========================= IO / EMAIL HELPERS ======================
Private Sub AppendShipments(ws As Worksheet, ByVal dt As String, ByVal sku As String, ByVal qty As Double, ByVal Cart As String)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    ws.Cells(r, "A").Value = dt
    ws.Cells(r, "B").Value = sku
    ws.Cells(r, "C").Value = qty
    ws.Cells(r, "D").Value = Cart
End Sub

Private Function BuildEmailBody(ByVal summary As Object, ByVal details As String, ByVal dt As String) As String
    Dim k As Variant, s As String
    s = "Hello Team," & vbCrLf & vbCrLf & _
        "The following shafts are ready to be transferred for sale:" & vbCrLf & vbCrLf
    For Each k In summary.keys
        s = s & "SKU: " & CStr(k) & vbTab & "Qty: " & summary(k) & vbCrLf
    Next k
    s = s & vbCrLf & "Here is a detailed view of this release:" & vbCrLf & vbCrLf & _
        details & vbCrLf & _
        "Thank you and have a great day!" & vbCrLf & "REVO Team"
    BuildEmailBody = s
End Function

Private Sub SendOutlookMail(ByVal toList As String, ByVal ccList As String, ByVal subj As String, ByVal body As String)
    Dim olApp As Object, olMail As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then Exit Sub
    
    Set olMail = olApp.CreateItem(0)
    With olMail
        .To = toList
        .CC = ccList
        .Subject = subj
        .body = body
        .Display
    End With
End Sub

'========================== ENSURE HELPERS =========================
Private Function EnsureSheetExists(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheetExists = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheetExists Is Nothing Then
        Set EnsureSheetExists = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheetExists.name = sheetName
    End If
End Function

Private Sub EnsureShipmentsHeaders(ws As Worksheet)
    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        ws.Range("A1:D1").Value = Array("Date", "SKU", "Qty", "Cart")
    End If
End Sub

Private Sub EnsureReleaseLogHeaders(ws As Worksheet)
    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        ws.Range("A1:I1").Value = Array( _
            "Date", "SKU", "Cart", "ReleasedQty", "WO", "Rework", "B Grade", "Reject", "Verified")
    End If
End Sub

Private Sub EnsureRejectLogHeaders(ws As Worksheet)
    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        ws.Range("A1:G1").Value = Array( _
            "Date", "WO", "SKU", "Cart", "RejectQty", "Reason", "User")
    End If
End Sub

Private Sub EnsureReworkHeaders(ws As Worksheet)
    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        ws.Range("A1:F1").Value = Array( _
            "Date", "WO", "SKU", "Cart", "Qty", "Status")
    End If
End Sub

'========================== SAFE/UTILS =============================
Private Function NzStr(v As Variant, Optional ByVal dflt As String = "") As String
    If IsError(v) Or IsNull(v) Or v = "" Then NzStr = dflt Else NzStr = CStr(v)
End Function

Private Function SafeDbl(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then
        SafeDbl = 0#
    Else
        SafeDbl = CDbl(v)
    End If
End Function

Private Function Boolish(v As Variant) As Boolean
    On Error Resume Next
    Select Case True
        Case VarType(v) = vbBoolean: Boolish = CBool(v)
        Case IsNumeric(v): Boolish = (CDbl(v) <> 0)
        Case TypeName(v) = "String": Boolish = (UCase$(Trim$(CStr(v))) Like "TRUE" Or Trim$(CStr(v)) = "1" Or UCase$(Trim$(CStr(v))) = "X")
        Case Else: Boolish = False
    End Select
End Function

' ---- userform shims (safe gets, control check) ----
Private Function GetIfExistsDbl(ByVal frm As Object, ByVal propName As String, Optional ByVal dflt As Double = 0#) As Double
    On Error Resume Next
    GetIfExistsDbl = CDbl(CallByName(frm, propName, VbGet))
    If Err.Number <> 0 Then GetIfExistsDbl = dflt
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetIfExistsStr(ByVal frm As Object, ByVal propName As String, Optional ByVal dflt As String = "") As String
    On Error Resume Next
    GetIfExistsStr = CStr(CallByName(frm, propName, VbGet))
    If Err.Number <> 0 Then GetIfExistsStr = dflt
    Err.Clear
    On Error GoTo 0
End Function

Private Function HasControl(ByVal frm As Object, ByVal ctrlName As String) As Boolean
    On Error Resume Next
    HasControl = Not frm.Controls(ctrlName) Is Nothing
    If Err.Number <> 0 Then HasControl = False
    Err.Clear
    On Error GoTo 0
End Function

'==================== DICTIONARY ADD/ACCUMULATE HELPER ====================
Private Sub Accumulate(ByRef dict As Object, ByVal key As String, ByVal addVal As Double)
    If Len(key) = 0 Then Exit Sub
    If Not dict.exists(key) Then
        dict.Add key, addVal
    Else
        dict(key) = dict(key) + addVal
    End If
End Sub




'==== tiny helpers ====
Private Function SafeText(v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        SafeText = ""
    Else
        SafeText = Trim$(CStr(v))
    End If
End Function


Public Sub UpdateProductionForecast_ByWeeks()
    ' ===== Source: REVO Floor (new layout) =====
    Const SRC_SHEET As String = "REVO Floor"
    Const SRC_COL_SKU As Long = 3           ' C
    Const SRC_COL_FORECAST As Long = 6      ' E
    Const SRC_COL_DATE As Long = 7          ' F
    Const SRC_FIRST_DATA_ROW As Long = 12


    Dim scrn&, evt&, calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:     Application.EnableEvents = False
    calc = Application.Calculation:     Application.Calculation = xlCalculationManual

    Dim stage As String
    On Error GoTo CleanFail

    Dim wsS As Worksheet, wsD As Worksheet
    Set wsS = ThisWorkbook.Worksheets(SRC_SHEET)
    Set wsD = ThisWorkbook.Worksheets(SH_DST)

    ' ---- DEST bounds & date header map ----
    stage = "dest bounds"
    Dim lastDstRow As Long, lastDstCol As Long
    lastDstRow = wsD.Cells(wsD.Rows.Count, DST_COL_SKU).End(xlUp).row
    lastDstCol = wsD.Cells(DST_ROW_DATES, wsD.Columns.Count).End(xlToLeft).Column
    If lastDstCol < 1 Then GoTo CleanExit

    stage = "dest date map"
    Dim colByDate As Object: Set colByDate = CreateObject("Scripting.Dictionary")
    Dim c As Long, dv As Variant, dk As Long
    For c = 1 To lastDstCol
        dv = wsD.Cells(DST_ROW_DATES, c).Value
        dk = dateKey(dv)
        If dk <> 0 Then If Not colByDate.exists(dk) Then colByDate.Add dk, c
    Next c
        Dim kk As Variant
    Debug.Print "-- Dest date headers --"
    For Each kk In colByDate.keys
        Debug.Print "DateKey:"; kk; " Col:"; colByDate(kk); " Display:"; wsD.Cells(DST_ROW_DATES, colByDate(kk)).text
    Next kk

    If colByDate.Count = 0 Then GoTo CleanExit

    ' ---- SOURCE: get last row ----
    stage = "source bounds"
    Dim lastSrc As Long
    lastSrc = wsS.Cells(wsS.Rows.Count, SRC_COL_SKU).End(xlUp).row
    If lastSrc < SRC_FIRST_DATA_ROW Then GoTo CleanExit

    ' ---- 1) Collect ordered unique WEEK dates from REVO Floor (column F) ----
    stage = "collect weeks"
    Dim weekDates() As Long, wCount As Long
    ReDim weekDates(1 To 200) ' generous cap; resize later

    Dim r As Long, dt As Variant, dKey As Long, prevKey As Long
    For r = SRC_FIRST_DATA_ROW To lastSrc
        dt = wsS.Cells(r, SRC_COL_DATE).Value
        dKey = dateKey(dt)
        If dKey <> 0 Then
            ' treat a change in date key as a new week boundary
            If dKey <> prevKey Then
                wCount = wCount + 1
                If wCount > UBound(weekDates) Then ReDim Preserve weekDates(1 To wCount + 50)
                weekDates(wCount) = dKey
                prevKey = dKey
            End If
        End If
    Next r
        Debug.Print "-- Week dates from REVO Floor --"
    For r = 1 To wCount
        Debug.Print r, weekDates(r)
    Next r

    If wCount = 0 Then GoTo CleanExit
    ReDim Preserve weekDates(1 To wCount)

    ' First week date on REVO Floor drives the destination start:
    Dim startDateKey As Long: startDateKey = weekDates(1)

    ' ---- 2) Build sums by (SKU × WeekIndex) from REVO Floor ----
    stage = "sum by SKU×Week"
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary") ' key: sku|weekIdx
    Dim sku As String, qty As Double, wkIdx As Long

    prevKey = 0: wkIdx = 0
    For r = SRC_FIRST_DATA_ROW To lastSrc
        sku = SafeText(wsS.Cells(r, SRC_COL_SKU).Value)
        If Len(sku) = 0 Then GoTo NextRow

        dKey = dateKey(wsS.Cells(r, SRC_COL_DATE).Value)
        If dKey = 0 Then GoTo NextRow

        ' increment week index when date changes
        If dKey <> prevKey Then
            wkIdx = wkIdx + 1
            prevKey = dKey
        End If
        If wkIdx < 1 Or wkIdx > wCount Then GoTo NextRow

            On Error GoTo SumErr
    qty = SafeDbl(wsS.Cells(r, SRC_COL_FORECAST).Value)
    On Error GoTo CleanFail

        If qty = 0# Then GoTo NextRow

        Dim k As String: k = sku & "|" & CStr(wkIdx)
        If sums.exists(k) Then
            sums(k) = CDbl(sums(k)) + qty
        Else
            sums.Add k, qty
        End If
            GoTo NextRow

SumErr:
    Debug.Print "Type mismatch at row"; r, _
                "SKU:"; sku, _
                "RawVal:"; wsS.Cells(r, SRC_COL_FORECAST).text
    Resume NextRow

NextRow:
    Next r
        Debug.Print "-- sums dict size --"; sums.Count
    Dim dbgKey As Variant, cnt As Long
    cnt = 0
    For Each dbgKey In sums.keys
        Debug.Print dbgKey, sums(dbgKey)
        cnt = cnt + 1
        If cnt >= 20 Then Exit For ' just sample
    Next dbgKey


    ' ---- 3) Clear and write destination columns:
    ' For each week index j, target date = startDateKey + 7*(j-1) ----
    stage = "write dest"
    Dim j As Long, targetDate As Long, col As Long, outRow As Long
    Dim rowBySku As Object: Set rowBySku = CreateObject("Scripting.Dictionary")

    ' Build SKU -> row map once
    Dim rr As Long, a As String
    For rr = DST_FIRST_DATA_ROW To lastDstRow
        a = SafeText(wsD.Cells(rr, DST_COL_SKU).Value)
        If Len(a) > 0 Then If Not rowBySku.exists(a) Then rowBySku.Add a, rr
    Next rr

    ' Clear all target week columns first (only those we’ll write)
   For j = 1 To wCount
    targetDate = startDateKey + 7 * (j - 1)
    If colByDate.exists(targetDate) Then
        col = colByDate(targetDate)
        ' Do NOT clear the total/subtotal row; only clear data rows
        wsD.Range(wsD.Cells(DST_FIRST_DATA_ROW, col), wsD.Cells(lastDstRow, col)).ClearContents
    End If
Next j

    
        Debug.Print "-- rowBySku count --"; rowBySku.Count
    ' Optional: peek a couple
    Dim rs As Variant, peek As Long
    peek = 0
    For Each rs In rowBySku.keys
        Debug.Print "SKU:", rs, "Row:", rowBySku(rs)
        peek = peek + 1
        If peek >= 20 Then Exit For
    Next rs


    ' Write values
    Dim key As Variant, parts() As String, sk As String, idx As Long
      For Each key In sums.keys
        parts = Split(CStr(key), "|")
        If UBound(parts) = 1 Then
            sk = parts(0)
            idx = val(parts(1))
            targetDate = startDateKey + 7 * (idx - 1)

            If Not rowBySku.exists(sk) Then
                ' SKU didn't match dest
                ' Debug.Print "No row for SKU"; sk
            ElseIf Not colByDate.exists(targetDate) Then
                ' Date didn't match dest
                ' Debug.Print "No col for date"; targetDate
            Else
                outRow = rowBySku(sk)
                col = colByDate(targetDate)
                wsD.Cells(outRow, col).Value = sums(key)
            End If
        End If
    Next key



        Debug.Print "==== UpdateProductionForecast_ByWeeks Debug ===="
    Debug.Print "stage: "; stage
    Debug.Print "lastDstRow: "; lastDstRow
    Debug.Print "lastDstCol: "; lastDstCol
    Debug.Print "colByDate.Count: "; IIf(Not colByDate Is Nothing, colByDate.Count, -1)
    Debug.Print "lastSrc: "; lastSrc
    Debug.Print "wCount: "; wCount
    Debug.Print "sums.Count: "; IIf(Not sums Is Nothing, sums.Count, -1)


CleanExit:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Exit Sub

CleanFail:
    MsgBox "UpdateProductionForecast_ByWeeks failed at stage: " & stage & vbCrLf & _
           "Err " & Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function TryDateKey(ByVal v As Variant, ByRef outKey As Long) As Boolean
    ' Returns True and fills outKey with the Excel date serial if v is a valid date (text or serial); otherwise False.
    On Error GoTo fail
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    ' Already a date literal?
    If IsDate(v) Then
        outKey = CLng(CDate(v)): TryDateKey = True: Exit Function
    End If

    ' Numeric serial (sometimes stored as Double)
    If IsNumeric(v) Then
        Dim d As Double: d = CDbl(v)
        If d > 0# And d < 60000# Then outKey = CLng(d): TryDateKey = True: Exit Function
    End If

    ' Text that looks like a date — try CDate cautiously
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) > 0 And IsDate(s) Then
        outKey = CLng(CDate(s)): TryDateKey = True: Exit Function
    End If
    Exit Function
fail:
    TryDateKey = False
End Function

Private Function TryDbl(ByVal v As Variant) As Double
    ' Tolerant numeric parse: strips commas/$/() and unicode minus.
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ",", "")
    s = Replace(s, "$", "")
    s = Replace(s, "–", "-")
    If LEFT$(s, 1) = "(" And Right$(s, 1) = ")" Then s = "-" & Mid$(s, 2, Len(s) - 2)
    If IsNumeric(s) Then TryDbl = CDbl(s) Else TryDbl = 0#
End Function


Private Function SafeCLng(v As Variant) As Long
    On Error GoTo Bad
    Dim s As String: s = Trim$(CStr(v))
    s = Replace(s, ",", "")
    s = Replace(s, "–", "-")
    If LEFT$(s, 1) = "(" And Right$(s, 1) = ")" Then
        s = "-" & Mid$(s, 2, Len(s) - 2)
    End If
    If IsNumeric(s) Then
        SafeCLng = CLng(CDbl(s))
    Else
        SafeCLng = 0
    End If
    Exit Function
Bad:
    SafeCLng = 0
End Function



Private Function dateKey(v As Variant) As Long
    ' returns CLng(DateValue) or 0 if not a valid date
    On Error GoTo Bad
    If IsDate(v) Then
        dateKey = CLng(CDate(v))
    Else
        dateKey = 0
    End If
    Exit Function
Bad:
    dateKey = 0
End Function





' One-time correction macro:
' - Reads dates from "Release Log" (A=Date, B=SKU, C=Cart, E=WO)
' - Applies date to "Rework Tracker" (A=Date, B=WO, C=SKU, D=Cart)
' - Match key: WO|SKU|Cart
' - Chooses the EARLIEST release date per key
' - By default fills blanks only (set OVERWRITE_EXISTING:=True to overwrite)

Public Sub BackfillReworkDatesFromReleaseLog(Optional ByVal OVERWRITE_EXISTING As Boolean = False)
    Dim scrn&, evt&, calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Const SH_RELEASE As String = "Release Log"
    Const SH_REWORK  As String = "Rework Tracker"

    ' --- column map (based on your helpers earlier)
    ' Release Log:  A=Date, B=SKU, C=Cart, D=ReleasedQty, E=WO, F=Rework, G=B Grade, H=Reject, I=Verified
    Const RL_COL_DATE As Long = 1
    Const RL_COL_SKU  As Long = 2
    Const RL_COL_CART As Long = 3
    Const RL_COL_WO   As Long = 5

    ' Rework Tracker: A=Date, B=WO, C=SKU, D=Cart, E=Qty, F=Status
    Const RW_COL_DATE As Long = 1
    Const RW_COL_WO   As Long = 2
    Const RW_COL_SKU  As Long = 3
    Const RW_COL_CART As Long = 4

    Dim wsRL As Worksheet, wsRW As Worksheet
    Set wsRL = ThisWorkbook.Worksheets(SH_RELEASE)
    Set wsRW = ThisWorkbook.Worksheets(SH_REWORK)

    Dim lastRL As Long, lastRW As Long
    lastRL = wsRL.Cells(wsRL.Rows.Count, RL_COL_WO).End(xlUp).row
    lastRW = wsRW.Cells(wsRW.Rows.Count, RW_COL_WO).End(xlUp).row
    If lastRL < 2 Or lastRW < 2 Then GoTo CleanExit

    ' --- Build map: key -> earliestDate
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String, dt As Variant, WO$, sku$, Cart$

    For r = 2 To lastRL
        WO = NzTrim(wsRL.Cells(r, RL_COL_WO).Value)
        sku = NzTrim(wsRL.Cells(r, RL_COL_SKU).Value)
        Cart = NzTrim(wsRL.Cells(r, RL_COL_CART).Value)
        If Len(WO) = 0 And Len(sku) = 0 And Len(Cart) = 0 Then GoTo NextRL

        dt = wsRL.Cells(r, RL_COL_DATE).Value
        If IsDate(dt) Then
            k = MakeKey(WO, sku, Cart)
            If d.exists(k) Then
                ' keep earliest date
                If CLng(CDate(dt)) < d(k) Then d(k) = CLng(CDate(dt))
            Else
                d.Add k, CLng(CDate(dt))
            End If
        End If
NextRL:
    Next r

    ' --- Apply to Rework Tracker
    Dim applied&, skipped_no_match&, skipped_has_date&
    Dim rw As Long, curDate As Variant

    For rw = 2 To lastRW
        WO = NzTrim(wsRW.Cells(rw, RW_COL_WO).Value)
        sku = NzTrim(wsRW.Cells(rw, RW_COL_SKU).Value)
        Cart = NzTrim(wsRW.Cells(rw, RW_COL_CART).Value)
        If Len(WO) = 0 And Len(sku) = 0 And Len(Cart) = 0 Then GoTo NextRW

        k = MakeKey(WO, sku, Cart)
        If d.exists(k) Then
            curDate = wsRW.Cells(rw, RW_COL_DATE).Value
            If OVERWRITE_EXISTING Or Not IsDate(curDate) Then
                wsRW.Cells(rw, RW_COL_DATE).Value = CDate(d(k))
                applied = applied + 1
            Else
                skipped_has_date = skipped_has_date + 1
            End If
        Else
            skipped_no_match = skipped_no_match + 1
        End If
NextRW:
    Next rw

    MsgBox "Backfill complete:" & vbCrLf & _
           "  Dates applied: " & applied & vbCrLf & _
           "  Skipped (no match in Release Log): " & skipped_no_match & vbCrLf & _
           "  Skipped (date already present): " & skipped_has_date, vbInformation

CleanExit:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    Exit Sub

CleanFail:
    Application.ScreenUpdating = CBool(scrn)
    Application.EnableEvents = CBool(evt)
    Application.Calculation = calc
    MsgBox "BackfillReworkDatesFromReleaseLog failed: " & Err.Description, vbExclamation
End Sub

' -------- helpers --------
Private Function MakeKey(ByVal WO As String, ByVal sku As String, ByVal Cart As String) As String
    MakeKey = UCase$(Trim$(WO)) & "|" & UCase$(Trim$(sku)) & "|" & UCase$(Trim$(Cart))
End Function

Private Function NzTrim(v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        NzTrim = ""
    Else
        NzTrim = Trim$(CStr(v))
    End If
End Function

Public Sub Backfill_NoOverwrite()
    On Error Resume Next
    Application.Run "BackfillReworkDatesFromReleaseLog", False
End Sub

Public Sub Backfill_Overwrite()
    On Error Resume Next
    Application.Run "BackfillReworkDatesFromReleaseLog", True
End Sub

Public Sub ArrangeREVOFloorByWeeks(ByRef weekCaps() As Double, ByRef weekMaxOp() As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SRC_SHEET)
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim i As Long, w As Long
    Dim opCols() As Long, opVals() As Long, opN As Long
    
    ' --- guards
    If (Not Not weekCaps) = 0 Or (Not Not weekMaxOp) = 0 Then
        Err.Raise vbObjectError + 901, , "Week arrays not provided."
    End If
    w = UBound(weekCaps) ' 1-based expected
    If w <> UBound(weekMaxOp) Then Err.Raise vbObjectError + 902, , "Week arrays must be same size."
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail
    
    ' --- bounds
    lastRow = ws.Cells(ws.Rows.Count, COL_SKU).End(xlUp).row
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    
    ' --- detect OP columns from row 2 (numeric headers)
    DetectOpColumns ws, HEADER_ROW, lastCol, opCols, opVals, opN
    If opN = 0 Then Err.Raise vbObjectError + 903, , "No operation columns detected on row " & HEADER_ROW & "."
    
    ' --- build assignment arrays
    Dim nRows As Long: nRows = Application.Max(0, lastRow - DATA_FIRST_ROW + 1)
    If nRows = 0 Then GoTo CleanExit
    
    Dim rowWeek() As Long: ReDim rowWeek(1 To nRows)
    Dim rowIdx() As Long:  ReDim rowIdx(1 To nRows)
    Dim rowQty() As Double: ReDim rowQty(1 To nRows)
    
    ' capture original order + qty
    For r = DATA_FIRST_ROW To lastRow
        i = r - DATA_FIRST_ROW + 1
        rowIdx(i) = r
        rowQty(i) = SafeDbl(ws.Cells(r, COL_QTY).Value)
        rowWeek(i) = 0
    Next r
    
    ' --- assign rows week-by-week
    Dim capLeft() As Double: ReDim capLeft(1 To w)
    For i = 1 To w: capLeft(i) = weekCaps(i): Next i
    
    Dim eligible As Boolean
    For w = 1 To w
        For i = 1 To nRows
            If rowWeek(i) = 0 Then
                If rowQty(i) > 0 Then
                    eligible = RowEligibleForWeek(ws, rowIdx(i), opCols, opVals, opN, weekMaxOp(w))
                    If eligible Then
                        If rowQty(i) <= capLeft(w) Then
                            rowWeek(i) = w
                            capLeft(w) = capLeft(w) - rowQty(i)
                        End If
                    End If
                End If
            End If
        Next i
    Next w
    
    ' sanity: any unassigned rows ? append to last week
    For i = 1 To nRows
        If rowWeek(i) = 0 Then rowWeek(i) = w
    Next i
    
    ' --- write temp keys and week tags, then sort entire rows
    Dim tmpKeyCol As Long, tmpWeekCol As Long
    tmpKeyCol = lastCol + 1
    tmpWeekCol = lastCol + 2
    ws.Cells(HEADER_ROW - 1, tmpKeyCol).Value = "__TMP_KEY__"
    ws.Cells(HEADER_ROW - 1, tmpWeekCol).Value = "__TMP_WEEK__"
    
    Dim sortKey As Double, seq As Long: seq = 0
    For i = 1 To nRows
        seq = seq + 1
        r = rowIdx(i)
        sortKey = CDbl(rowWeek(i)) * 10000000# + CDbl(seq)
        ws.Cells(r, tmpKeyCol).Value = sortKey
        ws.Cells(r, tmpWeekCol).Value = rowWeek(i)
    Next i
    
    ' sort by tmpKey
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Cells(DATA_FIRST_ROW, tmpKeyCol), _
                        SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(DATA_FIRST_ROW, 1), ws.Cells(lastRow, tmpWeekCol))
        .header = xlNo
        .Apply
    End With
    
    ' --- insert spacer rows and merge/label column A per week
    Dim cur As Long, startRow As Long, endRow As Long, curWeek As Long, nextWeek As Long
    r = DATA_FIRST_ROW
    If r > ws.Cells(ws.Rows.Count, tmpWeekCol).End(xlUp).row Then GoTo CleanupHelpers
    
    curWeek = SafeLng(ws.Cells(r, tmpWeekCol).Value)
    startRow = r
    
    Do
        ' find end of this week block
        Do While r <= ws.Cells(ws.Rows.Count, tmpWeekCol).End(xlUp).row
            nextWeek = SafeLng(ws.Cells(r + 1, tmpWeekCol).Value)
            If nextWeek <> curWeek Then Exit Do
            r = r + 1
        Loop
        endRow = r
        
        ' merge & label column A for this block
        If endRow >= startRow Then
            With ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 1))
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = "Week " & curWeek
            End With
        End If
        
        ' insert single spacer row after the block (unless last block)
        If nextWeek <> 0 Then
            ws.Rows(endRow + 1).Insert Shift:=xlDown
            r = endRow + 2          ' skip spacer
            startRow = r
            curWeek = nextWeek
        Else
            Exit Do
        End If
    Loop
    
CleanupHelpers:
    ' remove helper columns
    ws.Columns(tmpKeyCol).EntireColumn.Delete
    ws.Columns(tmpWeekCol - 1).EntireColumn.Delete   ' after previous delete, tmpWeekCol shifted left
    
    ShadeWeeksMediumLight

    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "ArrangeREVOFloorByWeeks failed: " & Err.Description, vbExclamation
End Sub

'---------------- helpers ----------------
Private Sub DetectOpColumns(ws As Worksheet, ByVal hdrRow As Long, ByVal lastCol As Long, _
                            ByRef opCols() As Long, ByRef opVals() As Long, ByRef opN As Long)
    Dim c As Long, v As Variant
    Dim tmpCols() As Long, tmpVals() As Long, n As Long
    ReDim tmpCols(1 To lastCol)
    ReDim tmpVals(1 To lastCol)
    For c = 1 To lastCol
        v = ws.Cells(hdrRow, c).Value
        If IsNumeric(v) Then
            n = n + 1
            tmpCols(n) = c
            tmpVals(n) = CLng(v)
        End If
    Next c
    If n > 0 Then
        ReDim opCols(1 To n)
        ReDim opVals(1 To n)
        For c = 1 To n
            opCols(c) = tmpCols(c)
            opVals(c) = tmpVals(c)
        Next c
    End If
    opN = n
End Sub

Private Function RowEligibleForWeek(ws As Worksheet, ByVal rowNum As Long, _
                                    ByRef opCols() As Long, ByRef opVals() As Long, ByVal opN As Long, _
                                    ByVal maxOp As Long) As Boolean
    Dim j As Long
    For j = 1 To opN
        If opVals(j) <= maxOp Then
            If Boolish(ws.Cells(rowNum, opCols(j)).Value) Then
                RowEligibleForWeek = True
                Exit Function
            End If
        End If
    Next j
    RowEligibleForWeek = False
End Function



Private Function SafeLng(v As Variant) As Long
    If IsError(v) Or IsEmpty(v) Or v = "" Then SafeLng = 0 Else SafeLng = CLng(v)
End Function

Public Sub ShadeWeeksMediumLight(Optional ByVal sheetName As String = "REVO Floor")
    Dim ws As Worksheet, lastRow As Long, r As Long, band As Range
    Dim colors As Variant, ci As Long
    ' 15 medium-light colors (pastels). Will rotate if weeks > 15.
    colors = Array( _
        RGB(221, 235, 247), _
        RGB(226, 239, 218), _
        RGB(255, 242, 204), _
        RGB(248, 229, 229), _
        RGB(234, 241, 221), _
        RGB(242, 242, 242), _
        RGB(217, 225, 242), _
        RGB(222, 235, 246), _
        RGB(255, 229, 204), _
        RGB(230, 230, 250), _
        RGB(210, 245, 255), _
        RGB(255, 230, 240), _
        RGB(230, 255, 230), _
        RGB(255, 250, 210), _
        RGB(235, 220, 255) _
    )

    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Application.ScreenUpdating = False
    ci = 0

    r = 2 ' assume headers on row 1
    Do While r <= lastRow
        ' Look for merged "Week n" label starting at this row
        If ws.Cells(r, 1).MergeCells Then
            Set band = ws.Cells(r, 1).MergeArea
            If LCase$(Trim$(band.Cells(1, 1).Value)) Like "week *" Then
                ' Shade columns B:F for the whole merged block
                ws.Range(ws.Cells(band.row, 2), ws.Cells(band.row + band.Rows.Count - 1, 6)) _
                  .Interior.Color = colors(ci Mod (UBound(colors) + 1))
                ci = ci + 1
                r = band.row + band.Rows.Count ' skip to row after this band
                GoTo NextRow
            End If
        End If
        r = r + 1
NextRow:
    Loop

    Application.ScreenUpdating = True
End Sub



Public Sub PlanWeeks_UI()
    Load frmWeekPlanner
    frmWeekPlanner.Show
    Unload frmWeekPlanner
End Sub


Public Sub Diagnose_frmWeekPlanner()
    Dim vbComp As Object, hitCount As Long, formFound As Boolean
    Dim modName As String, lineText As String, i As Long, cType As Long
    
    Debug.Print "=== Components ==="
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        cType = vbComp.Type                      ' 0=StdModule,1=Class,2=MSForm,3=UserControl, etc.
        Debug.Print vbComp.name & "  Type=" & cType
        If StrComp(vbComp.name, "frmWeekPlanner", vbTextCompare) = 0 Then
            formFound = (cType = 3) Or (cType = 2)      ' 3 is UserForm in many hosts; some list 2
        End If
    Next vbComp
    Debug.Print "Form component named 'frmWeekPlanner' found?  "; formFound

    Debug.Print "=== Text collisions (any use of the name) ==="
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Or vbComp.Type = 0 Then ' Class or Std module
            For i = 1 To vbComp.CodeModule.CountOfLines
                lineText = vbComp.CodeModule.lines(i, 1)
                If InStr(1, lineText, "frmWeekPlanner", vbTextCompare) > 0 Then
                    Debug.Print vbComp.name & "  line " & i & ":  "; lineText
                    hitCount = hitCount + 1
                End If
            Next i
        End If
    Next vbComp
    Debug.Print "Total hits: "; hitCount
    Debug.Print "Tip: If you see something like `Public frmWeekPlanner As ...` or a Sub named `frmWeekPlanner`, rename it."
End Sub

Public Sub Build_EST_REVO_INV()
    Dim scrn As Boolean, evt As Boolean
    Dim calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim wsEst As Worksheet
    Set wsEst = EnsureSheet(wb, SH_EST)
    ' Clear existing
    wsEst.Cells.Clear

    ' Date range: current month through Dec 2026
    Dim startMonth As Date, endMonth As Date
    startMonth = DateSerial(Year(Date), Month(Date), 1)
    endMonth = DateSerial(2026, 12, 1)

    Dim monthCount As Long
    monthCount = DateDiff("m", startMonth, endMonth) + 1

    ' Build headers
    wsEst.Cells(1, 1).Value = "SKU"
    Dim i As Long
    For i = 0 To monthCount - 1
        wsEst.Cells(1, 2 + i).Value = Format$(DateAdd("m", i, startMonth), "mmm yyyy")
    Next i

    ' Build data sources
    Dim dInv As Object, dSales As Object, dProd As Object, dRemCurr As Object
    Dim dSku As Object
    Set dInv = CreateObject("Scripting.Dictionary")
    Set dSales = CreateObject("Scripting.Dictionary")
    Set dProd = CreateObject("Scripting.Dictionary")
    Set dRemCurr = CreateObject("Scripting.Dictionary")
    Set dSku = CreateObject("Scripting.Dictionary")

    ' Load REVO_INV (SKU in col A, Qty in col B)
    Load_Inv wb, dInv, dSku

    ' Load Sales Forecast (REVO_SALES_FRCST: SKU col A, months in row 1)
    Load_SalesForecast wb, dSales, dSku

    ' Load Production Forecast (Production Forecast sheet)
    Load_ProdForecast wb, dProd, dSku

    ' Load remaining current-month REVO Floor production
    Load_RemainingCurrentMonth wb, dRemCurr, dSku, startMonth

    ' Dump SKUs sorted
    Dim skuArr() As String
    skuArr = SortedKeys(dSku)

    Dim rowOut As Long
    rowOut = 2

    Dim currMonthKey As Long
    currMonthKey = MonthKey(startMonth)

    Dim m As Long, mKey As Long
    Dim sku As String
    Dim estPrev As Double, estCurr As Double
    Dim onHand As Double, remCurr As Double
    Dim key As String

    For i = LBound(skuArr) To UBound(skuArr)
        sku = skuArr(i)
        If Len(sku) > 0 Then
            wsEst.Cells(rowOut, 1).Value = sku

            estPrev = 0

            For m = 0 To monthCount - 1
                mKey = MonthKey(DateAdd("m", m, startMonth))
                key = mKey & "|" & sku

                If m = 0 Then
                    ' Current month:
                    onHand = GetDictDbl(dInv, sku)
                    remCurr = GetDictDbl(dRemCurr, sku) ' only current month stored by SKU
                    estCurr = onHand + remCurr - GetDictDbl(dSales, key)
                Else
                    ' Future months:
                    estCurr = estPrev _
                              + GetDictDbl(dProd, key) _
                              - GetDictDbl(dSales, key)
                End If

                wsEst.Cells(rowOut, 2 + m).Value = estCurr
                estPrev = estCurr
            Next m

            rowOut = rowOut + 1
        End If
    Next i

    ' Format as table
    Dim lastRow As Long, lastCol As Long
    lastRow = wsEst.Cells(wsEst.Rows.Count, 1).End(xlUp).row
    lastCol = wsEst.Cells(1, wsEst.Columns.Count).End(xlToLeft).Column

    If lastRow >= 2 And lastCol >= 2 Then
        Dim lo As ListObject
        ' Delete existing table if re-running
        On Error Resume Next
        For Each lo In wsEst.ListObjects
            lo.Unlist
        Next lo
        On Error GoTo CleanFail

        Dim tbl As ListObject
        Set tbl = wsEst.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    Source:=wsEst.Range(wsEst.Cells(1, 1), wsEst.Cells(lastRow, lastCol)), _
                    XlListObjectHasHeaders:=xlYes)
        tbl.name = "tbl_EST_REVO_INV"
    End If


CleanExit:
    Application.ScreenUpdating = scrn
    Application.EnableEvents = evt
    Application.Calculation = calc
    Exit Sub

CleanFail:
    ' Minimal error surfacing; adjust if you want silent fail
    MsgBox "Build_EST_REVO_INV failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' ========= HELPERS =========



Private Function MonthKey(d As Date) As Long
    If d = 0 Then
        MonthKey = 0
    Else
        MonthKey = Year(d) * 100& + Month(d)
    End If
End Function

Private Function GetDictDbl(d As Object, k As String) As Double
    If Not d Is Nothing Then
        If d.exists(k) Then
            If IsNumeric(d(k)) Then
                GetDictDbl = CDbl(d(k))
            End If
        End If
    End If
End Function

Private Sub AddQtyToDict(ByVal key As String, ByVal qty As Double, ByRef d As Object)
    If qty = 0 Then Exit Sub
    If Not d.exists(key) Then
        d(key) = qty
    Else
        d(key) = CDbl(d(key)) + qty
    End If
End Sub

Private Sub AddSKUKey(ByVal sku As String, ByRef dSku As Object)
    If Len(sku) = 0 Then Exit Sub
    If Not dSku.exists(sku) Then dSku(sku) = 1
End Sub

Private Function SortedKeys(d As Object) As String()
    Dim arr() As String
    Dim i As Long, j As Long
    Dim tmp As String

    If d Is Nothing Or d.Count = 0 Then
        ReDim arr(0 To -1)
        SortedKeys = arr
        Exit Function
    End If

    ReDim arr(0 To d.Count - 1)
    i = 0
    Dim k As Variant
    For Each k In d.keys
        arr(i) = CStr(k)
        i = i + 1
    Next k

    ' Simple bubble sort (N is small-ish here, fine)
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    SortedKeys = arr
End Function

' ==== LOADERS ====

Private Sub Load_Inv(wb As Workbook, ByRef dInv As Object, ByRef dSku As Object)
    On Error GoTo Skip
    Dim ws As Worksheet
    Set ws = wb.Worksheets(SH_REVO_INV)

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim sku As String, qty As Double

    For r = 2 To lastRow
        sku = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(sku) > 0 Then
            qty = CDbl(SafeVal(ws.Cells(r, 3).Value))
            AddQtyToDict sku, qty, dInv
            AddSKUKey sku, dSku
        End If
    Next r
Skip:
End Sub

Private Sub Load_SalesForecast( _
    wb As Workbook, _
    ByRef dSales As Object, _
    ByRef dSku As Object)

    On Error GoTo Skip

    Dim ws As Worksheet
    Set ws = wb.Worksheets(SH_REVO_FCST)   ' "REVO_SALES_FRCST"

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, FC_COL_SKU).End(xlUp).row

    Dim sku As String
    Dim d As Variant
    Dim mKey As Long
    Dim qty As Double
    Dim key As String

    For r = 2 To lastRow
        sku = Trim$(CStr(ws.Cells(r, FC_COL_SKU).Value))
        If Len(sku) > 0 Then

            d = ws.Cells(r, FC_COL_DATE).Value
            qty = SafeVal(ws.Cells(r, FC_COL_QTY).Value)

            If IsDate(d) And qty <> 0 Then
                mKey = MonthKey(CDate(d))
                key = mKey & "|" & sku

                AddQtyToDict key, qty, dSales
                AddSKUKey sku, dSku
            End If
        End If
    Next r

Skip:
End Sub


Private Sub Load_ProdForecast(wb As Workbook, ByRef dProd As Object, ByRef dSku As Object)
    On Error GoTo Skip
    Dim ws As Worksheet
    Set ws = wb.Worksheets(SH_DST)

    Dim lastCol As Long, lastRow As Long
    lastCol = ws.Cells(DST_ROW_DATES, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, DST_COL_SKU).End(xlUp).row

    Dim c As Long, r As Long
    Dim d As Date, mKey As Long
    Dim sku As String, qty As Double
    Dim key As String

    For c = 2 To lastCol
        If IsDate(ws.Cells(DST_ROW_DATES, c).Value) Then
            d = CDate(ws.Cells(DST_ROW_DATES, c).Value)
            mKey = MonthKey(d)

            For r = DST_FIRST_DATA_ROW To lastRow
                sku = Trim$(CStr(ws.Cells(r, DST_COL_SKU).Value))
                If Len(sku) > 0 Then
                    qty = Round(CDbl(SafeVal(ws.Cells(r, c).Value)) * 0.96, 0)
                        If qty <> 0 Then
                        key = mKey & "|" & sku
                        AddQtyToDict key, qty, dProd
                        AddSKUKey sku, dSku
                    End If
                End If
            Next r
        End If
    Next c
Skip:
End Sub

Private Sub Load_RemainingCurrentMonth(wb As Workbook, ByRef dRem As Object, ByRef dSku As Object, ByVal startMonth As Date)
    On Error GoTo Skip
    Dim ws As Worksheet
    Set ws = wb.Worksheets(SRC_SHEET)

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, SRC_COL_SKU).End(xlUp).row

    Dim sku As String
    Dim qty As Double
    Dim d As Variant
    Dim relVal As String
    Dim relCol As Long
    relCol = ws.Columns(COL_RELEASE).Column

    Dim mKeyRow As Long, mKeyCurr As Long
    mKeyCurr = MonthKey(startMonth)

    For r = SRC_FIRST_DATA_ROW To lastRow
        sku = Trim$(CStr(ws.Cells(r, SRC_COL_SKU).Value))
        If Len(sku) > 0 Then
            d = ws.Cells(r, SRC_COL_DATE).Value
            relVal = UCase$(Trim$(CStr(ws.Cells(r, relCol).Value)))
            qty = CDbl(SafeVal(ws.Cells(r, SRC_COL_FORECAST).Value))

            If IsDate(d) And qty <> 0 Then
                mKeyRow = MonthKey(CDate(d))
                If mKeyRow = mKeyCurr _
                   And CDate(d) > Date _
                   And relVal <> "RELEASED" Then

                    AddQtyToDict sku, qty, dRem
                    AddSKUKey sku, dSku
                End If
            End If
        End If
    Next r
Skip:
End Sub

Private Function SafeVal(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or Trim$(CStr(v)) = "" Then
        SafeVal = 0#
    ElseIf IsNumeric(v) Then
        SafeVal = CDbl(v)
    Else
        SafeVal = 0#
    End If
End Function

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

    Const SH_BOM As String = "BOM"
    Const BOM_COL_SKU As Long = 1
    Const BOM_COL_COMP As Long = 2
    Const BOM_COL_QTY As Long = 3
    
    
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
    wsOut.Cells(1, 6).Value = "Box_Comp_Type"

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
                    wsOut.Cells(rowOut, 6).Value = BoxCompType(compID, "PREASSEMBLY")

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
                            wsOut.Cells(rowOut, 6).Value = BoxCompType(subID, "SUBCOMPONENT")

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
                    wsOut.Cells(rowOut, 6).Value = BoxCompType(compID, "COMPONENT")

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
    With wsOut.Range("A1:F" & rowOut - 1)

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

Private Function BoxCompType(ByVal itemSku As String, ByVal itemType As String) As String
    Dim s As String: s = UCase$(Trim$(itemSku))
    Dim t As String: t = UCase$(Trim$(itemType))

    Dim isCompOrSub As Boolean
    isCompOrSub = (t = "SUBCOMPONENT") Or (t = "COMPONENT")
    Dim isPreOrComp As Boolean
    isPreOrComp = (t = "PREASSEMBLY") Or (t = "COMPONENT")

    'TUBE: allow COMPONENT or SUBCOMPONENT
    If isCompOrSub And InStr(1, s, "TUBE", vbTextCompare) > 0 Then
        BoxCompType = "TUBE": Exit Function
    End If
    
        'SIL: allow COMPONENT or SUBCOMPONENT
    If isCompOrSub And InStr(1, s, "SIL", vbTextCompare) > 0 Then
        BoxCompType = "SIL": Exit Function
    End If

    'TIP: allow COMPONENT or SUBCOMPONENT, must start with P PRE TIP
    If isCompOrSub And LEFT$(s, 9) = "P PRE TIP" Then
        BoxCompType = "TIP": Exit Function
    End If

    'VP: allow COMPONENT or SUBCOMPONENT
    If isCompOrSub And InStr(1, s, "VAULT PLATE", vbTextCompare) > 0 Then
        BoxCompType = "VP": Exit Function
    End If

    'INSERT: allow COMPONENT or SUBCOMPONENT
    If isCompOrSub And InStr(1, s, "INSERT", vbTextCompare) > 0 Then
        BoxCompType = "INSERT": Exit Function
    End If


    'PLUG: PREASSEMBLY only
    If t = "PREASSEMBLY" And InStr(1, s, "PLUG", vbTextCompare) > 0 Then
        BoxCompType = "PLUG": Exit Function
    End If

    'WS: PREASSEMBLY only (we'll prefer W TAPE during selection)
    If isPreOrComp And InStr(1, s, " WS", vbTextCompare) > 0 Then
        BoxCompType = "WS": Exit Function
    End If

    BoxCompType = ""
End Function



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


Private Sub RenderBlock(ByVal ws As Worksheet, ByVal title As String, _
                        ByVal r As Long, ByVal c As Long, _
                        ByRef weekKeys() As Long, ByRef loads() As Double, _
                        ByVal cap As Double)

    ws.Range(ws.Cells(r, c), ws.Cells(r, c + 3)).Merge
    ws.Cells(r, c).Value = title
    ws.Cells(r, c).Font.Bold = True
    ws.Cells(r, c).Interior.Color = RGB(230, 235, 242)

    ws.Cells(r + 1, c).Resize(1, 4).Value = Array("Week", "Capacity", "Load", "%")

    Dim i As Long
    For i = 1 To UBound(weekKeys)
        ws.Cells(r + 1 + i, c).Value = CDate(weekKeys(i))
        ws.Cells(r + 1 + i, c + 1).Value = cap
        ws.Cells(r + 1 + i, c + 2).Value = loads(i)
        ws.Cells(r + 1 + i, c + 3).FormulaR1C1 = "=IF(RC[-2]=0,"""",RC[-1]/RC[-2])"
    Next i

    ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 1 + UBound(weekKeys), c + 3)).NumberFormat = "#,##0;-#,##0;"""""

    ws.Range(ws.Cells(r + 2, c + 3), ws.Cells(r + 1 + UBound(weekKeys), c + 3)).NumberFormat = "0%"

    With ws.Range(ws.Cells(r + 2, c + 3), ws.Cells(r + 1 + UBound(weekKeys), c + 3)) _
        .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=1")
        .Interior.Color = RGB(255, 200, 200)
    End With
End Sub

Private Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureSheet.name = sheetName
    End If
End Function

Private Sub RenderCapacityBlock(ByVal ws As Worksheet, ByVal title As String, _
                                ByVal r As Long, ByVal c As Long, _
                                ByRef weekKeys() As Long, ByRef loads() As Double, _
                                ByVal cap As Double)

    ' Title row
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + 3))
        .Merge
        .Value = title
        .Font.name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Interior.Color = RGB(230, 235, 242)
        .Font.Color = RGB(40, 40, 40)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' Column headers
    ws.Cells(r + 1, c).Resize(1, 4).Value = Array("Week", "Capacity", "Load", "% Util")

    With ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 1, c + 3))
        .Font.Bold = True
        .Font.name = "Segoe UI"
        .Font.Size = 9
        .Interior.Color = RGB(230, 235, 242)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(210, 214, 220)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With

    Dim i As Long
    For i = 1 To UBound(weekKeys)
        ws.Cells(r + 1 + i, c).Value = CDate(weekKeys(i))
        ws.Cells(r + 1 + i, c + 1).Value = cap
        ws.Cells(r + 1 + i, c + 2).Value = loads(i)
        ws.Cells(r + 1 + i, c + 3).FormulaR1C1 = "=IF(RC[-2]=0,"""",RC[-1]/RC[-2])"
    Next i

    ' Formats
    ws.Range(ws.Cells(r + 2, c), ws.Cells(r + 1 + UBound(weekKeys), c)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r + 2, c + 1), ws.Cells(r + 1 + UBound(weekKeys), c + 2)).NumberFormat = "#,##0;-#,##0;"""""
    ws.Range(ws.Cells(r + 2, c + 3), ws.Cells(r + 1 + UBound(weekKeys), c + 3)).NumberFormat = "0%"

    ' Risk highlighting: >100% util
    With ws.Range(ws.Cells(r + 2, c + 3), ws.Cells(r + 1 + UBound(weekKeys), c + 3)).FormatConditions
        .Delete
        With .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=1")
            .Interior.Color = RGB(255, 200, 200)
        End With
    End With

    ' Borders
    With ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 1 + UBound(weekKeys), c + 3)).Borders
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(220, 223, 228)
    End With
End Sub


Private Function InferFamily(ByVal comp As String) As String
    Dim s As String: s = UCase$(comp)
    If InStr(s, "PLUG") > 0 Or InStr(s, "AIR") > 0 Or InStr(s, "REAR") > 0 Then
        InferFamily = "PLUG"
    ElseIf InStr(s, "WS") > 0 Or InStr(s, "WEIGHT") > 0 Then
        InferFamily = "WS"
    ElseIf InStr(s, "TIP") > 0 Or InStr(s, "WVP") > 0 Or InStr(s, "BVP") > 0 Then
        InferFamily = "TIP"
    Else
        InferFamily = ""
    End If
End Function





Private Function FlattenToLeaves_ForPreAsm(ByVal sku As String, ByVal children As Object, ByVal preAsm As Object) As Object
    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    out.CompareMode = vbTextCompare

    Call FlattenWalkLeaves(KeyNorm(sku), 1#, children, preAsm, out)

    Set FlattenToLeaves_ForPreAsm = out
End Function

Private Sub FlattenWalkLeaves(ByVal parent As String, ByVal mult As Double, ByVal children As Object, ByVal preAsm As Object, ByVal out As Object)
    Dim i As Long
    Dim arr As Variant
    Dim c As Collection
    Dim child As String, qPer As Double

    If Not children.exists(parent) Then Exit Sub
    Set c = children(parent)

    For i = 1 To c.Count
        arr = c(i)
        child = CStr(arr(0))
        qPer = SafeDbl(arr(1))
        If qPer = 0 Then qPer = 1

        ' If child is itself a preassembly, keep walking down
        If preAsm.exists(child) Then
            FlattenWalkLeaves child, mult * qPer, children, preAsm, out
        Else
            ' leaf component
            AddQty out, child, mult * qPer
        End If
    Next i
End Sub


Private Function WeekIdxForDate(ByVal d As Date, ByRef weekKeys() As Long) As Long
    ' Returns the wkIdx whose week start is <= d < next week start
    Dim i As Long, n As Long
    n = UBound(weekKeys)

    If n = 0 Then Exit Function
    If CLng(d) < weekKeys(1) Then
        WeekIdxForDate = 1
        Exit Function
    End If

    For i = 1 To n - 1
        If CLng(d) >= weekKeys(i) And CLng(d) < weekKeys(i + 1) Then
            WeekIdxForDate = i
            Exit Function
        End If
    Next i

    ' If on/after last week start, bucket into last week
    If CLng(d) >= weekKeys(n) Then WeekIdxForDate = n
End Function

Private Function MinDbl(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDbl = a Else MinDbl = b
End Function



Private Function SheetExists(ByVal nm As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(nm) Is Nothing
    On Error GoTo 0
End Function

Private Function BuildPFRowMap(ws As Worksheet, _
                               ByVal firstRow As Long, _
                               ByVal lastRow As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim r As Long
    Dim v As Variant
    Dim sku As String

    For r = firstRow To lastRow
        ' Read cell very defensively
        On Error Resume Next
        v = ws.Cells(r, 1).Value
        If Err.Number <> 0 Then
            Debug.Print "MapPF: error reading A" & r & " - skipped"
            Err.Clear
            On Error GoTo 0
            GoTo NextRow
        End If
        On Error GoTo 0

        ' If it's an array (can happen with certain merged/linked setups), skip
        If IsArray(v) Then
            Debug.Print "MapPF: A" & r & " is array/merged - skipped"
            GoTo NextRow
        End If

        ' Normalize to SKU text
        sku = KeyNorm(v)

        ' Skip empty
        If LenB(sku) = 0 Then GoTo NextRow

        ' Skip week labels or other section labels
        If LCase$(sku) Like "week *" Then GoTo NextRow

        ' Only first occurrence wins
        If Not d.exists(sku) Then
            d.Add sku, r
        End If

NextRow:
    Next r

    ' Optional debug:
    Debug.Print "MapPF: built", d.Count, "SKU rows from", firstRow, "to", lastRow

    Set BuildPFRowMap = d
End Function

Private Function ColLtr(ByVal colNum As Long) As String
    ' Convert 1-based column number to letters
    Dim q As Long, r As Long, s As String
    q = colNum
    Do While q > 0
        r = (q - 1) Mod 26
        s = Chr$(65 + r) & s
        q = (q - r) \ 26
    Loop
    ColLtr = s
End Function

Private Function row() As String
    ' Helper for CF formula string building; just returns current row placeholder
    row = "" ' not used directly (we inject row numbers concretely via address building)
End Function

Private Sub AddQty(d As Object, key As String, val As Double)
    If Len(key) = 0 Or val = 0 Then Exit Sub
    If Not d.exists(key) Then
        d.Add key, val
    Else
        d(key) = CDbl(d(key)) + val
    End If
End Sub

Private Function FlattenBOM_ForRCF( _
    ByVal sku As String, _
    ByVal children As Object, _
    ByVal parents As Object, _
    ByVal cache As Object, _
    ByVal preAsm As Object, _
    Optional ByVal path As Object = Nothing) As Object

    Dim sk As String
    sk = KeyNorm(sku)
    If sk = "" Then Exit Function

    ' Memoized?
    If cache.exists(sk) Then
        Set FlattenBOM_ForRCF = cache(sk)
        Exit Function
    End If

    ' Cycle guard
    If path Is Nothing Then
        Set path = CreateObject("Scripting.Dictionary")
    End If
    If path.exists(sk) Then
        Set FlattenBOM_ForRCF = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    path(sk) = True

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim pair As Variant
    Dim childID As String
    Dim mult As Double
    Dim flatChild As Object
    Dim fk As Variant

    If children.exists(sk) Then
        ' sku is a parent; walk its children
        For Each pair In children(sk)
            childID = KeyNorm(pair(0))
            mult = SafeDbl(pair(1))
            If mult = 0 Then mult = 1

            If preAsm.exists(childID) Then
                ' ?? Treat PREASSEMBLY as a leaf for THIS forecast
                AddQty result, childID, mult

            ElseIf children.exists(childID) Then
                ' child is some other assembly: recurse to its leaves
                Set flatChild = FlattenBOM_ForRCF(childID, children, parents, cache, preAsm, path)
                For Each fk In flatChild.keys
                    AddQty result, CStr(fk), mult * SafeDbl(flatChild(fk))
                Next fk

            Else
                ' plain component leaf
                AddQty result, childID, mult
            End If
        Next pair

    Else
        ' No children: treat this as a leaf
        AddQty result, sk, 1
    End If

    path.Remove sk
    Set cache(sk) = result
    Set FlattenBOM_ForRCF = result
End Function













Public Sub BuildConsolidatedMonthlyProduction()
    Dim scrn As Boolean, evt As Boolean
    Dim calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual
    
    Const OUT_SHEET As String = "Monthly Prod Consolidated"
    
    ' Predator color theme
    Const PRED_YELLOW As Long = 13041663 ' approx RGB(255,199,0)
    Const DARK_HEADER As Long = 9211020  ' dark grey for header
    Const LIGHT_GREY As Long = 15132390  ' light grey for data banding
    
    On Error GoTo CleanFail
    
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Source: Production Forecast (weekly layout)
    Dim wsPF As Worksheet
    Set wsPF = wb.Worksheets(SH_PROD)
    
    Dim lastCol As Long, lastRow As Long
    lastCol = wsPF.Cells(PF_ROW_DATES, wsPF.Columns.Count).End(xlToLeft).Column
    lastRow = wsPF.Cells(wsPF.Rows.Count, PF_COL_SKU).End(xlUp).row
    
    Dim dict As Object                 ' monthly totals by belor|month
    Dim dictMonths As Object           ' unique months
    Dim dictBelors As Object           ' unique belors
    Dim dictMonthWeeks As Object       ' # of weekly buckets per month
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set dictMonths = CreateObject("Scripting.Dictionary")
    Set dictBelors = CreateObject("Scripting.Dictionary")
    Set dictMonthWeeks = CreateObject("Scripting.Dictionary")
    
    Dim c As Long, r As Long
    Dim vDate As Variant, vQty As Double
    Dim monthStart As Date
    Dim sku As String, belor As String
    Dim key As String, mKey As String
    
    ' === Build consolidated monthly totals by belor ===
    For c = PF_COL_SKU + 1 To lastCol
        vDate = wsPF.Cells(PF_ROW_DATES, c).Value
        If IsDate(vDate) Then
            monthStart = DateSerial(Year(vDate), Month(vDate), 1)
            mKey = CStr(monthStart)
            
            If Not dictMonths.exists(mKey) Then
                dictMonths.Add mKey, monthStart
            End If
            
            ' Count weekly buckets contributing to each month
            If dictMonthWeeks.exists(mKey) Then
                dictMonthWeeks(mKey) = dictMonthWeeks(mKey) + 1
            Else
                dictMonthWeeks.Add mKey, 1
            End If
            
            For r = PF_FIRST_DATA_ROW To lastRow
                sku = Trim$(CStr(wsPF.Cells(r, PF_COL_SKU).Value))
                If Len(sku) > 0 Then
                    vQty = SafeDbl(wsPF.Cells(r, c).Value)
                    If vQty <> 0 Then
                        belor = GetBelorFromSKU(sku)
                        If Len(belor) > 0 Then
                            If Not dictBelors.exists(belor) Then
                                dictBelors.Add belor, belor
                            End If
                            
                            key = belor & "|" & mKey
                            If dict.exists(key) Then
                                dict(key) = dict(key) + vQty
                            Else
                                dict.Add key, vQty
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    Next c
    
    ' === Output to consolidated sheet ===
    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = wb.Worksheets(OUT_SHEET)
    On Error GoTo CleanFail
    
    If wsOut Is Nothing Then
        Set wsOut = wb.Worksheets.Add(After:=wsPF)
        wsOut.name = OUT_SHEET
    End If
    
    wsOut.Cells.Clear
    
    ' If no data, bail gracefully
    If dictMonths.Count = 0 Or dictBelors.Count = 0 Then
        wsOut.Range("A1").Value = "Consolidated Monthly Production Forecast"
        wsOut.Range("A3").Value = "No data found."
        GoTo CleanExit
    End If
    
    ' ----- Monthly title -----
    With wsOut.Range("A1")
        .Value = "Consolidated Monthly Production Forecast"
        .Font.name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = PRED_YELLOW
    End With
    
    ' ----- Build sorted month array -----
    Dim months() As Variant
    Dim i As Long, j As Long
    Dim tmp As Variant, k As Variant
    
    ReDim months(1 To dictMonths.Count)
    i = 0
    For Each k In dictMonths.keys
        i = i + 1
        months(i) = dictMonths(k)
    Next k
    
    ' Sort months ascending
    For i = LBound(months) To UBound(months) - 1
        For j = i + 1 To UBound(months)
            If months(j) < months(i) Then
                tmp = months(i)
                months(i) = months(j)
                months(j) = tmp
            End If
        Next j
    Next i
    
    ' ----- Monthly table header -----
    wsOut.Range("A3").Value = "Diameter"
    
    Dim col As Long
    col = 2
    For i = LBound(months) To UBound(months)
        wsOut.Cells(3, col).Value = months(i)
        wsOut.Cells(3, col).NumberFormat = "mmm yyyy"
        col = col + 1
    Next i
    
    Dim lastMonthCol As Long
    lastMonthCol = col - 1
    
    Dim totalCol As Long
    totalCol = lastMonthCol + 1
    wsOut.Cells(3, totalCol).Value = "Total"
    
    ' ===== Belor order preference =====
    Dim belorOrder As Variant
    belorOrder = Array("11.8", "12.4", "12.9", "AIR", "CRM", "12.4 30", "12.9 30")
    
    Dim rowOut As Long
    rowOut = 4
    
    ' First output in preferred order if present
    For i = LBound(belorOrder) To UBound(belorOrder)
        belor = CStr(belorOrder(i))
        If dictBelors.exists(belor) Then
            WriteBelorRow wsOut, rowOut, belor, months, dict, totalCol, lastMonthCol
            rowOut = rowOut + 1
        End If
    Next i
    
    ' Then output any additional belors not in the preferred list (alphabetical)
    Dim extraBelors() As Variant
    Dim extraCount As Long
    
    extraCount = 0
    For Each k In dictBelors.keys
        If Not IsInArray(CStr(k), belorOrder) Then
            extraCount = extraCount + 1
        End If
    Next k
    
    If extraCount > 0 Then
        ReDim extraBelors(1 To extraCount)
        extraCount = 0
        For Each k In dictBelors.keys
            If Not IsInArray(CStr(k), belorOrder) Then
                extraCount = extraCount + 1
                extraBelors(extraCount) = CStr(k)
            End If
        Next k
        
        ' Sort extra belors alphabetically
        For i = LBound(extraBelors) To UBound(extraBelors) - 1
            For j = i + 1 To UBound(extraBelors)
                If extraBelors(j) < extraBelors(i) Then
                    tmp = extraBelors(i)
                    extraBelors(i) = extraBelors(j)
                    extraBelors(j) = tmp
                End If
            Next j
        Next i
        
        For i = LBound(extraBelors) To UBound(extraBelors)
            belor = extraBelors(i)
            WriteBelorRow wsOut, rowOut, belor, months, dict, totalCol, lastMonthCol
            rowOut = rowOut + 1
        Next i
    End If
    
    Dim lastRowOut As Long
    lastRowOut = rowOut - 1
    
    ' ===== Totals row for monthly table =====
    Dim totalRow As Long
    totalRow = lastRowOut + 1
    
    Dim rData As Long, cMonth As Long
    
    wsOut.Cells(totalRow, 1).Value = "Total"
    
    ' Column totals per Month
    For cMonth = 2 To lastMonthCol
        wsOut.Cells(totalRow, cMonth).FormulaR1C1 = "=SUM(R4C:R" & lastRowOut & "C)"
        wsOut.Cells(totalRow, cMonth).NumberFormat = "0"
    Next cMonth
    
    ' Grand total (bottom-right)
    wsOut.Cells(totalRow, totalCol).FormulaR1C1 = "=SUM(R4C" & totalCol & ":R" & lastRowOut & "C" & totalCol & ")"
    wsOut.Cells(totalRow, totalCol).NumberFormat = "0"
    
    ' ===== Formatting for monthly table =====
    Dim lastColOut As Long
    lastColOut = totalCol
    
    With wsOut
        ' Merge monthly title across all columns
        .Range(.Cells(1, 1), .Cells(1, lastColOut)).Merge
        With .Range(.Cells(1, 1), .Cells(1, lastColOut))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Color = vbBlack
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Header row
        With .Range(.Cells(3, 1), .Cells(3, lastColOut))
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = DARK_HEADER
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' Data region
        With .Range(.Cells(4, 1), .Cells(lastRowOut, lastColOut))
            .Borders.LineStyle = xlContinuous
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
        
        ' Banding
        For rData = 4 To lastRowOut
            If (rData Mod 2) = 0 Then
                .Range(.Cells(rData, 1), .Cells(rData, lastColOut)).Interior.Color = LIGHT_GREY
            End If
        Next rData
        
        ' Totals row
        With .Range(.Cells(totalRow, 1), .Cells(totalRow, lastColOut))
            .Font.Bold = True
            .Interior.Color = PRED_YELLOW
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        ' Left-align Belor labels
        .Range(.Cells(4, 1), .Cells(lastRowOut, 1)).HorizontalAlignment = xlLeft
    End With
    
    ' ======================================================
    '   Weekly average table (mirror layout)
    ' ======================================================
    Dim weeklyTitleRow As Long
    Dim wkHeaderRow As Long, wkDataStartRow As Long
    Dim wkLastRow As Long, wkLastMonthCol As Long
    Dim wkTotalCol As Long, wkTotalRow As Long
    Dim wkLastColOut As Long
    
    ' "Go down 4 rows" after the monthly total row
    weeklyTitleRow = totalRow + 2
    wkHeaderRow = totalRow + 4
    wkDataStartRow = wkHeaderRow + 1
    
    ' Weekly title
    With wsOut.Range("A" & weeklyTitleRow)
        .Value = "Consolidated Weekly Production Forecast"
        .Font.name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = PRED_YELLOW
    End With
    
    ' Weekly header row
    wsOut.Cells(wkHeaderRow, 1).Value = "Diameter"
    col = 2
    For i = LBound(months) To UBound(months)
        wsOut.Cells(wkHeaderRow, col).Value = months(i)
        wsOut.Cells(wkHeaderRow, col).NumberFormat = "mmm yyyy"
        col = col + 1
    Next i
    
    wkLastMonthCol = col - 1
    wkTotalCol = wkLastMonthCol + 1
    wsOut.Cells(wkHeaderRow, wkTotalCol).Value = "Average"
    
    ' Weekly rows – same belor order
    rowOut = wkDataStartRow
    
    ' Preferred belors
    For i = LBound(belorOrder) To UBound(belorOrder)
        belor = CStr(belorOrder(i))
        If dictBelors.exists(belor) Then
            WriteBelorWeeklyRow wsOut, rowOut, belor, months, dict, dictMonthWeeks, wkTotalCol, wkLastMonthCol
            rowOut = rowOut + 1
        End If
    Next i
    
    ' Extra belors
    If extraCount > 0 Then
        For i = LBound(extraBelors) To UBound(extraBelors)
            belor = extraBelors(i)
            WriteBelorWeeklyRow wsOut, rowOut, belor, months, dict, dictMonthWeeks, wkTotalCol, wkLastMonthCol
            rowOut = rowOut + 1
        Next i
    End If
    
    wkLastRow = rowOut - 1
    wkTotalRow = wkLastRow + 1
    wkLastColOut = wkTotalCol
    
' Weekly totals row (mirror of monthly)
wsOut.Cells(wkTotalRow, 1).Value = "Total"

For cMonth = 2 To wkLastMonthCol
    wsOut.Cells(wkTotalRow, cMonth).FormulaR1C1 = _
        "=SUM(R[" & (wkDataStartRow - wkTotalRow) & "]C:R[" & (wkLastRow - wkTotalRow) & "]C)"
    wsOut.Cells(wkTotalRow, cMonth).NumberFormat = "0"
Next cMonth

' Average column (same approach)
wsOut.Cells(wkTotalRow, wkTotalCol).FormulaR1C1 = _
    "=AVERAGE(R[" & (wkDataStartRow - wkTotalRow) & "]C:R[" & (wkLastRow - wkTotalRow) & "]C)"
wsOut.Cells(wkTotalRow, wkTotalCol).NumberFormat = "0"

    
    ' Formatting for weekly table
    With wsOut
        ' Merge weekly title
        .Range(.Cells(weeklyTitleRow, 1), .Cells(weeklyTitleRow, wkLastColOut)).Merge
        With .Range(.Cells(weeklyTitleRow, 1), .Cells(weeklyTitleRow, wkLastColOut))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Color = vbBlack
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Header row
        With .Range(.Cells(wkHeaderRow, 1), .Cells(wkHeaderRow, wkLastColOut))
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = DARK_HEADER
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' Data region
        With .Range(.Cells(wkDataStartRow, 1), .Cells(wkLastRow, wkLastColOut))
            .Borders.LineStyle = xlContinuous
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
        
        ' Banding
        For rData = wkDataStartRow To wkLastRow
            If (rData Mod 2) = 0 Then
                .Range(.Cells(rData, 1), .Cells(rData, wkLastColOut)).Interior.Color = LIGHT_GREY
            End If
        Next rData
        
        ' Weekly totals row
        With .Range(.Cells(wkTotalRow, 1), .Cells(wkTotalRow, wkLastColOut))
            .Font.Bold = True
            .Interior.Color = PRED_YELLOW
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        ' Left-align Belor labels
        .Range(.Cells(wkDataStartRow, 1), .Cells(wkLastRow, 1)).HorizontalAlignment = xlLeft
        
        ' Auto-fit everything
        .Columns.AutoFit
    End With
    
CleanExit:
    Application.ScreenUpdating = scrn
    Application.EnableEvents = evt
    Application.Calculation = calc
    Exit Sub

CleanFail:
    MsgBox "BuildConsolidatedMonthlyProduction failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

'-------------------- HELPERS --------------------

Private Sub WriteBelorRow( _
    ByVal ws As Worksheet, _
    ByVal rowOut As Long, _
    ByVal belor As String, _
    ByRef months() As Variant, _
    ByVal dict As Object, _
    ByVal totalCol As Long, _
    ByVal lastMonthCol As Long)
    
    Dim col As Long, i As Long
    Dim key As String
    Dim monthStart As Date
    Dim v As Variant
    
    ws.Cells(rowOut, 1).Value = belor
    
    col = 2
    For i = LBound(months) To UBound(months)
        monthStart = months(i)
        key = belor & "|" & CStr(monthStart)
        If dict.exists(key) Then
            v = dict(key)
            If v <> 0 Then
                ws.Cells(rowOut, col).Value = v
                ws.Cells(rowOut, col).NumberFormat = "0"
            Else
                ws.Cells(rowOut, col).ClearContents
            End If
        Else
            ws.Cells(rowOut, col).ClearContents
        End If
        col = col + 1
    Next i
    
    ' Row total
    ws.Cells(rowOut, totalCol).FormulaR1C1 = "=SUM(RC2:RC" & lastMonthCol & ")"
    ws.Cells(rowOut, totalCol).NumberFormat = "0"
End Sub

Private Sub WriteBelorWeeklyRow( _
    ByVal ws As Worksheet, _
    ByVal rowOut As Long, _
    ByVal belor As String, _
    ByRef months() As Variant, _
    ByVal dict As Object, _
    ByVal dictMonthWeeks As Object, _
    ByVal wkTotalCol As Long, _
    ByVal wkLastMonthCol As Long)
    
    Dim col As Long, i As Long
    Dim key As String, mKey As String
    Dim monthStart As Date
    Dim weeksInMonth As Long
    Dim v As Variant
    
    ws.Cells(rowOut, 1).Value = belor
    
    col = 2
    For i = LBound(months) To UBound(months)
        monthStart = months(i)
        mKey = CStr(monthStart)
        key = belor & "|" & mKey
        
        If dict.exists(key) And dictMonthWeeks.exists(mKey) Then
            weeksInMonth = CLng(dictMonthWeeks(mKey))
            If weeksInMonth > 0 Then
                v = dict(key) / weeksInMonth
                ws.Cells(rowOut, col).Value = v
                ws.Cells(rowOut, col).NumberFormat = "0"
            Else
                ws.Cells(rowOut, col).ClearContents
            End If
        Else
            ws.Cells(rowOut, col).ClearContents
        End If
        
        col = col + 1
    Next i
    
    ' Row total for weekly averages (sum of avg per month)
    ws.Cells(rowOut, wkTotalCol).FormulaR1C1 = "=SUM(RC2:RC" & wkLastMonthCol & ")"
    ws.Cells(rowOut, wkTotalCol).NumberFormat = "0"
End Sub


Private Function GetBelorFromSKU(ByVal sku As String) As String
    Dim s As String
    s = UCase$(sku)
    
    ' --- AIR / CRM override numeric groups ---
    If InStr(s, "AIR") > 0 Then
        GetBelorFromSKU = "AIR"
        Exit Function
    End If
    
    If InStr(s, "CRM") > 0 Then
        GetBelorFromSKU = "CRM"
        Exit Function
    End If
    
    ' --- 30-inch variants ---
    If InStr(s, "12.4") > 0 And InStr(s, "30") > 0 Then
        GetBelorFromSKU = "12.4 30"
        Exit Function
    End If
    
    ' 12.9 30 bucket includes BK 30
    If (InStr(s, "12.9") > 0 Or InStr(s, "BK") > 0) And InStr(s, "30") > 0 Then
        GetBelorFromSKU = "12.9 30"
        Exit Function
    End If
    
    ' --- Standard diameters ---
    If InStr(s, "11.8") > 0 Then
        GetBelorFromSKU = "11.8"
        Exit Function
    End If
    
    If InStr(s, "12.4") > 0 Then
        GetBelorFromSKU = "12.4"
        Exit Function
    End If
    
    ' 12.9 bucket includes BK
    If InStr(s, "12.9") > 0 Or InStr(s, "BK") > 0 Then
        GetBelorFromSKU = "12.9"
        Exit Function
    End If
    
    ' If no match, return blank (row is ignored)
    GetBelorFromSKU = ""
End Function

Private Function IsInArray(ByVal val As String, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If CStr(arr(i)) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Public Sub BuildRevoProductionAllocator()
    Const SH_RCF As String = "REVO Component Forecast"
    Const SH_BOM As String = "BOM"
    Const SH_BIN As String = "Bin_Inv"
    Const OUT_SHEET As String = "REVO Production Allocator"

    ' REVO Component Forecast layout
    Const RCF_HDR_ROW As Long = 4
    Const RCF_FIRST_ROW As Long = 5
    Const RCF_COL_SKU As Long = 1
    Const RCF_COL_SS As Long = 2
    Const RCF_COL_STOCK As Long = 3
    Const RCF_COL_WIP As Long = 4
    Const RCF_COL_TOTAL As Long = 5
    Const RCF_COL_P1 As Long = 6
    Const RCF_COL_P2 As Long = 7
    Const RCF_COL_P3 As Long = 8
    Const RCF_FIRST_WEEK_COL As Long = 9   ' weekly demand starts here

    ' BOM layout
    Const BOM_COL_SKU As Long = 1
    Const BOM_COL_COMP As Long = 2
    Const BOM_COL_QTY As Long = 3

    Dim scrn As Boolean, evt As Boolean
    Dim calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual

    Dim stage As String
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim wsRCF As Worksheet, wsBOM As Worksheet, wsBin As Worksheet, wsOut As Worksheet
    Set wb = ThisWorkbook
    Set wsRCF = wb.Worksheets(SH_RCF)
    Set wsBOM = wb.Worksheets(SH_BOM)
    Set wsBin = wb.Worksheets(SH_BIN)

    '====================================================
    ' 1) Read REVO Component Forecast: weeks + demand
    '====================================================
    stage = "Read REVO Component Forecast"

    Dim lastRCFRow As Long, lastWeekCol As Long
    lastRCFRow = wsRCF.Cells(wsRCF.Rows.Count, RCF_COL_SKU).End(xlUp).row
    If lastRCFRow < RCF_FIRST_ROW Then GoTo CleanExit

    lastWeekCol = wsRCF.Cells(RCF_HDR_ROW, wsRCF.Columns.Count).End(xlToLeft).Column
    If lastWeekCol < RCF_FIRST_WEEK_COL Then GoTo CleanExit

    Dim w As Long
    w = lastWeekCol - RCF_FIRST_WEEK_COL + 1

    Dim weekDates() As Double
    ReDim weekDates(1 To w)

    Dim c As Long, wkIdx As Long
    For c = RCF_FIRST_WEEK_COL To lastWeekCol
        wkIdx = c - RCF_FIRST_WEEK_COL + 1
        weekDates(wkIdx) = CDbl(CDate(wsRCF.Cells(RCF_HDR_ROW, c).Value))
    Next c

    ' Preassemblies with demand > 0
    Dim preRows As Object      ' SKU -> row
    Set preRows = CreateObject("Scripting.Dictionary")
    preRows.CompareMode = vbTextCompare

    Dim r As Long
    Dim sku As String
    Dim totalDemand As Double, val As Double

    For r = RCF_FIRST_ROW To lastRCFRow
        sku = KeyNorm(wsRCF.Cells(r, RCF_COL_SKU).Value)
        If Len(sku) > 0 Then
            totalDemand = 0
            For c = RCF_FIRST_WEEK_COL To lastWeekCol
                val = ToDbl(wsRCF.Cells(r, c).Value)
                totalDemand = totalDemand + val
            Next c

            If totalDemand > 0 Then
                If Not preRows.exists(sku) Then preRows.Add sku, r
            End If
        End If
    Next r

    If preRows.Count = 0 Then GoTo CleanExit

    ' Sorted list of preassemblies (controls allocation priority)
    Dim preCount As Long
    preCount = preRows.Count

    Dim preArr() As String, i As Long
    ReDim preArr(1 To preCount)
    i = 1
    Dim k As Variant
    For Each k In preRows.keys
        preArr(i) = CStr(k)
        i = i + 1
    Next k
    If preCount > 1 Then QuickSortStrings preArr, 1, preCount

    '====================================================
    ' 2) Build BOM children map (parent -> components)
    '====================================================
    stage = "Build BOM map"

    Dim children As Object
    Set children = CreateObject("Scripting.Dictionary")
    children.CompareMode = vbTextCompare

    Dim lastBOMRow As Long
    lastBOMRow = wsBOM.Cells(wsBOM.Rows.Count, BOM_COL_SKU).End(xlUp).row

    Dim parentID As String, compID As String
    Dim qPer As Double

    For r = 2 To lastBOMRow
        parentID = KeyNorm(wsBOM.Cells(r, BOM_COL_SKU).Value)
        compID = KeyNorm(wsBOM.Cells(r, BOM_COL_COMP).Value)
        qPer = ToDbl(wsBOM.Cells(r, BOM_COL_QTY).Value)
        If qPer = 0 Then qPer = 1

        If Len(parentID) > 0 And Len(compID) > 0 Then
            If Not children.exists(parentID) Then children.Add parentID, New Collection
            children(parentID).Add Array(compID, qPer)
        End If
    Next r

    ' Filter preassemblies: demand > 0 AND has BOM rows
    Dim validPre As Object
    Set validPre = CreateObject("Scripting.Dictionary")
    validPre.CompareMode = vbTextCompare

    For i = 1 To preCount
        sku = preArr(i)
        If children.exists(sku) Then validPre(sku) = preRows(sku)
    Next i

    If validPre.Count = 0 Then GoTo CleanExit

    ' Rebuild preArr using only valid preassemblies
    preCount = validPre.Count
    ReDim preArr(1 To preCount)
    i = 1
    For Each k In validPre.keys
        preArr(i) = CStr(k)
        i = i + 1
    Next k
    If preCount > 1 Then QuickSortStrings preArr, 1, preCount

    '====================================================
    ' 3) Read Bin_Inv (component stock & WIP)
    '====================================================
    stage = "Read Bin_Inv"

    Dim stockInv As Object, wipInv As Object
    Set stockInv = CreateObject("Scripting.Dictionary")
    Set wipInv = CreateObject("Scripting.Dictionary")
    stockInv.CompareMode = vbTextCompare
    wipInv.CompareMode = vbTextCompare

    Dim lastBinRow As Long
    lastBinRow = wsBin.Cells(wsBin.Rows.Count, 1).End(xlUp).row

    Dim binName As String, skuBin As String
    Dim qBin As Double

    For r = 2 To lastBinRow
        binName = UCase$(Trim$(CStr(wsBin.Cells(r, 1).Value)))
        skuBin = KeyNorm(wsBin.Cells(r, 2).Value)
        qBin = ToDbl(wsBin.Cells(r, 3).Value)

        If Len(skuBin) > 0 And qBin <> 0 Then
            If InStr(binName, "WIP") > 0 Then
                ' any bin with "WIP" is WIP
                AddQty wipInv, skuBin, qBin
            Else
                ' anything else counts as stock (covers TIP / STOCK ROOM / etc)
                AddQty stockInv, skuBin, qBin
            End If
        End If
    Next r

    ' Global remaining component inventory (stock + WIP)
    Dim remCompInv As Object
    Set remCompInv = CreateObject("Scripting.Dictionary")
    remCompInv.CompareMode = vbTextCompare

    For Each k In stockInv.keys
        remCompInv(k) = ToDbl(stockInv(k))
    Next k
    For Each k In wipInv.keys
        If remCompInv.exists(k) Then
            remCompInv(k) = ToDbl(remCompInv(k)) + ToDbl(wipInv(k))
        Else
            remCompInv(k) = ToDbl(wipInv(k))
        End If
    Next k

    '====================================================
    ' 4) Preassembly-level arrays (demand, inventory, allocation)
    '====================================================
    stage = "Build preassembly arrays"

    Dim preRow() As Long
    ReDim preRow(1 To preCount)

    Dim preSS() As Double, preInv() As Double, preWip() As Double, preTotalInv() As Double
    Dim preP1() As Double, preP2() As Double, preP3() As Double
    Dim preDemandTotal() As Double, preProdTotal() As Double
    Dim preRemainInv() As Double
    Dim coverageWeekIdx() As Long
    Dim coverageDate() As Variant

    ReDim preSS(1 To preCount)
    ReDim preInv(1 To preCount)
    ReDim preWip(1 To preCount)
    ReDim preTotalInv(1 To preCount)
    ReDim preP1(1 To preCount)
    ReDim preP2(1 To preCount)
    ReDim preP3(1 To preCount)
    ReDim preDemandTotal(1 To preCount)
    ReDim preProdTotal(1 To preCount)
    ReDim preRemainInv(1 To preCount)
    ReDim coverageWeekIdx(1 To preCount)
    ReDim coverageDate(1 To preCount)

    Dim demand() As Double
    Dim allocInv() As Double
    Dim allocComp() As Double

    ReDim demand(1 To preCount, 1 To w)
    ReDim allocInv(1 To preCount, 1 To w)
    ReDim allocComp(1 To preCount, 1 To w)

    Dim rowRCF As Long

    For i = 1 To preCount
        sku = preArr(i)
        rowRCF = CLng(validPre(sku))
        preRow(i) = rowRCF

        preSS(i) = ToDbl(wsRCF.Cells(rowRCF, RCF_COL_SS).Value)
        preInv(i) = ToDbl(wsRCF.Cells(rowRCF, RCF_COL_STOCK).Value)
        preWip(i) = ToDbl(wsRCF.Cells(rowRCF, RCF_COL_WIP).Value)
        preTotalInv(i) = ToDbl(wsRCF.Cells(rowRCF, RCF_COL_TOTAL).Value)
        preP1(i) = ToDbl(wsRCF.Cells(rowRCF, RCF_COL_P1).Value)
        preP2(i) = ToDbl(wsRCF.Cells(rowRCF, RCF_COL_P2).Value)
        preP3(i) = ToDbl(wsRCF.Cells(rowRCF, RCF_COL_P3).Value)

        preRemainInv(i) = preTotalInv(i)
        preDemandTotal(i) = 0
        preProdTotal(i) = 0
        coverageWeekIdx(i) = 0
        coverageDate(i) = ""

        For wkIdx = 1 To w
            val = ToDbl(wsRCF.Cells(rowRCF, RCF_FIRST_WEEK_COL + wkIdx - 1).Value)
            demand(i, wkIdx) = val
            preDemandTotal(i) = preDemandTotal(i) + val
        Next wkIdx
    Next i

    '====================================================
    ' 5) Week-by-week global allocation
    '====================================================
    stage = "Run allocator"

    Dim firstComponent As Boolean
    Dim maxBuildable As Double, buildFromComp As Double, needAfterInv As Double
    Dim demandThis As Double, allocFromInv As Double
    Dim childCol As Collection, itemArr As Variant
    Dim compStock As Double, compQtyPer As Double, compAvail As Double
    Dim compMax As Double

    Dim preDemandCum() As Double
    Dim preProdCum() As Double
    Dim coverageBroken() As Boolean

    ReDim preDemandCum(1 To preCount)
    ReDim preProdCum(1 To preCount)
    ReDim coverageBroken(1 To preCount)

    For i = 1 To preCount
        preDemandCum(i) = 0
        preProdCum(i) = 0
        coverageBroken(i) = False
    Next i

    For wkIdx = 1 To w
        For i = 1 To preCount
            sku = preArr(i)
            demandThis = demand(i, wkIdx)

            If demandThis <= 0 Then
                ' Still update cumulative demand for coverage math
                preDemandCum(i) = preDemandCum(i) + demandThis
                ' coverage: if already broken, stay; else coverageWeekIdx stays where it was
                GoTo NextPre
            End If

            ' 1) Allocate from existing preassembly inventory first
            allocFromInv = preRemainInv(i)
            If allocFromInv > demandThis Then allocFromInv = demandThis
            preRemainInv(i) = preRemainInv(i) - allocFromInv
            needAfterInv = demandThis - allocFromInv

            ' 2) Allocate from components (global inventory)
            buildFromComp = 0
            If needAfterInv > 0 And children.exists(sku) Then
                Set childCol = children(sku)

                firstComponent = True
                maxBuildable = 0

                ' Determine max units we can build given all components
                For Each itemArr In childCol
                    compID = CStr(itemArr(0))
                    compQtyPer = ToDbl(itemArr(1))
                    If compQtyPer = 0 Then compQtyPer = 1

                    compStock = 0
                    If remCompInv.exists(compID) Then compStock = ToDbl(remCompInv(compID))

                    If compStock <= 0 Then
                        compMax = 0
                    Else
                        compMax = Fix(compStock / compQtyPer)
                    End If

                    If firstComponent Then
                        maxBuildable = compMax
                        firstComponent = False
                    Else
                        If compMax < maxBuildable Then maxBuildable = compMax
                    End If
                Next itemArr

                If firstComponent Then maxBuildable = 0 ' no children somehow

                buildFromComp = needAfterInv
                If buildFromComp > maxBuildable Then buildFromComp = maxBuildable

                ' Deduct component inventory if we actually build anything
                If buildFromComp > 0 Then
                    For Each itemArr In childCol
                        compID = CStr(itemArr(0))
                        compQtyPer = ToDbl(itemArr(1))
                        If compQtyPer = 0 Then compQtyPer = 1

                        compAvail = 0
                        If remCompInv.exists(compID) Then compAvail = ToDbl(remCompInv(compID))
                        compAvail = compAvail - buildFromComp * compQtyPer
                        If compAvail < 0 Then compAvail = 0
                        remCompInv(compID) = compAvail
                    Next itemArr
                End If
            End If

            ' Record allocations
            allocInv(i, wkIdx) = allocFromInv
            allocComp(i, wkIdx) = buildFromComp

            ' Update cumulative demand / production and coverage
            preDemandCum(i) = preDemandCum(i) + demandThis
            preProdCum(i) = preProdCum(i) + allocFromInv + buildFromComp

            If Not coverageBroken(i) Then
                If preProdCum(i) >= preDemandCum(i) Then
                    coverageWeekIdx(i) = wkIdx
                    coverageDate(i) = weekDates(wkIdx)
                Else
                    coverageBroken(i) = True
                End If
            End If

NextPre:
        Next i
    Next wkIdx

    ' Final total production per preassembly
    For i = 1 To preCount
        preProdTotal(i) = preProdCum(i)
    Next i

    '====================================================
    ' 6) Create / reset output sheet
    '====================================================
    stage = "Create output sheet"

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(OUT_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail

    Set wsOut = wb.Worksheets.Add(After:=wsRCF)
    wsOut.name = OUT_SHEET

    ' Title + last updated
    With wsOut.Range("A1")
        .Value = "REVO Production Allocator (Demand-Driven)"
        .Font.name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
    End With

    With wsOut.Range("A2")
        .Value = "Last Updated " & Format(Now, "dddd, m/d/yyyy h:mm AM/PM")
        .Font.name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(90, 90, 90)
    End With

    Dim outRow As Long
    outRow = 4

    ' Detail header (single table, filterable)
    wsOut.Cells(outRow, 1).Value = "Preassembly SKU"
    wsOut.Cells(outRow, 2).Value = "Week Ending"
    wsOut.Cells(outRow, 3).Value = "Week #"
    wsOut.Cells(outRow, 4).Value = "Weekly Demand"
    wsOut.Cells(outRow, 5).Value = "Alloc from Finished Inv"
    wsOut.Cells(outRow, 6).Value = "Alloc from Components"
    wsOut.Cells(outRow, 7).Value = "Total Allocated"
    wsOut.Cells(outRow, 8).Value = "Weekly Shortfall"

    wsOut.Cells(outRow, 9).Value = "Total Demand (All Weeks)"
    wsOut.Cells(outRow, 10).Value = "Total Produced"
    wsOut.Cells(outRow, 11).Value = "Total Shortfall"
    wsOut.Cells(outRow, 12).Value = "Coverage Weeks"
    wsOut.Cells(outRow, 13).Value = "Coverage Through"

    With wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 13))
        .Font.Bold = True
        .Font.name = "Segoe UI"
        .Font.Size = 10
        .Interior.Color = RGB(230, 235, 242)
    End With

    outRow = outRow + 1

    '====================================================
    ' 7) Write rows: Preassembly × Week (only where something happens)
    '====================================================
    stage = "Write allocator rows"

    Dim weeklyAlloc As Double, weeklyShort As Double
    Dim totalShort As Double
    Dim covWeeks As Long
    Dim covDate As Variant

    For i = 1 To preCount
        sku = preArr(i)
        totalShort = preDemandTotal(i) - preProdTotal(i)
        If totalShort < 0 Then totalShort = 0
        covWeeks = coverageWeekIdx(i)
        covDate = coverageDate(i)

        For wkIdx = 1 To w
            demandThis = demand(i, wkIdx)
            weeklyAlloc = allocInv(i, wkIdx) + allocComp(i, wkIdx)

            ' Only show rows where there is demand or allocation
            If demandThis <> 0 Or weeklyAlloc <> 0 Then
                wsOut.Cells(outRow, 1).Value = sku
                wsOut.Cells(outRow, 2).Value = CDate(weekDates(wkIdx))
                wsOut.Cells(outRow, 3).Value = wkIdx

                wsOut.Cells(outRow, 4).Value = Round(demandThis, 0)
                wsOut.Cells(outRow, 5).Value = Round(allocInv(i, wkIdx), 0)
                wsOut.Cells(outRow, 6).Value = Round(allocComp(i, wkIdx), 0)
                wsOut.Cells(outRow, 7).Value = Round(weeklyAlloc, 0)

                weeklyShort = demandThis - weeklyAlloc
                If weeklyShort < 0 Then weeklyShort = 0
                wsOut.Cells(outRow, 8).Value = Round(weeklyShort, 0)

                wsOut.Cells(outRow, 9).Value = Round(preDemandTotal(i), 0)
                wsOut.Cells(outRow, 10).Value = Round(preProdTotal(i), 0)
                wsOut.Cells(outRow, 11).Value = Round(totalShort, 0)

                If covWeeks > 0 Then
                    wsOut.Cells(outRow, 12).Value = covWeeks
                    wsOut.Cells(outRow, 13).Value = CDate(covDate)
                    wsOut.Cells(outRow, 13).NumberFormat = "m/d/yyyy"
                Else
                    wsOut.Cells(outRow, 12).Value = ""
                    wsOut.Cells(outRow, 13).Value = ""
                End If

                outRow = outRow + 1
            End If
        Next wkIdx

        ' blank spacer row between SKUs
        outRow = outRow + 1
    Next i

    '====================================================
    ' 8) Format final table (COO-friendly)
    '====================================================
    Dim lastRow As Long
    lastRow = outRow - 1

    If lastRow >= 5 Then
        With wsOut.Range(wsOut.Cells(4, 1), wsOut.Cells(lastRow, 13))
            .Font.name = "Segoe UI"
            .Font.Size = 9
            .Borders.LineStyle = xlContinuous
            .Columns.AutoFit
        End With

        ' Zebra striping on data rows
        With wsOut.Range(wsOut.Cells(5, 1), wsOut.Cells(lastRow, 13)).FormatConditions.Add( _
                Type:=xlExpression, _
                Formula1:="=MOD(ROW(),2)=0")
            .Interior.Color = RGB(245, 247, 250)
        End With

        ' Conditional formatting: weekly shortfall > 0 in light red
        With wsOut.Range(wsOut.Cells(5, 8), wsOut.Cells(lastRow, 8)).FormatConditions.Add( _
                Type:=xlCellValue, _
                Operator:=xlGreater, _
                Formula1:="0")
            .Interior.Color = RGB(255, 220, 220)
        End With

        ' Conditional formatting: total shortfall >0
        With wsOut.Range(wsOut.Cells(5, 11), wsOut.Cells(lastRow, 11)).FormatConditions.Add( _
                Type:=xlCellValue, _
                Operator:=xlGreater, _
                Formula1:="0")
            .Interior.Color = RGB(255, 236, 179)
        End With

        ' Coverage formatting
        With wsOut.Range(wsOut.Cells(5, 12), wsOut.Cells(lastRow, 12)).FormatConditions.Add( _
                Type:=xlCellValue, _
                Operator:=xlLess, _
                Formula1:="4")
            .Interior.Color = RGB(255, 204, 204) ' < 4 weeks = red
        End With
        With wsOut.Range(wsOut.Cells(5, 12), wsOut.Cells(lastRow, 12)).FormatConditions.Add( _
                Type:=xlCellValue, _
                Operator:=xlBetween, _
                Formula1:="4", _
                Formula2:="8")
            .Interior.Color = RGB(255, 250, 205) ' 4–8 weeks = yellow
        End With
        With wsOut.Range(wsOut.Cells(5, 12), wsOut.Cells(lastRow, 12)).FormatConditions.Add( _
                Type:=xlCellValue, _
                Operator:=xlGreater, _
                Formula1:="8")
            .Interior.Color = RGB(225, 235, 245) ' >8 weeks = cool blue
        End With
    End If

CleanExit:
    Application.ScreenUpdating = scrn
    Application.EnableEvents = evt
    Application.Calculation = calc
    Exit Sub

CleanFail:
    Application.ScreenUpdating = scrn
    Application.EnableEvents = evt
    Application.Calculation = calc
    MsgBox "BuildRevoProductionAllocator failed at stage: " & stage & vbCrLf & _
           "Err " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' Helper: safely convert any value to Double (blank/text/error -> 0)
Private Function ToDbl(ByVal v As Variant) As Double
    If IsError(v) Then
        ToDbl = 0#
    ElseIf IsNull(v) Then
        ToDbl = 0#
    ElseIf VarType(v) = vbString Then
        If Trim$(v) = "" Then
            ToDbl = 0#
        Else
            ToDbl = CDbl(val(v))
        End If
    ElseIf v = "" Then
        ToDbl = 0#
    Else
        ToDbl = CDbl(v)
    End If
End Function
Public Sub BuildRevoAllocDashboard()
    Const SH_DATA As String = "REVO Production Allocator"
    Const SH_DASH As String = "REVO Alloc Dashboard"
    Const TBL_NAME As String = "tblRevoAlloc"

    Dim scrn As Boolean, evt As Boolean
    Dim calc As XlCalculation
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    evt = Application.EnableEvents:    Application.EnableEvents = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim wsData As Worksheet, wsDash As Worksheet
    Dim lastRow As Long, lastCol As Long

    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets(SH_DATA)

    '-----------------------------
    ' 1) Make sure data is a Table
    '-----------------------------
    ' Find last row / col of allocator table (headers start row 4)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row
    lastCol = wsData.Cells(4, wsData.Columns.Count).End(xlToLeft).Column
    If lastRow < 5 Or lastCol < 3 Then GoTo CleanExit

    Dim srcRange As Range
    Set srcRange = wsData.Range(wsData.Cells(4, 1), wsData.Cells(lastRow, lastCol))

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsData.ListObjects(TBL_NAME)
    On Error GoTo 0

    If lo Is Nothing Then
        Set lo = wsData.ListObjects.Add(SourceType:=xlSrcRange, Source:=srcRange, _
                                        XlListObjectHasHeaders:=xlYes)
        lo.name = TBL_NAME
    Else
        ' Resize to current data
        lo.Resize srcRange
    End If

    '-----------------------------
    ' 2) Create / reset dashboard sheet
    '-----------------------------
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(SH_DASH).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDash = wb.Worksheets.Add(After:=wsData)
    wsDash.name = SH_DASH

    '-----------------------------
    ' 3) KPI header / cards
    '-----------------------------
    With wsDash.Range("A1")
        .Value = "REVO Production Coverage Dashboard"
        .Font.name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
    End With

    With wsDash.Range("A2")
        .Value = "Source: REVO Production Allocator  |  Last Updated " & _
                 Format(Now, "m/d/yyyy h:mm AM/PM")
        .Font.name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = RGB(90, 90, 90)
    End With

    ' Card 1: Worst coverage (min coverage weeks > 0)
    With wsDash.Range("A4:E6")
        .Merge
        .Interior.Color = RGB(32, 45, 60)
        .Font.Color = RGB(255, 255, 255)
        .Font.name = "Segoe UI"
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    wsDash.Range("A4").Value = "Lowest Coverage (Weeks)"

    wsDash.Range("A7").Formula = _
        "=MIN(IF(tblRevoAlloc[Coverage Weeks]>0,tblRevoAlloc[Coverage Weeks]))"
    wsDash.Range("A7").Font.name = "Segoe UI"
    wsDash.Range("A7").Font.Size = 18
    wsDash.Range("A7").Font.Bold = True

    ' Card 2: Total shortfall across all preassemblies
    With wsDash.Range("G4:K6")
        .Merge
        .Interior.Color = RGB(230, 90, 90)
        .Font.Color = RGB(255, 255, 255)
        .Font.name = "Segoe UI"
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    wsDash.Range("G4").Value = "Total Shortfall (Units)"

    wsDash.Range("G7").Formula = "=SUM(tblRevoAlloc[Total Shortfall])"
    wsDash.Range("G7").Font.name = "Segoe UI"
    wsDash.Range("G7").Font.Size = 18
    wsDash.Range("G7").Font.Bold = True

    ' Card 3: Total demand vs total produced
    With wsDash.Range("M4:Q6")
        .Merge
        .Interior.Color = RGB(32, 120, 60)
        .Font.Color = RGB(255, 255, 255)
        .Font.name = "Segoe UI"
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    wsDash.Range("M4").Value = "Total Demand vs Produced"

    wsDash.Range("M7").Formula = "=SUM(tblRevoAlloc[Total Demand (All Weeks)])"
    wsDash.Range("N7").Formula = "=SUM(tblRevoAlloc[Total Produced])"
    wsDash.Range("M7:N7").Font.name = "Segoe UI"
    wsDash.Range("M7:N7").Font.Size = 11
    wsDash.Range("M7:N7").Font.Bold = True
    wsDash.Range("M6").Value = "Demand"
    wsDash.Range("N6").Value = "Produced"

    '-----------------------------
    ' 4) Pivot: Coverage by Preassembly
    '-----------------------------
    Dim pc As PivotCache
    Dim pt1 As PivotTable
    Dim ptRange As Range

    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, _
                                   SourceData:=lo.Range)

    Set ptRange = wsDash.Range("A10")
    Set pt1 = pc.CreatePivotTable(TableDestination:=ptRange, tableName:="ptCoverage")

    With pt1
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels

        With .PivotFields("Preassembly SKU")
            .Orientation = xlRowField
            .Position = 1
        End With

        .AddDataField .PivotFields("Coverage Weeks"), "Max Coverage Weeks", xlMax
        .AddDataField .PivotFields("Total Shortfall"), "Total Shortfall", xlSum
    End With

    wsDash.Range("A9").Value = "Coverage by Preassembly"
    wsDash.Range("A9").Font.Bold = True
    wsDash.Range("A9").Font.name = "Segoe UI"

    ' Format the pivot a bit
    pt1.RowAxisLayout xlTabularRow
    pt1.ShowTableStyleRowStripes = True
    pt1.TableStyle2 = "PivotStyleMedium9"

    '-----------------------------
    ' 5) Pivot: Weekly Shortfall Heatmap
    '-----------------------------
    Dim pt2 As PivotTable
    Set ptRange = wsDash.Range("J10")
    Set pt2 = pc.CreatePivotTable(TableDestination:=ptRange, tableName:="ptShortfall")

    With pt2
        .RowAxisLayout xlTabularRow

        With .PivotFields("Preassembly SKU")
            .Orientation = xlRowField
            .Position = 1
        End With

        With .PivotFields("Week #")
            .Orientation = xlColumnField
            .Position = 1
        End With

        .AddDataField .PivotFields("Weekly Shortfall"), "Weekly Shortfall", xlSum
    End With

    wsDash.Range("J9").Value = "Weekly Shortfall Heatmap"
    wsDash.Range("J9").Font.Bold = True
    wsDash.Range("J9").Font.name = "Segoe UI"

    pt2.TableStyle2 = "PivotStyleLight16"

    ' Conditional formatting for heatmap
    Dim dataBody As Range
    On Error Resume Next
    Set dataBody = pt2.DataBodyRange
    On Error GoTo 0

    If Not dataBody Is Nothing Then
        Dim fc As FormatCondition
        dataBody.FormatConditions.Delete
        Set fc = dataBody.FormatConditions.AddColorScale(ColorScaleType:=3)
        With fc
            ' low = green, mid = yellow, high = red
            .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .ColorScaleCriteria(1).FormatColor.Color = RGB(225, 235, 245)
            .ColorScaleCriteria(2).Type = xlConditionValuePercentile
            .ColorScaleCriteria(2).Value = 50
            .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 250, 205)
            .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
            .ColorScaleCriteria(3).FormatColor.Color = RGB(255, 204, 204)
        End With
    End If

    ' Autofit & tidy
    wsDash.Columns("A:Q").AutoFit

CleanExit:
    Application.ScreenUpdating = scrn
    Application.EnableEvents = evt
    Application.Calculation = calc
    Exit Sub

CleanFail:
    Application.ScreenUpdating = scrn
    Application.EnableEvents = evt
    Application.Calculation = calc
    MsgBox "BuildRevoAllocDashboard failed: " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit
End Sub


'=========================================================
' Finished SKU -> Preassembly requirements (stop-walk at PREASSEMBLY)
' Returns Dictionary: {PreAsmSKU -> qtyPerFinished}
'=========================================================
Private Function FlattenToPreAssemblies_ForFinished(ByVal finishedSku As String, _
                                                    ByVal children As Object, _
                                                    ByVal preAsm As Object) As Object
    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    out.CompareMode = vbTextCompare

    FlattenWalkToPreAsm KeyNorm(finishedSku), 1#, children, preAsm, out

    Set FlattenToPreAssemblies_ForFinished = out
End Function

Private Sub FlattenWalkToPreAsm(ByVal sku As String, ByVal mult As Double, _
                                ByVal children As Object, ByVal preAsm As Object, _
                                ByRef out As Object)
    If Len(sku) = 0 Then Exit Sub

    ' If THIS node is a preassembly, stop and record it
    If preAsm.exists(sku) Then
        AddQty out, sku, mult
        Exit Sub
    End If

    ' Otherwise walk down to reach preassemblies
    If Not children.exists(sku) Then Exit Sub

    Dim col As Collection
    Set col = children(sku)

    Dim i As Long, compArr As Variant
    Dim childSku As String, qPer As Double

    For i = 1 To col.Count
        compArr = col(i)                 ' Array(childSku, qPer)
        childSku = KeyNorm(compArr(0))
        qPer = SafeDbl(compArr(1))
        If qPer = 0 Then qPer = 1#
        If Len(childSku) > 0 Then
            FlattenWalkToPreAsm childSku, mult * qPer, children, preAsm, out
        End If
    Next i
End Sub
