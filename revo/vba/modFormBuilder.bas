Attribute VB_Name = "modFormBuilder"
Option Explicit

'======================================================================
' modFormBuilder  -  builds frmDisposition
'
' REQUIRES: File > Options > Trust Center > Trust Center Settings >
'           Macro Settings > tick "Trust access to the VBA project
'           object model". One time, then this runs.
'
' Run BuildDispositionForm once. It creates a real designer UserForm
' you can open in the VBE and tweak by hand afterwards. Delete and
' re-run to reset it.
'
' One form serves both flows. Caller sets Mode:
'   frmDisposition.Prime "NEW",    sku, cart, wo, qty
'   frmDisposition.Prime "REWORK", sku, cart, wo, qty
'
' ShowDispositionTest is the smoke test - it primes the form with a
' dummy cart and reports what came back, without writing to any log
' unless you actually enter a disposition quantity.
'
' Changes from the first draft:
'   - Cosmetic UserForm properties (Width/Height/BackColor) are set
'     through the Properties collection, which throws on some Excel
'     builds. Those are now individually guarded so a failed BackColor
'     no longer aborts the whole build.
'   - The three combo boxes were Style 0 (free text). Op codes have to
'     match IC/ASSYM/10/.../FQC exactly or the escape analysis silently
'     buckets typos as separate operations. All three are now
'     fmStyleDropDownList with MatchRequired.
'   - AddCombo now takes the list style so the choice is visible at the
'     call site instead of buried in the helper.
'======================================================================

Private Const FRM As String = "frmDisposition"

Private Const DQ As String = """"

' MSForms constants, spelled out so this compiles with or without the
' MSForms library referenced.
Private Const STYLE_COMBO As Long = 0       ' fmStyleDropDownCombo - free text
Private Const STYLE_LIST As Long = 2        ' fmStyleDropDownList  - pick only

Public Sub BuildDispositionForm()

    Dim vbc As Object
    Dim d As Object

    On Error GoTo Fail

    RemoveIfPresent FRM

    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(3)   ' 3 = vbext_ct_MSForm
    vbc.Name = FRM
    Set d = vbc.Designer

    SetProp vbc, "Caption", "REVO  -  Release / Disposition"
    SetProp vbc, "Width", 420
    SetProp vbc, "Height", 470
    SetProp vbc, "BackColor", RGB(245, 246, 248)

    ' ---------- header band ----------
    AddLabel d, "lblBand", 0, 0, 414, 46, "", RGB(255, 255, 255), RGB(68, 84, 106), 8, False
    AddLabel d, "lblTitle", 14, 8, 320, 18, "RELEASE / DISPOSITION", RGB(255, 255, 255), RGB(68, 84, 106), 11, True
    AddLabel d, "lblMode", 14, 26, 320, 14, "", RGB(198, 210, 226), RGB(68, 84, 106), 8, False

    ' ---------- context strip ----------
    AddLabel d, "lblSKUCap", 14, 58, 40, 12, "SKU", RGB(110, 118, 130), -1, 8, False
    AddLabel d, "lblSKU", 14, 71, 250, 14, "-", RGB(32, 38, 48), -1, 10, True
    AddLabel d, "lblCartCap", 280, 58, 40, 12, "CART", RGB(110, 118, 130), -1, 8, False
    AddLabel d, "lblCart", 280, 71, 120, 14, "-", RGB(32, 38, 48), -1, 10, True
    AddLabel d, "lblWOCap", 14, 92, 40, 12, "WO", RGB(110, 118, 130), -1, 8, False
    AddLabel d, "lblWO", 14, 105, 250, 14, "-", RGB(32, 38, 48), -1, 9, False
    AddLabel d, "lblQtyCap", 280, 92, 90, 12, "QTY ON CART", RGB(110, 118, 130), -1, 8, False
    AddLabel d, "lblQty", 280, 105, 120, 14, "0", RGB(32, 38, 48), -1, 9, False

    AddLabel d, "lblRule1", 14, 126, 386, 1, "", -1, RGB(214, 219, 226), 8, False

    ' ---------- quantities ----------
    AddLabel d, "lblQtyProc", 14, 138, 130, 14, "Qty to process", RGB(32, 38, 48), -1, 9, True
    AddText d, "txtQty", 150, 136, 70, 20

    AddLabel d, "lblRework", 14, 168, 130, 14, "Rework", RGB(32, 38, 48), -1, 9, False
    AddText d, "txtRework", 150, 166, 70, 20
    AddLabel d, "lblBGrade", 14, 194, 130, 14, "B-Grade", RGB(32, 38, 48), -1, 9, False
    AddText d, "txtBGrade", 150, 192, 70, 20
    AddLabel d, "lblReject", 14, 220, 130, 14, "Reject", RGB(32, 38, 48), -1, 9, False
    AddText d, "txtReject", 150, 218, 70, 20

    ' ---------- released tile ----------
    AddLabel d, "lblRelTile", 250, 138, 150, 74, "", -1, RGB(226, 240, 228), 8, False
    AddLabel d, "lblRelCap", 262, 148, 130, 12, "RELEASED TO INVENTORY", RGB(60, 100, 70), RGB(226, 240, 228), 7, True
    AddLabel d, "lblReleased", 262, 162, 130, 30, "0", RGB(34, 90, 50), RGB(226, 240, 228), 22, True
    AddLabel d, "lblRemain", 262, 194, 130, 12, "Remaining on cart: 0", RGB(60, 100, 70), RGB(226, 240, 228), 7, False

    ' ---------- defect capture ----------
    AddLabel d, "lblRule2", 14, 248, 386, 1, "", -1, RGB(214, 219, 226), 8, False
    AddLabel d, "lblDefHdr", 14, 258, 386, 14, "DEFECT DETAIL  (required when any disposition is entered)", RGB(150, 60, 60), -1, 8, True

    AddLabel d, "lblDefCode", 14, 280, 130, 14, "Defect code", RGB(32, 38, 48), -1, 9, False
    AddCombo d, "cboDefect", 150, 278, 250, 20, STYLE_LIST
    AddLabel d, "lblOpDet", 14, 306, 130, 14, "Op detected", RGB(32, 38, 48), -1, 9, False
    AddCombo d, "cboOpDetected", 150, 304, 110, 20, STYLE_LIST
    AddLabel d, "lblOpCau", 14, 332, 130, 14, "Op caused", RGB(32, 38, 48), -1, 9, False
    AddCombo d, "cboOpCaused", 150, 330, 110, 20, STYLE_LIST
    AddLabel d, "lblNotes", 14, 358, 130, 14, "Notes", RGB(32, 38, 48), -1, 9, False
    AddText d, "txtNotes", 150, 356, 250, 20

    AddLabel d, "lblEscape", 270, 306, 130, 40, "", RGB(150, 60, 60), -1, 8, True

    ' ---------- buttons ----------
    AddButton d, "btnSubmit", 250, 392, 150, 30, "SUBMIT RELEASE"
    AddButton d, "btnCancel", 150, 392, 90, 30, "Cancel"

    vbc.CodeModule.AddFromString FormCode()

    LogLine "BuildDispositionForm: " & FRM & " built with " & d.Controls.Count & " controls"
    Say FRM & " built." & vbCrLf & vbCrLf & _
        "Smoke test it with:  ShowDispositionTest" & vbCrLf & vbCrLf & _
        "Call it for real with:" & vbCrLf & _
        "  frmDisposition.Prime " & DQ & "NEW" & DQ & ", sku, cart, wo, qty" & vbCrLf & _
        "  frmDisposition.Show" & vbCrLf & _
        "  If Not frmDisposition.UserCancelled Then ... read properties", "REVO"
    Exit Sub

Fail:
    LogLine "*** BuildDispositionForm failed: " & Err.Number & " " & Err.Description
    If Err.Number = 1004 Or InStr(1, Err.Description, "programmatic", vbTextCompare) > 0 Then
        Say "Enable this first:" & vbCrLf & vbCrLf & _
            "File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
            "Macro Settings > Trust access to the VBA project object model", _
            "One-time setting", vbExclamation
    Else
        Say "Build failed: " & Err.Description, "REVO", vbCritical
    End If
End Sub

'----------------------------------------------------------------------
' Smoke test - Phase 3 acceptance. Primes the form with a dummy cart,
' shows it, and reports what came back.
'----------------------------------------------------------------------
Public Sub ShowDispositionTest()

    Dim before As Long, after As Long
    Dim ws As Worksheet

    On Error GoTo Fail

    If Not SheetExists(SH_DEFECTS) Then
        Say SH_DEFECTS & " does not exist. Run QualityBuild first.", "REVO", vbExclamation
        Exit Sub
    End If

    Set ws = GetSheet(SH_QLOG)
    If Not ws Is Nothing Then before = LastRow(ws)

    frmDisposition.Prime "NEW", "S PRE REVO TEST", "9999", "WO-TEST", 100
    frmDisposition.Show

    If frmDisposition.UserCancelled Then
        Say "Cancelled - nothing written.", "REVO"
    Else
        If Not ws Is Nothing Then after = LastRow(ws)
        Say "Submitted." & vbCrLf & vbCrLf & _
            "Qty to process:  " & frmDisposition.RequestedQty & vbCrLf & _
            "Released:        " & frmDisposition.ReleasedToInventory & vbCrLf & _
            "Rework/BG/Rej:   " & frmDisposition.Rework & " / " & frmDisposition.BGrade & " / " & frmDisposition.Reject & vbCrLf & _
            "Defect code:     " & frmDisposition.DefectCode & vbCrLf & _
            "Caused / caught: " & frmDisposition.OpCaused & " / " & frmDisposition.OpDetected & vbCrLf & vbCrLf & _
            SH_QLOG & " rows written: " & (after - before), "REVO"
    End If

    Unload frmDisposition
    Exit Sub
Fail:
    Say "Smoke test failed: " & Err.Description & vbCrLf & vbCrLf & _
        "If this says 'variable not defined', BuildDispositionForm has not run yet.", _
        "REVO", vbCritical
End Sub

'======================================================================
' CONTROL HELPERS
'======================================================================

' Cosmetic properties are individually guarded: some Excel builds refuse
' Width/Height/BackColor through the Properties collection, and losing
' the whole form over a background colour is not a trade worth making.
Private Sub SetProp(ByVal vbc As Object, ByVal nm As String, ByVal v As Variant)
    On Error Resume Next
    vbc.Properties.Item(nm).Value = v
    If Err.Number <> 0 Then LogLine "  form property '" & nm & "' not set: " & Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub AddLabel(d As Object, nm As String, l As Single, t As Single, w As Single, h As Single, _
                     cap As String, foreC As Long, backC As Long, sz As Single, bold As Boolean)
    Dim c As Object
    Set c = d.Controls.Add("Forms.Label.1", nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Caption = cap
    If foreC >= 0 Then c.ForeColor = foreC
    If backC >= 0 Then
        c.BackStyle = 1
        c.BackColor = backC
    Else
        c.BackStyle = 0
    End If
    c.Font.Size = sz
    c.Font.bold = bold
End Sub

Private Sub AddText(d As Object, nm As String, l As Single, t As Single, w As Single, h As Single)
    Dim c As Object
    Set c = d.Controls.Add("Forms.TextBox.1", nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = 10
    c.BorderStyle = 1
End Sub

Private Sub AddCombo(d As Object, nm As String, l As Single, t As Single, w As Single, h As Single, _
                     Optional ByVal listStyle As Long = STYLE_LIST)
    Dim c As Object
    Set c = d.Controls.Add("Forms.ComboBox.1", nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = 9
    c.Style = listStyle
    c.MatchRequired = (listStyle = STYLE_LIST)
End Sub

Private Sub AddButton(d As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String)
    Dim c As Object
    Set c = d.Controls.Add("Forms.CommandButton.1", nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Caption = cap
    c.Font.Size = 9
    c.Font.bold = True
End Sub

Private Sub RemoveIfPresent(ByVal nm As String)
    Dim vbc As Object
    On Error Resume Next
    Set vbc = ThisWorkbook.VBProject.VBComponents(nm)
    If Not vbc Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove vbc
    Err.Clear
    On Error GoTo 0
End Sub

'======================================================================
' THE FORM'S OWN CODE  (injected)
'======================================================================
Private Function FormCode() As String

    Dim s As String
    Dim NL As String
    NL = vbCrLf

    s = "Option Explicit" & NL & NL
    s = s & "Public UserCancelled As Boolean" & NL
    s = s & "Public Mode As String" & NL
    s = s & "Public SKU As String" & NL
    s = s & "Public Cart As String" & NL
    s = s & "Public WO As String" & NL
    s = s & "Public TotalQty As Long" & NL
    s = s & "Private mLoading As Boolean" & NL & NL

    ' ---- Prime ----
    s = s & "Public Sub Prime(ByVal pMode As String, ByVal pSKU As String, ByVal pCart As String, ByVal pWO As String, ByVal pQty As Long)" & NL
    s = s & "    mLoading = True" & NL
    s = s & "    UserCancelled = True" & NL
    s = s & "    Mode = UCase$(pMode): SKU = pSKU: Cart = pCart: WO = pWO: TotalQty = pQty" & NL
    s = s & "    lblSKU.Caption = pSKU" & NL
    s = s & "    lblCart.Caption = pCart" & NL
    s = s & "    lblWO.Caption = pWO" & NL
    s = s & "    lblQty.Caption = CStr(pQty)" & NL
    s = s & "    If Mode = " & DQ & "REWORK" & DQ & " Then" & NL
    s = s & "        lblTitle.Caption = " & DQ & "REWORK RELEASE" & DQ & NL
    s = s & "        lblMode.Caption = " & DQ & "Returning reworked units to inventory" & DQ & NL
    s = s & "        lblRework.Enabled = False: txtRework.Enabled = False: txtRework.Value = " & DQ & "0" & DQ & NL
    s = s & "    Else" & NL
    s = s & "        lblTitle.Caption = " & DQ & "NEW PRODUCTION RELEASE" & DQ & NL
    s = s & "        lblMode.Caption = " & DQ & "Releasing from the floor to inventory" & DQ & NL
    s = s & "        lblRework.Enabled = True: txtRework.Enabled = True" & NL
    s = s & "    End If" & NL
    s = s & "    txtQty.Value = CStr(pQty)" & NL
    s = s & "    txtRework.Value = " & DQ & "0" & DQ & ": txtBGrade.Value = " & DQ & "0" & DQ & ": txtReject.Value = " & DQ & "0" & DQ & NL
    s = s & "    txtNotes.Value = vbNullString" & NL
    s = s & "    LoadLists" & NL
    s = s & "    mLoading = False" & NL
    s = s & "    Recalc" & NL
    s = s & "End Sub" & NL & NL

    ' ---- LoadLists ----
    s = s & "Private Sub LoadLists()" & NL
    s = s & "    Dim v As Variant, i As Long, ops As Variant" & NL
    s = s & "    cboDefect.Clear: cboOpDetected.Clear: cboOpCaused.Clear" & NL
    s = s & "    v = QualityDefectCodes()" & NL
    s = s & "    If Not IsEmpty(v) Then" & NL
    s = s & "        For i = LBound(v) To UBound(v)" & NL
    s = s & "            cboDefect.AddItem v(i)" & NL
    s = s & "        Next i" & NL
    s = s & "    End If" & NL
    s = s & "    ops = OpsArray()" & NL
    s = s & "    For i = LBound(ops) To UBound(ops)" & NL
    s = s & "        cboOpDetected.AddItem ops(i)" & NL
    s = s & "        cboOpCaused.AddItem ops(i)" & NL
    s = s & "    Next i" & NL
    s = s & "    On Error Resume Next" & NL
    s = s & "    cboOpDetected.Value = " & DQ & "FQC" & DQ & NL
    s = s & "    On Error GoTo 0" & NL
    s = s & "End Sub" & NL & NL

    ' ---- properties ----
    s = s & "Public Property Get RequestedQty() As Long" & NL & "    RequestedQty = SafeLng(txtQty.Value)" & NL & "End Property" & NL & NL
    s = s & "Public Property Get Rework() As Long" & NL & "    Rework = SafeLng(txtRework.Value)" & NL & "End Property" & NL & NL
    s = s & "Public Property Get BGrade() As Long" & NL & "    BGrade = SafeLng(txtBGrade.Value)" & NL & "End Property" & NL & NL
    s = s & "Public Property Get Reject() As Long" & NL & "    Reject = SafeLng(txtReject.Value)" & NL & "End Property" & NL & NL
    s = s & "Public Property Get DefectCode() As String" & NL & "    DefectCode = QualityBareCode(CStr(cboDefect.Value & " & DQ & DQ & "))" & NL & "End Property" & NL & NL
    s = s & "Public Property Get OpDetected() As String" & NL & "    OpDetected = CStr(cboOpDetected.Value & " & DQ & DQ & ")" & NL & "End Property" & NL & NL
    s = s & "Public Property Get OpCaused() As String" & NL & "    OpCaused = CStr(cboOpCaused.Value & " & DQ & DQ & ")" & NL & "End Property" & NL & NL
    s = s & "Public Property Get Notes() As String" & NL & "    Notes = Trim$(CStr(txtNotes.Value))" & NL & "End Property" & NL & NL
    s = s & "Public Property Get CartRemaining() As Long" & NL
    s = s & "    CartRemaining = TotalQty - RequestedQty" & NL
    s = s & "    If CartRemaining < 0 Then CartRemaining = 0" & NL
    s = s & "End Property" & NL & NL
    s = s & "Public Property Get ReleasedToInventory() As Long" & NL
    s = s & "    Dim v As Long" & NL
    s = s & "    v = RequestedQty - Rework - BGrade - Reject" & NL
    s = s & "    If v < 0 Then v = 0" & NL
    s = s & "    ReleasedToInventory = v" & NL
    s = s & "End Property" & NL & NL
    s = s & "Public Property Get TotalDispositioned() As Long" & NL
    s = s & "    TotalDispositioned = Rework + BGrade + Reject" & NL
    s = s & "End Property" & NL & NL

    ' ---- recalc ----
    s = s & "Private Sub Recalc()" & NL
    s = s & "    If mLoading Then Exit Sub" & NL
    s = s & "    Dim needDefect As Boolean" & NL
    s = s & "    lblReleased.Caption = CStr(ReleasedToInventory)" & NL
    s = s & "    lblRemain.Caption = " & DQ & "Remaining on cart: " & DQ & " & CartRemaining" & NL
    s = s & "    needDefect = (TotalDispositioned > 0)" & NL
    s = s & "    lblDefHdr.Visible = needDefect" & NL
    s = s & "    lblDefCode.Visible = needDefect: cboDefect.Visible = needDefect" & NL
    s = s & "    lblOpDet.Visible = needDefect: cboOpDetected.Visible = needDefect" & NL
    s = s & "    lblOpCau.Visible = needDefect: cboOpCaused.Visible = needDefect" & NL
    s = s & "    lblNotes.Visible = needDefect: txtNotes.Visible = needDefect" & NL
    s = s & "    lblEscape.Visible = needDefect" & NL
    s = s & "    If needDefect And Len(OpDetected) > 0 And Len(OpCaused) > 0 Then" & NL
    s = s & "        If OpDetected <> OpCaused Then" & NL
    s = s & "            lblEscape.Caption = " & DQ & "ESCAPE - made at " & DQ & " & OpCaused & " & DQ & ", caught at " & DQ & " & OpDetected" & NL
    s = s & "        Else" & NL
    s = s & "            lblEscape.Caption = " & DQ & "Caught at source" & DQ & NL
    s = s & "        End If" & NL
    s = s & "    Else" & NL
    s = s & "        lblEscape.Caption = vbNullString" & NL
    s = s & "    End If" & NL
    s = s & "End Sub" & NL & NL

    ' ---- events ----
    s = s & "Private Sub UserForm_Initialize()" & NL & "    UserCancelled = True" & NL & "End Sub" & NL & NL
    s = s & "Private Sub txtQty_Change()" & NL & "    Recalc" & NL & "End Sub" & NL & NL
    s = s & "Private Sub txtRework_Change()" & NL & "    Recalc" & NL & "End Sub" & NL & NL
    s = s & "Private Sub txtBGrade_Change()" & NL & "    Recalc" & NL & "End Sub" & NL & NL
    s = s & "Private Sub txtReject_Change()" & NL & "    Recalc" & NL & "End Sub" & NL & NL
    s = s & "Private Sub cboOpDetected_Change()" & NL & "    Recalc" & NL & "End Sub" & NL & NL
    s = s & "Private Sub cboOpCaused_Change()" & NL & "    Recalc" & NL & "End Sub" & NL & NL

    ' ---- submit ----
    s = s & "Private Sub btnSubmit_Click()" & NL
    s = s & "    Dim q As Long" & NL
    s = s & "    q = RequestedQty" & NL
    s = s & "    If q <= 0 Then" & NL
    s = s & "        MsgBox " & DQ & "Enter a quantity greater than zero." & DQ & ", vbExclamation" & NL
    s = s & "        txtQty.SetFocus: Exit Sub" & NL
    s = s & "    End If" & NL
    s = s & "    If q > TotalQty Then" & NL
    s = s & "        MsgBox " & DQ & "Qty to process cannot exceed the cart qty (" & DQ & " & TotalQty & " & DQ & ")." & DQ & ", vbExclamation" & NL
    s = s & "        txtQty.SetFocus: Exit Sub" & NL
    s = s & "    End If" & NL
    s = s & "    If Rework < 0 Or BGrade < 0 Or Reject < 0 Then" & NL
    s = s & "        MsgBox " & DQ & "Negative values are not allowed." & DQ & ", vbExclamation: Exit Sub" & NL
    s = s & "    End If" & NL
    s = s & "    If TotalDispositioned > q Then" & NL
    s = s & "        MsgBox " & DQ & "Rework + B-Grade + Reject cannot exceed the qty being processed." & DQ & ", vbExclamation: Exit Sub" & NL
    s = s & "    End If" & NL
    s = s & "    If TotalDispositioned > 0 Then" & NL
    s = s & "        If Len(DefectCode) = 0 Then" & NL
    s = s & "            MsgBox " & DQ & "Select a defect code." & DQ & ", vbExclamation" & NL
    s = s & "            cboDefect.SetFocus: Exit Sub" & NL
    s = s & "        End If" & NL
    s = s & "        If Len(OpDetected) = 0 Or Len(OpCaused) = 0 Then" & NL
    s = s & "            MsgBox " & DQ & "Select both the op it was caught at and the op that caused it." & DQ & ", vbExclamation" & NL
    s = s & "            cboOpCaused.SetFocus: Exit Sub" & NL
    s = s & "        End If" & NL
    s = s & "        WriteDefects" & NL
    s = s & "    End If" & NL
    s = s & "    UserCancelled = False" & NL
    s = s & "    Me.Hide" & NL
    s = s & "End Sub" & NL & NL

    ' ---- write ----
    s = s & "Private Sub WriteDefects()" & NL
    s = s & "    Dim sh As String, who As String, d As Date" & NL
    s = s & "    d = Date" & NL
    s = s & "    who = Application.UserName" & NL
    s = s & "    If Hour(Now) >= 14 Or Hour(Now) < 5 Then sh = " & DQ & "Night" & DQ & " Else sh = " & DQ & "Day" & DQ & NL
    s = s & "    If Rework > 0 Then QualityLogDefect d, sh, Cart, WO, SKU, Rework, " & DQ & "Rework" & DQ & ", OpDetected, OpCaused, DefectCode, Notes, who" & NL
    s = s & "    If BGrade > 0 Then QualityLogDefect d, sh, Cart, WO, SKU, BGrade, " & DQ & "B-Grade" & DQ & ", OpDetected, OpCaused, DefectCode, Notes, who" & NL
    s = s & "    If Reject > 0 Then QualityLogDefect d, sh, Cart, WO, SKU, Reject, " & DQ & "Reject" & DQ & ", OpDetected, OpCaused, DefectCode, Notes, who" & NL
    s = s & "End Sub" & NL & NL

    s = s & "Private Sub btnCancel_Click()" & NL & "    UserCancelled = True" & NL & "    Me.Hide" & NL & "End Sub" & NL & NL
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & NL
    s = s & "    If CloseMode = vbFormControlMenu Then" & NL
    s = s & "        Cancel = True: UserCancelled = True: Me.Hide" & NL
    s = s & "    End If" & NL
    s = s & "End Sub" & NL & NL
    s = s & "Private Function SafeLng(ByVal v As Variant) As Long" & NL
    s = s & "    If IsNumeric(v) Then SafeLng = CLng(Val(CStr(v)))" & NL
    s = s & "End Function" & NL

    FormCode = s

End Function
