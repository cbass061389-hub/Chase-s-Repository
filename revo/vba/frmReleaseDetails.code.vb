'==============================================================================
' frmReleaseDetails  -  code-behind
'
' This file is the source of truth for the form's code. REVO_FormBuilder
' injects it when it builds the form, so edit it here and rebuild rather than
' editing in the VBE and losing the change on the next build.
'
' The form captures a cart release plus a coded quality disposition for each
' of rework, B grade and reject. Dropdowns are populated from the
' Quality_Codes sheet at run time - adding a defect type is a sheet edit.
'
' NOTE on the mLoading guard: UserForm_Initialize deliberately does not touch
' mLoading. The first reference to Me.Controls from Prime is what triggers
' Initialize, so an Initialize that cleared the flag would switch the guard off
' half way through Prime and let every _Change handler fire during setup. That
' was a live bug in the original form.
'==============================================================================
Option Explicit

Public UserCancelled As Boolean

Private mLoading   As Boolean
Private mSku       As String
Private mCart      As String
Private mWO        As String
Private mRemaining As Long      ' units still on the cart - the cap for this pass
Private mCartQty   As Long      ' full cart quantity, for display
Private mPriorRel  As Long      ' released on earlier passes, for display

'==================== PRIME ====================================================
Public Sub Prime(ByVal sSku As String, ByVal sCart As String, ByVal sWO As String, _
                 ByVal lRemaining As Long, ByVal lCartQty As Long, ByVal lPriorRel As Long)
    mLoading = True

    mSku = sSku
    mCart = sCart
    mWO = sWO
    mRemaining = lRemaining
    mCartQty = lCartQty
    mPriorRel = lPriorRel

    UserCancelled = True          ' nothing commits unless Submit is pressed

    Me.Caption = "Release cart " & sCart & "  -  " & sSku

    SetCap "lblSKU", sSku
    SetCap "lblCart", sCart
    SetCap "lblWO", sWO
    SetCap "lblCartQty", CStr(lCartQty)
    SetCap "lblPriorRel", CStr(lPriorRel)
    SetCap "lblRemaining", CStr(lRemaining)

    FillCombo Me.cboReworkDefect, REVO_Quality.DefectTypeList()
    FillCombo Me.cboReworkLocation, REVO_Quality.LocationList()
    FillCombo Me.cboReworkOp, REVO_Quality.OpList()
    FillCombo Me.cboReworkRoot, REVO_Quality.RootCauseList()
    FillCombo Me.cboReworkAction, REVO_Quality.ActionList()

    FillCombo Me.cboRejectDefect, REVO_Quality.DefectTypeList()
    FillCombo Me.cboRejectLocation, REVO_Quality.LocationList()
    FillCombo Me.cboRejectOp, REVO_Quality.OpList()
    FillCombo Me.cboRejectRoot, REVO_Quality.RootCauseList()
    FillCombo Me.cboRejectAction, REVO_Quality.ActionList()

    FillCombo Me.cboBGradeDefect, REVO_Quality.DefectTypeList()
    FillCombo Me.cboBGradeLocation, REVO_Quality.LocationList()
    FillCombo Me.cboBGradeRoot, REVO_Quality.RootCauseList()

    Me.txtCartQtyRelease.Value = lRemaining     ' default: release what is left
    Me.txtRework.Value = 0
    Me.txtBGrade.Value = 0
    Me.txtReject.Value = 0
    Me.txtReworkNotes.Value = vbNullString
    Me.txtRejectNotes.Value = vbNullString
    Me.txtBGradeNotes.Value = vbNullString

    mLoading = False

    SyncPanels
    UpdateReleaseAmount
End Sub

'==================== PROPERTIES ==============================================
Public Property Get RequestedQty() As Long
    RequestedQty = SafeLong(Me.txtCartQtyRelease.Value)
End Property

Public Property Get Rework() As Long
    Rework = SafeLong(Me.txtRework.Value)
End Property

Public Property Get BGrade() As Long
    BGrade = SafeLong(Me.txtBGrade.Value)
End Property

Public Property Get Reject() As Long
    Reject = SafeLong(Me.txtReject.Value)
End Property

Public Property Get ReleasedToInventory() As Long
    Dim v As Long
    v = RequestedQty - Rework - BGrade - Reject
    If v < 0 Then v = 0
    ReleasedToInventory = v
End Property

Public Property Get CartRemainingAfter() As Long
    Dim v As Long
    v = mRemaining - RequestedQty
    If v < 0 Then v = 0
    CartRemainingAfter = v
End Property

Public Property Get ReworkDefect() As String
    ReworkDefect = ComboText(Me.cboReworkDefect)
End Property
Public Property Get ReworkLocation() As String
    ReworkLocation = ComboText(Me.cboReworkLocation)
End Property
Public Property Get ReworkOp() As String
    ReworkOp = ComboText(Me.cboReworkOp)
End Property
Public Property Get ReworkRootCause() As String
    ReworkRootCause = ComboText(Me.cboReworkRoot)
End Property
Public Property Get ReworkAction() As String
    ReworkAction = ComboText(Me.cboReworkAction)
End Property
Public Property Get ReworkNotes() As String
    ReworkNotes = Trim$(CStr(Me.txtReworkNotes.Value))
End Property

Public Property Get RejectDefect() As String
    RejectDefect = ComboText(Me.cboRejectDefect)
End Property
Public Property Get RejectLocation() As String
    RejectLocation = ComboText(Me.cboRejectLocation)
End Property
Public Property Get RejectOp() As String
    RejectOp = ComboText(Me.cboRejectOp)
End Property
Public Property Get RejectRootCause() As String
    RejectRootCause = ComboText(Me.cboRejectRoot)
End Property
Public Property Get RejectAction() As String
    RejectAction = ComboText(Me.cboRejectAction)
End Property
Public Property Get RejectNotes() As String
    RejectNotes = Trim$(CStr(Me.txtRejectNotes.Value))
End Property

Public Property Get BGradeDefect() As String
    BGradeDefect = ComboText(Me.cboBGradeDefect)
End Property
Public Property Get BGradeLocation() As String
    BGradeLocation = ComboText(Me.cboBGradeLocation)
End Property
Public Property Get BGradeRootCause() As String
    BGradeRootCause = ComboText(Me.cboBGradeRoot)
End Property
Public Property Get BGradeNotes() As String
    BGradeNotes = Trim$(CStr(Me.txtBGradeNotes.Value))
End Property

'==================== UI ======================================================
Private Sub UserForm_Initialize()
    ' Deliberately minimal - see the header note about mLoading.
    UserCancelled = True
End Sub

Private Sub UpdateReleaseAmount()
    Dim req As Long, rel As Long, disp As Long

    If mLoading Then Exit Sub

    req = RequestedQty
    disp = Rework + BGrade + Reject
    rel = req - disp
    If rel < 0 Then rel = 0

    Me.lblToInventory.Caption = CStr(rel)
    Me.lblAfterThisPass.Caption = "Left on cart after this pass: " & CStr(mRemaining - req)

    ' Live warnings rather than a surprise at Submit.
    If req > mRemaining Then
        Me.lblWarn.Caption = "Release qty exceeds the " & mRemaining & " left on this cart."
        Me.lblWarn.ForeColor = RGB(192, 0, 0)
    ElseIf disp > req Then
        Me.lblWarn.Caption = "Rework + B Grade + Reject (" & disp & ") exceeds the release qty (" & req & ")."
        Me.lblWarn.ForeColor = RGB(192, 0, 0)
    ElseIf req <= 0 Then
        Me.lblWarn.Caption = "Enter a release quantity."
        Me.lblWarn.ForeColor = RGB(192, 0, 0)
    Else
        Me.lblWarn.Caption = ""
    End If
End Sub

' Quality panels are enabled only when their quantity is non-zero, so the
' operator is never asked for a root cause on a disposition of zero.
Private Sub SyncPanels()
    If mLoading Then Exit Sub
    EnablePanel Me.fraRework, (Rework > 0)
    EnablePanel Me.fraBGrade, (BGrade > 0)
    EnablePanel Me.fraReject, (Reject > 0)
End Sub

Private Sub EnablePanel(ByVal fra As MSForms.Frame, ByVal on_ As Boolean)
    Dim c As MSForms.Control
    fra.Enabled = on_
    fra.ForeColor = IIf(on_, RGB(0, 0, 0), RGB(150, 150, 150))
    For Each c In fra.Controls
        c.Enabled = on_
    Next c
End Sub

Private Sub txtCartQtyRelease_Change()
    UpdateReleaseAmount
End Sub

Private Sub txtRework_Change()
    SyncPanels
    UpdateReleaseAmount
End Sub

Private Sub txtBGrade_Change()
    SyncPanels
    UpdateReleaseAmount
End Sub

Private Sub txtReject_Change()
    SyncPanels
    UpdateReleaseAmount
End Sub

'==================== ACTIONS =================================================
Private Sub btnSubmit_Click()
    Dim req As Long, disp As Long

    req = RequestedQty
    disp = Rework + BGrade + Reject

    If req <= 0 Then
        Complain "Enter a release quantity greater than zero.", Me.txtCartQtyRelease
        Exit Sub
    End If

    If req > mRemaining Then
        Complain "Release qty (" & req & ") is more than the " & mRemaining & _
                 " still on this cart.", Me.txtCartQtyRelease
        Exit Sub
    End If

    If disp > req Then
        Complain "Rework + B Grade + Reject (" & disp & ") cannot exceed the release qty (" & req & ").", _
                 Me.txtRework
        Exit Sub
    End If

    '--- quality capture is mandatory once a disposition is non-zero --------
    If Rework > 0 Then
        If Len(ReworkDefect) = 0 Then
            Complain "Pick a defect type for the " & Rework & " going to rework.", Me.cboReworkDefect
            Exit Sub
        End If
        If Len(ReworkRootCause) = 0 Then
            Complain "Pick a root cause for the rework." & vbCrLf & vbCrLf & _
                     "If it genuinely is not known yet, choose 'Not Determined' - that is a " & _
                     "reportable answer and a blank is not.", Me.cboReworkRoot
            Exit Sub
        End If
    End If

    If Reject > 0 Then
        If Len(RejectDefect) = 0 Then
            Complain "Pick a defect type for the " & Reject & " being rejected.", Me.cboRejectDefect
            Exit Sub
        End If
        If Len(RejectRootCause) = 0 Then
            Complain "Pick a root cause for the reject." & vbCrLf & vbCrLf & _
                     "If it genuinely is not known yet, choose 'Not Determined'.", Me.cboRejectRoot
            Exit Sub
        End If
    End If

    If BGrade > 0 Then
        If Len(BGradeDefect) = 0 Then
            Complain "Pick a defect type for the " & BGrade & " going to B grade.", Me.cboBGradeDefect
            Exit Sub
        End If
    End If

    UserCancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    UserCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        UserCancelled = True
        Me.Hide
    End If
End Sub

'==================== HELPERS =================================================
Private Sub Complain(ByVal msg As String, ByVal ctl As Object)
    MsgBox msg, vbExclamation, "Check this first"
    On Error Resume Next
    ctl.SetFocus
    On Error GoTo 0
End Sub

Private Sub SetCap(ByVal ctlName As String, ByVal txt As String)
    On Error Resume Next
    Me.Controls(ctlName).Caption = txt
    On Error GoTo 0
End Sub

Private Sub FillCombo(ByVal cbo As MSForms.ComboBox, ByRef items() As String)
    Dim i As Long
    On Error Resume Next
    cbo.Clear
    For i = LBound(items) To UBound(items)
        If Len(items(i)) > 0 Then cbo.AddItem items(i)
    Next i
    cbo.Value = vbNullString
    On Error GoTo 0
End Sub

Private Function ComboText(ByVal cbo As MSForms.ComboBox) As String
    On Error Resume Next
    ComboText = Trim$(CStr(cbo.Value))
    On Error GoTo 0
End Function

Private Function SafeLong(ByVal v As Variant) As Long
    On Error GoTo Fallback
    If IsNumeric(v) Then
        SafeLong = CLng(v)
    Else
        SafeLong = 0
    End If
    Exit Function
Fallback:
    SafeLong = 0
End Function
