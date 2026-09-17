'==============================================================================
' frmReleaseDetails  -  code-behind
'
' This file is the source of truth for the form's code. REVO_FormBuilder injects
' it when it builds the form, so edit it here and rebuild rather than editing in
' the VBE and losing the change next build.
'
' QUALITY DETAIL IS A LIST, NOT A FIELD
'   Reject 10 shafts and it is rarely one defect - 3 collet crush, 5 chip at
'   butt, 2 breakthrough. One defect per disposition would force the operator to
'   pick whichever was most common and throw the rest away, which is exactly the
'   data loss this system exists to stop.
'
'   So each disposition carries as many detail lines as it needs. The operator
'   enters a quantity and a cause, presses Add, and repeats. Submit is blocked
'   until each disposition's lines add up to its total, so the counts can never
'   drift from the dispositions.
'
' NOTE on the mLoading guard: UserForm_Initialize deliberately does not touch
' it. The first reference to Me.Controls from Prime is what triggers Initialize,
' so an Initialize that cleared the flag would switch the guard off half way
' through Prime and let every _Change handler fire during setup.
'==============================================================================
Option Explicit

Private Const MAX_LINES As Long = 60

Public UserCancelled As Boolean

Private mLoading   As Boolean
Private mSku       As String
Private mCart      As String
Private mWO        As String
Private mRemaining As Long      ' units still on the cart - the cap for this pass
Private mCartQty   As Long
Private mPriorRel  As Long

' Quality detail lines, parallel arrays. Small and fixed - a cart release never
' carries sixty distinct defects, and MAX_LINES is a guard rail not a target.
Private mN        As Long
Private mDisp(1 To MAX_LINES)   As String
Private mQty(1 To MAX_LINES)    As Long
Private mDefect(1 To MAX_LINES) As String
Private mLoc(1 To MAX_LINES)    As String
Private mOp(1 To MAX_LINES)     As String
Private mRoot(1 To MAX_LINES)   As String
Private mAct(1 To MAX_LINES)    As String
Private mNote(1 To MAX_LINES)   As String

'==================== PRIME ===================================================
Public Sub Prime(ByVal sSku As String, ByVal sCart As String, ByVal sWO As String, _
                 ByVal lRemaining As Long, ByVal lCartQty As Long, ByVal lPriorRel As Long)
    mLoading = True

    mSku = sSku
    mCart = sCart
    mWO = sWO
    mRemaining = lRemaining
    mCartQty = lCartQty
    mPriorRel = lPriorRel
    mN = 0

    UserCancelled = True          ' nothing commits unless Submit is pressed

    Me.Caption = "Release cart " & sCart & "  -  " & sSku

    SetCap "lblSKU", sSku
    SetCap "lblCart", sCart
    SetCap "lblWO", sWO
    SetCap "lblCartQty", CStr(lCartQty)
    SetCap "lblPriorRel", CStr(lPriorRel)
    SetCap "lblRemaining", CStr(lRemaining)

    FillCombo Me.cboDisposition, Array("REWORK", "B GRADE", "REJECT")
    FillCombo Me.cboDefect, REVO_Quality.DefectTypeList()
    FillCombo Me.cboLocation, REVO_Quality.LocationList()
    FillCombo Me.cboOp, REVO_Quality.OpList()
    FillCombo Me.cboRoot, REVO_Quality.RootCauseList()
    FillCombo Me.cboAction, REVO_Quality.ActionList()

    Me.txtCartQtyRelease.Value = lRemaining     ' default: release what is left
    Me.txtRework.Value = 0
    Me.txtBGrade.Value = 0
    Me.txtReject.Value = 0

    ClearEntryRow
    SetupList
    RefreshList

    mLoading = False

    UpdateReleaseAmount
End Sub

'==================== TOTALS ==================================================
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

'==================== DETAIL LINES (read by REVO_Release) =====================
Public Property Get LineCount() As Long
    LineCount = mN
End Property

Public Property Get LineDisposition(ByVal i As Long) As String
    If i >= 1 And i <= mN Then LineDisposition = mDisp(i)
End Property

Public Property Get LineQty(ByVal i As Long) As Long
    If i >= 1 And i <= mN Then LineQty = mQty(i)
End Property

Public Property Get LineDefect(ByVal i As Long) As String
    If i >= 1 And i <= mN Then LineDefect = mDefect(i)
End Property

Public Property Get LineLocation(ByVal i As Long) As String
    If i >= 1 And i <= mN Then LineLocation = mLoc(i)
End Property

Public Property Get LineOp(ByVal i As Long) As String
    If i >= 1 And i <= mN Then LineOp = mOp(i)
End Property

Public Property Get LineRootCause(ByVal i As Long) As String
    If i >= 1 And i <= mN Then LineRootCause = mRoot(i)
End Property

Public Property Get LineAction(ByVal i As Long) As String
    If i >= 1 And i <= mN Then LineAction = mAct(i)
End Property

Public Property Get LineNotes(ByVal i As Long) As String
    If i >= 1 And i <= mN Then LineNotes = mNote(i)
End Property

' Units already accounted for against one disposition.
Public Function AllocatedFor(ByVal disp As String) As Long
    Dim i As Long, t As Long
    For i = 1 To mN
        If StrComp(mDisp(i), disp, vbTextCompare) = 0 Then t = t + mQty(i)
    Next i
    AllocatedFor = t
End Function

' A one-line summary of the reject causes, for the legacy Reject Log column.
Public Function RejectSummary() As String
    Dim i As Long, s As String
    For i = 1 To mN
        If StrComp(mDisp(i), "REJECT", vbTextCompare) = 0 Then
            If Len(s) > 0 Then s = s & "; "
            s = s & mQty(i) & "x " & mDefect(i)
            If Len(mLoc(i)) > 0 And mLoc(i) <> "Not Specified" Then s = s & " (" & mLoc(i) & ")"
            If Len(mRoot(i)) > 0 Then s = s & " - " & mRoot(i)
        End If
    Next i
    RejectSummary = s
End Function

'==================== UI ======================================================
Private Sub UserForm_Initialize()
    ' Deliberately minimal - see the header note about mLoading.
    UserCancelled = True
End Sub

Private Sub SetupList()
    On Error Resume Next
    With Me.lstLines
        .Clear
        .ColumnCount = 7
        .ColumnWidths = "58 pt;30 pt;90 pt;62 pt;38 pt;96 pt;120 pt"
        .ColumnHeads = False
    End With
    On Error GoTo 0
End Sub

Private Sub RefreshList()
    Dim i As Long
    On Error Resume Next
    Me.lstLines.Clear
    For i = 1 To mN
        Me.lstLines.AddItem mDisp(i)
        Me.lstLines.List(i - 1, 1) = CStr(mQty(i))
        Me.lstLines.List(i - 1, 2) = mDefect(i)
        Me.lstLines.List(i - 1, 3) = mLoc(i)
        Me.lstLines.List(i - 1, 4) = mOp(i)
        Me.lstLines.List(i - 1, 5) = mRoot(i)
        Me.lstLines.List(i - 1, 6) = mNote(i)
    Next i
    On Error GoTo 0
    UpdateAllocation
End Sub

Private Sub ClearEntryRow()
    On Error Resume Next
    Me.cboDisposition.Value = vbNullString
    Me.txtLineQty.Value = vbNullString
    Me.cboDefect.Value = vbNullString
    Me.cboLocation.Value = vbNullString
    Me.cboOp.Value = vbNullString
    Me.cboRoot.Value = vbNullString
    Me.cboAction.Value = vbNullString
    Me.txtLineNotes.Value = vbNullString
    On Error GoTo 0
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

    If req > mRemaining Then
        Warn "Release qty exceeds the " & mRemaining & " left on this cart."
    ElseIf disp > req Then
        Warn "Rework + B Grade + Reject (" & disp & ") exceeds the release qty (" & req & ")."
    ElseIf req <= 0 Then
        Warn "Enter a release quantity."
    Else
        Warn ""
    End If

    UpdateAllocation
End Sub

' The running score: how much of each disposition still needs a cause.
Private Sub UpdateAllocation()
    Dim s As String
    Dim shortfall As Boolean

    If mLoading Then Exit Sub

    s = DispState("REWORK", Rework, shortfall) & "    " & _
        DispState("B GRADE", BGrade, shortfall) & "    " & _
        DispState("REJECT", Reject, shortfall)

    On Error Resume Next
    Me.lblAllocation.Caption = s
    If shortfall Then
        Me.lblAllocation.ForeColor = RGB(150, 75, 0)
    Else
        Me.lblAllocation.ForeColor = RGB(0, 97, 0)
    End If
    On Error GoTo 0
End Sub

Private Function DispState(ByVal disp As String, ByVal total As Long, _
                           ByRef shortfall As Boolean) As String
    Dim got As Long
    got = AllocatedFor(disp)
    If total = 0 And got = 0 Then
        DispState = disp & ": -"
    Else
        DispState = disp & ": " & got & " of " & total
        If got <> total Then shortfall = True
    End If
End Function

Private Sub Warn(ByVal msg As String)
    On Error Resume Next
    Me.lblWarn.Caption = msg
    Me.lblWarn.ForeColor = RGB(192, 0, 0)
    On Error GoTo 0
End Sub

'==================== EVENTS ==================================================
Private Sub txtCartQtyRelease_Change()
    UpdateReleaseAmount
End Sub

Private Sub txtRework_Change()
    PrefillEntry "REWORK", Rework
    UpdateReleaseAmount
End Sub

Private Sub txtBGrade_Change()
    PrefillEntry "B GRADE", BGrade
    UpdateReleaseAmount
End Sub

Private Sub txtReject_Change()
    PrefillEntry "REJECT", Reject
    UpdateReleaseAmount
End Sub

' Typing a disposition quantity pre-loads the entry row with the unallocated
' balance, so the common case - one cause for the whole lot - is Add and done.
Private Sub PrefillEntry(ByVal disp As String, ByVal total As Long)
    Dim outstanding As Long

    If mLoading Then Exit Sub
    If total <= 0 Then Exit Sub

    outstanding = total - AllocatedFor(disp)
    If outstanding <= 0 Then Exit Sub

    On Error Resume Next
    If Len(Trim$(CStr(Me.cboDisposition.Value))) = 0 Or _
       StrComp(CStr(Me.cboDisposition.Value), disp, vbTextCompare) = 0 Then
        Me.cboDisposition.Value = disp
        Me.txtLineQty.Value = outstanding
    End If
    On Error GoTo 0
End Sub

'==================== ADD / REMOVE ============================================
Private Sub btnAddLine_Click()
    Dim disp As String, q As Long
    Dim outstanding As Long

    disp = UCase$(Trim$(CStr(Me.cboDisposition.Value)))
    q = SafeLong(Me.txtLineQty.Value)

    If disp <> "REWORK" And disp <> "B GRADE" And disp <> "REJECT" Then
        Complain "Pick a disposition for this line - Rework, B Grade or Reject.", Me.cboDisposition
        Exit Sub
    End If

    If q <= 0 Then
        Complain "Enter how many shafts this cause accounts for.", Me.txtLineQty
        Exit Sub
    End If

    outstanding = DispTotal(disp) - AllocatedFor(disp)
    If DispTotal(disp) = 0 Then
        Complain "The " & disp & " quantity is zero. Set it above before adding a line for it.", _
                 Me.cboDisposition
        Exit Sub
    End If
    If q > outstanding Then
        Complain "That is " & q & " but only " & outstanding & " of the " & DispTotal(disp) & _
                 " " & disp & " units are still unaccounted for.", Me.txtLineQty
        Exit Sub
    End If

    If Len(Trim$(CStr(Me.cboDefect.Value))) = 0 Then
        Complain "Pick a defect type for this line.", Me.cboDefect
        Exit Sub
    End If

    If Len(Trim$(CStr(Me.cboRoot.Value))) = 0 Then
        Complain "Pick a root cause for this line." & vbCrLf & vbCrLf & _
                 "If it genuinely is not known yet, choose 'Not Determined' - that is a " & _
                 "reportable answer and a blank is not.", Me.cboRoot
        Exit Sub
    End If

    If mN >= MAX_LINES Then
        Complain "That is " & MAX_LINES & " detail lines on one cart, which is almost " & _
                 "certainly a mistake. Submit what you have and raise the rest separately.", _
                 Me.cboDisposition
        Exit Sub
    End If

    mN = mN + 1
    mDisp(mN) = disp
    mQty(mN) = q
    mDefect(mN) = Trim$(CStr(Me.cboDefect.Value))
    mLoc(mN) = Trim$(CStr(Me.cboLocation.Value))
    mOp(mN) = Trim$(CStr(Me.cboOp.Value))
    mRoot(mN) = Trim$(CStr(Me.cboRoot.Value))
    mAct(mN) = Trim$(CStr(Me.cboAction.Value))
    mNote(mN) = Trim$(CStr(Me.txtLineNotes.Value))

    If Len(mLoc(mN)) = 0 Then mLoc(mN) = "Not Specified"
    If Len(mOp(mN)) = 0 Then mOp(mN) = "Unknown"

    ClearEntryRow
    RefreshList
    Warn ""

    ' Straight into the next cause for the same disposition if any remain.
    PrefillEntry disp, DispTotal(disp)
    On Error Resume Next
    Me.cboDefect.SetFocus
    On Error GoTo 0
End Sub

Private Sub btnRemoveLine_Click()
    Dim idx As Long, i As Long

    idx = -1
    On Error Resume Next
    idx = Me.lstLines.ListIndex
    On Error GoTo 0

    If idx < 0 Then
        Complain "Select a line in the list first.", Me.lstLines
        Exit Sub
    End If

    For i = idx + 1 To mN - 1
        mDisp(i) = mDisp(i + 1)
        mQty(i) = mQty(i + 1)
        mDefect(i) = mDefect(i + 1)
        mLoc(i) = mLoc(i + 1)
        mOp(i) = mOp(i + 1)
        mRoot(i) = mRoot(i + 1)
        mAct(i) = mAct(i + 1)
        mNote(i) = mNote(i + 1)
    Next i
    mN = mN - 1

    RefreshList
End Sub

Private Function DispTotal(ByVal disp As String) As Long
    Select Case UCase$(disp)
        Case "REWORK":  DispTotal = Rework
        Case "B GRADE": DispTotal = BGrade
        Case "REJECT":  DispTotal = Reject
    End Select
End Function

'==================== SUBMIT ==================================================
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

    ' Every dispositioned shaft must carry a cause, and the causes must add up.
    If Not DispComplete("REWORK", Rework, Me.txtRework) Then Exit Sub
    If Not DispComplete("B GRADE", BGrade, Me.txtBGrade) Then Exit Sub
    If Not DispComplete("REJECT", Reject, Me.txtReject) Then Exit Sub

    UserCancelled = False
    Me.Hide
End Sub

Private Function DispComplete(ByVal disp As String, ByVal total As Long, _
                              ByVal ctl As Object) As Boolean
    Dim got As Long
    got = AllocatedFor(disp)

    If total = 0 And got > 0 Then
        Complain "There are detail lines for " & disp & " totalling " & got & _
                 ", but the " & disp & " quantity is zero." & vbCrLf & vbCrLf & _
                 "Either set the quantity or remove those lines.", ctl
        Exit Function
    End If

    If got <> total Then
        Complain disp & ": " & got & " of " & total & " shafts have a cause recorded." & vbCrLf & vbCrLf & _
                 "Add a detail line for the remaining " & (total - got) & "." & vbCrLf & _
                 "If they are all the same cause, one line covers them.", ctl
        Exit Function
    End If

    DispComplete = True
End Function

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

Private Sub FillCombo(ByVal cbo As Object, ByVal items As Variant)
    Dim i As Long
    On Error Resume Next
    cbo.Clear
    For i = LBound(items) To UBound(items)
        If Len(CStr(items(i))) > 0 Then cbo.AddItem CStr(items(i))
    Next i
    cbo.Value = vbNullString
    On Error GoTo 0
End Sub

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
    Resume Done
Done:
End Function
