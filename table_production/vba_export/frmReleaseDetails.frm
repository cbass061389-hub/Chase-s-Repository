Option Explicit


Public UserCancelled As Boolean
Public sku As String
Public Cart As String
Public TotalQty As Long

Public Property Get ReleasedToInventory() As Long
    ReleasedToInventory = TotalQty - Rework - BGrade - Reject
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

Private Function SafeLong(val As Variant) As Long
    If IsNumeric(val) Then
        SafeLong = CLng(val)
    Else
        SafeLong = 0
    End If
End Function

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Initialize()
    txtBGrade.Value = 0
    txtRework.Value = 0
    txtReject.Value = 0
    txtRejectReason.Value = ""


    lblRejectReason.Visible = False
    txtRejectReason.Visible = False

    Call UpdateReleaseAmount
End Sub

Private Sub UpdateReleaseAmount()
    Dim rel As Long
    rel = TotalQty - Rework - BGrade - Reject
    If rel < 0 Then rel = 0
    If rel > TotalQty Then rel = TotalQty
    txtRelease.Caption = rel
End Sub

Private Sub txtBGrade_Change()
    Call UpdateReleaseAmount
End Sub

Private Sub txtRework_Change()
    UpdateReleaseAmount
End Sub


Private Sub txtReject_Change()
    Call UpdateReleaseAmount

    Dim rejectQty As Long
    rejectQty = SafeLong(txtReject.Value)

    If rejectQty > 0 Then
        lblRejectReason.Visible = True
        txtRejectReason.Visible = True
    Else
        lblRejectReason.Visible = False
        txtRejectReason.Visible = False
        txtRejectReason.Value = "" ' Clear any previous entry
    End If
End Sub



Private Sub btnCancel_Click()
    UserCancelled = True
    Me.Hide
End Sub

Private Sub btnSubmit_Click()
    Dim rejectQty As Long
    rejectQty = SafeLong(txtReject.Value)

    If rejectQty > 0 And Trim(txtRejectReason.Value) = "" Then
        MsgBox "Please enter a reject reason before submitting.", vbExclamation
        Exit Sub
    End If
    If (Rework + BGrade + Reject) > TotalQty Then
    MsgBox "Rework + B Grade + Reject cannot exceed Total Qty (" & TotalQty & ").", vbExclamation
    Exit Sub
End If


    UserCancelled = False
    Me.Hide
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.UserCancelled = True
        Debug.Print "? X button was clicked. UserCancelled set to TRUE."
    End If
End Sub
