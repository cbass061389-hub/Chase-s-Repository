Option Explicit

'==== Public flag ====
Public UserCancelled As Boolean

'==== Private state ====
Private mSKU As String
Private mCart As String
Private mRowQty As Long           ' original row qty (TotalQty)
' You can expose these via labels; no need for public fields

'==============================
' Lifecycle
'==============================
Private Sub UserForm_Initialize()
    txtQty.Value = "0"
    txtBGrade.Value = "0"
    txtReject.Value = "0"
    txtRejectReason.Value = vbNullString
    lblReleased.Caption = "0"
    UserCancelled = True
End Sub

' Prime the form with the row context; default Qty-to-process to full row qty
Public Sub Prime(ByVal pSKU As String, ByVal pCart As String, ByVal pQty As Long)
    mSKU = pSKU
    mCart = pCart
    mRowQty = pQty

    lblSKU.Caption = pSKU
    lblCart.Caption = pCart & " REWORK"
    lblQty.Caption = CStr(pQty)

    txtQty.Value = CStr(pQty)     ' default to full qty for this run
    txtBGrade.Value = "0"
    txtReject.Value = "0"
    txtRejectReason.Value = vbNullString

    RecalcReleased
End Sub

'==============================
' Exposed values (properties)
'==============================
Public Property Get RequestedQty() As Long
    RequestedQty = SafeLng(txtQty.Value)
End Property

Public Property Get BGrade() As Long
    BGrade = SafeLng(txtBGrade.Value)
End Property

Public Property Get Reject() As Long
    Reject = SafeLng(txtReject.Value)
End Property

Public Property Get ReleasedToInventory() As Long
    Dim rq As Long: rq = RequestedQty
    Dim rel As Long
    rel = rq - BGrade - Reject
    If rel < 0 Then rel = 0
    ReleasedToInventory = rel
End Property

'==============================
' Buttons
'==============================
Private Sub cmdCancel_Click()
    UserCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim rq As Long, bg As Long, rj As Long

    rq = RequestedQty
    bg = BGrade
    rj = Reject

    ' --- validations for partial release flow ---
    If rq <= 0 Then
        MsgBox "Enter a Qty to process greater than 0.", vbExclamation
        Exit Sub
    End If
    If rq > mRowQty Then
        MsgBox "Qty to process cannot exceed the row quantity (" & mRowQty & ").", vbExclamation
        Exit Sub
    End If
    If (bg < 0) Or (rj < 0) Then
        MsgBox "Negative values are not allowed.", vbExclamation
        Exit Sub
    End If
    If (bg + rj) > rq Then
        MsgBox "B-Grade + Reject cannot exceed the Qty to process.", vbExclamation
        Exit Sub
    End If
    If rj > 0 And Trim$(txtRejectReason.Value) = vbNullString Then
        MsgBox "Please enter a Reject Description.", vbExclamation
        Exit Sub
    End If

    ' All good — caller will read properties
    UserCancelled = False
    Me.Hide
End Sub

'==============================
' Live recalc / UX
'==============================
Private Sub txtQty_Change():      RecalcReleased: End Sub
Private Sub txtBGrade_Change():   RecalcReleased: End Sub
Private Sub txtReject_Change()
    RecalcReleased
    ' Show/hide reason box
    If SafeLng(txtReject.Value) > 0 Then
        txtRejectReason.Visible = True
        ' if you have a label for reason, show it too:
        ' lblRejectReason.Visible = True
    Else
        txtRejectReason.Visible = False
        ' lblRejectReason.Visible = False
        txtRejectReason.Value = vbNullString
    End If
End Sub

Private Sub RecalcReleased()
    Dim rq As Long, rel As Long
    rq = RequestedQty
    rel = rq - SafeLng(txtBGrade.Value) - SafeLng(txtReject.Value)
    If rel < 0 Then rel = 0
    lblReleased.Caption = CStr(rel)
End Sub

'==============================
' Helpers
'==============================
Private Function SafeLng(v As Variant) As Long
    If IsError(v) Or IsEmpty(v) Then
        SafeLng = 0
    ElseIf IsNumeric(v) Then
        SafeLng = CLng(val(CStr(v)))
    Else
        SafeLng = 0
    End If
End Function
