Option Explicit

' ===== Layout assumptions for REVO Floor =====
Private Const SRC_SHEET As String = "REVO Floor"
Private Const COL_QTY As Long = 5         ' E: FORECAST
Private Const HEADER_ROW As Long = 2      ' row with numeric op headers (10,20,...)
Private Const DATA_FIRST_ROW As Long = 3  ' first data row

' ===== State =====
Private totalForecast As Double
Private ops() As Long         ' available operation values from row 2
Private opCols() As Long
Private opCount As Long

' Editable plan (1-based arrays)
Private weekCaps() As Double
Private weekMaxOp() As Long
Private w As Long
Private mInitDone As Boolean  ' ensures Activate runs once

Private Sub UserForm_Initialize()
    ' Keep initialize feather-light to avoid type mismatch during load
    Me.Caption = "REVO Weekly Planner"
End Sub

Private Sub UserForm_Activate()
    Static ran As Boolean
    If ran Then Exit Sub
    ran = True

    ' --- UI props (tolerant) ---
    On Error Resume Next
    Me.lblTitle.Font.Bold = True

    Me.lstWeeks.Clear
    Me.lstWeeks.ColumnCount = 3
    Me.lstWeeks.ColumnHeads = False

    ' ColumnWidths can be finicky on some builds/regions; skip if it errors.
    Err.Clear
    Me.lstWeeks.ColumnWidths = "50;80;80"
    If Err.Number <> 0 Then
        Err.Clear
        ' If you *really* want widths, try commas; otherwise leave defaults
        Me.lstWeeks.ColumnWidths = "50,80,80"
        Err.Clear
    End If
    On Error GoTo 0

    ' --- Load ops & totals ---
    LoadOperations                      ' fills ops(), opCount
    totalForecast = ComputeTotalForecast
    Me.lblTotalVal.Caption = Format$(totalForecast, "#,##0")

    ' --- Fill Max Op dropdown ---
    Dim i As Long
    Me.cboMaxOp.Clear
    If opCount > 0 Then
        For i = 1 To opCount
            Me.cboMaxOp.AddItem CStr(ops(i))
        Next i
    End If

    Me.btnApply.Enabled = False
    Me.txtWeeklyCap.Value = ""
End Sub



' ---------- build initial weeks ----------
Private Sub btnBuild_Click()
    Dim cap As Double
    cap = val(txtWeeklyCap.Value)
    If cap <= 0 Then
        MsgBox "Enter a positive Default Weekly Capacity.", vbExclamation
        Exit Sub
    End If
    If totalForecast <= 0 Then
        MsgBox "No FORECAST detected.", vbInformation
        Exit Sub
    End If

    ' compute weeks needed (ceil)
    w = CeilDiv(totalForecast, cap)
    If w < 1 Then w = 1

    ' init arrays
    ReDim weekCaps(1 To w)
    ReDim weekMaxOp(1 To w)

    Dim i As Long, maxOp As Long
    maxOp = IIf(opCount > 0, ops(opCount), 0)
    For i = 1 To w
        weekCaps(i) = cap
        weekMaxOp(i) = maxOp
    Next i

    ' show in list
    RefreshWeekList

    ' select first row for edit
    If lstWeeks.ListCount > 0 Then
        lstWeeks.ListIndex = 0
        txtCap.Value = CStr(weekCaps(1))
        cboMaxOp.Value = CStr(weekMaxOp(1))
        btnApply.Enabled = True
    End If
End Sub

' ---------- list selection -> load editors ----------
Private Sub lstWeeks_Change()
    Dim idx As Long
    idx = lstWeeks.ListIndex
    If idx < 0 Then Exit Sub
    ' listindex is 0-based, our arrays are 1-based
    txtCap.Value = CStr(weekCaps(idx + 1))
    cboMaxOp.Value = CStr(weekMaxOp(idx + 1))
    btnApply.Enabled = True
End Sub

' ---------- apply edit to selected week ----------
Private Sub btnApply_Click()
    Dim idx As Long
    idx = lstWeeks.ListIndex
    If idx < 0 Then
        MsgBox "Select a week to edit.", vbExclamation
        Exit Sub
    End If

    Dim cap As Double, mOp As Long
    cap = val(txtCap.Value)
    If cap <= 0 Then
        MsgBox "Capacity must be > 0.", vbExclamation
        Exit Sub
    End If
    If cboMaxOp.ListCount = 0 Then
        MsgBox "No operations detected to choose from.", vbExclamation
        Exit Sub
    End If
    If Len(Trim$(cboMaxOp.Value)) = 0 Or Not IsNumeric(cboMaxOp.Value) Then
        MsgBox "Pick a numeric Max Op from the list.", vbExclamation
        Exit Sub
    End If
    mOp = CLng(cboMaxOp.Value)

    ' validate MaxOp exists in ops set
    If Not OpExists(mOp) Then
        MsgBox "Max Op value not in detected operation headers.", vbExclamation
        Exit Sub
    End If

    weekCaps(idx + 1) = cap
    weekMaxOp(idx + 1) = mOp

    ' update list row
    lstWeeks.list(idx, 1) = Format(cap, "#,##0")
    lstWeeks.list(idx, 2) = CStr(mOp)
End Sub

' ---------- finalize & arrange ----------
Private Sub btnSubmit_Click()
    If w = 0 Or lstWeeks.ListCount = 0 Then
        MsgBox "Build your weeks first.", vbExclamation
        Exit Sub
    End If

    ' call your arrange macro and shading
    On Error GoTo fail
    ArrangeREVOFloorByWeeks weekCaps, weekMaxOp
    ShadeWeeksMediumLight
    Unload Me
    Exit Sub
fail:
    MsgBox "Arrange failed: " & Err.Description, vbExclamation
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

' ---------- helpers ----------
Private Sub LoadOperations()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SRC_SHEET)

    Dim lastCol As Long, c As Long, v As Variant
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    ' collect numeric headers from HEADER_ROW
    Dim tmpOps() As Long, n As Long
    ReDim tmpOps(1 To lastCol)
    For c = 1 To lastCol
        v = ws.Cells(HEADER_ROW, c).Value
        If IsNumeric(v) Then
            n = n + 1
            tmpOps(n) = CLng(v)
        End If
    Next c

    If n > 0 Then
        ReDim ops(1 To n)
        Dim i As Long
        ' copy (already left-to-right)
        For i = 1 To n
            ops(i) = tmpOps(i)
        Next i
        opCount = n
    Else
        opCount = 0
    End If
End Sub

Private Function ComputeTotalForecast() As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SRC_SHEET)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_QTY).End(xlUp).row

    Dim r As Long, s As Double
    For r = DATA_FIRST_ROW To lastRow
        s = s + SafeDbl(ws.Cells(r, COL_QTY).Value)
    Next r
    ComputeTotalForecast = s
End Function

Private Sub RefreshWeekList()
    Dim i As Long
    lstWeeks.Clear
    For i = 1 To w
        lstWeeks.AddItem
        lstWeeks.list(i - 1, 0) = i                 ' Week #
        lstWeeks.list(i - 1, 1) = Format(weekCaps(i), "#,##0")
        lstWeeks.list(i - 1, 2) = CStr(weekMaxOp(i))
    Next i
End Sub

Private Function OpExists(ByVal v As Long) As Boolean
    Dim i As Long
    For i = 1 To opCount
        If ops(i) = v Then OpExists = True: Exit Function
    Next i
End Function

Private Function SafeDbl(v As Variant) As Double
    ' Converts to Double safely; returns 0 if not a clean number
    Select Case True
        Case IsError(v), IsEmpty(v), VarType(v) = vbNull
            SafeDbl = 0#
        Case IsNumeric(Trim$(CStr(v)))
            SafeDbl = CDbl(Trim$(CStr(v)))
        Case Else
            SafeDbl = 0#
    End Select
End Function


Private Function CeilDiv(ByVal total As Double, ByVal byValCap As Double) As Long
    ' ceiling(total / cap)
    CeilDiv = CLng(Application.WorksheetFunction.RoundUp(total / byValCap, 0))
    If CeilDiv < 1 Then CeilDiv = 1
End Function
