Attribute VB_Name = "REVO_FormBuilder"
'==============================================================================
' REVO_FormBuilder  -  builds frmReleaseDetails from code
'
' WHY THIS EXISTS
'   A UserForm's layout lives in a binary stream inside the VBA project, so it
'   cannot be delivered as a text file you paste in. Building it from code
'   means the form is version-controlled, reproducible, and one command to
'   rebuild after a taxonomy change.
'
' PREREQUISITE (one time, one checkbox)
'   File > Options > Trust Center > Trust Center Settings > Macro Settings
'     [x] Trust access to the VBA project object model
'   Without it Excel blocks programmatic access to the project and this module
'   stops with a clear message rather than a runtime error.
'
' The form's code-behind is read from frmReleaseDetails.code.vb, which lives
' beside the workbook. Edit that file and rebuild - never edit the form's code
' in the VBE, because the next rebuild overwrites it.
'
' Everything here is late-bound, so no VBIDE reference is needed.
'==============================================================================
Option Explicit

Private Const VB_EXT_CT_MSFORM As Long = 3
Private Const FORM_NAME        As String = "frmReleaseDetails"
Private Const CODE_FILE        As String = "frmReleaseDetails.code.vb"

'==============================================================================
Public Sub REVO_BuildReleaseForm(Optional ByVal codePath As String = "")
    Dim vbp As Object, vbc As Object, dsn As Object
    Dim code As String
    Dim fra As Object

    If Not VBAccessOK() Then Exit Sub

    codePath = ResolveCodePath(codePath)
    If Len(codePath) = 0 Then Exit Sub

    code = ReadTextFile(codePath)
    If Len(code) = 0 Then
        MsgBox "Could not read the form code from:" & vbCrLf & codePath, vbExclamation, "Form Builder"
        Exit Sub
    End If

    Set vbp = ThisWorkbook.VBProject

    '--- start clean --------------------------------------------------------
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents(FORM_NAME)
    On Error GoTo BuildFailed

    Set vbc = vbp.VBComponents.Add(VB_EXT_CT_MSFORM)
    vbc.name = FORM_NAME
    vbc.Properties("Caption") = "Release cart"
    vbc.Properties("Width") = 664
    vbc.Properties("Height") = 500

    Set dsn = vbc.Designer

    '========================= header block =================================
    AddLabel dsn, "lblHdr1", "SKU", 12, 8, 34, 14, True
    AddLabel dsn, "lblSKU", "", 48, 8, 268, 14, False
    AddLabel dsn, "lblHdr2", "Cart", 324, 8, 30, 14, True
    AddLabel dsn, "lblCart", "", 356, 8, 74, 14, False
    AddLabel dsn, "lblHdr3", "WO", 438, 8, 26, 14, True
    AddLabel dsn, "lblWO", "", 466, 8, 90, 14, False

    AddLabel dsn, "lblHdr4", "Cart qty", 12, 26, 50, 14, True
    AddLabel dsn, "lblCartQty", "", 64, 26, 50, 14, False
    AddLabel dsn, "lblHdr5", "Already released", 130, 26, 92, 14, True
    AddLabel dsn, "lblPriorRel", "", 224, 26, 50, 14, False
    AddLabel dsn, "lblHdr6", "Remaining", 290, 26, 60, 14, True
    AddLabel dsn, "lblRemaining", "", 352, 26, 50, 14, False

    AddLine dsn, "lnTop", 8, 46, 640

    '========================= release quantity =============================
    AddLabel dsn, "lblHdr7", "Release qty this pass", 12, 58, 122, 16, True
    AddTextBox dsn, "txtCartQtyRelease", 138, 56, 56, 20

    AddLabel dsn, "lblHdr8", "To inventory", 216, 58, 72, 16, True
    AddBigLabel dsn, "lblToInventory", "0", 292, 54, 66, 22

    AddLabel dsn, "lblAfterThisPass", "", 372, 58, 274, 16, False
    AddLabel dsn, "lblWarn", "", 12, 78, 634, 14, False

    '========================= disposition quantities =======================
    AddLabel dsn, "lblHdr9", "Rework", 12, 100, 46, 16, True
    AddTextBox dsn, "txtRework", 60, 98, 46, 20
    AddLabel dsn, "lblHdr10", "B Grade", 128, 100, 50, 16, True
    AddTextBox dsn, "txtBGrade", 180, 98, 46, 20
    AddLabel dsn, "lblHdr11", "Reject", 248, 100, 40, 16, True
    AddTextBox dsn, "txtReject", 290, 98, 46, 20
    AddLabel dsn, "lblHdr12", "Anything non-zero must be coded below.", 350, 100, 296, 16, False

    '========================= rework detail ================================
    Set fra = AddFrame(dsn, "fraRework", "Rework detail", 8, 122, 644, 92)
    AddLabel fra, "lblRwD", "Defect", 8, 16, 44, 14, True
    AddCombo fra, "cboReworkDefect", 54, 14, 128, 18
    AddLabel fra, "lblRwL", "Location", 190, 16, 48, 14, True
    AddCombo fra, "cboReworkLocation", 240, 14, 104, 18
    AddLabel fra, "lblRwO", "At op", 352, 16, 32, 14, True
    AddCombo fra, "cboReworkOp", 386, 14, 68, 18
    AddLabel fra, "lblRwR", "Root cause", 8, 40, 60, 14, True
    AddCombo fra, "cboReworkRoot", 70, 38, 128, 18
    AddLabel fra, "lblRwA", "Action", 206, 40, 40, 14, True
    AddCombo fra, "cboReworkAction", 248, 38, 128, 18
    AddLabel fra, "lblRwN", "Notes", 8, 64, 40, 14, True
    AddTextBox fra, "txtReworkNotes", 54, 62, 574, 18

    '========================= B grade detail ===============================
    Set fra = AddFrame(dsn, "fraBGrade", "B grade detail", 8, 220, 644, 70)
    AddLabel fra, "lblBgD", "Defect", 8, 16, 44, 14, True
    AddCombo fra, "cboBGradeDefect", 54, 14, 128, 18
    AddLabel fra, "lblBgL", "Location", 190, 16, 48, 14, True
    AddCombo fra, "cboBGradeLocation", 240, 14, 104, 18
    AddLabel fra, "lblBgR", "Root cause", 352, 16, 60, 14, True
    AddCombo fra, "cboBGradeRoot", 414, 14, 128, 18
    AddLabel fra, "lblBgN", "Notes", 8, 40, 40, 14, True
    AddTextBox fra, "txtBGradeNotes", 54, 38, 574, 18

    '========================= reject detail ================================
    Set fra = AddFrame(dsn, "fraReject", "Reject detail", 8, 296, 644, 92)
    AddLabel fra, "lblRjD", "Defect", 8, 16, 44, 14, True
    AddCombo fra, "cboRejectDefect", 54, 14, 128, 18
    AddLabel fra, "lblRjL", "Location", 190, 16, 48, 14, True
    AddCombo fra, "cboRejectLocation", 240, 14, 104, 18
    AddLabel fra, "lblRjO", "At op", 352, 16, 32, 14, True
    AddCombo fra, "cboRejectOp", 386, 14, 68, 18
    AddLabel fra, "lblRjR", "Root cause", 8, 40, 60, 14, True
    AddCombo fra, "cboRejectRoot", 70, 38, 128, 18
    AddLabel fra, "lblRjA", "Action", 206, 40, 40, 14, True
    AddCombo fra, "cboRejectAction", 248, 38, 128, 18
    AddLabel fra, "lblRjN", "Notes", 8, 64, 40, 14, True
    AddTextBox fra, "txtRejectNotes", 54, 62, 574, 18

    '========================= buttons ======================================
    AddButton dsn, "btnSubmit", "Submit release", 430, 400, 108, 28, True
    AddButton dsn, "btnCancel", "Cancel", 548, 400, 100, 28, False

    '========================= code =========================================
    With vbc.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString code
    End With

    REVO_Core.Audit "REVO_FormBuilder", "BUILD", FORM_NAME, "Rebuilt from " & codePath

    MsgBox "frmReleaseDetails rebuilt." & vbCrLf & vbCrLf & _
           "Controls: release qty, rework / B grade / reject quantities, and a " & _
           "coded defect, location, detected-at op, root cause and action for each." & vbCrLf & vbCrLf & _
           "Dropdowns read from the " & SH_QUAL_CODES & " sheet every time the form opens.", _
           vbInformation, "Form Builder"
    Exit Sub

BuildFailed:
    MsgBox "Building the form failed." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "If this says programmatic access is not trusted, tick" & vbCrLf & _
           "File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
           "Macro Settings > Trust access to the VBA project object model.", _
           vbCritical, "Form Builder"
End Sub

'==============================================================================
' CONTROL HELPERS
'
' parent is either the form's Designer or a Frame. Both expose Add, but the
' Designer exposes it via .Controls, so try the frame form first and fall back.
'==============================================================================
Private Function AddCtl(ByVal parent As Object, ByVal progID As String, _
                        ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                        ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    On Error Resume Next
    Set c = parent.Controls.Add(progID, nm, True)
    If c Is Nothing Then Set c = parent.Add(progID, nm, True)
    On Error GoTo 0
    If c Is Nothing Then Err.Raise vbObjectError + 1, , "Could not add control " & nm
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    Set AddCtl = c
End Function

Private Sub AddLabel(ByVal parent As Object, ByVal nm As String, ByVal cap As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                     ByVal h As Single, ByVal bold As Boolean)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.Label.1", nm, l, t, w, h)
    c.Caption = cap
    c.Font.name = "Segoe UI"
    c.Font.Size = 9
    c.Font.bold = bold
    If bold Then c.ForeColor = RGB(70, 70, 70)
End Sub

Private Sub AddBigLabel(ByVal parent As Object, ByVal nm As String, ByVal cap As String, _
                        ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.Label.1", nm, l, t, w, h)
    c.Caption = cap
    c.Font.name = "Segoe UI"
    c.Font.Size = 14
    c.Font.bold = True
    c.ForeColor = RGB(0, 97, 0)
End Sub

Private Sub AddTextBox(ByVal parent As Object, ByVal nm As String, _
                       ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.TextBox.1", nm, l, t, w, h)
    c.Font.name = "Segoe UI"
    c.Font.Size = 9
End Sub

Private Sub AddCombo(ByVal parent As Object, ByVal nm As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.ComboBox.1", nm, l, t, w, h)
    c.Font.name = "Segoe UI"
    c.Font.Size = 9
    c.MatchRequired = False        ' free text still allowed; validation is in the form
    c.Style = 0                    ' fmStyleDropDownCombo
End Sub

Private Function AddFrame(ByVal parent As Object, ByVal nm As String, ByVal cap As String, _
                          ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = AddCtl(parent, "Forms.Frame.1", nm, l, t, w, h)
    c.Caption = cap
    c.Font.name = "Segoe UI"
    c.Font.Size = 9
    c.Font.bold = True
    Set AddFrame = c
End Function

Private Sub AddButton(ByVal parent As Object, ByVal nm As String, ByVal cap As String, _
                      ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                      ByVal h As Single, ByVal isDefault As Boolean)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.CommandButton.1", nm, l, t, w, h)
    c.Caption = cap
    c.Font.name = "Segoe UI"
    c.Font.Size = 9
    c.Font.bold = isDefault
    On Error Resume Next
    c.Default = isDefault
    c.Cancel = Not isDefault
    On Error GoTo 0
End Sub

Private Sub AddLine(ByVal parent As Object, ByVal nm As String, _
                    ByVal l As Single, ByVal t As Single, ByVal w As Single)
    Dim c As Object
    On Error Resume Next
    Set c = AddCtl(parent, "Forms.Label.1", nm, l, t, w, 1)
    If Not c Is Nothing Then
        c.Caption = ""
        c.BackColor = RGB(200, 200, 200)
        c.BackStyle = 1
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' SUPPORT
'==============================================================================
Public Function VBAccessOK() As Boolean
    Dim n As Long
    On Error GoTo Blocked
    n = ThisWorkbook.VBProject.VBComponents.Count
    VBAccessOK = True
    Exit Function
Blocked:
    MsgBox "Excel is blocking programmatic access to the VBA project." & vbCrLf & vbCrLf & _
           "Tick this once and re-run:" & vbCrLf & vbCrLf & _
           "  File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
           "  Macro Settings > Trust access to the VBA project object model" & vbCrLf & vbCrLf & _
           "If your IT policy will not allow it, see revo/docs/MANUAL_FORM_BUILD.md " & _
           "for the control list to add by hand.", _
           vbExclamation, "Form Builder"
    VBAccessOK = False
End Function

Private Function ResolveCodePath(ByVal supplied As String) As String
    Dim p As String
    Dim fd As Object

    If Len(supplied) > 0 Then
        If Len(Dir$(supplied)) > 0 Then ResolveCodePath = supplied: Exit Function
    End If

    p = ThisWorkbook.Path
    If Len(p) > 0 Then
        If Len(Dir$(p & Application.PathSeparator & CODE_FILE)) > 0 Then
            ResolveCodePath = p & Application.PathSeparator & CODE_FILE
            Exit Function
        End If
        If Len(Dir$(p & Application.PathSeparator & "revo" & Application.PathSeparator & _
                    "vba" & Application.PathSeparator & CODE_FILE)) > 0 Then
            ResolveCodePath = p & Application.PathSeparator & "revo" & _
                              Application.PathSeparator & "vba" & _
                              Application.PathSeparator & CODE_FILE
            Exit Function
        End If
    End If

    Set fd = Application.FileDialog(3)      ' msoFileDialogFilePicker
    fd.Title = "Locate " & CODE_FILE
    fd.Filters.Clear
    fd.Filters.Add "Form code", "*.vb; *.txt; *.bas"
    If fd.Show = -1 Then ResolveCodePath = fd.SelectedItems(1)
End Function

Private Function ReadTextFile(ByVal path As String) As String
    Dim ff As Integer, s As String
    On Error GoTo Fail
    ff = FreeFile
    Open path For Input As #ff
    s = Input$(LOF(ff), ff)
    Close #ff
    ReadTextFile = s
    Exit Function
Fail:
    On Error Resume Next
    Close #ff
    ReadTextFile = ""
End Function

'==============================================================================
' BUTTON WIRING
'
' The REVO Floor button still points at the legacy
' ReleaseToShippingAndReceiving. Re-point it here rather than asking anyone to
' right-click and reassign.
'==============================================================================
Public Sub REVO_RewireButtons()
    Dim ws As Worksheet, shp As Shape
    Dim n As Long
    Dim rpt As String

    Set ws = REVO_Core.GetSheet(SH_FLOOR)
    If ws Is Nothing Then Exit Sub

    For Each shp In ws.Shapes
        On Error Resume Next
        If InStr(1, shp.OnAction, "ReleaseToShippingAndReceiving", vbTextCompare) > 0 Then
            shp.OnAction = "REVO_ReleaseCarts"
            n = n + 1
            rpt = rpt & "  " & ws.name & " / " & shp.name & vbCrLf
        End If
        On Error GoTo 0
    Next shp

    ' The Release Log "undo" button points at a macro that no longer exists.
    Set ws = REVO_Core.GetSheet(SH_RELEASE_LOG)
    If Not ws Is Nothing Then
        For Each shp In ws.Shapes
            On Error Resume Next
            If InStr(1, shp.OnAction, "UndoReleaseForCart", vbTextCompare) > 0 Then
                shp.OnAction = "REVO_ResetCart"
                n = n + 1
                rpt = rpt & "  " & ws.name & " / " & shp.name & " (was a dead macro)" & vbCrLf
            End If
            On Error GoTo 0
        Next shp
    End If

    REVO_Core.Audit "REVO_FormBuilder", "REWIRE", "Buttons", n & " button(s) re-pointed"

    If n = 0 Then
        MsgBox "No buttons needed re-pointing.", vbInformation, "Rewire"
    Else
        MsgBox n & " button(s) re-pointed:" & vbCrLf & vbCrLf & rpt, vbInformation, "Rewire"
    End If
End Sub
