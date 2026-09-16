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

' Set as the build walks, so a failure names the step rather than a bare number.
Private step_ As String

' Read by REVO_Install so it can report a build failure instead of swallowing it.
Public LastBuildError As String

'==============================================================================
Public Sub REVO_BuildReleaseForm(Optional ByVal codePath As String = "")
    Dim vbp As Object, vbc As Object, dsn As Object
    Dim code As String
    Dim fra As Object

    If Not VBAccessOK() Then Exit Sub

    codePath = ResolveCodePath(codePath)
    If Len(codePath) = 0 Then Exit Sub

    step_ = "reading " & CODE_FILE
    code = ReadTextFile(codePath)
    If Len(code) = 0 Then
        MsgBox "Could not read the form code from:" & vbCrLf & codePath, vbExclamation, "Form Builder"
        Exit Sub
    End If

    step_ = "opening the VBA project"
    Set vbp = ThisWorkbook.VBProject

    '--- start clean --------------------------------------------------------
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents(FORM_NAME)
    On Error GoTo BuildFailed

    step_ = "creating the UserForm component"
    Set vbc = vbp.VBComponents.Add(VB_EXT_CT_MSFORM)
    vbc.name = FORM_NAME

    ' Caption / Width / Height go through the VBComponent Properties collection,
    ' which is not exposed identically across Excel builds. Cosmetic only - a
    ' failure here must not cost us the form.
    On Error Resume Next
    vbc.Properties("Caption") = "Release cart"
    vbc.Properties("Width") = 664
    vbc.Properties("Height") = 500
    Err.Clear
    On Error GoTo BuildFailed

    step_ = "opening the form designer"
    Set dsn = vbc.Designer

    '========================= header block =================================
    step_ = "adding header controls"
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


    '========================= release quantity =============================
    step_ = "adding the release quantity controls"
    AddLabel dsn, "lblHdr7", "Release qty this pass", 12, 58, 122, 16, True
    AddTextBox dsn, "txtCartQtyRelease", 138, 56, 56, 20

    AddLabel dsn, "lblHdr8", "To inventory", 216, 58, 72, 16, True
    AddBigLabel dsn, "lblToInventory", "0", 292, 54, 66, 22

    AddLabel dsn, "lblAfterThisPass", "", 372, 58, 274, 16, False
    AddLabel dsn, "lblWarn", "", 12, 78, 634, 14, False

    '========================= disposition quantities =======================
    step_ = "adding the disposition quantity controls"
    AddLabel dsn, "lblHdr9", "Rework", 12, 100, 46, 16, True
    AddTextBox dsn, "txtRework", 60, 98, 46, 20
    AddLabel dsn, "lblHdr10", "B Grade", 128, 100, 50, 16, True
    AddTextBox dsn, "txtBGrade", 180, 98, 46, 20
    AddLabel dsn, "lblHdr11", "Reject", 248, 100, 40, 16, True
    AddTextBox dsn, "txtReject", 290, 98, 46, 20
    AddLabel dsn, "lblHdr12", "Anything non-zero must be coded below.", 350, 100, 296, 16, False

    '========================= rework detail ================================
    step_ = "adding the rework panel"
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
    step_ = "adding the B grade panel"
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
    step_ = "adding the reject panel"
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
    step_ = "adding the buttons"
    AddButton dsn, "btnSubmit", "Submit release", 430, 400, 108, 28, True
    AddButton dsn, "btnCancel", "Cancel", 548, 400, 100, 28, False

    '========================= code =========================================
    step_ = "injecting the form code"
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
    LastBuildError = ""
    Exit Sub

BuildFailed:
    ' Capture Err FIRST - anything called below can clear it.
    Dim eNum As Long, eDesc As String
    eNum = Err.Number
    eDesc = Err.Description

    LastBuildError = "Error " & eNum & ": " & eDesc & "  (step: " & step_ & ")"
    REVO_Core.Audit "REVO_FormBuilder", "ERROR", FORM_NAME, LastBuildError

    MsgBox "Building the form failed." & vbCrLf & vbCrLf & _
           "Error " & eNum & ": " & eDesc & vbCrLf & vbCrLf & _
           "Failed at: " & step_ & vbCrLf & vbCrLf & _
           "If this says programmatic access is not trusted, tick" & vbCrLf & _
           "File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
           "Macro Settings > Trust access to the VBA project object model.", _
           vbCritical, "Form Builder"

    ' Resume, not a bare End Sub. Falling off an active handler leaves the error
    ' pending and it surfaces in the CALLER's handler with Err already cleared -
    ' which is how the installer reported "Error 0".
    Resume CleanExit
CleanExit:
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
    If c Is Nothing Then
        Err.Raise vbObjectError + 513, "REVO_FormBuilder", _
                  "Could not add control '" & nm & "' (" & progID & ")"
    End If

    ' Geometry is structural - if this fails the form is unusable, so let it
    ' raise. Fonts and styles are not, and are guarded at each call site.
    c.Left = l
    c.Top = t
    c.Width = w
    c.Height = h
    Set AddCtl = c
End Function

' Cosmetic property set that must never fail a build.
Private Sub Cosmetic(ByVal c As Object, ByVal prop As String, ByVal v As Variant)
    On Error Resume Next
    CallByName c, prop, VbLet, v
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub SetFont(ByVal c As Object, ByVal sz As Single, ByVal bold As Boolean)
    On Error Resume Next
    c.Font.name = "Segoe UI"
    c.Font.Size = sz
    c.Font.bold = bold
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub AddLabel(ByVal parent As Object, ByVal nm As String, ByVal cap As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                     ByVal h As Single, ByVal bold As Boolean)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.Label.1", nm, l, t, w, h)
    Cosmetic c, "Caption", cap
    SetFont c, 9, bold
    If bold Then Cosmetic c, "ForeColor", RGB(70, 70, 70)
End Sub

Private Sub AddBigLabel(ByVal parent As Object, ByVal nm As String, ByVal cap As String, _
                        ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.Label.1", nm, l, t, w, h)
    Cosmetic c, "Caption", cap
    SetFont c, 14, True
    Cosmetic c, "ForeColor", RGB(0, 97, 0)
End Sub

Private Sub AddTextBox(ByVal parent As Object, ByVal nm As String, _
                       ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.TextBox.1", nm, l, t, w, h)
    SetFont c, 9, False
End Sub

Private Sub AddCombo(ByVal parent As Object, ByVal nm As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.ComboBox.1", nm, l, t, w, h)
    SetFont c, 9, False
    ' Free text stays allowed - validation lives in the form, not the control.
    Cosmetic c, "MatchRequired", False
    Cosmetic c, "Style", 0          ' fmStyleDropDownCombo
End Sub

Private Function AddFrame(ByVal parent As Object, ByVal nm As String, ByVal cap As String, _
                          ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = AddCtl(parent, "Forms.Frame.1", nm, l, t, w, h)
    Cosmetic c, "Caption", cap
    SetFont c, 9, True
    Set AddFrame = c
End Function

Private Sub AddButton(ByVal parent As Object, ByVal nm As String, ByVal cap As String, _
                      ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                      ByVal h As Single, ByVal isDefault As Boolean)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.CommandButton.1", nm, l, t, w, h)
    Cosmetic c, "Caption", cap
    SetFont c, 9, isDefault
    Cosmetic c, "Default", isDefault
    Cosmetic c, "Cancel", Not isDefault
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
    LastBuildError = "Programmatic access to the VBA project is blocked."
    MsgBox "Excel is blocking programmatic access to the VBA project." & vbCrLf & vbCrLf & _
           "Tick this once and re-run:" & vbCrLf & vbCrLf & _
           "  File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
           "  Macro Settings > Trust access to the VBA project object model" & vbCrLf & vbCrLf & _
           "If your IT policy will not allow it, see revo/docs/MANUAL_FORM_BUILD.md " & _
           "for the control list to add by hand.", _
           vbExclamation, "Form Builder"
    VBAccessOK = False
    Resume Done
Done:
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
    Resume Done
Done:
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
