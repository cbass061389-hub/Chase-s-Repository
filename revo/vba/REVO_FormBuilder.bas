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

    ' Armed FIRST. File probing happens below, and an unguarded Dir$ against a
    ' OneDrive / SharePoint path raises error 52 before any handler existed -
    ' which escaped to the installer and aborted the whole run.
    On Error GoTo BuildFailed
    LastBuildError = ""

    If Not VBAccessOK() Then
        LastBuildError = "Programmatic access to the VBA project is blocked."
        Exit Sub
    End If

    step_ = "locating " & CODE_FILE
    codePath = ResolveCodePath(codePath)

    step_ = "reading the form code"
    If Len(codePath) = 0 Then
        code = ReadCachedCode()
    Else
        code = ReadTextFile(codePath)
        If Len(code) = 0 Then code = ReadCachedCode()
    End If

    If Len(code) = 0 Then
        LastBuildError = "Could not find or read " & CODE_FILE & "."
        MsgBox "Could not read the form code." & vbCrLf & vbCrLf & _
               "Looked beside the workbook, in revo\vba, in Downloads and on the " & _
               "Desktop, and in the cached copy inside this workbook." & vbCrLf & vbCrLf & _
               "Run REVO_BuildReleaseForm again and point the file picker at " & _
               CODE_FILE & ".", vbExclamation, "Form Builder"
        Exit Sub
    End If

    step_ = "opening the VBA project"
    Set vbp = ThisWorkbook.VBProject

    '--- start clean --------------------------------------------------------
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents(FORM_NAME)
    Err.Clear
    On Error GoTo BuildFailed

    step_ = "creating the UserForm component"
    Set vbc = vbp.VBComponents.Add(VB_EXT_CT_MSFORM)
    vbc.name = FORM_NAME

    ' Caption / Width / Height go through the VBComponent Properties collection,
    ' which is not exposed identically across Excel builds. Cosmetic only - a
    ' failure here must not cost us the form.
    On Error Resume Next
    vbc.Properties("Caption") = "Release cart"
    vbc.Properties("Width") = 684
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
    AddLabel dsn, "lblHdr7", "Release qty this pass", 12, 52, 122, 16, True
    AddTextBox dsn, "txtCartQtyRelease", 138, 50, 56, 20

    AddLabel dsn, "lblHdr8", "To inventory", 216, 52, 72, 16, True
    AddBigLabel dsn, "lblToInventory", "0", 292, 48, 66, 22

    AddLabel dsn, "lblAfterThisPass", "", 372, 52, 290, 16, False
    AddLabel dsn, "lblWarn", "", 12, 72, 650, 14, False

    '========================= disposition quantities =======================
    step_ = "adding the disposition quantity controls"
    AddLabel dsn, "lblHdr9", "Rework", 12, 94, 46, 16, True
    AddTextBox dsn, "txtRework", 60, 92, 46, 20
    AddLabel dsn, "lblHdr10", "B Grade", 128, 94, 50, 16, True
    AddTextBox dsn, "txtBGrade", 180, 92, 46, 20
    AddLabel dsn, "lblHdr11", "Reject", 248, 94, 40, 16, True
    AddTextBox dsn, "txtReject", 290, 92, 46, 20
    AddLabel dsn, "lblAllocation", "", 348, 94, 314, 16, True

    '========================= quality detail entry =========================
    step_ = "adding the quality detail entry row"
    Set fra = AddFrame(dsn, "fraDetail", _
        "Quality detail  -  one line per cause. Split a disposition across as many causes as it needs.", _
        8, 118, 656, 118)

    AddLabel fra, "lblEnD", "Disposition", 8, 18, 62, 14, True
    AddCombo fra, "cboDisposition", 72, 16, 86, 18
    AddLabel fra, "lblEnQ", "Qty", 168, 18, 22, 14, True
    AddTextBox fra, "txtLineQty", 192, 16, 44, 20
    AddLabel fra, "lblEnDef", "Defect", 250, 18, 40, 14, True
    AddCombo fra, "cboDefect", 292, 16, 126, 18
    AddLabel fra, "lblEnL", "Location", 428, 18, 48, 14, True
    AddCombo fra, "cboLocation", 478, 16, 104, 18

    AddLabel fra, "lblEnO", "At op", 8, 44, 34, 14, True
    AddCombo fra, "cboOp", 72, 42, 86, 18
    AddLabel fra, "lblEnR", "Root cause", 168, 44, 62, 14, True
    AddCombo fra, "cboRoot", 232, 42, 136, 18
    AddLabel fra, "lblEnA", "Action", 380, 44, 40, 14, True
    AddCombo fra, "cboAction", 424, 42, 158, 18

    AddLabel fra, "lblEnN", "Notes", 8, 70, 34, 14, True
    AddTextBox fra, "txtLineNotes", 72, 68, 510, 18

    AddButton fra, "btnAddLine", "Add line", 8, 92, 86, 22, False
    AddButton fra, "btnRemoveLine", "Remove selected", 100, 92, 110, 22, False
    AddLabel fra, "lblEnHint", _
        "Example: reject 10 = 3 collet crush, 5 chip at butt, 2 breakthrough. Add each as its own line.", _
        220, 96, 420, 14, False

    '========================= the lines ====================================
    step_ = "adding the detail list"
    AddLabel dsn, "lblHdrList", "Lines recorded for this release", 12, 244, 200, 14, True
    AddListBox dsn, "lstLines", 8, 260, 656, 150

    '========================= buttons ======================================
    step_ = "adding the buttons"
    AddButton dsn, "btnSubmit", "Submit release", 446, 420, 108, 28, True
    AddButton dsn, "btnCancel", "Cancel", 564, 420, 100, 28, False

    '========================= code =========================================
    step_ = "injecting the form code"
    With vbc.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString code
    End With

    REVO_Core.Audit "REVO_FormBuilder", "BUILD", FORM_NAME, "Rebuilt from " & codePath

    MsgBox "frmReleaseDetails rebuilt." & vbCrLf & vbCrLf & _
           "Release qty, rework / B grade / reject totals, and a quality detail list " & _
           "that takes as many cause lines as a disposition needs - so 10 rejects can " & _
           "be 3 of one defect, 5 of another and 2 of a third." & vbCrLf & vbCrLf & _
           "Submit is blocked until each disposition's lines add up to its total." & vbCrLf & vbCrLf & _
           "Dropdowns read from the " & SH_QUAL_CODES & " sheet every time the form opens.", _
           vbInformation, "Form Builder"
    LastBuildError = ""
    CacheCode code
    Exit Sub

BuildFailed:
    ' Capture Err FIRST - anything called below can clear it.
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    LastBuildError = "Error " & errNum & ": " & errDesc & "  (step: " & step_ & ")"
    REVO_Core.Audit "REVO_FormBuilder", "ERROR", FORM_NAME, LastBuildError

    MsgBox "Building the form failed." & vbCrLf & vbCrLf & _
           "Error " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
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

Private Sub AddListBox(ByVal parent As Object, ByVal nm As String, _
                      ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = AddCtl(parent, "Forms.ListBox.1", nm, l, t, w, h)
    SetFont c, 9, False
    Cosmetic c, "ColumnCount", 7
    Cosmetic c, "ColumnWidths", "58 pt;30 pt;90 pt;62 pt;38 pt;96 pt;120 pt"
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

' Does this path point at a readable file? Never raises - Dir$ throws error 52
' on a URL, a disconnected drive, or a malformed path, and a probe must not.
Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    If Len(Trim$(path)) = 0 Then
        FileExists = False
    ElseIf IsUrlPath(path) Then
        FileExists = False          ' Dir$ cannot read http/https at all
    Else
        FileExists = (Len(Dir$(path)) > 0)
    End If
    If Err.Number <> 0 Then FileExists = False
    Err.Clear
    On Error GoTo 0
End Function

' A workbook opened from OneDrive or SharePoint reports an https:// path.
Public Function IsUrlPath(ByVal path As String) As Boolean
    Dim p As String
    p = LCase$(Trim$(path))
    IsUrlPath = (Left$(p, 7) = "http://" Or Left$(p, 8) = "https://")
End Function

Private Function Joined(ByVal folder As String, ByVal leaf As String) As String
    If Len(folder) = 0 Then Exit Function
    If Right$(folder, 1) = Application.PathSeparator Then
        Joined = folder & leaf
    Else
        Joined = folder & Application.PathSeparator & leaf
    End If
End Function

' Every place the form code plausibly is, cheapest first. Nothing here can
' raise; each candidate is probed through FileExists.
Private Function ResolveCodePath(ByVal supplied As String) As String
    Dim fd As Object
    Dim wbPath As String, userPath As String
    Dim cand As String
    Dim i As Long
    Dim folders(1 To 8) As String

    If FileExists(supplied) Then ResolveCodePath = supplied: Exit Function

    On Error Resume Next
    wbPath = ThisWorkbook.path
    userPath = Environ$("USERPROFILE")
    Err.Clear
    On Error GoTo 0

    ' A OneDrive / SharePoint workbook has no usable local folder of its own.
    If IsUrlPath(wbPath) Then wbPath = ""

    folders(1) = wbPath
    folders(2) = Joined(Joined(wbPath, "revo"), "vba")
    folders(3) = Joined(userPath, "Downloads")
    folders(4) = Joined(Joined(userPath, "Downloads"), "revo")
    folders(5) = Joined(Joined(Joined(userPath, "Downloads"), "revo"), "vba")
    folders(6) = Joined(userPath, "Desktop")
    folders(7) = Joined(Joined(userPath, "Desktop"), "vba")
    folders(8) = CurDirSafe()

    For i = 1 To 8
        If Len(folders(i)) > 0 Then
            cand = Joined(folders(i), CODE_FILE)
            If FileExists(cand) Then ResolveCodePath = cand: Exit Function
        End If
    Next i

    ' Nothing found. The cached copy inside the workbook is tried by the caller;
    ' offer the picker first because a real file is always preferable.
    On Error Resume Next
    Set fd = Application.FileDialog(3)      ' msoFileDialogFilePicker
    If Not fd Is Nothing Then
        fd.Title = "Locate " & CODE_FILE
        fd.Filters.Clear
        fd.Filters.Add "Form code", "*.vb; *.txt; *.bas"
        fd.Filters.Add "All files", "*.*"
        If fd.Show = -1 Then ResolveCodePath = fd.SelectedItems(1)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function CurDirSafe() As String
    On Error Resume Next
    CurDirSafe = CurDir$
    Err.Clear
    On Error GoTo 0
End Function

'==============================================================================
' CACHED SOURCE
'
' After a successful build the form code is stored on a veryHidden sheet, so a
' rebuild works even when the original file has moved, the workbook lives on
' SharePoint, or someone opens it on a machine that never had the repo. The
' file on disk stays the source of truth when it is reachable.
'==============================================================================
Private Const SH_FORM_SRC As String = "REVO_FormSrc"

Private Sub CacheCode(ByVal code As String)
    Dim ws As Worksheet
    Dim parts As Variant
    Dim i As Long

    On Error GoTo GiveUp
    Set ws = REVO_Core.GetOrCreateSheet(SH_FORM_SRC)
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Cached source of " & CODE_FILE & " - written by REVO_BuildReleaseForm. Do not edit."

    parts = Split(Replace$(code, vbCrLf, vbLf), vbLf)
    For i = LBound(parts) To UBound(parts)
        ' Leading apostrophe keeps Excel from interpreting a line as a formula.
        ws.Cells(i + 2, 1).Value = "'" & CStr(parts(i))
    Next i

    ws.Visible = xlSheetVeryHidden
    Exit Sub
GiveUp:
    Resume Done
Done:
End Sub

Private Function ReadCachedCode() As String
    Dim ws As Worksheet
    Dim lastR As Long, i As Long
    Dim sb As String, ln As String

    On Error GoTo GiveUp
    Set ws = REVO_Core.GetSheet(SH_FORM_SRC)
    If ws Is Nothing Then Exit Function

    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastR < 2 Then Exit Function

    For i = 2 To lastR
        ln = CStr(ws.Cells(i, 1).Value)
        If Len(ln) > 0 Then
            If Left$(ln, 1) = "'" Then ln = Mid$(ln, 2)
        End If
        sb = sb & ln & vbCrLf
    Next i

    ReadCachedCode = sb
    Exit Function
GiveUp:
    ReadCachedCode = ""
    Resume Done
Done:
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
