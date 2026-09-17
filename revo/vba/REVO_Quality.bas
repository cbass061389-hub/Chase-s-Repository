Attribute VB_Name = "REVO_Quality"
'==============================================================================
' REVO_Quality  -  structured quality capture
'
' Replaces the single free-text "Description" on Reject Log with a coded event
' model: every rework, B-grade and reject writes one row to Quality Log
' carrying defect type, location, detected-at op and root cause.
'
' The code lists live on the Quality_Codes sheet, not in this module. The
' UserForm reads its dropdowns from that sheet at run time, so adding a defect
' type is a sheet edit, not a code change. That is the whole point - the
' taxonomy will change and nobody should need the VBE to change it.
'
' Reject Log and Rework Tracker keep receiving their existing rows, so nothing
' downstream that reads them breaks. Quality Log is additive.
'==============================================================================
Option Explicit

Public Const QL_HEADER_ROW As Long = 1

' Quality Log column order. Changing this means changing WriteQualityEvent.
Public Const QLC_EVENTID  As Long = 1
Public Const QLC_DATE     As Long = 2
Public Const QLC_SHIFT    As Long = 3
Public Const QLC_SOURCE   As Long = 4
Public Const QLC_WO       As Long = 5
Public Const QLC_SKU      As Long = 6
Public Const QLC_CART     As Long = 7
Public Const QLC_DIA      As Long = 8
Public Const QLC_JOINT    As Long = 9
Public Const QLC_VP       As Long = 10
Public Const QLC_DISP     As Long = 11
Public Const QLC_QTY      As Long = 12
Public Const QLC_DEFECT   As Long = 13
Public Const QLC_LOCATION As Long = 14
Public Const QLC_OP       As Long = 15
Public Const QLC_ROOT     As Long = 16
Public Const QLC_ACTION   As Long = 17
Public Const QLC_MULTI    As Long = 18
Public Const QLC_NOTES    As Long = 19
Public Const QLC_USER     As Long = 20
Public Const QLC_RELID    As Long = 21
Public Const QL_COL_COUNT As Long = 21

'==============================================================================
' A single quality event, passed from the UserForm to the writer.
'==============================================================================
Public Type QualityEvent
    EventDate   As Date
    source      As String     ' RELEASE | REWORK RELEASE | BACKFILL | MANUAL
    WO          As String
    sku         As String
    Cart        As String
    Disposition As String     ' REWORK | B GRADE | REJECT
    Qty         As Double
    DefectType  As String
    Location    As String
    DetectedOp  As String
    RootCause   As String
    Action      As String
    MultiDefect As Boolean
    Notes       As String
    ReleaseID   As String
End Type

'==============================================================================
' SHEET PROVISIONING
'==============================================================================
Public Sub EnsureQualityLog()
    Dim ws As Worksheet
    Set ws = REVO_Core.GetOrCreateSheet(SH_QUAL_LOG)

    If Len(REVO_Core.SafeS(ws.Cells(QL_HEADER_ROW, 1).Value)) = 0 Then
        REVO_Core.WriteHeaders ws, QL_HEADER_ROW, Array( _
            "Event ID", "Date", "Shift", "Source", "WO", "SKU", "Cart", _
            "Diameter", "Joint", "VP", "Disposition", "Qty", _
            "Defect Type", "Defect Location", "Detected At Op", "Root Cause", _
            "Corrective Action", "Multi-Defect", "Notes", "Entered By", "Release ID")
        ws.Columns("A:A").ColumnWidth = 24
        ws.Columns("B:B").ColumnWidth = 11
        ws.Columns("F:F").ColumnWidth = 28
        ws.Columns("M:Q").ColumnWidth = 18
        ws.Columns("S:S").ColumnWidth = 40
    End If
End Sub

'==============================================================================
' CODE LISTS
'
' Seeded from the 549 historical reject rows and 16 hold reasons actually in
' this workbook. Edit the sheet freely afterwards; nothing here overwrites a
' populated Quality_Codes sheet.
'==============================================================================
Public Sub EnsureQualityCodes(Optional ByVal forceReseed As Boolean = False)
    Dim ws As Worksheet
    Dim defects As Variant, locations As Variant, roots As Variant, ops As Variant
    Dim actions As Variant
    Dim i As Long

    Set ws = REVO_Core.GetOrCreateSheet(SH_QUAL_CODES)

    If Not forceReseed And Len(REVO_Core.SafeS(ws.Cells(2, 1).Value)) > 0 Then Exit Sub

    ws.Cells.Clear

    ws.Range("A1").Value = "QUALITY CODES  -  edit these lists to change the UserForm dropdowns"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A2").Value = "The release form reads these columns live. Add a row and it appears in the dropdown next time the form opens. Do not rename the header row."
    ws.Range("A2").Font.Italic = True

    ws.Range("A4").Value = "Defect Type"
    ws.Range("B4").Value = "Defect Location"
    ws.Range("C4").Value = "Root Cause"
    ws.Range("D4").Value = "Detected At Op"
    ws.Range("E4").Value = "Corrective Action"
    ws.Range("A4:E4").Font.Bold = True
    ws.Range("A4:E4").Interior.Color = RGB(31, 56, 100)
    ws.Range("A4:E4").Font.Color = RGB(255, 255, 255)

    defects = Array("Chip", "Crack", "Breakthrough", "Collet Crush", "Splinter", _
                    "Plug Defect", "VP Gap", "Dimensional", "Cosmetic / Finish", _
                    "Bond / Adhesive", "Contamination", "Other", "Unclassified")

    locations = Array("Butt", "Tip", "Mid Shaft", "Joint", "Bore", "Not Specified")

    roots = Array("Operator Error", "Machine Setup", "Tooling / Fixture", _
                  "Material / Incoming", "Process / Method", "Maintenance", _
                  "Handling / Transport", "Supplier", "Not Determined")

    ops = Array("IC", "10", "20", "30", "40", "50", "60", "70", "100", "110", _
                "120", "130", "140", "160", "ASSYM", "FQC", "Unknown")

    actions = Array("Reworked in house", "Scrapped", "Downgraded to B", _
                    "Returned to supplier", "Machine adjusted", "Tooling replaced", _
                    "Operator retrained", "Process updated", "None", "Pending")

    For i = LBound(defects) To UBound(defects)
        ws.Cells(5 + i - LBound(defects), 1).Value = defects(i)
    Next i
    For i = LBound(locations) To UBound(locations)
        ws.Cells(5 + i - LBound(locations), 2).Value = locations(i)
    Next i
    For i = LBound(roots) To UBound(roots)
        ws.Cells(5 + i - LBound(roots), 3).Value = roots(i)
    Next i
    For i = LBound(ops) To UBound(ops)
        ws.Cells(5 + i - LBound(ops), 4).Value = ops(i)
    Next i
    For i = LBound(actions) To UBound(actions)
        ws.Cells(5 + i - LBound(actions), 5).Value = actions(i)
    Next i

    ws.Columns("A:E").ColumnWidth = 22

    '--- keyword map: drives backfill and the form's suggestion box ----------
    ws.Range("G4").Value = "Keyword"
    ws.Range("H4").Value = "Defect Type"
    ws.Range("I4").Value = "Root Cause"
    ws.Range("G4:I4").Font.Bold = True
    ws.Range("G4:I4").Interior.Color = RGB(31, 56, 100)
    ws.Range("G4:I4").Font.Color = RGB(255, 255, 255)
    ws.Range("G3").Value = "Classifier map - longest keyword wins. Used to code historical free text and to pre-select on the form."
    ws.Range("G3").Font.Italic = True

    ' Built one call at a time on purpose: VBA allows at most 24 line
    ' continuations in a single statement, and a nested Array() literal for
    ' this many rows blows straight past it and refuses to import.
    Dim kmRow As Long
    kmRow = 5

    AddKeyword ws, kmRow, "collet crush", "Collet Crush", "Tooling / Fixture"
    AddKeyword ws, kmRow, "cullet crush", "Collet Crush", "Tooling / Fixture"
    AddKeyword ws, kmRow, "collet", "Collet Crush", "Tooling / Fixture"
    AddKeyword ws, kmRow, "breakthrough", "Breakthrough", "Process / Method"
    AddKeyword ws, kmRow, "break through", "Breakthrough", "Process / Method"
    AddKeyword ws, kmRow, "bt at tip", "Breakthrough", "Process / Method"
    AddKeyword ws, kmRow, "bt butt", "Breakthrough", "Process / Method"
    AddKeyword ws, kmRow, "bt ", "Breakthrough", "Process / Method"
    AddKeyword ws, kmRow, "splintered", "Splinter", "Process / Method"
    AddKeyword ws, kmRow, "splinter", "Splinter", "Process / Method"
    AddKeyword ws, kmRow, "cracked", "Crack", "Not Determined"
    AddKeyword ws, kmRow, "crack", "Crack", "Not Determined"
    AddKeyword ws, kmRow, "chipped", "Chip", "Not Determined"
    AddKeyword ws, kmRow, "chip", "Chip", "Not Determined"
    AddKeyword ws, kmRow, "plugs pulling out", "Plug Defect", "Process / Method"
    AddKeyword ws, kmRow, "plug stuck", "Plug Defect", "Process / Method"
    AddKeyword ws, kmRow, "bore plugs", "Plug Defect", "Process / Method"
    AddKeyword ws, kmRow, "bore plug", "Plug Defect", "Process / Method"
    AddKeyword ws, kmRow, "plug", "Plug Defect", "Process / Method"
    AddKeyword ws, kmRow, "vp gaps", "VP Gap", "Process / Method"
    AddKeyword ws, kmRow, "vp gap", "VP Gap", "Process / Method"
    AddKeyword ws, kmRow, "tip gaps", "VP Gap", "Process / Method"
    AddKeyword ws, kmRow, "tip gap", "VP Gap", "Process / Method"
    AddKeyword ws, kmRow, "missing bushing", "Chip", "Tooling / Fixture"
    AddKeyword ws, kmRow, "operator error", "Other", "Operator Error"
    AddKeyword ws, kmRow, "tip problems", "Other", "Not Determined"
    AddKeyword ws, kmRow, "uid", "Other", "Not Determined"
    AddKeyword ws, kmRow, "uld", "Other", "Not Determined"
    AddKeyword ws, kmRow, "not reported", "Unclassified", "Not Determined"
    AddKeyword ws, kmRow, "n/a", "Unclassified", "Not Determined"

    ws.Columns("G:I").ColumnWidth = 22

    REVO_Core.Audit "REVO_Quality", "SEED", SH_QUAL_CODES, _
                    "Seeded defect/root-cause taxonomy"
End Sub

' Read one code column off Quality_Codes into a 1-based string array.
Public Function CodeList(ByVal col As Long) As String()
    Dim ws As Worksheet, r As Long, lastR As Long, n As Long
    Dim out() As String
    Dim v As String

    ReDim out(1 To 1)
    out(1) = ""

    Set ws = REVO_Core.GetSheet(SH_QUAL_CODES)
    If ws Is Nothing Then CodeList = out: Exit Function

    lastR = ws.Cells(ws.Rows.Count, col).End(xlUp).row
    If lastR < 5 Then CodeList = out: Exit Function

    ReDim out(1 To lastR - 4)
    For r = 5 To lastR
        v = REVO_Core.SafeS(ws.Cells(r, col).Value)
        If Len(v) > 0 Then
            n = n + 1
            out(n) = v
        End If
    Next r

    If n = 0 Then
        ReDim out(1 To 1)
        out(1) = ""
    Else
        ReDim Preserve out(1 To n)
    End If
    CodeList = out
End Function

Public Function DefectTypeList() As String()
    DefectTypeList = CodeList(1)
End Function
Public Function LocationList() As String()
    LocationList = CodeList(2)
End Function
Public Function RootCauseList() As String()
    RootCauseList = CodeList(3)
End Function
Public Function OpList() As String()
    OpList = CodeList(4)
End Function
Public Function ActionList() As String()
    ActionList = CodeList(5)
End Function

'==============================================================================
' CLASSIFIER
'
' Maps free text to a defect type / root cause using the keyword map on
' Quality_Codes. Longest keyword wins so "collet crush" beats "crush" and
' "bore plugs" beats "plug".
'==============================================================================
Public Sub ClassifyText(ByVal txt As String, _
                        ByRef outDefect As String, _
                        ByRef outRoot As String, _
                        ByRef outLocation As String, _
                        ByRef outOp As String, _
                        ByRef outMulti As Boolean)
    Dim ws As Worksheet, r As Long, lastR As Long
    Dim s As String, kw As String
    Dim bestLen As Long

    outDefect = "Unclassified"
    outRoot = "Not Determined"
    outLocation = "Not Specified"
    outOp = "Unknown"
    outMulti = False

    s = LCase$(" " & Trim$(txt) & " ")
    If Len(Trim$(txt)) = 0 Then Exit Sub

    '--- defect + root from the keyword map --------------------------------
    Set ws = REVO_Core.GetSheet(SH_QUAL_CODES)
    If Not ws Is Nothing Then
        lastR = ws.Cells(ws.Rows.Count, 7).End(xlUp).row
        For r = 5 To lastR
            kw = LCase$(REVO_Core.SafeS(ws.Cells(r, 7).Value))
            If Len(kw) > 0 Then
                If InStr(s, kw) > 0 And Len(kw) > bestLen Then
                    bestLen = Len(kw)
                    outDefect = REVO_Core.SafeS(ws.Cells(r, 8).Value, "Unclassified")
                    outRoot = REVO_Core.SafeS(ws.Cells(r, 9).Value, "Not Determined")
                End If
            End If
        Next r
    End If

    '--- "operator error" is a root cause even when a defect word also hits --
    If InStr(s, "operator error") > 0 Then outRoot = "Operator Error"
    If InStr(s, "missing bushing") > 0 Then outRoot = "Tooling / Fixture"

    '--- location -----------------------------------------------------------
    If InStr(s, "butt") > 0 And InStr(s, "tip") > 0 Then
        outLocation = "Not Specified"
        outMulti = True
    ElseIf InStr(s, "butt") > 0 Then
        outLocation = "Butt"
    ElseIf InStr(s, "tip") > 0 Then
        outLocation = "Tip"
    ElseIf InStr(s, "bore") > 0 Then
        outLocation = "Bore"
    ElseIf InStr(s, "joint") > 0 Then
        outLocation = "Joint"
    End If

    '--- detected-at op: first "op<digits>" token ---------------------------
    outOp = ExtractOp(s)

    '--- multiple defects called out in one description ---------------------
    If InStr(s, ";") > 0 Then outMulti = True
    If InStr(s, " and ") > 0 And bestLen > 0 Then outMulti = True
End Sub

' Pull an op number out of text such as "op10 chip", "splintered tip op100",
' "op 100 error". Returns "Unknown" when absent.
Public Function ExtractOp(ByVal s As String) As String
    Dim i As Long, j As Long, n As String
    Dim t As String

    t = LCase$(s)
    i = InStr(t, "op")
    Do While i > 0
        j = i + 2
        ' skip a single space between "op" and the number
        Do While j <= Len(t) And Mid$(t, j, 1) = " "
            j = j + 1
        Loop
        n = ""
        Do While j <= Len(t) And Mid$(t, j, 1) Like "#"
            n = n & Mid$(t, j, 1)
            j = j + 1
        Loop
        If Len(n) > 0 Then
            ExtractOp = CStr(CLng(n))
            Exit Function
        End If
        i = InStr(i + 1, t, "op")
    Loop
    ExtractOp = "Unknown"
End Function

'==============================================================================
' WRITER
'==============================================================================
Public Function WriteQualityEvent(ByRef ev As QualityEvent) As String
    Dim ws As Worksheet, r As Long
    Dim a As SkuAttrs
    Dim id As String

    EnsureQualityLog
    Set ws = REVO_Core.GetSheet(SH_QUAL_LOG)
    If ws Is Nothing Then Exit Function

    a = REVO_Core.ParseSku(ev.sku)
    id = REVO_Core.NewEventID("QE")

    r = ws.Cells(ws.Rows.Count, QLC_EVENTID).End(xlUp).row + 1
    If r <= QL_HEADER_ROW Then r = QL_HEADER_ROW + 1

    ws.Cells(r, QLC_EVENTID).Value = id
    ws.Cells(r, QLC_DATE).Value = ev.EventDate
    ws.Cells(r, QLC_DATE).NumberFormat = "m/d/yyyy"
    ws.Cells(r, QLC_SHIFT).Value = REVO_Core.ShiftFor(ev.EventDate)
    ws.Cells(r, QLC_SOURCE).Value = ev.source
    ws.Cells(r, QLC_WO).Value = ev.WO
    ws.Cells(r, QLC_SKU).Value = ev.sku
    ws.Cells(r, QLC_CART).Value = ev.Cart
    ws.Cells(r, QLC_DIA).Value = a.Diameter
    ws.Cells(r, QLC_JOINT).Value = a.Joint
    ws.Cells(r, QLC_VP).Value = a.VP
    ws.Cells(r, QLC_DISP).Value = ev.Disposition
    ws.Cells(r, QLC_QTY).Value = ev.Qty
    ws.Cells(r, QLC_DEFECT).Value = ev.DefectType
    ws.Cells(r, QLC_LOCATION).Value = ev.Location
    ws.Cells(r, QLC_OP).Value = ev.DetectedOp
    ws.Cells(r, QLC_ROOT).Value = ev.RootCause
    ws.Cells(r, QLC_ACTION).Value = ev.Action
    ws.Cells(r, QLC_MULTI).Value = ev.MultiDefect
    ws.Cells(r, QLC_NOTES).Value = ev.Notes
    ws.Cells(r, QLC_USER).Value = REVO_Core.CurrentUser
    ws.Cells(r, QLC_RELID).Value = ev.ReleaseID

    WriteQualityEvent = id
End Function

'==============================================================================
' BACKFILL
'
' Codes the historical free text on Reject Log and Rework Tracker into Quality
' Log so the dashboard has history on day one rather than starting empty.
'
' Idempotent: rows already carrying a BACKFILL event for the same
' date+cart+sku+qty are skipped, so running it twice does not double-count.
'==============================================================================
Public Function BackfillQualityLog(Optional ByVal showResult As Boolean = True) As Long
    Dim wsRej As Worksheet, wsRew As Worksheet, wsQL As Worksheet
    Dim hdr As Long, lastR As Long, r As Long
    Dim colDate As Long, cSku As Long, cWO As Long, cCart As Long, cQty As Long, cDesc As Long
    Dim seen As Object
    Dim ev As QualityEvent
    Dim added As Long, skipped As Long
    Dim k As String
    Dim dt As Date, txt As String
    Dim dDefect As String, dRoot As String, dLoc As String, dOp As String
    Dim dMulti As Boolean

    EnsureQualityCodes
    EnsureQualityLog

    Set wsQL = REVO_Core.GetSheet(SH_QUAL_LOG)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    '--- index what backfill has already written ---------------------------
    lastR = wsQL.Cells(wsQL.Rows.Count, QLC_EVENTID).End(xlUp).row
    For r = QL_HEADER_ROW + 1 To lastR
        If UCase$(REVO_Core.SafeS(wsQL.Cells(r, QLC_SOURCE).Value)) = "BACKFILL" Then
            k = BackfillKey(REVO_Core.SafeDate(wsQL.Cells(r, QLC_DATE).Value), _
                            REVO_Core.SafeS(wsQL.Cells(r, QLC_SKU).Value), _
                            REVO_Core.SafeS(wsQL.Cells(r, QLC_CART).Value), _
                            REVO_Core.SafeD(wsQL.Cells(r, QLC_QTY).Value), _
                            REVO_Core.SafeS(wsQL.Cells(r, QLC_DISP).Value))
            If Not seen.exists(k) Then seen.Add k, True
        End If
    Next r

    REVO_Core.PushAppState

    '======================= Reject Log =====================================
    Set wsRej = REVO_Core.GetSheet(SH_REJECT_LOG)
    If Not wsRej Is Nothing Then
        hdr = REVO_Core.FindHeaderRow(wsRej, Array("Date", "SKU"), 10, 12)
        If hdr > 0 Then
            colDate = REVO_Core.FindCol(wsRej, hdr, Array("Date"))
            cSku = REVO_Core.FindCol(wsRej, hdr, Array("SKU"))
            cWO = REVO_Core.FindCol(wsRej, hdr, Array("WO#", "WO", "Work Order"))
            cCart = REVO_Core.FindCol(wsRej, hdr, Array("Cart"))
            cQty = REVO_Core.FindCol(wsRej, hdr, Array("Quantity", "Qty"))
            cDesc = REVO_Core.FindCol(wsRej, hdr, Array("Description", "Reason", "Reject Reason"))

            If colDate > 0 And cSku > 0 And cQty > 0 Then
                lastR = REVO_Core.LastDataRow(wsRej, cSku, hdr)
                For r = hdr + 1 To lastR
                    dt = REVO_Core.SafeDate(wsRej.Cells(r, colDate).Value, 0)
                    If dt > 0 Then
                        txt = IIf(cDesc > 0, REVO_Core.SafeS(wsRej.Cells(r, cDesc).Value), "")
                        ev.EventDate = dt
                        ev.source = "BACKFILL"
                        ev.WO = IIf(cWO > 0, REVO_Core.SafeS(wsRej.Cells(r, cWO).Value), "")
                        ev.sku = REVO_Core.SafeS(wsRej.Cells(r, cSku).Value)
                        ev.Cart = IIf(cCart > 0, REVO_Core.SafeS(wsRej.Cells(r, cCart).Value), "")
                        ev.Disposition = "REJECT"
                        ev.Qty = REVO_Core.SafeD(wsRej.Cells(r, cQty).Value)
                        ev.ReleaseID = ""
                        ev.Action = ""

                        k = BackfillKey(dt, ev.sku, ev.Cart, ev.Qty, "REJECT")
                        If seen.exists(k) Then
                            skipped = skipped + 1
                        Else
                            ClassifyText txt, dDefect, dRoot, dLoc, dOp, dMulti
                            ev.DefectType = dDefect
                            ev.RootCause = dRoot
                            ev.Location = dLoc
                            ev.DetectedOp = dOp
                            ev.MultiDefect = dMulti
                            ev.Notes = txt
                            WriteQualityEvent ev
                            seen.Add k, True
                            added = added + 1
                        End If
                    End If
                Next r
            End If
        End If
    End If

    '======================= Rework Tracker =================================
    Set wsRew = REVO_Core.GetSheet(SH_REWORK)
    If Not wsRew Is Nothing Then
        hdr = REVO_Core.FindHeaderRow(wsRew, Array("Date", "SKU"), 10, 12)
        If hdr > 0 Then
            colDate = REVO_Core.FindCol(wsRew, hdr, Array("Date"))
            cSku = REVO_Core.FindCol(wsRew, hdr, Array("SKU"))
            cWO = REVO_Core.FindCol(wsRew, hdr, Array("WO", "WO#", "Work Order"))
            cCart = REVO_Core.FindCol(wsRew, hdr, Array("Cart"))
            cQty = REVO_Core.FindCol(wsRew, hdr, Array("Qty", "Quantity"))
            cDesc = REVO_Core.FindCol(wsRew, hdr, Array("Status", "Reason", "Notes"))

            If colDate > 0 And cSku > 0 And cQty > 0 Then
                lastR = REVO_Core.LastDataRow(wsRew, cSku, hdr)
                For r = hdr + 1 To lastR
                    dt = REVO_Core.SafeDate(wsRew.Cells(r, colDate).Value, 0)
                    If dt > 0 Then
                        txt = IIf(cDesc > 0, REVO_Core.SafeS(wsRew.Cells(r, cDesc).Value), "")
                        ev.EventDate = dt
                        ev.source = "BACKFILL"
                        ev.WO = IIf(cWO > 0, REVO_Core.SafeS(wsRew.Cells(r, cWO).Value), "")
                        ev.sku = REVO_Core.SafeS(wsRew.Cells(r, cSku).Value)
                        ev.Cart = IIf(cCart > 0, REVO_Core.SafeS(wsRew.Cells(r, cCart).Value), "")
                        ev.Disposition = "REWORK"
                        ev.Qty = REVO_Core.SafeD(wsRew.Cells(r, cQty).Value)
                        ev.ReleaseID = ""
                        ev.Action = ""

                        k = BackfillKey(dt, ev.sku, ev.Cart, ev.Qty, "REWORK")
                        If seen.exists(k) Then
                            skipped = skipped + 1
                        Else
                            ' Rework "Status" is workflow state, not a defect
                            ' description, so it seeds notes only.
                            ClassifyText txt, dDefect, dRoot, dLoc, dOp, dMulti
                            ev.DefectType = dDefect
                            ev.RootCause = dRoot
                            ev.Location = dLoc
                            ev.DetectedOp = ReworkOpFromRouting(wsRew, hdr, r)
                            ev.MultiDefect = dMulti
                            ev.Notes = txt
                            WriteQualityEvent ev
                            seen.Add k, True
                            added = added + 1
                        End If
                    End If
                Next r
            End If
        End If
    End If

    REVO_Core.PopAppState
    REVO_Core.Audit "REVO_Quality", "BACKFILL", SH_QUAL_LOG, _
                    added & " events added, " & skipped & " already present"

    BackfillQualityLog = added

    If showResult Then
        MsgBox "Quality Log backfill complete." & vbCrLf & vbCrLf & _
               "Events added:      " & added & vbCrLf & _
               "Already present:   " & skipped & vbCrLf & vbCrLf & _
               "Historical free text has been coded to defect type and root cause. " & _
               "Anything the classifier could not read is marked Unclassified - " & _
               "filter Quality Log on that to clean it up.", _
               vbInformation, "REVO Quality"
    End If
End Function

' The rework routing columns (10..160) are booleans marking which ops a cart
' must revisit. The first ticked op is where the rework starts.
Private Function ReworkOpFromRouting(ByVal ws As Worksheet, ByVal hdr As Long, _
                                     ByVal r As Long) As String
    Dim c As Long, lastCol As Long, capt As String
    lastCol = ws.Cells(hdr, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        capt = REVO_Core.SafeS(ws.Cells(hdr, c).Value)
        If Len(capt) > 0 And IsNumeric(capt) Then
            If REVO_Core.Boolish(ws.Cells(r, c).Value) Then
                ReworkOpFromRouting = capt
                Exit Function
            End If
        End If
    Next c
    ReworkOpFromRouting = "Unknown"
End Function

Private Function BackfillKey(ByVal d As Date, ByVal sku As String, ByVal cart As String, _
                             ByVal qty As Double, ByVal disp As String) As String
    BackfillKey = Format$(d, "yyyymmdd") & "|" & UCase$(Trim$(sku)) & "|" & _
                  UCase$(Trim$(cart)) & "|" & CStr(qty) & "|" & UCase$(Trim$(disp))
End Function

' Writes one classifier row and advances the cursor. Keeps the seeding code
' free of line continuations.
Private Sub AddKeyword(ByVal ws As Worksheet, ByRef r As Long, ByVal keyword As String, _
                       ByVal defect As String, ByVal root As String)
    ws.Cells(r, 7).Value = keyword
    ws.Cells(r, 8).Value = defect
    ws.Cells(r, 9).Value = root
    r = r + 1
End Sub
