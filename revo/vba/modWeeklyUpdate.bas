Attribute VB_Name = "modWeeklyUpdate"
Option Explicit

'======================================================================
' modWeeklyUpdate  -  builds Osvaldo's Monday REVO Operations Update
'
' Depends on: modConfig
'
' Public:
'   BuildWeeklyInputSheet  - creates Weekly_Update_Input. RUN THIS FIRST.
'                            Nothing else in this module works without it.
'   RefreshWeeklyMetrics   - fills the yellow cells from Release Log
'                            and Reject_Log_V2 for the stated week.
'   BuildWeeklyUpdate      - builds the HTML and OPENS it in Outlook.
'                            It never sends. Osvaldo presses Send.
'
' Section order matches the emails already going out:
'   Summary, Safety, Quality, Throughput, Next Four Weeks,
'   2026 Weekly Plan vs. Actual, Human Resources, Machinery, Supply
'
' Summary prints as paragraphs. Every other section prints as bullets.
'
'----------------------------------------------------------------------
' RELEASE LOG SEMANTICS  -  verified against the live data, not assumed
'
'   A Date | B SKU | C Cart | D Qty Released | E Work Order
'   F Qty Rework | G Qty B Grade | H Qty Reject
'   I Recalled | J Cart Qty Released | K Cart Remaining
'
' D is NET. On every row where J is populated, D + F + G + H = J exactly
' (row 758: 94 + 0 + 2 + 5 = 101; row 762: 52 + 0 + 5 + 3 = 60). So:
'
'   Total processed off the cart = D + F + G + H
'   A-Grade released to inventory = D          <- NOT D - G
'
' The first draft computed A-Grade as "released minus B-Grade", which
' subtracted the B-Grade quantity a second time and understated A-Grade
' by exactly that amount in the executive email. Fixed here.
'
' Reject rate now divides by total processed, which is what "reject rate"
' means on the floor. {REJECTRATE_REL} exposes the old released-only
' basis if the historical trend line has to stay comparable.
'======================================================================

' Input sheet geography - BuildWeeklyInputSheet owns this layout.
Private Const R_WEEKSTART As Long = 6
Private Const R_WEEKEND   As Long = 7
Private Const R_FROM      As Long = 8
Private Const R_TO        As Long = 9
Private Const R_CC        As Long = 10
Private Const R_SIGTITLE  As Long = 11
Private Const R_SIGADDR   As Long = 12
Private Const R_SIGPHONE  As Long = 13

Private Const R_PROCESSED As Long = 14
Private Const R_RELEASED  As Long = 15      ' = A-Grade
Private Const R_BGRADE    As Long = 16
Private Const R_REWORK    As Long = 17
Private Const R_REJECT    As Long = 18
Private Const R_REJRATE   As Long = 19
Private Const R_PLAN      As Long = 20
Private Const R_VARIANCE  As Long = 21
Private Const R_WEEKS     As Long = 22
Private Const R_TOPDEFECT As Long = 23

Private Const R_NARRHDR   As Long = 27
Private Const R_NARRATIVE As Long = 28

Private Const SECTION_ORDER As String = "Summary|Safety|Quality|Throughput|Next Four Weeks|2026 Weekly Plan vs. Actual|Human Resources|Machinery|Supply"

Private Const TO_PLACEHOLDER As String = "REPLACE-WITH-THE-NINE-RECIPIENTS@predatorgroup.com"

'======================================================================
' INPUT SHEET
'======================================================================

Public Sub BuildWeeklyInputSheet()

    Dim ws As Worksheet
    Dim secs As Variant
    Dim i As Long, r As Long

    On Error GoTo Fail
    FastOn

    Set ws = AddSheetAfter(SH_WEEKLYINPUT, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count).Name)

    If Len(ws.Range("B1").Value) = 0 Then

        ws.Range("B1").Value = "REVO WEEKLY OPERATIONS UPDATE - INPUT"
        ws.Range("B1").Font.Size = 14
        ws.Range("B1").Font.Bold = True
        ws.Range("B3").Value = "Yellow cells are yours. Grey cells are calculated - run RefreshWeeklyMetrics, don't type in them."
        ws.Range("B3").Font.Italic = True

        ws.Range("B5").Value = "EMAIL"
        StyleTitle ws.Range("B5:C5")
        Lbl ws, R_WEEKSTART, "Week start"
        Lbl ws, R_WEEKEND, "Week end"
        Lbl ws, R_FROM, "Sender name"
        Lbl ws, R_TO, "To (semicolon separated)"
        Lbl ws, R_CC, "CC (semicolon separated)"
        Lbl ws, R_SIGTITLE, "Signature title"
        Lbl ws, R_SIGADDR, "Signature address (use | between lines)"
        Lbl ws, R_SIGPHONE, "Signature phone"

        ws.Cells(R_WEEKSTART, 3).Value = Date - Weekday(Date, vbMonday) - 6
        ws.Cells(R_WEEKEND, 3).Value = Date - Weekday(Date, vbMonday)
        ws.Cells(R_WEEKSTART, 3).NumberFormat = "yyyy-mm-dd"
        ws.Cells(R_WEEKEND, 3).NumberFormat = "yyyy-mm-dd"
        ws.Cells(R_FROM, 3).Value = "Osvaldo Santiago"
        ws.Cells(R_TO, 3).Value = TO_PLACEHOLDER
        ws.Cells(R_CC, 3).Value = vbNullString
        ws.Cells(R_SIGTITLE, 3).Value = "Director, Supply Chain & Manufacturing"
        ws.Cells(R_SIGADDR, 3).Value = "4901 Belfort Rd. " & ChrW(8226) & " Suite 100|Jacksonville " & ChrW(8226) & " Florida " & ChrW(8226) & " 32256 " & ChrW(8226) & " USA"
        ws.Cells(R_SIGPHONE, 3).Value = "Corporate Office: +1 904.448.8748"
        Yellow ws.Range(ws.Cells(R_WEEKSTART, 3), ws.Cells(R_SIGPHONE, 3))

        Lbl ws, R_PROCESSED, "Total processed off cart"
        Lbl ws, R_RELEASED, "Released to inventory (A-Grade)"
        Lbl ws, R_BGRADE, "B-Grade"
        Lbl ws, R_REWORK, "Rework"
        Lbl ws, R_REJECT, "Reject"
        Lbl ws, R_REJRATE, "Reject rate (of processed)"
        Lbl ws, R_PLAN, "Weekly plan"
        Lbl ws, R_VARIANCE, "Variance to plan"
        Lbl ws, R_WEEKS, "Weeks remaining"
        Lbl ws, R_TOPDEFECT, "Top defect"
        Grey ws.Range(ws.Cells(R_PROCESSED, 3), ws.Cells(R_REJRATE, 3))
        Yellow ws.Range(ws.Cells(R_PLAN, 3), ws.Cells(R_PLAN, 3))
        Grey ws.Range(ws.Cells(R_VARIANCE, 3), ws.Cells(R_VARIANCE, 3))
        Yellow ws.Range(ws.Cells(R_WEEKS, 3), ws.Cells(R_WEEKS, 3))
        Grey ws.Range(ws.Cells(R_TOPDEFECT, 3), ws.Cells(R_TOPDEFECT, 3))
        ws.Cells(R_REJRATE, 3).NumberFormat = "0.00%"
        ws.Range(ws.Cells(R_PROCESSED, 3), ws.Cells(R_REJECT, 3)).NumberFormat = "#,##0"
        ws.Cells(R_PLAN, 3).NumberFormat = "#,##0"
        ws.Cells(R_VARIANCE, 3).NumberFormat = "#,##0;[Red]-#,##0"

        ' --- narrative -------------------------------------------------
        ws.Cells(R_NARRHDR, 2).Resize(1, 3).Value = Array("Section", "Text", "Include")
        StyleHeader ws.Cells(R_NARRHDR, 2).Resize(1, 3)

        secs = Split(SECTION_ORDER, "|")
        r = R_NARRATIVE
        For i = LBound(secs) To UBound(secs)
            ws.Cells(r, 2).Value = secs(i)
            ws.Cells(r, 3).Value = SeedText(CStr(secs(i)))
            ws.Cells(r, 4).Value = True
            r = r + 1
            ' two spare rows per section so bullets can be added without inserting
            ws.Cells(r, 2).Value = secs(i): ws.Cells(r, 4).Value = False: r = r + 1
            ws.Cells(r, 2).Value = secs(i): ws.Cells(r, 4).Value = False: r = r + 1
        Next i
        AddListValidation ws.Range("D" & R_NARRATIVE & ":D" & (r + 200)), "TRUE,FALSE"
        AddListValidation ws.Range("B" & R_NARRATIVE & ":B" & (r + 200)), Replace(SECTION_ORDER, "|", ",")
        Yellow ws.Range("C" & R_NARRATIVE & ":C" & (r - 1))

        ' --- token cheat sheet, parked out of the way -------------------
        ws.Range("F5").Value = "TOKENS - type these in the Text column"
        ws.Range("F5").Font.Bold = True
        WriteTokens ws, 6

        ws.Columns("B:B").ColumnWidth = 34
        ws.Columns("C:C").ColumnWidth = 90
        ws.Columns("D:D").ColumnWidth = 9
        ws.Columns("F:G").AutoFit
        ws.Rows(R_NARRATIVE & ":" & (r - 1)).VerticalAlignment = xlTop
        ws.Range("C" & R_NARRATIVE & ":C" & (r - 1)).WrapText = True
    End If

    FastOff
    LogLine "BuildWeeklyInputSheet: " & SH_WEEKLYINPUT & " ready"

    If InStr(1, CStr(ws.Cells(R_TO, 3).Value), "REPLACE", vbTextCompare) > 0 Then
        Say SH_WEEKLYINPUT & " built." & vbCrLf & vbCrLf & _
            "ACTION REQUIRED: cell C" & R_TO & " still holds a placeholder." & vbCrLf & _
            "Put the nine recipient addresses there, semicolon separated, before " & _
            "running BuildWeeklyUpdate.", "REVO Weekly Update", vbExclamation
    Else
        Say SH_WEEKLYINPUT & " built.", "REVO Weekly Update"
    End If
    Exit Sub

Fail:
    FastReset
    LogLine "*** BuildWeeklyInputSheet failed: " & Err.Description
    Say "Build failed: " & Err.Description, "REVO Weekly Update", vbCritical
End Sub

Private Function SeedText(ByVal sec As String) As String
    Select Case sec
        Case "Summary"
            SeedText = "We processed {PROCESSED} shafts in the week of {WEEKSTART} - {WEEKEND} and released " & _
                       "{RELEASED} to inventory against a plan of {PLAN}, a variance of {VARIANCE}. " & _
                       "Reject rate was {REJECTRATE}."
        Case "Quality"
            SeedText = "Reject {REJECT}, rework {REWORK}, B-Grade {BGRADE}. Top defect this week was {TOPDEFECT}."
        Case "Throughput"
            SeedText = "Released {RELEASED} against plan {PLAN} ({VARIANCE})."
        Case "2026 Weekly Plan vs. Actual"
            SeedText = "{WEEKS} weeks remaining on the 2026 plan."
        Case "Safety"
            SeedText = "No recordable incidents."
        Case Else
            SeedText = vbNullString
    End Select
End Function

Private Sub WriteTokens(ByVal ws As Worksheet, ByVal r As Long)
    Dim t As Variant, i As Long
    t = Array( _
        "{PROCESSED}", "Total off the cart (D+F+G+H)", _
        "{RELEASED}", "Released to inventory (A-Grade)", _
        "{AGRADE}", "Same as {RELEASED}", _
        "{BGRADE}", "B-Grade quantity", _
        "{REWORK}", "Rework quantity", _
        "{REJECT}", "Reject quantity", _
        "{REJECTRATE}", "Reject / total processed", _
        "{REJECTRATE_REL}", "Reject / released (legacy basis)", _
        "{PLAN}", "Weekly plan (you type it)", _
        "{VARIANCE}", "Released minus plan", _
        "{WEEKS}", "Weeks remaining (you type it)", _
        "{TOPDEFECT}", "Highest-qty defect code", _
        "{WEEKSTART}", "Week start, m/d", _
        "{WEEKEND}", "Week end, m/d")
    For i = LBound(t) To UBound(t) Step 2
        ws.Cells(r, 6).Value = t(i)
        ws.Cells(r, 7).Value = t(i + 1)
        r = r + 1
    Next i
End Sub

Private Sub Lbl(ByVal ws As Worksheet, ByVal r As Long, ByVal s As String)
    ws.Cells(r, 2).Value = s
    ws.Cells(r, 2).Font.Bold = True
End Sub

Private Sub Yellow(ByVal rng As Range)
    rng.Interior.Color = RGB(255, 242, 204)
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(191, 191, 191)
End Sub

Private Sub Grey(ByVal rng As Range)
    rng.Interior.Color = RGB(233, 233, 233)
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(191, 191, 191)
End Sub

Private Sub StyleHeader(ByVal rng As Range)
    rng.Font.Bold = True
    rng.Interior.Color = RGB(217, 217, 217)
    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Sub StyleTitle(ByVal rng As Range)
    rng.Font.Bold = True
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = RGB(68, 84, 106)
End Sub

Private Sub AddListValidation(ByVal rng As Range, ByVal csv As String)
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=csv
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

'======================================================================
' METRICS
'======================================================================

Public Sub RefreshWeeklyMetrics()

    Dim wsIn As Worksheet, wsRel As Worksheet
    Dim src As Variant
    Dim i As Long, lr As Long
    Dim d1 As Date, d2 As Date
    Dim rel As Double, bg As Double, rwk As Double, rej As Double
    Dim processed As Double, plan As Double

    On Error GoTo Fail
    Set wsIn = RequireSheet(SH_WEEKLYINPUT)

    d1 = NzDate(wsIn.Cells(R_WEEKSTART, 3).Value, Date - 7)
    d2 = NzDate(wsIn.Cells(R_WEEKEND, 3).Value, Date)

    ' ---- Release Log ------------------------------------------------
    Set wsRel = GetSheet(SH_RELEASELOG)
    If wsRel Is Nothing Then
        Err.Raise vbObjectError + 521, , SH_RELEASELOG & " is missing - cannot compute metrics."
    End If

    lr = LastRow(wsRel)
    If lr >= 2 Then
        src = wsRel.Range("A2:H" & lr).Value
        For i = 1 To UBound(src, 1)
            If IsDate(src(i, 1)) Then
                If CDate(src(i, 1)) >= d1 And CDate(src(i, 1)) <= d2 Then
                    rel = rel + Nz(src(i, 4))      ' D - net released, already excludes F/G/H
                    rwk = rwk + Nz(src(i, 6))      ' F
                    bg = bg + Nz(src(i, 7))        ' G
                    rej = rej + Nz(src(i, 8))      ' H
                End If
            End If
        Next i
        Erase src
    End If

    processed = rel + rwk + bg + rej

    wsIn.Cells(R_PROCESSED, 3).Value = processed
    wsIn.Cells(R_RELEASED, 3).Value = rel
    wsIn.Cells(R_BGRADE, 3).Value = bg
    wsIn.Cells(R_REWORK, 3).Value = rwk
    wsIn.Cells(R_REJECT, 3).Value = rej
    wsIn.Cells(R_REJRATE, 3).Value = IIf(processed = 0, 0, rej / processed)

    plan = Nz(wsIn.Cells(R_PLAN, 3).Value)
    wsIn.Cells(R_VARIANCE, 3).Value = rel - plan
    wsIn.Cells(R_TOPDEFECT, 3).Value = TopDefect(d1, d2)

    LogLine "RefreshWeeklyMetrics " & Format$(d1, "yyyy-mm-dd") & ".." & Format$(d2, "yyyy-mm-dd") & _
            ": processed " & processed & ", released " & rel & ", reject " & rej

    Say "Metrics refreshed for " & Format$(d1, "m/d") & " - " & Format$(d2, "m/d") & "." & vbCrLf & vbCrLf & _
        "Processed: " & Format$(processed, "#,##0") & vbCrLf & _
        "Released:  " & Format$(rel, "#,##0") & vbCrLf & _
        "Reject rate: " & Format$(IIf(processed = 0, 0, rej / processed), "0.00%"), _
        "REVO Weekly Update"
    Exit Sub
Fail:
    LogLine "*** RefreshWeeklyMetrics failed: " & Err.Description
    Say "Refresh failed: " & Err.Description, "REVO Weekly Update", vbCritical
End Sub

Private Function TopDefect(ByVal d1 As Date, ByVal d2 As Date) As String

    Dim ws As Worksheet, src As Variant, dic As Object
    Dim i As Long, lr As Long
    Dim best As String, bestQty As Double
    Dim k As Variant, code As String

    Set ws = GetSheet(SH_QLOG)
    If ws Is Nothing Then Exit Function
    lr = LastRow(ws)
    If lr < 2 Then Exit Function

    Set dic = CreateObject("Scripting.Dictionary")
    src = ws.Range("A2:J" & lr).Value
    For i = 1 To UBound(src, 1)
        If IsDate(src(i, 1)) Then
            If CDate(src(i, 1)) >= d1 And CDate(src(i, 1)) <= d2 Then
                code = Trim$(CStr(src(i, 10) & ""))
                If Len(code) > 0 Then
                    If dic.Exists(code) Then
                        dic(code) = dic(code) + Nz(src(i, 6))
                    Else
                        dic.Add code, Nz(src(i, 6))
                    End If
                End If
            End If
        End If
    Next i

    For Each k In dic.keys
        If dic(k) > bestQty Then
            bestQty = dic(k)
            best = CStr(k)
        End If
    Next k

    If Len(best) > 0 Then TopDefect = best & " (" & Format$(bestQty, "#,##0") & ")"

End Function

'======================================================================
' EMAIL
'
' This routine calls .Display and nothing else. There is no .Send
' anywhere in this module and there must never be one - Osvaldo reviews
' every update before it leaves.
'======================================================================

Public Sub BuildWeeklyUpdate()

    Dim wsIn As Worksheet
    Dim ol As Object, mail As Object
    Dim html As String, sendTo As String
    Dim d1 As Date, d2 As Date

    On Error GoTo Fail
    Set wsIn = RequireSheet(SH_WEEKLYINPUT)

    d1 = NzDate(wsIn.Cells(R_WEEKSTART, 3).Value, Date - 7)
    d2 = NzDate(wsIn.Cells(R_WEEKEND, 3).Value, Date)
    sendTo = Trim$(CStr(wsIn.Cells(R_TO, 3).Value))

    If InStr(1, sendTo, "REPLACE", vbTextCompare) > 0 Or Len(sendTo) = 0 Then
        Say "The To line on " & SH_WEEKLYINPUT & " cell C" & R_TO & " is still a placeholder." & vbCrLf & vbCrLf & _
            "Add the nine recipients before building the update.", "REVO Weekly Update", vbExclamation
        Exit Sub
    End If

    If Nz(wsIn.Cells(R_RELEASED, 3).Value) = 0 Then
        If Not Confirm("Released shafts is zero. Build the email anyway?" & vbCrLf & vbCrLf & _
                       "No = stop, run RefreshWeeklyMetrics first.", "REVO Weekly Update") Then Exit Sub
    End If

    html = BuildHtml(wsIn)

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo Fail
    If ol Is Nothing Then Err.Raise vbObjectError + 520, , "Outlook is not available on this machine."

    Set mail = ol.CreateItem(0)
    With mail
        .Subject = "REVO Production Update (" & Format$(d1, "m/d") & " - " & Format$(d2, "m/d") & ")"
        .To = sendTo
        .CC = CStr(wsIn.Cells(R_CC, 3).Value)
        .HTMLBody = html
        .Display                      ' review, then send by hand
    End With

    LogLine "BuildWeeklyUpdate: draft opened for " & Format$(d1, "m/d") & "-" & Format$(d2, "m/d")
    Exit Sub
Fail:
    LogLine "*** BuildWeeklyUpdate failed: " & Err.Description
    Say "Build failed: " & Err.Description, "REVO Weekly Update", vbCritical
End Sub

Private Function BuildHtml(ByVal wsIn As Worksheet) As String

    Dim s As String
    Dim sections As Variant
    Dim addr As Variant
    Dim i As Long

    sections = Split(SECTION_ORDER, "|")

    s = "<div style=""font-family:Calibri,sans-serif;font-size:11pt;color:#1f2328"">"
    s = s & "<p>Hi Team,</p>"
    s = s & "<p>Below is this week's REVO Operations Update.</p>"

    For i = LBound(sections) To UBound(sections)
        s = s & SectionHtml(wsIn, CStr(sections(i)), (CStr(sections(i)) = "Summary"))
    Next i

    s = s & "<p>Please let me know if you have any questions or feedback.</p>"
    s = s & "<p>Regards,</p>"
    s = s & "<p style=""margin-bottom:0"">" & HtmlEsc(CStr(wsIn.Cells(R_FROM, 3).Value)) & _
            " | <i>" & HtmlEsc(CStr(wsIn.Cells(R_SIGTITLE, 3).Value)) & "</i><br>"
    addr = Split(CStr(wsIn.Cells(R_SIGADDR, 3).Value), "|")
    For i = LBound(addr) To UBound(addr)
        s = s & HtmlEsc(CStr(addr(i))) & "<br>"
    Next i
    s = s & HtmlEsc(CStr(wsIn.Cells(R_SIGPHONE, 3).Value)) & "</p>"
    s = s & "</div>"

    BuildHtml = s

End Function

' Returns one section. Summary renders as paragraphs, everything else as bullets.
Private Function SectionHtml(ByVal wsIn As Worksheet, ByVal sectionName As String, _
                             ByVal asParagraphs As Boolean) As String

    Dim src As Variant
    Dim i As Long, lr As Long, n As Long
    Dim body As String, txt As String

    lr = wsIn.Cells(wsIn.Rows.Count, 2).End(xlUp).Row
    If lr < R_NARRATIVE Then Exit Function

    src = wsIn.Range("B" & R_NARRATIVE & ":D" & lr).Value

    For i = 1 To UBound(src, 1)
        If StrComp(Trim$(CStr(src(i, 1) & "")), sectionName, vbTextCompare) = 0 Then
            If src(i, 3) <> False Then
                txt = Trim$(CStr(src(i, 2) & ""))
                If Len(txt) > 0 Then
                    txt = ApplyTokens(wsIn, txt)
                    n = n + 1
                    If asParagraphs Then
                        body = body & "<p>" & HtmlEsc(txt) & "</p>"
                    Else
                        body = body & "<li style=""margin-bottom:4px"">" & HtmlEsc(txt) & "</li>"
                    End If
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Function      ' empty section is skipped entirely

    SectionHtml = "<p style=""margin-bottom:2px""><b>" & sectionName & "</b></p>"
    If asParagraphs Then
        SectionHtml = SectionHtml & body
    Else
        SectionHtml = SectionHtml & "<ul style=""margin-top:4px"">" & body & "</ul>"
    End If

End Function

' Replaces {TOKENS} with the calculated values.
' Renamed from Substitute - that is a worksheet function name and
' shadowing it makes the module confusing to read.
Private Function ApplyTokens(ByVal wsIn As Worksheet, ByVal txt As String) As String

    Dim s As String
    Dim rel As Double, rej As Double

    rel = Nz(wsIn.Cells(R_RELEASED, 3).Value)
    rej = Nz(wsIn.Cells(R_REJECT, 3).Value)
    s = txt

    s = Replace(s, "{PROCESSED}", Format$(Nz(wsIn.Cells(R_PROCESSED, 3).Value), "#,##0"))
    s = Replace(s, "{RELEASED}", Format$(rel, "#,##0"))
    s = Replace(s, "{AGRADE}", Format$(rel, "#,##0"))
    s = Replace(s, "{BGRADE}", Format$(Nz(wsIn.Cells(R_BGRADE, 3).Value), "#,##0"))
    s = Replace(s, "{REWORK}", Format$(Nz(wsIn.Cells(R_REWORK, 3).Value), "#,##0"))
    s = Replace(s, "{REJECT}", Format$(rej, "#,##0"))
    s = Replace(s, "{REJECTRATE}", Format$(Nz(wsIn.Cells(R_REJRATE, 3).Value), "0.00%"))
    s = Replace(s, "{REJECTRATE_REL}", Format$(IIf(rel = 0, 0, rej / rel), "0.00%"))
    s = Replace(s, "{PLAN}", Format$(Nz(wsIn.Cells(R_PLAN, 3).Value), "#,##0"))
    s = Replace(s, "{VARIANCE}", Format$(Nz(wsIn.Cells(R_VARIANCE, 3).Value), "#,##0;-#,##0"))
    s = Replace(s, "{WEEKS}", CStr(wsIn.Cells(R_WEEKS, 3).Value))
    s = Replace(s, "{TOPDEFECT}", CStr(wsIn.Cells(R_TOPDEFECT, 3).Value))
    s = Replace(s, "{WEEKSTART}", Format$(NzDate(wsIn.Cells(R_WEEKSTART, 3).Value, Date), "m/d"))
    s = Replace(s, "{WEEKEND}", Format$(NzDate(wsIn.Cells(R_WEEKEND, 3).Value, Date), "m/d"))

    ApplyTokens = s

End Function

Private Function HtmlEsc(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    HtmlEsc = t
End Function
