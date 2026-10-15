Option Explicit

'========================================================
' REVO Weekly Scorecard Email Module
'========================================================

Private Const SH_SUMMARY As String = "Summary"
Private Const EMAIL_FONT As String = "Aptos, Calibri, Arial, sans-serif"

Public Sub SendWeeklyREVOScorecard()

    If Not SheetExists(SH_SUMMARY) Then
        MsgBox "Missing '" & SH_SUMMARY & "' sheet. Email skipped.", vbExclamation
        Exit Sub
    End If

    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ThisWorkbook
    Set ws = GetSheet(SH_SUMMARY, wb)

    Dim wtdPct As Double, mtdPct As Double, ytdPct As Double
    Dim wtdAct As Double, wtdFc As Double
    Dim mtdAct As Double, mtdFc As Double
    Dim ytdAct As Double, ytdFc As Double

    wtdPct = SafeDbl(ws.Range("F31").Value)
    mtdPct = SafeDbl(ws.Range("L31").Value)
    ytdPct = SafeDbl(ws.Range("R31").Value)

    wtdAct = SafeDbl(ws.Range("D31").Value)
    wtdFc = SafeDbl(ws.Range("C31").Value)

    mtdAct = SafeDbl(ws.Range("J31").Value)
    mtdFc = SafeDbl(ws.Range("I31").Value)

    ytdAct = SafeDbl(ws.Range("P31").Value)
    ytdFc = SafeDbl(ws.Range("O31").Value)

    Dim body As String
    body = BuildREVOScorecardHtml( _
                wtdPct, wtdAct, wtdFc, _
                mtdPct, mtdAct, mtdFc, _
                ytdPct, ytdAct, ytdFc)

    Dim attachFile As String
    attachFile = BuildAttachmentCopy(wb, ws)

    Dim OutApp As Object
    Dim OutMail As Object

    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If OutApp Is Nothing Then
        MsgBox "Outlook could not be opened. Email was not created.", vbExclamation
        Exit Sub
    End If

    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = "cfournier@predatorgroup.com"
        .Subject = "REVO Weekly Scorecard - " & Format(Date, "mm.dd.yy")
        .Importance = 2
        .HTMLBody = body

        If Len(Dir$(attachFile)) > 0 Then
            .Attachments.Add attachFile
        End If

        .Display
    End With

End Sub

'========================================================
' HTML BUILDER
'========================================================

Private Function BuildREVOScorecardHtml( _
    ByVal wtdPct As Double, ByVal wtdAct As Double, ByVal wtdFc As Double, _
    ByVal mtdPct As Double, ByVal mtdAct As Double, ByVal mtdFc As Double, _
    ByVal ytdPct As Double, ByVal ytdAct As Double, ByVal ytdFc As Double) As String

    Dim html As String

    html = ""
    html = html & "<html>"
    html = html & "<body style='margin:0;padding:0;background:#ffffff;" & FontCss() & "color:#0f172a;'>"
    html = html & "<a name='top'></a>"

    html = html & "<table width='760' align='left' cellpadding='0' cellspacing='0' style='width:760px;border-collapse:collapse;background:#ffffff;" & FontCss() & "color:#0f172a;'>"

    html = html & "<tr><td style='" & FontCss() & "background:#0b1f3a;color:#ffffff;padding:18px 22px;'>"
    html = html & "<div style='" & FontCss() & "font-size:22pt;font-weight:700;letter-spacing:-0.3px;'>REVO Weekly Scorecard</div>"
    html = html & "<div style='" & FontCss() & "font-size:10pt;color:#d9e2ec;margin-top:6px;'>Generated " & Format(Date, "mmmm d, yyyy") & "</div>"
    html = html & "</td></tr>"

    html = html & "<tr><td style='" & FontCss() & "padding:14px 0 18px 0;'>"
    html = html & BuildScorecardNavClean(wtdPct)
    html = html & "</td></tr>"

    html = html & "<tr><td style='" & FontCss() & "padding:0 0 16px 0;'>"
    html = html & BuildThroughputScorecard(wtdPct, wtdAct, wtdFc, mtdPct, mtdAct, mtdFc, ytdPct, ytdAct, ytdFc)
    html = html & BuildFillableScorecardSection("safety", "Safety")
    html = html & BuildFillableScorecardSection("quality", "Quality")
    html = html & BuildFillableScorecardSection("forecast", "Forecast")
    html = html & BuildFillableScorecardSection("hr", "HR")
    html = html & BuildFillableScorecardSection("machinery", "Machinery")
    html = html & BuildFillableScorecardSection("supply", "Supply")
    html = html & "</td></tr>"

    html = html & "</table>"
    html = html & "</body></html>"

    BuildREVOScorecardHtml = html

End Function

Private Function BuildScorecardNavClean(ByVal throughputPct As Double) As String

    Dim html As String

    html = ""
    html = html & "<table width='100%' cellpadding='0' cellspacing='0' style='" & FontCss() & "border-collapse:collapse;'>"
    html = html & "<tr>"
    html = html & BuildNavCellClean("Throughput", "throughput", GetThroughputColor(throughputPct))
    html = html & BuildNavCellClean("Safety", "safety", "#00B050")
    html = html & BuildNavCellClean("Quality", "quality", "#00B050")
    html = html & BuildNavCellClean("Forecast", "forecast", "#00B050")
    html = html & BuildNavCellClean("HR", "hr", "#00B050")
    html = html & BuildNavCellClean("Machinery", "machinery", "#00B050")
    html = html & BuildNavCellClean("Supply", "supply", "#00B050")
    html = html & "</tr>"
    html = html & "</table>"

    BuildScorecardNavClean = html

End Function

Private Function BuildNavCellClean(ByVal label As String, _
                                   ByVal anchorName As String, _
                                   ByVal colorHex As String) As String

    BuildNavCellClean = _
        "<td style='" & FontCss() & "border-right:1px solid #d9dee5;text-align:center;padding:10px 6px;width:14.28%;'>" & _
            "<a href='#" & anchorName & "' style='" & FontCss() & "color:#003087;text-decoration:underline;font-size:10.5pt;font-weight:700;'>" & label & "</a><br>" & _
            "<span style='" & FontCss() & "font-size:17pt;line-height:24pt;color:" & colorHex & ";'>&#9679;</span>" & _
        "</td>"

End Function

Private Function BuildThroughputScorecard( _
    ByVal wtdPct As Double, ByVal wtdAct As Double, ByVal wtdFc As Double, _
    ByVal mtdPct As Double, ByVal mtdAct As Double, ByVal mtdFc As Double, _
    ByVal ytdPct As Double, ByVal ytdAct As Double, ByVal ytdFc As Double) As String

    Dim html As String

    html = ""
    html = html & BuildPanelHeader("throughput", "Throughput")

    html = html & "<table width='100%' cellpadding='0' cellspacing='0' style='" & FontCss() & "border-collapse:collapse;border-left:1px solid #d0d7de;border-right:1px solid #d0d7de;border-bottom:1px solid #d0d7de;margin-bottom:16px;'>"

    html = html & "<tr>"
    html = html & "<th style='" & FontCss() & "border:1px solid #d0d7de;background:#ffffff;padding:9px;text-align:left;font-size:10.5pt;font-weight:700;'>Metric</th>"
    html = html & "<th style='" & FontCss() & "border:1px solid #d0d7de;background:#ffffff;padding:9px;text-align:center;font-size:10.5pt;font-weight:700;'>Actual</th>"
    html = html & "<th style='" & FontCss() & "border:1px solid #d0d7de;background:#ffffff;padding:9px;text-align:center;font-size:10.5pt;font-weight:700;'>Plan / Forecast</th>"
    html = html & "<th style='" & FontCss() & "border:1px solid #d0d7de;background:#ffffff;padding:9px;text-align:center;font-size:10.5pt;font-weight:700;'>Variance</th>"
    html = html & "</tr>"

    html = html & BuildProductionRow("WTD Production", wtdAct, wtdFc, wtdPct)
    html = html & BuildProductionRow("MTD Production", mtdAct, mtdFc, mtdPct)
    html = html & BuildProductionRow("YTD Production", ytdAct, ytdFc, ytdPct)

    html = html & BuildSimpleThroughputRow("Next Week Planned Qty", "", "Enter Qty", "")
    html = html & BuildSimpleThroughputRow("Production Risk Level", "<span style='color:#00B050;font-weight:700;'>Low</span>", "", "")

    html = html & "</table>"

    BuildThroughputScorecard = html

End Function

Private Function BuildProductionRow(ByVal metricName As String, _
                                    ByVal actualQty As Double, _
                                    ByVal forecastQty As Double, _
                                    ByVal pct As Double) As String

    Dim colorHex As String
    Dim varianceQty As Double
    Dim varianceText As String

    colorHex = GetThroughputColor(pct)
    varianceQty = actualQty - forecastQty
    varianceText = Format(varianceQty, "#,##0") & " (" & Format(pct, "0.00%") & ")"

    BuildProductionRow = _
        "<tr>" & _
            "<td style='" & FontCss() & "border:1px solid #d0d7de;padding:10px;font-size:10.5pt;'>" & metricName & "</td>" & _
            "<td style='" & FontCss() & "border:1px solid #d0d7de;padding:10px;text-align:center;font-size:10.5pt;'>" & Format(actualQty, "#,##0") & "</td>" & _
            "<td style='" & FontCss() & "border:1px solid #d0d7de;padding:10px;text-align:center;font-size:10.5pt;'>" & Format(forecastQty, "#,##0") & "</td>" & _
            "<td style='" & FontCss() & "border:1px solid #d0d7de;padding:10px;text-align:center;color:" & colorHex & ";font-weight:700;font-size:10.5pt;'>" & varianceText & "</td>" & _
        "</tr>"

End Function

Private Function BuildSimpleThroughputRow(ByVal metricName As String, _
                                          ByVal actualText As String, _
                                          ByVal forecastText As String, _
                                          ByVal varianceText As String) As String

    BuildSimpleThroughputRow = _
        "<tr>" & _
            "<td style='" & FontCss() & "border:1px solid #d0d7de;padding:10px;font-size:10.5pt;'>" & metricName & "</td>" & _
            "<td style='" & FontCss() & "border:1px solid #d0d7de;padding:10px;text-align:center;font-size:10.5pt;'>" & actualText & "</td>" & _
            "<td style='" & FontCss() & "border:1px solid #d0d7de;padding:10px;text-align:center;font-size:10.5pt;'>" & forecastText & "</td>" & _
            "<td style='" & FontCss() & "border:1px solid #d0d7de;padding:10px;text-align:center;font-size:10.5pt;'>" & varianceText & "</td>" & _
        "</tr>"

End Function

Private Function BuildFillableScorecardSection(ByVal anchorName As String, _
                                               ByVal sectionTitle As String) As String

    BuildFillableScorecardSection = _
        BuildPanelHeader(anchorName, sectionTitle) & _
        "<table width='100%' cellpadding='0' cellspacing='0' style='" & FontCss() & "border-collapse:collapse;border-left:1px solid #d0d7de;border-right:1px solid #d0d7de;border-bottom:1px solid #d0d7de;margin-bottom:16px;'>" & _
            "<tr>" & _
                "<td style='" & FontCss() & "padding:13px 14px;height:42px;color:#64748b;font-size:10.5pt;background:#ffffff;'>Enter update here...</td>" & _
            "</tr>" & _
        "</table>"

End Function

Private Function BuildPanelHeader(ByVal anchorName As String, _
                                  ByVal sectionTitle As String) As String

    BuildPanelHeader = _
        "<a name='" & anchorName & "'></a>" & _
        "<table width='100%' cellpadding='0' cellspacing='0' style='" & FontCss() & "border-collapse:collapse;margin-top:10px;'>" & _
            "<tr>" & _
                "<td style='" & FontCss() & "background:#eef3f8;border:1px solid #d0d7de;border-bottom:none;padding:10px 12px;font-size:13pt;font-weight:700;color:#0f172a;'>" & sectionTitle & "</td>" & _
                "<td style='" & FontCss() & "background:#eef3f8;border:1px solid #d0d7de;border-bottom:none;border-left:none;padding:10px 12px;text-align:right;font-size:10pt;'>" & _
                    "<a href='#top' style='" & FontCss() & "color:#003087;text-decoration:underline;'>Back to Top</a>" & _
                "</td>" & _
            "</tr>" & _
        "</table>"

End Function

'========================================================
' COLORS / FORMATTING
'========================================================

Private Function FontCss() As String
    FontCss = "font-family:" & EMAIL_FONT & ";"
End Function

Private Function GetThroughputColor(ByVal pct As Double) As String

    If pct >= 0 Then
        GetThroughputColor = "#00B050"
    ElseIf pct >= -0.05 Then
        GetThroughputColor = "#FFC000"
    ElseIf pct >= -0.1 Then
        GetThroughputColor = "#C00000"
    Else
        GetThroughputColor = "#7030A0"
    End If

End Function

'========================================================
' ATTACHMENT COPY
'========================================================

Private Function BuildAttachmentCopy(ByVal wb As Workbook, ByVal ws As Worksheet) As String

    Dim weekLabel As String
    weekLabel = NzStr(ws.Range("C1").Value)

    If Len(weekLabel) = 0 Then weekLabel = Format(Date, "yyyy-mm-dd")

    weekLabel = Replace$(weekLabel, "/", "-")
    weekLabel = Replace$(weekLabel, ":", "-")
    weekLabel = Replace$(weekLabel, "\", "-")

    Dim ext As String
    Dim p As Long

    p = InStrRev(wb.name, ".")

    If p > 0 Then
        ext = Mid$(wb.name, p)
    Else
        ext = ".xlsm"
    End If

    Dim basePath As String
    basePath = wb.path

    If Len(basePath) = 0 Or InStr(1, basePath, "http", vbTextCompare) > 0 Then
        basePath = Application.DefaultFilePath

        If Len(Dir$(basePath, vbDirectory)) = 0 Then
            basePath = Environ$("USERPROFILE") & "\Documents"
        End If
    End If

    BuildAttachmentCopy = basePath & "\REVO Scorecard for week " & weekLabel & ext

    On Error Resume Next
    wb.SaveCopyAs BuildAttachmentCopy
    On Error GoTo 0

End Function

'========================================================
' GENERAL HELPERS
'========================================================

Private Function SheetExists(ByVal sheetName As String, _
                             Optional ByVal wb As Workbook) As Boolean

    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing

End Function

Private Function GetSheet(ByVal sheetName As String, _
                          Optional ByVal wb As Workbook) As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook

    On Error Resume Next
    Set GetSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetSheet Is Nothing Then
        Err.Raise vbObjectError + 513, "GetSheet", _
                  "Worksheet '" & sheetName & "' was not found."
    End If

End Function

Private Function SafeDbl(ByVal v As Variant) As Double

    If IsError(v) Then
        SafeDbl = 0#
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeDbl = 0#
    ElseIf Trim$(CStr(v)) = "" Then
        SafeDbl = 0#
    ElseIf Not IsNumeric(v) Then
        SafeDbl = 0#
    Else
        SafeDbl = CDbl(v)
    End If

End Function

Private Function NzStr(ByVal v As Variant, Optional ByVal fallback As String = "") As String

    If IsError(v) Then
        NzStr = fallback
    ElseIf IsEmpty(v) Or IsNull(v) Then
        NzStr = fallback
    Else
        NzStr = Trim$(CStr(v))
    End If

End Function
