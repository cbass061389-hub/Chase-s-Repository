Option Explicit



Private Sub CommitScheduleToValues()
    Dim wsPE As Worksheet: Set wsPE = ThisWorkbook.Sheets("Plan Engine")
    Dim peCols As Variant, scCols As Variant, i As Long
    peCols = Array("AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD")
    scCols = Array("J", "N", "R", "V", "Z", "AD", "AH", "AL", "AP", "AT", "AX", "BB", "BF")
    Application.Calculation = xlCalculationAutomatic
    For i = LBound(peCols) To UBound(peCols)
        wsPE.Range(peCols(i) & "2:" & peCols(i) & "121").Formula = _
          "=IFERROR(INDEX(Scheduler!" & scCols(i) & "$6:" & scCols(i) & _
          "$125,MATCH(ROW(),Scheduler!$B$6:$B$125,0)),"""")"
    Next i
    Application.CalculateFull
    For i = LBound(peCols) To UBound(peCols)
        wsPE.Range(peCols(i) & "2:" & peCols(i) & "121").Value = _
            wsPE.Range(peCols(i) & "2:" & peCols(i) & "121").Value
    Next i
End Sub

Private Sub ClearCarriedForward()
    With ThisWorkbook.Sheets("Daily Production Plan").Range("B23:E200")
        .ClearContents: .Interior.ColorIndex = xlNone: .Font.Italic = False
    End With
End Sub

Public Sub RebuildSchedule()
    If MsgBox("Rebuild the entire schedule from today? This replaces all committed dates.", _
        vbYesNo + vbQuestion, "Rebuild Schedule") <> vbYes Then Exit Sub
    ThisWorkbook.Sheets("Plan Config").Range("M6").Value = Date
    CommitScheduleToValues
    ClearCarriedForward
    ApplyHolds                       ' <-- freeze held carts at their op
    MsgBox "Schedule rebuilt for " & Format(Date, "m/d/yyyy") & ".", vbInformation
End Sub



Sub ApplyHolds()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Plan Engine")
    Application.Calculate                      ' refresh HoldOp helper

    Dim firstOpCol As Long, lastOpCol As Long, holdOpCol As Long
    firstOpCol = ws.Range("AR1").Column        ' IC
    lastOpCol = ws.Range("BD1").Column         ' 160
    holdOpCol = ws.Range("BM1").Column         ' HoldOp helper

    Dim r As Long, c As Long, heldCol As Long
    Dim holdOp As Variant

    For r = 2 To 121
        holdOp = ws.Cells(r, holdOpCol).Value
        If Len(CStr(holdOp)) > 0 Then
            heldCol = 0
            For c = firstOpCol To lastOpCol
                If CStr(ws.Cells(1, c).Value) = CStr(holdOp) Then heldCol = c: Exit For
            Next c
            ' keep the held op, clear every op scheduled after it
            If heldCol > 0 And heldCol < lastOpCol Then
                ws.Range(ws.Cells(r, heldCol + 1), ws.Cells(r, lastOpCol)).ClearContents
            End If
        End If
    Next r
End Sub
