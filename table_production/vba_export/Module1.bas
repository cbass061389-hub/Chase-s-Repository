Public Sub SetupRejectDashboardNames()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet   ' or specify sheet name

    '====================
    ' WTD (Row 6 values)
    '====================
    AddName "WTD_Released", ws.Range("B6")
    AddName "WTD_Reject", ws.Range("E6")
    AddName "WTD_RejectPct", ws.Range("H6")
    AddName "WTD_Rework", ws.Range("K6")
    AddName "WTD_ReworkPct", ws.Range("N6")
    AddName "WTD_BGrade", ws.Range("Q6")

    '====================
    ' MTD (Row 11 values)
    '====================
    AddName "MTD_Released", ws.Range("B11")
    AddName "MTD_Reject", ws.Range("E11")
    AddName "MTD_RejectPct", ws.Range("H11")
    AddName "MTD_Rework", ws.Range("K11")
    AddName "MTD_ReworkPct", ws.Range("N11")
    AddName "MTD_BGrade", ws.Range("Q11")

    '====================
    ' YTD (Row 16 values)
    '====================
    AddName "YTD_Released", ws.Range("B16")
    AddName "YTD_Reject", ws.Range("E16")
    AddName "YTD_RejectPct", ws.Range("H16")
    AddName "YTD_Rework", ws.Range("K16")
    AddName "YTD_ReworkPct", ws.Range("N16")
    AddName "YTD_BGrade", ws.Range("Q16")
    
    '----------------------------
' Lane | Date Range labels
'----------------------------
AddName "WTD_RangeLabel", ws.Range("B6")
AddName "MTD_RangeLabel", ws.Range("B13")
AddName "YTD_RangeLabel", ws.Range("B18")

    MsgBox "Dashboard Names Created Successfully"

End Sub


Private Sub AddName(ByVal nm As String, ByVal rng As Range)

    Dim n As name
    
    ' Delete workbook-level names
    For Each n In ThisWorkbook.names
        If n.name = nm Or n.name Like "*!" & nm Then
            n.Delete
        End If
    Next n

    ' Delete sheet-level names
    For Each n In rng.Worksheet.names
        If n.name Like "*!" & nm Then
            n.Delete
        End If
    Next n

    ' Add fresh workbook-level name
    ThisWorkbook.names.Add name:=nm, refersTo:=rng

End Sub
