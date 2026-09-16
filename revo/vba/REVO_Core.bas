Attribute VB_Name = "REVO_Core"
'==============================================================================
' REVO_Core  -  shared foundation for the REVO production system
'
' Single source of truth for: sheet names, SKU attribute parsing, safe type
' conversion, header discovery, application state, and the statistics used by
' the cart-velocity engine.
'
' Everything else in the REVO_* family depends on this module. Import it first.
'
' DESIGN NOTE - header discovery
'   Nothing in this system hardcodes a data row. Sheets grow title blocks over
'   time (REVO Floor already has eight rows above its header). Every reader
'   locates its header row by looking for a known column caption, then works
'   relative to that. A layout change moves the data; it no longer breaks the
'   code.
'==============================================================================
Option Explicit

'---------------------------------------------------------------- sheet names
Public Const SH_FLOOR       As String = "REVO Floor"
Public Const SH_RELEASE_LOG As String = "Release Log"
Public Const SH_REJECT_LOG  As String = "Reject Log"
Public Const SH_REWORK      As String = "Rework Tracker"
Public Const SH_SHIPMENTS   As String = "Shipments"
Public Const SH_FLOOR_LOG   As String = "Floor Log"
Public Const SH_HOLDS       As String = "Holds"
Public Const SH_CART_TRK    As String = "Cart Tracker"
Public Const SH_PLAN_CFG    As String = "Plan Config"
Public Const SH_DAILY_PLAN  As String = "Daily Production Plan"
Public Const SH_WEEK_PLAN   As String = "Weekly Production Plan"

' --- sheets this system creates ---
Public Const SH_QUAL_LOG    As String = "Quality Log"
Public Const SH_QUAL_CODES  As String = "Quality_Codes"
Public Const SH_QUAL_DASH   As String = "Quality Dashboard"
Public Const SH_SCORECARD   As String = "REVO Scorecard"
Public Const SH_VELOCITY    As String = "Cart Velocity"
Public Const SH_AUDIT       As String = "REVO Audit"
Public Const SH_SELFTEST    As String = "REVO Self Test"

'------------------------------------------------------------------- defaults
Public Const DEFAULT_SHAFTS_PER_CART As Long = 98
Public Const DEFAULT_OPS_IN_ROUTING  As Long = 15

'==============================================================================
' SKU ATTRIBUTES
'
' Observed SKU grammar across Release Log, Reject Log and REVO Floor:
'   S PRE REVO 12.4 RAD 30 BVP      S PRE REVO 11.8 10T WVP
'   S PRE REVO 12.4 UNI BVP         S PRE REVO BK 30 BVP
'   S PRE REVO BK BVP               S PRE REVO AIR
'   S CRM REVO 3C ULD UNI - CAT2
'
' Bucket follows the convention already established by the legacy
' GetBelorFromSKU: AIR and CRM override the numeric groups, and BK rolls into
' the 12.9 family. That convention is preserved so historical reporting stays
' comparable.
'==============================================================================
Public Type SkuAttrs
    Raw       As String
    Diameter  As String   ' 11.8 | 12.4 | 12.9 | BK | AIR | CRM | UNKNOWN
    Joint     As String   ' UNI | RAD | RAD30 | 10T | ULD | 30 | STD
    VP        As String   ' BVP | WVP | NA
    Lane      As String   ' BVP | WVP  (legacy lane, AIR/CRM count as WVP)
    Family    As String   ' "<Diameter> <Joint>" - the cart-velocity grouping key
    Belor     As String   ' legacy GetBelorFromSKU bucket, kept for continuity
    IsValid   As Boolean
End Type

Public Function ParseSku(ByVal sku As String) As SkuAttrs
    Dim a As SkuAttrs
    Dim s As String

    a.Raw = Trim$(CStr(sku))
    s = " " & UCase$(a.Raw) & " "
    s = Replace$(s, "-", " ")
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop

    If Len(Trim$(s)) = 0 Then
        a.Diameter = "UNKNOWN": a.Joint = "STD": a.VP = "NA"
        a.Lane = "": a.Family = "UNKNOWN": a.Belor = "": a.IsValid = False
        ParseSku = a
        Exit Function
    End If

    '--- diameter / model family -------------------------------------------
    If InStr(s, " AIR ") > 0 Then
        a.Diameter = "AIR"
    ElseIf InStr(s, " CRM ") > 0 Then
        a.Diameter = "CRM"
    ElseIf InStr(s, " 11.8 ") > 0 Then
        a.Diameter = "11.8"
    ElseIf InStr(s, " 12.4 ") > 0 Then
        a.Diameter = "12.4"
    ElseIf InStr(s, " 12.9 ") > 0 Then
        a.Diameter = "12.9"
    ElseIf InStr(s, " BK ") > 0 Then
        a.Diameter = "BK"
    Else
        a.Diameter = "UNKNOWN"
    End If

    '--- joint type ---------------------------------------------------------
    ' RAD 30 and BK 30 are distinct from plain RAD / BK, so test the 30
    ' variants first.
    If InStr(s, " RAD 30 ") > 0 Then
        a.Joint = "RAD30"
    ElseIf InStr(s, " RAD ") > 0 Then
        a.Joint = "RAD"
    ElseIf InStr(s, " 10T ") > 0 Then
        a.Joint = "10T"
    ElseIf InStr(s, " ULD ") > 0 Then
        a.Joint = "ULD"
    ElseIf InStr(s, " UNI ") > 0 Then
        a.Joint = "UNI"
    ElseIf InStr(s, " 30 ") > 0 Then
        a.Joint = "30"
    Else
        a.Joint = "STD"
    End If

    '--- valve / porting ----------------------------------------------------
    If InStr(s, " BVP ") > 0 Then
        a.VP = "BVP"
    ElseIf InStr(s, " WVP ") > 0 Then
        a.VP = "WVP"
    Else
        a.VP = "NA"
    End If

    '--- legacy lane (AIR and CRM report as WVP, matching LaneFromSKU) -------
    If a.VP = "WVP" Or a.Diameter = "AIR" Or a.Diameter = "CRM" Then
        a.Lane = "WVP"
    ElseIf a.VP = "BVP" Then
        a.Lane = "BVP"
    Else
        a.Lane = ""
    End If

    a.Belor = BelorBucket(a)
    a.Family = a.Diameter & " " & a.Joint
    a.IsValid = (a.Diameter <> "UNKNOWN")

    ParseSku = a
End Function

' Legacy bucket, reproduced exactly so historical comparisons hold.
Private Function BelorBucket(ByRef a As SkuAttrs) As String
    Dim is30 As Boolean
    is30 = (a.Joint = "RAD30" Or a.Joint = "30")

    Select Case a.Diameter
        Case "AIR":  BelorBucket = "AIR"
        Case "CRM":  BelorBucket = "CRM"
        Case "11.8": BelorBucket = "11.8"
        Case "12.4": BelorBucket = IIf(is30, "12.4 30", "12.4")
        Case "12.9": BelorBucket = IIf(is30, "12.9 30", "12.9")
        Case "BK":   BelorBucket = IIf(is30, "12.9 30", "12.9")
        Case Else:   BelorBucket = ""
    End Select
End Function

'==============================================================================
' SAFE CONVERSION
'==============================================================================
Public Function SafeD(ByVal v As Variant, Optional ByVal dflt As Double = 0#) As Double
    On Error GoTo Fallback
    If IsEmpty(v) Or IsNull(v) Then SafeD = dflt: Exit Function
    If IsError(v) Then SafeD = dflt: Exit Function
    If VarType(v) = vbBoolean Then SafeD = IIf(CBool(v), 1#, 0#): Exit Function
    If Not IsNumeric(v) Then SafeD = dflt: Exit Function
    SafeD = CDbl(v)
    Exit Function
Fallback:
    SafeD = dflt
End Function

Public Function SafeL(ByVal v As Variant, Optional ByVal dflt As Long = 0) As Long
    Dim d As Double
    d = SafeD(v, CDbl(dflt))
    If d > 2147483647# Then d = 2147483647#
    If d < -2147483648# Then d = -2147483648#
    On Error GoTo Fallback
    SafeL = CLng(d)
    Exit Function
Fallback:
    SafeL = dflt
End Function

Public Function SafeS(ByVal v As Variant, Optional ByVal dflt As String = "") As String
    On Error GoTo Fallback
    If IsEmpty(v) Or IsNull(v) Then SafeS = dflt: Exit Function
    If IsError(v) Then SafeS = dflt: Exit Function
    SafeS = Trim$(CStr(v))
    Exit Function
Fallback:
    SafeS = dflt
End Function

Public Function SafeDate(ByVal v As Variant, Optional ByVal dflt As Date = 0) As Date
    On Error GoTo Fallback
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then SafeDate = dflt: Exit Function
    If IsDate(v) Then SafeDate = CDate(v): Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 1 And CDbl(v) < 80000 Then SafeDate = CDate(CDbl(v)): Exit Function
    End If
    SafeDate = dflt
    Exit Function
Fallback:
    SafeDate = dflt
End Function

' TRUE only for a genuine boolean-true, a non-zero number, or the strings
' TRUE / 1 / X / YES. Status text such as "RELEASED" is deliberately FALSE.
Public Function Boolish(ByVal v As Variant) As Boolean
    Dim s As String
    On Error GoTo Fallback
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Boolish = False: Exit Function
    If VarType(v) = vbBoolean Then Boolish = CBool(v): Exit Function
    If IsNumeric(v) Then Boolish = (CDbl(v) <> 0): Exit Function
    s = UCase$(Trim$(CStr(v)))
    Boolish = (s = "TRUE" Or s = "1" Or s = "X" Or s = "YES" Or s = "Y")
    Exit Function
Fallback:
    Boolish = False
End Function

'==============================================================================
' SHEETS AND HEADERS
'==============================================================================
Public Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Function GetSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Public Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheet(nm)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = nm
    End If
    Set GetOrCreateSheet = ws
End Function

' Locate the header row by searching the first `scanRows` rows for any of the
' supplied captions. Returns 0 when not found - callers must handle that
' rather than assume a row.
Public Function FindHeaderRow(ByVal ws As Worksheet, ByVal captions As Variant, _
                              Optional ByVal scanRows As Long = 25, _
                              Optional ByVal scanCols As Long = 40) As Long
    Dim r As Long, c As Long, i As Long
    Dim cell As String

    For r = 1 To scanRows
        For c = 1 To scanCols
            cell = UCase$(SafeS(ws.Cells(r, c).Value))
            If Len(cell) > 0 Then
                For i = LBound(captions) To UBound(captions)
                    If cell = UCase$(Trim$(CStr(captions(i)))) Then
                        FindHeaderRow = r
                        Exit Function
                    End If
                Next i
            End If
        Next c
    Next r
    FindHeaderRow = 0
End Function

' Find a column by any of its accepted captions on a known header row.
' Returns 0 when absent.
Public Function FindCol(ByVal ws As Worksheet, ByVal headerRow As Long, _
                        ByVal captions As Variant, _
                        Optional ByVal scanCols As Long = 60) As Long
    Dim c As Long, i As Long
    Dim cell As String

    If headerRow <= 0 Then FindCol = 0: Exit Function

    For c = 1 To scanCols
        cell = UCase$(SafeS(ws.Cells(headerRow, c).Value))
        If Len(cell) > 0 Then
            For i = LBound(captions) To UBound(captions)
                If cell = UCase$(Trim$(CStr(captions(i)))) Then
                    FindCol = c
                    Exit Function
                End If
            Next i
        End If
    Next c
    FindCol = 0
End Function

' Find a column, or append it to the right of the used header band if missing.
Public Function FindOrAddCol(ByVal ws As Worksheet, ByVal headerRow As Long, _
                             ByVal caption As String, _
                             Optional ByVal captions As Variant) As Long
    Dim c As Long, lastCol As Long
    Dim lookFor As Variant

    If IsMissing(captions) Then
        lookFor = Array(caption)
    Else
        lookFor = captions
    End If

    c = FindCol(ws, headerRow, lookFor)
    If c > 0 Then FindOrAddCol = c: Exit Function

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1
    c = lastCol + 1
    ws.Cells(headerRow, c).Value = caption
    ws.Cells(headerRow, c).Font.Bold = True
    FindOrAddCol = c
End Function

Public Sub WriteHeaders(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal arr As Variant)
    Dim i As Long, c As Long
    c = 1
    For i = LBound(arr) To UBound(arr)
        ws.Cells(headerRow, c).Value = arr(i)
        c = c + 1
    Next i
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, c - 1))
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

' Last row carrying data in a column, never above the header.
Public Function LastDataRow(ByVal ws As Worksheet, ByVal col As Long, _
                            ByVal headerRow As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).row
    If r < headerRow Then r = headerRow
    LastDataRow = r
End Function

'==============================================================================
' APPLICATION STATE
'
' Nested-safe. Push before a bulk operation, pop in the cleanup handler. A
' nested push does not clobber the outer restore values, and popping more than
' pushed cannot leave events disabled.
'==============================================================================
Private mDepth As Long
Private mScreen As Boolean
Private mEvents As Boolean
Private mCalc As XlCalculation

Public Sub PushAppState(Optional ByVal quietScreen As Boolean = True)
    If mDepth = 0 Then
        mScreen = Application.ScreenUpdating
        mEvents = Application.EnableEvents
        mCalc = Application.Calculation
    End If
    mDepth = mDepth + 1
    If quietScreen Then Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    On Error GoTo 0
End Sub

Public Sub PopAppState()
    If mDepth > 0 Then mDepth = mDepth - 1
    If mDepth = 0 Then
        Application.ScreenUpdating = mScreen
        Application.EnableEvents = mEvents
        On Error Resume Next
        Application.Calculation = mCalc
        On Error GoTo 0
    End If
End Sub

' Force everything back on regardless of depth. For error handlers only.
Public Sub ResetAppState()
    mDepth = 0
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    On Error GoTo 0
End Sub

' A modal UserForm will not paint reliably while ScreenUpdating is off. Wrap
' every .Show with these two.
Public Sub BeginModalDialog()
    Application.ScreenUpdating = True
    DoEvents
End Sub

Public Sub EndModalDialog()
    If mDepth > 0 Then Application.ScreenUpdating = False
End Sub

'==============================================================================
' SHIFT MODEL
'
' Release and FQC happen on days only - the floor cannot release a cart at
' night. The scorecard needs that distinction to tell overtime apart from
' "wait for first shift".
'==============================================================================
Public Const DAY_SHIFT_START  As Double = 6#    ' 06:00
Public Const DAY_SHIFT_END    As Double = 16.5  ' 16:30
Public Const NIGHT_SHIFT_END  As Double = 2.5   ' 02:30 next day

Public Function ShiftFor(ByVal whenTS As Date) As String
    Dim h As Double
    h = Hour(whenTS) + Minute(whenTS) / 60#
    If h >= DAY_SHIFT_START And h < DAY_SHIFT_END Then
        ShiftFor = "DAY"
    Else
        ShiftFor = "NIGHT"
    End If
End Function

Public Function CanReleaseNow(Optional ByVal whenTS As Variant) As Boolean
    Dim t As Date
    If IsMissing(whenTS) Then t = Now Else t = SafeDate(whenTS, Now)
    CanReleaseNow = (ShiftFor(t) = "DAY")
End Function

' Production days between two dates, Sunday-to-Friday (Saturday is dark),
' excluding whole-site blackout rows on Plan Config.
Public Function ProductionDaysBetween(ByVal d1 As Date, ByVal d2 As Date) As Long
    Dim d As Date, n As Long
    If d2 < d1 Then ProductionDaysBetween = 0: Exit Function
    For d = d1 To d2
        If Weekday(d, vbSunday) <> vbSaturday Then
            If Not IsBlackoutDay(d) Then n = n + 1
        End If
    Next d
    ProductionDaysBetween = n
End Function

Public Function IsBlackoutDay(ByVal d As Date, Optional ByVal opName As String = "ALL") As Boolean
    Dim ws As Worksheet, hdr As Long, r As Long, lastR As Long
    Dim cStart As Long, cEnd As Long, cOp As Long
    Dim s As Date, e As Date, o As String

    Set ws = GetSheet(SH_PLAN_CFG)
    If ws Is Nothing Then Exit Function

    hdr = FindHeaderRow(ws, Array("Start Date"), 40, 10)
    If hdr = 0 Then Exit Function

    cStart = FindCol(ws, hdr, Array("Start Date"))
    cEnd = FindCol(ws, hdr, Array("End Date"))
    cOp = FindCol(ws, hdr, Array("Op (or All)", "Op"))
    If cStart = 0 Or cEnd = 0 Then Exit Function

    lastR = LastDataRow(ws, cStart, hdr)
    For r = hdr + 1 To lastR
        s = SafeDate(ws.Cells(r, cStart).Value, 0)
        e = SafeDate(ws.Cells(r, cEnd).Value, 0)
        If s > 0 And e > 0 Then
            If d >= s And d <= e Then
                o = UCase$(SafeS(ws.Cells(r, cOp).Value, "ALL"))
                If o = "ALL" Or Len(o) = 0 Or o = UCase$(opName) Then
                    IsBlackoutDay = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

'==============================================================================
' RELEASE HISTORY
'
' Shafts released between two dates, from Release Log, ignoring recalled rows.
' Lives here because the scorecard and the quality dashboard both need the same
' number and must never disagree about it.
'==============================================================================
Public Function ReleasedBetween(ByVal d1 As Date, ByVal d2 As Date) As Double
    Dim ws As Worksheet, hdr As Long, lastR As Long, r As Long
    Dim cDate As Long, cQty As Long, cRecall As Long
    Dim d As Date, t As Double

    Set ws = GetSheet(SH_RELEASE_LOG)
    If ws Is Nothing Then Exit Function

    hdr = FindHeaderRow(ws, Array("Qty Released"), 10, 15)
    If hdr = 0 Then Exit Function

    cDate = FindCol(ws, hdr, Array("Date"))
    cQty = FindCol(ws, hdr, Array("Qty Released"))
    cRecall = FindCol(ws, hdr, Array("Recalled"))
    If cDate = 0 Or cQty = 0 Then Exit Function

    lastR = LastDataRow(ws, cDate, hdr)
    For r = hdr + 1 To lastR
        d = SafeDate(ws.Cells(r, cDate).Value, 0)
        If d >= d1 And d <= d2 Then
            If cRecall = 0 Or Not Boolish(ws.Cells(r, cRecall).Value) Then
                t = t + SafeD(ws.Cells(r, cQty).Value)
            End If
        End If
    Next r
    ReleasedBetween = t
End Function

'==============================================================================
' STATISTICS  -  used by the cart-velocity engine
'
' Outlier rejection uses the median absolute deviation rather than standard
' deviation. A held cart that sat for nine days is exactly the kind of value
' that inflates a standard deviation enough to hide itself; MAD does not move.
'==============================================================================
Public Function Median(ByRef v() As Double, ByVal n As Long) As Double
    Dim s() As Double, i As Long
    If n <= 0 Then Median = 0: Exit Function
    ReDim s(1 To n)
    For i = 1 To n
        s(i) = v(i)
    Next i
    QuickSortD s, 1, n
    If n Mod 2 = 1 Then
        Median = s((n + 1) \ 2)
    Else
        Median = (s(n \ 2) + s(n \ 2 + 1)) / 2#
    End If
End Function

Public Function MedianAbsDev(ByRef v() As Double, ByVal n As Long, ByVal med As Double) As Double
    Dim d() As Double, i As Long
    If n <= 0 Then MedianAbsDev = 0: Exit Function
    ReDim d(1 To n)
    For i = 1 To n
        d(i) = Abs(v(i) - med)
    Next i
    MedianAbsDev = Median(d, n)
End Function

' Returns the count kept, and overwrites v() in place with the retained values.
' threshold is in robust sigmas (3.0 is the usual choice).
Public Function RejectOutliers(ByRef v() As Double, ByVal n As Long, _
                               Optional ByVal threshold As Double = 3#) As Long
    Dim med As Double, mad As Double, sigma As Double
    Dim keep() As Double, i As Long, k As Long

    If n <= 3 Then RejectOutliers = n: Exit Function

    med = Median(v, n)
    mad = MedianAbsDev(v, n, med)
    sigma = mad * 1.4826          ' MAD -> sigma for a normal distribution

    ' All-identical sample: nothing to reject.
    If sigma <= 0.000001 Then RejectOutliers = n: Exit Function

    ReDim keep(1 To n)
    For i = 1 To n
        If Abs(v(i) - med) <= threshold * sigma Then
            k = k + 1
            keep(k) = v(i)
        End If
    Next i

    If k = 0 Then RejectOutliers = n: Exit Function

    For i = 1 To k
        v(i) = keep(i)
    Next i
    RejectOutliers = k
End Function

Public Function MeanOf(ByRef v() As Double, ByVal n As Long) As Double
    Dim i As Long, t As Double
    If n <= 0 Then Exit Function
    For i = 1 To n
        t = t + v(i)
    Next i
    MeanOf = t / n
End Function

Public Sub QuickSortD(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim p As Double, t As Double
    If lo >= hi Then Exit Sub
    i = lo: j = hi
    p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p
            i = i + 1
        Loop
        Do While a(j) > p
            j = j - 1
        Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortD a, lo, j
    If i < hi Then QuickSortD a, i, hi
End Sub

'==============================================================================
' IDS AND AUDIT
'==============================================================================
Public Function NewEventID(ByVal prefix As String) As String
    Static seq As Long
    seq = seq + 1
    NewEventID = prefix & "-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(seq, "000")
End Function

Public Function CurrentUser() As String
    Dim u As String
    On Error Resume Next
    u = Application.UserName
    If Len(Trim$(u)) = 0 Then u = Environ$("USERNAME")
    On Error GoTo 0
    If Len(Trim$(u)) = 0 Then u = "UNKNOWN"
    CurrentUser = Trim$(u)
End Function

' Append a line to the audit sheet. Every macro that writes to a sheet a human
' also edits should call this, so "who changed this" is answerable.
Public Sub Audit(ByVal source As String, ByVal action As String, _
                 ByVal target As String, ByVal detail As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = GetOrCreateSheet(SH_AUDIT)
    If ws Is Nothing Then Exit Sub

    If Len(SafeS(ws.Cells(1, 1).Value)) = 0 Then
        WriteHeaders ws, 1, Array("Timestamp", "User", "Shift", "Source", "Action", "Target", "Detail")
        ws.Columns("A:G").ColumnWidth = 22
    End If

    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(r, 2).Value = CurrentUser
    ws.Cells(r, 3).Value = ShiftFor(Now)
    ws.Cells(r, 4).Value = source
    ws.Cells(r, 5).Value = action
    ws.Cells(r, 6).Value = target
    ws.Cells(r, 7).Value = detail
    On Error GoTo 0
End Sub
