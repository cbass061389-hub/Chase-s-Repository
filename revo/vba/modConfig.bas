Attribute VB_Name = "modConfig"
Option Explicit

'======================================================================
' modConfig  -  single point of change for the whole workbook
'
' Every sheet name lives here. Nothing else hard-codes a sheet name.
' Rename a tab, change one line, done.
'
' GetSheet() is the guard: it returns Nothing instead of raising 1004
' when a sheet has been removed. Callers decide how to degrade. This is
' what lets the inventory/BOM tabs come out without every code path
' that touches them dying.
'
' Also owns:
'   SilentMode  - suppresses every MsgBox so the PowerShell driver can
'                 run the whole migration unattended. A MsgBox in an
'                 automated run hangs Excel forever with no window to
'                 click, and the driver times out holding a file lock.
'   FastOn/Off  - saves and RESTORES calculation state. The original
'                 modules hard-set xlCalculationAutomatic on exit, which
'                 forced a full recalc between every migration phase.
'   LogLine     - appends to REVO_Migration_Log.txt beside the workbook.
'                 The driver reads this; it is the audit trail.
'======================================================================

Public Const APP_VERSION As String = "2.1"

' --- Floor / planning -------------------------------------------------
Public Const SH_FLOOR       As String = "REVO Floor"
Public Const SH_PLANCONFIG  As String = "Plan Config"
Public Const SH_PLANENGINE  As String = "Plan Engine"
Public Const SH_SCHEDULER   As String = "Scheduler"
Public Const SH_HOLDS       As String = "Holds"
Public Const SH_OVERRIDES   As String = "Overrides"
Public Const SH_COMMAND     As String = "Command Center"

' --- Release / disposition -------------------------------------------
Public Const SH_RELEASELOG  As String = "Release Log"
Public Const SH_REJECTLOG   As String = "Reject Log"
Public Const SH_REWORK      As String = "Rework Tracker"
Public Const SH_REWORKINV   As String = "Rework From Inventory"

' --- Reporting --------------------------------------------------------
Public Const SH_WEEKLYRPT   As String = "Weekly Reporting"
Public Const SH_PRODANALYSIS As String = "Production Analysis"
Public Const SH_REJECTDASH  As String = "Reject_Dashboard"
Public Const SH_WEEKLYINPUT As String = "Weekly_Update_Input"

' --- Quality layer (new) ---------------------------------------------
Public Const SH_DEFECTS     As String = "Defect_Master"
Public Const SH_QLOG        As String = "Reject_Log_V2"
Public Const SH_QANALYSIS   As String = "Quality_Analysis"
Public Const SH_RCA         As String = "RCA_Register"

' --- Shared constants -------------------------------------------------
Public Const OPS_CSV        As String = "IC,ASSYM,10,20,30,100,40,110,50,60,70,120,130,160,FQC"
Public Const DISPOSITIONS   As String = "Reject,Rework,B-Grade"
Public Const SHIFTS         As String = "Day,Night"
Public Const VALIDATION_ROWS As Long = 5000

Public Const LOG_FILE       As String = "REVO_Migration_Log.txt"

' --- Runtime state ----------------------------------------------------
Public SilentMode As Boolean            ' set True by the driver

Private mDepth As Long
Private mCalc As XlCalculation
Private mScreen As Boolean
Private mEvents As Boolean
Private mAlerts As Boolean

'----------------------------------------------------------------------
' Safe sheet accessor. Returns Nothing if the sheet is gone.
'----------------------------------------------------------------------
Public Function GetSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

'----------------------------------------------------------------------
' Same, but raises a clear error naming the missing sheet. Use where
' the routine genuinely cannot continue.
'----------------------------------------------------------------------
Public Function RequireSheet(ByVal nm As String) As Worksheet
    Set RequireSheet = GetSheet(nm)
    If RequireSheet Is Nothing Then
        Err.Raise vbObjectError + 513, "modConfig.RequireSheet", _
                  "Required sheet '" & nm & "' was not found in this workbook."
    End If
End Function

' Deliberately NOT called SheetExists. Update_Analysis already exports a
' Public Function SheetExists with a different signature, and two public
' procedures of the same name in different standard modules make every
' unqualified call an "Ambiguous name detected" compile error. Renaming
' ours is the one change that costs nothing and touches no legacy code.
Public Function HasSheet(ByVal nm As String) As Boolean
    HasSheet = Not GetSheet(nm) Is Nothing
End Function

'----------------------------------------------------------------------
' Adds a sheet after anchorName, or at the end if the anchor is gone.
' The original modules passed GetSheet(...) straight into After:=, which
' raises 1004 the moment the anchor does not exist yet.
'----------------------------------------------------------------------
Public Function AddSheetAfter(ByVal newName As String, ByVal anchorName As String) As Worksheet
    Dim ws As Worksheet, anchor As Worksheet
    Set ws = GetSheet(newName)
    If Not ws Is Nothing Then
        Set AddSheetAfter = ws
        Exit Function
    End If
    Set anchor = GetSheet(anchorName)
    If anchor Is Nothing Then
        Set anchor = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=anchor)
    ws.Name = newName
    Set AddSheetAfter = ws
End Function

Public Function OpsArray() As Variant
    OpsArray = Split(OPS_CSV, ",")
End Function

Public Function LastRow(ByVal ws As Worksheet, Optional ByVal col As String = "A") As Long
    If ws Is Nothing Then Exit Function
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Function Nz(ByVal v As Variant) As Double
    If IsNumeric(v) Then Nz = CDbl(v)
End Function

Public Function NzDate(ByVal v As Variant, ByVal fallback As Date) As Date
    If IsDate(v) Then NzDate = CDate(v) Else NzDate = fallback
End Function

'======================================================================
' APPLICATION STATE  -  save and restore, never hard-set
'======================================================================

Public Sub FastOn()
    If mDepth = 0 Then
        mCalc = Application.Calculation
        mScreen = Application.ScreenUpdating
        mEvents = Application.EnableEvents
        mAlerts = Application.DisplayAlerts
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
    End If
    mDepth = mDepth + 1
End Sub

Public Sub FastOff()
    If mDepth > 0 Then mDepth = mDepth - 1
    If mDepth = 0 Then
        On Error Resume Next
        Application.DisplayAlerts = mAlerts
        Application.EnableEvents = mEvents
        Application.ScreenUpdating = mScreen
        Application.Calculation = mCalc
        On Error GoTo 0
    End If
End Sub

' Force the stack back to zero after an unhandled error.
Public Sub FastReset()
    mDepth = 1
    FastOff
End Sub

'======================================================================
' OUTPUT  -  MsgBox when a human is driving, log file always
'======================================================================

' The driver calls this before anything else. In SilentMode no MsgBox
' is ever shown, because a modal dialog in an unattended COM run hangs
' Excel with no window for anyone to click.
Public Sub SetSilent(ByVal b As Boolean)
    SilentMode = b
    LogLine "SilentMode = " & b
End Sub

Public Sub Say(ByVal msg As String, Optional ByVal title As String = "REVO", _
               Optional ByVal style As VbMsgBoxStyle = vbInformation)
    LogLine title & ": " & Replace(Replace(msg, vbCrLf, " | "), vbCr, " ")
    If Not SilentMode Then MsgBox msg, style, title
End Sub

' Returns True to proceed. In SilentMode it always proceeds - the driver
' has already taken responsibility for working on a copy.
Public Function Confirm(ByVal msg As String, Optional ByVal title As String = "REVO") As Boolean
    If SilentMode Then
        LogLine title & ": [auto-confirmed] " & Replace(msg, vbCrLf, " | ")
        Confirm = True
        Exit Function
    End If
    Confirm = (MsgBox(msg, vbYesNo + vbExclamation, title) = vbYes)
End Function

Public Sub LogLine(ByVal s As String)
    Dim f As Integer, p As String
    On Error Resume Next
    p = ThisWorkbook.Path
    If Len(p) = 0 Then Exit Sub
    f = FreeFile
    Open p & Application.PathSeparator & LOG_FILE For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & s
    Close #f
    On Error GoTo 0
End Sub
