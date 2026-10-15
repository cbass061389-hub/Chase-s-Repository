Sub Copy_Reject_to_Weekly()
Attribute Copy_Reject_to_Weekly.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Copy_Reject_to_Weekly Macro
'

'
    Range("B4:V18").Select
    Selection.Copy
    Sheets("Weekly Reporting").Select
    Range("A20").Select
    ActiveSheet.Paste
End Sub
