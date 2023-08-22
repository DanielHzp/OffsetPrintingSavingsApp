Attribute VB_Name = "Module3"
Option Explicit


'Controls data export worksheet behavior
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Sheet1").Select
    Sheets("Sheet2").Visible = True
    
    
    
End Sub
