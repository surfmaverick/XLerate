VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' In ThisWorkbook
Private Sub Workbook_Open()
    Debug.Print "Workbook_Open triggered"
    ' Initialize number formats
    ModNumberFormat.InitializeFormats
    ModNumberFormat.GetFormatList
    Application.OnKey "^+1", "CycleNumberFormat"
    Debug.Print "Shortcut Ctrl + Shift + 1 assigned to CycleNumberFormat"
    
    ' Initialize cell formats
    ModCellFormat.InitializeCellFormats
    ModCellFormat.GetCellFormatList
    Application.OnKey "^+2", "CycleCellFormat"
    Debug.Print "Shortcut Ctrl + Shift + 2 assigned to CycleCellFormat"
    
    ' Initialize date formats
    ModDateFormat.InitializeDateFormats
    ModDateFormat.GetFormatList
    Application.OnKey "^+3", "CycleDateFormat"
    Debug.Print "Shortcut Ctrl + Shift + 3 assigned to CycleDateFormat"
    
    ' Add reset shortcut
    Application.OnKey "^+0", "ResetAllFormatsToDefaults"  ' Remove module name
    Debug.Print "Shortcut Ctrl + Shift + 0 assigned to ResetAllFormatsToDefaults"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clear the shortcuts
    Application.OnKey "^+1"
    Application.OnKey "^+2"
    Application.OnKey "^+3"
    Application.OnKey "^+0"
End Sub
