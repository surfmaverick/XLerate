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
    ModNumberFormat.InitializeFormats
    ModNumberFormat.GetFormatList ' Force array initialization
    Application.OnKey "^+1", "CycleNumberFormat"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.OnKey "^+1" ' Clear the shortcut
End Sub
