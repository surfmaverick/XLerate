VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListBoxHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' In clsListBoxHandler
Option Explicit
Private WithEvents ListBox As MSForms.ListBox
Attribute ListBox.VB_VarHelpID = -1
Private ParentForm As Object  ' Change from specific form type to Object

Public Sub Initialize(ByRef lBox As MSForms.ListBox, ByRef frm As Object)  ' Change parameter type to Object
    Set ListBox = lBox
    Set ParentForm = frm
End Sub

Private Sub ListBox_Click()
   Debug.Print "=== ListBox_Click triggered in handler ==="
   
   If ListBox.ListIndex >= 0 Then
       Debug.Print "Selected index changed to: " & ListBox.ListIndex
       Debug.Print "Selected text: " & ListBox.List(ListBox.ListIndex)
       
       If TypeOf ParentForm Is frmNumberSettings Then
           Dim formats() As clsFormatType
           formats = ModNumberFormat.GetFormatList()
           ParentForm.UpdateTextBoxes formats(ListBox.ListIndex).Name, formats(ListBox.ListIndex).FormatCode
           
       ElseIf TypeOf ParentForm Is frmDateSettings Then
           Dim dateFormats() As clsFormatType
           dateFormats = ModDateFormat.GetFormatList()
           ParentForm.UpdateTextBoxes dateFormats(ListBox.ListIndex).Name, dateFormats(ListBox.ListIndex).FormatCode
           
       ElseIf TypeOf ParentForm Is frmCellSettings Then
           Dim cellFormats() As clsCellFormatType
           cellFormats = ModCellFormat.GetCellFormatList()
           ParentForm.UpdateTextBoxes cellFormats(ListBox.ListIndex).Name
           
       Else
           Debug.Print "WARNING: Unknown form type encountered in ListBox_Click"
           Debug.Print "Form type: " & TypeName(ParentForm)
       End If
   End If
End Sub
