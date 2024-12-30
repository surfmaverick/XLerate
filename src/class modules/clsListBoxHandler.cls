VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListBoxHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ListBox As MSForms.ListBox
Attribute ListBox.VB_VarHelpID = -1
Private ParentForm As frmSettingsManager

Public Sub Initialize(ByRef lBox As MSForms.ListBox, ByRef frm As frmSettingsManager)
    Set ListBox = lBox
    Set ParentForm = frm
End Sub

Private Sub ListBox_Click()
    Debug.Print "=== ListBox_Click triggered in handler ==="
    
    If ListBox.ListIndex >= 0 Then
        Dim formats() As clsFormatType
        formats = ModNumberFormat.GetFormatList()
        
        Dim currentIndex As Long
        currentIndex = ListBox.ListIndex
        
        Debug.Print "Selected index changed to: " & currentIndex
        Debug.Print "Selected text: " & ListBox.List(currentIndex)
        
        ' Update text boxes through parent form
        If currentIndex >= LBound(formats) And currentIndex <= UBound(formats) Then
            ParentForm.UpdateTextBoxes formats(currentIndex).Name, formats(currentIndex).FormatCode
        End If
    End If
End Sub
