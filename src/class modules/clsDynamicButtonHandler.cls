VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDynamicButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents DynamicButton As MSForms.CommandButton
Attribute DynamicButton.VB_VarHelpID = -1
Private ButtonRole As String
Private ParentForm As frmSettingsManager

Public Sub Initialize(ByRef Button As MSForms.CommandButton, ByVal Role As String, ByRef Parent As frmSettingsManager)
    Set DynamicButton = Button
    ButtonRole = Role
    Set ParentForm = Parent
End Sub

Private Sub DynamicButton_Click()
    Select Case ButtonRole
        Case "Add"
            ParentForm.btnAdd_Click
        Case "Remove"
            ParentForm.btnRemove_Click
        Case "Save"
            ParentForm.btnOK_Click
        Case "Cancel"
            ParentForm.btnCancel_Click
    End Select
End Sub

