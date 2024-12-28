VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDependents 
   Caption         =   "Trace Dependents"
   ClientHeight    =   3500
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5610
   OleObjectBlob   =   "frmDependents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDependents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstDependents_Click()
    On Error Resume Next
    Dim dependentAddress As String
    dependentAddress = lstDependents.Value
    Debug.Print "dependentAddress: " & dependentAddress
    
    'Parse out sheet name and cell address
    Dim exclamationPosition As Integer
    exclamationPosition = InStr(dependentAddress, "!")
    If exclamationPosition > 0 Then
        Dim sheetName As String
        sheetName = Mid(dependentAddress, InStrRev(dependentAddress, "]") + 1, exclamationPosition - InStrRev(dependentAddress, "]") - 1)
        Dim cellAddress As String
        cellAddress = Mid(dependentAddress, exclamationPosition + 1)
        
        ' Check for trailing single quote
        If Right(sheetName, 1) = "'" Then
            ' Remove trailing single quote
            sheetName = Left(sheetName, Len(sheetName) - 1)
        End If
        
        ' Activate the Sheet
        Worksheets(sheetName).Activate
        Worksheets(sheetName).Range(cellAddress).Select
    Else
        Range(dependentAddress).Select
    End If
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    
    Unload Me

End Sub

Private Sub UserForm_Click()

End Sub

' linked to the ListBox (lstdependents) since that's the element that likely has focus when the UserForm is open
Private Sub lstDependents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


