VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrecedents 
   Caption         =   "Trace Precedents"
   ClientHeight    =   3500
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5610
   OleObjectBlob   =   "frmPrecedents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrecedents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstPrecedents_Click()
    On Error Resume Next
    Dim precedentAddress As String
    precedentAddress = lstPrecedents.value
    Debug.Print "precedentAddress: " & precedentAddress
    
    'Parse out sheet name and cell address
    Dim exclamationPosition As Integer
    exclamationPosition = InStr(precedentAddress, "!")
    If exclamationPosition > 0 Then
        Dim sheetName As String
        sheetName = Mid(precedentAddress, InStrRev(precedentAddress, "]") + 1, exclamationPosition - InStrRev(precedentAddress, "]") - 1)
        Dim cellAddress As String
        cellAddress = Mid(precedentAddress, exclamationPosition + 1)
        
        ' Check for trailing single quote
        If Right(sheetName, 1) = "'" Then
            ' Remove trailing single quote
            sheetName = Left(sheetName, Len(sheetName) - 1)
        End If
        
        ' Activate the Sheet
        Worksheets(sheetName).Activate
        Worksheets(sheetName).Range(cellAddress).Select
    Else
        Range(precedentAddress).Select
    End If
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    
    Unload Me

End Sub

Private Sub UserForm_Click()

End Sub

' linked to the ListBox (lstPrecedents) since that's the element that likely has focus when the UserForm is open
Private Sub lstPrecedents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

