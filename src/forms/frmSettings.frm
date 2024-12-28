VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Number Formats"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' In the frmSettings UserForm code module
Option Explicit

Private Sub UserForm_Initialize()
    ' Set form caption
    Me.Caption = "Format Settings"
    
    ' Center the form
    Me.StartUpPosition = 0 ' Manual
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    
    LoadSettings ' Load current settings from document properties
    
    ' Load current settings into form controls
    txtFormat1.Text = Format1
    txtFormat2.Text = Format2
    txtFormat3.Text = Format3
    
    chkFormat1.value = (EnabledFormats And 1) > 0
    chkFormat2.value = (EnabledFormats And 2) > 0
    chkFormat3.value = (EnabledFormats And 4) > 0
    
    ' Set up labels
    lblFormat1.Caption = "Format 1:"
    lblFormat2.Caption = "Format 2:"
    lblFormat3.Caption = "Format 3:"
    
    ' Set up tooltips
    txtFormat1.ControlTipText = "Enter the first number format (e.g. General)"
    txtFormat2.ControlTipText = "Enter the second number format (e.g. #,##0.0;(#,##0.0);""-"")"
    txtFormat3.ControlTipText = "Enter the third number format (e.g. #,##0.00;(#,##0.00);""-"")"
    
    ' Set up buttons
    btnOK.Caption = "OK"
    btnCancel.Caption = "Cancel"
    
    ' Set up checkboxes
    chkFormat1.Caption = "Enable"
    chkFormat2.Caption = "Enable"
    chkFormat3.Caption = "Enable"
End Sub

Private Sub btnOK_Click()
    ' Validate formats
    If Not IsValidNumberFormat(txtFormat1.Text) Or _
       Not IsValidNumberFormat(txtFormat2.Text) Or _
       Not IsValidNumberFormat(txtFormat3.Text) Then
        MsgBox "One or more number formats are invalid. Please check your input.", vbExclamation
        Exit Sub
    End If
    
    ' Save settings
    Format1 = txtFormat1.Text
    Format2 = txtFormat2.Text
    Format3 = txtFormat3.Text
    
    Dim newEnabledFormats As Integer
    newEnabledFormats = 0
    If chkFormat1.value Then newEnabledFormats = newEnabledFormats Or 1
    If chkFormat2.value Then newEnabledFormats = newEnabledFormats Or 2
    If chkFormat3.value Then newEnabledFormats = newEnabledFormats Or 4
    
    ' Ensure at least one format is enabled
    If newEnabledFormats = 0 Then
        MsgBox "At least one format must be enabled.", vbExclamation
        Exit Sub
    End If
    
    EnabledFormats = newEnabledFormats
    SaveSettings
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function IsValidNumberFormat(ByVal formatString As String) As Boolean
    On Error Resume Next
    Dim testCell As Range
    Set testCell = ActiveSheet.Range("A1") ' Use any cell
    
    Dim originalFormat As String
    originalFormat = testCell.NumberFormat
    
    testCell.NumberFormat = formatString
    IsValidNumberFormat = (Err.Number = 0)
    
    testCell.NumberFormat = originalFormat
    On Error GoTo 0
End Function

' Optional: Add these event handlers for better UX
Private Sub txtFormat1_Change()
    HighlightInvalidFormat txtFormat1
End Sub

Private Sub txtFormat2_Change()
    HighlightInvalidFormat txtFormat2
End Sub

Private Sub txtFormat3_Change()
    HighlightInvalidFormat txtFormat3
End Sub

Private Sub HighlightInvalidFormat(txt As MSForms.TextBox)
    If txt.Text = "" Then
        txt.BackColor = &H80000005 ' White
    ElseIf Not IsValidNumberFormat(txt.Text) Then
        txt.BackColor = &H8080FF ' Light red
    Else
        txt.BackColor = &H80000005 ' White
    End If
End Sub



