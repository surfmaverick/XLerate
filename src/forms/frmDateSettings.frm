VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDateSettings 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmDateSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDateSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmDateSettings
Option Explicit

Private FormatListBox As MSForms.ListBox
Private txtName As MSForms.TextBox
Private txtFormat As MSForms.TextBox
Private DynamicButtonHandlers As Collection
Private ListBoxHandler As clsListBoxHandler

Public Sub InitializeInPanel(parentFrame As MSForms.Frame)
    On Error GoTo ErrorHandler
    Debug.Print "=== Date Settings Initialize START ==="
    
    Set DynamicButtonHandlers = New Collection
    
    ModDateFormat.InitializeDateFormats
    InitializeControlsInPanel parentFrame
        
    RefreshFormatListBox
    
    Set ListBoxHandler = New clsListBoxHandler
    ListBoxHandler.Initialize FormatListBox, Me
    
    If FormatListBox.ListCount > 0 Then
        FormatListBox.ListIndex = 0
        Dim formats() As clsFormatType
        formats = ModDateFormat.GetFormatList()
        UpdateTextBoxes formats(0).Name, formats(0).FormatCode
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "Error in InitializeInPanel: " & Err.Description
    Resume Next
End Sub

Private Sub InitializeControlsInPanel(parentFrame As MSForms.Frame)
    Set FormatListBox = parentFrame.Controls.Add("Forms.ListBox.1", "FormatListBox")
    With FormatListBox
        .Left = 10
        .Top = 10
        .Width = 390
        .Height = 200
        .MultiSelect = 0
    End With
    
    Set txtName = parentFrame.Controls.Add("Forms.TextBox.1", "txtName")
    With txtName
        .Left = 10
        .Top = 270
        .Width = 290
        .Height = 20
    End With
    
    Set txtFormat = parentFrame.Controls.Add("Forms.TextBox.1", "txtFormat")
    With txtFormat
        .Left = 10
        .Top = 320
        .Width = 290
        .Height = 20
    End With
    
    Dim lblName As MSForms.Label
    Set lblName = parentFrame.Controls.Add("Forms.Label.1", "lblName")
    With lblName
        .Left = 10
        .Top = 250
        .Caption = "Name:"
    End With
    
    Dim lblFormat As MSForms.Label
    Set lblFormat = parentFrame.Controls.Add("Forms.Label.1", "lblFormat")
    With lblFormat
        .Left = 10
        .Top = 300
        .Caption = "Format:"
    End With
    
    CreateActionButtonsInPanel parentFrame
End Sub

Private Sub CreateActionButtonsInPanel(parentFrame As MSForms.Frame)
    Debug.Print "=== Creating Action Buttons START ==="
    On Error GoTo ErrorHandler
    
    Dim btnAdd As MSForms.CommandButton
    Set btnAdd = parentFrame.Controls.Add("Forms.CommandButton.1", "btnAdd")
    With btnAdd
        .Left = 310
        .Top = 220
        .Width = 80
        .Height = 25
        .Caption = "Add"
    End With
    Debug.Print "Created button: " & btnAdd.Caption
    AttachButtonHandler btnAdd, "Add"
    
    Dim btnRemove As MSForms.CommandButton
    Set btnRemove = parentFrame.Controls.Add("Forms.CommandButton.1", "btnRemove")
    With btnRemove
        .Left = 310
        .Top = 250
        .Width = 80
        .Height = 25
        .Caption = "Remove"
    End With
    Debug.Print "Created button: " & btnRemove.Caption
    AttachButtonHandler btnRemove, "Remove"
    
    Dim btnSave As MSForms.CommandButton
    Set btnSave = parentFrame.Controls.Add("Forms.CommandButton.1", "btnSave")
    With btnSave
        .Left = 310
        .Top = 290
        .Width = 80
        .Height = 25
        .Caption = "Save"
    End With
    Debug.Print "Created button: " & btnSave.Caption
    AttachButtonHandler btnSave, "Save"
    
    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = parentFrame.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Left = 310
        .Top = 320
        .Width = 80
        .Height = 25
        .Caption = "Cancel"
    End With
    Debug.Print "Created button: " & btnCancel.Caption
    AttachButtonHandler btnCancel, "Cancel"
    
    Exit Sub
ErrorHandler:
        Debug.Print "Error in CreateActionButtonsInPanel: " & Err.Description
        Resume Next

End Sub

Private Sub AttachButtonHandler(ByRef Button As MSForms.CommandButton, ByVal Role As String)
    Dim ButtonHandler As clsDynamicButtonHandler
    Set ButtonHandler = New clsDynamicButtonHandler
    ButtonHandler.Initialize Button, Role, Me
    DynamicButtonHandlers.Add ButtonHandler
End Sub

Public Sub btnAdd_Click()
    Debug.Print "=== btnAdd_Click START ==="
    On Error GoTo ErrorHandler
    Dim newName As String
    newName = InputBox("Enter the name for the new format:", "Add Format")
    If newName <> "" Then
        Dim newFormat As New clsFormatType
        Debug.Print "Creating new format: " & newName
        newFormat.Name = newName
        newFormat.FormatCode = "yyyy"
        ModDateFormat.AddFormat newFormat
        RefreshFormatListBox
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "Error in btnAdd_Click: " & Err.Description
    Resume Next
End Sub

Public Sub btnRemove_Click()
    If FormatListBox.ListIndex >= 0 Then
        ModDateFormat.RemoveFormat FormatListBox.ListIndex
        RefreshFormatListBox
    Else
        MsgBox "Please select a format to remove.", vbExclamation
    End If
End Sub

Public Sub btnSave_Click()
   Debug.Print "=== btnSave_Click START ==="
   On Error GoTo ErrorHandler
   
   If FormatListBox.ListIndex >= 0 Then
       Dim selectedIndex As Integer
       selectedIndex = FormatListBox.ListIndex
       Debug.Print "Selected index: " & selectedIndex
       
       Dim updatedFormat As New clsFormatType
       updatedFormat.Name = txtName.Text
       updatedFormat.FormatCode = txtFormat.Text
       Debug.Print "Updating format - Name: " & updatedFormat.Name & " | Format: " & updatedFormat.FormatCode
       
       ModDateFormat.UpdateFormat selectedIndex, updatedFormat
       ModDateFormat.SaveFormatsToWorkbook
       RefreshFormatListBox
       FormatListBox.ListIndex = selectedIndex
       Debug.Print "Format updated and saved successfully"
   Else
       MsgBox "Please select a format to save.", vbExclamation
       Debug.Print "No format selected"
   End If
   Debug.Print "=== btnSave_Click END ==="
   Exit Sub
   
ErrorHandler:
   Debug.Print "Error in btnSave_Click: " & Err.Description
   MsgBox "Error saving format: " & Err.Description, vbExclamation
   Resume Next
End Sub

Public Sub btnCancel_Click()
    ModDateFormat.LoadFormatsFromWorkbook  ' Reload original formats
    RefreshFormatListBox
End Sub

Public Sub UpdateTextBoxes(ByVal Name As String, ByVal format As String)
    txtName.Text = Name
    txtFormat.Text = format
End Sub

Private Sub txtName_Change()
    If FormatListBox.ListIndex >= 0 Then
        Dim updatedFormat As New clsFormatType
        updatedFormat.Name = txtName.Text
        updatedFormat.FormatCode = txtFormat.Text
        ModDateFormat.UpdateFormat FormatListBox.ListIndex, updatedFormat
        FormatListBox.List(FormatListBox.ListIndex) = txtName.Text
    End If
End Sub

Private Sub txtFormat_Change()
    If FormatListBox.ListIndex >= 0 Then
        Dim updatedFormat As New clsFormatType
        updatedFormat.Name = txtName.Text
        updatedFormat.FormatCode = txtFormat.Text
        ModDateFormat.UpdateFormat FormatListBox.ListIndex, updatedFormat
    End If
End Sub

Private Sub RefreshFormatListBox()
    Debug.Print "=== RefreshFormatListBox START ==="
    FormatListBox.Clear
    
    Dim formats() As clsFormatType
    formats = ModDateFormat.GetFormatList()
    Debug.Print "Got " & (UBound(formats) - LBound(formats) + 1) & " formats"
    
    Dim i As Integer
    For i = LBound(formats) To UBound(formats)
        Debug.Print "Adding to ListBox: " & formats(i).Name
        FormatListBox.AddItem formats(i).Name
    Next i
    
    If FormatListBox.ListCount > 0 Then
        FormatListBox.ListIndex = 0
        Debug.Print "Set initial selection"
    End If
    Debug.Print "=== RefreshFormatListBox END ==="
End Sub
