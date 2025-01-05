VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNumberSettings 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmNumberSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNumberSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmNumberSettings
Option Explicit

Private FormatListBox As MSForms.ListBox
Private txtName As MSForms.TextBox
Private txtFormat As MSForms.TextBox
Private DynamicButtonHandlers As Collection
Private ListBoxHandler As clsListBoxHandler

Public Sub InitializeInPanel(parentFrame As MSForms.Frame)
    Debug.Print "Number Settings Initialize started"
    On Error GoTo ErrorHandler
    
    ' Initialize the formats module
    Debug.Print "Initializing formats"
    ModNumberFormat.InitializeFormats
    
    Debug.Print "Initializing controls in panel"
    InitializeControlsInPanel parentFrame
    
    ' Initialize the collection
    Debug.Print "Initializing button handlers collection"
    If DynamicButtonHandlers Is Nothing Then Set DynamicButtonHandlers = New Collection
    
    ' Load initial data
    Debug.Print "Refreshing format list box"
    RefreshFormatListBox
    
    ' Set up list box handler
    Debug.Print "Setting up list box handler"
    Set ListBoxHandler = New clsListBoxHandler
    ListBoxHandler.Initialize FormatListBox, Me
    
    If FormatListBox.ListCount > 0 Then
        FormatListBox.ListIndex = 0
        Dim formats() As clsFormatType
        formats = ModNumberFormat.GetFormatList()
        UpdateTextBoxes formats(0).Name, formats(0).FormatCode
    End If

    Debug.Print "Number Settings Initialize completed"
    Exit Sub

ErrorHandler:
    Debug.Print "Error in Number Settings Initialize: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Private Sub InitializeControlsInPanel(parentFrame As MSForms.Frame)
    ' Create format list box
    Set FormatListBox = parentFrame.Controls.Add("Forms.ListBox.1", "FormatListBox")
    With FormatListBox
        .Left = 10
        .Top = 10
        .Width = 390
        .Height = 200
        .MultiSelect = 0
    End With
    
    ' Create text boxes and labels
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
    
    ' Add labels
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
    Debug.Print "Creating action buttons in panel"
    
    ' Create standard buttons
    Dim btnAdd As MSForms.CommandButton
    Set btnAdd = parentFrame.Controls.Add("Forms.CommandButton.1", "btnAdd")
    With btnAdd
        .Left = 310
        .Top = 220
        .Width = 80
        .Height = 25
        .Caption = "Add"
    End With
    Debug.Print "Add button created"
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
    Debug.Print "Remove button created"
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
    Debug.Print "Save button created"
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
    Debug.Print "Cancel button created"
    AttachButtonHandler btnCancel, "Cancel"
End Sub

Private Sub AttachButtonHandler(ByRef Button As MSForms.CommandButton, ByVal Role As String)
    Debug.Print "Attaching button handler for role: " & Role
    On Error GoTo ErrorHandler
    
    Dim ButtonHandler As clsDynamicButtonHandler
    Set ButtonHandler = New clsDynamicButtonHandler
    ButtonHandler.Initialize Button, Role, Me
    
    If DynamicButtonHandlers Is Nothing Then
        Debug.Print "Creating new DynamicButtonHandlers collection"
        Set DynamicButtonHandlers = New Collection
    End If
    DynamicButtonHandlers.Add ButtonHandler
    Debug.Print "Handler attached successfully for role: " & Role
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AttachButtonHandler: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Public Sub btnAdd_Click()
    Debug.Print "btnAdd_Click started"
    Dim newName As String
    Dim newFormat As clsFormatType
    newName = InputBox("Enter the name for the new format:", "Add Format")
    If newName <> "" Then
        Debug.Print "Creating new format: " & newName
        Set newFormat = New clsFormatType
        newFormat.Name = newName
        newFormat.FormatCode = "General"
        ModNumberFormat.AddFormat newFormat
        RefreshFormatListBox
    End If
    Debug.Print "btnAdd_Click completed"
End Sub

Public Sub btnRemove_Click()
    Debug.Print "btnRemove_Click started"
    If FormatListBox.ListIndex >= 0 Then
        Dim selectedIndex As Integer
        selectedIndex = FormatListBox.ListIndex
        Debug.Print "Removing format at index: " & selectedIndex
        ModNumberFormat.RemoveFormat selectedIndex
        RefreshFormatListBox
    Else
        MsgBox "Please select a format to remove.", vbExclamation
    End If
    Debug.Print "btnRemove_Click completed"
End Sub

Public Sub btnSave_Click()
    Debug.Print "btnSave_Click started"
    If FormatListBox.ListIndex >= 0 Then
        Dim selectedIndex As Integer
        selectedIndex = FormatListBox.ListIndex
        Debug.Print "Saving format at index: " & selectedIndex
               
        Dim updatedFormat As clsFormatType
        Set updatedFormat = New clsFormatType
        updatedFormat.Name = txtName.Text
        updatedFormat.FormatCode = txtFormat.Text
        
        ModNumberFormat.UpdateFormat selectedIndex, updatedFormat
        ModNumberFormat.SaveFormatsToWorkbook
        RefreshFormatListBox
        FormatListBox.ListIndex = selectedIndex
    Else
        MsgBox "Please select a format to save.", vbExclamation
    End If
    Debug.Print "btnSave_Click completed"
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
        ModNumberFormat.UpdateFormat FormatListBox.ListIndex, updatedFormat
        FormatListBox.List(FormatListBox.ListIndex) = txtName.Text
    End If
End Sub

Private Sub txtFormat_Change()
    If FormatListBox.ListIndex >= 0 Then
        Dim updatedFormat As New clsFormatType
        updatedFormat.Name = txtName.Text
        updatedFormat.FormatCode = txtFormat.Text
        ModNumberFormat.UpdateFormat FormatListBox.ListIndex, updatedFormat
    End If
End Sub

Private Sub RefreshFormatListBox()
    FormatListBox.Clear
    
    Dim formats() As clsFormatType
    formats = ModNumberFormat.GetFormatList()
    
    Dim i As Integer
    For i = LBound(formats) To UBound(formats)
        Debug.Print "Adding format to ListBox [" & i & "]: " & formats(i).Name
        FormatListBox.AddItem formats(i).Name
    Next i
    
    If FormatListBox.ListCount > 0 Then
        FormatListBox.ListIndex = 0  ' Select first item
        Debug.Print "Set initial selection to index 0"
    End If
End Sub
