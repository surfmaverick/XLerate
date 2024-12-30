VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettingsManager 
   Caption         =   "Settings"
   ClientHeight    =   6640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10780
   OleObjectBlob   =   "frmSettingsManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettingsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Control declarations
Private lstCategories As MSForms.ListBox
Private NumbersPanel As MSForms.Frame
Private FormatListBox As MSForms.ListBox
Private txtName As MSForms.TextBox
Private txtFormat As MSForms.TextBox
Private DynamicButtonHandlers As Collection
Private ListBoxHandler As clsListBoxHandler

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Debug.Print "Form Initialize started"
    
    ' Initialize the formats module at form load
    ModNumberFormat.InitializeFormats
    
    ' Initialize form layout
    Me.BackColor = RGB(255, 255, 255)
    Me.Caption = "Settings"
    Me.Width = 600
    Me.Height = 400
    Debug.Print "Form layout set"
    
    ' Initialize the collection
    If DynamicButtonHandlers Is Nothing Then Set DynamicButtonHandlers = New Collection

    ' Create navigation listbox
    Set lstCategories = Me.Controls.Add("Forms.ListBox.1", "lstCategories")
    With lstCategories
        .Left = 12
        .Top = 12
        .Width = 150
        .Height = 300
    End With
    Debug.Print "Created lstCategories"
    
    ' Create Numbers panel frame
    Set NumbersPanel = Me.Controls.Add("Forms.Frame.1", "NumbersPanel")
    With NumbersPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 350
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
    End With
    Debug.Print "Created NumbersPanel"
    
    ' Create format list box within NumbersPanel
    Set FormatListBox = NumbersPanel.Controls.Add("Forms.ListBox.1", "FormatListBox")
    With FormatListBox
        .Left = 10
        .Top = 10
        .Width = 390
        .Height = 200
        .MultiSelect = 0      ' Ensure single selection mode
        Debug.Print "FormatListBox Settings:"
        Debug.Print "  Style: " & .Style
        Debug.Print "  ListStyle: " & .ListStyle
    End With
    Debug.Print "Created FormatListBox"
    
    ' Create text boxes for Name and Format within NumbersPanel
    Set txtName = NumbersPanel.Controls.Add("Forms.TextBox.1", "txtName")
    With txtName
        .Left = 10
        .Top = 270
        .Width = 290
        .Height = 20
    End With
    
    Set txtFormat = NumbersPanel.Controls.Add("Forms.TextBox.1", "txtFormat")
    With txtFormat
        .Left = 10
        .Top = 320
        .Width = 290
        .Height = 20
    End With
    Debug.Print "Created text boxes"
    
    ' In UserForm_Initialize, add labels
    Dim lblName As MSForms.Label
    Set lblName = NumbersPanel.Controls.Add("Forms.Label.1", "lblName")
    With lblName
        .Left = 10
        .Top = 250
        .Caption = "Name:"
    End With
    
    Dim lblFormat As MSForms.Label
    Set lblFormat = NumbersPanel.Controls.Add("Forms.Label.1", "lblFormat")
    With lblFormat
        .Left = 10
        .Top = 300
        .Caption = "Format:"
    End With
    
    ' Create Add button
    Dim btnAdd As MSForms.CommandButton
    Set btnAdd = NumbersPanel.Controls.Add("Forms.CommandButton.1", "btnAdd")
    With btnAdd
        .Left = 310
        .Top = 220
        .Width = 80
        .Height = 25
        .Caption = "Add"
    End With
    AttachButtonHandler btnAdd, "Add"
    Debug.Print "Created Add button"
    
    ' Create Remove button
    Dim btnRemove As MSForms.CommandButton
    Set btnRemove = NumbersPanel.Controls.Add("Forms.CommandButton.1", "btnRemove")
    With btnRemove
        .Left = 310
        .Top = 250
        .Width = 80
        .Height = 25
        .Caption = "Remove"
    End With
    AttachButtonHandler btnRemove, "Remove"
    Debug.Print "Created Remove button"
    
    ' In UserForm_Initialize, after other controls
    Dim btnSave As MSForms.CommandButton
    Set btnSave = NumbersPanel.Controls.Add("Forms.CommandButton.1", "btnSave")
    With btnSave
        .Left = 310
        .Top = 290
        .Width = 80
        .Height = 25
        .Caption = "Save"
    End With
    AttachButtonHandler btnSave, "Save"
    
    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = NumbersPanel.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Left = 310
        .Top = 320
        .Width = 80
        .Height = 25
        .Caption = "Cancel"
    End With
    AttachButtonHandler btnCancel, "Cancel"
        
    ' Initialize the numbers panel
    RefreshFormatListBox  ' Load the formats first
    
    Set ListBoxHandler = New clsListBoxHandler  ' Set up handler
    ListBoxHandler.Initialize FormatListBox, Me
    
    InitializeHierarchyList  ' This will select Numbers and show the panel
    
    ' Populate the textboxes with the first item's data
    If FormatListBox.ListCount > 0 Then
        FormatListBox.ListIndex = 0
        Dim formats() As clsFormatType
        formats = ModNumberFormat.GetFormatList()
        UpdateTextBoxes formats(0).Name, formats(0).FormatCode
        Debug.Print "Initial Selection: Name = " & txtName.Text & ", Format = " & txtFormat.Text
    End If

ErrorHandler:
    Debug.Print "Error in UserForm_Initialize: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub


Private Sub InitializeNumbersPanel()
    RefreshFormatListBox
    ShowPanel "Numbers"
End Sub

Public Sub UpdateTextBoxes(ByVal Name As String, ByVal format As String)
    txtName.Text = Name
    txtFormat.Text = format
End Sub

' Show the requested panel and hide others
Private Sub ShowPanel(panelName As String)
    Select Case panelName
        Case "Numbers"
            NumbersPanel.Visible = True
        Case "None"
            NumbersPanel.Visible = False
        Case Else
            NumbersPanel.Visible = False
    End Select
End Sub

' Add this event handler for the listbox
Private Sub lstCategories_Click()
    Dim selectedCategory As String
    selectedCategory = Trim(lstCategories.Text)
    
    ' If header is clicked, force selection back to Numbers
    If lstCategories.List(lstCategories.ListIndex, 1) = "HEADER" Then
        lstCategories.ListIndex = 1  ' Select Numbers
        Exit Sub
    End If
    
    Select Case selectedCategory
        Case "  Numbers"
            ShowPanel "Numbers"
        Case Else
            ShowPanel "None"
    End Select
End Sub


Private Sub InitializeFormLayout()
    Me.BackColor = RGB(255, 255, 255)
    Me.Caption = "Settings"
    Me.Width = 600
    Me.Height = 400
    Debug.Print "Form layout set"
End Sub

Private Sub InitializeHierarchyList()
    lstCategories.Clear

    ' Add top-level category as header (disabled)
    With lstCategories
        .AddItem "Formatting"
        .List(.ListCount - 1, 1) = "HEADER"  ' Use second column to mark as header
        .AddItem "  Numbers"  ' Subcategory with indentation
        .ListIndex = 1  ' Select Numbers by default
    End With
    
    ' Initialize with Numbers panel shown
    ShowPanel "Numbers"
End Sub

Private Sub AttachButtonHandler(ByRef Button As MSForms.CommandButton, ByVal Role As String)
    Dim ButtonHandler As clsDynamicButtonHandler
    ' Create a new handler for the button
    Set ButtonHandler = New clsDynamicButtonHandler
    ' Initialize the handler with the button, its role, and the parent form
    ButtonHandler.Initialize Button, Role, Me
    ' Add the handler to the collection
    DynamicButtonHandlers.Add ButtonHandler
End Sub


' Add a new format to the list
Public Sub btnAdd_Click()
    Dim newName As String
    Dim newFormat As clsFormatType

    newName = InputBox("Enter the name for the new format:", "Add Format")
    If newName <> "" Then
        ' Create a new format object
        Set newFormat = New clsFormatType
        newFormat.Name = newName
        newFormat.FormatCode = "General" ' Default format code

        ' Add to the ModNumberFormat list
        ModNumberFormat.AddFormat newFormat

        ' Refresh the ListBox
        RefreshFormatListBox
    End If
End Sub

Public Sub btnRemove_Click()
    If FormatListBox.ListIndex >= 0 Then
        Dim selectedIndex As Integer
        selectedIndex = FormatListBox.ListIndex
        
        ' Remove from ModNumberFormat list
        ModNumberFormat.RemoveFormat selectedIndex
        
        ' Refresh the ListBox to reflect the removed item
        RefreshFormatListBox
    Else
        MsgBox "Please select a format to remove.", vbExclamation
    End If
End Sub

' Save changes to the selected format
Public Sub btnOK_Click()
    If FormatListBox.ListIndex >= 0 Then
        Dim selectedIndex As Integer
        selectedIndex = FormatListBox.ListIndex
               
        ' Update the selected format in FormatList
        Dim updatedFormat As clsFormatType
        Set updatedFormat = New clsFormatType
        updatedFormat.Name = txtName.Text
        updatedFormat.FormatCode = txtFormat.Text
        
        ModNumberFormat.UpdateFormat selectedIndex, updatedFormat
        
        ' Save the updated FormatList to the workbook
        ModNumberFormat.SaveFormatsToWorkbook
        
        ' Refresh the ListBox to reflect the updated name
        RefreshFormatListBox
        
        ' Reselect the updated item to show changes
        FormatListBox.ListIndex = selectedIndex
    Else
        MsgBox "Please select a format to save.", vbExclamation
    End If
End Sub


' Cancel changes and close the form
Public Sub btnCancel_Click()
    Unload Me
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

