VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCellSettings 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmCellSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCellSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmCellSettings
Option Explicit

Private CellFormatListBox As MSForms.ListBox
Private txtCellName As MSForms.TextBox
Private btnBackColor As MSForms.CommandButton
Private btnBorderColor As MSForms.CommandButton
Private cboBorderStyle As MSForms.ComboBox
Private DynamicButtonHandlers As Collection
Private ListBoxHandler As clsListBoxHandler
Private cboFillPattern As MSForms.ComboBox
Private btnFillColor As MSForms.CommandButton
Private cboFontStyle As MSForms.ComboBox
Private btnFontColor As MSForms.CommandButton

Public Sub InitializeInPanel(parentFrame As MSForms.Frame)
' Initializes the cell settings panel within a parent frame, setting up all UI controls,
' event handlers, and loading initial cell format data. Acts as the main setup routine
' for the cell formatting interface.

    On Error GoTo ErrorHandler
    
    ' Initialize the cell formats module
    ModCellFormat.InitializeCellFormats
    
    ' Create GUI controls
    InitializeControlsInPanel parentFrame
    
    ' Initialize button functions
    If DynamicButtonHandlers Is Nothing Then Set DynamicButtonHandlers = New Collection
    
    ' Load initial data
    RefreshCellFormatListBox
    
    ' Set up list box handler
    Set ListBoxHandler = New clsListBoxHandler
    ListBoxHandler.Initialize CellFormatListBox, Me

    ' Checking ListBox count
    If CellFormatListBox.ListCount > 0 Then
        CellFormatListBox.ListIndex = 0
        Dim formats() As clsCellFormatType
        formats = ModCellFormat.GetCellFormatList()
        UpdateTextBoxes formats(0).Name
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Error in Cell Settings Initialize at line: " & Erl
    Debug.Print "Error in Cell Settings Initialize: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Public Sub UpdateTextBoxes(ByVal Name As String)
' Updates all UI elements (textboxes, comboboxes, color buttons) to reflect the properties
' of the selected cell format. Ensures the interface stays in sync with the underlying data.

    On Error GoTo ErrorHandler
    
    If CellFormatListBox.ListIndex >= 0 Then
        Dim formats() As clsCellFormatType
        formats = ModCellFormat.GetCellFormatList()
        
        txtCellName.Text = Name
        cboBorderStyle.Text = ModCellFormat.GetBorderStyleName(formats(CellFormatListBox.ListIndex).BorderStyle)
        cboFillPattern.Text = ModCellFormat.GetFillPatternName(formats(CellFormatListBox.ListIndex).FillPattern)
        cboFontStyle.Text = ModCellFormat.GetFontStyleName(formats(CellFormatListBox.ListIndex).FontStyle)
        
        btnBorderColor.BackColor = formats(CellFormatListBox.ListIndex).BorderColor
        btnFillColor.BackColor = formats(CellFormatListBox.ListIndex).BackColor
        btnFontColor.BackColor = formats(CellFormatListBox.ListIndex).FontColor
    End If
    Debug.Print "UpdateTextBoxes completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in UpdateTextBoxes at line: " & Erl
    Debug.Print "Error in UpdateTextBoxes: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Private Sub InitializeControlsInPanel(parentFrame As MSForms.Frame)
' Creates and positions all UI controls within the parent frame, including the format listbox,
' name fields, border style selector, and color picker buttons. Handles the core UI layout.

    Debug.Print "Checking for required references..."
    On Error Resume Next
    Dim testShape As Shape
    If Err.Number <> 0 Then
        Debug.Print "WARNING: Microsoft Office Object Library reference may be missing"
        MsgBox "Required references may be missing. Please ensure Microsoft Office Object Library is referenced.", vbExclamation
    End If
    On Error GoTo 0
    
    ' Create format list box
    Set CellFormatListBox = parentFrame.Controls.Add("Forms.ListBox.1", "CellFormatListBox")
    With CellFormatListBox
        .Left = 10
        .Top = 10
        .Width = 390
        .Height = 200
        .MultiSelect = 0
    End With
    
    ' Create name textbox and label
    Set txtCellName = parentFrame.Controls.Add("Forms.TextBox.1", "txtCellName")
    With txtCellName
        .Left = 10
        .Top = 270
        .Width = 290
        .Height = 20
    End With
    
    Dim lblCellName As MSForms.Label
    Set lblCellName = parentFrame.Controls.Add("Forms.Label.1", "lblCellName")
    With lblCellName
        .Left = 10
        .Top = 250
        .Caption = "Name:"
    End With
    
    ' Create border style combo box
    Dim lblBorderStyle As MSForms.Label
    Set lblBorderStyle = parentFrame.Controls.Add("Forms.Label.1", "lblBorderStyle")
    With lblBorderStyle
        .Left = 10
        .Top = 300
        .Caption = "Border:"
    End With
    
    Set cboBorderStyle = parentFrame.Controls.Add("Forms.ComboBox.1", "cboBorderStyle")
    With cboBorderStyle
        .Left = 10
        .Top = 320
        .Width = 140
        .Height = 20
        .AddItem "None"
        .AddItem "Thin"
        .AddItem "Medium"
        .AddItem "Thick"
        .AddItem "Double"
        .AddItem "Dashed"
        .AddItem "Dotted"
    End With
    
    ' Fill Section
    Dim lblFillPattern As MSForms.Label
    Set lblFillPattern = parentFrame.Controls.Add("Forms.Label.1", "lblFillPattern")
    With lblFillPattern
        .Left = 10
        .Top = 350
        .Caption = "Fill:"
    End With
    
    Set cboFillPattern = parentFrame.Controls.Add("Forms.ComboBox.1", "cboFillPattern")
    With cboFillPattern
        .Left = 10
        .Top = 370
        .Width = 140
        .Height = 20
        .AddItem "None"
        .AddItem "Solid"
        .AddItem "25% Gray"
        .AddItem "50% Gray"
        .AddItem "75% Gray"
        .AddItem "Horizontal"
        .AddItem "Vertical"
        .AddItem "Diagonal Up"
        .AddItem "Diagonal Down"
    End With
    
    ' Font Section
    Dim lblFontStyle As MSForms.Label
    Set lblFontStyle = parentFrame.Controls.Add("Forms.Label.1", "lblFontStyle")
    With lblFontStyle
        .Left = 10
        .Top = 400
        .Caption = "Font:"
    End With
    
    Set cboFontStyle = parentFrame.Controls.Add("Forms.ComboBox.1", "cboFontStyle")
    With cboFontStyle
        .Left = 10
        .Top = 420
        .Width = 140
        .Height = 20
        .AddItem "Normal"
        .AddItem "Bold"
        .AddItem "Italic"
        .AddItem "Underline"
        .AddItem "Strike Through"
    End With
    
    ' Add buttons
    CreateActionButtonsInPanel parentFrame
End Sub

Private Sub CreateActionButtonsInPanel(parentFrame As MSForms.Frame)
' Creates and configures the action buttons (Add, Remove, Save) in the panel.
' Sets up their positions, sizes, and attaches appropriate event handlers.
   
    Dim btnAdd As MSForms.CommandButton
    Set btnAdd = parentFrame.Controls.Add("Forms.CommandButton.1", "btnAdd")
    With btnAdd
        .Left = 10
        .Top = 220
        .Width = 80
        .Height = 25
        .Caption = "Add"
    End With
    AttachButtonHandler btnAdd, "Add"
    
    Dim btnRemove As MSForms.CommandButton
    Set btnRemove = parentFrame.Controls.Add("Forms.CommandButton.1", "btnRemove")
    With btnRemove
        .Left = 100
        .Top = 220
        .Width = 80
        .Height = 25
        .Caption = "Remove"
    End With
    AttachButtonHandler btnRemove, "Remove"
    
    Dim btnSave As MSForms.CommandButton
    Set btnSave = parentFrame.Controls.Add("Forms.CommandButton.1", "btnSave")
    With btnSave
        .Left = 310
        .Top = 270
        .Width = 80
        .Height = 20
        .Caption = "Save"
    End With
    AttachButtonHandler btnSave, "Save"
    
    Set btnBackColor = parentFrame.Controls.Add("Forms.CommandButton.1", "btnBackColor")
    With btnBackColor
        .Left = 160
        .Top = 320
        .Width = 60
        .Height = 20
        .Caption = "Fill"
    End With
    AttachButtonHandler btnBackColor, "BackColor"
    
    Set btnBorderColor = parentFrame.Controls.Add("Forms.CommandButton.1", "btnBorderColor")
    With btnBorderColor
        .Left = 160
        .Top = 320
        .Width = 60
        .Height = 20
        .Caption = "Color"
    End With
    AttachButtonHandler btnBorderColor, "BorderColor"
    
    Set btnFillColor = parentFrame.Controls.Add("Forms.CommandButton.1", "btnFillColor")
    With btnFillColor
        .Left = 160
        .Top = 370
        .Width = 60
        .Height = 20
        .Caption = "Color"
    End With
    AttachButtonHandler btnFillColor, "FillColor"
    
    Set btnFontColor = parentFrame.Controls.Add("Forms.CommandButton.1", "btnFontColor")
    With btnFontColor
        .Left = 160
        .Top = 420
        .Width = 60
        .Height = 20
        .Caption = "Color"
    End With
    AttachButtonHandler btnFontColor, "FontColor"
    
End Sub

Private Sub AttachButtonHandler(ByRef Button As MSForms.CommandButton, ByVal Role As String)
' Attaches a dynamic event handler to a button based on its role. Enables flexible
' button behavior management through the DynamicButtonHandlers collection.
    
    Debug.Print "=== AttachButtonHandler START ==="
    On Error GoTo ErrorHandler
    
    If DynamicButtonHandlers Is Nothing Then
        Debug.Print "Creating new DynamicButtonHandlers collection"
        Set DynamicButtonHandlers = New Collection
    End If
    
    Dim ButtonHandler As clsDynamicButtonHandler
    Set ButtonHandler = New clsDynamicButtonHandler
    ButtonHandler.Initialize Button, Role, Me
    DynamicButtonHandlers.Add ButtonHandler
    Debug.Print "ButtonHandler added to collection for role: " & Role
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AttachButtonHandler: " & Err.Description
    Resume Next
End Sub

Private Sub RefreshCellFormatListBox()
' Reloads the list box with current cell format data from the ModCellFormat module.
' Ensures the UI displays the most up-to-date format information.
    
    CellFormatListBox.Clear
    
    Dim formats() As clsCellFormatType
    formats = ModCellFormat.GetCellFormatList()
    
    Dim i As Integer
    For i = LBound(formats) To UBound(formats)
        Debug.Print "Adding cell format to ListBox [" & i & "]: " & formats(i).Name
        CellFormatListBox.AddItem formats(i).Name
    Next i
    
    If CellFormatListBox.ListCount > 0 Then
        CellFormatListBox.ListIndex = 0
    End If
End Sub


Public Sub btnAdd_Click()
' Creates a new cell format with default properties.
    
    Debug.Print "=== btnAdd_Click started ==="
    Dim newName As String
    Dim newFormat As clsCellFormatType

    newName = InputBox("Enter the name for the new cell format:", "Add Format")
    Debug.Print "New name entered: " & newName
    If newName <> "" Then
        Set newFormat = New clsCellFormatType
        With newFormat
            .Name = newName
            .BackColor = RGB(255, 255, 255)
            .BorderStyle = xlContinuous
            .BorderColor = RGB(0, 0, 0)
            .FillPattern = xlSolid
            .FontStyle = 0  ' Normal
            .FontColor = RGB(0, 0, 0)
        End With

        Debug.Print "Calling ModCellFormat.AddFormat"
        ModCellFormat.AddFormat newFormat
        Debug.Print "Refreshing ListBox"
        RefreshCellFormatListBox
    End If
    Debug.Print "=== btnAdd_Click completed ==="
End Sub

Public Sub btnRemove_Click()
' Deletes the selected cell format from the collection and updates the UI accordingly.
    
    Debug.Print "=== btnRemove_Click started ==="
    Debug.Print "ListBox Index: " & CellFormatListBox.ListIndex
    If CellFormatListBox.ListIndex >= 0 Then
        ModCellFormat.RemoveFormat CellFormatListBox.ListIndex
        RefreshCellFormatListBox
    End If
    Debug.Print "=== btnRemove_Click completed ==="
End Sub

Public Sub btnSave_Click()
' Updates the selected cell format with current UI values and saves changes to the workbook.
    
    Debug.Print "=== btnSave_Click started ==="
    Debug.Print "ListBox Index: " & CellFormatListBox.ListIndex
    If CellFormatListBox.ListIndex >= 0 Then
        Dim selectedIndex As Integer
        selectedIndex = CellFormatListBox.ListIndex
        
        Dim updatedFormat As New clsCellFormatType
        With updatedFormat
            .Name = txtCellName.Text
            .BorderStyle = ModCellFormat.GetBorderStyleValue(cboBorderStyle.Text)
            .BorderColor = btnBorderColor.BackColor
            .FillPattern = ModCellFormat.GetFillPatternValue(cboFillPattern.Text)
            .BackColor = btnFillColor.BackColor
            .FontStyle = ModCellFormat.GetFontStyleValue(cboFontStyle.Text)
            .FontColor = btnFontColor.BackColor
        End With
        
        Debug.Print "Calling ModCellFormat.UpdateFormat"
        ModCellFormat.UpdateFormat selectedIndex, updatedFormat
        Debug.Print "Calling SaveCellFormatsToWorkbook"
        ModCellFormat.SaveCellFormatsToWorkbook
        
        RefreshCellFormatListBox
        CellFormatListBox.ListIndex = selectedIndex
    End If
    Debug.Print "=== btnSave_Click completed ==="
End Sub

Private Function ShowColorDialog() As Long
' Displays Excel's color picker dialog and returns the selected color value.
' Handles all color dialog initialization and error checking.
    
    On Error Resume Next
    
    ' Default return value
    ShowColorDialog = -1
    
    Dim testDialog As Object
    Set testDialog = Application.Dialogs(xlDialogEditColor)
    If Err.Number <> 0 Then
        Debug.Print "Error getting dialog: " & Err.Description
        Exit Function
    End If
    
    ' Store current workbook colors(1) to restore later
    Dim originalColor As Long
    originalColor = ActiveWorkbook.Colors(1)
    If Err.Number <> 0 Then
        Debug.Print "Error accessing workbook colors: " & Err.Description
        Exit Function
    End If
    
    ' Test basic dialog show without parameters first
    Dim success As Boolean
    success = Application.Dialogs(xlDialogEditColor).Show(1)
    If Err.Number <> 0 Then
        Debug.Print "Error showing dialog: " & Err.Description & " (Error " & Err.Number & ")"
        Exit Function
    End If
    
    If success Then
        ShowColorDialog = ActiveWorkbook.Colors(1)
    End If
    
    On Error GoTo 0
End Function

Public Sub btnBackColor_Click()
' Opens color picker and updates the button's background color with the selected color.

    Debug.Print "btnBackColor_Click triggered"
    Dim selectedColor As Long
    selectedColor = ShowColorDialog
    If selectedColor <> -1 Then
        btnBackColor.BackColor = selectedColor
    End If
End Sub

Public Sub btnBorderColor_Click()
' Handles the border color button click event. Opens color picker and
' updates the button's border color with the selected color.

    Debug.Print "btnBorderColor_Click triggered"
    Dim selectedColor As Long
    selectedColor = ShowColorDialog
    If selectedColor <> -1 Then
        btnBorderColor.BackColor = selectedColor
    End If
End Sub

Public Sub btnFillColor_Click()
    Debug.Print "btnFillColor_Click triggered"
    Dim selectedColor As Long
    selectedColor = ShowColorDialog
    If selectedColor <> -1 Then
        btnFillColor.BackColor = selectedColor
    End If
End Sub

Public Sub btnFontColor_Click()
    Debug.Print "btnFontColor_Click triggered"
    Dim selectedColor As Long
    selectedColor = ShowColorDialog
    If selectedColor <> -1 Then
        btnFontColor.BackColor = selectedColor
    End If
End Sub

