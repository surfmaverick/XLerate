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
' frmSettingsManager
Option Explicit
' Control declarations
Private NumbersPanel As MSForms.Frame
Private CellsPanel As MSForms.Frame
Private DatesPanel As MSForms.Frame
Private numberSettings As frmNumberSettings
Private WithEvents lstCategories As MSForms.ListBox
Private AutoColorPanel As MSForms.Frame
Private autoColorSettings As frmAutoColor


Private Sub UserForm_Initialize()
    Debug.Print "SettingsManager Initialize started"
    On Error GoTo ErrorHandler
    
    ' Initialize form layout
    InitializeFormLayout
    Debug.Print "Form layout initialized"
    
    ' Create navigation listbox with event handling
    Debug.Print "Creating categories listbox"
    Set lstCategories = Me.Controls.Add("Forms.ListBox.1", "lstCategories")
    With lstCategories
        .Left = 12
        .Top = 12
        .Width = 150
        .Height = 450
    End With
    Debug.Print "Categories list created"
    
    ' Create panels
    InitializePanels
    Debug.Print "Panels initialized"
    
    InitializeHierarchyList
    Debug.Print "Hierarchy list initialized"
    
    Debug.Print "SettingsManager Initialize completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in SettingsManager Initialize: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Private Sub InitializePanels()
    ' Create Numbers panel frame
    Set NumbersPanel = Me.Controls.Add("Forms.Frame.1", "NumbersPanel")
    With NumbersPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Create Cells panel frame
    Set CellsPanel = Me.Controls.Add("Forms.Frame.1", "CellsPanel")
    With CellsPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Create Dates panel frame
    Set DatesPanel = Me.Controls.Add("Forms.Frame.1", "DatesPanel")
    With DatesPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Create Auto-Color panel frame
    Set AutoColorPanel = Me.Controls.Add("Forms.Frame.1", "AutoColorPanel")
    With AutoColorPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Initialize all settings within their respective panels
    InitializeNumberSettings NumbersPanel
    InitializeCellSettings CellsPanel
    InitializeDateSettings DatesPanel
    InitializeAutoColorSettings AutoColorPanel
End Sub

Private Sub InitializeNumberSettings(parentFrame As MSForms.Frame)
    ' Create a new instance of frmNumberSettings
    Dim numberSettings As New frmNumberSettings
    ' Initialize it within the panel
    numberSettings.InitializeInPanel parentFrame
End Sub

Private Sub InitializeCellSettings(parentFrame As MSForms.Frame)
    ' Create a new instance of frmCellSettings
    Dim cellSettings As New frmCellSettings
    ' Initialize it within the panel
    cellSettings.InitializeInPanel parentFrame
End Sub

Private Sub InitializeDateSettings(parentFrame As MSForms.Frame)
    Debug.Print "Initializing date settings"
    Dim dateSettings As New frmDateSettings
    dateSettings.InitializeInPanel parentFrame
    Debug.Print "Date settings initialized"
End Sub

Private Sub InitializeAutoColorSettings(parentFrame As MSForms.Frame)
    Debug.Print "Initializing auto-color settings"
    Set autoColorSettings = New frmAutoColor
    autoColorSettings.InitializeInPanel parentFrame
    Debug.Print "Auto-color settings initialized"
End Sub

' Show the requested panel and hide others
Private Sub ShowPanel(panelName As String)
    Debug.Print vbNewLine & "=== ShowPanel called ==="
    Debug.Print "panelName: '" & panelName & "'"
    
    DebugPanelState
    
    ' Hide all panels first
    NumbersPanel.Visible = False
    DatesPanel.Visible = False
    CellsPanel.Visible = False
    AutoColorPanel.Visible = False
    
    Select Case panelName
        Case "Numbers"
            NumbersPanel.Visible = True
        Case "Dates"
            DatesPanel.Visible = True
        Case "Cells"
            CellsPanel.Visible = True
        Case "Auto-Color"
            AutoColorPanel.Visible = True
    End Select
    
    DebugPanelState
    Debug.Print "=== ShowPanel completed ==="
End Sub

Private Sub UserForm_Terminate()
    Set numberSettings = Nothing
    Set NumbersPanel = Nothing
    Set CellsPanel = Nothing
    Set DatesPanel = Nothing
    Set AutoColorPanel = Nothing
    Set autoColorSettings = Nothing
    Set lstCategories = Nothing
End Sub

' Add this event handler for the listbox
Private Sub lstCategories_Click()
    Debug.Print vbNewLine & "=== lstCategories_Click triggered ==="
    On Error GoTo ErrorHandler
    
    Dim selectedCategory As String
    selectedCategory = Trim(lstCategories.Text)
    Debug.Print "Selected category: '" & selectedCategory & "'"
    Debug.Print "Category length: " & Len(selectedCategory)
    Debug.Print "ASCII codes: "
    Dim i As Integer
    For i = 1 To Len(selectedCategory)
        Debug.Print "Position " & i & ": " & Asc(Mid(selectedCategory, i, 1))
    Next i
    
    If lstCategories.List(lstCategories.ListIndex, 1) = "HEADER" Then
        Debug.Print "Header clicked, selecting default item"
        lstCategories.ListIndex = 1
        Exit Sub
    End If
    
    Debug.Print "Processing category selection"
    ' Remove any potential hidden characters and extra spaces
    selectedCategory = Replace(selectedCategory, vbTab, "")
    selectedCategory = Replace(selectedCategory, vbCr, "")
    selectedCategory = Replace(selectedCategory, vbLf, "")
    selectedCategory = Trim(selectedCategory)
    
    Select Case selectedCategory
        Case "Numbers"
            ShowPanel "Numbers"
        Case "Cells"
            ShowPanel "Cells"
        Case "Dates"
            ShowPanel "Dates"
        Case "Auto-Color"
            ShowPanel "Auto-Color"
        Case Else
            Debug.Print "Unknown category selected"
            ShowPanel "None"
    End Select
    Debug.Print "=== lstCategories_Click completed ==="
    Exit Sub

ErrorHandler:
    Debug.Print "Error in lstCategories_Click: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub
Private Sub DebugPanelState()
    Debug.Print vbNewLine & "=== Panel State Debug ==="
    Debug.Print "NumbersPanel is Nothing: " & (NumbersPanel Is Nothing)
    If Not NumbersPanel Is Nothing Then Debug.Print "NumbersPanel.Visible: " & NumbersPanel.Visible
    
    Debug.Print "DatesPanel is Nothing: " & (DatesPanel Is Nothing)
    If Not DatesPanel Is Nothing Then Debug.Print "DatesPanel.Visible: " & DatesPanel.Visible
    
    Debug.Print "CellsPanel is Nothing: " & (CellsPanel Is Nothing)
    If Not CellsPanel Is Nothing Then Debug.Print "CellsPanel.Visible: " & CellsPanel.Visible
End Sub

Private Sub InitializeFormLayout()
    Me.BackColor = RGB(255, 255, 255)
    Me.Caption = "Settings"
    Me.Width = 600
    Me.Height = 500
    Debug.Print "Form layout set"
End Sub

Private Sub InitializeHierarchyList()
    Debug.Print "Initializing hierarchy list"
    lstCategories.Clear
    
    With lstCategories
        .AddItem "Formatting"
        .List(.ListCount - 1, 1) = "HEADER"
        .AddItem "Numbers"
        .AddItem "Dates"
        .AddItem "Cells"
        .AddItem "Auto-Color"
        .ListIndex = 1
    End With
    
    ShowPanel "Numbers"
End Sub

