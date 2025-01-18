VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTextStyle
   Caption         =   "Text Style Settings"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmTextStyle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTextStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' UI Controls
Private WithEvents StyleListBox As MSForms.ListBox
Private WithEvents txtStyleName As MSForms.TextBox
Private WithEvents cboFontName As MSForms.ComboBox
Private WithEvents txtFontSize As MSForms.TextBox
Private WithEvents chkBold As MSForms.CheckBox
Private WithEvents chkItalic As MSForms.CheckBox
Private WithEvents chkUnderline As MSForms.CheckBox
Private WithEvents btnFontColor As MSForms.CommandButton
Private WithEvents btnBackColor As MSForms.CommandButton
Private WithEvents cboBorderStyle As MSForms.ComboBox
Private WithEvents chkBorderTop As MSForms.CheckBox
Private WithEvents chkBorderBottom As MSForms.CheckBox
Private WithEvents chkBorderLeft As MSForms.CheckBox
Private WithEvents chkBorderRight As MSForms.CheckBox
Private WithEvents btnAdd As MSForms.CommandButton
Private WithEvents btnRemove As MSForms.CommandButton
Private WithEvents btnSave As MSForms.CommandButton
Private lblPreview As MSForms.Label
Private WithEvents cboBorderWeight As MSForms.ComboBox

Public Sub InitializeInPanel(parentFrame As MSForms.Frame)
    On Error GoTo ErrorHandler
    Debug.Print vbNewLine & "=== frmTextStyle.InitializeInPanel START ==="
    
    ' Initialize the text styles module
    Debug.Print "Calling ModTextStyle.InitializeTextStyles"
    ModTextStyle.InitializeTextStyles
    
    ' Create GUI controls
    Debug.Print "Creating GUI controls"
    CreateControls parentFrame
    
    ' Load initial data
    Debug.Print "Refreshing style list box"
    RefreshStyleListBox
    Debug.Print "Populating font combo box"
    PopulateFontComboBox
    
    ' Select first item if exists
    If StyleListBox.ListCount > 0 Then
        Debug.Print "Setting initial selection to first item"
        StyleListBox.ListIndex = 0
        UpdateControlsFromStyle 0
    End If
    
    Debug.Print "=== frmTextStyle.InitializeInPanel END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in frmTextStyle.InitializeInPanel: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Private Sub CreateControls(parentFrame As MSForms.Frame)
    Dim top As Long: top = 10
    Dim labelWidth As Long: labelWidth = 120
    Dim controlWidth As Long: controlWidth = 120
    Dim height As Long: height = 20
    Dim spacing As Long: spacing = 25
    
    ' Style List
    Set StyleListBox = parentFrame.Controls.Add("Forms.ListBox.1", "StyleListBox")
    With StyleListBox
        .Left = 10
        .top = top
        .Width = 390
        .height = 150
    End With
    
    top = top + 160  ' Space after list box
    
    ' Style Name
    CreateLabel parentFrame, "Style Name:", 10, top
    Set txtStyleName = parentFrame.Controls.Add("Forms.TextBox.1", "txtStyleName")
    With txtStyleName
        .Left = 10
        .top = top + 20
        .Width = controlWidth
    End With
    
    ' Font Name
    CreateLabel parentFrame, "Font:", 140, top
    Set cboFontName = parentFrame.Controls.Add("Forms.ComboBox.1", "cboFontName")
    With cboFontName
        .Left = 140
        .top = top + 20
        .Width = controlWidth
    End With
    
    ' Font Size
    CreateLabel parentFrame, "Size:", 270, top
    Set txtFontSize = parentFrame.Controls.Add("Forms.TextBox.1", "txtFontSize")
    With txtFontSize
        .Left = 270
        .top = top + 20
        .Width = 60
    End With
    
    top = top + spacing + 20
    
    ' Font Styles
    Set chkBold = parentFrame.Controls.Add("Forms.CheckBox.1", "chkBold")
    With chkBold
        .Left = 10
        .top = top
        .Caption = "Bold"
        .Width = 60
    End With
    
    Set chkItalic = parentFrame.Controls.Add("Forms.CheckBox.1", "chkItalic")
    With chkItalic
        .Left = 80
        .top = top
        .Caption = "Italic"
        .Width = 60
    End With
    
    Set chkUnderline = parentFrame.Controls.Add("Forms.CheckBox.1", "chkUnderline")
    With chkUnderline
        .Left = 150
        .top = top
        .Caption = "Underline"
        .Width = 80
    End With
    
    top = top + spacing
    
    ' Colors (all in one row)
    CreateLabel parentFrame, "Colors:", 10, top
    Set btnFontColor = parentFrame.Controls.Add("Forms.CommandButton.1", "btnFontColor")
    With btnFontColor
        .Left = 80  ' Moved right to be next to "Colors:" label
        .top = top - 3  ' Slight adjustment to align with label
        .Width = 90
        .Caption = "Font Color"
    End With
    
    Set btnBackColor = parentFrame.Controls.Add("Forms.CommandButton.1", "btnBackColor")
    With btnBackColor
        .Left = 180  ' Positioned next to Font Color button
        .top = top - 3  ' Slight adjustment to align with label
        .Width = 90
        .Caption = "Back Color"
    End With
    
    top = top + spacing  ' Only add one spacing since we're not using an extra line
    
    ' Border Style and Weight (side by side)
    CreateLabel parentFrame, "Border Style:", 10, top
    CreateLabel parentFrame, "Border Weight:", 140, top
    
    Set cboBorderStyle = parentFrame.Controls.Add("Forms.ComboBox.1", "cboBorderStyle")
    With cboBorderStyle
        .Left = 10
        .top = top + 20
        .Width = controlWidth
        .Style = fmStyleDropDownList
        .AddItem "None"
        .AddItem "Continuous"
        .AddItem "Double"
        .AddItem "Dash"
        .AddItem "Dot"
        .ListIndex = 0
    End With
    
    Set cboBorderWeight = parentFrame.Controls.Add("Forms.ComboBox.1", "cboBorderWeight")
    With cboBorderWeight
        .Left = 140
        .top = top + 20
        .Width = controlWidth
        .Style = fmStyleDropDownList
        .AddItem "Hairline"
        .AddItem "Thin"
        .AddItem "Medium"
        .AddItem "Thick"
        .ListIndex = 1
    End With
    
    top = top + spacing + 20
    
    ' Border Position (in one row)
    CreateLabel parentFrame, "Border Position:", 10, top
    
    Set chkBorderTop = parentFrame.Controls.Add("Forms.CheckBox.1", "chkBorderTop")
    With chkBorderTop
        .Left = 10
        .top = top + 20
        .Width = 60
        .Caption = "Top"
    End With
    
    Set chkBorderBottom = parentFrame.Controls.Add("Forms.CheckBox.1", "chkBorderBottom")
    With chkBorderBottom
        .Left = 80
        .top = top + 20
        .Width = 60
        .Caption = "Bottom"
    End With
    
    Set chkBorderLeft = parentFrame.Controls.Add("Forms.CheckBox.1", "chkBorderLeft")
    With chkBorderLeft
        .Left = 150
        .top = top + 20
        .Width = 60
        .Caption = "Left"
    End With
    
    Set chkBorderRight = parentFrame.Controls.Add("Forms.CheckBox.1", "chkBorderRight")
    With chkBorderRight
        .Left = 220
        .top = top + 20
        .Width = 60
        .Caption = "Right"
    End With
    
    top = top + spacing + 20
    
    ' Preview Label
    Set lblPreview = parentFrame.Controls.Add("Forms.Label.1", "lblPreview")
    With lblPreview
        .Left = 10
        .top = top
        .Width = 390
        .height = 40
        .BorderStyle = 1  ' 1 = Single border
        .Caption = "Preview Text"
        .TextAlign = fmTextAlignCenter
    End With
    
    top = top + 50
    
    ' Buttons row
    Set btnAdd = parentFrame.Controls.Add("Forms.CommandButton.1", "btnAdd")
    With btnAdd
        .Left = 10
        .top = top
        .Width = 90
        .Caption = "Add Style"
    End With
    
    Set btnRemove = parentFrame.Controls.Add("Forms.CommandButton.1", "btnRemove")
    With btnRemove
        .Left = 110
        .top = top
        .Width = 90
        .Caption = "Remove"
    End With
    
    Set btnSave = parentFrame.Controls.Add("Forms.CommandButton.1", "btnSave")
    With btnSave
        .Left = 210
        .top = top
        .Width = 90
        .Caption = "Save"
    End With
End Sub

Private Sub CreateLabel(parentFrame As MSForms.Frame, Caption As String, Left As Long, top As Long)
    Dim lbl As MSForms.Label
    Set lbl = parentFrame.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = Caption
        .Left = Left
        .top = top
        .AutoSize = True
    End With
End Sub

' Helper functions and event handlers to be continued...

Private Sub RefreshStyleListBox()
    On Error GoTo ErrorHandler
    Debug.Print vbNewLine & "=== RefreshStyleListBox START ==="
    
    Debug.Print "Clearing style list box"
    StyleListBox.Clear
    
    Debug.Print "Getting text style list"
    Dim styles() As clsTextStyleType
    styles = ModTextStyle.GetTextStyleList()
    
    Debug.Print "Adding styles to list box"
    Dim i As Integer
    For i = LBound(styles) To UBound(styles)
        Debug.Print "Adding style: " & styles(i).Name
        StyleListBox.AddItem styles(i).Name
    Next i
    
    Debug.Print "=== RefreshStyleListBox END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in RefreshStyleListBox: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Private Sub PopulateFontComboBox()
    ' Get system fonts
    Dim fontCount As Long
    fontCount = Application.FontNames.Count
    
    Dim i As Long
    For i = 1 To fontCount
        cboFontName.AddItem Application.FontNames(i)
    Next i
End Sub

Private Function GetBorderStyleIndex(excelBorderStyle As Long) As Integer
    Select Case excelBorderStyle
        Case xlNone
            GetBorderStyleIndex = 0
        Case xlContinuous
            GetBorderStyleIndex = 1
        Case xlDouble
            GetBorderStyleIndex = 2
        Case xlDash
            GetBorderStyleIndex = 3
        Case xlDot
            GetBorderStyleIndex = 4
        Case Else
            GetBorderStyleIndex = 0
    End Select
End Function

Private Function GetExcelBorderStyle(index As Integer) As XlLineStyle
    Debug.Print "GetExcelBorderStyle called with index: " & index
    
    Dim result As XlLineStyle
    Select Case index
        Case 0: result = xlLineStyleNone  ' None (-4142)
        Case 1: result = xlContinuous     ' Continuous (1)
        Case 2: result = xlDouble         ' Double (-4119)
        Case 3: result = xlDash           ' Dash (-4115)
        Case 4: result = xlDot            ' Dot (-4118)
        Case Else: result = xlLineStyleNone
    End Select
    
    Debug.Print "Converting border style:"
    Debug.Print "  Input Index: " & index
    Debug.Print "  Excel Constant: " & result
    Debug.Print "  Constant Name: " & Choose(index + 1, "xlLineStyleNone", "xlContinuous", "xlDouble", "xlDash", "xlDot")
    
    GetExcelBorderStyle = result
End Function

Private Function GetExcelBorderWeight(index As Integer) As XlBorderWeight
    Select Case index
        Case 0: GetExcelBorderWeight = xlHairline
        Case 1: GetExcelBorderWeight = xlThin
        Case 2: GetExcelBorderWeight = xlMedium
        Case 3: GetExcelBorderWeight = xlThick
        Case Else: GetExcelBorderWeight = xlThin
    End Select
End Function

Private Sub UpdateControlsFromStyle(index As Integer)
    On Error GoTo ErrorHandler
    Debug.Print "=== UpdateControlsFromStyle START ==="
    
    Dim styles() As clsTextStyleType
    styles = ModTextStyle.GetTextStyleList()
    
    With styles(index)
        Debug.Print "Updating controls for style: " & .Name
        txtStyleName.Text = .Name
        cboFontName.Text = .FontName
        txtFontSize.Text = CStr(.FontSize)
        chkBold.value = .Bold
        chkItalic.value = .Italic
        chkUnderline.value = .Underline
        btnFontColor.BackColor = .FontColor
        btnBackColor.BackColor = .BackColor
        cboBorderStyle.ListIndex = GetBorderStyleIndex(.BorderStyle)
        chkBorderTop.value = .BorderTop
        chkBorderBottom.value = .BorderBottom
        chkBorderLeft.value = .BorderLeft
        chkBorderRight.value = .BorderRight
        
        ' Set border weight based on stored value
        Select Case .BorderWeight
            Case xlHairline: cboBorderWeight.ListIndex = 0
            Case xlThin: cboBorderWeight.ListIndex = 1
            Case xlMedium: cboBorderWeight.ListIndex = 2
            Case xlThick: cboBorderWeight.ListIndex = 3
            Case Else: cboBorderWeight.ListIndex = 1
        End Select
        
        ' Update preview
        UpdatePreview
    End With
    
    Debug.Print "=== UpdateControlsFromStyle END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in UpdateControlsFromStyle: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Private Sub UpdatePreview()
    With lblPreview
        .Font.Name = cboFontName.Text
        .Font.Size = Val(txtFontSize.Text)
        .Font.Bold = chkBold.value
        .Font.Italic = chkItalic.value
        .Font.Underline = chkUnderline.value
        .ForeColor = btnFontColor.BackColor
        .BackColor = btnBackColor.BackColor
    End With
End Sub

Private Function ShowColorDialog() As Long
    On Error Resume Next
    ShowColorDialog = -1
    
    ' Store current workbook colors(1) to restore later
    Dim originalColor As Long
    originalColor = ActiveWorkbook.Colors(1)
    
    ' Show color picker
    If Application.Dialogs(xlDialogEditColor).Show(1) Then
        ShowColorDialog = ActiveWorkbook.Colors(1)
    End If
    
    ' Restore original color
    ActiveWorkbook.Colors(1) = originalColor
    On Error GoTo 0
End Function

' Event handlers
Private Sub StyleListBox_Click()
    Debug.Print "StyleListBox_Click triggered"
    If StyleListBox.ListIndex >= 0 Then
        UpdateControlsFromStyle StyleListBox.ListIndex
    End If
End Sub

Private Sub btnAdd_Click()
    Debug.Print "btnAdd_Click triggered"
    Dim newName As String
    newName = InputBox("Enter name for new style:", "Add Style")
    
    If newName <> "" Then
        Dim newStyle As New clsTextStyleType
        With newStyle
            .Name = newName
            .FontName = "Calibri"
            .FontSize = 11
            .Bold = False
            .Italic = False
            .Underline = False
            .FontColor = RGB(0, 0, 0)
            .BackColor = RGB(255, 255, 255)
            .BorderStyle = xlContinuous
            .BorderTop = False
            .BorderBottom = False
            .BorderLeft = False
            .BorderRight = False
        End With
        
        ModTextStyle.AddStyle newStyle
        RefreshStyleListBox
        StyleListBox.ListIndex = StyleListBox.ListCount - 1
    End If
End Sub

Private Sub btnRemove_Click()
    Debug.Print "btnRemove_Click triggered"
    If StyleListBox.ListIndex >= 0 Then
        If MsgBox("Are you sure you want to remove this style?", _
                 vbQuestion + vbYesNo) = vbYes Then
            ModTextStyle.RemoveStyle StyleListBox.ListIndex
            RefreshStyleListBox
            If StyleListBox.ListCount > 0 Then StyleListBox.ListIndex = 0
        End If
    End If
End Sub

Private Sub btnSave_Click()
    Debug.Print "btnSave_Click triggered"
    If StyleListBox.ListIndex >= 0 Then
        Dim updatedStyle As New clsTextStyleType
        
        ' Get the Excel border style and print debug info
        Dim borderStyle As XlLineStyle
        borderStyle = GetExcelBorderStyle(cboBorderStyle.ListIndex)
        Debug.Print "Border Style Selection:"
        Debug.Print "  ComboBox Index: " & cboBorderStyle.ListIndex
        Debug.Print "  ComboBox Text: " & cboBorderStyle.Text
        Debug.Print "  Excel Style Constant: " & borderStyle
        
        With updatedStyle
            .Name = txtStyleName.Text
            .FontName = cboFontName.Text
            .FontSize = Val(txtFontSize.Text)
            .Bold = chkBold.value
            .Italic = chkItalic.value
            .Underline = chkUnderline.value
            .FontColor = btnFontColor.BackColor
            .BackColor = btnBackColor.BackColor
            .BorderStyle = borderStyle
            .BorderTop = chkBorderTop.value
            .BorderBottom = chkBorderBottom.value
            .BorderLeft = chkBorderLeft.value
            .BorderRight = chkBorderRight.value
            
            Debug.Print "Style Properties:"
            Debug.Print "  Name: " & .Name
            Debug.Print "  BorderStyle: " & .BorderStyle
            Debug.Print "  Borders (T/B/L/R): " & .BorderTop & "/" & .BorderBottom & "/" & .BorderLeft & "/" & .BorderRight
        End With
        
        ModTextStyle.UpdateStyle StyleListBox.ListIndex, updatedStyle
        ModTextStyle.SaveTextStylesToWorkbook
        RefreshStyleListBox
        StyleListBox.ListIndex = StyleListBox.ListIndex
    End If
End Sub

Private Sub btnFontColor_Click()
    Debug.Print "btnFontColor_Click triggered"
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnFontColor.BackColor = newColor
        UpdatePreview
    End If
End Sub

Private Sub btnBackColor_Click()
    Debug.Print "btnBackColor_Click triggered"
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnBackColor.BackColor = newColor
        UpdatePreview
    End If
End Sub

' Control change events for preview updates
Private Sub cboFontName_Change()
    Debug.Print "cboFontName_Change triggered"
    UpdatePreview
End Sub

Private Sub txtFontSize_Change()
    Debug.Print "txtFontSize_Change triggered"
    UpdatePreview
End Sub

Private Sub chkBold_Click()
    Debug.Print "chkBold_Click triggered"
    UpdatePreview
End Sub

Private Sub chkItalic_Click()
    Debug.Print "chkItalic_Click triggered"
    UpdatePreview
End Sub

Private Sub chkUnderline_Click()
    Debug.Print "chkUnderline_Click triggered"
    UpdatePreview
End Sub

