VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAutoColor
   Caption         =   "Auto-Color Settings"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmAutoColor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAutoColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NAME_PREFIX As String = "AutoColor_"

' UI Controls
Private WithEvents btnInputColor As MSForms.CommandButton
Private WithEvents btnFormulaColor As MSForms.CommandButton
Private WithEvents btnWorksheetLinkColor As MSForms.CommandButton
Private WithEvents btnWorkbookLinkColor As MSForms.CommandButton
Private WithEvents btnExternalColor As MSForms.CommandButton
Private WithEvents btnHyperlinkColor As MSForms.CommandButton
Private WithEvents btnPartialInputColor As MSForms.CommandButton
Private WithEvents btnResetDefaults As MSForms.CommandButton

Public Sub InitializeInPanel(parentFrame As MSForms.Frame)
    ' Create labels and color buttons
    CreateControls parentFrame
    
    ' Load saved colors
    LoadSavedColors
End Sub

Private Sub CreateControls(parentFrame As MSForms.Frame)
    Dim top As Long: top = 10
    Dim labelWidth As Long: labelWidth = 120
    Dim buttonWidth As Long: buttonWidth = 60
    Dim height As Long: height = 20
    Dim spacing As Long: spacing = 25
    
    ' Inputs
    CreateLabelAndButton parentFrame, "Inputs:", "btnInputColor", top, labelWidth, buttonWidth, height
    Set btnInputColor = parentFrame.Controls("btnInputColor")
    
    ' Formulas
    top = top + spacing
    CreateLabelAndButton parentFrame, "Formulas:", "btnFormulaColor", top, labelWidth, buttonWidth, height
    Set btnFormulaColor = parentFrame.Controls("btnFormulaColor")
    
    ' Worksheet Links
    top = top + spacing
    CreateLabelAndButton parentFrame, "Worksheet Links:", "btnWorksheetLinkColor", top, labelWidth, buttonWidth, height
    Set btnWorksheetLinkColor = parentFrame.Controls("btnWorksheetLinkColor")
    
    ' Workbook Links
    top = top + spacing
    CreateLabelAndButton parentFrame, "Workbook Links:", "btnWorkbookLinkColor", top, labelWidth, buttonWidth, height
    Set btnWorkbookLinkColor = parentFrame.Controls("btnWorkbookLinkColor")
    
    ' External References
    top = top + spacing
    CreateLabelAndButton parentFrame, "External References:", "btnExternalColor", top, labelWidth, buttonWidth, height
    Set btnExternalColor = parentFrame.Controls("btnExternalColor")
    
    ' Hyperlinks
    top = top + spacing
    CreateLabelAndButton parentFrame, "Hyperlinks:", "btnHyperlinkColor", top, labelWidth, buttonWidth, height
    Set btnHyperlinkColor = parentFrame.Controls("btnHyperlinkColor")
    
    ' Partial Inputs
    top = top + spacing
    CreateLabelAndButton parentFrame, "Partial Inputs:", "btnPartialInputColor", top, labelWidth, buttonWidth, height
    Set btnPartialInputColor = parentFrame.Controls("btnPartialInputColor")
    
    ' Add Reset Defaults button at the bottom
    Set btnResetDefaults = parentFrame.Controls.Add("Forms.CommandButton.1", "btnResetDefaults")
    With btnResetDefaults
        .Left = 10
        .Top = top + spacing * 8  ' Position below all other controls
        .Width = 120
        .Height = height
        .Caption = "Reset to Defaults"
    End With
End Sub

Private Sub CreateLabelAndButton(parentFrame As MSForms.Frame, labelText As String, buttonName As String, _
                               top As Long, labelWidth As Long, buttonWidth As Long, height As Long)
    ' Create label
    Dim lbl As MSForms.Label
    Set lbl = parentFrame.Controls.Add("Forms.Label.1")
    With lbl
        .Left = 10
        .Top = top
        .Width = labelWidth
        .Height = height
        .Caption = labelText
    End With
    
    ' Create color button
    Dim btn As MSForms.CommandButton
    Set btn = parentFrame.Controls.Add("Forms.CommandButton.1", buttonName)
    With btn
        .Left = labelWidth + 20
        .Top = top
        .Width = buttonWidth
        .Height = height
        .Caption = "Color"
    End With
End Sub

Private Sub LoadSavedColors()
    ' Load colors from Names or set defaults if not found
    btnInputColor.BackColor = GetSavedColor("Input", 16711680)         ' Blue
    btnFormulaColor.BackColor = GetSavedColor("Formula", 0)            ' Black
    btnWorksheetLinkColor.BackColor = GetSavedColor("WorksheetLink", 32768)     ' Green
    btnWorkbookLinkColor.BackColor = GetSavedColor("WorkbookLink", 16751052)    ' Light Purple
    btnExternalColor.BackColor = GetSavedColor("External", 15773696)   ' Light Blue
    btnHyperlinkColor.BackColor = GetSavedColor("Hyperlink", 33023)    ' Orange
    btnPartialInputColor.BackColor = GetSavedColor("PartialInput", 128)         ' Purple
End Sub

Private Function GetSavedColor(colorName As String, defaultColor As Long) As Long
    On Error Resume Next
    Dim colorValue As String
    colorValue = ThisWorkbook.Names(NAME_PREFIX & colorName).RefersTo
    If Err.Number = 0 And colorValue <> "" Then
        GetSavedColor = CLng(Mid(colorValue, 2)) ' Remove the = sign
    Else
        GetSavedColor = defaultColor
    End If
    On Error GoTo 0
End Function

Private Sub SaveColor(colorName As String, colorValue As Long)
    On Error Resume Next
    ThisWorkbook.Names.Add NAME_PREFIX & colorName, "=" & colorValue
    On Error GoTo 0
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

' Color button click events
Private Sub btnInputColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnInputColor.BackColor = newColor
        SaveColor "Input", newColor
    End If
End Sub

Private Sub btnFormulaColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnFormulaColor.BackColor = newColor
        SaveColor "Formula", newColor
    End If
End Sub

Private Sub btnWorksheetLinkColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnWorksheetLinkColor.BackColor = newColor
        SaveColor "WorksheetLink", newColor
    End If
End Sub

Private Sub btnWorkbookLinkColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnWorkbookLinkColor.BackColor = newColor
        SaveColor "WorkbookLink", newColor
    End If
End Sub

Private Sub btnExternalColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnExternalColor.BackColor = newColor
        SaveColor "External", newColor
    End If
End Sub

Private Sub btnHyperlinkColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnHyperlinkColor.BackColor = newColor
        SaveColor "Hyperlink", newColor
    End If
End Sub

Private Sub btnPartialInputColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnPartialInputColor.BackColor = newColor
        SaveColor "PartialInput", newColor
    End If
End Sub

' Add the reset button click handler
Private Sub btnResetDefaults_Click()
    If MsgBox("Are you sure you want to reset all colors to their defaults?", _
              vbQuestion + vbYesNo, "Reset Colors") = vbYes Then
              
        ' Reset all colors to defaults
        btnInputColor.BackColor = 16711680        ' Blue
        btnFormulaColor.BackColor = 0             ' Black
        btnWorksheetLinkColor.BackColor = 32768   ' Green
        btnWorkbookLinkColor.BackColor = 16751052 ' Light Purple
        btnExternalColor.BackColor = 15773696     ' Light Blue
        btnHyperlinkColor.BackColor = 33023       ' Orange
        btnPartialInputColor.BackColor = 128      ' Purple
        
        ' Save default colors
        SaveColor "Input", 16711680
        SaveColor "Formula", 0
        SaveColor "WorksheetLink", 32768
        SaveColor "WorkbookLink", 16751052
        SaveColor "External", 15773696
        SaveColor "Hyperlink", 33023
        SaveColor "PartialInput", 128
        
        MsgBox "Colors have been reset to defaults.", vbInformation
    End If
End Sub 