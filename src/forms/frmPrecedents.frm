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
Option Explicit

' Form-level dimension constants
Private Const FORM_PADDING As Long = 5        ' Consistent padding around all edges

' Column width constants
Private Const COLUMN_WIDTH_ADDRESS As Long = 125
Private Const COLUMN_WIDTH_VALUE As Long = 50
Private Const COLUMN_WIDTH_FORMULA As Long = 125

' Control spacing constants
Private Const CONTROL_SPACING As Long = 15     ' Vertical space between controls
Private Const HEADER_OFFSET As Long = 12       ' Space between headers and listbox
Private Const SCROLLBAR_WIDTH As Long = 20     ' Width of vertical scrollbar

Private Const INNER_WIDTH As Long = COLUMN_WIDTH_ADDRESS + COLUMN_WIDTH_VALUE + COLUMN_WIDTH_FORMULA + SCROLLBAR_WIDTH  ' Width of content
Private Const FORM_WIDTH As Long = INNER_WIDTH + (4 * FORM_PADDING)  ' Total form width including padding
Private Const FORM_HEIGHT As Long = 250       ' Total form height

' Formula box constants
Private Const FORMULA_HEIGHT As Long = 20
Private Const FORMULA_TOP As Long = FORM_PADDING
Private Const FORMULA_WIDTH As Long = INNER_WIDTH  ' Width matches the listbox content

' ListBox constants
Private Const LISTBOX_TOP As Long = FORMULA_TOP + FORMULA_HEIGHT + CONTROL_SPACING
Private Const LISTBOX_WIDTH As Long = INNER_WIDTH
Private Const LISTBOX_HEIGHT As Long = FORM_HEIGHT - LISTBOX_TOP - FORM_PADDING  ' Account for bottom padding


' Function to get column widths string
Public Function GetColumnWidths() As String
    GetColumnWidths = COLUMN_WIDTH_ADDRESS & ";" & _
                     COLUMN_WIDTH_VALUE & ";" & _
                     COLUMN_WIDTH_FORMULA
End Function

Public Sub AddHeaders()
    With lstPrecedents
        ' Add headers using a Label control for each column
        Dim headerLabel1 As MSForms.Label
        Set headerLabel1 = Me.Controls.Add("Forms.Label.1", "lblHeader1")
        With headerLabel1
            .Top = lstPrecedents.Top - HEADER_OFFSET
            .Left = lstPrecedents.Left + FORM_PADDING
            .Caption = "Address"
            .Width = COLUMN_WIDTH_ADDRESS
        End With
        
        Dim headerLabel2 As MSForms.Label
        Set headerLabel2 = Me.Controls.Add("Forms.Label.1", "lblHeader2")
        With headerLabel2
            .Top = lstPrecedents.Top - HEADER_OFFSET
            .Left = lstPrecedents.Left + COLUMN_WIDTH_ADDRESS + FORM_PADDING
            .Caption = "Value"
            .Width = COLUMN_WIDTH_VALUE
        End With
        
        Dim headerLabel3 As MSForms.Label
        Set headerLabel3 = Me.Controls.Add("Forms.Label.1", "lblHeader3")
        With headerLabel3
            .Top = lstPrecedents.Top - HEADER_OFFSET
            .Left = lstPrecedents.Left + COLUMN_WIDTH_ADDRESS + COLUMN_WIDTH_VALUE + FORM_PADDING
            .Caption = "Formula"
            .Width = COLUMN_WIDTH_FORMULA
        End With
    End With
End Sub

Private Sub UserForm_Initialize()
    ' Set form caption and size
    Me.Caption = "Trace Precedents"
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT + (4 * FORM_PADDING) - 2
    
    ' Initialize the formula text box
    With Me.Controls.Add("Forms.TextBox.1", "txtFormula")
        .Top = FORMULA_TOP
        .Left = FORM_PADDING
        .Width = FORMULA_WIDTH
        .Height = FORMULA_HEIGHT
        .BackColor = RGB(240, 240, 240)
        .Locked = True
        .MultiLine = True
        .Font.Size = 10
    End With
    
    ' Initialize the list box with adjusted positioning
    With lstPrecedents
        .Top = LISTBOX_TOP
        .Left = FORM_PADDING
        .Width = LISTBOX_WIDTH
        .Height = LISTBOX_HEIGHT
        .ColumnCount = 3
        .ColumnWidths = GetColumnWidths()
        .Font.Size = 10
    End With
End Sub

Private Sub lstPrecedents_Click()
    On Error Resume Next
    If lstPrecedents.ListIndex < 0 Then Exit Sub
    
    Dim precedentAddress As String
    precedentAddress = lstPrecedents.List(lstPrecedents.ListIndex, 0) ' Get first column value
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
            sheetName = Left(sheetName, Len(sheetName) - 1)
        End If
        
        ' Activate the Sheet and select the cell
        Worksheets(sheetName).Activate
        With Worksheets(sheetName).Range(cellAddress)
            .Select
            ' Update formula display
            Me.Controls("txtFormula").Text = .Formula
        End With
    Else
        ' Remove any indentation before trying to use the address
        precedentAddress = Trim(precedentAddress)
        Range(precedentAddress).Select
        Me.Controls("txtFormula").Text = Selection.Formula
    End If
    On Error GoTo 0
End Sub


Private Sub lstPrecedents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub



