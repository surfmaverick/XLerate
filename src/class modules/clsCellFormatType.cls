' =============================================================================
' File: clsCellFormatType.cls
' Version: 2.0.0
' Description: Cell format type class for comprehensive cell formatting
' Author: XLerate Development Team
' Created: Enhanced for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCellFormatType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' Private properties
Private pName As String
Private pBackColor As Long
Private pBorderStyle As Long
Private pBorderColor As Long
Private pFillPattern As Long
Private pFontStyle As Long
Private pFontColor As Long

' Property accessors
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(Value As String)
    pName = Value
End Property

Public Property Get BackColor() As Long
    BackColor = pBackColor
End Property

Public Property Let BackColor(Value As Long)
    pBackColor = Value
End Property

Public Property Get BorderStyle() As Long
    BorderStyle = pBorderStyle
End Property

Public Property Let BorderStyle(Value As Long)
    pBorderStyle = Value
End Property

Public Property Get BorderColor() As Long
    BorderColor = pBorderColor
End Property

Public Property Let BorderColor(Value As Long)
    pBorderColor = Value
End Property

Public Property Get FillPattern() As Long
    FillPattern = pFillPattern
End Property

Public Property Let FillPattern(Value As Long)
    pFillPattern = Value
End Property

Public Property Get FontStyle() As Long
    FontStyle = pFontStyle
End Property

Public Property Let FontStyle(Value As Long)
    pFontStyle = Value
End Property

Public Property Get FontColor() As Long
    FontColor = pFontColor
End Property

Public Property Let FontColor(Value As Long)
    pFontColor = Value
End Property

Private Sub Class_Initialize()
    ' Set default values
    pName = "Default"
    pBackColor = RGB(255, 255, 255)    ' White
    pBorderStyle = xlContinuous
    pBorderColor = RGB(0, 0, 0)        ' Black
    pFillPattern = xlSolid
    pFontStyle = 0                     ' Normal
    pFontColor = RGB(0, 0, 0)          ' Black
End Sub

Public Function GetDisplayName() As String
    ' Return formatted display name for UI
    GetDisplayName = pName
End Function

Public Function IsValid() As Boolean
    ' Validate the format
    IsValid = (Len(pName) > 0)
End Function

Public Sub CopyFrom(sourceFormat As clsCellFormatType)
    ' Copy properties from another format
    pName = sourceFormat.Name
    pBackColor = sourceFormat.BackColor
    pBorderStyle = sourceFormat.BorderStyle
    pBorderColor = sourceFormat.BorderColor
    pFillPattern = sourceFormat.FillPattern
    pFontStyle = sourceFormat.FontStyle
    pFontColor = sourceFormat.FontColor
End Sub

Public Function Clone() As clsCellFormatType
    ' Create a copy of this format
    Dim newFormat As New clsCellFormatType
    newFormat.CopyFrom Me
    Set Clone = newFormat
End Function

Public Sub ApplyToRange(targetRange As Range)
    ' Apply this format to a range
    On Error Resume Next
    
    With targetRange
        .Interior.Pattern = pFillPattern
        .Interior.Color = pBackColor
        .Font.Color = pFontColor
        
        ' Apply font styles
        .Font.Bold = CBool(pFontStyle And 1)           ' FONT_BOLD = 1
        .Font.Italic = CBool(pFontStyle And 2)         ' FONT_ITALIC = 2
        .Font.Underline = CBool(pFontStyle And 4)      ' FONT_UNDERLINE = 4
        .Font.Strikethrough = CBool(pFontStyle And 8)  ' FONT_STRIKETHROUGH = 8
        
        ' Apply borders
        ApplyBordersToRange targetRange
    End With
    
    On Error GoTo 0
End Sub

Private Sub ApplyBordersToRange(targetRange As Range)
    ' Apply border formatting to a range
    On Error Resume Next
    
    Dim edges As Variant
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    
    Dim edge As Variant
    For Each edge In edges
        With targetRange.Borders(edge)
            .LineStyle = pBorderStyle
            If pBorderStyle <> xlNone Then
                .Color = pBorderColor
            End If
        End With
    Next edge
    
    ' Apply inside borders if range has multiple cells
    If targetRange.Cells.Count > 1 Then
        With targetRange.Borders(xlInsideHorizontal)
            .LineStyle = pBorderStyle
            If pBorderStyle <> xlNone Then
                .Color = pBorderColor
            End If
        End With
        
        With targetRange.Borders(xlInsideVertical)
            .LineStyle = pBorderStyle
            If pBorderStyle <> xlNone Then
                .Color = pBorderColor
            End If
        End With
    End If
    
    On Error GoTo 0
End Sub

Public Function GetColorName(colorValue As Long) As String
    ' Get a friendly name for common colors
    Select Case colorValue
        Case RGB(255, 255, 255): GetColorName = "White"
        Case RGB(0, 0, 0): GetColorName = "Black"
        Case RGB(255, 0, 0): GetColorName = "Red"
        Case RGB(0, 255, 0): GetColorName = "Green"
        Case RGB(0, 0, 255): GetColorName = "Blue"
        Case RGB(255, 255, 0): GetColorName = "Yellow"
        Case RGB(255, 0, 255): GetColorName = "Magenta"
        Case RGB(0, 255, 255): GetColorName = "Cyan"
        Case Else: GetColorName = "Custom (" & colorValue & ")"
    End Select
End Function