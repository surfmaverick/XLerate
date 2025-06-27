' =============================================================================
' File: clsTextStyleType.cls
' Version: 2.0.0
' Description: Text style type class for comprehensive text formatting
' Author: XLerate Development Team
' Created: Enhanced for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextStyleType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' Private member variables
Private m_Name As String
Private m_FontName As String
Private m_FontSize As Long
Private m_Bold As Boolean
Private m_Italic As Boolean
Private m_Underline As Boolean
Private m_FontColor As Long
Private m_BackColor As Long
Private m_BorderStyle As Long
Private m_BorderTop As Boolean
Private m_BorderBottom As Boolean
Private m_BorderLeft As Boolean
Private m_BorderRight As Boolean
Private m_BorderWeight As Long

Private Sub Class_Initialize()
    ' Set default values
    m_Name = "Default Style"
    m_FontName = "Calibri"
    m_FontSize = 11
    m_Bold = False
    m_Italic = False
    m_Underline = False
    m_FontColor = RGB(0, 0, 0)      ' Black
    m_BackColor = RGB(255, 255, 255) ' White
    m_BorderStyle = xlLineStyleNone
    m_BorderTop = False
    m_BorderBottom = False
    m_BorderLeft = False
    m_BorderRight = False
    m_BorderWeight = xlThin
End Sub

' Property accessors
Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get FontName() As String
    FontName = m_FontName
End Property

Public Property Let FontName(Value As String)
    m_FontName = Value
End Property

Public Property Get FontSize() As Long
    FontSize = m_FontSize
End Property

Public Property Let FontSize(Value As Long)
    If Value > 0 And Value <= 72 Then
        m_FontSize = Value
    End If
End Property

Public Property Get Bold() As Boolean
    Bold = m_Bold
End Property

Public Property Let Bold(Value As Boolean)
    m_Bold = Value
End Property

Public Property Get Italic() As Boolean
    Italic = m_Italic
End Property

Public Property Let Italic(Value As Boolean)
    m_Italic = Value
End Property

Public Property Get Underline() As Boolean
    Underline = m_Underline
End Property

Public Property Let Underline(Value As Boolean)
    m_Underline = Value
End Property

Public Property Get FontColor() As Long
    FontColor = m_FontColor
End Property

Public Property Let FontColor(Value As Long)
    m_FontColor = Value
End Property

Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property

Public Property Let BackColor(Value As Long)
    m_BackColor = Value
End Property

Public Property Get BorderStyle() As Long
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(Value As Long)
    m_BorderStyle = Value
End Property

Public Property Get BorderTop() As Boolean
    BorderTop = m_BorderTop
End Property

Public Property Let BorderTop(Value As Boolean)
    m_BorderTop = Value
End Property

Public Property Get BorderBottom() As Boolean
    BorderBottom = m_BorderBottom
End Property

Public Property Let BorderBottom(Value As Boolean)
    m_BorderBottom = Value
End Property

Public Property Get BorderLeft() As Boolean
    BorderLeft = m_BorderLeft
End Property

Public Property Let BorderLeft(Value As Boolean)
    m_BorderLeft = Value
End Property

Public Property Get BorderRight() As Boolean
    BorderRight = m_BorderRight
End Property

Public Property Let BorderRight(Value As Boolean)
    m_BorderRight = Value
End Property

Public Property Get BorderWeight() As Long
    BorderWeight = m_BorderWeight
End Property

Public Property Let BorderWeight(Value As Long)
    m_BorderWeight = Value
End Property

' Methods
Public Function GetDisplayName() As String
    ' Return formatted display name for UI
    GetDisplayName = m_Name & " (" & m_FontName & " " & m_FontSize & "pt)"
End Function

Public Function IsValid() As Boolean
    ' Validate the style
    IsValid = (Len(m_Name) > 0 And Len(m_FontName) > 0 And m_FontSize > 0)
End Function

Public Sub CopyFrom(sourceStyle As clsTextStyleType)
    ' Copy properties from another style
    m_Name = sourceStyle.Name
    m_FontName = sourceStyle.FontName
    m_FontSize = sourceStyle.FontSize
    m_Bold = sourceStyle.Bold
    m_Italic = sourceStyle.Italic
    m_Underline = sourceStyle.Underline
    m_FontColor = sourceStyle.FontColor
    m_BackColor = sourceStyle.BackColor
    m_BorderStyle = sourceStyle.BorderStyle
    m_BorderTop = sourceStyle.BorderTop
    m_BorderBottom = sourceStyle.BorderBottom
    m_BorderLeft = sourceStyle.BorderLeft
    m_BorderRight = sourceStyle.BorderRight
    m_BorderWeight = sourceStyle.BorderWeight
End Sub

Public Function Clone() As clsTextStyleType
    ' Create a copy of this style
    Dim newStyle As New clsTextStyleType
    newStyle.CopyFrom Me
    Set Clone = newStyle
End Function

Public Sub ApplyToRange(targetRange As Range)
    ' Apply this text style to a range
    On Error Resume Next
    
    With targetRange
        ' Apply font properties
        .Font.Name = m_FontName
        .Font.Size = m_FontSize
        .Font.Bold = m_Bold
        .Font.Italic = m_Italic
        .Font.Underline = m_Underline
        .Font.Color = m_FontColor
        
        ' Apply background color
        .Interior.Color = m_BackColor
        
        ' Clear existing borders first
        .Borders.LineStyle = xlNone
        
        ' Apply border style if not None
        If m_BorderStyle <> xlLineStyleNone Then
            ApplyBordersToRange targetRange
        End If
    End With
    
    On Error GoTo 0
End Sub

Public Function GetStyleDescription() As String
    ' Return a description of the style properties
    Dim desc As String
    desc = m_Name & ": " & m_FontName & " " & m_FontSize & "pt"
    
    If m_Bold Then desc = desc & ", Bold"
    If m_Italic Then desc = desc & ", Italic"
    If m_Underline Then desc = desc & ", Underline"
    
    If m_BorderTop Or m_BorderBottom Or m_BorderLeft Or m_BorderRight Then
        desc = desc & ", Borders"
    End If
    
    GetStyleDescription = desc
End Function

Public Function HasBorders() As Boolean
    ' Check if any borders are enabled
    HasBorders = (m_BorderTop Or m_BorderBottom Or m_BorderLeft Or m_BorderRight)
End Function

Public Sub SetAllBorders(enabled As Boolean)
    ' Enable or disable all borders
    m_BorderTop = enabled
    m_BorderBottom = enabled
    m_BorderLeft = enabled
    m_BorderRight = enabled
End Sub

Public Sub SetFontStyle(fontName As String, fontSize As Long, Optional bold As Boolean = False, _
                       Optional italic As Boolean = False, Optional underline As Boolean = False)
    ' Set multiple font properties at once
    m_FontName = fontName
    m_FontSize = fontSize
    m_Bold = bold
    m_Italic = italic
    m_Underline = underline
End Sub

Public Sub SetColors(fontColor As Long, Optional backColor As Long = -1)
    ' Set font and background colors
    m_FontColor = fontColor
    If backColor <> -1 Then
        m_BackColor = backColor
    End If
End Sub

Private Sub ApplyBordersToRange(targetRange As Range)
    ' Apply selective borders based on style settings
    On Error Resume Next
    
    If m_BorderTop Then
        With targetRange.Borders(xlEdgeTop)
            .LineStyle = m_BorderStyle
            .Weight = m_BorderWeight
        End With
    End If
    
    If m_BorderBottom Then
        With targetRange.Borders(xlEdgeBottom)
            .LineStyle = m_BorderStyle
            .Weight = m_BorderWeight
        End With
    End If
    
    If m_BorderLeft Then
        With targetRange.Borders(xlEdgeLeft)
            .LineStyle = m_BorderStyle
            .Weight = m_BorderWeight
        End With
    End If
    
    If m_BorderRight Then
        With targetRange.Borders(xlEdgeRight)
            .LineStyle = m_BorderStyle
            .Weight = m_BorderWeight
        End With
    End If