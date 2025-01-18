VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestStyleType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
    m_Name = ""
    m_FontName = "Calibri"
    m_FontSize = 11
    m_Bold = False
    m_Italic = False
    m_Underline = False
    m_FontColor = RGB(0, 0, 0)
    m_BackColor = RGB(255, 255, 255)
    m_BorderStyle = 0
    m_BorderTop = False
    m_BorderBottom = False
    m_BorderLeft = False
    m_BorderRight = False
    m_BorderWeight = xlThin  ' Default to thin
End Sub

' Name property
Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(value As String)
    m_Name = value
End Property

' FontName property
Public Property Get FontName() As String
    FontName = m_FontName
End Property
Public Property Let FontName(value As String)
    m_FontName = value
End Property

' FontSize property
Public Property Get FontSize() As Long
    FontSize = m_FontSize
End Property
Public Property Let FontSize(value As Long)
    m_FontSize = value
End Property

' Bold property
Public Property Get Bold() As Boolean
    Bold = m_Bold
End Property
Public Property Let Bold(value As Boolean)
    m_Bold = value
End Property

' Italic property
Public Property Get Italic() As Boolean
    Italic = m_Italic
End Property
Public Property Let Italic(value As Boolean)
    m_Italic = value
End Property

' Underline property
Public Property Get Underline() As Boolean
    Underline = m_Underline
End Property
Public Property Let Underline(value As Boolean)
    m_Underline = value
End Property

' FontColor property
Public Property Get FontColor() As Long
    FontColor = m_FontColor
End Property
Public Property Let FontColor(value As Long)
    m_FontColor = value
End Property

' BackColor property
Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property
Public Property Let BackColor(value As Long)
    m_BackColor = value
End Property

' BorderStyle property
Public Property Get BorderStyle() As Long
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(value As Long)
    m_BorderStyle = value
End Property

' Border position properties
Public Property Get BorderTop() As Boolean
    BorderTop = m_BorderTop
End Property
Public Property Let BorderTop(value As Boolean)
    m_BorderTop = value
End Property

Public Property Get BorderBottom() As Boolean
    BorderBottom = m_BorderBottom
End Property
Public Property Let BorderBottom(value As Boolean)
    m_BorderBottom = value
End Property

Public Property Get BorderLeft() As Boolean
    BorderLeft = m_BorderLeft
End Property
Public Property Let BorderLeft(value As Boolean)
    m_BorderLeft = value
End Property

Public Property Get BorderRight() As Boolean
    BorderRight = m_BorderRight
End Property
Public Property Let BorderRight(value As Boolean)
    m_BorderRight = value
End Property

' BorderWeight property
Public Property Get BorderWeight() As Long
    BorderWeight = m_BorderWeight
End Property

Public Property Let BorderWeight(value As Long)
    m_BorderWeight = value
End Property
