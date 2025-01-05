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

Private pName As String
Private pBackColor As Long
Private pBorderStyle As Long
Private pBorderColor As Long
Private pFillPattern As Long
Private pFontStyle As Long
Private pFontColor As Long

Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get BackColor() As Long
    BackColor = pBackColor
End Property

Public Property Let BackColor(value As Long)
    pBackColor = value
End Property

Public Property Get BorderStyle() As Long
    BorderStyle = pBorderStyle
End Property

Public Property Let BorderStyle(value As Long)
    pBorderStyle = value
End Property

Public Property Get BorderColor() As Long
    BorderColor = pBorderColor
End Property

Public Property Let BorderColor(value As Long)
    pBorderColor = value
End Property

Public Property Get FillPattern() As Long
    FillPattern = pFillPattern
End Property
Public Property Let FillPattern(value As Long)
    pFillPattern = value
End Property

Public Property Get FontStyle() As Long
    FontStyle = pFontStyle
End Property
Public Property Let FontStyle(value As Long)
    pFontStyle = value
End Property

Public Property Get FontColor() As Long
    FontColor = pFontColor
End Property
Public Property Let FontColor(value As Long)
    pFontColor = value
End Property

Private Sub Class_Initialize()
    ' Set default values
    pName = "Default"
    pBackColor = RGB(255, 255, 255)
    pBorderStyle = xlContinuous
    pBorderColor = RGB(0, 0, 0)
    pFillPattern = xlSolid
    pFontStyle = 0
    pFontColor = RGB(0, 0, 0)
End Sub
