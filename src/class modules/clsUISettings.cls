VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUISettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type UISettings
    BackgroundColor As Long
    FontName As String
    FontSize As Integer
    AccentColor As Long
End Type

Private Settings As UISettings

Public Property Get BackgroundColor() As Long
    BackgroundColor = Settings.BackgroundColor
End Property

Public Property Let BackgroundColor(ByVal value As Long)
    Settings.BackgroundColor = value
End Property

Public Property Get FontName() As String
    FontName = Settings.FontName
End Property

Public Property Let FontName(ByVal value As String)
    Settings.FontName = value
End Property

Public Property Get FontSize() As Integer
    FontSize = Settings.FontSize
End Property

Public Property Let FontSize(ByVal value As Integer)
    Settings.FontSize = value
End Property

Public Property Get AccentColor() As Long
    AccentColor = Settings.AccentColor
End Property

Public Property Let AccentColor(ByVal value As Long)
    Settings.AccentColor = value
End Property

Private Sub Class_Initialize()
    ' Set default values
    Settings.BackgroundColor = RGB(255, 255, 255) ' White
    Settings.FontName = "Segoe UI"
    Settings.FontSize = 9
    Settings.AccentColor = RGB(0, 120, 212) ' Blue accent
End Sub
