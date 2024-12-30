Attribute VB_Name = "ModGlobalSettings"
Option Explicit

Private GlobalSettings As clsUISettings

Public Sub InitializeUISettings()
    If GlobalSettings Is Nothing Then
        Set GlobalSettings = New clsUISettings
    End If
    
    If Not LoadSettingsFromWorkbook() Then
        ' Use defaults from class
        SaveSettingsToWorkbook
    End If
End Sub

Public Function GetUISettings() As clsUISettings
    If GlobalSettings Is Nothing Then
        InitializeUISettings
    End If
    Set GetUISettings = GlobalSettings
End Function

Private Sub SaveSettingsToWorkbook()
    Dim propValue As String
    
    ' Serialize settings to string
    propValue = GlobalSettings.BackgroundColor & "|" & _
                GlobalSettings.FontName & "|" & _
                GlobalSettings.FontSize & "|" & _
                GlobalSettings.AccentColor
    
    ' Save to custom property
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties("UISettings").Delete
    On Error GoTo 0
    ThisWorkbook.CustomDocumentProperties.Add Name:="UISettings", _
        LinkToContent:=False, Type:=msoPropertyTypeString, value:=propValue
        
    ThisWorkbook.Save
End Sub

Private Function LoadSettingsFromWorkbook() As Boolean
    On Error Resume Next
    Dim propValue As String
    propValue = ThisWorkbook.CustomDocumentProperties("UISettings")
    On Error GoTo 0
    
    If propValue = "" Then
        LoadSettingsFromWorkbook = False
        Exit Function
    End If
    
    Dim parts() As String
    parts = Split(propValue, "|")
    
    If UBound(parts) = 3 Then
        GlobalSettings.BackgroundColor = CLng(parts(0))
        GlobalSettings.FontName = parts(1)
        GlobalSettings.FontSize = CInt(parts(2))
        GlobalSettings.AccentColor = CLng(parts(3))
        LoadSettingsFromWorkbook = True
    Else
        LoadSettingsFromWorkbook = False
    End If
End Function

