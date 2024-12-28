Attribute VB_Name = "ModNumberFormat"
Option Explicit

Private Const DEFAULT_FORMAT_1 As String = "General"
Private Const DEFAULT_FORMAT_2 As String = "#,##0.0;(#,##0.0);""-"""
Private Const DEFAULT_FORMAT_3 As String = "#,##0.00;(#,##0.00);""-"""

Private Type FormatSettings
    Format1 As String
    Format2 As String
    Format3 As String
    EnabledFormats As Integer  ' Bitmap of enabled formats
End Type

Private Settings As FormatSettings

Public Sub InitializeSettings()
    If Settings.Format1 = "" Then
        Settings.Format1 = DEFAULT_FORMAT_1
        Settings.Format2 = DEFAULT_FORMAT_2
        Settings.Format3 = DEFAULT_FORMAT_3
        Settings.EnabledFormats = 7 ' All formats enabled (111 in binary)
    End If
End Sub

Public Sub SaveSettings()
    With ThisWorkbook
        .CustomDocumentProperties("Format1") = Settings.Format1
        .CustomDocumentProperties("Format2") = Settings.Format2
        .CustomDocumentProperties("Format3") = Settings.Format3
        .CustomDocumentProperties("EnabledFormats") = Settings.EnabledFormats
    End With
End Sub

Public Sub LoadSettings()
    On Error Resume Next
    CreateCustomProperties
    
    With ThisWorkbook
        Settings.Format1 = .CustomDocumentProperties("Format1")
        Settings.Format2 = .CustomDocumentProperties("Format2")
        Settings.Format3 = .CustomDocumentProperties("Format3")
        Settings.EnabledFormats = .CustomDocumentProperties("EnabledFormats")
    End With
    
    If Err.Number <> 0 Then
        InitializeSettings
        SaveSettings
    End If
End Sub

Public Sub CycleNumberFormat(Optional control As IRibbonControl)
    If Selection Is Nothing Then Exit Sub
    
    Dim currentFormat As String
    currentFormat = Selection.NumberFormat
    
    ' Determine next format based on current format
    Dim nextFormat As String
    Select Case currentFormat
        Case Settings.Format1
            If (Settings.EnabledFormats And 2) Then
                nextFormat = Settings.Format2
            ElseIf (Settings.EnabledFormats And 4) Then
                nextFormat = Settings.Format3
            Else
                nextFormat = Settings.Format1
            End If
        Case Settings.Format2
            If (Settings.EnabledFormats And 4) Then
                nextFormat = Settings.Format3
            ElseIf (Settings.EnabledFormats And 1) Then
                nextFormat = Settings.Format1
            Else
                nextFormat = Settings.Format2
            End If
        Case Settings.Format3
            If (Settings.EnabledFormats And 1) Then
                nextFormat = Settings.Format1
            ElseIf (Settings.EnabledFormats And 2) Then
                nextFormat = Settings.Format2
            Else
                nextFormat = Settings.Format3
            End If
        Case Else
            If (Settings.EnabledFormats And 1) Then
                nextFormat = Settings.Format1
            ElseIf (Settings.EnabledFormats And 2) Then
                nextFormat = Settings.Format2
            ElseIf (Settings.EnabledFormats And 4) Then
                nextFormat = Settings.Format3
            End If
    End Select
    
    Selection.NumberFormat = nextFormat
End Sub

Private Sub CreateCustomProperties()
    Dim prop As DocumentProperty
    
    With ThisWorkbook.CustomDocumentProperties
        ' Check and create Format1 property
        On Error Resume Next
        Set prop = .item("Format1")
        If Err.Number <> 0 Then
            .Add Name:="Format1", LinkToContent:=False, Type:=msoPropertyTypeString, value:=DEFAULT_FORMAT_1
        End If
        
        ' Check and create Format2 property
        Set prop = .item("Format2")
        If Err.Number <> 0 Then
            .Add Name:="Format2", LinkToContent:=False, Type:=msoPropertyTypeString, value:=DEFAULT_FORMAT_2
        End If
        
        ' Check and create Format3 property
        Set prop = .item("Format3")
        If Err.Number <> 0 Then
            .Add Name:="Format3", LinkToContent:=False, Type:=msoPropertyTypeString, value:=DEFAULT_FORMAT_3
        End If
        
        ' Check and create EnabledFormats property
        Set prop = .item("EnabledFormats")
        If Err.Number <> 0 Then
            .Add Name:="EnabledFormats", LinkToContent:=False, Type:=msoPropertyTypeNumber, value:=7
        End If
        On Error GoTo 0
    End With
End Sub

Public Property Get Format1() As String
    Format1 = Settings.Format1
End Property

Public Property Let Format1(ByVal value As String)
    Settings.Format1 = value
End Property

Public Property Get Format2() As String
    Format2 = Settings.Format2
End Property

Public Property Let Format2(ByVal value As String)
    Settings.Format2 = value
End Property

Public Property Get Format3() As String
    Format3 = Settings.Format3
End Property

Public Property Let Format3(ByVal value As String)
    Settings.Format3 = value
End Property

Public Property Get EnabledFormats() As Integer
    EnabledFormats = Settings.EnabledFormats
End Property

Public Property Let EnabledFormats(ByVal value As Integer)
    Settings.EnabledFormats = value
End Property
