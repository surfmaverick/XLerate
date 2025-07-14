Attribute VB_Name = "ModGlobalSettings"
' ModGlobalSettings.cls
Option Explicit

Private GlobalCellFormats As Collection  ' Store formats in a collection
Public Const FONT_BOLD As Long = 1
Public Const FONT_ITALIC As Long = 2
Public Const FONT_UNDERLINE As Long = 4
Public Const FONT_STRIKETHROUGH As Long = 8

Public Sub InitializeCellFormats()
    If GlobalCellFormats Is Nothing Then
        Set GlobalCellFormats = New Collection
    End If
    
    If Not LoadCellFormatsFromWorkbook() Then
        ' Create default formats
        Dim defaultFormat As clsCellFormatType
        Set defaultFormat = New clsCellFormatType
        With defaultFormat
            .Name = "Default"
            .BackColor = RGB(255, 255, 255)
            .BorderStyle = xlContinuous
            .BorderColor = RGB(0, 0, 0)
            .FillPattern = xlSolid
            .FontStyle = 0
            .FontColor = RGB(0, 0, 0)
        End With
        GlobalCellFormats.Add defaultFormat
        
        ' Add more default formats as needed...
        
        SaveCellFormatsToWorkbook
    End If
End Sub

Public Function GetCellFormatList() As Collection
    If GlobalCellFormats Is Nothing Then
        InitializeCellFormats
    End If
    Set GetCellFormatList = GlobalCellFormats
End Function

Private Sub SaveCellFormatsToWorkbook()
    Dim propValue As String
    Dim format As clsCellFormatType
    
    For Each format In GlobalCellFormats
        propValue = propValue & format.Name & "|" & _
                   format.BackColor & "|" & _
                   format.BorderStyle & "|" & _
                   format.BorderColor & "|" & _
                   format.FillPattern & "|" & _
                   format.FontStyle & "|" & _
                   format.FontColor & "||"
    Next format
    
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties("SavedCellFormats").Delete
    On Error GoTo 0
    ThisWorkbook.CustomDocumentProperties.Add Name:="SavedCellFormats", _
        LinkToContent:=False, Type:=msoPropertyTypeString, value:=propValue
        
    ThisWorkbook.Save
End Sub

Private Function LoadCellFormatsFromWorkbook() As Boolean
    On Error Resume Next
    Dim propValue As String
    propValue = ThisWorkbook.CustomDocumentProperties("SavedCellFormats")
    On Error GoTo 0
    
    If propValue = "" Then
        LoadCellFormatsFromWorkbook = False
        Exit Function
    End If
    
    Set GlobalCellFormats = New Collection
    
    Dim formatsArray() As String
    formatsArray = Split(propValue, "||")
    
    Dim i As Long
    For i = LBound(formatsArray) To UBound(formatsArray) - 1
        If formatsArray(i) <> "" Then
            Dim formatParts() As String
            formatParts = Split(formatsArray(i), "|")
            
            Dim newFormat As clsCellFormatType
            Set newFormat = New clsCellFormatType
            With newFormat
                .Name = formatParts(0)
                .BackColor = CLng(formatParts(1))
                .BorderStyle = CLng(formatParts(2))
                .BorderColor = CLng(formatParts(3))
                .FillPattern = CLng(formatParts(4))
                .FontStyle = CLng(formatParts(5))
                .FontColor = CLng(formatParts(6))
            End With
            GlobalCellFormats.Add newFormat
        End If
    Next i
    
    LoadCellFormatsFromWorkbook = (GlobalCellFormats.Count > 0)
End Function

Public Sub AddFormat(newFormat As clsCellFormatType)
    If GlobalCellFormats Is Nothing Then InitializeCellFormats
    GlobalCellFormats.Add newFormat
    SaveCellFormatsToWorkbook
End Sub

Public Sub RemoveFormat(index As Integer)
    If GlobalCellFormats Is Nothing Then Exit Sub
    If index <= 0 Or index > GlobalCellFormats.Count Then Exit Sub
    GlobalCellFormats.Remove index
    SaveCellFormatsToWorkbook
End Sub

Public Sub UpdateFormat(index As Integer, updatedFormat As clsCellFormatType)
    If GlobalCellFormats Is Nothing Then Exit Sub
    If index <= 0 Or index > GlobalCellFormats.Count Then Exit Sub
    
    GlobalCellFormats.Remove index
    GlobalCellFormats.Add updatedFormat, , , index - 1
    SaveCellFormatsToWorkbook
End Sub
