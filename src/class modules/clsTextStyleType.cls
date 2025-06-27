' =============================================================================
' File: clsFormatType.cls
' Version: 2.0.0
' Description: Format type class for number and date formatting
' Author: XLerate Development Team
' Created: Enhanced for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' Properties for format definition
Public Name As String
Public FormatCode As String

Private Sub Class_Initialize()
    ' Set default values
    Name = "Default Format"
    FormatCode = "General"
End Sub

Public Function GetDisplayName() As String
    ' Return formatted display name for UI
    GetDisplayName = Name & " (" & FormatCode & ")"
End Function

Public Function IsValid() As Boolean
    ' Validate the format
    IsValid = (Len(Name) > 0 And Len(FormatCode) > 0)
End Function

Public Sub CopyFrom(sourceFormat As clsFormatType)
    ' Copy properties from another format
    Name = sourceFormat.Name
    FormatCode = sourceFormat.FormatCode
End Sub

Public Function Clone() As clsFormatType
    ' Create a copy of this format
    Dim newFormat As New clsFormatType
    newFormat.Name = Name
    newFormat.FormatCode = FormatCode
    Set Clone = newFormat
End Function