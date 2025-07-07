' ModUtilityHelpers.bas
' Version: 1.0.0
' Date: 2025-01-04
' Author: XLerate Development Team
' 
' CHANGELOG:
' v1.0.0 - Initial implementation of utility helper functions
'        - Status bar management functions
'        - Common utility operations
'        - Helper functions for other modules
'
' DESCRIPTION:
' Utility helper functions used by other XLerate modules
' Provides common functionality to reduce code duplication

Attribute VB_Name = "ModUtilityHelpers"
Option Explicit

Public Sub ClearStatusBar()
    ' Clears the Excel status bar
    ' Called by other modules after operations complete
    
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
    
    Debug.Print "Status bar cleared"
End Sub

Public Sub SetStatusBar(message As String)
    ' Sets a message in the Excel status bar
    ' Used for progress indication during long operations
    
    On Error Resume Next
    Application.StatusBar = message
    On Error GoTo 0
    
    Debug.Print "Status bar set: " & message
End Sub

 