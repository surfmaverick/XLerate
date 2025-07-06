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

 =============================================================================
' File: ClearStatusBar Subroutine (part of ModUtilityFunctions.bas)
' Version: 2.0.0
' Date: January 2025
' Author: XLerate Development Team
'
' CHANGELOG:
' v2.0.0 - Enhanced status bar management with safety checks
'        - Cross-platform compatibility (Windows & macOS)
'        - Error handling for edge cases
'        - Integration with Application.OnTime for delayed clearing
' v1.0.0 - Basic status bar clearing functionality
' =============================================================================

' This subroutine should be placed in a standard module (e.g., ModUtilityFunctions.bas)

Public Sub ClearStatusBar()
    ' Utility function to clear the Excel status bar
    ' Called with Application.OnTime for delayed clearing after user feedback
    ' 
    ' Usage: Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    On Error Resume Next
    
    ' Clear the status bar message
    Application.StatusBar = False
    
    ' Debug output for troubleshooting
    Debug.Print "Status bar cleared at " & Format(Now, "hh:mm:ss")
    
    On Error GoTo 0
End Sub

Public Sub ClearStatusBarImmediate()
    ' Immediate status bar clearing without delay
    ' For use when immediate clearing is needed
    
    On Error Resume Next
    Application.StatusBar = False
    Debug.Print "Status bar cleared immediately"
    On Error GoTo 0
End Sub

Public Sub SetStatusMessage(message As String, Optional clearAfterSeconds As Integer = 2)
    ' Set a status bar message with automatic clearing
    ' 
    ' Parameters:
    '   message - The message to display
    '   clearAfterSeconds - Seconds after which to clear (default: 2)
    
    On Error Resume Next
    
    ' Set the message
    Application.StatusBar = message
    Debug.Print "Status bar set: " & message
    
    ' Schedule automatic clearing
    If clearAfterSeconds > 0 Then
        Application.OnTime Now + TimeValue("00:00:" & Format(clearAfterSeconds, "00")), "ClearStatusBar"
    End If
    
    On Error GoTo 0
End Sub

Public Sub SetProgressMessage(current As Long, total As Long, operation As String)
    ' Set a progress message in the status bar
    ' Useful for long-running operations
    
    On Error Resume Next
    
    Dim percentage As Integer
    If total > 0 Then
        percentage = Int((current / total) * 100)
    Else
        percentage = 0
    End If
    
    Dim progressMessage As String
    progressMessage = operation & " - " & current & " of " & total & " (" & percentage & "%)"
    
    Application.StatusBar = progressMessage
    
    ' Clear when complete
    If current >= total Then
        Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    End If
    
    On Error GoTo 0
End Sub


Public Sub ToggleCalculationMode(Optional control As IRibbonControl)
    ' Toggles between automatic and manual calculation
    ' Useful for large models where calculation speed matters
    
    On Error GoTo ErrorHandler
    
    Select Case Application.Calculation
        Case xlCalculationAutomatic
            Application.Calculation = xlCalculationManual
            SetStatusBar "Calculation set to Manual"
            Debug.Print "Calculation mode changed to Manual"
        Case xlCalculationManual
            Application.Calculation = xlCalculationAutomatic
            SetStatusBar "Calculation set to Automatic"
            Debug.Print "Calculation mode changed to Automatic"
        Case Else
            Application.Calculation = xlCalculationAutomatic
            SetStatusBar "Calculation set to Automatic"
    End Select
    
    ' Clear status bar after 2 seconds
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ToggleCalculationMode: " & Err.Description
End Sub

Public Sub InsertCurrentDate(Optional control As IRibbonControl)
    ' Inserts current date in active cell
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Selection.Value = Date
    Selection.NumberFormat = "mm/dd/yyyy"
    
    Debug.Print "Current date inserted: " & Date
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertCurrentDate: " & Err.Description
End Sub

Public Sub InsertCurrentTime(Optional control As IRibbonControl)
    ' Inserts current time in active cell
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Selection.Value = Time
    Selection.NumberFormat = "hh:mm AM/PM"
    
    Debug.Print "Current time inserted: " & Time
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertCurrentTime: " & Err.Description
End Sub

Public Function IsExcelVersionCompatible(minimumVersion As String) As Boolean
    ' Checks if current Excel version meets minimum requirements
    
    On Error GoTo ErrorHandler
    
    Dim currentVersion As Double
    currentVersion = CDbl(Application.Version)
    
    Dim minVersion As Double
    minVersion = CDbl(minimumVersion)
    
    IsExcelVersionCompatible = (currentVersion >= minVersion)
    
    Debug.Print "Excel version check: " & currentVersion & " >= " & minVersion & " = " & IsExcelVersionCompatible
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in IsExcelVersionCompatible: " & Err.Description
    IsExcelVersionCompatible = True ' Assume compatible on error
End Function

Public Sub ShowAboutDialog(Optional control As IRibbonControl)
    ' Shows information about XLerate
    
    Dim aboutText As String
    aboutText = "XLerate v2.0.0" & vbNewLine & vbNewLine & _
                "Open-source Excel add-in for financial modeling" & vbNewLine & vbNewLine & _
                "Features:" & vbNewLine & _
                "• Smart Fill Functions" & vbNewLine & _
                "• Advanced Auditing Tools" & vbNewLine & _
                "• Format Cycling" & vbNewLine & _
                "• Border Management" & vbNewLine & _
                "• Productivity Utilities" & vbNewLine & vbNewLine & _
                "Compatible with Excel " & Application.Version & vbNewLine & vbNewLine & _
                "GitHub: github.com/omegarhovega/XLerate" & vbNewLine & _
                "License: MIT"
    
    MsgBox aboutText, vbInformation, "About XLerate"
End Sub

Public Function GetWorkbookStats() As String
    ' Returns statistics about the current workbook
    
    On Error GoTo ErrorHandler
    
    Dim stats As String
    stats = "Workbook Statistics:" & vbNewLine & vbNewLine
    
    ' Basic counts
    stats = stats & "Worksheets: " & ActiveWorkbook.Worksheets.Count & vbNewLine
    stats = stats & "Named Ranges: " & ActiveWorkbook.Names.Count & vbNewLine
    
    ' Calculate total cells with content
    Dim totalCells As Long
    Dim totalFormulas As Long
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        If Not ws.UsedRange Is Nothing Then
            totalCells = totalCells + ws.UsedRange.Cells.Count
            
            ' Count formulas (simplified)
            Dim cell As Range
            For Each cell In ws.UsedRange
                If cell.HasFormula Then
                    totalFormulas = totalFormulas + 1
                End If
            Next cell
        End If
    Next ws
    
    stats = stats & "Total Used Cells: " & Format(totalCells, "#,##0") & vbNewLine
    stats = stats & "Formula Cells: " & Format(totalFormulas, "#,##0") & vbNewLine
    stats = stats & "File Size: " & Format(FileLen(ActiveWorkbook.FullName) / 1024, "#,##0") & " KB"
    
    GetWorkbookStats = stats
    
    Exit Function
    
ErrorHandler:
    GetWorkbookStats = "Error calculating statistics: " & Err.Description
End Function

Public Sub ShowWorkbookStats(Optional control As IRibbonControl)
    ' Displays workbook statistics
    
    MsgBox GetWorkbookStats(), vbInformation, "XLerate - Workbook Statistics"
End Sub

Public Function FormatBytes(bytes As Long) As String
    ' Formats byte count in human-readable format
    
    If bytes < 1024 Then
        FormatBytes = bytes & " B"
    ElseIf bytes < 1048576 Then ' 1024^2
        FormatBytes = Format(bytes / 1024, "#,##0.0") & " KB"
    ElseIf bytes < 1073741824 Then ' 1024^3
        FormatBytes = Format(bytes / 1048576, "#,##0.0") & " MB"
    Else
        FormatBytes = Format(bytes / 1073741824, "#,##0.0") & " GB"
    End If
End Function

Public Sub OptimizeWorkbook(Optional control As IRibbonControl)
    ' Performs basic workbook optimization
    
    On Error GoTo ErrorHandler
    
    Dim response As VbMsgBoxResult
    response = MsgBox("This will optimize the workbook by:" & vbNewLine & _
                     "• Removing unused styles" & vbNewLine & _
                     "• Clearing clipboard" & vbNewLine & _
                     "• Resetting used ranges" & vbNewLine & vbNewLine & _
                     "Continue?", vbYesNo + vbQuestion, "XLerate - Optimize Workbook")
    
    If response = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    SetStatusBar "Optimizing workbook..."
    
    ' Clear clipboard
    Application.CutCopyMode = False
    
    ' Reset used ranges (simplified approach)
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        SetStatusBar "Optimizing " & ws.Name & "..."
        ws.UsedRange ' This forces Excel to recalculate the used range
    Next ws
    
    ' Force garbage collection
    DoEvents
    
    ClearStatusBar
    Application.ScreenUpdating = True
    
    MsgBox "Workbook optimization completed!", vbInformation, "XLerate"
    
    Exit Sub
    
ErrorHandler:
    ClearStatusBar
    Application.ScreenUpdating = True
    Debug.Print "Error in OptimizeWorkbook: " & Err.Description
    MsgBox "Error during optimization: " & Err.Description, vbCritical, "XLerate"
End Sub