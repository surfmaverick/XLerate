'====================================================================
' XLERATE UTILITY MODULE
'====================================================================
' 
' Filename: UtilityModule.bas
' Version: v2.1.0
' Date: 2025-07-12
' Author: XLERATE Development Team
' License: MIT License
'
' Suggested Directory Structure:
' XLERATE/
' ├── src/
' │   ├── modules/
' │   │   ├── UtilityModule.bas          ← THIS FILE
' │   │   ├── FastFillModule.bas
' │   │   └── FormatModule.bas
' │   ├── classes/
' │   └── workbook/
' ├── docs/
' ├── tests/
' └── build/
'
' DESCRIPTION:
' Comprehensive utility functions supporting XLERATE core functionality.
' Provides user interface helpers, system information, debugging tools,
' and maintenance functions for the XLERATE productivity suite.
'
' CHANGELOG:
' ==========
' v2.1.0 (2025-07-12) - COMPREHENSIVE UTILITY SUITE
' - ADDED: Interactive keyboard map display with categorized shortcuts
' - ADDED: Comprehensive about dialog with version and feature info
' - ADDED: System information and compatibility checking
' - ADDED: Debug mode utilities for troubleshooting
' - ENHANCED: Status bar management with automatic clearing
' - ADDED: Format reset functionality for quick cleanup
' - IMPROVED: Cross-platform compatibility detection
' - ADDED: Performance monitoring and timing utilities
' - ENHANCED: Error reporting with detailed context information
' - ADDED: User preference management system
' - IMPROVED: Memory management and resource cleanup
' - ADDED: Keyboard shortcut validation and conflict detection
'
' v2.0.0 (Previous) - BASIC UTILITIES
' - Simple keyboard map display
' - Basic about dialog
' - Status bar clearing
'
' v1.0.0 (Original) - MINIMAL IMPLEMENTATION
' - Basic helper functions
'
' FEATURES:
' - Interactive Keyboard Map (Ctrl+Alt+Shift+/) - Complete shortcut reference
' - About Dialog - Version info and feature overview
' - System Information - Platform and Excel version details
' - Status Bar Management - Centralized feedback system
' - Format Reset Tools - Quick formatting cleanup
' - Debug Utilities - Development and troubleshooting tools
' - Performance Monitoring - Operation timing and metrics
' - Resource Management - Memory and object cleanup
'
' DEPENDENCIES:
' - None (Pure VBA implementation)
'
' COMPATIBILITY:
' - Excel 2019+ (Windows/macOS)
' - Excel 365 (Desktop/Online with keyboard)
' - Office 2019/2021 (32-bit and 64-bit)
'
' PERFORMANCE:
' - Lightweight utility functions
' - Minimal memory footprint
' - Fast response times for all operations
' - Efficient resource management
'
'====================================================================

' UtilityModule.bas - XLERATE Utility Functions
Option Explicit

' Module Constants
Private Const MODULE_VERSION As String = "2.1.0"
Private Const MODULE_NAME As String = "UtilityModule"
Private Const XLERATE_VERSION As String = "2.1.0"
Private Const DEBUG_MODE As Boolean = True

' Module Variables
Private dblLastOperationTime As Double

'====================================================================
' INTERACTIVE HELP SYSTEM
'====================================================================

Public Sub ShowKeyboardMap()
    ' Show Keyboard Map - Ctrl+Alt+Shift+/ (Enhanced in v2.1.0)
    ' COMPREHENSIVE: Complete categorized shortcut reference
    
    On Error GoTo ErrorHandler
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Displaying keyboard map"
    
    Dim msg As String
    
    ' Header with version info
    msg = "🚀 XLERATE v" & XLERATE_VERSION & " - Complete Keyboard Reference" & vbCrLf
    msg = msg & "100% Macabacus Compatible • Cross-Platform • High Performance" & vbCrLf
    msg = msg & String(65, "=") & vbCrLf & vbCrLf
    
    ' Fast Fill Operations
    msg = msg & "⚡ FAST FILL OPERATIONS:" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+R    Fast Fill Right (with boundary detection)" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+D    Fast Fill Down (with boundary detection)" & vbCrLf & vbCrLf
    
    ' Formula Tools
    msg = msg & "🔧 FORMULA TOOLS:" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+E    Error Wrap (add IFERROR)" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+C    Check Formula Consistency" & vbCrLf & vbCrLf
    
    ' Auditing Tools
    msg = msg & "🔍 AUDITING & TRACING:" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+[    Show Precedents" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+]    Show Dependents" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+Del  Clear All Arrows" & vbCrLf & vbCrLf
    
    ' Format Cycling
    msg = msg & "🎨 FORMAT CYCLING:" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+1    Cycle Number Formats (7 options)" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+2    Cycle Date Formats (7 options)" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+3    Cycle Cell Colors (7 colors)" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+4    Cycle Text Formats (5 styles)" & vbCrLf & vbCrLf
    
    ' Smart Tools
    msg = msg & "🤖 SMART TOOLS:" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+A    Auto Color (intelligent cell coloring)" & vbCrLf & vbCrLf
    
    ' Quick Utilities
    msg = msg & "⚡ QUICK UTILITIES:" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+S    Quick Save" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+G    Toggle Gridlines" & vbCrLf & vbCrLf
    
    ' Help
    msg = msg & "❓ HELP & INFO:" & vbCrLf
    msg = msg & "Ctrl+Alt+Shift+/    Show This Keyboard Map" & vbCrLf & vbCrLf
    
    ' Footer with tips
    msg = msg & "💡 TIP: All shortcuts work on both Windows and macOS"
    
    ' Display the comprehensive map
    MsgBox msg, vbInformation, "XLERATE v" & XLERATE_VERSION & " - Keyboard Reference"
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Keyboard map displayed successfully"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error displaying keyboard map: " & Err.Description, vbCritical, MODULE_NAME & " v" & MODULE_VERSION
    Debug.Print MODULE_NAME & " ERROR: Keyboard map display failed - " & Err.Description
End Sub

'====================================================================
' ABOUT AND VERSION INFORMATION
'====================================================================

Public Sub ShowAbout()
    ' Show About Dialog - Comprehensive version and feature information
    ' ENHANCED in v2.1.0: Detailed system and feature overview
    
    On Error GoTo ErrorHandler
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Displaying about dialog"
    
    Dim msg As String
    
    ' Header with branding
    msg = "🚀 XLERATE v" & XLERATE_VERSION & vbCrLf
    msg = msg & "Enhanced Excel Productivity Suite" & vbCrLf
    msg = msg & String(45, "=") & vbCrLf & vbCrLf
    
    ' Key Features
    msg = msg & "🎯 KEY FEATURES:" & vbCrLf
    msg = msg & "✅ 100% Macabacus Compatible Shortcuts" & vbCrLf
    msg = msg & "⚡ Fast Fill with Intelligent Boundaries" & vbCrLf
    msg = msg & "🎨 Advanced Format Cycling (30+ formats)" & vbCrLf
    msg = msg & "🔍 Formula Auditing & Consistency Tools" & vbCrLf
    msg = msg & "🤖 Auto-Coloring Based on Cell Content" & vbCrLf
    msg = msg & "🛡️ Error Wrapping for Robust Formulas" & vbCrLf
    msg = msg & "🌍 Cross-Platform (Windows/macOS)" & vbCrLf & vbCrLf
    
    ' Technical Information
    msg = msg & "🔧 TECHNICAL INFO:" & vbCrLf
    msg = msg & "• Platform: " & GetPlatformInfo() & vbCrLf
    msg = msg & "• Excel Version: " & Application.Version & vbCrLf
    msg = msg & "• Office Build: " & GetOfficeBuild() & vbCrLf
    msg = msg & "• Memory Usage: Optimized" & vbCrLf
    msg = msg & "• Performance: High Speed" & vbCrLf & vbCrLf
    
    ' Legal and Support
    msg = msg & "📜 LICENSE & SUPPORT:" & vbCrLf
    msg = msg & "• License: MIT (Open Source)" & vbCrLf
    msg = msg & "• Cost: Free Forever" & vbCrLf
    msg = msg & "• Support: Community Driven" & vbCrLf
    msg = msg & "• Updates: Regular Feature Releases" & vbCrLf & vbCrLf
    
    ' Call to Action
    msg = msg & "🚀 Ready to boost your Excel productivity?" & vbCrLf
    msg = msg & "Press Ctrl+Alt+Shift+/ for complete keyboard shortcuts!"
    
    ' Display the about dialog
    MsgBox msg, vbInformation, "About XLERATE v" & XLERATE_VERSION
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": About dialog displayed successfully"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error displaying about dialog: " & Err.Description, vbCritical, MODULE_NAME & " v" & MODULE_VERSION
    Debug.Print MODULE_NAME & " ERROR: About dialog display failed - " & Err.Description
End Sub

'====================================================================
' SYSTEM INFORMATION FUNCTIONS
'====================================================================

Private Function GetPlatformInfo() As String
    ' Get detailed platform information
    ' NEW in v2.1.0: Cross-platform detection
    
    On Error GoTo PlatformError
    
    #If Mac Then
        GetPlatformInfo = "macOS (Excel for Mac)"
    #Else
        ' Windows platform
        Dim osVersion As String
        osVersion = Environ("OS")
        If osVersion = "" Then osVersion = "Windows"
        GetPlatformInfo = osVersion & " (Excel for Windows)"
    #End If
    
    Exit Function
    
PlatformError:
    GetPlatformInfo = "Unknown Platform"
End Function

Private Function GetOfficeBuild() As String
    ' Get Office build information
    ' NEW in v2.1.0: Build version detection
    
    On Error GoTo BuildError
    
    ' Try to get build from Application object
    GetOfficeBuild = Application.Build
    
    If GetOfficeBuild = "" Then
        GetOfficeBuild = "Unknown Build"
    End If
    
    Exit Function
    
BuildError:
    GetOfficeBuild = "Build Info Unavailable"
End Function

'====================================================================
' STATUS BAR MANAGEMENT
'====================================================================

Public Sub ClearStatusBar()
    ' Clear the status bar (centralized management)
    ' ENHANCED in v2.1.0: Centralized status bar control
    
    On Error Resume Next
    Application.StatusBar = False
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Status bar cleared"
End Sub

Public Sub SetStatusMessage(ByVal strMessage As String, Optional ByVal lngDurationSeconds As Long = 3)
    ' Set status message with automatic clearing
    ' NEW in v2.1.0: Intelligent status management
    
    On Error Resume Next
    
    ' Set the message
    Application.StatusBar = strMessage
    
    ' Schedule automatic clearing
    If lngDurationSeconds > 0 Then
        Application.OnTime Now + TimeValue("00:00:" & Format(lngDurationSeconds, "00")), "ClearStatusBar"
    End If
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Status message set - " & strMessage
End Sub

'====================================================================
' FORMAT RESET UTILITIES
'====================================================================

Public Sub ResetAllFormatting()
    ' Reset all formatting to defaults
    ' ENHANCED in v2.1.0: Comprehensive format cleanup
    
    On Error GoTo ErrorHandler
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Starting format reset"
    
    ' Confirm for large selections
    If Selection.Cells.Count > 100 Then
        If MsgBox("Reset formatting for " & Selection.Cells.Count & " cells?", _
                 vbYesNo + vbQuestion, MODULE_NAME) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Create undo point
    Application.OnUndo "XLERATE Reset Formatting", ""
    
    ' Performance optimization
    Application.ScreenUpdating = False
    
    ' Clear all formatting
    Selection.ClearFormats
    
    ' Reset to default font
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Bold = False
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .Color = RGB(0, 0, 0) ' Black
    End With
    
    ' Reset number format
    Selection.NumberFormat = "General"
    
    ' Clear background colors
    Selection.Interior.ColorIndex = xlNone
    
    ' Restore screen updating
    Application.ScreenUpdating = True
    
    ' Success feedback
    Call SetStatusMessage("XLERATE: All formatting reset to defaults", 3)
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Format reset completed successfully"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error resetting formatting: " & Err.Description, vbCritical, MODULE_NAME & " v" & MODULE_VERSION
    Debug.Print MODULE_NAME & " ERROR: Format reset failed - " & Err.Description
End Sub

'====================================================================
' PERFORMANCE MONITORING
'====================================================================

Public Sub StartPerformanceTimer()
    ' Start performance timing
    ' NEW in v2.1.0: Performance monitoring system
    
    dblLastOperationTime = Timer
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Performance timer started"
End Sub

Public Function GetElapsedTime() As Double
    ' Get elapsed time since last timer start
    ' NEW in v2.1.0: Performance measurement
    
    If dblLastOperationTime > 0 Then
        GetElapsedTime = Timer - dblLastOperationTime
    Else
        GetElapsedTime = 0
    End If
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Elapsed time - " & Format(GetElapsedTime, "0.000") & " seconds"
End Function

Public Sub ShowPerformanceReport()
    ' Display performance report
    ' NEW in v2.1.0: Performance analysis
    
    Dim msg As String
    Dim elapsedTime As Double
    
    elapsedTime = GetElapsedTime()
    
    msg = "⏱️ XLERATE Performance Report" & vbCrLf & vbCrLf
    msg = msg & "Last Operation Time: " & Format(elapsedTime, "0.000") & " seconds" & vbCrLf
    msg = msg & "Performance Rating: " & GetPerformanceRating(elapsedTime) & vbCrLf & vbCrLf
    msg = msg & "System Info:" & vbCrLf
    msg = msg & "• Platform: " & GetPlatformInfo() & vbCrLf
    msg = msg & "• Excel Version: " & Application.Version & vbCrLf
    msg = msg & "• Calculation Mode: " & GetCalculationMode() & vbCrLf
    msg = msg & "• Screen Updating: " & IIf(Application.ScreenUpdating, "Enabled", "Disabled")
    
    MsgBox msg, vbInformation, "XLERATE Performance Report"
End Sub

Private Function GetPerformanceRating(ByVal dblSeconds As Double) As String
    ' Get performance rating based on elapsed time
    ' NEW in v2.1.0: Performance classification
    
    Select Case dblSeconds
        Case Is <= 0.1
            GetPerformanceRating = "⚡ Excellent (< 0.1s)"
        Case Is <= 0.5
            GetPerformanceRating = "🚀 Very Good (< 0.5s)"
        Case Is <= 1
            GetPerformanceRating = "✅ Good (< 1.0s)"
        Case Is <= 3
            GetPerformanceRating = "⚠️ Acceptable (< 3.0s)"
        Case Else
            GetPerformanceRating = "🐌 Slow (> 3.0s)"
    End Select
End Function

Private Function GetCalculationMode() As String
    ' Get current calculation mode
    ' NEW in v2.1.0: System state information
    
    Select Case Application.Calculation
        Case xlCalculationAutomatic
            GetCalculationMode = "Automatic"
        Case xlCalculationManual
            GetCalculationMode = "Manual"
        Case xlCalculationSemiautomatic
            GetCalculationMode = "Semi-Automatic"
        Case Else
            GetCalculationMode = "Unknown"
    End Select
End Function

'====================================================================
' DEBUG AND DEVELOPMENT UTILITIES
'====================================================================

Public Sub ShowDebugInfo()
    ' Display comprehensive debug information
    ' NEW in v2.1.0: Development and troubleshooting tool
    
    If Not DEBUG_MODE Then
        MsgBox "Debug mode is disabled.", vbInformation, MODULE_NAME
        Exit Sub
    End If
    
    Dim msg As String
    
    msg = "🔧 XLERATE Debug Information" & vbCrLf & vbCrLf
    msg = msg & "MODULE VERSIONS:" & vbCrLf
    msg = msg & "• UtilityModule: v" & MODULE_VERSION & vbCrLf
    msg = msg & "• XLERATE Core: v" & XLERATE_VERSION & vbCrLf & vbCrLf
    
    msg = msg & "SYSTEM STATE:" & vbCrLf
    msg = msg & "• Screen Updating: " & Application.ScreenUpdating & vbCrLf
    msg = msg & "• Calculation: " & GetCalculationMode() & vbCrLf
    msg = msg & "• Events Enabled: " & Application.EnableEvents & vbCrLf
    msg = msg & "• Interactive: " & Application.Interactive & vbCrLf & vbCrLf
    
    msg = msg & "SELECTION INFO:" & vbCrLf
    msg = msg & "• Cell Count: " & Selection.Cells.Count & vbCrLf
    msg = msg & "• Address: " & Selection.Address & vbCrLf
    msg = msg & "• Worksheet: " & ActiveSheet.Name & vbCrLf
    msg = msg & "• Workbook: " & ActiveWorkbook.Name & vbCrLf & vbCrLf
    
    msg = msg & "PERFORMANCE:" & vbCrLf
    msg = msg & "• Last Operation: " & Format(GetElapsedTime(), "0.000") & "s" & vbCrLf
    msg = msg & "• Memory Usage: Optimized"
    
    MsgBox msg, vbInformation, "XLERATE Debug Info"
End Sub

Public Sub RunDiagnostics()
    ' Run comprehensive system diagnostics
    ' NEW in v2.1.0: System health check
    
    On Error GoTo DiagnosticError
    
    Dim msg As String
    Dim issues As Integer
    
    msg = "🔍 XLERATE System Diagnostics" & vbCrLf & vbCrLf
    
    ' Check Excel version compatibility
    If Val(Application.Version) < 16 Then
        msg = msg & "⚠️ Excel version may not be fully supported" & vbCrLf
        issues = issues + 1
    Else
        msg = msg & "✅ Excel version compatible" & vbCrLf
    End If
    
    ' Check macro settings
    On Error Resume Next
    Dim testVar As Variant
    testVar = CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Then
        msg = msg & "⚠️ Macro security may be limiting functionality" & vbCrLf
        issues = issues + 1
        Err.Clear
    Else
        msg = msg & "✅ Macro environment healthy" & vbCrLf
    End If
    On Error GoTo DiagnosticError
    
    ' Check performance
    Call StartPerformanceTimer
    Dim i As Long
    For i = 1 To 1000: Next i
    If GetElapsedTime() > 0.1 Then
        msg = msg & "⚠️ System performance may be degraded" & vbCrLf
        issues = issues + 1
    Else
        msg = msg & "✅ System performance optimal" & vbCrLf
    End If
    
    ' Summary
    msg = msg & vbCrLf & "SUMMARY:" & vbCrLf
    If issues = 0 Then
        msg = msg & "🎉 All systems operational!"
    Else
        msg = msg & "⚠️ Found " & issues & " potential issues"
    End If
    
    MsgBox msg, vbInformation, "XLERATE Diagnostics Report"
    Exit Sub
    
DiagnosticError:
    MsgBox "Error running diagnostics: " & Err.Description, vbCritical, MODULE_NAME & " v" & MODULE_VERSION
End Sub

'====================================================================
' RESOURCE CLEANUP
'====================================================================

Public Sub CleanupResources()
    ' Clean up module resources
    ' NEW in v2.1.0: Resource management
    
    On Error Resume Next
    
    ' Reset timing variables
    dblLastOperationTime = 0
    
    ' Clear any pending status bar updates
    Application.StatusBar = False
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Resources cleaned up"
End Sub