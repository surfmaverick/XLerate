' =========================================================================
' UPDATED: RibbonCallbacks.bas v2.1.0 - Enhanced Ribbon Callbacks
' File: src/modules/RibbonCallbacks.bas
' Version: 2.1.0 (UPDATED from existing v2.0.0)
' Date: 2025-07-06
' Author: XLerate Development Team
' =========================================================================
'
' CHANGELOG v2.1.0:
' - ADDED: FastFillDown callback for new Ctrl+Alt+Shift+D functionality
' - ADDED: CurrencyCycling callback for new Ctrl+Alt+Shift+6 functionality
' - ENHANCED: Macabacus-compatible callback naming
' - IMPROVED: Error handling for all callbacks
' - ADDED: Version information and diagnostics
' - RETAINED: All existing callback functionality
'
' CHANGES FROM v2.0.0:
' - Added DoFastFillDown callback
' - Added DoCycleCurrency callback  
' - Enhanced error messaging with module identification
' - Added callback version tracking
' =========================================================================

Attribute VB_Name = "RibbonCallbacks"
Option Explicit

' Version information
Private Const CALLBACKS_VERSION As String = "2.1.0"

' Callback for customUI.onLoad
Public myRibbon As IRibbonUI

' =========================================================================
' RIBBON INITIALIZATION
' =========================================================================

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    ' Store ribbon reference and initialize
    Set myRibbon = ribbon
    Debug.Print "XLerate v" & CALLBACKS_VERSION & " - Ribbon callbacks loaded successfully"
End Sub

' =========================================================================
' AUDITING CALLBACKS (Existing - Enhanced)
' =========================================================================

Public Sub FindAndDisplayPrecedents(control As IRibbonControl)
    ' Callback for Trace Precedents button - Ctrl+Alt+Shift+[
    On Error GoTo ErrorHandler
    Application.Run "TraceUtils.ShowTracePrecedents"
    Debug.Print "Trace Precedents executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("FindAndDisplayPrecedents", Err.Description)
End Sub

Public Sub FindAndDisplayDependents(control As IRibbonControl)
    ' Callback for Trace Dependents button - Ctrl+Alt+Shift+]
    On Error GoTo ErrorHandler
    Application.Run "TraceUtils.ShowTraceDependents"
    Debug.Print "Trace Dependents executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("FindAndDisplayDependents", Err.Description)
End Sub

Public Sub OnCheckHorizontalConsistency(control As IRibbonControl)
    ' Callback for Formula Consistency button - Ctrl+Alt+Shift+C
    On Error GoTo ErrorHandler
    Application.Run "FormulaConsistency.CheckHorizontalConsistency"
    Debug.Print "Formula Consistency executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("OnCheckHorizontalConsistency", Err.Description)
End Sub

' =========================================================================
' FORMAT CYCLING CALLBACKS (Existing - Enhanced)
' =========================================================================

Public Sub DoCycleNumberFormat(control As IRibbonControl)
    ' Callback for Number Format Cycling - Ctrl+Alt+Shift+1
    On Error GoTo ErrorHandler
    Application.Run "ModNumberFormat.CycleNumberFormat"
    Debug.Print "Number format cycle executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("DoCycleNumberFormat", Err.Description)
End Sub

Public Sub DoCycleCellFormat(control As IRibbonControl)
    ' Callback for Cell Format Cycling - Ctrl+Alt+Shift+3
    On Error GoTo ErrorHandler
    Application.Run "ModCellFormat.CycleCellFormat"
    Debug.Print "Cell format cycle executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("DoCycleCellFormat", Err.Description)
End Sub

Public Sub DoCycleDateFormat(control As IRibbonControl)
    ' Callback for Date Format Cycling - Ctrl+Alt+Shift+2
    On Error GoTo ErrorHandler
    Application.Run "ModDateFormat.CycleDateFormat"
    Debug.Print "Date format cycle executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("DoCycleDateFormat", Err.Description)
End Sub

Public Sub DoCycleTextStyle(control As IRibbonControl)
    ' Callback for Text Style Cycling - Ctrl+Alt+Shift+4
    On Error GoTo ErrorHandler
    Application.Run "ModTextStyle.CycleTextStyle"
    Debug.Print "Text style cycle executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("DoCycleTextStyle", Err.Description)
End Sub

' =========================================================================
' NEW CALLBACKS - v2.1.0 Additions
' =========================================================================

Public Sub DoCycleCurrency(control As IRibbonControl)
    ' NEW: Callback for Currency Cycling - Ctrl+Alt+Shift+6
    On Error GoTo ErrorHandler
    Application.Run "ModCurrencyCycling.CycleCurrency", control
    Debug.Print "Currency cycle executed via ribbon (v" & CALLBACKS_VERSION & ") - NEW!"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("DoCycleCurrency", Err.Description)
End Sub

Public Sub DoFastFillDown(control As IRibbonControl)
    ' NEW: Callback for Fast Fill Down - Ctrl+Alt+Shift+D
    On Error GoTo ErrorHandler
    Application.Run "ModFastFillDown.FastFillDown", control
    Debug.Print "Fast Fill Down executed via ribbon (v" & CALLBACKS_VERSION & ") - NEW!"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("DoFastFillDown", Err.Description)
End Sub

Public Sub DoSmartFillDown(control As IRibbonControl)
    ' NEW: Callback for Enhanced Smart Fill Down
    On Error GoTo ErrorHandler
    Application.Run "ModFastFillDown.SmartFillDown", control
    Debug.Print "Smart Fill Down executed via ribbon (v" & CALLBACKS_VERSION & ") - NEW!"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("DoSmartFillDown", Err.Description)
End Sub

' =========================================================================
' MODELING CALLBACKS (Existing - Enhanced)
' =========================================================================

Public Sub SmartFillRight(control As IRibbonControl)
    ' Callback for Smart Fill Right - Ctrl+Alt+Shift+R
    On Error GoTo ErrorHandler
    Application.Run "ModSmartFillRight.SmartFillRight", control
    Debug.Print "Smart Fill Right executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("SmartFillRight", Err.Description)
End Sub

Public Sub SwitchCellSign(control As IRibbonControl)
    ' Callback for Switch Sign
    On Error GoTo ErrorHandler
    Application.Run "ModSwitchSign.SwitchCellSign", control
    Debug.Print "Switch Sign executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("SwitchCellSign", Err.Description)
End Sub

Public Sub WrapWithError(control As IRibbonControl)
    ' Callback for Error Wrap - Ctrl+Alt+Shift+E
    On Error GoTo ErrorHandler
    Application.Run "ModErrorWrap.WrapWithError", control
    Debug.Print "Error Wrap executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("WrapWithError", Err.Description)
End Sub

' =========================================================================
' COLOR AND FORMATTING CALLBACKS (Existing - Enhanced)
' =========================================================================

Public Sub AutoColor(control As IRibbonControl)
    ' Callback for Auto Color - Ctrl+Alt+Shift+A
    On Error GoTo ErrorHandler
    Application.Run "AutoColorModule.AutoColorSelection", control
    Debug.Print "Auto Color executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("AutoColor", Err.Description)
End Sub

Public Sub ResetFormatting(control As IRibbonControl)
    ' Callback for Format Reset
    On Error GoTo ErrorHandler
    Application.Run "ModFormatReset.ResetFormatting", control
    Debug.Print "Reset Formatting executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("ResetFormatting", Err.Description)
End Sub

' =========================================================================
' SETTINGS AND HELP CALLBACKS (Enhanced)
' =========================================================================

Public Sub ShowSettings(control As IRibbonControl)
    ' Callback for Settings Manager
    On Error GoTo ErrorHandler
    ' Check if the form exists and show it
    On Error Resume Next
    UserForms("frmSettingsManager").Show
    If Err.Number <> 0 Then
        MsgBox "Settings dialog is not available in this version.", vbInformation, "XLerate Settings"
        Debug.Print "Settings form not found: " & Err.Description
    End If
    On Error GoTo 0
    Debug.Print "Settings executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("ShowSettings", Err.Description)
End Sub

Public Sub ShowNumberSettings(control As IRibbonControl)
    ' Callback for Number Format Settings
    On Error GoTo ErrorHandler
    On Error Resume Next
    UserForms("frmNumberSettings").Show
    If Err.Number <> 0 Then
        MsgBox "Number settings dialog is not available.", vbInformation, "Number Settings"
    End If
    On Error GoTo 0
    Debug.Print "Number Settings executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("ShowNumberSettings", Err.Description)
End Sub

Public Sub ShowCurrencySettings(control As IRibbonControl)
    ' NEW: Callback for Currency Settings
    On Error GoTo ErrorHandler
    ' For now, show help since we don't have a dedicated form yet
    Application.Run "ModCurrencyCycling.ShowCurrencyHelp"
    Debug.Print "Currency Settings executed via ribbon (v" & CALLBACKS_VERSION & ") - NEW!"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("ShowCurrencySettings", Err.Description)
End Sub

Public Sub ShowAbout(control As IRibbonControl)
    ' Enhanced About dialog with version information
    On Error GoTo ErrorHandler
    
    Dim aboutText As String
    aboutText = "XLerate v2.1.0 - Enhanced Excel Productivity" & vbCrLf & vbCrLf
    aboutText = aboutText & "🚀 Features:" & vbCrLf
    aboutText = aboutText & "• 100% Macabacus shortcut compatibility" & vbCrLf
    aboutText = aboutText & "• Fast Fill Down with pattern detection" & vbCrLf
    aboutText = aboutText & "• Advanced currency cycling (20+ formats)" & vbCrLf
    aboutText = aboutText & "• Enhanced formula tracing and consistency" & vbCrLf
    aboutText = aboutText & "• Cross-platform support (Windows & macOS)" & vbCrLf & vbCrLf
    aboutText = aboutText & "🎯 New in v2.1.0:" & vbCrLf
    aboutText = aboutText & "• Fast Fill Down: Ctrl+Alt+Shift+D" & vbCrLf
    aboutText = aboutText & "• Currency Cycling: Ctrl+Alt+Shift+6" & vbCrLf
    aboutText = aboutText & "• Enhanced error handling" & vbCrLf
    aboutText = aboutText & "• Improved Macabacus compatibility" & vbCrLf & vbCrLf
    aboutText = aboutText & "Ribbon Callbacks Version: " & CALLBACKS_VERSION & vbCrLf
    aboutText = aboutText & "Built: " & Date & vbCrLf & vbCrLf
    aboutText = aboutText & "Open source • MIT License • Free forever"
    
    MsgBox aboutText, vbInformation, "About XLerate v2.1.0"
    Debug.Print "About dialog executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("ShowAbout", Err.Description)
End Sub

' =========================================================================
' ENHANCED HELP CALLBACKS
' =========================================================================

Public Sub ShowKeyboardShortcuts(control As IRibbonControl)
    ' NEW: Show comprehensive keyboard shortcuts
    On Error GoTo ErrorHandler
    
    Dim shortcutsText As String
    shortcutsText = "XLerate v2.1.0 - Keyboard Shortcuts" & vbCrLf & vbCrLf
    shortcutsText = shortcutsText & "🚀 CORE SHORTCUTS (Macabacus Compatible):" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+R - Fast Fill Right" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+D - Fast Fill Down (NEW!)" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+E - Error Wrap" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+[ - Pro Precedents" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+] - Pro Dependents" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+A - AutoColor" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+C - Formula Consistency" & vbCrLf & vbCrLf
    shortcutsText = shortcutsText & "📊 FORMAT CYCLING:" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+1 - Number Formats" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+2 - Date Formats" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+3 - Cell Formats" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+4 - Text Styles" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+6 - Currency Formats (NEW!)" & vbCrLf & vbCrLf
    shortcutsText = shortcutsText & "🔧 UTILITIES:" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+S - Quick Save" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+G - Toggle Gridlines" & vbCrLf
    shortcutsText = shortcutsText & "Ctrl+Alt+Shift+Del - Clear Arrows"
    
    MsgBox shortcutsText, vbInformation, "XLerate Keyboard Shortcuts"
    Debug.Print "Keyboard shortcuts help executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("ShowKeyboardShortcuts", Err.Description)
End Sub

' =========================================================================
' UTILITY AND DIAGNOSTIC CALLBACKS
' =========================================================================

Public Sub RunDiagnostics(control As IRibbonControl)
    ' NEW: Run system diagnostics
    On Error GoTo ErrorHandler
    
    Dim diagnostics As String
    diagnostics = "XLerate v2.1.0 - System Diagnostics" & vbCrLf & vbCrLf
    diagnostics = diagnostics & "Ribbon Callbacks Version: " & CALLBACKS_VERSION & vbCrLf
    diagnostics = diagnostics & "Excel Version: " & Application.Version & vbCrLf
    diagnostics = diagnostics & "Platform: " & Application.OperatingSystem & vbCrLf
    diagnostics = diagnostics & "Current Date: " & Now & vbCrLf
    diagnostics = diagnostics & "Active Workbook: " & ActiveWorkbook.Name & vbCrLf
    diagnostics = diagnostics & "Active Sheet: " & ActiveSheet.Name & vbCrLf
    diagnostics = diagnostics & "Selection: " & Selection.Address & vbCrLf & vbCrLf
    
    ' Test key modules
    diagnostics = diagnostics & "MODULE AVAILABILITY:" & vbCrLf
    diagnostics = diagnostics & "• ModFastFillDown: " & TestModuleAvailability("ModFastFillDown") & vbCrLf
    diagnostics = diagnostics & "• ModCurrencyCycling: " & TestModuleAvailability("ModCurrencyCycling") & vbCrLf
    diagnostics = diagnostics & "• ModSmartFillRight: " & TestModuleAvailability("ModSmartFillRight") & vbCrLf
    diagnostics = diagnostics & "• TraceUtils: " & TestModuleAvailability("TraceUtils") & vbCrLf
    diagnostics = diagnostics & "• AutoColorModule: " & TestModuleAvailability("AutoColorModule") & vbCrLf
    
    MsgBox diagnostics, vbInformation, "XLerate Diagnostics"
    Debug.Print "Diagnostics executed via ribbon (v" & CALLBACKS_VERSION & ")"
    Exit Sub
ErrorHandler:
    Call HandleCallbackError("RunDiagnostics", Err.Description)
End Sub

' =========================================================================
' ERROR HANDLING AND UTILITIES
' =========================================================================

Private Sub HandleCallbackError(functionName As String, errorDescription As String)
    ' Centralized error handling for all callbacks
    Debug.Print "RibbonCallback Error in " & functionName & ": " & errorDescription
    
    Dim errorMsg As String
    errorMsg = "XLerate Ribbon Error" & vbCrLf & vbCrLf
    errorMsg = errorMsg & "Function: " & functionName & vbCrLf
    errorMsg = errorMsg & "Error: " & errorDescription & vbCrLf & vbCrLf
    errorMsg = errorMsg & "Callbacks Version: " & CALLBACKS_VERSION & vbCrLf
    errorMsg = errorMsg & "Please check that all required modules are installed."
    
    MsgBox errorMsg, vbExclamation, "XLerate Error"
End Sub

Private Sub UseExistingClearStatusBar()
    ' Helper to use existing ClearStatusBar from ModUtilityFunctions
    ' FIXED: Avoids naming conflicts by using existing function
    On Error Resume Next
    Application.Run "ModUtilityFunctions.ClearStatusBar"
    If Err.Number <> 0 Then
        ' Fallback if ModUtilityFunctions.ClearStatusBar doesn't exist
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        Application.StatusBar = False
    End If
    On Error GoTo 0
End Sub

Private Function TestModuleAvailability(moduleName As String) As String
    ' Test if a VBA module is available
    On Error Resume Next
    
    Dim testModule As Object
    Set testModule = ThisWorkbook.VBProject.VBComponents(moduleName)
    
    If Err.Number = 0 Then
        TestModuleAvailability = "✓ Available"
    Else
        TestModuleAvailability = "✗ Missing"
    End If
    
    On Error GoTo 0
End Function

' =========================================================================
' VERSION INFORMATION
' =========================================================================

Public Function GetCallbacksVersion() As String
    ' Return callbacks version for diagnostics
    GetCallbacksVersion = CALLBACKS_VERSION
End Function

Public Sub RefreshRibbon()
    ' Refresh the ribbon interface
    On Error Resume Next
    If Not myRibbon Is Nothing Then
        myRibbon.Invalidate
        Debug.Print "Ribbon refreshed (v" & CALLBACKS_VERSION & ")"
    End If
    On Error GoTo 0
End Sub