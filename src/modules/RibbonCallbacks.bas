' RibbonCallbacks.bas
' Version: 2.0.0
' Date: 2025-01-04
' Author: XLerate Development Team
' 
' CHANGELOG:
' v2.0.0 - Enhanced ribbon callbacks for comprehensive functionality
'        - Added border management callbacks
'        - Added productivity utility callbacks
'        - Enhanced error handling and debugging
'        - Aligned with Macabacus workflow patterns
' v1.0.0 - Initial ribbon callback implementation
'
' DESCRIPTION:
' Comprehensive ribbon callback functions for XLerate add-in
' Provides interface between ribbon controls and core functionality

Attribute VB_Name = "RibbonCallbacks"
Option Explicit

' Callback for customUI.onLoad
Public myRibbon As IRibbonUI

' Store ribbon reference
Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
    Debug.Print "XLerate v2.0.0 - Ribbon loaded successfully"
End Sub

' === AUDITING CALLBACKS ===

Public Sub FindAndDisplayPrecedents(control As IRibbonControl)
    ' Callback for Trace Precedents button
    On Error GoTo ErrorHandler
    Application.Run "ShowTracePrecedents"
    Debug.Print "Trace Precedents executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in FindAndDisplayPrecedents: " & Err.Description
End Sub

Public Sub FindAndDisplayDependents(control As IRibbonControl)
    ' Callback for Trace Dependents button
    On Error GoTo ErrorHandler
    Application.Run "ShowTraceDependents"
    Debug.Print "Trace Dependents executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in FindAndDisplayDependents: " & Err.Description
End Sub

Public Sub OnCheckHorizontalConsistency(control As IRibbonControl)
    ' Callback for Horizontal Formula Consistency button
    On Error GoTo ErrorHandler
    Application.Run "CheckHorizontalConsistency"
    Debug.Print "Formula Consistency check executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in OnCheckHorizontalConsistency: " & Err.Description
End Sub

' === FORMAT CYCLING CALLBACKS ===

Public Sub DoCycleNumberFormat(control As IRibbonControl)
    ' Callback for Number Format Cycling
    On Error GoTo ErrorHandler
    Application.Run "ModNumberFormat.CycleNumberFormat"
    Debug.Print "Number format cycle executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in DoCycleNumberFormat: " & Err.Description
End Sub

Public Sub DoCycleCellFormat(control As IRibbonControl)
    ' Callback for Cell Format Cycling
    On Error GoTo ErrorHandler
    Application.Run "ModCellFormat.CycleCellFormat"
    Debug.Print "Cell format cycle executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in DoCycleCellFormat: " & Err.Description
End Sub

Public Sub DoCycleDateFormat(control As IRibbonControl)
    ' Callback for Date Format Cycling
    On Error GoTo ErrorHandler
    Application.Run "ModDateFormat.CycleDateFormat"
    Debug.Print "Date format cycle executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in DoCycleDateFormat: " & Err.Description
End Sub

Public Sub DoCycleTextStyle(control As IRibbonControl)
    ' Callback for Text Style Cycling
    On Error GoTo ErrorHandler
    Application.Run "ModTextStyle.CycleTextStyle"
    Debug.Print "Text style cycle executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in DoCycleTextStyle: " & Err.Description
End Sub

' === MODELING CALLBACKS ===

Public Sub SmartFillRight(control As IRibbonControl)
    ' Callback for Smart Fill Right
    On Error GoTo ErrorHandler
    Application.Run "SmartFillRight"
    Debug.Print "Smart Fill Right executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in SmartFillRight: " & Err.Description
End Sub

Public Sub SmartFillDown(control As IRibbonControl)
    ' Callback for Smart Fill Down
    On Error GoTo ErrorHandler
    Application.Run "SmartFillDown"
    Debug.Print "Smart Fill Down executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in SmartFillDown: " & Err.Description
End Sub

Public Sub SwitchCellSign(control As IRibbonControl)
    ' Callback for Switch Sign
    On Error GoTo ErrorHandler
    Application.Run "SwitchCellSign", control
    Debug.Print "Switch Sign executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in SwitchCellSign: " & Err.Description
End Sub

Public Sub WrapWithError(control As IRibbonControl)
    ' Callback for Error Wrap
    On Error GoTo ErrorHandler
    Application.Run "WrapWithError", control
    Debug.Print "Error Wrap executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in WrapWithError: " & Err.Description
End Sub

Public Sub InsertCAGRFormula(control As IRibbonControl)
    ' Callback for Insert CAGR Formula
    On Error GoTo ErrorHandler
    Application.Run "InsertCAGRFormula"
    Debug.Print "Insert CAGR Formula executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in InsertCAGRFormula: " & Err.Description
End Sub

' === BORDER CALLBACKS ===

Public Sub ApplyBottomBorder(control As IRibbonControl)
    ' Callback for Bottom Border
    On Error GoTo ErrorHandler
    Application.Run "ApplyBottomBorder"
    Debug.Print "Bottom Border applied via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ApplyBottomBorder: " & Err.Description
End Sub

Public Sub ApplyTopBorder(control As IRibbonControl)
    ' Callback for Top Border
    On Error GoTo ErrorHandler
    Application.Run "ApplyTopBorder"
    Debug.Print "Top Border applied via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ApplyTopBorder: " & Err.Description
End Sub

Public Sub ApplyLeftBorder(control As IRibbonControl)
    ' Callback for Left Border
    On Error GoTo ErrorHandler
    Application.Run "ApplyLeftBorder"
    Debug.Print "Left Border applied via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ApplyLeftBorder: " & Err.Description
End Sub

Public Sub ApplyRightBorder(control As IRibbonControl)
    ' Callback for Right Border
    On Error GoTo ErrorHandler
    Application.Run "ApplyRightBorder"
    Debug.Print "Right Border applied via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ApplyRightBorder: " & Err.Description
End Sub

Public Sub ApplyOutsideBorder(control As IRibbonControl)
    ' Callback for Outside Border
    On Error GoTo ErrorHandler
    Application.Run "ApplyOutsideBorder"
    Debug.Print "Outside Border applied via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ApplyOutsideBorder: " & Err.Description
End Sub

Public Sub RemoveAllBorders(control As IRibbonControl)
    ' Callback for No Border
    On Error GoTo ErrorHandler
    Application.Run "RemoveAllBorders"
    Debug.Print "All borders removed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in RemoveAllBorders: " & Err.Description
End Sub

Public Sub ApplyThickBottomBorder(control As IRibbonControl)
    ' Callback for Thick Bottom Border
    On Error GoTo ErrorHandler
    Application.Run "ApplyThickBottomBorder"
    Debug.Print "Thick Bottom Border applied via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ApplyThickBottomBorder: " & Err.Description
End Sub

Public Sub ApplyDoubleBorder(control As IRibbonControl)
    ' Callback for Double Border
    On Error GoTo ErrorHandler
    Application.Run "ApplyDoubleBorder"
    Debug.Print "Double Border applied via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ApplyDoubleBorder: " & Err.Description
End Sub

Public Sub CycleBorderStyle(control As IRibbonControl)
    ' Callback for Border Style Cycling
    On Error GoTo ErrorHandler
    Application.Run "CycleBorderStyle"
    Debug.Print "Border Style cycled via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in CycleBorderStyle: " & Err.Description
End Sub

' === UTILITY CALLBACKS ===

Public Sub PasteValuesOnly(control As IRibbonControl)
    ' Callback for Paste Values Only
    On Error GoTo ErrorHandler
    Application.Run "PasteValuesOnly"
    Debug.Print "Paste Values Only executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in PasteValuesOnly: " & Err.Description
End Sub

Public Sub QuickSaveWithTimestamp(control As IRibbonControl)
    ' Callback for Quick Save with Timestamp
    On Error GoTo ErrorHandler
    Application.Run "QuickSaveWithTimestamp"
    Debug.Print "Quick Save with Timestamp executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in QuickSaveWithTimestamp: " & Err.Description
End Sub

Public Sub ToggleGridlines(control As IRibbonControl)
    ' Callback for Toggle Gridlines
    On Error GoTo ErrorHandler
    Application.Run "ToggleGridlines"
    Debug.Print "Toggle Gridlines executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ToggleGridlines: " & Err.Description
End Sub

Public Sub InsertTimestamp(control As IRibbonControl)
    ' Callback for Insert Timestamp
    On Error GoTo ErrorHandler
    Application.Run "InsertTimestamp"
    Debug.Print "Insert Timestamp executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in InsertTimestamp: " & Err.Description
End Sub

Public Sub ZoomToSelection(control As IRibbonControl)
    ' Callback for Zoom to Selection
    On Error GoTo ErrorHandler
    Application.Run "ZoomToSelection"
    Debug.Print "Zoom to Selection executed via ribbon"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ZoomToSelection: " & Err.Description
End Sub

' === SETTINGS AND MANAGEMENT CALLBACKS ===

Public Sub ShowSettingsForm(control As IRibbonControl)
    ' Callback for Settings Manager
    On Error GoTo ErrorHandler
    Debug.Print "ShowSettingsForm callback triggered via ribbon"
    ShowSettings
    Debug.Print "Settings form displayed successfully"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ShowSettingsForm: " & Err.Description
End Sub

Public Sub ResetAllFormatsToDefaults(control As IRibbonControl)
    ' Callback for Reset All Formats
    On Error GoTo ErrorHandler
    Debug.Print "ResetAllFormatsToDefaults callback triggered via ribbon"
    Application.Run "ResetAllFormatsToDefaults"
    Debug.Print "Format reset completed successfully"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in ResetAllFormatsToDefaults: " & Err.Description
End Sub

Public Sub DoAutoColorCells(control As IRibbonControl)
    ' Callback for Auto Color Cells
    On Error GoTo ErrorHandler
    Debug.Print "DoAutoColorCells callback started via ribbon"
    AutoColorCells control
    Debug.Print "Auto Color Cells completed successfully"
    Exit Sub
ErrorHandler:
    Debug.Print "Error in DoAutoColorCells: " & Err.Description
End Sub

' === RIBBON MANAGEMENT ===

Public Sub RefreshRibbon()
    ' Refresh the entire ribbon - useful after settings changes
    On Error GoTo ErrorHandler
    If Not myRibbon Is Nothing Then
        myRibbon.Invalidate
        Debug.Print "Ribbon refreshed successfully"
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "Error refreshing ribbon: " & Err.Description
End Sub

Public Sub InvalidateControl(controlId As String)
    ' Invalidate a specific control - useful for dynamic updates
    On Error GoTo ErrorHandler
    If Not myRibbon Is Nothing Then
        myRibbon.InvalidateControl controlId
        Debug.Print "Control invalidated: " & controlId
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "Error invalidating control " & controlId & ": " & Err.Description
End Sub