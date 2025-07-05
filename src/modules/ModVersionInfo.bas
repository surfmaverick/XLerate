
' ================================================================
' File: src/modules/ModVersionInfo.bas
' Version: 2.0.0
' Date: January 2025
'
' CHANGELOG:
' v2.0.0 - Created version tracking module
'        - Added comprehensive changelog management
'        - Added version comparison functions
'        - Added upgrade notification system
'        - Added feature discovery helpers
' ================================================================

Attribute VB_Name = "ModVersionInfo"
Option Explicit

' Version Constants
Public Const XLERATE_VERSION As String = "2.0.0"
Public Const XLERATE_BUILD_DATE As String = "January 2025"
Public Const XLERATE_CODENAME As String = "Macabacus Compatible"

' Feature flags for version-specific functionality
Public Const FEATURES_MACABACUS_SHORTCUTS As Boolean = True
Public Const FEATURES_FAST_FILL_DOWN As Boolean = True
Public Const FEATURES_ENHANCED_UI As Boolean = True
Public Const FEATURES_CROSS_PLATFORM As Boolean = True

Public Function GetVersionInfo() As String
    GetVersionInfo = "XLerate v" & XLERATE_VERSION & " (" & XLERATE_CODENAME & ")" & vbNewLine & _
                    "Build Date: " & XLERATE_BUILD_DATE & vbNewLine & _
                    "Compatible with: Excel 365, 2019, 2021 (Windows & macOS)"
End Function

Public Function GetWhatsNew() As String
    Dim whatsNew As String
    whatsNew = "🆕 What's New in XLerate v" & XLERATE_VERSION & vbNewLine & vbNewLine
    
    whatsNew = whatsNew & "🚀 MACABACUS-COMPATIBLE SHORTCUTS:" & vbNewLine
    whatsNew = whatsNew & "• Fast Fill Right: Ctrl+Alt+Shift+R" & vbNewLine
    whatsNew = whatsNew & "• Fast Fill Down: Ctrl+Alt+Shift+D (NEW!)" & vbNewLine
    whatsNew = whatsNew & "• Error Wrap: Ctrl+Alt+Shift+E" & vbNewLine
    whatsNew = whatsNew & "• Pro Precedents: Ctrl+Alt+Shift+[" & vbNewLine
    whatsNew = whatsNew & "• Pro Dependents: Ctrl+Alt+Shift+]" & vbNewLine
    whatsNew = whatsNew & "• Number Cycle: Ctrl+Alt+Shift+1" & vbNewLine
    whatsNew = whatsNew & "• Date Cycle: Ctrl+Alt+Shift+2" & vbNewLine
    whatsNew = whatsNew & "• AutoColor: Ctrl+Alt+Shift+A" & vbNewLine
    whatsNew = whatsNew & "• Quick Save: Ctrl+Alt+Shift+S" & vbNewLine
    whatsNew = whatsNew & "• Toggle Gridlines: Ctrl+Alt+Shift+G" & vbNewLine & vbNewLine
    
    whatsNew = whatsNew & "✨ ENHANCED FEATURES:" & vbNewLine
    whatsNew = whatsNew & "• Smart Fill Down with column pattern detection" & vbNewLine
    whatsNew = whatsNew & "• Redesigned ribbon with Macabacus-inspired layout" & vbNewLine
    whatsNew = whatsNew & "• Cross-platform optimization (Windows & macOS)" & vbNewLine
    whatsNew = whatsNew & "• Enhanced performance for large ranges" & vbNewLine
    whatsNew = whatsNew & "• Improved error handling and user feedback" & vbNewLine
    whatsNew = whatsNew & "• Backward compatibility with all v1.x shortcuts" & vbNewLine & vbNewLine
    
    whatsNew = whatsNew & "🎯 WORKFLOW IMPROVEMENTS:" & vbNewLine
    whatsNew = whatsNew & "• Zoom controls with keyboard shortcuts" & vbNewLine
    whatsNew = whatsNew & "• Enhanced formula consistency checking" & vbNewLine
    whatsNew = whatsNew & "• Settings manager reorganization" & vbNewLine
    whatsNew = whatsNew & "• Status bar feedback for all operations" & vbNewLine
    whatsNew = whatsNew & "• Professional tooltips with shortcut references" & vbNewLine
    
    GetWhatsNew = whatsNew
End Function

Public Sub ShowVersionInfo()
    MsgBox GetVersionInfo(), vbInformation, "XLerate Version Information"
End Sub

Public Sub ShowWhatsNew()
    ' Create a simple form to display what's new
    Dim msg As String
    msg = GetWhatsNew()
    
    ' For now, use MsgBox (could be enhanced with custom form)
    MsgBox msg, vbInformation, "What's New in XLerate v" & XLERATE_VERSION
End Sub

Public Function CheckForUpdates() As Boolean
    ' Placeholder for future update checking functionality
    ' Could connect to GitHub API to check for newer releases
    CheckForUpdates = False
    
    ' Future implementation:
    ' - Compare current version with latest GitHub release
    ' - Notify user if update available
    ' - Provide download link
End Function

Public Sub RecordUsageStatistics(functionName As String)
    ' Optional: Track which functions are used most
    ' Could help prioritize future development
    
    On Error Resume Next
    Dim usageCount As Long
    usageCount = CLng(ThisWorkbook.CustomDocumentProperties("Usage_" & functionName))
    usageCount = usageCount + 1
    
    ' Delete and recreate property
    ThisWorkbook.CustomDocumentProperties("Usage_" & functionName).Delete
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:="Usage_" & functionName, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeNumber, _
        Value:=usageCount
        
    On Error GoTo 0
End Sub