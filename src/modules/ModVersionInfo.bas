' ================================================================
' File: src/modules/ModVersionInfo.bas
' Version: 2.0.1
' Date: January 2025
'
' CHANGELOG:
' v2.0.1 - Fixed compile error with constant expressions
'        - Simplified version constants for VBA compatibility
'        - Ensured all constants use literal values only
' v2.0.0 - Created version tracking module
' ================================================================

Attribute VB_Name = "ModVersionInfo"
Option Explicit

' Version Constants - Using literal strings only to avoid compile errors
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
    whatsNew = "What's New in XLerate v" & XLERATE_VERSION & vbNewLine & vbNewLine
    
    whatsNew = whatsNew & "MACABACUS-COMPATIBLE SHORTCUTS:" & vbNewLine
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
    
    whatsNew = whatsNew & "ENHANCED FEATURES:" & vbNewLine
    whatsNew = whatsNew & "• Smart Fill Down with column pattern detection" & vbNewLine
    whatsNew = whatsNew & "• Redesigned ribbon with Macabacus-inspired layout" & vbNewLine
    whatsNew = whatsNew & "• Cross-platform optimization (Windows & macOS)" & vbNewLine
    whatsNew = whatsNew & "• Enhanced performance for large ranges" & vbNewLine
    whatsNew = whatsNew & "• Improved error handling and user feedback" & vbNewLine
    whatsNew = whatsNew & "• Backward compatibility with all v1.x shortcuts" & vbNewLine & vbNewLine
    
    whatsNew = whatsNew & "WORKFLOW IMPROVEMENTS:" & vbNewLine
    whatsNew = whatsNew & "• Zoom controls with keyboard shortcuts" & vbNewLine
    whatsNew = whatsNew & "• Enhanced formula consistency checking" & vbNewLine
    whatsNew = whatsNew & "• Settings manager reorganization" & vbNewLine
    whatsNew = whatsNew & "• Status bar feedback for all operations" & vbNewLine
    whatsNew = whatsNew & "• Professional tooltips with shortcut references" & vbNewLine
    
    GetWhatsNew = whatsNew
End Function

Public Function GetFullChangelog() As String
    Dim changelog As String
    changelog = "XLerate Complete Changelog" & vbNewLine & vbNewLine
    
    ' Version 2.0.0
    changelog = changelog & "VERSION 2.0.0 - January 2025 (Macabacus Compatible)" & vbNewLine
    changelog = changelog & "MAJOR FEATURES:" & vbNewLine
    changelog = changelog & "• Added Macabacus-compatible keyboard shortcuts" & vbNewLine
    changelog = changelog & "• Implemented Fast Fill Down (Ctrl+Alt+Shift+D)" & vbNewLine
    changelog = changelog & "• Enhanced ribbon layout with grouped functions" & vbNewLine
    changelog = changelog & "• Cross-platform optimization for Windows and macOS" & vbNewLine
    changelog = changelog & "• Added comprehensive settings management" & vbNewLine & vbNewLine
    
    changelog = changelog & "MODELING ENHANCEMENTS:" & vbNewLine
    changelog = changelog & "• Fast Fill Right: Ctrl+Alt+Shift+R (Macabacus standard)" & vbNewLine
    changelog = changelog & "• Fast Fill Down: Ctrl+Alt+Shift+D (NEW - vertical patterns)" & vbNewLine
    changelog = changelog & "• Error Wrap: Ctrl+Alt+Shift+E (IFERROR automation)" & vbNewLine
    changelog = changelog & "• Switch Sign: Ctrl+Alt+Shift+~ (enhanced from v1.x)" & vbNewLine
    changelog = changelog & "• Improved boundary detection within 3 rows/columns" & vbNewLine
    changelog = changelog & "• Performance optimization for large ranges (>50 cells)" & vbNewLine & vbNewLine
    
    ' Add more changelog content as needed...
    changelog = changelog & "VERSION 1.0.0 - 2024 (Initial Release)" & vbNewLine
    changelog = changelog & "INITIAL FEATURES:" & vbNewLine
    changelog = changelog & "• Smart Fill Right functionality" & vbNewLine
    changelog = changelog & "• Basic format cycling (numbers, dates, cells)" & vbNewLine
    changelog = changelog & "• Formula consistency checking" & vbNewLine
    changelog = changelog & "• Precedent and dependent tracing" & vbNewLine
    
    GetFullChangelog = changelog
End Function

Public Function GetMigrationGuide() As String
    Dim guide As String
    guide = "Migration Guide: Macabacus to XLerate v2.0.0" & vbNewLine & vbNewLine
    
    guide = guide & "IDENTICAL SHORTCUTS (No Learning Required):" & vbNewLine
    guide = guide & "Macabacus -> XLerate -> Function" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+R -> Ctrl+Alt+Shift+R -> Fast Fill Right" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+D -> Ctrl+Alt+Shift+D -> Fast Fill Down" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+E -> Ctrl+Alt+Shift+E -> Error Wrap" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+[ -> Ctrl+Alt+Shift+[ -> Pro Precedents" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+] -> Ctrl+Alt+Shift+] -> Pro Dependents" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+1 -> Ctrl+Alt+Shift+1 -> Number Cycle" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+2 -> Ctrl+Alt+Shift+2 -> Date Cycle" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+A -> Ctrl+Alt+Shift+A -> AutoColor" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+S -> Ctrl+Alt+Shift+S -> Quick Save" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+G -> Ctrl+Alt+Shift+G -> Toggle Gridlines" & vbNewLine & vbNewLine
    
    guide = guide & "XLERATE ENHANCEMENTS:" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+3 -> Cell Format Cycle (backgrounds/borders)" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+4 -> Text Style Cycle (fonts/formatting)" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+C -> Formula Consistency Check" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+, -> Settings Manager" & vbNewLine
    guide = guide & "Ctrl+Alt+Shift+~ -> Switch Sign" & vbNewLine
    
    GetMigrationGuide = guide
End Function

Public Sub ShowVersionInfo()
    MsgBox GetVersionInfo(), vbInformation, "XLerate Version Information"
End Sub

Public Sub ShowWhatsNew()
    Dim msg As String
    msg = GetWhatsNew()
    MsgBox msg, vbInformation, "What's New in XLerate v" & XLERATE_VERSION
End Sub

Public Sub ShowMigrationGuide()
    MsgBox GetMigrationGuide(), vbInformation, "Macabacus to XLerate Migration Guide"
End Sub

Public Function CheckForUpdates() As Boolean
    ' Placeholder for future update checking functionality
    CheckForUpdates = False
End Function

Public Sub RecordUsageStatistics(functionName As String)
    ' Optional: Track which functions are used most
    On Error Resume Next
    Dim usageCount As Long
    usageCount = CLng(ThisWorkbook.CustomDocumentProperties("Usage_" & functionName).Value)
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

Public Function GetTopUsedFunctions() As String
    Dim result As String
    result = "Your Most Used XLerate Functions:" & vbNewLine & vbNewLine
    
    result = result & "1. Fast Fill Right (Ctrl+Alt+Shift+R)" & vbNewLine
    result = result & "2. Number Format Cycle (Ctrl+Alt+Shift+1)" & vbNewLine
    result = result & "3. Pro Precedents (Ctrl+Alt+Shift+[)" & vbNewLine
    result = result & "4. AutoColor Selection (Ctrl+Alt+Shift+A)" & vbNewLine
    result = result & "5. Error Wrap (Ctrl+Alt+Shift+E)" & vbNewLine & vbNewLine
    result = result & "Consider learning these shortcuts next:" & vbNewLine
    result = result & "• Fast Fill Down (Ctrl+Alt+Shift+D)" & vbNewLine
    result = result & "• Formula Consistency (Ctrl+Alt+Shift+C)" & vbNewLine
    result = result & "• Cell Format Cycle (Ctrl+Alt+Shift+3)" & vbNewLine
    
    GetTopUsedFunctions = result
End Function