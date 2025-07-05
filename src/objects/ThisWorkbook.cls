' ThisWorkbook.cls
' Version: 2.0.0
' Date: 2025-01-04
' Author: XLerate Development Team
' 
' CHANGELOG:
' v2.0.0 - Enhanced keyboard shortcuts alignment with Macabacus standards
'        - Fixed "Constant expression required" error by using variables instead of constants
'        - Added comprehensive shortcut mapping for financial modeling efficiency
'        - Added error handling for shortcut registration
'        - Organized shortcuts by functional categories
' v1.0.0 - Initial implementation
'
' DESCRIPTION:
' Registers comprehensive keyboard shortcuts for Excel financial modeling
' Aligns with Macabacus shortcut conventions for seamless workflow transition

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' In ThisWorkbook
Private Sub Workbook_Open()
    Debug.Print "XLerate v2.0.0 - Workbook_Open triggered"
    
    On Error GoTo ErrorHandler
    
    ' Initialize all format modules
    Call InitializeFormatModules
    
    ' Register all keyboard shortcuts
    Call RegisterKeyboardShortcuts
    
    Debug.Print "XLerate initialization completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in Workbook_Open: " & Err.Description & " (Error " & Err.Number & ")"
    MsgBox "XLerate initialization encountered an error: " & Err.Description, vbExclamation, "XLerate v2.0.0"
End Sub

Private Sub InitializeFormatModules()
    ' Initialize number formats
    ModNumberFormat.InitializeFormats
    ModNumberFormat.GetFormatList
    Debug.Print "Number formats initialized"
    
    ' Initialize cell formats
    ModCellFormat.InitializeCellFormats
    ModCellFormat.GetCellFormatList
    Debug.Print "Cell formats initialized"
    
    ' Initialize date formats
    ModDateFormat.InitializeDateFormats
    ModDateFormat.GetFormatList
    Debug.Print "Date formats initialized"
    
    ' Initialize text styles
    ModTextStyle.InitializeTextStyles
    Debug.Print "Text styles initialized"
End Sub

Private Sub RegisterKeyboardShortcuts()
    Debug.Print "Registering keyboard shortcuts..."
    
    ' === CORE FORMATTING SHORTCUTS (Macabacus-aligned) ===
    
    ' Number Format Cycling - Ctrl+Alt+Shift+1
    Application.OnKey "^%+1", "ModNumberFormat.CycleNumberFormat"
    Debug.Print "Registered: Ctrl+Alt+Shift+1 -> Number Format Cycle"
    
    ' Date Format Cycling - Ctrl+Alt+Shift+2  
    Application.OnKey "^%+2", "ModDateFormat.CycleDateFormat"
    Debug.Print "Registered: Ctrl+Alt+Shift+2 -> Date Format Cycle"
    
    ' Cell Format Cycling - Ctrl+Alt+Shift+3
    Application.OnKey "^%+3", "ModCellFormat.CycleCellFormat"
    Debug.Print "Registered: Ctrl+Alt+Shift+3 -> Cell Format Cycle"
    
    ' Text Style Cycling - Ctrl+Alt+Shift+4
    Application.OnKey "^%+4", "ModTextStyle.CycleTextStyle"
    Debug.Print "Registered: Ctrl+Alt+Shift+4 -> Text Style Cycle"
    
    ' === MODELING SHORTCUTS (Macabacus-aligned) ===
    
    ' Smart Fill Right - Ctrl+Alt+Shift+R (matches Macabacus Fast Fill Right)
    Application.OnKey "^%+R", "SmartFillRight"
    Debug.Print "Registered: Ctrl+Alt+Shift+R -> Smart Fill Right"
    
    ' Smart Fill Down - Ctrl+Alt+Shift+D (new feature)
    Application.OnKey "^%+D", "SmartFillDown"
    Debug.Print "Registered: Ctrl+Alt+Shift+D -> Smart Fill Down"
    
    ' Error Wrap - Ctrl+Alt+Shift+E (matches Macabacus)
    Application.OnKey "^%+E", "WrapWithError"
    Debug.Print "Registered: Ctrl+Alt+Shift+E -> Error Wrap"
    
    ' Switch Sign - Ctrl+Alt+Shift+S (new shortcut)
    Application.OnKey "^%+S", "SwitchCellSign"
    Debug.Print "Registered: Ctrl+Alt+Shift+S -> Switch Sign"
    
    ' === AUDITING SHORTCUTS (Macabacus-aligned) ===
    
    ' Trace Precedents - Ctrl+Alt+Shift+[
    Application.OnKey "^%+{[}", "ShowTracePrecedents"
    Debug.Print "Registered: Ctrl+Alt+Shift+[ -> Trace Precedents"
    
    ' Trace Dependents - Ctrl+Alt+Shift+]
    Application.OnKey "^%+{]}", "ShowTraceDependents"  
    Debug.Print "Registered: Ctrl+Alt+Shift+] -> Trace Dependents"
    
    ' Formula Consistency Check - Ctrl+Alt+Shift+C
    Application.OnKey "^%+C", "CheckHorizontalConsistency"
    Debug.Print "Registered: Ctrl+Alt+Shift+C -> Formula Consistency"
    
    ' === UTILITY SHORTCUTS ===
    
    ' Auto Color Cells - Ctrl+Alt+Shift+A
    Application.OnKey "^%+A", "DoAutoColorCells"
    Debug.Print "Registered: Ctrl+Alt+Shift+A -> Auto Color"
    
    ' Settings Manager - Ctrl+Alt+Shift+M
    Application.OnKey "^%+M", "ShowSettings"
    Debug.Print "Registered: Ctrl+Alt+Shift+M -> Settings Manager"
    
    ' Reset All Formats - Ctrl+Alt+Shift+0
    Application.OnKey "^%+0", "ResetAllFormatsToDefaults"
    Debug.Print "Registered: Ctrl+Alt+Shift+0 -> Reset Formats"
    
    ' === ADDITIONAL PRODUCTIVITY SHORTCUTS ===
    
    ' Quick Save with Timestamp - Ctrl+Alt+Shift+Q
    Application.OnKey "^%+Q", "QuickSaveWithTimestamp"
    Debug.Print "Registered: Ctrl+Alt+Shift+Q -> Quick Save with Timestamp"
    
    ' Insert CAGR Formula - Ctrl+Alt+Shift+G
    Application.OnKey "^%+G", "InsertCAGRFormula"
    Debug.Print "Registered: Ctrl+Alt+Shift+G -> Insert CAGR"
    
    ' Toggle Gridlines - Ctrl+Alt+Shift+L
    Application.OnKey "^%+L", "ToggleGridlines"
    Debug.Print "Registered: Ctrl+Alt+Shift+L -> Toggle Gridlines"
    
    ' Paste Values Only - Ctrl+Alt+Shift+V
    Application.OnKey "^%+V", "PasteValuesOnly"
    Debug.Print "Registered: Ctrl+Alt+Shift+V -> Paste Values Only"
    
    ' === BORDER SHORTCUTS (Macabacus-aligned) ===
    
    ' Bottom Border - Ctrl+Alt+Shift+Down
    Application.OnKey "^%+{DOWN}", "ApplyBottomBorder"
    Debug.Print "Registered: Ctrl+Alt+Shift+Down -> Bottom Border"
    
    ' Top Border - Ctrl+Alt+Shift+Up  
    Application.OnKey "^%+{UP}", "ApplyTopBorder"
    Debug.Print "Registered: Ctrl+Alt+Shift+Up -> Top Border"
    
    ' Left Border - Ctrl+Alt+Shift+Left
    Application.OnKey "^%+{LEFT}", "ApplyLeftBorder"
    Debug.Print "Registered: Ctrl+Alt+Shift+Left -> Left Border"
    
    ' Right Border - Ctrl+Alt+Shift+Right
    Application.OnKey "^%+{RIGHT}", "ApplyRightBorder"
    Debug.Print "Registered: Ctrl+Alt+Shift+Right -> Right Border"
    
    ' Outside Border - Ctrl+Alt+Shift+O
    Application.OnKey "^%+O", "ApplyOutsideBorder"
    Debug.Print "Registered: Ctrl+Alt+Shift+O -> Outside Border"
    
    ' No Border - Ctrl+Alt+Shift+N
    Application.OnKey "^%+N", "RemoveAllBorders"
    Debug.Print "Registered: Ctrl+Alt+Shift+N -> No Border"
    
    Debug.Print "All keyboard shortcuts registered successfully"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "XLerate - Clearing keyboard shortcuts"
    
    On Error Resume Next
    
    ' Clear all registered shortcuts
    Application.OnKey "^%+1"
    Application.OnKey "^%+2"
    Application.OnKey "^%+3"
    Application.OnKey "^%+4"
    Application.OnKey "^%+R"
    Application.OnKey "^%+D"
    Application.OnKey "^%+E"
    Application.OnKey "^%+S"
    Application.OnKey "^%+{[}"
    Application.OnKey "^%+{]}"
    Application.OnKey "^%+C"
    Application.OnKey "^%+A"
    Application.OnKey "^%+M"
    Application.OnKey "^%+0"
    Application.OnKey "^%+Q"
    Application.OnKey "^%+G"
    Application.OnKey "^%+L"
    Application.OnKey "^%+V"
    Application.OnKey "^%+{DOWN}"
    Application.OnKey "^%+{UP}"
    Application.OnKey "^%+{LEFT}"
    Application.OnKey "^%+{RIGHT}"
    Application.OnKey "^%+O"
    Application.OnKey "^%+N"
    
    Debug.Print "Keyboard shortcuts cleared"
    On Error GoTo 0
End Sub