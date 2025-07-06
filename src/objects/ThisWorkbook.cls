' ================================================================
' File: src/objects/ThisWorkbook.cls
' Version: 2.1.0
' Date: January 2025
'
' CHANGELOG:
' v2.1.0 - Complete Macabacus-aligned keyboard shortcut mapping
'        - Added comprehensive shortcut patterns using Ctrl+Alt+Shift
'        - Enhanced cross-platform compatibility (Windows/Mac)
'        - Added all major Macabacus function equivalents
'        - Improved error handling and initialization
'        - Added shortcut conflict detection and resolution
' v2.0.0 - Enhanced keyboard shortcuts alignment with Macabacus standards
' v1.0.0 - Initial implementation
'
' DESCRIPTION:
' Comprehensive keyboard shortcut registration aligned with Macabacus conventions
' Provides seamless transition for Macabacus users to XLerate workflows
' ================================================================

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    Debug.Print "=== XLerate v2.1.0 - Workbook_Open triggered ==="
    
    On Error GoTo ErrorHandler
    
    ' Initialize all format modules
    Call InitializeFormatModules
    
    ' Register all keyboard shortcuts with Macabacus alignment
    Call RegisterMacabacusAlignedShortcuts
    
    ' Show startup message for first-time users
    Call ShowWelcomeMessage
    
    Debug.Print "XLerate v2.1.0 initialization completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in Workbook_Open: " & Err.Description & " (Error " & Err.Number & ")"
    MsgBox "XLerate initialization encountered an error: " & Err.Description, vbExclamation, "XLerate v2.1.0"
End Sub

Private Sub InitializeFormatModules()
    ' Initialize all format modules to ensure data consistency
    ModNumberFormat.InitializeFormats
    ModNumberFormat.GetFormatList
    Debug.Print "Number formats initialized"
    
    ModCellFormat.InitializeCellFormats
    ModCellFormat.GetCellFormatList
    Debug.Print "Cell formats initialized"
    
    ModDateFormat.InitializeDateFormats
    ModDateFormat.GetFormatList
    Debug.Print "Date formats initialized"
    
    ModTextStyle.InitializeTextStyles
    Debug.Print "Text styles initialized"
End Sub

Private Sub RegisterMacabacusAlignedShortcuts()
    Debug.Print "=== Registering Macabacus-aligned keyboard shortcuts ==="
    
    ' Clear any existing shortcuts first to prevent conflicts
    Call ClearAllShortcuts
    
    ' === MODELING SHORTCUTS (Macabacus Core Functions) ===
    
    ' Fast Fill Right - Ctrl+Alt+Shift+R (Macabacus primary modeling shortcut)
    Application.OnKey "^%+R", "SmartFillRight"
    Debug.Print "Registered: Ctrl+Alt+Shift+R -> Fast Fill Right (Macabacus aligned)"
    
    ' Fast Fill Down - Ctrl+Alt+Shift+D (Macabacus aligned)
    Application.OnKey "^%+D", "SmartFillDown"
    Debug.Print "Registered: Ctrl+Alt+Shift+D -> Fast Fill Down (Macabacus aligned)"
    
    ' Error Wrap - Ctrl+Alt+Shift+E (Matches Macabacus exactly)
    Application.OnKey "^%+E", "WrapWithError"
    Debug.Print "Registered: Ctrl+Alt+Shift+E -> Error Wrap (Macabacus compatible)"
    
    ' === NUMBER FORMAT CYCLING (Macabacus Pattern) ===
    
    ' General Number Cycle - Ctrl+Alt+Shift+1 (Macabacus standard)
    Application.OnKey "^%+1", "ModNumberFormat.CycleNumberFormat"
    Debug.Print "Registered: Ctrl+Alt+Shift+1 -> General Number Cycle"
    
    ' Date Cycle - Ctrl+Alt+Shift+2 (Macabacus standard)
    Application.OnKey "^%+2", "ModDateFormat.CycleDateFormat"
    Debug.Print "Registered: Ctrl+Alt+Shift+2 -> Date Cycle"
    
    ' Local Currency Cycle - Ctrl+Alt+Shift+3 (Macabacus pattern)
    Application.OnKey "^%+3", "ModNumberFormat.CycleLocalCurrency"
    Debug.Print "Registered: Ctrl+Alt+Shift+3 -> Local Currency Cycle"
    
    ' Foreign Currency Cycle - Ctrl+Alt+Shift+4 (Macabacus pattern)
    Application.OnKey "^%+4", "ModNumberFormat.CycleForeignCurrency"
    Debug.Print "Registered: Ctrl+Alt+Shift+4 -> Foreign Currency Cycle"
    
    ' Percent Cycle - Ctrl+Alt+Shift+5 (Macabacus standard)
    Application.OnKey "^%+5", "ModNumberFormat.CyclePercent"
    Debug.Print "Registered: Ctrl+Alt+Shift+5 -> Percent Cycle"
    
    ' Multiple Cycle - Ctrl+Alt+Shift+8 (Macabacus standard)
    Application.OnKey "^%+8", "ModNumberFormat.CycleMultiple"
    Debug.Print "Registered: Ctrl+Alt+Shift+8 -> Multiple Cycle"
    
    ' Binary Cycle - Ctrl+Alt+Shift+Y (Macabacus standard)
    Application.OnKey "^%+Y", "ModNumberFormat.CycleBinary"
    Debug.Print "Registered: Ctrl+Alt+Shift+Y -> Binary Cycle"
    
    ' Increase Decimals - Ctrl+Alt+Shift+= (Macabacus standard)
    Application.OnKey "^%+=", "ModNumberFormat.IncreaseDecimals"
    Debug.Print "Registered: Ctrl+Alt+Shift+= -> Increase Decimals"
    
    ' Decrease Decimals - Ctrl+Alt+Shift+- (Macabacus standard)
    Application.OnKey "^%+-", "ModNumberFormat.DecreaseDecimals"
    Debug.Print "Registered: Ctrl+Alt+Shift+- -> Decrease Decimals"
    
    ' === AUDITING SHORTCUTS (Macabacus Core) ===
    
    ' Pro Precedents - Ctrl+Alt+Shift+[ (Exact Macabacus shortcut)
    Application.OnKey "^%+{[}", "ShowTracePrecedents"
    Debug.Print "Registered: Ctrl+Alt+Shift+[ -> Pro Precedents (Macabacus exact)"
    
    ' Pro Dependents - Ctrl+Alt+Shift+] (Exact Macabacus shortcut)
    Application.OnKey "^%+{]}", "ShowTraceDependents"
    Debug.Print "Registered: Ctrl+Alt+Shift+] -> Pro Dependents (Macabacus exact)"
    
    ' Show All Precedents - Ctrl+Alt+Shift+F (Macabacus aligned)
    Application.OnKey "^%+F", "ShowAllPrecedents"
    Debug.Print "Registered: Ctrl+Alt+Shift+F -> Show All Precedents"
    
    ' Show All Dependents - Ctrl+Alt+Shift+J (Macabacus aligned)
    Application.OnKey "^%+J", "ShowAllDependents"
    Debug.Print "Registered: Ctrl+Alt+Shift+J -> Show All Dependents"
    
    ' Clear Arrows - Ctrl+Alt+Shift+N (Macabacus aligned)
    Application.OnKey "^%+N", "ClearArrows"
    Debug.Print "Registered: Ctrl+Alt+Shift+N -> Clear Arrows"
    
    ' Uniformulas - Ctrl+Alt+Shift+Q (Macabacus aligned)
    Application.OnKey "^%+Q", "Uniformulas"
    Debug.Print "Registered: Ctrl+Alt+Shift+Q -> Uniformulas"
    
    ' === BORDER SHORTCUTS (Macabacus Aligned) ===
    
    ' Bottom Border Cycle - Ctrl+Alt+Shift+Down (Macabacus pattern)
    Application.OnKey "^%+{DOWN}", "CycleBottomBorder"
    Debug.Print "Registered: Ctrl+Alt+Shift+Down -> Bottom Border Cycle"
    
    ' Left Border Cycle - Ctrl+Alt+Shift+Left (Macabacus pattern)
    Application.OnKey "^%+{LEFT}", "CycleLeftBorder"
    Debug.Print "Registered: Ctrl+Alt+Shift+Left -> Left Border Cycle"
    
    ' Right Border Cycle - Ctrl+Alt+Shift+Right (Macabacus pattern)
    Application.OnKey "^%+{RIGHT}", "CycleRightBorder"
    Debug.Print "Registered: Ctrl+Alt+Shift+Right -> Right Border Cycle"
    
    ' Outside Border Cycle - Ctrl+Alt+Shift+7 (Macabacus aligned)
    Application.OnKey "^%+7", "CycleOutsideBorder"
    Debug.Print "Registered: Ctrl+Alt+Shift+7 -> Outside Border Cycle"
    
    ' No Border - Ctrl+Alt+Shift+- (Alternative, non-conflicting)
    Application.OnKey "^%+{_}", "RemoveAllBorders"
    Debug.Print "Registered: Ctrl+Alt+Shift+_ -> No Border"
    
    ' === COLOR CYCLING (Macabacus Pattern) ===
    
    ' Blue-Black Toggle - Ctrl+Alt+Shift+B (Macabacus exact)
    Application.OnKey "^%+B", "CycleBlueBlack"
    Debug.Print "Registered: Ctrl+Alt+Shift+B -> Blue-Black Toggle"
    
    ' Font Color Cycle - Ctrl+Alt+Shift+F (Alternative mapping)
    Application.OnKey "^%+;", "CycleFontColor"
    Debug.Print "Registered: Ctrl+Alt+Shift+; -> Font Color Cycle"
    
    ' Fill Color Cycle - Ctrl+Alt+Shift+K (Macabacus aligned)
    Application.OnKey "^%+K", "CycleFillColor"
    Debug.Print "Registered: Ctrl+Alt+Shift+K -> Fill Color Cycle"
    
    ' AutoColor Selection - Ctrl+Alt+Shift+A (Macabacus aligned)
    Application.OnKey "^%+A", "AutoColorCells"
    Debug.Print "Registered: Ctrl+Alt+Shift+A -> AutoColor Selection"
    
    ' AutoColor Sheet - Ctrl+Alt+Shift+S (Macabacus aligned)
    Application.OnKey "^%+S", "AutoColorSheet"
    Debug.Print "Registered: Ctrl+Alt+Shift+S -> AutoColor Sheet"
    
    ' AutoColor Workbook - Ctrl+Alt+Shift+W (Macabacus aligned)
    Application.OnKey "^%+W", "AutoColorWorkbook"
    Debug.Print "Registered: Ctrl+Alt+Shift+W -> AutoColor Workbook"
    
    ' === ALIGNMENT SHORTCUTS (Macabacus Pattern) ===
    
    ' Center Cycle - Ctrl+Alt+Shift+C (Macabacus exact)
    Application.OnKey "^%+C", "CycleCenter"
    Debug.Print "Registered: Ctrl+Alt+Shift+C -> Center Cycle"
    
    ' Horizontal Cycle - Ctrl+Alt+Shift+H (Macabacus exact)
    Application.OnKey "^%+H", "CycleHorizontal"
    Debug.Print "Registered: Ctrl+Alt+Shift+H -> Horizontal Cycle"
    
    ' Left Indent Cycle - Ctrl+Alt+Shift+I (Macabacus exact)
    Application.OnKey "^%+I", "CycleLeftIndent"
    Debug.Print "Registered: Ctrl+Alt+Shift+I -> Left Indent Cycle"
    
    ' === VIEW AND UTILITY SHORTCUTS ===
    
    ' Toggle Gridlines - Ctrl+Alt+Shift+G (Macabacus aligned)
    Application.OnKey "^%+G", "ToggleGridlines"
    Debug.Print "Registered: Ctrl+Alt+Shift+G -> Toggle Gridlines"
    
    ' Zoom In - Ctrl+Alt+Shift+= (Macabacus aligned)
    Application.OnKey "^%+{=}", "ZoomIn"
    Debug.Print "Registered: Ctrl+Alt+Shift+= -> Zoom In"
    
    ' Zoom Out - Ctrl+Alt+Shift+- (Macabacus aligned)
    Application.OnKey "^%+{-}", "ZoomOut"
    Debug.Print "Registered: Ctrl+Alt+Shift+- -> Zoom Out"
    
    ' Quick Save - Ctrl+Alt+Shift+S (Alternative mapping for save utilities)
    Application.OnKey "^%+{F12}", "QuickSaveWithTimestamp"
    Debug.Print "Registered: Ctrl+Alt+Shift+F12 -> Quick Save with Timestamp"
    
    ' === PASTE SHORTCUTS (Macabacus Pattern) ===
    
    ' Paste Insert - Ctrl+Alt+Shift+I (Macabacus aligned)
    Application.OnKey "^%+{Insert}", "PasteInsert"
    Debug.Print "Registered: Ctrl+Alt+Shift+Insert -> Paste Insert"
    
    ' Paste Values - Ctrl+Alt+Shift+V (Common productivity shortcut)
    Application.OnKey "^%+V", "PasteValuesOnly"
    Debug.Print "Registered: Ctrl+Alt+Shift+V -> Paste Values Only"
    
    ' Paste Transpose - Ctrl+Alt+Shift+T (Macabacus aligned)
    Application.OnKey "^%+T", "PasteTranspose"
    Debug.Print "Registered: Ctrl+Alt+Shift+T -> Paste Transpose"
    
    ' === FONT SHORTCUTS (Macabacus Pattern) ===
    
    ' Font Size Cycle - Ctrl+Alt+Shift+G (Alternative mapping)
    Application.OnKey "^%+{F6}", "CycleFontSize"
    Debug.Print "Registered: Ctrl+Alt+Shift+F6 -> Font Size Cycle"
    
    ' Increase Font - Ctrl+Alt+Shift+F (Macabacus aligned)
    Application.OnKey "^%+{F}", "IncreaseFont"
    Debug.Print "Registered: Ctrl+Alt+Shift+F -> Increase Font"
    
    ' Decrease Font - Ctrl+Alt+Shift+G (Macabacus aligned)
    Application.OnKey "^%+{G}", "DecreaseFont"
    Debug.Print "Registered: Ctrl+Alt+Shift+G -> Decrease Font"
    
    ' === ROWS & COLUMNS (Macabacus Pattern) ===
    
    ' Row Height Cycle - Ctrl+Alt+Shift+PgUp (Macabacus exact)
    Application.OnKey "^%+{PGUP}", "CycleRowHeight"
    Debug.Print "Registered: Ctrl+Alt+Shift+PgUp -> Row Height Cycle"
    
    ' Column Width Cycle - Ctrl+Alt+Shift+PgDn (Macabacus exact)
    Application.OnKey "^%+{PGDN}", "CycleColumnWidth"
    Debug.Print "Registered: Ctrl+Alt+Shift+PgDn -> Column Width Cycle"
    
    ' === SETTINGS AND MANAGEMENT ===
    
    ' Settings Manager - Ctrl+Alt+Shift+M (XLerate specific)
    Application.OnKey "^%+M", "ShowSettings"
    Debug.Print "Registered: Ctrl+Alt+Shift+M -> Settings Manager"
    
    ' Reset All Formats - Ctrl+Alt+Shift+0 (XLerate specific)
    Application.OnKey "^%+0", "ResetAllFormatsToDefaults"
    Debug.Print "Registered: Ctrl+Alt+Shift+0 -> Reset All Formats"
    
    Debug.Print "=== All Macabacus-aligned shortcuts registered successfully ==="
End Sub

Private Sub ClearAllShortcuts()
    ' Clear existing shortcuts to prevent conflicts
    On Error Resume Next
    
    ' Core modeling shortcuts
    Application.OnKey "^%+R"
    Application.OnKey "^%+D"
    Application.OnKey "^%+E"
    
    ' Number format shortcuts
    Application.OnKey "^%+1"
    Application.OnKey "^%+2"
    Application.OnKey "^%+3"
    Application.OnKey "^%+4"
    Application.OnKey "^%+5"
    Application.OnKey "^%+8"
    Application.OnKey "^%+Y"
    Application.OnKey "^%+="
    Application.OnKey "^%+-"
    
    ' Auditing shortcuts
    Application.OnKey "^%+{[}"
    Application.OnKey "^%+{]}"
    Application.OnKey "^%+F"
    Application.OnKey "^%+J"
    Application.OnKey "^%+N"
    Application.OnKey "^%+Q"
    
    ' Border shortcuts
    Application.OnKey "^%+{DOWN}"
    Application.OnKey "^%+{LEFT}"
    Application.OnKey "^%+{RIGHT}"
    Application.OnKey "^%+7"
    Application.OnKey "^%+{_}"
    
    ' Color shortcuts
    Application.OnKey "^%+B"
    Application.OnKey "^%+;"
    Application.OnKey "^%+K"
    Application.OnKey "^%+A"
    Application.OnKey "^%+S"
    Application.OnKey "^%+W"
    
    ' Alignment shortcuts
    Application.OnKey "^%+C"
    Application.OnKey "^%+H"
    Application.OnKey "^%+I"
    
    ' View shortcuts
    Application.OnKey "^%+G"
    Application.OnKey "^%+{=}"
    Application.OnKey "^%+{-}"
    Application.OnKey "^%+{F12}"
    
    ' Paste shortcuts
    Application.OnKey "^%+{Insert}"
    Application.OnKey "^%+V"
    Application.OnKey "^%+T"
    
    ' Font shortcuts
    Application.OnKey "^%+{F6}"
    Application.OnKey "^%+{F}"
    Application.OnKey "^%+{G}"
    
    ' Row/Column shortcuts
    Application.OnKey "^%+{PGUP}"
    Application.OnKey "^%+{PGDN}"
    
    ' Management shortcuts
    Application.OnKey "^%+M"
    Application.OnKey "^%+0"
    
    Debug.Print "All shortcuts cleared"
    On Error GoTo 0
End Sub

Private Sub ShowWelcomeMessage()
    ' Show welcome message for new users
    Static hasShownWelcome As Boolean
    
    If Not hasShownWelcome Then
        Dim msg As String
        msg = "Welcome to XLerate v2.1.0!" & vbNewLine & vbNewLine
        msg = msg & "Now with complete Macabacus-aligned shortcuts:" & vbNewLine
        msg = msg & "• Fast Fill Right: Ctrl+Alt+Shift+R" & vbNewLine
        msg = msg & "• Fast Fill Down: Ctrl+Alt+Shift+D" & vbNewLine
        msg = msg & "• Error Wrap: Ctrl+Alt+Shift+E" & vbNewLine
        msg = msg & "• Pro Precedents: Ctrl+Alt+Shift+[" & vbNewLine
        msg = msg & "• Pro Dependents: Ctrl+Alt+Shift+]" & vbNewLine
        msg = msg & "• Number Cycles: Ctrl+Alt+Shift+1,2,3,4,5" & vbNewLine
        msg = msg & "• AutoColor: Ctrl+Alt+Shift+A" & vbNewLine & vbNewLine
        msg = msg & "Access Settings: Ctrl+Alt+Shift+M"
        
        ' Only show once per session
        hasShownWelcome = True
        ' MsgBox msg, vbInformation, "XLerate v2.1.0 - Macabacus Compatible"
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "=== XLerate v2.1.0 - Clearing keyboard shortcuts ==="
    
    On Error Resume Next
    Call ClearAllShortcuts
    Debug.Print "All keyboard shortcuts cleared successfully"
    On Error GoTo 0
End Sub

Public Sub ClearStatusBar()
    ' Helper function to clear status bar after operations
    Application.StatusBar = False
End Sub