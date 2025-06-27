VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' Updated ThisWorkbook to match Macabacus keyboard shortcuts
Private Sub Workbook_Open()
    Debug.Print "Workbook_Open triggered - Setting up Macabacus-style shortcuts"
    
    ' Initialize formats
    ModNumberFormat.InitializeFormats
    ModNumberFormat.GetFormatList
    ModCellFormat.InitializeCellFormats
    ModCellFormat.GetCellFormatList
    ModDateFormat.InitializeDateFormats
    ModDateFormat.GetFormatList
    ModTextStyle.InitializeTextStyles
    
    ' === MACABACUS-STYLE KEYBOARD SHORTCUTS ===
    
    ' MODELING SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+R", "SmartFillRight"              ' Fast Fill Right: Ctrl+Alt+Shift+R
    Application.OnKey "^%+D", "SmartFillDown"               ' Fast Fill Down: Ctrl+Alt+Shift+D  
    Application.OnKey "^%+E", "WrapWithError"               ' Error Wrap: Ctrl+Alt+Shift+E
    Application.OnKey "^%+V", "SimplifyFormula"             ' Simplify Formula: Ctrl+Alt+Shift+V
    
    ' AUDITING SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+[", "ShowTracePrecedents"         ' Pro Precedents: Ctrl+Alt+Shift+[
    Application.OnKey "^%+]", "ShowTraceDependents"         ' Pro Dependents: Ctrl+Alt+Shift+]
    Application.OnKey "^%+F", "FormulaConsistency.CheckHorizontalConsistency"  ' Show All Precedents: Ctrl+Alt+Shift+F
    Application.OnKey "^%+J", "ShowAllDependents"           ' Show All Dependents: Ctrl+Alt+Shift+J
    Application.OnKey "^%+N", "ClearArrows"                 ' Clear Arrows: Ctrl+Alt+Shift+N
    Application.OnKey "^%+Q", "Uniformulas"                 ' Uniformulas: Ctrl+Alt+Shift+Q
    
    ' NUMBER FORMATTING SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+1", "ModNumberFormat.CycleNumberFormat"     ' General Number Cycle: Ctrl+Alt+Shift+1
    Application.OnKey "^%+2", "ModDateFormat.CycleDateFormat"         ' Date Cycle: Ctrl+Alt+Shift+2
    Application.OnKey "^%+3", "CycleLocalCurrency"          ' Local Currency Cycle: Ctrl+Alt+Shift+3
    Application.OnKey "^%+4", "CycleForeignCurrency"        ' Foreign Currency Cycle: Ctrl+Alt+Shift+4
    Application.OnKey "^%+5", "CyclePercent"                ' Percent Cycle: Ctrl+Alt+Shift+5
    Application.OnKey "^%+8", "CycleMultiple"               ' Multiple Cycle: Ctrl+Alt+Shift+8
    Application.OnKey "^%+Y", "CycleBinary"                 ' Binary Cycle: Ctrl+Alt+Shift+Y
    Application.OnKey "^%+]", "IncreaseDecimals"            ' Increase Decimals: Ctrl+Alt+Shift+]
    Application.OnKey "^%+[", "DecreaseDecimals"            ' Decrease Decimals: Ctrl+Alt+Shift+[
    
    ' COLOR SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+9", "CycleBlueBlack"              ' Blue-Black Toggle: Ctrl+Alt+Shift+9
    Application.OnKey "^%+C", "CycleFontColor"              ' Font Color Cycle: Ctrl+Alt+Shift+C
    Application.OnKey "^%+K", "CycleFillColor"              ' Fill Color Cycle: Ctrl+Alt+Shift+K
    Application.OnKey "^%+'", "CycleBorderColor"            ' Border Color Cycle: Ctrl+Alt+Shift+'
    Application.OnKey "^%+A", "AutoColorModule.AutoColorCells"  ' AutoColor Selection: Ctrl+Alt+Shift+A
    Application.OnKey "^%+S", "AutoColorSheet"              ' AutoColor Sheet: Ctrl+Alt+Shift+S
    Application.OnKey "^%+Q", "AutoColorWorkbook"           ' AutoColor Workbook: Ctrl+Alt+Shift+Q
    
    ' ALIGNMENT SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+C", "CycleCenter"                 ' Center Cycle: Ctrl+Alt+Shift+C
    Application.OnKey "^%+H", "CycleHorizontal"             ' Horizontal Cycle: Ctrl+Alt+Shift+H
    Application.OnKey "^%+I", "CycleLeftIndent"             ' Left Indent Cycle: Ctrl+Alt+Shift+I
    
    ' BORDER SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+{DOWN}", "CycleBottomBorder"      ' Bottom Border Cycle: Ctrl+Alt+Shift+Down
    Application.OnKey "^%+{LEFT}", "CycleLeftBorder"        ' Left Border Cycle: Ctrl+Alt+Shift+Left
    Application.OnKey "^%+{RIGHT}", "CycleRightBorder"      ' Right Border Cycle: Ctrl+Alt+Shift+Right
    Application.OnKey "^%+7", "CycleOutsideBorder"          ' Outside Border Cycle: Ctrl+Alt+Shift+7
    Application.OnKey "^%+-", "RemoveBorders"               ' No Border: Ctrl+Alt+Shift+-
    
    ' FONT SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+F", "IncreaseFontSize"            ' Increase Font: Ctrl+Alt+Shift+F
    Application.OnKey "^%+G", "DecreaseFontSize"            ' Decrease Font: Ctrl+Alt+Shift+G
    Application.OnKey "^%+U", "CycleUnderline"              ' Underline Cycle: Ctrl+Alt+Shift+U
    Application.OnKey "^%+L", "CycleList"                   ' List Cycle: Ctrl+Alt+Shift+L
    Application.OnKey "^%+L", "CycleLeaderDots"             ' Leader Dots: Ctrl+Alt+Shift+L
    Application.OnKey "^%+6", "CycleFootnote"               ' Footnote Cycle: Ctrl+Alt+Shift+6
    Application.OnKey "^%+W", "ToggleWrapText"              ' Wrap Text: Ctrl+Alt+Shift+W
    
    ' PAINTBRUSH SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+C", "CapturePaintbrushStyle"      ' Capture Paintbrush Style: Ctrl+Alt+Shift+C
    Application.OnKey "^%+P", "ApplyPaintbrushStyle"        ' Apply Paintbrush Style: Ctrl+Alt+Shift+P
    
    ' PASTE SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+I", "PasteInsert"                 ' Paste Insert: Ctrl+Alt+Shift+I
    Application.OnKey "^%+D", "PasteDuplicate"              ' Paste Duplicate: Ctrl+Alt+Shift+D
    Application.OnKey "^%+T", "PasteTranspose"              ' Paste Transpose: Ctrl+Alt+Shift+T
    
    ' VIEW SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+=", "ZoomIn"                      ' Zoom in: Ctrl+Alt+Shift+=
    Application.OnKey "^%+-", "ZoomOut"                     ' Zoom out: Ctrl+Alt+Shift+-
    Application.OnKey "^%+G", "ToggleGridlines"             ' Toggle Gridlines: Ctrl+Alt+Shift+G
    Application.OnKey "^%+B", "HidePageBreaks"              ' Hide Page Breaks: Ctrl+Alt+Shift+B
    
    ' ROWS & COLUMNS SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+{PGUP}", "CycleRowHeight"         ' Row Height Cycle: Ctrl+Alt+Shift+PgUp
    Application.OnKey "^%+{PGDN}", "CycleColumnWidth"       ' Column Width Cycle: Ctrl+Alt+Shift+PgDn
    Application.OnKey "^%+{RIGHT}", "GroupRow"              ' Group Row: Ctrl+Alt+Shift+Right
    Application.OnKey "^%+{DOWN}", "GroupColumn"            ' Group Column: Ctrl+Alt+Shift+Down
    Application.OnKey "^%+{LEFT}", "UngroupRow"             ' Ungroup Row: Ctrl+Alt+Shift+Left
    Application.OnKey "^%+{UP}", "UngroupColumn"            ' Ungroup Column: Ctrl+Alt+Shift+Up
    Application.OnKey "^%+=", "ExpandAllRows"               ' Expand All Rows: Ctrl+Alt+Shift+=
    Application.OnKey "^%+=", "ExpandAllColumns"            ' Expand All Columns: Ctrl+Alt+Shift+=
    Application.OnKey "^%+-", "CollapseAllRows"             ' Collapse All Rows: Ctrl+Alt+Shift+-
    Application.OnKey "^%+-", "CollapseAllColumns"          ' Collapse All Columns: Ctrl+Alt+Shift+-
    
    ' EXPORT SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+{LEFT}", "ExportMatchWidth"       ' Export Match Width: Ctrl+Alt+Shift+Left
    Application.OnKey "^%+{DOWN}", "ExportMatchHeight"      ' Export Match Height: Ctrl+Alt+Shift+Down
    Application.OnKey "^%+{RIGHT}", "ExportMatchNone"       ' Export Match None: Ctrl+Alt+Shift+Right
    Application.OnKey "^%+{UP}", "ExportMatchBoth"          ' Export Match Both: Ctrl+Alt+Shift+Up
    
    ' UTILITIES SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+S", "QuickSave"                   ' Quick Save: Ctrl+Alt+Shift+S
    Application.OnKey "^%+S", "QuickSaveAll"                ' Quick Save All: Ctrl+Alt+Shift+S
    Application.OnKey "^{F12}", "QuickSaveAs"               ' Quick Save As: Ctrl+Alt+Shift+F12
    Application.OnKey "^%+{F12}", "QuickSaveUp"             ' Quick Save Up: Ctrl+Alt+Shift+F12
    Application.OnKey "^%+D", "DeleteCommentsAndNotes"     ' Delete Comments & Notes: Ctrl+Alt+Shift+D
    
    ' OTHER FORMATTING SHORTCUTS (matching Macabacus)
    Application.OnKey "^%+1", "CycleCustomFormat1"          ' Custom Cycles 1-8: Ctrl+Alt+Shift+1-8
    Application.OnKey "^%+2", "CycleCustomFormat2"
    Application.OnKey "^%+3", "CycleCustomFormat3"
    Application.OnKey "^%+4", "CycleCustomFormat4"
    Application.OnKey "^%+5", "CycleCustomFormat5"
    Application.OnKey "^%+6", "CycleCustomFormat6"
    Application.OnKey "^%+7", "CycleCustomFormat7"
    Application.OnKey "^%+8", "CycleCustomFormat8"
    
    ' CELL FORMAT SHORTCUTS (updated to match Macabacus pattern)
    Application.OnKey "^%+2", "ModCellFormat.CycleCellFormat"   ' Cell Format Cycle: Ctrl+Alt+Shift+2
    
    ' TEXT STYLE SHORTCUTS (updated to match Macabacus pattern)
    Application.OnKey "^%+T", "ModTextStyle.CycleTextStyle"     ' Text Style Cycle: Ctrl+Alt+Shift+T
    
    ' RESET SHORTCUT (keeping this useful feature)
    Application.OnKey "^%+0", "ResetAllFormatsToDefaults"       ' Reset All Formats: Ctrl+Alt+Shift+0
    
    ' SWITCH SIGN SHORTCUT (keeping this useful feature)
    Application.OnKey "^%+~", "SwitchCellSign"                  ' Switch Sign: Ctrl+Alt+Shift+~
    
    Debug.Print "All Macabacus-style shortcuts registered successfully"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "Clearing all keyboard shortcuts"
    
    ' Clear all the shortcuts
    Application.OnKey "^%+R"
    Application.OnKey "^%+D"
    Application.OnKey "^%+E"
    Application.OnKey "^%+V"
    Application.OnKey "^%+["
    Application.OnKey "^%+]"
    Application.OnKey "^%+F"
    Application.OnKey "^%+J"
    Application.OnKey "^%+N"
    Application.OnKey "^%+Q"
    Application.OnKey "^%+1"
    Application.OnKey "^%+2"
    Application.OnKey "^%+3"
    Application.OnKey "^%+4"
    Application.OnKey "^%+5"
    Application.OnKey "^%+8"
    Application.OnKey "^%+Y"
    Application.OnKey "^%+9"
    Application.OnKey "^%+C"
    Application.OnKey "^%+K"
    Application.OnKey "^%+'"
    Application.OnKey "^%+A"
    Application.OnKey "^%+S"
    Application.OnKey "^%+H"
    Application.OnKey "^%+I"
    Application.OnKey "^%+{DOWN}"
    Application.OnKey "^%+{LEFT}"
    Application.OnKey "^%+{RIGHT}"
    Application.OnKey "^%+{UP}"
    Application.OnKey "^%+7"
    Application.OnKey "^%+-"
    Application.OnKey "^%+F"
    Application.OnKey "^%+G"
    Application.OnKey "^%+U"
    Application.OnKey "^%+L"
    Application.OnKey "^%+6"
    Application.OnKey "^%+W"
    Application.OnKey "^%+P"
    Application.OnKey "^%+T"
    Application.OnKey "^%+="
    Application.OnKey "^%+B"
    Application.OnKey "^%+{PGUP}"
    Application.OnKey "^%+{PGDN}"
    Application.OnKey "^{F12}"
    Application.OnKey "^%+{F12}"
    Application.OnKey "^%+0"
    Application.OnKey "^%+~"
    
    Debug.Print "All shortcuts cleared"
End Sub

' Helper functions for missing Macabacus-style functions (implement these as needed)
Private Sub SmartFillDown()
    ' Implement fast fill down functionality
    Debug.Print "Smart Fill Down called"
End Sub

Private Sub SimplifyFormula()
    ' Implement formula simplification
    Debug.Print "Simplify Formula called"
End Sub

Private Sub ShowAllDependents()
    ' Show all dependents functionality
    Debug.Print "Show All Dependents called"
End Sub

Private Sub ClearArrows()
    ' Clear all precedent/dependent arrows
    Range("A1").Worksheet.ClearArrows
    Debug.Print "Clear Arrows called"
End Sub

Private Sub Uniformulas()
    ' Implement uniformulas functionality
    Debug.Print "Uniformulas called"
End Sub

' Add other missing function stubs as needed...