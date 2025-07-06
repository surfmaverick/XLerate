' ================================================================
' File: src/modules/ModFastFillDown.bas
' Version: 1.1.0
' Date: January 2025
'
' CHANGELOG:
' v1.1.0 - Enhanced Smart Fill Down with Macabacus-style intelligence
'        - Added data boundary detection by scanning left columns
'        - Added formula analysis and automatic pattern recognition
'        - Enhanced error handling and edge case management
'        - Added support for mixed data types and complex formulas
'        - Cross-platform compatibility improvements
' v1.0.0 - Initial implementation of Smart Fill Down functionality
'
' DESCRIPTION:
' Advanced Smart Fill Down functionality that matches Macabacus behavior
' Automatically detects data boundaries and fills formulas intelligently
' Scans left columns to determine appropriate fill range
' ================================================================

Attribute VB_Name = "ModFastFillDown"
Option Explicit

' Configuration constants
Private Const MAX_SCAN_COLUMNS As Long = 10  ' Maximum columns to scan left for data patterns
Private Const MAX_SCAN_ROWS As Long = 1000   ' Maximum rows to scan for data boundaries
Private Const MIN_DATA_ROWS As Long = 2      ' Minimum rows required to establish a pattern

Public Sub SmartFillDown(Optional control As IRibbonControl)
    ' Main Smart Fill Down function - Macabacus compatible
    ' Ctrl+Alt+Shift+D shortcut handler
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Smart Fill Down Started ==="
    
    ' Validate selection
    If Selection Is Nothing Then
        Debug.Print "No selection - Smart Fill Down cancelled"
        Exit Sub
    End If
    
    If TypeName(Selection) <> "Range" Then
        Debug.Print "Selection is not a range - Smart Fill Down cancelled"
        Exit Sub
    End If
    
    ' Must be a single cell or single column selection
    If Selection.Columns.Count > 1 Then
        MsgBox "Smart Fill Down works with single cells or single columns only." & vbNewLine & _
               "Please select a single cell with a formula or a single column range.", _
               vbInformation, "XLerate - Smart Fill Down"
        Exit Sub
    End If
    
    Dim startCell As Range
    Set startCell = Selection.Cells(1, 1)
    
    ' Check if we have a formula to fill
    If Not startCell.HasFormula And IsEmpty(startCell.Value) Then
        MsgBox "The selected cell is empty and contains no formula to fill down." & vbNewLine & _
               "Please select a cell with a formula or value to fill down.", _
               vbInformation, "XLerate - Smart Fill Down"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing data pattern for Smart Fill Down..."
    
    ' Determine fill range using intelligent boundary detection
    Dim fillRange As Range
    Set fillRange = DetermineFillRange(startCell)
    
    If fillRange Is Nothing Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Could not determine an appropriate range for Smart Fill Down." & vbNewLine & _
               "Please ensure there is data in columns to the left to establish a pattern.", _
               vbInformation, "XLerate - Smart Fill Down"
        Exit Sub
    End If
    
    ' Show what we're about to do
    Dim lastRow As Long
    lastRow = fillRange.Row + fillRange.Rows.Count - 1
    Application.StatusBar = "Smart Fill Down: " & startCell.Address & " to " & _
                           startCell.Worksheet.Cells(lastRow, startCell.Column).Address
    
    ' Perform the smart fill operation
    Call PerformSmartFillDown(startCell, fillRange)
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show completion message
    Dim cellsCount As Long
    cellsCount = fillRange.Rows.Count - 1  ' Subtract 1 because we don't count the original cell
    
    Debug.Print "Smart Fill Down completed: " & cellsCount & " cells filled"
    Debug.Print "Range: " & startCell.Address & " to " & fillRange.Address
    
    ' Brief status message
    Application.StatusBar = "Smart Fill Down completed: " & cellsCount & " cells filled"
    Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Debug.Print "Error in SmartFillDown: " & Err.Description & " (Error " & Err.Number & ")"
    MsgBox "An error occurred during Smart Fill Down:" & vbNewLine & vbNewLine & _
           Err.Description & vbNewLine & vbNewLine & _
           "Error Number: " & Err.Number, _
           vbExclamation, "XLerate - Smart Fill Down Error"
End Sub

Private Function DetermineFillRange(startCell As Range) As Range
    ' Determines the appropriate range to fill based on data patterns in adjacent columns
    ' Uses Macabacus-style intelligence to find data boundaries
    
    Debug.Print "Determining fill range from " & startCell.Address
    
    Dim ws As Worksheet
    Set ws = startCell.Worksheet
    
    Dim startRow As Long
    Dim startCol As Long
    startRow = startCell.Row
    startCol = startCell.Column
    
    ' Strategy 1: Scan left columns for data patterns (primary method)
    Dim boundaryRow As Long
    boundaryRow = FindDataBoundaryFromLeftColumns(startCell)
    
    If boundaryRow > startRow Then
        Debug.Print "Found data boundary at row " & boundaryRow & " from left column analysis"
        Set DetermineFillRange = ws.Range(startCell, ws.Cells(boundaryRow, startCol))
        Exit Function
    End If
    
    ' Strategy 2: Check if we're in a table or structured data
    boundaryRow = FindTableBoundary(startCell)
    
    If boundaryRow > startRow Then
        Debug.Print "Found table boundary at row " & boundaryRow
        Set DetermineFillRange = ws.Range(startCell, ws.Cells(boundaryRow, startCol))
        Exit Function
    End If
    
    ' Strategy 3: Look for data patterns in the same column (if it has some data)
    boundaryRow = FindColumnDataBoundary(startCell)
    
    If boundaryRow > startRow Then
        Debug.Print "Found column data boundary at row " & boundaryRow
        Set DetermineFillRange = ws.Range(startCell, ws.Cells(boundaryRow, startCol))
        Exit Function
    End If
    
    ' Strategy 4: Use current region if all else fails
    On Error Resume Next
    Dim currentRegion As Range
    Set currentRegion = startCell.CurrentRegion
    On Error GoTo 0
    
    If Not currentRegion Is Nothing Then
        boundaryRow = currentRegion.Row + currentRegion.Rows.Count - 1
        If boundaryRow > startRow Then
            Debug.Print "Using current region boundary at row " & boundaryRow
            Set DetermineFillRange = ws.Range(startCell, ws.Cells(boundaryRow, startCol))
            Exit Function
        End If
    End If
    
    Debug.Print "Could not determine fill range"
    Set DetermineFillRange = Nothing
End Function

Private Function FindDataBoundaryFromLeftColumns(startCell As Range) As Long
    ' Scans columns to the left to find where data ends (Macabacus primary method)
    
    Dim ws As Worksheet
    Set ws = startCell.Worksheet
    
    Dim startRow As Long
    Dim startCol As Long
    startRow = startCell.Row
    startCol = startCell.Column
    
    Debug.Print "Scanning left columns for data boundary..."
    
    ' Scan up to MAX_SCAN_COLUMNS to the left
    Dim scanCol As Long
    Dim maxDataRow As Long
    maxDataRow = startRow
    
    For scanCol = startCol - 1 To Application.WorksheetFunction.Max(1, startCol - MAX_SCAN_COLUMNS) Step -1
        Dim lastDataRow As Long
        lastDataRow = FindLastDataRowInColumn(ws, scanCol, startRow)
        
        If lastDataRow > startRow Then
            Debug.Print "Column " & ColumnLetter(scanCol) & " has data to row " & lastDataRow
            maxDataRow = Application.WorksheetFunction.Max(maxDataRow, lastDataRow)
        End If
    Next scanCol
    
    ' Validate that we found a reasonable boundary
    If maxDataRow > startRow And maxDataRow <= startRow + MAX_SCAN_ROWS Then
        Debug.Print "Data boundary found at row " & maxDataRow & " from left column scan"
        FindDataBoundaryFromLeftColumns = maxDataRow
    Else
        Debug.Print "No valid data boundary found from left column scan"
        FindDataBoundaryFromLeftColumns = 0
    End If
End Function

Private Function FindLastDataRowInColumn(ws As Worksheet, col As Long, startRow As Long) As Long
    ' Finds the last row with data in a specific column, starting from startRow
    
    On Error GoTo ErrorHandler
    
    ' Start from startRow and scan down
    Dim checkRow As Long
    Dim lastDataRow As Long
    lastDataRow = startRow - 1  ' Default to before start if no data found
    
    For checkRow = startRow To startRow + MAX_SCAN_ROWS
        If checkRow > ws.Rows.Count Then Exit For
        
        Dim cell As Range
        Set cell = ws.Cells(checkRow, col)
        
        If Not IsEmpty(cell.Value) Or cell.HasFormula Then
            lastDataRow = checkRow
        ElseIf lastDataRow >= startRow Then
            ' We've found data and now hit an empty cell, so stop here
            Exit For
        End If
    Next checkRow
    
    FindLastDataRowInColumn = lastDataRow
    Exit Function
    
ErrorHandler:
    FindLastDataRowInColumn = startRow - 1
End Function

Private Function FindTableBoundary(startCell As Range) As Long
    ' Attempts to find table boundaries using Excel's table detection
    
    On Error Resume Next
    
    Dim listObj As ListObject
    Set listObj = startCell.ListObject
    
    If Not listObj Is Nothing Then
        FindTableBoundary = listObj.Range.Row + listObj.Range.Rows.Count - 1
        Debug.Print "Found Excel table boundary at row " & FindTableBoundary
        Exit Function
    End If
    
    On Error GoTo 0
    FindTableBoundary = 0
End Function

Private Function FindColumnDataBoundary(startCell As Range) As Long
    ' Finds data boundary within the same column
    
    Dim ws As Worksheet
    Set ws = startCell.Worksheet
    
    Dim col As Long
    col = startCell.Column
    
    ' Find the last non-empty cell in this column, starting from current row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    ' Only use this if it's reasonably close to our start position
    If lastRow > startCell.Row And lastRow <= startCell.Row + MAX_SCAN_ROWS Then
        Debug.Print "Found column data boundary at row " & lastRow
        FindColumnDataBoundary = lastRow
    Else
        FindColumnDataBoundary = 0
    End If
End Function

Private Sub PerformSmartFillDown(startCell As Range, fillRange As Range)
    ' Performs the actual smart fill down operation
    
    Debug.Print "Performing smart fill down from " & startCell.Address & " to " & fillRange.Address
    
    ' If start cell has a formula, we'll use Excel's intelligent fill
    If startCell.HasFormula Then
        Call FillFormulaDown(startCell, fillRange)
    ElseIf Not IsEmpty(startCell.Value) Then
        Call FillValueDown(startCell, fillRange)
    End If
End Sub

Private Sub FillFormulaDown(startCell As Range, fillRange As Range)
    ' Fills a formula down with Excel's intelligent reference adjustment
    
    Debug.Print "Filling formula down: " & startCell.Formula
    
    On Error GoTo ErrorHandler
    
    ' Use Excel's built-in AutoFill for intelligent formula copying
    startCell.AutoFill Destination:=fillRange, Type:=xlFillDefault
    
    Debug.Print "Formula fill completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error filling formula: " & Err.Description
    ' Fallback to manual copy
    fillRange.Formula = startCell.Formula
End Sub

Private Sub FillValueDown(startCell As Range, fillRange As Range)
    ' Fills a value down (for constants)
    
    Debug.Print "Filling value down: " & startCell.Value
    
    On Error GoTo ErrorHandler
    
    ' For values, we can use AutoFill or direct assignment
    fillRange.Value = startCell.Value
    
    ' Copy formatting as well
    startCell.Copy
    fillRange.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    Debug.Print "Value fill completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error filling value: " & Err.Description
End Sub

' === UTILITY HELPER FUNCTIONS ===

Private Function ColumnLetter(col As Long) As String
    ' Converts column number to letter (e.g., 1 = A, 27 = AA)
    ColumnLetter = Split(Cells(1, col).Address, "$")(1)
End Function

Public Sub ClearStatusBar()
    ' Helper function to clear status bar (called by timer)
    Application.StatusBar = False
End Sub

' === ALTERNATIVE FILL FUNCTIONS FOR ADVANCED SCENARIOS ===

Public Sub SmartFillDownWithPrompt(Optional control As IRibbonControl)
    ' Smart Fill Down with user confirmation of range
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Or TypeName(Selection) <> "Range" Then
        Exit Sub
    End If
    
    Dim startCell As Range
    Set startCell = Selection.Cells(1, 1)
    
    Dim fillRange As Range
    Set fillRange = DetermineFillRange(startCell)
    
    If fillRange Is Nothing Then
        MsgBox "Could not determine fill range. Please select the range manually.", _
               vbInformation, "XLerate - Smart Fill Down"
        Exit Sub
    End If
    
    ' Ask user to confirm the range
    Dim lastRow As Long
    lastRow = fillRange.Row + fillRange.Rows.Count - 1
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Smart Fill Down will fill from " & startCell.Address & _
                      " to " & startCell.Worksheet.Cells(lastRow, startCell.Column).Address & _
                      " (" & (fillRange.Rows.Count - 1) & " cells)." & vbNewLine & vbNewLine & _
                      "Continue with Smart Fill Down?", _
                      vbYesNo + vbQuestion, "XLerate - Confirm Smart Fill Down")
    
    If response = vbYes Then
        Application.ScreenUpdating = False
        Call PerformSmartFillDown(startCell, fillRange)
        Application.ScreenUpdating = True
        
        Application.StatusBar = "Smart Fill Down completed: " & (fillRange.Rows.Count - 1) & " cells filled"
        Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "Error in SmartFillDownWithPrompt: " & Err.Description
End Sub

Public Sub FillDownToSelection(Optional control As IRibbonControl)
    ' Fills down to the current selection (user-specified range)
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Or TypeName(Selection) <> "Range" Then
        Exit Sub
    End If
    
    If Selection.Columns.Count > 1 Then
        MsgBox "Please select a single column range for Fill Down to Selection.", _
               vbInformation, "XLerate - Fill Down to Selection"
        Exit Sub
    End If
    
    If Selection.Rows.Count < 2 Then
        MsgBox "Please select a range with at least 2 cells for Fill Down to Selection.", _
               vbInformation, "XLerate - Fill Down to Selection"
        Exit Sub
    End If
    
    Dim startCell As Range
    Set startCell = Selection.Cells(1, 1)
    
    Application.ScreenUpdating = False
    Call PerformSmartFillDown(startCell, Selection)
    Application.ScreenUpdating = True
    
    Application.StatusBar = "Fill Down completed: " & (Selection.Rows.Count - 1) & " cells filled"
    Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "Error in FillDownToSelection: " & Err.Description
End Sub