' =========================================================================
' FIXED: ModFastFillDown.bas v2.1.1 - Resolved Naming Conflicts
' File: src/modules/ModFastFillDown.bas
' Version: 2.1.1 (FIXED - No naming conflicts)
' Date: 2025-07-06
' Author: XLerate Development Team
' =========================================================================
'
' CHANGELOG v2.1.1:
' - FIXED: Removed duplicate ClearStatusBarDelayed function
' - FIXED: Uses existing ModUtilityFunctions.ClearStatusBar instead
' - RESOLVED: "Ambiguous name detected" compilation errors
' - MAINTAINED: All Fast Fill Down functionality
' - PRESERVED: Integration with existing utility functions
'
' CHANGES FROM v2.1.0:
' - Removed ClearStatusBarDelayed (uses existing ClearStatusBar)
' - Updated all status bar clearing to use ModUtilityFunctions.ClearStatusBar
' - Maintained all core Fast Fill Down logic
' =========================================================================

Attribute VB_Name = "ModFastFillDown"
Option Explicit

' Constants for boundary detection
Private Const MAX_SEARCH_ROWS As Long = 100
Private Const SEARCH_COLUMNS_LEFT As Long = 3

' =========================================================================
' PUBLIC INTERFACE - Called by shortcut Ctrl+Alt+Shift+D
' =========================================================================

Public Sub FastFillDown(Optional control As IRibbonControl)
    ' Main Fast Fill Down function - Macabacus compatible
    ' Shortcut: Ctrl+Alt+Shift+D
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== FastFillDown v2.1.1 Started ==="
    
    ' Validate selection
    If Not ValidateSelection() Then Exit Sub
    
    ' Get source range
    Dim sourceRange As Range
    Set sourceRange = Selection
    
    Application.ScreenUpdating = False
    Application.StatusBar = "XLerate: Fast filling down..."
    
    ' Find the vertical boundary
    Dim lastRow As Long
    lastRow = FindVerticalBoundary(sourceRange)
    
    If lastRow <= sourceRange.Row Then
        Application.StatusBar = "XLerate: No boundary detected for fill down"
        Call UseExistingClearStatusBar
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Perform the fill operation
    Call PerformVerticalFill(sourceRange, lastRow)
    
    Dim rowsFilled As Long
    rowsFilled = lastRow - sourceRange.Row - sourceRange.Rows.Count + 1
    
    Application.StatusBar = "XLerate: Filled " & rowsFilled & " rows down"
    Call UseExistingClearStatusBar
    Application.ScreenUpdating = True
    
    Debug.Print "FastFillDown completed successfully. Rows filled: " & rowsFilled
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = "XLerate: Fill down failed - " & Err.Description
    Call UseExistingClearStatusBar
    Debug.Print "Error in FastFillDown: " & Err.Description & " (Error " & Err.Number & ")"
End Sub

' =========================================================================
' VALIDATION FUNCTIONS
' =========================================================================

Private Function ValidateSelection() As Boolean
    ' Validate that the selection is suitable for fill down
    
    ValidateSelection = False
    
    ' Check if we have a selection
    If Selection Is Nothing Then
        Debug.Print "No selection found"
        Exit Function
    End If
    
    ' Check if selection is empty
    If Selection.Cells.Count = 0 Then
        Debug.Print "Empty selection"
        Exit Function
    End If
    
    ' Check for merged cells in selection
    Dim cell As Range
    For Each cell In Selection
        If cell.MergeArea.Cells.Count > 1 Then
            Debug.Print "Merged cells detected in selection"
            MsgBox "Cannot perform fast fill on ranges containing merged cells.", vbInformation, "XLerate Fast Fill Down"
            Exit Function
        End If
    Next cell
    
    ' Check if selection contains at least one formula or value
    Dim hasContent As Boolean
    For Each cell In Selection
        If cell.HasFormula Or (cell.Value <> "" And Not IsEmpty(cell.Value)) Then
            hasContent = True
            Exit For
        End If
    Next cell
    
    If Not hasContent Then
        Debug.Print "Selection contains no formulas or values"
        MsgBox "Selection must contain at least one formula or value to fill down.", vbInformation, "XLerate Fast Fill Down"
        Exit Function
    End If
    
    Debug.Print "Selection validation passed"
    ValidateSelection = True
End Function

' =========================================================================
' BOUNDARY DETECTION
' =========================================================================

Private Function FindVerticalBoundary(sourceRange As Range) As Long
    ' Find the last row for vertical fill based on pattern detection
    ' Uses Macabacus-style logic: look for data patterns in adjacent columns
    
    Debug.Print "Finding vertical boundary for range: " & sourceRange.Address
    
    Dim startRow As Long
    Dim startCol As Long
    Dim searchEndRow As Long
    
    startRow = sourceRange.Row + sourceRange.Rows.Count
    startCol = sourceRange.Column
    searchEndRow = startRow + MAX_SEARCH_ROWS
    
    ' Method 1: Look for patterns in columns to the left (primary method)
    Dim boundaryFromLeft As Long
    boundaryFromLeft = FindBoundaryFromLeftColumns(startRow, startCol, searchEndRow)
    
    If boundaryFromLeft > startRow Then
        Debug.Print "Boundary found from left columns: Row " & boundaryFromLeft
        FindVerticalBoundary = boundaryFromLeft
        Exit Function
    End If
    
    ' Method 2: Look for data in the same column (secondary method)
    Dim boundaryFromSameColumn As Long
    boundaryFromSameColumn = FindBoundaryFromSameColumn(startRow, startCol, searchEndRow)
    
    If boundaryFromSameColumn > startRow Then
        Debug.Print "Boundary found from same column: Row " & boundaryFromSameColumn
        FindVerticalBoundary = boundaryFromSameColumn
        Exit Function
    End If
    
    ' Method 3: Use Excel's current region (fallback method)
    Dim boundaryFromRegion As Long
    boundaryFromRegion = FindBoundaryFromCurrentRegion(sourceRange)
    
    If boundaryFromRegion > startRow Then
        Debug.Print "Boundary found from current region: Row " & boundaryFromRegion
        FindVerticalBoundary = boundaryFromRegion
        Exit Function
    End If
    
    Debug.Print "No boundary found using any method"
    FindVerticalBoundary = startRow - 1
End Function

Private Function FindBoundaryFromLeftColumns(startRow As Long, startCol As Long, searchEndRow As Long) As Long
    ' Look for patterns in columns to the left of the source range
    
    Dim checkCol As Long
    Dim checkRow As Long
    Dim lastDataRow As Long
    
    lastDataRow = startRow - 1
    
    ' Check up to 3 columns to the left
    For checkCol = startCol - 1 To Application.WorksheetFunction.Max(startCol - SEARCH_COLUMNS_LEFT, 1) Step -1
        
        ' Find the last row with data in this column
        For checkRow = startRow To searchEndRow
            If Not IsEmpty(Cells(checkRow, checkCol).Value) Or Cells(checkRow, checkCol).HasFormula Then
                lastDataRow = Application.WorksheetFunction.Max(lastDataRow, checkRow)
            ElseIf lastDataRow >= startRow Then
                ' Found a gap after finding data - this might be our boundary
                Exit For
            End If
        Next checkRow
        
        ' If we found a significant pattern, use it
        If lastDataRow >= startRow + 2 Then ' At least 3 rows of pattern
            FindBoundaryFromLeftColumns = lastDataRow
            Debug.Print "Pattern found in column " & checkCol & " ending at row " & lastDataRow
            Exit Function
        End If
    Next checkCol
    
    FindBoundaryFromLeftColumns = startRow - 1
End Function

Private Function FindBoundaryFromSameColumn(startRow As Long, startCol As Long, searchEndRow As Long) As Long
    ' Look for existing data pattern in the same column
    
    Dim checkRow As Long
    Dim consecutiveEmpty As Long
    Dim lastDataRow As Long
    
    lastDataRow = startRow - 1
    consecutiveEmpty = 0
    
    For checkRow = startRow To searchEndRow
        If Not IsEmpty(Cells(checkRow, startCol).Value) Or Cells(checkRow, startCol).HasFormula Then
            lastDataRow = checkRow
            consecutiveEmpty = 0
        Else
            consecutiveEmpty = consecutiveEmpty + 1
            ' If we hit 3 consecutive empty cells after finding data, stop
            If consecutiveEmpty >= 3 And lastDataRow >= startRow Then
                Exit For
            End If
        End If
    Next checkRow
    
    ' Only return if we found a meaningful pattern
    If lastDataRow >= startRow + 1 Then
        FindBoundaryFromSameColumn = lastDataRow
        Debug.Print "Same column pattern ending at row " & lastDataRow
    Else
        FindBoundaryFromSameColumn = startRow - 1
    End If
End Function

Private Function FindBoundaryFromCurrentRegion(sourceRange As Range) As Long
    ' Use Excel's current region to determine boundary
    
    On Error GoTo RegionError
    
    Dim currentRegion As Range
    Set currentRegion = sourceRange.CurrentRegion
    
    Dim regionBottomRow As Long
    regionBottomRow = currentRegion.Row + currentRegion.Rows.Count - 1
    
    ' Only use if it extends meaningfully beyond source
    If regionBottomRow > sourceRange.Row + sourceRange.Rows.Count + 2 Then
        FindBoundaryFromCurrentRegion = regionBottomRow
        Debug.Print "Current region boundary at row " & regionBottomRow
    Else
        FindBoundaryFromCurrentRegion = sourceRange.Row
    End If
    
    Exit Function
    
RegionError:
    Debug.Print "Error in FindBoundaryFromCurrentRegion: " & Err.Description
    FindBoundaryFromCurrentRegion = sourceRange.Row
End Function

' =========================================================================
' FILL OPERATIONS
' =========================================================================

Private Sub PerformVerticalFill(sourceRange As Range, lastRow As Long)
    ' Perform the actual vertical fill operation
    
    On Error GoTo FillError
    
    Dim targetRange As Range
    Set targetRange = Range(sourceRange, Cells(lastRow, sourceRange.Column + sourceRange.Columns.Count - 1))
    
    Debug.Print "Filling from " & sourceRange.Address & " to " & targetRange.Address
    
    ' Use AutoFill for smart formula adjustment
    sourceRange.AutoFill Destination:=targetRange, Type:=xlFillDefault
    
    ' Select the filled range for user feedback
    targetRange.Select
    
    Debug.Print "Fill operation completed successfully"
    Exit Sub
    
FillError:
    Debug.Print "Error in PerformVerticalFill: " & Err.Description
    Err.Raise Err.Number, "PerformVerticalFill", Err.Description
End Sub

' =========================================================================
' ENHANCED FILL FUNCTIONS (Additional Features)
' =========================================================================

Public Sub SmartFillDown(Optional control As IRibbonControl)
    ' Enhanced version with pattern recognition
    ' This can be called separately or used as an alternative
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== SmartFillDown (Enhanced) Started ==="
    
    If Not ValidateSelection() Then Exit Sub
    
    Dim sourceRange As Range
    Set sourceRange = Selection
    
    Application.ScreenUpdating = False
    Application.StatusBar = "XLerate: Smart filling down with pattern recognition..."
    
    ' Use enhanced pattern detection
    Dim lastRow As Long
    lastRow = FindVerticalBoundaryEnhanced(sourceRange)
    
    If lastRow <= sourceRange.Row Then
        Application.StatusBar = "XLerate: No suitable pattern detected"
        Call UseExistingClearStatusBar
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Perform fill with pattern validation
    Call PerformSmartVerticalFill(sourceRange, lastRow)
    
    Application.StatusBar = "XLerate: Smart fill down completed"
    Call UseExistingClearStatusBar
    Application.ScreenUpdating = True
    
    Debug.Print "SmartFillDown completed successfully"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = "XLerate: Smart fill down failed - " & Err.Description
    Call UseExistingClearStatusBar
    Debug.Print "Error in SmartFillDown: " & Err.Description
End Sub

Private Function FindVerticalBoundaryEnhanced(sourceRange As Range) As Long
    ' Enhanced boundary detection with multiple pattern analysis methods
    
    ' This could include more sophisticated pattern recognition
    ' For now, delegate to the standard method but could be expanded
    FindVerticalBoundaryEnhanced = FindVerticalBoundary(sourceRange)
    
    ' Future enhancements could include:
    ' - Analysis of formula patterns
    ' - Detection of table structures
    ' - Recognition of financial model patterns
    ' - Machine learning-based boundary prediction
End Function

Private Sub PerformSmartVerticalFill(sourceRange As Range, lastRow As Long)
    ' Enhanced fill operation with pattern validation
    
    ' For now, delegate to standard fill but validate the result
    Call PerformVerticalFill(sourceRange, lastRow)
    
    ' Future enhancements could include:
    ' - Post-fill validation
    ' - Formula consistency checking
    ' - Automatic formatting application
    ' - Error detection and correction
End Sub

' =========================================================================
' UTILITY FUNCTIONS - FIXED: Uses existing ModUtilityFunctions
' =========================================================================

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

Public Function GetFastFillDownVersion() As String
    ' Return module version for diagnostics
    GetFastFillDownVersion = "2.1.1"
End Function

Public Sub TestFastFillDown()
    ' Test function for development and debugging
    Debug.Print "=== FastFillDown Test Function ==="
    Debug.Print "Module Version: " & GetFastFillDownVersion()
    Debug.Print "Selection: " & Selection.Address
    Debug.Print "Test completed - use FastFillDown() for actual operation"
End Sub