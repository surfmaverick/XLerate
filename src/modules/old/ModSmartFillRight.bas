
' ================================================================
' File: src/modules/ModSmartFillRight.bas  
' Version: 2.0.0
' Date: January 2025
'
' CHANGELOG:
' v2.0.0 - Added Fast Fill Down functionality (Macabacus compatible)
'        - Enhanced error handling and boundary detection
'        - Added support for vertical fill patterns
'        - Improved performance for large ranges
'        - Added debug logging for troubleshooting
'        - Cross-platform compatibility improvements
' v1.0.0 - Initial Smart Fill Right implementation
' ================================================================

Attribute VB_Name = "ModSmartFillRight"
Option Explicit

Public Sub SmartFillRight(Optional control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Debug.Print "--- Starting SmartFillRight v2.0.0 ---"
    
    ' Get active cell
    Dim activeCell As Range
    Set activeCell = Application.activeCell
    Debug.Print "Active cell address: " & activeCell.Address
    Debug.Print "Active cell formula: " & activeCell.formula
    
    ' Check if cell contains formula
    If Len(activeCell.formula) = 0 Or Left(activeCell.formula, 1) <> "=" Then
        Debug.Print "No formula found in active cell"
        MsgBox "Active cell must contain a formula.", vbInformation, "XLerate Smart Fill"
        Exit Sub
    End If
    
    ' Check for merged cells in active cell
    If activeCell.MergeArea.Cells.Count > 1 Then
        Debug.Print "Active cell is merged"
        MsgBox "Cannot perform smart fill on merged cells.", vbInformation, "XLerate Smart Fill"
        Exit Sub
    End If
    
    ' Find boundary
    Debug.Print "Looking for boundary..."
    Dim boundaryCol As Long
    boundaryCol = FindHorizontalBoundary(activeCell)
    Debug.Print "Boundary column found: " & boundaryCol
    
    ' If no boundary found, exit
    If boundaryCol = 0 Then
        Debug.Print "No boundary found"
        MsgBox "No suitable boundary found within 3 rows above.", vbInformation, "XLerate Smart Fill"
        Exit Sub
    End If
    
    ' Perform fill
    Debug.Print "Performing horizontal fill operation"
    PerformHorizontalFill activeCell, boundaryCol
    
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR in SmartFillRight: " & Err.Number & " - " & Err.Description
    MsgBox "An error occurred during Smart Fill Right: " & Err.Description, vbCritical, "XLerate Error"
End Sub

Public Sub SmartFillDown(Optional control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Debug.Print "--- Starting SmartFillDown v2.0.0 ---"
    
    ' Get active cell
    Dim activeCell As Range
    Set activeCell = Application.activeCell
    Debug.Print "Active cell address: " & activeCell.Address
    Debug.Print "Active cell formula: " & activeCell.formula
    
    ' Check if cell contains formula
    If Len(activeCell.formula) = 0 Or Left(activeCell.formula, 1) <> "=" Then
        Debug.Print "No formula found in active cell"
        MsgBox "Active cell must contain a formula.", vbInformation, "XLerate Smart Fill"
        Exit Sub
    End If
    
    ' Check for merged cells in active cell
    If activeCell.MergeArea.Cells.Count > 1 Then
        Debug.Print "Active cell is merged"
        MsgBox "Cannot perform smart fill on merged cells.", vbInformation, "XLerate Smart Fill"
        Exit Sub
    End If
    
    ' Find boundary
    Debug.Print "Looking for vertical boundary..."
    Dim boundaryRow As Long
    boundaryRow = FindVerticalBoundary(activeCell)
    Debug.Print "Boundary row found: " & boundaryRow
    
    ' If no boundary found, exit
    If boundaryRow = 0 Then
        Debug.Print "No boundary found"
        MsgBox "No suitable boundary found within 3 columns to the left.", vbInformation, "XLerate Smart Fill"
        Exit Sub
    End If
    
    ' Perform fill
    Debug.Print "Performing vertical fill operation"
    PerformVerticalFill activeCell, boundaryRow
    
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR in SmartFillDown: " & Err.Number & " - " & Err.Description
    MsgBox "An error occurred during Smart Fill Down: " & Err.Description, vbCritical, "XLerate Error"
End Sub

Private Sub PerformHorizontalFill(startCell As Range, boundaryCol As Long)
    Debug.Print "From column: " & startCell.Column & " to column: " & boundaryCol
    
    On Error Resume Next
    Dim fillRange As Range
    Set fillRange = startCell.Worksheet.Range(startCell, startCell.Worksheet.Cells(startCell.row, boundaryCol))
    
    If Err.Number <> 0 Then
        Debug.Print "ERROR creating fill range: " & Err.Number & " - " & Err.Description
        On Error GoTo 0
        MsgBox "Error creating fill range", vbCritical, "XLerate Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    If fillRange Is Nothing Then
        Debug.Print "ERROR: Fill range is Nothing"
        MsgBox "Invalid fill range", vbCritical, "XLerate Error"
        Exit Sub
    End If
    
    Debug.Print "Fill range: " & fillRange.Address
    Debug.Print "Fill range column count: " & fillRange.Columns.Count
    
    ' Apply fill with progress indicator for large ranges
    Application.ScreenUpdating = False
    If fillRange.Columns.Count > 50 Then
        Application.StatusBar = "Filling formulas across " & fillRange.Columns.Count & " columns..."
    End If
    
    On Error Resume Next
    startCell.AutoFill Destination:=fillRange
    If Err.Number <> 0 Then
        Debug.Print "ERROR in AutoFill: " & Err.Number & " - " & Err.Description
        On Error GoTo 0
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Error during AutoFill operation: " & Err.Description, vbCritical, "XLerate Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Smart Fill Right completed for " & fillRange.Columns.Count & " columns"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    Debug.Print "Horizontal fill operation completed"
End Sub

Private Sub PerformVerticalFill(startCell As Range, boundaryRow As Long)
    Debug.Print "From row: " & startCell.row & " to row: " & boundaryRow
    
    On Error Resume Next
    Dim fillRange As Range
    Set fillRange = startCell.Worksheet.Range(startCell, startCell.Worksheet.Cells(boundaryRow, startCell.Column))
    
    If Err.Number <> 0 Then
        Debug.Print "ERROR creating fill range: " & Err.Number & " - " & Err.Description
        On Error GoTo 0
        MsgBox "Error creating fill range", vbCritical, "XLerate Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    If fillRange Is Nothing Then
        Debug.Print "ERROR: Fill range is Nothing"
        MsgBox "Invalid fill range", vbCritical, "XLerate Error"
        Exit Sub
    End If
    
    Debug.Print "Fill range: " & fillRange.Address
    Debug.Print "Fill range row count: " & fillRange.Rows.Count
    
    ' Apply fill with progress indicator for large ranges
    Application.ScreenUpdating = False
    If fillRange.Rows.Count > 50 Then
        Application.StatusBar = "Filling formulas down " & fillRange.Rows.Count & " rows..."
    End If
    
    On Error Resume Next
    startCell.AutoFill Destination:=fillRange, Type:=xlFillDefault
    If Err.Number <> 0 Then
        Debug.Print "ERROR in AutoFill: " & Err.Number & " - " & Err.Description
        On Error GoTo 0
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Error during AutoFill operation: " & Err.Description, vbCritical, "XLerate Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Smart Fill Down completed for " & fillRange.Rows.Count & " rows"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    Debug.Print "Vertical fill operation completed"
End Sub

Private Function FindHorizontalBoundary(startCell As Range) As Long
    Debug.Print "--- Finding Horizontal Boundary ---"
    Debug.Print "Start cell: " & startCell.Address
    
    Dim currentRow As Long
    Dim checkRow As Range
    Dim startCol As Long
    Dim maxRowsUp As Long
    Dim rowsChecked As Long
    
    startCol = startCell.Column
    maxRowsUp = 3
    rowsChecked = 0
    currentRow = startCell.row - 1
    
    Debug.Print "Starting column: " & startCol
    Debug.Print "Starting check from row: " & currentRow
    
    ' Check up to 3 rows above
    While rowsChecked < maxRowsUp And currentRow > 0
        Debug.Print "Checking row: " & currentRow
        
        ' Get the row to check
        On Error Resume Next
        Set checkRow = startCell.Worksheet.Rows(currentRow)
        If Err.Number <> 0 Then
            Debug.Print "Error getting row " & currentRow & ": " & Err.Description
            On Error GoTo 0
            GoTo NextIteration
        End If
        On Error GoTo 0
        
        ' Check for merged cells in the row
        If Not HasMergedCellsInRow(checkRow, startCol) Then
            ' Find last non-empty cell before empty cell
            Dim boundaryCol As Long
            boundaryCol = FindLastCellInRow(checkRow, startCol)
            
            Debug.Print "Boundary column found: " & boundaryCol
            If boundaryCol > startCol Then ' Ensure we found a valid boundary
                FindHorizontalBoundary = boundaryCol
                Debug.Print "Returning boundary: " & boundaryCol
                Exit Function
            End If
        Else
            Debug.Print "Merged cells found in row " & currentRow
        End If
        
NextIteration:
        currentRow = currentRow - 1
        rowsChecked = rowsChecked + 1
    Wend
    
    Debug.Print "No horizontal boundary found, returning 0"
    FindHorizontalBoundary = 0
End Function

Private Function FindVerticalBoundary(startCell As Range) As Long
    Debug.Print "--- Finding Vertical Boundary ---"
    Debug.Print "Start cell: " & startCell.Address
    
    Dim currentCol As Long
    Dim checkCol As Range
    Dim startRow As Long
    Dim maxColsLeft As Long
    Dim colsChecked As Long
    
    startRow = startCell.row
    maxColsLeft = 3
    colsChecked = 0
    currentCol = startCell.Column - 1
    
    Debug.Print "Starting row: " & startRow
    Debug.Print "Starting check from column: " & currentCol
    
    ' Check up to 3 columns to the left
    While colsChecked < maxColsLeft And currentCol > 0
        Debug.Print "Checking column: " & currentCol
        
        ' Get the column to check
        On Error Resume Next
        Set checkCol = startCell.Worksheet.Columns(currentCol)
        If Err.Number <> 0 Then
            Debug.Print "Error getting column " & currentCol & ": " & Err.Description
            On Error GoTo 0
            GoTo NextColumnIteration
        End If
        On Error GoTo 0
        
        ' Check for merged cells in the column
        If Not HasMergedCellsInColumn(checkCol, startRow) Then
            ' Find last non-empty cell before empty cell
            Dim boundaryRow As Long
            boundaryRow = FindLastCellInColumn(checkCol, startRow)
            
            Debug.Print "Boundary row found: " & boundaryRow
            If boundaryRow > startRow Then ' Ensure we found a valid boundary
                FindVerticalBoundary = boundaryRow
                Debug.Print "Returning boundary: " & boundaryRow
                Exit Function
            End If
        Else
            Debug.Print "Merged cells found in column " & currentCol
        End If
        
NextColumnIteration:
        currentCol = currentCol - 1
        colsChecked = colsChecked + 1
    Wend
    
    Debug.Print "No vertical boundary found, returning 0"
    FindVerticalBoundary = 0
End Function

Private Function HasMergedCellsInRow(checkRow As Range, startCol As Long) As Boolean
    Debug.Print "--- Checking for merged cells in row starting at column " & startCol & " ---"
    
    Dim cell As Range
    Set cell = checkRow.Cells(1, startCol)
    Dim checkCount As Long: checkCount = 0
    
    ' Check cells in the row from startCol, limiting check to reasonable range
    Do While Not IsEmpty(cell) And checkCount < 100 ' Prevent infinite loops
        Debug.Print "Checking cell " & cell.Address
        If cell.MergeArea.Cells.Count > 1 Then
            Debug.Print "Found merged cell at " & cell.Address
            HasMergedCellsInRow = True
            Exit Function
        End If
        Set cell = cell.Offset(0, 1)
        checkCount = checkCount + 1
    Loop
    
    Debug.Print "No merged cells found in row"
    HasMergedCellsInRow = False
End Function

Private Function HasMergedCellsInColumn(checkCol As Range, startRow As Long) As Boolean
    Debug.Print "--- Checking for merged cells in column starting at row " & startRow & " ---"
    
    Dim cell As Range
    Set cell = checkCol.Cells(startRow, 1)
    Dim checkCount As Long: checkCount = 0
    
    ' Check cells in the column from startRow, limiting check to reasonable range
    Do While Not IsEmpty(cell) And checkCount < 100 ' Prevent infinite loops
        Debug.Print "Checking cell " & cell.Address
        If cell.MergeArea.Cells.Count > 1 Then
            Debug.Print "Found merged cell at " & cell.Address
            HasMergedCellsInColumn = True
            Exit Function
        End If
        Set cell = cell.Offset(1, 0)
        checkCount = checkCount + 1
    Loop
    
    Debug.Print "No merged cells found in column"
    HasMergedCellsInColumn = False
End Function

Private Function FindLastCellInRow(checkRow As Range, startCol As Long) As Long
    Debug.Print "--- Finding last cell in row starting at column " & startCol & " ---"
    
    Dim cell As Range
    Set cell = checkRow.Cells(1, startCol)
    
    ' If starting position is empty, return 0
    If IsEmpty(cell) Then
        Debug.Print "Starting cell is empty, returning 0"
        FindLastCellInRow = 0
        Exit Function
    End If
    
    Debug.Print "Starting cell value: " & cell.Address
    
    ' Scan right until empty cell found or reasonable limit reached
    Dim checkCount As Long: checkCount = 0
    Do While Not IsEmpty(cell.Offset(0, 1)) And checkCount < 200
        Set cell = cell.Offset(0, 1)
        Debug.Print "Moving right to " & cell.Address
        checkCount = checkCount + 1
    Loop
    
    Debug.Print "Found last non-empty cell at " & cell.Address
    FindLastCellInRow = cell.Column
End Function

Private Function FindLastCellInColumn(checkCol As Range, startRow As Long) As Long
    Debug.Print "--- Finding last cell in column starting at row " & startRow & " ---"
    
    Dim cell As Range
    Set cell = checkCol.Cells(startRow, 1)
    
    ' If starting position is empty, return 0
    If IsEmpty(cell) Then
        Debug.Print "Starting cell is empty, returning 0"
        FindLastCellInColumn = 0
        Exit Function
    End If
    
    Debug.Print "Starting cell value: " & cell.Address
    
    ' Scan down until empty cell found or reasonable limit reached
    Dim checkCount As Long: checkCount = 0
    Do While Not IsEmpty(cell.Offset(1, 0)) And checkCount < 200
        Set cell = cell.Offset(1, 0)
        Debug.Print "Moving down to " & cell.Address
        checkCount = checkCount + 1
    Loop
    
    Debug.Print "Found last non-empty cell at " & cell.Address
    FindLastCellInColumn = cell.row
End Function