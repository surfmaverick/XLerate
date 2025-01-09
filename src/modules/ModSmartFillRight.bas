Attribute VB_Name = "ModSmartFillRight"
Option Explicit

Public Sub SmartFillRight(Optional control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Debug.Print "--- Starting SmartFillRight ---"
    
    ' Get active cell
    Dim activeCell As Range
    Set activeCell = Application.activeCell
    Debug.Print "Active cell address: " & activeCell.Address
    Debug.Print "Active cell formula: " & activeCell.formula
    
    ' Check if cell contains formula
    If Len(activeCell.formula) = 0 Or Left(activeCell.formula, 1) <> "=" Then
        Debug.Print "No formula found in active cell"
        MsgBox "Active cell must contain a formula.", vbInformation
        Exit Sub
    End If
    
    ' Check for merged cells in active cell
    If activeCell.MergeArea.Cells.Count > 1 Then
        Debug.Print "Active cell is merged"
        MsgBox "Cannot perform smart fill on merged cells.", vbInformation
        Exit Sub
    End If
    
    ' Find boundary
    Debug.Print "Looking for boundary..."
    Dim boundaryCol As Long
    boundaryCol = FindBoundary(activeCell)
    Debug.Print "Boundary column found: " & boundaryCol
    
    ' If no boundary found, exit
    If boundaryCol = 0 Then
        Debug.Print "No boundary found"
        MsgBox "No suitable boundary found within 3 rows above.", vbInformation
        Exit Sub
    End If
    
    ' Perform fill
    Debug.Print "Performing fill operation"
    Debug.Print "From column: " & activeCell.Column & " to column: " & boundaryCol
    
    On Error Resume Next
    Dim fillRange As Range
    Set fillRange = activeCell.Worksheet.Range(activeCell, activeCell.Worksheet.Cells(activeCell.row, boundaryCol))
    
    If Err.Number <> 0 Then
        Debug.Print "ERROR creating fill range: " & Err.Number & " - " & Err.Description
        On Error GoTo 0
        MsgBox "Error creating fill range", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    If fillRange Is Nothing Then
        Debug.Print "ERROR: Fill range is Nothing"
        MsgBox "Invalid fill range", vbCritical
        Exit Sub
    End If
    
    Debug.Print "Fill range: " & fillRange.Address
    Debug.Print "Fill range row: " & fillRange.row
    Debug.Print "Fill range column count: " & fillRange.Columns.Count
    Debug.Print "Starting cell value: " & activeCell.formula
    
    On Error Resume Next
    activeCell.AutoFill Destination:=fillRange
    If Err.Number <> 0 Then
        Debug.Print "ERROR in AutoFill: " & Err.Number & " - " & Err.Description
        On Error GoTo 0
        MsgBox "Error during AutoFill operation: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    Debug.Print "Fill operation completed"
    
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR in SmartFillRight: " & Err.Number & " - " & Err.Description
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Private Function FindBoundary(startCell As Range) As Long
    Debug.Print "--- Starting FindBoundary ---"
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
        If checkRow Is Nothing Then
            Debug.Print "Row " & currentRow & " is Nothing"
            On Error GoTo 0
            GoTo NextIteration
        End If
        On Error GoTo 0
        
        ' Check for merged cells in the row
        Debug.Print "Checking for merged cells"
        If Not HasMergedCells(checkRow, startCol) Then
            Debug.Print "No merged cells found, looking for last cell"
            ' Find last non-empty cell before empty cell
            Dim boundaryCol As Long
            boundaryCol = FindLastCellInRow(checkRow, startCol)
            
            Debug.Print "Boundary column found: " & boundaryCol
            If boundaryCol > 0 Then
                FindBoundary = boundaryCol
                Debug.Print "Returning boundary: " & boundaryCol
                Exit Function
            End If
        Else
            Debug.Print "Merged cells found in row " & currentRow
        End If
        
NextIteration:
        currentRow = currentRow - 1
        rowsChecked = rowsChecked + 1
        Debug.Print "Moving to next row. rowsChecked: " & rowsChecked
    Wend
    
    Debug.Print "No boundary found, returning 0"
    FindBoundary = 0 ' No boundary found
End Function

Private Function HasMergedCells(checkRow As Range, startCol As Long) As Boolean
    Debug.Print "--- Checking for merged cells starting at column " & startCol & " ---"
    
    Dim cell As Range
    Set cell = checkRow.Cells(1, startCol)
    
    ' Check if any cell in the row from startCol is merged
    Do While Not IsEmpty(cell)
        Debug.Print "Checking cell " & cell.Address
        If cell.MergeArea.Cells.Count > 1 Then
            Debug.Print "Found merged cell at " & cell.Address
            HasMergedCells = True
            Exit Function
        End If
        Set cell = cell.Offset(0, 1)
    Loop
    
    Debug.Print "No merged cells found"
    HasMergedCells = False
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
    
    ' Scan right until empty cell found
    Do While Not IsEmpty(cell.Offset(0, 1))
        Set cell = cell.Offset(0, 1)
        Debug.Print "Moving right to " & cell.Address
    Loop
    
    Debug.Print "Found last non-empty cell at " & cell.Address
    FindLastCellInRow = cell.Column
End Function
