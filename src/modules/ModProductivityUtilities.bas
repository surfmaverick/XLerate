' ModProductivityUtilities.bas
' Version: 1.0.0
' Date: 2025-01-04
' Author: XLerate Development Team
' 
' CHANGELOG:
' v1.0.0 - Initial implementation of productivity utility functions
'        - Quick save with timestamp functionality
'        - Enhanced paste operations
'        - View management utilities
'        - CAGR formula insertion helper
'
' DESCRIPTION:
' Additional productivity utilities to enhance financial modeling workflow
' Complements core XLerate functionality with convenience features

Attribute VB_Name = "ModProductivityUtilities"
Option Explicit

Public Sub QuickSaveWithTimestamp(Optional control As IRibbonControl)
    ' Saves the current workbook with a timestamp appended to filename
    ' Useful for creating quick backups during modeling sessions
    
    On Error GoTo ErrorHandler
    
    Dim originalName As String
    Dim newName As String
    Dim timestamp As String
    Dim baseName As String
    Dim extension As String
    Dim dotPosition As Long
    
    If ActiveWorkbook.Path = "" Then
        MsgBox "Please save the workbook first before using Quick Save with Timestamp.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    originalName = ActiveWorkbook.Name
    dotPosition = InStrRev(originalName, ".")
    
    If dotPosition > 0 Then
        baseName = Left(originalName, dotPosition - 1)
        extension = Mid(originalName, dotPosition)
    Else
        baseName = originalName
        extension = ".xlsx"
    End If
    
    timestamp = Format(Now, "yyyymmdd_hhmmss")
    newName = baseName & "_" & timestamp & extension
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & newName
    Application.DisplayAlerts = True
    
    Debug.Print "Saved workbook as: " & newName
    MsgBox "Workbook saved as: " & newName, vbInformation, "XLerate - Quick Save"
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Debug.Print "Error in QuickSaveWithTimestamp: " & Err.Description
    MsgBox "Error saving workbook: " & Err.Description, vbExclamation, "XLerate"
End Sub

Public Sub PasteValuesOnly(Optional control As IRibbonControl)
    ' Pastes only values (no formulas or formatting)
    ' Essential for financial modeling workflows
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Debug.Print "Pasted values only to " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteValuesOnly: " & Err.Description
    Application.CutCopyMode = False
End Sub

Public Sub ToggleGridlines(Optional control As IRibbonControl)
    ' Toggles gridlines on/off for the active sheet
    ' Useful for presentation and printing
    
    On Error GoTo ErrorHandler
    
    With ActiveWindow
        .DisplayGridlines = Not .DisplayGridlines
    End With
    
    Debug.Print "Toggled gridlines: " & ActiveWindow.DisplayGridlines
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ToggleGridlines: " & Err.Description
End Sub

Public Sub InsertCAGRFormula(Optional control As IRibbonControl)
    ' Inserts a CAGR formula template in the active cell
    ' Prompts user for range selection
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell for the CAGR formula.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Dim rangeAddress As String
    rangeAddress = InputBox("Enter the range for CAGR calculation (e.g., A1:A10):", _
                           "XLerate - Insert CAGR Formula", _
                           "A1:A10")
    
    If rangeAddress = "" Then Exit Sub
    
    ' Validate the range
    Dim testRange As Range
    On Error Resume Next
    Set testRange = Range(rangeAddress)
    On Error GoTo ErrorHandler
    
    If testRange Is Nothing Then
        MsgBox "Invalid range address. Please try again.", vbExclamation, "XLerate"
        Exit Sub
    End If
    
    ' Insert the CAGR formula
    Selection.Formula = "=CAGR(" & rangeAddress & ")"
    
    Debug.Print "Inserted CAGR formula for range: " & rangeAddress
    MsgBox "CAGR formula inserted successfully.", vbInformation, "XLerate"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertCAGRFormula: " & Err.Description
    MsgBox "Error inserting CAGR formula: " & Err.Description, vbExclamation, "XLerate"
End Sub

Public Sub SmartFillDown(Optional control As IRibbonControl)
    ' Smart fill down - similar to SmartFillRight but for vertical filling
    ' Analyzes pattern in columns to the left and fills down accordingly
    
    On Error GoTo ErrorHandler
    
    Debug.Print "--- Starting SmartFillDown ---"
    
    Dim activeCell As Range
    Set activeCell = Application.activeCell
    Debug.Print "Active cell address: " & activeCell.Address
    
    ' Check if cell contains formula
    If Len(activeCell.Formula) = 0 Or Left(activeCell.Formula, 1) <> "=" Then
        Debug.Print "No formula found in active cell"
        MsgBox "Active cell must contain a formula.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    ' Check for merged cells
    If activeCell.MergeArea.Cells.Count > 1 Then
        Debug.Print "Active cell is merged"
        MsgBox "Cannot perform smart fill on merged cells.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    ' Find boundary by looking at columns to the left
    Dim boundaryRow As Long
    boundaryRow = FindVerticalBoundary(activeCell)
    Debug.Print "Boundary row found: " & boundaryRow
    
    If boundaryRow = 0 Then
        Debug.Print "No boundary found"
        MsgBox "No suitable boundary found within 3 columns to the left.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    ' Perform fill
    Debug.Print "Performing vertical fill operation"
    
    Dim fillRange As Range
    Set fillRange = activeCell.Worksheet.Range(activeCell, activeCell.Worksheet.Cells(boundaryRow, activeCell.Column))
    
    Debug.Print "Fill range: " & fillRange.Address
    activeCell.AutoFill Destination:=fillRange
    Debug.Print "Vertical fill operation completed"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in SmartFillDown: " & Err.Description
    MsgBox "An error occurred: " & Err.Description, vbCritical, "XLerate"
End Sub

Private Function FindVerticalBoundary(startCell As Range) As Long
    ' Helper function to find vertical boundary for SmartFillDown
    
    Debug.Print "--- Finding vertical boundary ---"
    
    Dim currentCol As Long
    Dim checkCol As Range
    Dim startRow As Long
    Dim maxColsLeft As Long
    Dim colsChecked As Long
    
    startRow = startCell.Row
    maxColsLeft = 3
    colsChecked = 0
    currentCol = startCell.Column - 1
    
    Debug.Print "Starting row: " & startRow
    Debug.Print "Starting check from column: " & currentCol
    
    ' Check up to 3 columns to the left
    While colsChecked < maxColsLeft And currentCol > 0
        Debug.Print "Checking column: " & currentCol
        
        On Error Resume Next
        Set checkCol = startCell.Worksheet.Columns(currentCol)
        If Err.Number <> 0 Then
            Debug.Print "Error getting column " & currentCol
            GoTo NextColumn
        End If
        On Error GoTo 0
        
        ' Find last non-empty cell in this column starting from startRow
        Dim boundaryRow As Long
        boundaryRow = FindLastCellInColumn(checkCol, startRow)
        
        If boundaryRow > 0 Then
            FindVerticalBoundary = boundaryRow
            Debug.Print "Returning boundary row: " & boundaryRow
            Exit Function
        End If
        
NextColumn:
        currentCol = currentCol - 1
        colsChecked = colsChecked + 1
    Wend
    
    FindVerticalBoundary = 0 ' No boundary found
End Function

Private Function FindLastCellInColumn(checkCol As Range, startRow As Long) As Long
    ' Helper function to find the last non-empty cell in a column
    
    Dim cell As Range
    Set cell = checkCol.Cells(startRow, 1)
    
    ' If starting position is empty, return 0
    If IsEmpty(cell) Then
        FindLastCellInColumn = 0
        Exit Function
    End If
    
    ' Scan down until empty cell found
    Do While Not IsEmpty(cell.Offset(1, 0))
        Set cell = cell.Offset(1, 0)
    Loop
    
    FindLastCellInColumn = cell.Row
End Function

Public Sub ZoomToSelection(Optional control As IRibbonControl)
    ' Zooms the view to fit the current selection perfectly
    ' Useful for focusing on specific model sections
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Application.Goto Selection, True
    
    ' Calculate appropriate zoom level
    Dim zoomLevel As Long
    zoomLevel = 100 ' Default zoom
    
    ' Adjust zoom based on selection size
    If Selection.Cells.Count <= 20 Then
        zoomLevel = 150
    ElseIf Selection.Cells.Count <= 100 Then
        zoomLevel = 125
    ElseIf Selection.Cells.Count <= 500 Then
        zoomLevel = 100
    Else
        zoomLevel = 75
    End If
    
    ActiveWindow.Zoom = zoomLevel
    
    Debug.Print "Zoomed to selection: " & Selection.Address & " at " & zoomLevel & "%"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ZoomToSelection: " & Err.Description
End Sub

Public Sub InsertTimestamp(Optional control As IRibbonControl)
    ' Inserts current timestamp in the active cell
    ' Useful for version tracking and audit trails
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell for the timestamp.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Selection.Value = Now
    Selection.NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    
    Debug.Print "Inserted timestamp in " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertTimestamp: " & Err.Description
End Sub