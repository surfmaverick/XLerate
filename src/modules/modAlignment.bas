' =============================================================================
' File: ModAlignment.bas
' Version: 2.0.0
' Description: Alignment and layout functions with Macabacus-style cycling
' Author: XLerate Development Team
' Created: New module for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

Attribute VB_Name = "ModAlignment"
' Alignment and Layout Functions (Macabacus-style)
Option Explicit

' === CENTER CYCLE FUNCTION ===

Public Sub CycleCenter(Optional control As IRibbonControl)
    Debug.Print "CycleCenter called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Select Case Selection.HorizontalAlignment
        Case xlLeft
            Selection.HorizontalAlignment = xlCenter
            Debug.Print "Changed to Center alignment"
        Case xlCenter
            Selection.HorizontalAlignment = xlRight
            Debug.Print "Changed to Right alignment"
        Case xlRight
            Selection.HorizontalAlignment = xlLeft
            Debug.Print "Changed to Left alignment"
        Case Else
            Selection.HorizontalAlignment = xlCenter
            Debug.Print "Set to Center alignment (default)"
    End Select
    On Error GoTo 0
End Sub

' === HORIZONTAL ALIGNMENT CYCLE ===

Public Sub CycleHorizontal(Optional control As IRibbonControl)
    Debug.Print "CycleHorizontal called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Select Case Selection.HorizontalAlignment
        Case xlGeneral
            Selection.HorizontalAlignment = xlLeft
            Debug.Print "Changed to Left alignment"
        Case xlLeft
            Selection.HorizontalAlignment = xlCenter
            Debug.Print "Changed to Center alignment"
        Case xlCenter
            Selection.HorizontalAlignment = xlRight
            Debug.Print "Changed to Right alignment"
        Case xlRight
            Selection.HorizontalAlignment = xlJustify
            Debug.Print "Changed to Justify alignment"
        Case xlJustify
            Selection.HorizontalAlignment = xlGeneral
            Debug.Print "Changed to General alignment"
        Case Else
            Selection.HorizontalAlignment = xlLeft
            Debug.Print "Set to Left alignment (default)"
    End Select
    On Error GoTo 0
End Sub

' === VERTICAL ALIGNMENT CYCLE ===

Public Sub CycleVertical(Optional control As IRibbonControl)
    Debug.Print "CycleVertical called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Select Case Selection.VerticalAlignment
        Case xlTop
            Selection.VerticalAlignment = xlCenter
            Debug.Print "Changed to Middle alignment"
        Case xlCenter
            Selection.VerticalAlignment = xlBottom
            Debug.Print "Changed to Bottom alignment"
        Case xlBottom
            Selection.VerticalAlignment = xlJustify
            Debug.Print "Changed to Justify alignment"
        Case xlJustify
            Selection.VerticalAlignment = xlTop
            Debug.Print "Changed to Top alignment"
        Case Else
            Selection.VerticalAlignment = xlTop
            Debug.Print "Set to Top alignment (default)"
    End Select
    On Error GoTo 0
End Sub

' === LEFT INDENT CYCLE ===

Public Sub CycleLeftIndent(Optional control As IRibbonControl)
    Debug.Print "CycleLeftIndent called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentIndent As Integer
    currentIndent = Selection.IndentLevel
    
    ' Cycle through indent levels 0, 1, 2, 3, then back to 0
    Dim nextIndent As Integer
    nextIndent = (currentIndent + 1) Mod 4
    
    Selection.IndentLevel = nextIndent
    Debug.Print "Changed indent level from " & currentIndent & " to " & nextIndent
    On Error GoTo 0
End Sub

' === TEXT ORIENTATION CYCLE ===

Public Sub CycleTextOrientation(Optional control As IRibbonControl)
    Debug.Print "CycleTextOrientation called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentOrientation As Integer
    currentOrientation = Selection.Orientation
    
    ' Cycle through orientations: 0° (horizontal), 90° (vertical up), -90° (vertical down), 45°, -45°
    Select Case currentOrientation
        Case 0
            Selection.Orientation = 90
            Debug.Print "Changed to 90° orientation"
        Case 90
            Selection.Orientation = -90
            Debug.Print "Changed to -90° orientation"
        Case -90
            Selection.Orientation = 45
            Debug.Print "Changed to 45° orientation"
        Case 45
            Selection.Orientation = -45
            Debug.Print "Changed to -45° orientation"
        Case -45
            Selection.Orientation = 0
            Debug.Print "Changed to 0° orientation"
        Case Else
            Selection.Orientation = 0
            Debug.Print "Reset to 0° orientation"
    End Select
    On Error GoTo 0
End Sub

' === WRAP TEXT TOGGLE ===

Public Sub ToggleWrapText(Optional control As IRibbonControl)
    Debug.Print "ToggleWrapText called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentWrap As Boolean
    currentWrap = Selection.WrapText
    
    Selection.WrapText = Not currentWrap
    Debug.Print "Wrap text changed from " & currentWrap & " to " & (Not currentWrap)
    On Error GoTo 0
End Sub

' === MERGE CELLS CYCLE ===

Public Sub CycleMergeCells(Optional control As IRibbonControl)
    Debug.Print "CycleMergeCells called"
    
    If Selection Is Nothing Then Exit Sub
    If Selection.Cells.Count < 2 Then
        MsgBox "Please select multiple cells to merge.", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    If Selection.MergeCells Then
        ' Unmerge cells
        Selection.UnMerge
        Debug.Print "Cells unmerged"
    Else
        ' Merge cells
        Selection.Merge
        Debug.Print "Cells merged"
    End If
    On Error GoTo 0
End Sub

' === ADVANCED ALIGNMENT FUNCTIONS ===

Public Sub AlignToGrid(Optional control As IRibbonControl)
    Debug.Print "AlignToGrid called"
    
    If Selection Is Nothing Then Exit Sub
    
    ' Snap selection to Excel's grid alignment
    On Error Resume Next
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    On Error GoTo 0
    
    Debug.Print "Selection aligned to grid standards"
End Sub

Public Sub CopyAlignment(Optional control As IRibbonControl)
    Debug.Print "CopyAlignment called"
    
    If Selection Is Nothing Then Exit Sub
    
    ' Store alignment properties in a custom property for later pasting
    With Selection
        On Error Resume Next
        ThisWorkbook.CustomDocumentProperties("CopiedHAlign").Delete
        ThisWorkbook.CustomDocumentProperties("CopiedVAlign").Delete
        ThisWorkbook.CustomDocumentProperties("CopiedWrapText").Delete
        ThisWorkbook.CustomDocumentProperties("CopiedOrientation").Delete
        ThisWorkbook.CustomDocumentProperties("CopiedIndent").Delete
        On Error GoTo 0
        
        On Error Resume Next
        ThisWorkbook.CustomDocumentProperties.Add "CopiedHAlign", False, msoPropertyTypeNumber, .HorizontalAlignment
        ThisWorkbook.CustomDocumentProperties.Add "CopiedVAlign", False, msoPropertyTypeNumber, .VerticalAlignment
        ThisWorkbook.CustomDocumentProperties.Add "CopiedWrapText", False, msoPropertyTypeBoolean, .WrapText
        ThisWorkbook.CustomDocumentProperties.Add "CopiedOrientation", False, msoPropertyTypeNumber, .Orientation
        ThisWorkbook.CustomDocumentProperties.Add "CopiedIndent", False, msoPropertyTypeNumber, .IndentLevel
        On Error GoTo 0
    End With
    
    MsgBox "Alignment copied. Use 'Paste Alignment' to apply to other cells.", vbInformation
End Sub

Public Sub PasteAlignment(Optional control As IRibbonControl)
    Debug.Print "PasteAlignment called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim hAlign As Long, vAlign As Long, wrapText As Boolean, orientation As Long, indent As Long
    
    hAlign = ThisWorkbook.CustomDocumentProperties("CopiedHAlign").Value
    vAlign = ThisWorkbook.CustomDocumentProperties("CopiedVAlign").Value
    wrapText = ThisWorkbook.CustomDocumentProperties("CopiedWrapText").Value
    orientation = ThisWorkbook.CustomDocumentProperties("CopiedOrientation").Value
    indent = ThisWorkbook.CustomDocumentProperties("CopiedIndent").Value
    
    If Err.Number <> 0 Then
        MsgBox "No alignment data found. Please copy alignment first.", vbExclamation
        Exit Sub
    End If
    
    With Selection
        .HorizontalAlignment = hAlign
        .VerticalAlignment = vAlign
        .WrapText = wrapText
        .Orientation = orientation
        .IndentLevel = indent
    End With
    
    On Error GoTo 0
    Debug.Print "Alignment pasted to selection"
    MsgBox "Alignment applied to selection.", vbInformation
End Sub

' === ROW HEIGHT AND COLUMN WIDTH FUNCTIONS ===

Public Sub CycleRowHeight(Optional control As IRibbonControl)
    Debug.Print "CycleRowHeight called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    ' Define standard row heights
    Dim heights As Variant
    heights = Array(15, 18, 20, 24, 30, 36, 48)
    
    Dim currentHeight As Double
    currentHeight = Selection.Rows(1).RowHeight
    
    ' Find next height in the cycle
    Dim nextHeight As Double
    nextHeight = heights(0)  ' Default
    
    Dim i As Integer
    For i = LBound(heights) To UBound(heights)
        If Abs(currentHeight - heights(i)) < 0.1 Then  ' Close enough match
            nextHeight = heights(IIf(i < UBound(heights), i + 1, 0))
            Exit For
        End If
    Next i
    
    Selection.Rows.RowHeight = nextHeight
    Debug.Print "Row height changed from " & currentHeight & " to " & nextHeight
    On Error GoTo 0
End Sub

Public Sub CycleColumnWidth(Optional control As IRibbonControl)
    Debug.Print "CycleColumnWidth called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    ' Define standard column widths
    Dim widths As Variant
    widths = Array(8.43, 10, 12, 15, 20, 25, 30)
    
    Dim currentWidth As Double
    currentWidth = Selection.Columns(1).ColumnWidth
    
    ' Find next width in the cycle
    Dim nextWidth As Double
    nextWidth = widths(0)  ' Default
    
    Dim i As Integer
    For i = LBound(widths) To UBound(widths)
        If Abs(currentWidth - widths(i)) < 0.1 Then  ' Close enough match
            nextWidth = widths(IIf(i < UBound(widths), i + 1, 0))
            Exit For
        End If
    Next i
    
    Selection.Columns.ColumnWidth = nextWidth
    Debug.Print "Column width changed from " & currentWidth & " to " & nextWidth
    On Error GoTo 0
End Sub

Public Sub AutoFitRowsAndColumns(Optional control As IRibbonControl)
    Debug.Print "AutoFitRowsAndColumns called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
    On Error GoTo 0
    
    Debug.Print "Rows and columns auto-fitted"
End Sub

' === DISTRIBUTE SPACING ===

Public Sub DistributeRowsEvenly(Optional control As IRibbonControl)
    Debug.Print "DistributeRowsEvenly called"
    
    If Selection Is Nothing Then Exit Sub
    If Selection.Rows.Count < 2 Then
        MsgBox "Please select multiple rows to distribute.", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    ' Calculate average row height and apply to all selected rows
    Dim totalHeight As Double
    Dim row As Range
    For Each row In Selection.Rows
        totalHeight = totalHeight + row.RowHeight
    Next row
    
    Dim averageHeight As Double
    averageHeight = totalHeight / Selection.Rows.Count
    
    Selection.Rows.RowHeight = averageHeight
    Debug.Print "Distributed " & Selection.Rows.Count & " rows evenly with height " & averageHeight
    On Error GoTo 0
End Sub

Public Sub DistributeColumnsEvenly(Optional control As IRibbonControl)
    Debug.Print "DistributeColumnsEvenly called"
    
    If Selection Is Nothing Then Exit Sub
    If Selection.Columns.Count < 2 Then
        MsgBox "Please select multiple columns to distribute.", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    ' Calculate average column width and apply to all selected columns
    Dim totalWidth As Double
    Dim col As Range
    For Each col In Selection.Columns
        totalWidth = totalWidth + col.ColumnWidth
    Next col
    
    Dim averageWidth As Double
    averageWidth = totalWidth / Selection.Columns.Count
    
    Selection.Columns.ColumnWidth = averageWidth
    Debug.Print "Distributed " & Selection.Columns.Count & " columns evenly with width " & averageWidth
    On Error GoTo 0
End Sub