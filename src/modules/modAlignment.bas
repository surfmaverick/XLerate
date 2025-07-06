' =============================================================================
' File: ModAlignment.bas
' Version: 2.0.0
' Date: January 2025
' Author: XLerate Development Team
'
' CHANGELOG:
' v2.0.0 - Enhanced alignment functions with Macabacus-style cycling
'        - Center cycle, horizontal cycle, vertical cycle functions
'        - Text orientation and indentation management
'        - Row height and column width cycling
'        - Cross-platform compatibility (Windows & macOS)
'        - Professional formatting standards for financial modeling
' v1.0.0 - Basic alignment functionality
' =============================================================================

Attribute VB_Name = "ModAlignment"
Option Explicit

' === CENTER CYCLE FUNCTION (Macabacus-aligned) ===

Public Sub CycleCenter(Optional control As IRibbonControl)
    ' Cycle through Left → Center → Right alignment
    ' Matches Macabacus Center Cycle - Ctrl+Alt+Shift+C
    Debug.Print "CycleCenter called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentAlignment As Long
    currentAlignment = Selection.HorizontalAlignment
    
    Select Case currentAlignment
        Case xlLeft, xlGeneral
            Selection.HorizontalAlignment = xlCenter
            Application.StatusBar = "Alignment: Center"
            Debug.Print "Changed to Center alignment"
        Case xlCenter
            Selection.HorizontalAlignment = xlRight
            Application.StatusBar = "Alignment: Right"
            Debug.Print "Changed to Right alignment"
        Case xlRight
            Selection.HorizontalAlignment = xlLeft
            Application.StatusBar = "Alignment: Left"
            Debug.Print "Changed to Left alignment"
        Case Else
            Selection.HorizontalAlignment = xlCenter
            Application.StatusBar = "Alignment: Center"
            Debug.Print "Set to Center alignment (default)"
    End Select
    
    ' Clear status after 1 second
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    On Error GoTo 0
End Sub

' === HORIZONTAL ALIGNMENT CYCLE ===

Public Sub CycleHorizontal(Optional control As IRibbonControl)
    ' Cycle through all horizontal alignment options
    ' General → Left → Center → Right → Justify → General
    Debug.Print "CycleHorizontal called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentAlignment As Long
    currentAlignment = Selection.HorizontalAlignment
    
    Select Case currentAlignment
        Case xlGeneral
            Selection.HorizontalAlignment = xlLeft
            Application.StatusBar = "Horizontal: Left"
            Debug.Print "Changed to Left alignment"
        Case xlLeft
            Selection.HorizontalAlignment = xlCenter
            Application.StatusBar = "Horizontal: Center"
            Debug.Print "Changed to Center alignment"
        Case xlCenter
            Selection.HorizontalAlignment = xlRight
            Application.StatusBar = "Horizontal: Right"
            Debug.Print "Changed to Right alignment"
        Case xlRight
            Selection.HorizontalAlignment = xlJustify
            Application.StatusBar = "Horizontal: Justify"
            Debug.Print "Changed to Justify alignment"
        Case xlJustify
            Selection.HorizontalAlignment = xlGeneral
            Application.StatusBar = "Horizontal: General"
            Debug.Print "Changed to General alignment"
        Case Else
            Selection.HorizontalAlignment = xlLeft
            Application.StatusBar = "Horizontal: Left"
            Debug.Print "Set to Left alignment (default)"
    End Select
    
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    On Error GoTo 0
End Sub

' === VERTICAL ALIGNMENT CYCLE ===

Public Sub CycleVertical(Optional control As IRibbonControl)
    ' Cycle through vertical alignment options
    ' Top → Middle → Bottom → Justify → Top
    Debug.Print "CycleVertical called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentAlignment As Long
    currentAlignment = Selection.VerticalAlignment
    
    Select Case currentAlignment
        Case xlTop
            Selection.VerticalAlignment = xlCenter
            Application.StatusBar = "Vertical: Middle"
            Debug.Print "Changed to Middle alignment"
        Case xlCenter
            Selection.VerticalAlignment = xlBottom
            Application.StatusBar = "Vertical: Bottom"
            Debug.Print "Changed to Bottom alignment"
        Case xlBottom
            Selection.VerticalAlignment = xlJustify
            Application.StatusBar = "Vertical: Justify"
            Debug.Print "Changed to Justify alignment"
        Case xlJustify
            Selection.VerticalAlignment = xlTop
            Application.StatusBar = "Vertical: Top"
            Debug.Print "Changed to Top alignment"
        Case Else
            Selection.VerticalAlignment = xlTop
            Application.StatusBar = "Vertical: Top"
            Debug.Print "Set to Top alignment (default)"
    End Select
    
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    On Error GoTo 0
End Sub

' === LEFT INDENT CYCLE ===

Public Sub CycleLeftIndent(Optional control As IRibbonControl)
    ' Cycle through indent levels 0, 1, 2, 3, then back to 0
    ' Matches Macabacus Left Indent Cycle functionality
    Debug.Print "CycleLeftIndent called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentIndent As Integer
    currentIndent = Selection.IndentLevel
    
    ' Cycle through indent levels 0, 1, 2, 3, then back to 0
    Dim nextIndent As Integer
    nextIndent = (currentIndent + 1) Mod 4
    
    Selection.IndentLevel = nextIndent
    
    Application.StatusBar = "Indent Level: " & nextIndent
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Changed indent level from " & currentIndent & " to " & nextIndent
    On Error GoTo 0
End Sub

' === TEXT ORIENTATION CYCLE ===

Public Sub CycleTextOrientation(Optional control As IRibbonControl)
    ' Cycle through text orientations: 0°, 90°, -90°, 45°, -45°
    Debug.Print "CycleTextOrientation called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentOrientation As Integer
    currentOrientation = Selection.Orientation
    
    ' Cycle through orientations
    Select Case currentOrientation
        Case 0
            Selection.Orientation = 90
            Application.StatusBar = "Text Orientation: 90°"
            Debug.Print "Changed to 90° orientation"
        Case 90
            Selection.Orientation = -90
            Application.StatusBar = "Text Orientation: -90°"
            Debug.Print "Changed to -90° orientation"
        Case -90
            Selection.Orientation = 45
            Application.StatusBar = "Text Orientation: 45°"
            Debug.Print "Changed to 45° orientation"
        Case 45
            Selection.Orientation = -45
            Application.StatusBar = "Text Orientation: -45°"
            Debug.Print "Changed to -45° orientation"
        Case -45
            Selection.Orientation = 0
            Application.StatusBar = "Text Orientation: 0°"
            Debug.Print "Changed to 0° orientation"
        Case Else
            Selection.Orientation = 0
            Application.StatusBar = "Text Orientation: 0°"
            Debug.Print "Reset to 0° orientation"
    End Select
    
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    On Error GoTo 0
End Sub

' === WRAP TEXT TOGGLE ===

Public Sub ToggleWrapText(Optional control As IRibbonControl)
    ' Toggle wrap text on/off - Ctrl+Alt+Shift+W
    Debug.Print "ToggleWrapText called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentWrap As Boolean
    currentWrap = Selection.WrapText
    
    Selection.WrapText = Not currentWrap
    
    Application.StatusBar = "Wrap Text: " & IIf(Not currentWrap, "ON", "OFF")
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Wrap text changed from " & currentWrap & " to " & (Not currentWrap)
    On Error GoTo 0
End Sub

' === MERGE CELLS OPERATIONS ===

Public Sub CycleMergeCells(Optional control As IRibbonControl)
    ' Toggle merge/unmerge cells
    Debug.Print "CycleMergeCells called"
    
    If Selection Is Nothing Then Exit Sub
    If Selection.Cells.Count < 2 Then
        MsgBox "Please select multiple cells to merge.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    On Error Resume Next
    If Selection.MergeCells Then
        ' Unmerge cells
        Selection.UnMerge
        Application.StatusBar = "Cells unmerged"
        Debug.Print "Cells unmerged"
    Else
        ' Merge cells
        Selection.Merge
        Application.StatusBar = "Cells merged"
        Debug.Print "Cells merged"
    End If
    
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    On Error GoTo 0
End Sub

' === ROW HEIGHT CYCLING ===

Public Sub CycleRowHeight(Optional control As IRibbonControl)
    ' Cycle through standard row heights used in financial modeling
    Debug.Print "CycleRowHeight called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    ' Define standard row heights for financial models
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
            If i < UBound(heights) Then
                nextHeight = heights(i + 1)
            Else
                nextHeight = heights(0)  ' Cycle back to start
            End If
            Exit For
        End If
    Next i
    
    Selection.Rows.RowHeight = nextHeight
    
    Application.StatusBar = "Row Height: " & nextHeight & " points"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Row height changed from " & currentHeight & " to " & nextHeight
    On Error GoTo 0
End Sub

' === COLUMN WIDTH CYCLING ===

Public Sub CycleColumnWidth(Optional control As IRibbonControl)
    ' Cycle through standard column widths used in financial modeling
    Debug.Print "CycleColumnWidth called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    ' Define standard column widths for financial models
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
            If i < UBound(widths) Then
                nextWidth = widths(i + 1)
            Else
                nextWidth = widths(0)  ' Cycle back to start
            End If
            Exit For
        End If
    Next i
    
    Selection.Columns.ColumnWidth = nextWidth
    
    Application.StatusBar = "Column Width: " & nextWidth & " characters"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Column width changed from " & currentWidth & " to " & nextWidth
    On Error GoTo 0
End Sub

' === AUTO-FIT FUNCTIONS ===

Public Sub AutoFitRowsAndColumns(Optional control As IRibbonControl)
    ' Auto-fit both rows and columns for selected range
    Debug.Print "AutoFitRowsAndColumns called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
    
    Application.StatusBar = "Rows and columns auto-fitted"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Rows and columns auto-fitted"
    On Error GoTo 0
End Sub

Public Sub AutoFitRows(Optional control As IRibbonControl)
    ' Auto-fit row heights only
    Debug.Print "AutoFitRows called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Selection.Rows.AutoFit
    
    Application.StatusBar = "Rows auto-fitted"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Rows auto-fitted"
    On Error GoTo 0
End Sub

Public Sub AutoFitColumns(Optional control As IRibbonControl)
    ' Auto-fit column widths only
    Debug.Print "AutoFitColumns called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Selection.Columns.AutoFit
    
    Application.StatusBar = "Columns auto-fitted"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Columns auto-fitted"
    On Error GoTo 0
End Sub

' === DISTRIBUTE SPACING ===

Public Sub DistributeRowsEvenly(Optional control As IRibbonControl)
    ' Distribute selected rows with equal height
    Debug.Print "DistributeRowsEvenly called"
    
    If Selection Is Nothing Then Exit Sub
    If Selection.Rows.Count < 2 Then
        MsgBox "Please select multiple rows to distribute.", vbInformation, "XLerate"
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
    
    Application.StatusBar = "Distributed " & Selection.Rows.Count & " rows evenly"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Distributed " & Selection.Rows.Count & " rows evenly with height " & averageHeight
    On Error GoTo 0
End Sub

Public Sub DistributeColumnsEvenly(Optional control As IRibbonControl)
    ' Distribute selected columns with equal width
    Debug.Print "DistributeColumnsEvenly called"
    
    If Selection Is Nothing Then Exit Sub
    If Selection.Columns.Count < 2 Then
        MsgBox "Please select multiple columns to distribute.", vbInformation, "XLerate"
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
    
    Application.StatusBar = "Distributed " & Selection.Columns.Count & " columns evenly"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Distributed " & Selection.Columns.Count & " columns evenly with width " & averageWidth
    On Error GoTo 0
End Sub

' === ADVANCED ALIGNMENT FUNCTIONS ===

Public Sub AlignToGrid(Optional control As IRibbonControl)
    ' Apply standard grid alignment for financial models
    Debug.Print "AlignToGrid called"
    
    If Selection Is Nothing Then Exit Sub
    
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
    
    Application.StatusBar = "Grid alignment applied"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Selection aligned to grid standards"
    On Error GoTo 0
End Sub

Public Sub CopyAlignment(Optional control As IRibbonControl)
    ' Copy alignment properties for later pasting
    Debug.Print "CopyAlignment called"
    
    If Selection Is Nothing Then Exit Sub
    
    ' Store alignment properties in workbook custom properties
    With Selection
        On Error Resume Next
        ' Delete existing properties
        ThisWorkbook.CustomDocumentProperties("CopiedHAlign").Delete
        ThisWorkbook.CustomDocumentProperties("CopiedVAlign").Delete
        ThisWorkbook.CustomDocumentProperties("CopiedWrapText").Delete
        ThisWorkbook.CustomDocumentProperties("CopiedOrientation").Delete
        ThisWorkbook.CustomDocumentProperties("CopiedIndent").Delete
        On Error GoTo 0
        
        On Error Resume Next
        ' Add new properties
        ThisWorkbook.CustomDocumentProperties.Add "CopiedHAlign", False, msoPropertyTypeNumber, .HorizontalAlignment
        ThisWorkbook.CustomDocumentProperties.Add "CopiedVAlign", False, msoPropertyTypeNumber, .VerticalAlignment
        ThisWorkbook.CustomDocumentProperties.Add "CopiedWrapText", False, msoPropertyTypeBoolean, .WrapText
        ThisWorkbook.CustomDocumentProperties.Add "CopiedOrientation", False, msoPropertyTypeNumber, .Orientation
        ThisWorkbook.CustomDocumentProperties.Add "CopiedIndent", False, msoPropertyTypeNumber, .IndentLevel
        On Error GoTo 0
    End With
    
    Application.StatusBar = "Alignment copied - use Paste Alignment"
    Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
    
    Debug.Print "Alignment properties copied"
End Sub

Public Sub PasteAlignment(Optional control As IRibbonControl)
    ' Paste previously copied alignment properties
    Debug.Print "PasteAlignment called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim hAlign As Long, vAlign As Long, wrapText As Boolean, orientation As Long, indent As Long
    
    ' Retrieve stored properties
    hAlign = ThisWorkbook.CustomDocumentProperties("CopiedHAlign").Value
    vAlign = ThisWorkbook.CustomDocumentProperties("CopiedVAlign").Value
    wrapText = ThisWorkbook.CustomDocumentProperties("CopiedWrapText").Value
    orientation = ThisWorkbook.CustomDocumentProperties("CopiedOrientation").Value
    indent = ThisWorkbook.CustomDocumentProperties("CopiedIndent").Value
    
    If Err.Number <> 0 Then
        MsgBox "No alignment data found. Please copy alignment first.", vbExclamation, "XLerate"
        Exit Sub
    End If
    
    ' Apply alignment properties
    With Selection
        .HorizontalAlignment = hAlign
        .VerticalAlignment = vAlign
        .WrapText = wrapText
        .Orientation = orientation
        .IndentLevel = indent
    End With
    
    Application.StatusBar = "Alignment pasted successfully"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Alignment pasted to selection"
    On Error GoTo 0
End Sub