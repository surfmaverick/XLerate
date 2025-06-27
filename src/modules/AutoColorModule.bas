' =============================================================================
' File: AutoColorModule.bas
' Version: 2.0.0
' Description: Enhanced auto-color functions with sheet and workbook support
' Author: XLerate Development Team
' Created: Enhanced for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

Attribute VB_Name = "AutoColorModule"
Option Explicit

Private Const NAME_PREFIX As String = "AutoColor_"
Private Const MAX_CELLS As Long = 50000  ' Maximum number of cells to process in one go

' Function to get color from saved settings or return default
Private Function GetSavedColor(colorName As String, defaultColor As Long) As Long
    On Error Resume Next
    Dim colorValue As String
    colorValue = ThisWorkbook.Names(NAME_PREFIX & colorName).RefersTo
    If Err.Number = 0 And colorValue <> "" Then
        GetSavedColor = CLng(Mid(colorValue, 2)) ' Remove the = sign
    Else
        GetSavedColor = defaultColor
    End If
    On Error GoTo 0
End Function

' === MAIN AUTO-COLOR FUNCTIONS (Macabacus-style) ===

Sub AutoColorCells(Optional control As IRibbonControl)
    Debug.Print "AutoColorCells started"
    
    Application.ScreenUpdating = False
    
    ' Get saved colors or use defaults
    Dim colorInput As Long: colorInput = GetSavedColor("Input", 16711680)         ' Default: Blue
    Dim colorFormula As Long: colorFormula = GetSavedColor("Formula", 0)          ' Default: Black
    Dim colorWorksheetLink As Long: colorWorksheetLink = GetSavedColor("WorksheetLink", 32768)   ' Default: Green
    Dim colorWorkbookLink As Long: colorWorkbookLink = GetSavedColor("WorkbookLink", 16751052)  ' Default: Light Purple
    Dim colorExternal As Long: colorExternal = GetSavedColor("External", 15773696)      ' Default: Light Blue
    Dim colorHyperlink As Long: colorHyperlink = GetSavedColor("Hyperlink", 33023)      ' Default: Orange
    Dim colorPartialInput As Long: colorPartialInput = GetSavedColor("PartialInput", 128)     ' Default: Purple
    
    Dim rng As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection
    
    ProcessRangeAutoColor rng, colorInput, colorFormula, colorWorksheetLink, colorWorkbookLink, colorExternal, colorHyperlink, colorPartialInput
    
    Application.ScreenUpdating = True
    Debug.Print "AutoColorCells ended"
End Sub

Sub AutoColorSheet(Optional control As IRibbonControl)
    Debug.Print "AutoColorSheet started"
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Auto-coloring sheet... Please wait."
    
    ' Get saved colors
    Dim colorInput As Long: colorInput = GetSavedColor("Input", 16711680)
    Dim colorFormula As Long: colorFormula = GetSavedColor("Formula", 0)
    Dim colorWorksheetLink As Long: colorWorksheetLink = GetSavedColor("WorksheetLink", 32768)
    Dim colorWorkbookLink As Long: colorWorkbookLink = GetSavedColor("WorkbookLink", 16751052)
    Dim colorExternal As Long: colorExternal = GetSavedColor("External", 15773696)
    Dim colorHyperlink As Long: colorHyperlink = GetSavedColor("Hyperlink", 33023)
    Dim colorPartialInput As Long: colorPartialInput = GetSavedColor("PartialInput", 128)
    
    ' Process the entire used range of the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If Not ws.UsedRange Is Nothing Then
        ProcessRangeAutoColor ws.UsedRange, colorInput, colorFormula, colorWorksheetLink, colorWorkbookLink, colorExternal, colorHyperlink, colorPartialInput
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Debug.Print "AutoColorSheet ended"
    
    MsgBox "Auto-coloring completed for worksheet: " & ws.Name, vbInformation
End Sub

Sub AutoColorWorkbook(Optional control As IRibbonControl)
    Debug.Print "AutoColorWorkbook started"
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Auto-coloring workbook... Please wait."
    
    ' Get saved colors
    Dim colorInput As Long: colorInput = GetSavedColor("Input", 16711680)
    Dim colorFormula As Long: colorFormula = GetSavedColor("Formula", 0)
    Dim colorWorksheetLink As Long: colorWorksheetLink = GetSavedColor("WorksheetLink", 32768)
    Dim colorWorkbookLink As Long: colorWorkbookLink = GetSavedColor("WorkbookLink", 16751052)
    Dim colorExternal As Long: colorExternal = GetSavedColor("External", 15773696)
    Dim colorHyperlink As Long: colorHyperlink = GetSavedColor("Hyperlink", 33023)
    Dim colorPartialInput As Long: colorPartialInput = GetSavedColor("PartialInput", 128)
    
    Dim ws As Worksheet
    Dim totalSheets As Integer
    Dim currentSheet As Integer
    
    totalSheets = ActiveWorkbook.Worksheets.Count
    currentSheet = 0
    
    ' Process all worksheets in the workbook
    For Each ws In ActiveWorkbook.Worksheets
        currentSheet = currentSheet + 1
        Application.StatusBar = "Auto-coloring workbook... Sheet " & currentSheet & " of " & totalSheets & " (" & ws.Name & ")"
        
        If Not ws.UsedRange Is Nothing Then
            ProcessRangeAutoColor ws.UsedRange, colorInput, colorFormula, colorWorksheetLink, colorWorkbookLink, colorExternal, colorHyperlink, colorPartialInput
        End If
    Next ws
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Debug.Print "AutoColorWorkbook ended"
    
    MsgBox "Auto-coloring completed for entire workbook (" & totalSheets & " sheets processed).", vbInformation
End Sub

' === CORE PROCESSING FUNCTION ===

Private Sub ProcessRangeAutoColor(rng As Range, colorInput As Long, colorFormula As Long, _
                                 colorWorksheetLink As Long, colorWorkbookLink As Long, _
                                 colorExternal As Long, colorHyperlink As Long, colorPartialInput As Long)
    
    On Error Resume Next
    
    ' Get only cells with content (formulas or values), excluding blanks
    Dim usedCells As Range
    Set usedCells = rng.SpecialCells(xlCellTypeConstants)
    
    Dim formulaCells As Range
    Set formulaCells = rng.SpecialCells(xlCellTypeFormulas)
    
    ' Combine the ranges if both exist
    If Not usedCells Is Nothing Then
        If Not formulaCells Is Nothing Then
            Set usedCells = Union(usedCells, formulaCells)
        End If
    Else
        If Not formulaCells Is Nothing Then
            Set usedCells = formulaCells
        End If
    End If
    
    On Error GoTo 0
    
    ' If no cells with content found, exit
    If usedCells Is Nothing Then Exit Sub
    
    ' Process cells in batches to avoid performance issues
    Dim cellCount As Long
    cellCount = 0
    
    ' Process only the cells that contain something
    Dim cell As Range
    For Each cell In usedCells
        cellCount = cellCount + 1
        
        ' Update status bar every 1000 cells
        If cellCount Mod 1000 = 0 Then
            Application.StatusBar = "Processing cell " & cellCount & "..."
        End If
        
        If HasFormula(cell) Then
            If IsPartialInput(cell) Then
                cell.Font.Color = colorPartialInput
            ElseIf IsWorkbookLink(cell) Then
                cell.Font.Color = colorWorkbookLink
            ElseIf IsWorksheetLink(cell) Then
                cell.Font.Color = colorWorksheetLink
            ElseIf IsExternalReference(cell) Then
                cell.Font.Color = colorExternal
            ElseIf IsInput(cell) Then
                cell.Font.Color = colorInput
            Else
                cell.Font.Color = colorFormula
            End If
        ElseIf IsHyperlink(cell) Then
            cell.Font.Color = colorHyperlink
        ElseIf IsInput(cell) Then
            cell.Font.Color = colorInput
        End If
        
        ' Break if processing too many cells at once
        If cellCount > MAX_CELLS Then
            Debug.Print "Reached maximum cell limit for single operation: " & MAX_CELLS
            Exit For
        End If
    Next cell
    
    Debug.Print "Processed " & cellCount & " cells"
End Sub

' === ENHANCED COLOR CYCLING FUNCTIONS ===

Public Sub CycleFontColor(Optional control As IRibbonControl)
    If Selection Is Nothing Then Exit Sub
    
    ' Define color cycle (Macabacus-style)
    Dim colors As Variant
    colors = Array(RGB(0, 0, 0), RGB(255, 0, 0), RGB(0, 0, 255), RGB(0, 128, 0), RGB(128, 0, 128), RGB(255, 165, 0))
    
    CycleColorProperty colors, "Font"
End Sub

Public Sub CycleFillColor(Optional control As IRibbonControl)
    If Selection Is Nothing Then Exit Sub
    
    ' Define fill color cycle
    Dim colors As Variant
    colors = Array(RGB(255, 255, 255), RGB(255, 255, 0), RGB(192, 192, 192), RGB(255, 192, 203), RGB(173, 216, 230), RGB(144, 238, 144))
    
    CycleColorProperty colors, "Fill"
End Sub

Public Sub CycleBorderColor(Optional control As IRibbonControl)
    If Selection Is Nothing Then Exit Sub
    
    ' Define border color cycle
    Dim colors As Variant
    colors = Array(RGB(0, 0, 0), RGB(128, 128, 128), RGB(255, 0, 0), RGB(0, 0, 255), RGB(0, 128, 0))
    
    CycleColorProperty colors, "Border"
End Sub

Public Sub CycleBlueBlack(Optional control As IRibbonControl)
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim currentColor As Long
    currentColor = Selection.Font.Color
    
    If currentColor = RGB(0, 0, 255) Then  ' Blue
        Selection.Font.Color = RGB(0, 0, 0)        ' Black
    Else
        Selection.Font.Color = RGB(0, 0, 255)      ' Blue
    End If
    On Error GoTo 0
End Sub

Private Sub CycleColorProperty(colors As Variant, propertyType As String)
    On Error Resume Next
    
    Dim currentColor As Long
    Dim nextColorIndex As Integer
    Dim i As Integer
    
    ' Get current color based on property type
    Select Case propertyType
        Case "Font"
            currentColor = Selection.Font.Color
        Case "Fill"
            currentColor = Selection.Interior.Color
        Case "Border"
            currentColor = Selection.Borders(xlEdgeTop).Color
    End Select
    
    ' Find current color in array and get next one
    nextColorIndex = 0  ' Default to first color
    For i = LBound(colors) To UBound(colors)
        If colors(i) = currentColor Then
            nextColorIndex = IIf(i < UBound(colors), i + 1, LBound(colors))
            Exit For
        End If
    Next i
    
    ' Apply the next color
    Select Case propertyType
        Case "Font"
            Selection.Font.Color = colors(nextColorIndex)
        Case "Fill"
            Selection.Interior.Color = colors(nextColorIndex)
        Case "Border"
            Dim edges As Variant
            edges = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            Dim edge As Variant
            For Each edge In edges
                Selection.Borders(edge).Color = colors(nextColorIndex)
            Next edge
    End Select
    
    On Error GoTo 0
End Sub

' === EXISTING HELPER FUNCTIONS (Unchanged) ===

Private Function HasFormula(cell As Range) As Boolean
    HasFormula = cell.HasFormula
End Function

Private Function IsWorksheetLink(cell As Range) As Boolean
    If Not cell.HasFormula Then Exit Function
    
    Dim formula As String
    formula = cell.Formula
    
    ' Check if formula references another sheet but not another workbook
    IsWorksheetLink = (InStr(1, formula, "!") > 0) And (Left(formula, 1) <> "[")
End Function

Private Function IsWorkbookLink(cell As Range) As Boolean
    If Not cell.HasFormula Then Exit Function
    
    Dim formula As String
    formula = cell.Formula
    
    ' Look for [workbook] pattern anywhere in the formula
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\[[^\]]+\]"  ' Matches anything in square brackets
    
    IsWorkbookLink = regEx.Test(formula)
End Function

Private Function IsHyperlink(cell As Range) As Boolean
    IsHyperlink = cell.Hyperlinks.Count > 0
End Function

Private Function IsExternalReference(cell As Range) As Boolean
    If Not cell.HasFormula Then Exit Function
    
    ' Check for common external data functions
    Dim formula As String
    formula = UCase(cell.Formula)
    
    IsExternalReference = (InStr(1, formula, "WEBSERVICE") > 0) Or _
                         (InStr(1, formula, "ODBC") > 0) Or _
                         (InStr(1, formula, "SQL") > 0)
End Function

Private Function IsInput(cell As Range) As Boolean
    ' Consider a cell as input if:
    ' 1. It has a value but no formula, or
    ' 2. It's a formula that only contains numbers and operators, or
    ' 3. It's a formula that doesn't reference any cells
    If IsEmpty(cell.Value) Then Exit Function
    
    ' Check if cell contains text
    If VarType(cell.Value) = vbString Then Exit Function
    
    ' Check if cell contains a date
    If IsDate(cell.Value) Then Exit Function
    
    ' If it's a formula, check if it's only numbers/operators or has no references
    If cell.HasFormula Then
        Dim formula As String
        formula = cell.Formula
        
        ' Check for cell references
        Dim regEx As Object
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        
        ' Pattern to match any cell reference (A1 style or R1C1 style)
        regEx.Pattern = "[$]?[A-Za-z]+[$]?[0-9]+|R[0-9]*C[0-9]*"
        
        ' If no cell references found, treat as input
        If Not regEx.Test(formula) Then
            IsInput = True
            Exit Function
        End If
        
        ' Also check if it's only numbers and operators
        IsInput = IsOnlyNumbersAndOperators(formula)
        Exit Function
    End If
    
    ' If we get here, it's a numeric input
    IsInput = True
End Function

Private Function IsOnlyNumbersAndOperators(formula As String) As Boolean
    ' Remove the equals sign if present
    If Left(formula, 1) = "=" Then formula = Mid(formula, 2)
    
    ' Create regex to match only numbers, decimals, and basic operators
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.Pattern = "^[-+*/\d\s\.,()]*$"
    
    IsOnlyNumbersAndOperators = regEx.Test(formula)
End Function

Private Function IsPartialInput(cell As Range) As Boolean
    ' Don't process cells with text or dates
    If Not cell.HasFormula Then Exit Function
    If VarType(cell.Value) = vbString Then Exit Function
    If IsDate(cell.Value) Then Exit Function
    
    Dim formula As String
    formula = cell.Formula
    
    ' If formula only contains numbers and basic operators, treat as input
    If IsOnlyNumbersAndOperators(formula) Then
        Exit Function
    End If
    
    ' Look for numbers in the formula (excluding cell references and function names)
    If Left(formula, 1) = "=" Then formula = Mid(formula, 2)
    
    ' Handle sheet references by replacing them with a placeholder
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    
    ' Replace sheet references (including workbook references) with placeholder
    regEx.Pattern = "(\[[^\]]+\])?'?[^!]+!'?"
    formula = regEx.Replace(formula, "SHEET_REF!")
    
    ' Exclude common functions that might contain numbers
    Dim commonFuncs As Variant
    commonFuncs = Array("SUM", "AVERAGE", "COUNT", "LEFT", "RIGHT", "MID", "ROUND")
    Dim func As Variant
    For Each func In commonFuncs
        formula = Replace(formula, func, "")
    Next func
    
    ' Remove Excel-specific symbols and cell references
    ' Remove Excel-specific symbols ($, %, etc.)
    regEx.Pattern = "[$%]"
    formula = regEx.Replace(formula, "")
    
    ' Remove A1 style references (including with $ signs)
    regEx.Pattern = "[$]?[A-Za-z]+[$]?[0-9]+"
    formula = regEx.Replace(formula, "")
    
    ' Remove R1C1 style references
    regEx.Pattern = "R[0-9]*C[0-9]*"
    formula = regEx.Replace(formula, "")
    
    ' Now look for remaining numbers
    regEx.Pattern = "[0-9]+"
    IsPartialInput = regEx.Test(formula)
End Function