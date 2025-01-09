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

Sub AutoColorCells(control As IRibbonControl)
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
    If usedCells Is Nothing Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Process only the cells that contain something
    Dim cell As Range
    For Each cell In usedCells
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
    Next cell
    
    Application.ScreenUpdating = True
    Debug.Print "AutoColorCells ended"
End Sub

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