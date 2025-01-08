Option Explicit

' Color constants - using direct color values instead of RGB function
Private Const COLOR_INPUT As Long = 16711680         ' Blue (RGB 0, 0, 255)
Private Const COLOR_FORMULA As Long = 0              ' Black (RGB 0, 0, 0)
Private Const COLOR_WORKSHEET_LINK As Long = 32768   ' Green (RGB 0, 128, 0)
Private Const COLOR_WORKBOOK_LINK As Long = 128      ' Purple (RGB 128, 0, 128)
Private Const COLOR_EXTERNAL As Long = 15773696      ' Light Blue (RGB 0, 176, 240)
Private Const COLOR_HYPERLINK As Long = 33023        ' Orange (RGB 255, 128, 0)
Private Const COLOR_PARTIAL_INPUT As Long = 16751052 ' Light Purple (RGB 204, 153, 255)

Sub AutoColorCells(control As IRibbonControl)
    Debug.Print "AutoColorCells started"
    
    ' Use selected range only
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Dim rng As Range
    Set rng = Selection
    Debug.Print "Using selected range: " & rng.Address
    
    ApplyAutoColor rng
    Debug.Print "AutoColorCells ended"
End Sub

Private Sub ApplyAutoColor(rng As Range)
    Debug.Print "ApplyAutoColor started for range: " & rng.Address
    
    ' Apply colors
    Dim cell As Range
    For Each cell In rng
        Debug.Print "Processing cell: " & cell.Address & " Formula: " & cell.Formula
        
        ' Apply new color based on cell content
        If HasFormula(cell) Then
            If IsWorkbookLink(cell) Then
                cell.Font.Color = COLOR_WORKBOOK_LINK
                Debug.Print "Workbook Link detected"
            ElseIf IsWorksheetLink(cell) Then
                cell.Font.Color = COLOR_WORKSHEET_LINK
                Debug.Print "Worksheet Link detected"
            ElseIf IsPartialInput(cell) Then
                cell.Font.Color = COLOR_PARTIAL_INPUT
                Debug.Print "Partial Input detected"
            ElseIf IsExternalReference(cell) Then
                cell.Font.Color = COLOR_EXTERNAL
                Debug.Print "External Reference detected"
            Else
                cell.Font.Color = COLOR_FORMULA
                Debug.Print "Regular Formula detected"
            End If
        ElseIf IsHyperlink(cell) Then
            cell.Font.Color = COLOR_HYPERLINK
            Debug.Print "Hyperlink detected"
        ElseIf IsInput(cell) Then
            cell.Font.Color = COLOR_INPUT
            Debug.Print "Input detected"
        End If
    Next cell
    
    Debug.Print "ApplyAutoColor ended"
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
    
    ' Check if formula references another workbook
    IsWorkbookLink = (Left(cell.Formula, 1) = "[")
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
    ' Consider a cell as input if it has a value but no formula
    ' Exclude text and dates
    If cell.HasFormula Then Exit Function
    If IsEmpty(cell.Value) Then Exit Function
    
    ' Check if cell contains text
    If VarType(cell.Value) = vbString Then Exit Function
    
    ' Check if cell contains a date
    If IsDate(cell.Value) Then Exit Function
    
    ' If we get here, it's a numeric input
    IsInput = True
End Function

Private Function IsPartialInput(cell As Range) As Boolean
    ' Don't process cells with text or dates
    If Not cell.HasFormula Then Exit Function
    If VarType(cell.Value) = vbString Then Exit Function
    If IsDate(cell.Value) Then Exit Function
    
    Dim formula As String
    formula = cell.Formula
    
    ' Look for numbers in the formula (excluding cell references and function names)
    If Left(formula, 1) = "=" Then formula = Mid(formula, 2)
    
    ' Exclude common functions that might contain numbers (like LEFT, RIGHT, MID)
    Dim commonFuncs As Variant
    commonFuncs = Array("SUM", "AVERAGE", "COUNT", "LEFT", "RIGHT", "MID", "ROUND")
    Dim func As Variant
    For Each func In commonFuncs
        formula = Replace(formula, func, "")
    Next func
    
    ' Remove cell references (both A1 and R1C1 style)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    
    ' Remove A1 style references
    regEx.Pattern = "[A-Za-z]+[0-9]+"
    formula = regEx.Replace(formula, "")
    
    ' Remove R1C1 style references
    regEx.Pattern = "R[0-9]*C[0-9]*"
    formula = regEx.Replace(formula, "")
    
    ' Now look for remaining numbers
    regEx.Pattern = "[0-9]+"
    IsPartialInput = regEx.Test(formula)
End Function 