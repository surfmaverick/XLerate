Attribute VB_Name = "ModErrorWrap"
' ModErrorWrap
Option Explicit

Sub WrapWithError(control As IRibbonControl)
    Dim selectedCell As Range
    Dim cellCount As Long
    Dim errorCount As Long
    
    cellCount = 0
    errorCount = 0

    ' Iterate through each cell in the selection
    For Each selectedCell In Selection
        ' Check if the cell contains a formula
        If selectedCell.HasFormula Then
            ' Wrap the formula with IFERROR
            selectedCell.formula = "=IFERROR(" & Mid(selectedCell.formula, 2) & ", NA())"
            cellCount = cellCount + 1
        Else
            errorCount = errorCount + 1
        End If
    Next selectedCell
    
    ' Provide feedback based on the results
    If cellCount = 0 Then
        MsgBox "None of the selected cells contained a formula that could be wrapped.", vbExclamation
    ElseIf errorCount > 0 Then
        MsgBox "Some of the selected cells did not contain formulas and were skipped.", vbExclamation
    End If
End Sub
