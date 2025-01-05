Attribute VB_Name = "ModSwitchSign"
' ModSwitchSign
Option Explicit

Sub SwitchCellSign(control As IRibbonControl)
    Dim cell As Range
    Dim formulaStr As String
    Dim arrayFormula As Boolean
    
    ' Check if any cells are selected
    If Selection Is Nothing Then
        MsgBox "Please select one or more cells.", vbExclamation
        Exit Sub
    End If
    
    ' Use error handling for the undo record
    On Error Resume Next
    Application.EnableEvents = False
    
    ' Start undo group if available
    If Not Application.UndoRecord Is Nothing Then
        Application.UndoRecord.StartCustomRecord "Switch Sign"
    End If
    
    ' Performance optimization for large selections
    If Selection.Cells.Count > 1000 Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.StatusBar = "Switching signs... Please wait."
    End If
    
    For Each cell In Selection
        If cell.HasFormula Then
            formulaStr = cell.Formula
            arrayFormula = cell.HasArray
            ' Handle array formulas separately
            If arrayFormula Then
                ' 1. Remove outer {}
                formulaStr = Mid(formulaStr, 2, Len(formulaStr) - 2)
            End If
            ' 2. Apply the sign switch using -()
            If Left(formulaStr, 1) = "=" Then
                ' Formula starts with =
                formulaStr = "=-(" & Mid(formulaStr, 2) & ")"
            Else
                ' Formula might be a named range or start directly with a value or function
                formulaStr = "-(" & formulaStr & ")"
            End If
            ' 3. Restore array formula if needed
            If arrayFormula Then
                cell.FormulaArray = "=" & formulaStr
            Else
                cell.Formula = formulaStr
            End If
        Else
            ' Handle cells with values (numbers)
            If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                cell.value = -cell.value
            End If
        End If
    Next cell
    
    ' Restore settings
    If Selection.Cells.Count > 1000 Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
    End If
    
    ' End undo group if available
    If Not Application.UndoRecord Is Nothing Then
        Application.UndoRecord.EndCustomRecord
    End If
    
    Application.EnableEvents = True
    On Error GoTo 0
End Sub
