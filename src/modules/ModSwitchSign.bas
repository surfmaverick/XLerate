' =============================================================================
' File: ModSwitchSign.bas
' Version: 2.0.0
' Description: Sign switching functions for Macabacus-style value manipulation
' Author: XLerate Development Team
' Created: Enhanced for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

Attribute VB_Name = "ModSwitchSign"
Option Explicit

Public Sub SwitchCellSign(Optional control As IRibbonControl)
    Debug.Print "SwitchCellSign called"
    
    ' Check if any cells are selected
    If Selection Is Nothing Then
        MsgBox "Please select one or more cells.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim processedCount As Long
    processedCount = 0
    
    ' Performance optimization for large selections
    If Selection.Cells.Count > 1000 Then
        Application.Calculation = xlCalculationManual
        Application.StatusBar = "Switching signs... Please wait."
    End If
    
    On Error Resume Next
    
    Dim cell As Range
    For Each cell In Selection
        If ProcessCellSignSwitch(cell) Then
            processedCount = processedCount + 1
        End If
        
        ' Update status for large operations
        If processedCount Mod 500 = 0 And Selection.Cells.Count > 1000 Then
            Application.StatusBar = "Processed " & processedCount & " cells..."
        End If
    Next cell
    
    ' Restore settings
    If Selection.Cells.Count > 1000 Then
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    On Error GoTo 0
    
    If processedCount > 0 Then
        Debug.Print "SwitchCellSign completed - " & processedCount & " cells processed"
    Else
        MsgBox "No suitable cells found to switch signs.", vbInformation
    End If
End Sub

Private Function ProcessCellSignSwitch(cell As Range) As Boolean
    ProcessCellSignSwitch = False
    
    On Error Resume Next
    
    If cell.HasFormula Then
        ' Handle formula cells
        If SwitchFormulaSign(cell) Then
            ProcessCellSignSwitch = True
        End If
    ElseIf IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
        ' Handle numeric values
        If cell.Value <> 0 Then  ' Don't switch zero
            cell.Value = -cell.Value
            ProcessCellSignSwitch = True
        End If
    End If
    
    On Error GoTo 0
End Function

Private Function SwitchFormulaSign(cell As Range) As Boolean
    SwitchFormulaSign = False
    
    On Error Resume Next
    
    Dim originalFormula As String
    originalFormula = cell.Formula
    
    Dim newFormula As String
    
    ' Check if formula is already negated with a simple minus sign
    If Left(Trim(Mid(originalFormula, 2)), 1) = "-" And Mid(originalFormula, 2, 1) <> "(" Then
        ' Remove the minus sign
        newFormula = "=" & Mid(Trim(Mid(originalFormula, 2)), 2)
    ElseIf Left(originalFormula, 3) = "=-(" And Right(originalFormula, 1) = ")" Then
        ' Remove the -( ) wrapper
        newFormula = "=" & Mid(originalFormula, 4, Len(originalFormula) - 4)
    Else
        ' Add minus sign wrapper
        newFormula = "=-(" & Mid(originalFormula, 2) & ")"
    End If
    
    ' Apply the new formula
    If cell.HasArray Then
        ' Handle array formulas specially
        cell.FormulaArray = newFormula
    Else
        cell.Formula = newFormula
    End If
    
    SwitchFormulaSign = True
    On Error GoTo 0
End Function

' Advanced sign switching with options
Public Sub SwitchSignAdvanced(Optional control As IRibbonControl)
    Debug.Print "SwitchSignAdvanced called"
    
    If Selection Is Nothing Then
        MsgBox "Please select one or more cells.", vbExclamation
        Exit Sub
    End If
    
    ' Show options dialog
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox("Choose sign switching method:" & vbNewLine & vbNewLine & _
                       "YES = Multiply by -1 (preserves formatting)" & vbNewLine & _
                       "NO = Toggle formula sign (for formulas)" & vbNewLine & _
                       "CANCEL = Cancel operation", _
                       vbYesNoCancel + vbQuestion, "Advanced Sign Switch")
    
    Select Case userChoice
        Case vbYes
            SwitchSignMultiply
        Case vbNo
            SwitchCellSign control
        Case vbCancel
            Exit Sub
    End Select
End Sub

Private Sub SwitchSignMultiply()
    Debug.Print "SwitchSignMultiply called"
    
    Application.ScreenUpdating = False
    
    Dim processedCount As Long
    processedCount = 0
    
    ' Use Paste Special multiply to switch signs
    ' First, put -1 in clipboard
    Dim tempCell As Range
    Set tempCell = ActiveSheet.Cells(1048576, 16384)  ' Use last cell as temp
    tempCell.Value = -1
    tempCell.Copy
    
    On Error Resume Next
    
    ' Apply multiply operation to selection
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlPasteSpecialOperationMultiply
    
    ' Clean up
    Application.CutCopyMode = False
    tempCell.Clear
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    Debug.Print "SwitchSignMultiply completed"
End Sub

' Batch operations
Public Sub SwitchSignInRange(Optional control As IRibbonControl)
    Debug.Print "SwitchSignInRange called"
    
    Dim targetRange As Range
    On Error Resume Next
    Set targetRange = Application.InputBox("Select range to switch signs:", "Range Selection", Selection.Address, Type:=8)
    On Error GoTo 0
    
    If targetRange Is Nothing Then Exit Sub
    
    ' Temporarily select the range and process
    Dim originalSelection As Range
    Set originalSelection = Selection
    
    targetRange.Select
    SwitchCellSign control
    
    ' Restore original selection
    originalSelection.Select
End Sub

Public Sub SwitchSignFormulasOnly(Optional control As IRibbonControl)
    Debug.Print "SwitchSignFormulasOnly called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim processedCount As Long
    processedCount = 0
    
    On Error Resume Next
    
    ' Get only formula cells
    Dim formulaCells As Range
    Set formulaCells = Selection.SpecialCells(xlCellTypeFormulas)
    
    If Not formulaCells Is Nothing Then
        Dim cell As Range
        For Each cell In formulaCells
            If SwitchFormulaSign(cell) Then
                processedCount = processedCount + 1
            End If
        Next cell
    End If
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    If processedCount > 0 Then
        MsgBox processedCount & " formula(s) had their signs switched.", vbInformation
    Else
        MsgBox "No formulas found in selection.", vbInformation
    End If
    
    Debug.Print "SwitchSignFormulasOnly completed - " & processedCount & " formulas processed"
End Sub

Public Sub SwitchSignValuesOnly(Optional control As IRibbonControl)
    Debug.Print "SwitchSignValuesOnly called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim processedCount As Long
    processedCount = 0
    
    On Error Resume Next
    
    ' Get only constant cells (values, not formulas)
    Dim constantCells As Range
    Set constantCells = Selection.SpecialCells(xlCellTypeConstants, xlNumbers)
    
    If Not constantCells Is Nothing Then
        Dim cell As Range
        For Each cell In constantCells
            If IsNumeric(cell.Value) And cell.Value <> 0 Then
                cell.Value = -cell.Value
                processedCount = processedCount + 1
            End If
        Next cell
    End If
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    If processedCount > 0 Then
        MsgBox processedCount & " value(s) had their signs switched.", vbInformation
    Else
        MsgBox "No numeric values found in selection.", vbInformation
    End If
    
    Debug.Print "SwitchSignValuesOnly completed - " & processedCount & " values processed"
End Sub

' Conditional sign switching
Public Sub SwitchSignIfPositive(Optional control As IRibbonControl)
    Debug.Print "SwitchSignIfPositive called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim processedCount As Long
    processedCount = 0
    
    Dim cell As Range
    For Each cell In Selection
        On Error Resume Next
        
        If cell.HasFormula Then
            ' For formulas, check the result
            If IsNumeric(cell.Value) And cell.Value > 0 Then
                If SwitchFormulaSign(cell) Then
                    processedCount = processedCount + 1
                End If
            End If
        ElseIf IsNumeric(cell.Value) And cell.Value > 0 Then
            cell.Value = -cell.Value
            processedCount = processedCount + 1
        End If
        
        On Error GoTo 0
    Next cell
    
    Application.ScreenUpdating = True
    
    If processedCount > 0 Then
        MsgBox processedCount & " positive value(s) switched to negative.", vbInformation
    Else
        MsgBox "No positive values found in selection.", vbInformation
    End If
    
    Debug.Print "SwitchSignIfPositive completed - " & processedCount & " cells processed"
End Sub

Public Sub SwitchSignIfNegative(Optional control As IRibbonControl)
    Debug.Print "SwitchSignIfNegative called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim processedCount As Long
    processedCount = 0
    
    Dim cell As Range
    For Each cell In Selection
        On Error Resume Next
        
        If cell.HasFormula Then
            ' For formulas, check the result
            If IsNumeric(cell.Value) And cell.Value < 0 Then
                If SwitchFormulaSign(cell) Then
                    processedCount = processedCount + 1
                End If
            End If
        ElseIf IsNumeric(cell.Value) And cell.Value < 0 Then
            cell.Value = -cell.Value
            processedCount = processedCount + 1
        End If
        
        On Error GoTo 0
    Next cell
    
    Application.ScreenUpdating = True
    
    If processedCount > 0 Then
        MsgBox processedCount & " negative value(s) switched to positive.", vbInformation
    Else
        MsgBox "No negative values found in selection.", vbInformation
    End If
    
    Debug.Print "SwitchSignIfNegative completed - " & processedCount & " cells processed"
End Sub

' Utility functions
Public Sub MakeAllPositive(Optional control As IRibbonControl)
    Debug.Print "MakeAllPositive called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim processedCount As Long
    processedCount = 0
    
    Dim cell As Range
    For Each cell In Selection
        On Error Resume Next
        
        If cell.HasFormula Then
            If IsNumeric(cell.Value) And cell.Value < 0 Then
                If SwitchFormulaSign(cell) Then
                    processedCount = processedCount + 1
                End If
            End If
        ElseIf IsNumeric(cell.Value) And cell.Value < 0 Then
            cell.Value = Abs(cell.Value)
            processedCount = processedCount + 1
        End If
        
        On Error GoTo 0
    Next cell
    
    Application.ScreenUpdating = True
    
    If processedCount > 0 Then
        MsgBox processedCount & " value(s) made positive.", vbInformation
    Else
        MsgBox "No negative values found to convert.", vbInformation
    End If
    
    Debug.Print "MakeAllPositive completed - " & processedCount & " cells processed"
End Sub

Public Sub MakeAllNegative(Optional control As IRibbonControl)
    Debug.Print "MakeAllNegative called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim processedCount As Long
    processedCount = 0
    
    Dim cell As Range
    For Each cell In Selection
        On Error Resume Next
        
        If cell.HasFormula Then
            If IsNumeric(cell.Value) And cell.Value > 0 Then
                If SwitchFormulaSign(cell) Then
                    processedCount = processedCount + 1
                End If
            End If
        ElseIf IsNumeric(cell.Value) And cell.Value > 0 Then
            cell.Value = -Abs(cell.Value)
            processedCount = processedCount + 1
        End If
        
        On Error GoTo 0
    Next cell
    
    Application.ScreenUpdating = True
    
    If processedCount > 0 Then
        MsgBox processedCount & " value(s) made negative.", vbInformation
    Else
        MsgBox "No positive values found to convert.", vbInformation
    End If
    
    Debug.Print "MakeAllNegative completed - " & processedCount & " cells processed"
End Sub