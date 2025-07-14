' =============================================================================
' File: ModErrorWrap.bas
' Version: 2.0.0
' Description: Error wrapping functions for Macabacus-style formula protection
' Author: XLerate Development Team
' Created: Enhanced for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

Attribute VB_Name = "ModErrorWrap"
Option Explicit

Public Sub WrapWithError(Optional control As IRibbonControl)
    Debug.Print "WrapWithError called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim wrappedCount As Long
    wrappedCount = 0
    
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            Dim originalFormula As String
            originalFormula = cell.Formula
            
            ' Check if already wrapped with IFERROR
            If Not IsAlreadyWrapped(originalFormula) Then
                ' Wrap with IFERROR
                cell.Formula = "=IFERROR(" & Mid(originalFormula, 2) & ","""")"
                wrappedCount = wrappedCount + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    If wrappedCount > 0 Then
        MsgBox wrappedCount & " formula(s) wrapped with IFERROR.", vbInformation
    Else
        MsgBox "No formulas to wrap or all formulas already wrapped.", vbInformation
    End If
    
    Debug.Print "WrapWithError completed - " & wrappedCount & " formulas wrapped"
End Sub

Public Sub WrapWithIfNA(Optional control As IRibbonControl)
    Debug.Print "WrapWithIfNA called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim wrappedCount As Long
    wrappedCount = 0
    
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            Dim originalFormula As String
            originalFormula = cell.Formula
            
            ' Check if already wrapped
            If Not IsAlreadyWrappedWithIfNA(originalFormula) Then
                ' Wrap with IFNA (for #N/A errors specifically)
                cell.Formula = "=IFNA(" & Mid(originalFormula, 2) & ","""")"
                wrappedCount = wrappedCount + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    If wrappedCount > 0 Then
        MsgBox wrappedCount & " formula(s) wrapped with IFNA.", vbInformation
    Else
        MsgBox "No formulas to wrap or all formulas already wrapped.", vbInformation
    End If
    
    Debug.Print "WrapWithIfNA completed - " & wrappedCount & " formulas wrapped"
End Sub

Public Sub UnwrapFormulas(Optional control As IRibbonControl)
    Debug.Print "UnwrapFormulas called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim unwrappedCount As Long
    unwrappedCount = 0
    
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            Dim originalFormula As String
            originalFormula = cell.Formula
            
            Dim unwrappedFormula As String
            unwrappedFormula = UnwrapErrorFormula(originalFormula)
            
            If unwrappedFormula <> originalFormula Then
                cell.Formula = unwrappedFormula
                unwrappedCount = unwrappedCount + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    If unwrappedCount > 0 Then
        MsgBox unwrappedCount & " formula(s) unwrapped.", vbInformation
    Else
        MsgBox "No wrapped formulas found to unwrap.", vbInformation
    End If
    
    Debug.Print "UnwrapFormulas completed - " & unwrappedCount & " formulas unwrapped"
End Sub

Public Sub WrapWithCustomError(Optional control As IRibbonControl)
    Debug.Print "WrapWithCustomError called"
    
    If Selection Is Nothing Then Exit Sub
    
    ' Get custom error message from user
    Dim errorMessage As String
    errorMessage = InputBox("Enter custom error message:", "Custom Error Wrap", "N/A")
    
    If errorMessage = "" Then Exit Sub  ' User cancelled
    
    Application.ScreenUpdating = False
    
    Dim wrappedCount As Long
    wrappedCount = 0
    
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            Dim originalFormula As String
            originalFormula = cell.Formula
            
            ' Check if already wrapped
            If Not IsAlreadyWrapped(originalFormula) Then
                ' Wrap with custom error message
                cell.Formula = "=IFERROR(" & Mid(originalFormula, 2) & ",""" & errorMessage & """)"
                wrappedCount = wrappedCount + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    If wrappedCount > 0 Then
        MsgBox wrappedCount & " formula(s) wrapped with custom error message.", vbInformation
    Else
        MsgBox "No formulas to wrap or all formulas already wrapped.", vbInformation
    End If
    
    Debug.Print "WrapWithCustomError completed - " & wrappedCount & " formulas wrapped"
End Sub

Public Sub WrapWithZero(Optional control As IRibbonControl)
    Debug.Print "WrapWithZero called"
    
    If Selection Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim wrappedCount As Long
    wrappedCount = 0
    
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            Dim originalFormula As String
            originalFormula = cell.Formula
            
            ' Check if already wrapped
            If Not IsAlreadyWrapped(originalFormula) Then
                ' Wrap with zero as error value
                cell.Formula = "=IFERROR(" & Mid(originalFormula, 2) & ",0)"
                wrappedCount = wrappedCount + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    If wrappedCount > 0 Then
        MsgBox wrappedCount & " formula(s) wrapped with zero error value.", vbInformation
    Else
        MsgBox "No formulas to wrap or all formulas already wrapped.", vbInformation
    End If
    
    Debug.Print "WrapWithZero completed - " & wrappedCount & " formulas wrapped"
End Sub

' === HELPER FUNCTIONS ===

Private Function IsAlreadyWrapped(formula As String) As Boolean
    ' Check if formula is already wrapped with IFERROR or IFNA
    Dim upperFormula As String
    upperFormula = UCase(formula)
    
    IsAlreadyWrapped = (Left(upperFormula, 9) = "=IFERROR(" Or Left(upperFormula, 6) = "=IFNA(")
End Function

Private Function IsAlreadyWrappedWithIfNA(formula As String) As Boolean
    ' Check if formula is already wrapped with IFNA
    Dim upperFormula As String
    upperFormula = UCase(formula)
    
    IsAlreadyWrappedWithIfNA = (Left(upperFormula, 6) = "=IFNA(")
End Function

Private Function UnwrapErrorFormula(formula As String) As String
    ' Remove IFERROR or IFNA wrapping from a formula
    Dim upperFormula As String
    upperFormula = UCase(formula)
    
    If Left(upperFormula, 9) = "=IFERROR(" Then
        ' Extract the inner formula from IFERROR(innerformula, errorvalue)
        Dim innerFormula As String
        innerFormula = Mid(formula, 10)  ' Remove "=IFERROR("
        
        ' Find the last comma that separates the formula from the error value
        Dim lastCommaPos As Integer
        lastCommaPos = FindLastCommaInIFERROR(innerFormula)
        
        If lastCommaPos > 0 Then
            innerFormula = Left(innerFormula, lastCommaPos - 1)
            UnwrapErrorFormula = "=" & innerFormula
        Else
            UnwrapErrorFormula = formula  ' Couldn't parse, return original
        End If
        
    ElseIf Left(upperFormula, 6) = "=IFNA(" Then
        ' Extract the inner formula from IFNA(innerformula, errorvalue)
        innerFormula = Mid(formula, 7)  ' Remove "=IFNA("
        
        ' Find the last comma
        lastCommaPos = FindLastCommaInIFERROR(innerFormula)
        
        If lastCommaPos > 0 Then
            innerFormula = Left(innerFormula, lastCommaPos - 1)
            UnwrapErrorFormula = "=" & innerFormula
        Else
            UnwrapErrorFormula = formula  ' Couldn't parse, return original
        End If
        
    Else
        UnwrapErrorFormula = formula  ' Not wrapped
    End If
End Function

Private Function FindLastCommaInIFERROR(formula As String) As Integer
    ' Find the comma that separates the main formula from the error value
    ' Need to account for nested functions and quoted strings
    
    Dim i As Integer
    Dim parenCount As Integer
    Dim inQuotes As Boolean
    Dim lastCommaPos As Integer
    
    parenCount = 0
    inQuotes = False
    lastCommaPos = 0
    
    For i = Len(formula) To 1 Step -1
        Dim char As String
        char = Mid(formula, i, 1)
        
        If char = """" Then
            inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If char = ")" Then
                parenCount = parenCount + 1
            ElseIf char = "(" Then
                parenCount = parenCount - 1
            ElseIf char = "," And parenCount = 0 Then
                lastCommaPos = i
                Exit For
            End If
        End If
    Next i
    
    FindLastCommaInIFERROR = lastCommaPos
End Function

' === ERROR ANALYSIS FUNCTIONS ===

Public Sub AnalyzeErrors(Optional control As IRibbonControl)
    Debug.Print "AnalyzeErrors called"
    
    If Selection Is Nothing Then Exit Sub
    
    Dim errorCells As Collection
    Set errorCells = New Collection
    
    Dim cell As Range
    For Each cell In Selection
        If IsError(cell.Value) Then
            errorCells.Add cell.Address & " - " & CStr(cell.Value)
        End If
    Next cell
    
    If errorCells.Count > 0 Then
        Dim msg As String
        msg = "Errors found in selection:" & vbNewLine & vbNewLine
        
        Dim i As Integer
        For i = 1 To errorCells.Count
            msg = msg & errorCells(i) & vbNewLine
            If i >= 10 Then  ' Limit display
                msg = msg & "... and " & (errorCells.Count - 10) & " more errors"
                Exit For
            End If
        Next i
        
        MsgBox msg, vbExclamation, "Error Analysis"
    Else
        MsgBox "No errors found in the selection.", vbInformation
    End If
End Sub

Public Sub WrapAllErrorsInSheet(Optional control As IRibbonControl)
    Debug.Print "WrapAllErrorsInSheet called"
    
    If MsgBox("This will wrap ALL formulas in the current sheet with IFERROR. Continue?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Wrapping formulas with IFERROR..."
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim wrappedCount As Long
    wrappedCount = 0
    
    ' Process all cells with formulas in the used range
    On Error Resume Next
    Dim formulaCells As Range
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    
    If Not formulaCells Is Nothing Then
        Dim cell As Range
        For Each cell In formulaCells
            Dim originalFormula As String
            originalFormula = cell.Formula
            
            If Not IsAlreadyWrapped(originalFormula) Then
                cell.Formula = "=IFERROR(" & Mid(originalFormula, 2) & ","""")"
                wrappedCount = wrappedCount + 1
                
                ' Update status every 100 cells
                If wrappedCount Mod 100 = 0 Then
                    Application.StatusBar = "Wrapped " & wrappedCount & " formulas..."
                End If
            End If
        Next cell
    End If
    
    On Error GoTo 0
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Completed! " & wrappedCount & " formulas wrapped with IFERROR.", vbInformation
    Debug.Print "WrapAllErrorsInSheet completed - " & wrappedCount & " formulas wrapped"
End Sub