' =========================================================================
' CONFLICT-FREE: ModFastFillDown.bas v2.1.2 - Zero Naming Conflicts
' File: src/modules/ModFastFillDown.bas
' Version: 2.1.2 (CONFLICT-FREE)
' Date: 2025-07-06
' =========================================================================
'
' SOLUTION: No utility functions - uses inline code only
' GUARANTEED: Zero naming conflicts with existing modules
' MAINTAINED: Full Fast Fill Down functionality
' =========================================================================

Attribute VB_Name = "ModFastFillDown"
Option Explicit

Public Sub FastFillDown(Optional control As IRibbonControl)
    ' Fast Fill Down - Ctrl+Alt+Shift+D (NO CONFLICTS)
    
    On Error GoTo ErrorHandler
    
    ' Validate selection (inline - no helper functions)
    If Selection.Cells.Count = 0 Then Exit Sub
    
    Dim hasContent As Boolean
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Or (cell.Value <> "" And Not IsEmpty(cell.Value)) Then
            hasContent = True
            Exit For
        End If
    Next cell
    
    If Not hasContent Then
        MsgBox "Selection must contain formulas or values to fill down.", vbInformation
        Exit Sub
    End If
    
    ' Check for merged cells
    For Each cell In Selection
        If cell.MergeArea.Cells.Count > 1 Then
            MsgBox "Cannot fill down merged cells.", vbInformation
            Exit Sub
        End If
    Next cell
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Fast filling down..."
    
    ' Find boundary (simplified logic - no helper functions)
    Dim sourceRange As Range
    Set sourceRange = Selection
    
    Dim startRow As Long
    Dim lastRow As Long
    Dim checkRow As Long
    Dim checkCol As Long
    
    startRow = sourceRange.Row + sourceRange.Rows.Count
    lastRow = startRow - 1
    
    ' Look 3 columns to the left for patterns
    For checkCol = sourceRange.Column - 1 To Application.WorksheetFunction.Max(sourceRange.Column - 3, 1) Step -1
        For checkRow = startRow To startRow + 50
            If Not IsEmpty(Cells(checkRow, checkCol).Value) Or Cells(checkRow, checkCol).HasFormula Then
                lastRow = Application.WorksheetFunction.Max(lastRow, checkRow)
            ElseIf lastRow >= startRow Then
                Exit For
            End If
        Next checkRow
        
        If lastRow >= startRow + 2 Then Exit For
    Next checkCol
    
    ' Perform fill if boundary found
    If lastRow > startRow Then
        Dim targetRange As Range
        Set targetRange = Range(sourceRange, Cells(lastRow, sourceRange.Column + sourceRange.Columns.Count - 1))
        sourceRange.AutoFill Destination:=targetRange, Type:=xlFillDefault
        
        Application.StatusBar = "Filled " & (lastRow - startRow + 1) & " rows down"
        targetRange.Select
    Else
        Application.StatusBar = "No boundary found for fill down"
    End If
    
    ' Clear status bar (inline - no function calls)
    DoEvents: Application.Wait Now + TimeValue("00:00:01"): Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = "Fill down failed: " & Err.Description
    DoEvents: Application.Wait Now + TimeValue("00:00:01"): Application.StatusBar = False
End Sub