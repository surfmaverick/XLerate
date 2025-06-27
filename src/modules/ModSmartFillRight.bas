' =============================================================================
' File: ModSmartFillRight.bas
' Version: 2.0.0
' Description: Smart fill functions for Macabacus-style modeling efficiency
' Author: XLerate Development Team
' Created: Enhanced for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

Attribute VB_Name = "ModSmartFillRight"
Option Explicit

Public Sub SmartFillRight(Optional control As IRibbonControl)
    Debug.Print "SmartFillRight called"
    
    If Selection Is Nothing Then Exit Sub
    If Selection.Cells.Count < 2 Then
        MsgBox "Please select at least 2 cells to smart fill.", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    ' Get the leftmost cell as the source
    Dim sourceCell As Range
    Set sourceCell = Selection.Cells(1, 1)
    
    ' Determine the pattern based on the source cell
    If sourceCell.HasFormula Then
        ' Smart fill formulas
        SmartFillFormulas sourceCell, Selection
    ElseIf IsNumeric(sourceCell.Value) Then
        ' Smart fill numbers
        SmartFillNumbers sourceCell, Selection
    ElseIf IsDate(sourceCell.Value) Then
        ' Smart fill dates
        SmartFillDates sourceCell, Selection
    Else
        ' Smart fill text/series
        SmartFillText sourceCell, Selection
    End If
    
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    Debug.Print "SmartFillRight completed"
End Sub

Private Sub SmartFillFormulas(sourceCell As Range, targetRange As Range)
    Debug.Print "Smart filling formulas"
    
    Dim formula As String
    formula = sourceCell.FormulaR1C1
    
    ' Apply formula to all cells in the range
    Dim cell As Range
    For Each cell In targetRange
        If Not (cell.Row = sourceCell.Row And cell.Column = sourceCell.Column) Then
            cell.FormulaR1C1 = formula
        End If
    Next cell
End Sub

Private Sub SmartFillNumbers(sourceCell As Range, targetRange As Range)
    Debug.Print "Smart filling numbers"
    
    ' Look for a pattern in existing data
    Dim increment As Double
    increment = 1  ' Default increment
    
    ' If there are at least 2 cells with values, calculate increment
    If targetRange.Columns.Count > 1 Then
        Dim nextCell As Range
        Set nextCell = sourceCell.Offset(0, 1)
        
        If IsNumeric(nextCell.Value) And nextCell.Value <> "" Then
            increment = nextCell.Value - sourceCell.Value
        End If
    End If
    
    ' Fill the series
    Dim col As Integer
    For col = 1 To targetRange.Columns.Count
        Dim cell As Range
        Set cell = targetRange.Cells(1, col)
        If col = 1 Then
            ' Keep source value
        Else
            cell.Value = sourceCell.Value + (increment * (col - 1))
        End If
    Next col
End Sub

Private Sub SmartFillDates(sourceCell As Range, targetRange As Range)
    Debug.Print "Smart filling dates"
    
    ' Default to monthly increment
    Dim increment As Integer
    increment = 1  ' months
    
    ' Fill date series
    Dim col As Integer
    For col = 1 To targetRange.Columns.Count
        Dim cell As Range
        Set cell = targetRange.Cells(1, col)
        If col = 1 Then
            ' Keep source value
        Else
            cell.Value = DateAdd("m", col - 1, sourceCell.Value)
        End If
    Next col
End Sub

Private Sub SmartFillText(sourceCell As Range, targetRange As Range)
    Debug.Print "Smart filling text"
    
    Dim sourceText As String
    sourceText = CStr(sourceCell.Value)
    
    ' Try to detect if it's a series (Q1, Q2, etc.)
    If DetectQuarterSeries(sourceText) Then
        FillQuarterSeries sourceCell, targetRange
    ElseIf DetectMonthSeries(sourceText) Then
        FillMonthSeries sourceCell, targetRange
    Else
        ' Just copy the text
        Dim cell As Range
        For Each cell In targetRange
            If Not (cell.Row = sourceCell.Row And cell.Column = sourceCell.Column) Then
                cell.Value = sourceText
            End If
        Next cell
    End If
End Sub

Private Function DetectQuarterSeries(text As String) As Boolean
    DetectQuarterSeries = (UCase(text) Like "Q[1-4]*" Or UCase(text) Like "*Q[1-4]*")
End Function

Private Function DetectMonthSeries(text As String) As Boolean
    Dim months As Variant
    months = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    
    Dim i As Integer
    For i = LBound(months) To UBound(months)
        If InStr(UCase(text), months(i)) > 0 Then
            DetectMonthSeries = True
            Exit Function
        End If
    Next i
    
    DetectMonthSeries = False
End Function

Private Sub FillQuarterSeries(sourceCell As Range, targetRange As Range)
    ' Extract quarter number and year from source
    Dim sourceText As String
    sourceText = UCase(CStr(sourceCell.Value))
    
    Dim quarterNum As Integer
    Dim yearNum As Integer
    
    ' Simple extraction (assumes format like "Q1 2024" or "2024 Q1")
    If InStr(sourceText, "Q1") > 0 Then quarterNum = 1
    If InStr(sourceText, "Q2") > 0 Then quarterNum = 2
    If InStr(sourceText, "Q3") > 0 Then quarterNum = 3
    If InStr(sourceText, "Q4") > 0 Then quarterNum = 4
    
    ' Extract year (look for 4-digit number)
    Dim i As Integer
    For i = 1 To Len(sourceText) - 3
        If IsNumeric(Mid(sourceText, i, 4)) And Val(Mid(sourceText, i, 4)) > 2000 Then
            yearNum = Val(Mid(sourceText, i, 4))
            Exit For
        End If
    Next i
    
    ' Fill the series
    Dim col As Integer
    For col = 1 To targetRange.Columns.Count
        Dim cell As Range
        Set cell = targetRange.Cells(1, col)
        
        If col > 1 Then
            Dim newQuarter As Integer
            Dim newYear As Integer
            
            newQuarter = ((quarterNum - 1 + col - 1) Mod 4) + 1
            newYear = yearNum + Int((quarterNum - 1 + col - 1) / 4)
            
            cell.Value = "Q" & newQuarter & " " & newYear
        End If
    Next col
End Sub

Private Sub FillMonthSeries(sourceCell As Range, targetRange As Range)
    ' This would implement month series filling
    ' For now, just copy the source
    Dim cell As Range
    For Each cell In targetRange
        If Not (cell.Row = sourceCell.Row And cell.Column = sourceCell.Column) Then
            cell.Value = sourceCell.Value
        End If
    Next cell
End Sub

' Fast Fill Down function
Public Sub SmartFillDown(Optional control As IRibbonControl)
    Debug.Print "SmartFillDown called"
    
    If Selection Is Nothing Then Exit Sub
    If Selection.Rows.Count < 2 Then
        MsgBox "Please select at least 2 rows to smart fill.", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    ' Get the top cell as the source
    Dim sourceCell As Range
    Set sourceCell = Selection.Cells(1, 1)
    
    ' Fill down based on content type
    If sourceCell.HasFormula Then
        ' Fill formula down
        Dim formula As String
        formula = sourceCell.FormulaR1C1
        
        Dim cell As Range
        For Each cell In Selection
            If Not (cell.Row = sourceCell.Row And cell.Column = sourceCell.Column) Then
                cell.FormulaR1C1 = formula
            End If
        Next cell
    Else
        ' Fill value down
        Dim value As Variant
        value = sourceCell.Value
        
        For Each cell In Selection
            If Not (cell.Row = sourceCell.Row And cell.Column = sourceCell.Column) Then
                cell.Value = value
            End If
        Next cell
    End If
    
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    Debug.Print "SmartFillDown completed"
End Sub