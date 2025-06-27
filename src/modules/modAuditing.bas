' =============================================================================
' File: ModAuditing.bas
' Version: 2.0.0
' Description: Enhanced auditing functions with Macabacus-style features
' Author: XLerate Development Team
' Created: New module for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

Attribute VB_Name = "ModAuditing"
' Enhanced Auditing Module with Macabacus-style functions
Option Explicit

' === ENHANCED PRECEDENTS AND DEPENDENTS ===

Public Sub ShowAllPrecedents(Optional control As IRibbonControl)
    Debug.Print "ShowAllPrecedents called"
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Clear existing arrows first
    ActiveSheet.ClearArrows
    
    ' Show precedents for each cell in selection
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            On Error Resume Next
            cell.ShowPrecedents
            On Error GoTo 0
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    MsgBox "All precedents displayed. Use 'Clear Arrows' to remove.", vbInformation
End Sub

Public Sub ShowAllDependents(Optional control As IRibbonControl)
    Debug.Print "ShowAllDependents called"
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Clear existing arrows first
    ActiveSheet.ClearArrows
    
    ' Show dependents for each cell in selection
    Dim cell As Range
    For Each cell In Selection
        On Error Resume Next
        cell.ShowDependents
        On Error GoTo 0
    Next cell
    
    Application.ScreenUpdating = True
    
    MsgBox "All dependents displayed. Use 'Clear Arrows' to remove.", vbInformation
End Sub

Public Sub ClearArrows(Optional control As IRibbonControl)
    Debug.Print "ClearArrows called"
    
    On Error Resume Next
    ActiveSheet.ClearArrows
    On Error GoTo 0
    
    Debug.Print "All arrows cleared"
End Sub

' === UNIFORMULAS FUNCTION ===

Public Sub Uniformulas(Optional control As IRibbonControl)
    Debug.Print "Uniformulas called"
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells.", vbExclamation
        Exit Sub
    End If
    
    If Selection.Cells.Count < 2 Then
        MsgBox "Please select at least 2 cells.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Get the first cell's formula as the template
    Dim templateCell As Range
    Set templateCell = Selection.Cells(1)
    
    If Not templateCell.HasFormula Then
        MsgBox "The first cell in the selection must contain a formula.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim templateFormula As String
    templateFormula = templateCell.FormulaR1C1
    
    ' Apply the template formula to all other cells in the selection
    Dim cell As Range
    Dim changedCount As Long
    changedCount = 0
    
    For Each cell In Selection
        If Not (cell.Row = templateCell.Row And cell.Column = templateCell.Column) Then
            ' Skip the template cell itself
            If cell.FormulaR1C1 <> templateFormula Then
                cell.FormulaR1C1 = templateFormula
                changedCount = changedCount + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    MsgBox "Uniformulas complete. " & changedCount & " cells updated to match the template formula.", vbInformation
End Sub

' === ADVANCED FORMULA ANALYSIS ===

Public Sub AnalyzeFormulaComplexity(Optional control As IRibbonControl)
    Debug.Print "AnalyzeFormulaComplexity called"
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells.", vbExclamation
        Exit Sub
    End If
    
    Dim complexFormulas As Collection
    Set complexFormulas = New Collection
    
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            Dim complexity As Integer
            complexity = CalculateFormulaComplexity(cell.Formula)
            
            If complexity > 5 Then  ' Threshold for "complex"
                complexFormulas.Add cell.Address & " (Complexity: " & complexity & ")"
            End If
        End If
    Next cell
    
    If complexFormulas.Count > 0 Then
        Dim msg As String
        msg = "Complex formulas found:" & vbNewLine & vbNewLine
        
        Dim i As Integer
        For i = 1 To complexFormulas.Count
            msg = msg & complexFormulas(i) & vbNewLine
            If i >= 10 Then  ' Limit display to first 10
                msg = msg & "... and " & (complexFormulas.Count - 10) & " more"
                Exit For
            End If
        Next i
        
        MsgBox msg, vbInformation, "Formula Complexity Analysis"
    Else
        MsgBox "No complex formulas found in the selection.", vbInformation
    End If
End Sub

Private Function CalculateFormulaComplexity(formula As String) As Integer
    ' Simple complexity calculation based on various factors
    Dim complexity As Integer
    complexity = 0
    
    ' Count nested functions
    complexity = complexity + (Len(formula) - Len(Replace(formula, "(", "")))
    
    ' Count IF statements (more complex)
    complexity = complexity + (Len(UCase(formula)) - Len(Replace(UCase(formula), "IF(", ""))) * 2
    
    ' Count VLOOKUP/HLOOKUP/INDEX/MATCH (complex functions)
    Dim complexFunctions As Variant
    complexFunctions = Array("VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "SUMPRODUCT", "SUMIFS", "COUNTIFS")
    
    Dim func As Variant
    For Each func In complexFunctions
        complexity = complexity + (Len(UCase(formula)) - Len(Replace(UCase(formula), func, ""))) / Len(func) * 3
    Next func
    
    ' Count array formulas
    If Left(formula, 1) = "{" Then complexity = complexity + 5
    
    CalculateFormulaComplexity = complexity
End Function

' === FORMULA RELATIONSHIP MAPPING ===

Public Sub MapFormulaRelationships(Optional control As IRibbonControl)
    Debug.Print "MapFormulaRelationships called"
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells.", vbExclamation
        Exit Sub
    End If
    
    ' Create a new worksheet for the relationship map
    Dim mapSheet As Worksheet
    Set mapSheet = ActiveWorkbook.Worksheets.Add
    mapSheet.Name = "Formula_Map_" & Format(Now, "hhmmss")
    
    ' Headers
    mapSheet.Cells(1, 1).Value = "Cell Address"
    mapSheet.Cells(1, 2).Value = "Formula"
    mapSheet.Cells(1, 3).Value = "Precedents"
    mapSheet.Cells(1, 4).Value = "Dependents"
    mapSheet.Cells(1, 5).Value = "Complexity"
    
    ' Format headers
    With mapSheet.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    Dim row As Long
    row = 2
    
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            mapSheet.Cells(row, 1).Value = cell.Address(External:=True)
            mapSheet.Cells(row, 2).Value = cell.Formula
            mapSheet.Cells(row, 3).Value = GetPrecedentsString(cell)
            mapSheet.Cells(row, 4).Value = GetDependentsString(cell)
            mapSheet.Cells(row, 5).Value = CalculateFormulaComplexity(cell.Formula)
            row = row + 1
        End If
    Next cell
    
    ' Auto-fit columns
    mapSheet.Columns("A:E").AutoFit
    
    MsgBox "Formula relationship map created in worksheet: " & mapSheet.Name, vbInformation
End Sub

Private Function GetPrecedentsString(cell As Range) As String
    On Error Resume Next
    
    Dim precedents As String
    Dim precedentRange As Range
    
    ' This is a simplified version - Excel's precedent detection is complex
    ' For a full implementation, you'd need to parse the formula
    
    If cell.HasFormula Then
        Dim formula As String
        formula = cell.Formula
        
        ' Look for cell references in the formula
        Dim regEx As Object
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.Pattern = "[$]?[A-Za-z]+[$]?[0-9]+"
        
        Dim matches As Object
        Set matches = regEx.Execute(formula)
        
        Dim i As Integer
        For i = 0 To matches.Count - 1
            If i > 0 Then precedents = precedents & ", "
            precedents = precedents & matches(i).Value
        Next i
    End If
    
    GetPrecedentsString = precedents
    On Error GoTo 0
End Function

Private Function GetDependentsString(cell As Range) As String
    On Error Resume Next
    
    ' This would require scanning the entire worksheet for formulas that reference this cell
    ' For performance reasons, we'll return a placeholder
    GetDependentsString = "[Scan required]"
    
    On Error GoTo 0
End Function

' === FORMULA VALIDATION ===

Public Sub ValidateFormulas(Optional control As IRibbonControl)
    Debug.Print "ValidateFormulas called"
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells.", vbExclamation
        Exit Sub
    End If
    
    Dim errorCells As Collection
    Set errorCells = New Collection
    
    Dim warningCells As Collection
    Set warningCells = New Collection
    
    Application.ScreenUpdating = False
    
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            ' Check for errors
            If IsError(cell.Value) Then
                errorCells.Add cell.Address & " - " & CStr(cell.Value)
            End If
            
            ' Check for potential issues
            If InStr(UCase(cell.Formula), "VLOOKUP") > 0 And InStr(cell.Formula, ",0)") = 0 And InStr(cell.Formula, ",FALSE)") = 0 Then
                warningCells.Add cell.Address & " - VLOOKUP without exact match"
            End If
            
            If InStr(cell.Formula, "#REF!") > 0 Then
                errorCells.Add cell.Address & " - Contains #REF! error"
            End If
            
            ' Check for circular references (simplified)
            If InStr(cell.Formula, cell.Address) > 0 Then
                warningCells.Add cell.Address & " - Potential circular reference"
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    ' Display results
    Dim msg As String
    msg = "Formula Validation Results:" & vbNewLine & vbNewLine
    
    If errorCells.Count > 0 Then
        msg = msg & "ERRORS FOUND:" & vbNewLine
        Dim i As Integer
        For i = 1 To errorCells.Count
            msg = msg & "• " & errorCells(i) & vbNewLine
            If i >= 5 Then
                msg = msg & "... and " & (errorCells.Count - 5) & " more errors" & vbNewLine
                Exit For
            End If
        Next i
        msg = msg & vbNewLine
    End If
    
    If warningCells.Count > 0 Then
        msg = msg & "WARNINGS:" & vbNewLine
        For i = 1 To warningCells.Count
            msg = msg & "• " & warningCells(i) & vbNewLine
            If i >= 5 Then
                msg = msg & "... and " & (warningCells.Count - 5) & " more warnings" & vbNewLine
                Exit For
            End If
        Next i
    End If
    
    If errorCells.Count = 0 And warningCells.Count = 0 Then
        msg = msg & "No issues found. All formulas appear to be working correctly."
    End If
    
    MsgBox msg, IIf(errorCells.Count > 0, vbCritical, IIf(warningCells.Count > 0, vbExclamation, vbInformation)), "Formula Validation"
End Sub

' === FORMULA OPTIMIZATION SUGGESTIONS ===

Public Sub SuggestFormulaOptimizations(Optional control As IRibbonControl)
    Debug.Print "SuggestFormulaOptimizations called"
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells.", vbExclamation
        Exit Sub
    End If
    
    Dim suggestions As Collection
    Set suggestions = New Collection
    
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            Dim formula As String
            formula = UCase(cell.Formula)
            
            ' Check for optimization opportunities
            If InStr(formula, "VLOOKUP") > 0 Then
                suggestions.Add cell.Address & " - Consider using INDEX/MATCH instead of VLOOKUP for better performance"
            End If
            
            If InStr(formula, "SUMPRODUCT") > 0 And InStr(formula, "--") > 0 Then
                suggestions.Add cell.Address & " - SUMPRODUCT with double negative can be slow; consider SUMIFS"
            End If
            
            If InStr(formula, "INDIRECT") > 0 Then
                suggestions.Add cell.Address & " - INDIRECT is volatile and slows calculation; consider alternatives"
            End If
            
            If InStr(formula, "OFFSET") > 0 Then
                suggestions.Add cell.Address & " - OFFSET is volatile; consider using structured references"
            End If
            
            ' Check for array formulas that could be optimized
            If Left(cell.Formula, 1) = "{" And Right(cell.Formula, 1) = "}" Then
                suggestions.Add cell.Address & " - Array formula detected; verify if it can be simplified"
            End If
        End If
    Next cell
    
    ' Display suggestions
    If suggestions.Count > 0 Then
        Dim msg As String
        msg = "Formula Optimization Suggestions:" & vbNewLine & vbNewLine
        
        Dim i As Integer
        For i = 1 To suggestions.Count
            msg = msg & "• " & suggestions(i) & vbNewLine & vbNewLine
            If i >= 8 Then  ' Limit display
                msg = msg & "... and " & (suggestions.Count - 8) & " more suggestions"
                Exit For
            End If
        Next i
        
        MsgBox msg, vbInformation, "Optimization Suggestions"
    Else
        MsgBox "No optimization suggestions found for the selected formulas.", vbInformation
    End If
End Sub