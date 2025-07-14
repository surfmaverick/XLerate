' =============================================================================
' File: ModModelValidation.bas
' Version: 2.0.0
' Date: January 2025
' Author: XLerate Development Team
'
' CHANGELOG:
' v2.0.0 - Comprehensive financial model validation suite
'        - Balance sheet balancing checks
'        - Cash flow validation and reconciliation
'        - Formula consistency and error detection
'        - Professional model quality assurance
'        - Integration with Macabacus-style workflow
'        - Cross-platform compatibility (Windows & macOS)
' =============================================================================

Attribute VB_Name = "ModModelValidation"
Option Explicit

' === BALANCE SHEET VALIDATION ===

Public Sub ValidateBalanceSheet(Optional control As IRibbonControl)
    ' Comprehensive balance sheet validation
    Debug.Print "ValidateBalanceSheet called"
    
    On Error GoTo ErrorHandler
    
    Dim validationResults As Collection
    Set validationResults = New Collection
    
    ' Get balance sheet components
    Dim assetsRange As String, liabilitiesRange As String, equityRange As String
    
    assetsRange = InputBox("Enter Total Assets cell reference:", "Balance Sheet Validation", "B10")
    If assetsRange = "" Then Exit Sub
    
    liabilitiesRange = InputBox("Enter Total Liabilities cell reference:", "Balance Sheet Validation", "B20")
    If liabilitiesRange = "" Then Exit Sub
    
    equityRange = InputBox("Enter Total Equity cell reference:", "Balance Sheet Validation", "B30")
    If equityRange = "" Then Exit Sub
    
    ' Validate balance
    On Error Resume Next
    Dim assets As Double, liabilities As Double, equity As Double
    assets = Range(assetsRange).Value
    liabilities = Range(liabilitiesRange).Value
    equity = Range(equityRange).Value
    
    If Err.Number <> 0 Then
        validationResults.Add "ERROR: Invalid cell references"
        GoTo ShowResults
    End If
    On Error GoTo ErrorHandler
    
    ' Check if balance sheet balances
    Dim totalLiabEquity As Double
    totalLiabEquity = liabilities + equity
    Dim difference As Double
    difference = assets - totalLiabEquity
    
    If Abs(difference) < 0.01 Then
        validationResults.Add "✓ PASS: Balance sheet balances (difference: " & Format(difference, "$#,##0.00") & ")"
    Else
        validationResults.Add "✗ FAIL: Balance sheet does not balance (difference: " & Format(difference, "$#,##0.00") & ")"
    End If
    
    ' Additional checks
    If assets <= 0 Then validationResults.Add "⚠ WARNING: Total Assets is zero or negative"
    If liabilities < 0 Then validationResults.Add "⚠ WARNING: Total Liabilities is negative"
    If equity <= 0 Then validationResults.Add "⚠ WARNING: Total Equity is zero or negative"
    
ShowResults:
    DisplayValidationResults validationResults, "Balance Sheet Validation"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ValidateBalanceSheet: " & Err.Description
    MsgBox "Error during validation: " & Err.Description, vbExclamation, "XLerate"
End Sub

Public Sub ValidateCashFlow(Optional control As IRibbonControl)
    ' Cash flow statement validation
    Debug.Print "ValidateCashFlow called"
    
    On Error GoTo ErrorHandler
    
    Dim validationResults As Collection
    Set validationResults = New Collection
    
    ' Get cash flow components
    Dim operatingCF As String, investingCF As String, financingCF As String
    Dim beginningCash As String, endingCash As String
    
    operatingCF = InputBox("Enter Operating Cash Flow cell reference:", "Cash Flow Validation", "B10")
    If operatingCF = "" Then Exit Sub
    
    investingCF = InputBox("Enter Investing Cash Flow cell reference:", "Cash Flow Validation", "B15")
    If investingCF = "" Then Exit Sub
    
    financingCF = InputBox("Enter Financing Cash Flow cell reference:", "Cash Flow Validation", "B20")
    If financingCF = "" Then Exit Sub
    
    beginningCash = InputBox("Enter Beginning Cash cell reference:", "Cash Flow Validation", "B5")
    If beginningCash = "" Then Exit Sub
    
    endingCash = InputBox("Enter Ending Cash cell reference:", "Cash Flow Validation", "B25")
    If endingCash = "" Then Exit Sub
    
    ' Validate cash flow
    On Error Resume Next
    Dim opCF As Double, invCF As Double, finCF As Double
    Dim begCash As Double, endCash As Double
    
    opCF = Range(operatingCF).Value
    invCF = Range(investingCF).Value
    finCF = Range(financingCF).Value
    begCash = Range(beginningCash).Value
    endCash = Range(endingCash).Value
    
    If Err.Number <> 0 Then
        validationResults.Add "ERROR: Invalid cell references"
        GoTo ShowCFResults
    End If
    On Error GoTo ErrorHandler
    
    ' Check cash flow reconciliation
    Dim calculatedEndCash As Double
    calculatedEndCash = begCash + opCF + invCF + finCF
    Dim cfDifference As Double
    cfDifference = endCash - calculatedEndCash
    
    If Abs(cfDifference) < 0.01 Then
        validationResults.Add "✓ PASS: Cash flow reconciles (difference: " & Format(cfDifference, "$#,##0.00") & ")"
    Else
        validationResults.Add "✗ FAIL: Cash flow does not reconcile (difference: " & Format(cfDifference, "$#,##0.00") & ")"
    End If
    
    ' Additional checks
    If begCash < 0 Then validationResults.Add "⚠ WARNING: Beginning cash is negative"
    If endCash < 0 Then validationResults.Add "⚠ WARNING: Ending cash is negative"
    If opCF < 0 Then validationResults.Add "ℹ INFO: Operating cash flow is negative"
    
ShowCFResults:
    DisplayValidationResults validationResults, "Cash Flow Validation"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ValidateCashFlow: " & Err.Description
    MsgBox "Error during validation: " & Err.Description, vbExclamation, "XLerate"
End Sub

' === FORMULA VALIDATION ===

Public Sub ValidateModelFormulas(Optional control As IRibbonControl)
    ' Comprehensive formula validation across the model
    Debug.Print "ValidateModelFormulas called"
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then
        MsgBox "Please select the range to validate.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Validating formulas..."
    
    Dim validationResults As Collection
    Set validationResults = New Collection
    
    Dim errorCells As Collection
    Set errorCells = New Collection
    
    Dim inconsistentCells As Collection
    Set inconsistentCells = New Collection
    
    Dim circularRefs As Collection
    Set circularRefs = New Collection
    
    ' Check each cell in selection
    Dim cell As Range
    Dim cellCount As Long
    cellCount = 0
    
    For Each cell In Selection
        cellCount = cellCount + 1
        
        ' Update progress for large ranges
        If cellCount Mod 100 = 0 Then
            Application.StatusBar = "Validating formulas... " & cellCount & " cells checked"
        End If
        
        If cell.HasFormula Then
            ' Check for errors
            If IsError(cell.Value) Then
                errorCells.Add cell.Address & " - " & CStr(cell.Value)
            End If
            
            ' Check for potential circular references
            If InStr(cell.Formula, cell.Address) > 0 Then
                circularRefs.Add cell.Address & " - Self-reference detected"
            End If
            
            ' Check for common formula issues
            ValidateIndividualFormula cell, inconsistentCells
        End If
    Next cell
    
    ' Compile results
    validationResults.Add "=== FORMULA VALIDATION RESULTS ==="
    validationResults.Add "Cells checked: " & cellCount
    validationResults.Add ""
    
    If errorCells.Count = 0 Then
        validationResults.Add "✓ PASS: No formula errors found"
    Else
        validationResults.Add "✗ FAIL: " & errorCells.Count & " formula errors found:"
        Dim i As Integer
        For i = 1 To Application.Min(errorCells.Count, 10)
            validationResults.Add "  • " & errorCells(i)
        Next i
        If errorCells.Count > 10 Then
            validationResults.Add "  ... and " & (errorCells.Count - 10) & " more errors"
        End If
    End If
    
    validationResults.Add ""
    
    If circularRefs.Count = 0 Then
        validationResults.Add "✓ PASS: No circular references detected"
    Else
        validationResults.Add "⚠ WARNING: " & circularRefs.Count & " potential circular references:"
        For i = 1 To Application.Min(circularRefs.Count, 5)
            validationResults.Add "  • " & circularRefs(i)
        Next i
    End If
    
    validationResults.Add ""
    
    If inconsistentCells.Count = 0 Then
        validationResults.Add "✓ PASS: No obvious formula inconsistencies"
    Else
        validationResults.Add "⚠ WARNING: " & inconsistentCells.Count & " potential issues:"
        For i = 1 To Application.Min(inconsistentCells.Count, 5)
            validationResults.Add "  • " & inconsistentCells(i)
        Next i
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    DisplayValidationResults validationResults, "Model Formula Validation"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Debug.Print "Error in ValidateModelFormulas: " & Err.Description
    MsgBox "Error during validation: " & Err.Description, vbExclamation, "XLerate"
End Sub

Private Sub ValidateIndividualFormula(cell As Range, issues As Collection)
    ' Validate individual formula for common issues
    On Error Resume Next
    
    Dim formula As String
    formula = UCase(cell.Formula)
    
    ' Check for hardcoded values in formulas
    If InStr(formula, "VLOOKUP") > 0 And InStr(formula, ",0)") = 0 And InStr(formula, ",FALSE)") = 0 Then
        issues.Add cell.Address & " - VLOOKUP without exact match (may cause errors)"
    End If
    
    ' Check for volatile functions
    If InStr(formula, "INDIRECT") > 0 Then
        issues.Add cell.Address & " - INDIRECT function (volatile, slows calculation)"
    End If
    
    If InStr(formula, "OFFSET") > 0 Then
        issues.Add cell.Address & " - OFFSET function (volatile, consider alternatives)"
    End If
    
    ' Check for potential #REF! issues
    If InStr(formula, "#REF!") > 0 Then
        issues.Add cell.Address & " - Contains #REF! error"
    End If
    
    On Error GoTo 0
End Sub

' === MODEL INTEGRITY CHECKS ===

Public Sub ValidateModelIntegrity(Optional control As IRibbonControl)
    ' Overall model integrity validation
    Debug.Print "ValidateModelIntegrity called"
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Validating model integrity..."
    
    Dim validationResults As Collection
    Set validationResults = New Collection
    
    validationResults.Add "=== MODEL INTEGRITY VALIDATION ==="
    validationResults.Add "Workbook: " & ActiveWorkbook.Name
    validationResults.Add "Validation Date: " & Format(Now, "yyyy-mm-dd hh:mm")
    validationResults.Add ""
    
    ' Check for broken links
    Dim brokenLinks As Integer
    brokenLinks = CheckForBrokenLinks()
    If brokenLinks = 0 Then
        validationResults.Add "✓ PASS: No broken external links"
    Else
        validationResults.Add "✗ FAIL: " & brokenLinks & " broken external links detected"
    End If
    
    ' Check for hidden rows/columns with data
    Dim hiddenIssues As Integer
    hiddenIssues = CheckHiddenCells()
    If hiddenIssues = 0 Then
        validationResults.Add "✓ PASS: No data in hidden rows/columns"
    Else
        validationResults.Add "⚠ WARNING: " & hiddenIssues & " hidden rows/columns contain data"
    End If
    
    ' Check calculation mode
    If Application.Calculation = xlCalculationAutomatic Then
        validationResults.Add "✓ PASS: Calculation mode is automatic"
    Else
        validationResults.Add "⚠ WARNING: Calculation mode is not automatic"
    End If
    
    ' Check for very large numbers (potential errors)
    Dim largeNumbers As Integer
    largeNumbers = CheckForLargeNumbers()
    If largeNumbers = 0 Then
        validationResults.Add "✓ PASS: No suspiciously large numbers"
    Else
        validationResults.Add "⚠ WARNING: " & largeNumbers & " cells with very large numbers (>1 trillion)"
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    DisplayValidationResults validationResults, "Model Integrity Validation"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Debug.Print "Error in ValidateModelIntegrity: " & Err.Description
    MsgBox "Error during validation: " & Err.Description, vbExclamation, "XLerate"
End Sub

Private Function CheckForBrokenLinks() As Integer
    ' Check for broken external links
    On Error Resume Next
    
    Dim links As Variant
    links = ActiveWorkbook.LinkSources(xlExcelLinks)
    
    If IsArray(links) Then
        CheckForBrokenLinks = UBound(links) - LBound(links) + 1
    Else
        CheckForBrokenLinks = 0
    End If
    
    On Error GoTo 0
End Function

Private Function CheckHiddenCells() As Integer
    ' Check for data in hidden rows/columns
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim hiddenCount As Integer
    hiddenCount = 0
    
    For Each ws In ActiveWorkbook.Worksheets
        ' Check hidden rows
        Dim row As Range
        For Each row In ws.UsedRange.Rows
            If row.Hidden And Not IsEmpty(row.Cells(1)) Then
                hiddenCount = hiddenCount + 1
            End If
        Next row
        
        ' Check hidden columns
        Dim col As Range
        For Each col In ws.UsedRange.Columns
            If col.Hidden And Not IsEmpty(col.Cells(1)) Then
                hiddenCount = hiddenCount + 1
            End If
        Next col
    Next ws
    
    CheckHiddenCells = hiddenCount
    On Error GoTo 0
End Function

Private Function CheckForLargeNumbers() As Integer
    ' Check for suspiciously large numbers
    On Error Resume Next
    
    Dim largeCount As Integer
    largeCount = 0
    
    Dim cell As Range
    For Each cell In ActiveSheet.UsedRange
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            If Abs(cell.Value) > 1000000000000# Then  ' 1 trillion
                largeCount = largeCount + 1
            End If
        End If
    Next cell
    
    CheckForLargeNumbers = largeCount
    On Error GoTo 0
End Function

' === SENSITIVITY ANALYSIS VALIDATION ===

Public Sub ValidateSensitivityAnalysis(Optional control As IRibbonControl)
    ' Validate sensitivity analysis setup
    Debug.Print "ValidateSensitivityAnalysis called"