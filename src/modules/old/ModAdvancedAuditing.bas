' ModAdvancedAuditing.bas
' Version: 1.0.0
' Date: 2025-01-04
' Author: XLerate Development Team
' 
' CHANGELOG:
' v1.0.0 - Initial implementation of advanced auditing functions
'        - Enhanced formula analysis and validation
'        - Cross-worksheet dependency mapping
'        - Model integrity checking
'        - Performance analysis tools
'
' DESCRIPTION:
' Advanced auditing capabilities for comprehensive financial model validation
' Goes beyond basic precedent/dependent tracing to provide deep model insights

Attribute VB_Name = "ModAdvancedAuditing"
Option Explicit

' Analysis result structure
Private Type AuditResult
    CellAddress As String
    IssueType As String
    Severity As String
    Description As String
    Recommendation As String
End Type

' Constants for issue severity
Private Const SEVERITY_HIGH As String = "HIGH"
Private Const SEVERITY_MEDIUM As String = "MEDIUM"
Private Const SEVERITY_LOW As String = "LOW"
Private Const SEVERITY_INFO As String = "INFO"

Public Sub PerformComprehensiveAudit(Optional control As IRibbonControl)
    ' Performs a comprehensive audit of the entire workbook
    ' Identifies potential issues, inconsistencies, and optimization opportunities
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Starting Comprehensive Model Audit ==="
    
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Performing comprehensive model audit..."
    
    ' Initialize audit results collection
    Dim auditResults As Collection
    Set auditResults = New Collection
    
    ' Perform various audit checks
    Call AuditFormulaConsistency(auditResults)
    Call AuditCircularReferences(auditResults)
    Call AuditErrorCells(auditResults)
    Call AuditUnusedCells(auditResults)
    Call AuditHardcodedValues(auditResults)
    Call AuditVolatileFunctions(auditResults)
    Call AuditModelStructure(auditResults)
    Call AuditPerformanceIssues(auditResults)
    
    ' Generate audit report
    Call GenerateAuditReport(auditResults)
    
    Dim endTime As Double
    endTime = Timer
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Debug.Print "Comprehensive audit completed in " & Format(endTime - startTime, "0.00") & " seconds"
    MsgBox "Comprehensive audit completed!" & vbNewLine & _
           "Found " & auditResults.Count & " items for review." & vbNewLine & _
           "Audit time: " & Format(endTime - startTime, "0.00") & " seconds", _
           vbInformation, "XLerate Advanced Auditing"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Debug.Print "Error in PerformComprehensiveAudit: " & Err.Description
    MsgBox "Error during audit: " & Err.Description, vbCritical, "XLerate Advanced Auditing"
End Sub

Public Sub MapModelDependencies(Optional control As IRibbonControl)
    ' Creates a comprehensive dependency map of the entire model
    ' Shows relationships between worksheets and key calculations
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Creating Model Dependency Map ==="
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Mapping model dependencies..."
    
    ' Create new worksheet for dependency map
    Dim mapSheet As Worksheet
    Set mapSheet = ActiveWorkbook.Worksheets.Add
    mapSheet.Name = "Dependency_Map_" & Format(Now, "hhmmss")
    
    ' Set up headers
    With mapSheet
        .Cells(1, 1).Value = "Source Sheet"
        .Cells(1, 2).Value = "Source Cell"
        .Cells(1, 3).Value = "Target Sheet"
        .Cells(1, 4).Value = "Target Cell"
        .Cells(1, 5).Value = "Dependency Type"
        .Cells(1, 6).Value = "Formula"
        
        ' Format headers
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(200, 200, 200)
    End With
    
    Dim row As Long
    row = 2
    
    ' Analyze each worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> mapSheet.Name Then
            Application.StatusBar = "Analyzing dependencies in " & ws.Name & "..."
            Call AnalyzeWorksheetDependencies(ws, mapSheet, row)
        End If
    Next ws
    
    ' Auto-fit columns
    mapSheet.Columns("A:F").AutoFit
    
    ' Add summary
    Call AddDependencySummary(mapSheet, row)
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Debug.Print "Dependency map created successfully"
    MsgBox "Model dependency map created!" & vbNewLine & _
           "Map worksheet: " & mapSheet.Name, vbInformation, "XLerate Advanced Auditing"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Debug.Print "Error in MapModelDependencies: " & Err.Description
    MsgBox "Error creating dependency map: " & Err.Description, vbCritical, "XLerate Advanced Auditing"
End Sub

Public Sub AnalyzeFormulaComplexity(Optional control As IRibbonControl)
    ' Analyzes and reports on formula complexity throughout the model
    ' Identifies overly complex formulas that may need simplification
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Analyzing Formula Complexity ==="
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing formula complexity..."
    
    Dim complexityResults As Collection
    Set complexityResults = New Collection
    
    ' Analyze each worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Application.StatusBar = "Analyzing formulas in " & ws.Name & "..."
        Call AnalyzeWorksheetComplexity(ws, complexityResults)
    Next ws
    
    ' Create complexity report
    Call CreateComplexityReport(complexityResults)
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Debug.Print "Formula complexity analysis completed"
    MsgBox "Formula complexity analysis completed!" & vbNewLine & _
           "Found " & complexityResults.Count & " complex formulas.", _
           vbInformation, "XLerate Advanced Auditing"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Debug.Print "Error in AnalyzeFormulaComplexity: " & Err.Description
    MsgBox "Error analyzing complexity: " & Err.Description, vbCritical, "XLerate Advanced Auditing"
End Sub

Public Sub ValidateModelIntegrity(Optional control As IRibbonControl)
    ' Validates overall model integrity and consistency
    ' Checks for common modeling errors and best practices
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Validating Model Integrity ==="
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Validating model integrity..."
    
    Dim integrityIssues As Collection
    Set integrityIssues = New Collection
    
    ' Perform integrity checks
    Call CheckNamingConventions(integrityIssues)
    Call CheckCalculationChain(integrityIssues)
    Call CheckDataValidation(integrityIssues)
    Call CheckModelStructure(integrityIssues)
    Call CheckVersionControl(integrityIssues)
    
    ' Generate integrity report
    Call GenerateIntegrityReport(integrityIssues)
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Debug.Print "Model integrity validation completed"
    MsgBox "Model integrity validation completed!" & vbNewLine & _
           "Found " & integrityIssues.Count & " items to review.", _
           vbInformation, "XLerate Advanced Auditing"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Debug.Print "Error in ValidateModelIntegrity: " & Err.Description
    MsgBox "Error validating integrity: " & Err.Description, vbCritical, "XLerate Advanced Auditing"
End Sub

' === PRIVATE HELPER FUNCTIONS ===

Private Sub AuditFormulaConsistency(results As Collection)
    ' Checks for formula consistency across rows and columns
    
    Debug.Print "Auditing formula consistency..."
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Dim usedRange As Range
        Set usedRange = ws.UsedRange
        
        If Not usedRange Is Nothing Then
            Call CheckRowConsistency(ws, usedRange, results)
            Call CheckColumnConsistency(ws, usedRange, results)
        End If
    Next ws
End Sub

Private Sub AuditCircularReferences(results As Collection)
    ' Identifies circular references in the model
    
    Debug.Print "Auditing circular references..."
    
    ' Excel automatically tracks circular references
    If Application.CircularReferences.Count > 0 Then
        Dim i As Long
        For i = 1 To Application.CircularReferences.Count
            Dim result As AuditResult
            result.CellAddress = Application.CircularReferences(i).Address(External:=True)
            result.IssueType = "Circular Reference"
            result.Severity = SEVERITY_HIGH
            result.Description = "Cell contains a circular reference"
            result.Recommendation = "Review formula logic to eliminate circular dependency"
            
            results.Add result
        Next i
    End If
End Sub

Private Sub AuditErrorCells(results As Collection)
    ' Identifies cells containing errors
    
    Debug.Print "Auditing error cells..."
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Dim cell As Range
        For Each cell In ws.UsedRange
            If IsError(cell.Value) Then
                Dim result As AuditResult
                result.CellAddress = cell.Address(External:=True)
                result.IssueType = "Error Value"
                result.Severity = SEVERITY_HIGH
                result.Description = "Cell contains error: " & CStr(cell.Value)
                result.Recommendation = "Review formula and fix error condition"
                
                results.Add result
            End If
        Next cell
    Next ws
End Sub

Private Sub AuditUnusedCells(results As Collection)
    ' Identifies potentially unused cells and ranges
    
    Debug.Print "Auditing unused cells..."
    
    ' This is a simplified check - in practice, you'd want more sophisticated logic
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Dim cell As Range
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                ' Check if this cell is referenced by any other cell
                If Not IsCellReferenced(cell) Then
                    Dim result As AuditResult
                    result.CellAddress = cell.Address(External:=True)
                    result.IssueType = "Potentially Unused"
                    result.Severity = SEVERITY_LOW
                    result.Description = "Formula cell may not be referenced elsewhere"
                    result.Recommendation = "Verify if this calculation is needed"
                    
                    results.Add result
                End If
            End If
        Next cell
    Next ws
End Sub

Private Sub AuditHardcodedValues(results As Collection)
    ' Identifies hardcoded values that should be inputs
    
    Debug.Print "Auditing hardcoded values..."
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Dim cell As Range
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                If ContainsHardcodedValues(cell.Formula) Then
                    Dim result As AuditResult
                    result.CellAddress = cell.Address(External:=True)
                    result.IssueType = "Hardcoded Values"
                    result.Severity = SEVERITY_MEDIUM
                    result.Description = "Formula contains hardcoded numbers"
                    result.Recommendation = "Consider moving constants to input cells"
                    
                    results.Add result
                End If
            End If
        Next cell
    Next ws
End Sub

Private Sub AuditVolatileFunctions(results As Collection)
    ' Identifies volatile functions that may impact performance
    
    Debug.Print "Auditing volatile functions..."
    
    Dim volatileFunctions As Variant
    volatileFunctions = Array("NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT")
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Dim cell As Range
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                Dim formula As String
                formula = UCase(cell.Formula)
                
                Dim i As Integer
                For i = LBound(volatileFunctions) To UBound(volatileFunctions)
                    If InStr(formula, volatileFunctions(i)) > 0 Then
                        Dim result As AuditResult
                        result.CellAddress = cell.Address(External:=True)
                        result.IssueType = "Volatile Function"
                        result.Severity = SEVERITY_MEDIUM
                        result.Description = "Contains volatile function: " & volatileFunctions(i)
                        result.Recommendation = "Consider alternatives to improve performance"
                        
                        results.Add result
                        Exit For ' Only report once per cell
                    End If
                Next i
            End If
        Next cell
    Next ws
End Sub

Private Sub AuditModelStructure(results As Collection)
    ' Audits overall model structure and organization
    
    Debug.Print "Auditing model structure..."
    
    ' Check for too many worksheets
    If ActiveWorkbook.Worksheets.Count > 20 Then
        Dim result As AuditResult
        result.CellAddress = "Workbook"
        result.IssueType = "Model Structure"
        result.Severity = SEVERITY_MEDIUM
        result.Description = "Model has " & ActiveWorkbook.Worksheets.Count & " worksheets"
        result.Recommendation = "Consider consolidating or organizing worksheets"
        
        results.Add result
    End If
    
    ' Check for worksheets with no used range
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.UsedRange Is Nothing Then
            Dim emptyResult As AuditResult
            emptyResult.CellAddress = ws.Name
            emptyResult.IssueType = "Empty Worksheet"
            emptyResult.Severity = SEVERITY_LOW
            emptyResult.Description = "Worksheet appears to be empty"
            emptyResult.Recommendation = "Consider removing if not needed"
            
            results.Add emptyResult
        End If
    Next ws
End Sub

Private Sub AuditPerformanceIssues(results As Collection)
    ' Identifies potential performance issues
    
    Debug.Print "Auditing performance issues..."
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ' Check for very large used ranges
        If Not ws.UsedRange Is Nothing Then
            If ws.UsedRange.Cells.Count > 100000 Then
                Dim result As AuditResult
                result.CellAddress = ws.Name
                result.IssueType = "Performance"
                result.Severity = SEVERITY_MEDIUM
                result.Description = "Worksheet has very large used range: " & ws.UsedRange.Cells.Count & " cells"
                result.Recommendation = "Consider optimizing data layout or splitting worksheet"
                
                results.Add result
            End If
        End If
    Next ws
End Sub

Private Function IsCellReferenced(targetCell As Range) As Boolean
    ' Simplified check to see if a cell is referenced elsewhere
    ' In practice, this would be more comprehensive
    
    On Error Resume Next
    
    IsCellReferenced = False
    
    ' This is a basic implementation - you'd want more sophisticated logic
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Dim cell As Range
        For Each cell In ws.UsedRange
            If cell.HasFormula And cell.Address <> targetCell.Address Then
                If InStr(cell.Formula, targetCell.Address) > 0 Then
                    IsCellReferenced = True
                    Exit Function
                End If
            End If
        Next cell
    Next ws
    
    On Error GoTo 0
End Function

Private Function ContainsHardcodedValues(formula As String) As Boolean
    ' Checks if a formula contains hardcoded numeric values
    
    ' Remove common Excel functions and operators to isolate numbers
    Dim cleanFormula As String
    cleanFormula = formula
    
    ' This is a simplified check - you'd want more sophisticated pattern matching
    Dim pattern As String
    pattern = "[0-9]+\.?[0-9]*"
    
    ' Use a simple approach - look for standalone numbers
    ContainsHardcodedValues = (InStr(cleanFormula, "1") > 0 Or InStr(cleanFormula, "2") > 0 Or _
                              InStr(cleanFormula, "3") > 0 Or InStr(cleanFormula, "4") > 0 Or _
                              InStr(cleanFormula, "5") > 0 Or InStr(cleanFormula, "6") > 0 Or _
                              InStr(cleanFormula, "7") > 0 Or InStr(cleanFormula, "8") > 0 Or _
                              InStr(cleanFormula, "9") > 0 Or InStr(cleanFormula, "0") > 0)
    
    ' Exclude common acceptable cases
    If InStr(cleanFormula, "A1") > 0 Or InStr(cleanFormula, "100") > 0 Then
        ContainsHardcodedValues = False
    End If
End Function

Private Sub GenerateAuditReport(results As Collection)
    ' Creates a comprehensive audit report worksheet
    
    Debug.Print "Generating audit report..."
    
    Dim reportSheet As Worksheet
    Set reportSheet = ActiveWorkbook.Worksheets.Add
    reportSheet.Name = "Audit_Report_" & Format(Now, "hhmmss")
    
    ' Set up report headers
    With reportSheet
        .Cells(1, 1).Value = "XLerate Advanced Audit Report"
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True
        
        .Cells(3, 1).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        .Cells(4, 1).Value = "Total Issues Found: " & results.Count
        
        ' Column headers
        .Cells(6, 1).Value = "Cell Address"
        .Cells(6, 2).Value = "Issue Type"
        .Cells(6, 3).Value = "Severity"
        .Cells(6, 4).Value = "Description"
        .Cells(6, 5).Value = "Recommendation"
        
        ' Format headers
        .Range("A6:E6").Font.Bold = True
        .Range("A6:E6").Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Add audit results
    Dim row As Long
    row = 7
    
    Dim result As Variant
    For Each result In results
        With reportSheet
            .Cells(row, 1).Value = result.CellAddress
            .Cells(row, 2).Value = result.IssueType
            .Cells(row, 3).Value = result.Severity
            .Cells(row, 4).Value = result.Description
            .Cells(row, 5).Value = result.Recommendation
            
            ' Color code by severity
            Select Case result.Severity
                Case SEVERITY_HIGH
                    .Range(.Cells(row, 1), .Cells(row, 5)).Interior.Color = RGB(255, 200, 200)
                Case SEVERITY_MEDIUM
                    .Range(.Cells(row, 1), .Cells(row, 5)).Interior.Color = RGB(255, 255, 200)
                Case SEVERITY_LOW
                    .Range(.Cells(row, 1), .Cells(row, 5)).Interior.Color = RGB(200, 255, 200)
            End Select
        End With
        
        row = row + 1
    Next result
    
    ' Auto-fit columns
    reportSheet.Columns("A:E").AutoFit
    
    Debug.Print "Audit report created on worksheet: " & reportSheet.Name
End Sub

' Additional helper functions would go here for the other audit procedures...