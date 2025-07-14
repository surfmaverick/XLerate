' =============================================================================
' File: ModFinancialFunctions.bas
' Version: 2.0.0
' Date: January 2025
' Author: XLerate Development Team
'
' CHANGELOG:
' v2.0.0 - Initial comprehensive financial functions module
'        - CAGR, IRR, NPV calculation utilities
'        - Financial ratio calculations
'        - Sensitivity analysis tools
'        - Macabacus-aligned functionality
'        - Cross-platform compatibility (Windows & macOS)
'        - Enhanced error handling and validation
' =============================================================================

Attribute VB_Name = "ModFinancialFunctions"
Option Explicit

' === CORE FINANCIAL CALCULATION FUNCTIONS ===

Public Sub InsertCAGRFormula(Optional control As IRibbonControl)
    ' Inserts a CAGR formula with user-friendly interface
    ' Matches Macabacus CAGR insertion functionality
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell for the CAGR formula.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    ' Get range for CAGR calculation
    Dim rangeAddress As String
    rangeAddress = InputBox("Enter the range for CAGR calculation:" & vbNewLine & _
                           "Examples:" & vbNewLine & _
                           "  A1:A10 (time series data)" & vbNewLine & _
                           "  B5,B15 (start value, end value)" & vbNewLine & _
                           "  C10:C20 (revenue growth series)", _
                           "XLerate - Insert CAGR Formula", _
                           Selection.Address)
    
    If rangeAddress = "" Then Exit Sub
    
    ' Validate the range
    Dim testRange As Range
    On Error Resume Next
    Set testRange = Range(rangeAddress)
    On Error GoTo ErrorHandler
    
    If testRange Is Nothing Then
        MsgBox "Invalid range address. Please try again.", vbExclamation, "XLerate"
        Exit Sub
    End If
    
    ' Determine CAGR formula type based on range
    Dim cagrFormula As String
    If testRange.Cells.Count = 2 Then
        ' Two-cell CAGR: start and end values
        cagrFormula = "=POWER(" & testRange.Cells(2).Address & "/" & testRange.Cells(1).Address & ",1/(ROWS(" & rangeAddress & ")-1))-1"
    ElseIf InStr(rangeAddress, ",") > 0 Then
        ' Comma-separated cells
        Dim rangeParts() As String
        rangeParts = Split(rangeAddress, ",")
        If UBound(rangeParts) = 1 Then
            cagrFormula = "=POWER(" & Trim(rangeParts(1)) & "/" & Trim(rangeParts(0)) & ",1/YEARFRAC(" & Trim(rangeParts(0)) & "," & Trim(rangeParts(1)) & "))-1"
        Else
            cagrFormula = "=CAGR(" & rangeAddress & ")"
        End If
    Else
        ' Range of cells - use custom CAGR function
        cagrFormula = "=CAGR(" & rangeAddress & ")"
    End If
    
    ' Insert the formula
    Selection.Formula = cagrFormula
    Selection.NumberFormat = "0.0%"
    
    Debug.Print "Inserted CAGR formula: " & cagrFormula
    MsgBox "CAGR formula inserted successfully." & vbNewLine & "Formula: " & cagrFormula, vbInformation, "XLerate"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertCAGRFormula: " & Err.Description
    MsgBox "Error inserting CAGR formula: " & Err.Description, vbExclamation, "XLerate"
End Sub

Public Sub InsertIRRFormula(Optional control As IRibbonControl)
    ' Inserts IRR formula with enhanced functionality
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell for the IRR formula.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Dim cashFlowRange As String
    cashFlowRange = InputBox("Enter the cash flow range for IRR calculation:" & vbNewLine & _
                            "Example: A1:A10 (include initial investment as negative)", _
                            "XLerate - Insert IRR Formula", _
                            "A1:A10")
    
    If cashFlowRange = "" Then Exit Sub
    
    ' Validate range
    On Error Resume Next
    Dim testRange As Range
    Set testRange = Range(cashFlowRange)
    On Error GoTo ErrorHandler
    
    If testRange Is Nothing Then
        MsgBox "Invalid range address.", vbExclamation, "XLerate"
        Exit Sub
    End If
    
    ' Insert IRR formula with error handling
    Dim irrFormula As String
    irrFormula = "=IFERROR(IRR(" & cashFlowRange & "),""#N/A"")"
    
    Selection.Formula = irrFormula
    Selection.NumberFormat = "0.0%"
    
    MsgBox "IRR formula inserted successfully.", vbInformation, "XLerate"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertIRRFormula: " & Err.Description
    MsgBox "Error: " & Err.Description, vbExclamation, "XLerate"
End Sub

Public Sub InsertNPVFormula(Optional control As IRibbonControl)
    ' Inserts NPV formula with discount rate input
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell for the NPV formula.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Dim discountRate As String
    discountRate = InputBox("Enter discount rate (as cell reference or percentage):" & vbNewLine & _
                           "Examples: D5, 0.1, 10%", _
                           "XLerate - Discount Rate", "10%")
    
    If discountRate = "" Then Exit Sub
    
    Dim cashFlowRange As String
    cashFlowRange = InputBox("Enter cash flow range (excluding initial investment):" & vbNewLine & _
                            "Example: B1:B10", _
                            "XLerate - Cash Flow Range", "B1:B10")
    
    If cashFlowRange = "" Then Exit Sub
    
    Dim initialInvestment As String
    initialInvestment = InputBox("Enter initial investment (cell reference or value):" & vbNewLine & _
                                "Example: A1, -1000000", _
                                "XLerate - Initial Investment", "A1")
    
    ' Build NPV formula
    Dim npvFormula As String
    If initialInvestment <> "" Then
        npvFormula = "=NPV(" & discountRate & "," & cashFlowRange & ")+" & initialInvestment
    Else
        npvFormula = "=NPV(" & discountRate & "," & cashFlowRange & ")"
    End If
    
    Selection.Formula = npvFormula
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
    
    MsgBox "NPV formula inserted successfully.", vbInformation, "XLerate"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertNPVFormula: " & Err.Description
    MsgBox "Error: " & Err.Description, vbExclamation, "XLerate"
End Sub

' === FINANCIAL RATIO CALCULATIONS ===

Public Sub InsertFinancialRatios(Optional control As IRibbonControl)
    ' Inserts common financial ratios with templates
    
    On Error GoTo ErrorHandler
    
    Dim ratioType As String
    ratioType = InputBox("Select ratio type:" & vbNewLine & _
                        "1. ROE (Return on Equity)" & vbNewLine & _
                        "2. ROA (Return on Assets)" & vbNewLine & _
                        "3. Current Ratio" & vbNewLine & _
                        "4. Debt-to-Equity" & vbNewLine & _
                        "5. P/E Ratio" & vbNewLine & _
                        "Enter number (1-5):", _
                        "XLerate - Financial Ratios", "1")
    
    If ratioType = "" Then Exit Sub
    
    Dim formula As String
    Dim formatCode As String
    
    Select Case ratioType
        Case "1" ' ROE
            formula = GetROEFormula()
            formatCode = "0.0%"
        Case "2" ' ROA
            formula = GetROAFormula()
            formatCode = "0.0%"
        Case "3" ' Current Ratio
            formula = GetCurrentRatioFormula()
            formatCode = "0.0x"
        Case "4" ' Debt-to-Equity
            formula = GetDebtToEquityFormula()
            formatCode = "0.0x"
        Case "5" ' P/E Ratio
            formula = GetPEFormula()
            formatCode = "0.0x"
        Case Else
            MsgBox "Invalid selection.", vbExclamation, "XLerate"
            Exit Sub
    End Select
    
    If formula <> "" Then
        Selection.Formula = formula
        Selection.NumberFormat = formatCode
        MsgBox "Financial ratio formula inserted.", vbInformation, "XLerate"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertFinancialRatios: " & Err.Description
End Sub

Private Function GetROEFormula() As String
    Dim netIncome As String, equity As String
    netIncome = InputBox("Enter Net Income cell reference:", "ROE - Net Income", "B10")
    If netIncome = "" Then Exit Function
    
    equity = InputBox("Enter Shareholders' Equity cell reference:", "ROE - Equity", "B20")
    If equity = "" Then Exit Function
    
    GetROEFormula = "=" & netIncome & "/" & equity
End Function

Private Function GetROAFormula() As String
    Dim netIncome As String, assets As String
    netIncome = InputBox("Enter Net Income cell reference:", "ROA - Net Income", "B10")
    If netIncome = "" Then Exit Function
    
    assets = InputBox("Enter Total Assets cell reference:", "ROA - Assets", "B15")
    If assets = "" Then Exit Function
    
    GetROAFormula = "=" & netIncome & "/" & assets
End Function

Private Function GetCurrentRatioFormula() As String
    Dim currentAssets As String, currentLiabilities As String
    currentAssets = InputBox("Enter Current Assets cell reference:", "Current Ratio - Assets", "B5")
    If currentAssets = "" Then Exit Function
    
    currentLiabilities = InputBox("Enter Current Liabilities cell reference:", "Current Ratio - Liabilities", "B25")
    If currentLiabilities = "" Then Exit Function
    
    GetCurrentRatioFormula = "=" & currentAssets & "/" & currentLiabilities
End Function

Private Function GetDebtToEquityFormula() As String
    Dim totalDebt As String, equity As String
    totalDebt = InputBox("Enter Total Debt cell reference:", "D/E - Debt", "B30")
    If totalDebt = "" Then Exit Function
    
    equity = InputBox("Enter Shareholders' Equity cell reference:", "D/E - Equity", "B20")
    If equity = "" Then Exit Function
    
    GetDebtToEquityFormula = "=" & totalDebt & "/" & equity
End Function

Private Function GetPEFormula() As String
    Dim stockPrice As String, eps As String
    stockPrice = InputBox("Enter Stock Price cell reference:", "P/E - Price", "B2")
    If stockPrice = "" Then Exit Function
    
    eps = InputBox("Enter Earnings Per Share cell reference:", "P/E - EPS", "B12")
    If eps = "" Then Exit Function
    
    GetPEFormula = "=" & stockPrice & "/" & eps
End Function

' === SENSITIVITY ANALYSIS TOOLS ===

Public Sub CreateSensitivityTable(Optional control As IRibbonControl)
    ' Creates a sensitivity analysis table
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    
    Dim formulaCell As String
    formulaCell = InputBox("Enter the formula cell to analyze:", "Sensitivity Analysis", Selection.Address)
    If formulaCell = "" Then Exit Sub
    
    Dim inputCell1 As String
    inputCell1 = InputBox("Enter first input cell reference:", "Input 1", "A1")
    If inputCell1 = "" Then Exit Sub
    
    Dim inputCell2 As String
    inputCell2 = InputBox("Enter second input cell reference:", "Input 2", "A2")
    If inputCell2 = "" Then Exit Sub
    
    ' Create data table structure
    Dim startRow As Long, startCol As Long
    startRow = Selection.Row
    startCol = Selection.Column
    
    ' Set up headers
    Selection.Offset(0, 0).Value = "Sensitivity Analysis"
    Selection.Offset(1, 1).Value = "Input 2 →"
    Selection.Offset(2, 0).Value = "Input 1 ↓"
    
    ' Formula reference cell
    Selection.Offset(1, 0).Formula = "=" & formulaCell
    
    MsgBox "Sensitivity table structure created. " & vbNewLine & _
           "Now add your input values and use Data Table feature.", vbInformation, "XLerate"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in CreateSensitivityTable: " & Err.Description
End Sub

' === VALUATION HELPERS ===

Public Sub InsertMultipleAnalysis(Optional control As IRibbonControl)
    ' Inserts common trading multiples analysis
    
    On Error GoTo ErrorHandler
    
    Dim multipleType As String
    multipleType = InputBox("Select multiple type:" & vbNewLine & _
                           "1. EV/Revenue" & vbNewLine & _
                           "2. EV/EBITDA" & vbNewLine & _
                           "3. P/E Ratio" & vbNewLine & _
                           "4. P/B Ratio" & vbNewLine & _
                           "5. PEG Ratio" & vbNewLine & _
                           "Enter number (1-5):", _
                           "XLerate - Trading Multiples", "2")
    
    If multipleType = "" Then Exit Sub
    
    Dim formula As String
    
    Select Case multipleType
        Case "1" ' EV/Revenue
            formula = GetEVRevenueFormula()
        Case "2" ' EV/EBITDA
            formula = GetEVEBITDAFormula()
        Case "3" ' P/E
            formula = GetPEFormula()
        Case "4" ' P/B
            formula = GetPBFormula()
        Case "5" ' PEG
            formula = GetPEGFormula()
        Case Else
            MsgBox "Invalid selection.", vbExclamation, "XLerate"
            Exit Sub
    End Select
    
    If formula <> "" Then
        Selection.Formula = formula
        Selection.NumberFormat = "0.0x"
        MsgBox "Multiple formula inserted.", vbInformation, "XLerate"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertMultipleAnalysis: " & Err.Description
End Sub

Private Function GetEVRevenueFormula() As String
    Dim ev As String, revenue As String
    ev = InputBox("Enter Enterprise Value cell reference:", "EV/Revenue - EV", "B5")
    If ev = "" Then Exit Function
    
    revenue = InputBox("Enter Revenue cell reference:", "EV/Revenue - Revenue", "B10")
    If revenue = "" Then Exit Function
    
    GetEVRevenueFormula = "=" & ev & "/" & revenue
End Function

Private Function GetEVEBITDAFormula() As String
    Dim ev As String, ebitda As String
    ev = InputBox("Enter Enterprise Value cell reference:", "EV/EBITDA - EV", "B5")
    If ev = "" Then Exit Function
    
    ebitda = InputBox("Enter EBITDA cell reference:", "EV/EBITDA - EBITDA", "B15")
    If ebitda = "" Then Exit Function
    
    GetEVEBITDAFormula = "=" & ev & "/" & ebitda
End Function

Private Function GetPBFormula() As String
    Dim price As String, bookValue As String
    price = InputBox("Enter Stock Price cell reference:", "P/B - Price", "B2")
    If price = "" Then Exit Function
    
    bookValue = InputBox("Enter Book Value per Share cell reference:", "P/B - Book Value", "B20")
    If bookValue = "" Then Exit Function
    
    GetPBFormula = "=" & price & "/" & bookValue
End Function

Private Function GetPEGFormula() As String
    Dim pe As String, growth As String
    pe = InputBox("Enter P/E Ratio cell reference:", "PEG - P/E", "B25")
    If pe = "" Then Exit Function
    
    growth = InputBox("Enter Growth Rate cell reference (as %):", "PEG - Growth", "B30")
    If growth = "" Then Exit Function
    
    GetPEGFormula = "=" & pe & "/(" & growth & "*100)"
End Function

' === SCENARIO ANALYSIS ===

Public Sub CreateScenarioAnalysis(Optional control As IRibbonControl)
    ' Creates a scenario analysis template
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    
    Dim scenarios() As String
    ReDim scenarios(2)
    scenarios(0) = "Base Case"
    scenarios(1) = "Bull Case"
    scenarios(2) = "Bear Case"
    
    ' Create headers
    Selection.Value = "Scenario Analysis"
    Selection.Offset(1, 0).Value = "Scenario"
    Selection.Offset(1, 1).Value = "Assumptions"
    Selection.Offset(1, 2).Value = "Results"
    
    ' Add scenarios
    Dim i As Integer
    For i = 0 To UBound(scenarios)
        Selection.Offset(2 + i, 0).Value = scenarios(i)
    Next i
    
    ' Format headers
    With Selection.Resize(1, 3)
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
    End With
    
    With Selection.Offset(1, 0).Resize(1, 3)
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
    End With
    
    MsgBox "Scenario analysis template created.", vbInformation, "XLerate"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in CreateScenarioAnalysis: " & Err.Description
End Sub

' === UTILITY FUNCTIONS ===

Public Function CAGR(dataRange As Range) As Double
    ' Custom CAGR function for use in worksheets
    ' Calculates compound annual growth rate from a range of values
    
    On Error GoTo ErrorHandler
    
    If dataRange.Cells.Count < 2 Then
        CAGR = CVErr(xlErrValue)
        Exit Function
    End If
    
    Dim startValue As Double, endValue As Double
    Dim periods As Long
    
    startValue = dataRange.Cells(1).Value
    endValue = dataRange.Cells(dataRange.Cells.Count).Value
    periods = dataRange.Cells.Count - 1
    
    If startValue <= 0 Or endValue <= 0 Or periods <= 0 Then
        CAGR = CVErr(xlErrValue)
        Exit Function
    End If
    
    CAGR = (endValue / startValue) ^ (1 / periods) - 1
    
    Exit Function
    
ErrorHandler:
    CAGR = CVErr(xlErrValue)
End Function

Public Sub ClearStatusBar()
    ' Utility function to clear the status bar
    ' Called with Application.OnTime for delayed clearing
    
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub
