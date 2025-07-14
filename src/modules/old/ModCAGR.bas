Attribute VB_Name = "ModCAGR"
Function CAGR(rng As Range) As Double
    ' Calculate CAGR using first and last values from range and count as periods
    ' Formula: CAGR = (End Value/Start Value)^(1/n) - 1
    ' where n is the number of periods (count of items - 1)
    
    On Error GoTo ErrorHandler
    
    Dim firstValue As Double
    Dim lastValue As Double
    Dim periodCount As Long
    
    ' Get first and last values from range
    firstValue = rng.Cells(1).value
    lastValue = rng.Cells(rng.Cells.Count).value
    
    ' Calculate number of periods (count - 1 because we're measuring intervals)
    periodCount = rng.Cells.Count - 1
    
    ' Check for valid inputs
    If firstValue <= 0 Or lastValue <= 0 Then
        CAGR = CVErr(xlErrValue)
        Exit Function
    End If
    
    If periodCount <= 0 Then
        CAGR = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Calculate CAGR
    CAGR = (lastValue / firstValue) ^ (1 / periodCount) - 1
    
    Exit Function

ErrorHandler:
    CAGR = CVErr(xlErrValue)
End Function
