' =========================================================================
' CONFLICT-FREE: ModCurrencyCycling.bas v2.1.2 - Zero Naming Conflicts
' File: src/modules/ModCurrencyCycling.bas
' Version: 2.1.2 (CONFLICT-FREE)
' Date: 2025-07-06
' =========================================================================
'
' SOLUTION: No utility functions - uses inline code only
' GUARANTEED: Zero naming conflicts with existing modules
' MAINTAINED: Full currency cycling functionality (20+ formats)
' =========================================================================

Attribute VB_Name = "ModCurrencyCycling"
Option Explicit

' Static variables for cycling state
Private currentCurrencyIndex As Integer

Public Sub CycleCurrency(Optional control As IRibbonControl)
    ' Currency Cycling - Ctrl+Alt+Shift+6 (NO CONFLICTS)
    
    On Error GoTo ErrorHandler
    
    If Selection.Cells.Count = 0 Then Exit Sub
    
    Application.StatusBar = "Cycling currency formats..."
    
    ' Currency formats array (inline - no initialization functions)
    Dim formats As Variant
    Dim formatNames As Variant
    
    formats = Array( _
        "$#,##0", "$#,##0.00", "$#,##0_);($#,##0)", "$#,##0.00_);($#,##0.00)", _
        "€#,##0", "€#,##0.00", "€#,##0_);(€#,##0)", "€#,##0.00_);(€#,##0.00)", _
        "£#,##0", "£#,##0.00", "£#,##0_);(£#,##0)", "£#,##0.00_);(£#,##0.00)", _
        "¥#,##0", "¥#,##0.00", "₹#,##0.00", "₩#,##0", _
        "C$#,##0.00", "A$#,##0.00", "#,##0.00 ""CHF""", "R$#,##0.00" _
    )
    
    formatNames = Array( _
        "USD Simple", "USD Decimals", "USD Negative", "USD Decimal Negative", _
        "EUR Simple", "EUR Decimals", "EUR Negative", "EUR Decimal Negative", _
        "GBP Simple", "GBP Decimals", "GBP Negative", "GBP Decimal Negative", _
        "JPY", "CNY", "INR", "KRW", _
        "CAD", "AUD", "CHF", "BRL" _
    )
    
    ' Apply current format
    Selection.NumberFormat = formats(currentCurrencyIndex)
    
    ' Update status with format name
    Application.StatusBar = "Applied: " & formatNames(currentCurrencyIndex)
    
    ' Move to next format
    currentCurrencyIndex = (currentCurrencyIndex + 1) Mod (UBound(formats) + 1)
    
    ' Clear status bar (inline - no function calls)
    DoEvents: Application.Wait Now + TimeValue("00:00:01"): Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Currency cycling failed: " & Err.Description
    DoEvents: Application.Wait Now + TimeValue("00:00:01"): Application.StatusBar = False
End Sub

Public Sub ShowCurrencyFormats()
    ' Display available currency formats (NO CONFLICTS)
    
    Dim helpText As String
    helpText = "XLerate Currency Cycling" & vbCrLf & vbCrLf
    helpText = helpText & "Shortcut: Ctrl+Alt+Shift+6" & vbCrLf & vbCrLf
    helpText = helpText & "Available Formats:" & vbCrLf
    helpText = helpText & "• USD (4 variations)" & vbCrLf
    helpText = helpText & "• EUR (4 variations)" & vbCrLf
    helpText = helpText & "• GBP (4 variations)" & vbCrLf
    helpText = helpText & "• JPY, CNY, INR, KRW" & vbCrLf
    helpText = helpText & "• CAD, AUD, CHF, BRL" & vbCrLf & vbCrLf
    helpText = helpText & "Usage:" & vbCrLf
    helpText = helpText & "1. Select cells" & vbCrLf
    helpText = helpText & "2. Press Ctrl+Alt+Shift+6" & vbCrLf
    helpText = helpText & "3. Repeat to cycle through formats"
    
    MsgBox helpText, vbInformation, "Currency Cycling Help"
End Sub