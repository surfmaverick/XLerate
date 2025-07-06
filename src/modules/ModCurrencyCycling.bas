' =========================================================================
' FIXED: ModCurrencyCycling.bas v2.1.1 - Resolved Naming Conflicts
' File: src/modules/ModCurrencyCycling.bas
' Version: 2.1.1 (FIXED - No naming conflicts)
' Date: 2025-07-06
' Author: XLerate Development Team
' =========================================================================
'
' CHANGELOG v2.1.1:
' - FIXED: Removed duplicate ClearStatusBarDelayed function
' - FIXED: Uses existing ModUtilityFunctions.ClearStatusBar instead
' - RESOLVED: "Ambiguous name detected" compilation errors
' - MAINTAINED: All currency cycling functionality
' - PRESERVED: Integration with existing utility functions
'
' CHANGES FROM v2.1.0:
' - Removed ClearStatusBarDelayed (uses existing ClearStatusBar)
' - Updated all status bar clearing to use ModUtilityFunctions.ClearStatusBar
' - Maintained all core currency cycling logic and 20+ currency formats
' =========================================================================

Attribute VB_Name = "ModCurrencyCycling"
Option Explicit

' Module constants
Private Const MODULE_VERSION As String = "2.1.1"
Private Const SETTINGS_KEY As String = "CurrencySettings"

' Currency cycling state
Private currentCurrencyIndex As Integer
Private currentFormatIndex As Integer

' Currency format collections
Private Type CurrencyFormat
    Name As String
    Symbol As String
    Code As String
    FormatPositive As String
    FormatNegative As String
    FormatZero As String
End Type

Private CurrencyFormats() As CurrencyFormat
Private bFormatsInitialized As Boolean

' =========================================================================
' PUBLIC INTERFACE - Called by shortcut Ctrl+Alt+Shift+6
' =========================================================================

Public Sub CycleCurrency(Optional control As IRibbonControl)
    ' Main Currency Cycling function - Enhanced beyond Macabacus
    ' Shortcut: Ctrl+Alt+Shift+6
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== CycleCurrency v" & MODULE_VERSION & " Started ==="
    
    ' Validate selection
    If Selection.Cells.Count = 0 Then
        Debug.Print "No cells selected"
        Exit Sub
    End If
    
    ' Initialize currency formats if needed
    If Not bFormatsInitialized Then
        Call InitializeCurrencyFormats
    End If
    
    Application.StatusBar = "XLerate: Cycling currency formats..."
    
    ' Get current format and cycle to next
    Dim newFormat As String
    newFormat = GetNextCurrencyFormat()
    
    ' Apply to selection
    Selection.NumberFormat = newFormat
    
    ' Provide user feedback
    Dim currencyInfo As String
    currencyInfo = GetCurrentCurrencyInfo()
    
    Application.StatusBar = "XLerate: Applied " & currencyInfo
    Call UseExistingClearStatusBar
    
    Debug.Print "CycleCurrency completed. Applied format: " & newFormat
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "XLerate: Currency cycling failed - " & Err.Description
    Call UseExistingClearStatusBar
    Debug.Print "Error in CycleCurrency: " & Err.Description & " (Error " & Err.Number & ")"
End Sub

' =========================================================================
' FORMAT INITIALIZATION
' =========================================================================

Private Sub InitializeCurrencyFormats()
    ' Initialize the collection of currency formats
    
    Debug.Print "Initializing currency formats..."
    
    ReDim CurrencyFormats(0 To 19) ' 20 currency formats
    
    ' US Dollar formats
    With CurrencyFormats(0)
        .Name = "US Dollar - Simple"
        .Symbol = "$"
        .Code = "USD"
        .FormatPositive = "$#,##0"
        .FormatNegative = "$#,##0_);($#,##0)"
        .FormatZero = "$#,##0"
    End With
    
    With CurrencyFormats(1)
        .Name = "US Dollar - Decimals"
        .Symbol = "$"
        .Code = "USD"
        .FormatPositive = "$#,##0.00"
        .FormatNegative = "$#,##0.00_);($#,##0.00)"
        .FormatZero = "$#,##0.00"
    End With
    
    With CurrencyFormats(2)
        .Name = "US Dollar - Negative Red"
        .Symbol = "$"
        .Code = "USD"
        .FormatPositive = "$#,##0.00"
        .FormatNegative = "$#,##0.00_);[Red]($#,##0.00)"
        .FormatZero = "$#,##0.00"
    End With
    
    With CurrencyFormats(3)
        .Name = "US Dollar - Code Format"
        .Symbol = "$"
        .Code = "USD"
        .FormatPositive = "#,##0 ""USD"""
        .FormatNegative = "#,##0 ""USD""_);(#,##0 ""USD"")"
        .FormatZero = "#,##0 ""USD"""
    End With
    
    ' Euro formats
    With CurrencyFormats(4)
        .Name = "Euro - Simple"
        .Symbol = "€"
        .Code = "EUR"
        .FormatPositive = "€#,##0"
        .FormatNegative = "€#,##0_);(€#,##0)"
        .FormatZero = "€#,##0"
    End With
    
    With CurrencyFormats(5)
        .Name = "Euro - Decimals"
        .Symbol = "€"
        .Code = "EUR"
        .FormatPositive = "€#,##0.00"
        .FormatNegative = "€#,##0.00_);(€#,##0.00)"
        .FormatZero = "€#,##0.00"
    End With
    
    With CurrencyFormats(6)
        .Name = "Euro - Code Format"
        .Symbol = "€"
        .Code = "EUR"
        .FormatPositive = "#,##0 ""EUR"""
        .FormatNegative = "#,##0 ""EUR""_);(#,##0 ""EUR"")"
        .FormatZero = "#,##0 ""EUR"""
    End With
    
    ' British Pound formats
    With CurrencyFormats(7)
        .Name = "British Pound - Simple"
        .Symbol = "£"
        .Code = "GBP"
        .FormatPositive = "£#,##0"
        .FormatNegative = "£#,##0_);(£#,##0)"
        .FormatZero = "£#,##0"
    End With
    
    With CurrencyFormats(8)
        .Name = "British Pound - Decimals"
        .Symbol = "£"
        .Code = "GBP"
        .FormatPositive = "£#,##0.00"
        .FormatNegative = "£#,##0.00_);(£#,##0.00)"
        .FormatZero = "£#,##0.00"
    End With
    
    ' Japanese Yen formats
    With CurrencyFormats(9)
        .Name = "Japanese Yen"
        .Symbol = "¥"
        .Code = "JPY"
        .FormatPositive = "¥#,##0"
        .FormatNegative = "¥#,##0_);(¥#,##0)"
        .FormatZero = "¥#,##0"
    End With
    
    ' Chinese Yuan formats
    With CurrencyFormats(10)
        .Name = "Chinese Yuan"
        .Symbol = "¥"
        .Code = "CNY"
        .FormatPositive = "¥#,##0.00"
        .FormatNegative = "¥#,##0.00_);(¥#,##0.00)"
        .FormatZero = "¥#,##0.00"
    End With
    
    ' Canadian Dollar formats
    With CurrencyFormats(11)
        .Name = "Canadian Dollar"
        .Symbol = "C$"
        .Code = "CAD"
        .FormatPositive = "C$#,##0.00"
        .FormatNegative = "C$#,##0.00_);(C$#,##0.00)"
        .FormatZero = "C$#,##0.00"
    End With
    
    ' Australian Dollar formats
    With CurrencyFormats(12)
        .Name = "Australian Dollar"
        .Symbol = "A$"
        .Code = "AUD"
        .FormatPositive = "A$#,##0.00"
        .FormatNegative = "A$#,##0.00_);(A$#,##0.00)"
        .FormatZero = "A$#,##0.00"
    End With
    
    ' Swiss Franc formats
    With CurrencyFormats(13)
        .Name = "Swiss Franc"
        .Symbol = "CHF"
        .Code = "CHF"
        .FormatPositive = "#,##0.00 ""CHF"""
        .FormatNegative = "#,##0.00 ""CHF""_);(#,##0.00 ""CHF"")"
        .FormatZero = "#,##0.00 ""CHF"""
    End With
    
    ' Indian Rupee formats
    With CurrencyFormats(14)
        .Name = "Indian Rupee"
        .Symbol = "₹"
        .Code = "INR"
        .FormatPositive = "₹#,##0.00"
        .FormatNegative = "₹#,##0.00_);(₹#,##0.00)"
        .FormatZero = "₹#,##0.00"
    End With
    
    ' Korean Won formats
    With CurrencyFormats(15)
        .Name = "Korean Won"
        .Symbol = "₩"
        .Code = "KRW"
        .FormatPositive = "₩#,##0"
        .FormatNegative = "₩#,##0_);(₩#,##0)"
        .FormatZero = "₩#,##0"
    End With
    
    ' Brazilian Real formats
    With CurrencyFormats(16)
        .Name = "Brazilian Real"
        .Symbol = "R$"
        .Code = "BRL"
        .FormatPositive = "R$#,##0.00"
        .FormatNegative = "R$#,##0.00_);(R$#,##0.00)"
        .FormatZero = "R$#,##0.00"
    End With
    
    ' Russian Ruble formats
    With CurrencyFormats(17)
        .Name = "Russian Ruble"
        .Symbol = "₽"
        .Code = "RUB"
        .FormatPositive = "#,##0.00 ""₽"""
        .FormatNegative = "#,##0.00 ""₽""_);(#,##0.00 ""₽"")"
        .FormatZero = "#,##0.00 ""₽"""
    End With
    
    ' Mexican Peso formats
    With CurrencyFormats(18)
        .Name = "Mexican Peso"
        .Symbol = "$"
        .Code = "MXN"
        .FormatPositive = "$#,##0.00 ""MXN"""
        .FormatNegative = "$#,##0.00 ""MXN""_);($#,##0.00 ""MXN"")"
        .FormatZero = "$#,##0.00 ""MXN"""
    End With
    
    ' Generic Currency format
    With CurrencyFormats(19)
        .Name = "Generic Currency"
        .Symbol = "¤"
        .Code = "CUR"
        .FormatPositive = "¤#,##0.00"
        .FormatNegative = "¤#,##0.00_);(¤#,##0.00)"
        .FormatZero = "¤#,##0.00"
    End With
    
    bFormatsInitialized = True
    Debug.Print "Currency formats initialized. Total formats: " & (UBound(CurrencyFormats) + 1)
End Sub

' =========================================================================
' FORMAT CYCLING LOGIC
' =========================================================================

Private Function GetNextCurrencyFormat() As String
    ' Get the next currency format in the cycle
    
    ' Determine current format
    Dim currentFormat As String
    currentFormat = Selection.Cells(1, 1).NumberFormat
    
    ' Find current index
    Dim foundIndex As Integer
    foundIndex = FindCurrentFormatIndex(currentFormat)
    
    If foundIndex >= 0 Then
        ' Move to next format
        currentCurrencyIndex = (foundIndex + 1) Mod (UBound(CurrencyFormats) + 1)
    Else
        ' Start from beginning if current format not recognized
        currentCurrencyIndex = 0
    End If
    
    ' Return the format string
    GetNextCurrencyFormat = CurrencyFormats(currentCurrencyIndex).FormatPositive
    
    Debug.Print "Next currency format: " & CurrencyFormats(currentCurrencyIndex).Name & " (" & GetNextCurrencyFormat & ")"
End Function

Private Function FindCurrentFormatIndex(currentFormat As String) As Integer
    ' Find the index of the current format in our array
    
    Dim i As Integer
    
    For i = 0 To UBound(CurrencyFormats)
        If CurrencyFormats(i).FormatPositive = currentFormat Or _
           CurrencyFormats(i).FormatNegative = currentFormat Or _
           CurrencyFormats(i).FormatZero = currentFormat Then
            FindCurrentFormatIndex = i
            Exit Function
        End If
    Next i
    
    ' Not found
    FindCurrentFormatIndex = -1
End Function

Private Function GetCurrentCurrencyInfo() As String
    ' Get descriptive information about the current currency format
    
    If currentCurrencyIndex >= 0 And currentCurrencyIndex <= UBound(CurrencyFormats) Then
        GetCurrentCurrencyInfo = CurrencyFormats(currentCurrencyIndex).Name & " (" & CurrencyFormats(currentCurrencyIndex).Code & ")"
    Else
        GetCurrentCurrencyInfo = "Unknown Currency Format"
    End If
End Function

' =========================================================================
' ADVANCED CYCLING FUNCTIONS
' =========================================================================

Public Sub CycleCurrencyByRegion(Optional region As String = "")
    ' Cycle through currencies for a specific region
    ' region: "NA" (North America), "EU" (Europe), "ASIA", "ALL" (default)
    
    On Error GoTo ErrorHandler
    
    If Selection.Cells.Count = 0 Then Exit Sub
    
    If Not bFormatsInitialized Then Call InitializeCurrencyFormats
    
    Dim regionFormats() As Integer
    Call GetRegionFormats(region, regionFormats)
    
    If UBound(regionFormats) < 0 Then
        MsgBox "No currency formats found for region: " & region, vbInformation
        Exit Sub
    End If
    
    ' Cycle through region-specific formats
    Static regionIndex As Integer
    If regionIndex >= UBound(regionFormats) + 1 Then regionIndex = 0
    
    Dim formatIndex As Integer
    formatIndex = regionFormats(regionIndex)
    
    Selection.NumberFormat = CurrencyFormats(formatIndex).FormatPositive
    regionIndex = regionIndex + 1
    
    Application.StatusBar = "XLerate: Applied " & CurrencyFormats(formatIndex).Name & " (Region: " & region & ")"
    Call UseExistingClearStatusBar
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in CycleCurrencyByRegion: " & Err.Description
End Sub

Private Sub GetRegionFormats(region As String, ByRef regionFormats() As Integer)
    ' Get currency format indices for a specific region
    
    Dim tempArray(0 To 19) As Integer
    Dim count As Integer
    
    Select Case UCase(region)
        Case "NA", "NORTH AMERICA"
            tempArray(0) = 0: tempArray(1) = 1: tempArray(2) = 2: tempArray(3) = 3  ' USD
            tempArray(4) = 11  ' CAD
            tempArray(5) = 18  ' MXN
            count = 5
            
        Case "EU", "EUROPE"
            tempArray(0) = 4: tempArray(1) = 5: tempArray(2) = 6  ' EUR
            tempArray(3) = 7: tempArray(4) = 8  ' GBP
            tempArray(5) = 13  ' CHF
            tempArray(6) = 17  ' RUB
            count = 6
            
        Case "ASIA"
            tempArray(0) = 9   ' JPY
            tempArray(1) = 10  ' CNY
            tempArray(2) = 14  ' INR
            tempArray(3) = 15  ' KRW
            count = 3
            
        Case Else ' "ALL" or unspecified
            Dim i As Integer
            For i = 0 To UBound(CurrencyFormats)
                tempArray(i) = i
            Next i
            count = UBound(CurrencyFormats)
    End Select
    
    ' Resize and copy to output array
    ReDim regionFormats(0 To count)
    Dim j As Integer
    For j = 0 To count
        regionFormats(j) = tempArray(j)
    Next j
End Sub

' =========================================================================
' SETTINGS AND CUSTOMIZATION
' =========================================================================

Public Sub AddCustomCurrency(name As String, symbol As String, code As String, formatString As String)
    ' Add a custom currency format to the collection
    ' This could be enhanced to save to settings
    
    On Error GoTo ErrorHandler
    
    If Not bFormatsInitialized Then Call InitializeCurrencyFormats
    
    ' For now, just add to debug log
    ' Future enhancement: dynamically expand array and save to settings
    Debug.Print "Custom currency requested: " & name & " (" & code & ") - " & formatString
    
    MsgBox "Custom currency support will be available in a future version." & vbCrLf & _
           "Requested: " & name & " (" & code & ")", vbInformation, "Custom Currency"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddCustomCurrency: " & Err.Description
End Sub

Public Function GetAvailableCurrencies() As String
    ' Return a list of all available currencies
    
    If Not bFormatsInitialized Then Call InitializeCurrencyFormats
    
    Dim currencyList As String
    Dim i As Integer
    
    currencyList = "Available Currencies:" & vbCrLf
    
    For i = 0 To UBound(CurrencyFormats)
        currencyList = currencyList & "• " & CurrencyFormats(i).Name & " (" & CurrencyFormats(i).Code & ")" & vbCrLf
    Next i
    
    GetAvailableCurrencies = currencyList
End Function

' =========================================================================
' UTILITY FUNCTIONS - FIXED: Uses existing ModUtilityFunctions
' =========================================================================

Private Sub UseExistingClearStatusBar()
    ' Helper to use existing ClearStatusBar from ModUtilityFunctions
    ' FIXED: Avoids naming conflicts by using existing function
    On Error Resume Next
    Application.Run "ModUtilityFunctions.ClearStatusBar"
    If Err.Number <> 0 Then
        ' Fallback if ModUtilityFunctions.ClearStatusBar doesn't exist
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        Application.StatusBar = False
    End If
    On Error GoTo 0
End Sub

Public Function GetCurrencyCyclingVersion() As String
    ' Return module version for diagnostics
    GetCurrencyCyclingVersion = MODULE_VERSION
End Function

Public Sub ShowCurrencyHelp()
    ' Display help information about currency cycling
    
    Dim helpText As String
    helpText = "XLerate Currency Cycling v" & MODULE_VERSION & vbCrLf & vbCrLf
    helpText = helpText & "Shortcut: Ctrl+Alt+Shift+6" & vbCrLf & vbCrLf
    helpText = helpText & "Features:" & vbCrLf
    helpText = helpText & "• 20 currency formats including USD, EUR, GBP, JPY" & vbCrLf
    helpText = helpText & "• Multiple format variations (simple, decimals, negatives)" & vbCrLf
    helpText = helpText & "• Regional cycling support" & vbCrLf
    helpText = helpText & "• Custom currency support (coming soon)" & vbCrLf & vbCrLf
    helpText = helpText & "Usage:" & vbCrLf
    helpText = helpText & "1. Select cells to format" & vbCrLf
    helpText = helpText & "2. Press Ctrl+Alt+Shift+6 to cycle currencies" & vbCrLf
    helpText = helpText & "3. Repeat to cycle through all formats" & vbCrLf & vbCrLf
    helpText = helpText & GetAvailableCurrencies()
    
    MsgBox helpText, vbInformation, "XLerate Currency Cycling Help"
End Sub

Public Sub TestCurrencyCycling()
    ' Test function for development and debugging
    Debug.Print "=== CurrencyCycling Test Function ==="
    Debug.Print "Module Version: " & GetCurrencyCyclingVersion()
    Debug.Print "Formats Initialized: " & bFormatsInitialized
    
    If Not bFormatsInitialized Then Call InitializeCurrencyFormats
    Debug.Print "Total Currency Formats: " & (UBound(CurrencyFormats) + 1)
    Debug.Print "Current Selection: " & Selection.Address
    Debug.Print "Test completed - use CycleCurrency() for actual operation"
End Sub