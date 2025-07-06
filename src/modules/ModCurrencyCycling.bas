' ================================================================
' File: src/modules/ModCurrencyCycling.bas
' Version: 1.1.0
' Date: January 2025
'
' CHANGELOG:
' v1.1.0 - Enhanced currency cycling with comprehensive international support
'        - Added local and foreign currency cycling (Macabacus Ctrl+Alt+Shift+3/4)
'        - Added automatic locale detection and regional defaults
'        - Enhanced decimal precision management and scaling options
'        - Added support for cryptocurrency and commodity formats
'        - Cross-platform compatibility and cultural formatting
' v1.0.0 - Initial implementation of basic currency cycling
'
' DESCRIPTION:
' Advanced currency format cycling system aligned with Macabacus patterns
' Provides comprehensive local and foreign currency formatting options
' Includes intelligent decimal scaling and professional financial presentation
' ================================================================

Attribute VB_Name = "ModCurrencyCycling"
Option Explicit

' Currency cycling arrays for different regions
Private LocalCurrencyFormats As Collection
Private ForeignCurrencyFormats As Collection
Private CurrentLocalIndex As Long
Private CurrentForeignIndex As Long

' Currency format constants
Private Const DEFAULT_CURRENCY_SYMBOL As String = "$"
Private Const MAX_DECIMAL_PLACES As Long = 4
Private Const SCALING_THRESHOLD As Double = 1000

Public Sub InitializeCurrencyFormats()
    ' Initialize currency format collections based on system locale
    
    Set LocalCurrencyFormats = New Collection
    Set ForeignCurrencyFormats = New Collection
    
    ' Reset indices
    CurrentLocalIndex = 0
    CurrentForeignIndex = 0
    
    ' Detect system locale and populate appropriate formats
    Call PopulateLocalCurrencyFormats
    Call PopulateForeignCurrencyFormats
    
    Debug.Print "Currency formats initialized - Local: " & LocalCurrencyFormats.Count & _
                ", Foreign: " & ForeignCurrencyFormats.Count
End Sub

Private Sub PopulateLocalCurrencyFormats()
    ' Populate local currency formats based on system locale
    
    Dim systemCurrency As String
    systemCurrency = GetSystemCurrencySymbol()
    
    ' Clear existing formats
    Set LocalCurrencyFormats = New Collection
    
    ' Standard local currency formats (Macabacus-style progression)
    Select Case systemCurrency
        Case "$"  ' US Dollar
            LocalCurrencyFormats.Add "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
            LocalCurrencyFormats.Add "_($* #,##0.0_);_($* (#,##0.0);_($* ""-""??_);_(@_)"
            LocalCurrencyFormats.Add "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            LocalCurrencyFormats.Add "$#,##0"
            LocalCurrencyFormats.Add "$#,##0.0"
            LocalCurrencyFormats.Add "$#,##0.00"
            LocalCurrencyFormats.Add "$#,##0,,"" M"""
            LocalCurrencyFormats.Add "$#,##0.0,,"" M"""
            LocalCurrencyFormats.Add "$#,##0,,,"" B"""
            LocalCurrencyFormats.Add "$#,##0.0,,,"" B"""
            
        Case "€"  ' Euro
            LocalCurrencyFormats.Add "_-€* #,##0_-;-€* #,##0_-;_-€* ""-""_-;_-@_-"
            LocalCurrencyFormats.Add "_-€* #,##0.0_-;-€* #,##0.0_-;_-€* ""-""?_-;_-@_-"
            LocalCurrencyFormats.Add "_-€* #,##0.00_-;-€* #,##0.00_-;_-€* ""-""??_-;_-@_-"
            LocalCurrencyFormats.Add "€#,##0"
            LocalCurrencyFormats.Add "€#,##0.0"
            LocalCurrencyFormats.Add "€#,##0.00"
            LocalCurrencyFormats.Add "€#,##0,,"" M"""
            LocalCurrencyFormats.Add "€#,##0.0,,"" M"""
            LocalCurrencyFormats.Add "€#,##0,,,"" B"""
            LocalCurrencyFormats.Add "€#,##0.0,,,"" B"""
            
        Case "£"  ' British Pound
            LocalCurrencyFormats.Add "_-£* #,##0_-;-£* #,##0_-;_-£* ""-""_-;_-@_-"
            LocalCurrencyFormats.Add "_-£* #,##0.0_-;-£* #,##0.0_-;_-£* ""-""?_-;_-@_-"
            LocalCurrencyFormats.Add "_-£* #,##0.00_-;-£* #,##0.00_-;_-£* ""-""??_-;_-@_-"
            LocalCurrencyFormats.Add "£#,##0"
            LocalCurrencyFormats.Add "£#,##0.0"
            LocalCurrencyFormats.Add "£#,##0.00"
            LocalCurrencyFormats.Add "£#,##0,,"" M"""
            LocalCurrencyFormats.Add "£#,##0.0,,"" M"""
            LocalCurrencyFormats.Add "£#,##0,,,"" B"""
            LocalCurrencyFormats.Add "£#,##0.0,,,"" B"""
            
        Case "¥"  ' Japanese Yen / Chinese Yuan
            LocalCurrencyFormats.Add "_-¥* #,##0_-;-¥* #,##0_-;_-¥* ""-""_-;_-@_-"
            LocalCurrencyFormats.Add "¥#,##0"
            LocalCurrencyFormats.Add "¥#,##0,,"" M"""
            LocalCurrencyFormats.Add "¥#,##0,,,"" B"""
            LocalCurrencyFormats.Add "¥#,##0,,,,"" T"""
            
        Case Else  ' Default to USD format
            LocalCurrencyFormats.Add "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
            LocalCurrencyFormats.Add "_($* #,##0.0_);_($* (#,##0.0);_($* ""-""??_);_(@_)"
            LocalCurrencyFormats.Add "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            LocalCurrencyFormats.Add "$#,##0"
            LocalCurrencyFormats.Add "$#,##0.0"
            LocalCurrencyFormats.Add "$#,##0.00"
            LocalCurrencyFormats.Add "$#,##0,,"" M"""
            LocalCurrencyFormats.Add "$#,##0.0,,"" M"""
            LocalCurrencyFormats.Add "$#,##0,,,"" B"""
            LocalCurrencyFormats.Add "$#,##0.0,,,"" B"""
    End Select
End Sub

Private Sub PopulateForeignCurrencyFormats()
    ' Populate foreign currency formats (non-local currencies)
    
    Set ForeignCurrencyFormats = New Collection
    
    ' Major world currencies (Macabacus foreign currency pattern)
    
    ' US Dollar variants
    ForeignCurrencyFormats.Add "[$USD] #,##0"
    ForeignCurrencyFormats.Add "[$USD] #,##0.0"
    ForeignCurrencyFormats.Add "[$USD] #,##0.00"
    ForeignCurrencyFormats.Add "[$USD] #,##0,,"" M"""
    ForeignCurrencyFormats.Add "[$USD] #,##0,,,"" B"""
    
    ' Euro variants
    ForeignCurrencyFormats.Add "[$EUR] #,##0"
    ForeignCurrencyFormats.Add "[$EUR] #,##0.0"
    ForeignCurrencyFormats.Add "[$EUR] #,##0.00"
    ForeignCurrencyFormats.Add "[$EUR] #,##0,,"" M"""
    ForeignCurrencyFormats.Add "[$EUR] #,##0,,,"" B"""
    
    ' British Pound variants
    ForeignCurrencyFormats.Add "[$GBP] #,##0"
    ForeignCurrencyFormats.Add "[$GBP] #,##0.0"
    ForeignCurrencyFormats.Add "[$GBP] #,##0.00"
    ForeignCurrencyFormats.Add "[$GBP] #,##0,,"" M"""
    ForeignCurrencyFormats.Add "[$GBP] #,##0,,,"" B"""
    
    ' Japanese Yen variants
    ForeignCurrencyFormats.Add "[$JPY] #,##0"
    ForeignCurrencyFormats.Add "[$JPY] #,##0,,"" M"""
    ForeignCurrencyFormats.Add "[$JPY] #,##0,,,"" B"""
    
    ' Chinese Yuan variants
    ForeignCurrencyFormats.Add "[$CNY] #,##0"
    ForeignCurrencyFormats.Add "[$CNY] #,##0.0"
    ForeignCurrencyFormats.Add "[$CNY] #,##0.00"
    ForeignCurrencyFormats.Add "[$CNY] #,##0,,"" M"""
    ForeignCurrencyFormats.Add "[$CNY] #,##0,,,"" B"""
    
    ' Swiss Franc variants
    ForeignCurrencyFormats.Add "[$CHF] #,##0"
    ForeignCurrencyFormats.Add "[$CHF] #,##0.0"
    ForeignCurrencyFormats.Add "[$CHF] #,##0.00"
    ForeignCurrencyFormats.Add "[$CHF] #,##0,,"" M"""
    
    ' Canadian Dollar variants
    ForeignCurrencyFormats.Add "[$CAD] #,##0"
    ForeignCurrencyFormats.Add "[$CAD] #,##0.0"
    ForeignCurrencyFormats.Add "[$CAD] #,##0.00"
    ForeignCurrencyFormats.Add "[$CAD] #,##0,,"" M"""
    
    ' Australian Dollar variants
    ForeignCurrencyFormats.Add "[$AUD] #,##0"
    ForeignCurrencyFormats.Add "[$AUD] #,##0.0"
    ForeignCurrencyFormats.Add "[$AUD] #,##0.00"
    ForeignCurrencyFormats.Add "[$AUD] #,##0,,"" M"""
    
    ' Cryptocurrency variants (modern addition)
    ForeignCurrencyFormats.Add "[$BTC] #,##0.0000"
    ForeignCurrencyFormats.Add "[$BTC] #,##0.00000000"
    ForeignCurrencyFormats.Add "[$ETH] #,##0.000"
    ForeignCurrencyFormats.Add "[$ETH] #,##0.0000"
End Sub

Public Sub CycleLocalCurrency(Optional control As IRibbonControl)
    ' Cycles through local currency formats (Macabacus Ctrl+Alt+Shift+3)
    
    On Error GoTo ErrorHandler