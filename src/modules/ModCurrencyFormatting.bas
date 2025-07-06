' =============================================================================
' File: ModCurrencyFormatting.bas
' Version: 2.0.0
' Date: January 2025
' Author: XLerate Development Team
'
' CHANGELOG:
' v2.0.0 - Comprehensive currency formatting with Macabacus alignment
'        - Local and foreign currency cycling (USD, EUR, GBP, JPY, etc.)
'        - Professional financial formatting standards
'        - Multiple decimal precision options
'        - Thousands/millions scaling support
'        - Cross-platform compatibility (Windows & macOS)
' v1.0.0 - Basic currency formatting
' =============================================================================

Attribute VB_Name = "ModCurrencyFormatting"
Option Explicit

' Currency format arrays
Private LocalCurrencyFormats() As clsFormatType
Private ForeignCurrencyFormats() As clsFormatType
Private Initialized As Boolean

' === INITIALIZATION ===

Public Sub InitializeCurrencyFormats()
    ' Initialize all currency format arrays
    Debug.Print "InitializeCurrencyFormats called"
    
    If Initialized Then Exit Sub
    
    InitializeLocalCurrencyFormats
    InitializeForeignCurrencyFormats
    
    Initialized = True
    Debug.Print "Currency formats initialized"
End Sub

Private Sub InitializeLocalCurrencyFormats()
    ' Initialize local currency formats (USD-focused)
    Debug.Print "Initializing local currency formats (USD)"
    
    ReDim LocalCurrencyFormats(7)
    
    ' USD - No decimals
    Set LocalCurrencyFormats(0) = New clsFormatType
    LocalCurrencyFormats(0).Name = "USD - No Decimals"
    LocalCurrencyFormats(0).FormatCode = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
    
    ' USD - 2 decimals
    Set LocalCurrencyFormats(1) = New clsFormatType
    LocalCurrencyFormats(1).Name = "USD - 2 Decimals"
    LocalCurrencyFormats(1).FormatCode = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    ' USD - Thousands
    Set LocalCurrencyFormats(2) = New clsFormatType
    LocalCurrencyFormats(2).Name = "USD - Thousands"
    LocalCurrencyFormats(2).FormatCode = "_($* #,##0,_);_($* (#,##0,);_($* ""-""_);_(@_)"
    
    ' USD - Thousands with 1 decimal
    Set LocalCurrencyFormats(3) = New clsFormatType
    LocalCurrencyFormats(3).Name = "USD - Thousands (1 Dec)"
    LocalCurrencyFormats(3).FormatCode = "_($* #,##0.0,_);_($* (#,##0.0,);_($* ""-""_);_(@_)"
    
    ' USD - Millions
    Set LocalCurrencyFormats(4) = New clsFormatType
    LocalCurrencyFormats(4).Name = "USD - Millions"
    LocalCurrencyFormats(4).FormatCode = "_($* #,##0,,_);_($* (#,##0,,);_($* ""-""_);_(@_)"
    
    ' USD - Millions with 1 decimal
    Set LocalCurrencyFormats(5) = New clsFormatType
    LocalCurrencyFormats(5).Name = "USD - Millions (1 Dec)"
    LocalCurrencyFormats(5).FormatCode = "_($* #,##0.0,,_);_($* (#,##0.0,,);_($* ""-""_);_(@_)"
    
    ' USD - Billions
    Set LocalCurrencyFormats(6) = New clsFormatType
    LocalCurrencyFormats(6).Name = "USD - Billions"
    LocalCurrencyFormats(6).FormatCode = "_($* #,##0,,,_);_($* (#,##0,,,);_($* ""-""_);_(@_)"
    
    ' USD - Billions with 1 decimal
    Set LocalCurrencyFormats(7) = New clsFormatType
    LocalCurrencyFormats(7).Name = "USD - Billions (1 Dec)"
    LocalCurrencyFormats(7).FormatCode = "_($* #,##0.0,,,_);_($* (#,##0.0,,,);_($* ""-""_);_(@_)"
End Sub

Private Sub InitializeForeignCurrencyFormats()
    ' Initialize foreign currency formats
    Debug.Print "Initializing foreign currency formats"
    
    ReDim ForeignCurrencyFormats(11)
    
    ' EUR - No decimals
    Set ForeignCurrencyFormats(0) = New clsFormatType
    ForeignCurrencyFormats(0).Name = "EUR - No Decimals"
    ForeignCurrencyFormats(0).FormatCode = "_(€* #,##0_);_(€* (#,##0);_(€* ""-""_);_(@_)"
    
    ' EUR - 2 decimals
    Set ForeignCurrencyFormats(1) = New clsFormatType
    ForeignCurrencyFormats(1).Name = "EUR -