' =============================================================================
' File: ModNumberFormat.bas
' Version: 2.0.0
' Description: Enhanced number formatting with Macabacus-style multiple format types
' Author: XLerate Development Team
' Created: Enhanced for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

Attribute VB_Name = "ModNumberFormat"
' Enhanced ModNumberFormat with Macabacus-style multiple format types
Option Explicit

Private FormatList() As clsFormatType
Private LocalCurrencyFormats() As clsFormatType
Private ForeignCurrencyFormats() As clsFormatType
Private PercentFormats() As clsFormatType
Private MultipleFormats() As clsFormatType
Private BinaryFormats() As clsFormatType

Public Function GetFormatList() As clsFormatType()
    Debug.Print "--- GetFormatList called ---"
    If Not IsArrayInitialized(FormatList) Then
        Debug.Print "FormatList not initialized - initializing now"
        InitializeFormats
    End If
    
    ' Debug print the current state
    Debug.Print "FormatList contains " & (UBound(FormatList) - LBound(FormatList) + 1) & " items"
    Dim i As Integer
    For i = LBound(FormatList) To UBound(FormatList)
        Debug.Print "  Item " & i & ": " & FormatList(i).Name
    Next i
    
    GetFormatList = FormatList
End Function

Private Function IsArrayInitialized(ByRef arr As Variant) As Boolean
    On Error Resume Next
    IsArrayInitialized = (UBound(arr) >= 0)
    On Error GoTo 0
End Function

Public Sub InitializeFormats()
    Debug.Print "InitializeFormats called"
    
    ' Clear existing FormatList
    Erase FormatList
    
    If LoadFormatsFromWorkbook() Then
        Debug.Print "Successfully loaded formats from workbook"
    Else
        Debug.Print "Loading default formats"
        InitializeDefaultFormats
        InitializeLocalCurrencyFormats
        InitializeForeignCurrencyFormats
        InitializePercentFormats
        InitializeMultipleFormats
        InitializeBinaryFormats
        SaveFormatsToWorkbook
    End If
    Debug.Print "Format initialization complete, count: " & UBound(FormatList) - LBound(FormatList) + 1
End Sub

Private Sub InitializeDefaultFormats()
    Dim formatObj As clsFormatType
    ReDim FormatList(4)  ' 5 default general number formats
    
    Set formatObj = New clsFormatType
    formatObj.Name = "General"
    formatObj.FormatCode = "General"
    Set FormatList(0) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Comma 0 Dec"
    formatObj.FormatCode = "_(* #,##0_);(* (#,##0);_(* ""-""_);_(@_)"
    Set FormatList(1) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Comma 1 Dec"
    formatObj.FormatCode = "_(* #,##0.0_);(* (#,##0.0);_(* ""-""_);_(@_)"
    Set FormatList(2) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Comma 2 Dec"
    formatObj.FormatCode = "_(* #,##0.00_);(* (#,##0.00);_(* ""-""_);_(@_)"
    Set FormatList(3) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Thousands"
    formatObj.FormatCode = "_(* #,##0,_);(* (#,##0,);_(* ""-""_);_(@_)"
    Set FormatList(4) = formatObj
End Sub

Private Sub InitializeLocalCurrencyFormats()
    Dim formatObj As clsFormatType
    ReDim LocalCurrencyFormats(3)
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Currency 0 Dec"
    formatObj.FormatCode = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
    Set LocalCurrencyFormats(0) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Currency 2 Dec"
    formatObj.FormatCode = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Set LocalCurrencyFormats(1) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Currency Thousands"
    formatObj.FormatCode = "_($* #,##0,_);_($* (#,##0,);_($* ""-""_);_(@_)"
    Set LocalCurrencyFormats(2) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Currency Millions"
    formatObj.FormatCode = "_($* #,##0,,_);_($* (#,##0,,);_($* ""-""_);_(@_)"
    Set LocalCurrencyFormats(3) = formatObj
End Sub

Private Sub InitializeForeignCurrencyFormats()
    Dim formatObj As clsFormatType
    ReDim ForeignCurrencyFormats(3)
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Euro 0 Dec"
    formatObj.FormatCode = "_(€* #,##0_);_(€* (#,##0);_(€* ""-""_);_(@_)"
    Set ForeignCurrencyFormats(0) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Euro 2 Dec"
    formatObj.FormatCode = "_(€* #,##0.00_);_(€* (#,##0.00);_(€* ""-""??_);_(@_)"
    Set ForeignCurrencyFormats(1) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "GBP 0 Dec"
    formatObj.FormatCode = "_(£* #,##0_);_(£* (#,##0);_(£* ""-""_);_(@_)"
    Set ForeignCurrencyFormats(2) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "GBP 2 Dec"
    formatObj.FormatCode = "_(£* #,##0.00_);_(£* (#,##0.00);_(£* ""-""??_);_(@_)"
    Set ForeignCurrencyFormats(3) = formatObj
End Sub

Private Sub InitializePercentFormats()
    Dim formatObj As clsFormatType
    ReDim PercentFormats(3)
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Percent 0 Dec"
    formatObj.FormatCode = "0%"
    Set PercentFormats(0) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Percent 1 Dec"
    formatObj.FormatCode = "0.0%"
    Set PercentFormats(1) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Percent 2 Dec"
    formatObj.FormatCode = "0.00%"
    Set PercentFormats(2) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Percent Basis Points"
    formatObj.FormatCode = "0.00""bps"""
    Set PercentFormats(3) = formatObj
End Sub

Private Sub InitializeMultipleFormats()
    Dim formatObj As clsFormatType
    ReDim MultipleFormats(3)
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Multiple 1 Dec"
    formatObj.FormatCode = "0.0""x"""
    Set MultipleFormats(0) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Multiple 2 Dec"
    formatObj.FormatCode = "0.00""x"""
    Set MultipleFormats(1) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "EV/EBITDA"
    formatObj.FormatCode = "0.0""x"""
    Set MultipleFormats(2) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "P/E Ratio"
    formatObj.FormatCode = "0.0""x"""
    Set MultipleFormats(3) = formatObj
End Sub

Private Sub InitializeBinaryFormats()
    Dim formatObj As clsFormatType
    ReDim BinaryFormats(2)
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Yes/No"
    formatObj.FormatCode = "[>0]""Yes"";[=0]""No"";""N/A"""
    Set BinaryFormats(0) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "True/False"
    formatObj.FormatCode = "[>0]""True"";[=0]""False"";""N/A"""
    Set BinaryFormats(1) = formatObj
    
    Set formatObj = New clsFormatType
    formatObj.Name = "Pass/Fail"
    formatObj.FormatCode = "[>0]""Pass"";[=0]""Fail"";""N/A"""
    Set BinaryFormats(2) = formatObj
End Sub

' === MACABACUS-STYLE CYCLING FUNCTIONS ===

Public Sub CyclePercent()
    If Selection Is Nothing Then Exit Sub
    
    If Not IsArrayInitialized(PercentFormats) Then
        InitializeFormats
    End If
    
    CycleFormatArray PercentFormats
End Sub

Public Sub CycleMultiple()
    If Selection Is Nothing Then Exit Sub
    
    If Not IsArrayInitialized(MultipleFormats) Then
        InitializeFormats
    End If
    
    CycleFormatArray MultipleFormats
End Sub

Public Sub CycleBinary()
    If Selection Is Nothing Then Exit Sub
    
    If Not IsArrayInitialized(BinaryFormats) Then
        InitializeFormats
    End If
    
    CycleFormatArray BinaryFormats
End Sub

' Generic cycling function for any format array
Private Sub CycleFormatArray(formatArray() As clsFormatType)
    Dim currentFormat As String, nextFormat As String
    Dim found As Boolean
    
    ' Get the format of the first cell in the selection
    currentFormat = Selection.Cells(1).NumberFormat
    
    ' If the selection has multiple cells with different formats,
    ' use the first format in our list
    Dim cell As Range
    For Each cell In Selection
        If cell.NumberFormat <> currentFormat Then
            currentFormat = formatArray(0).FormatCode
            found = True
            Exit For
        End If
    Next cell
    
    If Not found Then
        ' Find the next format in the cycle
        Dim i As Integer
        For i = LBound(formatArray) To UBound(formatArray)
            If formatArray(i).FormatCode = currentFormat Then
                If i < UBound(formatArray) Then
                    nextFormat = formatArray(i + 1).FormatCode
                Else
                    nextFormat = formatArray(LBound(formatArray)).FormatCode
                End If
                found = True
                Exit For
            End If
        Next i
    End If
    
    ' If no match found or cells had different formats, use first format
    If Not found Then nextFormat = formatArray(LBound(formatArray)).FormatCode
    
    ' Apply the format to all selected cells
    Selection.NumberFormat = nextFormat
    
    Debug.Print "Applied new format: " & nextFormat & " to " & Selection.Cells.Count & " cells"
End Sub

' === DECIMAL MANIPULATION FUNCTIONS ===

Public Sub IncreaseDecimals()
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            Dim currentFormat As String
            currentFormat = cell.NumberFormat
            
            ' Add a decimal place by modifying the format
            If InStr(currentFormat, ".") > 0 Then
                ' Already has decimals, add one more
                If Right(currentFormat, 1) = "0" Then
                    cell.NumberFormat = currentFormat & "0"
                ElseIf InStr(currentFormat, "0.") > 0 Then
                    ' Find the last 0 after decimal and add another
                    Dim pos As Integer
                    pos = InStrRev(currentFormat, "0")
                    If pos > 0 Then
                        cell.NumberFormat = Left(currentFormat, pos) & "0" & Mid(currentFormat, pos + 1)
                    End If
                End If
            Else
                ' No decimals, add .0
                If currentFormat = "General" Then
                    cell.NumberFormat = "0.0"
                Else
                    ' Insert .0 before the last 0
                    Dim lastZero As Integer
                    lastZero = InStrRev(currentFormat, "0")
                    If lastZero > 0 Then
                        cell.NumberFormat = Left(currentFormat, lastZero - 1) & "0.0" & Mid(currentFormat, lastZero + 1)
                    End If
                End If
            End If
        End If
    Next cell
    On Error GoTo 0
End Sub

Public Sub DecreaseDecimals()
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            Dim currentFormat As String
            currentFormat = cell.NumberFormat
            
            ' Remove a decimal place by modifying the format
            If InStr(currentFormat, ".") > 0 Then
                ' Find the last 0 after the decimal point and remove it
                Dim decimalPos As Integer
                decimalPos = InStr(currentFormat, ".")
                
                If decimalPos > 0 Then
                    Dim afterDecimal As String
                    afterDecimal = Mid(currentFormat, decimalPos + 1)
                    
                    If Right(afterDecimal, 1) = "0" And Len(afterDecimal) > 1 Then
                        ' Remove the last 0
                        cell.NumberFormat = Left(currentFormat, Len(currentFormat) - 1)
                    ElseIf afterDecimal = "0" Then
                        ' Remove the entire .0 part
                        cell.NumberFormat = Left(currentFormat, decimalPos - 1) & Mid(currentFormat, decimalPos + 2)
                    End If
                End If
            End If
        End If
    Next cell
    On Error GoTo 0
End Sub

' === EXISTING FUNCTIONS (Updated) ===

Public Sub AddFormat(newFormat As clsFormatType)
    Debug.Print "Adding format: " & newFormat.Name
    Dim newIndex As Integer
    newIndex = UBound(FormatList) + 1
    ReDim Preserve FormatList(newIndex)
    Set FormatList(newIndex) = newFormat
    SaveFormatsToWorkbook
End Sub

Public Sub RemoveFormat(index As Integer)
    Debug.Print "Removing format at index: " & index
    Dim i As Integer
    For i = index To UBound(FormatList) - 1
        Set FormatList(i) = FormatList(i + 1)
    Next i
    ReDim Preserve FormatList(UBound(FormatList) - 1)
    SaveFormatsToWorkbook
End Sub

Public Sub UpdateFormat(index As Integer, updatedFormat As clsFormatType)
    Debug.Print "Updating format at index " & index & " to Name: " & updatedFormat.Name & ", FormatCode: " & updatedFormat.FormatCode
    If index >= 0 And index <= UBound(FormatList) Then
        Set FormatList(index) = updatedFormat
        SaveFormatsToWorkbook
    Else
        Debug.Print "UpdateFormat index out of bounds"
    End If
End Sub

Public Sub SaveFormatsToWorkbook()
    Debug.Print "Saving formats to workbook"
    
    ' Save main formats
    Dim propValue As String, i As Integer
    For i = LBound(FormatList) To UBound(FormatList)
        Debug.Print "Saving format: Name = " & FormatList(i).Name & ", FormatCode = " & FormatList(i).FormatCode
        propValue = propValue & FormatList(i).Name & "|" & FormatList(i).FormatCode & "||"
    Next i

    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties("SavedFormats").Delete
    On Error GoTo 0
    ThisWorkbook.CustomDocumentProperties.Add Name:="SavedFormats", _
        LinkToContent:=False, Type:=msoPropertyTypeString, Value:=propValue
    
    ' Save specialized format arrays
    SaveSpecializedFormats "LocalCurrency", LocalCurrencyFormats
    SaveSpecializedFormats "ForeignCurrency", ForeignCurrencyFormats
    SaveSpecializedFormats "Percent", PercentFormats
    SaveSpecializedFormats "Multiple", MultipleFormats
    SaveSpecializedFormats "Binary", BinaryFormats
    
    ThisWorkbook.Save
End Sub

Private Sub SaveSpecializedFormats(formatType As String, formatArray() As clsFormatType)
    If Not IsArrayInitialized(formatArray) Then Exit Sub
    
    Dim propValue As String, i As Integer
    For i = LBound(formatArray) To UBound(formatArray)
        propValue = propValue & formatArray(i).Name & "|" & formatArray(i).FormatCode & "||"
    Next i

    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties("Saved" & formatType & "Formats").Delete
    On Error GoTo 0
    ThisWorkbook.CustomDocumentProperties.Add Name:="Saved" & formatType & "Formats", _
        LinkToContent:=False, Type:=msoPropertyTypeString, Value:=propValue
End Sub

Private Function LoadFormatsFromWorkbook() As Boolean
    Debug.Print "=== LoadFormatsFromWorkbook START ==="
    On Error Resume Next
    Dim propValue As String
    propValue = ThisWorkbook.CustomDocumentProperties("SavedFormats")
    
    If Err.Number <> 0 Then
        Debug.Print "Error reading CustomDocumentProperties: " & Err.Description
        LoadFormatsFromWorkbook = False
        Exit Function
    End If
    On Error GoTo 0

    If propValue = "" Then
        Debug.Print "No saved formats found in workbook"
        LoadFormatsFromWorkbook = False
        Exit Function
    End If

    Debug.Print "Found saved formats string: " & Left(propValue, 50) & "..."
    
    ' Load main formats
    If LoadFormatArray(propValue, FormatList) Then
        ' Load specialized formats
        LoadSpecializedFormats "LocalCurrency", LocalCurrencyFormats
        LoadSpecializedFormats "ForeignCurrency", ForeignCurrencyFormats
        LoadSpecializedFormats "Percent", PercentFormats
        LoadSpecializedFormats "Multiple", MultipleFormats
        LoadSpecializedFormats "Binary", BinaryFormats
        
        LoadFormatsFromWorkbook = True
    Else
        LoadFormatsFromWorkbook = False
    End If
    
    Debug.Print "=== LoadFormatsFromWorkbook END ==="
End Function

Private Sub LoadSpecializedFormats(formatType As String, formatArray() As clsFormatType)
    On Error Resume Next
    Dim propValue As String
    propValue = ThisWorkbook.CustomDocumentProperties("Saved" & formatType & "Formats")
    
    If Err.Number = 0 And propValue <> "" Then
        LoadFormatArray propValue, formatArray
    End If
    On Error GoTo 0
End Sub

Private Function LoadFormatArray(propValue As String, formatArray() As clsFormatType) As Boolean
    Dim formatsArray() As String, formatParts() As String
    formatsArray = Split(propValue, "||")
    ReDim formatArray(UBound(formatsArray) - 1)

    Dim i As Integer
    For i = LBound(formatsArray) To UBound(formatsArray) - 1
        If formatsArray(i) <> "" Then
            formatParts = Split(formatsArray(i), "|")
            Set formatArray(i) = New clsFormatType
            formatArray(i).Name = formatParts(0)
            formatArray(i).FormatCode = formatParts(1)
            Debug.Print "Loaded format [" & i & "]: " & formatArray(i).Name & " | " & formatArray(i).FormatCode
        End If
    Next i

    LoadFormatArray = True
End FunctionycleNumberFormat()
    If Selection Is Nothing Then Exit Sub
    
    If Not IsArrayInitialized(FormatList) Then
        InitializeFormats
    End If
    
    If Not IsArrayInitialized(FormatList) Then
        Exit Sub
    End If
    
    CycleFormatArray FormatList
End Sub

Public Sub CycleLocalCurrency()
    If Selection Is Nothing Then Exit Sub
    
    If Not IsArrayInitialized(LocalCurrencyFormats) Then
        InitializeFormats
    End If
    
    CycleFormatArray LocalCurrencyFormats
End Sub

Public Sub CycleForeignCurrency()
    If Selection Is Nothing Then Exit Sub
    
    If Not IsArrayInitialized(ForeignCurrencyFormats) Then
        InitializeFormats
    End If
    
    CycleFormatArray ForeignCurrencyFormats
End Sub

Public Sub C