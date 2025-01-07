Attribute VB_Name = "ModNumberFormat"
' ModNumberFormat
Option Explicit

Private FormatList() As clsFormatType

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
    IsArrayInitialized = (UBound(arr) >= 0)  ' Check if array has any elements
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
        Dim formatObj As clsFormatType
        ReDim FormatList(2)
        
        Set formatObj = New clsFormatType
        formatObj.Name = "Comma 0 Dec Lg Align"
        formatObj.FormatCode = "_(* #,##0_);(* (#,##0);_(* ""-""_);_(@_)"
        Set FormatList(0) = formatObj
        
        Set formatObj = New clsFormatType
        formatObj.Name = "Comma 1 Dec Lg Align"
        formatObj.FormatCode = "_(* #,##0.0_);(* (#,##0.0);_(* ""-""_);_(@_)"
        Set FormatList(1) = formatObj
        
        Set formatObj = New clsFormatType
        formatObj.Name = "Comma 2 Dec Lg Align"
        formatObj.FormatCode = "_(* #,##0.00_);(* (#,##0.00);_(* ""-""_);_(@_)"
        Set FormatList(2) = formatObj
        
        SaveFormatsToWorkbook
    End If
    Debug.Print "Format initialization complete, count: " & UBound(FormatList) - LBound(FormatList) + 1
End Sub


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
    Dim propValue As String, i As Integer
    For i = LBound(FormatList) To UBound(FormatList)
        Debug.Print "Saving format: Name = " & FormatList(i).Name & ", FormatCode = " & FormatList(i).FormatCode
        propValue = propValue & FormatList(i).Name & "|" & FormatList(i).FormatCode & "||"
    Next i

    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties("SavedFormats").Delete
    On Error GoTo 0
    ThisWorkbook.CustomDocumentProperties.Add Name:="SavedFormats", _
        LinkToContent:=False, Type:=msoPropertyTypeString, value:=propValue
    ThisWorkbook.Save
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

    Debug.Print "Found saved formats string: " & Left(propValue, 50) & "..."  ' Print first 50 chars
    
    Dim formatsArray() As String, formatParts() As String
    formatsArray = Split(propValue, "||")
    ReDim FormatList(UBound(formatsArray) - 1)

    Dim i As Integer
    For i = LBound(formatsArray) To UBound(formatsArray) - 1
        If formatsArray(i) <> "" Then
            formatParts = Split(formatsArray(i), "|")
            Set FormatList(i) = New clsFormatType
            FormatList(i).Name = formatParts(0)
            FormatList(i).FormatCode = formatParts(1)
            Debug.Print "Loaded format [" & i & "]: " & FormatList(i).Name & " | " & FormatList(i).FormatCode
        End If
    Next i

    LoadFormatsFromWorkbook = True
    Debug.Print "=== LoadFormatsFromWorkbook END ==="
End Function

Public Sub CycleNumberFormat()
    If Selection Is Nothing Then Exit Sub
    
    Dim currentFormat As String, nextFormat As String
    Dim found As Boolean
    currentFormat = Selection.NumberFormat
    
    Dim i As Integer
    For i = LBound(FormatList) To UBound(FormatList)
        If FormatList(i).FormatCode = currentFormat Then
            If i < UBound(FormatList) Then
                nextFormat = FormatList(i + 1).FormatCode
            Else
                nextFormat = FormatList(LBound(FormatList)).FormatCode
            End If
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then nextFormat = FormatList(LBound(FormatList)).FormatCode
    Selection.NumberFormat = nextFormat
End Sub


