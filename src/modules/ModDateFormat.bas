Attribute VB_Name = "ModDateFormat"
' ModDateFormat
Option Explicit

Private FormatList() As clsFormatType

Public Function GetFormatList() As clsFormatType()
    Debug.Print "--- GetFormatList called ---"
    If Not IsArrayInitialized(FormatList) Then
        Debug.Print "FormatList not initialized - checking saved formats"
        If Not LoadFormatsFromWorkbook() Then
            Debug.Print "No saved formats found - initializing defaults"
            InitializeDateFormats
        End If
    End If
    GetFormatList = FormatList
End Function

Private Function IsArrayInitialized(ByRef arr As Variant) As Boolean
    On Error Resume Next
    IsArrayInitialized = (UBound(arr) >= 0)
    On Error GoTo 0
End Function

' In ModDateFormat.InitializeDateFormats
Public Sub InitializeDateFormats()
    On Error GoTo ErrorHandler
    Debug.Print "=== InitializeDateFormats START ==="
    
    ' First try to load saved formats
    If Not LoadFormatsFromWorkbook() Then
        Debug.Print "No saved formats found - creating defaults"
        ' Create default formats only if no saved formats exist
        Dim formatObj As clsFormatType
        ReDim FormatList(2)
        
        Set formatObj = New clsFormatType
        formatObj.Name = "Year Only"
        formatObj.FormatCode = "yyyy"
        Set FormatList(0) = formatObj
        Debug.Print "Created format 0: " & FormatList(0).Name
        
        Set formatObj = New clsFormatType
        formatObj.Name = "Month Year"
        formatObj.FormatCode = "mmm-yyyy"
        Set FormatList(1) = formatObj
        Debug.Print "Created format 1: " & FormatList(1).Name
        
        Set formatObj = New clsFormatType
        formatObj.Name = "Full Date"
        formatObj.FormatCode = "dd-mmm-yy"
        Set FormatList(2) = formatObj
        Debug.Print "Created format 2: " & FormatList(2).Name
        
        SaveFormatsToWorkbook
    End If
    
    Debug.Print "=== InitializeDateFormats END ==="
    Exit Sub

ErrorHandler:
    Debug.Print "Error in InitializeDateFormats: " & Err.Description
    Resume Next
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
    If index >= 0 And index <= UBound(FormatList) Then
        Set FormatList(index) = updatedFormat
        SaveFormatsToWorkbook
    End If
End Sub

Public Sub SaveFormatsToWorkbook()
   Debug.Print "=== SaveFormatsToWorkbook START ==="
   
   ' Wrap the delete in its own error handler
   On Error Resume Next
   ThisWorkbook.CustomDocumentProperties("SavedDateFormats").Delete
   If Err.Number <> 0 Then Debug.Print "Error deleting old property: " & Err.Description
   On Error GoTo ErrorHandler
   
   Dim propValue As String, i As Integer
   For i = LBound(FormatList) To UBound(FormatList)
       Debug.Print "Format " & i & ": " & FormatList(i).Name & " | " & FormatList(i).FormatCode
       propValue = propValue & FormatList(i).Name & "|" & FormatList(i).FormatCode & "||"
   Next i
   
   ThisWorkbook.CustomDocumentProperties.Add Name:="SavedDateFormats", _
       LinkToContent:=False, Type:=msoPropertyTypeString, value:=propValue
       
   Debug.Print "Property added successfully"
   ThisWorkbook.Save
   Debug.Print "Workbook saved successfully"
   Debug.Print "=== SaveFormatsToWorkbook END ==="
   Exit Sub

ErrorHandler:
   Debug.Print "Error in SaveFormatsToWorkbook: " & Err.Description
   MsgBox "Error saving formats: " & Err.Description, vbExclamation
   Resume Next
End Sub

Private Function LoadFormatsFromWorkbook() As Boolean
    Debug.Print "=== LoadFormatsFromWorkbook Debug ==="
    On Error Resume Next
    Dim propValue As String
    propValue = ThisWorkbook.CustomDocumentProperties("SavedDateFormats")
    Debug.Print "Loaded propValue: " & propValue
    If Err.Number <> 0 Then
        Debug.Print "Error loading property: " & Err.Description
        LoadFormatsFromWorkbook = False
        Exit Function
    End If
    On Error GoTo 0

    If propValue = "" Then
        Debug.Print "No saved formats found"
        LoadFormatsFromWorkbook = False
        Exit Function
    End If
    
    Dim formatsArray() As String, formatParts() As String
    formatsArray = Split(propValue, "||")
    Debug.Print "Found " & (UBound(formatsArray) - 1) & " format entries"
    
    ReDim FormatList(UBound(formatsArray) - 1)
    Dim i As Integer
    For i = LBound(formatsArray) To UBound(formatsArray) - 1
        If formatsArray(i) <> "" Then
            Debug.Print "Processing format " & i & ": " & formatsArray(i)
            formatParts = Split(formatsArray(i), "|")
            Set FormatList(i) = New clsFormatType
            FormatList(i).Name = formatParts(0)
            FormatList(i).FormatCode = formatParts(1)
            Debug.Print "Successfully loaded format: " & FormatList(i).Name
        End If
    Next i

    Debug.Print "=== LoadFormatsFromWorkbook Completed ==="
    LoadFormatsFromWorkbook = True
End Function

Public Sub CycleDateFormat()
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

