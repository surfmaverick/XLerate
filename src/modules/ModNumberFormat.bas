Attribute VB_Name = "ModNumberFormat"
Option Explicit

Private FormatList() As clsFormatType
Private Type UndoAction
    RangeAddress As String
    OldFormat As String
End Type

Private UndoStack() As UndoAction
Private UndoStackSize As Long

Public Function GetFormatList() As clsFormatType()
    Debug.Print "--- GetFormatList called ---"
    If Not IsArrayInitialized(FormatList) Then
        Debug.Print "FormatList not initialized - initializing now"
        InitializeFormats
    End If
    
    GetFormatList = FormatList
End Function

Private Sub InitializeUndoStack()
    ReDim UndoStack(0 To 99)  ' Support up to 100 undo actions
    UndoStackSize = 0
End Sub

Private Sub PushUndo(ByVal RangeAddress As String, ByVal OldFormat As String)
    If UndoStackSize = 0 Then
        InitializeUndoStack
    End If

    ' Shift everything down if we're at capacity
    If UndoStackSize = 100 Then
        Dim i As Long
        For i = 0 To 98
            UndoStack(i) = UndoStack(i + 1)
        Next i
        UndoStackSize = 99
    End If

    ' Store the new undo action
    With UndoStack(UndoStackSize)
        .RangeAddress = RangeAddress
        .OldFormat = OldFormat
    End With
    UndoStackSize = UndoStackSize + 1

    ' Set up the undo button with dynamic description
    Application.OnUndo "Undo Format Change for Range: " & RangeAddress, "UndoLastNumberFormat"
End Sub

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
    Debug.Print "Format initialization complete"
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

Public Sub CycleNumberFormat()
    If Selection Is Nothing Then Exit Sub
    
    ' Check if FormatList is initialized
    If Not IsArrayInitialized(FormatList) Then
        InitializeFormats
    End If
    
    ' Store current format and selection for undo
    Dim currentFormat As String
    currentFormat = Selection.NumberFormat
    
    ' Store the selection address
    Dim selAddress As String
    selAddress = Selection.Address
    
    ' Add to undo stack before making changes
    PushUndo selAddress, currentFormat
    
    ' Find and apply next format
    Dim nextFormat As String
    Dim found As Boolean
    
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
    
    ' Apply the new format
    Selection.NumberFormat = nextFormat
    
    Debug.Print "Applied new format: " & nextFormat
End Sub

Public Sub UndoLastNumberFormat()
    If UndoStackSize > 0 Then
        UndoStackSize = UndoStackSize - 1
        With UndoStack(UndoStackSize)
            Range(.RangeAddress).NumberFormat = .OldFormat
            Debug.Print "Reverted format for Range: " & .RangeAddress
        End With
        
        ' Set up undo for the next item if there is one
        If UndoStackSize > 0 Then
            Application.OnUndo "Undo Format Change for Range: " & UndoStack(UndoStackSize - 1).RangeAddress, "UndoLastNumberFormat"
        Else
            Application.OnUndo "No further undo actions available", ""
        End If
    Else
        MsgBox "No undo actions available.", vbExclamation
    End If
End Sub

Public Sub RevertAllFormatting()
    If UndoStackSize = 0 Then
        MsgBox "No changes to revert.", vbInformation
        Exit Sub
    End If

    While UndoStackSize > 0
        UndoLastNumberFormat
    Wend

    MsgBox "All formatting changes have been reverted.", vbInformation
End Sub

Private Function IsArrayInitialized(ByRef arr As Variant) As Boolean
    On Error Resume Next
    IsArrayInitialized = (UBound(arr) >= 0)
    On Error GoTo 0
End Function

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


