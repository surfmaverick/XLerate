Attribute VB_Name = "ModCellFormat"
' ModCellFormat.cls
Option Explicit

Private CellFormatList() As clsCellFormatType
Public Const FONT_BOLD As Long = 1
Public Const FONT_ITALIC As Long = 2
Public Const FONT_UNDERLINE As Long = 4
Public Const FONT_STRIKETHROUGH As Long = 8

Public Function GetCellFormatList() As clsCellFormatType()
' Returns the array of cell format types. Initializes the formats if they haven't been loaded yet.
' @return: Array of clsCellFormatType objects containing all available cell formats

    Debug.Print "GetCellFormatList called"
    If Not IsArrayInitialized(CellFormatList) Then
        Debug.Print "CellFormatList not initialized, calling InitializeCellFormats"
        InitializeCellFormats
    End If
    Debug.Print "Returning CellFormatList with " & UBound(CellFormatList) - LBound(CellFormatList) + 1 & " items"
    GetCellFormatList = CellFormatList
End Function

Private Function IsArrayInitialized(ByRef arr As Variant) As Boolean
' Checks if an array has been initialized and contains elements.
' @param arr: Array to check
' @return: Boolean indicating if array is initialized

    On Error Resume Next
    IsArrayInitialized = (UBound(arr) >= 0)
    On Error GoTo 0
End Function

Public Sub InitializeCellFormats()
' Sets up the cell format types, either loading from workbook or creating default formats.
' Creates three default formats if none exist: Default (white), Highlight (yellow), and Important (red)

     Debug.Print "=== InitializeCellFormats debug ==="
    Debug.Print "LoadCellFormatsFromWorkbook result: " & LoadCellFormatsFromWorkbook()
    If IsArrayInitialized(CellFormatList) Then
        Debug.Print "CellFormatList is initialized with " & UBound(CellFormatList) - LBound(CellFormatList) + 1 & " items"
    Else
        Debug.Print "CellFormatList is not initialized"
    End If
    
    On Error GoTo ErrorHandler
    
    Debug.Print "InitializeCellFormats started"

    If LoadCellFormatsFromWorkbook() Then
        Debug.Print "Formats loaded from workbook"
    Else
        Debug.Print "Loading default formats"
        
        Dim formatObj As clsCellFormatType
        ReDim CellFormatList(4)  ' Changed to 5 formats (0-4)
        
        ' Normal format
        Set formatObj = New clsCellFormatType
        With formatObj
            .Name = "Normal"
            .BackColor = RGB(255, 255, 255)  ' White background
            .BorderStyle = xlNone            ' No borders
            .BorderColor = RGB(0, 0, 0)      ' Black (unused)
            .FillPattern = xlSolid
            .FontStyle = 0                   ' Normal
            .FontColor = RGB(0, 0, 0)        ' Black text
        End With
        Set CellFormatList(0) = formatObj
        
        ' Inputs format
        Set formatObj = New clsCellFormatType
        With formatObj
            .Name = "Inputs"
            .BackColor = RGB(255, 255, 204)  ' Light yellow (Excel note color)
            .BorderStyle = xlContinuous      ' Thin border
            .BorderColor = RGB(128, 128, 128) ' Dark gray border
            .FillPattern = xlSolid
            .FontStyle = 0                   ' Normal
            .FontColor = RGB(0, 0, 255)      ' Blue text
        End With
        Set CellFormatList(1) = formatObj
        
        ' Good format
        Set formatObj = New clsCellFormatType
        With formatObj
            .Name = "Good"
            .BackColor = RGB(198, 239, 206)  ' Light green (Excel good color)
            .BorderStyle = xlContinuous      ' Thin border
            .BorderColor = RGB(128, 128, 128) ' Dark gray border
            .FillPattern = xlSolid
            .FontStyle = 0                   ' Normal
            .FontColor = RGB(0, 97, 0)       ' Dark green text
        End With
        Set CellFormatList(2) = formatObj
        
        ' Bad format
        Set formatObj = New clsCellFormatType
        With formatObj
            .Name = "Bad"
            .BackColor = RGB(255, 199, 206)  ' Light red (Excel bad color)
            .BorderStyle = xlContinuous      ' Thin border
            .BorderColor = RGB(128, 128, 128) ' Dark gray border
            .FillPattern = xlSolid
            .FontStyle = 0                   ' Normal
            .FontColor = RGB(156, 0, 6)      ' Dark red text
        End With
        Set CellFormatList(3) = formatObj

        ' Important format
        Set formatObj = New clsCellFormatType
        With formatObj
            .Name = "Important"
            .BackColor = RGB(255, 255, 0)    ' Yellow background
            .BorderStyle = xlNone            ' No borders
            .BorderColor = RGB(0, 0, 0)      ' Black (unused)
            .FillPattern = xlSolid
            .FontStyle = 0                   ' Normal
            .FontColor = RGB(0, 0, 0)        ' Black text
        End With
        Set CellFormatList(4) = formatObj

        SaveCellFormatsToWorkbook
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Error in InitializeCellFormats: " & Err.Description & " (Error " & Err.Number & ")"
End Sub

Private Function DoFormatsMatch(ByVal target As Range, ByVal format As clsCellFormatType) As Boolean
' Compares a range's formatting with a format type to check if they match.
' @param target: Range to check
' @param format: Format type to compare against
' @return: Boolean indicating if formats match

    Debug.Print "  Checking format match for: " & format.Name
    
    ' Check interior pattern/color first
    Debug.Print "    Checking Pattern - Target: " & target.Interior.Pattern & ", Format: " & format.FillPattern
    If target.Interior.Pattern <> format.FillPattern Then
        Debug.Print "    Pattern mismatch"
        DoFormatsMatch = False
        Exit Function
    End If
    
    Debug.Print "    Checking Color - Target: " & target.Interior.color & ", Format: " & format.BackColor
    If target.Interior.color <> format.BackColor Then
        Debug.Print "    Color mismatch"
        DoFormatsMatch = False
        Exit Function
    End If
    
    ' Check all border edges
    Debug.Print "    Checking Borders..."
    Dim edges As Variant
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    Dim edge As Variant
    
    For Each edge In edges
        With target.Borders(edge)
            If format.BorderStyle = xlNone Then
                If .lineStyle <> xlNone Then
                    Debug.Print "    Border style mismatch"
                    DoFormatsMatch = False
                    Exit Function
                End If
            Else
                If .lineStyle <> format.BorderStyle Or .color <> format.BorderColor Then
                    Debug.Print "    Border style/color mismatch"
                    DoFormatsMatch = False
                    Exit Function
                End If
            End If
        End With
    Next edge
    
    Debug.Print "  Format matches!"
    DoFormatsMatch = True
End Function

Public Function GetBorderStyleValue(styleName As String) As XlLineStyle
' Converts a border style name (as shown in UI) to its corresponding Excel
' line style constant value. Ensures proper data translation between UI and Excel.

    Select Case styleName
        Case "None"
            GetBorderStyleValue = xlNone
        Case "Thin"
            GetBorderStyleValue = xlContinuous
        Case "Medium"
            GetBorderStyleValue = xlMedium
        Case "Thick"
            GetBorderStyleValue = xlThick
        Case "Double"
            GetBorderStyleValue = xlDouble
        Case "Dashed"
            GetBorderStyleValue = xlDash
        Case "Dotted"
            GetBorderStyleValue = xlDot
        Case Else
            GetBorderStyleValue = xlContinuous
    End Select
End Function

Public Function GetBorderStyleName(styleValue As XlLineStyle) As String
' Converts an Excel line style constant to its corresponding display name for the UI.
' Provides human-readable border style names.

    Select Case styleValue
        Case xlNone
            GetBorderStyleName = "None"
        Case xlContinuous
            GetBorderStyleName = "Thin"
        Case xlMedium
            GetBorderStyleName = "Medium"
        Case xlThick
            GetBorderStyleName = "Thick"
        Case xlDouble
            GetBorderStyleName = "Double"
        Case xlDash
            GetBorderStyleName = "Dashed"
        Case xlDot
            GetBorderStyleName = "Dotted"
        Case Else
            GetBorderStyleName = "Thin"
    End Select
End Function

Public Function GetFillPatternName(patternValue As Long) As String
    Select Case patternValue
        Case xlNone
            GetFillPatternName = "None"
        Case xlSolid
            GetFillPatternName = "Solid"
        Case xlGray25
            GetFillPatternName = "25% Gray"
        Case xlGray50
            GetFillPatternName = "50% Gray"
        Case xlGray75
            GetFillPatternName = "75% Gray"
        Case xlHorizontal
            GetFillPatternName = "Horizontal"
        Case xlVertical
            GetFillPatternName = "Vertical"
        Case xlUpward
            GetFillPatternName = "Diagonal Up"
        Case xlDownward
            GetFillPatternName = "Diagonal Down"
        Case Else
            GetFillPatternName = "None"
    End Select
End Function

Public Function GetFillPatternValue(patternName As String) As XlPattern
    Select Case patternName
        Case "None"
            GetFillPatternValue = xlNone
        Case "Solid"
            GetFillPatternValue = xlSolid
        Case "25% Gray"
            GetFillPatternValue = xlGray25
        Case "50% Gray"
            GetFillPatternValue = xlGray50
        Case "75% Gray"
            GetFillPatternValue = xlGray75
        Case "Horizontal"
            GetFillPatternValue = xlHorizontal
        Case "Vertical"
            GetFillPatternValue = xlVertical
        Case "Diagonal Up"
            GetFillPatternValue = xlUpward
        Case "Diagonal Down"
            GetFillPatternValue = xlDownward
        Case Else
            GetFillPatternValue = xlNone
    End Select
End Function

Public Function GetFontStyleValue(styleName As String) As Long
' Helper function to fetch matching font style

    Dim styleValue As Long
    styleValue = 0
    
    Select Case styleName
        Case "Bold"
            styleValue = styleValue + FONT_BOLD
        Case "Italic"
            styleValue = styleValue + FONT_ITALIC
        Case "Underline"
            styleValue = styleValue + FONT_UNDERLINE
        Case "Strike Through"
            styleValue = styleValue + FONT_STRIKETHROUGH
    End Select
    
    GetFontStyleValue = styleValue
End Function

Public Function GetFontStyleName(styleValue As Long) As String
' Helper function to fetch matching font style name

    If styleValue = 0 Then
        GetFontStyleName = "Normal"
        Exit Function
    End If
    
    If (styleValue And FONT_BOLD) Then
        GetFontStyleName = "Bold"
        Exit Function
    End If
    If (styleValue And FONT_ITALIC) Then
        GetFontStyleName = "Italic"
        Exit Function
    End If
    If (styleValue And FONT_UNDERLINE) Then
        GetFontStyleName = "Underline"
        Exit Function
    End If
    If (styleValue And FONT_STRIKETHROUGH) Then
        GetFontStyleName = "Strike Through"
        Exit Function
    End If
    
    GetFontStyleName = "Normal"
End Function

Public Sub CycleCellFormat()
' Cycles through available cell formats for the selected range.
' When applied to a selection, changes to the next format in the sequence.

    Debug.Print "=== CycleCellFormat START ==="
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.MergeCells Then Exit Sub

    ' Check if CellFormatList is initialized
    If Not IsArrayInitialized(CellFormatList) Then
        InitializeCellFormats
    End If
    
    ' Get next format to apply
    Dim nextFormatIndex As Long
    Dim found As Boolean
    Dim i As Long  ' Added declaration
    
    ' First try to find a match
    For i = 0 To UBound(CellFormatList)
        If DoFormatsMatch(Selection, CellFormatList(i)) Then
            nextFormatIndex = IIf(i < UBound(CellFormatList), i + 1, 0)
            found = True
            Exit For
        End If
    Next i
    
    ' If no match found, apply first format
    If Not found Then nextFormatIndex = 0
    
    ' Apply the format
    Debug.Print "Applying format: " & CellFormatList(nextFormatIndex).Name
    With Selection
        .Interior.Pattern = CellFormatList(nextFormatIndex).FillPattern
        .Interior.color = CellFormatList(nextFormatIndex).BackColor
        .Font.color = CellFormatList(nextFormatIndex).FontColor
        
        .Font.Bold = CBool(CellFormatList(nextFormatIndex).FontStyle And FONT_BOLD)
        .Font.Italic = CBool(CellFormatList(nextFormatIndex).FontStyle And FONT_ITALIC)
        .Font.Underline = CBool(CellFormatList(nextFormatIndex).FontStyle And FONT_UNDERLINE)
        .Font.Strikethrough = CBool(CellFormatList(nextFormatIndex).FontStyle And FONT_STRIKETHROUGH)
    End With

    ' Apply borders
    ApplyBorders Selection, CellFormatList(nextFormatIndex).BorderStyle, CellFormatList(nextFormatIndex).BorderColor
    
    Debug.Print "Applied format: " & CellFormatList(nextFormatIndex).Name
    Debug.Print "=== CycleCellFormat END ==="
End Sub

Private Sub ApplyBorders(target As Range, lineStyle As Long, color As Long)
' Applies border formatting to a range with specified style and color.
' @param target: Range to format
' @param lineStyle: Border line style to apply
' @param color: Border color to apply

    Dim edges As Variant
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    Dim edge As Variant
    
    For Each edge In edges
        With target.Borders(edge)
            .lineStyle = lineStyle
            If lineStyle <> xlNone Then
                .color = color
            End If
        End With
    Next edge
    
    ' Only apply inside borders if the range has multiple cells
    If target.Cells.Count > 1 Then
        With target.Borders(xlInsideHorizontal)
            .lineStyle = lineStyle
            If lineStyle <> xlNone Then
                .color = color
            End If
        End With
        With target.Borders(xlInsideVertical)
            .lineStyle = lineStyle
            If lineStyle <> xlNone Then
                .color = color
            End If
        End With
    End If
End Sub

Public Sub SaveCellFormatsToWorkbook()
' Saves all cell formats to the workbook's custom properties.
' Formats are stored as concatenated strings with properties separated by delimiters.

    Debug.Print "=== SaveCellFormatsToWorkbook Debug ==="
    Dim propValue As String, i As Integer
    
    Debug.Print "CellFormatList bounds: " & LBound(CellFormatList) & " to " & UBound(CellFormatList)
    
    For i = LBound(CellFormatList) To UBound(CellFormatList)
        Debug.Print "Saving format " & i & ":"
        Debug.Print "  Name: " & CellFormatList(i).Name
        Debug.Print "  BackColor: " & CellFormatList(i).BackColor
        Debug.Print "  BorderStyle: " & CellFormatList(i).BorderStyle
        Debug.Print "  BorderColor: " & CellFormatList(i).BorderColor
        Debug.Print "  FillPattern: " & CellFormatList(i).FillPattern
        Debug.Print "  FontStyle: " & CellFormatList(i).FontStyle
        Debug.Print "  FontColor: " & CellFormatList(i).FontColor
        
        propValue = propValue & CellFormatList(i).Name & "|" & _
                   CellFormatList(i).BackColor & "|" & _
                   CellFormatList(i).BorderStyle & "|" & _
                   CellFormatList(i).BorderColor & "|" & _
                   CellFormatList(i).FillPattern & "|" & _
                   CellFormatList(i).FontStyle & "|" & _
                   CellFormatList(i).FontColor & "||"
    Next i

    Debug.Print "Final propValue: " & propValue
    
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties("SavedCellFormats").Delete
    If Err.Number <> 0 Then Debug.Print "Delete error: " & Err.Description
    On Error GoTo 0
    
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties.Add Name:="SavedCellFormats", _
        LinkToContent:=False, Type:=msoPropertyTypeString, value:=propValue
    If Err.Number <> 0 Then Debug.Print "Add error: " & Err.Description
    On Error GoTo 0
    
    ThisWorkbook.Save
    Debug.Print "=== SaveCellFormatsToWorkbook Completed ==="
End Sub

Private Function LoadCellFormatsFromWorkbook() As Boolean
' Attempts to load saved cell formats from workbook's custom properties.
' @return: Boolean indicating if formats were successfully loaded

    Debug.Print "=== LoadCellFormatsFromWorkbook Debug ==="
    
    On Error Resume Next
    Dim propValue As String
    propValue = ThisWorkbook.CustomDocumentProperties("SavedCellFormats")
    Dim errNum As Long
    errNum = Err.Number
    Dim errDesc As String
    errDesc = Err.Description
    On Error GoTo 0
    
    Debug.Print "Loaded propValue: " & propValue
    
    If propValue = "" Then
        Debug.Print "No saved formats found"
        LoadCellFormatsFromWorkbook = False
        Exit Function
    End If

    Dim formatsArray() As String
    formatsArray = Split(propValue, "||")
    Debug.Print "Found " & UBound(formatsArray) & " format entries"
    
    ' Check if we have valid format entries
    Dim validFormats As Integer
    validFormats = 0
    Dim i As Integer
    For i = 0 To UBound(formatsArray) - 1
        If Len(Trim(formatsArray(i))) > 0 Then validFormats = validFormats + 1
    Next i
    
    Debug.Print "Valid formats found: " & validFormats
    
    If validFormats = 0 Then
        LoadCellFormatsFromWorkbook = False
        Exit Function
    End If
    
    ReDim CellFormatList(validFormats - 1)
    
    Dim currentFormat As Integer
    currentFormat = 0
    
    For i = 0 To UBound(formatsArray) - 1
        If Len(Trim(formatsArray(i))) > 0 Then
            Debug.Print "Processing format " & i & ": " & formatsArray(i)
            Dim formatParts() As String
            formatParts = Split(formatsArray(i), "|")
            
            Set CellFormatList(currentFormat) = New clsCellFormatType
            With CellFormatList(currentFormat)
                .Name = formatParts(0)
                .BackColor = CLng(formatParts(1))
                .BorderStyle = CLng(formatParts(2))
                .BorderColor = CLng(formatParts(3))
                ' Add checks for FillPattern, FontStyle, FontColor if they exist
                If UBound(formatParts) >= 4 Then .FillPattern = CLng(formatParts(4))
                If UBound(formatParts) >= 5 Then .FontStyle = CLng(formatParts(5))
                If UBound(formatParts) >= 6 Then .FontColor = CLng(formatParts(6))
            End With
            Debug.Print "Successfully loaded format: " & CellFormatList(currentFormat).Name
            currentFormat = currentFormat + 1
        End If
    Next i
    
    LoadCellFormatsFromWorkbook = True
    Debug.Print "=== LoadCellFormatsFromWorkbook Completed ==="
End Function


Public Sub AddFormat(newFormat As clsCellFormatType)
' Adds a new format to the list of available formats.
' @param newFormat: New format type to add to the collection

    Dim newIndex As Integer
    newIndex = UBound(CellFormatList) + 1
    ReDim Preserve CellFormatList(newIndex)
    Set CellFormatList(newIndex) = newFormat
    SaveCellFormatsToWorkbook
End Sub

Public Sub RemoveFormat(index As Integer)
' Removes a format at the specified index from the list.
' @param index: Index of format to remove

    Dim i As Integer
    For i = index To UBound(CellFormatList) - 1
        Set CellFormatList(i) = CellFormatList(i + 1)
    Next i
    ReDim Preserve CellFormatList(UBound(CellFormatList) - 1)
    SaveCellFormatsToWorkbook
End Sub

Public Sub UpdateFormat(index As Integer, updatedFormat As clsCellFormatType)
' Updates an existing format at the specified index.
' @param index: Index of format to update
' @param updatedFormat: New format settings to apply

    If index >= 0 And index <= UBound(CellFormatList) Then
        Set CellFormatList(index) = updatedFormat
        SaveCellFormatsToWorkbook
    End If
End Sub
