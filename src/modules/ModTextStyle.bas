Attribute VB_Name = "ModTestStyle"
' ModTextStyle.bas
Option Explicit

Private Const NAME_PREFIX As String = "TextStyle_"
Private TextStyles() As clsTextStyleType
Private Initialized As Boolean
Private CurrentStyleIndex As Integer

Public Sub InitializeTextStyles()
    If Initialized Then Exit Sub
    
    Debug.Print "Initializing text styles"
    
    ' Try to load existing styles
    If Not LoadTextStylesFromWorkbook() Then
        Debug.Print "Creating default styles"
        ' Create default styles if none exist
        CreateDefaultStyles
        SaveTextStylesToWorkbook
    End If
    
    Initialized = True
    Debug.Print "Text styles initialized"
End Sub

Private Sub CreateDefaultStyles()
    ReDim TextStyles(0 To 2)
    
    ' Heading style
    Set TextStyles(0) = New clsTextStyleType
    With TextStyles(0)
        .Name = "Heading"
        .FontName = "Calibri"
        .FontSize = 14
        .Bold = True
        .Italic = False
        .Underline = False
        .FontColor = RGB(0, 0, 0)  ' Black
        .BackColor = RGB(240, 240, 240)  ' Light Gray
        .BorderStyle = xlContinuous
        .BorderWeight = xlMedium
        .BorderTop = True
        .BorderBottom = True
        .BorderLeft = False
        .BorderRight = False
    End With
    
    ' Subheading style
    Set TextStyles(1) = New clsTextStyleType
    With TextStyles(1)
        .Name = "Subheading"
        .FontName = "Calibri"
        .FontSize = 12
        .Bold = True
        .Italic = False
        .Underline = False
        .FontColor = RGB(89, 89, 89)  ' Dark Gray
        .BackColor = RGB(245, 245, 245)  ' Very Light Gray
        .BorderStyle = xlContinuous
        .BorderWeight = xlThin
        .BorderTop = False
        .BorderBottom = True
        .BorderLeft = False
        .BorderRight = False
    End With
    
    ' Sum style
    Set TextStyles(2) = New clsTextStyleType
    With TextStyles(2)
        .Name = "Sum"
        .FontName = "Calibri"
        .FontSize = 11
        .Bold = True
        .Italic = False
        .Underline = True
        .FontColor = RGB(0, 0, 0)  ' Black
        .BackColor = RGB(255, 255, 255)  ' White
        .BorderStyle = xlDouble
        .BorderWeight = xlThick
        .BorderTop = True
        .BorderBottom = False
        .BorderLeft = False
        .BorderRight = False
    End With
End Sub

Public Function GetTextStyleList() As clsTextStyleType()
    GetTextStyleList = TextStyles
End Function

Private Function LoadTextStylesFromWorkbook() As Boolean
    On Error Resume Next
    Debug.Print "=== LoadTextStylesFromWorkbook START ==="
    
    Dim styleString As String
    Dim fullName As String
    fullName = NAME_PREFIX & "List"
    
    ' Check if the storage sheet exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TextStyleStorage")
    If ws Is Nothing Then
        Debug.Print "No storage sheet found"
        LoadTextStylesFromWorkbook = False
        Exit Function
    End If
    
    ' Get the stored string
    styleString = ws.Range("A1").value
    Debug.Print "Loaded style string: " & styleString
    
    If Len(styleString) = 0 Then
        Debug.Print "Empty style string"
        LoadTextStylesFromWorkbook = False
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' Parse styles
    Dim styleStrings() As String
    styleStrings = Split(styleString, ";")
    
    If UBound(styleStrings) < 0 Then
        Debug.Print "No styles found in string"
        LoadTextStylesFromWorkbook = False
        Exit Function
    End If
    
    ReDim TextStyles(0 To UBound(styleStrings))
    
    Dim i As Integer
    For i = LBound(styleStrings) To UBound(styleStrings)
        Dim parts() As String
        parts = Split(styleStrings(i), "|")
        
        If UBound(parts) < 13 Then
            Debug.Print "Invalid style data for index " & i
            LoadTextStylesFromWorkbook = False
            Exit Function
        End If
        
        Set TextStyles(i) = New clsTextStyleType
        
        ' Remove quotes from strings
        Dim styleName As String
        Dim FontName As String
        styleName = Replace(Replace(parts(0), """", ""), "'", "")
        FontName = Replace(Replace(parts(1), """", ""), "'", "")
        
        With TextStyles(i)
            .Name = styleName
            .FontName = FontName
            .FontSize = CLng(parts(2))
            .Bold = CBool(CLng(parts(3)))
            .Italic = CBool(CLng(parts(4)))
            .Underline = CBool(CLng(parts(5)))
            .FontColor = CLng(parts(6))
            .BackColor = CLng(parts(7))
            .BorderStyle = CLng(parts(8))
            .BorderWeight = CLng(parts(9))
            .BorderTop = CBool(CLng(parts(10)))
            .BorderBottom = CBool(CLng(parts(11)))
            .BorderLeft = CBool(CLng(parts(12)))
            .BorderRight = CBool(CLng(parts(13)))
        End With
        Debug.Print "Loaded style: " & TextStyles(i).Name
    Next i
    
    LoadTextStylesFromWorkbook = True
    Debug.Print "=== LoadTextStylesFromWorkbook END ==="
    Exit Function
    
ErrorHandler:
    Debug.Print "Error loading styles: " & Err.Description & " (Error " & Err.Number & ")"
    LoadTextStylesFromWorkbook = False
End Function

Public Sub SaveTextStylesToWorkbook()
    Debug.Print "=== SaveTextStylesToWorkbook START ==="
    
    ' Convert styles to a serialized string format
    Dim styleString As String
    styleString = SerializeStyles()
    Debug.Print "Serialized styles: " & styleString
    
    ' Get or create the storage sheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TextStyleStorage")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "TextStyleStorage"
        ws.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0
    
    ' Save the string to cell A1
    ws.Range("A1").value = styleString
    
    ' Save to workbook name
    Dim fullName As String
    fullName = NAME_PREFIX & "List"
    
    ' Delete existing name if it exists
    On Error Resume Next
    ThisWorkbook.Names(fullName).Delete
    On Error GoTo 0
    
    ' Add new name referencing the cell
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=fullName, RefersTo:="=TextStyleStorage!$A$1"
    
    If Err.Number <> 0 Then
        Debug.Print "Error saving styles: " & Err.Description & " (Error " & Err.Number & ")"
        MsgBox "Error saving styles: " & Err.Description, vbExclamation
    Else
        Debug.Print "Styles saved successfully"
        ThisWorkbook.Save  ' Ensure the workbook is saved with the new styles
    End If
    
    Debug.Print "=== SaveTextStylesToWorkbook END ==="
End Sub

Private Function SerializeStyles() As String
    ' Format: Name|FontName|FontSize|Bold|Italic|Underline|FontColor|BackColor|BorderStyle|BorderWeight|BorderTop|BorderBottom|BorderLeft|BorderRight;NextStyle...
    Dim result As String
    Dim i As Integer
    
    For i = LBound(TextStyles) To UBound(TextStyles)
        With TextStyles(i)
            result = result & """" & .Name & """|" & _
                     """" & .FontName & """|" & _
                     CStr(.FontSize) & "|" & _
                     CStr(Abs(.Bold)) & "|" & _
                     CStr(Abs(.Italic)) & "|" & _
                     CStr(Abs(.Underline)) & "|" & _
                     CStr(.FontColor) & "|" & _
                     CStr(.BackColor) & "|" & _
                     CStr(.BorderStyle) & "|" & _
                     CStr(.BorderWeight) & "|" & _
                     CStr(Abs(.BorderTop)) & "|" & _
                     CStr(Abs(.BorderBottom)) & "|" & _
                     CStr(Abs(.BorderLeft)) & "|" & _
                     CStr(Abs(.BorderRight))
            
            If i < UBound(TextStyles) Then
                result = result & ";"
            End If
        End With
    Next i
    
    SerializeStyles = result
End Function

Public Sub AddStyle(newStyle As clsTextStyleType)
    ReDim Preserve TextStyles(LBound(TextStyles) To UBound(TextStyles) + 1)
    Set TextStyles(UBound(TextStyles)) = newStyle
    SaveTextStylesToWorkbook
End Sub

Public Sub RemoveStyle(index As Integer)
    If index < LBound(TextStyles) Or index > UBound(TextStyles) Then Exit Sub
    
    ' Shift remaining styles up
    Dim i As Integer
    For i = index To UBound(TextStyles) - 1
        Set TextStyles(i) = TextStyles(i + 1)
    Next i
    
    ' Resize array
    ReDim Preserve TextStyles(LBound(TextStyles) To UBound(TextStyles) - 1)
    SaveTextStylesToWorkbook
End Sub

Public Sub UpdateStyle(index As Integer, updatedStyle As clsTextStyleType)
    If index < LBound(TextStyles) Or index > UBound(TextStyles) Then Exit Sub
    Set TextStyles(index) = updatedStyle
    SaveTextStylesToWorkbook
End Sub


Public Sub CycleTextStyle()
    If Not Initialized Then InitializeTextStyles
    
    ' Get selected range
    Dim rng As Range
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    
    If rng Is Nothing Then Exit Sub
    
    ' Increment style index
    CurrentStyleIndex = (CurrentStyleIndex + 1) Mod (UBound(TextStyles) + 1)
    
    ' Apply style
    ApplyStyleToRange rng, TextStyles(CurrentStyleIndex)
End Sub

Private Sub ApplyStyleToRange(rng As Range, style As clsTextStyleType)
    Debug.Print vbNewLine & "=== ApplyStyleToRange START ==="
    Debug.Print "Applying style with:"
    Debug.Print "  BorderStyle: " & style.BorderStyle & " (" & _
        Choose(Abs(style.BorderStyle), _
            "xlContinuous", _
            "Unknown", _
            "Unknown", _
            "xlDashDot", _
            "xlDashDotDot", _
            "Unknown", _
            "Unknown", _
            "Unknown", _
            "Unknown", _
            "Unknown", _
            "Unknown", _
            "Unknown", _
            "xlSlantDashDot", _
            "xlDouble(-4119)", _
            "xlDash(-4115)", _
            "xlDot(-4118)") & ")"
    Debug.Print "  Borders (T/B/L/R): " & style.BorderTop & "/" & _
                                         style.BorderBottom & "/" & _
                                         style.BorderLeft & "/" & _
                                         style.BorderRight
    
    With rng
        .Font.Name = style.FontName
        .Font.Size = style.FontSize
        .Font.Bold = style.Bold
        .Font.Italic = style.Italic
        .Font.Underline = style.Underline
        .Font.color = style.FontColor
        .Interior.color = style.BackColor
        
        ' Clear existing borders
        .Borders.lineStyle = xlNone
        
        ' Apply border style if not None (changed from > 0 to <> 0)
        If style.BorderStyle <> 0 Then
            Debug.Print "  Applying border style: " & style.BorderStyle
            
            Dim borderWeight As XlBorderWeight
            Select Case style.BorderStyle
                Case xlContinuous
                    borderWeight = xlMedium
                Case xlDouble
                    borderWeight = xlThick
                Case xlDash
                    borderWeight = xlThin
                Case xlDot
                    borderWeight = xlThin
                Case Else
                    borderWeight = xlThin
            End Select
            
            If style.BorderTop Then
                With .Borders(xlEdgeTop)
                    .lineStyle = style.BorderStyle
                    .Weight = borderWeight
                End With
            End If
            If style.BorderBottom Then
                With .Borders(xlEdgeBottom)
                    .lineStyle = style.BorderStyle
                    .Weight = borderWeight
                End With
            End If
            If style.BorderLeft Then
                With .Borders(xlEdgeLeft)
                    .lineStyle = style.BorderStyle
                    .Weight = borderWeight
                End With
            End If
            If style.BorderRight Then
                With .Borders(xlEdgeRight)
                    .lineStyle = style.BorderStyle
                    .Weight = borderWeight
                End With
            End If
        End If
    End With
    Debug.Print "=== ApplyStyleToRange END ==="
End Sub



