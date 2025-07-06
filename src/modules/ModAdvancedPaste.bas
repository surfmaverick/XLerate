' ModAdvancedPaste.bas
' Version: 1.0.0
' Date: 2025-01-04
' Author: XLerate Development Team
' 
' CHANGELOG:
' v1.0.0 - Initial implementation of advanced paste operations
'        - Smart paste operations with context awareness
'        - Multiple clipboard management (simulated)
'        - Enhanced paste special operations
'        - Data transformation during paste
'
' DESCRIPTION:
' Advanced paste operations for enhanced productivity in financial modeling
' Provides intelligent paste functions that adapt to content and context

Attribute VB_Name = "ModAdvancedPaste"
Option Explicit

' Clipboard data structure for multiple clipboard simulation
Private Type ClipboardData
    Content As Variant
    Format As String
    Source As String
    Timestamp As Date
End Type

' Multiple clipboard storage (simulated)
Private ClipboardHistory(1 To 10) As ClipboardData
Private ClipboardIndex As Integer

Public Sub SmartPaste(Optional control As IRibbonControl)
    ' Intelligent paste that adapts based on source and destination content
    ' Analyzes clipboard content and destination to choose optimal paste method
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Smart Paste Operation ==="
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' Analyze clipboard content
    Dim clipboardData As ClipboardData
    clipboardData = AnalyzeClipboardContent()
    
    ' Analyze destination
    Dim destinationInfo As String
    destinationInfo = AnalyzeDestination(Selection)
    
    ' Determine optimal paste method
    Dim pasteMethod As String
    pasteMethod = DetermineOptimalPaste(clipboardData, destinationInfo)
    
    ' Perform the paste operation
    Call ExecuteSmartPaste(pasteMethod, Selection)
    
    Debug.Print "Smart paste completed using method: " & pasteMethod
    
    ' Store in clipboard history
    Call AddToClipboardHistory(clipboardData)
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in SmartPaste: " & Err.Description
    MsgBox "Error during smart paste: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

Public Sub PasteTranspose(Optional control As IRibbonControl)
    ' Pastes data with automatic transposition
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, _
                           SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Debug.Print "Transpose paste completed"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteTranspose: " & Err.Description
    MsgBox "Error during transpose paste: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

Public Sub PasteAdd(Optional control As IRibbonControl)
    ' Pastes values by adding to existing values
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd
    Application.CutCopyMode = False
    
    Debug.Print "Add paste completed"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteAdd: " & Err.Description
    MsgBox "Error during add paste: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

Public Sub PasteMultiply(Optional control As IRibbonControl)
    ' Pastes values by multiplying with existing values
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply
    Application.CutCopyMode = False
    
    Debug.Print "Multiply paste completed"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteMultiply: " & Err.Description
    MsgBox "Error during multiply paste: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

Public Sub PasteFormulasAsValues(Optional control As IRibbonControl)
    ' Converts formulas to values during paste while preserving formatting
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' First paste everything
    Selection.PasteSpecial Paste:=xlPasteAll
    
    ' Then convert formulas to values
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            cell.Value = cell.Value
        End If
    Next cell
    
    Application.CutCopyMode = False
    
    Debug.Print "Formulas converted to values during paste"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteFormulasAsValues: " & Err.Description
    MsgBox "Error during formula-to-value paste: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

Public Sub PasteWithScaling(Optional control As IRibbonControl)
    ' Pastes values with optional scaling (multiply by factor)
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' Get scaling factor from user
    Dim scaleFactor As String
    scaleFactor = InputBox("Enter scaling factor (e.g., 1000 for thousands, 0.01 for percent):", _
                          "XLerate Advanced Paste - Scale Values", "1")
    
    If scaleFactor = "" Then Exit Sub
    
    Dim scaleValue As Double
    On Error Resume Next
    scaleValue = CDbl(scaleFactor)
    On Error GoTo ErrorHandler
    
    If scaleValue = 0 Then
        MsgBox "Invalid scaling factor. Must be a non-zero number.", vbExclamation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    ' Paste values first
    Selection.PasteSpecial Paste:=xlPasteValues
    
    ' Apply scaling
    If scaleValue <> 1 Then
        Dim cell As Range
        For Each cell In Selection
            If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                cell.Value = cell.Value * scaleValue
            End If
        Next cell
    End If
    
    Application.CutCopyMode = False
    
    Debug.Print "Scaled paste completed with factor: " & scaleValue
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteWithScaling: " & Err.Description
    MsgBox "Error during scaled paste: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

Public Sub PasteSkipBlanks(Optional control As IRibbonControl)
    ' Pastes data while skipping blank cells (doesn't overwrite existing data with blanks)
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Selection.PasteSpecial Paste:=xlPasteAll, SkipBlanks:=True
    Application.CutCopyMode = False
    
    Debug.Print "Skip blanks paste completed"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteSkipBlanks: " & Err.Description
    MsgBox "Error during skip blanks paste: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

Public Sub ShowClipboardHistory(Optional control As IRibbonControl)
    ' Shows simulated clipboard history for multiple paste operations
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Showing Clipboard History ==="
    
    Dim historyText As String
    historyText = "Clipboard History:" & vbNewLine & vbNewLine
    
    Dim i As Integer
    Dim itemCount As Integer
    itemCount = 0
    
    For i = 1 To 10
        If ClipboardHistory(i).Content <> Empty Then
            itemCount = itemCount + 1
            historyText = historyText & itemCount & ". " & _
                         Format(ClipboardHistory(i).Timestamp, "hh:mm:ss") & " - " & _
                         ClipboardHistory(i).Format & " from " & _
                         ClipboardHistory(i).Source & vbNewLine
        End If
    Next i
    
    If itemCount = 0 Then
        historyText = historyText & "No items in clipboard history."
    Else
        historyText = historyText & vbNewLine & "Note: This is a simulated clipboard history." & vbNewLine & _
                     "Real multiple clipboard functionality would require Windows API integration."
    End If
    
    MsgBox historyText, vbInformation, "XLerate Advanced Paste - Clipboard History"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ShowClipboardHistory: " & Err.Description
    MsgBox "Error showing clipboard history: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

Public Sub PasteNumbersOnly(Optional control As IRibbonControl)
    ' Pastes only numeric values, ignoring text
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' Get clipboard data
    Dim clipboardRange As Range
    Set clipboardRange = Application.Selection ' This would need to be the copied range
    
    ' Paste values first
    Selection.PasteSpecial Paste:=xlPasteValues
    
    ' Remove non-numeric values
    Dim cell As Range
    For Each cell In Selection
        If Not IsNumeric(cell.Value) Or IsEmpty(cell.Value) Then
            cell.Clear
        End If
    Next cell
    
    Application.CutCopyMode = False
    
    Debug.Print "Numbers-only paste completed"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteNumbersOnly: " & Err.Description
    MsgBox "Error during numbers-only paste: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

Public Sub PasteTextOnly(Optional control As IRibbonControl)
    ' Pastes only text values, ignoring numbers
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' Paste values first
    Selection.PasteSpecial Paste:=xlPasteValues
    
    ' Remove numeric values
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            cell.Clear
        End If
    Next cell
    
    Application.CutCopyMode = False
    
    Debug.Print "Text-only paste completed"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteTextOnly: " & Err.Description
    MsgBox "Error during text-only paste: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub

' === PRIVATE HELPER FUNCTIONS ===

Private Function AnalyzeClipboardContent() As ClipboardData
    ' Analyzes the current clipboard content to determine optimal paste method
    
    Dim result As ClipboardData
    
    ' This is a simplified analysis - in practice, you'd use Windows API
    ' to get detailed clipboard information
    result.Content = "Unknown"
    result.Format = "Excel Range"
    result.Source = "Excel"
    result.Timestamp = Now
    
    ' Basic analysis based on selection (would be more sophisticated)
    If TypeName(Selection) = "Range" Then
        If Selection.HasFormula Then
            result.Format = "Formulas"
        ElseIf IsNumeric(Selection.Value) Then
            result.Format = "Numbers"
        Else
            result.Format = "Text"
        End If
    End If
    
    AnalyzeClipboardContent = result
End Function

Private Function AnalyzeDestination(destination As Range) As String
    ' Analyzes the destination range to understand context
    
    Dim analysis As String
    analysis = "Unknown"
    
    If destination.Cells.Count = 1 Then
        analysis = "Single Cell"
    ElseIf destination.Rows.Count = 1 Then
        analysis = "Single Row"
    ElseIf destination.Columns.Count = 1 Then
        analysis = "Single Column"
    Else
        analysis = "Multiple Cells"
    End If
    
    ' Check if destination has existing data
    If Not IsEmpty(destination.Value) Then
        analysis = analysis & " (Has Data)"
    Else
        analysis = analysis & " (Empty)"
    End If
    
    ' Check if destination has formulas
    Dim cell As Range
    For Each cell In destination
        If cell.HasFormula Then
            analysis = analysis & " (Has Formulas)"
            Exit For
        End If
    Next cell
    
    AnalyzeDestination = analysis
End Function

Private Function DetermineOptimalPaste(clipData As ClipboardData, destInfo As String) As String
    ' Determines the optimal paste method based on content and destination
    
    Dim method As String
    method = "Standard"
    
    ' Smart logic for determining paste method
    If clipData.Format = "Formulas" And InStr(destInfo, "Has Data") > 0 Then
        method = "Values Only"
    ElseIf clipData.Format = "Numbers" And InStr(destInfo, "Has Formulas") > 0 Then
        method = "Values Only"
    ElseIf InStr(destInfo, "Single Row") > 0 And clipData.Format <> "Single Row" Then
        method = "Transpose"
    ElseIf InStr(destInfo, "Single Column") > 0 And clipData.Format <> "Single Column" Then
        method = "Transpose"
    End If
    
    DetermineOptimalPaste = method
End Function

Private Sub ExecuteSmartPaste(method As String, destination As Range)
    ' Executes the determined paste method
    
    Select Case method
        Case "Values Only"
            destination.PasteSpecial Paste:=xlPasteValues
        Case "Transpose"
            destination.PasteSpecial Paste:=xlPasteAll, Transpose:=True
        Case "Formulas Only"
            destination.PasteSpecial Paste:=xlPasteFormulas
        Case "Formats Only"
            destination.PasteSpecial Paste:=xlPasteFormats
        Case Else
            destination.PasteSpecial Paste:=xlPasteAll
    End Select
    
    Application.CutCopyMode = False
End Sub

Private Sub AddToClipboardHistory(clipData As ClipboardData)
    ' Adds item to simulated clipboard history
    
    ' Shift existing items down
    Dim i As Integer
    For i = 10 To 2 Step -1
        ClipboardHistory(i) = ClipboardHistory(i - 1)
    Next i
    
    ' Add new item at top
    ClipboardHistory(1) = clipData
    
    Debug.Print "Added item to clipboard history"
End Sub

Public Sub ClearClipboardHistory(Optional control As IRibbonControl)
    ' Clears the simulated clipboard history
    
    Dim i As Integer
    For i = 1 To 10
        ClipboardHistory(i).Content = Empty
        ClipboardHistory(i).Format = ""
        ClipboardHistory(i).Source = ""
        ClipboardHistory(i).Timestamp = 0
    Next i
    
    Debug.Print "Clipboard history cleared"
    MsgBox "Clipboard history cleared.", vbInformation, "XLerate Advanced Paste"
End Sub

Public Sub PasteAsLink(Optional control As IRibbonControl)
    ' Creates a link to the source data instead of copying values
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate Advanced Paste"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Selection.PasteSpecial Paste:=xlPasteFormulas, Link:=True
    Application.CutCopyMode = False
    
    Debug.Print "Paste as link completed"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteAsLink: " & Err.Description
    MsgBox "Error during paste as link: " & Err.Description, vbCritical, "XLerate Advanced Paste"
End Sub