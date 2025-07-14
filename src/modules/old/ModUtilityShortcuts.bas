' =============================================================================
' File: ModUtilityShortcuts.bas
' Version: 2.0.0
' Date: January 2025
' Author: XLerate Development Team
'
' CHANGELOG:
' v2.0.0 - Comprehensive utility shortcuts aligned with Macabacus workflow
'        - Enhanced paste operations (values, formats, transpose)
'        - Quick save functions with versioning
'        - View management utilities
'        - Professional productivity enhancements
'        - Cross-platform compatibility (Windows & macOS)
' v1.0.0 - Basic utility functions
' =============================================================================

Attribute VB_Name = "ModUtilityShortcuts"
Option Explicit

' === PASTE OPERATIONS (Macabacus-aligned) ===

Public Sub PasteValuesOnly(Optional control As IRibbonControl)
    ' Paste only values without formulas or formatting
    ' Matches Macabacus Paste Values - Ctrl+Alt+Shift+V
    Debug.Print "PasteValuesOnly called"
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' Store original selection
    Dim originalSelection As Range
    Set originalSelection = Selection
    
    ' Paste values only
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Application.StatusBar = "Values pasted to " & originalSelection.Address
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Pasted values only to " & originalSelection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteValuesOnly: " & Err.Description
    Application.CutCopyMode = False
End Sub

Public Sub PasteFormatsOnly(Optional control As IRibbonControl)
    ' Paste only formatting without values or formulas
    Debug.Print "PasteFormatsOnly called"
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Selection.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    Application.StatusBar = "Formats pasted successfully"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Pasted formats only"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteFormatsOnly: " & Err.Description
    Application.CutCopyMode = False
End Sub

Public Sub PasteTranspose(Optional control As IRibbonControl)
    ' Paste with transpose operation
    ' Matches Macabacus Paste Transpose - Ctrl+Alt+Shift+T
    Debug.Print "PasteTranspose called"
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Application.StatusBar = "Data pasted with transpose"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Pasted with transpose"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteTranspose: " & Err.Description
    Application.CutCopyMode = False
End Sub

Public Sub PasteInsert(Optional control As IRibbonControl)
    ' Paste and insert cells (shift existing cells)
    ' Matches Macabacus Paste Insert functionality
    Debug.Print "PasteInsert called"
    
    On Error GoTo ErrorHandler
    
    If Application.CutCopyMode = False Then
        MsgBox "No data in clipboard to paste.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' Insert and paste
    Selection.Insert Shift:=xlShiftDown
    Selection.PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    
    Application.StatusBar = "Data pasted with insert"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Pasted with insert"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteInsert: " & Err.Description
    Application.CutCopyMode = False
End Sub

Public Sub PasteDuplicate(Optional control As IRibbonControl)
    ' Duplicate current selection by copying and pasting adjacent
    ' Matches Macabacus Paste Duplicate functionality
    Debug.Print "PasteDuplicate called"
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' Copy current selection
    Selection.Copy
    
    ' Determine where to paste (to the right)
    Dim pasteRange As Range
    Set pasteRange = Selection.Offset(0, Selection.Columns.Count)
    
    ' Paste
    pasteRange.PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    
    ' Select the pasted range
    pasteRange.Select
    
    Application.StatusBar = "Selection duplicated"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Selection duplicated"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PasteDuplicate: " & Err.Description
    Application.CutCopyMode = False
End Sub

' === QUICK SAVE FUNCTIONS (Macabacus-aligned) ===

Public Sub QuickSave(Optional control As IRibbonControl)
    ' Quick save current workbook
    ' Matches Macabacus Quick Save - Ctrl+Alt+Shift+S
    Debug.Print "QuickSave called"
    
    On Error GoTo ErrorHandler
    
    If ActiveWorkbook.Path = "" Then
        ' First save - prompt for location
        Application.Dialogs(xlDialogSaveAs).Show
    Else
        ' Regular save
        ActiveWorkbook.Save
    End If
    
    Application.StatusBar = "Workbook saved"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Workbook saved"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in QuickSave: " & Err.Description
    MsgBox "Error saving workbook: " & Err.Description, vbExclamation, "XLerate"
End Sub

Public Sub QuickSaveAs(Optional control As IRibbonControl)
    ' Quick Save As with suggested filename
    Debug.Print "QuickSaveAs called"
    
    On Error GoTo ErrorHandler
    
    Dim originalName As String
    Dim suggestedName As String
    
    originalName = ActiveWorkbook.Name
    
    ' Remove extension for suggested name
    If InStr(originalName, ".") > 0 Then
        suggestedName = Left(originalName, InStrRev(originalName, ".") - 1)
    Else
        suggestedName = originalName
    End If
    
    ' Add version suffix
    suggestedName = suggestedName & "_v" & Format(Now, "yyyymmdd")
    
    ' Show Save As dialog with suggested name
    Application.DisplayAlerts = False
    Application.Dialogs(xlDialogSaveAs).Show suggestedName
    Application.DisplayAlerts = True
    
    Application.StatusBar = "Workbook saved as new version"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Save As completed"
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Debug.Print "Error in QuickSaveAs: " & Err.Description
End Sub

Public Sub QuickSaveWithTimestamp(Optional control As IRibbonControl)
    ' Save workbook with timestamp appended
    ' Enhanced version for version control
    Debug.Print "QuickSaveWithTimestamp called"
    
    On Error GoTo ErrorHandler
    
    If ActiveWorkbook.Path = "" Then
        MsgBox "Please save the workbook first before using timestamped save.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Dim originalName As String
    Dim newName As String
    Dim timestamp As String
    Dim baseName As String
    Dim extension As String
    Dim dotPosition As Long
    
    originalName = ActiveWorkbook.Name
    dotPosition = InStrRev(originalName, ".")
    
    If dotPosition > 0 Then
        baseName = Left(originalName, dotPosition - 1)
        extension = Mid(originalName, dotPosition)
    Else
        baseName = originalName
        extension = ".xlsx"
    End If
    
    timestamp = Format(Now, "yyyymmdd_hhmmss")
    newName = baseName & "_" & timestamp & extension
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & Application.PathSeparator & newName
    Application.DisplayAlerts = True
    
    Application.StatusBar = "Saved as: " & newName
    Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
    
    Debug.Print "Saved workbook with timestamp: " & newName
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Debug.Print "Error in QuickSaveWithTimestamp: " & Err.Description
    MsgBox "Error saving with timestamp: " & Err.Description, vbExclamation, "XLerate"
End Sub

Public Sub QuickSaveAll(Optional control As IRibbonControl)
    ' Save all open workbooks
    ' Matches Macabacus Quick Save All functionality
    Debug.Print "QuickSaveAll called"
    
    On Error GoTo ErrorHandler
    
    Dim savedCount As Integer
    savedCount = 0
    
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Path <> "" Then  ' Only save already-saved workbooks
            wb.Save
            savedCount = savedCount + 1
        End If
    Next wb
    
    Application.StatusBar = savedCount & " workbooks saved"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Saved " & savedCount & " workbooks"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in QuickSaveAll: " & Err.Description
End Sub

' === TIMESTAMP AND DATE FUNCTIONS ===

Public Sub InsertTimestamp(Optional control As IRibbonControl)
    ' Insert current timestamp in active cell
    Debug.Print "InsertTimestamp called"
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell for the timestamp.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Selection.Value = Now
    Selection.NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    
    Application.StatusBar = "Timestamp inserted"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Inserted timestamp in " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertTimestamp: " & Err.Description
End Sub

Public Sub InsertDate(Optional control As IRibbonControl)
    ' Insert current date only (no time)
    Debug.Print "InsertDate called"
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select a single cell for the date.", vbInformation, "XLerate"
        Exit Sub
    End If
    
    Selection.Value = Date
    Selection.NumberFormat = "mm/dd/yyyy"
    
    Application.StatusBar = "Date inserted"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Inserted date in " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertDate: " & Err.Description
End Sub

' === NAVIGATION SHORTCUTS ===

Public Sub GoToLastCell(Optional control As IRibbonControl)
    ' Navigate to last used cell in worksheet
    Debug.Print "GoToLastCell called"
    
    On Error Resume Next
    ActiveSheet.UsedRange.Cells(ActiveSheet.UsedRange.Cells.Count).Select
    
    Application.StatusBar = "Navigated to last cell"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Navigated to last used cell"
    On Error GoTo 0
End Sub

Public Sub GoToFirstCell(Optional control As IRibbonControl)
    ' Navigate to first used cell in worksheet
    Debug.Print "GoToFirstCell called"
    
    On Error Resume Next
    ActiveSheet.UsedRange.Cells(1).Select
    
    Application.StatusBar = "Navigated to first cell"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Navigated to first used cell"
    On Error GoTo 0
End Sub

Public Sub SelectCurrentRegion(Optional control As IRibbonControl)
    ' Select current region around active cell
    Debug.Print "SelectCurrentRegion called"
    
    On Error Resume Next
    If Not Selection Is Nothing Then
        Selection.CurrentRegion.Select
        
        Application.StatusBar = "Current region selected"
        Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
        
        Debug.Print "Selected current region"
    End If
    On Error GoTo 0
End Sub

' === WORKSHEET MANAGEMENT ===

Public Sub InsertWorksheet(Optional control As IRibbonControl)
    ' Insert new worksheet with formatted name
    Debug.Print "InsertWorksheet called"
    
    On Error GoTo ErrorHandler
    
    Dim newSheet As Worksheet
    Set newSheet = ActiveWorkbook.Worksheets.Add
    
    ' Suggest a name based on existing sheets
    Dim sheetCount As Integer
    sheetCount = ActiveWorkbook.Worksheets.Count
    newSheet.Name = "Sheet" & sheetCount
    
    Application.StatusBar = "New worksheet added: " & newSheet.Name
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "New worksheet added: " & newSheet.Name
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertWorksheet: " & Err.Description
End Sub

Public Sub DeleteCurrentWorksheet(Optional control As IRibbonControl)
    ' Delete current worksheet with confirmation
    Debug.Print "DeleteCurrentWorksheet called"
    
    On Error GoTo ErrorHandler
    
    If ActiveWorkbook.Worksheets.Count = 1 Then
        MsgBox "Cannot delete the only worksheet in the workbook.", vbExclamation, "XLerate"
        Exit Sub
    End If
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete worksheet '" & ActiveSheet.Name & "'?", _
                     vbYesNo + vbQuestion, "XLerate - Delete Worksheet")
    
    If response = vbYes Then
        Dim sheetName As String
        sheetName = ActiveSheet.Name
        
        Application.DisplayAlerts = False
        ActiveSheet.Delete
        Application.DisplayAlerts = True
        
        Application.StatusBar = "Worksheet '" & sheetName & "' deleted"
        Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
        
        Debug.Print "Worksheet deleted: " & sheetName
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Debug.Print "Error in DeleteCurrentWorksheet: " & Err.Description
End Sub

' === CALCULATION CONTROLS ===

Public Sub ForceCalculation(Optional control As IRibbonControl)
    ' Force full calculation of all open workbooks
    Debug.Print "ForceCalculation called"
    
    On Error Resume Next
    Application.CalculateFull
    
    Application.StatusBar = "Full calculation completed"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Forced full calculation"
    On Error GoTo 0
End Sub

Public Sub ToggleCalculationMode(Optional control As IRibbonControl)
    ' Toggle between automatic and manual calculation
    Debug.Print "ToggleCalculationMode called"
    
    On Error Resume Next
    If Application.Calculation = xlCalculationAutomatic Then
        Application.Calculation = xlCalculationManual
        Application.StatusBar = "Calculation: Manual"
        Debug.Print "Changed to manual calculation"
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = "Calculation: Automatic"
        Debug.Print "Changed to automatic calculation"
    End If
    
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    On Error GoTo 0
End Sub

' === SCREEN AND DISPLAY UTILITIES ===

Public Sub RefreshScreen(Optional control As IRibbonControl)
    ' Refresh screen display
    Debug.Print "RefreshScreen called"
    
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    Application.StatusBar = "Screen refreshed"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Screen refreshed"
    On Error GoTo 0
End Sub

Public Sub ResetView(Optional control As IRibbonControl)
    ' Reset view to standard settings
    Debug.Print "ResetView called"
    
    On Error Resume Next
    With ActiveWindow
        .Zoom = 100
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayFormulas = False
        .DisplayZeros = True
        .View = xlNormalView
    End With
    
    Application.StatusBar = "View reset to defaults"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "View reset to defaults"
    On Error GoTo 0
End Sub