Attribute VB_Name = "FormulaConsistency"
' FormulaConsistency.cls
Option Explicit

' Constants for cell patterns and colors
Private Const PATTERN_HORIZONTAL As Long = xlHorizontal     ' Horizontal lines
Private Const COLOR_CONSISTENT As Long = 14348800           ' Green
Private Const COLOR_INCONSISTENT As Long = 255              ' Red
Private Const ORIGINAL_FORMAT_SHEET As String = "OriginalFormat"  ' Name of hidden sheet to store formats
Private Const FORMATTING_FLAG_CELL As String = "Z1"         ' Cell to store formatting state

Public Sub CheckHorizontalConsistency()
    ' Check if formatting is already applied
    If IsFormattingApplied() Then
        RemoveFormatting
        Exit Sub
    End If
    
    ' Store original formatting before proceeding
    StoreOriginalFormatting
    
    ' Store formatting state
    ActiveSheet.Range(FORMATTING_FLAG_CELL).Value = "Formatted"
    
    Application.ScreenUpdating = False
    
    ' Get the used range of the active sheet
    Dim usedRng As Range
    Set usedRng = ActiveSheet.UsedRange
    
    ' First pass - check and store horizontal consistency information
    Dim horizontalConsistentFormulasR1C1 As New Collection
    Dim R As Long, c As Long
    
    ' Find horizontally consistent formula patterns
    For R = usedRng.Row To usedRng.Row + usedRng.Rows.Count - 1
        For c = usedRng.Column To usedRng.Column + usedRng.Columns.Count - 2
            If Cells(R, c).HasFormula And Cells(R, c + 1).HasFormula Then
                If Cells(R, c).FormulaR1C1 = Cells(R, c + 1).FormulaR1C1 Then
                    AddToCollection horizontalConsistentFormulasR1C1, Cells(R, c).FormulaR1C1
                End If
            End If
        Next c
    Next R
    
    ' Check each cell for consistency with neighbors
    Dim cell As Range
    For Each cell In usedRng
        If cell.HasFormula Then
            Dim isHorizConsistent As Boolean
            Dim checkHoriz As Boolean
            
            ' Initialize check
            checkHoriz = False
            
            ' Check horizontal consistency
            If cell.Column < usedRng.Columns.Count + usedRng.Column - 1 Then
                If cell.Offset(0, 1).HasFormula Then
                    checkHoriz = True
                    isHorizConsistent = (cell.FormulaR1C1 = cell.Offset(0, 1).FormulaR1C1)
                Else
                    ' Check if this is the last cell in a consistent sequence
                    isHorizConsistent = IsInCollection(horizontalConsistentFormulasR1C1, cell.FormulaR1C1)
                    checkHoriz = isHorizConsistent
                End If
            Else
                ' Last column - check if part of consistent sequence
                isHorizConsistent = IsInCollection(horizontalConsistentFormulasR1C1, cell.FormulaR1C1)
                checkHoriz = isHorizConsistent
            End If
            
            ' Apply formatting based on consistency
            With cell.Interior
                ' Clear existing pattern
                .Pattern = xlNone
                
                ' Apply pattern if we need to check horizontal consistency
                If checkHoriz Then
                    .Pattern = PATTERN_HORIZONTAL
                    .PatternColor = IIf(isHorizConsistent, COLOR_CONSISTENT, COLOR_INCONSISTENT)
                End If
            End With
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    ' Show summary
    Dim msg As String
    msg = "Horizontal Formula Consistency Check Complete:" & vbNewLine & vbNewLine
    msg = msg & "Cells marked with:" & vbNewLine
    msg = msg & "- Green horizontal lines: Consistent with right neighbor" & vbNewLine
    msg = msg & "- Red horizontal lines: Inconsistent with right neighbor"
    
    MsgBox msg, vbInformation
End Sub

Private Sub AddToCollection(col As Collection, item As String)
    On Error Resume Next
    col.Add item, item  ' Using item as key prevents duplicates
    On Error GoTo 0
End Sub

Private Function IsInCollection(col As Collection, item As String) As Boolean
    Dim v As Variant
    For Each v In col
        If v = item Then
            IsInCollection = True
            Exit Function
        End If
    Next v
    IsInCollection = False
End Function

Private Function IsFormattingApplied() As Boolean
    On Error Resume Next
    IsFormattingApplied = (ActiveSheet.Range(FORMATTING_FLAG_CELL).Value = "Formatted")
    On Error GoTo 0
End Function

Private Sub StoreOriginalFormatting()
    ' Create a hidden sheet to store original formatting if it doesn't exist
    On Error Resume Next
    With ThisWorkbook
        If Not SheetExists(ORIGINAL_FORMAT_SHEET) Then
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = ORIGINAL_FORMAT_SHEET
            .Sheets(ORIGINAL_FORMAT_SHEET).Visible = xlSheetVeryHidden
        End If
    End With
    On Error GoTo 0
    
    ' Store the original interior format of cells with formulas
    Dim cell As Range
    Dim row As Long: row = 1
    
    With ThisWorkbook.Sheets(ORIGINAL_FORMAT_SHEET)
        .Cells.Clear
        For Each cell In ActiveSheet.UsedRange
            If cell.HasFormula Then
                .Cells(row, 1).Value = cell.Address
                .Cells(row, 2).Value = cell.Interior.Pattern
                .Cells(row, 3).Value = cell.Interior.Color
                .Cells(row, 4).Value = cell.Interior.PatternColor
                row = row + 1
            End If
        Next cell
    End With
End Sub

Private Sub RemoveFormatting()
    ' Get the used range of the active sheet
    Dim usedRng As Range
    Set usedRng = ActiveSheet.UsedRange
    
    Application.ScreenUpdating = False
    
    ' Restore original formatting
    If SheetExists(ORIGINAL_FORMAT_SHEET) Then
        Dim formatSheet As Worksheet
        Set formatSheet = ThisWorkbook.Sheets(ORIGINAL_FORMAT_SHEET)
        
        Dim row As Long
        For row = 1 To formatSheet.Cells(formatSheet.Rows.Count, 1).End(xlUp).row
            Dim cellAddress As String
            cellAddress = formatSheet.Cells(row, 1).Value
            
            With ActiveSheet.Range(cellAddress).Interior
                .Pattern = formatSheet.Cells(row, 2).Value
                If .Pattern <> xlNone Then
                    .Color = formatSheet.Cells(row, 3).Value
                    .PatternColor = formatSheet.Cells(row, 4).Value
                End If
            End With
        Next row
        
        ' Make sheet visible before deleting
        Application.DisplayAlerts = False
        formatSheet.Visible = xlSheetVisible
        formatSheet.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Clear the formatting flag
    ActiveSheet.Range(FORMATTING_FLAG_CELL).Value = ""
    
    Application.ScreenUpdating = True
    
    MsgBox "Formula consistency formatting has been removed.", vbInformation
End Sub

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
