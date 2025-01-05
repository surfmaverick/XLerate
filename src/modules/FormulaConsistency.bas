Attribute VB_Name = "FormulaConsistency"
' FormulaConsistency.cls
Option Explicit

' Constants for cell patterns and colors
Private Const PATTERN_HORIZONTAL As Long = xlHorizontal     ' Horizontal lines
Private Const COLOR_CONSISTENT As Long = 14348800           ' Green
Private Const COLOR_INCONSISTENT As Long = 255              ' Red

Public Sub CheckHorizontalConsistency()
    ' Get the used range of the active sheet
    Dim usedRng As Range
    Set usedRng = ActiveSheet.UsedRange
    
    Application.ScreenUpdating = False
    
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
