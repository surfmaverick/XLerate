Attribute VB_Name = "RibbonCallbacks"
Option Explicit

'Callback for customUI.onLoad
Public myRibbon As IRibbonUI

'Store ribbon reference
Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
End Sub

'Callback for Trace Precedents button
Public Sub FindAndDisplayPrecedents(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ShowTracePrecedents"
    On Error GoTo 0
End Sub

'Callback for Trace Dependents button
Public Sub FindAndDisplayDependents(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ShowTraceDependents"
    On Error GoTo 0
End Sub

'Callback for Horizontal Formula Consistency button
Public Sub OnCheckHorizontalConsistency(control As IRibbonControl)
    On Error Resume Next
    Application.Run "CheckHorizontalConsistency"
    On Error GoTo 0
End Sub

' Switch Sign callback
Public Sub SwitchCellSign(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModSwitchSign.SwitchCellSign", control
    On Error GoTo 0
End Sub

' Smart Fill Right callback
Public Sub SmartFillRight(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModSmartFillRight.SmartFillRight", control
    On Error GoTo 0
End Sub

' Error Wrap callback
Public Sub WrapWithError(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModErrorWrap.WrapWithError", control
    On Error GoTo 0
End Sub

' Format callbacks
Public Sub OnFormatMain(control As IRibbonControl)
    ' Main format button action - cycle through number formats
    DoCycleNumberFormat control
End Sub

Public Sub DoCycleNumberFormat(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModNumberFormat.CycleNumberFormat"
    On Error GoTo 0
End Sub

Public Sub DoCycleCellFormat(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModCellFormat.CycleCellFormat"
    On Error GoTo 0
End Sub

Public Sub DoCycleDateFormat(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModDateFormat.CycleDateFormat"
    On Error GoTo 0
End Sub

Public Sub DoCycleTextStyle(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModTextStyle.CycleTextStyle"
    On Error GoTo 0
End Sub

Public Sub ShowSettingsForm(control As IRibbonControl)
    Debug.Print "ShowSettingsForm callback was triggered"
    ShowSettings   ' Direct call instead of Application.Run
End Sub

Public Sub DoAutoColorCells(control As IRibbonControl)
    Debug.Print "DoAutoColorCells callback started"
    AutoColorModule.AutoColorCells control
    Debug.Print "DoAutoColorCells callback ended"
End Sub

' === NEW MACABACUS-STYLE CALLBACKS ===

' Number Format Cycles
Public Sub CycleLocalCurrency(Optional control As IRibbonControl)
    ' Implement local currency cycle
    Debug.Print "Cycle Local Currency called"
End Sub

Public Sub CycleForeignCurrency(Optional control As IRibbonControl)
    ' Implement foreign currency cycle  
    Debug.Print "Cycle Foreign Currency called"
End Sub

Public Sub CyclePercent(Optional control As IRibbonControl)
    ' Implement percent cycle
    Debug.Print "Cycle Percent called"
End Sub

Public Sub CycleMultiple(Optional control As IRibbonControl)
    ' Implement multiple cycle
    Debug.Print "Cycle Multiple called"
End Sub

Public Sub CycleBinary(Optional control As IRibbonControl)
    ' Implement binary cycle
    Debug.Print "Cycle Binary called"
End Sub

Public Sub IncreaseDecimals(Optional control As IRibbonControl)
    ' Increase decimal places
    On Error Resume Next
    Selection.NumberFormat = Selection.Cells(1).NumberFormat & "0"
    On Error GoTo 0
End Sub

Public Sub DecreaseDecimals(Optional control As IRibbonControl)
    ' Decrease decimal places
    On Error Resume Next
    Dim currentFormat As String
    currentFormat = Selection.Cells(1).NumberFormat
    If Right(currentFormat, 1) = "0" Then
        Selection.NumberFormat = Left(currentFormat, Len(currentFormat) - 1)
    End If
    On Error GoTo 0
End Sub

' Color Cycles
Public Sub CycleBlueBlack(Optional control As IRibbonControl)
    ' Toggle between blue and black font colors
    On Error Resume Next
    If Selection.Font.Color = RGB(0, 0, 255) Then  ' Blue
        Selection.Font.Color = RGB(0, 0, 0)        ' Black
    Else
        Selection.Font.Color = RGB(0, 0, 255)      ' Blue
    End If
    On Error GoTo 0
End Sub

Public Sub CycleFontColor(Optional control As IRibbonControl)
    ' Cycle through font colors
    Debug.Print "Cycle Font Color called"
End Sub

Public Sub CycleFillColor(Optional control As IRibbonControl)
    ' Cycle through fill colors
    Debug.Print "Cycle Fill Color called"
End Sub

Public Sub CycleBorderColor(Optional control As IRibbonControl)
    ' Cycle through border colors
    Debug.Print "Cycle Border Color called"
End Sub

Public Sub AutoColorSheet(Optional control As IRibbonControl)
    ' Auto color entire sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.UsedRange.Select
    DoAutoColorCells control
End Sub

Public Sub AutoColorWorkbook(Optional control As IRibbonControl)
    ' Auto color entire workbook
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ws.UsedRange.Select
        DoAutoColorCells control
    Next ws
End Sub

' Alignment Cycles
Public Sub CycleCenter(Optional control As IRibbonControl)
    ' Cycle through center alignments
    On Error Resume Next
    Select Case Selection.HorizontalAlignment
        Case xlLeft
            Selection.HorizontalAlignment = xlCenter
        Case xlCenter
            Selection.HorizontalAlignment = xlRight
        Case xlRight
            Selection.HorizontalAlignment = xlLeft
        Case Else
            Selection.HorizontalAlignment = xlCenter
    End Select
    On Error GoTo 0
End Sub

Public Sub CycleHorizontal(Optional control As IRibbonControl)
    ' Cycle through horizontal alignments
    CycleCenter control
End Sub

Public Sub CycleLeftIndent(Optional control As IRibbonControl)
    ' Cycle through left indent levels
    On Error Resume Next
    Selection.IndentLevel = (Selection.IndentLevel + 1) Mod 4
    On Error GoTo 0
End Sub

' Border Cycles
Public Sub CycleBottomBorder(Optional control As IRibbonControl)
    ' Cycle bottom border styles
    On Error Resume Next
    With Selection.Borders(xlEdgeBottom)
        Select Case .LineStyle
            Case xlNone
                .LineStyle = xlContinuous
                .Weight = xlThin
            Case xlContinuous
                .Weight = xlMedium
            Case xlMedium
                .Weight = xlThick
            Case Else
                .LineStyle = xlNone
        End Select
    End With
    On Error GoTo 0
End Sub

Public Sub CycleLeftBorder(Optional control As IRibbonControl)
    ' Cycle left border styles
    On Error Resume Next
    With Selection.Borders(xlEdgeLeft)
        Select Case .LineStyle
            Case xlNone
                .LineStyle = xlContinuous
                .Weight = xlThin
            Case xlContinuous
                .Weight = xlMedium
            Case xlMedium
                .Weight = xlThick
            Case Else
                .LineStyle = xlNone
        End Select
    End With
    On Error GoTo 0
End Sub

Public Sub CycleRightBorder(Optional control As IRibbonControl)
    ' Cycle right border styles
    On Error Resume Next
    With Selection.Borders(xlEdgeRight)
        Select Case .LineStyle
            Case xlNone
                .LineStyle = xlContinuous
                .Weight = xlThin
            Case xlContinuous
                .Weight = xlMedium
            Case xlMedium
                .Weight = xlThick
            Case Else
                .LineStyle = xlNone
        End Select
    End With
    On Error GoTo 0
End Sub

Public Sub CycleOutsideBorder(Optional control As IRibbonControl)
    ' Cycle outside border styles
    On Error Resume Next
    Dim edges As Variant
    edges = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
    
    Dim currentStyle As XlLineStyle
    currentStyle = Selection.Borders(xlEdgeTop).LineStyle
    
    Dim edge As Variant
    For Each edge In edges
        With Selection.Borders(edge)
            Select Case currentStyle
                Case xlNone
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                Case xlContinuous
                    .Weight = xlMedium
                Case xlMedium
                    .Weight = xlThick
                Case Else
                    .LineStyle = xlNone
            End Select
        End With
    Next edge
    On Error GoTo 0
End Sub

Public Sub RemoveBorders(Optional control As IRibbonControl)
    ' Remove all borders
    On Error Resume Next
    Selection.Borders.LineStyle = xlNone
    On Error GoTo 0
End Sub

' Font Functions
Public Sub IncreaseFontSize(Optional control As IRibbonControl)
    ' Increase font size
    On Error Resume Next
    Selection.Font.Size = Selection.Font.Size + 1
    On Error GoTo 0
End Sub

Public Sub DecreaseFontSize(Optional control As IRibbonControl)
    ' Decrease font size
    On Error Resume Next
    If Selection.Font.Size > 6 Then
        Selection.Font.Size = Selection.Font.Size - 1
    End If
    On Error GoTo 0
End Sub

Public Sub CycleUnderline(Optional control As IRibbonControl)
    ' Cycle underline styles
    On Error Resume Next
    Select Case Selection.Font.Underline
        Case xlUnderlineStyleNone
            Selection.Font.Underline = xlUnderlineStyleSingle
        Case xlUnderlineStyleSingle
            Selection.Font.Underline = xlUnderlineStyleDouble
        Case xlUnderlineStyleDouble
            Selection.Font.Underline = xlUnderlineStyleNone
    End Select
    On Error GoTo 0
End Sub

Public Sub ToggleWrapText(Optional control As IRibbonControl)
    ' Toggle wrap text
    On Error Resume Next
    Selection.WrapText = Not Selection.WrapText
    On Error GoTo 0
End Sub

' View Functions
Public Sub ZoomIn(Optional control As IRibbonControl)
    ' Zoom in
    On Error Resume Next
    ActiveWindow.Zoom = ActiveWindow.Zoom + 10
    On Error GoTo 0
End Sub

Public Sub ZoomOut(Optional control As IRibbonControl)
    ' Zoom out
    On Error Resume Next
    If ActiveWindow.Zoom > 10 Then
        ActiveWindow.Zoom = ActiveWindow.Zoom - 10
    End If
    On Error GoTo 0
End Sub

Public Sub ToggleGridlines(Optional control As IRibbonControl)
    ' Toggle gridlines
    On Error Resume Next
    ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
    On Error GoTo 0
End Sub

Public Sub HidePageBreaks(Optional control As IRibbonControl)
    ' Hide page breaks
    On Error Resume Next
    ActiveSheet.DisplayPageBreaks = False
    On Error GoTo 0
End Sub

' Quick Save Functions
Public Sub QuickSave(Optional control As IRibbonControl)
    ' Quick save
    On Error Resume Next
    ActiveWorkbook.Save
    On Error GoTo 0
End Sub

Public Sub QuickSaveAll(Optional control As IRibbonControl)
    ' Quick save all open workbooks
    On Error Resume Next
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        wb.Save
    Next wb
    On Error GoTo 0
End Sub