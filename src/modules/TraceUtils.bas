Attribute VB_Name = "TraceUtils"
Option Explicit

Private Function fullAddress(inCell As Range) As String
    fullAddress = inCell.Address(External:=True)
End Function

Public Sub ShowTracePrecedents()
    ' Ensure user has selected a range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range."
        Exit Sub
    End If

    Dim selectedRange As Range
    Set selectedRange = Selection

    ' Create the UserForm
    Dim frmPrecedents As New frmPrecedents
    With frmPrecedents
        ' Clear existing items
        .lstPrecedents.Clear
        
        ' Set up list box properties
        .lstPrecedents.ColumnCount = 3
        .lstPrecedents.ColumnWidths = .GetColumnWidths()
        
        ' Add headers using separate labels
        .AddHeaders
        
        ' Display formula of selected cell
        .Controls("txtFormula").Text = selectedRange.Formula
        
        ' Add indentation prefix for hierarchy
        Const INDENT_CHAR As String = "  "
        
        ' Iterate through each cell in the selected range
        Dim cell As Range
        For Each cell In selectedRange
            .lstPrecedents.AddItem
            With .lstPrecedents
                .List(.ListCount - 1, 0) = cell.Worksheet.Name & "!" & cell.Address
                .List(.ListCount - 1, 1) = GetCellValueAsString(cell)
                .List(.ListCount - 1, 2) = cell.Formula
            End With
        Next cell
    
        ' Collect precedents for the entire range
        Dim precedentsStr As String
        precedentsStr = findPrecedents(selectedRange)
    
        ' Populate the ListBox with precedents
        If precedentsStr <> "" Then
            Dim precedentArr() As String
            precedentArr = Split(precedentsStr, Chr(13))
            
            Dim i As Long
            For i = LBound(precedentArr) To UBound(precedentArr)
                If precedentArr(i) <> "" Then
                    Dim precedentAddress As String
                    precedentAddress = precedentArr(i)
                    
                    ' Get the actual range object for the precedent
                    Dim precedentRange As Range
                    On Error Resume Next
                    Set precedentRange = Range(precedentAddress)
                    If Not precedentRange Is Nothing Then
                        .lstPrecedents.AddItem
                        With .lstPrecedents
                            .List(.ListCount - 1, 0) = INDENT_CHAR & precedentAddress
                            .List(.ListCount - 1, 1) = GetCellValueAsString(precedentRange)
                            .List(.ListCount - 1, 2) = precedentRange.Formula
                        End With
                    Else
                        .lstPrecedents.AddItem
                        With .lstPrecedents
                            .List(.ListCount - 1, 0) = INDENT_CHAR & precedentAddress
                        End With
                    End If
                    On Error GoTo 0
                End If
            Next i
        End If

        ' Select the first row if available
        If .lstPrecedents.ListCount > 0 Then
            .lstPrecedents.ListIndex = 0
        End If

        .Show vbModeless
    End With
End Sub
Private Function GetCellValueAsString(cell As Range) As String
    On Error Resume Next
    If IsError(cell.value) Then
        GetCellValueAsString = "#ERROR"
    ElseIf IsEmpty(cell.value) Then
        GetCellValueAsString = ""
    Else
        GetCellValueAsString = CStr(cell.value)
    End If
    On Error GoTo 0
End Function


Private Function findPrecedents(ByVal inRange As Range) As String
    Dim sheetIdx As Integer
    sheetIdx = Sheets(inRange.Parent.Name).index

    Dim inAddresses As String, returnSelection As Range
    Dim i As Long, pCount As Long, qCount As Long
    Set returnSelection = Selection
    
    ' Collect addresses of all cells in the range
    Dim cell As Range
    For Each cell In inRange
        inAddresses = inAddresses & fullAddress(cell) & "|"
    Next cell
    inAddresses = Left(inAddresses, Len(inAddresses) - 1)
    
    Application.ScreenUpdating = False

    With inRange
        .ShowPrecedents
        .NavigateArrow True, 1

        Dim loopCount As Long
        loopCount = 0
        Do
            loopCount = loopCount + 1
            If loopCount > 1000 Then Exit Do  ' Prevent infinite loop
            
            pCount = pCount + 1
            .NavigateArrow True, pCount

            ' Check if current cell is not in the original range
            If InStr(inAddresses, fullAddress(Selection)) = 0 Then
                If ActiveSheet.Name <> returnSelection.Parent.Name Then
                    Do
                        qCount = qCount + 1
                        .NavigateArrow True, pCount, qCount
                        findPrecedents = findPrecedents & fullAddress(Selection) & Chr(13)

                        On Error Resume Next
                        .NavigateArrow True, pCount, qCount + 1
                    Loop Until Err.Number <> 0
                    .NavigateArrow True, pCount + 1
                Else
                    findPrecedents = findPrecedents & fullAddress(Selection) & Chr(13)
                    .NavigateArrow True, pCount + 1
                End If
            Else
                .NavigateArrow True, pCount + 1
            End If

            ' Exit condition
            Dim foundMatch As Boolean
            For Each cell In inRange
                If fullAddress(ActiveCell) = fullAddress(cell) Then
                    foundMatch = True
                    Exit For
                End If
            Next cell
        Loop Until foundMatch
        
        .Parent.ClearArrows
    End With

    With returnSelection
        .Parent.Activate
        .Select
    End With

    Sheets(sheetIdx).Activate
End Function

Public Sub ShowTraceDependents()
    ' Ensure user has selected a range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range."
        Exit Sub
    End If

    Dim selectedRange As Range
    Set selectedRange = Selection

    ' Create the UserForm
    Dim frmDependents As New frmDependents
    With frmDependents
        ' Clear existing items
        .lstDependents.Clear
        
        ' Set up list box properties
        .lstDependents.ColumnCount = 3
        .lstDependents.ColumnWidths = .GetColumnWidths()
        
        ' Add headers using separate labels
        .AddHeaders
        
        ' Display formula of selected cell
        .Controls("txtFormula").Text = selectedRange.Formula
        
        ' Add indentation prefix for hierarchy
        Const INDENT_CHAR As String = "  "
        
        ' Iterate through each cell in the selected range
        Dim cell As Range
        For Each cell In selectedRange
            .lstDependents.AddItem
            With .lstDependents
                .List(.ListCount - 1, 0) = cell.Worksheet.Name & "!" & cell.Address
                .List(.ListCount - 1, 1) = GetCellValueAsString(cell)
                .List(.ListCount - 1, 2) = cell.Formula
            End With
        Next cell
    
        ' Collect dependents for the entire range
        Dim dependentsStr As String
        dependentsStr = findDependents(selectedRange)
    
        ' Populate the ListBox with dependents
        If dependentsStr <> "" Then
            Dim dependentArr() As String
            dependentArr = Split(dependentsStr, Chr(13))
            
            Dim i As Long
            For i = LBound(dependentArr) To UBound(dependentArr)
                If dependentArr(i) <> "" Then
                    Dim dependentAddress As String
                    dependentAddress = dependentArr(i)
                    
                    ' Get the actual range object for the dependent
                    Dim dependentRange As Range
                    On Error Resume Next
                    Set dependentRange = Range(dependentAddress)
                    If Not dependentRange Is Nothing Then
                        .lstDependents.AddItem
                        With .lstDependents
                            .List(.ListCount - 1, 0) = INDENT_CHAR & dependentAddress
                            .List(.ListCount - 1, 1) = GetCellValueAsString(dependentRange)
                            .List(.ListCount - 1, 2) = dependentRange.Formula
                        End With
                    Else
                        .lstDependents.AddItem
                        With .lstDependents
                            .List(.ListCount - 1, 0) = INDENT_CHAR & dependentAddress
                        End With
                    End If
                    On Error GoTo 0
                End If
            Next i
        End If

        ' Select the first row if available
        If .lstDependents.ListCount > 0 Then
            .lstDependents.ListIndex = 0
        End If

        .Show vbModeless
    End With
End Sub

Private Function findDependents(ByVal inRange As Range) As String
    Dim sheetIdx As Integer
    sheetIdx = Sheets(inRange.Parent.Name).index

    Dim inAddresses As String, returnSelection As Range
    Dim i As Long, pCount As Long, qCount As Long
    Set returnSelection = Selection
    
    ' Collect addresses of all cells in the range
    Dim cell As Range
    For Each cell In inRange
        inAddresses = inAddresses & fullAddress(cell) & "|"
    Next cell
    inAddresses = Left(inAddresses, Len(inAddresses) - 1)

    Application.ScreenUpdating = False

    With inRange
        .ShowDependents
        .NavigateArrow False, 1

        Dim loopCount As Long
        loopCount = 0
        Do
            loopCount = loopCount + 1
            If loopCount > 1000 Then Exit Do  ' Prevent infinite loop
            
            pCount = pCount + 1
            .NavigateArrow False, pCount

            ' Check if current cell is not in the original range
            If InStr(inAddresses, fullAddress(Selection)) = 0 Then
                If ActiveSheet.Name <> returnSelection.Parent.Name Then
                    Do
                        qCount = qCount + 1
                        .NavigateArrow False, pCount, qCount
                        findDependents = findDependents & fullAddress(Selection) & Chr(13)

                        On Error Resume Next
                        .NavigateArrow False, pCount, qCount + 1
                    Loop Until Err.Number <> 0
                    .NavigateArrow False, pCount + 1
                Else
                    findDependents = findDependents & fullAddress(Selection) & Chr(13)
                    .NavigateArrow False, pCount + 1
                End If
            Else
                .NavigateArrow False, pCount + 1
            End If

            ' Exit condition
            Dim foundMatch As Boolean
            For Each cell In inRange
                If fullAddress(ActiveCell) = fullAddress(cell) Then
                    foundMatch = True
                    Exit For
                End If
            Next cell
        Loop Until foundMatch
        
        .Parent.ClearArrows
    End With

    With returnSelection
        .Parent.Activate
        .Select
    End With

    Sheets(sheetIdx).Activate
End Function
