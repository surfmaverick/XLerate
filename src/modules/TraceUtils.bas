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
        
        ' Iterate through each cell in the selected range
        Dim cell As Range
        For Each cell In selectedRange
            ' Add the cell details to the list
            .lstPrecedents.AddItem cell.Worksheet.Name & "!" & cell.Address
        Next cell
    
        ' Collect precedents for the entire range
        Dim dependentsStr As String
        dependentsStr = findPrecedents(selectedRange)
    
        ' Populate the ListBox with precedents
        Dim precedentArr() As String
        precedentArr = Split(dependentsStr, Chr(13))
        
        Dim i As Long
        For i = LBound(precedentArr) To UBound(precedentArr)
            If precedentArr(i) <> "" Then
                Dim precedentAddress As String
                precedentAddress = precedentArr(i)
    
                .lstPrecedents.AddItem precedentAddress
            End If
        Next i

        .Show vbModeless
    End With
End Sub

Private Function findPrecedents(ByVal inRange As Range) As String
    Dim sheetIdx As Integer
    sheetIdx = Sheets(inRange.Parent.Name).Index

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
        
        ' Iterate through each cell in the selected range
        Dim cell As Range
        For Each cell In selectedRange
            ' Add the cell details to the list
            .lstDependents.AddItem cell.Worksheet.Name & "!" & cell.Address
        Next cell

        Dim dependentsStr As String
        dependentsStr = findDependents(selectedRange)

        ' Populate the ListBox
        Dim dependentArr() As String
        dependentArr = Split(dependentsStr, Chr(13))

        Dim i As Long
        For i = LBound(dependentArr) To UBound(dependentArr)
            If dependentArr(i) <> "" Then
                Dim dependentAddress As String
                dependentAddress = dependentArr(i)

                .lstDependents.AddItem dependentAddress
            End If
        Next i

        .Show vbModeless
    End With
End Sub

Private Function findDependents(ByVal inRange As Range) As String
    Dim sheetIdx As Integer
    sheetIdx = Sheets(inRange.Parent.Name).Index

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
