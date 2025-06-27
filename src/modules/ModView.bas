' =============================================================================
' File: ModView.bas
' Version: 2.0.0
' Description: View and display functions with Macabacus-style controls
' Author: XLerate Development Team
' Created: New module for Macabacus compatibility
' Last Modified: 2025-06-27
' =============================================================================

Attribute VB_Name = "ModView"
' View and Display Functions (Macabacus-style)
Option Explicit

' === ZOOM FUNCTIONS ===

Public Sub ZoomIn(Optional control As IRibbonControl)
    Debug.Print "ZoomIn called"
    
    On Error Resume Next
    Dim currentZoom As Integer
    currentZoom = ActiveWindow.Zoom
    
    ' Increase zoom in 10% increments, max 400%
    Dim newZoom As Integer
    newZoom = currentZoom + 10
    If newZoom > 400 Then newZoom = 400
    
    ActiveWindow.Zoom = newZoom
    Debug.Print "Zoom changed from " & currentZoom & "% to " & newZoom & "%"
    On Error GoTo 0
End Sub

Public Sub ZoomOut(Optional control As IRibbonControl)
    Debug.Print "ZoomOut called"
    
    On Error Resume Next
    Dim currentZoom As Integer
    currentZoom = ActiveWindow.Zoom
    
    ' Decrease zoom in 10% increments, min 10%
    Dim newZoom As Integer
    newZoom = currentZoom - 10
    If newZoom < 10 Then newZoom = 10
    
    ActiveWindow.Zoom = newZoom
    Debug.Print "Zoom changed from " & currentZoom & "% to " & newZoom & "%"
    On Error GoTo 0
End Sub

Public Sub ZoomToSelection(Optional control As IRibbonControl)
    Debug.Print "ZoomToSelection called"
    
    If Selection Is Nothing Then Exit Sub
    
    On Error Resume Next
    ActiveWindow.Zoom = True  ' Zoom to fit selection
    Debug.Print "Zoomed to fit selection"
    On Error GoTo 0
End Sub

Public Sub ZoomToFit(Optional control As IRibbonControl)
    Debug.Print "ZoomToFit called"
    
    On Error Resume Next
    ' Zoom to fit the entire used range
    If Not ActiveSheet.UsedRange Is Nothing Then
        ActiveSheet.UsedRange.Select
        ActiveWindow.Zoom = True
        Debug.Print "Zoomed to fit used range"
    End If
    On Error GoTo 0
End Sub

Public Sub SetZoom100(Optional control As IRibbonControl)
    Debug.Print "SetZoom100 called"
    
    On Error Resume Next
    ActiveWindow.Zoom = 100
    Debug.Print "Zoom set to 100%"
    On Error GoTo 0
End Sub

' === GRIDLINES AND DISPLAY ===

Public Sub ToggleGridlines(Optional control As IRibbonControl)
    Debug.Print "ToggleGridlines called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveWindow.DisplayGridlines
    
    ActiveWindow.DisplayGridlines = Not currentState
    Debug.Print "Gridlines changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

Public Sub ToggleHeadings(Optional control As IRibbonControl)
    Debug.Print "ToggleHeadings called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveWindow.DisplayHeadings
    
    ActiveWindow.DisplayHeadings = Not currentState
    Debug.Print "Headings changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

Public Sub ToggleFormulas(Optional control As IRibbonControl)
    Debug.Print "ToggleFormulas called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveWindow.DisplayFormulas
    
    ActiveWindow.DisplayFormulas = Not currentState
    Debug.Print "Formula display changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

Public Sub ToggleZeros(Optional control As IRibbonControl)
    Debug.Print "ToggleZeros called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveWindow.DisplayZeros
    
    ActiveWindow.DisplayZeros = Not currentState
    Debug.Print "Zero display changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

' === PAGE BREAKS ===

Public Sub HidePageBreaks(Optional control As IRibbonControl)
    Debug.Print "HidePageBreaks called"
    
    On Error Resume Next
    ActiveSheet.DisplayPageBreaks = False
    Debug.Print "Page breaks hidden"
    On Error GoTo 0
End Sub

Public Sub ShowPageBreaks(Optional control As IRibbonControl)
    Debug.Print "ShowPageBreaks called"
    
    On Error Resume Next
    ActiveSheet.DisplayPageBreaks = True
    Debug.Print "Page breaks shown"
    On Error GoTo 0
End Sub

Public Sub TogglePageBreaks(Optional control As IRibbonControl)
    Debug.Print "TogglePageBreaks called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveSheet.DisplayPageBreaks
    
    ActiveSheet.DisplayPageBreaks = Not currentState
    Debug.Print "Page breaks changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

' === FREEZE PANES ===

Public Sub ToggleFreezePanes(Optional control As IRibbonControl)
    Debug.Print "ToggleFreezePanes called"
    
    On Error Resume Next
    If ActiveWindow.FreezePanes Then
        ' Unfreeze panes
        ActiveWindow.FreezePanes = False
        Debug.Print "Panes unfrozen"
    Else
        ' Freeze panes at current selection
        ActiveWindow.FreezePanes = True
        Debug.Print "Panes frozen at current selection"
    End If
    On Error GoTo 0
End Sub

Public Sub FreezeTopRow(Optional control As IRibbonControl)
    Debug.Print "FreezeTopRow called"
    
    On Error Resume Next
    ActiveWindow.FreezePanes = False  ' Clear existing freeze
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    Debug.Print "Top row frozen"
    On Error GoTo 0
End Sub

Public Sub FreezeFirstColumn(Optional control As IRibbonControl)
    Debug.Print "FreezeFirstColumn called"
    
    On Error Resume Next
    ActiveWindow.FreezePanes = False  ' Clear existing freeze
    Range("B1").Select
    ActiveWindow.FreezePanes = True
    Debug.Print "First column frozen"
    On Error GoTo 0
End Sub

' === SPLIT WINDOWS ===

Public Sub ToggleSplitWindow(Optional control As IRibbonControl)
    Debug.Print "ToggleSplitWindow called"
    
    On Error Resume Next
    If ActiveWindow.Split Then
        ' Remove split
        ActiveWindow.Split = False
        Debug.Print "Window split removed"
    Else
        ' Create split at current selection
        ActiveWindow.Split = True
        Debug.Print "Window split at current selection"
    End If
    On Error GoTo 0
End Sub

' === VIEW MODES ===

Public Sub CycleViewMode(Optional control As IRibbonControl)
    Debug.Print "CycleViewMode called"
    
    On Error Resume Next
    Select Case ActiveWindow.View
        Case xlNormalView
            ActiveWindow.View = xlPageBreakPreview
            Debug.Print "Changed to Page Break Preview"
        Case xlPageBreakPreview
            ActiveWindow.View = xlPageLayoutView
            Debug.Print "Changed to Page Layout View"
        Case xlPageLayoutView
            ActiveWindow.View = xlNormalView
            Debug.Print "Changed to Normal View"
        Case Else
            ActiveWindow.View = xlNormalView
            Debug.Print "Set to Normal View"
    End Select
    On Error GoTo 0
End Sub

Public Sub SetNormalView(Optional control As IRibbonControl)
    Debug.Print "SetNormalView called"
    
    On Error Resume Next
    ActiveWindow.View = xlNormalView
    Debug.Print "Set to Normal View"
    On Error GoTo 0
End Sub

Public Sub SetPageBreakPreview(Optional control As IRibbonControl)
    Debug.Print "SetPageBreakPreview called"
    
    On Error Resume Next
    ActiveWindow.View = xlPageBreakPreview
    Debug.Print "Set to Page Break Preview"
    On Error GoTo 0
End Sub

Public Sub SetPageLayoutView(Optional control As IRibbonControl)
    Debug.Print "SetPageLayoutView called"
    
    On Error Resume Next
    ActiveWindow.View = xlPageLayoutView
    Debug.Print "Set to Page Layout View"
    On Error GoTo 0
End Sub

' === NAVIGATION HELPERS ===

Public Sub GoToCell(Optional control As IRibbonControl)
    Debug.Print "GoToCell called"
    
    Application.Dialogs(xlDialogFormulaGoto).Show
End Sub

Public Sub GoToSpecial(Optional control As IRibbonControl)
    Debug.Print "GoToSpecial called"
    
    Application.Dialogs(xlDialogSelectSpecial).Show
End Sub

Public Sub FindAndReplace(Optional control As IRibbonControl)
    Debug.Print "FindAndReplace called"
    
    Application.Dialogs(xlDialogFindFile).Show
End Sub

' === WINDOW MANAGEMENT ===

Public Sub NewWindow(Optional control As IRibbonControl)
    Debug.Print "NewWindow called"
    
    On Error Resume Next
    ActiveWorkbook.NewWindow
    Debug.Print "New window created"
    On Error GoTo 0
End Sub

Public Sub ArrangeWindows(Optional control As IRibbonControl)
    Debug.Print "ArrangeWindows called"
    
    On Error Resume Next
    Windows.Arrange xlArrangeStyleTiled
    Debug.Print "Windows arranged in tiled view"
    On Error GoTo 0
End Sub

Public Sub CascadeWindows(Optional control As IRibbonControl)
    Debug.Print "CascadeWindows called"
    
    On Error Resume Next
    Windows.Arrange xlArrangeStyleCascade
    Debug.Print "Windows arranged in cascade view"
    On Error GoTo 0
End Sub

' === FULL SCREEN MODE ===

Public Sub ToggleFullScreen(Optional control As IRibbonControl)
    Debug.Print "ToggleFullScreen called"
    
    On Error Resume Next
    Application.DisplayFullScreen = Not Application.DisplayFullScreen
    Debug.Print "Full screen mode toggled"
    On Error GoTo 0
End Sub

' === HIDE/SHOW ELEMENTS ===

Public Sub ToggleFormulaBar(Optional control As IRibbonControl)
    Debug.Print "ToggleFormulaBar called"
    
    On Error Resume Next
    Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
    Debug.Print "Formula bar toggled"
    On Error GoTo 0
End Sub

Public Sub ToggleStatusBar(Optional control As IRibbonControl)
    Debug.Print "ToggleStatusBar called"
    
    On Error Resume Next
    Application.DisplayStatusBar = Not Application.DisplayStatusBar
    Debug.Print "Status bar toggled"
    On Error GoTo 0
End Sub

Public Sub ToggleScrollBars(Optional control As IRibbonControl)
    Debug.Print "ToggleScrollBars called"
    
    On Error Resume Next
    With ActiveWindow
        .DisplayHorizontalScrollBar = Not .DisplayHorizontalScrollBar
        .DisplayVerticalScrollBar = Not .DisplayVerticalScrollBar
        Debug.Print "Scroll bars toggled"
    End With
    On Error GoTo 0
End Sub

' === WORKSHEET TABS ===

Public Sub ToggleWorksheetTabs(Optional control As IRibbonControl)
    Debug.Print "ToggleWorksheetTabs called"
    
    On Error Resume Next
    ActiveWindow.DisplayWorkbookTabs = Not ActiveWindow.DisplayWorkbookTabs
    Debug.Print "Worksheet tabs toggled"
    On Error GoTo 0
End Sub

' === RULER AND GUIDES ===

Public Sub ToggleRuler(Optional control As IRibbonControl)
    Debug.Print "ToggleRuler called"
    
    On Error Resume Next
    ' Note: Ruler is only available in Page Layout view
    If ActiveWindow.View = xlPageLayoutView Then
        ActiveWindow.DisplayRuler = Not ActiveWindow.DisplayRuler
        Debug.Print "Ruler toggled"
    Else
        MsgBox "Ruler is only available in Page Layout view.", vbInformation
    End If
    On Error GoTo 0
End Sub

' === CUSTOM VIEW FUNCTIONS ===

Public Sub SaveCustomView(Optional control As IRibbonControl)
    Debug.Print "SaveCustomView called"
    
    Dim viewName As String
    viewName = InputBox("Enter a name for this custom view:", "Save Custom View")
    
    If viewName <> "" Then
        On Error Resume Next
        ActiveWorkbook.CustomViews.Add viewName
        If Err.Number = 0 Then
            Debug.Print "Custom view saved: " & viewName
            MsgBox "Custom view '" & viewName & "' saved successfully.", vbInformation
        Else
            Debug.Print "Error saving custom view: " & Err.Description
            MsgBox "Error saving custom view: " & Err.Description, vbExclamation
        End If
        On Error GoTo 0
    End If
End Sub

Public Sub ShowCustomViews(Optional control As IRibbonControl)
    Debug.Print "ShowCustomViews called"
    
    On Error Resume Next
    Application.Dialogs(xlDialogCustomViews).Show
    On Error GoTo 0
End Sub