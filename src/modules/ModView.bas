' =============================================================================
' File: ModView.bas
' Version: 2.0.0
' Date: January 2025
' Author: XLerate Development Team
'
' CHANGELOG:
' v2.0.0 - Enhanced view management with Macabacus-aligned shortcuts
'        - Comprehensive zoom controls with keyboard shortcuts
'        - Advanced display options (gridlines, headings, formulas)
'        - Window management and navigation utilities
'        - Cross-platform optimization (Windows & macOS)
'        - Professional view state management
' v1.0.0 - Basic view functionality
' =============================================================================

Attribute VB_Name = "ModView"
Option Explicit

' === ZOOM FUNCTIONS (Macabacus-aligned) ===

Public Sub ZoomIn(Optional control As IRibbonControl)
    ' Zoom in by 10% increments - Ctrl+Alt+Shift+=
    Debug.Print "ZoomIn called"
    
    On Error Resume Next
    Dim currentZoom As Integer
    currentZoom = ActiveWindow.Zoom
    
    ' Increase zoom in 10% increments, max 400%
    Dim newZoom As Integer
    newZoom = currentZoom + 10
    If newZoom > 400 Then newZoom = 400
    
    ActiveWindow.Zoom = newZoom
    Application.StatusBar = "Zoom: " & newZoom & "%"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Zoom changed from " & currentZoom & "% to " & newZoom & "%"
    On Error GoTo 0
End Sub

Public Sub ZoomOut(Optional control As IRibbonControl)
    ' Zoom out by 10% increments - Ctrl+Alt+Shift+-
    Debug.Print "ZoomOut called"
    
    On Error Resume Next
    Dim currentZoom As Integer
    currentZoom = ActiveWindow.Zoom
    
    ' Decrease zoom in 10% increments, min 10%
    Dim newZoom As Integer
    newZoom = currentZoom - 10
    If newZoom < 10 Then newZoom = 10
    
    ActiveWindow.Zoom = newZoom
    Application.StatusBar = "Zoom: " & newZoom & "%"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Zoom changed from " & currentZoom & "% to " & newZoom & "%"
    On Error GoTo 0
End Sub

Public Sub ZoomToSelection(Optional control As IRibbonControl)
    ' Zoom to fit current selection - matches Macabacus functionality
    Debug.Print "ZoomToSelection called"
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    On Error Resume Next
    ' Store current selection
    Dim originalSelection As Range
    Set originalSelection = Selection
    
    ' Zoom to selection
    ActiveWindow.Zoom = True
    
    ' Restore selection (zoom might change it)
    originalSelection.Select
    
    Application.StatusBar = "Zoomed to selection: " & originalSelection.Address
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Zoomed to fit selection: " & originalSelection.Address
    On Error GoTo 0
End Sub

Public Sub ZoomToFit(Optional control As IRibbonControl)
    ' Zoom to fit entire used range
    Debug.Print "ZoomToFit called"
    
    On Error Resume Next
    If Not ActiveSheet.UsedRange Is Nothing Then
        Dim originalSelection As Range
        Set originalSelection = Selection
        
        ActiveSheet.UsedRange.Select
        ActiveWindow.Zoom = True
        
        ' Restore original selection
        originalSelection.Select
        
        Application.StatusBar = "Zoomed to fit worksheet data"
        Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
        
        Debug.Print "Zoomed to fit used range"
    End If
    On Error GoTo 0
End Sub

Public Sub SetZoom100(Optional control As IRibbonControl)
    ' Set zoom to 100% - standard view
    Debug.Print "SetZoom100 called"
    
    On Error Resume Next
    ActiveWindow.Zoom = 100
    Application.StatusBar = "Zoom: 100%"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Zoom set to 100%"
    On Error GoTo 0
End Sub

Public Sub SetZoom75(Optional control As IRibbonControl)
    ' Set zoom to 75% - good for overview
    Debug.Print "SetZoom75 called"
    
    On Error Resume Next
    ActiveWindow.Zoom = 75
    Application.StatusBar = "Zoom: 75%"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Zoom set to 75%"
    On Error GoTo 0
End Sub

Public Sub SetZoom125(Optional control As IRibbonControl)
    ' Set zoom to 125% - good for detailed work
    Debug.Print "SetZoom125 called"
    
    On Error Resume Next
    ActiveWindow.Zoom = 125
    Application.StatusBar = "Zoom: 125%"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Zoom set to 125%"
    On Error GoTo 0
End Sub

' === DISPLAY TOGGLES (Macabacus-aligned) ===

Public Sub ToggleGridlines(Optional control As IRibbonControl)
    ' Toggle gridlines on/off - Ctrl+Alt+Shift+G
    Debug.Print "ToggleGridlines called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveWindow.DisplayGridlines
    
    ActiveWindow.DisplayGridlines = Not currentState
    
    Application.StatusBar = "Gridlines: " & IIf(Not currentState, "ON", "OFF")
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Gridlines changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

Public Sub ToggleHeadings(Optional control As IRibbonControl)
    ' Toggle row and column headings
    Debug.Print "ToggleHeadings called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveWindow.DisplayHeadings
    
    ActiveWindow.DisplayHeadings = Not currentState
    
    Application.StatusBar = "Row/Column Headers: " & IIf(Not currentState, "ON", "OFF")
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Headings changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

Public Sub ToggleFormulas(Optional control As IRibbonControl)
    ' Toggle formula display - show formulas vs values
    Debug.Print "ToggleFormulas called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveWindow.DisplayFormulas
    
    ActiveWindow.DisplayFormulas = Not currentState
    
    Application.StatusBar = "Show Formulas: " & IIf(Not currentState, "ON", "OFF")
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Formula display changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

Public Sub ToggleZeros(Optional control As IRibbonControl)
    ' Toggle zero value display
    Debug.Print "ToggleZeros called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveWindow.DisplayZeros
    
    ActiveWindow.DisplayZeros = Not currentState
    
    Application.StatusBar = "Show Zeros: " & IIf(Not currentState, "ON", "OFF")
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Zero display changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

' === PAGE BREAKS (Macabacus-aligned) ===

Public Sub HidePageBreaks(Optional control As IRibbonControl)
    ' Hide page breaks - Ctrl+Alt+Shift+B
    Debug.Print "HidePageBreaks called"
    
    On Error Resume Next
    ActiveSheet.DisplayPageBreaks = False
    Application.StatusBar = "Page breaks hidden"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Page breaks hidden"
    On Error GoTo 0
End Sub

Public Sub ShowPageBreaks(Optional control As IRibbonControl)
    ' Show page breaks
    Debug.Print "ShowPageBreaks called"
    
    On Error Resume Next
    ActiveSheet.DisplayPageBreaks = True
    Application.StatusBar = "Page breaks visible"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Page breaks shown"
    On Error GoTo 0
End Sub

Public Sub TogglePageBreaks(Optional control As IRibbonControl)
    ' Toggle page break display
    Debug.Print "TogglePageBreaks called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveSheet.DisplayPageBreaks
    
    ActiveSheet.DisplayPageBreaks = Not currentState
    
    Application.StatusBar = "Page Breaks: " & IIf(Not currentState, "ON", "OFF")
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Page breaks changed from " & currentState & " to " & (Not currentState)
    On Error GoTo 0
End Sub

' === FREEZE PANES ===

Public Sub ToggleFreezePanes(Optional control As IRibbonControl)
    ' Toggle freeze panes at current selection
    Debug.Print "ToggleFreezePanes called"
    
    On Error Resume Next
    If ActiveWindow.FreezePanes Then
        ' Unfreeze panes
        ActiveWindow.FreezePanes = False
        Application.StatusBar = "Panes unfrozen"
        Debug.Print "Panes unfrozen"
    Else
        ' Freeze panes at current selection
        ActiveWindow.FreezePanes = True
        Application.StatusBar = "Panes frozen at " & ActiveCell.Address
        Debug.Print "Panes frozen at current selection"
    End If
    
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    On Error GoTo 0
End Sub

Public Sub FreezeTopRow(Optional control As IRibbonControl)
    ' Freeze top row only
    Debug.Print "FreezeTopRow called"
    
    On Error Resume Next
    Dim originalSelection As Range
    Set originalSelection = Selection
    
    ActiveWindow.FreezePanes = False  ' Clear existing freeze
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    ' Restore selection
    originalSelection.Select
    
    Application.StatusBar = "Top row frozen"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    Debug.Print "Top row frozen"
    On Error GoTo 0
End Sub

Public Sub FreezeFirstColumn(Optional control As IRibbonControl)
    ' Freeze first column only
    Debug.Print "FreezeFirstColumn called"
    
    On Error Resume Next
    Dim originalSelection As Range
    Set originalSelection = Selection
    
    ActiveWindow.FreezePanes = False  ' Clear existing freeze
    Range("B1").Select
    ActiveWindow.FreezePanes = True
    
    ' Restore selection
    originalSelection.Select
    
    Application.StatusBar = "First column frozen"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    Debug.Print "First column frozen"
    On Error GoTo 0
End Sub

' === VIEW MODES ===

Public Sub CycleViewMode(Optional control As IRibbonControl)
    ' Cycle through Normal → Page Break → Page Layout
    Debug.Print "CycleViewMode called"
    
    On Error Resume Next
    Select Case ActiveWindow.View
        Case xlNormalView
            ActiveWindow.View = xlPageBreakPreview
            Application.StatusBar = "Page Break Preview"
            Debug.Print "Changed to Page Break Preview"
        Case xlPageBreakPreview
            ActiveWindow.View = xlPageLayoutView
            Application.StatusBar = "Page Layout View"
            Debug.Print "Changed to Page Layout View"
        Case xlPageLayoutView
            ActiveWindow.View = xlNormalView
            Application.StatusBar = "Normal View"
            Debug.Print "Changed to Normal View"
        Case Else
            ActiveWindow.View = xlNormalView
            Application.StatusBar = "Normal View"
            Debug.Print "Set to Normal View"
    End Select
    
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    On Error GoTo 0
End Sub

Public Sub SetNormalView(Optional control As IRibbonControl)
    ' Set to Normal View
    Debug.Print "SetNormalView called"
    
    On Error Resume Next
    ActiveWindow.View = xlNormalView
    Application.StatusBar = "Normal View"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Set to Normal View"
    On Error GoTo 0
End Sub

Public Sub SetPageBreakPreview(Optional control As IRibbonControl)
    ' Set to Page Break Preview
    Debug.Print "SetPageBreakPreview called"
    
    On Error Resume Next
    ActiveWindow.View = xlPageBreakPreview
    Application.StatusBar = "Page Break Preview"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Set to Page Break Preview"
    On Error GoTo 0
End Sub

Public Sub SetPageLayoutView(Optional control As IRibbonControl)
    ' Set to Page Layout View
    Debug.Print "SetPageLayoutView called"
    
    On Error Resume Next
    ActiveWindow.View = xlPageLayoutView
    Application.StatusBar = "Page Layout View"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Set to Page Layout View"
    On Error GoTo 0
End Sub

' === WINDOW MANAGEMENT ===

Public Sub NewWindow(Optional control As IRibbonControl)
    ' Create new window for current workbook
    Debug.Print "NewWindow called"
    
    On Error Resume Next
    ActiveWorkbook.NewWindow
    Application.StatusBar = "New window created"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "New window created"
    On Error GoTo 0
End Sub

Public Sub ArrangeWindows(Optional control As IRibbonControl)
    ' Arrange windows in tiled view
    Debug.Print "ArrangeWindows called"
    
    On Error Resume Next
    Windows.Arrange xlArrangeStyleTiled
    Application.StatusBar = "Windows arranged (tiled)"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Windows arranged in tiled view"
    On Error GoTo 0
End Sub

Public Sub CascadeWindows(Optional control As IRibbonControl)
    ' Arrange windows in cascade view
    Debug.Print "CascadeWindows called"
    
    On Error Resume Next
    Windows.Arrange xlArrangeStyleCascade
    Application.StatusBar = "Windows arranged (cascade)"
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    Debug.Print "Windows arranged in cascade view"
    On Error GoTo 0
End Sub

' === FULL SCREEN MODE ===

Public Sub ToggleFullScreen(Optional control As IRibbonControl)
    ' Toggle full screen mode
    Debug.Print "ToggleFullScreen called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = Application.DisplayFullScreen
    
    Application.DisplayFullScreen = Not currentState
    
    ' Status bar message (only visible when not in full screen)
    If Not Application.DisplayFullScreen Then
        Application.StatusBar = "Full screen mode: OFF"
        Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    End If
    
    Debug.Print "Full screen mode: " & IIf(Not currentState, "ON", "OFF")
    On Error GoTo 0
End Sub

' === INTERFACE ELEMENTS ===

Public Sub ToggleFormulaBar(Optional control As IRibbonControl)
    ' Toggle formula bar visibility
    Debug.Print "ToggleFormulaBar called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = Application.DisplayFormulaBar
    
    Application.DisplayFormulaBar = Not currentState
    
    Application.StatusBar = "Formula Bar: " & IIf(Not currentState, "ON", "OFF")
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Formula bar: " & IIf(Not currentState, "ON", "OFF")
    On Error GoTo 0
End Sub

Public Sub ToggleStatusBar(Optional control As IRibbonControl)
    ' Toggle status bar visibility
    Debug.Print "ToggleStatusBar called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = Application.DisplayStatusBar
    
    Application.DisplayStatusBar = Not currentState
    Debug.Print "Status bar: " & IIf(Not currentState, "ON", "OFF")
    On Error GoTo 0
End Sub

Public Sub ToggleScrollBars(Optional control As IRibbonControl)
    ' Toggle scroll bars visibility
    Debug.Print "ToggleScrollBars called"
    
    On Error Resume Next
    With ActiveWindow
        Dim currentHState As Boolean, currentVState As Boolean
        currentHState = .DisplayHorizontalScrollBar
        currentVState = .DisplayVerticalScrollBar
        
        .DisplayHorizontalScrollBar = Not currentHState
        .DisplayVerticalScrollBar = Not currentVState
        
        Application.StatusBar = "Scroll Bars: " & IIf(Not currentHState, "ON", "OFF")
        Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
        
        Debug.Print "Scroll bars: " & IIf(Not currentHState, "ON", "OFF")
    End With
    On Error GoTo 0
End Sub

Public Sub ToggleWorksheetTabs(Optional control As IRibbonControl)
    ' Toggle worksheet tabs visibility
    Debug.Print "ToggleWorksheetTabs called"
    
    On Error Resume Next
    Dim currentState As Boolean
    currentState = ActiveWindow.DisplayWorkbookTabs
    
    ActiveWindow.DisplayWorkbookTabs = Not currentState
    
    Application.StatusBar = "Worksheet Tabs: " & IIf(Not currentState, "ON", "OFF")
    Application.OnTime Now + TimeValue("00:00:01"), "ClearStatusBar"
    
    Debug.Print "Worksheet tabs: " & IIf(Not currentState, "ON", "OFF")
    On Error GoTo 0
End Sub

' === NAVIGATION HELPERS ===

Public Sub GoToCell(Optional control As IRibbonControl)
    ' Open Go To dialog
    Debug.Print "GoToCell called"
    
    On Error Resume Next
    Application.Dialogs(xlDialogFormulaGoto).Show
    On Error GoTo 0
End Sub

Public Sub GoToSpecial(Optional control As IRibbonControl)
    ' Open Go To Special dialog
    Debug.Print "GoToSpecial called"
    
    On Error Resume Next
    Application.Dialogs(xlDialogSelectSpecial).Show
    On Error GoTo 0
End Sub

Public Sub FindAndReplace(Optional control As IRibbonControl)
    ' Open Find & Replace dialog
    Debug.Print "FindAndReplace called"
    
    On Error Resume Next
    Application.Dialogs(xlDialogFindFile).Show
    On Error GoTo 0
End Sub

' === WORKSPACE MANAGEMENT ===

Public Sub SaveCustomView(Optional control As IRibbonControl)
    ' Save current view as custom view
    Debug.Print "SaveCustomView called"
    
    Dim viewName As String
    viewName = InputBox("Enter a name for this custom view:", "Save Custom View", _
                       "View_" & Format(Now, "yyyymmdd_hhmmss"))
    
    If viewName <> "" Then
        On Error Resume Next
        ActiveWorkbook.CustomViews.Add viewName
        If Err.Number = 0 Then
            Debug.Print "Custom view saved: " & viewName
            Application.StatusBar = "Custom view '" & viewName & "' saved"
            Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
            MsgBox "Custom view '" & viewName & "' saved successfully.", vbInformation, "XLerate"
        Else
            Debug.Print "Error saving custom view: " & Err.Description
            MsgBox "Error saving custom view: " & Err.Description, vbExclamation, "XLerate"
        End If
        On Error GoTo 0
    End If
End Sub

Public Sub ShowCustomViews(Optional control As IRibbonControl)
    ' Show Custom Views dialog
    Debug.Print "ShowCustomViews called"
    
    On Error Resume Next
    Application.Dialogs(xlDialogCustomViews).Show
    On Error GoTo 0
End Sub

' === PROFESSIONAL VIEW STATES ===

Public Sub SetModelingView(Optional control As IRibbonControl)
    ' Set optimal view for financial modeling
    Debug.Print "SetModelingView called"
    
    On Error Resume Next
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayFormulas = False
        .DisplayZeros = True
        .View = xlNormalView
        .Zoom = 100
    End With
    
    With Application
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
    End With
    
    Application.StatusBar = "Modeling view applied"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Modeling view state applied"
    On Error GoTo 0
End Sub

Public Sub SetPresentationView(Optional control As IRibbonControl)
    ' Set optimal view for presentations
    Debug.Print "SetPresentationView called"
    
    On Error Resume Next
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayFormulas = False
        .DisplayZeros = False
        .View = xlNormalView
        .Zoom = 100
    End With
    
    ActiveSheet.DisplayPageBreaks = False
    
    Application.StatusBar = "Presentation view applied"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Presentation view state applied"
    On Error GoTo 0
End Sub

Public Sub SetAuditView(Optional control As IRibbonControl)
    ' Set optimal view for formula auditing
    Debug.Print "SetAuditView called"
    
    On Error Resume Next
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayFormulas = True
        .DisplayZeros = True
        .View = xlNormalView
        .Zoom = 85  ' Smaller zoom to see more
    End With
    
    Application.StatusBar = "Audit view applied (formulas visible)"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
    
    Debug.Print "Audit view state applied"
    On Error GoTo 0
End Sub