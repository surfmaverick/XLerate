' ModBorderUtilities.bas
' Version: 1.0.0
' Date: 2025-01-04
' Author: XLerate Development Team
' 
' CHANGELOG:
' v1.0.0 - Initial implementation of border utility functions
'        - Comprehensive border application functions for financial modeling
'        - Consistent border styles and weights
'        - Error handling for edge cases
'
' DESCRIPTION:
' Provides comprehensive border management functions for Excel financial modeling
' Supports Macabacus-style border shortcuts and formatting

Attribute VB_Name = "ModBorderUtilities"
Option Explicit

' Border style constants
Private Const DEFAULT_BORDER_STYLE As Long = xlContinuous
Private Const DEFAULT_BORDER_WEIGHT As Long = xlThin
Private Const THICK_BORDER_WEIGHT As Long = xlThick
Private Const DEFAULT_BORDER_COLOR As Long = RGB(0, 0, 0) ' Black

Public Sub ApplyBottomBorder(Optional control As IRibbonControl)
    ' Applies bottom border to selected range
    ' Matches Macabacus bottom border functionality
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = DEFAULT_BORDER_STYLE
        .Weight = DEFAULT_BORDER_WEIGHT
        .Color = DEFAULT_BORDER_COLOR
    End With
    
    Debug.Print "Applied bottom border to " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ApplyBottomBorder: " & Err.Description
End Sub

Public Sub ApplyTopBorder(Optional control As IRibbonControl)
    ' Applies top border to selected range
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = DEFAULT_BORDER_STYLE
        .Weight = DEFAULT_BORDER_WEIGHT
        .Color = DEFAULT_BORDER_COLOR
    End With
    
    Debug.Print "Applied top border to " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ApplyTopBorder: " & Err.Description
End Sub

Public Sub ApplyLeftBorder(Optional control As IRibbonControl)
    ' Applies left border to selected range
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = DEFAULT_BORDER_STYLE
        .Weight = DEFAULT_BORDER_WEIGHT
        .Color = DEFAULT_BORDER_COLOR
    End With
    
    Debug.Print "Applied left border to " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ApplyLeftBorder: " & Err.Description
End Sub

Public Sub ApplyRightBorder(Optional control As IRibbonControl)
    ' Applies right border to selected range
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = DEFAULT_BORDER_STYLE
        .Weight = DEFAULT_BORDER_WEIGHT
        .Color = DEFAULT_BORDER_COLOR
    End With
    
    Debug.Print "Applied right border to " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ApplyRightBorder: " & Err.Description
End Sub

Public Sub ApplyOutsideBorder(Optional control As IRibbonControl)
    ' Applies border around the entire selection
    ' Matches Macabacus outside border functionality
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    ' Apply all outside borders
    With Selection
        .Borders(xlEdgeTop).LineStyle = DEFAULT_BORDER_STYLE
        .Borders(xlEdgeTop).Weight = DEFAULT_BORDER_WEIGHT
        .Borders(xlEdgeTop).Color = DEFAULT_BORDER_COLOR
        
        .Borders(xlEdgeBottom).LineStyle = DEFAULT_BORDER_STYLE
        .Borders(xlEdgeBottom).Weight = DEFAULT_BORDER_WEIGHT
        .Borders(xlEdgeBottom).Color = DEFAULT_BORDER_COLOR
        
        .Borders(xlEdgeLeft).LineStyle = DEFAULT_BORDER_STYLE
        .Borders(xlEdgeLeft).Weight = DEFAULT_BORDER_WEIGHT
        .Borders(xlEdgeLeft).Color = DEFAULT_BORDER_COLOR
        
        .Borders(xlEdgeRight).LineStyle = DEFAULT_BORDER_STYLE
        .Borders(xlEdgeRight).Weight = DEFAULT_BORDER_WEIGHT
        .Borders(xlEdgeRight).Color = DEFAULT_BORDER_COLOR
    End With
    
    Debug.Print "Applied outside border to " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ApplyOutsideBorder: " & Err.Description
End Sub

Public Sub RemoveAllBorders(Optional control As IRibbonControl)
    ' Removes all borders from selected range
    ' Matches Macabacus no border functionality
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Selection.Borders.LineStyle = xlNone
    
    Debug.Print "Removed all borders from " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in RemoveAllBorders: " & Err.Description
End Sub

Public Sub ApplyThickBottomBorder(Optional control As IRibbonControl)
    ' Applies thick bottom border - useful for totals and section separators
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = DEFAULT_BORDER_STYLE
        .Weight = THICK_BORDER_WEIGHT
        .Color = DEFAULT_BORDER_COLOR
    End With
    
    Debug.Print "Applied thick bottom border to " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ApplyThickBottomBorder: " & Err.Description
End Sub

Public Sub ApplyDoubleBorder(Optional control As IRibbonControl)
    ' Applies double bottom border - useful for final totals
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThin
        .Color = DEFAULT_BORDER_COLOR
    End With
    
    Debug.Print "Applied double border to " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ApplyDoubleBorder: " & Err.Description
End Sub

Public Sub CycleBorderStyle(Optional control As IRibbonControl)
    ' Cycles through different border styles on the bottom edge
    ' Useful for quickly formatting financial statements
    
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Dim currentStyle As Long
    currentStyle = Selection.Borders(xlEdgeBottom).LineStyle
    
    Select Case currentStyle
        Case xlNone
            ' No border -> Thin border
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = DEFAULT_BORDER_COLOR
            End With
            
        Case xlContinuous
            ' Check current weight
            If Selection.Borders(xlEdgeBottom).Weight = xlThin Then
                ' Thin -> Thick
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                    .Color = DEFAULT_BORDER_COLOR
                End With
            Else
                ' Thick -> Double
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlDouble
                    .Weight = xlThin
                    .Color = DEFAULT_BORDER_COLOR
                End With
            End If
            
        Case xlDouble
            ' Double -> No border (complete cycle)
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
            
        Case Else
            ' Unknown style -> Start with thin
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = DEFAULT_BORDER_COLOR
            End With
    End Select
    
    Debug.Print "Cycled border style for " & Selection.Address
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in CycleBorderStyle: " & Err.Description
End Sub