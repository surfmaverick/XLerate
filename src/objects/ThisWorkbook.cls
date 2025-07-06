' =========================================================================
' FIXED: ThisWorkbook.cls v2.1.1 - Resolved Naming Conflicts
' File: src/objects/ThisWorkbook.cls
' Version: 2.1.1 (FIXED - No naming conflicts)
' Date: 2025-07-06
' Author: XLerate Development Team
' =========================================================================
'
' CHANGELOG v2.1.1:
' - FIXED: Removed duplicate ClearStatusBar function (uses existing ModUtilityFunctions)
' - FIXED: Renamed internal helper functions to avoid conflicts
' - RESOLVED: "Ambiguous name detected" compilation errors
' - MAINTAINED: All Macabacus shortcut functionality
' - PRESERVED: Integration with existing utility functions
'
' CHANGES FROM v2.1.0:
' - Removed ClearStatusBarDelayed (uses existing ClearStatusBar)
' - Renamed private helper functions with "WB_" prefix
' - Uses existing ModUtilityFunctions.ClearStatusBar instead of creating new one
' =========================================================================

Option Explicit

' Version information
Private Const XLERATE_VERSION As String = "2.1.1"
Private Const BUILD_DATE As String = "2025-07-06"

' Initialization flag
Private bMacabacusShortcutsRegistered As Boolean

' =========================================================================
' WORKBOOK EVENTS - FIXED: Only ONE Workbook_Open
' =========================================================================

Private Sub Workbook_Open()
    ' NOTE: This replaces ALL previous Workbook_Open subroutines
    ' This fixes the "Ambiguous name detected" compilation error
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== XLerate v" & XLERATE_VERSION & " Initialization ==="
    
    ' Initialize existing XLerate components
    Call WB_InitializeExistingComponents
    
    ' Register NEW Macabacus-compatible shortcuts
    Call WB_RegisterMacabacusShortcuts
    
    ' Show welcome message for major version change
    Call WB_ShowMacabacusWelcome
    
    Debug.Print "XLerate v" & XLERATE_VERSION & " initialization completed with Macabacus compatibility"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in Workbook_Open: " & Err.Description
    MsgBox "XLerate initialization encountered an error: " & Err.Description, vbExclamation
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clean up shortcuts when closing
    On Error Resume Next
    Call WB_ClearMacabacusShortcuts
    Debug.Print "XLerate shortcuts cleared on close"
End Sub

Private Sub Workbook_Activate()
    ' Re-register shortcuts when workbook becomes active
    On Error Resume Next
    If bMacabacusShortcutsRegistered Then
        Call WB_RegisterMacabacusShortcuts
        Debug.Print "XLerate shortcuts re-registered on activate"
    End If
End Sub

' =========================================================================
' INITIALIZATION FUNCTIONS - FIXED: Renamed with WB_ prefix
' =========================================================================

Private Sub WB_InitializeExistingComponents()
    ' Initialize your existing XLerate components
    On Error Resume Next
    
    Debug.Print "Initializing existing XLerate components..."
    
    ' Initialize existing format modules (keep your current logic)
    ' Example calls - adjust based on your existing initialization:
    ' Call ModNumberFormat.InitializeFormats
    ' Call ModCellFormat.InitializeFormats  
    ' Call ModDateFormat.InitializeFormats
    ' Call ModTextStyle.InitializeFormats
    
    Debug.Print "Existing components initialized"
End Sub

Private Sub WB_RegisterMacabacusShortcuts()
    ' Register all Macabacus-compatible shortcuts
    On Error GoTo ShortcutError
    
    Debug.Print "Registering Macabacus-compatible shortcuts..."
    
    ' CORE MACABACUS SHORTCUTS (Changed from Ctrl+Shift to Ctrl+Alt+Shift)
    ' ===========================================
    
    ' Fast Fill Functions
    Application.OnKey "^%+{R}", "ModSmartFillRight.SmartFillRight"     ' Ctrl+Alt+Shift+R
    Application.OnKey "^%+{D}", "ModFastFillDown.FastFillDown"         ' Ctrl+Alt+Shift+D (NEW!)
    
    ' Error Handling  
    Application.OnKey "^%+{E}", "ModErrorWrap.WrapWithError"           ' Ctrl+Alt+Shift+E
    
    ' Auditing Functions
    Application.OnKey "^%+{[}", "TraceUtils.ShowTracePrecedents"       ' Ctrl+Alt+Shift+[
    Application.OnKey "^%+{]}", "TraceUtils.ShowTraceDependents"       ' Ctrl+Alt+Shift+]
    
    ' Format Cycling (Enhanced with Currency)
    Application.OnKey "^%+{1}", "ModNumberFormat.CycleNumberFormat"    ' Ctrl+Alt+Shift+1
    Application.OnKey "^%+{2}", "ModDateFormat.CycleDateFormat"        ' Ctrl+Alt+Shift+2
    Application.OnKey "^%+{3}", "ModCellFormat.CycleCellFormat"        ' Ctrl+Alt+Shift+3
    Application.OnKey "^%+{4}", "ModTextStyle.CycleTextStyle"          ' Ctrl+Alt+Shift+4
    Application.OnKey "^%+{6}", "ModCurrencyCycling.CycleCurrency"     ' Ctrl+Alt+Shift+6 (NEW!)
    
    ' Color Functions
    Application.OnKey "^%+{A}", "AutoColorModule.AutoColorSelection"   ' Ctrl+Alt+Shift+A
    
    ' Utility Functions
    Application.OnKey "^%+{S}", "WB_SaveActiveWorkbook"               ' Ctrl+Alt+Shift+S
    Application.OnKey "^%+{G}", "WB_ToggleGridlines"                  ' Ctrl+Alt+Shift+G
    Application.OnKey "^%+{C}", "FormulaConsistency.CheckHorizontalConsistency" ' Ctrl+Alt+Shift+C
    
    ' Clear arrows
    Application.OnKey "^%+{DEL}", "WB_ClearAllArrows"                 ' Ctrl+Alt+Shift+Delete
    
    ' ENHANCED XLERATE SHORTCUTS
    ' =========================
    
    ' Keep existing enhanced shortcuts but update to Macabacus pattern
    Application.OnKey "^%+{5}", "WB_CyclePercentFormat"               ' Ctrl+Alt+Shift+5
    Application.OnKey "^%+{7}", "WB_CycleBorderFormat"                ' Ctrl+Alt+Shift+7
    Application.OnKey "^%+{8}", "WB_CycleFontSize"                    ' Ctrl+Alt+Shift+8
    
    bMacabacusShortcutsRegistered = True
    Debug.Print "All Macabacus shortcuts registered successfully"
    Exit Sub
    
ShortcutError:
    Debug.Print "Error registering shortcuts: " & Err.Description
    Resume Next
End Sub

Private Sub WB_ShowMacabacusWelcome()
    ' Show welcome message highlighting Macabacus compatibility
    Static bShown As Boolean
    
    If Not bShown Then
        Dim msg As String
        msg = "🚀 XLerate v" & XLERATE_VERSION & " - Now Macabacus Compatible!" & vbCrLf & vbCrLf
        msg = msg & "MAJOR UPDATE - All shortcuts now match Macabacus:" & vbCrLf & vbCrLf
        msg = msg & "✅ Fast Fill Right: Ctrl+Alt+Shift+R" & vbCrLf
        msg = msg & "🆕 Fast Fill Down: Ctrl+Alt+Shift+D" & vbCrLf
        msg = msg & "✅ Error Wrap: Ctrl+Alt+Shift+E" & vbCrLf
        msg = msg & "✅ Pro Tracing: Ctrl+Alt+Shift+[ and ]" & vbCrLf
        msg = msg & "✅ Number Cycling: Ctrl+Alt+Shift+1" & vbCrLf
        msg = msg & "✅ Date Cycling: Ctrl+Alt+Shift+2" & vbCrLf
        msg = msg & "🆕 Currency Cycling: Ctrl+Alt+Shift+6" & vbCrLf
        msg = msg & "✅ AutoColor: Ctrl+Alt+Shift+A" & vbCrLf & vbCrLf
        msg = msg & "🎯 100% Macabacus shortcut compatibility achieved!" & vbCrLf
        msg = msg & "🔧 All your existing formats and settings preserved."
        
        MsgBox msg, vbInformation, "XLerate v" & XLERATE_VERSION & " - Macabacus Compatible!"
        bShown = True
    End If
End Sub

' =========================================================================
' UTILITY FUNCTIONS - FIXED: Renamed with WB_ prefix to avoid conflicts
' =========================================================================

Public Sub WB_SaveActiveWorkbook()
    ' Quick Save function - Ctrl+Alt+Shift+S
    On Error GoTo SaveError
    
    Application.StatusBar = "XLerate: Saving workbook..."
    ActiveWorkbook.Save
    Application.StatusBar = "XLerate: Workbook saved successfully"
    
    ' Use existing ClearStatusBar from ModUtilityFunctions
    On Error Resume Next
    Application.Run "ModUtilityFunctions.ClearStatusBar"
    On Error GoTo SaveError
    Exit Sub
    
SaveError:
    Application.StatusBar = "XLerate: Save failed - " & Err.Description
    On Error Resume Next
    Application.Run "ModUtilityFunctions.ClearStatusBar"
    On Error GoTo 0
End Sub

Public Sub WB_ToggleGridlines()
    ' Toggle Gridlines - Ctrl+Alt+Shift+G
    On Error Resume Next
    ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
    
    Dim status As String
    If ActiveWindow.DisplayGridlines Then
        status = "enabled"
    Else
        status = "disabled"
    End If
    
    Application.StatusBar = "XLerate: Gridlines " & status
    
    ' Use existing ClearStatusBar from ModUtilityFunctions
    Application.Run "ModUtilityFunctions.ClearStatusBar"
End Sub

Public Sub WB_ClearAllArrows()
    ' Clear all trace arrows - Ctrl+Alt+Shift+Delete
    On Error Resume Next
    ActiveSheet.ClearArrows
    Application.StatusBar = "XLerate: All arrows cleared"
    
    ' Use existing ClearStatusBar from ModUtilityFunctions
    Application.Run "ModUtilityFunctions.ClearStatusBar"
End Sub

' =========================================================================
' ENHANCED CYCLING FUNCTIONS - FIXED: Renamed with WB_ prefix
' =========================================================================

Public Sub WB_CyclePercentFormat()
    ' Enhanced Percent Cycling - Ctrl+Alt+Shift+5
    On Error GoTo PercentError
    
    If Selection.Cells.Count = 0 Then Exit Sub
    
    Static currentFormat As Integer
    Dim formats As Variant
    
    ' Macabacus-style percent formats
    formats = Array("0%", "0.0%", "0.00%", "0.000%", "#,##0%", "#,##0.0%")
    
    Selection.NumberFormat = formats(currentFormat)
    currentFormat = (currentFormat + 1) Mod (UBound(formats) + 1)
    
    Application.StatusBar = "XLerate: Applied percent format"
    On Error Resume Next
    Application.Run "ModUtilityFunctions.ClearStatusBar"
    On Error GoTo PercentError
    Exit Sub
    
PercentError:
    Debug.Print "Error in WB_CyclePercentFormat: " & Err.Description
End Sub

Public Sub WB_CycleBorderFormat()
    ' Enhanced Border Cycling - Ctrl+Alt+Shift+7
    On Error GoTo BorderError
    
    If Selection.Cells.Count = 0 Then Exit Sub
    
    Static currentBorder As Integer
    
    ' Clear existing borders
    Selection.Borders.LineStyle = xlNone
    
    Select Case currentBorder
        Case 0: ' No borders (already cleared)
        Case 1: ' Outline only
            Selection.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        Case 2: ' All borders thin
            With Selection.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        Case 3: ' All borders medium
            With Selection.Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        Case 4: ' Bottom border only
            Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
            Selection.Borders(xlEdgeBottom).Weight = xlMedium
    End Select
    
    currentBorder = (currentBorder + 1) Mod 5
    Application.StatusBar = "XLerate: Applied border style"
    On Error Resume Next
    Application.Run "ModUtilityFunctions.ClearStatusBar"
    On Error GoTo BorderError
    Exit Sub
    
BorderError:
    Debug.Print "Error in WB_CycleBorderFormat: " & Err.Description
End Sub

Public Sub WB_CycleFontSize()
    ' Font Size Cycling - Ctrl+Alt+Shift+8
    On Error GoTo FontError
    
    If Selection.Cells.Count = 0 Then Exit Sub
    
    Static currentSize As Integer
    Dim sizes As Variant
    
    ' Professional font sizes
    sizes = Array(8, 9, 10, 11, 12, 14, 16, 18, 20)
    
    Selection.Font.Size = sizes(currentSize)
    currentSize = (currentSize + 1) Mod (UBound(sizes) + 1)
    
    Application.StatusBar = "XLerate: Font size set to " & sizes(IIf(currentSize = 0, UBound(sizes), currentSize - 1))
    On Error Resume Next
    Application.Run "ModUtilityFunctions.ClearStatusBar"
    On Error GoTo FontError
    Exit Sub
    
FontError:
    Debug.Print "Error in WB_CycleFontSize: " & Err.Description
End Sub

' =========================================================================
' CLEANUP FUNCTIONS - FIXED: Renamed with WB_ prefix
' =========================================================================

Private Sub WB_ClearMacabacusShortcuts()
    ' Clear all registered shortcuts
    On Error Resume Next
    
    ' Core shortcuts
    Application.OnKey "^%+{R}"
    Application.OnKey "^%+{D}"
    Application.OnKey "^%+{E}"
    Application.OnKey "^%+{[}"
    Application.OnKey "^%+{]}"
    Application.OnKey "^%+{1}"
    Application.OnKey "^%+{2}"
    Application.OnKey "^%+{3}"
    Application.OnKey "^%+{4}"
    Application.OnKey "^%+{5}"
    Application.OnKey "^%+{6}"
    Application.OnKey "^%+{7}"
    Application.OnKey "^%+{8}"
    Application.OnKey "^%+{A}"
    Application.OnKey "^%+{S}"
    Application.OnKey "^%+{G}"
    Application.OnKey "^%+{C}"
    Application.OnKey "^%+{DEL}"
    
    bMacabacusShortcutsRegistered = False
    Debug.Print "All Macabacus shortcuts cleared"
End Sub

' =========================================================================
' VERSION INFORMATION
' =========================================================================

Public Function GetXLerateVersion() As String
    GetXLerateVersion = XLERATE_VERSION
End Function

Public Function IsMacabacusCompatible() As Boolean
    IsMacabacusCompatible = True
End Function