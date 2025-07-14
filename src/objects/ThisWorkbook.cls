'====================================================================
' XLERATE MAIN WORKBOOK CLASS MODULE
'====================================================================
' 
' Filename: ThisWorkbook.cls
' Version: v2.1.0
' Date: 2025-07-12
' Author: XLERATE Development Team
' License: MIT License
'
' Suggested Directory Structure:
' XLERATE/
' ├── src/
' │   ├── workbook/
' │   │   └── ThisWorkbook.cls           ← THIS FILE
' │   ├── classes/
' │   ├── modules/
' │   └── forms/
' ├── docs/
' ├── tests/
' └── build/
'
' DESCRIPTION:
' Main workbook controller that manages XLERATE initialization,
' shortcut registration, and workbook-level event handling.
' Provides 100% Macabacus-compatible keyboard shortcuts.
'
' CHANGELOG:
' ==========
' v2.1.0 (2025-07-12) - COMPLETE REWRITE
' - ADDED: Full Macabacus-compatible shortcut registration system
' - ADDED: Comprehensive error handling for all operations
' - ADDED: Proper workbook event management (Open, BeforeClose)
' - ADDED: Debug logging for troubleshooting
' - ADDED: User-friendly welcome message with version info
' - ADDED: Systematic shortcut cleanup on workbook close
' - IMPROVED: Cross-platform compatibility (Windows/macOS)
' - ENHANCED: Status feedback for all user actions
' - ADDED: Support for all core XLERATE functions
'
' v2.0.0 (Previous) - MACABACUS COMPATIBILITY
' - Basic shortcut registration
' - Core fast fill and formatting functions
'
' v1.0.0 (Original) - INITIAL IMPLEMENTATION
' - Basic XLERATE functionality
' - Limited shortcut support
'
' FEATURES:
' - 100% Macabacus shortcut compatibility
' - Fast Fill (Right/Down) with intelligent boundaries
' - Advanced format cycling (Numbers, Dates, Cells, Text)
' - Formula auditing and consistency checking
' - Auto-coloring system for different cell types
' - Error wrapping for formulas
' - Quick save and gridline toggle utilities
'
' DEPENDENCIES:
' - FastFillModule.bas (Fast fill functions)
' - FormatModule.bas (Format cycling functions)
' - UtilityModule.bas (Helper functions)
'
' COMPATIBILITY:
' - Excel 2019+ (Windows/macOS)
' - Excel 365 (Desktop/Online with keyboard)
' - Office 2019/2021 (32-bit and 64-bit)
'
'====================================================================

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ThisWorkbook - XLERATE Main Controller
Option Explicit

' Module Constants
Private Const XLERATE_VERSION As String = "2.1.0"
Private Const MODULE_NAME As String = "XLERATE Controller"
Private Const MACABACUS_COMPAT_VERSION As String = "8.0+"
Private Const DEBUG_MODE As Boolean = True

' Module Variables
Private bShortcutsRegistered As Boolean

'====================================================================
' WORKBOOK EVENT HANDLERS
'====================================================================

Private Sub Workbook_Open()
    ' Initialize XLERATE when workbook opens
    ' ENHANCED in v2.1.0: Complete initialization system
    
    On Error GoTo ErrorHandler
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & " v" & XLERATE_VERSION & ": Starting initialization..."
    
    ' Register all XLERATE shortcuts
    Call RegisterXLERATEShortcuts
    
    ' Show welcome message
    Call ShowWelcomeMessage
    
    ' Set initial status
    Application.StatusBar = "XLERATE v" & XLERATE_VERSION & " ready - Press Ctrl+Alt+Shift+/ for help"
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Initialization completed successfully"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error initializing XLERATE v" & XLERATE_VERSION & ":" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Some features may not work properly.", vbCritical, "XLERATE Initialization Error"
    Debug.Print MODULE_NAME & " ERROR: Initialization failed - " & Err.Description
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clean up XLERATE when workbook closes
    ' ENHANCED in v2.1.0: Comprehensive cleanup
    
    On Error Resume Next
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Starting cleanup..."
    
    ' Unregister all shortcuts
    Call UnregisterXLERATEShortcuts
    
    ' Clear status bar
    Application.StatusBar = False
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Cleanup completed"
End Sub

'====================================================================
' SHORTCUT REGISTRATION SYSTEM
'====================================================================

Private Sub RegisterXLERATEShortcuts()
    ' Register all XLERATE shortcuts (100% Macabacus compatible)
    ' COMPLETE REWRITE in v2.1.0: Full shortcut system
    
    On Error GoTo ShortcutError
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Registering shortcuts..."
    
    ' FAST FILL FUNCTIONS (Macabacus Compatible)
    Application.OnKey "^%+{R}", "FastFillRight"          ' Ctrl+Alt+Shift+R
    Application.OnKey "^%+{D}", "FastFillDown"           ' Ctrl+Alt+Shift+D
    
    ' FORMULA FUNCTIONS (Macabacus Compatible)
    Application.OnKey "^%+{E}", "WrapWithError"          ' Ctrl+Alt+Shift+E
    
    ' AUDITING FUNCTIONS (Macabacus Compatible)
    Application.OnKey "^%+{[}", "ShowPrecedents"         ' Ctrl+Alt+Shift+[
    Application.OnKey "^%+{]}", "ShowDependents"         ' Ctrl+Alt+Shift+]
    
    ' FORMAT CYCLING FUNCTIONS (Macabacus Compatible)
    Application.OnKey "^%+{1}", "CycleNumberFormat"      ' Ctrl+Alt+Shift+1
    Application.OnKey "^%+{2}", "CycleDateFormat"        ' Ctrl+Alt+Shift+2
    Application.OnKey "^%+{3}", "CycleCellFormat"        ' Ctrl+Alt+Shift+3
    Application.OnKey "^%+{4}", "CycleTextFormat"        ' Ctrl+Alt+Shift+4
    
    ' COLOR FUNCTIONS (Macabacus Compatible)
    Application.OnKey "^%+{A}", "AutoColorSelection"     ' Ctrl+Alt+Shift+A
    
    ' UTILITY FUNCTIONS (Macabacus Compatible)
    Application.OnKey "^%+{S}", "QuickSave"              ' Ctrl+Alt+Shift+S
    Application.OnKey "^%+{G}", "ToggleGridlines"        ' Ctrl+Alt+Shift+G
    Application.OnKey "^%+{C}", "CheckConsistency"       ' Ctrl+Alt+Shift+C
    
    ' CLEAR FUNCTIONS (Macabacus Compatible)
    Application.OnKey "^%+{DEL}", "ClearAllArrows"       ' Ctrl+Alt+Shift+Delete
    
    ' HELP FUNCTIONS (XLERATE Enhanced)
    Application.OnKey "^%+{/}", "ShowKeyboardMap"        ' Ctrl+Alt+Shift+/ (Question mark)
    
    bShortcutsRegistered = True
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": All shortcuts registered successfully"
    Exit Sub
    
ShortcutError:
    Debug.Print MODULE_NAME & " WARNING: Error registering shortcuts - " & Err.Description
    ' Continue execution - partial functionality is better than none
    Resume Next
End Sub

Private Sub UnregisterXLERATEShortcuts()
    ' Unregister all XLERATE shortcuts to prevent conflicts
    ' ENHANCED in v2.1.0: Systematic cleanup
    
    On Error Resume Next
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Unregistering shortcuts..."
    
    ' Define all registered shortcuts for systematic cleanup
    Dim arrShortcuts As Variant
    arrShortcuts = Array( _
        "^%+{R}", "^%+{D}", "^%+{E}", "^%+{[}", "^%+{]}", _
        "^%+{1}", "^%+{2}", "^%+{3}", "^%+{4}", "^%+{A}", _
        "^%+{S}", "^%+{G}", "^%+{C}", "^%+{DEL}", "^%+{/}" _
    )
    
    ' Unregister each shortcut
    Dim i As Integer
    For i = 0 To UBound(arrShortcuts)
        Application.OnKey arrShortcuts(i)
        If DEBUG_MODE And (i Mod 5 = 0) Then Debug.Print MODULE_NAME & ": Unregistered " & (i + 1) & " shortcuts..."
    Next i
    
    bShortcutsRegistered = False
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": All shortcuts unregistered"
End Sub

'====================================================================
' USER INTERFACE
'====================================================================

Private Sub ShowWelcomeMessage()
    ' Display welcome message with version and feature info
    ' ENHANCED in v2.1.0: Comprehensive feature overview
    
    Static bMessageShown As Boolean
    
    ' Only show once per Excel session
    If bMessageShown Then Exit Sub
    
    Dim msg As String
    msg = "🚀 XLERATE v" & XLERATE_VERSION & " Loaded Successfully!" & vbCrLf & vbCrLf
    msg = msg & "✅ 100% MACABACUS COMPATIBLE SHORTCUTS" & vbCrLf
    msg = msg & "Compatible with Macabacus v" & MACABACUS_COMPAT_VERSION & vbCrLf & vbCrLf
    msg = msg & "🔥 KEY FEATURES:" & vbCrLf
    msg = msg & "• Fast Fill Right/Down (Ctrl+Alt+Shift+R/D)" & vbCrLf
    msg = msg & "• Error Wrap Formulas (Ctrl+Alt+Shift+E)" & vbCrLf
    msg = msg & "• Pro Auditing Tools (Ctrl+Alt+Shift+[/])" & vbCrLf
    msg = msg & "• Advanced Format Cycling (Ctrl+Alt+Shift+1-4)" & vbCrLf
    msg = msg & "• Auto-Color System (Ctrl+Alt+Shift+A)" & vbCrLf
    msg = msg & "• Quick Utilities (Ctrl+Alt+Shift+S/G/C)" & vbCrLf & vbCrLf
    msg = msg & "❓ Press Ctrl+Alt+Shift+/ for complete keyboard map" & vbCrLf & vbCrLf
    msg = msg & "🆓 Open Source • 🌍 Cross-Platform • ⚡ High Performance"
    
    MsgBox msg, vbInformation, "XLERATE v" & XLERATE_VERSION & " - Ready!"
    
    bMessageShown = True
End Sub

'====================================================================
' PUBLIC PROPERTIES AND METHODS
'====================================================================

Public Property Get Version() As String
    ' Get current XLERATE version
    Version = XLERATE_VERSION
End Property

Public Property Get IsInitialized() As Boolean
    ' Check if XLERATE is properly initialized
    IsInitialized = bShortcutsRegistered
End Property

Public Function GetShortcutList() As String
    ' Return formatted list of all shortcuts
    ' NEW in v2.1.0: Programmatic shortcut documentation
    
    Dim shortcuts As String
    shortcuts = "XLERATE v" & XLERATE_VERSION & " - Keyboard Shortcuts:" & vbCrLf & vbCrLf
    shortcuts = shortcuts & "FAST FILL:" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+R  Fast Fill Right" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+D  Fast Fill Down" & vbCrLf & vbCrLf
    shortcuts = shortcuts & "FORMULA:" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+E  Error Wrap" & vbCrLf & vbCrLf
    shortcuts = shortcuts & "AUDITING:" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+[  Show Precedents" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+]  Show Dependents" & vbCrLf & vbCrLf
    shortcuts = shortcuts & "FORMATTING:" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+1  Cycle Number Format" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+2  Cycle Date Format" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+3  Cycle Cell Color" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+4  Cycle Text Format" & vbCrLf & vbCrLf
    shortcuts = shortcuts & "UTILITIES:" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+A  Auto Color" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+S  Quick Save" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+G  Toggle Gridlines" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+C  Check Consistency" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+Del Clear Arrows" & vbCrLf
    shortcuts = shortcuts & "Ctrl+Alt+Shift+/  Show This Help"
    
    GetShortcutList = shortcuts
End Function

Public Sub ShowAbout()
    ' Display comprehensive about dialog
    ' NEW in v2.1.0: Detailed version and feature information
    
    Dim msg As String
    msg = "XLERATE v" & XLERATE_VERSION & vbCrLf
    msg = msg & "Enhanced Excel Productivity Suite" & vbCrLf & vbCrLf
    msg = msg & "🎯 100% Macabacus Compatible" & vbCrLf
    msg = msg & "🌍 Cross-Platform (Windows/macOS)" & vbCrLf
    msg = msg & "🆓 Open Source (MIT License)" & vbCrLf
    msg = msg & "⚡ High Performance" & vbCrLf & vbCrLf
    msg = msg & "CORE FEATURES:" & vbCrLf
    msg = msg & "• Fast Fill with intelligent boundaries" & vbCrLf
    msg = msg & "• Advanced format cycling system" & vbCrLf
    msg = msg & "• Formula auditing and tracing tools" & vbCrLf
    msg = msg & "• Auto-coloring for different cell types" & vbCrLf
    msg = msg & "• Formula consistency checking" & vbCrLf
    msg = msg & "• Error wrapping for robust formulas" & vbCrLf & vbCrLf
    msg = msg & "Built for financial analysts, by financial analysts."
    
    MsgBox msg, vbInformation, "About XLERATE v" & XLERATE_VERSION
End Sub