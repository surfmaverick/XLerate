'====================================================================
' XLERATE DYNAMIC BUTTON HANDLER CLASS MODULE
'====================================================================
' 
' Filename: clsDynamicButtonHandler.cls
' Version: v2.1.0
' Date: 2025-07-12
' Author: XLERATE Development Team
' License: MIT License
'
' Suggested Directory Structure:
' XLERATE/
' ├── src/
' │   ├── classes/
' │   │   └── clsDynamicButtonHandler.cls    ← THIS FILE
' │   ├── modules/
' │   └── forms/
' ├── docs/
' └── tests/
'
' CHANGELOG:
' ==========
' v2.1.0 (2025-07-12) - CRITICAL FIX RELEASE
' - FIXED: Undefined Parent reference causing runtime errors
' - ADDED: Optional Parent parameter with auto-detection fallback
' - ENHANCED: Comprehensive error handling and user feedback
' - IMPROVED: Input validation for Initialize method
' - ADDED: Null reference protection throughout
' - ENHANCED: Debug logging for troubleshooting
' - IMPROVED: Error messages with context information
'
' v1.0.0 (Previous) - ORIGINAL IMPLEMENTATION
' - Basic dynamic button event handling
' - Support for standard button roles (Add, Remove, Save, Cancel, etc.)
' - Simple error handling with debug output
' - Event-driven architecture with WithEvents
'
' DEPENDENCIES:
' - None (Pure VBA implementation)
'
' COMPATIBILITY:
' - Excel 2019+ (Windows/macOS)
' - Excel 365 (Desktop/Online)
' - Office 2019/2021
'
' USAGE EXAMPLE:
' Dim btnHandler As clsDynamicButtonHandler
' Set btnHandler = New clsDynamicButtonHandler
' btnHandler.Initialize myButton, "Add", Me  ' Include parent form!
'
'====================================================================

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDynamicButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsDynamicButtonHandler - Enhanced Dynamic Button Event Handler
Option Explicit

' Module Constants
Private Const MODULE_VERSION As String = "2.1.0"
Private Const MODULE_NAME As String = "clsDynamicButtonHandler"
Private Const DEBUG_MODE As Boolean = True

' Event-enabled button object
Public WithEvents DynamicButton As Object

' Private member variables
Private ButtonRole As String
Private ParentForm As Object
Private bInitialized As Boolean

'====================================================================
' PUBLIC INITIALIZATION METHOD
'====================================================================

Public Sub Initialize(ByRef Button As Object, ByVal Role As String, Optional ByRef Parent As Object = Nothing)
    ' Enhanced initialization with proper parent handling
    ' 
    ' FIXED in v2.1.0: Added Optional Parent parameter to resolve undefined reference
    ' PARAMETERS:
    '   Button - The button control to handle (required)
    '   Role - The button's functional role (required)
    '   Parent - The parent form object (optional, auto-detected if not provided)
    
    On Error GoTo ErrorHandler
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & " v" & MODULE_VERSION & ": Initializing button handler for role: " & Role
    
    ' Input validation (NEW in v2.1.0)
    If Button Is Nothing Then
        Err.Raise vbObjectError + 1001, MODULE_NAME, "Button object cannot be Nothing"
        Exit Sub
    End If
    
    If Len(Trim(Role)) = 0 Then
        Err.Raise vbObjectError + 1002, MODULE_NAME, "Role parameter cannot be empty"
        Exit Sub
    End If
    
    ' Set button reference
    Set DynamicButton = Button
    ButtonRole = Trim(Role)
    
    ' CRITICAL FIX v2.1.0: Handle parent reference properly
    If Not Parent Is Nothing Then
        ' Use provided parent form
        Set ParentForm = Parent
        If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Using provided parent form"
    Else
        ' Auto-detect parent from button (fallback method)
        On Error Resume Next
        Set ParentForm = Button.Parent
        On Error GoTo ErrorHandler
        If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Auto-detected parent form"
    End If
    
    ' Validate parent form availability
    If ParentForm Is Nothing Then
        Debug.Print MODULE_NAME & " WARNING: Parent form could not be determined for role: " & Role
    End If
    
    bInitialized = True
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Button handler initialized successfully"
    
    Exit Sub
    
ErrorHandler:
    bInitialized = False
    Debug.Print MODULE_NAME & " ERROR: Failed to initialize button handler - " & Err.Description & " (Code: " & Err.Number & ")"
    MsgBox "Failed to initialize button handler for role: " & Role & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, MODULE_NAME & " v" & MODULE_VERSION
    Err.Clear
End Sub

'====================================================================
' ENHANCED BUTTON CLICK HANDLER
'====================================================================

Private Sub DynamicButton_Click()
    ' Enhanced click handler with comprehensive error handling
    ' IMPROVED in v2.1.0: Better error handling and user feedback
    
    On Error GoTo ErrorHandler
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": DynamicButton_Click triggered for role: " & ButtonRole
    
    ' Validate initialization (NEW in v2.1.0)
    If Not bInitialized Then
        Debug.Print MODULE_NAME & " ERROR: Button handler not properly initialized"
        MsgBox "Button handler error: Not properly initialized", vbCritical, MODULE_NAME
        Exit Sub
    End If
    
    ' Validate parent form availability (ENHANCED in v2.1.0)
    If ParentForm Is Nothing Then
        Debug.Print MODULE_NAME & " ERROR: Parent form not available for role: " & ButtonRole
        MsgBox "Button handler error: Parent form not available for action: " & ButtonRole, _
               vbExclamation, MODULE_NAME & " v" & MODULE_VERSION
        Exit Sub
    End If
    
    ' Execute the appropriate action based on button role
    Select Case UCase(Trim(ButtonRole))
        Case "ADD"
            If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Calling btnAdd_Click"
            ParentForm.btnAdd_Click
            
        Case "REMOVE", "DELETE"
            If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Calling btnRemove_Click"
            ParentForm.btnRemove_Click
            
        Case "SAVE"
            If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Calling btnSave_Click"
            ParentForm.btnSave_Click
            
        Case "CANCEL"
            If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Calling btnCancel_Click"
            ParentForm.btnCancel_Click
            
        Case "BORDERCOLOR", "BORDER_COLOR"
            If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Calling btnBorderColor_Click"
            ParentForm.btnBorderColor_Click
            
        Case "FILLCOLOR", "FILL_COLOR", "BACKGROUND"
            If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Calling btnFillColor_Click"
            ParentForm.btnFillColor_Click
            
        Case "FONTCOLOR", "FONT_COLOR", "TEXT_COLOR"
            If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Calling btnFontColor_Click"
            ParentForm.btnFontColor_Click
            
        Case Else
            Debug.Print MODULE_NAME & " WARNING: Unrecognized button role: " & ButtonRole
            MsgBox "Unknown button action: " & ButtonRole, vbInformation, MODULE_NAME
    End Select
    
    Exit Sub
    
ErrorHandler:
    Debug.Print MODULE_NAME & " ERROR in DynamicButton_Click: " & Err.Description & " (Error " & Err.Number & ") - Role: " & ButtonRole
    
    ' Enhanced user feedback (NEW in v2.1.0)
    MsgBox "Button action failed: " & ButtonRole & vbCrLf & vbCrLf & _
           "Error Details: " & Err.Description & vbCrLf & _
           "Error Code: " & Err.Number, vbExclamation, MODULE_NAME & " v" & MODULE_VERSION
    
    Err.Clear
End Sub

'====================================================================
' PUBLIC PROPERTIES (NEW in v2.1.0)
'====================================================================

Public Property Get Version() As String
    ' Get module version
    Version = MODULE_VERSION
End Property

Public Property Get IsInitialized() As Boolean
    ' Check if handler is properly initialized
    IsInitialized = bInitialized
End Property

Public Property Get Role() As String
    ' Get current button role
    Role = ButtonRole
End Property

'====================================================================
' CLEANUP METHODS (ENHANCED in v2.1.0)
'====================================================================

Public Sub Cleanup()
    ' Clean up object references and reset state
    On Error Resume Next
    
    Set DynamicButton = Nothing
    Set ParentForm = Nothing
    ButtonRole = ""
    bInitialized = False
    
    If DEBUG_MODE Then Debug.Print MODULE_NAME & ": Cleanup completed"
End Sub

Private Sub Class_Terminate()
    ' Automatic cleanup when class is destroyed
    Call Cleanup
    If DEBUG_MODE Then Debug.Print MODULE_NAME & " v" & MODULE_VERSION & ": Class terminated"
End Sub