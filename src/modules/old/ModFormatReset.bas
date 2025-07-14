Attribute VB_Name = "ModFormatReset"
' ModFormatReset
Option Explicit  ' At the top of the module

Public Sub ResetAllFormatsToDefaults()
    Debug.Print "=== ResetAllFormatsToDefaults START ==="
    MsgBox "Starting format reset...", vbInformation  ' Add this line for testing
    
    ' Delete all saved formats
    On Error Resume Next
    With ThisWorkbook.CustomDocumentProperties
        .item("SavedCellFormats").Delete
        Debug.Print "Deleted SavedCellFormats"
        .item("SavedDateFormats").Delete
        Debug.Print "Deleted SavedDateFormats"
        .item("SavedNumberFormats").Delete
        Debug.Print "Deleted SavedNumberFormats"
    End With
    On Error GoTo 0
    
    ThisWorkbook.Save
    Debug.Print "Saved formats deleted"
    
    ' Force close the settings form if it's open
    On Error Resume Next
    Unload frmSettingsManager
    On Error GoTo 0
    
    ' Reinitialize all formats
    ModCellFormat.InitializeCellFormats
    ModDateFormat.InitializeDateFormats
    ModNumberFormat.InitializeFormats
    
    Debug.Print "Formats reinitialized"
    Debug.Print "=== ResetAllFormatsToDefaults END ==="
    
    MsgBox "All formats have been reset to defaults." & vbNewLine & _
           "Please close and reopen Excel to see the changes.", vbInformation
End Sub

' Add this to ModFormatReset
Public Sub TestShortcutRegistration()
    Debug.Print "Testing shortcut registration..."
    Application.OnKey "^+0", "ResetAllFormatsToDefaults"
    MsgBox "Shortcut re-registered. Try Ctrl+Shift+0 now.", vbInformation
End Sub
