Attribute VB_Name = "ModSettings"
Option Explicit

Public Sub ShowSettings()
    Debug.Print "ShowSettings procedure called"
    
    Static isShowing As Boolean
    If isShowing Then Exit Sub  ' Prevent multiple instances
    isShowing = True
    
    On Error GoTo ErrorHandler
    
    Dim frm As frmSettingsManager
    Set frm = New frmSettingsManager
    frm.Show vbModal
    
    isShowing = False
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ShowSettings: " & Err.Description
    isShowing = False
End Sub
