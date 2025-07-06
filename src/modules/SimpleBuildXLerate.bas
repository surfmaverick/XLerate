' =============================================================================
' File: SimpleBuildXLerate.bas
' Version: 2.1.0 - Clean & Simple Build Script
' Description: Simplified, bulletproof build script for XLerate
' Author: XLerate Development Team
' Created: January 2025
' =============================================================================

Option Explicit

Private Const XLERATE_VERSION As String = "2.1.0"
Private Const BUILD_CODENAME As String = "Macabacus Professional"

Public Sub BuildXLerate()
    ' Simple, clean build process
    Debug.Print "=== XLerate v" & XLERATE_VERSION & " Simple Build ==="
    
    On Error GoTo BuildError
    
    ' Set up paths - EDIT THESE IF NEEDED
    Dim sourcePath As String
    Dim outputPath As String
    
    sourcePath = "C:\Mac\Home\Documents\Coding\GitHub\XLerate\src\"
    outputPath = "C:\Users\chris\Desktop\XLerate_v" & Replace(XLERATE_VERSION, ".", "_") & "_" & Replace(BUILD_CODENAME, " ", "_") & ".xlam"
    
    Debug.Print "Source: " & sourcePath
    Debug.Print "Output: " & outputPath
    
    ' Quick validation
    If Not FolderExists(sourcePath) Then
        MsgBox "Source folder not found: " & sourcePath, vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Create new workbook for add-in
    Debug.Print "Creating new add-in workbook..."
    Dim newAddin As Workbook
    Set newAddin = Workbooks.Add
    
    ' Remove extra sheets
    Do While newAddin.Worksheets.Count > 1
        newAddin.Worksheets(newAddin.Worksheets.Count).Delete
    Loop
    
    ' Set up info sheet
    newAddin.Worksheets(1).Name = "XLerate_Info"
    With newAddin.Worksheets(1)
        .Cells(1, 1) = "XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ")"
        .Cells(2, 1) = "Macabacus-Compatible Excel Add-in"
        .Cells(3, 1) = "Built: " & Now
        .Cells(4, 1) = ""
        .Cells(5, 1) = "Quick Start Shortcuts:"
        .Cells(6, 1) = "• Fast Fill Right: Ctrl+Alt+Shift+R"
        .Cells(7, 1) = "• Fast Fill Down: Ctrl+Alt+Shift+D"
        .Cells(8, 1) = "• Pro Precedents: Ctrl+Alt+Shift+["
        .Cells(9, 1) = "• Number Cycle: Ctrl+Alt+Shift+1"
        .Cells(10, 1) = "• AutoColor: Ctrl+Alt+Shift+A"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
    End With
    
    ' Import all modules
    Debug.Print "Importing modules..."
    ImportAllModules sourcePath, newAddin
    
    ' Set add-in properties
    Debug.Print "Setting add-in properties..."
    With newAddin
        .Title = "XLerate v" & XLERATE_VERSION
        .Subject = "Macabacus-Compatible Excel Add-in"
        .Author = "XLerate Development Team"
        .Comments = "Enhanced Excel add-in with Macabacus-style shortcuts"
        .IsAddin = True
    End With
    
    ' Save as add-in
    Debug.Print "Saving add-in..."
    If FileExists(outputPath) Then Kill outputPath
    newAddin.SaveAs outputPath, xlAddIn
    
    ' Clean up
    newAddin.Close False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Debug.Print "=== Build Complete! ==="
    
    ' Success message
    MsgBox "XLerate v" & XLERATE_VERSION & " build completed!" & vbNewLine & vbNewLine & _
           "Saved to: " & outputPath & vbNewLine & vbNewLine & _
           "Next steps:" & vbNewLine & _
           "1. Install the add-in in Excel" & vbNewLine & _
           "2. Enable macros" & vbNewLine & _
           "3. Try Ctrl+Alt+Shift+R for Fast Fill Right!", _
           vbInformation, "Build Complete"
    
    Exit Sub
    
BuildError:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Debug.Print "Build Error: " & Err.Description
    MsgBox "Build failed: " & Err.Description, vbCritical
    If Not newAddin Is Nothing Then newAddin.Close False
End Sub

Private Sub ImportAllModules(sourcePath As String, targetWB As Workbook)
    ' Import all modules with simple error handling
    
    ' Standard modules
    Debug.Print "  Importing standard modules..."
    ImportModulesFromFolder sourcePath & "modules\", targetWB, "*.bas"
    
    ' Class modules  
    Debug.Print "  Importing class modules..."
    ImportModulesFromFolder sourcePath & "class modules\", targetWB, "*.cls"
    
    ' Forms
    Debug.Print "  Importing forms..."
    ImportModulesFromFolder sourcePath & "forms\", targetWB, "*.frm"
    
    ' Update ThisWorkbook
    Debug.Print "  Updating ThisWorkbook..."
    UpdateThisWorkbook sourcePath & "objects\ThisWorkbook.cls", targetWB
End Sub

Private Sub ImportModulesFromFolder(folderPath As String, targetWB As Workbook, filePattern As String)
    ' Import all files matching pattern from folder
    
    If Not FolderExists(folderPath) Then
        Debug.Print "    Skipping missing folder: " & folderPath
        Exit Sub
    End If
    
    Dim fileName As String
    fileName = Dir(folderPath & filePattern)
    
    Do While fileName <> ""
        Dim filePath As String
        filePath = folderPath & fileName
        
        On Error Resume Next
        targetWB.VBProject.VBComponents.Import filePath
        If Err.Number = 0 Then
            Debug.Print "    ✓ " & fileName
        Else
            Debug.Print "    ✗ " & fileName & " (Error: " & Err.Description & ")"
        End If
        On Error GoTo 0
        
        fileName = Dir()
    Loop
End Sub

Private Sub UpdateThisWorkbook(filePath As String, targetWB As Workbook)
    ' Update ThisWorkbook with source code
    
    If Not FileExists(filePath) Then
        Debug.Print "    ✗ ThisWorkbook.cls not found"
        Exit Sub
    End If
    
    On Error Resume Next
    Dim fileContent As String
    fileContent = ReadFile(filePath)
    
    If fileContent <> "" Then
        With targetWB.VBProject.VBComponents("ThisWorkbook").CodeModule
            .DeleteLines 1, .CountOfLines
            .AddFromString fileContent
        End With
        Debug.Print "    ✓ ThisWorkbook updated"
    Else
        Debug.Print "    ✗ Could not read ThisWorkbook.cls"
    End If
    On Error GoTo 0
End Sub

' === UTILITY FUNCTIONS ===

Private Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

Private Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Private Function ReadFile(filePath As String) As String
    On Error GoTo ReadError
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    ReadFile = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    Exit Function
    
ReadError:
    ReadFile = ""
    If fileNum > 0 Then Close #fileNum
End Function

' === QUICK TEST FUNCTIONS ===

Public Sub QuickTest()
    ' Test if paths are correct
    Dim sourcePath As String
    sourcePath = "C:\Mac\Home\Documents\Coding\GitHub\XLerate\src\"
    
    Debug.Print "=== Quick Path Test ==="
    Debug.Print "Source exists: " & FolderExists(sourcePath)
    Debug.Print "Modules exists: " & FolderExists(sourcePath & "modules\")
    Debug.Print "Class modules exists: " & FolderExists(sourcePath & "class modules\")
    Debug.Print "Objects exists: " & FolderExists(sourcePath & "objects\")
    Debug.Print "Forms exists: " & FolderExists(sourcePath & "forms\")
    
    ' Test key files
    Debug.Print "ThisWorkbook.cls exists: " & FileExists(sourcePath & "objects\ThisWorkbook.cls")
    Debug.Print "ModNumberFormat.bas exists: " & FileExists(sourcePath & "modules\ModNumberFormat.bas")
    Debug.Print "RibbonCallbacks.bas exists: " & FileExists(sourcePath & "modules\RibbonCallbacks.bas")
    
    Debug.Print "=== Test Complete ==="
End Sub

Public Sub ListModules()
    ' List all modules that will be imported
    Dim sourcePath As String
    sourcePath = "C:\Mac\Home\Documents\Coding\GitHub\XLerate\src\"
    
    Debug.Print "=== Modules to Import ==="
    
    Debug.Print "Standard Modules (.bas):"
    ListFilesInFolder sourcePath & "modules\", "*.bas"
    
    Debug.Print "Class Modules (.cls):"
    ListFilesInFolder sourcePath & "class modules\", "*.cls"
    
    Debug.Print "Forms (.frm):"
    ListFilesInFolder sourcePath & "forms\", "*.frm"
    
    Debug.Print "=== List Complete ==="
End Sub

Private Sub ListFilesInFolder(folderPath As String, filePattern As String)
    If Not FolderExists(folderPath) Then
        Debug.Print "  Folder not found: " & folderPath
        Exit Sub
    End If
    
    Dim fileName As String
    fileName = Dir(folderPath & filePattern)
    
    Do While fileName <> ""
        Debug.Print "  " & fileName
        fileName = Dir()
    Loop
End Sub

' === INSTALLATION HELPER ===

Public Sub InstallXLerate()
    ' Helper to install the built add-in
    Dim addinPath As String
    addinPath = "C:\Users\chris\Desktop\XLerate_v" & Replace(XLERATE_VERSION, ".", "_") & "_" & Replace(BUILD_CODENAME, " ", "_") & ".xlam"
    
    If Not FileExists(addinPath) Then
        MsgBox "Add-in not found. Please build it first using BuildXLerate()", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    AddIns.Add addinPath
    AddIns("XLerate").Installed = True
    
    If Err.Number = 0 Then
        MsgBox "XLerate installed successfully!" & vbNewLine & _
               "Check Excel Add-ins to enable it.", vbInformation
    Else
        MsgBox "Installation failed: " & Err.Description & vbNewLine & vbNewLine & _
               "Try installing manually through Excel Add-ins.", vbExclamation
    End If
    On Error GoTo 0
End Sub