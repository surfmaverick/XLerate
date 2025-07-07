' =========================================================================
' File: src/modules/SimpleBuildFix.bas
' Version: 2.2.1
' Date: July 2025
' Description: Simple build script fix with no Unicode characters
'
' CHANGELOG:
' v2.2.1 - Completely removed all Unicode characters for maximum compatibility
'        - Simplified build process with clear ASCII-only messaging
'        - Added file selection dialog
'        - Enhanced error reporting
' =========================================================================

Attribute VB_Name = "SimpleBuildFix"
Option Explicit

Private Const XLERATE_VERSION As String = "2.2.1"

Public Sub BuildXLerateSimple()
    ' Simplified build with file selection and no Unicode issues
    
    Dim startTime As Date
    startTime = Now()
    
    Debug.Print "========================================"
    Debug.Print "XLerate v" & XLERATE_VERSION & " Build Started"
    Debug.Print "Time: " & Format(startTime, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "========================================"
    
    On Error GoTo BuildError
    
    ' Check VBA access first
    Debug.Print "Checking VBA project access..."
    If Not TestVBAAccess() Then
        MsgBox "VBA project access is required for building." & vbNewLine & vbNewLine & _
               "Please enable:" & vbNewLine & _
               "File > Options > Trust Center > Macro Settings >" & vbNewLine & _
               "'Trust access to VBA project object model'", vbCritical, "VBA Access Required"
        Exit Sub
    End If
    Debug.Print "[SUCCESS] VBA project access enabled"
    
    ' Get source path
    Debug.Print "Configuring source path..."
    Dim sourcePath As String
    sourcePath = GetSourcePath()
    If sourcePath = "" Then
        MsgBox "Could not find or select source path.", vbCritical, "Source Path Error"
        Exit Sub
    End If
    Debug.Print "[SUCCESS] Source path: " & sourcePath
    
    ' Get output file location
    Debug.Print "Selecting output location..."
    Dim outputPath As String
    outputPath = GetOutputPath()
    If outputPath = "" Then
        Debug.Print "Build cancelled by user"
        Exit Sub
    End If
    Debug.Print "[SUCCESS] Output file: " & outputPath
    
    ' Validate source structure
    Debug.Print "Validating source structure..."
    If Not ValidateSource(sourcePath) Then
        MsgBox "Source validation failed. Check Debug window for details.", vbCritical, "Source Validation Error"
        Exit Sub
    End If
    Debug.Print "[SUCCESS] Source validation passed"
    
    ' Start build process
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Building XLerate v" & XLERATE_VERSION & "..."
    
    Debug.Print "Creating new add-in workbook..."
    Dim newAddin As Workbook
    Set newAddin = Workbooks.Add
    
    ' Configure workbook
    Do While newAddin.Worksheets.Count > 1
        newAddin.Worksheets(newAddin.Worksheets.Count).Delete
    Loop
    newAddin.Worksheets(1).Name = "XLerate_Info"
    
    ' Add info to worksheet
    With newAddin.Worksheets(1)
        .Cells(1, 1) = "XLerate v" & XLERATE_VERSION
        .Cells(2, 1) = "Macabacus-Compatible Excel Add-in"
        .Cells(3, 1) = "Built: " & Format(startTime, "yyyy-mm-dd hh:nn:ss")
        .Cells(4, 1) = ""
        .Cells(5, 1) = "MACABACUS SHORTCUTS:"
        .Cells(6, 1) = "Fast Fill Right: Ctrl+Alt+Shift+R"
        .Cells(7, 1) = "Fast Fill Down: Ctrl+Alt+Shift+D"
        .Cells(8, 1) = "Pro Precedents: Ctrl+Alt+Shift+["
        .Cells(9, 1) = "Number Cycle: Ctrl+Alt+Shift+1"
        .Cells(10, 1) = "AutoColor: Ctrl+Alt+Shift+A"
    End With
    
    Debug.Print "[SUCCESS] Workbook configured"
    
    ' Import modules
    Debug.Print "Importing modules..."
    Dim moduleCount As Integer
    moduleCount = ImportAllModules(sourcePath, newAddin)
    Debug.Print "[SUCCESS] Imported " & moduleCount & " modules"
    
    ' Set add-in properties
    Debug.Print "Setting add-in properties..."
    With newAddin
        .Title = "XLerate v" & XLERATE_VERSION
        .Subject = "Macabacus-Compatible Excel Add-in"
        .Author = "XLerate Development Team"
        .IsAddin = True
    End With
    Debug.Print "[SUCCESS] Properties set"
    
    ' Save add-in
    Debug.Print "Saving add-in..."
    If Dir(outputPath) <> "" Then Kill outputPath ' Delete existing file
    newAddin.SaveAs outputPath, xlAddIn
    Debug.Print "[SUCCESS] Add-in saved"
    
    ' Validate saved file
    newAddin.Close False
    Set newAddin = Nothing
    
    If Dir(outputPath) <> "" Then
        Dim fileSize As Long
        fileSize = FileLen(outputPath)
        Debug.Print "[SUCCESS] File created: " & Format(fileSize, "#,##0") & " bytes"
    Else
        Debug.Print "[ERROR] File was not created"
    End If
    
    ' Cleanup and completion
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Dim buildTime As String
    buildTime = Format(Now() - startTime, "nn:ss")
    
    Debug.Print "========================================"
    Debug.Print "BUILD COMPLETED SUCCESSFULLY"
    Debug.Print "Build time: " & buildTime
    Debug.Print "========================================"
    
    MsgBox "XLerate v" & XLERATE_VERSION & " built successfully!" & vbNewLine & vbNewLine & _
           "Output: " & outputPath & vbNewLine & _
           "Build time: " & buildTime & vbNewLine & vbNewLine & _
           "Next steps:" & vbNewLine & _
           "1. Install the add-in in Excel" & vbNewLine & _
           "2. Enable macros" & vbNewLine & _
           "3. Test Ctrl+Alt+Shift+R for Fast Fill Right!", _
           vbInformation, "Build Successful"
    
    Exit Sub
    
BuildError:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    If Not newAddin Is Nothing Then
        newAddin.Close False
        Set newAddin = Nothing
    End If
    
    Debug.Print "========================================"
    Debug.Print "BUILD FAILED"
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Debug.Print "========================================"
    
    MsgBox "Build failed!" & vbNewLine & vbNewLine & _
           "Error: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
           "Check the Debug window (Ctrl+G) for details.", _
           vbCritical, "Build Failed"
End Sub

Private Function GetSourcePath() As String
    ' Try common source paths first
    Dim paths As Variant
    paths = Array( _
        ThisWorkbook.Path & "\src\", _
        ThisWorkbook.Path & "\", _
        "C:\Mac\Home\Documents\Coding\GitHub\XLerate\src\", _
        Environ("USERPROFILE") & "\Documents\XLerate\src\" _
    )
    
    Dim i As Integer
    For i = 0 To UBound(paths)
        Debug.Print "Checking: " & paths(i)
        If Dir(paths(i), vbDirectory) <> "" Then
            GetSourcePath = paths(i)
            Exit Function
        End If
    Next i
    
    ' If not found, ask user
    Debug.Print "Auto-detection failed, prompting user..."
    Dim folderPicker As Object
    Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With folderPicker
        .Title = "Select XLerate Source Folder"
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then
            GetSourcePath = .SelectedItems(1) & "\"
        End If
    End With
End Function

Private Function GetOutputPath() As String
    Dim saveDialog As Object
    Set saveDialog = Application.FileDialog(msoFileDialogSaveAs)
    
    Dim defaultName As String
    defaultName = "XLerate_v" & Replace(XLERATE_VERSION, ".", "_") & ".xlam"
    
    With saveDialog
        .Title = "Save XLerate Add-in As"
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\" & defaultName
        .Filters.Clear
        .Filters.Add "Excel Add-ins", "*.xlam"
        
        If .Show = -1 Then
            GetOutputPath = .SelectedItems(1)
        End If
    End With
End Function

Private Function ValidateSource(sourcePath As String) As Boolean
    ValidateSource = True
    
    ' Check for required directories
    If Dir(sourcePath & "modules\", vbDirectory) = "" Then
        Debug.Print "[ERROR] Missing modules directory"
        ValidateSource = False
    Else
        Debug.Print "[SUCCESS] Found modules directory"
    End If
    
    If Dir(sourcePath & "objects\", vbDirectory) = "" Then
        Debug.Print "[ERROR] Missing objects directory"
        ValidateSource = False
    Else
        Debug.Print "[SUCCESS] Found objects directory"
    End If
    
    ' Check for critical files
    If Dir(sourcePath & "objects\ThisWorkbook.cls") = "" Then
        Debug.Print "[ERROR] Missing ThisWorkbook.cls"
        ValidateSource = False
    Else
        Debug.Print "[SUCCESS] Found ThisWorkbook.cls"
    End If
End Function

Private Function ImportAllModules(sourcePath As String, targetWB As Workbook) As Integer
    Dim count As Integer
    count = 0
    
    ' Import standard modules
    count = count + ImportModulesFromFolder(sourcePath & "modules\", targetWB, "*.bas")
    
    ' Import class modules
    count = count + ImportModulesFromFolder(sourcePath & "class modules\", targetWB, "*.cls")
    
    ' Import forms
    count = count + ImportModulesFromFolder(sourcePath & "forms\", targetWB, "*.frm")
    
    ' Update ThisWorkbook
    UpdateThisWorkbook sourcePath & "objects\ThisWorkbook.cls", targetWB
    
    ImportAllModules = count
End Function

Private Function ImportModulesFromFolder(folderPath As String, targetWB As Workbook, pattern As String) As Integer
    If Dir(folderPath, vbDirectory) = "" Then
        Debug.Print "[WARNING] Folder not found: " & folderPath
        ImportModulesFromFolder = 0
        Exit Function
    End If
    
    Dim fileName As String
    Dim count As Integer
    count = 0
    
    fileName = Dir(folderPath & pattern)
    Do While fileName <> ""
        On Error Resume Next
        targetWB.VBProject.VBComponents.Import folderPath & fileName
        If Err.Number = 0 Then
            Debug.Print "[SUCCESS] Imported: " & fileName
            count = count + 1
        Else
            Debug.Print "[ERROR] Failed to import: " & fileName & " - " & Err.Description
        End If
        On Error GoTo 0
        fileName = Dir
    Loop
    
    ImportModulesFromFolder = count
End Function

Private Sub UpdateThisWorkbook(filePath As String, targetWB As Workbook)
    If Dir(filePath) = "" Then
        Debug.Print "[ERROR] ThisWorkbook.cls not found"
        Exit Sub
    End If
    
    ' Read file content
    Dim fileNum As Integer
    Dim content As String
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    content = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ' Update ThisWorkbook code
    On Error Resume Next
    With targetWB.VBProject.VBComponents("ThisWorkbook").CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString content
    End With
    
    If Err.Number = 0 Then
        Debug.Print "[SUCCESS] ThisWorkbook updated"
    Else
        Debug.Print "[ERROR] ThisWorkbook update failed: " & Err.Description
    End If
    On Error GoTo 0
End Sub

Private Function TestVBAAccess() As Boolean
    On Error Resume Next
    Dim test As Long
    test = ThisWorkbook.VBProject.VBComponents.Count
    TestVBAAccess = (Err.Number = 0)
    On Error GoTo 0
End Function

' Quick test function
Public Sub QuickTest()
    Debug.Print "=== XLerate Build Environment Test ==="
    Debug.Print "Excel Version: " & Application.Version
    Debug.Print "VBA Access: " & IIf(TestVBAAccess(), "Enabled", "DISABLED")
    Debug.Print "Current Path: " & ThisWorkbook.Path
    Debug.Print "Temp Path: " & Environ("TEMP")
    Debug.Print "User Profile: " & Environ("USERPROFILE")
    Debug.Print "====================================="
End Sub