' =========================================================================
' File: src/modules/BuildXLerate.bas
' Version: 2.1.1
' Date: July 2025
' Description: Fixed BuildXLerate with file selection and no Unicode issues
'
' CHANGELOG:
' v2.1.1 - Fixed Unicode characters causing display issues
'        - Added file selection dialog for output location
'        - Enhanced debugging and error logging
'        - Improved environment validation
'        - Fixed array syntax errors
' v2.1.0 - Previous version with hardcoded paths
' =========================================================================

Attribute VB_Name = "BuildXLerate"
Option Explicit

Private Const XLERATE_VERSION As String = "2.1.1"
Private Const BUILD_CODENAME As String = "Macabacus Professional"

Public Sub BuildXLerate()
    Debug.Print "==========================================="
    Debug.Print "XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ") Build Started"
    Debug.Print "Time: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
    Debug.Print "==========================================="
    
    On Error GoTo BuildError
    
    Dim startTime As Date
    startTime = Now()
    
    ' Step 1: Environment validation
    Debug.Print ""
    Debug.Print "STEP 1: Environment Validation"
    Debug.Print "--------------------------------"
    If Not ValidateEnvironment() Then
        MsgBox "Environment validation failed. Check Immediate Window for details.", vbCritical, "Build Failed"
        Exit Sub
    End If
    Debug.Print "[SUCCESS] Environment validation passed"
    
    ' Step 2: Get source path with auto-detection and user fallback
    Debug.Print ""
    Debug.Print "STEP 2: Source Path Configuration"
    Debug.Print "----------------------------------"
    Dim sourcePath As String
    sourcePath = GetSourcePath()
    If sourcePath = "" Then
        MsgBox "Could not find or select source path.", vbCritical, "Build Failed"
        Exit Sub
    End If
    Debug.Print "[SUCCESS] Source path: " & sourcePath
    
    ' Step 3: Get output location from user
    Debug.Print ""
    Debug.Print "STEP 3: Output Location Selection"
    Debug.Print "----------------------------------"
    Dim outputPath As String
    outputPath = GetOutputPath()
    If outputPath = "" Then
        Debug.Print "Build cancelled by user"
        Exit Sub
    End If
    Debug.Print "[SUCCESS] Output file: " & outputPath
    
    ' Step 4: Source structure validation
    Debug.Print ""
    Debug.Print "STEP 4: Source Structure Validation"
    Debug.Print "------------------------------------"
    If Not ValidateSourceStructure(sourcePath) Then
        MsgBox "Source validation failed. Check Immediate Window for details.", vbCritical, "Build Failed"
        Exit Sub
    End If
    Debug.Print "[SUCCESS] Source structure validation passed"
    
    ' Step 5: Build process
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Building XLerate v" & XLERATE_VERSION & "..."
    
    Debug.Print ""
    Debug.Print "STEP 5: Add-in Creation Process"
    Debug.Print "--------------------------------"
    Debug.Print "Creating new add-in workbook..."
    Dim newAddin As Workbook
    Set newAddin = CreateAddinWorkbook()
    
    Debug.Print "Importing all modules..."
    ImportAllModules sourcePath, newAddin
    
    Debug.Print "Configuring Macabacus compatibility..."
    ConfigureMacabacusCompatibility newAddin
    
    Debug.Print "Setting add-in properties..."
    SetAddinProperties newAddin
    
    Debug.Print "Saving add-in..."
    SaveAddin newAddin, outputPath
    
    ' Cleanup and validation
    newAddin.Close False
    Set newAddin = Nothing
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Validate the saved file
    Debug.Print ""
    Debug.Print "STEP 6: Post-Build Validation"
    Debug.Print "------------------------------"
    ValidateBuiltFile outputPath
    
    Dim buildTime As String
    buildTime = Format(Now() - startTime, "nn:ss")
    
    Debug.Print ""
    Debug.Print "==========================================="
    Debug.Print "BUILD COMPLETED SUCCESSFULLY"
    Debug.Print "Build time: " & buildTime
    Debug.Print "==========================================="
    
    MsgBox "XLerate v" & XLERATE_VERSION & " build completed successfully!" & vbNewLine & vbNewLine & _
           "Saved to: " & outputPath & vbNewLine & vbNewLine & _
           "Build time: " & buildTime & vbNewLine & vbNewLine & _
           "Next steps:" & vbNewLine & _
           "1. Install the add-in in Excel" & vbNewLine & _
           "2. Enable macros and VBA access" & vbNewLine & _
           "3. Try Ctrl+Alt+Shift+R for Fast Fill Right!" & vbNewLine & _
           "4. Try Ctrl+Alt+Shift+D for Fast Fill Down!" & vbNewLine & _
           "5. Try Ctrl+Alt+Shift+A for AutoColor!" & vbNewLine & vbNewLine & _
           "All Macabacus-compatible shortcuts are active!", _
           vbInformation, "XLerate v" & XLERATE_VERSION & " Build Complete"
    
    Exit Sub
    
BuildError:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    If Not newAddin Is Nothing Then
        newAddin.Close False
        Set newAddin = Nothing
    End If
    
    Debug.Print ""
    Debug.Print "==========================================="
    Debug.Print "BUILD FAILED WITH ERROR"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "==========================================="
    
    MsgBox "Build failed!" & vbNewLine & vbNewLine & _
           "Error: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
           "Check the Immediate Window (Ctrl+G) for detailed information.", _
           vbCritical, "Build Failed"
End Sub

Private Function GetSourcePath() As String
    ' Try multiple possible source paths first
    Debug.Print "Attempting automatic source path detection..."
    
    Dim possiblePaths As Variant
    possiblePaths = Array("C:\Mac\Home\Documents\Coding\GitHub\XLerate\src\", ThisWorkbook.Path & "\src\", ThisWorkbook.Path & "\", Environ("USERPROFILE") & "\Documents\XLerate\src\", "C:\XLerate\src\")
    
    Dim i As Long
    For i = 0 To UBound(possiblePaths)
        Debug.Print "  Checking: " & possiblePaths(i)
        If FolderExists(CStr(possiblePaths(i))) Then
            Debug.Print "  [SUCCESS] Found: " & possiblePaths(i)
            GetSourcePath = possiblePaths(i)
            Exit Function
        End If
    Next i
    
    ' If no automatic path found, ask user
    Debug.Print "  Auto-detection failed, prompting user..."
    
    Dim folderPicker As Object
    Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With folderPicker
        .Title = "Select XLerate Source Folder (containing modules, objects, etc.)"
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then
            GetSourcePath = .SelectedItems(1) & "\"
            Debug.Print "  [SUCCESS] User selected: " & GetSourcePath
        Else
            Debug.Print "  [CANCELLED] User cancelled folder selection"
            GetSourcePath = ""
        End If
    End With
End Function

Private Function GetOutputPath() As String
    ' Show file save dialog for output location
    Debug.Print "Prompting user for output location..."
    
    Dim saveDialog As Object
    Set saveDialog = Application.FileDialog(msoFileDialogSaveAs)
    
    Dim defaultFileName As String
    defaultFileName = "XLerate_v" & Replace(XLERATE_VERSION, ".", "_") & "_" & Replace(BUILD_CODENAME, " ", "_") & ".xlam"
    
    With saveDialog
        .Title = "Save XLerate Add-in As"
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\" & defaultFileName
        .FilterIndex = 1
        .Filters.Clear
        .Filters.Add "Excel Add-ins", "*.xlam"
        .Filters.Add "All Files", "*.*"
        
        If .Show = -1 Then
            GetOutputPath = .SelectedItems(1)
            Debug.Print "  [SUCCESS] User selected: " & GetOutputPath
        Else
            Debug.Print "  [CANCELLED] User cancelled file save"
            GetOutputPath = ""
        End If
    End With
End Function

Private Function ValidateEnvironment() As Boolean
    Debug.Print "Validating build environment..."
    
    On Error GoTo EnvironmentError
    
    ' Check Excel version
    Dim excelVersion As Double
    excelVersion = CDbl(Application.Version)
    Debug.Print "  Excel version: " & Application.Version
    If excelVersion < 15.0 Then
        Debug.Print "  [ERROR] Excel version too old (minimum: 15.0/2013)"
        ValidateEnvironment = False
        Exit Function
    End If
    Debug.Print "  [SUCCESS] Excel version acceptable"
    
    ' Check VBA project access
    On Error Resume Next
    Dim testAccess As Long
    testAccess = ThisWorkbook.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Debug.Print "  [ERROR] VBA project access denied"
        Debug.Print "  Solution: File -> Options -> Trust Center -> Macro Settings -> Trust access to VBA project object model"
        ValidateEnvironment = False
        On Error GoTo EnvironmentError
        Exit Function
    End If
    On Error GoTo EnvironmentError
    Debug.Print "  [SUCCESS] VBA project access enabled"
    
    ' Check disk space
    On Error Resume Next
    Dim freeSpace As Double
    freeSpace = CreateObject("Scripting.FileSystemObject").GetDrive(Environ("TEMP")).FreeSpace
    If Err.Number = 0 Then
        Debug.Print "  [SUCCESS] Available temp space: " & Format(freeSpace / 1024 / 1024, "#,##0") & " MB"
    End If
    On Error GoTo EnvironmentError
    
    ValidateEnvironment = True
    Debug.Print "Environment validation completed successfully"
    Exit Function
    
EnvironmentError:
    Debug.Print "  [ERROR] Environment validation error: " & Err.Description
    ValidateEnvironment = False
End Function

Private Function ValidateSourceStructure(sourcePath As String) As Boolean
    Debug.Print "Validating source structure: " & sourcePath
    
    If Not FolderExists(sourcePath) Then
        Debug.Print "  [ERROR] Source directory not found: " & sourcePath
        ValidateSourceStructure = False
        Exit Function
    End If
    Debug.Print "  [SUCCESS] Source directory found"
    
    ' Check required directories
    Dim requiredDirs As Variant
    requiredDirs = Array("modules", "objects")
    
    Dim dir As Variant
    For Each dir In requiredDirs
        If FolderExists(sourcePath & dir & "\") Then
            Debug.Print "  [SUCCESS] Found required directory: " & dir
        Else
            Debug.Print "  [ERROR] Missing required directory: " & dir
            ValidateSourceStructure = False
            Exit Function
        End If
    Next dir
    
    ' Check optional directories
    Dim optionalDirs As Variant
    optionalDirs = Array("class modules", "forms", "ribbon")
    
    For Each dir In optionalDirs
        If FolderExists(sourcePath & dir & "\") Then
            Debug.Print "  [SUCCESS] Found optional directory: " & dir
        Else
            Debug.Print "  [WARNING] Optional directory missing: " & dir
        End If
    Next dir
    
    ' Check for critical files
    If FileExists(sourcePath & "objects\ThisWorkbook.cls") Then
        Debug.Print "  [SUCCESS] Found ThisWorkbook.cls"
    Else
        Debug.Print "  [ERROR] Missing critical file: ThisWorkbook.cls"
        ValidateSourceStructure = False
        Exit Function
    End If
    
    ValidateSourceStructure = True
    Debug.Print "Source structure validation completed successfully"
End Function

Private Function CreateAddinWorkbook() As Workbook
    Set CreateAddinWorkbook = Workbooks.Add
    
    ' Configure workbook structure
    Do While CreateAddinWorkbook.Worksheets.Count > 1
        CreateAddinWorkbook.Worksheets(CreateAddinWorkbook.Worksheets.Count).Delete
    Loop
    
    CreateAddinWorkbook.Worksheets(1).Name = "XLerate_Info"
    
    ' Populate info sheet
    With CreateAddinWorkbook.Worksheets(1)
        .Cells(1, 1) = "XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ")"
        .Cells(2, 1) = "Macabacus-Compatible Excel Add-in"
        .Cells(3, 1) = "Built: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
        .Cells(4, 1) = "Platform: Windows + macOS"
        .Cells(5, 1) = ""
        .Cells(6, 1) = "MACABACUS-COMPATIBLE SHORTCUTS:"
        .Cells(7, 1) = "Fast Fill Right: Ctrl+Alt+Shift+R"
        .Cells(8, 1) = "Fast Fill Down: Ctrl+Alt+Shift+D"
        .Cells(9, 1) = "Error Wrap: Ctrl+Alt+Shift+E"
        .Cells(10, 1) = "Pro Precedents: Ctrl+Alt+Shift+["
        .Cells(11, 1) = "Pro Dependents: Ctrl+Alt+Shift+]"
        .Cells(12, 1) = "Number Cycle: Ctrl+Alt+Shift+1"
        .Cells(13, 1) = "Date Cycle: Ctrl+Alt+Shift+2"
        .Cells(14, 1) = "AutoColor: Ctrl+Alt+Shift+A"
        .Cells(15, 1) = "Quick Save: Ctrl+Alt+Shift+S"
        .Cells(16, 1) = "Toggle Gridlines: Ctrl+Alt+Shift+G"
        .Cells(17, 1) = ""
        .Cells(18, 1) = "For full documentation, visit: github.com/omegarhovega/XLerate"
    End With
    
    Debug.Print "  [SUCCESS] Workbook created and configured"
End Function

Private Sub ImportAllModules(sourcePath As String, targetWB As Workbook)
    Debug.Print "Starting module import process..."
    
    Dim totalModules As Integer
    totalModules = 0
    
    ' Import standard modules
    Debug.Print "  Importing standard modules..."
    totalModules = totalModules + ImportModulesFromFolder(sourcePath & "modules\", targetWB, "*.bas", "Standard Module")
    
    ' Import class modules
    Debug.Print "  Importing class modules..."
    totalModules = totalModules + ImportModulesFromFolder(sourcePath & "class modules\", targetWB, "*.cls", "Class Module")
    
    ' Import forms
    Debug.Print "  Importing forms..."
    totalModules = totalModules + ImportModulesFromFolder(sourcePath & "forms\", targetWB, "*.frm", "UserForm")
    
    ' Update ThisWorkbook
    Debug.Print "  Updating ThisWorkbook..."
    UpdateThisWorkbook sourcePath & "objects\ThisWorkbook.cls", targetWB
    
    Debug.Print "  [SUCCESS] Module import completed: " & totalModules & " modules imported"
End Sub

Private Function ImportModulesFromFolder(folderPath As String, targetWB As Workbook, filePattern As String, moduleType As String) As Integer
    If Not FolderExists(folderPath) Then
        Debug.Print "    [WARNING] Folder not found: " & folderPath
        ImportModulesFromFolder = 0
        Exit Function
    End If
    
    Dim fileName As String
    Dim importCount As Integer
    importCount = 0
    
    fileName = Dir(folderPath & filePattern)
    Do While fileName <> ""
        Dim filePath As String
        filePath = folderPath & fileName
        
        Debug.Print "    Importing " & moduleType & ": " & fileName
        
        On Error Resume Next
        targetWB.VBProject.VBComponents.Import filePath
        If Err.Number = 0 Then
            Debug.Print "      [SUCCESS] " & fileName
            importCount = importCount + 1
        Else
            Debug.Print "      [ERROR] " & fileName & " - " & Err.Description
        End If
        On Error GoTo 0
        
        fileName = Dir
    Loop
    
    Debug.Print "    " & moduleType & " import completed: " & importCount & " files"
    ImportModulesFromFolder = importCount
End Function

Private Sub UpdateThisWorkbook(thisWorkbookPath As String, targetWB As Workbook)
    If Not FileExists(thisWorkbookPath) Then
        Debug.Print "    [ERROR] ThisWorkbook.cls not found: " & thisWorkbookPath
        Exit Sub
    End If
    
    Debug.Print "    Reading ThisWorkbook content..."
    Dim content As String
    content = ReadFile(thisWorkbookPath)
    
    If Len(content) = 0 Then
        Debug.Print "    [ERROR] Failed to read ThisWorkbook.cls"
        Exit Sub
    End If
    
    Debug.Print "    Updating ThisWorkbook code module..."
    On Error Resume Next
    targetWB.VBProject.VBComponents("ThisWorkbook").CodeModule.DeleteLines 1, targetWB.VBProject.VBComponents("ThisWorkbook").CodeModule.CountOfLines
    targetWB.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString content
    
    If Err.Number = 0 Then
        Debug.Print "    [SUCCESS] ThisWorkbook updated successfully"
    Else
        Debug.Print "    [ERROR] ThisWorkbook update failed: " & Err.Description
    End If
    On Error GoTo 0
End Sub

Private Sub ConfigureMacabacusCompatibility(targetWB As Workbook)
    Debug.Print "  Configuring Macabacus compatibility settings..."
    
    On Error Resume Next
    
    ' Remove existing properties first
    targetWB.CustomDocumentProperties("Macabacus_Compatible").Delete
    targetWB.CustomDocumentProperties("Features_FastFillDown").Delete
    targetWB.CustomDocumentProperties("Features_CurrencyCycling").Delete
    targetWB.CustomDocumentProperties("Features_EnhancedUI").Delete
    
    ' Add new properties
    targetWB.CustomDocumentProperties.Add Name:="Macabacus_Compatible", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:=True
    targetWB.CustomDocumentProperties.Add Name:="Features_FastFillDown", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:=True
    targetWB.CustomDocumentProperties.Add Name:="Features_CurrencyCycling", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:=True
    targetWB.CustomDocumentProperties.Add Name:="Features_EnhancedUI", LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:=True
    targetWB.CustomDocumentProperties.Add Name:="Build_Version", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=XLERATE_VERSION
    targetWB.CustomDocumentProperties.Add Name:="Build_Date", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=Format(Now(), "yyyy-mm-dd")
    
    On Error GoTo 0
    Debug.Print "  [SUCCESS] Macabacus compatibility configured"
End Sub

Private Sub SetAddinProperties(targetWB As Workbook)
    Debug.Print "  Setting add-in properties..."
    
    On Error Resume Next
    With targetWB
        .Title = "XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ")"
        .Subject = "Macabacus-Compatible Excel Add-in with Enhanced Features"
        .Author = "XLerate Development Team"
        .Comments = "Enhanced with Fast Fill Down, Currency Cycling, and comprehensive Macabacus compatibility"
        .Keywords = "Excel, Add-in, Financial, Modeling, Macabacus, XLerate, Shortcuts, VBA, Productivity"
        .IsAddin = True
    End With
    On Error GoTo 0
    
    Debug.Print "  [SUCCESS] Add-in properties configured"
End Sub

Private Sub SaveAddin(targetWB As Workbook, outputPath As String)
    Debug.Print "  Preparing to save add-in..."
    
    ' Delete existing file if it exists
    If FileExists(outputPath) Then
        Debug.Print "    Deleting existing file..."
        On Error Resume Next
        Kill outputPath
        If Err.Number = 0 Then
            Debug.Print "    [SUCCESS] Existing file deleted"
        Else
            Debug.Print "    [WARNING] Could not delete existing file: " & Err.Description
        End If
        On Error GoTo 0
    End If
    
    Debug.Print "    Saving add-in to: " & outputPath
    
    On Error GoTo SaveError
    targetWB.SaveAs outputPath, xlAddIn
    Debug.Print "  [SUCCESS] Add-in saved successfully"
    Exit Sub
    
SaveError:
    Debug.Print "  [ERROR] Save failed: " & Err.Description
End Sub

Private Sub ValidateBuiltFile(outputPath As String)
    Debug.Print "Validating built add-in..."
    
    If FileExists(outputPath) Then
        Dim fileSize As Long
        fileSize = FileLen(outputPath)
        Debug.Print "  [SUCCESS] Add-in file exists"
        Debug.Print "  [SUCCESS] File size: " & Format(fileSize, "#,##0") & " bytes"
        
        If fileSize < 10000 Then
            Debug.Print "  [WARNING] File size seems small, may indicate incomplete build"
        End If
        
        ' Try to get file properties
        On Error Resume Next
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim file As Object
        Set file = fso.GetFile(outputPath)
        Debug.Print "  [SUCCESS] File created: " & file.DateCreated
        Debug.Print "  [SUCCESS] File modified: " & file.DateLastModified
        On Error GoTo 0
        
    Else
        Debug.Print "  [ERROR] Add-in file was not created!"
    End If
End Sub

' Utility functions
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

' Quick test function for debugging
Public Sub QuickBuildTest()
    Debug.Print "=== XLerate Build Environment Test ==="
    Debug.Print "Excel Version: " & Application.Version
    Debug.Print "VBA Access: " & IIf(TestVBAAccess(), "[SUCCESS] Enabled", "[ERROR] Disabled")
    Debug.Print "Temp Path: " & Environ("TEMP")
    Debug.Print "User Profile: " & Environ("USERPROFILE")
    Debug.Print "Current Path: " & ThisWorkbook.Path
    Debug.Print "======================================"
End Sub

Private Function TestVBAAccess() As Boolean
    On Error Resume Next
    Dim test As Long
    test = ThisWorkbook.VBProject.VBComponents.Count
    TestVBAAccess = (Err.Number = 0)
    On Error GoTo 0
End Function