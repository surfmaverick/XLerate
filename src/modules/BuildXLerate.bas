Option Explicit

Private Const XLERATE_VERSION As String = "2.1.0"
Private Const BUILD_CODENAME As String = "Macabacus Professional"

Public Sub BuildXLerate()
    Debug.Print "=== XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ") Build Started ==="
    
    On Error GoTo BuildError
    
    Dim sourcePath As String
    Dim outputPath As String
    
    sourcePath = "C:\Mac\Home\Documents\Coding\GitHub\XLerate\src\"
    outputPath = "C:\Users\chris\Desktop\XLerate_v2_1_0_Macabacus_Professional.xlam"
    
    Debug.Print "Source: " & sourcePath
    Debug.Print "Output: " & outputPath
    Debug.Print "Platform: Windows + macOS Compatible"
    Debug.Print ""
    
    If Not ValidateEnvironment() Then
        MsgBox "Environment validation failed. Check Immediate Window for details.", vbCritical
        Exit Sub
    End If
    
    If Not ValidateSourceStructure(sourcePath) Then
        MsgBox "Source validation failed. Check Immediate Window for details.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Building XLerate v" & XLERATE_VERSION & "..."
    
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
    
    newAddin.Close False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Debug.Print "=== Build Complete! ==="
    
    MsgBox "XLerate v" & XLERATE_VERSION & " build completed successfully!" & vbNewLine & vbNewLine & _
           "Saved to: " & outputPath & vbNewLine & vbNewLine & _
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
    Debug.Print "Build Error: " & Err.Description & " (Error " & Err.Number & ")"
    MsgBox "Build failed: " & Err.Description & vbNewLine & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Check the Immediate Window (Ctrl+G) for detailed information.", _
           vbCritical, "Build Failed"
    If Not newAddin Is Nothing Then newAddin.Close False
End Sub

Private Function ValidateEnvironment() As Boolean
    Debug.Print "Validating build environment..."
    
    On Error GoTo EnvironmentError
    
    Dim excelVersion As Double
    excelVersion = CDbl(Application.Version)
    If excelVersion < 15.0 Then
        Debug.Print "  ERROR: Excel version too old: " & Application.Version & " (minimum: 15.0)"
        ValidateEnvironment = False
        Exit Function
    End If
    Debug.Print "  SUCCESS: Excel version: " & Application.Version
    
    On Error Resume Next
    Dim testAccess As Long
    testAccess = ThisWorkbook.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Debug.Print "  ERROR: VBA project access denied"
        Debug.Print "    Enable: File -> Options -> Trust Center -> Macro Settings -> Trust access to VBA project object model"
        ValidateEnvironment = False
        On Error GoTo EnvironmentError
        Exit Function
    End If
    On Error GoTo EnvironmentError
    Debug.Print "  SUCCESS: VBA project access enabled"
    
    ValidateEnvironment = True
    Debug.Print "Environment validation passed"
    Exit Function
    
EnvironmentError:
    Debug.Print "Environment validation error: " & Err.Description
    ValidateEnvironment = False
End Function

Private Function ValidateSourceStructure(sourcePath As String) As Boolean
    Debug.Print "Validating source structure: " & sourcePath
    
    If Not FolderExists(sourcePath) Then
        Debug.Print "  ERROR: Source directory not found: " & sourcePath
        ValidateSourceStructure = False
        Exit Function
    End If
    Debug.Print "  SUCCESS: Source directory found"
    
    Dim requiredDirs As Variant
    requiredDirs = Array("modules", "class modules", "objects", "forms")
    
    Dim dir As Variant
    For Each dir In requiredDirs
        If FolderExists(sourcePath & dir & "\") Then
            Debug.Print "  SUCCESS: Found directory: " & dir
        Else
            If dir = "forms" Then
                Debug.Print "  Warning: Optional directory missing: " & dir
            Else
                Debug.Print "  ERROR: Required directory missing: " & dir
                ValidateSourceStructure = False
                Exit Function
            End If
        End If
    Next dir
    
    If FileExists(sourcePath & "objects\ThisWorkbook.cls") Then
        Debug.Print "  SUCCESS: Found ThisWorkbook.cls"
    Else
        Debug.Print "  ERROR: Missing critical file: ThisWorkbook.cls"
        ValidateSourceStructure = False
        Exit Function
    End If
    
    ValidateSourceStructure = True
    Debug.Print "Source structure validation passed"
End Function

Private Function CreateAddinWorkbook() As Workbook
    Set CreateAddinWorkbook = Workbooks.Add
    
    Do While CreateAddinWorkbook.Worksheets.Count > 1
        CreateAddinWorkbook.Worksheets(CreateAddinWorkbook.Worksheets.Count).Delete
    Loop
    
    CreateAddinWorkbook.Worksheets(1).Name = "XLerate_Info"
    With CreateAddinWorkbook.Worksheets(1)
        .Cells(1, 1) = "XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ")"
        .Cells(2, 1) = "Macabacus-Compatible Excel Add-in"
        .Cells(3, 1) = "Built: " & Now
        .Cells(4, 1) = "Platform: Windows + macOS"
        .Cells(5, 1) = ""
        .Cells(6, 1) = "Quick Start Shortcuts:"
        .Cells(7, 1) = "Fast Fill Right: Ctrl+Alt+Shift+R"
        .Cells(8, 1) = "Fast Fill Down: Ctrl+Alt+Shift+D (NEW!)"
        .Cells(9, 1) = "Error Wrap: Ctrl+Alt+Shift+E"
        .Cells(10, 1) = "Pro Precedents: Ctrl+Alt+Shift+["
        .Cells(11, 1) = "Pro Dependents: Ctrl+Alt+Shift+]"
        .Cells(12, 1) = "Number Cycle: Ctrl+Alt+Shift+1"
        .Cells(13, 1) = "Date Cycle: Ctrl+Alt+Shift+2"
        .Cells(14, 1) = "Local Currency: Ctrl+Alt+Shift+3 (NEW!)"
        .Cells(15, 1) = "Foreign Currency: Ctrl+Alt+Shift+4 (NEW!)"
        .Cells(16, 1) = "AutoColor: Ctrl+Alt+Shift+A"
        .Cells(17, 1) = "Settings: Ctrl+Alt+Shift+M"
        .Cells(18, 1) = ""
        .Cells(19, 1) = "100% Macabacus-Compatible Shortcuts"
        .Cells(20, 1) = "Enhanced with Fast Fill Down, Currency Cycling, and Border Utilities"
        
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Color = RGB(0, 100, 200)
        .Range("A6:A20").Font.Color = RGB(0, 120, 0)
        .Cells(19, 1).Font.Bold = True
        .Cells(19, 1).Font.Color = RGB(200, 0, 0)
        .Columns("A:A").AutoFit
    End With
End Function

Private Sub ImportAllModules(sourcePath As String, targetWB As Workbook)
    Debug.Print "Importing all modules..."
    
    Debug.Print "  Importing standard modules..."
    ImportModulesFromFolder sourcePath & "modules\", targetWB, "*.bas"
    
    Debug.Print "  Importing class modules..."
    ImportModulesFromFolder sourcePath & "class modules\", targetWB, "*.cls"
    
    Debug.Print "  Importing forms..."
    ImportModulesFromFolder sourcePath & "forms\", targetWB, "*.frm"
    
    Debug.Print "  Updating ThisWorkbook..."
    UpdateThisWorkbook sourcePath & "objects\ThisWorkbook.cls", targetWB
    
    VerifyCriticalModules targetWB
End Sub

Private Sub ImportModulesFromFolder(folderPath As String, targetWB As Workbook, filePattern As String)
    If Not FolderExists(folderPath) Then
        Debug.Print "    Warning: Folder not found: " & folderPath
        Exit Sub
    End If
    
    Dim fileName As String
    Dim importCount As Long
    Dim errorCount As Long
    
    fileName = Dir(folderPath & filePattern)
    
    Do While fileName <> ""
        Dim filePath As String
        filePath = folderPath & fileName
        
        On Error Resume Next
        targetWB.VBProject.VBComponents.Import filePath
        If Err.Number = 0 Then
            Debug.Print "    SUCCESS: " & fileName
            importCount = importCount + 1
        Else
            Debug.Print "    ERROR: " & fileName & " - " & Err.Description
            errorCount = errorCount + 1
        End If
        On Error GoTo 0
        
        fileName = Dir()
    Loop
    
    Debug.Print "    Summary: " & importCount & " imported, " & errorCount & " errors"
End Sub

Private Sub UpdateThisWorkbook(filePath As String, targetWB As Workbook)
    If Not FileExists(filePath) Then
        Debug.Print "    Warning: ThisWorkbook.cls not found"
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
        Debug.Print "    SUCCESS: ThisWorkbook updated with Macabacus shortcuts"
    Else
        Debug.Print "    ERROR: Could not read ThisWorkbook.cls"
    End If
    On Error GoTo 0
End Sub

Private Sub VerifyCriticalModules(targetWB As Workbook)
    Debug.Print "Verifying critical modules..."
    
    Dim criticalModules As Variant
    criticalModules = Array( _
        "ModNumberFormat", "ModCellFormat", "ModDateFormat", _
        "ModFastFillDown", "ModCurrencyCycling", "ModBorderUtilities", _
        "ModSmartFillRight", "ModErrorWrap", "AutoColorModule", _
        "RibbonCallbacks", "TraceUtils", "FormulaConsistency" _
    )
    
    Dim moduleName As Variant
    For Each moduleName In criticalModules
        On Error Resume Next
        Dim testModule As Object
        Set testModule = targetWB.VBProject.VBComponents(CStr(moduleName))
        If Err.Number = 0 Then
            Debug.Print "    SUCCESS: " & moduleName
        Else
            Debug.Print "    Warning: Missing: " & moduleName
        End If
        On Error GoTo 0
    Next moduleName
End Sub

Private Sub ConfigureMacabacusCompatibility(targetWB As Workbook)
    On Error Resume Next
    
    targetWB.CustomDocumentProperties("Macabacus_Compatible").Delete
    targetWB.CustomDocumentProperties.Add _
        Name:="Macabacus_Compatible", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeBoolean, _
        Value:=True
    
    targetWB.CustomDocumentProperties("Features_FastFillDown").Delete
    targetWB.CustomDocumentProperties.Add _
        Name:="Features_FastFillDown", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeBoolean, _
        Value:=True
    
    targetWB.CustomDocumentProperties("Features_CurrencyCycling").Delete
    targetWB.CustomDocumentProperties.Add _
        Name:="Features_CurrencyCycling", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeBoolean, _
        Value:=True
    
    On Error GoTo 0
    Debug.Print "  SUCCESS: Macabacus compatibility configured"
End Sub

Private Sub SetAddinProperties(targetWB As Workbook)
    On Error Resume Next
    With targetWB
        .Title = "XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ")"
        .Subject = "Macabacus-Compatible Excel Add-in with Enhanced Features"
        .Author = "XLerate Development Team"
        .Comments = "Enhanced with Fast Fill Down, Currency Cycling, Border Utilities, and 100% Macabacus-compatible shortcuts"
        .Keywords = "Excel, Add-in, Financial, Modeling, Macabacus, XLerate, Shortcuts, VBA, Productivity"
        .IsAddin = True
    End With
    
    targetWB.CustomDocumentProperties("Build_Version").Delete
    targetWB.CustomDocumentProperties.Add _
        Name:="Build_Version", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=XLERATE_VERSION
    
    On Error GoTo 0
    Debug.Print "  SUCCESS: Add-in properties configured"
End Sub

Private Sub SaveAddin(targetWB As Workbook, outputPath As String)
    If FileExists(outputPath) Then
        Kill outputPath
        Debug.Print "  SUCCESS: Deleted existing file"
    End If
    
    targetWB.SaveAs outputPath, xlAddIn
    
    If FileExists(outputPath) Then
        Dim fileSize As Long
        fileSize = FileLen(outputPath)
        Debug.Print "  SUCCESS: Add-in saved successfully"
        Debug.Print "  SUCCESS: File size: " & Format(fileSize, "#,##0") & " bytes"
    Else
        Debug.Print "  ERROR: File was not created"
    End If
End Sub

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

Public Sub QuickTest()
    Dim sourcePath As String
    sourcePath = "C:\Mac\Home\Documents\Coding\GitHub\XLerate\src\"
    
    Debug.Print "=== XLerate Quick Test ==="
    Debug.Print "Source exists: " & FolderExists(sourcePath)
    Debug.Print "Modules exists: " & FolderExists(sourcePath & "modules\")
    Debug.Print "Class modules exists: " & FolderExists(sourcePath & "class modules\")
    Debug.Print "Objects exists: " & FolderExists(sourcePath & "objects\")
    Debug.Print "Forms exists: " & FolderExists(sourcePath & "forms\")
    
    Debug.Print "ThisWorkbook.cls exists: " & FileExists(sourcePath & "objects\ThisWorkbook.cls")
    Debug.Print "ModNumberFormat.bas exists: " & FileExists(sourcePath & "modules\ModNumberFormat.bas")
    Debug.Print "ModFastFillDown.bas exists: " & FileExists(sourcePath & "modules\ModFastFillDown.bas")
    Debug.Print "ModCurrencyCycling.bas exists: " & FileExists(sourcePath & "modules\ModCurrencyCycling.bas")
    Debug.Print "ModBorderUtilities.bas exists: " & FileExists(sourcePath & "modules\ModBorderUtilities.bas")
    Debug.Print "RibbonCallbacks.bas exists: " & FileExists(sourcePath & "modules\RibbonCallbacks.bas")
    
    Debug.Print "=== Test Complete ==="
End Sub

Public Sub ValidateCurrentBuild()
    Debug.Print "=== XLerate Installation Validation ==="
    
    Dim requiredModules As Variant
    requiredModules = Array( _
        "ModNumberFormat", "ModCellFormat", "ModDateFormat", "ModTextStyle", _
        "ModFastFillDown", "ModCurrencyCycling", "ModBorderUtilities", _
        "ModSmartFillRight", "ModErrorWrap", "ModSwitchSign", _
        "AutoColorModule", "FormulaConsistency", "TraceUtils", "RibbonCallbacks" _
    )
    
    Dim foundModules As Long
    Dim totalModules As Long
    totalModules = UBound(requiredModules) - LBound(requiredModules) + 1
    
    Dim moduleName As Variant
    For Each moduleName In requiredModules
        On Error Resume Next
        Dim testModule As Object
        Set testModule = ThisWorkbook.VBProject.VBComponents(CStr(moduleName))
        If Err.Number = 0 Then
            Debug.Print "SUCCESS: " & moduleName
            foundModules = foundModules + 1
        Else
            Debug.Print "MISSING: " & moduleName
        End If
        On Error GoTo 0
    Next moduleName
    
    Debug.Print ""
    Debug.Print "Validation Summary:"
    Debug.Print "Found: " & foundModules & "/" & totalModules & " modules"
    Debug.Print "Success Rate: " & Format((foundModules / totalModules) * 100, "0.0") & "%"
    
    If foundModules = totalModules Then
        Debug.Print "SUCCESS: Installation validation PASSED"
        MsgBox "XLerate installation validation passed!" & vbNewLine & _
               "All " & totalModules & " required modules are present.", _
               vbInformation, "Validation Success"
    Else
        Debug.Print "ERROR: Installation validation FAILED"
        MsgBox "XLerate installation validation failed!" & vbNewLine & _
               "Found " & foundModules & " of " & totalModules & " required modules." & vbNewLine & _
               "Run BuildXLerate() to rebuild the add-in.", _
               vbExclamation, "Validation Failed"
    End If
End Sub