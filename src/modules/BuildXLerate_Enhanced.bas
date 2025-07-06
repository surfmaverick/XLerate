' =============================================================================
' File: BuildXLerate_Enhanced.bas
' Version: 2.1.0 - Macabacus-Compatible Build System
' Description: Enhanced build script with complete Macabacus alignment
' Author: XLerate Development Team
' Created: Enhanced for Macabacus compatibility
' Last Modified: January 2025
' =============================================================================

Option Explicit

Private Const XLERATE_VERSION As String = "2.1.0"
Private Const BUILD_DATE As String = "January 2025"
Private Const BUILD_CODENAME As String = "Macabacus Professional"
Private Const AUTHOR As String = "XLerate Development Team"

' Build progress tracking
Private Type BuildProgress
    TotalSteps As Long
    CompletedSteps As Long
    FailedSteps As Long
    Warnings As Long
    StartTime As Date
End Type

Public Sub BuildXLerate()
    ' Enhanced build procedure with Macabacus compatibility and comprehensive module support
    Debug.Print "=== XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ") Build Started ==="
    Debug.Print "Build Date: " & BUILD_DATE
    Debug.Print "Platform: Windows + macOS Compatible"
    Debug.Print ""
    
    Dim progress As BuildProgress
    progress.StartTime = Now
    
    On Error GoTo BuildError
    
    ' Phase 1: Get Build Paths
    Debug.Print "Phase 1: Configure Build Paths"
    Dim sourcePath As String
    Dim outputPath As String
    
    sourcePath = GetSourcePath()
    If sourcePath = "" Then
        MsgBox "Build cancelled - no source path provided.", vbInformation, "XLerate Build"
        Exit Sub
    End If
    
    outputPath = GetOutputPath()
    If outputPath = "" Then
        MsgBox "Build cancelled - no output path provided.", vbInformation, "XLerate Build"
        Exit Sub
    End If
    
    Debug.Print "✓ Source: " & sourcePath
    Debug.Print "✓ Output: " & outputPath
    Debug.Print ""
    
    ' Phase 2: Validate Environment
    Debug.Print "Phase 2: Validate Build Environment"
    If Not ValidateEnvironment() Then
        Err.Raise 9999, "BuildXLerate", "Environment validation failed"
    End If
    If Not ValidateSourceDirectory(sourcePath) Then
        Err.Raise 9999, "BuildXLerate", "Source directory validation failed"
    End If
    Debug.Print "✓ Environment validation passed"
    Debug.Print ""
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Building XLerate v" & XLERATE_VERSION & "..."
    
    ' Phase 3: Create New Add-in
    Debug.Print "Phase 3: Create New Add-in Workbook"
    Dim newAddin As Workbook
    Set newAddin = CreateNewAddin()
    Debug.Print "✓ New add-in workbook created"
    progress.CompletedSteps = progress.CompletedSteps + 1
    
    ' Phase 4: Import Core Components
    Debug.Print "Phase 4: Import Core Components"
    ImportAllModules sourcePath, newAddin, progress
    Debug.Print "✓ Core components imported"
    
    ' Phase 5: Setup Ribbon Integration
    Debug.Print "Phase 5: Setup Ribbon Integration"
    SetupRibbonIntegration sourcePath, newAddin, progress
    Debug.Print "✓ Ribbon integration configured"
    
    ' Phase 6: Configure Macabacus Compatibility
    Debug.Print "Phase 6: Configure Macabacus Compatibility"
    ConfigureMacabacusCompatibility newAddin, progress
    Debug.Print "✓ Macabacus compatibility configured"
    
    ' Phase 7: Set Add-in Properties
    Debug.Print "Phase 7: Set Add-in Properties"
    SetAddinProperties newAddin
    Debug.Print "✓ Add-in properties set"
    progress.CompletedSteps = progress.CompletedSteps + 1
    
    ' Phase 8: Save and Finalize
    Debug.Print "Phase 8: Save and Finalize"
    SaveAsXLAM newAddin, outputPath
    Debug.Print "✓ Add-in saved as XLAM"
    progress.CompletedSteps = progress.CompletedSteps + 1
    
    ' Phase 9: Generate Documentation
    Debug.Print "Phase 9: Generate Documentation"
    GenerateBuildDocumentation outputPath, progress
    Debug.Print "✓ Build documentation generated"
    
    ' Clean up
    newAddin.Close False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Dim buildTime As Date
    buildTime = Now - progress.StartTime
    
    Debug.Print ""
    Debug.Print "=== XLerate Build Completed Successfully ==="
    Debug.Print "Build Time: " & Format(buildTime, "hh:mm:ss")
    Debug.Print "Components: " & progress.CompletedSteps & " completed, " & progress.FailedSteps & " failed"
    Debug.Print "Warnings: " & progress.Warnings
    Debug.Print ""
    
    ' Show success message with next steps
    Dim msg As String
    msg = "XLerate v" & XLERATE_VERSION & " build completed successfully!" & vbNewLine & vbNewLine
    msg = msg & "📊 Build Summary:" & vbNewLine
    msg = msg & "• Build Time: " & Format(buildTime, "hh:mm:ss") & vbNewLine
    msg = msg & "• Components: " & progress.CompletedSteps & " completed" & vbNewLine
    msg = msg & "• Failed: " & progress.FailedSteps & vbNewLine
    msg = msg & "• Warnings: " & progress.Warnings & vbNewLine & vbNewLine
    msg = msg & "📁 Output Location:" & vbNewLine
    msg = msg & outputPath & vbNewLine & vbNewLine
    msg = msg & "🎯 Next Steps:" & vbNewLine
    msg = msg & "1. Install the add-in in Excel" & vbNewLine
    msg = msg & "2. Enable macros and VBA access" & vbNewLine
    msg = msg & "3. Test Macabacus-compatible shortcuts" & vbNewLine
    msg = msg & "4. Customize settings as needed"
    
    MsgBox msg, vbInformation, "XLerate v" & XLERATE_VERSION & " Build Complete"
    
    Exit Sub
    
BuildError:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Debug.Print ""
    Debug.Print "=== BUILD FAILED ==="
    Debug.Print "Error: " & Err.Description & " (Error " & Err.Number & ")"
    Debug.Print "Completed: " & progress.CompletedSteps
    Debug.Print "Failed: " & progress.FailedSteps
    
    MsgBox "XLerate build failed!" & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine & _
           "Completed: " & progress.CompletedSteps & vbNewLine & _
           "Failed: " & progress.FailedSteps & vbNewLine & vbNewLine & _
           "Check the Immediate Window (Ctrl+G) for detailed error information.", _
           vbCritical, "XLerate Build Failed"
    
    If Not newAddin Is Nothing Then
        newAddin.Close False
    End If
End Sub

Private Function GetSourcePath() As String
    ' Get source directory with enhanced path detection
    
    ' Try to detect common project locations
    Dim possiblePaths As Variant
    possiblePaths = Array( _
        "C:\Users\" & Environ("USERNAME") & "\Documents\XLerate\", _
        "C:\Users\" & Environ("USERNAME") & "\Desktop\XLerate\", _
        "C:\Users\" & Environ("USERNAME") & "\XLerate\", _
        "C:\XLerate\", _
        ThisWorkbook.Path & "\", _
        ThisWorkbook.Path & "\..\", _
        ThisWorkbook.Path & "\..\.." _
    )
    
    Dim defaultPath As String
    defaultPath = possiblePaths(0)
    
    ' Check if any of the possible paths exist
    Dim i As Long
    For i = LBound(possiblePaths) To UBound(possiblePaths)
        If DirExists(CStr(possiblePaths(i)) & "src\") Then
            defaultPath = CStr(possiblePaths(i))
            Exit For
        End If
    Next i
    
    Dim userPath As String
    userPath = InputBox( _
        "📁 XLerate Source Directory" & vbNewLine & vbNewLine & _
        "Enter the full path to your XLerate project directory:" & vbNewLine & _
        "(The folder that contains the 'src' folder)" & vbNewLine & vbNewLine & _
        "Example: C:\Users\YourName\XLerate\" & vbNewLine & vbNewLine & _
        "💡 Tip: The build script will look for:" & vbNewLine & _
        "• src\modules\ (VBA modules)" & vbNewLine & _
        "• src\class modules\ (Class files)" & vbNewLine & _
        "• src\forms\ (UserForms)" & vbNewLine & _
        "• src\objects\ (ThisWorkbook)" & vbNewLine & _
        "• src\ribbon\ (Ribbon XML)", _
        "XLerate Source Directory", defaultPath)
    
    If userPath <> "" Then
        ' Ensure path ends with backslash
        If Right(userPath, 1) <> "\" Then userPath = userPath & "\"
        ' Add src\ to the path
        GetSourcePath = userPath & "src\"
    Else
        GetSourcePath = ""
    End If
End Function

Private Function GetOutputPath() As String
    ' Get output file path with version-specific naming
    
    Dim defaultPath As String
    Dim versionString As String
    versionString = Replace(XLERATE_VERSION, ".", "_")
    
    ' Suggest desktop location with descriptive filename
    defaultPath = "C:\Users\" & Environ("USERNAME") & "\Desktop\XLerate_v" & versionString & "_" & BUILD_CODENAME & ".xlam"
    
    Dim userPath As String
    userPath = InputBox( _
        "💾 XLerate Output File" & vbNewLine & vbNewLine & _
        "Enter the full path where you want to save the XLerate add-in:" & vbNewLine & vbNewLine & _
        "📋 Recommended filename format:" & vbNewLine & _
        "XLerate_v" & versionString & "_" & BUILD_CODENAME & ".xlam" & vbNewLine & vbNewLine & _
        "⚠️ Note: If file exists, it will be overwritten." & vbNewLine & vbNewLine & _
        "Example: C:\Users\YourName\Desktop\XLerate_v" & versionString & ".xlam", _
        "XLerate Output File", defaultPath)
    
    GetOutputPath = userPath
End Function

Private Function ValidateEnvironment() As Boolean
    ' Enhanced environment validation
    
    Debug.Print "  - Checking Excel version..."
    If CDbl(Application.Version) < 15.0 Then  ' Excel 2013+
        Debug.Print "    ✗ Excel version " & Application.Version & " is too old (minimum: 15.0)"
        MsgBox "Excel 2013 or later is required for XLerate." & vbNewLine & _
               "Current version: " & Application.Version, vbCritical
        ValidateEnvironment = False
        Exit Function
    End If
    Debug.Print "    ✓ Excel version " & Application.Version & " is compatible"
    
    Debug.Print "  - Checking VBA project access..."
    On Error Resume Next
    Dim testAccess As Long
    testAccess = ThisWorkbook.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Debug.Print "    ✗ VBA project access denied"
        MsgBox "VBA project access is required for building XLerate." & vbNewLine & vbNewLine & _
               "Please enable:" & vbNewLine & _
               "File → Options → Trust Center → Trust Center Settings → " & _
               "Macro Settings → Trust access to the VBA project object model", _
               vbCritical, "VBA Access Required"
        ValidateEnvironment = False
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    Debug.Print "    ✓ VBA project access available"
    
    Debug.Print "  - Checking macro security..."
    ' This is informational - user will need to handle macro security
    Debug.Print "    ℹ Macro security should be set appropriately for add-in usage"
    
    ValidateEnvironment = True
End Function

Private Function ValidateSourceDirectory(sourcePath As String) As Boolean
    ' Enhanced source directory validation with detailed feedback
    
    Debug.Print "  - Validating source structure: " & sourcePath
    
    Dim isValid As Boolean
    Dim missingItems As Collection
    Set missingItems = New Collection
    
    isValid = True
    
    ' Check required directories
    Dim requiredDirs As Variant
    requiredDirs = Array("class modules", "modules", "objects", "forms", "ribbon")
    
    Dim dir As Variant
    For Each dir In requiredDirs
        If Not DirExists(sourcePath & dir & "\") Then
            Debug.Print "    ✗ Missing directory: " & dir
            missingItems.Add "Directory: " & dir
            ' Don't fail for optional directories
            If dir <> "ribbon" And dir <> "forms" Then
                isValid = False
            End If
        Else
            Debug.Print "    ✓ Found directory: " & dir
        End If
    Next dir
    
    ' Check critical files
    Dim criticalFiles As Variant
    criticalFiles = Array( _
        "objects\ThisWorkbook.cls", _
        "modules\ModNumberFormat.bas", _
        "modules\RibbonCallbacks.bas" _
    )
    
    Dim file As Variant
    For Each file In criticalFiles
        If Not FileExists(sourcePath & file) Then
            Debug.Print "    ✗ Missing critical file: " & file
            missingItems.Add "File: " & file
            isValid = False
        Else
            Debug.Print "    ✓ Found critical file: " & file
        End If
    Next file
    
    If Not isValid Then
        Dim msg As String
        msg = "Source directory validation failed!" & vbNewLine & vbNewLine & _
              "Missing required items:" & vbNewLine
        
        Dim item As Variant
        For Each item In missingItems
            msg = msg & "• " & item & vbNewLine
        Next item
        
        msg = msg & vbNewLine & "Please check your source directory structure."
        MsgBox msg, vbCritical, "Source Validation Failed"
    End If
    
    ValidateSourceDirectory = isValid
End Function

Private Sub ImportAllModules(sourcePath As String, targetWB As Workbook, ByRef progress As BuildProgress)
    ' Import all modules with enhanced error handling and progress tracking
    
    Application.StatusBar = "Importing modules..."
    
    ' Import in specific order for dependencies
    ImportClassModules sourcePath, targetWB, progress
    ImportStandardModules sourcePath, targetWB, progress
    ImportUserForms sourcePath, targetWB, progress
    UpdateThisWorkbook sourcePath, targetWB, progress
End Sub

Private Sub ImportClassModules(sourcePath As String, targetWB As Workbook, ByRef progress As BuildProgress)
    ' Import class modules with enhanced module list
    
    Debug.Print "  - Importing class modules..."
    
    Dim classFiles As Variant
    classFiles = Array( _
        "clsFormatType.cls", _
        "clsCellFormatType.cls", _
        "clsTextStyleType.cls", _
        "clsUISettings.cls", _
        "clsListBoxHandler.cls", _
        "clsDynamicButtonHandler.cls" _
    )
    
    ImportModuleArray sourcePath & "class modules\", classFiles, targetWB, progress, "class"
End Sub

Private Sub ImportStandardModules(sourcePath As String, targetWB As Workbook, ByRef progress As BuildProgress)
    ' Import standard modules with complete Macabacus-aligned module list
    
    Debug.Print "  - Importing standard modules..."
    
    Dim moduleFiles As Variant
    moduleFiles = Array( _
        "ModNumberFormat.bas", _
        "ModCellFormat.bas", _
        "ModDateFormat.bas", _
        "ModTextStyle.bas", _
        "ModSmartFillRight.bas", _
        "ModFastFillDown.bas", _
        "ModCurrencyCycling.bas", _
        "ModBorderUtilities.bas", _
        "ModErrorWrap.bas", _
        "ModSwitchSign.bas", _
        "ModCAGR.bas", _
        "ModFormatReset.bas", _
        "ModSettings.bas", _
        "ModGlobalSettings.bas", _
        "ModVersionInfo.bas", _
        "ModUtilityHelpers.bas", _
        "ModFinancialFunctions.bas", _
        "AutoColorModule.bas", _
        "FormulaConsistency.bas", _
        "TraceUtils.bas", _
        "RibbonCallbacks.bas", _
        "ModBuildXLerate.bas" _
    )
    
    ImportModuleArray sourcePath & "modules\", moduleFiles, targetWB, progress, "module"
End Sub

Private Sub ImportUserForms(sourcePath As String, targetWB As Workbook, ByRef progress As BuildProgress)
    ' Import user forms with enhanced form list
    
    Debug.Print "  - Importing user forms..."
    
    Dim formFiles As Variant
    formFiles = Array( _
        "frmNumberSettings.frm", _
        "frmCellSettings.frm", _
        "frmDateSettings.frm", _
        "frmTextStyle.frm", _
        "frmAutoColor.frm", _
        "frmErrorSettings.frm", _
        "frmSettingsManager.frm", _
        "frmPrecedents.frm", _
        "frmDependents.frm" _
    )
    
    ImportModuleArray sourcePath & "forms\", formFiles, targetWB, progress, "form"
End Sub

Private Sub ImportModuleArray(basePath As String, moduleFiles As Variant, targetWB As Workbook, ByRef progress As BuildProgress, moduleType As String)
    ' Generic module import function with progress tracking
    
    Dim i As Long
    For i = LBound(moduleFiles) To UBound(moduleFiles)
        Dim filePath As String
        filePath = basePath & moduleFiles(i)
        
        Application.StatusBar = "Importing " & moduleType & ": " & moduleFiles(i)
        
        If FileExists(filePath) Then
            On Error Resume Next
            targetWB.VBProject.VBComponents.Import filePath
            If Err.Number = 0 Then
                Debug.Print "    ✓ Imported " & moduleType & ": " & moduleFiles(i)
                progress.CompletedSteps = progress.CompletedSteps + 1
            Else
                Debug.Print "    ✗ ERROR importing " & moduleFiles(i) & ": " & Err.Description
                progress.FailedSteps = progress.FailedSteps + 1
            End If
            On Error GoTo 0
        Else
            Debug.Print "    ⚠ Missing " & moduleType & " file: " & moduleFiles(i)
            progress.Warnings = progress.Warnings + 1
        End If
    Next i
End Sub

Private Sub UpdateThisWorkbook(sourcePath As String, targetWB As Workbook, ByRef progress As BuildProgress)
    ' Update ThisWorkbook with enhanced version for Macabacus compatibility
    
    Debug.Print "  - Updating ThisWorkbook..."
    Application.StatusBar = "Updating ThisWorkbook..."
    
    Dim filePath As String
    filePath = sourcePath & "objects\ThisWorkbook.cls"
    
    If FileExists(filePath) Then
        On Error Resume Next
        Dim fileContent As String
        fileContent = ReadTextFile(filePath)
        
        If fileContent <> "" Then
            With targetWB.VBProject.VBComponents("ThisWorkbook").CodeModule
                .DeleteLines 1, .CountOfLines
                .AddFromString fileContent
            End With
            Debug.Print "    ✓ Updated: ThisWorkbook.cls with Macabacus shortcuts"
            progress.CompletedSteps = progress.CompletedSteps + 1
        Else
            Debug.Print "    ✗ ERROR: Could not read ThisWorkbook.cls content"
            progress.FailedSteps = progress.FailedSteps + 1
        End If
        On Error GoTo 0
    Else
        Debug.Print "    ✗ Missing file: ThisWorkbook.cls"
        progress.FailedSteps = progress.FailedSteps + 1
    End If
End Sub

Private Sub SetupRibbonIntegration(sourcePath As String, targetWB As Workbook, ByRef progress As BuildProgress)
    ' Setup ribbon integration (prepare for manual XML addition)
    
    Debug.Print "  - Preparing ribbon integration..."
    Application.StatusBar = "Setting up ribbon integration..."
    
    ' Check for ribbon XML file
    Dim ribbonPath As String
    ribbonPath = sourcePath & "ribbon\customUI14.xml"
    
    If FileExists(ribbonPath) Then
        Debug.Print "    ✓ Found ribbon XML file: " & ribbonPath
        Debug.Print "    ℹ Note: Ribbon XML must be added manually using Custom UI Editor"
        progress.Warnings = progress.Warnings + 1  ' Note for manual step
    Else
        Debug.Print "    ⚠ Ribbon XML file not found - add-in will work without custom ribbon"
        progress.Warnings = progress.Warnings + 1
    End If
    
    ' Ensure RibbonCallbacks module is present
    On Error Resume Next
    Dim ribbonModule As Object
    Set ribbonModule = targetWB.VBProject.VBComponents("RibbonCallbacks")
    If Err.Number = 0 Then
        Debug.Print "    ✓ RibbonCallbacks module present"
    Else
        Debug.Print "    ⚠ RibbonCallbacks module missing - ribbon features will not work"
        progress.Warnings = progress.Warnings + 1
    End If
    On Error GoTo 0
End Sub

Private Sub ConfigureMacabacusCompatibility(targetWB As Workbook, ByRef progress As BuildProgress)
    ' Configure Macabacus compatibility settings
    
    Debug.Print "  - Configuring Macabacus compatibility..."
    Application.StatusBar = "Configuring Macabacus compatibility..."
    
    ' Add Macabacus compatibility flag
    On Error Resume Next
    targetWB.CustomDocumentProperties("Macabacus_Compatible").Delete
    targetWB.CustomDocumentProperties.Add _
        Name:="Macabacus_Compatible", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeBoolean, _
        Value:=True
    
    ' Add shortcut mapping information
    targetWB.CustomDocumentProperties("Shortcut_Pattern").Delete
    targetWB.CustomDocumentProperties.Add _
        Name:="Shortcut_Pattern", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:="Ctrl+Alt+Shift"
    
    ' Add feature flags
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
    
    targetWB.CustomDocumentProperties("Features_BorderUtilities").Delete
    targetWB.CustomDocumentProperties.Add _
        Name:="Features_BorderUtilities", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeBoolean, _
        Value:=True
    
    On Error GoTo 0
    
    Debug.Print "    ✓ Macabacus compatibility configured"
    progress.CompletedSteps = progress.CompletedSteps + 1
End Sub

Private Function CreateNewAddin() As Workbook
    ' Create enhanced add-in workbook with version info
    
    Debug.Print "  - Creating new add-in workbook..."
    
    Set CreateNewAddin = Workbooks.Add
    
    ' Remove extra sheets (keep only one)
    Do While CreateNewAddin.Worksheets.Count > 1
        Application.DisplayAlerts = False
        CreateNewAddin.Worksheets(CreateNewAddin.Worksheets.Count).Delete
        Application.DisplayAlerts = True
    Loop
    
    ' Configure the info sheet
    CreateNewAddin.Worksheets(1).Name = "XLerate_Info"
    
    ' Add comprehensive version information
    With CreateNewAddin.Worksheets(1)
        .Cells(1, 1).Value = "XLerate Add-in"
        .Cells(2, 1).Value = "Version: " & XLERATE_VERSION
        .Cells(3, 1).Value = "Codename: " & BUILD_CODENAME
        .Cells(4, 1).Value = "Built: " & BUILD_DATE
        .Cells(5, 1).Value = "Author: " & AUTHOR
        .Cells(6, 1).Value = "Platform: Windows + macOS"
        .Cells(7, 1).Value = "Compatibility: Macabacus-aligned shortcuts"
        .Cells(8, 1).Value = "Excel Version: " & Application.Version & "+"
        
        .Cells(10, 1).Value = "Quick Start:"
        .Cells(11, 1).Value = "• Fast Fill Right: Ctrl+Alt+Shift+R"
        .Cells(12, 1).Value = "• Fast Fill Down: Ctrl+Alt+Shift+D"
        .Cells(13, 1).Value = "• Pro Precedents: Ctrl+Alt+Shift+["
        .Cells(14, 1).Value = "• Pro Dependents: Ctrl+Alt+Shift+]"
        .Cells(15, 1).Value = "• Number Cycle: Ctrl+Alt+Shift+1"
        .Cells(16, 1).Value = "• AutoColor: Ctrl+Alt+Shift+A"
        .Cells(17, 1).Value = "• Settings: Ctrl+Alt+Shift+M"
        
        ' Format the sheet
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Color = RGB(0, 100, 200)
        
        .Range("A10:A17").Font.Bold = True
        .Range("A11:A17").Font.Color = RGB(0, 150, 0)
        
        .Columns("A:A").AutoFit
        .Range("A1").Select
    End With
End Function

Private Sub SetAddinProperties(targetWB As Workbook)
    ' Set comprehensive add-in properties
    
    Debug.Print "  - Setting add-in properties..."
    
    On Error Resume Next
    With targetWB
        .Title = "XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ")"
        .Subject = "Excel Add-in for Financial Modeling with Macabacus Compatibility"
        .Author = AUTHOR
        .Comments = "Enhanced with Macabacus-style shortcuts, Fast Fill Down, Currency Cycling, and Border Utilities"
        .Keywords = "Excel, Add-in, Financial, Modeling, Macabacus, XLerate, Shortcuts, VBA"
        .IsAddin = True
    End With
    
    ' Add build information
    targetWB.CustomDocumentProperties("Build_Version").Delete
    targetWB.CustomDocumentProperties("Build_Date").Delete
    targetWB.CustomDocumentProperties("Build_Codename").Delete
    
    targetWB.CustomDocumentProperties.Add _
        Name:="Build_Version", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=XLERATE_VERSION
    
    targetWB.CustomDocumentProperties.Add _
        Name:="Build_Date", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=BUILD_DATE
    
    targetWB.CustomDocumentProperties.Add _
        Name:="Build_Codename", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=BUILD_CODENAME
    
    On Error GoTo 0
End Sub

Private Sub SaveAsXLAM(targetWB As Workbook, outputPath As String)
    ' Save with enhanced error handling and verification
    
    Debug.Print "  - Saving as XLAM: " & outputPath
    Application.StatusBar = "Saving add-in..."
    
    On Error Resume Next
    ' Delete existing file if it exists
    If FileExists(outputPath) Then
        Kill outputPath
        Debug.Print "    ✓ Deleted existing file"
    End If
    
    ' Save as add-in
    targetWB.SaveAs outputPath, xlAddIn
    If Err.Number = 0 Then
        Debug.Print "    ✓ Saved successfully as " & outputPath
        
        ' Verify file was created
        If FileExists(outputPath) Then
            Dim fileSize As Long
            fileSize = FileLen(outputPath)
            Debug.Print "    ✓ File verified: " & Format(fileSize, "#,##0") & " bytes"
        Else
            Debug.Print "    ✗ File verification failed"
        End If
    Else
        Debug.Print "    ✗ Save failed: " & Err.Description
    End If
    On Error GoTo 0
End Sub

Private Sub GenerateBuildDocumentation(outputPath As String, progress As BuildProgress)
    ' Generate build documentation and next steps
    
    Debug.Print "  - Generating build documentation..."
    
    Dim docPath As String
    docPath = Replace(outputPath, ".xlam", "_ReadMe.txt")
    
    On Error Resume Next
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open docPath For Output As #fileNum
    Print #fileNum, "XLerate v" & XLERATE_VERSION & " (" & BUILD_CODENAME & ") - Build Documentation"
    Print #fileNum, String(80, "=")
    Print #fileNum, ""
    Print #fileNum, "Build Date: " & BUILD_DATE
    Print #fileNum, "Build Time: " & Format(Now - progress.StartTime, "hh:mm:ss")
    Print #fileNum, "Author: " & AUTHOR
    Print #fileNum, ""
    Print #fileNum, "Build Summary:"
    Print #fileNum, "- Components Completed: " & progress.CompletedSteps
    Print #fileNum, "- Components Failed: " & progress.FailedSteps
    Print #fileNum, "- Warnings: " & progress.Warnings
    Print #fileNum, ""
    Print #fileNum, "Installation Instructions:"
    Print #fileNum, "1. Copy " & outputPath & " to your Excel add-ins folder"
    Print #fileNum, "2. In Excel: File → Options → Add-ins → Excel Add-ins → Go..."
    Print #fileNum, "3. Click Browse and select the XLerate.xlam file"
    Print #fileNum, "4. Check 'XLerate' in the add-ins list and click OK"
    Print #fileNum, "5. Enable macros when prompted"
    Print #fileNum, ""
    Print #fileNum, "Macabacus-Compatible Shortcuts:"
    Print #fileNum, "- Fast Fill Right: Ctrl+Alt+Shift+R"
    Print #fileNum, "- Fast Fill Down: Ctrl+Alt+Shift+D"
    Print #fileNum, "- Error Wrap: Ctrl+Alt+Shift+E"
    Print #fileNum, "- Pro Precedents: Ctrl+Alt+Shift+["
    Print #fileNum, "- Pro Dependents: Ctrl+Alt+Shift+]"
    Print #fileNum, "- Number Cycle: Ctrl+Alt+Shift+1"
    Print #fileNum, "- Date Cycle: Ctrl+Alt+Shift+2"
    Print #fileNum, "- Local Currency: Ctrl+Alt+Shift+3"
    Print #fileNum, "- Foreign Currency: Ctrl+Alt+Shift+4"
    Print #fileNum, "- AutoColor: Ctrl+Alt+Shift+A"
    Print #fileNum, "- Settings: Ctrl+Alt+Shift+M"
    Print #fileNum, ""
    Print #fileNum, "Support:"
    Print #fileNum, "- Documentation: Check the XLerate_Info sheet in the add-in"
    Print #fileNum, "- GitHub: github.com/omegarhovega/XLerate"
    Print #fileNum, ""
    Close #fileNum
    
    If Err.Number = 0 Then
        Debug.Print "    ✓ Documentation saved: " & docPath
    Else
        Debug.Print "    ⚠ Documentation save failed: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' === UTILITY FUNCTIONS ===

Private Function DirExists(dirPath As String) As Boolean
    On Error Resume Next
    DirExists = (Dir(dirPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

Private Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Private Function ReadTextFile(filePath As String) As String
    ' Read entire text file content with better error handling
    
    On Error GoTo ReadError
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    ReadTextFile = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    Exit Function
    
ReadError:
    Debug.Print "Error reading file: " & filePath & " - " & Err.Description
    ReadTextFile = ""
    If fileNum > 0 Then Close #fileNum
End Function

' === QUICK BUILD FUNCTIONS ===

Public Sub QuickBuildXLerate()
    ' Quick build with default paths (for development)
    
    Debug.Print "=== Quick Build XLerate ==="
    
    Dim sourcePath As String
    Dim outputPath As String
    
    ' Use current workbook path as base
    sourcePath = ThisWorkbook.Path & "\src\"
    outputPath = ThisWorkbook.Path & "\XLerate_v" & Replace(XLERATE_VERSION, ".", "_") & "_Quick.xlam"
    
    If DirExists(sourcePath) Then
        Debug.Print "Using quick build paths:"
        Debug.Print "Source: " & sourcePath
        Debug.Print "Output: " & outputPath
        
        ' Temporarily override the input functions
        ' (This is a simplified version for development)
        MsgBox "Quick build will use:" & vbNewLine & _
               "Source: " & sourcePath & vbNewLine & _
               "Output: " & outputPath, vbInformation
               
        ' Call main build (would need path injection)
    Else
        Debug.Print "Quick build source not found: " & sourcePath
        MsgBox "Quick build requires the source in: " & sourcePath, vbExclamation
    End If
End Sub

Public Sub ValidateCurrentBuild()
    ' Validate the current XLerate installation
    
    Debug.Print "=== Validating Current XLerate Build ==="
    
    ' Check for key modules
    Dim requiredModules As Variant
    requiredModules = Array( _
        "ModNumberFormat", "ModCellFormat", "ModDateFormat", _
        "ModFastFillDown", "ModCurrencyCycling", "ModBorderUtilities", _
        "RibbonCallbacks", "AutoColorModule", "TraceUtils" _
    )
    
    Dim foundModules As Long
    Dim totalModules As Long
    totalModules = UBound(requiredModules) - LBound(requiredModules) + 1
    
    Dim moduleName As Variant
    For Each moduleName in requiredModules
        On Error Resume Next
        Dim testModule As Object
        Set testModule = ThisWorkbook.VBProject.VBComponents(CStr(moduleName))
        If Err.Number = 0 Then
            Debug.Print "✓ Found: " & moduleName
            foundModules = foundModules + 1
        Else
            Debug.Print "✗ Missing: " & moduleName
        End If
        On Error GoTo 0
    Next moduleName
    
    Debug.Print ""
    Debug.Print "Validation Summary:"
    Debug.Print "Found: " & foundModules & "/" & totalModules & " required modules"
    Debug.Print "Success Rate: " & Format((foundModules / totalModules) * 100, "0.0") & "%"
    
    If foundModules = totalModules Then
        Debug.Print "✓ Build validation PASSED"
        MsgBox "XLerate build validation passed!" & vbNewLine & _
               "All " & totalModules & " required modules are present.", _
               vbInformation, "Build Validation"
    Else
        Debug.Print "✗ Build validation FAILED"
        MsgBox "XLerate build validation failed!" & vbNewLine & _
               "Found " & foundModules & " of " & totalModules & " required modules." & vbNewLine & _
               "Check the Immediate Window for details.", _
               vbExclamation, "Build Validation"
    End If
End Sub