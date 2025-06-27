⚡XLerate - Enhanced with Macabacus-style Shortcuts
XLerate is an open-source Excel add-in for Windows and Mac designed to speed up financial modelling tasks and spot potential errors with ease, featuring advanced auditing and formula consistency tools with Macabacus-compatible keyboard shortcuts.

<img src="/XLerate.png" alt="XLerate Add-in" width="800" height="auto"/>

🚀 Enhanced Features

Advanced Formula Tracer

Pro Precedents (Ctrl+Alt+Shift+[): Enhanced precedent tracing with navigation

Pro Dependents (Ctrl+Alt+Shift+]): Enhanced dependent tracing with navigation

Show All Precedents (Ctrl+Alt+Shift+F): Display all precedent relationships

Show All Dependents (Ctrl+Alt+Shift+J): Display all dependent relationships

Clear Arrows (Ctrl+Alt+Shift+N): Remove all tracing arrows

Enhanced Formula Consistency & Auditing
Horizontal Consistency: Visual highlighting of formula pattern breaks
Uniformulas (Ctrl+Alt+Shift+Q): Make formulas uniform across ranges
Formula Analysis: Complexity analysis and optimization suggestions
Formula Validation: Automatic error detection and warnings
Macabacus-style Format Cycling
General Number Cycle (Ctrl+Alt+Shift+1): Cycle through general number formats
Date Cycle (Ctrl+Alt+Shift+2): Cycle through date formats
Local Currency Cycle (Ctrl+Alt+Shift+3): Cycle through local currency formats
Foreign Currency Cycle (Ctrl+Alt+Shift+4): Cycle through foreign currency formats
Percent Cycle (Ctrl+Alt+Shift+5): Cycle through percentage formats
Multiple Cycle (Ctrl+Alt+Shift+8): Cycle through multiple/ratio formats
Binary Cycle (Ctrl+Alt+Shift+Y): Cycle through Yes/No, True/False formats
Increase/Decrease Decimals (Ctrl+Alt+Shift+] / Ctrl+Alt+Shift+[)
Advanced Color Functions
Blue-Black Toggle (Ctrl+Alt+Shift+9): Toggle between blue and black font
Font Color Cycle (Ctrl+Alt+Shift+C): Cycle through font colors
Fill Color Cycle (Ctrl+Alt+Shift+K): Cycle through background colors
Border Color Cycle (Ctrl+Alt+Shift+'): Cycle through border colors
AutoColor Selection (Ctrl+Alt+Shift+A): Auto-color selected cells
AutoColor Sheet (Ctrl+Alt+Shift+S): Auto-color entire sheet
AutoColor Workbook (Ctrl+Alt+Shift+Q): Auto-color entire workbook
Smart Modeling Functions
Fast Fill Right (Ctrl+Alt+Shift+R): Smart-fill formulas right based on patterns
Fast Fill Down (Ctrl+Alt+Shift+D): Smart-fill formulas down
Error Wrap (Ctrl+Alt+Shift+E): Wrap formulas with IFERROR statements
Simplify Formula (Ctrl+Alt+Shift+V): Simplify complex formulas
Switch Sign (Ctrl+Alt+Shift+~): Change signs of values/formulas
Alignment & Layout
Center Cycle (Ctrl+Alt+Shift+C): Cycle through center alignments
Horizontal Cycle (Ctrl+Alt+Shift+H): Cycle through horizontal alignments
Left Indent Cycle (Ctrl+Alt+Shift+I): Cycle through indent levels
Row Height Cycle (Ctrl+Alt+Shift+PgUp): Cycle through standard row heights
Column Width Cycle (Ctrl+Alt+Shift+PgDn): Cycle through standard column widths
Border Management
Bottom Border Cycle (Ctrl+Alt+Shift+Down): Cycle bottom border styles
Left Border Cycle (Ctrl+Alt+Shift+Left): Cycle left border styles
Right Border Cycle (Ctrl+Alt+Shift+Right): Cycle right border styles
Outside Border Cycle (Ctrl+Alt+Shift+7): Cycle outside border styles
No Border (Ctrl+Alt+Shift+-): Remove all borders
View & Display
Zoom In/Out (Ctrl+Alt+Shift+= / Ctrl+Alt+Shift+-): Adjust zoom levels
Toggle Gridlines (Ctrl+Alt+Shift+G): Show/hide gridlines
Hide Page Breaks (Ctrl+Alt+Shift+B): Hide page break lines
Cycle View Mode: Switch between Normal, Page Break, and Page Layout views
Toggle Freeze Panes: Freeze/unfreeze panes at selection
Font Functions
Increase/Decrease Font (Ctrl+Alt+Shift+F / Ctrl+Alt+Shift+G): Adjust font size
Underline Cycle (Ctrl+Alt+Shift+U): Cycle through underline styles
Wrap Text (Ctrl+Alt+Shift+W): Toggle text wrapping
Utilities
Quick Save (Ctrl+Alt+Shift+S): Save current workbook
Quick Save All: Save all open workbooks
Reset All Formats (Ctrl+Alt+Shift+0): Reset all customizations to defaults
📖 Usage
After installation, XLerate provides over 50 Macabacus-compatible keyboard shortcuts for maximum productivity:

Most Used Shortcuts (Macabacus-style)
MODELING
Ctrl+Alt+Shift+R    Fast Fill Right
Ctrl+Alt+Shift+E    Error Wrap
Ctrl+Alt+Shift+~    Switch Sign

AUDITING  
Ctrl+Alt+Shift+[    Pro Precedents
Ctrl+Alt+Shift+]    Pro Dependents
Ctrl+Alt+Shift+N    Clear Arrows

NUMBERS
Ctrl+Alt+Shift+1    General Number Cycle
Ctrl+Alt+Shift+2    Date Cycle
Ctrl+Alt+Shift+3    Local Currency Cycle
Ctrl+Alt+Shift+4    Foreign Currency Cycle
Ctrl+Alt+Shift+5    Percent Cycle

COLORS
Ctrl+Alt+Shift+9    Blue-Black Toggle
Ctrl+Alt+Shift+A    AutoColor Selection
Ctrl+Alt+Shift+K    Fill Color Cycle

ALIGNMENT
Ctrl+Alt+Shift+C    Center Cycle
Ctrl+Alt+Shift+H    Horizontal Cycle
Ctrl+Alt+Shift+W    Wrap Text
Enhanced Formula Tracer
Select a cell with a formula
Press Ctrl+Alt+Shift+[ for Pro Precedents or Ctrl+Alt+Shift+] for Pro Dependents
Navigate through dependencies using arrow keys in the interactive dialog
Press Esc to close the dialog
Consistency Checker
Select your worksheet data
Press Ctrl+Alt+Shift+F or click "Horizontal Consistency" in the XLerate ribbon
Review highlighted inconsistencies (red) and consistent form



XLerate Compilation Guide: Building the XLAM File
Prerequisites

Microsoft Excel (2016 or newer)
Developer Tab enabled
VBA Editor access (Alt + F11)
Trust settings configured for VBA macros

Method 1: Manual Import Method (Recommended)
Step 1: Create Base XLAM File

Open Excel
Press Alt + F11 to open VBA Editor
In VBA Editor: File → New → Microsoft Excel Add-In
Save as XLerate.xlam in your working directory

Step 2: Import All Source Files
Import the following files in this order:
Class Modules (Import via VBA Editor → File → Import File)

clsFormatType.cls
clsCellFormatType.cls
clsTextStyleType.cls
clsUISettings.cls
clsListBoxHandler.cls
clsDynamicButtonHandler.cls

Standard Modules

ModNumberFormat.bas
ModCellFormat.bas
ModDateFormat.bas
ModTextStyle.bas
AutoColorModule.bas
ModAuditing.bas (new)
ModAlignment.bas (new)
ModView.bas (new)
ModErrorWrap.bas
ModSwitchSign.bas
ModSmartFillRight.bas
ModCAGR.bas
ModFormatReset.bas
ModSettings.bas
ModGlobalSettings.bas
FormulaConsistency.bas
TraceUtils.bas
RibbonCallbacks.bas

Forms (UserForms)

frmNumberSettings.frm
frmCellSettings.frm
frmDateSettings.frm
frmTextStyle.frm
frmAutoColor.frm
frmErrorSettings.frm
frmSettingsManager.frm
frmPrecedents.frm
frmDependents.frm

Objects

Replace ThisWorkbook.cls content with the updated version

Step 3: Add Ribbon Customization

Close the XLAM file in Excel
Rename the file from XLerate.xlam to XLerate.zip
Extract the ZIP file
Navigate to the customUI folder (create if it doesn't exist)
Add the file customUI14.xml with the updated ribbon content
Re-zip the folder and rename back to XLerate.xlam

Step 4: Test and Validate

Open Excel
File → Options → Add-ins → Manage Excel Add-ins → Go
Browse and select your XLerate.xlam file
Check the XLerate checkbox and click OK
Verify the XLerate tab appears in the ribbon
Test some keyboard shortcuts:

Ctrl + Alt + Shift + R (Fast Fill Right)
Ctrl + Alt + Shift + 1 (Cycle Number Formats)
Ctrl + Alt + Shift + A (Auto Color)



Method 2: Automated VBA Build Script
Create this VBA script in a separate workbook to automate the build:
vbSub BuildXLerate()
    Dim sourcePath As String
    Dim targetPath As String
    
    sourcePath = "C:\path\to\xlerate\src\"  ' Update this path
    targetPath = "C:\path\to\xlerate\dist\XLerate.xlam"
    
    ' Create new add-in
    Dim newAddin As Workbook
    Set newAddin = Workbooks.Add(xlWBATemplate)
    
    ' Import all modules and classes
    ImportAllModules sourcePath, newAddin
    
    ' Save as XLAM
    newAddin.SaveAs targetPath, xlAddIn
    newAddin.Close
    
    MsgBox "XLerate build complete: " & targetPath
End Sub

Private Sub ImportAllModules(sourcePath As String, targetWB As Workbook)
    ' Import class modules
    targetWB.VBProject.VBComponents.Import sourcePath & "class modules\clsFormatType.cls"
    targetWB.VBProject.VBComponents.Import sourcePath & "class modules\clsCellFormatType.cls"
    ' ... (add all other imports)
End Sub
Key Updated Features (Macabacus-style)
Enhanced Keyboard Shortcuts

Modeling: Ctrl+Alt+Shift+R (Fast Fill Right), Ctrl+Alt+Shift+E (Error Wrap)
Auditing: Ctrl+Alt+Shift+[ (Pro Precedents), Ctrl+Alt+Shift+] (Pro Dependents)
Numbers: Ctrl+Alt+Shift+1-8 (Various number format cycles)
Colors: Ctrl+Alt+Shift+9 (Blue-Black Toggle), Ctrl+Alt+Shift+A (Auto Color)
Alignment: Ctrl+Alt+Shift+C (Center Cycle), Ctrl+Alt+Shift+H (Horizontal Cycle)
View: Ctrl+Alt+Shift+= (Zoom In), Ctrl+Alt+Shift+G (Toggle Gridlines)

New Functions Added

Multiple Number Format Types: Local Currency, Foreign Currency, Percent, Multiple, Binary
Enhanced Auto-Color: Sheet and Workbook level coloring
Advanced Auditing: Show All Precedents/Dependents, Uniformulas, Formula Analysis
Alignment Functions: Center Cycle, Indent Cycle, Text Orientation
View Management: Zoom controls, Freeze Panes, View Mode cycling

File Structure Summary
XLerate/
├── src/
│   ├── class modules/
│   │   ├── clsFormatType.cls
│   │   ├── clsCellFormatType.cls
│   │   ├── clsTextStyleType.cls
│   │   ├── clsUISettings.cls
│   │   ├── clsListBoxHandler.cls
│   │   └── clsDynamicButtonHandler.cls
│   ├── modules/
│   │   ├── ModNumberFormat.bas (enhanced)
│   │   ├── ModCellFormat.bas
│   │   ├── ModDateFormat.bas
│   │   ├── ModTextStyle.bas
│   │   ├── AutoColorModule.bas (enhanced)
│   │   ├── ModAuditing.bas (new)
│   │   ├── ModAlignment.bas (new)
│   │   ├── ModView.bas (new)
│   │   ├── ModErrorWrap.bas
│   │   ├── ModSwitchSign.bas
│   │   ├── ModSmartFillRight.bas
│   │   ├── ModCAGR.bas
│   │   ├── ModFormatReset.bas
│   │   ├── ModSettings.bas
│   │   ├── ModGlobalSettings.bas
│   │   ├── FormulaConsistency.bas
│   │   ├── TraceUtils.bas
│   │   └── RibbonCallbacks.bas (enhanced)
│   ├── forms/
│   │   ├── frmNumberSettings.frm
│   │   ├── frmCellSettings.frm
│   │   ├── frmDateSettings.frm
│   │   ├── frmTextStyle.frm
│   │   ├── frmAutoColor.frm
│   │   ├── frmErrorSettings.frm
│   │   ├── frmSettingsManager.frm
│   │   ├── frmPrecedents.frm
│   │   └── frmDependents.frm
│   ├── objects/
│   │   └── ThisWorkbook.cls (updated)
│   └── ribbon/
│       └── customUI14.xml (enhanced)
├── dist/
│   └── XLerate.xlam (compiled output)
└── README.md