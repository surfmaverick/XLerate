# ⚡XLerate v2.0.0
**Enhanced Excel productivity add-in with Macabacus-compatible shortcuts**

XLerate is an open-source Excel add-in for Windows and Mac designed to speed up financial modeling tasks and spot potential errors with ease, featuring advanced auditing and formula consistency tools with Macabacus-compatible keyboard shortcuts.

<img src="/XLerate.png" alt="XLerate Add-in" width="800" height="auto"/>

## 🆕 What's New in v2.0.0

### Macabacus-Compatible Shortcuts
- **Fast Fill Right**: `Ctrl+Alt+Shift+R` (Macabacus standard)
- **Fast Fill Down**: `Ctrl+Alt+Shift+D` (Macabacus standard)  
- **Error Wrap**: `Ctrl+Alt+Shift+E` (Macabacus standard)
- **Pro Precedents**: `Ctrl+Alt+Shift+[` (Macabacus standard)
- **Pro Dependents**: `Ctrl+Alt+Shift+]` (Macabacus standard)
- **General Number Cycle**: `Ctrl+Alt+Shift+1` (Macabacus standard)
- **Date Cycle**: `Ctrl+Alt+Shift+2` (Macabacus standard)
- **AutoColor Selection**: `Ctrl+Alt+Shift+A` (Macabacus standard)
- **Quick Save**: `Ctrl+Alt+Shift+S` (Macabacus standard)
- **Toggle Gridlines**: `Ctrl+Alt+Shift+G` (Macabacus standard)
- **Zoom In/Out**: `Ctrl+Alt+Shift+=/−` (Macabacus standard)

### Enhanced Features
- **Smart Fill Down**: Vertical filling based on column patterns
- **Enhanced UI**: Redesigned ribbon with Macabacus-inspired layout
- **Cross-Platform**: Optimized for both Windows and macOS
- **Backward Compatibility**: All original shortcuts still work
- **Improved Performance**: Faster processing for large ranges

## 🚀 Core Features

### Advanced Formula Tracer
- **Pro Precedents** (`Ctrl+Alt+Shift+[`): Trace all precedents with interactive navigation
- **Pro Dependents** (`Ctrl+Alt+Shift+]`): Trace all dependents with enhanced visualization
- Quick navigation through complex formula chains
- Clear all arrows with `Ctrl+Alt+Shift+Delete`

### Smart Fill Functions
- **Fast Fill Right** (`Ctrl+Alt+Shift+R`): Intelligently fills formulas right based on row patterns
- **Fast Fill Down** (`Ctrl+Alt+Shift+D`): Intelligently fills formulas down based on column patterns
- Automatic boundary detection within 3 rows/columns
- Handles merged cells and complex ranges

### Formula Consistency Checker
- **Formula Consistency** (`Ctrl+Alt+Shift+C`): Visual highlighting of formula pattern breaks
- Green highlighting for consistent formulas
- Red highlighting for inconsistencies
- Toggle on/off to compare before and after

### Enhanced Format Cycling
- **Number Formats** (`Ctrl+Alt+Shift+1`): Cycle through custom number formats
- **Date Formats** (`Ctrl+Alt+Shift+2`): Cycle through custom date formats  
- **Cell Formats** (`Ctrl+Alt+Shift+3`): Cycle through background colors and borders
- **Text Styles** (`Ctrl+Alt+Shift+4`): Cycle through font styles and formatting
- Fully customizable through Settings Manager

### AutoColor System
- **AutoColor Selection** (`Ctrl+Alt+Shift+A`): Automatically color cells by content type
  - **Blue**: Input values and constants
  - **Black**: Standard formulas  
  - **Green**: Worksheet links
  - **Purple**: Workbook links and external references
  - **Orange**: Hyperlinks
  - **Custom colors**: Configurable through settings

### Error Management
- **Error Wrap** (`Ctrl+Alt+Shift+E`): Wrap formulas with IFERROR statements
- **Switch Sign** (`Ctrl+Alt+Shift+~`): Toggle positive/negative values
- Customizable error values (NA(), 0, "", etc.)

### View Controls
- **Toggle Gridlines** (`Ctrl+Alt+Shift+G`): Show/hide worksheet gridlines
- **Zoom In** (`Ctrl+Alt+Shift+=`): Increase zoom by 10%
- **Zoom Out** (`Ctrl+Alt+Shift+-`): Decrease zoom by 10%

### Utility Functions
- **Quick Save** (`Ctrl+Alt+Shift+S`): Save with visual confirmation
- **Settings Manager** (`Ctrl+Alt+Shift+,`): Configure all XLerate options
- **CAGR Function**: Built-in compound annual growth rate calculation
- **Reset Formats** (`Ctrl+Shift+0`): Reset all customizations to defaults

## 💾 Installation

### Windows 🪟
1. Download `XLerate.xlam` from the `dist` folder
2. Place in your Excel add-ins folder: `C:\Users\[Username]\AppData\Roaming\Microsoft\AddIns`
3. Enable in Excel: File → Options → Add-ins → Excel Add-ins → Go → Check "XLerate" → OK

**Note:** You may need to unblock the file:
1. Right-click `XLerate.xlam` → Properties
2. Check "Unblock" under Security → OK

### Mac 🍎
1. Download `XLerate.xlam` from the `dist` folder
2. Place in Excel add-ins folder:
   - **Office 365 (Big Sur+)**: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/Add-ins`
   - **Legacy versions**: `/Users/<username>/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins`
3. Enable in Excel: Excel → Tools → Excel Add-ins → Browse → Select `XLerate.xlam` → Check "XLerate" → OK

**Note:** If you see a security warning:
1. System Preferences → Security & Privacy
2. Click "Open Anyway" for XLerate.xlam

## 📖 Usage Guide

### Quick Start with Macabacus-Compatible Shortcuts

XLerate v2.0.0 uses the same keyboard shortcuts as Macabacus for seamless transition:

| Function | Macabacus Shortcut | XLerate Shortcut | Description |
|----------|-------------------|------------------|-------------|
| Fast Fill Right | `Ctrl+Alt+Shift+R` | `Ctrl+Alt+Shift+R` | Smart horizontal fill |
| Fast Fill Down | `Ctrl+Alt+Shift+D` | `Ctrl+Alt+Shift+D` | Smart vertical fill |
| Error Wrap | `Ctrl+Alt+Shift+E` | `Ctrl+Alt+Shift+E` | Add IFERROR wrapping |
| Pro Precedents | `Ctrl+Alt+Shift+[` | `Ctrl+Alt+Shift+[` | Advanced precedent trace |
| Pro Dependents | `Ctrl+Alt+Shift+]` | `Ctrl+Alt+Shift+]` | Advanced dependent trace |
| General Number | `Ctrl+Alt+Shift+1` | `Ctrl+Alt+Shift+1` | Cycle number formats |
| Date Cycle | `Ctrl+Alt+Shift+2` | `Ctrl+Alt+Shift+2` | Cycle date formats |
| AutoColor | `Ctrl+Alt+Shift+A` | `Ctrl+Alt+Shift+A` | Auto-color by content |
| Quick Save | `Ctrl+Alt+Shift+S` | `Ctrl+Alt+Shift+S` | Save with confirmation |
| Toggle Gridlines | `Ctrl+Alt+Shift+G` | `Ctrl+Alt+Shift+G` | Show/hide gridlines |

### Modeling Workflows

#### Fast Fill Operations
1. **Right Fill**: Select cell with formula → `Ctrl+Alt+Shift+R`
   - XLerate scans 3 rows above for data patterns
   - Automatically fills to the boundary of data
   
2. **Down Fill**: Select cell with formula → `Ctrl+Alt+Shift+D`  
   - XLerate scans 3 columns left for data patterns
   - Fills down to match the data boundary

#### Error Handling
1. **Wrap with IFERROR**: Select formulas → `Ctrl+Alt+Shift+E`
   - Wraps selected formulas: `=IFERROR(original_formula, NA())`
   - Configurable error values in Settings

2. **Switch Signs**: Select cells → `Ctrl+Alt+Shift+~`
   - Toggles positive/negative for values and formulas
   - Handles both numbers and formula references

### Auditing Workflows

#### Advanced Tracing
1. **Pro Precedents**: Select cell → `Ctrl+Alt+Shift+[`
   - Interactive dialog with all precedents
   - Navigate with arrow keys, `Esc` to close
   - Shows cell addresses, values, and formulas

2. **Pro Dependents**: Select cell → `Ctrl+Alt+Shift+]`
   - Interactive dialog with all dependents
   - Click any item to navigate to that cell
   - Real-time formula preview

#### Formula Consistency
1. **Check Consistency**: Select range → `Ctrl+Alt+Shift+C`
   - Green bars: Formulas consistent with neighbors
   - Red bars: Formulas inconsistent (potential errors)
   - Toggle off: Press `Ctrl+Alt+Shift+C` again

### Formatting Workflows

#### Format Cycling
All format cycles are **fully customizable** through Settings Manager:

1. **Number Formats** (`Ctrl+Alt+Shift+1`):
   - Default: General → Comma 0 → Comma 1 → Comma 2 → (repeat)
   - Add custom formats like thousands, millions, percentages

2. **Date Formats** (`Ctrl+Alt+Shift+2`):
   - Default: yyyy → mmm-yyyy → dd-mmm-yy → (repeat)
   - Add quarterly, weekly, or fiscal year formats

3. **Cell Formats** (`Ctrl+Alt+Shift+3`):
   - Default: Normal → Inputs → Good → Bad → Important → (repeat)
   - Customize colors, borders, and patterns

4. **Text Styles** (`Ctrl+Alt+Shift+4`):
   - Default: Heading → Subheading → Sum → (repeat)
   - Configure fonts, sizes, colors, and borders

#### AutoColor System
**AutoColor Selection** (`Ctrl+Alt+Shift+A`) applies intelligent coloring:

- **Input Detection**: Constants, user-entered values
- **Formula Types**: Simple formulas vs. complex calculations  
- **Link Classification**: 
  - Worksheet links (same workbook)
  - Workbook links (external workbooks)
  - External references (web services, databases)
- **Partial Inputs**: Formulas containing hardcoded numbers

Customize all colors in Settings → Auto-Color.

### Settings and Customization

#### Access Settings Manager
- **Ribbon**: XLerate tab → Utilities → Settings
- **Keyboard**: `Ctrl+Alt+Shift+,`

#### Configuration Options
1. **Numbers**: Add/edit/remove number format cycles
2. **Dates**: Customize date format patterns  
3. **Cells**: Configure background colors and border styles
4. **Text Styles**: Set up font combinations with borders
5. **Auto-Color**: Customize colors for each content type
6. **Error**: Set default error values for wrapping

#### Reset to Defaults
- **All Formats**: `Ctrl+Shift+0` - Resets everything to defaults
- **Individual**: Use Settings Manager to reset specific categories

## 🔧 Advanced Features

### CAGR Function
Built-in compound annual growth rate calculation:
```excel
=CAGR(A1:A5)  ' Calculates CAGR using first/last values and count
```

### Backward Compatibility
Original XLerate shortcuts still work:
- `Ctrl+Shift+1`: Number format cycle
- `Ctrl+Shift+2`: Cell format cycle  
- `Ctrl+Shift+3`: Date format cycle
- `Ctrl+Shift+R`: Smart Fill Right (original)

### Cross-Platform Notes
**Windows vs. macOS**:
- All shortcuts work identically on both platforms
- File paths differ for installation
- Performance optimized for both Office versions

**Office Versions**:
- Supports Office 365, Office 2019, Office 2021
- Compatible with both 32-bit and 64-bit installations
- Ribbon adapts to Office UI themes

## 💡 Tips and Best Practices

### Maximizing Productivity
1. **Learn the "Big 5" shortcuts**:
   - `Ctrl+Alt+Shift+R/D`: Fast Fill  
   - `Ctrl+Alt+Shift+[/]`: Pro Tracing
   - `Ctrl+Alt+Shift+1`: Number cycling

2. **Customize format cycles** for your workflow:
   - Add organization-specific number formats
   - Set up consistent cell coloring schemes
   - Configure date formats for reporting periods

3. **Use AutoColor systematically**:
   - Color inputs first with `Ctrl+Alt+Shift+A`
   - Review red-colored partial inputs for hardcoded values
   - Check green worksheet links for broken references

### Troubleshooting
- **Shortcuts not working**: Check if another add-in conflicts
- **Performance issues**: Reduce range sizes for large worksheets  
- **Format cycles stopped**: Reset with `Ctrl+Shift+0`
- **Settings not saving**: Ensure macro permissions are enabled

## 🛠️ Development and Contributing

### For Developers
XLerate is built with:
- **VBA**: Core functionality and ribbon interface
- **XML**: Custom ribbon definition (customUI14.xml)
- **Class Modules**: Object-oriented formatting and settings
- **Module Architecture**: Separated concerns for maintainability

### Contributing Guidelines
1. **Fork** the repository
2. **Create feature branch**: `git checkout -b feature/macabacus-shortcuts`
3. **Follow naming conventions**: Use descriptive function names
4. **Add version info**: Update changelog in file headers
5. **Test thoroughly**: Verify on both Windows and macOS
6. **Submit pull request**: Include description of changes

### Project Structure
```
src/
├── modules/           # Core functionality modules
├── forms/            # Settings UI forms  
├── class modules/    # Data type definitions
├── ribbon/           # Ribbon XML definition
└── objects/          # Workbook and worksheet events
```

## 📊 Comparison with Macabacus

| Feature | Macabacus | XLerate v2.0.0 | Notes |
|---------|-----------|----------------|-------|
| Fast Fill Right | ✅ `Ctrl+Alt+Shift+R` | ✅ `Ctrl+Alt+Shift+R` | Same shortcut |
| Fast Fill Down | ✅ `Ctrl+Alt+Shift+D` | ✅ `Ctrl+Alt+Shift+D` | Same shortcut |
| Error Wrap | ✅ `Ctrl+Alt+Shift+E` | ✅ `Ctrl+Alt+Shift+E` | Same shortcut |
| Pro Precedents | ✅ `Ctrl+Alt+Shift+[` | ✅ `Ctrl+Alt+Shift+[` | Same shortcut |
| Pro Dependents | ✅ `Ctrl+Alt+Shift+]` | ✅ `Ctrl+Alt+Shift+]` | Same shortcut |
| Number Cycle | ✅ `Ctrl+Alt+Shift+1` | ✅ `Ctrl+Alt+Shift+1` | Same shortcut |
| Date Cycle | ✅ `Ctrl+Alt+Shift+2` | ✅ `Ctrl+Alt+Shift+2` | Same shortcut |
| AutoColor | ✅ `Ctrl+Alt+Shift+A` | ✅ `Ctrl+Alt+Shift+A` | Same shortcut |
| Quick Save | ✅ `Ctrl+Alt+Shift+S` | ✅ `Ctrl+Alt+Shift+S` | Same shortcut |
| Toggle Gridlines | ✅ `Ctrl+Alt+Shift+G` | ✅ `Ctrl+Alt+Shift+G` | Same shortcut |
| Open Source | ❌ | ✅ | MIT License |
| Cost | 💰 Paid | 🆓 Free | Always free |
| Customization | Limited | ✅ Full | Complete control |

## 📄 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

## 💪 Support

- **Report bugs**: [GitHub Issues](https://github.com/omegarhovega/XLerate/issues)
- **Request features**: [Feature Request](https://github.com/omegarhovega/XLerate/issues/new?template=feature_request.md)
- **Discussions**: [GitHub Discussions](https://github.com/omegarhovega/XLerate/discussions)
- **Documentation**: [Wiki](https://github.com/omegarhovega/XLerate/wiki)

## 🙏 Acknowledgments

- **Inspired by Macabacus**: XLerate adopts the same keyboard shortcuts for seamless transition
- **Built by financial analysts**: For financial analysts who need speed and accuracy
- **Community driven**: Open source project welcoming contributions
- **Cross-platform**: Equal support for Windows and macOS users

---

**XLerate v2.0.0** - Making Excel faster, one shortcut at a time ⚡