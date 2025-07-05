# ⚡XLerate v2.0.0
XLerate is an open-source Excel add-in for Windows and Mac designed to speed up financial modeling tasks and spot potential errors with ease. Version 2.0 introduces comprehensive Macabacus-aligned keyboard shortcuts and enhanced productivity features.

<img src="/XLerate.png" alt="XLerate Add-in" width="800" height="auto"/>

## 🚀 What's New in v2.0.0

### 🎯 Macabacus-Aligned Keyboard Shortcuts
- **Complete shortcut mapping** using `Ctrl+Alt+Shift+Key` combinations
- **Seamless transition** from Macabacus workflows
- **Consistent patterns** for financial modeling efficiency

### 🔧 Enhanced Border Management
- **Quick border application** with directional shortcuts
- **Border cycling** through none → thin → thick → double
- **Professional formatting** for financial statements

### 📊 Advanced Productivity Features
- **Smart Fill Down** for vertical pattern recognition
- **Quick save with timestamp** for version control
- **Enhanced paste operations** (values only, etc.)
- **View management utilities** (zoom to selection, toggle gridlines)

## 📋 Complete Keyboard Shortcut Reference

### 🎨 Format Cycling
| Shortcut | Function | Description |
|----------|----------|-------------|
| `Ctrl+Alt+Shift+1` | Number Format Cycle | Cycle through configured number formats |
| `Ctrl+Alt+Shift+2` | Date Format Cycle | Cycle through configured date formats |
| `Ctrl+Alt+Shift+3` | Cell Format Cycle | Cycle through configured cell formats |
| `Ctrl+Alt+Shift+4` | Text Style Cycle | Cycle through configured text styles |

### 🧮 Modeling & Analysis
| Shortcut | Function | Description |
|----------|----------|-------------|
| `Ctrl+Alt+Shift+R` | Smart Fill Right | Auto-fill formulas based on row patterns |
| `Ctrl+Alt+Shift+D` | Smart Fill Down | Auto-fill formulas based on column patterns |
| `Ctrl+Alt+Shift+E` | Error Wrap | Wrap formulas with IFERROR statements |
| `Ctrl+Alt+Shift+S` | Switch Sign | Toggle positive/negative values |
| `Ctrl+Alt+Shift+G` | Insert CAGR | Insert CAGR formula template |

### 🔍 Auditing & Tracing
| Shortcut | Function | Description |
|----------|----------|-------------|
| `Ctrl+Alt+Shift+[` | Trace Precedents | Show all cells feeding into selection |
| `Ctrl+Alt+Shift+]` | Trace Dependents | Show all cells dependent on selection |
| `Ctrl+Alt+Shift+C` | Formula Consistency | Check formula consistency across rows |

### 🎨 Border Management
| Shortcut | Function | Description |
|----------|----------|-------------|
| `Ctrl+Alt+Shift+↓` | Bottom Border | Apply bottom border to selection |
| `Ctrl+Alt+Shift+↑` | Top Border | Apply top border to selection |
| `Ctrl+Alt+Shift+←` | Left Border | Apply left border to selection |
| `Ctrl+Alt+Shift+→` | Right Border | Apply right border to selection |
| `Ctrl+Alt+Shift+O` | Outside Border | Apply border around entire selection |
| `Ctrl+Alt+Shift+N` | No Border | Remove all borders from selection |

### 🛠️ Productivity Utilities
| Shortcut | Function | Description |
|----------|----------|-------------|
| `Ctrl+Alt+Shift+A` | Auto-Color Cells | Color cells by content type (inputs, formulas, links) |
| `Ctrl+Alt+Shift+V` | Paste Values Only | Paste values without formulas or formatting |
| `Ctrl+Alt+Shift+Q` | Quick Save + Timestamp | Save with timestamp appended to filename |
| `Ctrl+Alt+Shift+L` | Toggle Gridlines | Show/hide worksheet gridlines |

### ⚙️ Settings & Management
| Shortcut | Function | Description |
|----------|----------|-------------|
| `Ctrl+Alt+Shift+M` | Settings Manager | Open comprehensive settings interface |
| `Ctrl+Alt+Shift+0` | Reset All Formats | Reset all format settings to defaults |

## 🚀 Features

### Advanced Formula Tracer
- **Comprehensive dependency mapping** - Trace all precedents and dependents in a single operation
- **Interactive navigation** - Click through formula chains with ease
- **Multi-sheet support** - Follow references across worksheets and workbooks

### Formula Consistency Checker
- **Pattern recognition** - Instantly identify inconsistencies in adjacent formulas
- **Visual highlighting** - Green for consistent, red for inconsistent patterns
- **Error prevention** - Catch common modeling mistakes before they propagate

### Enhanced Format Cycling
- **Four format categories** - Numbers, dates, cells, and text styles
- **Customizable formats** - Add, remove, and modify format collections
- **Instant application** - One keystroke to cycle through professional formats

### Smart Fill Functions
- **Intelligent pattern detection** - Analyzes data in surrounding rows/columns
- **Bi-directional filling** - Fill right or down based on context
- **Formula-aware** - Maintains proper cell references during filling

### Comprehensive Border Management
- **Directional shortcuts** - Apply borders using arrow key combinations
- **Style cycling** - Rotate through none → thin → thick → double borders
- **Professional formatting** - Perfect for financial statement presentation

### Advanced Auto-Coloring
- **Content-type recognition** - Distinguishes inputs, formulas, links, and external references
- **Customizable color schemes** - Modify colors to match your workflow
- **Instant application** - Color entire selections with one command

### Additional Productivity Features
- **Error wrapping** - Automatically wrap formulas with IFERROR statements
- **Sign switching** - Toggle positive/negative values instantly
- **CAGR function** - Built-in compound annual growth rate calculations
- **Quick save with versioning** - Timestamp-based file saving
- **View management** - Gridline toggling, zoom to selection

## 💾 Installation

### Windows 🪟
1. Download the latest release (`XLerate.xlam`) from the `dist` folder
2. Place the file in your Excel add-ins folder (typically `C:\Users\[Username]\AppData\Roaming\Microsoft\AddIns`)
3. Enable the add-in in Excel:
   - File → Options → Add-ins
   - Manage: Excel Add-ins → Go
   - Check "XLerate" and click OK

**Note:** Due to Windows security settings, you might need to unblock the add-in: 
1. Right-click the `XLerate.xlam` file and select Properties 
2. On the General tab under Security, check the "Unblock" box

### Mac 🍎
1. Download the latest release (`XLerate.xlam`) from the `dist` folder
2. Place the file in your Excel add-ins folder:
   - Modern Office 365 (Big Sur and newer): `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/Add-ins`
   - Legacy versions: `/Users/<username>/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins`
   - If not available, search for "Add-ins" folder
3. Enable the add-in in Excel:
   - Excel → Tools → Excel Add-ins
   - Click "Browse" if XLerate isn't listed
   - Locate and select `XLerate.xlam`
   - Check "XLerate" and click OK

**Note:** On Mac, if you receive a security warning, you may need to:
1. Open System Preferences → Security & Privacy
2. Click "Open Anyway" for `XLerate.xlam`

## 📖 Usage Guide

### Quick Start
After installation, you'll find the **XLerate v2.0** tab in your Excel ribbon with six main groups:

#### 1. Modeling Group
- **Smart Fill**: Intelligent formula filling based on surrounding data patterns
- **Switch Sign**: Instantly toggle positive/negative values
- **Error Wrap**: Add IFERROR protection to formulas
- **Insert CAGR**: Add compound annual growth rate calculations

#### 2. Auditing Group  
- **Trace Precedents/Dependents**: Interactive formula dependency mapping
- **Formula Consistency**: Visual highlighting of formula pattern breaks

#### 3. Formatting Group
- **Format Cycling**: Rotate through number, date, cell, and text formats
- **Auto-Color**: Intelligent cell coloring based on content type

#### 4. Borders Group
- **Directional Borders**: Apply borders using intuitive shortcuts
- **Style Options**: Choose from thin, thick, double, or no borders
- **Quick Formatting**: Professional financial statement borders

#### 5. Utilities Group
- **Paste Operations**: Values-only pasting and other specialized paste functions
- **File Management**: Quick save with timestamps for version control
- **View Controls**: Gridline toggling and zoom management

#### 6. Settings Group
- **Configuration**: Customize all format cycles and color schemes
- **Reset Options**: Restore default settings when needed

### Advanced Usage Tips

#### Financial Model Setup
1. Use **Auto-Color** (`Ctrl+Alt+Shift+A`) to distinguish inputs from calculations
2. Apply **Outside Borders** (`Ctrl+Alt+Shift+O`) to define model sections
3. Use **Formula Consistency** (`Ctrl+Alt+Shift+C`) to validate formula patterns
4. Apply **Bottom Borders** (`Ctrl+Alt+Shift+↓`) for section totals

#### Efficient Data Entry
1. Set up custom number formats in Settings for your standard formats
2. Use **Number Format Cycle** (`Ctrl+Alt+Shift+1`) to quickly format values
3. Employ **Smart Fill Right** (`Ctrl+Alt+Shift+R`) for extending formulas
4. Use **Error Wrap** (`Ctrl+Alt+Shift+E`) to add error handling

#### Model Auditing Workflow
1. Select key calculation cells
2. Use **Trace Precedents** (`Ctrl+Alt+Shift+[`) to verify inputs
3. Use **Trace Dependents** (`Ctrl+Alt+Shift+]`) to check impact
4. Run **Formula Consistency** check across model sections

#### Presentation Preparation
1. Use **Toggle Gridlines** (`Ctrl+Alt+Shift+L`) for clean presentation view
2. Apply consistent borders using the border shortcuts
3. Ensure proper formatting with format cycling functions
4. Save presentation version with **Quick Save + Timestamp**

## ⚙️ Customization

### Format Cycling Configuration
Access the Settings Manager (`Ctrl+Alt+Shift+M`) to customize:

- **Number Formats**: Add accounting, percentage, or custom number formats
- **Date Formats**: Configure region-specific or custom date displays  
- **Cell Formats**: Set up highlighting schemes for different cell types
- **Text Styles**: Define heading, subheading, and emphasis styles
- **Auto-Color Schemes**: Customize colors for different content types

### Keyboard Shortcut Customization
XLerate v2.0 uses a consistent `Ctrl+Alt+Shift+Key` pattern aligned with Macabacus conventions. This ensures:
- **Muscle memory compatibility** for Macabacus users
- **No conflicts** with existing Excel shortcuts
- **Logical grouping** of related functions

## 🔧 Troubleshooting

### Common Issues

#### "Constant expression required" Error
This VBA compile error has been resolved in v2.0.0 by replacing constants with properly scoped variables.

#### Shortcuts Not Working
1. Ensure the add-in is properly loaded (check for XLerate tab in ribbon)
2. Try restarting Excel to refresh the shortcut registration
3. Use **Reset All Formats** (`Ctrl+Alt+Shift+0`) to reinitialize settings

#### Performance with Large Models
- XLerate includes optimizations for large workbooks (>1000 cells)
- Uses screen updating and calculation controls during operations
- Provides progress feedback for long-running operations

#### Mac-Specific Issues
- Ensure Excel has proper permissions to access the add-ins folder
- Some shortcuts may require using `Cmd` instead of `Ctrl` depending on Excel version
- Check System Preferences → Security & Privacy if the add-in won't load

### Getting Help
- **Debug Information**: Enable debug output in Settings for detailed logging
- **Error Reporting**: Use the built-in error reporting in the Settings Manager
- **Documentation**: Access built-in help through the Settings interface

## 🛠️ Development & Contributing

### Development Environment Setup
1. Enable "Trust access to the VBA project object model" in Excel's Trust Center
2. Clone the repository: `git clone https://github.com/omegarhovega/XLerate.git`
3. Open the XLAM file in Excel Developer mode
4. Use the provided export utility to manage code changes

### Version 2.0.0 Architecture
- **Modular design** with separated concerns (formatting, borders, utilities)
- **Enhanced error handling** throughout all modules
- **Consistent naming conventions** for maintainability
- **Comprehensive logging** for debugging and support

### Code Organization
```
src/
├── modules/           # Core functionality modules
├── forms/            # User interface forms
├── class modules/    # Data type definitions
├── ribbon/          # Ribbon interface definition
└── objects/         # Workbook and application events
```

### Contributing Guidelines
1. Fork the repo and create your branch from `master`
2. Follow the established coding standards and naming conventions
3. Add comprehensive error handling and debug output
4. Test on both Windows and Mac platforms
5. Update documentation for any new features
6. Submit pull request with detailed description of changes

### Development Process
- **Feature branches** for new functionality
- **Version numbering** follows semantic versioning (major.minor.patch)
- **Changelog maintenance** for all modifications
- **Cross-platform testing** before releases

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 💪 Support & Community

- **Bug Reports**: [GitHub Issues](https://github.com/omegarhovega/XLerate/issues)
- **Feature Requests**: Use our [Feature Request Template](https://github.com/omegarhovega/XLerate/issues/new?template=feature_request.md)
- **Discussions**: Join community discussions in the GitHub Discussions section
- **Documentation**: Comprehensive guides available in the `/docs` folder

## 🙏 Acknowledgments

- Built with 💖 for financial analysts by financial analysts
- Inspired by Macabacus workflow patterns and keyboard shortcut conventions
- Special thanks to the Excel VBA community for best practices and optimization techniques
- Thanks to all contributors and beta testers who helped shape v2.0.0

## 🔄 Migration from v1.x

### Shortcut Changes
Version 2.0.0 introduces a new, more consistent shortcut scheme:

**Old (v1.x)** → **New (v2.0.0)**
- `Ctrl+Shift+1` → `Ctrl+Alt+Shift+1` (Number formats)
- `Ctrl+Shift+2` → `Ctrl+Alt+Shift+3` (Cell formats) 
- `Ctrl+Shift+3` → `Ctrl+Alt+Shift+2` (Date formats)
- `Ctrl+Shift+R` → `Ctrl+Alt+Shift+R` (Smart fill right)

### New Features in v2.0.0
- Complete border management system
- Enhanced productivity utilities
- Improved error handling and debugging
- Macabacus-aligned workflow patterns
- Advanced auto-coloring capabilities

### Settings Migration
Your existing format settings will be automatically migrated when you first open v2.0.0. If you encounter any issues, use the "Reset All Formats" function to start fresh with the new defaults.

---

**XLerate v2.0.0** - Transforming Excel financial modeling workflows with professional-grade productivity tools.