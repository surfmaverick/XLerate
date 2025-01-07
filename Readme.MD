# ⚡XLerate
XLerate is an open-source Excel add-in for Windows and Mac designed to speed up financial modeling tasks and spot potential errors with ease, incl. advanced auditing and formula consistency tools.

<img src="/XLerate.png" alt="XLerate Add-in" width="800" height="auto"/>

## 🚀 Features

### Advanced Formula Tracer
- Trace all precedents and dependents of a cell in a single operation
- Quick navigation through formula chains

### Formula Consistency Checker
- Instantly identify inconsistencies in adjacent formulas
- Visual highlighting of formula pattern breaks to help prevent common modeling errors

### Cell Format Cycling
- Cycling through set number, cell and date formats with shortcuts `Ctrl + Shift + 1-3`
- Settings menu allows to add/remove/modify custom formats

### Other Features
- Wrap cells in IFERROR statements instantly
- Switch signs of values and formulas instantly
- New CAGR function to quickly add compound annual growth rates
- Currently in development (see below how to contribute)

## 💾 Installation

### Windows 🪟
1. Download the latest release (`XLerate.xlam`) from the `dist` folder
2. Place the file in your Excel add-ins folder (typically `C:\Users\[Username]\AppData\Roaming\Microsoft\AddIns`)
3. Enable the add-in in Excel:
   - File → Options → Add-ins
   - Manage: Excel Add-ins → Go
   - Check "XLerate" and click OK

**Note:** Due to Windows security settings, you might have to unblock the add-in: 
1. Right-click the `XLerate.xlam` file and select Properties 
2. On the General tab under Security tick the "Unblock" box

### Mac 🍎
1. Download the latest release (`XLerate.xlam`) from the `dist` folder
2. Place the file in your Excel add-ins folder:
   - Modern Office 365 (Big Sur and newer): `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/Add-ins`
   - Legacy versions: `/Users/<username>/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins`
   - If not available search for "Add-ins" folder
3. Enable the add-in in Excel:
   - Excel → Tools → Excel Add-ins
   - Click "Browse" if XLerate isn't listed
   - Locate and select `XLerate.xlam`
   - Check "XLerate" and click OK

**Note:** On Mac, if you receive a security warning, you may need to:
1. Open System Preferences → Security & Privacy
2. Click "Open Anyway" for `XLerate.xlam`

## 📖 Usage

After installation, you'll find the XLerate tab in your Excel ribbon:

### Formula Tracer
1. Select a cell with a formula
2. Click "Trace Precedents" or "Trace Dependents" in the XLerate ribbon
3. Navigate through dependencies using the arrow keys in the interactive dialog box
4. To quickly close the dialog window hit `Esc`

### Consistency Checker
1. With an open sheet click "Horizontal Consistency" in the XLerate ribbon under Auditing   
2. Review highlighted inconsistencies in the selection (red) and check which formulas are consistent (green)
3. Click "Horizontal Consistency" again to remove the formatting   

### Number, Cell and Date Formatting Cycles
1. Select a cell or range of cells
2. Click "Format" in the XLerate ribbon under Formatting and select one of the options repeatedly to cycle through the different preset formats
3. For faster access, the number, cell and date cycles can be accessed via shortcuts: Pressing `Ctrl + Shift + 1` repeatedly cycles through set number formats, `+ 2` through cell fromats, and `+ 3` through date fromats
4. Click on the "Settings" button in ribbon to access the predifined styles for numbers, cells and dates and to add your own styles to the cycles, or remove or modify existing ones

### Error Wrapping
1. Select a cell or range of cells
2. Click "Error Wrap" in the XLerate ribbon under Formatting to add an IFERROR statement to any formulas within your selection
3. If one of the selected cells does not contain a formula and is not affected a pop-up window will alert you to this

### Switch Sings
1. Select a cell or range of cells
2. Click "Switch Sign" in the XLerate ribbon under Formulas to change the sign of a cell or formula from positive to negative or vice versa

### CAGR Function
1. Type `+CAGR` and select a range of cells with numbers
2. The formula will calculate an annual CAGR based on the first and last number in the selection and the number of values (years)

We **highly recommend** adding individual functions you use often to Excel's quick access toolbar (Excel → File → Options → Quick Access Toolbar). In the "Chose commands from:", dropdown select "XLerate Tab" and add any function as needed.

## 🛠️ Contributing

We love your input! We want to make contributing to XLerate as easy and transparent as possible, whether it's:

- Reporting a bug
- Discussing the current state of the code
- Submitting a fix
- Proposing new features
- Becoming a maintainer

### Development Process

1. Fork the repo and create your branch from `master`
2. If you've added code that should be tested, add tests
3. Ensure the test suite passes
4. Make sure your code follows the existing style
5. Issue that pull request!

### Development Environment Setup

1. Enable "Trust access to the VBA project object model" in Excel's Trust Center
2. Clone the repository
3. Open the XLAM file in Excel Developer mode
4. Use the provided export utility to export changes

### Any contributions you make will be under the MIT Software License
In short, when you submit code changes, your submissions are understood to be under the same [MIT License](http://choosealicense.com/licenses/mit/) that covers the project. Feel free to contact the maintainers if that's a concern.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 💪 Support

- Report bugs via [GitHub Issues](https://github.com/[your-username]/XLerate/issues)
- Request features through our [Feature Request](https://github.com/[your-username]/XLerate/issues/new?template=feature_request.md) template

## 🙏 Acknowledgments

Built with 💖 for financial analysts by financial analysts.
