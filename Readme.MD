# ⚡XLerate
XLerate is an Excel add-in for Windows and Mac designed to speed up financial modeling tasks and spot potential errors with ease, incl. advanced auditing and formula consistency tools.

## 🚀 Features

### Advanced Formula Tracer
- Trace all precedents and dependents in a single operation
- Visualize formula dependencies across multiple worksheets
- Quick navigation through formula chains

### Formula Consistency Checker
- Instantly identify inconsistencies in adjacent formulas
- Visual highlighting of formula pattern breaks
- Help prevent common modeling errors
- Support for complex financial modeling patterns

### Formatting cycling
- Cycling through set number formats with one shortcut
- Setting custom formattings in menu

### Other Features
- Currently in development (see below how to contribute)

## 💾 Installation

### Windows 🪟
1. Download the latest release (`XLerate.xlam`) from the `dist` folder
2. Place the file in your Excel add-ins folder (typically `C:\Users\[Username]\AppData\Roaming\Microsoft\AddIns`)
3. Enable the add-in in Excel:
   - File → Options → Add-ins
   - Manage: Excel Add-ins → Go
   - Check "XLerate" and click OK

### Mac 🍎
1. Download the latest release (`XLerate.xlam`) from the `dist` folder
2. Place the file in your Excel add-ins folder:
   - Modern Office 365 (Big Sur and newer): `~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/Add-ins`
   - Legacy versions: `~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins`
3. Enable the add-in in Excel:
   - Excel → Tools → Excel Add-ins
   - Click "Browse" if XLerate isn't listed
   - Locate and select XLerate.xlam
   - Check "XLerate" and click OK

Note: On Mac, if you receive a security warning, you may need to:
1. Open System Preferences → Security & Privacy
2. Click "Open Anyway" for XLerate.xlam

## 📖 Usage

After installation, you'll find the XLerate tab in your Excel ribbon with two main tools:

### Formula Tracer
1. Select a cell with a formula
2. Click "Trace Precedents" or "Trace Dependents" in the XLerate ribbon (or add them to the quick access toolbar for faster use)
3. Navigate through dependencies using the arrow keys in the interactive dialog box

### Consistency Checker
1. Select a range of cells
2. Click "Horizontal Check" in the XLerate ribbon under Formula Consistency
3. Review highlighted inconsistencies in the selection (red) and check which formulas are consistent (green)

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
3. If you've changed APIs, update the documentation
4. Ensure the test suite passes
5. Make sure your code follows the existing style
6. Issue that pull request!

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
