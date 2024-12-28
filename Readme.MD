# XLerate

XLerate is an Excel add-in designed to speed up financial modeling tasks and spot potential errors with ease incl. advanced auditing and formula consistency tools.

## Features

### Advanced Formula Tracer
- Trace all precedents and dependents in a single operation
- Visualize formula dependencies across multiple worksheets
- Quick navigation through formula chains

### Formula Consistency Checker
- Instantly identify inconsistencies in adjacent formulas
- Visual highlighting of formula pattern breaks
- Help prevent common modeling errors
- Support for complex financial modeling patterns

## Installation

### Windows
1. Download the latest release (`XLerate.xlam`) from the `dist` folder
2. Place the file in your Excel add-ins folder (typically `C:\Users\[Username]\AppData\Roaming\Microsoft\AddIns`)
3. Enable the add-in in Excel:
   - File → Options → Add-ins
   - Manage: Excel Add-ins → Go
   - Check "XLerate" and click OK

### Mac
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

## Usage

After installation, you'll find the XLerate tab in your Excel ribbon with two main tools:

### Formula Tracer
1. Select a cell with a formula
2. Click "Trace All" in the XLerate ribbon
3. Navigate through dependencies using the interactive map

### Consistency Checker
1. Select a range of cells
2. Click "Check Consistency" in the XLerate ribbon
3. Review highlighted inconsistencies in the selection

## Development

See [CONTRIBUTING.md](docs/contributing.md) for development setup and guidelines.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- Report bugs via [GitHub Issues](https://github.com/[your-username]/XLerate/issues)
- Request features through our [Feature Request](https://github.com/[your-username]/XLerate/issues/new?template=feature_request.md) template

## Acknowledgments

Built with love for financial analysts by financial analysts.
