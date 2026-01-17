# ExcelJSON

Convert structured spreadsheet data to d3.js-compatible treemap JSON.

## Available Versions

### 📊 Microsoft Excel (VBA)
- **File**: `ExcelJSON.xlsm`
- **Language**: Visual Basic for Applications (VBA)
- **Fixed Version**: `Arkusz1_FIXED.cls` (bug-free code)
- **See**: [BUG_FIX.md](BUG_FIX.md) for bug fixes and installation

### ☁️ Google Sheets (Apps Script)
- **File**: `Code.gs`
- **Language**: JavaScript (Google Apps Script)
- **See**: [GOOGLE_SHEETS_SETUP.md](GOOGLE_SHEETS_SETUP.md) for installation

## Quick Start

### For Excel Users
1. Download `ExcelJSON.xlsm`
2. Open in Microsoft Excel
3. Enable macros if prompted
4. Set up your data in sheets 1 and 2
5. Click the button to generate JSON

**⚠️ Bug Fix Available**: See [BUG_FIX.md](BUG_FIX.md) to apply fixes for better reliability

### For Google Sheets Users
1. Copy the code from `Code.gs`
2. Open Google Sheets → Extensions → Apps Script
3. Paste the code and save
4. Refresh your sheet and use the **ExcelJSON** menu
5. See [GOOGLE_SHEETS_SETUP.md](GOOGLE_SHEETS_SETUP.md) for detailed instructions

## What It Does

Converts hierarchical spreadsheet data into JSON format suitable for d3.js treemap visualizations.

**Input**: Structured data with parent-child relationships
**Output**: Nested JSON ready for d3.js

## Example Use Case

Perfect for creating interactive treemap visualizations from spreadsheet data, such as:
- Organizational hierarchies
- Budget breakdowns
- File system visualizations
- Product categorizations
- Any hierarchical data structure

## Files

- `ExcelJSON.xlsm` - Excel macro-enabled workbook (VBA version)
- `treemapdata03.xlsx` - Example data file
- `Arkusz1_FIXED.cls` - Bug-fixed VBA code (recommended)
- `Code.gs` - Google Apps Script version
- `BUG_FIX.md` - Documentation of bugs found and fixes
- `GOOGLE_SHEETS_SETUP.md` - Google Sheets installation guide

## License

MIT License (see LICENSE file)
d3js compatible JSON from strucured Excel sheet
VBA code inside the main file "ExcelJSON", the other files are only documentation or examples
use it with the d3js layout "collapsible tree"
cheers, karl.
