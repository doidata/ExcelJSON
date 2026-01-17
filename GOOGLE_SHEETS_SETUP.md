# ExcelJSON for Google Sheets

This is a Google Apps Script version of the ExcelJSON VBA macro, designed to convert structured spreadsheet data into d3.js-compatible treemap JSON.

## Features

- ✅ Native JavaScript with built-in JSON support (much simpler than VBA!)
- ✅ Custom menu for easy access
- ✅ Outputs JSON to a dedicated sheet
- ✅ Can display JSON in a sidebar
- ✅ Proper error handling with user-friendly messages
- ✅ Works with the same data structure as the Excel version

## Installation

### Step 1: Open Your Google Sheet

1. Open Google Sheets and create a new spreadsheet or open an existing one
2. Set up your data structure (see Data Format section below)

### Step 2: Open the Script Editor

1. In Google Sheets, click **Extensions** → **Apps Script**
2. This will open the Google Apps Script editor in a new tab

### Step 3: Add the Code

1. Delete any existing code in the editor
2. Copy the entire contents of the `Code.gs` file from this repository
3. Paste it into the script editor
4. Click the **Save** icon (💾) or press `Ctrl+S` (Windows) / `Cmd+S` (Mac)
5. Give your project a name (e.g., "ExcelJSON")

### Step 4: Grant Permissions

1. Close the Apps Script tab and return to your Google Sheet
2. Refresh the page (F5 or Cmd+R)
3. You should see a new menu called **ExcelJSON** appear in the menu bar
4. Click **ExcelJSON** → **Generate JSON**
5. Google will ask you to authorize the script:
   - Click **Continue**
   - Select your Google account
   - Click **Advanced** → **Go to [Your Project Name] (unsafe)**
   - Click **Allow**

### Step 5: You're Ready!

The script is now installed and ready to use!

## Data Format

### Sheet 1: Nodes

The first sheet should contain your node hierarchy:

| id | parent | attribute1 | attribute2 | attribute3 |
|----|--------|------------|------------|------------|
| root_node | root | value1 | value2 | value3 |
| child1 | root_node | value1 | value2 | value3 |
| child2 | root_node | value1 | value2 | value3 |
| grandchild1 | child1 | value1 | value2 | value3 |

**Columns:**
- **Column A (id)**: Unique identifier for each node
- **Column B (parent)**: Parent node id (use "root" for top-level nodes)
- **Column C+**: Additional attributes (headers in row 1, values in data rows)

### Sheet 2: Children (Optional)

The second sheet can contain additional child data:

| parent | child_attribute1 | child_attribute2 |
|--------|------------------|------------------|
| node1 | value1 | value2 |
| node1 | value3 | value4 |
| node2 | value5 | value6 |

**Columns:**
- **Column A (parent)**: Parent node id
- **Column B+**: Child attributes

## Usage

### Method 1: Generate JSON to Sheet (Recommended)

1. Click **ExcelJSON** → **Generate JSON**
2. The script will create a new sheet called "JSON Output"
3. The JSON will be displayed in cell A2
4. Copy the JSON and use it in your d3.js visualization

### Method 2: View JSON in Sidebar

1. Click **ExcelJSON** → **Show JSON in Sidebar**
2. A sidebar will open on the right side of your screen
3. The formatted JSON will be displayed
4. You can copy it from there

## Output Format

The script generates JSON in the d3.js treemap format:

```json
{
  "name": "root_node",
  "attribute1": "value1",
  "attribute2": "value2",
  "children": [
    {
      "name": "child1",
      "attribute1": "value1",
      "children": [...]
    },
    {
      "name": "child2",
      "attribute1": "value1",
      "children": [...]
    }
  ]
}
```

## Differences from Excel VBA Version

| Feature | Excel VBA | Google Sheets |
|---------|-----------|---------------|
| **Language** | VBA (Visual Basic) | JavaScript (Apps Script) |
| **JSON Conversion** | Custom toString() function | Native JSON.stringify() |
| **UI Control** | ActiveX Button & TextBox | Custom menu |
| **Output** | TextBox control | Sheet or Sidebar |
| **Error Handling** | On Error GoTo | try/catch blocks |
| **Data Structures** | Collections & Dictionaries | Arrays & Objects |
| **Classes** | Custom VBA classes | JavaScript objects |

## Advantages of the Google Sheets Version

1. **Simpler Code**: JavaScript's native JSON support eliminates the need for custom serialization
2. **No Binary Files**: Everything is text-based and version-controllable
3. **Cloud-Based**: Works anywhere with internet access
4. **Auto-Save**: Changes are saved automatically
5. **Collaboration**: Multiple people can work on the same sheet
6. **No Installation**: No need to enable macros or worry about security warnings

## Troubleshooting

### Menu doesn't appear
- Refresh the page (F5)
- Close and reopen the spreadsheet
- Check that the script was saved properly

### "Authorization Required" error
- Follow the permission steps in the Installation section
- Make sure you clicked "Allow" when prompted

### "No root node found" error
- Check that at least one node has "root" as its parent in Sheet 1, column B

### JSON looks wrong
- Verify your data format matches the examples above
- Check that column headers are in row 1
- Make sure there are no extra empty rows between data

### Empty output
- Ensure Sheet 1 has data starting from row 2 (row 1 is headers)
- Check that column A (id) and column B (parent) have values

## Support

For issues or questions:
1. Check the [BUG_FIX.md](BUG_FIX.md) for common problems
2. Review your data structure against the examples
3. Check the Apps Script logs: **Extensions** → **Apps Script** → **View** → **Logs**

## License

Same license as the parent ExcelJSON project (see LICENSE file).
