# Excel Auto Filter Macro

This project contains a powerful VBA macro for Excel that automates the process of splitting a large dataset into separate worksheets based on the unique values of a specific column.

## Features

- **Automated Splitting**: Automatically creates a new worksheet for every unique value in your chosen column.
- **Smart Validation**: Checks for user cancellation and ensures valid selections.
- **Safe Sheet Naming**: Automatically sanitizes sheet names (removing invalid characters like `*`, `/`, `:`) and handles duplicate names by appending a timestamp.
- **Header Preservation**: Intelligently identifies and copies header rows to strictly preserve formatting.
- **Serial Number Generation**: Optionally regenerates serial numbers (1, 2, 3...) in the first column of the new 
sheets.
- **Performance Optimized**: Disables screen updating during processing to ensure maximum speed.

## Installation & Usage

### Method 1: For Beginners (One-Time Use)
1. Open your Excel workbook.
2. Press `ALT + F11` to open the VBA Editor.
3. Go to `Insert > Module`.
4. Copy the code from `auto_Filter_Macro.vba` in this repository.
5. Paste it into the empty module window.
6. Close the VBA Editor.
7. To run: Press `ALT + F8`, select `AutoFilter`, and click **Run**.

### Method 2: For Advanced Users (Personal Macro Workbook)
To make this macro available in *all* your Excel workbooks:
1. Open Excel and record a blank macro (View > Macros > Record Macro), storing it in "Personal Macro Workbook".
2. Stop recording immediately.
3. Press `ALT + F11` to open the editor.
4. Locate `VBAProject (PERSONAL.XLSB)` in the Project Explorer pane (left side).
5. Open `Modules > Module1` (or the new module created).
6. Paste the `AutoFilter` code there.
7. Save the Personal workbook (`CTRL + S` inside the editor).
8. **Pro Tip**: Add this macro to your Quick Access Toolbar or Ribbon for one-click access.

## How It Works in Detail
1. **Selection**: You select your entire data table (including headers).
2. **Column Pick**: You click a single cell in the column you want to group by (e.g., if you want a sheet for every "Region", click a cell in the "Region" column).
3. **Processing**:
   - The macro scans the column for unique values.
   - It filters the original list for each value.
   - It creates a new sheet named after that value (e.g., "North", "South").
   - It copies the visible data to the new sheet.

## Files
- `auto_Filter_Macro.vba`: The source code for the macro.

## License
MIT License. See `LICENSE` for details.
