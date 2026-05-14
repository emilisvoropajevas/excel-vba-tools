# VBA Transfer Data Macro

## Overview

This VBA macro copies data from `Test_WorkingSheet` into `Working Sheet` using a fixed column-to-column mapping.

It is designed to:
- Activate the `Working Sheet` tab
- Read rows from `Test_WorkingSheet`
- Copy selected source columns into specific destination columns
- Start copying from row 3 in the source sheet
- Paste data beginning at row 2 in the destination sheet

## Files / Sheets Required

This macro expects the workbook to contain these worksheets exactly:

- `Test_WorkingSheet`
- `Working Sheet`

If either sheet name is different, the macro will fail. You can adjust the code and the source/destination sheets to your liking.

## What the Macro Does

1. Runs `Open_Working_Sheet` to activate `Working Sheet`
2. Finds the last used row in column `B` of `Test_WorkingSheet`
3. Loops from row 3 through the last row
4. Copies values from source columns to destination columns based on the mapping list
5. Shows a confirmation message when finished

## Column Mapping

| Source Column | Destination Column |
|---|---|
| B | C |
| E | K |
| F | N |
| G | A |
| H | E |
| I | F |
| J | G |
| K | D |
| L | H |
| M | I |
| N | J |
| O | O |
| P | L |
| Q | B |
| R | M |

Column headers are specific to this workbook. The mappings were chosen specifically to map data on products to a fresh sheet to prep for uploads. You can adjust this accordingly.
## Row Mapping

- Source starts at row **3**
- Destination starts at row **2** because the code uses:
  ```vb
  destSheet.Cells(i - 1, ...)
  ```

So:
- Source row 3 → Destination row 2
- Source row 4 → Destination row 3
- Source row 5 → Destination row 4
- and so on

## Macro: `Open_Working_Sheet()`

A helper macro called automatically during `Transfer_Data()`. It activates the `Working Sheet` worksheet.

## How to Use

1. Open the Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Insert the macro into a standard module.
4. Make sure the sheet names match exactly.
5. Run `Transfer_Data`.

## Notes

- The macro copies **values only**, not formulas or formatting.
- Existing data in the destination sheet will be overwritten in the mapped cells.
- The macro does not clear old data before writing new data.
- If there is no data below row 2 in `Test_WorkingSheet`, only the message box will appear after the sheet is activated.

---

## Possible Improvements

- Add error handling for missing sheets
- Clear destination rows before copying
- Copy formatting as well as values
- Use named constants for sheet names to make maintenance easier
