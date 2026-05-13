# Pull Data VBA Macro

## Overview

The `PullData` VBA macro is designed to pull file names from a selected folder into Excel.  
It is primarily intended for importing image file names, but it can also be used for any file type stored in a directory.

The macro automatically:

- Reads a folder path from the workbook
- Retrieves all file names from the selected folder
- Pastes the file names into a worksheet
- Sorts the file list numerically
- Clears previous imported data before loading new data

---

# Features

- Automatic file retrieval
- Numeric file sorting
- Clears old imported data before each run
- Works well with image folders
- Lightweight and fast execution
- Simple Excel integration

---

# Workbook Structure

The macro uses the following worksheets:

| Worksheet | Purpose |
|---|---|
| `Introduction` | Stores the selected folder path |
| `Test_WorkingSheet` | Displays imported file names |

---

# Required Cell Setup

The folder path must be stored in:

| Sheet | Cell |
|---|---|
| `Introduction` | `D3` |

Example:

```text
C:\Users\YourName\Pictures
```

---

# How To Use

## Step 1 — Open the Workbook

Open the Excel workbook containing the VBA macros and ensure:

- Macros are enabled
- The workbook is saved as a `.xlsm` file

---

## Step 2 — Select the Source Folder

Run the following macro:

```vb
GetUserDirectory
```

This macro allows the user to browse and select a folder containing image files or other documents.

The selected folder path will automatically populate:

```text
Introduction!D3
```

Example:

```text
C:\Users\YourName\Pictures
```

---

## Step 3 — Pull File Names Into Excel

Run the following macro:

```vb
PullData
```

or:

```vb
GetFileList
```

The macro will automatically:

1. Read the folder path from `Introduction!D3`
2. Retrieve all file names from the selected folder
3. Clear previous imported data
4. Paste the file names into `Test_WorkingSheet`
5. Sort the files numerically
6. Display a completion message

---

## Step 4 — View Imported Data

Imported file names will appear in:

| Worksheet | Starting Cell |
|---|---|
| `Test_WorkingSheet` | `A3` |

Example output:

```text
1.jpg
2.jpg
3.jpg
10.jpg
11.jpg
```

---

# Sorting Logic

The macro automatically sorts files numerically rather than alphabetically.

### Example

Without numeric sorting:

```text
1.jpg
10.jpg
11.jpg
2.jpg
```

After numeric sorting:

```text
1.jpg
2.jpg
10.jpg
11.jpg
```

The numeric value is extracted using:

```vb
Val(nameWithoutExt)
```

---

# Main VBA Procedures

| Procedure | Purpose |
|---|---|
| `PullData` | Main macro entry point |
| `GetFileList` | Pulls file names from folder |
| `ExecCmd` | Executes Windows command line |
| `SortFileList` | Sorts file names numerically |

---

# Output Location

Imported file names are placed starting from:

| Worksheet | Starting Cell |
|---|---|
| `Test_WorkingSheet` | `A3` |

---

# Requirements

- Microsoft Excel
- VBA enabled workbook (`.xlsm`)
- Windows operating system
- Macro permissions enabled

---

# Limitations

- Only retrieves file names
- Does not import actual image files
- Does not scan subfolders
- Sorting assumes filenames begin with numeric values

---

# Troubleshooting

## Folder Path Not Found

If the following message appears:

```text
Folder path not found
```

Check:

- The folder exists
- The path in `Introduction!D3` is correct
- The folder path does not contain invalid characters

---

## No Files Imported

Check:

- The selected folder contains files
- VBA macros are enabled
- Worksheet names are correct

---

# Example Workflow

1. Run `GetUserDirectory`
2. Select a folder containing images/files
3. Run `PullData`
4. Review imported file names in `Test_WorkingSheet`

---

# Future Improvements

This VBA solution can easily be extended to support:

- File type filtering
- Subfolder scanning
- Image previews
- File metadata extraction
- Automated renaming workflows
- Duplicate file detection

---
