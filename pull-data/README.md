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

# Workbook Structure

The macro uses the following worksheets:

| Worksheet | Purpose |
|---|---|
| `Introduction` | Stores the selected folder path |
| `Test_WorkingSheet` | Displays imported file names |


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

## Step 1 — Select the Source Folder

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

## Step 2 — Pull File Names Into Excel

Run the following macro:

```vb
PullData
```

or :

```vb
GetFileList
```

---

## Step 3 — View Imported Data

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

# Future Improvements

This VBA solution can easily be extended to support:

- File type filtering
- Subfolder scanning
- Image previews
- File metadata extraction
- Automated renaming workflows
- Duplicate file detection
