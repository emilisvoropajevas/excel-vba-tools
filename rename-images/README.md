# VBA Image Renaming Macro

This VBA macro automates the process of renaming image ```.jpg``` files based on values stored in an Excel worksheet.

## Features

- Reads existing file paths from a worksheet
- Reads new file names from another column
- Renames `.jpg` image files automatically
- Processes multiple files in bulk
- Displays completion confirmation when finished 

---

## Requirements

This macro runs on **Microsoft Windows** only. It relies on the Windows ``ren`` shell command to rename files. This command is only available in the Windows command prompt and does not exist on macOS or Linux.

### Column Structure

| Column | Purpose |
|--------|---------|
| D | Existing image full path |
| E | New image name |

### Example
To generate full file paths, see the ```get-user-directory``` macro.

| D (Old Path) | E (New Name) |
|---|---|
| `C:\Images\IMG001.jpg` | `Product_001` |
| `C:\Images\IMG002.jpg` | `Product_002.jpg` |

> If `.jpg` is not included in the new name, the macro automatically appends it.

---

## VBA Modules

### `RenameImages`

Main procedure that:

1. Reads rows from the worksheet
2. Retrieves old and new file names
3. Calls `RenameFile` for each valid row

### `RenameFile`

Handles:

- Path validation
- `.jpg` extension handling
- Windows command creation
- File rename execution

### ExecCmd Function

The `ExecCmd` procedure executes Windows command-line instructions from VBA using `WScript.Shell`. In this macro, it is used to run the Windows `ren` command to rename image files silently in the background.

```cmd
ExecCmd "cmd.exe /c ren ""oldfile.jpg"" ""newfile.jpg"""
```

## Important Notes

### Limitation

This macro only works for ```.jpg``` extensions

### Full File Paths Required!

The old file name **must include the full file path**.
