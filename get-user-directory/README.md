# Get User Directory VBA Macro

## Overview
This VBA macro retrieves the current Windows user's directory path and writes it into cell B3 in Excel. Validation is added if target cell is empty.

## Requirements

This macro runs on **Microsoft Windows** only. It relies on `Environ("USERPROFILE")`, 
a Windows environment variable that returns the current user's home directory. 
This variable does not exist on macOS or Linux.

## Usage
Run the following macro:

```vb
GetCurrentUserDirectory
```
The user directory path will appear in cell B3

```
C:\Users\YourName\
```

I've modified this output path to direct me to my set of folders with images. Freely adjust yours.

### Warning Message

If cell `A3` is empty, the following message will appear:

```text
Please enter folder name
```

---

## Macro: `Intro_End()`

A helper macro called automatically after `GetCurrentUserDirectory()` completes. 
It activates the `Test_WorkingSheet` worksheet.

## Setup

1. Open Excel and press `Alt + F11` to open the VBA editor
2. Insert a new module via **Insert → Module**
3. Paste the macro code into the module
4. Run `GetCurrentUserDirectory` via **Run → Run Macro** or `F5`
