# Excel VBA Macros

## Overview
A collection of Excel VBA macros designed to automate image file 
renaming workflows in Windows. Built to remove manual command-line 
steps and reduce repetitive work for non-technical users.

## Requirements

Open the Excel workbook containing the VBA macros and ensure:

- Macros are enabled
- The workbook is saved as a `.xlsm` file

## Purpose
The macros were designed to automate renaming images, specifically ```.jpg``` files. Originally names of images were pulled manually from Windows ```cmd``` and formatted with ```Text -> Column```. Now the process is automated by separating each sub-process into a macro. Overall this reduced the need for non technical people to pull image names for manual renaming. In total this saved around 30 mins of work every day which compounds to 2.5 hours of work per week.

## Macros

| Macro | File | Purpose |
|---|---|---|
| GetCurrentUserDirectory | get-user-directory.bas | Gets Windows user path |
| PullData | pull-data.bas | Pulls image file paths into the Test Sheet |
| RenameImages | rename-images.bas | Renames .jpg files |
| TransferData | transfer-data.bas | Copies mapped columns to Working Sheet |

## Setup

1. Open the `.xlsm` workbook
2. Enable macros when prompted
3. Ensure sheet names match exactly: `Test_WorkingSheet` and `Working Sheet`
4. Run macros from `Alt + F11` or the Developer tab


## Limitations

- Windows only — relies on `Environ("USERPROFILE")` and the `ren` shell command
- `.jpg` files only
- Sheet names must match exactly or macros will fail


