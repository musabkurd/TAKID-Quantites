# TAKID Quantities Excel Automation

One-click Excel VBA automation for cleaning TAKID quantity confirmation source reports.

This project is built for daily warehouse/distributor follow-up work. Open the source Excel file, click the TAKID macro button, and the macro creates a clean confirmation workbook ready to send.

## What It Does

- Reads the active Excel workbook.
- Uses the distributor name from cell `D3`.
- Creates a clean TAKID result file named:

```text
Distributor Name - تأکید كمیات.xlsx
```

- Keeps only the useful columns:
  - `Item ID`
  - `Item No`
  - `Lot Number`
  - `Item Name 1`
  - `Item Qty`
  - `Expiry Date`
  - `Expiry Status`
- Formats the report with clean headers, borders, filters, and readable column widths.
- Adds a total under `Item Qty`.
- Highlights `.EXP` and `.BUG` text inside item names so promotion items stand out.
- Saves the result in the same folder as the source workbook.

## Main Macro

Use this macro from Excel:

```vb
Clean_FIFO_Source_Report
```

The macro is designed to run from `PERSONAL.XLSB`, so it can be used from any opened Excel file.

## Files In This Repo

| File | Purpose |
| --- | --- |
| `V1.2.bas` | Current VBA macro code |
| `CODE V1.1.txt` | Previous readable macro copy |
| `TEMPLATE.xlsx` | Excel template/reference file |
| `PERSONAL_MACRO_BACKUP/PERSONAL.XLSB` | Personal macro workbook backup |
| `PERSONAL_MACRO_BACKUP/Excel Customizations.exportedUI` | Quick Access Toolbar buttons/icons backup |
| `PERSONAL_MACRO_BACKUP/README.md` | Restore steps for a new laptop |

## Daily Use

1. Open the TAKID source Excel file.
2. Make sure the source workbook is the active workbook.
3. Click the TAKID button from the Quick Access Toolbar.
4. Wait for the success message.
5. Send the generated `Distributor Name - تأکید كمیات.xlsx` file.

## Install In PERSONAL.XLSB

1. Open Excel.
2. Unhide `PERSONAL.XLSB` if needed.
3. Press `Alt + F11`.
4. Open `VBAProject (PERSONAL.XLSB)`.
5. Insert a module.
6. Paste the code from `V1.2.bas`.
7. Save `PERSONAL.XLSB`.
8. Hide `PERSONAL.XLSB` again.

Important: when running global macros from `PERSONAL.XLSB`, the macro must work on `ActiveWorkbook`, not `ThisWorkbook`.

## Add The One-Click Button

1. Open Excel.
2. Go to `File > Options > Quick Access Toolbar`.
3. Change `Choose commands from` to `Macros`.
4. Add the TAKID macro.
5. Click `Modify...`.
6. Rename it to `TAKID`.
7. Pick an icon.
8. Save.

## Restore On A New Laptop

Use the backup folder:

```text
PERSONAL_MACRO_BACKUP
```

Restore both files:

- `PERSONAL.XLSB`
- `Excel Customizations.exportedUI`

`PERSONAL.XLSB` restores the macros. `Excel Customizations.exportedUI` restores the toolbar buttons/icons.

Full restore steps are inside:

```text
PERSONAL_MACRO_BACKUP/README.md
```

## Required Source Layout

The source file must contain the expected TAKID report headers and distributor name in:

```text
D3
```

If the source layout changes, update the header detection logic in `V1.2.bas`.

## Notes

- Keep only `PERSONAL.XLSB` inside the Excel `XLSTART` folder.
- Do not store random `.xlsx` files in `XLSTART`, because Excel will open them every time.
- Re-export `Excel Customizations.exportedUI` after changing toolbar buttons.

Made for Musab's daily stock and warehouse follow-up workflow.
