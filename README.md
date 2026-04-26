# ⚡ TAKID Excel Automation (Global VBA Setup)

## 🎯 Goal

Run a VBA macro from **any Excel file** using **one button**, without copying code into each file.

---

## 🧠 What We Built

* Global macro using `PERSONAL.XLSB`
* One-click button in Excel toolbar
* Auto-save result file in same folder as source file
* Clean Excel startup (no unwanted files opening)

---

## 🛠️ Setup (From Scratch)

### 1. Create PERSONAL.XLSB

* Open Excel
* View → Record Macro
* Store macro in: **Personal Macro Workbook**
* Stop recording

---

### 2. Add Your Code

* Press `Alt + F11`
* Open: `VBAProject (PERSONAL.XLSB)`
* Insert → Module
* Paste your VBA code
* Make sure main macro name is:

```vb
Public Sub takid()
```

---

### 3. Fix Save Location (IMPORTANT)

Use this logic:

```vb
sourcePath = wsSource.Parent.Path
sourceName = wsSource.Parent.Name
```

This ensures output saves in the **same folder as the source file**, not PERSONAL.

---

### 4. Add One-Click Button

* File → Options → Quick Access Toolbar
* Choose commands from: **Macros**
* Add: `PERSONAL.XLSB!takid`
* (Optional) Rename + change icon

---

### 5. Clean Startup Folder

Go to:

```
%AppData%\Microsoft\Excel\XLSTART
```

Keep ONLY:

```
PERSONAL.XLSB
```

Delete:

* old result files
* template files
* any `.xlsx` files

---

## ⚠️ Key Concepts

| Concept           | Meaning                             |
| ----------------- | ----------------------------------- |
| `ThisWorkbook`    | Where the code is stored (PERSONAL) |
| `ActiveWorkbook`  | File you're working on              |
| `wsSource.Parent` | Best reference to source file       |

---

## 🚀 Final Workflow

1. Open any Excel file
2. Go to correct sheet
3. Click ⚡ TAKID button
4. Done

---

## 💡 What We Learned

* How to use `PERSONAL.XLSB` for global macros
* Why `ThisWorkbook` caused wrong save location
* How Excel auto-loads files from XLSTART
* How to create a reusable automation system

---

## 🏁 Result

* 1-click automation
* Works on any file
* Clean & professional setup

---

## 🔥 Future Upgrades (Optional)

* Auto-detect correct sheet
* Auto-close result file
* Add success notification (Arabic/English)
* Chain multiple macros

---

Made by Musab Qasim ⚡
## PERSONAL.XLSB Quick Setup

Exactly. That workflow is correct.

Your clean version is:

1. Unhide `PERSONAL.XLSB`
2. Press `Alt + F11`
3. Paste the VBA into a module in `PERSONAL.XLSB`
4. Replace `ThisWorkbook` with `ActiveWorkbook` where needed
5. Save
6. Hide `PERSONAL.XLSB` again
7. In Excel Options > Quick Access Toolbar
8. Change command list to `Macros`
9. Add the macro you want
10. Click `Modify...` to change the name and icon
11. Done

That is the best setup for your daily Excel automation across all files.

