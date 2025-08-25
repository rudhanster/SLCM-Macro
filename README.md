# 📘 SLCM Attendance Automation (Windows)

Automates marking **student attendance** on **SLCM (Salesforce Lightning)** using **Python + Selenium**, launched directly from **Excel VBA** on Windows.

---

## ✅ What you’ll do (high level)
1. Convert your workbook from **.xlsx** to **.xlsm** (macro-enabled).
2. Import the provided VBA module `RunAttendance.bas` into the .xlsm.
3. Set two paths at the top of the module (your Python & `maa.py` paths).
4. Select a **date header cell** and run the macro → the Python script automates attendance in Chrome.

---

## 🖥️ Requirements
- Windows 10/11
- Google Chrome (latest)
-
### Python
- Install **Python 3.11+**.
- **Windows**:
  - Open **Microsoft Store**, search for **Python**, install it.
- Verify:
  ```bash
  python --version
  ```

### Dependencies
Install in one line:
```bash
pip install selenium pandas openpyxl webdriver-manager
```
###- Your Excel workbook with:
  - **Attendance** sheet:
    - Row 2 = headers; one column named like **Reg. No.**
    - Data starts from row 3; absentees marked `ab` or `ABSENT`
  - **Initial Setup** sheet:
    - B1: Course Name
    - B2: Course Code
    - B3: Semester
    - B4: Class Section (e.g., `B` or `B-1`)
    - B5: Session No (optional)

---

## 🔄 Convert XLSX to XLSM (keep your data)
1. Open your current `.xlsx` in Excel.
2. **File → Save As** (or **Save a Copy**).
3. Choose **Save as type**: **Excel Macro-Enabled Workbook (*.xlsm)**
4. Save as e.g. `AttendanceWorkbook.xlsm` (you may keep the .xlsx as backup).

> `.xlsx` cannot store macros. Use `.xlsm` for the macro-enabled version.

---

## ➕ Import the macro into the .xlsm
1. Open `AttendanceWorkbook.xlsm`.
2. Press **ALT+F11** to open the VBA editor.
3. **File → Import File…** → select `RunAttendance.bas`.
4. In the imported module (top of file), edit these two constants:
   ```vb
   Private Const PYTHON_EXE As String = "C:\Path\To\Python\python.exe"
   Private Const PY_SCRIPT  As String = "C:\Path\To\maa.py"
   ```
   - Find Python path via: `where python` (Command Prompt).
5. Close the editor and **save**. Reopen Excel if prompted and click **Enable Content** (to allow macros).

---

## ▶️ Run the automation
1. In the **Attendance** sheet, click the **date header cell** you wish to submit.
2. Press **ALT+F8** → choose `RunAttendanceForActiveWorkbook` → **Run**.
3. A Command Prompt opens and runs `maa.py`:
   - Complete SSO in Chrome if asked; return to console when prompted.
   - Script finds your class event, opens **Attendance**, unticks absentees, and submits.

---

## 🔧 Customization
- **Close console automatically**: in VBA change `cmd.exe /K` to `cmd.exe /C`.
- **Headless** Chrome: in `maa.py`, uncomment `--headless=new` (recommended only after stabilizing).
- **Timeouts**: adjust `PANEL_READY_TIMEOUT`, `EVENT_SEARCH_TIMEOUT`, etc., in `maa.py` for slow pages.

---

## 🧪 Example console output
```
📅 Selected Date : 2025-08-01
🧑‍🎓 Absentees   : 230905016, 230905064
✅ Opened Calendar
✅ Opened Attendance tab
✔️ Unticked: 230905016
❌ Not found: 230905064
✅ Confirmed submission
🎉 SLCM Attendance automation completed!
```

---

## ❗ Troubleshooting
- **Macro disabled**: File → Options → Trust Center → Trust Center Settings → Macro Settings → enable / trusted location.
- **Date column not found**: Ensure you selected the header cell; the macro formats real date cells as `m/d/yyyy`.
- **Chrome/driver mismatch**: `webdriver-manager` fetches correct driver automatically; ensure internet access or install manually.
- **Event not found**: Ensure Course Code, Semester, Section in *Initial Setup* exactly match the Salesforce event text. Note: `B` will **not** match `B-1/B-2` by design.
- **Paths invalid**: Update `PYTHON_EXE` and `PY_SCRIPT` constants to your actual paths.




---

## 👨‍💻 Author
Developed by **Anirudhan Adukkathayar C**, SCE, MIT

