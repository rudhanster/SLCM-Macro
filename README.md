# 📘 SLCM Attendance Automation [Windows]

This project automates marking **student attendance** on **SLCM (Salesforce Lightning)** using **Python + Selenium**, launched directly from **Excel VBA** on Windows.

---

## 🚀 Features
- Select a **date header cell** in Excel → run the macro → automation takes over.
- Reads:
  - **Absentees** from the `Attendance` sheet (row 2 = headers, “Reg. No.” column, `ab/ABSENT` marks).
  - **Course details** from the `Initial Setup` sheet (B1–B5).
- Opens Chrome, navigates to Salesforce Lightning Calendar, finds the correct class event.
- Unticks absentees, clicks **Submit Attendance**, confirms submission.
- Supports **section specificity** (`B` will not match `B-1` / `B-2`).

---

## 📂 Project Structure

```
slcm-attendance/
├─ maa.py                  # Python Selenium automation script
├─ excel/
│  └─ RunAttendance.bas    # VBA macro module (Windows)
├─ README.md               # This file
├─ .gitignore
└─ (optional) workbook.xlsm
```

---

## 🖥️ Requirements

- Windows 10/11  
- **Google Chrome** (latest version)  
- **Python 3.10+** with:
  ```bash
  pip install pandas selenium webdriver-manager
  ```
- Excel workbook with:
  - Sheet **Attendance**
    - Row 2 = headers
    - One column labeled like “Reg. No.”
    - Absentees marked `ab` or `ABSENT`
  - Sheet **Initial Setup**
    - B1: Course Name  
    - B2: Course Code  
    - B3: Semester  
    - B4: Class Section (e.g., `B` or `B-1`)  
    - B5: Session No (optional)

---

## ⚙️ Setup

1. **Clone this repo** or download the files.  
2. Place `maa.py` somewhere accessible (e.g., `C:\Users\<you>\Desktop\testSlcm\maa.py`).  
3. Open your Excel workbook (**.xlsm** format).  
4. Open VBA editor (**ALT+F11**) → **Insert → Module**.  
5. Import `excel/RunAttendance.bas`.  
6. At the top of the module, update paths:
   ```vb
   Private Const PYTHON_EXE As String = "C:\Path\To\Python\python.exe"
   Private Const PY_SCRIPT  As String = "C:\Path\To\maa.py"
   ```
   - Run `where python` in Command Prompt to find your Python path.

---

## ▶️ Running the Automation

1. In Excel, go to the **Attendance** sheet.  
2. Select the **date header cell** you want to process.  
3. Run the macro:
   - Press **ALT+F8**
   - Select `RunAttendanceForActiveWorkbook`
   - Click **Run**  
4. A **Command Prompt** will open and run `maa.py`.  
   - If SSO login appears, complete it in Chrome and press Enter in console.  
   - Watch Selenium untick absentees and submit attendance.  

---

## 🔧 Customization

- **Console closes automatically**: change `cmd.exe /K` to `cmd.exe /C` in VBA.  
- **Headless mode**: uncomment `--headless=new` in `maa.py` to hide Chrome.  
- **Timeouts**: adjust constants like `EVENT_SEARCH_TIMEOUT` in `maa.py` if pages load slowly.  

---

## 📊 Example Workflow

1. Select **Aug 1, 2025** header cell in Excel.  
2. Macro gathers absentees from that column.  
3. Reads course code, semester, section from *Initial Setup*.  
4. Launches Chrome → finds event tile → opens **Attendance** tab.  
5. Unticks absentees → clicks **Submit Attendance** → confirms.  
6. Console prints a summary:  
   ```
   ✔️ Unticked: 230905016
   ❌ Not found: 230905064
   ✅ Confirmed submission
   🎉 SLCM Attendance automation completed!
   ```

---

## ⚠️ Troubleshooting

- **Macro not running** → Enable macros in Excel (Trust Center Settings).  
- **Date not found** → Ensure the selected header matches the format (`m/d/yyyy` if Excel date).  
- **ChromeDriver error** → Let `webdriver-manager` auto-install, or update Chrome.  
- **Event not found** → Ensure Course Code, Semester, Section match Salesforce event text exactly.  

---

## 👨‍💻 Author

Developed by **Anirudhan Adukkathayar C**  
SCE, MIT
